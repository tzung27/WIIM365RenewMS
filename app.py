import io
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st



# ------------------------------------------------------------
# Page config / styles
# ------------------------------------------------------------
st.set_page_config(
    page_title="Weblink M365 續約概況與精準行銷推廣平台",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 2.2rem; padding-bottom: 2rem;}
    .kpi-label {font-size: 0.82rem; color: #5b6b7a; margin-bottom: 0.2rem;}
    .kpi-value {font-size: 1.6rem; font-weight: 700; color: #183153; line-height: 1.1;}
    .section-title {font-size: 1.05rem; font-weight: 700; margin: 0.2rem 0 0.6rem 0;}
    .card {
        background: #ffffff;
        border: 1px solid rgba(49, 51, 63, 0.12);
        border-radius: 16px;
        padding: 1rem 1.1rem;
        box-shadow: 0 4px 14px rgba(0,0,0,0.04);
        min-height: 118px;
    }
    .subtle {color: #6b7280; font-size: 0.92rem;}
    /* ── Sidebar 字體縮小 ── */
    section[data-testid="stSidebar"] {font-size: 0.82rem;}
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {font-size: 0.95rem !important; margin-bottom: 0.4rem;}
    section[data-testid="stSidebar"] label {font-size: 0.82rem !important;}
    section[data-testid="stSidebar"] .stCheckbox label,
    section[data-testid="stSidebar"] .stRadio label,
    section[data-testid="stSidebar"] .stMultiSelect label,
    section[data-testid="stSidebar"] .stDateInput label,
    section[data-testid="stSidebar"] .stSelectbox label {font-size: 0.82rem !important;}
    section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] {font-size: 0.75rem;}
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span {font-size: 0.82rem;}
    section[data-testid="stSidebar"] hr {margin: 0.5rem 0;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Constants
# ------------------------------------------------------------
DEFAULT_LOCAL_XLSX_PATH = r"D:\DeskT\Austin 自動化\新增資料夾\CSP訂單資料_raw.xlsx"
STATUS_TO_REMOVE = {"下單異常", "已取消", "已退貨"}

# 每次修改資料清洗邏輯（prepare_cleaned_df）後請更新此版本號，
# 以強制 @st.cache_data 快取失效，避免舊快取傳回未過濾資料。
_CLEAN_CACHE_VERSION = "v5"  # v4→v5: 保留成交單價未稅欄位供 CSP 年化收入計算

CUSTOMER_ALIASES = {
    "最終用戶": "最終客戶",
    "最終客户": "最終客戶",
    "客户": "最終客戶",
}

COLUMN_ALIASES = {
    "資格經銷商": "資格",
    "訂單編號": "經銷商訂單編號",
    "最終用戶": "最終客戶",
}

DROP_CANDIDATE_GROUPS = [
    ["客戶微軟ID"],
    ["資格經銷商", "資格"],
    ["訂單編號", "經銷商訂單編號"],
    ["展碁料號"],
    ["微軟料號"],
    ["展碁業務部門"],
    ["展碁內勤部門"],
    ["微軟下單日"],
    ["展碁PNS"],
    ["微軟Order ID"],
    ["微軟SubID"],
    ["退貨日"],
    ["退貨單號"],
    ["原訂閱之微軟SubID"],
    ["原訂閱之訂單下單日"],
    ["原訂閱之訂單到期日"],
]

NUMERIC_COLUMNS = ["數量", "成交單價未稅", "展碁COST單價未稅", "展碁COST未稅小計", "成交價未稅小計"]
DATE_COLUMNS = [
    "開單日",
    "訂單下單日",
    "訂閱到期日",
    "微軟下單日",
    "退貨日",
    "原訂閱之訂單下單日",
    "原訂閱之訂單到期日",
]
REQUIRED_FOR_ANALYSIS = ["訂閱到期日", "訂單下單日", "經銷商", "最終客戶", "商品名稱", "成交價未稅小計"]

FISCAL_MONTH_MAP = {7: 1, 8: 2, 9: 3, 10: 4, 11: 5, 12: 6, 1: 7, 2: 8, 3: 9, 4: 10, 5: 11, 6: 12}
QUARTER_BY_MONTH = {7: "Q1", 8: "Q1", 9: "Q1", 10: "Q2", 11: "Q2", 12: "Q2", 1: "Q3", 2: "Q3", 3: "Q3", 4: "Q4", 5: "Q4", 6: "Q4"}
QUARTER_ORDER = ["Q1", "Q2", "Q3", "Q4"]


@dataclass
class AnalysisPeriod:
    label: str
    start: pd.Timestamp
    end: pd.Timestamp


# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------
def fmt_int(value) -> str:
    if pd.isna(value):
        return "-"
    try:
        return f"{int(round(float(value))):,}"
    except Exception:
        return "-"


def fmt_currency(value) -> str:
    if pd.isna(value):
        return "-"
    try:
        return f"{float(value):,.0f}"
    except Exception:
        return "-"


def fmt_currency_compact(value) -> str:
    """以 M / K 縮寫顯示大金額，例如 $405.8M、$38.5M、$123K。"""
    if pd.isna(value):
        return "-"
    try:
        v = float(value)
        if abs(v) >= 1_000_000:
            return f"{v/1_000_000:.1f}M"
        if abs(v) >= 1_000:
            return f"{v/1_000:.1f}K"
        return f"{v:,.0f}"
    except Exception:
        return "-"


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    cols = []
    for col in df.columns:
        c = str(col).strip()
        c = CUSTOMER_ALIASES.get(c, c)
        c = COLUMN_ALIASES.get(c, c)
        cols.append(c)
    df.columns = cols
    return df


def robust_to_datetime(series: pd.Series) -> pd.Series:
    dt = pd.to_datetime(series, errors="coerce")
    mask = dt.isna() & series.notna()
    if mask.any():
        cleaned = (
            series.astype(str)
            .str.replace("上午", "AM", regex=False)
            .str.replace("下午", "PM", regex=False)
            .str.replace("/", "-", regex=False)
            .str.strip()
        )
        dt2 = pd.to_datetime(cleaned, errors="coerce")
        dt = dt.fillna(dt2)
    return dt


def infer_default_periods() -> tuple[AnalysisPeriod, AnalysisPeriod]:
    return (
        AnalysisPeriod("去年度", pd.Timestamp("2025-07-01"), pd.Timestamp("2026-06-30")),
        AnalysisPeriod("今年度", pd.Timestamp("2026-07-01"), pd.Timestamp("2027-06-30")),
    )


def date_input_to_tuple(value):
    if isinstance(value, tuple) and len(value) == 2:
        return value
    if isinstance(value, list) and len(value) == 2:
        return tuple(value)
    return None


def get_fy_label_from_date(dt: pd.Timestamp) -> str:
    if pd.isna(dt):
        return ""
    fy_num = dt.year + 1 if dt.month >= 7 else dt.year
    return f"FY{str(fy_num % 100).zfill(2)}"


def quarter_labels_for_period(period: AnalysisPeriod) -> list[str]:
    fy = get_fy_label_from_date(period.start)
    return [f"{fy} Q1", f"{fy} Q2", f"{fy} Q3", f"{fy} Q4"]


def month_sequence_for_period(period: AnalysisPeriod) -> list[pd.Timestamp]:
    start = pd.Timestamp(period.start).replace(day=1)
    return list(pd.date_range(start=start, periods=12, freq="MS"))


def month_labels_for_period(period: AnalysisPeriod) -> list[str]:
    return [m.strftime("%Y-%m") for m in month_sequence_for_period(period)]


def add_fy_columns(df: pd.DataFrame, expiry_col: str = "訂閱到期日") -> pd.DataFrame:
    df = df.copy()
    if expiry_col not in df.columns:
        df["FY年度"] = pd.NA
        df["季度代碼"] = pd.NA
        df["_quarter_short"] = pd.NA
        df["_fiscal_month_order"] = pd.NA
        df["_fiscal_month_label"] = pd.NA
        return df

    expiry = pd.to_datetime(df[expiry_col], errors="coerce")
    month = expiry.dt.month
    year = expiry.dt.year

    fy_num = pd.Series(np.where(month >= 7, year + 1, year), index=df.index)
    fy_num = pd.to_numeric(fy_num, errors="coerce")

    fy_label = pd.Series(pd.NA, index=df.index, dtype="string")
    valid_fy = expiry.notna() & fy_num.notna()
    fy_label.loc[valid_fy] = "FY" + fy_num.loc[valid_fy].astype(int).astype(str).str[-2:].str.zfill(2)

    quarter_short = month.map(QUARTER_BY_MONTH).astype("string")
    quarter_code = pd.Series(pd.NA, index=df.index, dtype="string")
    valid_q = fy_label.notna() & quarter_short.notna()
    quarter_code.loc[valid_q] = fy_label.loc[valid_q] + " " + quarter_short.loc[valid_q]

    fiscal_month_order = month.map(FISCAL_MONTH_MAP).astype("Int64")
    fiscal_month_label = pd.Series(pd.NA, index=df.index, dtype="string")
    fiscal_month_label.loc[expiry.notna()] = expiry.loc[expiry.notna()].dt.strftime("%Y-%m")

    df["FY年度"] = fy_label
    df["季度代碼"] = quarter_code
    df["_quarter_short"] = quarter_short
    df["_fiscal_month_order"] = fiscal_month_order
    df["_fiscal_month_label"] = fiscal_month_label
    return df


def drop_candidate_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    df = df.copy()
    existing = set(df.columns)
    to_drop = []
    dropped = []
    for group in DROP_CANDIDATE_GROUPS:
        for candidate in group:
            canonical = COLUMN_ALIASES.get(candidate, candidate)
            if candidate in existing:
                to_drop.append(candidate)
                dropped.append(candidate)
                break
            if canonical in existing:
                to_drop.append(canonical)
                dropped.append(canonical)
                break
    if to_drop:
        df = df.drop(columns=list(dict.fromkeys(to_drop)), errors="ignore")
    return df, list(dict.fromkeys(dropped))


@st.cache_data(show_spinner=False)
def load_excel_from_bytes(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes))


@st.cache_data(show_spinner=False)
def load_excel_from_path(path_str: str) -> pd.DataFrame:
    return pd.read_excel(path_str)


def prepare_cleaned_df(raw_df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    df = normalize_columns(raw_df)
    info = {
        "warnings": [],
        "dropped_columns": [],
        "missing_columns": [],
        "date_parse_failures": {},
    }

    for col in DATE_COLUMNS:
        if col in df.columns:
            before_non_null = int(df[col].notna().sum())
            parsed = robust_to_datetime(df[col])
            after_non_null = int(parsed.notna().sum())
            info["date_parse_failures"][col] = max(before_non_null - after_non_null, 0)
            df[col] = parsed

    if "訂單狀態" in df.columns:
        mask_remove = df["訂單狀態"].astype("string").str.strip().isin(STATUS_TO_REMOVE)
        df = df.loc[~mask_remove].copy()
    else:
        info["warnings"].append("缺少【訂單狀態】欄位，無法執行狀態排除。")

    # 刪除資格為「教育」或「非營利」的資料列（完全比對）
    if "資格" in df.columns:
        _qual_mask = df["資格"].astype("string").str.strip().isin({"教育", "非營利"})
        df = df.loc[~_qual_mask].copy()
    else:
        info["warnings"].append("缺少【資格】欄位，無法執行教育／非營利資料排除。")

    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = add_fy_columns(df, "訂閱到期日")
    df, dropped = drop_candidate_columns(df)
    info["dropped_columns"] = dropped

    for col in REQUIRED_FOR_ANALYSIS:
        if col not in df.columns:
            info["missing_columns"].append(col)

    return df, info


@st.cache_data(show_spinner=False)
def clean_data_from_bytes(file_bytes: bytes, cache_version: str = _CLEAN_CACHE_VERSION) -> tuple[pd.DataFrame, dict, pd.DataFrame]:
    raw_df = load_excel_from_bytes(file_bytes)
    cleaned_df, info = prepare_cleaned_df(raw_df)
    return cleaned_df, info, raw_df


@st.cache_data(show_spinner=False)
def clean_data_from_path(path_str: str, cache_version: str = _CLEAN_CACHE_VERSION) -> tuple[pd.DataFrame, dict, pd.DataFrame]:
    raw_df = load_excel_from_path(path_str)
    cleaned_df, info = prepare_cleaned_df(raw_df)
    return cleaned_df, info, raw_df


def apply_filters(
    df: pd.DataFrame,
    reseller_values: list[str],
    customer_values: list[str],
    expiry_range: tuple[date, date] | None,
    order_range: tuple[date, date] | None,
    staff_values: list[str] | None = None,
) -> pd.DataFrame:
    out = df.copy()

    if "經銷商" in out.columns and reseller_values:
        out = out[out["經銷商"].astype(str).isin(reseller_values)]
    if "最終客戶" in out.columns and customer_values:
        out = out[out["最終客戶"].astype(str).isin(customer_values)]
    if "展碁業務" in out.columns and staff_values:
        out = out[out["展碁業務"].astype(str).isin(staff_values)]

    if expiry_range and "訂閱到期日" in out.columns:
        start_dt = pd.Timestamp(expiry_range[0])
        end_dt = pd.Timestamp(expiry_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        out = out[out["訂閱到期日"].between(start_dt, end_dt, inclusive="both")]

    if order_range and "訂單下單日" in out.columns:
        start_dt = pd.Timestamp(order_range[0])
        end_dt = pd.Timestamp(order_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        out = out[out["訂單下單日"].between(start_dt, end_dt, inclusive="both")]

    return out


def build_period_mask(df: pd.DataFrame, period: AnalysisPeriod, date_col: str = "訂閱到期日") -> pd.Series:
    if date_col not in df.columns:
        return pd.Series(False, index=df.index)
    end_ts = pd.Timestamp(period.end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    return df[date_col].between(pd.Timestamp(period.start), end_ts, inclusive="both")


def calculate_forecast(df: pd.DataFrame, last_period: AnalysisPeriod, this_period: AnalysisPeriod) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = df.copy().reset_index(drop=True)
    df["__row_id"] = np.arange(len(df))
    df["Forecast"] = 0.0

    required = ["最終客戶", "商品名稱", "成交價未稅小計", "訂閱到期日"]
    empty_cols = [
        "最終客戶", "商品名稱", "去年度金額", "今年度金額", "Forecast",
        "_quarter_short", "_fiscal_month_order", "_fiscal_month_label"
    ]
    if not all(col in df.columns for col in required):
        return df.drop(columns=["__row_id"], errors="ignore"), pd.DataFrame(columns=empty_cols)

    work = df.copy()
    work["_in_last_period"] = build_period_mask(work, last_period)
    work["_in_this_period"] = build_period_mask(work, this_period)
    work["分析年度"] = np.select(
        [work["_in_last_period"], work["_in_this_period"]],
        [last_period.label, this_period.label],
        default=None,
    )

    base = work[work["分析年度"].isin([last_period.label, this_period.label])].copy()
    if base.empty:
        return df.drop(columns=["__row_id"], errors="ignore"), pd.DataFrame(columns=empty_cols)

    grouped = (
        base.groupby(["最終客戶", "商品名稱", "分析年度"], dropna=False, as_index=False)["成交價未稅小計"]
        .sum(min_count=1)
    )

    pivot = (
        grouped.pivot_table(
            index=["最終客戶", "商品名稱"],
            columns="分析年度",
            values="成交價未稅小計",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    if last_period.label not in pivot.columns:
        pivot[last_period.label] = 0.0
    if this_period.label not in pivot.columns:
        pivot[this_period.label] = 0.0

    pivot["去年度金額"] = pd.to_numeric(pivot[last_period.label], errors="coerce").fillna(0.0)
    pivot["今年度金額"] = pd.to_numeric(pivot[this_period.label], errors="coerce").fillna(0.0)
    pivot["Forecast"] = (pivot["去年度金額"] - pivot["今年度金額"]).clip(lower=0) * 0.8

    last_rows = work[work["_in_last_period"]].copy()
    if not last_rows.empty:
        key_last_total = (
            last_rows.groupby(["最終客戶", "商品名稱"], dropna=False)["成交價未稅小計"]
            .sum(min_count=1)
            .reset_index()
            .rename(columns={"成交價未稅小計": "_key_last_total"})
        )
        forecast_map = pivot[["最終客戶", "商品名稱", "Forecast"]].rename(columns={"Forecast": "_key_forecast"})
        last_rows = last_rows.merge(forecast_map, on=["最終客戶", "商品名稱"], how="left")
        last_rows = last_rows.merge(key_last_total, on=["最終客戶", "商品名稱"], how="left")
        last_rows["_key_forecast"] = pd.to_numeric(last_rows["_key_forecast"], errors="coerce").fillna(0.0)
        last_rows["_key_last_total"] = pd.to_numeric(last_rows["_key_last_total"], errors="coerce").fillna(0.0)
        last_rows["成交價未稅小計"] = pd.to_numeric(last_rows["成交價未稅小計"], errors="coerce").fillna(0.0)
        last_rows["_row_forecast"] = np.where(
            last_rows["_key_last_total"] > 0,
            last_rows["_key_forecast"] * (last_rows["成交價未稅小計"] / last_rows["_key_last_total"]),
            0.0,
        )
        row_assign = last_rows[["__row_id", "_row_forecast"]].copy()
        df = df.merge(row_assign, on="__row_id", how="left")
        df["Forecast"] = pd.to_numeric(df["_row_forecast"], errors="coerce").fillna(df["Forecast"])
        df["Forecast"] = pd.to_numeric(df["Forecast"], errors="coerce").fillna(0.0)
        df = df.drop(columns=["_row_forecast"], errors="ignore")

        slot_base = (
            last_rows.groupby(
                ["最終客戶", "商品名稱", "_quarter_short", "_fiscal_month_order", "_fiscal_month_label"],
                dropna=False,
                as_index=False,
            )["成交價未稅小計"]
            .sum(min_count=1)
        )
        slot_base["成交價未稅小計"] = pd.to_numeric(slot_base["成交價未稅小計"], errors="coerce").fillna(0.0)
        slot_base = slot_base.sort_values(
            by=["最終客戶", "商品名稱", "成交價未稅小計", "_fiscal_month_order"],
            ascending=[True, True, False, True],
        )
        slot_pick = slot_base.drop_duplicates(subset=["最終客戶", "商品名稱"], keep="first")
    else:
        slot_pick = pd.DataFrame(columns=["最終客戶", "商品名稱", "_quarter_short", "_fiscal_month_order", "_fiscal_month_label"])

    forecast_key_df = pivot[["最終客戶", "商品名稱", "去年度金額", "今年度金額", "Forecast"]].copy()
    forecast_key_df = forecast_key_df.merge(
        slot_pick[["最終客戶", "商品名稱", "_quarter_short", "_fiscal_month_order", "_fiscal_month_label"]],
        on=["最終客戶", "商品名稱"],
        how="left",
    )

    df = df.drop(columns=["__row_id"], errors="ignore")

    # 訂閱到期日 <= 今天 → Forecast 歸零（續約機會已過）
    today_ts = pd.Timestamp(date.today())
    if "訂閱到期日" in df.columns:
        expired = df["訂閱到期日"].notna() & (df["訂閱到期日"] <= today_ts)
        df.loc[expired, "Forecast"] = 0.0

    # forecast_key_df 同步：依 _fiscal_month_label 判斷是否已過期
    if not forecast_key_df.empty and "_fiscal_month_label" in forecast_key_df.columns:
        exp_month = pd.to_datetime(forecast_key_df["_fiscal_month_label"], format="%Y-%m", errors="coerce")
        # 月份的最後一天 <= 今天 → 該月已完全過期
        exp_month_end = exp_month + pd.offsets.MonthEnd(0)
        forecast_key_df.loc[exp_month_end.notna() & (exp_month_end <= today_ts), "Forecast"] = 0.0

    return df, forecast_key_df


def build_kpi_summary(df: pd.DataFrame, period: AnalysisPeriod) -> dict:
    part = df[build_period_mask(df, period)].copy()
    return {
        "筆數": int(len(part)),
        "最終客戶數": int(part["最終客戶"].nunique()) if "最終客戶" in part.columns else 0,
        "經銷商數": int(part["經銷商"].nunique()) if "經銷商" in part.columns else 0,
        "金額合計": float(pd.to_numeric(part.get("成交價未稅小計", 0), errors="coerce").fillna(0).sum()),
    }


# ── Product tier & Upsell Motion helpers ────────────────────────────────────
_TIER_RULES: list[tuple[str, str]] = [
    ("business premium",  "BP"),
    ("business standard", "BS"),
    ("business basic",    "BB"),
    ("copilot",           "Copilot"),
    (" e5",               "ME5"),
    ("e5 ",               "ME5"),
    (" e3",               "ME3"),
    ("e3 ",               "ME3"),
    ("office 365",        "O365"),
    ("microsoft 365",     "M365"),
]


def _product_tier(name: str) -> str:
    if pd.isna(name):
        return "Other"
    p = str(name).lower()
    for pattern, code in _TIER_RULES:
        if pattern in p:
            return code
    return "Other"


def _upsell_motion(exp_prod: str, ren_prod: str) -> str:
    """Classify Upsell Motion based on product tier change."""
    if pd.isna(exp_prod) or pd.isna(ren_prod):
        return "Other Upsell"
    if str(exp_prod).strip() == str(ren_prod).strip():
        return "Expansion"
    et, rt = _product_tier(exp_prod), _product_tier(ren_prod)
    if et == "BB" and rt in ("BS", "BP"):    return "Basic to Prem. or Std."
    if et == "BS" and rt == "BP":            return "Std. to Prem."
    if et == "ME3" and rt == "ME5":          return "ME3 to ME5"
    if et == "O365" and rt in ("M365", "ME3", "ME5"): return "O to M"
    return "Other Upsell"


def build_csp_renewal_metrics(
    df: pd.DataFrame,
    last_period: AnalysisPeriod,
    this_period: AnalysisPeriod,
) -> dict:
    """
    計算微軟 CSP Renewal Dashboard 全套標準指標。
    Expiration = last_period（訂閱到期日）；Renewal = this_period
    Is Renewed = (最終客戶, 商品名稱) 於今年度出現
    Annualized Revenue = 成交單價未稅 × 數量 × 12
    TTM 系列以 Expiration Initial（首筆）為分母；
    Annualized 系列以 Expiration Ending（末筆）為分母。
    """
    req = ["最終客戶", "商品名稱", "成交價未稅小計", "數量", "訂閱到期日", "訂單下單日"]
    if not all(c in df.columns for c in req):
        return {}

    last_mask = build_period_mask(df, last_period)
    this_mask = build_period_mask(df, this_period)
    last_df = df[last_mask].copy()
    this_df = df[this_mask].copy()

    if last_df.empty:
        return {}

    for d in (last_df, this_df):
        d["_ann_rev"] = (
            pd.to_numeric(d["成交單價未稅"], errors="coerce").fillna(0)
            * pd.to_numeric(d["數量"], errors="coerce").fillna(0)
            * 12
        )
        d["成交價未稅小計"] = pd.to_numeric(d["成交價未稅小計"], errors="coerce").fillna(0)
        d["數量"] = pd.to_numeric(d["數量"], errors="coerce").fillna(0)

    exp_billed = float(last_df["成交價未稅小計"].sum())

    last_sorted = last_df.sort_values("訂單下單日")
    exp_ending = last_sorted.groupby(["最終客戶", "商品名稱"], dropna=False).last().reset_index()
    exp_ending_ann   = float(exp_ending["_ann_rev"].sum())
    exp_ending_seats = float(exp_ending["數量"].sum())

    if not this_df.empty:
        this_keys = set(
            zip(this_df["最終客戶"].astype(str), this_df["商品名稱"].astype(str))
        )
    else:
        this_keys = set()

    exp_ending["_is_renewed"] = exp_ending.apply(
        lambda r: (str(r["最終客戶"]), str(r["商品名稱"])) in this_keys, axis=1
    )

    renewed_pairs = exp_ending.loc[exp_ending["_is_renewed"], ["最終客戶", "商品名稱"]]

    if not renewed_pairs.empty and not this_df.empty:
        ren_df = this_df.merge(renewed_pairs, on=["最終客戶", "商品名稱"])
        ren_billed = float(ren_df["成交價未稅小計"].sum())
        ren_initial = (
            ren_df.sort_values("訂單下單日")
            .groupby(["最終客戶", "商品名稱"], dropna=False).first().reset_index()
        )
        ren_initial_ann   = float(ren_initial["_ann_rev"].sum())
        ren_initial_seats = float(ren_initial["數量"].sum())
        exp_renewed_only = exp_ending[exp_ending["_is_renewed"]]
        exp_renewed_billed = float(
            last_df.merge(renewed_pairs, on=["最終客戶", "商品名稱"])["成交價未稅小計"].sum()
        )
        exp_renewed_ending_ann = float(exp_renewed_only["_ann_rev"].sum())
    else:
        ren_billed = ren_initial_ann = ren_initial_seats = 0.0
        exp_renewed_billed = exp_renewed_ending_ann = 0.0

    def _safe_div(n, d):
        return n / d if d > 0 else 0.0

    csp_renewal_rate        = _safe_div(ren_billed,       exp_billed)
    csp_recapture_rate      = _safe_div(ren_billed,       exp_renewed_billed)
    annualized_renewal_rate = _safe_div(ren_initial_ann,  exp_ending_ann)
    annualized_recapture_rate = _safe_div(ren_initial_ann, exp_renewed_ending_ann)
    seats_renewal_rate      = _safe_div(ren_initial_seats, exp_ending_seats)

    cust_table = exp_ending[
        ["最終客戶", "商品名稱", "數量", "_ann_rev", "成交價未稅小計", "_is_renewed"]
    ].copy().rename(columns={
        "數量":          "到期席位數",
        "_ann_rev":     "到期年化收入",
        "成交價未稅小計": "到期計費收入",
        "_is_renewed":  "已續約",
    })

    if not this_df.empty and not renewed_pairs.empty:
        _this_init = (
            this_df.sort_values("訂單下單日")
            .groupby(["最終客戶", "商品名稱"], dropna=False).first().reset_index()
        )[["最終客戶", "商品名稱", "數量", "_ann_rev"]].rename(columns={
            "數量":      "續約席位數",
            "_ann_rev": "續約年化收入",
        })
        cust_table = cust_table.merge(_this_init, on=["最終客戶", "商品名稱"], how="left")
    else:
        cust_table["續約席位數"]  = float("nan")
        cust_table["續約年化收入"] = float("nan")

    if "續約年化收入" in cust_table.columns and "到期年化收入" in cust_table.columns:
        cust_table["年化回收率"] = (
            cust_table["續約年化收入"].fillna(0)
            / cust_table["到期年化收入"].replace(0, float("nan"))
        )
    else:
        cust_table["年化回收率"] = float("nan")

    cust_table = cust_table.sort_values("到期計費收入", ascending=False).reset_index(drop=True)

    return {
        "exp_billed":               exp_billed,
        "exp_ending_ann":           exp_ending_ann,
        "exp_ending_seats":         exp_ending_seats,
        "ren_billed":               ren_billed,
        "ren_initial_ann":          ren_initial_ann,
        "ren_initial_seats":        ren_initial_seats,
        "csp_renewal_rate":         csp_renewal_rate,
        "csp_recapture_rate":       csp_recapture_rate,
        "annualized_renewal_rate":  annualized_renewal_rate,
        "annualized_recapture_rate": annualized_recapture_rate,
        "seats_renewal_rate":       seats_renewal_rate,
        "renewed_count":            int(exp_ending["_is_renewed"].sum()),
        "total_subs":               len(exp_ending),
        "cust_table":               cust_table,
    }

# ── FULL build_csp_renewal_metrics replacement (inserted below) ─────────────
def _build_csp_metrics_full(
    df: pd.DataFrame,
    last_period: AnalysisPeriod,
    this_period: AnalysisPeriod,
) -> dict:
    """完整版 CSP 指標計算：包含 TTM / Price-Qty 分解 / Upsell Motion / Copilot。"""
    req = ["最終客戶", "商品名稱", "成交價未稅小計", "數量", "訂閱到期日", "訂單下單日"]
    if not all(c in df.columns for c in req):
        return {}

    last_mask = build_period_mask(df, last_period)
    this_mask = build_period_mask(df, this_period)
    last_df = df[last_mask].copy()
    this_df = df[this_mask].copy()
    if last_df.empty:
        return {}

    def _num(d, col):
        return pd.to_numeric(d[col], errors="coerce").fillna(0) if col in d.columns else pd.Series(0.0, index=d.index)

    def _div(n, d):
        return float(n) / float(d) if d and float(d) != 0 else 0.0

    for d in (last_df, this_df):
        d["數量"]           = _num(d, "數量")
        d["成交價未稅小計"]  = _num(d, "成交價未稅小計")
        d["成交單價未稅"]    = _num(d, "成交單價未稅") if "成交單價未稅" in d.columns else 0.0
        d["_ann_rev"]       = d["成交單價未稅"] * d["數量"] * 12

    # ── Expiration ────────────────────────────────────────────────────────
    exp_billed = float(last_df["成交價未稅小計"].sum())
    last_sorted = last_df.sort_values("訂單下單日")

    exp_init = (last_sorted.groupby(["最終客戶", "商品名稱"], dropna=False)
                .first().reset_index())
    exp_end  = (last_sorted.groupby(["最終客戶", "商品名稱"], dropna=False)
                .last().reset_index())

    exp_initial_ann    = float(exp_init["_ann_rev"].sum())
    exp_initial_seats  = float(exp_init["數量"].sum())
    exp_ending_ann     = float(exp_end["_ann_rev"].sum())
    exp_ending_seats   = float(exp_end["數量"].sum())

    # ── Is Renewed (customer+product key) ─────────────────────────────────
    this_keys = (set(zip(this_df["最終客戶"].astype(str), this_df["商品名稱"].astype(str)))
                 if not this_df.empty else set())
    exp_end["_is_renewed"] = exp_end.apply(
        lambda r: (str(r["最終客戶"]), str(r["商品名稱"])) in this_keys, axis=1)
    exp_init["_is_renewed"] = exp_end["_is_renewed"].values
    renewed_pairs = exp_end.loc[exp_end["_is_renewed"], ["最終客戶", "商品名稱"]]
    renewed_count = int(exp_end["_is_renewed"].sum())

    # ── Renewal ───────────────────────────────────────────────────────────
    if not renewed_pairs.empty and not this_df.empty:
        ren_all  = this_df.merge(renewed_pairs, on=["最終客戶", "商品名稱"])
        ren_billed = float(ren_all["成交價未稅小計"].sum())
        ren_init = (this_df.sort_values("訂單下單日")
                    .merge(renewed_pairs, on=["最終客戶", "商品名稱"])
                    .groupby(["最終客戶", "商品名稱"], dropna=False).first().reset_index())
        ren_initial_ann   = float(ren_init["_ann_rev"].sum())
        ren_initial_seats = float(ren_init["數量"].sum())

        exp_ren_end   = exp_end[exp_end["_is_renewed"]]
        exp_ren_init  = exp_init[exp_init["_is_renewed"]]
        exp_ren_billed          = float(last_df.merge(renewed_pairs, on=["最終客戶", "商品名稱"])["成交價未稅小計"].sum())
        exp_ren_ending_ann      = float(exp_ren_end["_ann_rev"].sum())
        exp_ren_ending_seats    = float(exp_ren_end["數量"].sum())
        exp_ren_initial_ann     = float(exp_ren_init["_ann_rev"].sum())
        exp_ren_initial_seats   = float(exp_ren_init["數量"].sum())
    else:
        ren_billed = ren_initial_ann = ren_initial_seats = 0.0
        exp_ren_billed = exp_ren_ending_ann = exp_ren_ending_seats = 0.0
        exp_ren_initial_ann = exp_ren_initial_seats = 0.0
        ren_init = pd.DataFrame()

    # ── Standard rates ────────────────────────────────────────────────────
    csp_renewal_rate          = _div(ren_billed,       exp_billed)
    csp_recapture_rate        = _div(ren_billed,       exp_ren_billed)
    annualized_renewal_rate   = _div(ren_initial_ann,  exp_ending_ann)
    annualized_recapture_rate = _div(ren_initial_ann,  exp_ren_ending_ann)
    seats_renewal_rate        = _div(ren_initial_seats, exp_ending_seats)
    seats_recapture_rate      = _div(ren_initial_seats, exp_ren_ending_seats)

    # ── TTM rates (Initial denominator) ───────────────────────────────────
    ttm_ann_renewal_rate   = _div(ren_initial_ann,  exp_initial_ann)
    ttm_ann_recapture_rate = _div(ren_initial_ann,  exp_ren_initial_ann)
    ttm_arr_growth         = ren_initial_ann   - exp_initial_ann
    ttm_arr_growth_seats   = ren_initial_seats - exp_initial_seats
    ttm_seats_renewal_rate = _div(ren_initial_seats, exp_initial_seats)

    # ── Price / Quantity decomposition (on renewed subs) ─────────────────
    q0   = exp_ren_initial_seats
    q1   = ren_initial_seats
    p0   = _div(exp_ren_initial_ann, q0)
    p1   = _div(ren_initial_ann, q1)
    _p   = (p1 - p0) * q0
    _q   = (q1 - q0) * p0
    pxq  = (p1 - p0) * (q1 - q0)
    _pq  = _p + _q
    price_effect    = _p + (pxq * _div(_p, _pq) if _pq else pxq / 2)
    quantity_effect = _q + (pxq * _div(_q, _pq) if _pq else pxq / 2)
    arr_growth_base = ren_initial_ann - exp_ren_initial_ann
    price_pct    = _div(price_effect,    arr_growth_base)
    quantity_pct = _div(quantity_effect, arr_growth_base)

    # ── Customer-level primary product (for Upsell Motion) ────────────────
    def _primary_product(src_df: pd.DataFrame, grp_col: str, val_col: str, name_col: str) -> pd.DataFrame:
        return (src_df.groupby([grp_col, name_col], dropna=False)[val_col]
                .sum().reset_index()
                .sort_values(val_col, ascending=False)
                .drop_duplicates(subset=[grp_col])
                [[grp_col, name_col]])

    _last_prim = _primary_product(last_df, "最終客戶", "成交價未稅小計", "商品名稱").rename(columns={"商品名稱": "_exp_prod"})
    _this_prim = (_primary_product(this_df, "最終客戶", "成交價未稅小計", "商品名稱").rename(columns={"商品名稱": "_ren_prod"})
                  if not this_df.empty else pd.DataFrame(columns=["最終客戶", "_ren_prod"]))
    _motion_map = (_last_prim.merge(_this_prim, on="最終客戶", how="left")
                   .assign(Upsell_Motion=lambda r: r.apply(
                       lambda x: _upsell_motion(x["_exp_prod"], x.get("_ren_prod", float("nan"))), axis=1)))

    # ── Build customer-product detail table ───────────────────────────────
    ct = exp_end[["最終客戶", "商品名稱", "數量", "_ann_rev", "成交價未稅小計", "_is_renewed"]].copy()
    ct = ct.rename(columns={"數量": "到期席位數", "_ann_rev": "到期年化收入",
                             "成交價未稅小計": "到期計費收入", "_is_renewed": "已續約"})
    ct["到期初始年化收入"] = exp_init["_ann_rev"].values
    ct["到期初始席位數"]   = exp_init["數量"].values

    if not ren_init.empty:
        _rj = ren_init[["最終客戶", "商品名稱", "數量", "_ann_rev"]].rename(
            columns={"數量": "續約初始席位數", "_ann_rev": "續約初始年化收入"})
        ct = ct.merge(_rj, on=["最終客戶", "商品名稱"], how="left")
    else:
        ct["續約初始席位數"] = float("nan")
        ct["續約初始年化收入"] = float("nan")

    ct["年化回收率（Ending）"] = ct["續約初始年化收入"].fillna(0) / ct["到期年化收入"].replace(0, float("nan"))
    ct["TTM年化回收率（Initial）"] = ct["續約初始年化收入"].fillna(0) / ct["到期初始年化收入"].replace(0, float("nan"))
    ct["席位變化"] = ct["續約初始席位數"].fillna(0) - ct["到期席位數"]

    ct = ct.merge(_motion_map[["最終客戶", "Upsell_Motion"]], on="最終客戶", how="left")

    # Override same-product renewals with seat-change classification
    same_ren = ct["已續約"] & (ct["Upsell_Motion"] == "Expansion")
    ct.loc[same_ren & (ct["席位變化"] > 0),  "Upsell_Motion"] = "Expansion"
    ct.loc[same_ren & (ct["席位變化"] == 0), "Upsell_Motion"] = "Pure Renewal"
    ct.loc[same_ren & (ct["席位變化"] < 0),  "Upsell_Motion"] = "Reduction"
    ct.loc[~ct["已續約"],                    "Upsell_Motion"] = "Not Renewed"

    ct["Upsell Summary"] = ct["Upsell_Motion"].map(
        lambda m: "Expansion" if m == "Expansion"
        else ("Upsell" if m not in ("Not Renewed", "Pure Renewal", "Reduction", "Not Renewed") else m)
    )
    ct = ct.rename(columns={"Upsell_Motion": "Upsell Motion"})

    # ── Expansion / Upsell rates ──────────────────────────────────────────
    def _motion_rate(motion_filter) -> float:
        _keys = ct.loc[motion_filter, ["最終客戶", "商品名稱"]]
        if _keys.empty or this_df.empty:
            return 0.0
        _r = this_df.merge(_keys, on=["最終客戶", "商品名稱"])["成交價未稅小計"].sum()
        _e = last_df.merge(_keys, on=["最終客戶", "商品名稱"])["成交價未稅小計"].sum()
        return _div(_r, _e)

    expansion_renewal_rate = _motion_rate(ct["Upsell Motion"] == "Expansion")
    upsell_renewal_rate    = _motion_rate(ct["已續約"] & ~ct["Upsell Motion"].isin(["Expansion", "Pure Renewal"]))

    # Motion summary
    motion_summary = (
        ct.groupby("Upsell Motion", dropna=False)
        .agg(訂閱數=("已續約", "count"),
             到期計費收入=("到期計費收入", "sum"),
             到期年化收入=("到期年化收入", "sum"),
             續約年化收入=("續約初始年化收入", lambda s: s.fillna(0).sum()))
        .reset_index()
    )
    motion_summary["年化回收率"] = (
        motion_summary["續約年化收入"] / motion_summary["到期年化收入"].replace(0, float("nan"))
    )

    # ── Copilot metrics ───────────────────────────────────────────────────
    if not this_df.empty and not ren_init.empty:
        _cop = (ren_init["商品名稱"].astype(str).str.lower().str.contains("copilot", na=False) &
                ~ren_init["商品名稱"].astype(str).str.lower().str.contains("copilot studio", na=False))
        copilot_renewal_seats = float(ren_init.loc[_cop, "數量"].fillna(0).sum())
        copilot_accounts      = int(ren_init.loc[_cop, "最終客戶"].nunique())
    else:
        copilot_renewal_seats = 0.0
        copilot_accounts      = 0

    _elig = exp_end["商品名稱"].astype(str).str.lower().str.contains(
        "business premium|business standard|m365 e3|m365 e5|microsoft 365 e3|microsoft 365 e5|o365 e3|o365 e5",
        regex=True, na=False)
    copilot_eligible_seats = float(exp_end.loc[_elig, "數量"].fillna(0).sum())
    copilot_attach_rate    = _div(copilot_renewal_seats, copilot_eligible_seats)

    ct = ct.sort_values("到期計費收入", ascending=False).reset_index(drop=True)

    return {
        # Expiration
        "exp_billed":               exp_billed,
        "exp_ending_ann":           exp_ending_ann,
        "exp_ending_seats":         exp_ending_seats,
        "exp_initial_ann":          exp_initial_ann,
        "exp_initial_seats":        exp_initial_seats,
        # Renewal
        "ren_billed":               ren_billed,
        "ren_initial_ann":          ren_initial_ann,
        "ren_initial_seats":        ren_initial_seats,
        # Standard rates
        "csp_renewal_rate":         csp_renewal_rate,
        "csp_recapture_rate":       csp_recapture_rate,
        "annualized_renewal_rate":  annualized_renewal_rate,
        "annualized_recapture_rate": annualized_recapture_rate,
        "seats_renewal_rate":       seats_renewal_rate,
        "seats_recapture_rate":     seats_recapture_rate,
        # TTM rates
        "ttm_ann_renewal_rate":     ttm_ann_renewal_rate,
        "ttm_ann_recapture_rate":   ttm_ann_recapture_rate,
        "ttm_arr_growth":           ttm_arr_growth,
        "ttm_arr_growth_seats":     ttm_arr_growth_seats,
        "ttm_seats_renewal_rate":   ttm_seats_renewal_rate,
        # Price / Quantity decomposition
        "price_effect":             price_effect,
        "quantity_effect":          quantity_effect,
        "price_pct":                price_pct,
        "quantity_pct":             quantity_pct,
        "arr_growth_base":          arr_growth_base,
        # Upsell Motion
        "expansion_renewal_rate":   expansion_renewal_rate,
        "upsell_renewal_rate":      upsell_renewal_rate,
        "motion_summary":           motion_summary,
        # Copilot
        "copilot_renewal_seats":    copilot_renewal_seats,
        "copilot_accounts":         copilot_accounts,
        "copilot_attach_rate":      copilot_attach_rate,
        "copilot_eligible_seats":   copilot_eligible_seats,
        # Summary
        "renewed_count":            renewed_count,
        "total_subs":               len(exp_end),
        "cust_table":               ct,
    }


def build_month_summary(
    df: pd.DataFrame,
    forecast_key_df: pd.DataFrame,
    last_period: AnalysisPeriod,
    this_period: AnalysisPeriod,
    selected_month_orders: list[int] | None = None,
) -> pd.DataFrame:
    month_map = pd.DataFrame({"_fiscal_month_order": list(range(1, 13)), "期間": month_labels_for_period(last_period)})
    last_df = df[build_period_mask(df, last_period)].copy()
    this_df = df[build_period_mask(df, this_period)].copy()

    last_grp = last_df.groupby("_fiscal_month_order", as_index=False)["成交價未稅小計"].sum(min_count=1).rename(columns={"成交價未稅小計": "去年度金額"})
    this_grp = this_df.groupby("_fiscal_month_order", as_index=False)["成交價未稅小計"].sum(min_count=1).rename(columns={"成交價未稅小計": "今年度金額"})
    fc_grp = forecast_key_df.groupby("_fiscal_month_order", as_index=False)["Forecast"].sum(min_count=1)

    summary = month_map.merge(last_grp, on="_fiscal_month_order", how="left").merge(this_grp, on="_fiscal_month_order", how="left").merge(fc_grp, on="_fiscal_month_order", how="left").fillna(0)
    summary["差異"] = summary["今年度金額"] - summary["去年度金額"]

    if selected_month_orders:
        summary = summary[summary["_fiscal_month_order"].isin(selected_month_orders)].copy()
    return summary


def build_quarter_summary(
    df: pd.DataFrame,
    forecast_key_df: pd.DataFrame,
    last_period: AnalysisPeriod,
    this_period: AnalysisPeriod,
    selected_quarters: list[str] | None = None,
) -> pd.DataFrame:
    base_q = pd.DataFrame({"期間": QUARTER_ORDER})
    last_df = df[build_period_mask(df, last_period)].copy()
    this_df = df[build_period_mask(df, this_period)].copy()

    last_grp = last_df.groupby("_quarter_short", as_index=False)["成交價未稅小計"].sum(min_count=1).rename(columns={"_quarter_short": "期間", "成交價未稅小計": "去年度金額"})
    this_grp = this_df.groupby("_quarter_short", as_index=False)["成交價未稅小計"].sum(min_count=1).rename(columns={"_quarter_short": "期間", "成交價未稅小計": "今年度金額"})
    fc_grp = forecast_key_df.groupby("_quarter_short", as_index=False)["Forecast"].sum(min_count=1).rename(columns={"_quarter_short": "期間"})

    summary = base_q.merge(last_grp, on="期間", how="left").merge(this_grp, on="期間", how="left").merge(fc_grp, on="期間", how="left").fillna(0)
    summary["差異"] = summary["今年度金額"] - summary["去年度金額"]

    if selected_quarters:
        summary = summary[summary["期間"].isin(selected_quarters)].copy()
    return summary


def build_period_summary(
    df: pd.DataFrame,
    forecast_key_df: pd.DataFrame,
    mode: str,
    last_period: AnalysisPeriod,
    this_period: AnalysisPeriod,
    selected_values: list[str] | list[int] | None = None,
) -> pd.DataFrame:
    if mode == "Month":
        return build_month_summary(df, forecast_key_df, last_period, this_period, selected_values if selected_values else None)
    return build_quarter_summary(df, forecast_key_df, last_period, this_period, selected_values if selected_values else None)


def build_trend_figure(summary_df: pd.DataFrame, mode: str, last_period: AnalysisPeriod, this_period: AnalysisPeriod) -> go.Figure:
    if summary_df.empty:
        return go.Figure()

    x_values = summary_df["期間"].tolist()
    last_vals = pd.to_numeric(summary_df["去年度金額"], errors="coerce").fillna(0.0).tolist()
    this_vals = pd.to_numeric(summary_df["今年度金額"], errors="coerce").fillna(0.0).tolist()
    forecast_vals = pd.to_numeric(summary_df["Forecast"], errors="coerce").fillna(0.0).tolist()
    diff_vals = pd.to_numeric(summary_df["差異"], errors="coerce").fillna(0.0).tolist()

    title = "四季趨勢（去年度 / 今年度 / Forecast）" if mode == "Quarter" else "月份趨勢（去年度 / 今年度 / Forecast）"
    x_title = "季度" if mode == "Quarter" else "月份"

    row_heights = [0.56, 0.44] if mode == "Month" else [0.66, 0.34]
    v_spacing = 0.26 if mode == "Month" else 0.28
    fig = make_subplots(
        rows=2,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=v_spacing,
        row_heights=row_heights,
        subplot_titles=(title, ""),   # 下方標題改用 annotation 手動分層
    )

    _fmt_v = lambda v: f"{v/1e6:.1f}M" if abs(v) >= 1e6 else f"{v:,.0f}"
    _fmt = lambda vals: [_fmt_v(v) for v in vals]
    _nonzero_text = lambda vals: [t if v != 0 else "" for v, t in zip(vals, _fmt(vals))]
    # Quarter 模式直接用 bar text 標示；Month 模式改用 annotation 框（避免旋轉/太小）
    _upper_bar_text = (lambda vals: _nonzero_text(vals)) if mode == "Quarter" else (lambda vals: [""] * len(vals))
    _fs_main = 11
    _fs_diff = 11

    # 上方主圖：去年度獨立柱；今年度柱 + Forecast 堆疊於其上
    fig.add_trace(
        go.Bar(
            name=last_period.label,
            x=x_values,
            y=last_vals,
            marker_color="#F97316",
            offsetgroup="last",
            legendgroup="last",
            text=_upper_bar_text(last_vals),
            textposition="outside",
            textangle=0,
            textfont=dict(size=_fs_main, color="#F97316"),
            cliponaxis=False,
            hovertemplate=f"{last_period.label}<br>%{{x}}<br>%{{y:,.0f}}<extra></extra>",
        ),
        row=1,
        col=1,
    )
    fig.add_trace(
        go.Bar(
            name=this_period.label,
            x=x_values,
            y=this_vals,
            marker_color="#1D4ED8",
            offsetgroup="this",
            legendgroup="this",
            text=_upper_bar_text(this_vals),
            textposition="outside",
            textangle=0,
            textfont=dict(size=_fs_main, color="#1D4ED8"),
            cliponaxis=False,
            hovertemplate=f"{this_period.label}<br>%{{x}}<br>%{{y:,.0f}}<extra></extra>",
        ),
        row=1,
        col=1,
    )
    fig.add_trace(
        go.Bar(
            name="Forecast",
            x=x_values,
            y=forecast_vals,
            marker_color="#93C5FD",
            offsetgroup="this",
            legendgroup="forecast",
            text=_upper_bar_text(forecast_vals),
            textposition="outside",
            textangle=0,
            textfont=dict(size=_fs_main, color="#1E40AF"),
            cliponaxis=False,
            hovertemplate="Forecast<br>%{x}<br>%{y:,.0f}<extra></extra>",
        ),
        row=1,
        col=1,
    )

    diff_colors = ["#059669" if v >= 0 else "#DC2626" for v in diff_vals]

    # Month：bar 不顯示文字，改用 annotation 強制大字；Quarter：直接用 bar text
    bar_text = [""] * len(diff_vals) if mode == "Month" else _nonzero_text(diff_vals)
    fig.add_trace(
        go.Bar(
            name="差異",
            x=x_values,
            y=diff_vals,
            marker_color=diff_colors,
            text=bar_text,
            textposition="outside",
            textangle=0,
            textfont=dict(size=_fs_diff, color=diff_colors),
            cliponaxis=False,
            showlegend=False,
            hovertemplate="差異<br>%{x}<br>%{y:,.0f}<extra></extra>",
        ),
        row=2,
        col=1,
    )

    if mode == "Month":
        # ── 下方差異圖 annotation ──
        # 相鄰月份若差值相近會堆疊，改以奇偶 index 交錯 yshift 避免重疊
        for i, (x_val, y_val, color) in enumerate(zip(x_values, diff_vals, diff_colors)):
            if y_val == 0:
                continue
            _base_shift = 5
            _extra = 14 if i % 2 == 1 else 0   # 奇數月份多往外推一層
            _yshift = (_base_shift + _extra) if y_val >= 0 else -(_base_shift + _extra)
            fig.add_annotation(
                x=x_val, y=y_val,
                text=f"<b>{_fmt_v(y_val)}</b>",
                xref="x2", yref="y2",
                font=dict(size=11, color=color),
                showarrow=False,
                yanchor="bottom" if y_val >= 0 else "top",
                yshift=_yshift,
                bgcolor="white", bordercolor=color, borderwidth=1,
            )

        # ── 上方主圖 annotation（以 xshift 區分左右柱群）──
        def _add_main_ann(x, y, text, color, xshift, extra_yshift=0):
            if y == 0:
                return
            fig.add_annotation(
                x=x, y=y,
                text=f"<b>{text}</b>",
                xref="x", yref="y",
                xshift=xshift,
                font=dict(size=11, color=color),
                showarrow=False,
                yanchor="bottom",
                yshift=5 + extra_yshift,
                bgcolor="white", bordercolor=color, borderwidth=1,
            )

        for i, (x_val, l_val, t_val, f_val) in enumerate(zip(x_values, last_vals, this_vals, forecast_vals)):
            # 奇偶月份交錯：讓兩組標籤在 y 軸方向錯開，降低相鄰月份標籤重疊機率
            _stagger = 18 if i % 2 == 1 else 0
            # 去年度：左側柱群
            _add_main_ann(x_val, l_val, _fmt_v(l_val), "#F97316", xshift=-20, extra_yshift=_stagger)
            # 今年度：右側柱群（顯示在 今年度 頂端）
            _add_main_ann(x_val, t_val + f_val, _fmt_v(t_val), "#1D4ED8", xshift=20, extra_yshift=_stagger)
            # Forecast：疊加在今年度上方，若非零則再顯示一層
            if f_val != 0:
                _add_main_ann(x_val, t_val + f_val, _fmt_v(f_val), "#1E40AF", xshift=20, extra_yshift=22 + _stagger)

    chart_height = 1050 if mode == "Month" else 680
    fig.update_layout(
        barmode="relative",
        height=chart_height,
        margin=dict(l=20, r=20, t=80, b=40),
        legend_title_text="年度",
        plot_bgcolor="white",
        paper_bgcolor="white",
        bargap=0.25,
        bargroupgap=0.08,
    )
    fig.update_yaxes(title_text="金額", tickformat=",", row=1, col=1, gridcolor="#E5E7EB")
    fig.update_yaxes(title_text="差異（今年-去年）", tickformat=",", zeroline=True, zerolinecolor="#9CA3AF", gridcolor="#F3F4F6", row=2, col=1)
    # 隱藏兩個子圖的軸刻度，改由 annotation 完全控制位置
    fig.update_xaxes(showgrid=False, showticklabels=False, row=1, col=1)
    fig.update_xaxes(showgrid=False, showticklabels=False, side="top", row=2, col=1)

    # ── 7 等分間隔，由上而下放在第 2、4、6 格中央 ──────────────────
    _avail        = 1.0 - v_spacing
    _r2_top_paper = row_heights[1] * _avail          # 間隔底端（下方子圖頂端）
    _seg          = v_spacing / 7                    # 每一格高度

    # 第 2 格中央（由上往下）→ 月份 / 季度 標題
    _y_xtitle  = _r2_top_paper + _seg * (7 - 1.5)   # = + 5.5 * seg
    # 第 4 格中央 → 刻度標籤
    _y_ticks   = _r2_top_paper + _seg * (7 - 3.5)   # = + 3.5 * seg
    # 第 6 格中央 → 去年度與今年度差異
    _y_diff    = _r2_top_paper + _seg * (7 - 5.5)   # = + 1.5 * seg

    # 月份 / 季度 標題
    fig.add_annotation(
        text=f"<b>{x_title}</b>",
        xref="paper", yref="paper",
        x=0.5, y=_y_xtitle,
        showarrow=False,
        font=dict(size=16, color="#444444"),
        xanchor="center", yanchor="middle",
    )

    # 刻度標籤（每個 x 值獨立 annotation，xref="x" 對齊柱子）
    for x_val in x_values:
        fig.add_annotation(
            text=str(x_val),
            xref="x", yref="paper",
            x=x_val, y=_y_ticks,
            showarrow=False,
            font=dict(size=13, color="#555555"),
            xanchor="center", yanchor="middle",
        )

    # 去年度與今年度差異 標題
    fig.add_annotation(
        text="去年度與今年度差異",
        xref="paper", yref="paper",
        x=0.5, y=_y_diff,
        showarrow=False,
        font=dict(size=16, color="#444444"),
        xanchor="center", yanchor="middle",
    )
    return fig


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="cleaned_data")
    output.seek(0)
    return output.read()


def render_marketing_kpi_cards(last_metrics: dict, this_metrics: dict) -> None:
    """精準行銷分頁專用：精緻緊湊的 KPI 卡片（去年度 / 今年度 / 差異 三合一）"""
    diff = {k: this_metrics[k] - last_metrics[k] for k in last_metrics}

    cards_cfg = [
        ("📋", "筆數",      "#6366F1", "#EEF2FF", False),
        ("👥", "最終客戶數", "#0EA5E9", "#F0F9FF", False),
        ("🏪", "經銷商數",   "#F59E0B", "#FFFBEB", False),
        ("💰", "金額合計",   "#10B981", "#F0FDF4", True ),
    ]

    cols = st.columns(4)
    for col, (icon, label, accent, bg, is_money) in zip(cols, cards_cfg):
        with col:
            fmt       = fmt_currency if is_money else fmt_int
            last_str  = fmt(last_metrics[label])
            this_str  = fmt(this_metrics[label])
            diff_v    = diff[label]
            diff_str  = fmt(abs(diff_v))
            arrow     = "▲" if diff_v >= 0 else "▼"
            d_color   = "#059669" if diff_v >= 0 else "#DC2626"
            d_bg      = "#D1FAE5" if diff_v >= 0 else "#FEE2E2"

            st.markdown(f"""
<div style="
    background:{bg};
    border-radius:14px;
    border-top:6px solid {accent};
    padding:14px 16px 12px;
    box-shadow:0 2px 10px rgba(0,0,0,0.08);
    margin-bottom:4px;
">
  <!-- 標題 -->
  <div style="display:flex;align-items:center;gap:6px;margin-bottom:10px;">
    <span style="font-size:20px;line-height:1;">{icon}</span>
    <span style="font-size:14px;font-weight:700;color:#374151;letter-spacing:0.3px;">{label}</span>
  </div>
  <!-- 去年度 / 今年度 並排 -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:3px 10px;margin-bottom:10px;">
    <div style="font-size:12px;color:#9CA3AF;font-weight:600;">去年度</div>
    <div style="font-size:12px;color:{accent};font-weight:700;">今年度</div>
    <div style="font-size:18px;font-weight:800;color:#1F2937;white-space:nowrap;">{last_str}</div>
    <div style="font-size:18px;font-weight:800;color:{accent};white-space:nowrap;">{this_str}</div>
  </div>
  <!-- 差異標籤 -->
  <div style="
    display:inline-flex;align-items:center;gap:5px;
    padding:3px 10px;border-radius:20px;
    background:{d_bg};
  ">
    <span style="font-size:12px;font-weight:700;color:{d_color};">{arrow} {diff_str}</span>
    <span style="font-size:11px;color:#6B7280;">YOY</span>
  </div>
</div>
""", unsafe_allow_html=True)


def render_kpi_cards(title: str, metrics: dict, color: str = "#183153", bg_color: str = "#ffffff") -> None:
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    labels = ["筆數", "最終客戶數", "經銷商數", "金額合計"]
    for col, label in zip(cols, labels):
        with col:
            value = fmt_currency(metrics[label]) if label == "金額合計" else fmt_int(metrics[label])
            st.markdown(
                f"""
                <div class="card" style="background:{bg_color};">
                    <div class="kpi-label">{label}</div>
                    <div class="kpi-value" style="color:{color};">{value}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def render_diff_cards(last_metrics: dict, this_metrics: dict) -> None:
    diff = {
        "筆數": this_metrics["筆數"] - last_metrics["筆數"],
        "最終客戶數": this_metrics["最終客戶數"] - last_metrics["最終客戶數"],
        "經銷商數": this_metrics["經銷商數"] - last_metrics["經銷商數"],
        "金額合計": this_metrics["金額合計"] - last_metrics["金額合計"],
    }
    st.markdown("<div class='section-title'>差異（今年度 - 去年度）</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    labels = ["筆數", "最終客戶數", "經銷商數", "金額合計"]
    for col, label in zip(cols, labels):
        with col:
            raw = diff[label]
            arrow = "▲" if raw >= 0 else "▼"
            color = "#059669" if raw >= 0 else "#DC2626"
            bg_color = "#F0FDF4" if raw >= 0 else "#FFF1F2"
            value = fmt_currency(abs(raw)) if label == "金額合計" else fmt_int(abs(raw))
            st.markdown(
                f"""
                <div class="card" style="background:{bg_color};">
                    <div class="kpi-label">{label} YOY</div>
                    <div class="kpi-value" style="color:{color};">{arrow} {value}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def render_main_kpi_cards(last_metrics: dict, this_metrics: dict, this_label: str = "今年度", last_label: str = "去年度") -> None:
    """主要分析分頁專用：單行四卡，每卡整合今年度大值＋去年度＋YOY差異。"""
    cards_cfg = [
        # (icon, label, accent_color, bg_color, unit_suffix, desc, is_money)
        ("📋", "Enrollment 數量",  "#4F6BF4", "#EEF2FF", "筆", f"{this_label}累計交易",   False),
        ("👥", "最終客戶數",       "#0EA5E9", "#F0F9FF", "戶", f"{this_label}主體客戶數", False),
        ("🏪", "經銷商數",         "#F59E0B", "#FFFBEB", "家", f"{this_label}出貨經銷商", False),
        ("💲", "金額合計（未稅）", "#10B981", "#F0FDF4", "",   f"{this_label}累計成交價", True ),
    ]

    cols = st.columns(4)
    for col, (icon, label, accent, bg, unit, desc, is_money) in zip(cols, cards_cfg):
        with col:
            key     = "金額合計" if "金額合計" in label else ("筆數" if "Enrollment" in label else label)
            this_v  = this_metrics[key]
            last_v  = last_metrics[key]
            diff_v  = this_v - last_v
            arrow   = "+" if diff_v >= 0 else "-"
            d_color = "#059669" if diff_v >= 0 else "#DC2626"
            d_bg    = "#D1FAE5" if diff_v >= 0 else "#FEE2E2"

            # ── 純文字數值（不使用 HTML 字串變數）──
            pct      = (diff_v / last_v * 100) if last_v else 0.0
            pct_str  = f"{abs(pct):.1f}%"               # "25.9%"

            if is_money:
                main_num  = fmt_currency(this_v)
                last_disp = f"$ {fmt_currency_compact(last_v)}"
                abs_disp  = f"$ {fmt_currency_compact(abs(diff_v))}"
                num_line  = f"$ {main_num}"
            else:
                main_num  = fmt_int(this_v)
                last_disp = f"{fmt_int(last_v)} {unit}"
                abs_disp  = fmt_int(abs(diff_v))
                num_line  = f"{main_num} {unit}" if unit else main_num

            # YOY badge：+/-百分比 (+/-絕對數)
            yoy_disp = f"{arrow} {pct_str}  ({arrow}{abs_disp})"

            st.markdown(f"""<div style="background:{bg};border-radius:16px;border:1px solid rgba(0,0,0,0.08);padding:18px 18px 14px;box-shadow:0 2px 12px rgba(0,0,0,0.06);min-height:138px;">
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;"><span style="font-size:13px;font-weight:600;color:#6B7280;">{label}</span><span style="font-size:22px;">{icon}</span></div>
<div style="font-size:28px;font-weight:800;color:#111827;line-height:1.1;letter-spacing:-0.5px;margin-bottom:4px;white-space:nowrap;">{num_line}</div>
<div style="font-size:11.5px;color:#9CA3AF;margin-bottom:10px;border-bottom:1px solid rgba(0,0,0,0.07);padding-bottom:10px;">{desc}</div>
<div style="display:flex;justify-content:space-between;align-items:center;">
<div><span style="font-size:11px;color:#9CA3AF;margin-right:4px;">{last_label}</span><span style="font-size:13px;font-weight:700;color:#374151;">{last_disp}</span></div>
<div style="display:inline-flex;align-items:center;gap:3px;padding:3px 10px;border-radius:20px;background:{d_bg};white-space:nowrap;"><span style="font-size:10px;color:#6B7280;font-weight:600;">YOY</span><span style="font-size:12px;font-weight:800;color:{d_color};">{yoy_disp}</span></div>
</div>
</div>""", unsafe_allow_html=True)


# ------------------------------------------------------------
# App
# ------------------------------------------------------------
st.title("Weblink M365 續約概況與精準行銷推廣平台")

# ── 上傳區（頂部常駐，收合後仍可見上傳狀態）──────────────────────────────────
_data_exp = st.expander("📂 資料載入與摘要", expanded=True)
with _data_exp:
    _ul_col, _info_col = st.columns([2, 3])
    with _ul_col:
        uploaded_file = st.file_uploader(
            "上傳 Excel 訂單資料（.xlsx）",
            type=["xlsx"],
            help="例如：CSP訂單資料_raw.xlsx",
        )
    with _info_col:
        st.info(
            "**支援功能**\n"
            "- 資料自動清洗（排除異常、教育、非營利）\n"
            "- FY 年度 / 季度衍生欄位 / Forecast 分析\n"
            "- KPI 趨勢圖 / Top15 / Copilot 滲透率\n"
            "- 精準行銷 90 天示警 / 微軟 CSP 指標（Recapture Rate 等）"
        )
default_last_period, default_this_period = infer_default_periods()
default_local_file_exists = Path(DEFAULT_LOCAL_XLSX_PATH).exists()

with st.sidebar:
    # ── 分析年度選擇器（最上方） ───────────────────────────────────────────
    _base_year_default = default_this_period.start.year   # 今年度起始年，預設 2026
    _year_options = list(range(_base_year_default - 4, _base_year_default + 5))

    def _on_base_year_change():
        _by = st.session_state["base_year"]
        st.session_state["last_range"] = (date(_by - 1, 7, 1), date(_by,     6, 30))
        st.session_state["this_range"] = (date(_by,     7, 1), date(_by + 1, 6, 30))

    st.selectbox(
        "📅 分析年度",
        options=_year_options,
        index=_year_options.index(_base_year_default),
        format_func=lambda y: (
            f"{y}（去 FY{str(y % 100).zfill(2)}｜今 FY{str((y + 1) % 100).zfill(2)}）"
        ),
        key="base_year",
        on_change=_on_base_year_change,
    )
    st.markdown("---")
    st.header("篩選條件")
    show_time_ranges = st.checkbox("顯示時間範圍(以訂閱到期日為基準)", value=False)
    if show_time_ranges:
        _lr_ss = st.session_state.get("last_range", None)
        _tr_ss = st.session_state.get("this_range", None)
        _lp = AnalysisPeriod("去年度", pd.Timestamp(_lr_ss[0]), pd.Timestamp(_lr_ss[1])) if _lr_ss and len(_lr_ss) == 2 else default_last_period
        _tp = AnalysisPeriod("今年度", pd.Timestamp(_tr_ss[0]), pd.Timestamp(_tr_ss[1])) if _tr_ss and len(_tr_ss) == 2 else default_this_period
        def _fy_quarter_months(period):
            sy = period.start.year
            fy = f"FY{str((sy + 1) % 100).zfill(2)}"
            return [
                (f"{fy} Q1", sy,     [7, 8, 9]),
                (f"{fy} Q2", sy,     [10, 11, 12]),
                (f"{fy} Q3", sy + 1, [1, 2, 3]),
                (f"{fy} Q4", sy + 1, [4, 5, 6]),
            ]

        st.info(f"去年度：{_lp.start.date()} ～ {_lp.end.date()}")
        for _ql, _yr, _ms in _fy_quarter_months(_lp):
            st.caption(f"　{_ql}：{_yr} {'、'.join(str(m) for m in _ms)} 月")
        st.info(f"今年度：{_tp.start.date()} ～ {_tp.end.date()}")
        for _ql, _yr, _ms in _fy_quarter_months(_tp):
            st.caption(f"　{_ql}：{_yr} {'、'.join(str(m) for m in _ms)} 月")
    st.markdown("**訂單狀態排除規則**")
    st.caption("系統固定排除：下單異常、已取消、已退貨")

    mode = st.radio("比較模式", ["Month", "Quarter"], index=1)

    if mode == "Quarter":
        st.caption("Quarter 可選擇單季或多季")
        selected_period_values = st.multiselect(
            "quarter_sel_label",
            options=QUARTER_ORDER,
            default=QUARTER_ORDER,
            key="quarter_sel",
            label_visibility="collapsed",
        )
    else:
        st.caption("Month 可選擇單月或多月")
        _last_range_ss = st.session_state.get("last_range", None)
        if _last_range_ss and isinstance(_last_range_ss, (tuple, list)) and len(_last_range_ss) == 2:
            _month_period_start = pd.Timestamp(_last_range_ss[0])
        else:
            _month_period_start = default_last_period.start
        _month_period = AnalysisPeriod("去年度", _month_period_start, _month_period_start + pd.DateOffset(years=1) - pd.Timedelta(days=1))
        _month_options = list(enumerate(month_labels_for_period(_month_period), start=1))
        _month_display = {idx: label for idx, label in _month_options}
        _default_month_orders = list(_month_display.keys())
        selected_period_values = st.multiselect(
            "month_sel_label",
            options=_default_month_orders,
            default=_default_month_orders,
            format_func=lambda x: _month_display.get(x, str(x)),
            key="month_sel",
            label_visibility="collapsed",
        )

    st.markdown("---")
    st.markdown("**分析年度區間**")
    # 若 session_state 已由年度選擇器寫入，優先使用；否則沿用 infer_default_periods 的預設值
    _last_range_default = st.session_state.get("last_range") or (default_last_period.start.date(), default_last_period.end.date())
    _this_range_default = st.session_state.get("this_range") or (default_this_period.start.date(), default_this_period.end.date())
    last_range_raw = st.date_input("去年度區間", value=_last_range_default, key="last_range")
    this_range_raw = st.date_input("今年度區間", value=_this_range_default, key="this_range")

# 資料來源：上傳優先，否則固定路徑
try:
    if uploaded_file is not None:
        cleaned_df, info, raw_df = clean_data_from_bytes(uploaded_file.getvalue())
        data_source_label = f"目前資料來源：使用上傳檔案【{uploaded_file.name}】"
    elif default_local_file_exists:
        cleaned_df, info, raw_df = clean_data_from_path(DEFAULT_LOCAL_XLSX_PATH)
        data_source_label = f"目前資料來源：自動載入固定路徑【{DEFAULT_LOCAL_XLSX_PATH}】"
    else:
        st.info(f"請先上傳 Excel 檔案，或確認固定路徑檔案存在：{DEFAULT_LOCAL_XLSX_PATH}")
        st.stop()
except Exception as exc:
    st.error(f"讀取 Excel 失敗：{exc}")
    st.stop()

last_range = date_input_to_tuple(last_range_raw)
this_range = date_input_to_tuple(this_range_raw)
last_period = AnalysisPeriod("去年度", pd.Timestamp(last_range[0]), pd.Timestamp(last_range[1])) if last_range else default_last_period
this_period = AnalysisPeriod("今年度", pd.Timestamp(this_range[0]), pd.Timestamp(this_range[1])) if this_range else default_this_period

with st.sidebar:
    reseller_options = sorted(cleaned_df["經銷商"].dropna().astype(str).unique().tolist()) if "經銷商" in cleaned_df.columns else []
    customer_options = sorted(cleaned_df["最終客戶"].dropna().astype(str).unique().tolist()) if "最終客戶" in cleaned_df.columns else []
    staff_options    = sorted(cleaned_df["展碁業務"].dropna().astype(str).unique().tolist()) if "展碁業務" in cleaned_df.columns else []
    selected_resellers = st.multiselect("經銷商", reseller_options)
    selected_customers = st.multiselect("最終客戶", customer_options)
    selected_staff     = st.multiselect("展碁業務", staff_options)
    promo_365_copilot  = st.checkbox(
        "365 Copilot Business 推廣",
        value=False,
        help="勾選後僅顯示：去年度曾購買 M365 Business 系列，且訂閱到期日在 90 天內的客戶",
    )

    st.markdown("---")
    st.markdown("**頁面資料篩選區間**")
    expiry_default_min = cleaned_df["訂閱到期日"].min().date() if "訂閱到期日" in cleaned_df.columns and cleaned_df["訂閱到期日"].notna().any() else default_last_period.start.date()
    expiry_default_max = cleaned_df["訂閱到期日"].max().date() if "訂閱到期日" in cleaned_df.columns and cleaned_df["訂閱到期日"].notna().any() else default_this_period.end.date()
    order_default_min = cleaned_df["訂單下單日"].min().date() if "訂單下單日" in cleaned_df.columns and cleaned_df["訂單下單日"].notna().any() else default_last_period.start.date()
    order_default_max = cleaned_df["訂單下單日"].max().date() if "訂單下單日" in cleaned_df.columns and cleaned_df["訂單下單日"].notna().any() else date.today()

    # ── 到期日快速篩選按鈕（Microsoft dashboard image003 設計） ────────────
    from datetime import timedelta as _td
    def _on_quick_expiry():
        _today = date.today()
        _sel = st.session_state.get("_quick_expiry", "全部")
        if _sel == "已到期":
            st.session_state["expiry_range"] = (expiry_default_min, _today - _td(days=1))
        elif _sel == "30天內":
            st.session_state["expiry_range"] = (_today, _today + _td(days=30))
        elif _sel == "60天內":
            st.session_state["expiry_range"] = (_today, _today + _td(days=60))
        elif _sel == "90天內":
            st.session_state["expiry_range"] = (_today, _today + _td(days=90))
        elif _sel == "180天內":
            st.session_state["expiry_range"] = (_today, _today + _td(days=180))
        # "全部" → do not modify, keep existing range

    st.caption("**到期日快速篩選**")
    st.radio(
        "_quick_expiry_label",
        ["全部", "已到期", "30天內", "60天內", "90天內", "180天內"],
        horizontal=True,
        key="_quick_expiry",
        on_change=_on_quick_expiry,
        label_visibility="collapsed",
    )

    expiry_range_raw = st.date_input("訂閱到期日區間（可手動調整）", value=(expiry_default_min, expiry_default_max), key="expiry_range")
    order_range_raw = st.date_input("訂單下單日區間", value=(order_default_min, order_default_max), key="order_range")

expiry_range = date_input_to_tuple(expiry_range_raw)
order_range = date_input_to_tuple(order_range_raw)

with _data_exp:
    st.caption(data_source_label)
    if info["warnings"]:
        for msg in info["warnings"]:
            st.warning(msg)
    if info["missing_columns"]:
        st.error("缺少必要分析欄位：" + "、".join(info["missing_columns"]))
        st.stop()
    with st.expander("原始資料摘要", expanded=False):
        c1, c2 = st.columns(2)
        c1.metric("原始資料筆數", fmt_int(len(raw_df)))
        c2.metric("原始資料欄位數", fmt_int(raw_df.shape[1]))
        st.dataframe(raw_df.head(200), use_container_width=True, height=300)
    with st.expander("清洗資訊與清洗後資料摘要", expanded=False):
        c1, c2 = st.columns(2)
        c1.metric("清洗後資料筆數", fmt_int(len(cleaned_df)))
        c2.metric("清洗後資料欄位數", fmt_int(cleaned_df.shape[1]))
        if info["dropped_columns"]:
            st.caption("已刪除欄位：" + "、".join(info["dropped_columns"]))
        failures = [f"{k}={v}" for k, v in info["date_parse_failures"].items() if v > 0]
        st.caption("日期轉換失敗筆數：" + ("；".join(failures) if failures else "0"))
        st.dataframe(cleaned_df.head(200), use_container_width=True, height=320)

filtered_df = apply_filters(cleaned_df, selected_resellers, selected_customers, expiry_range, order_range, selected_staff)
filtered_df, forecast_key_df = calculate_forecast(filtered_df, last_period, this_period)

# Month / Quarter 子選擇 mask（連動左側 Month/Quarter 篩選至所有表格與圖表）
if mode == "Month" and selected_period_values:
    _period_sel_mask = filtered_df["_fiscal_month_order"].isin(selected_period_values)
elif mode == "Quarter" and selected_period_values:
    _period_sel_mask = filtered_df["_quarter_short"].isin(selected_period_values)
else:
    _period_sel_mask = pd.Series(True, index=filtered_df.index)
filtered_df_sel = filtered_df[_period_sel_mask].copy()

# ── 365 Copilot Business 推廣 篩選 ───────────────────────────────────────
if promo_365_copilot and not filtered_df_sel.empty:
    _promo_today = pd.Timestamp.today().normalize()

    if "訂閱到期日" in filtered_df_sel.columns and \
       "最終客戶"   in filtered_df_sel.columns and \
       "商品名稱"   in filtered_df_sel.columns:

        _prod_col = filtered_df_sel["商品名稱"].astype(str)

        # 條件 1：訂閱到期日在今天起 90 天內（0 ~ 90 天）
        _exp_days = (
            pd.to_datetime(filtered_df_sel["訂閱到期日"], errors="coerce")
            .sub(_promo_today).dt.days
        )
        _exp_90_custs = set(
            filtered_df_sel.loc[
                (_exp_days >= 0) & (_exp_days <= 90), "最終客戶"
            ].dropna().astype(str)
        )

        # 條件 2：去年度曾購買三種 M365 Business 產品之一（包含比對）
        _last_mask_tmp = build_period_mask(filtered_df_sel, last_period)
        _m365_mask = (
            _prod_col.str.contains("Microsoft 365 Business Basic",    case=False, na=False) |
            _prod_col.str.contains("Microsoft 365 Business Standard", case=False, na=False) |
            _prod_col.str.contains("Microsoft 365 Business Premium",  case=False, na=False)
        )
        _m365_custs = set(
            filtered_df_sel.loc[_last_mask_tmp & _m365_mask, "最終客戶"]
            .dropna().astype(str)
        )

        # 條件 3：從未購買 Copilot 相關產品（全資料範圍排除）
        _cop_mask = (
            _prod_col.str.contains("Copilot", case=False, na=False) &
            ~_prod_col.str.contains("Microsoft Copilot Studio", case=False, na=False)
        )
        _cop_custs = set(
            filtered_df_sel.loc[_cop_mask, "最終客戶"].dropna().astype(str)
        )

        # 三條件交集：90 天內到期 ∩ 去年買 M365 ∩ 從未買 Copilot
        _promo_custs    = (_exp_90_custs & _m365_custs) - _cop_custs
        filtered_df_sel = filtered_df_sel[
            filtered_df_sel["最終客戶"].astype(str).isin(_promo_custs)
        ].copy()
        filtered_df = filtered_df[
            filtered_df["最終客戶"].astype(str).isin(_promo_custs)
        ].copy()

if filtered_df.empty:
    st.warning("目前篩選條件下沒有資料。請調整客戶或日期區間。")
    st.stop()

# ── 共用遮罩：避免重複呼叫 build_period_mask ────────────────────────────
_sel_last_mask = build_period_mask(filtered_df_sel, last_period)
_sel_this_mask = build_period_mask(filtered_df_sel, this_period)
_sel_last_df   = filtered_df_sel[_sel_last_mask]
_sel_this_df   = filtered_df_sel[_sel_this_mask]


# ── Copilot New Penetration 資料（Tab 2 / Grouped Detail 共用）──────────
def _copilot_product_mask(df: pd.DataFrame) -> pd.Series:
    col = df["商品名稱"].astype(str)
    return (
        col.str.contains("Copilot", case=False, na=False) &
        ~col.str.contains("Microsoft Copilot Studio", case=False, na=False)
    )

if "商品名稱" in filtered_df_sel.columns and "最終客戶" in filtered_df_sel.columns:
    _last_cop_cust = set(
        _sel_last_df[_copilot_product_mask(_sel_last_df)]["最終客戶"].dropna().astype(str)
    )
    _this_cop_df  = _sel_this_df[_copilot_product_mask(_sel_this_df)].copy()
    _new_pen_df   = _this_cop_df[~_this_cop_df["最終客戶"].astype(str).isin(_last_cop_cust)].copy()
else:
    _new_pen_df = pd.DataFrame()


# ── 向量化 Grouped Detail 表格建構（所有分頁共用）──────────────────────
def _build_grouped_detail(src_df: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    """
    以全向量化方式建立 Master-Detail 大綱表格，不使用 iterrows()。
    回傳 (outline_df, n_groups, n_rows)
    """
    if src_df.empty:
        return pd.DataFrame(), 0, 0

    df = src_df.copy()

    # 確保 FY年度 存在
    if "FY年度" not in df.columns:
        df["FY年度"] = pd.to_datetime(
            df["訂閱到期日"], errors="coerce"
        ).map(lambda d: get_fy_label_from_date(d) if pd.notna(d) else "")

    # 1. 群組摘要
    summary = (
        df.groupby(["最終客戶", "FY年度"], dropna=False)["成交價未稅小計"]
        .sum().reset_index()
        .rename(columns={"成交價未稅小計": "_總金額"})
        .sort_values("_總金額", ascending=False)
        .reset_index(drop=True)
    )
    summary["_rank"] = range(1, len(summary) + 1)

    # 2. 標題列
    header = pd.DataFrame({
        "_rank":         summary["_rank"],
        "_row_order":    0,
        "排序":           summary["_rank"],
        "最終客戶":       summary["最終客戶"].astype(str),
        "訂閱到期日年度": summary["FY年度"].astype(str),
        "加總金額":       summary["_總金額"].map(lambda x: f"{x:,.0f}"),
        "訂單動作":       None,
        "訂閱到期日":     None,
        "商品名稱":       None,
        "數量":           None,
        "成交單價未稅":   None,
        "成交價未稅小計": None,
        "經銷商":         None,
        "展碁業務":       None,
    })

    # 3. 明細列（向量化格式化）
    detail = df.merge(
        summary[["最終客戶", "FY年度", "_rank"]],
        on=["最終客戶", "FY年度"], how="left"
    ).sort_values(["_rank", "訂閱到期日"])

    detail["_dt_str"] = pd.to_datetime(
        detail["訂閱到期日"], errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    def _vec_fmt_money(s: pd.Series) -> pd.Series:
        return pd.to_numeric(s, errors="coerce").map(
            lambda v: f"{v:,.0f}" if pd.notna(v) else None
        )
    def _vec_fmt_qty(s: pd.Series) -> pd.Series:
        return pd.to_numeric(s, errors="coerce").map(
            lambda v: f"{int(round(v)):,}" if pd.notna(v) else None
        )

    detail_rows = pd.DataFrame({
        "_rank":         detail["_rank"],
        "_row_order":    detail.groupby("_rank").cumcount() + 1,
        "排序":           None,
        "最終客戶":       None,
        "訂閱到期日年度": None,
        "加總金額":       None,
        "訂單動作":       detail["訂閱動作"]     if "訂閱動作"     in detail.columns else None,
        "訂閱到期日":     detail["_dt_str"],
        "商品名稱":       detail["商品名稱"]     if "商品名稱"     in detail.columns else None,
        "數量":           _vec_fmt_qty(detail["數量"])          if "數量"          in detail.columns else None,
        "成交單價未稅":   _vec_fmt_money(detail["成交單價未稅"]) if "成交單價未稅"  in detail.columns else None,
        "成交價未稅小計": _vec_fmt_money(detail["成交價未稅小計"]) if "成交價未稅小計" in detail.columns else None,
        "經銷商":         detail["經銷商"]       if "經銷商"       in detail.columns else None,
        "展碁業務":       detail["展碁業務"]     if "展碁業務"     in detail.columns else None,
    })

    # 4. 合併排序，去除工作欄
    _outline_cols = [
        "排序", "最終客戶", "訂閱到期日年度", "加總金額",
        "訂單動作", "訂閱到期日", "商品名稱",
        "數量", "成交單價未稅", "成交價未稅小計",
        "經銷商", "展碁業務",
    ]
    combined = (
        pd.concat([header, detail_rows], ignore_index=True)
        .sort_values(["_rank", "_row_order"])
        .drop(columns=["_rank", "_row_order"])
        .reset_index(drop=True)
    )
    combined = combined[[c for c in _outline_cols if c in combined.columns]]
    return combined, len(summary), len(df)


tab_main, tab_copilot, tab_marketing, tab_csp = st.tabs(
    ["📊 主要分析", "🤖 Copilot New Penetration", "🎯 精準行銷", "📈 CSP 指標（微軟標準）"]
)


with tab_main:
    st.markdown("---")
    st.subheader("KPI 指標（以「成交價未稅小計」分析）")
    last_metrics = build_kpi_summary(filtered_df_sel, last_period)
    this_metrics = build_kpi_summary(filtered_df_sel, this_period)
    render_main_kpi_cards(
        last_metrics, this_metrics,
        this_label=this_period.label,
        last_label=last_period.label,
    )

    st.markdown("---")
    st.subheader("比較分析")
    summary_df = build_period_summary(filtered_df, forecast_key_df, mode, last_period, this_period, selected_period_values)
    with st.expander("展開數據表", expanded=False):
        if summary_df.empty:
            st.info("目前分析期間內沒有可彙總資料。")
        else:
            display_summary = summary_df.copy()
            for col in ["去年度金額", "今年度金額", "Forecast", "差異"]:
                if col in display_summary.columns:
                    display_summary[col] = pd.to_numeric(display_summary[col], errors="coerce").fillna(0).map(lambda x: f"{x:,.0f}")
            st.dataframe(display_summary, use_container_width=True, hide_index=True)

    st.markdown("---")
    chart_col, side_col = st.columns([3.3, 1.1])
    with chart_col:
        fig = build_trend_figure(summary_df, mode, last_period, this_period)
        st.plotly_chart(fig, use_container_width=True)
    with side_col:
        with st.expander("Forecast 說明", expanded=False):
            st.markdown(
                """
                <div class="card">
                    <div class="subtle">
                    比對 key：最終客戶 + 商品名稱<br><br>
                    金額欄位：成交價未稅小計<br><br>
                    規則：<br>
                    1. 去年度有、今年度沒有 ⇒ Forecast = 去年度金額 × 0.8<br>
                    2. 今年度部分續約 ⇒ Forecast = max(去年度金額 - 今年度金額, 0) × 0.8<br><br>
                    Forecast 以 key level 計算後，再依去年度明細金額占比分攤回列資料。
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if not forecast_key_df.empty:
                preview = forecast_key_df[["最終客戶", "商品名稱", "去年度金額", "今年度金額", "Forecast"]].copy().sort_values("Forecast", ascending=False).head(20)
                for col in ["去年度金額", "今年度金額", "Forecast"]:
                    preview[col] = pd.to_numeric(preview[col], errors="coerce").fillna(0).map(lambda x: f"{x:,.0f}")
                st.dataframe(preview, use_container_width=True, height=320, hide_index=True)

with tab_main:
    st.markdown("---")
    st.subheader("Top 15 客戶（以去年度成交價未稅小計合計排序）")

    # FY 標籤（供欄位標題與 FY年度 欄使用）
    _last_fy = get_fy_label_from_date(last_period.start)
    _this_fy = get_fy_label_from_date(this_period.start)
    _fy_label = f"{_last_fy} / {_this_fy}"

    _last_mask = _sel_last_mask
    _this_mask = _sel_this_mask

    # 以「最終客戶」為單位加總
    _last_grp = (
        filtered_df_sel[_last_mask]
        .groupby("最終客戶", dropna=False)["成交價未稅小計"]
        .sum(min_count=1).reset_index()
        .rename(columns={"成交價未稅小計": "去年度總金額"})
    )
    _this_grp = (
        filtered_df_sel[_this_mask]
        .groupby("最終客戶", dropna=False)["成交價未稅小計"]
        .sum(min_count=1).reset_index()
        .rename(columns={"成交價未稅小計": "今年度總金額"})
    )
    _top15_df = (
        _last_grp.merge(_this_grp, on="最終客戶", how="outer")
        .fillna({"去年度總金額": 0, "今年度總金額": 0})
    )
    _top15_df["去年度總金額"] = pd.to_numeric(_top15_df["去年度總金額"], errors="coerce").fillna(0)
    _top15_df["今年度總金額"] = pd.to_numeric(_top15_df["今年度總金額"], errors="coerce").fillna(0)
    _top15_df["差異金額"]     = _top15_df["今年度總金額"] - _top15_df["去年度總金額"]
    _top15_df["FY年度"]       = _fy_label
    _top15_df = (
        _top15_df.sort_values("去年度總金額", ascending=False)
        .head(15)
        .reset_index(drop=True)
    )
    _top15_df.insert(0, "排序", range(1, len(_top15_df) + 1))
    _top15_display = _top15_df[["排序", "最終客戶", "去年度總金額", "今年度總金額", "差異金額", "FY年度"]].copy()
    for _c in ["去年度總金額", "今年度總金額", "差異金額"]:
        _top15_display[_c] = _top15_display[_c].map(lambda x: f"{x:,.0f}")
    st.dataframe(_top15_display, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Top 15 經銷商（以去年度成交價未稅小計合計排序）")

    _last_reseller_grp = (
        filtered_df_sel[_last_mask]
        .groupby("經銷商", dropna=False)["成交價未稅小計"]
        .sum(min_count=1).reset_index()
        .rename(columns={"成交價未稅小計": "去年度總金額"})
    )
    _this_reseller_grp = (
        filtered_df_sel[_this_mask]
        .groupby("經銷商", dropna=False)["成交價未稅小計"]
        .sum(min_count=1).reset_index()
        .rename(columns={"成交價未稅小計": "今年度總金額"})
    )
    _top15_reseller_df = (
        _last_reseller_grp.merge(_this_reseller_grp, on="經銷商", how="outer")
        .fillna({"去年度總金額": 0, "今年度總金額": 0})
    )
    _top15_reseller_df["去年度總金額"] = pd.to_numeric(_top15_reseller_df["去年度總金額"], errors="coerce").fillna(0)
    _top15_reseller_df["今年度總金額"] = pd.to_numeric(_top15_reseller_df["今年度總金額"], errors="coerce").fillna(0)
    _top15_reseller_df["差異金額"]     = _top15_reseller_df["今年度總金額"] - _top15_reseller_df["去年度總金額"]
    _top15_reseller_df["FY年度"]       = _fy_label
    _top15_reseller_df = (
        _top15_reseller_df.sort_values("去年度總金額", ascending=False)
        .head(15)
        .reset_index(drop=True)
    )
    _top15_reseller_df.insert(0, "排序", range(1, len(_top15_reseller_df) + 1))
    _top15_reseller_display = _top15_reseller_df[["排序", "經銷商", "去年度總金額", "今年度總金額", "差異金額", "FY年度"]].copy()
    for _c in ["去年度總金額", "今年度總金額", "差異金額"]:
        _top15_reseller_display[_c] = _top15_reseller_display[_c].map(lambda x: f"{x:,.0f}")
    st.dataframe(_top15_reseller_display, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Top 15 商品名稱")

    def _build_top15_product_fig(df: pd.DataFrame, period: AnalysisPeriod, bar_color: str, title: str) -> go.Figure:
        mask = build_period_mask(df, period)
        grp = (
            df[mask]
            .groupby("商品名稱", dropna=False)["成交價未稅小計"]
            .sum(min_count=1).reset_index()
            .rename(columns={"成交價未稅小計": "金額"})
        )
        grp["金額"] = pd.to_numeric(grp["金額"], errors="coerce").fillna(0)
        top15 = grp.sort_values("金額", ascending=False).head(15).reset_index(drop=True)
        # 反轉順序讓第1名在最上方
        top15 = top15.iloc[::-1].reset_index(drop=True)

        _fmt_bar = lambda v: f"{v/1e6:.1f}M" if abs(v) >= 1e6 else f"{v:,.0f}"
        text_vals = [_fmt_bar(v) for v in top15["金額"]]

        fig = go.Figure(go.Bar(
            x=top15["金額"],
            y=top15["商品名稱"],
            orientation="h",
            marker_color=bar_color,
            text=text_vals,
            textposition="outside",
            textfont=dict(size=12),
            cliponaxis=False,
            hovertemplate="%{y}<br>%{x:,.0f}<extra></extra>",
        ))
        fig.update_layout(
            title=dict(text=title, font=dict(size=14), x=0.5, xanchor="center"),
            height=520,
            margin=dict(l=10, r=80, t=50, b=30),
            plot_bgcolor="white",
            paper_bgcolor="white",
            xaxis=dict(showgrid=True, gridcolor="#E5E7EB", tickformat=","),
            yaxis=dict(showgrid=False, automargin=True, tickfont=dict(family="Arial Black, Arial Bold, sans-serif", size=11)),
        )
        return fig

    _last_fy_lbl = get_fy_label_from_date(last_period.start)
    _this_fy_lbl = get_fy_label_from_date(this_period.start)

    _prod_col1, _prod_col2 = st.columns(2)
    with _prod_col1:
        _fig_last_prod = _build_top15_product_fig(
            filtered_df_sel, last_period, "#F97316",
            f"去年度 Top 15 商品名稱（{_last_fy_lbl}）"
        )
        st.plotly_chart(_fig_last_prod, use_container_width=True)
    with _prod_col2:
        _fig_this_prod = _build_top15_product_fig(
            filtered_df_sel, this_period, "#1D4ED8",
            f"今年度 Top 15 商品名稱（{_this_fy_lbl}）"
        )
        st.plotly_chart(_fig_this_prod, use_container_width=True)

    st.markdown("---")
    st.subheader("Top 15 商品名稱（增減比較）")

    def _build_top15_change_fig(
        df: pd.DataFrame,
        last_period: AnalysisPeriod,
        this_period: AnalysisPeriod,
        value_col: str,
        title: str,
        fmt_func,
    ) -> go.Figure:
        _lm = build_period_mask(df, last_period)
        _tm = build_period_mask(df, this_period)
        _lg = (
            df[_lm].groupby("商品名稱", dropna=False)[value_col]
            .sum(min_count=1).reset_index().rename(columns={value_col: "去年度"})
        )
        _tg = (
            df[_tm].groupby("商品名稱", dropna=False)[value_col]
            .sum(min_count=1).reset_index().rename(columns={value_col: "今年度"})
        )
        _mg = _lg.merge(_tg, on="商品名稱", how="outer").fillna(0)
        _mg["去年度"] = pd.to_numeric(_mg["去年度"], errors="coerce").fillna(0)
        _mg["今年度"] = pd.to_numeric(_mg["今年度"], errors="coerce").fillna(0)
        _mg["差異"]   = _mg["今年度"] - _mg["去年度"]

        # 依去年度數值降序取 Top 15，第 1 名排最上方
        _top = (
            _mg.sort_values("去年度", ascending=False)
            .head(15)
            .iloc[::-1]
            .reset_index(drop=True)
        )

        _last_fy = get_fy_label_from_date(last_period.start)
        _this_fy = get_fy_label_from_date(this_period.start)

        fig = go.Figure()
        fig.add_trace(go.Bar(
            name=f"去年度（{_last_fy}）",
            y=_top["商品名稱"],
            x=_top["去年度"],
            orientation="h",
            marker_color="#F97316",
            text=[fmt_func(v) for v in _top["去年度"]],
            textposition="outside",
            textfont=dict(size=11),
            cliponaxis=False,
            hovertemplate="%{y}<br>去年度：%{x:,.0f}<extra></extra>",
        ))
        fig.add_trace(go.Bar(
            name=f"今年度（{_this_fy}）",
            y=_top["商品名稱"],
            x=_top["今年度"],
            orientation="h",
            marker_color="#1D4ED8",
            text=[fmt_func(v) for v in _top["今年度"]],
            textposition="outside",
            textfont=dict(size=11),
            cliponaxis=False,
            hovertemplate="%{y}<br>今年度：%{x:,.0f}<extra></extra>",
        ))
        fig.update_layout(
            title=dict(text=title, font=dict(size=14), x=0.5, xanchor="center"),
            barmode="group",
            height=600,
            margin=dict(l=10, r=100, t=55, b=60),
            plot_bgcolor="white",
            paper_bgcolor="white",
            legend=dict(orientation="h", yanchor="top", y=-0.06, xanchor="center", x=0.5),
            xaxis=dict(showgrid=True, gridcolor="#E5E7EB", tickformat=","),
            yaxis=dict(showgrid=False, automargin=True, tickfont=dict(family="Arial Black, Arial Bold, sans-serif", size=11)),
        )
        return fig

    _fmt_money  = lambda v: f"{v/1e6:.1f}M" if abs(v) >= 1e6 else f"{v:,.0f}"
    _fmt_qty    = lambda v: f"{int(round(v)):,}"

    _chg_col1, _chg_col2 = st.columns(2)
    with _chg_col1:
        st.plotly_chart(
            _build_top15_change_fig(
                filtered_df_sel, last_period, this_period,
                "成交價未稅小計",
                "Top 15 商品名稱（金額之增減｜今年度 vs 去年度）",
                _fmt_money,
            ),
            use_container_width=True,
        )
    with _chg_col2:
        st.plotly_chart(
            _build_top15_change_fig(
                filtered_df_sel, last_period, this_period,
                "數量",
                "Top 15 商品名稱（數量之增減｜今年度 vs 去年度）",
                _fmt_qty,
            ),
            use_container_width=True,
        )

# ── Copilot New Penetration ────────────────────────────────────────────
with tab_copilot:
    st.subheader("Copilot New Penetration")
    st.caption("定義：去年度（訂閱到期日）無 Copilot 相關購買、今年度有購買的客戶（排除 Microsoft Copilot Studio）")

    if "商品名稱" in filtered_df_sel.columns and "最終客戶" in filtered_df_sel.columns:
        # 使用預先計算的 _new_pen_df（在 tabs 宣告前已建立）
        if _new_pen_df.empty:
            st.info("目前篩選條件下無 Copilot New Penetration 資料。")
        else:
            st.markdown(f"**符合條件客戶數：{_new_pen_df['最終客戶'].nunique():,} 位　｜　新增筆數：{len(_new_pen_df):,} 筆**")

            # ── 1. 商品名稱彙總 ──────────────────────────────────────
            st.markdown("#### 商品名稱彙總")
            _pen_prod = (
                _new_pen_df.groupby("商品名稱", dropna=False)
                .agg(加總金額=("成交價未稅小計", "sum"), 加總數量=("數量", "sum"))
                .reset_index()
                .sort_values("加總金額", ascending=False)
                .reset_index(drop=True)
            )
            _pen_prod.insert(0, "排序", range(1, len(_pen_prod) + 1))
            _pen_prod_display = _pen_prod.copy()
            _pen_prod_display["加總金額"] = _pen_prod_display["加總金額"].map(lambda x: f"{x:,.0f}")
            _pen_prod_display["加總數量"] = _pen_prod_display["加總數量"].map(lambda x: f"{int(round(x)):,}")
            st.dataframe(_pen_prod_display, use_container_width=True, hide_index=True)

            # ── 2. Top 15 經銷商 & 客戶 並排 ─────────────────────────
            st.markdown("#### Top 15 經銷商 ／ 客戶")
            _pen_r_col, _pen_c_col = st.columns(2)

            def _pen_top15_fig(grp_col: str, bar_color: str, title: str) -> go.Figure:
                _agg_cols = {"金額": ("成交價未稅小計", "sum")}
                if "數量" in _new_pen_df.columns:
                    _agg_cols["數量"] = ("數量", "sum")
                _grp = (
                    _new_pen_df.groupby(grp_col, dropna=False)
                    .agg(**_agg_cols)
                    .reset_index()
                    .sort_values("金額", ascending=False)
                    .head(15)
                    .iloc[::-1]
                    .reset_index(drop=True)
                )
                _fmt_amt = lambda v: f"{v/1e6:.1f}M" if abs(v) >= 1e6 else f"{v:,.0f}"
                _fmt_qty = lambda v: f"{int(round(v)):,}"
                if "數量" in _grp.columns:
                    _labels = [
                        f"{_fmt_amt(a)}  ({_fmt_qty(q)})"
                        for a, q in zip(_grp["金額"], _grp["數量"])
                    ]
                else:
                    _labels = [_fmt_amt(v) for v in _grp["金額"]]
                fig = go.Figure(go.Bar(
                    x=_grp["金額"], y=_grp[grp_col],
                    orientation="h",
                    marker_color=bar_color,
                    text=_labels,
                    textposition="outside",
                    textfont=dict(size=11),
                    cliponaxis=False,
                    hovertemplate="%{y}<br>金額：%{x:,.0f}<extra></extra>",
                ))
                fig.update_layout(
                    title=dict(text=title, font=dict(size=13), x=0.5, xanchor="center"),
                    height=500,
                    margin=dict(l=10, r=90, t=45, b=30),
                    plot_bgcolor="white", paper_bgcolor="white",
                    xaxis=dict(showgrid=True, gridcolor="#E5E7EB", tickformat=","),
                    yaxis=dict(showgrid=False, automargin=True,
                               tickfont=dict(family="Arial Black, Arial Bold, sans-serif", size=11)),
                )
                return fig

            with _pen_r_col:
                st.plotly_chart(_pen_top15_fig("經銷商", "#F97316", "Top 15 經銷商（Copilot New Penetration）"), use_container_width=True)
            with _pen_c_col:
                st.plotly_chart(_pen_top15_fig("最終客戶", "#1D4ED8", "Top 15 客戶（Copilot New Penetration）"), use_container_width=True)
    else:
        st.warning("缺少【商品名稱】或【最終客戶】欄位，無法計算 Copilot New Penetration。")

    # ── Copilot New Penetration Grouped Detail ────────────────────────
    st.markdown("---")
    st.subheader("Copilot New Penetration Grouped Detail")

    if "商品名稱" in filtered_df_sel.columns and "最終客戶" in filtered_df_sel.columns:
        if _new_pen_df.empty:
            st.info("目前篩選條件下無 Copilot New Penetration 資料。")
        else:
            _gd_result, _gd_n_grp, _gd_n_rows = _build_grouped_detail(_new_pen_df)
            st.caption(f"共 {_gd_n_grp} 個客戶群組 ／ {_gd_n_rows:,} 筆明細　（依加總金額降序排列）")
            st.dataframe(_gd_result, use_container_width=True, hide_index=True)
    else:
        st.info("缺少必要欄位，無法顯示 Grouped Detail。")

with tab_main:
    detail_columns = [
        "訂閱到期日", "訂單下單日", "經銷商", "最終客戶", "商品名稱", "數量",
        "展碁COST單價未稅", "展碁COST未稅小計", "成交價未稅小計", "訂閱動作", "FY年度", "季度代碼", "Forecast"
    ]

    def _make_detail_df(base_df: pd.DataFrame, period: AnalysisPeriod) -> pd.DataFrame:
        mask = build_period_mask(base_df, period)
        cols = [c for c in detail_columns if c in base_df.columns]
        df = base_df.loc[mask, cols].copy()
        for dt_col in ["訂閱到期日", "訂單下單日"]:
            if dt_col in df.columns:
                df[dt_col] = pd.to_datetime(df[dt_col], errors="coerce").dt.strftime("%Y-%m-%d")
        return df

    _last_fy_label = get_fy_label_from_date(last_period.start)
    _this_fy_label = get_fy_label_from_date(this_period.start)

    last_detail_df = _make_detail_df(filtered_df_sel, last_period)
    this_detail_df = _make_detail_df(filtered_df_sel, this_period)

    st.markdown("---")
    st.subheader(f"去年度明細表（{_last_fy_label}）")
    with st.expander(f"展開去年度明細表（{_last_fy_label}）", expanded=False):
        st.caption(f"共 {len(last_detail_df):,} 筆")
        st.dataframe(last_detail_df, use_container_width=True, height=520, hide_index=True)

    st.markdown("---")
    st.subheader(f"今年度明細表（{_this_fy_label}）")
    with st.expander(f"展開今年度明細表（{_this_fy_label}）", expanded=False):
        st.caption(f"共 {len(this_detail_df):,} 筆")
        st.dataframe(this_detail_df, use_container_width=True, height=520, hide_index=True)

    st.markdown("---")
    st.subheader("下載")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.download_button(f"下載去年度明細（CSV）", data=last_detail_df.to_csv(index=False).encode("utf-8-sig"), file_name=f"detail_{_last_fy_label}.csv", mime="text/csv", use_container_width=True)
    with col2:
        st.download_button(f"下載去年度明細（Excel）", data=to_excel_bytes(last_detail_df), file_name=f"detail_{_last_fy_label}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with col3:
        st.download_button(f"下載今年度明細（CSV）", data=this_detail_df.to_csv(index=False).encode("utf-8-sig"), file_name=f"detail_{_this_fy_label}.csv", mime="text/csv", use_container_width=True)
    with col4:
        st.download_button(f"下載今年度明細（Excel）", data=to_excel_bytes(this_detail_df), file_name=f"detail_{_this_fy_label}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.markdown("---")
    st.markdown("### 設計說明")
    st.markdown(
        """
1. **Forecast 趨勢圖調整**
   上半部改為：去年度柱狀體獨立顯示；今年度柱狀體與 Forecast 柱狀體採同一組堆疊，讓 Forecast 顯示在今年度柱狀體上方。
   下半部新增去年度與今年度差異柱狀圖，並將正負數值顯示於柱體外側，避免互相遮擋。

2. **Month / Quarter 選擇位置調整**
   已將 `Quarter 可選擇單季或多季` 與 `Month 可選擇單月或多月` 移到比較模式邏輯區下方，並依目前模式動態顯示。

3. **Forecast 穩定版設計**
   採用 `最終客戶 + 商品名稱` 的 key-based 計算流程；先做 key level Forecast，再依去年度明細金額占比分攤回列資料，避免圖表與明細表不一致。
        """
    )

# ── 精準行銷 分頁 ──────────────────────────────────────────────────────
with tab_marketing:
    st.markdown("---")

    _mkt_last_mask = _sel_last_mask
    _mkt_this_mask = _sel_this_mask

    # 向量化取「每個客戶最常出現的值」（無需 Python-level 自訂函式）
    def _vec_mode_by_group(df: pd.DataFrame, grp_col: str, val_col: str) -> pd.Series:
        """回傳 Series，index=grp_col，value=val_col 的眾數（最常出現值）"""
        if val_col not in df.columns:
            return pd.Series(dtype=str)
        return (
            df[[grp_col, val_col]].dropna(subset=[val_col])
            .groupby([grp_col, val_col]).size()
            .reset_index(name="_n")
            .sort_values("_n", ascending=False)
            .drop_duplicates(subset=[grp_col])
            .set_index(grp_col)[val_col]
        )

    # 去年度金額
    _mkt_last_grp = (
        filtered_df_sel[_mkt_last_mask]
        .groupby("最終客戶", dropna=False)["成交價未稅小計"]
        .sum(min_count=1).reset_index()
        .rename(columns={"成交價未稅小計": "去年度總金額"})
    )

    # 今年度金額
    _mkt_this_grp = (
        filtered_df_sel[_mkt_this_mask]
        .groupby("最終客戶", dropna=False)["成交價未稅小計"]
        .sum(min_count=1).reset_index()
        .rename(columns={"成交價未稅小計": "今年度總金額"})
    )

    # 經銷商 / 展碁業務：今年度優先，不足補去年度（全向量化）
    _ty = filtered_df_sel[_mkt_this_mask]
    _ly = filtered_df_sel[_mkt_last_mask]
    _reseller_ty  = _vec_mode_by_group(_ty, "最終客戶", "經銷商")
    _reseller_ly  = _vec_mode_by_group(_ly, "最終客戶", "經銷商")
    _staff_ty     = _vec_mode_by_group(_ty, "最終客戶", "展碁業務")
    _staff_ly     = _vec_mode_by_group(_ly, "最終客戶", "展碁業務")
    # 合併：今年度有值用今年度，否則用去年度
    _reseller_map = _reseller_ty.combine_first(_reseller_ly)
    _staff_map    = _staff_ty.combine_first(_staff_ly)

    # 合併去年度 / 今年度金額
    _mkt_tbl = _mkt_last_grp.merge(_mkt_this_grp, on="最終客戶", how="outer")
    _mkt_tbl["去年度總金額"] = pd.to_numeric(_mkt_tbl["去年度總金額"], errors="coerce").fillna(0)
    _mkt_tbl["今年度總金額"] = pd.to_numeric(_mkt_tbl["今年度總金額"], errors="coerce").fillna(0)
    _mkt_tbl["差異金額"]     = _mkt_tbl["今年度總金額"] - _mkt_tbl["去年度總金額"]

    # 對應填入經銷商 / 展碁業務
    if not _reseller_map.empty:
        _mkt_tbl["經銷商"] = _mkt_tbl["最終客戶"].map(_reseller_map).fillna("")
    if not _staff_map.empty:
        _mkt_tbl["展碁業務"] = _mkt_tbl["最終客戶"].map(_staff_map).fillna("")

    # 排序
    _mkt_tbl = _mkt_tbl.sort_values("去年度總金額", ascending=False).reset_index(drop=True)

    # ── 未續約示警：依去年度最晚到期日距今天數判斷 ──────────────────────
    _today = pd.Timestamp.today().normalize()

    if "訂閱到期日" in _sel_last_df.columns and not _sel_last_df.empty:
        _exp_tmp = _sel_last_df[["最終客戶", "訂閱到期日"]].copy()
        _exp_tmp["訂閱到期日"] = pd.to_datetime(_exp_tmp["訂閱到期日"], errors="coerce")
        _max_expiry_map = _exp_tmp.groupby("最終客戶")["訂閱到期日"].max()
        _mkt_tbl["_days"] = (
            _mkt_tbl["最終客戶"].map(_max_expiry_map)
            .sub(_today)
            .dt.days
        )
    else:
        _mkt_tbl["_days"] = float("nan")

    _renewed     = _mkt_tbl["今年度總金額"] > 0
    _d           = _mkt_tbl["_days"]
    _not_renewed = ~_renewed & _d.notna()
    _warn        = pd.Series("", index=_mkt_tbl.index)

    # 已續約：綠方塊（今年度 >= 去年度）/ 紅方塊（今年度 < 去年度）
    _warn.loc[_renewed & (_mkt_tbl["今年度總金額"] >= _mkt_tbl["去年度總金額"])] = "🟩 已續約"
    _warn.loc[_renewed & (_mkt_tbl["今年度總金額"] <  _mkt_tbl["去年度總金額"])] = "🟥 已續約"

    # 未續約：依距到期日天數
    _warn.loc[_not_renewed & (_d <   0              )] = "🔴 已到期"
    _warn.loc[_not_renewed & (_d >=  0) & (_d <= 15 )] = "🟡 15天"
    _warn.loc[_not_renewed & (_d >  15) & (_d <= 30 )] = "🟠 30天"
    _warn.loc[_not_renewed & (_d >  30) & (_d <= 45 )] = "🟣 45天"
    _warn.loc[_not_renewed & (_d >  45) & (_d <= 60 )] = "🟤 60天"
    _warn.loc[_not_renewed & (_d >  60) & (_d <= 90 )] = "⚫ 90天"

    # _days 保留給後續續約率計算使用（不會出現在 _mkt_display 的欄位清單中）

    # 插入固定欄
    _mkt_tbl.insert(0, "排序", range(1, len(_mkt_tbl) + 1))
    _mkt_tbl.insert(1, "未續約示警", _warn)

    _mkt_display_cols = ["排序", "未續約示警", "最終客戶", "去年度總金額", "今年度總金額", "差異金額"]
    for _c in ["經銷商", "展碁業務"]:
        if _c in _mkt_tbl.columns:
            _mkt_display_cols.append(_c)

    _mkt_display = _mkt_tbl[_mkt_display_cols].copy()
    for _c in ["去年度總金額", "今年度總金額", "差異金額"]:
        _mkt_display[_c] = _mkt_display[_c].map(lambda x: f"{x:,.0f}")

    # ── 未續約示警分析報告 ───────────────────────────────────────────────
    # 統計數值
    _W_COLORS = {
        "🔴 已到期": "#EF4444",
        "🟡 15天":   "#EAB308",
        "🟠 30天":   "#F97316",
        "🟣 45天":   "#A855F7",
        "🟤 60天":   "#92400E",
        "⚫ 90天":   "#374151",
        "🟩 已續約": "#22C55E",
        "🟥 已續約": "#F43F5E",
    }
    _W_ORDER = list(_W_COLORS.keys())

    # 示警分析只計算「去年度有購買記錄」的客戶
    # （今年才新增、去年沒買過的客戶不列入續約率計算）
    _mkt_tbl_ly = _mkt_tbl[_mkt_tbl["去年度總金額"] > 0].copy()

    _warn_counts = (
        _mkt_tbl_ly["未續約示警"]
        .replace("", "（無資料）")
        .value_counts()
        .reset_index()
        .rename(columns={"未續約示警": "狀態", "count": "客戶數"})
    )
    _warn_counts["_ord"] = _warn_counts["狀態"].map(
        {v: i for i, v in enumerate(_W_ORDER)}
    ).fillna(len(_W_ORDER))
    _warn_counts = _warn_counts.sort_values("_ord").drop(columns=["_ord"]).reset_index(drop=True)

    _wc = _warn_counts.set_index("狀態")["客戶數"].to_dict()
    _total_cust    = len(_mkt_tbl_ly)                  # 去年度客戶數（總覽用）
    _renewed_count = int(_mkt_tbl_ly["未續約示警"].str.contains("已續約", na=False).sum())
    _warn_total    = _total_cust - _renewed_count

    # 續約成功率：分母限縮為「截至今日已到期的去年度客戶」（_days < 0）
    _expired_ly          = _mkt_tbl_ly[_mkt_tbl_ly["_days"].notna() & (_mkt_tbl_ly["_days"] < 0)]
    _expired_cnt         = len(_expired_ly)
    _expired_renewed_cnt = int((_expired_ly["今年度總金額"] > 0).sum())
    _renewal_rate        = _expired_renewed_cnt / _expired_cnt * 100 if _expired_cnt > 0 else 0
    _cnt_expired   = int(_wc.get("🔴 已到期", 0))
    _cnt_15        = int(_wc.get("🟡 15天",   0))
    _cnt_30        = int(_wc.get("🟠 30天",   0))
    _cnt_45        = int(_wc.get("🟣 45天",   0))
    _cnt_60        = int(_wc.get("🟤 60天",   0))
    _cnt_90        = int(_wc.get("⚫ 90天",   0))
    _cnt_ren_up    = int(_wc.get("🟩 已續約", 0))   # 今年度 >= 去年度
    _cnt_ren_down  = int(_wc.get("🟥 已續約", 0))   # 今年度 <  去年度
    _today_str     = pd.Timestamp.today().strftime("%Y/%m/%d")
    _avg_per_stage = (_warn_total - _cnt_expired) / 5 if _warn_total > _cnt_expired else 0

    # ── 標題列 ────────────────────────────────────────────────────────────
    _rpt_h1, _rpt_h2 = st.columns([3, 1])
    with _rpt_h1:
        st.markdown(
            "<h3 style='margin-bottom:2px;'>🔴 90天內未續約示警分析報告</h3>"
            "<p style='color:#6B7280;font-size:13px;margin-top:0;'>實時監控合約到期狀況與續約表現</p>",
            unsafe_allow_html=True,
        )
    with _rpt_h2:
        st.markdown(
            f"<p style='text-align:right;color:#9CA3AF;font-size:12px;padding-top:14px;'>"
            f"🕐 數據最後更新於：{_today_str}</p>",
            unsafe_allow_html=True,
        )

    # ── KPI 三欄 ─────────────────────────────────────────────────────────
    _k1, _k2, _k3 = st.columns(3)
    for _col, _icon, _ibg, _label, _val, _vcolor, _suffix in [
        (_k1, "👥", "#EFF6FF", "去年度客戶數", f"{_total_cust:,}",       "#1E3A8A", ""),
        (_k2, "✅", "#F0FDF4", "續約成功率", f"{_renewal_rate:.1f}%",   "#16A34A", ""),
        (_k3, "⏰", "#FFF1F2", "待跟進示警", f"{_warn_total:,}",        "#DC2626", " 位"),
    ]:
        _col.markdown(
            f"""<div style="background:#fff;border-radius:14px;padding:20px 22px;
                box-shadow:0 2px 8px rgba(0,0,0,0.07);border:1px solid #F3F4F6;
                margin-bottom:4px;">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
                <span style="background:{_ibg};border-radius:10px;padding:6px 8px;
                             font-size:16px;line-height:1;">{_icon}</span>
                <span style="color:#6B7280;font-size:13px;font-weight:500;">{_label}</span>
              </div>
              <div style="font-size:30px;font-weight:800;color:{_vcolor};letter-spacing:-0.5px;">
                {_val}<span style="font-size:14px;font-weight:400;color:#9CA3AF;">{_suffix}</span>
              </div>
            </div>""",
            unsafe_allow_html=True,
        )

    st.markdown("<div style='margin-top:12px'></div>", unsafe_allow_html=True)

    # ── 中段：長條圖 ＋ 甜甜圈 ────────────────────────────────────────────
    _mid_l, _mid_r = st.columns([3, 2])

    with _mid_l:
        with st.container(border=True):
            st.markdown("<span style='font-size:13px;font-weight:700;'>續約警示時間軸（人數）</span>",
                        unsafe_allow_html=True)

            # 已續約在最下方，示警在上方；倒序讓「已到期」顯示在最頂端
            _bc_cats = [
                "🟥 已續約（減少）", "🟩 已續約（增加/持平）",
                "⚫ 90天內", "🟤 60天內", "🟣 45天內",
                "🟠 30天內", "🟡 15天內", "🔴 已到期",
            ]
            _bc_vals = [
                _cnt_ren_down, _cnt_ren_up,
                _cnt_90, _cnt_60, _cnt_45,
                _cnt_30, _cnt_15, _cnt_expired,
            ]
            _bc_clrs = [
                "#F43F5E", "#22C55E",
                "#374151", "#92400E", "#A855F7",
                "#F97316", "#EAB308", "#EF4444",
            ]

            _bar_fig = go.Figure(go.Bar(
                x=_bc_vals,
                y=_bc_cats,
                orientation="h",
                marker_color=_bc_clrs,
                text=[str(v) if v > 0 else "" for v in _bc_vals],
                textposition="outside",
                textfont=dict(size=12, color="#374151"),
                cliponaxis=False,
                hovertemplate="%{y}：%{x} 位<extra></extra>",
            ))
            # 分隔線：在「⚫ 90天內」與「🟩 已續約」之間加橫線
            _bar_fig.update_layout(
                height=370,
                margin=dict(l=10, r=60, t=10, b=10),
                paper_bgcolor="white",
                plot_bgcolor="white",
                xaxis=dict(showgrid=False, visible=False),
                yaxis=dict(showgrid=False, tickfont=dict(size=12)),
                showlegend=False,
                shapes=[dict(
                    type="line",
                    xref="paper", x0=0, x1=1,
                    yref="y",    y0=1.5, y1=1.5,
                    line=dict(color="#D1D5DB", width=1.5, dash="dot"),
                )],
            )
            st.plotly_chart(_bar_fig, use_container_width=True)

            _bs1, _bs2 = st.columns(2)
            _bs1.markdown(
                f"<div style='text-align:center;padding:6px 0;'>"
                f"<div style='font-size:12px;color:#6B7280;'>最高風險</div>"
                f"<div style='font-size:14px;font-weight:700;color:#EF4444;'>"
                f"已到期 ({_cnt_expired})</div></div>",
                unsafe_allow_html=True,
            )
            _bs2.markdown(
                f"<div style='text-align:center;padding:6px 0;'>"
                f"<div style='font-size:12px;color:#6B7280;'>平均關注</div>"
                f"<div style='font-size:14px;font-weight:700;color:#374151;'>"
                f"~{_avg_per_stage:.0f}位/階段</div></div>",
                unsafe_allow_html=True,
            )

    with _mid_r:
        with st.container(border=True):
            st.markdown("<span style='font-size:13px;font-weight:700;'>續約與示警分佈</span>",
                        unsafe_allow_html=True)

            # 甜甜圈只計算「已到期」客戶的續約 vs 未續約
            _expired_not_renewed = _expired_cnt - _expired_renewed_cnt
            _donut_fig = go.Figure(go.Pie(
                labels=["已到期已續約", "已到期未續約"],
                values=[_expired_renewed_cnt, _expired_not_renewed],
                marker=dict(
                    colors=["#22C55E", "#F97316"],
                    line=dict(color="#FFFFFF", width=3),
                ),
                hole=0.62,
                textinfo="none",
                hovertemplate="%{label}<br>%{value} 位 (%{percent})<extra></extra>",
            ))
            _donut_fig.update_layout(
                height=220,
                margin=dict(l=10, r=10, t=10, b=10),
                paper_bgcolor="white",
                showlegend=False,
                annotations=[dict(
                    text=f"<b>{_renewal_rate:.1f}%</b><br>續約率",
                    x=0.5, y=0.5, font_size=14, showarrow=False,
                )],
            )
            st.plotly_chart(_donut_fig, use_container_width=True)

            _not_renewed_rate = 100 - _renewal_rate
            for _lbl, _clr, _pct in [
                ("已到期已續約", "#22C55E", _renewal_rate),
                ("已到期未續約", "#F97316", _not_renewed_rate),
            ]:
                st.markdown(
                    f"<div style='display:flex;align-items:center;gap:8px;"
                    f"font-size:13px;margin:5px 0;'>"
                    f"<span style='width:12px;height:12px;border-radius:50%;"
                    f"background:{_clr};display:inline-block;flex-shrink:0;'></span>"
                    f"<span style='color:#374151;'>{_lbl}</span>"
                    f"<span style='margin-left:auto;font-weight:700;'>{_pct:.1f}%</span>"
                    f"</div>",
                    unsafe_allow_html=True,
                )

            if _cnt_expired > 0 and _warn_total > 0:
                _ep = _cnt_expired / _warn_total * 100
                st.markdown(
                    f"<div style='margin-top:12px;background:#FFFBEB;"
                    f"border-left:3px solid #F59E0B;border-radius:6px;"
                    f"padding:9px 11px;font-size:12px;color:#92400E;'>"
                    f"💡 提示：雖然續約率高，但「已到期」客戶佔了示警區的近 "
                    f"<b>{_ep:.0f}%</b>，應優先介入處理。</div>",
                    unsafe_allow_html=True,
                )

    st.markdown("<div style='margin-top:12px'></div>", unsafe_allow_html=True)

    # ── 後續行動建議 ──────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("<span style='font-size:13px;font-weight:700;'>後續行動建議</span>",
                    unsafe_allow_html=True)
        _ac1, _ac2, _ac3 = st.columns(3)
        for _acol, _dot, _title, _body in [
            (_ac1, "#EF4444", "🔴 立即搶救",
             f"處理「已到期」的 <b>{_cnt_expired}</b> 位客戶，確認是否已轉向競爭對手或有操作困難。"),
            (_ac2, "#F97316", "🟠 關鍵溝通",
             f"針對「15天」與「30天」內的 <b>{_cnt_15 + _cnt_30}</b> 位客戶，發送續約優惠或關懷信箱。"),
            (_ac3, "#374151", "⚫ 持續監控",
             f"針對「45-90天」客戶 <b>{_cnt_45 + _cnt_60 + _cnt_90}</b> 位進行例行維護，"
             f"確保自動續約流程正常運作。"),
        ]:
            _acol.markdown(
                f"""<div style="padding:16px 18px;border-radius:10px;
                    background:#FAFAFA;border:1px solid #E5E7EB;height:100%;">
                  <div style="font-weight:700;font-size:14px;margin-bottom:7px;">{_title}</div>
                  <div style="font-size:13px;color:#4B5563;line-height:1.6;">{_body}</div>
                </div>""",
                unsafe_allow_html=True,
            )

    # ── 數值說明 ──────────────────────────────────────────────────────────
    with st.expander("📖 數值計算說明", expanded=False):
        st.markdown("""
### 資料基礎

所有數值的母體皆來自 **客戶續約概況表（`_mkt_tbl`）**，建立流程：

```
原始資料
  → 排除 資格＝教育／非營利、訂單動作＝下單異常／已取消／已退貨
  → 套用 Sidebar 篩選（經銷商 / 最終客戶 / 展碁業務 / 日期區間 / 推廣勾選）
  → 套用 Month / Quarter 子篩選
  → 去年度 / 今年度各自依「最終客戶」加總 成交價未稅小計
  → Outer Join → 每列 = 一位客戶
```

---

### KPI 卡片

| 指標 | 計算公式 |
|---|---|
| **去年度客戶數** | 去年度有購買記錄的唯一客戶數（今年才新增的客戶不列入）|
| **續約成功率** | 截至今日**已到期**的去年度客戶中，今年度有購買的比例（已到期已續約數 ÷ 已到期客戶數 × 100%）|
| **待跟進示警** | 去年度客戶中，今年度尚未購買的客戶數 |

---

### 未續約示警判斷邏輯

依序判斷，**優先判斷已續約**：

| 狀態 | 條件 |
|---|---|
| 🟩 已續約（增加/持平） | 今年度 > 0　且　今年度 ≥ 去年度 |
| 🟥 已續約（減少） | 今年度 > 0　且　今年度 < 去年度 |
| 🔴 已到期 | 今年度 = 0，最晚到期日距今 **< 0 天** |
| 🟡 15天內 | 今年度 = 0，距今 **0 ～ 15 天** |
| 🟠 30天內 | 今年度 = 0，距今 **16 ～ 30 天** |
| 🟣 45天內 | 今年度 = 0，距今 **31 ～ 45 天** |
| 🟤 60天內 | 今年度 = 0，距今 **46 ～ 60 天** |
| ⚫ 90天內 | 今年度 = 0，距今 **61 ～ 90 天** |
| （空白）  | 今年度 = 0，距今 > 90 天，或去年度無訂閱到期日資料 |

> **到期日來源**：取該客戶在**去年度**資料中最晚的「訂閱到期日」與今天計算天數差。

---

### 甜甜圈圖

| 區塊 | 說明 |
|---|---|
| **已到期已續約** | 去年度訂閱**已到期**（到期日 < 今天），且今年度有購買記錄 |
| **已到期未續約** | 去年度訂閱**已到期**（到期日 < 今天），今年度尚無購買記錄 |

> 訂閱尚未到期（🟡🟠🟣🟤⚫ 示警中）的客戶不計入甜甜圈，因其合約仍有效，不影響目前的續約率。|

---

### 長條圖底部統計

| 指標 | 計算 |
|---|---|
| **最高風險** | 固定顯示「🔴 已到期」人數 |
| **平均關注** | （待跟進總數 − 已到期數）÷ 5　（15/30/45/60/90 天共 5 個階段的平均值）|

---

### 後續行動建議（動態數值）

| 卡片 | 引用數值 |
|---|---|
| 🔴 立即搶救 | 🔴 已到期 人數 |
| 🟠 關鍵溝通 | 🟡 15天 ＋ 🟠 30天 人數合計 |
| ⚫ 持續監控 | 🟣 45天 ＋ 🟤 60天 ＋ ⚫ 90天 人數合計 |
""")

    # ── KPI 指標 ──────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("KPI 指標（以「成交價未稅小計」分析）")
    _mkt_last_metrics = build_kpi_summary(filtered_df_sel, last_period)
    _mkt_this_metrics = build_kpi_summary(filtered_df_sel, this_period)
    render_marketing_kpi_cards(_mkt_last_metrics, _mkt_this_metrics)

    # ── 客戶續約概況表 ────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("客戶續約概況表")
    with st.expander("📋 未續約示警規則說明", expanded=False):
        st.markdown("""
**判斷基礎**：取該客戶在**去年度**資料中最晚的「訂閱到期日」，與今天計算天數差（`_days`）。

#### 已續約（今年度總金額 > 0）

| 圖示 | 條件 |
|------|------|
| 🟩 已續約 | 今年度金額 **≥** 去年度金額（持平或增加） |
| 🟥 已續約 | 今年度金額 **<** 去年度金額（減少） |

#### 未續約（今年度總金額 = 0）

| 圖示 | 示警類別 | 條件 |
|------|---------|------|
| 🔴 | 已到期 | 到期日已過（`_days < 0`） |
| 🟡 | 15 天內 | `0 ≤ _days ≤ 15` |
| 🟠 | 30 天內 | `15 < _days ≤ 30` |
| 🟣 | 45 天內 | `30 < _days ≤ 45` |
| 🟤 | 60 天內 | `45 < _days ≤ 60` |
| ⚫ | 90 天內 | `60 < _days ≤ 90` |
| （空白） | 超過 90 天或無到期日 | `_days > 90` 或資料缺失 |

> 統計範圍：僅計入**去年度有購買記錄**（去年度總金額 > 0）的客戶。今年度新客戶不列入續約分析。
""")
    st.caption(f"共 {len(_mkt_display):,} 位客戶　｜　依去年度總金額降序排列")
    st.dataframe(_mkt_display, use_container_width=True, hide_index=True, height=600)

    # ── Grouped Detail 共用欄寬設定 ────────────────────────────────────
    _gd_col_cfg = {
        "排序":           st.column_config.NumberColumn ("排序",           width=60),
        "最終客戶":       st.column_config.TextColumn   ("最終客戶",       width=200),
        "訂閱到期日年度": st.column_config.TextColumn   ("訂閱到期日年度", width=120),
        "加總金額":       st.column_config.TextColumn   ("加總金額",       width=120),
        "訂單動作":       st.column_config.TextColumn   ("訂單動作",       width=100),
        "訂閱到期日":     st.column_config.TextColumn   ("訂閱到期日",     width=110),
        "商品名稱":       st.column_config.TextColumn   ("商品名稱",       width=260),
        "數量":           st.column_config.TextColumn   ("數量",           width=80),
        "成交單價未稅":   st.column_config.TextColumn   ("成交單價未稅",   width=110),
        "成交價未稅小計": st.column_config.TextColumn   ("成交價未稅小計", width=120),
        "經銷商":         st.column_config.TextColumn   ("經銷商",         width=160),
        "展碁業務":       st.column_config.TextColumn   ("展碁業務",       width=100),
    }

    # ── 去年度客戶續約 Grouped Detail ──────────────────────────────────
    st.markdown("---")
    st.subheader("去年度客戶續約 Grouped Detail")

    _gd_last_df = _sel_last_df.copy()
    if _gd_last_df.empty:
        st.info("去年度目前篩選條件下無資料。")
    else:
        _gd_last_result, _gd_last_n_grp, _gd_last_n_rows = _build_grouped_detail(_gd_last_df)
        st.caption(f"共 {_gd_last_n_grp} 個客戶群組 ／ {_gd_last_n_rows:,} 筆明細　｜　依加總金額降序排列")
        st.dataframe(_gd_last_result, use_container_width=True, hide_index=True,
                     column_config=_gd_col_cfg)

    # ── 今年度客戶續約 Grouped Detail ──────────────────────────────────
    st.markdown("---")
    st.subheader("今年度客戶續約 Grouped Detail")

    _gd_this_df = _sel_this_df.copy()
    if _gd_this_df.empty:
        st.info("今年度目前篩選條件下無資料。")
    else:
        _gd_this_result, _gd_this_n_grp, _gd_this_n_rows = _build_grouped_detail(_gd_this_df)
        st.caption(f"共 {_gd_this_n_grp} 個客戶群組 ／ {_gd_this_n_rows:,} 筆明細　｜　依加總金額降序排列")
        st.dataframe(_gd_this_result, use_container_width=True, hide_index=True,
                     column_config=_gd_col_cfg)

    # ── Top 15 商品名稱（增減比較） ────────────────────────────────────────
    st.markdown("---")
    st.subheader("Top 15 商品名稱（增減比較）")

    _mkt_fmt_money = lambda v: f"{v/1e6:.1f}M" if abs(v) >= 1e6 else f"{v:,.0f}"
    _mkt_fmt_qty   = lambda v: f"{int(round(v)):,}"

    _mkt_chg_col1, _mkt_chg_col2 = st.columns(2)
    with _mkt_chg_col1:
        st.plotly_chart(
            _build_top15_change_fig(
                filtered_df_sel, last_period, this_period,
                "成交價未稅小計",
                "Top 15 商品名稱（金額之增減｜今年度 vs 去年度）",
                _mkt_fmt_money,
            ),
            use_container_width=True,
            key="mkt_top15_money",
        )
    with _mkt_chg_col2:
        st.plotly_chart(
            _build_top15_change_fig(
                filtered_df_sel, last_period, this_period,
                "數量",
                "Top 15 商品名稱（數量之增減｜今年度 vs 去年度）",
                _mkt_fmt_qty,
            ),
            use_container_width=True,
            key="mkt_top15_qty",
        )


# ── CSP 指標（微軟標準）分頁 ─────────────────────────────────────────────────
with tab_csp:
    st.markdown("---")
    st.subheader("CSP Renewal Dashboard（微軟標準指標）")
    st.caption(
        "依微軟 CSP Renewal Dashboard 完整定義計算。"
        "**TTM**（Trailing Twelve Months，過去12個月滾動）以 Expiration Initial（首筆）為分母；"
        "**Annualized**（年化）以 Expiration Ending（末筆）為分母。"
        "**ARR**（Annual Recurring Revenue，年度經常性收入）= 成交單價未稅 × 數量 × 12。"
    )

    _csp_m = _build_csp_metrics_full(filtered_df_sel, last_period, this_period)

    if not _csp_m:
        st.warning("資料不足，無法計算 CSP 指標（需要：最終客戶、商品名稱、成交價未稅小計、數量、訂閱到期日、訂單下單日）")
    else:
        _last_fy_c = get_fy_label_from_date(last_period.start)
        _this_fy_c = get_fy_label_from_date(this_period.start)
        _fmt_v = lambda v: f"${v/1e6:.1f}M" if abs(v) >= 1e6 else f"${v:,.0f}"

        def _rate_card(col, title, rate, desc, accent, bg):
            pct = f"{rate * 100:.1f}%"
            sc  = "#059669" if rate >= 1.0 else ("#F59E0B" if rate >= 0.8 else "#DC2626")
            sl  = "超越去年" if rate > 1.0 else ("達標" if rate >= 0.8 else "需關注")
            col.markdown(
                f"""<div style="background:{bg};border-radius:16px;border-top:5px solid {accent};
                    padding:16px 18px 12px;box-shadow:0 2px 10px rgba(0,0,0,0.06);min-height:140px;">
                  <div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:8px;line-height:1.4;">{title}</div>
                  <div style="font-size:32px;font-weight:800;color:{accent};line-height:1;">{pct}</div>
                  <div style="font-size:10.5px;color:#9CA3AF;margin:6px 0 10px;">{desc}</div>
                  <span style="padding:2px 10px;border-radius:12px;background:{sc}22;
                    font-size:11px;font-weight:700;color:{sc};">{sl}</span>
                </div>""", unsafe_allow_html=True)

        def _analysis_box(lines):
            inner = "".join(
                f'<div style="font-size:12.5px;color:{c};margin-bottom:5px;line-height:1.55;">{i} {t}</div>'
                for i, t, c in lines
            )
            st.markdown(
                f'<div style="background:#F8FAFC;border-left:4px solid #6366F1;border-radius:0 8px 8px 0;'
                f'padding:12px 16px;margin:10px 0 4px;">{inner}</div>',
                unsafe_allow_html=True
            )


        # ═══════════════════════════════════════════════════════════════════════
        # Block 1 ── TTM 核心指標  （微軟最重視的 KPI 群，置於最前）
        # ═══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.markdown("### ⭐ 1. TTM 核心指標")

        with st.expander("📖 指標定義與術語說明（TTM 系列）", expanded=False):
            st.markdown("""
**TTM（Trailing Twelve Months）過去 12 個月滾動**：以過去連續 12 個月的訂閱到期資料為計算基礎，反映最新續約健康度。

**ARR（Annual Recurring Revenue）年度經常性收入**：= 成交單價未稅 × 數量 × 12，將月費標準化至年度基準以便跨商品比較。

**CSP（Cloud Solution Provider）雲端解決方案提供商**：微軟授權的訂閱轉售計畫，由展碁等 Indirect Reseller 向終端客戶提供服務。

**Expiration Initial（到期初始）**：去年度（到期期間）每筆 客戶×商品 訂閱的**第一筆**交易年化收入/席位，作為 TTM 系列的計算分母。

**Renewal Initial（續約初始）**：今年度（續約期間）每筆已續約訂閱的**第一筆**交易年化收入/席位。

| 指標 | 英文全名 | 公式 | 計算範圍 |
|---|---|---|---|
| **TTM 年化回收率** ⭐ | TTM Annualized Recapture Rate | Renewal Initial Ann ÷ Exp Initial Ann | 僅已續約 |
| **TTM 年化續約率** | TTM Annualized Renewal Rate | Renewal Initial Ann ÷ Exp Initial Ann | 全部到期 |
| **TTM 席位續約率** | TTM Seats Renewal Rate | Renewal Initial Seats ÷ Exp Initial Seats | 全部到期 |
| **TTM ARR 成長額** | TTM ARR Growth | Renewal Initial Ann − Exp Initial Ann | 全部到期 |

> ⭐ **TTM Annualized Recapture Rate 是微軟最重視的核心指標**，代表「已成功留下的客戶，其年化收入相較去年初始水準的回收狀況」。超過 100% 代表有效 Upsell 或提價；低於 100% 代表整體收入縮水。
""")

        _sc1, _sc2, _sc3, _sc4 = st.columns(4)
        _rate_card(_sc1, "⭐ TTM 年化回收率\nTTM Annualized Recapture Rate",
                   _csp_m["ttm_ann_recapture_rate"],
                   "Renewal Initial Ann ÷ Exp Initial Ann（僅已續約）", "#EF4444", "#FFF1F2")
        _rate_card(_sc2, "TTM 年化續約率\nTTM Annualized Renewal Rate",
                   _csp_m["ttm_ann_renewal_rate"],
                   "Renewal Initial Ann ÷ Exp Initial Ann（全部到期）", "#F97316", "#FFF7ED")
        _rate_card(_sc3, "TTM 席位續約率\nTTM Seats Renewal Rate",
                   _csp_m["ttm_seats_renewal_rate"],
                   "Renewal Initial Seats ÷ Exp Initial Seats", "#8B5CF6", "#F5F3FF")
        _sc4.markdown(
            f"""<div style="background:#F0FDF4;border-radius:16px;border-top:5px solid #10B981;
                padding:16px 18px 12px;box-shadow:0 2px 10px rgba(0,0,0,0.06);min-height:140px;">
              <div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;line-height:1.4;">
                TTM ARR 成長額<br>TTM ARR Growth</div>
              <div style="font-size:24px;font-weight:800;color:{'#059669' if _csp_m['ttm_arr_growth']>=0 else '#DC2626'};line-height:1.1;">
                {'▲' if _csp_m['ttm_arr_growth']>=0 else '▼'} {_fmt_v(abs(_csp_m['ttm_arr_growth']))}</div>
              <div style="font-size:10.5px;color:#9CA3AF;margin:4px 0 6px;">Renewal − Exp Initial Ann Rev</div>
              <div style="font-size:12px;color:#374151;">席位: {'▲' if _csp_m['ttm_arr_growth_seats']>=0 else '▼'} {abs(_csp_m['ttm_arr_growth_seats']):,.0f}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<div style='margin-top:12px'></div>", unsafe_allow_html=True)

        _ttm_rec  = _csp_m["ttm_ann_recapture_rate"] * 100
        _ttm_ren  = _csp_m["ttm_ann_renewal_rate"]   * 100
        _ttm_seat = _csp_m["ttm_seats_renewal_rate"]  * 100
        _ttm_gap  = _ttm_rec - _ttm_ren
        _a1_icon  = "✅" if _ttm_rec >= 100 else ("⚠️" if _ttm_rec >= 85 else "🔴")
        _a1_clr   = "#059669" if _ttm_rec >= 100 else ("#D97706" if _ttm_rec >= 85 else "#DC2626")
        _a1_txt   = (
            f"TTM Annualized Recapture Rate {_ttm_rec:.1f}%——已續約客戶年化收入超越去年初始水準，顯示有效 Upsell 或提價。"
            if _ttm_rec >= 100 else
            f"TTM Annualized Recapture Rate {_ttm_rec:.1f}%——接近 100% 目標，仍有 {100-_ttm_rec:.1f}% 回收缺口，建議關注尚未完全回收的高值訂閱。"
            if _ttm_rec >= 85 else
            f"TTM Annualized Recapture Rate 僅 {_ttm_rec:.1f}%——年化收入回收不足，建議優先盤點流失高值客戶並啟動挽回方案。"
        )
        _a2_txt = (
            f"Recapture（{_ttm_rec:.1f}%）vs Renewal（{_ttm_ren:.1f}%）差距 {_ttm_gap:.1f}%——未續約客戶顯著拉低整體 Renewal Rate，流失客戶是主要缺口。"
            if _ttm_gap > 5 else
            f"Recapture 與 Renewal 差距僅 {_ttm_gap:.1f}%——流失客戶對整體續約率影響有限，健康。"
        )
        _a3_txt = (
            f"TTM Seats Renewal Rate {_ttm_seat:.1f}%——席位數持續成長，客戶使用規模擴大。"
            if _ttm_seat >= 100 else
            f"TTM Seats Renewal Rate {_ttm_seat:.1f}%——席位數小幅縮減，整體規模基本持平。"
            if _ttm_seat >= 90 else
            f"TTM Seats Renewal Rate {_ttm_seat:.1f}%——席位數明顯下滑，建議確認是否有大型客戶縮減授權。"
        )
        _analysis_box([
            (_a1_icon, _a1_txt, _a1_clr),
            ("📌", _a2_txt, "#374151"),
            ("💺", _a3_txt, "#374151"),
        ])

        # ═══════════════════════════════════════════════════════════════════════
        # Block 2 ── ARR Growth 分解（緊接 Block 1：TTM ARR Growth 的細部拆解）
        # ═══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.markdown("### 💡 2. ARR 成長分解（Price Effect vs Quantity Effect）")

        with st.expander("📖 指標定義——ARR 成長分解", expanded=False):
            st.markdown("""
**ARR Growth 分解**：將「已續約客戶」的 ARR 成長拆解為兩個驅動因素，幫助判斷成長品質。

| 效應 | 英文 | 中文意義 | 公式（簡化） |
|---|---|---|---|
| **價格效應** | Price Effect | 單價變動（升級、提價、降級）對 ARR 的貢獻 | ≈ (p₁ − p₀) × q₀ |
| **數量效應** | Quantity Effect | 席位增減對 ARR 的貢獻 | ≈ (q₁ − q₀) × p₀ |

> p₀ = Expiration Initial 平均單席年化單價；q₀ = Expiration Initial 席位數
> p₁ = Renewal Initial 平均單席年化單價；q₁ = Renewal Initial 席位數
> 交叉效應（Δp × Δq）按比例分配至兩者

**解讀方式：**
- Price Effect 為正 → 客戶升級或提價，代表 Upsell 成功，每席收入提升
- Quantity Effect 為正 → 客戶增購席位，代表 Expansion 成功，使用規模擴大
- 兩者均為負 → 客戶降級且縮減席位，屬高度流失風險，需緊急介入
- **理想狀態**：兩者均正，且 Price Effect 佔比較高，代表高品質成長
""")

        _pg1, _pg2, _pg3 = st.columns(3)
        _p_eff     = _csp_m["price_effect"]
        _q_eff     = _csp_m["quantity_effect"]
        _total_arr = _csp_m["arr_growth_base"]
        for _col, _label, _val, _pct in [
            (_pg1, "價格效應\nPrice Effect",        _p_eff,    _csp_m["price_pct"]),
            (_pg2, "數量效應\nQuantity Effect",     _q_eff,    _csp_m["quantity_pct"]),
            (_pg3, "ARR 成長（已續約）\nARR Growth", _total_arr, 1.0),
        ]:
            _arrow = "▲" if _val >= 0 else "▼"
            _vc = "#059669" if _val >= 0 else "#DC2626"
            _col.markdown(
                f"""<div style="border:1px solid #E5E7EB;border-radius:12px;padding:16px 18px;background:#fff;">
                  <div style="font-size:12px;color:#6B7280;margin-bottom:8px;white-space:pre-line;">{_label}</div>
                  <div style="font-size:26px;font-weight:800;color:{_vc};">{_arrow} {_fmt_v(abs(_val))}</div>
                  <div style="font-size:12px;color:#9CA3AF;margin-top:6px;">佔比 {abs(_pct)*100:.1f}%</div>
                </div>""", unsafe_allow_html=True)

        _pq_fig = go.Figure()
        _pq_fig.add_trace(go.Bar(
            x=["Price Effect\n（價格效應）", "Quantity Effect\n（數量效應）"],
            y=[_p_eff, _q_eff],
            marker_color=["#4F6BF4" if _p_eff >= 0 else "#F87171", "#10B981" if _q_eff >= 0 else "#F87171"],
            text=[f"{_fmt_v(_p_eff)} ({_csp_m['price_pct']*100:.1f}%)",
                  f"{_fmt_v(_q_eff)} ({_csp_m['quantity_pct']*100:.1f}%)"],
            textposition="outside", textfont=dict(size=12), cliponaxis=False,
        ))
        _pq_fig.update_layout(height=260, margin=dict(l=20, r=20, t=20, b=60),
                               plot_bgcolor="white", paper_bgcolor="white",
                               yaxis=dict(showgrid=True, gridcolor="#E5E7EB"),
                               showlegend=False)
        st.plotly_chart(_pq_fig, use_container_width=True)

        _dominant = "Price Effect（價格/升級驅動）" if abs(_p_eff) >= abs(_q_eff) else "Quantity Effect（席位擴張驅動）"
        _b2_lines = []
        if _p_eff >= 0 and _q_eff >= 0:
            _b2_lines.append(("✅", "Price Effect 與 Quantity Effect 均為正——已續約客戶同時做到升值與擴量，增長品質高。", "#059669"))
        elif _p_eff >= 0 > _q_eff:
            _b2_lines.append(("📌", f"Price Effect 正（{_fmt_v(_p_eff)}）但 Quantity Effect 負（{_fmt_v(_q_eff)}）——客戶傾向升級方案但縮減席位，建議關注席位流失原因。", "#D97706"))
        elif _q_eff >= 0 > _p_eff:
            _b2_lines.append(("📌", f"Quantity Effect 正（{_fmt_v(_q_eff)}）但 Price Effect 負（{_fmt_v(_p_eff)}）——客戶增加席位但有降級趨勢，需確認商品組合健康度。", "#D97706"))
        else:
            _b2_lines.append(("🔴", "Price Effect 與 Quantity Effect 均為負——已續約客戶整體縮減，建議深入分析降級與席位縮減的主要來源。", "#DC2626"))
        _b2_lines.append(("💡", f"ARR 成長主要驅動力為 {_dominant}（佔比 {max(abs(_csp_m['price_pct']), abs(_csp_m['quantity_pct']))*100:.1f}%）。", "#374151"))
        _analysis_box(_b2_lines)

        # ═══════════════════════════════════════════════════════════════════════
        # Block 3 ── 年化收入回收指標（Annualized Revenue）
        # 與 TTM 系列同屬年化收入，差異在分母（Ending vs Initial），緊接比較
        # ═══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.markdown("### 📈 3. 年化收入回收指標（Annualized Revenue）")

        with st.expander("📖 指標定義——Annualized 年化系列 vs TTM 系列", expanded=False):
            st.markdown("""
**Annualized（年化）系列**：以 **Expiration Ending（到期末筆）** 為分母，反映客戶到期前已調整的最終訂閱狀態。

**Expiration Ending（到期末筆）**：去年度期間內每筆訂閱的**最後一筆**交易年化收入/席位，最能代表客戶「實際訂閱規模」。

| 指標 | 英文 | 公式 | 計算範圍 |
|---|---|---|---|
| **年化回收率** | Annualized Recapture Rate | Renewal Initial Ann ÷ Exp **Ending** Ann | 僅已續約 |
| **年化續約率** | Annualized Renewal Rate | Renewal Initial Ann ÷ Exp **Ending** Ann | 全部到期 |

**TTM 系列 vs Annualized 系列差異：**

| 指標系列 | 分母 | 意義 |
|---|---|---|
| TTM | Expiration **Initial**（首筆） | 以「到期初始規模」為基準 |
| Annualized | Expiration **Ending**（末筆） | 以「到期最終規模」為基準 |

> 若客戶在到期期間有「增購」，Ending > Initial → Annualized 分母較大 → 比率通常低於 TTM。
> 兩者差距大，代表客戶在到期期間有顯著的席位或商品調整行為。
""")

        _ac1, _ac2, _ac3, _ac4 = st.columns(4)
        _rate_card(_ac1, "年化回收率\nAnnualized Recapture Rate",
                   _csp_m["annualized_recapture_rate"],
                   "Renewal Initial Ann ÷ Exp Ending Ann（已續約）", "#4F6BF4", "#EEF2FF")
        _rate_card(_ac2, "年化續約率\nAnnualized Renewal Rate",
                   _csp_m["annualized_renewal_rate"],
                   "Renewal Initial Ann ÷ Exp Ending Ann（全部）", "#0EA5E9", "#F0F9FF")

        _ann_rec = _csp_m["annualized_recapture_rate"] * 100
        _ann_ren = _csp_m["annualized_renewal_rate"]   * 100
        _diff_ann_ttm = _ann_rec - _ttm_rec
        _ac3.markdown(
            f"""<div style="border:1px solid #E5E7EB;border-radius:12px;padding:14px 16px;background:#fff;min-height:140px;">
              <div style="font-size:11px;color:#9CA3AF;margin-bottom:6px;">Ann vs TTM 差距</div>
              <div style="font-size:22px;font-weight:800;color:{'#059669' if _diff_ann_ttm>=0 else '#F59E0B'};">
                {'▲' if _diff_ann_ttm>=0 else '▼'} {abs(_diff_ann_ttm):.1f}%</div>
              <div style="font-size:10.5px;color:#D1D5DB;margin-top:4px;">Annualized Recapture − TTM Recapture</div>
              <div style="font-size:11px;color:#374151;margin-top:6px;">
                {'Ending > Initial：到期期間有增購' if _diff_ann_ttm < 0 else 'Initial > Ending：到期期間有縮減'}</div>
            </div>""", unsafe_allow_html=True)
        _ac4.markdown(
            f"""<div style="border:1px solid #E5E7EB;border-radius:12px;padding:14px 16px;background:#fff;min-height:140px;">
              <div style="font-size:11px;color:#9CA3AF;margin-bottom:6px;">已續約 vs 全部 差距</div>
              <div style="font-size:22px;font-weight:800;color:#6366F1;">{_ann_rec - _ann_ren:.1f}%</div>
              <div style="font-size:10.5px;color:#D1D5DB;margin-top:4px;">Recapture − Renewal Rate</div>
              <div style="font-size:11px;color:#374151;margin-top:6px;">差距越大代表流失客戶拉低整體率</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)
        _b3_lines = []
        if _ann_rec >= 100:
            _b3_lines.append(("✅", f"Annualized Recapture Rate {_ann_rec:.1f}%——以末筆為基準，已續約客戶年化收入超越到期末筆水準，Upsell 成效顯著。", "#059669"))
        elif _ann_rec >= 85:
            _b3_lines.append(("⚠️", f"Annualized Recapture Rate {_ann_rec:.1f}%——仍有 {100-_ann_rec:.1f}% 缺口，已續約客戶收入略有收縮，建議加強 Upsell 推動。", "#D97706"))
        else:
            _b3_lines.append(("🔴", f"Annualized Recapture Rate {_ann_rec:.1f}%——已續約客戶收入顯著低於去年末筆基準，建議確認是否有大量降級情形。", "#DC2626"))
        if abs(_diff_ann_ttm) > 5:
            _dir = "增購（Ending 基準較高，Annualized 比率偏低）" if _diff_ann_ttm < 0 else "縮減（Ending 基準較低，Annualized 比率偏高）"
            _b3_lines.append(("📌", f"Annualized 與 TTM Recapture Rate 差距達 {abs(_diff_ann_ttm):.1f}%——客戶在到期期間內有明顯的 {_dir} 行為。", "#374151"))
        _analysis_box(_b3_lines)

        # ═══════════════════════════════════════════════════════════════════════
        # Block 4 ── 計費收入指標（CSP Billed Revenue）
        # 與 Annualized 系列互補：Billed 反映實際計費，無年化換算
        # ═══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.markdown("### 📄 4. 計費收入指標（CSP Billed Revenue）")

        with st.expander("📖 指標定義——Billed Revenue 計費收入系列", expanded=False):
            st.markdown("""
**Billed Revenue（計費收入）**：訂閱期間內所有交易的實際計費金額（成交價未稅小計合計），包含中途調整、增購等，**不進行年化換算**。

| 指標 | 英文 | 公式 | 計算範圍 |
|---|---|---|---|
| **CSP 計費續約率** | CSP Renewal Rate | Renewal Billed ÷ Expiration Billed | 全部到期 |
| **CSP 計費回收率** | CSP Recapture Rate | Renewal Billed ÷ Expiration Billed | 僅已續約 |

**Billed vs Annualized 的差異：**

| 系列 | 特性 | 適用場景 |
|---|---|---|
| **Billed（計費）** | 反映實際收款，受訂閱期長短、中途調整影響 | 財務對帳、現金流分析 |
| **Annualized（年化）** | 標準化至年度基準，消除期數差異 | 跨商品/客戶比較、趨勢判斷 |

> 微軟建議以 **Annualized 系列**作為主要判斷依據，Billed 系列作為輔助驗證。兩者差距大時，代表訂閱期長度有顯著差異（例如年約 vs 月約混合）。
""")

        _bc1, _bc2, _bc3, _bc4 = st.columns(4)
        _rate_card(_bc1, "CSP 計費續約率\nCSP Renewal Rate",
                   _csp_m["csp_renewal_rate"],
                   "Renewal Billed ÷ Expiration Billed（全部到期）", "#10B981", "#F0FDF4")
        _rate_card(_bc2, "CSP 計費回收率\nCSP Recapture Rate",
                   _csp_m["csp_recapture_rate"],
                   "Renewal Billed ÷ Expiration Billed（僅已續約）", "#F59E0B", "#FFFBEB")

        _csp_ren = _csp_m["csp_renewal_rate"]   * 100
        _csp_rec = _csp_m["csp_recapture_rate"] * 100
        _billed_gap = _csp_rec - _csp_ren
        _bc3.markdown(
            f"""<div style="border:1px solid #E5E7EB;border-radius:12px;padding:14px 16px;background:#fff;min-height:140px;">
              <div style="font-size:11px;color:#9CA3AF;margin-bottom:6px;">Recapture − Renewal 差距</div>
              <div style="font-size:22px;font-weight:800;color:#F59E0B;">{_billed_gap:.1f}%</div>
              <div style="font-size:10.5px;color:#D1D5DB;margin-top:4px;">差距越大代表未續約客戶越多</div>
              <div style="font-size:11px;color:#374151;margin-top:6px;">
                {'差距偏大，需關注流失客戶挽回' if _billed_gap > 15 else '差距正常，流失影響可控'}</div>
            </div>""", unsafe_allow_html=True)
        _bc4.markdown(
            f"""<div style="border:1px solid #E5E7EB;border-radius:12px;padding:14px 16px;background:#fff;min-height:140px;">
              <div style="font-size:11px;color:#9CA3AF;margin-bottom:6px;">Exp Billed → Ren Billed</div>
              <div style="font-size:20px;font-weight:800;color:#1D4ED8;">
                {_fmt_v(_csp_m['exp_billed'])} → {_fmt_v(_csp_m['ren_billed'])}</div>
              <div style="font-size:10.5px;color:#D1D5DB;margin-top:4px;">到期計費 → 續約計費</div>
              <div style="font-size:11px;color:{'#059669' if _csp_m['ren_billed'] >= _csp_m['exp_billed'] else '#DC2626'};margin-top:6px;">
                {'計費收入成長' if _csp_m['ren_billed'] >= _csp_m['exp_billed'] else '計費收入縮減'}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)
        _b4_lines = [("📊", f"CSP 計費續約率 {_csp_ren:.1f}%，回收率 {_csp_rec:.1f}%（差距 {_billed_gap:.1f}%）。", "#374151")]
        if _billed_gap > 15:
            _b4_lines.append(("⚠️", f"Recapture 與 Renewal 差距達 {_billed_gap:.1f}%——有相當數量到期客戶未續約，建議強化 90 天前的主動提醒與精準行銷。", "#D97706"))
        if _csp_rec >= 100:
            _b4_lines.append(("✅", "CSP Recapture Rate ≥ 100%——已續約客戶的計費收入超越去年，有明顯增購行為。", "#059669"))
        elif _csp_rec < 85:
            _b4_lines.append(("🔴", "CSP Recapture Rate < 85%——已續約客戶計費收入縮減，建議確認是否有大量降級或縮短訂閱期情形。", "#DC2626"))
        _analysis_box(_b4_lines)

        # ═══════════════════════════════════════════════════════════════════════
        # Block 5 ── 席位續約指標（Seats）
        # 與收入指標平行：席位數是衡量使用規模的獨立維度
        # ═══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.markdown("### 🪑 5. 席位續約指標（Seats Renewal）")

        with st.expander("📖 指標定義——席位（Seats）系列與 Expansion/Upsell", expanded=False):
            st.markdown("""
**席位（Seats）**：訂閱授權數量（Microsoft 365 等雲端服務以「席位」計費，一席 = 一個使用者授權）。

**SKU（Stock Keeping Unit）商品型號**：微軟 CSP 商品的具體規格，例如「Microsoft 365 Business Premium」。

| 指標 | 英文 | 公式 | 計算範圍 |
|---|---|---|---|
| **席位續約率** | CSP Seats Renewal Rate | Renewal Initial Seats ÷ Exp **Ending** Seats | 全部到期 |
| **席位回收率** | Seats Recapture Rate | CSP Seats Renewal Rate | 僅已續約 |
| **增購型回收率** | Expansion Renewal Rate | CSP Recapture Rate（Expansion 類客戶） | 僅 Expansion |
| **升級型回收率** | Upsell Renewal Rate | CSP Recapture Rate（非 Expansion 升級類客戶） | 僅 Upsell |

**Expansion（增購）**：同一 SKU 持續續約，可能伴隨席位成長——橫向擴張，提升客戶黏性與規模。

**Upsell（升級）**：客戶換成更高階 SKU（如 Business Basic → Business Premium、E3 → E5）——縱向提升，代表更高的 ARPU（Average Revenue Per User，每用戶平均收入）。
""")

        _s5c1, _s5c2, _s5c3, _s5c4 = st.columns(4)
        _rate_card(_s5c1, "席位續約率\nCSP Seats Renewal Rate",
                   _csp_m["seats_renewal_rate"],
                   "Renewal Initial Seats ÷ Exp Ending Seats（全部）", "#6366F1", "#EEF2FF")
        _rate_card(_s5c2, "席位回收率\nSeats Recapture Rate",
                   _csp_m["seats_recapture_rate"],
                   "CSP Seats Renewal Rate（僅已續約）", "#A855F7", "#FAF5FF")
        _rate_card(_s5c3, "增購型回收率\nExpansion Renewal Rate",
                   _csp_m["expansion_renewal_rate"],
                   "Expansion（同 SKU）客戶的計費收入回收率", "#14B8A6", "#F0FDFA")
        _rate_card(_s5c4, "升級型回收率\nUpsell Renewal Rate",
                   _csp_m["upsell_renewal_rate"],
                   "非 Expansion 升級客戶的計費收入回收率", "#F43F5E", "#FFF1F2")

        st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)
        _seat_ren = _csp_m["seats_renewal_rate"]     * 100
        _seat_rec = _csp_m["seats_recapture_rate"]   * 100
        _exp_ren  = _csp_m["expansion_renewal_rate"] * 100
        _up_ren   = _csp_m["upsell_renewal_rate"]    * 100
        _b5_lines = []
        if _seat_ren >= 100:
            _b5_lines.append(("✅", f"席位續約率 {_seat_ren:.1f}%——授權席位持續成長，客戶使用規模擴大。", "#059669"))
        elif _seat_ren >= 90:
            _b5_lines.append(("⚠️", f"席位續約率 {_seat_ren:.1f}%——席位略有縮減，建議關注是否有大客戶在縮減授權規模。", "#D97706"))
        else:
            _b5_lines.append(("🔴", f"席位續約率 {_seat_ren:.1f}%——席位顯著下滑，建議盤點主要縮減來源並評估後續影響。", "#DC2626"))
        if _exp_ren > _up_ren:
            _b5_lines.append(("💡", f"Expansion Rate（{_exp_ren:.1f}%）高於 Upsell Rate（{_up_ren:.1f}%）——收入回收主力為同 SKU 增購，建議同步推動方案升級以提升 ARPU。", "#374151"))
        else:
            _b5_lines.append(("💡", f"Upsell Rate（{_up_ren:.1f}%）高於 Expansion Rate（{_exp_ren:.1f}%）——方案升級客戶帶來較高回收貢獻，升級策略奏效，可持續強化。", "#374151"))
        _analysis_box(_b5_lines)

        # ═══════════════════════════════════════════════════════════════════════
        # Block 6 ── 收入結構摘要與達成率總覽
        # 彙整所有收入維度的視覺對比，作為 Blocks 1-5 的綜合佐證
        # ═══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.markdown("### 📊 6. 收入結構摘要與達成率總覽")

        with st.expander("📖 說明——各收入欄位意義與達成率圖解讀", expanded=False):
            st.markdown("""
**各收入欄位對照：**

| 欄位名稱 | 英文 | 意義 | 用於計算 |
|---|---|---|---|
| **到期計費** | Exp Billed | 到期期間所有交易的實際計費合計 | CSP Renewal/Recapture Rate（分母）|
| **到期初始年化** | Exp Initial Ann | 到期期間第一筆交易的年化收入 | TTM 系列（分母）|
| **到期末筆年化** | Exp Ending Ann | 到期期間最後一筆交易的年化收入 | Annualized 系列（分母）|
| **續約初始年化** | Ren Initial Ann | 續約期間第一筆交易的年化收入 | 所有年化比率（分子）|
| **續約計費** | Ren Billed | 續約期間所有交易的實際計費合計 | CSP Renewal/Recapture Rate（分子）|

**達成率圖解讀（100% 虛線為基準）：**
- 🟥 超過 100% = 年化/計費收入或席位超越去年，有 Upsell/Expansion
- 🟡 80–100% = 基本維持，可接受
- ⬜ 低於 80% = 需關注，收入或席位明顯縮減
""")

        _sm = st.columns(6)
        _sm_data = [
            ("到期訂閱數",   f"{_csp_m['total_subs']:,}",            "客戶×商品 組合數",      "#6366F1"),
            ("已續約訂閱數", f"{_csp_m['renewed_count']:,}",          "去年→今年續約",          "#10B981"),
            ("到期計費收入", f"{_fmt_v(_csp_m['exp_billed'])}",       "Exp Billed",             "#F97316"),
            ("續約計費收入", f"{_fmt_v(_csp_m['ren_billed'])}",       "Ren Billed",             "#1D4ED8"),
            ("到期末筆年化", f"{_fmt_v(_csp_m['exp_ending_ann'])}",   "Exp Ending Ann（基準）", "#F59E0B"),
            ("續約初始年化", f"{_fmt_v(_csp_m['ren_initial_ann'])}",  "Ren Initial Ann",        "#EF4444"),
        ]
        for _col, (_lbl, _val, _sub, _clr) in zip(_sm, _sm_data):
            _col.markdown(
                f"""<div style="border:1px solid #E5E7EB;border-radius:12px;padding:12px 14px;
                    background:#fff;box-shadow:0 1px 6px rgba(0,0,0,0.04);">
                  <div style="font-size:10.5px;color:#9CA3AF;margin-bottom:5px;">{_lbl}</div>
                  <div style="font-size:20px;font-weight:800;color:{_clr};">{_val}</div>
                  <div style="font-size:10px;color:#D1D5DB;margin-top:3px;">{_sub}</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("<div style='margin-top:16px'></div>", unsafe_allow_html=True)

        _cv1, _cv2 = st.columns(2)
        with _cv1:
            _bar = go.Figure()
            _cats  = ["Exp\nBilled", "Exp\nInitial Ann", "Exp\nEnding Ann", "Ren\nInitial Ann", "Ren\nBilled"]
            _cvals = [_csp_m["exp_billed"], _csp_m["exp_initial_ann"], _csp_m["exp_ending_ann"],
                      _csp_m["ren_initial_ann"], _csp_m["ren_billed"]]
            _bclrs = ["#F97316", "#FBBF24", "#FDE68A", "#3B82F6", "#1D4ED8"]
            _bar.add_trace(go.Bar(x=_cats, y=_cvals, marker_color=_bclrs,
                                  text=[_fmt_v(v) for v in _cvals],
                                  textposition="outside", textfont=dict(size=11), cliponaxis=False))
            _bar.update_layout(
                title=dict(text="收入結構比較（橘=到期 藍=續約）", font=dict(size=12), x=0.5),
                height=340, margin=dict(l=10, r=10, t=45, b=50),
                plot_bgcolor="white", paper_bgcolor="white",
                yaxis=dict(showgrid=True, gridcolor="#E5E7EB"),
                xaxis=dict(showgrid=False), showlegend=False)
            st.plotly_chart(_bar, use_container_width=True)

        with _cv2:
            _rate_names_g = [
                "⭐ TTM Ann\nRecapture", "TTM Ann\nRenewal", "Ann\nRecapture", "Ann\nRenewal",
                "CSP\nRenewal", "CSP\nRecapture", "Seats\nRenewal", "Seats\nRecapture",
                "Expansion\nRenewal", "Upsell\nRenewal"]
            _rate_vals_g = [_csp_m[k] * 100 for k in [
                "ttm_ann_recapture_rate", "ttm_ann_renewal_rate",
                "annualized_recapture_rate", "annualized_renewal_rate",
                "csp_renewal_rate", "csp_recapture_rate",
                "seats_renewal_rate", "seats_recapture_rate",
                "expansion_renewal_rate", "upsell_renewal_rate"]]
            _clrs_g = ["#EF4444" if v >= 100 else ("#F59E0B" if v >= 80 else "#6B7280") for v in _rate_vals_g]
            _fig_g = go.Figure(go.Bar(
                x=_rate_vals_g, y=_rate_names_g, orientation="h",
                marker_color=_clrs_g,
                text=[f"{v:.1f}%" for v in _rate_vals_g],
                textposition="outside", textfont=dict(size=10), cliponaxis=False))
            _fig_g.add_vline(x=100, line=dict(color="#374151", dash="dash", width=1.5))
            _fig_g.update_layout(
                title=dict(text="各指標達成率（虛線=100% 基準）", font=dict(size=12), x=0.5),
                height=340, margin=dict(l=10, r=70, t=45, b=30),
                plot_bgcolor="white", paper_bgcolor="white",
                xaxis=dict(ticksuffix="%", showgrid=True, gridcolor="#E5E7EB"),
                yaxis=dict(automargin=True), showlegend=False)
            st.plotly_chart(_fig_g, use_container_width=True)

        _above100 = sum(1 for v in _rate_vals_g if v >= 100)
        _below80  = sum(1 for v in _rate_vals_g if v < 80)
        _b6_lines = [("📊", f"10 項指標中，{_above100} 項超過 100%（超越去年基準），{_below80} 項低於 80%（需關注）。", "#374151")]
        if _below80 >= 5:
            _b6_lines.append(("🔴", "超過半數指標低於 80%，整體續約健康度偏弱，建議啟動全面客戶挽回計畫。", "#DC2626"))
        elif _below80 >= 3:
            _b6_lines.append(("⚠️", f"有 {_below80} 項指標低於 80%，需優先關注對應客群的續約狀況。", "#D97706"))
        else:
            _b6_lines.append(("✅", "多數指標表現健康，建議持續維持並強化已超越 100% 的指標。", "#059669"))
        _analysis_box(_b6_lines)

        # ═══════════════════════════════════════════════════════════════════════
        # Block 7 ── Upsell Motion 分析（客戶行為分類）
        # 獨立主題：深入分析客戶行為，與收入/席位指標互補
        # ═══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.markdown("### 🔝 7. Upsell Motion 分析（客戶行為分類）")

        with st.expander("📖 Upsell Motion 分類說明", expanded=False):
            st.markdown("""
**Upsell Motion（升級行為分類）**：依客戶在去年度（到期）與今年度（續約）主要商品的跨期變化，自動分類續約行為類型，協助判斷銷售策略成效。

| Motion | 中文說明 | 觸發條件 |
|---|---|---|
| **Expansion** | 增購（同商品擴量） | 同 SKU 續約，可含席位成長 |
| **Pure Renewal** | 純續約（維持現狀） | 同 SKU，席位不變 |
| **Reduction** | 縮減（同商品縮席位） | 同 SKU，席位減少 |
| **Basic to Prem. or Std.** | 基礎版升級 | Business Basic → Standard/Premium |
| **Std. to Prem.** | 標準版升至頂規 | Business Standard → Premium |
| **ME3 to ME5** | 企業版升級 | Microsoft 365 E3 → E5 |
| **O to M** | Office 升 Microsoft 365 | Office 365 → Microsoft 365 |
| **Other Upsell** | 其他升級 | 其他跨商品升級 |
| **Not Renewed** | 未續約 | 今年度無購買記錄 |

**Expansion（增購）**：橫向擴張，同 SKU 加席位，提升客戶黏性與規模。
**Upsell（升級）**：縱向提升，換更高階 SKU，代表更高的 ARPU（Average Revenue Per User，每用戶平均收入）。
**SKU（Stock Keeping Unit）商品型號**：如「Microsoft 365 Business Premium」。
""")

        _um1, _um2, _um3, _um4 = st.columns(4)
        _rate_card(_um1, "增購型回收率\nExpansion Renewal Rate",
                   _csp_m["expansion_renewal_rate"],
                   "Expansion 客戶的 CSP Recapture Rate", "#10B981", "#F0FDF4")
        _rate_card(_um2, "升級型回收率\nUpsell Renewal Rate",
                   _csp_m["upsell_renewal_rate"],
                   "升級型客戶的 CSP Recapture Rate（非 Expansion）", "#8B5CF6", "#F5F3FF")

        _m_summary = _csp_m["motion_summary"].copy()
        _MOTION_ORDER = ["Expansion", "Pure Renewal", "Reduction", "Not Renewed",
                         "Basic to Prem. or Std.", "Std. to Prem.", "ME3 to ME5", "O to M", "Other Upsell"]
        _MOTION_COLORS = {
            "Expansion": "#10B981", "Pure Renewal": "#3B82F6", "Reduction": "#F59E0B",
            "Not Renewed": "#EF4444", "Basic to Prem. or Std.": "#8B5CF6", "Std. to Prem.": "#A855F7",
            "ME3 to ME5": "#EC4899", "O to M": "#14B8A6", "Other Upsell": "#6366F1",
        }
        _exp_cnt    = int(_m_summary.loc[_m_summary["Upsell Motion"] == "Expansion", "訂閱數"].sum()) if "Expansion" in _m_summary["Upsell Motion"].values else 0
        _upsell_cnt = int(_m_summary.loc[~_m_summary["Upsell Motion"].isin(["Expansion", "Pure Renewal", "Not Renewed", "Reduction"]), "訂閱數"].sum())
        _not_ren_cnt = int(_m_summary.loc[_m_summary["Upsell Motion"] == "Not Renewed", "訂閱數"].sum()) if "Not Renewed" in _m_summary["Upsell Motion"].values else 0
        _um3.metric("Expansion 訂閱數", f"{_exp_cnt:,}", help="同SKU增購（席位成長）")
        _um4.metric("Upsell 訂閱數",    f"{_upsell_cnt:,}", help="SKU升級（方案提升）")

        st.markdown("<div style='margin-top:12px'></div>", unsafe_allow_html=True)

        _ms_col, _ms_chart = st.columns([2, 3])
        with _ms_col:
            st.markdown("**Motion 彙總表**")
            _ms_disp = _m_summary.copy()
            for _c in ["到期計費收入", "到期年化收入", "續約年化收入"]:
                if _c in _ms_disp.columns:
                    _ms_disp[_c] = pd.to_numeric(_ms_disp[_c], errors="coerce").map(
                        lambda v: f"{v:,.0f}" if pd.notna(v) else "-")
            if "年化回收率" in _ms_disp.columns:
                _ms_disp["年化回收率"] = pd.to_numeric(_ms_disp["年化回收率"], errors="coerce").map(
                    lambda v: f"{v*100:.1f}%" if pd.notna(v) else "-")
            st.dataframe(_ms_disp, use_container_width=True, hide_index=True)

        with _ms_chart:
            st.markdown("**Motion 分佈（到期年化收入）**")
            _ms_sorted = _m_summary.sort_values("到期年化收入", ascending=True)
            _mc_colors = [_MOTION_COLORS.get(str(m), "#9CA3AF") for m in _ms_sorted["Upsell Motion"]]
            _mc_fig = go.Figure(go.Bar(
                x=pd.to_numeric(_ms_sorted["到期年化收入"], errors="coerce").fillna(0),
                y=_ms_sorted["Upsell Motion"].astype(str),
                orientation="h",
                marker_color=_mc_colors,
                text=[_fmt_v(v) for v in pd.to_numeric(_ms_sorted["到期年化收入"], errors="coerce").fillna(0)],
                textposition="outside", textfont=dict(size=11), cliponaxis=False,
            ))
            _mc_fig.update_layout(height=300, margin=dict(l=10, r=80, t=20, b=30),
                                   plot_bgcolor="white", paper_bgcolor="white",
                                   xaxis=dict(showgrid=True, gridcolor="#E5E7EB"),
                                   yaxis=dict(automargin=True), showlegend=False)
            st.plotly_chart(_mc_fig, use_container_width=True)

        _total_subs_m = _csp_m["total_subs"]
        _not_ren_pct  = _not_ren_cnt / _total_subs_m * 100 if _total_subs_m > 0 else 0
        _b7_lines = []
        if _not_ren_pct > 20:
            _b7_lines.append(("🔴", f"未續約（Not Renewed）佔 {_not_ren_pct:.1f}%——流失訂閱比例偏高，建議分析未續約客戶的商品類型與業務負責人，優先啟動挽回行動。", "#DC2626"))
        elif _not_ren_pct > 10:
            _b7_lines.append(("⚠️", f"未續約（Not Renewed）佔 {_not_ren_pct:.1f}%——仍有改善空間，建議加強 90 天前主動聯繫與精準行銷。", "#D97706"))
        else:
            _b7_lines.append(("✅", f"未續約（Not Renewed）佔 {_not_ren_pct:.1f}%——流失率低，整體續約執行良好。", "#059669"))
        if _exp_cnt > _upsell_cnt:
            _b7_lines.append(("💡", f"Expansion（{_exp_cnt:,} 筆）多於 Upsell（{_upsell_cnt:,} 筆）——成長動能主要來自加購席位，建議同步推動方案升級以提升 ARPU。", "#374151"))
        else:
            _b7_lines.append(("💡", f"Upsell（{_upsell_cnt:,} 筆）多於 Expansion（{_exp_cnt:,} 筆）——客戶方案升級積極，建議持續推動 E3→E5 及 Copilot 附加方案。", "#374151"))
        _analysis_box(_b7_lines)

        # ═══════════════════════════════════════════════════════════════════════
        # Block 8 ── Copilot Seat Attach Rate（AI 席位附掛率）
        # 獨立主題：AI 滲透率，與其他指標相對獨立，置於 Upsell 之後合理
        # ═══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.markdown("### 🤖 8. Copilot Seat Attach Rate（AI 席位附掛率）")

        with st.expander("📖 Copilot 指標說明", expanded=False):
            st.markdown("""
**Copilot（Microsoft 365 Copilot）**：微軟 AI 助理，需以 M365 E3/E5/Business Standard/Premium 為基礎授權方可附加。

**Seat Attach Rate（席位附掛率）**：在所有符合條件的席位中，實際搭配附加 Copilot 的比例——衡量 AI 滲透率的關鍵指標。

| 指標 | 英文 | 定義 |
|---|---|---|
| **Copilot 續約席位** | Renewal Copilot Seats Attached | 續約首筆交易中，Copilot 相關商品的席位數 |
| **Copilot 合規到期席位** | Copilot Eligible Expiring Seats | 到期末筆中，M365 E3/E5/Business Std/Premium 的席位（有資格加掛 Copilot）|
| **Copilot 附掛率** | Copilot Seat Attach Rate | Copilot 續約席位 ÷ 合規到期席位 |
| **有 Copilot 的客戶數** | Accounts with Copilot Attached | 續約後有 Copilot 席位的唯一客戶數 |

> 目前微軟建議 **Seat Attach Rate ≥ 10%** 為健康基準。Copilot 滲透率低的客戶，合規到期席位即為直接銷售機會。
""")

        _cop1, _cop2, _cop3, _cop4 = st.columns(4)
        _cop_rate = _csp_m["copilot_attach_rate"]
        _cop1.metric("Copilot 續約席位\nRenewal Copilot Seats",
                     f"{_csp_m['copilot_renewal_seats']:,.0f}",
                     help="續約初始席位中 Copilot 相關商品席位數")
        _cop2.metric("合規到期席位\nCopilot Eligible Seats",
                     f"{_csp_m['copilot_eligible_seats']:,.0f}",
                     help="到期末筆 M365 E3/E5/Business Std/Premium 席位")
        _cop3.metric("有 Copilot 的客戶數\nAccounts w/ Copilot",
                     f"{_csp_m['copilot_accounts']:,}",
                     help="續約後有 Copilot 席位的唯一客戶數")
        _cop_clr = "#059669" if _cop_rate >= 0.1 else ("#F59E0B" if _cop_rate >= 0.05 else "#DC2626")
        _cop4.markdown(
            f"""<div style="border:1px solid #E5E7EB;border-radius:12px;padding:14px 16px;background:#fff;">
              <div style="font-size:11px;color:#9CA3AF;margin-bottom:6px;">Copilot 附掛率<br>Seat Attach Rate</div>
              <div style="font-size:28px;font-weight:800;color:{_cop_clr};">{_cop_rate*100:.1f}%</div>
              <div style="font-size:10px;color:#D1D5DB;margin-top:4px;">Copilot Seats ÷ Eligible Seats</div>
              <div style="font-size:11px;color:#374151;margin-top:4px;">目標基準：≥ 10%</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)
        _b8_lines = []
        if _cop_rate >= 0.1:
            _b8_lines.append(("✅", f"Copilot Seat Attach Rate {_cop_rate*100:.1f}%——超過微軟建議的 10% 基準，AI 滲透率良好，建議持續推動 Copilot 在更多客戶的採購。", "#059669"))
        elif _cop_rate >= 0.05:
            _b8_lines.append(("⚠️", f"Copilot Seat Attach Rate {_cop_rate*100:.1f}%——尚未達到 10% 目標，建議針對有 M365 E3/E5 的客戶主動提案 Copilot 試用方案。", "#D97706"))
        else:
            _b8_lines.append(("🔴", f"Copilot Seat Attach Rate 僅 {_cop_rate*100:.1f}%——AI 附掛率極低，建議優先安排 Copilot 示範演示（Demo）並鎖定業務決策者為主要目標。", "#DC2626"))
        _gap_seats = _csp_m["copilot_eligible_seats"] - _csp_m["copilot_renewal_seats"]
        if _gap_seats > 0:
            _b8_lines.append(("💡", f"合規到期席位共 {_csp_m['copilot_eligible_seats']:,.0f} 席，目前 Copilot 附掛 {_csp_m['copilot_renewal_seats']:,.0f} 席，尚有 {_gap_seats:,.0f} 席未附掛——為直接銷售機會。", "#374151"))
        _analysis_box(_b8_lines)

        # ═══════════════════════════════════════════════════════════════════════
        # Block 9 ── 客戶 × 商品 完整明細（明細表置於最後，供深入查閱）
        # ═══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.subheader("📋 9. 客戶 × 商品 完整明細")
        st.caption("依「到期計費收入」降序排列；年化收入 = 成交單價未稅 × 數量 × 12")

        _ct = _csp_m["cust_table"].copy()
        for _mc in ["到期計費收入", "到期年化收入", "到期初始年化收入", "續約初始年化收入"]:
            if _mc in _ct.columns:
                _ct[_mc] = pd.to_numeric(_ct[_mc], errors="coerce").map(
                    lambda v: f"{v:,.0f}" if pd.notna(v) else "-")
        for _sc in ["到期席位數", "到期初始席位數", "續約初始席位數", "席位變化"]:
            if _sc in _ct.columns:
                _ct[_sc] = pd.to_numeric(_ct[_sc], errors="coerce").map(
                    lambda v: f"{int(round(v)):,}" if pd.notna(v) else "-")
        for _rc in ["年化回收率（Ending）", "TTM年化回收率（Initial）"]:
            if _rc in _ct.columns:
                _ct[_rc] = pd.to_numeric(_ct[_rc], errors="coerce").map(
                    lambda v: f"{v*100:.1f}%" if pd.notna(v) else "-")
        if "已續約" in _ct.columns:
            _ct["已續約"] = _ct["已續約"].map(lambda v: "✅" if v else "❌")
        st.dataframe(_ct, use_container_width=True, hide_index=True, height=540)

        # ═══════════════════════════════════════════════════════════════════════
        # Block 10 ── 全指標數值參考表（彙整所有 16 項指標數值）
        # ═══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.subheader("📑 10. 全指標數值參考表")
        st.caption("所有微軟標準 CSP Renewal 指標的計算結果彙整，供匯出或對外報告使用。")
        _all_metrics = pd.DataFrame([
            ("⭐ TTM 年化回收率", "TTM Annualized Recapture Rate",   f"{_csp_m['ttm_ann_recapture_rate']*100:.2f}%",    "Renewal Initial Ann ÷ Exp Initial Ann（已續約）",   "TTM"),
            ("TTM 年化續約率",    "TTM Annualized Renewal Rate",      f"{_csp_m['ttm_ann_renewal_rate']*100:.2f}%",      "Renewal Initial Ann ÷ Exp Initial Ann（全部）",     "TTM"),
            ("TTM ARR 成長額",    "TTM ARR Growth ($)",               f"${_csp_m['ttm_arr_growth']:,.0f}",               "Renewal Initial Ann − Exp Initial Ann",             "TTM"),
            ("TTM ARR 席位成長",  "TTM ARR Growth (Seats)",           f"{_csp_m['ttm_arr_growth_seats']:,.0f}",          "Renewal Initial Seats − Exp Initial Seats",         "TTM"),
            ("TTM 席位續約率",    "TTM Seats Renewal Rate",           f"{_csp_m['ttm_seats_renewal_rate']*100:.2f}%",    "Renewal Initial Seats ÷ Exp Initial Seats",         "TTM"),
            ("年化回收率",        "Annualized Recapture Rate",        f"{_csp_m['annualized_recapture_rate']*100:.2f}%", "Renewal Initial Ann ÷ Exp Ending Ann（已續約）",    "Annualized"),
            ("年化續約率",        "Annualized Renewal Rate",          f"{_csp_m['annualized_renewal_rate']*100:.2f}%",   "Renewal Initial Ann ÷ Exp Ending Ann（全部）",      "Annualized"),
            ("CSP 計費續約率",    "CSP Renewal Rate",                 f"{_csp_m['csp_renewal_rate']*100:.2f}%",          "Renewal Billed ÷ Expiration Billed",                "Billed"),
            ("CSP 計費回收率",    "CSP Recapture Rate",               f"{_csp_m['csp_recapture_rate']*100:.2f}%",        "CSP Renewal Rate（已續約）",                        "Billed"),
            ("席位續約率",        "CSP Seats Renewal Rate",           f"{_csp_m['seats_renewal_rate']*100:.2f}%",        "Renewal Initial Seats ÷ Exp Ending Seats",          "Seats"),
            ("席位回收率",        "Seats Recapture Rate",             f"{_csp_m['seats_recapture_rate']*100:.2f}%",      "CSP Seats Renewal Rate（已續約）",                  "Seats"),
            ("增購型回收率",      "Expansion Renewal Rate",           f"{_csp_m['expansion_renewal_rate']*100:.2f}%",    "CSP Recapture Rate（Expansion 類）",                "Upsell"),
            ("升級型回收率",      "Upsell Renewal Rate",              f"{_csp_m['upsell_renewal_rate']*100:.2f}%",       "CSP Recapture Rate（非 Expansion 升級類）",         "Upsell"),
            ("價格效應",          "Price Effect",                     f"${_csp_m['price_effect']:,.0f}",                 "ARR 成長中來自單價變動的部分",                      "Decomp"),
            ("數量效應",          "Quantity Effect",                  f"${_csp_m['quantity_effect']:,.0f}",              "ARR 成長中來自席位變動的部分",                      "Decomp"),
            ("Copilot 附掛率",    "Copilot Seat Attach Rate",         f"{_csp_m['copilot_attach_rate']*100:.2f}%",       "Copilot Renewal Seats ÷ Eligible Expiring Seats",   "Copilot"),
        ], columns=["中文名稱", "英文名稱", "數值", "說明", "類別"])
        st.dataframe(_all_metrics, use_container_width=True, hide_index=True)
