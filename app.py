import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import re
from difflib import SequenceMatcher

# =========================================================
# Page Config
# =========================================================
st.set_page_config(
    page_title="経営会議ナビ",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# Style
# =========================================================
st.markdown("""
<style>
:root {
    --bg: #f6f7fb;
    --card: #ffffff;
    --text: #111827;
    --sub: #6b7280;
    --line: #e5e7eb;
    --blue: #2563eb;
    --green: #16a34a;
    --yellow: #d97706;
    --red: #dc2626;
    --soft-blue: #eff6ff;
    --soft-green: #ecfdf5;
    --soft-yellow: #fffbeb;
    --soft-red: #fef2f2;
    --shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
    --radius-xl: 24px;
    --radius-lg: 18px;
    --radius-md: 14px;
}

html, body, [data-testid="stAppViewContainer"] {
    background: linear-gradient(180deg, #f8fafc 0%, #f3f4f6 100%);
    color: var(--text);
}

.block-container {
    max-width: 1380px;
    padding-top: 1rem;
    padding-bottom: 2rem;
}

[data-testid="stHeader"] {
    background: rgba(255,255,255,0);
}

[data-testid="stSidebar"] {
    background: #ffffff;
    border-right: 1px solid var(--line);
}

.hero {
    background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
    border: 1px solid #edf2f7;
    border-radius: var(--radius-xl);
    padding: 22px 24px;
    box-shadow: var(--shadow);
    margin-bottom: 16px;
}

.hero-title {
    font-size: clamp(28px, 4vw, 44px);
    font-weight: 900;
    letter-spacing: -0.04em;
    line-height: 1.08;
    margin-bottom: 8px;
}

.hero-sub {
    color: var(--sub);
    font-size: 14px;
    line-height: 1.8;
}

.section-card {
    background: var(--card);
    border: 1px solid #edf2f7;
    border-radius: var(--radius-lg);
    box-shadow: var(--shadow);
    padding: 16px 18px;
    margin-bottom: 14px;
}

.section-title {
    font-size: 20px;
    font-weight: 800;
    letter-spacing: -0.03em;
    margin-bottom: 4px;
}

.section-sub {
    font-size: 13px;
    color: var(--sub);
    margin-bottom: 4px;
}

.kpi {
    background: #ffffff;
    border: 1px solid #edf2f7;
    border-radius: 18px;
    padding: 18px 18px 16px 18px;
    box-shadow: var(--shadow);
    min-height: 150px;
}

.kpi-label {
    font-size: 12px;
    color: var(--sub);
    font-weight: 700;
    margin-bottom: 8px;
}

.kpi-value {
    font-size: 30px;
    font-weight: 900;
    letter-spacing: -0.04em;
    line-height: 1.1;
    margin-bottom: 6px;
}

.kpi-sub {
    font-size: 12px;
    color: var(--sub);
    line-height: 1.6;
}

.badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    font-size: 12px;
    font-weight: 800;
    border-radius: 999px;
    padding: 6px 10px;
    margin-bottom: 10px;
}

.badge-blue { background: var(--soft-blue); color: var(--blue); }
.badge-green { background: var(--soft-green); color: var(--green); }
.badge-yellow { background: var(--soft-yellow); color: var(--yellow); }
.badge-red { background: var(--soft-red); color: var(--red); }

.priority-card {
    background: #ffffff;
    border: 1px solid #edf2f7;
    border-radius: 16px;
    padding: 14px 16px;
    box-shadow: var(--shadow);
    margin-bottom: 10px;
}

.priority-title {
    font-size: 17px;
    font-weight: 800;
    margin-bottom: 8px;
}

.priority-text {
    font-size: 14px;
    color: var(--sub);
    line-height: 1.7;
    margin-bottom: 4px;
}

.info-box {
    background: #ffffff;
    border: 1px solid #edf2f7;
    border-radius: 16px;
    padding: 14px 16px;
    box-shadow: var(--shadow);
    margin-bottom: 10px;
}

.analysis-box {
    background: #f8fafc;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 12px 14px;
    margin-bottom: 10px;
}

.analysis-title {
    font-size: 15px;
    font-weight: 800;
    margin-bottom: 6px;
}

.warn {
    background: #fef2f2;
    border: 1px solid #fecaca;
    color: #991b1b;
    padding: 12px 14px;
    border-radius: 14px;
    margin-bottom: 10px;
    font-size: 13px;
    line-height: 1.7;
}

.ok {
    background: #ecfdf5;
    border: 1px solid #bbf7d0;
    color: #166534;
    padding: 12px 14px;
    border-radius: 14px;
    margin-bottom: 10px;
    font-size: 13px;
    line-height: 1.7;
}

.note {
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    color: #1d4ed8;
    padding: 12px 14px;
    border-radius: 14px;
    margin-bottom: 10px;
    font-size: 13px;
    line-height: 1.7;
}

.small-note {
    font-size: 12px;
    color: var(--sub);
}

div[data-baseweb="select"] > div,
[data-testid="stFileUploader"] {
    border-radius: 14px !important;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background: #ffffff;
    border: 1px solid #edf2f7;
    padding: 6px;
    border-radius: 14px;
    width: fit-content;
}

.stTabs [data-baseweb="tab"] {
    border-radius: 10px;
    padding: 10px 14px;
    font-weight: 700;
    color: var(--sub);
}

.stTabs [aria-selected="true"] {
    background: #f8fafc !important;
    color: var(--text) !important;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# Helpers
# =========================================================
def yen(value):
    try:
        if pd.isna(value):
            return "—"
        return f"{int(round(float(value))):,}円"
    except Exception:
        return "—"

def percent(value):
    try:
        if pd.isna(value):
            return "—"
        return f"{float(value):.1f}%"
    except Exception:
        return "—"

def safe_float(v, default=np.nan):
    try:
        if pd.isna(v):
            return default
        return float(v)
    except Exception:
        return default

def make_unique_columns(df):
    cols = []
    seen = {}
    for c in df.columns:
        c_str = str(c).strip()
        if c_str in seen:
            seen[c_str] += 1
            cols.append(f"{c_str}__{seen[c_str]}")
        else:
            seen[c_str] = 0
            cols.append(c_str)
    out = df.copy()
    out.columns = cols
    return out

def read_csv_flexible(uploaded_file):
    encodings = ["utf-8-sig", "utf-8", "cp932", "shift-jis"]
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding=enc)
            return make_unique_columns(df)
        except Exception:
            pass
    uploaded_file.seek(0)
    df = pd.read_csv(uploaded_file)
    return make_unique_columns(df)

def normalize_colname(name):
    s = str(name).strip().lower()
    s = s.replace("　", "")
    s = s.replace(" ", "")
    s = s.replace("_", "")
    s = s.replace("-", "")
    s = s.replace("／", "/")
    s = s.replace("・", "")
    s = s.replace("㈱", "株式会社")
    s = s.replace("％", "%")
    s = s.replace("累月", "累計")
    s = s.replace("予算", "計画")
    s = s.replace("実績値", "実績")
    s = re.sub(r"[()\[\]{}]", "", s)
    return s

def similarity(a, b):
    return SequenceMatcher(None, normalize_colname(a), normalize_colname(b)).ratio()

def looks_like_month_value(v):
    s = str(v).strip()
    if s in ["", "nan", "NaT", "None"]:
        return False
    patterns = [
        r"^\d{4}[-/]\d{1,2}$",
        r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$",
        r"^\d{4}年\d{1,2}月$",
        r"^\d{4}年\d{1,2}月\d{1,2}日$",
        r"^\d{6}$",
        r"^\d{4}\.\d{1,2}$"
    ]
    return any(re.match(p, s) for p in patterns)

def parse_month_series(series):
    if isinstance(series, pd.DataFrame):
        raise ValueError("月列が重複しています。CSV内の同名列を確認してください。")

    raw = series.astype(str).str.strip()
    normalized = raw.copy()
    normalized = normalized.str.replace("/", "-", regex=False)
    normalized = normalized.str.replace(".", "-", regex=False)
    normalized = normalized.str.replace("年", "-", regex=False)
    normalized = normalized.str.replace("月", "", regex=False)
    normalized = normalized.str.replace("日", "", regex=False)
    normalized = normalized.str.replace(" ", "", regex=False)

    dt = pd.to_datetime(normalized, errors="coerce")
    if dt.notna().sum() == 0:
        dt = pd.to_datetime(raw, errors="coerce")

    return dt.dt.to_period("M")

def month_display(period_series):
    return period_series.astype(str)

def strict_numeric_from_string(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if s == "":
        return np.nan

    s = s.replace(",", "")
    s = s.replace("¥", "")
    s = s.replace("￥", "")
    s = s.replace("円", "")
    s = s.replace("%", "")
    s = s.replace(" ", "")
    s = s.replace("△", "-")
    s = s.replace("▲", "-")
    s = s.replace("−", "-")

    if re.fullmatch(r"\(.*\)", s):
        s = "-" + s[1:-1]

    if re.fullmatch(r"[-+]?\d+(\.\d+)?", s):
        try:
            return float(s)
        except Exception:
            return np.nan
    return np.nan

def to_numeric_safe(series):
    if isinstance(series, pd.DataFrame):
        raise ValueError("数値列が重複しています。CSV内の同名列を確認してください。")
    return series.apply(strict_numeric_from_string)

def non_null_ratio(series):
    if len(series) == 0:
        return 0
    return series.notna().mean()

def month_ratio(series):
    if len(series) == 0:
        return 0
    sample = series.dropna().astype(str).head(50)
    if len(sample) == 0:
        return 0
    hits = sum(looks_like_month_value(v) for v in sample)
    return hits / len(sample)

def value_profile_score(series):
    numeric = to_numeric_safe(series)
    valid = numeric.dropna()
    if len(valid) == 0:
        return {"numeric_ratio": 0, "mean": np.nan, "std": np.nan, "positive_ratio": 0}
    return {
        "numeric_ratio": numeric.notna().mean(),
        "mean": valid.mean(),
        "std": valid.std() if len(valid) > 1 else 0,
        "positive_ratio": (valid >= 0).mean()
    }

# =========================================================
# Detection Rules
# =========================================================
METRIC_RULES = {
    "month": {
        "keywords": ["対象月", "年月", "計上月", "月", "date", "期間", "月次", "伝票月", "発生日", "計画月", "予算月"],
        "exclude": ["率", "%", "前年差", "前年差率", "構成比"],
        "prefer": ["対象月", "年月", "月"]
    },
    "sales": {
        "keywords": ["売上計画", "売上予算", "計画売上", "売上高", "売上", "営業収益", "売上金額", "月次売上", "収益", "sales", "営業売上", "部門別売上"],
        "exclude": ["差額", "前年差", "粗利", "利益率", "計画比", "%", "原価", "販管費", "営業利益"],
        "prefer": ["売上計画", "売上高", "売上", "部門別売上"]
    },
    "cost": {
        "keywords": ["原価計画", "原価予算", "計画原価", "売上原価", "原価", "仕入原価", "cost", "原材料費", "外注原価", "製造原価"],
        "exclude": ["率", "%", "粗利", "売上", "利益", "差額"],
        "prefer": ["原価計画", "売上原価", "原価"]
    },
    "gross": {
        "keywords": ["粗利計画", "粗利予算", "計画粗利", "粗利益", "粗利", "grossprofit", "gross", "商品別粗利"],
        "exclude": ["率", "%", "差額", "営業利益", "売上計画比"],
        "prefer": ["粗利計画", "粗利", "商品別粗利"]
    },
    "sga": {
        "keywords": ["販管費計画", "販管費予算", "計画販管費", "販管費", "販売管理費", "販売費及び一般管理費", "販売費", "一般管理費", "sga"],
        "exclude": ["率", "%", "差額", "計画比", "売上", "粗利", "営業利益"],
        "prefer": ["販管費計画", "販管費"]
    },
    "op": {
        "keywords": ["営業利益計画", "営業利益予算", "計画営業利益", "営業利益", "営業損益", "operatingprofit", "operating", "op", "営業利益額"],
        "exclude": ["率", "%", "差額", "粗利率", "計画比"],
        "prefer": ["営業利益計画", "営業利益"]
    }
}

def score_column_name(col, rule):
    col_n = normalize_colname(col)
    score = 0
    for ex in rule.get("exclude", []):
        if normalize_colname(ex) in col_n:
            score -= 120
    for i, kw in enumerate(rule.get("keywords", [])):
        kw_n = normalize_colname(kw)
        if col_n == kw_n:
            score += 180 - i
        elif kw_n in col_n:
            score += 90 - min(i, 25)
        else:
            sim = similarity(col_n, kw_n)
            if sim >= 0.78:
                score += 35
    for pk in rule.get("prefer", []):
        if normalize_colname(pk) in col_n:
            score += 15
    return score

def score_column_by_values(series, metric_name):
    score = 0
    profile = value_profile_score(series)

    if metric_name == "month":
        score += month_ratio(series) * 220
        score += non_null_ratio(series) * 15
        return score

    score += profile["numeric_ratio"] * 80
    vals = to_numeric_safe(series).dropna()
    if len(vals) == 0:
        return score

    abs_mean = np.abs(vals).mean()
    if abs_mean >= 1000:
        score += 20
    if (vals != 0).mean() > 0.5:
        score += 12

    return score

def infer_flow_type_from_name(col_name):
    n = normalize_colname(col_name)
    if "累計" in n or "累月" in n or ("累" in n and "計" in n):
        return "累計"
    if "当月" in n or "単月" in n or "月次" in n:
        return "単月"
    return "不明"

def infer_flow_type_from_values(df, month_col, value_col):
    try:
        temp = pd.DataFrame({
            "月": parse_month_series(df[month_col]),
            "値": to_numeric_safe(df[value_col])
        }).dropna(subset=["月", "値"])

        if len(temp) == 0:
            return "不明", 0.0

        agg = temp.groupby("月", as_index=False)["値"].sum().sort_values("月")
        vals = agg["値"].values

        if len(vals) < 3:
            return "不明", 0.0

        diffs = np.diff(vals)
        non_decreasing_ratio = (diffs >= 0).mean()
        large_reset_ratio = (np.abs(diffs) > (np.abs(vals[:-1]) * 0.8)).mean()

        if non_decreasing_ratio >= 0.8 and large_reset_ratio <= 0.2:
            return "累計候補", non_decreasing_ratio
        return "単月候補", 1 - non_decreasing_ratio
    except Exception:
        return "不明", 0.0

def detect_best_column_advanced(df, metric_name):
    rule = METRIC_RULES[metric_name]
    results = []
    for col in df.columns:
        name_score = score_column_name(col, rule)
        value_score = score_column_by_values(df[col], metric_name)
        total = name_score + value_score
        results.append({
            "column": col,
            "name_score": round(name_score, 1),
            "value_score": round(value_score, 1),
            "total_score": round(total, 1)
        })
    results = sorted(results, key=lambda x: x["total_score"], reverse=True)
    best = results[0]["column"] if len(results) > 0 else None
    confidence = results[0]["total_score"] - results[1]["total_score"] if len(results) >= 2 else (results[0]["total_score"] if results else 0)
    return best, results[:6], confidence

def confidence_label(score_gap):
    if score_gap >= 60:
        return "高"
    elif score_gap >= 25:
        return "中"
    return "低"

def build_auto_mapping(df, source_name):
    mapping = {}
    details = {}

    for metric in ["month", "sales", "cost", "gross", "sga", "op"]:
        best, ranked, conf = detect_best_column_advanced(df, metric)
        mapping[metric] = best
        details[metric] = {
            "best": best,
            "ranked": ranked,
            "confidence_gap": conf,
            "confidence_label": confidence_label(conf)
        }

    flows = {}
    month_col = mapping["month"]
    for metric in ["sales", "cost", "gross", "sga", "op"]:
        col = mapping.get(metric)
        if col and month_col:
            name_hint = infer_flow_type_from_name(col)
            value_hint, value_conf = infer_flow_type_from_values(df, month_col, col)
            if name_hint == "累計":
                flow = "累計"
            elif name_hint == "単月":
                flow = "単月"
            elif value_hint == "累計候補" and value_conf >= 0.85:
                flow = "累計"
            else:
                flow = "単月"
        else:
            flow = "単月"
        flows[metric] = flow

    mapping["flows"] = flows
    mapping["source_name"] = source_name
    return mapping, details

# =========================================================
# Validation / Integrity
# =========================================================
def relative_error(a, b):
    a = np.array(a, dtype=float)
    b = np.array(b, dtype=float)
    denom = np.maximum(np.abs(b), 1.0)
    return np.abs(a - b) / denom

def choose_best_series(label, raw_series, calc_series, tolerance=0.05):
    raw_valid = raw_series.notna().sum() > 0
    calc_valid = calc_series.notna().sum() > 0

    if not raw_valid and calc_valid:
        return calc_series, "calculated", f"{label}列が見つからないため計算値を採用"

    if raw_valid and not calc_valid:
        return raw_series, "csv", f"{label}列のみ利用可能のためCSV値を採用"

    if raw_valid and calc_valid:
        err = relative_error(raw_series.fillna(0), calc_series.fillna(0))
        avg_err = np.nanmean(err) if len(err) > 0 else 0
        if avg_err <= tolerance:
            return raw_series, "csv", f"{label}列が計算値と整合したためCSV値を採用"
        return calc_series, "calculated", f"{label}列が計算値と乖離したため計算値を採用"

    return pd.Series([np.nan] * len(raw_series)), "missing", f"{label}を確定できません"

def detect_unit_scale(series):
    vals = pd.Series(series).dropna().astype(float)
    if len(vals) == 0:
        return None
    med = np.median(np.abs(vals))
    if med >= 100000:
        return "円"
    if 100 <= med < 100000:
        return "千円候補"
    return "小さすぎる"

def detect_unit_mismatch(freee_sales, bixid_sales):
    fs = detect_unit_scale(freee_sales)
    bs = detect_unit_scale(bixid_sales)
    if fs is None or bs is None:
        return None
    if fs == "円" and bs == "千円候補":
        return "freeeは円、bixidは千円の可能性があります。単位を確認してください。"
    if fs == "千円候補" and bs == "円":
        return "freeeは千円、bixidは円の可能性があります。単位を確認してください。"
    return None

def anomaly_check(latest):
    messages = []
    severe = False

    sales = safe_float(latest.get("売上_実績"), np.nan)
    cost = safe_float(latest.get("原価_実績"), np.nan)
    gross = safe_float(latest.get("粗利_実績"), np.nan)
    sga = safe_float(latest.get("販管費_実績"), np.nan)
    op = safe_float(latest.get("営業利益_実績"), np.nan)

    gross_rate = safe_float(latest.get("粗利率(%)"), np.nan)
    op_rate = safe_float(latest.get("営業利益率(%)"), np.nan)

    if not pd.isna(gross_rate) and gross_rate > 100:
        severe = True
        messages.append(f"粗利率が {gross_rate:.1f}% です。通常値を大きく超えているため、表示を停止します。")

    if not pd.isna(sales) and not pd.isna(gross) and sales > 0 and gross > sales:
        severe = True
        messages.append("粗利が売上を超えています。列定義または計算元データに問題があるため、表示を停止します。")

    if not pd.isna(gross) and not pd.isna(op) and op > gross:
        severe = True
        messages.append("営業利益が粗利を超えています。列定義または計算元データに問題があるため、表示を停止します。")

    if not pd.isna(cost) and cost < 0:
        severe = True
        messages.append("原価がマイナスです。符号処理または列定義に問題があるため、表示を停止します。")

    if not pd.isna(sga) and sga < 0:
        severe = True
        messages.append("販管費がマイナスです。符号処理または列定義に問題があるため、表示を停止します。")

    if not pd.isna(op_rate) and op_rate > 100:
        severe = True
        messages.append(f"営業利益率が {op_rate:.1f}% です。異常値のため、表示を停止します。")

    return severe, messages

# =========================================================
# Data Build
# =========================================================
def monthly_from_flow(df, month_col, value_col, flow_type="単月"):
    temp = pd.DataFrame({
        "月period": parse_month_series(df[month_col]),
        "値": to_numeric_safe(df[value_col])
    }).dropna(subset=["月period", "値"])

    if len(temp) == 0:
        return pd.DataFrame(columns=["月period", "値"])

    monthly_sum = (
        temp.groupby("月period", as_index=False)["値"]
        .sum()
        .sort_values("月period")
        .reset_index(drop=True)
    )

    monthly_max = (
        temp.groupby("月period", as_index=False)["値"]
        .max()
        .sort_values("月period")
        .reset_index(drop=True)
    )

    result = monthly_max if flow_type == "累計" else monthly_sum
    return result

def safe_monthly_metric(df, month_col, value_col, flow_type="単月"):
    if value_col in [None, "使わない"]:
        return pd.DataFrame(columns=["月period", "値"])

    return monthly_from_flow(df, month_col, value_col, flow_type)

def build_freee_actual_monthly(df, mapping):
    month_col = mapping["month"]

    sales_col = mapping["sales"]
    cost_col = mapping["cost"]
    gross_col = mapping["gross"]
    sga_col = mapping["sga"]
    op_col = mapping["op"]

    sales_df = safe_monthly_metric(df, month_col, sales_col, "単月").rename(columns={"値": "売上_実績_raw"})
    cost_df = safe_monthly_metric(df, month_col, cost_col, "単月").rename(columns={"値": "原価_実績_raw"})
    gross_df = safe_monthly_metric(df, month_col, gross_col, "単月").rename(columns={"値": "粗利_実績_raw"})
    sga_df = safe_monthly_metric(df, month_col, sga_col, "単月").rename(columns={"値": "販管費_実績_raw"})
    op_df = safe_monthly_metric(df, month_col, op_col, "単月").rename(columns={"値": "営業利益_実績_raw"})

    frames = [sales_df, cost_df, gross_df, sga_df, op_df]
    result = None
    for f in frames:
        if result is None:
            result = f
        else:
            result = result.merge(f, on="月period", how="outer")

    if result is None or len(result) == 0:
        return pd.DataFrame()

    for c in ["売上_実績_raw", "原価_実績_raw", "粗利_実績_raw", "販管費_実績_raw", "営業利益_実績_raw"]:
        if c not in result.columns:
            result[c] = np.nan

    calc_gross = result["売上_実績_raw"] - result["原価_実績_raw"]
    chosen_gross, gross_source, gross_note = choose_best_series("粗利", result["粗利_実績_raw"], calc_gross)

    result["粗利_実績"] = chosen_gross
    result["粗利_実績_source"] = gross_source
    result["粗利_実績_note"] = gross_note

    calc_op_from_sga = result["粗利_実績"] - result["販管費_実績_raw"]
    chosen_op, op_source, op_note = choose_best_series("営業利益", result["営業利益_実績_raw"], calc_op_from_sga)

    result["営業利益_実績"] = chosen_op
    result["営業利益_実績_source"] = op_source
    result["営業利益_実績_note"] = op_note

    result["売上_実績"] = result["売上_実績_raw"]
    result["原価_実績"] = result["原価_実績_raw"]
    result["販管費_実績"] = result["販管費_実績_raw"]

    result["月"] = month_display(result["月period"])
    result = result.sort_values("月period").reset_index(drop=True)

    return result[[
        "月period", "月",
        "売上_実績", "原価_実績", "粗利_実績", "販管費_実績", "営業利益_実績",
        "粗利_実績_source", "粗利_実績_note",
        "営業利益_実績_source", "営業利益_実績_note"
    ]]

def build_bixid_plan_monthly(df, mapping):
    month_col = mapping["month"]
    flows = mapping["flows"]

    sales = safe_monthly_metric(df, month_col, mapping.get("sales"), flows.get("sales", "単月")).rename(columns={"値": "売上_計画_raw"})
    cost = safe_monthly_metric(df, month_col, mapping.get("cost"), flows.get("cost", "単月")).rename(columns={"値": "原価_計画_raw"})
    gross = safe_monthly_metric(df, month_col, mapping.get("gross"), flows.get("gross", "単月")).rename(columns={"値": "粗利_計画_raw"})
    sga = safe_monthly_metric(df, month_col, mapping.get("sga"), flows.get("sga", "単月")).rename(columns={"値": "販管費_計画_raw"})
    op = safe_monthly_metric(df, month_col, mapping.get("op"), flows.get("op", "単月")).rename(columns={"値": "営業利益_計画_raw"})

    frames = [sales, cost, gross, sga, op]
    result = None
    for f in frames:
        if result is None:
            result = f
        else:
            result = result.merge(f, on="月period", how="outer")

    if result is None or len(result) == 0:
        return pd.DataFrame()

    for c in ["売上_計画_raw", "原価_計画_raw", "粗利_計画_raw", "販管費_計画_raw", "営業利益_計画_raw"]:
        if c not in result.columns:
            result[c] = np.nan

    calc_gross = result["売上_計画_raw"] - result["原価_計画_raw"]
    chosen_gross, gross_source, gross_note = choose_best_series("粗利計画", result["粗利_計画_raw"], calc_gross)

    result["粗利_計画"] = chosen_gross
    result["粗利_計画_source"] = gross_source
    result["粗利_計画_note"] = gross_note

    calc_op = result["粗利_計画"] - result["販管費_計画_raw"]
    chosen_op, op_source, op_note = choose_best_series("営業利益計画", result["営業利益_計画_raw"], calc_op)

    result["営業利益_計画"] = chosen_op
    result["営業利益_計画_source"] = op_source
    result["営業利益_計画_note"] = op_note

    result["売上_計画"] = result["売上_計画_raw"]
    result["原価_計画"] = result["原価_計画_raw"]
    result["販管費_計画"] = result["販管費_計画_raw"]

    result["月"] = month_display(result["月period"])
    result = result.sort_values("月period").reset_index(drop=True)

    return result[[
        "月period", "月",
        "売上_計画", "原価_計画", "粗利_計画", "販管費_計画", "営業利益_計画",
        "粗利_計画_source", "粗利_計画_note",
        "営業利益_計画_source", "営業利益_計画_note"
    ]]

# =========================================================
# Analytics
# =========================================================
def plan_rate(actual, plan):
    if pd.isna(actual) or pd.isna(plan) or plan == 0:
        return np.nan
    return actual / plan * 100

def calc_growth_rate(current, prev):
    if pd.isna(current) or pd.isna(prev) or prev == 0:
        return np.nan
    return (current / prev - 1) * 100

def delta_label(current, prev):
    if pd.isna(current) or pd.isna(prev):
        return "前月比較なし"
    diff = current - prev
    if diff > 0:
        return f"前月比 +{yen(diff)}"
    elif diff < 0:
        return f"前月比 {yen(diff)}"
    return "前月比 ±0円"

def diff_badge(diff):
    if pd.isna(diff):
        return "badge-blue", "比較不可"
    if diff > 0:
        return "badge-green", "計画超過"
    elif diff < 0:
        return "badge-red", "未達"
    return "badge-blue", "計画通り"

def build_headline(latest):
    sales_ratio = safe_float(latest["売上計画比(%)"], np.nan)
    gross_ratio = safe_float(latest["粗利率(%)"], np.nan)
    op_ratio = safe_float(latest["営業利益計画比(%)"], np.nan)
    margin = safe_float(latest["営業利益率(%)"], np.nan)

    score = 0
    notes = []

    if not pd.isna(sales_ratio):
        if sales_ratio >= 100:
            score += 1
            notes.append("売上は計画達成")
        elif sales_ratio >= 90:
            notes.append("売上は概ね計画線")
        else:
            score -= 1
            notes.append("売上は未達")

    if not pd.isna(gross_ratio):
        if 35 <= gross_ratio <= 80:
            score += 1
            notes.append("粗利率は健全")
        elif gross_ratio > 80:
            notes.append("粗利率は異常値の可能性")
        elif gross_ratio >= 30:
            notes.append("粗利率は注意水準")
        else:
            score -= 1
            notes.append("粗利率は低め")

    if not pd.isna(op_ratio):
        if op_ratio >= 100:
            score += 1
            notes.append("利益は計画達成")
        elif op_ratio >= 90:
            notes.append("利益は要監視")
        else:
            score -= 1
            notes.append("利益は未達")

    if not pd.isna(margin):
        if margin >= 15:
            notes.append("利益率は良好")
        elif margin < 8:
            notes.append("利益率は低め")

    if score >= 2:
        return "今月は攻めの判断をしやすい状態です", " / ".join(notes[:4]), "badge-green"
    elif score >= 0:
        return "今月は守りと攻めのバランス判断が必要です", " / ".join(notes[:4]), "badge-yellow"
    return "今月は守りを優先した判断が必要です", " / ".join(notes[:4]), "badge-red"

def build_cross_analysis(latest, prev):
    analyses = []

    sales_ratio = safe_float(latest["売上計画比(%)"], np.nan)
    gross_ratio = safe_float(latest["粗利率(%)"], np.nan)
    sga_ratio = safe_float(latest["販管費率(%)"], np.nan)
    op_ratio = safe_float(latest["営業利益計画比(%)"], np.nan)
    op_margin = safe_float(latest["営業利益率(%)"], np.nan)

    sales_growth = calc_growth_rate(latest["売上_実績"], prev["売上_実績"]) if prev is not None else np.nan
    gross_growth = calc_growth_rate(latest["粗利_実績"], prev["粗利_実績"]) if prev is not None else np.nan
    op_growth = calc_growth_rate(latest["営業利益_実績"], prev["営業利益_実績"]) if prev is not None else np.nan

    if not pd.isna(gross_ratio) and gross_ratio > 100:
        analyses.append({
            "title": "粗利率が100%超です",
            "text": f"粗利率は {percent(gross_ratio)} です。列定義または符号処理に問題がある可能性があります。"
        })

    if not pd.isna(sales_ratio) and not pd.isna(gross_ratio):
        if sales_ratio < 100 and gross_ratio < 30:
            analyses.append({
                "title": "売上未達と粗利率低下が同時発生",
                "text": f"売上計画比は {percent(sales_ratio)}、粗利率は {percent(gross_ratio)} です。営業量だけでなく価格・原価・商品構成も見直し対象です。"
            })
        elif sales_ratio >= 100 and gross_ratio < 30:
            analyses.append({
                "title": "売上は取れているが利益の質が弱い",
                "text": f"売上計画比は {percent(sales_ratio)}、粗利率は {percent(gross_ratio)} です。値引きや原価上昇が利益を削っている可能性があります。"
            })
        elif sales_ratio < 90 and 35 <= gross_ratio <= 80:
            analyses.append({
                "title": "中心課題は売上量の不足",
                "text": f"売上計画比は {percent(sales_ratio)}、粗利率は {percent(gross_ratio)} です。件数・受注率が先の論点です。"
            })

    if not pd.isna(gross_ratio) and not pd.isna(sga_ratio):
        if gross_ratio < 30 and sga_ratio > 30:
            analyses.append({
                "title": "粗利不足に対して販管費負担が重い",
                "text": f"粗利率 {percent(gross_ratio)} に対し販管費率は {percent(sga_ratio)} です。固定費の重さが利益を圧迫しています。"
            })
        elif 35 <= gross_ratio <= 80 and sga_ratio > 30:
            analyses.append({
                "title": "粗利はあるが販管費で利益が残りにくい",
                "text": f"粗利率は {percent(gross_ratio)}、販管費率は {percent(sga_ratio)} です。広告費・人件費・外注費の優先順位見直しが必要です。"
            })

    if not pd.isna(sales_ratio) and not pd.isna(op_ratio):
        if sales_ratio >= 100 and op_ratio < 90:
            analyses.append({
                "title": "売上達成でも利益未達",
                "text": f"売上計画比は {percent(sales_ratio)}、営業利益計画比は {percent(op_ratio)} です。売上増が利益増に直結していません。"
            })

    if not pd.isna(sales_growth) and not pd.isna(gross_growth):
        if sales_growth > 0 and gross_growth < 0:
            analyses.append({
                "title": "売上増でも粗利減",
                "text": f"前月比で売上は {sales_growth:.1f}% 増、粗利は {gross_growth:.1f}% です。値引き・原価上昇・案件ミックス変化の可能性があります。"
            })

    if not pd.isna(op_growth) and op_growth < -20:
        analyses.append({
            "title": "営業利益の落ち込みが大きい",
            "text": f"営業利益は前月比 {op_growth:.1f}% です。単月要因か、販管費増か、粗利率低下かを切り分けてください。"
        })

    if len(analyses) == 0:
        analyses.append({
            "title": "大きな構造歪みは限定的",
            "text": f"売上計画比 {percent(sales_ratio)}、粗利率 {percent(gross_ratio)}、営業利益率 {percent(op_margin)} を見る限り、極端な崩れは見えていません。"
        })

    return analyses[:4]

def build_priority_topics(latest, prev):
    sales_ratio = safe_float(latest["売上計画比(%)"], 0)
    gross_ratio = safe_float(latest["粗利率(%)"], 0)
    sga_ratio = safe_float(latest["販管費率(%)"], 0)
    op_ratio = safe_float(latest["営業利益計画比(%)"], 0)
    op_margin = safe_float(latest["営業利益率(%)"], 0)

    sales_diff = safe_float(latest["売上差額"], 0)
    gross_diff = safe_float(latest["粗利差額"], 0)
    op_diff = safe_float(latest["営業利益差額"], 0)

    sales_growth = calc_growth_rate(latest["売上_実績"], prev["売上_実績"]) if prev is not None else np.nan
    gross_growth = calc_growth_rate(latest["粗利_実績"], prev["粗利_実績"]) if prev is not None else np.nan
    op_growth = calc_growth_rate(latest["営業利益_実績"], prev["営業利益_実績"]) if prev is not None else np.nan

    topics = []

    score_sales = 0
    if sales_ratio < 90:
        score_sales += 45
    elif sales_ratio < 100:
        score_sales += 20
    if sales_diff < 0:
        score_sales += min(abs(sales_diff) / 50000, 20)
    if not pd.isna(sales_growth) and sales_growth < -10:
        score_sales += 15
    topics.append({
        "score": score_sales,
        "title": "売上未達の主因確認",
        "why": f"売上計画比は {percent(sales_ratio)}、差額は {yen(sales_diff)} です。",
        "focus": "件数・単価・受注率のどこが崩れたか",
        "decision": "営業導線・案件化率・集客投資のどこに先に手を入れるか"
    })

    score_gross = 0
    if gross_ratio > 100:
        score_gross += 70
    elif gross_ratio < 30:
        score_gross += 45
    elif gross_ratio < 35:
        score_gross += 20
    if gross_diff < 0:
        score_gross += min(abs(gross_diff) / 30000, 20)
    if not pd.isna(gross_growth) and gross_growth < -10:
        score_gross += 15
    if sales_ratio >= 100 and gross_ratio < 30:
        score_gross += 10
    topics.append({
        "score": score_gross,
        "title": "粗利率・原価構造の確認",
        "why": f"粗利率は {percent(gross_ratio)}、粗利差額は {yen(gross_diff)} です。",
        "focus": "価格・原価・商品ミックスのどこで利益が薄くなっているか",
        "decision": "値付け改善か原価改善か、どちらを先にやるか"
    })

    score_profit = 0
    if op_ratio < 90:
        score_profit += 45
    elif op_ratio < 100:
        score_profit += 20
    if op_diff < 0:
        score_profit += min(abs(op_diff) / 30000, 20)
    if sga_ratio > 30:
        score_profit += 15
    if not pd.isna(op_growth) and op_growth < -15:
        score_profit += 15
    topics.append({
        "score": score_profit,
        "title": "利益未達と販管費負担の確認",
        "why": f"営業利益計画比は {percent(op_ratio)}、営業利益率は {percent(op_margin)}、販管費率は {percent(sga_ratio)} です。",
        "focus": "利益を削っている費用が一時的か固定的か",
        "decision": "止める費用と続ける投資を分ける"
    })

    score_growth = 0
    if sales_ratio >= 100 and 35 <= gross_ratio <= 80 and op_ratio >= 100:
        score_growth = 40
    topics.append({
        "score": score_growth,
        "title": "次の成長投資の優先順位",
        "why": f"売上計画比 {percent(sales_ratio)}、粗利率 {percent(gross_ratio)}、営業利益計画比 {percent(op_ratio)} です。",
        "focus": "再現性のある利益が出る領域に資源配分できているか",
        "decision": "採用・営業強化・新規投資のどれを先に進めるか"
    })

    topics = sorted(topics, key=lambda x: x["score"], reverse=True)
    return topics[:3]

# =========================================================
# UI Helpers
# =========================================================
def render_kpi(title, value, sub, badge_class=None, badge_text=None):
    badge_html = f'<div class="badge {badge_class}">{badge_text}</div>' if badge_class and badge_text else ""
    st.markdown(f"""
    <div class="kpi">
        {badge_html}
        <div class="kpi-label">{title}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)

def summary_table(latest, merged_info):
    return pd.DataFrame([
        {"項目": "売上", "実績": yen(latest["売上_実績"]), "計画": yen(latest["売上_計画"]), "差額": yen(latest["売上差額"]), "計画比": percent(latest["売上計画比(%)"])},
        {"項目": "粗利", "実績": yen(latest["粗利_実績"]), "計画": yen(latest["粗利_計画"]), "差額": yen(latest["粗利差額"]), "計画比": percent(plan_rate(latest["粗利_実績"], latest["粗利_計画"]))},
        {"項目": "営業利益", "実績": yen(latest["営業利益_実績"]), "計画": yen(latest["営業利益_計画"]), "差額": yen(latest["営業利益差額"]), "計画比": percent(latest["営業利益計画比(%)"])},
        {"項目": "粗利 実績出所", "実績": merged_info["gross_actual_source"], "計画": "—", "差額": "—", "計画比": "—"},
        {"項目": "営業利益 実績出所", "実績": merged_info["op_actual_source"], "計画": "—", "差額": "—", "計画比": "—"},
        {"項目": "粗利率", "実績": percent(latest["粗利率(%)"]), "計画": "—", "差額": "—", "計画比": "—"},
        {"項目": "販管費率", "実績": percent(latest["販管費率(%)"]), "計画": "—", "差額": "—", "計画比": "—"},
        {"項目": "営業利益率", "実績": percent(latest["営業利益率(%)"]), "計画": "—", "差額": "—", "計画比": "—"},
    ])

def render_mapping_editor(df, title, mapping, details, key_prefix):
    st.markdown(f"#### {title}")
    cols = df.columns.tolist()

    month_default = mapping["month"] if mapping["month"] in cols else cols[0]
    mapping["month"] = st.selectbox(
        "月列",
        options=cols,
        index=cols.index(month_default),
        key=f"{key_prefix}_month"
    )

    for metric_key, label in [
        ("sales", "売上"),
        ("cost", "原価"),
        ("gross", "粗利"),
        ("sga", "販管費"),
        ("op", "営業利益"),
    ]:
        info = details[metric_key]
        best = mapping[metric_key] if mapping[metric_key] in cols else cols[0]
        conf = info["confidence_label"]
        conf_badge = "🟢" if conf == "高" else "🟡" if conf == "中" else "🔴"

        use_metric = st.checkbox(
            f"{label}を使う {conf_badge}",
            value=(mapping[metric_key] != "使わない"),
            key=f"{key_prefix}_{metric_key}_use"
        )
        if use_metric:
            selected = st.selectbox(
                f"{label}列",
                options=cols,
                index=cols.index(best),
                key=f"{key_prefix}_{metric_key}_select"
            )
            mapping[metric_key] = selected
            mapping["flows"][metric_key] = st.selectbox(
                f"{label}の種類",
                options=["単月", "累計"],
                index=0 if mapping["flows"].get(metric_key, "単月") == "単月" else 1,
                key=f"{key_prefix}_{metric_key}_flow"
            )
            with st.expander(f"{label}候補を見る", expanded=False):
                st.dataframe(pd.DataFrame(info["ranked"]), use_container_width=True, hide_index=True)
        else:
            mapping[metric_key] = "使わない"

    return mapping

# =========================================================
# Plot Helpers
# =========================================================
def base_layout(fig, height=360):
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=10, t=48, b=20),
        font=dict(size=12, color="#111827"),
        hovermode="x unified",
        legend_title_text="",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    fig.update_xaxes(showgrid=False, linecolor="rgba(0,0,0,0.08)")
    fig.update_yaxes(gridcolor="rgba(0,0,0,0.08)", zeroline=False, tickformat=",")
    return fig

def fig_plan_actual(merged, plan_col, actual_col, title):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=merged["月"],
        y=merged[plan_col],
        name="計画",
        opacity=0.30,
        hovertemplate="%{x}<br>計画 %{y:,.0f}円<extra></extra>"
    ))
    fig.add_trace(go.Scatter(
        x=merged["月"],
        y=merged[actual_col],
        name="実績",
        mode="lines+markers",
        line=dict(width=3),
        hovertemplate="%{x}<br>実績 %{y:,.0f}円<extra></extra>"
    ))
    fig.update_layout(title=title)
    return base_layout(fig, 360)

def fig_variance(merged, diff_col, title):
    colors = ["#16a34a" if (not pd.isna(v) and v >= 0) else "#dc2626" for v in merged[diff_col]]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=merged["月"],
        y=merged[diff_col],
        marker_color=colors,
        text=[f"{v:,.0f}" if not pd.isna(v) else "" for v in merged[diff_col]],
        textposition="outside",
        cliponaxis=False,
        hovertemplate="%{x}<br>%{y:,.0f}円<extra></extra>"
    ))
    fig.add_hline(y=0, line_width=1, line_color="rgba(0,0,0,0.25)")
    fig.update_layout(title=title)
    return base_layout(fig, 340)

def fig_rate_trend(merged):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=merged["月"], y=merged["粗利率(%)"],
        mode="lines+markers", name="粗利率",
        line=dict(width=3),
        hovertemplate="%{x}<br>粗利率 %{y:.1f}%<extra></extra>"
    ))
    fig.add_trace(go.Scatter(
        x=merged["月"], y=merged["販管費率(%)"],
        mode="lines+markers", name="販管費率",
        line=dict(width=3),
        hovertemplate="%{x}<br>販管費率 %{y:.1f}%<extra></extra>"
    ))
    fig.add_trace(go.Scatter(
        x=merged["月"], y=merged["営業利益率(%)"],
        mode="lines+markers", name="営業利益率",
        line=dict(width=3),
        hovertemplate="%{x}<br>営業利益率 %{y:.1f}%<extra></extra>"
    ))
    fig.update_layout(title="率の推移")
    fig.update_yaxes(ticksuffix="%")
    return base_layout(fig, 390)

def fig_waterfall_latest(latest):
    sales = safe_float(latest["売上_実績"], 0)
    cost = safe_float(latest["原価_実績"], 0)
    sga = safe_float(latest["販管費_実績"], 0)
    gross = safe_float(latest["粗利_実績"], 0)
    op = safe_float(latest["営業利益_実績"], 0)

    fig = go.Figure(go.Waterfall(
        name="今月構造",
        orientation="v",
        measure=["relative", "relative", "total", "relative", "total"],
        x=["売上", "原価", "粗利", "販管費", "営業利益"],
        y=[sales, -cost, 0, -sga, 0],
        text=[yen(sales), yen(-cost), yen(gross), yen(-sga), yen(op)],
        textposition="outside"
    ))
    fig.update_layout(title="今月の損益構造")
    return base_layout(fig, 380)

# =========================================================
# Header
# =========================================================
st.markdown("""
<div class="hero">
    <div class="hero-title">経営会議ナビ</div>
    <div class="hero-sub">
        freee実績 × bixid計画を自動突合し、今月の経営判断ポイントを整理します。
        この版は、見た目よりも「数字を間違えないこと」を優先した安全設計です。
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.markdown("## データ設定")
    freee_file = st.file_uploader("freee 実績CSV", type=["csv"], key="freee_csv")
    bixid_file = st.file_uploader("bixid 計画CSV", type=["csv"], key="bixid_csv")

    st.markdown("---")
    show_mapping = st.checkbox("列マッピングを手動調整", value=False)
    show_logs = st.checkbox("自動突合ログを表示", value=False)
    show_raw = st.checkbox("正規化データを表示", value=False)

if freee_file is None or bixid_file is None:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">使い方</div>
        <div class="section-sub">左のサイドバーから freee 実績CSV と bixid 計画CSV をアップロードしてください。</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# =========================================================
# Read CSV
# =========================================================
try:
    freee_df = read_csv_flexible(freee_file)
    bixid_df = read_csv_flexible(bixid_file)
except Exception as e:
    st.error(f"CSVの読み込みに失敗しました: {e}")
    st.stop()

# =========================================================
# Auto Mapping
# =========================================================
try:
    freee_mapping, freee_details = build_auto_mapping(freee_df, "freee 実績CSV")
    bixid_mapping, bixid_details = build_auto_mapping(bixid_df, "bixid 計画CSV")
except Exception as e:
    st.error(f"自動判定に失敗しました: {e}")
    st.stop()

# =========================================================
# Optional Mapping UI
# =========================================================
if show_mapping:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">列マッピング調整</div>
        <div class="section-sub">自動判定が怪しい列だけ直してください。数字の信用を優先するため、迷う時は必ず確認してください。</div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        freee_mapping = render_mapping_editor(freee_df, "freee 実績CSV", freee_mapping, freee_details, "freee")
    with c2:
        bixid_mapping = render_mapping_editor(bixid_df, "bixid 計画CSV", bixid_mapping, bixid_details, "bixid")

# =========================================================
# Low Confidence Warning
# =========================================================
low_conf_items = []
for source_name, details in [("freee", freee_details), ("bixid", bixid_details)]:
    for k in ["month", "sales", "cost", "gross", "sga", "op"]:
        if details[k]["confidence_label"] == "低":
            low_conf_items.append(f"{source_name}:{k}")

if low_conf_items:
    st.markdown(
        f'<div class="note">自動判定の信頼度が低い項目があります: {", ".join(low_conf_items)}。数字を確定する前に列マッピングを確認してください。</div>',
        unsafe_allow_html=True
    )

# =========================================================
# Build Monthly
# =========================================================
try:
    freee_monthly = build_freee_actual_monthly(freee_df, freee_mapping)
    bixid_monthly = build_bixid_plan_monthly(bixid_df, bixid_mapping)
except Exception as e:
    st.error(f"月次データ作成中にエラーが出ました: {e}")
    st.stop()

if freee_monthly.empty or bixid_monthly.empty:
    st.error("月次データを作成できませんでした。月列や主要列の設定を確認してください。")
    st.stop()

# =========================================================
# Merge
# =========================================================
merged = pd.merge(freee_monthly, bixid_monthly, on=["月period", "月"], how="outer")
merged = merged.sort_values("月period").reset_index(drop=True)

# KPI計算
merged["売上計画比(%)"] = merged.apply(lambda r: plan_rate(r["売上_実績"], r["売上_計画"]), axis=1)
merged["営業利益計画比(%)"] = merged.apply(lambda r: plan_rate(r["営業利益_実績"], r["営業利益_計画"]), axis=1)
merged["粗利率(%)"] = np.where((merged["売上_実績"].notna()) & (merged["売上_実績"] != 0), merged["粗利_実績"] / merged["売上_実績"] * 100, np.nan)
merged["営業利益率(%)"] = np.where((merged["売上_実績"].notna()) & (merged["売上_実績"] != 0), merged["営業利益_実績"] / merged["売上_実績"] * 100, np.nan)
merged["販管費率(%)"] = np.where((merged["売上_実績"].notna()) & (merged["売上_実績"] != 0), merged["販管費_実績"] / merged["売上_実績"] * 100, np.nan)

merged["売上差額"] = merged["売上_実績"] - merged["売上_計画"]
merged["粗利差額"] = merged["粗利_実績"] - merged["粗利_計画"]
merged["営業利益差額"] = merged["営業利益_実績"] - merged["営業利益_計画"]

latest = merged.iloc[-1]
prev = merged.iloc[-2] if len(merged) >= 2 else None

headline, headline_sub, headline_badge = build_headline(latest)
priority_topics = build_priority_topics(latest, prev)
cross_analyses = build_cross_analysis(latest, prev)
severe_anomaly, anomaly_messages = anomaly_check(latest)

sales_badge_class, sales_badge_text = diff_badge(latest["売上差額"])
gross_badge_class, gross_badge_text = diff_badge(latest["粗利差額"])
profit_badge_class, profit_badge_text = diff_badge(latest["営業利益差額"])

unit_warning = detect_unit_mismatch(merged["売上_実績"], merged["売上_計画"])

merged_info = {
    "gross_actual_source": latest.get("粗利_実績_source", "—"),
    "op_actual_source": latest.get("営業利益_実績_source", "—"),
}

# =========================================================
# Summary
# =========================================================
st.markdown(f"""
<div class="section-card">
    <div class="badge {headline_badge}">MONTHLY DECISION</div>
    <div class="section-title">{headline}</div>
    <div class="section-sub">{headline_sub}</div>
</div>
""", unsafe_allow_html=True)

if unit_warning:
    st.markdown(f'<div class="warn">{unit_warning}</div>', unsafe_allow_html=True)

if anomaly_messages:
    for msg in anomaly_messages:
        st.markdown(f'<div class="warn">{msg}</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="ok">重大な異常値は検知されていません。</div>', unsafe_allow_html=True)

if latest.get("粗利_実績_note"):
    st.markdown(f'<div class="note">粗利採用ロジック: {latest["粗利_実績_note"]}</div>', unsafe_allow_html=True)

if latest.get("営業利益_実績_note"):
    st.markdown(f'<div class="note">営業利益採用ロジック: {latest["営業利益_実績_note"]}</div>', unsafe_allow_html=True)

# =========================================================
# Stop on Severe Anomaly
# =========================================================
if severe_anomaly:
    st.error("重大な整合性エラーがあるため、KPIとグラフの表示を停止しました。列マッピングまたはCSV内容を確認してください。")

    if show_logs:
        with st.expander("自動突合ログ", expanded=True):
            st.write("freee 実績CSV")
            st.write({
                "月": freee_mapping["month"],
                "売上": freee_mapping["sales"],
                "原価": freee_mapping["cost"],
                "粗利": freee_mapping["gross"],
                "販管費": freee_mapping["sga"],
                "営業利益": freee_mapping["op"],
            })

            st.write("bixid 計画CSV")
            st.write({
                "月": bixid_mapping["month"],
                "売上": f"{bixid_mapping['sales']} / {bixid_mapping['flows']['sales']}",
                "原価": f"{bixid_mapping['cost']} / {bixid_mapping['flows']['cost']}",
                "粗利": f"{bixid_mapping['gross']} / {bixid_mapping['flows']['gross']}",
                "販管費": f"{bixid_mapping['sga']} / {bixid_mapping['flows']['sga']}",
                "営業利益": f"{bixid_mapping['op']} / {bixid_mapping['flows']['op']}",
            })

    if show_raw:
        with st.expander("正規化データ", expanded=True):
            st.write("freee 月次")
            st.dataframe(freee_monthly, use_container_width=True, hide_index=True)
            st.write("bixid 月次")
            st.dataframe(bixid_monthly, use_container_width=True, hide_index=True)
            st.write("統合後")
            st.dataframe(merged, use_container_width=True, hide_index=True)

    st.stop()

# =========================================================
# KPI
# =========================================================
k1, k2, k3, k4 = st.columns(4)
with k1:
    render_kpi(
        "売上",
        yen(latest["売上_実績"]),
        f"計画 {yen(latest['売上_計画'])}<br>差額 {yen(latest['売上差額'])}<br>計画比 {percent(latest['売上計画比(%)'])}",
        sales_badge_class,
        sales_badge_text
    )
with k2:
    render_kpi(
        "粗利",
        yen(latest["粗利_実績"]),
        f"計画 {yen(latest['粗利_計画'])}<br>差額 {yen(latest['粗利差額'])}<br>粗利率 {percent(latest['粗利率(%)'])}<br>出所 {latest.get('粗利_実績_source', '—')}",
        gross_badge_class,
        gross_badge_text
    )
with k3:
    render_kpi(
        "営業利益",
        yen(latest["営業利益_実績"]),
        f"計画 {yen(latest['営業利益_計画'])}<br>差額 {yen(latest['営業利益差額'])}<br>計画比 {percent(latest['営業利益計画比(%)'])}<br>出所 {latest.get('営業利益_実績_source', '—')}",
        profit_badge_class,
        profit_badge_text
    )
with k4:
    render_kpi(
        "販管費",
        yen(latest["販管費_実績"]),
        f"販管費率 {percent(latest['販管費率(%)'])}<br>{delta_label(latest['販管費_実績'], prev['販管費_実績']) if prev is not None else '前月比較なし'}"
    )

# =========================================================
# Priority + Summary Table
# =========================================================
p1, p2 = st.columns([1.15, 0.85])

with p1:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">会議で先に決めること</div>
        <div class="section-sub">数字の優先順位から、最初に話すべき論点を3つに整理</div>
    </div>
    """, unsafe_allow_html=True)

    tone_map = {1: "badge-red", 2: "badge-yellow", 3: "badge-blue"}

    for i, topic in enumerate(priority_topics, start=1):
        st.markdown(f"""
        <div class="priority-card">
            <div class="badge {tone_map[i]}">Priority {i}</div>
            <div class="priority-title">{topic['title']}</div>
            <div class="priority-text"><strong>なぜ今これか：</strong> {topic['why']}</div>
            <div class="priority-text"><strong>確認すること：</strong> {topic['focus']}</div>
            <div class="priority-text"><strong>決めること：</strong> {topic['decision']}</div>
        </div>
        """, unsafe_allow_html=True)

with p2:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">今月の数字一覧</div>
        <div class="section-sub">会議の土台になる主要数値と出所</div>
    </div>
    """, unsafe_allow_html=True)
    st.dataframe(summary_table(latest, merged_info), use_container_width=True, hide_index=True)

# =========================================================
# Cross Analysis
# =========================================================
st.markdown("""
<div class="section-card">
    <div class="section-title">数字の読み解き</div>
    <div class="section-sub">単体ではなく、数字どうしの関係から今月の状態を把握</div>
</div>
""", unsafe_allow_html=True)

for item in cross_analyses:
    st.markdown(f"""
    <div class="analysis-box">
        <div class="analysis-title">{item['title']}</div>
        <div style="font-size:14px; line-height:1.7; color:#6b7280;">{item['text']}</div>
    </div>
    """, unsafe_allow_html=True)

# =========================================================
# Charts
# =========================================================
st.markdown("""
<div class="section-card">
    <div class="section-title">推移を見る</div>
    <div class="section-sub">会計構造が分かるグラフだけに絞っています</div>
</div>
""", unsafe_allow_html=True)

tabs = st.tabs(["売上", "利益", "率", "今月構造"])

with tabs[0]:
    a, b = st.columns(2)
    with a:
        st.plotly_chart(
            fig_plan_actual(merged, "売上_計画", "売上_実績", "売上の推移（計画 vs 実績）"),
            use_container_width=True
        )
    with b:
        st.plotly_chart(
            fig_variance(merged, "売上差額", "売上差額の推移"),
            use_container_width=True
        )

with tabs[1]:
    a, b = st.columns(2)
    with a:
        st.plotly_chart(
            fig_plan_actual(merged, "営業利益_計画", "営業利益_実績", "営業利益の推移（計画 vs 実績）"),
            use_container_width=True
        )
    with b:
        profit_long = merged[["月", "粗利_実績", "販管費_実績", "営業利益_実績"]].melt(
            id_vars="月", var_name="区分", value_name="金額"
        )
        fig = px.line(
            profit_long,
            x="月",
            y="金額",
            color="区分",
            markers=True,
            title="粗利・販管費・営業利益の推移"
        )
        fig.update_traces(line=dict(width=3), hovertemplate="%{fullData.name}<br>%{x}<br>%{y:,.0f}円<extra></extra>")
        st.plotly_chart(base_layout(fig, 360), use_container_width=True)

with tabs[2]:
    st.plotly_chart(fig_rate_trend(merged), use_container_width=True)

with tabs[3]:
    st.plotly_chart(fig_waterfall_latest(latest), use_container_width=True)

# =========================================================
# Meeting Memo
# =========================================================
st.markdown("""
<div class="section-card">
    <div class="section-title">会議メモ</div>
    <div class="section-sub">外部要因や確認事項をその場で残せます</div>
</div>
""", unsafe_allow_html=True)

m1, m2, m3 = st.columns(3)
with m1:
    external_event = st.text_input("外部で起きていること", placeholder="例：原材料価格が上昇")
with m2:
    external_impact = st.text_input("自社への影響", placeholder="例：粗利率低下の可能性")
with m3:
    external_action = st.text_input("会議で確認すること", placeholder="例：値上げ余地、仕入先見直し")

if external_event or external_impact or external_action:
    st.markdown('<div class="info-box">', unsafe_allow_html=True)
    if external_event:
        st.write(f"**外部で起きていること**：{external_event}")
    if external_impact:
        st.write(f"**自社への影響**：{external_impact}")
    if external_action:
        st.write(f"**会議で確認すること**：{external_action}")
    st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# Debug / Logs
# =========================================================
if show_logs:
    with st.expander("自動突合ログ", expanded=False):
        st.write("freee 実績CSV")
        st.write({
            "月": freee_mapping["month"],
            "売上": freee_mapping["sales"],
            "原価": freee_mapping["cost"],
            "粗利": freee_mapping["gross"],
            "販管費": freee_mapping["sga"],
            "営業利益": freee_mapping["op"],
        })

        st.write("bixid 計画CSV")
        st.write({
            "月": bixid_mapping["month"],
            "売上": f"{bixid_mapping['sales']} / {bixid_mapping['flows']['sales']}",
            "原価": f"{bixid_mapping['cost']} / {bixid_mapping['flows']['cost']}",
            "粗利": f"{bixid_mapping['gross']} / {bixid_mapping['flows']['gross']}",
            "販管費": f"{bixid_mapping['sga']} / {bixid_mapping['flows']['sga']}",
            "営業利益": f"{bixid_mapping['op']} / {bixid_mapping['flows']['op']}",
        })

        st.write("実績採用根拠")
        st.write({
            "粗利": latest.get("粗利_実績_note", "—"),
            "営業利益": latest.get("営業利益_実績_note", "—")
        })

if show_raw:
    with st.expander("正規化データ", expanded=False):
        st.write("freee 月次")
        st.dataframe(freee_monthly, use_container_width=True, hide_index=True)
        st.write("bixid 月次")
        st.dataframe(bixid_monthly, use_container_width=True, hide_index=True)
        st.write("統合後")
        st.dataframe(merged, use_container_width=True, hide_index=True)
