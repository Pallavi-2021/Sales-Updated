"""
Sales Intelligence Hub v3.0
Enhancements:
  (a) Time-based filters: Year / Quarter / Month derived from Date
  (b) Opportunity & Issue analysis with RBAC (Management / Regional Manager / Salesperson)
  (c) Opportunity Scope Assessment with visualisations (matplotlib inline)
  (d) Downloadable reports: CSV, Excel (.xlsx), PDF summary
"""

import io, math, warnings, textwrap
from itertools import combinations
from collections import defaultdict

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.backends.backend_pdf import PdfPages

warnings.filterwarnings("ignore")

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sales Intelligence Hub",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.stApp{background:#080b12;color:#dde3ef;}
[data-testid="stSidebar"]{background:#0c0f18!important;border-right:1px solid #161d2e;}
.hero{background:linear-gradient(120deg,#0a0f1e 0%,#0f1a30 60%,#080b12 100%);
  border:1px solid #1a2540;border-radius:14px;padding:28px 34px 22px;margin-bottom:20px;
  position:relative;overflow:hidden;}
.hero-title{font-family:'JetBrains Mono',monospace;font-size:1.65rem;font-weight:700;
  color:#63b3ed;letter-spacing:-.3px;margin:0 0 4px;}
.hero-sub{font-size:.87rem;color:#4a5568;margin:0;}
.mstrip{display:flex;gap:12px;margin-bottom:18px;flex-wrap:wrap;}
.mcard{flex:1;min-width:115px;background:#0c0f18;border:1px solid #161d2e;
  border-radius:10px;padding:13px 17px;}
.mlabel{font-family:'JetBrains Mono',monospace;font-size:.6rem;color:#3a4460;
  text-transform:uppercase;letter-spacing:1.4px;margin-bottom:4px;}
.mval{font-family:'JetBrains Mono',monospace;font-size:1.45rem;font-weight:700;color:#e2e8f0;}
.c-blue{color:#63b3ed;}.c-red{color:#fc8181;}.c-grn{color:#68d391;}
.sh{font-family:'JetBrains Mono',monospace;font-size:.64rem;color:#63b3ed;
  text-transform:uppercase;letter-spacing:2.5px;border-bottom:1px solid #161d2e;
  padding-bottom:6px;margin:20px 0 11px;}
.ibox{background:rgba(99,179,237,.06);border-left:3px solid #63b3ed;border-radius:6px;
  padding:9px 14px;font-size:.82rem;color:#718096;margin-bottom:11px;}
.wbox{background:rgba(252,129,74,.06);border-left:3px solid #fc814a;border-radius:6px;
  padding:9px 14px;font-size:.82rem;color:#718096;margin-bottom:11px;}
.gbox{background:rgba(104,211,145,.06);border-left:3px solid #68d391;border-radius:6px;
  padding:9px 14px;font-size:.82rem;color:#718096;margin-bottom:11px;}
.badge{display:inline-block;padding:3px 11px;border-radius:20px;
  font-family:'JetBrains Mono',monospace;font-size:.64rem;font-weight:600;}
.bc{background:rgba(252,129,74,.12);color:#fc8181;border:1px solid rgba(252,129,74,.25);}
.bh{background:rgba(246,173,85,.12);color:#f6ad55;border:1px solid rgba(246,173,85,.25);}
.bg{background:rgba(104,211,145,.12);color:#68d391;border:1px solid rgba(104,211,145,.25);}
.stButton>button{background:#2b6cb0!important;color:#e2e8f0!important;
  font-family:'JetBrains Mono',monospace!important;font-weight:600!important;
  border:none!important;border-radius:8px!important;padding:9px 26px!important;}
.stButton>button:hover{background:#3182ce!important;}
.lc{background:#0c0f18;border:1px solid #161d2e;border-radius:10px;padding:15px 19px;margin-bottom:9px;}
.lt{font-family:'JetBrains Mono',monospace;font-size:.7rem;color:#63b3ed;letter-spacing:1px;
  margin-bottom:6px;font-weight:600;}
.lb{font-size:.81rem;color:#718096;line-height:1.75;}
.pinfo{font-family:'JetBrains Mono',monospace;font-size:.68rem;color:#4a5568;
  text-align:center;margin:5px 0;}
.rbac-tag{display:inline-block;padding:4px 14px;border-radius:20px;font-size:.7rem;
  font-family:'JetBrains Mono',monospace;font-weight:700;margin-bottom:8px;}
.role-mgmt{background:rgba(99,179,237,.15);color:#63b3ed;border:1px solid #63b3ed55;}
.role-rm{background:rgba(246,173,85,.15);color:#f6ad55;border:1px solid #f6ad5555;}
.role-sp{background:rgba(104,211,145,.15);color:#68d391;border:1px solid #68d39155;}
</style>
""", unsafe_allow_html=True)

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
CRITICAL_DECLINE_PCT    = -50.0
MIN_PREV_PERIOD_SALES   = 500
MIN_INVOICES_FOR_PERIOD = 2
MIN_STORE_INVOICES      = 5
MIN_MISSING_PROD_STORES = 3
MIN_MISSING_AVG_SALES   = 1000
MIN_SUPPORT_COUNT       = 3
MIN_CONFIDENCE_PCT      = 40.0
PAGE_SIZE               = 200

# Updated expected columns for new dataset structure
EXPECTED_COLS = [
    "StoreID", "Store_Name", "Salesperson_Code", "Salesperson_Name",
    "Date", "Region", "Regional_Manager",
    "Store CheckIn Time", "Store Check Out Time",
    "Avg_Call_Duration", "StoreLatitude", "StoreLongitude",
    "Visit Frequency", "InvoiceNo", "Product", "Qty", "Rate",
    "LineTotal", "InvoiceTotal",
]

DISPLAY_COLS = [
    "Store_ID", "Store_Name", "Region", "Regional_Manager",
    "Salesperson_Name", "Product",
    "Reason_Type", "Period_Type", "Observation_Type",
    "Prev_Period", "Curr_Period",
    "Prev_Sales", "Curr_Sales", "Pct_Change", "Impact", "Analysis_Type",
]
OUTPUT_COLS = DISPLAY_COLS + [
    "Reason_Description", "Priority",
    "Store_Latitude", "Store_Longitude",
    "Salesperson_Code",
]

PERIOD_TYPE_MAP = {"Mth": "Monthly", "Qtr": "Quarterly", "Yr": "Yearly"}
OBS_TYPE_MAP = {
    "Full Sales Drop":            "Issue",
    "Critical Decline":           "Issue",
    "Missing High-Value Product": "Opportunity",
    "Bundle Opportunity":         "Opportunity",
    "Low Visit Frequency":        "Opportunity",
    "High Potential Region":      "Opportunity",
    "Under-Serviced Store":       "Opportunity",
    "Salesperson Productivity":   "Opportunity",
}

# ── ROLE DEFINITIONS ─────────────────────────────────────────────────────────
ROLE_MANAGEMENT = "Management"
ROLE_REGIONAL   = "Regional Manager"
ROLE_SALES      = "Salesperson"
ROLES           = [ROLE_MANAGEMENT, ROLE_REGIONAL, ROLE_SALES]

# ── DATA LOAD ─────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, fname: str):
    buf = io.BytesIO(file_bytes)
    df  = pd.read_csv(buf) if fname.lower().endswith(".csv") else pd.read_excel(buf)
    df.columns = [c.strip() for c in df.columns]
    missing    = [c for c in EXPECTED_COLS if c not in df.columns]

    df["Date"]        = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
    df["LineTotal"]   = pd.to_numeric(df.get("LineTotal",   0), errors="coerce").fillna(0)
    df["InvoiceTotal"]= pd.to_numeric(df.get("InvoiceTotal",0), errors="coerce").fillna(0)
    df["Qty"]         = pd.to_numeric(df.get("Qty",         0), errors="coerce").fillna(0)
    df["Rate"]        = pd.to_numeric(df.get("Rate",        0), errors="coerce").fillna(0)

    # Normalise Visit Frequency — strip non-numeric chars (e.g. "3x", "2/week")
    if "Visit Frequency" in df.columns:
        df["Visit Frequency"] = pd.to_numeric(
            df["Visit Frequency"].astype(str).str.extract(r"([\d.]+)", expand=False),
            errors="coerce",
        ).fillna(0)

    # Normalise Avg_Call_Duration — support HH:MM:SS / HH:MM strings and plain numbers
    if "Avg_Call_Duration" in df.columns:
        def _parse_duration(v):
            if pd.isna(v):
                return np.nan
            s = str(v).strip()
            if ":" in s:
                parts = s.split(":")
                try:
                    h   = float(parts[0])
                    m   = float(parts[1])
                    sec = float(parts[2]) if len(parts) > 2 else 0.0
                    return h * 60 + m + sec / 60
                except (ValueError, IndexError):
                    return np.nan
            try:
                return float(s)
            except ValueError:
                return np.nan
        df["Avg_Call_Duration"] = df["Avg_Call_Duration"].apply(_parse_duration).fillna(0)

    df.dropna(subset=["Date", "StoreID", "Product"], inplace=True)

    # Derived time columns
    df["_year"]    = df["Date"].dt.year
    df["_month"]   = df["Date"].dt.to_period("M")
    df["_quarter"] = df["Date"].dt.to_period("Q")
    df["_month_num"] = df["Date"].dt.month

    for col in ["LineTotal", "Qty", "Rate", "InvoiceTotal"]:
        if col in df.columns:
            df[col] = df[col].astype("float32")

    return df, missing


# ── TIME FILTER HELPER ────────────────────────────────────────────────────────
def apply_time_filters(df: pd.DataFrame,
                       sel_years, sel_quarters, sel_months) -> pd.DataFrame:
    """Filter df by selected year/quarter/month values (dynamic, data-driven)."""
    if sel_years:
        df = df[df["_year"].isin(sel_years)]
    if sel_quarters:
        # sel_quarters are strings like "2024Q1"
        df = df[df["_quarter"].astype(str).isin(sel_quarters)]
    if sel_months:
        # sel_months are ints 1–12
        df = df[df["_month_num"].isin(sel_months)]
    return df


# ── RBAC FILTER ───────────────────────────────────────────────────────────────
def apply_rbac(df: pd.DataFrame,
               role: str,
               identity: str) -> pd.DataFrame:
    """Return only the rows the current role/identity is allowed to see."""
    if role == ROLE_MANAGEMENT:
        return df
    elif role == ROLE_REGIONAL:
        if "Regional_Manager" in df.columns:
            return df[df["Regional_Manager"] == identity]
        elif "Region" in df.columns:
            return df[df["Region"] == identity]
        return df
    elif role == ROLE_SALES:
        for col in ["Salesperson_Name", "Salesperson_Code"]:
            if col in df.columns:
                return df[df[col] == identity]
    return df


# ── PERIOD ENGINE ─────────────────────────────────────────────────────────────
def _period_compare(df, period_col, cur_lbl, prev_lbl, pname):
    group_keys = ["StoreID", "Store_Name", "Region", "Regional_Manager",
                  "Salesperson_Name", "Salesperson_Code",
                  "StoreLatitude", "StoreLongitude", "Product"]
    group_keys = [k for k in group_keys if k in df.columns]

    grp = (df.groupby(group_keys + [period_col], as_index=False)
             .agg(Sales=("LineTotal", "sum"), Inv=("InvoiceNo", "nunique")))
    cur  = grp[grp[period_col] == cur_lbl]
    prev = grp[grp[period_col] == prev_lbl]
    if cur.empty or prev.empty:
        return pd.DataFrame()

    m = prev.merge(cur, on=group_keys, suffixes=("_p", "_c"), how="outer").fillna(0)
    rows = []
    for _, r in m.iterrows():
        ps  = float(r["Sales_p"]); cs = float(r["Sales_c"])
        inv = max(r.get("Inv_p", 0), r.get("Inv_c", 0))
        if ps < MIN_PREV_PERIOD_SALES or inv < MIN_INVOICES_FOR_PERIOD:
            continue
        pct      = ((cs - ps) / ps * 100) if ps > 0 else 0.0
        zero_drop   = (ps > 0 and cs == 0)
        is_critical = zero_drop or pct <= CRITICAL_DECLINE_PCT
        if not is_critical:
            continue
        impact = round(abs(cs - ps), 0)
        rtype  = "Full Sales Drop" if zero_drop else "Critical Decline"
        desc   = (f"Lost ALL sales in {cur_lbl} (prev: {ps:,.0f})." if zero_drop
                  else f"Fell {abs(pct):.1f}%: {ps:,.0f}→{cs:,.0f}. At-risk: {impact:,.0f}.")
        rows.append({
            "Store_ID":        r["StoreID"],
            "Store_Name":      r.get("Store_Name", ""),
            "Region":          r.get("Region", ""),
            "Regional_Manager":r.get("Regional_Manager", ""),
            "Salesperson_Name":r.get("Salesperson_Name", ""),
            "Salesperson_Code":r.get("Salesperson_Code", ""),
            "Product":         r["Product"],
            "Reason_Type":     rtype,
            "Reason_Description": desc,
            "Prev_Period":     str(prev_lbl),
            "Curr_Period":     str(cur_lbl),
            "Prev_Sales":      round(ps, 0),
            "Curr_Sales":      round(cs, 0),
            "Pct_Change":      round(pct, 1),
            "Impact":          impact,
            "Priority":        "Critical",
            "Analysis_Type":   f"Period ({pname})",
            "Period_Type":     PERIOD_TYPE_MAP.get(pname, pname),
            "Observation_Type":OBS_TYPE_MAP.get(rtype, "Issue"),
            "Store_Latitude":  r.get("StoreLatitude", None),
            "Store_Longitude": r.get("StoreLongitude", None),
        })
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def run_all_period_analyses(df_bytes: bytes):
    df = pd.read_parquet(io.BytesIO(df_bytes))
    frames, summary = [], {}
    months = sorted(df["_month"].unique())
    if len(months) >= 2:
        cm, pm = months[-1], months[-2]
        d = _period_compare(df, "_month", cm, pm, "Mth")
        summary["Monthly"] = {"current": str(cm), "previous": str(pm), "critical_rows": len(d)}
        if not d.empty: frames.append(d)
    else:
        summary["Monthly"] = {"note": "Need ≥ 2 months"}
    quarters = sorted(df["_quarter"].unique())
    if len(quarters) >= 2:
        cq, pq = quarters[-1], quarters[-2]
        d = _period_compare(df, "_quarter", cq, pq, "Qtr")
        summary["Quarterly"] = {"current": str(cq), "previous": str(pq), "critical_rows": len(d)}
        if not d.empty: frames.append(d)
    else:
        summary["Quarterly"] = {"note": "Need ≥ 2 quarters"}
    years = sorted(df["_year"].unique())
    if len(years) >= 2:
        cy, py = years[-1], years[-2]
        d = _period_compare(df, "_year", cy, py, "Yr")
        summary["Yearly"] = {"current": str(cy), "previous": str(py), "critical_rows": len(d)}
        if not d.empty: frames.append(d)
    else:
        summary["Yearly"] = {"note": "Need ≥ 2 years"}
    if not frames:
        return pd.DataFrame(), summary
    out = pd.concat(frames, ignore_index=True)
    out["_s"] = out["Reason_Type"].map({"Full Sales Drop": 0, "Critical Decline": 1}).fillna(2)
    out = (out.sort_values("_s")
              .drop_duplicates(subset=["Store_ID", "Product", "Analysis_Type"])
              .drop(columns=["_s"])
              .sort_values("Impact", ascending=False)
              .reset_index(drop=True))
    return out, summary


# ── CROSS-SELL ENGINE ─────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def run_cross_sell_analysis(df_bytes: bytes):
    df = pd.read_parquet(io.BytesIO(df_bytes))
    meta_cols = ["StoreID", "Store_Name", "Region", "Regional_Manager",
                 "Salesperson_Name", "Salesperson_Code", "StoreLatitude", "StoreLongitude"]
    meta_cols = [c for c in meta_cols if c in df.columns]
    store_meta = (df.drop_duplicates("StoreID")[meta_cols].set_index("StoreID").to_dict("index"))
    prod_stats  = df.groupby("Product").agg(gavg=("LineTotal", "mean"),
                                             nstores=("StoreID", "nunique"))
    baskets     = df.groupby(["StoreID", "InvoiceNo"])["Product"].apply(list).reset_index()
    sic         = baskets.groupby("StoreID")["InvoiceNo"].nunique().to_dict()
    spro        = df.groupby("StoreID")["Product"].apply(set).to_dict()
    latest_month    = df["_month"].max()
    active_products = set(df[df["_month"] == latest_month]["Product"].unique())
    records = []

    for sid, meta in store_meta.items():
        ic = sic.get(sid, 0)
        if ic < MIN_STORE_INVOICES: continue
        existing = spro.get(sid, set())
        for prod, stats in prod_stats.iterrows():
            if prod in existing: continue
            if prod not in active_products: continue
            if stats["nstores"] < MIN_MISSING_PROD_STORES: continue
            if stats["gavg"] < MIN_MISSING_AVG_SALES: continue
            impact = round(stats["gavg"] * (ic / 4), 0)
            records.append({
                "Store_ID": sid, "Store_Name": meta.get("Store_Name", ""),
                "Region": meta.get("Region", ""),
                "Regional_Manager": meta.get("Regional_Manager", ""),
                "Salesperson_Name": meta.get("Salesperson_Name", ""),
                "Salesperson_Code": meta.get("Salesperson_Code", ""),
                "Product": prod,
                "Reason_Type": "Missing High-Value Product",
                "Reason_Description": (f"'{prod}' absent; active in {int(stats['nstores'])} "
                                       f"stores avg {stats['gavg']:,.0f}/inv. Uplift ~{impact:,.0f}."),
                "Prev_Period": "N/A", "Curr_Period": str(latest_month),
                "Prev_Sales": 0, "Curr_Sales": 0, "Pct_Change": 0,
                "Impact": impact, "Priority": "High",
                "Analysis_Type": "Cross Selling", "Period_Type": "N/A",
                "Observation_Type": "Opportunity",
                "Store_Latitude": meta.get("StoreLatitude"), "Store_Longitude": meta.get("StoreLongitude"),
            })

    pair_counts = defaultdict(lambda: defaultdict(int))
    for sid, grp in baskets.groupby("StoreID"):
        if sic.get(sid, 0) < MIN_STORE_INVOICES: continue
        for _, row in grp.iterrows():
            ps = list(set(row["Product"]))
            for a, b in combinations(sorted(ps), 2):
                pair_counts[sid][(a, b)] += 1
    for sid, pairs in pair_counts.items():
        meta = store_meta.get(sid)
        if not meta: continue
        itot = sic.get(sid, 0)
        for (a, b), cnt in pairs.items():
            if cnt < MIN_SUPPORT_COUNT: continue
            conf = cnt / itot * 100
            if conf < MIN_CONFIDENCE_PCT: continue
            ga = prod_stats.loc[a, "gavg"] if a in prod_stats.index else 0
            gb = prod_stats.loc[b, "gavg"] if b in prod_stats.index else 0
            impact = round((ga + gb) * (itot - cnt) * 0.3, 0)
            if impact < MIN_MISSING_AVG_SALES: continue
            desc = (f"'{a}'+'{b}' co-bought {cnt}/{itot} ({conf:.1f}%). Bundle uplift ~{impact:,.0f}.")
            for prod in [a, b]:
                records.append({
                    "Store_ID": sid, "Store_Name": meta.get("Store_Name", ""),
                    "Region": meta.get("Region", ""),
                    "Regional_Manager": meta.get("Regional_Manager", ""),
                    "Salesperson_Name": meta.get("Salesperson_Name", ""),
                    "Salesperson_Code": meta.get("Salesperson_Code", ""),
                    "Product": prod,
                    "Reason_Type": "Bundle Opportunity", "Reason_Description": desc,
                    "Prev_Period": "N/A", "Curr_Period": str(df["_month"].max()),
                    "Prev_Sales": 0, "Curr_Sales": 0, "Pct_Change": 0,
                    "Impact": impact, "Priority": "High",
                    "Analysis_Type": "Cross Selling", "Period_Type": "N/A",
                    "Observation_Type": "Opportunity",
                    "Store_Latitude": meta.get("StoreLatitude"), "Store_Longitude": meta.get("StoreLongitude"),
                })

    if not records:
        return pd.DataFrame()
    return (pd.DataFrame(records)
              .sort_values("Impact", ascending=False)
              .drop_duplicates(subset=["Store_ID", "Product", "Reason_Type"])
              .reset_index(drop=True))


# ── OPPORTUNITY SCOPE ENGINE ──────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def run_opportunity_scope(df_bytes: bytes):
    """
    Five opportunity dimensions:
      1. Revenue growth opportunities (regions/products with high upside)
      2. Under-serviced / low-coverage stores
      3. High-potential regions / products
      4. Visit optimisation
      5. Salesperson productivity
    Returns dict of DataFrames per dimension.
    """
    df = pd.read_parquet(io.BytesIO(df_bytes))

    # ── Pre-coerce any column that must be numeric before groupby ─────────────
    # Handles time-string formats ("HH:MM:SS", "HH:MM") as well as plain numbers.
    def _to_numeric_minutes(series: pd.Series) -> pd.Series:
        """Convert a series to float (minutes).
        Accepts: numeric values, 'HH:MM:SS', 'HH:MM', or plain strings of numbers.
        Non-parseable values become NaN.
        """
        def _parse_one(v):
            if pd.isna(v):
                return np.nan
            s = str(v).strip()
            # HH:MM:SS or HH:MM
            if ":" in s:
                parts = s.split(":")
                try:
                    h = float(parts[0])
                    m = float(parts[1])
                    sec = float(parts[2]) if len(parts) > 2 else 0.0
                    return h * 60 + m + sec / 60
                except (ValueError, IndexError):
                    return np.nan
            try:
                return float(s)
            except ValueError:
                return np.nan
        return series.apply(_parse_one)

    if "Avg_Call_Duration" in df.columns:
        df["Avg_Call_Duration"] = _to_numeric_minutes(df["Avg_Call_Duration"])

    if "Visit Frequency" in df.columns:
        df["Visit Frequency"] = pd.to_numeric(
            df["Visit Frequency"].astype(str).str.extract(r"([\d.]+)", expand=False),
            errors="coerce"
        )

    results = {}

    # ── 1. Revenue growth — region × product combos below median ──────────────
    rp = (df.groupby(["Region", "Product"], as_index=False)
            .agg(Revenue=("LineTotal", "sum"), Invoices=("InvoiceNo", "nunique")))
    median_rev = rp["Revenue"].median()
    rp["Gap"]  = (median_rev - rp["Revenue"]).clip(lower=0)
    rp["Opportunity_Type"] = "Revenue Growth"
    results["revenue_growth"] = rp[rp["Gap"] > 0].sort_values("Gap", ascending=False)

    # ── 2. Under-serviced stores — invoice count below 25th percentile ────────
    us_group_keys = [k for k in ["StoreID", "Store_Name", "Region", "Regional_Manager"]
                     if k in df.columns]
    si = (df.groupby(us_group_keys, as_index=False)
            .agg(Invoices=("InvoiceNo", "nunique"), Revenue=("LineTotal", "sum")))
    p25 = si["Invoices"].quantile(0.25)
    si["Gap_Invoices"] = (p25 - si["Invoices"]).clip(lower=0)
    avg_rev_per_inv  = df["LineTotal"].sum() / max(df["InvoiceNo"].nunique(), 1)
    si["Est_Uplift"] = si["Gap_Invoices"] * avg_rev_per_inv
    results["under_serviced"] = si[si["Gap_Invoices"] > 0].sort_values("Est_Uplift", ascending=False)

    # ── 3. High-potential regions — revenue per invoice ────────────────────────
    rg = (df.groupby("Region", as_index=False)
            .agg(Revenue=("LineTotal", "sum"), Invoices=("InvoiceNo", "nunique"),
                 Stores=("StoreID", "nunique")))
    rg["Rev_per_Invoice"]  = rg["Revenue"] / rg["Invoices"].replace(0, np.nan)
    rg["Opportunity_Type"] = "High Potential Region"
    results["high_potential_regions"] = rg.sort_values("Rev_per_Invoice", ascending=False)

    # ── 4. Visit optimisation ─────────────────────────────────────────────────
    if "Visit Frequency" in df.columns and df["Visit Frequency"].notna().any():
        vf_group = [k for k in ["StoreID", "Store_Name", "Region"] if k in df.columns]
        vf = (df.groupby(vf_group, as_index=False)
                .agg(Avg_Visit_Freq=("Visit Frequency", "mean"),
                     Revenue=("LineTotal", "sum")))
        vf["Avg_Visit_Freq"] = pd.to_numeric(vf["Avg_Visit_Freq"], errors="coerce").fillna(0)
        med_freq = vf["Avg_Visit_Freq"].median()
        med_rev  = vf["Revenue"].median()
        vf["Visit_Gap"] = (med_freq - vf["Avg_Visit_Freq"]).clip(lower=0)
        vf["Priority"]  = np.where(
            (vf["Avg_Visit_Freq"] < med_freq) & (vf["Revenue"] > med_rev),
            "High", "Normal",
        )
        results["visit_optimisation"] = vf.sort_values("Visit_Gap", ascending=False)
    else:
        results["visit_optimisation"] = pd.DataFrame()

    # ── 5. Salesperson productivity ───────────────────────────────────────────
    sp_keys = [k for k in ["Salesperson_Code", "Salesperson_Name", "Region", "Regional_Manager"]
               if k in df.columns]
    if ("Avg_Call_Duration" in df.columns
            and df["Avg_Call_Duration"].notna().any()
            and sp_keys):
        sp = (df.groupby(sp_keys, as_index=False)
                .agg(Revenue=("LineTotal", "sum"),
                     Avg_Call=("Avg_Call_Duration", "mean"),
                     Stores=("StoreID", "nunique")))
        sp["Avg_Call"] = pd.to_numeric(sp["Avg_Call"], errors="coerce").fillna(0)
        sp["Rev_per_Call_Min"] = sp["Revenue"] / sp["Avg_Call"].replace(0, np.nan)
        sp["Rev_per_Call_Min"] = pd.to_numeric(sp["Rev_per_Call_Min"], errors="coerce").fillna(0)
        med_rpc = sp["Rev_per_Call_Min"].median()
        sp["Productivity_Gap"] = (med_rpc - sp["Rev_per_Call_Min"]).clip(lower=0)
        results["salesperson_productivity"] = sp.sort_values("Productivity_Gap", ascending=False)
    else:
        results["salesperson_productivity"] = pd.DataFrame()

    return results


# ── CHART HELPERS ─────────────────────────────────────────────────────────────
_DARK   = "#080b12"
_GRID   = "#161d2e"
_BLUE   = "#63b3ed"
_GREEN  = "#68d391"
_ORANGE = "#f6ad55"
_RED    = "#fc8181"
_TEXT   = "#dde3ef"
_MUTED  = "#4a5568"


def _fig(w=9, h=4):
    fig, ax = plt.subplots(figsize=(w, h))
    fig.patch.set_facecolor(_DARK)
    ax.set_facecolor(_DARK)
    ax.tick_params(colors=_MUTED, labelsize=8)
    ax.spines[:].set_color(_GRID)
    ax.xaxis.label.set_color(_MUTED)
    ax.yaxis.label.set_color(_MUTED)
    ax.title.set_color(_BLUE)
    ax.yaxis.set_tick_params(color=_GRID)
    ax.xaxis.set_tick_params(color=_GRID)
    ax.grid(axis="y", color=_GRID, linewidth=0.5)
    return fig, ax


def chart_revenue_growth(df_rg):
    top = df_rg.nlargest(12, "Gap")
    if top.empty:
        return None
    fig, ax = _fig(10, 4)
    labels = [f"{r['Region'][:12]}\n{r['Product'][:12]}" for _, r in top.iterrows()]
    bars = ax.bar(range(len(top)), top["Gap"], color=_BLUE, width=0.6, alpha=0.85)
    ax.set_xticks(range(len(top)))
    ax.set_xticklabels(labels, fontsize=6.5, rotation=30, ha="right")
    ax.set_title("Revenue Growth Gap — Top Region × Product Combos", fontsize=10, pad=10)
    ax.set_ylabel("Gap to Median (₹)", fontsize=8)
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, h + h * 0.01,
                f"{h:,.0f}", ha="center", va="bottom", fontsize=6, color=_MUTED)
    fig.tight_layout()
    return fig


def chart_under_serviced(df_us):
    top = df_us.nlargest(10, "Est_Uplift")
    if top.empty:
        return None
    fig, ax = _fig(10, 4)
    labels = [f"{r['Store_Name'][:15]}\n({r['Region'][:10]})" for _, r in top.iterrows()]
    ax.barh(range(len(top)), top["Est_Uplift"], color=_ORANGE, alpha=0.85)
    ax.set_yticks(range(len(top)))
    ax.set_yticklabels(labels, fontsize=7)
    ax.set_title("Under-Serviced Stores — Estimated Uplift if Visit Gap Closed", fontsize=10, pad=10)
    ax.set_xlabel("Estimated Revenue Uplift (₹)", fontsize=8)
    ax.invert_yaxis()
    fig.tight_layout()
    return fig


def chart_region_potential(df_rg):
    if df_rg.empty:
        return None
    fig, ax = _fig(8, 4)
    regions = df_rg.sort_values("Rev_per_Invoice", ascending=False)
    colors  = [_GREEN if i < 3 else _BLUE for i in range(len(regions))]
    ax.bar(regions["Region"], regions["Rev_per_Invoice"], color=colors, alpha=0.85)
    ax.set_title("Revenue per Invoice by Region", fontsize=10, pad=10)
    ax.set_ylabel("Rev / Invoice (₹)", fontsize=8)
    ax.set_xticklabels(regions["Region"], rotation=25, ha="right", fontsize=7)
    top_patch    = mpatches.Patch(color=_GREEN, label="Top 3 regions")
    other_patch  = mpatches.Patch(color=_BLUE,  label="Others")
    ax.legend(handles=[top_patch, other_patch], facecolor=_DARK, labelcolor=_TEXT, fontsize=7)
    fig.tight_layout()
    return fig


def chart_visit_optimisation(df_vo):
    if df_vo.empty:
        return None
    top = df_vo.nlargest(12, "Visit_Gap")
    fig, ax = _fig(10, 4)
    colors = [_RED if p == "High" else _BLUE for p in top["Priority"]]
    ax.bar(range(len(top)), top["Visit_Gap"], color=colors, alpha=0.85)
    ax.set_xticks(range(len(top)))
    labels = [f"{r['Store_Name'][:12]}" for _, r in top.iterrows()]
    ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=7)
    ax.set_title("Visit Frequency Gap by Store", fontsize=10, pad=10)
    ax.set_ylabel("Gap to Median Visit Freq", fontsize=8)
    high_p = mpatches.Patch(color=_RED,  label="High-revenue, under-visited")
    norm_p = mpatches.Patch(color=_BLUE, label="Normal priority")
    ax.legend(handles=[high_p, norm_p], facecolor=_DARK, labelcolor=_TEXT, fontsize=7)
    fig.tight_layout()
    return fig


def chart_salesperson_productivity(df_sp):
    if df_sp.empty:
        return None
    top = df_sp.nlargest(12, "Productivity_Gap")
    fig, ax = _fig(10, 4)
    ax.barh(range(len(top)), top["Productivity_Gap"], color=_RED, alpha=0.8)
    labels = [f"{r['Salesperson_Name'][:18]}" for _, r in top.iterrows()]
    ax.set_yticks(range(len(top)))
    ax.set_yticklabels(labels, fontsize=7)
    ax.set_title("Salesperson Productivity Gap (Rev/Call-Minute vs Median)", fontsize=10, pad=10)
    ax.set_xlabel("Gap (₹ per Call Minute)", fontsize=8)
    ax.invert_yaxis()
    fig.tight_layout()
    return fig


# ── EXPORT HELPERS ────────────────────────────────────────────────────────────
def to_excel_bytes(dfs_dict: dict) -> bytes:
    """Write multiple DataFrames to separate sheets in one xlsx file."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet, df in dfs_dict.items():
            if df is not None and not df.empty:
                safe_name = sheet[:31]   # Excel sheet name limit
                df.to_excel(writer, sheet_name=safe_name, index=False)
    return buf.getvalue()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def to_pdf_bytes(dfs_dict: dict, figs_dict: dict, meta: dict) -> bytes:
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        # Cover page
        fig_cover, ax_cover = plt.subplots(figsize=(11, 8.5))
        fig_cover.patch.set_facecolor("#0a0f1e")
        ax_cover.set_facecolor("#0a0f1e")
        ax_cover.axis("off")
        ax_cover.text(0.5, 0.72, "Sales Intelligence Hub", ha="center", va="center",
                      fontsize=26, color="#63b3ed", fontweight="bold",
                      fontfamily="monospace", transform=ax_cover.transAxes)
        ax_cover.text(0.5, 0.60, "Opportunity & Issue Analysis Report", ha="center",
                      fontsize=14, color="#dde3ef", transform=ax_cover.transAxes)
        for i, (k, v) in enumerate(meta.items()):
            ax_cover.text(0.5, 0.45 - i * 0.07, f"{k}: {v}", ha="center",
                          fontsize=10, color="#4a5568", transform=ax_cover.transAxes)
        pdf.savefig(fig_cover, bbox_inches="tight")
        plt.close(fig_cover)

        # Charts
        for title, fig in figs_dict.items():
            if fig is not None:
                pdf.savefig(fig, bbox_inches="tight")
                plt.close(fig)

        # Summary tables (first 50 rows each)
        for title, df in dfs_dict.items():
            if df is None or df.empty:
                continue
            subset = df.head(50)
            n_rows = len(subset)
            fig_t, ax_t = plt.subplots(figsize=(16, max(2, n_rows * 0.28 + 1)))
            fig_t.patch.set_facecolor("#0a0f1e")
            ax_t.set_facecolor("#0a0f1e")
            ax_t.axis("off")
            ax_t.set_title(title, fontsize=10, color="#63b3ed", pad=6, fontfamily="monospace")
            cols = [c[:18] for c in subset.columns.tolist()]
            data = [[str(v)[:20] for v in row] for row in subset.values]
            tbl  = ax_t.table(cellText=data, colLabels=cols,
                               cellLoc="center", loc="center")
            tbl.auto_set_font_size(False)
            tbl.set_fontsize(6)
            tbl.scale(1, 1.3)
            for (r, c), cell in tbl.get_celld().items():
                if r == 0:
                    cell.set_facecolor("#0f1a30")
                    cell.set_text_props(color="#63b3ed", fontweight="bold")
                else:
                    cell.set_facecolor("#0c0f18" if r % 2 == 0 else "#080b12")
                    cell.set_text_props(color="#dde3ef")
                cell.set_edgecolor("#161d2e")
            pdf.savefig(fig_t, bbox_inches="tight")
            plt.close(fig_t)

    return buf.getvalue()


# ── PAGINATED TABLE ────────────────────────────────────────────────────────────
def paginated_table(df_full: pd.DataFrame, key: str, height: int = 400):
    total = len(df_full)
    if total == 0:
        st.info("No rows to display.")
        return
    show_cols = [c for c in DISPLAY_COLS if c in df_full.columns]
    df_view   = df_full[show_cols]
    n_pages   = max(1, math.ceil(total / PAGE_SIZE))
    pkey      = f"_pg_{key}"
    if pkey not in st.session_state:
        st.session_state[pkey] = 1
    cl, cc, cr = st.columns([1, 4, 1])
    with cl:
        if st.button("← Prev", key=f"prev_{key}"):
            st.session_state[pkey] = max(1, st.session_state[pkey] - 1)
    with cr:
        if st.button("Next →", key=f"next_{key}"):
            st.session_state[pkey] = min(n_pages, st.session_state[pkey] + 1)
    with cc:
        if n_pages > 1:
            page = st.slider("", 1, n_pages, st.session_state[pkey],
                             key=f"sl_{key}", label_visibility="collapsed")
            st.session_state[pkey] = page
        else:
            page = 1
    s = (page - 1) * PAGE_SIZE
    e = min(s + PAGE_SIZE, total)
    st.markdown(f'<p class="pinfo">Rows {s+1}–{e} of {total:,}  |  Page {page}/{n_pages}</p>',
                unsafe_allow_html=True)
    try:
        st.dataframe(df_view.iloc[s:e], use_container_width=True, hide_index=True, height=height)
    except Exception as _df_err:
        st.warning(f"Table render error: {_df_err} — try the 🔄 Clear Cache button in the sidebar.")
        st.write(df_view.iloc[s:e].to_dict(orient="records"))


# ═══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""<div style='font-family:"JetBrains Mono",monospace;font-size:.61rem;
         color:#63b3ed;letter-spacing:2.5px;text-transform:uppercase;
         padding:10px 0 15px;border-bottom:1px solid #161d2e;'>
        ◈ SALES INTELLIGENCE HUB v3</div>""", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # Cache-clear safety valve — shown at top of sidebar so users can recover
    # from "Failed to fetch dynamically imported module" stale-asset errors
    # without needing a hard browser refresh or Streamlit Cloud redeploy.
    if st.button("🔄  Clear Cache & Reload", key="cache_clear",
                 help="Fixes 'Failed to fetch' JS errors by clearing all cached data"):
        st.cache_data.clear()
        st.rerun()

    uploaded_file = st.file_uploader("Upload Sales Data (CSV / Excel)",
                                     type=["csv", "xlsx", "xls"])
    st.markdown("---")

    # ── RBAC ──────────────────────────────────────────────────────────────────
    st.markdown('<p style="font-family:JetBrains Mono;font-size:.6rem;color:#63b3ed;'
                'letter-spacing:2px;text-transform:uppercase;">ACCESS CONTROL</p>',
                unsafe_allow_html=True)
    user_role = st.selectbox("Login as", ROLES, key="user_role")

    role_class = {"Management": "role-mgmt", "Regional Manager": "role-rm",
                  "Salesperson": "role-sp"}.get(user_role, "role-mgmt")
    st.markdown(f'<span class="rbac-tag {role_class}">{user_role}</span>',
                unsafe_allow_html=True)

    # Identity picker populated after data load
    identity_placeholder = st.empty()
    st.markdown("---")

    # ── ANALYSIS OPTIONS ─────────────────────────────────────────────────────
    run_period = st.checkbox("Period-to-Period Analysis",  value=True)
    run_cross  = st.checkbox("Cross-Selling Analysis",     value=True)
    run_scope  = st.checkbox("Opportunity Scope Analysis", value=True)
    st.markdown("---")

    st.markdown("""<div style='font-size:.63rem;color:#2d3748;
        font-family:"JetBrains Mono",monospace;'>
        v3.0 · Time Filters · RBAC · Scope Charts · Multi-format Export</div>""",
        unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  HERO
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""<div class="hero">
  <p class="hero-title">◈ Sales Intelligence Hub</p>
  <p class="hero-sub">
    Period Variance · Cross-Selling · Opportunity Scope · RBAC ·
    Year / Quarter / Month Filters · Multi-format Export
  </p>
</div>""", unsafe_allow_html=True)

with st.expander("▸ Analysis Logic & Business Rules", expanded=False):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
<div class="lc"><div class="lt">PERIOD — 3 COMPARISONS</div><div class="lb">
(a) Latest Month vs Previous Month<br>
(b) Latest Quarter vs Previous Quarter<br>
(c) Latest Year vs Previous Year<br><br>
Time filters narrow the dataset before analysis.
</div></div>
<div class="lc"><div class="lt">CRITICAL THRESHOLD</div><div class="lb">
✗ Full Sales Drop — sold prev, zero in current<br>
✗ Critical Decline — fell ≥ 50% period-on-period<br>
Noise filters: prev sales ≥ 500, ≥ 2 invoices/period
</div></div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("""
<div class="lc"><div class="lt">OPPORTUNITY SCOPE (5 DIMENSIONS)</div><div class="lb">
1. Revenue growth gap — region×product vs median<br>
2. Under-serviced stores — invoice gap vs 25th pct<br>
3. High-potential regions — revenue per invoice<br>
4. Visit optimisation — visit freq vs median<br>
5. Salesperson productivity — rev per call-minute
</div></div>
<div class="lc"><div class="lt">ROLE-BASED ACCESS</div><div class="lb">
Management — full visibility<br>
Regional Manager — own region only<br>
Salesperson — own performance & stores only
</div></div>""", unsafe_allow_html=True)

# ── NO FILE STATE ─────────────────────────────────────────────────────────────
if uploaded_file is None:
    st.markdown("""<div class="ibox">
    ← Upload your sales CSV or Excel file from the sidebar to begin.
    Results are paginated in 200-row chunks to handle large datasets.
    </div>""", unsafe_allow_html=True)
    st.markdown('<p class="sh">EXPECTED COLUMNS</p>', unsafe_allow_html=True)
    # Use a plain markdown table — avoids the DataFrame JS widget which can
    # crash with "Failed to fetch dynamically imported module" on cold starts
    # or when the Streamlit deployment has mismatched cached frontend assets.
    rows_md = "\n".join(f"| `{c}` |" for c in EXPECTED_COLS)
    st.markdown(
        "| Column |\n|--------|\n" + rows_md,
        unsafe_allow_html=False,
    )
    st.stop()

# ── LOAD DATA ─────────────────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    file_bytes     = uploaded_file.read()
    df_raw, missing_cols = load_data(file_bytes, uploaded_file.name)

if missing_cols:
    st.markdown(f'<div class="wbox">⚠ Missing columns (will proceed): '
                f'{", ".join(missing_cols)}</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  (a) TIME-BASED FILTERS — fully dynamic from data
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<p class="sh">TIME FILTERS</p>', unsafe_allow_html=True)
tf1, tf2, tf3 = st.columns(3)

all_years    = sorted(df_raw["_year"].unique().tolist())
all_quarters = sorted(df_raw["_quarter"].astype(str).unique().tolist())
all_month_nums = sorted(df_raw["_month_num"].unique().tolist())
month_label_map = {m: pd.Timestamp(2000, m, 1).strftime("%B") for m in all_month_nums}

with tf1:
    sel_years = st.multiselect("Year", options=all_years,
                               default=[], key="tf_year",
                               help="Leave empty to include all years")
with tf2:
    sel_quarters = st.multiselect("Quarter", options=all_quarters,
                                  default=[], key="tf_quarter",
                                  help="e.g. 2024Q1, 2024Q2")
with tf3:
    sel_months = st.multiselect(
        "Month",
        options=all_month_nums,
        format_func=lambda m: month_label_map.get(m, str(m)),
        default=[], key="tf_month",
    )

# Apply time filters
df = apply_time_filters(df_raw.copy(), sel_years, sel_quarters, sel_months)
if df.empty:
    st.warning("No data for the selected time period. Adjust filters.")
    st.stop()

# ═══════════════════════════════════════════════════════════════════════════════
#  (b) RBAC — populate identity picker from filtered data
# ═══════════════════════════════════════════════════════════════════════════════
if user_role == ROLE_REGIONAL:
    rm_options = sorted(df["Regional_Manager"].dropna().unique().tolist()) \
                 if "Regional_Manager" in df.columns else \
                 sorted(df["Region"].dropna().unique().tolist())
    user_identity = identity_placeholder.selectbox(
        "Select Regional Manager", rm_options, key="rb_identity")
elif user_role == ROLE_SALES:
    sp_options = sorted(df["Salesperson_Name"].dropna().unique().tolist()) \
                 if "Salesperson_Name" in df.columns else []
    user_identity = identity_placeholder.selectbox(
        "Select Salesperson", sp_options, key="rb_identity")
else:
    user_identity = None
    identity_placeholder.empty()

# Apply RBAC to the working dataframe
df = apply_rbac(df, user_role, user_identity)
if df.empty:
    st.warning("No data available for the selected role / identity. "
               "Check your selection or data coverage.")
    st.stop()

# ── METRICS ───────────────────────────────────────────────────────────────────
ts  = df["LineTotal"].sum()
ns  = df["StoreID"].nunique()
np_ = df["Product"].nunique()
ni  = df["InvoiceNo"].nunique()
nm  = df["_month"].nunique()
dr  = f"{df['Date'].min().strftime('%d %b %y')} → {df['Date'].max().strftime('%d %b %y')}"

extra_metrics = ""
if "Salesperson_Name" in df.columns:
    nsp = df["Salesperson_Name"].nunique()
    extra_metrics += (f'<div class="mcard"><div class="mlabel">Salespersons</div>'
                      f'<div class="mval">{nsp}</div></div>')
if "Regional_Manager" in df.columns:
    nrm = df["Regional_Manager"].nunique()
    extra_metrics += (f'<div class="mcard"><div class="mlabel">Reg. Managers</div>'
                      f'<div class="mval">{nrm}</div></div>')

st.markdown(f"""<div class="mstrip">
  <div class="mcard"><div class="mlabel">Revenue</div>
    <div class="mval c-blue">{ts/1_000_000:.1f}M</div></div>
  <div class="mcard"><div class="mlabel">Stores</div>
    <div class="mval">{ns:,}</div></div>
  <div class="mcard"><div class="mlabel">Products</div>
    <div class="mval">{np_}</div></div>
  <div class="mcard"><div class="mlabel">Invoices</div>
    <div class="mval">{ni:,}</div></div>
  <div class="mcard"><div class="mlabel">Months</div>
    <div class="mval">{nm}</div></div>
  {extra_metrics}
  <div class="mcard"><div class="mlabel">Date Range</div>
    <div class="mval" style="font-size:.8rem;color:#718096;">{dr}</div></div>
</div>""", unsafe_allow_html=True)

# ── SERIALISE FILTERED DF TO PARQUET (for cache-friendly function args) ────────
@st.cache_data(show_spinner=False)
def _df_to_parquet(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_parquet(buf, index=False)
    return buf.getvalue()

df_bytes = _df_to_parquet(df)

# ═══════════════════════════════════════════════════════════════════════════════
#  RUN
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<p class="sh">RUN ANALYSIS</p>', unsafe_allow_html=True)

if st.button("▶  Run Analysis"):
    all_frames = []

    if run_period:
        with st.spinner("Period analysis…"):
            df_period, period_summary = run_all_period_analyses(df_bytes)
        st.session_state["period_summary"] = period_summary
        if not df_period.empty:
            keep = [c for c in OUTPUT_COLS if c in df_period.columns]
            st.session_state["df_period"]     = df_period[keep]
            st.session_state["df_period_csv"] = df_period[keep].to_csv(index=False).encode()
            all_frames.append(df_period)
        else:
            for k in ["df_period", "df_period_csv"]:
                st.session_state.pop(k, None)

    if run_cross:
        with st.spinner("Cross-sell analysis…"):
            df_cross = run_cross_sell_analysis(df_bytes)
        if not df_cross.empty:
            keep = [c for c in OUTPUT_COLS if c in df_cross.columns]
            st.session_state["df_cross"]     = df_cross[keep]
            st.session_state["df_cross_csv"] = df_cross[keep].to_csv(index=False).encode()
            all_frames.append(df_cross)
        else:
            for k in ["df_cross", "df_cross_csv"]:
                st.session_state.pop(k, None)

    if run_scope:
        with st.spinner("Opportunity scope analysis…"):
            scope_results = run_opportunity_scope(df_bytes)
        st.session_state["scope_results"] = scope_results

    if all_frames:
        combined = (pd.concat(all_frames, ignore_index=True)
                      .sort_values("Impact", ascending=False)
                      .reset_index(drop=True))
        keep = [c for c in OUTPUT_COLS if c in combined.columns]
        combined = combined[keep]
        st.session_state["combined_csv"]      = combined.to_csv(index=False).encode()
        st.session_state["combined_len"]      = len(combined)
        st.session_state["combined_stores"]   = combined["Store_ID"].nunique()
        st.session_state["combined_products"] = combined["Product"].nunique()
        st.session_state["period_rows"] = combined[
            combined["Analysis_Type"].str.startswith("Period", na=False)].shape[0]
        st.session_state["cross_rows"] = combined[
            combined["Analysis_Type"] == "Cross Selling"].shape[0]
        # Persist filter/role context for export metadata
        st.session_state["export_meta"] = {
            "Role": user_role,
            "Identity": user_identity or "All",
            "Years":    ", ".join(map(str, sel_years))    or "All",
            "Quarters": ", ".join(sel_quarters)           or "All",
            "Months":   ", ".join(month_label_map.get(m, str(m)) for m in sel_months) or "All",
            "Records":  len(combined),
            "Date Range": dr,
        }
        st.success(f"✓ {len(combined):,} findings — "
                   f"{combined['Store_ID'].nunique()} stores, "
                   f"{combined['Product'].nunique()} products.")
    elif run_scope and "scope_results" in st.session_state:
        st.info("No critical period / cross-sell findings, but scope results are available below.")
    else:
        st.warning("No findings. Try loosening filters or check data coverage.")
        for k in ["combined_csv", "df_period", "df_cross", "df_period_csv",
                  "df_cross_csv", "scope_results"]:
            st.session_state.pop(k, None)


# ═══════════════════════════════════════════════════════════════════════════════
#  DISPLAY TABS
# ═══════════════════════════════════════════════════════════════════════════════
has_combined = "combined_csv" in st.session_state
has_scope    = "scope_results" in st.session_state

if has_combined or has_scope:

    pr  = st.session_state.get("period_rows", 0)
    cr_ = st.session_state.get("cross_rows",  0)
    tot = st.session_state.get("combined_len", 0)

    st.markdown(f"""<div style="display:flex;gap:10px;margin:10px 0 18px;flex-wrap:wrap;">
      <span class="badge bc">Period: {pr}</span>
      <span class="badge bh">Cross-Sell: {cr_}</span>
      <span class="badge bc">Total Findings: {tot:,}</span>
      <span class="badge bg">Role: {user_role}</span>
    </div>""", unsafe_allow_html=True)

    tab_labels = ["📉 Period Analysis", "🔗 Cross-Selling",
                  "📋 Combined Output", "🎯 Opportunity Scope", "⬇ Export"]
    tab1, tab2, tab3, tab4, tab5 = st.tabs(tab_labels)

    # ── TAB 1: Period ────────────────────────────────────────────────────────
    with tab1:
        if "df_period" in st.session_state:
            dp = st.session_state["df_period"]
            if "period_summary" in st.session_state:
                ps = st.session_state["period_summary"]
                c1_, c2_, c3_ = st.columns(3)
                for col, lbl in zip([c1_, c2_, c3_], ["Monthly", "Quarterly", "Yearly"]):
                    info = ps.get(lbl, {})
                    with col:
                        if "note" in info:
                            st.markdown(f'<div class="wbox"><b>{lbl}</b><br>{info["note"]}</div>',
                                        unsafe_allow_html=True)
                        else:
                            st.markdown(f'<div class="ibox"><b>{lbl}</b><br>'
                                        f'{info.get("previous","")} → {info.get("current","")}<br>'
                                        f'Critical rows: <b>{info.get("critical_rows",0)}</b></div>',
                                        unsafe_allow_html=True)

            # Region filter (RBAC-aware — Salesperson sees only their data)
            regions = ["All"] + sorted(dp["Region"].dropna().unique().tolist()) \
                      if "Region" in dp.columns else ["All"]
            sel_r_p = st.selectbox("Filter by Region", regions, key="p_region")
            obs_types = ["All"] + sorted(dp["Observation_Type"].dropna().unique().tolist()) \
                        if "Observation_Type" in dp.columns else ["All"]
            sel_obs = st.selectbox("Filter by Observation Type", obs_types, key="p_obs")

            view = dp.copy()
            if sel_r_p != "All":
                view = view[view["Region"] == sel_r_p]
            if sel_obs != "All":
                view = view[view["Observation_Type"] == sel_obs]
            paginated_table(view, key="period_tab")
        else:
            st.markdown('<div class="ibox">No critical period findings.</div>',
                        unsafe_allow_html=True)

    # ── TAB 2: Cross-Selling ────────────────────────────────────────────────
    with tab2:
        if "df_cross" in st.session_state:
            dc = st.session_state["df_cross"]
            regions_c = ["All"] + sorted(dc["Region"].dropna().unique().tolist()) \
                        if "Region" in dc.columns else ["All"]
            sel_r_c = st.selectbox("Filter by Region", regions_c, key="c_region")
            types_c = ["All"] + sorted(dc["Reason_Type"].dropna().unique().tolist())
            sel_c   = st.selectbox("Filter by Opportunity Type", types_c, key="ct_sel")
            view_c  = dc.copy()
            if sel_r_c != "All":
                view_c = view_c[view_c["Region"] == sel_r_c]
            if sel_c != "All":
                view_c = view_c[view_c["Reason_Type"] == sel_c]
            paginated_table(view_c, key="cross_tab")
        else:
            st.markdown('<div class="ibox">No cross-selling findings.</div>',
                        unsafe_allow_html=True)

    # ── TAB 3: Combined ─────────────────────────────────────────────────────
    with tab3:
        if has_combined:
            st.markdown(f'<p class="sh">COMBINED — {tot:,} ROWS</p>',
                        unsafe_allow_html=True)
            combined_disp = pd.read_csv(io.BytesIO(st.session_state["combined_csv"]))
            top10 = (combined_disp.groupby("Product")["Impact"]
                                   .sum().nlargest(10).reset_index())
            top10["Impact"] = top10["Impact"].apply(lambda x: f"{x:,.0f}")
            with st.expander("▸ Top 10 products by at-risk / opportunity revenue"):
                st.dataframe(top10, use_container_width=True, hide_index=True)

            obs_all = ["All"] + sorted(combined_disp["Observation_Type"].dropna().unique()) \
                      if "Observation_Type" in combined_disp.columns else ["All"]
            sel_obs_all = st.selectbox("Filter by Observation Type",
                                       obs_all, key="comb_obs")
            regions_all = ["All"] + sorted(combined_disp["Region"].dropna().unique())
            sel_r_all   = st.selectbox("Filter by Region", regions_all, key="comb_reg")
            view_all = combined_disp.copy()
            if sel_obs_all != "All":
                view_all = view_all[view_all["Observation_Type"] == sel_obs_all]
            if sel_r_all != "All":
                view_all = view_all[view_all["Region"] == sel_r_all]
            paginated_table(view_all, key="all_tab", height=420)
        else:
            st.markdown('<div class="ibox">Run analysis first to see combined output.</div>',
                        unsafe_allow_html=True)

    # ── TAB 4: Opportunity Scope ────────────────────────────────────────────
    with tab4:
        st.markdown('<p class="sh">OPPORTUNITY SCOPE ASSESSMENT</p>',
                    unsafe_allow_html=True)
        if has_scope:
            sc = st.session_state["scope_results"]

            # 1. Revenue Growth
            st.markdown("#### 1 · Revenue Growth Opportunities")
            df_rg = sc.get("revenue_growth", pd.DataFrame())
            if not df_rg.empty:
                total_gap = df_rg["Gap"].sum()
                st.markdown(f'<div class="gbox">💰 Total identified revenue gap: '
                            f'<b>{total_gap:,.0f}</b> across '
                            f'{df_rg["Region"].nunique()} regions × '
                            f'{df_rg["Product"].nunique()} products</div>',
                            unsafe_allow_html=True)
                fig_rg = chart_revenue_growth(df_rg)
                if fig_rg:
                    st.pyplot(fig_rg, use_container_width=True)
                    plt.close(fig_rg)
                with st.expander("▸ Full data table"):
                    st.dataframe(df_rg, use_container_width=True, hide_index=True)
            else:
                st.info("Insufficient data for revenue growth analysis.")

            st.divider()

            # 2. Under-Serviced Stores
            st.markdown("#### 2 · Under-Serviced / Low-Coverage Stores")
            df_us = sc.get("under_serviced", pd.DataFrame())
            if not df_us.empty:
                total_uplift = df_us["Est_Uplift"].sum()
                st.markdown(f'<div class="gbox">🏪 Estimated uplift if visit gaps closed: '
                            f'<b>{total_uplift:,.0f}</b> across '
                            f'{len(df_us)} stores</div>', unsafe_allow_html=True)
                fig_us = chart_under_serviced(df_us)
                if fig_us:
                    st.pyplot(fig_us, use_container_width=True)
                    plt.close(fig_us)
                with st.expander("▸ Full data table"):
                    st.dataframe(df_us, use_container_width=True, hide_index=True)
            else:
                st.info("Insufficient data for under-serviced store analysis.")

            st.divider()

            # 3. High-Potential Regions
            st.markdown("#### 3 · High-Potential Regions & Products")
            df_hp = sc.get("high_potential_regions", pd.DataFrame())
            if not df_hp.empty:
                top_region = df_hp.iloc[0]["Region"] if not df_hp.empty else "N/A"
                top_rpi    = df_hp.iloc[0]["Rev_per_Invoice"] if not df_hp.empty else 0
                st.markdown(f'<div class="gbox">📍 Top region: <b>{top_region}</b> — '
                            f'Revenue per invoice: <b>{top_rpi:,.0f}</b></div>',
                            unsafe_allow_html=True)
                fig_hp = chart_region_potential(df_hp)
                if fig_hp:
                    st.pyplot(fig_hp, use_container_width=True)
                    plt.close(fig_hp)
                with st.expander("▸ Full data table"):
                    st.dataframe(df_hp, use_container_width=True, hide_index=True)
            else:
                st.info("Insufficient data for regional potential analysis.")

            st.divider()

            # 4. Visit Optimisation
            st.markdown("#### 4 · Store Visit Optimisation")
            df_vo = sc.get("visit_optimisation", pd.DataFrame())
            if not df_vo.empty:
                high_pri = df_vo[df_vo.get("Priority", pd.Series()) == "High"]
                st.markdown(f'<div class="wbox">⚠ <b>{len(high_pri)}</b> high-revenue stores '
                            f'are below median visit frequency — '
                            f'immediate priority for increased visits.</div>',
                            unsafe_allow_html=True)
                fig_vo = chart_visit_optimisation(df_vo)
                if fig_vo:
                    st.pyplot(fig_vo, use_container_width=True)
                    plt.close(fig_vo)
                with st.expander("▸ Full data table"):
                    st.dataframe(df_vo, use_container_width=True, hide_index=True)
            else:
                st.info("Visit Frequency column not found — skipping visit optimisation.")

            st.divider()

            # 5. Salesperson Productivity
            st.markdown("#### 5 · Sales Productivity Improvement Areas")
            df_sp = sc.get("salesperson_productivity", pd.DataFrame())
            if not df_sp.empty:
                bottom_q = df_sp[df_sp["Productivity_Gap"] > df_sp["Productivity_Gap"].median()]
                st.markdown(f'<div class="wbox">👤 <b>{len(bottom_q)}</b> salespersons '
                            f'below median productivity — targeted coaching recommended.</div>',
                            unsafe_allow_html=True)
                fig_sp = chart_salesperson_productivity(df_sp)
                if fig_sp:
                    st.pyplot(fig_sp, use_container_width=True)
                    plt.close(fig_sp)
                with st.expander("▸ Full data table"):
                    st.dataframe(df_sp, use_container_width=True, hide_index=True)
            else:
                st.info("Avg_Call_Duration / Salesperson_Name columns not found — "
                        "skipping productivity analysis.")
        else:
            st.markdown('<div class="ibox">Run analysis with "Opportunity Scope Analysis" '
                        'enabled to view this section.</div>', unsafe_allow_html=True)

    # ── TAB 5: Export ───────────────────────────────────────────────────────
    with tab5:
        st.markdown('<p class="sh">EXPORT REPORTS</p>', unsafe_allow_html=True)

        meta = st.session_state.get("export_meta", {
            "Role": user_role, "Identity": user_identity or "All",
            "Years": "All", "Quarters": "All", "Months": "All",
            "Date Range": dr,
        })
        st.markdown(f'<div class="ibox">Current export scope: '
                    f'Role=<b>{meta.get("Role")}</b> · '
                    f'Identity=<b>{meta.get("Identity")}</b> · '
                    f'Years={meta.get("Years")} · '
                    f'Quarters={meta.get("Quarters")} · '
                    f'Months={meta.get("Months")}</div>',
                    unsafe_allow_html=True)

        ecol1, ecol2, ecol3 = st.columns(3)

        # ── CSV ───────────────────────────────────────────────────────────
        with ecol1:
            st.markdown("**CSV Export**")
            if has_combined:
                st.download_button(
                    "⬇  Download Combined CSV",
                    data=st.session_state["combined_csv"],
                    file_name="sales_analysis_combined.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            if "df_period_csv" in st.session_state:
                st.download_button(
                    "⬇  Period Analysis CSV",
                    data=st.session_state["df_period_csv"],
                    file_name="sales_period_analysis.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            if "df_cross_csv" in st.session_state:
                st.download_button(
                    "⬇  Cross-Sell Analysis CSV",
                    data=st.session_state["df_cross_csv"],
                    file_name="sales_cross_sell.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

        # ── EXCEL ─────────────────────────────────────────────────────────
        with ecol2:
            st.markdown("**Excel Export (.xlsx)**")
            sheets = {}
            if has_combined:
                sheets["Combined"] = pd.read_csv(
                    io.BytesIO(st.session_state["combined_csv"]))
            if "df_period" in st.session_state:
                sheets["Period Analysis"] = st.session_state["df_period"]
            if "df_cross" in st.session_state:
                sheets["Cross Selling"]   = st.session_state["df_cross"]
            if has_scope:
                sc_ = st.session_state["scope_results"]
                for key, label in [
                    ("revenue_growth",        "Revenue Growth Gaps"),
                    ("under_serviced",        "Under-Serviced Stores"),
                    ("high_potential_regions","High Potential Regions"),
                    ("visit_optimisation",    "Visit Optimisation"),
                    ("salesperson_productivity","Salesperson Productivity"),
                ]:
                    dft = sc_.get(key, pd.DataFrame())
                    if dft is not None and not dft.empty:
                        sheets[label] = dft
            if sheets:
                xlsx_bytes = to_excel_bytes(sheets)
                st.download_button(
                    "⬇  Download Excel Report",
                    data=xlsx_bytes,
                    file_name="sales_intelligence_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.markdown(f'<div class="gbox" style="font-size:.74rem;">'
                            f'Sheets: {", ".join(sheets.keys())}</div>',
                            unsafe_allow_html=True)
            else:
                st.info("Run analysis first.")

        # ── PDF ───────────────────────────────────────────────────────────
        with ecol3:
            st.markdown("**PDF Summary Report**")
            if st.button("🖨  Generate PDF", use_container_width=True):
                with st.spinner("Generating PDF…"):
                    pdf_figs, pdf_dfs = {}, {}
                    if has_scope:
                        sc_ = st.session_state["scope_results"]
                        df_rg2 = sc_.get("revenue_growth",         pd.DataFrame())
                        df_us2 = sc_.get("under_serviced",         pd.DataFrame())
                        df_hp2 = sc_.get("high_potential_regions", pd.DataFrame())
                        df_vo2 = sc_.get("visit_optimisation",     pd.DataFrame())
                        df_sp2 = sc_.get("salesperson_productivity",pd.DataFrame())
                        pdf_figs["Revenue Growth Opportunities"] = chart_revenue_growth(df_rg2)
                        pdf_figs["Under-Serviced Stores"]        = chart_under_serviced(df_us2)
                        pdf_figs["Region Potential"]             = chart_region_potential(df_hp2)
                        pdf_figs["Visit Optimisation"]           = chart_visit_optimisation(df_vo2)
                        pdf_figs["Salesperson Productivity"]     = chart_salesperson_productivity(df_sp2)
                        for k, v in [("Revenue Growth", df_rg2), ("Under-Serviced", df_us2),
                                     ("High Potential", df_hp2), ("Visit Opt.", df_vo2),
                                     ("SP Productivity", df_sp2)]:
                            if v is not None and not v.empty:
                                pdf_dfs[k] = v
                    if has_combined:
                        pdf_dfs["Combined Findings"] = pd.read_csv(
                            io.BytesIO(st.session_state["combined_csv"]))
                    pdf_bytes = to_pdf_bytes(pdf_dfs, pdf_figs, meta)
                    st.session_state["pdf_bytes"] = pdf_bytes

            if "pdf_bytes" in st.session_state:
                st.download_button(
                    "⬇  Download PDF",
                    data=st.session_state["pdf_bytes"],
                    file_name="sales_intelligence_summary.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

    # ── RAW PREVIEW ─────────────────────────────────────────────────────────
    with st.expander("▸ Raw input data preview (first 100 rows — filtered & RBAC scoped)"):
        preview_cols = ["StoreID", "Store_Name", "Region", "Regional_Manager",
                        "Salesperson_Name", "Date", "InvoiceNo",
                        "Product", "Qty", "LineTotal", "Visit Frequency"]
        show = [c for c in preview_cols if c in df.columns]
        st.dataframe(df[show].head(100), use_container_width=True, hide_index=True)
