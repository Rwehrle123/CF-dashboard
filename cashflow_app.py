"""
SHG Cash Flow Dashboard — Complete App
=======================================
Corrected version: 3-Month Focus and Weekly 4+13 Outlook both use TOTAL CASH
(UK + Ireland), not UK cash, so the two tabs reconcile on the same basis.

Run:
    streamlit run cashflow_app.py

Install:
    pip install streamlit pandas openpyxl plotly
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import warnings
import calendar

warnings.filterwarnings("ignore")

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SHG Cash Dashboard",
    page_icon="💷",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── SHG brand styling ─────────────────────────────────────────────────────────
st.markdown("""
<style>
div[data-testid="stTabs"] button[aria-selected="true"] {
    color: #1C1464 !important;
    border-bottom-color: #1C1464 !important;
    font-weight: 600;
}
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] .stMarkdown strong {
    color: #C9A84C;
}
[data-testid="stMetricLabel"] { color: #1C1464 !important; }
hr { border-color: #C9A84C !important; opacity: 0.4; }
thead tr th { background-color: #1C1464 !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ACCOUNT → ENTITY MAPPING
# ══════════════════════════════════════════════════════════════════════════════
ACCOUNT_ENTITY = {
    # UK entities
    "SHTL GBP": "UK",
    "SHTL EUR": "UK",
    "SHTL USD": "UK",
    "SHTL PAY GBP": "UK",
    "SHTL PAY": "UK",
    "TMOOD GBP": "UK",
    "TMOOD EUR": "UK",
    "TMOOD USD": "UK",
    "TMOOD CAD": "UK",
    "HJT GBP": "UK",
    "HJT AH GBP": "UK",

    # Ireland entities
    "SHGI EUR": "Ireland",
    "SHGI GBP": "Ireland",
    "SHGI USD": "Ireland",
    "SHGI CAD": "Ireland",
    "AHD EURO CURRENT ACCOUNT": "Ireland",
    "BOA": "Ireland",
}

UK_GBP_ACCS = ["SHTL GBP", "SHTL PAY GBP", "SHTL PAY", "TMOOD GBP", "HJT GBP", "HJT AH GBP"]
UK_ACCS = [k for k, v in ACCOUNT_ENTITY.items() if v == "UK"]
IRELAND_ACCS = [k for k, v in ACCOUNT_ENTITY.items() if v == "Ireland"]

KEY_INFLOW = [
    "AGENT RECEIPTS", "FD RECEIPT", "DIRECT RECEIPTS", "FX TRADE IN",
    "OTHER RECEIPTS", "TUI RECEIPT", "INTERCO"
]
KEY_OUTFLOW = [
    "AP COGS", "AP OVH", "FLIGHT COSTS", "CUSTOMER REFUNDS",
    "FX TRADE OUT", "PAYROLL", "OTHER COSTS", "TAX", "INTERCO"
]

RECEIPT_LABELS = [
    "DIRECT RECEIPTS", "AGENT RECEIPTS", "FD RECEIPT", "CUSTOMER REFUND",
    "TUI RECEIPT", "OTHER RECEIPT", "FX TRADE IN", "INTERCO (net)", "OD INTEREST"
]
PAYMENT_LABELS = [
    "AP COGS", "AP OVH", "PAYROLL", "TAX", "OTHER CASH OUT", "FX TRADE OUT", "FLIGHT COSTS"
]

WEEKLY_LOCK = {"PAYROLL"}
MN = {1:"Jan", 2:"Feb", 3:"Mar", 4:"Apr", 5:"May", 6:"Jun", 7:"Jul", 8:"Aug", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dec"}

BLUE = "#1C1464"
GREEN = "#3B6D11"
LBLUE = "#3D3580"
LGREEN = "#639922"
RED = "#E24B4A"
AMBER = "#C9A84C"
GREY = "rgba(0,0,0,0.06)"
DEFAULT_FX = {"EUR": 0.86, "USD": 0.76, "CAD": 0.53, "GBP": 1.0}

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(v, decimals=1):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    sign = "-" if v < 0 else ""
    a = abs(float(v))
    if a >= 1e6:
        return f"{sign}£{a/1e6:.{decimals}f}m"
    if a >= 1e3:
        return f"{sign}£{a/1e3:.0f}k"
    return f"{sign}£{a:,.0f}"


def fkw(v, dash=True):
    if v is None:
        return ""
    if abs(v) < 0.0001 and dash:
        return "—"
    return ("−" if v < 0 else "") + "£" + f"{abs(round(float(v))):,}k"


def safe_get(df, yr, mn, col):
    try:
        if (yr, mn) in df.index and col in df.columns:
            v = df.loc[(yr, mn), col]
            return float(v) if not pd.isna(v) else 0.0
    except Exception:
        pass
    return 0.0


def safe_entity(entity_m, yr, mn, ent, field):
    try:
        return float(entity_m.loc[(yr, mn, ent), field])
    except Exception:
        return 0.0


def get_currency(acc):
    acc = str(acc).strip().upper()
    if "EUR" in acc:
        return "EUR"
    if "USD" in acc:
        return "USD"
    if "CAD" in acc:
        return "CAD"
    if acc == "BOA":
        return "USD"
    return "GBP"

# ── Data loading ───────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_transactions(b):
    df = pd.read_excel(io.BytesIO(b), sheet_name="Data Sheet")
    df["PostDate"] = pd.to_datetime(df["PostDate"])
    df["AccountName"] = df["AccountName"].astype(str).str.strip()
    df["entity"] = df["AccountName"].map(ACCOUNT_ENTITY)

    unmapped = df.loc[df["entity"].isna(), "AccountName"].unique()
    if len(unmapped) > 0:
        st.warning(
            f"⚠️ Unmapped accounts detected: **{', '.join(unmapped)}**. "
            f"Defaulting to Ireland. Add these to ACCOUNT_ENTITY."
        )
    df["entity"] = df["entity"].fillna("Ireland")

    df["TrnSpec"] = df["TrnSpec"].fillna("OTHER").astype(str).str.strip()
    df.loc[df["TrnSpec"].str.upper().str.strip() == "PAYROLL", "TrnSpec"] = "PAYROLL"
    df["Currency"] = df["AccountName"].apply(get_currency)
    df["Year"] = df["PostDate"].dt.year.astype(int)
    df["Month"] = df["PostDate"].dt.month.astype(int)
    df["WeekStart"] = df["PostDate"].dt.to_period("W-SUN").apply(lambda p: p.start_time)

    df_op = df[df["TrnSpec"] != "OVERNIGHT DEPOSIT"].copy()
    df_op["inflow"] = df_op["Amount"].clip(lower=0)
    df_op["outflow"] = df_op["Amount"].clip(upper=0)
    return df, df_op


@st.cache_data(show_spinner=False)
def load_forecast(b):
    raw = pd.read_excel(io.BytesIO(b), sheet_name="Sheet1", header=None)
    dates = pd.to_datetime(raw.iloc[0, 1:].values)
    fc = pd.DataFrame({
        "client_money": raw.iloc[1, 1:].values.astype(float),
        "cash_uk": raw.iloc[2, 1:].values.astype(float),
        "cash_ireland": raw.iloc[3, 1:].values.astype(float),
    }, index=dates)
    fc.index.name = "Date"
    fc["Year"] = fc.index.year.astype(int)
    fc["Month"] = fc.index.month.astype(int)
    return fc


def extend_forecast_horizon(fc, latest_bank_date):
    hard_min = pd.Timestamp("2027-12-01")
    rolling_min = (latest_bank_date + pd.DateOffset(months=12)).replace(day=1)
    required_end = max(hard_min, rolling_min)

    if fc.index.max() >= required_end:
        return fc

    extra_dates = pd.date_range(fc.index.max() + pd.DateOffset(months=1), required_end, freq="MS")
    rows = []
    for dt in extra_dates:
        prior_dt = dt - pd.DateOffset(years=1)
        base = fc.loc[prior_dt] if prior_dt in fc.index else fc.iloc[-1]
        rows.append({
            "client_money": float(base["client_money"]),
            "cash_uk": float(base["cash_uk"]),
            "cash_ireland": float(base["cash_ireland"]),
        })
    extra = pd.DataFrame(rows, index=extra_dates)
    extra.index.name = "Date"
    extra["Year"] = extra.index.year.astype(int)
    extra["Month"] = extra.index.month.astype(int)
    return pd.concat([fc, extra]).loc[lambda d: ~d.index.duplicated(keep="first")].sort_index()


def build_actuals(df_op):
    entity_m = df_op.groupby(["Year", "Month", "entity"]).agg(
        inflow=("inflow", "sum"),
        outflow=("outflow", "sum"),
        net=("Amount", "sum"),
    ).round(0)

    def piv(ent, flow):
        col = "inflow" if flow == "in" else "outflow"
        return (
            df_op[df_op["entity"] == ent]
            .groupby(["Year", "Month", "TrnSpec"])[col]
            .sum()
            .round(0)
            .unstack("TrnSpec")
            .fillna(0)
        )

    uk_in = piv("UK", "in")
    uk_out = piv("UK", "out")
    ie_in = piv("Ireland", "in")
    ie_out = piv("Ireland", "out")
    return entity_m, uk_in, uk_out, ie_in, ie_out


def build_weekly(df_raw, fx_rates):
    """
    Weekly pivot — TOTAL GROUP CASH, FX converted to GBP at budget rates.
    This includes UK + Ireland and is used by both 3-Month Focus and 4+13 Outlook.
    """
    df = df_raw.copy()
    df["AccountName"] = df["AccountName"].astype(str).str.strip()
    df["entity"] = df["AccountName"].map(ACCOUNT_ENTITY).fillna("Ireland")
    df["TrnSpec"] = df["TrnSpec"].fillna("OTHER").astype(str).str.strip()
    df.loc[df["TrnSpec"].str.upper().str.strip() == "PAYROLL", "TrnSpec"] = "PAYROLL"
    df["Currency"] = df["AccountName"].apply(get_currency)
    df["PostDate"] = pd.to_datetime(df["PostDate"])
    df["WeekStart"] = df["PostDate"].dt.to_period("W-SUN").apply(lambda p: p.start_time)
    df["Amount_GBP"] = df.apply(lambda r: r["Amount"] * fx_rates.get(r["Currency"], 1.0), axis=1)

    df_group = df[df["TrnSpec"] != "OVERNIGHT DEPOSIT"].copy()
    df_od = df[df["TrnSpec"] == "OVERNIGHT DEPOSIT"].copy()

    weekly = (
        df_group.groupby(["WeekStart", "TrnSpec"])["Amount_GBP"]
        .sum().round(0).unstack("TrnSpec").fillna(0)
    )
    od_weekly = df_od.groupby("WeekStart")["Amount_GBP"].sum().round(0)
    return weekly, od_weekly


def build_weekly_forecast(weekly, od_weekly, n_fc=13):
    """Forecast weeks use prior year same ISO week with simple fallbacks. Interco defaults to zero."""
    last_date = weekly.index.max()
    fc_weeks = [last_date + pd.Timedelta(weeks=i+1) for i in range(n_fc)]

    all_specs = [
        "DIRECT RECEIPTS", "AGENT RECEIPTS", "FD RECEIPT", "CUSTOMER REFUNDS", "TUI RECEIPT",
        "OTHER RECEIPTS", "FX TRADE IN", "INTERCO", "AP COGS", "AP OVH", "PAYROLL", "TAX",
        "OTHER COSTS", "FX TRADE OUT", "FLIGHT COSTS"
    ]
    forecast = {spec: [] for spec in all_specs}
    forecast["INTERCO"] = [0] * n_fc
    forecast["OD_INTEREST"] = [11000] * n_fc

    # Payroll cycle defaults
    large, small, days_since = -250000, -30000, 21
    if "PAYROLL" in weekly.columns:
        pay_s = weekly["PAYROLL"].replace(0, np.nan).dropna()
        if len(pay_s) > 0:
            med = pay_s.median()
            if pd.notna(med) and med != 0:
                large_s = pay_s[pay_s < med * 2]
                small_s = pay_s[pay_s >= med]
                if pd.notna(large_s.median()):
                    large = float(large_s.median())
                if len(small_s) > 0 and pd.notna(small_s.median()):
                    small = float(small_s.median())
            days_since = int((last_date - pay_s.index.max()).days)

    pay_fc = []
    for fw in fc_weeks:
        da = (fw - last_date).days + days_since
        if da % 28 < 7:
            pay_fc.append(int(large))
        elif (da + 14) % 28 < 7:
            pay_fc.append(int(small))
        else:
            pay_fc.append(0)
    forecast["PAYROLL"] = pay_fc

    for spec in all_specs:
        if spec in ("INTERCO", "PAYROLL"):
            continue
        vals = []
        for fw in fc_weeks:
            py_date = fw - pd.Timedelta(weeks=52)
            base = None
            for delta in [0, 1, -1, 2, -2]:
                sd = py_date + pd.Timedelta(weeks=delta)
                if spec in weekly.columns and sd in weekly.index:
                    base = float(weekly.loc[sd, spec])
                    break
            if base is None:
                if spec in weekly.columns:
                    same_mn = weekly[spec][weekly.index.month == fw.month]
                    base = float(same_mn.median()) if len(same_mn) > 0 else 0.0
                else:
                    base = 0.0
            vals.append(round(base))
        forecast[spec] = vals

    return forecast, fc_weeks

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 💷 SHG Cash Dashboard")
    st.divider()
    st.markdown("**1 · Bank transactions**")
    tx_file = st.file_uploader("Upload bank Excel", type=["xlsx"], key="tx", help="Must contain 'Data Sheet' tab")
    st.markdown("**2 · Forecast / client money**")
    fc_file = st.file_uploader("Upload forecast Excel", type=["xlsx"], key="fc", help="Rows: Client money / Cash UK / Cash Ireland")
    st.divider()

    uk_pct = st.slider("UK cash requirement (%)", 50, 100, 70, 5)
    st.divider()
    st.markdown("**FX budget rates** (£ per 1 foreign unit)")
    eur_rate = st.number_input("EUR → GBP", value=0.86, min_value=0.50, max_value=1.50, step=0.01, format="%.3f")
    usd_rate = st.number_input("USD → GBP", value=0.76, min_value=0.30, max_value=1.20, step=0.01, format="%.3f")
    cad_rate = st.number_input("CAD → GBP", value=0.53, min_value=0.20, max_value=1.00, step=0.01, format="%.3f")
    fx_rates = {"EUR": eur_rate, "USD": usd_rate, "CAD": cad_rate, "GBP": 1.0}

    st.divider()
    st.markdown("**Actual cash positions (GBP equiv)**")
    pos_date = st.date_input("As at date", value=None)
    pos_sht = st.number_input("SHT (£)", value=0, min_value=0, max_value=999_999_999, step=1000, format="%d")
    pos_shgi = st.number_input("SHGI (£)", value=0, min_value=0, max_value=999_999_999, step=1000, format="%d")
    pos_tmd = st.number_input("TMD (£)", value=0, min_value=0, max_value=999_999_999, step=1000, format="%d")

    pos_total_uk = pos_sht + pos_tmd
    pos_total_ireland = pos_shgi
    pos_total = pos_sht + pos_shgi + pos_tmd
    use_actual_pos = pos_total > 0
    if use_actual_pos:
        st.success(f"Total: £{pos_total:,.0f} · UK: £{pos_total_uk:,.0f} · IE: £{pos_total_ireland:,.0f}")
    else:
        st.caption("Enter values above to use actual position.")

    st.divider()
    st.markdown("**Client money override**")
    cm_override = st.number_input("Client money (£) — leave 0 to use forecast file", value=0, min_value=0, max_value=999_999_999, step=100_000, format="%d")
    use_cm_override = cm_override > 0

    st.divider()
    st.markdown("**Remaining flows adjustment**")
    st.caption("Adjust forecast receipts / AP versus LY run rate.")
    adj_receipts = st.slider("Receipts vs LY remaining (%)", -50, 50, 0, 5)
    adj_payments = st.slider("AP payments vs LY remaining (%)", -50, 50, 0, 5)

if not tx_file or not fc_file:
    st.title("💷 SHG Cash Flow Dashboard")
    c1, c2 = st.columns(2)
    c1.info("📂 Upload **bank transactions** Excel in the sidebar. `Data Sheet` tab required.")
    c2.info("📂 Upload **forecast / client money** Excel in the sidebar.")
    st.stop()

# ── Load and enrich ───────────────────────────────────────────────────────────
with st.spinner("Loading data..."):
    df_raw, df_op = load_transactions(tx_file.read())
    fc_raw = load_forecast(fc_file.read())
    entity_m, uk_in_spec, uk_out_spec, ie_in_spec, ie_out_spec = build_actuals(df_op)

    latest_date = df_raw["PostDate"].max()
    fc_raw = extend_forecast_horizon(fc_raw, latest_date)

    weekly_raw, od_weekly = build_weekly(df_raw, fx_rates)

    hard_min_w = pd.Timestamp("2027-12-31")
    rolling_min_w = latest_date + pd.Timedelta(weeks=52)
    required_w = max(hard_min_w, rolling_min_w)
    last_act_wk = weekly_raw.index.max()
    n_fc_dynamic = max(13, int((required_w - last_act_wk).days // 7) + 1)
    fc_base, fc_weeks = build_weekly_forecast(weekly_raw, od_weekly, n_fc=n_fc_dynamic)

fc = fc_raw.copy()
fc["uk_required"] = fc["client_money"] * uk_pct / 100
fc["uk_headroom"] = fc["cash_uk"] - fc["uk_required"]
fc["total_cash"] = fc["cash_uk"] + fc["cash_ireland"]
fc["total_headroom"] = fc["total_cash"] - fc["uk_required"]
fc["headroom_pct"] = fc["uk_headroom"] / fc["uk_required"]
fc["total_headroom_pct"] = fc["total_headroom"] / fc["uk_required"]
fc["fc_implied_net_total"] = fc["total_cash"].diff()
fc["fc_implied_net_uk"] = fc["cash_uk"].diff()

latest_yr = int(latest_date.year)
latest_mn = int(latest_date.month)
latest_day = int(latest_date.day)

def is_actual(yr, mn):
    return yr < latest_yr or (yr == latest_yr and mn <= latest_mn)

fc["is_actual"] = fc.apply(lambda r: is_actual(int(r["Year"]), int(r["Month"])), axis=1)
fc["is_partial"] = fc.apply(lambda r: int(r["Year"]) == latest_yr and int(r["Month"]) == latest_mn, axis=1)

latest_fc = fc[fc["is_actual"]].iloc[-1]
kpi_client_mon = float(cm_override) if use_cm_override else float(latest_fc["client_money"])
req_now = kpi_client_mon * uk_pct / 100

# KPI month-end forecast, with actual-position projection if entered
fc_me_uk = float(latest_fc["cash_uk"])
fc_me_ie = float(latest_fc["cash_ireland"])
fc_me_total = float(latest_fc["total_cash"])

if use_actual_pos and pos_date is not None:
    days_in_month = calendar.monthrange(pos_date.year, pos_date.month)[1]
    days_remaining = max(0, days_in_month - pos_date.day)
    weeks_remaining = days_remaining / 7.0

    core_in = ["AGENT RECEIPTS", "FD RECEIPT", "DIRECT RECEIPTS", "OTHER RECEIPTS", "TUI RECEIPT"]
    core_out = ["AP COGS", "AP OVH", "FLIGHT COSTS", "PAYROLL", "OTHER COSTS", "TAX"]

    df_op_w = df_op.copy()
    df_op_w["WeekStart"] = df_op_w["PostDate"].dt.to_period("W-SUN").apply(lambda p: p.start_time)
    wk_core_in = df_op_w[df_op_w["TrnSpec"].isin(core_in)].groupby("WeekStart")["inflow"].sum()
    wk_core_out = df_op_w[df_op_w["TrnSpec"].isin(core_out)].groupby("WeekStart")["outflow"].sum()
    avg_wk_in = float(wk_core_in.tail(4).mean()) if len(wk_core_in) else 0.0
    avg_wk_out = float(wk_core_out.tail(4).mean()) if len(wk_core_out) else 0.0

    rem_in = avg_wk_in * weeks_remaining * (1 + adj_receipts / 100)
    rem_out = avg_wk_out * weeks_remaining * (1 + adj_payments / 100)

    kpi_total_cash = pos_total + rem_in + rem_out
    # Keep UK/Ireland shown separately from forecast split for compliance context
    kpi_uk_cash = pos_total_uk + (rem_in + rem_out) * (fc_me_uk / fc_me_total if fc_me_total else 0)
    kpi_ie_cash = kpi_total_cash - kpi_uk_cash
    kpi_subtitle = f"Forecast {pd.Timestamp(latest_yr, latest_mn, days_in_month).strftime('%d %b')} · {days_remaining}d remaining · total cash basis"
else:
    kpi_uk_cash = fc_me_uk
    kpi_ie_cash = fc_me_ie
    kpi_total_cash = fc_me_total
    kpi_subtitle = f"Forecast {MN[latest_mn]} {latest_yr} month-end"

hroom_now = kpi_uk_cash - req_now
hpct_now = hroom_now / req_now if req_now > 0 else 0

total_hroom_now = kpi_total_cash - req_now
total_hpct_now = total_hroom_now / req_now if req_now > 0 else 0

# ── Header / KPIs ─────────────────────────────────────────────────────────────
st.title(f"💷 SHG Cash — Latest: {latest_date.strftime('%d %b %Y')}")
if use_actual_pos:
    pos_date_str = pos_date.strftime("%d %b %Y") if pos_date else "date not set"
    st.info(
        f"📍 **Actual cash position as at {pos_date_str}** — "
        f"SHT: £{pos_sht:,.0f} · SHGI: £{pos_shgi:,.0f} · TMD: £{pos_tmd:,.0f} · "
        f"**Total: £{pos_total:,.0f}** · UK: £{pos_total_uk:,.0f} · Ireland: £{pos_total_ireland:,.0f}"
    )

if hpct_now >= 0.20:
    st.success(f"✅ COMPLIANT — UK headroom {fmt(hroom_now)} ({hpct_now:.0%} of req {fmt(req_now)}) · {kpi_subtitle}")
elif hpct_now >= 0:
    st.warning(f"⚠️ CAUTION — UK headroom {fmt(hroom_now)} ({hpct_now:.0%}) · {kpi_subtitle}")
else:
    st.error(f"🚨 BREACH — UK cash {fmt(abs(hroom_now))} below requirement · {kpi_subtitle}")

k = st.columns(6)
k[0].metric("UK Cash (fcst m/e)", fmt(kpi_uk_cash), delta=f"actual now: {fmt(pos_total_uk)}" if use_actual_pos else None)
k[1].metric("Ireland Cash (fcst m/e)", fmt(kpi_ie_cash), delta=f"actual now: {fmt(pos_total_ireland)}" if use_actual_pos else None)
k[2].metric("Total Cash (fcst m/e)", fmt(kpi_total_cash), delta=f"actual now: {fmt(pos_total)}" if use_actual_pos else None)
k[3].metric("Client Money", fmt(kpi_client_mon), delta="overridden" if use_cm_override else None, delta_color="off")
k[4].metric("UK Required", fmt(req_now), delta=f"{uk_pct}% of {'override' if use_cm_override else 'forecast file'}", delta_color="off")
k[5].metric("Total Headroom", fmt(total_hroom_now), delta=f"{total_hpct_now:.0%}", delta_color="normal" if total_hroom_now >= 0 else "inverse")

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SHARED WEEKLY 4+13 DATA — TOTAL CASH BASIS
# ══════════════════════════════════════════════════════════════════════════════
if "weekly_ov" not in st.session_state:
    st.session_state.weekly_ov = {}

n_actuals = 4
actual_weeks = weekly_raw.index[-n_actuals:].tolist()
all_weeks = actual_weeks + fc_weeks
N_W = len(all_weeks)

first_wk = actual_weeks[0]
book2_before = fc[fc.index < first_wk]
book2_open = float(book2_before.iloc[-1]["total_cash"]) / 1000 if not book2_before.empty else 0.0
open_base = pos_total / 1000 if use_actual_pos else book2_open

ROW_SPECS = [
    ("DIRECT RECEIPTS", "DIRECT RECEIPTS", False),
    ("AGENT RECEIPTS", "AGENT RECEIPTS", False),
    ("FD RECEIPT", "FD RECEIPT", False),
    ("CUSTOMER REFUND", "CUSTOMER REFUNDS", False),
    ("TUI RECEIPT", "TUI RECEIPT", False),
    ("OTHER RECEIPT", "OTHER RECEIPTS", True),
    ("FX TRADE IN", "FX TRADE IN", True),
    ("INTERCO (net)", "INTERCO", True),
    ("OD INTEREST", "_OD", False),
    ("AP COGS", "AP COGS", False),
    ("AP OVH", "AP OVH", False),
    ("PAYROLL", "PAYROLL", False),
    ("TAX", "TAX", False),
    ("OTHER CASH OUT", "OTHER COSTS", True),
    ("FX TRADE OUT", "FX TRADE OUT", True),
    ("FLIGHT COSTS", "FLIGHT COSTS", False),
]

SLIDER_RECEIPT_ROWS = {"DIRECT RECEIPTS", "AGENT RECEIPTS", "FD RECEIPT", "TUI RECEIPT"}
SLIDER_AP_ROWS = {"AP COGS", "AP OVH"}

data = {}
for label, spec, blank_fc in ROW_SPECS:
    row = []
    for i, wk in enumerate(all_weeks):
        ov_key = f"{label}_{i}"
        if ov_key in st.session_state.weekly_ov:
            row.append(st.session_state.weekly_ov[ov_key] / 1000)
            continue

        is_fc_wk = i >= n_actuals
        if not is_fc_wk:
            if spec == "_OD":
                v = float(od_weekly.get(wk, 0)) / 1000
            elif spec in weekly_raw.columns and wk in weekly_raw.index:
                v = float(weekly_raw.loc[wk, spec]) / 1000
            else:
                v = 0.0
        else:
            fi = i - n_actuals
            if blank_fc:
                v = 0.0
            elif spec == "_OD":
                v = 11.0
            elif spec == "PAYROLL":
                v = fc_base.get("PAYROLL", [0] * len(fc_weeks))[fi] / 1000
            else:
                v = fc_base.get(spec, [0] * len(fc_weeks))[fi] / 1000 if spec else 0.0

            if label in SLIDER_RECEIPT_ROWS and adj_receipts != 0:
                v *= (1 + adj_receipts / 100)
            if label in SLIDER_AP_ROWS and adj_payments != 0:
                v *= (1 + adj_payments / 100)

        row.append(round(v, 1))
    data[label] = row

totR = [sum(data[l][i] for l in RECEIPT_LABELS) for i in range(N_W)]
totP = [sum(data[l][i] for l in PAYMENT_LABELS) for i in range(N_W)]
net = [totR[i] + totP[i] for i in range(N_W)]
opens = [0.0] * N_W
closes = [0.0] * N_W
opens[0] = open_base
closes[0] = open_base + net[0]
for i in range(1, N_W):
    opens[i] = closes[i-1]
    closes[i] = closes[i-1] + net[i]


def shared_month_end_close(yr, mn):
    """Return weekly-outlook TOTAL closing balance (£) at month end."""
    wks = [(i, w) for i, w in enumerate(all_weeks) if w.year == yr and w.month == mn]
    if not wks:
        return None
    last_i = max(wks, key=lambda x: x[0])[0]
    return closes[last_i] * 1000

# ── Tabs ─────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "📈 Cash vs Forecast",
    "🛡️ Compliance",
    "📊 UK YoY",
    "📊 Ireland YoY",
    "🎯 3-Month Focus",
    "⚡ Opportunities & Risks",
    "📅 Weekly 4+13 Outlook",
    "ℹ️ How it works",
])

fc_chart = fc.copy()
fc_chart["label"] = fc_chart["Month"].map(MN) + " " + fc_chart["Year"].astype(str)
labels = fc_chart["label"].tolist()

def split_act_fc(series):
    act = series.where(fc_chart["is_actual"] & ~fc_chart["is_partial"])
    fc_s = series.where(~fc_chart["is_actual"] | fc_chart["is_partial"])
    idx = fc_chart[fc_chart["is_actual"]].index[-1]
    if idx in fc_s.index:
        fc_s.loc[idx] = series.loc[idx]
    return act.tolist(), fc_s.tolist()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Cash vs Forecast
# ══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    fig = go.Figure()
    for series, name, color, width in [
        (fc["cash_uk"], "UK cash", BLUE, 2.5),
        (fc["total_cash"], "Total cash", GREEN, 2.5),
        (fc["uk_headroom"], "UK headroom", LBLUE, 1.5),
        (fc["total_headroom"], "Total headroom", LGREEN, 1.5),
    ]:
        act, fcast = split_act_fc(series / 1e6)
        fig.add_trace(go.Scatter(x=labels, y=act, name=f"{name} (actual)", line=dict(color=color, width=width), mode="lines+markers", marker=dict(size=3)))
        fig.add_trace(go.Scatter(x=labels, y=fcast, name=f"{name} (forecast)", line=dict(color=color, width=max(1, width-0.5), dash="dot"), mode="lines+markers", marker=dict(size=2, symbol="circle-open")))
    fig.add_trace(go.Scatter(x=labels, y=(fc["uk_required"] / 1e6).tolist(), name="UK required", line=dict(color=RED, dash="dash", width=1.5), mode="lines"))
    fig.add_hline(y=0, line_color="rgba(0,0,0,0.15)", line_width=0.5)
    fig.update_layout(height=380, plot_bgcolor="white", paper_bgcolor="white", yaxis=dict(tickformat=",.1f", ticksuffix="m", gridcolor=GREY, title="£m"), xaxis=dict(tickangle=45), legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(size=10)), margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig, use_container_width=True)

    rows = []
    for _, r in fc.iterrows():
        hp = r["headroom_pct"]
        thp = r["total_headroom_pct"]
        rows.append({
            "Year": int(r["Year"]),
            "Month": MN[int(r["Month"])],
            "Type": "Actual" if r["is_actual"] else "Forecast",
            "Client Money": fmt(r["client_money"]),
            "UK Cash": fmt(r["cash_uk"]),
            "Ireland Cash": fmt(r["cash_ireland"]),
            "Total Cash": fmt(r["total_cash"]),
            "UK Required": fmt(r["uk_required"]),
            "UK Headroom": fmt(r["uk_headroom"]),
            "Total Headroom": fmt(r["total_headroom"]),
            "UK Headroom %": f"{hp:.0%}" if not np.isnan(hp) else "—",
            "Total Headroom %": f"{thp:.0%}" if not np.isnan(thp) else "—",
            "Status": "✅ Compliant" if hp >= 0.20 else ("⚠️ Caution" if hp >= 0 else "🚨 Breach"),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Compliance
# ══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    breaches = fc[(fc["uk_headroom"] < 0) & (~fc["is_actual"])]
    cautions = fc[(fc["uk_headroom"] >= 0) & (fc["headroom_pct"] < 0.20) & (~fc["is_actual"])]
    if not breaches.empty:
        st.error("🚨 **Forecast UK breach months:** " + ", ".join(f"{MN[int(r['Month'])]} {int(r['Year'])} ({fmt(r['uk_headroom'])})" for _, r in breaches.iterrows()))
    if not cautions.empty:
        st.warning("⚠️ **UK caution months:** " + ", ".join(f"{MN[int(r['Month'])]} {int(r['Year'])} ({r['headroom_pct']:.0%})" for _, r in cautions.iterrows()))

    fig2 = make_subplots(rows=2, cols=1, subplot_titles=["UK headroom (£m)", "Total headroom (£m)"], vertical_spacing=0.14)
    for series, ri, col in [(fc["uk_headroom"], 1, BLUE), (fc["total_headroom"], 2, GREEN)]:
        act, fcast = split_act_fc(series / 1e6)
        rgb = tuple(int(col[i:i+2], 16) for i in (1, 3, 5))
        fig2.add_trace(go.Scatter(x=labels, y=act, line=dict(color=col, width=2), fill="tozeroy", fillcolor=f"rgba({rgb[0]},{rgb[1]},{rgb[2]},0.08)", name="Actual"), row=ri, col=1)
        fig2.add_trace(go.Scatter(x=labels, y=fcast, line=dict(color=col, width=1.5, dash="dot"), name="Forecast"), row=ri, col=1)
        fig2.add_hline(y=0, line_color=RED, line_dash="dash", line_width=1.5, row=ri, col=1)
    fig2.update_layout(height=480, plot_bgcolor="white", paper_bgcolor="white", showlegend=False, margin=dict(l=0, r=0, t=40, b=0))
    fig2.update_yaxes(tickformat=",.1f", ticksuffix="m", gridcolor=GREY)
    st.plotly_chart(fig2, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# YoY helper
# ══════════════════════════════════════════════════════════════════════════════
def build_yoy_table(in_spec, out_spec, cats, is_out, entity_label):
    spec = out_spec if is_out else in_spec
    if spec.empty:
        st.info(f"No {entity_label} data available.")
        return
    all_ym = sorted(set(spec.index.tolist()))
    months_avail = sorted(set(mn for _, mn in all_ym))
    years_avail = sorted(set(yr for yr, _ in all_ym))
    rows = []
    for cat in cats:
        if cat not in spec.columns:
            continue
        row = {"Category": cat}
        for mn in months_avail:
            for yr in years_avail:
                v = safe_get(spec, yr, mn, cat)
                if is_out:
                    v = abs(v)
                row[f"{MN[mn]} {yr}"] = v
        rows.append(row)
    if not rows:
        st.info(f"No {entity_label} {'outflow' if is_out else 'inflow'} categories available.")
        return
    df_s = pd.DataFrame(rows).fillna(0)
    for col in df_s.columns[1:]:
        df_s[col] = df_s[col].apply(lambda v: fmt(v) if v != 0 else "—")
    st.dataframe(df_s, use_container_width=True, hide_index=True)

    fig = go.Figure()
    for yi, yr in enumerate(years_avail):
        vals = []
        for mn in months_avail:
            total = sum(abs(safe_get(spec, yr, mn, c)) if is_out else safe_get(spec, yr, mn, c) for c in cats if c in spec.columns)
            vals.append(total / 1e6)
        clr = ["rgba(136,135,128,0.55)", "rgba(28,20,100,0.7)", "rgba(59,109,17,0.7)", "rgba(201,168,76,0.7)"][yi % 4]
        fig.add_trace(go.Bar(x=[MN[m] for m in months_avail], y=vals, name=str(yr), marker_color=clr, offsetgroup=yi))
    fig.update_layout(barmode="group", height=240, title=f"{entity_label} {'outflows' if is_out else 'inflows'} by month (£m)", plot_bgcolor="white", paper_bgcolor="white", yaxis=dict(tickformat=",.1f", ticksuffix="m", gridcolor=GREY), legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(size=10)), margin=dict(l=0, r=0, t=40, b=0))
    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 / 4 — YoY
# ══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    uk_sub = st.tabs(["📥 UK Inflow YoY", "📤 UK Outflow YoY"])
    with uk_sub[0]:
        avail_in = [c for c in KEY_INFLOW if c in uk_in_spec.columns and c not in {"FX TRADE IN", "FX TRADE OUT"}]
        st.caption("Like-for-like: same calendar month year-on-year. FX trade excluded.")
        build_yoy_table(uk_in_spec, uk_out_spec, avail_in, False, "UK")
    with uk_sub[1]:
        avail_out = [c for c in KEY_OUTFLOW if c in uk_out_spec.columns and c not in {"FX TRADE IN", "FX TRADE OUT"}]
        st.caption("Outflows shown as positive. FX trade excluded.")
        build_yoy_table(uk_in_spec, uk_out_spec, avail_out, True, "UK")

with tabs[3]:
    ie_sub = st.tabs(["📥 Ireland Inflow YoY", "📤 Ireland Outflow YoY"])
    with ie_sub[0]:
        ie_in_cats = [c for c in KEY_INFLOW if c in ie_in_spec.columns and c not in {"FX TRADE IN", "FX TRADE OUT"}]
        st.caption("Ireland inflows. FX trade excluded.")
        build_yoy_table(ie_in_spec, ie_out_spec, ie_in_cats, False, "Ireland")
    with ie_sub[1]:
        ie_out_cats = [c for c in KEY_OUTFLOW if c in ie_out_spec.columns and c not in {"FX TRADE IN", "FX TRADE OUT"}]
        st.caption("Ireland outflows shown as positive. FX trade excluded.")
        build_yoy_table(ie_in_spec, ie_out_spec, ie_out_cats, True, "Ireland")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — 3-Month Focus — TOTAL CASH BASIS
# ══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    focus_months = []
    for delta in range(3):
        mn = latest_mn + delta
        yr = latest_yr
        while mn > 12:
            mn -= 12
            yr += 1
        focus_months.append((yr, mn))

    st.caption(
        f"Bank data to **{latest_date.strftime('%d %b %Y')}**. Current month + next 2. "
        f"This tab now uses **total cash** and the same close chain as the Weekly 4+13 Outlook."
    )

    cols = st.columns(3)
    for ci, (fyr, fmn) in enumerate(focus_months):
        fc_row_df = fc[(fc["Year"] == fyr) & (fc["Month"] == fmn)]
        if fc_row_df.empty:
            with cols[ci]:
                st.info(f"{MN[fmn]} {fyr} — no forecast data")
            continue

        fc_row = fc_row_df.iloc[0]
        is_part = fyr == latest_yr and fmn == latest_mn

        # Opening balance — TOTAL CASH
        prev_mn_3, prev_yr_3 = (fmn - 1, fyr) if fmn > 1 else (12, fyr - 1)
        prev_fc = fc[(fc["Year"] == prev_yr_3) & (fc["Month"] == prev_mn_3)]
        book2_open = float(prev_fc.iloc[0]["total_cash"]) if not prev_fc.empty else float(fc_row["total_cash"])
        if is_part and use_actual_pos:
            open_cash = pos_total
        else:
            wk_open = shared_month_end_close(prev_yr_3, prev_mn_3)
            open_cash = wk_open if wk_open is not None else book2_open

        fc_uk_close = float(fc_row["cash_uk"])
        fc_ie_close = float(fc_row["cash_ireland"])
        fc_close = float(fc_row["total_cash"])
        uk_req = float(fc_row["uk_required"])
        headroom = fc_close - uk_req
        hpct = headroom / uk_req if uk_req > 0 else 0

        wk_close_this = shared_month_end_close(fyr, fmn)
        wk_headroom = wk_close_this - uk_req if wk_close_this is not None else None

        def total_in(yr, mn, cat):
            return safe_get(uk_in_spec, yr, mn, cat) + safe_get(ie_in_spec, yr, mn, cat)

        def total_out_abs(yr, mn, cat):
            return abs(safe_get(uk_out_spec, yr, mn, cat) + safe_get(ie_out_spec, yr, mn, cat))

        INFLOW_CORE = ["AGENT RECEIPTS", "FD RECEIPT", "DIRECT RECEIPTS"]
        ly_inflow = sum(total_in(fyr - 1, fmn, c) for c in INFLOW_CORE)
        ly_apcogs = total_out_abs(fyr - 1, fmn, "AP COGS")
        ly_apovh = total_out_abs(fyr - 1, fmn, "AP OVH")
        ly_ap_total = ly_apcogs + ly_apovh
        ly_flight = total_out_abs(fyr - 1, fmn, "FLIGHT COSTS")
        ly_payroll = total_out_abs(fyr - 1, fmn, "PAYROLL")
        ly_tax = total_out_abs(fyr - 1, fmn, "TAX")
        ly_fixed = ly_flight + ly_payroll + ly_tax

        days_in = (pd.Timestamp(fyr, fmn, 1) + pd.offsets.MonthEnd(1)).day
        weeks_in = days_in / 7.0

        if is_part:
            rem = max(0, days_in - latest_day)
            frac_rem = rem / days_in
            weeks_rem = rem / 7.0
            rem_in = ly_inflow * frac_rem * (1 + adj_receipts / 100)
            rem_fixed = ly_fixed * frac_rem
            ap_mtd = total_out_abs(fyr, fmn, "AP COGS") + total_out_abs(fyr, fmn, "AP OVH")
            total_ap_budget = open_cash + rem_in - fc_close - rem_fixed
            allow_ap_remain = max(0, total_ap_budget - ap_mtd)
            allow_ap_per_wk = allow_ap_remain / max(weeks_rem, 0.5)
        else:
            rem = days_in
            frac_rem = 1.0
            weeks_rem = weeks_in
            rem_in = ly_inflow * (1 + adj_receipts / 100)
            rem_fixed = ly_fixed
            ap_mtd = 0.0
            total_ap_budget = open_cash + rem_in - fc_close - rem_fixed
            allow_ap_remain = max(0, total_ap_budget)
            allow_ap_per_wk = allow_ap_remain / weeks_in

        ly_ap_for_period_adj = ly_ap_total * frac_rem * (1 + adj_payments / 100)
        cy_inflow_mtd = sum(total_in(fyr, fmn, c) for c in INFLOW_CORE)
        hold_vs_ly = max(0, ly_ap_for_period_adj - allow_ap_remain)
        ly_ap_per_wk = ly_ap_total / weeks_in if weeks_in else 0

        badge = "🚨 Breach" if hpct < 0 else ("⚠️ Monitor" if hpct < 0.20 else "✅ On track")

        with cols[ci]:
            st.markdown(f"### {MN[fmn]} {fyr} {badge}")
            if is_part:
                st.caption(f"**{latest_day} of {days_in} days elapsed · {rem} days / ~{weeks_rem:.1f} weeks remaining**")

            st.markdown("**Cash position**")
            wk_cl_fmt = fmt(wk_close_this) if wk_close_this is not None else "—"
            wk_hr_fmt = fmt(wk_headroom) if wk_headroom is not None else "—"
            wk_hr_pct = f"({wk_headroom / uk_req:.0%})" if wk_headroom is not None and uk_req > 0 else ""
            var_close = wk_close_this - fc_close if wk_close_this is not None else None
            var_fmt = f"**{'+' if var_close >= 0 else ''}{fmt(var_close)}**" if var_close is not None else "—"

            st.markdown(
                "| | Forecast target | Weekly outlook |\n"
                "|---|---|---|\n"
                f"| Opening total cash | {fmt(book2_open)} | **{fmt(open_cash)}** |\n"
                f"| UK close | {fmt(fc_uk_close)} | — |\n"
                f"| Ireland close | {fmt(fc_ie_close)} | — |\n"
                f"| **Total close** | **{fmt(fc_close)}** | **{wk_cl_fmt}** |\n"
                f"| UK required | {fmt(uk_req)} | {fmt(uk_req)} |\n"
                f"| Total headroom | **{fmt(headroom)}** ({hpct:.0%}) | **{wk_hr_fmt}** {wk_hr_pct} |\n"
                f"| Variance vs fcst total close | | {var_fmt} |"
            )

            if var_close is not None and var_close < -500000:
                st.error(f"⚠️ Weekly outlook {fmt(abs(var_close))} **below** total forecast target — adjust AP or receipts")
            elif var_close is not None and var_close > 500000:
                st.success(f"✅ Weekly outlook {fmt(var_close)} **ahead** of total forecast target")

            st.markdown("**Expected receipts** (Direct + Agent + FD, LY basis · total group)")
            st.markdown(
                "| | LY full month | " + ("Remaining est." if is_part else "Full month") + " |\n"
                "|---|---|---|\n"
                f"| Core receipts | {fmt(ly_inflow)} | {fmt(rem_in)} |"
                + (f"\n| CY actual to date | | {fmt(cy_inflow_mtd)} |" if is_part else "")
            )

            st.markdown("---")
            st.markdown("### 🔑 AP payment capacity")
            if hold_vs_ly > 50000:
                st.error(f"**Hold back {fmt(hold_vs_ly)} vs LY run rate** to hit forecast close")
            else:
                st.success("✅ AP can run at LY pace — no hold back required")

            m1, m2 = st.columns(2)
            adj_note = ""
            if adj_receipts != 0: adj_note += f" · receipts {adj_receipts:+d}%"
            if adj_payments != 0: adj_note += f" · AP {adj_payments:+d}%"
            m1.metric("Total AP budget " + ("remaining" if is_part else "this month"), fmt(allow_ap_remain), delta=f"LY rate: {fmt(ly_ap_for_period_adj)}{adj_note}", delta_color="off")
            m2.metric("Weekly AP capacity", fmt(allow_ap_per_wk), delta=f"LY weekly avg: {fmt(ly_ap_per_wk)}", delta_color="off")

            if is_part and ap_mtd > 0:
                st.caption(f"AP COGS + OVH already paid this month: **{fmt(ap_mtd)}** · Remaining capacity: {fmt(allow_ap_remain)}")

            if ly_ap_total > 0:
                cogs_frac = ly_apcogs / ly_ap_total
                ovh_frac = ly_apovh / ly_ap_total
                st.markdown(
                    "| | LY run rate | Budget |\n"
                    "|---|---|---|\n"
                    f"| AP COGS | {fmt(ly_apcogs * frac_rem)} | **{fmt(allow_ap_remain * cogs_frac)}** |\n"
                    f"| AP OVH | {fmt(ly_apovh * frac_rem)} | **{fmt(allow_ap_remain * ovh_frac)}** |"
                )

            with st.expander("Other outflows for reference"):
                est_label = "Remaining est." if is_part else "Full month"
                st.markdown(
                    f"| Category | LY full month | {est_label} |\n"
                    f"|---|---|---|\n"
                    f"| Flight costs | {fmt(ly_flight)} | {fmt(ly_flight * frac_rem)} |\n"
                    f"| Payroll | {fmt(ly_payroll)} | {fmt(ly_payroll * frac_rem)} |\n"
                    f"| Tax | {fmt(ly_tax)} | {fmt(ly_tax * frac_rem)} |"
                )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — Opportunities & Risks — TOTAL GROUP
# ══════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    st.caption("Total group cash flows (UK + Ireland). Actuals vs total forecast-implied and prior year same month.")
    act_data = fc[fc["is_actual"] & ~fc["is_partial"]]

    monthly_data = []
    for _, row in act_data.iterrows():
        yr, mn = int(row["Year"]), int(row["Month"])
        cy_net = (
            safe_entity(entity_m, yr, mn, "UK", "inflow") + safe_entity(entity_m, yr, mn, "UK", "outflow") +
            safe_entity(entity_m, yr, mn, "Ireland", "inflow") + safe_entity(entity_m, yr, mn, "Ireland", "outflow")
        )
        ly_net = (
            safe_entity(entity_m, yr-1, mn, "UK", "inflow") + safe_entity(entity_m, yr-1, mn, "UK", "outflow") +
            safe_entity(entity_m, yr-1, mn, "Ireland", "inflow") + safe_entity(entity_m, yr-1, mn, "Ireland", "outflow")
        )
        fc_imp = float(row.get("fc_implied_net_total", 0) or 0)
        monthly_data.append({"label": f"{MN[mn]} {yr}", "cy_net": cy_net, "ly_net": ly_net, "fc_imp": fc_imp})

    all_labels_c, all_act_c, all_fc_c, all_ly_c = [], [], [], []
    for _, row in fc.iterrows():
        yr, mn = int(row["Year"]), int(row["Month"])
        is_act_complete = bool(row["is_actual"]) and not bool(row["is_partial"])
        ly_net = (
            safe_entity(entity_m, yr-1, mn, "UK", "inflow") + safe_entity(entity_m, yr-1, mn, "UK", "outflow") +
            safe_entity(entity_m, yr-1, mn, "Ireland", "inflow") + safe_entity(entity_m, yr-1, mn, "Ireland", "outflow")
        ) / 1e6
        act_net = None
        if is_act_complete:
            act_net = (
                safe_entity(entity_m, yr, mn, "UK", "inflow") + safe_entity(entity_m, yr, mn, "UK", "outflow") +
                safe_entity(entity_m, yr, mn, "Ireland", "inflow") + safe_entity(entity_m, yr, mn, "Ireland", "outflow")
            ) / 1e6
        all_labels_c.append(f"{MN[mn]} {yr}")
        all_act_c.append(act_net)
        all_fc_c.append(float(row.get("fc_implied_net_total", 0) or 0) / 1e6)
        all_ly_c.append(ly_net if abs(ly_net) > 0.1 else None)

    st.subheader("Monthly net cash — total group")
    fig_or = go.Figure()
    bar_vals = [all_act_c[i] if all_act_c[i] is not None else all_fc_c[i] for i in range(len(all_labels_c))]
    bar_colors = [GREEN if (v or 0) >= 0 else RED for v in bar_vals]
    fig_or.add_trace(go.Bar(x=all_labels_c, y=bar_vals, name="Actual / Forecast", marker_color=bar_colors, opacity=0.85))
    fig_or.add_trace(go.Scatter(x=all_labels_c, y=all_ly_c, name="Prior year same month", line=dict(color=AMBER, width=2, dash="dot"), mode="lines+markers", connectgaps=False))
    fig_or.add_hline(y=0, line_color="rgba(0,0,0,0.2)", line_width=0.8)
    fig_or.update_layout(height=320, plot_bgcolor="white", paper_bgcolor="white", yaxis=dict(tickformat=",.1f", ticksuffix="m", gridcolor=GREY, title="£m"), xaxis=dict(tickangle=45), legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(size=10)), margin=dict(l=0, r=0, t=40, b=0))
    st.plotly_chart(fig_or, use_container_width=True)

    if monthly_data:
        trend_vs_ly = np.mean([d["cy_net"] - d["ly_net"] for d in monthly_data])
        trend_vs_fc = np.mean([d["cy_net"] - d["fc_imp"] for d in monthly_data if abs(d["fc_imp"]) > 100000]) if any(abs(d["fc_imp"]) > 100000 for d in monthly_data) else 0.0
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### Opportunities")
            if trend_vs_ly > 0.2e6:
                st.success(f"📈 Running **{fmt(trend_vs_ly)}/month ahead** of prior year on average.")
            if trend_vs_fc > 0.2e6:
                st.success(f"📈 Running **{fmt(trend_vs_fc)}/month ahead** of forecast file on average.")
            if trend_vs_ly <= 0.2e6 and trend_vs_fc <= 0.2e6:
                st.info("No material positive trend detected.")
        with c2:
            st.markdown("### Risks")
            if trend_vs_ly < -0.2e6:
                st.error(f"📉 Running **{fmt(abs(trend_vs_ly))}/month below** prior year on average.")
            if trend_vs_fc < -0.2e6:
                st.error(f"📉 Running **{fmt(abs(trend_vs_fc))}/month below** forecast file.")
            if trend_vs_ly >= -0.2e6 and trend_vs_fc >= -0.2e6:
                st.success("No material negative trend detected.")

    buf = io.BytesIO()
    fc.reset_index().to_excel(buf, index=False, sheet_name="Dashboard Data")
    buf.seek(0)
    st.download_button("⬇️ Download dashboard data (.xlsx)", data=buf, file_name="shg_cash_dashboard.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — Weekly 4+13 Outlook — TOTAL CASH BASIS
# ══════════════════════════════════════════════════════════════════════════════
with tabs[6]:
    st.caption(
        f"All amounts £GBP equivalent at budget rates (EUR×{eur_rate:.3f} · USD×{usd_rate:.3f} · CAD×{cad_rate:.3f}). "
        f"Actual: last 4 weeks from bank data. Forecast: prior year same ISO week. "
        f"Basis: **total group cash** across UK + Ireland. Interco nets naturally at total-group level."
    )

    if use_actual_pos:
        st.info(
            f"📍 Opening balance anchored to actual total cash position: "
            f"SHT £{pos_sht:,.0f} + SHGI £{pos_shgi:,.0f} + TMD £{pos_tmd:,.0f} = **Total £{pos_total:,.0f}** "
            f"(W{actual_weeks[0].isocalendar()[1]} {actual_weeks[0].strftime('%d %b')} opening = **£{open_base:,.0f}k**)."
        )

    _, col_btn = st.columns([8, 1])
    with col_btn:
        if st.button("Reset overrides"):
            st.session_state.weekly_ov = {}
            st.rerun()

    ovr_labels = ["AP COGS", "AP OVH", "PAYROLL", "FD RECEIPT", "AGENT RECEIPTS", "FX TRADE IN", "FX TRADE OUT", "FLIGHT COSTS", "INTERCO (net)", "OTHER RECEIPT", "OTHER CASH OUT"]
    spec_map = {l: s for l, s, _ in ROW_SPECS}

    with st.expander("Edit forecast — nearest 4 weeks (values in £)", expanded=False):
        st.caption("Overrides feed directly into this 4+13 table and the 3-Month Focus tab.")
        cols_ov = st.columns(4)
        for fi in range(min(4, len(fc_weeks))):
            i = fi + n_actuals
            wk = all_weeks[i]
            with cols_ov[fi]:
                st.markdown(f"**{wk.strftime('%d %b %Y')}**")
                for label in ovr_labels:
                    spec = spec_map.get(label, label)
                    ov_key = f"{label}_{i}"
                    is_lk = label in WEEKLY_LOCK
                    is_ico = "INTERCO" in label
                    spec_blank = next((b for l2, s2, b in ROW_SPECS if l2 == label), False)
                    if spec_blank:
                        fc_def = 0.0
                    elif is_lk:
                        fc_def = float(fc_base.get("PAYROLL", [0] * len(fc_weeks))[fi])
                    else:
                        fc_def = float(fc_base.get(spec, [0] * len(fc_weeks))[fi] if spec in fc_base else 0.0)

                    curr = float(st.session_state.weekly_ov.get(ov_key, fc_def))
                    hint = "🔒 " if is_lk else ("🔵 " if is_ico else ("⬜ " if spec_blank else ""))
                    new_v = st.number_input(
                        f"{hint}{label}",
                        value=float(np.clip(curr, -9_999_999.0, 9_999_999.0)),
                        min_value=-9_999_999.0,
                        max_value=9_999_999.0,
                        step=1000.0,
                        format="%.0f",
                        key=f"wov_{label}_{fi}",
                    )
                    if abs(new_v - fc_def) > 500:
                        st.session_state.weekly_ov[ov_key] = new_v
                    elif ov_key in st.session_state.weekly_ov:
                        del st.session_state.weekly_ov[ov_key]

    wk_hdrs = [f"{'W' if i < n_actuals else '~W'}{all_weeks[i].isocalendar()[1]}\n{all_weeks[i].strftime('%d/%m')}" for i in range(N_W)]
    display_rows = []

    def dr(label, vals, kind="data", indent=False):
        row = {"Row": ("  " + label if indent else label)}
        for i, v in enumerate(vals):
            row[wk_hdrs[i]] = fkw(v, kind not in ("total", "balance"))
        display_rows.append({"_kind": kind, "_label": label, **row})

    dr("Opening total cash balance", opens, "balance")
    display_rows.append({"_kind": "header", "_label": "RECEIPTS", "Row": "RECEIPTS", **{h: "" for h in wk_hdrs}})
    for label in RECEIPT_LABELS:
        dr(label, data[label], indent=True)
    dr("Total receipts", totR, "total")
    display_rows.append({"_kind": "sep", "_label": "", "Row": "", **{h: "" for h in wk_hdrs}})
    display_rows.append({"_kind": "header", "_label": "PAYMENTS", "Row": "PAYMENTS", **{h: "" for h in wk_hdrs}})
    for label in PAYMENT_LABELS:
        dr(label, data[label], indent=True)
    dr("Total payments", totP, "total")
    display_rows.append({"_kind": "sep", "_label": "", "Row": "", **{h: "" for h in wk_hdrs}})
    dr("Net cash", net, "total")
    dr("Closing total cash balance", closes, "balance")

    def get_book2_for_week(wk):
        """Return total forecast close, client money and total headroom in £k at month end only."""
        mn, yr = wk.month, wk.year
        wks_in_month = [w for w in all_weeks if w.month == mn and w.year == yr]
        if not wks_in_month or wk != max(wks_in_month):
            return None, None, None
        book2_row = fc[(fc["Year"] == yr) & (fc["Month"] == mn)]
        if book2_row.empty:
            return None, None, None
        row = book2_row.iloc[0]
        return float(row["total_cash"]) / 1000, float(row["client_money"]) / 1000, float(row["total_headroom"]) / 1000

    b2_fc_close, b2_cm, b2_headroom, b2_variance = [], [], [], []
    for i, wk in enumerate(all_weeks):
        fc_c, cm_c, hr_c = get_book2_for_week(wk)
        b2_fc_close.append(fc_c)
        b2_cm.append(cm_c)
        b2_headroom.append(hr_c)
        b2_variance.append(closes[i] - fc_c if fc_c is not None else None)

    def fkw_me(v):
        return "" if v is None else fkw(v, dash=False)

    def fkw_var(v):
        if v is None:
            return ""
        if abs(v) < 1:
            return "—"
        return f"{'+' if v > 0 else ''}{fkw(v, dash=False)}"

    display_rows.append({"_kind": "sep", "_label": "", "Row": "", **{h: "" for h in wk_hdrs}})
    display_rows.append({"_kind": "header", "_label": "MONTH-END TRACKING", "Row": "MONTH-END TRACKING (Forecast)", **{h: "" for h in wk_hdrs}})

    def dr_me(label, vals_fn, kind="forecast"):
        row = {"Row": label}
        for i in range(N_W):
            row[wk_hdrs[i]] = vals_fn(i)
        display_rows.append({"_kind": kind, "_label": label, **row})

    dr_me("Forecast total close", lambda i: fkw_me(b2_fc_close[i]), "forecast")
    dr_me("Client money", lambda i: fkw_me(b2_cm[i]), "forecast")
    dr_me("Total headroom vs UK req", lambda i: fkw_me(b2_headroom[i]), "forecast")
    dr_me("Variance vs fcst total close", lambda i: fkw_var(b2_variance[i]), "variance")

    df_display = pd.DataFrame([{k: v for k, v in r.items() if not k.startswith("_")} for r in display_rows])
    df_display["Row"] = df_display["Row"].astype(str).str.strip()
    st.dataframe(df_display, use_container_width=True, hide_index=True, height=min(80 + len(display_rows) * 32, 900))

    ch1, ch2 = st.columns(2)
    wk_lbl = [all_weeks[i].strftime("%d %b") for i in range(N_W)]

    with ch1:
        st.caption("Closing total cash balance (£k) — solid = actual, dashed = forecast")
        fig_bal = go.Figure()
        fig_bal.add_trace(go.Scatter(x=wk_lbl[:n_actuals], y=closes[:n_actuals], name="Actual", line=dict(color=BLUE, width=2.5), mode="lines+markers", marker=dict(size=4)))
        fig_bal.add_trace(go.Scatter(x=wk_lbl[n_actuals-1:], y=closes[n_actuals-1:], name="Forecast", line=dict(color=BLUE, width=1.5, dash="dot"), mode="lines+markers", marker=dict(size=3, symbol="circle-open")))
        fig_bal.update_layout(height=220, plot_bgcolor="white", paper_bgcolor="white", yaxis=dict(tickformat=",.0f", ticksuffix="k", gridcolor=GREY), legend=dict(orientation="h", y=1.1, font=dict(size=10)), margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig_bal, use_container_width=True)

    with ch2:
        st.caption("Weekly receipts vs payments (£k) — total group")
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(x=wk_lbl, y=totR, name="Receipts", marker_color=[BLUE if i < n_actuals else "#7B72C8" for i in range(N_W)], offsetgroup=0))
        fig_bar.add_trace(go.Bar(x=wk_lbl, y=[abs(v) for v in totP], name="Payments", marker_color=[RED if i < n_actuals else "#F09595" for i in range(N_W)], offsetgroup=1))
        fig_bar.update_layout(barmode="group", height=220, plot_bgcolor="white", paper_bgcolor="white", yaxis=dict(tickformat=",.0f", ticksuffix="k", gridcolor=GREY), legend=dict(orientation="h", y=1.1, font=dict(size=10)), margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig_bar, use_container_width=True)

    st.caption("Forecast closing total cash balance — 🟢 net inflow week · 🟡 positive but below opening · 🔴 negative")
    gauge_weeks = fc_weeks[:13]
    if gauge_weeks:
        g_cols = st.columns(len(gauge_weeks))
        for fi, wk in enumerate(gauge_weeks):
            i = fi + n_actuals
            if i >= len(closes):
                continue
            cl, op = closes[i], opens[i]
            signal = "🟢" if cl > op else ("🟡" if cl > 0 else "🔴")
            with g_cols[fi]:
                st.metric(label=wk.strftime("%d %b"), value=f"£{cl:,.0f}k", delta=f"open: £{op:,.0f}k · {signal}", delta_color="off")

    st.divider()
    exp_rows = {"Opening total cash balance": opens}
    for l in RECEIPT_LABELS:
        exp_rows[l] = data[l]
    exp_rows["Total receipts"] = totR
    for l in PAYMENT_LABELS:
        exp_rows[l] = data[l]
    exp_rows["Total payments"] = totP
    exp_rows["Net cash"] = net
    exp_rows["Closing total cash balance"] = closes
    col_names = [f"{'ACT_' if i < n_actuals else 'FC_'}{all_weeks[i].strftime('%d%b%Y')}" for i in range(N_W)]
    export_df = pd.DataFrame(exp_rows, index=col_names).T
    buf2 = io.BytesIO()
    export_df.to_excel(buf2, sheet_name="Weekly Cash Flow")
    buf2.seek(0)
    st.download_button("⬇️ Export weekly cash flow (.xlsx)", data=buf2, file_name=f"weekly_cashflow_{latest_date.strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 8 — How it works
# ══════════════════════════════════════════════════════════════════════════════
with tabs[7]:
    st.header("How this dashboard works")
    st.caption("Reference guide for banking and treasury team.")

    with st.expander("📂 Data inputs", expanded=True):
        st.markdown("""
**Two files are uploaded each session via the sidebar — nothing is stored permanently.**

| File | What it contains | Format |
|---|---|---|
| Bank transactions | Daily transaction-level data for all accounts | Excel — `Data Sheet` tab |
| Forecast / client money file | Monthly forecast closing balances + client money | Excel — Row 1: dates, Row 2: client money, Row 3: UK cash, Row 4: Ireland cash |

Optional sidebar inputs:
- SHT / SHGI / TMD actual positions — GBP equivalent from daily cash balance sheet.
- Client money override.
- FX budget rates.
- Remaining receipts / AP sliders.
""")

    with st.expander("🏦 Account mapping"):
        st.markdown("""
Account → entity classification is hard-coded in `ACCOUNT_ENTITY`.

| Entity | Accounts |
|---|---|
| **UK** | SHTL GBP, SHTL EUR, SHTL USD, SHTL PAY GBP, SHTL PAY, TMOOD GBP, TMOOD EUR, TMOOD USD, TMOOD CAD, HJT GBP, HJT AH GBP |
| **Ireland** | SHGI EUR, SHGI GBP, SHGI USD, SHGI CAD, AHD EURO CURRENT ACCOUNT, BOA |
""")

    with st.expander("🎯 3-Month Focus and 4+13 alignment", expanded=True):
        st.markdown("""
This version runs both tabs on the same basis:

- **Weekly 4+13 Outlook** uses total group cash: UK + Ireland.
- **3-Month Focus** also uses the same total-cash close chain.
- Month-end tracking compares against **forecast total close**, not UK close.
- Variance is now **weekly forecast total close less forecast file total close**.

That is the correction that fixes the mismatch between the two tabs.
""")

    with st.expander("🛡️ Compliance"):
        st.markdown("""
Compliance still uses the UK cash requirement:

```text
UK required = client money × selected requirement %
UK headroom = UK cash − UK required
Total headroom = total cash − UK required
```

The monthly and weekly planning views use total cash because cash can be managed across UK / Ireland, while the compliance badge still highlights the UK requirement position.
""")

# ── Footer ───────────────────────────────────────────────────────────────────
st.divider()
c1, c2, c3 = st.columns(3)
c1.caption(f"Bank data: {df_raw['PostDate'].min().strftime('%d %b %Y')} – {latest_date.strftime('%d %b %Y')}")
c2.caption(f"Forecast: {fc.index[0].strftime('%b %Y')} – {fc.index[-1].strftime('%b %Y')} ({len(fc)} months)")
c3.caption(f"Transactions: {len(df_raw):,} · Accounts: {df_raw['AccountName'].nunique()} · FX: EUR {eur_rate:.3f} · USD {usd_rate:.3f} · CAD {cad_rate:.3f}")
