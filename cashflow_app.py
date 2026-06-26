"""
SHG Cash Flow Dashboard — Complete App
=======================================
Integrates: monthly cash vs forecast, compliance, YoY analysis,
3-month focus, opportunities & risks, weekly 4+13 outlook with
FX budget rates and interco netting.

Run:
    streamlit run shg_dashboard_complete.py

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
warnings.filterwarnings('ignore')

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
/* SHG navy primary colour on key interactive elements */
div[data-testid="stTabs"] button[aria-selected="true"] {
    color: #1C1464 !important;
    border-bottom-color: #1C1464 !important;
    font-weight: 600;
}
/* SHG gold accent on sidebar headers */
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] .stMarkdown strong {
    color: #C9A84C;
}
/* SHG navy on metric labels */
[data-testid="stMetricLabel"] { color: #1C1464 !important; }
/* Gold dividers */
hr { border-color: #C9A84C !important; opacity: 0.4; }
/* SHG navy on dataframe header */
thead tr th { background-color: #1C1464 !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ACCOUNT → ENTITY MAPPING  (hard-coded — no upload required)
# Update this dict if accounts are added or entity changes.
# ══════════════════════════════════════════════════════════════════════════════
ACCOUNT_ENTITY = {
    # ── UK entities ───────────────────────────────────────────────────────────
    'SHTL GBP':                 'UK',
    'SHTL EUR':                 'UK',
    'SHTL USD':                 'UK',
    'SHTL PAY GBP':             'UK',
    'SHTL PAY':                 'UK',
    'TMOOD GBP':                'UK',
    'TMOOD EUR':                'UK',
    'TMOOD USD':                'UK',
    'TMOOD CAD':                'UK',
    'HJT GBP':                  'UK',
    'HJT AH GBP':               'UK',
    # ── Ireland entities ─────────────────────────────────────────────────────
    'SHGI EUR':                 'Ireland',
    'SHGI GBP':                 'Ireland',
    'SHGI USD':                 'Ireland',
    'SHGI CAD':                 'Ireland',
    'AHD EURO CURRENT ACCOUNT': 'Ireland',
    'BOA':                      'Ireland',   # Bank of America — Ireland entity
}

# GBP-only UK accounts used for client money compliance
UK_GBP_ACCS = ['SHTL GBP', 'SHTL PAY GBP', 'SHTL PAY', 'TMOOD GBP', 'HJT GBP', 'HJT AH GBP']

UK_ACCS     = [k for k, v in ACCOUNT_ENTITY.items() if v == 'UK']
IRELAND_ACCS= [k for k, v in ACCOUNT_ENTITY.items() if v == 'Ireland']

KEY_INFLOW  = ['AGENT RECEIPTS', 'FD RECEIPT', 'DIRECT RECEIPTS', 'FX TRADE IN',
               'OTHER RECEIPTS', 'TUI RECEIPT', 'INTERCO']
KEY_OUTFLOW = ['AP COGS', 'AP OVH', 'FLIGHT COSTS', 'CUSTOMER REFUNDS',
               'FX TRADE OUT', 'PAYROLL', 'OTHER COSTS', 'TAX', 'INTERCO']

RECEIPT_ROWS = [
    ('DIRECT RECEIPTS', 'DIRECT RECEIPTS'),
    ('AGENT RECEIPTS',  'AGENT RECEIPTS'),
    ('FD RECEIPT',      'FD RECEIPT'),
    ('CUSTOMER REFUND', 'CUSTOMER REFUNDS'),
    ('TUI RECEIPT',     'TUI RECEIPT'),
    ('OTHER RECEIPT',   'OTHER RECEIPTS'),
    ('FX TRADE IN',     'FX TRADE IN'),
    ('INTERCO (net)',   'INTERCO'),
    ('OD INTEREST',     'OVERNIGHT DEPOSIT'),
]
PAYMENT_ROWS = [
    ('AP COGS',        'AP COGS'),
    ('AP OVH',         'AP OVH'),
    ('PAYROLL',        'PAYROLL'),
    ('TAX',            'TAX'),
    ('OTHER CASH OUT', 'OTHER COSTS'),
    ('FX TRADE OUT',   'FX TRADE OUT'),
    ('FLIGHT COSTS',   'FLIGHT COSTS'),
]

WEEKLY_OVR  = {'AP COGS', 'AP OVH', 'PAYROLL', 'FD RECEIPT', 'AGENT RECEIPTS',
               'FX TRADE IN', 'FX TRADE OUT', 'FLIGHT COSTS', 'INTERCO (net)'}
WEEKLY_LOCK = {'PAYROLL'}

MN = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr',  5:'May',  6:'Jun',
      7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}

BLUE   = '#1C1464'  # SHG navy
GREEN  = '#3B6D11'
LBLUE  = '#3D3580'  # SHG navy light
LGREEN = '#639922'
RED    = '#E24B4A'
AMBER  = '#C9A84C'  # SHG gold
GREY   = 'rgba(0,0,0,0.06)'
DEFAULT_FX = {'EUR': 0.86, 'USD': 0.76, 'CAD': 0.53}


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(v, decimals=1):
    if v is None or (isinstance(v, float) and np.isnan(v)): return '—'
    sign = '-' if v < 0 else ''
    a = abs(v)
    if a >= 1e6: return f"{sign}£{a/1e6:.{decimals}f}m"
    if a >= 1e3: return f"{sign}£{a/1e3:.0f}k"
    return f"{sign}£{a:,.0f}"

def safe_get(df, yr, mn, col):
    try:
        if (yr, mn) in df.index and col in df.columns:
            v = df.loc[(yr, mn), col]
            return float(v) if not pd.isna(v) else 0.0
    except: pass
    return 0.0

def safe_entity(entity_m, yr, mn, ent, field):
    try: return float(entity_m.loc[(yr, mn, ent), field])
    except: return 0.0

def get_currency(acc):
    """Derive transaction currency from account name."""
    acc = str(acc).strip()
    if 'EUR' in acc:      return 'EUR'
    if 'USD' in acc:      return 'USD'
    if 'CAD' in acc:      return 'CAD'
    if acc == 'BOA':      return 'USD'   # BOA = USD account
    return 'GBP'


# ── Data loading ───────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_transactions(b):
    df = pd.read_excel(io.BytesIO(b), sheet_name='Data Sheet')
    df['PostDate']  = pd.to_datetime(df['PostDate'])
    df['AccountName'] = df['AccountName'].astype(str).str.strip()

    # Apply hard-coded entity mapping — no upload needed
    df['entity'] = df['AccountName'].map(ACCOUNT_ENTITY)

    # Warn on any unmapped accounts (new accounts added to the bank feed)
    unmapped = df.loc[df['entity'].isna(), 'AccountName'].unique()
    if len(unmapped) > 0:
        st.warning(
            f"⚠️ Unmapped accounts detected: **{', '.join(unmapped)}**. "
            f"Defaulting to Ireland. Ask your developer to add these to ACCOUNT_ENTITY in the code."
        )
    df['entity'] = df['entity'].fillna('Ireland')

    df['TrnSpec']  = df['TrnSpec'].fillna('OTHER').str.strip()
    df.loc[df['TrnSpec'].str.upper().str.strip() == 'PAYROLL  ', 'TrnSpec'] = 'PAYROLL'
    df['Currency'] = df['AccountName'].apply(get_currency)
    df['Year']     = df['PostDate'].dt.year.astype(int)
    df['Month']    = df['PostDate'].dt.month.astype(int)
    df['WeekStart']= df['PostDate'].dt.to_period('W-SUN').apply(lambda p: p.start_time)

    df_op = df[df['TrnSpec'] != 'OVERNIGHT DEPOSIT'].copy()
    df_op['inflow']  = df_op['Amount'].clip(lower=0)
    df_op['outflow'] = df_op['Amount'].clip(upper=0)
    return df, df_op


@st.cache_data(show_spinner=False)
def load_forecast(b):
    raw   = pd.read_excel(io.BytesIO(b), sheet_name='Sheet1', header=None)
    dates = pd.to_datetime(raw.iloc[0, 1:].values)
    fc    = pd.DataFrame({
        'client_money': raw.iloc[1, 1:].values.astype(float),
        'cash_uk':      raw.iloc[2, 1:].values.astype(float),
        'cash_ireland': raw.iloc[3, 1:].values.astype(float),
    }, index=dates)
    fc.index.name = 'Date'
    fc['Year']  = fc.index.year.astype(int)
    fc['Month'] = fc.index.month.astype(int)
    return fc


def extend_forecast_horizon(fc, latest_bank_date):
    """
    Ensure forecast always covers:
      (a) at least December 2027, AND
      (b) at least 12 rolling months beyond the latest bank data date.
    Missing months are extrapolated by seasonal carry-forward:
    same month prior year from the uploaded forecast.
    """
    hard_min    = pd.Timestamp('2027-12-01')
    rolling_min = (latest_bank_date + pd.DateOffset(months=12)).replace(day=1)
    required_end = max(hard_min, rolling_min)

    if fc.index.max() >= required_end:
        return fc   # already covers required horizon

    extra_dates = pd.date_range(
        start = fc.index.max() + pd.DateOffset(months=1),
        end   = required_end,
        freq  = 'MS'
    )

    rows = []
    for dt in extra_dates:
        prior_dt = dt - pd.DateOffset(years=1)
        base = fc.loc[prior_dt] if prior_dt in fc.index else fc.iloc[-1]
        rows.append({
            'client_money': float(base['client_money']),
            'cash_uk':      float(base['cash_uk']),
            'cash_ireland': float(base['cash_ireland']),
        })

    extra = pd.DataFrame(rows, index=extra_dates)
    extra.index.name = 'Date'
    extra['Year']  = extra.index.year.astype(int)
    extra['Month'] = extra.index.month.astype(int)

    out = pd.concat([fc, extra])
    return out[~out.index.duplicated(keep='first')].sort_index()


def build_actuals(df_op):
    entity_m = df_op.groupby(['Year', 'Month', 'entity']).agg(
        inflow=('inflow', 'sum'), outflow=('outflow', 'sum'), net=('Amount', 'sum')
    ).round(0)
    uk_in  = df_op[df_op['entity']=='UK'].groupby(['Year','Month','TrnSpec'])['inflow'].sum().round(0).unstack('TrnSpec').fillna(0)
    uk_out = df_op[df_op['entity']=='UK'].groupby(['Year','Month','TrnSpec'])['outflow'].sum().round(0).unstack('TrnSpec').fillna(0)
    ie_in  = df_op[df_op['entity']=='Ireland'].groupby(['Year','Month','TrnSpec'])['inflow'].sum().round(0).unstack('TrnSpec').fillna(0)
    ie_out = df_op[df_op['entity']=='Ireland'].groupby(['Year','Month','TrnSpec'])['outflow'].sum().round(0).unstack('TrnSpec').fillna(0)
    return entity_m, uk_in, uk_out, ie_in, ie_out


def build_weekly(df_raw, fx_rates):
    """
    Weekly pivot — UK entities only, FX converted to GBP at budget rates.
    Interco: aggregated net per week (intra-UK transfers cancel automatically;
    only net UK↔Ireland movement remains).
    """
    df = df_raw.copy()
    df['AccountName'] = df['AccountName'].astype(str).str.strip()
    df['entity']      = df['AccountName'].map(ACCOUNT_ENTITY).fillna('Ireland')
    df['TrnSpec']     = df['TrnSpec'].fillna('OTHER').str.strip()
    df.loc[df['TrnSpec'].str.upper().str.strip() == 'PAYROLL  ', 'TrnSpec'] = 'PAYROLL'
    df['Currency']    = df['AccountName'].apply(get_currency)
    df['PostDate']    = pd.to_datetime(df['PostDate'])
    df['WeekStart']   = pd.to_datetime(
        df['PostDate'].dt.to_period('W-SUN').apply(lambda p: p.start_time))

    # Apply FX budget rates
    df['Amount_GBP'] = df.apply(
        lambda r: r['Amount'] * fx_rates.get(r['Currency'], 1.0), axis=1)

    # UK operational flows (excl overnight deposits)
    df_uk = df[(df['TrnSpec'] != 'OVERNIGHT DEPOSIT') & (df['entity'] == 'UK')].copy()

    # Overnight deposit net per week (proxy for interest income)
    df_od = df[(df['TrnSpec'] == 'OVERNIGHT DEPOSIT') & (df['entity'] == 'UK')].copy()
    od_weekly = df_od.groupby('WeekStart')['Amount_GBP'].sum().round(0)

    weekly = (df_uk
              .groupby(['WeekStart', 'TrnSpec'])['Amount_GBP']
              .sum().round(0)
              .unstack('TrnSpec')
              .fillna(0))

    return weekly, od_weekly


def build_weekly_forecast(weekly, od_weekly, n_fc=13, fx_rates=None):
    """13-week forecast: prior year same ISO week × YoY trend. Interco = 0 (blank)."""
    if fx_rates is None:
        fx_rates = DEFAULT_FX
    last_date = weekly.index.max()
    fc_weeks  = [last_date + pd.Timedelta(weeks=i+1) for i in range(n_fc)]
    all_specs = [s for _, s in RECEIPT_ROWS[:-1]] + [s for _, s in PAYMENT_ROWS]
    all_specs = [s for s in all_specs if s != 'OVERNIGHT DEPOSIT']

    # No trend multiplier — prior year same week IS the forecast.
    # Trend factors were distorting badly because the data window is short.
    # Same week last year is the correct seasonal anchor for this business.
    forecast = {spec: [] for spec in all_specs}
    forecast['INTERCO']     = [0] * n_fc   # blank — net zero assumed
    forecast['OD_INTEREST'] = [11000] * n_fc

    # Payroll cycle detection — with NaN guards throughout
    large = -250000; small = -30000; days_since = 21   # safe defaults
    if 'PAYROLL' in weekly.columns:
        pay_s = weekly['PAYROLL'].replace(0, np.nan).dropna()
        if len(pay_s) > 0:
            med = pay_s.median()
            if pd.notna(med) and med != 0:
                large_s = pay_s[pay_s < med * 2]
                small_s = pay_s[pay_s >= med]
                large_med = large_s.median()
                small_med = small_s.median()
                if pd.notna(large_med): large = float(large_med)
                if pd.notna(small_med) and len(small_s) > 0: small = float(small_med)
            last_pay_idx = pay_s.index.max()
            if pd.notna(last_pay_idx):
                days_since = int((last_date - last_pay_idx).days)

    pay_fc = []
    for fw in fc_weeks:
        da = (fw - last_date).days + days_since
        if da % 28 < 7:          pay_fc.append(int(large))
        elif (da + 14) % 28 < 7: pay_fc.append(int(small))
        else:                     pay_fc.append(0)
    forecast['PAYROLL'] = pay_fc

    for spec in all_specs:
        if spec in ('INTERCO', 'PAYROLL'): continue
        col_fc = []
        for fw in fc_weeks:
            # Use prior year same ISO week — search ±2 weeks if exact match missing
            py_date = fw - pd.Timedelta(weeks=52)
            base = None
            for delta in [0, 1, -1, 2, -2]:
                sd = py_date + pd.Timedelta(weeks=delta)
                if spec in weekly.columns and sd in weekly.index:
                    base = float(weekly.loc[sd, spec])
                    break
            if base is None:
                # Fall back to median of same month across all years in data
                if spec in weekly.columns:
                    same_mn = weekly[spec][weekly.index.month == fw.month]
                    base = float(same_mn.median()) if len(same_mn) > 0 else 0.0
                else:
                    base = 0.0
            col_fc.append(round(base))
        forecast[spec] = col_fc

    return forecast, fc_weeks


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 💷 SHG Cash Dashboard")
    st.divider()
    st.markdown("**1 · Bank transactions**")
    tx_file = st.file_uploader("Upload bank Excel", type=['xlsx'], key='tx',
                               help="Must contain 'Data Sheet' tab")
    st.markdown("**2 · Forecast / client money**")
    fc_file = st.file_uploader("Upload forecast Excel", type=['xlsx'], key='fc',
                               help="Forecast file format — rows: Client money / Cash UK / Cash Ireland")
    st.divider()
    uk_pct = st.slider("UK cash requirement (%)", 50, 100, 70, 5)
    st.divider()
    st.markdown("**FX budget rates** (£ per 1 foreign unit)")
    st.caption("Spot 23 Jun 2026: EUR 0.8631 · USD 0.7582 · CAD 0.5332")
    eur_rate = st.number_input("EUR → GBP", value=0.86, min_value=0.50, max_value=1.50, step=0.01, format="%.3f")
    usd_rate = st.number_input("USD → GBP", value=0.76, min_value=0.30, max_value=1.20, step=0.01, format="%.3f")
    cad_rate = st.number_input("CAD → GBP", value=0.53, min_value=0.20, max_value=1.00, step=0.01, format="%.3f")
    fx_rates = {'EUR': eur_rate, 'USD': usd_rate, 'CAD': cad_rate, 'GBP': 1.0}
    st.divider()
    st.markdown("**Actual cash positions (GBP equiv)**")
    st.caption("Enter from the daily cash balance sheet. Overrides forecast file as opening balance.")
    pos_date = st.date_input("As at date", value=None, help="Date of the cash position snapshot")
    pos_sht  = st.number_input("SHT (£)",  value=0, min_value=0, max_value=999_999_999, step=1000, format="%d")
    pos_shgi = st.number_input("SHGI (£)", value=0, min_value=0, max_value=999_999_999, step=1000, format="%d")
    pos_tmd  = st.number_input("TMD (£)",  value=0, min_value=0, max_value=999_999_999, step=1000, format="%d")
    pos_total_uk      = pos_sht + pos_tmd
    pos_total_ireland = pos_shgi
    pos_total         = pos_sht + pos_shgi + pos_tmd
    use_actual_pos    = pos_total > 0
    if use_actual_pos:
        st.success(f"Total: £{pos_total:,.0f}  ·  UK: £{pos_total_uk:,.0f}  ·  IE: £{pos_total_ireland:,.0f}")
    else:
        st.caption("Enter values above to use actual position (currently using forecast file)")
    st.divider()
    st.markdown("**Client money override**")
    st.caption("Override forecast file client money for current month if you have a more accurate figure.")
    cm_override = st.number_input(
        "Client money (£) — leave 0 to use forecast file",
        value=0, min_value=0, max_value=999_999_999, step=100_000, format="%d")
    use_cm_override = cm_override > 0
    if use_cm_override:
        st.success(f"Client money: £{cm_override:,.0f} · UK req: £{cm_override*0.70:,.0f}")

    if use_actual_pos:
        st.divider()
        st.markdown("**Remaining flows adjustment**")
        st.caption("Adjust remaining month flows vs prior year implied. 0% = follow LY exactly.")
        adj_receipts = st.slider("Receipts vs LY remaining (%)", -50, 50, 0, 5,
                                  help="Positive = collecting more than LY pace, negative = less")
        adj_payments = st.slider("AP payments vs LY remaining (%)", -50, 50, 0, 5,
                                  help="Negative = holding back payments vs LY, positive = paying more")
    else:
        adj_receipts = 0
        adj_payments = 0
    st.divider()
    st.caption(
        "Account mapping is hard-coded — no upload required.\n"
        "BOA = Ireland entity.\n"
        "Overnight deposits excluded (net zero).\n"
        "Opening balance from forecast file prior month close (or actual position above).\n"
        "Interco nets within UK entities.\n"
        "Forecast always extends to Dec 2027 minimum + 12 months rolling beyond latest bank data."
    )

if not tx_file or not fc_file:
    st.title("💷 SHG Cash Flow Dashboard")
    c1, c2 = st.columns(2)
    with c1: st.info("📂 Upload **bank transactions** Excel (sidebar)\n`Data Sheet` tab required")
    with c2: st.info("📂 Upload **forecast / client money** Excel (sidebar)\nForecast file format: rows = Client money / Cash UK / Cash Ireland")
    st.stop()

# ── Load & enrich ─────────────────────────────────────────────────────────────
with st.spinner("Loading data..."):
    df_raw, df_op = load_transactions(tx_file.read())
    fc_raw         = load_forecast(fc_file.read())
    entity_m, uk_in_spec, uk_out_spec, ie_in_spec, ie_out_spec = build_actuals(df_op)
    weekly_raw, od_weekly = build_weekly(df_raw, fx_rates)

    # ── Dynamic forecast horizon ──────────────────────────────────────────────
    # Always show at least Dec 2027 AND rolling 12 months beyond latest bank data.
    # extend_forecast_horizon() carries forward missing months using prior-year
    # seasonal values from the uploaded Book2 file.
    latest_bank_tmp = df_raw['PostDate'].max()
    fc_raw          = extend_forecast_horizon(fc_raw, latest_bank_tmp)

    # Weekly forecast: extend to cover the same horizon
    hard_min_w    = pd.Timestamp('2027-12-31')
    rolling_min_w = latest_bank_tmp + pd.Timedelta(weeks=52)
    required_w    = max(hard_min_w, rolling_min_w)
    last_act_wk   = weekly_raw.index.max()
    n_fc_dynamic  = max(13, int((required_w - last_act_wk).days // 7) + 1)
    fc_base, fc_weeks = build_weekly_forecast(weekly_raw, od_weekly,
                                               n_fc=n_fc_dynamic, fx_rates=fx_rates)

fc = fc_raw.copy()
fc['uk_required']    = fc['client_money'] * uk_pct / 100
fc['uk_headroom']    = fc['cash_uk'] - fc['uk_required']
fc['total_cash']     = fc['cash_uk'] + fc['cash_ireland']
fc['total_headroom'] = fc['total_cash'] - fc['uk_required']
fc['headroom_pct']   = fc['uk_headroom'] / fc['uk_required']
fc['fc_implied_net'] = fc['cash_uk'].diff()

latest_date = df_raw['PostDate'].max()
latest_yr   = int(latest_date.year)
latest_mn   = int(latest_date.month)
latest_day  = int(latest_date.day)

def is_actual(yr, mn):
    if yr < latest_yr:                    return True
    if yr == latest_yr and mn <= latest_mn: return True
    return False

fc['is_actual']  = fc.apply(lambda r: is_actual(int(r['Year']), int(r['Month'])), axis=1)
fc['is_partial'] = fc.apply(lambda r: int(r['Year']) == latest_yr and int(r['Month']) == latest_mn, axis=1)

latest_fc      = fc[fc['is_actual']].iloc[-1]
kpi_client_mon = float(cm_override) if use_cm_override else float(latest_fc['client_money'])
req_now        = kpi_client_mon * uk_pct / 100
# If actual position entered, use that for headroom — more accurate than Book2
if use_actual_pos:
    hroom_now = pos_total_uk - req_now
else:
    hroom_now = float(latest_fc['uk_headroom'])
hpct_now  = hroom_now / req_now if req_now > 0 else 0

# ── Page header / KPIs ────────────────────────────────────────────────────────
st.title(f"💷 SHG Cash — Latest: {latest_date.strftime('%d %b %Y')}")
if use_actual_pos:
    pos_date_str = pos_date.strftime("%d %b %Y") if pos_date else "date not set"
    st.info(
        f"📍 **Actual cash position as at {pos_date_str}** — "
        f"SHT: £{pos_sht:,.0f} · SHGI: £{pos_shgi:,.0f} · TMD: £{pos_tmd:,.0f} · "
        f"**Total: £{pos_total:,.0f}** · UK: £{pos_total_uk:,.0f} · Ireland: £{pos_total_ireland:,.0f}"
    )

# Compliance badge rendered inside KPI block below

# ── KPI calculation ───────────────────────────────────────────────────────────
# Book2 month-end forecast values for current month
fc_me_uk  = float(latest_fc['cash_uk'])
fc_me_ie  = float(latest_fc['cash_ireland'])
fc_me_tot = float(latest_fc['total_cash'])
kpi_client_mon = float(latest_fc['client_money'])

if use_actual_pos and pos_date is not None:
    import calendar
    days_in_month  = calendar.monthrange(pos_date.year, pos_date.month)[1]
    days_elapsed   = pos_date.day
    days_remaining = days_in_month - days_elapsed
    weeks_remaining = days_remaining / 7.0

    # ── Remaining flows: rolling 4-week average (excludes lumpy FX/Interco) ──
    # This avoids distortion from prior year opening-month anomalies.
    # We use the last 4 weeks of actual weekly flows as the run rate,
    # then pro-rate by days remaining in the month.
    KEY_IN_CORE  = ['AGENT RECEIPTS','FD RECEIPT','DIRECT RECEIPTS',
                    'OTHER RECEIPTS','TUI RECEIPT']
    KEY_OUT_CORE = ['AP COGS','AP OVH','FLIGHT COSTS','PAYROLL','OTHER COSTS','TAX']

    # Build weekly pivot from df_op (UK only, operational)
    df_op_w = df_op[df_op['entity'] == 'UK'].copy()
    df_op_w['WeekStart'] = pd.to_datetime(
        df_op_w['PostDate'].dt.to_period('W-SUN').apply(lambda p: p.start_time))

    wk_core_in  = (df_op_w[df_op_w['TrnSpec'].isin(KEY_IN_CORE)]
                   .groupby('WeekStart')['inflow'].sum())
    wk_core_out = (df_op_w[df_op_w['TrnSpec'].isin(KEY_OUT_CORE)]
                   .groupby('WeekStart')['outflow'].sum())

    avg_wk_in  = float(wk_core_in.tail(4).mean())  if len(wk_core_in)  >= 1 else 0.0
    avg_wk_out = float(wk_core_out.tail(4).mean()) if len(wk_core_out) >= 1 else 0.0

    # Apply sidebar adjustment sliders
    raw_rem_uk_in  = avg_wk_in  * weeks_remaining
    raw_rem_uk_out = avg_wk_out * weeks_remaining   # negative

    rem_uk_in  = raw_rem_uk_in  * (1 + adj_receipts / 100)
    rem_uk_out = raw_rem_uk_out * (1 + adj_payments / 100)

    # Ireland: simpler — use same fraction of LY month (Ireland less affected by opening anomaly)
    yr_ly = latest_yr - 1; mn_ly = latest_mn
    ly_ie_in  = sum(safe_get(ie_in_spec,  yr_ly, mn_ly, c)
                    for c in KEY_INFLOW if c in ie_in_spec.columns)
    ly_ie_out = sum(safe_get(ie_out_spec, yr_ly, mn_ly, c)
                    for c in KEY_OUTFLOW if c in ie_out_spec.columns)
    rem_ie_in  = ly_ie_in  * (days_remaining / days_in_month) * (1 + adj_receipts / 100)
    rem_ie_out = ly_ie_out * (days_remaining / days_in_month) * (1 + adj_payments / 100)

    # Forecast month-end = actual snapshot + remaining expected flows
    kpi_uk_cash    = pos_total_uk      + rem_uk_in + rem_uk_out
    kpi_ie_cash    = pos_total_ireland + rem_ie_in + rem_ie_out
    kpi_total_cash = kpi_uk_cash + kpi_ie_cash

    # ── Room to pay ──────────────────────────────────────────────────────────
    # Key question: how much can I pay out over remaining days and still hit
    # the forecast month-end close?
    # Room to pay = Actual UK now + Expected remaining receipts − Forecast close
    # Positive = you have budget to pay; negative = you need to hold back
    room_to_pay     = pos_total_uk + rem_uk_in - fc_me_uk
    current_run_rate = rem_uk_out  # what outflows look like at current pace (negative)
    ap_headroom      = room_to_pay - abs(current_run_rate)  # spare capacity vs run rate

    still_to_collect = rem_uk_in
    still_to_pay     = rem_uk_out   # negative
    adj_note = ""
    if adj_receipts != 0: adj_note += f" · receipts {adj_receipts:+d}%"
    if adj_payments != 0: adj_note += f" · AP {adj_payments:+d}%"
    kpi_subtitle = (f"Forecast {pd.Timestamp(latest_yr, latest_mn, days_in_month).strftime('%d %b')} · "
                    f"{days_remaining}d remaining · 4wk run rate{adj_note}")
else:
    # No actual position — use Book2 as-is
    kpi_uk_cash      = fc_me_uk
    kpi_ie_cash      = fc_me_ie
    kpi_total_cash   = fc_me_tot
    still_to_collect = None
    still_to_pay     = None
    room_to_pay      = None
    ap_headroom      = None
    days_remaining   = 0
    kpi_subtitle     = f"Forecast {MN[latest_mn]} {latest_yr} month-end"

# Headroom: uses overridden req_now if client money override is set
hroom_now = kpi_uk_cash - req_now
hpct_now  = hroom_now / req_now if req_now > 0 else 0

if hpct_now >= 0.20:
    st.success(f"✅ COMPLIANT — UK headroom {fmt(hroom_now)} ({hpct_now:.0%} of req {fmt(req_now)}) · {kpi_subtitle}")
elif hpct_now >= 0:
    st.warning(f"⚠️ CAUTION — UK headroom {fmt(hroom_now)} ({hpct_now:.0%}) · {kpi_subtitle}")
else:
    st.error(f"🚨 BREACH — UK cash {fmt(abs(hroom_now))} below requirement · {kpi_subtitle}")

k = st.columns(6)
k[0].metric("UK Cash (fcst m/e)",      fmt(kpi_uk_cash),
            delta=f"actual now: {fmt(pos_total_uk)}" if use_actual_pos else None)
k[1].metric("Ireland Cash (fcst m/e)", fmt(kpi_ie_cash),
            delta=f"actual now: {fmt(pos_total_ireland)}" if use_actual_pos else None)
k[2].metric("Total Cash (fcst m/e)",   fmt(kpi_total_cash),
            delta=f"actual now: {fmt(pos_total)}" if use_actual_pos else None)
k[3].metric("Client Money",
            fmt(kpi_client_mon),
            delta="overridden" if use_cm_override else None,
            delta_color="off")
k[4].metric("UK Required",
            fmt(req_now),
            delta=f"{uk_pct}% of {'override' if use_cm_override else 'forecast file'}",
            delta_color="off")
k[5].metric("UK Headroom (fcst m/e)",  fmt(hroom_now),
            delta=f"{hpct_now:.0%}",
            delta_color="normal" if hroom_now >= 0 else "inverse")

if use_actual_pos and still_to_collect is not None and room_to_pay is not None:
    # ── AP payment room — clear action box ───────────────────────────────────
    ap_over = room_to_pay < 0
    ap_tight = (not ap_over) and ap_headroom < 0
    ap_ok    = ap_headroom >= 0

    if ap_over:
        st.error(
            f"🚨 **AP payments must be reduced this month.** "
            f"At current pace you will miss the forecast close by **{fmt(abs(room_to_pay))}**. "
            f"Hold back at least **{fmt(abs(room_to_pay))}** of AP before month-end."
        )
    elif ap_tight:
        st.warning(
            f"⚠️ **AP payments running slightly hot.** "
            f"You have room to pay **{fmt(room_to_pay)}** total over the remaining {days_remaining} days, "
            f"but your current run rate implies **{fmt(abs(still_to_pay))}** — "
            f"reduce AP by **{fmt(abs(ap_headroom))}** to hit forecast."
        )
    else:
        st.success(
            f"✅ **AP payments on track.** "
            f"Room to pay **{fmt(room_to_pay)}** over remaining {days_remaining} days · "
            f"current run rate **{fmt(abs(still_to_pay))}** · "
            f"spare capacity **{fmt(ap_headroom)}**."
        )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric(
        f"Expected receipts · {days_remaining}d left",
        fmt(still_to_collect),
        delta="4-week rolling average", delta_color="off")
    c2.metric(
        f"AP run rate · {days_remaining}d left",
        fmt(abs(still_to_pay)),
        delta="4-week rolling average", delta_color="off")
    c3.metric(
        "Total AP room to month-end",
        fmt(room_to_pay),
        delta=f"= actual UK + receipts − forecast close",
        delta_color="off")
    c4.metric(
        "Spare capacity vs run rate",
        fmt(ap_headroom),
        delta="release AP" if ap_ok else "hold back AP",
        delta_color="normal" if ap_ok else "inverse")

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# SHARED WEEKLY CLOSE CHAIN
# Computed once here, used by both the Weekly 4+13 tab and 3-Month Focus tab.
# Includes session state overrides so both tabs see the same adjusted numbers.
# ══════════════════════════════════════════════════════════════════════════════
def _build_shared_closes(weekly_raw, od_weekly, fc_weeks, fc_base, fc,
                          adj_receipts, adj_payments,
                          use_actual_pos, pos_total_uk, latest_mn, latest_yr):
    """Build the weekly closing balance chain, applying overrides and sliders.
    Returns: (all_weeks, closes_k) where closes_k is in £k."""
    n_act     = 4
    act_wks   = weekly_raw.index[-n_act:].tolist()
    all_wks   = act_wks + list(fc_weeks)
    N         = len(all_wks)

    CORE_IN   = ['DIRECT RECEIPTS','AGENT RECEIPTS','FD RECEIPT','TUI RECEIPT']
    AP_ROWS   = ['AP COGS','AP OVH']
    FIXED_OUT = ['PAYROLL','TAX','FLIGHT COSTS']
    ALL_SPEC  = CORE_IN + AP_ROWS + FIXED_OUT

    SLIDER_REC = set(CORE_IN)
    SLIDER_AP  = set(AP_ROWS)

    def base_val(spec, wk, i):
        if i < n_act:
            return float(weekly_raw.loc[wk, spec]) / 1000                    if spec in weekly_raw.columns and wk in weekly_raw.index else 0.0
        else:
            return fc_base.get(spec, [0]*13)[i - n_act] / 1000                    if spec in fc_base else 0.0

    # Build row data respecting overrides and sliders
    # Overrides stored in £ (not £k) — divide by 1000
    ov = st.session_state.get('weekly_ov', {})
    ROW_SPECS_SHARED = [
        ('DIRECT RECEIPTS', False), ('AGENT RECEIPTS', False),
        ('FD RECEIPT', False),      ('TUI RECEIPT',    False),
        ('OTHER RECEIPT', True),    ('FX TRADE IN',    True),
        ('INTERCO (net)', True),    ('OD INTEREST',    False),
        ('AP COGS',       False),   ('AP OVH',         False),
        ('PAYROLL',       False),   ('TAX',            False),
        ('OTHER CASH OUT',True),    ('FX TRADE OUT',   True),
        ('FLIGHT COSTS',  False),
    ]

    net_per_wk = []
    for i, wk in enumerate(all_wks):
        wk_net = 0.0
        for label, blank_fc in ROW_SPECS_SHARED:
            ov_key = f"{label}_{i}"
            if ov_key in ov:
                wk_net += ov[ov_key] / 1000
                continue
            if i < n_act:
                spec = label if label != 'OD INTEREST' else None
                if label == 'OD INTEREST':
                    # Use actual OD net for actuals (same as weekly tab)
                    v = float(od_weekly.get(wk, 0)) / 1000
                elif spec and spec in weekly_raw.columns and wk in weekly_raw.index:
                    v = float(weekly_raw.loc[wk, spec]) / 1000
                else:
                    v = 0.0
            else:
                if blank_fc:
                    v = 0.0
                elif label == 'OD INTEREST':
                    v = 11.0
                elif label == 'PAYROLL':
                    v = fc_base.get('PAYROLL', [0]*13)[i - n_act] / 1000
                else:
                    v = fc_base.get(label, [0]*13)[i - n_act] / 1000                         if label in fc_base else 0.0
                # Apply sliders to forecast weeks
                if label in SLIDER_REC and adj_receipts != 0:
                    v *= (1 + adj_receipts / 100)
                elif label in SLIDER_AP and adj_payments != 0:
                    v *= (1 + adj_payments / 100)
            wk_net += v
        net_per_wk.append(wk_net)

    # Opening balance
    book2_before = fc[fc.index < act_wks[0]]
    open_k = (pos_total_uk / 1000) if use_actual_pos else              (float(book2_before.iloc[-1]['cash_uk']) / 1000
              if not book2_before.empty else 0.0)

    closes_k = [0.0] * N
    closes_k[0] = open_k + net_per_wk[0]
    for i in range(1, N):
        closes_k[i] = closes_k[i-1] + net_per_wk[i]

    return all_wks, closes_k


# Initialise session state before computing (weekly tab also does this but we may run first)
if 'weekly_ov' not in st.session_state:
    st.session_state.weekly_ov = {}

_shared_all_weeks, _shared_closes_k = _build_shared_closes(
    weekly_raw, od_weekly, fc_weeks, fc_base, fc,
    adj_receipts, adj_payments,
    use_actual_pos, pos_total_uk, latest_mn, latest_yr)

def shared_month_end_close(yr, mn):
    """Return weekly-outlook closing balance (£) at end of given month.
    Uses the pre-tab chain which includes all overrides and sliders.
    This runs before any tab renders so is always consistent."""
    wks = [(i, w) for i, w in enumerate(_shared_all_weeks)
           if w.year == yr and w.month == mn]
    if not wks: return None
    last_i = max(wks, key=lambda x: x[0])[0]
    return _shared_closes_k[last_i] * 1000   # £k → £

# ── Tabs ──────────────────────────────────────────────────────────────────────
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
fc_chart['label'] = fc_chart['Month'].map(MN) + ' ' + fc_chart['Year'].astype(str)
labels = fc_chart['label'].tolist()

def split_act_fc(series):
    act  = series.where(fc_chart['is_actual'] & ~fc_chart['is_partial'])
    fc_s = series.where(~fc_chart['is_actual'] | fc_chart['is_partial'])
    idx  = fc_chart[fc_chart['is_actual']].index[-1]
    if idx in fc_s.index: fc_s.loc[idx] = series.loc[idx]
    return act.tolist(), fc_s.tolist()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Cash vs Forecast
# ══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    fig = go.Figure()
    for series, name, color, width in [
        (fc['cash_uk'],        'UK cash',        BLUE,   2.5),
        (fc['total_cash'],     'Total cash',     GREEN,  2.5),
        (fc['uk_headroom'],    'UK headroom',    LBLUE,  1.5),
        (fc['total_headroom'], 'Total headroom', LGREEN, 1.5),
    ]:
        act, fcast = split_act_fc(series / 1e6)
        fig.add_trace(go.Scatter(x=labels, y=act,   name=f'{name} (actual)',
            line=dict(color=color, width=width), mode='lines+markers', marker=dict(size=3)))
        fig.add_trace(go.Scatter(x=labels, y=fcast, name=f'{name} (forecast)',
            line=dict(color=color, width=max(1, width-0.5), dash='dot'),
            mode='lines+markers', marker=dict(size=2, symbol='circle-open')))
    fig.add_trace(go.Scatter(x=labels, y=(fc['uk_required']/1e6).tolist(),
        name='UK required', line=dict(color=RED, dash='dash', width=1.5), mode='lines'))
    fig.add_hline(y=0, line_color='rgba(0,0,0,0.15)', line_width=0.5)
    fig.update_layout(height=380, plot_bgcolor='white', paper_bgcolor='white',
        yaxis=dict(tickformat=',.1f', ticksuffix='m', gridcolor=GREY, title='£m'),
        xaxis=dict(tickangle=45),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, font=dict(size=10)),
        margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig, use_container_width=True)

    rows = []
    for _, r in fc.iterrows():
        hp = r['headroom_pct']
        rows.append({
            'Year': int(r['Year']), 'Month': MN[int(r['Month'])],
            'Type': 'Actual' if r['is_actual'] else 'Forecast',
            'Client Money':    fmt(r['client_money']),
            'UK Cash':         fmt(r['cash_uk']),
            'Ireland Cash':    fmt(r['cash_ireland']),
            'Total Cash':      fmt(r['total_cash']),
            'UK Required':     fmt(r['uk_required']),
            'UK Headroom':     fmt(r['uk_headroom']),
            'Total Headroom':  fmt(r['total_headroom']),
            'Headroom %':      f"{hp:.0%}" if not np.isnan(hp) else '—',
            'Status':          ('✅ Compliant' if hp >= 0.20 else ('⚠️ Caution' if hp >= 0 else '🚨 Breach')),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Compliance
# ══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    breaches = fc[(fc['uk_headroom'] < 0) & (~fc['is_actual'])]
    cautions = fc[(fc['uk_headroom'] >= 0) & (fc['headroom_pct'] < 0.20) & (~fc['is_actual'])]
    if not breaches.empty:
        st.error("🚨 **Forecast breach months:** " + ", ".join(
            f"{MN[int(r['Month'])]} {int(r['Year'])} ({fmt(r['uk_headroom'])})"
            for _, r in breaches.iterrows()))
    if not cautions.empty:
        st.warning("⚠️ **Caution months:** " + ", ".join(
            f"{MN[int(r['Month'])]} {int(r['Year'])} ({r['headroom_pct']:.0%})"
            for _, r in cautions.iterrows()))

    fig2 = make_subplots(rows=2, cols=1,
        subplot_titles=['UK headroom (£m)', 'Total headroom (£m)'],
        vertical_spacing=0.14)
    for series, ri, col in [(fc['uk_headroom'], 1, BLUE), (fc['total_headroom'], 2, GREEN)]:
        act, fcast = split_act_fc(series / 1e6)
        ri_int, gi_int, bi_int = int(col[1:3], 16), int(col[3:5], 16), int(col[5:], 16)
        fig2.add_trace(go.Scatter(x=labels, y=act,
            line=dict(color=col, width=2), fill='tozeroy',
            fillcolor=f'rgba({ri_int},{gi_int},{bi_int},0.08)', name='Actual'),
            row=ri, col=1)
        fig2.add_trace(go.Scatter(x=labels, y=fcast,
            line=dict(color=col, width=1.5, dash='dot'), name='Forecast'),
            row=ri, col=1)
        fig2.add_hline(y=0, line_color=RED, line_dash='dash', line_width=1.5, row=ri, col=1)
    fig2.update_layout(height=480, plot_bgcolor='white', paper_bgcolor='white',
        showlegend=False, margin=dict(l=0, r=0, t=40, b=0))
    fig2.update_yaxes(tickformat=',.1f', ticksuffix='m', gridcolor=GREY)
    st.plotly_chart(fig2, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# YoY helper
# ══════════════════════════════════════════════════════════════════════════════
def build_yoy_table(in_spec, out_spec, cats, is_out, entity_label):
    spec        = out_spec if is_out else in_spec
    all_ym      = sorted(set(spec.index.tolist()))
    months_avail= sorted(set(mn for (_, mn) in all_ym))
    years_avail = sorted(set(yr for (yr, _) in all_ym))
    rows = []
    for cat in cats:
        if cat not in spec.columns: continue
        row = {'Category': cat}
        yoy_pairs = []
        for mn in months_avail:
            for yr in years_avail:
                v = safe_get(spec, yr, mn, cat)
                if is_out: v = abs(v)
                row[f"{MN[mn]} {yr}"] = v
            yr_vals = [(yr, abs(safe_get(spec, yr, mn, cat)) if is_out else safe_get(spec, yr, mn, cat))
                       for yr in years_avail if safe_get(spec, yr, mn, cat) != 0]
            for i in range(len(yr_vals) - 1):
                y1, v1 = yr_vals[i]; y2, v2 = yr_vals[i+1]
                if v1 != 0 and y2 == y1 + 1:
                    yoy_pairs.append((y1, y2, mn, (v2 - v1) / abs(v1)))
        row['_yoy'] = yoy_pairs
        rows.append(row)
    if not rows: return
    col_order = [f"{MN[mn]} {yr}" for mn in months_avail for yr in years_avail]
    col_order  = [c for c in col_order if any(c in r for r in rows)]
    df_s = pd.DataFrame(rows).fillna(0)
    df_s = df_s[['Category'] + [c for c in col_order if c in df_s.columns] + ['_yoy']]
    disp = df_s.drop('_yoy', axis=1).copy()
    for col in disp.columns[1:]:
        disp[col] = disp[col].apply(lambda v: fmt(v) if v != 0 else '—')
    for i in range(len(years_avail) - 1):
        y1, y2  = years_avail[i], years_avail[i+1]
        yoy_col = f"YoY {y1}→{y2}"
        def calc_yoy(rd, y1=y1, y2=y2):
            pairs = [p for p in rd['_yoy'] if p[0] == y1 and p[1] == y2]
            return np.mean([p[3] for p in pairs]) if pairs else None
        df_s[yoy_col] = df_s.apply(calc_yoy, axis=1)
        disp[yoy_col] = df_s[yoy_col].apply(lambda v: f"{v:+.0%}" if v is not None else '—')
    st.dataframe(disp, use_container_width=True, hide_index=True)
    if len(years_avail) >= 2:
        fig = go.Figure()
        for yi, yr in enumerate(years_avail):
            vals = []
            for mn in months_avail:
                total = sum(abs(safe_get(spec, yr, mn, c)) if is_out else safe_get(spec, yr, mn, c)
                            for c in cats if c in spec.columns)
                vals.append(total / 1e6)
            clr = ['rgba(136,135,128,0.55)', 'rgba(28,20,100,0.7)',
                   'rgba(59,109,17,0.7)', 'rgba(201,168,76,0.7)'][yi % 4]
            fig.add_trace(go.Bar(x=[MN[m] for m in months_avail], y=vals,
                name=str(yr), marker_color=clr, offsetgroup=yi))
        fig.update_layout(barmode='group', height=220,
            title=f'{entity_label} {"outflows" if is_out else "inflows"} by month (£m)',
            plot_bgcolor='white', paper_bgcolor='white',
            yaxis=dict(tickformat=',.1f', ticksuffix='m', gridcolor=GREY),
            legend=dict(orientation='h', yanchor='bottom', y=1.02, font=dict(size=10)),
            margin=dict(l=0, r=0, t=40, b=0))
        st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TABs 3, 4, 5 — YoY
# ══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    uk_sub = st.tabs(["📥 UK Inflow YoY", "📤 UK Outflow YoY"])
    with uk_sub[0]:
        YOY_EXCL = {'FX TRADE IN', 'FX TRADE OUT'}
        avail_in = [c for c in KEY_INFLOW if c in uk_in_spec.columns and c not in YOY_EXCL]
        st.caption("Like-for-like: same calendar month year-on-year. FX trade excluded (lumpy/treasury).")
        build_yoy_table(uk_in_spec, uk_out_spec, avail_in, False, "UK")
    with uk_sub[1]:
        YOY_EXCL = {'FX TRADE IN', 'FX TRADE OUT'}
        avail_out = [c for c in KEY_OUTFLOW if c in uk_out_spec.columns and c not in YOY_EXCL]
        st.caption("Outflows as positive. FX trade excluded. Red YoY = payments running above prior year.")
        build_yoy_table(uk_in_spec, uk_out_spec, avail_out, True, "UK")

with tabs[3]:
    st.caption("Ireland entities — like-for-like month comparison.")
    ie_in_cats  = [c for c in KEY_INFLOW  if c in ie_in_spec.columns]
    ie_out_cats = [c for c in KEY_OUTFLOW if c in ie_out_spec.columns]
    ie_sub = st.tabs(["📥 Ireland Inflow YoY", "📤 Ireland Outflow YoY"])
    with ie_sub[0]:
        YOY_EXCL = {'FX TRADE IN', 'FX TRADE OUT'}
        ie_in_cats = [c for c in KEY_INFLOW if c in ie_in_spec.columns and c not in YOY_EXCL]
        if ie_in_cats:
            st.caption("Inflows only. FX trade excluded. Green YoY = growing vs prior year. Red = declining.")
            build_yoy_table(ie_in_spec, ie_out_spec, ie_in_cats, False, "Ireland")
        else:
            st.info("No Ireland inflow data available.")
    with ie_sub[1]:
        YOY_EXCL = {'FX TRADE IN', 'FX TRADE OUT'}
        ie_out_cats = [c for c in KEY_OUTFLOW if c in ie_out_spec.columns and c not in YOY_EXCL]
        if ie_out_cats:
            st.caption("Outflows as positive. FX trade excluded. Red YoY = running above prior year.")
            build_yoy_table(ie_in_spec, ie_out_spec, ie_out_cats, True, "Ireland")
        else:
            st.info("No Ireland outflow data available.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — 3-Month Focus
# ══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    focus_months = []
    for delta in range(3):
        mn = latest_mn + delta; yr = latest_yr
        while mn > 12: mn -= 12; yr += 1
        focus_months.append((yr, mn))

    # ── Use shared weekly close chain (includes overrides + sliders) ─────────────
    _wk_outlook_month_close = shared_month_end_close


    st.caption(
        f"Bank data to **{latest_date.strftime('%d %b %Y')}**. "
        f"Current month ({MN[latest_mn]} {latest_yr}) + next 2. "
        f"**AP capacity uses weekly outlook close as opening** — reconciled with forecast targetarget."
    )
    cols = st.columns(3)
    for ci, (fyr, fmn) in enumerate(focus_months):
        fc_row = fc[(fc['Year'] == fyr) & (fc['Month'] == fmn)]
        if fc_row.empty:
            with cols[ci]: st.info(f"{MN[fmn]} {fyr} — no forecast data"); continue
        fc_row   = fc_row.iloc[0]
        is_part  = (fyr == latest_yr and fmn == latest_mn)

        # Opening balance: weekly outlook end of prior month (includes overrides + sliders)
        prev_mn_3, prev_yr_3 = (fmn-1, fyr) if fmn > 1 else (12, fyr-1)
        wk_open    = shared_month_end_close(prev_yr_3, prev_mn_3)
        prev_fc    = fc[(fc['Year'] == prev_yr_3) & (fc['Month'] == prev_mn_3)]
        book2_open = float(prev_fc.iloc[0]['cash_uk']) if not prev_fc.empty else float(fc_row['cash_uk'])
        open_cash  = wk_open if wk_open is not None else book2_open

        # Forecast file targets — UK cash and total (UK + Ireland)
        fc_uk_close  = float(fc_row['cash_uk'])
        fc_ie_close  = float(fc_row['cash_ireland'])
        fc_close     = fc_uk_close + fc_ie_close   # total cash
        uk_req       = float(fc_row['uk_required'])
        headroom     = fc_close - uk_req            # total headroom (easy to move between UK/IE)
        hpct         = headroom / uk_req if uk_req > 0 else 0

        # Weekly outlook close for THIS month — from shared chain (includes overrides + sliders)
        wk_close_this   = shared_month_end_close(fyr, fmn)
        # Total headroom: add Ireland forecast close to weekly UK outlook
        wk_total_close  = (wk_close_this + fc_ie_close) if wk_close_this is not None else None
        wk_headroom     = (wk_total_close - uk_req) if wk_total_close is not None else None
        # ── Category definitions — explicit, no FX, no interco, no sweep ───────
        INFLOW_CORE  = ['AGENT RECEIPTS', 'FD RECEIPT', 'DIRECT RECEIPTS']
        # Fixed outflows: things that happen regardless and can't be deferred
        OUTFLOW_FIXED = ['FLIGHT COSTS', 'PAYROLL', 'TAX']
        # Controllable AP — the main lever
        OUTFLOW_AP    = ['AP COGS', 'AP OVH']
        # Deliberately excluded from all calculations:
        # FX TRADE OUT, INTERCO, SWEEP, CUSTOMER REFUNDS, OTHER COSTS
        # These are either treasury flows, non-cash, or unpredictable

        ly_inflow    = sum(safe_get(uk_in_spec,  fyr-1, fmn, c) for c in INFLOW_CORE   if c in uk_in_spec.columns)
        ly_apcogs    = abs(safe_get(uk_out_spec, fyr-1, fmn, 'AP COGS'))
        ly_apovh     = abs(safe_get(uk_out_spec, fyr-1, fmn, 'AP OVH'))
        ly_ap_total  = ly_apcogs + ly_apovh
        ly_flight    = abs(safe_get(uk_out_spec, fyr-1, fmn, 'FLIGHT COSTS'))
        ly_payroll   = abs(safe_get(uk_out_spec, fyr-1, fmn, 'PAYROLL'))
        ly_tax       = abs(safe_get(uk_out_spec, fyr-1, fmn, 'TAX'))
        ly_fixed     = ly_flight + ly_payroll + ly_tax

        days_in  = (pd.Timestamp(fyr, fmn, 1) + pd.offsets.MonthEnd(1)).day
        weeks_in = days_in / 7.0

        if is_part:
            rem       = days_in - latest_day
            frac_rem  = rem / days_in
            weeks_rem = rem / 7.0
            # Remaining inflows = LY full month × fraction remaining × slider adjustment
            rem_in    = ly_inflow * frac_rem * (1 + adj_receipts / 100)
            # Remaining fixed outflows (flight, payroll, tax — not AP)
            rem_fixed = ly_fixed  * frac_rem
            # AP already paid this month
            ap_mtd    = (abs(safe_get(uk_out_spec, fyr, fmn, 'AP COGS')) +
                         abs(safe_get(uk_out_spec, fyr, fmn, 'AP OVH')))
            # Total AP budget = opening + expected remaining inflows − forecast close − remaining fixed
            # AP slider adjusts the AP run rate benchmark (not the budget — budget is cash-driven)
            total_ap_budget = open_cash + rem_in - fc_close - rem_fixed
            allow_ap_remain = max(0, total_ap_budget - ap_mtd)
            allow_ap_per_wk = allow_ap_remain / max(weeks_rem, 0.5)
        else:
            rem       = days_in
            frac_rem  = 1.0
            weeks_rem = weeks_in
            # Full future month — apply sliders to full month estimate
            rem_in    = ly_inflow * (1 + adj_receipts / 100)
            rem_fixed = ly_fixed
            ap_mtd    = 0.0
            total_ap_budget  = open_cash + rem_in - fc_close - rem_fixed
            allow_ap_remain  = max(0, total_ap_budget)
            allow_ap_per_wk  = allow_ap_remain / weeks_in

        # LY AP adjusted by AP slider for comparison
        ly_ap_for_period_adj = ly_ap_total * frac_rem * (1 + adj_payments / 100)

        # CY actuals to date
        cy_inflow_mtd = sum(safe_get(uk_in_spec, fyr, fmn, c)
                            for c in INFLOW_CORE if c in uk_in_spec.columns)

        # How much to hold vs LY run rate
        ly_ap_for_period = ly_ap_total * frac_rem
        # hold_vs_ly uses slider-adjusted LY rate so AP slider affects the verdict
        hold_vs_ly       = max(0, ly_ap_for_period_adj - allow_ap_remain)
        ly_ap_per_wk     = ly_ap_total / weeks_in

        badge = "🚨 Breach" if hpct < 0 else ("⚠️ Monitor" if hpct < 0.20 else "✅ On track")

        with cols[ci]:
            st.markdown(f"### {MN[fmn]} {fyr}  {badge}")
            if is_part:
                st.caption(
                    f"**{latest_day} of {days_in} days elapsed · "
                    f"{rem} days / ~{weeks_rem:.1f} weeks remaining**"
                )

            # ── Cash position — forecast file vs weekly outlook ────────────────
            st.markdown("**Cash position**")
            wk_cl_fmt  = fmt(wk_total_close) if wk_total_close is not None else "—"
            wk_hr_fmt  = fmt(wk_headroom)    if wk_headroom    is not None else "—"
            wk_hr_pct  = f"({wk_headroom/uk_req:.0%})" if wk_headroom is not None and uk_req > 0 else ""
            # Compare total close: weekly UK + IE forecast close vs total forecast close
            var_close  = (wk_total_close - fc_close) if wk_total_close is not None else None
            var_fmt    = (f"**{'+' if var_close>=0 else ''}{fmt(var_close)}**"
                         if var_close is not None else "—")
            st.markdown(
                "| | Forecast target | Weekly outlook |\n"
                "|---|---|---|\n"
                f"| Opening (UK) | {fmt(book2_open)} | **{fmt(open_cash)}** |\n"
                f"| UK close | {fmt(fc_uk_close)} | {wk_cl_fmt} |\n"
                f"| Ireland close | {fmt(fc_ie_close)} | — |\n"
                f"| **Total close** | **{fmt(fc_close)}** | **{wk_cl_fmt}** |\n"
                f"| UK required | {fmt(uk_req)} | {fmt(uk_req)} |\n"
                f"| Total headroom | **{fmt(headroom)}** ({hpct:.0%}) | **{wk_hr_fmt}** {wk_hr_pct} |\n"
                f"| Variance vs fcst | | {var_fmt} |"
            )
            if var_close is not None and var_close < -500000:
                st.error(f"⚠️ Weekly outlook {fmt(abs(var_close))} **below** forecast target — "
                         f"adjust AP or receipts to close the gap")
            elif var_close is not None and var_close > 500000:
                st.success(f"✅ Weekly outlook {fmt(var_close)} **ahead** of forecast target")

            # ── Receipts ──────────────────────────────────────────────────────
            st.markdown("**Expected receipts** (Direct + Agent + FD, LY basis)")
            st.markdown(
                "| | LY full month | " + ("Remaining est." if is_part else "Full month") + " |\n"
                "|---|---|---|\n"
                f"| Core receipts | {fmt(ly_inflow)} | {fmt(rem_in)} |"
                + (f"\n| CY actual to date | | {fmt(cy_inflow_mtd)} |" if is_part else "")
            )

            # ── AP PAYMENT CAPACITY — THE KEY NUMBER ─────────────────────────
            st.markdown("---")
            st.markdown("### 🔑 AP payment capacity")

            if hold_vs_ly > 50000:
                st.error(
                    f"**Hold back {fmt(hold_vs_ly)} vs LY run rate** to hit forecast close"
                )
            else:
                st.success("✅ AP can run at LY pace — no hold back required")

            # Monthly budget
            m1, m2 = st.columns(2)
            adj_note_3mo = ""
            if adj_receipts != 0: adj_note_3mo += f" · receipts {adj_receipts:+d}%"
            if adj_payments != 0: adj_note_3mo += f" · AP {adj_payments:+d}%"
            m1.metric(
                "Total AP budget this " + ("month (remaining)" if is_part else "month"),
                fmt(allow_ap_remain),
                delta=f"LY rate: {fmt(ly_ap_for_period_adj)}" + (adj_note_3mo if adj_note_3mo else ""),
                delta_color="off"
            )
            m2.metric(
                "Weekly AP capacity",
                fmt(allow_ap_per_wk),
                delta=f"LY weekly avg: {fmt(ly_ap_per_wk)}",
                delta_color="off"
            )
            if is_part and ap_mtd > 0:
                st.caption(
                    f"AP COGS + OVH already paid this month: **{fmt(ap_mtd)}** · "
                    f"Remaining capacity: {fmt(allow_ap_remain)}"
                )

            # Breakdown: COGS vs OVH
            if ly_ap_total > 0:
                cogs_frac = ly_apcogs / ly_ap_total
                ovh_frac  = ly_apovh  / ly_ap_total
                allow_cogs = allow_ap_remain * cogs_frac
                allow_ovh  = allow_ap_remain * ovh_frac
                st.markdown(
                    "| | LY run rate | Budget (" +
                    ("remaining" if is_part else "full month") + ") |\n"
                    "|---|---|---|\n"
                    f"| AP COGS | {fmt(ly_apcogs * frac_rem)} | **{fmt(allow_cogs)}** |\n"
                    f"| AP OVH  | {fmt(ly_apovh  * frac_rem)} | **{fmt(allow_ovh)}** |"
                )

            # ── Other outflows ─────────────────────────────────────────────────
            st.markdown("---")
            with st.expander("Other outflows (for reference — excluded from AP calc)"):
                st.caption("FX trade out, interco and sweep are excluded — treasury flows, not operational.")
                est_label = 'Remaining est.' if is_part else 'Full month'
                st.markdown(
                    f"| Category | LY full month | {est_label} |\n"
                    f"|---|---|---|\n"
                    f"| Flight costs | {fmt(ly_flight)} | {fmt(ly_flight * frac_rem)} |\n"
                    f"| Payroll | {fmt(ly_payroll)} | {fmt(ly_payroll * frac_rem)} |\n"
                    f"| Tax | {fmt(ly_tax)} | {fmt(ly_tax * frac_rem)} |"
                )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — Opportunities & Risks
# ══════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    st.caption(
        "Total group cash flows (UK + Ireland). "
        "Actuals vs forecast-implied and prior year same month. "
        "Deviations >£500k are explained where drivers can be identified."
    )

    act_data = fc[fc['is_actual'] & ~fc['is_partial']]

    # ── Build monthly data for driver callouts (actual months only) ───────────
    monthly_data = []
    for _, row in act_data.iterrows():
        yr, mn  = int(row['Year']), int(row['Month'])
        uk_in   = safe_entity(entity_m, yr,   mn, 'UK',      'inflow')
        uk_out  = safe_entity(entity_m, yr,   mn, 'UK',      'outflow')
        ie_in   = safe_entity(entity_m, yr,   mn, 'Ireland', 'inflow')
        ie_out  = safe_entity(entity_m, yr,   mn, 'Ireland', 'outflow')
        tot_net = uk_in + uk_out + ie_in + ie_out
        ly_uk_in  = safe_entity(entity_m, yr-1, mn, 'UK',      'inflow')
        ly_uk_out = safe_entity(entity_m, yr-1, mn, 'UK',      'outflow')
        ly_ie_in  = safe_entity(entity_m, yr-1, mn, 'Ireland', 'inflow')
        ly_ie_out = safe_entity(entity_m, yr-1, mn, 'Ireland', 'outflow')
        ly_net   = ly_uk_in + ly_uk_out + ly_ie_in + ly_ie_out
        fc_imp   = float(row.get('fc_implied_net', 0) or 0)
        yoy_var  = tot_net - ly_net if abs(ly_net) > 100000 else None
        drivers  = []
        for dlabel, cat, spec, is_in in [
            ('FD receipts',    'FD RECEIPT',    uk_in_spec,  True),
            ('Agent receipts', 'AGENT RECEIPTS',uk_in_spec,  True),
            ('Direct receipts','DIRECT RECEIPTS',uk_in_spec, True),
            ('AP COGS',        'AP COGS',       uk_out_spec, False),
            ('AP OVH',         'AP OVH',        uk_out_spec, False),
            ('Flight costs',   'FLIGHT COSTS',  uk_out_spec, False),
            ('Payroll',        'PAYROLL',       uk_out_spec, False),
        ]:
            cy_v = safe_get(spec, yr,   mn, cat)
            ly_v = safe_get(spec, yr-1, mn, cat)
            if not is_in: cy_v = abs(cy_v); ly_v = abs(ly_v)
            if ly_v > 100000 and abs((cy_v-ly_v)/ly_v) > 0.20:
                drivers.append((dlabel, (cy_v-ly_v)/ly_v,
                                'up' if cy_v > ly_v else 'down', cy_v, ly_v))
        monthly_data.append({
            'label': f"{MN[mn]} {yr}", 'yr': yr, 'mn': mn,
            'tot_net': tot_net, 'ly_net': ly_net, 'fc_imp': fc_imp,
            'yoy_var': yoy_var, 'drivers': drivers,
            'uk_in': uk_in, 'uk_out': uk_out, 'ie_in': ie_in, 'ie_out': ie_out,
        })

    # ── Extended chart data: actuals + forecast ───────────────────────────────
    fc_ext = fc.copy()
    fc_ext['fc_implied_net'] = fc_ext['cash_uk'].diff()
    all_labels_c, all_act_c, all_fc_c, all_ly_c, all_is_act_c = [], [], [], [], []
    for _, row in fc_ext.iterrows():
        yr, mn  = int(row['Year']), int(row['Month'])
        is_act  = bool(row['is_actual']) and not bool(row.get('is_partial', False))
        fc_imp  = float(row.get('fc_implied_net', 0) or 0)
        ly_net  = (safe_entity(entity_m, yr-1, mn, 'UK', 'inflow') +
                   safe_entity(entity_m, yr-1, mn, 'UK', 'outflow') +
                   safe_entity(entity_m, yr-1, mn, 'Ireland', 'inflow') +
                   safe_entity(entity_m, yr-1, mn, 'Ireland', 'outflow')) / 1e6
        if is_act:
            act_net = (safe_entity(entity_m, yr, mn, 'UK', 'inflow') +
                       safe_entity(entity_m, yr, mn, 'UK', 'outflow') +
                       safe_entity(entity_m, yr, mn, 'Ireland', 'inflow') +
                       safe_entity(entity_m, yr, mn, 'Ireland', 'outflow')) / 1e6
        else:
            act_net = None
        all_labels_c.append(f"{MN[mn]} {yr}")
        all_act_c.append(act_net)
        all_fc_c.append(fc_imp / 1e6 if abs(fc_imp) > 0 else None)
        all_ly_c.append(ly_net if abs(ly_net) > 0.1 else None)
        all_is_act_c.append(is_act)

    if not monthly_data:
        st.info("No complete actual months available yet.")
    else:
        # ── Chart 1: Total net cash — actuals + forecast vs LY ───────────────
        st.subheader("Monthly net cash — total group (UK + Ireland) · actuals + forecast")
        fig_or  = go.Figure()
        bar_vals   = [all_act_c[i] if all_act_c[i] is not None else all_fc_c[i]
                      for i in range(len(all_labels_c))]
        bar_colors = []
        for i in range(len(all_labels_c)):
            v = bar_vals[i] or 0
            if all_act_c[i] is not None:
                bar_colors.append(GREEN if v >= 0 else RED)
            else:
                bar_colors.append('rgba(28,20,100,0.35)' if v >= 0 else 'rgba(226,75,74,0.35)')
        fig_or.add_trace(go.Bar(
            x=all_labels_c, y=bar_vals,
            name='Actual (solid) / Forecast (faded)',
            marker_color=bar_colors, opacity=0.9))
        fig_or.add_trace(go.Scatter(
            x=all_labels_c, y=all_ly_c,
            name='Prior year same month',
            line=dict(color=AMBER, width=2, dash='dot'),
            mode='lines+markers', marker=dict(size=4), connectgaps=False))
        # Boundary line — use shape + annotation separately (Plotly 6 compatibility)
        last_act_lbl = [l for l, a in zip(all_labels_c, all_is_act_c) if a]
        if last_act_lbl:
            x_idx = all_labels_c.index(last_act_lbl[-1])
            fig_or.add_shape(type='line',
                             xref='x', yref='paper',
                             x0=last_act_lbl[-1], x1=last_act_lbl[-1],
                             y0=0, y1=1,
                             line=dict(color='rgba(0,0,0,0.25)', dash='dash', width=1))
            fig_or.add_annotation(
                x=last_act_lbl[-1], y=1, xref='x', yref='paper',
                text='← actual | forecast →',
                showarrow=False, font=dict(size=9, color='rgba(0,0,0,0.4)'),
                xanchor='left', yanchor='top')
        fig_or.add_hline(y=0, line_color='rgba(0,0,0,0.2)', line_width=0.8)
        fig_or.update_layout(
            height=320, plot_bgcolor='white', paper_bgcolor='white',
            yaxis=dict(tickformat=',.1f', ticksuffix='m', gridcolor=GREY, title='£m'),
            xaxis=dict(tickangle=45),
            legend=dict(orientation='h', yanchor='bottom', y=1.02, font=dict(size=10)),
            margin=dict(l=0, r=0, t=40, b=0))
        st.plotly_chart(fig_or, use_container_width=True)

        # ── Chart 2: YoY variance — actual months only ───────────────────────
        st.subheader("Year-on-year variance — actual months (CY vs prior year same month)")
        labels_m   = [d['label'] for d in monthly_data]
        yoy_clean  = [d['yoy_var']/1e6 if d['yoy_var'] is not None else 0
                      for d in monthly_data]
        fig_yoy = go.Figure(go.Bar(
            x=labels_m, y=yoy_clean,
            marker_color=[GREEN if v >= 0 else RED for v in yoy_clean],
            text=[f"{'+' if v >= 0 else ''}{v:.1f}m" for v in yoy_clean],
            textposition='outside'))
        fig_yoy.add_hline(y=0, line_color='rgba(0,0,0,0.2)', line_width=0.8)
        fig_yoy.update_layout(
            height=240, plot_bgcolor='white', paper_bgcolor='white',
            yaxis=dict(tickformat=',.1f', ticksuffix='m', gridcolor=GREY, title='£m vs LY'),
            margin=dict(l=0, r=0, t=10, b=0), showlegend=False)
        st.plotly_chart(fig_yoy, use_container_width=True)

        # ── Driver callouts ───────────────────────────────────────────────────
        st.subheader("Variance drivers — where known (>20% move vs prior year)")

        has_drivers = any(d['drivers'] for d in monthly_data)
        large_yoy   = [d for d in monthly_data if d['yoy_var'] is not None and abs(d['yoy_var']) > 0.5e6]

        if not large_yoy:
            st.info("No month shows a YoY variance above £500k — flows are broadly in line with prior year.")
        else:
            for d in large_yoy:
                yov  = d['yoy_var']
                sign = "🟢 above" if yov > 0 else "🔴 below"
                with st.expander(
                    f"{d['label']}  —  {sign} prior year by **{fmt(abs(yov))}**"
                    + (f"  ·  {len(d['drivers'])} driver(s) identified" if d['drivers'] else "  ·  no specific driver identified")
                ):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("**CY actuals**")
                        st.markdown(
                            "| | |\n|---|---|\n"
                            + f"| UK inflow | {fmt(d['uk_in'])} |\n"
                            + f"| UK outflow | {fmt(d['uk_out'])} |\n"
                            + f"| Ireland inflow | {fmt(d['ie_in'])} |\n"
                            + f"| Ireland outflow | {fmt(d['ie_out'])} |\n"
                            + f"| **Total net** | **{fmt(d['tot_net'])}** |"
                        )
                    with c2:
                        st.markdown("**Prior year same month**")
                        ly_uk_in  = safe_entity(entity_m, d['yr']-1, d['mn'], 'UK',      'inflow')
                        ly_uk_out = safe_entity(entity_m, d['yr']-1, d['mn'], 'UK',      'outflow')
                        ly_ie_in  = safe_entity(entity_m, d['yr']-1, d['mn'], 'Ireland', 'inflow')
                        ly_ie_out = safe_entity(entity_m, d['yr']-1, d['mn'], 'Ireland', 'outflow')
                        st.markdown(
                            "| | |\n|---|---|\n"
                            + f"| UK inflow | {fmt(ly_uk_in)} |\n"
                            + f"| UK outflow | {fmt(ly_uk_out)} |\n"
                            + f"| Ireland inflow | {fmt(ly_ie_in)} |\n"
                            + f"| Ireland outflow | {fmt(ly_ie_out)} |\n"
                            + f"| **Total net** | **{fmt(d['ly_net'])}** |"
                        )
                    if d['drivers']:
                        st.markdown("**Key drivers (>20% vs LY):**")
                        for dlabel, chg, direction, cy_v, ly_v in d['drivers']:
                            arrow = "⬆️" if direction == 'up' else "⬇️"
                            st.markdown(
                                f"- {arrow} **{dlabel}**: {fmt(cy_v)} vs LY {fmt(ly_v)} "
                                f"({chg:+.0%})"
                            )
                    else:
                        st.caption("No single category moved >20% vs prior year. "
                                   "The variance may reflect timing differences or multiple small movements.")

        # ── Trend summary & forward look ──────────────────────────────────────
        st.divider()
        act_nets_arr = np.array([d['tot_net'] for d in monthly_data])
        ly_nets_arr  = np.array([d['ly_net']  for d in monthly_data if abs(d['ly_net']) > 100000])
        fc_nets_arr  = np.array([d['fc_imp']  for d in monthly_data if abs(d['fc_imp'])  > 100000])

        trend_vs_ly = float(np.mean(act_nets_arr - np.array([d['ly_net'] for d in monthly_data])))
        trend_vs_fc = float(np.mean([d['tot_net'] - d['fc_imp'] for d in monthly_data if abs(d['fc_imp']) > 100000]))                       if any(abs(d['fc_imp']) > 100000 for d in monthly_data) else 0.0

        breach_fc = fc[~fc['is_actual'] & (fc['uk_headroom'] < 0)]
        tight_fc  = fc[~fc['is_actual'] & (fc['uk_headroom'] >= 0) & (fc['headroom_pct'] < 0.15)]

        tc1, tc2 = st.columns(2)
        with tc1:
            st.markdown("### Opportunities")
            if trend_vs_ly > 0.2e6:
                st.success(f"📈 Running **{fmt(trend_vs_ly)}/month ahead** of prior year on average.")
            if trend_vs_fc > 0.2e6:
                st.success(f"📈 Running **{fmt(trend_vs_fc)}/month ahead** of forecast file on average.")
            if not breach_fc.empty and trend_vs_fc > 0:
                st.info(
                    f"💡 If positive trend continues, forecast breach months could see "
                    f"~{fmt(trend_vs_fc * len(breach_fc))} aggregate headroom improvement."
                )
            if trend_vs_ly <= 0.2e6 and trend_vs_fc <= 0.2e6:
                st.info("No material positive trend vs prior year or forecast detected.")
        with tc2:
            st.markdown("### Risks")
            if trend_vs_ly < -0.2e6:
                st.error(f"📉 Running **{fmt(abs(trend_vs_ly))}/month below** prior year on average.")
            if trend_vs_fc < -0.2e6:
                st.error(f"📉 Running **{fmt(abs(trend_vs_fc))}/month below** forecast file.")
            if not breach_fc.empty:
                st.error("🚨 **Forecast breach months:** " + ", ".join(
                    f"{MN[int(r['Month'])]} {int(r['Year'])} ({fmt(r['uk_headroom'])})"
                    for _, r in breach_fc.iterrows()))
            if not tight_fc.empty:
                st.warning("⚠️ **Tight months (<15% headroom):** " + ", ".join(
                    f"{MN[int(r['Month'])]} {int(r['Year'])} ({r['headroom_pct']:.0%})"
                    for _, r in tight_fc.iterrows()))
            if trend_vs_ly >= -0.2e6 and trend_vs_fc >= -0.2e6 and breach_fc.empty:
                st.success("No material negative trends or forecast breaches detected.")

    st.divider()
    buf = io.BytesIO()
    fc.reset_index().to_excel(buf, index=False, sheet_name='Dashboard Data')
    buf.seek(0)
    st.download_button("⬇️ Download dashboard data (.xlsx)", data=buf,
        file_name='shg_cash_dashboard.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — Weekly 4+13 Outlook
# ══════════════════════════════════════════════════════════════════════════════
with tabs[6]:
    st.caption(
        f"All amounts £GBP equivalent at budget rates "
        f"(EUR×{eur_rate:.3f} · USD×{usd_rate:.3f} · CAD×{cad_rate:.3f}). "
        f"Actual: last 4 weeks from bank data. "
        f"Forecast: prior year same ISO week × YoY trend. "
        f"Interco: net UK↔Ireland only — intra-UK transfers cancel. "
        f"OD interest: actual net for actuals / £11k/week forecast."
    )

    if 'weekly_ov' not in st.session_state:
        st.session_state.weekly_ov = {}

    n_actuals    = 4
    actual_weeks = weekly_raw.index[-n_actuals:].tolist()
    all_weeks    = actual_weeks + fc_weeks
    N_W          = len(all_weeks)

    # Opening balance: use actual cash position if entered, else fall back to Book2
    first_wk     = actual_weeks[0]
    book2_before = fc[fc.index < first_wk]
    book2_open   = float(book2_before.iloc[-1]['cash_uk']) / 1000 if not book2_before.empty else 0.0
    if use_actual_pos:
        # Use total UK GBP equiv entered in sidebar (SHT + TMD)
        # Divide by 1000 as weekly table works in £k
        open_base = pos_total_uk / 1000
    else:
        open_base = book2_open

    ROW_SPECS = [
        ('DIRECT RECEIPTS', 'DIRECT RECEIPTS',  False),
        ('AGENT RECEIPTS',  'AGENT RECEIPTS',   False),
        ('FD RECEIPT',      'FD RECEIPT',        False),
        ('CUSTOMER REFUND', 'CUSTOMER REFUNDS',  False),
        ('TUI RECEIPT',     'TUI RECEIPT',       False),
        ('OTHER RECEIPT',   'OTHER RECEIPTS',    True),  # blank — unpredictable
        ('FX TRADE IN',     'FX TRADE IN',       True),  # blank — treasury/lumpy
        ('INTERCO (net)',   'INTERCO',           True),  # blank — net zero assumed
        ('OD INTEREST',     '_OD',               False),
        ('AP COGS',         'AP COGS',           False),
        ('AP OVH',          'AP OVH',            False),
        ('PAYROLL',         'PAYROLL',           False),
        ('TAX',             'TAX',               False),
        ('OTHER CASH OUT',  'OTHER COSTS',       True),  # blank — unpredictable
        ('FX TRADE OUT',    'FX TRADE OUT',      True),  # blank — treasury/lumpy
        ('FLIGHT COSTS',    'FLIGHT COSTS',      False),
    ]
    RECEIPT_LABELS = ['DIRECT RECEIPTS','AGENT RECEIPTS','FD RECEIPT','CUSTOMER REFUND',
                      'TUI RECEIPT','OTHER RECEIPT','FX TRADE IN','INTERCO (net)','OD INTEREST']
    PAYMENT_LABELS = ['AP COGS','AP OVH','PAYROLL','TAX','OTHER CASH OUT','FX TRADE OUT','FLIGHT COSTS']

    # Rows affected by the sidebar adjustment sliders
    # (only core receipts and AP — not FX, interco, payroll, flight, tax)
    SLIDER_RECEIPT_ROWS = {'DIRECT RECEIPTS', 'AGENT RECEIPTS', 'FD RECEIPT', 'TUI RECEIPT'}
    SLIDER_AP_ROWS      = {'AP COGS', 'AP OVH'}

    # All forecast weeks — sliders apply across the full 13-week outlook
    current_month_fc_weeks = set(range(n_actuals, N_W))

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
                if spec == '_OD':
                    v = float(od_weekly.get(wk, 0)) / 1000
                elif spec in weekly_raw.columns and wk in weekly_raw.index:
                    v = float(weekly_raw.loc[wk, spec]) / 1000
                else:
                    v = 0.0
            else:
                if blank_fc:
                    v = 0.0
                elif spec == '_OD':
                    v = 11.0
                elif spec == 'PAYROLL':
                    v = fc_base.get('PAYROLL', [0]*13)[i - n_actuals] / 1000
                else:
                    v = fc_base.get(spec, [0]*13)[i - n_actuals] / 1000 if spec else 0.0

                # Apply slider adjustments across all 13 forecast weeks
                if i in current_month_fc_weeks:
                    if label in SLIDER_RECEIPT_ROWS and adj_receipts != 0:
                        v = v * (1 + adj_receipts / 100)
                    elif label in SLIDER_AP_ROWS and adj_payments != 0:
                        v = v * (1 + adj_payments / 100)

            row.append(round(v, 1))
        data[label] = row

    totR  = [sum(data[l][i] for l in RECEIPT_LABELS) for i in range(N_W)]
    totP  = [sum(data[l][i] for l in PAYMENT_LABELS) for i in range(N_W)]
    net   = [totR[i] + totP[i] for i in range(N_W)]
    opens = [0.0] * N_W;  closes = [0.0] * N_W
    opens[0]  = open_base
    closes[0] = open_base + net[0]
    for i in range(1, N_W):
        opens[i]  = closes[i-1]
        closes[i] = closes[i-1] + net[i]

    # Note: shared_month_end_close uses the pre-tab chain (_shared_closes_k)
    # which is computed before any tab renders using the same overrides + sliders.
    # No session state sync needed.

    # ── Override expander ─────────────────────────────────────────────────────
    ovr_labels = ['AP COGS','AP OVH','PAYROLL','FD RECEIPT','AGENT RECEIPTS',
                  'FX TRADE IN','FX TRADE OUT','FLIGHT COSTS','INTERCO (net)',
                  'OTHER RECEIPT','OTHER CASH OUT']
    # ⬜ = blank by default in forecast (FX, interco, other — enter only if known)

    _, col_btn = st.columns([8, 1])
    with col_btn:
        if st.button("Reset overrides"):
            st.session_state.weekly_ov = {}
            st.rerun()

    with st.expander("Edit forecast — nearest 4 weeks (−£9,999,999 to +£9,999,999)", expanded=False):
        st.caption(
            "All values in £. 🔒 Payroll auto-detected from cycle — still editable. "
            "🔵 Interco blank by default — enter known UK↔Ireland transfers only."
        )
        cols_ov = st.columns(4)
        spec_map = {l: s for l, s, _ in ROW_SPECS}
        for fi in range(min(4, len(fc_weeks))):
            i  = fi + n_actuals
            wk = all_weeks[i]
            with cols_ov[fi]:
                st.markdown(f"**{wk.strftime('%d %b %Y')}**")
                for label in ovr_labels:
                    spec    = spec_map.get(label, label)
                    ov_key  = f"{label}_{i}"
                    is_lk   = label in WEEKLY_LOCK
                    is_ico  = 'INTERCO' in label
                    # fc_def: what the forecast would be without override (in £)
                    # blank_fc rows default to 0; others use fc_base
                    spec_blank = next((b for l2, s2, b in ROW_SPECS if l2 == label), False)
                    if spec_blank:
                        fc_def = 0.0
                    elif is_lk:
                        fc_def = float(fc_base.get('PAYROLL', [0]*13)[fi])
                    else:
                        fc_def = float(fc_base.get(spec, [0]*13)[fi] if spec in fc_base else 0.0)

                    # Use stored override if exists, else forecast default
                    curr     = float(st.session_state.weekly_ov.get(ov_key, fc_def))
                    hint     = "🔒 " if is_lk else ("🔵 " if is_ico else ("⬜ " if spec_blank else ""))
                    safe_val = float(np.clip(curr, -9_999_999.0, 9_999_999.0))
                    new_v    = st.number_input(
                        f"{hint}{label}",
                        value=safe_val,
                        min_value=-9_999_999.0, max_value=9_999_999.0,
                        step=1000.0, format="%.0f",
                        key=f"wov_{label}_{fi}")
                    # Save if different from forecast default (use tolerance of £500)
                    if abs(new_v - fc_def) > 500:
                        st.session_state.weekly_ov[ov_key] = new_v
                    elif ov_key in st.session_state.weekly_ov:
                        del st.session_state.weekly_ov[ov_key]

    # ── Display table ─────────────────────────────────────────────────────────
    def fkw(v, dash=True):
        if v == 0 and dash: return '—'
        return ('−' if v < 0 else '') + '£' + f"{abs(round(v)):,}k"

    wk_hdrs = [
        f"{'W' if i < n_actuals else '~W'}{all_weeks[i].isocalendar()[1]}\n{all_weeks[i].strftime('%d/%m')}"
        for i in range(N_W)
    ]

    display_rows = []

    def dr(label, vals, kind='data', indent=False):
        row = {'Row': ('  ' + label if indent else label)}
        for i, v in enumerate(vals):
            row[wk_hdrs[i]] = fkw(v, kind not in ('total', 'balance'))
        display_rows.append({'_kind': kind, '_label': label, **row})

    dr('Opening balance', opens, 'balance')
    display_rows.append({'_kind': 'header', '_label': 'RECEIPTS', 'Row': 'RECEIPTS',
                          **{h: '' for h in wk_hdrs}})
    for label in RECEIPT_LABELS:
        dr(label, data[label], indent=True)
    dr('Total receipts', totR, 'total')
    display_rows.append({'_kind': 'sep', '_label': '', 'Row': '', **{h: '' for h in wk_hdrs}})
    display_rows.append({'_kind': 'header', '_label': 'PAYMENTS', 'Row': 'PAYMENTS',
                          **{h: '' for h in wk_hdrs}})
    for label in PAYMENT_LABELS:
        dr(label, data[label], indent=True)
    dr('Total payments', totP, 'total')
    display_rows.append({'_kind': 'sep', '_label': '', 'Row': '', **{h: '' for h in wk_hdrs}})
    dr('Net cash',        net,    'total')
    dr('Closing balance', closes, 'balance')

    # ── Book2 month-end targets — show in the last week of each calendar month ─
    # For each week column, find if it's the last week whose WeekStart falls in
    # a given month, then show forecast file close / client money / headroom.
    # Only shown once per month (the week that closes out that month).

    def get_book2_for_week(wk):
        """Return (fc_close_k, client_money_k, headroom_k) for the month containing wk,
        but only if wk is the LAST week in that calendar month among all_weeks."""
        mn = wk.month; yr = wk.year
        # Last week in all_weeks that falls in this month
        wks_in_month = [w for w in all_weeks if w.month == mn and w.year == yr]
        if not wks_in_month or wk != max(wks_in_month):
            return None, None, None
        # Look up Book2
        book2_row = fc[(fc['Year'] == yr) & (fc['Month'] == mn)]
        if book2_row.empty:
            return None, None, None
        row = book2_row.iloc[0]
        # Total cash close = UK + Ireland
        fc_close_k  = (float(row['cash_uk']) + float(row['cash_ireland'])) / 1000
        cm_k        = float(row['client_money'])  / 1000
        headroom_k  = float(row['uk_headroom'])   / 1000
        return fc_close_k, cm_k, headroom_k

    b2_fc_close  = []
    b2_cm        = []
    b2_headroom  = []
    b2_variance  = []   # closing balance (running) minus forecast month-end target
    for i, wk in enumerate(all_weeks):
        fc_c, cm_c, hr_c = get_book2_for_week(wk)
        b2_fc_close.append(fc_c)
        b2_cm.append(cm_c)
        b2_headroom.append(hr_c)
        # Variance: only meaningful where Book2 target exists
        if fc_c is not None:
            b2_variance.append(closes[i] - fc_c)
        else:
            b2_variance.append(None)

    # Format helpers for these rows (show '—' where no month-end target)
    def fkw_me(v):
        if v is None: return ''     # blank — not end of month
        return fkw(v, dash=False)

    def fkw_var(v):
        if v is None: return ''
        if abs(v) < 1: return '—'
        prefix = '+' if v > 0 else ''
        return f"{prefix}{fkw(v, dash=False)}"

    # Add separator then month-end tracking rows
    display_rows.append({'_kind': 'sep', '_label': '', 'Row': '', **{h: '' for h in wk_hdrs}})
    display_rows.append({'_kind': 'header', '_label': 'MONTH-END TRACKING',
                         'Row': 'MONTH-END TRACKING (Forecast)', **{h: '' for h in wk_hdrs}})

    def dr_me(label, vals_fn, kind='forecast'):
        row = {'Row': label}
        for i in range(N_W):
            row[wk_hdrs[i]] = vals_fn(i)
        display_rows.append({'_kind': kind, '_label': label, **row})

    dr_me('Forecast file close',   lambda i: fkw_me(b2_fc_close[i]),  'forecast')
    dr_me('Client money',           lambda i: fkw_me(b2_cm[i]),         'forecast')
    dr_me('UK headroom vs req',     lambda i: fkw_me(b2_headroom[i]),   'forecast')
    dr_me('Variance vs forecast close',lambda i: fkw_var(b2_variance[i]),  'variance')

    df_display = pd.DataFrame(
        [{k: v for k, v in r.items() if not k.startswith('_')} for r in display_rows]
    ).set_index('Row')

    # Render weekly table — plain dataframe, no custom styler (avoids pandas KeyError)
    # Highlight locked/override rows with emoji prefix in the Row label instead
    df_plain = df_display.copy().reset_index()
    df_plain['Row'] = df_plain['Row'].apply(lambda x: x.strip())
    st.dataframe(
        df_plain,
        use_container_width=True,
        hide_index=True,
        height=min(80 + len(display_rows) * 32, 900),
    )

    # ── Charts ────────────────────────────────────────────────────────────────
    ch1, ch2 = st.columns(2)
    wk_lbl = [all_weeks[i].strftime('%d %b') for i in range(N_W)]

    with ch1:
        st.caption("Closing balance (£k) — solid = actual, dashed = forecast")
        fig_bal = go.Figure()
        fig_bal.add_trace(go.Scatter(
            x=wk_lbl[:n_actuals], y=closes[:n_actuals],
            name='Actual', line=dict(color=BLUE, width=2.5),
            mode='lines+markers', marker=dict(size=4)))
        fig_bal.add_trace(go.Scatter(
            x=wk_lbl[n_actuals-1:], y=closes[n_actuals-1:],
            name='Forecast', line=dict(color=BLUE, width=1.5, dash='dot'),
            mode='lines+markers', marker=dict(size=3, symbol='circle-open')))
        fig_bal.update_layout(height=220, plot_bgcolor='white', paper_bgcolor='white',
            yaxis=dict(tickformat=',.0f', ticksuffix='k', gridcolor=GREY),
            legend=dict(orientation='h', y=1.1, font=dict(size=10)),
            margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig_bal, use_container_width=True)

    with ch2:
        st.caption("Weekly receipts vs payments (£k)")
        bc = [BLUE if i < n_actuals else '#7B72C8' for i in range(N_W)]
        pc = [RED  if i < n_actuals else '#F09595' for i in range(N_W)]
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(x=wk_lbl, y=totR,
            name='Receipts', marker_color=bc, offsetgroup=0))
        fig_bar.add_trace(go.Bar(x=wk_lbl, y=[abs(v) for v in totP],
            name='Payments', marker_color=pc, offsetgroup=1))
        fig_bar.update_layout(barmode='group', height=220,
            plot_bgcolor='white', paper_bgcolor='white',
            yaxis=dict(tickformat=',.0f', ticksuffix='k', gridcolor=GREY),
            legend=dict(orientation='h', y=1.1, font=dict(size=10)),
            margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig_bar, use_container_width=True)

    # ── Closing balance gauges ────────────────────────────────────────────────
    st.caption("Forecast closing balance — 🟢 net inflow week  🟡 positive but below opening  🔴 negative")
    # Show first 13 forecast weeks in gauges (cap at 13 columns for readability)
    gauge_weeks = fc_weeks[:13]
    n_gauges = len(gauge_weeks)
    if n_gauges > 0:
        g_cols = st.columns(n_gauges)
        for fi, wk in enumerate(gauge_weeks):
            i  = fi + n_actuals
            if i >= len(closes): continue
            cl = closes[i]; op = opens[i]
            signal = "🟢" if cl > op else ("🟡" if cl > 0 else "🔴")
            with g_cols[fi]:
                st.metric(
                    label=wk.strftime('%d %b'),
                    value=f"£{cl:,.0f}k",
                    delta=f"open: £{op:,.0f}k · {signal}",
                    delta_color="off")

    # ── Export ────────────────────────────────────────────────────────────────
    st.divider()
    exp_rows = {'Opening balance': opens}
    for l in RECEIPT_LABELS: exp_rows[l] = data[l]
    exp_rows['Total receipts']  = totR
    for l in PAYMENT_LABELS:   exp_rows[l] = data[l]
    exp_rows['Total payments']  = totP
    exp_rows['Net cash']        = net
    exp_rows['Closing balance'] = closes
    col_names = [
        f"{'ACT_' if i < n_actuals else 'FC_'}{all_weeks[i].strftime('%d%b%Y')}"
        for i in range(N_W)
    ]
    export_df = pd.DataFrame(exp_rows, index=col_names).T
    buf2 = io.BytesIO()
    export_df.to_excel(buf2, sheet_name='Weekly Cash Flow')
    buf2.seek(0)
    st.download_button("⬇️ Export weekly cash flow (.xlsx)", data=buf2,
        file_name=f"weekly_cashflow_{latest_date.strftime('%Y%m%d')}.xlsx",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ══════════════════════════════════════════════════════════════════════════════
# TAB 8 — How it works
# ══════════════════════════════════════════════════════════════════════════════
with tabs[7]:
    st.header("How this dashboard works")
    st.caption("Reference guide for the banking and treasury team.")

    with st.expander("📂 Data inputs", expanded=True):
        st.markdown("""
**Two files are uploaded each session via the sidebar — nothing is stored permanently.**

| File | What it contains | Format |
|---|---|---|
| Bank transactions | Daily transaction-level data for all accounts | Excel — `Data Sheet` tab |
| Forecast / client money file | Monthly forecast closing balances + client money | Excel — Row 1: dates, Row 2: client money, Row 3: UK cash, Row 4: Ireland cash |

**Optional sidebar inputs (entered manually each session):**
- **SHT / SHGI / TMD actual positions** — GBP equivalent from the daily cash balance sheet. Overrides the forecast file as the opening balance anchor.
- **Client money override** — enter the expected month-end client money figure if it differs from the forecast file.
- **FX budget rates** — EUR, USD and CAD to GBP conversion rates applied to all non-GBP transactions.
- **Remaining flows sliders** — adjust expected receipts and AP payments vs prior year run rate for the rest of the current month.
""")

    with st.expander("🏦 Account mapping"):
        st.markdown("""
Account → entity classification is hard-coded in the app. No upload required.

| Entity | Accounts |
|---|---|
| **UK** | SHTL GBP, SHTL EUR, SHTL USD, SHTL PAY GBP, SHTL PAY, TMOOD GBP, TMOOD EUR, TMOOD USD, TMOOD CAD, HJT GBP, HJT AH GBP |
| **Ireland** | SHGI EUR, SHGI GBP, SHGI USD, SHGI CAD, AHD EURO CURRENT ACCOUNT, BOA |

If a new account appears in the bank feed that isn't mapped, the app shows a yellow warning and defaults it to Ireland. Contact the developer to add it to the mapping.
""")

    with st.expander("💱 FX conversion"):
        st.markdown("""
All non-GBP transactions are converted to GBP equivalent using the **budget rates** set in the sidebar.

- Default rates are based on spot rates as at 23 Jun 2026 (EUR 0.86, USD 0.76, CAD 0.53)
- Rates can be adjusted at any time — the entire dashboard recalculates immediately
- FX Trade In / Out amounts are already captured in GBP in the GBP accounts, so there is no double-counting
- Overnight deposits are excluded from all cash flow analysis — they go out and come back each day and net to approximately zero per week (the small positive residual shown as OD Interest represents the interest earned)
""")

    with st.expander("🔁 Interco netting"):
        st.markdown("""
The weekly cash flow table shows **net UK↔Ireland interco only**.

Within the UK entity group, transfers between SHTL, TMOOD and HJT cancel each other out when viewed on a consolidated basis. Only the net movement of cash between UK and Ireland entities appears in the table.

In the forecast, interco defaults to zero (the assumption being that intra-group transfers will net to nil over the forecast period). If you know a specific transfer is planned, enter it in the override panel.
""")

    with st.expander("📅 Forecast methodology"):
        st.markdown("""
**Monthly forecast** (Cash vs Forecast, Compliance tabs):
- Source: forecast file uploaded each session
- Extended automatically to December 2027 minimum, plus 12 rolling months beyond the latest bank data date
- Missing months beyond the forecast file are filled by carrying forward the same month one year prior (seasonal carry-forward)

**Weekly forecast** (4+13 Outlook tab):
- For each forecast week, the app looks up the same ISO week number from the prior year in the actual bank data
- This preserves the seasonal pattern (e.g. high FD receipts in summer, payroll timing)
- If the exact prior year week is missing, the nearest available week (±2 weeks) is used
- Payroll is detected automatically from the payment cycle pattern in actuals
- Interco is left blank (zero) in the forecast — editable if needed

**No trend multipliers are applied.** Prior year same week is used directly. This avoids distortion from the short data history (which only starts May 2025).
""")

    with st.expander("🛡️ Client money compliance"):
        st.markdown("""
The regulatory requirement is that **70% of client money must be held in UK GBP cash at all times**.

- Client money figure: from forecast file (or overridden in sidebar)
- UK required = client money × 70% (percentage adjustable in sidebar)
- UK headroom = UK cash − UK required
- Status: ✅ Compliant (headroom ≥ 20% of requirement) · ⚠️ Caution (0–20%) · 🚨 Breach (<0%)

The compliance tab shows the full forecast horizon so breach months can be identified and planned for in advance.
""")

    with st.expander("🎯 3-Month Focus — AP payment capacity"):
        st.markdown("""
The 3-Month Focus tab answers the key question: **how much can we pay suppliers this month and still hit the forecast closing balance?**

**Calculation:**
```
Max AP this month = Opening balance + Expected receipts − Forecast close − Fixed outflows
```

- **Expected receipts** = prior year Direct Receipts + Agent Receipts + FD Receipts (core receipts only — FX and interco excluded as they are unpredictable)
- **Fixed outflows** = prior year Flight Costs + Payroll + Tax (pro-rated for remaining days if current month)
- **AP budget** = what's left after fixed outflows and required receipts land

For the current (partial) month, amounts already paid in AP this month are subtracted from the budget to show the **remaining AP capacity**.

The "hold back vs LY" figure tells you exactly how much less AP to release compared to last year's same month run rate.
""")

    with st.expander("⚡ Opportunities & Risks"):
        st.markdown("""
This tab compares **actual monthly net cash flows** against the **implied flow from the forecast file** (month-on-month change in forecast closing balance).

A deviation of more than £500k triggers an insight, with reasons identified by comparing actuals to prior year in key categories:
- FD receipts, FX inflows (inflow drivers)
- AP COGS, FX outflows (outflow drivers)

**Trend direction** — if actuals have been consistently above or below the implied forecast, the trend is highlighted and projected forward to show the cumulative impact on forecast breach months.
""")

    with st.expander("🔄 Updating the dashboard"):
        st.markdown("""
**Monthly refresh:**
1. Download latest bank transactions Excel from your banking system
2. Update the forecast file with the latest actual closing balances and forward forecast
3. Open the app → upload both files via the sidebar
4. Enter the latest SHT/SHGI/TMD GBP equiv from the daily cash position sheet
5. Override client money if the month-end estimate has changed

**Adding new accounts:**
If a new bank account appears, the app will flag it with a yellow warning. Contact the developer (or edit `ACCOUNT_ENTITY` in the Python file directly) to add the mapping.

**Changing FX rates:**
Adjust the EUR/USD/CAD sliders in the sidebar. The dashboard recalculates immediately for the session. To change the default rates permanently, update `DEFAULT_FX` in the Python file.
""")

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
c1, c2, c3 = st.columns(3)
c1.caption(f"Bank data: {df_raw['PostDate'].min().strftime('%d %b %Y')} – {latest_date.strftime('%d %b %Y')}")
c2.caption(f"Forecast: {fc.index[0].strftime('%b %Y')} – {fc.index[-1].strftime('%b %Y')} ({len(fc)} months)")
c3.caption(f"Transactions: {len(df_raw):,} · Accounts: {df_raw['AccountName'].nunique()} · FX: EUR {eur_rate:.3f} · USD {usd_rate:.3f} · CAD {cad_rate:.3f}")
