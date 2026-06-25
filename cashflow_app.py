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

BLUE   = '#185FA5'; GREEN  = '#3B6D11'; LBLUE  = '#378ADD'; LGREEN = '#639922'
RED    = '#E24B4A'; AMBER  = '#EF9F27'; GREY   = 'rgba(0,0,0,0.06)'
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

    # YoY trend per spec
    trend = {}
    for spec in all_specs:
        if spec not in weekly.columns:
            trend[spec] = 1.0; continue
        rec      = weekly[spec].tail(13).sum()
        py_end   = last_date - pd.Timedelta(weeks=52)
        py_start = py_end   - pd.Timedelta(weeks=12)
        pri = weekly.loc[(weekly.index >= py_start) & (weekly.index <= py_end), spec].sum()
        trend[spec] = float(np.clip(rec / pri, 0.3, 3.0)) if abs(pri) > 1000 else 1.0

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
            py_date = fw - pd.Timedelta(weeks=52)
            base = 0.0
            for delta in [0, 1, -1, 2, -2]:
                sd = py_date + pd.Timedelta(weeks=delta)
                if spec in weekly.columns and sd in weekly.index:
                    base = float(weekly.loc[sd, spec]); break
            else:
                base = float(weekly[spec].tail(8).mean()) if spec in weekly.columns else 0.0
            col_fc.append(round(base * trend.get(spec, 1.0)))
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
                               help="Book2 format — rows: Client money / Cash UK / Cash Ireland")
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
    st.caption(
        "Account mapping is hard-coded — no upload required.\n"
        "BOA = Ireland entity.\n"
        "Overnight deposits excluded (net zero).\n"
        "Opening balance from Book2 prior month close.\n"
        "Interco nets within UK entities.\n"
        "Forecast always extends to Dec 2027 minimum + 12 months rolling beyond latest bank data."
    )

if not tx_file or not fc_file:
    st.title("💷 SHG Cash Flow Dashboard")
    c1, c2 = st.columns(2)
    with c1: st.info("📂 Upload **bank transactions** Excel (sidebar)\n`Data Sheet` tab required")
    with c2: st.info("📂 Upload **forecast / client money** Excel (sidebar)\nBook2 format: rows = Client money / Cash UK / Cash Ireland")
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

latest_fc = fc[fc['is_actual']].iloc[-1]
req_now   = float(latest_fc['uk_required'])
hroom_now = float(latest_fc['uk_headroom'])
hpct_now  = hroom_now / req_now if req_now > 0 else 0

# ── Page header / KPIs ────────────────────────────────────────────────────────
st.title(f"💷 SHG Cash — Latest: {latest_date.strftime('%d %b %Y')}")

if hpct_now >= 0.20:
    st.success(f"✅ COMPLIANT — UK headroom {fmt(hroom_now)} ({hpct_now:.0%} of req {fmt(req_now)})")
elif hpct_now >= 0:
    st.warning(f"⚠️ CAUTION — UK headroom {fmt(hroom_now)} ({hpct_now:.0%}). Approaching minimum.")
else:
    st.error(f"🚨 BREACH — UK cash {fmt(abs(hroom_now))} below requirement")

k = st.columns(6)
k[0].metric("UK Cash",      fmt(latest_fc['cash_uk']))
k[1].metric("Ireland Cash", fmt(latest_fc['cash_ireland']))
k[2].metric("Total Cash",   fmt(latest_fc['total_cash']))
k[3].metric("Client Money", fmt(latest_fc['client_money']))
k[4].metric("UK Required",  fmt(req_now))
k[5].metric("UK Headroom",  fmt(hroom_now),
            delta=f"{hpct_now:.0%}",
            delta_color="normal" if hroom_now >= 0 else "inverse")

st.divider()

# ── Tabs ──────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "📈 Cash vs Forecast",
    "🛡️ Compliance",
    "📊 UK Inflow YoY",
    "📊 UK Outflow YoY",
    "📊 Ireland YoY",
    "🎯 3-Month Focus",
    "⚡ Opportunities & Risks",
    "📅 Weekly 4+13 Outlook",
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
            clr = ['rgba(136,135,128,0.55)', 'rgba(24,95,165,0.7)',
                   'rgba(59,109,17,0.7)', 'rgba(239,159,39,0.7)'][yi % 4]
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
    avail_in = [c for c in KEY_INFLOW if c in uk_in_spec.columns]
    st.caption("Like-for-like: same calendar month year-on-year. Actual months only.")
    build_yoy_table(uk_in_spec, uk_out_spec, avail_in, False, "UK")

with tabs[3]:
    avail_out = [c for c in KEY_OUTFLOW if c in uk_out_spec.columns]
    st.caption("Outflows as positive. Red YoY = payments running above prior year.")
    build_yoy_table(uk_in_spec, uk_out_spec, avail_out, True, "UK")

with tabs[4]:
    st.caption("Ireland entities — like-for-like month comparison.")
    ie_in_cats  = [c for c in KEY_INFLOW  if c in ie_in_spec.columns]
    ie_out_cats = [c for c in KEY_OUTFLOW if c in ie_out_spec.columns]
    if ie_in_cats:
        st.markdown("**Inflows**")
        build_yoy_table(ie_in_spec, ie_out_spec, ie_in_cats, False, "Ireland")
    if ie_out_cats:
        st.markdown("**Outflows**")
        build_yoy_table(ie_in_spec, ie_out_spec, ie_out_cats, True, "Ireland")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — 3-Month Focus
# ══════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    focus_months = []
    for delta in range(3):
        mn = latest_mn + delta; yr = latest_yr
        while mn > 12: mn -= 12; yr += 1
        focus_months.append((yr, mn))

    st.caption(
        f"Bank data to **{latest_date.strftime('%d %b %Y')}**. "
        f"Current month ({MN[latest_mn]} {latest_yr}) + next 2. "
        f"AP limits to hit forecast closing cash. Prior year same month = expected run rate."
    )
    cols = st.columns(3)
    for ci, (fyr, fmn) in enumerate(focus_months):
        fc_row = fc[(fc['Year'] == fyr) & (fc['Month'] == fmn)]
        if fc_row.empty:
            with cols[ci]: st.info(f"{MN[fmn]} {fyr} — no forecast data"); continue
        fc_row    = fc_row.iloc[0]
        is_part   = (fyr == latest_yr and fmn == latest_mn)
        prev_mn, prev_yr = (fmn-1, fyr) if fmn > 1 else (12, fyr-1)
        prev_fc   = fc[(fc['Year'] == prev_yr) & (fc['Month'] == prev_mn)]
        open_cash = float(prev_fc.iloc[0]['cash_uk']) if not prev_fc.empty else float(fc_row['cash_uk'])
        fc_close  = float(fc_row['cash_uk'])
        uk_req    = float(fc_row['uk_required'])
        headroom  = float(fc_row['uk_headroom'])
        hpct      = headroom / uk_req if uk_req > 0 else 0
        ly_apcogs = abs(safe_get(uk_out_spec, fyr-1, fmn, 'AP COGS'))
        ly_apovh  = abs(safe_get(uk_out_spec, fyr-1, fmn, 'AP OVH'))
        ly_other  = sum(abs(safe_get(uk_out_spec, fyr-1, fmn, c))
                        for c in KEY_OUTFLOW if c not in ['AP COGS','AP OVH']
                        and c in uk_out_spec.columns)
        ly_inflow = sum(safe_get(uk_in_spec, fyr-1, fmn, c)
                        for c in KEY_INFLOW if c in uk_in_spec.columns)
        if is_part:
            days_in = (pd.Timestamp(fyr, fmn, 1) + pd.offsets.MonthEnd(1)).day
            rem     = days_in - latest_day
            rem_in  = ly_inflow * (rem / days_in)
            rem_oth = ly_other  * (rem / days_in)
            allow   = open_cash + rem_in - fc_close - rem_oth
            ap_mtd  = abs(safe_get(uk_out_spec, fyr, fmn, 'AP COGS'))
            ovh_mtd = abs(safe_get(uk_out_spec, fyr, fmn, 'AP OVH'))
            denom   = ly_apcogs + ly_apovh + 0.001
            allow_c = max(0, allow * (ly_apcogs / denom) - ap_mtd)
            allow_o = max(0, allow * (ly_apovh  / denom) - ovh_mtd)
        else:
            allow   = open_cash + ly_inflow - fc_close - ly_other
            denom   = ly_apcogs + ly_apovh + 0.001
            allow_c = max(0, allow * (ly_apcogs / denom))
            allow_o = max(0, allow * (ly_apovh  / denom))
        save_c = max(0, ly_apcogs - allow_c)
        save_o = max(0, ly_apovh  - allow_o)
        badge  = "🚨 Breach" if hpct < 0 else ("⚠️ Monitor" if hpct < 0.20 else "✅ On track")
        with cols[ci]:
            st.markdown(f"### {MN[fmn]} {fyr}  {badge}")
            if is_part:
                st.caption(f"Current month — {latest_day} of {days_in} days elapsed")
            st.markdown(
                f"| | |\n|---|---|\n"
                f"| Forecast close | **{fmt(fc_close)}** |\n"
                f"| UK required | {fmt(uk_req)} |\n"
                f"| Headroom | **{fmt(headroom)}** ({hpct:.0%}) |\n"
                f"| Expected inflows (LY) | {fmt(ly_inflow)} |"
            )
            st.markdown("**AP payment limits to hit forecast**")
            st.metric("AP COGS — can release", fmt(allow_c),
                      delta=f"LY: {fmt(ly_apcogs)} · {'hold '+fmt(save_c)+' vs LY' if save_c > 0 else 'on track'}",
                      delta_color="off")
            st.metric("AP OVH — can release", fmt(allow_o),
                      delta=f"LY: {fmt(ly_apovh)} · {'hold '+fmt(save_o)+' vs LY' if save_o > 0 else 'on track'}",
                      delta_color="off")
            total_save = save_c + save_o
            if total_save > 0.05e6:
                st.warning(f"Defer **{fmt(total_save)}** AP vs LY run rate")
            else:
                st.success("AP can run at prior year levels")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — Opportunities & Risks
# ══════════════════════════════════════════════════════════════════════════════
with tabs[6]:
    st.subheader("Forecast deviation analysis")
    st.caption("Actual flows vs forecast-implied. Signals where cash is tracking ahead or behind plan.")
    insights = []
    act_data = fc[fc['is_actual'] & ~fc['is_partial']]
    for _, row in act_data.iterrows():
        yr, mn  = int(row['Year']), int(row['Month'])
        fc_imp  = float(row.get('fc_implied_net', 0) or 0)
        act_net = safe_entity(entity_m, yr, mn, 'UK', 'net')
        dev     = act_net - fc_imp if fc_imp != 0 else 0
        if abs(dev) > 0.5e6:
            reasons = []
            checks = [
                ('FD receipts', 'FD RECEIPT',   uk_in_spec,  'inflow'),
                ('FX inflows',  'FX TRADE IN',  uk_in_spec,  'inflow'),
                ('AP COGS',     'AP COGS',       uk_out_spec, 'outflow'),
                ('FX outflows', 'FX TRADE OUT',  uk_out_spec, 'outflow'),
            ]
            for label, cat, spec, direction in checks:
                ly_v = safe_get(spec, yr-1, mn, cat)
                ac_v = safe_get(spec, yr,   mn, cat)
                if direction == 'outflow': ly_v = abs(ly_v); ac_v = abs(ac_v)
                if ly_v > 0:
                    chg = (ac_v - ly_v) / ly_v
                    if direction == 'inflow':
                        if chg >  0.15: reasons.append(f"{label} +{chg:.0%} vs LY")
                        if chg < -0.15: reasons.append(f"{label} {chg:.0%} vs LY")
                    else:
                        if chg >  0.15: reasons.append(f"{label} +{chg:.0%} vs LY (higher spend)")
                        if chg < -0.15: reasons.append(f"{label} {chg:.0%} vs LY (deferred)")
            insights.append({
                'month':    f"{MN[mn]} {yr}",
                'deviation': dev,
                'type':     'opportunity' if dev > 0 else 'risk',
                'reasons':   reasons,
                'uk_in':    safe_entity(entity_m, yr, mn, 'UK', 'inflow'),
                'uk_out':   abs(safe_entity(entity_m, yr, mn, 'UK', 'outflow')),
            })

    act_nets  = np.array([safe_entity(entity_m, int(r['Year']), int(r['Month']), 'UK', 'net')
                          for _, r in act_data.iterrows()])
    fc_nets   = act_data['fc_implied_net'].fillna(0).values
    trend_dev = float(np.mean(act_nets - fc_nets)) if len(act_nets) > 0 else 0.0

    breach_fc = fc[~fc['is_actual'] & (fc['uk_headroom'] < 0)]
    tight_fc  = fc[~fc['is_actual'] & (fc['uk_headroom'] >= 0) & (fc['headroom_pct'] < 0.15)]

    o_col, r_col = st.columns(2)
    with o_col:
        st.markdown("### Opportunities")
        if trend_dev > 0.2e6:
            st.success(f"📈 Cash tracking ahead by avg {fmt(trend_dev)}/month. Breaches may reduce.")
        for ins in [i for i in insights if i['type'] == 'opportunity']:
            with st.expander(f"✅ {ins['month']} — {fmt(ins['deviation'])} ahead of plan"):
                for r in ins['reasons']: st.markdown(f"- {r}")
                if not ins['reasons']: st.markdown("- Flows broadly in line; net better than implied")
                st.caption(f"UK inflow: {fmt(ins['uk_in'])} · outflow: {fmt(ins['uk_out'])}")
        if not breach_fc.empty and trend_dev > 0:
            st.info(f"💡 If trend continues, breach months could see ~{fmt(trend_dev * len(breach_fc))} aggregate improvement.")
    with r_col:
        st.markdown("### Risks")
        if trend_dev < -0.2e6:
            st.error(f"📉 Cash tracking below forecast by avg {fmt(abs(trend_dev))}/month.")
        if not breach_fc.empty:
            st.error("🚨 **Forecast breach months:** " + ", ".join(
                f"{MN[int(r['Month'])]} {int(r['Year'])} ({fmt(r['uk_headroom'])})"
                for _, r in breach_fc.iterrows()))
        if not tight_fc.empty:
            st.warning("⚠️ **Tight months (<15%):** " + ", ".join(
                f"{MN[int(r['Month'])]} {int(r['Year'])} ({r['headroom_pct']:.0%})"
                for _, r in tight_fc.iterrows()))
        for ins in [i for i in insights if i['type'] == 'risk']:
            with st.expander(f"🔴 {ins['month']} — {fmt(abs(ins['deviation']))} behind plan"):
                for r in ins['reasons']: st.markdown(f"- {r}")
                if not ins['reasons']: st.markdown("- Flows broadly in line; net below implied")
                st.caption(f"UK inflow: {fmt(ins['uk_in'])} · outflow: {fmt(ins['uk_out'])}")

    st.divider()
    if not act_data.empty:
        act_labels_list = (act_data['Month'].map(MN) + ' ' + act_data['Year'].astype(str)).tolist()
        actual_nets_m   = [safe_entity(entity_m, int(r['Year']), int(r['Month']), 'UK', 'net') / 1e6
                           for _, r in act_data.iterrows()]
        fc_nets_m       = [v / 1e6 for v in act_data['fc_implied_net'].fillna(0).tolist()]
        fig3 = make_subplots(rows=2, cols=1,
            subplot_titles=['Monthly net cash — actual (UK)', 'Forecast implied net (UK)'],
            vertical_spacing=0.18)
        fig3.add_trace(go.Bar(x=act_labels_list, y=actual_nets_m,
            marker_color=[GREEN if v >= 0 else RED for v in actual_nets_m], name='Actual'), row=1, col=1)
        fig3.add_trace(go.Bar(x=act_labels_list, y=fc_nets_m,
            marker_color='rgba(136,135,128,0.5)', name='Forecast implied'), row=2, col=1)
        fig3.update_layout(height=380, plot_bgcolor='white', paper_bgcolor='white',
            showlegend=True, margin=dict(l=0, r=0, t=40, b=0))
        fig3.update_yaxes(tickformat=',.1f', ticksuffix='m', gridcolor=GREY)
        st.plotly_chart(fig3, use_container_width=True)

    st.divider()
    buf = io.BytesIO()
    fc.reset_index().to_excel(buf, index=False, sheet_name='Dashboard Data')
    buf.seek(0)
    st.download_button("⬇️ Download dashboard data (.xlsx)", data=buf,
        file_name='shg_cash_dashboard.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ══════════════════════════════════════════════════════════════════════════════
# TAB 8 — Weekly 4+13 Outlook
# ══════════════════════════════════════════════════════════════════════════════
with tabs[7]:
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

    # Opening balance: last Book2 month close before first actual week
    first_wk     = actual_weeks[0]
    book2_before = fc[fc.index < first_wk]
    open_base    = float(book2_before.iloc[-1]['cash_uk']) / 1000 if not book2_before.empty else 0.0

    ROW_SPECS = [
        ('DIRECT RECEIPTS', 'DIRECT RECEIPTS',  False),
        ('AGENT RECEIPTS',  'AGENT RECEIPTS',   False),
        ('FD RECEIPT',      'FD RECEIPT',        False),
        ('CUSTOMER REFUND', 'CUSTOMER REFUNDS',  False),
        ('TUI RECEIPT',     'TUI RECEIPT',       False),
        ('OTHER RECEIPT',   'OTHER RECEIPTS',    False),
        ('FX TRADE IN',     'FX TRADE IN',       False),
        ('INTERCO (net)',   'INTERCO',           True),   # blank in forecast
        ('OD INTEREST',     '_OD',               False),
        ('AP COGS',         'AP COGS',           False),
        ('AP OVH',          'AP OVH',            False),
        ('PAYROLL',         'PAYROLL',           False),
        ('TAX',             'TAX',               False),
        ('OTHER CASH OUT',  'OTHER COSTS',       False),
        ('FX TRADE OUT',    'FX TRADE OUT',      False),
        ('FLIGHT COSTS',    'FLIGHT COSTS',      False),
    ]
    RECEIPT_LABELS = ['DIRECT RECEIPTS','AGENT RECEIPTS','FD RECEIPT','CUSTOMER REFUND',
                      'TUI RECEIPT','OTHER RECEIPT','FX TRADE IN','INTERCO (net)','OD INTEREST']
    PAYMENT_LABELS = ['AP COGS','AP OVH','PAYROLL','TAX','OTHER CASH OUT','FX TRADE OUT','FLIGHT COSTS']

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

    # ── Override expander ─────────────────────────────────────────────────────
    ovr_labels = ['AP COGS','AP OVH','PAYROLL','FD RECEIPT','AGENT RECEIPTS',
                  'FX TRADE IN','FX TRADE OUT','FLIGHT COSTS','INTERCO (net)']

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
                    fc_def  = fc_base.get('PAYROLL', [0]*13)[fi] if is_lk else \
                              (fc_base.get(spec, [0]*13)[fi] if spec in fc_base else 0.0)
                    curr    = st.session_state.weekly_ov.get(ov_key, fc_def)
                    hint    = "🔒 " if is_lk else ("🔵 " if is_ico else "")
                    new_v   = st.number_input(
                        f"{hint}{label}",
                        value=float(curr),
                        min_value=-9_999_999.0, max_value=9_999_999.0,
                        step=1000.0, format="%.0f",
                        key=f"wov_{label}_{fi}")
                    if abs(new_v - fc_def) > 1:
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

    df_display = pd.DataFrame(
        [{k: v for k, v in r.items() if not k.startswith('_')} for r in display_rows]
    ).set_index('Row')

    def style_fn(df):
        s = pd.DataFrame('', index=df.index, columns=df.columns)
        act_cols = [wk_hdrs[i] for i in range(n_actuals)]
        fc_cols  = [wk_hdrs[i] for i in range(n_actuals, N_W)]
        for rd in display_rows:
            idx = rd['Row']
            if idx not in df.index: continue
            kind = rd['_kind']
            if kind == 'header':
                s.loc[idx] = 'background-color:#1F3864;color:white;font-weight:bold'
            elif kind in ('total', 'balance'):
                s.loc[idx] = 'background-color:#DAE3F3;font-weight:bold;color:#1F3864'
            else:
                for c in act_cols:
                    s.loc[idx, c] = 'background-color:#FAFAFA'
                lbl = rd['_label']
                for ci, c in enumerate(fc_cols):
                    wi     = ci + n_actuals
                    ov_key = f"{lbl}_{wi}"
                    if 'INTERCO' in lbl:
                        s.loc[idx, c] = 'background-color:#F2F6FC;color:#5F5E5A'
                    elif ov_key in st.session_state.weekly_ov:
                        s.loc[idx, c] = 'background-color:#FFF0CC;color:#633806;font-weight:bold'
                    elif lbl in WEEKLY_LOCK:
                        s.loc[idx, c] = 'background-color:#FFF8E0;color:#633806'
                    else:
                        s.loc[idx, c] = 'background-color:#F2F6FC'
        return s

    st.dataframe(
        df_display.style.apply(style_fn, axis=None),
        use_container_width=True,
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
        bc = [BLUE if i < n_actuals else '#85B7EB' for i in range(N_W)]
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
    g_cols = st.columns(min(13, len(fc_weeks)))
    for fi, wk in enumerate(fc_weeks):
        i  = fi + n_actuals
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


# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
c1, c2, c3 = st.columns(3)
c1.caption(f"Bank data: {df_raw['PostDate'].min().strftime('%d %b %Y')} – {latest_date.strftime('%d %b %Y')}")
c2.caption(f"Forecast: {fc.index[0].strftime('%b %Y')} – {fc.index[-1].strftime('%b %Y')} ({len(fc)} months)")
c3.caption(f"Transactions: {len(df_raw):,} · Accounts: {df_raw['AccountName'].nunique()} · FX: EUR {eur_rate:.3f} · USD {usd_rate:.3f} · CAD {cad_rate:.3f}")
