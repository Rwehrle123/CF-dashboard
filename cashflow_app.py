# ══════════════════════════════════════════════════════════════════════════════
# REPLACE build_weekly() WITH THIS VERSION
# ══════════════════════════════════════════════════════════════════════════════
def build_weekly(df_raw, fx_rates):
    """
    Weekly pivot — TOTAL GROUP CASH, FX converted to GBP at budget rates.

    Important:
    - This now includes UK + Ireland entities.
    - This is what drives both the 4+13 Outlook and 3-Month Focus tabs.
    - Interco is naturally netted at total-group level if both sides are in the feed.
    - Overnight deposits are tracked separately as OD interest / sweep residual.
    """
    df = df_raw.copy()
    df['AccountName'] = df['AccountName'].astype(str).str.strip()
    df['entity']      = df['AccountName'].map(ACCOUNT_ENTITY).fillna('Ireland')
    df['TrnSpec']     = df['TrnSpec'].fillna('OTHER').str.strip()
    df.loc[df['TrnSpec'].str.upper().str.strip() == 'PAYROLL  ', 'TrnSpec'] = 'PAYROLL'
    df['Currency']    = df['AccountName'].apply(get_currency)
    df['PostDate']    = pd.to_datetime(df['PostDate'])
    df['WeekStart']   = pd.to_datetime(
        df['PostDate'].dt.to_period('W-SUN').apply(lambda p: p.start_time)
    )

    # Apply FX budget rates
    df['Amount_GBP'] = df.apply(
        lambda r: r['Amount'] * fx_rates.get(r['Currency'], 1.0), axis=1
    )

    # TOTAL GROUP operational flows: UK + Ireland, excl overnight deposits
    df_group = df[df['TrnSpec'] != 'OVERNIGHT DEPOSIT'].copy()

    # Overnight deposit net per week across total group
    df_od = df[df['TrnSpec'] == 'OVERNIGHT DEPOSIT'].copy()
    od_weekly = df_od.groupby('WeekStart')['Amount_GBP'].sum().round(0)

    weekly = (
        df_group
        .groupby(['WeekStart', 'TrnSpec'])['Amount_GBP']
        .sum()
        .round(0)
        .unstack('TrnSpec')
        .fillna(0)
    )

    return weekly, od_weekly


# ══════════════════════════════════════════════════════════════════════════════
# REPLACE THE SHARED WEEKLY DATA OPENING-BALANCE BLOCK WITH THIS VERSION
# This is the block immediately after:
#   n_actuals    = 4
#   actual_weeks = weekly_raw.index[-n_actuals:].tolist()
#   all_weeks    = actual_weeks + fc_weeks
#   N_W          = len(all_weeks)
# ══════════════════════════════════════════════════════════════════════════════

# Opening balance: TOTAL CASH basis
# Use actual total cash position if entered, else fall back to Book2 total cash
first_wk     = actual_weeks[0]
book2_before = fc[fc.index < first_wk]
book2_open   = float(book2_before.iloc[-1]['total_cash']) / 1000 if not book2_before.empty else 0.0

if use_actual_pos:
    # Use total group GBP equivalent entered in sidebar: SHT + SHGI + TMD
    # Divide by 1000 as weekly table works in £k
    open_base = pos_total / 1000
else:
    open_base = book2_open


# ══════════════════════════════════════════════════════════════════════════════
# REPLACE shared_month_end_close() WITH THIS VERSION
# ══════════════════════════════════════════════════════════════════════════════
def shared_month_end_close(yr, mn):
    """
    Return weekly-outlook TOTAL closing balance (£) at end of given month.

    This reads from the shared total-cash closes[] array built before the tabs,
    so the 3-Month Focus tab and Weekly 4+13 tab are on exactly the same basis.
    """
    wks = [
        (i, w) for i, w in enumerate(all_weeks)
        if w.year == yr and w.month == mn
    ]
    if not wks:
        return None

    last_i = max(wks, key=lambda x: x[0])[0]
    return closes[last_i] * 1000   # £k → £


# ══════════════════════════════════════════════════════════════════════════════
# IN TAB 5 — 3-MONTH FOCUS
# REPLACE THE CASH / AP CALC SECTION INSIDE THE MONTH LOOP WITH THIS VERSION
# Starts around:
#   # Opening balance:
# and runs down to before:
#   badge = ...
# ══════════════════════════════════════════════════════════════════════════════

        # Opening balance:
        # - Current month partial: use actual TOTAL cash if entered
        # - Future months: use weekly outlook TOTAL close of prior month
        prev_mn_3, prev_yr_3 = (fmn - 1, fyr) if fmn > 1 else (12, fyr - 1)
        prev_fc = fc[(fc['Year'] == prev_yr_3) & (fc['Month'] == prev_mn_3)]

        book2_open = (
            float(prev_fc.iloc[0]['total_cash'])
            if not prev_fc.empty
            else float(fc_row['total_cash'])
        )

        if is_part and use_actual_pos:
            # Current month: actual TOTAL cash position entered in sidebar
            open_cash = pos_total
        else:
            wk_open = shared_month_end_close(prev_yr_3, prev_mn_3)
            open_cash = wk_open if wk_open is not None else book2_open

        # Forecast file targets — TOTAL CASH basis
        fc_uk_close     = float(fc_row['cash_uk'])
        fc_ie_close     = float(fc_row['cash_ireland'])
        fc_total_close  = float(fc_row['total_cash'])
        fc_close        = fc_total_close

        uk_req   = float(fc_row['uk_required'])
        headroom = fc_close - uk_req
        hpct     = headroom / uk_req if uk_req > 0 else 0

        # Weekly outlook close for THIS month — TOTAL CASH chain
        wk_close_this = shared_month_end_close(fyr, fmn)
        wk_headroom   = (wk_close_this - uk_req) if wk_close_this is not None else None

        # ── Category definitions — total group basis for receipts / AP / fixed outflows ──
        INFLOW_CORE = ['AGENT RECEIPTS', 'FD RECEIPT', 'DIRECT RECEIPTS']

        # Fixed outflows: things that happen regardless and cannot easily be deferred
        OUTFLOW_FIXED = ['FLIGHT COSTS', 'PAYROLL', 'TAX']

        # Controllable AP — the main lever
        OUTFLOW_AP = ['AP COGS', 'AP OVH']

        def safe_get_total_in(yr, mn, cat):
            return (
                safe_get(uk_in_spec, yr, mn, cat) +
                safe_get(ie_in_spec, yr, mn, cat)
            )

        def safe_get_total_out_abs(yr, mn, cat):
            return abs(
                safe_get(uk_out_spec, yr, mn, cat) +
                safe_get(ie_out_spec, yr, mn, cat)
            )

        ly_inflow = sum(
            safe_get_total_in(fyr - 1, fmn, c)
            for c in INFLOW_CORE
        )

        ly_apcogs   = safe_get_total_out_abs(fyr - 1, fmn, 'AP COGS')
        ly_apovh    = safe_get_total_out_abs(fyr - 1, fmn, 'AP OVH')
        ly_ap_total = ly_apcogs + ly_apovh

        ly_flight = safe_get_total_out_abs(fyr - 1, fmn, 'FLIGHT COSTS')
        ly_payroll = safe_get_total_out_abs(fyr - 1, fmn, 'PAYROLL')
        ly_tax = safe_get_total_out_abs(fyr - 1, fmn, 'TAX')
        ly_fixed = ly_flight + ly_payroll + ly_tax

        days_in  = (pd.Timestamp(fyr, fmn, 1) + pd.offsets.MonthEnd(1)).day
        weeks_in = days_in / 7.0

        if is_part:
            rem       = days_in - latest_day
            frac_rem  = rem / days_in
            weeks_rem = rem / 7.0

            # Remaining inflows: LY pro-rated x receipts slider
            rem_in = ly_inflow * frac_rem * (1 + adj_receipts / 100)

            # Remaining fixed outflows pro-rated
            rem_fixed = ly_fixed * frac_rem

            # AP already paid this month — total group
            ap_mtd = (
                safe_get_total_out_abs(fyr, fmn, 'AP COGS') +
                safe_get_total_out_abs(fyr, fmn, 'AP OVH')
            )

            # Total AP budget on TOTAL CASH basis:
            # actual/opening total cash + remaining receipts - total forecast close - fixed outflows
            total_ap_budget = open_cash + rem_in - fc_close - rem_fixed
            allow_ap_remain = max(0, total_ap_budget - ap_mtd)
            allow_ap_per_wk = allow_ap_remain / max(weeks_rem, 0.5)

        else:
            rem       = days_in
            frac_rem  = 1.0
            weeks_rem = weeks_in

            # Full future month — apply sliders to full month estimate
            rem_in = ly_inflow * (1 + adj_receipts / 100)
            rem_fixed = ly_fixed
            ap_mtd = 0.0

            total_ap_budget = open_cash + rem_in - fc_close - rem_fixed
            allow_ap_remain = max(0, total_ap_budget)
            allow_ap_per_wk = allow_ap_remain / weeks_in

        # LY AP adjusted by AP slider for comparison
        ly_ap_for_period_adj = ly_ap_total * frac_rem * (1 + adj_payments / 100)

        # CY actual receipts to date — total group
        cy_inflow_mtd = sum(
            safe_get_total_in(fyr, fmn, c)
            for c in INFLOW_CORE
        )

        # How much to hold vs LY run rate
        ly_ap_for_period = ly_ap_total * frac_rem
        hold_vs_ly       = max(0, ly_ap_for_period_adj - allow_ap_remain)
        ly_ap_per_wk     = ly_ap_total / weeks_in


# ══════════════════════════════════════════════════════════════════════════════
# STILL IN TAB 5 — 3-MONTH FOCUS
# REPLACE THE CASH POSITION MARKDOWN BLOCK WITH THIS VERSION
# Starts around:
#   st.markdown("**Cash position**")
# ══════════════════════════════════════════════════════════════════════════════

            # ── Cash position — forecast file vs weekly outlook ────────────────
            st.markdown("**Cash position**")

            wk_cl_fmt = fmt(wk_close_this) if wk_close_this is not None else "—"
            wk_hr_fmt = fmt(wk_headroom) if wk_headroom is not None else "—"
            wk_hr_pct = (
                f"({wk_headroom / uk_req:.0%})"
                if wk_headroom is not None and uk_req > 0
                else ""
            )

            # Variance: TOTAL weekly chain close vs TOTAL forecast close
            # This now matches the 4+13 "Variance vs forecast close" row.
            var_close = (
                wk_close_this - fc_close
                if wk_close_this is not None
                else None
            )
            var_fmt = (
                f"**{'+' if var_close >= 0 else ''}{fmt(var_close)}**"
                if var_close is not None
                else "—"
            )

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
                st.error(
                    f"⚠️ Weekly outlook {fmt(abs(var_close))} **below** total forecast target — "
                    f"adjust AP or receipts to close the gap"
                )
            elif var_close is not None and var_close > 500000:
                st.success(
                    f"✅ Weekly outlook {fmt(var_close)} **ahead** of total forecast target"
                )


# ══════════════════════════════════════════════════════════════════════════════
# IN TAB 7 — WEEKLY 4+13 OUTLOOK
# REPLACE THE CAPTION AT THE TOP OF THE TAB WITH THIS VERSION
# ══════════════════════════════════════════════════════════════════════════════

    st.caption(
        f"All amounts £GBP equivalent at budget rates "
        f"(EUR×{eur_rate:.3f} · USD×{usd_rate:.3f} · CAD×{cad_rate:.3f}). "
        f"Actual: last 4 weeks from bank data. "
        f"Forecast: prior year same ISO week. "
        f"Basis: **total group cash** across UK + Ireland. "
        f"Interco nets naturally at total-group level where both sides are in the feed. "
        f"OD interest: actual net for actuals / £11k/week forecast."
    )


# ══════════════════════════════════════════════════════════════════════════════
# IN TAB 7 — WEEKLY 4+13 OUTLOOK
# REPLACE THE ACTUAL POSITION BANNER WITH THIS VERSION
# ══════════════════════════════════════════════════════════════════════════════

    # ── Actual position banner ────────────────────────────────────────────────
    if use_actual_pos:
        st.info(
            f"📍 Opening balance anchored to actual total cash position: "
            f"SHT £{pos_sht:,.0f} + SHGI £{pos_shgi:,.0f} + TMD £{pos_tmd:,.0f} = "
            f"**Total £{pos_total:,.0f}** "
            f"(W{actual_weeks[0].isocalendar()[1]} {actual_weeks[0].strftime('%d %b')} opening = "
            f"**£{open_base:,.0f}k**). "
            f"Closes chain from this base through {all_weeks[-1].strftime('%b %Y')}."
        )


# ══════════════════════════════════════════════════════════════════════════════
# IN TAB 7 — WEEKLY 4+13 OUTLOOK
# REPLACE get_book2_for_week() WITH THIS VERSION
# ══════════════════════════════════════════════════════════════════════════════

    def get_book2_for_week(wk):
        """
        Return (forecast_total_close_k, client_money_k, total_headroom_k)
        for the month containing wk, but only if wk is the LAST week in that
        calendar month among all_weeks.

        This is now TOTAL CASH basis, not UK cash.
        """
        mn = wk.month
        yr = wk.year

        # Last week in all_weeks that falls in this month
        wks_in_month = [w for w in all_weeks if w.month == mn and w.year == yr]
        if not wks_in_month or wk != max(wks_in_month):
            return None, None, None

        # Look up Book2 / forecast file
        book2_row = fc[(fc['Year'] == yr) & (fc['Month'] == mn)]
        if book2_row.empty:
            return None, None, None

        row = book2_row.iloc[0]

        # TOTAL cash close — matches closes[] which is now total-cash chain
        fc_close_k = float(row['total_cash']) / 1000
        cm_k       = float(row['client_money']) / 1000
        headroom_k = float(row['total_headroom']) / 1000

        return fc_close_k, cm_k, headroom_k


# ══════════════════════════════════════════════════════════════════════════════
# IN TAB 7 — WEEKLY 4+13 OUTLOOK
# REPLACE THE MONTH-END TRACKING ROW LABELS WITH THESE
# ══════════════════════════════════════════════════════════════════════════════

    dr_me('Forecast total close',            lambda i: fkw_me(b2_fc_close[i]), 'forecast')
    dr_me('Client money',                    lambda i: fkw_me(b2_cm[i]),       'forecast')
    dr_me('Total headroom vs UK req',        lambda i: fkw_me(b2_headroom[i]), 'forecast')
    dr_me('Variance vs fcst total close',    lambda i: fkw_var(b2_variance[i]), 'variance')


# ══════════════════════════════════════════════════════════════════════════════
# IN TAB 7 — WEEKLY 4+13 OUTLOOK
# REPLACE THIS COMMENT / VARIANCE SECTION IF PRESENT
# ══════════════════════════════════════════════════════════════════════════════

    b2_fc_close = []
    b2_cm = []
    b2_headroom = []
    b2_variance = []   # closing balance minus forecast total month-end target

    for i, wk in enumerate(all_weeks):
        fc_c, cm_c, hr_c = get_book2_for_week(wk)

        b2_fc_close.append(fc_c)
        b2_cm.append(cm_c)
        b2_headroom.append(hr_c)

        # Variance: only meaningful where Book2 total target exists
        if fc_c is not None:
            b2_variance.append(closes[i] - fc_c)
        else:
            b2_variance.append(None)


# ══════════════════════════════════════════════════════════════════════════════
# OPTIONAL BUT RECOMMENDED:
# REPLACE THE EXPORT ROW LABEL IF YOU WANT THE FILE TO BE CLEARER
# ══════════════════════════════════════════════════════════════════════════════

    exp_rows = {'Opening total cash balance': opens}
    for l in RECEIPT_LABELS:
        exp_rows[l] = data[l]
    exp_rows['Total receipts'] = totR
    for l in PAYMENT_LABELS:
        exp_rows[l] = data[l]
    exp_rows['Total payments'] = totP
    exp_rows['Net cash'] = net
    exp_rows['Closing total cash balance'] = closes
