import io, re, calendar, datetime
import numpy as np
import pandas as pd
import streamlit as st


# ═══════════════════════════════════════════════════════
#  SHARED CONSTANTS & HELPERS
# ═══════════════════════════════════════════════════════

COMPANY_CUTOFFS = {
    "dexcom":  (11, 10),
    "micron":  (21, 20),
    "eti":     (21, 20),
    "ybs":     (21, 20),
    "atns":    (7,  6),
    "sustio":  (16, 15),
    "te":      (16, 15),
    "jabil":   (24, 23),
    "wd":      (1,  None),
    "shopee":  (1,  None),
}

MANUAL_MAP = {
    "Auto-detect": None, "Dexcom": "dexcom", "Micron/ETI/YBS": "micron",
    "ATnS": "atns", "Sustio": "sustio", "TE": "te", "Jabil": "jabil", "WD/Shopee": "wd",
}

# Sustio shift codes
SUSTIO_PRESENT = {
    "M","N","M8","N8","RN8","RM8","ON8","PM8","PN8",
    "PN","PM","RN","RM","MR","NR","NR8","MR8","ON","OM8","OM","OM9","OM9.5","OM10",
    "MR9","MR9.5","MR10","NR9","NR9.5",
}

def detect_company(val):
    import re
    if pd.isna(val): return None
    s = str(val).lower()
    for key in COMPANY_CUTOFFS:
        # short keys (≤3 chars) must match as whole word to avoid false positives
        if len(key) <= 3:
            if not re.search(r"\b" + re.escape(key) + r"\b", s): continue
        if key in s: return key
    return None

def get_period(year, month, ck):
    if ck is None or ck not in COMPANY_CUTOFFS:
        last = calendar.monthrange(year, month)[1]
        return pd.Timestamp(year, month, 1), pd.Timestamp(year, month, last)
    start_day, end_day = COMPANY_CUTOFFS[ck]
    # FIX: WD/Shopee (end_day=None) = 1st to last day of that exact month
    if end_day is None:
        last = calendar.monthrange(year, month)[1]
        return pd.Timestamp(year, month, 1), pd.Timestamp(year, month, last)
    py = year if month > 1 else year - 1
    pm = month - 1 if month > 1 else 12
    ps = pd.Timestamp(py, pm, min(start_day, calendar.monthrange(py, pm)[1]))
    pe = pd.Timestamp(year, month, min(end_day, calendar.monthrange(year, month)[1]))
    return ps, pe

def date_to_billing_period(d, ck):
    if pd.isna(d): return None, None
    if ck is None or ck not in COMPANY_CUTOFFS:
        last = calendar.monthrange(d.year, d.month)[1]
        return pd.Timestamp(d.year, d.month, 1), pd.Timestamp(d.year, d.month, last)
    start_day, end_day = COMPANY_CUTOFFS[ck]
    # FIX: WD/Shopee (end_day=None = whole-month rule) — period is the date's own month
    if end_day is None:
        return get_period(d.year, d.month, ck)
    nm, ny = (d.month+1, d.year) if d.day >= start_day and d.month < 12 else \
             (1, d.year+1) if d.day >= start_day else (d.month, d.year)
    return get_period(ny, nm, ck)

def claim_window(j, ck=None):
    if pd.isna(j): return pd.NaT, pd.NaT
    cs = pd.Timestamp(j).normalize()
    ee = (cs + pd.DateOffset(months=3) - pd.Timedelta(days=1)).normalize()
    # FIX: WD/Shopee — snap eligible end to last day of that same month
    if ck in ("wd", "shopee"):
        last = calendar.monthrange(ee.year, ee.month)[1]
        ee = pd.Timestamp(ee.year, ee.month, last)
    return cs, ee

def fmt_date(ts):
    try: return pd.Timestamp(ts).strftime("%d/%m/%Y")
    except: return ""

def clean_str(s):
    return "" if pd.isna(s) else " ".join(str(s).strip().split()).upper()

def clean_emp(s):
    if pd.isna(s): return ""
    v = str(s).strip()
    return v[:-2] if v.endswith(".0") else v

def parse_time(val):
    if val is None: return np.nan
    if isinstance(val, datetime.time):
        return val.hour + val.minute/60 + val.second/3600
    if isinstance(val, float):
        if np.isnan(val): return np.nan
        return round(val*24*60)/60
    if isinstance(val, str):
        s = val.strip()
        if not s: return np.nan
        try:
            t = pd.to_datetime(s, errors="coerce")
            if pd.notna(t): return t.hour + t.minute/60
        except: pass
        try: return float(s)
        except: return np.nan
    try:
        t = pd.Timestamp(val); return t.hour + t.minute/60
    except: return np.nan

def calc_hours(t_in, t_out):
    if np.isnan(t_in) or np.isnan(t_out): return 0.0
    dur = t_out - t_in
    if dur < 0: dur += 24
    return min(max(dur, 0), 24)

def html_grid(df, col_labels, ps, pe, tick_col_label="Total Workdays"):
    """Render a coloured attendance tick grid as HTML."""
    th = "padding:4px 6px;background:#1e3a5f;color:white;font-size:11px;text-align:center;white-space:nowrap;border:1px solid #0d2137;"
    td = "padding:3px 5px;font-size:11px;text-align:center;border:1px solid #d0d0d0;"
    ti = "padding:3px 6px;font-size:11px;text-align:left;border:1px solid #d0d0d0;white-space:nowrap;"

    rows_html = ""
    for _, row in df.iterrows():
        cs = row.get("Claim Start", pd.NaT)
        ee = row.get("Eligible End", pd.NaT)
        cells  = f'<td style="{ti}">{row.get("Emp No", row.get("Worker No",""))}</td>'
        cells += f'<td style="{ti}">{row.get("Emp Name", row.get("Worker Name",""))}</td>'
        cells += f'<td style="{ti}">{row.get("Recruiter","")}</td>'
        cells += f'<td style="{ti}">{fmt_date(row.get("Join Date", row.get("JOIN_DATE","")))}</td>'
        cells += f'<td style="{ti}">{fmt_date(cs)}</td>'
        cells += f'<td style="{ti}">{fmt_date(ee)}</td>'

        for col in col_labels:
            val = int(row.get(col, 0))
            try:
                parts = col.split("/")
                m_num = int(parts[1]); d_num = int(parts[0])
                y_num = ps.year if m_num >= ps.month else pe.year
                day_ts = pd.Timestamp(y_num, m_num, d_num)
                outside = (pd.notna(cs) and pd.notna(ee) and
                           (day_ts < pd.Timestamp(cs) or day_ts > pd.Timestamp(ee)))
            except:
                outside = False

            if val == 1:
                cells += f'<td style="{td}background:#d4edda;font-weight:bold;">✔</td>'
            elif outside:
                cells += f'<td style="{td}background:#f0f0f0;color:#bbb;">—</td>'
            else:
                cells += f'<td style="{td}color:#ccc;">·</td>'

        cells += f'<td style="{td}font-weight:bold;background:#fff3cd;">{int(row.get(tick_col_label,0))}</td>'
        rows_html += f"<tr>{cells}</tr>\n"

    hdrs = ["Emp No","Name","Recruiter","Join Date","Claim Start","Eligible End"] + col_labels + ["Total"]
    hdr_html = "".join(f'<th style="{th}">{h}</th>' for h in hdrs)
    return f"""
    <div style="overflow-x:auto;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.2);margin-bottom:12px;">
    <table style="border-collapse:collapse;font-family:monospace;width:100%;">
      <thead><tr>{hdr_html}</tr></thead>
      <tbody>{rows_html}</tbody>
    </table></div>
    <p style="font-size:11px;color:#888;">✔ = counted day &nbsp;|&nbsp; · = not worked &nbsp;|&nbsp; — = outside claim window</p>
    """

def excel_sheet(writer, df, period_label, col_labels, ps, pe, df_summary, df_grand,
                total_col="Total Workdays"):
    wb = writer.book
    fmt_hdr   = wb.add_format({"bold":True,"bg_color":"#1e3a5f","font_color":"white","border":1,"align":"center","font_size":9,"text_wrap":True})
    fmt_tick  = wb.add_format({"bg_color":"#d4edda","align":"center","border":1,"font_size":9,"bold":True})
    fmt_dot   = wb.add_format({"align":"center","border":1,"font_size":9,"font_color":"#cccccc"})
    fmt_grey  = wb.add_format({"bg_color":"#f0f0f0","font_color":"#aaaaaa","align":"center","border":1,"font_size":9})
    fmt_total = wb.add_format({"bold":True,"bg_color":"#fff3cd","align":"center","border":1,"font_size":9})
    fmt_info  = wb.add_format({"border":1,"font_size":9})
    fmt_bold  = wb.add_format({"bold":True})
    fmt_title = wb.add_format({"bold":True,"font_size":11,"font_color":"#1e3a5f"})

    emp_col   = "Emp No" if "Emp No" in df.columns else "Worker No"
    name_col  = "Emp Name" if "Emp Name" in df.columns else "Worker Name"
    join_col  = "Join Date" if "Join Date" in df.columns else "JOIN_DATE"
    cs_col    = "Claim Start"
    ee_col    = "Eligible End" if "Eligible End" in df.columns else "Eligible End Date"

    info_cols = [emp_col, name_col, "Recruiter", join_col, cs_col, ee_col]
    all_hdrs  = info_cols + col_labels + [total_col]
    n_info    = len(info_cols)

    sheet = period_label[:31]
    ws = wb.add_worksheet(sheet)
    writer.sheets[sheet] = ws

    ws.merge_range(0, 0, 0, len(all_hdrs)-1, f"Period: {period_label}", fmt_title)
    for ci, h in enumerate(all_hdrs):
        ws.write(1, ci, h, fmt_hdr)
    for ci in range(n_info): ws.set_column(ci, ci, 16)
    for ci in range(n_info, n_info+len(col_labels)): ws.set_column(ci, ci, 5)
    ws.set_column(n_info+len(col_labels), n_info+len(col_labels), 12)
    ws.set_row(1, 28)

    for ri, (_, row) in enumerate(df.iterrows()):
        r = ri + 2
        cs = row.get(cs_col, pd.NaT)
        ee = row.get(ee_col, pd.NaT)
        ws.write(r, 0, str(row.get(emp_col,"")), fmt_info)
        ws.write(r, 1, str(row.get(name_col,"")), fmt_info)
        ws.write(r, 2, str(row.get("Recruiter","")), fmt_info)
        ws.write(r, 3, fmt_date(row.get(join_col,"")), fmt_info)
        ws.write(r, 4, fmt_date(cs), fmt_info)
        ws.write(r, 5, fmt_date(ee), fmt_info)

        for ci, col in enumerate(col_labels):
            val = int(row.get(col, 0))
            col_ci = n_info + ci
            try:
                parts = col.split("/")
                m_num = int(parts[1]); d_num = int(parts[0])
                y_num = ps.year if m_num >= ps.month else pe.year
                day_ts = pd.Timestamp(y_num, m_num, d_num)
                outside = (pd.notna(cs) and pd.notna(ee) and
                           (day_ts < pd.Timestamp(cs) or day_ts > pd.Timestamp(ee)))
            except:
                outside = False
            if val == 1: ws.write(r, col_ci, "✔", fmt_tick)
            elif outside: ws.write(r, col_ci, "—", fmt_grey)
            else: ws.write(r, col_ci, "·", fmt_dot)

        ws.write(r, n_info+len(col_labels), int(row.get(total_col, 0)), fmt_total)

    s0 = len(df) + 4
    ws.write(s0, 0, f"Period: {period_label}", fmt_bold)
    ws.write(s0+1, 0, "Per-Recruiter Summary", fmt_bold)
    df_summary.to_excel(writer, sheet_name=sheet, startrow=s0+2, index=False)
    s1 = s0 + len(df_summary) + 4
    ws.write(s1, 0, "Grand Total", fmt_bold)
    df_grand.to_excel(writer, sheet_name=sheet, startrow=s1+1, index=False)


# ═══════════════════════════════════════════════════════
#  PAGE CONFIG & TABS
# ═══════════════════════════════════════════════════════
st.set_page_config(page_title="Agency Claim Generator", layout="wide")
st.title("🧾 Agency Claim Table Generator")

tab_dexcom, tab_sustio = st.tabs(["🔵 Dexcom / General (Timecard)", "🟢 Sustio (Shift Code)"])


# ═══════════════════════════════════════════════════════
#  TAB 1 — DEXCOM / GENERAL  (time-in / time-out)
# ═══════════════════════════════════════════════════════
with tab_dexcom:
    st.markdown("**Eligibility:** Claim window = join date → join date + 3 months − 1 day")

    with st.expander("📋 Company Cutoff Reference", expanded=False):
        st.table(pd.DataFrame([
            {"Company": "Dexcom",            "Billing Period": "11th → 10th"},
            {"Company": "Micron / ETI / YBS","Billing Period": "21st → 20th"},
            {"Company": "ATnS",              "Billing Period": "7th → 6th"},
            {"Company": "Sustio",            "Billing Period": "16th → 15th"},
            {"Company": "TE",               "Billing Period": "16th → 15th"},
            {"Company": "Jabil",             "Billing Period": "24th → 23rd"},
            {"Company": "WD / Shopee",       "Billing Period": "1st → end of month"},
        ]))

    with st.sidebar:
        st.header("🔵 General Settings")
        hours_per_day = st.number_input("Hours = 1 workday", 1.0, 24.0, 8.0, key="g_hrs")
        grace_minutes = st.number_input("Grace window (min)", 0, 120, 15, key="g_grace")
        day_rate_g    = st.number_input("Rate per claim day (RM)", 0.0, 1000.0, 3.0, key="g_rate")
        day_first     = st.checkbox("Dates are DD/MM/YYYY", value=True, key="g_dayfirst")
        st.divider()
        st.subheader("Company Cutoff Override")
        manual_company = st.selectbox("Force company cutoff",
            ["Auto-detect","Dexcom","Micron/ETI/YBS","ATnS","Sustio","TE","Jabil","WD/Shopee"], key="g_co")
        forced_company = MANUAL_MAP[manual_company]

    effective_threshold = hours_per_day - grace_minutes / 60

    att_file_g = st.file_uploader("Upload Timecard",   type=["csv","xlsx","xls"], key="g_att")
    mst_file_g = st.file_uploader("Upload Masterlist", type=["xlsx","xls"],       key="g_mst")

    if att_file_g and mst_file_g:
        try:
            att = (pd.read_excel(att_file_g) if not att_file_g.name.endswith(".csv")
                   else pd.read_csv(att_file_g))
            mst = pd.read_excel(mst_file_g) if not mst_file_g.name.endswith(".xls") else \
                  pd.read_excel(mst_file_g, engine="xlrd")
            att.columns = att.columns.str.strip()
            mst.columns = mst.columns.str.strip()

            st.subheader("1️⃣ Map Timecard Columns")
            t_emp  = st.selectbox("Emp No (Timecard)", att.columns, key="g_t_emp")
            t_name = st.selectbox("Name (Timecard)",   att.columns, key="g_t_name")
            t_date = st.selectbox("Date",              att.columns, key="g_t_date")
            t_in   = st.selectbox("Time In",           att.columns, key="g_t_in")
            t_out  = st.selectbox("Time Out",          att.columns, key="g_t_out")

            st.subheader("2️⃣ Map Masterlist Columns")
            m_emp  = st.selectbox("Emp No (Masterlist)", mst.columns, key="g_m_emp")
            m_name = st.selectbox("Name (Masterlist)",   mst.columns, key="g_m_name")
            m_join = st.selectbox("Joined Date",          mst.columns, key="g_m_join")
            m_recr = st.selectbox("Recruiter",   ["(none)"]+mst.columns.tolist(), key="g_m_recr")
            m_comp = st.selectbox("Company (for cutoff)",["(none)"]+mst.columns.tolist(), key="g_m_comp")

            # Normalise
            att["_emp"]   = att[t_emp].apply(clean_emp)
            att["_name"]  = att[t_name].apply(clean_str)
            att["_date"]  = pd.to_datetime(att[t_date], errors="coerce")
            att = att[att["_date"].dt.year >= 2000].copy()
            att["_in"]    = att[t_in].apply(parse_time)
            att["_out"]   = att[t_out].apply(parse_time)
            att["_hours"] = [calc_hours(i, o) for i, o in zip(att["_in"], att["_out"])]

            mst["_emp"]  = mst[m_emp].apply(clean_emp)
            mst["_name"] = mst[m_name].apply(clean_str)
            mst["_join"] = pd.to_datetime(mst[m_join], errors="coerce")

            jmap_e = dict(zip(mst["_emp"],  mst["_join"]))
            jmap_n = dict(zip(mst["_name"], mst["_join"]))
            att["JOIN_DATE"] = att["_emp"].map(jmap_e)
            att.loc[att["JOIN_DATE"].isna(), "JOIN_DATE"] = \
                att.loc[att["JOIN_DATE"].isna(), "_name"].map(jmap_n)

            n_matched   = att.dropna(subset=["JOIN_DATE"])["_emp"].nunique()
            n_total_emp = att["_emp"].nunique()
            if n_matched < n_total_emp:
                st.info(f"ℹ️ {n_total_emp - n_matched} of {n_total_emp} timecard employees "
                        f"not in masterlist — ignored. Output covers {n_matched} matched employee(s).")

            att = att.dropna(subset=["JOIN_DATE"]).copy()
            if att.empty:
                st.warning("No timecard rows matched any masterlist employee.")
                st.stop()

            # Company & Recruiter
            if m_comp != "(none)":
                cmap_e = dict(zip(mst["_emp"],  mst[m_comp]))
                cmap_n = dict(zip(mst["_name"], mst[m_comp]))
                att["_company"] = att["_emp"].map(cmap_e)
                att.loc[att["_company"].isna(),"_company"] = att.loc[att["_company"].isna(),"_name"].map(cmap_n)
            else:
                att["_company"] = np.nan
            att["_ck"] = forced_company if forced_company else att["_company"].apply(detect_company)

            if m_recr != "(none)":
                rmap_e = dict(zip(mst["_emp"],  mst[m_recr]))
                rmap_n = dict(zip(mst["_name"], mst[m_recr]))
                att["Recruiter"] = att["_emp"].map(rmap_e)
                att.loc[att["Recruiter"].isna(),"Recruiter"] = \
                    att.loc[att["Recruiter"].isna(),"_name"].map(rmap_n)
            else:
                att["Recruiter"] = "Unassigned"
            att["Recruiter"] = att["Recruiter"].fillna("Unassigned")

            att[["CLAIM_START","ELIGIBLE_END"]] = att[["JOIN_DATE","_ck"]].apply(
                lambda row: pd.Series(claim_window(row["JOIN_DATE"], row["_ck"])), axis=1)

            # Daily aggregation
            daily = (
                att.groupby(["_emp","_name","Recruiter","_ck",
                             "JOIN_DATE","CLAIM_START","ELIGIBLE_END","_date"], as_index=False)
                .agg(hrs=("_hours","sum"))
            )
            daily["worked"] = np.where(
                (daily["hrs"] >= effective_threshold) &
                (daily["_date"] >= daily["CLAIM_START"]) &
                (daily["_date"] <= daily["ELIGIBLE_END"]), 1, 0)

            def assign_period(row):
                ps, pe = date_to_billing_period(row["_date"], row["_ck"])
                return pd.Series({"_ps": ps, "_pe": pe})
            daily[["_ps","_pe"]] = daily.apply(assign_period, axis=1)
            daily["_pk"] = daily["_ps"].astype(str)

            all_pks = sorted(daily["_pk"].dropna().unique())
            if len(all_pks) > 1:
                st.info(f"ℹ️ {len(all_pks)} billing periods found in this file. "
                        f"Each is shown as a separate section below and exported as its own sheet in the Excel.")

            tables    = {}
            summaries = {}

            for pk in all_pks:
                pdata = daily[daily["_pk"] == pk].copy()
                ps, pe = pdata["_ps"].iloc[0], pdata["_pe"].iloc[0]
                all_dates  = pd.date_range(ps, pe)
                col_labels = [d.strftime("%-d/%-m") for d in all_dates]
                date_to_col = {d.normalize(): d.strftime("%-d/%-m") for d in all_dates}

                base = (
                    pdata.groupby(["_emp","_name","Recruiter","JOIN_DATE","CLAIM_START","ELIGIBLE_END"],
                                  as_index=False)["worked"].sum()
                    .rename(columns={"worked":"Total Workdays"})
                ).reset_index(drop=True)
                for col in col_labels: base[col] = 0

                for _, row in pdata.iterrows():
                    d = row["_date"]
                    if pd.isna(d): continue
                    col = date_to_col.get(d.normalize())
                    if col is None: continue
                    idx = base.index[base["_emp"] == row["_emp"]].tolist()
                    if idx: base.at[idx[0], col] = int(row["worked"])

                base = base.rename(columns={"_emp":"Emp No","_name":"Emp Name",
                    "JOIN_DATE":"Join Date","CLAIM_START":"Claim Start","ELIGIBLE_END":"Eligible End"})
                ordered = (["Emp No","Emp Name","Recruiter","Join Date","Claim Start","Eligible End"]
                           + col_labels + ["Total Workdays"])
                base = base[[c for c in ordered if c in base.columns]].reset_index(drop=True)

                period_label = f"{ps.strftime('%d %b')} – {pe.strftime('%d %b %Y')}"
                tables[pk] = (base, period_label, col_labels, ps, pe)

                summary = base.groupby("Recruiter", as_index=False)["Total Workdays"].sum()
                summary["Rate (RM)"]   = day_rate_g
                summary["Amount (RM)"] = summary["Total Workdays"] * day_rate_g
                grand = pd.DataFrame({"Recruiter":["TOTAL"],
                    "Total Workdays":[summary["Total Workdays"].sum()],
                    "Rate (RM)":[day_rate_g],
                    "Amount (RM)":[summary["Amount (RM)"].sum()]})
                summaries[pk] = (summary, grand)

            # Display
            for pk in tables:
                df, period_label, col_labels, ps, pe = tables[pk]
                st.subheader(f"📅 Period: {period_label}")
                with st.expander("👁️ Employee Attendance Grid", expanded=True):
                    st.markdown(html_grid(df, col_labels, ps, pe), unsafe_allow_html=True)
                st.caption("Per-Recruiter Summary")
                st.dataframe(summaries[pk][0], use_container_width=True)
                st.write("Grand Total")
                st.dataframe(summaries[pk][1], use_container_width=True)
                st.divider()

            # Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                for pk in tables:
                    df, period_label, col_labels, ps, pe = tables[pk]
                    excel_sheet(writer, df, period_label, col_labels, ps, pe,
                                summaries[pk][0], summaries[pk][1])
            st.download_button("⬇️ Download Excel", buffer.getvalue(),
                file_name="claim_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
    else:
        st.info("Upload both Timecard and Masterlist files to continue.")


# ═══════════════════════════════════════════════════════
#  TAB 2 — SUSTIO  (shift-code grid)
# ═══════════════════════════════════════════════════════
with tab_sustio:
    st.markdown("""
    **Sustio attendance format:** Each row is an employee. Day columns contain shift codes.  
    **Present codes:** M, N, M8, N8, OM, ON, OM8, ON8, PM, PN, RM, RN, MR, NR (and OT variants)  
    **Billing period:** 16th of previous month → 15th of current month  
    **Eligibility:** Join date → join date + 3 months − 1 day
    """)

    with st.sidebar:
        st.markdown("---")
        st.header("🟢 Sustio Settings")
        day_rate_s  = st.number_input("Rate per claim day (RM)", 0.0, 1000.0, 3.0, key="s_rate")
        excl_unassigned = st.checkbox("Exclude Unassigned from Excel export", value=True, key="s_excl")

        st.markdown("**Select claim month (period ends on 15th):**")
        today = datetime.date.today()
        years = [today.year-1, today.year, today.year+1]
        sel_year  = st.selectbox("Year",  years,  index=years.index(today.year), key="s_yr")
        sel_month = st.selectbox("Month", list(calendar.month_abbr)[1:], index=today.month-1, key="s_mo")
        sel_month_num = list(calendar.month_abbr).index(sel_month)
        cycle_end_date = datetime.date(sel_year, sel_month_num, 15)

    # Derived period bounds for Sustio
    cycle_end_ts = pd.Timestamp(cycle_end_date)
    prev_m = cycle_end_ts.month - 1 if cycle_end_ts.month > 1 else 12
    prev_y = cycle_end_ts.year if cycle_end_ts.month > 1 else cycle_end_ts.year - 1
    sustio_ps = pd.Timestamp(prev_y, prev_m, 16)
    sustio_pe = pd.Timestamp(cycle_end_ts.year, cycle_end_ts.month, 15)
    st.info(f"Sustio claim window: **{sustio_ps.strftime('%d %b %Y')} → {sustio_pe.strftime('%d %b %Y')}**")

    att_file_s = st.file_uploader("Upload Sustio Attendance (xlsx/xls)", type=["xlsx","xls"], key="s_att")
    mst_file_s = st.file_uploader("Upload Sustio Masterlist (xlsx/xls)", type=["xlsx","xls"], key="s_mst")

    if att_file_s and mst_file_s:
        try:
            # ── Load attendance — skip legend rows, find header ──
            engine_a = "xlrd" if att_file_s.name.endswith(".xls") else "openpyxl"
            raw_att  = pd.read_excel(att_file_s, header=None, engine=engine_a)

            header_idx = None
            for i in range(min(20, len(raw_att))):
                vals = [str(v).strip().lower() for v in raw_att.iloc[i].tolist()]
                if any(k in vals for k in ["emp no","employee no","worker no","emp_no","no pekerja"]):
                    header_idx = i; break
            if header_idx is None:
                st.error("Could not find header row in attendance file (looking for 'EMP NO' / 'NAME').")
                st.stop()

            att_s = pd.read_excel(att_file_s, header=header_idx, engine=engine_a)
            att_s.columns = [str(c).strip() for c in att_s.columns]

            # Rename standard columns
            col_map = {}
            for c in att_s.columns:
                cl = c.lower().replace(" ","").replace("_","")
                if cl in {"empno","employeeno","workerno","noperkerja","empid","id"}: col_map[c] = "EMP_NO"
                elif cl in {"name","workername","employeename","nama"}: col_map[c] = "NAME"
                elif cl in {"joineddate","joindate","datejoined","tarikh masuk","doj"}: col_map[c] = "JOINED_DATE"
            att_s = att_s.rename(columns=col_map)

            for req in ["EMP_NO","NAME"]:
                if req not in att_s.columns:
                    st.error(f"Could not find '{req}' column in attendance. Please check your file.")
                    st.stop()

            # Identify day columns: integers 1–31 (may be int or string)
            raw_day_cols = []
            for c in att_s.columns:
                try:
                    n = int(str(c).strip())
                    if 1 <= n <= 31: raw_day_cols.append(c)
                except: pass

            # Split: days >= 16 are PREVIOUS month, days 1-15 are CURRENT month
            prev_month_days = sorted([c for c in raw_day_cols if int(str(c)) >= 16], key=lambda x: int(str(x)))
            curr_month_days = sorted([c for c in raw_day_cols if int(str(c)) <= 15],  key=lambda x: int(str(x)))
            ordered_day_cols = prev_month_days + curr_month_days  # 16..31, 1..15

            # Drop fully empty rows
            att_s = att_s.dropna(subset=["NAME"]).reset_index(drop=True)

            # Normalise emp/name
            att_s["_emp"]  = att_s["EMP_NO"].apply(lambda s: str(s).strip().upper() if not pd.isna(s) else "")
            att_s["_name"] = att_s["NAME"].apply(clean_str)

            # ── Load masterlist ──
            engine_m = "xlrd" if mst_file_s.name.endswith(".xls") else "openpyxl"
            mst_s = pd.read_excel(mst_file_s, engine=engine_m)
            mst_s.columns = mst_s.columns.str.strip()

            st.subheader("Map Masterlist Columns")
            m_emp_s  = st.selectbox("Emp No col",   mst_s.columns, key="s_m_emp",
                index=next((i for i,c in enumerate(mst_s.columns) if "emp" in c.lower()), 0))
            m_name_s = st.selectbox("Name col",     mst_s.columns, key="s_m_name",
                index=next((i for i,c in enumerate(mst_s.columns) if "name" in c.lower()), 0))
            m_join_s = st.selectbox("Joined Date col", mst_s.columns, key="s_m_join",
                index=next((i for i,c in enumerate(mst_s.columns) if "join" in c.lower() or "date" in c.lower()), 0))
            m_recr_s = st.selectbox("Recruiter col", ["(none)"]+mst_s.columns.tolist(), key="s_m_recr",
                index=next((i+1 for i,c in enumerate(mst_s.columns) if "recruit" in c.lower()), 0))

            mst_s["_emp"]  = mst_s[m_emp_s].apply(lambda s: str(s).strip().upper() if not pd.isna(s) else "")
            mst_s["_name"] = mst_s[m_name_s].apply(clean_str)
            mst_s["_join"] = pd.to_datetime(mst_s[m_join_s], errors="coerce")
            mst_s["_recr"] = mst_s[m_recr_s].apply(clean_str) if m_recr_s != "(none)" else "Unassigned"

            # ── Match attendance → masterlist ──
            jmap_e = dict(zip(mst_s["_emp"],  mst_s["_join"]))
            jmap_n = dict(zip(mst_s["_name"], mst_s["_join"]))
            rmap_e = dict(zip(mst_s["_emp"],  mst_s["_recr"]))
            rmap_n = dict(zip(mst_s["_name"], mst_s["_recr"]))

            att_s["JOIN_DATE"] = att_s["_emp"].map(jmap_e)
            att_s.loc[att_s["JOIN_DATE"].isna(),"JOIN_DATE"] = att_s.loc[att_s["JOIN_DATE"].isna(),"_name"].map(jmap_n)

            att_s["Recruiter"] = att_s["_emp"].map(rmap_e)
            att_s.loc[att_s["Recruiter"].isna(),"Recruiter"] = att_s.loc[att_s["Recruiter"].isna(),"_name"].map(rmap_n)
            att_s["Recruiter"] = att_s["Recruiter"].fillna("Unassigned")

            n_matched = att_s.dropna(subset=["JOIN_DATE"])["_emp"].nunique()
            n_total   = att_s["_emp"].nunique()
            if n_matched < n_total:
                st.info(f"ℹ️ {n_total - n_matched} of {n_total} attendance employees not matched "
                        f"to masterlist — ignored.")
            att_s = att_s.dropna(subset=["JOIN_DATE"]).copy()
            if att_s.empty:
                st.warning("No attendance rows matched any masterlist employee.")
                st.stop()

            att_s[["CLAIM_START","ELIGIBLE_END"]] = att_s["JOIN_DATE"].apply(
                lambda d: pd.Series(claim_window(d)))

            # ── Convert shift codes → 0/1 ──
            def is_present(v):
                if pd.isna(v): return 0
                s = str(v).strip().upper()
                # Handle typos like OM9, OM9.5
                base = re.sub(r"\d+(\.\d+)?$", "", s)
                return 1 if (s in SUSTIO_PRESENT or base in SUSTIO_PRESENT) else 0

            for c in ordered_day_cols:
                att_s[c] = att_s[c].apply(is_present)

            # ── Build col labels (DD/MM) and cap by claim window ──
            def day_to_date(day_num):
                d = int(str(day_num))
                if d >= 16:
                    return pd.Timestamp(sustio_ps.year, sustio_ps.month, d)
                else:
                    return pd.Timestamp(sustio_pe.year, sustio_pe.month, d)

            col_labels_s = [day_to_date(c).strftime("%-d/%-m") for c in ordered_day_cols]
            day_to_label = {c: day_to_date(c).strftime("%-d/%-m") for c in ordered_day_cols}

            # Cap days outside individual claim windows
            for c in ordered_day_cols:
                actual_date = day_to_date(c)
                outside_mask = (
                    (att_s["CLAIM_START"].notna() & (actual_date < att_s["CLAIM_START"])) |
                    (att_s["ELIGIBLE_END"].notna() & (actual_date > att_s["ELIGIBLE_END"]))
                )
                att_s.loc[outside_mask, c] = 0

            att_s["Total Working Days"] = att_s[ordered_day_cols].sum(axis=1).astype(int)

            # Rename day columns to DD/MM labels
            rename_days = {c: day_to_label[c] for c in ordered_day_cols}
            att_s = att_s.rename(columns=rename_days)

            # Build output DataFrame
            out_df = att_s[["_emp","_name","Recruiter","JOIN_DATE","CLAIM_START","ELIGIBLE_END"]
                           + col_labels_s + ["Total Working Days"]].copy()
            out_df = out_df.rename(columns={
                "_emp":"Worker No","_name":"Worker Name",
                "JOIN_DATE":"Join Date","CLAIM_START":"Claim Start","ELIGIBLE_END":"Eligible End"
            })

            period_label_s = f"{sustio_ps.strftime('%d %b')} – {sustio_pe.strftime('%d %b %Y')}"

            # ── Display ──
            st.subheader(f"📅 Sustio Period: {period_label_s}")
            with st.expander("👁️ Employee Attendance Grid", expanded=True):
                st.markdown(html_grid(out_df, col_labels_s, sustio_ps, sustio_pe,
                                      tick_col_label="Total Working Days"),
                            unsafe_allow_html=True)

            st.caption("Per-Recruiter Summary")
            summary_s = out_df.groupby("Recruiter", as_index=False)["Total Working Days"].sum()
            summary_s["Rate (RM)"]   = day_rate_s
            summary_s["Amount (RM)"] = summary_s["Total Working Days"] * day_rate_s
            grand_s = pd.DataFrame({"Recruiter":["TOTAL"],
                "Total Working Days":[summary_s["Total Working Days"].sum()],
                "Rate (RM)":[day_rate_s],
                "Amount (RM)":[summary_s["Amount (RM)"].sum()]})
            st.dataframe(summary_s, use_container_width=True)
            st.write("Grand Total")
            st.dataframe(grand_s, use_container_width=True)

            # ── Excel export ──
            export_df = out_df[out_df["Recruiter"] != "Unassigned"].copy() if excl_unassigned else out_df.copy()
            sum_export = export_df.groupby("Recruiter", as_index=False)["Total Working Days"].sum()
            sum_export["Rate (RM)"]   = day_rate_s
            sum_export["Amount (RM)"] = sum_export["Total Working Days"] * day_rate_s
            grand_exp = pd.DataFrame({"Recruiter":["TOTAL"],
                "Total Working Days":[sum_export["Total Working Days"].sum()],
                "Rate (RM)":[day_rate_s],
                "Amount (RM)":[sum_export["Amount (RM)"].sum()]})

            buf_s = io.BytesIO()
            with pd.ExcelWriter(buf_s, engine="xlsxwriter") as writer_s:
                excel_sheet(writer_s, export_df, period_label_s, col_labels_s,
                            sustio_ps, sustio_pe, sum_export, grand_exp,
                            total_col="Total Working Days")
            st.download_button("⬇️ Download Sustio Excel", buf_s.getvalue(),
                file_name="sustio_claim_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
    else:
        st.info("Upload both Sustio Attendance and Masterlist files to continue.")
