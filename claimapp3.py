import io
import calendar
import datetime
import numpy as np
import pandas as pd
import streamlit as st


# ───────────────────────── Company Cutoff Config ─────────────────────────
COMPANY_CUTOFFS = {
    "dexcom":  (11, 10),
    "micron":  (21, 20),
    "eti":     (21, 20),
    "ybs":     (21, 20),
    "atns":    (7,  6),
    "sustio":  (16, 15),
    "jabil":   (24, 23),
    "wd":      (1,  None),
    "shopee":  (1,  None),
}

MANUAL_MAP = {
    "Auto-detect": None, "Dexcom": "dexcom", "Micron/ETI/YBS": "micron",
    "ATnS": "atns", "Sustio": "sustio", "Jabil": "jabil", "WD/Shopee": "wd",
}

def detect_company(val):
    if pd.isna(val): return None
    s = str(val).lower()
    for key in COMPANY_CUTOFFS:
        if key in s: return key
    return None

def get_period(year, month, ck):
    if ck is None or ck not in COMPANY_CUTOFFS:
        last = calendar.monthrange(year, month)[1]
        return pd.Timestamp(year, month, 1), pd.Timestamp(year, month, last)
    start_day, end_day = COMPANY_CUTOFFS[ck]
    py = year if month > 1 else year - 1
    pm = month - 1 if month > 1 else 12
    ps = pd.Timestamp(py, pm, min(start_day, calendar.monthrange(py, pm)[1]))
    if end_day is None:
        pe = pd.Timestamp(year, month, calendar.monthrange(year, month)[1])
    else:
        pe = pd.Timestamp(year, month, min(end_day, calendar.monthrange(year, month)[1]))
    return ps, pe

def date_to_billing_period(d, ck):
    if pd.isna(d): return None, None
    if ck is None or ck not in COMPANY_CUTOFFS:
        last = calendar.monthrange(d.year, d.month)[1]
        return pd.Timestamp(d.year, d.month, 1), pd.Timestamp(d.year, d.month, last)
    start_day, _ = COMPANY_CUTOFFS[ck]
    if d.day >= start_day:
        nm = d.month + 1 if d.month < 12 else 1
        ny = d.year if d.month < 12 else d.year + 1
    else:
        nm, ny = d.month, d.year
    return get_period(ny, nm, ck)


# ───────────────────────── UI ─────────────────────────
st.set_page_config(page_title="Agency Claim Table Generator", layout="wide")
st.title("🧾 Agency Claim Table Generator")
st.markdown("**Eligibility:** Claim window = join date → join date + 3 months − 1 day")

with st.expander("📋 Company Cutoff Reference", expanded=False):
    st.table(pd.DataFrame([
        {"Company": "Dexcom",            "Billing Period": "11th → 10th"},
        {"Company": "Micron / ETI / YBS","Billing Period": "21st → 20th"},
        {"Company": "ATnS",              "Billing Period": "7th → 6th"},
        {"Company": "Sustio",            "Billing Period": "16th → 15th"},
        {"Company": "Jabil",             "Billing Period": "24th → 23rd"},
        {"Company": "WD / Shopee",       "Billing Period": "1st → end of month"},
    ]))

# ───────────────────────── Sidebar ─────────────────────────
with st.sidebar:
    st.header("Settings")
    hours_per_day = st.number_input("Hours = 1 workday", 1.0, 24.0, 8.0)
    grace_minutes = st.number_input("Grace window (min)", 0, 120, 15)
    day_rate      = st.number_input("Rate per claim day (RM)", 0.0, 1000.0, 3.0)
    day_first     = st.checkbox("Dates are DD/MM/YYYY", value=True)
    st.divider()
    st.subheader("Company Cutoff Override")
    st.caption("Force one company's cutoff for ALL records, or leave as Auto-detect.")
    manual_company = st.selectbox("Force company cutoff",
        ["Auto-detect", "Dexcom", "Micron/ETI/YBS", "ATnS", "Sustio", "Jabil", "WD/Shopee"])
    forced_company = MANUAL_MAP[manual_company]

effective_threshold = hours_per_day - grace_minutes / 60


# ───────────────────────── Upload ─────────────────────────
att_file = st.file_uploader("Upload Timecard",   type=["csv","xlsx","xls"])
mst_file = st.file_uploader("Upload Masterlist", type=["xlsx","xls"])


# ───────────────────────── Helpers ─────────────────────────
def clean_str(s):
    """Strip whitespace, normalise inner spaces, uppercase."""
    if pd.isna(s): return ""
    return " ".join(str(s).strip().split()).upper()

def clean_emp(s):
    """Return emp no as a plain string (no .0 suffix)."""
    if pd.isna(s): return ""
    v = str(s).strip()
    if v.endswith(".0"):
        v = v[:-2]
    return v

def parse_time(val):
    """
    Return hours as float from any of:
      - datetime.time  (native Excel time after openpyxl loads it)
      - float          (Excel fractional day stored as number)
      - str            (HH:MM or blank)
      - NaN / None
    """
    if val is None: return np.nan
    if isinstance(val, datetime.time):
        return val.hour + val.minute / 60 + val.second / 3600
    if isinstance(val, float):
        if np.isnan(val): return np.nan
        # fractional day  (e.g. 0.781944 ≈ 18:46)
        total_minutes = round(val * 24 * 60)
        return total_minutes / 60
    if isinstance(val, str):
        s = val.strip()
        if not s: return np.nan
        try:
            t = pd.to_datetime(s, errors="coerce")
            if pd.notna(t): return t.hour + t.minute / 60
        except: pass
        try: return float(s)
        except: return np.nan
    # datetime / Timestamp
    try:
        t = pd.Timestamp(val)
        return t.hour + t.minute / 60
    except: return np.nan

def calc_hours(t_in, t_out):
    if np.isnan(t_in) or np.isnan(t_out): return 0.0
    dur = t_out - t_in
    if dur < 0: dur += 24          # overnight / night-shift
    return min(max(dur, 0), 24)

def claim_window(j):
    if pd.isna(j): return pd.NaT, pd.NaT
    cs = pd.Timestamp(j).normalize()
    ee = (cs + pd.DateOffset(months=3) - pd.Timedelta(days=1)).normalize()
    return cs, ee

def fmt_date(ts):
    try: return pd.Timestamp(ts).strftime("%d/%m/%Y")
    except: return ""


# ───────────────────────── Main ─────────────────────────
if att_file and mst_file:
    try:
        # ── Load ──
        att = pd.read_excel(att_file) if not att_file.name.endswith(".csv") else pd.read_csv(att_file)
        mst = pd.read_excel(mst_file)

        # Strip column name whitespace
        att.columns = att.columns.str.strip()
        mst.columns = mst.columns.str.strip()

        # ── Column Mapping ──
        st.subheader("1️⃣ Map Timecard Columns")
        t_emp  = st.selectbox("Emp No (Timecard)", att.columns)
        t_name = st.selectbox("Name (Timecard)",   att.columns)
        t_date = st.selectbox("Date",              att.columns)
        t_in   = st.selectbox("Time In",           att.columns)
        t_out  = st.selectbox("Time Out",          att.columns)

        st.subheader("2️⃣ Map Masterlist Columns")
        m_emp  = st.selectbox("Emp No (Masterlist)", mst.columns)
        m_name = st.selectbox("Name (Masterlist)",   mst.columns)
        m_join = st.selectbox("Joined Date",          mst.columns)
        m_recr = st.selectbox("Recruiter",           ["(none)"] + mst.columns.tolist())
        m_comp = st.selectbox("Company (for cutoff)",["(none)"] + mst.columns.tolist())

        # ── Normalise timecard ──
        att["_emp"]   = att[t_emp].apply(clean_emp)
        att["_name"]  = att[t_name].apply(clean_str)
        att["_date"]  = pd.to_datetime(att[t_date], errors="coerce")
        # Filter out garbage dates (e.g. serial 0 → 1970)
        att = att[att["_date"].dt.year >= 2000].copy()
        att["_in"]    = att[t_in].apply(parse_time)
        att["_out"]   = att[t_out].apply(parse_time)
        att["_hours"] = [calc_hours(i, o) for i, o in zip(att["_in"], att["_out"])]

        # ── Normalise masterlist ──
        mst["_emp"]  = mst[m_emp].apply(clean_emp)
        mst["_name"] = mst[m_name].apply(clean_str)
        mst["_join"] = pd.to_datetime(mst[m_join], errors="coerce")

        # ── Lookup join date: emp no first, then name ──
        jmap_e = dict(zip(mst["_emp"],  mst["_join"]))
        jmap_n = dict(zip(mst["_name"], mst["_join"]))
        att["JOIN_DATE"] = att["_emp"].map(jmap_e)
        miss = att["JOIN_DATE"].isna()
        att.loc[miss, "JOIN_DATE"] = att.loc[miss, "_name"].map(jmap_n)

        n_matched   = att.dropna(subset=["JOIN_DATE"])["_emp"].nunique()
        n_total_emp = att["_emp"].nunique()

        if n_matched < n_total_emp:
            st.info(
                f"ℹ️ {n_total_emp - n_matched} of {n_total_emp} employees in the timecard "
                f"are not in the masterlist and will be ignored. "
                f"Output will cover the {n_matched} matched employee(s)."
            )

        # Drop rows with no join date (unmatched employees) — silently ignored
        att = att.dropna(subset=["JOIN_DATE"]).copy()

        if att.empty:
            st.warning("No timecard rows matched any employee in the masterlist. Please check you have the right files.")
            st.stop()

        # ── Company ──
        if m_comp != "(none)":
            cmap_e = dict(zip(mst["_emp"],  mst[m_comp]))
            cmap_n = dict(zip(mst["_name"], mst[m_comp]))
            att["_company"] = att["_emp"].map(cmap_e)
            miss = att["_company"].isna()
            att.loc[miss, "_company"] = att.loc[miss, "_name"].map(cmap_n)
        else:
            att["_company"] = np.nan

        att["_ck"] = forced_company if forced_company else att["_company"].apply(detect_company)

        # ── Recruiter ──
        if m_recr != "(none)":
            rmap_e = dict(zip(mst["_emp"],  mst[m_recr]))
            rmap_n = dict(zip(mst["_name"], mst[m_recr]))
            att["Recruiter"] = att["_emp"].map(rmap_e)
            miss = att["Recruiter"].isna()
            att.loc[miss, "Recruiter"] = att.loc[miss, "_name"].map(rmap_n)
        else:
            att["Recruiter"] = "Unassigned"
        att["Recruiter"] = att["Recruiter"].fillna("Unassigned")

        # ── Claim window ──
        att[["CLAIM_START","ELIGIBLE_END"]] = att["JOIN_DATE"].apply(
            lambda d: pd.Series(claim_window(d))
        )

        # ── Daily aggregation ──
        daily = (
            att.groupby(
                ["_emp","_name","Recruiter","_ck",
                 "JOIN_DATE","CLAIM_START","ELIGIBLE_END","_date"],
                as_index=False
            ).agg(hrs=("_hours","sum"))
        )

        daily["worked"] = np.where(
            (daily["hrs"] >= effective_threshold) &
            (daily["_date"] >= daily["CLAIM_START"]) &
            (daily["_date"] <= daily["ELIGIBLE_END"]),
            1, 0
        )

        # ── Assign billing period ──
        def assign_period(row):
            ps, pe = date_to_billing_period(row["_date"], row["_ck"])
            return pd.Series({"_ps": ps, "_pe": pe})

        daily[["_ps","_pe"]] = daily.apply(assign_period, axis=1)
        daily["_pk"] = daily["_ps"].astype(str)

        periods = sorted(daily["_pk"].dropna().unique())
        tables    = {}
        summaries = {}

        for pk in periods:
            pdata = daily[daily["_pk"] == pk].copy()
            ps    = pdata["_ps"].iloc[0]
            pe    = pdata["_pe"].iloc[0]

            all_dates  = pd.date_range(ps, pe)
            col_labels = [d.strftime("%-d/%-m") for d in all_dates]
            date_to_col = {d.normalize(): d.strftime("%-d/%-m") for d in all_dates}

            # One row per employee
            base = (
                pdata.groupby(
                    ["_emp","_name","Recruiter","JOIN_DATE","CLAIM_START","ELIGIBLE_END"],
                    as_index=False
                )["worked"].sum()
                .rename(columns={"worked":"Total Workdays"})
            )
            base = base.reset_index(drop=True)

            for col in col_labels:
                base[col] = 0

            # Fill attendance per day
            for _, row in pdata.iterrows():
                d = row["_date"]
                if pd.isna(d): continue
                col = date_to_col.get(d.normalize())
                if col is None: continue
                idx_list = base.index[base["_emp"] == row["_emp"]].tolist()
                if idx_list:
                    base.at[idx_list[0], col] = int(row["worked"])

            base = base.rename(columns={
                "_emp":"Emp No", "_name":"Emp Name",
                "JOIN_DATE":"Join Date","CLAIM_START":"Claim Start","ELIGIBLE_END":"Eligible End"
            })

            ordered = (["Emp No","Emp Name","Recruiter","Join Date","Claim Start","Eligible End"]
                       + col_labels + ["Total Workdays"])
            base = base[[c for c in ordered if c in base.columns]].reset_index(drop=True)

            period_label = f"{ps.strftime('%d %b')} – {pe.strftime('%d %b %Y')}"
            tables[pk] = (base, period_label, col_labels, ps, pe)

            summary = base.groupby("Recruiter", as_index=False)["Total Workdays"].sum()
            summary["Rate (RM)"]   = day_rate
            summary["Amount (RM)"] = summary["Total Workdays"] * day_rate
            grand = pd.DataFrame({
                "Recruiter":["TOTAL"],
                "Total Workdays":[summary["Total Workdays"].sum()],
                "Rate (RM)":[day_rate],
                "Amount (RM)":[summary["Amount (RM)"].sum()],
            })
            summaries[pk] = (summary, grand)


        # ─────────────── DISPLAY ───────────────
        for pk in tables:
            df, period_label, col_labels, ps, pe = tables[pk]
            st.subheader(f"📅 Period: {period_label}")

            with st.expander("👁️ Employee Attendance Grid", expanded=True):
                th = "padding:4px 6px;background:#1e3a5f;color:white;font-size:11px;text-align:center;white-space:nowrap;border:1px solid #0d2137;"
                td = "padding:3px 5px;font-size:11px;text-align:center;border:1px solid #d0d0d0;"
                ti = "padding:3px 6px;font-size:11px;text-align:left;border:1px solid #d0d0d0;white-space:nowrap;"

                rows_html = ""
                for _, row in df.iterrows():
                    cs = row["Claim Start"]
                    ee = row["Eligible End"]
                    cells  = f'<td style="{ti}">{row["Emp No"]}</td>'
                    cells += f'<td style="{ti}">{row["Emp Name"]}</td>'
                    cells += f'<td style="{ti}">{row["Recruiter"]}</td>'
                    cells += f'<td style="{ti}">{fmt_date(row["Join Date"])}</td>'
                    cells += f'<td style="{ti}">{fmt_date(cs)}</td>'
                    cells += f'<td style="{ti}">{fmt_date(ee)}</td>'

                    for col in col_labels:
                        val = int(row.get(col, 0))
                        try:
                            parts = col.split("/")
                            # figure out which year this day belongs to
                            m_num = int(parts[1])
                            y_num = ps.year if m_num >= ps.month else pe.year
                            day_ts = pd.Timestamp(y_num, m_num, int(parts[0]))
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

                    cells += f'<td style="{td}font-weight:bold;background:#fff3cd;">{int(row["Total Workdays"])}</td>'
                    rows_html += f"<tr>{cells}</tr>\n"

                header_items = ["Emp No","Name","Recruiter","Join Date","Claim Start","Eligible End"] + col_labels + ["Total"]
                hdr = "".join(f'<th style="{th}">{h}</th>' for h in header_items)

                st.markdown(f"""
                <div style="overflow-x:auto;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.2);margin-bottom:12px;">
                <table style="border-collapse:collapse;font-family:monospace;width:100%;">
                  <thead><tr>{hdr}</tr></thead>
                  <tbody>{rows_html}</tbody>
                </table></div>
                <p style="font-size:11px;color:#888;">✔ = counted day &nbsp;|&nbsp; · = not worked &nbsp;|&nbsp; — = outside claim window</p>
                """, unsafe_allow_html=True)

            st.caption("Per-Recruiter Summary")
            st.dataframe(summaries[pk][0], use_container_width=True)
            st.write("Grand Total")
            st.dataframe(summaries[pk][1], use_container_width=True)
            st.divider()


        # ─────────────── EXCEL ───────────────
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            wb = writer.book
            fmt_hdr   = wb.add_format({"bold":True,"bg_color":"#1e3a5f","font_color":"white","border":1,"align":"center","font_size":9,"text_wrap":True})
            fmt_tick  = wb.add_format({"bg_color":"#d4edda","align":"center","border":1,"font_size":9,"bold":True})
            fmt_dot   = wb.add_format({"align":"center","border":1,"font_size":9,"font_color":"#cccccc"})
            fmt_grey  = wb.add_format({"bg_color":"#f0f0f0","font_color":"#aaaaaa","align":"center","border":1,"font_size":9})
            fmt_total = wb.add_format({"bold":True,"bg_color":"#fff3cd","align":"center","border":1,"font_size":9})
            fmt_info  = wb.add_format({"border":1,"font_size":9})
            fmt_bold  = wb.add_format({"bold":True})
            fmt_title = wb.add_format({"bold":True,"font_size":11,"font_color":"#1e3a5f"})

            for pk in tables:
                df, period_label, col_labels, ps, pe = tables[pk]
                df_summary, df_grand = summaries[pk]

                sheet = period_label[:31]
                ws = wb.add_worksheet(sheet)
                writer.sheets[sheet] = ws

                info_cols = ["Emp No","Emp Name","Recruiter","Join Date","Claim Start","Eligible End"]
                all_hdrs  = info_cols + col_labels + ["Total Workdays"]
                n_info    = len(info_cols)
                n_days    = len(col_labels)

                # Row 0: period title
                ws.merge_range(0, 0, 0, len(all_hdrs)-1, f"Period: {period_label}", fmt_title)
                # Row 1: headers
                for ci, h in enumerate(all_hdrs):
                    ws.write(1, ci, h, fmt_hdr)

                # Column widths
                for ci in range(n_info):
                    ws.set_column(ci, ci, 16)
                for ci in range(n_info, n_info + n_days):
                    ws.set_column(ci, ci, 5)
                ws.set_column(n_info + n_days, n_info + n_days, 12)
                ws.set_row(1, 28)

                # Data rows
                for ri, (_, row) in enumerate(df.iterrows()):
                    r = ri + 2
                    cs = row["Claim Start"]
                    ee = row["Eligible End"]
                    ws.write(r, 0, str(row["Emp No"]),   fmt_info)
                    ws.write(r, 1, str(row["Emp Name"]), fmt_info)
                    ws.write(r, 2, str(row["Recruiter"]),fmt_info)
                    ws.write(r, 3, fmt_date(row["Join Date"]), fmt_info)
                    ws.write(r, 4, fmt_date(cs),              fmt_info)
                    ws.write(r, 5, fmt_date(ee),              fmt_info)

                    for ci, col in enumerate(col_labels):
                        val = int(row.get(col, 0))
                        col_ci = n_info + ci
                        try:
                            parts  = col.split("/")
                            m_num  = int(parts[1])
                            y_num  = ps.year if m_num >= ps.month else pe.year
                            day_ts = pd.Timestamp(y_num, m_num, int(parts[0]))
                            outside = (pd.notna(cs) and pd.notna(ee) and
                                       (day_ts < pd.Timestamp(cs) or day_ts > pd.Timestamp(ee)))
                        except:
                            outside = False

                        if val == 1:
                            ws.write(r, col_ci, "✔", fmt_tick)
                        elif outside:
                            ws.write(r, col_ci, "—", fmt_grey)
                        else:
                            ws.write(r, col_ci, "·", fmt_dot)

                    ws.write(r, n_info + n_days, int(row["Total Workdays"]), fmt_total)

                # Summary below
                s0 = len(df) + 4
                ws.write(s0,   0, f"Period: {period_label}", fmt_bold)
                ws.write(s0+1, 0, "Per-Recruiter Summary",   fmt_bold)
                df_summary.to_excel(writer, sheet_name=sheet, startrow=s0+2, index=False)
                s1 = s0 + len(df_summary) + 4
                ws.write(s1,   0, "Grand Total", fmt_bold)
                df_grand.to_excel(writer, sheet_name=sheet, startrow=s1+1, index=False)

        st.download_button(
            "⬇️ Download Excel",
            buffer.getvalue(),
            file_name="claim_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)

else:
    st.info("Upload both files to continue.")
