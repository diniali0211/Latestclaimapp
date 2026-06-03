import io
import calendar
import numpy as np
import pandas as pd
import streamlit as st


# ───────────────────────── Company Cutoff Config ─────────────────────────
# Each entry: company keyword (lowercase) → (cutoff_start_day, cutoff_end_day)
# cutoff_start_day: the day of month the claim period STARTS
# cutoff_end_day:   the day of month the claim period ENDS (in next/same month)
# e.g. Dexcom: 11th–10th means period runs from 11th of prev month to 10th of current month

COMPANY_CUTOFFS = {
    "dexcom":  (11, 10),
    "micron":  (21, 20),
    "eti":     (21, 20),
    "ybs":     (21, 20),
    "atns":    (7,  6),
    "sustio":  (16, 15),
    "jabil":   (24, 23),
    "wd":      (1,  None),   # None = end of month (30th/31st)
    "shopee":  (1,  None),
}

def detect_company(recruiter_name: str) -> str | None:
    """Return matched company key from recruiter name, or None."""
    if pd.isna(recruiter_name):
        return None
    s = str(recruiter_name).lower()
    for key in COMPANY_CUTOFFS:
        if key in s:
            return key
    return None


def get_claim_periods_for_month(year: int, month: int, company_key: str | None):
    """
    Return the (period_start, period_end) date range that corresponds to
    the billing cycle containing most of `month` for the given company.

    If company_key is None or unrecognised, fall back to calendar month.
    """
    if company_key is None or company_key not in COMPANY_CUTOFFS:
        # fallback: full calendar month
        last = calendar.monthrange(year, month)[1]
        return pd.Timestamp(year, month, 1), pd.Timestamp(year, month, last)

    start_day, end_day = COMPANY_CUTOFFS[company_key]

    # Period starts on start_day of the PREVIOUS month, ends on end_day of `month`
    # e.g. Dexcom for "June" → 11 May to 10 June
    prev_year  = year if month > 1 else year - 1
    prev_month = month - 1 if month > 1 else 12

    # Clamp start_day to valid day in prev_month
    max_day_prev = calendar.monthrange(prev_year, prev_month)[1]
    actual_start_day = min(start_day, max_day_prev)
    period_start = pd.Timestamp(prev_year, prev_month, actual_start_day)

    if end_day is None:
        # End of current month
        last = calendar.monthrange(year, month)[1]
        period_end = pd.Timestamp(year, month, last)
    else:
        max_day_curr = calendar.monthrange(year, month)[1]
        actual_end_day = min(end_day, max_day_curr)
        period_end = pd.Timestamp(year, month, actual_end_day)

    return period_start, period_end


# ───────────────────────── UI ─────────────────────────
st.set_page_config(page_title="Agency Claim Table Generator", layout="wide")
st.title("🧾 Agency Claim Table Generator")

st.markdown("""
**Eligibility rule**

• All employees → claim window starts from **join date**, runs for **3 months − 1 day**
""")

# ───────────────────────── Company Cutoff Reference ─────────────────────────
with st.expander("📋 Company Cutoff Day Reference", expanded=False):
    cutoff_rows = [
        {"Company": "Dexcom",        "Period": "11th → 10th"},
        {"Company": "Micron / ETI / YBS", "Period": "21st → 20th"},
        {"Company": "ATnS",          "Period": "7th → 6th"},
        {"Company": "Sustio",        "Period": "16th → 15th"},
        {"Company": "Jabil",         "Period": "24th → 23rd"},
        {"Company": "WD / Shopee",   "Period": "1st → 30th/31st"},
    ]
    st.table(pd.DataFrame(cutoff_rows))


# ───────────────────────── Sidebar ─────────────────────────
with st.sidebar:

    st.header("Settings")

    hours_per_day = st.number_input(
        "Hours considered 1 workday", 1.0, 24.0, 8.0
    )

    grace_minutes = st.number_input(
        "Grace window (minutes)", 0, 120, 15
    )

    day_rate = st.number_input(
        "Rate per claim day (RM)", 0.0, 1000.0, 3.0
    )

    day_first = st.checkbox(
        "Dates are DD/MM/YYYY", value=True
    )

    st.divider()
    st.subheader("Company Cutoff Override")
    st.caption(
        "If a recruiter name doesn't auto-detect, you can manually pick the company cutoff below. "
        "Leave as 'Auto-detect' to use name matching."
    )
    manual_company = st.selectbox(
        "Force company cutoff for ALL records",
        ["Auto-detect", "Dexcom", "Micron/ETI/YBS", "ATnS", "Sustio", "Jabil", "WD/Shopee"],
    )

    MANUAL_MAP = {
        "Auto-detect":    None,
        "Dexcom":         "dexcom",
        "Micron/ETI/YBS": "micron",
        "ATnS":           "atns",
        "Sustio":         "sustio",
        "Jabil":          "jabil",
        "WD/Shopee":      "wd",
    }
    forced_company = MANUAL_MAP[manual_company]

effective_threshold = hours_per_day - grace_minutes / 60


# ───────────────────────── Upload ─────────────────────────
att_file = st.file_uploader(
    "Upload Timecard", type=["csv","xlsx","xls"]
)

mst_file = st.file_uploader(
    "Upload Masterlist", type=["xlsx","xls"]
)


# ───────────────────────── Helpers ─────────────────────────
def clean_name(s):
    if pd.isna(s):
        return ""
    return str(s).strip().replace("\u2019","'")


def clean_emp(s):
    if pd.isna(s):
        return ""
    return str(s).strip().replace(".0","")


def parse_time(val):
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    if s == "":
        return np.nan
    try:
        t = pd.to_datetime(s, errors="coerce")
        if pd.notna(t):
            return t.hour + t.minute / 60
    except:
        pass
    try:
        return float(s)
    except:
        return np.nan


def calc_hours(t_in, t_out):
    if pd.isna(t_in) or pd.isna(t_out):
        return 0.0
    dur = t_out - t_in
    if dur < 0:
        dur += 24
    return min(max(dur, 0), 24)


# ───────────────────────── Claim window logic ─────────────────────────
def claim_window(j):
    """
    For ANY join date: claim starts on join date, ends 3 months − 1 day later.
    """
    if pd.isna(j):
        return pd.NaT, pd.NaT

    claim_start  = j.normalize()
    eligible_end = (
        claim_start
        + pd.DateOffset(months=3)
        - pd.Timedelta(days=1)
    ).normalize()

    return claim_start, eligible_end


# ───────────────────────── Main ─────────────────────────
if att_file and mst_file:

    try:

        att = (
            pd.read_excel(att_file)
            if not att_file.name.endswith(".csv")
            else pd.read_csv(att_file)
        )
        mst = pd.read_excel(mst_file)

        # ───── Column Mapping ─────
        st.subheader("1️⃣ Map Timecard Columns")
        t_emp  = st.selectbox("Emp No (Timecard)", att.columns)
        t_name = st.selectbox("Name (Timecard)", att.columns)
        t_date = st.selectbox("Date", att.columns)
        t_in   = st.selectbox("Time In", att.columns)
        t_out  = st.selectbox("Time Out", att.columns)

        st.subheader("2️⃣ Map Masterlist Columns")
        m_emp  = st.selectbox("Emp No (Masterlist)", mst.columns)
        m_name = st.selectbox("Name (Masterlist)", mst.columns)
        m_join = st.selectbox("Joined Date", mst.columns)
        m_recr = st.selectbox("Recruiter", ["(none)"] + mst.columns.tolist())
        m_comp = st.selectbox(
            "Company column (for cutoff detection)",
            ["(none)"] + mst.columns.tolist(),
        )

        # ───── Normalize Timecard ─────
        att["__Emp"]   = att[t_emp].apply(clean_emp)
        att["__Name"]  = att[t_name].apply(clean_name)
        att["__Date"]  = pd.to_datetime(att[t_date], dayfirst=day_first, errors="coerce")
        att["__InH"]   = att[t_in].apply(parse_time)
        att["__OutH"]  = att[t_out].apply(parse_time)
        att["__Hours"] = [calc_hours(i, o) for i, o in zip(att["__InH"], att["__OutH"])]

        # ───── Normalize Masterlist ─────
        mst["__Emp"]  = mst[m_emp].apply(clean_emp)
        mst["__Name"] = mst[m_name].apply(clean_name)
        mst[m_join]   = pd.to_datetime(mst[m_join], errors="coerce")

        # ───── Join date mapping ─────
        join_by_emp  = dict(zip(mst["__Emp"],  mst[m_join]))
        join_by_name = dict(zip(mst["__Name"], mst[m_join]))

        att["JOIN_DATE"] = att["__Emp"].map(join_by_emp)
        missing = att["JOIN_DATE"].isna()
        att.loc[missing, "JOIN_DATE"] = att.loc[missing, "__Name"].map(join_by_name)

        # ───── Company mapping ─────
        if m_comp != "(none)":
            comp_by_emp  = dict(zip(mst["__Emp"],  mst[m_comp]))
            comp_by_name = dict(zip(mst["__Name"], mst[m_comp]))
            att["__Company"] = att["__Emp"].map(comp_by_emp)
            miss = att["__Company"].isna()
            att.loc[miss, "__Company"] = att.loc[miss, "__Name"].map(comp_by_name)
        else:
            att["__Company"] = np.nan

        # Determine company key per row
        if forced_company:
            att["__CompanyKey"] = forced_company
        else:
            att["__CompanyKey"] = att["__Company"].apply(detect_company)

        # ───── Recruiter mapping ─────
        if m_recr != "(none)":
            recr_by_emp  = dict(zip(mst["__Emp"],  mst[m_recr]))
            recr_by_name = dict(zip(mst["__Name"], mst[m_recr]))
            att["Recruiter"] = att["__Emp"].map(recr_by_emp)
            miss = att["Recruiter"].isna()
            att.loc[miss, "Recruiter"] = att.loc[miss, "__Name"].map(recr_by_name)
        else:
            att["Recruiter"] = "Unassigned"

        att["Recruiter"] = att["Recruiter"].fillna("Unassigned")

        # ───── Claim window ─────
        att[["CLAIM_START", "ELIGIBLE_END"]] = (
            att["JOIN_DATE"].apply(lambda d: pd.Series(claim_window(d)))
        )

        # ───── Daily aggregation ─────
        daily = (
            att.groupby(
                ["__Emp", "__Name", "Recruiter", "__CompanyKey",
                 "JOIN_DATE", "CLAIM_START", "ELIGIBLE_END", "__Date"],
                as_index=False,
            )
            .agg(Total_Hours=("__Hours", "sum"))
        )

        daily["Worked_Day"] = np.where(
            (daily["Total_Hours"] >= effective_threshold) &
            (daily["__Date"] >= daily["CLAIM_START"]) &
            (daily["__Date"] <= daily["ELIGIBLE_END"]),
            1,
            0,
        )

        # ───── Determine billing periods ─────
        # For company-cutoff schemes the "month" label is the end-month of the period.
        # We assign each attendance date to its billing period end-month.

        def date_to_billing_month(row):
            d = row["__Date"]
            ck = row["__CompanyKey"]
            if pd.isna(d):
                return pd.NaT
            if ck is None or ck not in COMPANY_CUTOFFS:
                return d.to_period("M")
            start_day, end_day = COMPANY_CUTOFFS[ck]
            # If date is on/after start_day → it belongs to the period ending next month
            if d.day >= start_day:
                # period ends in next month
                nm = d.month + 1 if d.month < 12 else 1
                ny = d.year if d.month < 12 else d.year + 1
                return pd.Period(f"{ny}-{nm:02d}", freq="M")
            else:
                # period ends this month
                return d.to_period("M")

        daily["__BillingMonth"] = daily.apply(date_to_billing_month, axis=1)

        months = sorted(daily["__BillingMonth"].dropna().unique())

        tables    = {}
        summaries = {}

        # ───── Build Monthly Tables ─────
        for m in months:

            month_df = daily[daily["__BillingMonth"] == m].copy()

            # Determine company key for this month (use most common)
            ck = (
                month_df["__CompanyKey"].dropna().mode().iloc[0]
                if not month_df["__CompanyKey"].dropna().empty
                else None
            )

            period_start, period_end = get_claim_periods_for_month(m.year, m.month, ck)

            # Number of calendar days in the period
            period_days = (period_end - period_start).days + 1
            day_labels  = [
                (period_start + pd.Timedelta(days=i))
                for i in range(period_days)
            ]

            base = (
                month_df.groupby(
                    ["__Emp", "__Name", "Recruiter",
                     "JOIN_DATE", "CLAIM_START", "ELIGIBLE_END"],
                    as_index=False,
                )["Worked_Day"]
                .sum()
                .rename(columns={"Worked_Day": "Total Workdays"})
            )

            # Add a column per day in the period, labelled as "DD/MM"
            for dl in day_labels:
                base[dl.strftime("%-d/%-m")] = 0

            for _, r in month_df.iterrows():
                d = r["__Date"]
                if pd.isna(d):
                    continue
                col = d.strftime("%-d/%-m")
                if col not in base.columns:
                    continue
                mask = (
                    (base["__Emp"]   == r["__Emp"]) &
                    (base["__Name"]  == r["__Name"])
                )
                base.loc[mask, col] = r["Worked_Day"]

            base = base.rename(columns={
                "__Emp":         "Emp No",
                "__Name":        "Emp Name",
                "JOIN_DATE":     "Join Date",
                "CLAIM_START":   "Claim Start",
                "ELIGIBLE_END":  "Eligible End",
            })

            day_cols = [dl.strftime("%-d/%-m") for dl in day_labels]
            ordered_cols = (
                ["Emp No", "Emp Name", "Recruiter",
                 "Join Date", "Claim Start", "Eligible End"]
                + day_cols
                + ["Total Workdays"]
            )
            base = base[[c for c in ordered_cols if c in base.columns]]

            period_label = (
                f"{period_start.strftime('%d %b')} – {period_end.strftime('%d %b %Y')}"
            )
            tables[str(m)] = (base, period_label)

            summary = (
                base.groupby("Recruiter", as_index=False)["Total Workdays"].sum()
            )
            summary["Rate (RM)"]   = day_rate
            summary["Amount (RM)"] = summary["Total Workdays"] * day_rate

            grand = pd.DataFrame({
                "Recruiter":      ["TOTAL"],
                "Total Workdays": [summary["Total Workdays"].sum()],
                "Rate (RM)":      [day_rate],
                "Amount (RM)":    [summary["Amount (RM)"].sum()],
            })

            summaries[str(m)] = (summary, grand)

        # ───── Display ─────
        for m in tables:
            df, period_label = tables[m]
            st.subheader(f"📅 {m}  —  Period: {period_label}")
            st.dataframe(df, use_container_width=True)
            st.caption("Per-Recruiter Summary")
            st.dataframe(summaries[m][0], use_container_width=True)
            st.write("Grand Total")
            st.dataframe(summaries[m][1], use_container_width=True)

        # ───── Excel Export ─────
        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:

            for m in tables:
                df, period_label = tables[m]
                df_summary, df_grand = summaries[m]

                sheet = m[:31]
                df.to_excel(writer, sheet_name=sheet, index=False)

                ws  = writer.sheets[sheet]
                wb  = writer.book
                bold = wb.add_format({"bold": True})

                for i, col in enumerate(df.columns):
                    ws.set_column(i, i, 15)

                start = len(df) + 2
                ws.write(start, 0, f"Period: {period_label}", bold)
                ws.write(start + 1, 0, "Per-Recruiter Summary", bold)

                df_summary.to_excel(
                    writer, sheet_name=sheet, startrow=start + 2, index=False
                )

                start = start + len(df_summary) + 4
                ws.write(start, 0, "Grand Total", bold)
                df_grand.to_excel(
                    writer, sheet_name=sheet, startrow=start + 1, index=False
                )

        st.download_button(
            "⬇️ Download Excel (one sheet per billing period)",
            buffer.getvalue(),
            file_name="claim_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)

else:
    st.info("Upload both files to continue.")
