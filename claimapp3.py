import io
import calendar
import numpy as np
import pandas as pd
import streamlit as st


# ───────────────────────── UI ─────────────────────────
st.set_page_config(page_title="Agency Claim Table Generator", layout="wide")
st.title("🧾 Agency Claim Table Generator")

st.markdown("""
**Eligibility rules**

• Joined **before 1 Jan 2026** → claim immediately for **3 months**  
• Joined **on/after 1 Jan 2026** → **wait 1 calendar month**, then claim for **3 months − 1 day**
""")


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

    return (
        str(s)
        .strip()
        .replace("’","'")
    )


def clean_emp(s):

    if pd.isna(s):
        return ""

    return (
        str(s)
        .strip()
        .replace(".0","")
    )


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

    return min(max(dur,0),24)


# ───────────────────────── Main ─────────────────────────
if att_file and mst_file:

    try:

        att = pd.read_excel(att_file) if not att_file.name.endswith(".csv") else pd.read_csv(att_file)
        mst = pd.read_excel(mst_file)


        # ───── Column Mapping ─────
        st.subheader("1️⃣ Map Timecard Columns")

        t_emp = st.selectbox("Emp No (Timecard)", att.columns)
        t_name = st.selectbox("Name (Timecard)", att.columns)
        t_date = st.selectbox("Date", att.columns)
        t_in = st.selectbox("Time In", att.columns)
        t_out = st.selectbox("Time Out", att.columns)

        st.subheader("2️⃣ Map Masterlist Columns")

        m_emp = st.selectbox("Emp No (Masterlist)", mst.columns)
        m_name = st.selectbox("Name (Masterlist)", mst.columns)
        m_join = st.selectbox("Joined Date", mst.columns)
        m_recr = st.selectbox("Recruiter", ["(none)"] + mst.columns.tolist())


        # ───── Normalize Timecard ─────
        att["__Emp"] = att[t_emp].apply(clean_emp)
        att["__Name"] = att[t_name].apply(clean_name)

        att["__Date"] = pd.to_datetime(
            att[t_date],
            dayfirst=day_first,
            errors="coerce"
        )

        att["__InH"] = att[t_in].apply(parse_time)
        att["__OutH"] = att[t_out].apply(parse_time)

        att["__Hours"] = [
            calc_hours(i,o) for i,o in zip(att["__InH"],att["__OutH"])
        ]


        # ───── Normalize Masterlist ─────
        mst["__Emp"] = mst[m_emp].apply(clean_emp)
        mst["__Name"] = mst[m_name].apply(clean_name)

        mst[m_join] = pd.to_datetime(mst[m_join], errors="coerce")


        # ───── Mapping (EMP NO FIRST) ─────
        join_by_emp = dict(zip(mst["__Emp"], mst[m_join]))
        join_by_name = dict(zip(mst["__Name"], mst[m_join]))

        att["JOIN_DATE"] = att["__Emp"].map(join_by_emp)

        missing = att["JOIN_DATE"].isna()

        att.loc[missing,"JOIN_DATE"] = (
            att.loc[missing,"__Name"].map(join_by_name)
        )


        # ───── Recruiter mapping ─────
        if m_recr != "(none)":

            recr_by_emp = dict(zip(mst["__Emp"], mst[m_recr]))
            recr_by_name = dict(zip(mst["__Name"], mst[m_recr]))

            att["Recruiter"] = att["__Emp"].map(recr_by_emp)

            miss = att["Recruiter"].isna()

            att.loc[miss,"Recruiter"] = (
                att.loc[miss,"__Name"].map(recr_by_name)
            )

        else:

            att["Recruiter"] = "Unassigned"

        att["Recruiter"] = att["Recruiter"].fillna("Unassigned")


        # ───── Claim window logic ─────
        NEW_POLICY = pd.Timestamp("2026-01-01")

        def claim_window(j):

            if pd.isna(j):
                return pd.NaT, pd.NaT

            if j >= NEW_POLICY:
                claim_start = (j + pd.DateOffset(months=1)).normalize()
            else:
                claim_start = j.normalize()

            eligible_end = (
                claim_start
                + pd.DateOffset(months=3)
                - pd.Timedelta(days=1)
            ).normalize()

            return claim_start, eligible_end


        att[["CLAIM_START","ELIGIBLE_END"]] = (
            att["JOIN_DATE"]
            .apply(lambda d: pd.Series(claim_window(d)))
        )


        # ───── Daily aggregation ─────
        daily = (
            att.groupby(
                [
                    "__Emp",
                    "__Name",
                    "Recruiter",
                    "JOIN_DATE",
                    "CLAIM_START",
                    "ELIGIBLE_END",
                    "__Date"
                ],
                as_index=False
            )
            .agg(Total_Hours=("__Hours","sum"))
        )


        daily["Worked_Day"] = np.where(
            (daily["Total_Hours"] >= effective_threshold) &
            (daily["__Date"] >= daily["CLAIM_START"]) &
            (daily["__Date"] <= daily["ELIGIBLE_END"]),
            1,
            0
        )


        months = sorted(
            daily["__Date"].dt.to_period("M").unique()
        )


        tables = {}
        summaries = {}


        # ───── Build Monthly Tables ─────
        for m in months:

            month_df = daily[
                daily["__Date"].dt.to_period("M") == m
            ]

            days = calendar.monthrange(m.year,m.month)[1]

            base = (
                month_df.groupby(
                    [
                        "__Emp",
                        "__Name",
                        "Recruiter",
                        "JOIN_DATE",
                        "CLAIM_START",
                        "ELIGIBLE_END"
                    ],
                    as_index=False
                )["Worked_Day"]
                .sum()
                .rename(columns={"Worked_Day":"Total Workdays"})
            )


            for d in range(1,days+1):
                base[str(d)] = 0


            for _,r in month_df.iterrows():

                mask = (
                    (base["__Emp"] == r["__Emp"]) &
                    (base["__Name"] == r["__Name"])
                )

                base.loc[
                    mask,
                    str(r["__Date"].day)
                ] = r["Worked_Day"]


            base = base.rename(columns={
                "__Emp":"Emp No",
                "__Name":"Emp Name",
                "JOIN_DATE":"Join Date",
                "CLAIM_START":"Claim Start",
                "ELIGIBLE_END":"Eligible End"
            })


            day_cols = [str(d) for d in range(1,days+1)]

            ordered_cols = (
                [
                    "Emp No",
                    "Emp Name",
                    "Recruiter",
                    "Join Date",
                    "Claim Start",
                    "Eligible End"
                ]
                + day_cols
                + ["Total Workdays"]
            )

            base = base[ordered_cols]

            tables[str(m)] = base


            summary = (
                base.groupby("Recruiter",as_index=False)["Total Workdays"]
                .sum()
            )

            summary["Rate (RM)"] = day_rate
            summary["Amount (RM)"] = summary["Total Workdays"] * day_rate


            grand = pd.DataFrame({
                "Recruiter":["TOTAL"],
                "Total Workdays":[summary["Total Workdays"].sum()],
                "Rate (RM)":[day_rate],
                "Amount (RM)":[summary["Amount (RM)"].sum()]
            })


            summaries[str(m)] = (summary,grand)


        # ───── Display ─────
        for m in tables:

            st.subheader(f"📅 {m}")
            st.dataframe(tables[m], use_container_width=True)

            st.caption("Per-Recruiter Summary")
            st.dataframe(summaries[m][0], use_container_width=True)

            st.write("Grand Total")
            st.dataframe(summaries[m][1], use_container_width=True)


        # ───── Excel Export ─────
        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:

            for m in tables:

                sheet = m[:31]

                df_claim = tables[m]
                df_summary, df_grand = summaries[m]

                df_claim.to_excel(writer, sheet_name=sheet, index=False)

                ws = writer.sheets[sheet]
                wb = writer.book
                bold = wb.add_format({"bold":True})

                for i,col in enumerate(df_claim.columns):
                    ws.set_column(i,i,15)

                start = len(df_claim) + 2

                ws.write(start,0,"Per-Recruiter Summary",bold)

                df_summary.to_excel(
                    writer,
                    sheet_name=sheet,
                    startrow=start+1,
                    index=False
                )

                start = start + len(df_summary) + 3

                ws.write(start,0,"Grand Total",bold)

                df_grand.to_excel(
                    writer,
                    sheet_name=sheet,
                    startrow=start+1,
                    index=False
                )


        st.download_button(
            "⬇️ Download Excel (one sheet per month)",
            buffer.getvalue(),
            file_name="claim_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


    except Exception as e:

        st.error(f"Error: {e}")
        st.exception(e)

else:

    st.info("Upload both files to continue.")

