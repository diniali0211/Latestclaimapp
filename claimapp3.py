import io
import re
import calendar
from pathlib import Path
from datetime import datetime as dt

import numpy as np
import pandas as pd
import streamlit as st


# ───────────────────────────── UI header ─────────────────────────────
st.set_page_config(page_title="Agency Claim Table Generator (Auto Recruiter)", layout="wide")
st.title("🧾 Agency Claim Table Generator (Auto-detect Recruiter)")

st.markdown(
    """
- **Masterlist**: must include **Name**, **Joined Date**, and **Recruiter** (optional but recommended).  
- **Timecard**: must include **Emp No**, **Name**, **Date**, and **one IN + one OUT** column.  
- A workday counts **1** if **daily hours ≥ (hours per day − grace)** and **not on leave**.
- Eligibility: **JOIN_DATE → JOIN_DATE + 3 months − 1 day** (inclusive).
"""
)

# ───────────────────────────── Sidebar ─────────────────────────────
with st.sidebar:
    st.header("Settings")
    hours_per_day = st.number_input("Hours considered 1 workday", 1.0, 24.0, 8.0, 0.5)
    grace_minutes = st.number_input("Grace window (minutes)", 0, 120, 15, 5)
    counting_rule = st.selectbox(
        "Counting rule",
        ["Per-day ≥ threshold", "Floor(total hours ÷ threshold)"],
        index=0,
        help=(
            "• Per-day ≥ threshold: a day counts if daily hours ≥ (threshold − grace) AND no leave.\n"
            "• Floor(total hours ÷ threshold): floor(total_hours / threshold)."
        ),
    )
    day_rate = st.number_input("Rate per claim day (RM)", 0.0, 1000.0, 3.0, 0.5)
    exclude_not_in_master = st.checkbox("Exclude employees not in Masterlist", value=True)
    # NEW: date parsing control
    day_first = st.checkbox("Dates are day-first (DD/MM/YYYY)", value=True,
                            help="Turn OFF if your timecard uses MM/DD/YYYY")

effective_threshold = max(0.0, float(hours_per_day) - float(grace_minutes) / 60.0)

# ───────────────────────────── Uploaders ─────────────────────────────
att_file = st.file_uploader("Upload **Timecard** (CSV/XLSX/XLS/XLSM)", type=["csv", "xlsx", "xls", "xlsm"])
mst_file = st.file_uploader("Upload **Masterlist** (XLSX/XLS/XLSM)", type=["xlsx", "xls", "xlsm"])


# ───────────────────────────── Helpers ─────────────────────────────
def ensure_unique_headers(df: pd.DataFrame) -> pd.DataFrame:
    counts, new_cols = {}, []
    for c in df.columns:
        base = str(c).strip()
        if base not in counts:
            counts[base] = 1
            new_cols.append(base)
        else:
            counts[base] += 1
            new_cols.append(f"{base}_{counts[base]}")
    out = df.copy()
    out.columns = new_cols
    return out


def _norm_empid(s):
    if pd.isna(s):
        return np.nan
    s = str(s).strip().upper().replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _norm_name(s):
    if pd.isna(s):
        return ""
    return str(s).strip()


def _norm_recruiter(s):
    if pd.isna(s) or str(s).strip() == "":
        return "Unassigned"
    return str(s).strip().title()


def _is_leave(val):
    if val is None:
        return False
    s = str(val).strip().lower()
    if s in ("", "-", "nan", "0", "0.0"):
        return False
    keys = ["unpaid", "annual", "absent", "emergency", "medical", "sick", "mc", "leave"]
    return any(k in s for k in keys)


def _to_hours_any(val):
    """Robust hour parser for IN/OUT."""
    if val is None:
        return np.nan
    if isinstance(val, (int, float)) and not pd.isna(val):
        f = float(val)
        if 0 <= f <= 1.5:
            return f * 24.0
        if 0 <= f <= 24:
            return f
        return np.nan
    s = str(val).strip()
    if s == "":
        return np.nan
    s_num = s.replace(",", ".")
    try:
        f = float(s_num)
        if 0 <= f <= 1.5:
            return f * 24.0
        if 0 <= f <= 24 and ":" not in s:
            return f
    except Exception:
        pass
    mhm = re.match(r"^\s*(\d{1,2})\.(\d{2})\s*$", s_num)
    if mhm:
        h = int(mhm.group(1)); mi = int(mhm.group(2))
        if 0 <= h <= 24 and 0 <= mi < 60:
            return h + mi / 60.0
    m = re.match(r"^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([APap][Mm])?\s*$", s)
    if m:
        h = int(m.group(1)); mi = int(m.group(2)); se = int(m.group(3) or 0); ampm = m.group(4)
        if ampm:
            if ampm.lower() == "pm" and h < 12: h += 12
            if ampm.lower() == "am" and h == 12: h = 0
        if 0 <= h <= 24 and 0 <= mi < 60 and 0 <= se < 60:
            return h + mi / 60.0 + se / 3600.0
    t = pd.to_datetime(s, errors="coerce")
    if pd.notna(t):
        return t.hour + t.minute / 60.0 + t.second / 3600.0
    td = pd.to_timedelta(s, errors="coerce")
    if pd.notna(td):
        return td.total_seconds() / 3600.0
    return np.nan


def _pair_duration(in_v, out_v):
    hi = _to_hours_any(in_v); ho = _to_hours_any(out_v)
    if pd.isna(hi) or pd.isna(ho):
        return 0.0
    dur = ho - hi
    if dur < 0:
        dur += 24.0
    return float(np.clip(dur, 0.0, 24.0))


# NEW: robust date parser with day-first toggle + Excel serial fallback
def _parse_dates(series, day_first_flag=True):
    s = series.astype(str).str.strip().replace({"": np.nan, "-": np.nan})
    # Pass 1: string dates, honoring day-first if toggled
    d = pd.to_datetime(s, errors="coerce", dayfirst=day_first_flag)
    # Pass 2: Excel serials like 45503 or "45503.0"
    is_num = d.isna() & s.str.fullmatch(r"\d+(\.0)?").fillna(False)
    if is_num.any():
        nums = pd.to_numeric(s[is_num], errors="coerce")
        d.loc[is_num] = pd.to_datetime(nums, unit="d", origin="1899-12-30")
    return d


def load_masterlist(file_like) -> pd.DataFrame:
    """Load masterlist and try to detect header row."""
    try:
        raw = pd.read_excel(file_like, header=None, dtype=str, keep_default_na=False)
    except Exception:
        raw = pd.read_excel(file_like, header=None)
    hdr = None
    for i in range(min(50, len(raw))):
        vals = raw.iloc[i].astype(str).str.strip().str.lower().tolist()
        has_name = any("name" in v for v in vals)
        has_join = any(any(k in v for k in ["join", "start", "date", "doj", "hire"]) for v in vals)
        if has_name and has_join:
            hdr = i; break
    if hdr is None:
        return pd.read_excel(file_like, sheet_name=0, dtype=str, keep_default_na=False)
    return pd.read_excel(file_like, sheet_name=0, header=hdr, dtype=str, keep_default_na=False)


def guess_timecard_columns(df: pd.DataFrame):
    """Return (Date, Name, EmpNo, IN, OUT, Leave)."""
    cols = list(df.columns)
    def first_match(names):
        low = {c.lower().strip(): c for c in cols}
        for n in names:
            if n in low: return low[n]
        for c in cols:
            if any(n in c.lower() for n in names): return c
        return None
    date_col = first_match(["date", "work date", "day"])
    name_col = first_match(["name", "employee name"])
    emp_col  = first_match(["emp no", "employee id", "new id", "employee number", "id"])
    in_col, out_col = None, None
    for c in cols:
        if re.fullmatch(r"in(_\d+)?", c.lower()): in_col = c; break
    for c in cols:
        cl = c.lower()
        if re.fullmatch(r"out(_\d+)?", cl) and not re.search(r"\be\s*out\b|\bearly\s*out\b", cl):
            out_col = c; break
    leave_col = first_match(["leave", "leave type", "absence"])
    return date_col, name_col, emp_col, in_col, out_col, leave_col


# ── Auto-detect masterlist columns (with robust Recruiter detection) ──
def autodetect_master_columns(mst: pd.DataFrame):
    cols = mst.columns.tolist()
    low = {c.lower().strip(): c for c in cols}

    def pick(*keys, contains=None):
        for k in keys:
            if k in low: return low[k]
        if contains:
            for c in cols:
                if any(k in c.lower() for k in contains): return c
        return None

    # Name
    name_col = pick("name", contains=["employee name", "full name"]) or cols[0]
    # Emp No
    emp_col  = pick("new id", "emp no", "employee id", "employee number", "id")
    # Join
    join_col = pick("joined date", "join date", "date joined", "doj", "start date", "hire date",
                    contains=["join", "date"]) or cols[0]

    # Recruiter: prefer header containing 'recruit'
    recr_col = None
    for c in cols:
        if "recruit" in c.lower():
            recr_col = c
            break
    if recr_col is None:
        # Heuristic: looks like names
        candidates = []
        for c in cols:
            if c in {name_col, emp_col, join_col}: continue
            series = mst[c].astype(str).str.strip()
            nonempty = series[series.str.len() > 0]
            if len(nonempty) == 0: continue
            alpha_ratio = (nonempty.str.contains(r"[A-Za-z]")).mean()
            digit_ratio = (nonempty.str.fullmatch(r"\d+(\.0)?").fillna(False)).mean()
            if alpha_ratio >= 0.7 and digit_ratio <= 0.2:
                candidates.append((c, alpha_ratio - digit_ratio))
        if candidates:
            candidates.sort(key=lambda x: x[1], reverse=True)
            recr_col = candidates[0][0]

    return name_col, emp_col, join_col, recr_col


# ───────────────────────────── Main ─────────────────────────────
if att_file and mst_file:
    try:
        # ----- Load timecard -----
        if Path(att_file.name).suffix.lower() != ".csv":
            xls = pd.ExcelFile(att_file)
            att_sheet = st.selectbox("Timecard sheet", xls.sheet_names, index=0)
            att_raw = pd.read_excel(att_file, sheet_name=att_sheet, dtype=str, keep_default_na=False)
        else:
            att_raw = pd.read_csv(att_file, dtype=str, keep_default_na=False)

        att_raw = ensure_unique_headers(att_raw)

        # ----- Load masterlist -----
        mst = load_masterlist(mst_file)
        mst = ensure_unique_headers(mst)

        # ----- Map timecard columns -----
        st.subheader("1) Map Timecard Columns")
        d_guess, n_guess, e_guess, in_guess, out_guess, l_guess = guess_timecard_columns(att_raw)

        c1, c2, c3 = st.columns(3)
        with c1:
            t_date = st.selectbox("Timecard — Date", att_raw.columns.tolist(),
                                  index=(att_raw.columns.get_loc(d_guess) if d_guess in att_raw.columns else 0))
        with c2:
            t_name = st.selectbox("Timecard — Name", att_raw.columns.tolist(),
                                  index=(att_raw.columns.get_loc(n_guess) if n_guess in att_raw.columns else 0))
        with c3:
            t_emp = st.selectbox("Timecard — Emp No", att_raw.columns.tolist(),
                                 index=(att_raw.columns.get_loc(e_guess) if e_guess in att_raw.columns else 0))
        c4, c5 = st.columns(2)
        with c4:
            t_in = st.selectbox("Timecard — IN (single)", att_raw.columns.tolist(),
                                index=(att_raw.columns.get_loc(in_guess) if in_guess in att_raw.columns else 0))
        with c5:
            t_out = st.selectbox("Timecard — OUT (single)", att_raw.columns.tolist(),
                                 index=(att_raw.columns.get_loc(out_guess) if out_guess in att_raw.columns else 0))
        t_leave = st.selectbox("Timecard — Leave (optional)",
                               ["(none)"] + att_raw.columns.tolist(),
                               index=(att_raw.columns.get_loc(l_guess) + 1 if l_guess in att_raw.columns else 0))

        # ----- Normalize timecard -----
        att = att_raw.copy()
        # NEW: robust date parsing + optional clean
        att["__Date"] = _parse_dates(att[t_date], day_first_flag=day_first)
        att = att[att["__Date"].notna()].copy()

        att["__Name"] = att[t_name].apply(_norm_name)
        att["__EmpID"] = att[t_emp].apply(_norm_empid)

        att["__InRaw"] = att[t_in]
        att["__OutRaw"] = att[t_out]
        att["__InH"] = att["__InRaw"].apply(_to_hours_any)
        att["__OutH"] = att["__OutRaw"].apply(_to_hours_any)
        att["__Hours"] = [_pair_duration(r[t_in], r[t_out]) for _, r in att.iterrows()]

        if t_leave != "(none)":
            att["__LeaveFlagRow"] = att[t_leave].apply(_is_leave)
            att["__LeaveText"] = att[t_leave].astype(str)
        else:
            att["__LeaveFlagRow"] = False
            att["__LeaveText"] = ""

        with st.expander("Preview timecard parsing"):
            st.dataframe(att[["__Date","__Name","__EmpID","__InRaw","__OutRaw","__InH","__OutH","__Hours"]].head(15),
                         use_container_width=True)

        # ----- Auto-detect masterlist columns (with override) -----
        st.subheader("2) Masterlist Mapping (auto-detected, you can override)")

        m_name_d, m_emp_d, m_join_d, m_recr_d = autodetect_master_columns(mst)
        mcols = mst.columns.tolist()

        c6, c7 = st.columns(2)
        with c6:
            mst_name = st.selectbox("Masterlist — Name", mcols, index=mcols.index(m_name_d) if m_name_d in mcols else 0)
        with c7:
            mst_emp  = st.selectbox("Masterlist — Emp No", mcols, index=mcols.index(m_emp_d) if m_emp_d in mcols else 0)
        c8, c9 = st.columns(2)
        with c8:
            mst_join = st.selectbox("Masterlist — Joined Date", mcols, index=mcols.index(m_join_d) if m_join_d in mcols else 0)
        with c9:
            choices = (["(none)"] + mcols) if (m_recr_d not in mcols) else mcols
            idx = (choices.index(m_recr_d) if m_recr_d in choices else 0)
            mst_recr = st.selectbox("Masterlist — Recruiter", choices, index=idx)

        # Normalize masterlist fields
        mst[mst_name] = mst[mst_name].apply(_norm_name)
        if mst_emp in mst.columns:
            mst[mst_emp] = mst[mst_emp].apply(_norm_empid)
        mst[mst_join] = pd.to_datetime(mst[mst_join], errors="coerce")
        if mst_recr != "(none)":
            mst[mst_recr] = mst[mst_recr].apply(_norm_recruiter)
        else:
            mst_recr = None

        # Lookups with A-suffix alias + fallback by name
        join_by_emp, recr_by_emp = {}, {}
        if mst_emp in mst.columns:
            for _, rr in mst.iterrows():
                eid = rr[mst_emp]
                jd  = rr[mst_join]
                rc  = rr[mst_recr] if mst_recr else None
                if pd.notna(eid) and str(eid).strip() != "":
                    join_by_emp[eid] = jd
                    if mst_recr: recr_by_emp[eid] = rc
                    m = re.fullmatch(r"(\d+)[A-Z]", str(eid))
                    if m:
                        base = m.group(1)
                        if base not in join_by_emp:
                            join_by_emp[base] = jd
                            if mst_recr: recr_by_emp[base] = rc

        join_by_name = dict(zip(mst[mst_name], mst[mst_join]))
        recr_by_name = dict(zip(mst[mst_name], mst[mst_recr])) if mst_recr else {}

        def get_join(row):
            jd = join_by_emp.get(row["__EmpID"]) if mst_emp in mst.columns else None
            if jd is None or pd.isna(jd): jd = join_by_name.get(row["__Name"])
            return jd

        def get_recr(row):
            if not mst_recr: return "Unassigned"
            rc = recr_by_emp.get(row["__EmpID"]) if mst_emp in mst.columns else None
            if rc is None or (isinstance(rc, float) and pd.isna(rc)):
                rc = recr_by_name.get(row["__Name"])
            return _norm_recruiter(rc)

        att["JOIN_DATE"] = att.apply(get_join, axis=1)
        att["Recruiter"] = att.apply(get_recr, axis=1)

        # Optional: drop rows not present in Masterlist
        if exclude_not_in_master:
            if mst_emp in mst.columns:
                att = att[(att["JOIN_DATE"].notna()) | (att["__EmpID"].isin(join_by_emp.keys()))]
            else:
                att = att[(att["JOIN_DATE"].notna()) | (att["__Name"].isin(join_by_name.keys()))]

        # Eligibility window
        join_safe = pd.to_datetime(att["JOIN_DATE"], errors="coerce")
        att["ELIGIBLE_END"] = join_safe.apply(
            lambda d: (d + pd.DateOffset(months=3) - pd.Timedelta(days=1)) if pd.notna(d) else pd.NaT
        ).dt.normalize()

        # Claim can only start after 30 days from join (DISPLAY purpose)
        att["CLAIM_START"] = join_safe.apply(
            lambda d: (d + pd.Timedelta(days=30)) if pd.notna(d) else pd.NaT
        ).dt.normalize()


        # Eligibility filter
        eligible = att[
    (att["JOIN_DATE"].notna()) &
    (att["CLAIM_START"].notna()) &
    (att["__Date"] >= att["CLAIM_START"]) &
    (att["__Date"] <= att["ELIGIBLE_END"])
].copy()


        # Per-day aggregate within eligibility
        keys = ["__EmpID", "__Name", "JOIN_DATE", "CLAIM_START", "ELIGIBLE_END", "Recruiter", "__Date"]

        daily = (
            eligible.groupby(keys, as_index=False)
            .agg(
                Total_Hours=("__Hours", "sum"),
                Any_Leave=("__LeaveFlagRow", "max"),
                Leave_Types=("__LeaveText", lambda s: ", ".join(sorted({x for x in map(str, s) if x and x.lower()!="nan"}))),
            )
        )
        daily["Worked_Day"] = np.where(
            daily["Any_Leave"].astype(bool), 0, (daily["Total_Hours"] >= effective_threshold).astype(int)
        )

        if daily.empty:
            st.warning("No eligible rows after filtering. Check mappings/JOIN_DATEs.")
            st.stop()

        # Build month tables
        months = sorted(daily["__Date"].dropna().dt.to_period("M").unique())
        tables_by_month, summaries_by_month = {}, {}

        for m in months:
            tt = pd.Timestamp(m.start_time)
            days_in_month = calendar.monthrange(tt.year, tt.month)[1]
            mm = daily["__Date"].dt.to_period("M") == m

            base = (
                daily[mm]
                .groupby(["__EmpID","__Name","JOIN_DATE","CLAIM_START","ELIGIBLE_END","Recruiter"], as_index=False)
                .agg(Total_Hours=("Total_Hours","sum"), Claim_Days=("Worked_Day","sum"))
            )

            for d in range(1, days_in_month + 1):
                base[str(d)] = 0

            for _, r in daily[mm].iterrows():
                d = int(r["__Date"].day)
                mask = (base["__EmpID"]==r["__EmpID"]) & (base["__Name"]==r["__Name"]) & (base["Recruiter"]==r["Recruiter"])
                base.loc[mask, str(d)] = int(r["Worked_Day"])

            if counting_rule == "Floor(total hours ÷ threshold)":
                base["Claim_Days"] = np.floor(base["Total_Hours"] / hours_per_day).astype(int)

            base["TOTAL WORKING"] = base["Claim_Days"].astype(int)
            base = base.sort_values(["JOIN_DATE","__Name"], kind="stable").reset_index(drop=True)
            base = base.rename(columns={"__EmpID":"Emp No","__Name":"Name"})

            ordered = ["Emp No","Name","JOIN_DATE","CLAIM_START","ELIGIBLE_END","Recruiter"] + \
                      [str(d) for d in range(1, days_in_month+1)] + ["TOTAL WORKING","Total_Hours"]
            if "Total_Hours" not in base.columns:
                base["Total_Hours"] = 0.0
            tables_by_month[str(m)] = base[ordered]

            summary = (base.groupby("Recruiter", as_index=False)["TOTAL WORKING"]
                       .sum().rename(columns={"TOTAL WORKING":"Days"}))
            summary["Rate (RM)"] = day_rate
            summary["Amount (RM)"] = summary["Days"] * day_rate
            grand = pd.DataFrame({"Recruiter":["TOTAL"],
                                  "Days":[summary["Days"].sum()],
                                  "Rate (RM)":[day_rate],
                                  "Amount (RM)":[summary["Amount (RM)"].sum()]})
            summaries_by_month[str(m)] = (summary, grand)

        st.success(f"Built claim tables for {len(tables_by_month)} month(s). "
                   f"(Effective threshold = {effective_threshold:.2f} h with {grace_minutes}m grace)")

        for mon in sorted(tables_by_month.keys()):
            st.subheader(f"📅 {mon}")
            st.dataframe(tables_by_month[mon], use_container_width=True)
            st.caption("Per-Recruiter Summary")
            summary, grand = summaries_by_month[mon]
            st.dataframe(summary, use_container_width=True)
            st.write("**Grand Total**")
            st.dataframe(grand, use_container_width=True)

        # -------- Employee Debugger (shows In/Out + daily hours) --------
        st.markdown("---")
        st.header("🔎 Employee Debugger — time in/out + daily hours")

        def _fmt_h(h):
            if pd.isna(h): return ""
            h = float(h) % 24.0
            hh = int(h); mm = int(round((h - hh) * 60))
            if mm == 60: hh = (hh + 1) % 24; mm = 0
            return f"{hh:02d}:{mm:02d}"

        dbg_raw = att[(att["JOIN_DATE"].notna()) &
                      (att["__Date"] >= att["JOIN_DATE"]) &
                      (att["__Date"] <= att["ELIGIBLE_END"])].copy()

        if dbg_raw.empty:
            st.info("No rows in eligibility window.")
        else:
            dbg_raw["Month"] = dbg_raw["__Date"].dt.to_period("M").astype(str)
            months_avail = sorted(dbg_raw["Month"].unique().tolist())
            names_avail  = sorted(dbg_raw["__Name"].unique().tolist())
            cA, cB = st.columns(2)
            with cA: dbg_month = st.selectbox("Month", months_avail, index=len(months_avail)-1)
            with cB: dbg_name  = st.selectbox("Employee", names_avail, index=0)

            sub = dbg_raw[(dbg_raw["Month"]==dbg_month) & (dbg_raw["__Name"]==dbg_name)].copy()
            if sub.empty:
                st.info("No rows for that selection.")
            else:
                per_day = (sub.groupby(["__Date"], as_index=False)
                           .agg(First_In_h=("__InH","min"),
                                Last_Out_h=("__OutH","max"),
                                Daily_Hours=("__Hours","sum"),
                                Any_Leave=("__LeaveFlagRow","max"),
                                Join=("JOIN_DATE","first"),
                                Eligible_End=("ELIGIBLE_END","first"),
                                Recruiter=("Recruiter","first")))
                per_day["In"]  = per_day["First_In_h"].apply(_fmt_h)
                per_day["Out"] = per_day["Last_Out_h"].apply(_fmt_h)
                per_day["Counted(1/0)"] = np.where(
                    per_day["Any_Leave"].astype(bool), 0,
                    (per_day["Daily_Hours"] >= effective_threshold).astype(int)
                )
                st.write(f"Threshold used: **{effective_threshold:.2f} h** (base {hours_per_day} − {grace_minutes}m grace)")
                st.dataframe(per_day[["__Date","In","Out","Daily_Hours","Any_Leave","Counted(1/0)","Join","Eligible_End","Recruiter"]]
                             .sort_values("__Date"), use_container_width=True)

        # ---------- Export ----------
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            for mon in sorted(tables_by_month.keys()):
                sheet = f"{mon}"[:31]
                df_out = tables_by_month[mon]
                df_out.to_excel(writer, sheet_name=sheet, index=False)

                workbook = writer.book
                ws = writer.sheets[sheet]
                bold = workbook.add_format({"bold": True})
                ws.freeze_panes(1, 5)  # freeze header and ID/Name/Join/Eligible/Recruiter

                start = len(df_out) + 2
                ws.write(start, 0, "Per-Recruiter Summary", bold); start += 1
                summary, grand = summaries_by_month[mon]
                summary.to_excel(writer, sheet_name=sheet, startrow=start, index=False)
                start = start + len(summary) + 2
                ws.write(start, 0, "Grand Total", bold); start += 1
                grand.to_excel(writer, sheet_name=sheet, startrow=start, index=False)

        st.download_button(
            "⬇️ Download Excel (one sheet per month: table + summary + grand total)",
            data=buffer.getvalue(),
            file_name=f"claim_tables_{dt.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error while building claim table: {e}")
        st.exception(e)
else:
    st.info("Upload both files to continue.")
