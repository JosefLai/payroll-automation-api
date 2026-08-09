from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse
import pandas as pd
import numpy as np
import io
from openpyxl.styles import PatternFill

app = FastAPI(title="Payroll Automation API")

# ----------------------------------------------------
# Rule 4: Employee-specific loaded rates
# ----------------------------------------------------
EMPLOYEE_RATES = {
    "DEFAULT": {"rate": 40.0, "loading": 18.0},          # Adj Rate = $58.00
    "ayman": {"rate": 50.0, "loading": 22.50},            # Adj Rate = $72.50
    "christian": {"rate": 27.0, "loading": 12.15},        # Adj Rate = $39.15
}

# ----------------------------------------------------
# Rule 1 & 2: Pruning rules (confirmed with Dan, Aug 4-5 email thread)
# ----------------------------------------------------
EXCLUDED_JOBCODES = ["general admin", "vacation", "sick", "sick- unpaid", "management"]
# ASSUMPTION (open item - see email): these 3 employees are hard-coded exceptions because
# they log admin time under jobcodes shared with billable staff (e.g. "Quotes", "My Home
# Handyman"). Confirmed with Dan on Aug 4. Preferred long-term fix is asking them to stop
# using those jobcodes rather than growing this hard-coded list.
EXCLUDED_EMPLOYEES = ["mark biggins", "william lebherz", "jonathan kruger"]

# Penny-accurate minute table per Dan's Aug 5 request (60 values, 4 significant figures,
# rounded to the minute - not the second, per Dan's confirmation).
MINUTE_FRACTION = {m: round(m / 60, 4) for m in range(60)}

LUNCH_PAID_THRESHOLD_HR = 0.5  # Dan: first 30 min of a lunch break is paid, anything beyond is a flag

# OT-EXEMPT job codes (corrected per Dan's review, Aug 9): these entries are identified
# directly by the jobcode_1 column QBO already assigns - NOT by detecting time gaps
# between shifts. They must never be assigned OT hours/dollars themselves; any OT for
# the day is carried entirely by the other (billable) shifts on that day.
#
# IMPORTANT - narrowed during testing (Aug 9): Dan initially asked to exempt "lunch break
# or any other downtime activities." I tested treating jobcode_1 == "Downtime" the same
# as "Lunch break", but checked it against Dan's actual June working paper first and it
# does NOT hold - his own file assigns OT to "Downtime"-coded rows in 20 of 164 cases
# (11.79 OT hours total), while "Lunch break" rows carry OT in only 1 of 78 cases. So only
# "Lunch break" is treated as OT-exempt below. STILL OPEN: what Dan meant by "other
# downtime activities" if not the literal "Downtime" jobcode - flagged for his reply.
OT_EXEMPT_JOBCODES = ["lunch break"]


def get_employee_rate(fname, lname):
    full_name = f"{str(fname).strip()} {str(lname).strip()}".strip().lower()
    for name_key, rates in EMPLOYEE_RATES.items():
        if name_key != "DEFAULT" and name_key in full_name:
            return rates["rate"], rates["loading"]
    return EMPLOYEE_RATES["DEFAULT"]["rate"], EMPLOYEE_RATES["DEFAULT"]["loading"]


def flag_lunch_breaks(df):
    """Flag lunch break entries identified by jobcode_1 == 'Lunch break' (the literal
    QBO code Dan enters), NOT by detecting a time gap between shifts. Only flags the
    ones over the 30-minute paid threshold, since that's the case Dan needs to review.
    This is a REVIEW flag only - the automation does not change any hours based on it."""
    is_lunch = df["jobcode_1"].str.lower() == "lunch break"
    over_threshold = is_lunch & (df["hours"] > LUNCH_PAID_THRESHOLD_HR)
    df["Lunch_Flag"] = ""
    df.loc[over_threshold, "Lunch_Flag"] = "Lunch break >30 min - review"
    return df


def calculate_alberta_ot(df):
    # ----------------------------------------------------
    # Data cleaning & column pre-processing
    # ----------------------------------------------------
    df["hours"] = pd.to_numeric(df["hours"], errors="coerce").fillna(0)
    df["full_name"] = df["fname"].astype(str).str.strip() + " " + df["lname"].astype(str).str.strip()
    df["full_name_clean"] = df["full_name"].str.lower().str.strip()

    for col in ["jobcode_1", "jobcode_2", "class", "service item"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).str.strip()

    df["local_date"] = pd.to_datetime(df["local_date"])
    df["local_start_time"] = pd.to_datetime(df["local_start_time"])
    df["local_end_time"] = pd.to_datetime(df["local_end_time"])

    # ----------------------------------------------------
    # Rule 3: drop hours = 0 (QBO raw export column, confirmed with Dan Jul 24)
    # ----------------------------------------------------
    df = df[df["hours"] > 0].copy()

    # ----------------------------------------------------
    # Rule 1 & 2: pruning (jobcodes, admin class/service item, exception employees)
    # ----------------------------------------------------
    is_excluded_user = df["full_name_clean"].isin(EXCLUDED_EMPLOYEES)
    is_excluded_jobcode = df["jobcode_1"].str.lower().isin(EXCLUDED_JOBCODES)
    is_admin_class = (
        df["class"].str.lower().str.contains("admin", na=False)
        | df["service item"].str.lower().str.contains("admin", na=False)
    )
    df = df[~(is_excluded_user | is_excluded_jobcode | is_admin_class)].copy()

    if df.empty:
        return df

    # ----------------------------------------------------
    # Lunch break >30 min flag (identified by jobcode_1, not by time gaps - Aug 9 correction)
    # ----------------------------------------------------
    df = flag_lunch_breaks(df)

    # Mark OT-exempt rows (Lunch break / Downtime, per Aug 9 correction). These entries
    # are excluded from the OT allocation below - they always get paid straight time for
    # their actual hours, and never absorb any of the day's overtime.
    df["is_ot_exempt"] = df["jobcode_1"].str.lower().isin(OT_EXEMPT_JOBCODES)

    # ----------------------------------------------------
    # Alberta 8/44 OT split
    # ----------------------------------------------------
    df = df.sort_values(by=["full_name", "local_date", "local_start_time"]).reset_index(drop=True)

    # 1. Daily OT (>8h). Uses ALL hours worked that day (including lunch/downtime) to
    #    determine whether the 8-hour daily threshold was crossed - this is unchanged.
    daily = df.groupby(["full_name", "local_date"])["hours"].sum().reset_index()
    daily["daily_reg"] = daily["hours"].apply(lambda h: min(h, 8.0))
    daily["daily_ot"] = daily["hours"].apply(lambda h: max(0.0, h - 8.0))

    # 2. Weekly OT (>44h)
    daily["year_week"] = daily["local_date"].dt.strftime("%G-%V")
    daily["weekly_cum_reg"] = daily.groupby(["full_name", "year_week"])["daily_reg"].cumsum()

    def apply_weekly_cap(row):
        cum = row["weekly_cum_reg"]
        reg = row["daily_reg"]
        ot = row["daily_ot"]
        if cum > 44.0:
            excess = cum - 44.0
            actual_reg = max(0.0, reg - excess)
            actual_ot = ot + (reg - actual_reg)
            return pd.Series([actual_reg, actual_ot])
        else:
            return pd.Series([reg, ot])

    daily[["final_reg", "final_ot"]] = daily.apply(apply_weekly_cap, axis=1)

    # 3. Set aside the OT-exempt hours as straight Reg time, then split the REMAINING
    #    Reg/OT pool only across the non-exempt (billable) shifts of the day, in
    #    proportion to each one's share of the non-exempt hours.
    #    ASSUMPTION (OPEN ITEM - still not confirmed with Dan): when a day has more than
    #    one billable shift under different job codes, we split the remaining OT
    #    proportionally by hours. Testing found at least one case (Moein Moradi, Jun 24)
    #    where Dan's actual entry put ALL of a day's OT on a single shift instead of
    #    splitting it - so this proportional assumption is still not confirmed to match
    #    Dan's real method. The DAY TOTAL will be correct either way; the split between
    #    individual billable shifts on a multi-shift day may not be. Flagged below.
    exempt_hours = df[df["is_ot_exempt"]].groupby(["full_name", "local_date"])["hours"].sum()
    exempt_hours.name = "exempt_hours"
    daily = daily.merge(exempt_hours, on=["full_name", "local_date"], how="left")
    daily["exempt_hours"] = daily["exempt_hours"].fillna(0.0)

    # Guard against the rare edge case where exempt hours alone exceed the day's Reg cap
    # (e.g. an unusually long downtime block) - keeps the Reg/OT pools from going negative.
    daily["reg_pool"] = (daily["final_reg"] - daily["exempt_hours"]).clip(lower=0.0)
    shortfall = (daily["exempt_hours"] - daily["final_reg"]).clip(lower=0.0)
    daily["ot_pool"] = daily["final_ot"] + shortfall
    daily["eligible_hours"] = daily["hours"] - daily["exempt_hours"]

    df = df.merge(
        daily[["full_name", "local_date", "reg_pool", "ot_pool", "eligible_hours"]],
        on=["full_name", "local_date"],
    )

    df["shift_ratio"] = np.where(
        (~df["is_ot_exempt"]) & (df["eligible_hours"] > 0),
        df["hours"] / df["eligible_hours"],
        0,
    )

    # Count BILLABLE (non-exempt) shifts per employee-day - this is what actually drives
    # the proportional-split assumption above. A day with several exempt entries but only
    # one billable shift has no real ambiguity, since 100% of the remaining pool goes to
    # that one shift regardless of method.
    eligible_shift_counts = df[~df["is_ot_exempt"]].groupby(["full_name", "local_date"]).size().rename("n_eligible_shifts")
    df = df.merge(eligible_shift_counts, on=["full_name", "local_date"], how="left")
    df["n_eligible_shifts"] = df["n_eligible_shifts"].fillna(0).astype(int)
    df["Multi_Shift_Day"] = np.where(df["n_eligible_shifts"] > 1, "Review - OT split assumption applied", "")

    # ----------------------------------------------------
    # KEY FIX (Aug 9, still applies): keep full floating-point precision from Reg/OT all
    # the way through to the Hours/Minutes split. Rounding Reg/OT to 2 decimals BEFORE
    # deriving Hours/Minutes double-rounds and can land on the wrong minute (e.g. 11.05
    # exact -> minute 3, but 11.04 rounded first -> minute 2). Reg/OT/Adj Total below are
    # rounded only for on-screen display; Hours/Minutes/Dollars use the un-rounded value.
    # ----------------------------------------------------
    reg_raw = np.where(df["is_ot_exempt"], df["hours"], df["reg_pool"] * df["shift_ratio"])
    ot_raw = np.where(df["is_ot_exempt"], 0.0, df["ot_pool"] * df["shift_ratio"])
    reg_raw = pd.Series(reg_raw, index=df.index)
    ot_raw = pd.Series(ot_raw, index=df.index)
    adj_total_raw = reg_raw + ot_raw * 1.5

    df["Reg"] = reg_raw.round(2)
    df["OT"] = ot_raw.round(2)
    df["Adj Total"] = adj_total_raw.round(2)

    # 4. Jobber hours/minutes split (pink / purple columns)
    df["Hours"] = adj_total_raw.apply(lambda x: int(np.floor(x)))
    df["Minutes"] = adj_total_raw.apply(lambda x: int(round((x % 1) * 60)))
    rolled_over = df["Minutes"] == 60
    df.loc[rolled_over, "Hours"] += 1
    df.loc[rolled_over, "Minutes"] = 0

    # ----------------------------------------------------
    # Rate + penny-accurate Dollars (per Dan's Aug 5 request: 60-value, 4-sig-fig table)
    # ----------------------------------------------------
    rates = df.apply(lambda r: get_employee_rate(r["fname"], r["lname"]), axis=1)
    df["Rate"] = [r[0] for r in rates]
    df["Loading"] = [r[1] for r in rates]
    df["Adj Rate"] = df["Rate"] + df["Loading"]
    df["Dollars"] = ((df["Hours"] + df["Minutes"].map(MINUTE_FRACTION)) * df["Adj Rate"]).round(2)

    # Job / Cost Code - Phase 2 (not yet built). Left blank intentionally.
    df["Job"] = ""

    output_cols = [
        "fname", "lname", "local_date", "local_day", "local_start_time", "local_end_time",
        "hours", "Reg", "OT", "Adj Total", "Hours", "Minutes",
        "Rate", "Loading", "Adj Rate", "Dollars", "Job",
        "jobcode_1", "jobcode_2", "class", "service item", "notes", "approved_status",
        "Lunch_Flag", "Multi_Shift_Day",
    ]
    for col in output_cols:
        if col not in df.columns:
            df[col] = ""
    df = df[output_cols]

    # Multi-level sort: 1. jobcode_1  2. local_date (old->new)  3. fname (A-Z)
    df = df.sort_values(by=["jobcode_1", "local_date", "fname"], ascending=[True, True, True]).reset_index(drop=True)

    return df


@app.get("/")
def health_check():
    return {"status": "ok", "message": "Payroll API is running successfully."}


@app.post("/process-timesheet")
async def process_timesheet(file: UploadFile = File(...)):
    contents = await file.read()
    df_raw = pd.read_csv(io.BytesIO(contents))

    df_processed = calculate_alberta_ot(df_raw)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_processed.to_excel(writer, index=False, sheet_name="Processed Payroll")
        ws = writer.sheets["Processed Payroll"]

        hours_col_idx = df_processed.columns.get_loc("Hours") + 1
        minutes_col_idx = df_processed.columns.get_loc("Minutes") + 1
        lunch_col_idx = df_processed.columns.get_loc("Lunch_Flag") + 1
        multi_col_idx = df_processed.columns.get_loc("Multi_Shift_Day") + 1

        fill_hours = PatternFill(start_color="EAD1DC", end_color="EAD1DC", fill_type="solid")
        fill_minutes = PatternFill(start_color="D9D2E9", end_color="D9D2E9", fill_type="solid")
        fill_flag = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        for row in range(1, ws.max_row + 1):
            ws.cell(row=row, column=hours_col_idx).fill = fill_hours
            ws.cell(row=row, column=minutes_col_idx).fill = fill_minutes
            if row > 1:
                if ws.cell(row=row, column=lunch_col_idx).value:
                    ws.cell(row=row, column=lunch_col_idx).fill = fill_flag
                if ws.cell(row=row, column=multi_col_idx).value:
                    ws.cell(row=row, column=multi_col_idx).fill = fill_flag

    output.seek(0)

    filename = file.filename.replace(".csv", "_Processed.xlsx")
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )
