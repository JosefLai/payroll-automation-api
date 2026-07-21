from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse
import pandas as pd
import numpy as np
import io

app = FastAPI(title="Payroll Automation API")

# Default Rate Mapping (Can be updated dynamically)
EMPLOYEE_RATES = {
    # Default rate for most employees
    "DEFAULT": {"rate": 40.0, "loading": 18.0},
    # Specific exceptions can be added here
    # "Christian": {"rate": 40.0, "loading": 18.0},
}

def get_employee_rate(fname, lname):
    full_name = f"{fname} {lname}".strip()
    for name_key, rates in EMPLOYEE_RATES.items():
        if name_key != "DEFAULT" and name_key.lower() in full_name.lower():
            return rates["rate"], rates["loading"]
    return EMPLOYEE_RATES["DEFAULT"]["rate"], EMPLOYEE_RATES["DEFAULT"]["loading"]

def calculate_alberta_ot(df):
    """
    Implements Alberta 8/44 Overtime Rule:
    Daily OT = hours over 8 per day.
    Weekly OT = total regular hours in a week (Mon-Sun) over 44.
    Take whichever produces higher OT.
    """
    df['local_date'] = pd.to_datetime(df['local_date'])
    df['hours'] = df['hours'].astype(float)
    
    # Sort data logically
    df = df.sort_values(by=['fname', 'lname', 'local_date', 'local_start_time']).reset_index(drop=True)
    
    # Group by Employee and Date for Daily OT calculation
    daily = df.groupby(['fname', 'lname', 'local_date'])['hours'].sum().reset_index()
    daily['daily_reg'] = daily['hours'].apply(lambda h: min(h, 8.0))
    daily['daily_ot'] = daily['hours'].apply(lambda h: max(0.0, h - 8.0))
    
    # Add Year-Week identifier (Monday as start of week)
    daily['year_week'] = daily['local_date'].dt.strftime('%G-%V')
    
    # Calculate weekly cumulative regular hours to check > 44h rule
    daily['weekly_cum_reg'] = daily.groupby(['fname', 'lname', 'year_week'])['daily_reg'].cumsum()
    
    def apply_weekly_cap(row):
        cum = row['weekly_cum_reg']
        reg = row['daily_reg']
        ot = row['daily_ot']
        
        if cum > 44.0:
            excess = cum - 44.0
            actual_reg = max(0.0, reg - excess)
            actual_ot = ot + (reg - actual_reg)
            return pd.Series([actual_reg, actual_ot])
        else:
            return pd.Series([reg, ot])
            
    daily[['final_reg', 'final_ot']] = daily.apply(apply_weekly_cap, axis=1)
    
    # Merge daily calculation back to individual shift entries
    df = df.merge(
        daily[['fname', 'lname', 'local_date', 'hours', 'final_reg', 'final_ot']],
        on=['fname', 'lname', 'local_date'],
        suffixes=('', '_daily_total')
    )
    
    # Pro-rate hours for multiple shifts in the same day
    df['shift_ratio'] = np.where(df['hours_daily_total'] > 0, df['hours'] / df['hours_daily_total'], 0)
    df['Reg'] = (df['final_reg'] * df['shift_ratio']).round(2)
    df['OT'] = (df['final_ot'] * df['shift_ratio']).round(2)
    df['Adj Total'] = (df['Reg'] + (df['OT'] * 1.5)).round(2)
    
    # Calculate Hours and Minutes display
    df['Hours'] = df['Adj Total'].apply(lambda x: int(np.floor(x)))
    df['Minutes'] = df['Adj Total'].apply(lambda x: int(round((x % 1) * 60)))
    
    # Apply Loaded Rates
    rates = df.apply(lambda r: get_employee_rate(r['fname'], r['lname']), axis=1)
    df['Rate'] = [r[0] for r in rates]
    df['Loading'] = [r[1] for r in rates]
    df['Adj Rate'] = df['Rate'] + df['Loading']
    df['Dollars'] = (df['Adj Total'] * df['Adj Rate']).round(2)
    
    # Job column placeholder for Step 2 Jobber matching
    df['Job'] = ""
    
    # Reorder columns to match Dan's exact processed sheet format
    output_cols = [
        'fname', 'lname', 'local_date', 'local_day', 'local_start_time', 'local_end_time',
        'hours', 'Reg', 'OT', 'Adj Total', 'Hours', 'Minutes',
        'Rate', 'Loading', 'Adj Rate', 'Dollars', 'Job',
        'jobcode_1', 'class', 'service item', 'notes', 'approved_status'
    ]
    
    # Retain existing columns and fill missing if any
    for col in output_cols:
        if col not in df.columns:
            df[col] = ""
            
    return df[output_cols]

@app.get("/")
def health_check():
    return {"status": "ok", "message": "Payroll API is running."}

@app.post("/process-timesheet")
async def process_timesheet(file: UploadFile = File(...)):
    # Read CSV uploaded by Make.com
    contents = await file.read()
    df_raw = pd.read_csv(io.BytesIO(contents))
    
    # Process Payroll
    df_processed = calculate_alberta_ot(df_raw)
    
    # Export to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_processed.to_excel(writer, index=False, sheet_name='Processed Payroll')
    output.seek(0)
    
    # Return Excel File to Make.com
    filename = file.filename.replace('.csv', '_Processed.xlsx')
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
