from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse
import pandas as pd
import numpy as np
import io

app = FastAPI(title="Payroll Automation API")

# Hardcoded Rates Mapping
# Default: Rate = 40.0, Loading = 18.0 (Ratio-based)
# Exceptions: Christian & Ayman maintain the exact same ratio
EMPLOYEE_RATES = {
    "DEFAULT": {"rate": 40.0, "loading": 18.0},
    "christian": {"rate": 40.0, "loading": 18.0},
    "ayman": {"rate": 40.0, "loading": 18.0},
}

def get_employee_rate(fname, lname):
    full_name = f"{str(fname).strip()} {str(lname).strip()}".strip().lower()
    
    # Check for name exceptions using Full Name search
    for name_key, rates in EMPLOYEE_RATES.items():
        if name_key != "DEFAULT" and name_key in full_name:
            return rates["rate"], rates["loading"]
            
    # Fallback to default $40 rate and $18 loading
    return EMPLOYEE_RATES["DEFAULT"]["rate"], EMPLOYEE_RATES["DEFAULT"]["loading"]

def calculate_alberta_ot(df):
    df['local_date'] = pd.to_datetime(df['local_date'])
    df['hours'] = df['hours'].astype(float)
    
    # Generate Unique Full Name Identifier (UID)
    df['full_name'] = df['fname'].astype(str).str.strip() + " " + df['lname'].astype(str).str.strip()
    
    # Sort logically
    df = df.sort_values(by=['full_name', 'local_date', 'local_start_time']).reset_index(drop=True)
    
    # Group by Full Name UID and Date for Daily OT
    daily = df.groupby(['full_name', 'local_date'])['hours'].sum().reset_index()
    daily['daily_reg'] = daily['hours'].apply(lambda h: min(h, 8.0))
    daily['daily_ot'] = daily['hours'].apply(lambda h: max(0.0, h - 8.0))
    
    # Add ISO Year-Week identifier (Monday start)
    daily['year_week'] = daily['local_date'].dt.strftime('%G-%V')
    
    # Calculate weekly cumulative regular hours for Alberta > 44h rule
    daily['weekly_cum_reg'] = daily.groupby(['full_name', 'year_week'])['daily_reg'].cumsum()
    
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
    
    # Merge back to original shift records
    df = df.merge(
        daily[['full_name', 'local_date', 'hours', 'final_reg', 'final_ot']],
        on=['full_name', 'local_date'],
        suffixes=('', '_daily_total')
    )
    
    # Pro-rate hours for multiple daily shifts
    df['shift_ratio'] = np.where(df['hours_daily_total'] > 0, df['hours'] / df['hours_daily_total'], 0)
    df['Reg'] = (df['final_reg'] * df['shift_ratio']).round(2)
    df['OT'] = (df['final_ot'] * df['shift_ratio']).round(2)
    df['Adj Total'] = (df['Reg'] + (df['OT'] * 1.5)).round(2)
    
    # Hours & Minutes formatting
    df['Hours'] = df['Adj Total'].apply(lambda x: int(np.floor(x)))
    df['Minutes'] = df['Adj Total'].apply(lambda x: int(round((x % 1) * 60)))
    
    # Apply Hardcoded Rates using Full Name UID
    rates = df.apply(lambda r: get_employee_rate(r['fname'], r['lname']), axis=1)
    df['Rate'] = [r[0] for r in rates]
    df['Loading'] = [r[1] for r in rates]
    df['Adj Rate'] = df['Rate'] + df['Loading']
    df['Dollars'] = (df['Adj Total'] * df['Adj Rate']).round(2)
    
    df['Job'] = ""
    
    output_cols = [
        'fname', 'lname', 'local_date', 'local_day', 'local_start_time', 'local_end_time',
        'hours', 'Reg', 'OT', 'Adj Total', 'Hours', 'Minutes',
        'Rate', 'Loading', 'Adj Rate', 'Dollars', 'Job',
        'jobcode_1', 'class', 'service item', 'notes', 'approved_status'
    ]
    
    for col in output_cols:
        if col not in df.columns:
            df[col] = ""
            
    return df[output_cols]

@app.get("/")
def health_check():
    return {"status": "ok", "message": "Payroll API is running successfully."}

@app.post("/process-timesheet")
async def process_timesheet(file: UploadFile = File(...)):
    contents = await file.read()
    df_raw = pd.read_csv(io.BytesIO(contents))
    
    df_processed = calculate_alberta_ot(df_raw)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_processed.to_excel(writer, index=False, sheet_name='Processed Payroll')
    output.seek(0)
    
    filename = file.filename.replace('.csv', '_Processed.xlsx')
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
