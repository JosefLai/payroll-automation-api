from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import StreamingResponse
import pandas as pd
import numpy as np
import io
import json

app = FastAPI(title="Payroll Automation API")

def parse_rates_json(rates_json_str: str):
    """
    Parses rates provided dynamically from Make.com (sourced from Google Sheet).
    Expected JSON structure:
    {
        "Cooper Kelly": {"rate": 40.0, "loading": 18.0},
        "Ayman X": {"rate": 40.0, "loading": 18.0},
        "DEFAULT": {"rate": 40.0, "loading": 18.0}
    }
    """
    default_rates = {"rate": 40.0, "loading": 18.0}
    if not rates_json_str:
        return {"DEFAULT": default_rates}
    
    try:
        rates_dict = json.loads(rates_json_str)
        if "DEFAULT" not in rates_dict:
            rates_dict["DEFAULT"] = default_rates
        return rates_dict
    except Exception:
        return {"DEFAULT": default_rates}

def get_employee_rate(fname, lname, rates_dict):
    # Combine First Name + Last Name to form a Unique Identifier (UID)
    full_name = f"{str(fname).strip()} {str(lname).strip()}".strip().lower()
    
    # Check exact full name match in lower case
    for name_key, rates in rates_dict.items():
        if name_key.lower() == full_name:
            return float(rates.get("rate", 40.0)), float(rates.get("loading", 18.0))
            
    # Fallback to DEFAULT rates
    default = rates_dict.get("DEFAULT", {"rate": 40.0, "loading": 18.0})
    return float(default.get("rate", 40.0)), float(default.get("loading", 18.0))

def calculate_alberta_ot(df, rates_dict):
    df['local_date'] = pd.to_datetime(df['local_date'])
    df['hours'] = df['hours'].astype(float)
    
    # Sort data by unique full name + start time
    df['full_name'] = df['fname'].astype(str).str.strip() + " " + df['lname'].astype(str).str.strip()
    df = df.sort_values(by=['full_name', 'local_date', 'local_start_time']).reset_index(drop=True)
    
    # Group by Unique Full Name and Date for Daily OT calculation
    daily = df.groupby(['full_name', 'local_date'])['hours'].sum().reset_index()
    daily['daily_reg'] = daily['hours'].apply(lambda h: min(h, 8.0))
    daily['daily_ot'] = daily['hours'].apply(lambda h: max(0.0, h - 8.0))
    
    # Add ISO Year-Week identifier (Monday as start of week)
    daily['year_week'] = daily['local_date'].dt.strftime('%G-%V')
    
    # Calculate weekly cumulative regular hours to enforce Alberta > 44h rule
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
    
    # Merge back to original shift entries
    df = df.merge(
        daily[['full_name', 'local_date', 'hours', 'final_reg', 'final_ot']],
        on=['full_name', 'local_date'],
        suffixes=('', '_daily_total')
    )
    
    # Pro-rate shift hours
    df['shift_ratio'] = np.where(df['hours_daily_total'] > 0, df['hours'] / df['hours_daily_total'], 0)
    df['Reg'] = (df['final_reg'] * df['shift_ratio']).round(2)
    df['OT'] = (df['final_ot'] * df['shift_ratio']).round(2)
    df['Adj Total'] = (df['Reg'] + (df['OT'] * 1.5)).round(2)
    
    # Hours & Minutes formatting
    df['Hours'] = df['Adj Total'].apply(lambda x: int(np.floor(x)))
    df['Minutes'] = df['Adj Total'].apply(lambda x: int(round((x % 1) * 60)))
    
    # Apply Loaded Rates using Full Name UID + Rates Dict
    rates = df.apply(lambda r: get_employee_rate(r['fname'], r['lname'], rates_dict), axis=1)
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
    return {"status": "ok", "message": "Payroll API with Dynamic Rates is running."}

@app.post("/process-timesheet")
async def process_timesheet(
    file: UploadFile = File(...),
    rates_json: str = Form(None)  # Optional JSON string sent from Make.com
):
    contents = await file.read()
    df_raw = pd.read_csv(io.BytesIO(contents))
    
    # Parse rates passed dynamically from Google Sheet
    rates_dict = parse_rates_json(rates_json)
    
    # Process Payroll
    df_processed = calculate_alberta_ot(df_raw, rates_dict)
    
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
