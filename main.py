from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse
import pandas as pd
import numpy as np
import io
from openpyxl.styles import PatternFill

app = FastAPI(title="Payroll Automation API")

# ----------------------------------------------------
# Rule 4: 特殊員工費率與預設費率對照
# ----------------------------------------------------
EMPLOYEE_RATES = {
    "DEFAULT": {"rate": 40.0, "loading": 18.0},          # Adj Rate = $58.00
    "ayman": {"rate": 50.0, "loading": 22.50},            # Adj Rate = $72.50
    "christian": {"rate": 27.0, "loading": 12.15},        # Adj Rate = $39.15
}

# ----------------------------------------------------
# Rule 1 & 2: 過濾條件 (Pruning Rules)
# ----------------------------------------------------
EXCLUDED_JOBCODES = ["general admin", "vacation", "sick", "sick- unpaid", "management"]
EXCLUDED_EMPLOYEES = ["mark biggins", "william lebherz", "jonathan kruger"]

def get_employee_rate(fname, lname):
    full_name = f"{str(fname).strip()} {str(lname).strip()}".strip().lower()
    for name_key, rates in EMPLOYEE_RATES.items():
        if name_key != "DEFAULT" and name_key in full_name:
            return rates["rate"], rates["loading"]
    return EMPLOYEE_RATES["DEFAULT"]["rate"], EMPLOYEE_RATES["DEFAULT"]["loading"]

def calculate_alberta_ot(df):
    # ----------------------------------------------------
    # Data Cleaning & Column Pre-processing
    # ----------------------------------------------------
    df['hours'] = pd.to_numeric(df['hours'], errors='coerce').fillna(0)
    df['full_name'] = df['fname'].astype(str).str.strip() + " " + df['lname'].astype(str).str.strip()
    df['full_name_clean'] = df['full_name'].str.lower().str.strip()
    
    # 包含 jobcode_2 欄位預處理
    for col in ['jobcode_1', 'jobcode_2', 'class', 'service item']:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).str.strip()

    # ----------------------------------------------------
    # Rule 3: 剔除 hours = 0 的紀錄
    # ----------------------------------------------------
    df = df[df['hours'] > 0].copy()

    # ----------------------------------------------------
    # Rule 1 & 2: 執行 Pruning (Jobcodes, Admin Class & Exception Employees)
    # ----------------------------------------------------
    is_excluded_user = df['full_name_clean'].isin(EXCLUDED_EMPLOYEES)
    is_excluded_jobcode = df['jobcode_1'].str.lower().isin(EXCLUDED_JOBCODES)
    is_admin_class = (
        df['class'].str.lower().str.contains('admin', na=False) |
        df['service item'].str.lower().str.contains('admin', na=False)
    )
    
    df = df[~(is_excluded_user | is_excluded_jobcode | is_admin_class)].copy()

    if df.empty:
        return df

    # ----------------------------------------------------
    # Alberta 8/44 加班費拆分算法
    # ----------------------------------------------------
    df['local_date'] = pd.to_datetime(df['local_date'])
    df = df.sort_values(by=['full_name', 'local_date', 'local_start_time']).reset_index(drop=True)
    
    # 1. 單日 Daily OT (>8h)
    daily = df.groupby(['full_name', 'local_date'])['hours'].sum().reset_index()
    daily['daily_reg'] = daily['hours'].apply(lambda h: min(h, 8.0))
    daily['daily_ot'] = daily['hours'].apply(lambda h: max(0.0, h - 8.0))
    
    # 2. 每週 Weekly OT (>44h)
    daily['year_week'] = daily['local_date'].dt.strftime('%G-%V')
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
    
    # 3. 按多班次比例 (Shift Ratio) 拆分
    df = df.merge(
        daily[['full_name', 'local_date', 'hours', 'final_reg', 'final_ot']],
        on=['full_name', 'local_date'],
        suffixes=('', '_daily_total')
    )
    
    df['shift_ratio'] = np.where(df['hours_daily_total'] > 0, df['hours'] / df['hours_daily_total'], 0)
    df['Reg'] = (df['final_reg'] * df['shift_ratio']).round(2)
    df['OT'] = (df['final_ot'] * df['shift_ratio']).round(2)
    df['Adj Total'] = (df['Reg'] + (df['OT'] * 1.5)).round(2)
    
    # 4. Jobber 時分拆分
    df['Hours'] = df['Adj Total'].apply(lambda x: int(np.floor(x)))
    df['Minutes'] = df['Adj Total'].apply(lambda x: int(round((x % 1) * 60)))
    
    # ----------------------------------------------------
    # 套用費率與 Jobber 高精度 Dollars 計算
    # ----------------------------------------------------
    rates = df.apply(lambda r: get_employee_rate(r['fname'], r['lname']), axis=1)
    df['Rate'] = [r[0] for r in rates]
    df['Loading'] = [r[1] for r in rates]
    df['Adj Rate'] = df['Rate'] + df['Loading']
    df['Dollars'] = ((df['Hours'] + (df['Minutes'] / 60.0)) * df['Adj Rate']).round(2)
    
    df['Job'] = ""
    
    # 插入 jobcode_2 於 jobcode_1 旁
    output_cols = [
        'fname', 'lname', 'local_date', 'local_day', 'local_start_time', 'local_end_time',
        'hours', 'Reg', 'OT', 'Adj Total', 'Hours', 'Minutes',
        'Rate', 'Loading', 'Adj Rate', 'Dollars', 'Job',
        'jobcode_1', 'jobcode_2', 'class', 'service item', 'notes', 'approved_status'
    ]
    
    for col in output_cols:
        if col not in df.columns:
            df[col] = ""
            
    df = df[output_cols]

    # 按 jobcode_1 排序
    df = df.sort_values(by='jobcode_1', ascending=True).reset_index(drop=True)
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
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_processed.to_excel(writer, index=False, sheet_name='Processed Payroll')
        ws = writer.sheets['Processed Payroll']
        
        # 取得 Hours 與 Minutes 欄位索引 (1-based)
        hours_col_idx = df_processed.columns.get_loc('Hours') + 1
        minutes_col_idx = df_processed.columns.get_loc('Minutes') + 1
        
        # 定義背景填滿色彩
        fill_hours = PatternFill(start_color="EAD1DC", end_color="EAD1DC", fill_type="solid")
        fill_minutes = PatternFill(start_color="D9D2E9", end_color="D9D2E9", fill_type="solid")
        
        # 套用顏色至整欄 (包含標題與數據列)
        for row in range(1, ws.max_row + 1):
            ws.cell(row=row, column=hours_col_idx).fill = fill_hours
            ws.cell(row=row, column=minutes_col_idx).fill = fill_minutes

    output.seek(0)
    
    filename = file.filename.replace('.csv', '_Processed.xlsx')
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
