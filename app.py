import streamlit as st
import pandas as pd
import io
import datetime as dt
from datetime import time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go
import plotly.express as px

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Attendance Report System",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1A3C6E, #2D6099);
        color: white; padding: 20px 30px; border-radius: 10px;
        margin-bottom: 20px; text-align: center;
    }
    .metric-card {
        background: white; border-radius: 10px; padding: 15px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1); text-align: center;
        border-left: 4px solid #2D6099;
    }
    .status-p   { background:#E3F2FD; color:#0D47A1; padding:2px 8px; border-radius:4px; font-weight:bold; }
    .status-a   { background:#FFCDD2; color:#B71C1C; padding:2px 8px; border-radius:4px; font-weight:bold; }
    .status-wop { background:#C8E6C9; color:#1B5E20; padding:2px 8px; border-radius:4px; font-weight:bold; }
    .status-wo  { background:#EEEEEE; color:#555;    padding:2px 8px; border-radius:4px; font-weight:bold; }
    .status-miss{ background:#FFE0B2; color:#E65100; padding:2px 8px; border-radius:4px; font-weight:bold; }
    .miss-cell  { background:#FFE0B2 !important; }
    div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
    .stTabs [data-baseweb="tab"] { font-size: 15px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
SUNDAYS   = {1,8,15,22}
SATURDAYS = {7,14,21,28}
DAY_ABBR  = {d: dt.date(2026,2,d).strftime('%a') for d in range(1,29)}
GRACE     = 5/1440.0
OT_THR    = 29/1440.0

def parse_time_str(t):
    if pd.isna(t): return None
    s = str(t).strip()
    if ':' in s:
        try:
            p = s.split(':')
            return time(int(p[0]), int(p[1]))
        except: return None
    return None

def parse_time_cell(val):
    if pd.isna(val) or val == '' or val == 0: return None
    s = str(val).strip()
    if s in ('0','0.0','nan',''): return None
    if ':' in s and len(s) <= 5:
        try:
            p = s.split(':')
            h, m = int(p[0]), int(p[1])
            if 0 <= h <= 23 and 0 <= m <= 59:
                return time(h, m)
        except: pass
    if hasattr(val, 'hour'):
        return time(val.hour, val.minute)
    return None

def time_val(t):
    if t is None: return None
    return (t.hour*60 + t.minute) / 1440.0

def t2m(t):
    if t is None: return 0
    return t.hour*60 + t.minute

def m2hm(m):
    if not m or m <= 0: return "0:00"
    return f"{int(m//60)}:{int(m%60):02d}"

def calc_day(d, in_t, out_t, si_t, so_t, is_open=False, std_hours=9.5):
    is_sun = d in SUNDAYS
    r = dict(status='', work=0, ot=0, late=0, early=0)
    if in_t is None and out_t is None:
        r['status'] = 'WO' if is_sun else 'A'; return r
    if in_t is None or out_t is None:
        r['status'] = 'MISS'; return r

    in_m  = t2m(in_t)
    out_m = t2m(out_t)

    if is_sun:
        cap = 16*60
        if is_open:
            r['work']  = max(0, min(out_m, cap) - in_m)
        else:
            si_m = t2m(si_t) if si_t else 0
            eff_in_m = si_m if in_m <= si_m + 5 else in_m
            eff_in_m = max(si_m, eff_in_m)
            r['work'] = max(0, min(out_m, cap) - eff_in_m)
        r['status'] = 'WOP'
        return r

    if is_open:
        r['work']  = max(0, out_m - in_m)
        std_mins   = int(std_hours * 60)
        extra      = r['work'] - std_mins
        r['ot']    = extra if extra > 29 else 0
        r['late']  = 0
        r['early'] = 0
        r['status'] = 'P'
    else:
        si_m = t2m(si_t) if si_t else 0
        so_m = t2m(so_t) if so_t else 0
        eff_in_m = si_m if in_m <= si_m + 5 else in_m
        eff_in_m = max(si_m, eff_in_m)
        r['work']  = max(0, min(out_m, so_m) - eff_in_m)
        r['late']  = max(0, eff_in_m - si_m) if eff_in_m > si_m else 0
        r['early'] = max(0, so_m - out_m) if out_m < so_m else 0
        extra = out_m - so_m
        r['ot']    = extra if extra > 29 else 0
        r['status'] = 'P'
    return r

def load_master(file):
    df = pd.read_excel(file, sheet_name=0, header=None)
    
    # ── FIX: Drop fully-blank trailing columns before assigning names ──
    df = df.dropna(axis=1, how='all')
    
    # Now assign column names based on actual column count
    expected_cols = ['EmpCode','EmpName','Company','Department','Hour','Timing','ShiftIn','ShiftOut','Shift']
    actual_cols = len(df.columns)
    
    if actual_cols >= len(expected_cols):
        df = df.iloc[:, :len(expected_cols)]  # keep only first 9 columns
        df.columns = expected_cols
    else:
        # Pad with dummy names if fewer columns than expected
        dummy = [f'_col{i}' for i in range(actual_cols - len(expected_cols), 0)]
        df.columns = expected_cols[:actual_cols] + dummy

    df = df.iloc[1:].reset_index(drop=True)
    df['ShiftInTime']  = df['ShiftIn'].apply(parse_time_str)
    df['ShiftOutTime'] = df['ShiftOut'].apply(parse_time_str)
    df['EmpCode'] = df['EmpCode'].apply(lambda x: int(float(str(x))) if not pd.isna(x) else x)

    def is_open(row):
        shift_val = str(row['Shift']).strip().lower()
        if shift_val == 'open shift': return True
        if pd.isna(row['ShiftInTime']) or pd.isna(row['ShiftOutTime']): return True
        return False
    df['IsOpen'] = df.apply(is_open, axis=1)

    def get_std_hours(row):
        try: return float(str(row['Hour']).strip())
        except: return 9.5
    df['StdHours'] = df.apply(get_std_hours, axis=1)
    return df

def load_attendance(file):
    att = {}
    try:
        xl = pd.ExcelFile(file)
    except:
        return att

    for sheet_name in xl.sheet_names:
        try:
            df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        except:
            continue

        global_day_col_map = None
        for ri in range(len(df)):
            if str(df.iloc[ri, 0]).strip() == 'Days':
                global_day_col_map = {}
                for ci in range(len(df.iloc[ri])):
                    cv = str(df.iloc[ri, ci]).strip().split()
                    if cv and cv[0].isdigit():
                        global_day_col_map[ci] = int(cv[0])
                break

        i = 0
        while i < len(df):
            if str(df.iloc[i, 0]).strip() == 'Emp. Code:':
                emp_code = None
                for j in range(1, len(df.iloc[i])):
                    v = df.iloc[i, j]
                    if not pd.isna(v) and str(v).strip() not in ['','nan','Emp. Name:']:
                        try: emp_code = int(float(str(v))); break
                        except: pass
                if emp_code is None or i+3 >= len(df): i += 1; continue
                next_label = str(df.iloc[i+1, 0]).strip()
                if next_label == 'Status':
                    day_col_map = global_day_col_map
                    in_row = df.iloc[i+2]; out_row = df.iloc[i+3]
                else:
                    day_col_map = {}
                    for ci in range(len(df.iloc[i+1])):
                        cv = str(df.iloc[i+1, ci]).strip().split()
                        if cv and cv[0].isdigit():
                            day_col_map[ci] = int(cv[0])
                    in_row = df.iloc[i+2]; out_row = df.iloc[i+3]
                if emp_code not in att:
                    att[emp_code] = {}
                if day_col_map:
                    for ci, dn in day_col_map.items():
                        if ci < len(in_row) and ci < len(out_row):
                            att[emp_code][dn] = {
                                'in':  parse_time_cell(in_row.iloc[ci]),
                                'out': parse_time_cell(out_row.iloc[ci])
                            }
                i += 4
            else:
                i += 1
    return att

def get_gp_deduction(gate_passes, code, day):
    total = 0
    for gp in gate_passes:
        if gp['code'] == code and gp['day'] == day:
            total += gp['duration_mins']
    return total

def compute_summary(emp_df, attendance, gate_passes=None):
    if gate_passes is None: gate_passes = []
    rows = []
    for _, emp in emp_df.iterrows():
        try:    code = int(float(str(emp['EmpCode'])))
        except: code = str(emp['EmpCode'])
        si_t      = emp['ShiftInTime']
        so_t      = emp['ShiftOutTime']
        is_open   = bool(emp.get('IsOpen', False))
        std_hours = float(emp.get('StdHours', 9.5))
        emp_att   = attendance.get(code, {})
        totals = dict(P=0,WOP=0,WO=0,A=0,MISS=0,work=0,ot=0,late=0,early=0,late_days=0,early_days=0,gp=0,gp_days=0)
        for d in range(1,29):
            dd  = emp_att.get(d,{})
            r   = calc_day(d, dd.get('in'), dd.get('out'), si_t, so_t, is_open, std_hours)
            gp_mins = get_gp_deduction(gate_passes, code, d)
            st2 = r['status']
            if st2 in totals: totals[st2] += 1
            work_after_gp = max(0, r['work'] - gp_mins)
            totals['work']  += work_after_gp
            totals['ot']    += r['ot']
            totals['late']  += r['late']
            totals['early'] += r['early']
            totals['gp']    += gp_mins
            if gp_mins > 0: totals['gp_days'] += 1
            if r['late']  > 0: totals['late_days']  += 1
            if r['early'] > 0: totals['early_days'] += 1

        avg = totals['work'] / max(totals['P']+totals['WOP'],1)
        net = max(0, totals['work'] + totals['ot'] - totals['late'])
        shift_label = 'Open Shift' if is_open else f"{si_t.strftime('%H:%M') if si_t else '?'}–{so_t.strftime('%H:%M') if so_t else '?'}"
        rows.append({
            'Code': code, 'Name': str(emp['EmpName']),
            'Company': str(emp['Company']), 'Department': str(emp['Department']),
            'Shift': shift_label, 'Type': '🔓 Open' if is_open else '🔒 Fixed',
            'Present': totals['P'], 'WOP': totals['WOP'], 'W/Off': totals['WO'],
            'Absent': totals['A'], 'MISS': totals['MISS'],
            'Work Hrs': m2hm(totals['work']),
            'OT Hrs': m2hm(totals['ot']),
            'Late Hrs': m2hm(totals['late']), 'Late Days': totals['late_days'],
            'Early Hrs': m2hm(totals['early']), 'Early Days': totals['early_days'],
            'Gate Pass Hrs': m2hm(totals['gp']), 'GP Days': totals['gp_days'],
            'Dur+OT': m2hm(totals['work']+totals['ot']),
            'Net Work Hrs': m2hm(net),
            'Avg Hrs/Day': m2hm(int(avg)),
            '_work_mins': totals['work'], '_ot_mins': totals['ot'],
            '_present_total': totals['P']+totals['WOP'],
        })
    return pd.DataFrame(rows)

def build_excel(emp_df, attendance, gate_passes=None):
    if gate_passes is None: gate_passes = []
    thin=Side(style='thin'); med=Side(style='medium')
    def tb(): return Border(left=thin,right=thin,top=thin,bottom=thin)
    def mb(): return Border(left=med,right=med,top=med,bottom=med)
    def fl(h): return PatternFill('solid',start_color=h)
    def fn(bold=False,size=9,color='000000'):
        return Font(name='Arial',bold=bold,size=size,color=color)
    def al(h='center',v='center',wrap=True):
        return Alignment(horizontal=h,vertical=v,wrap_text=wrap)

    LABEL_COL=1; SHIFTIN_COL=2; SHIFTOUT_COL=3; DAY1_COL=4; DAY_END=31
    SUM1_COL=32
    SUM_LABELS=['Present','WOP','W/Off','Absent','MISS',
                'Total Work\nDuration','Total OT\nHrs',
                'Late By\nHrs','Late By\nDays',
                'Early By\nHrs','Early By\nDays',
                'Total Dur\n(+OT)','Avg Work\nHrs',
                'Net Work\nHrs\n(Work+OT\n−Late)',
                'Total\nDur+OT\n(All Days)']
    LAST_COL = SUM1_COL+len(SUM_LABELS)-1

    GRACE_F = 5.0/1440.0
    OT_F    = 29.0/1440.0

    wb = Workbook(); wb.remove(wb.active)

    for company in emp_df['Company'].unique():
        c_emps = emp_df[emp_df['Company']==company].copy()
        if c_emps.empty: continue
        ws = wb.create_sheet(title=str(company)[:31])

        for r,txt,sz in [(1,"Monthly Status Report (Detailed Work Duration)",14),
                         (2,"Feb 01 2026  To  Feb 28 2026",10)]:
            ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=LAST_COL)
            c=ws.cell(r,1,txt); c.font=fn(True,sz,'FFFFFF')
            c.fill=fl('1A3C6E'); c.alignment=al()
            ws.row_dimensions[r].height=22 if r==1 else 16

        ws.merge_cells(start_row=3,start_column=1,end_row=3,end_column=LAST_COL)
        ws.cell(3,1).fill=fl('1A3C6E'); ws.row_dimensions[3].height=4

        ws.merge_cells(start_row=4,start_column=1,end_row=4,end_column=8)
        c=ws.cell(4,1,f"Company:  {company}"); c.font=fn(True,10); c.alignment=al('left')
        ws.row_dimensions[4].height=16

        ws.row_dimensions[5].height=32
        for col,txt2,fg,bg in [(1,'Days','000000','EEF4FA'),(2,'Sh.In','1565C0','DDEEFF'),(3,'Sh.Out','1565C0','DDEEFF')]:
            c=ws.cell(5,col,txt2); c.font=fn(True,8,fg); c.fill=fl(bg); c.alignment=al(); c.border=tb()
        for d in range(1,29):
            col=DAY1_COL+d-1
            c=ws.cell(5,col,f"{d}\n{DAY_ABBR[d][0]}")
            if d in SUNDAYS:     c.fill=fl('FFA726'); c.font=fn(True,8,'7B1FA2')
            elif d in SATURDAYS: c.fill=fl('A5D6A7'); c.font=fn(True,8,'1B5E20')
            else:                c.fill=fl('D0E4F5'); c.font=fn(True,8)
            c.alignment=al(wrap=True); c.border=tb()
        for si2,sh in enumerate(SUM_LABELS):
            col=SUM1_COL+si2; c=ws.cell(5,col,sh)
            c.font=fn(True,8,'FFFFFF')
            c.fill=fl('4A235A') if si2==14 else fl('1F4E79')
            c.alignment=al(wrap=True); c.border=tb()

        current_row=6
        for dept in c_emps['Department'].unique():
            d_emps=c_emps[c_emps['Department']==dept]
            ws.merge_cells(start_row=current_row,start_column=1,end_row=current_row,end_column=LAST_COL)
            c=ws.cell(current_row,1,f"  Department:  {dept}")
            c.font=fn(True,10,'FFFFFF'); c.fill=fl('2D5F8A'); c.alignment=al('left')
            ws.row_dimensions[current_row].height=16; current_row+=1

            for _,emp in d_emps.iterrows():
                try:    emp_code=int(float(str(emp['EmpCode'])))
                except: emp_code=str(emp['EmpCode'])
                emp_name  = str(emp['EmpName'])
                si_t      = emp['ShiftInTime']; so_t = emp['ShiftOutTime']
                is_open   = bool(emp.get('IsOpen', False))
                std_hours = float(emp.get('StdHours', 9.5))
                std_mins_frac = (std_hours * 60) / 1440.0

                siv = time_val(si_t); sov = time_val(so_t)
                if is_open:
                    shift_str = f"Open Shift  (Std: {std_hours}h)"
                else:
                    shift_str = (f"{si_t.strftime('%H:%M')}–{so_t.strftime('%H:%M')}"
                                 if si_t and so_t else str(emp.get('Shift','Open')))
                emp_att=attendance.get(emp_code,{})

                HR=current_row; RS=HR+1; RI=HR+2; RO=HR+3
                RD=HR+4; RL=HR+5; RE=HR+6; ROT=HR+7; RDP=HR+8; RSP=HR+9

                for rr,h in [(HR,18),(RS,14),(RI,14),(RO,14),(RD,14),(RL,14),(RE,14),(ROT,14),(RDP,14),(RSP,3)]:
                    ws.row_dimensions[rr].height=h

                c=ws.cell(HR,1,f"  {emp_code} : {emp_name}")
                c.font=fn(True,9); c.fill=fl('E8F0E8' if is_open else 'E8F4FD')
                c.alignment=al('left'); c.border=tb()

                c=ws.cell(HR,SHIFTIN_COL)
                if is_open:
                    c.value = None
                    c.font=fn(False,8,'2E7D32'); c.fill=fl('E8F0E8')
                else:
                    c.value=siv; c.number_format='h:mm'
                    c.font=fn(False,8,'1565C0'); c.fill=fl('DDEEFF')
                c.alignment=al(); c.border=tb()

                c=ws.cell(HR,SHIFTOUT_COL)
                if is_open:
                    c.value = std_mins_frac
                    c.number_format='[h]:mm'
                    c.font=fn(False,8,'2E7D32'); c.fill=fl('E8F0E8')
                else:
                    c.value=sov; c.number_format='h:mm'
                    c.font=fn(False,8,'1565C0'); c.fill=fl('DDEEFF')
                c.alignment=al(); c.border=tb()

                ws.merge_cells(start_row=HR,start_column=4,end_row=HR,end_column=LAST_COL)
                hint = "Open Shift — Duration = Out−In | OT if Duration > Std Hrs | No Late/Early tracking" if is_open else "Fill MISS punch in InTime/OutTime rows ↓   |   All formulas auto-recalculate"
                c=ws.cell(HR,4,f"Shift: {shift_str}   |   {hint}")
                c.font=fn(False,8,'2E7D32' if is_open else '1A6B3A')
                c.fill=fl('E8F0E8' if is_open else 'E8F4FD')
                c.alignment=al('left'); c.border=tb()

                for rr2,lbl,bg in [(RS,'Status','F8FBFF'),(RI,'InTime','FFFFF0'),(RO,'OutTime','F0FFF0'),
                                   (RD,'Duration','F0F0FF'),(RL,'Late By','FFF3E0'),(RE,'Early By','FFF8F0'),
                                   (ROT,'OT','F0FFF4'),(RDP,'Dur+OT','EEE8F7')]:
                    c=ws.cell(rr2,1,lbl); c.font=fn(True,8); c.fill=fl('EEF4FA')
                    c.alignment=al(wrap=False); c.border=tb()
                    ws.cell(rr2,2).fill=fl(bg); ws.cell(rr2,2).border=tb()
                    ws.cell(rr2,3).fill=fl(bg); ws.cell(rr2,3).border=tb()
                for col in range(1,LAST_COL+1):
                    ws.cell(RSP,col).fill=fl('DDE8F0')

                si_ref=f"$B${HR}"; so_ref=f"$C${HR}"
                for d in range(1,29):
                    col=DAY1_COL+d-1; cl=get_column_letter(col); sun=d in SUNDAYS
                    dd=emp_att.get(d,{}); raw_in=dd.get('in'); raw_out=dd.get('out')
                    has_in=raw_in is not None; has_out=raw_out is not None
                    miss=(has_in and not has_out) or (not has_in and has_out)

                    in_bg  = 'FFE0B2' if miss else ('F0FFF8' if is_open else 'FFFFF0')
                    out_bg = 'FFE0B2' if miss else ('F0FFF8' if is_open else 'F0FFF0')

                    c=ws.cell(RI,col)
                    if has_in: c.value=time_val(raw_in); c.number_format='h:mm'
                    c.fill=fl(in_bg); c.font=fn(size=8); c.alignment=al(wrap=False); c.border=tb()

                    c=ws.cell(RO,col)
                    if has_out: c.value=time_val(raw_out); c.number_format='h:mm'
                    c.fill=fl(out_bg); c.font=fn(size=8); c.alignment=al(wrap=False); c.border=tb()

                    ir=f"{cl}{RI}"; or_=f"{cl}{RO}"

                    gp_mins = get_gp_deduction(gate_passes, emp_code, d)
                    gp_frac = gp_mins / 1440.0

                    if is_open:
                        if sun:
                            raw_dur = f'MAX(0,MIN({or_},TIME(16,0,0))-{ir})'
                        else:
                            raw_dur = f'MAX(0,{or_}-{ir})'
                    else:
                        eff = f"MAX({si_ref},IF({ir}<={si_ref}+{GRACE_F},{si_ref},{ir}))"
                        if sun:
                            raw_dur = f'MAX(0,MIN({or_},TIME(16,0,0))-({eff}))'
                        else:
                            raw_dur = f'MAX(0,MIN({or_},{so_ref})-({eff}))'

                    if gp_mins > 0:
                        df_ = f'=IF(OR({ir}="",{or_}=""),"",MAX(0,{raw_dur}-{gp_frac}))'
                        dur_fill = 'FFD580'
                    else:
                        df_ = f'=IF(OR({ir}="",{or_}=""),"",{raw_dur})'
                        dur_fill = 'E8F5E9' if is_open else 'F0F0FF'

                    c=ws.cell(RD,col); c.value=df_; c.number_format='[h]:mm'
                    c.font=fn(size=8); c.fill=fl(dur_fill); c.alignment=al(wrap=False); c.border=tb()

                    if is_open or sun:
                        lf = '=""'
                    else:
                        lf = f'=IF({ir}="","",IF({ir}>{si_ref}+{GRACE_F},{ir}-{si_ref},""))'
                    c=ws.cell(RL,col); c.value=lf; c.number_format='[h]:mm'
                    c.font=fn(size=8,color='B85C00'); c.fill=fl('FFF3E0'); c.alignment=al(wrap=False); c.border=tb()

                    if is_open or sun:
                        ef = '=""'
                    else:
                        ef = f'=IF({or_}="","",IF({or_}<{so_ref},{so_ref}-{or_},""))'
                    c=ws.cell(RE,col); c.value=ef; c.number_format='[h]:mm'
                    c.font=fn(size=8,color='795548'); c.fill=fl('FFF8F0'); c.alignment=al(wrap=False); c.border=tb()

                    dr_ref = f"{cl}{RD}"
                    if sun:
                        of = '=""'
                    elif is_open:
                        of = (f'=IF({or_}="","",IF({dr_ref}>{so_ref}+{OT_F},'
                              f'{dr_ref}-{so_ref},""))')
                    else:
                        of = f'=IF({or_}="","",IF({or_}-{so_ref}>{OT_F},{or_}-{so_ref},""))'
                    c=ws.cell(ROT,col); c.value=of; c.number_format='[h]:mm'
                    c.font=fn(size=8,color='2E7D32'); c.fill=fl('F0FFF4'); c.alignment=al(wrap=False); c.border=tb()

                    ot_ref=f"{cl}{ROT}"
                    dpf=(f'=IF(AND({dr_ref}="",{ot_ref}=""),"",IFERROR('
                         f'IF(ISNUMBER({dr_ref}),{dr_ref},0)+IF(ISNUMBER({ot_ref}),{ot_ref},0),""))')
                    c=ws.cell(RDP,col); c.value=dpf; c.number_format='[h]:mm'
                    c.font=fn(size=8,color='4A235A'); c.fill=fl('EEE8F7'); c.alignment=al(wrap=False); c.border=tb()

                    sf=(f'=IF(AND({ir}="",{or_}=""),"WO",IF(OR({ir}="",{or_}=""),"MISS","WOP"))' if sun
                        else f'=IF(AND({ir}="",{or_}=""),"A",IF(OR({ir}="",{or_}=""),"MISS","P"))')
                    c=ws.cell(RS,col); c.value=sf; c.font=fn(size=8)
                    c.fill=fl('F8FBFF'); c.alignment=al(wrap=False); c.border=tb()

                sc2=get_column_letter(DAY1_COL); ec2=get_column_letter(DAY_END)
                sst=f"{sc2}{RS}:{ec2}{RS}"; sdr=f"{sc2}{RD}:{ec2}{RD}"
                slr=f"{sc2}{RL}:{ec2}{RL}"; ser=f"{sc2}{RE}:{ec2}{RE}"; sor=f"{sc2}{ROT}:{ec2}{ROT}"
                def ss(rng): return f"SUMPRODUCT(IF(ISNUMBER({rng}),{rng},0))"
                sw=ss(sdr); so2=ss(sor); sl=ss(slr); se2=ss(ser)
                pc=f'COUNTIF({sst},"P")'; wc=f'COUNTIF({sst},"WOP")'

                fmls=[
                    f'=COUNTIF({sst},"P")',f'=COUNTIF({sst},"WOP")',f'=COUNTIF({sst},"WO")',
                    f'=COUNTIF({sst},"A")',f'=COUNTIF({sst},"MISS")',
                    f'=IF({sw}=0,"0:00",TEXT({sw},"[h]:mm"))',
                    f'=IF({so2}=0,"0:00",TEXT({so2},"[h]:mm"))',
                    f'=IF({sl}=0,"0:00",TEXT({sl},"[h]:mm"))',
                    f'=SUMPRODUCT(({slr}<>"")*ISNUMBER({slr})*1)',
                    f'=IF({se2}=0,"0:00",TEXT({se2},"[h]:mm"))',
                    f'=SUMPRODUCT(({ser}<>"")*ISNUMBER({ser})*1)',
                    f'=IF(({sw}+{so2})=0,"0:00",TEXT({sw}+{so2},"[h]:mm"))',
                    f'=IFERROR(IF({sw}=0,"0:00",TEXT({sw}/MAX({pc}+{wc},1),"[h]:mm")),"0:00")',
                    f'=IF(MAX({sw}+{so2}-{sl},0)=0,"0:00",TEXT(MAX({sw}+{so2}-{sl},0),"[h]:mm"))',
                    f'=IF(({sw}+{so2})=0,"0:00",TEXT({sw}+{so2},"[h]:mm"))',
                ]
                for si_idx,formula in enumerate(fmls):
                    col=SUM1_COL+si_idx
                    ws.merge_cells(start_row=RS,start_column=col,end_row=RDP,end_column=col)
                    c=ws.cell(RS,col,formula)
                    c.font=fn(True,9,'4A235A') if si_idx==14 else fn(True,9)
                    c.fill=fl('EEE8F7') if si_idx==14 else fl('EBF5FB')
                    c.alignment=al(); c.border=mb()
                current_row=RSP+1

        ws.column_dimensions['A'].width=10; ws.column_dimensions['B'].width=5.5; ws.column_dimensions['C'].width=5.5
        for d in range(1,29): ws.column_dimensions[get_column_letter(DAY1_COL+d-1)].width=6.2
        for si3,sw_ in enumerate([8,6,6,7,6,11,9,10,8,10,8,12,10,11,11]):
            ws.column_dimensions[get_column_letter(SUM1_COL+si3)].width=sw_
        ws.freeze_panes="D6"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
for key in ['emp_df','attendance','summary_df','miss_edits','gate_passes']:
    if key not in st.session_state:
        st.session_state[key] = None
if 'miss_edits' not in st.session_state or st.session_state.miss_edits is None:
    st.session_state.miss_edits = {}
if 'gate_passes' not in st.session_state or st.session_state.gate_passes is None:
    st.session_state.gate_passes = []

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📋 Attendance System")
    st.markdown("---")
    st.markdown("### Step 1: Upload Files")

    master_file = st.file_uploader("👥 Employee Master (.xlsx)", type=['xlsx','xls'], key='master_up')
    att_file    = st.file_uploader("🕐 Attendance Data (.xls/.xlsx)", type=['xlsx','xls'], key='att_up')

    if master_file and att_file:
        if st.button("⚙️ Process Files", use_container_width=True, type='primary'):
            with st.spinner("Loading data..."):
                try:
                    emp_df = load_master(master_file)
                    attendance = load_attendance(att_file)

                    master_codes = set(emp_df['EmpCode'].tolist())
                    att_codes    = set(attendance.keys())
                    missing_from_master = att_codes - master_codes

                    st.session_state.emp_df     = emp_df
                    st.session_state.attendance = attendance
                    st.session_state.summary_df = compute_summary(emp_df, attendance, st.session_state.gate_passes)
                    st.session_state.miss_edits = {}

                    st.success(f"✅ {len(emp_df)} employees | {len(attendance)} in attendance")
                    if missing_from_master:
                        st.warning(f"⚠️ {len(missing_from_master)} employees in attendance not in master: {sorted(missing_from_master)}")
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.exception(e)

    st.markdown("---")
    if st.session_state.emp_df is not None:
        df = st.session_state.emp_df
        st.markdown(f"**Employees:** {len(df)}")
        st.markdown(f"**Companies:** {df['Company'].nunique()}")
        st.markdown(f"**Departments:** {df['Department'].nunique()}")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
st.markdown('<div class="main-header"><h1>📋 Monthly Attendance Report System</h1><p>Feb 2026 — Formula-Enabled Excel Generator</p></div>', unsafe_allow_html=True)

if st.session_state.emp_df is None:
    st.info("👈 Upload Employee Master and Attendance files in the sidebar to get started.")
    with st.expander("📖 How to use"):
        st.markdown("""
        1. **Upload Employee Master** — the xlsx with employee codes, names, companies, shifts
        2. **Upload Attendance XLS** — the BasicWorkDurationReport file from ESSL
        3. Click **Process Files**
        4. Use the tabs to: view dashboard, fix MISS punches, and download the report
        """)
    st.stop()

emp_df     = st.session_state.emp_df
attendance = st.session_state.attendance
summary_df = st.session_state.summary_df

# Apply any MISS edits to attendance
att_working = {k: {d: dict(v) for d,v in v2.items()} for k,v2 in attendance.items()}
for (code, day), edits in st.session_state.miss_edits.items():
    if code not in att_working: att_working[code] = {}
    if day not in att_working[code]: att_working[code][day] = {'in':None,'out':None}
    if 'in'  in edits: att_working[code][day]['in']  = edits['in']
    if 'out' in edits: att_working[code][day]['out'] = edits['out']

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Dashboard", "👁️ View Report", "✏️ Fix MISS Punches", "🚪 Gate Pass", "⬇️ Download Excel"])

# ─────────── TAB 1: DASHBOARD ────────────────
with tab1:
    st.subheader("📊 Attendance Summary Dashboard")

    total_p   = summary_df['Present'].sum()
    total_a   = summary_df['Absent'].sum()
    total_wop = summary_df['WOP'].sum()
    total_miss= summary_df['MISS'].sum()
    avg_work  = summary_df['_work_mins'].mean()

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("👥 Total Employees", len(summary_df))
    c2.metric("✅ Total Present Days", total_p)
    c3.metric("❌ Total Absent Days", total_a)
    c4.metric("🟠 MISS Punches", total_miss)
    c5.metric("⏱️ Avg Work Hrs/Day", m2hm(int(avg_work)))

    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        comp_data = summary_df.groupby('Company')[['Present','WOP','Absent','MISS']].sum().reset_index()
        fig = px.bar(comp_data.melt(id_vars='Company', var_name='Status', value_name='Days'),
                     x='Company', y='Days', color='Status', barmode='group',
                     color_discrete_map={'Present':'#1976D2','WOP':'#388E3C','Absent':'#D32F2F','MISS':'#F57C00'},
                     title="Attendance by Company")
        fig.update_layout(height=350, margin=dict(l=0,r=0,t=40,b=0))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        late_df = summary_df[['Name','_work_mins','_ot_mins','Late Days']].nlargest(10,'Late Days')
        fig2 = px.bar(late_df, x='Name', y='Late Days', title="Top 10 Late Employees",
                      color='Late Days', color_continuous_scale='Oranges')
        fig2.update_layout(height=350, margin=dict(l=0,r=0,t=40,b=0), xaxis_tickangle=-30)
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        ot_data = summary_df.groupby('Company')['_ot_mins'].sum().reset_index()
        ot_data['OT Hrs'] = ot_data['_ot_mins']/60
        fig3 = px.pie(ot_data, names='Company', values='OT Hrs', title="OT Hours Distribution")
        fig3.update_layout(height=320, margin=dict(l=0,r=0,t=40,b=0))
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        summary_df['Work Hrs Num'] = summary_df['_work_mins']/60
        fig4 = px.histogram(summary_df, x='Work Hrs Num', nbins=20,
                            title="Work Hours Distribution (All Employees)",
                            color_discrete_sequence=['#1565C0'])
        fig4.update_layout(height=320, margin=dict(l=0,r=0,t=40,b=0))
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown("---")
    st.subheader("📋 Full Summary Table")

    fc1,fc2,fc3 = st.columns(3)
    sel_company = fc1.multiselect("Company", options=summary_df['Company'].unique(), default=list(summary_df['Company'].unique()))
    sel_dept    = fc2.multiselect("Department", options=summary_df['Department'].unique(), default=list(summary_df['Department'].unique()))
    show_miss   = fc3.checkbox("Show only MISS employees", value=False)

    filtered = summary_df[summary_df['Company'].isin(sel_company) & summary_df['Department'].isin(sel_dept)]
    if show_miss:
        filtered = filtered[filtered['MISS'] > 0]

    display_cols = ['Code','Name','Company','Department','Type','Shift','Present','WOP','W/Off','Absent','MISS','Work Hrs','OT Hrs','Gate Pass Hrs','GP Days','Dur+OT','Late Hrs','Late Days','Net Work Hrs','Avg Hrs/Day']
    st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True,
                 column_config={
                     'Present': st.column_config.NumberColumn('✅ P', width='small'),
                     'Absent':  st.column_config.NumberColumn('❌ A', width='small'),
                     'MISS':    st.column_config.NumberColumn('🟠 M', width='small'),
                     'WOP':     st.column_config.NumberColumn('🟢 WOP', width='small'),
                 })

# ─────────── TAB 2: VIEW REPORT ──────────────
with tab2:
    st.subheader("👁️ View Attendance Report")

    companies = list(emp_df['Company'].unique())
    sel_comp  = st.selectbox("Select Company", companies)
    c_emps    = emp_df[emp_df['Company']==sel_comp]
    depts     = list(c_emps['Department'].unique())
    sel_dept2 = st.selectbox("Select Department", depts)
    d_emps    = c_emps[c_emps['Department']==sel_dept2]

    for _, emp in d_emps.iterrows():
        try:    code = int(float(str(emp['EmpCode'])))
        except: code = str(emp['EmpCode'])
        si_t      = emp['ShiftInTime']; so_t = emp['ShiftOutTime']
        is_open   = bool(emp.get('IsOpen', False))
        std_hours = float(emp.get('StdHours', 9.5))
        emp_att   = att_working.get(code, {})
        shift_label = f"Open Shift (Std: {std_hours}h)" if is_open else f"{si_t.strftime('%H:%M') if si_t else '?'} – {so_t.strftime('%H:%M') if so_t else '?'}"

        with st.expander(f"{'🔓' if is_open else '🔒'} {code} : {emp['EmpName']}  |  {shift_label}"):
            rows = {'Status':[], 'InTime':[], 'OutTime':[], 'Duration':[], 'OT':[], 'Dur+OT':[]}

            for d in range(1,29):
                dd   = emp_att.get(d,{})
                r    = calc_day(d, dd.get('in'), dd.get('out'), si_t, so_t, is_open, std_hours)
                in_t  = dd.get('in')
                out_t = dd.get('out')
                rows['Status'].append(r['status'] if r['status'] else '')
                rows['InTime'].append(in_t.strftime('%H:%M') if in_t else '')
                rows['OutTime'].append(out_t.strftime('%H:%M') if out_t else '')
                rows['Duration'].append(m2hm(r['work']) if r['work']>0 else '')
                rows['OT'].append(m2hm(r['ot']) if r['ot']>0 else '')
                rows['Dur+OT'].append(m2hm(r['work']+r['ot']) if (r['work']+r['ot'])>0 else '')

            view_df = pd.DataFrame(rows, index=['Status','InTime','OutTime','Duration','OT','Dur+OT'])
            view_df.columns = [f"{d} {DAY_ABBR[d][:2]}" for d in range(1,29)]

            def color_status(val):
                colors = {'P':'background-color:#E3F2FD;color:#0D47A1;font-weight:bold',
                          'A':'background-color:#FFCDD2;color:#B71C1C;font-weight:bold',
                          'WOP':'background-color:#C8E6C9;color:#1B5E20;font-weight:bold',
                          'WO':'background-color:#EEEEEE;color:#555;font-weight:bold',
                          'MISS':'background-color:#FFE0B2;color:#E65100;font-weight:bold'}
                return colors.get(val,'')

            styled = view_df.style.applymap(color_status, subset=pd.IndexSlice[['Status'],:])
            st.dataframe(styled, use_container_width=True)

            sc1,sc2,sc3,sc4,sc5,sc6,sc7 = st.columns(7)
            sr = summary_df[summary_df['Code']==code]
            if not sr.empty:
                row = sr.iloc[0]
                sc1.metric("Present",   row['Present'])
                sc2.metric("WOP",       row['WOP'])
                sc3.metric("Absent",    row['Absent'])
                sc4.metric("MISS",      row['MISS'])
                sc5.metric("Work Hrs",  row['Work Hrs'])
                sc6.metric("OT Hrs",    row['OT Hrs'])
                sc7.metric("Dur+OT",    row['Dur+OT'])

# ─────────── TAB 3: FIX MISS ─────────────────
with tab3:
    st.subheader("✏️ Fix MISS Punches")
    st.info("🟠 Orange cells below have a missing punch. Enter the missing time and click Save.")

    miss_list = []
    for _, emp in emp_df.iterrows():
        try:    code = int(float(str(emp['EmpCode'])))
        except: code = str(emp['EmpCode'])
        si_t = emp['ShiftInTime']; so_t = emp['ShiftOutTime']
        emp_att = att_working.get(code, {})
        for d in range(1,29):
            dd = emp_att.get(d, {})
            in_t = dd.get('in'); out_t = dd.get('out')
            if (in_t is None) != (out_t is None):
                miss_list.append({
                    'Code': code, 'Name': str(emp['EmpName']),
                    'Company': str(emp['Company']), 'Department': str(emp['Department']),
                    'Day': d, 'Date': f"Feb {d:02d}",
                    'InTime': in_t.strftime('%H:%M') if in_t else '❌ MISSING',
                    'OutTime': out_t.strftime('%H:%M') if out_t else '❌ MISSING',
                    '_missing': 'in' if in_t is None else 'out'
                })

    if not miss_list:
        st.success("✅ No MISS punches found! All employees have complete punch data.")
    else:
        st.warning(f"Found **{len(miss_list)} MISS punches** across {len(set(m['Code'] for m in miss_list))} employees")

        miss_df = pd.DataFrame(miss_list)
        fc1, fc2 = st.columns(2)
        f_comp = fc1.selectbox("Filter by Company", ['All']+list(miss_df['Company'].unique()))
        f_dept = fc2.selectbox("Filter by Dept",    ['All']+list(miss_df['Department'].unique()))

        filtered_miss = miss_df.copy()
        if f_comp != 'All': filtered_miss = filtered_miss[filtered_miss['Company']==f_comp]
        if f_dept != 'All': filtered_miss = filtered_miss[filtered_miss['Department']==f_dept]

        st.markdown(f"Showing **{len(filtered_miss)}** MISS punches")

        for _, miss_row in filtered_miss.iterrows():
            code   = miss_row['Code']
            day    = miss_row['Day']
            emp_name = miss_row['Name']
            missing_side = miss_row['_missing']

            col_a, col_b, col_c, col_d, col_e = st.columns([3,2,2,2,1])
            col_a.markdown(f"**{code} : {emp_name}**  \n`Feb {day:02d} ({DAY_ABBR[day]})`")
            col_b.markdown(f"**InTime:** `{miss_row['InTime']}`")
            col_c.markdown(f"**OutTime:** `{miss_row['OutTime']}`")

            edit_key = f"miss_{code}_{day}_{missing_side}"
            new_time = col_d.text_input(
                f"Enter {'InTime' if missing_side=='in' else 'OutTime'}",
                placeholder="e.g. 09:05",
                key=edit_key,
                label_visibility='collapsed'
            )
            if col_e.button("💾", key=f"save_{code}_{day}"):
                if new_time.strip():
                    t = parse_time_str(new_time.strip())
                    if t:
                        edit_key2 = (code, day)
                        if edit_key2 not in st.session_state.miss_edits:
                            st.session_state.miss_edits[edit_key2] = {}
                        st.session_state.miss_edits[edit_key2][missing_side] = t
                        att_working[code][day][missing_side] = t
                        st.session_state.summary_df = compute_summary(emp_df, att_working, st.session_state.gate_passes)
                        st.success(f"✅ Saved {new_time} for {emp_name} Day {day}")
                        st.rerun()
                    else:
                        st.error("Invalid time format. Use HH:MM")

        if st.session_state.miss_edits:
            st.markdown("---")
            st.success(f"✅ {len(st.session_state.miss_edits)} MISS punches fixed so far")
            if st.button("🔄 Reset All Edits"):
                st.session_state.miss_edits = {}
                st.session_state.summary_df = compute_summary(emp_df, attendance, st.session_state.gate_passes)
                st.rerun()

# ─────────── TAB 4: GATE PASS ────────────────
with tab4:
    st.subheader("🚪 Gate Pass Manager")
    st.markdown("Record mid-day outside visits. Gate pass duration is **automatically deducted from work hours**.")

    gate_passes = st.session_state.gate_passes

    with st.container(border=True):
        st.markdown("### ➕ Add New Gate Pass")

        ga, gb = st.columns(2)
        emp_options = {f"{int(float(str(r['EmpCode'])))} — {r['EmpName']} ({r['Company']})": int(float(str(r['EmpCode'])))
                       for _, r in emp_df.iterrows()}
        selected_label = ga.selectbox("👤 Select Employee", options=list(emp_options.keys()), key="gp_emp")
        selected_code  = emp_options[selected_label]
        selected_name  = selected_label.split('—')[1].split('(')[0].strip()

        gp_day = gb.selectbox("📅 Date", options=list(range(1,29)),
                              format_func=lambda d: f"Feb {d:02d} ({DAY_ABBR[d]})", key="gp_day")

        gc, gd, ge, gf = st.columns(4)
        gp_out_str  = gc.text_input("🚶 Out Time (left)",    placeholder="e.g. 12:30", key="gp_out")
        gp_in_str   = gd.text_input("🔙 In Time (returned)", placeholder="e.g. 14:15", key="gp_in")
        gp_reason   = ge.text_input("📝 Reason / Purpose",   placeholder="e.g. Bank work", key="gp_reason")
        gp_approved = gf.text_input("✅ Approved By",         placeholder="e.g. Manager", key="gp_approved")

        if st.button("➕ Add Gate Pass", type="primary", use_container_width=True):
            gp_out_t = parse_time_str(gp_out_str.strip()) if gp_out_str.strip() else None
            gp_in_t  = parse_time_str(gp_in_str.strip())  if gp_in_str.strip()  else None
            if not gp_out_t or not gp_in_t:
                st.error("❌ Please enter valid Out Time and In Time (format: HH:MM)")
            elif t2m(gp_in_t) <= t2m(gp_out_t):
                st.error("❌ In Time must be after Out Time")
            else:
                dur = t2m(gp_in_t) - t2m(gp_out_t)
                new_gp = {
                    'code': selected_code,
                    'name': selected_name,
                    'day': gp_day,
                    'date_str': f"Feb {gp_day:02d} ({DAY_ABBR[gp_day]})",
                    'gp_out': gp_out_t.strftime('%H:%M'),
                    'gp_in':  gp_in_t.strftime('%H:%M'),
                    'duration_mins': dur,
                    'duration_str': m2hm(dur),
                    'reason': gp_reason.strip() or '—',
                    'approved_by': gp_approved.strip() or '—',
                }
                st.session_state.gate_passes.append(new_gp)
                st.session_state.summary_df = compute_summary(emp_df, att_working, st.session_state.gate_passes)
                st.success(f"✅ Gate pass added for **{selected_name}** on Feb {gp_day} — {m2hm(dur)} deducted")
                st.rerun()

    st.markdown("---")

    if not gate_passes:
        st.info("📭 No gate passes recorded yet. Add one above.")
    else:
        st.markdown(f"### 📋 Recorded Gate Passes ({len(gate_passes)} total)")

        gf1, gf2 = st.columns(2)
        gp_df_raw = pd.DataFrame(gate_passes)

        filter_emp = gf1.selectbox("Filter by Employee",
                                   ['All'] + sorted(gp_df_raw['name'].unique().tolist()),
                                   key="gp_filter_emp")
        filter_day = gf2.selectbox("Filter by Date",
                                   ['All'] + [f"Feb {d:02d} ({DAY_ABBR[d]})" for d in range(1,29)],
                                   key="gp_filter_day")

        filtered_gp = gate_passes.copy()
        if filter_emp != 'All':
            filtered_gp = [g for g in filtered_gp if g['name'] == filter_emp]
        if filter_day != 'All':
            day_num = int(filter_day.split()[1])
            filtered_gp = [g for g in filtered_gp if g['day'] == day_num]

        disp_rows = []
        for i, gp in enumerate(filtered_gp):
            disp_rows.append({
                '#': gate_passes.index(gp),
                'Emp Code': gp['code'],
                'Employee': gp['name'],
                'Date': gp['date_str'],
                'Out Time': gp['gp_out'],
                'In Time': gp['gp_in'],
                'Duration': gp['duration_str'],
                'Reason': gp['reason'],
                'Approved By': gp['approved_by'],
            })

        gp_display_df = pd.DataFrame(disp_rows)
        st.dataframe(gp_display_df.drop(columns=['#']), use_container_width=True, hide_index=True,
                     column_config={
                         'Duration': st.column_config.TextColumn('⏱️ Duration', width='small'),
                         'Out Time': st.column_config.TextColumn('🚶 Out', width='small'),
                         'In Time':  st.column_config.TextColumn('🔙 In',  width='small'),
                     })

        st.markdown("### 📊 Gate Pass Summary by Employee")
        gp_summary = {}
        for gp in gate_passes:
            k = (gp['code'], gp['name'])
            if k not in gp_summary:
                gp_summary[k] = {'count': 0, 'total_mins': 0, 'days': set()}
            gp_summary[k]['count']      += 1
            gp_summary[k]['total_mins'] += gp['duration_mins']
            gp_summary[k]['days'].add(gp['day'])

        sum_rows = []
        for (code, name), v in sorted(gp_summary.items()):
            sum_rows.append({
                'Code': code, 'Employee': name,
                'Gate Passes': v['count'],
                'Days Used': len(v['days']),
                'Total Time Out': m2hm(v['total_mins']),
            })
        st.dataframe(pd.DataFrame(sum_rows), use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("### 🗑️ Delete a Gate Pass")
        del_col1, del_col2 = st.columns([3,1])
        del_options = {f"#{i} | {gp['name']} | {gp['date_str']} | {gp['gp_out']}–{gp['gp_in']} | {gp['reason']}": i
                       for i, gp in enumerate(gate_passes)}
        if del_options:
            del_sel = del_col1.selectbox("Select gate pass to delete", list(del_options.keys()), key="gp_del_sel")
            if del_col2.button("🗑️ Delete", key="gp_del_btn"):
                idx_to_del = del_options[del_sel]
                st.session_state.gate_passes.pop(idx_to_del)
                st.session_state.summary_df = compute_summary(emp_df, att_working, st.session_state.gate_passes)
                st.success("✅ Gate pass deleted")
                st.rerun()

        if st.button("🔄 Clear ALL Gate Passes", key="gp_clear_all"):
            st.session_state.gate_passes = []
            st.session_state.summary_df = compute_summary(emp_df, att_working, st.session_state.gate_passes)
            st.rerun()

        st.markdown("---")
        if gate_passes:
            gp_export = pd.DataFrame([{
                'Emp Code': g['code'], 'Employee': g['name'],
                'Date': g['date_str'], 'Out Time': g['gp_out'],
                'In Time': g['gp_in'], 'Duration': g['duration_str'],
                'Reason': g['reason'], 'Approved By': g['approved_by']
            } for g in gate_passes])
            csv_buf = io.StringIO()
            gp_export.to_csv(csv_buf, index=False)
            st.download_button(
                "📥 Export Gate Passes to CSV",
                data=csv_buf.getvalue(),
                file_name="GatePasses_Feb2026.csv",
                mime="text/csv",
            )

# ─────────── TAB 5: DOWNLOAD ─────────────────
with tab5:
    st.subheader("⬇️ Download Formula-Enabled Excel Report")
    st.markdown("""
    The downloaded Excel file includes:
    - ✅ **Company sheets** — each with all employees
    - ✅ **9 rows per employee**: Status, InTime, OutTime, Duration, Late By, Early By, OT, Dur+OT
    - ✅ **All formulas live** — edit any InTime/OutTime and everything recalculates
    - ✅ **MISS punches** highlighted in orange — just type the missing time
    - ✅ **15 summary columns** including Net Work Hrs, Total Dur+OT
    - ✅ **All attendance rules** applied: Grace period, OT threshold, Sunday WOP
    - ✅ **Gate pass deductions** applied to Duration where recorded
    """)

    if st.session_state.gate_passes:
        st.info(f"🚪 **{len(st.session_state.gate_passes)} gate passes** will be included in this report")

    if st.button("🔨 Generate Excel Report", type="primary", use_container_width=True):
        with st.spinner("Building Excel report... this may take 30-60 seconds for large employee sets"):
            try:
                buf = build_excel(emp_df, att_working, st.session_state.gate_passes)
                st.download_button(
                    label="📥 Download Attendance_Feb2026_Report.xlsx",
                    data=buf,
                    file_name="Attendance_Feb2026_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                st.success("✅ Report ready! Click the button above to download.")
            except Exception as e:
                st.error(f"Error generating report: {e}")
                st.exception(e)
