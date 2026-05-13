import streamlit as st
import pandas as pd
import io
import openpyxl
import re
from openpyxl.styles import Font, PatternFill
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.formatting.rule import FormulaRule
from datetime import date

st.set_page_config(page_title="WIP Summary Consolidator Pro v25", layout="wide")

def extract_area_from_filename(filename):
    match = re.search(r"Open Orders List\s+(.*?)\s+\d+", filename, re.IGNORECASE)
    if match: return match.group(1).strip()
    parts = filename.replace('.xlsx', '').split(' ')
    return parts[3] if len(parts) >= 4 else "Unknown"

def parse_excel_wip(file_obj, filename):
    try:
        df_raw = pd.read_excel(file_obj, header=None)
    except Exception as e:
        st.error(f"Error reading {filename}: {e}")
        return []

    data = []
    today = date.today()
    area = extract_area_from_filename(filename)

    header_idx = -1
    col_map = {}
    for i in range(min(50, len(df_raw))):
        row_list = [str(x).strip() for x in df_raw.iloc[i]]
        if 'Pending Days' in row_list:
            header_idx = i
            for j, val in enumerate(row_list):
                if val and val != 'nan':
                    col_map[val] = j
            break
    
    if header_idx == -1: return []

    curr_dname = "Unknown"
    i = header_idx + 1
    
    while i < len(df_raw):
        row = df_raw.iloc[i]
        if str(row[0]).strip() == 'Dealer':
            curr_dname = str(df_raw.iloc[i, 6]).split('.')[0] if not pd.isna(df_raw.iloc[i, 6]) else "Unknown"
            i += 1
            continue
            
        try:
            val_0 = row[0]
            if not pd.isna(val_0) and (isinstance(val_0, (int, float)) or str(val_0).isdigit()):
                def get_v(label):
                    idx = col_map.get(label)
                    return row[idx] if idx is not None else None

                ord_no = str(get_v('Ord.No.')).split('.')[0]
                cr_dt_raw = get_v('Cr.Dt')
                
                # Calculation forced to Whole Number
                calc_pending = ""
                if cr_dt_raw:
                    try:
                        creation_date = pd.to_datetime(cr_dt_raw).date()
                        calc_pending = int((today - creation_date).days)
                    except:
                        try:
                            calc_pending = int(float(val_0))
                        except:
                            calc_pending = val_0

                record = {
                    "SNO.": "",
                    "Pending Days": calc_pending,
                    "D.Code &JC NO.": f"{curr_dname}{ord_no}",
                    "Dname": curr_dname,
                    "Area": area,
                    "Dealer Name": area,
                    "Req. Delv. Dt": get_v('Req. Delv. Dt'),
                    "Cr.Dt": cr_dt_raw,
                    "Total": get_v('Total'),
                    "PartsTotal": get_v('PartsTotal'),
                    "Ord.No.": ord_no,
                    "Ord.Ty.": get_v('Ord.Ty.'),
                    "Regn.No": get_v('Regn.No'),
                    "Chassis/SL No.": get_v('Chassis/SL No.'),
                    "Notes": "",
                    "Cust.No": get_v('Cust.No'),
                    "Cust Name": get_v('Cust Name'),
                    "Remarks": "",
                    "Category": "",
                    "Invoice no.": "",
                    "Date": "",
                    "Status": "Open"
                }
                
                if i + 1 < len(df_raw) and 'Notes' in [str(x).strip() for x in df_raw.iloc[i+1]]:
                    row_notes = df_raw.iloc[i+1]
                    for k in range(min(21, len(df_raw.columns)), len(df_raw.columns)):
                        if not pd.isna(row_notes[k]) and str(row_notes[k]).strip() not in ["", "Notes"]:
                            record['Notes'] = row_notes[k]
                            break
                    i += 1
                data.append(record)
        except: pass
        i += 1
    return data

def process_paycode_report(file_obj):
    try:
        df_raw = pd.read_excel(file_obj, header=None)
        header_row = -1
        for i, row in df_raw.head(25).iterrows():
            if "Ord.ID" in row.values:
                header_row = i
                break
        if header_row != -1:
            df = df_raw.iloc[header_row+1:].copy()
            df.columns = df_raw.iloc[header_row].values
            return set(df['Ord.ID'].dropna().astype(str).unique())
    except: pass
    return set()

def apply_rich_remarks(text):
    if not text or not isinstance(text, str) or text.strip() == "": return text
    red_bold = InlineFont(color="FFFF0000", b=True)
    normal = InlineFont()
    rt = CellRichText()
    parts = text.split('-', 1)
    rt.append(TextBlock(red_bold, parts[0]))
    if len(parts) > 1:
        suffix = "-" + parts[1]
        if "Job completed" in suffix:
            sub_parts = suffix.split("Job completed")
            for idx, p in enumerate(sub_parts):
                rt.append(TextBlock(normal, p))
                if idx < len(sub_parts) - 1: rt.append(TextBlock(red_bold, "Job completed"))
        else: rt.append(TextBlock(normal, suffix))
    return rt

st.title("📊 WIP Summary Consolidator Pro v25")

with st.sidebar:
    st.header("Files")
    summary_file = st.file_uploader("1. Existing Summary File", type=["xlsx"])
    open_order_files = st.file_uploader("2. New Open Order Lists", type=["xlsx"], accept_multiple_files=True)
    paycode_file = st.file_uploader("3. Paycodewise Report", type=["xlsx"])

if st.button("Generate & Format Report"):
    if not summary_file:
        st.warning("Please upload the Summary file.")
    else:
        wb = openpyxl.load_workbook(summary_file, keep_links=True)
        target_name = "Master Data" if "Master Data" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[target_name]

        data_rows = list(ws.iter_rows(values_only=True))
        header = data_rows[0]
        col_map_ws = {name: i for i, name in enumerate(header)}
        
        # Capture existing data and convert Pending Days to Int
        existing_data = []
        for row in data_rows[1:]:
            d = {header[i]: val for i, val in enumerate(row)}
            if 'Pending Days' in d and d['Pending Days'] is not None:
                try:
                    d['Pending Days'] = int(float(d