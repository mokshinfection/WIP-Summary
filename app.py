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

st.set_page_config(page_title="WIP Summary Consolidator Pro v23", layout="wide")

def extract_area_from_filename(filename):
    """Extracts area name from filename."""
    match = re.search(r"Open Orders List\s+(.*?)\s+\d+", filename, re.IGNORECASE)
    if match: return match.group(1).strip()
    parts = filename.replace('.xlsx', '').split(' ')
    return parts[3] if len(parts) >= 4 else "Unknown"

def parse_excel_wip(file_obj, filename):
    """Parses Excel using the scanning logic from the user-provided file."""
    try:
        df_raw = pd.read_excel(file_obj, header=None)
    except Exception as e:
        st.error(f"Error reading {filename}: {e}")
        return []

    data = []
    today = date.today()
    area = extract_area_from_filename(filename)

    # Find the Main Header Row (looking for 'Pending Days')
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
        
        # Detect Dealer Context row
        if str(row[0]).strip() == 'Dealer':
            # Column index 6 for Dname as per provided app.py
            curr_dname = str(df_raw.iloc[i, 6]).split('.')[0] if not pd.isna(df_raw.iloc[i, 6]) else "Unknown"
            i += 1
            continue
            
        try:
            val_0 = row[0]
            # Check if row is a data row (starts with an index/number)
            if not pd.isna(val_0) and (isinstance(val_0, (int, float)) or str(val_0).isdigit()):
                def get_v(label):
                    idx = col_map.get(label)
                    return row[idx] if idx is not None else None

                ord_no = str(get_v('Ord.No.')).split('.')[0]
                cr_dt_raw = get_v('Cr.Dt')
                
                # Calculate Pending Days
                calc_pending = ""
                if cr_dt_raw:
                    try:
                        creation_date = pd.to_datetime(cr_dt_raw).date()
                        calc_pending = (today - creation_date).days
                    except:
                        calc_pending = val_0

                record = {
                    "SNO.": "",
                    "Pending Days": calc_pending,
                    "D.Code &JC NO.": f"{curr_dname}{ord_no}",
                    "Dname": curr_dname,
                    "Area": area,
                    "Dealer Name": area, # CHANGED: Just the area name
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
                
                # Extract Notes (scan following row)
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
    """Extracts Order IDs from Paycodewise report."""
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
    """Applies Red/Bold formatting based on rules."""
    if not text or not isinstance(text, str) or text.strip() == "": return text
    red_bold = InlineFont(color="FFFF0000", b=True)
    normal = InlineFont()
    rt = CellRichText()
    
    parts = text.split('-', 1)
    prefix = parts[0]
    rt.append(TextBlock(red_bold, prefix))
    
    if len(parts) > 1:
        suffix = "-" + parts[1]
        if "Job completed" in suffix:
            sub_parts = suffix.split("Job completed")
            for idx, p in enumerate(sub_parts):
                rt.append(TextBlock(normal, p))
                if idx < len(sub_parts) - 1:
                    rt.append(TextBlock(red_bold, "Job completed"))
        else:
            rt.append(TextBlock(normal, suffix))
    else:
        if "Job completed" in text:
            sub_parts = text.split("Job completed")
            for idx, p in enumerate(sub_parts):
                rt.append(TextBlock(normal, p))
                if idx < len(sub_parts) - 1:
                    rt.append(TextBlock(red_bold, "Job completed"))
        else:
            return text
    return rt

st.title("📊 WIP Summary Consolidator Pro v23")

with st.sidebar:
    st.header("Files")
    summary_file = st.file_uploader("1. Existing Summary File", type=["xlsx"])
    open_order_files = st.file_uploader("2. New Open Order Lists", type=["xlsx"], accept_multiple_files=True)
    paycode_file = st.file_uploader("3. Paycodewise Report", type=["xlsx"])

if st.button("Generate Final Report"):
    if not summary_file:
        st.warning("Please upload the Summary file.")
    else:
        wb = openpyxl.load_workbook(summary_file, keep_links=True)
        target_name = "Master Data" if "Master Data" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[target_name]

        data_rows = list(ws.iter_rows(values_only=True))
        header = data_rows[0]
        col_map_ws = {name: i for i, name in enumerate(header)}
        existing_data = [{header[i]: val for i, val in enumerate(row)} for row in data_rows[1:]]

        new_data_raw = []
        for f in open_order_files:
            new_data_raw.extend(parse_excel_wip(f, f.name))
        
        existing_ord_nos = {str(r.get('Ord.No.', '')) for r in existing_data if r.get('Ord.No.')}
        unique_new = [n for n in new_data_raw if str(n['Ord.No.']) not in existing_ord_nos]
        full_data = existing_data + unique_new
        full_data = [r for r in full_data if not str(r.get('Chassis/SL No.', '')).startswith('B')]
        
        paycode_ids = process_paycode_report(paycode_file)

        ws.delete_rows(2, ws.max_row)

        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        for r_idx, row_data in enumerate(full_data, start=2):
            ord_no = str(row_data.get('Ord.No.', ''))
            inv_no = row_data.get('Invoice no.', '')
            has_inv = inv_no is not None and str(inv_no).strip() != "" and str(inv_no).strip().lower() != 'nan'
            
            cat_val = str(row_data.get('Category', '') or '').strip()
            status = "Open"

            if has_inv:
                cat_val = "JOB CARD CLOSED"
                status = "Closed"
            elif ord_no in paycode_ids and "JOB CARD CLOSED" not in cat_val.upper():
                cat_val = "Cancelled"
                status = "Closed"
            elif "CANCEL" in cat_val.upper() or "CLOSED" in cat_val.upper():
                status = "Closed"

            for name, c_idx in col_map_ws.items():
                cell = ws.cell(row=r_idx, column=c_idx+1)
                if name == "SNO.": cell.value = r_idx - 1
                elif name == "Category": cell.value = cat_val
                elif name == "Status": cell.value = status
                elif name == "Remarks":
                    cell.value = apply_rich_remarks(str(row_data.get('Remarks', '') or ''))
                else:
                    cell.value = row_data.get(name, "")

        last_row = len(full_data) + 1
        c_idx_cat = col_map_ws.get('Category', 19) + 1
        c_idx_stat = col_map_ws.get('Status', 22) + 1
        cat_let = openpyxl.utils.get_column_letter(c_idx_cat)
        stat_let = openpyxl.utils.get_column_letter(c_idx_stat)

        ws.conditional_formatting.add(f'{cat_let}2:{cat_let}{last_row + 500}', 
            FormulaRule(formula=[f'OR({cat_let}2="JOB CARD CLOSED", {cat_let}2="Cancelled", {cat_let}2="CANCELLED")'], fill=green_fill, stopIfTrue=True))
        ws.conditional_formatting.add(f'{cat_let}2:{cat_let}{last_row + 500}', 
            FormulaRule(formula=[f'AND({cat_let}2<>"", {cat_let}2<>"JOB CARD CLOSED", {cat_let}2<>"Cancelled", {cat_let}2<>"CANCELLED")'], fill=blue_fill))

        red_text = Font(color="9C0006", bold=True)
        green_text = Font(color="006100", bold=True)
        ws.conditional_formatting.add(f'{stat_let}2:{stat_let}{last_row + 500}', 
            FormulaRule(formula=[f'{stat_let}2="Closed"'], fill=red_fill, font=red_text))
        ws.conditional_formatting.add(f'{stat_let}2:{stat_let}{last_row + 500}', 
            FormulaRule(formula=[f'{stat_let}2="Open"'], fill=green_fill, font=green_text))

        output = io.BytesIO()
        wb.save(output)
        st.success("WIP Summary generated! Dealer Name now set to Area.")
        st.download_button("📥 Download Updated Summary", output.getvalue(), file_name=summary_file.name)
