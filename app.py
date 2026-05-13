import streamlit as st
import pandas as pd
import io
import openpyxl
import re
from openpyxl.styles import Font, PatternFill
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from datetime import date

st.set_page_config(page_title="WIP Summary Consolidator Pro v18", layout="wide")

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
    for i, row in df_raw.head(20).iterrows():
        if "Ord.No." in row.values:
            header_idx = i
            break
    
    if header_idx == -1: return []

    df_clean = df_raw.iloc[header_idx+1:].copy()
    df_clean.columns = df_raw.iloc[header_idx].values

    for _, row in df_clean.iterrows():
        if pd.isna(row.get('Ord.No.')): continue
        
        cr_dt = row.get('Cr.Dt')
        pending_days = ""
        if pd.notnull(cr_dt):
            try:
                cr_dt_parsed = pd.to_datetime(cr_dt).date() if isinstance(cr_dt, str) else cr_dt.date()
                pending_days = (today - cr_dt_parsed).days
            except: pass

        data.append({
            "SNO.": "",
            "Pending Days": pending_days,
            "D.Code &JC NO.": f"{str(row.get('Dname', ''))}{str(row.get('Ord.No.', ''))}",
            "Dname": row.get('Dname', ''),
            "Area": area,
            "Dealer Name": row.get('Dealer Name', ''),
            "Req. Delv. Dt": row.get('Req. Delv. Dt', ''),
            "Cr.Dt": cr_dt,
            "Total": row.get('Total', 0),
            "PartsTotal": row.get('PartsTotal', 0),
            "Ord.No.": row.get('Ord.No.', ''),
            "Ord.Ty.": row.get('Ord.Ty.', ''),
            "Regn.No": row.get('Regn.No', ''),
            "Chassis/SL No.": row.get('Chassis/SL No.', ''),
            "Notes": row.get('Notes', ''),
            "Cust.No": row.get('Cust.No', ''),
            "Cust Name": row.get('Cust Name', ''),
            "Remarks": "",
            "Category": "",
            "Invoice no.": "",
            "Date": "",
            "Status": "Open"
        })
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
    prefix = parts[0]
    rt.append(TextBlock(red_bold, prefix))
    
    if len(parts) > 1:
        suffix = "-" + parts[1]
        if "Job completed" in suffix:
            sub_parts = suffix.split("Job completed")
            for i, p in enumerate(sub_parts):
                rt.append(TextBlock(normal, p))
                if i < len(sub_parts) - 1:
                    rt.append(TextBlock(red_bold, "Job completed"))
        else:
            rt.append(TextBlock(normal, suffix))
    return rt

st.title("📊 WIP Summary Consolidator Pro v18")

with st.sidebar:
    st.header("Files")
    summary_file = st.file_uploader("1. Existing Summary File", type=["xlsx"])
    open_order_files = st.file_uploader("2. New Open Order Lists", type=["xlsx"], accept_multiple_files=True)
    paycode_file = st.file_uploader("3. Paycodewise Report", type=["xlsx"])

if st.button("Generate & Format Summary"):
    if not summary_file:
        st.warning("Please upload the Summary file.")
    else:
        wb = openpyxl.load_workbook(summary_file, keep_links=True)
        target_name = "Master Data" if "Master Data" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[target_name]

        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        col_map = {name: i for i, name in enumerate(header)}
        existing_data = [{header[i]: val for i, val in enumerate(row)} for row in rows[1:]]

        new_data = []
        for f in open_order_files:
            new_data.extend(parse_excel_wip(f, f.name))
        
        existing_ids = {str(r.get('Ord.No.', '')) for r in existing_data if r.get('Ord.No.')}
        unique_new = [n for n in new_data if str(n['Ord.No.']) not in existing_ids]
        full_data = existing_data + unique_new
        full_data = [r for r in full_data if not str(r.get('Chassis/SL No.', '')).startswith('B')]
        
        paycode_ids = process_paycode_report(paycode_file)

        ws.delete_rows(2, ws.max_row)

        # Updated Color Palette
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        red_status_text = Font(color="9C0006", bold=True)
        green_status_text = Font(color="006100", bold=True)
        none_fill = PatternFill(fill_type=None)

        for r_idx, row_data in enumerate(full_data, start=2):
            ord_no = str(row_data.get('Ord.No.', ''))
            inv_val_raw = row_data.get('Invoice no.', '')
            has_inv = inv_val_raw is not None and str(inv_val_raw).strip() != "" and str(inv_val_raw).strip().lower() != 'nan'
            
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

            for name, c_idx in col_map.items():
                cell = ws.cell(row=r_idx, column=c_idx+1)
                if name == "SNO.": cell.value = r_idx - 1
                elif name == "Category":
                    cell.value = cat_val
                    # Formatting: Green for Closed/Cancelled, Blue for the rest
                    if cat_val in ["JOB CARD CLOSED", "Cancelled"]:
                        cell.fill = green_fill
                        cell.font = Font() # Standard font
                    elif cat_val != "":
                        cell.fill = blue_fill
                        cell.font = Font() # Standard font
                    else:
                        cell.fill = none_fill
                        cell.font = Font()
                elif name == "Status":
                    cell.value = status
                    cell.fill = red_fill if status == "Closed" else green_fill
                    cell.font = red_status_text if status == "Closed" else green_status_text
                elif name == "Remarks":
                    cell.value = apply_rich_remarks(str(row_data.get('Remarks', '') or ''))
                else:
                    cell.value = row_data.get(name, "")

        output = io.BytesIO()
        wb.save(output)
        st.success("WIP Summary generated with updated color logic!")
        st.download_button("📥 Download Final Summary", output.getvalue(), file_name=summary_file.name)