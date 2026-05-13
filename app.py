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

st.set_page_config(page_title="WIP Summary Consolidator Pro v13", layout="wide")

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

        d_code = str(row.get('Dname', ''))
        ord_no = str(row.get('Ord.No.', ''))

        record = {
            "SNO.": "",
            "Pending Days": pending_days,
            "D.Code &JC NO.": f"{d_code}{ord_no}",
            "Dname": d_code,
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
            "Closing Date": "",
            "Remarks": "",
            "Category": "",
            "Invoice no.": "",
            "Date": "",
            "Status": "Open"
        }
        data.append(record)
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

def apply_remarks_formatting(text):
    """Applies rich text formatting based on user requirements."""
    if not text or not isinstance(text, str):
        return text
    
    red_bold = InlineFont(color="FFFF0000", b=True)
    normal = InlineFont()
    
    # Logic: 
    # 1. Everything before '-' is red/bold.
    # 2. 'Job completed' is red/bold.
    
    # Split by hyphen
    parts = text.split('-', 1)
    rt = CellRichText()
    
    if len(parts) > 1:
        prefix = parts[0]
        suffix = "-" + parts[1]
        
        # Check if 'Job completed' is in prefix (it already gets red/bold)
        rt.append(TextBlock(red_bold, prefix))
        
        # Check suffix for 'Job completed'
        if "Job completed" in suffix:
            sub_parts = suffix.split("Job completed")
            for i, part in enumerate(sub_parts):
                rt.append(TextBlock(normal, part))
                if i < len(sub_parts) - 1:
                    rt.append(TextBlock(red_bold, "Job completed"))
        else:
            rt.append(TextBlock(normal, suffix))
    else:
        # No hyphen - check for 'Job completed' only
        if "Job completed" in text:
            sub_parts = text.split("Job completed")
            for i, part in enumerate(sub_parts):
                rt.append(TextBlock(normal, part))
                if i < len(sub_parts) - 1:
                    rt.append(TextBlock(red_bold, "Job completed"))
        else:
            # Just plain text
            return text
            
    return rt

st.title("📊 WIP Summary Consolidator Pro v13")

with st.sidebar:
    st.header("Files")
    summary_file = st.file_uploader("1. Summary File to Update", type=["xlsx"])
    open_order_files = st.file_uploader("2. New Open Order Lists", type=["xlsx"], accept_multiple_files=True)
    paycode_file = st.file_uploader("3. Paycodewise Report", type=["xlsx"])

if st.button("Edit & Update Summary"):
    if not summary_file:
        st.warning("Please upload the Summary file.")
    else:
        wb = openpyxl.load_workbook(summary_file)
        target_sheet_name = "Master Data" if "Master Data" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[target_sheet_name]

        # Read current Data
        data_gen = ws.values
        header = next(data_gen)
        col_map = {name: i for i, name in enumerate(header)}
        
        existing_data = []
        for row in ws.iter_rows(min_row=2):
            row_dict = {}
            for name, idx in col_map.items():
                cell = row[idx]
                row_dict[name] = cell.value
            existing_data.append(row_dict)

        # Parse new data
        new_data_raw = []
        for f in open_order_files:
            new_data_raw.extend(parse_excel_wip(f, f.name))
        
        existing_ord_nos = {str(r.get('Ord.No.', '')) for r in existing_data if r.get('Ord.No.')}
        unique_new = [n for n in new_data_raw if str(n['Ord.No.']) not in existing_ord_nos]
        full_data = existing_data + unique_new
        full_data = [r for r in full_data if not str(r.get('Chassis/SL No.', '')).startswith('B')]

        paycode_ids = process_paycode_report(paycode_file) if paycode_file else set()

        # Surgical Edit
        ws.delete_rows(2, ws.max_row)

        # Styles
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        green_font_style = Font(color="006100", bold=True)
        red_font_style = Font(color="9C0006", bold=True)

        idx_inv = col_map.get('Invoice no.', 20)
        idx_cat = col_map.get('Category', 19)
        idx_stat = col_map.get('Status', 22)
        idx_rem = col_map.get('Remarks', 18)
        idx_ord = col_map.get('Ord.No.', 10)
        
        inv_letter = openpyxl.utils.get_column_letter(idx_inv + 1)
        cat_letter = openpyxl.utils.get_column_letter(idx_cat + 1)

        for r_idx, row_data in enumerate(full_data, start=2):
            ord_val = str(row_data.get('Ord.No.', ''))
            inv_val = row_data.get('Invoice no.', '')
            has_inv = inv_val is not None and str(inv_val).strip() != ""
            in_paycode = ord_val in paycode_ids

            # Status and Category Logic
            category = row_data.get('Category', '')
            status = "Open"

            if in_paycode:
                category = "Cancelled"
                status = "Closed"
            elif has_inv:
                category = "JOB CARD CLOSED"
                status = "Closed"
            elif "CANCEL" in str(category).upper() or "CLOSED" in str(category).upper():
                status = "Closed"

            # Write values
            for name, idx in col_map.items():
                cell = ws.cell(row=r_idx, column=idx + 1)
                
                if name == "SNO.":
                    cell.value = r_idx - 1
                elif name == "Category":
                    cell.value = category
                elif name == "Status":
                    cell.value = status
                    if status == "Closed":
                        cell.fill = red_fill
                        cell.font = red_font_style
                    else:
                        cell.fill = green_fill
                        cell.font = green_font_style
                elif name == "Remarks":
                    raw_rem = row_data.get("Remarks", "")
                    cell.value = apply_remarks_formatting(raw_rem)
                else:
                    cell.value = row_data.get(name, "")

            # Apply Green Formatting for Category if it has value (simulating formula check)
            if str(ws.cell(row=r_idx, column=idx_cat+1).value).strip() != "":
                ws.cell(row=r_idx, column=idx_cat+1).fill = green_fill
                ws.cell(row=r_idx, column=idx_cat+1).font = green_font_style

        # Save
        output = io.BytesIO()
        wb.save(output)
        
        st.success("Update Complete! All requirements applied.")
        st.download_button(label="📥 Download Updated Summary", data=output.getvalue(), 
                           file_name=summary_file.name)
