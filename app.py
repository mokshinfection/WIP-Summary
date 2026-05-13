import streamlit as st
import pandas as pd
import io
import openpyxl
import re
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from datetime import date

st.set_page_config(page_title="WIP Summary Consolidator Pro v12", layout="wide")

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

st.title("📊 WIP Summary Consolidator Pro v12")
st.info("Directly updates 'Master Data' while preserving all other sheets (Pivots, Formulas, Charts) and their formatting.")

with st.sidebar:
    st.header("Files")
    summary_file = st.file_uploader("1. Summary File to Update", type=["xlsx"])
    open_order_files = st.file_uploader("2. New Open Order Lists", type=["xlsx"], accept_multiple_files=True)
    paycode_file = st.file_uploader("3. Paycodewise Report", type=["xlsx"])

if st.button("Edit & Update Summary"):
    if not summary_file:
        st.warning("Please upload the Summary file.")
    else:
        # Load workbook with openpyxl (this keeps all sheets, charts, and formulas intact)
        wb = openpyxl.load_workbook(summary_file)
        target_sheet_name = "Master Data" if "Master Data" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[target_sheet_name]

        # 1. Read current Master Data for deduplication
        data_gen = ws.values
        header = next(data_gen)
        col_map = {name: i for i, name in enumerate(header)}
        
        existing_data = []
        # Re-read with cell objects to capture formatting for Remarks
        for row in ws.iter_rows(min_row=2):
            row_dict = {}
            for name, idx in col_map.items():
                cell = row[idx]
                row_dict[name] = cell.value
                if name == "Remarks":
                    row_dict["_remarks_color"] = cell.font.color.rgb if cell.font and cell.font.color else None
            existing_data.append(row_dict)

        # 2. Parse new data
        new_data_raw = []
        for f in open_order_files:
            new_data_raw.extend(parse_excel_wip(f, f.name))
        
        # 3. Deduplicate (only add what isn't there)
        existing_ord_nos = {str(r.get('Ord.No.', '')) for r in existing_data if r.get('Ord.No.')}
        unique_new = [n for n in new_data_raw if str(n['Ord.No.']) not in existing_ord_nos]
        
        # Combine existing and new
        full_data = existing_data + unique_new
        
        # Filter Chassis 'B'
        full_data = [r for r in full_data if not str(r.get('Chassis/SL No.', '')).startswith('B')]

        # 4. Paycode Cross-ref
        paycode_ids = process_paycode_report(paycode_file) if paycode_file else set()

        # 5. SURGICAL EDIT: Clear data rows only, then repopulate
        # This keeps the header and all other sheets exactly as they are.
        ws.delete_rows(2, ws.max_row)

        # Styles
        red_font = Font(color="FFFF0000", bold=True)
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        green_font = Font(color="006100", bold=True)

        inv_col_idx = col_map.get('Invoice no.', 20) + 1
        cat_col_idx = col_map.get('Category', 19) + 1
        stat_col_idx = col_map.get('Status', 22) + 1
        rem_col_idx = col_map.get('Remarks', 18) + 1
        ord_col_idx = col_map.get('Ord.No.', 10) + 1
        
        inv_letter = openpyxl.utils.get_column_letter(inv_col_idx)
        cat_letter = openpyxl.utils.get_column_letter(cat_col_idx)

        for r_idx, row_data in enumerate(full_data, start=2):
            for name, idx in col_map.items():
                cell = ws.cell(row=r_idx, column=idx + 1)
                
                if name == "SNO.":
                    cell.value = r_idx - 1
                elif name == "Category":
                    cell.value = f'=IF(ISBLANK({inv_letter}{r_idx}), "", "JOB CARD CLOSED")'
                elif name == "Status":
                    pass # logic below
                else:
                    cell.value = row_data.get(name, "")

                # Keep the red color for remarks if it was originally red
                if name == "Remarks" and row_data.get("_remarks_color") == "FFFF0000":
                    cell.font = red_font

            # Live Status Calculation
            ord_val = str(ws.cell(row=r_idx, column=ord_col_idx).value)
            inv_val = ws.cell(row=r_idx, column=inv_col_idx).value
            
            is_closed = (ord_val in paycode_ids) or (inv_val is not None and str(inv_val).strip() != "")
            
            status_cell = ws.cell(row=r_idx, column=stat_col_idx)
            if is_closed:
                status_cell.value = "Closed"
                status_cell.fill = red_fill
                status_cell.font = Font(color="9C0006", bold=True)
            else:
                status_cell.value = "Open"
                status_cell.fill = green_fill
                status_cell.font = green_font
        
        # Apply conditional formatting to Category column
        ws.conditional_formatting.add(f'{cat_letter}2:{cat_letter}{len(full_data)+1}',
            FormulaRule(formula=[f'NOT(ISBLANK({inv_letter}2))'], fill=green_fill, font=green_font))

        # Save to buffer
        output = io.BytesIO()
        wb.save(output)
        
        st.success("Master Data updated. All other sheets (Pivots, Charts, etc.) remain untouched.")
        st.download_button(label="📥 Download Updated Summary", data=output.getvalue(), 
                           file_name=summary_file.name)
