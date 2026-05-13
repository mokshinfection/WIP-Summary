import streamlit as st
import pandas as pd
import io
import openpyxl
import re
import copy
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from datetime import date

st.set_page_config(page_title="WIP Summary Consolidator Pro v10", layout="wide")

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

st.title("📊 WIP Summary Consolidator Pro v10")
st.info("Updates only the Master Data sheet while leaving all other sheets (Pivots, Formulas, Charts) untouched.")

with st.sidebar:
    st.header("Files")
    summary_file = st.file_uploader("1. Summary File to Update", type=["xlsx"])
    open_order_files = st.file_uploader("2. New Open Order Lists", type=["xlsx"], accept_multiple_files=True)
    paycode_file = st.file_uploader("3. Paycodewise Report", type=["xlsx"])

if st.button("Update Summary Data"):
    if not summary_file:
        st.warning("Please upload the Summary file.")
    else:
        # Load workbook with openpyxl directly to keep EVERYTHING
        wb = openpyxl.load_workbook(summary_file)
        target_sheet_name = "Master Data" if "Master Data" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[target_sheet_name]

        # Extract current data from Master sheet for deduplication
        data = ws.values
        cols = next(data)
        master_df = pd.DataFrame(data, columns=cols)

        # Parse new data
        new_data_list = []
        for f in open_order_files:
            new_data_list.extend(parse_excel_wip(f, f.name))
        new_df = pd.DataFrame(new_data_list)

        # Deduplicate
        if not master_df.empty and not new_df.empty:
            existing_orders = set(master_df['Ord.No.'].astype(str).unique())
            new_df = new_df[~new_df['Ord.No.'].astype(str).isin(existing_orders)]

        # Combine
        combined_df = pd.concat([master_df, new_df], ignore_index=True)
        
        if not combined_df.empty:
            # Filter Chassis
            if 'Chassis/SL No.' in combined_df.columns:
                combined_df = combined_df[~combined_df['Chassis/SL No.'].astype(str).str.startswith('B', na=False)]

            paycode_ids = process_paycode_report(paycode_file) if paycode_file else set()
            
            # Instead of overwriting the whole file, we overwrite ONLY the cells in the Master sheet
            # Clear current Master sheet (except header)
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.value = None

            # Styles
            red_font = Font(color="FF0000", bold=True)
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            green_font = Font(color="006100", bold=True)

            col_map = {val: i+1 for i, val in enumerate(combined_df.columns)}
            
            # Re-write the combined data back to the sheet
            for r_idx, row_data in enumerate(combined_df.to_dict('records'), start=2):
                for col_name, value in row_data.items():
                    c_idx = col_map[col_name]
                    cell = ws.cell(row=r_idx, column=c_idx)
                    
                    # 1. Formatting for Remarks (Red check)
                    if col_name == "Remarks" and str(value).lower() == "checked": # or other logic
                         pass # existing logic per cell would need to be re-applied or held

                    # 2. Re-apply Category Formula
                    if col_name == "Category":
                        inv_letter = openpyxl.utils.get_column_letter(col_map["Invoice no."])
                        cell.value = f'=IF(ISBLANK({inv_letter}{r_idx}), "", "JOB CARD CLOSED")'
                    else:
                        cell.value = value

                    # 3. Status logic
                    if col_name == "Status":
                        ord_val = str(row_data.get("Ord.No.", ""))
                        inv_val = row_data.get("Invoice no.", "")
                        cat_val = str(row_data.get("Category", "")).upper()
                        
                        is_closed = (ord_val in paycode_ids) or (inv_val and str(inv_val).strip() != "") or ("CLOSED" in cat_val)
                        
                        if is_closed:
                            cell.value = "Closed"
                            cell.fill = red_fill
                            cell.font = Font(color="9C0006", bold=True)
                        else:
                            cell.value = "Open"
                            cell.fill = green_fill
                            cell.font = green_font

            # Finalize
            output = io.BytesIO()
            wb.save(output)
            
            st.success("Successfully updated the Master Data while preserving all other sheets and formulas!")
            st.download_button(label="📥 Download Updated Summary", data=output.getvalue(), 
                               file_name=summary_file.name)
