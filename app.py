import streamlit as st
import pandas as pd
import io
import openpyxl
import re
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from datetime import date

st.set_page_config(page_title="WIP Summary Consolidator Pro v9", layout="wide")

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

st.title("📊 WIP Summary Consolidator Pro v9")
st.info("Upload your existing Summary file to update it with new data while preserving all formatting and sheets.")

with st.sidebar:
    st.header("Files")
    summary_file = st.file_uploader("1. Summary File to Update", type=["xlsx"])
    open_order_files = st.file_uploader("2. New Open Order Lists", type=["xlsx"], accept_multiple_files=True)
    paycode_file = st.file_uploader("3. Paycodewise Report", type=["xlsx"])

if st.button("Update Summary Data"):
    if not summary_file:
        st.warning("Please upload the Summary file you want to edit.")
    else:
        # Load Existing Summary exactly as it is
        wb_orig = openpyxl.load_workbook(summary_file)
        xls = pd.ExcelFile(summary_file)
        existing_sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
        target_sheet_name = "Master Data" if "Master Data" in existing_sheets else xls.sheet_names[0]
        master_df = existing_sheets[target_sheet_name]

        # Parse new data
        new_data = []
        for f in open_order_files:
            new_data.extend(parse_excel_wip(f, f.name))
        new_df = pd.DataFrame(new_data)

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
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Write updated data to the target sheet
                combined_df.to_excel(writer, index=False, sheet_name=target_sheet_name)
                
                # Copy all other sheets over exactly
                for sheet_name in wb_orig.sheet_names:
                    if sheet_name != target_sheet_name:
                        existing_sheets[sheet_name].to_excel(writer, index=False, sheet_name=sheet_name)

                ws = writer.book[target_sheet_name]
                ws_orig = wb_orig[target_sheet_name]

                # Formatting and Logic
                red_font = Font(color="FF0000", bold=True)
                red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                green_font = Font(color="006100", bold=True)

                cols = {val: i+1 for i, val in enumerate(combined_df.columns)}
                cat_col = cols.get('Category')
                inv_col = cols.get('Invoice no.')
                stat_col = cols.get('Status')
                rem_col = cols.get('Remarks')
                ord_col = cols.get('Ord.No.')

                inv_letter = openpyxl.utils.get_column_letter(inv_col)
                cat_letter = openpyxl.utils.get_column_letter(cat_col)

                for r in range(2, len(combined_df) + 2):
                    # Keep Red Remarks
                    if r <= ws_orig.max_row:
                        orig_cell = ws_orig.cell(row=r, column=rem_col)
                        if orig_cell.font and orig_cell.font.color and orig_cell.font.color.rgb == "FFFF0000":
                            ws.cell(row=r, column=rem_col).font = red_font

                    # Fix Category
                    ws.cell(row=r, column=cat_col).value = f'=IF(ISBLANK({inv_letter}{r}), \"\", \"JOB CARD CLOSED\")'
                    ws.conditional_formatting.add(f'{cat_letter}{r}',
                        FormulaRule(formula=[f'NOT(ISBLANK({inv_letter}{r}))'], fill=green_fill, font=green_font))

                    # Status Logic
                    ord_val = str(ws.cell(row=r, column=ord_col).value)
                    # We evaluate formula-like logic here for the Status value
                    inv_val = ws.cell(row=r, column=inv_col).value
                    is_closed = (ord_val in paycode_ids) or (inv_val is not None and str(inv_val).strip() != "")
                    
                    status_cell = ws.cell(row=r, column=stat_col)
                    if is_closed:
                        status_cell.value = "Closed"
                        status_cell.fill = red_fill
                        status_cell.font = Font(color="9C0006", bold=True)
                    else:
                        status_cell.value = "Open"
                        status_cell.fill = green_fill
                        status_cell.font = green_font

            st.success("Update Complete!")
            st.download_button(label="📥 Download Updated Summary", data=output.getvalue(), 
                               file_name=summary_file.name) # Use the same filename
