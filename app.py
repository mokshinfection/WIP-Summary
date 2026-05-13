import streamlit as st
import pandas as pd
import io
import openpyxl
import re
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule
from datetime import date

st.set_page_config(page_title="WIP Summary Consolidator Pro v7", layout="wide")

def extract_area_from_filename(filename):
    """Extracts the area name from filenames like 'Open Orders List Nellore 30th Apr 2026.xlsx'"""
    match = re.search(r"Open Orders List\s+(.*?)\s+\d+", filename, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    parts = filename.replace('.xlsx', '').split(' ')
    if len(parts) >= 4:
        return parts[3]
    return "Unknown"

def parse_excel_wip(file_obj, filename):
    """Parses Open Order Excel files."""
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
    
    if header_idx == -1:
        st.warning(f"Could not find 'Ord.No.' header in {filename}. Skipping.")
        return []

    df_clean = df_raw.iloc[header_idx+1:].copy()
    df_clean.columns = df_raw.iloc[header_idx].values

    for _, row in df_clean.iterrows():
        if pd.isna(row.get('Ord.No.')):
            continue
            
        cr_dt = row.get('Cr.Dt')
        pending_days = ""
        if pd.notnull(cr_dt):
            try:
                if isinstance(cr_dt, str):
                    cr_dt_parsed = pd.to_datetime(cr_dt).date()
                else:
                    cr_dt_parsed = cr_dt.date() if hasattr(cr_dt, 'date') else cr_dt
                pending_days = (today - cr_dt_parsed).days
            except:
                pass

        d_code = str(row.get('Dname', ''))
        ord_no = str(row.get('Ord.No.', ''))
        d_code_jc = f"{d_code}{ord_no}"

        record = {
            "SNO.": "",
            "Pending Days": pending_days,
            "D.Code &JC NO.": d_code_jc,
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
    """Extracts a set of Order IDs from the Paycodewise report."""
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
    except Exception as e:
        st.error(f"Error processing Paycode report: {e}")
    return set()

st.title("📊 WIP Summary Consolidator Pro v7")

with st.sidebar:
    st.header("Upload Files")
    summary_file = st.file_uploader("1. Existing Summary File", type=["xlsx"])
    open_order_files = st.file_uploader("2. New Open Order Lists", type=["xlsx"], accept_multiple_files=True)
    paycode_file = st.file_uploader("3. Paycodewise Report", type=["xlsx"])

if st.button("Generate Consolidated WIP Report"):
    if not open_order_files and not summary_file:
        st.warning("Please upload at least one file.")
    else:
        existing_sheets = {}
        master_df = pd.DataFrame()
        target_sheet_name = "Master Data"
        
        if summary_file:
            xls = pd.ExcelFile(summary_file)
            existing_sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
            if "Master Data" in existing_sheets:
                master_df = existing_sheets["Master Data"]
                target_sheet_name = "Master Data"
            else:
                target_sheet_name = xls.sheet_names[0]
                master_df = existing_sheets[target_sheet_name]

        new_data = []
        for f in open_order_files:
            new_data.extend(parse_excel_wip(f, f.name))
        new_df = pd.DataFrame(new_data)

        if not master_df.empty and not new_df.empty:
            existing_orders = set(master_df['Ord.No.'].astype(str).unique())
            new_df = new_df[~new_df['Ord.No.'].astype(str).isin(existing_orders)]

        combined_df = pd.concat([master_df, new_df], ignore_index=True)
        
        if not combined_df.empty:
            if 'Chassis/SL No.' in combined_df.columns:
                combined_df = combined_df[~combined_df['Chassis/SL No.'].astype(str).str.startswith('B', na=False)]

            paycode_ids = set()
            if paycode_file:
                paycode_ids = process_paycode_report(paycode_file)

            if 'Status' not in combined_df.columns:
                combined_df['Status'] = 'Open'
            
            # Recalculate Status but DO NOT overwrite existing Category text unless empty
            def determine_status(row):
                if str(row.get('Ord.No.', '')) in paycode_ids:
                    return "Closed"
                cat = str(row.get('Category', '')).upper()
                if "JOB CARD CLOSED" in cat or "CANCELLED" in cat:
                    return "Closed"
                return "Open"

            combined_df['Status'] = combined_df.apply(determine_status, axis=1)
            combined_df['SNO.'] = range(1, len(combined_df) + 1)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, sheet_name=target_sheet_name)
                for sheet_name, sheet_df in existing_sheets.items():
                    if sheet_name != target_sheet_name:
                        sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

                ws = writer.book[target_sheet_name]

                # Styles
                header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                red_font = Font(color="9C0006", bold=True)
                green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                green_font = Font(color="006100", bold=True)

                for col in range(1, len(combined_df.columns) + 1):
                    cell = ws.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 20

                cols = {val: i+1 for i, val in enumerate(combined_df.columns)}
                cat_col = cols.get('Category')
                inv_col = cols.get('Invoice no.')
                stat_col = cols.get('Status')
                rem_col = cols.get('Remarks')

                cat_letter = openpyxl.utils.get_column_letter(cat_col)
                inv_letter = openpyxl.utils.get_column_letter(inv_col)

                for r in range(2, len(combined_df) + 2):
                    # 1. FIX: Ensure Category isn't empty by re-applying the formula logic
                    # This ensures the column shows "JOB CARD CLOSED" if Invoice No exists
                    ws.cell(row=r, column=cat_col).value = f'=IF(ISBLANK({inv_letter}{r}), "", "JOB CARD CLOSED")'
                    
                    # 2. Formatting for Category based on Invoice
                    ws.conditional_formatting.add(f'{cat_letter}{r}',
                        FormulaRule(formula=[f'NOT(ISBLANK({inv_letter}{r}))'], fill=green_fill, font=green_font))

                    # 3. Status Formatting
                    status_cell = ws.cell(row=r, column=stat_col)
                    if status_cell.value == "Closed":
                        status_cell.fill = red_fill
                        status_cell.font = red_font
                    else:
                        status_cell.fill = green_fill
                        status_cell.font = green_font

                ws.auto_filter.ref = ws.dimensions

            st.success("Consolidation Complete!")
            st.download_button(
                label="📥 Download Updated WIP Summary",
                data=output.getvalue(),
                file_name=f"Updated_WIP_Summary_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
