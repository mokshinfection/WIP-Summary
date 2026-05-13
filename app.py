import streamlit as st
import pandas as pd
import io
import openpyxl
import re
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule
from datetime import date

st.set_page_config(page_title="WIP Summary Consolidator Pro v6", layout="wide")

def extract_area_from_filename(filename):
    """
    Extracts the area name from filenames like 'Open Orders List Nellore 30th Apr 2026.xlsx'
    """
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

    # Detect header row (looking for 'Ord.No.')
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
        # Basic check to avoid empty rows
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

        # D.Code & JC No construction
        d_code = str(row.get('Dname', ''))
        ord_no = str(row.get('Ord.No.', ''))
        d_code_jc = f"{d_code}{ord_no}"

        record = {
            "SNO.": "", # Will re-index later
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
            "Status": "Open" # Default
        }
        data.append(record)
    
    return data

def process_paycode_report(file_obj):
    """Extracts a set of Order IDs from the Paycodewise report."""
    try:
        df_raw = pd.read_excel(file_obj, header=None)
        # Find the row containing 'Ord.ID'
        header_row = -1
        for i, row in df_raw.head(25).iterrows():
            if "Ord.ID" in row.values:
                header_row = i
                break
        
        if header_row != -1:
            df = df_raw.iloc[header_row+1:].copy()
            df.columns = df_raw.iloc[header_row].values
            # Return set of valid Order IDs
            return set(df['Ord.ID'].dropna().astype(str).unique())
    except Exception as e:
        st.error(f"Error processing Paycode report: {e}")
    return set()

st.title("📊 WIP Summary Consolidator Pro v6")
st.markdown("Consolidate Open Order lists into an existing Summary file with cross-referencing.")

with st.sidebar:
    st.header("Upload Files")
    summary_file = st.file_uploader("1. Existing Summary File (Optional)", type=["xlsx"])
    open_order_files = st.file_uploader("2. New Open Order Lists", type=["xlsx"], accept_multiple_files=True)
    paycode_file = st.file_uploader("3. Paycodewise Report (Optional)", type=["xlsx"])

if st.button("Generate Consolidated WIP Report"):
    if not open_order_files and not summary_file:
        st.warning("Please upload at least one file.")
    else:
        # 1. Load Existing Summary
        existing_sheets = {}
        master_df = pd.DataFrame()
        target_sheet_name = "Master Data" # Default name
        
        if summary_file:
            xls = pd.ExcelFile(summary_file)
            existing_sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
            # Identify master sheet (user said "Master Data" usually)
            if "Master Data" in existing_sheets:
                master_df = existing_sheets["Master Data"]
                target_sheet_name = "Master Data"
            else:
                target_sheet_name = xls.sheet_names[0]
                master_df = existing_sheets[target_sheet_name]
            st.info(f"Loaded existing summary from sheet: '{target_sheet_name}'")

        # 2. Process New Files
        new_data = []
        for f in open_order_files:
            new_data.extend(parse_excel_wip(f, f.name))
        
        new_df = pd.DataFrame(new_data)

        # 3. Filter New Records (Deduplication)
        if not master_df.empty and not new_df.empty:
            existing_orders = set(master_df['Ord.No.'].astype(str).unique())
            new_df = new_df[~new_df['Ord.No.'].astype(str).isin(existing_orders)]
            st.success(f"Found {len(new_df)} new unique orders to add.")

        # Combine
        combined_df = pd.concat([master_df, new_df], ignore_index=True)
        
        if not combined_df.empty:
            # 4. Remove Chassis starting with 'B'
            if 'Chassis/SL No.' in combined_df.columns:
                initial_count = len(combined_df)
                combined_df = combined_df[~combined_df['Chassis/SL No.'].astype(str).str.startswith('B', na=False)]
                removed = initial_count - len(combined_df)
                if removed > 0:
                    st.info(f"Removed {removed} rows with Chassis starting with 'B'.")

            # 5. Cross-reference Paycodewise Report
            paycode_ids = set()
            if paycode_file:
                paycode_ids = process_paycode_report(paycode_file)
                st.info(f"Loaded {len(paycode_ids)} Order IDs from Paycodewise Report.")

            # 6. Status Column Logic
            # Ensure 'Status' exists
            if 'Status' not in combined_df.columns:
                combined_df['Status'] = 'Open'
            
            # Update Status based on Category and Paycode
            def determine_status(row):
                # If in Paycode report -> Closed
                if str(row.get('Ord.No.', '')) in paycode_ids:
                    return "Closed"
                # If Category is Closed/Cancelled -> Closed
                cat = str(row.get('Category', '')).upper()
                if "JOB CARD CLOSED" in cat or "CANCELLED" in cat:
                    return "Closed"
                return "Open"

            combined_df['Status'] = combined_df.apply(determine_status, axis=1)

            # Re-generate SNO
            combined_df['SNO.'] = range(1, len(combined_df) + 1)

            # Prepare to save
            output = io.BytesIO()
            
            # Use ExcelWriter to preserve other sheets
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Write the updated master sheet
                combined_df.to_excel(writer, index=False, sheet_name=target_sheet_name)
                
                # Write other sheets back
                for sheet_name, sheet_df in existing_sheets.items():
                    if sheet_name != target_sheet_name:
                        sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

                # Access workbook for formatting
                workbook = writer.book
                ws = workbook[target_sheet_name]

                # Formatting Styles
                header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                
                red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                red_font = Font(color="9C0006", bold=True)
                
                green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                green_font = Font(color="006100", bold=True)

                # Header formatting
                for col in range(1, len(combined_df.columns) + 1):
                    cell = ws.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 20

                # Column Indices (1-based for openpyxl)
                cols = {val: i+1 for i, val in enumerate(combined_df.columns)}
                
                cat_col_letter = openpyxl.utils.get_column_letter(cols.get('Category', 1))
                inv_col_letter = openpyxl.utils.get_column_letter(cols.get('Invoice no.', 1))
                stat_col_letter = openpyxl.utils.get_column_letter(cols.get('Status', 1))

                # Apply logic and formatting
                for r in range(2, len(combined_df) + 2):
                    # Category Formula (Old requirement)
                    if 'Category' in cols and 'Invoice no.' in cols:
                        ws[f'{cat_col_letter}{r}'] = f'=IF(ISBLANK({inv_col_letter}{r}), "", "JOB CARD CLOSED")'
                    
                    # Status Formatting
                    status_cell = ws[f'{stat_col_letter}{r}']
                    if status_cell.value == "Closed":
                        status_cell.fill = red_fill
                        status_cell.font = red_font
                    elif status_cell.value == "Open":
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
            st.dataframe(combined_df.head(20))
