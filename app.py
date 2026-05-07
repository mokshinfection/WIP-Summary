import streamlit as st
import pandas as pd
import io
import openpyxl
import re
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule
from datetime import date

st.set_page_config(page_title="WIP Summary Consolidator Pro v5", layout="wide")

def extract_area_from_filename(filename):
    """
    Extracts the area name from filenames like 'Open Orders List Nellore 30th Apr 2026.xlsx'
    It looks for the word(s) immediately following 'Open Orders List'.
    """
    # Regex to find text between 'Open Orders List' and a date/year or end of string
    match = re.search(r"Open Orders List\s+(.*?)\s+\d+", filename, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    
    # Fallback: If filename doesn't match standard, try splitting
    parts = filename.replace('.xlsx', '').split(' ')
    if len(parts) >= 4:
        return parts[3] # Usually the 4th word in your format
    
    return "Unknown"

def parse_excel_wip(file_obj, filename):
    """Parses Excel with dynamic Area extraction and Identity calculation."""
    df_raw = pd.read_excel(file_obj, header=None)
    data = []
    today = date.today()

    # REQUIREMENT: Pull Area Name dynamically from the file name
    hub_area = extract_area_from_filename(filename)

    # Find the Header Row
    header_idx = -1
    col_map = {}
    for i in range(len(df_raw)):
        row_list = [str(x).strip() for x in df_raw.iloc[i]]
        if 'Pending Days' in row_list:
            header_idx = i
            for j, val in enumerate(row_list):
                if val and val != 'nan':
                    col_map[val] = j
            break
    
    if header_idx == -1: return []

    curr_dname, full_dealer_string = None, None
    i = header_idx + 1
    
    while i < len(df_raw):
        row = df_raw.iloc[i]
        
        # Detect Dealer Header row
        if str(row[0]).strip() == 'Dealer':
            curr_dname = str(df_raw.iloc[i, 6]).split('.')[0]
            # Dealer name as it is in the source file
            full_dealer_string = str(df_raw.iloc[i, 7]) if not pd.isna(df_raw.iloc[i, 7]) else ""
            
        try:
            val_0 = row[0]
            if not pd.isna(val_0) and (isinstance(val_0, (int, float)) or str(val_0).isdigit()):
                def get_v(label):
                    idx = col_map.get(label)
                    return row[idx] if idx is not None else None

                ord_no = str(get_v('Ord.No.')).split('.')[0]
                cr_dt_raw = get_v('Cr.Dt')
                
                # Calculate Pending Days
                try:
                    creation_date = pd.to_datetime(cr_dt_raw).date()
                    calc_pending = (today - creation_date).days
                except:
                    calc_pending = val_0

                record = {
                    'Pending Days': calc_pending,
                    # Column "D.Code & JC No." (Concat Dname and Ord No)
                    'D.Code & JC No.': f"{curr_dname}{ord_no}",
                    'Dname': curr_dname,
                    'Area': hub_area, # Pulled from filename
                    'Dealer Name': full_dealer_string,
                    'Req. Delv. Dt': get_v('Req. Delv. Dt'),
                    'Cr.Dt': cr_dt_raw,
                    'Total': get_v('Total'),
                    'PartsTotal': get_v('PartsTotal'),
                    'Ord.No.': ord_no,
                    'Ord.Ty.': get_v('Ord.Ty.'),
                    'Regn.No': get_v('Regn.No'),
                    'Chassis/SL No.': get_v('Chassis/SL No.'),
                    'Cust.No': get_v('Cust.No'),
                    'Cust Name': get_v('Cust Name'),
                    'Notes': "",
                    'Category': "", # Handled by Excel Logic
                    'Invoice No.': "", 
                    'Date': ""        
                }
                
                # Extract Notes
                if i + 1 < len(df_raw) and 'Notes' in [str(x).strip() for x in df_raw.iloc[i+1]]:
                    row_notes = df_raw.iloc[i+1]
                    for k in range(21, len(df_raw.columns)):
                        if not pd.isna(row_notes[k]) and str(row_notes[k]).strip() not in ["", "Notes"]:
                            record['Notes'] = row_notes[k]
                            break
                    i += 1
                data.append(record)
        except: pass
        i += 1
    return data

# --- Streamlit UI ---
st.title("📊 WIP Summary Consolidator Pro")
st.markdown("Automated Area Extraction | Status Tracking | Identity Concatenation")

uploaded_files = st.file_uploader("Upload Source Excel Files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    if st.button("Generate Final Report"):
        all_data = []
        for f in uploaded_files:
            all_data.extend(parse_excel_wip(f, f.name))
        
        if all_data:
            df = pd.DataFrame(all_data)
            df.insert(0, 'SNO.', range(1, len(df) + 1))
            
            target_cols = [
                'SNO.', 'Pending Days', 'D.Code & JC No.', 'Dname', 'Area', 'Dealer Name', 
                'Req. Delv. Dt', 'Cr.Dt', 'Total', 'PartsTotal', 'Ord.No.', 'Ord.Ty.', 
                'Regn.No', 'Chassis/SL No.', 'Notes', 'Cust.No', 'Cust Name', 
                'Closing Date', 'Remarks', 'Category', 'Invoice No.', 'Date'
            ]
            
            for col in target_cols:
                if col not in df.columns: df[col] = ""
            df = df[target_cols]

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='WIP Summary')
                ws = writer.sheets['WIP Summary']
                
                # Header Styling
                header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                for col in range(1, len(df.columns) + 1):
                    cell = ws.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 22

                # Excel Logic for Category Column
                # Invoice No is col U (21), Category is col T (20)
                green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                green_font = Font(color="006100", bold=True)
                
                for r in range(2, len(df) + 2):
                    # Set the Formula for Category column
                    ws[f'T{r}'] = f'=IF(ISBLANK(U{r}), "", "JOB CARD CLOSED")'
                    
                    # Add green formatting if U is not blank
                    ws.conditional_formatting.add(f'T{r}',
                        FormulaRule(formula=[f'NOT(ISBLANK(U{r}))'], stopIfTrue=True, fill=green_fill, font=green_font))

                ws.auto_filter.ref = ws.dimensions

            st.success("Success! Area names pulled from filenames and report generated.")
            st.download_button("📥 Download Excel Report", output.getvalue(), f"WIP_Summary_{date.today()}.xlsx")