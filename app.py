import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import date

st.set_page_config(page_title="WIP Summary Consolidator v3", layout="wide")

def parse_excel_wip(file_obj, filename):
    """Parses Excel and calculates Pending Days based on Today's Date."""
    df_raw = pd.read_excel(file_obj, header=None)
    data = []
    
    # Get Today's Date for calculation
    today = date.today()

    # Determine the Hub/Area from the filename
    hub_area = "Unknown"
    fn = filename.upper()
    if "KOTHAGUDEM" in fn: hub_area = "Kothagudem"
    elif "NELLORE" in fn: hub_area = "Nellore"
    elif "RAMGUNDAM" in fn or "RAMAGUNDAM" in fn: hub_area = "Ramagundam"
    elif "HOSKOTE" in fn: hub_area = "Hoskote"
    elif "HYD" in fn: hub_area = "HYD"

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
    
    if header_idx == -1:
        return []

    curr_dname, curr_site_name = None, None
    i = header_idx + 1
    
    while i < len(df_raw):
        row = df_raw.iloc[i]
        
        # Detect Dealer Header row
        if str(row[0]).strip() == 'Dealer':
            curr_dname = str(df_raw.iloc[i, 6]).split('.')[0]
            info_str = str(df_raw.iloc[i, 7]) if not pd.isna(df_raw.iloc[i, 7]) else ""
            parts = [p.strip() for p in info_str.split('|')]
            
            if len(parts) >= 3:
                curr_site_name = parts[2]
            elif len(parts) == 2:
                curr_site_name = parts[1]
            else:
                curr_site_name = parts[0]
            
        # Detect Data row (check if first cell is a number)
        try:
            val_0 = row[0]
            if not pd.isna(val_0) and (isinstance(val_0, (int, float)) or str(val_0).isdigit()):
                def get_v(label):
                    idx = col_map.get(label)
                    return row[idx] if idx is not None else None

                # --- NEW CALCULATION LOGIC ---
                creation_date_raw = get_v('Cr.Dt')
                calculated_pending_days = 0
                
                try:
                    # Convert Cr.Dt to a date object
                    if isinstance(creation_date_raw, pd.Timestamp):
                        creation_date = creation_date_raw.date()
                    else:
                        creation_date = pd.to_datetime(creation_date_raw).date()
                    
                    # Calculate difference
                    calculated_pending_days = (today - creation_date).days
                except:
                    # Fallback to the file's value if date parsing fails
                    calculated_pending_days = val_0

                record = {
                    'Pending Days': calculated_pending_days,
                    'Dname': curr_dname,
                    'Area': hub_area,
                    'Dealer Name': curr_site_name,
                    'Req. Delv. Dt': get_v('Req. Delv. Dt'),
                    'Cr.Dt': creation_date_raw,
                    'Total': get_v('Total'),
                    'PartsTotal': get_v('PartsTotal'),
                    'Ord.No.': get_v('Ord.No.'),
                    'Ord.Ty.': get_v('Ord.Ty.'),
                    'Regn.No': get_v('Regn.No'),
                    'Chassis/SL No.': get_v('Chassis/SL No.'),
                    'Cust.No': get_v('Cust.No'),
                    'Cust Name': get_v('Cust Name'),
                    'Notes': ""
                }
                
                # Check for Notes in the next row
                if i + 1 < len(df_raw):
                    next_row_vals = [str(x).strip() for x in df_raw.iloc[i+1]]
                    if 'Notes' in next_row_vals:
                        note_label_idx = next_row_vals.index('Notes')
                        for k in range(note_label_idx + 1, len(df_raw.columns)):
                            note_val = df_raw.iloc[i+1, k]
                            if not pd.isna(note_val) and str(note_val).strip() != "":
                                record['Notes'] = note_val
                                break
                        i += 1
                data.append(record)
        except:
            pass
        i += 1
    return data

# --- Streamlit UI ---
st.title("📊 WIP Summary Consolidator (Live Calculator)")
st.subheader(f"Today's Date: {date.today().strftime('%d-%b-%Y')}")

uploaded_files = st.file_uploader("Upload Hub Excel Files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    if st.button("Generate Consolidated Report"):
        all_data = []
        for uploaded_file in uploaded_files:
            all_data.extend(parse_excel_wip(uploaded_file, uploaded_file.name))
        
        if all_data:
            df = pd.DataFrame(all_data)
            df.insert(0, 'SNO.', range(1, len(df) + 1))
            
            # Reorder columns to match summary format
            target_cols = ['SNO.', 'Pending Days', 'Dname', 'Area', 'Dealer Name', 'Req. Delv. Dt', 
                           'Cr.Dt', 'Total', 'PartsTotal', 'Ord.No.', 'Ord.Ty.', 'Regn.No', 
                           'Chassis/SL No.', 'Notes', 'Cust.No', 'Cust Name', 'Closing Date', 'Remarks', 'Category']
            for col in target_cols:
                if col not in df.columns: df[col] = ""
            df = df[target_cols]

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='WIP Summary')
                ws = writer.sheets['WIP Summary']
                header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                
                for col in range(1, len(df.columns) + 1):
                    cell = ws.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center')
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 18
                ws.auto_filter.ref = ws.dimensions

            st.success(f"Processed {len(df)} orders. Pending Days calculated relative to {date.today()}!")
            st.download_button(
                label="📥 Download Updated Excel",
                data=output.getvalue(),
                file_name=f"Consolidated_WIP_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )