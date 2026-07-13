import pandas as pd
import numpy as np
import re
import io
import os
import streamlit as st
import threading
import time
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# ==========================================
# 0. KEEP-ALIVE BACKGROUND THREAD
# ==========================================
APP_URL = "https://webactivityreport-p2mxvecaaq35ramit9iw6u.streamlit.app" 

def ping_server():
    while True:
        try:
            requests.get(APP_URL)
        except Exception:
            pass
        time.sleep(600) 

if 'keep_awake_thread' not in st.session_state:
    thread = threading.Thread(target=ping_server, daemon=True)
    thread.start()
    st.session_state['keep_awake_thread'] = True

# ==========================================
# 1. UI & CONTROL FLAGS
# ==========================================
st.set_page_config(page_title="Diamond Web Log Processor", layout="centered")
st.title("💎 Interactive Web-Activity Report Generator")

FILL_HIGHEST_ACTION = False  
FILL_TYPE_COLUMN = True      
FILL_IP_COUNT = False        
OUTPUT_FILE_NAME = 'Weblog_Action_Report.xlsm' # Changed to output Macro file
TEMPLATE_FILE = 'macro_template.xlsm'         # Looks for this in your GitHub repo

# ==========================================
# 2. FILE UPLOADERS
# ==========================================
st.write("Upload your source files below to generate the report.")
weblog_file = st.file_uploader("1. Upload Weblog Data (Excel) *Required*", type=['xlsx'])
master_file = st.file_uploader("2. Upload Color Master Data (Excel) *Required*", type=['xlsx'])
sales_file = st.file_uploader("3. Upload Sales and Bid Data (Excel) *Optional*", type=['xlsx'])
critical_file = st.file_uploader("4. Upload Critical Search Data (Excel) *Required*", type=['xlsx'])

if st.button("Generate Interactive Report"):
    
    if weblog_file and master_file and critical_file:
        
        # FATAL ERROR CHECK: Ensure template exists in GitHub before doing any heavy processing
        if not os.path.exists(TEMPLATE_FILE):
            st.error(f"❌ ERROR: '{TEMPLATE_FILE}' is missing! Please upload it to your GitHub repository.")
            st.stop()
            
        with st.spinner("Crunching data and building macros (this takes time for large files)..."):
            try:
                # 3. DATA CLEANING & LOADING
                # Loading full dataset for the macro drill-down
                logs_df = pd.read_excel(weblog_file)
                master_df = pd.read_excel(master_file)
                critical_df = pd.read_excel(critical_file)

                logs_df.columns = logs_df.columns.str.strip()
                master_df.columns = master_df.columns.str.strip()
                critical_df.columns = critical_df.columns.str.strip()

                if 'Zone' not in master_df.columns:
                    master_df['Zone'] = ''

                logs_df['PARTY_COMPANY_NAME'] = logs_df['PARTY_COMPANY_NAME'].astype(str).str.strip().str.upper()
                logs_df['EMPLOYEE_SHORT_NAME'] = logs_df['EMPLOYEE_SHORT_NAME'].fillna('Unknown').astype(str).str.strip().str.upper()
                logs_df['description'] = logs_df['description'].astype(str).str.strip().str.upper()

                master_df['Company Name'] = master_df['Company Name'].astype(str).str.strip().str.upper()
                master_df['Color'] = master_df['Color'].astype(str).str.strip().str.upper()
                master_df['Zone'] = master_df['Zone'].astype(str).str.strip()

                # 4. CRITICAL SEARCH PREP
                critical_df['Grp'] = critical_df['Grp'].astype(str).str.strip().str.upper()
                critical_filtered = critical_df[critical_df['Grp'].isin(['F3', 'CRITICAL'])]

                critical_keywords = set(
                    critical_filtered['Name'].dropna().astype(str).str.strip().str.upper().tolist() + 
                    critical_filtered['Short Name'].dropna().astype(str).str.strip().str.upper().tolist()
                )

                def extract_action(desc, critical_keys):
                    if pd.isna(desc) or desc == 'NAN': return 'OTHER'
                    if 'SEARCHID' in desc or 'SEARCH PERFORMED' in desc:
                        for kw in critical_keys:
                            if re.search(r'\b' + re.escape(kw) + r'\b', desc): return 'Critical_Search'
                        return 'Detail' 
                    elif any(keyword in desc for keyword in ['VIDEO', 'PHYGITAL', 'CERTIFICATE', 'DETAIL', 'IMAGE', 'PLOTING']): return 'MEDIA'
                    elif 'EXCEL' in desc: return 'EXCEL'
                    elif 'LAYOUT' in desc: return 'Layout'
                    elif 'NEW ARRIVAL' in desc: return 'NEW ARRIVAL'
                    elif 'TWIN STONES' in desc: return 'TWIN STONES'
                    elif 'WISHLIST' in desc: return 'WISHLIST'
                    return 'OTHER'

                logs_df['Action'] = logs_df['description'].apply(lambda x: extract_action(x, critical_keywords))
                
                # Reorder to ensure 'Action' is at the front for the VBA macro to easily find it
                cols = logs_df.columns.tolist()
                cols.insert(0, cols.pop(cols.index('Action')))
                logs_df = logs_df[cols]

                # 5. DATA AGGREGATION
                action_counts = pd.pivot_table(logs_df, index=['PARTY_COMPANY_NAME', 'EMPLOYEE_SHORT_NAME'], columns='Action', aggfunc='size', fill_value=0).reset_index()

                required_cols = ['Detail', 'Critical_Search', 'EXCEL', 'MEDIA', 'Layout', 'NEW ARRIVAL', 'TWIN STONES', 'WISHLIST']
                for col in required_cols:
                    if col not in action_counts.columns: action_counts[col] = 0

                action_counts['Grand Total'] = action_counts[required_cols].sum(axis=1)

                ip_counts = logs_df.groupby(['PARTY_COMPANY_NAME', 'EMPLOYEE_SHORT_NAME'])['ipAddress'].nunique().reset_index()
                ip_counts.rename(columns={'ipAddress': 'IP Counts'}, inplace=True)

                report_df = pd.merge(action_counts, ip_counts, on=['PARTY_COMPANY_NAME', 'EMPLOYEE_SHORT_NAME'], how='left')
                report_df.rename(columns={'EMPLOYEE_SHORT_NAME': 'Sales_Person'}, inplace=True)

                if sales_file is not None:
                    sales_df = pd.read_excel(sales_file, usecols=['Sold Party', 'Type', 'AMT'])
                    sales_df.columns = sales_df.columns.str.strip()
                    sales_df['Sold Party'] = sales_df['Sold Party'].astype(str).str.strip().str.upper()
                    sales_df['Type'] = sales_df['Type'].astype(str).str.strip().str.upper()
                    sales_df['AMT'] = pd.to_numeric(sales_df['AMT'], errors='coerce').fillna(0)
                    valid_sales = sales_df[sales_df['Type'].isin(['SALE', 'BID'])]
                    amt_grouped = valid_sales.groupby('Sold Party')['AMT'].sum().reset_index()
                    report_df = pd.merge(report_df, amt_grouped, left_on='PARTY_COMPANY_NAME', right_on='Sold Party', how='left')
                    report_df['AMT'] = report_df['AMT'].fillna(0)
                else:
                    report_df['AMT'] = 0

                # 6. MASTER LIST FILTERING
                final_df = pd.merge(report_df, master_df[['Company Name', 'Color', 'Zone']], left_on='PARTY_COMPANY_NAME', right_on='Company Name', how='inner')
                final_df.rename(columns={'Zone': 'Type'}, inplace=True)
                final_df['Remark'] = ''
                final_columns = ['PARTY_COMPANY_NAME', 'Type', 'Remark', 'Sales_Person', 'Critical_Search', 'Detail', 'EXCEL', 'MEDIA', 'Layout', 'NEW ARRIVAL', 'TWIN STONES', 'WISHLIST', 'IP Counts', 'Grand Total', 'AMT', 'Color']
                final_df = final_df[final_columns]

                # 7. INJECT DATA INTO MACRO TEMPLATE
                wb = load_workbook(TEMPLATE_FILE, keep_vba=True)
                ws_report = wb['Report']
                ws_weblog = wb['Weblog Data']

                for r in dataframe_to_rows(final_df, index=False, header=True): ws_report.append(r)
                for r in dataframe_to_rows(logs_df, index=False, header=True): ws_weblog.append(r)

                # 8. STYLING
                COLOR_MAP = {'GREEN': '00FF00', 'RED': 'FF0000', 'BLUE': '0000FF', 'YELLOW': 'FFFF00', 'ORANGE': 'FFA500'}
                
                headers_report = [cell.value for cell in ws_report[1]]
                action_indices = [headers_report.index(col) + 1 for col in required_cols] 
                ip_index = headers_report.index('IP Counts') + 1
                color_col_index = headers_report.index('Color') + 1
                type_col_index = headers_report.index('Type') + 1 
                party_name_index = headers_report.index('PARTY_COMPANY_NAME') + 1

                for row_num in range(2, ws_report.max_row + 1):
                    party_name = str(ws_report.cell(row=row_num, column=party_name_index).value).strip().upper()
                    
                    if party_name == "K GIRDHARLAL & CO.":
                        special_fill = PatternFill(start_color="E6B8B7", end_color="E6B8B7", fill_type='solid')
                        for col_idx in range(1, len(headers_report)): ws_report.cell(row=row_num, column=col_idx).fill = special_fill
                        continue 

                    type_val = str(ws_report.cell(row=row_num, column=type_col_index).value).upper()
                    if 'RED' in type_val: hex_code = 'FF0000'
                    elif 'ORANGE' in type_val: hex_code = 'FFA500'
                    elif 'GREEN' in type_val: hex_code = '00FF00'
                    elif 'YELLOW' in type_val: hex_code = 'FFFF00'
                    elif 'BLUE' in type_val: hex_code = '0000FF'
                    else:
                        color_name = str(ws_report.cell(row=row_num, column=color_col_index).value).upper().strip()
                        hex_code = COLOR_MAP.get(color_name, 'CCCCCC') 
                    
                    fill_style = PatternFill(start_color=hex_code, end_color=hex_code, fill_type='solid')

                    if FILL_TYPE_COLUMN: ws_report.cell(row=row_num, column=type_col_index).fill = fill_style

                    if FILL_HIGHEST_ACTION:
                        action_values = [ws_report.cell(row=row_num, column=idx).value or 0 for idx in action_indices]
                        max_val = max(action_values) if action_values else 0
                        if max_val > 0: 
                            for idx, val in zip(action_indices, action_values):
                                if val == max_val: ws_report.cell(row=row_num, column=idx).fill = fill_style
                                
                    if FILL_IP_COUNT:
                        ip_val = ws_report.cell(row=row_num, column=ip_index).value or 0
                        if ip_val >= 5: ws_report.cell(row=row_num, column=ip_index).fill = fill_style

                    for idx in action_indices:
                        cell_val = ws_report.cell(row=row_num, column=idx).value or 0
                        if cell_val > 0:
                            ws_report.cell(row=row_num, column=idx).hyperlink = "#'Weblog Data'!A1"
                            ws_report.cell(row=row_num, column=idx).font = Font(color="0000FF", underline="single")

                ws_report.delete_cols(color_col_index)
                ws_report.freeze_panes = 'A2' 

                center_align = Alignment(horizontal='center', vertical='center')
                thin_border = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'), top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
                header_font = Font(bold=True)
                header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid") 

                current_headers = [cell.value for cell in ws_report[1]]
                for row in ws_report.iter_rows(min_row=1, max_row=ws_report.max_row, min_col=1, max_col=ws_report.max_column):
                    for cell in row:
                        cell.alignment = center_align
                        cell.border = thin_border
                        if cell.row > 1 and cell.column <= len(current_headers) and current_headers[cell.column - 1] == 'AMT': cell.number_format = '#,##0.00'
                        if cell.row == 1:
                            cell.font = header_font
                            cell.fill = header_fill

                for col in ws_report.columns:
                    max_length = 0
                    column_letter = col[0].column_letter 
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                        except: pass
                    ws_report.column_dimensions[column_letter].width = min(max_length + 3, 50) 

                ws_weblog = wb['Weblog Data']
                headers_weblog = [cell.value for cell in ws_weblog[1]]
                ws_weblog.auto_filter.ref = ws_weblog.dimensions
                ws_weblog.freeze_panes = 'A2'
                
                if 'Action' in headers_weblog:
                    action_col_idx = headers_weblog.index('Action') + 1
                    critical_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type='solid')
                    for row_num in range(2, ws_weblog.max_row + 1):
                        if ws_weblog.cell(row=row_num, column=action_col_idx).value == 'Critical_Search':
                            for col_idx in range(1, ws_weblog.max_column + 1):
                                ws_weblog.cell(row=row_num, column=col_idx).fill = critical_fill

                for cell in ws_weblog[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center_align

                final_output = io.BytesIO()
                wb.save(final_output)
                final_output.seek(0)
                
                st.success("Report Generated Successfully!")
                st.download_button(label="📥 Download Interactive Report (.xlsm)", data=final_output, file_name=OUTPUT_FILE_NAME, mime="application/vnd.ms-excel.sheet.macroEnabled.12")

            except Exception as e:
                st.error(f"An error occurred: {e}")
                
    else:
        st.warning("Please upload the required files (Weblog, Color Master, and Critical Search).")
