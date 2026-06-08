import streamlit as st
import openpyxl
import urllib.parse
import io
import pandas as pd

# Page Configuration
st.set_page_config(page_title="WP WhatsApp Link Generator", page_icon="💬", layout="wide")

st.title("💬 WP WhatsApp Link Generator")
st.write("Upload either a **New Orders** or **Today Delivery** Excel file (`.xlsx`) to automatically append a column with formatted, clickable WhatsApp links.")

# File Uploader
uploaded_file = st.file_uploader("Choose an Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    try:
        # Load the workbook from the memory buffer
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        
        # Step 1: Detect the target sheet (prefers 'Summary Data')
        if "Summary Data" in wb.sheetnames:
            sheet_name = "Summary Data"
        else:
            # Fallback: find the first sheet that contains the required column headers
            sheet_name = wb.sheetnames[0]
            for name in wb.sheetnames:
                ws = wb[name]
                headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
                if "Sales Order (Remarks) Tel" in headers and "Sales Order Cus. P.O. No." in headers:
                    sheet_name = name
                    break
                    
        ws = wb[sheet_name]
        
        # Step 2: Extract column headers to find cell coordinates
        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
        
        if "Sales Order (Remarks) Tel" not in headers or "Sales Order Cus. P.O. No." not in headers:
            st.error(f"⚠️ Could not find the required columns ('Sales Order (Remarks) Tel' and 'Sales Order Cus. P.O. No.') in the sheet '{sheet_name}'.")
        else:
            tel_idx = headers.index("Sales Order (Remarks) Tel") + 1
            po_idx = headers.index("Sales Order Cus. P.O. No.") + 1
            
            # Determine or create the WhatsApp Link column index
            if "WhatsApp Link" in headers:
                link_col_idx = headers.index("WhatsApp Link") + 1
            else:
                link_col_idx = len(headers) + 1
                ws.cell(row=1, column=link_col_idx, value="WhatsApp Link")
                
            count = 0
            # Step 3: Populate each row with customized WhatsApp links
            for row_idx in range(2, ws.max_row + 1):
                tel_val = ws.cell(row=row_idx, column=tel_idx).value
                po_val = ws.cell(row=row_idx, column=po_idx).value
                
                if tel_val is not None and po_val is not None:
                    # Clean and format phone number
                    p_str = str(tel_val).strip()
                    if p_str.endswith('.0'):  # Remove floating point suffix if read as float
                        p_str = p_str[:-2]
                    p_str = "".join(c for c in p_str if c.isdigit())
                    if len(p_str) == 8:
                        p_str = "852" + p_str  # Add Hong Kong country code prefix
                        
                    # Format order number
                    po_str = str(po_val).strip()
                    
                    # Message template matching user format with explicit non-breaking spaces (\xa0)
                    text = f"Hello, Thank you for your order with WP at home\nPlease be informed that we will deliver your order {po_str} today, between 11:00 am - 6:00 pm. \nThank\xa0you\xa0and\xa0enjoy"
                    encoded_text = urllib.parse.quote(text)
                    whatsapp_url = f"https://api.whatsapp.com/send?phone={p_str}&text={encoded_text}"
                    
                    # Assign value, hyperlink target, and native Excel hyperlink styling
                    cell = ws.cell(row=row_idx, column=link_col_idx, value="Send WhatsApp")
                    cell.hyperlink = whatsapp_url
                    cell.font = openpyxl.styles.Font(color="0563C1", underline="single")
                    count += 1
            
            # Save the modified workbook to an in-memory buffer for downloading
            out_buffer = io.BytesIO()
            wb.save(out_buffer)
            out_buffer.seek(0)
            
            # Step 4: Generate web browser preview table
            data = []
            updated_headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
            for row_idx in range(2, ws.max_row + 1):
                row_vals = [ws.cell(row=row_idx, column=c_idx).value for c_idx in range(1, ws.max_column + 1)]
                if len(row_vals) >= link_col_idx:
                    h_target = ws.cell(row=row_idx, column=link_col_idx).hyperlink
                    row_vals[link_col_idx - 1] = h_target.target if h_target else ""
                data.append(row_vals)
                
            preview_df = pd.DataFrame(data, columns=updated_headers)
            
            st.success(f"🎉 Successfully processed {count} records in sheet: **{sheet_name}**!")
            
            # File Download Button
            st.download_button(
                label="📥 Download Updated Excel File",
                data=out_buffer,
                file_name=f"whatsapp_links_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # In-App Dataframe Preview
            st.subheader("👀 Preview of Data and Links")
            st.dataframe(
                preview_df,
                column_config={
                    "WhatsApp Link": st.column_config.LinkColumn("WhatsApp Link", display_text="Test Link")
                },
                use_container_width=True
            )
    except Exception as e:
        st.error(f"❌ An error occurred while parsing the spreadsheet: {e}")
