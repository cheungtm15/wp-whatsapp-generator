import streamlit as st
import openpyxl
import urllib.parse
import io
import pandas as pd

# 1. Page Config & Layout
st.set_page_config(page_title="WP WhatsApp Generator", page_icon="💬", layout="wide")

st.title("💬 WP WhatsApp Link Generator")
st.write("Upload your spreadsheet below. The app will automatically append **two separate columns** containing your dynamic WhatsApp links side-by-side.")

# 2. File Uploader (No sidebars or toggles needed anymore!)
uploaded_file = st.file_uploader("Upload 'New Orders.xlsx' or 'Today delivery.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        
        # Determine target sheet
        sheet_name = "Summary Data" if "Summary Data" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        
        # Read headers
        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
        
        if "Sales Order (Remarks) Tel" not in headers or "Sales Order Cus. P.O. No." not in headers:
            st.error(f"❌ Required columns missing in sheet '{sheet_name}'. Need 'Sales Order (Remarks) Tel' and 'Sales Order Cus. P.O. No.'.")
        else:
            tel_idx = headers.index("Sales Order (Remarks) Tel") + 1
            po_idx = headers.index("Sales Order Cus. P.O. No.") + 1
            etd_idx = headers.index("Sales Order ETD") + 1 if "Sales Order ETD" in headers else None
            
            # Setup Column 1: WhatsApp Reminder Link
            if "WhatsApp Reminder Link" in headers:
                reminder_col_idx = headers.index("WhatsApp Reminder Link") + 1
            else:
                reminder_col_idx = len(headers) + 1
                ws.cell(row=1, column=reminder_col_idx, value="WhatsApp Reminder Link")
                headers.append("WhatsApp Reminder Link") # update tracking list inline
                
            # Setup Column 2: WhatsApp Confirmation Link
            if "WhatsApp Confirmation Link" in headers:
                confirm_col_idx = headers.index("WhatsApp Confirmation Link") + 1
            else:
                confirm_col_idx = len(headers) + 1
                ws.cell(row=1, column=confirm_col_idx, value="WhatsApp Confirmation Link")
                
            count = 0
            
            # Process all data rows
            for row_idx in range(2, ws.max_row + 1):
                tel_val = ws.cell(row=row_idx, column=tel_idx).value
                po_val = ws.cell(row=row_idx, column=po_idx).value
                etd_val = ws.cell(row=row_idx, column=etd_idx).value if etd_idx else None
                
                if tel_val is not None:
                    # Clean up phone numbers safely
                    p_str = str(tel_val).strip().split('.')[0]
                    p_str = "".join(c for c in p_str if c.isdigit())
                    if len(p_str) == 8:
                        p_str = "852" + p_str
                        
                    po_str = str(po_val).strip() if po_val else ""
                    
                    # Clean up and interpret delivery dates safely
                    date_str = "today"
                    if etd_val:
                        try:
                            date_str = pd.to_datetime(etd_val).strftime("%d %B %Y")
                        except:
                            date_str = str(etd_val).strip()

                    # --- MESSAGE 1: REMINDER TEXT generation ---
                    text_reminder = f"Hello, Thank you for your order with WP at home\nPlease be informed that we will deliver your order {po_str} today, between 11:00 am - 6:00 pm.\xa0\nThank you and enjoy"
                    url_reminder = f"https://api.whatsapp.com/send?phone={p_str}&text={urllib.parse.quote(text_reminder)}"
                    
                    cell_rem = ws.cell(row=row_idx, column=reminder_col_idx, value="Send Reminder")
                    cell_rem.hyperlink = url_reminder
                    cell_rem.font = openpyxl.styles.Font(color="0563C1", underline="single")

                    # --- MESSAGE 2: CONFIRMATION TEXT generation ---
                    text_confirm = (
                        f"Dear Customer,\n\xa0\nThank you for your order!\n\n"
                        f"Your requested delivery date has been confirmed.\n\n"
                        f"We will deliver your order on\xa0 {date_str}, between 11:00 am - 6:00 pm.\n\n"
                        f"Please make sure the phone number provided is correct and have someone can access the door.\n\n"
                        f"Thank you and enjoy!\n\nWP at Home\nPhone: +852 2880 0000\nEmail: wpathome@wavespacific.com"
                    )
                    url_confirm = f"https://api.whatsapp.com/send?phone={p_str}&text={urllib.parse.quote(text_confirm)}"
                    
                    cell_conf = ws.cell(row=row_idx, column=confirm_col_idx, value="Send Confirmation")
                    cell_conf.hyperlink = url_confirm
                    cell_conf.font = openpyxl.styles.Font(color="0563C1", underline="single")
                    
                    count += 1
            
            # Save the updated spreadsheet out to memory buffer
            out_buffer = io.BytesIO()
            wb.save(out_buffer)
            out_buffer.seek(0)
            
            st.success(f"🎉 Generated links for {count} rows across both reminder & confirmation formats!")
            
            # Download asset trigger
            st.download_button(
                label="📥 Download Updated Spreadsheet",
                data=out_buffer,
                file_name=f"processed_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Safe text table formatting logic for visual preview dashboard
            data = []
            updated_headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
            for row_idx in range(2, ws.max_row + 1):
                row_vals = [ws.cell(row=row_idx, column=c_idx).value for c_idx in range(1, ws.max_column + 1)]
                
                # Extract link addresses out to display cleanly in browser framework
                if len(row_vals) >= reminder_col_idx:
                    h_rem = ws.cell(row=row_idx, column=reminder_col_idx).hyperlink
                    row_vals[reminder_col_idx - 1] = h_rem.target if h_rem else ""
                if len(row_vals) >= confirm_col_idx:
                    h_conf = ws.cell(row=row_idx, column=confirm_col_idx).hyperlink
                    row_vals[confirm_col_idx - 1] = h_conf.target if h_conf else ""
                    
                row_vals = ["" if x is None else str(x) for x in row_vals]
                data.append(row_vals)
                
            preview_df = pd.DataFrame(data, columns=updated_headers)
            
            st.markdown("---")
            st.subheader("👀 Generated Data Preview")
            st.dataframe(
                preview_df,
                column_config={
                    "WhatsApp Reminder Link": st.column_config.LinkColumn("WhatsApp Reminder Link", display_text="Test Reminder Link"),
                    "WhatsApp Confirmation Link": st.column_config.LinkColumn("WhatsApp Confirmation Link", display_text="Test Confirm Link")
                },
                width="stretch"
            )
            
    except Exception as e:
        st.error(f"An unexpected system reading error occurred: {e}")
