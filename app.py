import streamlit as st
import openpyxl
import urllib.parse
import io
import pandas as pd

# 1. Page Config & Layout
st.set_page_config(page_title="WP WhatsApp Generator", page_icon="💬")
st.title("💬 WP WhatsApp Link Generator")

st.markdown("---") 

# 2. THE TOGGLE 
st.subheader("Step 1: Choose Your Message Template")
message_type = st.radio(
    "Which notification text should be generated?",
    options=[
        "🚚 Today's Delivery Reminder (Short version with Order No.)",
        "📅 Future Delivery Confirmation (Long version with Delivery Date)"
    ],
    index=0
)

st.markdown("---")

# 3. File Uploader
st.subheader("Step 2: Upload Spreadsheet")
uploaded_file = st.file_uploader("Upload 'New Orders.xlsx' or 'Today delivery.xlsx'", type=["xlsx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        
        # Determine Sheet
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
            
            # Link Column setup
            if "WhatsApp Link" in headers:
                link_col_idx = headers.index("WhatsApp Link") + 1
            else:
                link_col_idx = len(headers) + 1
                ws.cell(row=1, column=link_col_idx, value="WhatsApp Link")
                
            count = 0
            
            for row_idx in range(2, ws.max_row + 1):
                tel_val = ws.cell(row=row_idx, column=tel_idx).value
                po_val = ws.cell(row=row_idx, column=po_idx).value
                etd_val = ws.cell(row=row_idx, column=etd_idx).value if etd_idx else None
                
                if tel_val is not None:
                    # Phone clean
                    p_str = str(tel_val).strip().split('.')[0]
                    p_str = "".join(c for c in p_str if c.isdigit())
                    if len(p_str) == 8:
                        p_str = "852" + p_str
                        
                    po_str = str(po_val).strip() if po_val else ""
                    
                    # Date formatting
                    date_str = "today"
                    if etd_val:
                        try:
                            date_str = pd.to_datetime(etd_val).strftime("%d %B %Y")
                        except:
                            date_str = str(etd_val).strip()

                    # Message Generation
                    if "🚚 Today's Delivery" in message_type:
                        text = f"Hello, Thank you for your order with WP at home\nPlease be informed that we will deliver your order {po_str} today, between 11:00 am - 6:00 pm.\xa0\nThank you and enjoy"
                    else:
                        text = (
                            f"Dear Customer,\n\xa0\nThank you for your order!\n\n"
                            f"Your requested delivery date has been confirmed.\n\n"
                            f"We will deliver your order on\xa0 {date_str}, between 11:00 am - 6:00 pm.\n\n"
                            f"Please make sure the phone number provided is correct and have someone can access the door.\n\n"
                            f"Thank you and enjoy!\n\nWP at Home\nPhone: +852 2880 0000\nEmail: wpathome@wavespacific.com"
                        )
                    
                    encoded_text = urllib.parse.quote(text)
                    whatsapp_url = f"https://api.whatsapp.com/send?phone={p_str}&text={encoded_text}"
                    
                    cell = ws.cell(row=row_idx, column=link_col_idx, value="Send WhatsApp")
                    cell.hyperlink = whatsapp_url
                    cell.font = openpyxl.styles.Font(color="0563C1", underline="single")
                    count += 1
            
            # Build download asset
            out_buffer = io.BytesIO()
            wb.save(out_buffer)
            out_buffer.seek(0)
            
            st.success(f"🎉 Created {count} WhatsApp links successfully!")
            st.download_button(
                label="📥 Download Updated Spreadsheet",
                data=out_buffer,
                file_name=f"processed_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Reconstruct preview data safely to prevent Arrow Type Errors
            data = []
            updated_headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
            for row_idx in range(2, ws.max_row + 1):
                row_vals = [ws.cell(row=row_idx, column=c_idx).value for c_idx in range(1, ws.max_column + 1)]
                if len(row_vals) >= link_col_idx:
                    h_target = ws.cell(row=row_idx, column=link_col_idx).hyperlink
                    row_vals[link_col_idx - 1] = h_target.target if h_target else ""
                # Convert all items to clean strings to prevent backend preview crashes
                row_vals = ["" if x is None else str(x) for x in row_vals]
                data.append(row_vals)
                
            preview_df = pd.DataFrame(data, columns=updated_headers)
            
            st.markdown("---")
            st.subheader("👀 Generated Data Preview")
            st.dataframe(
                preview_df,
                column_config={
                    "WhatsApp Link": st.column_config.LinkColumn("WhatsApp Link", display_text="Test Link")
                },
                width="stretch"  # Updated configuration fix matching modern Streamlit spec
            )
            
    except Exception as e:
        st.error(f"An unexpected system reading error occurred: {e}")
