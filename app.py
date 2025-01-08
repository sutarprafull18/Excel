import streamlit as st
import pandas as pd
import io
from datetime import datetime

# Set page config
st.set_page_config(
    page_title="Excel Sheet Matcher",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for better UI
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #FF4B4B;
        color: white;
        font-weight: bold;
    }
    .success-message {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #D4EDDA;
        color: #155724;
        margin: 1rem 0;
    }
    .error-message {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #F8D7DA;
        color: #721C24;
        margin: 1rem 0;
    }
    .info-message {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #E2E3E5;
        color: #383D41;
        margin: 1rem 0;
    }
    .stDataFrame {
        margin-top: 1rem;
        margin-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

def find_sheet_names(sheets):
    """Find NOC/NOV and REC sheets regardless of case"""
    noc_sheet = None
    rec_sheet = None
    
    for sheet in sheets:
        sheet_lower = sheet.lower()
        if sheet_lower in ['noc', 'nov']:
            noc_sheet = sheet
        elif sheet_lower == 'rec':
            rec_sheet = sheet
            
    return noc_sheet, rec_sheet

def process_sheets(noc_df, rec_df):
    try:
        # Create a mapping of order_id to item_name from NOC sheet
        noc_df = noc_df.fillna('')
        rec_df = rec_df.fillna('')
        
        # Create mapping from NOC sheet
        order_map = {}
        for _, row in noc_df.iterrows():
            if row.iloc[0]:  # If order ID exists
                order_map[str(row.iloc[0]).strip()] = row.iloc[1]  # Map order ID to item name
        
        # Update REC sheet
        for idx, row in rec_df.iterrows():
            order_id = str(row['Order ID']).strip() if 'Order ID' in rec_df.columns else None
            if order_id and order_id in order_map:
                rec_df.at[idx, 'ITEM NAME'] = order_map[order_id]
        
        return rec_df
    except Exception as e:
        st.error(f"Error in processing: {str(e)}")
        return None

def main():
    st.title("📊 Excel Sheet Matcher")
    st.markdown("### Match and populate item names between sheets")
    
    # Session state for storing DataFrames
    if 'noc_df' not in st.session_state:
        st.session_state.noc_df = None
    if 'rec_df' not in st.session_state:
        st.session_state.rec_df = None
    
    # File upload section
    st.markdown("### Upload Excel File")
    uploaded_file = st.file_uploader("Choose Excel file containing NOC and REC sheets", type=['xlsx'])

    if uploaded_file:
        try:
            # Read the Excel file
            xls = pd.ExcelFile(uploaded_file)
            sheets = xls.sheet_names
            
            # Find sheets regardless of case
            noc_sheet, rec_sheet = find_sheet_names(sheets)
            
            if noc_sheet and rec_sheet:
                st.markdown(f'<div class="success-message">✅ Found sheets: {noc_sheet} and {rec_sheet}</div>', 
                          unsafe_allow_html=True)
                
                # Read both sheets and store in session state
                st.session_state.noc_df = pd.read_excel(uploaded_file, sheet_name=noc_sheet)
                st.session_state.rec_df = pd.read_excel(uploaded_file, sheet_name=rec_sheet)
                
                # Show data previews in tabs
                st.markdown("### Sheet Contents")
                tab1, tab2 = st.tabs([f"{noc_sheet} Sheet", f"{rec_sheet} Sheet"])
                
                with tab1:
                    st.markdown(f"#### {noc_sheet} Sheet Data")
                    st.dataframe(st.session_state.noc_df, use_container_width=True)
                    st.markdown(f"Total rows: {len(st.session_state.noc_df)}")
                
                with tab2:
                    st.markdown(f"#### {rec_sheet} Sheet Data")
                    st.dataframe(st.session_state.rec_df, use_container_width=True)
                    st.markdown(f"Total rows: {len(st.session_state.rec_df)}")
                
                # Process button
                if st.button("Process Sheets"):
                    with st.spinner("Processing sheets..."):
                        result_df = process_sheets(st.session_state.noc_df, st.session_state.rec_df)
                        
                        if result_df is not None:
                            st.markdown("### Results")
                            st.dataframe(result_df, use_container_width=True)
                            
                            # Prepare download
                            buffer = io.BytesIO()
                            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                result_df.to_excel(writer, sheet_name='Updated_REC', index=False)
                            
                            buffer.seek(0)
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            
                            # Download button
                            st.download_button(
                                label="📥 Download Processed File",
                                data=buffer,
                                file_name=f"processed_sheets_{timestamp}.xlsx",
                                mime="application/vnd.ms-excel"
                            )
                            
                            # Show success message with counts
                            matched_count = result_df['ITEM NAME'].notna().sum()
                            total_count = len(result_df)
                            st.markdown(
                                f'<div class="success-message">✅ Processing completed successfully!<br>'
                                f'Matched {matched_count} out of {total_count} records.</div>',
                                unsafe_allow_html=True
                            )
            else:
                missing_sheets = []
                if not noc_sheet:
                    missing_sheets.append('NOC/NOV')
                if not rec_sheet:
                    missing_sheets.append('REC')
                    
                st.markdown(
                    f'<div class="error-message">❌ Missing required sheets: {", ".join(missing_sheets)}. '
                    f'Please ensure your Excel file contains both NOC/NOV and REC sheets.</div>',
                    unsafe_allow_html=True
                )
                
                st.markdown("Found sheets in uploaded file:")
                for sheet in sheets:
                    st.markdown(f"- {sheet}")

        except Exception as e:
            st.markdown(
                f'<div class="error-message">❌ Error reading file: {str(e)}</div>',
                unsafe_allow_html=True
            )

    else:
        st.markdown(
            '<div class="info-message">ℹ️ Please upload an Excel file containing both NOC/NOV and REC sheets.</div>',
            unsafe_allow_html=True
        )

if __name__ == "__main__":
    main()
