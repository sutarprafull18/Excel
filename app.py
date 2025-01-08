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
    </style>
""", unsafe_allow_html=True)

def process_sheets(nov_df, rec_df):
    try:
        # Create a mapping of order_id to item_name from NOV sheet
        nov_df = nov_df.fillna('')  # Handle any NaN values
        rec_df = rec_df.fillna('')
        
        # Create mapping from NOV sheet (assuming first column is Order ID and second is Item Name)
        order_map = {}
        for _, row in nov_df.iterrows():
            if row.iloc[0]:  # If order ID exists
                order_map[row.iloc[0]] = row.iloc[1]  # Map order ID to item name
        
        # Update REC sheet
        for idx, row in rec_df.iterrows():
            order_id = row['Order ID'] if 'Order ID' in rec_df.columns else None
            if order_id and order_id in order_map:
                rec_df.at[idx, 'ITEM NAME'] = order_map[order_id]
        
        return rec_df
    except Exception as e:
        st.error(f"Error in processing: {str(e)}")
        return None

def main():
    st.title("📊 Excel Sheet Matcher")
    st.markdown("### Match and populate item names between NOV and REC sheets")
    
    # File upload section
    st.markdown("### Upload Excel File")
    uploaded_file = st.file_uploader("Choose Excel file containing NOV and REC sheets", type=['xlsx'])

    if uploaded_file:
        try:
            # Read the Excel file
            xls = pd.ExcelFile(uploaded_file)
            sheets = xls.sheet_names
            
            # Check if required sheets exist
            if 'NOV' in sheets and 'REC' in sheets:
                st.markdown('<div class="success-message">✅ Both NOV and REC sheets found!</div>', unsafe_allow_html=True)
                
                # Read both sheets
                nov_df = pd.read_excel(uploaded_file, sheet_name='NOV')
                rec_df = pd.read_excel(uploaded_file, sheet_name='REC')
                
                # Show data previews in tabs
                st.markdown("### Preview of Sheets")
                tab1, tab2 = st.tabs(["NOV Sheet", "REC Sheet"])
                
                with tab1:
                    st.dataframe(nov_df.head(), use_container_width=True)
                
                with tab2:
                    st.dataframe(rec_df.head(), use_container_width=True)
                
                # Process button
                if st.button("Process Sheets"):
                    with st.spinner("Processing sheets..."):
                        result_df = process_sheets(nov_df, rec_df)
                        
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
                            
                            st.markdown(
                                '<div class="success-message">✅ Processing completed successfully!</div>',
                                unsafe_allow_html=True
                            )
            else:
                missing_sheets = []
                if 'NOV' not in sheets:
                    missing_sheets.append('NOV')
                if 'REC' not in sheets:
                    missing_sheets.append('REC')
                    
                st.markdown(
                    f'<div class="error-message">❌ Missing required sheets: {", ".join(missing_sheets)}. '
                    f'Please ensure your Excel file contains both NOV and REC sheets.</div>',
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
            '<div class="info-message">ℹ️ Please upload an Excel file containing both NOV and REC sheets.</div>',
            unsafe_allow_html=True
        )

if __name__ == "__main__":
    main()
