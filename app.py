import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
import base64

def process_excel(uploaded_file):
    # Read both sheets
    nov_df = pd.read_excel(uploaded_file, sheet_name='NOV')
    rec_df = pd.read_excel(uploaded_file, sheet_name='REC')
    
    # Create a mapping dictionary from NOV sheet (Order ID to Item Name)
    order_to_item = dict(zip(nov_df['Order ID'], nov_df['ITEM NAME']))
    
    # Map the item names to REC sheet based on Order ID
    rec_df['ITEM NAME'] = rec_df['Order ID'].map(order_to_item)
    
    return rec_df

def download_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='REC')
    
    b64 = base64.b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="processed_data.xlsx">Download Processed Excel File</a>'
    return href

# Set page config
st.set_page_config(page_title="Excel Processor", layout="wide")

# Custom CSS
st.markdown("""
    <style>
        .main {
            padding: 2rem;
        }
        .stButton>button {
            width: 100%;
            margin-top: 1rem;
        }
        .upload-section {
            padding: 2rem;
            border-radius: 0.5rem;
            border: 2px dashed #4CAF50;
            margin-bottom: 2rem;
        }
        .title {
            color: #2196F3;
            text-align: center;
            margin-bottom: 2rem;
        }
    </style>
""", unsafe_allow_html=True)

# App title
st.markdown("<h1 class='title'>Excel Data Processor</h1>", unsafe_allow_html=True)

# File upload section
st.markdown("<div class='upload-section'>", unsafe_allow_html=True)
st.write("### Upload Your Excel File")
st.write("Please upload an Excel file containing 'NOV' and 'REC' sheets.")
uploaded_file = st.file_uploader("Choose a file", type=['xlsx'])
st.markdown("</div>", unsafe_allow_html=True)

if uploaded_file is not None:
    try:
        # Process the file
        with st.spinner('Processing your file...'):
            result_df = process_excel(uploaded_file)
        
        # Display success message
        st.success('File processed successfully!')
        
        # Show preview of processed data
        st.write("### Preview of Processed Data")
        st.dataframe(result_df.head(10))
        
        # Download button
        st.markdown("### Download Processed File")
        st.markdown(download_excel(result_df), unsafe_allow_html=True)
        
        # Display statistics
        st.write("### Processing Statistics")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total Records", len(result_df))
        
        with col2:
            filled_items = result_df['ITEM NAME'].notna().sum()
            st.metric("Matched Items", filled_items)
        
        with col3:
            match_rate = (filled_items / len(result_df)) * 100
            st.metric("Match Rate", f"{match_rate:.2f}%")

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        st.write("Please make sure your Excel file has the correct format with 'NOV' and 'REC' sheets.")

# Instructions section
with st.expander("How to Use"):
    st.write("""
    1. Prepare your Excel file with two sheets named 'NOV' and 'REC'
    2. The NOV sheet should contain 'Order ID' and 'ITEM NAME' columns
    3. The REC sheet should contain 'Order ID' column
    4. Upload your file using the upload button above
    5. The app will process the data and show you a preview
    6. You can download the processed file using the download link
    """)
