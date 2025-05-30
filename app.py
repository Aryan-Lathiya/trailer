import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Set page config
st.set_page_config(
    page_title="Excel Data Comparison Dashboard",
    page_icon="📊",
    layout="wide"
)

# Title
st.title("📊 Excel Data Comparison Dashboard")
st.markdown("Compare up to 3 rows from your Excel data")

# Load Excel file from directory
excel_file_path = "trailer.xlsx"  # Replace with your actual file name

# Check if file exists
import os
if not os.path.exists(excel_file_path):
    st.error(f"Excel file '{excel_file_path}' not found in the current directory!")
    st.info("Please make sure your Excel file is in the same directory as this script.")
    st.stop()

# Load the file automatically
try:
    # Read Excel file from directory
    df = pd.read_excel(excel_file_path)
    
    # Get the first column name and unique values
    first_column = df.columns[0]
    unique_values = df[first_column].unique().tolist()
    
    st.markdown("---")
    st.subheader("🔍 Select Items to Compare")
    
    # Create three columns for dropdowns
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**First Selection**")
        selection1 = st.selectbox(
            f"Choose from {first_column}:",
            options=["None"] + unique_values,
            key="select1"
        )
    
    with col2:
        st.markdown("**Second Selection**")
        selection2 = st.selectbox(
            f"Choose from {first_column}:",
            options=["None"] + unique_values,
            key="select2"
        )
    
    with col3:
        st.markdown("**Third Selection**")
        selection3 = st.selectbox(
            f"Choose from {first_column}:",
            options=["None"] + unique_values,
            key="select3"
        )
    
    # Filter selected rows
    selected_data = []
    selections = [selection1, selection2, selection3]
    selection_names = ["Selection 1", "Selection 2", "Selection 3"]
    
    for i, selection in enumerate(selections):
        if selection != "None":
            row_data = df[df[first_column] == selection].iloc[0]
            selected_data.append((selection_names[i], selection, row_data))
    
    if selected_data:
        st.markdown("---")
        st.subheader("📋 Comparison Results")
        
        # Create comparison dataframe in vertical format
        comparison_df = pd.DataFrame()
        
        for name, selection, row_data in selected_data:
            comparison_df[f"{name}\n({selection})"] = row_data
        
        # Display comparison table in vertical format (features as rows, selections as columns)
        st.markdown("### Detailed Comparison Table")
        st.dataframe(
            comparison_df,  # Remove .T to keep it vertical
            use_container_width=True,
            height=600  # Increased height since we have more rows now
        )
        
    else:
        st.info("👆 Please select at least one item from the dropdowns to see the comparison.")

except Exception as e:
    st.error(f"Error reading the Excel file: {str(e)}")
    st.info("Please make sure your Excel file is properly formatted and not corrupted.")