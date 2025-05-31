import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

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
if not os.path.exists(excel_file_path):
    st.error(f"Excel file '{excel_file_path}' not found in the current directory!")
    st.info("Please make sure your Excel file is in the same directory as this script.")
    st.stop()

# Initialize session state for editable data
if 'editable_data' not in st.session_state:
    st.session_state.editable_data = {}

# Initialize session state for tracking data refresh
if 'data_refresh_needed' not in st.session_state:
    st.session_state.data_refresh_needed = False

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
            row_data = df[df[first_column] == selection].iloc[0].copy()
            selected_data.append((selection_names[i], selection, row_data, i))
    
    if selected_data:
        st.markdown("---")
        
        # Create editable input fields section
        st.markdown("Adjust the values below to recalculate costs:")
        
        # Create columns for editable inputs
        input_cols = st.columns(len(selected_data))
        
        for idx, (name, selection, row_data, original_idx) in enumerate(selected_data):
            with input_cols[idx]:
                
                # Initialize session state for this selection if not exists
                selection_key = f"{selection}_{original_idx}"
                if selection_key not in st.session_state.editable_data:
                    # Load actual values from Excel data, not default zeros
                    try:
                        excel_excess_km_charge = float(row_data.get('Excess KM charge (Per KM)', 0)) if pd.notna(row_data.get('Excess KM charge (Per KM)', 0)) else 0.0
                    except (ValueError, TypeError):
                        excel_excess_km_charge = 0.0
                    
                    try:
                        excel_estimated_km_per_month = float(row_data.get('Estimated KM (Per Month)', 0)) if pd.notna(row_data.get('Estimated KM (Per Month)', 0)) else 0.0
                    except (ValueError, TypeError):
                        excel_estimated_km_per_month = 0.0
                    
                    st.session_state.editable_data[selection_key] = {
                        'excess_km_charge': excel_excess_km_charge,
                        'estimated_km_per_month': excel_estimated_km_per_month
                    }
                
                # Editable fields - show actual values from Excel or session state
                excess_km_charge = st.number_input(
                    "Excess KM Charge (Per KM)",
                    value=st.session_state.editable_data[selection_key]['excess_km_charge'],
                    min_value=0.0,
                    step=0.1,
                    format="%.2f",
                    key=f"excess_km_charge_{selection_key}",
                    help="Current value from Excel file or your last edit"
                )
                
                estimated_km_per_month = st.number_input(
                    "Estimated KM (Per Month)",
                    value=st.session_state.editable_data[selection_key]['estimated_km_per_month'],
                    min_value=0.0,
                    step=1.0,
                    format="%.1f",
                    key=f"estimated_km_per_month_{selection_key}",
                    help="Current value from Excel file or your last edit"
                )
                
                # Update session state
                st.session_state.editable_data[selection_key]['excess_km_charge'] = excess_km_charge
                st.session_state.editable_data[selection_key]['estimated_km_per_month'] = estimated_km_per_month
        
        # Calculate and display comparison
        st.markdown("### 📊 Updated Comparison")
        
        # Create card-based comparison
        updated_rows = []
        
        # Create columns for cards
        card_cols = st.columns(len(selected_data))
        
        for idx, (name, selection, row_data, original_idx) in enumerate(selected_data):
            selection_key = f"{selection}_{original_idx}"
            
            # Get editable values
            excess_km_charge = st.session_state.editable_data[selection_key]['excess_km_charge']
            estimated_km_per_month = st.session_state.editable_data[selection_key]['estimated_km_per_month']
            
            # Get existing values with proper type conversion
            try:
                kilo_meters_per_month = float(row_data.get('Kilo Meters (Per Month)', 0)) if pd.notna(row_data.get('Kilo Meters (Per Month)', 0)) else 0.0
            except (ValueError, TypeError):
                kilo_meters_per_month = 0.0
                
            try:
                rent_cost_over_lease_period = float(row_data.get('Rent Cost (Over Lease Period)', 0)) if pd.notna(row_data.get('Rent Cost (Over Lease Period)', 0)) else 0.0
            except (ValueError, TypeError):
                rent_cost_over_lease_period = 0.0
            
            try:
                sticker_cost = float(row_data.get('Sticker Cost', 0)) if pd.notna(row_data.get('Sticker Cost', 0)) else 0.0
            except (ValueError, TypeError):
                sticker_cost = 0.0
            
            try:
                insurance_cost = float(row_data.get('Insurance Cost', 0)) if pd.notna(row_data.get('Insurance Cost', 0)) else 0.0
            except (ValueError, TypeError):
                insurance_cost = 0.0
            
            try:
                months = float(row_data.get('MONTH', 48)) if pd.notna(row_data.get('MONTH', 48)) else 48.0  # Default to 48 months if not found
            except (ValueError, TypeError):
                months = 48.0
            
            # Calculate Excess KM Cost (Over the Lease Term) - using the MONTH column value
            if kilo_meters_per_month == 0:
                excess_km_cost_over_lease_term = 0.0
            elif estimated_km_per_month > kilo_meters_per_month:
                excess_km_cost_over_lease_term = (estimated_km_per_month - kilo_meters_per_month) * excess_km_charge * months
            else:
                excess_km_cost_over_lease_term = 0.0
            
            # Calculate Total Cost over the Lease Term
            total_cost_over_lease_term = rent_cost_over_lease_period + excess_km_cost_over_lease_term + sticker_cost + insurance_cost
            
            # Calculate Cost per Month
            cost_per_month = total_cost_over_lease_term / months if months > 0 else 0.0
            
            # Store raw values for saving
            raw_updated_row = row_data.copy()
            raw_updated_row['Excess KM charge (Per KM)'] = excess_km_charge
            raw_updated_row['Estimated KM (Per Month)'] = estimated_km_per_month
            raw_updated_row['Excess KM Cost (Over the Lease Term)'] = excess_km_cost_over_lease_term
            raw_updated_row['Total Cost over the Lease Term'] = total_cost_over_lease_term
            raw_updated_row['Cost per Month'] = cost_per_month
            updated_rows.append((selection, raw_updated_row))
            
            # Display card with fixed height
            with card_cols[idx]:
                st.markdown(f"""
                <div style="
                    border: 1px solid #e0e0e0;
                    border-radius: 8px;
                    padding: 20px;
                    margin-bottom: 10px;
                    background-color: #f8f9fa;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                ">
                    <h4 style="margin-top: 0; color: #333; border-bottom: 2px solid #007bff; padding-bottom: 8px;">
                        {selection}
                    </h4>
                    <div style="flex-grow: 1;">
                """, unsafe_allow_html=True)
                
                # Display all fields from the row
                for col, value in row_data.items():
                    if col != first_column:  # Skip the first column as it's already shown in header
                        # Format the value based on the column
                        if col in ['Excess KM charge (Per KM)', 'Estimated KM (Per Month)', 'Excess KM Cost (Over the Lease Term)', 'Total Cost over the Lease Term', 'Cost per Month']:
                            if col == 'Excess KM charge (Per KM)':
                                formatted_value = f"{excess_km_charge:,.2f}"
                            elif col == 'Estimated KM (Per Month)':
                                formatted_value = f"{estimated_km_per_month:,.1f} km"
                            elif col == 'Excess KM Cost (Over the Lease Term)':
                                formatted_value = f"{excess_km_cost_over_lease_term:,.2f}"
                            elif col == 'Total Cost over the Lease Term':
                                formatted_value = f"{total_cost_over_lease_term:,.2f}"
                            elif col == 'Cost per Month':
                                formatted_value = f"{cost_per_month:,.2f}"
                        elif col == 'Rent Cost (Over Lease Period)':
                            formatted_value = f"{rent_cost_over_lease_period:,.2f}"
                        elif col == 'Kilo Meters (Per Month)':
                            formatted_value = f"{kilo_meters_per_month:,.0f} km"
                        elif col in ['Sticker Cost', 'Insurance Cost']:
                            cost_value = sticker_cost if col == 'Sticker Cost' else insurance_cost
                            formatted_value = f"{cost_value:,.2f}"
                        else:
                            if pd.isna(value):
                                formatted_value = "-"
                            else:
                                formatted_value = str(value)
                        
                        # Highlight calculated fields
                        if col in ['Rent Cost (Over Lease Period)', 'Sticker Cost', 'Insurance Cost', 'Excess KM Cost (Over the Lease Term)', 'Total Cost over the Lease Term', 'Cost per Month']:
                            st.markdown(f"**{col}:** <span style='color: #007bff; font-weight: bold;'>{formatted_value}</span>", unsafe_allow_html=True)
                        else:
                            st.markdown(f"**{col}:** {formatted_value}")
                
                st.markdown("</div></div>", unsafe_allow_html=True)
        
        # Highlight calculated fields
        st.markdown("### 🧮 Calculated Values Summary")
        summary_cols = st.columns(len(selected_data))
        
        for idx, (name, selection, row_data, original_idx) in enumerate(selected_data):
            with summary_cols[idx]:
                selection_key = f"{selection}_{original_idx}"
                excess_km_charge = st.session_state.editable_data[selection_key]['excess_km_charge']
                estimated_km_per_month = st.session_state.editable_data[selection_key]['estimated_km_per_month']
                
                # Get existing values with proper type conversion
                try:
                    kilo_meters_per_month = float(row_data.get('Kilo Meters (Per Month)', 0)) if pd.notna(row_data.get('Kilo Meters (Per Month)', 0)) else 0.0
                except (ValueError, TypeError):
                    kilo_meters_per_month = 0.0
                    
                try:
                    rent_cost_over_lease_period = float(row_data.get('Rent Cost (Over Lease Period)', 0)) if pd.notna(row_data.get('Rent Cost (Over Lease Period)', 0)) else 0.0
                except (ValueError, TypeError):
                    rent_cost_over_lease_period = 0.0
                
                try:
                    sticker_cost = float(row_data.get('Sticker Cost', 0)) if pd.notna(row_data.get('Sticker Cost', 0)) else 0.0
                except (ValueError, TypeError):
                    sticker_cost = 0.0
                
                try:
                    insurance_cost = float(row_data.get('Insurance Cost', 0)) if pd.notna(row_data.get('Insurance Cost', 0)) else 0.0
                except (ValueError, TypeError):
                    insurance_cost = 0.0
                
                try:
                    months = float(row_data.get('MONTH', 48)) if pd.notna(row_data.get('MONTH', 48)) else 48.0
                except (ValueError, TypeError):
                    months = 48.0
                
                if kilo_meters_per_month == 0:
                    excess_km_cost_over_lease_term = 0.0
                elif estimated_km_per_month > kilo_meters_per_month:
                    excess_km_cost_over_lease_term = (estimated_km_per_month - kilo_meters_per_month) * excess_km_charge * months
                else:
                    excess_km_cost_over_lease_term = 0.0
                
                total_cost_over_lease_term = rent_cost_over_lease_period + excess_km_cost_over_lease_term + sticker_cost + insurance_cost
                cost_per_month = total_cost_over_lease_term / months if months > 0 else 0.0
                
                st.markdown(f"**{selection}**")
                st.metric("Excess KM Cost (Over the Lease Term)", f"{excess_km_cost_over_lease_term:,.2f}")
                st.metric("Total Cost over the Lease Term", f"{total_cost_over_lease_term:,.2f}")
                st.metric("Cost per Month", f"{cost_per_month:,.2f}")
        
        # Save and Download buttons
        st.markdown("---")
        col_save, col_download, col_info = st.columns([1, 1, 2])
        
        with col_save:
            if st.button("💾 Save Changes to Excel", type="primary"):
                try:
                    # Read the original Excel file
                    df_original = pd.read_excel(excel_file_path)
                    
                    # Update the dataframe with calculated values
                    for selection, updated_row in updated_rows:
                        # Find the row index in original dataframe
                        row_index = df_original[df_original[first_column] == selection].index[0]
                        
                        # Update the specific columns
                        df_original.loc[row_index, 'Excess KM charge (Per KM)'] = updated_row['Excess KM charge (Per KM)']
                        df_original.loc[row_index, 'Estimated KM (Per Month)'] = updated_row['Estimated KM (Per Month)']
                        df_original.loc[row_index, 'Excess KM Cost (Over the Lease Term)'] = updated_row['Excess KM Cost (Over the Lease Term)']
                        df_original.loc[row_index, 'Total Cost over the Lease Term'] = updated_row['Total Cost over the Lease Term']
                        df_original.loc[row_index, 'Cost per Month'] = updated_row['Cost per Month']
                    
                    # Save back to Excel
                    df_original.to_excel(excel_file_path, index=False)
                    st.success("✅ Changes saved successfully to Excel file!")
                    
                    # Clear session state to force reload of fresh data
                    st.session_state.editable_data = {}
                    st.session_state.data_refresh_needed = True
                    
                    # Small delay and then rerun to show updated values
                    import time
                    time.sleep(0.5)
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Error saving to Excel: {str(e)}")
        
        with col_download:
            try:
                # Prepare the updated dataframe for download
                df_for_download = pd.read_excel(excel_file_path)
                
                # Update the dataframe with current calculated values
                for selection, updated_row in updated_rows:
                    # Find the row index in dataframe
                    row_index = df_for_download[df_for_download[first_column] == selection].index[0]
                    
                    # Update the specific columns with current values
                    df_for_download.loc[row_index, 'Excess KM charge (Per KM)'] = updated_row['Excess KM charge (Per KM)']
                    df_for_download.loc[row_index, 'Estimated KM (Per Month)'] = updated_row['Estimated KM (Per Month)']
                    df_for_download.loc[row_index, 'Excess KM Cost (Over the Lease Term)'] = updated_row['Excess KM Cost (Over the Lease Term)']
                    df_for_download.loc[row_index, 'Total Cost over the Lease Term'] = updated_row['Total Cost over the Lease Term']
                    df_for_download.loc[row_index, 'Cost per Month'] = updated_row['Cost per Month']
                
                # Convert dataframe to Excel bytes
                from io import BytesIO
                excel_buffer = BytesIO()
                df_for_download.to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_buffer.seek(0)
                
                # Create download button
                st.download_button(
                    label="📥 Download Data",
                    data=excel_buffer.getvalue(),
                    file_name=f"updated_{excel_file_path}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download Excel file with current calculated values"
                )
                
            except Exception as e:
                st.error(f"❌ Error preparing download: {str(e)}")
        
        with col_info:
            st.info("💡 'Save Changes' updates the original file permanently. 'Download Data' gives you a copy with current calculations.")
        
    else:
        st.info("👆 Please select at least one item from the dropdowns to see the comparison.")

except Exception as e:
    st.error(f"Error reading the Excel file: {str(e)}")
    st.info("Please make sure your Excel file is properly formatted and not corrupted.")