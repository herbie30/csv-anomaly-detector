import streamlit as st
import pandas as pd
import numpy as np
from scipy import stats
import os
import sys
import io
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime

# App title and sidebar setup
st.set_page_config(page_title="Container Data Analysis Tool", layout="wide")
st.title("Container Data Analysis Tool")

# Create tabs for different functionalities
tab1, tab2 = st.tabs(["Container Comparison", "Anomaly Detection"])

# Function to read file based on its extension
def read_file(file):
    """Read a file based on its extension (CSV or Excel)."""
    if file is None:
        return None
        
    file_name = file.name.lower()
    
    if file_name.endswith('.csv'):
        return pd.read_csv(file)
    elif file_name.endswith(('.xlsx', '.xls')):
        return pd.read_excel(file, engine='openpyxl')
    else:
        st.error(f"Unsupported file format: {file_name}")
        return None

# Identify container column in dataframe
def identify_container_column(df):
    """Try to identify the container number column in the dataframe."""
    if df is None:
        return None
        
    # Common container column names
    possible_column_names = [
        'CONTAINER NUMBER', 'CONTAINER', 'CONTAINER NO', 'CONTAINER_NUMBER',
        'container', 'container_number', 'container_no', 'containerno', 'container_id'
    ]
    
    # Check exact matches first
    for col_name in possible_column_names:
        if col_name in df.columns:
            return col_name
    
    # Then check for partial matches
    for col in df.columns:
        if 'container' in col.lower():
            return col
    
    return None

# Function to compare container spreadsheets
def compare_container_spreadsheets(tops_df, cyman_df, container_column):
    """
    Compare container numbers between TOPS and Cyman dataframes and identify mismatches.
    
    Parameters:
    tops_df (DataFrame): TOPS dataframe
    cyman_df (DataFrame): Cyman dataframe
    container_column (str): Name of the column containing container numbers
    
    Returns:
    DataFrame: Results dataframe with mismatches
    """
    if tops_df is None or cyman_df is None or container_column is None:
        return None
    
    # Check if container column exists in both dataframes
    if container_column not in tops_df.columns or container_column not in cyman_df.columns:
        st.error(f"Column '{container_column}' not found in one or both spreadsheets")
        return None
    
    # Extract container numbers
    tops_containers = set(tops_df[container_column].astype(str).str.strip())
    cyman_containers = set(cyman_df[container_column].astype(str).str.strip())
    
    # Find mismatches
    tops_only = tops_containers - cyman_containers
    cyman_only = cyman_containers - tops_containers
    
    # Create results DataFrame
    results = []
    
    for container in sorted(tops_only):
        results.append({
            'CONTAINER NUMBER': container,
            'CYMAN': 'Missing',
            'TOPS': 'Present'
        })
    
    for container in sorted(cyman_only):
        results.append({
            'CONTAINER NUMBER': container,
            'CYMAN': 'Present',
            'TOPS': 'Missing'
        })
    
    if not results:
        return pd.DataFrame(columns=['CONTAINER NUMBER', 'CYMAN', 'TOPS'])
    
    results_df = pd.DataFrame(results)
    return results_df

# Function to format Excel file
def format_excel_output(excel_data, results_df):
    """
    Format the Excel output with colors and styling
    
    Parameters:
    excel_data (BytesIO): BytesIO object to write to
    results_df (DataFrame): DataFrame with results
    
    Returns:
    BytesIO: Formatted Excel file as BytesIO object
    """
    row_count = len(results_df)
    
    # Write dataframe to Excel file
    with pd.ExcelWriter(excel_data, engine='openpyxl') as writer:
        results_df.to_excel(writer, index=False, sheet_name='Results')
    
    # Open the workbook from the BytesIO object
    excel_data.seek(0)
    wb = openpyxl.load_workbook(excel_data)
    ws = wb.active
    
    # Define styles
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    missing_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    present_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    # Format header row
    for col in range(1, 4):  # Columns A, B, C
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_alignment
    
    # Format data rows
    for row in range(2, row_count + 2):  # Skip header row
        # Center align container number
        ws.cell(row=row, column=1).alignment = center_alignment
        
        # Format CYMAN column (column B)
        cyman_cell = ws.cell(row=row, column=2)
        cyman_cell.alignment = center_alignment
        if cyman_cell.value == 'Missing':
            cyman_cell.fill = missing_fill
            cyman_cell.value = '❌'
        else:
            cyman_cell.fill = present_fill
            cyman_cell.value = '✓'
        
        # Format TOPS column (column C)
        tops_cell = ws.cell(row=row, column=3)
        tops_cell.alignment = center_alignment
        if tops_cell.value == 'Missing':
            tops_cell.fill = missing_fill
            tops_cell.value = '❌'
        else:
            tops_cell.fill = present_fill
            tops_cell.value = '✓'
    
    # Set column widths
    ws.column_dimensions[get_column_letter(1)].width = 20  # Container Number
    ws.column_dimensions[get_column_letter(2)].width = 12  # CYMAN
    ws.column_dimensions[get_column_letter(3)].width = 12  # TOPS
    
    # Add a title
    ws.insert_rows(1)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "Container Number Comparison Report"
    title_cell.font = Font(size=14, bold=True)
    ws.merge_cells('A1:C1')
    title_cell.alignment = Alignment(horizontal='center')
    
    # Add timestamp
    ws.insert_rows(2)
    date_cell = ws.cell(row=2, column=1)
    date_cell.value = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws.merge_cells('A2:C2')
    date_cell.alignment = Alignment(horizontal='center')
    
    # Add summary
    ws.append([])  # Empty row
    summary_row = ws.max_row + 1
    summary_cell = ws.cell(row=summary_row, column=1)
    summary_cell.value = f"Total mismatches found: {row_count}"
    summary_cell.font = Font(bold=True)
    ws.merge_cells(f'A{summary_row}:C{summary_row}')
    
    # Save the workbook to BytesIO
    excel_data.seek(0)
    excel_data.truncate(0)
    wb.save(excel_data)
    excel_data.seek(0)
    
    return excel_data

# Define anomaly detection functions
def detect_missing_values(df, container_col='container_number'):
    """Detect missing values and return sorted by container number."""
    if df is None or df.empty:
        return pd.DataFrame()
        
    missing_by_column = df.isnull().sum()
    cols_with_missing = missing_by_column[missing_by_column > 0].index.tolist()
    
    if not cols_with_missing:
        return pd.DataFrame()
    
    # Filter rows that have missing values in any of the identified columns
    rows_with_missing = df[df[cols_with_missing].isnull().any(axis=1)]
    
    # Add a flag column to indicate missing values
    rows_with_missing = rows_with_missing.copy()
    rows_with_missing['missing_in_columns'] = rows_with_missing[cols_with_missing].isnull().apply(
        lambda x: ', '.join(x.index[x].tolist()), axis=1
    )
    
    # Sort by container number if it exists
    if container_col in rows_with_missing.columns:
        rows_with_missing = rows_with_missing.sort_values(by=container_col)
    
    return rows_with_missing

def detect_duplicate_rows(df, container_col='container_number'):
    """Detect duplicate rows and return sorted by container number."""
    if df is None or df.empty:
        return pd.DataFrame()
        
    duplicates = df[df.duplicated(keep=False)]
    
    # Sort by container number if it exists
    if not duplicates.empty and container_col in duplicates.columns:
        duplicates = duplicates.sort_values(by=container_col)
    
    return duplicates

def detect_outliers_zscore(df, threshold=3, container_col='container_number'):
    """Detect outliers using Z-score and return sorted by container number."""
    if df is None or df.empty:
        return pd.DataFrame()
        
    numeric_df = df.select_dtypes(include=[np.number])
    if numeric_df.empty:
        return pd.DataFrame()
    
    z_scores = np.abs(stats.zscore(numeric_df, nan_policy='omit'))
    
    # Flag rows where any column's z-score exceeds the threshold
    outlier_mask = (z_scores > threshold).any(axis=1)
    outliers = df[outlier_mask].copy()
    
    # Add information about which columns are outliers
    for col in numeric_df.columns:
        col_z_scores = np.abs(stats.zscore(numeric_df[col], nan_policy='omit'))
        outlier_cols = []
        for i, is_outlier in enumerate(outlier_mask):
            if is_outlier and i < len(col_z_scores) and col_z_scores[i] > threshold:
                outlier_cols.append(i)
        
        if outlier_cols:
            outliers[f'{col}_is_outlier'] = outliers.index.isin(outlier_cols)
    
    # Sort by container number if it exists
    if container_col in outliers.columns:
        outliers = outliers.sort_values(by=container_col)
    
    return outliers

def detect_custom_anomaly(df, container_col='container_number'):
    """Detect custom anomalies (negative values) and return sorted by container number."""
    if df is None or df.empty:
        return pd.DataFrame()
        
    numeric_df = df.select_dtypes(include=[np.number])
    if numeric_df.empty:
        return pd.DataFrame()
    
    # Find rows with any negative values
    anomaly_mask = (numeric_df < 0).any(axis=1)
    anomalies = df[anomaly_mask].copy()
    
    # Add information about which columns have negative values
    anomalies['negative_columns'] = numeric_df[anomaly_mask].apply(
        lambda row: ', '.join([col for col, val in row.items() if val < 0]), 
        axis=1
    )
    
    # Sort by container number if it exists
    if container_col in anomalies.columns:
        anomalies = anomalies.sort_values(by=container_col)
    
    return anomalies

# TAB 1: Container Comparison
with tab1:
    st.header("Container Spreadsheet Comparison")
    st.write("Compare container numbers between TOPS and Cyman spreadsheets to identify mismatches.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        tops_file = st.file_uploader("Upload TOPS spreadsheet (Excel or CSV)", type=["xlsx", "xls", "csv"])
    
    with col2:
        cyman_file = st.file_uploader("Upload Cyman spreadsheet (Excel or CSV)", type=["xlsx", "xls", "csv"])
    
    # Load the data
    tops_df = read_file(tops_file)
    cyman_df = read_file(cyman_file)
    
    # Determine container column
    if tops_df is not None and cyman_df is not None:
        tops_container_col = identify_container_column(tops_df)
        cyman_container_col = identify_container_column(cyman_df)
        
        if tops_container_col and cyman_container_col:
            if tops_container_col == cyman_container_col:
                detected_container_col = tops_container_col
                st.success(f"Container column automatically detected: '{detected_container_col}'")
            else:
                st.warning(f"Different container columns detected: TOPS: '{tops_container_col}', Cyman: '{cyman_container_col}'")
                detected_container_col = st.selectbox(
                    "Select which container column to use:",
                    [tops_container_col, cyman_container_col]
                )
        elif tops_container_col:
            detected_container_col = tops_container_col
            st.success(f"Container column detected in TOPS: '{detected_container_col}'")
        elif cyman_container_col:
            detected_container_col = cyman_container_col
            st.success(f"Container column detected in Cyman: '{detected_container_col}'")
        else:
            st.warning("Could not automatically detect container number column.")
            # Display available columns
            if tops_df is not None:
                st.write("Available columns in TOPS:")
                st.write(", ".join(tops_df.columns.tolist()))
            if cyman_df is not None:
                st.write("Available columns in Cyman:")
                st.write(", ".join(cyman_df.columns.tolist()))
            
            detected_container_col = None
        
        # Allow manual container column selection
        container_column = st.text_input(
            "Enter container column name (or confirm detected column):",
            value=detected_container_col if detected_container_col else ""
        )
        
        if st.button("Compare Containers"):
            if not container_column:
                st.error("Please enter a container column name.")
            elif tops_df is None:
                st.error("Please upload a TOPS spreadsheet.")
            elif cyman_df is None:
                st.error("Please upload a Cyman spreadsheet.")
            else:
                with st.spinner("Comparing container numbers..."):
                    results_df = compare_container_spreadsheets(tops_df, cyman_df, container_column)
                    
                    if results_df is not None:
                        if results_df.empty:
                            st.success("No mismatches found! All container numbers match between TOPS and Cyman.")
                        else:
                            st.write(f"Found {len(results_df)} mismatches:")
                            st.dataframe(results_df)
                            
                            # Create formatted Excel output
                            excel_data = io.BytesIO()
                            excel_data = format_excel_output(excel_data, results_df)
                            
                            # Add download button for the formatted Excel
                            st.download_button(
                                label="Download Excel Report",
                                data=excel_data,
                                file_name=f"container_mismatches_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error("Error comparing containers. Please check the data and try again.")

# TAB 2: Anomaly Detection
with tab2:
    st.header("File Anomaly Detector")
    st.write("Upload files to detect anomalies in container data.")
    
    # File uploader allowing multiple files (CSV and Excel)
    uploaded_files = st.file_uploader(
        "Upload CSV or Excel files (max 8)", 
        type=["csv", "xlsx", "xls"], 
        accept_multiple_files=True,
        key="anomaly_files"
    )
    
    # Limit to 8 files if more are uploaded
    if uploaded_files and len(uploaded_files) > 8:
        st.error("Please upload a maximum of 8 files.")
        uploaded_files = uploaded_files[:8]
    
    # Multi-select menu for selecting anomaly detection methods
    anomaly_methods = st.multiselect(
        "Select anomaly detection methods",
        ["MovementPre", "MovementIn", "In Activity", "Out Activity", "PreRelease", "Job Complete", "In Progress"],
        default=["MovementPre"]
    )
    
    # Process each uploaded file
    if uploaded_files:
        for file_index, uploaded_file in enumerate(uploaded_files):
            st.subheader(f"File: {uploaded_file.name}")
            try:
                # Read the file based on its extension
                df = read_file(uploaded_file)
                
                if df is None:
                    continue
                
                # Display file information
                st.write(f"Total rows: {len(df)}")
                st.write(f"Total columns: {len(df.columns)}")
                
                # Identify container column
                container_col = identify_container_column(df)
                
                if container_col:
                    st.write(f"Container number column identified: '{container_col}'")
                else:
                    st.warning("No container number column detected. Please enter the column name:")
                    container_col = st.text_input(
                        f"Container number column name for {uploaded_file.name}:", 
                        "container_number", 
                        key=f"container_col_{file_index}"
                    )
                
                # Display a preview with option to see more
                with st.expander(f"Preview data - {uploaded_file.name}"):
                    preview_rows = st.slider(
                        "Number of rows to preview", 
                        5, 100, 5, 
                        key=f"preview_slider_{file_index}"
                    )
                    st.dataframe(df.head(preview_rows))
                
            except Exception as e:
                st.error(f"Error reading {uploaded_file.name}: {e}")
                continue
            
            if not anomaly_methods:
                st.warning("Please select at least one anomaly detection method.")
            else:
                st.write("Anomaly Detection Results:")
                
                # Create a dictionary to store anomaly results for all methods
                all_anomalies = {}
                
                # Process each selected method
                for method_index, method in enumerate(anomaly_methods):
                    method_key = f"{file_index}_{method_index}"
                    st.write(f"### {method}")
                    
                    # Filter data based on the selected status if column exists
                    status_col = None
                    possible_status_cols = ['status', 'state', 'condition', 'stage']
                    
                    for col in possible_status_cols:
                        if col in df.columns:
                            status_col = col
                            break
                    
                    if status_col and method in df[status_col].values:
                        filtered_df = df[df[status_col] == method]
                        st.write(f"Found {len(filtered_df)} containers with status '{method}'.")
                    else:
                        # If no status column or value not found, use all data
                        filtered_df = df
                        st.write(f"Processing all {len(filtered_df)} rows for '{method}' anomalies.")
                    
                    # Store method-specific anomalies
                    method_anomalies = {
                        'Missing Values': detect_missing_values(filtered_df, container_col),
                        'Duplicate Rows': detect_duplicate_rows(filtered_df, container_col),
                        'Outliers': detect_outliers_zscore(filtered_df, container_col=container_col),
                        'Negative Values': detect_custom_anomaly(filtered_df, container_col)
                    }
                    
                    # Display results in tables
                    for anomaly_type_index, (anomaly_type, anomaly_df) in enumerate(method_anomalies.items()):
                        anomaly_key = f"{method_key}_{anomaly_type_index}"
                        if not anomaly_df.empty:
                            with st.expander(f"{anomaly_type} ({len(anomaly_df)} containers) - {method}"):
                                st.write(f"Found {len(anomaly_df)} containers with {anomaly_type.lower()}:")
                                
                                # Create a simplified view for the table with focus on container numbers
                                if container_col in anomaly_df.columns:
                                    # Determine which columns to show
                                    display_cols = [container_col]
                                    
                                    # Add status column if it exists
                                    if status_col in anomaly_df.columns:
                                        display_cols.append(status_col)
                                    
                                    # Add anomaly-specific columns
                                    if anomaly_type == 'Missing Values' and 'missing_in_columns' in anomaly_df.columns:
                                        display_cols.append('missing_in_columns')
                                    elif anomaly_type == 'Negative Values' and 'negative_columns' in anomaly_df.columns:
                                        display_cols.append('negative_columns')
                                    
                                    # Display the simplified table
                                    st.dataframe(anomaly_df[display_cols])
                                    
                                    # Option to view full details
                                    if st.button(
                                        f"View full details for {anomaly_type}", 
                                        key=f"view_details_{anomaly_key}"
                                    ):
                                        st.dataframe(anomaly_df)
                                else:
                                    # If container column not found, show all data
                                    st.dataframe(anomaly_df)
                                
                                # Save for aggregate view
                                key = f"{method} - {anomaly_type}"
                                all_anomalies[key] = anomaly_df
                    
                    # Add a separator between methods
                    if method != anomaly_methods[-1]:
                        st.markdown("---")
                
                # Aggregate view of all anomalies
                st.header("Container Anomaly Summary")
                
                # Create a set of all containers with anomalies
                all_container_anomalies = set()
                for anomaly_df in all_anomalies.values():
                    if not anomaly_df.empty and container_col in anomaly_df.columns:
                        all_container_anomalies.update(anomaly_df[container_col].astype(str).tolist())
                
                # Create a summary table of all containers and their anomalies
                if all_container_anomalies:
                    # Create a DataFrame for the summary
                    summary_data = []
                    
                    for container in sorted(all_container_anomalies):
                        container_anomalies = []
                        
                        for method_type, anomaly_df in all_anomalies.items():
                            if (not anomaly_df.empty and 
                                container_col in anomaly_df.columns and 
                                container in anomaly_df[container_col].astype(str).values):
                                container_anomalies.append(method_type)
                        
                        summary_data.append({
                            'Container Number': container,
                            'Anomaly Types': ', '.join(container_anomalies),
                            'Total Anomalies': len(container_anomalies)
                        })
                    
                    # Create and display the summary DataFrame
                    summary_df = pd.DataFrame(summary_data)
                    summary_df = summary_df.sort_values(by=['Total Anomalies', 'Container Number'], ascending=[False, True])
                    
                    st.write(f"Found {len(summary_df)} containers with anomalies:")
                    st.dataframe(summary_df)
                    
                    # Allow file download of the summary
                    csv = summary_df.to_csv(index=False)
                    st.download_button(
                        label=f"Download Container Anomaly Summary - {uploaded_file.name}",
                        data=csv,
                        file_name=f"container_anomaly_summary_{uploaded_file.name.split('.')[0]}.csv",
                        mime="text/csv",
                        key=f"download_button_{file_index}"
                    )
                else:
                    st.write("No container anomalies detected.")
