import streamlit as st
import pandas as pd
import numpy as np
from scipy import stats

# App title and instructions
st.title("File Anomaly Detector")
st.write("Upload up to 8 CSV or Excel files, select anomaly detection methods, and view container anomalies in order.")

# File uploader allowing multiple files (CSV and Excel)
uploaded_files = st.file_uploader("Upload CSV or Excel files (max 8)", type=["csv", "xlsx", "xls"], accept_multiple_files=True)

# Limit to 8 files if more are uploaded
if uploaded_files and len(uploaded_files) > 8:
    st.error("Please upload a maximum of 8 files.")
    uploaded_files = uploaded_files[:8]

# Multi-select menu for selecting anomaly detection methods with updated labels
anomaly_methods = st.multiselect(
    "Select anomaly detection methods",
    ["MovementPre", "MovementIn", "In Activity", "Out Activity", "PreRelease", "Job Complete", "In Progress"],
    default=["MovementPre"]
)

# Define anomaly detection functions
def detect_missing_values(df, container_col='container_number'):
    """Detect missing values and return sorted by container number."""
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
    duplicates = df[df.duplicated(keep=False)]
    
    # Sort by container number if it exists
    if not duplicates.empty and container_col in duplicates.columns:
        duplicates = duplicates.sort_values(by=container_col)
    
    return duplicates

def detect_outliers_zscore(df, threshold=3, container_col='container_number'):
    """Detect outliers using Z-score and return sorted by container number."""
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

# Function to read file based on its extension
def read_file(file):
    """Read a file based on its extension (CSV or Excel)."""
    file_name = file.name.lower()
    
    if file_name.endswith('.csv'):
        return pd.read_csv(file)
    elif file_name.endswith(('.xlsx', '.xls')):
        return pd.read_excel(file)
    else:
        raise ValueError(f"Unsupported file format: {file_name}")

# Process each uploaded file
if uploaded_files:
    for uploaded_file in uploaded_files:
        st.subheader(f"File: {uploaded_file.name}")
        try:
            # Read the file based on its extension
            df = read_file(uploaded_file)
            
            # Display file information
            st.write(f"Total rows: {len(df)}")
            st.write(f"Total columns: {len(df.columns)}")
            
            # Try to identify the container number column
            container_col = None
            possible_container_cols = ['container', 'container_number', 'container_no', 'containerno', 'container_id']
            for col in possible_container_cols:
                if col in df.columns:
                    container_col = col
                    break
            
            if container_col is None:
                # Look for columns that might contain container information
                for col in df.columns:
                    if 'container' in col.lower():
                        container_col = col
                        break
            
            if container_col:
                st.write(f"Container number column identified: '{container_col}'")
            else:
                st.warning("No container number column detected. Please enter the column name:")
                container_col = st.text_input("Container number column name:", "container_number")
            
            # Display a preview with option to see more
            with st.expander("Preview data"):
                preview_rows = st.slider("Number of rows to preview", 5, 100, 5)
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
            for method in anomaly_methods:
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
                for anomaly_type, anomaly_df in method_anomalies.items():
                    if not anomaly_df.empty:
                        with st.expander(f"{anomaly_type} ({len(anomaly_df)} containers)"):
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
                                if st.button(f"View full details for {anomaly_type}", key=f"{method}_{anomaly_type}"):
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
                    label="Download Container Anomaly Summary",
                    data=csv,
                    file_name="container_anomaly_summary.csv",
                    mime="text/csv"
                )
            else:
                st.write("No container anomalies detected.")
  
   
   

  
            
            
           
                
