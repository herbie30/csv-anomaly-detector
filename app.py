import streamlit as st
import pandas as pd
import numpy as np
from scipy import stats

# App title and instructions
st.title("CSV Anomaly Detector")
st.write("Upload up to 8 CSV files, select anomaly detection methods, and preview the results.")

# CSV uploader allowing multiple files (limit to CSV)
uploaded_files = st.file_uploader("Upload CSV files (max 8)", type="csv", accept_multiple_files=True)

# Limit to 8 files if more are uploaded
if uploaded_files and len(uploaded_files) > 8:
    st.error("Please upload a maximum of 8 files.")
    uploaded_files = uploaded_files[:8]

# Multi-select menu for selecting anomaly detection methods
anomaly_methods = st.multiselect(
    "Select anomaly detection methods",
    ["Missing Values", "Duplicate Rows", "Outliers (Z-score)", "Custom Anomaly"],
    default=["Missing Values"]  # Default selection
)

# Define anomaly detection functions
def detect_missing_values(df):
    missing = df.isnull().sum()
    return missing[missing > 0]

def detect_duplicate_rows(df):
    duplicates = df[df.duplicated()]
    return duplicates

def detect_outliers_zscore(df, threshold=3):
    numeric_df = df.select_dtypes(include=[np.number])
    if numeric_df.empty:
        return pd.DataFrame()  # No numeric columns to process.
    z_scores = np.abs(stats.zscore(numeric_df, nan_policy='omit'))
    # Flag rows where any column's z-score exceeds the threshold.
    outlier_mask = (z_scores > threshold).any(axis=1)
    return df[outlier_mask]

def detect_custom_anomaly(df):
    # Example custom logic: detect rows where any numeric value is negative.
    numeric_df = df.select_dtypes(include=[np.number])
    if numeric_df.empty:
        return pd.DataFrame()
    anomaly_mask = (numeric_df < 0).any(axis=1)
    return df[anomaly_mask]

# Process each uploaded file
if uploaded_files:
    for uploaded_file in uploaded_files:
        st.subheader(f"File: {uploaded_file.name}")
        try:
            df = pd.read_csv(uploaded_file)
        except Exception as e:
            st.error(f"Error reading {uploaded_file.name}: {e}")
            continue

        st.write("Preview (first 5 rows):")
        st.dataframe(df.head())

        if not anomaly_methods:
            st.warning("Please select at least one anomaly detection method.")
        else:
            st.write("Anomaly Detection Results:")
            
            # Process each selected method
            for method in anomaly_methods:
                st.write(f"### {method}")
                
                if method == "Missing Values":
                    missing = detect_missing_values(df)
                    if missing.empty:
                        st.write("No missing values detected.")
                    else:
                        st.write("Missing values detected (count per column):")
                        st.write(missing)
                
                elif method == "Duplicate Rows":
                    duplicates = detect_duplicate_rows(df)
                    if duplicates.empty:
                        st.write("No duplicate rows detected.")
                    else:
                        st.write("Duplicate rows found:")
                        st.dataframe(duplicates)
                
                elif method == "Outliers (Z-score)":
                    outliers = detect_outliers_zscore(df)
                    if outliers.empty:
                        st.write("No outliers detected.")
                    else:
                        st.write("Outliers detected based on z-score analysis:")
                        st.dataframe(outliers)
                
                elif method == "Custom Anomaly":
                    anomalies = detect_custom_anomaly(df)
                    if anomalies.empty:
                        st.write("No custom anomalies detected.")
                    else:
                        st.write("Custom anomalies detected (rows with negative numeric values):")
                        st.dataframe(anomalies)
                
                # Add a separator between methods
                if method != anomaly_methods[-1]:
                    st.markdown("---")
