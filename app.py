import streamlit as st
import pandas as pd
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime
import io
import csv
from collections import Counter
import re

def identify_container_column(df):
    """
    Identify the container number column in the dataframe based on known terminology.
    """
    if df is None:
        return None

    possible_column_names = [
        'CONTAINER NUMBER', 'CONTAINER', 'CONTAINER NO', 'CONTAINER_NUMBER',
        'container', 'container_number', 'container_no', 'containerno', 'container_id',
        'Unit No', 'UNIT NO', 'unit no', 'unit_no', 'unitno', 'UNIT_NO'
    ]

    # Exact matches first
    for col_name in possible_column_names:
        if col_name in df.columns:
            return col_name

    # Then partial matches
    for col in df.columns:
        if 'container' in col.lower() or 'unit' in col.lower():
            return col

    return None

def identify_columns(df, system_type):
    """
    Identify relevant columns in the dataframes based on the system type.
    
    Args:
        df: The dataframe to analyze
        system_type: Either 'TOPS' or 'Cyman'
        
    Returns:
        dict: A dictionary of identified column names
    """
    identified_columns = {}
    
    if df is None:
        return identified_columns
    
    if system_type == 'TOPS':
        # Identify container column
        container_col = identify_container_column(df)
        identified_columns['container_col'] = container_col
        
        # Identify status column
        for col in df.columns:
            if any(status in str(col).lower() for status in ['status', 'state', 'progress']):
                identified_columns['status_col'] = col
                break
        
        # Identify location column
        for col in df.columns:
            if any(loc in str(col).lower() for loc in ['location', 'unload', 'terminal']):
                identified_columns['location_col'] = col
                break
    
    elif system_type == 'Cyman':
        # Identify container column
        container_col = identify_container_column(df)
        identified_columns['container_col'] = container_col
        
        # Identify activity column
        for col in df.columns:
            if any(act in str(col).lower() for act in ['activity', 'status', 'standard']):
                identified_columns['activity_col'] = col
                break
    
    return identified_columns

def format_excel_workbook(wb, row_count):
    """
    Apply formatting to the Excel workbook (in memory).
    """
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
        # Container Number column
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
    ws.column_dimensions[get_column_letter(1)].width = 20
    ws.column_dimensions[get_column_letter(2)].width = 12
    ws.column_dimensions[get_column_letter(3)].width = 12

    # Add a title row
    ws.insert_rows(1)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "Container Number Comparison Report"
    title_cell.font = Font(size=14, bold=True)
    ws.merge_cells('A1:C1')
    title_cell.alignment = Alignment(horizontal='center')

    # Add terminology note
    ws.insert_rows(2)
    terminology_cell = ws.cell(row=2, column=1)
    terminology_cell.value = "Note: TOPS 'Container Number' = Cyman 'Unit No'"
    terminology_cell.font = Font(italic=True)
    ws.merge_cells('A2:C2')
    terminology_cell.alignment = Alignment(horizontal='center')

    # Add timestamp
    ws.insert_rows(3)
    date_cell = ws.cell(row=3, column=1)
    date_cell.value = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws.merge_cells('A3:C3')
    date_cell.alignment = Alignment(horizontal='center')

    # Add summary at the bottom
    ws.append([])  # Empty row
    summary_row = ws.max_row + 1
    summary_cell = ws.cell(row=summary_row, column=1)
    summary_cell.value = f"Total mismatches found: {row_count}"
    summary_cell.font = Font(bold=True)
    ws.merge_cells(f'A{summary_row}:C{summary_row}')

def format_for_display(results_df):
    """
    Format the results DataFrame for on-screen display with the same style as the Excel output.
    """
    # Create a copy of the dataframe for display
    display_df = results_df.copy()
    
    # Replace values with symbols for display
    display_df = display_df.replace({'Missing': '❌', 'Present': '✓'})
    
    return display_df

def compare_container_spreadsheets(tops_file, cyman_file, tops_container_col=None, cyman_container_col=None, check_single_boxes=False):
    """
    Compare container numbers between TOPS and Cyman spreadsheets with enhanced logic.
    Returns a tuple containing:
    - BytesIO for Excel output
    - DataFrame for display
    - StringIO for CSV output
    """
    # Load spreadsheets
    try:
        tops_df = pd.read_excel(tops_file)
        cyman_df = pd.read_excel(cyman_file)
    except Exception as e:
        st.error(f"Error loading spreadsheets: {e}")
        return None, None, None

    # Identify columns in both spreadsheets
    tops_columns = identify_columns(tops_df, 'TOPS')
    cyman_columns = identify_columns(cyman_df, 'Cyman')
    
    # Auto-detect container columns if not provided
    if tops_container_col is None:
        tops_container_col = tops_columns.get('container_col')
        if tops_container_col:
            st.write(f"TOPS container column detected: {tops_container_col}")
        else:
            st.error("Could not automatically detect container number column in TOPS. Available columns:")
            st.write(list(tops_df.columns))
            return None, None, None

    if cyman_container_col is None:
        cyman_container_col = cyman_columns.get('container_col')
        if cyman_container_col:
            st.write(f"Cyman container column detected: {cyman_container_col}")
        else:
            st.error("Could not automatically detect container number column in Cyman. Available columns:")
            st.write(list(cyman_df.columns))
            return None, None, None
            
    # Initialize columns for status and location/activity
    tops_status_col = tops_columns.get('status_col', None)
    tops_location_col = tops_columns.get('location_col', None)
    cyman_activity_col = cyman_columns.get('activity_col', None)
    
    # If we couldn't detect columns automatically, look for them by position based on the sample data
    if tops_status_col is None and len(tops_df.columns) >= 1:
        tops_status_col = tops_df.columns[0]  # First column in TOPS sample
        st.write(f"Using first column as TOPS status: {tops_status_col}")
    
    if tops_location_col is None and len(tops_df.columns) >= 3:
        tops_location_col = tops_df.columns[2]  # Third column in TOPS sample
        st.write(f"Using third column as TOPS location: {tops_location_col}")
    
    if cyman_activity_col is None and len(cyman_df.columns) >= 6:
        cyman_activity_col = cyman_df.columns[5]  # Sixth column in Cyman sample
        st.write(f"Using sixth column as Cyman activity: {cyman_activity_col}")

    try:
        # Convert container columns to strings and strip whitespace
        tops_df[tops_container_col] = tops_df[tops_container_col].astype(str).str.strip()
        cyman_df[cyman_container_col] = cyman_df[cyman_container_col].astype(str).str.strip()
        
        # Filter TOPS data to only include "Complete" or "In Progress" in Status Name column
       
