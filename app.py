import streamlit as st
import pandas as pd
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime
import io

# Dictionary of equivalent terms between TOPS and Cyman
TERMINOLOGY_MAPPING = {
    # Container related terms
    "Container Number": "Unit No",
    
    # Location terms
    "Unload Location": "In Activity",
    
    # Status terms
    "Job Complete": "In Activity",
    "In Progress": "MovementPre"
}

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

def compare_container_spreadsheets(tops_file, cyman_file, tops_container_col=None, cyman_container_col=None):
    """
    Compare container numbers between TOPS and Cyman spreadsheets with enhanced logic.
    Returns a BytesIO stream with the output Excel file.
    """
    # Load spreadsheets
    try:
        tops_df = pd.read_excel(tops_file)
        cyman_df = pd.read_excel(cyman_file)
    except Exception as e:
        st.error(f"Error loading spreadsheets: {e}")
        return None

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
            return None

    if cyman_container_col is None:
        cyman_container_col = cyman_columns.get('container_col')
        if cyman_container_col:
            st.write(f"Cyman container column detected: {cyman_container_col}")
        else:
            st.error("Could not automatically detect container number column in Cyman. Available columns:")
            st.write(list(cyman_df.columns))
            return None
            
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
        
        # Create dictionaries for faster lookups
        tops_containers = set(tops_df[tops_container_col])
        cyman_containers = set(cyman_df[cyman_container_col])
        
        tops_data = {row[tops_container_col]: row for _, row in tops_df.iterrows()}
        cyman_data = {row[cyman_container_col]: row for _, row in cyman_df.iterrows()}
    except KeyError as e:
        st.error(f"Error: Column not found - {e}")
        return None

    # Process containers based on the enhanced requirements
    results = []
    
    # Process containers in TOPS but not in Cyman
    for container in tops_containers - cyman_containers:
        results.append({
            'CONTAINER NUMBER': container,
            'CYMAN': 'Missing',
            'TOPS': 'Present'
        })
    
    # Process containers in Cyman but not in TOPS
    for container in cyman_containers - tops_containers:
        results.append({
            'CONTAINER NUMBER': container,
            'CYMAN': 'Present',
            'TOPS': 'Missing'
        })
    
    # Process containers that appear in both systems
    for container in tops_containers.intersection(cyman_containers):
        tops_row = tops_data[container]
        cyman_row = cyman_data[container]
        
        # Skip if container has "James Kemball Holding Centre" in TOPS and "Standard" in Cyman
        tops_location = str(tops_row.get(tops_location_col, "")) if tops_location_col else ""
        cyman_activity = str(cyman_row.get(cyman_activity_col, "")) if cyman_activity_col else ""
        tops_status = str(tops_row.get(tops_status_col, "")) if tops_status_col else ""
        
        # Skip matching containers with specified location/activity
        if "james kemball holding centre" in tops_location.lower() and "standard" in cyman_activity.lower():
            continue
        
        # Check for special case: "Standard" in Cyman and "In Progress" in TOPS
        if "standard" in cyman_activity.lower() and "in progress" in tops_status.lower():
            results.append({
                'CONTAINER NUMBER': container,
                'CYMAN': 'Present',
                'TOPS': 'Missing'
            })
        # Otherwise, this is a mismatch to report
        else:
            results.append({
                'CONTAINER NUMBER': container,
                'CYMAN': 'Present',
                'TOPS': 'Present'
            })

    if not results:
        st.info("No mismatches found based on the specified criteria.")
        empty_df = pd.DataFrame(columns=['CONTAINER NUMBER', 'CYMAN', 'TOPS'])
        output_buffer = io.BytesIO()
        empty_df.to_excel(output_buffer, index=False)
        output_buffer.seek(0)
        return output_buffer

    results_df = pd.DataFrame(results)

    # Write results to an in-memory Excel file
    output_buffer = io.BytesIO()
    results_df.to_excel(output_buffer, index=False)
    output_buffer.seek(0)

    # Load workbook from the BytesIO buffer
    try:
        wb = openpyxl.load_workbook(output_buffer)
    except Exception as e:
        st.error(f"Error processing workbook: {e}")
        return None

    # Apply formatting
    format_excel_workbook(wb, len(results_df))

    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)

    st.success("Enhanced comparison complete!")
    return final_buffer

def display_terminology_mapping():
    st.write("### Terminology Mapping between TOPS and Cyman:")
    st.write("=" * 50)
    for tops_term, cyman_term in TERMINOLOGY_MAPPING.items():
        st.write(f"{tops_term:<25} {cyman_term:<25}")
    st.write("=" * 50)

def main():
    st.title("Container Spreadsheet Comparison Tool")
    st.write("This tool compares container numbers between TOPS and Cyman systems with enhanced logic.")

    display_terminology_mapping()

    # Upload file widgets
    tops_file = st.file_uploader("Upload TOPS spreadsheet", type=["xlsx", "xls"])
    cyman_file = st.file_uploader("Upload Cyman spreadsheet", type=["xlsx", "xls"])

    # Text inputs for container column names (optional)
    tops_col = st.text_input("TOPS container column name (leave blank for auto-detection)")
    if tops_col.strip() == "":
        tops_col = None

    cyman_col = st.text_input("Cyman container column name (leave blank for auto-detection)")
    if cyman_col.strip() == "":
        cyman_col = None

    if st.button("Run Enhanced Comparison"):
        if not tops_file or not cyman_file:
            st.error("Please upload both TOPS and Cyman spreadsheets.")
        else:
            result_buffer = compare_container_spreadsheets(tops_file, cyman_file, tops_col, cyman_col)
            if result_buffer:
                st.download_button(
                    label="Download Comparison Report",
                    data=result_buffer,
                    file_name="enhanced_container_comparison_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
