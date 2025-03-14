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
    Compare container numbers between TOPS and Cyman spreadsheets.
    Returns a BytesIO stream with the output Excel file.
    """
    # Load spreadsheets (tops_file and cyman_file can be file-like objects from Streamlit)
    try:
        tops_df = pd.read_excel(tops_file)
        cyman_df = pd.read_excel(cyman_file)
    except Exception as e:
        st.error(f"Error loading spreadsheets: {e}")
        return None

    # Auto-detect container columns if not provided
    if tops_container_col is None:
        tops_container_col = identify_container_column(tops_df)
        if tops_container_col:
            st.write(f"TOPS container column detected: {tops_container_col}")
        else:
            st.error("Could not automatically detect container number column in TOPS. Available columns:")
            st.write(list(tops_df.columns))
            return None

    if cyman_container_col is None:
        cyman_container_col = identify_container_column(cyman_df)
        if cyman_container_col:
            st.write(f"Cyman container column detected: {cyman_container_col}")
        else:
            st.error("Could not automatically detect container number column in Cyman. Available columns:")
            st.write(list(cyman_df.columns))
            return None

    try:
        tops_containers = set(tops_df[tops_container_col].astype(str).str.strip())
        cyman_containers = set(cyman_df[cyman_container_col].astype(str).str.strip())
    except KeyError:
        st.error(f"Error: Columns not found - TOPS: '{tops_container_col}', Cyman: '{cyman_container_col}'")
        return None

    # Find mismatches
    tops_only = tops_containers - cyman_containers
    cyman_only = cyman_containers - tops_containers

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
        st.info("No mismatches found. All container numbers match between TOPS and Cyman.")
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

    st.success("Comparison complete!")
    return final_buffer

def display_terminology_mapping():
    st.write("### Terminology Mapping between TOPS and Cyman:")
    st.write("=" * 50)
    for tops_term, cyman_term in TERMINOLOGY_MAPPING.items():
        st.write(f"{tops_term:<25} {cyman_term:<25}")
    st.write("=" * 50)

def main():
    st.title("Container Spreadsheet Comparison Tool")
    st.write("This tool compares container numbers between TOPS and Cyman systems.")

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

    if st.button("Run Comparison"):
        if not tops_file or not cyman_file:
            st.error("Please upload both TOPS and Cyman spreadsheets.")
        else:
            result_buffer = compare_container_spreadsheets(tops_file, cyman_file, tops_col, cyman_col)
            if result_buffer:
                st.download_button(
                    label="Download Comparison Report",
                    data=result_buffer,
                    file_name="container_comparison_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
