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
    Compare container numbers between TOPS and Cyman spreadsheets using the following prompt logic:
    
    Input Files:
      - TOPS: Focus on columns A (Status Name), H (Container Number), and O (Unload Location)
      - CYMAN: Focus on columns E (In Activity) and G (Unit No.)
    
    Filter Data:
      - TOPS: Keep only rows where Unload Location (Column O) contains "JAMES KEMBALL HOLDING CENTER" 
              (case-insensitive; allow minor typos)
      - CYMAN: Keep all rows
    
    Comparison Logic:
      For each row in the filtered TOPS data, check if the Container Number (Column H) exists in CYMAN’s Unit No. (Column G).
      If it DOES NOT match, retain the row. If it matches, exclude the row.
    
    Output:
      Return a table with the columns:
        - TOPS Container Number
        - CYMAN Unit No.
        - TOPS Unload Location
        - CYMAN In Activity
    """
    # Load spreadsheets
    try:
        tops_df = pd.read_excel(tops_file)
        cyman_df = pd.read_excel(cyman_file)
    except Exception as e:
        st.error(f"Error loading spreadsheets: {e}")
        return None, None, None

    # Subset required columns based on prompt instructions
    # TOPS: Columns A (Status Name), H (Container Number), O (Unload Location)
    tops_required = ["Status Name", "Container Number", "Unload Location"]
    if not all(col in tops_df.columns for col in tops_required):
        st.error("TOPS file missing one or more required columns: " + ", ".join(tops_required))
        return None, None, None
    tops_df = tops_df[tops_required]

    # CYMAN: Columns E (In Activity) and G (Unit No.)
    cyman_required = ["In Activity", "Unit No."]
    if not all(col in cyman_df.columns for col in cyman_required):
        st.error("CYMAN file missing one or more required columns: " + ", ".join(cyman_required))
        return None, None, None
    cyman_df = cyman_df[cyman_required]

    # Filter TOPS data:
    # Keep only rows where Unload Location (Column O) contains "JAMES KEMBALL HOLDING CENTER"
    # Allow for minor typos by using a regex that accepts both "center" and "centre"
    pattern = re.compile(r'james kemball holding cent(er|re)?', re.IGNORECASE)
    tops_df_filtered = tops_df[tops_df["Unload Location"].astype(str).apply(lambda x: bool(pattern.search(x)))]
    
    if tops_df_filtered.empty:
        st.info("No TOPS rows meet the Unload Location criteria.")
        empty_df = pd.DataFrame(columns=["TOPS Container Number", "CYMAN Unit No.", "TOPS Unload Location", "CYMAN In Activity"])
        output_buffer = io.BytesIO()
        empty_df.to_excel(output_buffer, index=False)
        output_buffer.seek(0)
        csv_buffer = io.StringIO()
        empty_df.to_csv(csv_buffer, index=False)
        csv_buffer.seek(0)
        return output_buffer, empty_df, csv_buffer

    # Normalize the container numbers (TOPS) and unit numbers (CYMAN)
    tops_df_filtered["Container Number"] = tops_df_filtered["Container Number"].astype(str).str.strip().str.upper()
    cyman_df["Unit No."] = cyman_df["Unit No."].astype(str).str.strip().str.upper()

    # Create a set of CYMAN Unit No. values for quick lookup
    cyman_units = set(cyman_df["Unit No."].tolist())

    # Compare each TOPS row: if the Container Number does NOT exist in CYMAN, retain the row.
    results = []
    for idx, row in tops_df_filtered.iterrows():
        container = row["Container Number"]
        unload_location = row["Unload Location"]
        if container not in cyman_units:
            results.append({
                "TOPS Container Number": container,
                "CYMAN Unit No.": "Not Found",
                "TOPS Unload Location": unload_location,
                "CYMAN In Activity": "Not Found"
            })

    # Create DataFrame from results
    results_df = pd.DataFrame(results)
    
    if results_df.empty:
        st.info("All filtered TOPS container numbers have matching entries in CYMAN.")
        empty_df = pd.DataFrame(columns=["TOPS Container Number", "CYMAN Unit No.", "TOPS Unload Location", "CYMAN In Activity"])
        output_buffer = io.BytesIO()
        empty_df.to_excel(output_buffer, index=False)
        output_buffer.seek(0)
        csv_buffer = io.StringIO()
        empty_df.to_csv(csv_buffer, index=False)
        csv_buffer.seek(0)
        return output_buffer, empty_df, csv_buffer

    # Write results to an in-memory Excel file
    output_buffer = io.BytesIO()
    results_df.to_excel(output_buffer, index=False)
    output_buffer.seek(0)

    # Create CSV buffer
    csv_buffer = io.StringIO()
    results_df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)

    st.success("Comparison complete!")
    return output_buffer, results_df, csv_buffer

def main():
    st.title("Container Spreadsheet Comparison Tool")

    # Upload file widgets
    tops_file = st.file_uploader("Upload TOPS spreadsheet", type=["xlsx", "xls"])
    cyman_file = st.file_uploader("Upload CYMAN spreadsheet", type=["xlsx", "xls"])

    # Text inputs for container column names (optional)
    tops_col = st.text_input("TOPS container column name (leave blank for auto-detection)")
    if tops_col.strip() == "":
        tops_col = None

    cyman_col = st.text_input("CYMAN container column name (leave blank for auto-detection)")
    if cyman_col.strip() == "":
        cyman_col = None
        
    # Option to check for single boxes
    check_single_boxes = st.checkbox("Check for single boxes across both yards")
    
    # Output format selection
    output_format = st.radio("Select output format:", ["View on screen", "Download as Excel", "Download as CSV"])

    if st.button("Run Comparison"):
        if not tops_file or not cyman_file:
            st.error("Please upload both TOPS and CYMAN spreadsheets.")
        else:
            excel_buffer, result_df, csv_buffer = compare_container_spreadsheets(
                tops_file, cyman_file, tops_col, cyman_col, check_single_boxes
            )
            
            if excel_buffer is not None and result_df is not None and csv_buffer is not None:
                if output_format == "View on screen":
                    st.subheader("Container Number Comparison Report")
                    st.caption("Comparison based on prompt instructions:")
                    st.caption("TOPS: Columns A (Status Name), H (Container Number), O (Unload Location)")
                    st.caption("CYMAN: Columns E (In Activity), G (Unit No.)")
                    st.caption(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    
                    display_df = format_for_display(result_df)
                    
                    st.markdown("""
                    <style>
                    .dataframe th {
                        background-color: #1F4E78;
                        color: white;
                        text-align: center;
                    }
                    .dataframe td {
                        text-align: center;
                    }
                    </style>
                    """, unsafe_allow_html=True)
                    
                    st.dataframe(display_df)
                    st.write(f"Total mismatches found: {len(result_df)}")
                    
                elif output_format == "Download as Excel":
                    st.download_button(
                        label="Download Comparison Report (Excel)",
                        data=excel_buffer,
                        file_name="container_comparison_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                elif output_format == "Download as CSV":
                    st.download_button(
                        label="Download Comparison Report (CSV)",
                        data=csv_buffer.getvalue(),
                        file_name="container_comparison_report.csv",
                        mime="text/csv"
                    )

if __name__ == "__main__":
    main()
