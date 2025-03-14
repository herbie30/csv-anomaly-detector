import pandas as pd
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime

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
    Try to identify the container number column in the dataframe based on known terminology.
    
    Parameters:
    df (DataFrame): The dataframe to analyze
    
    Returns:
    str: The identified container column name or None if not found
    """
    if df is None:
        return None
    
    # Common container column names (including both TOPS and Cyman terminology)
    possible_column_names = [
        'CONTAINER NUMBER', 'CONTAINER', 'CONTAINER NO', 'CONTAINER_NUMBER',
        'container', 'container_number', 'container_no', 'containerno', 'container_id',
        'Unit No', 'UNIT NO', 'unit no', 'unit_no', 'unitno', 'UNIT_NO'
    ]
    
    # Check exact matches first
    for col_name in possible_column_names:
        if col_name in df.columns:
            return col_name
    
    # Then check for partial matches
    for col in df.columns:
        if 'container' in col.lower() or 'unit' in col.lower():
            return col
    
    return None

def compare_container_spreadsheets(tops_file, cyman_file, tops_container_col=None, cyman_container_col=None, output_file=None):
    """
    Compare container numbers between TOPS and Cyman spreadsheets and identify mismatches.
    Outputs results to a nicely formatted Excel spreadsheet.
    
    Parameters:
    tops_file (str): Path to the TOPS spreadsheet
    cyman_file (str): Path to the Cyman spreadsheet
    tops_container_col (str): Name of the column containing container numbers in TOPS spreadsheet
    cyman_container_col (str): Name of the column containing container numbers in Cyman spreadsheet
    output_file (str): Path for the output Excel file
    
    Returns:
    str: Path to the created output file
    """
    # Check if files exist
    if not os.path.exists(tops_file):
        print(f"Error: TOPS file not found at {tops_file}")
        return None
    
    if not os.path.exists(cyman_file):
        print(f"Error: Cyman file not found at {cyman_file}")
        return None
    
    # Load spreadsheets
    try:
        tops_df = pd.read_excel(tops_file)
        cyman_df = pd.read_excel(cyman_file)
    except Exception as e:
        print(f"Error loading spreadsheets: {e}")
        return None
    
    # Auto-detect container columns if not specified
    if tops_container_col is None:
        tops_container_col = identify_container_column(tops_df)
        if tops_container_col:
            print(f"TOPS container column detected: {tops_container_col}")
        else:
            print("Could not automatically detect container number column in TOPS.")
            print("Available columns in TOPS:")
            for col in tops_df.columns:
                print(f"  - {col}")
            return None
    
    if cyman_container_col is None:
        cyman_container_col = identify_container_column(cyman_df)
        if cyman_container_col:
            print(f"Cyman container column detected: {cyman_container_col}")
        else:
            print("Could not automatically detect container number column in Cyman.")
            print("Available columns in Cyman:")
            for col in cyman_df.columns:
                print(f"  - {col}")
            return None
    
    # Extract container numbers, making sure they're treated as strings
    try:
        tops_containers = set(tops_df[tops_container_col].astype(str).str.strip())
        cyman_containers = set(cyman_df[cyman_container_col].astype(str).str.strip())
    except KeyError:
        print(f"Error: Columns not found - TOPS: '{tops_container_col}', Cyman: '{cyman_container_col}'")
        return None
    
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
        print("No mismatches found. All container numbers match between TOPS and Cyman.")
        empty_df = pd.DataFrame(columns=['CONTAINER NUMBER', 'CYMAN', 'TOPS'])
        if output_file:
            empty_df.to_excel(output_file, index=False)
            return output_file
        return None
    
    results_df = pd.DataFrame(results)
    
    # Set default output file name if not provided
    if not output_file:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"container_mismatches_{timestamp}.xlsx"
    
    # Export to Excel
    results_df.to_excel(output_file, index=False)
    
    # Apply formatting to the Excel file
    format_excel_output(output_file, len(results_df))
    
    print(f"Results saved to {output_file}")
    return output_file

def format_excel_output(file_path, row_count):
    """
    Format the Excel output file with colors and styling
    """
    # Open the workbook
    wb = openpyxl.load_workbook(file_path)
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
    
    # Add summary
    ws.append([])  # Empty row
    summary_row = ws.max_row + 1
    summary_cell = ws.cell(row=summary_row, column=1)
    summary_cell.value = f"Total mismatches found: {row_count}"
    summary_cell.font = Font(bold=True)
    ws.merge_cells(f'A{summary_row}:C{summary_row}')
    
    # Save the workbook
    wb.save(file_path)

def display_terminology_mapping():
    """
    Display the terminology mapping between TOPS and Cyman
    """
    print("\nTerminology Mapping between TOPS and Cyman:")
    print("=" * 50)
    print(f"{'TOPS Terms':<25} {'CYMAN Terms':<25}")
    print("-" * 50)
    for tops_term, cyman_term in TERMINOLOGY_MAPPING.items():
        print(f"{tops_term:<25} {cyman_term:<25}")
    print("=" * 50)
    print()

if __name__ == "__main__":
    print("Container Spreadsheet Comparison Tool")
    print("=====================================")
    print("This tool compares container numbers between TOPS and Cyman systems.")
    
    # Display terminology mapping
    display_terminology_mapping()
    
    # Get input file paths
    tops_file = input("Enter path to TOPS spreadsheet: ")
    cyman_file = input("Enter path to Cyman spreadsheet: ")
    
    # Optionally allow specifying container column names
    print("\nContainer number column names (leave blank for auto-detection):")
    tops_col = input(f"TOPS container column name (default: 'Container Number'): ").strip()
    tops_col = tops_col if tops_col else None
    
    cyman_col = input(f"Cyman container column name (default: 'Unit No'): ").strip()
    cyman_col = cyman_col if cyman_col else None
    
    # Optionally specify output file
    output_file = input("\nEnter output Excel file name (or press Enter for auto-generated name): ").strip()
    output_file = output_file if output_file else None
    
    # Run the comparison
    result_file = compare_container_spreadsheets(tops_file, cyman_file, tops_col, cyman_col, output_file)
    
    if result_file:
        print(f"\nComparison complete! Results saved to: {result_file}")
        print("\nThe Excel file contains:")
        print("- Color-coded indicators (red ❌ for missing, green ✓ for present)")
        print("- Summary of total mismatches")
        print("- Terminology mapping note")
    else:
        if tops_file and cyman_file:  # No mismatches and files were specified
            print("\nNo output file created as there were no mismatches to report.")
        else:
            print("\nOperation could not be completed. Please check the inputs and try again.")
