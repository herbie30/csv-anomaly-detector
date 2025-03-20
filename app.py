import streamlit as st
import pandas as pd

st.title("Container Comparison Tool")

# Upload files
tops_file = st.file_uploader("Upload TOPS Spreadsheet", type=["xlsx"])
cyman_file = st.file_uploader("Upload CYMAN Spreadsheet", type=["xlsx"])

if tops_file is not None and cyman_file is not None:
    # Read the spreadsheets
    tops_df = pd.read_excel(tops_file)
    cyman_df = pd.read_excel(cyman_file)

    # Display the original column names for verification
    st.write("### TOPS Columns", tops_df.columns.tolist())
    st.write("### CYMAN Columns", cyman_df.columns.tolist())

    # Normalize column names to lowercase and strip extra spaces
    tops_df.columns = tops_df.columns.str.lower().str.strip()
    cyman_df.columns = cyman_df.columns.str.lower().str.strip()

    # Check if the necessary columns exist in TOPS
    if 'container number' not in tops_df.columns:
        st.error("The TOPS file does not have a 'container number' column. Please verify the column names.")
    elif 'status name' not in tops_df.columns or 'unload location' not in tops_df.columns:
        st.error("One or more required columns ('status name', 'unload location') are missing in the TOPS file.")
    elif 'unit no' not in cyman_df.columns or 'in activity' not in cyman_df.columns or 'in haulier' not in cyman_df.columns:
        st.error("One or more required columns ('unit no', 'in activity', 'in haulier') are missing in the CYMAN file.")
    else:
        # Clean and standardize string columns
        tops_df['container number'] = tops_df['container number'].astype(str).str.strip()
        tops_df['status name'] = tops_df['status name'].astype(str).str.lower().str.strip()
        tops_df['unload location'] = tops_df['unload location'].astype(str).str.upper().str.strip()

        cyman_df['unit no'] = cyman_df['unit no'].astype(str).str.strip()
        cyman_df['in activity'] = cyman_df['in activity'].astype(str).str.lower().str.strip()
        cyman_df['in haulier'] = cyman_df['in haulier'].astype(str).str.upper().str.strip()

        # Filter TOPS: select rows with job complete and the specific unload location
        tops_filtered = tops_df[
            (tops_df['status name'] == "job complete") &
            (tops_df['unload location'] == "JAMES KEMBALL HOLDING CENTER")
        ]

        # Filter CYMAN: select rows with in activity as standard, unit no present, and in haulier as KEMBALL
        cyman_filtered = cyman_df[
            (cyman_df['in activity'] == "standard") &
            (cyman_df['unit no'].notnull()) &
            (cyman_df['in haulier'] == "KEMBALL")
        ]

        # Create sets of container/unit numbers for comparison
        tops_set = set(tops_filtered['container number'])
        cyman_set = set(cyman_filtered['unit no'])

        # Identify mismatches: those in TOPS but not in CYMAN, and vice versa
        missing_in_cyman = tops_set - cyman_set
        missing_in_tops = cyman_set - tops_set

        # Build a summary DataFrame for the differences
        results = []

        for container in missing_in_cyman:
            results.append({
                "Container/Unit No": container,
                "Source System": "TOPS",
                "Status / In Activity": "Job Complete / N/A",
                "Unload Location / In Haulier": "JAMES KEMBALL HOLDING CENTER / (Missing in CYMAN)",
                "Notes": "Missing in CYMAN"
            })

        for unit in missing_in_tops:
            results.append({
                "Container/Unit No": unit,
                "Source System": "CYMAN",
                "Status / In Activity": "N/A / Standard",
                "Unload Location / In Haulier": "(Missing in TOPS) / KEMBALL",
                "Notes": "Missing in TOPS"
            })

        result_df = pd.DataFrame(results)

        st.write("### Comparison Results")
        st.dataframe(result_df)
else:
    st.write("Please upload both spreadsheets to run the comparison.")
