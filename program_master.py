import pandas as pd
import os
import streamlit as st

@st.cache_data
def load_program_master():
    """
    Loads program_master.xlsx from the "Sheet3" sheet.
    It directly uses the SECOND column (index 1) as the unique key.
    If duplicates exist, they are grouped into a list for the user to choose from later.
    Returns a tuple: (data_dictionary, error_message)
    """
    file_path = "program_master.xlsx"
    if not os.path.exists(file_path):
        return None, f"Error: Program master file '{file_path}' not found."

    try:
        # Load the excel file
        df = pd.read_excel(file_path, sheet_name="Sheet3", header=0)

        # Check if the DataFrame has at least two columns
        if len(df.columns) < 2:
            return None, "Error: The target sheet in program_master.xlsx must have at least two columns."

        # --- DIRECTLY SELECT THE KEY COLUMN ---
        key_column_name = df.columns[1]
        
        # Clean the keys to prevent fake duplicates caused by hidden spaces
        df[key_column_name] = df[key_column_name].astype(str).str.strip()
        df = df.fillna('')
        
        # --- GROUP DUPLICATES ---
        # Instead of throwing a fatal error, we group by the key column.
        program_dict = df.groupby(key_column_name).apply(lambda x: x.to_dict(orient='records')).to_dict()

        return program_dict, None
        
    except Exception as e:
        return None, f"Error loading '{file_path}': {e}"


def check_missing_programs(source_program_nos: list, master_dict: dict):
    """
    Checks which program numbers from a source list are missing in the master dictionary.
    """
    if not master_dict: 
        return sorted(list(set(source_program_nos)))
    master_keys = set(master_dict.keys())
    source_keys = set(map(str, source_program_nos))
    missing_keys = sorted(list(source_keys - master_keys))
    return missing_keys


def get_program_details(prog_no, master_dict):
    """
    Retrieves the details for a given program number.
    If duplicates exist, it prompts the user to select one via a Streamlit dropdown.
    Returns a single dictionary of the selected row, or None if not found.
    """
    # Convert input to string and strip whitespace to ensure a clean match
    clean_prog_no = str(prog_no).strip()
    
    if not master_dict or clean_prog_no not in master_dict:
        return None

    # Fetch the list of records for this prog_no
    records = master_dict[clean_prog_no]

    # Scenario A: Only one entry exists. Return it directly.
    if len(records) == 1:
        return records[0]

    # Scenario B: Duplicates exist. Ask the user to choose.
    st.warning(f"⚠️ Multiple entries found for Program No: **{clean_prog_no}**. Please select which one to use for this session.")
    
    # Create readable options for the dropdown
    options = []
    for i, rec in enumerate(records):
        # Checks for uppercase first, then lowercase, then defaults to 'N/A'
        abbr = rec.get('ABBR', rec.get('abbr', 'N/A'))
        degnm = rec.get('DEGNM', rec.get('degnm', 'N/A'))
        
        options.append(f"Option {i+1} | Abbr: {abbr} | Degree: {degnm}")
    
    # Render the Streamlit selectbox
    selected_option = st.selectbox(
        "Select Entry:", 
        options, 
        key=f"duplicate_select_{clean_prog_no}" # Unique key is required
    )
    
    # Return the specific dictionary the user selected
    selected_index = options.index(selected_option)
    return records[selected_index]