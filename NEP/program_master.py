import pandas as pd
import os
import streamlit as st

@st.cache_data
def load_program_master():
    """
    Loads program_master.xlsx from "Sheet3".
    Uses the FIRST column (index 0) as the unique Program ID key (PROG_NO).
    Returns a tuple: (data_dictionary, error_message)
    """
    # 1. Check if file exists
    dir_path = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(dir_path, "program_master.xlsx")
    if not os.path.exists(file_path):
        return {}, f"Error: Master file '{file_path}' not found in the folder."

    try:
        # 2. Load the excel file (Sheet3)
        # Using header=0 to treat the first row as column names
        df = pd.read_excel(file_path, sheet_name="Sheet3", header=0)

        # 3. Validation: Must have at least two columns
        if len(df.columns) < 2:
            return {}, "Error: 'Sheet3' must have at least two columns (PROG_NO, ABBR, etc)."

        # 4. Identify the Key Column (The 1st Column - PROG_NO)
        # CHANGED FROM index 1 TO index 0
        key_col = df.columns[0]
        
        # 5. Data Cleaning
        # Convert the key column to string, remove spaces, and drop rows where key is empty
        df[key_col] = df[key_col].astype(str).str.strip()
        
        # Remove empty rows or rows that became 'nan' string
        df = df[df[key_col] != 'nan']
        df = df[df[key_col] != '']
        
        # 6. Check for Duplicates
        if not df[key_col].is_unique:
            dup_val = df[df.duplicated(subset=[key_col])][key_col].iloc[0]
            return {}, f"FATAL ERROR: Duplicate Program ID '{dup_val}' found in master file."

        # 7. Fill empty cells with empty strings to prevent 'nan' appearing in final output
        df = df.fillna('')

        # 8. Convert to Dictionary for high-speed lookups
        # Format: { '3335261': {'ABBR': 'BAFN', ...}, ... }
        program_dict = df.set_index(key_col).to_dict('index')

        return program_dict, None
        
    except Exception as e:
        return {}, f"Error reading '{file_path}': {str(e)}"

def check_missing_programs(uploaded_filenames, master_dict):
    """
    Helper to check which uploaded files don't have a matching entry in the Master.
    """
    if not master_dict:
        return uploaded_filenames
    
    missing = []
    for fname in uploaded_filenames:
        # Extract ID from filename (e.g., "3335261.xlsx" -> "3335261")
        prog_id = os.path.splitext(fname)[0].strip()
        
        # Check if this ID exists in the master dictionary keys
        if prog_id not in master_dict:
            missing.append(prog_id)
            
    return sorted(list(set(missing)))