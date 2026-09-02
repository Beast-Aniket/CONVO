import os
import pandas as pd
import streamlit as st

def get_program_master_path(filename="program_master.xlsx"):
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    candidates = [
        os.path.join(base_dir, "data", filename),
        os.path.join(base_dir, filename),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), filename),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return candidates[0]

def load_program_master(filename="program_master.xlsx"):
    """
    Loads program master Excel file into a fast dictionary keyed by Program ID (PROG_NO / ABBR).
    Returns (dict, error_message).
    """
    file_path = get_program_master_path(filename)
    
    if not os.path.exists(file_path):
        return {}, f"Program master file '{filename}' not found."

    try:
        # Load the excel file
        # Check sheet names first
        excel_file = pd.ExcelFile(file_path)
        sheet_to_use = "Sheet3" if "Sheet3" in excel_file.sheet_names else excel_file.sheet_names[0]
        
        df = pd.read_excel(file_path, sheet_name=sheet_to_use, header=0, dtype=str)

        if len(df.columns) < 2:
            return {}, f"Error: '{sheet_to_use}' must have at least two columns."

        # The first column is the program key (PROG_NO)
        key_col = df.columns[0]
        df = df.copy()
        df.loc[:, key_col] = df[key_col].str.strip()
        
        # Remove empty or nan rows
        df = df[df[key_col] != "nan"]
        df = df[df[key_col] != ""]
        df = df.fillna("")

        # Drop duplicates, keeping first
        df_unique = df.drop_duplicates(subset=[key_col], keep="first")
        master_dict = df_unique.set_index(key_col).to_dict(orient="index")

        return master_dict, None

    except Exception as e:
        return {}, f"Error loading program master: {str(e)}"

def load_nep_program_master():
    return load_program_master("nep_program_master.xlsx")

def get_program_details(prog_code, master_dict):
    """
    Safely retrieves program details for a given code from the loaded master dictionary.
    """
    if not prog_code or not master_dict:
        return None
    code_str = str(prog_code).strip()
    return master_dict.get(code_str)
