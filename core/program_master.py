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

def _detect_key_column(df):
    """
    Intelligently detects the Program Code / ID column, skipping serial numbers.
    """
    # 1. Exact priority matches
    for col in df.columns:
        norm = str(col).strip().upper().replace("_", "").replace(".", "").replace(" ", "")
        if norm in {"PROGNO", "PROGCD", "PROGRAMNO", "PROGRAMCODE", "PROGCODE", "CODE", "PROGRAM"}:
            return col

    # 2. Substring 'PROG' matches
    for col in df.columns:
        if "PROG" in str(col).strip().upper():
            return col

    # 3. First non-serial number column
    for col in df.columns:
        norm = str(col).strip().upper().replace("_", "").replace(".", "").replace(" ", "")
        if norm not in {"SR", "SRNO", "NO", "INDEX", "ID", "SLNO"}:
            return col

    return df.columns[0]

def _detect_abbr_column(df):
    """Detects the Abbreviation column if present."""
    for col in df.columns:
        norm = str(col).strip().upper().replace("_", "").replace(".", "").replace(" ", "")
        if norm in {"ABBR", "ABBREVIATION", "DEGREEABBR", "PROGABBR"}:
            return col
    return None

def load_program_master(filename="program_master.xlsx"):
    """
    Loads program master Excel file into a fast, case-insensitive dictionary
    keyed by both Program Code (e.g., '1S00146', '3335261') and Abbr (e.g., 'BSCC').
    Returns (dict, error_message).
    """
    file_path = get_program_master_path(filename)
    
    if not os.path.exists(file_path):
        return {}, f"Program master file '{filename}' not found."

    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_to_use = "Sheet3" if "Sheet3" in excel_file.sheet_names else excel_file.sheet_names[0]
        
        df = pd.read_excel(file_path, sheet_name=sheet_to_use, header=0, dtype=str)

        if len(df.columns) < 2:
            return {}, f"Error: '{sheet_to_use}' must have at least two columns."

        # Detect true key column (e.g. PROG_NO)
        key_col = _detect_key_column(df)
        abbr_col = _detect_abbr_column(df)

        df = df.fillna("").astype(str)
        df = df.map(lambda x: "" if str(x).strip().lower() in {"nan", "none", "nat"} else str(x).strip())

        # Remove empty key rows
        df = df[df[key_col] != ""]

        master_dict = {}

        for _, row in df.iterrows():
            raw_key = row[key_col].strip()
            if not raw_key:
                continue

            # Build standardized record with canonical uppercase keys + case-tolerant aliases
            record = {}
            for col in df.columns:
                val = str(row[col]).strip()
                if val.lower() in {"nan", "none"}:
                    val = ""
                col_upper = col.upper().strip()
                record[col_upper] = val
                record[col] = val  # keep original column name casing as well
            
            # Map canonical standard field aliases
            canonical_aliases = {
                "FACULTY": ["FACULTY", "Faculty"],
                "ABBR": ["ABBR", "Abbr", "Abbreviation"],
                "DEGNM": ["DEGNM", "Degnm", "DegreeName", "DEG_NAME"],
                "MDEGNM": ["MDEGNM", "Mdegnm", "MarathiDegreeName", "MDEG_NAME"],
                "SUBDEGNM": ["SUBDEGNM", "Subdegnm", "SUB_DEG_NAME"],
                "MSUBDEGNM": ["MSUBDEGNM", "Msubdegnm", "MSUB_DEG_NAME"],
                "SUB1_NAME": ["SUB1_NAME", "Sub1_Name", "SUB1"],
                "SUB1_NAMEM": ["SUB1_NAMEM", "Sub1_Namem", "SUB1M"],
            }
            for canonical_key, aliases in canonical_aliases.items():
                found_val = ""
                for alias in aliases:
                    if alias in record and record[alias]:
                        found_val = record[alias]
                        break
                    if alias.upper() in record and record[alias.upper()]:
                        found_val = record[alias.upper()]
                        break
                record[canonical_key] = found_val
                # Also assign title/capitalized form
                record[canonical_key.capitalize()] = found_val

            # Index by primary code (exact, uppercase, and stripped of .0)
            key_exact = raw_key
            key_upper = raw_key.upper()
            key_no_dot = raw_key.split(".")[0] if raw_key.endswith(".0") else raw_key

            master_dict[key_exact] = record
            master_dict[key_upper] = record
            if key_no_dot != key_exact:
                master_dict[key_no_dot] = record
                master_dict[key_no_dot.upper()] = record

            # Index by ABBR if available and not yet keyed
            if abbr_col and row.get(abbr_col):
                abbr_val = row[abbr_col].strip()
                if abbr_val and abbr_val not in master_dict:
                    master_dict[abbr_val] = record
                    master_dict[abbr_val.upper()] = record

        return master_dict, None

    except Exception as e:
        return {}, f"Error loading program master: {str(e)}"

def load_nep_program_master():
    return load_program_master("nep_program_master.xlsx")

def get_program_details(prog_code, master_dict):
    """
    Safely retrieves program details for a given code from the loaded master dictionary.
    Handles case variations, whitespace, and numerical representations.
    """
    if not prog_code or not master_dict:
        return None

    code_str = str(prog_code).strip()
    if not code_str:
        return None

    # 1. Direct match
    if code_str in master_dict:
        return master_dict[code_str]

    # 2. Uppercase match
    code_upper = code_str.upper()
    if code_upper in master_dict:
        return master_dict[code_upper]

    # 3. Clean floating point suffix (e.g. '1234.0' -> '1234')
    if code_str.endswith(".0"):
        code_no_dot = code_str[:-2]
        if code_no_dot in master_dict:
            return master_dict[code_no_dot]
        if code_no_dot.upper() in master_dict:
            return master_dict[code_no_dot.upper()]

    # 4. Case-insensitive search across keys
    for k, v in master_dict.items():
        if k.upper() == code_upper:
            return v

    return None
