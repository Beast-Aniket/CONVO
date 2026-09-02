import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import sys
import zipfile
import tempfile
import re
from datetime import datetime
from dbfread import DBF
from openpyxl.styles import numbers

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from core.name_lookup import translate_name_series, transliterate_text
from core.college_master import get_college_by_no
from core.program_master import load_nep_program_master as load_program_master, get_program_details

STRUCTURE_COLUMNS = [ "LotNo", "Conv ID", "Faculty", "PRNERN", "ProgType", "APPL_NO", "SEAT_NO", "COLL_NO", "COLL_NAME", "COLL_NAMEM", "StudLastName", "StudFirstName", "StudMidddleName", "StudMotherName", "NAME", "NAME_MARAT", "SEX", "ABBR", "CLASS", "MCLASS", "SUB1", "SUB1_NAME", "SUB1_NAMEM", "SUB2", "SUB2_NAME", "SUB2_NAMEM", "DEGNM", "MDEGNM", "SUBDEGNM", "MSUBDEGNM", "PER", "MPER"]

AUTO_MAP_RULES = { 
    "PRNERN": "Enrolment Number", 
    "SEAT_NO": "Roll Number", 
    "NAME": "Student Name", 
    "CLASS": "CGPA", 
    "COLL_NO": "Organizational Unit"
}

SENSITIVE_FIELDS = ["PRNERN", "APPL_NO", "SEAT_NO"]

# -------------------------
# Helpers
# -------------------------
def to_marathi_digits(x):
    text = str(x)
    map_digits = str.maketrans("0123456789", "०१२३४५६७८९")
    return text.translate(map_digits)

def translate_full_name_by_word(name, d): 
    if not isinstance(name, str): return ""
    words = name.strip().split()
    translated_words = [d.get(w.upper(), w) for w in words] 
    return " ".join(translated_words)

def clean_prn(val):
    if pd.isna(val) or val == "": return ""
    try:
        if isinstance(val, float): return str(int(val))
        val_str = str(val).strip()
        if 'E' in val_str.upper() or 'e' in val_str:
            return str(int(float(val_str)))
        return val_str
    except: return str(val)

def extract_college_code(val):
    if not isinstance(val, str): return val
    match = re.search(r"MU-(\d+)\s*:", val)
    if match: return str(int(match.group(1))) 
    return val

def map_gender_value(val):
    s = str(val).strip().upper()
    if s.startswith('M'): return "1"
    if s.startswith('F'): return "2"
    return ""

def load_data(uploaded_file):
    file_name = uploaded_file.name.lower()
    if file_name.endswith('.dbf'):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".dbf") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        try:
            df = pd.DataFrame(iter(DBF(tmp_path, load=True, char_decode_errors="ignore")))
            return df
        finally:
            os.unlink(tmp_path)
    else:
        return pd.read_excel(uploaded_file)

# -------------------------
# Main App
# -------------------------
def run_nep_exam_app():
    st.title("🎓 Bulk Excel Processor")

    if PROGRAM_MASTER_ERROR:
        st.error(PROGRAM_MASTER_ERROR)

    uploaded_files = st.file_uploader("Choose Excel/DBF files", type=['xlsx', 'xls', 'dbf'], accept_multiple_files=True)

    if uploaded_files:
        st.markdown("---")
        with st.expander("Manual Overrides (Optional)", expanded=False):
            manual_per = st.text_input("Value for `PER` (e.g., March 2025)", key="manual_per")
            st.info("Subject names and Degree info will now be fetched automatically from Program Master.")

        if st.button("🚀 Process All Files", type="primary", use_container_width=True):
            if not PROGRAM_MASTER_DICT:
                st.error("Cannot proceed: Program Master not loaded.")
            else:
                zip_buffer = io.BytesIO()
                skipped_files = []
                processed_count = 0
                
                progress_bar = st.progress(0)
                status_text = st.empty()

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
                    for i, file in enumerate(uploaded_files):
                        file_name = file.name
                        prog_no = os.path.splitext(file_name)[0].strip()
                        
                        status_text.text(f"Processing: {file_name}...")
                        progress_bar.progress((i + 1) / len(uploaded_files))

                        if prog_no not in PROGRAM_MASTER_DICT:
                            skipped_files.append(f"{file_name} (ID '{prog_no}' not in Master)")
                            continue

                        program_details = PROGRAM_MASTER_DICT[prog_no]

                        try:
                            df = load_data(file)
                            col_map = {str(c).lower().strip().replace("_", " "): c for c in df.columns}
                            def get_col(name):
                                return col_map.get(name.lower().strip().replace("_", " "))

                            # --- 1. MAPPING ---
                            out = pd.DataFrame()
                            
                            prn_col = get_col(AUTO_MAP_RULES["PRNERN"])
                            if prn_col: out["PRNERN"] = df[prn_col].apply(clean_prn)
                                
                            seat_col = get_col(AUTO_MAP_RULES["SEAT_NO"])
                            if seat_col: out["SEAT_NO"] = df[seat_col]
                            
                            name_col = get_col(AUTO_MAP_RULES["NAME"])
                            if name_col: 
                                # LOGIC: Ensure Student Name is ALL CAPITAL
                                out["NAME"] = df[name_col].astype(str).str.upper().str.strip()
                            
                            cgpa_col = get_col(AUTO_MAP_RULES["CLASS"])
                            if cgpa_col: out["CLASS"] = df[cgpa_col]
                            
                            org_col = get_col(AUTO_MAP_RULES["COLL_NO"])
                            if org_col: out["COLL_NO"] = df[org_col].apply(extract_college_code)
                            
                            gender_col = get_col("Gender") or get_col("Sex") or get_col("Student Gender")
                            if gender_col: out["SEX"] = df[gender_col].apply(map_gender_value)

                            # --- 2. AUDIT & FILTER ---
                            res_status_col = get_col("Result Status") 
                            remarks_col = get_col("Remarks") 
                            
                            if res_status_col:
                                s_status = df[res_status_col].astype(str).str.strip().str.upper()
                                s_remarks = df[remarks_col].astype(str).str.strip().str.upper().replace("NAN", "") if remarks_col else pd.Series([""] * len(df))
                                
                                mask = (s_status == "PASS") & (~(s_remarks.str.contains("RLE") | s_remarks.str.contains("RPV")))
                                out = out.loc[mask].reset_index(drop=True)

                            # --- 3. ENRICHMENT (AUTO-FETCH FROM MASTER) ---
                            out['APPL_NO'] = prog_no
                            out['LotNo'] = out.index + 1
                            out['ProgType'] = 'Degree'
                            
                            # List of columns to pull automatically from the Excel Master
                            master_cols = [
                                'Faculty', 'ABBR', 'DEGNM', 'MDEGNM', 'SUBDEGNM', 'MSUBDEGNM',
                                'SUB1', 'SUB1_NAME', 'SUB1_NAMEM', 
                                'SUB2', 'SUB2_NAME', 'SUB2_NAMEM'
                            ]
                            for c in master_cols:
                                # Fetch from master (handling case sensitivity)
                                val = program_details.get(c) or program_details.get(c.upper())
                                if val: out[c] = val

                            # College Lookup
                            if "COLL_NO" in out.columns:
                                out["COLL_NAME"] = out["COLL_NO"].map(lambda x: get_college_by_no(x)["COLL_NAME"])
                                out["COLL_NAMEM"] = out["COLL_NO"].map(lambda x: get_college_by_no(x)["COLL_NAMEM"])

                            # Translations
                            if manual_per:
                                out['PER'] = manual_per
                                out['MPER'] = transliterate_text(manual_per)

                            if "NAME" in out.columns:
                                out['NAME_MARAT'], _ = translate_name_series(out['NAME'])

                            if "CLASS" in out.columns:
                                out['MCLASS'] = out['CLASS'].astype(str).map(to_marathi_digits)

                            # Finalize Structure
                            out = out.fillna('').reindex(columns=STRUCTURE_COLUMNS, fill_value='')

                            # --- 4. EXPORT ---
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                                out.to_excel(writer, index=False, sheet_name="Structured")
                            
                            archive.writestr(f"{prog_no}_Structured.xlsx", excel_buffer.getvalue())
                            processed_count += 1
                            
                        except Exception as e:
                            skipped_files.append(f"{file_name} (Error: {str(e)})")

                status_text.empty()
                st.success(f"✅ Processed: {processed_count}, Skipped: {len(skipped_files)}")
                
                if skipped_files:
                    with st.expander("⚠️ View Skipped Files"):
                        for f in skipped_files: st.write(f"- {f}")

                st.download_button(
                    label="📦 Download Processed Files (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Batch_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )

if __name__ == "__main__":
    st.set_page_config(page_title="Bulk Data Processor", layout="wide", page_icon="🎓")
    run_nep_app()