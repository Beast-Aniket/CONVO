import streamlit as st
import pandas as pd
import numpy as np
import io
import tempfile
import os
import sys
from datetime import datetime
from dbfread import DBF
from openpyxl.styles import numbers

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from core.data_utils import clean_identifier_columns
from core.name_lookup import translate_name_series, transliterate_text, universal_dictionary_available
from core.college_master import get_college_by_no
from core.program_master import load_program_master, get_program_details

@st.cache_data(show_spinner=False)
def get_program_master_cache():
    return load_program_master()

def dictionary_available():
    return universal_dictionary_available()

STRUCTURE_COLUMNS = [
    "LotNo", "Conv ID", "Faculty", "PRNERN", "ProgType", "APPL_NO", "SEAT_NO",
    "COLL_NO", "COLL_NAME", "COLL_NAMEM", "StudLastName", "StudFirstName",
    "StudMidddleName", "StudMotherName", "NAME", "NAME_MARAT", "SEX", "ABBR",
    "CLASS", "MCLASS", "SUB1", "SUB1_NAME", "SUB1_NAMEM", "SUB2", "SUB2_NAME",
    "SUB2_NAMEM", "DEGNM", "MDEGNM", "SUBDEGNM", "MSUBDEGNM", "PER", "MPER"
]

AUTO_MAP_RULES = {
    "PRNERN": "PRN",
    "SEAT_NO": "SEAT_NO",
    "COLL_NO": "COLL_NO",
    "NAME": "NAME",
    "SEX": "SEX"
}

SENSITIVE_FIELDS = ["PRNERN", "APPL_NO", "SEAT_NO"]

def to_marathi_digits(x):
    text = str(x)
    map_digits = str.maketrans("0123456789", "०१२३४५६७८९")
    return text.translate(map_digits)

def safe_num_series(s):
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def run_rra_exam_app():
    program_master_dict, program_master_error = get_program_master_cache()
    PROGRAM_MASTER_ERROR = program_master_error

    if dictionary_available():
        st.sidebar.success("✅ Universal dictionary (dic.py) loaded.")
    else:
        st.sidebar.warning("⚠️ Universal `dic.py` not found. Translation will fallback to API.")

    if not PROGRAM_MASTER_ERROR:
        st.sidebar.success("✅ Program master (program_master.xlsx) loaded.")
    else:
        st.sidebar.error(f"❌ {PROGRAM_MASTER_ERROR}")

    if 'mappings' not in st.session_state:
        st.session_state.mappings = {}

    if 'manual_program_details' not in st.session_state:
        st.session_state.manual_program_details = {}

    st.markdown("## 📋 RRA Students Data Processor")
    st.info("💡 **RRA Filter Active**: This module exclusively processes and retains students with remark **`RES = RRA`** and **`RSLT = P / PASS`**.", icon="ℹ️")

    st.markdown("### 📁 Step 1: Upload Data File")
    uploaded_file = st.file_uploader("Choose a DBF, Excel, or CSV file", type=None, key="rra_data_uploader")

    if uploaded_file:
        file_name_lower = uploaded_file.name.lower()
        is_dbf_ext = file_name_lower.endswith(".dbf")
        is_excel_ext = file_name_lower.endswith((".xlsx", ".xls", ".xlsm", ".xlsb", ".csv"))
        
        default_index = 0 if is_dbf_ext else 1

        file_type_choice = st.radio(
            "How should this file be processed?",
            ('As a DBF file', 'As an Excel file'),
            index=default_index,
            horizontal=True,
            key="rra_data_file_type"
        )

        # Immediate validation for file format mismatch
        if file_type_choice == 'As an Excel file' and is_dbf_ext:
            st.error(f"⚠️ **File Format Mismatch**: You uploaded a DBF file (`{uploaded_file.name}`) but selected **'As an Excel file'**. Please select **'As a DBF file'** above to proceed.")
            return

        if file_type_choice == 'As a DBF file' and is_excel_ext:
            st.error(f"⚠️ **File Format Mismatch**: You uploaded an Excel file (`{uploaded_file.name}`) but selected **'As a DBF file'**. Please select **'As an Excel file'** above to proceed.")
            return

        try:
            with st.spinner("Processing file..."):
                if file_type_choice == 'As a DBF file':
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".dbf") as tmp:
                        tmp.write(uploaded_file.getvalue())
                        tmp_path = tmp.name

                    try:
                        df = pd.DataFrame(
                            iter(DBF(tmp_path, load=True, char_decode_errors="ignore"))
                        )
                        df.columns = [str(c).strip().upper() for c in df.columns]
                        st.success(f"✅ Loaded as DBF with {len(df)} total rows")
                    except Exception as dbf_err:
                        st.error(f"❌ **Failed to read DBF file**: `{uploaded_file.name}`. Please ensure the file is a valid database file.")
                        with st.expander("Technical details"):
                            st.caption(str(dbf_err))
                        return
                    finally:
                        if os.path.exists(tmp_path):
                            os.unlink(tmp_path)

                else:
                    try:
                        if file_name_lower.endswith(".csv"):
                            df = pd.read_csv(uploaded_file, dtype=str)
                        else:
                            df = pd.read_excel(uploaded_file, dtype=str)
                        df.columns = [str(c).strip().upper() for c in df.columns]
                        st.success(f"✅ Loaded as Excel with {len(df)} total rows")
                    except Exception as excel_err:
                        st.error(f"❌ **Failed to read Excel file**: `{uploaded_file.name}`. Please ensure the file is a valid spreadsheet.")
                        with st.expander("Technical details"):
                            st.caption(str(excel_err))
                        return

            # Column auto mapping
            mappings = {}
            for target, source in AUTO_MAP_RULES.items():
                if source in df.columns:
                    mappings[target] = source
                elif target in df.columns:
                    mappings[target] = target
                else:
                    matched = None
                    for c in df.columns:
                        if c.strip().upper() == source.upper() or c.strip().upper() == target.upper():
                            matched = c
                            break
                    if matched:
                        mappings[target] = matched
            st.session_state.mappings = mappings

            st.dataframe(df.head(), width="stretch")

            st.markdown("---")
            st.markdown("### 🔑 Step 2: Enter Program Information")
            st.info("Enter the Program Number (e.g., `1S00146`, `2C00146`, `3A00146`) to auto-fetch details from Program Master.")

            manual_prog_no = st.text_input("**Program Number (PROG_NO)**", key="rra_manual_prog_no").strip()

            program_details_found = False
            selected_program_details = {}

            if manual_prog_no and program_master_dict:
                fetched_details = get_program_details(manual_prog_no, program_master_dict)

                if fetched_details:
                    st.success(f"✅ Program '{manual_prog_no}' details loaded: **{fetched_details.get('DEGNM')}** ({fetched_details.get('FACULTY')})")
                    program_details_found = True
                    selected_program_details = fetched_details
                else:
                    st.warning(f"⚠️ Program '{manual_prog_no}' not found in master. Please provide the details below.")
                    with st.expander("Manually Enter Program Details", expanded=True):
                        st.session_state.manual_program_details['Faculty'] = st.text_input("Faculty", key="rra_man_fac")
                        st.session_state.manual_program_details['ABBR'] = st.text_input("ABBR", key="rra_man_abbr")
                        st.session_state.manual_program_details['DEGNM'] = st.text_input("DEGNM", key="rra_man_degnm")
                        st.session_state.manual_program_details['MDEGNM'] = st.text_input("MDEGNM", key="rra_man_mdegnm")
                        st.session_state.manual_program_details['SUBDEGNM'] = st.text_input("SUBDEGNM", key="rra_man_subdegnm")
                        st.session_state.manual_program_details['MSUBDEGNM'] = st.text_input("MSUBDEGNM", key="rra_man_msubdegnm")

            st.markdown("---")
            st.markdown("### 🔗 Step 3: Set Manual and Calculated Columns")

            with st.expander("Expand to set values", expanded=True):
                st.subheader("CLASS Calculation")
                cols = st.columns(2)
                class_col1 = cols[0].selectbox("Column 1 for CLASS / GRADE", ["-- None --"] + list(df.columns), key="rra_class_col1")
                class_col2 = cols[1].selectbox("Column 2 for CLASS (Optional)", ["-- None --"] + list(df.columns), key="rra_class_col2")
                st.text("If numeric, the sum is calculated. If a Grade (A, B+, I, S, P, etc.), it maps automatically.")

                st.subheader("Other Manual Columns")
                st.text_input("Manual Value for `PER` (e.g., March 2025)", key="rra_manual_per")
                st.text("The value for `PER` will be fully translated to Marathi numerals and text for `MPER`.")

                # Subject selection
                subject_column_mapping = {}
                subject_mapping_fields = ['SUB1', 'SUB1_NAME', 'SUB1_NAMEM', 'SUB2', 'SUB2_NAME', 'SUB2_NAMEM']

                for col in subject_mapping_fields:
                    selected_col = st.selectbox(
                        f"Select source column for `{col}`",
                        ["-- None --"] + list(df.columns),
                        key=f"rra_map_{col}"
                    )
                    subject_column_mapping[col] = selected_col

            st.markdown("---")
            st.markdown("### ⚙️ Step 4: Generate Structured File (RRA Filtered)")

            if st.button("✅ Generate RRA Structured File", type="primary", width="stretch", key="rra_generate_btn"):
                if not manual_prog_no:
                    st.error("Error: Please enter a Program Number in Step 2.")
                else:
                    with st.spinner("Filtering RRA students and generating file..."):
                        # 1. Base Output DataFrame
                        out = pd.DataFrame()
                        for target, source in st.session_state.mappings.items():
                            if source in df.columns:
                                out[target] = df[source]

                        # ---------------------------------------------------------
                        # RRA STRICT FILTER & AUDIT LOGIC (RES == 'RRA' & RSLT == 'P')
                        # ---------------------------------------------------------
                        df_ci = df.copy()
                        df_ci.columns = [str(c).lower().strip() for c in df_ci.columns]

                        audit_df = pd.DataFrame()
                        audit_df['Row_Index'] = df.index + 2

                        seat_source_col = st.session_state.mappings.get("SEAT_NO")
                        if seat_source_col and seat_source_col in df.columns:
                            audit_df['SEAT_NO'] = df[seat_source_col].astype(str).str.strip()
                        else:
                            audit_df['SEAT_NO'] = "N/A"

                        # Retrieve RES and RSLT series
                        has_res = 'res' in df_ci.columns
                        has_rslt = 'rslt' in df_ci.columns

                        if has_res and has_rslt:
                            s_res = df_ci["res"].astype(str).str.strip().str.upper().replace("NAN", "")
                            s_rslt = df_ci["rslt"].astype(str).str.strip().str.upper().replace("NAN", "")

                            audit_df['RES_Val'] = s_res
                            audit_df['RSLT_Val'] = s_rslt
                            if 'frem' in df_ci.columns:
                                audit_df['FREM_Val'] = df_ci["frem"].astype(str).str.strip().str.upper().replace("NAN", "")

                            # Strict RRA condition: RES must be 'RRA' and RSLT must be 'P' or 'PASS'
                            cond_res_rra = s_res == "RRA"
                            cond_rslt_pass = s_rslt.isin(["P", "PASS"])

                            mask = cond_res_rra & cond_rslt_pass

                            audit_df['Status'] = "Included"
                            audit_df['Reason'] = "Valid RRA Student (RES=RRA & RSLT=P)"

                            audit_df.loc[~cond_res_rra, 'Status'] = "Excluded"
                            audit_df.loc[~cond_res_rra, 'Reason'] = "RES is not RRA"

                            audit_df.loc[~cond_rslt_pass, 'Status'] = "Excluded"
                            audit_df.loc[~cond_rslt_pass, 'Reason'] = "RSLT is not P/PASS"

                            out = out.loc[mask].reset_index(drop=True)

                            st.success(f"🎯 **RRA Filter Applied**: Retained **{mask.sum()}** RRA student records out of **{len(df)}** total rows.")
                        else:
                            missing_cols = []
                            if not has_res: missing_cols.append("RES")
                            if not has_rslt: missing_cols.append("RSLT")
                            st.error(f"❌ Required column(s) {', '.join(missing_cols)} not found in uploaded file. Cannot filter for RRA students.")
                            return

                        # ---------------------------------------------------------
                        # DATA ENRICHMENT
                        # ---------------------------------------------------------
                        if program_details_found:
                            program_details = selected_program_details
                        else:
                            program_details = st.session_state.manual_program_details

                        out['APPL_NO'] = manual_prog_no

                        cols_from_program = ['Faculty', 'ABBR', 'DEGNM', 'MDEGNM', 'SUBDEGNM', 'MSUBDEGNM']
                        for col_name in cols_from_program:
                            value = program_details.get(col_name) or program_details.get(col_name.upper())
                            if value is not None:
                                out[col_name] = value

                        out['LotNo'] = out.index + 1
                        out['ProgType'] = 'Degree'

                        # Dynamic CLASS / GRADE Calculation
                        if st.session_state.rra_class_col1 != '-- None --' or st.session_state.rra_class_col2 != '-- None --':
                            raw_s1 = df[st.session_state.rra_class_col1] if st.session_state.rra_class_col1 != '-- None --' else pd.Series([''] * len(df))
                            raw_s2 = df[st.session_state.rra_class_col2] if st.session_state.rra_class_col2 != '-- None --' else pd.Series([''] * len(df))

                            raw_s1 = raw_s1.loc[mask].reset_index(drop=True)
                            raw_s2 = raw_s2.loc[mask].reset_index(drop=True)

                            GRADE_MAP = {
                                'A': ('‘A’ Grade', '‘ए’ श्रेणीत'),
                                'B': ('‘B’ Grade', '‘बी’ श्रेणीत'),
                                'C': ('‘C’ Grade', '‘सी’ श्रेणीत'),
                                'D': ('‘D’ Grade', '‘डी’ श्रेणीत'),
                                'E': ('‘E’ Grade', '‘इ’ श्रेणीत'),
                                'O': ('‘O’ Grade', '‘ओ’ श्रेणीत'),
                                'A+': ('‘A+’ Grade', '‘ए+’ श्रेणीत'),
                                'B+': ('‘B+’ Grade', '‘बी+’ श्रेणीत'),
                                'I': ('First Class', 'प्रथम श्रेणीत'),
                                'S': ('Second Class', 'द्वितीय श्रेणीत'),
                                'P': ('Pass Class', 'उत्तीर्ण श्रेणीत'),
                                'PASS': ('Pass Class', 'उत्तीर्ण श्रेणीत'),
                                'F': ('Fail', 'अनुत्तीर्ण'),
                                'FAIL': ('Fail', 'अनुत्तीर्ण')
                            }

                            class_eng_list = []
                            class_mar_list = []

                            for v1, v2 in zip(raw_s1, raw_s2):
                                v1_clean = str(v1).strip().upper()
                                v2_clean = str(v2).strip().upper()

                                is_v1_num = pd.to_numeric(v1, errors='coerce') is not np.nan and str(v1).strip() != ''
                                is_v2_num = pd.to_numeric(v2, errors='coerce') is not np.nan and str(v2).strip() != ''

                                if is_v1_num or is_v2_num:
                                    n1 = pd.to_numeric(v1, errors='coerce')
                                    n2 = pd.to_numeric(v2, errors='coerce')
                                    total = (n1 if pd.notna(n1) else 0.0) + (n2 if pd.notna(n2) else 0.0)
                                    class_eng_list.append(f"{total:.2f}")
                                    class_mar_list.append(to_marathi_digits(f"{total:.2f}"))
                                elif v1_clean in GRADE_MAP:
                                    class_eng_list.append(GRADE_MAP[v1_clean][0])
                                    class_mar_list.append(GRADE_MAP[v1_clean][1])
                                elif v2_clean in GRADE_MAP:
                                    class_eng_list.append(GRADE_MAP[v2_clean][0])
                                    class_mar_list.append(GRADE_MAP[v2_clean][1])
                                else:
                                    class_eng = str(v1).strip() if str(v1).strip() else str(v2).strip()
                                    class_eng_list.append(class_eng)
                                    class_mar_list.append(to_marathi_digits(class_eng))

                            out['CLASS'] = class_eng_list
                            out['MCLASS'] = class_mar_list

                        # Translation of PER -> MPER
                        if st.session_state.rra_manual_per:
                            per_text = st.session_state.rra_manual_per
                            out['PER'] = per_text
                            out['MPER'] = transliterate_text(per_text)

                        # Copy selected subject columns
                        for target_col, source_col in subject_column_mapping.items():
                            if source_col != "-- None --" and source_col in df.columns:
                                selected_series = df[source_col].loc[mask].reset_index(drop=True)
                                out[target_col] = selected_series

                        # Name translation via Universal Dictionary + Google Input Tools
                        if 'NAME' in out.columns:
                            existing_marathi = out['NAME_MARAT'] if 'NAME_MARAT' in out.columns else None
                            translated_names, missing_count = translate_name_series(out['NAME'], existing_marathi=existing_marathi)
                            out['NAME_MARAT'] = translated_names

                        # College Master lookup
                        if 'COLL_NO' in out.columns:
                            college_details = out['COLL_NO'].apply(get_college_by_no)
                            out['COLL_NAME'] = college_details.apply(lambda x: x.get('COLL_NAME', ''))
                            out['COLL_NAMEM'] = college_details.apply(lambda x: x.get('COLL_NAMEM', ''))

                        # Ensure all structure columns exist
                        out = out.fillna('')
                        out = out.reindex(columns=STRUCTURE_COLUMNS, fill_value='')

                        # Clean sensitive identifier columns
                        clean_identifier_columns(out, SENSITIVE_FIELDS)

                        # ---------------------------------------------------------
                        # EXCEL WRITING
                        # ---------------------------------------------------------
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            out.to_excel(writer, index=False, sheet_name="RRA_Structured")

                            workbook = writer.book
                            ws = writer.sheets["RRA_Structured"]
                            text_format = workbook.add_format({"num_format": "@"})
                            for col_idx, col_name in enumerate(out.columns):
                                if col_name in SENSITIVE_FIELDS:
                                    ws.set_column(col_idx, col_idx, None, text_format)

                            # Sheet 2: Detailed Audit Log
                            audit_df.to_excel(writer, index=False, sheet_name="Audit_Log")
                            ws_audit = writer.sheets["Audit_Log"]
                            ws_audit.set_column("B:B", 15)
                            ws_audit.set_column("C:C", 15)
                            ws_audit.set_column("D:D", 30)

                        st.success("🎉 RRA Structured Excel generated successfully!")
                        st.dataframe(out.head(50), width="stretch")

                        st.download_button(
                            "📥 Download RRA_Structured_Output.xlsx",
                            data=output.getvalue(),
                            file_name="RRA_Structured_Output.xlsx",
                            mime="application/vnd.ms-excel",
                            width="stretch",
                            key="rra_download_btn"
                        )

        except Exception as e:
            st.error(f"❌ An error occurred while processing the data: {e}")
            with st.expander("Technical details"):
                import traceback
                st.code(traceback.format_exc())

    else:
        st.info("👋 Welcome! Please upload a DBF, Excel, or CSV file to begin.")
