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

STRUCTURE_COLUMNS = [ "LotNo", "Conv ID", "Faculty", "PRNERN", "ProgType", "APPL_NO", "SEAT_NO", "COLL_NO", "COLL_NAME", "COLL_NAMEM", "StudLastName", "StudFirstName", "StudMidddleName", "StudMotherName", "NAME", "NAME_MARAT", "SEX", "ABBR", "CLASS", "MCLASS", "SUB1", "SUB1_NAME", "SUB1_NAMEM", "SUB2", "SUB2_NAME", "SUB2_NAMEM", "DEGNM", "MDEGNM", "SUBDEGNM", "MSUBDEGNM", "PER", "MPER"]
AUTO_MAP_RULES = { "PRNERN": "PRN", "SEAT_NO": "SEAT_NO", "COLL_NO": "COLL_NO", "NAME": "NAME", "SEX": "SEX" }
SENSITIVE_FIELDS = ["PRNERN", "APPL_NO", "SEAT_NO"]

def to_marathi_digits(x):
    text = str(x)
    map_digits = str.maketrans("0123456789", "०१२३४५६७८९")
    return text.translate(map_digits)

def safe_num_series(s): 
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def run_regular_data_app():
    program_master_dict, program_master_error = get_program_master_cache()
    PROGRAM_MASTER_ERROR = program_master_error

    if dictionary_available():
        st.sidebar.success("✅ Universal dictionary (dic.py) loaded.")
    else:
        st.sidebar.warning("⚠️ Universal `dic.py` not found. Name translation disabled.")

    if program_master_dict:
        st.sidebar.success("✅ Program master (program_master.xlsx) loaded.")
    else:
        st.sidebar.error(f"❌ {PROGRAM_MASTER_ERROR}")

    if 'mappings' not in st.session_state: 
        st.session_state.mappings = {}

    if 'manual_program_details' not in st.session_state: 
        st.session_state.manual_program_details = {}

    st.markdown("## 📁 Step 1: Upload Data File")
    uploaded_file = st.file_uploader("Choose a DBF or Excel file", type=None, key="regular_data_uploader")

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
            key="regular_data_file_type"
        )

        # 1. Immediate friendly validation for file type mismatch
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
                        st.success(f"✅ Loaded as DBF with {len(df)} rows")
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
                        df = pd.read_excel(uploaded_file, dtype=str)
                        df.columns = [str(c).strip().upper() for c in df.columns]
                        st.success(f"✅ Loaded as Excel with {len(df)} rows")
                    except Exception as excel_err:
                        st.error(f"❌ **Failed to read Excel file**: `{uploaded_file.name}`. Please ensure the file is a valid Excel spreadsheet.")
                        with st.expander("Technical details"):
                            st.caption(str(excel_err))
                        return

            # Robust, case-insensitive mapping logic
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
            st.markdown("## 🔑 Step 2: Enter Program Information")

            st.info(
                "Enter the single Program Number that applies to all records. This will be used to look up program details.",
                icon="ℹ️"
            )

            manual_prog_no = st.text_input("**Program Number (PROG_NO)**", key="regular_manual_prog_no").strip()

            program_details_found = False
            selected_program_details = {}

            if manual_prog_no and program_master_dict:
                fetched_details = get_program_details(
                    manual_prog_no,
                    program_master_dict
                )

                if fetched_details:
                    st.success(
                        f"✅ Program '{manual_prog_no}' data loaded successfully!"
                    )

                    program_details_found = True
                    selected_program_details = fetched_details

                else:
                    st.warning(
                        f"⚠️ Program '{manual_prog_no}' not found. Please provide the details below."
                    )

                    with st.expander("Manually Enter Program Details", expanded=True):
                        st.session_state.manual_program_details['Faculty'] = st.text_input("Faculty", key="reg_man_fac")
                        st.session_state.manual_program_details['ABBR'] = st.text_input("ABBR", key="reg_man_abbr")
                        st.session_state.manual_program_details['DEGNM'] = st.text_input("DEGNM", key="reg_man_degnm")
                        st.session_state.manual_program_details['MDEGNM'] = st.text_input("MDEGNM", key="reg_man_mdegnm")
                        st.session_state.manual_program_details['SUBDEGNM'] = st.text_input("SUBDEGNM", key="reg_man_subdegnm")
                        st.session_state.manual_program_details['MSUBDEGNM'] = st.text_input("MSUBDEGNM", key="reg_man_msubdegnm")

            st.markdown("---")
            st.markdown("## 🔗 Step 3: Set Manual and Calculated Columns")

            with st.expander("Expand to set values", expanded=True):

                st.subheader("CLASS Calculation")

                cols = st.columns(2)

                class_col1 = cols[0].selectbox(
                    "Column 1 for CLASS / GRADE",
                    ["-- None --"] + list(df.columns),
                    key="class_col1"
                )

                class_col2 = cols[1].selectbox(
                    "Column 2 for CLASS (Optional)",
                    ["-- None --"] + list(df.columns),
                    key="class_col2"
                )

                st.text(
                    "If numeric, the sum is calculated. If a Grade (A, B+, I, S, P, etc.), it maps automatically."
                )

                st.subheader("Other Manual Columns")

                st.text_input(
                    "Manual Value for `PER` (e.g., March 2025)",
                    key="manual_per"
                )

                st.text(
                    "The value for `PER` will be fully translated to create `MPER`."
                )

                # --- UPDATED: Select source columns instead of manual input ---
                subject_column_mapping = {}

                subject_mapping_fields = [
                    'SUB1',
                    'SUB1_NAME',
                    'SUB1_NAMEM',
                    'SUB2',
                    'SUB2_NAME',
                    'SUB2_NAMEM'
                ]

                for col in subject_mapping_fields:
                    selected_col = st.selectbox(
                        f"Select source column for `{col}`",
                        ["-- None --"] + list(df.columns),
                        key=f"map_{col}"
                    )

                    subject_column_mapping[col] = selected_col

            st.markdown("---")
            st.markdown("## ⚙️ Step 4: Generate Structured File")

            if st.button(
                "✅ Generate Structured File",
                type="primary",
                width="stretch",
                key="regular_generate_btn"
            ):

                if not manual_prog_no:
                    st.error("Error: Please enter a Program Number in Step 2.")

                else:
                    with st.spinner("Generating file..."):

                        # 1. Base Output DataFrame
                        out = pd.DataFrame()

                        for target, source in st.session_state.mappings.items():
                            if source in df.columns:
                                out[target] = df[source]

                        # ---------------------------------------------------------
                        # AUDIT & FILTER LOGIC
                        # ---------------------------------------------------------

                        df_ci = df.copy()
                        df_ci.columns = [str(c).lower().strip() for c in df_ci.columns]

                        audit_df = pd.DataFrame()
                        audit_df['Row_Index'] = df.index + 2

                        seat_source_col = st.session_state.mappings.get("SEAT_NO")

                        if seat_source_col and seat_source_col in df.columns:
                            audit_df['SEAT_NO'] = df[seat_source_col].astype(str).str.strip()
                        else:
                            audit_df['SEAT_NO'] = "N/A (Mapping Missing)"

                        if {'rslt', 'res', 'frem'}.issubset(df_ci.columns):
                            s_rslt = df_ci["rslt"].fillna("").astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": ""})
                            s_res = df_ci["res"].fillna("").astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": ""})
                            s_frem = df_ci["frem"].fillna("").astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": ""})

                            audit_df['RSLT_Val'] = s_rslt
                            audit_df['RES_Val'] = s_res
                            audit_df['FREM_Val'] = s_frem

                            # Result must be Pass
                            cond_rslt = s_rslt.isin(["P", "S", "I", "PASS", "PASS CLASS", "FIRST CLASS", "SECOND CLASS"])

                            # Final remark must be blank
                            cond_frem = s_frem == ""

                            # RES Remark: Accept only Blank and RRA. Exclude RLE, RPV, ABS, or any other remarks.
                            is_blank_res = s_res == ""
                            is_rra_res = (s_res == "RRA") | (s_res.str.contains(r"\bRRA\b", regex=True))
                            is_excluded_remark = s_res.str.contains(r"RLE|RPV|ABS|ADC|FAIL", regex=True)

                            cond_res = (is_blank_res | is_rra_res) & (~is_excluded_remark)

                            mask = cond_rslt & cond_res & cond_frem

                            audit_df['Status'] = "Included"
                            audit_df['Reason'] = "Valid (Blank or RRA Remark with Passing Result)"

                            audit_df.loc[~cond_frem, 'Status'] = "Excluded"
                            audit_df.loc[~cond_frem, 'Reason'] = "FREM is not empty"

                            audit_df.loc[~cond_res, 'Status'] = "Excluded"
                            audit_df.loc[~cond_res, 'Reason'] = "RES is not Blank or RRA (RLE/RPV/Other Remark Excluded)"

                            audit_df.loc[~cond_rslt, 'Status'] = "Excluded"
                            audit_df.loc[~cond_rslt, 'Reason'] = "RSLT is not Pass (P/S/I/PASS)"

                            out = out.loc[mask].reset_index(drop=True)

                            st.info(
                                f"ℹ️ Filtering Applied: {mask.sum()} records kept out of {len(df)}"
                            )

                        else:
                            st.warning(
                                "⚠️ Pass/Fail filter columns (RSLT, RES, FREM) not found. All included."
                            )

                            mask = pd.Series([True] * len(df))

                            audit_df['Status'] = "Included (No Filter)"
                            audit_df['Reason'] = "Filter Columns Missing"

                            out = out.reset_index(drop=True)

                        # ---------------------------------------------------------
                        # DATA ENRICHMENT
                        # ---------------------------------------------------------

                        if program_details_found:
                            program_details = selected_program_details
                        else:
                            program_details = st.session_state.manual_program_details

                        out['APPL_NO'] = manual_prog_no

                        cols_from_program = [
                            'Faculty',
                            'ABBR',
                            'DEGNM',
                            'MDEGNM',
                            'SUBDEGNM',
                            'MSUBDEGNM'
                        ]

                        for col_name in cols_from_program:
                            value = program_details.get(col_name) or program_details.get(col_name.upper())

                            if value is not None:
                                out[col_name] = value

                        out['LotNo'] = out.index + 1
                        out['ProgType'] = 'Degree'

                        # --- UPDATED: Dynamic CLASS / GRADE Calculation ---
                        if (
                            st.session_state.class_col1 != '-- None --'
                            or st.session_state.class_col2 != '-- None --'
                        ):

                            raw_s1 = (
                                df[st.session_state.class_col1]
                                if st.session_state.class_col1 != '-- None --'
                                else pd.Series([''] * len(df))
                            )

                            raw_s2 = (
                                df[st.session_state.class_col2]
                                if st.session_state.class_col2 != '-- None --'
                                else pd.Series([''] * len(df))
                            )

                            if 'mask' in locals():
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
                                'P': ('Pass Class', 'पास श्रेणीत')
                            }

                            class_eng_list = []
                            class_mar_list = []

                            for v1, v2 in zip(raw_s1, raw_s2):

                                val1_str = str(v1).strip().upper()
                                val2_str = str(v2).strip().upper()

                                if val1_str in GRADE_MAP:
                                    class_eng_list.append(GRADE_MAP[val1_str][0])
                                    class_mar_list.append(GRADE_MAP[val1_str][1])

                                elif val2_str in GRADE_MAP:
                                    class_eng_list.append(GRADE_MAP[val2_str][0])
                                    class_mar_list.append(GRADE_MAP[val2_str][1])

                                else:
                                    try:
                                        n1 = float(v1) if v1 != '' and str(v1).lower() != 'nan' else 0.0
                                    except ValueError:
                                        n1 = 0.0

                                    try:
                                        n2 = float(v2) if v2 != '' and str(v2).lower() != 'nan' else 0.0
                                    except ValueError:
                                        n2 = 0.0

                                    total = round(n1 + n2, 2)

                                    class_eng = f"{total:.2f}"

                                    class_eng_list.append(class_eng)
                                    class_mar_list.append(to_marathi_digits(class_eng))

                            out['CLASS'] = class_eng_list
                            out['MCLASS'] = class_mar_list

                        # Translation
                        if st.session_state.manual_per:
                            per_text = st.session_state.manual_per
                            out['PER'] = per_text
                            out['MPER'] = transliterate_text(per_text)

                        # --- UPDATED: Copy selected subject columns ---
                        for target_col, source_col in subject_column_mapping.items():

                            if source_col != "-- None --" and source_col in df.columns:

                                selected_series = df[source_col]

                                if 'mask' in locals():
                                    selected_series = selected_series.loc[mask].reset_index(drop=True)

                                out[target_col] = selected_series.values

                        if "COLL_NO" in out.columns:
                            out["COLL_NAME"] = out["COLL_NO"].map(
                                lambda x: get_college_by_no(x)["COLL_NAME"]
                            )

                            out["COLL_NAMEM"] = out["COLL_NO"].map(
                                lambda x: get_college_by_no(x)["COLL_NAMEM"]
                            )

                        if "NAME" in out.columns:
                            out['NAME_MARAT'], missing_names = translate_name_series(out['NAME'])
                            if missing_names:
                                st.info(f"{missing_names} unique name words used Google Input Tools fallback.")

                        # Final Cleanup
                        out = out.fillna('')

                        out = out.reindex(
                            columns=STRUCTURE_COLUMNS,
                            fill_value=''
                        )

                        clean_identifier_columns(out, SENSITIVE_FIELDS)

                        # ---------------------------------------------------------
                        # EXCEL WRITING
                        # ---------------------------------------------------------

                        output = io.BytesIO()

                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

                            # Sheet 1: Structured Data
                            out.to_excel(
                                writer,
                                index=False,
                                sheet_name="Structured"
                            )

                            workbook = writer.book
                            ws = writer.sheets["Structured"]
                            text_format = workbook.add_format({"num_format": "@"})
                            for col_idx, col_name in enumerate(out.columns):
                                if col_name in SENSITIVE_FIELDS:
                                    ws.set_column(col_idx, col_idx, None, text_format)

                            # Sheet 2: Detailed Audit Log
                            audit_df.to_excel(
                                writer,
                                index=False,
                                sheet_name="Audit_Log"
                            )

                            ws_audit = writer.sheets["Audit_Log"]
                            ws_audit.set_column("B:B", 15)
                            ws_audit.set_column("C:C", 15)
                            ws_audit.set_column("D:D", 30)

                        st.success("🎉 Structured Excel generated!")

                        st.dataframe(
                            out.head(50),
                            width="stretch"
                        )

                        st.download_button(
                            "📥 Download Structured_Output.xlsx",
                            data=output.getvalue(),
                            file_name="Structured_Output.xlsx",
                            mime="application/vnd.ms-excel",
                            width="stretch",
                            key="regular_download_btn"
                        )

        except Exception as e:
            st.error(f"❌ An error occurred while processing the data: {e}")
            with st.expander("Technical details"):
                import traceback
                st.code(traceback.format_exc())

    else:
        st.info("👋 Welcome! Please upload a file to begin.")
