import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import importlib.util
from dbfread import DBF
from openpyxl.styles import numbers
from deep_translator import GoogleTranslator

@st.cache_resource(show_spinner=False)
def load_name_dictionary():
    try:
        from dic import name_translation_dict
        return name_translation_dict, None
    except Exception as exc:
        return {}, str(exc)

def dictionary_available():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(current_dir)
    return (
        os.path.exists(os.path.join(root_dir, "dic.py"))
        or os.path.exists(os.path.join(current_dir, "dic.py"))
    )

STRUCTURE_COLUMNS = [
    "LotNo", "Conv ID", "Faculty", "PRNERN", "ProgType", "ProgNO", "SEAT_NO",
    "COLL_NO", "COLL_NAME", "COLL_NAMEM", "StudLastName", "StudFirstName",
    "StudMidddleName", "StudMotherName", "NAME", "NAME_MARAT", "SEX", "ABBR",
    "CLASS", "MCLASS", "SUB1", "SUB1_NAME", "SUB1_NAMEM", "SUB2", "SUB2_NAME",
    "SUB2_NAMEM", "DEGNM", "MDEGNM", "SUBDEGNM", "MSUBDEGNM", "PER", "MPER", "ABC_ID"
]

SENSITIVE_FIELDS = ["PRNERN", "ProgNO", "SEAT_NO", "ABC_ID"]

# --- Fixed Values (B.Sc.) ---
faculty_options = ["Science & Technology"]
PROGTYPE_FIXED = "Degree"
DEGNM_FIXED = "BACHELOR OF SCIENCE"
MDEGNM_FIXED = "विज्ञान स्नातक"
SUBDEGNM_FIXED = "(Three Year Degree Course)"
MSUBDEGNM_FIXED = "(त्रिवर्षीय पदवी अभ्यासक्रम)"

# --- Helper Functions ---
def to_marathi_digits(x):
    try:
        s = f"{float(x):.2f}" if isinstance(x, (int, float, np.number)) else str(x)
        trans = str.maketrans("0123456789", "०१२३४५६७८९")
        return s.translate(trans)
    except Exception:
        return str(x)

def translate_to_marathi_text(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""
    try:
        return GoogleTranslator(source="auto", target="mr").translate(str(value))
    except Exception:
        return to_marathi_digits(value)

def normalize_colname(name):
    return str(name).strip().lower().replace(" ", "").replace("_", "")

def find_best_match(target_col, source_cols):
    # Aliases sorted by priority (MAJ1/MAJ2 are checked first)
    aliases = {
        "prnern": ["prn", "prnern", "enrolment"],
        "sub1": ["maj1", "course1", "sub1"],
        "sub2": ["maj2", "course2", "sub2", "course4"]
    }
    
    target_norm = normalize_colname(target_col)
    
    # Priority 1: Check defined aliases in exact order of preference
    if target_norm in aliases:
        for alias in aliases[target_norm]:
            for src in source_cols:
                if normalize_colname(src) == alias:
                    return src
                    
    # Priority 2: Exact column name match
    for src in source_cols:
        if normalize_colname(src) == target_norm:
            return src
            
    # Priority 3: Substring match
    for src in source_cols:
        if target_norm in normalize_colname(src) or normalize_colname(src) in target_norm:
            return src
            
    return None

def run_bsc_exam_app():
    # Load subjects_marathi from BSCPRG directory dynamically
    current_dir = os.path.dirname(os.path.abspath(__file__))
    subjects_path = os.path.join(current_dir, 'subjects_marathi.py')
    spec = importlib.util.spec_from_file_location("subjects_marathi_bsc", subjects_path)
    subjects_marathi = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(subjects_marathi)

    # --- UI Status ---
    if dictionary_available():
        st.sidebar.success("✅ Name dictionary (dic.py) loaded.")
    else:
        st.sidebar.warning("⚠️ `dic.py` not found. Name translation disabled.")

    # --- File Upload ---
    st.markdown("## 📁 File Upload (B.Sc.)")
    uploaded_file = st.file_uploader("Choose a DBF or Excel file", type=["dbf", "xlsx", "csv"], key="bsc_file_uploader")

    if uploaded_file:
        file_name = uploaded_file.name.lower()
        file_base = os.path.splitext(uploaded_file.name)[0]
        try:
            with st.spinner("Processing file..."):
                if file_name.endswith(".dbf"):
                    import tempfile
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".dbf") as tmp:
                        tmp.write(uploaded_file.getvalue())
                        tmp_path = tmp.name
                    table = DBF(tmp_path, load=True, char_decode_errors="ignore")
                    df = pd.DataFrame(iter(table))
                    os.unlink(tmp_path)
                elif file_name.endswith(".csv"):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
            st.success(f"✅ Loaded file with {len(df)} rows and {len(df.columns)} cols")
            st.dataframe(df.head(), width="stretch")

            # --- Audit Log ---
            df.columns = [str(c).strip() for c in df.columns]
            audit_log = pd.DataFrame({"RowNo": range(1, len(df)+1), "Reason": "Included"})
            drop_index = set()

            if "RSLT" in df.columns:
                fail_idx = df[df["RSLT"].astype(str).str.upper() == "F"].index
                for i in fail_idx:
                    audit_log.loc[i, "Reason"] = "Excluded: RSLT=F"
                drop_index.update(fail_idx)

            for col in ["RES", "FREM"]:
                if col in df.columns:
                    non_blank = df[df[col].notna() & (df[col].astype(str).str.strip() != "")].index
                    for i in non_blank:
                        if i not in drop_index:
                            audit_log.loc[i, "Reason"] = f"Excluded: {col} has value"
                    drop_index.update(non_blank)

            df_filtered = df.drop(index=list(drop_index)).reset_index(drop=True)
            st.info(f"🧹 Filtered out {len(df) - len(df_filtered)} rows")
            st.dataframe(audit_log.head(30), width="stretch")

            # --- Column Mapping ---
            auto_map = {}
            for col in STRUCTURE_COLUMNS:
                match = find_best_match(col, df_filtered.columns)
                if match:
                    auto_map[col] = match

            st.markdown("---")
            st.markdown("## ⚙️ Settings & Manual Mapping")
            
            # --- Subject Selection Overrides ---
            with st.expander("Subject Column Mapping (Verify & Change)", expanded=True):
                cols = st.columns(2)
                default_sub1_idx = list(df.columns).index(auto_map.get("SUB1")) + 1 if auto_map.get("SUB1") in df.columns else 0
                default_sub2_idx = list(df.columns).index(auto_map.get("SUB2")) + 1 if auto_map.get("SUB2") in df.columns else 0
                
                sub1_col_selected = cols[0].selectbox("Select Column for SUB1 (MAJ1 priority)", ["-- None --"] + list(df.columns), index=default_sub1_idx, key="bsc_sub1_select")
                sub2_col_selected = cols[1].selectbox("Select Column for SUB2 (MAJ2 priority)", ["-- None --"] + list(df.columns), index=default_sub2_idx, key="bsc_sub2_select")

            # --- Manual Entries ---
            st.write("### Fixed Values")
            faculty_value = st.selectbox("Select Faculty", faculty_options, key="bsc_faculty_select")
            abbr_value = st.text_input("Enter ABBR (e.g. BSC, BA, BCOM)", "", key="bsc_abbr_input")
            per_value = st.text_input("Enter PER (e.g. 78.90 or PASS)", "", key="bsc_per_input")

            if st.button("✅ Generate Structured File", type="primary", width="stretch", key="bsc_generate_btn"):
                out = pd.DataFrame("", index=df_filtered.index, columns=STRUCTURE_COLUMNS)

                # Apply mapped columns
                for col, src in auto_map.items():
                    if src in df_filtered.columns:
                        out[col] = df_filtered[src].astype(str)

                # Fixed & manual values
                out["Faculty"] = faculty_value
                out["ProgType"] = PROGTYPE_FIXED
                out["DEGNM"] = DEGNM_FIXED
                out["MDEGNM"] = MDEGNM_FIXED
                out["SUBDEGNM"] = SUBDEGNM_FIXED
                out["MSUBDEGNM"] = MSUBDEGNM_FIXED
                out["ABBR"] = abbr_value
                out["PER"] = per_value
                out["ProgNO"] = file_base
                for coln in ["StudLastName", "StudFirstName", "StudMidddleName", "StudMotherName"]:
                    out[coln] = ""

                # CLASS & Marathi CLASS
                if "CGPA" in df_filtered.columns and "GCGPA" in df_filtered.columns:
                    out["CLASS"] = (
                        pd.to_numeric(df_filtered["CGPA"], errors="coerce").fillna(0) +
                        pd.to_numeric(df_filtered["GCGPA"], errors="coerce").fillna(0)
                    ).round(2)
                else:
                    out["CLASS"] = 0.00
                out["MCLASS"] = out["CLASS"].apply(to_marathi_digits)

                out["MPER"] = translate_to_marathi_text(per_value)
                out["LotNo"] = np.arange(1, len(out)+1)

                # --- Subject Mapping (Updated logic for strict matches & B.Sc.) ---
                def map_subjects(row):
                    # Use manually selected columns
                    sub1 = str(row.get(sub1_col_selected, "")).strip().upper() if sub1_col_selected != "-- None --" else ""
                    sub2 = str(row.get(sub2_col_selected, "")).strip().upper() if sub2_col_selected != "-- None --" else ""
                    
                    mapped = {"SUB1_NAME": "", "SUB1_NAMEM": "", "SUB2_NAME": "", "SUB2_NAMEM": ""}

                    # Sort dictionary keys by length (longest first)
                    sorted_codes = sorted(subjects_marathi.subject_dict.keys(), key=lambda x: len(str(x)), reverse=True)
                    
                    sub1_found, sub2_found = False, False
                    
                    for sub_code in sorted_codes:
                        sub_code_str = str(sub_code).strip().upper()
                        data = subjects_marathi.subject_dict[sub_code]
                        
                        if not sub1_found and sub1 and sub1.startswith(sub_code_str):
                            mapped["SUB1_NAME"] = data["name"]
                            mapped["SUB1_NAMEM"] = data["namem"]
                            sub1_found = True
                            
                        if not sub2_found and sub2 and sub2.startswith(sub_code_str):
                            mapped["SUB2_NAME"] = data["name"]
                            mapped["SUB2_NAMEM"] = data["namem"]
                            sub2_found = True
                            
                        if sub1_found and sub2_found:
                            break

                    # Apply B.Sc. Rule
                    if mapped["SUB1_NAME"] and mapped["SUB2_NAME"]:
                        if mapped["SUB1_NAME"] == mapped["SUB2_NAME"]:
                            # Same subject → SUB1 = Major
                            mapped["SUB1_NAME"] += " (Major)"
                            mapped["SUB1_NAMEM"] += " (प्रमुख)"
                            mapped["SUB2_NAME"] = ""
                            mapped["SUB2_NAMEM"] = ""
                        else:
                            # Different subjects → SUB2 = Major
                            mapped["SUB2_NAME"] += " (Major)"
                            mapped["SUB2_NAMEM"] += " (प्रमुख)"
                    elif mapped["SUB1_NAME"] and not mapped["SUB2_NAME"]:
                        mapped["SUB1_NAME"] += " (Major)"
                        mapped["SUB1_NAMEM"] += " (प्रमुख)"
                    elif mapped["SUB2_NAME"] and not mapped["SUB1_NAME"]:
                        mapped["SUB2_NAME"] += " (Major)"
                        mapped["SUB2_NAMEM"] += " (प्रमुख)"

                    return pd.Series(mapped)

                subj_map_df = df_filtered.apply(map_subjects, axis=1)
                out[["SUB1_NAME", "SUB1_NAMEM", "SUB2_NAME", "SUB2_NAMEM"]] = subj_map_df

                # --- College Lookup ---
                try:
                    from college_master import get_college_by_no
                    out["COLL_NAME"] = out["COLL_NO"].map(
                        lambda x: get_college_by_no(str(x))["COLL_NAME"] if pd.notna(x) else ""
                    )
                    out["COLL_NAMEM"] = out["COLL_NO"].map(
                        lambda x: get_college_by_no(str(x))["COLL_NAMEM"] if pd.notna(x) else ""
                    )
                except ImportError:
                    out["COLL_NAME"] = ""
                    out["COLL_NAMEM"] = ""

                # --- Export to Excel ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    out.to_excel(writer, index=False, sheet_name="Structured")
                    audit_log.to_excel(writer, index=False, sheet_name="Audit_Log")

                    workbook = writer.book
                    sheet = writer.sheets["Structured"]
                    if "CLASS" in out.columns:
                        class_col_idx = out.columns.get_loc("CLASS")
                        class_format = workbook.add_format({"num_format": "0.00"})
                        sheet.set_column(class_col_idx, class_col_idx, None, class_format)

                output.seek(0)
                st.download_button(
                    "📥 Download BSc_Structured_Output_With_Audit.xlsx",
                    data=output.getvalue(),
                    file_name="BSc_Structured_Output_With_Audit.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="bsc_download_btn"
                )

        except Exception as e:
            st.error(f"❌ Error: {e}")
    else:
        st.warning("📁 Please upload a DBF or Excel file to begin.")
