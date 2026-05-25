import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import sys
from dbfread import DBF
from openpyxl.styles import numbers
from deep_translator import GoogleTranslator
import subjects_marathi

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from data_utils import clean_identifier_columns

STRUCTURE_COLUMNS = [
    "LotNo", "Conv ID", "Faculty", "PRNERN", "ProgType", "ProgNO", "SEAT_NO",
    "COLL_NO", "COLL_NAME", "COLL_NAMEM", "StudLastName", "StudFirstName",
    "StudMidddleName", "StudMotherName", "NAME", "NAME_MARAT", "SEX", "ABBR",
    "CLASS", "MCLASS", "SUB1", "SUB1_NAME", "SUB1_NAMEM", "SUB2", "SUB2_NAME",
    "SUB2_NAMEM", "DEGNM", "MDEGNM", "SUBDEGNM", "MSUBDEGNM", "PER", "MPER", "ABC_ID"
]

SENSITIVE_FIELDS = ["PRNERN", "ProgNO", "SEAT_NO", "ABC_ID"]

st.set_page_config(page_title="DBF/Excel Auto Mapper", layout="wide", page_icon="📊")

# --- Fixed Values ---
faculty_options = ["Interdisciplinary Studies", "Humanities"]
PROGTYPE_FIXED = "Degree"
DEGNM_FIXED = "BACHELOR OF ARTS"
MDEGNM_FIXED = "मानव्यविद्या स्नातक"
SUBDEGNM_FIXED = "(Three Year Degree Course)"
MSUBDEGNM_FIXED = "(त्रिवर्षीय पदवी अभ्यासक्रम)"

# --- Helper Functions ---
def to_marathi_digits(x):
    try:
        s = str(x)
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
    target_norm = normalize_colname(target_col)
    for src in source_cols:
        if normalize_colname(src) == target_norm:
            return src
    for src in source_cols:
        if target_norm in normalize_colname(src) or normalize_colname(src) in target_norm:
            return src
    return None

# --- File Upload ---
st.markdown("## 📁 File Upload")
uploaded_file = st.file_uploader("Choose a DBF or Excel file", type=["dbf", "xlsx"])

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
            else:
                df = pd.read_excel(uploaded_file, dtype=str)
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

        st.write("Column mapping:", auto_map)

        # --- Manual Entries ---
        faculty_value = st.selectbox("Select Faculty", faculty_options)
        abbr_value = st.text_input("Enter ABBR (e.g. BSC, BA, BCOM)", "")
        per_value = st.text_input("Enter PER (e.g. 78.90 or PASS)", "")

        if st.button("✅ Generate Structured File"):
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
            for coln in ["StudLastName","StudFirstName","StudMidddleName","StudMotherName"]:
                out[coln] = ""

            # CLASS & Marathi CLASS
            if "CGPA" in df_filtered.columns and "GCGPA" in df_filtered.columns:
                out["CLASS"] = pd.to_numeric(df_filtered["CGPA"], errors="coerce").fillna(0) + pd.to_numeric(df_filtered["GCGPA"], errors="coerce").fillna(0)
            else:
                out["CLASS"] = 0
            out["MCLASS"] = out["CLASS"].apply(to_marathi_digits)

            out["MPER"] = translate_to_marathi_text(per_value)
            out["LotNo"] = np.arange(1, len(out)+1)

            # --- Subject Mapping ---
            def map_subjects(row):
                sub1, sub2 = str(row.get(auto_map.get("SUB1"), "")), str(row.get(auto_map.get("SUB2"), ""))
                mapped = {"SUB1_NAME": "", "SUB1_NAMEM": "", "SUB2_NAME": "", "SUB2_NAMEM": ""}
                for sub_code, data in subjects_marathi.subject_dict.items():
                    if sub1.startswith(str(sub_code)):
                        mapped["SUB1_NAME"] = data["name"]
                        mapped["SUB1_NAMEM"] = data["namem"]
                    if sub2.startswith(str(sub_code)):
                        mapped["SUB2_NAME"] = data["name"]
                        mapped["SUB2_NAMEM"] = data["namem"]
                if mapped["SUB1_NAME"] and mapped["SUB1_NAME"] == mapped["SUB2_NAME"]:
                    mapped["SUB1_NAME"] += " (Major)"
                    mapped["SUB1_NAMEM"] += " (Major)"
                    mapped["SUB2_NAME"] = ""
                    mapped["SUB2_NAMEM"] = ""
                return pd.Series(mapped)

            subj_map_df = df_filtered.apply(map_subjects, axis=1)
            out[["SUB1_NAME","SUB1_NAMEM","SUB2_NAME","SUB2_NAMEM"]] = subj_map_df

            # --- College Lookup (if get_college_by_no is available) ---
            try:
                from college_master import get_college_by_no
                out["COLL_NAME"] = out["COLL_NO"].map(lambda x: get_college_by_no(str(x))["COLL_NAME"] if pd.notna(x) else "")
                out["COLL_NAMEM"] = out["COLL_NO"].map(lambda x: get_college_by_no(str(x))["COLL_NAMEM"] if pd.notna(x) else "")
            except ImportError:
                out["COLL_NAME"] = ""
                out["COLL_NAMEM"] = ""

            clean_identifier_columns(out, SENSITIVE_FIELDS)

            # --- Export to Excel ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                out.to_excel(writer, index=False, sheet_name="Structured")
                audit_log.to_excel(writer, index=False, sheet_name="Audit_Log")
            output.seek(0)
            st.download_button(
                "📥 Download Structured_Output_With_Audit.xlsx",
                data=output.getvalue(),
                file_name="Structured_Output_With_Audit.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.warning("📁 Please upload a DBF or Excel file to begin.")
