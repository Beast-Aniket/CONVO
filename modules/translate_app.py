import streamlit as st
import pandas as pd
import io
import os
import zipfile
import re
import unicodedata
from core.name_lookup import translate_name_series, universal_dictionary_status

# -------------------------
# HELPER: Normalization
# -------------------------
def clean_text(text):
    """
    Smart cleaning:
    1. Convert to String.
    2. Normalize Unicode (fixes weird spaces, full-width chars, accents).
    3. Strip whitespace.
    4. Convert to Upper Case.
    """
    if pd.isna(text):
        return ""
    text = str(text)
    # NFKC Normalization fixes "full-width" characters often found in imported data
    # e.g., "ＢＡＩＤ" (Full width) becomes "BAID" (Standard)
    text = unicodedata.normalize('NFKC', text)
    return text.strip().upper()

def dictionary_status():
    return universal_dictionary_status()

def run_translate_app():
    # Sidebar Info
    st.sidebar.markdown("### 🛠 System Status")
    st.sidebar.info(dictionary_status())
    st.sidebar.markdown("---")
    st.sidebar.markdown("**Features Active:**")
    st.sidebar.markdown("- ✅ Case Insensitivity")
    st.sidebar.markdown("- ✅ Unicode Normalization")
    st.sidebar.markdown("- ✅ Punctuation Handling")
    st.sidebar.markdown("- ✅ Empty Cell Protection")

    st.markdown("""
    ### How it works
    This tool uses the universal root `dic.py` first, then Google Input Tools for missing English words. It safely handles:
    1.  **Messy Inputs:** `Rahul   Baid` (extra spaces) or `(Rahul)` (brackets).
    2.  **Case Issues:** `baid`, `BAID`, `Baid` all match.
    3.  **Weird Characters:** Handles invisible spaces or full-width characters.
    4.  **Fallback:** Words not found in `dic.py` are transliterated online.
    """)

    uploaded_files = st.file_uploader("📤 Upload Excel files (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="translate_uploader")

    if uploaded_files and st.button("🚀 Process Files", key="translate_process_btn"):
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            total_files = len(uploaded_files)
            
            for i, file in enumerate(uploaded_files):
                status_text.text(f"Processing {i+1}/{total_files}: {file.name}")
                
                try:
                    # Load Excel as String (prevents date/number conversion issues)
                    df = pd.read_excel(file, dtype=str)
                    df.columns = df.columns.str.strip() 
                    
                    if "NAME" in df.columns:
                        if "NAME_MARAT" not in df.columns:
                            df["NAME_MARAT"] = ""
                        
                        df["NAME_MARAT"], fallback_count = translate_name_series(
                            df["NAME"],
                            df["NAME_MARAT"],
                            preserve_existing=True,
                        )
                        if fallback_count:
                            st.info(f"{fallback_count} unique name words used Google Input Tools fallback.")
                        
                        # Save
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            df.to_excel(writer, index=False)
                        
                        zf.writestr(f"updated_{file.name}", output.getvalue())
                        
                    else:
                        # Passthrough for files without NAME column
                        file.seek(0)
                        zf.writestr(file.name, file.read())
                        
                except Exception as e:
                    st.error(f"Error processing {file.name}: {e}")
                
                progress_bar.progress((i + 1) / total_files)

        status_text.success("✅ All files processed successfully!")
        
        st.download_button(
            label="📦 Download Results (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="smart_translated_files.zip",
            mime="application/zip",
            key="translate_download_btn"
        )
