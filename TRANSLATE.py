import streamlit as st
import pandas as pd
import io
import os
import zipfile
import re
import unicodedata
from name_lookup import translate_name_series, universal_dictionary_status

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

# Session cache to store API-transliterated names and avoid lag
TRANSLIT_CACHE = {}

def transliterate_word(word):
    word = str(word).upper().strip()
    if not word:
        return ""
    if word in TRANSLIT_CACHE:
        return TRANSLIT_CACHE[word]
    
    # Check if word contains any alphabets (ignore pure digits/special chars)
    if not any(c.isalpha() for c in word):
        return word
        
    try:
        import urllib.request
        import urllib.parse
        import json
        url = "https://inputtools.google.com/request?" + urllib.parse.urlencode({
            "text": word,
            "itc": "mr-t-i0-und",
            "num": "1",
            "cp": "0",
            "cs": "1",
            "ie": "utf-8",
            "oe": "utf-8",
            "app": "test"
        })
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=3) as response:
            res = json.loads(response.read().decode('utf-8'))
            if res[0] == 'SUCCESS':
                transliterated = res[1][0][1][0]
                TRANSLIT_CACHE[word] = transliterated
                return transliterated
    except Exception:
        pass
    
    return word

# -------------------------
# CORE TRANSLATION LOGIC
# -------------------------
def process_row(name_eng, name_mar_existing):
    """
    Robust translation logic:
    1. Preserves original formatting (spaces/punctuation).
    2. Translates individual words if found in dictionary.
    3. Handles mixed data types gracefully.
    """
    # 1. Safety check for empty/NaN
    if pd.isna(name_eng) or str(name_eng).strip() == "":
        return ""

    translated, _ = translate_name_series(
        pd.Series([name_eng]),
        pd.Series([name_mar_existing]),
        preserve_existing=True,
    )
    return translated.iloc[0]

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
