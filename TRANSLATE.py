import streamlit as st
import pandas as pd
import io
import zipfile
import re
import unicodedata

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

# -------------------------
# Import custom dictionary
# -------------------------
try:
    try:
        from dic import CUSTOM_MAP
    except ImportError:
        from dic import name_translation_dict as CUSTOM_MAP
    
    # OPTIMIZATION:
    # Pre-process dictionary keys using the same 'clean_text' logic
    # This ensures "  baid  " in dict matches "Baid" in file.
    DICT_LOOKUP = {}
    for k, v in CUSTOM_MAP.items():
        clean_key = clean_text(k)
        if clean_key:
            DICT_LOOKUP[clean_key] = str(v).strip()
            
    DICT_KEYS = set(DICT_LOOKUP.keys())
    dict_status = f"✅ Loaded {len(DICT_LOOKUP)} words from dic.py"
except Exception as e:
    import traceback
    st.error(f"Failed to load custom dictionary from dic.py: {e}")
    DICT_LOOKUP = {}
    DICT_KEYS = set()
    dict_status = "❌ Dictionary load failed"

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
        
    # Keep original for reconstruction, but work with normalized version for matching
    original_text = str(name_eng)
    
    # 2. SMART TOKENIZATION
    # Split by anything that is NOT a word char (letters/numbers/underscore).
    # Capturing group () keeps the separators in the list.
    tokens = re.split(r'(\W+)', original_text)
    
    # Check if ANY meaningful word exists in Dictionary (for existing Marathi logic)
    # We use our clean_text helper to match the dictionary keys exactly
    words_found_in_dict = False
    for t in tokens:
        if re.search(r'\w', t): # If it's a word
            if clean_text(t) in DICT_KEYS:
                words_found_in_dict = True
                break
    
    # 3. LOGIC: Preserve Existing Marathi?
    # If we have valid Marathi AND the dictionary doesn't have any corrections for this specific name
    # then we trust the existing translation to save processing/correction.
    has_existing_marathi = pd.notna(name_mar_existing) and str(name_mar_existing).strip() != ""
    
    if has_existing_marathi and not words_found_in_dict:
        return name_mar_existing

    # 4. BUILD NEW TRANSLATION
    translated_tokens = []
    
    for token in tokens:
        # Check if the token is a word (contains letters/numbers)
        if re.search(r'\w', token):
            # clean_text handles Case, Unicode, and Stripping for the Lookup
            token_key = clean_text(token)
            
            if token_key in DICT_LOOKUP:
                # Found match -> Use Marathi
                translated_tokens.append(DICT_LOOKUP[token_key])
            else:
                # No match -> Transliterate using Google Input Tools fallback
                translated_tokens.append(transliterate_word(token))
        else:
            # Punctuation/Spaces -> Keep exactly as is (preserves formatting)
            translated_tokens.append(token)
            
    return "".join(translated_tokens)

def run_translate_app():
    # Sidebar Info
    st.sidebar.markdown("### 🛠 System Status")
    st.sidebar.info(dict_status)
    st.sidebar.markdown("---")
    st.sidebar.markdown("**Features Active:**")
    st.sidebar.markdown("- ✅ Case Insensitivity")
    st.sidebar.markdown("- ✅ Unicode Normalization")
    st.sidebar.markdown("- ✅ Punctuation Handling")
    st.sidebar.markdown("- ✅ Empty Cell Protection")

    st.markdown("""
    ### How it works
    This tool is **Offline** and **Robust**. It safely handles:
    1.  **Messy Inputs:** `Rahul   Baid` (extra spaces) or `(Rahul)` (brackets).
    2.  **Case Issues:** `baid`, `BAID`, `Baid` all match.
    3.  **Weird Characters:** Handles invisible spaces or full-width characters.
    4.  **Logic:** Only translates if the word exists in `dic.py`.
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
                        
                        # Apply Translation
                        df["NAME_MARAT"] = df.apply(
                            lambda row: process_row(row["NAME"], row["NAME_MARAT"]), axis=1
                        )
                        
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