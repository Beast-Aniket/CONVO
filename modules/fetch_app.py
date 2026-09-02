import io
import re
import unicodedata
import pandas as pd
import streamlit as st
from core.name_lookup import _transliterate_word_uncached

WORD_RE = re.compile(r"[A-Za-z]+")

def clean_token(token):
    return unicodedata.normalize("NFKC", str(token)).strip().upper()

def extract_english_words_from_text(text):
    words = set()
    if pd.isna(text):
        return words
    for match in WORD_RE.finditer(str(text)):
        w = clean_token(match.group(0))
        if w:
            words.add(w)
    return words

def run_fetch_app():
    st.markdown("## 🔍 Marathi Name Extractor & Dictionary Generator")
    st.markdown("""
    This utility helps you extract English words that remain un-translated in the Marathi name column (`NAME_MARAT`), 
    transliterates them phonetically using **Google Input Tools**, and generates an updated dictionary format (`new_dic.py`) 
    that you can append directly to your universal dictionary.
    """)

    uploaded_file = st.file_uploader(
        "Upload Excel or CSV file containing Marathi names", 
        type=["csv", "xlsx", "xls"],
        key="fetch_uploader"
    )

    if uploaded_file:
        try:
            # Load file
            file_name = uploaded_file.name.lower()
            if file_name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, dtype=str)

            st.success(f"✅ Loaded file with {len(df)} rows")
            st.dataframe(df.head(10), width="stretch")

            # Check for Marathi column name (flexible casing)
            marathi_col = None
            for col in df.columns:
                if str(col).strip().upper() in ["NAME_MARAT", "NAMEMARAT", "NAME_MARATHI", "NAMEMARATHI"]:
                    marathi_col = col
                    break

            if not marathi_col:
                st.error("❌ Column 'NAME_MARAT' not found in the uploaded file. Please make sure the column exists.")
                return

            # Extract pending English words
            pending_words = set()
            for name in df[marathi_col].dropna():
                pending_words.update(extract_english_words_from_text(name))

            if not pending_words:
                st.info("ℹ️ No pending English words found in the Marathi name column. All names are in Devanagari or empty.")
                return

            st.markdown(f"### 📋 Found `{len(pending_words)}` unique English words to transliterate")
            with st.expander("Show Extracted English Words", expanded=False):
                st.write(", ".join(sorted(list(pending_words))))

            col1, col2 = st.columns(2)
            
            # Button to trigger transliteration
            translate_btn = col1.button("🚀 Start Transliteration via Google Input Tools", key="start_fetch_translate_btn", type="primary", width="stretch")
            
            if translate_btn:
                progress_bar = st.progress(0)
                status_text = st.empty()
                translated_dict = {}

                sorted_pending = sorted(list(pending_words))
                total_words = len(sorted_pending)

                for idx, word in enumerate(sorted_pending):
                    status_text.text(f"Transliterating {idx+1}/{total_words}: {word}")
                    try:
                        transliterated = _transliterate_word_uncached(word)
                        translated_dict[word] = transliterated
                    except Exception as e:
                        st.warning(f"Error transliterating {word}: {e}")
                        translated_dict[word] = word
                    progress_bar.progress((idx + 1) / total_words)

                status_text.success("🎉 Transliteration completed successfully!")

                # Generate dictionary string
                dict_content = "name_translation_dict = {\n"
                for eng, mar in translated_dict.items():
                    escaped_mar = str(mar).replace("\\", "\\\\").replace('"', '\\"')
                    dict_content += f'    "{eng}": "{escaped_mar}",\n'
                dict_content += "}\n"

                # Provide download button
                st.markdown("### 📥 Step 3: Download Generated Dictionary")
                st.download_button(
                    label="📥 Download new_dic.py",
                    data=dict_content.encode('utf-8'),
                    file_name="new_dic.py",
                    mime="text/x-python",
                    key="fetch_download_btn",
                    width="stretch"
                )

        except Exception as e:
            st.error(f"Error loading or processing file: {e}")
