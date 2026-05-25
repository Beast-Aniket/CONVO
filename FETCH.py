import streamlit as st
import pandas as pd
from deep_translator import GoogleTranslator
import re
import io

def is_english_word(word):
    # Check if the word contains English alphabets
    return bool(re.match(r'^[A-Za-z]+$', word.strip()))

def run_fetch_app():
    st.markdown("## 🔍 Marathi Name Extractor & Dictionary Generator")
    st.markdown("""
    This utility helps you extract English words that are mixed in the Marathi name column (`NAME_MARAT`), 
    translates them online, and packages them into a fresh dictionary file format (`new_dic.py`) that you can append to your universal dictionary.
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

            if 'NAME_MARAT' not in df.columns:
                st.error("❌ Column 'NAME_MARAT' not found in the uploaded file. Please make sure the column exists.")
                return

            # Extract pending words
            pending_words = set()
            for name in df['NAME_MARAT'].dropna():
                words = str(name).split()
                for word in words:
                    if is_english_word(word):
                        pending_words.add(word.upper())

            if not pending_words:
                st.info("ℹ️ No pending English words found in 'NAME_MARAT' column. They are all correctly translated or empty.")
                return

            st.markdown(f"### 📋 Found `{len(pending_words)}` unique English words to translate")
            with st.expander("Show Extracted English Words", expanded=False):
                st.write(", ".join(sorted(list(pending_words))))

            col1, col2 = st.columns(2)
            
            # Button to trigger translation
            translate_btn = col1.button("🚀 Start Translation via Google Translate", key="start_fetch_translate_btn", type="primary", width="stretch")
            
            if translate_btn:
                progress_bar = st.progress(0)
                status_text = st.empty()
                translated_dict = {}

                translator = GoogleTranslator(source='en', target='mr')
                total_words = len(pending_words)

                for idx, word in enumerate(pending_words):
                    status_text.text(f"Translating {idx+1}/{total_words}: {word}")
                    try:
                        translated = translator.translate(word)
                        translated_dict[word] = translated
                    except Exception as e:
                        st.warning(f"Error translating {word}: {e}")
                        translated_dict[word] = word
                    progress_bar.progress((idx + 1) / total_words)

                status_text.success("🎉 Translation completed successfully!")

                # Generate dictionary string
                dict_content = "name_translation_dict = {\n"
                for eng, mar in translated_dict.items():
                    dict_content += f'    "{eng}": "{mar}",\n'
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
