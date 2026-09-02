import ast
import json
import os
import re
import unicodedata
import urllib.parse
import urllib.request
import pandas as pd
import streamlit as st

WORD_RE = re.compile(r"[A-Za-z]+")

def universal_dictionary_path():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    candidates = [
        os.path.join(base_dir, "data", "dic.py"),
        os.path.join(base_dir, "dic.py"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "dic.py"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return candidates[0]

def universal_dictionary_available():
    return os.path.exists(universal_dictionary_path())

def universal_dictionary_status():
    path = universal_dictionary_path()
    if not os.path.exists(path):
        return "dic.py not found"
    size_mb = os.path.getsize(path) / (1024 * 1024)
    return f"dic.py available ({size_mb:.1f} MB)"

def _clean_word(word):
    return unicodedata.normalize("NFKC", str(word)).strip().upper()

def _words_from_names(names):
    words = set()
    for name in names.dropna().astype(str):
        for match in WORD_RE.finditer(name):
            word = _clean_word(match.group(0))
            if word:
                words.add(word)
    return words

def _parse_dictionary_line(line):
    stripped = line.strip()
    if not stripped.startswith('"') or ":" not in stripped:
        return None, None

    key_end = stripped.find('"', 1)
    if key_end == -1:
        return None, None

    key = _clean_word(stripped[1:key_end])
    value_part = stripped[stripped.find(":", key_end) + 1:].strip().rstrip(",")

    try:
        value = ast.literal_eval(value_part)
    except (SyntaxError, ValueError):
        return None, None

    return key, value

@st.cache_data(show_spinner=False)
def _lookup_words(words_key):
    pending = set(words_key)
    found = {}

    if not pending:
        return found

    path = universal_dictionary_path()
    if not os.path.exists(path):
        return found

    with open(path, "r", encoding="utf-8") as handle:
        for line in handle:
            stripped = line.lstrip()
            if not stripped.startswith('"'):
                continue

            key_end = stripped.find('"', 1)
            if key_end == -1:
                continue

            key = _clean_word(stripped[1:key_end])
            if key not in pending:
                continue

            _, value = _parse_dictionary_line(stripped)
            if value is None:
                continue

            found[key] = value
            pending.remove(key)
            if not pending:
                return found

    return found

def _transliterate_word_uncached(word):
    if not word:
        return ""

    try:
        url = "https://inputtools.google.com/request?" + urllib.parse.urlencode({
            "text": word,
            "itc": "mr-t-i0-und",
            "num": "1",
            "cp": "0",
            "cs": "1",
            "ie": "utf-8",
            "oe": "utf-8",
            "app": "test",
        })
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=3) as response:
            result = json.loads(response.read().decode("utf-8"))
            if result[0] == "SUCCESS":
                return result[1][0][1][0]
    except Exception:
        pass

    return word

@st.cache_data(show_spinner=False)
def _transliterate_words(words_key):
    return {word: _transliterate_word_uncached(word) for word in words_key}

def translate_name_series(names, existing_marathi=None, preserve_existing=False):
    words = _words_from_names(names)
    lookup = _lookup_words(tuple(sorted(words)))
    missing_words = words - set(lookup)
    fallback = _transliterate_words(tuple(sorted(missing_words)))

    existing_values = None
    if existing_marathi is not None:
        existing_values = existing_marathi.fillna("").astype(str)

    def translate_text(text, existing_text=""):
        text = "" if text is None else str(text)
        if not text.strip():
            return ""

        row_words = {_clean_word(match.group(0)) for match in WORD_RE.finditer(text)}
        if preserve_existing and existing_text.strip() and row_words.isdisjoint(lookup):
            return existing_text

        def replace_match(match):
            original = match.group(0)
            key = _clean_word(original)
            return lookup.get(key) or fallback.get(key) or original

        return WORD_RE.sub(replace_match, text)

    if existing_values is None:
        translated = names.fillna("").astype(str).apply(translate_text)
    else:
        translated = names.fillna("").astype(str).combine(existing_values, translate_text)

    return translated, len(missing_words)

def transliterate_text(text):
    """
    Transliterates general text phrases (such as PER periods like 'March 2025') 
    to Marathi using Google Input Tools API.
    """
    if text is None:
        return ""
    text_str = str(text).strip()
    if not text_str or text_str.lower() == "nan":
        return ""

    try:
        url = "https://inputtools.google.com/request?" + urllib.parse.urlencode({
            "text": text_str,
            "itc": "mr-t-i0-und",
            "num": "1",
            "cp": "0",
            "cs": "1",
            "ie": "utf-8",
            "oe": "utf-8",
            "app": "test",
        })
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=5) as response:
            result = json.loads(response.read().decode("utf-8"))
            if result[0] == "SUCCESS":
                return " ".join(item[1][0] for item in result[1])
    except Exception:
        pass

    # Fallback: convert English digits to Marathi numerals
    trans_digits = str.maketrans("0123456789", "०१२३४५६७८९")
    return text_str.translate(trans_digits)
