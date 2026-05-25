import ast
import os

import pandas as pd
import streamlit as st


def _dictionary_paths(base_dir):
    candidates = [
        os.path.join(base_dir, "dic.py"),
        os.path.join(os.path.dirname(base_dir), "dic.py"),
    ]
    seen = set()
    for path in candidates:
        path = os.path.abspath(path)
        if path not in seen and os.path.exists(path):
            seen.add(path)
            yield path


def _words_from_names(names):
    words = set()
    for name in names.dropna().astype(str):
        for word in name.upper().strip().split():
            if word and any(ch.isalpha() for ch in word):
                words.add(word)
    return words


def _parse_dictionary_line(line):
    stripped = line.strip()
    if not stripped.startswith('"') or ":" not in stripped:
        return None, None

    key_end = stripped.find('"', 1)
    if key_end == -1:
        return None, None

    key = stripped[1:key_end].upper()
    value_part = stripped[stripped.find(":", key_end) + 1:].strip().rstrip(",")

    try:
        value = ast.literal_eval(value_part)
    except (SyntaxError, ValueError):
        return None, None

    return key, value


@st.cache_data(show_spinner=False)
def _lookup_words(words_key, base_dir):
    pending = set(words_key)
    found = {}

    if not pending:
        return found

    for path in _dictionary_paths(base_dir):
        with open(path, "r", encoding="utf-8") as handle:
            for line in handle:
                stripped = line.lstrip()
                if not stripped.startswith('"'):
                    continue

                key_end = stripped.find('"', 1)
                if key_end == -1:
                    continue

                key = stripped[1:key_end].upper()
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


def translate_name_series(names, base_dir):
    words = _words_from_names(names)
    lookup = _lookup_words(tuple(sorted(words)), os.path.abspath(base_dir))

    def translate_name(name):
        translated = []
        for word in str(name).upper().strip().split():
            translated.append(lookup.get(word, word))
        return " ".join(translated)

    translated_names = names.fillna("").astype(str).apply(translate_name)
    missing_count = len(words - set(lookup))
    return translated_names, missing_count
