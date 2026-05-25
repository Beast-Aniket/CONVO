from decimal import Decimal, InvalidOperation
import re

import pandas as pd


INTEGER_FLOAT_RE = re.compile(r"^[+-]?\d+\.0+$")
SCIENTIFIC_RE = re.compile(r"^[+-]?\d+(?:\.\d+)?[eE][+-]?\d+$")


def clean_identifier_value(value):
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""

    if INTEGER_FLOAT_RE.fullmatch(text):
        return text.split(".", 1)[0]

    if SCIENTIFIC_RE.fullmatch(text):
        try:
            number = Decimal(text)
            if number == number.to_integral_value():
                return format(number.quantize(Decimal(1)), "f")
        except (InvalidOperation, ValueError):
            pass

    return text


def clean_identifier_series(series):
    return series.apply(clean_identifier_value).astype(str)


def clean_identifier_columns(df, columns):
    for column in columns:
        if column in df.columns:
            df[column] = clean_identifier_series(df[column])
    return df
