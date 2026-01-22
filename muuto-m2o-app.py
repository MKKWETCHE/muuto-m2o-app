import os
from io import BytesIO

import pandas as pd
import streamlit as st


# =============================
# Files / Paths
# =============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

RAW_DATA_XLSX_PATH = os.path.join(BASE_DIR, "raw-data.xlsx")
MASTERDATA_TEMPLATE_XLSX_PATH = os.path.join(BASE_DIR, "Masterdata-output-template.xlsx")
LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")

DEFAULT_NO_SELECTION = "--- Please Select ---"

# 🔒 Currency/Market is now hidden from UI
# Choose what makes most sense for you:
# - If you want a specific market only: set DEFAULT_CURRENCY to that exact value found in raw-data.xlsx currency/market column.
# - If you want ALL data regardless of market: set DEFAULT_CURRENCY = "ALL"
DEFAULT_CURRENCY = "ALL"  # e.g. "EURO" or "DKK" or "IE - EUR"


# =============================
# Helpers
# =============================
def safe_str(x):
    return "" if pd.isna(x) else str(x)


def normalize_key(s: str) -> str:
    return (
        str(s)
        .replace(" ", "_")
        .replace("/", "_")
        .replace("(", "")
        .replace(")", "")
        .replace("__", "_")
    )


def find_currency_column(df: pd.DataFrame) -> str | None:
    """
    Tries to detect which column in raw data indicates currency/market.
    Adjust here if your raw-data.xlsx uses a specific column name.
    """
    candidates = [
        "Currency",
        "currency",
        "Market",
        "market",
        "Currency/Market",
        "Market/Currency",
        "Price Currency",
        "Sales Currency",
    ]
    for c in candidates:
        if c in df.columns:
            return c

    lowered = {c.lower(): c for c in df.columns}
    for key in ["currency", "market", "currency/market", "market/currency"]:
        if key in lowered:
            return lowered[key]

    return None


def load_template_columns(template_path: str) -> list[str]:
