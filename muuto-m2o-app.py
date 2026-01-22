import os
from io import BytesIO
import pandas as pd
import streamlit as st

# --------------------------------------------------
# Paths
# --------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DATA_XLSX_PATH = os.path.join(BASE_DIR, "raw-data.xlsx")
TEMPLATE_XLSX_PATH = os.path.join(BASE_DIR, "Masterdata-output-template.xlsx")
LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")

DEFAULT_NO_SELECTION = "--- Please Select ---"

# --------------------------------------------------
# Helpers
# --------------------------------------------------
def normalize_key(val: str) -> str:
    return (
        str(val)
        .replace(" ", "_")
        .replace("/", "_")
        .replace("(", "")
        .replace(")", "")
    )

def load_template_columns(template_path: str) -> list:
    """Load column headers from template (prices will be ignored later)."""
    if not os.path.exists(template_path):
        return []
    try:
        df = pd.read_excel(template_path, nrows=0)
        return list(df.columns)
    except Exception:
        return []

def remove_price_columns(cols: list) -> list:
    out = []
    for c in cols:
        lc = str(c).lower()
        if lc.startswith("wholesale price") or lc.startswith("retail price"):
            continue
        out.append(c)
    return out

def build_excel(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Masterdata")
    return buffer.getvalue()

# --------------------------------------------------
# Streamlit setup
# --------------------------------------------------
st.set_page_config(
    page_title="Muuto Made-to-Order Master Data Tool",
    layout="wide",
)

# --------------------------------------------------
# Session state
# --------------------------------------------------
for key, default in {
    "raw_df": None,
    "filtered_raw_df": pd.DataFrame(),
    "selected_family": DEFAULT_NO_SELECTION,
    "selected_items": {},
    "final_items": [],
    "template_cols": [],
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# --------------------------------------------------
# Header
# --------------------------------------------------
cols = st.columns([6, 1])
with cols[0]:
    st.title("Muuto Made-to-Order Master Data Tool")
with cols[1]:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_container_width=True)

st.markdown(
    """
    This tool allows you to select Made-to-Order products and generate master data.

    **Pricing and currency selection have been removed.**
    The tool is kept minimal until BC goes live.
    """
)

st.markdown("---")

# --------------------------------------------------
# Load data
# --------------------------------------------------
if not os.path.exists(RAW_DATA_XLSX_PATH):
    st.error("raw-data.xlsx not found.")
    st.stop()

st.session_state.raw_df = pd.read_excel(RAW_DATA_XLSX_PATH)
st.session_state.filtered_raw_df = st.session_state.raw_df.copy()

st.session_state.template_cols = remove_price_columns(
    load_template_columns(TEMPLATE_XLSX_PATH)
)

# --------------------------------------------------
# STEP 1 – Product selection (was Step 2)
# --------------------------------------------------
st.header("Step 1: Select Product Combinations")

df = st.session_state.filtered_raw_df

if "Product Family" not in df.columns:
    st.error("Column 'Product Family' missing in raw data.")
    st.stop()

families = [DEFAULT_NO_SELECTION] + sorted(df["Product Family"].dropna().unique())
st.session_state.selected_family = st.selectbox(
    "Select Product Family",
    families,
    index=families.index(st.session_state.selected_family)
    if st.session_state.selected_family in families
    else 0,
)

if st.session_state.selected_family == DEFAULT_NO_SELECTION:
    st.info("Select a product family to continue.")
    st.stop()

family_df = df[df["Product Family"] == st.session_state.selected_family]

if not {"Product Display Name", "Upholstery Type", "Upholstery Color"}.issubset(family_df.columns):
    st.error("Required columns missing in raw data.")
    st.stop()

products = sorted(family_df["Product Display Name"].unique())
upholstery_types = sorted(family_df["Upholstery Type"].unique())

st.markdown("### Select items")

for product in products:
    with st.expander(product):
        for uph_type in upholstery_types:
            colors = family_df[
                (family_df["Product Display Name"] == product)
                & (family_df["Upholstery Type"] == uph_type)
            ]["Upholstery Color"].unique()

            for color in colors:
                key = normalize_key(f"{product}_{uph_type}_{color}")
                st.checkbox(
                    f"{uph_type} – {color}",
                    key=key,
                )

# --------------------------------------------------
# STEP 2 – Review selections
# --------------------------------------------------
st.markdown("---")
st.header("Step 2: Review Selections")

selected = [
    key for key, val in st.session_state.items()
    if key not in ["raw_df", "filtered_raw_df", "selected_family", "template_cols"]
    and val is True
]

if not selected:
    st.info("No items selected yet.")
else:
    st.session_state.final_items = []
    for key in selected:
        parts = key.split("_")
        product, uph_type, color = parts[0], parts[1], parts[2]

        match = family_df[
            (family_df["Product Display Name"] == product)
            & (family_df["Upholstery Type"] == uph_type)
            & (family_df["Upholstery Color"] == color)
        ]

        if not match.empty:
            row = match.iloc[0]
            st.write(f"• {product} / {uph_type} / {color} (Item {row['Item No']})")
            st.session_state.final_items.append(row)

# --------------------------------------------------
# STEP 3 – Download
# --------------------------------------------------
st.markdown("---")
st.header("Step 3: Generate Master Data File")

if st.session_state.final_items:
    output_df = pd.DataFrame(st.session_state.final_items)

    if st.session_state.template_cols:
        output_df = output_df[
            [c for c in st.session_state.template_cols if c in output_df.columns]
        ]

    excel_bytes = build_excel(output_df)

    st.download_button(
        "Generate and Download Master Data File",
        excel_bytes,
        file_name="masterdata_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.button("Generate Master Data File", disabled=True)
