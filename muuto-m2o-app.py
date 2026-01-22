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

def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Pick the first matching column from candidates."""
    for c in candidates:
        if c in df.columns:
            return c
    # Try case-insensitive match
    lower_map = {str(col).strip().lower(): col for col in df.columns}
    for c in candidates:
        key = str(c).strip().lower()
        if key in lower_map:
            return lower_map[key]
    return None

def load_template_columns(template_path: str) -> list[str]:
    """Load columns from template header (row 1)."""
    if not os.path.exists(template_path):
        return []
    try:
        df = pd.read_excel(template_path, nrows=0, engine="openpyxl")
        return list(df.columns)
    except Exception:
        return []

def remove_price_columns(cols: list[str]) -> list[str]:
    """Drop any wholesale/retail price columns (incl variants)."""
    out = []
    for c in cols:
        lc = str(c).strip().lower()
        if lc.startswith("wholesale price") or lc.startswith("retail price"):
            continue
        out.append(c)
    return out

def build_excel(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Masterdata")
    return buffer.getvalue()

def safe_series_str(s: pd.Series) -> pd.Series:
    return s.astype(str).fillna("")

# --------------------------------------------------
# Streamlit setup
# --------------------------------------------------
st.set_page_config(page_title="Muuto Made-to-Order Master Data Tool", layout="wide")

# --------------------------------------------------
# Header
# --------------------------------------------------
top = st.columns([6, 1])
with top[0]:
    st.title("Muuto Made-to-Order Master Data Tool")
with top[1]:
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
    st.error(f"Missing file: {RAW_DATA_XLSX_PATH}")
    st.stop()

try:
    raw_df = pd.read_excel(RAW_DATA_XLSX_PATH, engine="openpyxl")
except Exception as e:
    st.error(f"Could not read raw-data.xlsx: {e}")
    st.stop()

# No currency filtering anymore – use all data
df = raw_df.copy()

# Load template columns (optional)
template_cols = remove_price_columns(load_template_columns(TEMPLATE_XLSX_PATH))

# --------------------------------------------------
# Column mapping (aliases)
# --------------------------------------------------
family_col = pick_col(df, ["Product Family", "Family", "Product_Family"])
product_col = pick_col(df, ["Product Display Name", "Item Name", "Product", "Item", "Product Name"])
uph_type_col = pick_col(df, ["Upholstery Type", "Upholstery", "Upholstery Type Cleaned", "Upholstery_Type"])
uph_color_col = pick_col(df, ["Upholstery Color", "Upholstery Colour", "Upholstery Color No", "Upholstery Color Number", "Upholstery_Color"])
item_no_col = pick_col(df, ["Item No", "Item No.", "Item Number", "Item_No", "ItemNo"])
article_no_col = pick_col(df, ["Article No", "Article No.", "Article Number", "Article_No", "ArticleNo"])

missing = []
if not family_col: missing.append("Product Family")
if not product_col: missing.append("Product Display Name / Item Name")
if not uph_type_col: missing.append("Upholstery Type")
if not uph_color_col: missing.append("Upholstery Color")

if missing:
    st.error("Required columns missing in raw-data.xlsx: " + ", ".join(missing))
    st.write("Columns found in raw-data.xlsx:")
    st.write(list(df.columns))
    st.stop()

# Normalize to strings for filtering/comparison safety
df[family_col] = safe_series_str(df[family_col])
df[product_col] = safe_series_str(df[product_col])
df[uph_type_col] = safe_series_str(df[uph_type_col])
df[uph_color_col] = safe_series_str(df[uph_color_col])

# --------------------------------------------------
# SESSION STATE
# --------------------------------------------------
if "selected_family" not in st.session_state:
    st.session_state.selected_family = DEFAULT_NO_SELECTION

if "selected_keys" not in st.session_state:
    st.session_state.selected_keys = set()

# --------------------------------------------------
# STEP 1 – Product selection (currency removed)
# --------------------------------------------------
st.header("Step 1: Select Product Combinations")

families = [DEFAULT_NO_SELECTION] + sorted([f for f in df[family_col].dropna().unique() if str(f).strip()])
if st.session_state.selected_family not in families:
    st.session_state.selected_family = DEFAULT_NO_SELECTION

selected_family = st.selectbox(
    "Select Product Family",
    families,
    index=families.index(st.session_state.selected_family),
)
st.session_state.selected_family = selected_family

if selected_family == DEFAULT_NO_SELECTION:
    st.info("Select a product family to continue.")
    st.stop()

family_df = df[df[family_col].astype(str) == str(selected_family)].copy()

products = sorted([p for p in family_df[product_col].dropna().unique() if str(p).strip()])
uph_types = sorted([u for u in family_df[uph_type_col].dropna().unique() if str(u).strip()])

if not products or not uph_types:
    st.warning("No products or upholstery types found for this family.")
    st.stop()

# Build selection UI
st.markdown("### Select items")
st.caption("Expand a product and tick the upholstery color combinations you want.")

for prod in products:
    with st.expander(prod):
        prod_df = family_df[family_df[product_col].astype(str) == str(prod)]
        for ut in uph_types:
            ut_df = prod_df[prod_df[uph_type_col].astype(str) == str(ut)]
            if ut_df.empty:
                continue
            colors = sorted([c for c in ut_df[uph_color_col].dropna().unique() if str(c).strip()])

            if not colors:
                continue

            st.markdown(f"**{ut}**")
            for col in colors:
                key = normalize_key(f"{selected_family}__{prod}__{ut}__{col}")
                checked = key in st.session_state.selected_keys

                new_val = st.checkbox(
                    f"{col}",
                    value=checked,
                    key=f"cb_{key}",
                )

                if new_val:
                    st.session_state.selected_keys.add(key)
                else:
                    st.session_state.selected_keys.discard(key)

# --------------------------------------------------
# STEP 2 – Review selections
# --------------------------------------------------
st.markdown("---")
st.header("Step 2: Review Selections")

selected_items = []
for k in sorted(st.session_state.selected_keys):
    parts = k.split("__")
    if len(parts) != 4:
        continue
    fam, prod, ut, col = parts

    m = family_df[
        (family_df[product_col].astype(str) == str(prod))
        & (family_df[uph_type_col].astype(str) == str(ut))
        & (family_df[uph_color_col].astype(str) == str(col))
    ]

    if m.empty:
        continue

    row = m.iloc[0]

    item_no = row.get(item_no_col) if item_no_col else None
    article_no = row.get(article_no_col) if article_no_col else None

    selected_items.append(row)

    st.write(
        f"• {prod} / {ut} / {col}"
        + (f" (Item: {item_no})" if item_no is not None and str(item_no).strip() else "")
        + (f" (Article: {article_no})" if article_no is not None and str(article_no).strip() else "")
    )

if not selected_items:
    st.info("No items selected yet.")
    st.markdown("---")
    st.header("Step 3: Generate Master Data File")
    st.button("Generate Master Data File", disabled=True)
    st.stop()

# --------------------------------------------------
# STEP 3 – Download
# --------------------------------------------------
st.markdown("---")
st.header("Step 3: Generate Master Data File")

output_df = pd.DataFrame(selected_items)

# Apply template ordering (optional)
if template_cols:
    # Keep only template columns that exist in output_df
    cols_in_df = [c for c in template_cols if c in output_df.columns]
    if cols_in_df:
        output_df = output_df[cols_in_df]

# Ensure price columns are not present even if they existed in raw/template
output_df = output_df[[c for c in output_df.columns if not str(c).lower().startswith(("wholesale price", "retail price"))]]

excel_bytes = build_excel(output_df)

st.download_button(
    "Generate and Download Master Data File",
    data=excel_bytes,
    file_name="masterdata_output.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
