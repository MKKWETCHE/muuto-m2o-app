import os
from io import BytesIO

import pandas as pd
import streamlit as st

# -----------------------------
# Paths
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DATA_XLSX_PATH = os.path.join(BASE_DIR, "raw-data.xlsx")
MASTERDATA_TEMPLATE_XLSX_PATH = os.path.join(BASE_DIR, "Masterdata-output-template.xlsx")
LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")

DEFAULT_NO_SELECTION = "--- Please Select ---"

# Optional: keep a stable currency list, since we no longer read currencies from price matrices
CURRENCY_OPTIONS_FALLBACK = [
    "DACH - EURO", "DKK", "EURO", "NOK", "PLN", "SEK", "AUD", "GBP", "IE - EUR"
]

# -----------------------------
# Helpers
# -----------------------------
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
    We keep this flexible because raw-data.xlsx may vary.
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
    # Try fuzzy match
    lowered = {c.lower(): c for c in df.columns}
    for key in ["currency", "market", "currency/market", "market/currency"]:
        if key in lowered:
            return lowered[key]
    return None

def load_template_columns(template_path: str) -> list[str]:
    """
    Reads the template and returns the header columns.
    If template doesn't exist, we will fallback later to raw df columns.
    """
    if not os.path.exists(template_path):
        return []
    try:
        # Read only header
        tmp = pd.read_excel(template_path, engine="openpyxl", nrows=0)
        return list(tmp.columns)
    except Exception:
        return []

def remove_price_columns(cols: list[str]) -> list[str]:
    """
    Drops "Wholesale price" / "Retail price" columns, including variants with currency in parentheses.
    """
    out = []
    for c in cols:
        lc = str(c).strip().lower()
        if lc.startswith("wholesale price") or lc.startswith("retail price"):
            continue
        out.append(c)
    return out

def prepare_excel_bytes(output_df: pd.DataFrame, template_path: str) -> bytes:
    """
    Writes output_df to an .xlsx in memory.
    If template exists, we keep the header order from template (minus price cols).
    """
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        output_df.to_excel(writer, index=False, sheet_name="Masterdata")
    return bio.getvalue()

def get_required_columns_warning(df: pd.DataFrame):
    needed = ["Product Family", "Upholstery Type", "Upholstery Color", "Item No", "Article No"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        st.warning(
            "Raw-data.xlsx mangler nogle forventede kolonner: "
            + ", ".join(missing)
            + ". Appen forsøger stadig at køre, men visning/udvalg kan blive begrænset."
        )

# -----------------------------
# Streamlit setup
# -----------------------------
st.set_page_config(page_title="Muuto Made-to-Order Master Data Tool", layout="wide")

# Basic styling (keep it simple & stable)
st.markdown(
    """
    <style>
      .stApp, body { background-color: #EFEEEB !important; }
      .main .block-container { background-color: #EFEEEB !important; padding-top: 2rem; }
      h1, h2, h3 { text-transform: none !important; }
      h1 { color: #333; }
      h2, h3 { color: #1E40AF; }
      .note-box {
        background: #fff;
        border: 1px solid #ddd;
        padding: 12px 14px;
        border-radius: 8px;
        margin-top: 8px;
      }
      .swatch-placeholder {
        width:25px; height:25px;
        border: 1px dashed #ddd;
        background-color: #f9f9f9;
        display: inline-block;
      }
      .product-name-cell { font-weight: 600; }
      .select-all-label { font-weight: 600; font-size: 0.85em; text-align: right; padding-right: 6px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# Session state
# -----------------------------
if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "filtered_raw_df" not in st.session_state:
    st.session_state.filtered_raw_df = pd.DataFrame()
if "selected_currency_session" not in st.session_state:
    st.session_state.selected_currency_session = None
if "selected_family_session" not in st.session_state:
    st.session_state.selected_family_session = DEFAULT_NO_SELECTION
if "matrix_selected_generic_items" not in st.session_state:
    st.session_state.matrix_selected_generic_items = {}
if "user_chosen_base_colors_for_items" not in st.session_state:
    st.session_state.user_chosen_base_colors_for_items = {}
if "final_items_for_download" not in st.session_state:
    st.session_state.final_items_for_download = []
if "template_cols" not in st.session_state:
    st.session_state.template_cols = []

# -----------------------------
# Header
# -----------------------------
top_col1, top_col_spacer, top_col2 = st.columns([5.5, 0.5, 1])

with top_col1:
    st.title("Muuto Made-to-Order Master Data Tool")

with top_col2:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_container_width=True)

st.markdown(
    """
    Welcome to Muuto's Made-to-Order (MTO) Product Configurator!

    This tool simplifies selecting MTO products and generating the data you need for your systems. Here's how it works:

    * **Step 1: Select Currency**
      * Choose your preferred currency/market for filtering available products (prices are NOT generated).
    * **Step 2: Select Product Combinations**
      * Choose a product family to view its available products and upholstery options.
      * Select your desired product, upholstery, and color combinations directly in the matrix.
      * Use the **"Select All"** checkbox at the top of an upholstery color column to select/deselect all available products in that column.
    * **Step 2a: Specify Base Colors**
      * For items requiring base colors, they are grouped by product family. You can apply a base color to all applicable items in a family, or choose per item.
    * **Step 3: Review Selections**
      * Review the final list and remove items if needed.
    * **Step 4: Generate Master Data File**
      * Download an Excel file with master data for the selected items.
    """,
)

st.markdown(
    """
    <div class="note-box">
      <b>Note (prices removed):</b>
      This version intentionally does <u>not</u> generate Wholesale/Retail prices to avoid outdated pricing.
      Currency selection is used only for filtering products/variants, not price output.
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("---")

# -----------------------------
# Load data
# -----------------------------
files_loaded_successfully = True

if not os.path.exists(RAW_DATA_XLSX_PATH):
    st.error(f"Missing raw data file: {RAW_DATA_XLSX_PATH}")
    files_loaded_successfully = False

# Template is optional
st.session_state.template_cols = load_template_columns(MASTERDATA_TEMPLATE_XLSX_PATH)
st.session_state.template_cols = remove_price_columns(st.session_state.template_cols)

if files_loaded_successfully:
    try:
        st.session_state.raw_df = pd.read_excel(RAW_DATA_XLSX_PATH, engine="openpyxl")
        # Normalize some expected string columns if they exist
        for col in ["Upholstery Color", "Upholstery Type", "Product Family", "Product Display Name"]:
            if col in st.session_state.raw_df.columns:
                st.session_state.raw_df[col] = st.session_state.raw_df[col].astype(str).fillna("")

        get_required_columns_warning(st.session_state.raw_df)

        # If template is missing, fallback to raw columns (minus price cols)
        if not st.session_state.template_cols:
            st.session_state.template_cols = remove_price_columns(list(st.session_state.raw_df.columns))

    except Exception as e:
        st.error(f"Failed to load raw-data.xlsx: {e}")
        files_loaded_successfully = False

# -----------------------------
# Step 1: Currency
# -----------------------------
st.header("Step 1: Select Currency")

if not files_loaded_successfully:
    st.stop()

raw_df = st.session_state.raw_df
currency_col = find_currency_column(raw_df)

if currency_col:
    # Currency list based on raw data column values + fallback
    raw_currencies = sorted(
        [c for c in raw_df[currency_col].dropna().astype(str).unique().tolist() if c.strip()]
    )
    merged = sorted(list(dict.fromkeys(raw_currencies + CURRENCY_OPTIONS_FALLBACK)))
    currency_options = [DEFAULT_NO_SELECTION] + merged
else:
    currency_options = [DEFAULT_NO_SELECTION] + CURRENCY_OPTIONS_FALLBACK

current_currency_idx = 0
if st.session_state.selected_currency_session and st.session_state.selected_currency_session in currency_options:
    current_currency_idx = currency_options.index(st.session_state.selected_currency_session)

prev_selected_currency = st.session_state.selected_currency_session
selected_currency_choice = st.selectbox(
    "Select Currency:",
    options=currency_options,
    index=current_currency_idx,
    key="currency_selector_main_key",
)

# Apply currency filter
try:
    if selected_currency_choice == DEFAULT_NO_SELECTION:
        st.session_state.selected_currency_session = None
        st.session_state.filtered_raw_df = pd.DataFrame(columns=raw_df.columns)
    else:
        st.session_state.selected_currency_session = selected_currency_choice
        if currency_col:
            st.session_state.filtered_raw_df = raw_df[raw_df[currency_col].astype(str) == str(selected_currency_choice)].copy()
        else:
            # No currency column available -> no filter (but still allow the flow)
            st.session_state.filtered_raw_df = raw_df.copy()

    # Reset selections if currency changed
    if prev_selected_currency and prev_selected_currency != st.session_state.selected_currency_session:
        st.session_state.matrix_selected_generic_items = {}
        st.session_state.user_chosen_base_colors_for_items = {}
        st.session_state.final_items_for_download = []
        st.session_state.selected_family_session = DEFAULT_NO_SELECTION

except Exception as e:
    st.error(f"Error processing currency selection/filtering: {e}")
    st.session_state.selected_currency_session = None
    st.session_state.filtered_raw_df = pd.DataFrame(columns=raw_df.columns)

st.markdown("---")

# -----------------------------
# Step 2: Matrix selection
# -----------------------------
st.header("Step 2: Select Product Combinations (Product / Upholstery / Color)")

if not st.session_state.selected_currency_session:
    st.info("Please select a currency in Step 1 to see available products.")
    st.markdown("---")
else:
    df_for_display = st.session_state.filtered_raw_df

    if df_for_display.empty:
        st.warning("No rows available after filtering. Check raw data and/or currency mapping.")
        st.markdown("---")
    else:
        # Family dropdown
        if "Product Family" in df_for_display.columns:
            available_families = [DEFAULT_NO_SELECTION] + sorted(df_for_display["Product Family"].dropna().astype(str).unique().tolist())
        else:
            available_families = [DEFAULT_NO_SELECTION]

        if st.session_state.selected_family_session not in available_families:
            st.session_state.selected_family_session = DEFAULT_NO_SELECTION

        selected_family = st.selectbox(
            "Select Product Family:",
            options=available_families,
            index=available_families.index(st.session_state.selected_family_session),
            key="family_selector_key",
        )
        st.session_state.selected_family_session = selected_family

        def handle_matrix_cb_toggle(prod_name, uph_type, uph_color, checkbox_key_matrix):
            is_checked = st.session_state.get(checkbox_key_matrix, False)
            current_family = st.session_state.selected_family_session
            generic_item_key = normalize_key(f"{current_family}_{prod_name}_{uph_type}_{uph_color}")

            if is_checked:
                family_df = st.session_state.filtered_raw_df
                item_exists_df = family_df[
                    (family_df.get("Product Display Name", family_df.get("Item Name", "")).astype(str) == str(prod_name))
                    & (family_df["Upholstery Type"].astype(str) == str(uph_type))
                    & (family_df["Upholstery Color"].astype(str).fillna("N/A") == str(uph_color))
                ] if ("Upholstery Type" in family_df.columns and "Upholstery Color" in family_df.columns) else pd.DataFrame()

                if item_exists_df.empty:
                    st.toast("Could not find item row for this selection.", icon="⚠️")
                    return

                # Base color handling
                base_col = "Base Color Cleaned" if "Base Color Cleaned" in item_exists_df.columns else None
                unique_base_colors = item_exists_df[base_col].dropna().astype(str).unique().tolist() if base_col else []
                requires_base_choice = len(unique_base_colors) > 1
                resolved_base_if_single = unique_base_colors[0] if len(unique_base_colors) == 1 else (pd.NA if not unique_base_colors else None)

                first_match = item_exists_df.iloc[0]
                item_data = {
                    "key": generic_item_key,
                    "family": current_family,
                    "product": prod_name,
                    "upholstery_type": uph_type,
                    "upholstery_color": uph_color,
                    "requires_base_choice": requires_base_choice,
                    "available_bases": unique_base_colors,
                    "resolved_base_if_single": resolved_base_if_single,
                    # If there is only one base (or none), we can resolve item/article directly from that row:
                    "item_no_if_single_base": first_match.get("Item No", None),
                    "article_no_if_single_base": first_match.get("Article No", None),
                }
                st.session_state.matrix_selected_generic_items[generic_item_key] = item_data

            else:
                if generic_item_key in st.session_state.matrix_selected_generic_items:
                    del st.session_state.matrix_selected_generic_items[generic_item_key]
                if generic_item_key in st.session_state.user_chosen_base_colors_for_items:
                    del st.session_state.user_chosen_base_colors_for_items[generic_item_key]

        def handle_select_all_column_toggle(uph_type_col, uph_color_col, products_in_family, select_all_key):
            is_all_selected_for_column_now = st.session_state.get(select_all_key, False)
            current_family = st.session_state.selected_family_session
            family_df = st.session_state.filtered_raw_df

            action_count = 0
            for prod_name in products_in_family:
                item_exists_df_col = family_df[
                    (family_df.get("Product Display Name", family_df.get("Item Name", "")).astype(str) == str(prod_name))
                    & (family_df["Upholstery Type"].astype(str) == str(uph_type_col))
                    & (family_df["Upholstery Color"].astype(str).fillna("N/A") == str(uph_color_col))
                ] if ("Upholstery Type" in family_df.columns and "Upholstery Color" in family_df.columns) else pd.DataFrame()

                if item_exists_df_col.empty:
                    continue

                generic_item_key_col = normalize_key(f"{current_family}_{prod_name}_{uph_type_col}_{uph_color_col}")

                if is_all_selected_for_column_now:
                    if generic_item_key_col not in st.session_state.matrix_selected_generic_items:
                        base_col = "Base Color Cleaned" if "Base Color Cleaned" in item_exists_df_col.columns else None
                        unique_base_colors_col = item_exists_df_col[base_col].dropna().astype(str).unique().tolist() if base_col else []
                        requires_base_choice = len(unique_base_colors_col) > 1
                        resolved_base_if_single = unique_base_colors_col[0] if len(unique_base_colors_col) == 1 else (pd.NA if not unique_base_colors_col else None)
                        first_match_col = item_exists_df_col.iloc[0]
                        item_data_col = {
                            "key": generic_item_key_col,
                            "family": current_family,
                            "product": prod_name,
                            "upholstery_type": uph_type_col,
                            "upholstery_color": uph_color_col,
                            "requires_base_choice": requires_base_choice,
                            "available_bases": unique_base_colors_col,
                            "resolved_base_if_single": resolved_base_if_single,
                            "item_no_if_single_base": first_match_col.get("Item No", None),
                            "article_no_if_single_base": first_match_col.get("Article No", None),
                        }
                        st.session_state.matrix_selected_generic_items[generic_item_key_col] = item_data_col
                        action_count += 1
                else:
                    if generic_item_key_col in st.session_state.matrix_selected_generic_items:
                        del st.session_state.matrix_selected_generic_items[generic_item_key_col]
                        action_count += 1
                    if generic_item_key_col in st.session_state.user_chosen_base_colors_for_items:
                        del st.session_state.user_chosen_base_colors_for_items[generic_item_key_col]

            action = "selected" if is_all_selected_for_column_now else "deselected"
            st.toast(f"Column '{uph_type_col} - {uph_color_col}' {action} ({action_count} items).", icon="✅" if is_all_selected_for_column_now else "❌")

        def handle_base_color_multiselect_change(item_key):
            ms_key = f"ms_base_{item_key}"
            chosen = st.session_state.get(ms_key, [])
            st.session_state.user_chosen_base_colors_for_items[item_key] = chosen

        def handle_family_base_color_select_all_toggle(family_name, base_color, items_in_family, cb_key):
            is_checked = st.session_state.get(cb_key, False)
            count = 0
            for item in items_in_family:
                if base_color not in item.get("available_bases", []):
                    continue
                current = st.session_state.user_chosen_base_colors_for_items.get(item["key"], [])
                if is_checked:
                    if base_color not in current:
                        st.session_state.user_chosen_base_colors_for_items[item["key"]] = current + [base_color]
                        count += 1
                else:
                    if base_color in current:
                        st.session_state.user_chosen_base_colors_for_items[item["key"]] = [b for b in current if b != base_color]
                        count += 1

            action_desc = "applied to" if is_checked else "removed from"
            st.toast(f"Base color '{base_color}' {action_desc} {count} item(s) in {family_name}.", icon="✅" if is_checked else "❌")

        # Build matrix
        if selected_family and selected_family != DEFAULT_NO_SELECTION and "Product Family" in df_for_display.columns:
            family_df = df_for_display[df_for_display["Product Family"].astype(str) == str(selected_family)].copy()

            # Determine product name field
            if "Product Display Name" in family_df.columns:
                product_name_col = "Product Display Name"
            elif "Item Name" in family_df.columns:
                product_name_col = "Item Name"
            else:
                product_name_col = None

            if product_name_col and "Upholstery Type" in family_df.columns and "Upholstery Color" in family_df.columns:
                products_in_family = sorted(family_df[product_name_col].dropna().astype(str).unique().tolist())
                upholstery_types_in_family = sorted(family_df["Upholstery Type"].dropna().astype(str).unique().tolist())

                if not products_in_family:
                    st.info(f"No products for {selected_family} for current currency/market.")
                elif not upholstery_types_in_family:
                    st.info(f"No upholstery types for {selected_family} for current currency/market.")
                else:
                    data_column_map = []
                    header_swatches = []

                    # Build column map in a stable order: by upholstery type then by upholstery color
                    for uph_type_clean in upholstery_types_in_family:
                        colors_for_type_df = (
                            family_df[family_df["Upholstery Type"].astype(str) == str(uph_type_clean)][
                                ["Upholstery Color"] + (["Image URL swatch"] if "Image URL swatch" in family_df.columns else [])
                            ]
                            .drop_duplicates()
                            .sort_values(by="Upholstery Color")
                        )
                        if colors_for_type_df.empty:
                            continue

                        for _, rowc in colors_for_type_df.iterrows():
                            color_val = safe_str(rowc.get("Upholstery Color"))
                            swatch_val = rowc.get("Image URL swatch", None) if "Image URL swatch" in colors_for_type_df.columns else None
                            data_column_map.append({"uph_type": uph_type_clean, "uph_color": color_val, "swatch": swatch_val})

                    num_data_columns = len(data_column_map)
                    if num_data_columns == 0:
                        st.info("No upholstery colors found for this family.")
                    else:
                        # Headers: Upholstery Type row
                        cols_type = st.columns([2.5] + [1] * num_data_columns)
                        cols_type[0].markdown("**Product**")
                        current_type = None
                        for i, col in enumerate(cols_type[1:], start=0):
                            m = data_column_map[i]
                            label = m["uph_type"] if m["uph_type"] != current_type else ""
                            col.caption(label)
                            current_type = m["uph_type"]

                        # Swatch row
                        cols_sw = st.columns([2.5] + [1] * num_data_columns)
                        cols_sw[0].markdown("<small>Click swatch to zoom</small>", unsafe_allow_html=True)
                        for i, col in enumerate(cols_sw[1:], start=0):
                            sw_url = data_column_map[i]["swatch"]
                            if sw_url and pd.notna(sw_url):
                                col.image(sw_url, width=30)
                            else:
                                col.markdown("<span class='swatch-placeholder'></span>", unsafe_allow_html=True)

                        # Color number row
                        cols_color = st.columns([2.5] + [1] * num_data_columns)
                        cols_color[0].write("")
                        for i, col in enumerate(cols_color[1:], start=0):
                            col.caption(f"<small>{data_column_map[i]['uph_color']}</small>", unsafe_allow_html=True)

                        # Select-all row per column
                        cols_sa = st.columns([2.5] + [1] * num_data_columns, vertical_alignment="center")
                        cols_sa[0].markdown("<div class='select-all-label'>Select All:</div>", unsafe_allow_html=True)

                        for i, col in enumerate(cols_sa[1:], start=0):
                            uph_type_for_col = data_column_map[i]["uph_type"]
                            uph_color_for_col = data_column_map[i]["uph_color"]

                            # check how many selectable items in this column
                            selectable = 0
                            all_selected = True
                            for prod_name_sa in products_in_family:
                                item_exists = family_df[
                                    (family_df[product_name_col].astype(str) == str(prod_name_sa))
                                    & (family_df["Upholstery Type"].astype(str) == str(uph_type_for_col))
                                    & (family_df["Upholstery Color"].astype(str).fillna("N/A") == str(uph_color_for_col))
                                ]
                                if item_exists.empty:
                                    continue
                                selectable += 1
                                key_chk = normalize_key(f"{selected_family}_{prod_name_sa}_{uph_type_for_col}_{uph_color_for_col}")
                                if key_chk not in st.session_state.matrix_selected_generic_items:
                                    all_selected = False

                            if selectable == 0:
                                all_selected = False

                            select_all_key = normalize_key(f"select_all_cb_{selected_family}_{uph_type_for_col}_{uph_color_for_col}")
                            with col:
                                if selectable > 0:
                                    st.checkbox(
                                        " ",
                                        value=all_selected,
                                        key=select_all_key,
                                        on_change=handle_select_all_column_toggle,
                                        args=(uph_type_for_col, uph_color_for_col, products_in_family, select_all_key),
                                        label_visibility="collapsed",
                                        help=f"Select/Deselect all for {uph_type_for_col} - {uph_color_for_col}",
                                    )
                                else:
                                    st.write("")

                        st.markdown("---")

                        # Data rows
                        for prod_name in products_in_family:
                            cols_row = st.columns([2.5] + [1] * num_data_columns, vertical_alignment="center")
                            cols_row[0].markdown(f"<div class='product-name-cell'>{prod_name}</div>", unsafe_allow_html=True)

                            for i, col in enumerate(cols_row[1:], start=0):
                                uph_type = data_column_map[i]["uph_type"]
                                uph_color = data_column_map[i]["uph_color"]

                                item_exists_df = family_df[
                                    (family_df[product_name_col].astype(str) == str(prod_name))
                                    & (family_df["Upholstery Type"].astype(str) == str(uph_type))
                                    & (family_df["Upholstery Color"].astype(str).fillna("N/A") == str(uph_color))
                                ]

                                if item_exists_df.empty:
                                    col.write("")
                                    continue

                                cb_key = normalize_key(f"cb_{selected_family}_{prod_name}_{uph_type}_{uph_color}")
                                generic_item_key = normalize_key(f"{selected_family}_{prod_name}_{uph_type}_{uph_color}")
                                is_selected = generic_item_key in st.session_state.matrix_selected_generic_items

                                col.checkbox(
                                    " ",
                                    value=is_selected,
                                    key=cb_key,
                                    on_change=handle_matrix_cb_toggle,
                                    args=(prod_name, uph_type, uph_color, cb_key),
                                    label_visibility="collapsed",
                                )
            else:
                st.info("Raw data must include Product Display Name/Item Name + Upholstery Type + Upholstery Color to show matrix.")
        else:
            if selected_family and selected_family != DEFAULT_NO_SELECTION:
                st.info(f"No data for {selected_family} with current currency/market.")
            else:
                st.info("Select product family.")

st.markdown("---")

# -----------------------------
# Step 2a: Base colors
# -----------------------------
st.subheader("Step 2a: Specify Base Colors")

items_needing_base_choice_now = [
    item for item in st.session_state.matrix_selected_generic_items.values()
    if item.get("requires_base_choice")
]

if not items_needing_base_choice_now:
    st.caption("No selected items currently require base color specification.")
else:
    items_by_family = {}
    for it in items_needing_base_choice_now:
        items_by_family.setdefault(it["family"], []).append(it)

    for fam, items_in_fam in items_by_family.items():
        st.markdown(f"#### {fam}")

        # Family-level apply base color
        unique_bases = set()
        for it in items_in_fam:
            unique_bases.update(it.get("available_bases", []))
        unique_bases = sorted([b for b in unique_bases if str(b).strip()])

        if unique_bases:
            st.markdown("<small>Apply specific base color to all applicable products in this family:</small>", unsafe_allow_html=True)
            for base in unique_bases:
                fam_base_key = normalize_key(f"fam_base_all_{fam}_{base}")

                # Determine if base is selected for all applicable
                is_all = True
                applicable = 0
                for it in items_in_fam:
                    if base not in it.get("available_bases", []):
                        continue
                    applicable += 1
                    chosen = st.session_state.user_chosen_base_colors_for_items.get(it["key"], [])
                    if base not in chosen:
                        is_all = False
                        break
                if applicable == 0:
                    is_all = False

                if applicable > 0:
                    st.checkbox(
                        f"{base}",
                        value=is_all,
                        key=fam_base_key,
                        on_change=lambda fam=fam, base=base, items=items_in_fam, k=fam_base_key:
                            handle_family_base_color_select_all_toggle(fam, base, items, k),
                    )
            st.markdown("---")

        # Item-level base selection
        for it in items_in_fam:
            item_key = it["key"]
            ms_key = f"ms_base_{item_key}"
            st.markdown(f"**{it['product']}** ({it['upholstery_type']} - {it['upholstery_color']})")
            st.multiselect(
                "Available base colors for this item:",
                options=it.get("available_bases", []),
                default=st.session_state.user_chosen_base_colors_for_items.get(item_key, []),
                key=ms_key,
                on_change=lambda k=item_key: handle_base_color_multiselect_change(k),
            )
            st.markdown("---")

st.markdown("---")

# -----------------------------
# Step 3: Review selections
# -----------------------------
st.header("Step 3: Review Selections")

_current_final_items = []
df_filtered = st.session_state.filtered_raw_df

if df_filtered is not None and not df_filtered.empty:
    for key, item in st.session_state.matrix_selected_generic_items.items():
        if not item.get("requires_base_choice"):
            # If there is exactly one base or none, use resolved row info
            _current_final_items.append({
                "description": f"{item['family']} / {item['product']} / {item['upholstery_type']} / {item['upholstery_color']}"
                               + (f" / Base: {item['resolved_base_if_single']}" if pd.notna(item.get("resolved_base_if_single")) else ""),
                "item_no": item.get("item_no_if_single_base"),
                "article_no": item.get("article_no_if_single_base"),
                "key_in_matrix": key
            })
        else:
            chosen_bases = st.session_state.user_chosen_base_colors_for_items.get(key, [])
            for bc in chosen_bases:
                # Find the actual matching row for chosen base
                mask = pd.Series([True] * len(df_filtered))
                if "Product Family" in df_filtered.columns:
                    mask &= (df_filtered["Product Family"].astype(str) == str(item["family"]))
                # Product match
                prod_col = "Product Display Name" if "Product Display Name" in df_filtered.columns else ("Item Name" if "Item Name" in df_filtered.columns else None)
                if prod_col:
                    mask &= (df_filtered[prod_col].astype(str) == str(item["product"]))
                if "Upholstery Type" in df_filtered.columns:
                    mask &= (df_filtered["Upholstery Type"].astype(str) == str(item["upholstery_type"]))
                if "Upholstery Color" in df_filtered.columns:
                    mask &= (df_filtered["Upholstery Color"].astype(str).fillna("N/A") == str(item["upholstery_color"]))
                if "Base Color Cleaned" in df_filtered.columns:
                    mask &= (df_filtered["Base Color Cleaned"].astype(str).fillna("N/A") == str(bc))

                specific = df_filtered[mask]
                if specific.empty:
                    continue
                row = specific.iloc[0]
                _current_final_items.append({
                    "description": f"{item['family']} / {item['product']} / {item['upholstery_type']} / {item['upholstery_color']} / Base: {bc}",
                    "item_no": row.get("Item No", None),
                    "article_no": row.get("Article No", None),
                    "key_in_matrix": key,
                    "chosen_base": bc,
                })

# De-duplicate review list
temp_final_list_review, seen = [], set()
for it in _current_final_items:
    ukey = f"{safe_str(it.get('item_no'))}_{safe_str(it.get('chosen_base', 'NO_BASE'))}"
    if ukey not in seen and safe_str(it.get("item_no")).strip():
        temp_final_list_review.append(it)
        seen.add(ukey)

st.session_state.final_items_for_download = temp_final_list_review

if st.session_state.final_items_for_download:
    for i, combo in enumerate(st.session_state.final_items_for_download):
        col1, col2 = st.columns([10, 1])
        col1.write(f"{i+1}. {combo['description']} (Item: {combo['item_no']})")

        remove_key = normalize_key(f"final_review_remove_{i}_{combo.get('item_no')}_{combo.get('chosen_base','nobase')}")
        if col2.button("Remove", key=remove_key):
            original_key = combo.get("key_in_matrix")
            if original_key in st.session_state.matrix_selected_generic_items:
                details = st.session_state.matrix_selected_generic_items[original_key]
                if details.get("requires_base_choice") and "chosen_base" in combo:
                    bc = combo["chosen_base"]
                    chosen_list = st.session_state.user_chosen_base_colors_for_items.get(original_key, [])
                    if bc in chosen_list:
                        chosen_list.remove(bc)
                    if not chosen_list:
                        st.session_state.user_chosen_base_colors_for_items.pop(original_key, None)
                        st.session_state.matrix_selected_generic_items.pop(original_key, None)
                    else:
                        st.session_state.user_chosen_base_colors_for_items[original_key] = chosen_list
                else:
                    st.session_state.matrix_selected_generic_items.pop(original_key, None)
                    st.session_state.user_chosen_base_colors_for_items.pop(original_key, None)

            st.toast(f"Removed: {combo['description']}", icon="🗑️")
            st.rerun()

    st.markdown("---")
else:
    st.info("No items selected for download yet.")

st.markdown("---")

# -----------------------------
# Step 4: Output (NO PRICES)
# -----------------------------
st.header("Step 4: Generate Master Data File")

def build_output_dataframe() -> pd.DataFrame:
    if not st.session_state.final_items_for_download:
        return pd.DataFrame(columns=st.session_state.template_cols)

    df = st.session_state.filtered_raw_df
    out_rows = []

    # Columns order from template, minus price cols already removed
    cols = st.session_state.template_cols or remove_price_columns(list(df.columns))

    for combo in st.session_state.final_items_for_download:
        item_no = combo.get("item_no")
        if not item_no or pd.isna(item_no):
            continue

        # Find the raw row by Item No
        if "Item No" in df.columns:
            matches = df[df["Item No"].astype(str) == str(item_no)]
        else:
            matches = pd.DataFrame()

        if matches.empty:
            continue

        row = matches.iloc[0]
        out = {}
        for c in cols:
            # Special mapping: template "Product" -> raw "Item Name" if present, else Product Display Name
            if str(c).strip().lower() == "product":
                if "Item Name" in row.index and pd.notna(row.get("Item Name")):
                    out[c] = row.get("Item Name")
                elif "Product Display Name" in row.index and pd.notna(row.get("Product Display Name")):
                    out[c] = row.get("Product Display Name")
                else:
                    out[c] = None
            else:
                out[c] = row.get(c, None) if c in row.index else None

        out_rows.append(out)

    return pd.DataFrame(out_rows, columns=cols)

can_download_now = bool(st.session_state.final_items_for_download and st.session_state.selected_currency_session)

if can_download_now:
    try:
        output_df = build_output_dataframe()
        file_bytes = prepare_excel_bytes(output_df, MASTERDATA_TEMPLATE_XLSX_PATH)
        filename = f"masterdata_output_{normalize_key(st.session_state.selected_currency_session)}.xlsx"

        st.download_button(
            label="Generate and Download Master Data File (NO PRICES)",
            data=file_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="final_download_button_no_prices_v1",
            help="Click to download. Output excludes Wholesale/Retail prices.",
        )
    except Exception as e:
        st.error(f"Could not generate Excel: {e}")
else:
    help_msg = "Select currency (Step 1) and add items (Step 2 & 3)."
    if not st.session_state.selected_currency_session:
        help_msg = "Select currency first."
    elif not st.session_state.final_items_for_download:
        help_msg = "Select items first."
    st.button("Generate Master Data File", disabled=True, help=help_msg)
