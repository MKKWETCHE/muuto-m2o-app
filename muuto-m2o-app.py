import streamlit as st
import pandas as pd
import io
import os

# --- Page Configuration (MUST BE THE FIRST STREAMLIT COMMAND) ---
st.set_page_config(
    layout="wide",
    page_title="Muuto M2O",
    page_icon="favicon.png"
)

# --- Configuration & Constants ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

RAW_DATA_XLSX_PATH = os.path.join(BASE_DIR, "raw-data.xlsx")
MASTERDATA_TEMPLATE_XLSX_PATH = os.path.join(BASE_DIR, "Masterdata-output-template.xlsx")
LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")

RAW_DATA_APP_SHEET = "APP"
DEFAULT_NO_SELECTION = "--- Please Select ---"

# --- Helper Function to Construct Product Display Name ---
def construct_product_display_name(row: pd.Series) -> str:
    name_parts = []
    product_type = row.get("Product Type")
    product_model = row.get("Product Model")
    sofa_direction = row.get("Sofa Direction")

    if pd.notna(product_type) and str(product_type).strip().upper() != "N/A":
        name_parts.append(str(product_type).strip())
    if pd.notna(product_model) and str(product_model).strip().upper() != "N/A":
        name_parts.append(str(product_model).strip())

    if str(product_type).strip().lower() == "sofa chaise longue":
        if pd.notna(sofa_direction) and str(sofa_direction).strip().upper() != "N/A":
            name_parts.append(str(sofa_direction).strip())

    return " - ".join(name_parts) if name_parts else "Unnamed Product"


# --- Session State Init ---
def ensure_state(key, default):
    if key not in st.session_state:
        st.session_state[key] = default


ensure_state("raw_df", None)
ensure_state("template_cols", None)
ensure_state("selected_family_session", DEFAULT_NO_SELECTION)
ensure_state("matrix_selected_generic_items", {})  # key -> item_data
ensure_state("user_chosen_base_colors_for_items", {})  # key -> [bases]
ensure_state("final_items_for_download", [])  # list of {description, item_no, article_no, key_in_matrix, chosen_base?}

# --- Logo and Title Section ---
top_col1, top_col_spacer, top_col2 = st.columns([5.5, 0.5, 1])
with top_col1:
    st.title("Muuto Made-to-Order Master Data Tool")
with top_col2:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)

st.markdown(
    """
This tool allows you to select Made-to-Order products and generate master data.

**Pricing and currency selection have been removed.** The tool is kept minimal until BC goes live.
"""
)

# --- Load Raw Data ---
files_loaded_successfully = True

if st.session_state.raw_df is None:
    if not os.path.exists(RAW_DATA_XLSX_PATH):
        st.error(f"Raw Data file not found: {RAW_DATA_XLSX_PATH}")
        files_loaded_successfully = False
    else:
        try:
            df = pd.read_excel(RAW_DATA_XLSX_PATH, sheet_name=RAW_DATA_APP_SHEET)

            required_cols = [
                "Product Type",
                "Product Model",
                "Sofa Direction",
                "Base Color",
                "Product Family",
                "Item No",
                "Article No",
                "Image URL swatch",
                "Upholstery Type",
                "Upholstery Color",
                "Item Name",
            ]
            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                st.error(
                    f"Required columns missing in '{os.path.basename(RAW_DATA_XLSX_PATH)}': {', '.join(missing)}."
                )
                files_loaded_successfully = False
            else:
                # Always construct these helper columns
                if "Product Display Name" not in df.columns:
                    df["Product Display Name"] = df.apply(construct_product_display_name, axis=1)

                df["Base Color Cleaned"] = df["Base Color"].astype(str).str.strip().replace("N/A", pd.NA)
                df["Upholstery Type"] = df["Upholstery Type"].astype(str).str.strip()
                df["Upholstery Color"] = df["Upholstery Color"].astype(str).str.strip()

                st.session_state.raw_df = df

        except Exception as e:
            st.error(f"Error loading Raw Data: {e}")
            files_loaded_successfully = False

# --- Load Template Columns (for output) ---
if files_loaded_successfully and st.session_state.template_cols is None:
    if not os.path.exists(MASTERDATA_TEMPLATE_XLSX_PATH):
        st.error(f"Template file not found: {MASTERDATA_TEMPLATE_XLSX_PATH}")
        files_loaded_successfully = False
    else:
        try:
            st.session_state.template_cols = pd.read_excel(MASTERDATA_TEMPLATE_XLSX_PATH).columns.tolist()
        except Exception as e:
            st.error(f"Error loading Template: {e}")
            files_loaded_successfully = False

# --- Callbacks ---
def _normalize_key(s: str) -> str:
    return (
        str(s)
        .replace(" ", "_")
        .replace("/", "_")
        .replace("(", "")
        .replace(")", "")
    )

def handle_matrix_cb_toggle(prod_name, uph_type, uph_color, checkbox_key_matrix):
    is_checked = st.session_state.get(checkbox_key_matrix, False)
    selected_family = st.session_state.selected_family_session
    generic_item_key = _normalize_key(f"{selected_family}_{prod_name}_{uph_type}_{uph_color}")

    df = st.session_state.raw_df
    if df is None:
        return

    if is_checked:
        matching = df[
            (df["Product Family"] == selected_family)
            & (df["Product Display Name"] == prod_name)
            & (df["Upholstery Type"].fillna("N/A") == uph_type)
            & (df["Upholstery Color"].astype(str).fillna("N/A") == uph_color)
        ]

        if not matching.empty:
            unique_bases = matching["Base Color Cleaned"].dropna().unique().tolist()
            first_row = matching.iloc[0]

            item_data = {
                "key": generic_item_key,
                "family": selected_family,
                "product": prod_name,
                "upholstery_type": uph_type,
                "upholstery_color": uph_color,
                "requires_base_choice": len(unique_bases) > 1,
                "available_bases": unique_bases if len(unique_bases) > 1 else [],
                "item_no_if_single_base": first_row["Item No"] if len(unique_bases) <= 1 else None,
                "article_no_if_single_base": first_row["Article No"] if len(unique_bases) <= 1 else None,
                "resolved_base_if_single": unique_bases[0] if len(unique_bases) == 1 else (pd.NA if not unique_bases else None),
            }
            st.session_state.matrix_selected_generic_items[generic_item_key] = item_data
    else:
        if generic_item_key in st.session_state.matrix_selected_generic_items:
            del st.session_state.matrix_selected_generic_items[generic_item_key]
        if generic_item_key in st.session_state.user_chosen_base_colors_for_items:
            del st.session_state.user_chosen_base_colors_for_items[generic_item_key]


def handle_select_all_column_toggle(uph_type_col, uph_color_col, products_in_col, select_all_key):
    is_all_selected = st.session_state.get(select_all_key, False)
    selected_family = st.session_state.selected_family_session
    df = st.session_state.raw_df
    if df is None:
        return

    family_df = df[df["Product Family"] == selected_family]

    for prod_name in products_in_col:
        exists_df = family_df[
            (family_df["Product Display Name"] == prod_name)
            & (family_df["Upholstery Type"].fillna("N/A") == uph_type_col)
            & (family_df["Upholstery Color"].astype(str).fillna("N/A") == uph_color_col)
        ]

        if exists_df.empty:
            continue

        generic_key = _normalize_key(f"{selected_family}_{prod_name}_{uph_type_col}_{uph_color_col}")

        if is_all_selected:
            if generic_key not in st.session_state.matrix_selected_generic_items:
                unique_bases = exists_df["Base Color Cleaned"].dropna().unique().tolist()
                first_row = exists_df.iloc[0]
                st.session_state.matrix_selected_generic_items[generic_key] = {
                    "key": generic_key,
                    "family": selected_family,
                    "product": prod_name,
                    "upholstery_type": uph_type_col,
                    "upholstery_color": uph_color_col,
                    "requires_base_choice": len(unique_bases) > 1,
                    "available_bases": unique_bases if len(unique_bases) > 1 else [],
                    "item_no_if_single_base": first_row["Item No"] if len(unique_bases) <= 1 else None,
                    "article_no_if_single_base": first_row["Article No"] if len(unique_bases) <= 1 else None,
                    "resolved_base_if_single": unique_bases[0] if len(unique_bases) == 1 else (pd.NA if not unique_bases else None),
                }
        else:
            if generic_key in st.session_state.matrix_selected_generic_items:
                del st.session_state.matrix_selected_generic_items[generic_key]
            if generic_key in st.session_state.user_chosen_base_colors_for_items:
                del st.session_state.user_chosen_base_colors_for_items[generic_key]

    action = "selected" if is_all_selected else "deselected"
    st.toast(f"All available items in column '{uph_type_col} - {uph_color_col}' {action}.",
             icon="✅" if is_all_selected else "❌")


def handle_base_color_multiselect_change(item_key_for_base_select):
    multiselect_key = f"ms_base_{item_key_for_base_select}"
    st.session_state.user_chosen_base_colors_for_items[item_key_for_base_select] = st.session_state.get(multiselect_key, [])


def handle_family_base_color_select_all_toggle(family_name_cb, base_color_cb, items_in_family_cb, checkbox_key_cb):
    is_checked = st.session_state.get(checkbox_key_cb, False)
    action_count = 0

    for item in items_in_family_cb:
        item_key = item["key"]
        if base_color_cb not in item["available_bases"]:
            continue

        current_bases = st.session_state.user_chosen_base_colors_for_items.get(item_key, [])

        if is_checked:
            if base_color_cb not in current_bases:
                st.session_state.user_chosen_base_colors_for_items[item_key] = current_bases + [base_color_cb]
                action_count += 1
        else:
            if base_color_cb in current_bases:
                st.session_state.user_chosen_base_colors_for_items[item_key] = [b for b in current_bases if b != base_color_cb]
                action_count += 1

    if action_count > 0:
        action_desc = "applied to" if is_checked else "removed from"
        st.toast(
            f"Base color '{base_color_cb}' {action_desc} {action_count} applicable product(s) in {family_name_cb}.",
            icon="✅" if is_checked else "❌",
        )

# --- Main UI ---
if not files_loaded_successfully:
    st.stop()

df = st.session_state.raw_df
if df is None or df.empty:
    st.error("Raw data is empty.")
    st.stop()

# --- Step 2: Select Product Combinations ---
st.header("Step 2: Select Product Combinations (Product / Upholstery / Color)")

families = [DEFAULT_NO_SELECTION] + sorted(df["Product Family"].dropna().unique().tolist())
if st.session_state.selected_family_session not in families:
    st.session_state.selected_family_session = DEFAULT_NO_SELECTION

selected_family = st.selectbox(
    "Select Product Family:",
    options=families,
    index=families.index(st.session_state.selected_family_session),
    key="family_selector_main",
)
st.session_state.selected_family_session = selected_family

if not selected_family or selected_family == DEFAULT_NO_SELECTION:
    st.info("Select a product family.")
else:
    family_df = df[df["Product Family"] == selected_family].copy()
    if family_df.empty:
        st.info(f"No data for {selected_family}.")
    else:
        # Ensure helper columns exist (failsafe)
        if "Product Display Name" not in family_df.columns:
            family_df["Product Display Name"] = family_df.apply(construct_product_display_name, axis=1)

        products_in_family = sorted(family_df["Product Display Name"].dropna().unique().tolist())
        upholstery_types = sorted(family_df["Upholstery Type"].dropna().unique().tolist())

        if not products_in_family:
            st.info(f"No products in {selected_family}.")
        elif not upholstery_types:
            st.info(f"No upholstery types in {selected_family}.")
        else:
            data_column_map = []

            for uph_type in upholstery_types:
                colors_df = (
                    family_df[family_df["Upholstery Type"] == uph_type][["Upholstery Color", "Image URL swatch"]]
                    .drop_duplicates()
                    .sort_values(by="Upholstery Color")
                )
                for _, r in colors_df.iterrows():
                    col_color = str(r["Upholstery Color"])
                    swatch = r["Image URL swatch"] if pd.notna(r["Image URL swatch"]) else None
                    data_column_map.append({"uph_type": uph_type, "uph_color": col_color, "swatch": swatch})

            num_cols = len(data_column_map)

            if num_cols == 0:
                st.info("No upholstery colors found.")
            else:
                # Upholstery Type header row
                cols_type = st.columns([2.5] + [1] * num_cols)
                current = None
                for i in range(1, num_cols + 1):
                    entry = data_column_map[i - 1]
                    with cols_type[i]:
                        if entry["uph_type"] != current:
                            st.caption(f"<div class='upholstery-header'>{entry['uph_type']}</div>", unsafe_allow_html=True)
                            current = entry["uph_type"]

                # Swatch row
                cols_sw = st.columns([2.5] + [1] * num_cols)
                cols_sw[0].markdown("<div class='zoom-instruction'><br>Click swatch to zoom</div>", unsafe_allow_html=True)
                for i in range(1, num_cols + 1):
                    sw_url = data_column_map[i - 1]["swatch"]
                    with cols_sw[i]:
                        if sw_url:
                            st.image(sw_url, width=30)
                        else:
                            st.markdown("<div class='swatch-placeholder'></div>", unsafe_allow_html=True)

                # Color number row
                cols_color = st.columns([2.5] + [1] * num_cols)
                for i in range(1, num_cols + 1):
                    with cols_color[i]:
                        st.caption(f"<small>{data_column_map[i - 1]['uph_color']}</small>", unsafe_allow_html=True)

                # Select All row
                cols_sa = st.columns([2.5] + [1] * num_cols, vertical_alignment="center")
                cols_sa[0].markdown("<div class='select-all-label'>Select All:</div>", unsafe_allow_html=True)

                for i in range(1, num_cols + 1):
                    entry = data_column_map[i - 1]
                    uph_type = entry["uph_type"]
                    uph_color = entry["uph_color"]

                    # Determine if all selectable in that column are selected
                    all_selected = True
                    selectable_count = 0

                    for prod in products_in_family:
                        exists_df = family_df[
                            (family_df["Product Display Name"] == prod)
                            & (family_df["Upholstery Type"].fillna("N/A") == uph_type)
                            & (family_df["Upholstery Color"].astype(str).fillna("N/A") == uph_color)
                        ]
                        if not exists_df.empty:
                            selectable_count += 1
                            gen_key = _normalize_key(f"{selected_family}_{prod}_{uph_type}_{uph_color}")
                            if gen_key not in st.session_state.matrix_selected_generic_items:
                                all_selected = False
                                break

                    if selectable_count == 0:
                        all_selected = False

                    select_all_key = _normalize_key(f"select_all_cb_{selected_family}_{uph_type}_{uph_color}")
                    with cols_sa[i]:
                        if selectable_count > 0:
                            st.checkbox(
                                " ",
                                value=all_selected,
                                key=select_all_key,
                                on_change=handle_select_all_column_toggle,
                                args=(uph_type, uph_color, products_in_family, select_all_key),
                                label_visibility="collapsed",
                                help=f"Select/Deselect all for {uph_type} - {uph_color}",
                            )
                        else:
                            st.markdown("<div class='checkbox-placeholder'></div>", unsafe_allow_html=True)

                st.markdown("---")

                # Product rows
                for prod in products_in_family:
                    cols_row = st.columns([2.5] + [1] * num_cols, vertical_alignment="center")
                    cols_row[0].markdown(f"<div class='product-name-cell'>{prod}</div>", unsafe_allow_html=True)

                    for i in range(1, num_cols + 1):
                        entry = data_column_map[i - 1]
                        uph_type = entry["uph_type"]
                        uph_color = entry["uph_color"]

                        exists_df = family_df[
                            (family_df["Product Display Name"] == prod)
                            & (family_df["Upholstery Type"].fillna("N/A") == uph_type)
                            & (family_df["Upholstery Color"].astype(str).fillna("N/A") == uph_color)
                        ]

                        with cols_row[i]:
                            if not exists_df.empty:
                                cb_key = _normalize_key(f"cb_{selected_family}_{prod}_{uph_type}_{uph_color}")
                                gen_key = _normalize_key(f"{selected_family}_{prod}_{uph_type}_{uph_color}")
                                is_selected = gen_key in st.session_state.matrix_selected_generic_items

                                st.checkbox(
                                    " ",
                                    value=is_selected,
                                    key=cb_key,
                                    on_change=handle_matrix_cb_toggle,
                                    args=(prod, uph_type, uph_color, cb_key),
                                    label_visibility="collapsed",
                                )
                            else:
                                st.markdown("<div class='checkbox-placeholder'></div>", unsafe_allow_html=True)

# --- Step 2a: Base Colors ---
items_needing_base_choice = [
    item for item in st.session_state.matrix_selected_generic_items.values()
    if item.get("requires_base_choice")
]

if items_needing_base_choice:
    st.subheader("Step 2a: Specify Base Colors")

    items_by_family = {}
    for item in items_needing_base_choice:
        items_by_family.setdefault(item["family"], []).append(item)

    for fam_name, items_in_fam in items_by_family.items():
        st.markdown(f"#### {fam_name}")

        unique_bases = set()
        for it in items_in_fam:
            unique_bases.update(it.get("available_bases", []))

        sorted_unique_bases = sorted(list(unique_bases))
        if sorted_unique_bases:
            st.markdown("<small>Apply specific base color to all applicable products in this family:</small>", unsafe_allow_html=True)
            for base_color in sorted_unique_bases:
                key_cb = _normalize_key(f"fam_base_all_{fam_name}_{base_color}")

                # Determine default checked state
                is_all = True
                applicable = 0
                for it in items_in_fam:
                    if base_color in it.get("available_bases", []):
                        applicable += 1
                        chosen = st.session_state.user_chosen_base_colors_for_items.get(it["key"], [])
                        if base_color not in chosen:
                            is_all = False
                            break
                if applicable == 0:
                    is_all = False

                if applicable > 0:
                    st.checkbox(
                        f"{base_color}",
                        value=is_all,
                        key=key_cb,
                        on_change=handle_family_base_color_select_all_toggle,
                        args=(fam_name, base_color, items_in_fam, key_cb),
                    )

            st.markdown("---")

        for it in items_in_fam:
            item_key = it["key"]
            ms_key = f"ms_base_{item_key}"
            st.markdown(f"**{it['product']}** ({it['upholstery_type']} - {it['upholstery_color']})")

            st.multiselect(
                label="Available base colors for this item:",
                options=it.get("available_bases", []),
                default=st.session_state.user_chosen_base_colors_for_items.get(item_key, []),
                key=ms_key,
                on_change=handle_base_color_multiselect_change,
                args=(item_key,),
            )
            st.markdown("---")

# --- Step 3: Review Selections ---
st.header("Step 3: Review Selections")

_current_final_items = []
df_all = st.session_state.raw_df

for key, item in st.session_state.matrix_selected_generic_items.items():
    if not item.get("requires_base_choice"):
        if item.get("item_no_if_single_base") is not None:
            desc = f"{item['family']} / {item['product']} / {item['upholstery_type']} / {item['upholstery_color']}"
            if pd.notna(item.get("resolved_base_if_single", pd.NA)):
                desc += f" / Base: {item['resolved_base_if_single']}"
            _current_final_items.append({
                "description": desc,
                "item_no": item["item_no_if_single_base"],
                "article_no": item["article_no_if_single_base"],
                "key_in_matrix": key,
            })
    else:
        chosen_bases = st.session_state.user_chosen_base_colors_for_items.get(key, [])
        for bc in chosen_bases:
            specific_df = df_all[
                (df_all["Product Family"] == item["family"])
                & (df_all["Product Display Name"] == item["product"])
                & (df_all["Upholstery Type"].fillna("N/A") == item["upholstery_type"])
                & (df_all["Upholstery Color"].astype(str).fillna("N/A") == item["upholstery_color"])
                & (df_all["Base Color Cleaned"].fillna("N/A") == bc)
            ]
            if not specific_df.empty:
                row = specific_df.iloc[0]
                _current_final_items.append({
                    "description": f"{item['family']} / {item['product']} / {item['upholstery_type']} / {item['upholstery_color']} / Base: {bc}",
                    "item_no": row["Item No"],
                    "article_no": row["Article No"],
                    "key_in_matrix": key,
                    "chosen_base": bc,
                })

# Remove duplicates
final_list = []
seen = set()
for it in _current_final_items:
    ukey = f"{it['item_no']}_{it.get('chosen_base', 'NO_BASE')}"
    if ukey not in seen:
        seen.add(ukey)
        final_list.append(it)

st.session_state.final_items_for_download = final_list

if st.session_state.final_items_for_download:
    st.markdown("**Current Selections for Download:**")
    for i in range(len(st.session_state.final_items_for_download) - 1, -1, -1):
        combo = st.session_state.final_items_for_download[i]
        c1, c2 = st.columns([0.9, 0.1])
        c1.write(f"{i+1}. {combo['description']} (Item: {combo['item_no']})")

        remove_key = _normalize_key(f"final_review_remove_{i}_{combo['item_no']}_{combo.get('chosen_base','nobase')}")
        if c2.button("Remove", key=remove_key):
            original_key = combo["key_in_matrix"]

            if original_key in st.session_state.matrix_selected_generic_items:
                details = st.session_state.matrix_selected_generic_items[original_key]
                if details.get("requires_base_choice") and "chosen_base" in combo:
                    bc = combo["chosen_base"]
                    if original_key in st.session_state.user_chosen_base_colors_for_items:
                        if bc in st.session_state.user_chosen_base_colors_for_items[original_key]:
                            st.session_state.user_chosen_base_colors_for_items[original_key].remove(bc)
                            if not st.session_state.user_chosen_base_colors_for_items[original_key]:
                                del st.session_state.user_chosen_base_colors_for_items[original_key]
                                if not st.session_state.user_chosen_base_colors_for_items.get(original_key):
                                    del st.session_state.matrix_selected_generic_items[original_key]
                else:
                    del st.session_state.matrix_selected_generic_items[original_key]
                    if original_key in st.session_state.user_chosen_base_colors_for_items:
                        del st.session_state.user_chosen_base_colors_for_items[original_key]

            st.toast(f"Removed: {combo['description']}", icon="🗑️")
            st.rerun()
else:
    st.info("No items selected for download yet.")

# --- Step 4: Generate Master Data File ---
st.header("Step 4: Generate Master Data File")

def prepare_excel_for_download_final():
    if not st.session_state.final_items_for_download:
        st.warning("No items selected.")
        return None

    output_data = []
    template_cols = st.session_state.template_cols or []

    # Output columns = template columns (NO price columns)
    final_output_cols = list(template_cols)

    for combo in st.session_state.final_items_for_download:
        item_no = combo["item_no"]
        item_df = df_all[df_all["Item No"] == item_no]
        if item_df.empty:
            st.warning(f"Item No {item_no} not found. Skipping.")
            continue

        item_series = item_df.iloc[0]
        out_row = {}

        for col_name in final_output_cols:
            value_to_assign = None

            if col_name.strip().lower() == "product":
                # Use Item Name primarily
                if "Item Name" in item_series.index:
                    value_to_assign = item_series["Item Name"]
                else:
                    value_to_assign = item_series.get("Product Display Name", None)
            elif col_name in item_series.index:
                value_to_assign = item_series[col_name]

            out_row[col_name] = value_to_assign

        output_data.append(out_row)

    if not output_data:
        st.info("No data to output.")
        return None

    out_df = pd.DataFrame(output_data, columns=final_output_cols)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Masterdata Output")

    return buffer.getvalue()

if st.session_state.final_items_for_download:
    file_bytes = prepare_excel_for_download_final()
    if file_bytes:
        st.download_button(
            label="Generate and Download Master Data File",
            data=file_bytes,
            file_name="masterdata_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="final_download_button_no_prices",
            help="Click to download.",
        )
else:
    st.button("Generate Master Data File", disabled=True, help="Select items first.")

# --- Styling ---
st.markdown(
    """
<style>
    .stApp, body { background-color: #EFEEEB !important; }
    .main .block-container { background-color: #EFEEEB !important; padding-top: 2rem; }

    h1, h2, h3 { text-transform: none !important; }
    h1 { color: #333; }
    h2 { color: #1E40AF; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; }
    h3 { color: #1E40AF; font-size: 1.25em; padding-bottom: 3px; margin-top: 20px; margin-bottom: 10px; }
    h4 { color: #102A63; font-size: 1.1em; margin-top: 15px; margin-bottom: 5px; }

    div[data-testid="stCaptionContainer"] > div > p {
        font-weight: bold;
        font-size: 0.8em !important;
        color: #31333F !important;
        text-align: center;
        white-space: normal;
        overflow-wrap:break-word;
        line-height: 1.2;
        padding: 2px;
    }
    .upholstery-header {
        white-space: normal !important;
        overflow: visible !important;
        text-overflow: clip !important;
        display: block;
        max-width: 100%;
        line-height: 1.2;
        color: #31333F !important;
        text-transform: capitalize !important;
        font-weight: bold !important;
        font-size: 0.8em !important;
    }
    div[data-testid="stCaptionContainer"] small { color: #31333F !important; font-weight: normal !important; font-size: 0.75em !important; }
    div[data-testid="stCaptionContainer"] img { max-height: 25px !important; width: 25px !important; object-fit: cover !important; margin-right:2px; }
    .swatch-placeholder { width:25px !important; height:25px !important; display: flex; align-items: center; justify-content: center; font-size: 0.6em; color: #ccc; border: 1px dashed #ddd; background-color: #f9f9f9; }
    .zoom-instruction { font-size: 0.6em; color: #555; text-align: left; padding-top: 10px; }

    .select-all-label {
        font-size: 0.75em;
        color: #31333F;
        text-align: right;
        padding-right: 5px;
        font-weight:bold;
        display: flex;
        align-items: center;
        height: 100%;
    }
    .checkbox-placeholder { width: 20px; height: 20px; margin: auto; }

    div[data-testid="stImage"], div[data-testid="stImage"] img { border-radius: 0 !important; overflow: visible !important; }

    .product-name-cell {
        display: flex;
        align-items: center;
        height: auto;
        min-height: 30px;
        line-height: 1.3;
        max-height: calc(1.3em * 2 + 4px);
        overflow-y: hidden;
        color: #31333F !important;
        font-weight: normal !important;
        font-size: 0.8em !important;
        padding-right: 5px;
        word-break: break-word;
        box-sizing: border-box;
    }

    div[data-testid="stHorizontalBlock"] > div[data-testid="stVerticalBlock"] {
        height: 30px !important;
        min-height: 30px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        padding: 0 !important;
        margin: 0 !important;
        box-sizing: border-box;
    }

    div[data-testid="stCheckbox"] {
        width: auto !important;
        min-height: 28px;
        display: flex;
        align-items: center;
    }

    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] {
        display: flex !important;
        align-items: center !important;
        width: auto !important;
        height: auto !important;
        padding: 0 !important;
        margin: 0 !important;
        cursor: pointer;
    }

    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] > span:first-child {
        background-color: #FFFFFF !important;
        border: 1px solid #5B4A14 !important;
        box-shadow: none !important;
        width: 20px !important;
        height: 20px !important;
        min-width: 20px !important;
        min-height: 20px !important;
        border-radius: 0.25rem !important;
        margin-right: 0.5rem !important;
        padding: 0 !important;
        box-sizing: border-box !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        flex-shrink: 0;
    }

    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"] > span:first-child svg {
        fill: #FFFFFF !important;
        width: 12px !important;
        height: 12px !important;
    }

    div[data-testid="stCheckbox"] > label[data-baseweb="checkbox"]:has(input[type="checkbox"][aria-checked="true"]) > span:first-child {
        background-color: #5B4A14 !important;
        border-color: #5B4A14 !important;
    }

    hr { margin-top: 0.5rem !important; margin-bottom: 0.5rem !important; border-top: 1px solid #dee2e6; }

    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"],
    div[data-testid="stButton"] button[data-testid^="stBaseButton"] {
        border: 1px solid #5B4A14 !important;
        background-color: #FFFFFF !important;
        color: #5B4A14 !important;
        padding: 0.375rem 0.75rem !important;
        font-size: 1rem !important;
        line-height: 1.5 !important;
        border-radius: 0.25rem !important;
        font-weight: 500 !important;
        text-transform: none !important;
    }

    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:hover,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:hover {
        background-color: #5B4A14 !important;
        color: #FFFFFF !important;
        border-color: #5B4A14 !important;
    }

    div[data-testid="stAlert"] {
        background-color: #f0f2f6 !important;
        border: 1px solid #D1D5DB !important;
        border-radius: 0.25rem !important;
    }
</style>
""",
    unsafe_allow_html=True
)
