import io
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="Pinestate Liquor Template Builder",
    page_icon="🍾",
    layout="wide",
)

APP_TITLE = "Pinestate Liquor Template Builder"

BASE_DIR = Path(__file__).parent
TEMPLATE_DIR = BASE_DIR / "templates"


# ============================================================
# HELPERS
# ============================================================

def clean(value):
    """Convert values to clean strings while preserving blanks."""
    if value is None:
        return ""

    if pd.isna(value):
        return ""

    return str(value).strip()


def read_input_file(uploaded_file):
    """
    Read CSV or Excel input files.
    All values are initially treated as strings to preserve
    UPCs, vendor numbers, leading zeroes, etc.
    """

    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):
        return pd.read_csv(
            uploaded_file,
            dtype=str,
            keep_default_na=False,
        )

    elif filename.endswith((".xlsx", ".xls")):
        return pd.read_excel(
            uploaded_file,
            dtype=str,
            keep_default_na=False,
        )

    else:
        raise ValueError(
            f"Unsupported file type: {uploaded_file.name}"
        )


def source_column(df, excel_column):
    """
    Get an Excel-style column from a pandas DataFrame.

    Example:
        A = first column
        B = second column
        H = eighth column
    """

    index = ord(excel_column.upper()) - ord("A")

    if index < 0 or index >= len(df.columns):
        raise ValueError(
            f"Input file does not contain column {excel_column}."
        )

    return df.iloc[:, index].map(clean)


def load_template(filename):
    """Load an Excel template from the templates folder."""

    path = TEMPLATE_DIR / filename

    if not path.exists():
        raise FileNotFoundError(
            f"Template '{filename}' was not found in the templates folder."
        )

    return load_workbook(path)


def write_to_template(template_filename, rows):
    """
    Write generated records into the supplied Excel template.

    Headers are on row 11.
    Data starts on row 12.
    """

    wb = load_template(template_filename)

    # Use the first worksheet in the supplied template.
    ws = wb[wb.sheetnames[0]]

    # Clear existing detail rows while keeping rows 1-11 intact.
    if ws.max_row >= 12:

        for row in ws.iter_rows(
            min_row=12,
            max_row=ws.max_row,
            min_col=1,
            max_col=ws.max_column,
        ):
            for cell in row:
                cell.value = None

    # Write records beginning on row 12.
    for row_number, row_data in enumerate(rows, start=12):

        for column_number, value in enumerate(row_data, start=1):

            ws.cell(
                row=row_number,
                column=column_number,
            ).value = value

    # Return workbook as bytes.
    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output.getvalue()


def remove_duplicates(rows):
    """
    Remove duplicate output records while preserving order.
    """

    seen = set()
    unique_rows = []

    for row in rows:

        key = tuple(row)

        if key not in seen:
            seen.add(key)
            unique_rows.append(row)

    return unique_rows


# ============================================================
# PRODUCTS LOOKUP
# ============================================================

def build_products_lookup(products):
    """
    Products File:

    Column D = lookup key
    Column H = value returned to output Column B
    """

    lookup = {}

    products_key = source_column(products, "D")
    products_h = source_column(products, "H")

    for key, value in zip(products_key, products_h):

        key = clean(key)

        if key and key not in lookup:
            lookup[key] = clean(value)

    return lookup


# ============================================================
# EG PROMO RETAIL
# ============================================================

def build_eg_promo_retail(promo_retail):
    """
    EG Promo Retail mapping:

    Output C = Promo Retail B
    Output F = Promo Retail C
    Output H = Promo Retail E
    Output K = Promo Retail F
    Output L = Promo Retail I
    Output N = 08784
    Output O = Pine State Liquor EG
    Output P = 0
    """

    rows = []

    for _, record in promo_retail.iterrows():

        # Output has columns A:P = 16 columns.
        output = [""] * 16

        # C <- Source B
        output[2] = clean(record.iloc[1])

        # F <- Source C
        output[5] = clean(record.iloc[2])

        # H <- Source E
        output[7] = clean(record.iloc[4])

        # K <- Source F
        output[10] = clean(record.iloc[5])

        # L <- Source I
        output[11] = clean(record.iloc[8])

        # N <- Default
        output[13] = "08784"

        # O <- Default
        output[14] = "Pine State Liquor EG"

        # P <- Default
        output[15] = "0"

        rows.append(output)

    # Remove duplicates.
    return remove_duplicates(rows)


# ============================================================
# EG STANDARD COST
# ============================================================

def build_eg_standard_cost(raw_cost, products):
    """
    Only include Raw Vendor Store Cost records where:

        Column L = 0

    Output:

    A = VC
    B = Products H
    C = Raw C
    D = Raw O
    E = blank
    F = Raw K
    G = Raw M
    H = blank
    I = Raw B
    J = blank
    K = blank
    L = blank

    Column B lookup:

        Output I -> Products D -> Products H

    If there is no match, Column B is blank.
    """

    products_lookup = build_products_lookup(products)

    rows = []

    for _, record in raw_cost.iterrows():

        # Raw Column L = Excel column L = position 11
        promo_flag = clean(record.iloc[11])

        # Standard Cost only
        if promo_flag != "0":
            continue

        output = [""] * 12

        # A
        output[0] = "VC"

        # C <- Raw Column C
        output[2] = clean(record.iloc[2])

        # D <- Raw Column O
        output[3] = clean(record.iloc[14])

        # F <- Raw Column K
        output[5] = clean(record.iloc[10])

        # G <- Raw Column M
        output[6] = clean(record.iloc[12])

        # I <- Raw Column B
        output[8] = clean(record.iloc[1])

        # ----------------------------------------------------
        # B is populated LAST.
        #
        # Output I is used as the lookup key against
        # Products Column D.
        # ----------------------------------------------------

        lookup_key = output[8]

        if lookup_key:
            output[1] = products_lookup.get(
                lookup_key,
                "",
            )

        rows.append(output)

    # Remove duplicates.
    return remove_duplicates(rows)


# ============================================================
# EG PROMO COST
# ============================================================

def build_eg_promo_cost(raw_cost, products):
    """
    Only include Raw Vendor Store Cost records where:

        Column L = 1

    Output:

    A = VC
    B = Products H
    C = Raw C
    D = Raw O
    E = Raw L
    F = Raw K
    G = Raw M
    H = Raw N
    I = Raw B
    J = blank
    K = blank
    L = blank

    Column B lookup:

        Output I -> Products D -> Products H

    If there is no match, Column B is blank.
    """

    products_lookup = build_products_lookup(products)

    rows = []

    for _, record in raw_cost.iterrows():

        # Raw Column L
        promo_flag = clean(record.iloc[11])

        # Promo Cost only
        if promo_flag != "1":
            continue

        output = [""] * 12

        # A
        output[0] = "VC"

        # C <- Raw Column C
        output[2] = clean(record.iloc[2])

        # D <- Raw Column O
        output[3] = clean(record.iloc[14])

        # E <- Raw Column L
        output[4] = clean(record.iloc[11])

        # F <- Raw Column K
        output[5] = clean(record.iloc[10])

        # G <- Raw Column M
        output[6] = clean(record.iloc[12])

        # H <- Raw Column N
        output[7] = clean(record.iloc[13])

        # I <- Raw Column B
        output[8] = clean(record.iloc[1])

        # ----------------------------------------------------
        # B is populated using Output I as lookup key.
        # ----------------------------------------------------

        lookup_key = output[8]

        if lookup_key:
            output[1] = products_lookup.get(
                lookup_key,
                "",
            )

        rows.append(output)

    # Remove duplicates.
    return remove_duplicates(rows)


# ============================================================
# STREAMLIT UI
# ============================================================

st.title(APP_TITLE)

st.markdown(
    """
    Upload the three required source files below and generate
    the three Pine State Liquor EG templates.
    """
)


# ------------------------------------------------------------
# FILE UPLOADS
# ------------------------------------------------------------

st.subheader("Required Files")

col1, col2, col3 = st.columns(3)

with col1:

    products_file = st.file_uploader(
        "1. Products File",
        type=["csv", "xlsx", "xls"],
        key="products_file",
    )

with col2:

    promo_retail_file = st.file_uploader(
        "2. Promo Retail File",
        type=["csv", "xlsx", "xls"],
        key="promo_retail_file",
    )

with col3:

    raw_cost_file = st.file_uploader(
        "3. Raw Vendor Store Cost File",
        type=["csv", "xlsx", "xls"],
        key="raw_cost_file",
    )


st.divider()


# ------------------------------------------------------------
# PROCESS BUTTON
# ------------------------------------------------------------

process = st.button(
    "🚀 Process Files",
    type="primary",
    use_container_width=True,
)


if process:

    # --------------------------------------------------------
    # Validate uploads
    # --------------------------------------------------------

    if not products_file:
        st.error("Please upload the Products File.")
        st.stop()

    if not promo_retail_file:
        st.error("Please upload the Promo Retail File.")
        st.stop()

    if not raw_cost_file:
        st.error("Please upload the Raw Vendor Store Cost File.")
        st.stop()


    try:

        with st.spinner("Processing files..."):

            # ------------------------------------------------
            # Read input files
            # ------------------------------------------------

            products = read_input_file(products_file)

            promo_retail = read_input_file(
                promo_retail_file
            )

            raw_cost = read_input_file(
                raw_cost_file
            )


            # ------------------------------------------------
            # Validate source columns
            # ------------------------------------------------

            if len(products.columns) < 8:
                raise ValueError(
                    "Products File must contain at least 8 columns."
                )

            if len(promo_retail.columns) < 9:
                raise ValueError(
                    "Promo Retail File must contain at least 9 columns."
                )

            if len(raw_cost.columns) < 15:
                raise ValueError(
                    "Raw Vendor Store Cost File must contain at least 15 columns."
                )


            # ------------------------------------------------
            # Generate outputs
            # ------------------------------------------------

            promo_retail_rows = build_eg_promo_retail(
                promo_retail
            )

            standard_cost_rows = build_eg_standard_cost(
                raw_cost,
                products,
            )

            promo_cost_rows = build_eg_promo_cost(
                raw_cost,
                products,
            )


            # ------------------------------------------------
            # Write to templates
            # ------------------------------------------------

            promo_retail_output = write_to_template(
                "EG_PromoRetail.xlsx",
                promo_retail_rows,
            )

            standard_cost_output = write_to_template(
                "EG_StandardCost.xlsx",
                standard_cost_rows,
            )

            promo_cost_output = write_to_template(
                "EG_PromoCost.xlsx",
                promo_cost_rows,
            )


            # ------------------------------------------------
            # Store outputs in session state
            # ------------------------------------------------

            st.session_state["promo_retail_output"] = (
                promo_retail_output
            )

            st.session_state["standard_cost_output"] = (
                standard_cost_output
            )

            st.session_state["promo_cost_output"] = (
                promo_cost_output
            )

            st.session_state["promo_retail_count"] = (
                len(promo_retail_rows)
            )

            st.session_state["standard_cost_count"] = (
                len(standard_cost_rows)
            )

            st.session_state["promo_cost_count"] = (
                len(promo_cost_rows)
            )

            st.session_state["processed"] = True


        st.success(
            "Files processed successfully!"
        )


    except Exception as error:

        st.error(
            f"Processing failed: {error}"
        )

        st.stop()


# ============================================================
# DOWNLOAD SECTION
# ============================================================

if st.session_state.get("processed"):

    st.subheader("Generated Files")


    # --------------------------------------------------------
    # Promo Retail
    # --------------------------------------------------------

    col1, col2 = st.columns([4, 1])

    with col1:

        st.write(
            f"**EG_Promo Retail File.xlsx**  \n"
            f"{st.session_state['promo_retail_count']:,} records"
        )

    with col2:

        st.download_button(
            label="Download",
            data=st.session_state[
                "promo_retail_output"
            ],
            file_name="EG_Promo Retail File.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_promo_retail",
            use_container_width=True,
        )


    # --------------------------------------------------------
    # Standard Cost
    # --------------------------------------------------------

    col1, col2 = st.columns([4, 1])

    with col1:

        st.write(
            f"**EG_Standard Cost File.xlsx**  \n"
            f"{st.session_state['standard_cost_count']:,} records"
        )

    with col2:

        st.download_button(
            label="Download",
            data=st.session_state[
                "standard_cost_output"
            ],
            file_name="EG_Standard Cost File.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_standard_cost",
            use_container_width=True,
        )


    # --------------------------------------------------------
    # Promo Cost
    # --------------------------------------------------------

    col1, col2 = st.columns([4, 1])

    with col1:

        st.write(
            f"**EG_Promo Cost File.xlsx**  \n"
            f"{st.session_state['promo_cost_count']:,} records"
        )

    with col2:

        st.download_button(
            label="Download",
            data=st.session_state[
                "promo_cost_output"
            ],
            file_name="EG_Promo Cost File.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_promo_cost",
            use_container_width=True,
        )


    st.divider()

    st.caption(
        "Duplicate records are automatically removed before "
        "the output files are generated."
    )
