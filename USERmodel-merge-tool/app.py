# app.py – Smart Streamlit Excel MODEL merge tool

import streamlit as st
import pandas as pd
from io import BytesIO

# ---------- Page config ----------
st.set_page_config(
    page_title="Model Merge Tool",
    layout="wide",
    page_icon="📊",
)

# ---------- Header ----------
st.markdown(
    """
    <style>
    .big-title {font-size: 30px; font-weight: 700;}
    .sub {color: #555; font-size: 14px;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<p class="big-title">📊 Model / Style Quantity Merger</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub">Upload any Excel with MODEL / STYLE and QTY / AMOUNT columns. '
    'App will auto-detect columns, merge duplicates, and calculate unit price.</p>',
    unsafe_allow_html=True,
)
st.markdown("---")

# ---------- Helper: smart column detection ----------
def detect_column(df_cols, keywords):
    """
    df_cols: list of uppercase column names
    keywords: list of strings to search
    returns first matching column or None
    """
    for key in keywords:
        for col in df_cols:
            if key in col:
                return col
    return None

# ---------- File upload ----------
uploaded_file = st.file_uploader("Excel file upload karo (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # Read Excel
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"File read error: {e}")
        st.stop()

    st.success(f"File loaded: **{uploaded_file.name}**")
    st.write(f"Total rows: **{len(df)}**")

    # Show preview
    with st.expander("🔍 Original data preview (top 15 rows)", expanded=False):
        st.dataframe(df.head(15), use_container_width=True)

    # Normalize column names
    df_original_cols = df.columns.copy()
    df.columns = df.columns.str.strip().str.upper()

    cols = list(df.columns)

    # Auto-detect columns
    auto_model_col = detect_column(cols, ["MODEL", "STYLE", "ITEM", "CODE"])
    auto_qty_col = detect_column(cols, ["QTY", "QUANTITY", "PCS", "QTY/PCS"])
    auto_amount_col = detect_column(cols, ["AMOUNT", "TOTAL", "VALUE", "AMT"])

    # ---------- Sidebar: settings ----------
    st.sidebar.header("⚙️ Settings")

    st.sidebar.write("**Column selection** (auto-detected, but you can change):")
    model_col = st.sidebar.selectbox("Model / Style column", options=cols, index=cols.index(auto_model_col) if auto_model_col in cols else 0)
    qty_col = st.sidebar.selectbox("Quantity column", options=cols, index=cols.index(auto_qty_col) if auto_qty_col in cols else 0)
    amount_col = st.sidebar.selectbox("Amount column", options=cols, index=cols.index(auto_amount_col) if auto_amount_col in cols else 0)

    st.sidebar.markdown("---")
    st.sidebar.write("**Result options**:")

    min_qty_filter = st.sidebar.number_input("Minimum Total_QTY to keep", value=0, min_value=0, step=1)
    sort_option = st.sidebar.selectbox("Sort result by", options=["MODEL", "Total_QTY", "Total_Amount", "Unit_Price"])

    st.sidebar.markdown("---")
    search_text = st.sidebar.text_input("Filter by model name (contains)", value="")

    # ---------- Cleaning ----------
    working_df = df.copy()

    # Clean model
    working_df[model_col] = working_df[model_col].astype(str).str.strip()
    working_df[model_col] = working_df[model_col].replace(["", "nan", "NaN", "NONE", "None"], pd.NA)
    working_df[model_col] = working_df[model_col].fillna(method="ffill")

    # Numeric
    working_df[qty_col] = pd.to_numeric(working_df[qty_col], errors="coerce").fillna(0)
    working_df[amount_col] = pd.to_numeric(working_df[amount_col], errors="coerce").fillna(0)

    # ---------- Grouping ----------
    grouped = (
        working_df.groupby(model_col, as_index=False)
        .agg({qty_col: "sum", amount_col: "sum"})
    )
    grouped.columns = ["MODEL", "Total_QTY", "Total_Amount"]
    grouped["Unit_Price"] = (grouped["Total_Amount"] / grouped["Total_QTY"]).round(2)

    # Apply filters
    if min_qty_filter > 0:
        grouped = grouped[grouped["Total_QTY"] >= min_qty_filter]

    if search_text:
        grouped = grouped[grouped["MODEL"].str.contains(search_text, case=False, na=False)]

    # Sort
    grouped = grouped.sort_values(sort_option).reset_index(drop=True)

    # ---------- Display result ----------
    st.markdown("### ✅ Merged result")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Models", len(grouped))
    with col2:
        st.metric("Total QTY", int(grouped["Total_QTY"].sum()))
    with col3:
        st.metric("Total Amount", f"{grouped['Total_Amount'].sum():,.2f}")

    st.dataframe(grouped, use_container_width=True, height=450)

    # ---------- Download Excel ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        grouped.to_excel(writer, index=False, sheet_name="Merged Data")
    excel_data = output.getvalue()

    st.markdown("### ⬇️ Download")
    st.download_button(
        label="Download merged Excel",
        data=excel_data,
        file_name="Model_Merged_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    st.info("Start karne ke liye upar se koi bhi related Excel (.xlsx) upload karo.")
