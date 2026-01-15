# app.py – Professional Streamlit Invoice Model Merger

import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------- Page config ----------
st.set_page_config(
    page_title="Invoice Processor",
    layout="wide",
    page_icon="📦",
    initial_sidebar_state="expanded",
)

# ---------- Custom CSS ----------
st.markdown(
    """
    <style>
    .main-title {
        font-size: 36px;
        font-weight: 800;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 10px;
    }
    .subtitle {
        color: #cccccc;
        font-size: 16px;
        margin-bottom: 20px;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 12px;
        margin: 10px 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<p class="main-title">📦 Invoice Model Merger</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">Koi bhi invoice Excel upload karo. App MODEL ke hisaab se '
    'duplicate rows merge karega, saare columns safe rahenge, aur professional Excel '
    'output dega.</p>',
    unsafe_allow_html=True,
)
st.markdown("---")


# ---------- Helper: detect column ----------
def detect_column(df_cols, keywords):
    df_cols_upper = [str(c).upper() for c in df_cols]
    for key in keywords:
        key_up = key.upper()
        for col, col_up in zip(df_cols, df_cols_upper):
            if key_up in col_up:
                return col
    return None


# ---------- Main app ----------
uploaded_file = st.file_uploader(
    "Upload your invoice Excel (.xlsx)", type=["xlsx"]
)

if uploaded_file is not None:
    # 1) Read
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ File error: {e}")
        st.stop()

    st.markdown(
        f'<div class="success-box">✅ File loaded: <b>{uploaded_file.name}</b> '
        f'| Rows: {len(df)}</div>',
        unsafe_allow_html=True,
    )

    with st.expander("🔍 Preview: Original data (first 10 rows)", expanded=False):
        st.dataframe(df.head(10), use_container_width=True)

    # 2) Normalize columns
    original_cols = df.columns.copy()
    df.columns = df.columns.str.strip().str.upper()
    cols = list(df.columns)

    # 3) Auto-detect important columns (use original names for operations)
    model_col = detect_column(original_cols, ["MODEL", "STYLE", "ITEM", "CODE"])
    qty_col = detect_column(original_cols, ["QTY", "QUANTITY", "PCS"])
    price_col = detect_column(original_cols, ["PRICE", "U.PRICE", "UNIT"])
    amount_col = detect_column(original_cols, ["AMOUNT", "TOTAL", "VALUE"])

    if not model_col:
        st.error("❌ MODEL / STYLE column nahi mil raha! Column naam check karo.")
        st.stop()

    # ---------- Sidebar ----------
    st.sidebar.header("⚙️ Settings")
    st.sidebar.write("**Detected columns:**")
    st.sidebar.info(
        f"🔹 Model: **{model_col}**\n\n"
        f"🔹 Qty: **{qty_col if qty_col else 'Not found'}**\n\n"
        f"🔹 Price: **{price_col if price_col else 'Not found'}**\n\n"
        f"🔹 Amount: **{amount_col if amount_col else 'Not found'}**"
    )

    st.sidebar.markdown("---")
    st.sidebar.write("**Result filters:**")
    min_qty = st.sidebar.number_input("Min Total QTY", value=0, min_value=0, step=1)
    search_model = st.sidebar.text_input("Filter model (contains)", value="")
    sort_by = st.sidebar.selectbox(
        "Sort by",
        options=[
            "MODEL",
            "Total_QTY",
            "Total_Amount",
        ]
        if amount_col and qty_col
        else (["MODEL", "Total_QTY"] if qty_col else ["MODEL"]),
        index=0,
    )

    # 4) Clean data
    working_df = df.copy()

    # Clean MODEL
    working_df[model_col] = working_df[model_col].astype(str).str.strip()
    working_df[model_col] = working_df[model_col].replace(
        ["", "nan", "NaN", "NONE", "None"], pd.NA
    )
    working_df[model_col] = working_df[model_col].fillna(method="ffill")

    # Convert numeric columns
    if qty_col:
        working_df[qty_col] = pd.to_numeric(
            working_df[qty_col], errors="coerce"
        ).fillna(0)
    if price_col:
        working_df[price_col] = pd.to_numeric(
            working_df[price_col], errors="coerce"
        ).fillna(0)
    if amount_col:
        working_df[amount_col] = pd.to_numeric(
            working_df[amount_col], errors="coerce"
        ).fillna(0)

    # 5) Groupby MODEL – sum QTY, AMOUNT; keep first of other columns
    agg_dict = {}
    if qty_col:
        agg_dict[qty_col] = "sum"
    if price_col:
        agg_dict[price_col] = "first"  # Keep first price
    if amount_col:
        agg_dict[amount_col] = "sum"

    for col in working_df.columns:
        if col != model_col and col not in agg_dict:
            agg_dict[col] = "first"

    grouped = working_df.groupby(model_col, as_index=False).agg(agg_dict)

    # Rename helper columns (only for display)
    display_df = grouped.copy()
    if qty_col:
        display_df["Total_QTY"] = display_df
