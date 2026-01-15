# app.py – Professional Streamlit Excel Merger (keeps all columns, groups by MODEL)

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
        color: #666;
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
    '<p class="subtitle">Upload any invoice Excel file. App will merge duplicate models '
    'while keeping all details. Beautiful professional output!</p>',
    unsafe_allow_html=True,
)
st.markdown("---")


# ---------- Helper: detect column ----------
def detect_column(df_cols, keywords):
    for key in keywords:
        for col in df_cols:
            if key in col:
                return col
    return None


# ---------- Main app ----------
uploaded_file = st.file_uploader("Upload your invoice Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # 1) Read
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ File error: {e}")
        st.stop()

    st.markdown(f'<div class="success-box">✅ File loaded: <b>{uploaded_file.name}</b> | Rows: {len(df)}</div>', unsafe_allow_html=True)

    with st.expander("🔍 Preview: Original data (first 10 rows)", expanded=False):
        st.dataframe(df.head(10), use_container_width=True)

    # 2) Normalize columns
    original_cols = df.columns.copy()
    df.columns = df.columns.str.strip().str.upper()
    cols = list(df.columns)

    # 3) Auto-detect MODEL column
    model_col = detect_column(cols, ["MODEL", "STYLE", "ITEM", "CODE"])
    qty_col = detect_column(cols, ["QTY", "QUANTITY", "PCS"])
    price_col = detect_column(cols, ["PRICE", "U.PRICE", "UNIT"])
    amount_col = detect_column(cols, ["AMOUNT", "TOTAL", "VALUE"])

    if not model_col:
        st.error("❌ MODEL / STYLE column nahi mil raha!")
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
        options=["MODEL", "Total_QTY", "Total_Amount"] if amount_col else ["MODEL", "Total_QTY"],
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
        working_df[qty_col] = pd.to_numeric(working_df[qty_col], errors="coerce").fillna(0)
    if price_col:
        working_df[price_col] = pd.to_numeric(working_df[price_col], errors="coerce").fillna(0)
    if amount_col:
        working_df[amount_col] = pd.to_numeric(working_df[amount_col], errors="coerce").fillna(0)

    # 5) Groupby MODEL – sum QTY, PRICE, AMOUNT; keep first of other columns
    agg_dict = {}
    if qty_col:
        agg_dict[qty_col] = "sum"
    if price_col:
        agg_dict[price_col] = "first"  # Keep first price
    if amount_col:
        agg_dict[amount_col] = "sum"

    # For other columns, keep first value
    for col in working_df.columns:
        if col != model_col and col not in agg_dict:
            agg_dict[col] = "first"

    grouped = working_df.groupby(model_col, as_index=False).agg(agg_dict)

    # Rename for clarity
    grouped_renamed = grouped.copy()
    if qty_col:
        grouped_renamed[f"Total_{qty_col}"] = grouped_renamed[qty_col]
    if amount_col:
        grouped_renamed[f"Total_{amount_col}"] = grouped_renamed[amount_col]

    # 6) Apply filters
    if min_qty > 0 and qty_col:
        grouped = grouped[grouped[qty_col] >= min_qty]

    if search_model:
        grouped = grouped[
            grouped[model_col].str.contains(search_model, case=False, na=False)
        ]

    # Sort
    sort_col = qty_col if sort_by == "Total_QTY" else (amount_col if sort_by == "Total_Amount" else model_col)
    grouped = grouped.sort_values(sort_col, ascending=False).reset_index(drop=True)

    # 7) Display
    st.markdown("### ✅ Merged Result")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Total Models", len(grouped))
    with c2:
        if qty_col:
            st.metric("Total Quantity", int(grouped[qty_col].sum()))
    with c3:
        if amount_col:
            st.metric("Total Amount", f"{grouped[amount_col].sum():,.2f}")

    st.dataframe(grouped, use_container_width=True, height=400)

    # 8) Create Beautiful Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        grouped.to_excel(writer, sheet_name="Merged Data", index=False)

        ws = writer.sheets["Merged Data"]

        # Styling
        header_fill = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        header_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )

        # Header row
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.border = border

        # Data rows – alternating colors
        light_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=len(grouped) + 1, min_col=1, max_col=len(grouped.columns)), start=2):
            fill = light_fill if (row_idx - 2) % 2 == 0 else white_fill
            for cell in row:
                cell.fill = fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border
                # Numeric formatting
                try:
                    if isinstance(cell.value, (int, float)) and "." in str(cell.value):
                        cell.number_format = "0.00"
                except:
                    pass

        # Auto column width
        for column in ws.columns:
            max_len = 0
            col_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 2, 30)

        # Freeze header
        ws.freeze_panes = "A2"

    excel_bytes = output.getvalue()

    st.markdown("### ⬇️ Download")
    st.download_button(
        label="📥 Download Professional Excel",
        data=excel_bytes,
        file_name=f"Invoice_Merged_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown("---")
    st.success("✨ Processing complete! Download your professional Excel file.")

else:
    st.info("📤 Start by uploading an Excel file above.")
