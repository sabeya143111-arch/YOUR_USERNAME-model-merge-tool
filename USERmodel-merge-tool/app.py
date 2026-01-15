# app.py  –  Streamlit Excel MODEL merge tool

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Model Merge Tool", layout="wide")

st.title("📊 Model Quantity Merger")
st.write(
    "Same MODEL / STYLE ki multiple quantities ko merge karo, "
    "total amount aur unit price automatically calculate ho jayega."
)

st.markdown("---")

# 1) File upload
uploaded_file = st.file_uploader("Excel file upload karo (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # 2) Read Excel
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"File read error: {e}")
        st.stop()

    st.subheader("Original data (top 10 rows)")
    st.dataframe(df.head(10))

    # 3) Normalize column names
    df.columns = df.columns.str.strip().str.upper()

    # Auto-detect MODEL / STYLE column
    model_col = None
    for col in df.columns:
        if "MODEL" in col or "STYLE" in col:
            model_col = col
            break

    # Auto-detect QTY column
    qty_col = None
    for col in df.columns:
        if "QTY" in col or "QUANTITY" in col:
            qty_col = col
            break

    # Auto-detect AMOUNT column
    amount_col = None
    for col in df.columns:
        if "AMOUNT" in col or "TOTAL" in col:
            amount_col = col
            break

    if not model_col or not qty_col or not amount_col:
        st.error(
            "MODEL / QTY / AMOUNT columns automatically detect nahi ho paaye. "
            "Please ensure column names me in words ka use ho."
        )
        st.stop()

    st.markdown("### Detected columns")
    st.write(f"**MODEL column:** {model_col}")
    st.write(f"**QTY column:** {qty_col}")
    st.write(f"**AMOUNT column:** {amount_col}")

    # 4) Clean MODEL: trim, blanks fill
    df[model_col] = df[model_col].astype(str).str.strip()
    df[model_col] = df[model_col].replace(["", "nan", "NaN", "NONE", "None"], pd.NA)
    # Forward fill: neeche blank MODEL ko upar wali value de do
    df[model_col] = df[model_col].fillna(method="ffill")

    # 5) Convert QTY & AMOUNT to numeric
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)

    st.subheader("Cleaned data sample")
    st.dataframe(df[[model_col, qty_col, amount_col]].head(10))

    # 6) Group by MODEL and sum
    grouped = (
        df.groupby(model_col, as_index=False)
          .agg({qty_col: "sum", amount_col: "sum"})
    )

    # Rename columns
    grouped.columns = ["MODEL", "Total_QTY", "Total_Amount"]

    # 7) Calculate Unit Price
    grouped["Unit_Price"] = (grouped["Total_Amount"] / grouped["Total_QTY"]).round(2)

    # Sort by model
    grouped = grouped.sort_values("MODEL").reset_index(drop=True)

    st.markdown("### Merged result")
    st.dataframe(grouped)

    # Summary
    st.info(
        f"Total models: **{len(grouped)}**, "
        f"Total QTY: **{int(grouped['Total_QTY'].sum())}**, "
        f"Total Amount: **{grouped['Total_Amount'].sum():,.2f}**"
    )

    # 8) Prepare Excel for download
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        grouped.to_excel(writer, index=False, sheet_name="Merged Data")
    excel_data = output.getvalue()

    st.markdown("### Download")
    st.download_button(
        label="⬇️ Download merged Excel",
        data=excel_data,
        file_name="Model_Merged_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upar se Excel (.xlsx) file upload karo, phir result yahan show hoga.")
