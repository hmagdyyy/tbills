import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="T-Bills NAV Helper", layout="wide")

st.title("T-Bills NAV Helper")

st.markdown("""
1. Upload the **bills Excel file**  
2. Enter the **NAV**  
3. Automatically get the processed table + sum row  
4. Download as Excel
""")

# ---------------------------------------------------------
# Main Processing Function
# ---------------------------------------------------------
def process_file(file, nav):
    xls = pd.ExcelFile(file)
    sheet_name = xls.sheet_names[0]

    # Read raw to get B3 (row index 2, col index 1)
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    today_raw = raw.iloc[2, 1] if raw.shape[0] > 2 else None
    today_date = pd.to_datetime(today_raw, errors="coerce")

    # Read main table starting from header row
    df = pd.read_excel(file, sheet_name=sheet_name, header=7)
    df = df.dropna(subset=["Bank"])

    # Convert numerics
    numeric_cols = [
        "FaceValue", "RatePercent", "PurchaseValue", "NumOfDays",
        "RemainPeriod", "PaidValue", "CurrentAccrude", "Tax"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # -------------------
    # Calculations
    # -------------------
    df["After Tax"] = df["RatePercent"] * 0.8
    df["WeightFromNAV"] = df["PurchaseValue"] / nav
    df["AvgReturn"] = df["After Tax"] * df["WeightFromNAV"]
    df["TodayDate"] = today_date
    df["Avg # Of Days"] = df["WeightFromNAV"] * df["RemainPeriod"]
    df["NAV"] = nav

    # -------------------
    # Format dates
    # -------------------
    date_cols = ["PurchaseDate", "MaturityDate", "TodayDate"]
    for c in date_cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.strftime("%d-%m-%Y")

    # -------------------
    # Round numbers
    # -------------------
    for col in df.columns:
        if df[col].dtype in ["float64", "int64"]:
            df[col] = df[col].round(2)

    # -------------------
    # Create SUM row
    # -------------------
    sum_row = pd.DataFrame([{
        "Bank": "TOTAL",
        "FaceValue": df["FaceValue"].sum(),
        "PurchaseValue": df["PurchaseValue"].sum(),
        "AvgReturn": df["AvgReturn"].sum(),
        "Avg # Of Days": df["Avg # Of Days"].sum(),
        "PaidValue": df["PaidValue"].sum(),
        "CurrentAccrude": df["CurrentAccrude"].sum(),
        "Tax": df["Tax"].sum(),
    }])

    # Round sum row
    for col in sum_row.columns:
        if col != "Bank":
            sum_row[col] = round(sum_row[col], 2)

    # Append to end
    df_final = pd.concat([df, sum_row], ignore_index=True)

    # -------------------
    # Final order
    # -------------------
    final_cols = [
        "Bank", "FaceValue", "RatePercent", "After Tax", "PurchaseValue",
        "PurchaseDate", "WeightFromNAV", "AvgReturn", "TodayDate", "MaturityDate",
        "NumOfDays", "RemainPeriod", "Avg # Of Days", "PaidValue",
        "CurrentAccrude", "Tax", "NAV"
    ]
    final_cols = [c for c in final_cols if c in df_final.columns]

    return df_final[final_cols]


# ---------------------------------------------------------
# Excel Export Helper
# ---------------------------------------------------------
def to_excel_bytes(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Result")
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------
uploaded_file = st.file_uploader("Upload Excel file", type=["xls", "xlsx"])
nav_value = st.number_input("Enter NAV", min_value=0.01, step=100.0, format="%.2f")

if uploaded_file and nav_value:
    try:
        result_df = process_file(uploaded_file, nav_value)

        st.subheader("Processed Table with SUM Row")
        st.dataframe(result_df, use_container_width=True)

        excel_data = to_excel_bytes(result_df)

        st.download_button(
            label="📥 Download Excel",
            data=excel_data,
            file_name="t_bills_nav_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")

else:
    st.info("Upload a file and enter NAV to proceed.")
