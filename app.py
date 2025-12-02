import io
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from datetime import datetime

st.set_page_config(page_title="T-Bills NAV Helper", layout="wide")
st.title("T-Bills NAV Helper")


# -----------------------------------------
# Process the uploaded file into a DataFrame (for preview)
# -----------------------------------------
def process_file(file, nav):
    xls = pd.ExcelFile(file)
    sheet_name = xls.sheet_names[0]

    # Get TodayDate from B3 (row index 2, col index 1)
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    today_raw = raw.iloc[2, 1] if raw.shape[0] > 2 and raw.shape[1] > 1 else None
    today_date = pd.to_datetime(today_raw, errors="coerce")

    # Main table (header row = index 7 -> row 8 in Excel)
    df = pd.read_excel(file, sheet_name=sheet_name, header=7)

    # Keep only rows with a Bank
    df = df.dropna(subset=["Bank"])

    # Numeric columns
    numeric_cols = [
        "FaceValue", "RatePercent", "PurchaseValue", "NumOfDays",
        "RemainPeriod", "PaidValue", "CurrentAccrude", "Tax"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Date columns
    for c in ["PurchaseDate", "MaturityDate"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # Core calculations (for preview)
    df["After Tax"] = df["RatePercent"] * 0.8
    df["WeightFromNAV"] = df["PurchaseValue"] / nav
    df["AvgReturn"] = df["After Tax"] * df["WeightFromNAV"]
    df["TodayDate"] = today_date
    df["Avg # Of Days"] = df["WeightFromNAV"] * df["RemainPeriod"]
    df["NAV"] = nav

    # New columns:
    # Price = 1 / ((NumOfDays * RatePercent)/365) + 1
    df["Price"] = 365 / ((df["NumOfDays"] * df["RatePercent"]) + 365)

    # PX = Price rounded to 5 dp (preview; Excel has ROUND)
    df["PX"] = df["Price"].round(5)


    # Accruals = (100 - (Price * 100)) * (TodayDate - PurchaseDate) / NumOfDays
    # Use day count difference
    days_diff = (df["TodayDate"] - df["PurchaseDate"]).dt.days
    df["Accruals"] = (100 - (df["Price"] * 100)) * (days_diff / df["NumOfDays"])

    # New PX = (PX * 100) + Accruals
    df["New PX"] = (df["PX"] * 100) + df["Accruals"]

    # Breakeven = 365 * ((100 / New PX) - 1) / RemainPeriod
    df["Breakeven"] = 365 * ((100 / df["New PX"]) - 1) / df["RemainPeriod"]

    # Sort by ascending remaining period
    if "RemainPeriod" in df.columns:
        df = df.sort_values("RemainPeriod", ascending=True)

    # Rounding for preview (Excel handles actual number formats)
    num_cols_preview = df.select_dtypes(include=["float64", "int64"]).columns
    for col in num_cols_preview:
        if col == "PX":
            df[col] = df[col].round(5)
        else:
            df[col] = df[col].round(2)

    # TOTAL row for preview only (same set as before)
    sum_row = {
        "Bank": "TOTAL",
        "FaceValue": df["FaceValue"].sum(),
        "PurchaseValue": df["PurchaseValue"].sum(),
        "AvgReturn": df["AvgReturn"].sum(),
        "Avg # Of Days": df["Avg # Of Days"].sum(),
        "PaidValue": df["PaidValue"].sum(),
        "CurrentAccrude": df["CurrentAccrude"].sum(),
        "Tax": df["Tax"].sum(),
        "NAV": nav
    }
    df = pd.concat([df, pd.DataFrame([sum_row])], ignore_index=True)

    # Final column order
    final_cols = [
        "Bank",
        "FaceValue",
        "RatePercent",
        "PurchaseValue",
        "PurchaseDate",
        "WeightFromNAV",
        "AvgReturn",
        "TodayDate",
        "MaturityDate",
        "NumOfDays",
        "RemainPeriod",
        "Avg # Of Days",
        "PaidValue",
        "CurrentAccrude",
        "Tax",
        "After Tax",
        "Price",
        "PX",
        "Accruals",
        "New PX",
        "Breakeven",
        "NAV",
    ]
    final_cols = [c for c in final_cols if c in df.columns]

    return df[final_cols]


# -----------------------------------------
# Build Excel with formulas & formatting
# -----------------------------------------
def to_excel_bytes(df: pd.DataFrame, nav: float) -> bytes:
    buffer = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Result"

    # Column order from DataFrame
    final_cols = list(df.columns)

    # Header row
    for col_idx, col_name in enumerate(final_cols, start=1):
        ws.cell(row=1, column=col_idx, value=col_name)

    # Data rows (exclude TOTAL row; that will be formula-based)
    data_df = df[df["Bank"] != "TOTAL"].reset_index(drop=True)
    rows_no_total = len(data_df)
    last_data_row = rows_no_total + 1   # data rows occupy Excel rows 2..last_data_row
    total_row_idx = last_data_row + 1   # TOTAL row index

    # Map: column name -> 1-based index
    col_index = {name: i + 1 for i, name in enumerate(final_cols)}

    # Write data rows with formulas
    for i, row in data_df.iterrows():
        r = i + 2  # Excel row (header is row 1)

        # Basic values
        ws.cell(row=r, column=col_index["Bank"], value=row["Bank"])
        if "FaceValue" in col_index:
            ws.cell(row=r, column=col_index["FaceValue"], value=row.get("FaceValue"))
        if "RatePercent" in col_index:
            ws.cell(row=r, column=col_index["RatePercent"], value=row.get("RatePercent"))
        if "PurchaseValue" in col_index:
            ws.cell(row=r, column=col_index["PurchaseValue"], value=row.get("PurchaseValue"))
        if "NumOfDays" in col_index:
            ws.cell(row=r, column=col_index["NumOfDays"], value=row.get("NumOfDays"))
        if "RemainPeriod" in col_index:
            ws.cell(row=r, column=col_index["RemainPeriod"], value=row.get("RemainPeriod"))
        if "PaidValue" in col_index:
            ws.cell(row=r, column=col_index["PaidValue"], value=row.get("PaidValue"))
        if "CurrentAccrude" in col_index:
            ws.cell(row=r, column=col_index["CurrentAccrude"], value=row.get("CurrentAccrude"))
        if "Tax" in col_index:
            ws.cell(row=r, column=col_index["Tax"], value=row.get("Tax"))

        # Date values (real dates; formatting applied later)
        if "PurchaseDate" in col_index:
            val = row.get("PurchaseDate")
            ws.cell(row=r, column=col_index["PurchaseDate"], value=val if pd.notna(val) else None)
        if "TodayDate" in col_index:
            val = row.get("TodayDate")
            ws.cell(row=r, column=col_index["TodayDate"], value=val if pd.notna(val) else None)
        if "MaturityDate" in col_index:
            val = row.get("MaturityDate")
            ws.cell(row=r, column=col_index["MaturityDate"], value=val if pd.notna(val) else None)

        # NAV column
        nav_col = col_index["NAV"]
        if r == 2:
            ws.cell(row=r, column=nav_col, value=nav)
        else:
            first_nav_cell = ws.cell(row=2, column=nav_col).coordinate
            ws.cell(row=r, column=nav_col, value=f"={first_nav_cell}")

        # Formulas

        # After Tax = RatePercent * 0.8
        if "After Tax" in col_index:
            c_rate = ws.cell(row=r, column=col_index["RatePercent"]).coordinate
            ws.cell(row=r, column=col_index["After Tax"], value=f"={c_rate}*0.8")

        # WeightFromNAV = PurchaseValue / NAV (same row)
        if "WeightFromNAV" in col_index:
            c_purchase = ws.cell(row=r, column=col_index["PurchaseValue"]).coordinate
            c_nav = ws.cell(row=r, column=nav_col).coordinate
            ws.cell(row=r, column=col_index["WeightFromNAV"], value=f"={c_purchase}/{c_nav}")

        # AvgReturn = After Tax * WeightFromNAV
        if "AvgReturn" in col_index:
            c_after_tax = ws.cell(row=r, column=col_index["After Tax"]).coordinate
            c_weight = ws.cell(row=r, column=col_index["WeightFromNAV"]).coordinate
            ws.cell(row=r, column=col_index["AvgReturn"], value=f"={c_after_tax}*{c_weight}")

        # Avg # Of Days = WeightFromNAV * RemainPeriod
        if "Avg # Of Days" in col_index:
            c_weight = ws.cell(row=r, column=col_index["WeightFromNAV"]).coordinate
            c_remain = ws.cell(row=r, column=col_index["RemainPeriod"]).coordinate
            ws.cell(row=r, column=col_index["Avg # Of Days"], value=f"={c_weight}*{c_remain}")

        # Price = 1 / ((NumOfDays * RatePercent) / 365) + 1
        if "Price" in col_index:
            c_days = ws.cell(row=r, column=col_index["NumOfDays"]).coordinate
            c_rate = ws.cell(row=r, column=col_index["RatePercent"]).coordinate
            ws.cell(
                row=r,
                column=col_index["Price"],
                value=f"=1/(({c_days}*{c_rate})/365)+1"
            )

        # PX = ROUND(Price, 5)
        if "PX" in col_index and "Price" in col_index:
            c_price = ws.cell(row=r, column=col_index["Price"]).coordinate
            ws.cell(
                row=r,
                column=col_index["PX"],
                value=f"=ROUND({c_price},5)"
            )


        # Accruals = (100 - (Price * 100)) * ((TodayDate - PurchaseDate) / NumOfDays)
        if "Accruals" in col_index:
            c_price = ws.cell(row=r, column=col_index["Price"]).coordinate
            c_today = ws.cell(row=r, column=col_index["TodayDate"]).coordinate
            c_pur = ws.cell(row=r, column=col_index["PurchaseDate"]).coordinate
            c_days = ws.cell(row=r, column=col_index["NumOfDays"]).coordinate
            ws.cell(
                row=r,
                column=col_index["Accruals"],
                value=f"=(100-({c_price}*100))*(({c_today}-{c_pur})/{c_days})"
            )

        # New PX = (PX * 100) + Accruals
        if "New PX" in col_index:
            c_px = ws.cell(row=r, column=col_index["PX"]).coordinate
            c_accr = ws.cell(row=r, column=col_index["Accruals"]).coordinate
            ws.cell(
                row=r,
                column=col_index["New PX"],
                value=f"=({c_px}*100)+{c_accr}"
            )

        # Breakeven = 365 * ((100/New PX) - 1) / RemainPeriod
        if "Breakeven" in col_index:
            c_newpx = ws.cell(row=r, column=col_index["New PX"]).coordinate
            c_remain = ws.cell(row=r, column=col_index["RemainPeriod"]).coordinate
            ws.cell(
                row=r,
                column=col_index["Breakeven"],
                value=f"=365*((100/{c_newpx})-1)/{c_remain}"
            )

    # TOTAL row with SUM formulas (same set as earlier)
    ws.cell(row=total_row_idx, column=col_index["Bank"], value="TOTAL")

    def add_sum(col_name):
        c_idx = col_index[col_name]
        col_letter = ws.cell(row=1, column=c_idx).column_letter
        ws.cell(
            row=total_row_idx,
            column=c_idx,
            value=f"=SUM({col_letter}2:{col_letter}{last_data_row})"
        )

    for cname in ["FaceValue", "PurchaseValue", "AvgReturn", "Avg # Of Days",
                  "PaidValue", "CurrentAccrude", "Tax"]:
        if cname in col_index:
            add_sum(cname)

    # NAV in TOTAL row = NAV in first data row
    ws.cell(
        row=total_row_idx,
        column=col_index["NAV"],
        value=f"={ws.cell(row=2, column=col_index['NAV']).coordinate}"
    )

    # Number formatting
    number_fmt_2 = '#,##0.00'
    number_fmt_5 = '#,##0.00000'

    numeric_columns = [
        "FaceValue", "RatePercent", "After Tax", "Price", "PX",
        "Accruals", "New PX", "Breakeven",
        "PurchaseValue","WeightFromNAV", "AvgReturn",
        "NumOfDays", "RemainPeriod", "Avg # Of Days",
        "PaidValue", "CurrentAccrude", "Tax", "NAV"
    ]

    for cname in numeric_columns:
        if cname not in col_index:
            continue
        c_idx = col_index[cname]
        fmt = number_fmt_5 if cname == "PX" else number_fmt_2
        for r in range(2, total_row_idx + 1):
            cell = ws.cell(row=r, column=c_idx)
            if isinstance(cell.value, (int, float)) or (isinstance(cell.value, str) and cell.value.startswith("=")):
                cell.number_format = fmt

    # Date formatting: dd-mm-yyyy
    date_fmt = "DD-MM-YYYY"
    for cname in ["PurchaseDate", "TodayDate", "MaturityDate"]:
        if cname not in col_index:
            continue
        c_idx = col_index[cname]
        for r in range(2, total_row_idx + 1):
            cell = ws.cell(row=r, column=c_idx)
            if isinstance(cell.value, datetime):
                cell.number_format = date_fmt

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------------------
# Streamlit UI
# -----------------------------------------
uploaded_file = st.file_uploader("Upload Excel file", type=["xls", "xlsx"])
nav_value = st.number_input("Enter NAV", min_value=0.01, step=100.0, format="%.2f")

if uploaded_file and nav_value:
    try:
        result_df = process_file(uploaded_file, nav_value)

        st.subheader("Processed Table (Preview)")
        st.dataframe(result_df, use_container_width=True)

        excel_data = to_excel_bytes(result_df, nav_value)

        st.download_button(
            label="📥 Download Excel (with formulas)",
            data=excel_data,
            file_name="t_bills_nav_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload a file and enter NAV to continue.")

