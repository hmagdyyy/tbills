import io
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from datetime import datetime

st.set_page_config(page_title="T-Bills NAV Helper", layout="wide")
st.title("T-Bills NAV Helper (Multi-Fund Workbook)")


# -----------------------------------------
# Per-sheet processing
# -----------------------------------------
def process_sheet(xls, sheet_name: str, nav: float) -> pd.DataFrame:
    """
    Process a single sheet into the final DataFrame for preview + export.
    - xls: pd.ExcelFile (opened on the uploaded workbook)
    - sheet_name: name of the sheet (fund)
    - nav: NAV for this fund
    """

    # --- Get TodayDate from B3 (row 3, col B: index [2,1]) ---
    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    today_raw = raw.iloc[2, 1] if raw.shape[0] > 2 and raw.shape[1] > 1 else None
    today_date = pd.to_datetime(today_raw, errors="coerce", dayfirst=True)

    # --- Main table starting row 8 (header row index = 7) ---
    df = pd.read_excel(xls, sheet_name=sheet_name, header=7)

    # Keep only rows with a Bank
    df = df.dropna(subset=["Bank"])

    # Numeric cols in original sheet
    numeric_cols = [
        "FaceValue", "RatePercent", "PurchaseValue", "NumOfDays",
        "RemainPeriod", "PaidValue", "CurrentAccrude", "Tax"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Make RatePercent a fraction (27 → 0.27) but keep full precision
    if "RatePercent" in df.columns:
        df["RatePercent"] = df["RatePercent"] / 100.0

    # Date columns – enforce dd/mm/yyyy (day-first)
    for c in ["PurchaseDate", "MaturityDate"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)

    # Rename for output
    df.rename(columns={"PurchaseValue": "PV", "PaidValue": "CurrentValue"}, inplace=True)

    # Core calculations
    df["After Tax"] = df["RatePercent"] * 0.8
    df["WeightFromNAV"] = df["PV"] / nav
    df["AvgReturn"] = df["After Tax"] * df["WeightFromNAV"]
    df["TodayDate"] = today_date
    df["Avg # Of Days"] = df["WeightFromNAV"] * df["RemainPeriod"]
    df["NAV"] = nav

    # Price = 365 / ((NumOfDays * RatePercent) + 365)
    df["Price"] = 365 / ((df["NumOfDays"] * df["RatePercent"]) + 365)

    # PX = Price rounded to 5dp (preview; Excel also uses ROUND)
    df["PX"] = df["Price"].round(5)

    # Accruals = (100 - (Price*100)) * (TodayDate - PurchaseDate) / NumOfDays
    days_diff = (df["TodayDate"] - df["PurchaseDate"]).dt.days
    df["Accruals"] = (100 - (df["Price"] * 100)) * (days_diff / df["NumOfDays"])

    # New PX = (PX*100) + Accruals
    df["New PX"] = (df["PX"] * 100) + df["Accruals"]

    # Breakeven = 365 * ((100 / New PX) - 1) / RemainPeriod
    df["Breakeven"] = 365 * ((100 / df["New PX"]) - 1) / df["RemainPeriod"]

    # Sort by ascending remaining period
    if "RemainPeriod" in df.columns:
        df = df.sort_values("RemainPeriod", ascending=True)

    # Rounding for preview (Excel recomputes from formulas)
    num_cols_preview = df.select_dtypes(include=["float64", "int64"]).columns
    for col in num_cols_preview:
        if col == "PX":
            df[col] = df[col].round(5)
        elif col == "Breakeven":
            df[col] = df[col].round(3)
        elif col == "RatePercent":
            # keep full precision, no rounding
            continue
        else:
            df[col] = df[col].round(2)

    # TOTAL row (preview only)
    sum_row = {
        "Bank": "TOTAL",
        "FaceValue": df["FaceValue"].sum() if "FaceValue" in df.columns else None,
        "PV": df["PV"].sum() if "PV" in df.columns else None,
        "AvgReturn": df["AvgReturn"].sum() if "AvgReturn" in df.columns else None,
        "Avg # Of Days": df["Avg # Of Days"].sum() if "Avg # Of Days" in df.columns else None,
        "CurrentValue": df["CurrentValue"].sum() if "CurrentValue" in df.columns else None,
        "CurrentAccrude": df["CurrentAccrude"].sum() if "CurrentAccrude" in df.columns else None,
        "Tax": df["Tax"].sum() if "Tax" in df.columns else None,
        "NAV": nav,
    }
    df = pd.concat([df, pd.DataFrame([sum_row])], ignore_index=True)

    # Final column order
    final_cols = [
        "Bank",
        "NAV",
        "FaceValue",
        "PV",
        "RatePercent",
        "After Tax",
        "PurchaseDate",
        "MaturityDate",
        "NumOfDays",
        "Price",
        "PX",
        "TodayDate",
        "RemainPeriod",
        "Breakeven",
        "CurrentValue",
        "New PX",
        "Accruals",
        "CurrentAccrude",
        "WeightFromNAV",
        "AvgReturn",
        "Avg # Of Days",
        "Tax",
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

    final_cols = list(df.columns)

    # Header row
    for col_idx, col_name in enumerate(final_cols, start=1):
        ws.cell(row=1, column=col_idx, value=col_name)

    # Data rows (exclude TOTAL; that row is sums)
    data_df = df[df["Bank"] != "TOTAL"].reset_index(drop=True)
    rows_no_total = len(data_df)
    last_data_row = rows_no_total + 1    # rows 2..last_data_row
    total_row_idx = last_data_row + 1    # row for TOTAL

    col_index = {name: i + 1 for i, name in enumerate(final_cols)}

    # --- Write rows and formulas ---
    for i, row in data_df.iterrows():
        r = i + 2  # Excel row index

        # Basic values
        ws.cell(row=r, column=col_index["Bank"], value=row["Bank"])
        for name in ["FaceValue", "RatePercent", "PV", "NumOfDays",
                     "RemainPeriod", "CurrentValue", "CurrentAccrude", "Tax"]:
            if name in col_index:
                ws.cell(row=r, column=col_index[name], value=row.get(name))

        # Dates
        for cname in ["PurchaseDate", "TodayDate", "MaturityDate"]:
            if cname in col_index:
                val = row.get(cname)
                ws.cell(row=r, column=col_index[cname], value=val if pd.notna(val) else None)

        # NAV column
        nav_col = col_index["NAV"]
        if r == 2:
            ws.cell(row=r, column=nav_col, value=nav)
        else:
            first_nav_cell = ws.cell(row=2, column=nav_col).coordinate
            ws.cell(row=r, column=nav_col, value=f"={first_nav_cell}")

        # After Tax = RatePercent * 0.8
        if "After Tax" in col_index:
            c_rate = ws.cell(row=r, column=col_index["RatePercent"]).coordinate
            ws.cell(row=r, column=col_index["After Tax"], value=f"={c_rate}*0.8")

        # WeightFromNAV = PV / NAV
        if "WeightFromNAV" in col_index:
            c_pv = ws.cell(row=r, column=col_index["PV"]).coordinate
            c_nav = ws.cell(row=r, column=nav_col).coordinate
            ws.cell(row=r, column=col_index["WeightFromNAV"], value=f"={c_pv}/{c_nav}")

        # AvgReturn = After Tax * WeightFromNAV
        if "AvgReturn" in col_index:
            c_at = ws.cell(row=r, column=col_index["After Tax"]).coordinate
            c_w = ws.cell(row=r, column=col_index["WeightFromNAV"]).coordinate
            ws.cell(row=r, column=col_index["AvgReturn"], value=f"={c_at}*{c_w}")

        # Avg # Of Days = WeightFromNAV * RemainPeriod
        if "Avg # Of Days" in col_index:
            c_w = ws.cell(row=r, column=col_index["WeightFromNAV"]).coordinate
            c_rem = ws.cell(row=r, column=col_index["RemainPeriod"]).coordinate
            ws.cell(row=r, column=col_index["Avg # Of Days"], value=f"={c_w}*{c_rem}")

        # Price = 365 / ((NumOfDays * RatePercent) + 365)
        if "Price" in col_index:
            c_days = ws.cell(row=r, column=col_index["NumOfDays"]).coordinate
            c_rate = ws.cell(row=r, column=col_index["RatePercent"]).coordinate
            ws.cell(
                row=r,
                column=col_index["Price"],
                value=f"=365/(({c_days}*{c_rate})+365)"
            )

        # PX = ROUND(Price, 5)
        if "PX" in col_index and "Price" in col_index:
            c_price = ws.cell(row=r, column=col_index["Price"]).coordinate
            ws.cell(
                row=r,
                column=col_index["PX"],
                value=f"=ROUND({c_price},5)"
            )

        # Accruals = (100 - (Price*100)) * ((TodayDate - PurchaseDate)/NumOfDays)
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

        # New PX = (PX*100) + Accruals
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
            c_rem = ws.cell(row=r, column=col_index["RemainPeriod"]).coordinate
            ws.cell(
                row=r,
                column=col_index["Breakeven"],
                value=f"=365*((100/{c_newpx})-1)/{c_rem}"
            )

    # --- TOTAL row with SUM formulas ---
    ws.cell(row=total_row_idx, column=col_index["Bank"], value="TOTAL")

    def add_sum(col_name: str):
        c_idx = col_index[col_name]
        col_letter = ws.cell(row=1, column=c_idx).column_letter
        ws.cell(
            row=total_row_idx,
            column=c_idx,
            value=f"=SUM({col_letter}2:{col_letter}{last_data_row})"
        )

    for cname in ["FaceValue", "PV", "AvgReturn", "Avg # Of Days",
                  "CurrentValue", "CurrentAccrude", "Tax"]:
        if cname in col_index:
            add_sum(cname)

    # NAV in TOTAL row = NAV in first data row
    ws.cell(
        row=total_row_idx,
        column=col_index["NAV"],
        value=f"={ws.cell(row=2, column=col_index['NAV']).coordinate}"
    )

    # --- Formatting ---
    number_fmt_2 = '#,##0.00'
    number_fmt_5 = '#,##0.00000'

    # Numeric columns (non-percentage)
    numeric_columns = [
        "FaceValue", "Price", "PX", "Accruals", "New PX",
        "PV", "WeightFromNAV", "AvgReturn",
        "NumOfDays", "RemainPeriod", "Avg # Of Days",
        "CurrentValue", "CurrentAccrude", "Tax", "NAV"
    ]
    for cname in numeric_columns:
        if cname not in col_index:
            continue
        c_idx = col_index[cname]
        fmt = number_fmt_5 if cname == "PX" else number_fmt_2
        for r in range(2, total_row_idx + 1):
            cell = ws.cell(row=r, column=c_idx)
            if isinstance(cell.value, (int, float)) or (
                isinstance(cell.value, str) and cell.value.startswith("=")
            ):
                cell.number_format = fmt

    # RatePercent as percentage, 4 decimals (no rounding of stored value)
    if "RatePercent" in col_index:
        c_idx = col_index["RatePercent"]
        for r in range(2, total_row_idx + 1):
            cell = ws.cell(row=r, column=c_idx)
            if isinstance(cell.value, (int, float)) or (
                isinstance(cell.value, str) and cell.value.startswith("=")
            ):
                cell.number_format = "0.0000%"

    # After Tax as percentage, 2 decimals
    if "After Tax" in col_index:
        c_idx = col_index["After Tax"]
        for r in range(2, total_row_idx + 1):
            cell = ws.cell(row=r, column=c_idx)
            if isinstance(cell.value, (int, float)) or (
                isinstance(cell.value, str) and cell.value.startswith("=")
            ):
                cell.number_format = "0.00%"

    # WeightFromNAV as percentage, 2 decimals
    if "WeightFromNAV" in col_index:
        c_idx = col_index["WeightFromNAV"]
        for r in range(2, total_row_idx + 1):
            cell = ws.cell(row=r, column=c_idx)
            if isinstance(cell.value, (int, float)) or (
                isinstance(cell.value, str) and cell.value.startswith("=")
            ):
                cell.number_format = "0.00%"

    # Breakeven as percentage, 3 decimals
    if "Breakeven" in col_index:
        c_idx = col_index["Breakeven"]
        for r in range(2, total_row_idx + 1):
            cell = ws.cell(row=r, column=c_idx)
            if isinstance(cell.value, (int, float)) or (
                isinstance(cell.value, str) and cell.value.startswith("=")
            ):
                cell.number_format = "0.000%"

    # Dates as dd/mm/yyyy
    date_fmt = "DD/MM/YYYY"
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
uploaded_file = st.file_uploader("Upload Excel workbook", type=["xls", "xlsx"])

if uploaded_file:
    # Load workbook once
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    st.markdown(f"**Found {len(sheet_names)} sheet(s):** {', '.join(sheet_names)}")

    # For each sheet (fund), extract Fund Name from B4 and show NAV input + export
    for sheet_name in sheet_names:
        # Read B4 = Fund Name (row 4, col B -> [3,1])
        raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if raw.shape[0] > 3 and raw.shape[1] > 1:
            fund_name = str(raw.iloc[3, 1]).strip()
        else:
            fund_name = sheet_name

        with st.expander(f"Fund: {fund_name} (sheet: {sheet_name})", expanded=False):
            nav_value = st.number_input(
                f"NAV for {fund_name}",
                min_value=0.01,
                step=100.0,
                format="%.2f",
                key=f"nav_{sheet_name}",
            )

            if nav_value and nav_value > 0:
                # Process this sheet
                df_result = process_sheet(xls, sheet_name, nav_value)

                st.write("Preview:")
                st.dataframe(df_result, use_container_width=True)

                # Build Excel bytes
                excel_data = to_excel_bytes(df_result, nav_value)

                safe_fund_name = fund_name.replace(" ", "_") or sheet_name
                st.download_button(
                    label=f"📥 Download Excel for {fund_name}",
                    data=excel_data,
                    file_name=f"{safe_fund_name}_t_bills_nav.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{sheet_name}",
                )
else:
    st.info("Upload a workbook with one or more T-Bills sheets to begin.")

