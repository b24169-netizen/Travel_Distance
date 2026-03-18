import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("EMS Customer Invoice Generator")

# -----------------------------
# Upload File
# -----------------------------
uploaded_file = st.file_uploader("Upload EMS Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Raw Data Preview")
    st.dataframe(df.head())

    # -----------------------------
    # Column Mapping
    # -----------------------------
    st.sidebar.header("Column Mapping")

    customer_col = st.sidebar.selectbox("Customer Column", df.columns)
    employee_col = st.sidebar.selectbox("Employee Column", df.columns)
    date_col = st.sidebar.selectbox("Date Column", df.columns)
    duration_col = st.sidebar.selectbox("Duration Column", df.columns)

    # -----------------------------
    # Select Customer
    # -----------------------------
    customers = df[customer_col].dropna().unique()
    selected_customer = st.selectbox("Select Customer", sorted(customers))

    # -----------------------------
    # Rate Input
    # -----------------------------
    st.sidebar.header("Rate Input")
    rate_per_hour = st.sidebar.number_input("Rate per Hour (£)", value=15.0)

    # -----------------------------
    # Filter Data
    # -----------------------------
    cust_df = df[df[customer_col] == selected_customer].copy()

    # -----------------------------
    # SAFE Duration Conversion (Fixes your error)
    # -----------------------------
    def convert_to_hours(x):
        try:
            # Case 1: numeric minutes
            return float(x) / 60
        except:
            try:
                # Case 2: time format HH:MM:SS
                t = pd.to_timedelta(x)
                return t.total_seconds() / 3600
            except:
                return 0

    cust_df["Hours"] = cust_df[duration_col].apply(convert_to_hours)

    # -----------------------------
    # Summary Calculation
    # -----------------------------
    summary = cust_df.groupby(employee_col).agg(
        Visits=(employee_col, "count"),
        Hours=("Hours", "sum")
    ).reset_index()

    summary["Rate"] = rate_per_hour
    summary["Cost"] = summary["Hours"] * rate_per_hour

    total_cost = summary["Cost"].sum()

    # -----------------------------
    # Display
    # -----------------------------
    st.subheader(f"Invoice Preview: {selected_customer}")
    st.dataframe(summary)

    st.success(f"Total Cost: £{round(total_cost, 2)}")

    st.subheader("Detailed Visits")
    st.dataframe(cust_df)

    # -----------------------------
    # Generate Formatted Excel Invoice
    # -----------------------------
    def generate_invoice(summary_df, detail_df):
        output = BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            sheet = workbook.add_worksheet("Invoice")
            writer.sheets["Invoice"] = sheet

            # Formats
            bold = workbook.add_format({'bold': True})
            title = workbook.add_format({'bold': True, 'font_size': 16})
            money = workbook.add_format({'num_format': '£#,##0.00'})

            # -----------------------------
            # Header
            # -----------------------------
            sheet.write("A1", "INVOICE", title)

            sheet.write("A3", "Customer:", bold)
            sheet.write("B3", selected_customer)

            sheet.write("A4", "Total Cost:", bold)
            sheet.write("B4", total_cost, money)

            # -----------------------------
            # Table Header
            # -----------------------------
            start_row = 6
            headers = ["Employee", "Visits", "Hours", "Rate (£)", "Cost (£)"]

            for col, h in enumerate(headers):
                sheet.write(start_row, col, h, bold)

            # -----------------------------
            # Table Data
            # -----------------------------
            for i, row in summary_df.iterrows():
                sheet.write(start_row + 1 + i, 0, row[employee_col])
                sheet.write(start_row + 1 + i, 1, row["Visits"])
                sheet.write(start_row + 1 + i, 2, row["Hours"])
                sheet.write(start_row + 1 + i, 3, row["Rate"])
                sheet.write(start_row + 1 + i, 4, row["Cost"], money)

            # -----------------------------
            # Grand Total
            # -----------------------------
            end_row = start_row + len(summary_df) + 2
            sheet.write(end_row, 3, "Grand Total", bold)
            sheet.write(end_row, 4, total_cost, money)

            # -----------------------------
            # Detailed Sheet
            # -----------------------------
            detail_df.to_excel(writer, sheet_name="Detailed Visits", index=False)

            # Adjust column width
            sheet.set_column("A:A", 25)
            sheet.set_column("B:E", 15)

        output.seek(0)
        return output

    # Generate file
    invoice_file = generate_invoice(summary, cust_df)

    # -----------------------------
    # Download Button
    # -----------------------------
    st.download_button(
        label="Download Invoice Excel",
        data=invoice_file,
        file_name=f"{selected_customer}_invoice.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
