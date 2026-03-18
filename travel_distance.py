import streamlit as st
import pandas as pd
from io import BytesIO

st.title("EMS Invoice Generator")

uploaded_file = st.file_uploader("Upload EMS File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # -----------------------------
    # Column Mapping
    # -----------------------------
    st.sidebar.header("Column Mapping")

    customer_col = st.sidebar.selectbox("Customer", df.columns)
    employee_col = st.sidebar.selectbox("Employee", df.columns)
    date_col = st.sidebar.selectbox("Date", df.columns)
    duration_col = st.sidebar.selectbox("Duration (Minutes)", df.columns)

    # -----------------------------
    # Select Customer
    # -----------------------------
    customers = df[customer_col].dropna().unique()
    selected_customer = st.selectbox("Select Customer", sorted(customers))

    # -----------------------------
    # Inputs
    # -----------------------------
    rate = st.number_input("Rate per Hour (£)", value=15.0)
    travel_rate = st.number_input("Travel Rate per KM (£)", value=0.5)

    # -----------------------------
    # Filter Data
    # -----------------------------
    cust_df = df[df[customer_col] == selected_customer].copy()
    cust_df["Hours"] = cust_df[duration_col] / 60

    # -----------------------------
    # Summary
    # -----------------------------
    summary = cust_df.groupby(employee_col).agg(
        Visits=(employee_col, "count"),
        Hours=("Hours", "sum")
    ).reset_index()

    summary["Rate"] = rate
    summary["Cost"] = summary["Hours"] * rate

    total_cost = summary["Cost"].sum()

    st.subheader("Invoice Preview")
    st.dataframe(summary)
    st.success(f"Total Cost: £{round(total_cost, 2)}")

    # -----------------------------
    # Generate Formatted Excel
    # -----------------------------
    def generate_invoice():
        output = BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            sheet = workbook.add_worksheet("Invoice")
            writer.sheets["Invoice"] = sheet

            # Formats
            bold = workbook.add_format({'bold': True})
            title = workbook.add_format({'bold': True, 'font_size': 16})
            money = workbook.add_format({'num_format': '£#,##0.00'})
            border = workbook.add_format({'border': 1})

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
            for i, row in summary.iterrows():
                sheet.write(start_row + 1 + i, 0, row[employee_col])
                sheet.write(start_row + 1 + i, 1, row["Visits"])
                sheet.write(start_row + 1 + i, 2, row["Hours"])
                sheet.write(start_row + 1 + i, 3, row["Rate"])
                sheet.write(start_row + 1 + i, 4, row["Cost"], money)

            # -----------------------------
            # Grand Total
            # -----------------------------
            end_row = start_row + len(summary) + 2
            sheet.write(end_row, 3, "Grand Total", bold)
            sheet.write(end_row, 4, total_cost, money)

            # Auto width
            sheet.set_column("A:A", 25)
            sheet.set_column("B:E", 15)

        output.seek(0)
        return output

    invoice_file = generate_invoice()

    st.download_button(
        label="Download Invoice Excel",
        data=invoice_file,
        file_name=f"{selected_customer}_invoice.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )