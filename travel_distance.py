import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("EMS Customer Invoice Generator")

uploaded_file = st.file_uploader("Upload EMS Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Raw Data")
    st.dataframe(df.head())

    # -----------------------------
    # Column Mapping
    # -----------------------------
    st.sidebar.header("Column Mapping")

    customer_col = st.sidebar.selectbox("Customer Column", df.columns)
    employee_col = st.sidebar.selectbox("Employee Column", df.columns)
    duration_col = st.sidebar.selectbox("Duration Column", df.columns)

    # -----------------------------
    # Select Customer
    # -----------------------------
    customers = df[customer_col].dropna().unique()
    selected_customer = st.selectbox("Select Customer", sorted(customers))

    cust_df = df[df[customer_col] == selected_customer].copy()

    # -----------------------------
    # Convert Duration → Hours
    # -----------------------------
    def convert_to_hours(x):
        if pd.isna(x):
            return 0

        x = str(x).strip()

        if x.replace('.', '', 1).isdigit():
            return float(x) / 60

        if "min" in x.lower():
            num = ''.join(filter(str.isdigit, x))
            return float(num) / 60 if num else 0

        try:
            t = pd.to_timedelta(x)
            return t.total_seconds() / 3600
        except:
            return 0

    cust_df["Hours"] = cust_df[duration_col].apply(convert_to_hours)

    # -----------------------------
    # Group by Employee
    # -----------------------------
    summary = cust_df.groupby(employee_col).agg(
        Visits=(employee_col, "count"),
        Hours=("Hours", "sum")
    ).reset_index()

    st.subheader(f"Employees serving {selected_customer}")
    st.dataframe(summary)

    # -----------------------------
    # USER INPUT PER EMPLOYEE
    # -----------------------------
    st.subheader("Enter Rate & Distance for Each Employee")

    employee_inputs = {}

    for i, row in summary.iterrows():
        emp = row[employee_col]

        col1, col2 = st.columns(2)

        with col1:
            rate = st.number_input(
                f"Rate (£/hr) - {emp}",
                min_value=0.0,
                value=15.0,
                key=f"rate_{i}"
            )

        with col2:
            distance = st.number_input(
                f"Distance (KM) - {emp}",
                min_value=0.0,
                value=0.0,
                key=f"dist_{i}"
            )

        employee_inputs[emp] = {
            "rate": rate,
            "distance": distance
        }

    # -----------------------------
    # Calculate Costs
    # -----------------------------
    travel_rate = st.number_input("Travel Rate per KM (£)", value=0.5)

    costs = []

    for i, row in summary.iterrows():
        emp = row[employee_col]
        hours = row["Hours"]

        rate = employee_inputs[emp]["rate"]
        distance = employee_inputs[emp]["distance"]

        care_cost = hours * rate
        travel_cost = distance * travel_rate
        total_cost = care_cost + travel_cost

        costs.append([
            emp,
            row["Visits"],
            hours,
            rate,
            distance,
            total_cost
        ])

    result_df = pd.DataFrame(costs, columns=[
        "Employee", "Visits", "Hours", "Rate", "Distance", "Total Cost"
    ])

    st.subheader("Final Cost Table")
    st.dataframe(result_df)

    grand_total = result_df["Total Cost"].sum()
    st.success(f"Grand Total: £{round(grand_total, 2)}")

    # -----------------------------
    # Excel Download
    # -----------------------------
    def generate_invoice(df):
        output = BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name="Invoice", index=False)

        output.seek(0)
        return output

    invoice_file = generate_invoice(result_df)

    st.download_button(
        "Download Invoice",
        data=invoice_file,
        file_name=f"{selected_customer}_invoice.xlsx"
    )
