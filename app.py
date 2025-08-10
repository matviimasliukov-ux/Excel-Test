import streamlit as st
import pandas as pd
import io
from datetime import date

# ----------------------
# Sample employee database (in-memory)
# ----------------------
employees = [
    {"name": "John Doe", "rate": 25, "equipment": "Drill"},
    {"name": "Jane Smith", "rate": 30, "equipment": "Lift"},
    {"name": "Alex Brown", "rate": 28, "equipment": "Saw"},
]

# ----------------------
# App title and layout
# ----------------------
st.title("Technician Weekly File Processor - Demo")

st.sidebar.header("Employees")
for emp in employees:
    st.sidebar.write(f"**{emp['name']}** - ${emp['rate']}/hr - {emp['equipment']}")

# Option to add a new employee
with st.sidebar.expander("Add Employee"):
    new_name = st.text_input("Name")
    new_rate = st.number_input("Rate", min_value=0, value=25)
    new_equipment = st.text_input("Equipment")
    if st.button("Add Employee"):
        if new_name.strip():
            employees.append({"name": new_name, "rate": new_rate, "equipment": new_equipment})
            st.success(f"Added {new_name}")

# ----------------------
# File upload and processing
# ----------------------

uploaded_file = st.file_uploader("Upload Weekly Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("### Uploaded Data Preview")
    st.dataframe(df.head())

    # Filter employees present in the file
    available_emps = [e for e in employees if e['name'] in df['Technician'].unique()]
    selected_emps = st.multiselect("Filter by Employee", [e['name'] for e in available_emps])

    if selected_emps:
        # Process for selected employees
        for emp_name in selected_emps:
            emp_data = df[df['Technician'] == emp_name].copy()
            emp_info = next(e for e in employees if e['name'] == emp_name)
            emp_data['Rate'] = emp_info['rate']
            emp_data['Equipment Charge'] = 50  # static for demo

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                emp_data.to_excel(writer, index=False)
            st.download_button(
                label=f"Download File for {emp_name}",
                data=output.getvalue(),
                file_name=f"{emp_name.replace(' ', '_')}_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
