import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from datetime import datetime
import shutil
import os

st.title("📝 Daily Visit Report Generator")

# Load source file
df = pd.read_excel("Data base.xlsx")

visits = st.number_input("Number of visits", min_value=1, step=1)

filled_data = []

for i in range(visits):
    st.subheader(f"Visit {i + 1}")

    is_office = st.checkbox(f"Is Visit {i + 1} an Office Visit?", key=f"office_{i}")

    if is_office:
        # Office visit: only "Office" as site name, everything else is blank
        filled_data.append([
            datetime.now().strftime('%Y-%m-%d'),  # Date
            "Office",                             # Site Name
            "", "", "", "", "",                   # Gov, Address, Contact Name, No., Type
            "", "", "",                           # Task Status, Task Type, Model
            "", "",                               # SN, Work
            "",                                   # Travel
            "",                                   # Tech Report
            ""                                    # Case No.
        ])
    else:
        serial_input = st.text_input(f"Serial Number {i + 1}", key=f"serial_{i}")
        row = df[df["Serial Number"].astype(str) == serial_input]

        if not row.empty:
            row = row.iloc[0]
            st.markdown(f"**Customer:** {row['Customer Name']}")
            st.markdown(f"**Governorate:** {row['Governorate']}")
            st.markdown(f"**Address:** {row['Address']}")
            st.markdown(f"**Model:** {row['Model']}")

            task_type = st.text_input(f"Task Type {i+1}", key=f"task_{i}")
            task_status = st.selectbox(f"Task Status {i+1}", ["Complete", "NOT Complete"], key=f"status_{i}")
            visit_type = st.selectbox(f"Type {i+1}", ["PPM", "Service"], key=f"type_{i}")
            work = st.text_area(f"Work Done {i+1}", key=f"work_{i}")
            tech_report = st.text_input(f"Technical Report No. {i+1}", key=f"report_{i}")

            filled_data.append([
                datetime.now().strftime('%Y-%m-%d'),
                row["Customer Name"],
                row["Governorate"],
                row["Address"],
                row["Contact Person 1"],
                row["Contact Number 1"],
                visit_type,
                task_status,
                task_type,
                row["Model"],
                serial_input,
                work,
                "",
                tech_report,
                ""
            ])
        else:
            st.error("❌ Serial not found.")


# When ready to export
if st.button("Generate Excel Report"):
    if filled_data:
        template = "Daily Report Form.xlsx"
        output = f"filled_visits_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        shutil.copy(template, output)
        wb = load_workbook(output)
        ws = wb.active

        for i, row_data in enumerate(filled_data, start=2):
            for j, val in enumerate(row_data, start=2):
                cell = ws.cell(row=i, column=j, value=val)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(bold=True)

        wb.save(output)
        st.success("✅ File generated!")
        with open(output, "rb") as f:
            st.download_button("📥 Download Report", f, file_name=output)
    else:
        st.warning("No data to export.")
