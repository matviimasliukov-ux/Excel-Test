import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, numbers

st.set_page_config(page_title="Weekly Report Processor", layout="wide")

# -----------------------------
# State & helpers
# -----------------------------
def init_employees_state():
    if "employees" not in st.session_state:
        # Demo list — editable in the sidebar: rate is in PERCENT
        st.session_state.employees = [
            {"name": "John Doe", "rate_pct": 25.0, "truck": True,  "meter": False},
            {"name": "Jane Smith", "rate_pct": 30.0, "truck": False, "meter": True},
            {"name": "Alex Brown", "rate_pct": 28.0, "truck": True,  "meter": True},
        ]

def get_employee_by_name(name: str):
    for e in st.session_state.employees:
        if e["name"] == name:
            return e
    return None

def auto_col_width(ws):
    """Autosize columns based on max content length."""
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            v = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(v))
        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 60)

def format_amount_col(ws, amount_col_idx: int, bold_rows: set):
    """Right-align amount col and format as currency; bold for specified rows."""
    for col_cells in ws.iter_cols(min_col=amount_col_idx, max_col=amount_col_idx, min_row=2, max_row=ws.max_row):
        for c in col_cells:
            c.alignment = Alignment(horizontal="right")
            if isinstance(c.value, (int, float)):
                c.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            if c.row in bold_rows:
                c.font = Font(bold=True)

def export_per_tech_xlsx(df_tech: pd.DataFrame, tech_info: dict, date_col: str, tech_col: str, jobfee_col: str) -> bytes:
    """
    Build a workbook for a single technician:
    - Amount = Job Fee * (Rate% / 100)
    - Append charges: Truck ($50/day, cap $150), Meter ($25), Penguin Data ($6.25) in far-left column
    - Add bold Total row
    - Auto-size columns
    """
    d = df_tech.copy()

    # Compute Amount from Job Fee and Rate (%)
    d["Rate (%)"] = float(tech_info["rate_pct"])
    base = pd.to_numeric(d[jobfee_col], errors="coerce").fillna(0.0)
    d["Amount"] = base * (d["Rate (%)"] / 100.0)

    # Workbook & sheet
    wb = Workbook()
    ws = wb.active
    ws.title = tech_info["name"][:31]

    # Order columns: Date, Technician, all others..., Rate (%), Amount
    cols = list(d.columns)
    ordered = []
    if date_col in cols: ordered.append(date_col)
    if tech_col in cols and tech_col not in ordered: ordered.append(tech_col)
    for c in cols:
        if c not in {date_col, tech_col, "Rate (%)", "Amount"}:
            ordered.append(c)
    ordered.extend(["Rate (%)", "Amount"])
    d = d[ordered]

    # Write dataframe
    for r in dataframe_to_rows(d, index=False, header=True):
        ws.append(r)

    # Determine column positions
    last_col_idx = ws.max_column   # after writing, last column is "Amount"
    amount_col_idx = last_col_idx  # we placed Amount last

    # Append charges rows (names in far-left col; amounts in last col)
    bold_rows = set()
    charges_total = 0.0

    # Truck: $50 per working day from unique dates, cap $150
    if tech_info.get("truck"):
        unique_days = pd.to_datetime(d[date_col], errors="coerce").dt.date.dropna().unique()
        truck_days = len(unique_days)
        truck_fee = min(3, truck_days) * 50.0
        if truck_fee > 0:
            ws.append(["Truck Charge"] + [None]*(last_col_idx - 2) + [truck_fee])
            charges_total += truck_fee

    # Meter: fixed $25
    if tech_info.get("meter"):
        ws.append(["Meter Fee"] + [None]*(last_col_idx - 2) + [25.0])
        charges_total += 25.0

    # Penguin Data: fixed $6.25 (always)
    ws.append(["Penguin Data Fee"] + [None]*(last_col_idx - 2) + [6.25])
    charges_total += 6.25

    # Sum of Amount column for original rows (exclude charge rows)
    data_end_row = 1 + len(d)  # header + data
    data_amount_sum = 0.0
    for r in range(2, data_end_row + 1):
        val = ws.cell(row=r, column=amount_col_idx).value
        if isinstance(val, (int, float)):
            data_amount_sum += float(val)

    total = data_amount_sum + charges_total
    ws.append(["Total:"] + [None]*(last_col_idx - 2) + [total])
    total_row_idx = ws.max_row

    # Style header row bold
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")

    # Make "Total:" row bold (text + numbers)
    for cell in ws[total_row_idx]:
        cell.font = Font(bold=True)

    # Format amount column & right-align; bold total row
    bold_rows.add(total_row_idx)
    format_amount_col(ws, amount_col_idx, bold_rows)

    # Auto-size columns to avoid cropping
    auto_col_width(ws)

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# -----------------------------
# UI
# -----------------------------
init_employees_state()

st.title("Weekly Company Report → Per-Technician Split & Charges")
st.caption(
    "Upload a weekly .xlsx, map Date/Technician/Job Fee columns, then export per-tech files with Rate(%), "
    "Truck ($50/day, max $150), Meter ($25), Penguin Data ($6.25), and a bold Total row. Columns are auto-sized."
)

# Sidebar: employee manager (rate in %, Truck/Meter checkboxes)
with st.sidebar:
    st.header("Technicians (edit rate % + charges)")
    if st.session_state.employees:
        for e in st.session_state.employees:
            st.write(f"**{e['name']}** — Rate: {e['rate_pct']}% · Truck: {'✅' if e['truck'] else '—'} · Meter: {'✅' if e['meter'] else '—'}")

    st.divider()
    st.subheader("Add / Edit technician")
    with st.form("emp_form"):
        names = [e["name"] for e in st.session_state.employees]
        mode = st.radio("Mode", ["Add", "Edit"], horizontal=True)
        sel_name = st.selectbox("Select name (for Edit)", names) if mode == "Edit" else None
        name = st.text_input("Name", value=(sel_name or ""))
        default_rate = get_employee_by_name(sel_name)["rate_pct"] if sel_name else 25.0
        rate_pct = st.number_input("Rate (%)", min_value=0.0, max_value=1000.0, value=float(default_rate))
        default_truck = get_employee_by_name(sel_name)["truck"] if sel_name else False
        default_meter = get_employee_by_name(sel_name)["meter"] if sel_name else False
        truck = st.checkbox("Truck charge applies", value=bool(default_truck))
        meter = st.checkbox("Meter fee applies", value=bool(default_meter))
        submitted = st.form_submit_button("Save")
        if submitted and name.strip():
            if mode == "Add" and name not in names:
                st.session_state.employees.append({"name": name.strip(), "rate_pct": rate_pct, "truck": truck, "meter": meter})
                st.success(f"Added {name}")
            elif mode == "Edit" and sel_name in names:
                emp = get_employee_by_name(sel_name)
                emp.update({"name": name.strip(), "rate_pct": rate_pct, "truck": truck, "meter": meter})
                st.success(f"Updated {sel_name}")

# Main: file upload and processing
uploaded = st.file_uploader("Upload weekly .xlsx report", type=["xlsx"])

if uploaded:
    try:
        df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Could not read Excel: {e}")
        st.stop()

    st.subheader("Preview")
    st.dataframe(df.head(20), use_container_width=True)

    # Column mapping (auto-guess; allow override)
    cols = list(df.columns)
    if len(cols) < 3:
        st.error("The report should have at least three columns (Date, Technician, Job Fee).")
        st.stop()

    default_date = cols[0]
    # Try to guess tech column
    guess_tech = None
    for cand in ["Technician", "Tech", "Worker", "Employee", "Name"]:
        if cand in cols:
            guess_tech = cand
            break
    if guess_tech is None:
        guess_tech = cols[1]

    default_jobfee = cols[-1]

    st.markdown("#### Column mapping")
    c1, c2, c3 = st.columns(3)
    with c1:
        date_col = st.selectbox("Date column", cols, index=cols.index(default_date))
    with c2:
        tech_col = st.selectbox("Technician column", cols, index=cols.index(guess_tech))
    with c3:
        jobfee_col = st.selectbox("Job fee column (multiplied by Rate %)", cols, index=cols.index(default_jobfee))

    # Determine technicians we can process (present in both the file and our system list)
    techs_in_file = sorted(df[tech_col].dropna().astype(str).unique())
    system_names = [e["name"] for e in st.session_state.employees]
    matched = [t for t in techs_in_file if t in system_names]

    if not matched:
        st.warning("No matching technicians between the file and the system list. Add them in the sidebar or adjust names.")
        st.stop()

    st.markdown("#### Technicians detected & matched")
    st.write(", ".join(matched))

    # Generate per-tech files into a ZIP
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
        export_count = 0
        for t in matched:
            tech_rows = df[df[tech_col].astype(str) == t]
            if tech_rows.empty:
                continue
            info = get_employee_by_name(t)
            out_bytes = export_per_tech_xlsx(tech_rows, info, date_col, tech_col, jobfee_col)
            fname = f"{t.replace(' ', '_')}_{datetime.now().date()}.xlsx"
            z.writestr(fname, out_bytes)
            export_count += 1

    st.success(f"Prepared {export_count} file(s).")
    st.download_button(
        "Download ZIP with technician files",
        data=zbuf.getvalue(),
        file_name=f"technician_breakdowns_{datetime.now().date()}.zip",
        mime="application/zip",
    )

    st.info("Notes: Charge names are placed in the far-left column under the dated rows; amounts are in the last column. "
            "Columns are auto-sized; 'Total:' row is bold (label and numbers).")
else:
    st.info("Upload a weekly .xlsx report to begin.")
