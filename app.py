"""
Gulati Extrusions Pvt. Ltd. — Payroll Manager
Mobile-friendly Streamlit app with AI Vision attendance scanning.
"""

import streamlit as st
import calendar
import json
import re
from datetime import datetime
from io import BytesIO

import pandas as pd
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import (
    Font, Alignment, Border, Side, PatternFill, numbers
)
from openpyxl.utils import get_column_letter

# ── optional: google-generativeai (Gemini Vision) ──────────────────
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

# ═══════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Payroll Manager",
    page_icon="🏭",
    layout="wide",
)

# Mobile-friendly CSS
st.markdown("""
<style>
    /* tighter padding on mobile */
    .block-container {padding: 1rem 0.8rem !important;}
    /* full-width buttons */
    .stButton > button, .stDownloadButton > button {width:100%;}
    /* smaller dataframe font */
    .stDataFrame td, .stDataFrame th {font-size:0.82rem; padding:4px 6px !important;}
    /* metric cards */
    [data-testid="stMetric"] {background:#f8f9fb; border-radius:8px; padding:10px 14px;}
    /* tabs */
    .stTabs [data-baseweb="tab"] {font-size:0.85rem;}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# AUTO-DETECT MONTH  &  SIDEBAR
# ═══════════════════════════════════════════════════════════════════
now = datetime.now()

with st.sidebar:
    st.header("⚙️ Settings")

    sel_month = st.selectbox(
        "Month",
        range(1, 13),
        index=now.month - 1,
        format_func=lambda m: calendar.month_name[m],
    )
    sel_year = st.number_input("Year", 2020, 2035, now.year)
    total_days = calendar.monthrange(sel_year, sel_month)[1]

    st.metric("📅 Total Days", total_days)

    duty_hours = st.number_input(
        "Default Duty Hours", min_value=1, max_value=24, value=9
    )

    st.divider()
    st.subheader("🔑 AI Vision Key")
    # Try secrets first, then manual input
    api_key = st.secrets.get("GEMINI_API_KEY", "") if hasattr(st, "secrets") else ""
    if not api_key:
        api_key = st.text_input(
            "Gemini API Key",
            type="password",
            help="Free key → https://aistudio.google.com/apikey",
        )

    st.divider()
    st.caption("Gulati Extrusions Pvt. Ltd.")

month_name = calendar.month_name[sel_month]

# ═══════════════════════════════════════════════════════════════════
# TITLE
# ═══════════════════════════════════════════════════════════════════
st.title("🏭 Payroll Manager")
st.caption(f"**{month_name} {sel_year}** · {total_days} days · Duty {duty_hours} hrs")

# ═══════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════
if "employees" not in st.session_state:
    st.session_state.employees = pd.DataFrame(
        columns=["Name", "Rate", "Working_Days", "Advance", "OT_Hours"]
    )
if "scan_done" not in st.session_state:
    st.session_state.scan_done = False


# ═══════════════════════════════════════════════════════════════════
# HELPER — GEMINI VISION EXTRACTION
# ═══════════════════════════════════════════════════════════════════
def extract_attendance(images: list, api_key: str) -> list[dict]:
    """Send attendance sheet images to Gemini and get structured data."""
    if not GEMINI_AVAILABLE:
        st.error("google-generativeai package not installed.")
        return []

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")

    prompt = f"""You are analyzing handwritten factory attendance register images.

Extract the following for EACH employee row you can read:
1. "name"  — Employee name (as written, Hindi or English)
2. "working_days" — Count of 'P' (Present) marks. Count ONLY clear P marks.
3. "advance" — Advance amount if noted anywhere for that employee, else 0

IMPORTANT RULES:
- This month has {total_days} days. Working days cannot exceed {total_days}.
- If you see numbers in an 'Advance' / 'Peshgi' column, use them.
- Ignore header rows, totals rows, and empty rows.
- Return ONLY a valid JSON array — no markdown, no explanation.

Example output:
[
  {{"name": "Ram Kumar", "working_days": 26, "advance": 2000}},
  {{"name": "Suresh Yadav", "working_days": 24, "advance": 0}}
]"""

    parts = [prompt] + images  # text + PIL images

    try:
        response = model.generate_content(parts)
        raw = response.text.strip()
        # Clean markdown fences if present
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
        if isinstance(data, list):
            return data
        return []
    except json.JSONDecodeError:
        st.error("AI returned invalid JSON. Please check the image quality and try again.")
        st.code(raw[:500], language="json")
        return []
    except Exception as e:
        st.error(f"AI Vision error: {e}")
        return []


# ═══════════════════════════════════════════════════════════════════
# HELPER — SALARY CALCULATION
# ═══════════════════════════════════════════════════════════════════
def calc_salary(rate: float, working_days: float, total_month_days: int) -> tuple:
    """Returns (salary_type, earned_salary)."""
    if rate > 1500:
        # Monthly rate
        earned = round((rate / total_month_days) * working_days, 2)
        return ("Monthly", earned)
    else:
        # Daily rate
        earned = round(rate * working_days, 2)
        return ("Daily", earned)


def calc_ot(rate: float, ot_hours: float, total_month_days: int, duty_hrs: int) -> float:
    """Calculate overtime amount."""
    if ot_hours <= 0:
        return 0.0
    if rate > 1500:
        # Monthly → per-hour = rate / total_days / duty_hours
        per_hour = rate / total_month_days / duty_hrs
    else:
        # Daily → per-hour = rate / duty_hours
        per_hour = rate / duty_hrs
    return round(per_hour * ot_hours, 2)


# ═══════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════
tab_scan, tab_edit, tab_salary, tab_download = st.tabs(
    ["📷 Scan", "✏️ Edit Data", "💰 Salary", "📥 Excel"]
)

# ───────────────────────────────────────────────────────────────────
# TAB 1 — SCAN ATTENDANCE SHEET
# ───────────────────────────────────────────────────────────────────
with tab_scan:
    st.subheader("📷 Scan Attendance Sheet")

    uploaded = st.file_uploader(
        "Upload attendance sheet image(s)",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        help="Photo of handwritten attendance register",
    )

    if uploaded:
        cols = st.columns(min(len(uploaded), 3))
        pil_images = []
        for i, f in enumerate(uploaded):
            img = Image.open(f)
            pil_images.append(img)
            with cols[i % len(cols)]:
                st.image(img, caption=f.name, use_container_width=True)

        if not api_key:
            st.warning("⬅️ Enter your Gemini API key in the sidebar first.")
        elif st.button("🔍 Extract Attendance with AI", type="primary"):
            with st.spinner("AI is reading your attendance sheet..."):
                rows = extract_attendance(pil_images, api_key)

            if rows:
                new_df = pd.DataFrame(rows)
                # Ensure required columns
                for col in ["name", "working_days", "advance"]:
                    if col not in new_df.columns:
                        new_df[col] = 0

                new_df = new_df.rename(columns={
                    "name": "Name",
                    "working_days": "Working_Days",
                    "advance": "Advance",
                })
                new_df["Rate"] = 0.0
                new_df["OT_Hours"] = 0.0

                # Keep only needed columns in order
                new_df = new_df[["Name", "Rate", "Working_Days", "Advance", "OT_Hours"]]
                new_df["Working_Days"] = pd.to_numeric(new_df["Working_Days"], errors="coerce").fillna(0)
                new_df["Advance"] = pd.to_numeric(new_df["Advance"], errors="coerce").fillna(0)

                st.session_state.employees = new_df
                st.session_state.scan_done = True
                st.success(f"✅ {len(new_df)} employees extracted! Go to **Edit Data** tab to add Rates.")
            else:
                st.error("No data could be extracted. Try a clearer image.")
    else:
        st.info("Upload photo(s) of the handwritten attendance register to begin.")

# ───────────────────────────────────────────────────────────────────
# TAB 2 — EDIT DATA (manual entry + corrections)
# ───────────────────────────────────────────────────────────────────
with tab_edit:
    st.subheader("✏️ Employee Data")
    st.caption("Edit any cell directly. Add Rate for each employee.")

    # Add employee manually
    with st.expander("➕ Add Employee Manually"):
        c1, c2 = st.columns(2)
        new_name = c1.text_input("Name")
        new_rate = c2.number_input("Rate (₹)", min_value=0.0, value=0.0, step=50.0)
        c3, c4 = st.columns(2)
        new_wd = c3.number_input("Working Days", 0, total_days, 0)
        new_adv = c4.number_input("Advance (₹)", min_value=0.0, value=0.0, step=100.0)
        new_ot = st.number_input("OT Hours", min_value=0.0, value=0.0, step=1.0)

        if st.button("➕ Add", type="primary"):
            if new_name.strip():
                new_row = pd.DataFrame([{
                    "Name": new_name.strip(),
                    "Rate": new_rate,
                    "Working_Days": new_wd,
                    "Advance": new_adv,
                    "OT_Hours": new_ot,
                }])
                st.session_state.employees = pd.concat(
                    [st.session_state.employees, new_row], ignore_index=True
                )
                st.success(f"Added {new_name.strip()}")
                st.rerun()
            else:
                st.warning("Please enter a name.")

    # Editable table
    if len(st.session_state.employees) > 0:
        edited = st.data_editor(
            st.session_state.employees,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Name": st.column_config.TextColumn("Name", width="medium"),
                "Rate": st.column_config.NumberColumn("Rate ₹", min_value=0, format="%.0f"),
                "Working_Days": st.column_config.NumberColumn(
                    "Working Days", min_value=0, max_value=total_days, format="%d"
                ),
                "Advance": st.column_config.NumberColumn("Advance ₹", min_value=0, format="%.0f"),
                "OT_Hours": st.column_config.NumberColumn("OT Hours", min_value=0, format="%.1f"),
            },
            key="emp_editor",
        )
        st.session_state.employees = edited

        st.info(f"👥 {len(edited)} employees · Click any cell to edit")

        # Clear all
        if st.button("🗑️ Clear All Data", type="secondary"):
            st.session_state.employees = pd.DataFrame(
                columns=["Name", "Rate", "Working_Days", "Advance", "OT_Hours"]
            )
            st.session_state.scan_done = False
            st.rerun()
    else:
        st.info("No employees yet. Scan an attendance sheet or add manually above.")

# ───────────────────────────────────────────────────────────────────
# TAB 3 — SALARY CALCULATION
# ───────────────────────────────────────────────────────────────────
with tab_salary:
    st.subheader("💰 Salary Breakdown")

    df = st.session_state.employees.copy()
    if len(df) == 0:
        st.info("Add employee data first (Scan or Edit tab).")
    else:
        # Convert types
        for col in ["Rate", "Working_Days", "Advance", "OT_Hours"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # Calculate
        results = []
        for _, row in df.iterrows():
            stype, earned = calc_salary(row["Rate"], row["Working_Days"], total_days)
            ot_amt = calc_ot(row["Rate"], row["OT_Hours"], total_days, duty_hours)
            net = earned + ot_amt - row["Advance"]
            results.append({
                "Name": row["Name"],
                "Rate": row["Rate"],
                "Type": stype,
                "Working Days": int(row["Working_Days"]),
                "Earned (₹)": earned,
                "Advance (₹)": row["Advance"],
                "OT Hrs": row["OT_Hours"],
                "OT Amt (₹)": ot_amt,
                "Net Salary (₹)": round(net, 2),
            })

        salary_df = pd.DataFrame(results)

        # Summary metrics
        m1, m2, m3 = st.columns(3)
        m1.metric("👥 Employees", len(salary_df))
        m2.metric("💰 Total Net", f"₹{salary_df['Net Salary (₹)'].sum():,.0f}")
        m3.metric("📦 Total Advance", f"₹{salary_df['Advance (₹)'].sum():,.0f}")

        # Full table
        st.dataframe(
            salary_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Rate": st.column_config.NumberColumn(format="₹%.0f"),
                "Earned (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                "Advance (₹)": st.column_config.NumberColumn(format="₹%.0f"),
                "OT Amt (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                "Net Salary (₹)": st.column_config.NumberColumn(format="₹%.2f"),
            },
        )

        # Totals row
        st.markdown(
            f"**TOTAL →** Earned: **₹{salary_df['Earned (₹)'].sum():,.2f}** · "
            f"OT: **₹{salary_df['OT Amt (₹)'].sum():,.2f}** · "
            f"Advance: **₹{salary_df['Advance (₹)'].sum():,.0f}** · "
            f"Net: **₹{salary_df['Net Salary (₹)'].sum():,.2f}**"
        )

# ───────────────────────────────────────────────────────────────────
# TAB 4 — DOWNLOAD EXCEL
# ───────────────────────────────────────────────────────────────────
with tab_download:
    st.subheader("📥 Download Salary Sheet")

    df = st.session_state.employees.copy()
    if len(df) == 0:
        st.info("Add employee data first to generate Excel.")
    else:
        for col in ["Rate", "Working_Days", "Advance", "OT_Hours"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        if st.button("📊 Generate Excel", type="primary"):
            # ── Build Excel ──
            wb = Workbook()
            ws = wb.active
            ws.title = "Salary Sheet"

            # Styles
            header_font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
            sub_font = Font(name="Calibri", bold=True, size=11)
            col_header_font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
            data_font = Font(name="Calibri", size=10)
            total_font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")

            header_fill = PatternFill("solid", fgColor="1F4E79")
            col_fill = PatternFill("solid", fgColor="2E75B6")
            total_fill = PatternFill("solid", fgColor="1F4E79")
            alt_fill = PatternFill("solid", fgColor="D6E4F0")

            center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            left = Alignment(horizontal="left", vertical="center")
            right_align = Alignment(horizontal="right", vertical="center")

            thin_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

            # Column setup
            columns = [
                ("S.No", 6),
                ("Employee Name", 22),
                ("Rate (₹)", 12),
                ("Type", 10),
                (f"Total Days\n({month_name})", 11),
                ("Working\nDays", 10),
                ("Earned\nSalary (₹)", 14),
                ("Advance (₹)", 12),
                ("OT Hours", 10),
                ("OT Amt (₹)", 12),
                ("Net Salary (₹)", 15),
            ]

            # Set column widths
            for i, (_, w) in enumerate(columns, 1):
                ws.column_dimensions[get_column_letter(i)].width = w

            # Row 1-2: Company header (merged)
            ws.merge_cells("A1:K1")
            cell = ws["A1"]
            cell.value = "GULATI EXTRUSIONS PVT. LTD."
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center

            ws.merge_cells("A2:K2")
            cell = ws["A2"]
            cell.value = f"SALARY SHEET — {month_name.upper()} {sel_year}  |  Total Days: {total_days}  |  Duty: {duty_hours} hrs"
            cell.font = sub_font
            cell.fill = PatternFill("solid", fgColor="B4C6E7")
            cell.alignment = center

            ws.row_dimensions[1].height = 30
            ws.row_dimensions[2].height = 22

            # Row 3: Column headers
            for i, (title, _) in enumerate(columns, 1):
                cell = ws.cell(row=3, column=i, value=title)
                cell.font = col_header_font
                cell.fill = col_fill
                cell.alignment = center
                cell.border = thin_border
            ws.row_dimensions[3].height = 30

            # Data rows (starting row 4)
            start_row = 4
            for idx, (_, row) in enumerate(df.iterrows()):
                stype, earned = calc_salary(row["Rate"], row["Working_Days"], total_days)
                ot_amt = calc_ot(row["Rate"], row["OT_Hours"], total_days, duty_hours)
                net = earned + ot_amt - row["Advance"]

                r = start_row + idx
                values = [
                    idx + 1,
                    row["Name"],
                    row["Rate"],
                    stype,
                    total_days,
                    int(row["Working_Days"]),
                    earned,
                    row["Advance"],
                    row["OT_Hours"],
                    ot_amt,
                    round(net, 2),
                ]
                for c, val in enumerate(values, 1):
                    cell = ws.cell(row=r, column=c, value=val)
                    cell.font = data_font
                    cell.border = thin_border
                    if c == 2:
                        cell.alignment = left
                    elif c == 4:
                        cell.alignment = center
                    else:
                        cell.alignment = right_align

                    # Number format for currency columns
                    if c in (3, 7, 8, 10, 11):
                        cell.number_format = '#,##0.00'
                    elif c == 9:
                        cell.number_format = '0.0'

                    # Alternating row color
                    if idx % 2 == 1:
                        cell.fill = alt_fill

            # Totals row
            total_row = start_row + len(df)
            totals_data = {
                1: "TOTAL",
                7: sum(
                    calc_salary(r["Rate"], r["Working_Days"], total_days)[1]
                    for _, r in df.iterrows()
                ),
                8: df["Advance"].sum(),
                9: df["OT_Hours"].sum(),
                10: sum(
                    calc_ot(r["Rate"], r["OT_Hours"], total_days, duty_hours)
                    for _, r in df.iterrows()
                ),
            }
            # Net total
            total_earned = totals_data[7]
            total_ot_amt = totals_data[10]
            total_adv = totals_data[8]
            totals_data[11] = round(total_earned + total_ot_amt - total_adv, 2)

            for c in range(1, 12):
                cell = ws.cell(row=total_row, column=c)
                cell.font = total_font
                cell.fill = total_fill
                cell.border = thin_border
                cell.alignment = right_align if c != 1 else center
                if c in totals_data:
                    cell.value = totals_data[c]
                if c in (7, 8, 10, 11):
                    cell.number_format = '#,##0.00'
                elif c == 9:
                    cell.number_format = '0.0'

            ws.merge_cells(f"A{total_row}:B{total_row}")

            # Freeze panes
            ws.freeze_panes = "A4"

            # Print setup
            ws.print_title_rows = "1:3"
            ws.sheet_properties.pageSetUpPr.fitToPage = True

            # Save to buffer
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            filename = f"Salary_{month_name}_{sel_year}.xlsx"

            st.download_button(
                label=f"⬇️ Download {filename}",
                data=buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

            st.success(f"✅ Excel ready with {len(df)} employees!")


# ═══════════════════════════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════════════════════════
st.divider()
st.caption(
    "Gulati Extrusions Pvt. Ltd. · Payroll Manager · "
    f"Rate > ₹1500 = Monthly | Rate ≤ ₹1500 = Daily"
)
