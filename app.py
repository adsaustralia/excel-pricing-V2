
import streamlit as st
import openpyxl
import pandas as pd
import re
from io import BytesIO
from openpyxl.utils import column_index_from_string, get_column_letter

st.set_page_config(page_title="Excel SQM & Pricing Tool", layout="wide")
st.title("📊 Excel SQM & Pricing Calculator — Multi-Sheet Version")

st.write("""
Upload your Excel file, define column/row settings, and input pricing rules.  
The app will calculate **SQM and prices** for each sheet, display previews, and let you **download the updated Excel file or a combined summary**.
""")


# ---------------------------------------
# Utility: Normalize strings for matching
# ---------------------------------------
def normalize(s):
    return re.sub(r'[^a-z0-9]+', '', str(s).lower()) if s else ""


# ---------------------------------------
# MATERIAL CATEGORY DETECTION
# ---------------------------------------
MATERIAL_CATEGORY_MAP = {
    # PVC & Laminates
    "038mmpvcmattlaminate": "0.38mm PVC Matt Laminate",
    "2mmpvc": "2MM PVC",

    # Magnetic
    "06mmmagnetic": "0.6mm Magnetic",
    "magnetic": "0.6mm Magnetic",

    # Gloss Paper
    "350gsmgloss": "350 GSM Gloss",
    "350gsm": "350 GSM Gloss",
}


# ---------------------------------------
# PER-SIDE MULTIPLIERS (ONLY RULES!)
# Prices will be entered per client.
# ---------------------------------------
SIDE_MULTIPLIERS = {
    "0.6mm Magnetic": {
        "Single sided": 1.0,
        "Double sided": 1.3,
    },
    "350 GSM Gloss": {
        "Single sided": 1.0,
        "Double sided": 1.2,
    },
    "0.38mm PVC Matt Laminate": {
        "Single sided": 1.0,
        "Double sided": 1.3,
    },
    "2MM PVC": {
        "Single sided": 1.0,
        "Double sided": 1.3,
    },
}


# Detect category from raw material cell text
def detect_category(material: str):
    if not material:
        return None
    m = normalize(material)
    for needle, category in MATERIAL_CATEGORY_MAP.items():
        if needle in m:
            return category
    return None  # fallback (material name itself will be category)


# ---------------------------------------
# DATA CLEANING HELPERS
# ---------------------------------------
def parse_size(raw):
    if not raw:
        return None, None
    s = str(raw).replace("×", "x").replace("X", "x").replace("*", "x")
    nums = re.findall(r'\d+(?:\.\d+)?', s)
    if len(nums) >= 2:
        return float(nums[0]), float(nums[1])
    return None, None


def parse_qty(raw):
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)

    m = re.search(r'\d+(?:\.\d+)?', str(raw).replace(",", ""))
    if m:
        return float(m.group(0))

    return None


def clean_value(v):
    if v is None:
        return None
    if isinstance(v, str) and v.strip().startswith("="):
        return None
    return v


# ---------------------------------------
# MAIN APP LOGIC
# ---------------------------------------
uploaded_file = st.file_uploader("📤 Upload Excel file", type=["xlsx"])

if uploaded_file:

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet_names = wb.sheetnames

    # ---------------------------------------
    # Excel structure settings
    # ---------------------------------------
    st.subheader("⚙️ Excel Structure Settings")
    col1, col2 = st.columns(2)

    with col1:
        start_col = st.text_input("Start Column", value="AC")
        end_col = st.text_input("End Column", value="IG")

    with col2:
        row_size = st.number_input("Row (Size)", value=5)
        row_material = st.number_input("Row (Material)", value=6)
        row_qty = st.number_input("Row (Quantity)", value=155)
        row_sqm = st.number_input("Row (SQM Output)", value=156)
        row_price = st.number_input("Row (Price Output)", value=157)

    # ---------------------------------------
    # Single-sided / Double-sided
    # ---------------------------------------
    st.subheader("🧮 Print Side")
    side_option = st.radio(
        "Select print side:",
        options=["Single sided", "Double sided"],
        index=0,
        horizontal=True
    )

    # ---------------------------------------
    # Sheet selector
    # ---------------------------------------
    st.subheader("📄 Sheet Selection")
    process_all = st.checkbox("Process all sheets", value=True)

    sheet_choice = None
    if not process_all:
        sheet_choice = st.selectbox("Select sheet", sheet_names)

    start_idx = column_index_from_string(start_col)
    end_idx = column_index_from_string(end_col)

    # ---------------------------------------
    # Detect materials in Excel
    # ---------------------------------------
    st.subheader("🧾 Materials Detected in Excel")

    detected_materials = set()
    source_sheets = sheet_names if process_all else [sheet_choice]

    for sname in source_sheets:
        ws = wb[sname]
        for c in range(start_idx, end_idx + 1):
            col = get_column_letter(c)
            raw = clean_value(ws[f"{col}{row_material}"].value)
            if raw:
                detected_materials.add(str(raw).strip())

    if not detected_materials:
        st.error("❌ No materials found. Check row/column settings.")
        st.stop()

    # Determine category list
    categories_present = set()
    for mat in detected_materials:
        cat = detect_category(mat)
        if cat is None:
            cat = mat  # fallback use as-is
        categories_present.add(cat)

    st.write("**Detected categories:**")
    for c in sorted(categories_present):
        st.write(f"- {c}")

    st.markdown("---")

    # ---------------------------------------
    # Per-client price input
    # ---------------------------------------
    st.subheader("💰 Enter base SINGLE-SIDED rate per category (per client)")

    base_rates = {}
    for cat in sorted(categories_present):
        base_rates[cat] = st.number_input(
            f"Base rate for '{cat}' (Single sided, AUD/m²):",
            min_value=0.0,
            value=0.0,
            step=0.1,
            key=f"base_{normalize(cat)}"
        )

    # Effective rate calculator
    def get_effective_rate(material, side):
        cat = detect_category(material) or material
        base = base_rates.get(cat, 0)
        mult = SIDE_MULTIPLIERS.get(cat, {}).get(side, 1.0)
        return base * mult if base else None

    # ---------------------------------------
    # PROCESS BUTTON
    # ---------------------------------------
    if st.button("🚀 Process & Calculate"):
        summary = []

        def process_sheet(ws, sheetname):
            rows = []
            total_cost = 0

            for c in range(start_idx, end_idx + 1):
                col = get_column_letter(c)

                raw_size = clean_value(ws[f"{col}{row_size}"].value)
                raw_mat = clean_value(ws[f"{col}{row_material}"].value)
                raw_qty = clean_value(ws[f"{col}{row_qty}"].value)

                w, h = parse_size(raw_size)
                qty = parse_qty(raw_qty)
                rate = get_effective_rate(raw_mat, side_option)

                sqm = price = None
                if w and h and qty:
                    sqm = (w / 1000) * (h / 1000) * qty
                    if rate:
                        price = round(sqm * rate, 2)
                        total_cost += price

                    ws[f"{col}{row_sqm}"].value = sqm
                    ws[f"{col}{row_price}"].value = price

                rows.append({
                    "Column": col,
                    "Material": raw_mat,
                    "Category": detect_category(raw_mat) or raw_mat,
                    "Size": raw_size,
                    "Qty": qty,
                    "Rate": rate,
                    "SQM": sqm,
                    "Price": price
                })

            # Write total to Excel
            ws[f"{end_col}{row_sqm}"] = "TOTAL"
            ws[f"{end_col}{row_price}"] = total_cost
            return pd.DataFrame(rows), total_cost

        # Process
        if process_all:
            for sname in sheet_names:
                df, total = process_sheet(wb[sname], sname)
                st.markdown(f"### 📄 {sname}")
                st.dataframe(df)
                st.success(f"Subtotal: ${total:,.2f}")
                summary.append({"Sheet": sname, "Total": total})
        else:
            df, total = process_sheet(wb[sheet_choice], sheet_choice)
            st.markdown(f"### 📄 {sheet_choice}")
            st.dataframe(df)
            st.success(f"Total: ${total:,.2f}")
            summary.append({"Sheet": sheet_choice, "Total": total})

        # Combined summary
        st.subheader("📘 Combined Totals")
        summary_df = pd.DataFrame(summary)
        st.dataframe(summary_df)
        st.success(f"GRAND TOTAL: ${summary_df['Total'].sum():,.2f}")

        # Download
        bytes_xlsx = BytesIO()
        wb.save(bytes_xlsx)
        bytes_xlsx.seek(0)

        st.download_button(
            "⬇️ Download Updated Excel",
            bytes_xlsx,
            "Updated_Pricing.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
