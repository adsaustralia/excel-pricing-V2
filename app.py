
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
The app will calculate **SQM and prices** for each sheet, display previews, and let you **download results or a combined summary**.
""")

def normalize(s):
    return re.sub(r'[^a-z0-9]+', '', str(s).lower()) if s else ""

def base_get_rate(material, prices):
    if material is None:
        return None
    m = normalize(material)
    if "jellyfish" in m:
        return prices["Jellyfish"]
    if "ferrous" in m:
        return prices["Ferrous"]
    if "syntheticbanner" in m:
        ds_keywords = ["ds","doublesided","2sided","twosided","railtop","pocket","dowel"]
        if any(k in m for k in ds_keywords):
            return prices["Synthetic Banner (DS)"]
        return prices["Synthetic Banner (SS)"]
    if "backlitshimmer" in m:
        return prices["Backlit Shimmer"]
    if "artboard" in m or "bflute" in m or "400gsm" in m:
        return prices["Artboard"]
    return None

def parse_size(raw):
    if not raw:
        return None, None
    s = str(raw).replace("×","x").replace("X","x").replace("*","x")
    nums = re.findall(r'\d+(?:\.\d+)?', s)
    if len(nums) >= 2:
        return float(nums[0]), float(nums[1])
    return None, None

def parse_qty(raw):
    if raw is None:
        return None
    if isinstance(raw,(int,float)):
        return float(raw)
    m = re.search(r'\d+(?:\.\d+)?', str(raw).replace(",",""))
    return float(m.group(0)) if m else None

def clean_value(v):
    if v is None:
        return None
    if isinstance(v,str) and v.strip().startswith("="):
        return None
    return v

uploaded_file = st.file_uploader("📤 Upload Excel file", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet_names = wb.sheetnames

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

    st.subheader("💰 Base Material Pricing Rules (AUD/m²)")
    col1,col2,col3 = st.columns(3)
    with col1:
        price_jellyfish = st.number_input("Jellyfish", value=7.55)
        price_ferrous = st.number_input("Ferrous", value=12.00)
    with col2:
        price_artboard = st.number_input("Artboard / B-Flute / 400gsm", value=6.42)
        price_ss = st.number_input("Synthetic Banner (SS)", value=13.00)
    with col3:
        price_ds = st.number_input("Synthetic Banner (DS)", value=19.50)
        price_backlit = st.number_input("Backlit Shimmer", value=40.00)

    prices = {
        "Jellyfish": price_jellyfish,
        "Ferrous": price_ferrous,
        "Artboard": price_artboard,
        "Synthetic Banner (SS)": price_ss,
        "Synthetic Banner (DS)": price_ds,
        "Backlit Shimmer": price_backlit,
    }

    st.subheader("📄 Sheet Selection")
    process_all = st.checkbox("Process all sheets", value=True)
    sheet_choice = None
    if not process_all:
        sheet_choice = st.selectbox("Select a sheet", sheet_names)

    start_idx = column_index_from_string(start_col)
    end_idx = column_index_from_string(end_col)

    st.subheader("🧾 Detected Materials & Custom Rates (for new ones)")
    detected_materials = set()

    sheets_to_scan = sheet_names if process_all else [sheet_choice]
    for sheet_name in sheets_to_scan:
        ws_scan = wb[sheet_name]
        for c in range(start_idx, end_idx+1):
            col_letter = get_column_letter(c)
            raw_mat = clean_value(ws_scan[f"{col_letter}{row_material}"].value)
            if raw_mat:
                detected_materials.add(str(raw_mat).strip())

    extra_rates={}
    for mat in sorted(detected_materials):
        if base_get_rate(mat, prices) is None:
            extra_rates[mat] = st.number_input(
                f"Rate for NEW material: '{mat}' (AUD/m²)", min_value=0.0, value=0.0
            )

    def get_effective_rate(material):
        r = base_get_rate(material, prices)
        if r is not None:
            return r
        if material is None:
            return None
        return extra_rates.get(str(material).strip())

    if st.button("🚀 Process & Calculate"):
        summary_data=[]

        def process_sheet(ws, sheet_name):
            total=0
            rows=[]
            for c in range(start_idx, end_idx+1):
                col = get_column_letter(c)
                raw_size = clean_value(ws[f"{col}{row_size}"].value)
                raw_mat = clean_value(ws[f"{col}{row_material}"].value)
                raw_qty = clean_value(ws[f"{col}{row_qty}"].value)

                w,h = parse_size(raw_size)
                qty = parse_qty(raw_qty)
                rate = get_effective_rate(raw_mat)
                sqm=price=None

                if w and h and qty:
                    sqm=(w/1000)*(h/1000)*qty
                    if rate is not None:
                        price=round(sqm*rate,2)
                        total += price
                    ws[f"{col}{row_sqm}"].value = sqm
                    ws[f"{col}{row_price}"].value = price

                rows.append({
                    "Column":col,
                    "Material":raw_mat,
                    "Size":raw_size,
                    "Qty":qty,
                    "Rate":rate,
                    "SQM":sqm,
                    "Price (AUD)":price
                })

            ws[f"{end_col}{row_sqm}"]="TOTAL"
            ws[f"{end_col}{row_price}"]=total
            return pd.DataFrame(rows), total

        if process_all:
            for name in sheet_names:
                ws=wb[name]
                df,total = process_sheet(ws,name)
                st.markdown(f"### 🧾 Preview — {name}")
                st.dataframe(df)
                st.info(f"Subtotal for {name}: ${total:,.2f}")
                summary_data.append({"Sheet":name, "Total (AUD)":total})
        else:
            ws=wb[sheet_choice]
            df,total = process_sheet(ws,sheet_choice)
            st.markdown(f"### 🧾 Preview — {sheet_choice}")
            st.dataframe(df)
            st.info(f"Total for {sheet_choice}: ${total:,.2f}")
            summary_data.append({"Sheet":sheet_choice, "Total (AUD)":total})

        if summary_data:
            summary_df=pd.DataFrame(summary_data)
            st.subheader("📘 Combined Totals Summary")
            st.dataframe(summary_df)
            grand = summary_df["Total (AUD)"].sum()
            st.success(f"✅ Grand Total: ${grand:,.2f}")

        excel_bytes=BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        st.download_button(
            "⬇️ Download Updated Excel Workbook",
            excel_bytes,
            "Updated_Pricing_All_Sheets.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
