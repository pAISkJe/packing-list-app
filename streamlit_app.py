"""
Packing List generator -- Streamlit web app.

Upload a parcel master workbook (.xlsx), pick the metal sheet, and download a
ready-made "<Metal> Packing List" workbook (computed columns, photos, sorting).

Deploy on Streamlit Community Cloud:
  * main file path : streamlit_app.py
  * see README.md for full steps
"""

import streamlit as st

from packing_core import (
    list_sheets,
    suggest_purity_unit,
    generate_packing_list,
)

st.set_page_config(page_title="Packing List Generator", page_icon="💎", layout="centered")

st.title("💎 Packing List Generator")
st.caption(
    "Upload a parcel master workbook, choose a metal sheet, and download the "
    "packing list. The USD columns stay blank until you type the USD-to-INR "
    "rate into cell **E1** of the generated sheet — Excel fills them in on open."
)

uploaded = st.file_uploader("Parcel master workbook (.xlsx)", type=["xlsx"])

if not uploaded:
    st.info("Upload an .xlsx file to begin.")
    st.stop()

file_bytes = uploaded.getvalue()

# Read sheet names (guard against unreadable files)
try:
    sheets = list_sheets(file_bytes)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not read that file: {exc}")
    st.stop()

if not sheets:
    st.error("No sheets found in the workbook.")
    st.stop()

st.subheader("Options")

col1, col2 = st.columns(2)
with col1:
    sheet = st.selectbox("Metal sheet", sheets, index=0)
with col2:
    metal = st.text_input("Metal label (used in Item text / titles)", value=sheet)

col3, col4 = st.columns(2)
with col3:
    unit_choice = st.selectbox(
        "Purity unit in Item text",
        ["Auto", "Kt (×24)", "Fineness (×1000)", "None"],
        index=0,
        help="Auto = Kt for gold, fineness for silver, none otherwise.",
    )
with col4:
    markup_pct = st.number_input(
        "Default mark-up on stone price (%)",
        min_value=0.0, max_value=100.0, value=4.0, step=0.5,
    )

unit_map = {"Auto": None, "Kt (×24)": "kt", "Fineness (×1000)": "fineness", "None": "none"}
purity_unit = unit_map[unit_choice]
if purity_unit is None:
    st.caption(f"Auto purity unit for **{metal}** → `{suggest_purity_unit(metal)}`")

if st.button("Generate packing list", type="primary"):
    try:
        with st.spinner("Building packing list…"):
            out_bytes, info = generate_packing_list(
                file_bytes,
                sheet=sheet,
                metal=metal.strip() or sheet,
                purity_unit=purity_unit,
                markup=markup_pct / 100.0,
            )
    except Exception as exc:  # noqa: BLE001
        st.error(f"Generation failed: {exc}")
        st.stop()

    base = uploaded.name.rsplit(".", 1)[0]
    out_name = f"{base} {info['metal']} Packing List.xlsx"

    st.success(
        f"Done — {info['items']} item(s) from **{info['sheet']}** "
        f"(purity unit: `{info['purity_unit']}`, mark-up: {info['markup']:.0%})."
    )
    st.download_button(
        "⬇️ Download packing list",
        data=out_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption(
        "Reminder: open the file and enter the USD-to-INR rate in cell **E1** "
        "to populate the USD columns."
    )

with st.expander("Notes & limits"):
    st.markdown(
        "- Columns are matched by **header text**, so Gold / Silver / Brass all work "
        "despite their shifted layouts.\n"
        "- The generated file has two sheets: the **packing list** and a copy of the "
        "**source sheet** (photos preserved).\n"
        "- The app reads cached cell values — if you just edited formulas in the master, "
        "open & save it once in Excel first so values are current.\n"
        "- Brass works too, but since it has no purity / fine-metal columns those fields "
        "are left out of its Item text."
    )
