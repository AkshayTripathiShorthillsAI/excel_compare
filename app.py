import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tempfile
from io import BytesIO

st.title("📊 Excel Difference Highlighter")

raw_file = st.file_uploader("Upload Raw Excel", type=["xlsx"])
changed_file = st.file_uploader("Upload Updated Excel", type=["xlsx"])

if raw_file and changed_file:

    # Load workbooks
    raw_wb = load_workbook(BytesIO(raw_file.read()), data_only=True)
    changed_wb = load_workbook(BytesIO(changed_file.read()))

    raw_sheets = set(raw_wb.sheetnames)
    changed_sheets = set(changed_wb.sheetnames)

    common_sheets = raw_sheets.intersection(changed_sheets)

    if not common_sheets:
        st.error("No matching sheets found.")
        st.stop()

    highlight_fill = PatternFill(
        start_color="FFFF00",
        end_color="FFFF00",
        fill_type="solid"
    )

    for sheet_name in common_sheets:

        st.write(f"Processing: {sheet_name}")

        raw_ws = raw_wb[sheet_name]
        changed_ws = changed_wb[sheet_name]

        max_row = changed_ws.max_row
        max_col = changed_ws.max_column

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):

                new_cell = changed_ws.cell(row=row, column=col)
                old_cell = raw_ws.cell(row=row, column=col)

                new_val = new_cell.value
                old_val = old_cell.value

                # compare safely
                if str(new_val) != str(old_val):
                    new_cell.fill = highlight_fill

    # Save highlighted UPDATED workbook
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        changed_wb.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.download_button(
                "⬇️ Download Highlighted Excel",
                data=f,
                file_name="highlighted_updated.xlsx"
            )
