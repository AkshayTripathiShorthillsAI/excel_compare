import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import tempfile

st.title("Excel Difference Highlighter")

raw_file = st.file_uploader("Upload Raw Excel", type=["xlsx"])
changed_file = st.file_uploader("Upload Changed Excel", type=["xlsx"])

if raw_file and changed_file:

    raw_excel = pd.ExcelFile(raw_file)
    changed_excel = pd.ExcelFile(changed_file)

    output_wb = Workbook()
    output_wb.remove(output_wb.active)

    highlight_fill = PatternFill(start_color="FFFF00",
                                 end_color="FFFF00",
                                 fill_type="solid")

    for sheet_name in changed_excel.sheet_names:

        raw_df = pd.read_excel(raw_excel, sheet_name=sheet_name)
        changed_df = pd.read_excel(changed_excel, sheet_name=sheet_name)

        raw_df = raw_df.reindex_like(changed_df)

        ws = output_wb.create_sheet(title=sheet_name)

        for col_idx, column in enumerate(changed_df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=column)

        for row_idx in range(len(changed_df)):
            for col_idx in range(len(changed_df.columns)):
                new_val = changed_df.iloc[row_idx, col_idx]
                old_val = raw_df.iloc[row_idx, col_idx]

                cell = ws.cell(row=row_idx+2, column=col_idx+1, value=new_val)

                if str(new_val) != str(old_val):
                    cell.fill = highlight_fill

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        output_wb.save(tmp.name)
        tmp.seek(0)
        st.download_button(
            "Download Highlighted Excel",
            data=tmp,
            file_name="diff_output.xlsx"
        )
