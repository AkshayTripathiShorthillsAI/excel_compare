import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import tempfile

st.title("🚀 Smart Excel Diff Tool")

raw_file = st.file_uploader("Upload Raw Excel", type=["xlsx"])
changed_file = st.file_uploader("Upload Updated Excel", type=["xlsx"])


# ---------- NORMALIZER ----------
def normalize(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return round(v, 6)
    return str(v).strip()


# ---------- COLORS ----------
MODIFIED = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
ADDED = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
DELETED = PatternFill(start_color="FF7F7F", end_color="FF7F7F", fill_type="solid")


if raw_file and changed_file:

    raw_wb = load_workbook(BytesIO(raw_file.read()), data_only=True)

    changed_wb = load_workbook(BytesIO(changed_file.read()))
    changed_val_wb = load_workbook(
        BytesIO(changed_file.getvalue()),
        data_only=True
    )

    summary = []

    common_sheets = (
        set(raw_wb.sheetnames)
        & set(changed_wb.sheetnames)
    )

    for sheet in common_sheets:

        st.write(f"Processing {sheet}")

        raw_ws = raw_wb[sheet]
        new_ws = changed_wb[sheet]
        new_ws_val = changed_val_wb[sheet]

        raw_rows = raw_ws.max_row
        new_rows = new_ws.max_row
        max_cols = new_ws.max_column

        modified_count = 0
        added_count = 0
        deleted_count = 0

        max_rows = max(raw_rows, new_rows)

        for r in range(1, max_rows + 1):

            # ---------- ADDED ROW ----------
            if r > raw_rows:
                added_count += 1
                for c in range(1, max_cols + 1):
                    new_ws.cell(r, c).fill = ADDED
                continue

            # ---------- DELETED ROW ----------
            if r > new_rows:
                deleted_count += 1
                continue

            # ---------- CELL COMPARISON ----------
            for c in range(1, max_cols + 1):

                old_val = normalize(raw_ws.cell(r, c).value)
                new_val = normalize(new_ws_val.cell(r, c).value)

                if old_val != new_val:
                    new_ws.cell(r, c).fill = MODIFIED
                    modified_count += 1

        summary.append([
            sheet,
            modified_count,
            added_count,
            deleted_count
        ])

    # ---------- SUMMARY SHEET ----------
    summary_ws = changed_wb.create_sheet("Diff_Summary")

    summary_ws.append(
        ["Sheet Name", "Modified Cells", "Added Rows", "Deleted Rows"]
    )

    for row in summary:
        summary_ws.append(row)

    # ---------- DOWNLOAD ----------
    with tempfile.NamedTemporaryFile(delete=False,
                                     suffix=".xlsx") as tmp:

        changed_wb.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.download_button(
                "⬇️ Download Smart Diff Excel",
                f,
                file_name="smart_diff.xlsx"
            )
