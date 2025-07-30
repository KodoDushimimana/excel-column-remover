import streamlit as st
import openpyxl
import io
from openpyxl.utils import get_column_letter
import time
import os

st.set_page_config(page_title="Excel Column Remover", layout="wide")

st.title("📊 Excel Column Remover")

@st.cache_data
def load_workbook(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    sheet = wb.active
    max_col = sheet.max_column
    headers = [
        sheet.cell(row=1, column=col).value or f"(empty {get_column_letter(col)})"
        for col in range(1, max_col + 1)
    ]
    return wb, sheet, max_col, headers

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xlsm"])

if uploaded_file:
    # Simulate upload progress
    progress = st.progress(0)
    status_text = st.empty()

    file_size = len(uploaded_file.getvalue())
    chunk_size = max(1, file_size // 20)  # 20 steps

    bytes_read = 0
    for i in range(1, 21):
        time.sleep(0.05)  # simulate delay
        bytes_read = min(file_size, i * chunk_size)
        percent = int((bytes_read / file_size) * 100)
        progress.progress(bytes_read / file_size)
        status_text.text(f"Uploading file... {percent}%")

    status_text.text("Upload complete!")
    progress.progress(1.0)

    wb, sheet, max_col, headers = load_workbook(uploaded_file)

    st.write(f"✅ Found **{max_col} columns** in `{sheet.title}`")

    st.subheader("Step 1: Select Columns to Delete")

    num_cols = 4
    cols = st.columns(num_cols)
    selections = {}

    for idx, header in enumerate(headers, start=1):
        col_idx = (idx - 1) % num_cols
        with cols[col_idx]:
            selections[idx] = st.checkbox(f"{idx}: {header}", key=f"col_{idx}")

    selected_cols = [idx for idx, checked in selections.items() if checked]
    st.info(f"Selected {len(selected_cols)} / {max_col} columns")

    if st.button("Clean and Download"):
        progress = st.progress(0)
        status_text = st.empty()

        new_wb = openpyxl.Workbook()
        new_sheet = new_wb.active
        new_sheet.title = f"{sheet.title}_cleaned"

        total_to_delete = len(selected_cols)
        for i, row in enumerate(sheet.iter_rows(values_only=True, max_col=max_col), start=1):
            new_row = [cell for idx, cell in enumerate(row, start=1) if idx not in selected_cols]
            new_sheet.append(new_row)

            if total_to_delete > 0 and i % 100 == 0:
                progress.progress(i / sheet.max_row)
                status_text.text(f"Processing row {i} of {sheet.max_row}...")

        output = io.BytesIO()
        new_wb.save(output)
        output.seek(0)

        st.success("✅ File cleaned successfully!")
        progress.progress(1.0)
        status_text.text("Done!")

        st.download_button(
            label="Download Cleaned Excel File",
            data=output,
            file_name="cleaned_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
