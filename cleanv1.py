import streamlit as st
import openpyxl
import io
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Excel Column Remover", layout="wide")

st.title("📊 Lightweight Excel Column Remover")

@st.cache_data
def get_headers(file):
    wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
    sheet = wb.active
    max_col = sheet.max_column
    headers = [
        sheet.cell(row=1, column=col).value or f"(empty {get_column_letter(col)})"
        for col in range(1, max_col + 1)
    ]
    return headers, max_col

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xlsm"])

if uploaded_file:
    headers, max_col = get_headers(uploaded_file)
    st.write(f"✅ Found **{max_col} columns**")

    selected_cols = st.multiselect(
        "Select columns to delete (searchable):",
        options=list(range(1, max_col + 1)),
        format_func=lambda x: f"{x}: {headers[x-1]}"
    )

    st.info(f"Selected {len(selected_cols)} / {max_col} columns")

    if st.button("Clean and Download"):
        wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
        sheet = wb.active
        new_wb = openpyxl.Workbook()
        new_sheet = new_wb.active
        new_sheet.title = f"{sheet.title}_cleaned"

        total_rows = sheet.max_row
        progress = st.progress(0)
        status_text = st.empty()

        for row_idx, row in enumerate(sheet.iter_rows(values_only=True, max_col=max_col), start=1):
            new_row = [cell for idx, cell in enumerate(row, start=1) if idx not in selected_cols]
            new_sheet.append(new_row)

            if row_idx % 500 == 0 or row_idx == total_rows:
                progress.progress(row_idx / total_rows)
                status_text.text(f"Processing {row_idx}/{total_rows} rows...")

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
