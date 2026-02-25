import streamlit as st
import zipfile
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Excel Auto Processor", layout="centered")

st.title("📊 Excel Auto Processor System")

# =====================================================
# CONFIG
# =====================================================

file_map = {
    "abc.xlsx": "ABC.xlsx",
    "xyz.xlsx": "XYZ.xlsx",
    "pqr.xlsx": "PQR.xlsx"
}

moves = [
    ("qYY", "r"),
    ("qY", "qYY"),
    ("q", "qY"),
    ("pYY", "q"),
    ("pY", "pYY"),
    ("Y", "pY")
]

# =====================================================
# SESSION STORAGE
# =====================================================

if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = {}

# =====================================================
# STEP 1 DELETE
# =====================================================

if st.button("🗑 Delete Old Files"):

    st.session_state.uploaded_files = {}

    st.success("Deleted successfully")


# =====================================================
# STEP 2 UPLOAD FORM (FIX)
# =====================================================

with st.form("upload_form"):

    files = st.file_uploader(
        "Upload files",
        type="xlsx",
        accept_multiple_files=True
    )

    upload_btn = st.form_submit_button("Upload")

    if upload_btn and files:

        for file in files:

            st.session_state.uploaded_files[file.name] = file.read()

        st.success("Upload completed")


# =====================================================
# STEP 3 PROCESS
# =====================================================

if st.button("⚙ Process Files"):

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:

        for input_name, output_name in file_map.items():

            if input_name not in st.session_state.uploaded_files:

                st.error(f"Missing {input_name}")

                continue

            wb = load_workbook(
                BytesIO(st.session_state.uploaded_files[input_name])
            )

            data_cache = {}

            for src, dst in moves:

                if src in wb.sheetnames:

                    sheet = wb[src]

                    data_cache[src] = [

                        [sheet.cell(row=r, column=c).value
                         for c in range(1, 16)]

                        for r in range(2, 52)

                    ]

            for src, dst in moves:

                if dst in wb.sheetnames and src in data_cache:

                    sheet = wb[dst]

                    for r, row in enumerate(data_cache[src], 2):

                        for c, val in enumerate(row, 1):

                            sheet.cell(r, c).value = val


            buffer = BytesIO()

            wb.save(buffer)

            zip_file.writestr(output_name, buffer.getvalue())

            st.success(f"Processed {output_name}")

    zip_buffer.seek(0)

    st.download_button(
        "⬇ Download ZIP",
        zip_buffer,
        "Processed.zip"
    )
