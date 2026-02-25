import streamlit as st
import zipfile
from io import BytesIO
from openpyxl import load_workbook

# =====================================================
# CONFIG
# =====================================================

st.set_page_config(page_title="Excel Auto Processor", layout="centered")

st.title("📊 Excel Auto Processor System")


# =====================================================
# FILE MAP
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

if "files" not in st.session_state:
    st.session_state.files = {}


# =====================================================
# STEP 1 DELETE
# =====================================================

st.header("Step 1: Delete Old Files")

if st.button("🗑 Delete Old Files"):

    st.session_state.files = {}

    st.success("All files deleted")


# =====================================================
# STEP 2 UPLOAD (FIXED)
# =====================================================

st.header("Step 2: Upload Files")

uploaded_files = st.file_uploader(
    "Upload abc.xlsx, xyz.xlsx, pqr.xlsx",
    type="xlsx",
    accept_multiple_files=True
)

if uploaded_files:

    for file in uploaded_files:

        st.session_state.files[file.name] = file.getvalue()

        st.success(f"Uploaded: {file.name}")


# =====================================================
# STEP 3 PROCESS
# =====================================================

st.header("Step 3: Process and Download")

if st.button("⚙ Process Files"):

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:

        for input_name, output_name in file_map.items():

            if input_name not in st.session_state.files:

                st.error(f"Missing: {input_name}")

                continue

            file_bytes = st.session_state.files[input_name]

            wb = load_workbook(BytesIO(file_bytes))

            data_cache = {}

            # READ

            for src, dst in moves:

                if src in wb.sheetnames:

                    sheet = wb[src]

                    data_cache[src] = [

                        [sheet.cell(row=r, column=c).value for c in range(1, 16)]

                        for r in range(2, 52)

                    ]

            # WRITE

            for src, dst in moves:

                if dst in wb.sheetnames and src in data_cache:

                    sheet = wb[dst]

                    for r_idx, row in enumerate(data_cache[src], start=2):

                        for c_idx, val in enumerate(row, start=1):

                            sheet.cell(row=r_idx, column=c_idx).value = val


            output_buffer = BytesIO()

            wb.save(output_buffer)

            zip_file.writestr(output_name, output_buffer.getvalue())

            st.success(f"Processed: {output_name}")


    zip_buffer.seek(0)


    st.download_button(

        "⬇ Download ZIP",

        zip_buffer,

        "Processed_Excel_Files.zip",

        "application/zip"
    )


# =====================================================
# REFRESH
# =====================================================

if st.button("🔄 Refresh"):

    st.rerun()
