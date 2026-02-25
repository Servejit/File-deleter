import streamlit as st
import os
from openpyxl import load_workbook
from io import BytesIO
import zipfile

st.set_page_config(page_title="Excel Auto Processor", layout="centered")

st.title("📊 Excel Auto Processor System")

# =====================================================
# FILE MAP
# =====================================================

file_map = {
    'abc.xlsx': 'ABC.xlsx',
    'xyz.xlsx': 'XYZ.xlsx',
    'pqr.xlsx': 'PQR.xlsx'
}

# =====================================================
# STEP 1: DELETE OLD FILES
# =====================================================

st.header("Step 1: Delete Old Files")

if st.button("🗑 Delete Old Files"):

    for f in list(file_map.keys()) + list(file_map.values()):

        if os.path.exists(f):
            os.remove(f)
            st.success(f"Deleted: {f}")


# =====================================================
# STEP 2: UPLOAD NEW FILES
# =====================================================

st.header("Step 2: Upload New Files")

uploaded_files = st.file_uploader(

    "Upload abc.xlsx, xyz.xlsx, pqr.xlsx",

    type=["xlsx"],

    accept_multiple_files=True

)

if uploaded_files:

    for file in uploaded_files:

        with open(file.name, "wb") as f:
            f.write(file.getbuffer())

        st.success(f"Uploaded: {file.name}")


# =====================================================
# STEP 3: PROCESS FILES
# =====================================================

st.header("Step 3: Process and Download")

moves = [

    ('qYY', 'r'),
    ('qY', 'qYY'),
    ('q', 'qY'),
    ('pYY', 'q'),
    ('pY', 'pYY'),
    ('Y', 'pY')

]


if st.button("⚙ Process Files and Prepare Download"):

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:

        for input_file, output_file in file_map.items():

            if not os.path.exists(input_file):

                st.error(f"{input_file} not uploaded")
                continue

            wb = load_workbook(input_file)

            data_cache = {}

            # Read
            for src, _ in moves:

                sheet = wb[src]

                data_cache[src] = [

                    [sheet.cell(row=r, column=c).value for c in range(1, 16)]

                    for r in range(2, 52)

                ]

            # Paste
            for src, dst in moves:

                sheet = wb[dst]

                for r_idx, row in enumerate(data_cache[src], start=2):

                    for c_idx, val in enumerate(row, start=1):

                        sheet.cell(row=r_idx, column=c_idx).value = val


            # Save each file to memory
            file_buffer = BytesIO()
            wb.save(file_buffer)

            zip_file.writestr(output_file, file_buffer.getvalue())

            st.success(f"Processed: {output_file}")


    zip_buffer.seek(0)

    # =====================================================
    # SINGLE DOWNLOAD BUTTON
    # =====================================================

    st.download_button(

        label="⬇ Download ALL Files (ABC, XYZ, PQR)",

        data=zip_buffer,

        file_name="Processed_Excel_Files.zip",

        mime="application/zip"

    )
