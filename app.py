import streamlit as st
import os
from openpyxl import load_workbook
from io import BytesIO

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

    files_deleted = False

    for f in list(file_map.keys()) + list(file_map.values()):

        if os.path.exists(f):

            os.remove(f)

            st.success(f"Deleted: {f}")

            files_deleted = True

    if not files_deleted:

        st.info("No old files found")


# =====================================================
# STEP 2: UPLOAD NEW FILES
# =====================================================

st.header("Step 2: Upload New Files")

uploaded_files = st.file_uploader(

    "Upload abc.xlsx, xyz.xlsx, pqr.xlsx",

    type=["xlsx"],

    accept_multiple_files=True

)

uploaded_names = []

if uploaded_files:

    for file in uploaded_files:

        with open(file.name, "wb") as f:

            f.write(file.getbuffer())

        uploaded_names.append(file.name)

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


if st.button("⚙ Process Files"):

    for input_file, output_file in file_map.items():

        if not os.path.exists(input_file):

            st.error(f"{input_file} not uploaded")

            continue

        wb = load_workbook(input_file)

        data_cache = {}

        # Read Data

        for src, _ in moves:

            sheet = wb[src]

            data_cache[src] = [

                [sheet.cell(row=r, column=c).value for c in range(1, 16)]

                for r in range(2, 52)

            ]


        # Paste Data

        for src, dst in moves:

            sheet = wb[dst]

            for r_idx, row in enumerate(data_cache[src], start=2):

                for c_idx, val in enumerate(row, start=1):

                    sheet.cell(row=r_idx, column=c_idx).value = val


        # Save in memory

        output = BytesIO()

        wb.save(output)

        output.seek(0)

        st.success(f"Processed: {output_file}")


        # =====================================================
        # STEP 4: DOWNLOAD FILE
        # =====================================================

        st.download_button(

            label=f"⬇ Download {output_file}",

            data=output,

            file_name=output_file,

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )


# =====================================================
# SHOW CURRENT FILES
# =====================================================

st.header("📁 Files in System")

files = [f for f in os.listdir() if f.endswith(".xlsx")]

if files:

    for f in files:

        st.write(f)

else:

    st.write("No Excel files present")
