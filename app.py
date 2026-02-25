import streamlit as st
import os
import zipfile
from io import BytesIO
from openpyxl import load_workbook

# =========================================================
# CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Excel Auto Processor",
    layout="centered"
)

st.title("📊 Excel Auto Processor System")

# =========================================================
# FILE CONFIG
# =========================================================

INPUT_FILES = ["abc.xlsx", "xyz.xlsx", "pqr.xlsx"]

OUTPUT_MAP = {
    "abc.xlsx": "ABC.xlsx",
    "xyz.xlsx": "XYZ.xlsx",
    "pqr.xlsx": "PQR.xlsx"
}

# Sheet move logic
MOVE_RULES = [
    ("qYY", "r"),
    ("qY", "qYY"),
    ("q", "qY"),
    ("pYY", "q"),
    ("pY", "pYY"),
    ("Y", "pY")
]

# =========================================================
# FUNCTION : DELETE FILES
# =========================================================

def delete_old_files():

    deleted_any = False

    for file in INPUT_FILES + list(OUTPUT_MAP.values()):

        if os.path.exists(file):

            os.remove(file)
            st.success(f"Deleted: {file}")
            deleted_any = True

    if not deleted_any:
        st.info("No old files found")


# =========================================================
# FUNCTION : SAVE UPLOADED FILES
# =========================================================

def save_uploaded_files(files):

    for file in files:

        with open(file.name, "wb") as f:
            f.write(file.getbuffer())

        st.success(f"Uploaded: {file.name}")


# =========================================================
# FUNCTION : PROCESS FILE
# =========================================================

def process_excel(input_file):

    wb = load_workbook(input_file)

    cache = {}

    # READ DATA
    for source, target in MOVE_RULES:

        if source in wb.sheetnames:

            sheet = wb[source]

            cache[source] = [

                [sheet.cell(row=r, column=c).value for c in range(1, 16)]

                for r in range(2, 52)

            ]

    # WRITE DATA
    for source, target in MOVE_RULES:

        if target in wb.sheetnames and source in cache:

            sheet = wb[target]

            for r_index, row in enumerate(cache[source], start=2):

                for c_index, value in enumerate(row, start=1):

                    sheet.cell(row=r_index, column=c_index).value = value

    buffer = BytesIO()

    wb.save(buffer)

    return buffer.getvalue()


# =========================================================
# STEP 1 : DELETE
# =========================================================

st.header("Step 1: Delete Old Files")

if st.button("🗑 Delete Old Files"):

    delete_old_files()


# =========================================================
# STEP 2 : UPLOAD
# =========================================================

st.header("Step 2: Upload Files")

uploaded = st.file_uploader(

    "Upload abc.xlsx, xyz.xlsx, pqr.xlsx",

    type=["xlsx"],

    accept_multiple_files=True

)

if uploaded:

    save_uploaded_files(uploaded)


# =========================================================
# STEP 3 : PROCESS
# =========================================================

st.header("Step 3: Process and Download")

if st.button("⚙ Process Files"):

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zipf:

        for input_file, output_file in OUTPUT_MAP.items():

            if not os.path.exists(input_file):

                st.error(f"Missing file: {input_file}")
                continue

            processed = process_excel(input_file)

            zipf.writestr(output_file, processed)

            st.success(f"Processed: {output_file}")

    zip_buffer.seek(0)

    st.download_button(

        label="⬇ Download ZIP",

        data=zip_buffer,

        file_name="Processed_Excel_Files.zip",

        mime="application/zip"
    )


# =========================================================
# REFRESH
# =========================================================

if st.button("🔄 Refresh"):

    st.rerun()
