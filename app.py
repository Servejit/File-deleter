import streamlit as st
import os

st.title("File Deleter")

files_to_delete = [
    'abc.xlsx', 'xyz.xlsx', 'pqr.xlsx',
    'ABC.xlsx', 'XYZ.xlsx', 'PQR.xlsx'
]

if st.button("Delete Files"):

    for f in files_to_delete:

        if os.path.exists(f):
            os.remove(f)
            st.write(f"Deleted: {f}")
        else:
            st.write(f"Not found: {f}")
          
