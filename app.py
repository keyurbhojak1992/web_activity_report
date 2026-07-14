import streamlit as st
import sys

st.set_page_config(page_title="Dependency Test")

st.title("Dependency diagnostic")
st.write("Python:", sys.version)
st.success("Streamlit imported successfully")

try:
    import numpy as np
    st.success(f"NumPy imported: {np.__version__}")
except Exception as exc:
    st.exception(exc)
    st.stop()

try:
    import pandas as pd
    st.success(f"Pandas imported: {pd.__version__}")
except Exception as exc:
    st.exception(exc)
    st.stop()

try:
    import openpyxl
    st.success(f"OpenPyXL imported: {openpyxl.__version__}")
except Exception as exc:
    st.exception(exc)
    st.stop()

st.success("All dependencies are working.")
