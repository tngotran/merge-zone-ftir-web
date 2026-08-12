"""Streamlit entry point. Routes between the app's tools.

Stays the deployment entry point (Streamlit Cloud main file) — the individual
tools live in *_page.py modules.
"""

import streamlit as st

st.set_page_config(page_title="FTIR Tools", page_icon="🧪", layout="centered")

pages = [
    st.Page(
        "zone_merger_page.py",
        title="FTIR Zone Merger",
        icon="🧪",
        default=True,
    ),
    st.Page(
        "dat_converter_page.py",
        title="DAT to Excel",
        icon="📊",
    ),
]

st.navigation(pages).run()
