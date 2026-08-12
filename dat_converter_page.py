"""DAT-to-Excel converter page.

Mirrors the Zone Merger's interaction model. Session-state keys are prefixed
`dat_` because st.navigation pages share one session state.
"""

import streamlit as st

from dat_pipeline import convert_dat_files


st.title("DAT to Excel Converter")
st.write(
    "Upload one or more `.dat` files. Each file's `#` header becomes a "
    "**Metadata** sheet and its numeric block becomes a **Data** sheet with "
    "proper column names. You get back a `.zip` holding one `.xlsx` per file."
)

# Session state init
if "dat_result" not in st.session_state:
    st.session_state.dat_result = None  # tuple (filename, bytes) or None
if "dat_log" not in st.session_state:
    st.session_state.dat_log = []  # list of str
if "dat_uploader_key" not in st.session_state:
    st.session_state.dat_uploader_key = 0

uploaded = st.file_uploader(
    "Drop .dat files here (or click to browse)",
    accept_multiple_files=True,
    type=["dat"],
    key=f"dat_uploader_{st.session_state.dat_uploader_key}",
)

if uploaded:
    st.write(f"**{len(uploaded)} file(s) queued:**")
    for f in uploaded:
        st.write(f"- {f.name}")

col_run, col_reset = st.columns([3, 1])
with col_run:
    run_clicked = st.button("Convert files", type="primary", disabled=not uploaded)
with col_reset:
    if st.button("Reset"):
        st.session_state.dat_result = None
        st.session_state.dat_log = []
        st.session_state.dat_uploader_key += 1
        st.rerun()

if run_clicked:
    st.session_state.dat_log = []
    st.session_state.dat_result = None

    log_placeholder = st.empty()

    def on_progress(msg: str) -> None:
        st.session_state.dat_log.append(msg)
        log_placeholder.code("\n".join(st.session_state.dat_log))

    files_payload = [(f.name, f.getvalue()) for f in uploaded]

    empty_files = [name for name, content in files_payload if not content]
    if empty_files:
        if len(empty_files) == len(files_payload):
            st.error(
                "Your file(s) are empty and cannot be converted: "
                + ", ".join(empty_files)
            )
            st.stop()
        else:
            st.warning(
                "The following file(s) are empty and will be skipped: "
                + ", ".join(empty_files)
            )

    try:
        with st.spinner("Converting..."):
            filename, zip_bytes = convert_dat_files(
                files_payload,
                progress_callback=on_progress,
            )
        st.session_state.dat_result = (filename, zip_bytes)
        st.success("Done.")
    except ValueError as e:
        st.error(f"Conversion error: {e}")
    except Exception as e:
        st.error(f"Unexpected error: {e}")
        st.exception(e)

if st.session_state.dat_log and not run_clicked:
    st.code("\n".join(st.session_state.dat_log))

if st.session_state.dat_result is not None:
    fname, data = st.session_state.dat_result
    col_dl, col_new = st.columns([3, 2])
    with col_dl:
        st.download_button(
            label=f"Download {fname}",
            data=data,
            file_name=fname,
            mime="application/zip",
            use_container_width=True,
        )
    with col_new:
        if st.button("Convert other files", use_container_width=True):
            st.session_state.dat_result = None
            st.session_state.dat_log = []
            st.session_state.dat_uploader_key += 1
            st.rerun()
