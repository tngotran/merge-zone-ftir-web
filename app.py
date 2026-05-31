"""Streamlit UI for the FTIR zone-merge pipeline."""

import streamlit as st

from pipeline import process_dpt_files


st.set_page_config(page_title="FTIR Zone Merger", page_icon="🧪", layout="centered")
st.title("FTIR Zone Merger")
st.write(
    "Upload one or more `.dpt` files. The app parses each spectrum, "
    "computes the standard peak-ratio metric, merges per-zone, and gives "
    "you a single multi-sheet Excel file to download."
)

# Session state init
if "result" not in st.session_state:
    st.session_state.result = None  # tuple (filename, bytes) or None
if "log" not in st.session_state:
    st.session_state.log = []  # list of str

uploaded = st.file_uploader(
    "Drop .dpt files here (or click to browse)",
    accept_multiple_files=True,
    type=["dpt"],
)

if uploaded:
    st.write(f"**{len(uploaded)} file(s) queued:**")
    for f in uploaded:
        st.write(f"- {f.name}")

col_run, col_reset = st.columns([3, 1])
with col_run:
    run_clicked = st.button("Process files", type="primary", disabled=not uploaded)
with col_reset:
    if st.button("Reset"):
        st.session_state.result = None
        st.session_state.log = []
        st.rerun()

if run_clicked:
    st.session_state.log = []
    st.session_state.result = None

    log_placeholder = st.empty()

    def on_progress(msg: str) -> None:
        st.session_state.log.append(msg)
        log_placeholder.code("\n".join(st.session_state.log))

    files_payload = [(f.name, f.getvalue()) for f in uploaded]

    empty_files = [name for name, content in files_payload if len(content) < 10]
    if empty_files:
        if len(empty_files) == len(files_payload):
            st.error(
                f"Your file(s) are empty and cannot be processed: "
                + ", ".join(empty_files)
            )
            st.stop()
        else:
            st.warning(
                f"The following file(s) are empty and will be skipped: "
                + ", ".join(empty_files)
            )

    try:
        with st.spinner("Processing..."):
            filename, xlsx_bytes = process_dpt_files(
                files_payload,
                progress_callback=on_progress,
            )
        st.session_state.result = (filename, xlsx_bytes)
        st.success("Done.")
    except ValueError as e:
        st.error(f"Pipeline error: {e}")
    except Exception as e:
        st.error(f"Unexpected error: {e}")
        st.exception(e)

if st.session_state.log and not run_clicked:
    st.code("\n".join(st.session_state.log))

if st.session_state.result is not None:
    fname, data = st.session_state.result
    col_dl, col_new = st.columns([3, 2])
    with col_dl:
        st.download_button(
            label=f"Download {fname}",
            data=data,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_new:
        if st.button("Process other files", use_container_width=True):
            st.session_state.result = None
            st.session_state.log = []
            st.rerun()
