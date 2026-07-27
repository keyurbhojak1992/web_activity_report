from __future__ import annotations

import gc
import hashlib
import traceback

import streamlit as st

from report_builder import (
    OUTPUT_FILE_NAME,
    BuildResult,
    InputValidationError,
    build_report,
)


st.set_page_config(
    page_title="Diamond Web-Activity Log Processor",
    page_icon="💎",
    layout="centered",
)

SAFE_WEBLOG_SIZE_MB = 15


def _megabytes(size_bytes: int) -> float:
    return size_bytes / (1024 * 1024)


def _input_signature(*files, critical_search_only_searchid: bool) -> str:
    digest = hashlib.sha256()
    for uploaded_file in files:
        if uploaded_file is None:
            digest.update(b"<none>")
            continue
        digest.update(uploaded_file.name.encode("utf-8", errors="replace"))
        digest.update(str(uploaded_file.size).encode("ascii"))
        digest.update(uploaded_file.getvalue()[:1024])
    digest.update(
        f"<critical-search-only-searchid:{int(critical_search_only_searchid)}>".encode(
            "ascii"
        )
    )
    return digest.hexdigest()


def _clear_generated_report() -> None:
    result = st.session_state.pop("generated_report", None)
    if isinstance(result, BuildResult):
        result.cleanup()
    st.session_state.pop("generated_signature", None)


st.title("💎 Web Action Report Generator")
st.caption(
    "Creates the sales action report, counts each matched critical shape, and keeps the uploaded Weblog worksheet in the output workbook."
)

with st.form("report_inputs", clear_on_submit=False):
    weblog_file = st.file_uploader(
        "1. Upload Weblog Data (Excel) *Required*",
        type=["xlsx"],
        help="The app uses the worksheet named 'Weblog Data', or the first worksheet when that name is absent.",
    )
    master_file = st.file_uploader(
        "2. Upload Color Master Data (Excel) *Required*",
        type=["xlsx"],
    )
    sales_file = st.file_uploader(
        "3. Upload Sales and Bid Data (Excel) *Optional*",
        type=["xlsx"],
    )
    critical_file = st.file_uploader(
        "4. Upload Critical Search Data (Excel) *Required*",
        type=["xlsx"],
    )

    critical_search_only_searchid = st.checkbox(
        "Count Critical Search only in rows containing SEARCHID",
        value=True,
        help=(
            "Enabled: scan only descriptions containing SEARCHID. "
            "Disabled: scan every Weblog description for critical master shapes."
        ),
    )

    allow_large_weblog = st.checkbox(
        f"Allow a Weblog file larger than {SAFE_WEBLOG_SIZE_MB} MB",
        value=False,
        help=(
            "Preserving an Excel worksheet requires substantial RAM. Enable this only on a server "
            "with at least 4 GB memory."
        ),
    )

    submitted = st.form_submit_button("Generate Report", type="primary", use_container_width=True)

if submitted:
    _clear_generated_report()

    missing_files = []
    if weblog_file is None:
        missing_files.append("Weblog Data")
    if master_file is None:
        missing_files.append("Color Master Data")
    if critical_file is None:
        missing_files.append("Critical Search Data")

    if missing_files:
        st.warning("Please upload: " + ", ".join(missing_files) + ".")
    elif weblog_file.size > SAFE_WEBLOG_SIZE_MB * 1024 * 1024 and not allow_large_weblog:
        st.error(
            f"The Weblog file is {_megabytes(weblog_file.size):.1f} MB. "
            f"To prevent a cloud-memory crash, the safe limit is {SAFE_WEBLOG_SIZE_MB} MB. "
            "Reduce the workbook size or enable the large-file option when using a higher-memory server."
        )
    else:
        status_text = st.empty()

        def update_status(message: str) -> None:
            status_text.info(message)

        try:
            with st.spinner("Generating report..."):
                result: BuildResult = build_report(
                    weblog_bytes=weblog_file.getvalue(),
                    master_bytes=master_file.getvalue(),
                    critical_bytes=critical_file.getvalue(),
                    sales_bytes=sales_file.getvalue() if sales_file is not None else None,
                    critical_search_only_searchid=critical_search_only_searchid,
                    progress_callback=update_status,
                )

            signature = _input_signature(
                weblog_file,
                master_file,
                sales_file,
                critical_file,
                critical_search_only_searchid=critical_search_only_searchid,
            )
            st.session_state["generated_report"] = result
            st.session_state["generated_signature"] = signature
            status_text.empty()
            st.success(
                f"Report generated successfully with {result.report_rows:,} report row(s)."
            )
        except InputValidationError as exc:
            status_text.empty()
            st.error(str(exc))
        except MemoryError:
            status_text.empty()
            st.error(
                "The server ran out of memory while processing the workbook. Reduce the Weblog file size "
                "or deploy the app on a server with more RAM."
            )
        except Exception as exc:
            status_text.empty()
            st.error(f"Report generation failed: {type(exc).__name__}: {exc}")
            with st.expander("Technical details"):
                st.code(traceback.format_exc(), language="text")
        finally:
            gc.collect()

result = st.session_state.get("generated_report")
generated_signature = st.session_state.get("generated_signature")
current_signature = (
    _input_signature(
        weblog_file,
        master_file,
        sales_file,
        critical_file,
        critical_search_only_searchid=critical_search_only_searchid,
    )
    if weblog_file is not None and master_file is not None and critical_file is not None
    else None
)

if isinstance(result, BuildResult) and current_signature != generated_signature:
    st.warning("The uploaded files changed after the last report was generated. Click Generate Report again.")
elif isinstance(result, BuildResult):
    st.caption(f"Output size: {_megabytes(result.file_size):.1f} MB")
    st.download_button(
        label=f"📥 Download {OUTPUT_FILE_NAME}",
        data=result.read_bytes,
        file_name=OUTPUT_FILE_NAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        on_click="ignore",
        key="download_generated_report",
    )

    if result.warnings:
        with st.expander("Excel compatibility warnings"):
            for warning in result.warnings:
                st.write(f"• {warning}")

    if st.button("Clear generated file from memory", use_container_width=True):
        _clear_generated_report()
        gc.collect()
        st.rerun()

st.divider()
st.caption(
    "Developed by Keyur Bhojak for Web-Activity analysis.
                © 2026 Weblog Tracker. All rights reserved."
)
