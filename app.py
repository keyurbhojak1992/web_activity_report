import gc
import os
import re
import tempfile
from datetime import date, datetime

import numpy as np
import pandas as pd
import streamlit as st
import xlsxwriter

st.set_page_config(page_title="Diamond WebLog Processor", layout="centered")

st.title("💎 Weblog Activity Report Generator")
st.caption("Memory-safe Excel export build")

FILL_HIGHEST_ACTION = True
FILL_TYPE_COLUMN = True
FILL_IP_COUNT = False
OUTPUT_FILE_NAME = "Weblog Report.xlsx"

REQUIRED_ACTION_COLUMNS = [
    "Detail",
    "Critical_Search",
    "EXCEL",
    "MEDIA",
    "Layout",
    "NEW ARRIVAL",
    "TWIN STONES",
    "WISHLIST",
]

COLOR_MAP = {
    "GREEN": "#00FF00",
    "RED": "#FF0000",
    "BLUE": "#0000FF",
    "YELLOW": "#FFFF00",
    "ORANGE": "#FFA500",
}

st.write("Upload your source files below to generate the report.")
weblog_file = st.file_uploader(
    "1. Upload Weblog Data (Excel) *Required*", type=["xlsx"]
)
master_file = st.file_uploader(
    "2. Upload Color Master Data (Excel) *Required*", type=["xlsx"]
)
sales_file = st.file_uploader(
    "3. Upload Sales and Bid Data (Excel) *Optional*", type=["xlsx"]
)
critical_file = st.file_uploader(
    "4. Upload Critical Search Data (Excel) *Required*", type=["xlsx"]
)


def require_columns(df: pd.DataFrame, required: list[str], label: str) -> None:
    missing = [column for column in required if column not in df.columns]
    if missing:
        raise ValueError(f"{label} is missing required columns: {', '.join(missing)}")


def normalize_excel_value(value):
    """Convert pandas/numpy values to types accepted by XlsxWriter."""
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, np.floating):
        return float(value)
    if isinstance(value, np.bool_):
        return bool(value)
    if isinstance(value, pd.Timestamp):
        if value.tzinfo is not None:
            value = value.tz_convert(None)
        return value.to_pydatetime()
    if isinstance(value, (list, tuple, dict, set)):
        return str(value)[:32767]
    if isinstance(value, str):
        return value[:32767]
    return value


def display_length(value) -> int:
    value = normalize_excel_value(value)
    if value is None:
        return 0
    return len(str(value))


def write_value(
    worksheet,
    row: int,
    col: int,
    value,
    cell_format,
    datetime_format=None,
) -> None:
    value = normalize_excel_value(value)

    if value is None:
        worksheet.write_blank(row, col, None, cell_format)
    elif isinstance(value, bool):
        worksheet.write_boolean(row, col, value, cell_format)
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        worksheet.write_number(row, col, value, cell_format)
    elif isinstance(value, (datetime, date)):
        worksheet.write_datetime(
            row,
            col,
            value if isinstance(value, datetime) else datetime.combine(value, datetime.min.time()),
            datetime_format or cell_format,
        )
    else:
        worksheet.write(row, col, value, cell_format)


def determine_row_color(type_value, color_value) -> str:
    type_text = str(type_value or "").upper()
    for name in ("RED", "ORANGE", "GREEN", "YELLOW", "BLUE"):
        if name in type_text:
            return COLOR_MAP[name]
    return COLOR_MAP.get(str(color_value or "").upper().strip(), "#CCCCCC")


def write_memory_safe_workbook(
    final_df: pd.DataFrame,
    logs_df: pd.DataFrame,
    output_path: str,
) -> None:
    """Write the final XLSX without loading a second in-memory workbook."""
    workbook = xlsxwriter.Workbook(
        output_path,
        {
            "constant_memory": True,
            "nan_inf_to_errors": True,
        },
    )

    common = {"align": "center", "valign": "vcenter", "border": 1}
    header_format = workbook.add_format(
        {**common, "bold": True, "bg_color": "#D9D9D9"}
    )
    body_format = workbook.add_format(common)
    amount_format = workbook.add_format({**common, "num_format": "#,##0.00"})
    date_format = workbook.add_format({**common, "num_format": "yyyy-mm-dd hh:mm:ss"})
    hyperlink_format = workbook.add_format(
        {**common, "font_color": "#0000FF", "underline": True}
    )

    critical_format = workbook.add_format({**common, "bg_color": "#FFF2CC"})
    critical_date_format = workbook.add_format(
        {**common, "bg_color": "#FFF2CC", "num_format": "yyyy-mm-dd hh:mm:ss"}
    )
    special_format = workbook.add_format({**common, "bg_color": "#E6B8B7"})
    special_amount_format = workbook.add_format(
        {**common, "bg_color": "#E6B8B7", "num_format": "#,##0.00"}
    )

    fill_formats: dict[tuple[str, str], object] = {}

    def get_fill_format(hex_color: str, kind: str):
        key = (hex_color, kind)
        if key in fill_formats:
            return fill_formats[key]

        properties = {**common, "bg_color": hex_color}
        if kind == "amount":
            properties["num_format"] = "#,##0.00"
        elif kind == "hyperlink":
            properties["font_color"] = "#0000FF"
            properties["underline"] = True

        fill_formats[key] = workbook.add_format(properties)
        return fill_formats[key]

    # ------------------------------
    # Report worksheet
    # ------------------------------
    report_sheet = workbook.add_worksheet("Report")
    output_columns = [column for column in final_df.columns if column != "Color"]
    report_widths = [len(str(column)) for column in output_columns]

    for column_index, column_name in enumerate(output_columns):
        report_sheet.write(0, column_index, column_name, header_format)

    action_columns = set(REQUIRED_ACTION_COLUMNS)

    for excel_row, values in enumerate(
        final_df.itertuples(index=False, name=None), start=1
    ):
        row_map = dict(zip(final_df.columns, values))
        party_name = str(row_map.get("PARTY_COMPANY_NAME", "")).strip().upper()
        special_row = party_name == "K GIRDHARLAL & CO."
        row_color = determine_row_color(row_map.get("Type"), row_map.get("Color"))

        action_values = [
            float(row_map.get(column, 0) or 0) for column in REQUIRED_ACTION_COLUMNS
        ]
        max_action = max(action_values) if action_values else 0

        for column_index, column_name in enumerate(output_columns):
            value = row_map.get(column_name)
            report_widths[column_index] = min(
                50, max(report_widths[column_index], display_length(value))
            )

            if special_row:
                fmt = special_amount_format if column_name == "AMT" else special_format
                write_value(report_sheet, excel_row, column_index, value, fmt)
                continue

            is_hyperlink = column_name in action_columns and float(value or 0) > 0
            use_fill = False

            if column_name == "Type" and FILL_TYPE_COLUMN:
                use_fill = True
            elif (
                column_name in action_columns
                and FILL_HIGHEST_ACTION
                and max_action > 0
                and float(value or 0) == max_action
            ):
                use_fill = True
            elif (
                column_name == "IP Counts"
                and FILL_IP_COUNT
                and float(value or 0) >= 5
            ):
                use_fill = True

            if is_hyperlink:
                fmt = (
                    get_fill_format(row_color, "hyperlink")
                    if use_fill
                    else hyperlink_format
                )
                report_sheet.write_url(
                    excel_row,
                    column_index,
                    "internal:'Weblog Data'!A1",
                    fmt,
                    str(normalize_excel_value(value)),
                )
            else:
                if use_fill:
                    fmt = get_fill_format(
                        row_color, "amount" if column_name == "AMT" else "body"
                    )
                else:
                    fmt = amount_format if column_name == "AMT" else body_format
                write_value(report_sheet, excel_row, column_index, value, fmt)

    for column_index, width in enumerate(report_widths):
        report_sheet.set_column(column_index, column_index, min(width + 3, 50))

    report_sheet.freeze_panes(1, 0)
    if output_columns:
        report_sheet.autofilter(
            0, 0, max(len(final_df), 1), len(output_columns) - 1
        )

    # ------------------------------
    # Raw weblog worksheet
    # ------------------------------
    weblog_sheet = workbook.add_worksheet("Weblog Data")
    weblog_columns = list(logs_df.columns)
    weblog_widths = [len(str(column)) for column in weblog_columns]
    action_index = weblog_columns.index("Action") if "Action" in weblog_columns else None

    for column_index, column_name in enumerate(weblog_columns):
        weblog_sheet.write(0, column_index, column_name, header_format)

    for excel_row, values in enumerate(
        logs_df.itertuples(index=False, name=None), start=1
    ):
        critical_row = (
            action_index is not None
            and values[action_index] == "Critical_Search"
        )
        row_format = critical_format if critical_row else body_format
        row_date_format = critical_date_format if critical_row else date_format

        for column_index, value in enumerate(values):
            weblog_widths[column_index] = min(
                50, max(weblog_widths[column_index], display_length(value))
            )
            write_value(
                weblog_sheet,
                excel_row,
                column_index,
                value,
                row_format,
                row_date_format,
            )

    for column_index, width in enumerate(weblog_widths):
        weblog_sheet.set_column(column_index, column_index, min(width + 3, 50))

    weblog_sheet.freeze_panes(1, 0)
    if weblog_columns:
        weblog_sheet.autofilter(
            0, 0, max(len(logs_df), 1), len(weblog_columns) - 1
        )

    workbook.close()


if st.button("Generate Report"):
    if not (weblog_file and master_file and critical_file):
        st.warning(
            "Please upload the required files (Weblog, Color Master, and Critical Search)."
        )
        st.stop()

    temp_output_path = None
    try:
        with st.status("Generating report...", expanded=True) as status:
            st.write("Reading Excel files...")
            print("[PROC 1] Reading input files", flush=True)

            logs_df = pd.read_excel(weblog_file, engine="openpyxl")
            master_df = pd.read_excel(master_file, engine="openpyxl")
            critical_df = pd.read_excel(critical_file, engine="openpyxl")

            logs_df.columns = logs_df.columns.astype(str).str.strip()
            master_df.columns = master_df.columns.astype(str).str.strip()
            critical_df.columns = critical_df.columns.astype(str).str.strip()

            require_columns(
                logs_df,
                [
                    "PARTY_COMPANY_NAME",
                    "EMPLOYEE_SHORT_NAME",
                    "description",
                    "ipAddress",
                ],
                "Weblog file",
            )
            require_columns(master_df, ["Company Name", "Color"], "Color Master file")
            require_columns(
                critical_df, ["Grp", "Name", "Short Name"], "Critical Search file"
            )

            if "Zone" not in master_df.columns:
                master_df["Zone"] = ""

            st.write(
                f"Loaded {len(logs_df):,} weblog rows, "
                f"{len(master_df):,} master rows, and "
                f"{len(critical_df):,} critical-search rows."
            )
            print(f"[PROC 2] Loaded weblog rows={len(logs_df)}", flush=True)

            st.write("Cleaning and classifying activity data...")
            logs_df["PARTY_COMPANY_NAME"] = (
                logs_df["PARTY_COMPANY_NAME"].astype(str).str.strip().str.upper()
            )
            logs_df["EMPLOYEE_SHORT_NAME"] = (
                logs_df["EMPLOYEE_SHORT_NAME"]
                .fillna("Unknown")
                .astype(str)
                .str.strip()
                .str.upper()
            )
            logs_df["description"] = (
                logs_df["description"].astype(str).str.strip().str.upper()
            )

            master_df["Company Name"] = (
                master_df["Company Name"].astype(str).str.strip().str.upper()
            )
            master_df["Color"] = (
                master_df["Color"].astype(str).str.strip().str.upper()
            )
            master_df["Zone"] = master_df["Zone"].astype(str).str.strip()

            critical_df["Grp"] = (
                critical_df["Grp"].astype(str).str.strip().str.upper()
            )
            critical_filtered = critical_df[
                critical_df["Grp"].isin(["F3", "CRITICAL"])
            ]
            critical_keywords = {
                item
                for item in (
                    critical_filtered["Name"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .str.upper()
                    .tolist()
                    + critical_filtered["Short Name"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .str.upper()
                    .tolist()
                )
                if item
            }

            def extract_action(description: str) -> str:
                if pd.isna(description) or description == "NAN":
                    return "OTHER"

                if "SEARCHID" in description or "SEARCH PERFORMED" in description:
                    for keyword in critical_keywords:
                        if re.search(r"\b" + re.escape(keyword) + r"\b", description):
                            return "Critical_Search"
                    return "Detail"

                if any(
                    keyword in description
                    for keyword in (
                        "VIDEO",
                        "PHYGITAL",
                        "CERTIFICATE",
                        "DETAIL",
                        "IMAGE",
                        "PLOTING",
                    )
                ):
                    return "MEDIA"
                if "EXCEL" in description:
                    return "EXCEL"
                if "LAYOUT" in description:
                    return "Layout"
                if "NEW ARRIVAL" in description:
                    return "NEW ARRIVAL"
                if "TWIN STONES" in description:
                    return "TWIN STONES"
                if "WISHLIST" in description:
                    return "WISHLIST"
                return "OTHER"

            logs_df["Action"] = logs_df["description"].map(extract_action)
            print("[PROC 3] Action classification completed", flush=True)

            st.write("Aggregating report totals...")
            action_counts = pd.pivot_table(
                logs_df,
                index=["PARTY_COMPANY_NAME", "EMPLOYEE_SHORT_NAME"],
                columns="Action",
                aggfunc="size",
                fill_value=0,
            ).reset_index()

            for column in REQUIRED_ACTION_COLUMNS:
                if column not in action_counts.columns:
                    action_counts[column] = 0

            action_counts["Grand Total"] = action_counts[
                REQUIRED_ACTION_COLUMNS
            ].sum(axis=1)

            ip_counts = (
                logs_df.groupby(
                    ["PARTY_COMPANY_NAME", "EMPLOYEE_SHORT_NAME"], sort=False
                )["ipAddress"]
                .nunique()
                .reset_index(name="IP Counts")
            )

            report_df = pd.merge(
                action_counts,
                ip_counts,
                on=["PARTY_COMPANY_NAME", "EMPLOYEE_SHORT_NAME"],
                how="left",
            ).rename(columns={"EMPLOYEE_SHORT_NAME": "Sales_Person"})

            if sales_file is not None:
                sales_df = pd.read_excel(
                    sales_file,
                    usecols=["Sold Party", "Type", "AMT"],
                    engine="openpyxl",
                )
                sales_df.columns = sales_df.columns.astype(str).str.strip()
                sales_df["Sold Party"] = (
                    sales_df["Sold Party"].astype(str).str.strip().str.upper()
                )
                sales_df["Type"] = (
                    sales_df["Type"].astype(str).str.strip().str.upper()
                )
                sales_df["AMT"] = pd.to_numeric(
                    sales_df["AMT"], errors="coerce"
                ).fillna(0)

                amt_grouped = (
                    sales_df[sales_df["Type"].isin(["SALE", "BID"])]
                    .groupby("Sold Party", sort=False)["AMT"]
                    .sum()
                    .reset_index()
                )
                report_df = pd.merge(
                    report_df,
                    amt_grouped,
                    left_on="PARTY_COMPANY_NAME",
                    right_on="Sold Party",
                    how="left",
                )
                report_df["AMT"] = report_df["AMT"].fillna(0)
                del sales_df, amt_grouped
            else:
                report_df["AMT"] = 0

            final_df = pd.merge(
                report_df,
                master_df[["Company Name", "Color", "Zone"]],
                left_on="PARTY_COMPANY_NAME",
                right_on="Company Name",
                how="inner",
            )
            final_df.rename(columns={"Zone": "Type"}, inplace=True)
            final_df["Remark"] = ""

            final_columns = [
                "PARTY_COMPANY_NAME",
                "Type",
                "Remark",
                "Sales_Person",
                "Critical_Search",
                "Detail",
                "EXCEL",
                "MEDIA",
                "Layout",
                "NEW ARRIVAL",
                "TWIN STONES",
                "WISHLIST",
                "IP Counts",
                "Grand Total",
                "AMT",
                "Color",
            ]
            final_df = final_df[final_columns]
            print(f"[PROC 4] Report rows={len(final_df)}", flush=True)

            st.write("Writing the Excel workbook in low-memory mode...")
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
                temp_output_path = temp_file.name

            write_memory_safe_workbook(final_df, logs_df, temp_output_path)
            print("[PROC 5] Workbook written", flush=True)

            del action_counts, ip_counts, report_df, master_df, critical_df, final_df
            gc.collect()

            with open(temp_output_path, "rb") as output_file:
                report_bytes = output_file.read()

            status.update(label="Report generated successfully", state="complete")

        st.success("Report Generated Successfully!")
        st.download_button(
            label="📥 Download Report.xlsx",
            data=report_bytes,
            file_name=OUTPUT_FILE_NAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as exc:
        print(f"[ERROR] {type(exc).__name__}: {exc}", flush=True)
        st.exception(exc)
    finally:
        if temp_output_path and os.path.exists(temp_output_path):
            try:
                os.remove(temp_output_path)
            except OSError:
                pass
