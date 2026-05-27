import streamlit as st
import pandas as pd
import zipfile
import tempfile
import os
import re
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="US Foods Report Generator",
    page_icon="📦",
    layout="wide"
)

st.title("📦 US Foods Report Generator")

st.write(
    "Upload the required reports and Outlook ZIP files to generate the US Foods report."
)

# ---------------------------------------------------
# Helper Functions
# ---------------------------------------------------

def normalize_po(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def extract_text_from_file(file_path):
    try:
        with open(file_path, "rb") as f:
            raw = f.read()

        try:
            return raw.decode("utf-8", errors="ignore")
        except:
            return raw.decode("latin-1", errors="ignore")

    except:
        return ""


def extract_requested_delivery_date(text):
    patterns = [
        r"<REQUESTED_DELIVERY_DATE>(.*?)</REQUESTED_DELIVERY_DATE>",
        r"<REQUEST_DELIVERY_DATE>(.*?)</REQUEST_DELIVERY_DATE>",
        r"<DELIVERY_DATE>(.*?)</DELIVERY_DATE>",
        r"Requested Delivery Date[:\s]+([0-9\/\-]+)"
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)

        if match:
            value = match.group(1).strip()
            parsed = pd.to_datetime(value, errors="coerce")

            if not pd.isna(parsed):
                return parsed

    return pd.NaT


def extract_po(text):
    patterns = [
        r"<TP_PO_NUMBER>(.*?)</TP_PO_NUMBER>",
        r"<ORDER_ID>(.*?)</ORDER_ID>",
        r"<ORDER_NUMBER>(.*?)</ORDER_NUMBER>",
        r"Customer PO Number[:\s]+([A-Za-z0-9\-]+)"
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)

        if match:
            return match.group(1).strip()

    return ""


def build_ship_date_lookup(zip_files, optional_files):

    lookup = {}

    with tempfile.TemporaryDirectory() as temp_dir:

        all_paths = []

        # ZIP FILES
        if zip_files:

            for uploaded_zip in zip_files:

                zip_path = os.path.join(temp_dir, uploaded_zip.name)

                with open(zip_path, "wb") as f:
                    f.write(uploaded_zip.read())

                try:
                    with zipfile.ZipFile(zip_path, "r") as z:

                        extract_dir = os.path.join(
                            temp_dir,
                            uploaded_zip.name.replace(".zip", "")
                        )

                        os.makedirs(extract_dir, exist_ok=True)

                        z.extractall(extract_dir)

                        for root, _, files in os.walk(extract_dir):

                            for file in files:

                                if file.lower().endswith(
                                    (".msg", ".eml", ".xml", ".txt")
                                ):

                                    all_paths.append(
                                        os.path.join(root, file)
                                    )

                except:
                    st.warning(f"Could not read ZIP: {uploaded_zip.name}")

        # OPTIONAL FILES
        if optional_files:

            optional_dir = os.path.join(temp_dir, "optional")

            os.makedirs(optional_dir, exist_ok=True)

            for uploaded_file in optional_files:

                path = os.path.join(optional_dir, uploaded_file.name)

                with open(path, "wb") as f:
                    f.write(uploaded_file.read())

                all_paths.append(path)

        # PROCESS FILES
        for path in all_paths:

            text = extract_text_from_file(path)

            po = extract_po(text)

            requested_date = extract_requested_delivery_date(text)

            if po and not pd.isna(requested_date):

                lookup[normalize_po(po)] = requested_date

    return lookup


def remove_tracking_rows(df, tracking_df):

    if tracking_df is None or tracking_df.empty:
        return df

    tracking_col = None
    po_col = None

    for col in tracking_df.columns:

        lower = str(col).lower()

        if "tracking" in lower:
            tracking_col = col

        if "po" in lower:
            po_col = col

    if tracking_col is None or po_col is None:
        return df

    shipped_pos = tracking_df[
        tracking_df[tracking_col].notna()
    ][po_col].astype(str).str.strip().tolist()

    main_po_col = None

    for col in df.columns:

        if "po" in str(col).lower():

            main_po_col = col
            break

    if main_po_col is None:
        return df

    df = df[
        ~df[main_po_col].astype(str).str.strip().isin(shipped_pos)
    ]

    return df


def get_cancelled_jobs(cancelled_df):

    if cancelled_df is None or cancelled_df.empty:
        return set()

    job_col = None

    for col in cancelled_df.columns:

        if "job" in str(col).lower():

            job_col = col
            break

    if job_col is None:
        return set()

    return set(
        cancelled_df[job_col]
        .dropna()
        .astype(str)
        .str.strip()
    )


def apply_ship_date_filter(df):

    if "Ship Date" not in df.columns:
        return df

    today = pd.Timestamp.today().normalize()

    df["_ShipDateParsed"] = pd.to_datetime(
        df["Ship Date"],
        errors="coerce"
    ).dt.normalize()

    keep_past_today = df["_ShipDateParsed"] <= today

    future_dates = df.loc[
        df["_ShipDateParsed"] > today,
        "_ShipDateParsed"
    ].dropna()

    if not future_dates.empty:

        next_date = future_dates.min()

        keep_next = df["_ShipDateParsed"] == next_date

        df = df[keep_past_today | keep_next]

    else:

        df = df[keep_past_today]

    df = df.drop(columns=["_ShipDateParsed"])

    return df


def split_sheets(df):

    laminate_col = None
    coating_col = None

    for col in df.columns:

        lower = str(col).lower()

        if "laminate" in lower:
            laminate_col = col

        if "coating" in lower:
            coating_col = col

    laminate_df = pd.DataFrame()
    uv_df = pd.DataFrame()
    trim_df = pd.DataFrame()

    if laminate_col:

        laminate_df = df[
            df[laminate_col]
            .astype(str)
            .str.upper()
            .str.strip()
            == "LAMINATE"
        ]

    if coating_col:

        uv_df = df[
            df[coating_col]
            .astype(str)
            .str.upper()
            .str.strip()
            != "NONE"
        ]

    if laminate_col and coating_col:

        trim_df = df[
            (
                df[laminate_col]
                .astype(str)
                .str.upper()
                .str.strip()
                == "NONE"
            )
            &
            (
                df[coating_col]
                .astype(str)
                .str.upper()
                .str.strip()
                == "NONE"
            )
        ]

    return {
        "All Open Jobs": df,
        "UV COATING": uv_df,
        "Laminate": laminate_df,
        "Trim To Size": trim_df
    }


def format_workbook(output, cancelled_jobs):

    wb = load_workbook(output)

    header_fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    cancelled_fill = PatternFill(
        "solid",
        fgColor="FF0000"
    )

    thin = Side(style="thin", color="000000")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    for ws in wb.worksheets:

        ws.freeze_panes = "A2"

        # HEADER FORMAT
        for cell in ws[1]:

            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center")

        # FIND JOB COL
        job_col_index = None

        for cell in ws[1]:

            if "job" in str(cell.value).lower():

                job_col_index = cell.column
                break

        # FORMAT ROWS
        for row in ws.iter_rows(min_row=2):

            for cell in row:

                cell.border = border

            if job_col_index:

                job_value = str(
                    ws.cell(
                        row=row[0].row,
                        column=job_col_index
                    ).value
                ).strip()

                if job_value in cancelled_jobs:

                    for cell in row:
                        cell.fill = cancelled_fill

        # AUTO WIDTH
        for col in ws.columns:

            max_length = 0

            column = get_column_letter(col[0].column)

            for cell in col:

                try:
                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )
                except:
                    pass

            ws.column_dimensions[column].width = min(
                max_length + 2,
                50
            )

    final_output = BytesIO()

    wb.save(final_output)

    final_output.seek(0)

    return final_output


# ---------------------------------------------------
# Upload Section
# ---------------------------------------------------

st.subheader("Upload Required Files")

main_report = st.file_uploader(
    "Master Tracking Numbers Report",
    type=["xls", "xlsx"],
    accept_multiple_files=False
)

cancelled_file = st.file_uploader(
    "Cancelled Status Report",
    type=["xls", "xlsx"],
    accept_multiple_files=False
)

zip_files = st.file_uploader(
    "US Foods Outlook Email ZIP file(s)",
    type=["zip"],
    accept_multiple_files=True
)

optional_files = st.file_uploader(
    "Optional: individual .msg, .eml, .xml, .txt files",
    type=["msg", "eml", "xml", "txt"],
    accept_multiple_files=True
)

# ---------------------------------------------------
# Generate Report
# ---------------------------------------------------

if st.button("Generate Report", type="primary"):

    if main_report is None:

        st.error("Please upload the Master Tracking Numbers Report.")

        st.stop()

    try:

        df = pd.read_excel(main_report)

    except Exception as e:

        st.error(f"Could not read report: {e}")

        st.stop()

    # CANCELLED REPORT
    cancelled_df = None

    if cancelled_file:

        try:
            cancelled_df = pd.read_excel(cancelled_file)

        except:
            st.warning("Could not read Cancelled Status Report.")

    # SHIP DATE LOOKUP
    st.info("Reading Outlook ZIP/XML files...")

    ship_date_lookup = build_ship_date_lookup(
        zip_files,
        optional_files
    )

    # FIND PO COLUMN
    po_col = None

    for col in df.columns:

        if "po" in str(col).lower():

            po_col = col
            break

    # ADD SHIP DATE
    if po_col:

        df["Ship Date"] = df[po_col].apply(
            lambda x: ship_date_lookup.get(
                normalize_po(x),
                pd.NaT
            )
        )

    else:

        df["Ship Date"] = pd.NaT

    # FILTER SHIP DATES
    df = apply_ship_date_filter(df)

    # SORT
    if "Ship Date" in df.columns:

        df = df.sort_values(
            by="Ship Date",
            ascending=True
        )

    # CANCELLED JOBS
    cancelled_jobs = get_cancelled_jobs(cancelled_df)

    # SPLIT SHEETS
    sheets = split_sheets(df)

    # WRITE EXCEL
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for sheet_name, sheet_df in sheets.items():

            sheet_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

    output.seek(0)

    final_output = format_workbook(
        output,
        cancelled_jobs
    )

    filename = (
        f"US_Foods_Report_"
        f"{datetime.today().strftime('%m%d%Y')}.xlsx"
    )

    st.success("Report generated successfully.")

    st.download_button(
        label="Download Finished Report",
        data=final_output,
        file_name=filename,
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    )