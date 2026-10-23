import io
import hashlib
import re
from datetime import datetime, timezone

import numpy as np
import pandas as pd
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

st.set_page_config(
    page_title="PGD Special Label Uploader",
    page_icon="📊",
    layout="wide",
)

st.title("📊 PGD Special Label Dashboard Uploader")
st.caption("Upload Excel → ETL v1.1 → audit & validation → update Google Sheets.")

# ---------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------
# Spreadsheet ID from the Google Sheets URL.
# Worksheet is resolved by its numeric gid (sheetId), not tab name.
REQUIRED_COLUMNS = [
    "Order Date", "PO#", "SO", "Article No.", "Model", "Line", "Building",
    "Order Qty", "Add Qty", "Add %", "Reason", "Label Item",
    "Label Category", "RCA", "ShipTo Country Name", "LPD", "PODD",
]
DATE_COLS = ["Order Date", "Taken Date", "LPD", "PODD"]
ID_COLS = ["PO#", "SO", "Article No.", "Model", "Label Item"]
CATEGORY_MASTER = {
    "SHOEBOX LABEL": "Shoebox Label",
    "HANGTAG RFID": "Hangtag RFID",
    "HANGTAG STANDART": "Hangtag Standart",
    "SPECIAL LABEL": "Special Label",
}
CATEGORY_ALIASES = {"HANGTAG STANDARD": "HANGTAG STANDART"}
COUNTRY_MAP = {"UNITED STATES": "USA", "US": "USA"}
REASON_MAP = {
    "PRODUCTION ISSUE": "Production Issue",
    "C GRADE": "C-Grade",
    "C-GRADE": "C-Grade",
    "REJECT": "Reject",
}

# Kontrak kolom yang dihasilkan ETL dan diharapkan dashboard.
ETL_COLUMNS = [
    "Year", "Quarter", "Month", "Month No", "Year Month", "Building Line",
    "Severity", "Add Qty Bucket", "Order Qty Bucket", "Request Flag",
    "Add % Calculated", "Add % Difference", "Category Review Flag",
]


def normalize_column_name(col):
    col = str(col).replace("\n", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", col).strip()


def normalize_headers(df):
    df = df.copy()
    df.columns = [normalize_column_name(c) for c in df.columns]
    aliases = {
        "Article No": "Article No.",
        "Add Qty (Pcs)": "Add Qty",
        "Model Name": "Model",
    }
    renames = {old: new for old, new in aliases.items()
               if old in df.columns and new not in df.columns}
    df.rename(columns=renames, inplace=True)
    dupes = df.columns[df.columns.duplicated()].tolist()
    if dupes:
        raise ValueError(f"Header duplikat setelah normalisasi: {dupes}")
    return df


def clean_text(series):
    s = series.astype("string")
    s = (s.str.replace("\xa0", " ", regex=False).str.strip()
         .str.replace(r"\s+", " ", regex=True))
    return s.mask(s.str.lower().isin(["", "nan", "none", "n/a", "null"]))


def to_numeric_safe(series):
    s = clean_text(series).str.replace(",", "", regex=False)
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    return pd.to_numeric(s, errors="coerce")


def parse_percent(series):
    result = []
    for value in series:
        if pd.isna(value):
            result.append(np.nan)
        elif isinstance(value, str):
            v = value.strip().replace(",", "")
            if not v:
                result.append(np.nan)
                continue
            has_percent = "%" in v
            try:
                number = float(v.replace("%", "").strip())
                result.append(number / 100 if has_percent else number)
            except (ValueError, TypeError):
                result.append(np.nan)
        else:
            try:
                result.append(float(value))
            except (ValueError, TypeError):
                result.append(np.nan)
    return pd.Series(result, index=series.index, dtype="float64")


def parse_date_safe(series):
    # Meniru ETL v1.1: teks tanggal diasumsikan MM/DD/YYYY.
    return pd.to_datetime(series, errors="coerce", dayfirst=False)


def standardize_label_category(series):
    s = clean_text(series)
    out = []
    for value in s:
        if pd.isna(value):
            out.append(pd.NA)
            continue
        key = CATEGORY_ALIASES.get(value.upper().strip(), value.upper().strip())
        out.append(CATEGORY_MASTER.get(key, value))
    return pd.Series(out, index=series.index, dtype="string")


def assert_schema(df):
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Kolom wajib tidak ditemukan: " + ", ".join(missing)
            + "\nKolom tersedia: " + ", ".join(map(str, df.columns))
        )


def clean_data(df_raw):
    """ETL Special Label sejalan dengan pipeline Colab v1.1."""
    df = normalize_headers(df_raw)
    assert_schema(df)

    for col in df.select_dtypes(include=["object", "string"]).columns:
        df[col] = clean_text(df[col])

    for col in ID_COLS:
        df[col] = clean_text(df[col])

    for col in ["Building", "Line"]:
        df[col] = df[col].str.upper().str.strip()

    df["ShipTo Country Name"] = df["ShipTo Country Name"].str.upper().replace(COUNTRY_MAP)
    df["Reason"] = (df["Reason"].str.upper()
                    .str.replace(r"\(\?\)", "", regex=True)
                    .str.replace("?", "", regex=False)
                    .str.replace(r"\s+", " ", regex=True).str.strip()
                    .replace(REASON_MAP))
    if "RCA" in df.columns:
        df["RCA"] = clean_text(df["RCA"])
    df["Label Item"] = clean_text(df["Label Item"])
    df["Label Category"] = standardize_label_category(df["Label Category"])

    for col in ["Order Qty", "Add Qty"]:
        df[col] = to_numeric_safe(df[col])
    df["Add %"] = parse_percent(df["Add %"])

    for col in DATE_COLS:
        if col in df.columns:
            df[col] = parse_date_safe(df[col])

    for col, default in {"Reason": "Unknown", "RCA": "No RCA", "Remark": "No Remark"}.items():
        if col in df.columns:
            df[col] = df[col].fillna(default)

    # Calendar attributes
    df["Year"] = df["Order Date"].dt.year.astype("Int64")
    df["Quarter"] = "Q" + df["Order Date"].dt.quarter.astype("Int64").astype("string")
    df["Month"] = df["Order Date"].dt.strftime("%b")
    df["Month No"] = df["Order Date"].dt.month.astype("Int64")
    df["Year Month"] = df["Order Date"].dt.strftime("%Y-%m")
    df["Building Line"] = df["Building"].fillna("Unknown") + "-" + df["Line"].fillna("Unknown")

    df["Severity"] = np.select(
        [df["Add %"] < 0.01, df["Add %"].between(0.01, 0.03, inclusive="both")],
        ["Low", "Medium"], default="High")
    df.loc[df["Add %"].isna(), "Severity"] = "Unknown"
    df["Add Qty Bucket"] = pd.cut(
        df["Add Qty"], [-np.inf, 100, 500, 1000, np.inf],
        labels=["0-100", "101-500", "501-1000", "1001+"], right=True)
    df["Order Qty Bucket"] = pd.cut(
        df["Order Qty"], [-np.inf, 5000, 10000, np.inf],
        labels=["0-5000", "5001-10000", "10001+"], right=True)
    df["Request Flag"] = 1
    df["Add % Calculated"] = np.where(df["Order Qty"] > 0,
                                      df["Add Qty"] / df["Order Qty"], np.nan)
    df["Add % Difference"] = df["Add %"] - df["Add % Calculated"]
    std_categories = list(CATEGORY_MASTER.values())
    df["Category Review Flag"] = np.where(
        df["Label Category"].isna(), "Needs Review",
        np.where(df["Label Category"].isin(std_categories), "Assigned",
                 "Review Unknown Category"))
    return df


def build_validation(df):
    checks = []
    def add(name, passed, detail):
        checks.append({"Validation": name, "Status": "PASS" if passed else "FAIL", "Detail": detail})
    dup = int(df.duplicated().sum())
    add("Duplicate Rows", dup == 0, f"{dup} duplicate rows")
    missing_po = int(df["PO#"].isna().sum())
    add("Missing PO#", missing_po == 0, f"{missing_po} missing")
    invalid_date = int(df["Order Date"].isna().sum())
    add("Missing/Invalid Order Date", invalid_date == 0, f"{invalid_date} missing/invalid")
    neg_order = int((df["Order Qty"] < 0).sum())
    add("Negative Order Qty", neg_order == 0, f"{neg_order} negative values")
    neg_add = int((df["Add Qty"] < 0).sum())
    add("Negative Add Qty", neg_add == 0, f"{neg_add} negative values")
    pct_valid = df["Add %"].dropna().between(0, 1).all()
    add("Add % Range", bool(pct_valid), "Non-null values should be between 0% and 100%")
    buildings = set(df["Building"].dropna().unique())
    add("Building Valid", buildings <= {"A", "B", "C", "D"}, f"Found: {sorted(buildings)}")
    standards = set(CATEGORY_MASTER.values())
    unknown_n = int((~df["Label Category"].isin(standards) & df["Label Category"].notna()).sum())
    add("Standard Label Category", unknown_n == 0, f"{unknown_n} nonstandard/unknown rows")
    missing_cat = int(df["Label Category"].isna().sum())
    add("Label Category Filled", missing_cat == 0, f"{missing_cat} missing categories")
    diff = df["Add % Difference"].abs()
    mismatch = int((diff.notna() & (diff > 0.005)).sum())
    add("Add % Consistency", mismatch == 0, f"{mismatch} rows differ by >0.5 percentage points")
    report = pd.DataFrame(checks)
    return report, round((report["Status"] == "PASS").mean() * 100, 2)


def validate_upload(uploaded_bytes):
    raw = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=0)
    if raw.empty:
        raise ValueError("File tidak memiliki data.")
    df = clean_data(raw)
    # Make dashboard-facing order stable: source fields first, ETL fields last.
    source_cols = [c for c in df.columns if c not in ETL_COLUMNS]
    ordered_cols = source_cols + [c for c in ETL_COLUMNS if c in df.columns]
    df = df[ordered_cols].copy()
    validation_df, quality_score = build_validation(df)
    known = set(CATEGORY_MASTER.values())
    unknown = sorted(str(v) for v in df["Label Category"].dropna().unique() if v not in known)
    return df, unknown, validation_df, quality_score


def get_google_credentials():
    if "gdrive_service_account" not in st.secrets:
        raise RuntimeError(
            "Secret [gdrive_service_account] belum ditemukan."
        )

    info = dict(st.secrets["gdrive_service_account"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    return service_account.Credentials.from_service_account_info(
        info,
        scopes=scopes,
    )


def get_sheets_service():
    credentials = get_google_credentials()
    return build(
        "sheets",
        "v4",
        credentials=credentials,
        cache_discovery=False,
    )


def get_config():
    drive_cfg = st.secrets.get("drive", {})
    spreadsheet_id = str(drive_cfg.get("target_file_id", "")).strip()
    worksheet_gid = str(drive_cfg.get("target_gid", "")).strip()

    if not spreadsheet_id:
        raise RuntimeError("drive.target_file_id belum diisi.")
    if not worksheet_gid:
        raise RuntimeError(
            "drive.target_gid belum diisi. Isi dengan gid tab tujuan."
        )

    try:
        worksheet_gid_int = int(worksheet_gid)
    except ValueError:
        raise RuntimeError("drive.target_gid harus berupa angka.")

    return spreadsheet_id, worksheet_gid_int


def validate_upload(uploaded_bytes):
    # Default: first worksheet in uploaded workbook.
    # Adjust sheet_name if your source worksheet is not the first tab.
    df = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=0)
    df.columns = [normalize_header(c) for c in df.columns]

    if df.columns.duplicated().any():
        dupes = df.columns[df.columns.duplicated()].tolist()
        raise ValueError(f"Header duplikat: {dupes}")

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Kolom wajib tidak ditemukan: "
            + ", ".join(missing)
            + "\nKolom yang ditemukan: "
            + ", ".join(map(str, df.columns))
        )

    if df.empty:
        raise ValueError("File tidak memiliki data.")

    # Keep expected field order and any extra columns after required fields.
    extras = [c for c in df.columns if c not in REQUIRED_COLUMNS]
    df = df[REQUIRED_COLUMNS + extras].copy()

    df["Label Category"] = df["Label Category"].map(standardize_category)

    known = set(CATEGORY_MAP.values())
    unknown = sorted(
        str(v)
        for v in df["Label Category"].dropna().unique()
        if str(v).strip() and v not in known
    )

    # Convert source rows to Google Sheets API-compatible scalar values.
    return df, unknown


def to_sheet_value(value):
    """Convert pandas/numpy values to JSON-compatible sheet cell values."""
    if pd.isna(value):
        return ""

    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d")

    # Preserve identifier-like fields as strings, especially PO/SO/article.
    if isinstance(value, (int,)):
        return int(value)
    if isinstance(value, (float,)):
        if value.is_integer():
            return int(value)
        return float(value)

    # Handle numpy scalar types.
    if hasattr(value, "item"):
        try:
            value = value.item()
        except Exception:
            pass

    if isinstance(value, (int, float, bool)):
        return value
    return str(value).strip()


def dataframe_to_values(df):
    headers = list(df.columns)
    rows = [headers]

    for row in df.itertuples(index=False, name=None):
        rows.append([to_sheet_value(v) for v in row])

    return rows


def get_spreadsheet_metadata(service, spreadsheet_id):
    return (
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields="spreadsheetId,properties(title),sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
        )
        .execute()
    )


def resolve_target_tab(metadata, target_gid):
    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if int(props.get("sheetId", -1)) == target_gid:
            return props
    available = [
        f"{s.get('properties', {}).get('title')} "
        f"(gid={s.get('properties', {}).get('sheetId')})"
        for s in metadata.get("sheets", [])
    ]
    raise RuntimeError(
        f"gid {target_gid} tidak ditemukan. Tab tersedia: "
        + "; ".join(available)
    )


def quote_sheet_title(title):
    # Escape single quotes in sheet title for A1 notation.
    return "'" + title.replace("'", "''") + "'"


def replace_worksheet_values(service, spreadsheet_id, tab_title, values):
    """
    Replace cell values in the target tab.
    Clears existing values first, then writes the new table.
    Formatting and worksheet/tab identity are retained.
    """
    quoted_title = quote_sheet_title(tab_title)

    # Clear all existing values in this worksheet.
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"{quoted_title}",
        body={},
    ).execute()

    # Write headers + data, starting at A1.
    response = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{quoted_title}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()

    return response


# ---------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------
with st.sidebar:
    st.header("Configuration")
    st.info(
        "Target Spreadsheet ID dan target tab gid dibaca dari Streamlit Secrets."
    )
    st.caption("Input: Excel .xlsx")
    st.caption("Output: Google Sheets native")
    st.caption("Worksheet sumber: tab pertama di file upload")

# ---------------------------------------------------------------------
# UPLOAD & VALIDATION
# ---------------------------------------------------------------------
upload = st.file_uploader(
    "Upload file Excel Special Label (.xlsx)",
    type=["xlsx"],
    help="Workbook ini akan menggantikan isi nilai pada worksheet target.",
)

if not upload:
    st.info("Upload file Excel untuk memulai.")
    st.stop()

file_bytes = upload.getvalue()
file_hash = hashlib.sha256(file_bytes).hexdigest()

st.write(f"**File:** {upload.name} · **Size:** {len(file_bytes):,} bytes")

try:
    df, unknown_categories, validation_df, quality_score = validate_upload(file_bytes)
except Exception as exc:
    st.error(f"Validasi gagal: {exc}")
    st.stop()

# Metrics
order_qty = pd.to_numeric(
    df["Order Qty"].astype(str).str.replace(",", "", regex=False),
    errors="coerce",
)
add_qty = pd.to_numeric(
    df["Add Qty"].astype(str).str.replace(",", "", regex=False),
    errors="coerce",
)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Rows", f"{len(df):,}")
m2.metric("Unique PO", f"{df['PO#'].nunique(dropna=True):,}")
m3.metric("Total Add Qty", f"{add_qty.sum():,.0f}")
m4.metric("Label Items", f"{df['Label Item'].nunique(dropna=True):,}")

st.success(f"ETL v1.1 selesai. Schema output: {len(df.columns)} kolom.")
st.metric("Data Quality Score", f"{quality_score:.2f}/100")
st.subheader("Validation Report")
st.dataframe(validation_df, use_container_width=True)
failed_checks = validation_df[validation_df["Status"] == "FAIL"]
if not failed_checks.empty:
    st.warning("Ada validation check berstatus FAIL. Periksa report sebelum publish.")
st.caption(f"SHA-256 file upload: `{file_hash}`")

if unknown_categories:
    st.warning(
        "Label Category di luar standard ditemukan: "
        + ", ".join(unknown_categories[:30])
    )
    st.caption(
        "Kategori di luar standard dipertahankan apa adanya. "
        "Mohon review sebelum update."
    )

if df["Label Category"].eq("").any():
    st.warning(
        f"Ada {(df['Label Category'] == '').sum():,} baris tanpa Label Category."
    )

st.subheader("Preview data")
st.dataframe(df.head(30), use_container_width=True)

st.subheader("ETL Dashboard Fields")
st.caption("Field hasil ETL yang ditambahkan: " + ", ".join(c for c in ETL_COLUMNS if c in df.columns))
st.subheader("Distribution by Label Category")
category_summary = (
    df["Label Category"]
    .replace("", "(Blank)")
    .value_counts(dropna=False)
    .rename_axis("Label Category")
    .reset_index(name="Rows")
)
st.dataframe(category_summary, use_container_width=True)

# ---------------------------------------------------------------------
# TARGET INSPECTION
# ---------------------------------------------------------------------
st.divider()
st.subheader("Target Google Sheets")

try:
    spreadsheet_id, target_gid = get_config()
    sheets_service = get_sheets_service()
    metadata = get_spreadsheet_metadata(sheets_service, spreadsheet_id)
    target_props = resolve_target_tab(metadata, target_gid)
    target_title = target_props["title"]

    st.write(f"**Spreadsheet:** {metadata.get('properties', {}).get('title', 'Unknown')}")
    st.write(f"**Target tab:** {target_title}")
    st.write(f"**gid:** `{target_gid}`")
    st.write(f"**Rows in uploaded data:** {len(df):,}")
    st.write(f"**Columns in uploaded data:** {len(df.columns):,}")
    st.success("Koneksi dan akses metadata Google Sheets berhasil.")

except HttpError as exc:
    st.error(
        "Gagal membaca Google Sheets. Pastikan Google Sheets API aktif, "
        "Spreadsheet ID benar, dan spreadsheet sudah di-share ke service account."
    )
    st.code(str(exc))
    st.stop()
except Exception as exc:
    st.error(f"Konfigurasi/koneksi gagal: {exc}")
    st.stop()

# ---------------------------------------------------------------------
# CONFIRM & WRITE
# ---------------------------------------------------------------------
st.warning(
    "PERHATIAN: Update akan menghapus nilai lama di tab target, lalu menulis "
    "header dan data baru mulai A1. Format sel/tab dipertahankan sejauh "
    "tidak diubah oleh operasi values API. Pastikan tab target memang khusus "
    "untuk dataset ini dan backup sudah tersedia."
)

confirm_target = st.text_input(
    f"Ketik UPDATE untuk menyetujui penggantian data pada tab: {target_title}",
    placeholder="UPDATE",
)
backup_confirm = st.checkbox(
    "Saya sudah membuat backup dan memastikan tab target benar."
)

# Mencegah publish bila schema dashboard belum lengkap.
missing_etl = [c for c in ETL_COLUMNS if c not in df.columns]
if missing_etl:
    st.error("Schema ETL tidak lengkap: " + ", ".join(missing_etl))
can_update = (confirm_target.strip().upper() == "UPDATE" and backup_confirm
              and not missing_etl and not df.empty)

if st.button(
    "☁️ Update Google Sheets",
    type="primary",
    disabled=not can_update,
    use_container_width=True,
):
    try:
        values = dataframe_to_values(df)

        # Guardrail: ensure nonempty data and expected header row.
        if len(values) < 2:
            st.error("Tidak ada baris data untuk ditulis.")
            st.stop()

        with st.spinner("Meng-update Google Sheets..."):
            response = replace_worksheet_values(
                sheets_service,
                spreadsheet_id,
                target_title,
                values,
            )

        st.success("Google Sheets berhasil diperbarui.")
        st.write(f"**Updated range:** {response.get('updatedRange', 'N/A')}")
        st.write(f"**Updated rows:** {response.get('updatedRows', 'N/A')}")
        st.write(f"**Updated columns:** {response.get('updatedColumns', 'N/A')}")
        st.write(f"**Updated cells:** {response.get('updatedCells', 'N/A')}")

        sheet_url = (
            f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
            f"#gid={target_gid}"
        )
        st.markdown(f"[Buka Google Sheets target]({sheet_url})")

        st.info(
            "Data di Google Sheets sudah ditulis. Perubahan Looker Studio "
            "bergantung pada konektor, cache, dan pengaturan refresh data source."
        )

    except HttpError as exc:
        st.error("Google Sheets API gagal saat menulis data.")
        st.code(str(exc))
        st.caption(
            "Periksa izin Editor, batas ukuran request, nama tab, dan konfigurasi API."
        )
    except Exception as exc:
        st.error(f"Update gagal: {exc}")

st.divider()
st.caption(
    "PGD Special Label Uploader · Gunakan backup dan akses aplikasi terbatas "
    "sebelum deployment untuk penggunaan operasional."
)
