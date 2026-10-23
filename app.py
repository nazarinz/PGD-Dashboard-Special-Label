import io
import hashlib
from datetime import datetime, timezone

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
st.caption("Upload Excel terbaru → validasi → update Google Sheets.")

# ---------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------
# Spreadsheet ID from the Google Sheets URL.
# Worksheet is resolved by its numeric gid (sheetId), not tab name.
REQUIRED_COLUMNS = [
    "Order Date",
    "PO#",
    "SO",
    "Article No.",
    "Model",
    "Line",
    "Building",
    "Order Qty",
    "Add Qty",
    "Add %",
    "Reason",
    "Label Item",
    "Label Category",
    "RCA",
    "Taken Date",
    "ShipTo Country Name",
    "LPD",
    "PODD",
    "Remark",
]

CATEGORY_MAP = {
    "SHOEBOX LABEL": "Shoebox Label",
    "HANGTAG RFID": "Hangtag RFID",
    "HANGTAG STANDART": "Hangtag Standart",
    "SPECIAL LABEL": "Special Label",
}
CATEGORY_ALIASES = {
    "HANGTAG STANDARD": "HANGTAG STANDART",
}


def normalize_header(value):
    return " ".join(str(value).replace("\xa0", " ").split()).strip()


def standardize_category(value):
    if pd.isna(value):
        return ""
    clean = " ".join(str(value).replace("\xa0", " ").split()).strip()
    if not clean:
        return ""
    key = clean.upper()
    key = CATEGORY_ALIASES.get(key, key)
    return CATEGORY_MAP.get(key, clean)


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
    df, unknown_categories = validate_upload(file_bytes)
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

st.success("Validasi struktur berhasil.")
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

can_update = confirm_target.strip().upper() == "UPDATE" and backup_confirm

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
