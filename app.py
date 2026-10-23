import io
import re
import hashlib
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

st.set_page_config(page_title="PGD Special Label Uploader", page_icon="📊", layout="wide")
st.title("📊 PGD Special Label Dashboard Uploader")
st.caption("Upload Excel → ETL v1.1 (identik dengan Colab) → audit & validation → update Google Sheets.")

VERSION = "v1.1"

# ================================================================
# 1. CONFIGURATION (sama dengan Colab v1.1)
# ================================================================
REQUIRED_COLS = [
    "Order Date", "PO#", "SO", "Article No.", "Model", "Line", "Building",
    "Order Qty", "Add Qty", "Add %", "Reason", "Label Item",
    "Label Category", "RCA", "ShipTo Country Name", "LPD", "PODD",
]
DATE_COLS = ["Order Date", "Taken Date", "LPD", "PODD"]
ID_COLS = ["PO#", "SO", "Article No.", "Model", "Label Item"]
DROP_COLS = []

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

ETL_COLUMNS = [
    "Year", "Quarter", "Month", "Month No", "Year Month", "Building Line",
    "Severity", "Add Qty Bucket", "Order Qty Bucket", "Request Flag",
    "Add % Calculated", "Add % Difference", "Category Review Flag",
]

# Tulis kolom ID sebagai teks di Google Sheets (cegah 00123 -> 123).
ID_AS_TEXT = True


# ================================================================
# 2. HELPERS
# ================================================================
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
    rename_dict = {old: new for old, new in aliases.items()
                   if old in df.columns and new not in df.columns}
    df.rename(columns=rename_dict, inplace=True)
    dupes = df.columns[df.columns.duplicated()].tolist()
    if dupes:
        raise ValueError("Terdapat nama kolom duplikat setelah normalisasi: " + str(dupes))
    return df


def clean_text(series):
    s = series.astype("string")
    s = (s.str.replace("\xa0", " ", regex=False)
          .str.strip()
          .str.replace(r"\s+", " ", regex=True))
    empty_values = ["", "nan", "None", "NaN", "N/A", "NULL"]
    return s.mask(s.isin(empty_values), pd.NA)


def normalize_id(series):
    return clean_text(series)


def to_numeric_safe(series):
    s = clean_text(series)
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    return pd.to_numeric(s, errors="coerce")


def parse_percent(series):
    """'0.15%' -> 0.0015 ; angka numerik dipakai apa adanya."""
    result = []
    for value in series:
        if pd.isna(value):
            result.append(np.nan)
            continue
        if isinstance(value, str):
            v = value.strip().replace(",", "")
            if not v:
                result.append(np.nan)
                continue
            has_percent = "%" in v
            v = v.replace("%", "").strip()
            try:
                number = float(v)
                if has_percent:
                    number /= 100
                result.append(number)
            except (ValueError, TypeError):
                result.append(np.nan)
        else:
            try:
                result.append(float(value))
            except (ValueError, TypeError):
                result.append(np.nan)
    return pd.Series(result, index=series.index, dtype="float64")


def parse_date_safe(series):
    # Teks tanggal diasumsikan MM/DD/YYYY. Ubah dayfirst=True bila sumber DD/MM/YYYY.
    return pd.to_datetime(series, errors="coerce", dayfirst=False)


def standardize_label_category(series):
    s = clean_text(series)
    standardized = []
    for value in s:
        if pd.isna(value):
            standardized.append(pd.NA)
            continue
        key = value.upper().strip()
        key = CATEGORY_ALIASES.get(key, key)
        category = CATEGORY_MASTER.get(key)
        standardized.append(category if category else value)
    return pd.Series(standardized, index=series.index, dtype="string")


def assert_schema(df):
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            "Kolom wajib tidak ditemukan:\n"
            + "\n".join(f"- {c}" for c in missing)
            + "\n\nKolom tersedia:\n" + ", ".join(map(str, df.columns))
        )


# ================================================================
# 3. ETL / CLEAN DATA (identik dengan Colab v1.1)
# ================================================================
def clean_data(df_raw):
    df = normalize_headers(df_raw)
    assert_schema(df)

    # 8A. Cleaning seluruh kolom teks
    for col in df.select_dtypes(include=["object", "string"]).columns:
        df[col] = clean_text(df[col])

    # 8B. Identifiers
    for col in ID_COLS:
        df[col] = normalize_id(df[col])

    # 8C. Building & Line
    for col in ["Building", "Line"]:
        df[col] = df[col].str.upper().str.strip()

    # 8D. Country
    df["ShipTo Country Name"] = df["ShipTo Country Name"].str.upper().replace(COUNTRY_MAP)

    # 8E. Reason
    df["Reason"] = (
        df["Reason"].str.upper()
        .str.replace(r"\(\?\)", "", regex=True)
        .str.replace("?", "", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .replace(REASON_MAP)
    )

    # 8F-8H. RCA, Label Item, Label Category
    if "RCA" in df.columns:
        df["RCA"] = clean_text(df["RCA"])
    df["Label Item"] = clean_text(df["Label Item"])
    df["Label Category"] = standardize_label_category(df["Label Category"])

    # 8I. Numeric
    for col in ["Order Qty", "Add Qty"]:
        df[col] = to_numeric_safe(df[col])
    df["Add %"] = parse_percent(df["Add %"])

    # 8J. Dates
    for col in DATE_COLS:
        if col in df.columns:
            df[col] = parse_date_safe(df[col])

    # 8K. Defaults
    for col, default in {"Reason": "Unknown", "RCA": "No RCA", "Remark": "No Remark"}.items():
        if col in df.columns:
            df[col] = df[col].fillna(default)

    # 8L. Calendar
    df["Year"] = df["Order Date"].dt.year.astype("Int64")
    df["Quarter"] = "Q" + df["Order Date"].dt.quarter.astype("Int64").astype("string")
    df["Month"] = df["Order Date"].dt.strftime("%b")
    df["Month No"] = df["Order Date"].dt.month.astype("Int64")
    df["Year Month"] = df["Order Date"].dt.strftime("%Y-%m")

    # 8M. Building Line
    df["Building Line"] = df["Building"].fillna("Unknown") + "-" + df["Line"].fillna("Unknown")

    # 8N. Severity
    df["Severity"] = np.select(
        [df["Add %"] < 0.01, df["Add %"].between(0.01, 0.03, inclusive="both")],
        ["Low", "Medium"],
        default="High",
    )
    df.loc[df["Add %"].isna(), "Severity"] = "Unknown"

    # 8O. Buckets
    df["Add Qty Bucket"] = pd.cut(
        df["Add Qty"], bins=[-np.inf, 100, 500, 1000, np.inf],
        labels=["0-100", "101-500", "501-1000", "1001+"], right=True)
    df["Order Qty Bucket"] = pd.cut(
        df["Order Qty"], bins=[-np.inf, 5000, 10000, np.inf],
        labels=["0-5000", "5001-10000", "10001+"], right=True)

    # 8P-8Q. Request flag & calculated Add %
    df["Request Flag"] = 1
    df["Add % Calculated"] = np.where(df["Order Qty"] > 0, df["Add Qty"] / df["Order Qty"], np.nan)
    df["Add % Difference"] = df["Add %"] - df["Add % Calculated"]

    # 8R. Category review flag
    std = list(CATEGORY_MASTER.values())
    df["Category Review Flag"] = np.where(
        df["Label Category"].isna(), "Needs Review",
        np.where(df["Label Category"].isin(std), "Assigned", "Review Unknown Category"))

    # 8S. Drop columns (opsional)
    existing_drop = [c for c in DROP_COLS if c in df.columns]
    if existing_drop:
        df.drop(columns=existing_drop, inplace=True)

    return df


# ================================================================
# 4. AUDIT, VALIDATION, SUMMARY (identik dengan Colab v1.1)
# ================================================================
def build_audit(df):
    has_date = df["Order Date"].notna().any()
    date_start = df["Order Date"].min() if has_date else pd.NaT
    date_end = df["Order Date"].max() if has_date else pd.NaT
    total_order = df["Order Qty"].sum(min_count=1)
    total_add = df["Add Qty"].sum(min_count=1)
    weighted = total_add / total_order if pd.notna(total_order) and total_order > 0 else np.nan

    metrics = {
        "Rows": len(df),
        "Columns": len(df.columns),
        "Duplicate Rows": int(df.duplicated().sum()),
        "Total Missing Cells": int(df.isna().sum().sum()),
        "Start Date": date_start,
        "End Date": date_end,
        "Total Order Qty": total_order,
        "Total Add Qty": total_add,
        "Average Add %": df["Add %"].mean(),
        "Weighted Add %": weighted,
        "Total PO": df["PO#"].nunique(),
        "Total SO": df["SO"].nunique(),
        "Total Model": df["Model"].nunique(),
        "Total Article": df["Article No."].nunique(),
        "Total Label Item": df["Label Item"].nunique(),
        "Total Label Category": df["Label Category"].nunique(),
        "Total Building": df["Building"].nunique(),
        "Total Line": df["Line"].nunique(),
        "Total Country": df["ShipTo Country Name"].nunique(),
        "Reason Category": df["Reason"].nunique(),
        "RCA Category": df["RCA"].nunique(),
        "Missing Label Category": int(df["Label Category"].isna().sum()),
        "Unknown Category Rows": int((df["Category Review Flag"] == "Review Unknown Category").sum()),
        "Invalid Order Date": int(df["Order Date"].isna().sum()),
        "Invalid Add %": int(df["Add %"].isna().sum()),
    }
    out = pd.DataFrame(list(metrics.items()), columns=["Metric", "Value"])
    out["Value"] = out["Value"].astype(str)  # kolom campuran -> aman untuk Streamlit/Arrow
    return out


def build_validation(df):
    checks = []

    def add_check(name, condition, detail=""):
        checks.append({"Validation": name, "Status": "PASS" if condition else "FAIL", "Detail": detail})

    dup = int(df.duplicated().sum())
    add_check("Duplicate Rows", dup == 0, f"{dup} duplicate rows")
    missing_po = int(df["PO#"].isna().sum())
    add_check("Missing PO#", missing_po == 0, f"{missing_po} missing")
    invalid_date = int(df["Order Date"].isna().sum())
    add_check("Missing/Invalid Order Date", invalid_date == 0, f"{invalid_date} missing/invalid")
    neg_order = int((df["Order Qty"] < 0).sum())
    add_check("Negative Order Qty", neg_order == 0, f"{neg_order} negative values")
    neg_add = int((df["Add Qty"] < 0).sum())
    add_check("Negative Add Qty", neg_add == 0, f"{neg_add} negative values")
    pct_valid = bool(df["Add %"].dropna().between(0, 1).all())
    add_check("Add % Range", pct_valid, "Non-null values should be between 0% and 100%")
    buildings = set(df["Building"].dropna().unique())
    add_check("Building Valid", buildings <= {"A", "B", "C", "D"}, f"Found: {sorted(buildings)}")

    standards = set(CATEGORY_MASTER.values())
    nonstd = int((~df["Label Category"].isin(standards) & df["Label Category"].notna()).sum())
    add_check("Standard Label Category", nonstd == 0, f"{nonstd} nonstandard/unknown rows")
    missing_cat = int(df["Label Category"].isna().sum())
    add_check("Label Category Filled", missing_cat == 0, f"{missing_cat} missing categories")

    diff = df["Add % Difference"].abs()
    mismatch = int((diff.notna() & (diff > 0.005)).sum())
    add_check("Add % Consistency", mismatch == 0, f"{mismatch} rows differ by >0.5 percentage points")

    validation = pd.DataFrame(checks)
    score = round((validation["Status"] == "PASS").mean() * 100, 2)
    return validation, score


def summarize(df, group_col):
    return (
        df.groupby(group_col, dropna=False, observed=True)
        .agg(
            Requests=("Request Flag", "sum"),
            Unique_PO=("PO#", "nunique"),
            Total_Order_Qty=("Order Qty", "sum"),
            Total_Add_Qty=("Add Qty", "sum"),
            Average_Add_Pct=("Add %", "mean"),
        )
        .reset_index()
        .sort_values("Total_Add_Qty", ascending=False)
    )


def build_metadata(df, score):
    """Metadata satu baris untuk dashboard."""
    if df["Order Date"].notna().any():
        period = (df["Order Date"].min().strftime("%d %b %Y")
                  + " - " + df["Order Date"].max().strftime("%d %b %Y"))
    else:
        period = ""
    return pd.DataFrame([{
        "Last Update": datetime.now().strftime("%Y-%m-%d"),
        "Data Period": period,
        "Total Rows": int(len(df)),
        "Total PO": int(df["PO#"].nunique()),
        "Total Add Qty": int(df["Add Qty"].sum()),
        "Average Add %": round(float(df["Add %"].mean()), 4),
        "Data Quality": f"{score:.0f}/100",
        "Version": VERSION,
    }])


@st.cache_data(show_spinner="Menjalankan ETL v1.1...")
def run_pipeline(file_bytes, source_name):
    raw = pd.read_excel(io.BytesIO(file_bytes))  # sheet pertama, sama dengan Colab
    if raw.empty:
        raise ValueError("File tidak memiliki data.")

    df = clean_data(raw)
    audit_df = build_audit(df)
    validation_df, score = build_validation(df)

    summaries = {
        "Summary_Category": summarize(df, "Label Category"),
        "Summary_Label_Item": summarize(df, "Label Item"),
        "Summary_Reason": summarize(df, "Reason"),
        "Summary_RCA": summarize(df, "RCA"),
        "Summary_Building": summarize(df, "Building"),
        "Summary_Line": summarize(df, "Building Line"),
        "Top_10_Model": summarize(df, "Model").head(10),
        "Summary_Country": summarize(df, "ShipTo Country Name"),
    }

    category_mapping = (
        df.groupby(["Label Item", "Label Category"], dropna=False)
        .size().reset_index(name="Rows")
        .sort_values(["Label Item", "Rows"], ascending=[True, False])
    )
    conflicts = (
        df.dropna(subset=["Label Item", "Label Category"])
        .groupby("Label Item")["Label Category"].nunique()
        .reset_index(name="Category Count")
    )
    conflicts = conflicts[conflicts["Category Count"] > 1]

    category_audit = (
        df["Label Category"].value_counts(dropna=False)
        .rename_axis("Label Category").reset_index(name="Rows")
    )
    category_audit["Is Standard"] = category_audit["Label Category"].isin(list(CATEGORY_MASTER.values()))

    metadata = build_metadata(df, score)

    return {
        "df": df, "audit": audit_df, "validation": validation_df, "score": score,
        "summaries": summaries, "mapping": category_mapping, "conflicts": conflicts,
        "category_audit": category_audit, "metadata": metadata,
    }


def build_excel_bytes(res):
    """Export multi-sheet, sama dengan output Colab."""
    from openpyxl.styles import Font, PatternFill, Alignment

    sheets = {
        "Fact_Special_Label": res["df"],
        "Audit_Report": res["audit"],
        "Validation_Report": res["validation"],
        **res["summaries"],
        "Label_Mapping": res["mapping"],
        "Category_Conflicts": res["conflicts"],
        "Category_Audit": res["category_audit"],
        "Metadata": res["metadata"],
    }
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for column_cells in ws.columns:
                letter = column_cells[0].column_letter
                max_len = max((len(str(c.value)) if c.value is not None else 0)
                              for c in column_cells[:500])
                ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 35)
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(fill_type="solid", fgColor="008C95")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 25
    return buf.getvalue()


# ================================================================
# 5. GOOGLE SHEETS
# ================================================================
def get_google_credentials():
    if "gdrive_service_account" not in st.secrets:
        raise RuntimeError("Secret [gdrive_service_account] belum ditemukan.")
    info = dict(st.secrets["gdrive_service_account"])
    return service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"])


def get_sheets_service():
    return build("sheets", "v4", credentials=get_google_credentials(), cache_discovery=False)


def get_config():
    drive_cfg = st.secrets.get("drive", {})
    spreadsheet_id = str(drive_cfg.get("target_file_id", "")).strip()
    worksheet_gid = str(drive_cfg.get("target_gid", "")).strip()
    if not spreadsheet_id:
        raise RuntimeError("drive.target_file_id belum diisi.")
    if not worksheet_gid:
        raise RuntimeError("drive.target_gid belum diisi. Isi dengan gid tab tujuan.")
    try:
        return spreadsheet_id, int(worksheet_gid)
    except ValueError:
        raise RuntimeError("drive.target_gid harus berupa angka.")


def get_spreadsheet_metadata(service, spreadsheet_id):
    return service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="spreadsheetId,properties(title),sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
    ).execute()


def resolve_target_tab(metadata, target_gid):
    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if int(props.get("sheetId", -1)) == target_gid:
            return props
    available = [f"{s['properties'].get('title')} (gid={s['properties'].get('sheetId')})"
                 for s in metadata.get("sheets", [])]
    raise RuntimeError(f"gid {target_gid} tidak ditemukan. Tab tersedia: " + "; ".join(available))


def quote_sheet_title(title):
    return "'" + title.replace("'", "''") + "'"


def to_sheet_value(value, as_text=False):
    if value is None or value is pd.NA:
        return ""
    if pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "item") and not isinstance(value, str):
        value = value.item()
    if isinstance(value, bool):
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else value
    if isinstance(value, int):
        return value
    text = str(value).strip()
    if as_text and text:
        return "'" + text  # apostrof = paksa teks di USER_ENTERED
    return text


def dataframe_to_values(df):
    headers = list(df.columns)
    text_idx = {i for i, c in enumerate(headers) if ID_AS_TEXT and c in ID_COLS}
    rows = [headers]
    for row in df.itertuples(index=False, name=None):
        rows.append([to_sheet_value(v, as_text=(i in text_idx)) for i, v in enumerate(row)])
    return rows


def ensure_grid_size(service, spreadsheet_id, props, n_rows, n_cols):
    grid = props.get("gridProperties", {})
    cur_rows, cur_cols = grid.get("rowCount", 0), grid.get("columnCount", 0)
    if cur_rows >= n_rows and cur_cols >= n_cols:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{
            "updateSheetProperties": {
                "properties": {
                    "sheetId": props["sheetId"],
                    "gridProperties": {"rowCount": max(cur_rows, n_rows),
                                       "columnCount": max(cur_cols, n_cols)},
                },
                "fields": "gridProperties.rowCount,gridProperties.columnCount",
            }
        }]},
    ).execute()


def replace_worksheet_values(service, spreadsheet_id, tab_props, values, chunk=5000, progress=None):
    """Clear tab, lalu tulis header + data bertahap (chunk baris)."""
    title = tab_props["title"]
    q = quote_sheet_title(title)
    n_rows, n_cols = len(values), len(values[0])

    ensure_grid_size(service, spreadsheet_id, tab_props, n_rows, n_cols)

    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id, range=q, body={}).execute()

    header, rows = values[0], values[1:]
    total_cells = 0
    n_chunks = max(1, -(-len(rows) // chunk))
    for k, i in enumerate(range(0, len(rows), chunk)):
        block = rows[i:i + chunk]
        if i == 0:
            data, start_row = [header] + block, 1
        else:
            data, start_row = block, i + 2  # baris 1 = header
        resp = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{q}!A{start_row}",
            valueInputOption="USER_ENTERED",
            body={"values": data},
        ).execute()
        total_cells += resp.get("updatedCells", 0)
        if progress:
            progress.progress((k + 1) / n_chunks)
    return {"updatedRows": n_rows, "updatedColumns": n_cols,
            "updatedCells": total_cells, "updatedRange": f"{title}!A1"}


# ================================================================
# 6. SIDEBAR
# ================================================================
with st.sidebar:
    st.header("Configuration")
    st.info("Target Spreadsheet ID dan target tab gid dibaca dari Streamlit Secrets.")
    st.caption("Input: Excel .xlsx (tab pertama)")
    st.caption("Output: Google Sheets native + Excel multi-sheet")
    st.caption(f"ETL version: {VERSION}")

# ================================================================
# 7. UPLOAD & ETL
# ================================================================
upload = st.file_uploader(
    "Upload file Excel Special Label (.xlsx)", type=["xlsx"],
    help="Workbook ini akan menggantikan isi nilai pada worksheet target.")
if not upload:
    st.info("Upload file Excel untuk memulai.")
    st.stop()

file_bytes = upload.getvalue()
file_hash = hashlib.sha256(file_bytes).hexdigest()
st.write(f"**File:** {upload.name} · **Size:** {len(file_bytes):,} bytes")

try:
    res = run_pipeline(file_bytes, upload.name)
except Exception as exc:
    st.error(f"Validasi gagal: {exc}")
    st.stop()

df = res["df"]
validation_df = res["validation"]
quality_score = res["score"]
known = set(CATEGORY_MASTER.values())
unknown_categories = sorted(str(v) for v in df["Label Category"].dropna().unique() if v not in known)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Rows", f"{len(df):,}")
m2.metric("Unique PO", f"{df['PO#'].nunique(dropna=True):,}")
m3.metric("Total Add Qty", f"{df['Add Qty'].sum():,.0f}")
m4.metric("Label Items", f"{df['Label Item'].nunique(dropna=True):,}")

st.success(f"ETL {VERSION} selesai. Schema output: {len(df.columns)} kolom.")
st.metric("Data Quality Score", f"{quality_score:.2f}/100")

st.subheader("Validation Report")
st.dataframe(validation_df, use_container_width=True)
if (validation_df["Status"] == "FAIL").any():
    st.warning("Ada validation check berstatus FAIL. Periksa report sebelum publish.")
st.caption(f"SHA-256 file upload: `{file_hash}`")

if unknown_categories:
    st.warning("Label Category di luar standard ditemukan: " + ", ".join(unknown_categories[:30]))
    st.caption("Kategori di luar standard dipertahankan apa adanya. Mohon review sebelum update.")

n_blank = int(df["Label Category"].isna().sum())
if n_blank:
    st.warning(f"Ada {n_blank:,} baris tanpa Label Category.")

st.subheader("Preview data")
st.dataframe(df.head(30), use_container_width=True)
st.caption("Field hasil ETL: " + ", ".join(c for c in ETL_COLUMNS if c in df.columns))

with st.expander("Audit Report"):
    st.dataframe(res["audit"], use_container_width=True)

with st.expander("Summary tables"):
    tabs = st.tabs(list(res["summaries"].keys()))
    for tab, (name, frame) in zip(tabs, res["summaries"].items()):
        with tab:
            st.dataframe(frame, use_container_width=True)

with st.expander("Label Mapping, Conflicts & Category Audit"):
    st.markdown("**Category Audit**")
    st.dataframe(res["category_audit"], use_container_width=True)
    st.markdown("**Category Conflicts**")
    st.dataframe(res["conflicts"], use_container_width=True)
    st.markdown("**Label Mapping (50 baris pertama)**")
    st.dataframe(res["mapping"].head(50), use_container_width=True)

st.download_button(
    "⬇️ Download hasil ETL (Excel multi-sheet)",
    data=build_excel_bytes(res),
    file_name="PGD_Special_Label_ETL_" + datetime.now().strftime("%Y%m%d_%H%M") + ".xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ================================================================
# 8. TARGET INSPECTION
# ================================================================
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
    st.error("Gagal membaca Google Sheets. Pastikan Google Sheets API aktif, "
             "Spreadsheet ID benar, dan spreadsheet sudah di-share ke service account.")
    st.code(str(exc))
    st.stop()
except Exception as exc:
    st.error(f"Konfigurasi/koneksi gagal: {exc}")
    st.stop()

# ================================================================
# 9. CONFIRM & WRITE
# ================================================================
st.warning(
    "PERHATIAN: Update akan menghapus nilai lama di tab target, lalu menulis header dan data "
    "baru mulai A1. Format sel yang sudah ada dipertahankan, tetapi kolom baru tidak otomatis "
    "mendapat format. Pastikan tab target khusus untuk dataset ini dan backup sudah tersedia."
)
confirm_target = st.text_input(
    f"Ketik UPDATE untuk menyetujui penggantian data pada tab: {target_title}",
    placeholder="UPDATE")
backup_confirm = st.checkbox("Saya sudah membuat backup dan memastikan tab target benar.")

missing_etl = [c for c in ETL_COLUMNS if c not in df.columns]
if missing_etl:
    st.error("Schema ETL tidak lengkap: " + ", ".join(missing_etl))

can_update = (confirm_target.strip().upper() == "UPDATE" and backup_confirm
              and not missing_etl and not df.empty)

if st.button("☁️ Update Google Sheets", type="primary",
             disabled=not can_update, use_container_width=True):
    try:
        values = dataframe_to_values(df)
        progress = st.progress(0.0)
        with st.spinner("Meng-update Google Sheets..."):
            response = replace_worksheet_values(
                sheets_service, spreadsheet_id, target_props, values, progress=progress)

        st.success("Google Sheets berhasil diperbarui.")
        st.write(f"**Updated range:** {response['updatedRange']}")
        st.write(f"**Updated rows:** {response['updatedRows']:,}")
        st.write(f"**Updated columns:** {response['updatedColumns']}")
        st.write(f"**Updated cells:** {response['updatedCells']:,}")
        st.markdown(f"[Buka Google Sheets target]"
                    f"(https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid={target_gid})")
        st.info("Perubahan di Looker Studio bergantung pada konektor, cache, dan refresh data source.")
    except HttpError as exc:
        st.error("Google Sheets API gagal saat menulis data. Tab target mungkin sudah "
                 "dikosongkan atau terisi sebagian; ulangi update atau pulihkan dari backup.")
        st.code(str(exc))
    except Exception as exc:
        st.error(f"Update gagal: {exc}. Tab target mungkin sudah dikosongkan atau terisi sebagian.")

st.divider()
st.caption("PGD Special Label Uploader · Gunakan backup dan akses aplikasi terbatas "
           "sebelum deployment untuk penggunaan operasional.")
