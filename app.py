import io
import hashlib
import re
from datetime import datetime, timezone

import numpy as np
import pandas as pd
import streamlit as st
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
