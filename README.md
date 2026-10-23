# PGD Special Label Dashboard Uploader

Streamlit app untuk mengganti isi file XLSX yang telah ada di Google Drive, sambil mempertahankan Drive File ID.

## 1) Setup Google Drive
1. Di Google Cloud Console, buat/pilih project.
2. Enable **Google Drive API**.
3. Buat Service Account dan JSON key.
4. Share file target Excel ke email Service Account sebagai **Editor**.
5. Ambil file ID dari URL Drive, contoh `https://drive.google.com/file/d/FILE_ID/view`.
6. Jangan commit atau membagikan key JSON di repo/chat.

## 2) Streamlit Community Cloud Secrets
App → Settings → Secrets. Isi TOML berikut dan ganti placeholder dengan nilai pada JSON key:

```toml
[drive]
target_file_id = "ISI_FILE_ID_TARGET"

[gdrive_service_account]
type = "service_account"
project_id = "PROJECT_ID"
private_key_id = "PRIVATE_KEY_ID"
private_key = """-----BEGIN PRIVATE KEY-----
PASTE_PRIVATE_KEY_DI_SINI
-----END PRIVATE KEY-----\n"""
client_email = "SERVICE_ACCOUNT_EMAIL"
client_id = "CLIENT_ID"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "SERVICE_ACCOUNT_CERT_URL"
universe_domain = "googleapis.com"
```

Salin nilai credential dengan teliti. Jangan memasukkan private key ke source code.

## 3) Deploy
1. Push `app.py` dan `requirements.txt` ke GitHub repository.
2. Streamlit Community Cloud → New app → pilih repo, branch, `app.py`.
3. Isi Secrets seperti bagian 2, lalu deploy.

## 4) Cara menggunakan
1. Upload workbook XLSX final.
2. Review row count, total Add Qty, preview, dan kategori.
3. Centang konfirmasi.
4. Klik **Update Google Drive**.

## Batasan penting
- Validasi dan preview membaca worksheet pertama.
- Kode mengunggah workbook asli. Kategori di preview distandardisasi, tetapi standardisasi tersebut **tidak ditulis ke file hasil upload**. Jalankan ETL/standardisasi terlebih dahulu jika output itu yang dibutuhkan.
- Target harus berupa XLSX; target Google Sheets native memerlukan Google Sheets API flow yang berbeda.
- File ID dipertahankan; konten lama diganti.
- Looker Studio refresh tergantung konektor, cache, dan data freshness settings, bukan selalu instan.
- Untuk produksi, tambahkan autentikasi/allowlist pengguna, audit log, backup/versioning, dan uji recovery.
