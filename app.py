import io, hashlib
import pandas as pd
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

st.set_page_config(page_title='PGD Special Label Uploader', page_icon='📊', layout='wide')
st.title('PGD Special Label Dashboard Uploader')
st.caption('Upload data terbaru dan update file target di Google Drive.')

REQUIRED_COLUMNS = ['Order Date','PO#','SO','Article No.','Model','Line','Building','Order Qty','Add Qty','Add %','Reason','Label Item','Label Category','ShipTo Country Name','LPD','PODD']
CATEGORY_MAP = {'SHOEBOX LABEL':'Shoebox Label','HANGTAG RFID':'Hangtag RFID','HANGTAG STANDART':'Hangtag Standart','SPECIAL LABEL':'Special Label'}
CATEGORY_ALIASES = {'HANGTAG STANDARD':'HANGTAG STANDART'}

def normalize_header(x):
    return ' '.join(str(x).replace('\xa0',' ').split()).strip()

def standardize_category(value):
    if pd.isna(value): return pd.NA
    value = ' '.join(str(value).replace('\xa0',' ').split()).strip()
    if not value: return pd.NA
    key = CATEGORY_ALIASES.get(value.upper(), value.upper())
    return CATEGORY_MAP.get(key, value)

def validate_upload(payload):
    df = pd.read_excel(io.BytesIO(payload), sheet_name=0)
    df.columns = [normalize_header(c) for c in df.columns]
    if df.columns.duplicated().any():
        raise ValueError(f'Header duplikat: {df.columns[df.columns.duplicated()].tolist()}')
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError('Kolom wajib tidak ditemukan: ' + ', '.join(missing))
    if df.empty: raise ValueError('Workbook tidak memiliki baris data.')
    df['Label Category'] = df['Label Category'].map(standardize_category)
    allowed = set(CATEGORY_MAP.values())
    unknown = sorted(str(v) for v in df['Label Category'].dropna().unique() if v not in allowed)
    add_qty = pd.to_numeric(df['Add Qty'].astype(str).str.replace(',','',regex=False), errors='coerce')
    stats = {'rows':len(df),'columns':len(df.columns),'po':df['PO#'].nunique(dropna=True),'add_qty':add_qty.sum(),
             'missing_category':int(df['Label Category'].isna().sum()),'unknown':unknown,
             'sha256':hashlib.sha256(payload).hexdigest()}
    return df, stats

def drive_service():
    info = dict(st.secrets['gdrive_service_account'])
    creds = service_account.Credentials.from_service_account_info(info, scopes=['https://www.googleapis.com/auth/drive'])
    return build('drive','v3',credentials=creds,cache_discovery=False)

def target_file_id():
    value = st.secrets.get('drive',{}).get('target_file_id','')
    if not value: raise RuntimeError('drive.target_file_id belum dikonfigurasi di Streamlit Secrets.')
    return str(value).strip()

with st.sidebar:
    st.subheader('Setup')
    st.info('Drive target dan service account dikonfigurasi melalui Streamlit Secrets.')
    st.caption('Upload format: .xlsx | Validasi: worksheet pertama')

upload = st.file_uploader('Upload file Excel Special Label terbaru', type=['xlsx'])
if upload:
    payload = upload.getvalue()
    st.write(f'**File:** {upload.name} · **Size:** {len(payload):,} bytes')
    try:
        df, metrics = validate_upload(payload)
        st.success('Validasi struktur berhasil.')
        a,b,c,d = st.columns(4)
        a.metric('Rows',f"{metrics['rows']:,}")
        b.metric('Columns',f"{metrics['columns']:,}")
        c.metric('Unique PO',f"{metrics['po']:,}")
        d.metric('Total Add Qty',f"{metrics['add_qty']:,.0f}" if pd.notna(metrics['add_qty']) else 'N/A')
        if metrics['missing_category']:
            st.warning(f"{metrics['missing_category']:,} baris tidak memiliki Label Category.")
        if metrics['unknown']:
            st.warning('Kategori di luar standard master: '+', '.join(metrics['unknown'][:30]))
        st.caption('Preview kategori sudah distandardisasi untuk pemeriksaan. Upload masih memakai workbook sumber asli.')
        st.dataframe(df.head(20),use_container_width=True)
        st.subheader('Ringkasan Label Category')
        st.dataframe(df['Label Category'].fillna('(Blank)').value_counts().rename_axis('Label Category').reset_index(name='Rows'),use_container_width=True)
        st.divider()
        st.warning('Update akan mengganti konten file target di Drive. File ID dipertahankan, tetapi konten lama tertimpa.')
        confirmed = st.checkbox('Saya sudah memeriksa file dan menyetujui penggantian file target.')
        if st.button('☁️ Update Google Drive',type='primary',disabled=not confirmed,use_container_width=True):
            try:
                service = drive_service()
                fid = target_file_id()
                target = service.files().get(fileId=fid,fields='id,name,mimeType,modifiedTime,webViewLink',supportsAllDrives=True).execute()
                acceptable = {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','application/octet-stream'}
                if target.get('mimeType') not in acceptable:
                    st.error('Target bukan XLSX yang dikenali. Operasi dibatalkan untuk melindungi Google Sheets/native file atau format lain.')
                    st.stop()
                st.write(f"**Target:** {target.get('name','(tanpa nama)')} · **Last modified:** {target.get('modifiedTime','Unknown')}")
                media = MediaIoBaseUpload(io.BytesIO(payload),mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',resumable=True)
                result = service.files().update(fileId=fid,media_body=media,fields='id,name,mimeType,modifiedTime,size,webViewLink',supportsAllDrives=True).execute()
                st.success('Konten file Google Drive berhasil diperbarui.')
                st.write(f"**File ID:** `{result.get('id')}` · **Modified:** {result.get('modifiedTime','N/A')}")
                if result.get('webViewLink'): st.markdown(f"[Buka file di Google Drive]({result['webViewLink']})")
                st.info('Refresh Looker Studio sesuai jenis konektor dan pengaturan cache/data freshness. Update Drive tidak menjamin tampilan instan.')
            except HttpError as e:
                st.error('Google Drive API gagal. Periksa Drive API, target file ID, dan izin Editor untuk service account.')
                st.code(str(e))
            except Exception as e: st.error(f'Update gagal: {e}')
    except Exception as e:
        st.error(f'Validasi gagal: {e}')
        st.info('Periksa header wajib dan worksheet pertama. Jika nama kolom atau header row berbeda, sesuaikan daftar REQUIRED_COLUMNS/logic pembacaan.')
else:
    st.info('Upload file .xlsx untuk memulai.')

st.divider()
st.caption('Production hardening yang disarankan: login/allowlist, audit log, backup/versioning, dan access review.')
