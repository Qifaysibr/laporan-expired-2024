#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Laporan Expired SBU - Python Cron Job
Sumber data: Supabase PostgreSQL
Output: File Excel (.xlsx) dengan 3 sheet [87, 60, 30] hari expired
Upload: Google Drive dengan folder per bulan
"""

import os
import sys
import json
from datetime import datetime, timedelta
from collections import defaultdict
import csv
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Database
import psycopg2
from psycopg2.extras import RealDictCursor

# Excel
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Google Drive
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import pickle

# ================== KONFIGURASI ==================
# Database Supabase (Direct PostgreSQL)
DB_CONFIG = {
    "host": os.getenv("DB_HOST", "your-db-host.supabase.co"),
    "port": os.getenv("DB_PORT", "5432"),
    "dbname": os.getenv("DB_NAME", "postgres"),
    "user": os.getenv("DB_USER", "postgres"),
    "password": os.getenv("DB_PASSWORD", "your-db-password"),
}

# Google Drive Configuration
GOOGLE_CREDENTIALS_PATH = os.getenv("GOOGLE_CREDENTIALS_PATH", "client_secrets.json")
TOKEN_PATH = os.getenv("TOKEN_PATH", "token.pickle")
ROOT_FOLDER_ID = os.getenv("GDRIVE_ROOT_FOLDER_ID", "1-wHPWaIRH_GyLQ8RxTTrPz129GQRrvE9")

# Kota Jabodetabek (diurutkan paling atas)
KOTA_JABODETABEK = ['bekasi', 'bogor', 'depok', 'jakarta timur']

# Interval laporan (hari)
INTERVALS = [87, 60, 30]

# Jendela penggabungan SBU (hari) & masa suppress dedup (hari)
WINDOW_DAYS = 7
SUPPRESS_DAYS = 7

# State file untuk dedup lintas laporan harian
STATE_PATH = os.path.join(os.getcwd(), "output", "state", "laporan_state.json")

# Timezone offset (GMT+7)
TZ_OFFSET = timedelta(hours=7)
# ================================================


def get_db_connection():
    """Buat koneksi ke PostgreSQL Supabase."""
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        return conn
    except Exception as e:
        print(f"[ERROR] Gagal koneksi ke database: {e}")
        sys.exit(1)


def get_conflicted_company_keys(conn):
    """
    Perusahaan yang punya kode subklasifikasi dengan status 'Pencabutan'
    DAN 'Disetujui' (SBU pengganti yang valid). Perusahaan ini tidak
    dimasukkan ke laporan. Return set kunci "<nama>_<telepon>".
    """
    query = """
    SELECT DISTINCT p.nama_badan_usaha, p.telepon, sd.id_subklasifikasi
    FROM perusahaan p
    JOIN sbu_details sd ON p.id = sd.perusahaan_id
    GROUP BY p.id, p.nama_badan_usaha, p.telepon, sd.id_subklasifikasi
    HAVING COUNT(*) FILTER (WHERE sd.status_sbu = 'Pencabutan') > 0
       AND COUNT(*) FILTER (WHERE sd.status_sbu = 'Disetujui') > 0
    """
    keys = set()
    with conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute(query)
        for row in cur.fetchall():
            keys.add(f"{row['nama_badan_usaha']}_{row['telepon']}")
    return keys


def load_state():
    """Muat state dedup dari file JSON. Return dict per interval."""
    if not os.path.exists(STATE_PATH):
        return {}
    try:
        with open(STATE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"[WARNING] Gagal membaca state file, mulai baru: {e}")
        return {}


def save_state(state):
    """Simpan state dedup ke file JSON."""
    try:
        os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
        with open(STATE_PATH, 'w', encoding='utf-8') as f:
            json.dump(state, f, indent=2)
    except Exception as e:
        print(f"[WARNING] Gagal menyimpan state file: {e}")


def get_gdrive_service():
    """Inisialisasi Google Drive API via OAuth2 (User Authentication)."""
    creds = None
    # Token.pickle menyimpan token akses dan refresh
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, 'rb') as token:
            creds = pickle.load(token)
    
    # Jika tidak ada token valid, minta user login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("[INFO] Refreshing Google Drive token...")
            creds.refresh(Request())
        else:
            if not os.path.exists(GOOGLE_CREDENTIALS_PATH):
                print(f"[ERROR] File {GOOGLE_CREDENTIALS_PATH} tidak ditemukan!")
                return None
            
            print("[INFO] Memulai proses login Google Drive di browser...")
            flow = InstalledAppFlow.from_client_secrets_file(
                GOOGLE_CREDENTIALS_PATH, 
                scopes=['https://www.googleapis.com/auth/drive']
            )
            creds = flow.run_local_server(port=0)
        
        # Simpan token untuk penggunaan berikutnya
        with open(TOKEN_PATH, 'wb') as token:
            pickle.dump(creds, token)

    try:
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        print(f"[WARNING] Gagal inisialisasi Google Drive: {e}")
        return None
proxy_get_gdrive_service = get_gdrive_service


def get_or_create_folder(service, parent_id, folder_name):
    """Cari folder berdasarkan nama, atau buat baru jika belum ada."""
    if not service:
        return None

    try:
        # Cek apakah parent_id ada dan bisa diakses (termasuk Shared Drives)
        service.files().get(fileId=parent_id, supportsAllDrives=True).execute()
    except Exception as e:
        print(f"[ERROR] Folder Utama (ROOT_FOLDER_ID) tidak ditemukan atau tidak bisa diakses: {e}")
        print(f"[TIP] Pastikan folder dengan ID {parent_id} sudah dishare ke email: python@daily-actvity.iam.gserviceaccount.com")
        return None

    query = f"mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and name='{folder_name}' and trashed=false"
    try:
        results = service.files().list(
            q=query, 
            spaces='drive', 
            fields='files(id, name)',
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        items = results.get('files', [])

        if items:
            return items[0]['id']
        else:
            metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_id]
            }
            folder = service.files().create(
                body=metadata, 
                fields='id',
                supportsAllDrives=True
            ).execute()
            return folder['id']
    except Exception as e:
        print(f"[ERROR] Gagal membuat/mencari folder bulan '{folder_name}': {e}")
        return None


def upload_to_drive(service, file_path, file_name, folder_id):
    """Upload file ke Google Drive folder tertentu."""
    if not service or not folder_id:
        print("[INFO] Google Drive service tidak tersedia, file tetap tersimpan lokal.")
        return None

    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaFileUpload(file_path, resumable=True)
    file = service.files().create(
        body=file_metadata, 
        media_body=media, 
        fields='id',
        supportsAllDrives=True
    ).execute()
    print(f"[SUCCESS] File diupload ke Google Drive: {file_name} (ID: {file.get('id')})")
    return file.get('id')


def fetch_expired_data(conn, start_date, end_date):
    """
    Ambil SBU perusahaan dengan tgl_expired dalam rentang [start_date, end_date).
    Rentang ini dipakai untuk menggabungkan SBU perusahaan yang habis dalam
    7 hari dari tanggal target. Perusahaan yang ditampilkan difilter di main()
    menjadi hanya yang expired tepat di start_date.
    Return list of dict.
    """
    query = """
    SELECT 
        p.id AS perusahaan_id,
        p.nama_badan_usaha,
        p.provinsi,
        p.kabupaten_kota,
        p.telepon,
        p.email,
        p.link_lpjk,
        p.tgl_akte,
        sd.id_izin,
        sd.id_subklasifikasi,
        sd.subklasifikasi,
        sd.kualifikasi,
        sd.tgl_expired,
        sd.status_sbu,
        sd.penerbit
    FROM perusahaan p
    JOIN sbu_details sd ON p.id = sd.perusahaan_id
    WHERE sd.tgl_expired >= %s AND sd.tgl_expired < %s
    ORDER BY p.nama_badan_usaha, sd.id_subklasifikasi
    """

    with conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute(query, (start_date, end_date))
        rows = cur.fetchall()
    return rows


def group_by_perusahaan(rows):
    """
    Group data per perusahaan (nama + telepon).
    Aggregate: sbu codes, subklasifikasi, kualifikasi, status, dll.
    """
    grouped = defaultdict(lambda: {
        'data': None,
        'sbu_codes': [],
        'subklasifikasi_list': [],
        'kualifikasi_list': [],
        'status_list': [],
        'id_izin_list': [],
        'tgl_expired_list': [],
        'count': 0
    })

    for row in rows:
        key = f"{row['nama_badan_usaha']}_{row['telepon']}"

        if grouped[key]['data'] is None:
            grouped[key]['data'] = row

        grouped[key]['sbu_codes'].append(row['id_subklasifikasi'] or '')
        grouped[key]['subklasifikasi_list'].append(row['subklasifikasi'] or '')
        grouped[key]['kualifikasi_list'].append(row['kualifikasi'] or '')
        grouped[key]['status_list'].append(row['status_sbu'] or '')
        grouped[key]['id_izin_list'].append(row['id_izin'] or '')
        grouped[key]['tgl_expired_list'].append(row['tgl_expired'])
        grouped[key]['count'] += 1

    return grouped


def generate_whatsapp_message(row_data, exp_date_formatted, days, sbu_codes, link_lpjk):
    """Generate pesan WhatsApp berdasarkan interval hari."""
    nama = row_data['nama_badan_usaha']
    
    if days == 87:
        # Template perkenalan (soft approach)
        return f"""Selamat siang Bapak/Ibu pimpinan {nama},

Semoga Bapak/Ibu dalam keadaan baik.
Berdasarkan data LPJK, masa berlaku SBU perusahaan Bapak/Ibu akan berakhir sekitar 3 bulan lagi (pada {exp_date_formatted}).
Saya Ibrahim dari PT. Esdea Assistance Management. Kami berfokus membantu perusahaan seperti perusahaan Bapak/Ibu agar urusan legalitas dan perizinan bisa berjalan lancar dan efisien.

Oleh karena itu, kami hadir untuk membantu tim Bapak/Ibu, baik untuk:
- Penerbitan dan perpanjangan SBU
- Sertifikasi Kompetensi Kerja (SKK)
- Sertifikasi ISO (Akreditasi dan Non Akreditasi)
- Sertifikasi SNI
- Sertifikasi K3, SIO, SILO
- Audit Akuntan Publik (LAI)
- Laporan keuangan dan pajak (Bulanan/tahunan)

Apakah saat ini ada kebutuhan terkait layanan di atas yang sedang menjadi prioritas perusahaan Bapak/Ibu?
Boleh info ke siapa saya bisa kirimkan company profile kami?

Ibrahim
PT Esdea Assistance Management
https://esdeassistance.com"""

    elif days == 60:
        # Template pengingat formal
        return f"""*[PENGINGAT MASA BERLAKU SBU {nama}]*

Yth. Pimpinan {nama},
Mohon izin menginformasikan status SBU Anda di sistem LPJK:

📄 SBU: {sbu_codes}
📅 Expired: {exp_date_formatted} (60 Hari Lagi)

Kami tim PT. ESDEA Assistance Management siap men-support proses perpanjangan segera agar administrasi proyek tidak terhambat

Apakah boleh saya bantu siapkan estimasi biaya dan waktu pengerjaannya?."""

    elif days == 30:
        # Template pengingat urgent
        return f"""*[Pengingat Masa Berlaku SBU {nama}]*

Yth. Pimpinan {nama},
Mohon izin menginformasikan status SBU Anda:

📄 SBU: {sbu_codes}
📅 Expired: {exp_date_formatted} (30 Hari Lagi)

Kami tim PT. ESDEA Assistance Management siap men-support proses perpanjangan segera agar administrasi proyek tidak terhambat

Apakah boleh saya bantu siapkan estimasi biaya dan waktu pengerjaannya?."""

    return ""


def get_email_template_content(days, data_dict):
    """Membaca file template email dan mengganti placeholder."""
    template_path = os.path.join(os.getcwd(), "template", f"email_{days}.html")
    if not os.path.exists(template_path):
        return None
        
    with open(template_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Replace placeholders
    for key, value in data_dict.items():
        placeholder = "{{" + str(key) + "}}"
        content = content.replace(placeholder, str(value))
        
    return content


def format_tanggal_indonesia(date_obj):
    """Format date ke 'dd MMMM yyyy' dalam Bahasa Indonesia."""
    month_names = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    return f"{date_obj.day} {month_names[date_obj.month - 1]} {date_obj.year}"


def clean_phone_number(tel):
    """Bersihkan nomor telepon untuk format WhatsApp."""
    if not tel:
        return None

    clean_tel = str(tel).replace(r"\D", "").strip()
    # Hapus semua non-digit
    clean_tel = ''.join(filter(str.isdigit, clean_tel))

    if clean_tel.startswith('0'):
        clean_tel = '62' + clean_tel[1:]
    elif clean_tel.startswith('8'):
        clean_tel = '62' + clean_tel
    elif clean_tel.startswith('62'):
        pass
    else:
        return None

    return clean_tel


def build_excel_report(today, all_intervals_data):
    """
    Bangun file Excel dengan 3 sheet.
    all_intervals_data: dict {90: [rows], 60: [rows], 30: [rows]}
    """
    wb = Workbook()
    # Hapus sheet default
    wb.remove(wb.active)

    # State dedup: perusahaan tidak muncul ulang dalam SUPPRESS_DAYS hari
    state = load_state()
    today_iso = today.date().isoformat()

    # Header sesuai target
    headers = [
        "Nama Perusahaan",
        "Kabupaten/Kota",
        "Telepon",
        "Email",
        "Link",
        "Jumlah KBLI 2020",
        "Kode Subklasifikasi",
        "Tanggal Ditetapkan",
        "Tanggal Habis Masa Berlaku",
        "Status",
        "Sisa Hari",
        "Tanggal Expired",
        "Subklasifikasi",
        "Kualifikasi",
        "ID Subklasifikasi",
        "Provinsi",
        "Link WhatsApp"
    ]

    # Buat folder output utama
    out_dir = os.path.join(os.getcwd(), "output")
    xlsx_dir = os.path.join(out_dir, "xlsx")
    csv_base_dir = os.path.join(out_dir, "csv")
    os.makedirs(xlsx_dir, exist_ok=True)
    os.makedirs(csv_base_dir, exist_ok=True)

    for days in INTERVALS:
        rows = all_intervals_data.get(days, [])
        sheet_name = f"{days}_Hari_Expired"
        ws = wb.create_sheet(title=sheet_name)

        # Folder CSV tersendiri untuk setiap interval
        interval_csv_dir = os.path.join(csv_base_dir, str(days))
        os.makedirs(interval_csv_dir, exist_ok=True)

        # Tulis header
        ws.append(headers)

        # Style header
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Lanjutkan proses jika ada data
        csv_kecil = []
        csv_smb = []
        shown_today = {}
        if rows:
            # Group data
            grouped = group_by_perusahaan(rows)

            # Prepare data untuk sorting: Jabodetabek di atas
            processed_rows = []
            for key, item in grouped.items():
                row_data = item['data']

                # Dedup: lewati perusahaan yang sudah tampil < SUPPRESS_DAYS lalu
                last_shown = state.get(str(days), {}).get(key)
                if last_shown:
                    last_dt = datetime.strptime(last_shown, "%Y-%m-%d").date()
                    if (today.date() - last_dt).days < SUPPRESS_DAYS:
                        continue

                # Aggregate
                sbu_codes = ", ".join(sorted(set(filter(None, item['sbu_codes']))))
                subklasifikasi = ", ".join(sorted(set(filter(None, item['subklasifikasi_list']))))
                kualifikasi = ", ".join(sorted(set(filter(None, item['kualifikasi_list']))))
                status = ", ".join(sorted(set(filter(None, item['status_list']))))
                id_izin = ", ".join(sorted(set(filter(None, item['id_izin_list']))))
                jumlah_kbli = item['count']

                # Tanggal paling awal di antara SBU yang digabung
                tgl_expired = min(item['tgl_expired_list'])
                tgl_expired_formatted = format_tanggal_indonesia(tgl_expired)
                tgl_akte = row_data['tgl_akte']
                tgl_akte_formatted = format_tanggal_indonesia(tgl_akte) if tgl_akte else ""

                # Cek Jabodetabek
                kota = str(row_data['kabupaten_kota'] or '').lower().strip()
                is_jabodetabek = any(k in kota for k in KOTA_JABODETABEK)

                # WhatsApp Link (Hanya untuk Jabodetabek dan bukan nomor kantor 021)
                tel = clean_phone_number(row_data['telepon'])
                wa_link = ""
                wa_url_raw = ""
                
                is_bukan_kantor = tel and not tel.startswith('6221')
                
                if tel and is_jabodetabek and is_bukan_kantor:
                    pesan = generate_whatsapp_message(
                        row_data,
                        tgl_expired_formatted,
                        days,
                        sbu_codes,
                        row_data['link_lpjk']
                    )
                    wa_url_raw = f"https://wa.me/{tel}?text={pesan.replace(chr(10), '%0A')}"
                    wa_link = f'=HYPERLINK("{wa_url_raw}", "Hubungi")'

                row_values = [
                    row_data['nama_badan_usaha'],
                    row_data['kabupaten_kota'],
                    row_data['telepon'],
                    row_data['email'],
                    row_data['link_lpjk'],
                    jumlah_kbli,
                    sbu_codes,           # Kode Subklasifikasi
                    tgl_akte_formatted,  # Tanggal Ditetapkan
                    tgl_expired,         # Tanggal Habis Masa Berlaku (as date)
                    status,
                    days,                # Sisa Hari
                    tgl_expired_formatted,
                    subklasifikasi,
                    kualifikasi,
                    id_izin,             # ID Subklasifikasi
                    row_data['provinsi'],
                    wa_link
                ]

                # Kategori kualifikasi: Kecil vs Spesialis/Menengah/Besar
                # Campuran (mis. "K, M") dianggap non-Kecil
                kualifikasi_set = set(filter(None, item['kualifikasi_list']))
                is_kecil = bool(kualifikasi_set) and kualifikasi_set <= {'K'}

                processed_rows.append({
                    'is_jabodetabek': is_jabodetabek,
                    'is_kecil': is_kecil,
                    'values': row_values,
                    'csv_values': row_values[:-1] + [wa_url_raw] # CSV simpan URL asli
                })
                shown_today[key] = today_iso

            # Sort: Jabodetabek dulu, lalu nama perusahaan
            processed_rows.sort(key=lambda x: (not x['is_jabodetabek'], x['values'][0] or ''))

            # Tulis ke sheet & list CSV
            for item in processed_rows:
                ws.append(item['values'])
                if item['is_kecil']:
                    csv_kecil.append(item['csv_values'])
                else:
                    csv_smb.append(item['csv_values'])

        # Simpan CSV untuk interval ini (dipisah berdasarkan kualifikasi)
        def _write_csv(path, data):
            with open(path, mode='w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(data)
            print(f"[SUCCESS] File CSV tersimpan: {path}")

        csv_kecil_path = os.path.join(interval_csv_dir, f"{days}-Hari-Expired-{today.strftime('%d-%m-%y')}-Kecil.csv")
        csv_smb_path = os.path.join(interval_csv_dir, f"{days}-Hari-Expired-{today.strftime('%d-%m-%y')}-SpesialisMenengahBesar.csv")
        _write_csv(csv_kecil_path, csv_kecil)
        _write_csv(csv_smb_path, csv_smb)

        # Update state dedup interval ini, pangkas entri tua
        interval_state = {k: v for k, v in state.get(str(days), {}).items()
                          if (today.date() - datetime.strptime(v, "%Y-%m-%d").date()).days < SUPPRESS_DAYS}
        interval_state.update(shown_today)
        state[str(days)] = interval_state

        # Format kolom Tanggal (index 9 dan 12) di Excel
        for r in range(2, ws.max_row + 1):
            # Tanggal Habis Masa Berlaku (kolom I = 9)
            cell = ws.cell(row=r, column=9)
            if isinstance(cell.value, datetime):
                cell.number_format = 'DD/MM/YYYY'

            # Link WhatsApp (kolom Q = 17) -> biru & underline
            cell_wa = ws.cell(row=r, column=17)
            if cell_wa.value and isinstance(cell_wa.value, str) and cell_wa.value.startswith('='):
                cell_wa.font = Font(color="0000FF", underline="single")

        # Auto resize kolom
        for col_idx in range(1, len(headers) + 1):
            max_length = 0
            column = get_column_letter(col_idx)
            for cell in ws[column]:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column].width = adjusted_width

    # Simpan file XLSX
    date_str = today.strftime("%d-%m-%y")
    file_name = f"v3-{date_str}-Laporan_Expired.xlsx"
    local_xlsx_path = os.path.join(xlsx_dir, file_name)

    wb.save(local_xlsx_path)
    save_state(state)
    print(f"[SUCCESS] File Excel tersimpan: {local_xlsx_path}")
    return local_xlsx_path, file_name


def main():
    print(f"[INFO] Memulai proses laporan expired: {datetime.now()}")

    # Hitung tanggal hari ini (GMT+7)
    today = datetime.now() + TZ_OFFSET
    today = today.replace(hour=0, minute=0, second=0, microsecond=0)

    # Koneksi DB
    conn = get_db_connection()
    print("[INFO] Koneksi database berhasil.")

    # Ambil data untuk masing-masing interval
    all_intervals_data = {}

    # Perusahaan dengan kode subklasifikasi berstatus 'Pencabutan' & 'Disetujui'
    # tidak dimasukkan ke laporan
    conflicted_keys = get_conflicted_company_keys(conn)
    if conflicted_keys:
        print(f"[INFO] {len(conflicted_keys)} perusahaan dengan konflik status SBU dilewati.")

    for days in INTERVALS:
        start_date = today + timedelta(days=days)
        end_date = start_date + timedelta(days=WINDOW_DAYS)
        print(f"[INFO] Mengambil data untuk interval {days} hari (target: {start_date.date()}, jendela penggabungan SBU s/d {end_date.date() - timedelta(days=1)})...")

        rows = fetch_expired_data(conn, start_date.date(), end_date.date())

        # Hanya tampilkan perusahaan yang setidaknya 1 SBU-nya expired TEPAT di tanggal
        # target. SBU lain dari perusahaan yang sama (habis dalam 7 hari ke depan) tetap
        # digabung menjadi satu baris oleh group_by_perusahaan.
        anchor = start_date.date()
        anchor_keys = {
            f"{r['nama_badan_usaha']}_{r['telepon']}"
            for r in rows if r['tgl_expired'] == anchor
        }
        rows = [r for r in rows if f"{r['nama_badan_usaha']}_{r['telepon']}" in anchor_keys]

        # Buang baris milik perusahaan yang punya konflik status
        if conflicted_keys:
            rows = [r for r in rows if f"{r['nama_badan_usaha']}_{r['telepon']}" not in conflicted_keys]

        all_intervals_data[days] = rows
        print(f"[INFO] Ditemukan {len(rows)} baris SBU untuk interval {days} hari (perusahaan expired tepat {anchor}).")

    conn.close()

    # Bangun Excel
    local_path, file_name = build_excel_report(today, all_intervals_data)

    # Upload ke Google Drive
    try:
        gdrive_service = get_gdrive_service()
        if gdrive_service:
            # Buat folder bulan
            month_names = [
                "Januari", "Februari", "Maret", "April", "Mei", "Juni",
                "Juli", "Agustus", "September", "Oktober", "November", "Desember"
            ]
            folder_name = f"{month_names[today.month - 1]} {today.year}"
            folder_id = get_or_create_folder(gdrive_service, ROOT_FOLDER_ID, folder_name)

            if folder_id:
                upload_to_drive(gdrive_service, local_path, file_name, folder_id)
            else:
                print("[WARNING] Upload Google Drive dilewati karena folder tujuan tidak ditemukan.")
    except Exception as e:
        print(f"[ERROR] Terjadi kesalahan saat upload ke Google Drive: {e}")

    print("[INFO] Proses selesai.")


if __name__ == "__main__":
    main()
