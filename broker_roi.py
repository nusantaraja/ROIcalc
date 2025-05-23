# -*- coding: utf-8 -*-
"""
Modifikasi skrip ROI Calculator untuk dijalankan di Streamlit,
menambahkan input administrasi, menghasilkan output PDF, mengunggah ke Google Drive,
menggunakan Streamlit Secrets (termasuk Base64), mencatat ke Google Sheets,
dan nomor proposal otomatis.
"""
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from datetime import datetime
import locale
import io
import base64
import json
import toml # Untuk membaca secrets jika dalam format dict/toml
from weasyprint import HTML
from jinja2 import Environment, FileSystemLoader, select_autoescape

# Import Google API libraries
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

# --- Konfigurasi Halaman Streamlit (Harus menjadi perintah st pertama) ---
st.set_page_config(layout="wide", page_title="Kalkulator ROI AI Voice Broker")

# --- Konfigurasi Awal Lainnya ---
PROVIDER_COMPANY_NAME = "MEDIA AI SOLUSI, group of PT. EKUITAS MEDIA INVESTAMA"
LOG_SHEET_NAME = "Log Proposal" # Nama sheet/tab di dalam Google Sheet

# Set locale ke Indonesian (Hapus warning jika gagal)
try:
    locale.setlocale(locale.LC_ALL, 'id_ID.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'id_ID')
    except locale.Error:
        # st.sidebar.warning("Locale Indonesia tidak ditemukan. Format angka mungkin tidak sesuai.") # Dihapus sesuai permintaan
        pass # Abaikan jika locale tidak bisa diset

# --- Fungsi Bantuan ---
def format_number_id(value, precision=2):
    """Format angka ke format Indonesia.
    Handles potential locale errors gracefully.
    """
    try:
        if isinstance(value, (int, float)):
            if value == float('inf') or value == float('-inf'):
                return "N/A"
            # Coba format dengan locale
            try:
                format_str = f"%.{precision}f"
                formatted_value = locale.format_string(format_str, value, grouping=True)
                # Ganti pemisah desimal default locale jika perlu (misal, jika locale id_ID pakai titik)
                if locale.localeconv()['decimal_point'] == '.':
                     formatted_value = formatted_value.replace('.', '#TEMP#').replace(',', '.').replace('#TEMP#', ',')
                return formatted_value
            except (locale.Error, KeyError):
                 # Fallback jika locale gagal: format manual
                 return f"{value:,.{precision}f}".replace(',', '#TEMP#').replace('.', ',').replace('#TEMP#', '.')
        return str(value) # Kembalikan sebagai string jika bukan angka
    except (TypeError, ValueError):
        return str(value)

def generate_pdf(data):
    """Menghasilkan PDF dari data menggunakan template Jinja2 dan WeasyPrint."""
    try:
        # Dapatkan direktori skrip saat ini
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            script_dir = os.getcwd() # Fallback jika __file__ tidak tersedia (misal, di REPL)

        env = Environment(
            loader=FileSystemLoader(script_dir),
            autoescape=select_autoescape(['html'])
        )
        env.filters['format_number'] = format_number_id
        template = env.get_template('template.html')
        html_out = template.render(data)

        if 'chart_path' not in data or not data['chart_path'].startswith('data:image'):
             data['chart_path'] = ''

        pdf_bytes = HTML(string=html_out, base_url=script_dir).write_pdf()
        return pdf_bytes
    except Exception as e:
        st.error(f"Error saat membuat PDF: {e}")
        return None

# --- Fungsi Google Drive & Sheets (Auth terpusat) ---

SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

def get_google_credentials(secrets):
    """Mendapatkan kredensial Service Account dari Streamlit Secrets (Base64 atau JSON) atau file upload."""
    credentials_info = None
    source = "Not Found"
    show_api_settings = True # Default tampilkan pengaturan jika ada masalah

    # Prioritas 1: Coba dari Secret Base64 (misal, google_service_account_b64)
    if 'google_service_account_b64' in secrets:
        try:
            b64_str = secrets['google_service_account_b64']
            if isinstance(b64_str, str) and b64_str:
                decoded_bytes = base64.b64decode(b64_str)
                decoded_str = decoded_bytes.decode('utf-8')
                credentials_info = json.loads(decoded_str)
                source = "Streamlit Secrets (Base64)"
                # st.sidebar.success("Kredensial Service Account ditemukan di Secrets (Base64).") # Kurangi pesan sukses
                show_api_settings = False # Sembunyikan jika sukses dari secrets
            else:
                st.sidebar.warning("Secret 'google_service_account_b64' ditemukan tapi kosong atau bukan string.")
        except base64.binascii.Error:
            st.sidebar.error("Gagal decode Base64 dari secret 'google_service_account_b64'. Pastikan string Base64 valid.")
            return None, "Error", show_api_settings
        except json.JSONDecodeError:
            st.sidebar.error("Gagal parse JSON setelah decode Base64 dari secret 'google_service_account_b64'.")
            return None, "Error", show_api_settings
        except Exception as e:
            st.sidebar.error(f"Error membaca secret Base64 'google_service_account_b64': {e}")
            return None, "Error", show_api_settings

    # Prioritas 2: Coba dari Secret JSON/TOML biasa (misal, google_service_account)
    if credentials_info is None and 'google_service_account' in secrets:
        try:
            secret_content = secrets['google_service_account']
            if isinstance(secret_content, str):
                credentials_info = json.loads(secret_content)
            elif isinstance(secret_content, dict) or isinstance(secret_content, toml.TomlDecoder):
                 credentials_info = dict(secret_content)
            else:
                 st.sidebar.warning("Format secret 'google_service_account' tidak dikenali (harus string JSON atau TOML).")

            if credentials_info: # Hanya jika berhasil di-parse
                source = "Streamlit Secrets (JSON/TOML)"
                # st.sidebar.success("Kredensial Service Account ditemukan di Secrets (JSON/TOML).") # Kurangi pesan sukses
                show_api_settings = False # Sembunyikan jika sukses dari secrets

        except json.JSONDecodeError:
            st.sidebar.warning("Gagal membaca secret 'google_service_account' sebagai JSON/TOML. Mencoba file upload...")
        except Exception as e:
            st.sidebar.warning(f"Error membaca secret 'google_service_account': {e}. Mencoba file upload...")

    # Prioritas 3: Fallback ke File Upload
    if credentials_info is None:
        uploaded_key_file = st.sidebar.file_uploader("Unggah File Kunci JSON Service Account", type=['json'], help="Jika tidak menggunakan Streamlit Secrets. Unduh dari Google Cloud Console.")
        if uploaded_key_file is not None:
            try:
                stringio = io.StringIO(uploaded_key_file.getvalue().decode("utf-8"))
                credentials_info = json.load(stringio)
                source = "File Upload"
                st.sidebar.success("File kunci JSON berhasil dibaca.")
                show_api_settings = True # Tetap tampilkan jika via upload
            except json.JSONDecodeError:
                st.sidebar.error("File kunci JSON tidak valid.")
                return None, "Error", True
            except Exception as e:
                st.sidebar.error(f"Error membaca file kunci: {e}")
                return None, "Error", True

    # Validasi format kredensial jika ditemukan dari sumber manapun
    if credentials_info:
        required_keys = ("type", "project_id", "private_key_id", "private_key", "client_email", "client_id", "auth_uri", "token_uri")
        if not all(k in credentials_info for k in required_keys):
            st.sidebar.error(f"Format kredensial dari '{source}' tidak lengkap. Pastikan file/secret JSON Service Account yang benar digunakan.")
            return None, "Error", True
        return credentials_info, source, show_api_settings
    else:
        # Jika tidak ada secret dan tidak ada file diupload
        if source == "Not Found":
             st.sidebar.info("Kredensial Google Service Account tidak ditemukan.")
        return None, "Not Found", True

def get_gdrive_service(credentials_info):
    """Membuat service Google Drive."""
    try:
        credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
        service = build('drive', 'v3', credentials=credentials)
        return service
    except ValueError as ve:
        st.error(f"Error memproses kredensial untuk Google Drive: {ve}.")
        return None
    except Exception as e:
        st.error(f"Gagal membuat service Google Drive: {e}")
        return None

def get_gsheets_service(credentials_info):
    """Membuat service Google Sheets."""
    try:
        credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        return service
    except ValueError as ve:
        st.error(f"Error memproses kredensial untuk Google Sheets: {ve}.")
        return None
    except Exception as e:
        st.error(f"Gagal membuat service Google Sheets: {e}")
        return None

def get_next_proposal_number(service, sheet_id):
    """Mendapatkan nomor proposal berikutnya dari Google Sheet."""
    try:
        # Asumsi: Kolom B (indeks 1) berisi nomor proposal
        # Asumsi: Data dimulai dari baris 2 (baris 1 header)
        # Baca kolom B saja untuk efisiensi
        range_name = f"{LOG_SHEET_NAME}!B2:B"
        result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_name).execute()
        values = result.get('values', [])

        today_prefix = "PROP-" + datetime.now().strftime("%y%m%d") + "-"
        last_number = 0

        if values:
            # Cari nomor terakhir untuk tanggal hari ini
            for row in reversed(values):
                if row and row[0].startswith(today_prefix):
                    try:
                        last_number = int(row[0].split('-')[-1])
                        break
                    except (IndexError, ValueError):
                        continue # Abaikan format yang salah

        next_number = last_number + 1
        return f"{today_prefix}{next_number:03d}"

    except HttpError as error:
        st.warning(f"Error saat membaca Google Sheet untuk nomor proposal: {error}. Menggunakan format default.")
        return "PROP-" + datetime.now().strftime("%y%m%d") + "-001"
    except Exception as e:
        st.warning(f"Error tidak terduga saat mendapatkan nomor proposal: {e}. Menggunakan format default.")
        return "PROP-" + datetime.now().strftime("%y%m%d") + "-001"

def log_to_gsheet(service, sheet_id, log_data):
    """Mencatat data proposal ke Google Sheet."""
    try:
        # Asumsi urutan kolom: Timestamp, Proposal No, Agent Name, Agent Email, Agent Phone, Prospect Name, Prospect Location, GDrive Link
        values = [
            [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                log_data.get('proposal_number', ''),
                log_data.get('agent_name', ''),
                log_data.get('agent_email', ''),
                log_data.get('agent_phone', ''),
                log_data.get('prospect_name', ''),
                log_data.get('prospect_location', ''),
                log_data.get('gdrive_link', '')
            ]
        ]
        body = {
            'values': values
        }
        result = service.spreadsheets().values().append(
            spreadsheetId=sheet_id,
            range=f"{LOG_SHEET_NAME}!A1", # Append akan mencari baris kosong pertama
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body=body).execute()
        # st.success(f"Proposal berhasil dicatat ke Google Sheet.") # Kurangi pesan sukses
        return True
    except HttpError as error:
        st.error(f"Error saat mencatat ke Google Sheet: {error}. Pastikan Service Account memiliki akses Editor.")
        return False
    except Exception as e:
        st.error(f"Error tidak terduga saat mencatat ke Google Sheet: {e}")
        return False

def find_or_create_folder(service, folder_name, parent_folder_id):
    """Mencari/membuat folder di Google Drive."""
    # (Kode fungsi ini tetap sama)
    try:
        safe_folder_name = "".join(c for c in folder_name if c.isalnum() or c in (' ', '_', '-')).strip()
        if not safe_folder_name:
             safe_folder_name = "Prospek Tanpa Nama"
        query = f"name='{safe_folder_name}' and mimeType='application/vnd.google-apps.folder' and '{parent_folder_id}' in parents and trashed=false"
        response = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        folders = response.get('files', [])
        if folders:
            return folders[0].get('id')
        else:
            st.info(f"Membuat folder GDrive: '{safe_folder_name}'...")
            file_metadata = {'name': safe_folder_name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_folder_id]}
            folder = service.files().create(body=file_metadata, fields='id').execute()
            st.success(f"Folder '{safe_folder_name}' berhasil dibuat.")
            return folder.get('id')
    except HttpError as error:
        st.error(f"Error GDrive (Folder): {error}. Pastikan Service Account memiliki akses Editor ke folder induk.")
        return None
    except Exception as e:
        st.error(f"Error GDrive (Folder): {e}")
        return None

def upload_to_drive(service, pdf_bytes, filename, prospect_folder_id):
    """Mengunggah file PDF ke Google Drive."""
    # (Kode fungsi ini tetap sama)
    gdrive_link = None
    file_id = None
    try:
        file_metadata = {'name': filename, 'parents': [prospect_folder_id]}
        media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype='application/pdf', resumable=True)
        request = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink')
        response = None
        progress_bar = st.progress(0, text="Mengunggah ke Google Drive...")
        while response is None:
            status, response = request.next_chunk()
            if status:
                progress = int(status.progress() * 100)
                progress_bar.progress(progress, text=f"Mengunggah... {progress}%")
        progress_bar.empty()
        file_id = response.get('id')
        gdrive_link = response.get('webViewLink')
        st.success(f"File '{filename}' berhasil diunggah.")
        st.markdown(f"[Lihat File di Google Drive]({gdrive_link})", unsafe_allow_html=True)
    except HttpError as error:
        st.error(f"Error GDrive (Upload): {error}")
        if 'progress_bar' in locals() and hasattr(progress_bar, 'empty'): progress_bar.empty()
    except Exception as e:
        st.error(f"Error GDrive (Upload): {e}")
        if 'progress_bar' in locals() and hasattr(progress_bar, 'empty'): progress_bar.empty()
    return file_id, gdrive_link # Kembalikan ID dan Link

# --- Judul dan Deskripsi Aplikasi ---
st.title("📊 Kalkulator ROI Solusi AI Voice untuk Broker Forex")
st.markdown("Masukkan data operasional dan asumsi untuk menghitung potensi ROI, lalu hasilkan proposal PDF dan simpan ke Google Drive.")

# --- Dapatkan Kredensial & ID dari Secrets --- 
secrets = st.secrets
credentials_info, cred_source, show_api_settings_sidebar = get_google_credentials(secrets)
gdrive_parent_folder_id_secret = secrets.get("gdrive_parent_folder_id")
google_sheet_id_secret = secrets.get("google_sheet_id")

# --- Dapatkan Nomor Proposal Berikutnya (jika memungkinkan) ---
next_proposal_num = "PROP-" + datetime.now().strftime("%y%m%d") + "-XXX"
if credentials_info and google_sheet_id_secret:
    gsheets_service = get_gsheets_service(credentials_info)
    if gsheets_service:
        next_proposal_num = get_next_proposal_number(gsheets_service, google_sheet_id_secret)
    else:
        st.sidebar.warning("Gagal terhubung ke Google Sheets untuk mendapatkan nomor proposal.")

# --- Input Data --- 
with st.sidebar:
    st.header("⚙️ Input Data")

    # Informasi Agent/Marketing
    st.subheader("Informasi Agent/Marketing")
    agent_name = st.text_input("Nama Agent/Marketing", "")
    agent_email = st.text_input("Email Agent/Marketing", "")
    agent_phone = st.text_input("No. HP/WA Agent/Marketing", "")

    # Informasi Proposal & Prospek
    st.subheader("Informasi Proposal & Prospek")
    # Tampilkan nomor proposal tapi disable inputnya
    st.text_input("Nomor Proposal (Otomatis)", value=next_proposal_num, disabled=True)
    prospect_name = st.text_input("Nama Prospek (Broker Forex)", "PT Contoh Broker")
    prospect_location = st.text_input("Lokasi Prospek", "Jakarta")

    # Metrik Operasional
    st.subheader("Metrik Operasional Saat Ini")
    cs_staff = st.number_input("Jumlah staf CS", min_value=1, value=10)
    avg_monthly_salary = st.number_input("Gaji bulanan/staf (IDR)", min_value=0, value=7000000, step=100000, format="%d")
    overhead_multiplier = st.number_input("Pengali overhead", min_value=1.0, value=1.3, step=0.1)
    usd_conversion_rate = st.number_input("Kurs IDR ke USD", min_value=1000, value=15500, step=100, format="%d")
    monthly_inquiries = st.number_input("Pertanyaan/bulan", min_value=0, value=5000, step=100, format="%d")
    avg_handling_time = st.number_input("Waktu penanganan/pertanyaan (menit)", min_value=0.0, value=5.0, step=0.5)
    avg_monthly_clients = st.number_input("Klien aktif/bulan", min_value=0, value=1000, step=50, format="%d")
    avg_monthly_client_value = st.number_input("Pendapatan/klien/bulan (USD)", min_value=0.0, value=50.0, step=5.0)
    current_retention_rate = st.number_input("Loyalitas klien tahunan (%) ", min_value=0.0, max_value=100.0, value=85.0, step=1.0)

    # Investasi Solusi AI Voice
    st.subheader("Investasi Solusi AI Voice")
    implementation_cost = st.number_input("Biaya implementasi (USD)", min_value=0.0, value=10000.0, step=1000.0)
    annual_subscription = st.number_input("Biaya langganan tahunan (USD)", min_value=0.0, value=5000.0, step=500.0)

    # Asumsi Dampak
    st.subheader("Asumsi Dampak Solusi AI Voice")
    automation_rate = st.slider("Otomatisasi pertanyaan (%) ", 0, 100, 75)
    staff_reduction = st.slider("Pengurangan staf CS (%) ", 0, 100, 35)
    retention_improvement = st.slider("Peningkatan loyalitas klien (%) ", 0.0, 20.0, 7.5, 0.5)
    handling_time_improvement = st.slider("Peningkatan waktu penanganan (%) ", 0, 100, 25)

    # --- Pengaturan Google API (Hanya tampil jika perlu) ---
    gdrive_parent_folder_id_input = None
    google_sheet_id_input = None
    if show_api_settings_sidebar:
        st.subheader("☁️ Pengaturan Google API (Manual)")
        # Tampilkan input ID Folder hanya jika tidak ada di secrets
        if not gdrive_parent_folder_id_secret:
            gdrive_parent_folder_id_input = st.text_input("ID Folder Induk Google Drive", "", help="Masukkan jika tidak diset di Streamlit Secrets.")
        else:
            st.info("ID Folder Induk GDrive ditemukan di Streamlit Secrets.")
        # Tampilkan input ID Sheet hanya jika tidak ada di secrets
        if not google_sheet_id_secret:
            google_sheet_id_input = st.text_input("ID Google Sheet (untuk Log)", "", help="Masukkan jika tidak diset di Streamlit Secrets.")
        else:
            st.info("ID Google Sheet ditemukan di Streamlit Secrets.")
        # Opsi upload kredensial sudah ditangani di get_google_credentials

    # Tombol Kalkulasi
    st.divider()
    calculate_button = st.button("🚀 Hitung ROI, Buat PDF & Unggah")

# --- Kalkulasi & Output --- 
if calculate_button:
    # Tentukan ID Folder & Sheet yang akan digunakan (prioritaskan secrets)
    gdrive_parent_folder_id = gdrive_parent_folder_id_secret or gdrive_parent_folder_id_input
    google_sheet_id = google_sheet_id_secret or google_sheet_id_input

    # Validasi input dasar
    if not agent_name or not agent_email or not agent_phone:
        st.sidebar.error("Harap isi semua informasi Agent/Marketing.")
        st.stop()
    if not prospect_name:
        st.sidebar.error("Harap isi Nama Prospek.")
        st.stop()

    # Validasi Kredensial dan ID jika diperlukan upload/log
    trigger_gdrive_upload = False
    trigger_gsheet_log = False
    show_upload_log_section = False # Flag untuk menampilkan section di area utama

    if credentials_info:
        if gdrive_parent_folder_id:
            trigger_gdrive_upload = True
            show_upload_log_section = True
        else:
            st.warning("ID Folder Induk Google Drive tidak ditemukan. PDF tidak akan diunggah.")

        if google_sheet_id:
            trigger_gsheet_log = True
            show_upload_log_section = True
        else:
            st.warning("ID Google Sheet tidak ditemukan. Proposal tidak akan dicatat.")
    elif cred_source != "Error":
        st.warning("Kredensial Google tidak ditemukan. PDF tidak akan diunggah atau dicatat.")

    st.header("📈 Hasil Analisis ROI")
    with st.spinner("Melakukan kalkulasi ROI..."):
        # --- Kalkulasi Inti (sama seperti sebelumnya) ---
        avg_annual_salary = avg_monthly_salary * 12
        avg_annual_salary_usd = avg_annual_salary / usd_conversion_rate if usd_conversion_rate else 0
        avg_client_value = avg_monthly_client_value * 12
        current_annual_labor_cost_usd = cs_staff * avg_annual_salary_usd * overhead_multiplier
        current_annual_labor_cost_idr = current_annual_labor_cost_usd * usd_conversion_rate
        current_inquiries_per_year = monthly_inquiries * 12
        current_handling_hours = (current_inquiries_per_year * avg_handling_time) / 60
        automated_inquiries = current_inquiries_per_year * (automation_rate / 100)
        remaining_manual_inquiries = current_inquiries_per_year - automated_inquiries
        new_handling_time = avg_handling_time * (1 - handling_time_improvement / 100)
        new_handling_hours = (remaining_manual_inquiries * new_handling_time) / 60
        new_staff_count = round(cs_staff * (1 - staff_reduction / 100))
        if new_handling_hours > 0 and new_staff_count == 0:
            new_staff_count = 1
        new_annual_labor_cost_usd = new_staff_count * avg_annual_salary_usd * overhead_multiplier
        new_annual_labor_cost_idr = new_annual_labor_cost_usd * usd_conversion_rate
        labor_savings_usd = current_annual_labor_cost_usd - new_annual_labor_cost_usd
        labor_savings_idr = labor_savings_usd * usd_conversion_rate
        new_retention_rate = min(100.0, current_retention_rate + retention_improvement)
        avg_annual_clients = avg_monthly_clients
        current_retained_clients = avg_annual_clients * (current_retention_rate / 100)
        new_retained_clients = avg_annual_clients * (new_retention_rate / 100)
        additional_retained_clients = new_retained_clients - current_retained_clients
        retention_revenue_impact = additional_retained_clients * avg_client_value
        total_annual_savings_usd = labor_savings_usd + retention_revenue_impact
        total_annual_savings_idr = labor_savings_idr + (retention_revenue_impact * usd_conversion_rate)
        first_year_net_usd = total_annual_savings_usd - implementation_cost - annual_subscription
        subsequent_years_net_usd = total_annual_savings_usd - annual_subscription
        total_first_year_investment = implementation_cost + annual_subscription
        total_three_year_investment = implementation_cost + annual_subscription * 3
        first_year_roi = (first_year_net_usd / total_first_year_investment * 100) if total_first_year_investment > 0 else float('inf')
        three_year_net_benefit = first_year_net_usd + subsequent_years_net_usd * 2
        three_year_roi = (three_year_net_benefit / total_three_year_investment * 100) if total_three_year_investment > 0 else float('inf')
        monthly_savings_usd = total_annual_savings_usd / 12 if total_annual_savings_usd else 0
        total_investment_usd = implementation_cost + annual_subscription
        payback_period = (total_investment_usd / monthly_savings_usd) if monthly_savings_usd > 0 else float('inf')
        years = range(1, 6)
        costs = [(implementation_cost + annual_subscription) if year == 1 else annual_subscription for year in years]
        benefits = [total_annual_savings_usd for _ in years]
        net_benefits = [benefits[i] - costs[i] for i in range(len(years))]
        cumulative_net = np.cumsum(net_benefits).tolist()
        five_year_net_benefit = cumulative_net[-1] if cumulative_net else 0
        five_year_projection_data = []
        for i in range(len(years)):
            five_year_projection_data.append({
                'year': years[i],
                'cost': costs[i],
                'benefit': benefits[i],
                'net_benefit': net_benefits[i],
                'cumulative_net': cumulative_net[i]
            })

    # --- Generate Chart (sama seperti sebelumnya) ---
    st.subheader("📊 Grafik Analisis")
    chart_buffer = io.BytesIO()
    chart_data_uri = None
    try:
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        plt.style.use('seaborn-v0_8-whitegrid')
        labels_cost = ['Saat Ini', 'Dengan AI Voice']
        cs_costs_idr = [current_annual_labor_cost_idr, new_annual_labor_cost_idr]
        colors_cost = ['#3A86FF', '#8338EC']
        bars = ax1.bar(labels_cost, cs_costs_idr, color=colors_cost)
        ax1.set_title('Perbandingan Biaya Tahunan (IDR)', fontsize=12)
        ax1.set_ylabel('Biaya (IDR)', fontsize=10)
        ax1.tick_params(axis='x', labelsize=10)
        ax1.tick_params(axis='y', labelsize=10)
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_number_id(x, 0)))
        for bar in bars:
            yval = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2.0, yval * 1.01, f'IDR {format_number_id(yval, 0)}', va='bottom', ha='center', fontsize=9)
        ax2.plot(years, cumulative_net, marker='o', linewidth=2, color='#FF006E', label='Manfaat Bersih Kumulatif')
        ax2.set_title('Manfaat Bersih Kumulatif 5 Tahun (USD)', fontsize=12)
        ax2.set_xlabel('Tahun', fontsize=10)
        ax2.set_ylabel('USD', fontsize=10)
        ax2.grid(True, linestyle='--', alpha=0.6)
        ax2.set_xticks(years)
        ax2.tick_params(axis='x', labelsize=10)
        ax2.tick_params(axis='y', labelsize=10)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${format_number_id(x, 0)}'))
        ax2.axhline(0, color='grey', linewidth=0.8, linestyle='--')
        for i, value in enumerate(cumulative_net):
             ax2.text(years[i], value, f'${format_number_id(value, 0)}', ha='center', va='bottom', fontsize=9)
        plt.tight_layout(pad=2.0)
        plt.savefig(chart_buffer, format='png', dpi=300, bbox_inches='tight')
        plt.close(fig)
        chart_buffer.seek(0)
        chart_base64 = base64.b64encode(chart_buffer.read()).decode()
        chart_data_uri = f"data:image/png;base64,{chart_base64}"
        st.image(chart_buffer, caption="Grafik Analisis ROI", use_container_width=True)
    except Exception as e:
        st.error(f"Gagal membuat grafik: {e}")

    # --- Tampilkan Ringkasan di Streamlit (sama seperti sebelumnya) ---
    st.subheader("Ringkasan Hasil Utama")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="ROI Tahun Pertama", value=f"{format_number_id(first_year_roi)}%")
        st.metric(label="ROI Tiga Tahun", value=f"{format_number_id(three_year_roi)}%")
        st.metric(label="Periode Pengembalian", value=f"{format_number_id(payback_period, 1)} bulan" if payback_period != float('inf') else "Tidak Tercapai")
    with col2:
        st.metric(label="Total Penghematan Tahunan (USD)", value=f"$ {format_number_id(total_annual_savings_usd)}")
        st.metric(label="Manfaat Bersih 5 Tahun (USD)", value=f"$ {format_number_id(five_year_net_benefit)}")
        st.metric(label="Pengurangan Staf", value=f"{cs_staff - new_staff_count} orang ({staff_reduction}%)")
    with col3:
        st.metric(label="Otomatisasi Pertanyaan", value=f"{format_number_id(automated_inquiries, 0)} ({automation_rate}%)")
        st.metric(label="Peningkatan Loyalitas Klien", value=f"+{retention_improvement}% (menjadi {new_retention_rate}%)")
        st.metric(label="Penghematan Biaya Tahunan (IDR)", value=f"Rp {format_number_id(labor_savings_idr)}")

    # --- Kesimpulan Teks (sama seperti sebelumnya) ---
    st.subheader("📝 Kesimpulan")
    # (Logika kesimpulan tetap sama)
    if first_year_roi != float('inf') and first_year_roi > 50 and payback_period < 18:
        conclusion_text = f"Implementasi Solusi AI Voice untuk **{prospect_name}** sangat direkomendasikan. Dengan ROI tahun pertama **{format_number_id(first_year_roi)}%** dan periode pengembalian hanya **{format_number_id(payback_period, 1)} bulan**, investasi ini menawarkan nilai finansial yang sangat signifikan dan cepat."
    elif first_year_roi != float('inf') and first_year_roi > 0:
        conclusion_text = f"Implementasi Solusi AI Voice untuk **{prospect_name}** direkomendasikan. ROI tahun pertama sebesar **{format_number_id(first_year_roi)}%** dan periode pengembalian **{format_number_id(payback_period, 1)} bulan** menunjukkan potensi pengembalian investasi yang solid dalam jangka menengah."
    elif three_year_roi != float('inf') and three_year_roi > 0:
         conclusion_text = f"Implementasi Solusi AI Voice untuk **{prospect_name}** patut dipertimbangkan. Meskipun ROI tahun pertama mungkin belum positif ({format_number_id(first_year_roi)}%), ROI tiga tahun sebesar **{format_number_id(three_year_roi)}%** mengindikasikan potensi keuntungan jangka panjang yang menarik."
    else:
        conclusion_text = f"Berdasarkan data dan asumsi saat ini, ROI untuk implementasi Solusi AI Voice bagi **{prospect_name}** terlihat kurang menarik ({format_number_id(first_year_roi)}% ROI tahun pertama). Perlu evaluasi lebih lanjut terhadap asumsi atau potensi manfaat lain sebelum melanjutkan."
    st.markdown(conclusion_text)

    # --- Persiapan Data untuk PDF & Log ---
    current_time = datetime.now()
    # Gunakan nomor proposal yang didapat dari GSheet atau default
    final_proposal_number = next_proposal_num
    pdf_data = {
        'proposal_number': final_proposal_number,
        'analysis_date': current_time.strftime('%d %B %Y'),
        'prospect_name': prospect_name,
        'prospect_location': prospect_location,
        'provider_company_name': PROVIDER_COMPANY_NAME,
        'agent_name': agent_name,
        'agent_email': agent_email,
        'agent_phone': agent_phone,
        # Data kalkulasi lainnya (sama seperti sebelumnya)
        'cs_staff': cs_staff,
        'current_annual_labor_cost_idr': current_annual_labor_cost_idr,
        'current_annual_labor_cost_usd': current_annual_labor_cost_usd,
        'current_inquiries_per_year': current_inquiries_per_year,
        'current_handling_hours': current_handling_hours,
        'current_retention_rate': current_retention_rate,
        'new_staff_count': new_staff_count,
        'new_annual_labor_cost_idr': new_annual_labor_cost_idr,
        'new_annual_labor_cost_usd': new_annual_labor_cost_usd,
        'automated_inquiries': automated_inquiries,
        'automation_rate': automation_rate,
        'new_handling_hours': new_handling_hours,
        'new_retention_rate': new_retention_rate,
        'labor_savings_idr': labor_savings_idr,
        'labor_savings_usd': labor_savings_usd,
        'retention_revenue_impact': retention_revenue_impact,
        'total_annual_savings_usd': total_annual_savings_usd,
        'first_year_net_usd': first_year_net_usd,
        'subsequent_years_net_usd': subsequent_years_net_usd,
        'first_year_roi': first_year_roi,
        'three_year_roi': three_year_roi,
        'payback_period': payback_period,
        'five_year_net_benefit': five_year_net_benefit,
        'five_year_projection': five_year_projection_data,
        'chart_path': chart_data_uri if chart_data_uri else '',
        'avg_monthly_salary': avg_monthly_salary,
        'overhead_multiplier': overhead_multiplier,
        'usd_conversion_rate': usd_conversion_rate,
        'staff_reduction': staff_reduction,
        'retention_improvement': retention_improvement,
        'handling_time_improvement': handling_time_improvement,
        'conclusion_text': conclusion_text
    }

    # --- Generate PDF --- 
    st.subheader("📄 Proposal PDF")
    pdf_bytes = None
    with st.spinner("Membuat file PDF proposal..."):
        pdf_bytes = generate_pdf(pdf_data)

    if pdf_bytes:
        safe_prospect_name = "".join(c for c in prospect_name if c.isalnum() or c in (' ', '_', '-')).strip()
        safe_location = "".join(c for c in prospect_location if c.isalnum() or c in (' ', '_', '-')).strip()
        pdf_filename = f"{final_proposal_number} {safe_prospect_name} {safe_location}.pdf" # Gunakan nomor proposal otomatis

        st.download_button(
            label="📥 Unduh PDF",
            data=pdf_bytes,
            file_name=pdf_filename,
            mime="application/pdf"
        )
        st.success(f"Proposal PDF siap diunduh: **{pdf_filename}**")

        # --- Unggah ke Google Drive & Log ke Google Sheet ---
        gdrive_pdf_link = None
        if show_upload_log_section:
             st.subheader("☁️ Unggah & Pencatatan")

        if trigger_gdrive_upload and credentials_info:
            gdrive_service = None
            with st.spinner("Menghubungkan ke Google Drive..."):
                gdrive_service = get_gdrive_service(credentials_info)

            if gdrive_service:
                prospect_folder_id = None
                with st.spinner(f"Mencari/membuat folder GDrive untuk '{safe_prospect_name}'..."):
                    prospect_folder_id = find_or_create_folder(gdrive_service, safe_prospect_name, gdrive_parent_folder_id)

                if prospect_folder_id:
                    _, gdrive_pdf_link = upload_to_drive(gdrive_service, pdf_bytes, pdf_filename, prospect_folder_id)
                else:
                    st.error("Gagal mendapatkan/membuat folder prospek GDrive. File tidak diunggah.")
            else:
                 st.error("Gagal terhubung ke Google Drive. File tidak diunggah.")

        # --- Log ke Google Sheet --- 
        if trigger_gsheet_log and credentials_info:
            gsheet_service = None
            # Coba dapatkan service lagi jika belum ada
            if 'gsheets_service' not in locals() or not gsheets_service:
                with st.spinner("Menghubungkan ke Google Sheets..."):
                    gsheet_service = get_gsheets_service(credentials_info)

            if gsheet_service:
                log_data = {
                    'proposal_number': final_proposal_number,
                    'agent_name': agent_name,
                    'agent_email': agent_email,
                    'agent_phone': agent_phone,
                    'prospect_name': prospect_name,
                    'prospect_location': prospect_location,
                    'gdrive_link': gdrive_pdf_link or "Upload Gagal/Tidak Dilakukan"
                }
                with st.spinner("Mencatat proposal ke Google Sheet..."):
                     log_success = log_to_gsheet(gsheet_service, google_sheet_id, log_data)
                     if log_success:
                         st.success("Proposal berhasil dicatat ke Google Sheet.")
            else:
                st.error("Gagal terhubung ke Google Sheets. Proposal tidak dicatat.")
        elif trigger_gsheet_log and not credentials_info:
             st.warning("Kredensial Google tidak valid/ditemukan. Proposal tidak dicatat ke Google Sheet.")

    else:
        st.error("Gagal menghasilkan file PDF. Tidak ada yang dapat diunduh atau diunggah.")

else:
    st.info("Silakan isi data di sidebar kiri dan klik tombol 'Hitung ROI, Buat PDF & Unggah' untuk melihat hasil.")

