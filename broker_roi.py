# -*- coding: utf-8 -*-
"""
Modifikasi skrip ROI Calculator untuk dijalankan di Streamlit,
menambahkan input administrasi, menghasilkan output PDF, mengunggah ke Google Drive,
menggunakan Streamlit Secrets (termasuk Base64), mencatat ke Google Sheets,
dan nomor proposal otomatis. Versi final dengan UX disempurnakan untuk marketing.
Versi ini menambahkan debugging detail untuk Google Sheets dan memperbaiki syntax error.
"""
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from datetime import datetime
import locale # Tetap import untuk jaga-jaga jika ada penggunaan lain, tapi format utama manual
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

# --- Fungsi Bantuan ---
def format_number_id(value, precision=2):
    """Format angka ke format Indonesia (manual, tanpa locale).
    Memastikan titik sebagai pemisah ribuan dan koma sebagai desimal.
    """
    try:
        if isinstance(value, (int, float)):
            if value == float("inf") or value == float("-inf"):
                return "N/A"
            formatted_str = f"{value:,.{precision}f}"
            parts = formatted_str.split(".")
            int_part = parts[0]
            dec_part = parts[1] if len(parts) > 1 else ""
            int_part_swapped = int_part.replace(",", ".")
            return f"{int_part_swapped},{dec_part}"
        return str(value)
    except (TypeError, ValueError):
        return str(value)

def generate_pdf(data):
    """Menghasilkan PDF dari data menggunakan template Jinja2 dan WeasyPrint."""
    try:
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            script_dir = os.getcwd()
        env = Environment(
            loader=FileSystemLoader(script_dir),
            autoescape=select_autoescape(["html"])
        )
        env.filters["format_number"] = format_number_id
        template = env.get_template("template.html")
        html_out = template.render(data)
        if "chart_path" not in data or not data["chart_path"].startswith("data:image"):
             data["chart_path"] = ""
        pdf_bytes = HTML(string=html_out, base_url=script_dir).write_pdf()
        return pdf_bytes
    except Exception as e:
        st.error(f"Error saat membuat PDF: {e}")
        return None

# --- Fungsi Google Drive & Sheets (Auth terpusat) ---
SCOPES = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]

def get_google_credentials(secrets):
    """Mendapatkan kredensial Service Account dari Streamlit Secrets (Base64 atau JSON) atau file upload."""
    credentials_info = None
    source = "Not Found"
    show_api_settings = True
    if "google_service_account_b64" in secrets:
        try:
            b64_str = secrets["google_service_account_b64"]
            if isinstance(b64_str, str) and b64_str:
                decoded_bytes = base64.b64decode(b64_str)
                decoded_str = decoded_bytes.decode("utf-8")
                credentials_info = json.loads(decoded_str)
                source = "Streamlit Secrets (Base64)"
                show_api_settings = False
        except (base64.binascii.Error, json.JSONDecodeError, Exception) as e:
            st.sidebar.warning(f"Gagal memproses secret Base64: {e}")
            pass

    if credentials_info is None and "google_service_account" in secrets:
        try:
            secret_content = secrets["google_service_account"]
            if isinstance(secret_content, str):
                credentials_info = json.loads(secret_content)
            elif isinstance(secret_content, dict) or isinstance(secret_content, toml.TomlDecoder):
                 credentials_info = dict(secret_content)
            if credentials_info:
                source = "Streamlit Secrets (JSON/TOML)"
                show_api_settings = False
        except (json.JSONDecodeError, Exception) as e:
            st.sidebar.warning(f"Gagal memproses secret JSON/TOML: {e}")
            pass

    if credentials_info is None and show_api_settings:
        uploaded_key_file = st.sidebar.file_uploader("Unggah File Kunci JSON Service Account", type=["json"], help="Jika tidak menggunakan Streamlit Secrets.")
        if uploaded_key_file is not None:
            try:
                stringio = io.StringIO(uploaded_key_file.getvalue().decode("utf-8"))
                credentials_info = json.load(stringio)
                source = "File Upload"
                st.sidebar.success("File kunci JSON berhasil dibaca.")
                show_api_settings = True
            except (json.JSONDecodeError, Exception) as e:
                st.sidebar.error(f"Error membaca file kunci: {e}")
                return None, "Error", True

    if credentials_info:
        required_keys = ("type", "project_id", "private_key_id", "private_key", "client_email", "client_id", "auth_uri", "token_uri")
        if not all(k in credentials_info for k in required_keys):
            # Perbaikan: Menggunakan kutip yang benar di f-string
            st.sidebar.error(f"Format kredensial dari {source} tidak lengkap.")
            return None, "Error", True
        return credentials_info, source, show_api_settings
    else:
        if source == "Not Found" and show_api_settings:
             st.sidebar.info("Kredensial Google Service Account tidak ditemukan.")
        return None, "Not Found", show_api_settings

def get_gdrive_service(credentials_info):
    """Membuat service Google Drive. Return None on failure."""
    try:
        credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
        service = build("drive", "v3", credentials=credentials)
        return service
    except Exception as e:
        return None

def get_gsheets_service(credentials_info):
    """Membuat service Google Sheets. Return None on failure."""
    try:
        credentials = service_account.Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
        service = build("sheets", "v4", credentials=credentials)
        return service
    except Exception as e:
        st.sidebar.error(f"DEBUG: Gagal koneksi Google Sheets: {e}")
        return None

def get_next_proposal_number(service, sheet_id):
    """Mendapatkan nomor proposal berikutnya dari Google Sheet. Menampilkan error detail."""
    fallback_num = "PROP-" + datetime.now().strftime("%y%m%d") + "-XXX"
    try:
        range_name = f"{LOG_SHEET_NAME}!B2:B"
        result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_name).execute()
        values = result.get("values", [])
        today_prefix = "PROP-" + datetime.now().strftime("%y%m%d") + "-"
        last_number = 0
        if values:
            for row in reversed(values):
                # Perbaikan: Pastikan row tidak kosong dan punya elemen pertama
                if row and len(row) > 0 and isinstance(row[0], str) and row[0].startswith(today_prefix):
                    try:
                        last_number = int(row[0].split("-")[-1])
                        break
                    except (IndexError, ValueError):
                        continue
        next_number = last_number + 1
        return f"{today_prefix}{next_number:03d}"
    except HttpError as e:
        st.sidebar.error(f"DEBUG: Error API saat baca GSheet (Nomor): {e}")
        return fallback_num
    except Exception as e:
        st.sidebar.error(f"DEBUG: Error umum saat baca GSheet (Nomor): {e}")
        return fallback_num

def log_to_gsheet(service, sheet_id, log_data):
    """Mencatat data proposal ke Google Sheet. Return True on success, False on failure. Menampilkan error detail."""
    try:
        values = [
            [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                log_data.get("proposal_number", ""),
                log_data.get("agent_name", ""),
                log_data.get("agent_email", ""),
                f"'{log_data.get('agent_phone', '')}" if str(log_data.get('agent_phone', '')).startswith('0') else log_data.get('agent_phone', ''), # Add prefix ' if starts with 0
                log_data.get("prospect_name", ""),
                log_data.get("prospect_location", ""),
                log_data.get("gdrive_link", "")
            ]
        ]
        body = {"values": values}
        result = service.spreadsheets().values().append(
            spreadsheetId=sheet_id,
            range=f"{LOG_SHEET_NAME}!A1",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body=body).execute()
        return True
    except HttpError as e:
        st.error(f"DEBUG: Error API saat catat GSheet: {e}")
        return False
    except Exception as e:
        st.error(f"DEBUG: Error umum saat catat GSheet: {e}")
        return False

def find_or_create_folder(service, folder_name, parent_folder_id):
    """Mencari/membuat folder di Google Drive. Return folder ID or None."""
    try:
        safe_folder_name = "".join(c for c in folder_name if c.isalnum() or c in (" ", "_", "-")).strip()
        if not safe_folder_name:
             safe_folder_name = "Prospek Tanpa Nama"
        # Perbaikan: Menggunakan kutip tunggal untuk nilai string dalam query
        query = f"name='{safe_folder_name}' and mimeType='application/vnd.google-apps.folder' and '{parent_folder_id}' in parents and trashed=false"
        response = service.files().list(q=query, spaces="drive", fields="files(id, name)").execute()
        folders = response.get("files", [])
        if folders:
            return folders[0].get("id")
        else:
            file_metadata = {"name": safe_folder_name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_folder_id]}
            folder = service.files().create(body=file_metadata, fields="id").execute()
            return folder.get("id")
    except HttpError as e:
        return None
    except Exception as e:
        return None

def upload_to_drive(service, pdf_bytes, filename, prospect_folder_id):
    """Mengunggah file PDF ke Google Drive. Return link or None."""
    gdrive_link = None
    try:
        file_metadata = {"name": filename, "parents": [prospect_folder_id]}
        media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype="application/pdf", resumable=True)
        request = service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink")
        response = None
        while response is None:
            status, response = request.next_chunk()
        file_id = response.get("id")
        gdrive_link = response.get("webViewLink")
    except HttpError as e:
        pass
    except Exception as e:
        pass
    return gdrive_link

# --- Judul dan Deskripsi Aplikasi ---
st.title("📊 Kalkulator ROI Solusi AI Voice untuk Broker Forex")
st.markdown("Masukkan data operasional dan asumsi untuk menghitung potensi ROI dan menghasilkan proposal PDF.")

# --- Dapatkan Kredensial & ID dari Secrets --- 
secrets = st.secrets
credentials_info, cred_source, show_api_settings_sidebar = get_google_credentials(secrets)
gdrive_parent_folder_id_secret = secrets.get("gdrive_parent_folder_id")
google_sheet_id_secret = secrets.get("google_sheet_id")

# --- Dapatkan Nomor Proposal Berikutnya (jika memungkinkan) ---
next_proposal_num = "PROP-" + datetime.now().strftime("%y%m%d") + "-XXX"
gsheets_service = None
if credentials_info and google_sheet_id_secret:
    gsheets_service = get_gsheets_service(credentials_info)
    if gsheets_service:
        next_proposal_num = get_next_proposal_number(gsheets_service, google_sheet_id_secret)

# --- Input Data --- 
with st.sidebar:
    st.header("⚙️ Input Data")
    st.subheader("Informasi Konsultan")
    agent_name = st.text_input("Nama Konsultan", "")
    agent_email = st.text_input("Email Konsultan", "")
    agent_phone = st.text_input("No. HP/WA Konsultan", "")
    st.subheader("Informasi Proposal & Prospek")
    st.text_input("Nomor Proposal (Otomatis)", value=next_proposal_num, disabled=True)
    prospect_name = st.text_input("Nama Prospek (Broker Forex)", "PT Contoh Broker")
    prospect_location = st.text_input("Lokasi Prospek", "Jakarta")
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
    st.subheader("Investasi Solusi AI Voice")
    implementation_cost = st.number_input("Biaya implementasi (USD)", min_value=0.0, value=10000.0, step=1000.0)
    annual_subscription = st.number_input("Biaya langganan tahunan (USD)", min_value=0.0, value=5000.0, step=500.0)
    st.subheader("Asumsi Dampak Solusi AI Voice")
    automation_rate = st.slider("Otomatisasi pertanyaan (%) ", 0, 100, 75)
    staff_reduction = st.slider("Pengurangan staf CS (%) ", 0, 100, 35)
    retention_improvement = st.slider("Peningkatan loyalitas klien (%) ", 0.0, 20.0, 7.5, 0.5)
    handling_time_improvement = st.slider("Peningkatan waktu penanganan (%) ", 0, 100, 25)
    gdrive_parent_folder_id_input = None
    google_sheet_id_input = None
    if show_api_settings_sidebar:
        st.subheader("☁️ Pengaturan Google API (Manual)")
        if not gdrive_parent_folder_id_secret:
            gdrive_parent_folder_id_input = st.text_input("ID Folder Induk Google Drive", "", help="Masukkan jika tidak diset di Streamlit Secrets.")
        else:
            st.info("ID Folder Induk GDrive ditemukan di Streamlit Secrets.")
        if not google_sheet_id_secret:
            google_sheet_id_input = st.text_input("ID Google Sheet (untuk Log)", "", help="Masukkan jika tidak diset di Streamlit Secrets.")
        else:
            st.info("ID Google Sheet ditemukan di Streamlit Secrets.")
    st.divider()
    calculate_button = st.button("📊 Hitung ROI & Buat Proposal PDF")

# --- Kalkulasi & Output --- 
if calculate_button:
    gdrive_parent_folder_id = gdrive_parent_folder_id_secret or gdrive_parent_folder_id_input
    google_sheet_id = google_sheet_id_secret or google_sheet_id_input
    if not agent_name or not agent_email or not agent_phone:
        st.sidebar.error("Harap isi semua informasi Agent/Marketing.")
        st.stop()
    if not prospect_name:
        st.sidebar.error("Harap isi Nama Prospek.")
        st.stop()

    trigger_gdrive_upload = False
    trigger_gsheet_log = False
    gdrive_service = None

    if credentials_info:
        if gdrive_parent_folder_id:
            gdrive_service = get_gdrive_service(credentials_info)
            if gdrive_service:
                trigger_gdrive_upload = True
        if google_sheet_id:
            if gsheets_service:
                trigger_gsheet_log = True

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
        first_year_roi = (first_year_net_usd / total_first_year_investment * 100) if total_first_year_investment > 0 else float("inf")
        three_year_net_benefit = first_year_net_usd + subsequent_years_net_usd * 2
        three_year_roi = (three_year_net_benefit / total_three_year_investment * 100) if total_three_year_investment > 0 else float("inf")
        monthly_savings_usd = total_annual_savings_usd / 12 if total_annual_savings_usd else 0
        total_investment_usd = implementation_cost + annual_subscription
        payback_period = (total_investment_usd / monthly_savings_usd) if monthly_savings_usd > 0 else float("inf")
        years = range(1, 6)
        costs = [(implementation_cost + annual_subscription) if year == 1 else annual_subscription for year in years]
        benefits = [total_annual_savings_usd for _ in years]
        net_benefits = [benefits[i] - costs[i] for i in range(len(years))]
        cumulative_net = np.cumsum(net_benefits).tolist()
        five_year_net_benefit = cumulative_net[-1] if cumulative_net else 0
        five_year_projection_data = []
        for i in range(len(years)):
            five_year_projection_data.append({
                "year": years[i],
                "cost": costs[i],
                "benefit": benefits[i],
                "net_benefit": net_benefits[i],
                "cumulative_net": cumulative_net[i]
            })

    st.subheader("📊 Grafik Analisis")
    chart_buffer = io.BytesIO()
    chart_data_uri = None
    try:
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        plt.style.use("seaborn-v0_8-whitegrid")
        labels_cost = ["Saat Ini", "Dengan AI Voice"]
        cs_costs_idr = [current_annual_labor_cost_idr, new_annual_labor_cost_idr]
        colors_cost = ["#3A86FF", "#8338EC"]
        bars = ax1.bar(labels_cost, cs_costs_idr, color=colors_cost)
        ax1.set_title("Perbandingan Biaya Tahunan (IDR)", fontsize=12)
        ax1.set_ylabel("Biaya (IDR)", fontsize=10)
        ax1.tick_params(axis="x", labelsize=10)
        ax1.tick_params(axis="y", labelsize=10)
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format_number_id(x, 0)))
        for bar in bars:
            yval = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2.0, yval * 1.01, f"IDR {format_number_id(yval, 0)}", va="bottom", ha="center", fontsize=9)
        ax2.plot(years, cumulative_net, marker="o", linewidth=2, color="#FF006E", label="Manfaat Bersih Kumulatif")
        ax2.set_title("Manfaat Bersih Kumulatif 5 Tahun (USD)", fontsize=12)
        ax2.set_xlabel("Tahun", fontsize=10)
        ax2.set_ylabel("USD", fontsize=10)
        ax2.grid(True, linestyle="--", alpha=0.6)
        ax2.set_xticks(years)
        ax2.tick_params(axis="x", labelsize=10)
        ax2.tick_params(axis="y", labelsize=10)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"$ {format_number_id(x, 0)}"))
        ax2.axhline(0, color="grey", linewidth=0.8, linestyle="--")
        for i, value in enumerate(cumulative_net):
             ax2.text(years[i], value, f"$ {format_number_id(value, 0)}", ha="center", va="bottom", fontsize=9)
        plt.tight_layout(pad=2.0)
        plt.savefig(chart_buffer, format="png", dpi=300, bbox_inches="tight")
        plt.close(fig)
        chart_buffer.seek(0)
        chart_base64 = base64.b64encode(chart_buffer.read()).decode()
        chart_data_uri = f"data:image/png;base64,{chart_base64}"
        st.image(chart_buffer, caption="Grafik Analisis ROI", use_container_width=True)
    except Exception as e:
        st.error(f"Gagal membuat grafik: {e}")

    st.subheader("Ringkasan Hasil Utama")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="ROI Tahun Pertama", value=f"{format_number_id(first_year_roi)}%")
        st.metric(label="ROI Tiga Tahun", value=f"{format_number_id(three_year_roi)}%")
        st.metric(label="Periode Pengembalian", value=f"{format_number_id(payback_period, 1)} bulan" if payback_period != float("inf") else "Tidak Tercapai")
    with col2:
        st.metric(label="Total Penghematan Tahunan (USD)", value=f"$ {format_number_id(total_annual_savings_usd)}")
        st.metric(label="Manfaat Bersih 5 Tahun (USD)", value=f"$ {format_number_id(five_year_net_benefit)}")
        st.metric(label="Pengurangan Staf", value=f"{cs_staff - new_staff_count} orang ({staff_reduction}%)")
    with col3:
        st.metric(label="Otomatisasi Pertanyaan", value=f"{format_number_id(automated_inquiries, 0)} ({automation_rate}%)")
        st.metric(label="Peningkatan Loyalitas Klien", value=f"+{retention_improvement}% (menjadi {new_retention_rate}%)")
        st.metric(label="Penghematan Biaya Tahunan (IDR)", value=f"Rp {format_number_id(labor_savings_idr)}")

    st.subheader("📝 Kesimpulan")
    if first_year_roi != float("inf") and first_year_roi > 50 and payback_period < 18:
        conclusion_text = f"Implementasi Solusi AI Voice untuk {prospect_name} sangat direkomendasikan. Dengan ROI tahun pertama {format_number_id(first_year_roi)}% dan periode pengembalian hanya {format_number_id(payback_period, 1)} bulan, investasi ini menawarkan nilai finansial yang sangat signifikan dan cepat."
    elif first_year_roi != float("inf") and first_year_roi > 0:
        conclusion_text = f"Implementasi Solusi AI Voice untuk {prospect_name} direkomendasikan. ROI tahun pertama sebesar {format_number_id(first_year_roi)}% dan periode pengembalian {format_number_id(payback_period, 1)} bulan menunjukkan potensi pengembalian investasi yang solid dalam jangka menengah."
    elif three_year_roi != float("inf") and three_year_roi > 0:
         conclusion_text = f"Implementasi Solusi AI Voice untuk {prospect_name} patut dipertimbangkan. Meskipun ROI tahun pertama mungkin belum positif ({format_number_id(first_year_roi)}%), ROI tiga tahun sebesar {format_number_id(three_year_roi)}% mengindikasikan potensi keuntungan jangka panjang yang menarik."
    else:
        conclusion_text = f"Berdasarkan data dan asumsi saat ini, ROI untuk implementasi Solusi AI Voice bagi {prospect_name} terlihat kurang menarik ({format_number_id(first_year_roi)}% ROI tahun pertama). Perlu evaluasi lebih lanjut terhadap asumsi atau potensi manfaat lain sebelum melanjutkan."
    st.markdown(conclusion_text)

    st.subheader("📄 Proposal PDF")
    pdf_bytes = None
    with st.spinner("Membuat file PDF proposal..."):
        # Perbaikan: Memasukkan data yang benar ke generate_pdf
        current_time = datetime.now() # Define current_time here
        pdf_data = {
            "proposal_number": next_proposal_num,
            "analysis_date": current_time.strftime("%d %B %Y"),
            "prospect_name": prospect_name,
            "prospect_location": prospect_location,
            "provider_company_name": PROVIDER_COMPANY_NAME,
            "agent_name": agent_name,
            "agent_email": agent_email,
            "agent_phone": agent_phone,
            "cs_staff": cs_staff,
            "current_annual_labor_cost_idr": current_annual_labor_cost_idr,
            "current_annual_labor_cost_usd": current_annual_labor_cost_usd,
            "current_inquiries_per_year": current_inquiries_per_year,
            "current_handling_hours": current_handling_hours,
            "current_retention_rate": current_retention_rate,
            "new_staff_count": new_staff_count,
            "new_annual_labor_cost_idr": new_annual_labor_cost_idr,
            "new_annual_labor_cost_usd": new_annual_labor_cost_usd,
            "automated_inquiries": automated_inquiries,
            "automation_rate": automation_rate,
            "new_handling_hours": new_handling_hours,
            "new_retention_rate": new_retention_rate,
            "labor_savings_idr": labor_savings_idr,
            "labor_savings_usd": labor_savings_usd,
            "retention_revenue_impact": retention_revenue_impact,
            "total_annual_savings_usd": total_annual_savings_usd,
            "first_year_net_usd": first_year_net_usd,
            "subsequent_years_net_usd": subsequent_years_net_usd,
            "first_year_roi": first_year_roi,
            "three_year_roi": three_year_roi,
            "payback_period": payback_period,
            "five_year_net_benefit": five_year_net_benefit,
            "five_year_projection": five_year_projection_data,
            "chart_path": chart_data_uri if chart_data_uri else "",
            "avg_monthly_salary": avg_monthly_salary,
            "overhead_multiplier": overhead_multiplier,
            "usd_conversion_rate": usd_conversion_rate,
            "staff_reduction": staff_reduction,
            "retention_improvement": retention_improvement,
            "handling_time_improvement": handling_time_improvement,
            "conclusion_text": conclusion_text
        }
        pdf_bytes = generate_pdf(pdf_data)

    if pdf_bytes:
        safe_prospect_name = "".join(c for c in prospect_name if c.isalnum() or c in (" ", "_", "-")).strip()
        safe_location = "".join(c for c in prospect_location if c.isalnum() or c in (" ", "_", "-")).strip()
        pdf_filename = f"{next_proposal_num} {safe_prospect_name} {safe_location}.pdf"

        st.download_button(
            label="📥 Unduh PDF",
            data=pdf_bytes,
            file_name=pdf_filename,
            mime="application/pdf"
        )
        st.success(f"Proposal PDF ({pdf_filename}) siap diunduh.")

        # --- Operasi Backend (GDrive & GSheet) ---
        gdrive_pdf_link = None
        if trigger_gdrive_upload and gdrive_service:
            prospect_folder_id = find_or_create_folder(gdrive_service, safe_prospect_name, gdrive_parent_folder_id)
            if prospect_folder_id:
                gdrive_pdf_link = upload_to_drive(gdrive_service, pdf_bytes, pdf_filename, prospect_folder_id)

        # --- Log ke Google Sheet (Menampilkan error jika gagal) --- 
        if trigger_gsheet_log and gsheets_service:
            log_data = {
                "proposal_number": next_proposal_num,
                "agent_name": agent_name,
                "agent_email": agent_email,
                "agent_phone": agent_phone,
                "prospect_name": prospect_name,
                "prospect_location": prospect_location,
                "gdrive_link": gdrive_pdf_link or "Upload Gagal/Tidak Dilakukan"
            }
            log_success = log_to_gsheet(gsheets_service, google_sheet_id, log_data)
            # Pesan error sudah ditangani di dalam log_to_gsheet

    else:
        pass # Error pembuatan PDF sudah ditangani di generate_pdf

else:
    # Perbaikan: Menggunakan kutip yang benar di f-string
    st.info("Silakan isi data di sidebar kiri dan klik tombol \"📊 Hitung ROI & Buat Proposal PDF\" untuk melihat hasil.")

