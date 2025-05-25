#!/usr/bin/env python3
# Kalkulator ROI 5 Tahun untuk Implementasi AI Voice di Rumah Sakit
# Created by: Medical Solutions
# Versi Streamlit 2.4 (2024) - Added Google Drive Upload

import streamlit as st
from datetime import datetime
import locale
import matplotlib.pyplot as plt
import numpy as np
from contextlib import suppress
import re
import os
import base64
from io import BytesIO
import logging
import unicodedata

# --- Google Drive Integration --- 
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaIoBaseUpload
    GOOGLE_API_AVAILABLE = True
except ImportError:
    GOOGLE_API_AVAILABLE = False
    st.warning("Google API libraries not found. Google Drive upload will be disabled. Please install them (`pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib`).")

# --- PDF Generation --- 
try:
    from weasyprint import HTML, CSS
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False
    # Warning is displayed only if Google API is available but WeasyPrint is not
    if GOOGLE_API_AVAILABLE:
        st.warning("WeasyPrint library not found. PDF generation (and thus Drive upload) will be disabled. Please install it (`pip install weasyprint`).")

# --- Configuration --- 
SERVICE_ACCOUNT_FILE = "/home/ubuntu/upload/service_account_key.json" # Path to the credentials file
PARENT_FOLDER_ID = "1bCG7m4T73K3RNoMvWTE4fjRdWCkwAiKR" # User provided parent folder ID
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ====================== HELPER FUNCTIONS ======================

def setup_locale():
    """Mengatur locale untuk format angka"""
    for loc in ["id_ID.UTF-8", "en_US.UTF-8", "C.UTF-8", ""]:
        with suppress(locale.Error):
            locale.setlocale(locale.LC_ALL, loc)
            break

def format_currency(amount, symbol="Rp "):
    """Format angka ke mata uang IDR dengan pemisah ribuan"""
    try:
        return locale.currency(float(amount), symbol=symbol, grouping=True, international=False)
    except (ValueError, TypeError):
         return f"{symbol}0"
    except Exception:
        try:
            return f"{symbol}{float(amount):,.0f}".replace(",", ".")
        except Exception:
            return f"{symbol}-"

def format_percent(value):
    """Formats a ratio as a percentage string."""
    try:
        return f"{float(value)*100:.1f}%"
    except (ValueError, TypeError):
        return "0.0%"

def format_months(months):
    """Formats payback period in months."""
    if months == float("inf"):
        return "Tidak Tercapai"
    if months < 0:
        return "Instan"
    try:
        return f"{float(months):.1f} Bulan"
    except (ValueError, TypeError):
        return "N/A"

def calculate_roi(investment, annual_net_gain, years):
    """Hitung ROI dalam persen untuk X tahun berdasarkan annual NET gain"""
    try:
        investment = float(investment)
        annual_net_gain = float(annual_net_gain)
        if investment == 0:
            return float("inf") if annual_net_gain > 0 else (0 if annual_net_gain == 0 else float("-inf"))
        total_net_gain = annual_net_gain * years
        roi = ((total_net_gain - investment) / investment) * 100
        return roi
    except (ValueError, TypeError):
        return 0

def format_roi(roi_value):
    """Formats ROI value for display."""
    if roi_value == float("inf"):
        return "Infinite"
    if roi_value == float("-inf"):
        return "-Infinite"
    try:
        return f"{float(roi_value):.1f}%"
    except (ValueError, TypeError):
        return "0.0%"

def sanitize_filename(name):
    """Sanitizes a string to be safe for filenames."""
    # Normalize unicode characters
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    # Replace spaces with underscores
    name = name.replace(" ", "_")
    # Remove invalid characters
    name = re.sub(r"[^a-zA-Z0-9_.-]", "", name)
    # Truncate if too long (optional)
    return name[:100] # Limit length

def generate_charts(data):
    """Generate grafik untuk visualisasi data dan return figure objects."""
    figs = {}
    try:
        # Grafik Proyeksi Arus Kas Kumulatif
        fig1, ax1 = plt.subplots(figsize=(8, 4))
        months = 60
        monthly_net_savings_val = float(data.get("total_monthly_net_savings", 0))
        total_investment_val = float(data.get("total_investment", 0))
        cumulative = [monthly_net_savings_val * m - total_investment_val for m in range(1, months+1)]
        ax1.plot(range(1, months+1), cumulative, color="#2E86C1", linewidth=1.5, marker=".", markersize=3)
        breakeven_month = next((m for m, c in enumerate(cumulative, 1) if c >= 0), None)
        ax1.axhline(0, color="grey", linestyle="--", linewidth=0.6)
        if breakeven_month:
            ax1.plot(breakeven_month, cumulative[breakeven_month-1], "ro", markersize=4, label=f"Breakeven: ~Bln {breakeven_month}")
            ax1.axvline(breakeven_month, color="red", linestyle="--", linewidth=0.6)
            ax1.legend(fontsize=8)
        ax1.set_title("Proyeksi Arus Kas Kumulatif Bersih (5 Tahun)", fontsize=10)
        ax1.set_xlabel("Bulan Sejak Implementasi", fontsize=9)
        ax1.set_ylabel("Arus Kas Kumulatif (IDR)", fontsize=9)
        ax1.grid(True, linestyle="--", alpha=0.5)
        ax1.ticklabel_format(style="plain", axis="y")
        ax1.tick_params(axis="both", which="major", labelsize=8)
        plt.tight_layout()
        figs["cashflow"] = fig1

        # Grafik Perbandingan Penghematan Tahunan
        fig2, ax2 = plt.subplots(figsize=(8, 4))
        categories = ["Staff", "No-Show", "Total Kotor"]
        staff_savings_annual = float(data.get("staff_savings_monthly", 0)) * 12
        noshow_savings_annual = float(data.get("noshow_savings_monthly", 0)) * 12
        gross_annual_savings_val = staff_savings_annual + noshow_savings_annual
        savings = [staff_savings_annual, noshow_savings_annual, gross_annual_savings_val]
        bars = ax2.bar(categories, savings, color=["#27AE60", "#F1C40F", "#E74C3C"], width=0.6)
        ax2.set_title("Komponen Penghematan Kotor Tahunan", fontsize=10)
        ax2.set_ylabel("Jumlah (IDR)", fontsize=9)
        ax2.ticklabel_format(style="plain", axis="y")
        ax2.tick_params(axis="both", which="major", labelsize=8)
        ax2.grid(True, axis="y", linestyle="--", alpha=0.5)
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2.0, yval, format_currency(yval).replace("Rp ", ""), va="bottom" if yval >= 0 else "top", ha="center", fontsize=7)
        plt.tight_layout()
        figs["savings"] = fig2

    except Exception as e:
        logger.error(f"Gagal membuat grafik: {e}", exc_info=True)
        st.error(f"Gagal membuat grafik: {e}")
    finally:
        plt.close(fig1) if "cashflow" in figs else None
        plt.close(fig2) if "savings" in figs else None
    return figs

def fig_to_base64(fig):
    """Converts a Matplotlib figure to a base64 encoded PNG."""
    if fig is None:
        return None
    try:
        img_buf = BytesIO()
        fig.savefig(img_buf, format="png", dpi=150, bbox_inches="tight")
        img_buf.seek(0)
        return base64.b64encode(img_buf.getvalue()).decode("utf-8")
    except Exception as e:
        logger.error(f"Error converting figure to base64: {e}", exc_info=True)
        return None

def generate_pdf_report(data, charts_base64):
    """Generates a PDF report from data using WeasyPrint. Returns PDF bytes."""
    if not WEASYPRINT_AVAILABLE:
        st.error("PDF generation failed: WeasyPrint library is not available.")
        return None

    # --- HTML Content --- (Using f-string for simplicity)
    # Ensure chart data exists before trying to embed
    cashflow_chart_html = f"<img src=\"data:image/png;base64,{charts_base64.get('cashflow', '')}\" alt=\"Grafik Arus Kas\">" if charts_base64.get("cashflow") else "<p><i>Grafik arus kas tidak tersedia.</i></p>"
    savings_chart_html = f"<img src=\"data:image/png;base64,{charts_base64.get('savings', '')}\" alt=\"Grafik Penghematan\">" if charts_base64.get("savings") else "<p><i>Grafik penghematan tidak tersedia.</i></p>"

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Laporan Analisis ROI - {data.get("hospital_name", "N/A")}</title>
        <style>
            @page {{ size: A4; margin: 1.5cm; }}
            body {{ font-family: "Helvetica Neue", Helvetica, Arial, sans-serif; font-size: 10pt; line-height: 1.4; color: #333; }}
            h1, h2, h3 {{ font-family: "Trebuchet MS", Helvetica, sans-serif; color: #1F618D; margin-bottom: 0.5em; line-height: 1.2; }}
            h1 {{ font-size: 18pt; border-bottom: 2px solid #2E86C1; padding-bottom: 5px; margin-bottom: 1em; }}
            h2 {{ font-size: 14pt; border-bottom: 1px solid #AED6F1; padding-bottom: 3px; margin-top: 1.5em; }}
            h3 {{ font-size: 11pt; color: #2E86C1; margin-top: 1.2em; margin-bottom: 0.3em; }}
            p {{ margin-bottom: 0.8em; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 1em; page-break-inside: avoid; }}
            th, td {{ border: 1px solid #ddd; padding: 6px 8px; text-align: left; vertical-align: top; }}
            th {{ background-color: #f2f2f2; font-weight: bold; color: #444; }}
            .summary-box {{ background-color: #f8f9fa; border: 1px solid #e0e0e0; border-left: 4px solid #2E86C1; padding: 15px; margin: 1.5em 0; border-radius: 4px; page-break-inside: avoid; }}
            .summary-box h3 {{ margin-top: 0; color: #1F618D; font-size: 12pt; }}
            .summary-box p {{ margin-bottom: 0.5em; font-size: 11pt; }}
            .summary-box strong {{ font-weight: 600; color: #111; }}
            .chart-container {{ text-align: center; margin: 1.5em 0; page-break-inside: avoid; }}
            .chart-container img {{ max-width: 95%; height: auto; border: 1px solid #ccc; padding: 5px; margin-top: 5px; }}
            .footer {{ text-align: center; font-size: 8pt; color: #777; margin-top: 2em; border-top: 1px solid #ccc; padding-top: 5px; position: fixed; bottom: 0.5cm; left: 1.5cm; right: 1.5cm; }}
            .label {{ font-weight: bold; color: #555; width: 45%; }}
            .value {{ }}
            .page-break {{ page-break-before: always; }}
        </style>
    </head>
    <body>
        <h1>Laporan Analisis ROI Implementasi AI Voice</h1>
        <p>Dokumen ini menyajikan analisis estimasi Return on Investment (ROI) selama 5 tahun untuk implementasi solusi AI Voice di <strong>{data.get("hospital_name", "N/A")}</strong>, berlokasi di {data.get("hospital_location", "N/A")}.</p>
        <p>Analisis ini dihitung pada: {data.get("calculation_timestamp", datetime.now()).strftime("%d %B %Y, %H:%M:%S")}</p>
        <p>Dipersiapkan oleh: <strong>{data.get("consultant_name", "N/A")}</strong> ({data.get("consultant_email", "N/A")}, {data.get("consultant_phone", "N/A")})</p>

        <div class="summary-box">
            <h3>Ringkasan Eksekutif</h3>
            <p>Implementasi AI Voice diproyeksikan menghasilkan <strong>Penghematan Bersih Tahunan</strong> sebesar <strong>{format_currency(data.get("annual_net_savings", 0))}</strong> setelah memperhitungkan biaya operasional. Investasi awal yang dibutuhkan adalah sekitar <strong>{format_currency(data.get("total_investment", 0))}</strong>. Dengan metrik ini, estimasi <strong>Payback Period</strong> adalah <strong>{format_months(data.get("payback_period_months", float("inf")))}</strong>. Proyeksi ROI untuk 1, 3, dan 5 tahun adalah <strong>{format_roi(data.get("roi_1_year", 0))}</strong>, <strong>{format_roi(data.get("roi_3_year", 0))}</strong>, dan <strong>{format_roi(data.get("roi_5_year", 0))}</strong> secara berturut-turut.</p>
        </div>

        <h2>Parameter Input Utama</h2>
        <table>
            <tr><th colspan="2">Operasional & Biaya Saat Ini</th></tr>
            <tr><td class="label">Rata-rata Janji Temu/Bulan</td><td class="value">{data.get("monthly_appointments", 0):,.0f}</td></tr>
            <tr><td class="label">Tingkat No-Show Saat Ini</td><td class="value">{format_percent(data.get("noshow_rate_before", 0))}</td></tr>
            <tr><td class="label">Staff Admin Terdampak</td><td class="value">{data.get("admin_staff", 0):,.0f} orang</td></tr>
            <tr><td class="label">Gaji Rata-rata Staff Admin</td><td class="value">{format_currency(data.get("avg_salary", 0))}/Bulan</td></tr>
            <tr><td class="label">Pendapatan Rata-rata/Janji Temu</td><td class="value">{format_currency(data.get("revenue_per_appointment", 0))}</td></tr>
            <tr><th colspan="2">Target Efisiensi & Biaya Implementasi</th></tr>
            <tr><td class="label">Target Pengurangan Beban Kerja Staff</td><td class="value">{format_percent(data.get("staff_reduction_target", 0))}</td></tr>
            <tr><td class="label">Target Pengurangan Angka No-Show</td><td class="value">{format_percent(data.get("noshow_reduction_target", 0))}</td></tr>
            <tr><td class="label">Asumsi Kurs USD-IDR</td><td class="value">{format_currency(data.get("exchange_rate", 1), symbol="Rp ")}</td></tr>
            <tr><td class="label">Biaya Langganan Tahunan (USD)</td><td class="value">{format_currency(data.get("annual_subscription_usd", 0), symbol="$ ")}</td></tr>
            <tr><td class="label">Biaya Pemeliharaan Lainnya (IDR/Bulan)</td><td class="value">{format_currency(data.get("maintenance_cost_monthly", 0))}</td></tr>
        </table>

        <div class="page-break"></div>

        <h2>Hasil Perhitungan Estimasi</h2>
        <table>
            <tr><th colspan="2">Estimasi Investasi Awal (IDR)</th></tr>
            <tr><td class="label">Biaya Setup Awal</td><td class="value">{format_currency(data.get("setup_cost_idr", 0))}</td></tr>
            <tr><td class="label">Biaya Integrasi Sistem</td><td class="value">{format_currency(data.get("integration_cost_idr", 0))}</td></tr>
            <tr><td class="label">Biaya Pelatihan Staff</td><td class="value">{format_currency(data.get("training_cost_idr", 0))}</td></tr>
            <tr><td class="label"><strong>Total Investasi Awal</strong></td><td class="value"><strong>{format_currency(data.get("total_investment", 0))}</strong></td></tr>
            <tr><th colspan="2">Estimasi Penghematan Kotor Tahunan (IDR)</th></tr>
            <tr><td class="label">Penghematan dari Efisiensi Staff</td><td class="value">{format_currency(data.get("staff_savings_monthly", 0) * 12)}</td></tr>
            <tr><td class="label">Penghematan dari Pengurangan No-Show</td><td class="value">{format_currency(data.get("noshow_savings_monthly", 0) * 12)}</td></tr>
            <tr><td class="label"><strong>Total Penghematan Kotor Tahunan</strong></td><td class="value"><strong>{format_currency(data.get("annual_gross_savings", 0))}</strong></td></tr>
            <tr><th colspan="2">Estimasi Biaya Operasional Tahunan (IDR)</th></tr>
            <tr><td class="label">Biaya Langganan Tahunan</td><td class="value">{format_currency(data.get("annual_subscription_idr", 0))}</td></tr>
            <tr><td class="label">Biaya Pemeliharaan & Support Lainnya</td><td class="value">{format_currency(data.get("maintenance_cost_monthly", 0) * 12)}</td></tr>
            <tr><td class="label"><strong>Total Biaya Operasional Tahunan</strong></td><td class="value"><strong>{format_currency(data.get("total_annual_operational_cost", 0))}</strong></td></tr>
             <tr><th colspan="2">Estimasi Penghematan Bersih Tahunan (IDR)</th></tr>
             <tr><td class="label">Total Penghematan Kotor Tahunan</td><td class="value">{format_currency(data.get("annual_gross_savings", 0))}</td></tr>
             <tr><td class="label">Dikurangi Total Biaya Operasional Tahunan</td><td class="value">{format_currency(data.get("total_annual_operational_cost", 0))}</td></tr>
             <tr><td class="label"><strong>Total Penghematan Bersih Tahunan</strong></td><td class="value"><strong>{format_currency(data.get("annual_net_savings", 0))}</strong></td></tr>
        </table>

        <h2>Metrik Utama ROI</h2>
        <table>
            <tr><td class="label">Payback Period</td><td class="value">{format_months(data.get("payback_period_months", float("inf")))}</td></tr>
            <tr><td class="label">ROI 1 Tahun</td><td class="value">{format_roi(data.get("roi_1_year", 0))}</td></tr>
            <tr><td class="label">ROI 3 Tahun</td><td class="value">{format_roi(data.get("roi_3_year", 0))}</td></tr>
            <tr><td class="label">ROI 5 Tahun</td><td class="value">{format_roi(data.get("roi_5_year", 0))}</td></tr>
        </table>

        <h2>Visualisasi Proyeksi</h2>
        <div class="chart-container">
            <h3>Proyeksi Arus Kas Kumulatif Bersih (5 Tahun)</h3>
            {cashflow_chart_html}
        </div>
        <div class="chart-container">
            <h3>Komponen Penghematan Kotor Tahunan</h3>
            {savings_chart_html}
        </div>

        <div class="footer">
            Analisis ini adalah estimasi berdasarkan input yang diberikan. Hasil aktual dapat bervariasi. | © {datetime.now().year} Medical AI Solutions Inc.
        </div>
    </body>
    </html>
    """

    try:
        # Use WeasyPrint to render HTML to PDF bytes
        pdf_bytes = HTML(string=html_content).write_pdf(stylesheets=[CSS(string='@page { size: A4; margin: 1.5cm; }')])
        return pdf_bytes
    except Exception as e:
        logger.error(f"Error generating PDF with WeasyPrint: {e}", exc_info=True)
        st.error(f"Gagal membuat file PDF: {e}")
        return None

# ====================== GOOGLE DRIVE FUNCTIONS ======================

def get_drive_service():
    """Authenticates and returns the Google Drive service object."""
    if not GOOGLE_API_AVAILABLE:
        logger.warning("Google API libraries not available.")
        return None
    try:
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build("drive", "v3", credentials=creds)
        logger.info("Google Drive service authenticated successfully.")
        return service
    except FileNotFoundError:
        logger.error(f"Service account file not found at {SERVICE_ACCOUNT_FILE}")
        st.error(f"File Kredensial Google ({SERVICE_ACCOUNT_FILE}) tidak ditemukan.")
        return None
    except Exception as e:
        logger.error(f"Error authenticating Google Drive service: {e}", exc_info=True)
        st.error(f"Gagal mengautentikasi ke Google Drive: {e}")
        return None

def find_or_create_folder(service, folder_name, parent_folder_id):
    """Finds a folder by name within a parent folder, creates it if not found."""
    if not service:
        return None
    try:
        # Sanitize folder name for query
        sanitized_name = folder_name.replace("'", "\'")
        query = f"name='{sanitized_name}' and mimeType='application/vnd.google-apps.folder' and '{parent_folder_id}' in parents and trashed=false"
        response = service.files().list(q=query, spaces="drive", fields="files(id, name)").execute()
        folders = response.get("files", [])

        if folders:
            folder_id = folders[0].get("id")
            logger.info(f"Found existing folder '{folder_name}' with ID: {folder_id}")
            return folder_id
        else:
            logger.info(f"Folder '{folder_name}' not found. Creating...")
            file_metadata = {
                "name": folder_name, # Use original name for creation
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [parent_folder_id]
            }
            folder = service.files().create(body=file_metadata, fields="id").execute()
            folder_id = folder.get("id")
            logger.info(f"Created folder '{folder_name}' with ID: {folder_id}")
            return folder_id
    except HttpError as error:
        logger.error(f"An HTTP error occurred while finding/creating folder: {error}", exc_info=True)
        st.error(f"Error Google Drive (Folder): {error}")
        return None
    except Exception as e:
        logger.error(f"An unexpected error occurred while finding/creating folder: {e}", exc_info=True)
        st.error(f"Error tak terduga saat proses folder Google Drive: {e}")
        return None

def upload_pdf_to_drive(service, pdf_bytes, filename, folder_id):
    """Uploads PDF bytes to the specified Google Drive folder."""
    if not service or not pdf_bytes or not folder_id:
        logger.error("Upload prerequisites not met (service, pdf_bytes, or folder_id missing).")
        return None
    try:
        file_metadata = {
            "name": filename,
            "parents": [folder_id]
        }
        media = MediaIoBaseUpload(BytesIO(pdf_bytes), mimetype="application/pdf", resumable=True)
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink" # Get ID and link of the uploaded file
        ).execute()
        file_id = file.get("id")
        file_link = file.get("webViewLink")
        logger.info(f"File '{filename}' uploaded successfully with ID: {file_id}. Link: {file_link}")
        return file_link # Return the web link
    except HttpError as error:
        logger.error(f"An HTTP error occurred during upload: {error}", exc_info=True)
        st.error(f"Error Google Drive (Upload): {error}")
        return None
    except Exception as e:
        logger.error(f"An unexpected error occurred during upload: {e}", exc_info=True)
        st.error(f"Error tak terduga saat upload ke Google Drive: {e}")
        return None

# ====================== TAMPILAN STREAMLIT ======================

def main():
    st.set_page_config(page_title="Kalkulator ROI AI Voice", page_icon="🏥", layout="wide")

    # CSS (same as before)
    st.markdown("""
    <style>
    /* [CSS styles omitted for brevity] */
    .stButton>button { border-radius: 5px; padding: 10px 15px; font-weight: bold; border: none; transition: background-color 0.3s ease; }
    div[data-testid="stHorizontalBlock"] button[kind="primary"] { background-color: #2E86C1; color: white; }
    div[data-testid="stHorizontalBlock"] button[kind="primary"]:hover { background-color: #1F618D; }
    .stSidebar .stButton>button { background-color: #27AE60; color: white; width: 100%; }
    .stSidebar .stButton>button:hover { background-color: #229954; }
    .stMetric { background-color: #f8f9fa; border-left: 5px solid #2E86C1; padding: 15px 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 10px; }
    .stMetric label { font-weight: 500; color: #5a5a5a; font-size: 0.95em; }
    .stMetric .stMetricValue { font-size: 1.7em; font-weight: 600; color: #1F618D; }
    .stExpander { border: 1px solid #e0e0e0; border-radius: 5px; margin-top: 15px; }
    .stExpander header { font-weight: 600; font-size: 1.1em; background-color: #f1f3f5; padding: 10px 15px !important; border-bottom: 1px solid #e0e0e0; }
    .stExpander header:hover { background-color: #e9ecef; }
    .stSidebar .stTextInput label, .stSidebar .stNumberInput label, .stSidebar .stSlider label { font-size: 0.9rem !important; font-weight: 500; margin-bottom: 2px; }
    .stSidebar .stTextInput, .stSidebar .stNumberInput, .stSidebar .stSlider { margin-bottom: 10px; }
    .stSidebar h3 { color: #1F618D; font-size: 1.1em; font-weight: 600; border-bottom: 2px solid #2E86C1; padding-bottom: 5px; margin-top: 15px; margin-bottom: 10px; }
    .stSidebar [data-testid="stNotification"] { border-radius: 5px; margin-top: 10px; }
    .sidebar-separator { margin-top: 15px; margin-bottom: 15px; border-bottom: 1px solid #e0e0e0; }
    </style>
    """, unsafe_allow_html=True)

    st.title("🏥 Kalkulator ROI Implementasi AI Voice")
    st.markdown("##### Analisis Estimasi *Return on Investment* (ROI) selama 5 tahun untuk solusi AI Voice di fasilitas kesehatan.")
    st.markdown("---")

    # --- Sidebar Inputs --- 
    with st.sidebar:
        st.header("⚙️ Parameter Input")
        # Consultant Info
        st.subheader("Informasi Konsultan")
        consultant_name = st.text_input("Nama Konsultan*", key="consultant_name", placeholder="Contoh: Budi Santoso")
        consultant_email = st.text_input("Email Konsultan*", key="consultant_email", placeholder="contoh@email.com")
        consultant_phone = st.text_input("Nomor HP/WA Konsultan*", key="consultant_phone", placeholder="+62 812 3456 7890")
        st.caption("*Wajib diisi")
        st.markdown("<div class=\"sidebar-separator\"></div>", unsafe_allow_html=True)
        # Hospital Info
        st.subheader("Informasi Rumah Sakit")
        hospital_name = st.text_input("Nama Rumah Sakit", "RS Sehat Sentosa", key="hospital_name")
        hospital_location = st.text_input("Lokasi (Kota/Area)", "Jakarta Selatan", key="hospital_location")
        st.markdown("<div class=\"sidebar-separator\"></div>", unsafe_allow_html=True)
        # Operational Params
        st.subheader("Parameter Operasional Saat Ini")
        col1_op, col2_op = st.columns(2)
        with col1_op: total_staff = st.number_input("Total Staff RS", 1, value=250, step=10, key="total_staff")
        with col2_op: admin_staff = st.number_input("Staff Admin Terkait", 1, value=25, step=1, key="admin_staff")
        monthly_appointments = st.number_input("Rata-rata Janji Temu/Bulan", 1, value=6000, step=100, key="monthly_appointments")
        noshow_rate = st.slider("Tingkat No-Show Saat Ini (%)", 0.0, 50.0, 15.0, step=0.5, key="noshow_rate", format="%.1f%%") / 100
        st.markdown("<div class=\"sidebar-separator\"></div>", unsafe_allow_html=True)
        # Cost Params
        st.subheader("Parameter Biaya Saat Ini")
        avg_salary = st.number_input("Gaji Rata-rata Staff Admin (IDR/Bulan)", 0, value=7500000, step=100000, key="avg_salary", format="%d")
        revenue_per_appointment = st.number_input("Rata-rata Pendapatan/Janji Temu (IDR)", 0, value=300000, step=10000, key="revenue_per_appointment", format="%d")
        st.markdown("<div class=\"sidebar-separator\"></div>", unsafe_allow_html=True)
        # Efficiency Targets
        st.subheader("Target Efisiensi dengan AI Voice")
        staff_reduction = st.slider("Target Pengurangan Beban Kerja Staff Admin (%)", 0.0, 100.0, 30.0, step=1.0, key="staff_reduction", format="%.1f%%") / 100
        noshow_reduction = st.slider("Target Pengurangan Angka No-Show (%)", 0.0, 100.0, 40.0, step=1.0, key="noshow_reduction", format="%.1f%%") / 100
        st.markdown("<div class=\"sidebar-separator\"></div>", unsafe_allow_html=True)
        # Implementation Costs
        st.subheader("Estimasi Biaya Implementasi & Operasional AI Voice")
        exchange_rate = st.number_input("Asumsi Kurs USD-IDR", 1000, value=16000, step=100, key="exchange_rate", format="%d")
        st.markdown("**Biaya Investasi Awal (Satu Kali):**")
        setup_cost_usd = st.number_input("Biaya Setup Awal (USD)", 0, value=15000, step=500, key="setup_cost_usd", format="%d")
        integration_cost_usd = st.number_input("Biaya Integrasi Sistem (USD)", 0, value=10000, step=500, key="integration_cost_usd", format="%d")
        training_cost_usd = st.number_input("Biaya Pelatihan Staff (USD)", 0, value=5000, step=500, key="training_cost_usd", format="%d")
        st.markdown("**Biaya Operasional Berjalan:**")
        annual_subscription_usd = st.number_input("Biaya Langganan per Tahun (USD)", 0, value=12000, step=100, key="annual_subscription_usd", format="%d")
        maintenance_cost = st.number_input("Biaya Pemeliharaan & Support Lainnya (IDR/Bulan)", 0, value=6000000, step=100000, key="maintenance_cost", format="%d")
        st.markdown("<div class=\"sidebar-separator\"></div>", unsafe_allow_html=True)

        # Calculate Button
        hitung_roi = st.button("🚀 HITUNG ROI SEKARANG", key="hitung_button")

    # --- Main Area --- 
    results_placeholder = st.empty()
    pdf_upload_placeholder = st.empty() # Placeholder for PDF/Upload button and status

    # Initialize session state
    if "results_calculated" not in st.session_state: st.session_state.results_calculated = False
    if "report_data" not in st.session_state: st.session_state.report_data = {}
    if "pdf_bytes" not in st.session_state: st.session_state.pdf_bytes = None
    if "pdf_filename" not in st.session_state: st.session_state.pdf_filename = None
    if "drive_upload_status" not in st.session_state: st.session_state.drive_upload_status = None # None, "success", "error", "pending"
    if "drive_file_link" not in st.session_state: st.session_state.drive_file_link = None

    # --- Calculation Logic --- 
    if hitung_roi:
        # Reset previous upload status on new calculation
        st.session_state.drive_upload_status = None
        st.session_state.drive_file_link = None
        st.session_state.pdf_bytes = None
        st.session_state.pdf_filename = None

        # Validation
        error_messages = []
        c_name = st.session_state.get("consultant_name", "").strip()
        c_email = st.session_state.get("consultant_email", "").strip()
        c_phone = st.session_state.get("consultant_phone", "").strip()
        if not c_name: error_messages.append("Nama Konsultan wajib diisi.")
        if not c_email: error_messages.append("Email Konsultan wajib diisi.")
        elif not re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", c_email): error_messages.append("Format Email Konsultan tidak valid.")
        if not c_phone: error_messages.append("Nomor HP/WA Konsultan wajib diisi.")
        elif not re.match(r"^\+?([\d\s\-()]{7,})$", c_phone): error_messages.append("Format Nomor HP/WA Konsultan tidak valid.")

        if error_messages:
            st.session_state.results_calculated = False
            st.session_state.report_data = {}
            results_placeholder.empty()
            pdf_upload_placeholder.empty()
            for msg in error_messages: st.sidebar.error(msg)
            st.stop()
        else:
            st.session_state.results_calculated = True

        # Proceed with calculations
        try:
            # Fetch values from session state
            h_name = st.session_state.get("hospital_name", "N/A")
            h_loc = st.session_state.get("hospital_location", "N/A")
            ex_rate = st.session_state.get("exchange_rate", 1)
            s_cost_usd = st.session_state.get("setup_cost_usd", 0)
            i_cost_usd = st.session_state.get("integration_cost_usd", 0)
            t_cost_usd = st.session_state.get("training_cost_usd", 0)
            ann_sub_usd = st.session_state.get("annual_subscription_usd", 0)
            m_cost = st.session_state.get("maintenance_cost", 0)
            a_staff = st.session_state.get("admin_staff", 0)
            s_reduc = st.session_state.get("staff_reduction", 0.0)
            a_salary = st.session_state.get("avg_salary", 0)
            m_appts = st.session_state.get("monthly_appointments", 0)
            n_rate = st.session_state.get("noshow_rate", 0.0)
            n_reduc = st.session_state.get("noshow_reduction", 0.0)
            r_appt = st.session_state.get("revenue_per_appointment", 0)

            # Calculations (Investment, Operational Costs, Savings, ROI, Payback)
            setup_cost = s_cost_usd * ex_rate
            integration_cost = i_cost_usd * ex_rate
            training_cost = t_cost_usd * ex_rate
            total_investment = setup_cost + integration_cost + training_cost
            annual_subscription_idr = ann_sub_usd * ex_rate
            monthly_subscription_idr = annual_subscription_idr / 12
            total_monthly_operational_cost = m_cost + monthly_subscription_idr
            total_annual_operational_cost = total_monthly_operational_cost * 12
            staff_savings_monthly = (a_staff * s_reduc) * a_salary
            noshow_reduction_count_monthly = m_appts * n_rate * n_reduc
            noshow_savings_monthly = noshow_reduction_count_monthly * r_appt
            total_monthly_gross_savings = staff_savings_monthly + noshow_savings_monthly
            annual_gross_savings = total_monthly_gross_savings * 12
            total_monthly_net_savings = total_monthly_gross_savings - total_monthly_operational_cost
            annual_net_savings = total_monthly_net_savings * 12
            payback_period_months = total_investment / total_monthly_net_savings if total_monthly_net_savings > 0 else float("inf")
            roi_1_year = calculate_roi(total_investment, annual_net_savings, 1)
            roi_3_year = calculate_roi(total_investment, annual_net_savings, 3)
            roi_5_year = calculate_roi(total_investment, annual_net_savings, 5)

            # Store results in session state
            st.session_state.report_data = {
                "calculation_timestamp": datetime.now(),
                "consultant_name": c_name, "consultant_email": c_email, "consultant_phone": c_phone,
                "hospital_name": h_name, "hospital_location": h_loc,
                # Inputs
                "monthly_appointments": m_appts, "noshow_rate_before": n_rate, "admin_staff": a_staff,
                "avg_salary": a_salary, "revenue_per_appointment": r_appt,
                "staff_reduction_target": s_reduc, "noshow_reduction_target": n_reduc,
                "exchange_rate": ex_rate, "annual_subscription_usd": ann_sub_usd, "maintenance_cost_monthly": m_cost,
                "setup_cost_usd": s_cost_usd, "integration_cost_usd": i_cost_usd, "training_cost_usd": t_cost_usd,
                # Calculated Costs & Savings
                "setup_cost_idr": setup_cost, "integration_cost_idr": integration_cost, "training_cost_idr": training_cost,
                "total_investment": total_investment,
                "annual_subscription_idr": annual_subscription_idr, "monthly_subscription_idr": monthly_subscription_idr,
                "total_monthly_operational_cost": total_monthly_operational_cost, "total_annual_operational_cost": total_annual_operational_cost,
                "staff_savings_monthly": staff_savings_monthly, "noshow_savings_monthly": noshow_savings_monthly,
                "total_monthly_gross_savings": total_monthly_gross_savings, "annual_gross_savings": annual_gross_savings,
                "total_monthly_net_savings": total_monthly_net_savings, "annual_net_savings": annual_net_savings,
                # Metrics
                "payback_period_months": payback_period_months,
                "roi_1_year": roi_1_year, "roi_3_year": roi_3_year, "roi_5_year": roi_5_year,
            }

            # --- Generate PDF --- 
            if WEASYPRINT_AVAILABLE:
                charts = generate_charts(st.session_state.report_data)
                charts_base64 = {k: fig_to_base64(v) for k, v in charts.items()}
                pdf_bytes = generate_pdf_report(st.session_state.report_data, charts_base64)
                if pdf_bytes:
                    st.session_state.pdf_bytes = pdf_bytes
                    # Generate filename: yymmdd namarumahsakit lokasi namakonsultan.pdf
                    ts = st.session_state.report_data["calculation_timestamp"].strftime("%y%m%d")
                    h_name_sanitized = sanitize_filename(st.session_state.report_data.get("hospital_name", "UnknownHospital"))
                    h_loc_sanitized = sanitize_filename(st.session_state.report_data.get("hospital_location", "UnknownLocation"))
                    c_name_sanitized = sanitize_filename(st.session_state.report_data.get("consultant_name", "UnknownConsultant"))
                    st.session_state.pdf_filename = f"{ts}_{h_name_sanitized}_{h_loc_sanitized}_{c_name_sanitized}.pdf"
                else:
                    st.session_state.results_calculated = False # Mark as failed if PDF fails
                    st.error("Gagal membuat laporan PDF.")
            else:
                st.session_state.results_calculated = False # Mark as failed if PDF lib not available
                st.error("Fungsi PDF tidak tersedia.")

        except Exception as e:
            logger.error(f"Terjadi kesalahan saat perhitungan atau pembuatan PDF: {e}", exc_info=True)
            st.error(f"Terjadi kesalahan saat perhitungan: {e}")
            st.session_state.results_calculated = False
            st.session_state.report_data = {}
            results_placeholder.empty()
            pdf_upload_placeholder.empty()
            st.stop()

    # --- Display Results --- 
    if st.session_state.get("results_calculated", False):
        data = st.session_state.get("report_data", {})
        with results_placeholder.container():
            st.header("📊 Ringkasan Hasil Analisis ROI")
            st.markdown(f"Analisis untuk **{data.get('hospital_name', 'N/A')}** ({data.get('hospital_location', 'N/A')}) | Dihitung oleh: **{data.get('consultant_name', 'N/A')}** | {data.get('calculation_timestamp', datetime.now()).strftime('%d %B %Y, %H:%M:%S')}")
            st.markdown("---")
            # Metrics
            col1_res, col2_res, col3_res = st.columns(3)
            with col1_res: st.metric("💰 Investasi Awal (IDR)", format_currency(data.get("total_investment", 0)))
            with col2_res: st.metric("📈 Penghematan Tahunan Bersih (IDR)", format_currency(data.get("annual_net_savings", 0)), help="Total penghematan kotor dikurangi biaya operasional tahunan.")
            with col3_res: st.metric("⏳ Payback Period", format_months(data.get("payback_period_months", float("inf"))), help="Waktu agar penghematan bersih kumulatif menutupi investasi awal.")
            st.markdown("##### Proyeksi Return on Investment (ROI)")
            col1_roi, col2_roi, col3_roi = st.columns(3)
            with col1_roi: st.metric("ROI 1 Tahun", format_roi(data.get("roi_1_year", 0)))
            with col2_roi: st.metric("ROI 3 Tahun", format_roi(data.get("roi_3_year", 0)))
            with col3_roi: st.metric("ROI 5 Tahun", format_roi(data.get("roi_5_year", 0)))
            st.markdown("---")
            # Charts
            st.subheader("📈 Visualisasi Proyeksi")
            if data.get("total_monthly_net_savings", 0) <= 0 and data.get("total_investment", 0) > 0:
                 st.warning("Penghematan bulanan bersih tidak positif. Grafik arus kas tidak akan menunjukkan breakeven.")
            elif data.get("total_investment", 0) == 0 and data.get("total_monthly_net_savings", 0) > 0:
                 st.info("Investasi awal nol dengan penghematan bersih positif. ROI sangat tinggi.")
                 # Regenerate charts for display if needed (or use stored ones if complex)
                 display_charts = generate_charts(data)
                 if display_charts.get("cashflow"): st.pyplot(display_charts["cashflow"])
                 if display_charts.get("savings"): st.pyplot(display_charts["savings"])
            elif data.get("total_investment", 0) == 0 and data.get("total_monthly_net_savings", 0) <= 0:
                 st.warning("Investasi awal nol dan penghematan bulanan bersih tidak positif. Analisis ROI tidak bermakna.")
            else:
                 display_charts = generate_charts(data)
                 if display_charts.get("cashflow"): st.pyplot(display_charts["cashflow"])
                 if display_charts.get("savings"): st.pyplot(display_charts["savings"])
            st.markdown("---")
            # Details Expander
            with st.expander("🔍 Lihat Detail Perhitungan & Asumsi"):
                # [Details content omitted for brevity - same as previous version]
                st.subheader("Komponen Penghematan Kotor Bulanan (Estimasi)")
                st.markdown(f"- Efisiensi Staff: `{format_currency(data.get('staff_savings_monthly', 0))}`")
                st.markdown(f"- Pengurangan No-Show: `{format_currency(data.get('noshow_savings_monthly', 0))}`")
                st.markdown(f"**= Total Penghematan Kotor Bulanan: `{format_currency(data.get('total_monthly_gross_savings', 0))}`**")
                st.markdown("&nbsp;")
                st.subheader("Biaya Operasional Berjalan Bulanan (Estimasi IDR)")
                st.markdown(f"- Langganan Bulanan: `{format_currency(data.get('monthly_subscription_idr', 0))}`")
                st.markdown(f"- Pemeliharaan Lainnya: `{format_currency(data.get('maintenance_cost_monthly', 0))}`")
                st.markdown(f"**= Total Biaya Operasional Bulanan: `{format_currency(data.get('total_monthly_operational_cost', 0))}`**")
                st.markdown("&nbsp;")
                st.subheader("Penghematan Bersih Bulanan (Estimasi IDR)")
                st.markdown(f"- Total Penghematan Kotor Bulanan: `{format_currency(data.get('total_monthly_gross_savings', 0))}`")
                st.markdown(f"- Dikurangi Total Biaya Operasional Bulanan: `{format_currency(data.get('total_monthly_operational_cost', 0))}`")
                st.markdown(f"**= Total Penghematan Bersih Bulanan: `{format_currency(data.get('total_monthly_net_savings', 0))}`**")
                st.markdown(f"**= Total Penghematan Bersih Tahunan: `{format_currency(data.get('annual_net_savings', 0))}`**")
                st.markdown("&nbsp;")
                st.subheader("Rincian Biaya Investasi Awal (Estimasi IDR)")
                st.markdown(f"- Setup: `{format_currency(data.get('setup_cost_idr', 0))}`")
                st.markdown(f"- Integrasi: `{format_currency(data.get('integration_cost_idr', 0))}`")
                st.markdown(f"- Pelatihan: `{format_currency(data.get('training_cost_idr', 0))}`")
                st.markdown(f"**= Total Investasi Awal: `{format_currency(data.get('total_investment', 0))}`**")
                st.markdown("&nbsp;")
                st.subheader("Asumsi Dasar")
                # [Assumptions content omitted for brevity]
                col_a1, col_a2 = st.columns(2)
                with col_a1:
                    st.markdown(f"- Janji Temu/Bulan: `{data.get('monthly_appointments', 0):,.0f}`")
                    st.markdown(f"- No-Show Awal: `{format_percent(data.get('noshow_rate_before', 0))}`")
                    st.markdown(f"- Target Reduksi No-Show: `{format_percent(data.get('noshow_reduction_target', 0))}`")
                    st.markdown(f"- Staff Admin: `{data.get('admin_staff', 0):,.0f}` org")
                with col_a2:
                    st.markdown(f"- Gaji Staff Admin: `{format_currency(data.get('avg_salary', 0))}`/Bln")
                    st.markdown(f"- Target Efisiensi Staff: `{format_percent(data.get('staff_reduction_target', 0))}`")
                    st.markdown(f"- Pendapatan/Janji Temu: `{format_currency(data.get('revenue_per_appointment', 0))}`")
                    st.markdown(f"- Kurs USD-IDR: `{format_currency(data.get('exchange_rate', 0))}`")

            # Footer
            st.markdown("---")
            st.caption(f"© {datetime.now().year} Medical AI Solutions Inc. | Analisis ini adalah estimasi.")

        # --- PDF Download and Upload Section --- 
        with pdf_upload_placeholder.container():
            if st.session_state.get("pdf_bytes") and st.session_state.get("pdf_filename"):
                pdf_filename = st.session_state.pdf_filename
                pdf_bytes = st.session_state.pdf_bytes
                st.markdown("--- ")
                st.subheader("📄 Unduh & Unggah Laporan PDF")

                col_dl, col_ul = st.columns([1, 2])

                with col_dl:
                    st.download_button(
                        label="📥 Unduh Laporan PDF",
                        data=pdf_bytes,
                        file_name=pdf_filename,
                        mime="application/pdf",
                        key="download_pdf_button"
                    )

                with col_ul:
                    # Display upload status
                    upload_status = st.session_state.get("drive_upload_status")
                    if upload_status == "success":
                        st.success(f"Laporan PDF berhasil diunggah ke Google Drive! [Lihat File]({st.session_state.drive_file_link})", icon="✅")
                    elif upload_status == "error":
                        st.error("Gagal mengunggah laporan PDF ke Google Drive.", icon="❌")
                    elif upload_status == "pending":
                        st.info("Sedang mengunggah laporan ke Google Drive...", icon="⏳")
                    else:
                        # Show upload button only if not already uploaded or pending
                        if GOOGLE_API_AVAILABLE and WEASYPRINT_AVAILABLE:
                            if st.button("☁️ Unggah ke Google Drive", key="upload_drive_button"):
                                st.session_state.drive_upload_status = "pending"
                                st.rerun() # Rerun to show pending status and trigger upload logic below
                        else:
                            st.warning("Fungsi unggah Google Drive tidak tersedia (cek library).", icon="⚠️")
            elif st.session_state.get("results_calculated"): # If calculation done but PDF failed
                 st.error("Pembuatan Laporan PDF gagal, tidak dapat mengunduh atau mengunggah.")

    # --- Google Drive Upload Logic (runs after button press + rerun) ---
    if st.session_state.get("drive_upload_status") == "pending":
        pdf_bytes = st.session_state.get("pdf_bytes")
        pdf_filename = st.session_state.get("pdf_filename")
        hospital_name_for_folder = sanitize_filename(st.session_state.report_data.get("hospital_name", "UnknownHospital"))

        if pdf_bytes and pdf_filename and hospital_name_for_folder:
            drive_service = get_drive_service()
            if drive_service:
                target_subfolder_id = find_or_create_folder(drive_service, hospital_name_for_folder, PARENT_FOLDER_ID)
                if target_subfolder_id:
                    file_link = upload_pdf_to_drive(drive_service, pdf_bytes, pdf_filename, target_subfolder_id)
                    if file_link:
                        st.session_state.drive_upload_status = "success"
                        st.session_state.drive_file_link = file_link
                    else:
                        st.session_state.drive_upload_status = "error"
                else:
                    st.session_state.drive_upload_status = "error" # Error finding/creating folder
            else:
                st.session_state.drive_upload_status = "error" # Error getting service
        else:
            st.session_state.drive_upload_status = "error" # PDF data missing
            st.error("Data PDF tidak ditemukan untuk diunggah.")

        # Rerun one last time to display the final success/error message
        st.rerun()

# ====================== EKSEKUSI APLIKASI ======================
if __name__ == "__main__":
    setup_locale()
    main()


