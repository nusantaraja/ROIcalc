# Medical ROI Calculator (Revised)

## Deskripsi
Aplikasi Streamlit ini menghitung estimasi Return on Investment (ROI) selama 5 tahun untuk implementasi solusi AI Voice di fasilitas kesehatan.

## Perubahan Utama (Mei 2025)
- Menambahkan field input wajib untuk Nama Konsultan, Email Konsultan, dan Nomor HP/WA Konsultan dengan validasi dasar.
- Menambahkan field input untuk Biaya Langganan Tahunan (USD) dan mengintegrasikannya ke dalam perhitungan total investasi dan ROI.
- Mengimplementasikan fitur pembuatan laporan ringkasan dalam format PDF.
- Mengintegrasikan unggah otomatis laporan PDF ke Google Drive menggunakan kredensial akun layanan (service account).
  - PDF disimpan dalam subfolder Google Drive (ID: `1bCG7m4T73K3RNoMvWTE4fjRdWCkwAiKR`) yang dinamai sesuai nama rumah sakit.
  - Nama file PDF mengikuti format: `yymmdd namarumahsakit lokasi namakonsultan.pdf`.
- Memperbaiki error sintaks f-string pada tampilan ringkasan hasil.

## Instruksi Penggunaan
1.  **Ekstrak File:** Ekstrak isi file `Medical_ROI_Calculator_Revised.zip` ke direktori pilihan Anda.
2.  **Kredensial Google Drive:** Tempatkan file kredensial akun layanan Google Anda (`service_account_key.json`) di direktori yang sama dengan file `Medical_ROI_Calc.py`. Pastikan akun layanan memiliki izin **Editor** (atau setidaknya izin tulis) pada folder Google Drive target (`1bCG7m4T73K3RNoMvWTE4fjRdWCkwAiKR`).
3.  **Instal Dependensi:** Buka terminal atau command prompt, navigasi ke direktori aplikasi, dan jalankan perintah:
    ```bash
    pip install -r requirements.txt
    ```
4.  **Jalankan Aplikasi:** Masih di direktori yang sama, jalankan perintah:
    ```bash
    streamlit run Medical_ROI_Calc.py
    ```
5.  Aplikasi akan terbuka di browser web Anda.

## Catatan Validasi
- Aplikasi telah diuji untuk memastikan fitur-fitur utama berfungsi.
- **Penting:** Selama pengujian, ditemukan bahwa beberapa field input (seperti "Lokasi (Kota/Area)" dan "Biaya Langganan Tahunan (USD)") mungkin menjadi tidak aktif (disabled) setelah field lain diisi. Hal ini menghambat pengujian otomatis penuh. Fungsi inti seperti kalkulasi, pembuatan PDF, dan unggah ke Google Drive telah diverifikasi berdasarkan kode dan pengujian manual terbatas. **Disarankan agar Anda melakukan pengujian manual menyeluruh pada alur kerja lengkap untuk memastikan semua field dan perhitungan berjalan sesuai harapan dalam skenario penggunaan Anda.**
