import streamlit as st
import matplotlib.pyplot as plt
from datetime import datetime
import os
import base64
from io import BytesIO

# Konfigurasi halaman
st.set_page_config(
    page_title="Kalkulator ROI AI Voice",
    page_icon="📈",
    layout="wide"
)

# Fungsi untuk generate laporan
def generate_report(params, fig):
    md_content = f"""# Analisis ROI Solusi AI Voice untuk {params['broker_name']}

**Tanggal Analisis:** {datetime.now().strftime('%d %B %Y')}  
**Broker:** {params['broker_name']}  
**Lokasi:** {params['broker_location']}  

## Informasi Operasional Saat Ini

| Metrik | Nilai |
|--------|-------|
| Jumlah Staf Layanan Pelanggan | {params['cs_staff']} |
| Biaya Tenaga Kerja Tahunan | IDR {params['current_annual_labor_cost_idr']:,.2f} (USD {params['current_annual_labor_cost_usd']:,.2f}) |
| Pertanyaan Tahunan | {params['current_inquiries_per_year']:,} |
| Jam Penanganan yang Dibutuhkan | {params['current_handling_hours']:,.2f} |
| Tingkat Loyalitas Klien | {params['current_retention_rate']}% |

## Proyeksi dengan Solusi AI Voice

| Metrik | Nilai |
|--------|-------|
| Jumlah Staf | {params['new_staff_count']} |
| Biaya Tenaga Kerja Tahunan | IDR {params['new_annual_labor_cost_idr']:,.2f} (USD {params['new_annual_labor_cost_usd']:,.2f}) |
| Pertanyaan Diotomatisasi | {params['automated_inquiries']:,.0f} ({params['automation_rate']}%) |
| ROI Tahun Pertama | {params['first_year_roi']:.2f}% |
| ROI 3 Tahun | {params['three_year_roi']:.2f}% |

![Proyeksi 5 Tahun](data:image/png;base64,{fig})
"""
    return md_content

# UI Streamlit
def main():
    st.title("📊 Kalkulator ROI Solusi AI Voice")
    st.caption("Aplikasi untuk menghitung Return on Investment implementasi solusi AI Voice")

    with st.sidebar:
        st.header("🔧 Parameter Input")
        
        st.subheader("Informasi Broker")
        broker_name = st.text_input("Nama Broker Forex")
        broker_location = st.text_input("Lokasi Broker")
        
        st.subheader("Data Operasional")
        cs_staff = st.number_input("Jumlah Staf Layanan Pelanggan", min_value=1, value=5)
        avg_monthly_salary = st.number_input("Rata-rata Gaji Bulanan per Staf (IDR)", min_value=0, value=15000000)
        overhead_multiplier = st.slider("Pengali Overhead", 1.0, 2.0, 1.3)
        usd_conversion_rate = st.number_input("Kurs IDR ke USD", min_value=1.0, value=15500.0)
        
        monthly_inquiries = st.number_input("Pertanyaan Bulanan", min_value=0, value=500)
        avg_handling_time = st.number_input("Waktu Penanganan per Pertanyaan (menit)", min_value=0.0, value=15.0)
        
        avg_monthly_clients = st.number_input("Klien Aktif Bulanan", min_value=0, value=200)
        avg_monthly_client_value = st.number_input("Pendapatan Bulanan per Klien (USD)", min_value=0.0, value=150.0)
        current_retention_rate = st.slider("Tingkat Loyalitas Klien (%)", 0, 100, 70)
        
        st.subheader("Biaya AI Voice")
        implementation_cost = st.number_input("Biaya Implementasi (USD)", min_value=0.0, value=10000.0)
        annual_subscription = st.number_input("Biaya Langganan Tahunan (USD)", min_value=0.0, value=5000.0)
        
        st.subheader("Asumsi Dampak")
        automation_rate = st.slider("Persentase Otomatisasi (%)", 0, 100, 70)
        staff_reduction = st.slider("Pengurangan Staf (%)", 0, 100, 30)
        retention_improvement = st.slider("Peningkatan Loyalitas (% point)", 0, 20, 5)
        handling_time_improvement = st.slider("Peningkatan Efisiensi Waktu (%)", 0, 50, 20)

    # Kalkulasi
    if st.button("🚀 Hitung ROI"):
        with st.spinner('Menghitung...'):
            # Konversi ke tahunan
            avg_annual_salary = avg_monthly_salary * 12
            avg_client_value = avg_monthly_client_value * 12
            
            # Perhitungan (sama seperti script asli)
            avg_annual_salary_usd = avg_annual_salary / usd_conversion_rate
            current_annual_labor_cost_usd = cs_staff * avg_annual_salary_usd * overhead_multiplier
            current_annual_labor_cost_idr = current_annual_labor_cost_usd * usd_conversion_rate
            current_inquiries_per_year = monthly_inquiries * 12
            current_handling_hours = (current_inquiries_per_year * avg_handling_time) / 60
            
            automated_inquiries = current_inquiries_per_year * (automation_rate / 100)
            remaining_manual_inquiries = current_inquiries_per_year - automated_inquiries
            new_handling_time = avg_handling_time * (1 - handling_time_improvement / 100)
            new_handling_hours = (remaining_manual_inquiries * new_handling_time) / 60
            
            new_staff_count = round(cs_staff * (1 - staff_reduction / 100))
            new_annual_labor_cost_usd = new_staff_count * avg_annual_salary_usd * overhead_multiplier
            new_annual_labor_cost_idr = new_annual_labor_cost_usd * usd_conversion_rate
            
            labor_savings_usd = current_annual_labor_cost_usd - new_annual_labor_cost_usd
            labor_savings_idr = labor_savings_usd * usd_conversion_rate
            
            new_retention_rate = min(100, current_retention_rate + retention_improvement)
            current_churned_clients = avg_monthly_clients * (1 - current_retention_rate / 100)
            new_churned_clients = avg_monthly_clients * (1 - new_retention_rate / 100)
            clients_saved = current_churned_clients - new_churned_clients
            retention_revenue_impact = clients_saved * avg_client_value
            
            total_annual_savings_usd = labor_savings_usd + retention_revenue_impact
            total_annual_savings_idr = labor_savings_idr + (retention_revenue_impact * usd_conversion_rate)
            first_year_net_usd = total_annual_savings_usd - implementation_cost - annual_subscription
            subsequent_years_net_usd = total_annual_savings_usd - annual_subscription
            
            first_year_roi = (first_year_net_usd / (implementation_cost + annual_subscription)) * 100
            three_year_roi = ((first_year_net_usd + subsequent_years_net_usd * 2) / (implementation_cost + annual_subscription * 3)) * 100
            
            monthly_savings_usd = total_annual_savings_usd / 12
            total_investment_usd = implementation_cost + annual_subscription
            payback_period = total_investment_usd / monthly_savings_usd
            
            # Grafik
            years = range(1, 6)
            costs = [implementation_cost + annual_subscription if year == 1 else annual_subscription for year in years]
            benefits = [total_annual_savings_usd for _ in years]
            net_benefits = [benefits[i] - costs[i] for i in range(len(years))]
            cumulative_net = [sum(net_benefits[:i+1]) for i in range(len(years))]
            
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.plot(years, cumulative_net, marker='o', color='#4CAF50')
            ax.set_title('Proyeksi 5 Tahun (USD)')
            ax.set_xlabel('Tahun')
            ax.set_ylabel('Manfaat Kumulatif')
            ax.grid(True)
            
            # Simpan gambar ke base64
            buf = BytesIO()
            plt.savefig(buf, format="png")
            plt.close()
            fig_base64 = base64.b64encode(buf.getvalue()).decode("utf-8")
            
            # Parameter untuk laporan
            params = {
                'broker_name': broker_name,
                'broker_location': broker_location,
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
                'first_year_roi': first_year_roi,
                'three_year_roi': three_year_roi
            }
            
            # Hasil utama
            st.success("Perhitungan Selesai!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("ROI Tahun Pertama", f"{first_year_roi:.2f}%")
                st.metric("Penghematan Tahunan", f"USD {total_annual_savings_usd:,.2f}")
                
            with col2:
                st.metric("ROI 3 Tahun", f"{three_year_roi:.2f}%")
                st.metric("Periode Pengembalian", f"{payback_period:.1f} bulan")
            
            st.pyplot(fig)
            
            # Generate dan download laporan
            report = generate_report(params, fig_base64)
            st.download_button(
                label="📥 Download Laporan (MD)",
                data=report,
                file_name=f"ROI_AI_Voice_{broker_name.replace(' ','_')}.md",
                mime="text/markdown"
            )

if __name__ == "__main__":
    main()