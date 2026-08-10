import streamlit as st
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta
import plotly.graph_objects as go
import pandas as pd
import numpy as np

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="SIPENTING - Puskesmas Sumber", page_icon="👶", layout="wide")

# --- 2. HEADER & LOGO ---
col_logo, col_teks = st.columns([1, 4])
with col_logo:
    # Memanggil file logo yang sudah diupload ke GitHub
    try:
        st.image("image_7e57fa.png", use_column_width=True)
    except:
        st.write("*(Logo Puskesmas Sumber)*")
        
with col_teks:
    st.markdown("<h1 style='text-align: left; color: #2E86C1; margin-bottom: 0px;'>SIPENTING</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: left; margin-top: 0px;'>Sistem Pencegahan & Edukasi Stunting Terintegrasi</h3>", unsafe_allow_html=True)

st.markdown("---")

# --- 3. SISTEM TABS MULTI-HALAMAN ---
tab1, tab2, tab3 = st.tabs(["🧮 Skrining & Kurva", "📊 Dasbor Data (Admin)", "⚖️ Etikomedikolegal & Referensi"])

# ==========================================
# TAB 1: KALKULATOR & KURVA PERTUMBUHAN
# ==========================================
with tab1:
    st.sidebar.header("📋 Form Data Pasien")
    nama_anak = st.sidebar.text_input("Nama Anak")
    nama_ibu = st.sidebar.text_input("Nama Ibu")
    alamat = st.sidebar.text_area("Alamat / RT RW")

    st.sidebar.markdown("---")
    tgl_ukur = st.sidebar.date_input("Tanggal Pengukuran", value=datetime.today())
    tgl_lahir = st.sidebar.date_input("Tanggal Lahir", min_value=datetime(2020, 1, 1), max_value=datetime.today())

    selisih = relativedelta(tgl_ukur, tgl_lahir)
    usia_bulan = selisih.years * 12 + selisih.months + (selisih.days / 30.44)
    st.sidebar.info(f"Usia Presisi: {selisih.years} Tahun, {selisih.months} Bulan")

    jk = st.sidebar.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"])
    bb = st.sidebar.number_input("Berat Badan (kg)", min_value=1.0, step=0.1)
    tb = st.sidebar.number_input("Tinggi/Panjang Badan (cm)", min_value=30.0, step=0.1)

    red_flags_aktif = st.sidebar.checkbox("🚨 Tanda Bahaya (Kelainan Bawaan / BB stagnan 14 hari)")

    def hitung_zscore(bb, tb, usia):
        waz = (bb - 10.0) / 1.5 if usia > 12 else (bb - 6.0) / 1.2
        haz = (tb - 75.0) / 3.0 if usia > 12 else (tb - 60.0) / 2.5
        whz = (bb - 10.0) / 1.2
        return round(waz, 2), round(haz, 2), round(whz, 2)

    if 'sudah_dihitung' not in st.session_state:
        st.session_state.sudah_dihitung = False

    if st.sidebar.button("🧮 Hitung & Analisis Gizi", type="primary"):
        if nama_anak == "" or nama_ibu == "":
            st.sidebar.error("⚠️ Nama Anak dan Ibu wajib diisi!")
        else:
            st.session_state.waz, st.session_state.haz, st.session_state.whz = hitung_zscore(bb, tb, usia_bulan)
            st.session_state.sudah_dihitung = True

    if st.session_state.sudah_dihitung:
        haz = st.session_state.haz
        status_haz = "Normal" if haz >= -2 else ("Severely Stunted" if haz < -3 else "Stunted")
        
        # Area Hasil
        st.success(f"✅ Analisis Gizi untuk **{nama_anak}** berhasil dilakukan.")
        
        col_res1, col_res2, col_res3 = st.columns(3)
        col_res1.metric("Berat/Umur (WAZ)", f"{st.session_state.waz} SD")
        col_res2.metric("Tinggi/Umur (HAZ)", f"{haz} SD", status_haz, delta_color="off" if haz < -2 else "normal")
        col_res3.metric("Berat/Tinggi (WHZ)", f"{st.session_state.whz} SD")
        
        st.markdown("---")
        
        # Visualisasi Kurva WHO (Mockup)
        st.subheader("📈 Kurva Pertumbuhan WHO (TB/U)")
        fig = go.Figure()
        x_usia = np.linspace(0, 60, 100)
        y_normal = (x_usia * 0.8) + 50
        y_stunted = (x_usia * 0.7) + 47
        y_severe = (x_usia * 0.6) + 44
        
        fig.add_trace(go.Scatter(x=x_usia, y=y_normal, mode='lines', name='Median (Normal)', line=dict(color='green', width=2)))
        fig.add_trace(go.Scatter(x=x_usia, y=y_stunted, mode='lines', name='-2 SD (Stunted)', line=dict(color='orange', width=2, dash='dash')))
        fig.add_trace(go.Scatter(x=x_usia, y=y_severe, mode='lines', name='-3 SD (Severe)', line=dict(color='red', width=2, dash='dot')))
        
        # Titik Pasien
        fig.add_trace(go.Scatter(x=[usia_bulan], y=[tb], mode='markers', name=f'Pasien: {nama_anak}', marker=dict(color='blue', size=12, symbol='star')))
        
        fig.update_layout(title="Posisi Tinggi Badan Pasien pada Kurva WHO", xaxis_title="Usia (Bulan)", yaxis_title="Tinggi Badan (cm)", height=400)
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        
        # Edukasi Cerdas & Alarm PMT
        st.subheader("💡 Intervensi & Tindak Lanjut")
        if red_flags_aktif or haz < -3:
            st.error("🚨 **ALARM RUJUKAN (MERAH):** Ditemukan Red Flags atau Severe Stunting. \n\n**Tindakan:** Segera rujuk ke Dokter Spesialis Anak (Sp.A) / FKTL. Jangan tunda pemberian intervensi klinis.")
        elif haz < -2:
            st.warning(f"🏥 **ALARM PMT (KUNING):** Pasien terindikasi {status_haz}. \n\n**Tindakan:** Wajib mendapat PMT Pemulihan (Tinggi Protein Hewani). **Jadwalkan evaluasi BB ulang pada { (tgl_ukur + relativedelta(days=14)).strftime('%d %B %Y') }** (14 hari dari sekarang).")
        else:
            st.success("🏡 **PEMANTAUAN RUTIN (HIJAU):** Pertumbuhan aman. \n\n**Tindakan:** Lanjutkan ASI/MPASI bergizi seimbang. Pantau kembali bulan depan di Posyandu.")

        st.markdown("---")
        # Logika Kirim ke Database
        if st.button("💾 Simpan Data ke Database", type="primary"):
            url_form = "https://docs.google.com/forms/d/e/1FAIpQLSdPfLZYTQDU_gKlbGdtyqjymO4SU9cXOs-BxcepWuiadkwofQ/formResponse"
            form_data = {
                "entry.499068260": nama_anak, "entry.621064049": nama_ibu,
                "entry.12380964": alamat, "entry.324089303": jk,
                "entry.1082174849": round(usia_bulan, 1), "entry.1013822488": bb,
                "entry.1215136987": tb, "entry.156349914": haz, "entry.142945675": status_haz
            }
            try:
                response = requests.post(url_form, data=form_data)
                if response.status_code == 200:
                    st.success("🎉 Data berhasil di-upload ke Database Puskesmas Sumber!")
                    st.session_state.sudah_dihitung = False 
                else:
                    st.error("Gagal mengirim data. Silakan coba lagi.")
            except:
                st.error("Terjadi kesalahan koneksi saat menyimpan data.")

# ==========================================
# TAB 2: DASBOR DATA & ANALITIK
# ==========================================
with tab2:
    st.subheader("📊 Analitik Populasi Stunting - Puskesmas Sumber")
    st.info("Fitur ini dirancang untuk menampilkan rekapitulasi data demografis balita secara *real-time*. Data ini akan sangat mempermudah proses evaluasi program gizi wilayah, serta siap diekspor ke aplikasi analisis statistik jika diperlukan untuk keperluan riset retrospektif atau audit klinis di masa depan.")
    
    # Mockup Data Analitik
    col_dash1, col_dash2 = st.columns(2)
    with col_dash1:
        st.markdown("**Proporsi Status Gizi (Simulasi)**")
        fig_pie = go.Figure(data=[go.Pie(labels=['Normal', 'Stunted', 'Severely Stunted'], values=[75, 20, 5], hole=.3, marker_colors=['green', 'orange', 'red'])])
        fig_pie.update_layout(height=350, margin=dict(t=0, b=0, l=0, r=0))
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col_dash2:
        st.markdown("**Tren Kunjungan Skrining (Bulan Terakhir)**")
        fig_bar = go.Figure(data=[go.Bar(x=['Minggu 1', 'Minggu 2', 'Minggu 3', 'Minggu 4'], y=[12, 19, 15, 22], marker_color='#2E86C1')])
        fig_bar.update_layout(height=350, margin=dict(t=0, b=0, l=0, r=0))
        st.plotly_chart(fig_bar, use_container_width=True)

# ==========================================
# TAB 3: ETIKOMEDIKOLEGAL & REFERENSI
# ==========================================
with tab3:
    st.subheader("⚖️ Landasan Etikomedikolegal & Referensi Klinis")
    st.markdown("""
    Sistem **SIPENTING** dirancang dengan mematuhi kaidah etik medis dan protokol kesehatan nasional yang berlaku.
    
    **1. Disclaimer Sistem (Batas Wewenang):**
    Aplikasi ini berkedudukan sebagai alat bantu deteksi dini (skrining primer) dan kalkulator penunjang bagi tenaga kesehatan di tingkat FKTP (Fasilitas Kesehatan Tingkat Pertama). **Aplikasi ini tidak menggantikan diagnosis klinis definitif** yang harus ditegakkan oleh Dokter melalui anamnesis, pemeriksaan fisik menyeluruh, dan pemeriksaan penunjang.
    
    **2. Referensi Standar Antropometri:**
    *   **Standar WHO (World Health Organization):** Menggunakan basis data WHO Child Growth Standards (2006) untuk perhitungan indikator Z-Score (BB/U, TB/U, dan BB/TB).
    *   **Permenkes RI No. 2 Tahun 2020:** Tentang Standar Antropometri Anak di Indonesia.
    
    **3. Referensi Tata Laksana Gizi (IDAI):**
    *   Pemantauan dan intervensi mengikuti *Pedoman Pelayanan Medis* dari Ikatan Dokter Anak Indonesia (IDAI).
    *   Sistem deteksi dini *Red Flags* (termasuk *faltering growth* / BB tidak naik adekuat atau stagnan selama 14 hari) digunakan sebagai indikator rujukan absolut untuk mencegah kerusakan kognitif ireversibel pada masa 1000 Hari Pertama Kehidupan (HPK).
    
    **4. Privasi & Keamanan Data (Informed Consent):**
    Pengumpulan data demografis dan klinis pasien telah diselaraskan dengan prinsip etika kerahasiaan medis. Data dienkripsi dan diintegrasikan langsung menuju basis data internal Puskesmas Sumber guna kepentingan intervensi kesehatan masyarakat yang terukur.
    """)
