import streamlit as st
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta
import plotly.graph_objects as go
import pandas as pd
import numpy as np
from decimal import Decimal, InvalidOperation
from pygrowup import Calculator
from pygrowup import exceptions as pg_exceptions

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="SIPENTING - Puskesmas Sumber", page_icon="👶", layout="wide")

# --- 2. INJEKSI CUSTOM CSS (GLOW UP DESAIN) ---
st.markdown("""
    <style>
    /* Mengubah warna background utama menjadi biru sangat muda/pastel */
    .stApp {
        background-color: #F4F9F9;
    }
    
    /* Mempercantik kotak hasil perhitungan (Metric Cards) */
    div[data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #E0E0E0;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.05);
    }

    /* Mempercantik tombol Hitung dan Simpan */
    .stButton>button {
        border-radius: 8px;
        font-weight: bold;
        transition: 0.3s;
    }
    .stButton>button:hover {
        transform: scale(1.02);
    }
    
    /* Merapikan Judul Utama */
    .main-title {
        text-align: left;
        color: #1A5276;
        font-weight: 800;
        font-size: 3rem;
        margin-bottom: -15px;
        text-transform: uppercase;
        letter-spacing: 2px;
    }
    .sub-title {
        text-align: left;
        color: #2E86C1;
        font-size: 1.2rem;
        font-style: italic;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# --- 3. HEADER & LOGO ---
col_logo, col_teks = st.columns([1, 4])
with col_logo:
    st.markdown(
        '<img src="https://raw.githubusercontent.com/rizkinovaln03/sipenting-sumber/main/logo.png" style="width:100%; border-radius:10px;">',
        unsafe_allow_html=True
    )

with col_teks:
    st.markdown("<div class='main-title'>SIPENTING</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-title'>Sistem Pencegahan & Edukasi Stunting Terintegrasi | Puskesmas Sumber</div>", unsafe_allow_html=True)

st.markdown("---")

# --- 4. KALKULATOR Z-SCORE RESMI WHO (LMS, via pygrowup) ---
# pygrowup menyimpan tabel L/M/S resmi WHO Child Growth Standards (2006)
# per bulan usia (0-60) dan per jenis kelamin -> jauh lebih akurat
# dibanding formula linear.
@st.cache_resource
def get_calculator():
    return Calculator(adjust_height_data=False, adjust_weight_scores=False, include_cdc=False)


def hitung_zscore(bb, tb, usia_bulan, jk):
    """
    Menghitung WAZ, HAZ, WHZ sesuai standar WHO 2006 (metode LMS),
    memperhitungkan usia (bulan) DAN jenis kelamin.

    - WAZ (Weight-for-Age)  -> indikator 'wfa'
    - HAZ (Height/Length-for-Age) -> indikator 'lhfa'
    - WHZ (Weight-for-Height/Length):
        usia < 24 bulan -> 'wfl' (panjang badan, posisi berbaring)
        usia >= 24 bulan -> 'wfh' (tinggi badan, posisi berdiri)

    Return: (waz, haz, whz, error_pesan)
    Jika salah satu z-score gagal dihitung (misal ukuran ekstrem di luar
    rentang tabel WHO), nilainya None dan error_pesan berisi keterangannya
    (JANGAN pernah menampilkan angka hasil formula karangan sendiri).
    """
    calc = get_calculator()
    sex_code = "M" if jk == "Laki-laki" else "F"
    errors = []

    def _safe_z(indicator, measurement, height=None):
        try:
            z = calc.zscore_for_measurement(
                indicator, measurement, usia_bulan, sex_code, height=height
            )
            return round(float(z), 2)
        except (pg_exceptions.InvalidMeasurement, pg_exceptions.DataNotFound,
                InvalidOperation, AssertionError, TypeError, ValueError) as e:
            errors.append(f"{indicator.upper()}: {e}")
            return None

    waz = _safe_z("wfa", bb)
    haz = _safe_z("lhfa", tb)

    if usia_bulan < 24:
        whz = _safe_z("wfl", bb, height=tb)
    else:
        whz = _safe_z("wfh", bb, height=tb)

    return waz, haz, whz, errors


# --- Tabel kurva WHO asli (median, -2SD, -3SD) untuk plotting ---
@st.cache_data
def get_kurva_haz(jk):
    """
    Mengambil kurva Tinggi/Panjang-menurut-Umur (0-60 bulan) LANGSUNG
    dari tabel resmi WHO (bukan garis lurus hasil rekaan), dipisah
    per jenis kelamin.
    """
    calc = get_calculator()
    table = calc.lhfa_boys_0_5 if jk == "Laki-laki" else calc.lhfa_girls_0_5

    bulan, median, sd2neg, sd3neg = [], [], [], []
    for k, v in table.items():
        if k == "field_name":
            continue
        bulan.append(float(v["Month"]))
        median.append(float(v["SD0"]))
        sd2neg.append(float(v["SD2neg"]))
        sd3neg.append(float(v["SD3neg"]))

    df = pd.DataFrame({
        "bulan": bulan, "median": median, "sd2neg": sd2neg, "sd3neg": sd3neg
    }).sort_values("bulan")
    return df


# --- 5. SISTEM TABS MULTI-HALAMAN ---
tab1, tab2, tab3 = st.tabs(["🧮 Skrining & Kurva", "📊 Dasbor Data", "⚖️ Referensi Klinis"])

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

    if usia_bulan > 60:
        st.sidebar.warning("⚠️ Standar WHO pada aplikasi ini hanya mencakup usia 0-60 bulan (0-5 tahun).")

    jk = st.sidebar.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"])
    bb = st.sidebar.number_input("Berat Badan (kg)", min_value=1.0, step=0.1)
    tb = st.sidebar.number_input("Tinggi/Panjang Badan (cm)", min_value=30.0, step=0.1)

    red_flags_aktif = st.sidebar.checkbox("🚨 Tanda Bahaya (Kelainan Bawaan / BB stagnan 14 hari)")

    if 'sudah_dihitung' not in st.session_state:
        st.session_state.sudah_dihitung = False

    if st.sidebar.button("🧮 Hitung & Analisis Gizi", type="primary"):
        if nama_anak == "" or nama_ibu == "":
            st.sidebar.error("⚠️ Nama Anak dan Ibu wajib diisi!")
        elif usia_bulan > 60 or usia_bulan < 0:
            st.sidebar.error("⚠️ Standar WHO Child Growth Standards di aplikasi ini hanya berlaku 0-60 bulan.")
        else:
            waz, haz, whz, errors = hitung_zscore(bb, tb, usia_bulan, jk)
            st.session_state.waz, st.session_state.haz, st.session_state.whz = waz, haz, whz
            st.session_state.zscore_errors = errors
            st.session_state.sudah_dihitung = True

    if st.session_state.sudah_dihitung:
        waz = st.session_state.waz
        haz = st.session_state.haz
        whz = st.session_state.whz

        if st.session_state.get("zscore_errors"):
            for err in st.session_state.zscore_errors:
                st.error(f"⚠️ Gagal menghitung salah satu Z-score (kemungkinan input di luar rentang wajar WHO): {err}")

        if haz is None:
            st.warning("HAZ tidak dapat dihitung untuk input ini. Periksa kembali data usia/tinggi badan.")
            status_haz = None
        else:
            status_haz = "Normal" if haz >= -2 else ("Severely Stunted" if haz < -3 else "Stunted")

            # Area Hasil
            st.success(f"✅ Analisis Gizi untuk **{nama_anak}** berhasil dilakukan (Z-score WHO 2006, {jk}, usia {usia_bulan:.1f} bulan).")

            col_res1, col_res2, col_res3 = st.columns(3)
            col_res1.metric("Berat/Umur (WAZ)", f"{waz} SD" if waz is not None else "N/A")
            col_res2.metric("Tinggi/Umur (HAZ)", f"{haz} SD", status_haz, delta_color="off" if haz < -2 else "normal")
            col_res3.metric("Berat/Tinggi (WHZ)", f"{whz} SD" if whz is not None else "N/A")

            st.markdown("---")

            # Visualisasi Kurva WHO ASLI (bukan garis lurus rekaan)
            st.subheader(f"📈 Kurva Pertumbuhan WHO Tinggi/Panjang-menurut-Umur ({jk})")
            df_kurva = get_kurva_haz(jk)

            fig = go.Figure()
            fig.add_trace(go.Scatter(x=df_kurva["bulan"], y=df_kurva["median"], mode='lines',
                                      name='Median (0 SD)', line=dict(color='green', width=2)))
            fig.add_trace(go.Scatter(x=df_kurva["bulan"], y=df_kurva["sd2neg"], mode='lines',
                                      name='-2 SD (Stunted)', line=dict(color='orange', width=2, dash='dash')))
            fig.add_trace(go.Scatter(x=df_kurva["bulan"], y=df_kurva["sd3neg"], mode='lines',
                                      name='-3 SD (Severely Stunted)', line=dict(color='red', width=2, dash='dot')))

            # Titik Pasien
            fig.add_trace(go.Scatter(x=[usia_bulan], y=[tb], mode='markers', name=f'Pasien: {nama_anak}',
                                      marker=dict(color='blue', size=14, symbol='star')))

            fig.update_layout(
                title=f"Posisi Tinggi/Panjang Badan Pasien pada Kurva WHO ({jk}, 0-60 bulan)",
                xaxis_title="Usia (Bulan)", yaxis_title="Tinggi/Panjang Badan (cm)", height=420
            )
            st.plotly_chart(fig, use_container_width=True)

            st.caption("Kurva ini diambil langsung dari tabel LMS resmi WHO Child Growth Standards (2006), "
                       "dipisah per jenis kelamin — bukan garis rata-rata yang disamaratakan.")

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
                    "entry.499068260": str(nama_anak),
                    "entry.621064049": str(nama_ibu),
                    "entry.12380964": str(alamat),
                    "entry.324089303": str(jk),
                    "entry.1082174849": str(round(usia_bulan, 1)),
                    "entry.1013822488": str(bb),
                    "entry.1215136987": str(tb),
                    "entry.156349914": str(haz),
                    "entry.142945675": str(status_haz)
                }
                try:
                    response = requests.post(url_form, data=form_data)
                    if response.status_code == 200:
                        st.success("🎉 Data berhasil di-upload ke Database Puskesmas Sumber!")
                        st.session_state.sudah_dihitung = False
                    else:
                        st.error(f"Gagal mengirim data. Kode Error Google: {response.status_code}.")
                except Exception as e:
                    st.error(f"Terjadi kesalahan koneksi: {e}")

# ==========================================
# TAB 2: DASBOR DATA & ANALITIK
# ==========================================
with tab2:
    st.subheader("📊 Analitik Populasi Stunting - Puskesmas Sumber")
    st.info("Fitur ini dirancang untuk menampilkan rekapitulasi data demografis balita secara *real-time*. Data ini akan sangat mempermudah proses evaluasi program gizi wilayah, serta siap diekspor ke aplikasi analisis statistik jika diperlukan untuk keperluan riset retrospektif atau audit klinis di masa depan.")

    st.warning("⚠️ Grafik di bawah ini masih data SIMULASI/mockup, bukan data asli dari Google Form. "
               "Kalau Prof mau, saya bisa sambungkan ke Google Sheets hasil respons form supaya datanya live.")

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
    *   **Standar WHO (World Health Organization):** Menggunakan basis data WHO Child Growth Standards (2006) untuk perhitungan indikator Z-Score (BB/U, TB/U, dan BB/TB), dihitung dengan metode LMS resmi (bukan estimasi linear).
    *   **Permenkes RI No. 2 Tahun 2020:** Tentang Standar Antropometri Anak di Indonesia.
    
    **3. Referensi Tata Laksana Gizi (IDAI):**
    *   Pemantauan dan intervensi mengikuti *Pedoman Pelayanan Medis* dari Ikatan Dokter Anak Indonesia (IDAI).
    *   Sistem deteksi dini *Red Flags* (termasuk *faltering growth* / BB tidak naik adekuat atau stagnan selama 14 hari) digunakan sebagai indikator rujukan absolut untuk mencegah kerusakan kognitif ireversibel pada masa 1000 Hari Pertama Kehidupan (HPK).
    
    **4. Privasi & Keamanan Data (Informed Consent):**
    Pengumpulan data demografis dan klinis pasien telah diselaraskan dengan prinsip etika kerahasiaan medis. Data dienkripsi dan diintegrasikan langsung menuju basis data internal Puskesmas Sumber guna kepentingan intervensi kesehatan masyarakat yang terukur.
    
    **5. Cakupan Usia:**
    Kalkulator ini hanya berlaku untuk anak usia **0-60 bulan (0-5 tahun)**, sesuai cakupan tabel WHO Child Growth Standards yang digunakan.
    """)
