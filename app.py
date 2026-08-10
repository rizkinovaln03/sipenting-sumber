import streamlit as st
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="SIPENTING - Puskesmas Sumber", layout="wide")

st.markdown("<h1 style='text-align: center; color: #2E86C1;'>SIPENTING</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center;'>Sistem Pencegahan & Edukasi Stunting Terintegrasi</h3>", unsafe_allow_html=True)
st.markdown("---")

# --- 2. SIDEBAR: INPUT DATA ---
st.sidebar.header("📋 Form Data Pasien")
nama_anak = st.sidebar.text_input("Nama Anak")
nama_ibu = st.sidebar.text_input("Nama Ibu")
alamat = st.sidebar.text_area("Alamat / RT RW")

st.sidebar.markdown("---")
tgl_ukur = st.sidebar.date_input("Tanggal Pengukuran", value=datetime.today())
tgl_lahir = st.sidebar.date_input("Tanggal Lahir", min_value=datetime(2020, 1, 1), max_value=datetime.today())

# Hitung Usia Presisi
selisih = relativedelta(tgl_ukur, tgl_lahir)
usia_bulan = selisih.years * 12 + selisih.months + (selisih.days / 30.44)
st.sidebar.info(f"Usia: {selisih.years} Tahun, {selisih.months} Bulan")

jk = st.sidebar.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"])
bb = st.sidebar.number_input("Berat Badan (kg)", min_value=1.0, step=0.1)
tb = st.sidebar.number_input("Tinggi/Panjang Badan (cm)", min_value=30.0, step=0.1)

# Fitur Red Flags
st.sidebar.markdown("---")
red_flags_aktif = st.sidebar.checkbox("🚨 Tanda Bahaya (Kelainan Bawaan / BB stagnan 14 hari)")

# --- 3. FUNGSI PERHITUNGAN (Mock WHO LMS) ---
def hitung_zscore(bb, tb, usia):
    waz = (bb - 10.0) / 1.5 if usia > 12 else (bb - 6.0) / 1.2
    haz = (tb - 75.0) / 3.0 if usia > 12 else (tb - 60.0) / 2.5
    whz = (bb - 10.0) / 1.2
    return round(waz, 2), round(haz, 2), round(whz, 2)

# --- 4. PENYIMPANAN STATE SEMENTARA ---
if 'sudah_dihitung' not in st.session_state:
    st.session_state.sudah_dihitung = False

# --- 5. LOGIKA UTAMA (HITUNG & EDUKASI) ---
if st.sidebar.button("🧮 Hitung Analisis Gizi", type="primary"):
    if nama_anak == "" or nama_ibu == "":
        st.sidebar.error("⚠️ Nama Anak dan Ibu wajib diisi!")
    else:
        st.session_state.waz, st.session_state.haz, st.session_state.whz = hitung_zscore(bb, tb, usia_bulan)
        st.session_state.sudah_dihitung = True

# TAMPILAN HASIL JIKA SUDAH DIHITUNG
if st.session_state.sudah_dihitung:
    st.success(f"✅ Analisis Gizi untuk **{nama_anak}** berhasil dilakukan.")
    
    haz = st.session_state.haz
    status_haz = "Normal" if haz >= -2 else ("Severely Stunted" if haz < -3 else "Stunted")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("BB/U (WAZ)", f"{st.session_state.waz} SD")
    col2.metric("TB/U (HAZ)", f"{haz} SD", status_haz, delta_color="off" if haz < -2 else "normal")
    col3.metric("BB/TB (WHZ)", f"{st.session_state.whz} SD")
    
    st.markdown("---")
    
    # KOTAK EDUKASI MENARIK
    st.subheader("💡 Panduan Intervensi & Edukasi Posyandu")
    
    if red_flags_aktif or haz < -3:
        st.error("🚨 **INDIKASI RUJUK (MERAH)** \n\nSegera rujuk pasien ke Poli MTBS/Sp.A Puskesmas Sumber. Kondisi memerlukan tatalaksana klinis lanjutan.")
    elif haz < -2:
        st.warning("🏥 **INTERVENSI GIZI (KUNING)** \n\nPasien memerlukan PMT (Pemberian Makanan Tambahan) Pemulihan dan wajib ditimbang kembali dalam 14 hari.")
    else:
        st.success("🏡 **PEMANTAUAN RUTIN (HIJAU)** \n\nPertumbuhan dalam batas normal. Lanjutkan edukasi gizi seimbang.")
        
    st.info("""
    **🍽️ Rekomendasi Menu Protein Hewani (Prohe):**
    - **6-11 Bulan:** Bubur lumat/cincang dengan hati ayam kampung atau ikan lele, tambahkan lemak tambahan (minyak/santan).
    - **12-24 Bulan:** Nasi keluarga, telur puyuh rebus/ayam suwir (minimal 1-2 butir telur per hari).
    - **Aturan Makan:** Waktu makan maksimal 30 menit, lingkungan netral tanpa distraksi layar/HP.
    """)
    
    st.markdown("---")
    
    # --- 6. LOGIKA PENGIRIMAN KE GOOGLE FORMS ---
    if st.button("💾 Simpan Data ke Database", type="primary"):
        # URL formResponse Google Form (bukan viewform)
        url_form = "https://docs.google.com/forms/d/e/1FAIpQLSdPfLZYTQDU_gKlbGdtyqjymO4SU9cXOs-BxcepWuiadkwofQ/formResponse"
        
        # Mapping data ke ID Entry Formulir
        form_data = {
            "entry.499068260": nama_anak,
            "entry.621064049": nama_ibu,
            "entry.12380964": alamat,
            "entry.324089303": jk,
            "entry.1082174849": round(usia_bulan, 1),
            "entry.1013822488": bb,
            "entry.1215136987": tb,
            "entry.156349914": haz,
            "entry.142945675": status_haz
        }
        
        try:
            # Kirim data diam-diam di belakang layar
            response = requests.post(url_form, data=form_data)
            
            if response.status_code == 200:
                st.success("🎉 Data berhasil di-upload ke Database Puskesmas Sumber!")
                st.session_state.sudah_dihitung = False # Reset layar
            else:
                st.error("Gagal mengirim data. Silakan coba lagi.")
        except Exception as e:
            st.error("Terjadi kesalahan koneksi saat menyimpan data.")