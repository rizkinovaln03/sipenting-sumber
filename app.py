import streamlit as st
import requests
import sqlite3
import os
import io
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import plotly.graph_objects as go
import pandas as pd
import numpy as np
from decimal import Decimal, InvalidOperation
from pygrowup import Calculator
from pygrowup import exceptions as pg_exceptions

# --- 0. KONFIGURASI DATABASE (SQLite lokal, di folder yang sama dengan app.py) ---
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sipenting.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS pasien (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            no_rm TEXT,
            nama TEXT NOT NULL,
            nama_ibu TEXT,
            alamat TEXT,
            tanggal_lahir TEXT NOT NULL,
            jenis_kelamin TEXT NOT NULL,
            dibuat_pada TEXT
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS pengukuran (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pasien_id INTEGER,
            tanggal_ukur TEXT,
            usia_bulan REAL,
            bb REAL,
            tb REAL,
            waz REAL,
            haz REAL,
            whz REAL,
            status_haz TEXT,
            tindak_lanjut TEXT,
            tb_skor INTEGER DEFAULT 0,
            is_gtm TEXT DEFAULT 'Tidak',
            red_flags INTEGER DEFAULT 0,
            FOREIGN KEY (pasien_id) REFERENCES pasien (id)
        )
    ''')
    conn.commit()
    conn.close()


def cari_pasien(keyword):
    """Cari pasien berdasarkan No. RM atau Nama (untuk cek 'sudah pernah diinput apa belum')."""
    if not keyword or keyword.strip() == "":
        return []
    conn = get_conn()
    like = f"%{keyword.strip()}%"
    rows = conn.execute(
        "SELECT * FROM pasien WHERE no_rm LIKE ? OR nama LIKE ? ORDER BY nama LIMIT 20",
        (like, like)
    ).fetchall()
    conn.close()
    return rows


def get_pasien(pasien_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM pasien WHERE id = ?", (pasien_id,)).fetchone()
    conn.close()
    return row


def get_semua_pasien():
    conn = get_conn()
    rows = conn.execute("SELECT id, nama, no_rm, tanggal_lahir FROM pasien ORDER BY nama").fetchall()
    conn.close()
    return rows


def simpan_pasien_baru(no_rm, nama, nama_ibu, alamat, tanggal_lahir, jenis_kelamin):
    conn = get_conn()
    cur = conn.execute(
        """INSERT INTO pasien (no_rm, nama, nama_ibu, alamat, tanggal_lahir, jenis_kelamin, dibuat_pada)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (no_rm, nama, nama_ibu, alamat, tanggal_lahir.isoformat(), jenis_kelamin,
         datetime.now().isoformat(timespec="seconds"))
    )
    conn.commit()
    pasien_id = cur.lastrowid
    conn.close()
    return pasien_id


def simpan_pengukuran(pasien_id, tgl_ukur, usia_bulan, bb, tb, waz, haz, whz, final_status_text, tindak_lanjut, red_flags, tb_skor, is_gtm):
    conn = get_conn()
    cur = conn.execute(
        """INSERT INTO pengukuran (pasien_id, tanggal_ukur, usia_bulan, bb, tb, waz, haz, whz, status_haz, tindak_lanjut, red_flags, tb_skor, is_gtm)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (pasien_id, tgl_ukur.isoformat(), usia_bulan, bb, tb, waz, haz, whz, final_status_text, tindak_lanjut, int(red_flags), tb_skor, is_gtm)
    )
    conn.commit()
    conn.close()


def get_riwayat_pengukuran(pasien_id):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM pengukuran WHERE pasien_id = ? ORDER BY tanggal_ukur ASC",
        (pasien_id,)
    ).fetchall()
    conn.close()
    return rows


def get_ringkasan_status_gizi():
    """Distribusi status gizi berdasarkan PENGUKURAN TERAKHIR tiap pasien (bukan semua baris riwayat)."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT p.status_haz, COUNT(*) as jumlah FROM pengukuran p
        INNER JOIN (
            SELECT pasien_id, MAX(tanggal_ukur) as tgl_terakhir
            FROM pengukuran GROUP BY pasien_id
        ) terakhir ON p.pasien_id = terakhir.pasien_id AND p.tanggal_ukur = terakhir.tgl_terakhir
        GROUP BY p.status_haz
    """).fetchall()
    conn.close()
    return {r["status_haz"]: r["jumlah"] for r in rows}


def get_semua_data_gabungan():
    """Semua pasien + seluruh riwayat pengukurannya, digabung (join) untuk keperluan export/rekap."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT
            pa.no_rm, pa.nama AS nama_anak, pa.nama_ibu, pa.alamat,
            pa.tanggal_lahir, pa.jenis_kelamin,
            pe.tanggal_ukur, pe.usia_bulan, pe.bb, pe.tb,
            pe.waz, pe.haz, pe.whz, pe.status_haz, pe.tindak_lanjut, pe.tb_skor, pe.is_gtm, pe.red_flags
        FROM pasien pa
        LEFT JOIN pengukuran pe ON pe.pasien_id = pa.id
        ORDER BY pa.nama, pe.tanggal_ukur
    """).fetchall()
    conn.close()
    return pd.DataFrame([dict(r) for r in rows])


def buat_file_excel(df_pasien, df_gabungan, df_riwayat_pasien_terpilih=None, nama_pasien_terpilih=None):
    """Bikin file Excel di memori (BytesIO) dengan beberapa sheet, siap didownload user."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_pasien.to_excel(writer, sheet_name="Daftar Pasien", index=False)
        df_gabungan.to_excel(writer, sheet_name="Rekap Semua Pengukuran", index=False)
        if df_riwayat_pasien_terpilih is not None and not df_riwayat_pasien_terpilih.empty:
            sheet_name = f"Riwayat - {nama_pasien_terpilih}"[:31]  # batas nama sheet Excel = 31 karakter
            df_riwayat_pasien_terpilih.to_excel(writer, sheet_name=sheet_name, index=False)

        # Rapikan lebar kolom biar gak mepet
        for sheet in writer.sheets.values():
            for col_cells in sheet.columns:
                panjang_maks = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
                sheet.column_dimensions[col_cells[0].column_letter].width = min(panjang_maks + 3, 40)

    buffer.seek(0)
    return buffer


def get_tren_kunjungan_mingguan():
    """Jumlah pengukuran per minggu (ISO week) dari seluruh riwayat, untuk grafik tren."""
    conn = get_conn()
    rows = conn.execute("SELECT tanggal_ukur FROM pengukuran ORDER BY tanggal_ukur").fetchall()
    conn.close()
    if not rows:
        return pd.DataFrame({"minggu": [], "jumlah": []})
    df = pd.DataFrame({"tanggal_ukur": [r["tanggal_ukur"] for r in rows]})
    df["tanggal_ukur"] = pd.to_datetime(df["tanggal_ukur"])
    df["minggu"] = df["tanggal_ukur"].dt.strftime("%Y-W%U")
    agg = df.groupby("minggu").size().reset_index(name="jumlah")
    return agg.tail(8)  # 8 minggu terakhir

init_db()

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="SIPENTING Sumber",
    page_icon="sipenting.png",
    layout="wide"
)

# --- 2. INJEKSI CUSTOM CSS ---
st.markdown("""
    <style>
    /* --- MENGHILANGKAN GARIS ATAS STREAMLIT --- */
    /* Ini untuk mematikan header atas STREAMLIT, sehingga menu burger ikut mati.
       Jadi Tampilan Benar-benar Custom */
    [data-testid="stHeader"] { visibility: hidden; }

    /* CSS Desain Bawaanmu */
    .stApp { background-color: #F4F9F9; }
    div[data-testid="metric-container"] {
        background-color: white; border: 1px solid #E0E0E0; padding: 15px;
        border-radius: 10px; box-shadow: 2px 2px 10px rgba(0,0,0,0.05);
    }
    .stButton>button { border-radius: 8px; font-weight: bold; transition: 0.3s; }
    .stButton>button:hover { transform: scale(1.02); }
    .main-title {
        text-align: left; color: #1A5276; font-weight: 800; font-size: 3rem;
        margin-bottom: -15px; text-transform: uppercase; letter-spacing: 2px;
    }
    .sub-title { text-align: left; color: #2E86C1; font-size: 1.2rem; font-style: italic; margin-bottom: 20px; }
    .kia-card {
        background-color: white; border: 1px solid #E0E0E0; border-radius: 10px;
        padding: 16px 20px; box-shadow: 2px 2px 10px rgba(0,0,0,0.05); margin-bottom: 12px;
    }
    </style>
""", unsafe_allow_html=True)

# --- 3. HEADER & LOGO ---
# Membagi 3 kolom: [ Teks Judul Kiri ] - [ Logo Puskesmas Tengah ] - [ Logo SiPENTING Kanan ]
col_kiri, col_tengah, col_kanan = st.columns([5, 1.2, 1.8])

with col_kiri:
    # 1. TEKS JUDUL (KIRI)
    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
    st.markdown("<div class='main-title'>SIPENTING</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-title'>Sistem Pencegahan & Edukasi Stunting Terintegrasi | Puskesmas Sumber</div>", unsafe_allow_html=True)

with col_tengah:
    # 2. LOGO PUSKESMAS (TENGAH)
    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
    st.markdown(
        '<img src="logo.png" style="width:100%; object-fit:contain;">',
        unsafe_allow_html=True
    )

with col_kanan:
    # 3. LOGO SIPENTING (KANAN)
    st.markdown(
        '<img src="sipenting.png" style="width:100%; object-fit:contain;">', 
        unsafe_allow_html=True
    )

st.markdown("---")

# --- 4. KALKULATOR Z-SCORE RESMI WHO (LMS, via pygrowup) ---
@st.cache_resource
def get_calculator():
    # Menginisialisasi kalkulator pygrowup dengan standar WHO
    return Calculator(adjust_height_data=False, adjust_weight_scores=False, include_cdc=False)


def hitung_zscore(bb, tb, usia_bulan, jk):
    calc = get_calculator()
    sex_code = "M" if jk == "Laki-laki" else "F"
    errors = []

    # Fungsi pembantu untuk menangkap error kalkulator
    def _safe_z(indicator, measurement, height=None):
        try:
            z = calc.zscore_for_measurement(indicator, measurement, usia_bulan, sex_code, height=height)
            return round(float(z), 2)
        except (pg_exceptions.InvalidMeasurement, pg_exceptions.DataNotFound, 
                InvalidOperation, AssertionError, TypeError, ValueError) as e:
            errors.append(f"{indicator.upper()}: {e}")
            return None

    # Hitung 3 indikator utama
    waz = _safe_z("wfa", bb)
    haz = _safe_z("lhfa", tb)
    
    # WHZ butuh perlakuan khusus Panjang/Tinggi badan (LMS Table WHO berbeda)
    # pygrowup menangani Panjang (<24 bln) vs Tinggi (>=24 bln) secara otomatis 
    # berdasarkan usia_bulan jika data tinggi badan disediakan.
    whz = _safe_z("wfl", bb, height=tb) if usia_bulan < 24 else _safe_z("wfh", bb, height=tb)

    return waz, haz, whz, errors


# Mengambil data kurva standar haz WHO (CACHED)
@st.cache_data
def get_kurva_haz(jk):
    calc = get_calculator()
    table = calc.lhfa_boys_0_5 if jk == "Laki-laki" else calc.lhfa_girls_0_5
    bulan, median, sd2neg, sd3neg = [], [], [], []
    for k, v in table.items():
        if k == "field_name": continue
        bulan.append(float(v["Month"]))
        median.append(float(v["SD0"]))
        sd2neg.append(float(v["SD2neg"]))
        sd3neg.append(float(v["SD3neg"]))
    df = pd.DataFrame({"bulan": bulan, "median": median, "sd2neg": sd2neg, "sd3neg": sd3neg}).sort_values("bulan")
    return df


# --- 5. SISTEM TABS MULTI-HALAMAN ---
tab1, tab2 = st.tabs(["🧮 Kalkulator & Skrining", "📖 Dashboard & Rekap Data"])

# ==========================================
# TAB 1: KALKULATOR & KURVA PERTUMBUHAN + REGISTRASI/SEARCH PASIEN
# ==========================================
with tab1:
    if "pasien_id_terpilih" not in st.session_state:
        st.session_state.pasien_id_terpilih = None
    if "sudah_dihitung" not in st.session_state:
        st.session_state.sudah_dihitung = False

    # Membagi 2 kolom Utama di Tab 1: [ Form Data Kiri ] - [ Hasil metric & Kurva Kanan ]
    col_form, col_hasil = st.columns([1.2, 2.5])

    # ---- KOLOM KIRI: FORM PENDAFTARAN & PENGUKURAN ----
    with col_form:
        st.header("🔍 Cari/Pilih Pasien")
        kata_kunci = st.text_input("No. RM atau Nama Pasien", key="kata_kunci_cari")
        
        # Panggil pencarian secara real-time
        hasil_cari = cari_pasien(kata_kunci) if kata_kunci else []

        if hasil_cari:
            # Buat opsi dropdown untuk pasien ditemukan
            opsi = {f"{r['nama']} (RM: {r['no_rm'] or '-'})": r["id"] for r in hasil_cari}
            pilihan = st.selectbox("Pasien ditemukan, pilih salah satu:", list(opsi.keys()))
            if st.button("✅ Gunakan Data Pasien Ini"):
                st.session_state.pasien_id_terpilih = opsi[pilihan]
                st.session_state.sudah_dihitung = False # Reset hasil hitung
                st.rerun() # Refresh agar state langsung terupdate
        elif kata_kunci:
            st.info("Pasien tidak ditemukan. Akan didaftarkan sebagai pasien baru di bawah.")

        # Tombol untuk mendaftarkan pasien baru
        if st.session_state.pasien_id_terpilih:
            if st.button("➕ Daftar Pasien Baru (Reset)"):
                st.session_state.pasien_id_terpilih = None
                st.session_state.sudah_dihitung = False
                st.rerun()

        st.markdown("---")
        st.header("📋 Form Data Pasien & Pengukuran")

        # Ambil data pasien lama jika ada yang terpilih
        pasien_lama = get_pasien(st.session_state.pasien_id_terpilih) if st.session_state.pasien_id_terpilih else None

        # Data Identitas (Jika pasien lama, input didisable)
        if pasien_lama:
            riwayat_lama = get_riwayat_pengukuran(pasien_lama["id"])
            st.success(f"📖 Pasien terdaftar: **{pasien_lama['nama']}** — {len(riwayat_lama)}x pengukuran"
                       + (f", terakhir {riwayat_lama[-1]['tanggal_ukur']}" if riwayat_lama else ""))
            
            # Non-aktifkan input identitas karena sudah ada
            no_rm = st.text_input("No. RM Puskesmas (opsional)", value=pasien_lama["no_rm"] or "", disabled=True)
            nama_anak = st.text_input("Nama Anak", value=pasien_lama["nama"], disabled=True)
            nama_ibu = st.text_input("Nama Ibu", value=pasien_lama["nama_ibu"] or "", disabled=True)
            alamat = st.text_area("Alamat / RT RW", value=pasien_lama["alamat"] or "", disabled=True)
            
            # Konversi tanggal lahir lama untuk date_input
            tgl_lahir_default = datetime.strptime(pasien_lama["tanggal_lahir"], "%Y-%m-%d").date()
            jk = pasien_lama["jenis_kelamin"]
            
            # Tampilkan Jenis Kelamin lama (disabled)
            st.text_input("Jenis Kelamin", value=jk, disabled=True)

        else:
            # Pasien baru, input aktif
            st.info("Mendaftarkan pasien baru.")
            no_rm = st.text_input("No. RM Puskesmas (opsional)")
            nama_anak = st.text_input("Nama Anak")
            nama_ibu = st.text_input("Nama Ibu")
            alamat = st.text_area("Alamat / RT RW")
            tgl_lahir_default = date(2023, 1, 1)
            jk = None

        st.markdown("---")
        # Data Pengukuran (Selalu Aktif)
        tgl_ukur = st.date_input("Tanggal Pengukuran (Hari Ini)", value=datetime.today())
        
        # Tanggal Lahir (Disable jika pasien lama)
        tgl_lahir = st.date_input("Tanggal Lahir Anak", value=tgl_lahir_default, 
                                 min_value=datetime(2020, 1, 1), max_value=datetime.today(),
                                 disabled=bool(pasien_lama))

        # Hitung usia presisi
        selisih = relativedelta(tgl_ukur, tgl_lahir)
        usia_bulan = selisih.years * 12 + selisih.months + (selisih.days / 30.44) # Usia rata-rata bulan
        st.info(f"Usia Anak saat pengukuran: **{selisih.years} Tahun, {selisih.months} Bulan**")
        
        # Peringatan cakupan usia WHO 0-5 Tahun (60 bulan)
        if usia_bulan > 60:
            st.warning("⚠️ Standar WHO Child Growth Standards (pygrowup) di aplikasi ini hanya mencakup usia 0-60 bulan (0-5 Tahun). Kalkulasi mungkin kurang akurat.")

        # Jenis kelamin (selectbox jika pasien baru)
        if not pasien_lama:
            jk = st.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"])

        bb = st.number_input("Berat Badan (kg)", min_value=1.0, step=0.1)
        tb = st.number_input("Tinggi/Panjang Badan (cm)", min_value=30.0, step=0.1)
        catatan_kunjungan = st.text_area("Catatan kunjungan (opsional)")

        st.markdown("---")
        # --- FIX: TB SKORING MENYESUAIKAN GAMBAR DOKTER ---
        # Menggunakan selectbox biar bisa milih skornya (0, 1, 2, atau 3)
        st.header("🦠 Skoring TB IDAI (Standard Nasional)")
        st.caption("Pilih level gejala sesuai pemeriksaan untuk menghitung skor otomatis (Total ≥ 6 = Diagnosis TB Tegak).")

        # Parameter 1: Kontak
        opt_kontak = {
            "0: Tidak jelas": 0,
            "1: Laporan keluarga, BTA (-)": 1,
            "3: BTA (+)": 3
        }
        tb_contacts_val = st.selectbox("1. Kontak dengan Pasien TB", list(opt_kontak.keys()))
        score_contacts = opt_kontak[tb_contacts_val]

        # Parameter 2: Uji Tuberkulin
        opt_tuberculin = {
            "0: Negatif": 0,
            "3: Positif (≥10mm / ≥5mm imunokompromais)": 3
        }
        tb_tuberculin_val = st.selectbox("2. Uji Tuberkulin (Mantoux)", list(opt_tuberculin.keys()))
        score_tuberculin = opt_tuberculin[tb_tuberculin_val]

        # Parameter 3: Status Gizi (image uses BB/TB or BB/U)
        opt_nutrition = {
            "0: Normal": 0,
            "2: Gizi kurang (BB/TB < 90% atau BB/U < 80%)": 2,
            "3: Gizi buruk (BB/TB < 70% atau BB/U < 60%)": 3
        }
        tb_nutrition_val = st.selectbox("3. Status Gizi Anak", list(opt_nutrition.keys()))
        score_nutrition = opt_nutrition[tb_nutrition_val]

        # Parameter 4: Gejala Klinis
        opt_clinics = {
            "0: Tidak Ada Gejala": 0,
            "1: Ada Gejala (Demam ≥ 2 mgg DAN/ATAU Batuk ≥ 3 mgg)": 1
        }
        tb_clinics_val = st.selectbox("4. Gejala Klinis", list(opt_clinics.keys()))
        score_clinics = opt_clinics[tb_clinics_val]

        # Parameter 5: Pembesaran Kelenjar Limfe
        opt_lymph = {
            "0: Tidak Ada": 0,
            "1: Ada (≥ 1 cm, > 1 KGB, tidak nyeri)": 1
        }
        tb_lymph_val = st.selectbox("5. Pembesaran Kelenjar Limfe", list(opt_lymph.keys()))
        score_lymph = opt_lymph[tb_lymph_val]

        # Parameter 6: Pembengkakan Tulang/Sendi
        opt_bones = {
            "0: Tidak Ada": 0,
            "1: Ada": 1
        }
        tb_bones_val = st.selectbox("6. Pembengkakan Tulang/Sendi", list(opt_bones.keys()))
        score_bones = opt_bones[tb_bones_val]

        # Parameter 7: Foto Toraks
        opt_xray = {
            "0: Normal / Kelainan tidak jelas": 0,
            "3: Gambaran Sugestif TB": 3
        }
        tb_xray_val = st.selectbox("7. Foto Rontgen Toraks", list(opt_xray.keys()))
        score_xray = opt_xray[tb_xray_val]

        # --- FIX: GTM SCREENING (PINDAH KE FORM UTAMA) ---
        # Ini penting agar Dokter/Kader langsung skrining inappropriate feeding
        st.header("🍽️ Skrining GTM (Kesulitan Makan)")
        is_gtm = st.radio("Apakah anak mengalami Gerakan Tutup Mulut (GTM)/Susah Makan?", ("Tidak", "Ya"))

        st.markdown("---")
        # --- FIX: RED FLAGS (PINDAH KE FORM UTAMA) ---
        # Skrining kelainan bawaan/BB Stagnan (Holliday-Segar fluid calculation trigger)
        st.header("🚨 Tanda Bahaya (Red Flags)")
        red_flags_aktif = st.checkbox("Ada Tanda Bahaya (Kelainan Bawaan/Faltering Growth/BB Stagnan)?")

        st.markdown("---")
        # Tombol Hitung
        # use_container_width=True biar tombol penuhi lebar kolom form
        if st.button("🧮 Hitung & Analisis Gizi", type="primary", use_container_width=True):
            if not pasien_lama and (nama_anak == "" or nama_ibu == ""):
                st.error("⚠️ Nama Anak dan Ibu wajib diisi untuk pasien baru!")
            elif usia_bulan > 60:
                st.error("⚠️ pygrowup tidak dapat menghitung Z-score untuk usia di atas 60 bulan.")
            elif usia_bulan < 0:
                st.error("⚠️ Tanggal Pengukuran tidak boleh sebelum Tanggal Lahir.")
            else:
                # Panggil fungsi kalkulasi Z-score
                waz, haz, whz, errors = hitung_zscore(bb, tb, usia_bulan, jk)
                
                if errors:
                    st.error("Gagal menghitung Z-Score. Periksa kembali input data.")
                    for err in errors:
                        st.caption(f"- {err}")
                else:
                    # Simpan hasil hitung ke session state
                    st.session_state.waz, st.session_state.haz, st.session_state.whz = waz, haz, whz
                    
                    # --- FIX: TB SCORING TOTAL CALCULATION ---
                    # Hitung total skor berdasarkan pilihan dropdown
                    skor_tb_total = score_contacts + score_tuberculin + score_nutrition + score_clinics + score_lymph + score_bones + score_xray
                    st.session_state.skor_tb_total = skor_tb_total
                    
                    st.session_state.is_gtm = is_gtm
                    st.session_state.red_flags = red_flags_aktif
                    st.session_state.catatan_kunjungan = catatan_kunjungan
                    st.session_state.tgl_ukur = tgl_ukur
                    st.session_state.jk = jk
                    st.session_state.no_rm_input = no_rm
                    st.session_state.nama_anak_input = nama_anak
                    st.session_state.nama_ibu_input = nama_ibu
                    st.session_state.alamat_input = alamat
                    st.session_state.tgl_lahir_input = tgl_lahir
                    st.session_state.usia_bulan = usia_bulan
                    st.session_state.bb_input = bb
                    st.session_state.tb_input = tb
                    
                    st.session_state.sudah_dihitung = True
                    st.success("Perhitungan selesai! Cek hasil di kolom sebelah kanan.")
                    # st.rerun() # Tidak perlu rerun karena metrics di kolom kanan langsung render


    # ---- KOLOM KANAN: HASIL METRICS, INTERVENSI, & KURVA ----
    with col_hasil:
        st.header("📊 Hasil Analisis Gizi WHO")
        
        if st.session_state.sudah_dihitung:
            # Ambil data dari session state
            waz = st.session_state.waz
            haz = st.session_state.haz
            whz = st.session_state.whz
            # --- FIX: TB SKORING TOTAL (AMBIL DARI SESSION) ---
            skor_tb_total = st.session_state.skor_tb_total
            
            is_gtm = st.session_state.is_gtm
            red_flags_aktif = st.session_state.red_flags
            
            # --- FIX: INDIKATOR TINGGI (HAZ) DIPERTAHANKAN ---
            # Menentukan status TB/U (HAZ) berdasarkan Kemenkes
            status_haz = ""
            if haz < -3: status_haz = "Severely Stunted (Sangat Pendek)"
            elif haz < -2: status_haz = "Stunted (Pendek)"
            elif haz <= 2: status_haz = "Normal"
            else: status_haz = "Tinggi"

            # Menentukan status BB/TB (WHZ) berdasarkan Kemenkes (Opsional buat intervensi)
            status_whz = ""
            if whz < -3: status_whz = "Gizi Buruk"
            elif whz < -2: status_whz = "Gizi Kurang"
            elif whz <= 1: status_whz = "Gizi Baik (Normal)"
            elif whz <= 2: status_whz = "Berisiko Gizi Lebih"
            elif whz <= 3: status_whz = "Gizi Lebih"
            else: status_whz = "Obesitas"

            # Tampilkan 3 metrik utama secara berdampingan
            col_metric1, col_metric2, col_metric3 = st.columns(3)
            col_metric1.metric("Berat / Umur (WAZ)", f"{waz} SD", help="WHO Weight-for-Age Z-score")
            
            # HAZ Metric (Indikator Stunting Utama)
            # help="" biar cursor hover tampilkan info
            col_metric2.metric("Tinggi / Umur (HAZ)", f"{haz} SD", help="WHO Height-for-Age Z-score (Indikator Stunting)",
                               delta=status_haz, delta_color="off" if haz >= -2 else "normal") # delta_color off biar gak merah kalau normal
            
            col_metric3.metric("Berat / Tinggi (WHZ)", f"{whz} SD", help="WHO Weight-for-Height Z-score")

            st.markdown("---")

            # --- FIX: LOGIKA INTERVENSI BARU & SKORING TB SESUAI GAMBAR ---
            st.header("💡 Analisis Lanjutan & Intervensi (Rekomendasi IDAI)")
            
            tindak_lanjut = ""
            # Gunakan st.error/st.warning untuk visualisasi prioritas
            
            # --- SKORING TB INTEGRATED (IDAI Pedoman) ---
            st.subheader(f"🦠 Hasil Skoring TB (Pedoman Nasional): Skor {skor_tb_total}")
            
            if skor_tb_total >= 6:
                st.error(f"**SKOR TB = {skor_tb_total} (Total ≥ 6). Diagnosa TB Klinis Tegak!**\n\n- WAJIB Mulai Pengobatan OAT (Obat Anti-TB) dosis anak.\n\n- Segera konsultasi ke Dokter Spesialis Anak/Poli DOTS Puskesmas.")
                st.subheader("❗ Tindakan Segera")
                tindak_lanjut += " RUJUKAN TB: Diagnosa TB Klinis (Skor >= 6)."
            elif skor_tb_total > 0:
                st.warning(f"Skor TB = {skor_tb_total} (Total < 6). Belum diagnosa tegak, perlu observasi gejala menetap.")
            else:
                st.success("Skor TB = 0. Tidak ada indikasi TB.")

            st.markdown("---")
            # --- FIX: LOGIKA TINGGI BADAN (HAZ) DIUTAMAKAN ---
            if red_flags_aktif:
                st.error("**🚨 ALERT: Red Flags AKTIF!** Terdapat Tanda Bahaya (Kelainan Bawaan/BB Stagnan).")
                st.subheader("❗ Tindakan Segera")
                st.info("Prioritas intervensi cairan Holliday-Segar (Puskesmas Sumber Protocol). Evaluasi Red Flags ketat.")
                tindak_lanjut += " Rujuk Spesialis Anak/IGD (🚨 Red Flags AKTIF!)."
                
                # Jika Haz bermasalah, tambah info
                if haz < -2:
                    st.subheader(f"Tren Tinggi (HAZ = {haz} SD)")
                    st.caption(f"Status: {status_haz}. Red Flags menyulitkan pertumbuhan linier.")

            elif whz < -2:
                # Kasus Gizi Kurang/Buruk (Wasting) didahulukan
                st.error(f"**GIZI BURUK/KURANG (WASTING)!** WHZ = {whz} SD ({status_whz}).")
                st.subheader("💡 Intervensi Nutrisi (Tatalaksana Gizi Buruk/Kurang Kemenkes)")
                st.info("Pemberian F100/PMT Pemulihan terukur. Skrining penyakit penyerta ketat (Skoring TB IDAI).")
                
                if haz < -2:
                    st.subheader(f"Tinggi (HAZ = {haz} SD): {status_haz}")
                    st.caption("Pertumbuhan Tinggi linier terganggu karena Wasting kronis.")

            elif haz < -2:
                # Kasus Stunted (Pendek) tanpa Wasting (Wasting -), GTM opsional
                st.warning(f"**STUNTED/PENDEK (HAZ = {haz} SD)!** Status: {status_haz}.")
                st.subheader("💡 Intervensi Pertumbuhan Linier")
                
                if is_gtm == "Ya":
                    st.error("Anak mengalami GTM (Kombinasi Stunted + Susah Makan).")
                    st.info("Pemberian Protein Hewani terukur (Telur, Ayam, Daging) di setiap porsi MPASI. **Wajib terapkan Feeding Rules IDAI**.")
                else:
                    st.info("Pemberian Protein Hewani terukur (Telur, Ayam, Daging) di setiap porsi MPASI. Skrining penyakit penyerta.")
                    
            else:
                # Gizi & Tinggi Normal
                st.success("**Pertumbuhan Gizi & Tinggi Normal.**")
                if is_gtm == "Ya":
                    st.warning("Gizi Normal, tapi ada GTM (Gerakan Tutup Mulut).")
                    st.info("Anjurkan Feeding Rules IDAI untuk mencegah gizi kurang di kemudian hari.")
                else:
                    st.info("Pertahankan pola makan menu lengkap gizi seimbang. Pemantauan rutin di Posyandu.")

            # Simpan ringkasan tindak lanjut
            if red_flags_aktif:
                final_status_text = status_haz + " / Rujuk Segera!"
                # tindak lanjut sudah diisi di blok red flags
            elif whz < -2 or haz < -2:
                final_status_text = status_haz
                if tindak_lanjut == "": tindak_lanjut = "Evaluasi Pola Asuh & Diet Protein Hewani (Telur/Ikan) terukur." # Default tgl-lanjut stunting
            else:
                final_status_text = "Gizi & Tinggi Normal."
                if is_gtm == "Ya": tindak_lanjut = "Feeding Rules IDAI (Masalah Makan)."
                else: tindak_lanjut = "Pemantauan rutin Posyandu (Pertahankan)."

            st.session_state.final_status_text = final_status_text
            st.session_state.tindak_lanjut_text = tindak_lanjut

            st.markdown("---")
            # --- KURVA PERTUMBUHAN LINIER (HAZ) DITAMPILKAN ---
            st.subheader(f"📈 Kurva Tren Pertumbuhan WHO (HAZ Linier — {st.session_state.jk})")
            
            with st.spinner("Membuat kurva..."):
                df_kurva = get_kurva_haz(st.session_state.jk)
                
                fig = go.Figure()

                # Garis Standar WHO (CACHED)
                fig.add_trace(go.Scatter(x=df_kurva["bulan"], y=df_kurva["median"], mode='lines', name='Median (WHO)', line=dict(color='green', width=2)))
                fig.add_trace(go.Scatter(x=df_kurva["bulan"], y=df_kurva["sd2neg"], mode='lines', name='-2 SD (Stunted)', line=dict(color='orange', width=2, dash='dash')))
                fig.add_trace(go.Scatter(x=df_kurva["bulan"], y=df_kurva["sd3neg"], mode='lines', name='-3 SD (Sev. Stunted)', line=dict(color='red', width=2, dash='dot')))

                # Data Pasien (Titik Tunggal)
                fig.add_trace(go.Scatter(x=[st.session_state.usia_bulan], y=[st.session_state.haz], 
                                         mode='markers', name='Kunjungan Hari Ini', 
                                         marker=dict(color='blue', size=14, symbol='star')))

                fig.update_layout(xaxis_title="Usia (Bulan)", yaxis_title="Z-Score TB/U (HAZ)", height=450)
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")
            # --- TOMBOL SIMPAN KE DATABASE ---
            if st.button("💾 Simpan Data ke Buku KIA Digital", type="primary"):
                # Menentukan pasien_id
                if pasien_lama:
                    pasien_id = pasien_lama["id"]
                else:
                    # Registrasi pasien baru dulu
                    pasien_id = simpan_pasien_baru(
                        st.session_state.no_rm_input or None,
                        st.session_state.nama_anak_input,
                        st.session_state.nama_ibu_input or None,
                        st.session_state.alamat_input or None,
                        st.session_state.tgl_lahir_input,
                        st.session_state.jk
                    )
                
                # --- FIX: SIMPAN TB SKOR TOTAL & GTM & RED FLAGS ---
                # Simpan data pengukuran
                simpan_pengukuran(
                    pasien_id, 
                    st.session_state.tgl_ukur, 
                    st.session_state.usia_bulan, 
                    st.session_state.bb_input, 
                    st.session_state.tb_input, 
                    st.session_state.waz, 
                    st.session_state.haz, 
                    st.session_state.whz, 
                    st.session_state.final_status_text, 
                    st.session_state.tindak_lanjut_text,
                    st.session_state.red_flags,
                    st.session_state.skor_tb_total, # Skor TB Total IDAI
                    st.session_state.is_gtm
                )
                
                st.success("🎉 Data Registrasi & Pengukuran BERHASIL disimpan ke database! Cek di Tab Buku KIA Digital.")
                
                # Reset state kalkulator
                st.session_state.sudah_dihitung = False
                st.session_state.pasien_id_terpilih = pasien_id # Tetap pilih pasien ini setelah simpan
                st.rerun()

        else:
            # Pesan default jika belum klik Hitung
            st.info("👈 Silakan isi form pendaftaran/pengukuran di sebelah kiri, lalu klik 'Hitung & Analisis Gizi' untuk menampilkan hasil Z-Score & Kurva.")
            st.markdown(
                """
                <div style="background-color: white; border: 1px solid #E0E0E0; border-radius: 10px; padding: 20px; box-shadow: 2px 2px 10px rgba(0,0,0,0.05); text-align: center; margin-top: 50px;">
                    <img src="sipenting.png" style="width:150px; object-fit:contain; margin-bottom: 20px;">
                    <h3 style="color: #1A5276;">Selamat Datang di SIPENTING</h3>
                    <p style="color: #2E86C1;"><i>Sistem Pencegahan & Edukasi Stunting Terintegrasi | Puskesmas Sumber</i></p>
                    <p>pygrowup mengimplementasikan WHO Child Growth Standards resmi.</p>
                </div>
                """, unsafe_allow_html=True
            )


# ==========================================
# TAB 2: DASHBOARD & REKAP DATA DIGITAL
# ==========================================
with tab2:
    st.subheader("📖 Buku KIA Digital — Riwayat Kunjungan Pasien")

    semua_pasien = get_semua_pasien()
    if not semua_pasien:
        st.info("Belum ada pasien terdaftar. Simpan data pasien pertama lewat tab 🧮 Kalkulator & Skrining.")
    else:
        # Pili Pasien dari List
        opsi_monitor = {f"{p['nama']} (RM: {p['no_rm'] or '-'})": p["id"] for p in semua_pasien}
        pilih_monitor = st.selectbox("Pilih pasien untuk monitoring:", list(opsi_monitor.keys()))
        
        pid_monitor = opsi_monitor[pilih_monitor]
        detail = get_pasien(pid_monitor)
        riwayat = get_riwayat_pengukuran(pid_monitor)

        # Kartu Identitas Pasien
        st.markdown(f"""
        <div class="kia-card">
        <b>Nama Anak:</b> {detail['nama']} &nbsp;|&nbsp; <b>No. RM:</b> {detail['no_rm'] or '-'}<br>
        <b>Nama Ibu:</b> {detail['nama_ibu'] or '-'} &nbsp;|&nbsp; <b>JK:</b> {detail['jenis_kelamin']}<br>
        <b>Tanggal Lahir:</b> {detail['tanggal_lahir']} &nbsp;|&nbsp; <b>Alamat:</b> {detail['alamat'] or '-'}
        </div>
        """, unsafe_allow_html=True)

        if not riwayat:
            st.warning("Pasien ini belum memiliki riwayat pengukuran.")
        else:
            # Tampilkan Tabel Riwayat (join query di sqlite)
            df_riwayat = pd.DataFrame([{
                "Tanggal": r["tanggal_ukur"],
                "Usia(bln)": round(r["usia_bulan"], 1),
                "BB(kg)": r["bb"],
                "TB(cm)": r["tb"],
                "Z-Score HAZ (Tinggi)": r["haz"],
                "Status Tinggi (HAZ)": r["status_haz"],
                # --- FIX: TB SKOR IDAI DI TABEL RIWAYAT ---
                "Skor TB IDAI": r["tb_skor"], 
                "GTM?": r["is_gtm"],
                "Tindak Lanjut": r["tindak_lanjut"]
            } for r in riwayat])
            
            st.markdown("**📋 Tabel Riwayat Kunjungan**")
            st.dataframe(df_riwayat, use_container_width=True)

            # Export Excel Pasien Ini
            df_pasien_ini = pd.DataFrame([dict(detail)])
            df_gabungan = get_semua_data_gabungan()
            excel_pasien = buat_file_excel(df_pasien_ini, df_gabungan, df_riwayat, detail["nama"])
            st.download_button(
                "📥 Export Riwayat Pasien Ini ke Excel",
                data=excel_pasien,
                file_name=f"riwayat_{detail['nama'].replace(' ', '_')}_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    st.markdown("---")
    st.subheader("💾 Export Rekapitulasi Data Lengkap FKTP")
    df_semua_pasien = pd.DataFrame([dict(p) for p in semua_pasien]) if semua_pasien else pd.DataFrame()
    df_gabungan_semua = get_semua_data_gabungan()
    
    if df_gabungan_semua.empty:
        st.info("Belum ada data pengukuran FKTP untuk diexport.")
    else:
        excel_semua = buat_file_excel(df_semua_pasien, df_gabungan_semua)
        st.download_button(
            "📥 Download Rekap Seluruh Pasien & Pengukuran FKTP (.xlsx)",
            data=excel_semua,
            file_name=f"sipenting_FKTP_Sumber_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        st.caption("Isinya: sheet Daftar Pasien + sheet Rekap Semua Pengukuran + sheet riwayat pasien terpilih (kalau ada). Cocok buat laporan Dinas Kesehatan.")

    st.markdown("---")
    st.subheader("📊 Statistik Pertumbuhan Populasi Puskesmas Sumber")
    
    col_dash1, col_dash2 = st.columns(2)

    with col_dash1:
        st.markdown("**Distribusi Status Tinggi Anak (HAZ — Pengukuran Terakhir)**")
        ringkasan = get_ringkasan_status_gizi()
        if ringkasan:
            # Buat chart pie sederhana
            fig_pie = go.Figure(data=[go.Pie(
                labels=list(ringkasan.keys()),
                values=list(ringkasan.values()),
                hole=.3,
                marker_colors=['#2E86C1', '#F1C40F', '#E74C3C', '#28B463']
            )])
            fig_pie.update_layout(height=350, margin=dict(t=0, b=0, l=0, r=0))
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Belum ada data pengukuran.")

    with col_dash2:
        st.markdown("**Tren Jumlah Skrining Bulanan**")
        df_tren = get_tren_kunjungan_mingguan()
        if not df_tren.empty:
            # Buat chart bar tren kunjungan
            fig_bar = go.Figure(data=[go.Bar(
                x=df_tren["minggu"],
                y=df_tren["jumlah"],
                marker_color='#1A5276'
            )])
            fig_bar.update_layout(height=350, margin=dict(t=0, b=0, l=0, r=0))
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Belum ada data kunjungan.")
