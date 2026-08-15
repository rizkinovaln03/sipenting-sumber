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
            jenis_kelamin TEXT NOT NULL
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
            status_gizi TEXT,
            red_flag INTEGER,
            tindak_lanjut TEXT,
            catatan TEXT,
            skor_tb INTEGER DEFAULT 0,
            is_gtm TEXT DEFAULT 'Tidak'
        )
    ''')
    
    # KODE SAKTI AUTO-MIGRASI 
    c.execute("PRAGMA table_info(pengukuran)")
    columns = [col[1] for col in c.fetchall()]
    if 'skor_tb' not in columns:
        c.execute("ALTER TABLE pengukuran ADD COLUMN skor_tb INTEGER DEFAULT 0")
    if 'is_gtm' not in columns:
        c.execute("ALTER TABLE pengukuran ADD COLUMN is_gtm TEXT DEFAULT 'Tidak'")
        
    conn.commit()
    conn.close()

def cari_pasien(keyword):
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
        """INSERT INTO pasien (no_rm, nama, nama_ibu, alamat, tanggal_lahir, jenis_kelamin)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (no_rm, nama, nama_ibu, alamat, tanggal_lahir.isoformat(), jenis_kelamin)
    )
    conn.commit()
    pasien_id = cur.lastrowid
    conn.close()
    return pasien_id

def simpan_pengukuran(pasien_id, tgl_ukur, usia_bulan, bb, tb, waz, haz, whz, status_gizi, red_flag, tindak_lanjut, catatan, skor_tb=0, is_gtm="Tidak"):
    conn = get_conn()
    c = conn.cursor()
    c.execute('''
        INSERT INTO pengukuran 
        (pasien_id, tanggal_ukur, usia_bulan, bb, tb, waz, haz, whz, status_gizi, red_flag, tindak_lanjut, catatan, skor_tb, is_gtm)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (pasien_id, tgl_ukur, usia_bulan, bb, tb, waz, haz, whz, status_gizi, int(red_flag), tindak_lanjut, catatan, skor_tb, is_gtm))
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
    conn = get_conn()
    rows = conn.execute("""
        SELECT p.status_gizi, COUNT(*) as jumlah FROM pengukuran p
        INNER JOIN (
            SELECT pasien_id, MAX(tanggal_ukur) as tgl_terakhir
            FROM pengukuran GROUP BY pasien_id
        ) terakhir ON p.pasien_id = terakhir.pasien_id AND p.tanggal_ukur = terakhir.tgl_terakhir
        GROUP BY p.status_gizi
    """).fetchall()
    conn.close()
    return {r["status_gizi"]: r["jumlah"] for r in rows}

def get_semua_data_gabungan():
    conn = get_conn()
    rows = conn.execute("""
        SELECT
            pa.no_rm, pa.nama AS nama_anak, pa.nama_ibu, pa.alamat,
            pa.tanggal_lahir, pa.jenis_kelamin,
            pe.tanggal_ukur, pe.usia_bulan, pe.bb, pe.tb,
            pe.waz, pe.haz, pe.whz, pe.status_gizi, pe.red_flag,
            pe.tindak_lanjut, pe.catatan, pe.skor_tb, pe.is_gtm
        FROM pasien pa
        LEFT JOIN pengukuran pe ON pe.pasien_id = pa.id
        ORDER BY pa.nama, pe.tanggal_ukur
    """).fetchall()
    conn.close()
    return pd.DataFrame([dict(r) for r in rows])

def buat_file_excel(df_pasien, df_gabungan, df_riwayat_pasien_terpilih=None, nama_pasien_terpilih=None):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_pasien.to_excel(writer, sheet_name="Daftar Pasien", index=False)
        df_gabungan.to_excel(writer, sheet_name="Rekap Semua Pengukuran", index=False)
        if df_riwayat_pasien_terpilih is not None and not df_riwayat_pasien_terpilih.empty:
            sheet_name = f"Riwayat - {nama_pasien_terpilih}"[:31] 
            df_riwayat_pasien_terpilih.to_excel(writer, sheet_name=sheet_name, index=False)

        for sheet in writer.sheets.values():
            for col_cells in sheet.columns:
                panjang_maks = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
                sheet.column_dimensions[col_cells[0].column_letter].width = min(panjang_maks + 3, 40)

    buffer.seek(0)
    return buffer

def get_tren_kunjungan_mingguan():
    conn = get_conn()
    rows = conn.execute("SELECT tanggal_ukur FROM pengukuran ORDER BY tanggal_ukur").fetchall()
    conn.close()
    if not rows:
        return pd.DataFrame({"minggu": [], "jumlah": []})
    df = pd.DataFrame({"tanggal_ukur": [r["tanggal_ukur"] for r in rows]})
    df["tanggal_ukur"] = pd.to_datetime(df["tanggal_ukur"])
    df["minggu"] = df["tanggal_ukur"].dt.strftime("%Y-W%U")
    agg = df.groupby("minggu").size().reset_index(name="jumlah")
    return agg.tail(8) 

def hapus_pengukuran(pengukuran_id):
    conn = get_conn()
    conn.execute("DELETE FROM pengukuran WHERE id = ?", (pengukuran_id,))
    conn.commit()
    conn.close()

def hapus_pasien_total(pasien_id):
    try:
        conn = get_conn()
        conn.execute("DELETE FROM pengukuran WHERE pasien_id = ?", (pasien_id,))
        conn.execute("DELETE FROM pasien WHERE id = ?", (pasien_id,))
        conn.commit()
        conn.close()
        return True, ""
    except Exception as e:
        return False, str(e)

init_db()

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="SIPENTING - Puskesmas Sumber", 
    page_icon="sipenting.png",
    layout="wide"
)

# --- 2. INJEKSI CUSTOM CSS ---
st.markdown("""
    <style>
    [data-testid="stHeader"] { visibility: hidden; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }

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
col_kiri, col_tengah, col_kanan = st.columns([1.2, 5, 1.8])

with col_kiri:
    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
    st.markdown(
        '<img src="https://raw.githubusercontent.com/rizkinovaln03/sipenting-sumber/main/logo.png" style="width:100%; object-fit:contain;">',
        unsafe_allow_html=True
    )

with col_tengah:
    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
    st.markdown("<div class='main-title'>SIPENTING</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-title'>Sistem Pencegahan & Edukasi Stunting Terintegrasi | Puskesmas Sumber</div>", unsafe_allow_html=True)

with col_kanan:
    st.markdown(
        '<img src="https://raw.githubusercontent.com/rizkinovaln03/sipenting-sumber/main/sipenting.png" style="width:100%; object-fit:contain;">', 
        unsafe_allow_html=True
    )

st.markdown("---")

# --- 4. KALKULATOR Z-SCORE RESMI WHO ---
@st.cache_resource
def get_calculator():
    return Calculator(adjust_height_data=False, adjust_weight_scores=False, include_cdc=False)

def hitung_zscore(bb, tb, usia_bulan, jk):
    calc = get_calculator()
    sex_code = "M" if jk == "Laki-laki" else "F"
    errors = []

    def _safe_z(indicator, measurement, height=None):
        try:
            z = calc.zscore_for_measurement(indicator, measurement, usia_bulan, sex_code, height=height)
            return round(float(z), 2)
        except (pg_exceptions.InvalidMeasurement, pg_exceptions.DataNotFound,
                InvalidOperation, AssertionError, TypeError, ValueError) as e:
            errors.append(f"{indicator.upper()}: {e}")
            return None

    waz = _safe_z("wfa", bb)
    haz = _safe_z("lhfa", tb)
    whz = _safe_z("wfl", bb, height=tb) if usia_bulan < 24 else _safe_z("wfh", bb, height=tb)
    return waz, haz, whz, errors

@st.cache_data
def get_kurva_haz(jk):
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
    df = pd.DataFrame({"bulan": bulan, "median": median, "sd2neg": sd2neg, "sd3neg": sd3neg}).sort_values("bulan")
    return df
    
@st.cache_data
def get_kurva_waz(jk):
    calc = get_calculator()
    table = calc.wfa_boys_0_5 if jk == "Laki-laki" else calc.wfa_girls_0_5
    bulan, median, sd2neg, sd3neg = [], [], [], []
    for k, v in table.items():
        if k == "field_name": continue
        bulan.append(float(v["Month"]))
        median.append(float(v["SD0"]))
        sd2neg.append(float(v["SD2neg"]))
        sd3neg.append(float(v["SD3neg"]))
    return pd.DataFrame({"bulan": bulan, "median": median, "sd2neg": sd2neg, "sd3neg": sd3neg}).sort_values("bulan")

@st.cache_data
def get_kurva_whz(jk, usia_bulan):
    calc = get_calculator()
    is_length = usia_bulan < 24
    if is_length:
        table = calc.wfl_boys_0_2 if jk == "Laki-laki" else calc.wfl_girls_0_2
        x_key = "Length"
    else:
        table = calc.wfh_boys_2_5 if jk == "Laki-laki" else calc.wfh_girls_2_5
        x_key = "Height"
        
    tinggi, median, sd2neg, sd3neg, sd2pos, sd3pos = [], [], [], [], [], []
    for k, v in table.items():
        if k == "field_name": continue
        tinggi.append(float(v[x_key]))
        median.append(float(v["SD0"]))
        sd2neg.append(float(v["SD2neg"]))
        sd3neg.append(float(v["SD3neg"]))
        sd2pos.append(float(v["SD2"]))
        sd3pos.append(float(v["SD3"]))
        
    df = pd.DataFrame({"tinggi": tinggi, "median": median, "sd2neg": sd2neg, "sd3neg": sd3neg, "sd2pos": sd2pos, "sd3pos": sd3pos})
    return df.sort_values("tinggi"), is_length

def tentukan_status_dan_tindak_lanjut(waz, haz, whz, red_flags_aktif, tgl_ukur, bb, usia_bulan, skor_tb, is_gtm):
    status_bb = "Normal" if waz is not None and waz >= -2 else ("Kurang" if waz and waz >= -3 else "Sangat Kurang")
    status_tb = "Normal" if haz is not None and haz >= -2 else ("Pendek" if haz and haz >= -3 else "Sangat Pendek")
    status_gizi = "Gizi Baik" if whz is not None and -2 <= whz <= 1 else ("Gizi Kurang" if whz and -3 <= whz < -2 else "Gizi Buruk/Lebih")

    if haz is None: return None, None, None, None
    status = "Normal" if haz >= -2 else ("Severely Stunted" if haz < -3 else "Stunted")
    
    usia_tahun = usia_bulan / 12
    if usia_bulan < 12: bb_ideal, rda = (usia_bulan + 9) / 2, 110
    elif usia_bulan <= 36: bb_ideal, rda = (2 * usia_tahun) + 8, 100
    else: bb_ideal, rda = (2 * usia_tahun) + 8, 90

    kalori = rda * bb_ideal
    teks_kalori = f"🍎 **Kebutuhan Kalori:** ± {int(kalori)} kkal/hari *(Berdasarkan RDA {rda} kkal x BB Ideal {bb_ideal:.1f} kg)*"

    if usia_bulan < 6:
        teks_usia = "- 🍼 **Nutrisi (0-6 Bulan):** ASI Eksklusif sesuka bayi (*on demand*). Tanpa tambahan air/makanan."
        teks_feeding_rules = ""
    elif usia_bulan < 12:
        teks_usia = "- 🥣 **Nutrisi (6-12 Bulan):** Mulai MPASI! Tekstur saring berlanjut ke lumat lalu cincang. Porsi 2-3 sdm hingga ½ mangkok (125ml)."
        teks_feeding_rules = "- ⏰ **Feeding Rules (IDAI):** Jadwal utama 3x, selingan 2x. Maks 30 menit/sesi. TANPA distraksi (HP/TV).\n"
    elif usia_bulan < 24:
        teks_usia = "- 🍲 **Nutrisi (1-2 Tahun):** Kenalkan Makanan Keluarga. Porsi ¾ mangkok (200ml). ASI hingga 2 tahun."
        teks_feeding_rules = "- ⏰ **Feeding Rules (IDAI):** Jadwal utama 3x, selingan 2x. Maks 30 menit/sesi. TANPA distraksi (HP/TV).\n"
    else:
        teks_usia = "- 🍛 **Nutrisi (> 2 Tahun):** Makanan Keluarga porsi utuh. Saatnya sapih ASI."
        teks_feeding_rules = "- ⏰ **Feeding Rules (IDAI):** Jadwal utama 3x, selingan 2x. Maks 30 menit/sesi. TANPA distraksi (HP/TV).\n"

    teks_tb = ""
    if skor_tb >= 6:
        teks_tb = f"\n- 🦠 **PERINGATAN TB (Skor = {skor_tb}):** Diagnosis Klinis TB (Skor ≥ 6). WAJIB RUJUK ke poli anak / DOTS Puskesmas untuk OAT!"
    elif skor_tb > 0:
        teks_tb = f"\n- 🦠 **Observasi TB (Skor = {skor_tb}):** Ada indikasi risiko. Pantau ketat, evaluasi uji Mantoux/Rontgen jika gejala menetap."

    teks_gtm = ""
    if is_gtm == "Ya":
        teks_gtm = "\n- 🚫 **TATA LAKSANA GTM:** Lakukan evaluasi *Red Flags* (sariawan, infeksi). Terapkan *Feeding Rules* sangat ketat! Jangan paksa anak makan (*force feeding*)."

    if red_flags_aktif or haz < -3 or skor_tb >= 6:
        gizi_tambahan = "- 🥩 **Gizi:** Wajib Protein Hewani setiap porsi MPASI!" if usia_bulan >= 6 else "- 🍼 **Laktasi:** Fokus perbaikan manajemen laktasi."
        tindak_lanjut = f"RUJUKAN (MERAH): Red flags / Severely Stunted / Suspek TB!\n\n{teks_kalori}\n\n**Edukasi & Tata Laksana Klinis:**\n{teks_usia}\n{gizi_tambahan}\n{teks_feeding_rules}{teks_tb}{teks_gtm}\n- 🩺 **Medis:** Skrining ketat penyakit penyerta."
    elif haz < -2:
        tgl_evaluasi = (tgl_ukur + relativedelta(days=14)).strftime('%d %B %Y')
        gizi_tambahan = "- 🥚 **Gizi:** Kejar tumbuh! Berikan PMT Puskesmas + ekstra Protein Hewani (1-2 telur/hari)." if usia_bulan >= 6 else "- 🍼 **Laktasi:** Pantau ketat kecukupan ASI."
        tindak_lanjut = f"PMT PEMULIHAN (KUNING): Status Stunted, evaluasi ulang {tgl_evaluasi}\n\n{teks_kalori}\n\n**Edukasi & Tata Laksana Klinis:**\n{teks_usia}\n{gizi_tambahan}\n{teks_feeding_rules}{teks_tb}{teks_gtm}"
    else:
        gizi_tambahan = "- 🥩 **Gizi:** Lanjutkan MPASI menu lengkap gizi seimbang." if usia_bulan >= 6 else "- 🤱 **Laktasi:** Lanjutkan pemberian ASI Eksklusif."
        tindak_lanjut = f"PEMANTAUAN RUTIN (HIJAU): Pertumbuhan Baik. Kontrol Posyandu.\n\n{teks_kalori}\n\n**Edukasi & Tata Laksana Klinis:**\n{teks_usia}\n{gizi_tambahan}\n{teks_feeding_rules}{teks_tb}{teks_gtm}\n- 🏃 **Tumbuh Kembang:** Stimulasi rutin."
        
    return status_bb, status_tb, status_gizi, tindak_lanjut.strip()

# --- 5. SISTEM TABS MULTI-HALAMAN ---
tab1, tab2, tab3 = st.tabs(["🧮 Skrining & Kurva", "📖 Buku KIA & Monitoring", "⚖️ Referensi Klinis"])

# ==========================================
# TAB 1: KALKULATOR & KURVA PERTUMBUHAN + REGISTRASI/SEARCH PASIEN
# ==========================================
with tab1:
    if "pasien_id_terpilih" not in st.session_state:
        st.session_state.pasien_id_terpilih = None
    if "sudah_dihitung" not in st.session_state:
        st.session_state.sudah_dihitung = False

    col_form, col_hasil = st.columns([1.2, 2.5])

    with col_form:
        st.header("🔍 Cari Pasien")
        kata_kunci = st.text_input("No. RM atau Nama Pasien", key="kata_kunci_cari")
        hasil_cari = cari_pasien(kata_kunci) if kata_kunci else []

        if hasil_cari:
            opsi = {f"{r['nama']} (RM: {r['no_rm'] or '-'})": r["id"] for r in hasil_cari}
            pilihan = st.selectbox("Pasien ditemukan, pilih salah satu:", list(opsi.keys()))
            if st.button("✅ Gunakan Data Pasien Ini"):
                st.session_state.pasien_id_terpilih = opsi[pilihan]
                st.session_state.sudah_dihitung = False
                st.rerun()
        elif kata_kunci:
            st.info("Tidak ditemukan. Akan didaftarkan sebagai pasien baru di bawah.")

        if st.session_state.pasien_id_terpilih:
            if st.button("➕ Ganti ke Pasien Baru"):
                st.session_state.pasien_id_terpilih = None
                st.session_state.sudah_dihitung = False
                st.rerun()

        st.markdown("---")
        st.header("📋 Form Data Pasien")

        pasien_lama = get_pasien(st.session_state.pasien_id_terpilih) if st.session_state.pasien_id_terpilih else None

        if pasien_lama:
            riwayat_lama = get_riwayat_pengukuran(pasien_lama["id"])
            st.success(f"📖 Pasien terdaftar: **{pasien_lama['nama']}** — {len(riwayat_lama)}x pengukuran"
                       + (f", terakhir {riwayat_lama[-1]['tanggal_ukur']}" if riwayat_lama else ""))

        with st.form("form_skrining"):
            if pasien_lama:
                no_rm = st.text_input("No. RM", value=pasien_lama["no_rm"] or "", disabled=True)
                nama_anak = st.text_input("Nama Anak", value=pasien_lama["nama"], disabled=True)
                nama_ibu = st.text_input("Nama Ibu", value=pasien_lama["nama_ibu"] or "", disabled=True)
                alamat = st.text_area("Alamat / RT RW", value=pasien_lama["alamat"] or "", disabled=True)
                tgl_lahir_default = datetime.strptime(pasien_lama["tanggal_lahir"], "%Y-%m-%d").date()
                jk = pasien_lama["jenis_kelamin"]
                st.text_input("Jenis Kelamin", value=jk, disabled=True)
            else:
                no_rm = st.text_input("No. RM Puskesmas (opsional)")
                nama_anak = st.text_input("Nama Anak")
                nama_ibu = st.text_input("Nama Ibu")
                alamat = st.text_area("Alamat / RT RW")
                tgl_lahir_default = date(2023, 1, 1)
                jk = st.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"])

            st.markdown("---")
            tgl_ukur = st.date_input("Tanggal Pengukuran", value=datetime.today())
            tgl_lahir = st.date_input("Tanggal Lahir", value=tgl_lahir_default, min_value=datetime(2020, 1, 1), max_value=datetime.today(), disabled=bool(pasien_lama))

            bb = st.number_input("Berat Badan (kg)", min_value=1.0, step=0.1)
            tb = st.number_input("Tinggi/Panjang Badan (cm)", min_value=30.0, step=0.1)
            red_flags_aktif = st.checkbox("🚨 Tanda Bahaya (Kelainan Bawaan / BB stagnan)")
            catatan_kunjungan = st.text_area("Catatan kunjungan (opsional)")
            
            st.markdown("---")
            st.markdown("**🩺 Skrining Klinis Tambahan (Opsional)**")
            
            with st.expander("🍽️ Skrining GTM & Feeding Rules (Buka jika anak susah makan)"):
                st.info("Skrining Inappropriate Feeding Practice (Berdasarkan IDAI). Centang 'Ya' jika terjadi kebiasaan berikut:")
                gtm_1 = st.selectbox("1. Waktu makan berlangsung lebih dari 30 menit?", ["Tidak (0)", "Ya (1)"])
                gtm_2 = st.selectbox("2. Makan sambil main HP, nonton TV, atau jalan-jalan?", ["Tidak (0)", "Ya (1)"])
                gtm_3 = st.selectbox("3. Anak dipaksa makan (force feeding) atau dikejar-kejar?", ["Tidak (0)", "Ya (1)"])
                gtm_4 = st.selectbox("4. Sering minum susu / ngemil berdekatan dengan jam makan?", ["Tidak (0)", "Ya (1)"])
                
            with st.expander("🦠 Form Skoring TB IDAI (Buka jika ada indikasi)"):
                st.info("Catatan: Skor ≥ 6 (dengan max 1 skor/parameter) mengindikasikan diagnosis klinis TB.")
                tb_1 = st.selectbox("1. Kontak dengan pasien TB Paru", ["Tidak Jelas (0)", "Laporan Keluarga / BTA tidak diketahui (2)", "BTA Positif (3)"])
                tb_2 = st.selectbox("2. Uji Tuberkulin (Mantoux)", ["Negatif / Tidak Dilakukan (0)", "Positif / ≥10mm (3)"])
                tb_3 = st.selectbox("3. Berat Badan / Keadaan Gizi", ["Normal (0)", "BB/TB < 90% atau BB/U < 80% (1)", "Klinis Gizi Buruk atau BGM (2)"])
                tb_4 = st.selectbox("4. Demam yang tidak diketahui penyebabnya (≥2 minggu)", ["Tidak (0)", "Ya (1)"])
                tb_5 = st.selectbox("5. Batuk kronik (≥3 minggu)", ["Tidak (0)", "Ya (1)"])
                tb_6 = st.selectbox("6. Pembesaran kelenjar limfe koli/aksila (>1cm, >1 kelenjar)", ["Tidak (0)", "Ya (1)"])
                tb_7 = st.selectbox("7. Pembengkakan tulang/sendi", ["Tidak (0)", "Ya (1)"])
                tb_8 = st.selectbox("8. Foto rontgen toraks", ["Normal / Tidak Diperiksa (0)", "Kesan TB (1)"])
            
            hitung_ditekan = st.form_submit_button("🧮 Hitung & Analisis Gizi", type="primary", use_container_width=True)

        selisih = relativedelta(tgl_ukur, tgl_lahir)
        usia_bulan = selisih.years * 12 + selisih.months + (selisih.days / 30.44)
        st.info(f"Usia Presisi: {selisih.years} Tahun, {selisih.months} Bulan, {selisih.days} Hari")

        if hitung_ditekan:
            skor_tb_total = (
                int(tb_1.split("(")[1].split(")")[0]) + int(tb_2.split("(")[1].split(")")[0]) +
                int(tb_3.split("(")[1].split(")")[0]) + int(tb_4.split("(")[1].split(")")[0]) +
                int(tb_5.split("(")[1].split(")")[0]) + int(tb_6.split("(")[1].split(")")[0]) +
                int(tb_7.split("(")[1].split(")")[0]) + int(tb_8.split("(")[1].split(")")[0])
            )
            st.session_state.skor_tb_input = skor_tb_total
            
            skor_gtm_total = (
                int(gtm_1.split("(")[1].split(")")[0]) + int(gtm_2.split("(")[1].split(")")[0]) +
                int(gtm_3.split("(")[1].split(")")[0]) + int(gtm_4.split("(")[1].split(")")[0])
            )
            st.session_state.is_gtm_input = "Ya" if skor_gtm_total > 0 else "Tidak"
            
            if not pasien_lama and (nama_anak == "" or nama_ibu == ""):
                st.error("⚠️ Nama Anak dan Ibu wajib diisi!")
            elif usia_bulan > 60 or usia_bulan < 0:
                st.error("⚠️ Standar WHO Child Growth Standards di aplikasi ini hanya berlaku 0-60 bulan.")
            else:
                waz, haz, whz, errors = hitung_zscore(bb, tb, usia_bulan, jk)
                st.session_state.waz, st.session_state.haz, st.session_state.whz = waz, haz, whz
                st.session_state.zscore_errors = errors
                st.session_state.sudah_dihitung = True

    # ---- BAGIAN KANAN: HASIL DAN KURVA ----
    with col_hasil:
        if st.session_state.sudah_dihitung:
            waz, haz, whz = st.session_state.waz, st.session_state.haz, st.session_state.whz

            if st.session_state.get("zscore_errors"):
                for err in st.session_state.zscore_errors:
                    st.error(f"⚠️ Gagal menghitung salah satu Z-score: {err}")

            if haz is None or whz is None or waz is None:
                st.warning("Data tidak lengkap untuk dihitung. Periksa kembali usia/berat/tinggi badan.")
            else:
                status_bb, status_tb, status_gizi, tindak_lanjut = tentukan_status_dan_tindak_lanjut(
                    waz, haz, whz, red_flags_aktif, tgl_ukur, bb, usia_bulan,
                    st.session_state.skor_tb_input, st.session_state.is_gtm_input
                )
                nama_tampil = pasien_lama["nama"] if pasien_lama else nama_anak

                st.success(f"✅ Analisis Gizi Kemenkes untuk **{nama_tampil}** ({jk}, usia {usia_bulan:.1f} bulan).")

                col_res1, col_res2, col_res3 = st.columns(3)
                col_res1.metric("Berat/Umur (BB/U)", f"{waz} SD", status_bb, delta_color="off" if (waz < -2) else "normal")
                col_res2.metric("Tinggi/Umur (TB/U)", f"{haz} SD", status_tb, delta_color="off" if (haz < -2) else "normal")
                col_res3.metric("Berat/Tinggi (BB/TB)", f"{whz} SD", status_gizi, delta_color="off" if (whz < -2 or whz > 2) else "normal")

                st.markdown("---")
                st.subheader("📈 Kurva Pertumbuhan Anak (WHO)")
                
                tab_kurva_bb, tab_kurva_tb, tab_kurva_gizi = st.tabs(["📉 BB/Umur", "📉 TB/Umur", "📉 BB/Tinggi (Gizi)"])
                
                with tab_kurva_tb:
                    df_tb = get_kurva_haz(jk)
                    fig_tb = go.Figure()
                    fig_tb.add_trace(go.Scatter(x=df_tb["bulan"], y=df_tb["median"], mode='lines', name='0 SD', line=dict(color='green')))
                    fig_tb.add_trace(go.Scatter(x=df_tb["bulan"], y=df_tb["sd2neg"], mode='lines', name='-2 SD (Stunted)', line=dict(color='orange', dash='dash')))
                    fig_tb.add_trace(go.Scatter(x=df_tb["bulan"], y=df_tb["sd3neg"], mode='lines', name='-3 SD (Sev. Stunted)', line=dict(color='red', dash='dot')))
                    
                    if pasien_lama:
                        riw = get_riwayat_pengukuran(pasien_lama["id"])
                        if riw: fig_tb.add_trace(go.Scatter(x=[r["usia_bulan"] for r in riw], y=[r["tb"] for r in riw], mode='lines+markers', name='Riwayat', line=dict(color='#7d3c98', dash='dot')))
                    fig_tb.add_trace(go.Scatter(x=[usia_bulan], y=[tb], mode='markers', name='Kunjungan Ini', marker=dict(color='blue', size=14, symbol='star')))
                    fig_tb.update_layout(xaxis_title="Usia (Bulan)", yaxis_title="Tinggi/Panjang (cm)", height=350, margin=dict(l=0, r=0, t=20, b=0))
                    st.plotly_chart(fig_tb, use_container_width=True)

                with tab_kurva_bb:
                    df_bb = get_kurva_waz(jk)
                    fig_bb = go.Figure()
                    fig_bb.add_trace(go.Scatter(x=df_bb["bulan"], y=df_bb["median"], mode='lines', name='0 SD', line=dict(color='green')))
                    fig_bb.add_trace(go.Scatter(x=df_bb["bulan"], y=df_bb["sd2neg"], mode='lines', name='-2 SD (Kurang)', line=dict(color='orange', dash='dash')))
                    fig_bb.add_trace(go.Scatter(x=df_bb["bulan"], y=df_bb["sd3neg"], mode='lines', name='-3 SD (Buruk)', line=dict(color='red', dash='dot')))
                    
                    if pasien_lama:
                        riw = get_riwayat_pengukuran(pasien_lama["id"])
                        if riw: fig_bb.add_trace(go.Scatter(x=[r["usia_bulan"] for r in riw], y=[r["bb"] for r in riw], mode='lines+markers', name='Riwayat', line=dict(color='#7d3c98', dash='dot')))
                    fig_bb.add_trace(go.Scatter(x=[usia_bulan], y=[bb], mode='markers', name='Kunjungan Ini', marker=dict(color='blue', size=14, symbol='star')))
                    fig_bb.update_layout(xaxis_title="Usia (Bulan)", yaxis_title="Berat Badan (kg)", height=350, margin=dict(l=0, r=0, t=20, b=0))
                    st.plotly_chart(fig_bb, use_container_width=True)

                with tab_kurva_gizi:
                    df_whz, is_length = get_kurva_whz(jk, usia_bulan)
                    label_x = "Panjang Badan (cm)" if is_length else "Tinggi Badan (cm)"
                    
                    fig_whz = go.Figure()
                    fig_whz.add_trace(go.Scatter(x=df_whz["tinggi"], y=df_whz["sd3pos"], mode='lines', name='+3 SD (Obesitas)', line=dict(color='red', dash='dot')))
                    fig_whz.add_trace(go.Scatter(x=df_whz["tinggi"], y=df_whz["sd2pos"], mode='lines', name='+2 SD (Gizi Lebih)', line=dict(color='orange', dash='dash')))
                    fig_whz.add_trace(go.Scatter(x=df_whz["tinggi"], y=df_whz["median"], mode='lines', name='0 SD (Normal)', line=dict(color='green')))
                    fig_whz.add_trace(go.Scatter(x=df_whz["tinggi"], y=df_whz["sd2neg"], mode='lines', name='-2 SD (Gizi Kurang)', line=dict(color='orange', dash='dash')))
                    fig_whz.add_trace(go.Scatter(x=df_whz["tinggi"], y=df_whz["sd3neg"], mode='lines', name='-3 SD (Gizi Buruk)', line=dict(color='red', dash='dot')))
                    
                    if pasien_lama:
                        riw = get_riwayat_pengukuran(pasien_lama["id"])
                        if riw: 
                            x_riw = [r["tb"] for r in riw if (r["usia_bulan"] < 24) == is_length]
                            y_riw = [r["bb"] for r in riw if (r["usia_bulan"] < 24) == is_length]
                            if x_riw: fig_whz.add_trace(go.Scatter(x=x_riw, y=y_riw, mode='lines+markers', name='Riwayat', line=dict(color='#7d3c98', dash='dot')))
                    
                    fig_whz.add_trace(go.Scatter(x=[tb], y=[bb], mode='markers', name='Kunjungan Ini', marker=dict(color='blue', size=14, symbol='star')))
                    fig_whz.update_layout(xaxis_title=label_x, yaxis_title="Berat Badan (kg)", height=350, margin=dict(l=0, r=0, t=20, b=0))
                    st.plotly_chart(fig_whz, use_container_width=True)

                st.markdown("---")
                st.subheader("💡 Intervensi & Tindak Lanjut")
                warna_map = {"RUJUKAN": st.error, "PMT": st.warning, "PEMANTAUAN": st.success}
                for kunci, fungsi in warna_map.items():
                    if tindak_lanjut.startswith(kunci):
                        fungsi(f"**{tindak_lanjut}**")

                st.markdown("---")
                if st.button("💾 Simpan ke Rekam Pasien (Database)", type="primary"):
                    if pasien_lama:
                        pasien_id = pasien_lama["id"]
                    else:
                        pasien_id = simpan_pasien_baru(no_rm or None, nama_anak, nama_ibu, alamat, tgl_lahir, jk)
                    
                    status_gabungan = f"BB/U: {status_bb} | TB/U: {status_tb} | BB/TB: {status_gizi}"
                    
                    simpan_pengukuran(
                        pasien_id, tgl_ukur, usia_bulan, bb, tb, waz, haz, whz,
                        status_gabungan, red_flags_aktif, tindak_lanjut, catatan_kunjungan, 
                        st.session_state.skor_tb_input, st.session_state.is_gtm_input
                    )
                    st.success("🎉 Data disimpan ke rekam pasien! Cek riwayatnya di tab Buku KIA.")
                    st.session_state.sudah_dihitung = False
                    st.session_state.pasien_id_terpilih = pasien_id
                    st.rerun()
        else:
            st.info("👈 Silakan isi form pendaftaran/pengukuran di sebelah kiri, lalu klik 'Hitung & Analisis Gizi' untuk menampilkan hasil Z-Score.")
            
# ==========================================
# TAB 2: BUKU KIA DIGITAL & MONITORING PASIEN
# ==========================================
with tab2:
    st.subheader("📖 Buku KIA Digital — Riwayat & Monitoring Per Pasien")

    semua_pasien = get_semua_pasien()
    if not semua_pasien:
        st.info("Belum ada pasien terdaftar. Simpan data pasien pertama lewat tab 🧮 Skrining & Kurva.")
    else:
        opsi_monitor = {f"{p['nama']} (RM: {p['no_rm'] or '-'})": p["id"] for p in semua_pasien}
        pilih_monitor = st.selectbox("Pilih pasien:", list(opsi_monitor.keys()))
        pid_monitor = opsi_monitor[pilih_monitor]
        detail = get_pasien(pid_monitor)
        riwayat = get_riwayat_pengukuran(pid_monitor)

        st.markdown(f"""
        <div class="kia-card">
        <b>Nama:</b> {detail['nama']} &nbsp;|&nbsp; <b>No. RM:</b> {detail['no_rm'] or '-'}<br>
        <b>Nama Ibu:</b> {detail['nama_ibu'] or '-'} &nbsp;|&nbsp; <b>JK:</b> {detail['jenis_kelamin']}<br>
        <b>Tanggal Lahir:</b> {detail['tanggal_lahir']} &nbsp;|&nbsp; <b>Alamat:</b> {detail['alamat'] or '-'}
        </div>
        """, unsafe_allow_html=True)

        if not riwayat:
            st.warning("Pasien ini belum punya riwayat pengukuran.")
        else:
            df_riwayat = pd.DataFrame([{
                "Tanggal": r["tanggal_ukur"], "Usia (bln)": round(r["usia_bulan"], 1),
                "BB (kg)": r["bb"], "TB (cm)": r["tb"],
                "WAZ": r["waz"], "HAZ": r["haz"], "WHZ": r["whz"],
                "Status Gizi": r["status_gizi"],
                "Red Flag": "🚨" if r["red_flag"] else "-",
                "Tindak Lanjut": r["tindak_lanjut"],
                "Catatan": r["catatan"] or "-"
            } for r in riwayat])
            st.markdown("**📋 Riwayat Pengukuran**")
            st.dataframe(df_riwayat, use_container_width=True, hide_index=True)

            st.markdown("**📈 Tren Pertumbuhan Z-Score (Riwayat Kunjungan)**")
            tab_tren_bb, tab_tren_tb, tab_tren_gizi = st.tabs(["📉 Tren BB/Umur (WAZ)", "📉 Tren TB/Umur (HAZ)", "📉 Tren Gizi (WHZ)"])
            
            with tab_tren_bb:
                fig_waz = go.Figure()
                fig_waz.add_trace(go.Scatter(x=df_riwayat["Tanggal"], y=df_riwayat["WAZ"], mode='lines+markers',
                                             name='WAZ', line=dict(color='#28B463', width=3), marker=dict(size=9)))
                fig_waz.add_hline(y=-2, line_dash="dash", line_color="orange", annotation_text="-2 SD (Kurang)")
                fig_waz.add_hline(y=-3, line_dash="dot", line_color="red", annotation_text="-3 SD (Sangat Kurang)")
                fig_waz.update_layout(xaxis_title="Tanggal Kunjungan", yaxis_title="Z-Score BB/U (WAZ)", height=350, margin=dict(t=20))
                st.plotly_chart(fig_waz, use_container_width=True)

            with tab_tren_tb:
                fig_haz = go.Figure()
                fig_haz.add_trace(go.Scatter(x=df_riwayat["Tanggal"], y=df_riwayat["HAZ"], mode='lines+markers',
                                             name='HAZ', line=dict(color='#1A5276', width=3), marker=dict(size=9)))
                fig_haz.add_hline(y=-2, line_dash="dash", line_color="orange", annotation_text="-2 SD (Pendek)")
                fig_haz.add_hline(y=-3, line_dash="dot", line_color="red", annotation_text="-3 SD (Sangat Pendek)")
                fig_haz.update_layout(xaxis_title="Tanggal Kunjungan", yaxis_title="Z-Score TB/U (HAZ)", height=350, margin=dict(t=20))
                st.plotly_chart(fig_haz, use_container_width=True)
                
            with tab_tren_gizi:
                fig_whz = go.Figure()
                fig_whz.add_trace(go.Scatter(x=df_riwayat["Tanggal"], y=df_riwayat["WHZ"], mode='lines+markers',
                                             name='WHZ', line=dict(color='#8E44AD', width=3), marker=dict(size=9)))
                fig_whz.add_hline(y=3, line_dash="dot", line_color="red", annotation_text="+3 SD (Obesitas)")
                fig_whz.add_hline(y=2, line_dash="dash", line_color="orange", annotation_text="+2 SD (Gizi Lebih)")
                fig_whz.add_hline(y=-2, line_dash="dash", line_color="orange", annotation_text="-2 SD (Gizi Kurang)")
                fig_whz.add_hline(y=-3, line_dash="dot", line_color="red", annotation_text="-3 SD (Gizi Buruk)")
                fig_whz.update_layout(xaxis_title="Tanggal Kunjungan", yaxis_title="Z-Score BB/TB (WHZ)", height=350, margin=dict(t=20))
                st.plotly_chart(fig_whz, use_container_width=True)

            df_pasien_ini = pd.DataFrame([dict(detail)])
            excel_pasien = buat_file_excel(
                df_pasien=df_pasien_ini,
                df_gabungan=get_semua_data_gabungan(),
                df_riwayat_pasien_terpilih=df_riwayat,
                nama_pasien_terpilih=detail["nama"]
            )
            st.download_button(
                "📥 Export Riwayat Pasien Ini ke Excel",
                data=excel_pasien,
                file_name=f"riwayat_{detail['nama'].replace(' ', '_')}_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.markdown("---")
            st.markdown("### ⚙️ Manajemen & Koreksi Data (Khusus Admin)")
            with st.expander("Buka Panel Koreksi Data (Butuh PIN)"):
                pin_admin = st.text_input("Masukkan PIN Admin Puskesmas:", type="password")
                
                if pin_admin == "sumber123":
                    st.success("Akses Terbuka. Gunakan dengan hati-hati!")
                    
                    if riwayat:
                        st.markdown("**1. Hapus Kunjungan/Pengukuran Tertentu** (Jika salah ketik BB/TB)")
                        opsi_hapus = {f"Tgl: {r['tanggal_ukur']} | BB: {r['bb']} kg | TB: {r['tb']} cm": r["id"] for r in riwayat}
                        pilih_hapus = st.selectbox("Pilih data kunjungan yang salah:", list(opsi_hapus.keys()))
                        id_pengukuran_salah = opsi_hapus[pilih_hapus]
                        
                        if st.button("🗑️ Hapus Pengukuran Ini"):
                            hapus_pengukuran(id_pengukuran_salah)
                            st.success("Satu baris data pengukuran berhasil dihapus!")
                            st.rerun() 
                    
                    st.markdown("---")
                    
                    st.markdown("**2. Hapus SELURUH Data Pasien Ini**")
                    st.warning("⚠️ Peringatan: Aksi ini akan menghapus permanen identitas pasien dan seluruh riwayat Buku KIA-nya!")
                    konfirmasi_hapus = st.checkbox(f"Saya yakin ingin menghapus permanen data {detail['nama']}")
                    
                    if konfirmasi_hapus:
                        if st.button("🚨 Hapus Permanen Pasien", type="primary"):
                            sukses, pesan_error = hapus_pasien_total(pid_monitor)
                            
                            if sukses:
                                if st.session_state.get("pasien_id_terpilih") == pid_monitor:
                                    st.session_state.pasien_id_terpilih = None
                                    st.session_state.sudah_dihitung = False
                                
                                st.success(f"Identitas {detail['nama']} dan seluruh riwayatnya BERHASIL DIMUSNAHKAN dari database!")
                                st.rerun() 
                            else:
                                st.error(f"Gagal menghapus pasien: {pesan_error}")
                            
                elif pin_admin != "":
                    st.error("❌ PIN Salah! Akses ditolak.")

    st.markdown("---")
    st.subheader("💾 Export Database Lengkap")
    df_semua_pasien = pd.DataFrame([dict(p) for p in semua_pasien]) if semua_pasien else pd.DataFrame()
    df_gabungan_semua = get_semua_data_gabungan()
    if df_gabungan_semua.empty:
        st.info("Belum ada data untuk diexport.")
    else:
        st.caption(f"Total {df_semua_pasien['nama'].nunique() if not df_semua_pasien.empty else 0} pasien, {len(df_gabungan_semua)} baris pengukuran.")
        excel_semua = buat_file_excel(df_pasien=df_semua_pasien, df_gabungan=df_gabungan_semua)
        st.download_button(
            "📥 Download Rekap Seluruh Pasien & Pengukuran (.xlsx)",
            data=excel_semua,
            file_name=f"sipenting_rekap_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        st.caption("Isinya: sheet Daftar Pasien + sheet Rekap Semua Pengukuran + sheet riwayat pasien terpilih (kalau ada). Cocok buat lapor ke dinas kesehatan atau backup manual.")

    st.markdown("---")
    st.subheader("📊 Ringkasan Populasi (data asli dari database, bukan simulasi)")
    col_dash1, col_dash2 = st.columns(2)

    with col_dash1:
        st.markdown("**Distribusi Status Gizi (pengukuran terakhir tiap pasien)**")
        ringkasan = get_ringkasan_status_gizi()
        if ringkasan:
            fig_pie = go.Figure(data=[go.Pie(
                labels=list(ringkasan.keys()), values=list(ringkasan.values()), hole=.3,
                marker_colors=['green', 'orange', 'red']
            )])
            fig_pie.update_layout(height=350, margin=dict(t=0, b=0, l=0, r=0))
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Belum ada data pengukuran.")

    with col_dash2:
        st.markdown("**Tren Kunjungan Skrining (8 Minggu Terakhir)**")
        df_tren = get_tren_kunjungan_mingguan()
        if not df_tren.empty:
            fig_bar = go.Figure(data=[go.Bar(x=df_tren["minggu"], y=df_tren["jumlah"], marker_color='#2E86C1')])
            fig_bar.update_layout(height=350, margin=dict(t=0, b=0, l=0, r=0))
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Belum ada data kunjungan.")

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
    *   **Standar WHO (World Health Organization):** Menggunakan basis data WHO Child Growth Standards (2006) untuk perhitungan indikator Z-Score (BB/U, TB/U, dan BB/TB), dihitung dengan metode LMS resmi.
    *   **Permenkes RI No. 2 Tahun 2020:** Tentang Standar Antropometri Anak di Indonesia.

    **3. Referensi Tata Laksana Gizi (IDAI):**
    *   Pemantauan dan intervensi mengikuti *Pedoman Pelayanan Medis* dari Ikatan Dokter Anak Indonesia (IDAI).
    *   Sistem deteksi dini *Red Flags* (termasuk *faltering growth* / BB tidak naik adekuat atau stagnan selama 14 hari) digunakan sebagai indikator rujukan absolut untuk mencegah kerusakan kognitif ireversibel pada masa 1000 Hari Pertama Kehidupan (HPK).

    **4. Privasi & Keamanan Data (Rekam Medis):**
    Data pasien tersimpan di database lokal Puskesmas Sumber (SQLite, di server internal), tidak dikirim ke pihak ketiga. Akses ke server tetap harus dibatasi hanya untuk tenaga kesehatan berwenang, sesuai prinsip kerahasiaan rekam medis.

    **5. Cakupan Usia:**
    Kalkulator ini hanya berlaku untuk anak usia **0-60 bulan (0-5 tahun)**, sesuai cakupan tabel WHO Child Growth Standards yang digunakan.
    """)
