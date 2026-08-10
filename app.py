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
    conn.execute("""
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
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS pengukuran (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pasien_id INTEGER NOT NULL,
            tanggal_ukur TEXT NOT NULL,
            usia_bulan REAL NOT NULL,
            bb REAL NOT NULL,
            tb REAL NOT NULL,
            waz REAL,
            haz REAL,
            whz REAL,
            status_gizi TEXT,
            red_flag INTEGER DEFAULT 0,
            tindak_lanjut TEXT,
            catatan TEXT,
            dicatat_pada TEXT,
            FOREIGN KEY (pasien_id) REFERENCES pasien(id)
        )
    """)
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


def simpan_pengukuran(pasien_id, tanggal_ukur, usia_bulan, bb, tb, waz, haz, whz,
                       status_gizi, red_flag, tindak_lanjut, catatan=""):
    conn = get_conn()
    conn.execute(
        """INSERT INTO pengukuran
           (pasien_id, tanggal_ukur, usia_bulan, bb, tb, waz, haz, whz,
            status_gizi, red_flag, tindak_lanjut, catatan, dicatat_pada)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (pasien_id, tanggal_ukur.isoformat(), usia_bulan, bb, tb, waz, haz, whz,
         status_gizi, int(red_flag), tindak_lanjut, catatan,
         datetime.now().isoformat(timespec="seconds"))
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
    """Semua pasien + seluruh riwayat pengukurannya, digabung (join) untuk keperluan export/rekap."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT
            pa.no_rm, pa.nama AS nama_anak, pa.nama_ibu, pa.alamat,
            pa.tanggal_lahir, pa.jenis_kelamin,
            pe.tanggal_ukur, pe.usia_bulan, pe.bb, pe.tb,
            pe.waz, pe.haz, pe.whz, pe.status_gizi, pe.red_flag,
            pe.tindak_lanjut, pe.catatan
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

def hapus_pengukuran(pengukuran_id):
    """Menghapus 1 baris riwayat pengukuran yang salah input"""
    conn = get_conn()
    conn.execute("DELETE FROM pengukuran WHERE id = ?", (pengukuran_id,))
    conn.commit()
    conn.close()

def hapus_pasien_total(pasien_id):
    """Menghapus seluruh data pasien beserta semua riwayat pengukurannya"""
    try:
        conn = get_conn()
        # Hapus riwayat pengukurannya dulu (karena ada foreign key constraint)
        conn.execute("DELETE FROM pengukuran WHERE pasien_id = ?", (pasien_id,))
        # Baru hapus data utama pasiennya!
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
    page_icon="https://raw.githubusercontent.com/rizkinovaln03/sipenting-sumber/main/sipenting.png", 
    layout="wide"
)

# --- 2. INJEKSI CUSTOM CSS ---
st.markdown("""
    <style>
    /* --- MENGHILANGKAN JEJAK STREAMLIT --- */
    [data-testid="stHeader"] { visibility: hidden; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }

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
col_logo1, col_logo2, col_teks = st.columns([1, 1.2, 4])

with col_logo1:
    # Logo Puskesmas Sumber (Lama)
    st.markdown(
        '<img src="https://raw.githubusercontent.com/rizkinovaln03/sipenting-sumber/main/logo.png" style="width:100%; border-radius:10px;">',
        unsafe_allow_html=True
    )

with col_logo2:
    # Logo SiPENTING (Baru)
    st.markdown(
        '<img src="https://raw.githubusercontent.com/rizkinovaln03/sipenting-sumber/main/sipenting.png" style="width:100%; border-radius:10px;">', 
        unsafe_allow_html=True
    )
        
with col_teks:
    # Margin agar teks sejajar dengan logo
    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
    st.markdown("<div class='main-title'>SIPENTING</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-title'>Sistem Pencegahan & Edukasi Stunting Terintegrasi | Puskesmas Sumber</div>", unsafe_allow_html=True)

st.markdown("---")

# --- 4. KALKULATOR Z-SCORE RESMI WHO (LMS, via pygrowup) ---
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


def tentukan_status_dan_tindak_lanjut(haz, red_flags_aktif, tgl_ukur):
    if haz is None:
        return None, None
    status = "Normal" if haz >= -2 else ("Severely Stunted" if haz < -3 else "Stunted")
    if red_flags_aktif or haz < -3:
        tindak_lanjut = "RUJUKAN (MERAH): Red flags/severe stunting - rujuk Sp.A/FKTL segera."
    elif haz < -2:
        tgl_evaluasi = (tgl_ukur + relativedelta(days=14)).strftime('%d %B %Y')
        tindak_lanjut = f"PMT PEMULIHAN (KUNING): status {status}, evaluasi ulang {tgl_evaluasi}."
    else:
        tindak_lanjut = "PEMANTAUAN RUTIN (HIJAU): lanjut ASI/MPASI, kontrol bulan depan di Posyandu."
    return status, tindak_lanjut




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

    st.sidebar.header("🔍 Cari Pasien")
    kata_kunci = st.sidebar.text_input("No. RM atau Nama Pasien", key="kata_kunci_cari")
    hasil_cari = cari_pasien(kata_kunci) if kata_kunci else []

    if hasil_cari:
        opsi = {f"{r['nama']} (RM: {r['no_rm'] or '-'}, lahir {r['tanggal_lahir']})": r["id"] for r in hasil_cari}
        pilihan = st.sidebar.selectbox("Pasien ditemukan, pilih salah satu:", list(opsi.keys()))
        if st.sidebar.button("✅ Gunakan Data Pasien Ini"):
            st.session_state.pasien_id_terpilih = opsi[pilihan]
            st.session_state.sudah_dihitung = False
    elif kata_kunci:
        st.sidebar.info("Tidak ditemukan. Akan didaftarkan sebagai pasien baru di bawah.")

    if st.session_state.pasien_id_terpilih:
        if st.sidebar.button("➕ Ganti ke Pasien Baru"):
            st.session_state.pasien_id_terpilih = None
            st.session_state.sudah_dihitung = False

    st.sidebar.markdown("---")
    st.sidebar.header("📋 Form Data Pasien")

    pasien_lama = get_pasien(st.session_state.pasien_id_terpilih) if st.session_state.pasien_id_terpilih else None

    if pasien_lama:
        riwayat_lama = get_riwayat_pengukuran(pasien_lama["id"])
        st.sidebar.success(f"📖 Pasien terdaftar: **{pasien_lama['nama']}** — {len(riwayat_lama)}x pengukuran sebelumnya"
                            + (f", terakhir {riwayat_lama[-1]['tanggal_ukur']}" if riwayat_lama else ""))
        no_rm = st.sidebar.text_input("No. RM", value=pasien_lama["no_rm"] or "", disabled=True)
        nama_anak = st.sidebar.text_input("Nama Anak", value=pasien_lama["nama"], disabled=True)
        nama_ibu = st.sidebar.text_input("Nama Ibu", value=pasien_lama["nama_ibu"] or "", disabled=True)
        alamat = st.sidebar.text_area("Alamat / RT RW", value=pasien_lama["alamat"] or "", disabled=True)
        tgl_lahir_default = datetime.strptime(pasien_lama["tanggal_lahir"], "%Y-%m-%d").date()
        jk = pasien_lama["jenis_kelamin"]
        st.sidebar.text_input("Jenis Kelamin", value=jk, disabled=True)
    else:
        no_rm = st.sidebar.text_input("No. RM Puskesmas (opsional)")
        nama_anak = st.sidebar.text_input("Nama Anak")
        nama_ibu = st.sidebar.text_input("Nama Ibu")
        alamat = st.sidebar.text_area("Alamat / RT RW")
        tgl_lahir_default = date(2023, 1, 1)
        jk = None

    st.sidebar.markdown("---")
    tgl_ukur = st.sidebar.date_input("Tanggal Pengukuran", value=datetime.today())
    tgl_lahir = st.sidebar.date_input("Tanggal Lahir", value=tgl_lahir_default,
                                       min_value=datetime(2020, 1, 1), max_value=datetime.today(),
                                       disabled=bool(pasien_lama))

    selisih = relativedelta(tgl_ukur, tgl_lahir)
    usia_bulan = selisih.years * 12 + selisih.months + (selisih.days / 30.44)
    st.sidebar.info(f"Usia Presisi: {selisih.years} Tahun, {selisih.months} Bulan")
    if usia_bulan > 60:
        st.sidebar.warning("⚠️ Standar WHO pada aplikasi ini hanya mencakup usia 0-60 bulan (0-5 tahun).")

    if not pasien_lama:
        jk = st.sidebar.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"])

    bb = st.sidebar.number_input("Berat Badan (kg)", min_value=1.0, step=0.1)
    tb = st.sidebar.number_input("Tinggi/Panjang Badan (cm)", min_value=30.0, step=0.1)
    red_flags_aktif = st.sidebar.checkbox("🚨 Tanda Bahaya (Kelainan Bawaan / BB stagnan 14 hari)")
    catatan_kunjungan = st.sidebar.text_area("Catatan kunjungan (opsional)")

    if st.sidebar.button("🧮 Hitung & Analisis Gizi", type="primary"):
        if not pasien_lama and (nama_anak == "" or nama_ibu == ""):
            st.sidebar.error("⚠️ Nama Anak dan Ibu wajib diisi!")
        elif usia_bulan > 60 or usia_bulan < 0:
            st.sidebar.error("⚠️ Standar WHO Child Growth Standards di aplikasi ini hanya berlaku 0-60 bulan.")
        else:
            waz, haz, whz, errors = hitung_zscore(bb, tb, usia_bulan, jk)
            st.session_state.waz, st.session_state.haz, st.session_state.whz = waz, haz, whz
            st.session_state.zscore_errors = errors
            st.session_state.sudah_dihitung = True

    if st.session_state.sudah_dihitung:
        waz, haz, whz = st.session_state.waz, st.session_state.haz, st.session_state.whz

        if st.session_state.get("zscore_errors"):
            for err in st.session_state.zscore_errors:
                st.error(f"⚠️ Gagal menghitung salah satu Z-score: {err}")

        if haz is None:
            st.warning("HAZ tidak dapat dihitung untuk input ini. Periksa kembali data usia/tinggi badan.")
        else:
            status_haz, tindak_lanjut = tentukan_status_dan_tindak_lanjut(haz, red_flags_aktif, tgl_ukur)
            nama_tampil = pasien_lama["nama"] if pasien_lama else nama_anak

            st.success(f"✅ Analisis Gizi untuk **{nama_tampil}** ({jk}, usia {usia_bulan:.1f} bulan, {tgl_ukur.strftime('%d %B %Y')}).")

            col_res1, col_res2, col_res3 = st.columns(3)
            col_res1.metric("Berat/Umur (WAZ)", f"{waz} SD" if waz is not None else "N/A")
            col_res2.metric("Tinggi/Umur (HAZ)", f"{haz} SD", status_haz, delta_color="off" if haz < -2 else "normal")
            col_res3.metric("Berat/Tinggi (WHZ)", f"{whz} SD" if whz is not None else "N/A")

            st.markdown("---")
            st.subheader(f"📈 Kurva Pertumbuhan WHO Tinggi/Panjang-menurut-Umur ({jk})")
            df_kurva = get_kurva_haz(jk)

            fig = go.Figure()
            fig.add_trace(go.Scatter(x=df_kurva["bulan"], y=df_kurva["median"], mode='lines',
                                      name='Median (0 SD)', line=dict(color='green', width=2)))
            fig.add_trace(go.Scatter(x=df_kurva["bulan"], y=df_kurva["sd2neg"], mode='lines',
                                      name='-2 SD (Stunted)', line=dict(color='orange', width=2, dash='dash')))
            fig.add_trace(go.Scatter(x=df_kurva["bulan"], y=df_kurva["sd3neg"], mode='lines',
                                      name='-3 SD (Severely Stunted)', line=dict(color='red', width=2, dash='dot')))

            # Titik-titik riwayat pengukuran sebelumnya (kalau pasien lama) -> jadi trajektori, bukan cuma 1 titik
            if pasien_lama:
                riwayat = get_riwayat_pengukuran(pasien_lama["id"])
                if riwayat:
                    x_riwayat = [r["usia_bulan"] for r in riwayat]
                    y_riwayat = [r["tb"] for r in riwayat]
                    fig.add_trace(go.Scatter(x=x_riwayat, y=y_riwayat, mode='lines+markers',
                                              name='Riwayat Kunjungan', line=dict(color='#7d3c98', width=2, dash='dot'),
                                              marker=dict(size=8)))

            fig.add_trace(go.Scatter(x=[usia_bulan], y=[tb], mode='markers', name=f'Kunjungan Ini: {nama_tampil}',
                                      marker=dict(color='blue', size=14, symbol='star')))
            fig.update_layout(title=f"Posisi Tinggi/Panjang Badan pada Kurva WHO ({jk}, 0-60 bulan)",
                               xaxis_title="Usia (Bulan)", yaxis_title="Tinggi/Panjang Badan (cm)", height=420)
            st.plotly_chart(fig, use_container_width=True)
            st.caption("Kurva dari tabel LMS resmi WHO Child Growth Standards (2006). Titik ungu = riwayat kunjungan pasien ini.")

            st.markdown("---")
            st.subheader("💡 Intervensi & Tindak Lanjut")
            warna_map = {"RUJUKAN": st.error, "PMT": st.warning, "PEMANTAUAN": st.success}
            for kunci, fungsi in warna_map.items():
                if tindak_lanjut.startswith(kunci):
                    fungsi(f"**{tindak_lanjut}**")

            st.markdown("---")
            if st.button("💾 Simpan ke Rekam Pasien (Database Puskesmas)", type="primary"):
                if pasien_lama:
                    pasien_id = pasien_lama["id"]
                else:
                    pasien_id = simpan_pasien_baru(no_rm or None, nama_anak, nama_ibu, alamat, tgl_lahir, jk)
                simpan_pengukuran(pasien_id, tgl_ukur, usia_bulan, bb, tb, waz, haz, whz,
                                   status_haz, red_flags_aktif, tindak_lanjut, catatan_kunjungan)
                st.success("🎉 Data pengukuran berhasil disimpan ke rekam pasien! Cek riwayatnya di tab 📖 Buku KIA & Monitoring.")
                st.session_state.sudah_dihitung = False
                st.session_state.pasien_id_terpilih = pasien_id

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

            st.markdown("**📈 Tren Pertumbuhan (HAZ dari waktu ke waktu)**")
            fig_tren = go.Figure()
            fig_tren.add_trace(go.Scatter(x=df_riwayat["Tanggal"], y=df_riwayat["HAZ"], mode='lines+markers',
                                           name='HAZ', line=dict(color='#1A5276', width=3), marker=dict(size=9)))
            fig_tren.add_hline(y=-2, line_dash="dash", line_color="orange", annotation_text="-2 SD (Stunted)")
            fig_tren.add_hline(y=-3, line_dash="dot", line_color="red", annotation_text="-3 SD (Severely Stunted)")
            fig_tren.update_layout(xaxis_title="Tanggal Kunjungan", yaxis_title="HAZ (SD)", height=350)
            st.plotly_chart(fig_tren, use_container_width=True)

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
    # --- MULAI DARI SINI: PANEL MANAJEMEN DATA (ADMIN) ---
            st.markdown("---")
            st.markdown("### ⚙️ Manajemen & Koreksi Data (Khusus Admin)")
            with st.expander("Buka Panel Koreksi Data (Butuh PIN)"):
                pin_admin = st.text_input("Masukkan PIN Admin Puskesmas:", type="password")
                
                # Ganti "sumber123" dengan PIN rahasiamu
                if pin_admin == "sumber123":
                    st.success("Akses Terbuka. Gunakan dengan hati-hati!")
                    
                    # FITUR 1: HAPUS 1 BARIS PENGUKURAN
                    if riwayat:
                        st.markdown("**1. Hapus Kunjungan/Pengukuran Tertentu** (Jika salah ketik BB/TB)")
                        # Buat opsi dropdown berdasarkan riwayat yang ada
                        opsi_hapus = {f"Tgl: {r['tanggal_ukur']} | BB: {r['bb']} kg | TB: {r['tb']} cm": r["id"] for r in riwayat}
                        pilih_hapus = st.selectbox("Pilih data kunjungan yang salah:", list(opsi_hapus.keys()))
                        id_pengukuran_salah = opsi_hapus[pilih_hapus]
                        
                        if st.button("🗑️ Hapus Pengukuran Ini"):
                            hapus_pengukuran(id_pengukuran_salah)
                            st.success("Satu baris data pengukuran berhasil dihapus!")
                            st.rerun() # Refresh halaman otomatis
                    
                    st.markdown("---")
                    
                    # FITUR 2: HAPUS TOTAL PASIEN
                    st.markdown("**2. Hapus SELURUH Data Pasien Ini**")
                    st.warning("⚠️ Peringatan: Aksi ini akan menghapus permanen identitas pasien dan seluruh riwayat Buku KIA-nya!")
                    konfirmasi_hapus = st.checkbox(f"Saya yakin ingin menghapus permanen data {detail['nama']}")
                    
                    if konfirmasi_hapus:
                        if st.button("🚨 Hapus Permanen Pasien", type="primary"):
                            # Panggil fungsi hapus yang baru
                            sukses, pesan_error = hapus_pasien_total(pid_monitor)
                            
                            if sukses:
                                # BERSIHKAN MEMORI DI TAB 1 JUGA (Penting!)
                                if st.session_state.get("pasien_id_terpilih") == pid_monitor:
                                    st.session_state.pasien_id_terpilih = None
                                    st.session_state.sudah_dihitung = False
                                
                                st.success(f"Identitas {detail['nama']} dan seluruh riwayatnya BERHASIL DIMUSNAHKAN dari database!")
                                st.rerun() # Refresh halaman otomatis
                            else:
                                st.error(f"Gagal menghapus pasien: {pesan_error}")
                            
                elif pin_admin != "":
                    st.error("❌ PIN Salah! Akses ditolak.")
            # --- BATAS AKHIR PANEL MANAJEMEN DATA ---

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
