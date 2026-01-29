import io
import numpy as np
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import blue, black, lightgrey
from reportlab.pdfgen import canvas
from datetime import datetime
import os

# Pastikan Anda sudah menginstal reportlab dan openpyxl:
# pip install streamlit numpy pandas reportlab openpyxl

# ===============================================
# === MODIFIKASI: Mapel Tetap (4 Mata Pelajaran) ===
# ===============================================
mapel_tetap = [
    "Bahasa Indonesia",
    "Matematika",
    "Bahasa Inggris",
    "Ilmu Pengetahuan Alam"
]
mapel_semua = mapel_tetap

# Mapping bulan Indonesia
bulan_id = {
    "January": "Januari", "February": "Februari", "March": "Maret",
    "April": "April", "May": "Mei", "June": "Juni",
    "July": "Juli", "August": "Agustus", "September": "September",
    "October": "Oktober", "November": "November", "December": "Desember"
}

st.header("Laporan Hasil Persiapan dan Pemantapan")
st.markdown("---")

# --- Pilihan di Streamlit ---
asesmen_opsi = [
    "TKA/TKAD Dikpora Kab Bantul 1",
    "TKA/TKAD MKKS SMP Kab Bantul 1",
    "TKA/TKAD MKKS SMP Kab Bantul 2",
    "TKA/TKAD Forum MKKS SMP D.I.Yogyakarta",
    "TKA/TKAD Dikpora Kab Bantul 2"
]
sel_asesmen = st.selectbox("Pilih Jenis Asesmen", asesmen_opsi)
tahun_opsi = [f"{th}/{th+1}" for th in range(2025, 2036)]
sel_tahun = st.selectbox("Pilih Tahun Pelajaran", tahun_opsi, index=0)

# ===============================================
# === PENAMBAHAN PILIHAN TANGGAL KEGIATAN ===
# ===============================================

# Definisikan Opsi Tanggal Kegiatan
tgl_kegiatan_opsi = [
    "Tanggal 20 - 23 Oktober 2025",
    "Tanggal 3 - 6 November 2025",
    "Tanggal 19 - 22 Januari 2026",
    "Tanggal 2 - 5 Februari 2026",
    "Tanggal 9 - 12 Maret 2026"
]

# Tambahkan Pilihan Tanggal Kegiatan ke Streamlit
sel_tgl_kegiatan = st.selectbox("Pilih Tanggal Kegiatan", tgl_kegiatan_opsi)

# ===============================================
# === AKHIR PENAMBAHAN PILIHAN TANGGAL KEGIATAN ===
# ===============================================

sel_tgl_ttd = st.date_input("Tanggal Penulisan Tanda Tangan (di dokumen PDF)", datetime.now(), format="DD/MM/YYYY")
st.markdown("---")

# === Template Excel ===
def generate_template():
    score_cols = []
    for i in range(1, 6):
        for m in mapel_semua:
            score_cols.append(f"{m}_TKAD{i}")

    cols = ["Kelas", "NIS", "Nama Siswa"] + score_cols
    df_template = pd.DataFrame(columns=cols)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False, sheet_name="Nilai")
    buffer.seek(0)
    return buffer

st.download_button(
    "📥 Download Template Excel (Nilai TO TKA dan TKAD)",
    data=generate_template(),
    file_name="Template_Nilai_4_Mapel_TKA_TKAD.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Upload file Excel
uploaded = st.file_uploader("Unggah file Excel daftar nilai (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Silakan unggah file Excel nilai")
    st.stop()

# baca file
try:
    df = pd.read_excel(uploaded, engine="openpyxl")
except Exception as e:
    st.error(f"Gagal membaca file Excel: {e}")
    st.stop()

df.columns = df.columns.str.strip()

# Pilih kelas dari file
if "Kelas" not in df.columns:
    st.error("Kolom 'Kelas' tidak ditemukan di file. Pastikan pakai template.")
    st.stop()

kelas_list = sorted(df["Kelas"].astype(str).unique())
sel_kelas = st.selectbox("Pilih Kelas", kelas_list)
df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

# ===============================================
mapel_urut = mapel_tetap
mapel_5_cols = []
for i in range(1, 6):
    for m in mapel_urut:
        mapel_5_cols.append(f"{m}_TKAD{i}")

expected_base = ["Kelas", "NIS", "Nama Siswa"]
missing_base = [c for c in expected_base if c not in df.columns]
if missing_base:
    st.error(f"Kolom wajib hilang: {missing_base}")
    st.stop()

missing_mapel_cols = [c for c in mapel_5_cols if c not in df.columns]
if missing_mapel_cols:
    st.warning(f"Kolom nilai wajib hilang di file Excel dan akan diperlakukan kosong: {missing_mapel_cols[:5]}...")
    for m in missing_mapel_cols:
        df[m] = np.nan
    df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

# === Bersihkan & konversi nilai ===
for col in [c for c in mapel_5_cols if c in df.columns]:
    df[col] = (
        df[col]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
        .str.strip()
    )
    df.loc[df[col] == "", col] = np.nan
    df[col] = pd.to_numeric(df[col], errors="coerce")

cols_to_keep = expected_base + [c for c in mapel_5_cols if c in df.columns]
df = df[cols_to_keep]
df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

siswa_list = df_kelas["Nama Siswa"].astype(str).tolist()
sel_siswa = st.selectbox("Pilih Siswa", ["-- Semua Siswa --"] + siswa_list)

# helper: format skor aman
def format_score(val):
    if pd.isna(val):
        return ""
    try:
        num = float(val)
        # SELALU tampilkan dua angka di belakang koma
        return f"{num:.2f}"
    except Exception:
        # coba convert dari string yang mungkin masih berformat koma
        try:
            num = float(str(val).strip().replace(",", "."))
            return f"{num:.2f}"
        except Exception:
            return str(val)

# Fungsi gambar halaman siswa DENGAN MARGIN
# CATATAN: sel_tgl_kegiatan ditambahkan ke parameter
def draw_student_page(c, row, sel_asesmen, sel_tahun, sel_tgl_kegiatan, mapel_urut, sel_tgl_ttd, asesmen_list):
    width, height = A4

    aksara_path = "assets/aksara_jawa.jpg"
    logo_path = "assets/logo_kiri.png"

    # === PENGATURAN MARGIN ===
    margin_left   = 30 * mm
    margin_right  = 15 * mm
    margin_top    = 25 * mm
    margin_bottom = 20 * mm

    content_width  = width - (margin_left + margin_right)
    y = height - margin_top

    # (Optional) logo kiri atas
    try:
        logo_w = 30*mm
        logo_h = 30*mm
        x_logo = margin_left - 10*mm
        y_logo = height - margin_top - (-10*mm) - logo_h
        c.drawImage(logo_path, x_logo, y_logo, width=logo_w, height=logo_h, preserveAspectRatio=True, mask='auto')
    except Exception as e:
        pass

    # --- 2. KOP SEDERHANA (CENTERED TEXT) ---
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "PEMERINTAH KABUPATEN BANTUL")
    y -= 5*mm
    c.drawCentredString(width/2, y, "DINAS PENDIDIKAN, KEPEMUDAAN, DAN OLAHRAGA")
    y -= 5*mm
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width/2, y, "SMP NEGERI 2 BANGUNTAPAN")
    y -= 1*mm

    # (Optional) aksara jawa
    try:
        aksara_w = 100*mm
        aksara_h = 10*mm
        x_aksara = (width - aksara_w) / 2
        y_aksara = y - aksara_h
        c.drawImage(aksara_path, x_aksara, y_aksara, width=aksara_w, height=aksara_h, preserveAspectRatio=True, mask='auto')
        y = y_aksara - 2*mm
    except Exception as e:
        y -= 4*mm

    c.setFont("Helvetica-Oblique", 10)
    c.drawCentredString(width/2, y, "Jalan Karangsari, Banguntapan, Kabupaten Bantul, Yogyakarta 55198 Telp. 382754")
    y -= 5*mm
    c.setFont("Helvetica", 10)
    c.setFillColor(blue)
    c.drawCentredString(width/2, y, "Website : www.smpn2banguntapan.sch.id Email : smp2banguntapan@yahoo.com")
    c.setFillColor(black)
    y -= 3*mm

    # Garis separator (double)
    c.setLineWidth(1)
    c.line(margin_left, y, width - margin_right, y)
    y -= 1.5*mm
    c.setLineWidth(0.5)
    c.line(margin_left, y, width - margin_right, y)
    y -= 10*mm

    # --- 3. JUDUL DOKUMEN (CENTERED TEXT) ---
    c.setFont("Helvetica-Bold", 12)

    # Baris 1: Judul Utama
    c.drawCentredString(width/2, y, "LAPORAN HASIL PERSIAPAN DAN PEMANTAPAN")
    y -= 6*mm

    # Baris 2: Nama Asesmen (Kapital Semua, menggunakan .upper())
    c.drawCentredString(width/2, y, sel_asesmen.upper())
    y -= 6*mm

    # Baris 3: Tahun Pelajaran
    c.drawCentredString(width/2, y, f"TAHUN PELAJARAN {sel_tahun}")
    y -= 6*mm

    # Baris 4: Tanggal Kegiatan (Hanya huruf awal bulan yang besar)
    # Kita menggunakan sel_tgl_kegiatan apa adanya (tidak .upper())
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, sel_tgl_kegiatan)
    y -= 12*mm # Jarak lebih besar ke identitas siswa

    # Identitas siswa (titik dua rata)
    id_margin_left = margin_left + 10*mm
    label_w = 20*mm
    colon_x = id_margin_left + label_w
    value_x = colon_x + 5
    c.setFont("Helvetica-Bold", 12)
    # Nama
    c.drawString(id_margin_left, y, "Nama")
    c.drawString(colon_x, y, ":")
    c.drawString(value_x, y, " " + str(row.get("Nama Siswa", "")))
    y -= 6*mm
    # NIS
    c.drawString(id_margin_left, y, "NIS")
    c.drawString(colon_x, y, ":")
    c.drawString(value_x, y, " " + str(row.get("NIS", "")))
    y -= 6*mm
    # Kelas
    c.drawString(id_margin_left, y, "Kelas")
    c.drawString(colon_x, y, ":")
    c.drawString(value_x, y, " " + str(row.get("Kelas", "")))
    y -= 10*mm

    # ===============================================
    # === PENGAMBILAN & PERHITUNGAN NILAI PER KOLOM ===
    # ===============================================
    nilai_cols_count = 5
    nilai_data = {}

    scores_by_column = [[] for _ in range(nilai_cols_count)]

    for subj in mapel_urut:
        scores = []
        for i in range(1, nilai_cols_count + 1):
            col_name = f"{subj}_TKAD{i}"
            raw = row.get(col_name, np.nan)

            val = np.nan
            if not pd.isna(raw):
                try:
                    val = float(raw)
                except Exception:
                    pass
            scores.append(val)

            if not np.isnan(val):
                scores_by_column[i-1].append(val)

        nilai_data[subj] = scores

    # --- Hitung Jumlah dan Rata-rata PER KOLOM ---
    jumlah_per_kolom = []
    rata2_per_kolom = []

    for col_scores in scores_by_column:
        if len(col_scores) == 0:
            col_sum = np.nan
            col_avg = np.nan
        else:
            col_sum = sum(col_scores)
            col_count = len(col_scores)
            col_avg = col_sum / col_count if col_count > 0 else np.nan

        jumlah_per_kolom.append(col_sum)
        rata2_per_kolom.append(col_avg)

    # ===============================================
    # === TABEL DRAWING ===
    # ===============================================

    row_height = 7*mm
    font_size = 11

    col_no_w = 15*mm
    col_mapel_w = 70*mm
    col_nilai_single_w = 15*mm

    table_width = col_no_w + col_mapel_w + (col_nilai_single_w * nilai_cols_count)

    x0 = margin_left + (content_width - table_width) / 2
    y0 = y

    nrows_content = len(mapel_urut) + 2
    total_rows_to_draw = nrows_content + 2

    # --- Header Row 1 & 2 Background ---
    header_height = 2 * row_height
    header_y_bottom = y0 - header_height
    c.setFillColor(lightgrey)
    c.rect(x0, header_y_bottom, table_width, header_height, stroke=0, fill=1)
    c.setFillColor(black)

    # --- Draw Grid (SEMUA KOLOM TERPISAH) ---
    c.setLineWidth(0.5)

    # 1. Garis vertikal
    c.line(x0, y0, x0, y0 - total_rows_to_draw*row_height)
    c.line(x0 + col_no_w, y0, x0 + col_no_w, y0 - total_rows_to_draw*row_height)
    c.line(x0 + col_no_w + col_mapel_w, y0, x0 + col_no_w + col_mapel_w, y0 - total_rows_to_draw*row_height)

    # Garis vertikal 5 kolom Nilai
    current_x_grid = x0 + col_no_w + col_mapel_w
    for i in range(1, nilai_cols_count):
        current_x_grid += col_nilai_single_w
        c.line(current_x_grid, y0 - row_height, current_x_grid, y0 - total_rows_to_draw*row_height)

    current_x_grid += col_nilai_single_w
    c.line(current_x_grid, y0, current_x_grid, y0 - total_rows_to_draw*row_height)

    # 2. Garis horizontal untuk semua baris
    for r in range(total_rows_to_draw + 1):
        if r == 1:
            continue
        c.line(x0, y0 - r*row_height, x0 + table_width, y0 - r*row_height)

    # 3. Garis pemisah horizontal antara Header Row 1 dan Row 2 HANYA di kolom Nilai
    nilai_start_x_grid = x0 + col_no_w + col_mapel_w
    c.line(nilai_start_x_grid, y0 - row_height, x0 + table_width, y0 - row_height)


    # --- Header Text Positioning ---
    y_center_2_rows = y0 - row_height
    adj_y_merged = y_center_2_rows - (font_size/3.5)

    y_center_top_row = y0 - row_height/2
    adj_y_top_row = y_center_top_row - (font_size/3.5)


    # 1. Text: No (MERGED)
    c.setFont("Helvetica-Bold", font_size)
    c.drawCentredString(x0 + col_no_w/2, adj_y_merged, "No")

    # 2. Text: Mata Pelajaran (MERGED)
    c.drawCentredString(x0 + col_no_w + col_mapel_w/2, adj_y_merged, "Mata Pelajaran")

    # 3. Text: Nilai TKA/TKAD (Merged Column 3-7, only in Row 1)
    nilai_merged_center_x = x0 + col_no_w + col_mapel_w + (col_nilai_single_w * nilai_cols_count) / 2
    c.drawCentredString(nilai_merged_center_x, adj_y_top_row, "Nilai TKA/TKAD")

    # --- Header Row 2 Text Positioning (1, 2, 3, 4, 5) ---
    c.setFont("Helvetica-Bold", font_size)
    y_center_bottom_row = y0 - (1.5 * row_height)
    adj_y_bottom_row = y_center_bottom_row - (font_size/3.5)

    # Header 5 Kolom Nilai
    current_x = x0 + col_no_w + col_mapel_w
    for i in range(nilai_cols_count):
        center_x = current_x + col_nilai_single_w / 2
        c.drawCentredString(center_x, adj_y_bottom_row, str(i+1))
        current_x += col_nilai_single_w

    c.setFont("Helvetica", font_size)
    y_text = y0 - 2 * row_height

    # isi tabel
    for i, subj in enumerate(mapel_urut, start=1):
        cell_middle = y_text - row_height/2
        adj_y = cell_middle - (font_size/3.5)

        # Kolom No & Mapel
        c.drawCentredString(x0 + col_no_w/2, adj_y, str(i))
        c.drawString(x0 + col_no_w + 2*mm, adj_y, subj)

        # Kolom Nilai (5 kolom)
        scores = nilai_data.get(subj, [np.nan] * nilai_cols_count)
        current_x = x0 + col_no_w + col_mapel_w
        for score in scores:
            val_str = format_score(score)
            c.drawCentredString(current_x + col_nilai_single_w / 2, adj_y, val_str)
            current_x += col_nilai_single_w

        y_text -= row_height

    # === Jumlah & Rata-rata (Per Kolom) ===
    # Jumlah
    cell_middle = y_text - row_height/2
    adj_y = cell_middle - (font_size/3.5)
    c.setFont("Helvetica-Bold", font_size)
    c.drawString(x0 + col_no_w + 2*mm, adj_y, "Jumlah")
    current_x = x0 + col_no_w + col_mapel_w
    for col_sum in jumlah_per_kolom:
        # Jika kolom tidak ada nilainya (semua kosong atau NaN), jangan tampilkan apapun
        if pd.isna(col_sum):
            val_str = ""  # kosongkan
        else:
            val_str = format_score(col_sum)
        c.drawCentredString(current_x + col_nilai_single_w / 2, adj_y, val_str)
        current_x += col_nilai_single_w
    y_text -= row_height

    # Rata-rata
    cell_middle = y_text - row_height/2
    adj_y = cell_middle - (font_size/3.5)
    c.drawString(x0 + col_no_w + 2*mm, adj_y, "Rata-rata")
    current_x = x0 + col_no_w + col_mapel_w
    for col_avg in rata2_per_kolom:
        if pd.isna(col_avg):
            val_str = ""
        else:
            val_str = format_score(col_avg)
        c.drawCentredString(current_x + col_nilai_single_w / 2, adj_y, val_str)
        current_x += col_nilai_single_w

    y_text -= row_height + 2

    # === KETERANGAN ===
    keterangan_x = margin_left
    y_keterangan = y_text - 5*mm

    c.setFont("Helvetica-Bold", 9)
    c.drawString(keterangan_x, y_keterangan, "Keterangan:")
    y_keterangan -= 5*mm

    c.setFont("Helvetica", 9)

    # Keterangan Nomor (1-5) akan menggunakan tanggal kegiatan
    for i, item in enumerate(asesmen_list, start=1):
        # Ambil tanggal dari list tgl_kegiatan_opsi
        tgl_item = tgl_kegiatan_opsi[i-1] if i-1 < len(tgl_kegiatan_opsi) else "TANGGAL TIDAK DIKETAHUI"
        c.drawString(keterangan_x + 5*mm, y_keterangan, f"{i}. {item} ({tgl_item})")
        y_keterangan -= 5*mm

    y_text = y_keterangan + -3*mm

    # tanda tangan
    # GUNAKAN TANGGAL PILIHAN DARI STREAMLIT (sel_tgl_ttd)
    ttd_date = sel_tgl_ttd
    bulan_eng = ttd_date.strftime('%B')
    tgl = f"{ttd_date.day} {bulan_id[bulan_eng]} {ttd_date.year}"

    # Posisikan tanda tangan berdasarkan margin_right dan margin_bottom
    x_ttd = width - margin_right - 70*mm
    y_ttd_start = margin_bottom + 60*mm

    c.setFont("Helvetica", 12)
    c.drawString(x_ttd, y_ttd_start, f"Banguntapan, {tgl}")
    y_ttd_start -= 8*mm
    c.drawString(x_ttd, y_ttd_start, "Mengetahui,")
    y_ttd_start -= 5*mm
    c.drawString(x_ttd, y_ttd_start, "Kepala Sekolah,")
    y_ttd_start -= -1*mm
    # Tambah gambar tanda tangan (cek file)
    ttd_path = "assets/ttd_kepsek.jpeg"
    if os.path.exists(ttd_path):
        try:
            c.drawImage(ttd_path, x_ttd, y_ttd_start - 22*mm, width=40*mm, height=20*mm, mask="auto")
        except Exception:
            pass

    # Nama & NIP (tetap ditampilkan)
    y_ttd_after = y_ttd_start - 25*mm
    c.drawString(x_ttd, y_ttd_after, "Alina Fiftiyani Nurjannah, M.Pd.")
    y_ttd_after -= 6*mm
    c.drawString(x_ttd, y_ttd_after, "NIP 198001052009032006")

# PDF generator
# CATATAN: sel_tgl_kegiatan ditambahkan ke parameter
def make_pdf_for_student(row, mapel_urut, sel_tgl_ttd, asesmen_list, sel_tgl_kegiatan):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    draw_student_page(c, row, sel_asesmen, sel_tahun, sel_tgl_kegiatan, mapel_urut, sel_tgl_ttd, asesmen_list)
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

def make_pdf_for_class(df_kelas, mapel_urut, sel_tgl_ttd, asesmen_list, sel_tgl_kegiatan):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    for _, row in df_kelas.iterrows():
        draw_student_page(c, row, sel_asesmen, sel_tahun, sel_tgl_kegiatan, mapel_urut, sel_tgl_ttd, asesmen_list)
        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

st.markdown("---")
st.subheader("Pilih Siswa & Unduh Laporan")

# Buttons (hanya tampil jika df_kelas tidak kosong)
if df_kelas.empty:
    st.warning("Tidak ada data siswa untuk kelas ini.")
else:
    if sel_siswa != "-- Semua Siswa --":
        row = df_kelas[df_kelas["Nama Siswa"] == sel_siswa].iloc[0]
        st.download_button("📄 Download PDF (Per Siswa)",
                           # CATATAN: sel_tgl_kegiatan ditambahkan di sini
                           data=make_pdf_for_student(row, mapel_urut, sel_tgl_ttd, asesmen_opsi, sel_tgl_kegiatan),
                           file_name=f"Laporan_{row['Nama Siswa']}.pdf",
                           mime="application/pdf")

    st.download_button("📄 Download PDF (Per Kelas)",
                       # CATATAN: sel_tgl_kegiatan ditambahkan di sini
                       data=make_pdf_for_class(df_kelas, mapel_urut, sel_tgl_ttd, asesmen_opsi, sel_tgl_kegiatan),
                       file_name=f"Laporan_{sel_kelas}.pdf",
                       mime="application/pdf")
