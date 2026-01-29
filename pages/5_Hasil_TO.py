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

# ===============================================
# === MODIFIKASI: Mapel Tetap (4 Mata Pelajaran) ===
# ===============================================
mapel_tetap = [
    "Bahasa Indonesia",
    "Matematika",
    "Bahasa Inggris",
    "Ilmu Pengetahuan Alam"
[cite_start]] [cite: 1]
[cite_start]mapel_semua = mapel_tetap [cite: 1]

# Mapping bulan Indonesia
bulan_id = {
    "January": "Januari", "February": "Februari", "March": "Maret",
    "April": "April", "May": "Mei", "June": "Juni",
    "July": "Juli", "August": "Agustus", "September": "September",
    "October": "Oktober", "November": "November", "December": "Desember"
[cite_start]} [cite: 1, 2]

st.header("Laporan Hasil Persiapan dan Pemantapan")
st.markdown("---")

# --- Pilihan di Streamlit ---
asesmen_opsi = [
    "TKA/TKAD Dikpora Kab Bantul 1",
    "TKA/TKAD MKKS SMP Kab Bantul 1",
    "TKA/TKAD MKKS SMP Kab Bantul 2",
    "TKA/TKAD Forum MKKS SMP D.I.Yogyakarta",
    "TKA/TKAD Dikpora Kab Bantul 2"
[cite_start]] [cite: 2]
[cite_start]sel_asesmen = st.selectbox("Pilih Jenis Asesmen", asesmen_opsi) [cite: 2]
[cite_start]tahun_opsi = [f"{th}/{th+1}" for th in range(2025, 2036)] [cite: 2]
[cite_start]sel_tahun = st.selectbox("Pilih Tahun Pelajaran", tahun_opsi, index=0) [cite: 2]

tgl_kegiatan_opsi = [
    "Tanggal 20 - 23 Oktober 2025",
    "Tanggal 3 - 6 November 2025",
    "Tanggal 19 - 22 Januari 2026",
    "Tanggal 2 - 5 Februari 2026",
    "Tanggal 9 - 12 Maret 2026"
[cite_start]] [cite: 3]

[cite_start]sel_tgl_kegiatan = st.selectbox("Pilih Tanggal Kegiatan", tgl_kegiatan_opsi) [cite: 3]
[cite_start]sel_tgl_ttd = st.date_input("Tanggal Penulisan Tanda Tangan (di dokumen PDF)", datetime.now(), format="DD/MM/YYYY") [cite: 3]
st.markdown("---")

# === Template Excel ===
def generate_template():
    score_cols = []
    for i in range(1, 6):
        for m in mapel_semua:
            [cite_start]score_cols.append(f"{m}_TKAD{i}") [cite: 3, 4]

    # [cite_start]MODIFIKASI: Menambahkan kolom 'Peringkat' [cite: 4]
    cols = ["Kelas", "NIS", "Nama Siswa", "Peringkat"] + score_cols
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
[cite_start]) [cite: 4]

# Upload file Excel
[cite_start]uploaded = st.file_uploader("Unggah file Excel daftar nilai (.xlsx)", type=["xlsx"]) [cite: 4, 5]
if not uploaded:
    st.info("Silakan unggah file Excel nilai")
    st.stop()

try:
    df = pd.read_excel(uploaded, engine="openpyxl")
except Exception as e:
    st.error(f"Gagal membaca file Excel: {e}")
    [cite_start]st.stop() [cite: 5]

df.columns = df.columns.str.strip()

if "Kelas" not in df.columns:
    st.error("Kolom 'Kelas' tidak ditemukan di file. Pastikan pakai template.")
    [cite_start]st.stop() [cite: 5, 6]

[cite_start]kelas_list = sorted(df["Kelas"].astype(str).unique()) [cite: 6]
[cite_start]sel_kelas = st.selectbox("Pilih Kelas", kelas_list) [cite: 6]
[cite_start]df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy() [cite: 6]

# ===============================================
mapel_urut = mapel_tetap
mapel_5_cols = []
for i in range(1, 6):
    for m in mapel_urut:
        [cite_start]mapel_5_cols.append(f"{m}_TKAD{i}") [cite: 6]

# [cite_start]MODIFIKASI: Menambahkan 'Peringkat' ke dalam expected_base [cite: 6]
expected_base = ["Kelas", "NIS", "Nama Siswa", "Peringkat"]
missing_base = [c for c in expected_base if c not in df.columns]
if missing_base:
    st.error(f"Kolom wajib hilang: {missing_base}")
    [cite_start]st.stop() [cite: 6]

missing_mapel_cols = [c for c in mapel_5_cols if c not in df.columns]
if missing_mapel_cols:
    st.warning(f"Kolom nilai wajib hilang di file Excel dan akan diperlakukan kosong: {missing_mapel_cols[:5]}...")
    for m in missing_mapel_cols:
        [cite_start]df[m] = np.nan [cite: 6, 7]
    [cite_start]df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy() [cite: 7]

for col in [c for c in mapel_5_cols if c in df.columns]:
    df[col] = (
        df[col]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
        .str.strip()
    )
    df.loc[df[col] == "", col] = np.nan
    [cite_start]df[col] = pd.to_numeric(df[col], errors="coerce") [cite: 7, 8]

cols_to_keep = expected_base + [c for c in mapel_5_cols if c in df.columns]
df = df[cols_to_keep]
[cite_start]df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy() [cite: 8]

siswa_list = df_kelas["Nama Siswa"].astype(str).tolist()
[cite_start]sel_siswa = st.selectbox("Pilih Siswa", ["-- Semua Siswa --"] + siswa_list) [cite: 8]

def format_score(val):
    if pd.isna(val):
        return ""
    try:
        num = float(val)
        return f"{num:.2f}"
    except Exception:
        try:
            num = float(str(val).strip().replace(",", "."))
            return f"{num:.2f}"
        except Exception:
            [cite_start]return str(val) [cite: 8, 9]

def draw_student_page(c, row, sel_asesmen, sel_tahun, sel_tgl_kegiatan, mapel_urut, sel_tgl_ttd, asesmen_list):
    [cite_start]width, height = A4 [cite: 9, 10]
    aksara_path = "assets/aksara_jawa.jpg"
    logo_path = "assets/logo_kiri.png"

    margin_left   = 30 * mm
    margin_right  = 15 * mm
    margin_top    = 25 * mm
    margin_bottom = 20 * mm

    content_width  = width - (margin_left + margin_right)
    [cite_start]y = height - margin_top [cite: 10]

    try:
        logo_w, logo_h = 30*mm, 30*mm
        x_logo, y_logo = margin_left - 10*mm, height - margin_top - (-10*mm) - logo_h
        c.drawImage(logo_path, x_logo, y_logo, width=logo_w, height=logo_h, preserveAspectRatio=True, mask='auto')
    [cite_start]except: pass [cite: 10, 11]

    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "PEMERINTAH KABUPATEN BANTUL")
    y -= 5*mm
    c.drawCentredString(width/2, y, "DINAS PENDIDIKAN, KEPEMUDAAN, DAN OLAHRAGA")
    y -= 5*mm
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width/2, y, "SMP NEGERI 2 BANGUNTAPAN")
    [cite_start]y -= 1*mm [cite: 11, 12]

    try:
        aksara_w, aksara_h = 100*mm, 10*mm
        x_aksara = (width - aksara_w) / 2
        y_aksara = y - aksara_h
        c.drawImage(aksara_path, x_aksara, y_aksara, width=aksara_w, height=aksara_h, preserveAspectRatio=True, mask='auto')
        y = y_aksara - 2*mm
    [cite_start]except: y -= 4*mm [cite: 12, 13]

    c.setFont("Helvetica-Oblique", 10)
    c.drawCentredString(width/2, y, "Jalan Karangsari, Banguntapan, Kabupaten Bantul, Yogyakarta 55198 Telp. 382754")
    y -= 5*mm
    c.setFont("Helvetica", 10)
    c.setFillColor(blue)
    c.drawCentredString(width/2, y, "Website : www.smpn2banguntapan.sch.id Email : smp2banguntapan@yahoo.com")
    c.setFillColor(black)
    [cite_start]y -= 3*mm [cite: 13, 14]

    c.setLineWidth(1)
    c.line(margin_left, y, width - margin_right, y)
    y -= 1.5*mm
    c.setLineWidth(0.5)
    c.line(margin_left, y, width - margin_right, y)
    [cite_start]y -= 10*mm [cite: 14]

    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "LAPORAN HASIL PERSIAPAN DAN PEMANTAPAN")
    y -= 6*mm
    c.drawCentredString(width/2, y, sel_asesmen.upper())
    y -= 6*mm
    c.drawCentredString(width/2, y, f"TAHUN PELAJARAN {sel_tahun}")
    y -= 6*mm
    c.drawCentredString(width/2, y, sel_tgl_kegiatan)
    [cite_start]y -= 12*mm [cite: 15, 16]

    # Identitas siswa
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
    y -= 6*mm

    # [cite_start]MODIFIKASI: Menambahkan tampilan Peringkat [cite: 16, 17]
    c.drawString(id_margin_left, y, "Peringkat")
    c.drawString(colon_x, y, ":")
    c.drawString(value_x, y, " " + str(row.get("Peringkat", "")))
    y -= 10*mm

    # --- Pengolahan Nilai ---
    nilai_cols_count = 5
    nilai_data = {}
    [cite_start]scores_by_column = [[] for _ in range(nilai_cols_count)] [cite: 17, 18]

    for subj in mapel_urut:
        scores = []
        for i in range(1, nilai_cols_count + 1):
            col_name = f"{subj}_TKAD{i}"
            raw = row.get(col_name, np.nan)
            val = np.nan
            try:
                val = float(raw)
            except: pass
            scores.append(val)
            if not np.isnan(val):
                [cite_start]scores_by_column[i-1].append(val) [cite: 18, 19, 20]
        nilai_data[subj] = scores

    jumlah_per_kolom, rata2_per_kolom = [], []
    for col_scores in scores_by_column:
        if len(col_scores) == 0:
            col_sum, col_avg = np.nan, np.nan
        else:
            col_sum = sum(col_scores)
            [cite_start]col_avg = col_sum / len(col_scores) [cite: 21]
        jumlah_per_kolom.append(col_sum)
        [cite_start]rata2_per_kolom.append(col_avg) [cite: 21]

    # --- Tabel ---
    row_height, font_size = 7*mm, 11
    col_no_w, col_mapel_w, col_nilai_single_w = 15*mm, 70*mm, 15*mm
    table_width = col_no_w + col_mapel_w + (col_nilai_single_w * nilai_cols_count)
    x0, y0 = margin_left + (content_width - table_width) / 2, y
    total_rows_to_draw = len(mapel_urut) + 4

    c.setFillColor(lightgrey)
    c.rect(x0, y0 - 2*row_height, table_width, 2*row_height, stroke=0, fill=1)
    [cite_start]c.setFillColor(black) [cite: 22, 23]

    c.setLineWidth(0.5)
    # Garis vertikal
    c.line(x0, y0, x0, y0 - total_rows_to_draw*row_height)
    c.line(x0 + col_no_w, y0, x0 + col_no_w, y0 - total_rows_to_draw*row_height)
    c.line(x0 + col_no_w + col_mapel_w, y0, x0 + col_no_w + col_mapel_w, y0 - total_rows_to_draw*row_height)
    
    current_x_grid = x0 + col_no_w + col_mapel_w
    for i in range(1, nilai_cols_count):
        current_x_grid += col_nilai_single_w
        [cite_start]c.line(current_x_grid, y0 - row_height, current_x_grid, y0 - total_rows_to_draw*row_height) [cite: 23, 24]
    
    c.line(x0 + table_width, y0, x0 + table_width, y0 - total_rows_to_draw*row_height)

    # Garis horizontal
    for r in range(total_rows_to_draw + 1):
        if r == 1: continue
        [cite_start]c.line(x0, y0 - r*row_height, x0 + table_width, y0 - r*row_height) [cite: 24, 25]
    [cite_start]c.line(x0 + col_no_w + col_mapel_w, y0 - row_height, x0 + table_width, y0 - row_height) [cite: 25]

    # Header Text
    adj_y_merged = y0 - row_height - (font_size/3.5)
    c.setFont("Helvetica-Bold", font_size)
    c.drawCentredString(x0 + col_no_w/2, adj_y_merged, "No")
    c.drawCentredString(x0 + col_no_w + col_mapel_w/2, adj_y_merged, "Mata Pelajaran")
    c.drawCentredString(x0 + col_no_w + col_mapel_w + (col_nilai_single_w * nilai_cols_count)/2, y0 - row_height/2 - (font_size/3.5), "Nilai TKA/TKAD")

    current_x = x0 + col_no_w + col_mapel_w
    for i in range(nilai_cols_count):
        c.drawCentredString(current_x + col_nilai_single_w/2, y0 - 1.5*row_height - (font_size/3.5), str(i+1))
        [cite_start]current_x += col_nilai_single_w [cite: 26, 27]

    # Isi Tabel
    y_text = y0 - 2 * row_height
    c.setFont("Helvetica", font_size)
    for i, subj in enumerate(mapel_urut, start=1):
        adj_y = y_text - row_height/2 - (font_size/3.5)
        c.drawCentredString(x0 + col_no_w/2, adj_y, str(i))
        c.drawString(x0 + col_no_w + 2*mm, adj_y, subj)
        
        scores = nilai_data.get(subj, [np.nan] * nilai_cols_count)
        current_x = x0 + col_no_w + col_mapel_w
        for score in scores:
            c.drawCentredString(current_x + col_nilai_single_w/2, adj_y, format_score(score))
            [cite_start]current_x += col_nilai_single_w [cite: 28, 29]
        y_text -= row_height

    # Baris Jumlah & Rata-rata
    c.setFont("Helvetica-Bold", font_size)
    for label, data_list in [("Jumlah", jumlah_per_kolom), ("Rata-rata", rata2_per_kolom)]:
        adj_y = y_text - row_height/2 - (font_size/3.5)
        c.drawString(x0 + col_no_w + 2*mm, adj_y, label)
        current_x = x0 + col_no_w + col_mapel_w
        for val in data_list:
            c.drawCentredString(current_x + col_nilai_single_w/2, adj_y, format_score(val) if not pd.isna(val) else "")
            [cite_start]current_x += col_nilai_single_w [cite: 30, 31, 32]
        y_text -= row_height

    # Keterangan & TTD
    y_keterangan = y_text - 5*mm
    c.setFont("Helvetica-Bold", 9)
    c.drawString(margin_left, y_keterangan, "Keterangan:")
    y_keterangan -= 5*mm
    c.setFont("Helvetica", 9)
    for i, item in enumerate(asesmen_list, start=1):
        tgl_item = tgl_kegiatan_opsi[i-1] if i-1 < len(tgl_kegiatan_opsi) else "TANGGAL TIDAK DIKETAHUI"
        c.drawString(margin_left + 5*mm, y_keterangan, f"{i}. {item} ({tgl_item})")
        [cite_start]y_keterangan -= 5*mm [cite: 33, 34]

    ttd_date = sel_tgl_ttd
    tgl_ttd_str = f"{ttd_date.day} {bulan_id[ttd_date.strftime('%B')]} {ttd_date.year}"
    x_ttd, y_ttd = width - margin_right - 70*mm, margin_bottom + 60*mm
    c.setFont("Helvetica", 12)
    c.drawString(x_ttd, y_ttd, f"Banguntapan, {tgl_ttd_str}")
    y_ttd -= 8*mm
    c.drawString(x_ttd, y_ttd, "Mengetahui,")
    y_ttd -= 5*mm
    c.drawString(x_ttd, y_ttd, "Kepala Sekolah,")
    
    if os.path.exists("assets/ttd_kepsek.jpeg"):
        try: c.drawImage("assets/ttd_kepsek.jpeg", x_ttd, y_ttd - 22*mm, width=40*mm, height=20*mm, mask="auto")
        [cite_start]except: pass [cite: 34, 35]

    y_ttd -= 25*mm
    c.drawString(x_ttd, y_ttd, "Alina Fiftiyani Nurjannah, M.Pd.")
    [cite_start]c.drawString(x_ttd, y_ttd - 6*mm, "NIP 198001052009032006") [cite: 36]

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
        [cite_start]c.showPage() [cite: 36, 37]
    c.save()
    buffer.seek(0)
    return buffer

st.markdown("---")
st.subheader("Pilih Siswa & Unduh Laporan")

if df_kelas.empty:
    st.warning("Tidak ada data siswa untuk kelas ini.")
else:
    if sel_siswa != "-- Semua Siswa --":
        row = df_kelas[df_kelas["Nama Siswa"] == sel_siswa].iloc[0]
        st.download_button("📄 Download PDF (Per Siswa)",
                           data=make_pdf_for_student(row, mapel_urut, sel_tgl_ttd, asesmen_opsi, sel_tgl_kegiatan),
                           file_name=f"Laporan_{row['Nama Siswa']}.pdf",
                           [cite_start]mime="application/pdf") [cite: 38, 39]

    st.download_button("📄 Download PDF (Per Kelas)",
                       data=make_pdf_for_class(df_kelas, mapel_urut, sel_tgl_ttd, asesmen_opsi, sel_tgl_kegiatan),
                       file_name=f"Laporan_{sel_kelas}.pdf",
                       [cite_start]mime="application/pdf") [cite: 39, 40]
