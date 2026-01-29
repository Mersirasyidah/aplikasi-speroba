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
# === KONFIGURASI MATA PELAJARAN [cite: 1] ===
# ===============================================
mapel_tetap = [
    "Bahasa Indonesia",
    "Matematika",
    "Bahasa Inggris",
    "Ilmu Pengetahuan Alam"
]
mapel_semua = mapel_tetap

# Mapping bulan Indonesia [cite: 2]
bulan_id = {
    "January": "Januari", "February": "Februari", "March": "Maret",
    "April": "April", "May": "Mei", "June": "Juni",
    "July": "Juli", "August": "Agustus", "September": "September",
    "October": "Oktober", "November": "November", "December": "Desember"
}

st.header("Laporan Hasil Persiapan dan Pemantapan")
st.markdown("---")

# --- Pilihan di Streamlit [cite: 2, 3] ---
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

tgl_kegiatan_opsi = [
    "Tanggal 20 - 23 Oktober 2025",
    "Tanggal 3 - 6 November 2025",
    "Tanggal 19 - 22 Januari 2026",
    "Tanggal 2 - 5 Februari 2026",
    "Tanggal 9 - 12 Maret 2026"
]

sel_tgl_kegiatan = st.selectbox("Pilih Tanggal Kegiatan", tgl_kegiatan_opsi)
sel_tgl_ttd = st.date_input("Tanggal Penulisan Tanda Tangan (di dokumen PDF)", datetime.now(), format="DD/MM/YYYY")
st.markdown("---")

# === Template Excel [cite: 4] ===
def generate_template():
    score_cols = []
    for i in range(1, 6):
        for m in mapel_semua:
            score_cols.append(f"{m}_TKAD{i}")

    # Tambah kolom 'Peringkat' untuk diisi manual atau dikosongkan (akan dihitung otomatis)
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
)

# Upload file Excel [cite: 5]
uploaded = st.file_uploader("Unggah file Excel daftar nilai (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Silakan unggah file Excel nilai")
    st.stop()

try:
    df = pd.read_excel(uploaded, engine="openpyxl")
except Exception as e:
    st.error(f"Gagal membaca file Excel: {e}")
    st.stop()

df.columns = df.columns.str.strip()

# Validasi Kolom [cite: 6]
expected_base = ["Kelas", "NIS", "Nama Siswa", "Peringkat"]
missing_base = [c for c in expected_base if c not in df.columns]
if missing_base:
    # Jika Peringkat tidak ada, tambahkan kolom kosong
    if "Peringkat" in missing_base:
        df["Peringkat"] = np.nan
        missing_base.remove("Peringkat")
    
    if missing_base:
        st.error(f"Kolom wajib hilang: {missing_base}")
        st.stop()

# Bersihkan dan Konversi Nilai [cite: 7, 8]
mapel_5_cols = [f"{m}_TKAD{i}" for i in range(1, 6) for m in mapel_tetap]
for col in [c for c in mapel_5_cols if c in df.columns]:
    df[col] = (
        df[col].astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
        .str.strip()
    )
    df.loc[df[col] == "", col] = np.nan
    df[col] = pd.to_numeric(df[col], errors="coerce")

# ===============================================
# === FITUR: PERHITUNGAN PERINGKAT OTOMATIS ===
# ===============================================
# Tentukan sesi asesmen mana yang sedang dipilih (1-5)
idx_asesmen = asesmen_opsi.index(sel_asesmen) + 1
current_cols = [f"{m}_TKAD{idx_asesmen}" for m in mapel_tetap]

# Hitung total nilai untuk sesi yang dipilih
df['Total_Skor_Temp'] = df[current_cols].sum(axis=1, skipna=True)
# Tandai siswa yang memiliki setidaknya satu nilai di sesi tersebut
df['Ada_Nilai'] = df[current_cols].notna().any(axis=1)

# Hitung peringkat per kelas berdasarkan total nilai
df['Peringkat_Auto'] = df.groupby('Kelas')['Total_Skor_Temp'].rank(ascending=False, method='min')

# Jika kolom 'Peringkat' di Excel kosong, gunakan hasil hitungan otomatis
df['Peringkat'] = df['Peringkat'].fillna(df['Peringkat_Auto'])
# Pastikan peringkat hanya muncul jika ada nilai
df.loc[~df['Ada_Nilai'], 'Peringkat'] = np.nan

# Format peringkat agar jadi angka bulat (1, 2, dst)
def format_rank(val):
    try: return str(int(float(val))) if not pd.isna(val) else "-"
    except: return str(val)

df['Peringkat'] = df['Peringkat'].apply(format_rank)

# Filter Kelas [cite: 6]
kelas_list = sorted(df["Kelas"].astype(str).unique())
sel_kelas = st.selectbox("Pilih Kelas", kelas_list)
df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

siswa_list = df_kelas["Nama Siswa"].astype(str).tolist()
sel_siswa = st.selectbox("Pilih Siswa", ["-- Semua Siswa --"] + siswa_list)

# ===============================================
# === PDF GENERATOR [cite: 10-17] ===
# ===============================================
def format_score(val):
    if pd.isna(val): return ""
    try: return f"{float(val):.2f}"
    except: return str(val)

def draw_student_page(c, row, sel_asesmen, sel_tahun, sel_tgl_kegiatan, mapel_urut, sel_tgl_ttd, asesmen_list):
    width, height = A4
    margin_left, margin_right, margin_top, margin_bottom = 30*mm, 15*mm, 25*mm, 20*mm
    content_width = width - (margin_left + margin_right)
    y = height - margin_top

    # Header & Kop [cite: 11-14]
    try:
        c.drawImage("assets/logo_kiri.png", margin_left - 10*mm, height - margin_top + 10*mm - 30*mm, width=30*mm, height=30*mm, preserveAspectRatio=True, mask='auto')
    except: pass

    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "PEMERINTAH KABUPATEN BANTUL")
    y -= 5*mm
    c.drawCentredString(width/2, y, "DINAS PENDIDIKAN, KEPEMUDAAN, DAN OLAHRAGA")
    y -= 5*mm
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width/2, y, "SMP NEGERI 2 BANGUNTAPAN")
    y -= 2*mm
    
    try:
        c.drawImage("assets/aksara_jawa.jpg", (width - 100*mm)/2, y - 10*mm, width=100*mm, height=10*mm, preserveAspectRatio=True, mask='auto')
        y -= 12*mm
    except: y -= 4*mm

    c.setFont("Helvetica-Oblique", 10)
    c.drawCentredString(width/2, y, "Jalan Karangsari, Banguntapan, Kabupaten Bantul, Yogyakarta 55198 Telp. 382754")
    y -= 5*mm
    c.setFont("Helvetica", 10)
    c.setFillColor(blue)
    c.drawCentredString(width/2, y, "Website : www.smpn2banguntapan.sch.id Email : smp2banguntapan@yahoo.com")
    c.setFillColor(black)
    y -= 3*mm

    c.setLineWidth(1)
    c.line(margin_left, y, width - margin_right, y)
    y -= 10*mm

    # Judul Dokumen [cite: 15]
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "LAPORAN HASIL PERSIAPAN DAN PEMANTAPAN")
    y -= 6*mm
    c.drawCentredString(width/2, y, sel_asesmen.upper())
    y -= 6*mm
    c.drawCentredString(width/2, y, f"TAHUN PELAJARAN {sel_tahun}")
    y -= 6*mm
    c.drawCentredString(width/2, y, sel_tgl_kegiatan)
    y -= 12*mm

    # Identitas Siswa & Peringkat 
    id_x, label_w = margin_left + 10*mm, 25*mm
    colon_x, value_x = id_x + label_w, id_x + label_w + 5
    
    for label, key in [("Nama", "Nama Siswa"), ("NIS", "NIS"), ("Kelas", "Kelas"), ("Peringkat", "Peringkat")]:
        c.drawString(id_x, y, label)
        c.drawString(colon_x, y, ":")
        c.drawString(value_x, y, f" {row.get(key, '')}")
        y -= 6*mm
    
    y -= 4*mm

    # --- Pengolahan & Tabel Nilai [cite: 18-32] ---
    scores_by_col = [[] for _ in range(5)]
    nilai_data = {}
    for subj in mapel_urut:
        row_scores = []
        for i in range(1, 6):
            val = row.get(f"{subj}_TKAD{i}", np.nan)
            row_scores.append(val)
            if not np.isnan(val): scores_by_col[i-1].append(val)
        nilai_data[subj] = row_scores

    # Header Tabel
    row_h, col_no, col_mapel, col_score = 7*mm, 15*mm, 70*mm, 15*mm
    table_w = col_no + col_mapel + (col_score * 5)
    x_start = margin_left + (content_width - table_w)/2
    
    c.setFillColor(lightgrey)
    c.rect(x_start, y - 2*row_h, table_w, 2*row_h, stroke=0, fill=1)
    c.setFillColor(black)
    c.setFont("Helvetica-Bold", 11)
    
    # Grid & Header Text
    c.drawCentredString(x_start + col_no/2, y - row_h - 2*mm, "No")
    c.drawCentredString(x_start + col_no + col_mapel/2, y - row_h - 2*mm, "Mata Pelajaran")
    c.drawCentredString(x_start + col_no + col_mapel + (col_score*5)/2, y - row_h/2 - 2*mm, "Nilai TKA/TKAD")
    for i in range(5):
        c.drawCentredString(x_start + col_no + col_mapel + (i*col_score) + col_score/2, y - 1.5*row_h - 2*mm, str(i+1))

    # Isi Nilai
    y_row = y - 2*row_h
    c.setFont("Helvetica", 11)
    for i, subj in enumerate(mapel_urut, start=1):
        c.drawCentredString(x_start + col_no/2, y_row - row_h/2 - 2*mm, str(i))
        c.drawString(x_start + col_no + 2*mm, y_row - row_h/2 - 2*mm, subj)
        for j, score in enumerate(nilai_data[subj]):
            c.drawCentredString(x_start + col_no + col_mapel + (j*col_score) + col_score/2, y_row - row_h/2 - 2*mm, format_score(score))
        y_row -= row_h

    # Baris Jumlah & Rata-rata [cite: 30-32]
    c.setFont("Helvetica-Bold", 11)
    for label, data in [("Jumlah", [sum(x) if x else np.nan for x in scores_by_col]), 
                        ("Rata-rata", [sum(x)/len(x) if x else np.nan for x in scores_by_col])]:
        c.drawString(x_start + col_no + 2*mm, y_row - row_h/2 - 2*mm, label)
        for j, val in enumerate(data):
            c.drawCentredString(x_start + col_no + col_mapel + (j*col_score) + col_score/2, y_row - row_h/2 - 2*mm, format_score(val))
        y_row -= row_h

    # Grid Lines
    c.setLineWidth(0.5)
    for i in range(len(mapel_urut) + 5):
        if i == 1: continue
        c.line(x_start, y - i*row_h, x_start + table_w, y - i*row_h)
    c.line(x_start, y, x_start, y_row)
    c.line(x_start + col_no, y, x_start + col_no, y_row)
    c.line(x_start + col_no + col_mapel, y, x_start + col_no + col_mapel, y_row)
    for i in range(6):
        c.line(x_start + col_no + col_mapel + i*col_score, y if i==0 or i==5 else y - row_h, x_start + col_no + col_mapel + i*col_score, y_row)

    # Keterangan & Tanda Tangan [cite: 33-36]
    y_ttd = y_row - 10*mm
    c.setFont("Helvetica-Bold", 9)
    c.drawString(margin_left, y_ttd, "Keterangan:")
    for i, item in enumerate(asesmen_list, start=1):
        y_ttd -= 4*mm
        tgl_item = tgl_kegiatan_opsi[i-1] if i-1 < len(tgl_kegiatan_opsi) else "-"
        c.setFont("Helvetica", 9)
        c.drawString(margin_left + 5*mm, y_ttd, f"{i}. {item} ({tgl_item})")

    ttd_date = f"{sel_tgl_ttd.day} {bulan_id[sel_tgl_ttd.strftime('%B')]} {sel_tgl_ttd.year}"
    x_ttd = width - margin_right - 60*mm
    y_sign = margin_bottom + 45*mm
    c.setFont("Helvetica", 11)
    c.drawString(x_ttd, y_sign + 15*mm, f"Banguntapan, {ttd_date}")
    c.drawString(x_ttd, y_sign + 10*mm, "Mengetahui,")
    c.drawString(x_ttd, y_sign + 5*mm, "Kepala Sekolah,")
    
    try:
        c.drawImage("assets/ttd_kepsek.jpeg", x_ttd, y_sign - 15*mm, width=40*mm, height=20*mm, mask='auto')
    except: pass
    
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_ttd, y_sign - 20*mm, "Alina Fiftiyani Nurjannah, M.Pd.")
    c.setFont("Helvetica", 11)
    c.drawString(x_ttd, y_sign - 25*mm, "NIP 198001052009032006")

# Fungsi Helper PDF [cite: 37-40]
def make_pdf(data_rows):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    for _, row in data_rows.iterrows():
        draw_student_page(c, row, sel_asesmen, sel_tahun, sel_tgl_kegiatan, mapel_urut, sel_tgl_ttd, asesmen_opsi)
        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

if not df_kelas.empty:
    if sel_siswa != "-- Semua Siswa --":
        row_sel = df_kelas[df_kelas["Nama Siswa"] == sel_siswa]
        st.download_button("📄 Download PDF (Siswa Terpilih)", data=make_pdf(row_sel), file_name=f"Laporan_{sel_siswa}.pdf", mime="application/pdf")
    
    st.download_button("📄 Download PDF (Satu Kelas)", data=make_pdf(df_kelas), file_name=f"Laporan_Kelas_{sel_kelas}.pdf", mime="application/pdf")
