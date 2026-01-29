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
# === KONFIGURASI MATA PELAJARAN ===
# ===============================================
mapel_tetap = [
    "Bahasa Indonesia",
    "Matematika",
    "Bahasa Inggris",
    "Ilmu Pengetahuan Alam"
]
mapel_semua = mapel_tetap
mapel_urut = mapel_tetap  # Mendefinisikan mapel_urut agar tidak error

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

# === Template Excel ===
def generate_template():
    score_cols = []
    for i in range(1, 6):
        for m in mapel_semua:
            score_cols.append(f"{m}_TKAD{i}")

    # Tambah kolom 'Peringkat'
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

# Upload file Excel
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

# Validasi Kolom Wajib
if "Kelas" not in df.columns or "Nama Siswa" not in df.columns:
    st.error("Kolom 'Kelas' atau 'Nama Siswa' tidak ditemukan.")
    st.stop()

# Tambahkan kolom Peringkat jika tidak ada di file
if "Peringkat" not in df.columns:
    df["Peringkat"] = np.nan

# Bersihkan Data Nilai
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
idx_asesmen = asesmen_opsi.index(sel_asesmen) + 1
current_cols = [f"{m}_TKAD{idx_asesmen}" for m in mapel_tetap]
available_cols = [c for c in current_cols if c in df.columns]

if available_cols:
    df['Total_Skor_Temp'] = df[available_cols].sum(axis=1, skipna=True)
    df['Ada_Nilai'] = df[available_cols].notna().any(axis=1)
    
    # Hitung peringkat per kelas berdasarkan total nilai
    df['Peringkat_Auto'] = df.groupby('Kelas')['Total_Skor_Temp'].rank(ascending=False, method='min')
    
    # Gunakan peringkat auto jika kolom Peringkat kosong
    df['Peringkat'] = df['Peringkat'].fillna(df['Peringkat_Auto'])
    df.loc[~df['Ada_Nilai'], 'Peringkat'] = np.nan
else:
    df['Peringkat'] = "-"

# Format Peringkat menjadi string bersih
def clean_rank(val):
    if pd.isna(val) or val == "": return "-"
    try: return str(int(float(val)))
    except: return str(val)

df['Peringkat'] = df['Peringkat'].apply(clean_rank)

# Filter Kelas
kelas_list = sorted(df["Kelas"].astype(str).unique())
sel_kelas = st.selectbox("Pilih Kelas", kelas_list)
df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

siswa_list = df_kelas["Nama Siswa"].astype(str).tolist()
sel_siswa = st.selectbox("Pilih Siswa", ["-- Semua Siswa --"] + siswa_list)

# ===============================================
# === PDF GENERATOR ===
# ===============================================
def format_score(val):
    if pd.isna(val): return ""
    try: return f"{float(val):.2f}"
    except: return str(val)

def draw_student_page(c, row, sel_asesmen, sel_tahun, sel_tgl_kegiatan, mapel_list, sel_tgl_ttd, asesmen_list):
    width, height = A4
    margin_left, margin_right, margin_top, margin_bottom = 30*mm, 15*mm, 25*mm, 20*mm
    content_width = width - (margin_left + margin_right)
    y = height - margin_top

    # Header / Kop Surat
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

    # Judul
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "LAPORAN HASIL PERSIAPAN DAN PEMANTAPAN")
    y -= 6*mm
    c.drawCentredString(width/2, y, sel_asesmen.upper())
    y -= 6*mm
    c.drawCentredString(width/2, y, f"TAHUN PELAJARAN {sel_tahun}")
    y -= 6*mm
    c.drawCentredString(width/2, y, sel_tgl_kegiatan)
    y -= 12*mm

    # Identitas & Peringkat
    id_x, label_w = margin_left + 10*mm, 25*mm
    colon_x, val_x = id_x + label_w, id_x + label_w + 5
    
    identitas = [
        ("Nama", row.get("Nama Siswa", "")),
        ("NIS", row.get("NIS", "")),
        ("Kelas", row.get("Kelas", "")),
        ("Peringkat", row.get("Peringkat", "-"))
    ]

    for label, val in identitas:
        c.drawString(id_x, y, label)
        c.drawString(colon_x, y, ":")
        c.drawString(val_x, y, f" {val}")
        y -= 6*mm
    
    y -= 4*mm

    # Tabel Nilai
    row_h, col_no, col_mapel, col_score = 7*mm, 15*mm, 70*mm, 15*mm
    table_w = col_no + col_mapel + (col_score * 5)
    x_s = margin_left + (content_width - table_w)/2
    
    c.setFillColor(lightgrey)
    c.rect(x_s, y - 2*row_h, table_w, 2*row_h, stroke=0, fill=1)
    c.setFillColor(black)
    c.setFont("Helvetica-Bold", 11)
    
    c.drawCentredString(x_s + col_no/2, y - row_h - 2*mm, "No")
    c.drawCentredString(x_s + col_no + col_mapel/2, y - row_h - 2*mm, "Mata Pelajaran")
    c.drawCentredString(x_s + col_no + col_mapel + (col_score*5)/2, y - row_h/2 - 2*mm, "Nilai TKA/TKAD")
    for i in range(5):
        c.drawCentredString(x_s + col_no + col_mapel + (i*col_score) + col_score/2, y - 1.5*row_h - 2*mm, str(i+1))

    # Data Baris
    y_row = y - 2*row_h
    c.setFont("Helvetica", 11)
    
    sums = [[] for _ in range(5)]
    for i, m in enumerate(mapel_list, 1):
        c.drawCentredString(x_s + col_no/2, y_row - row_h/2 - 2*mm, str(i))
        c.drawString(x_s + col_no + 2*mm, y_row - row_h/2 - 2*mm, m)
        for j in range(5):
            val = row.get(f"{m}_TKAD{j+1}", np.nan)
            c.drawCentredString(x_s + col_no + col_mapel + (j*col_score) + col_score/2, y_row - row_h/2 - 2*mm, format_score(val))
            if not pd.isna(val): sums[j].append(val)
        y_row -= row_h

    # Baris Jumlah & Rata-rata
    c.setFont("Helvetica-Bold", 11)
    for label, data in [("Jumlah", [sum(x) if x else np.nan for x in sums]), 
                        ("Rata-rata", [sum(x)/len(x) if x else np.nan for x in sums])]:
        c.drawString(x_s + col_no + 2*mm, y_row - row_h/2 - 2*mm, label)
        for j, val in enumerate(data):
            c.drawCentredString(x_s + col_no + col_mapel + (j*col_score) + col_score/2, y_row - row_h/2 - 2*mm, format_score(val))
        y_row -= row_h

    # Garis-garis Tabel
    c.setLineWidth(0.5)
    total_rows = len(mapel_list) + 2 + 2 # Header + Data + Jml/Rata
    for i in range(total_rows + 1):
        if i == 1: continue
        c.line(x_s, y - i*row_h, x_s + table_w, y - i*row_h)
    c.line(x_s, y, x_s, y_row)
    c.line(x_s + col_no, y, x_s + col_no, y_row)
    c.line(x_s + col_no + col_mapel, y, x_s + col_no + col_mapel, y_row)
    for i in range(6):
        c.line(x_s + col_no + col_mapel + i*col_score, y if i==0 or i==5 else y - row_h, x_s + col_no + col_mapel + i*col_score, y_row)

    # Tanda Tangan
    y_ttd = y_row - 10*mm
    ttd_date = f"{sel_tgl_ttd.day} {bulan_id[sel_tgl_ttd.strftime('%B')]} {sel_tgl_ttd.year}"
    x_t = width - margin_right - 60*mm
    c.setFont("Helvetica", 11)
    c.drawString(x_t, y_ttd - 10*mm, f"Banguntapan, {ttd_date}")
    c.drawString(x_t, y_ttd - 15*mm, "Kepala Sekolah,")
    
    try:
        c.drawImage("assets/ttd_kepsek.jpeg", x_t, y_ttd - 40*mm, width=40*mm, height=20*mm, mask='auto')
    except: pass
    
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_t, y_ttd - 45*mm, "Alina Fiftiyani Nurjannah, M.Pd.")
    c.setFont("Helvetica", 11)
    c.drawString(x_t, y_ttd - 50*mm, "NIP 198001052009032006")

def make_pdf(data_rows):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    for _, row in data_rows.iterrows():
        draw_student_page(c, row, sel_asesmen, sel_tahun, sel_tgl_kegiatan, mapel_urut, sel_tgl_ttd, asesmen_opsi)
        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# Tombol Unduh
if not df_kelas.empty:
    if sel_siswa != "-- Semua Siswa --":
        row_sel = df_kelas[df_kelas["Nama Siswa"] == sel_siswa]
        st.download_button("📄 Download PDF (Siswa Terpilih)", data=make_pdf(row_sel), file_name=f"Laporan_{sel_siswa}.pdf", mime="application/pdf")
    
    st.download_button("📄 Download PDF (Satu Kelas)", data=make_pdf(df_kelas), file_name=f"Laporan_Kelas_{sel_kelas}.pdf", mime="application/pdf")
