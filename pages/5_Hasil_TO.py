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
mapel_urut = mapel_tetap

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
sel_tgl_ttd = st.date_input("Tanggal Penulisan Tanda Tangan", datetime.now(), format="DD/MM/YYYY")

# === Template Excel ===
def generate_template():
    score_cols = [f"{m}_TKAD{i}" for i in range(1, 6) for m in mapel_tetap]
    cols = ["Kelas", "NIS", "Nama Siswa"] + score_cols
    df_template = pd.DataFrame(columns=cols)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False, sheet_name="Nilai")
    buffer.seek(0)
    return buffer

st.download_button("📥 Download Template Excel", data=generate_template(), file_name="Template_Nilai_TO.xlsx")

# Upload file Excel
uploaded = st.file_uploader("Unggah file Excel daftar nilai (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Silakan unggah file Excel nilai")
    st.stop()

try:
    df = pd.read_excel(uploaded, engine="openpyxl").fillna(np.nan)
    df.columns = df.columns.str.strip()
except Exception as e:
    st.error(f"Gagal membaca file Excel: {e}")
    st.stop()

# ===============================================
# === LOGIKA PERHITUNGAN PERINGKAT (RANKING) ===
# ===============================================
# Hitung peringkat untuk tiap sesi TKAD 1-5 secara otomatis
for i in range(1, 6):
    cols_sesi = [f"{m}_TKAD{i}" for m in mapel_tetap]
    available = [c for c in cols_sesi if c in df.columns]
    if available:
        # Hitung total skor per sesi
        df[f"Total_TKAD{i}"] = df[available].sum(axis=1, skipna=True)
        # Beri peringkat jika ada nilai
        mask = df[available].notna().any(axis=1)
        df.loc[mask, f"Peringkat_TKAD{i}"] = df[mask].groupby("Kelas")[f"Total_TKAD{i}"].rank(ascending=False, method="min")
    else:
        df[f"Peringkat_TKAD{i}"] = np.nan

# Filter Kelas & Siswa
kelas_list = sorted(df["Kelas"].astype(str).unique())
sel_kelas = st.selectbox("Pilih Kelas", kelas_list)
df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()
siswa_list = df_kelas["Nama Siswa"].astype(str).tolist()
sel_siswa = st.selectbox("Pilih Siswa", ["-- Semua Siswa --"] + siswa_list)

# helper: format skor
def format_score(val):
    if pd.isna(val) or val == "": return ""
    try:
        num = float(val)
        return f"{num:.2f}" if num % 1 != 0 else str(int(num))
    except: return str(val)

# === PDF GENERATOR ===
def draw_student_page(c, row, sel_asesmen, sel_tahun, sel_tgl_kegiatan, mapel_urut, sel_tgl_ttd, asesmen_list):
    width, height = A4
    ml, mr, mt, mb = 30*mm, 15*mm, 25*mm, 20*mm
    content_w = width - (ml + mr)
    y = height - mt

    # Kop Surat (Disederhanakan)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "PEMERINTAH KABUPATEN BANTUL")
    y -= 5*mm
    c.drawCentredString(width/2, y, "DINAS PENDIDIKAN, KEPEMUDAAN, DAN OLAHRAGA")
    y -= 5*mm
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width/2, y, "SMP NEGERI 2 BANGUNTAPAN")
    y -= 5*mm
    c.setFont("Helvetica", 10)
    c.drawCentredString(width/2, y, "Jalan Karangsari, Banguntapan, Bantul, Yogyakarta 55198")
    y -= 3*mm
    c.setLineWidth(1); c.line(ml, y, width - mr, y); y -= 8*mm

    # Judul
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "LAPORAN HASIL PERSIAPAN DAN PEMANTAPAN")
    y -= 6*mm
    c.drawCentredString(width/2, y, sel_asesmen.upper())
    y -= 6*mm
    c.drawCentredString(width/2, y, f"TAHUN PELAJARAN {sel_tahun}")
    y -= 6*mm
    c.drawCentredString(width/2, y, sel_tgl_kegiatan)
    y -= 10*mm

    # Identitas
    c.setFont("Helvetica-Bold", 11)
    identitas = [("Nama", row.get("Nama Siswa", "")), ("NIS", row.get("NIS", "")), ("Kelas", row.get("Kelas", ""))]
    for lbl, val in identitas:
        c.drawString(ml + 10*mm, y, lbl)
        c.drawString(ml + 30*mm, y, f": {val}")
        y -= 6*mm
    y -= 4*mm

    # --- Tabel Nilai ---
    row_h = 7*mm
    col_no_w, col_m_w, col_s_w = 12*mm, 65*mm, 17*mm
    tw = col_no_no_w = col_no_w + col_m_w + (col_s_w * 5)
    x0 = ml + (content_w - tw)/2
    
    # Header Tabel
    c.setFillColor(lightgrey); c.rect(x0, y - 2*row_h, tw, 2*row_h, fill=1); c.setFillColor(black)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(x0 + col_no_w/2, y - row_h - 2*mm, "No")
    c.drawCentredString(x0 + col_no_w + col_m_w/2, y - row_h - 2*mm, "Mata Pelajaran")
    c.drawCentredString(x0 + col_no_w + col_m_w + (col_s_w*5)/2, y - row_h/2 - 2*mm, "Nilai TKA/TKAD")
    for i in range(5):
        c.drawCentredString(x0 + col_no_w + col_m_w + (i*col_s_w) + col_s_w/2, y - 1.5*row_h - 2*mm, str(i+1))

    # Baris Mapel
    y_row = y - 2*row_h
    c.setFont("Helvetica", 10)
    sums = [[] for _ in range(5)]
    for i, m in enumerate(mapel_urut, 1):
        c.drawCentredString(x0 + col_no_w/2, y_row - row_h/2 - 2*mm, str(i))
        c.drawString(x0 + col_no_w + 2*mm, y_row - row_h/2 - 2*mm, m)
        for j in range(5):
            val = row.get(f"{m}_TKAD{j+1}", np.nan)
            c.drawCentredString(x0 + col_no_w + col_m_w + (j*col_s_w) + col_s_w/2, y_row - row_h/2 - 2*mm, format_score(val))
            if not pd.isna(val): sums[j].append(float(val))
        y_row -= row_h

    # Footer Tabel: Jumlah, Rata-rata, Peringkat
    c.setFont("Helvetica-Bold", 10)
    footer = [
        ("Jumlah", [sum(x) if x else np.nan for x in sums]),
        ("Rata-rata", [sum(x)/len(x) if x else np.nan for x in sums]),
        ("Peringkat", [row.get(f"Peringkat_TKAD{j+1}", np.nan) for j in range(5)])
    ]
    for lbl, data in footer:
        c.drawString(x0 + col_no_w + 2*mm, y_row - row_h/2 - 2*mm, lbl)
        for j, val in enumerate(data):
            c.drawCentredString(x0 + col_no_w + col_m_w + (j*col_s_w) + col_s_w/2, y_row - row_h/2 - 2*mm, format_score(val))
        y_row -= row_h

    # --- Menggambar Garis Grid (Perbaikan Garis yang Hilang) ---
    c.setLineWidth(0.5)
    total_table_rows = len(mapel_urut) + 2 + 3 # Header (2) + Mapel + Footer (3)
    for i in range(total_table_rows + 1):
        if i == 1: continue # Lewati garis horizontal tengah di header No/Mapel
        c.line(x0, y - i*row_h, x0 + tw, y - i*row_h)
    
    # Garis Vertikal
    c.line(x0, y, x0, y_row) # Paling Kiri
    c.line(x0 + col_no_w, y, x0 + col_no_w, y_row) # Setelah No
    c.line(x0 + col_no_w + col_m_w, y, x0 + col_no_w + col_m_w, y_row) # Setelah Mapel
    for i in range(1, 6):
        c.line(x0 + col_no_w + col_m_w + i*col_s_w, y - row_h, x0 + col_no_w + col_m_w + i*col_s_w, y_row)
    c.line(x0 + tw, y, x0 + tw, y_row) # Paling Kanan

    # --- Keterangan (Perbaikan Tampilan) ---
    y_ket = y_row - 10*mm
    c.setFont("Helvetica-Bold", 9)
    c.drawString(ml, y_ket, "Keterangan:")
    c.setFont("Helvetica", 9)
    for i, item in enumerate(asesmen_list, 1):
        y_ket -= 4*mm
        tgl_item = tgl_kegiatan_opsi[i-1] if i-1 < len(tgl_kegiatan_opsi) else "-"
        c.drawString(ml + 5*mm, y_ket, f"{i}. {item} ({tgl_item})")

    # Tanda Tangan
    tgl_ttd_str = f"{sel_tgl_ttd.day} {bulan_id[sel_tgl_ttd.strftime('%B')]} {sel_tgl_ttd.year}"
    x_ttd = width - mr - 65*mm
    y_sign = y_ket - 10*mm
    c.setFont("Helvetica", 11)
    c.drawString(x_ttd, y_sign, f"Banguntapan, {tgl_ttd_str}")
    c.drawString(x_ttd, y_sign - 5*mm, "Kepala Sekolah,")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_ttd, y_sign - 30*mm, "Alina Fiftiyani Nurjannah, M.Pd.")

def make_pdf(data_rows):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    for _, r in data_rows.iterrows():
        draw_student_page(c, r, sel_asesmen, sel_tahun, sel_tgl_kegiatan, mapel_urut, sel_tgl_ttd, asesmen_opsi)
        c.showPage()
    c.save(); buf.seek(0)
    return buf

# Tombol Download
if not df_kelas.empty:
    if sel_siswa != "-- Semua Siswa --":
        row_sel = df_kelas[df_kelas["Nama Siswa"] == sel_siswa]
        st.download_button("📄 Download PDF (Per Siswa)", make_pdf(row_sel), f"Laporan_{sel_siswa}.pdf")
    st.download_button("📄 Download PDF (Satu Kelas)", make_pdf(df_kelas), f"Laporan_Kelas_{sel_kelas}.pdf")
