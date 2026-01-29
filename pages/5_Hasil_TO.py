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
# === KONFIGURASI ===
# ===============================================
mapel_tetap = ["Bahasa Indonesia", "Matematika", "Bahasa Inggris", "Ilmu Pengetahuan Alam"]
mapel_urut = mapel_tetap

bulan_id = {
    "January": "Januari", "February": "Februari", "March": "Maret",
    "April": "April", "May": "Mei", "June": "Juni",
    "July": "Juli", "August": "Agustus", "September": "September",
    "October": "Oktober", "November": "November", "December": "Desember"
}

st.header("Laporan Hasil Persiapan dan Pemantapan")

# --- Input Pilihan ---
asesmen_opsi = [
    "TKA/TKAD Dikpora Kab Bantul 1", "TKA/TKAD MKKS SMP Kab Bantul 1",
    "TKA/TKAD MKKS SMP Kab Bantul 2", "TKA/TKAD Forum MKKS SMP D.I.Yogyakarta",
    "TKA/TKAD Dikpora Kab Bantul 2"
]
sel_asesmen = st.selectbox("Pilih Jenis Asesmen", asesmen_opsi)
tahun_opsi = [f"{th}/{th+1}" for th in range(2025, 2036)]
sel_tahun = st.selectbox("Pilih Tahun Pelajaran", tahun_opsi, index=0)
tgl_kegiatan_opsi = [
    "Tanggal 20 - 23 Oktober 2025", "Tanggal 3 - 6 November 2025",
    "Tanggal 19 - 22 Januari 2026", "Tanggal 2 - 5 Februari 2026", "Tanggal 9 - 12 Maret 2026"
]
sel_tgl_kegiatan = st.selectbox("Pilih Tanggal Kegiatan", tgl_kegiatan_opsi)
sel_tgl_ttd = st.date_input("Tanggal Tanda Tangan", datetime.now())

# === Template Excel ===
def generate_template():
    cols = ["Kelas", "NIS", "Nama Siswa"]
    for i in range(1, 6):
        for m in mapel_tetap:
            cols.append(f"{m}_TKAD{i}")
    df_template = pd.DataFrame(columns=cols)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer

st.download_button("📥 Download Template Excel", data=generate_template(), file_name="Template_Nilai.xlsx")

# === Upload & Olah Data ===
uploaded = st.file_uploader("Unggah file Excel", type=["xlsx"])
if not uploaded: st.stop()

df = pd.read_excel(uploaded).fillna(np.nan)
df.columns = df.columns.str.strip()

# Validasi Dasar
for c in ["Kelas", "Nama Siswa"]:
    if c not in df.columns:
        st.error(f"Kolom {c} tidak ada!"); st.stop()

# --- Hitung Peringkat Otomatis untuk Tiap Kolom TKAD ---
# Kita buat kolom Rank_TKAD1 sampai Rank_TKAD5 secara tersembunyi
for i in range(1, 6):
    cols_sesi = [f"{m}_TKAD{i}" for m in mapel_tetap]
    available = [c for c in cols_sesi if c in df.columns]
    if available:
        # Hitung total per siswa untuk sesi ini
        df[f"Total_Sesi_{i}"] = df[available].sum(axis=1, skipna=True)
        # Filter hanya yang punya nilai
        mask = df[available].notna().any(axis=1)
        df.loc[mask, f"Peringkat_TKAD{i}"] = df[mask].groupby("Kelas")[f"Total_Sesi_{i}"].rank(ascending=False, method="min")
    else:
        df[f"Peringkat_TKAD{i}"] = np.nan

# Filter Kelas
kelas_list = sorted(df["Kelas"].astype(str).unique())
sel_kelas = st.selectbox("Pilih Kelas", kelas_list)
df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

siswa_list = df_kelas["Nama Siswa"].astype(str).tolist()
sel_siswa = st.selectbox("Pilih Siswa", ["-- Semua Siswa --"] + siswa_list)

# ===============================================
# === PDF GENERATOR DENGAN BARIS PERINGKAT ===
# ===============================================
def format_val(val):
    if pd.isna(val) or val == "": return ""
    try: return f"{float(val):.2f}" if float(val) % 1 != 0 else str(int(float(val)))
    except: return str(val)

def draw_student_page(c, row):
    width, height = A4
    ml, mt = 30*mm, 25*mm
    y = height - mt

    # Header / Kop (Disederhanakan)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "SMP NEGERI 2 BANGUNTAPAN")
    y -= 15*mm
    c.setLineWidth(1); c.line(ml, y, width-15*mm, y); y -= 10*mm

    # Judul
    c.drawCentredString(width/2, y, "LAPORAN HASIL PERSIAPAN DAN PEMANTAPAN")
    y -= 15*mm

    # Identitas (Tanpa Peringkat di sini, karena sudah masuk tabel)
    for lbl, k in [("Nama", "Nama Siswa"), ("NIS", "NIS"), ("Kelas", "Kelas")]:
        c.setFont("Helvetica", 11)
        c.drawString(ml+10*mm, y, lbl)
        c.drawString(ml+35*mm, y, f": {row.get(k, '')}")
        y -= 6*mm
    y -= 5*mm

    # --- Tabel ---
    row_h, col_no, col_m, col_s = 7*mm, 12*mm, 65*mm, 17*mm
    tw = col_no + col_m + (col_s * 5)
    xs = ml + (width - ml - 15*mm - tw)/2

    # Header Tabel
    c.setFillColor(lightgrey); c.rect(xs, y-2*row_h, tw, 2*row_h, fill=1); c.setFillColor(black)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(xs+col_no/2, y-row_h-2*mm, "No")
    c.drawCentredString(xs+col_no+col_m/2, y-row_h-2*mm, "Mata Pelajaran")
    c.drawCentredString(xs+col_no+col_m+(col_s*5)/2, y-row_h/2-2*mm, "Nilai TKA/TKAD")
    for i in range(5):
        c.drawCentredString(xs+col_no+col_m+(i*col_s)+col_s/2, y-1.5*row_h-2*mm, str(i+1))

    # Data Mapel
    y_row = y - 2*row_h
    c.setFont("Helvetica", 10)
    sums = [[] for _ in range(5)]
    for i, m in enumerate(mapel_urut, 1):
        c.drawCentredString(xs+col_no/2, y_row-row_h/2-2*mm, str(i))
        c.drawString(xs+col_no+2*mm, y_row-row_h/2-2*mm, m)
        for j in range(5):
            val = row.get(f"{m}_TKAD{j+1}", np.nan)
            c.drawCentredString(xs+col_no+col_m+(j*col_s)+col_s/2, y_row-row_h/2-2*mm, format_val(val))
            if not pd.isna(val): sums[j].append(val)
        y_row -= row_h

    # Baris Tambahan: Jumlah, Rata-rata, dan PERINGKAT
    c.setFont("Helvetica-Bold", 10)
    footer_data = [
        ("Jumlah", [sum(x) if x else np.nan for x in sums]),
        ("Rata-rata", [sum(x)/len(x) if x else np.nan for x in sums]),
        ("Peringkat", [row.get(f"Peringkat_TKAD{j+1}", np.nan) for j in range(5)])
    ]

    for label, data_list in footer_data:
        c.drawString(xs+col_no+2*mm, y_row-row_h/2-2*mm, label)
        for j, v in enumerate(data_list):
            c.drawCentredString(xs+col_no+col_m+(j*col_s)+col_s/2, y_row-row_h/2-2*mm, format_val(v))
        y_row -= row_h

    # Garis-garis Tabel
    c.setLineWidth(0.5)
    total_grid_rows = len(mapel_urut) + 2 + 3 # +2 Header, +3 Footer
    for i in range(total_grid_rows + 1):
        if i == 1: continue
        c.line(xs, y - i*row_h, xs + tw, y - i*row_h)
    c.line(xs, y, xs, y_row) # Vertikal kiri
    c.line(xs+col_no, y, xs+col_no, y_row) # Vertikal No
    c.line(xs+col_no+col_m, y, xs+col_no+col_m, y_row) # Vertikal Mapel
    for i in range(1, 6): # Vertikal Nilai
        c.line(xs+col_no+col_m+i*col_s, y-row_h, xs+col_no+col_m+i*col_s, y_row)
    c.line(xs+tw, y, xs+tw, y_row) # Vertikal Kanan

    # Tanda Tangan
    y_sign = y_row - 15*mm
    tgl_str = f"{sel_tgl_ttd.day} {bulan_id[sel_tgl_ttd.strftime('%B')]} {sel_tgl_ttd.year}"
    c.setFont("Helvetica", 11)
    c.drawString(width-80*mm, y_sign, f"Banguntapan, {tgl_str}")
    c.drawString(width-80*mm, y_sign-5*mm, "Kepala Sekolah,")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(width-80*mm, y_sign-30*mm, "Alina Fiftiyani Nurjannah, M.Pd.")

def make_pdf(data_rows):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    for _, r in data_rows.iterrows():
        draw_student_page(c, r)
        c.showPage()
    c.save(); buf.seek(0)
    return buf

# Tombol Unduh
if sel_siswa != "-- Semua Siswa --":
    row_sel = df_kelas[df_kelas["Nama Siswa"] == sel_siswa]
    st.download_button("📄 PDF Siswa", make_pdf(row_sel), f"{sel_siswa}.pdf")
st.download_button("📄 PDF Satu Kelas", make_pdf(df_kelas), f"Kelas_{sel_kelas}.pdf")
