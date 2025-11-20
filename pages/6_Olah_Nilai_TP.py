import csv
import os
import streamlit as st
import pandas as pd
from io import BytesIO
import random
import numpy as np

# =========================================================
# KONFIGURASI DAN DATA LOADING
# =========================================================

# Variabel konfigurasi
COMMON_SUBJECTS = ["Matematika", "Bahasa Inggris", "IPA", "IPS", "Seni Budaya", "Informatika"]
CLASS_OPTIONS = ["7A", "7B", "7C", "7D", "8A", "8B", "9A", "9B"]
YEAR_OPTIONS = ["2025/2026", "2026/2027", "2027/2028", "2028/2029", "2029/2030"]

# Kolom-kolom nilai yang akan diinput/dihitung
SCORE_COLUMNS = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5', 'LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5', 'PTS', 'SAS', 'NR']
INPUT_SCORE_COLS = [c for c in SCORE_COLUMNS if c != 'NR']

# Peta untuk nama tampilan kolom
COLUMN_DISPLAY_MAP = {
    'TP1': 'TP-1', 'TP2': 'TP-2', 'TP3': 'TP-3', 'TP4': 'TP-4', 'TP5': 'TP-5',
    'LM_1': 'LM-1', 'LM_2': 'LM-2', 'LM_3': 'LM-3', 'LM_4': 'LM-4', 'LM_5': 'LM-5',
    'PTS': 'PTS', 'SAS': 'SAS/SAT', 'NR': 'NR',
    'Avg_TP': 'Rata-rata TP',
    'Avg_LM': 'Rata-rata LM',
    'Avg_PSA': 'Rata-rata PSA',
    'Deskripsi_NR': 'Deskripsi Rapor'
}
# Threshold KKM/Batas Ketuntasan untuk menentukan Tingkat Ketercapaian (TK)
KKM = 80

@st.cache_data
def load_dummy_data():
    """Membuat DataFrame dummy untuk semua siswa (sebagai fallback)."""
    data = {
        'NIS': [1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008],
        'Nama': ['Budi Santoso', 'Citra Dewi', 'Doni Pratama', 'Eka Fitriani', 'Fajar Nur', 'Gita Cahyani', 'Hendra Wijaya', 'Irma Suryani'],
        'Kelas': ['7A', '7A', '7A', '7A', '7B', '7B', '7C', '7C'],
    }
    df = pd.DataFrame(data)
    # Inisialisasi kolom nilai input dengan angka acak yang realistis
    for col in INPUT_SCORE_COLS:
        df[col] = np.random.randint(65, 95, size=len(df))
    df['NR'] = 0 # Initialize NR as integer
    df['NIS'] = df['NIS'].astype(str) # Pastikan NIS adalah string
    return df

@st.cache_data
def load_base_student_data():
    """
    Mencoba memuat data siswa dari daftar_siswa.csv secara default.
    Jika gagal atau file tidak ditemukan, gunakan data dummy.
    """
    file_path = "daftar_siswa.csv"
    try:
        if os.path.exists(file_path):
            df = pd.read_csv(file_path)

            # Standarisasi kolom
            column_rename_map = {
                'NISN': 'NIS', 'NAMA': 'Nama', 'KLAS': 'Kelas', 'KLS': 'Kelas',
            }

            df.columns = [col.strip() for col in df.columns]
            df = df.rename(columns=column_rename_map, errors='ignore')

            required_cols = ['NIS', 'Nama', 'Kelas']
            if all(col in df.columns for col in required_cols):
                # Tambahkan kolom nilai input yang hilang (termasuk LM_1 s/d LM_5) dan inisialisasi dengan 0
                for col in INPUT_SCORE_COLS:
                    if col not in df.columns:
                        df[col] = 0

                # Bersihkan tipe data dan spasi
                df['NIS'] = df['NIS'].astype(str).str.strip()
                df['Nama'] = df['Nama'].astype(str).str.strip()
                df['Kelas'] = df['Kelas'].astype(str).str.strip().str.upper()

                # Hapus kolom 'Nilai Rata-rata' dari CSV awal jika ada
                df = df.drop(columns=['Nilai Rata-rata'], errors='ignore')

                return df
            else:
                 return load_dummy_data()
        else:
            return load_dummy_data()

    except Exception as e:
        return load_dummy_data()


# Muat data dasar (prioritas CSV, fallback dummy)
df_all_students_base = load_base_student_data()

# =========================================================
# FUNGSI PERHITUNGAN DAN DESKRIPSI
# =========================================================

def calculate_nr(df_input: pd.DataFrame) -> pd.DataFrame:
    """
    Menghitung Nilai Rapor (NR) dan nilai perantara (Avg_TP, Avg_LM, Avg_PSA).
    Formula NR: NR = (Avg_TP + Avg_LM + 2 * Avg_PSA) / 4 (dengan pembagi adaptif)
    """
    df = df_input.copy()

    # 1. Hitung Rata-rata TP (hanya kolom TP1-TP5)
    tp_cols = [c for c in df.columns if c.startswith('TP') and len(c) == 3]
    df['Avg_TP'] = df[tp_cols].replace(0, np.nan).mean(axis=1).round(2)

    # 2. Hitung Rata-rata LM (dari kolom LM_1 hingga LM_5)
    lm_cols = [c for c in df.columns if c.startswith('LM_') and len(c) == 4]
    df['Avg_LM'] = df[lm_cols].replace(0, np.nan).mean(axis=1).round(2)

    # 3. Hitung Rata-rata Penilaian Sumatif Akhir (PSA)
    df['Avg_PSA'] = df[['PTS', 'SAS']].replace(0, np.nan).mean(axis=1).round(2)

    # 4. Hitung NR (Nilai Rapor): NR = (Avg_TP + Avg_LM + 2 * Avg_PSA) / (Jumlah Bobot Komponen Valid)

    # Komponen NR: Avg_TP (bobot 1), Avg_LM (bobot 1), Avg_PSA (bobot 2)
    nr_components = df[['Avg_TP', 'Avg_LM', 'Avg_PSA']].copy()
    nr_components['Avg_PSA_weighted'] = nr_components['Avg_PSA'] * 2

    # Hitung total nilai komponen yang valid (nilai > 0 atau bukan NaN)
    sum_components = (
        nr_components['Avg_TP'].fillna(0) +
        nr_components['Avg_LM'].fillna(0) +
        nr_components['Avg_PSA_weighted'].fillna(0)
    )

    # Hitung jumlah bobot komponen yang valid (maksimal 4)
    count_components = nr_components.apply(lambda row: sum([
        1 if pd.notna(row['Avg_TP']) and row['Avg_TP'] > 0 else 0,
        1 if pd.notna(row['Avg_LM']) and row['Avg_LM'] > 0 else 0,
        2 if pd.notna(row['Avg_PSA']) and row['Avg_PSA'] > 0 else 0
    ]), axis=1)

    df['NR'] = np.where(count_components > 0, sum_components / count_components, 0)

    # Bulatkan NR ke bilangan bulat terdekat
    df['NR'] = df['NR'].round(0).astype(int)

    return df

def calculate_tk_status(df_input: pd.DataFrame) -> pd.DataFrame:
    """
    Menambahkan kolom Tingkat Ketercapaian (TK) untuk setiap TP
    Berbasis threshold 80. Nilai >= 80 = 'T', Nilai < 80 = 'R'.
    Logika Khusus: Jika semua TP Tuntas (Nilai >= 80), ubah TP dengan nilai terkecil menjadi 'R'.
    Jika nilai TP adalah 0, status TK diatur menjadi "" (kosong).
    """
    df = df_input.copy()
    tp_cols = [c for c in df.columns if c.startswith('TP') and len(c) == 3]
    TK_THRESHOLD = KKM # Menggunakan KKM dari konfigurasi
    tk_cols = [f'TK_{col}' for col in tp_cols]

    # 1. Tentukan status T/R awal
    for tp_col in tp_cols:
        tk_col = f'TK_{tp_col}'

        def determine_tk(score):
            if score == 0:
                return ""
            elif score >= TK_THRESHOLD:
                return "T"
            else:
                return "R"

        df[tk_col] = df[tp_col].apply(determine_tk)

    # 2. Terapkan Logika Khusus (Demotion Rule)
    def apply_demotion_rule(row):
        # Filter TK hanya yang memiliki nilai (bukan "")
        active_tk_cols = [tk_col for tk_col in tk_cols if row.get(tk_col) != ""]

        if not active_tk_cols:
            return row

        # Cek apakah semua status TK yang aktif adalah 'T'
        is_all_t = all(row.get(tk_col) == 'T' for tk_col in active_tk_cols)

        if is_all_t:
            # Dapatkan skor TP yang TIDAK KOSONG
            active_tp_cols = [col.replace('TK_', '') for col in active_tk_cols]
            tp_scores = row[active_tp_cols]

            # Cari kolom TP (label) yang memiliki nilai terkecil di antara yang aktif
            min_score_tp_col = tp_scores.idxmin()

            # Konversi nama kolom TP ('TPx') menjadi nama kolom TK ('TK_TPx')
            tk_to_demote_col = f'TK_{min_score_tp_col}'

            # Ubah status TK kolom tersebut menjadi 'R'
            row[tk_to_demote_col] = 'R'

        return row

    # Terapkan logika hanya pada baris di mana terdapat input TP (total TP > 0)
    df['total_tp'] = df[tp_cols].sum(axis=1)
    df[tk_cols] = df.apply(lambda row: apply_demotion_rule(row) if row['total_tp'] > 0 else row, axis=1)[tk_cols]
    df = df.drop(columns=['total_tp'])

    return df

def generate_nr_description(df_input: pd.DataFrame) -> pd.DataFrame:
    """
    Membuat Deskripsi Naratif Nilai Rapor (Deskripsi_NR) berdasarkan status TK.
    """
    df = df_input.copy()
    tp_cols_prefix = [c for c in df.columns if c.startswith('TP') and len(c) == 3]
    tk_cols = [f'TK_{col}' for col in tp_cols_prefix]
    descriptions = []

    for index, row in df.iterrows():
        # Cari TP mana saja yang statusnya Remidi ('R')
        remidi_tps = [i + 1 for i, col in enumerate(tk_cols) if row.get(col) == 'R']

        if not remidi_tps:
            description = "Ananda telah menunjukkan penguasaan materi yang sangat baik dan tuntas pada seluruh Tujuan Pembelajaran."
        else:
            tp_list = [f"TP-{i}" for i in remidi_tps]
            if len(tp_list) > 1:
                tp_list_str = f"{', '.join(tp_list[:-1])}, dan {tp_list[-1]}"
            else:
                tp_list_str = tp_list[0]

            description = f"Ananda perlu meningkatkan pemahaman dan penguasaan pada materi di {tp_list_str}."

        descriptions.append(description)

    df['Deskripsi_NR'] = descriptions
    return df

# =========================================================
# FUNGSI EKSPOR EXCEL DENGAN BORDER
# =========================================================

def generate_excel_form_nilai_siswa(df, mapel, semester, kelas, tp, guru, nip):
    """
    Membuat buffer Excel untuk Form Nilai Siswa (Laporan Lengkap).
    - Header informasi: Keterangan di Kolom A (0), Nilai Isian di Kolom C (2) (dengan ':').
    - Kolom B (1) dikosongkan sebagai spacer visual.
    - Tabel data siswa: Dimulai dari Kolom A (NIS) pada baris ke-9.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Inisialisasi sheet.
        df_temp = pd.DataFrame()
        df_temp.to_excel(writer, sheet_name='Form Nilai', startrow=0, startcol=0, index=False)
        worksheet = writer.sheets['Form Nilai']

        # Definisi Format (Diasumsikan COLUMN_DISPLAY_MAP, KKM, dll. sudah didefinisikan)
        border_format = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        header_format = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'bold': True, 'fg_color': '#D9E1F2', 'text_wrap': True
        })
        text_format = workbook.add_format({
            'border': 1, 'align': 'left', 'valign': 'vcenter'
        })
        nr_format = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'bold': True
        })
        header_info_format = workbook.add_format({
            'align': 'left', 'valign': 'vcenter'
        })

        # ------------------------------------

        # Kolom yang diekspor dan pembersihan data
        LM_COLS_EXPORT = ['LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5']
        EXPORT_SCORE_COLS = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5'] + LM_COLS_EXPORT + ['PTS', 'SAS', 'Avg_TP', 'Avg_LM', 'Avg_PSA', 'NR', 'Deskripsi_NR']

        df_export = df[['NIS', 'Nama', 'Kelas'] + [c for c in EXPORT_SCORE_COLS if c in df.columns]]
        df_export.columns = ['NIS', 'NAMA SISWA', 'KELAS'] + [COLUMN_DISPLAY_MAP.get(c, c) for c in df_export.columns if c not in ['NIS', 'Nama', 'Kelas']]

        NUMERIC_COLS_EXPORT = [c for c in df_export.columns if c not in ['NIS', 'NAMA SISWA', 'KELAS', 'Deskripsi Rapor']]
        df_export[NUMERIC_COLS_EXPORT] = df_export[NUMERIC_COLS_EXPORT].fillna(0)

        # ====================================================================
        # 2. MENULIS HEADER INFORMASI (Keterangan di Kolom A, Nilai Isian di Kolom C)
        # ====================================================================
        header_data = {
            'Keterangan': [
                'Mata Pelajaran', 'Kelas', 'Semester', 'Tahun Pelajaran', 'KKTP',
                'Guru Mata Pelelajaran', 'NIP Guru'
            ],
            'Nilai': [
                mapel, kelas, semester, tp, KKM, guru, nip
            ]
        }
        header_df = pd.DataFrame(header_data)

        START_ROW_INFO = 0
        for r_idx, row in header_df.iterrows():
            # Kolom 0 (A): Keterangan
            worksheet.write(r_idx + START_ROW_INFO, 0, row['Keterangan'], header_info_format)

            # Kolom 2 (C): Nilai Isian dengan tanda :
            # 🔴 KOREKSI DITERAPKAN: Pindah Nilai Isian ke Kolom C (indeks 2)
            worksheet.write(r_idx + START_ROW_INFO, 2, ': ' + str(row['Nilai']), header_info_format)

        # Set lebar kolom A dan C untuk header informasi
        worksheet.set_column(0, 0, 20) # Kolom A (Keterangan)
        worksheet.set_column(2, 2, 30) # Kolom C (Nilai Isian)
        # Kolom B (indeks 1) dibiarkan sebagai spacer, lebar default

        # ====================================================================
        # 3. MENULIS DATA NILAI SISWA (Dimulai dari Kolom A / Index 0)
        # ====================================================================
        START_ROW_DATA = 9
        # Tabel data siswa dimulai dari Kolom A (indeks 0)
        COL_OFFSET = 0

        # Tulis Header Tabel Data Nilai
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(START_ROW_DATA, col_num + COL_OFFSET, value, header_format)

        # Tulis Data Siswa dengan Format Border
        num_rows = len(df_export)
        num_cols = len(df_export.columns)

        for row_num in range(START_ROW_DATA + 1, START_ROW_DATA + num_rows + 1):
            for col_num in range(num_cols):
                cell_value = df_export.iloc[row_num - (START_ROW_DATA + 1), col_num]
                column_name = df_export.columns[col_num]

                # Tentukan format
                current_format = border_format
                if column_name in ['NIS', 'NAMA SISWA', 'KELAS', 'Deskripsi Rapor']:
                    if isinstance(cell_value, (int, float)):
                        cell_value = str(cell_value)
                    current_format = text_format
                elif column_name == 'NR':
                    current_format = nr_format

                # Tulis data dengan offset
                worksheet.write(row_num, col_num + COL_OFFSET, cell_value, current_format)

        # ====================================================================
        # 4. SET LEBAR KOLOM UNTUK DATA SISWA
        # ====================================================================

        # NIS: Kolom A (indeks 0) -> Lebar 5
        worksheet.set_column(0 + COL_OFFSET, 0 + COL_OFFSET, 5)
        # Nama Siswa: Kolom B (indeks 1) -> Lebar 25
        worksheet.set_column(1 + COL_OFFSET, 1 + COL_OFFSET, 25)
        # Kelas: Kolom C (indeks 2) -> Lebar 5
        worksheet.set_column(2 + COL_OFFSET, 2 + COL_OFFSET, 7)

        # Kolom Nilai Numerik (D dan seterusnya)
        worksheet.set_column(3 + COL_OFFSET, num_cols - 2 + COL_OFFSET, 8)

        # Deskripsi Rapor (Kolom terakhir)
        worksheet.set_column(num_cols - 1 + COL_OFFSET, num_cols - 1 + COL_OFFSET, 60)

    return output.getvalue()

def generate_excel_report_tk(df, mapel, kelas, tp):
    """
    Membuat buffer Excel untuk Laporan Tingkat Ketercapaian (TK)
    dengan header pendek (1, 2, 3, 4, 5) dan border pada tabel data.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Inisialisasi sheet
        df_temp = pd.DataFrame()
        df_temp.to_excel(writer, sheet_name='Laporan TK', startrow=0, startcol=0, index=False)
        worksheet = writer.sheets['Laporan TK']

        # Definisi Format dengan Border
        border_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        header_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bold': True,
            'fg_color': '#D9E1F2'
        })
        text_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        # ------------------------------------

        # Tampilkan kolom NR dan semua kolom TK
        tk_cols = [c for c in df.columns if c.startswith('TK_TP') and len(c) == 6]
        df_export = df[['NIS', 'Nama', 'Kelas', 'NR'] + tk_cols]

        # MODIFIKASI: Gunakan NAMA KOLOM PENDEK (1, 2, 3, 4, 5)
        tk_headers_short = [f'KTP-{i}' for i in range(1, len(tk_cols) + 1)]
        df_export.columns = ['NIS', 'NAMA SISWA', 'KELAS', 'NR'] + tk_headers_short

        # --- PERBAIKAN ERROR NAN/INF ---
        df_export['NR'] = df_export['NR'].fillna(0)
        # -------------------------------

        # 1. SIAPKAN DATA HEADER
        KKTP = KKM
        keterangan_df = pd.DataFrame({
            'Keterangan': [
                'Mata Pelajaran', 'Kelas', 'Tahun Pelajaran', 'Batas Ketuntasan (KKTP)'
            ]
        })

        # --- PERUBAHAN: Tambahkan Tanda Titik Dua (:) ---
        nilai_isian_dengan_titik_dua = [
            ': ' + str(mapel),
            ': ' + str(kelas),
            ': ' + str(tp),
            ': ' + str(KKTP)
        ]

        nilai_isian_df = pd.DataFrame({
            'Nilai_Isian': nilai_isian_dengan_titik_dua
        })
        # --- AKHIR PERUBAHAN ---

        combined_header_df = pd.concat([
            keterangan_df,
            pd.DataFrame(np.nan, index=keterangan_df.index, columns=['Kolom Kosong 1']),
            nilai_isian_df
        ], axis=1)

        # 2. TULIS DATA HEADER KE EXCEL
        combined_header_df.to_excel(writer, startrow=0, startcol=0, sheet_name='Laporan TK', index=False, header=False)


        # Tulis data TK siswa di bawah header
        START_ROW_DATA = 6

        # Tulis ulang Header dengan Format Border
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(START_ROW_DATA, col_num, value, header_format)

        # Terapkan Format Border pada Data Siswa
        num_rows = len(df_export)
        num_cols = len(df_export.columns)

        for row_num in range(START_ROW_DATA + 1, START_ROW_DATA + num_rows + 1):
            for col_num in range(num_cols):
                cell_value = df_export.iloc[row_num - (START_ROW_DATA + 1), col_num]

                column_name = df_export.columns[col_num]

                # Format khusus untuk kolom teks (NIS, Nama, Kelas)
                if column_name in ['NIS', 'NAMA SISWA', 'KELAS']:
                    if isinstance(cell_value, (int, float)):
                        cell_value = str(cell_value)
                    worksheet.write(row_num, col_num, cell_value, text_format)
                else:
                    # NR (numerik) dan status TK (teks/angka 0)
                    worksheet.write(row_num, col_num, cell_value, border_format)

        # Set lebar kolom
        worksheet.set_column(1, 1, 25) # Nama Siswa

    return output.getvalue()


# =========================================================
# APLIKASI STREAMLIT UTAMA
# =========================================================

st.set_page_config(layout="wide", page_title="Editor Nilai Kurikulum Merdeka (LM x 5)")

# Injeksi CSS untuk gaya
st.markdown(
    """
    <style>
    /* Hilangkan index di Streamlit Data Editor */
    div[data-testid="stDataFrame"] .st-bd {
        margin-left: -5px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("Olah Nilai Rapor / Griya Rapor")
st.write("---")

# --- Bagian Sidebar untuk Unggah Data Siswa ---
st.sidebar.header("Data Siswa")
uploaded_file = st.sidebar.file_uploader(
    "Unggah File CSV Data Siswa (Kolom wajib: NIS, Nama, Kelas)", type="csv"
)

# Menentukan data yang akan digunakan (Uploaded > Base CSV > Dummy)
if uploaded_file is not None:
    try:
        df_all_students_raw = pd.read_csv(uploaded_file)

        # Bersihkan dan rename kolom seperti pada fungsi loading
        df_all_students_raw.columns = [col.strip() for col in df_all_students_raw.columns]
        df_all_students = df_all_students_raw.rename(columns={
            'NISN': 'NIS', 'NAMA': 'Nama', 'KLAS': 'Kelas', 'KLS': 'Kelas'
        }, errors='ignore')

        required_cols = ['NIS', 'Nama', 'Kelas']
        if all(col in df_all_students.columns for col in required_cols):
            df_all_students['NIS'] = df_all_students['NIS'].astype(str).str.strip()
            df_all_students['Nama'] = df_all_students['Nama'].astype(str).str.strip()
            df_all_students['Kelas'] = df_all_students['Kelas'].astype(str).str.strip().str.upper()
        else:
            st.sidebar.error("CSV yang diunggah harus memiliki kolom 'NIS', 'Nama', dan 'Kelas'. Menggunakan data dasar.")
            df_all_students = df_all_students_base

    except Exception as e:
        st.sidebar.error(f"Terjadi kesalahan saat memuat CSV yang diunggah: {e}. Menggunakan data dasar.")
        df_all_students = df_all_students_base
else:
    df_all_students = df_all_students_base

if len(df_all_students) > 0 and 'Nama' in df_all_students.columns:
    st.sidebar.info(f"Total data siswa yang dimuat: **{len(df_all_students)}**.")


st.header("⚙️ Pengaturan Form Nilai")

# --- Form Input Data Guru dan Kelas ---
col1, col2, col3 = st.columns(3)
with col1:
    mapel_terpilih = st.selectbox("Mata Pelajaran", COMMON_SUBJECTS)
with col2:
    # Tentukan opsi kelas berdasarkan data yang dimuat
    available_classes = sorted(df_all_students['Kelas'].unique().tolist()) if len(df_all_students) > 0 and 'Kelas' in df_all_students.columns else CLASS_OPTIONS

    default_class_index = 0
    if len(available_classes) > 0:
        try:
            most_common_class = df_all_students['Kelas'].mode().iloc[0]
            default_class_index = available_classes.index(most_common_class)
        except:
            default_class_index = 0

    kelas_input = st.selectbox("Pilih Kelas", available_classes, index=default_class_index if len(available_classes) > default_class_index else 0)

with col3:
    semester_input = st.selectbox("Semester", ["Ganjil", "Genap"])

col4, col5 = st.columns(2)
with col4:
    tahun_pelajaran_input = st.selectbox("Tahun Pelajaran", YEAR_OPTIONS)
with col5:
    guru_input = st.text_input("Nama Guru Pengampu", "")
    nip_guru_input = st.text_input("NIP Guru", "")


# Filter data siswa berdasarkan input kelas
if 'Kelas' in df_all_students.columns:
    df_base = df_all_students[df_all_students['Kelas'] == kelas_input].reset_index(drop=True)
else:
    df_base = pd.DataFrame() # DataFrame kosong jika tidak ada kolom Kelas

# Kunci state unik berdasarkan filter (Kelas dan Mapel)
state_key = f'edited_data_{kelas_input}_{mapel_terpilih}'


st.header(f"Tabel Input Nilai Kelas {kelas_input} ({mapel_terpilih})")

# 1. LOGIC INITIALIZATION AND RESET STATE
if df_base is not None and len(df_base) > 0:
    # Cek apakah data di session state perlu diinisialisasi atau direset karena filter berubah
    is_data_in_state_correct = (
        state_key in st.session_state and
        st.session_state[state_key].shape[0] == df_base.shape[0] and
        st.session_state[state_key]['NIS'].astype(str).equals(df_base['NIS'].astype(str))
    )

    if not is_data_in_state_correct:
        st.session_state[state_key] = df_base.copy()

        # Pastikan kolom nilai input (termasuk LM_1 s/d LM_5) sudah ada dan bertipe integer
        for col in INPUT_SCORE_COLS:
            if col not in st.session_state[state_key].columns:
                 st.session_state[state_key][col] = 0

            # Pastikan tipe data nilai adalah numerik
            st.session_state[state_key][col] = pd.to_numeric(st.session_state[state_key][col], errors='coerce').fillna(0).astype(int)

        # Inisialisasi awal, pastikan semua kolom terhitung
        df_init = calculate_nr(st.session_state[state_key])
        df_init = calculate_tk_status(df_init)
        df_init = generate_nr_description(df_init)
        st.session_state[state_key] = df_init.copy()

    st.info("Salin (**paste**) nilai dari Excel di kolom TP, **LM**, PTS, dan SAS/SAT.")

    # 2. KONFIGURASI COLUMN UNTUK DATA EDITOR
    INPUT_COLS_FOR_EDITOR = [c for c in SCORE_COLUMNS if c not in ['NR']]

    column_config = {
        'NIS': st.column_config.Column("NIS", disabled=True),
        'Nama': st.column_config.Column("Nama", disabled=True),
        # Konfigurasi kolom nilai input (TP, 5xLM, PTS, SAS)
        **{
            col: st.column_config.NumberColumn(
                COLUMN_DISPLAY_MAP.get(col, col),
                min_value=0, max_value=100, step=1, default=0, format="%d"
            )
            for col in INPUT_COLS_FOR_EDITOR
        },
        # NR di-disable
        'NR': st.column_config.NumberColumn("NR (Otomatis)", disabled=True, format="%d", help=f"Nilai Rapor dihitung. Batas Ketuntasan (TK): {KKM}"),
    }

    # Data yang ditampilkan di editor (dihilangkan kolom hitungan perantara dan Kelas)
    columns_to_drop_editor = ['Avg_TP', 'Avg_LM', 'Avg_PSA', 'Deskripsi_NR', 'Kelas'] + [c for c in st.session_state[state_key].columns if c.startswith('TK_') ]
    df_editor_input = st.session_state[state_key].drop(columns=columns_to_drop_editor, errors='ignore')

    # Tampilkan editor dan simpan hasilnya
    edited_df = st.data_editor(
        df_editor_input,
        column_config=column_config,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key=f"data_editor_{state_key}"
    )

    # 3. Lakukan Perhitungan NR, TK, dan DESKRIPSI setiap kali data diubah
    if not edited_df.equals(df_editor_input):
        # Convert edited columns back to numeric
        for col in INPUT_COLS_FOR_EDITOR:
            edited_df[col] = pd.to_numeric(edited_df[col], errors='coerce').fillna(0).astype(int)

        # Ambil kembali kolom yang tersembunyi ('Kelas', 'NIS', 'Nama')
        for col in ['NIS', 'Nama']:
            if col in df_base.columns:
                edited_df[col] = df_base[col]
        if 'Kelas' in df_base.columns and 'Kelas' not in edited_df.columns:
            edited_df.insert(2, 'Kelas', df_base['Kelas'])

        # Hitung ulang NR, TK, dan Deskripsi
        df_nr_calculated = calculate_nr(edited_df)
        df_tk_calculated = calculate_tk_status(df_nr_calculated)
        df_final_calculated = generate_nr_description(df_tk_calculated)

        # Update session state dengan data yang sudah dihitung
        st.session_state[state_key] = df_final_calculated.copy()

    # Tampilkan hasil perhitungan
    df_final_calculated = st.session_state[state_key]

    st.subheader("Hasil Nilai Akhir, dan Kriteria Tujuan Pembelajaran")

    # Kolom untuk ditampilkan: NIS, Nama, Nilai Rata-rata, NR, Status TK, Deskripsi
    tk_cols_display = [c for c in df_final_calculated.columns if c.startswith('TK_TP')]
    LM_COLS_DISPLAY = [f'LM_{i}' for i in range(1, 6)]
    SCORE_COLS_DISPLAY_KEPT = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5'] + LM_COLS_DISPLAY + ['PTS', 'SAS', 'NR', 'Deskripsi_NR']

    # Daftar kolom yang akan ditampilkan
    display_cols = ['NIS', 'Nama'] + [c for c in SCORE_COLS_DISPLAY_KEPT if c in df_final_calculated.columns] + tk_cols_display

    # Rename kolom TK
    display_map_tk = {c: c.replace('TK_', 'Status ') + ' (T/R)' for c in tk_cols_display}

    # Tampilkan data dengan kolom yang sudah difilter
    df_display = df_final_calculated[display_cols].rename(columns={**display_map_tk})

    st.dataframe(df_display,
                 hide_index=True, use_container_width=True)

    df_siswa_to_export = df_final_calculated

else:
    st.warning(f"⚠️ **Tidak ditemukan** data siswa yang **Kelas**-nya sama persis dengan **{kelas_input}** di data yang dimuat. Pastikan file data siswa (CSV) sudah dimuat dan memiliki kolom 'Kelas' yang sesuai.")
    df_siswa_to_export = None


st.write("---")

# =========================================================
# EKSPOR HASIL
# =========================================================
if df_siswa_to_export is not None and len(df_siswa_to_export) > 0:
    st.header("⬆️ Ekspor Hasil Nilai Rapor")

    col_form, col_tk = st.columns(2)

    with col_form:
        excel_buffer_nilai = generate_excel_form_nilai_siswa(
            df_siswa_to_export,
            mapel_terpilih,
            semester_input,
            kelas_input,
            tahun_pelajaran_input,
            guru_input,
            nip_guru_input
        )
        st.download_button(
            label="⬇️ Ekspor / Cetak Nilai Rapor dengan Deskripsi Rapor",
            file_name=f"Nilai_Akhir_{kelas_input}_{mapel_terpilih}_{semester_input}.xlsx",
            data=excel_buffer_nilai,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_nilai"
        )

    with col_tk:
        excel_buffer_tk = generate_excel_report_tk(
            df_siswa_to_export,
            mapel_terpilih,
            kelas_input,
            tahun_pelajaran_input
        )
        st.download_button(
            label="⬇️ Unduh Nilai Rapor dan Kriteria TP",
            file_name=f"Nilai_Rapor_K_TP_{kelas_input}_{mapel_terpilih}_{semester_input}.xlsx",
            data=excel_buffer_tk,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_tk_report"
        )

st.write("---")
