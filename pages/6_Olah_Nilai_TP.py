import csv
import os
import streamlit as st
import pandas as pd
from io import BytesIO
import random
import numpy as np
from typing import List, Dict, Any, Union

# =========================================================
# KONFIGURASI DAN DATA LOADING
# =========================================================

# Variabel konfigurasi
COMMON_SUBJECTS = ["Matematika", "Bahasa Inggris", "IPA", "IPS", "Bahasa Indonesia","Seni Budaya", "P.Pancasila", "Pendidikan Agama", "Bahasa Jawa", "PJOK", "Informatika", "Prakarya"]
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
                # Tambahkan kolom nilai input yang hilang (termasuk LM_1 s/d LM_5)
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
# (Termasuk Logika Demotion Rule)
# =========================================================

def calculate_nr(df_input: pd.DataFrame) -> pd.DataFrame:
    """
    Menghitung Nilai Rapor (NR) dan nilai perantara (Avg_TP, Avg_LM, Avg_PSA).
    LOGIKA BARU: NR = (1*Avg_TP + 2*Avg_LM + 1*Avg_PSA) / (Total Bobot Valid)
    """
    df = df_input.copy()

    # 1. Hitung Rata-rata TP (hanya kolom TP1-TP5)
    tp_cols = [c for c in df.columns if c.startswith('TP') and len(c) == 3]
    # Menggunakan replace(0, np.nan).mean() memastikan pembagi adaptif (mengabaikan nilai 0)
    df['Avg_TP'] = df[tp_cols].replace(0, np.nan).mean(axis=1).round(2)

    # 2. Hitung Rata-rata LM (dari kolom LM_1 hingga LM_5)
    lm_cols = [c for c in df.columns if c.startswith('LM_') and len(c) == 4]
    # Menggunakan replace(0, np.nan).mean() memastikan pembagi adaptif (mengabaikan nilai 0)
    df['Avg_LM'] = df[lm_cols].replace(0, np.nan).mean(axis=1).round(2)

    # 3. Hitung Rata-rata Penilaian Sumatif Akhir (PSA)
    # Rata-rata PSA (PTS dan SAS) dengan pembagi adaptif (total yang tidak 0)
    df['Avg_PSA'] = np.where(
        (df['PTS'] + df['SAS']) > 0,
        (df['PTS'] + df['SAS']) / 2,
        0
    ).round(2)

    # 4. Hitung NR (Nilai Rapor): NR = (1*Avg_TP + 2*Avg_LM + 1*Avg_PSA) / (Jumlah Bobot Komponen Valid)

    # Komponen NR: Avg_TP (bobot 1), Avg_LM (bobot 2), Avg_PSA (bobot 1)
    nr_components = df[['Avg_TP', 'Avg_LM', 'Avg_PSA']].copy()

    # Komponen nilai dikalikan dengan bobotnya
    nr_components['Avg_TP_weighted'] = nr_components['Avg_TP'] * 1
    nr_components['Avg_LM_weighted'] = nr_components['Avg_LM'] * 2  # **MODIFIKASI BOBOT (1 -> 2)**
    nr_components['Avg_PSA_weighted'] = nr_components['Avg_PSA'] * 1 # **MODIFIKASI BOBOT (2 -> 1)**

    # Hitung total nilai komponen yang valid (nilai > 0 atau bukan NaN)
    sum_components = (
        nr_components['Avg_TP_weighted'].fillna(0) +
        nr_components['Avg_LM_weighted'].fillna(0) +
        nr_components['Avg_PSA_weighted'].fillna(0)
    )

    # Hitung jumlah bobot komponen yang valid (maksimal 4)
    count_components = nr_components.apply(lambda row: sum([1 if pd.notna(row['Avg_TP']) and row['Avg_TP'] > 0 else 0,
                                                           2 if pd.notna(row['Avg_LM']) and row['Avg_LM'] > 0 else 0, # **MODIFIKASI BOBOT (1 -> 2)**
                                                           1 if pd.notna(row['Avg_PSA']) and row['Avg_PSA'] > 0 else 0]), axis=1) # **MODIFIKASI BOBOT (2 -> 1)**

    # NR akan 0 jika count_components = 0
    df['NR'] = np.where(count_components > 0, sum_components / count_components, 0)

    # RULE: NR must be 0 if Avg_PSA (PTS and SAS) is 0
    df['NR'] = np.where(df['Avg_PSA'] > 0, df['NR'], 0)

    # Bulatkan NR ke bilangan bulat terdekat
    df['NR'] = df['NR'].round(0).astype(int)

    return df

def calculate_tk_status(df_input: pd.DataFrame) -> pd.DataFrame:
    """
    Menentukan Status TP1–TP5 (T/R) termasuk validasi tambahan (Demotion Rule).
    """
    df = df_input.copy()
    # Pastikan kolom TP ada; gunakan urutan TP1..TP5
    tp_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    tk_cols = [f'TK_{tp}' for tp in tp_cols]
    threshold = KKM

    # ---- Tahap 1: Tentukan T/R awal ----
    for tp in tp_cols:
        tk = f'TK_{tp}'
        df[tk] = df[tp].apply(
            lambda x: "" if x == 0 or pd.isna(x)
            else "T" if x >= threshold
            else "R"
        )

    # ---- Tahap 2: Validasi: Demotion Rule ----
    # Jika SEMUA nilai terisi (non-zero) dan semuanya >= threshold,
    # maka pilih TP dengan nilai terkecil dan ubah statusnya menjadi 'R'
    def apply_validation_rule(row):
        # Ambil hanya TP yang terisi (non-zero)
        filled_tps = {tp: row[tp] for tp in tp_cols if pd.notna(row[tp]) and row[tp] > 0}

        # Jika kurang dari 2 nilai terisi, tidak perlu menerapkan aturan demotion
        if len(filled_tps) < 2:
            return row

        # Cek apakah semua nilai yang terisi >= threshold
        all_t = all(v >= threshold for v in filled_tps.values())

        if all_t:
            # Temukan TP dengan nilai terkecil (jika multiple equal, ambil yang pertama)
            smallest_tp = min(filled_tps, key=filled_tps.get)
            smallest_tk = f'TK_{smallest_tp}'
            row[smallest_tk] = 'R'

        return row

    df = df.apply(apply_validation_rule, axis=1)

    return df

def generate_nr_description(df_input: pd.DataFrame) -> pd.DataFrame:
    """
    Membuat Deskripsi Naratif Nilai Rapor (Deskripsi_NR) berdasarkan status TK yang SUDAH termasuk Demotion Rule.
    """
    df = df_input.copy()
    tp_cols_prefix = [c for c in df.columns if c.startswith('TP') and len(c) == 3]
    tk_cols = [f'TK_{col}' for col in tp_cols_prefix]
    descriptions = []

    for index, row in df.iterrows():
        # Cari TP mana saja yang statusnya Remidi ('R')
        remidi_tps = [i + 1 for i, col in enumerate(tk_cols) if row.get(col) == 'R']

        if not remidi_tps:
            # Memastikan bahwa tidak ada TP yang diisi (Total TP Score = 0)
            tp_scores_sum = row[tp_cols_prefix].sum()

            if tp_scores_sum == 0:
                 # Jika tidak ada nilai TP yang dimasukkan
                 description = "Nilai Tujuan Pembelajaran belum diinput."
            else:
                 # Jika semua TP Tuntas (sesuai Demotion Rule)
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
# HELPERS UNTUK EXCEL
# =========================================================

def col_idx_to_excel(col_idx):
    """Convert 0-based column index to Excel column letters (0->A)."""
    # convert to 1-based
    col_idx += 1
    letters = ""
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letters = chr(65 + remainder) + letters
        col_idx = int(col_idx)
    return letters

def write_form_nilai_sheet(df, mapel, semester, kelas, tp, guru, nip, writer, sheet_name):
    """
    Internal function: Menulis satu sheet Form Nilai Siswa (Laporan Lengkap) ke writer yang sudah ada.
    Status TP ditulis sebagai RUMUS Excel agar dapat diperbarui secara real-time,
    termasuk logika Demotion Rule yang kompleks.
    """
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name) # Tambahkan sheet baru

    # Definisi Format
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

    # Format untuk kolom Status TP (warna kuning lembut) dan locked (NILAI RUMUS)
    status_tp_protected_format = workbook.add_format({
        'bg_color': '#FFF2CC',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'locked': True
    })
    # Format untuk kolom Rumus lainnya (warna kuning lembut) dan locked
    formula_protected_format = workbook.add_format({
        'bg_color': '#FFF2CC',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'locked': True
    })
    header_info_format = workbook.add_format({
        'align': 'left', 'valign': 'vcenter'
    })


    # ------------------------------------

    # Kolom yang diekspor dan pembersihan data
    LM_COLS_EXPORT = ['LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5']
    TK_COLS_EXPORT = [c for c in df.columns if c.startswith('TK_TP') and len(c) == 6] # Hanya untuk header, tidak di df_export
    TP_SCORE_COLS = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    INPUT_SCORE_COLS_ALL = TP_SCORE_COLS + LM_COLS_EXPORT + ['PTS', 'SAS']

    # Kolom Nilai Inti (tanpa Status TK)
    CORE_SCORE_COLS = TP_SCORE_COLS + LM_COLS_EXPORT + ['PTS', 'SAS', 'Avg_TP', 'Avg_LM', 'Avg_PSA', 'NR']

    # Urutan Kolom Data di DataFrame (Data statis/numerik saja, Status TK dihilangkan)
    FINAL_COLS_ORDER_DATA = ['NIS', 'Nama', 'Kelas'] + \
                            [c for c in CORE_SCORE_COLS if c in df.columns] + \
                            ['Deskripsi_NR']

    # DataFrame hanya berisi data statis/numerik dan deskripsi
    df_export = df[FINAL_COLS_ORDER_DATA].copy()

    # Pastikan numeric diisi 0 (jika NaN)
    NUMERIC_COLS_EXPORT = [c for c in df_export.columns if c not in ['NIS', 'Nama', 'Kelas', 'Deskripsi_NR']]
    df_export[NUMERIC_COLS_EXPORT] = df_export[NUMERIC_COLS_EXPORT].fillna(0)

    # Buat list nama kolom UNTUK HEADER (termasuk Status TP)
    tk_display_map = {c: c.replace('TK_', 'Status ') for c in TK_COLS_EXPORT}

    # List nama kolom di Excel (urutan: Data, Status TK, Deskripsi Rapor)
    HEADER_COLS_ORDER = ['NIS', 'NAMA SISWA', 'KELAS'] + \
                  [COLUMN_DISPLAY_MAP.get(c, c) for c in CORE_SCORE_COLS if c in df.columns] + \
                  [tk_display_map.get(c, c) for c in TK_COLS_EXPORT] + \
                  [COLUMN_DISPLAY_MAP.get('Deskripsi_NR', 'Deskripsi Rapor')]

    # ====================================================================
    # 2. MENULIS HEADER INFORMASI
    # ====================================================================
    START_ROW_INFO = 0
    INFO_COL_START = 6 # Column G (untuk yang diletakkan di kanan)

    # ------------------------------------------------
    # Bagian Kiri (Kolom A & C)
    # ------------------------------------------------
    worksheet.write(0 + START_ROW_INFO, 0, 'Mata Pelajaran', header_info_format)
    worksheet.write(0 + START_ROW_INFO, 2, ': ' + str(mapel), header_info_format)
    worksheet.write(1 + START_ROW_INFO, 0, 'Kelas', header_info_format)
    worksheet.write(1 + START_ROW_INFO, 2, ': ' + str(kelas), header_info_format)
    worksheet.write(2 + START_ROW_INFO, 0, 'Semester', header_info_format)
    worksheet.write(2 + START_ROW_INFO, 2, ': ' + str(semester), header_info_format)

    # Tulis KKTP sebagai teks statis di kolom C (C4). Kolom D4 tidak lagi diisi.
    worksheet.write(3 + START_ROW_INFO, 0, 'KKTP', header_info_format)
    worksheet.write(3 + START_ROW_INFO, 2, f": {KKM}", header_info_format) # Tulis ': 80' di C4

    # ------------------------------------------------
    # Bagian Kanan (Kolom G & I)
    # ------------------------------------------------
    worksheet.write(0 + START_ROW_INFO, INFO_COL_START, 'Tahun Pelajaran', header_info_format)
    worksheet.write(0 + START_ROW_INFO, INFO_COL_START + 2, ': ' + str(tp), header_info_format)

    worksheet.write(1 + START_ROW_INFO, INFO_COL_START, 'Guru Mata Pelelajaran', header_info_format)
    worksheet.write(1 + START_ROW_INFO, INFO_COL_START + 2, ': ' + str(guru), header_info_format)

    worksheet.write(2 + START_ROW_INFO, INFO_COL_START, 'NIP Guru', header_info_format)
    worksheet.write(2 + START_ROW_INFO, INFO_COL_START + 2, ': ' + str(nip), header_info_format)

    # Set lebar kolom untuk header info
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(2, 2, 30)
    worksheet.set_column(INFO_COL_START, INFO_COL_START, 20)
    worksheet.set_column(INFO_COL_START + 2, INFO_COL_START + 2, 30)

    # ====================================================================
    # 3. MENULIS DATA NILAI SISWA
    # ====================================================================
    START_ROW_DATA = 6
    COL_OFFSET = 0

    # Tulis Header Tabel Data Nilai
    for col_num, value in enumerate(HEADER_COLS_ORDER): # Menggunakan HEADER_COLS_ORDER
        worksheet.write(START_ROW_DATA, col_num + COL_OFFSET, value, header_format)

    # Pre-calc column indexes by title (dynamic)
    cols = list(HEADER_COLS_ORDER)
    def idx_of(title):
        return cols.index(title) if title in cols else None

    # Indeks kolom hasil perhitungan yang berisi rumus
    idx_avg_tp = idx_of('Rata-rata TP')
    idx_avg_lm = idx_of('Rata-rata LM')
    idx_avg_psa = idx_of('Rata-rata PSA')
    idx_nr = idx_of('NR')

    # NEW: Find Status TP and corresponding TP Score indices
    status_tp_map = {f'Status TP{i}': f'TP-{i}' for i in range(1, 6)}

    # Status TP Index -> TP Score Index
    status_tp_indices = {
        idx_of(status_name): idx_of(tp_score_name)
        for status_name, tp_score_name in status_tp_map.items()
        if idx_of(status_name) is not None and idx_of(tp_score_name) is not None
    }

    # Kolom untuk rumus AVG dan NR
    tp_score_display_titles = ['TP-1','TP-2','TP-3','TP-4','TP-5']
    lm_display_titles = ['LM-1','LM-2','LM-3','LM-4','LM-5']
    tp_score_indices = [cols.index(t) for t in tp_score_display_titles if t in cols]
    lm_indices = [cols.index(t) for t in lm_display_titles if t in cols]
    idx_pts = cols.index('PTS') if 'PTS' in cols else None
    idx_sas = cols.index('SAS/SAT') if 'SAS/SAT' in cols else None

    num_rows = len(df_export)
    num_cols = len(HEADER_COLS_ORDER)

    # Reverse Map untuk mengambil data dari df_export
    REVERSE_COLUMN_MAP = {v: k for k, v in COLUMN_DISPLAY_MAP.items()}
    REVERSE_COLUMN_MAP['NAMA SISWA'] = 'Nama'
    REVERSE_COLUMN_MAP['KELAS'] = 'Kelas'

    # Tentukan range absolut TP Scores untuk Demotion Rule
    if tp_score_indices:
        first_tp_col_idx = tp_score_indices[0] + COL_OFFSET
        last_tp_col_idx = tp_score_indices[-1] + COL_OFFSET

        # Kolom huruf untuk TP1 dan TP5
        first_tp_col_letter = col_idx_to_excel(first_tp_col_idx)
        last_tp_col_letter = col_idx_to_excel(last_tp_col_idx)
    else:
        first_tp_col_letter, last_tp_col_letter = None, None

    for r_i in range(num_rows):
        row_num = START_ROW_DATA + 1 + r_i
        excel_row = row_num + 1

        # Full TP Score Range (absolute, e.g., $D8:$H8)
        if first_tp_col_letter:
            full_tp_range = (
                f"${first_tp_col_letter}{excel_row}:"
                f"${last_tp_col_letter}{excel_row}"
            )
            min_score = f"MIN({full_tp_range})"
        else:
            full_tp_range = ""
            min_score = ""


        for col_num, column_name in enumerate(HEADER_COLS_ORDER):

            # 1. Menulis Status TP (Menggunakan Rumus Excel)
            if column_name.startswith('Status TP'):

                KKM_VALUE = KKM
                tp_score_idx = status_tp_indices.get(col_num)

                if tp_score_idx is not None and full_tp_range:

                    # ------------------- LOGIKA RUMUS DEMOTION RULE ---------------------

                    # Column letter for current TP score (e.g., D for TP-1)
                    col_tp_score_letter = col_idx_to_excel(tp_score_idx + COL_OFFSET)

                    # Current TP Cell (e.g., D8)
                    current_tp_cell = f"{col_tp_score_letter}{excel_row}"

                    # Current TP Range for COUNTIF (semi-absolute, e.g., $D8:D8, $D8:E8)
                    current_countif_range = (
                        f"${first_tp_col_letter}{excel_row}:"
                        f"{col_tp_score_letter}{excel_row}"
                    )

                    # Demotion Check Logic: AND(Score=MIN(Range); COUNTIF($Start:Current; MIN(Range))=1)
                    # Ini memastikan hanya TP pertama dengan nilai minimum yang diturunkan 'R'
                    demotion_check = (
                        f"AND({current_tp_cell}={min_score};"
                        f"COUNTIF({current_countif_range};{min_score})=1)"
                    )

                    # Full Formula: IF(Score=0;""; IF(Score<KKM;"R"; IF(DemotionCheck;"R";"T")))
                    formula_tk = (
                        f'=IF({current_tp_cell}=0;"";'
                        f'IF({current_tp_cell}<{KKM_VALUE};"R";'
                        f'IF({demotion_check};"R";"T")))'
                    )

                    # ------------------- AKHIR LOGIKA RUMUS DEMOTION RULE ---------------------

                    worksheet.write_formula(row_num, col_num + COL_OFFSET, formula_tk, status_tp_protected_format)

                continue # Lanjut ke kolom berikutnya

            # 2. Skip kolom rumus lainnya yang akan ditulis nanti (Avg_TP, Avg_LM, Avg_PSA, NR)
            if column_name in ['Rata-rata TP', 'Rata-rata LM', 'Rata-rata PSA', 'NR']:
                continue

            # 3. Tulis nilai statis/input

            # Cari nama kolom data di df_export
            data_column_name = REVERSE_COLUMN_MAP.get(column_name, column_name)

            if data_column_name not in df_export.columns:
                 continue

            cell_value = df_export.iloc[r_i, df_export.columns.get_loc(data_column_name)]
            current_format = border_format

            # Cek apakah kolom tersebut adalah kolom input nilai
            is_input_score = data_column_name in INPUT_SCORE_COLS_ALL

            # PERBAIKAN: Jika input score bernilai 0, tulis sebagai string kosong ("") agar terlihat BLANK di Excel
            if is_input_score and cell_value == 0:
                cell_value = ""

            # Logic untuk format
            if data_column_name in ['NIS', 'Nama', 'Kelas', 'Deskripsi_NR']:
                if isinstance(cell_value, (int, float)):
                    cell_value = str(cell_value)
                current_format = text_format
            else:
                current_format = border_format # Untuk nilai numerik input

            # Tulis nilai statis/input
            worksheet.write(row_num, col_num + COL_OFFSET, cell_value, current_format)

        # TULIS RUMUS AVERAGEIF untuk Avg_TP dan Avg_LM
        if idx_avg_tp is not None and tp_score_indices:
            first_tp_col = col_idx_to_excel(tp_score_indices[0] + COL_OFFSET)
            last_tp_col = col_idx_to_excel(tp_score_indices[-1] + COL_OFFSET)

            # **MODIFIKASI: Menambahkan IFERROR untuk menampilkan 0 jika tidak ada nilai > 0**
            formula_avg_tp = (
                f"=IFERROR(AVERAGEIF({first_tp_col}{excel_row}:{last_tp_col}{excel_row};\">0\"); 0)"
            )
            worksheet.write_formula(row_num, idx_avg_tp + COL_OFFSET, formula_avg_tp, formula_protected_format)

        if idx_avg_lm is not None and lm_indices:
            first_lm_col = col_idx_to_excel(lm_indices[0] + COL_OFFSET)
            last_lm_col = col_idx_to_excel(lm_indices[-1] + COL_OFFSET)

            # **MODIFIKASI: Menambahkan IFERROR untuk menampilkan 0 jika tidak ada nilai > 0**
            formula_avg_lm = (
                f"=IFERROR(AVERAGEIF({first_lm_col}{excel_row}:{last_lm_col}{excel_row};\">0\"); 0)"
            )
            worksheet.write_formula(row_num, idx_avg_lm + COL_OFFSET, formula_avg_lm, formula_protected_format)

        # TULIS RUMUS Avg_PSA
        if idx_avg_psa is not None and idx_pts is not None and idx_sas is not None:
            col_pts = col_idx_to_excel(idx_pts + COL_OFFSET)
            col_sas = col_idx_to_excel(idx_sas + COL_OFFSET)
            # Rumus untuk rata-rata 2 nilai (PTS dan SAS).
            # Jika total > 0, hitung rata-rata
            formula_avg_psa = (
                f"=IF({col_pts}{excel_row}+{col_sas}{excel_row}>0; "
                f"({col_pts}{excel_row}+{col_sas}{excel_row})/2; 0)"
            )
            worksheet.write_formula(row_num, idx_avg_psa + COL_OFFSET, formula_avg_psa, formula_protected_format)

        # Rumus NR (Pembobotan 1:2:1)
        if idx_nr is not None and idx_avg_tp is not None and idx_avg_lm is not None and idx_avg_psa is not None:
            col_avg_tp = col_idx_to_excel(idx_avg_tp + COL_OFFSET)
            col_avg_lm = col_idx_to_excel(idx_avg_lm + COL_OFFSET)
            col_avg_psa = col_idx_to_excel(idx_avg_psa + COL_OFFSET)

            # Logika pembagi BARU: (Bobot TP=1 + Bobot LM=2 + Bobot PSA=1)
            calculation_denominator = (
                f"(({col_avg_tp}{excel_row}>0)*1+({col_avg_lm}{excel_row}>0)*2+({col_avg_psa}{excel_row}>0)*1)"
            )

            # Bagian utama perhitungan BARU: (Avg_TP + 2*Avg_LM + 1*Avg_PSA) / bobot_valid
            calculation_core = (
                f"({col_avg_tp}{excel_row}+2*{col_avg_lm}{excel_row}+{col_avg_psa}{excel_row})/"
                f"IF({calculation_denominator}=0;1;{calculation_denominator})" # Mencegah Div/0
            )

            # Rumus: =IF(Avg_PSA_Cell > 0; IFERROR(ROUND(core; 0); 0); 0)
            formula_nr = (
                f"=IF({col_avg_psa}{excel_row}>0; "
                f"IFERROR(ROUND({calculation_core};0);0);0)"
            )
            worksheet.write_formula(row_num, idx_nr + COL_OFFSET, formula_nr, formula_protected_format)

    # ====================================================================
    # 4. SET LEBAR KOLOM
    # ====================================================================
    worksheet.set_column(0 + COL_OFFSET, 0 + COL_OFFSET, 5) # NIS (A)
    worksheet.set_column(1 + COL_OFFSET, 1 + COL_OFFSET, 25) # Nama Siswa (B)
    worksheet.set_column(2 + COL_OFFSET, 2 + COL_OFFSET, 7) # Kelas (C)

    idx_desc = idx_of('Deskripsi Rapor')
    # Tentukan batas akhir kolom numerik/status
    end_col_numeric = idx_desc - 1 + COL_OFFSET if idx_desc is not None else len(HEADER_COLS_ORDER) - 1 + COL_OFFSET
    # Mulai dari kolom 3/D (TP-1) hingga kolom sebelum Deskripsi
    worksheet.set_column(3 + COL_OFFSET, end_col_numeric, 8) # Nilai Numerik dan Status

    if idx_desc is not None:
        worksheet.set_column(idx_desc + COL_OFFSET, idx_desc + COL_OFFSET, 60) # Deskripsi Rapor

    # Freeze row: START_ROW_DATA (baris header) + 2 (baris data pertama + 1)
    worksheet.freeze_panes(START_ROW_DATA + 2, 2)
    # Tidak perlu return output.getvalue() karena ini fungsi internal

def write_report_tk_sheet(df, mapel, kelas, tp, writer, sheet_name):
    """
    Internal function: Menulis satu sheet Laporan TK ke writer yang sudah ada.
    Status TK di sheet ini tetap ditulis sebagai nilai statis (Teks T/R) hasil perhitungan Python
    (termasuk Demotion Rule) agar deskripsi rapor tetap valid.
    """
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name) # Tambahkan sheet baru

    # Definisi Format
    border_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    header_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bold': True, 'fg_color': '#D9E1F2'
    })
    text_format = workbook.add_format({
        'border': 1, 'align': 'left', 'valign': 'vcenter'
    })
    header_info_format = workbook.add_format({
        'align': 'left', 'valign': 'vcenter'
    })
    # ------------------------------------

    # Tampilkan kolom NR dan semua kolom TK
    tk_cols = [c for c in df.columns if c.startswith('TK_TP') and len(c) == 6]
    df_export = df[['NIS', 'Nama', 'Kelas', 'NR'] + tk_cols].copy()

    # MODIFIKASI: Gunakan NAMA KOLOM PENDEK (1, 2, 3, 4, 5)
    tk_headers_short = [f'KTP-{i}' for i in range(1, len(tk_cols) + 1)]
    df_export.columns = ['NIS', 'NAMA SISWA', 'KELAS', 'NR'] + tk_headers_short

    df_export['NR'] = df_export['NR'].fillna(0)
    # Pastikan Status TK diubah menjadi string, dan NaN diubah string kosong ('')
    for col in [c for c in df_export.columns if c.startswith('KTP-')]:
        df_export[col] = df_export[col].astype(str).replace('nan', '')

    # Kita harus memastikan nilai 0 di kolom NR juga tidak ditampilkan
    df_export['NR'] = df_export['NR'].apply(lambda x: '' if x == 0 else x)


    # 1. SIAPKAN DATA HEADER
    KKTP = KKM
    header_data = {
        'Keterangan': [
            'Mata Pelajaran', 'Kelas', 'Tahun Pelajaran', 'Batas Ketuntasan (KKTP)'
        ],
        'Nilai_Isian': [
            ': ' + str(mapel),
            ': ' + str(kelas),
            ': ' + str(tp),
            ': ' + str(KKTP)
        ]
    }
    combined_header_df = pd.DataFrame(header_data)

    # 2. TULIS DATA HEADER KE EXCEL
    for r_idx, row in combined_header_df.iterrows():
        worksheet.write(r_idx, 0, row['Keterangan'], header_info_format)
        worksheet.write(r_idx, 2, row['Nilai_Isian'], header_info_format)

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
            if column_name in ['NIS', 'NAMA SISWA', 'KELAS']:
                if isinstance(cell_value, (int, float)):
                    cell_value = str(cell_value)
                worksheet.write(row_num, col_num, cell_value, text_format)
            elif column_name.startswith('KTP-'):
                # Tulis nilai T/R statis (hasil perhitungan Python)
                worksheet.write_string(row_num, col_num, str(cell_value), border_format)
            else:
                worksheet.write(row_num, col_num, cell_value, border_format)

    # Set lebar kolom
    worksheet.set_column(1, 1, 35) # Nama Siswa
    worksheet.set_column(0, num_cols-1, 8) # Lebar default 8
    # Freeze pane to keep header and first two columns visible
    worksheet.freeze_panes(START_ROW_DATA + 1, 2)
    # Tidak perlu return output.getvalue() karena ini fungsi internal

def export_multisheet_form_nilai(df_all: pd.DataFrame, classes: List[str], mapel, semester, tp, guru, nip):
    """Menghasilkan file Excel multisheet untuk Form Nilai."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for kelas in classes:
            df_kelas = df_all[df_all['Kelas'] == kelas].reset_index(drop=True)
            if not df_kelas.empty:
                # Batasi nama sheet agar tidak melebihi 31 karakter
                sheet_name = f"{kelas} - Form Nilai"
                write_form_nilai_sheet(df_kelas, mapel, semester, kelas, tp, guru, nip, writer, sheet_name)
    return output.getvalue()

def export_multisheet_report_tk(df_all: pd.DataFrame, classes: List[str], mapel, tp):
    """Menghasilkan file Excel multisheet untuk Laporan TK."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for kelas in classes:
            df_kelas = df_all[df_all['Kelas'] == kelas].reset_index(drop=True)
            if not df_kelas.empty:
                sheet_name = f"{kelas} - Laporan TK"
                write_report_tk_sheet(df_kelas, mapel, kelas, tp, writer, sheet_name)
    return output.getvalue()

# =========================================================
# APLIKASI STREAMLIT UTAMA
# =========================================================
st.set_page_config(layout="wide", page_title="Editor Nilai Kurikulum Merdeka (LM x 5)")

# Injeksi CSS untuk gaya & freeze header + freeze kolom Nama
st.markdown(
    """
    <style>
    /* Buat tabel st.dataframe bisa scroll horz & vert */
    div[data-testid="stDataFrame"] {
        width: 100%;
        overflow: auto;
        height: 60vh;
    }
    /* Freeze Header Row */
    div[data-testid="stDataFrame"] thead tr:first-child th {
        position: sticky;
        top: 0;
        background-color: #ffffff;
        z-index: 5;
        border-bottom: 1px solid #ddd;
    }
    /* Freeze kolom NIS (kolom ke-1) */
    div[data-testid="stDataFrame"] tbody tr td:nth-child(1), div[data-testid="stDataFrame"] thead tr th:nth-child(1) {
        position: sticky;
        left: 0;
        background-color: #ffffff;
        z-index: 4;
    }
    /* Freeze kolom Nama (kolom ke-2) */
    div[data-testid="stDataFrame"] tbody tr td:nth-child(2), div[data-testid="stDataFrame"] thead tr th:nth-child(2) {
        position: sticky;
        left: 80px; /* Streamlit auto width — 80px offset works with default */
        background-color: #ffffff;
        z-index: 3;
    }
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
    "Unggah File CSV Data Siswa (Kolom wajib: NIS, Nama, Kelas)",
    type="csv"
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

            # Tambahkan kolom nilai input yang hilang (termasuk LM_1 s/d LM_5) dan inisialisasi dengan 0
            for col in INPUT_SCORE_COLS:
                if col not in df_all_students.columns:
                    df_all_students[col] = 0

            # Hapus kolom 'Nilai Rata-rata' dari CSV awal jika ada
            df_all_students = df_all_students.drop(columns=['Nilai Rata-rata'], errors='ignore')

            # Ensure numeric columns are integer type
            for col in INPUT_SCORE_COLS:
                df_all_students[col] = pd.to_numeric(df_all_students[col], errors='coerce').fillna(0).astype(int)
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
    available_classes = sorted(df_all_students['Kelas'].unique().tolist())
    if len(available_classes) == 0:
        available_classes = CLASS_OPTIONS

    kelas_input_list = st.multiselect("Pilih Kelas", available_classes, default=available_classes[:1])

with col3:
    semester_input = st.selectbox("Semester", ["Ganjil", "Genap"])

col4, col5, col6 = st.columns(3)
with col4:
    tahun_pelajaran_input = st.selectbox("Tahun Pelajaran", YEAR_OPTIONS, index=0)
with col5:
    guru_input = st.text_input("Nama Guru Mata Pelajaran", "")
with col6:
    nip_guru_input = st.text_input("NIP Guru", "")

st.write("---")
st.header("📝 Editor Nilai")

if not kelas_input_list:
    st.warning("Silakan pilih minimal satu kelas untuk memulai editor nilai.")
else:
    # Filter data berdasarkan kelas yang dipilih
    df_selected_classes = df_all_students[df_all_students['Kelas'].isin(kelas_input_list)].reset_index(drop=True)

    # 1. Hitung NR, Avg_TP, Avg_LM, Avg_PSA (Perhitungan Python)
    df_calculated = calculate_nr(df_selected_classes)

    # 2. Hitung Status TK (T/R)
    df_calculated = calculate_tk_status(df_calculated)

    # 3. Hitung Deskripsi NR
    df_calculated = generate_nr_description(df_calculated)

    # Kolom untuk tampilan editor
    DISPLAY_COLS = ['NIS', 'Nama', 'Kelas'] + INPUT_SCORE_COLS

    # Kolom untuk Data Editor (hanya kolom input)
    editable_cols = [c for c in INPUT_SCORE_COLS if c in df_calculated.columns]

    # Ambil kolom tampilan (termasuk hasil perhitungan Avg & NR untuk referensi)
    editor_display_cols = DISPLAY_COLS + [
        'Avg_TP', 'Avg_LM', 'Avg_PSA', 'NR', 'Deskripsi_NR'
    ] + [c for c in df_calculated.columns if c.startswith('TK_TP')]

    # Ganti nama kolom untuk tampilan Streamlit
    df_display = df_calculated[editor_display_cols].rename(columns=COLUMN_DISPLAY_MAP)

    # Atur tipe data untuk data editor (hanya kolom input)
    column_config_map = {
        'NIS': st.column_config.TextColumn("NIS", disabled=True),
        'Nama': st.column_config.TextColumn("NAMA SISWA", disabled=True),
        'Kelas': st.column_config.TextColumn("KELAS", disabled=True),
    }
    # Konfigurasi kolom input nilai
    for col in INPUT_SCORE_COLS:
        column_config_map[COLUMN_DISPLAY_MAP.get(col, col)] = st.column_config.NumberColumn(
            COLUMN_DISPLAY_MAP.get(col, col),
            help="Nilai antara 0-100",
            min_value=0,
            max_value=100,
            default=0,
            format="%d"
        )
    # Konfigurasi kolom hasil (harus disabled)
    for col in ['Rata-rata TP', 'Rata-rata LM', 'Rata-rata PSA', 'NR', 'Deskripsi Rapor'] + [COLUMN_DISPLAY_MAP.get(c, c) for c in df_calculated.columns if c.startswith('TK_TP')]:
        column_config_map[col] = st.column_config.TextColumn(col, disabled=True)

    st.subheader("Tabel Olah Nilai")
    st.info("Nilai di kolom **TP-n**, **LM-n**, **PTS**, dan **SAS/SAT** dapat diubah langsung. Kolom perhitungan (Avg & NR) akan diperbarui otomatis.")

    edited_df = st.data_editor(
        df_display,
        column_config=column_config_map,
        hide_index=True,
        key="data_editor_nilai"
    )

    # --- Bagian Ekspor ---
    # Jika data diedit, kembalikan ke format kolom asli
    df_export_edited = edited_df.rename(columns={v: k for k, v in COLUMN_DISPLAY_MAP.items()})

    # Konversi kolom input kembali ke int (editor mengkonversi ke float jika diedit)
    for col in INPUT_SCORE_COLS:
        if col in df_export_edited.columns:
            df_export_edited[col] = pd.to_numeric(df_export_edited[col], errors='coerce').fillna(0).astype(int)

    # Recalculate based on edited data
    df_export_calculated = calculate_nr(df_export_edited)
    df_export_calculated = calculate_tk_status(df_export_calculated)
    df_export_calculated = generate_nr_description(df_export_calculated)

    col_form, col_tk = st.columns(2)

    with col_form:
        excel_buffer_nilai = export_multisheet_form_nilai(
            df_export_calculated,
            kelas_input_list,
            mapel_terpilih,
            semester_input,
            tahun_pelajaran_input,
            guru_input,
            nip_guru_input
        )
        st.download_button(
            label="⬇️ Download Orek Orek Nilai Rapor** (Pilih Banyak Kelas)",
            file_name=f"Form_Nilai_Rapor_Multi_{mapel_terpilih}_{semester_input}.xlsx",
            data=excel_buffer_nilai,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_nilai"
        )

    with col_tk:
        excel_buffer_tk = export_multisheet_report_tk(
            df_export_calculated,
            kelas_input_list,
            mapel_terpilih,
            tahun_pelajaran_input
        )
        st.download_button(
            label="⬇️ Unduh **Hasil Olah Nilai Rapor dan TP** (Pilih Banyak Kelas)",
            file_name=f"Laporan_TP_Multi_{mapel_terpilih}_{semester_input}.xlsx",
            data=excel_buffer_tk,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_tk"
        )
