import csv

def baca_data_siswa():
    try:
        with open('siswa.csv', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            data = list(reader)
        return data
    except FileNotFoundError:
        print("File siswa.csv tidak ditemukan.")
        return []

def tampilkan_daftar_per_kelas(kelas):
    siswa = baca_data_siswa()
    data_kelas = [s for s in siswa if s['Kelas'] == kelas]
    if not data_kelas:
        print(f"Tidak ada data untuk kelas {kelas}")
        return
    print(f"\nDaftar Siswa Kelas {kelas}:\n")
    print(f"{'No':<5} {'Nama':<25} {'NIS':<10}")
    print("=" * 45)
    for s in data_kelas:
        print(f"{s['No']:<5} {s['Nama']:<25} {s['NIS']:<10}")

if __name__ == "__main__":
    print("=== Sistem Informasi Daftar Siswa ===")
    kelas = input("Masukkan kelas yang ingin ditampilkan (contoh: 7A): ")
    tampilkan_daftar_per_kelas(kelas)
