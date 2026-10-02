# KONTEKS SESI — PENILAIAN PROJEK ASTS "PERSONAL BRANDING CREATIVE CAMPAIGN"

> Dokumen ini adalah **context/handoff** agar penilaian ini bisa dilanjutkan kapan saja
> tanpa mengulang seluruh proses verifikasi.
> Dibuat: **1 Oktober 2026** | Diperbarui: **2 Oktober 2026** | Status: **SELESAI — 68 siswa dinilai**

---

## 0. ISI FOLDER INI

| Path | Isi |
|---|---|
| `..\REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx` | **FILE UTAMA.** 75 sheet: rekap nilai, nilai komponen, 68 sheet detail siswa, rekap ketidaklengkapan, cek manual guru, data rekapan, nilai tambah resolusi, metode & rubrik |
| `..\LAPORAN PENILAIAN ASTS - Personal Branding XI DKV 2026-2027.md` | Laporan naratif lengkap, per siswa + feedback |
| `..\index.html` | Versi HTML yang menampilkan rekap interaktif **seluruh 68 siswa** |
| `..\rekap_nilai_asts_hasil.json` | Data nilai (machine-readable) |
| `.\skrip\` | Semua skrip yang dipakai (fetch → parse → OCR → hitung → Excel) |
| `.\data\` | Hasil mentah: `parsed.json`, `images.json`, `ocr_compact2.txt`, `nilai_akhir.json` |

---

## 1. KONTEKS PROYEK

| Item | Keterangan |
|---|---|
| Mata pelajaran | Pilihan DKV — **AI dalam Desain** |
| Kelas | XI DKV 1, 2, 3, 4 |
| Semester | 1 |
| TP | 2026/2027 |
| Sekolah | **SMK Negeri 9 Garut** |
| Guru | Dede Alamsyah, S.Pd |
| Projek | **PROJEK ASTS — Personal Branding Creative Campaign berbasis AI** |
| Batas akhir | Jumat, 2 Oktober 2026 23.59 WIB |
| Total Pengumpul | **68 siswa** (sebelumnya 69; 1 duplikat INDRI FITRIYANI dihapus pada 2 Oktober 2026) |
| Sumber rekapan | Google Form → `https://docs.google.com/spreadsheets/d/1eq2UyXOYX8z60ybOxGiSa9CLhDQRrRrhQViKrzZjIoU` |

---

## 2. HASIL PENILAIAN TERKINI (02 OKTOBER 2026)

- **Total Terverifikasi**: 68 Siswa
- **Rata-rata Nilai Akhir Kelas**: **78,95** (Rubrik 78,86 + Nilai Tambah Resolusi 0,09)
- **Nilai 0 (4 siswa, semuanya XI DKV 2)**: CEISHA SINTHIA (No 65), SINDIA SAPUTRI (No 66),
  SITI JENAB (No 67) — link Blogger mengarah ke halaman Editor (butuh login);
  RISMA SAPARANI (No 68) — link publik aktif, tetapi tidak ada visual/prompt yang dapat diperiksa.
- **Update Terakhir**:
  1. Penilaian otomatis & manual untuk 32 siswa susulan diselesaikan di `SESI PENILAIAN/skrip/make_xlsx.py`.
  2. Excel rekap (`REKAP NILAIAN ASTS...xlsx`) dibuat ulang mencakup seluruh 75 sheet.
  3. Duplikat INDRI FITRIYANI dihapus dari `make_xlsx.py` → 68 siswa, rata-rata 78,95.
  4. Perubahan disinkronkan sepenuhnya ke `index.html` menggunakan `docs/update_data.py`.
  5. Perubahan telah di-commit & di-push ke repository GitHub (`main` branch, commit `2914cb9`).

---

## 3. TUGAS/TINDAK LANJUT BERIKUTNYA

1. **Konfirmasi Link Broken**:
   - Meminta link artikel publik untuk 3 siswa yang terindikasi mengirim URL Editor Blogger:
     - **CEISHA SINTHIA** (XI DKV 2, No 65)
     - **SINDIA SAPUTRI** (XI DKV 2, No 66)
     - **SITI JENAB** (XI DKV 2, No 67)
   - **RISMA SAPARANI** (XI DKV 2, No 68): link publik, tetapi isi tidak dapat diverifikasi —
     perlu diputuskan apakah di nilai ulang atau tetap 0.
2. **Cek Manual Guru**:
   - Meninjau butir-butir pada sheet `4. CEK MANUAL GURU` (17 butir).
3. **Ekspor Laporan**:
   - Menghasilkan laporan cetak/PDF akhir bila diperlukan untuk pembagian hasil ASTS.

---
*Dokumen ini diperbarui secara otomatis setelah pembaruan data rekap dan integrasi index.html.*
