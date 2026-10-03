# KONTEKS SESI — PENILAIAN PROJEK ASTS "PERSONAL BRANDING CREATIVE CAMPAIGN"

> Dokumen ini adalah **context/handoff** agar penilaian ini bisa dilanjutkan kapan saja
> tanpa mengulang seluruh proses verifikasi.
> Dibuat: **1 Oktober 2026** | Diperbarui: **3 Oktober 2026** | Status: **SELESAI — 86 siswa dinilai**

---

## 0. ISI FOLDER INI

| Path | Isi |
|---|---|
| `..\REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx` | **FILE UTAMA.** 93 sheet: rekap nilai, nilai komponen, 86 sheet detail siswa, rekap ketidaklengkapan (330 temuan), cek manual guru (27 butir), data rekapan, nilai tambah resolusi, metode & rubrik |
| `..\LAPORAN PENILAIAN ASTS - Personal Branding XI DKV 2026-2027.md` | Laporan naratif lengkap, per siswa + feedback |
| `..\index.html` | Versi HTML yang menampilkan rekap interaktif **seluruh 86 siswa** |
| `..\rekap_nilai_asts_hasil.json` | Data nilai machine-readable (86 entri) |
| `.\skrip\` | Semua skrip yang dipakai (fetch → parse → OCR → hitung → Excel) |
| `.\skrip\tools\BRIEF-PENILAIAN.md` | Brief penilaian untuk AI batch (arsip — batch 2 Okt sudah selesai) |
| `.\data\` | Hasil mentah: `parsed.json`, `images.json`, `ocr_compact2.txt`, `nilai_akhir.json` |
| `.\README_CONTINUE.md` | **Catatan handoff utama** — baca ini lebih dulu |

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
| Total Pengumpul | **86 siswa** |
| Sumber rekapan | Google Form → `https://docs.google.com/spreadsheets/d/1eq2UyXOYX8z60ybOxGiSa9CLhDQRrRrhQViKrzZjIoU` |

---

## 2. HASIL PENILAIAN TERKINI (3 OKTOBER 2026)

- **Total Terverifikasi**: 86 siswa
- **Rata-rata Nilai Akhir**: **79,35** (rubrik 79,28 + nilai tambah resolusi 0,07)
- **Kategori**: Sangat Baik 23 · Baik 44 · Cukup 10 · Perlu Perbaikan 9
- **Sebar kelas**: DKV 1 = 18 · DKV 2 = 22 · DKV 3 = 25 · DKV 4 = 21
- **Nilai 0 (6 siswa)**, semua karena karya **tidak dapat diverifikasi**:
  - **URL *editor* Blogger** (butuh login): CEISHA SINTHIA, SINDIA SAPUTRI, SITI JENAB, NAZWA KURNIA
  - **HTTP 404**: RISMA SAPARANI, NURI MEITRI AENI
- **Update terakhir (3 Oktober 2026)**:
  1. Audit konsistensi 4 sumber data (`make_xlsx.py`, `parsed.json`, `.xlsx`, `index.html`) — 0 selisih pada nilai.
  2. `make_xlsx.py`: teks sheet `6. METODE & RUBRIK` diperbarui (86 kiriman, 555 berkas gambar),
     status akses nyata pada sheet `5. DATA REKAPAN`, 10 butir baru di `4. CEK MANUAL GURU` (17 → 27),
     2 key `RESOLUSI` yang hilang ditambahkan, key duplikat dihapus, `OUT` dibuat absolut.
  3. `.xlsx` diregenerasi (93 sheet) dan `index.html` disinkronkan via `docs/update_data.py`.
  4. `LAPORAN *.md`: 18 seksi siswa yang hilang ditambahkan (bagian C3), seluruh tabel skor
     86 seksi diselaraskan, tabel ringkasan A ditulis ulang (68 → 86 baris), bagian D & E dihitung ulang.
  5. `rekap_nilai_asts_hasil.json` diregenerasi (68 → 86 entri, 16 nilai basi dikoreksi).

---

## 3. TUGAS/TINDAK LANJUT BERIKUTNYA

1. **Minta link artikel** dari 6 siswa bernilai 0:
   - URL editor Blogger → **CEISHA SINTHIA**, **SINDIA SAPUTRI**, **SITI JENAB**, **NAZWA KURNIA**
   - Link 404 → **RISMA SAPARANI**, **NURI MEITRI AENI**
2. **Periksa orisinalitas 4 kiriman** — heading `{Monogram SSG}` muncul identik pada
   **MUHAMAD DIAZ PIRDAUS**, **NAZMA KAYVA GASANI**, **TENI DAMAYANTI**, **RESTI NURUL FADILA**.
3. **Cek Manual Guru**: tinjau **27 butir** pada sheet `4. CEK MANUAL GURU`.
4. **Ekspor Laporan**: buat PDF/laporan cetak bila diperlukan untuk pembagian hasil ASTS.

---

*Dokumen ini diperbarui mengikuti kondisi `make_xlsx.py`. Sumber kebenaran angka tetap
`SESI PENILAIAN/skrip/make_xlsx.py`.*