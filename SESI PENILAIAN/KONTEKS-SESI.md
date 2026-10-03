# KONTEKS SESI — PENILAIAN PROJEK ASTS "PERSONAL BRANDING CREATIVE CAMPAIGN"

> Context/handoff agar penilaian bisa dilanjutkan kapan saja tanpa mengulang verifikasi.
> Dibuat: **1 Oktober 2026** | Diperbarui: **3 Oktober 2026 (sore)** | Status: **SELESAI — 112 siswa dinilai**

---

## 0. ISI FOLDER INI

| Path | Isi |
|---|---|
| `..\REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx` | **FILE UTAMA.** 119 sheet: 112 `Detail - <nama>` + 7 sheet rekap |
| `..\LAPORAN PENILAIAN ASTS - Personal Branding XI DKV 2026-2027.md` | Laporan naratif, seksi per siswa (A ringkasan, B ketidaklengkapan, C/C2/C3/C4 per siswa, D cek manual, E ringkasan akhir) |
| `..\index.html` | Rekap interaktif **112 siswa** (objek `DATA` di-generate dari Excel) |
| `..\rekap_nilai_asts_hasil.json` | Data nilai machine-readable (112 entri) |
| `.\skrip\make_xlsx.py` | **SUMBER KEBENARAN nilai** |
| `.\skrip\tools\` | Toolchain penilaian (OCR RapidOCR, dossier, brief, entries) |
| `.\skrip\raw\` | HTML mentah per id (**berhati-hati**: id berubah nomor tiap fetch) |
| `.\README_CONTINUE.md` | **Catatan handoff utama — baca ini lebih dulu** |

---

## 1. KONTEKS PROYEK

| Item | Keterangan |
|---|---|
| Mata pelajaran | Pilihan DKV — **AI dalam Desain** |
| Kelas / Semester / TP | XI DKV 1–4 / 1 / 2026/2027 |
| Sekolah | **SMK Negeri 9 Garut** |
| Guru | Dede Alamsyah, S.Pd |
| Projek | **PROJEK ASTS — Personal Branding Creative Campaign berbasis AI** |
| Batas akhir | Jumat, 2 Oktober 2026 23.59 WIB |
| Total Pengumpul | **112 siswa** |
| Sumber rekapan | Google Form → `https://docs.google.com/spreadsheets/d/1eq2UyXOYX8z60ybOxGiSa9CLhDQRrRrhQViKrzZjIoU` |

---

## 2. HASIL PENILAIAN (3 OKTOBER 2026)

- **Total dinilai**: **112 siswa**
- **Rata-rata**: rubrik **78,53** + nilai tambah resolusi **0,09** = **78,62**
- **Kategori**: Sangat Baik 33 · Baik 51 · Cukup 15 · Perlu Perbaikan 13
- **Kelas**: DKV 1 = 27 · DKV 2 = 26 · DKV 3 = 31 · DKV 4 = 28
- **8 siswa bernilai 0** karena karya tidak dapat diverifikasi:
  - URL *editor* Blogger — CEISHA SINTHIA, SINDIA SAPUTRI, SITI JENAB, NAZWA KURNIA,
    LUSI NURAENI, RADIT KURNIAWAN
  - HTTP 404 — RISMA SAPARANI, ALIA ALAIKA NURFADILA
- **Perubahan besar pada batch terakhir**:
  - **NURI MEITRI AENI** 0 → **75,50** (link aktif, domain `nurimeiaeni.blogspot.com`)
  - **QUINSYA RAHMANESA SOLEHA** 52,50 → **92,75**
  - **DHEA EKA KHOERUNNISA** 58,50 → **87,00**
  - **INDRI FITRIYANI** 76,50 → **90,00**
  - **SAFINAH SYARA GARINI** 89,50 → **95,00**
  - **WULAN SUNDARI** 79,50 → **87,75** (dinilai dari arsip; link publik kini 404)
  - **NADIA FITRIANI** 88,25 → **87,50**, **SITI RAHMA SILPIANA** 88,50 → **91,00**

---

## 3. TUGAS BERIKUTNYA

1. **Minta link** dari 8 siswa bernilai 0.
2. **Periksa 38 butir** sheet `4. CEK MANUAL GURU`.
3. **Temuan yang perlu klarifikasi** (lihat `README_CONTINUE.md` bagian "Temuan"):
   RAFI FAUZAN (projek salah), DAPA MUSTOPA/DAFA MUSTOFA (nama), MUHAMAD REZA (identitas brand),
   AQILA (gambar privat), MUHAMMAD TAUFIQ (sisa teks AI), SAVINA (shotlist kosong),
   SAFINAH (duplikasi).
4. **Ekspor laporan** PDF/cetak bila perlu.

---

*Sumber kebenaran angka: `SESI PENILAIAN/skrip/make_xlsx.py`.*