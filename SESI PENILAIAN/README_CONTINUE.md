# Continuation Note

State of the ASTS assessment workflow, so the AI can continue without re-loading all data.

**Diperbarui: 3 Oktober 2026 — sinkron dengan 86 siswa, rata-rata 79,35.**
Versi sebelum 2 Oktober sudah usang: rubric bukan lagi 36/68 siswa.

## Lokasi
- Skrip penilaian: `SESI PENILAIAN/skrip/`
- Sumber kebenaran nilai: `SESI PENILAIAN/skrip/make_xlsx.py` (`S.append` + `RESOLUSI`)
- Data fetch: `SESI PENILAIAN/skrip/parsed.json` (**86 entri**), baseline identik di `SESI PENILAIAN/data/parsed.json`
- HTML mentah: `SESI PENILAIAN/skrip/raw/R*.html` (111 berkas)
- Bukti baca: `SESI PENILAIAN/skrip/reports/R*.txt` (86 berkas)
- Output: `REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx` (**93 sheet**)
- Turunan: `LAPORAN PENILAIAN ... .md`, `rekap_nilai_asts_hasil.json`, `index.html`

## Sumber data
- Spreadsheet Form Responses: `https://docs.google.com/spreadsheets/d/1eq2UyXOYX8z60ybOxGiSa9CLhDQRrRrhQViKrzZjIoU`
- `fetch_parse.py` memakai ekspor **xlsx** (`?export?format=xlsx`), bukan CSV.

## Status per 03/10/2026
- **86 siswa** dinilai penuh. Rata-rata **rubrik 79,28 + nilai tambah resolusi 0,07 = 79,35**.
- Kategori: Sangat Baik 23 · Baik 44 · Cukup 10 · Perlu Perbaikan 9.
- Sebar kelas: XI DKV 1 = 18 · DKV 2 = 22 · DKV 3 = 25 · DKV 4 = 21.
- **93 sheet**: 86 `Detail - <nama>` + `1. REKAP NILAI`, `2. NILAI KOMPONEN`,
  `3. REKAP KETIDAKLENGKAPAN` (330 temuan), `4. CEK MANUAL GURU` (**27 butir**),
  `5. DATA REKAPAN`, `7. NILAI TAMBAHAN RESOLUSI`, `6. METODE & RUBRIK`.
- **6 siswa bernilai 0** (karya tidak dapat diverifikasi, bukan karya kosong):
  - URL *editor* Blogger → CEISHA SINTHIA, SINDIA SAPUTRI, SITI JENAB, NAZWA KURNIA
  - HTTP 404 → RISMA SAPARANI, NURI MEITRI AENI
- **27 butir `4. CEK MANUAL GURU`**: 17 butir verifikasi visual lama + 4 butir orisinalitas
  heading `{Monogram SSG}` (MUHAMAD DIAZ, NAZMA KAYVA, TENI DAMAYANTI, RESTI NURUL FADILA)
  + 6 permintaan link publik/aktif.

## PENTING: cakupan make_xlsx.py
- `make_xlsx.py` **tidak membaca** `parsed.json`. Seluruh rubrik (12 komponen × 86 siswa)
  ditulis literal sebagai `S.append((...))` + `RESOLUSI`.
- Id `R##` pada `S.append` **tidak unik** dan tidak dirujuk apa pun — jangan memakainya
  sebagai kunci lookup. `TIDAK_AKSES` dan `CATATAN` sengaja di-key **berdasarkan nama**.
- `OUT` dihitung absolut dari lokasi skrip, jadi aman dijalankan dari direktori mana pun.

## Alur pemeliharaan
```bash
# 1. nilai baru: baca reports/R*.txt + raw/R*.html, lalu tambah S.append((...)) di make_xlsx.py
python "SESI PENILAIAN/skrip/make_xlsx.py"   # regenerate Excel (93 sheet)
python docs/update_data.py --check            # lihat apa yang berubah, tanpa menulis
python docs/update_data.py                    # tulis DATA ke index.html
git add -A && git commit -m "..." && git push origin main
```
Bila ada kiriman baru dari Form Responses:
```bash
cp "SESI PENILAIAN/skrip/parsed.json" "SESI PENILAIAN/data/parsed.json"  # baseline dulu
python "SESI PENILAIAN/skrip/fetch_parse.py"
python "SESI PENILAIAN/skrip/report.py"
```

## Urutan sinkronisasi 4 berkas (jangan sampaiidah)
`make_xlsx.py` → `.xlsx` → `docs/update_data.py` → `index.html`.
`LAPORAN *.md` dan `rekap_nilai_asts_hasil.json` **tidak** dibangun otomatis — keduanya
perlu diperbarui manual dari `make_xlsx.py` setiap kali nilai berubah.

## Jebakan yang sudah pernah terjadi
- **Baseline `data/parsed.json` tertinggal.** Kalau `fetch_parse.py` dijalankan saat
  baseline lebih tua dari `skrip/parsed.json`, ~(selisih) siswa terbaca sebagai "kiriman baru"
  yang palsu. Salin dulu sebelum fetch.
- **Duplikat nama.** DIRA RAHMAWATI sempat dikeluarkan lalu muncul lagi; INDRI FITRIYANI
  sempat tercatat dua kali. Halaman web mencari data **berdasarkan nama** (`norm()`), jadi
  nama harus unik.
- **Id R## bukan kunci.** Gunakan nama lengkap.

## Catatan teknis
- `docs/update_data.py` hanya menyalin ulang objek `DATA`; HTML/JS lain tidak tersentuh.
  Ia juga menulis ulang angka ringkasan (`meta.rata`, `meta.gambar`).
- `git push` ke repo `dedealamsyah-sch/...` gagal 403 bila akun aktif `gh` adalah
  `dedealamsyah`. Pakai `gh auth switch --user dedealamsyah-sch` sebelum push, lalu kembalikan.
- Output python yang panjang kadang tertangkap hook; tulis ke file lalu baca bila perlu.
- `.gitignore` **tidak** mengecualikan apa pun — seluruh isi project (termasuk `.xlsx`,
  `SESI PENILAIAN/`, gambar) ikut ter-*commit* ke repo.

## Tugas berikutnya
1. **Minta link** dari 6 siswa bernilai 0 (lihat daftar di atas).
2. **Tinjau 27 butir** sheet `4. CEK MANUAL GURU`.
3. **Periksa orisinalitas** 4 kiriman dengan heading `{Monogram SSG}`.
4. **Ekspor laporan** PDF/cetak bila perlu untuk pembagian hasil ASTS.