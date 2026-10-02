# Continuation Note

State of the ASTS assessment workflow, so the AI can continue without re-loading all data.

**Diperbarui: 2 Oktober 2026 (pukul 10.30 WIB) — sinkron dengan commit `2914cb9`.**
Versi sebelum 1 Oktober sudah usang: rubric bukan lagi 36 siswa, dan tidak ada siswa
yang belum dinilai.

## Lokasi
- Root: `/Users/dedealamsyah/Instructor/SAGAR/DKV/2026/MPP AI - XI DKV/ASTS/NILAI ASTS 1/`
- Skrip: `SESI PENILAIAN/skrip/` (`fetch_parse.py`, `report.py`, `make_xlsx.py`, `hitung.py`, `getimg.py`)
- Data: `SESI PENILAIAN/skrip/parsed.json` (hasil fetch terbaru, **68 entri**),
  `SESI PENILAIAN/data/parsed.json` (baseline, **masih 45 entri** — lihat "Jebakan baseline")
- HTML mentah: `SESI PENILAIAN/skrip/raw/R*.html` (67 berkas)
- Bukti baca: `SESI PENILAIAN/skrip/reports/R*.txt` (68 berkas, otomatis dari `report.py`)
- Output: `REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx` (**75 sheet**)
- Backup workbook: `SESI PENILAIAN/backup/REKAP NILAIAN ASTS - 69 baris (sebelum hapus duplikat).xlsx`
- Web: `index.html` (semua data dibaca dari objek JS `DATA`)

## Sumber data
- Spreadsheet Form Responses: `https://docs.google.com/spreadsheets/d/1eq2UyXOYX8z60ybOxGiSa9CLhDQRrRrhQViKrzZjIoU`
- `fetch_parse.py` memakai ekspor **xlsx** (`?export?format=xlsx`), bukan CSV.

## Status per 02/10/2026
- **68 siswa** dinilai penuh di `1. REKAP NILAI` (No 1–68, tanpa nomor lompatan).
- **75 sheet**: 68 `Detail - <nama>` + `1. REKAP NILAI`, `2. NILAI KOMPONEN`,
  `3. REKAP KETIDAKLENGKAPAN`, `4. CEK MANUAL GURU`, `5. DATA REKAPAN`,
  `7. NILAI TAMBAHAN RESOLUSI`, `6. METODE & RUBRIK`.
- **Rata-rata Nilai Akhir 78,95** (rubrik 78,86 + nilai tambah resolusi 0,09), 440 berkas gambar.
- Sebar per kelas: XI DKV 1 = 13, XI DKV 2 = 19, XI DKV 3 = 24, XI DKV 4 = 12.
- Kategori: Sangat Baik 16 · Baik 34 · Cukup 10 · Perlu Perbaikan 8.
- **Nilai 0 (4 siswa, semuanya XI DKV 2)**: CEISHA SINTHIA (No 65), SINDIA SAPUTRI (No 66),
  SITI JENAB (No 67) — link-nya berupa halaman *editor* Blogger sehingga tidak dapat diverifikasi;
  RISMA SAPARANI (No 68) — link publik aktif, tetapi tidak ada visual/prompt yang dapat diperiksa.
- 2 Oktober: duplikat **INDRI FITRIYANI** (satu orang, link `...fitriyani-xi.html?m=1` = versi
  mobile dari kiriman utama) dihapus dari `make_xlsx.py`. Sheet `Detail - Indri Fitriyani1` hilang,
  jumlah siswa 69 → 68, rata-rata 79,11 → 78,95. Skor yang dipakai: rubrik **90,5**.
- Web sudah sinkron dengan Excel (`python3 docs/update_data.py`, 0 selisih di semua baris),
  sudah di-commit & push (`2914cb9`).

## PENTING: cakupan make_xlsx.py
- `make_xlsx.py` **tidak membaca** `parsed.json`. Seluruh penilaian rubric (12 komponen × 68 siswa)
  ditulis manual sebagai literal `S.append((...))`, lengkap dengan narasi bukti per komponen.
- **68 entri `S.append`** = 68 siswa. Regenerasi **sudah diverifikasi identik (0 sel berbeda)**
  pada 02/10/2026, jadi aman selama `S` tidak diubah manual.
- Id `R##` pada `S.append` **tidak unik** (R16, R17, R19, R20, R21 dipakai dua kali) dan tidak
  dirujuk apa pun — jangan memakainya sebagai kunci lookup.
- Archives/kelengkapan nilai tetap di `index.html` dibaca dari Excel, bukan dari skrip.

## Jebakan baseline
`SESI PENILAIAN/data/parsed.json` masih berisi **45 entri**, sedangkan `skrip/parsed.json` berisi 68.
Kalau `fetch_parse.py` dijalankan sekarang, ~23 siswa akan terbaca sebagai "kiriman baru" yang
palsu. Salin `skrip/parsed.json` ke `data/parsed.json` lebih dulu bila ingin deteksi perubahan akurat.

## Yang belum dinilai
**Tidak ada.** Seluruh 68 siswa sudah punya nilai rubrik + nilai akhir di Excel maupun `index.html`.
Yang masih menunggu keputusan guru ada di bagian "Tugas التالية".

## Tugas berikutnya
1. **Konfirmasi 4 siswa nilai 0**: minta link artikel publik untuk CEISHA SINTHIA, SINDIA SAPUTRI,
   SITI JENAB (link editor), dan klarifikasi apakah RISMA SAPARANI perlu di nilai ulang.
2. **Cek manual guru**: tinjau butir pada sheet `4. CEK MANUAL GURU` (17 butir).
3. **Ekspor laporan**: buat PDF/laporan cetak bila diperlukan untuk pembagian hasil ASTS.

## Alur pemeliharaan (kalau ada perubahan nilai)
```bash
# 1. nilai baru: baca reports/R*.txt + raw/R*.html, lalu tambah S.append((...)) di make_xlsx.py
python3 "SESI PENILAIAN/skrip/make_xlsx.py"     # regenerate Excel (68 -> N siswa)
python3 docs/update_data.py --check              # lihat apa yang berubah, tanpa menulis
python3 docs/update_data.py                      # tulis DATA ke index.html
git add index.html && git commit -m "..." && git push origin main
```
Bila kiriman baru dari Form Responses:
```bash
cd "SESI PENILAIAN/skrip"
cp parsed.json ../data/parsed.json               # baseline dulu (lihat "Jebakan baseline")
python3 fetch_parse.py                           # deteksi kiriman baru / berubah
python3 report.py                                # bukti baca untuk siswa baru
```

## Catatan teknis
- `docs/update_data.py` hanya menyalin ulang objek `DATA`; HTML/JS lain tidak tersentuh.
  Ia juga menulis ulang angka ringkasan pada `6. METODE & RUBRIK` (jumlah kiriman & berkas gambar).
- Halaman web mencari data **berdasarkan nama** (`norm()`). Nama siswa harus unik — duplikat nama
  membuat baris kedua memakai angka baris pertama dan membuat `detail` tidak terpetakan.
- `git push` ke repo `dedealamsyah-sch/...` gagal 403 bila akun aktif `gh` adalah `dedealamsyah`.
  Pakai `gh auth switch --user dedealamsyah-sch` sebelum push, lalu kembalikan.
- Output python yang panjang kadang tertangkap hook; tulis ke file lalu baca bila perlu
  (`python3 skrip.py > /tmp/out.txt`).
- `.xlsx`, `LAPORAN *.md`, `rekap_nilai_asts_hasil.json`, dan `SESI PENILAIAN/` tidak ikut
  ter-upload ke GitHub (`.gitignore`); hanya `index.html`, `logo-dkv.png`, dan `docs/`.