# Continuation Note

State of the ASTS assessment workflow, so the AI can continue without re-loading all data.

**Diperbarui: 3 Oktober 2026 (batch sore) — sinkron dengan 112 siswa, rata-rata 78,62.**

## Lokasi
- Skrip penilaian: `SESI PENILAIAN/skrip/`
- Sumber kebenaran nilai: `SESI PENILAIAN/skrip/make_xlsx.py` (`S.append` + `RESOLUSI`)
- Data fetch: `SESI PENILAIAN/skrip/parsed.json` (**112 entri**)
- Bukti assess: `SESI PENILAIAN/skrip/tools/out/dossier/<RID>.txt`
- Output: `REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx` (**119 sheet**)
- Turunan: `LAPORAN *.md`, `rekap_nilai_asts_hasil.json`, `index.html`

## Toolchain (SUDAH BERUBAH — baca dulu)
 assessments sebelumnya memakai OCR Apple Vision (macOS). Di platform Windows itu tidak
 tersedia, dan **AI pun tidak memiliki vision** (gambar tidak pernah sampai ke model).
 Verifikasi visual kini memakai:

| Tool | Fungsi |
|---|---|
| `tools/ocr_batch.py` | OCR batch via **RapidOCR/ONNX** (python, lintas-platform). Cache: `tools/out/ocr_cache_win.json` |
| `tools/dossier_win.py` | Dossier per siswa: teks per bagian + tabel + `[IMG#nn] px=` + baris `OCR:` |
| `tools/dossier_arsip.py` | Dossier dari HTML terarsip (untuk siswa yang link publiknya mati) |
| `tools/BRIEF-BATCH.md` | Brief penilaian untuk sub-agent |
| `tools/out/entries_final.py` | 38 entry hasil penilaian batch terakhir (sudah digabung ke `make_xlsx.py`) |

Skrip `ocr` / `ocrcrop` (binary Mach-O) **tidak bisa dijalankan** di sini.

## Status per 03/10/2026 (sore)
- **112 siswa** dinilai. Rata-rata **rubrik 78,53 + nilai tambah 0,09 = 78,62**.
- Kategori: Sangat Baik 33 · Baik 51 · Cukup 15 · Perlu Perbaikan 13.
- Kelas: DKV 1 = 27 · DKV 2 = 26 · DKV 3 = 31 · DKV 4 = 28.
- **119 sheet** = 112 `Detail - <nama>` + 7 sheet rekap.
- `3. REKAP KETIDAKLENGKAPAN` = **484 temuan**. `4. CEK MANUAL GURU` = **38 butir**.
- **8 siswa bernilai 0** (karya tidak dapat diverifikasi):
  - URL *editor* Blogger (6): CEISHA SINTHIA, SINDIA SAPUTRI, SITI JENAB, NAZWA KURNIA,
    LUSI NURAENI, RADIT KURNIAWAN
  - HTTP 404 (2): RISMA SAPARANI, ALIA ALAIKA NURFADILA
- **NURI MEITRI AENI** sebelumnya 0 (404) → kini **75,50** (link aktif, domain
  `nurimeiaeni.blogspot.com` — domain lama `nurimeitriii` masih 404).
- **WULAN SUNDARI** linknya kini 404, tetapi karya dipulihkan dari arsip HTML commit
  `bcdf815` → **87,75** (bukan 0).

## Jebakan yang sudah diperbaiki (jangan dibalik!)
1. **Id `R##` berubah nomor** tiap `fetch_parse.py` (diurutkan ulang dari spreadsheet).
   Folder `img/R##` bisa berisi siswa lain. Selalu pakai **nama** sebagai kunci.
2. **`fetch_parse.py` menimpa `raw/<id>.html` dengan halaman gagal.** Sudah diperbaiki:
   hasil fetch yang gagal kini disimpan ke `raw/<id>.html.arsip` dan TIDAK menimpa bukti.
   Tetap periksa `git log -- raw/<id>.html` bilaPROOF hilang.
3. **`getimg_sel.py` menimpa `images.json`.** Sudah diperbaiki → merge.
4. **Validasi isi gambar, bukan ukuran.** `lh3.google.com` (Google Photos privat)
   membalas HTML 1,2 MB dengan HTTP 200. `is_image()` memeriksa magic bytes.
5. **Ekstraksi teks hanya `<p>`/`<li>` under-report hingga 98%** (siswa menulis di
   `<div>`/`<span>`). `dossier_win.py` kini menyertakan blok
   `TEKS ARTIKEL LENGKAP (get_text() penuh)` bila per-bagian < 70% `word_count`.
6. **Nama yang jadi awalan nama lain** ("INDRI" vs "INDRI FITRIYANI") bikin sheet Excel
   gagal dipetakan di `docs/update_data.py`. Sudah diperbaiki: cocokkan persis dulu.

## Temuan yang perlu keputusan guru
1. **RAFI FAUZAN NAJA LUTFIANA** mengirim ID Card + banner rental PS + logo Android
   (CorelDRAW) — **bukan projek Personal Branding**. Nilai 4,75.
2. **DAPA MUSTOPA** — rekapan & page title "DAPA MUSTOPA", seluruh isi & gambar
   "DAFA MUSTOFA". Perlu verifikasi identitas.
3. **MUHAMAD REZA RAMDANI** — identitas brand berubah 4× dalam satu artikel
   (logo "RR" → moodboard branding sekolah → "The Creative Studio"/Ahmad Faisal).
4. **AQILA NAZIL FALAQ** — 7 gambar di album Google Photos privat, tidak bisa diunduh.
5. **MUHAMMAD TAUFIQ ISMAIL** — sisa teks mentah AI (`( IndiBlogHub )`, `( Markuva )`).
6. **SAVINA KHOERUNNISA** — bagian SHOTLIST & STORYBOARD berisi deskripsi logo, bukan
   shot/scene.
7. **SAFINAH SYARA GARINI** — bagian mockup "Kemasan" duplikat identik dengan "Kartu Nama".
8. **Orisinalitas** — heading `{Monogram SSG}` identik di 4 kiriman (lama).
9. **8 siswa nilai 0** — minta link publik/aktif.

## Alur pemeliharaan
```bash
cp "SESI PENILAIAN/skrip/parsed.json" "SESI PENILAIAN/skrip/parsed_snap.json"  # opsional
python "SESI PENILAIAN/skrip/fetch_parse.py"
python "SESI PENILAIAN/skrip/getimg_sel.py <ID...>     # merge, validasi magic bytes
python "SESI PENILAIAN/skrip/tools/ocr_batch.py <ID...>
python "SESI PENILAIAN/skrip/tools/dossier_win.py <ID...>
# ... penilaian ... -> tools/out/entries_*.py, lalu gabung ke make_xlsx.py
python "SESI PENILAIAN/skrip/make_xlsx.py"
python docs/update_data.py
git add -A && git commit -m "..." && git push origin main
```

## Catatan gambar (PENTING)
Pada 3 Oktober 2026 folder `skrip/img/` dipangkas dari 806 → 337 file (131 MB → 45 MB).
Yang dihapus adalah gambar **yatim**: tidak terdaftar di `images.json` karena menumpang
di folder `R##` yang nomornya sudah bergeser, sehingga isinya milik siswa lain dan tidak
dapat dipercaya.

- **Nilai tidak bergantung pada gambar.** `RESOLUSI` di `make_xlsx.py` sudah berupa angka
  literal hasil pengukuran; tidak ada dokumen hasil yang menautkan `.jpg`.
- Gambar **16 siswa** dengan butir cek visual di sheet `4. CEK MANUAL GURU` **dipertahankan**
  (AHMAD FAUZI, AQILA NAZIL FALAQ, DEDE APRILIA KARTIKA, DHEA EKA KHOERUNNISA, ILMA LATIFAH,
  INTAN WIDIYANTI, JIHAN SHAFIRA KEAN PUTRI MULYADI, PUTRI INTAN NURAENI, QIANDRA KAIZAR NAHARI,
  QUINSYA RAHMANESA SOLEHA, SAFINAH SYARA GARINI, SAVINA KHOERUNNISA, SOPA ANIDATUL AISAH,
  WAHDAN SAPARI, WILDA AZKIA, WULAN SUNDARI). Jangan hapus sebelum butirnya selesai diperiksa.
- Gambar siswa nilai 0 (CEISHA SINTHIA, SINDIA SAPUTRI, SITI JENAB, NAZWA KURNIA) ikut
  terhapus — butir CEK mereka hanya "kirim link publik". Bisa diunduh ulang bila perlu.
- **Masih dapat dipulihkan** dari history: `git checkout <commit-sebelum> -- "SESI PENILAIAN/skrip/img"`.
- Menghapus gambar **tidak** mengecilkan `.git`; itu masalah terpisah (lihat di bawah).

## Catatan teknis
- `docs/update_data.py` hanya menyalin objek `DATA`; HTML/JS lain tidak tersentuh.
- `.gitignore` **tidak** mengecualikan apa pun — gambar & `.xlsx` ikut ter-*commit*.
- Output python panjang kadang tertangkap hook → tulis ke file lalu baca.
- `git push` gagal 403 bila akun aktif `gh` adalah `dedealamsyah`;
  pakai `gh auth switch --user dedealamsyah-sch`.