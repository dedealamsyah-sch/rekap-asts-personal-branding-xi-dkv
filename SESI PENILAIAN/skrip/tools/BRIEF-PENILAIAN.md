# BRIEF PENILAIAN — 27 SISWA (2 Oktober 2026)

Kamu menilai karya ASTS "Personal Branding Creative Campaign" untuk siswa yang **baru mengirim/
mengubah** kiriman. Semua data sudah di-fetch (lihat `dossier/<RID>.txt`).

## WAJIB DIBACA DULU
1. `SESI PENILAIAN/KETENTUAN-AI.md` → bab 3.2 (skala 0-4), bab 4 (checklist ketentuan per komponen),
   bab 5 (metodologi verifikasi/atasan). Baca juga catatan Revisi-1/2/3.
2. Contoh entry yang sudah jadi di `SESI PENILAIAN/skrip/make_xlsx.py` (cari `S.append((`) —
   imitate gaya, panjang, dan tingkat detail bukti. Contoh bagus: entry "R21" RIZKY (evidence detail).
   Contoh gaya ringkas yang dipakai untuk batch 2 Oktober: entry "R46" DEBI LESTARI.

## ATURAN KETENTUAN (berlaku, menggantikan soal tertulis)
- **Komponen 1 (Penamaan Judul & Identitas, bobot 5)**: judul pendek "Personal Branding [Nama] [Kelas]"
  (page title) + judul panjang "ASTS Personal Branding [Nama] [Kelas] SMKN 9 Garut" (heading artikel).
- **Komponen 2 (Personal Branding & Logo, bobot 12)**: logo + nama branding + tagline + NAMA SISWA pada
  branding/logo + deskripsi konsep **minimal 100 kata** (hitung!). "by Nama Siswa" tidak lagi wajib
  (Revisi-1) — yang wajib NAMA SISWA tercantum pada branding.
- **Komponen 3 (Moodboard, bobot 8)**: format **landscape** + 7 unsur (warna utama, typography,
  style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual).
- **Komponen 4 (Mockup, bobot 10)**: minimal 3 mockup + penjelasan fungsi tiap media. Kalau jumlah
  unit di dalam satu gambar tidak terbaca OCR → status `TIDAK DAPAT DIVERIFIKASI`.
- **Komponen 5 (Naskah, 8)**: Judul, Tema, Pesan Utama, Narasi/Dialog, Closing Tagline.
- **Komponen 6 (Storyline, 8)**: Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup.
- **Komponen 7 (Shotlist, 10)**: minimal 10 shot. **Tabel ATAU gambar sama-sama diterima** (Revisi-1);
  kolom lengkap (No/Adegan/Jenis Shot/Angle/Movement/Durasi/Deskripsi) appreciated. Shotlist tertulis
  sebagai teks berurutan tanpa kolom → skor 3 (`SEBAGIAN`). Fewer than 10 → 1-2.
- **Komponen 8 (Storyboard, 10)**: minimal 6 scene + keterangan (shot/angle/transisi/dialog).
  Scene tertulis lengkap di teks = bukti sah. Kalau hanya gambar dan OCR tidak dapat membaca
  scene → `TIDAK DAPAT DIVERIFIKASI`, skor 2.
- **Komponen 9 (AI Mascot, 8)**: **Full Body WAJIB**. Portrait & Bersama Logo opsional.
  Bukti: label OCR pada gambar (mis. "Full Body / Portrait / With Logo") atau deskripsi teks.
- **Komponen 10 (Prompt & Dokumentasi AI, 7)**: prompt terdokumentasi (jumlah + label).
- **Komponen 11 (Portfolio Blogger, 10)**: isi artikel + **5 label wajib** (Tugas Sekolah,
  Personal Branding, AI, Portofolio, SMKN 9 Garut) + link aktif. Label kosong = penalti jelas (skor 2).
- **Komponen 12 (Kreativitas & Profesionalisme, 4)**: orisinalitas + konsistensi identitas.

Skala: **4** sangat baik · **3** baik · **2** cukup · **1** kurang · **0** tidak ada.
Status yang dipakai: `LENGKAP`, `SEBAGIAN`, `BELUM MEMENUHI`, `TIDAK DIKUMPULKAN`,
`TIDAK DAPAT DIVERIFIKASI`, `DIKUMPULKAN`.

## BATASAN TEKNIS (WAJIB)
- Kamu **TIDAK bisa melihat gambar**. Bukti visual hanya dari: OCR Apple Vision (sudah ada di dossier),
  dimensi piksel asli (`px=`), dan struktur HTML (heading, tabel, `<img>`, label).
- Dossier memuat satu blok "TEKS ARTIKEL LENGKAP" (fallback) — itu isi artikel apa adanya. Baca entirety.
- Nilai tambah resolusi **TIDAK** kamu hitung: cukup isi `RESOLUSI` (lihat di bawah) dari blok
  `[IMG#nn] WxH px=...` di dossier.
- **Jangan mengarang**. Setiap klaim harus bisa ditunjuk ke teks/OCR/dimensi di dossier.
- Timecode/Nama FILE: `dossier/R23.txt` untuk R23, dst.

## FORMAT OUTPUT
Tulis **satu file Python** per batch: `/var/folders/pl/qb4gscmn3479hd69g5txzl5w0000gp/T/opencode/entries_<N>.py`
Isinya HANYA dua bentuk statement (tanpa `import`, tanpa print), untuk tiap siswa yang kamu nilai:

```python
S.append(("R69","FITRIYANI","XI DKV 4","02 Okt 2026",
"https://fitriyani011.blogspot.com/2026/09/personal-branding-fitriyani-xi-dkv-4.html","2 (terverifikasi teks)","10 (terverifikasi teks)","6 (terverifikasi teks)","3 (terverifikasi OCR)","9","2345 (memenuhi)",[4,3,3,3,4,4,3,3,4,4,2,4],[
  ("LENGKAP","Bukti faktual komponen 1..."),
  ("SEBAGIAN","Bukti faktual komponen 2..."),
  ... 12 baris, urutan komponen 1..12 ...
]))
RESOLUSI["FITRIYANI"] = (6,0,320,"")
```

Keterangan:
- Elemen 1 = id tebakan (pakai id dari `parsed.json`, mis. `R69`).
- Elemen 2 = NAMA SISWA **persis** seperti di Excel/parsed.json (uppercase).
- Elemen 3 = kelas. Elemen 4 = tanggal kirim `02 Okt 2026` (boleh jam, mis. `02 Okt 2026 08:48`).
- Elemen 5 = URL lengkap.
- Elemen 6-11 = ringkasan kolom rekap: mockup, shotlist, storyboard, mascot, prompt, jumlah kata.
  Ikuti gaya yang sudah dipakai di make_xlsx.py, contoh: `"3 (terverifikasi)"`, `"10 (terverifikasi OCR)"`,
  `"1 (full body)"`, `"8"`, `"384 (memenuhi)"`, `"TIDAK DAPAT DIVERIFIKASI"`.
- Elemen 12 = 12 skor 0-4 sesuai komponen 1..12.
- Elemen 13 = 12 tuple `(status, bukti)` — **bukti 1-3 kalimat, spesifik** (sebut kata/label/OCR/dimensi).
- `RESOLUSI[nama] = (jumlah_gambar, jumlah_ge_1000px, lebar_maks, "")` — hitung dari blok `[IMG#nn] WxH`
  di dossier. Gambar yang gagal diunduh / tidak terukur → 0 dan sertakan catatan singkat di slot ke-4.

## VERIFIKASI WAJIB SEBELUM MENULIS FILE
1. `python3 -c "compile(open('<file>').read(),'x','exec')"` harus lulus.
2. Hitung ulang nilai komponen: `nilai = skor/4*bobot` dengan `BOBOT=[5,12,8,10,8,8,10,10,8,7,10,4]`;
   pastikan tidak ada yang negatif/lebih dari bobot.
3. Pastikan **tepat 12** tuple status-bukti dan **tepat 12** skor per siswa.