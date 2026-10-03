# BRIEF PENILAIAN BATCH — ASTS "Personal Branding Creative Campaign"

Kamu menilai karya ASTS untuk siswa yang **baru mengirim / memperbaiki kiriman**.
Skor resmi butuh **bukti faktual**, bukan perkiraan.

## 0. YANG SUDAH DISIAPKAN — baca dossier, itu satu-satunya sumbermu

Semua bukti sudah di dalam satu berkas per siswa:

`SESI PENILAIAN/skrip/tools/out/dossier/<RID>.txt`

Isinya:
- **Teks artikel per bagian** (hasil ekstraksi heading→paragraf).
- **Blok `TEKS ARTIKEL LENGKAP (get_text() penuh, N kata)`** di akhir — WAJIB dibaca
  entirety bila ada. Bagian atas bisa *under-report* karena banyak siswa menulis teks
  di dalam `<div>`/`<span>`, bukan `<p>`.
- **Tabel HTML** yang dibaca baris per baris (jumlah baris = jumlah shot bila tabel shotlist).
- **Daftar gambar**: `[IMG#nn] px=WxH orientasi | bagian: ... | file: ...` +
  baris `OCR:` berisi teks yang terbaca dari gambar itu.

> **PENTING — soal gambar:** kamu **TIDAK** bisa melihat gambar, dan tidak perlu.
> Bukti visual sudah diterjemahkan jadiOCR di dossier. Jangan pernah mengklaim
> "saya melihat gambar" — rujuklah isi baris `OCR:`.
>
> Keterbatasan OCR yang harus kamudapat:
> - Teks di dalam gambar buatan AI sering **pseudoteks** (`"DREAM-DESXN-CXIATS"`,
>   `"B7SH0LST"`, `"Gouwn Irng"`). Angka/baris di dalam gambar **tidak bisa dihitung
>   andal**. Bila jumlah baris hanya kelihatan dari OCR yang kacau → tulis
>   `TIDAK DAPAT DIVERIFIKASI`, jangan menebak.
> - `OCR: (tidak ada teks yang terbaca)` = gambar benar-benar tanpa teks yang bisa
>   dibaca. Itu bukti tidak adanya caption/label, bukan bukti ketiadaan isi.
> - `OCR: (tidak ada di cache ...)` = OCR belum dijalankan untuk id itu. **Jangan
>   menilai komponen visual dari id itu**; tandai `TIDAK DAPAT DIVERIFIKASI`.

## 1. RUBRIK (WAJIB — 12 komponen, skor 0-4, total bobot 100)

| No | Komponen | Bobot |
|---:|---|---:|
| 1 | Penamaan Judul & Identitas | 5 |
| 2 | Personal Branding & Logo | 12 |
| 3 | Moodboard | 8 |
| 4 | Mockup Branding | 10 |
| 5 | Naskah Iklan | 8 |
| 6 | Storyline | 8 |
| 7 | Shotlist | 10 |
| 8 | Storyboard | 10 |
| 9 | AI Mascot Character | 8 |
| 10 | Prompt & Dokumentasi AI | 7 |
| 11 | Portfolio Blogger | 10 |
| 12 | Kreativitas & Profesionalisme | 4 |

### Ketentuan per komponen

- **K1 (b5)** Judul pendek `Personal Branding [Nama] [Kelas]` (page title) + judul panjang
  `ASTS Personal Branding [Nama] [Kelas] SMKN 9 Garut` (heading artikel). Kedua-duanya ada → 4.
- **K2 (b12)** Logo + nama branding + tagline + **NAMA SISWA tercantum pada branding/logo**
  + deskripsi konsep **minimal 100 kata** (hitung!). `by Nama Siswa` tidak wajib lagi
  (Revisi-1) — yang wajib NAMA SISWA ada. Skor 4 = semua ada & deskripsi ≥100 kata.
- **K3 (b8)** Moodboard format **landscape** + **7 unsur**: warna utama, typography,
  style visual, referensi desain, tone & mood, elemen grafis,cjika Inspirasi visual.
  Persegi/potret → kurangi. 7 unsur + landscape → 4.
- **K4 (b10)** Minimal **3 mockup** + penjelasan fungsi tiap media.
- **K5 (b8)** Naskah: Judul, Tema, Pesan Utama, Narasi/Dialog, Closing Tagline.
- **K6 (b8)** Storyline: Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup.
- **K7 (b10)** Shotlist minimal **10 shot**. Tabel **atau** gambar sama-sama diterima
  (Revisi-1). Kolom lengkap `No/Adegan/Jenis Shot/Angle/Movement/Durasi/Deskripsi` =
  nilai tinggi. Shotlist tertulis berurutan tanpa kolom → skor 3 (`SEBAGIAN`).
  Kurang dari 10 → skor 1-2 dan tulis `BELUM MEMENUHI`.
- **K8 (b10)** Minimal **6 scene** + keterangan (shot/angle/transisi/dialog).
  Scene tertulis lengkap di artikel = bukti sah. Kalau hanya gambar dan tidak terbaca → skor 2.
- **K9 (b8)** **Full Body WAJIB.** Portrait & Bersama Logo opsional.
  4 = ketiganya ada; 3 = Full Body ada (wajib terpenuhi); 2 = mascot ada tapi Full Body
  tidak dapat dipastikan; 1 = sangat terbatas; 0 = tidak ada gambar mascot.
- **K10 (b7)** Prompt terdokumentasi (hitung jumlahnya + apakah diberi label).
- **K11 (b10)** Isi artikel + **5 label wajib**: `Tugas Sekolah`, `Personal Branding`, `AI`,
  `Portofolio`, `SMKN 9 Garut`. Link aktif. Label kosong/tidak relevan → penalti jelas (skor 2).
- **K12 (b4)** Orisinalitas + konsistensi identitas visual.

Skala: **4** sangat baik · **3** baik · **2** cukup · **1** kurang · **0** tidak ada.

### Revisi-3 (nilai tambah resolusi)
Nilai Akhir = Nilai Rubrik + nilai tambah resolusi. TAdds **TIDAK** kamu hitung —
cukup isi `RESOLUSI`. Resolusi rendah tidak pernah menjadi penalti.

## 2. ATURAN STATUS (pilih salah satu, hanya itu)
`LENGKAP` · `DIKUMPULKAN` · `SEBAGIAN` · `BELUM MEMENUHI` · `TIDAK DIKUMPULKAN` ·
`TIDAK DAPAT DIVERIFIKASI`

- **JANGAN mengarang.** Setiap klaim harus bisa ditunjuk ke teks artikel / isi gambar /
  dimensi px / label.
- **JANGANG** anggap "DIKUMPULKAN" = "MEMENUHI". 8 shot → tulis `BELUM MEMENUHI — 8 shot`.
- **Jangan menilai hanya dari estetika.** Estetika bukan pengganti ketentuan.
- Kalau nama pada artikel/rekapan berbeda → catat di bukti K1 dan K12.
- Bahasa Indonesia, netral, objektif. Untuk indikasi tidak orisinal tulis
  "terindikasi tidak orisinal / perlu klarifikasi guru", jangan menuduh.

## 3. FORMAT OUTPUT — tulis SATU file `.py`

Hanya dua bentuk statement, **tanpa `import`, tanpa `print`**, satu blok per siswa:

```python
S.append(("R112","SITI MULYANI","XI DKV 1","03 Okt 2026 08:04",
"https://sitimulyani0127.blogspot.com/...","3 (terverifikasi)","10 (terverifikasi gambar)","6 (terverifikasi teks)","3 (terverifikasi label)","2","1682 (memenuhi)",[4,4,4,3,4,4,3,4,3,4,4,4],[
  ("LENGKAP","Page title 'Personal Branding Siti Mulyani XI DKV 1' + heading 'ASTS Personal Branding Siti Mulyani XI DKV 1 SMKN 9 Garut' - keduanya sesuai."),
  ("LENGKAP","Logo monogram SM + bola futsal (IMG#01); nama branding 'Siti Mulyani'; tagline 'Menciptakan Jejak Visual yang Dinamis'; deskripsi 172 kata; nama siswa tercetak pada logo."),
  ("LENGKAP","IMG#02 1024x572 landscape. 7 unsur terverifikasi di teks: Warna Utama (hex #001F4D dsb), Typography, Style Visual, Referensi Desain, Tone & Mood, Elemen Grafis, Inspirasi Visual."),
  ... 12 baris, urutan komponen 1..12 ...
]))
RESOLUSI["SITI MULYANI"] = (10,10,1024,"")
```

Keterangan elemen:
1. id dari `parsed.json` (mis. `R112`)
2. NAMA SISWA **persis** seperti di `parsed.json` (uppercase)
3. kelas · 4. tanggal kirim `DD MMM YYYY` (+ jam bila ada) · 5. URL lengkap
6-11. ringkasan kolom rekap: mockup, shotlist, storyboard, mascot, prompt, jumlah kata.
   Gaya yang sudah dipakai di repo: `"3 (terverifikasi)"`, `"10 (terverifikasi gambar)"`,
   `"6 (terverifikasi teks)"`, `"1 (full body)"`, `"8"`, `"1682 (memenuhi)"`,
   `"TIDAK DAPAT DIVERIFIKASI"`.
12. **tepat 12** skor 0-4 sesuai komponen 1..12
13. **tepat 12** tuple `(status, bukti)` — bukti **1-3 kalimat spesifik**, sebutkan
    kata/label/nama berkas/dimensi px yang dibaca.

`RESOLUSI[nama] = (jumlah_gambar, jumlah_ge_1000px, lebar_maks, catatan)`
Hitung dari blok `[IMG#nn] px=WxH` di dossier. Gambar gagal/unduh → 0 + catatan singkat
di slot ke-4.

> PENTING: `RESOLUSI` dihitung dari **dossier siswa itu saja**, bukan meniru nilai lain.

## 4. VERIFIKASI WAJIB SEBELUM MENULIS FILE

Jalankan ini dan pastikan lulus:
```bash
python -c "import ast,io,sys; ast.parse(open(r'<file>',encoding='utf-8').read()); print('OK')"
```
Lalu periksa sendiri:
- setiap siswa **tepat 12** skor dan **tepat 12** tuple status-bukti, urutan 1..12;
- skor semua dalam 0..4;
- nilai komponen = skor/4 × bobot, tidak negatif dan tidak melebihi bobot
  (`BOBOT=[5,12,8,10,8,8,10,10,8,7,10,4]`, jumlah = 100).

## 5. LAPORKAN KEMBALI (wajib, singkat)

1. Nama file `.py` yang ditulis + path.
2. Daftar siswa yang dinilai, dengan nilai rubrik & nilai akhir (hitung sendiri).
3. **Semua temuan tidak biasa** yang butuh perhatian guru, terutama:
   - nama pada rekapan ≠ nama pada artikel/branding;
   - link yang bukan projek personal branding (mis. portofolio tugas lain);
   - indikasi tidak orisinal / teks identik antar siswa;
   - komponen yang tidak dapat diverifikasi despite ada klaim;
   - link tidak dapat diakses (URL editor Blogger / HTTP 404).
4. Rata-rata nilai rubrik batch ini.

Jangan mengarang. Kalau tidak yakin, tulis `TIDAK DAPAT DIVERIFIKASI` dan laporkan.