# Panduan Update Data Rekap ASTS

Seluruh data pada halaman **`index.html`** dibaca dari objek JavaScript `DATA`.
Objek tersebut di-*generate ulang* dari file Excel sumber:

> **`REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx`**

## Cara memperbarui data

1. **Siapkan file Excel**
   - Letakkan workbook di root proyek (satu folder dengan `index.html`).
   - Jangan mengubah nama sheet. Skrip mengenali sheet berikut:
     `1. REKAP NILAI`, `2. NILAI KOMPONEN`, `3. REKAP KETIDAKLENGKAPAN`,
     `4. CEK MANUAL GURU`, `5. DATA REKAPAN`, `6. METODE & RUBRIK`,
     `7. NILAI TAMBAHAN RESOLUSI`, dan semua sheet `Detail - <nama>`.

2. **Jalankan skrip pembaruan**
   ```bash
   python3 docs/update_data.py
   ```
   Skrip akan menimpa **seluruh isi `DATA`** (bukan hanya `rekap`) dengan data
   terbaru dari Excel, yaitu:

   | Sheet Excel | Bagian `DATA` |
   |---|---|
   | `1. REKAP NILAI` | `rekap`, `meta` (rata-rata) |
   | `2. NILAI KOMPONEN` | `per_komponen`, `komponen` |
   | `3. REKAP KETIDAKLENGKAPAN` | `ketidaklengkapan` |
   | `4. CEK MANUAL GURU` | `cek_manual` |
   | `5. DATA REKAPAN` | `data_rekapan` |
   | `6. METODE & RUBRIK` | `metode`, `meta.revisi` |
   | `7. NILAI TAMBAHAN RESOLUSI` | `bonus`, `bonus_rules`, `meta.gambar` |
   | `Detail - <nama>` | `detail` |

   Untuk melihat perubahannya tanpa menulis file:
   ```bash
   python3 docs/update_data.py --check
   ```

3. **Commit perubahan**
   ```bash
   git add index.html
   git commit -m "Refresh data dari Excel"
   git push
   ```

## Catatan teknis

- Hanya objek `DATA` yang diubah; sisa `index.html` tidak tersentuh.
- `rekap` tetap ditulis multi-baris, bagian lain tetap satu baris (compact)
  agar format file tidak berubah banyak.
- Nilai tambah resolusi ditulis sebagai **angka** (`1.0`, `0.7`, `0.0`),
  bukan string `"+1,0"`, karena JavaScript membandingkannya dengan `> 0`.
- `detail` di-key dengan **nama lengkap** dari sheet `1. REKAP NILAI`. Nama
  sheet Excel dipotong 31 karakter, sehingga skrip memetakan ulang tiap sheet
  ke nama lengkap siswa agar pencarian `norm()` di halaman tetap cocok.
- `meta.gambar` = jumlah kolom "Jumlah Gambar" pada sheet 7.
- Skrip bersifat *idempotent*: dijalankan dua kali menghasilkan file sama.

## Troubleshooting

- **File tidak ditemukan** – pastikan workbook dan `index.html` berada di root
  proyek.
- **Header tidak ditemukan** – periksa apakah baris header sheet masih memuat
  label aslinya (mis. `Nama`, `Komponen`, `Jumlah Gambar`).
- **Tidak ada perubahan setelah push** – pastikan skrip berhasil menulis
  `index.html` (lihat pesan ringkasan yang dicetak).

Dengan mengikuti panduan ini, cukup satu perintah untuk menyelaraskan halaman
dengan data Excel terbaru.
