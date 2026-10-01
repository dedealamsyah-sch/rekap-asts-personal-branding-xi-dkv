# Panduan Update Data Rekap ASTS

File utama yang memuat data siswa adalah objek JavaScript `DATA` di dalam **index.html**.  Bagian penting yang berubah-ubah adalah array `rekap`.

## Cara memperbarui data

1. **Siapkan file Excel**
   - Pastikan file **`REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx`** berada di root proyek (sama lokasi `index.html`).
   - Baris header harus mengandung kolom **`Nama Siswa`** (sebagai penanda baris header).

2. **Jalankan skrip pembaruan**
   ```bash
   python docs/update_data.py
   ```
   Skrip ini akan:
   - Membaca seluruh baris (setelah header) dari Excel.
   - Membuat array JSON berisi objek dengan properti:
     `no, nama, kelas, kirim, url, mockup, shotlist, storyboard, mascot, prompt, kata, resolusi, rubrik, bonus, nilai, kategori`.
   - Mengganti hanya bagian `"rekap": [...]` di dalam `index.html` dengan data baru, mempertahankan `meta`, `komponen`, `per_komponen`, `detail`, dll.

3. **Commit perubahan**
   ```bash
   git add index.html
   git commit -m "Refresh data dari Excel"
   git push
   ```
   Setelah commit, situs akan menampilkan data terbaru secara otomatis.

## Catatan tambahan
- **Tidak mengubah** bagian lain dari `DATA`.  Jika Anda ingin menambah atau mengubah `per_komponen`, `detail`, atau `bonus`, lakukan secara manual di `index.html`.
- Skrip **`docs/update_data.py`** sudah disertakan di dalam repository.  Anda dapat membuka file tersebut untuk melihat detail implementasinya.
- Jika ada penambahan atau penghapusan kolom di Excel, sesuaikan kamus `cols` di dalam skrip.

## Troubleshooting
- **File tidak ditemukan** – pastikan jalur file Excel dan `index.html` sesuai dengan struktur repository.
- **Error JSON** – pastikan tidak ada karakter tak valid di data Excel (misal tanda kutip ganda). Skrip secara otomatis men‑escape string.
- **Tidak ada perubahan setelah push** – periksa apakah skrip berhasil menulis kembali ke `index.html`. Anda dapat membuka file dan mencari bagian `"rekap":` untuk memastikan array telah ter‑update.

Dengan mengikuti panduan ini, Anda dapat memperbarui data situs kapan saja hanya dengan satu perintah.
