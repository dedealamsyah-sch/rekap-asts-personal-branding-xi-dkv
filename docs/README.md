# Project Documentation

This repository contains the web application for the ASTS Personal Branding recap.

## How to keep the data up-to-date

The web page (`index.html`) reads all of its student data from the JavaScript
constant `DATA`. **Every** section of `DATA` is generated from the source Excel
file (`REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx`) by the
helper script `docs/update_data.py`.

```bash
python3 docs/update_data.py          # write the new data into index.html
python3 docs/update_data.py --check  # show what would change, write nothing
```

### Excel sheet → `DATA` section

| Excel sheet | `DATA` key(s) | Isi |
|---|---|---|
| `1. REKAP NILAI` | `rekap`, bagian dari `meta` | nilai akhir tiap siswa |
| `2. NILAI KOMPONEN` | `per_komponen`, `komponen` | skor 0–4 per komponen + bobot |
| `3. REKAP KETIDAKLENGKAPAN` | `ketidaklengkapan` | temuan ketidaklengkapan |
| `4. CEK MANUAL GURU` | `cek_manual` | butir yang perlu dicek guru |
| `5. DATA REKAPAN` | `data_rekapan` | rekapan pengumpulan Google Form |
| `6. METODE & RUBRIK` | `metode`, `meta.revisi` | metode, rubrik, dan revisi |
| `7. NILAI TAMBAHAN RESOLUSI` | `bonus`, `bonus_rules`, `meta.gambar` | nilai tambah resolusi |
| `Detail - <nama siswa>` | `detail` | bukti & hasil pemeriksaan per komponen |

### Notes

- `rekap` is written pretty-printed; every other section stays on a single
  compact line, preserving the existing file layout. Only the `DATA` object is
  touched — the rest of `index.html` is left byte-for-byte unchanged.
- `rekap[].bonus` and `per_komponen[].bonus` are written as **numbers**
  (`1.0`, `0.7`, `0.0`). The page compares them with `bonus > 0`.
- `detail` is keyed by the **full student name** from `1. REKAP NILAI`. Excel
  truncates sheet names to 31 characters (`Quinsya Rahmanesa Sole`,
  `Meylan Meliyanti Anast`, …), so the script maps each sheet back to the full
  name; otherwise the page's `norm()` lookup would fail.- `meta.gambar` is the sum of the "Jumlah Gambar" column in sheet 7, and
  `meta.rata` / `meta.rataRubrik` / `meta.rataBonus` are computed from `rekap`.

### After running the script

```bash
git add index.html
git commit -m "Refresh data dari Excel"
git push
```

## Files
- `index.html` – the interactive recap page (contains the `DATA` object).
- `docs/update_data.py` – regenerates `DATA` from the Excel workbook.
- `docs/update_instructions.md` – short usage guide.
- `rekap_nilai_asts_hasil.json` – machine-readable copy of the final scores.
