# Project Documentation

This repository contains the web application for the ASTS Personal Branding recap.

## How to keep the data up‑to‑date

The web page (`index.html`) reads the student data from the JavaScript constant `DATA`. To update the data after changes in the source Excel file (`REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx`), run the helper script located at `docs/update_data.py`.

```bash
python docs/update_data.py
```

The script will:
1. Load the Excel file.
2. Extract the rows from the sheet (starting after the header row that contains "Nama Siswa").
3. Build a JSON array of the `rekap` entries.
4. Replace only the `rekap` part of the `DATA` object in `index.html`, preserving the existing `meta`, `komponen`, `per_komponen`, `detail`, `bonus`, etc.
5. Write the updated `index.html` back to the repository.

After running the script, commit the changes:

```bash
git add index.html
git commit -m "Update rekap data from Excel"
git push
```

## Files
- `docs/update_data.py` – the data‑update script.
- `docs/README.md` – this documentation.

Feel free to add more documentation or automation scripts as needed.
