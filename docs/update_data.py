#!/usr/bin/env python3
import json, openpyxl, re, pathlib, sys

EXCEL_PATH = pathlib.Path(__file__).parents[1] / 'REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx'
HTML_PATH = pathlib.Path(__file__).parents[1] / 'index.html'

def load_excel():
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb.active
    header = None
    header_row = None
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if any(cell and isinstance(cell, str) and 'Nama Siswa' in cell for cell in row):
            header = list(row)
            header_row = i
            break
    if not header:
        sys.exit('Header not found in Excel')
    cols = {name: idx for idx, name in enumerate(header)}
    records = []
    for row in ws.iter_rows(min_row=header_row+1, values_only=True):
        if row[0] is None:
            continue
        rec = {
            'no': row[cols.get('No')],
            'nama': row[cols.get('Nama Siswa')],
            'kelas': row[cols.get('Kelas')],
            'kirim': row[cols.get('Waktu Kirim')],
            'url': row[cols.get('Link Blogger')],
            'mockup': row[cols.get('Mockup')],
            'shotlist': row[cols.get('Shotlist')],
            'storyboard': row[cols.get('Storyboard')],
            'mascot': row[cols.get('Mascot')],
            'prompt': str(row[cols.get('Prompt')]) if row[cols.get('Prompt')] is not None else None,
            'kata': row[cols.get('Jml Kata Deskripsi')],
            'resolusi': row[cols.get('Resolusi (px)')],
            'rubrik': row[cols.get('Nilai Rubrik /100')],
            'bonus': row[cols.get('Nilai Tambah Resolusi')],
            'nilai': row[cols.get('NILAI AKHIR /100')],
            'kategori': row[cols.get('Kategori')]
        }
        records.append(rec)
    return records

def update_html(records):
    text = HTML_PATH.read_text(encoding='utf-8')
    new_json = json.dumps(records, ensure_ascii=False, indent=2)
    new_text = re.sub(r'"rekap"\s*:\s*\[.*?\]', f'"rekap": {new_json}', text, flags=re.S)
    HTML_PATH.write_text(new_text, encoding='utf-8')
    print('index.html updated with', len(records), 'records')

if __name__ == '__main__':
    recs = load_excel()
    update_html(recs)
