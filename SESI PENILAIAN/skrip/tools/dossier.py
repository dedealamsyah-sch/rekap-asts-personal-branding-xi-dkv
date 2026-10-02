"""Dossier v2: teks artikel per heading (penuh), tabel, gambar+px, OCR.

Pakai: python3 dossier2.py R69 R71 ...
"""
import json
import os
import re
import subprocess
import sys

from bs4 import BeautifulSoup

SK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
data = {r['id']: r for r in json.load(open(os.path.join(SK, 'parsed.json')))}
images = json.load(open(os.path.join(SK, 'images.json')))
imgby = {}
for m in images:
    imgby.setdefault(m['id'], []).append(m)


def strip(s):
    return re.sub(r'\s+', ' ', s).strip()


def wc(s):
    return len(re.findall(r"[A-Za-z0-9'\-]+", s))


import json as _json
import textwrap
CACHE = os.path.join(TMP, 'ocr_cache.json')
cache = _json.load(open(CACHE)) if os.path.exists(CACHE) else {}

out = []
for rid in sys.argv[1:]:
    rec = data[rid]
    out.append('=' * 100)
    out.append('ID %s | %s | %s | ts %s' % (rid, rec['nama'], rec['kelas'], rec['timestamp']))
    out.append('URL: %s' % rec['url'])
    out.append('STATUS %s | words=%s | img=%s | tables=%s | labels=%s'
               % (rec.get('status_code'), rec.get('word_count'), rec.get('img_count'),
                  len(rec.get('tables', [])), rec.get('labels')))
    if not rec.get('accessible'):
        out.append('!! TIDAK DAPAT DIakses: %s' % rec.get('reason'))
        continue
    soup = BeautifulSoup(open(os.path.join(SK, 'raw', rid + '.html'), encoding='utf-8', errors='ignore').read(), 'html.parser')
    post = None
    for sel in ["div.post-body", "div.post-body.entry-content", "div[itemprop='articleBody']",
                "div#post-body-", "div.post-body.element magazine"]:
        post = soup.select_one(sel)
        if post and len(post.get_text(strip=True)) > 80:
            break
    if post is None:
        best, bl = None, 0
        for dd in soup.find_all('div'):
            l = len(dd.get_text(strip=True))
            if l > bl and l < 60000 and len(dd.find_all(['div'])) < 30:
                best, bl = dd, l
        post = best
    body = BeautifulSoup(str(post), 'html.parser')
    for b in body.find_all(['script', 'style']):
        b.decompose()

    # walk: heading -> teks sampai heading berikutnya
    cur = 'AWAL'
    buf = []
    nimg = 0
    for el in body.find_all(['h1', 'h2', 'h3', 'h4', 'p', 'li', 'table', 'img'], recursive=True):
        if el.name in ('h1', 'h2', 'h3', 'h4'):
            t = strip(el.get_text())
            if t:
                txt = strip(' '.join(buf))
                if wc(txt) > 2:
                    out.append('\n### [%s] (%d kata)' % (cur, wc(txt)))
                    out.append(txt[:1400])
                cur, buf = t, []
        elif el.name == 'img':
            src = el.get('src') or ''
            if 'profile_img' in src:
                continue
            nimg += 1
        elif el.name == 'table':
            rows = [[strip(td.get_text()) for td in tr.find_all(['td', 'th'])] for tr in el.find_all('tr')]
            rows = [r for r in rows if any(r)]
            if rows:
                out.append('\n### TABEL (%d baris x %d kolom)' % (len(rows), max(len(r) for r in rows)))
                for i, r in enumerate(rows[:30], 1):
                    out.append('  r%02d: %s' % (i, ' || '.join(r)[:400]))
        else:
            if el.find(['img', 'table']):
                continue
            t = strip(el.get_text())
            if wc(t) > 1:
                buf.append(t)
    txt = strip(' '.join(buf))
    if wc(txt) > 2:
        out.append('\n### [%s] (%d kata)' % (cur, wc(txt)))
        out.append(txt[:1400])

    if sum(wc(x) for x in out if x.startswith('###')) < 50:
        out.append('\n### TEKS ARTIKEL LENGKAP (fallback, %d kata)' % rec.get('word_count', 0))
        out.append(strip(rec.get('text', ''))[:26000])
        out.append('\n### JUMLAH KATA PER BAGIAN: %s' % rec.get('sections'))

    out.append('\n--- GAMBAR (%d) ---' % nimg)
    d = os.path.join(SK, 'img', rid)
    for m in imgby.get(rid, []):
        f = os.path.join(SK, m['file'])
        out.append('\n[IMG#%02d] %s px=%s %s' % (m['idx'], m.get('px'), m.get('orient'),
                                                (m.get('section') or '')[:60]))
        key = f
        if key in cache:
            ocr = cache[key]
        else:
            try:
                r = subprocess.run([os.path.join(SK, 'ocr'), f], capture_output=True, text=True,
                                   timeout=25)
                ocr = ' / '.join(x.strip() for x in (r.stdout or r.stderr).split('\n')[1:] if x.strip())
            except subprocess.TimeoutExpired:
                ocr = '(OCR timeout 25s, gambar dilewati)'
            cache[key] = ocr
            try:
                os.makedirs(os.path.dirname(CACHE), exist_ok=True)
                _json.dump(cache, open(CACHE, 'w'), ensure_ascii=False)
            except Exception:
                pass
        out.append('  OCR: ' + ocr[:900])

dd = os.path.join(TMP, 'dossier')
os.makedirs(dd, exist_ok=True)
for rid in sys.argv[1:]:
    body = '\n'.join(out)
    i = body.index('=' * 100 + '\nID %s' % rid)
    j = min([body.find('=' * 100 + '\nID %s ' % r) for r in sys.argv[1:] if r != rid
             and body.find('=' * 100 + '\nID %s ' % r) > i] or [len(body)])
    txt = textwrap.fill(body[i:j], 150)
    open(os.path.join(dd, rid + '.txt'), 'w', encoding='utf-8').write(txt)
    print(rid, len(txt), 'karakter')