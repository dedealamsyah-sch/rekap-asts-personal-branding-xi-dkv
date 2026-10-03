import json, re, sys, time, os
import requests
from bs4 import BeautifulSoup

# Selalu utf-8 untuk baca/tulis berkas: Windows default (cp1252) gagal pada teks Indonesia.
def _load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)

def _dump(obj, p):
    with open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)

UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"}
OUT = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(OUT, "raw")
os.makedirs(RAW, exist_ok=True)

SHEET_ID = "1eq2UyXOYX8z60ybOxGiSa9CLhDQRrRrhQViKrzZjIoU"
SHEET_XLSX = "https://docs.google.com/spreadsheets/d/%s/export?format=xlsx" % SHEET_ID

def baca_rekapan():
    """Ambil data rekapan langsung dari Google Spreadsheet (Form Responses 1).
    Mengembalikan [(id, timestamp, kelas, nama, url), ...]
    Kolom nama/link bisa berada di salah satu dari 4 pasang kolom,
    jadi semua dicek agar siswa yang mengisi kolom tidak pertama tetap terdeteksi."""
    import io
    import openpyxl
    r = requests.get(SHEET_XLSX, headers={"User-Agent": "Mozilla/5.0"}, timeout=90)
    r.raise_for_status()
    wb = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    hdr = [str(h or "").strip().lower() for h in rows[0]]
    ci_kelas = next((i for i, h in enumerate(hdr) if "kelas" in h), 1)
    ci_ts    = 0
    pasangan = []          # [(idx_nama, idx_link), ...]
    for i, h in enumerate(hdr):
        if "nama" in h:
            j = next((k for k in range(i + 1, len(hdr))
                      if "link" in hdr[k] or "blogger" in hdr[k]), None)
            if j is not None:
                pasangan.append((i, j))
    hasil, dipakai = [], set()
    for n, row in enumerate(rows[1:], 1):
        ts = row[ci_ts] if ci_ts < len(row) else ""
        if hasattr(ts, "strftime"):
            ts = ts.strftime("%Y-%m-%d %H:%M:%S")
        for ci_n, ci_l in pasangan:
            nama = str(row[ci_n]).strip() if ci_n < len(row) and row[ci_n] else ""
            link = str(row[ci_l]).strip() if ci_l < len(row) and row[ci_l] else ""
            if not nama or not link or link.lower() == "nan":
                continue
            kunci = nama.upper()
            if kunci in dipakai:      # cegah duplikat bila ada baris ganda
                continue
            dipakai.add(kunci)
            kelas = str(row[ci_kelas]).strip() if ci_kelas < len(row) and row[ci_kelas] else ""
            hasil.append(("R%02d" % (len(hasil) + 1), ts, kelas, nama, link))
    return hasil

ROWS = baca_rekapan()
print("Rekapan dari Google Spreadsheet: %d baris kiriman" % len(ROWS))
for r in ROWS:
    print("   %-4s %-12s %-30s %s" % (r[0], r[2], r[3][:30], r[4][:70]))
print("-" * 100)


KW = {
  "logo": r"\blogo\b", "moodboard": r"mood\s*board", "mockup": r"mock\s*-?\s*up",
  "naskah": r"naskah|skrip|script|dialog|iklan", "storyline": r"story\s*line",
  "shotlist": r"shot\s*list", "storyboard": r"story\s*board",
  "mascot": r"mascot|maskot", "prompt": r"prompt", "tagline": r"tagline|slogan",
  "deskripsi": r"deskripsi|konsep|filosofi|pengertian", "by_nama": r"\bby\b",
  "color": r"warna|color|palet", "tipografi": r"tipografi|typography|font",
}

def strip_tags(s):
    return re.sub(r"\s+", " ", s).strip()


def _arsipkan(fn, arsip, teks):
    """Simpan hasil fetch yang GAGAL ke <fn>.arsip, jangan menimpa bukti lama.
    Kalau <fn> sudah berisi artikel valid (punya <img>), jangan sentuh sama sekali."""
    try:
        if os.path.exists(fn):
            lama = open(fn, encoding="utf-8", errors="ignore").read()
            if lama.count("<img") > 0 or lama.count("post-body") == 0:
                return  # bukti lama masih utuh
        with open(arsip, "w", encoding="utf-8") as f:
            f.write(teks)
    except Exception:
        pass

def parse(rid, ts, kelas, nama, url):
    rec = {"id": rid, "timestamp": ts, "kelas": kelas, "nama": nama, "url": url}
    fn = os.path.join(RAW, rid + ".html")
    # Arsip bukti: JANGAN pernah menimpa raw/<id>.html dengan halaman gagal
    # (404 / login). Page HTML yang sebelumnya berhasil adalah satu-satunya bukti
    # karya bila link suatu saat mati — menimpanya menghilangkan bukti secara permanen.
    arsip = fn + ".arsip"
    try:
        r = requests.get(url, headers=UA, timeout=45, allow_redirects=True)
        rec["status_code"] = r.status_code
        rec["final_url"] = r.url
        rec["bytes"] = len(r.content)
        if "blogger.com" in r.url and "/post/edit" in r.url:
            rec["accessible"] = False
            rec["reason"] = "URL editor Blogger (butuh login) - tidak dapat diakses publik"
            _arsipkan(fn, arsip, r.text)
            return rec
        if r.status_code != 200:
            rec["accessible"] = False
            rec["reason"] = "HTTP %s" % r.status_code
            _arsipkan(fn, arsip, r.text)
            return rec
        open(fn, "w", encoding="utf-8").write(r.text)
    except Exception as e:
        rec["accessible"] = False
        rec["reason"] = "Error: %s" % e
        return rec

    soup = BeautifulSoup(r.text, "html.parser")
    t = soup.find("title")
    rec["page_title"] = strip_tags(t.get_text()) if t else ""
    h1 = soup.find("h1", class_=re.compile("title"))
    if h1: rec["h1"] = strip_tags(h1.get_text())

    post = None
    for sel in ["div.post-body", "div.post-body.entry-content", "div[itemprop='articleBody']",
                "div#post-body-", "div.post-body.element magazine"]:
        post = soup.select_one(sel)
        if post and len(post.get_text(strip=True)) > 80:
            break
    if post is None:
        cand = soup.find_all("div")
        best, bl = None, 0
        for d in cand:
            l = len(d.get_text(strip=True))
            if l > bl and l < 60000 and len(d.find_all(["div"])) < 30:
                best, bl = d, l
        post = best
    if post is None:
        rec["accessible"] = False
        rec["reason"] = "Badan artikel tidak ditemukan"
        return rec
    rec["accessible"] = True

    # clone to strip non-content widgets inside post
    body = BeautifulSoup(str(post), "html.parser")

    # headings
    rec["headings"] = [strip_tags(h.get_text()) for h in body.find_all(["h1","h2","h3","h4","h5"])]

    # images
    imgs = []
    for im in body.find_all("img"):
        src = im.get("src") or im.get("data-src") or ""
        if "profile_img" in src or "bp.blogspot" in src and "blogger" in src:
            pass
        w = im.get("width") or ""
        hgt = im.get("height") or ""
        imgs.append({"src": src[:180], "alt": strip_tags(im.get("alt") or ""), "w": w, "h": hgt})
    rec["images"] = imgs
    rec["img_count"] = len(imgs)

    # tables
    tables = []
    for tb in body.find_all("table"):
        rows = []
        for tr in tb.find_all("tr"):
            cells = [strip_tags(td.get_text()) for td in tr.find_all(["td","th"])]
            if any(c for c in cells):
                rows.append(cells)
        tables.append({"rows": rows, "nrows": len(rows), "maxcols": max([len(r) for r in rows], default=0)})
    rec["tables"] = tables

    # text
    for bad in body.find_all(["script","style"]):
        bad.decompose()
    text = strip_tags(body.get_text(" "))
    rec["text"] = text
    rec["word_count"] = len(re.findall(r"[A-Za-z0-9'\-]+", text))
    rec["kw"] = {k: len(re.findall(v, text, re.I)) for k, v in KW.items()}
    rec["kw_in_headings"] = {k: len(re.findall(v, " ".join(rec["headings"]), re.I)) for k, v in KW.items()}

    # labels / tags
    labs = []
    for a in soup.find_all("a", href=True):
        if "/search/label/" in a["href"]:
            labs.append(strip_tags(a.get_text()))
    rec["labels"] = sorted(set(labs))

    # section-level word counts: split text by headings
    sections = {}
    cur = "AWAL"
    buf = []
    for el in body.find_all(["h1","h2","h3","h4","p","li","td","div"], recursive=True):
        if el.name in ("h1","h2","h3","h4") and el.get_text(strip=True):
            sections[cur] = re.findall(r"[A-Za-z0-9'\-]+", " ".join(buf))
            cur = strip_tags(el.get_text())[:80]
            buf = []
        elif el.name in ("p","li"):
            buf.append(el.get_text(" "))
    sections[cur] = re.findall(r"[A-Za-z0-9'\-]+", " ".join(buf))
    rec["sections"] = {k: len(v) for k, v in sections.items() if v}
    rec["sections_max"] = max(rec["sections"].values(), default=0)
    rec["sections_keys"] = list(rec["sections"].keys())
    return rec

data = []
for row in ROWS:
    rec = parse(*row)
    data.append(rec)
    print(rec["id"], rec["nama"], "| acc:", rec.get("accessible"), "| words:", rec.get("word_count"),
          "| img:", rec.get("img_count"), "| tables:", len(rec.get("tables", [])), "| labels:", len(rec.get("labels", [])), flush=True)

# ---------- deteksi perubahan terhadap snapshot sebelumnya ----------
# Snapshot pembanding disimpan di ../data/parsed.json (bukan di folder skrip)
SNAP = os.path.join(os.path.dirname(OUT), "data", "parsed.json")
lama = {}
if os.path.exists(SNAP):
    lama = {r["nama"].upper(): r for r in _load(SNAP)}
    print("Snapshot pembanding: %s (%d siswa)" % (SNAP, len(lama)))
    # PENJAGA: bila snapshot jauh lebih tua dari skrip/parsed.json, maka selisihnya
    # akan dilaporkan sebagai "KIRIMAN BARU" padahal siswa itu sudah dinilai.
    # Sinkronkan dulu:  cp parsed.json ../data/parsed.json
    _sini = os.path.join(OUT, "parsed.json")
    if os.path.exists(_sini):
        _n_sini = len(json.load(open(_sini, encoding="utf-8")))
        if len(lama) < _n_sini - 3:
            print("  !! PERINGATAN: snapshot (%d) lebih tua dari parsed.json (%d)."
                  % (len(lama), _n_sini))
            print("  !! %d siswa akan Terlihat sebagai 'KIRIMAN BARU' padahal sudah dinilai."
                  % (_n_sini - len(lama)))
            print("  !! Jalankan dulu:  cp parsed.json ../data/parsed.json")
else:
    print("PERINGATAN: snapshot %s tidak ditemukan - perbandingan tidak dapat dilakukan" % SNAP)
baru_nama = [r["nama"].upper() for r in data]
tambahan = [n for n in baru_nama if n not in lama]
hilang   = [n for n in lama if n not in baru_nama]
ubah = []
for r in data:
    o = lama.get(r["nama"].upper())
    if not o:
        continue
    d = []
    if o.get("word_count") != r.get("word_count"): d.append("kata %s->%s" % (o.get("word_count"), r.get("word_count")))
    if o.get("img_count")  != r.get("img_count"):  d.append("gambar %s->%s" % (o.get("img_count"), r.get("img_count")))
    if len(o.get("tables", [])) != len(r.get("tables", [])): d.append("tabel %s->%s" % (len(o.get("tables", [])), len(r.get("tables", []))))
    if o.get("labels") != r.get("labels"):        d.append("label berubah")
    if o.get("accessible") != r.get("accessible"): d.append("AKSES LINK BERUBAH")
    if d: ubah.append((r["nama"], d))

_dump(data, os.path.join(OUT, "parsed.json"))
print("=" * 100)
print("RINGKASAN PERUBAHAN")
if tambahan:  print("KIRIMAN BARU (%d): %s" % (len(tambahan), ", ".join(tambahan)))
if hilang:    print("HILANG DARI REKAPAN (%d): %s" % (len(hilang), ", ".join(hilang)))
if ubah: 
    for n, d in ubah: print("BERUBAH - %s: %s" % (n, "; ".join(d)))
if not (tambahan or hilang or ubah): print("TIDAK ADA PERUBAHAN - tidak ada siswa yang mengunggah ulang")
print("TOTAL PERUBAHAN:", len(tambahan) + len(hilang) + len(ubah))
print("DONE")
