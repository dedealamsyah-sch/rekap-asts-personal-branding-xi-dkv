"""Dossier|Windows): teks artikel per heading + tabel + daftar gambar (px/orientasi).
Tidak memakai OCR — verifikasi visual dilakukan langsung oleh AI dengan membaca gambar.

Pakai: python dossier_win.py R20 R25 ...
"""
import json, os, re, sys, textwrap
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
SK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS = os.path.join(SK, "tools")
OUTD = os.path.join(TOOLS, "out", "dossier")


def _load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


data = {r["id"]: r for r in _load(os.path.join(SK, "parsed.json"))}
imgby = {}
for m in _load(os.path.join(SK, "images.json")):
    imgby.setdefault(m["id"], []).append(m)

# Cache OCR Windows (RapidOCR). Absen untuk siswa lama -> Maybe blok kosong.
OCRP = os.path.join(TOOLS, "out", "ocr_cache_win.json")
OCRC = _load(OCRP) if os.path.exists(OCRP) else {}

strip = lambda s: re.sub(r"\s+", " ", s).strip()
wc = lambda s: len(re.findall(r"[A-Za-z0-9'\-]+", s))


def build(rid):
    rec = data[rid]
    o = []
    o.append("=" * 100)
    o.append("ID %s | %s | %s | ts %s" % (rid, rec["nama"], rec["kelas"], rec["timestamp"]))
    o.append("URL: %s" % rec["url"])
    o.append("PAGE TITLE: %r" % (rec.get("page_title") or ""))
    o.append("STATUS %s | words=%s | img=%s | tabel=%d | labels=%s"
             % (rec.get("status_code"), rec.get("word_count"), rec.get("img_count"),
                len(rec.get("tables") or []), rec.get("labels")))
    if not rec.get("accessible"):
        o.append("!! TIDAK DAPAT DIAKSES: %s" % rec.get("reason"))
        return "\n".join(o)

    soup = BeautifulSoup(open(os.path.join(SK, "raw", rid + ".html"),
                              encoding="utf-8", errors="ignore").read(), "html.parser")
    post = None
    for sel in ["div.post-body", "div.post-body.entry-content", "div[itemprop='articleBody']",
                "div#post-body-", "div.post-body.element magazine"]:
        post = soup.select_one(sel)
        if post and len(post.get_text(strip=True)) > 80:
            break
    if post is None:
        best, bl = None, 0
        for dd in soup.find_all("div"):
            L = len(dd.get_text(strip=True))
            if L > bl and L < 60000 and len(dd.find_all(["div"])) < 30:
                best, bl = dd, L
        post = best
    body = BeautifulSoup(str(post), "html.parser")
    for b in body.find_all(["script", "style"]):
        b.decompose()

    cur, buf, nimg = "AWAL", [], 0
    n_sec = 0
    for el in body.find_all(["h1", "h2", "h3", "h4", "p", "li", "table", "img"], recursive=True):
        if el.name in ("h1", "h2", "h3", "h4"):
            t = strip(el.get_text())
            if t:
                txt = strip(" ".join(buf))
                if wc(txt) > 2:
                    o.append("\n### [%s] (%d kata)" % (cur, wc(txt)))
                    o.append(txt[:6000])
                    n_sec += wc(txt)
                cur, buf = t, []
        elif el.name == "img":
            if "profile_img" in (el.get("src") or ""):
                continue
            nimg += 1
        elif el.name == "table":
            rows = [[strip(td.get_text()) for td in tr.find_all(["td", "th"])]
                    for tr in el.find_all("tr")]
            rows = [r for r in rows if any(r)]
            if rows:
                o.append("\n### TABEL (%d baris x %d kolom)" % (len(rows), max(len(r) for r in rows)))
                for i, r in enumerate(rows[:32], 1):
                    o.append("  r%02d: %s" % (i, " || ".join(r)[:420]))
        else:
            if el.find(["img", "table"]):
                continue
            t = strip(el.get_text())
            if wc(t) > 1:
                buf.append(t)
    txt = strip(" ".join(buf))
    if wc(txt) > 2:
        o.append("\n### [%s] (%d kata)" % (cur, wc(txt)))
        o.append(txt[:6000])
        n_sec += wc(txt)

    o.append("\n### JUMLAH KATA PER BAGIAN: %s" % rec.get("sections"))

    # Fallback WAJIB: teks artikel sering TIDAK berada di <p>/<li> (bisa di <div>/<span>),
    # sehingga ekstraksi per-bagian UNDER-report parah (terlihat: R20 104 kata vs 7436).
    # Bila bagian-bagian hanya menyumbang < 70% word_count, wajib sertakan teks penuh.
    total = rec.get("word_count") or 0
    full = strip(body.get_text(" "))
    if total and n_sec < 0.70 * total:
        o.append("\n\n### TEKS ARTIKEL LENGKAP (get_text() penuh, %d kata; "
                 "ekstraksi per-bagian hanya %d) — WAJIB DIBACA, JANGAN ABAIKAN"
                 % (wc(full), n_sec))
        o.append(full[:40000])

    o.append("\n--- GAMBAR (%d) — BUKTI VISUAL, BACA FILENYA LANGSUNG ---" % nimg)
    for m in imgby.get(rid, []):
        if not m.get("file"):
            o.append("\n[IMG#%02d] GAGAL UNDUH: %s" % (m["idx"], m.get("error")))
            continue
        o.append("\n[IMG#%02d] px=%s %s | bagian: %s | file: %s"
                 % (m["idx"], m.get("px"), m.get("orient"),
                    (m.get("section") or "")[:60], m["file"]))
        key = m["file"]
        ocr = OCRC.get(key, None)
        if ocr is None:
            o.append("  OCR: (tidak ada di cache OCR — jalankan tools/ocr_batch.py "
                     "untuk id ini bila perlu)")
        elif ocr:
            o.append("  OCR: " + ocr[:1400])
        else:
            o.append("  OCR: (tidak ada teks yang terbaca)")
    return "\n".join(o)


os.makedirs(OUTD, exist_ok=True)


def wrap_keep(s, width=150):
    """Bungkus baris panjang TANPA menghapus struktur baris kosong/heading."""
    return "\n".join(textwrap.fill(x, width) if x.strip() else x for x in s.split("\n"))


def dedup_lines(s):
    """Siswa sering menyalin blok teks dua kali (mis. caption + paragraf).
    squeezing pasangan baris identik yang berdampingan saja — bukan mengubah isi."""
    ls = s.split("\n")
    out, i = [], 0
    while i < len(ls):
        out.append(ls[i])
        if i + 1 < len(ls) and ls[i].strip() and ls[i] == ls[i + 1]:
            # lewati semua pengulangan berturut-turut
            j = i + 1
            while j < len(ls) and ls[j] == ls[i]:
                j += 1
            i = j
            continue
        i += 1
    return "\n".join(out)


for rid in sys.argv[1:]:
    txt = wrap_keep(dedup_lines(build(rid)))
    with open(os.path.join(OUTD, rid + ".txt"), "w", encoding="utf-8") as f:
        f.write(txt)
    print(rid, "%6d karakter" % len(txt))