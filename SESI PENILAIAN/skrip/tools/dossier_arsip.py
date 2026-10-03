"""Bangun ulang dossier dari HTML yang terarsip (untuk siswa yang link-nya kini mati)."""
import json, os, re, sys, textwrap
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
SK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS = os.path.join(SK, "tools")
OCRC = {}
ocp = os.path.join(TOOLS, "out", "ocr_cache_win.json")
if os.path.exists(ocp):
    with open(ocp, encoding="utf-8") as f:
        OCRC = json.load(f)

strip = lambda s: re.sub(r"\s+", " ", s).strip()
wc = lambda s: len(re.findall(r"[A-Za-z0-9'\-]+", s))


def wrap_keep(s, w=150):
    return "\n".join(textwrap.fill(x, w) if x.strip() else x for x in s.split("\n"))


def main(rid):
    raw = os.path.join(SK, "raw", rid + ".html")
    if not os.path.exists(raw):
        print("TIDAK ADA arsip:", raw); return
    html = open(raw, encoding="utf-8", errors="ignore").read()
    soup = BeautifulSoup(html, "html.parser")
    t = soup.find("title")
    o = []
    o.append("=" * 100)
    o.append("ARSIP HTML LOKAL (link publik kini tidak aktif — dipakai sebagai sumber bukti)")
    o.append("PAGE TITLE: %r" % (strip(t.get_text()) if t else ""))
    o.append("sumber: %s (%d byte)" % (raw, os.path.getsize(raw)))

    post = None
    for sel in ["div.post-body", "div.post-body.entry-content", "div[itemprop='articleBody']",
                "div#post-body-", "div.post-body.element magazine"]:
        post = soup.select_one(sel)
        if post and len(post.get_text(strip=True)) > 80:
            break
    if post is None:
        best, bl = None, 0
        for d in soup.find_all("div"):
            L = len(d.get_text(strip=True))
            if L > bl and L < 60000 and len(d.find_all(["div"])) < 30:
                best, bl = d, L
        post = best
    body = BeautifulSoup(str(post), "html.parser")
    for b in body.find_all(["script", "style"]):
        b.decompose()

    o.append("kata artikel: %d" % wc(body.get_text(" ")))

    cur, buf = "AWAL", []
    nimg = 0
    for el in body.find_all(["h1", "h2", "h3", "h4", "p", "li", "table", "img"], recursive=True):
        if el.name in ("h1", "h2", "h3", "h4"):
            tt = strip(el.get_text())
            if tt:
                txt = strip(" ".join(buf))
                if wc(txt) > 2:
                    o.append("\n### [%s] (%d kata)" % (cur, wc(txt)))
                    o.append(txt[:6000])
                cur, buf = tt, []
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
            tt = strip(el.get_text())
            if wc(tt) > 1:
                buf.append(tt)
    txt = strip(" ".join(buf))
    if wc(txt) > 2:
        o.append("\n### [%s] (%d kata)" % (cur, wc(txt)))
        o.append(txt[:6000])

    labs = sorted({strip(a.get_text()) for a in soup.find_all("a", href=True)
                   if "/search/label/" in a["href"]})
    o.append("\n### LABEL: %s" % labs)

    o.append("\n--- GAMBAR (%d) ---" % nimg)
    for m in [x for x in _load_img() if x["id"] == rid]:
        o.append("\n[IMG#%02d] px=%s %s | bagian: %s | file: %s"
                 % (m["idx"], m.get("px"), m.get("orient"),
                    (m.get("section") or "")[:60], m.get("file")))
        if m.get("file"):
            o.append("  OCR: " + (OCRC.get(m["file"]) or "(tidak ada di cache)")[:1400])

    d = os.path.join(TOOLS, "out", "dossier")
    os.makedirs(d, exist_ok=True)
    out = os.path.join(d, rid + ".txt")
    with open(out, "w", encoding="utf-8") as f:
        f.write(wrap_keep("\n".join(o)))
    print("ditulis:", out, len("\n".join(o)), "karakter")


def _load_img():
    p = os.path.join(SK, "images.json")
    with open(p, encoding="utf-8") as f:
        return json.load(f)


if __name__ == "__main__":
    for r in sys.argv[1:]:
        main(r)