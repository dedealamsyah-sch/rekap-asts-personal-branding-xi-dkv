"""Unduh gambar asli hanya untuk sekumpulan id (default: kiriman baru saja)."""
import json, os, re, sys, requests
from concurrent.futures import ThreadPoolExecutor
from bs4 import BeautifulSoup

OUT = os.path.dirname(os.path.abspath(__file__))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126.0 Safari/537.36"}
IMG = os.path.join(OUT, "img")
os.makedirs(IMG, exist_ok=True)

with open(os.path.join(OUT, "parsed.json"), encoding="utf-8") as f:
    data = json.load(f)

ids = sys.argv[1:] or None
if ids:
    recs = [r for r in data if r["id"] in ids]
else:
    recs = [r for r in data if r.get("accessible")]
print("target:", len(recs), "entri", flush=True)

from PIL import Image

def is_image(data):
    """Validasi ISI berkas, bukan hanya ukurannya.
    Host Google Photos (lh3.google.com) sering membalas halaman login HTML
    dengan status 200 dan >1 MB — tanpa cek ini, HTML tersimpan sebagai .jpg."""
    if len(data) < 2000:
        return False
    if data[:3] == b"\xff\xd8\xff":          # JPEG
        return True
    if data[:8] == b"\x89PNG\r\n\x1a\n":     # PNG
        return True
    if data[:6] in (b"GIF87a", b"GIF89a"):   # GIF
        return True
    if data[:2] == b"BM" or data[:4] == b"RIFF":  # BMP / WEBP
        return True
    return False


def grab(rec):
    rid = rec["id"]
    d = os.path.join(IMG, rid)
    os.makedirs(d, exist_ok=True)
    raw = os.path.join(OUT, "raw", rid + ".html")
    if not os.path.exists(raw):
        return []
    soup = BeautifulSoup(open(raw, encoding="utf-8", errors="ignore").read(), "html.parser")
    post = None
    for sel in ["div.post-body", "div#post-body-", "div[itemprop='articleBody']"]:
        post = soup.select_one(sel)
        if post and len(post.get_text(strip=True)) > 80:
            break
    if post is None:
        return []
    body = BeautifulSoup(str(post), "html.parser")
    for bad in body.find_all(["script", "style"]):
        bad.decompose()
    out, cur, n = [], "AWAL", 0
    for el in body.find_all(["h1", "h2", "h3", "h4", "img"], recursive=True):
        if el.name in ("h1", "h2", "h3", "h4"):
            t = re.sub(r"\s+", " ", el.get_text()).strip()
            if t: cur = t
            continue
        src = el.get("src") or ""
        if not src.startswith("http") or "profile_img" in src:
            continue
        n += 1
        orig = re.sub(r"=s\d+(\S*)$", r"=s0\1", src)
        if not re.search(r"=s\d+", src):
            orig = src + "=s0"
        fp = os.path.join(d, "%02d.jpg" % n)
        err = None
        got = False
        for cand in (orig, src):
            try:
                r = requests.get(cand, headers=UA, timeout=45)
                if r.status_code != 200:
                    err = "HTTP %s" % r.status_code
                    continue
                if not is_image(r.content):
                    err = "bukan gambar (%d byte, awal=%r)" % (len(r.content), r.content[:24])
                    continue
                open(fp, "wb").write(r.content)
                got = True
                break
            except Exception as e:
                err = str(e)[:80]
        if not got:
            manifest.append({"id": rid, "nama": rec["nama"], "kelas": rec["kelas"], "idx": n,
                             "section": cur, "file": None, "error": err})
            continue
        try:
            with Image.open(fp) as im:
                w, h = im.size
        except Exception:
            w = h = 0
        out.append({"id": rid, "nama": rec["nama"], "kelas": rec["kelas"], "idx": n,
                    "section": cur, "file": os.path.relpath(fp, OUT).replace("\\", "/"),
                    "px": "%dx%d" % (w, h),
                    "orient": "landscape" if w > h else ("portrait" if h > w else "square"),
                    "bytes": os.path.getsize(fp)})
    return out

manifest = []
with ThreadPoolExecutor(max_workers=4) as ex:
    for res in ex.map(grab, recs):
        manifest.extend(res)

# MERGE, jangan timpa: images.json memuat seluruh siswa. Menimpa akan menghapus
# entri siswa lain yang gambarnya sudah terlanjur terunduh.
IMGP = os.path.join(OUT, "images.json")
existing = {}
if os.path.exists(IMGP):
    with open(IMGP, encoding="utf-8") as f:
        try:
            for m in json.load(f):
                existing[(m["id"], m["idx"])] = m
        except Exception:
            existing = {}
for m in manifest:
    existing[(m["id"], m["idx"])] = m
manifest = sorted(existing.values(), key=lambda m: (m["id"], m["idx"]))

with open(IMGP, "w", encoding="utf-8") as f:
    json.dump(manifest, f, ensure_ascii=False, indent=1)

for m in manifest:
    if m.get("file"):
        print("%-5s %-28s #%02d %-40s %-10s %s" % (m["id"], m["nama"][:28], m["idx"],
              m["section"][:40], m["px"], m["orient"]), flush=True)
    else:
        print("%-5s %-28s #%02d GAGAL: %s" % (m["id"], m["nama"][:28], m["idx"], m["error"]), flush=True)
print("TOTAL", len(manifest), "| berhasil", sum(1 for m in manifest if m.get("file")))