import json, os, re, subprocess, sys
import requests
from bs4 import BeautifulSoup

OUT = os.path.dirname(os.path.abspath(__file__))
data = json.load(open(os.path.join(OUT, "parsed.json")))
UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/126.0 Safari/537.36"}
IMG = os.path.join(OUT, "img")
os.makedirs(IMG, exist_ok=True)

SKIP = ("profile_img", "blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEiO")

def dims(path):
    try:
        o = subprocess.run(["sips", "-g", "pixelWidth", "-g", "pixelHeight", path],
                           capture_output=True, text=True).stdout
        w = re.search(r"pixelWidth: (\d+)", o)
        h = re.search(r"pixelHeight: (\d+)", o)
        return (int(w.group(1)) if w else 0, int(h.group(1)) if h else 0)
    except Exception:
        return (0, 0)

manifest = []
for rec in data:
    rid = rec["id"]
    if not rec.get("accessible"):
        continue
    d = os.path.join(IMG, rid)
    os.makedirs(d, exist_ok=True)
    soup = BeautifulSoup(open(os.path.join(OUT, "raw", rid + ".html"), encoding="utf-8", errors="ignore").read(), "html.parser")
    post = None
    for sel in ["div.post-body", "div#post-body-", "div[itemprop='articleBody']"]:
        post = soup.select_one(sel)
        if post and len(post.get_text(strip=True)) > 80:
            break
    body = BeautifulSoup(str(post), "html.parser")
    for bad in body.find_all(["script", "style"]):
        bad.decompose()

    # walk in order, track current heading
    cur = "AWAL"
    n = 0
    for el in body.find_all(["h1", "h2", "h3", "h4", "img"], recursive=True):
        if el.name in ("h1", "h2", "h3", "h4"):
            t = re.sub(r"\s+", " ", el.get_text()).strip()
            if t:
                cur = t
            continue
        src = el.get("src") or ""
        if not src.startswith("http"):
            continue
        if "profile_img" in src:
            continue
        n += 1
        orig = re.sub(r"=s\d+(\S*)$", r"=s0\1", src)
        if not re.search(r"=s\d+", src):
            orig = src + "=s0"
        fp = os.path.join(d, "%02d.jpg" % n)
        try:
            r = requests.get(orig, headers=UA, timeout=60)
            if r.status_code == 200 and len(r.content) > 2000:
                open(fp, "wb").write(r.content)
            else:
                r2 = requests.get(src, headers=UA, timeout=60)
                open(fp, "wb").write(r2.content)
        except Exception as e:
            manifest.append({"id": rid, "nama": rec["nama"], "kelas": rec["kelas"], "idx": n,
                             "section": cur, "file": None, "error": str(e)})
            continue
        w, h = dims(fp)
        manifest.append({"id": rid, "nama": rec["nama"], "kelas": rec["kelas"], "idx": n,
                         "section": cur, "file": os.path.relpath(fp, OUT), "px": "%dx%d" % (w, h),
                         "orient": "landscape" if w > h else ("portrait" if h > w else "square"),
                         "bytes": os.path.getsize(fp)})

json.dump(manifest, open(os.path.join(OUT, "images.json"), "w"), ensure_ascii=False, indent=1)
for m in manifest:
    print("%s %-28s #%02d %-42s %-10s %s" % (m["id"], m["nama"][:28], m["idx"], m["section"][:42], m.get("px", "-"), m.get("orient", "-")))
print("TOTAL", len(manifest))
