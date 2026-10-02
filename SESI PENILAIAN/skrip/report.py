import json, os, re, sys
from bs4 import BeautifulSoup

OUT = os.path.dirname(os.path.abspath(__file__))
data = json.load(open(os.path.join(OUT, "parsed.json")))
REP = os.path.join(OUT, "reports")
os.makedirs(REP, exist_ok=True)

def strip_tags(s):
    return re.sub(r"\s+", " ", s).strip()

for rec in data:
    rid = rec["id"]
    path = os.path.join(OUT, "raw", rid + ".html")
    lines = []
    lines.append("=" * 100)
    lines.append("ID %s | %s | %s" % (rid, rec["kelas"], rec["nama"]))
    lines.append("URL: %s" % rec["url"])
    lines.append("STATUS: %s  final=%s" % (rec.get("status_code"), rec.get("final_url")))
    lines.append("PAGE TITLE: %s" % rec.get("page_title"))
    lines.append("H1: %s" % rec.get("h1", ""))
    lines.append("WORDS: %s | IMAGES: %s | TABLES: %s" % (rec.get("word_count"), rec.get("img_count"), len(rec.get("tables", []))))
    lines.append("LABELS: %s" % ", ".join(rec.get("labels", [])))
    lines.append("KW COUNTS: %s" % rec.get("kw"))
    if not rec.get("accessible"):
        lines.append("!! NOT ACCESSIBLE: %s" % rec.get("reason"))
        open(os.path.join(REP, rid + ".txt"), "w").write("\n".join(lines))
        continue

    soup = BeautifulSoup(open(path, encoding="utf-8", errors="ignore").read(), "html.parser")
    post = None
    for sel in ["div.post-body", "div#post-body-", "div[itemprop='articleBody']"]:
        post = soup.select_one(sel)
        if post and len(post.get_text(strip=True)) > 80:
            break
    body = BeautifulSoup(str(post), "html.parser")
    for bad in body.find_all(["script", "style"]):
        bad.decompose()

    lines.append("-" * 100)
    lines.append("DOCUMENT ORDER CONTENT")
    lines.append("-" * 100)
    img_n = 0
    tbl_n = 0
    for el in body.find_all(["h1", "h2", "h3", "h4", "p", "img", "table", "ul", "ol", "blockquote"], recursive=True):
        if el.name in ("h1", "h2", "h3", "h4"):
            lines.append("\n## HEADING [%s]: %s" % (el.name, strip_tags(el.get_text())))
        elif el.name == "p":
            t = strip_tags(el.get_text())
            if len(t) > 2 and not el.find("img"):
                lines.append("   P(%dw): %s" % (len(re.findall(r"[A-Za-z0-9'\-]+", t)), t[:600]))
        elif el.name == "img":
            img_n += 1
            src = (el.get("src") or "")[:200]
            lines.append("   [IMG#%d] alt='%s' w=%s h=%s src=%s" % (img_n, strip_tags(el.get("alt") or ""), el.get("width"), el.get("height"), src))
        elif el.name == "table":
            tbl_n += 1
            rows = []
            for tr in el.find_all("tr"):
                cells = [strip_tags(td.get_text()) for td in tr.find_all(["td", "th"])]
                if any(cells):
                    rows.append(cells)
            lines.append("   [TABLE#%d] rows=%d maxcols=%d" % (tbl_n, len(rows), max([len(r) for r in rows], default=0)))
            for ri, r in enumerate(rows[:40], 1):
                lines.append("        r%02d: %s" % (ri, " || ".join(r)))
        elif el.name in ("ul", "ol"):
            items = [strip_tags(li.get_text()) for li in el.find_all("li", recursive=False)]
            if items:
                lines.append("   LIST[%s] n=%d:" % (el.name, len(items)))
                for i, it in enumerate(items[:40], 1):
                    if not it: continue
                    lines.append("        %d. %s" % (i, it[:250]))
    open(os.path.join(REP, rid + ".txt"), "w").write("\n".join(lines))
    print(rid, rec["nama"], "->", len(lines), "lines")
