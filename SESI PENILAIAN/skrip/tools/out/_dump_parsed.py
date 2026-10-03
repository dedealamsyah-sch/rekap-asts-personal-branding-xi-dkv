import json

p = r"D:\MyLearn\rekap-asts-personal-branding-xi-dkv\SESI PENILAIAN\skrip\parsed.json"
with open(p, encoding="utf-8") as f:
    d = json.load(f)

ids = ["R20", "R25", "R88", "R89", "R108", "R67"]
by = {}
for x in d:
    k = x.get("id") or x.get("rid") or x.get("RID")
    if k in ids:
        by[k] = x

out = ["TOTAL ENTRIES: %d" % len(d)]
for k in ids:
    x = by.get(k)
    out.append("===== %s =====" % k)
    if not x:
        out.append("  NOT FOUND")
        continue
    out.append("  nama      = %s" % x.get("nama"))
    out.append("  kelas     = %s" % x.get("kelas"))
    out.append("  timestamp = %s" % x.get("timestamp"))
    out.append("  url       = %s" % x.get("url"))
    out.append("  page_title= %s" % x.get("page_title"))
    out.append("  words     = %s" % x.get("word_count"))
    out.append("  labels    = %s" % (x.get("labels"),))
    out.append("  headings  = %s" % (x.get("headings"),))
    imgs = x.get("images") or []
    out.append("  n_images  = %d" % len(imgs))
    for i, im in enumerate(imgs):
        out.append("    IMG%02d %sx%s" % (i + 1, im.get("w"), im.get("h")))
    tabs = x.get("tables") or []
    out.append("  n_tables  = %d" % len(tabs))
    for ti, t in enumerate(tabs):
        rows = t.get("rows") or []
        out.append("    TBL%d nrows=%d maxcols=%s" % (ti, len(rows), t.get("maxcols")))
        if rows:
            out.append("      header=%s" % (rows[0],))
    out.append("")

open(r"D:\MyLearn\rekap-asts-personal-branding-xi-dkv\SESI PENILAIAN\skrip\tools\out\_parsed_report.txt",
     "w", encoding="utf-8").write("\n".join(out))
print("WROTE _parsed_report.txt for %d ids" % len(by))