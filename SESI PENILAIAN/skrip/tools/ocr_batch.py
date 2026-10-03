"""OCR batch lintas-platform (RapidOCR/ONNX) — pengganti Apple Vision untuk Windows.

Gambar kecil di- upscale dulu (padding putih) supaya teks AI pseudoteks lebih terbaca,
meniru langkah `sips -Z 2400` pada pipeline macOS lama.

Pakai:  python ocr_batch.py R20 R25 ...     (tanpa argumen = semua gambar di images.json)
"""
import json, os, sys, time
from concurrent.futures import ThreadPoolExecutor

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
SK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(SK, "tools", "out", "ocr_cache_win.json")


def load(p, default=None):
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {} if default is None else default


def save(p, o):
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(o, f, ensure_ascii=False, indent=1)
    os.replace(tmp, p)


def upscale(path, target=2000):
    """Perbesar gambar kecil + pad putih agar kontras teks AI terbaca."""
    from PIL import Image
    im = Image.open(path)
    if im.mode not in ("RGB",):
        im = im.convert("RGB")
    w, h = im.size
    if max(w, h) >= target:
        return im
    s = target / max(w, h)
    im = im.resize((int(w * s) + 1, int(h * s) + 1), Image.LANCZOS)
    canvas = Image.new("RGB", (im.width + 40, im.height + 40), "white")
    canvas.paste(im, (20, 20))
    return canvas


_ocr = None
_lock = __import__("threading").Lock()


def get_ocr():
    global _ocr
    with _lock:
        if _ocr is None:
            from rapidocr_onnxruntime import RapidOCR
            _ocr = RapidOCR()
    return _ocr


def ocr_one(path):
    # images.json menyimpan path relatif terhadap folder skrip; pastikan absolut.
    ap = path if os.path.isabs(path) else os.path.join(SK, path)
    key = os.path.relpath(ap, SK).replace("\\", "/")
    if not os.path.exists(ap):
        return key, "(berkas tidak ada: %s)" % ap
    try:
        img = upscale(ap)
        res, _ = get_ocr()(img)
    except Exception as e:
        return key, "(OCR error: %s)" % str(e)[:80]
    if not res:
        return key, ""
    parts = []
    for box, txt, conf in res:
        if conf < 0.45:
            continue
        parts.append(txt.strip())
    return key, " / ".join(p for p in parts if p)


def main():
    ids = sys.argv[1:]
    with open(os.path.join(SK, "images.json"), encoding="utf-8") as f:
        manifest = json.load(f)
    todo = [m for m in manifest if m.get("file")
            and (not ids or m["id"] in ids)]
    cache = load(CACHE)
    pending = [m for m in todo
               if os.path.relpath(m["file"], SK).replace("\\", "/") not in cache]
    print("gambar terdaftar: %d | sudah ada di cache: %d | perlu OCR: %d"
          % (len(todo), len(todo) - len(pending), len(pending)), flush=True)
    t0 = time.time()
    done = 0
    with ThreadPoolExecutor(max_workers=3) as ex:
        for key, txt in ex.map(ocr_one, [m["file"] for m in pending]):
            cache[key] = txt
            done += 1
            if done % 10 == 0:
                save(CACHE, cache)
                print("  %d/%d  (%.0fs)" % (done, len(pending), time.time() - t0), flush=True)
    save(CACHE, cache)
    print("Selesai %d gambar dalam %.0fs -> %s" % (len(pending), time.time() - t0, CACHE))


if __name__ == "__main__":
    main()