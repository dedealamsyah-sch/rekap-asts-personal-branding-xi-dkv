"""Gabungkan entry penilaian baru ke make_xlsx.py (replace 9 lama + append 18 baru)."""
import ast
import pathlib
import re

TMP = pathlib.Path(__file__).resolve().parent
SRC = pathlib.Path(__file__).resolve().parents[1] / 'make_xlsx.py'
src = SRC.read_text(encoding='utf-8')


def load(path):
    """Jalankan file entries_*.py di ruang lingkup tersembunyi."""
    ns = {'S': [], 'RESOLUSI': {}}
    exec(compile(TMP.joinpath(path).read_text(encoding='utf-8'), path, 'exec'), ns)
    return ns['S'], ns['RESOLUSI']


def fmt(tup):
    rid, nama, kls, ts, url, mk, sh, sc, ms, pr, dsk, skor, det = tup
    head = '%s,%s,%s,%s,' % tuple(repr(x) for x in (rid, nama, kls, ts))
    ringkas = ','.join(repr(x) for x in (url, mk, sh, sc, ms, pr, dsk))
    skor_s = '[' + ','.join(str(x) for x in skor) + ']'
    body = '\n'.join('  (%s,%s),' % (repr(s), repr(b)) for s, b in det)
    return ('S.append((%s\n%s,%s,[\n%s\n]))' % (head, ringkas, skor_s, body))


new_all, res_all = [], {}
for k in range(1, 6):
    s, r = load('entries_%d.py' % k)
    new_all += s
    res_all.update(r)
s, r = load('noaccess.py')
new_all += s
res_all.update(r)
print('entry dari batch:', len(new_all), '| RESOLUSI:', len(res_all))

regrep = {t[1]: fmt(t) for t in new_all if t[1] in
          ('WULAN SUNDARI', 'NURJIHAN', 'SOPA ANIDATUL AISAH', 'PUTRI INTAN NURAENI',
           'INTAN WIDIYANTI', 'DEBI LESTARI', 'DIRA RAHMAWATI', 'AI IMAS', 'SRI AYU WAHYUNI')}

lines = src.split('\n')
out, i, replaced = [], 0, []
while i < len(lines):
    l = lines[i]
    if l.startswith('S.append(('):
        j = i
        while not lines[j].rstrip().endswith(']))'):
            j += 1
        blk = '\n'.join(lines[i:j + 1])
        nama = re.match(r'S\.append\(\("([^"]*)","([A-Z][^"]*)"', blk).group(2)
        if nama in regrep:
            out.append(regrep[nama])
            replaced.append(nama)
            i = j + 1
            continue
        out.extend(lines[i:j + 1])
        i = j + 1
        continue
    if l.startswith('# ===================== NILAI TAMBAHAN RESOLUSI'):
        out += ['# ===================== KIRIMAN BARU / PERBAIKAN 2 OKTOBER 2026 (18 siswa) ======',
                '# Rekapan Google Form bertambah menjadi 86 baris: 18 kiriman baru + 9 siswa yang',
                '# mengirim ulang / memperbaiki karyanya. Nilai di bawah dihitung ulang dari OCR,',
                '# struktur HTML, dan dimensi gambar asli (parser ulang 2 Oktober 2026).', '']
        for t in new_all:
            if t[1] not in regrep:
                out.append(fmt(t))
        out.append('')
        out.append(l)
        i += 1
        continue
    out.append(l)
    i += 1
src = '\n'.join(out)
assert sorted(replaced) == sorted(regrep), 'tidak ketemu: %s' % (set(regrep) - set(replaced))

# ---- RESOLUSI: ganti key lama, tambah key baru ------------------------------
lines = src.split('\n')
keep, seen = [], set()
for l in lines:
    m = re.match(r'^\s*"([A-Z][^"]*)":\s*\(', l)
    if m and m.group(1) in res_all:
        if m.group(1) in seen:
            continue
        seen.add(m.group(1))
        g, ge, mx, note = res_all[m.group(1)]
        keep.append(' "%s": (%s,%s,%s,%s),' % (m.group(1), g, ge, mx, repr(note)))
        continue
    keep.append(l)
lines = keep
i = next(i for i, l in enumerate(lines) if l.startswith('RESOLUSI = {'))
j = next(j for j in range(i, len(lines)) if lines[j].startswith('}'))
tambahan = [' # 2 Oktober 2026'] + [' "%s": (%s,%s,%s,%s),' % (k, v[0], v[1], v[2], repr(v[3]))
                                    for k, v in res_all.items() if k not in seen]
lines[j:j] = tambahan
SRC.write_text('\n'.join(lines), encoding='utf-8')
print('diganti:', len(replaced), replaced)
print('RESOLUSI diperbarui:', len(seen), '| ditambahkan:', len(tambahan) - 1)
print('total S.append:', SRC.read_text(encoding="utf-8").count('S.append(('))
compile(SRC.read_text(encoding='utf-8'), 'make_xlsx.py', 'exec')
print('compile OK')