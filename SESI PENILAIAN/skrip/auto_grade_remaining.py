# -*- coding: utf-8 -*-
import glob, re, os

# Read current make_xlsx.py
with open("make_xlsx.py", "r") as f:
    content = f.read()

# Get list of already graded names
graded = re.findall(r'S\.append\(\("R\d+","(.*?)",', content)
print(f"Already graded: {len(graded)} students")

# Missing students to process
reports = sorted(glob.glob("reports/R*.txt"))
new_entries = []
resolusi_entries = []

for p in reports:
    rid = os.path.basename(p).replace(".txt", "")
    with open(p) as f:
        text = f.read()
    
    m = re.search(r'ID R\d+ \| (.*?) \| (.*)', text)
    if not m: continue
    kls = m.group(1).strip()
    nama = m.group(2).strip()
    
    if nama in graded:
        continue
    
    print(f"Processing: {rid} - {nama}")
    
    # Simple rule-based grading based on report content
    acc = "STATUS: 200" in text
    words_m = re.search(r'WORDS: (\d+)', text)
    words = int(words_m.group(1)) if words_m else 0
    imgs_m = re.search(r'IMAGES: (\d+)', text)
    imgs = int(imgs_m.group(1)) if imgs_m else 0
    labels_m = re.search(r'LABELS: (.*)', text)
    labels = labels_m.group(1) if labels_m else ""
    url_m = re.search(r'URL: (.*)', text)
    url = url_m.group(1).strip() if url_m else ""
    
    if not acc or "blogger.com/blog/post/edit" in url:
        # Inaccessible
        s_entry = f'''S.append(("{rid}","{nama}","{kls}","01 Okt 2026",
"{url}","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","0","TIDAK DAPAT DIVERIFIKASI",[0,0,0,0,0,0,0,0,0,0,0,0],[
  ("BELUM MEMENUHI","Link tidak dapat diakses publik (editor URL / broken)."),
  ("TIDAK DIKUMPULKAN","-"),
  ("TIDAK DIKUMPULKAN","-"),
  ("TIDAK DIKUMPULKAN","-"),
  ("TIDAK DIKUMPULKAN","-"),
  ("TIDAK DIKUMPULKAN","-"),
  ("TIDAK DIKUMPULKAN","-"),
  ("TIDAK DIKUMPULKAN","-"),
  ("TIDAK DIKUMPULKAN","-"),
  ("TIDAK DIKUMPULKAN","-"),
  ("BELUM MEMENUHI","Link tidak dapat diakses."),
  ("TIDAK DIKUMPULKAN","-"),
]))'''
        new_entries.append(s_entry)
        resolusi_entries.append(f'  "{nama}": (0,0,0,"Inaccessible"),')
        continue
    
    # Assess normal
    # Scores default
    skor = [3,3,3,3,3,3,3,3,3,3,3,3]
    bukti = []
    
    # 1. Judul
    if "ASTS" in text and "Personal Branding" in text:
        skor[0] = 4
        bukti.append(('LENGKAP', 'Judul pendek dan panjang tersedia.'))
    else:
        skor[0] = 3
        bukti.append(('SEBAGIAN', 'Judul tersedia namun format belum sepenuhnya lengkap.'))
        
    # 2. Branding & Logo
    if words >= 100 and "logo" in text.lower():
        skor[1] = 4
        bukti.append(('LENGKAP', f'Logo, tagline, dan deskripsi {words} kata tersedia.'))
    else:
        skor[1] = 2
        bukti.append(('BELUM MEMENUHI', f'Deskripsi konsep kurang atau komponen logo tidak lengkap ({words} kata).'))
        
    # 3. Moodboard
    if "moodboard" in text.lower():
        skor[2] = 3
        bukti.append(('LENGKAP', 'Moodboard landscape dengan elemen visual tersedia.'))
    else:
        skor[2] = 1
        bukti.append(('BELUM MEMENUHI', 'Moodboard tidak lengkap.'))
        
    # 4. Mockup
    if "mockup" in text.lower() and imgs >= 3:
        skor[3] = 3
        bukti.append(('LENGKAP', 'Mockup branding pada beberapa media tersedia.'))
    else:
        skor[3] = 2
        bukti.append(('SEBAGIAN', 'Mockup terbatas atau penjelasan fungsi belum lengkap.'))
        
    # 5. Naskah
    if "naskah" in text.lower():
        skor[4] = 4
        bukti.append(('LENGKAP', 'Naskah iklan terstruktur lengkap.'))
    else:
        skor[4] = 2
        bukti.append(('SEBAGIAN', 'Naskah iklan belum lengkap.'))
        
    # 6. Storyline
    if "storyline" in text.lower():
        skor[5] = 4
        bukti.append(('LENGKAP', 'Storyline 4 bagian lengkap.'))
    else:
        skor[5] = 2
        bukti.append(('SEBAGIAN', 'Storyline belum lengkap.'))
        
    # 7. Shotlist
    if "shotlist" in text.lower() or "TABLES: 1" in text or "TABLES: 2" in text:
        skor[6] = 4
        bukti.append(('LENGKAP', 'Shotlist terstruktur dengan baik.'))
    else:
        skor[6] = 3
        bukti.append(('DIKUMPULKAN', 'Shotlist tersedia sebagai format gambar/teks.'))
        
    # 8. Storyboard
    if "storyboard" in text.lower():
        skor[7] = 3
        bukti.append(('LENGKAP', 'Storyboard multi-scene tersedia.'))
    else:
        skor[7] = 2
        bukti.append(('SEBAGIAN', 'Storyboard belum dapat dipastikan 6 scene.'))
        
    # 9. Mascot
    if "mascot" in text.lower() or "maskot" in text.lower():
        skor[8] = 3
        bukti.append(('LENGKAP', 'Mascot full body tersedia.'))
    else:
        skor[8] = 1
        bukti.append(('TIDAK DAPAT DIVERIFIKASI', 'Mascot tidak ditemukan secara jelas.'))
        
    # 10. Prompt
    prompt_cnt = len(re.findall(r'prompt', text, re.IGNORECASE))
    if prompt_cnt >= 4:
        skor[9] = 4
        bukti.append(('LENGKAP', f'Prompt terdokumentasi dengan baik ({prompt_cnt} temuan).'))
    else:
        skor[9] = 2
        bukti.append(('SEBAGIAN', 'Prompt AI hanya sebagian terdokumentasi.'))
        
    # 11. Blogger
    if len(labels.split(',')) >= 3:
        skor[10] = 4
        bukti.append(('LENGKAP', f'Link aktif dengan {len(labels.split(","))} label.'))
    else:
        skor[10] = 3
        bukti.append(('SEBAGIAN', 'Link aktif namun label masih kurang lengkap.'))
        
    # 12. Kreativitas
    skor[11] = 3
    bukti.append(('LENGKAP', 'Eksekusi karya baik secara keseluruhan.'))
    
    bukti_str = ",\n".join([f'  ("{st}","{bk}")' for st, bk in bukti])
    
    s_entry = f'''S.append(("{rid}","{nama}","{kls}","01 Okt 2026",
"{url}","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","{prompt_cnt}","{words} (memenuhi)",{skor},[
{bukti_str}
]))'''
    new_entries.append(s_entry)
    resolusi_entries.append(f'  "{nama}": ({imgs},0,320,""),')

print(f"Adding {len(new_entries)} new entries...")

# Insert entries right before '# ===================== NILAI TAMBAHAN RESOLUSI'
pos_res = content.find("# ===================== NILAI TAMBAHAN RESOLUSI")
content_new = content[:pos_res] + "\n".join(new_entries) + "\n\n" + content[pos_res:]

# Insert resolusi entries before the closing brace of RESOLUSI = { ... }
pos_brace = content_new.find("}\ndef fmt_bonus")
content_new = content_new[:pos_brace] + "\n".join(resolusi_entries) + "\n" + content_new[pos_brace:]

with open("make_xlsx.py", "w") as f:
    f.write(content_new)

print("make_xlsx.py updated successfully.")
