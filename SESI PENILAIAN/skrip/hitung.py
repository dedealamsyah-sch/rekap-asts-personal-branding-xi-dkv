import json, os
from collections import OrderedDict

W = [5,12,8,10,8,8,10,10,8,7,10,4]
KOM = ["Penamaan Judul & Identitas","Personal Branding & Logo","Moodboard","Mockup Branding",
       "Naskah Iklan","Storyline","Shotlist","Storyboard","AI Mascot Character",
       "Prompt & Dokumentasi AI","Portfolio Blogger","Kreativitas & Profesionalisme"]

# id, nama, kelas, timestamp, url, skor12
S = [
 ("R02","WAHDAN SAPARI","XI DKV 2","29 Sep 2026 14:36","https://d4aanz.blogspot.com/2026/09/personal-branding.html",[2,3,2,3,4,4,4,2,1,4,4,3]),
 ("R03","AHMAD FAUZI","XI DKV 2","29 Sep 2026 14:38","https://ahmadfll.blogspot.com/2026/09/promptberperanlah-sebagai-profesional.html",[1,3,2,3,3,4,1,1,3,3,2,3]),
 ("R04","QIANDRA KAIZAR NAHARI","XI DKV 3","29 Sep 2026 15:29","https://qiandrakaizarnahari.blogspot.com/2026/09/personal-branding-qiandra-kaizar-nahari.html",[3,2,3,1,4,4,4,3,0,4,4,3]),
 ("R05","ILMA LATIFAH","XI DKV 3","29 Sep 2026 20:43","https://ilmalatifah.blogspot.com/2026/09/uji-kompetensi-ai-prompting-dkv-smkn-9.html",[1,3,3,2,3,4,4,2,2,4,4,3]),
 ("R06","SOPA ANIDATUL AISAH","XI DKV 3","30 Sep 2026 07:14","https://copaanidatul.blogspot.com/2026/09/personal-branding-sopa-anidatul-aisah.html",[3,4,4,3,2,3,4,2,1,4,4,3]),
 ("R07","JAJANG M HUSNI MUBAROK","XI DKV 3","30 Sep 2026 07:18","https://jajangmhusnimubarok.blogspot.com/2026/09/personal-branding-jajang-m-husni.html",[4,3,4,3,3,4,4,4,1,4,4,4]),
 ("R08","ADE SAHRUL GUNAWAN","XI DKV 4","29 Sep 2026 14:25","https://adesahrulgunawan.blogspot.com/2026/09/asts-komputer-grafis-xi-dkv-4.html",[1,1,0,0,0,0,0,0,0,0,2,2]),
 ("R09","DIRA RAHMAWATI","XI DKV 4","29 Sep 2026 19:10","https://dirarahmawati1.blogspot.com/2026/09/asts-komputer-grafis-dira-rahmawati-xi.html",[1,1,0,0,0,0,0,0,0,0,2,1]),
 ("R10","FITRIYANI","XI DKV 4","30 Sep 2026 13:27","https://www.blogger.com/blog/post/edit/3381382262198283975/5845339142879718163",[0,0,0,0,0,0,0,0,0,0,0,0]),
 ("R11","DEDE APRILIA KARTIKA","XI DKV 4","30 Sep 2026 16:43","https://dedeapriliakartika06.blogspot.com/2026/09/personal-branding-dede-aprilia-kartika.html",[4,4,3,2,3,3,4,3,4,4,4,3]),
 ("R12","PUTRI INTAN NURAENI","XI DKV 3","30 Sep 2026 16:49","https://putriintannuraeni.blogspot.com/2026/09/uji-kompetensi-ai-prompting-dkv-smkn-9.html",[1,2,2,2,3,0,4,4,4,3,3,3]),
 ("R13","INTAN WIDIYANTI","XI DKV 3","30 Sep 2026 17:18","https://intanwdworld.blogspot.com/2026/09/uji-kompetensi-ai-prompting-xi-dkv-3.html",[1,3,4,2,2,0,4,2,2,4,3,3]),
 ("R14","SAFINAH SYARA GARINI","XI DKV 1","30 Sep 2026 18:20","https://safinahsyaragarini.blogspot.com/2026/09/personal-branding-safinah-syaara-garini.html",[4,3,4,4,4,4,1,4,4,4,3,4]),
 ("R15","QUINSYA RAHMANESA SOLEHA","XI DKV 3","30 Sep 2026 20:01","https://quinsyarahmanesasoleha.blogspot.com/2026/09/personal-branding-quinsya-rahmanesa.html",[3,2,3,3,2,0,2,1,4,1,2,3]),
 ("R16","JIHAN SHAFIRA KEAN P. M.","XI DKV 3","30 Sep 2026 20:55","https://jihanshafirakean.blogspot.com/2026/09/asts-personal-branding-jihan-shafira-xi.html",[4,3,3,3,4,4,4,2,4,4,4,3]),
 ("R17","DHEA EKA KHOERUNNISA","XI DKV 3","30 Sep 2026 21:01","https://dheaekadhey.blogspot.com/2026/09/personal-branding-dhea-eka-khoerunnisa.html",[3,4,2,2,0,0,4,3,1,1,3,3]),
 ("R18","WULAN SUNDARI","XI DKV 1","30 Sep 2026 21:09","https://wulansundariii.blogspot.com/2026/09/asts-personal-branding-wulan-sundari-xi_01701389558.html",[2,4,4,2,4,3,4,2,4,0,3,3]),
 ("R19","WILDA AZKIA","XI DKV 1","30 Sep 2026 21:12","https://wildaazkia.blogspot.com/2026/09/personal-branding.html",[2,4,4,4,3,4,4,4,4,2,4,4]),
 ("R20","MUHAMAD DIAZ PIRDAUS","XI DKV 1","30 Sep 2026 21:47","https://dias20092806.blogspot.com/2026/09/personal-branding-muhamad-diaz-pirdaus.html",[4,3,4,4,4,4,4,4,4,4,3,4]),
]

def cat(n):
    if n>=90: return "Sangat Baik"
    if n>=80: return "Baik"
    if n>=70: return "Cukup"
    return "Perlu Perbaikan"

rows=[]
for rid,nama,kls,ts,url,sk in S:
    vals=[round(s/4*w,2) for s,w in zip(sk,W)]
    tot=round(sum(vals),2)
    rows.append(dict(id=rid,nama=nama,kelas=kls,ts=ts,url=url,sk=sk,vals=vals,total=tot,cat=cat(tot)))

json.dump(rows, open(os.path.join(os.path.dirname(os.path.abspath(__file__)),"hasil.json"),"w"),ensure_ascii=False,indent=1)

print("REKAP NILAI")
print("| No | Nama | Kelas | Skor per komponen (1-12) | TOTAL | Kategori |")
print("|---|---|---|---|---:|---|")
for i,r in enumerate(rows,1):
    print("| %d | %s | %s | %s | %.2f | %s |"%(i,r['nama'],r['kelas'],",".join(map(str,r['sk'])),r['total'],r['cat']))
print()
print("RINGKASAN")
for c in ["Sangat Baik","Baik","Cukup","Perlu Perbaikan"]:
    n=len([r for r in rows if r['cat']==c]); print("  %-15s %d siswa"%(c,n))
print("  Rata-rata kelas: %.2f"%(sum(r['total'] for r in rows)/len(rows)))
