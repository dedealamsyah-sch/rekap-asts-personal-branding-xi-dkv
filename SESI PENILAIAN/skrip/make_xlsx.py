# -*- coding: utf-8 -*-
"""Generate Excel rekap penilaian ASTS Personal Branding."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Selalu tulis ke root proyek, apa pun direktori kerja saat skrip dijalankan.
_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
OUT = os.path.join(_ROOT, "REKAP NILAIAN ASTS - Personal Branding XI DKV 2026-2027.xlsx")

KOM = ["Penamaan Judul & Identitas","Personal Branding & Logo","Moodboard","Mockup Branding",
       "Naskah Iklan","Storyline","Shotlist","Storyboard","AI Mascot Character",
       "Prompt & Dokumentasi AI","Portfolio Blogger","Kreativitas & Profesionalisme"]
BOBOT = [5,12,8,10,8,8,10,10,8,7,10,4]

# ===================== DATA SISWA =====================
# id, nama, kelas, timestamp, link, mockup, shot, scene, mascot, prompt, deskripsi_katak, skor12, [12 x (status, bukti)]
S = []

S.append(("R02","WAHDAN SAPARI","XI DKV 2","29 Sep 2026 14:36",
"https://d4aanz.blogspot.com/2026/09/personal-branding.html","3 (terverifikasi)","10","TIDAK DAPAT DIVERIFIKASI","1 (full body, wajib terpenuhi)","7","171 (memenuhi)",[2,4,2,3,4,4,4,2,3,4,4,4],[
 ("SEBAGIAN","Judul pendek ada di page title 'Personal Branding Wahdan Sapari XI-DKV-2'; judul panjang 'ASTS ... SMKN 9 Garut' TIDAK ditemukan di artikel."),
 ("LENGKAP","Logo monogram WS + nama branding + tagline 'FILMMAKER | VIDEOGRAPHER' + deskripsi 171 kata. Nama siswa 'WAHDAN SAPARI' tercetak pada logo (OCR)."),
 ("SEBAGIAN","1 gambar 320x179 (landscape). 7 unsur hanya tertulis sebagai echo prompt; tidak ada deskripsi/penjelasan moodboard."),
 ("LENGKAP","3 berkas terpisah: Packaging (laptop box), Laptop, Social Media Feed. Fungsi tiap media diuraikan sangat rinci (3 x 3 poin)."),
 ("LENGKAP","Judul 'The Future Frame', Tema, Pesan Utama, 5 scene lengkap dengan Visual/Audio-SFX/Voice Over, Closing Tagline + makna closing."),
 ("LENGKAP","4 bagian lengkap: Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup."),
 ("LENGKAP","10 shot. Kolom Shot, Jenis Shot, Angle, Movement, Durasi, Deskripsi terisi semua."),
 ("TIDAK DAPAT DIVERIFIKASI","1 gambar 320x179. OCR hanya terbaca 5 label scene (Hook/Intro, Process, The Build, Maks Karya, Outro) - belum dapat dipastikan 6 scene."),
 ("LENGKAP","1 output mascot 3D Character FULL BODY (terverifikasi OCR) - memenuhi ketentuan wajib. Portrait & bersama logo bersifat OPSIONAL dan tidak diunggah."),
 ("LENGKAP","7 prompt terdokumentasi: logo, moodboard, mockup, naskah, shotlist, storyboard, mascot."),
 ("LENGKAP","Link aktif (HTTP 200). 5 label sesuai: AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah."),
 ("LENGKAP","Perencanaan video paling lengkap di XI DKV 2; organisasi artikel tertata dan konsisten dari logo hingga storyboard. Resolusi 320 px TIDAK lagi menjadi pencilan (resolusi rendah tetap diperbolehkan)."),
]))

S.append(("R03","AHMAD FAUZI","XI DKV 2","29 Sep 2026 14:38",
"https://ahmadfll.blogspot.com/2026/09/promptberperanlah-sebagai-profesional.html","3 (terverifikasi OCR)","Ada (format gambar, jumlah TIDAK DAPAT DIVERIFIKASI)","Ada (format gambar, jumlah TIDAK DAPAT DIVERIFIKASI)","3 (terverifikasi OCR)","5","181 (memenuhi)",[1,4,2,3,3,4,2,2,4,3,2,3],[
 ("BELUM MEMENUHI","IDENTITAS PERLU VERIFIKASI. Judul artikel 'Biodata diri' - tidak mengikuti format projek. Judul pendek & panjang tidak ditemukan."),
 ("LENGKAP","Logo inisial AF + aperture + nama 'AF Visual Creator' + tagline + deskripsi 181 kata. Nama siswa 'AHMAD FAUZI' tercetak pada logo (OCR)."),
 ("SEBAGIAN","1 gambar 320x179 (landscape). 7 unsur hanya echo prompt; tidak ada deskripsi moodboard."),
 ("LENGKAP","1 board memuat 3 mockup terverifikasi OCR: '1. LAPTOP', '2. SOCIAL MEDIA FEED', '3. BILLBOARD'. Fungsi tiap media dijelaskan."),
 ("LENGKAP","Judul 'Abadi dalam Cerita', Tema, Pesan Utama, naskah Voice Over lengkap, Closing Tagline."),
 ("LENGKAP","4 bagian lengkap dengan detail adegan, durasi, dan reaksi tokoh."),
 ("DIKUMPULKAN","Shotlist dikumpulkan dalam bentuk GAMBAR (berkas #8 dan #9, keduanya bertuliskan 'SHOTLIST' dengan kolom Angle/Movement/Duration). Sesuai revisi rubrik, format gambar DITERIMA. JUMLAH SHOT TIDAK DAPAT DIVERIFIKASI karena teks baris tabel hasil AI tidak terbaca."),
 ("DIKUMPULKAN","4 berkas panel adegan (#4-#7; OCR terbaca jenis shot, angle, dan movement) + 1 berkas grid 'STORYBOARD IKLAN KOMERSIAL: AF'. Jumlah scene di dalam gambar tidak dapat diverifikasi."),
 ("LENGKAP","1 berkas memuat 3 output terverifikasi OCR: '1. MASCOT FULL BODY' (WAJIB), '2. MASCOT PORTRAIT' dan '3. MASCOT WITH LOGO' (opsional) - lengkap."),
 ("SEBAGIAN","5 prompt (logo, moodboard, mockup, storyboard, mascot). Prompt storyline tidak terdokumentasi. Prompt ditulis dalam bahasa Inggris."),
 ("BELUM MEMENUHI","Link aktif, tetapi TIDAK ADA label/tag sama sekali (0 label) dan judul artikel salah."),
 ("SEBAGIAN","Logo & maskot relevan; tetapi planning video tidak lengkap dan organisasi artikel lemah."),
]))

S.append(("R04","QIANDRA KAIZAR NAHARI","XI DKV 3","29 Sep 2026 15:29",
"https://qiandrakaizarnahari.blogspot.com/2026/09/personal-branding-qiandra-kaizar-nahari.html","TIDAK DAPAT DIVERIFIKASI","10 (terverifikasi)","6 (terverifikasi OCR)","0 (tidak ada gambar)","8","23 (KURANG - harus >=100)",[3,2,3,1,4,4,4,3,0,4,4,3],[
 ("SEBAGIAN","Judul pendek (page title) dan judul panjang (heading ASTS ... SMKN 9 Garut) keduanya ada. Typo: tertulis 'X DKV 3' seharusnya 'XI DKV 3'."),
 ("BELUM MEMENUHI","Logo monogram QKN + nama + tagline 'Tenang dalam Proses, Tegas dalam Karya'. Deskripsi konsep hanya 23 kata (kurang 77). 'by Nama Siswa' tidak ada."),
 ("SEBAGIAN","1024x1024 (PERSGI, bukan landscape). 7 unsur terverifikasi OCR: COLOR PALETTE, TYPOGRAPHY, VISUAL STYLE, DESIGN REFERENCES, TONE & MOOD, GRAPHIC ELEMENTS, VISUAL INSPIRATION."),
 ("TIDAK DAPAT DIVERIFIKASI","2 berkas gambar (1024x1024). Teks menyebut kaos, hoodie, social media feed; jumlah unit di dalam gambar tidak terbaca OCR."),
 ("LENGKAP","Judul 'Keheningan yang Berbicara: Identitas Visual QKN', Tema, Pesan Utama, 4 scene dengan Visual/SFX/VO dan timecode, Closing Tagline."),
 ("LENGKAP","4 bagian lengkap dengan Setting Tempat, Suasana, Pengenalan Tokoh, Hook Visual, Audio & Narasi."),
 ("LENGKAP","10 baris tabel dengan kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi - semua terisi."),
 ("LENGKAP","Berkas gambar memuat SCENE 1 sampai SCENE 6 (terverifikasi OCR) beserta shot, angle, transisi, dan VO."),
 ("TIDAK DIKUMPULKAN","Tidak ditemukan pada hasil pengumpulan. Maskot hanya dijelaskan dalam teks, TIDAK ada berkas gambar maskot."),
 ("LENGKAP","8 prompt terdokumentasi: branding, moodboard, mockup, maskot, naskah, storyline, shotlist, storyboard."),
 ("LENGKAP","Link aktif. 12 label sesuai ketentuan (Branding, Maskot, Mockup, Moodboard, Portofolio, Profil, SMKN 9 Garut, Tugas Sekolah, dll)."),
 ("SEBAGIAN","Planning video terbaik di XI DKV 3; namun satu komponen wajib (AI Mascot) tidak dikumpulkan dan deskripsi terlalu singkat."),
]))

S.append(("R05","ILMA LATIFAH","XI DKV 3","29 Sep 2026 20:43",
"https://ilmalatifah.blogspot.com/2026/09/uji-kompetensi-ai-prompting-dkv-smkn-9.html","TIDAK DAPAT DIVERIFIKASI","12 (terverifikasi)","TIDAK DAPAT DIVERIFIKASI","1 (jenis tidak dapat diverifikasi)","9","296 (memenuhi)",[1,4,3,2,3,4,4,2,2,4,4,3],[
 ("BELUM MEMENUHI","Judul artikel 'UJI KOMPETENSI AI PROMPTING DKV - ILMA LATIFAH (XI DKV 3)'. Tidak menyebut Personal Branding; judul pendek & panjang sesuai format tidak ditemukan."),
 ("LENGKAP","Logo monogram ML + nama 'Ilma Latifah Creative Visuals' + tagline + deskripsi 296 kata (paling mendalam kedua). Nama siswa 'ILMA LATIFAH' tercetak pada logo (OCR)."),
 ("LENGKAP","320x213 landscape. Uraian panjang mencakup warna emas/hitam, tipografi, style visual, referensi desain, tone & mood, elemen grafis,elemen grafis, dan inspirasi visual."),
 ("TIDAK DAPAT DIVERIFIKASI","1 berkas gambar (320x213). Teks menyebut 3 media (kartu nama, stiker, banner) dengan penjelasan; jumlah unit dalam gambar tidak terbaca."),
 ("LENGKAP","Judul 'Kisah Abadi dalam Setiap Frame', Tema, Pesan Utama, 3 scene lengkap Visual + Narator, Closing Tagline."),
 ("LENGKAP","4 bagian lengkap dengan narasi panjang tiap bagian."),
 ("LENGKAP","12 baris tabel; kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi semua terisi."),
 ("TIDAK DAPAT DIVERIFIKASI","1 gambar 'Storyboard Iklan - Kisah Abadi dalam Setiap Frame'. Jumlah scene tidak terbaca OCR."),
 ("TIDAK DAPAT DIVERIFIKASI","1 berkas gambar (320x320). Mascot ada, tetapi tidak dapat dipastikan apakah FULL BODY (wajib) terunggah; portrait & bersama logo (opsional) tidak terverifikasi."),
 ("LENGKAP","9 prompt terdokumentasi (paling banyak kedua)."),
 ("SEBAGIAN","Link aktif, 17 label (terbaik). Namun ada label tidak relevan dengan branding: 'MOCKUP SKINCARE'."),
 ("SEBAGIAN","Isi sangat matang dan dokumentasi AI terbaik; tetapi judul tidak sesuai ketentuan dan beberapa visual sulit diverifikasi."),
]))

S.append(('R20','SOPA ANIDATUL AISAH','XI DKV 3','30 Sep 2026 07:14',
'https://copaanidatul.blogspot.com/2026/09/personal-branding-sopa-anidatul-aisah.html','3 (terverifikasi)','11 (terverifikasi)','6 (terverifikasi teks)','2 (tidak dapat diverifikasi)','8','1696 (memenuhi)',[3,4,4,3,4,3,4,3,2,4,4,3],[
  ('SEBAGIAN',"Page title 'Personal Branding Sopa Anidatul Aisah XI DKV 3 SMKN 9 GARUT' menggabungkan judul pendek dan panjang dalam satu baris sehingga tidak terpisah; paragraf pembuka 'Personal Branding ASTS Sopa Anidatul Aisah XI_DKV 3 SMKN 9 GARUT' mendekati judul panjang yang diminta."),
  ('LENGKAP',"Logo SA dengan ikon kamera. OCR IMG#01 279x279 membaca 'by Sopa Anidatul Aisah'. Nama branding SA Photography, tagline 'Capture the beauty in every moment', dan deskripsi konsep 239 kata (minimum 100 terpenuhi)."),
  ('LENGKAP',"IMG#02 310x207 landscape, OCR membaca 'Cormorant Garamond'. Uraian menjawab 7 unsur: warna utama (soft pink/soft blue), tipografi (Cormorant Garamond), style visual, referensi desain, tone & mood, elemen grafis, dan inspirasi visual."),
  ('LENGKAP',"3 media mockup (Stiker SA, Kartu Nama, Social Media Feed) masing-masing disertai paragraf 'Penjelasan Fungsi Media'; didukung 3 berkas gambar 320x213."),
  ('LENGKAP',"Naskah kini lengkap: Judul 'Abadikan Momen Anda', Tema, Pesan Utama, Narasi/Dialog umum, Slogan Penutup, ditambah 'Narasi/Dialog per Adegan' 6 adegan (Pembukaan 0-3 detik sampai Penutup 15-18 detik) yang memuat Visual, VO, dan Dialog shutter pada tiap adegan."),
  ('SEBAGIAN','Keempat bagian storyline tersedia (Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup) dengan durasi, tetapi masing-masing hanya 1-2 kalimat tanpa perincian visual atau VO per adegan.'),
  ('LENGKAP','Tabel shotlist 12 baris x 7 kolom (header + shot 1-11) dengan kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, dan Deskripsi; seluruh baris terisi.'),
  ('LENGKAP',"Teks section 8 menguraikan storyboard 6 scene berurutan (Scene 1 sampai Scene 6, dari maskot membawa kamera sampai penutup logo + tagline) dan menyatakan tiap scene memuat visual adegan, keterangan shot, angle kamera, transisi, serta dialog/narasi. OCR IMG#07 membaca 'STORYBOARD / video Ikan Caplure rour laoment' tetapi nomor scene di dalam gambar tidak terbaca."),
  ('TIDAK DAPAT DIVERIFIKASI',"Berkas maskot IMG#06 213x320 tidak menghasilkan teks OCR sama sekali; teks hanya menyebut 'karakter 3D Character' tanpa pernyataan bahwa full body (wajib) terunggah, dan tidak ada label Portrait/Bersama Logo."),
  ('LENGKAP',"8 prompt terdokumentasi berlabel 'prompt yang digunakan' untuk logo, moodboard, mockup, maskot, naskah/perencanaan video, storyline, shotlist, dan storyboard."),
  ('LENGKAP','Link aktif, 12 label. Kelima label wajib lengkap: AI, Personal Branding, Portofolio, SMKN 9 Garut, dan tugas sekolah; label tambahan per komponen (MASKOT, MOCKUP, MOODBOARD_SOPA ANIDATUL AISAH) makin terstruktur.'),
  ('SEBAGIAN','Branding dan dokumentasi AI sangat baik dan warna soft pink/soft blue konsisten di semua komponen; tetapi planning video baru secukupnya (storyline 1 kalimat per bagian) dan output maskot full body belum dapat dipastikan.'),
]))

S.append(("R07","JAJANG M HUSNI MUBAROK","XI DKV 3","30 Sep 2026 07:18",
"https://jajangmhusnimubarok.blogspot.com/2026/09/personal-branding-jajang-m-husni.html","3 (terverifikasi OCR)","12 (terverifikasi)","6 (terverifikasi OCR)","1 (full body, wajib terpenuhi)","8","54 (KURANG - harus >=100)",[4,3,4,3,3,4,4,4,3,4,4,4],[
  ("LENGKAP","Judul pendek 'Personal Branding JAJANG M HUSNI MUBAROK XI DKV 3' dan judul panjang 'ASTS ... SMKN 9 Garut' keduanya ada, nama & kelas benar."),
  ("BELUM MEMENUHI","Logo monogram JHM + gunung + kamera; nama JHM; tagline 'Explore. Capture. Create. Be Better.'. 'by jajang' TERBUKTI ada (OCR pada logo, moodboard, mockup, storyboard). NAMUN deskripsi konsep hanya 54 kata."),
  ("LENGKAP","1536x1024 landscape. 7 unsur TERVERIFIKASI OCR lengkap dengan hex color dan nama typeface (Montserrat)."),
  ("LENGKAP","1 board 1536x1024 memuat 3 mockup terverifikasi OCR: T-SHIRT, HOODIE, SOCIAL MEDIA FEED (INSTAGRAM)."),
  ("SEBAGIAN","Judul, Tema, Pesan Utama, Narasi/Dialog 4 scene dengan Narator, dan Closing Tagline tersedia; masih relatif singkat."),
  ("LENGKAP","4 bagian lengkap dengan durasi, visual, dialog maskot, dan audio/SFX."),
  ("LENGKAP","12 baris tabel; seluruh kolom terisi dan logis dengan storyline."),
  ("LENGKAP","Berkas gambar memuat scene 1-6 (terverifikasi OCR: PEMBUKAAN, JUDUL, MENANGKAP MOMEN, HASIL FOTO, KREASI DESIGN, PENUTUP) lengkap shot, angle, movement, durasi, transisi, dialog."),
  ("LENGKAP","1 output mascot 3D Character full body, dinyatakan eksplisit dalam teks dan didukung 1 gambar. Memenuhi ketentuan WAJIB; portrait & bersama logo (opsional) tidak diunggah."),
  ("LENGKAP","8 prompt terdokumentasi untuk seluruh komponen."),
  ("LENGKAP","Link aktif. 12 label sesuai ketentuan."),
  ("SEBAGIAN","Kualitas gambar tertinggi di kelas (1536x1024) dan teks terbaca jelas; tetapi deskripsi konsep terlalu pendek dan maskot kurang."),
]))

S.append(("R08","SILVI BUDIA PUTRI","XI DKV 1","01 Okt 2026",
"https://silvibudiaputri.blogspot.com/2026/09/personal-branding-silvi-budia-putri-xi.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","3 (terverifikasi)","8","4224 (memenuhi)",[4,4,4,4,3,4,4,4,4,4,4,4],[
  ("LENGKAP","Judul pendek dan judul panjang tersedia. Nama lengkap tercantum."),
  ("LENGKAP","Logo monogram SBP + tagline + deskripsi 4224 kata. Nama siswa tercetak pada logo (OCR)."),
  ("LENGKAP","320x179 landscape + uraian 7 unsur moodboard lengkap."),
  ("LENGKAP","3 media mockup: kaos, hoodie, laptop dengan penjelasan filosofi."),
  ("LENGKAP","Judul, Tema, Pesan Utama, Narasi/Dialog, Closing Tagline tersedia."),
  ("LENGKAP","Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup lengkap."),
  ("LENGKAP","10 shotlist lengkap dengan 7 kolom."),
  ("LENGKAP","6 scene dengan shot, angle, transisi, dan dialog."),
  ("LENGKAP","1 berkas memuat output mascot full body, portrait, dan logo."),
  ("LENGKAP","8 prompt terdokumentasi (logo, moodboard, 3 mockup, mascot, naskah, storyline, shotlist, storyboard)."),
  ("LENGKAP","Link aktif. 5 label sesuai ketentuan."),
  ("LENGKAP","Konsistensi branding sangat baik."),
]))

S.append(('R16','NURJIHAN','XI DKV 2','01 Okt 2026 12:11',
'https://nurjihansa01.blogspot.com/2026/09/personal-branding-nurjihan.html','3 (terverifikasi)','10 (terverifikasi)','10 (terverifikasi teks)','3 (terverifikasi OCR)','8','3445 (memenuhi)',[3,4,4,3,4,4,3,4,4,4,3,3],[
  ('SEBAGIAN',"Page title 'PERSONAL BRANDING NURJIHAN' memuat judul pendek tanpa kelas; judul panjang ada pada heading 'ASTS Personal Branding Nurjihan XI DKV 2 SMKN NEGERI 9 GARUT' (ejaan 'SMKN NEGERI 9'). Nama dan kelas benar."),
  ('LENGKAP',"Logo NJ (topi koki, panci masak, kamera DSLR, foto cetak). OCR IMG#01 320x240: 'Nurjihan(tm) / Cook - Capture - Create / - Small steps, big dreams -'. Nama branding dan tagline tertera, deskripsi konsep 218 kata (minimum 100 terpenuhi)."),
  ('LENGKAP','IMG#02 320x180 landscape. Uraian memuat 7 poin bernomor: Logo Utama, Tipografi (3 gaya font), Palet Warna 5 hex (#FEF9EA, #311508, #F58A93, #FABF66, #64723D), Gaya Visual, Tone & Mood, Elemen Grafis, Inspirasi Visual.'),
  ('LENGKAP','Teks uraikan 3 media mockup (Banner, Laptop Skin, Stiker Vinyl) dan masing-masing punya paragraf Keterangan fungsi; berkas gambar mockup 1 buah (IMG#03 320x292) dengan unit di dalamnya tidak terverifikasi.'),
  ('LENGKAP',"Judul 'Mahakarya Kecil dari Dapur', Tema, Pesan Utama, Narasi 5 adegan bertimecode (00:00-00:06 s/d 00:25-00:30) lengkap Visual, SFX, Voice Over, dan Dialog, serta Closing Tagline 'Small steps, big dreams.'"),
  ('LENGKAP','Storyline 4 bagian lengkap dengan Setting, Adegan, Fokus Visual, dan Narasi Utama: 1. Pembukaan, 2. Alur Cerita, 3. Konflik/Fokus Visual, 4. Penutup - ditutup Closing Tagline dan audio penutup.'),
  ('SEBAGIAN',"10 shot tertulis berurutan (1-10) lengkap dengan jenis shot, angle, movement, durasi, dan deskripsi; ditambah poster shotlist IMG#05 (OCR 'TABEL SHOTLIST / SHOTLIST IKLAN SHORT MOVIE (30 DETIK)'). Sesuai Revisi-1c tabel/gambar diterima, tetapi ditulis sebagai daftar teks tanpa kolom sehingga tidak setara tabel utuh."),
  ('LENGKAP','Teks uraikan Scene 1 sampai Scene 10 (melewati minimum 6) dan tiap scene memuat Visual, Spesifikasi Shot, Angle, Transisi, serta Narasi/SFX. Berkas IMG#06 320x240 tidak terbaca OCR, tetapi scene tertulis adalah bukti sah.'),
  ('LENGKAP',"Teks melabeli ketiga output secara eksplisit: 'MASCOT FULL BODY' (WAJIB), 'MASCOT PORTRAIT', 'MASCOT BERSAMA LOGO BRANDING'; OCR IMG#04 320x292 membaca 'MASCOT BERSAMA LOGO BRANDING'. Ketiganya lengkap."),
  ('LENGKAP','8 prompt terdokumentasi dan berlabel (Prompt : / Promt :) untuk logo, moodboard, mockup, maskot, naskah, storyline, shotlist, dan storyboard - menutupi seluruh komponen visual.'),
  ('SEBAGIAN',"Link aktif, 7 label. Empat label wajib ada (AI, personal branding, portofolio, tugas sekolah); label wajib 'SMKN 9 Garut' tidak ada, dan 2 label tidak relevan terhadap projek ('kegiatan bersih bersih', 'CV')."),
  ('SEBAGIAN',"Konsistensi identitas kuat - inisial NJ, palet krem-cokelat-pink, dan tagline small steps big dreams terpakai seragam di logo, moodboard, mockup, maskot, sampai closing. Namun teks masih memuat sisa percakapan AI ('Baik, nah sekarang saya ingin membuat naskah iklan short movie 30 detik')."),
]))

S.append(("R11","DEDE APRILIA KARTIKA","XI DKV 4","30 Sep 2026 16:43",
"https://dedeapriliakartika06.blogspot.com/2026/09/personal-branding-dede-aprilia-kartika.html","TIDAK DAPAT DIVERIFIKASI","10 (terverifikasi OCR)","6 (terverifikasi teks)","3 (terverifikasi OCR)","6","146 (memenuhi)",[4,4,3,2,3,3,4,3,4,4,4,3],[
 ("LENGKAP","Judul pendek 'Personal Branding Dede Aprilia Kartika XI DKV 4' dan judul panjang 'ASTS ... SMKN 9 GARUT' keduanya ada."),
 ("LENGKAP","Logo inisial AK + nama 'AK Creative Studio' + tagline 'Warna Ide, dalam Karya Nyata' + deskripsi 146 kata. 'by Dede Aprilia Kartika' TERBUKTI ada (OCR pada logo & mockup)."),
 ("LENGKAP","320x213 landscape + uraian 7 unsur lengkap."),
 ("TIDAK DAPAT DIVERIFIKASI","1 berkas gambar (320x213). Teks menjelaskan 3 media (tote bag, kartu nama, social media feed) beserta fungsinya; jumlah unit dalam gambar tidak terbaca."),
 ("LENGKAP","Judul, Tema, Pesan Utama, Narasi/Dialog panjang, Closing Tagline + visual closing. Tidak ada judul film terpisah tetapi seluruh unsur ada."),
 ("LENGKAP","4 bagian lengkap dengan deskripsi adegan."),
 ("LENGKAP","10 baris pada tabel shotlist (terverifikasi OCR: kolom No/Adegan/Jenis Shot/Angle/Movement/Durasi/Deskripsi dengan baris 1-10)."),
 ("LENGKAP","6 scene diuraikan lengkap di teks (Scene 1-6 dengan shot, angle, transisi, narasi). Berkas gambar sangat kecil (320x124) sehingga tidak terbaca."),
 ("LENGKAP","1 berkas memuat 3 output terverifikasi OCR: mascot full body (WAJIB), '2 Mascot Potret' dan '3 Mascot Bersama Logo Branding' (opsional) - lengkap."),
 ("LENGKAP","6 prompt (A-F) paling terstruktur: logo, moodboard, mockup, mascot, video iklan, storyboard."),
 ("LENGKAP","Link aktif. 5 label sesuai: AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah."),
 ("SEBAGIAN","Dokumentasi prompt terbaik di kelas dan konsisten dari logo sampai closing; mockup & storyboard perlu resolusi diperbaiki."),
]))

S.append(('R22','PUTRI INTAN NURAENI','XI DKV 3','30 Sep 2026 16:49',
'https://putriintannuraeni.blogspot.com/2026/09/uji-kompetensi-ai-prompting-dkv-smkn-9.html','3 (terverifikasi teks)','10 (terverifikasi)','6 (terverifikasi OCR)','2 (terverifikasi teks)','6','1598 (memenuhi)',[1,3,2,3,3,0,4,4,3,3,3,3],[
  ('BELUM MEMENUHI',"Page title 'ASTS KOMPETENSI AI DKV - SMKN 9 GARUT' dan heading artikel 'ASTS KOMPETENSI AI DKV - PUTRI INTAN NURAENI (XI DKV 3)'. Nama dan kelas ada, tetapi tidak mengikuti format Personal Branding [Nama] [Kelas] maupun judul panjang yang diminta."),
  ('SEBAGIAN',"Logo monogram PIN (P berisi mikrofon, N berupa aperture kamera) dengan nama 'PUTRI INTAN NURAENI' dan tagline 'CREATIVE | SINGING | PHOTOGRAPHY' terbaca OCR IMG#01. Deskripsi konsep kini 227 kata (sebelumnya 50 kata, jadi sudah memenuhi minimum 100), tetapi uraian masih banyak menyapa pembaca dengan kata 'Anda' (identitas Anda, minat Anda) dan belum ada nama branding terpisah dari monogram."),
  ('BELUM MEMENUHI','IMG#02 320x320 PERSEGI, bukan landscape. Uraian moodboard hanya 3 sub-bagian (Judul & Header Atas, Logo Utama, Warna & Tekstur); tipografi, style visual, tone & mood, elemen grafis, referensi desain, dan inspirasi visual tidak diuraikan.'),
  ('SEBAGIAN',"Bagian 'FILOSOFI & ANALISIS MOCKUP IDENTITAS' kini memuat 3 media - Laptop Die-Cut Sticker, Kartu Nama dua sisi, dan Koleksi Stiker - masing-masing dengan paragraf 'Filosofi & Fungsi'. Namun hanya 2 berkas gambar 320x179 sehingga jumlah unit di dalam gambar tidak dapat diverifikasi."),
  ('LENGKAP',"Naskah memuat Judul Iklan 'Karya Berani, Warna Sejati bersama Putri Intan Nuraeni', Tema, Pesan Utama, NARASI & SCENE BREAKDOWN 4 scene lengkap Visual + Audio/SFX + VO, serta CLOSING TAGLINE dengan teks on-screen dan VO penutup. Durasi per adegan tidak dicantumkan."),
  ('TIDAK DIKUMPULKAN',"Heading 'STORYLINE' ada tetapi langsung disusul baris 'Shotlist Video Iklan Personal Branding: Putri Intan Nuraeni (PIN) - Total Durasi: 20s' tanpa satu pun isi. Tidak ditemukan pada hasil pengumpulan."),
  ('LENGKAP','Tabel shotlist 11 baris x 7 kolom (header + shot 1-10) dengan kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, dan Deskripsi Visual & Audio; seluruh baris terisi.'),
  ('LENGKAP','IMG#05 320x320 OCR membaca 6 label scene: SCENE 1 OPENING REVEAL, SCENE 2 CAPTURE FOKUS, SCENE 3 EDITING AJAIB, SCENE 4 EKSPRESI MUSIK, SCENE 5 CALL TO ACTION, SCENE 6 CLOSING TAGLINE - jumlah 6 scene memenuhi minimum.'),
  ('LENGKAP',"Teks menyatakan Panel 1 'Menampilkan maskot secara utuh (full body)' sehingga output wajib Full Body terbukti secara deskripsi teks, dan Panel 2 adalah close-up/portrait. Panel Mascot Bersama Logo hanya diminta pada prompt, tidak terverifikasi pada teks maupun OCR berkas 320x179."),
  ('SEBAGIAN',"6 prompt terdokumentasi berlabel 'PROMPT YANG SAYA GUNAKAN' (logo, moodboard, mockup, maskot, naskah, shotlist). Prompt untuk storyline dan storyboard tidak ada - dua section itu justru kosong."),
  ('SEBAGIAN',"Link aktif dengan 14 label; label wajib 'portofolio' dan 'tugas sekolah' ada, tetapi 'Personal Branding' dan 'SMKN 9 Garut' tidak ada dan judul artikel bukan 'Personal Branding ...'."),
  ('SEBAGIAN',"Monogram PIN dan palet magenta-ungu/emerald konsisten di logo, moodboard, mockup, dan maskot; namun uraian branding masih menggunakan kata sapaan 'Anda', dan dua section (STORYLINE serta isi STORYBOARD) tidak dikembangkan."),
]))

S.append(('R23','INTAN WIDIYANTI','XI DKV 3','30 Sep 2026 17:18',
'https://intanwdworld.blogspot.com/2026/09/uji-kompetensi-ai-prompting-xi-dkv-3.html','3 (terverifikasi teks)','10 (terverifikasi)','6 (terverifikasi teks)','1 (full body, teks)','8','3393 (memenuhi)',[4,4,4,3,4,4,4,4,3,4,4,3],[
  ('LENGKAP',"Page title 'Personal Branding Intan Widiyanti XI DKV 3' sudah persis mengikuti judul pendek, dan heading artikel 'ASTS PERSONAL BRANDING INTAN WIDIYANTI XI DKV 3 SMKN 9 GARUT' memenuhi judul panjang. Nama, kelas, dan sekolah benar."),
  ('LENGKAP',"Logo monogram WD (viewfinder, aperture, pita film) dengan nama 'INTAN WIDIYANTI' terbaca OCR IMG#01 dan profesi 'Visual Director | Photographer'; deskripsi konsep 432 kata (jauh melampaui minimum 100) dengan 4 sub-bagian terperinci."),
  ('LENGKAP','IMG#02 320x179 landscape plus uraian 7 unsur lengkap: Warna Utama (Obsidian Black #1A1A1A, Pure White #FFFFFF, Slate Gray #4A4A4A), Typography (Playfair Display/Cinzel, Montserrat/Inter), Style Visual, Tone & Mood, Elemen Grafis, Referensi Desain, Inspirasi Visual.'),
  ('SEBAGIAN','Teks menguraikan 3 media mockup - Kartu Nama (matte black + emboss emas), Stiker Branding (die-cut collector sheet), dan Roll-Up Banner - beserta fungsi masing-masing; tetapi hanya 1 berkas gambar (IMG#03 320x320) sehingga jumlah unit di dalamnya tidak dapat diverifikasi OCR.'),
  ('LENGKAP',"Naskah kini lengkap: Judul 'Through the Lens of Storytelling', Tema, Pesan Utama, Narasi/Dialog, Closing Tagline, ditambah 'Detail Adegan & Narasi / Dialog' 10 shot bertimecode lengkap dengan Jenis Shot, Angle, Movement, Visual, dan Narasi VO, serta Text on Screen & SFX."),
  ('LENGKAP','Storyline 4 bagian lengkap dengan Adegan, Fokus Narasi, dan Kesan Visual: 1. Pembukaan (Introduction), 2. Alur Cerita (Development), 3. Konflik/Fokus Visual (Climax & Execution), 4. Penutup (Resolution & Call to Action) - ditutup Closing Tagline.'),
  ('LENGKAP',"Tabel shotlist 11 baris x 7 kolom (header + shot 1-10) dengan kolom No, Adegan & Stage, Jenis Shot, Angle, Movement, Durasi, dan Deskripsi; ditutup 'Total Estimasi Durasi: 29 Detik'."),
  ('LENGKAP','Storyboard 6 scene tertulis lengkap (Scene 1 sampai Scene 6) dan tiap scene memuat Keterangan Shot, Angle Kamera, Transisi, serta Dialog/Narasi - memenuhi seluruh ketentuan komponen storyboard. Berkas IMG#05 320x179 tidak terbaca OCR, tetapi scene tertulis adalah bukti sah.'),
  ('LENGKAP',"Teks menyatakan maskot 3D 'karakter full-body' dan Widi tampil 'secara full-body' pada roll-up banner, sehingga output wajib Full Body terbukti lewat deskripsi teks. Hanya 1 berkas maskot (IMG#04 320x320, OCR tidak terbaca) sehingga Portrait dan Bersama Logo yang opsional tidak terverifikasi."),
  ('LENGKAP',"8 prompt terdokumentasi berlabel 'Prompt yang saya gunakan' untuk logo, moodboard, mockup, maskot, naskah, storyline, shotlist, dan storyboard - lengkap dan spesifik."),
  ('LENGKAP',"Link aktif dengan 14 label. Kelima label wajib seluruhnya ada: AI, personal branding, portofolio, SMKN 9 Garut, dan tugas sekolah. Only label ganda 'portopolio' (typo) selain 'portofolio'."),
  ('SEBAGIAN',"Identitas sangat konsisten - maskot 'Widi' berkepala tiga lensa, monogram WD, palet hitam-emas, dan tagline Framing Your Vision ada di logo sampai penutup, dengan konsep orisinal. Namun teks masih memuat 17 artefak '[cite: 3]', menyapa pembaca dengan 'Anda', dan ada blok kode bocor (```html) di section shotlist."),
]))

S.append(("R14","SAFINAH SYARA GARINI","XI DKV 1","30 Sep 2026 18:20",
"https://safinahsyaragarini.blogspot.com/2026/09/personal-branding-safinah-syaara-garini.html","3 (terverifikasi)","Ada (format gambar, jumlah TIDAK DAPAT DIVERIFIKASI)","6 (terverifikasi)","3 (terverifikasi OCR)","8","86 (KURANG - harus >=100)",[4,3,4,4,4,4,2,4,4,4,3,4],[
 ("LENGKAP","Judul pendek 'PERSONAL BRANDING SAFINAH SYARA GARINI XI DKV 1' dan judul panjang 'ASTS ... SMKN 9 GARUT' keduanya ada."),
 ("BELUM MEMENUHI","Logo monogram SSG + nama + tagline 'Stories, Motion, and Life.'. Deskripsi konsep hanya 86 kata (kurang 14). 'by Nama Siswa' tidak ditemukan."),
 ("LENGKAP","446x249 landscape + uraian 6 titik dengan hex color (#0C2C5E, #FFB145, #2DB89B) dan nama typeface."),
 ("LENGKAP","3 berkas terpisah: Kartu Nama, Feed Media Sosial, Packaging/merchandise. Masing-masing ada 'Filosofi & Fungsi'."),
 ("LENGKAP","Judul, Tema, Pesan Utama, RUNDOWN 4 adegan lengkap dengan Visual/Musik-SFX/Voice Over, Closing Tagline."),
 ("LENGKAP","4 bagian lengkap dengan Waktu & Tempat, Deskripsi, dan Fokus Utama."),
 ("DIKUMPULKAN","Shotlist dikumpulkan dalam bentuk GAMBAR (1 berkas 422x237; terbaca 'DOKUMEN SHOT LIST TERPERINCI (60 DETIK)' dengan kolom No/Adegan/Angle/Movement/Durasi/Dialog-VO). Sesuai revisi rubrik, format gambar DITERIMA. JUMLAH SHOT TIDAK DAPAT DIVERIFIKASI karena teks baris tabel hasil AI tidak terbaca."),
 ("LENGKAP","Enam panel terverifikasi (Panel 1-6): Persiapan & Fokus, Semangat Berlari, Adaptasi & Ketenangan, Pengamatan Seni, Menangkap Rasa, Penutup & Identitas - dengan shot, angle, transisi, dialog."),
 ("LENGKAP","1 berkas memuat 3 output terverifikasi OCR: 'Full Body Mascot', 'Portrait View', 'Mascot with Logo'."),
 ("LENGKAP","8 prompt terdokumentasi (logo, moodboard, 3 mockup, mascot, naskah, storyline, shotlist, storyboard)."),
 ("SEBAGIAN","Link aktif dan seluruh komponen ada, tetapi hanya 2 label yang saling duplikat ('PERSONAL BRANDING', 'PERSONAL BRANDING1')."),
 ("SEBAGIAN","Filosofi mockup dan maskot sangat kuat; tetapi shotlist tidak dapat dibuktikan ≥10 shot dan label Blogger tidak sesuai."),
]))

S.append(("R15","QUINSYA RAHMANESA SOLEHA","XI DKV 3","30 Sep 2026 20:01",
"https://quinsyarahmanesasoleha.blogspot.com/2026/09/personal-branding-quinsya-rahmanesa.html","3 (terverifikasi)","8 (BELUM MEMENUHI - kurang 2)","TIDAK DAPAT DIVERIFIKASI","3 (terverifikasi OCR)","1","100 (memenuhi, tepat batas)",[3,2,3,3,2,0,2,1,4,1,2,3],[
 ("SEBAGIAN","Judul pendek & panjang tersedia. H1 artikel berisi 'quinsya blog' sehingga judul utama artikel terpotong."),
 ("BELUM MEMENUHI","Logo QRS + deskripsi tepat 100 kata. NAMA BRANDING dan TAGLINE tidak dinyatakan eksplisit (hanya 'QRS' dan 'Master Every Cut' di storyboard). 'by Nama Siswa' tidak ada."),
 ("LENGKAP","320x179 landscape. 7 unsur bernomor terverifikasi OCR: 1. WARNA UTAMA, 2. TYPOGRAPHY, 3. STYLE VISUAL, 4. ELEMEN GRAFIS, 5. REFERENSI DESAIN, 6. TONE & MOOD, 7. INSPIRASI VISUAL."),
 ("LENGKAP","3 berkas terpisah: Mockup Stiker (Die-cut Vinyl), Mockup Laptop, Mockup Social Media Feed. Fungsi stiker & feed dijelaskan; fungsi laptop tidak."),
 ("BELUM MEMENUHI","Terdapat 7 baris Voice Over/Visual Action, tetapi Judul, Tema, dan Pesan Utama TIDAK diisi (hanya daftar syarat)."),
 ("TIDAK DIKUMPULKAN","Hanya daftar syarat (Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup) tanpa isi."),
 ("BELUM MEMENUHI","8 baris dengan nomor 1,2,3,4,5,6 lalu lompat ke 8 dan 9. MINIMUM 10 SHOT TIDAK TERPENUHI (kurang 2 shot) dan penomoran tidak berurutan."),
 ("TIDAK DAPAT DIVERIFIKASI","1 gambar 320x179 'STORYBOARD: Video Iklan QRS'. Jumlah scene tidak terbaca; di teks hanya daftar syarat."),
 ("LENGKAP","1 berkas memuat 3 output terverifikasi OCR: '1. MASCOT FULL BODY' (WAJIB), '2. MASKOT POTRET' dan '3. MASKOT BERSAMA LOGO BRANDING' (opsional) - lengkap."),
 ("BELUM MEMENUHI","Hanya 1 prompt (logo, ditulis dengan salah ejaan 'promt'). Prompt moodboard, mockup, naskah, storyline, shotlist, storyboard, mascot tidak ada."),
 ("BELUM MEMENUHI","Link aktif, tetapi TIDAK ADA label/tag sama sekali (0 label)."),
 ("SEBAGIAN","Konsep 'Ksatria QRS' orisinal dan konsisten; tetapi shotlist dan storyline tidak memenuhi ketentuan minimum."),
]))

S.append(("R16","JIHAN SHAFIRA KEAN PUTRI MULYADI","XI DKV 3","30 Sep 2026 20:55",
"https://jihanshafirakean.blogspot.com/2026/09/asts-personal-branding-jihan-shafira-xi.html","TIDAK DAPAT DIVERIFIKASI","11 (terverifikasi OCR)","4 (BELUM MEMENUHI - kurang 2)","3 (terverifikasi OCR)","8","153 (memenuhi)",[4,4,3,3,4,4,4,2,4,4,4,3],[
 ("LENGKAP","Judul pendek 'Personal Branding Jihan Shafira XI DKV 3' dan judul panjang 'ASTS ... SMKN 9 Garut' keduanya ada. Nama perlu dipastikan (Jihan Shafira Kean Putri Mulyadi)."),
 ("LENGKAP","Logo monogram 'Sh' (S = not balok) + nama 'SHAFIRA Creative Design' + tagline 'Calm. Premium. Modern.' + deskripsi 153 kata. Nama siswa tercetak pada logo (OCR)."),
 ("SEBAGIAN","Uraian moodboard sangat rinci (palet, pencahayaan, material & tekstur, aksesoris, elemen identitas); 7 unsur tidak dipecah eksplisit seperti format yang diminta."),
 ("TIDAK DAPAT DIVERIFIKASI","4-5 berkas identitas (1024-1376 px). Teks menyebut laptop, packaging, gelas kopi beserta fungsinya; jumlah unit dalam gambar tidak terbaca."),
 ("LENGKAP","Judul 'Harmoni Sebuah Identitas', Tema, Pesan Utama, 4 scene lengkap Visual/Audio-SFX/VO dengan timecode, Closing Tagline."),
 ("LENGKAP","4 bagian lengkap dengan deskripsi adegan panjang."),
 ("LENGKAP","11 baris tabel (terverifikasi OCR: No 1-11, kolom Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi)."),
 ("BELUM MEMENUHI","Hanya 4 scene (Scene 1-4; masing-masing menutupi 2 adegan). MINIMUM 6 SCENE TIDAK TERPENUHI (kurang 2 scene)."),
 ("LENGKAP","1 berkas 'MASCOT SHOWCASE' memuat 3 output terverifikasi OCR: MASCOT FULL BODY, MASCOT PORTRAIT, MASCOT BERSAMA BRANDING."),
 ("LENGKAP","8 prompt terdokumentasi (logo, moodboard, 3 mockup, mascot, naskah, storyline, storyboard)."),
 ("LENGKAP","Link aktif. 5 label sesuai: AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah."),
 ("SEBAGIAN","Shotlist dan naskah sangat detail; tetapi storyboard belum memenuhi minimum 6 scene."),
]))

S.append(("R17","DHEA EKA KHOERUNNISA","XI DKV 3","30 Sep 2026 21:01",
"https://dheaekadhey.blogspot.com/2026/09/personal-branding-dhea-eka-khoerunnisa.html","TIDAK DAPAT DIVERIFIKASI","12 (terverifikasi - terbaik)","TIDAK DAPAT DIVERIFIKASI","1 (jenis tidak dapat diverifikasi)","1","242 (memenuhi)",[3,4,2,2,0,0,4,3,2,1,3,3],[
 ("SEBAGIAN","Judul pendek & panjang ada. Nama pada judul panjang disingkat 'DHEA EKA' (nama lengkap: Dhea Eka Khoerunnisa)."),
 ("LENGKAP","Logo monogram DK + tagline 'Dream. Design. Create.' + deskripsi 242 kata. 'by Dhea Khoerunnisa' TERBUKTI ada (OCR pada logo)."),
 ("BELUM MEMENUHI","320x320 (PERSGI, bukan landscape). Tidak ada deskripsi moodboard - section kosong."),
 ("TIDAK DAPAT DIVERIFIKASI","1 berkas gambar (320x320). Teks menjelaskan 3 media (stiker, Instagram feed, kartu nama); jumlah unit tidak terbaca."),
 ("TIDAK DIKUMPULKAN","Heading '5. Naskah' ada tetapi TANPA ISI. Tidak ditemukan pada hasil pengumpulan."),
 ("TIDAK DIKUMPULKAN","Heading '6. Storyline' ada tetapi TANPA ISI. Tidak ditemukan pada hasil pengumpulan."),
 ("LENGKAP","12 baris dengan 7 kolom lengkap - kualitas shotlist terbaik di kelas. Konsisten dengan karakter humble/glassmorphism."),
 ("SEBAGIAN","1 berkas gambar (300x320). Teks menyebut 'delapan scene' dengan detail; jumlah scene di dalam gambar tidak terbaca OCR."),
 ("TIDAK DAPAT DIVERIFIKASI","1 berkas gambar (213x320). Mascot ada, tetapi tidak dapat dipastikan apakah FULL BODY (wajib) terunggah; opsional tidak terverifikasi."),
 ("BELUM MEMENUHI","Hanya 1 prompt (storyboard). Prompt logo, moodboard, mockup, shotlist tidak terdokumentasi."),
 ("SEBAGIAN","Link aktif. 15 label - PALING BAIK di kelas dan sesuai ketentuan."),
 ("SEBAGIAN","Branding dan shotlist sangat baik; tetapi dua komponen wajib (naskah & storyline) kosong."),
]))

S.append(('R02','WULAN SUNDARI','XI DKV 1','30 Sep 2026 21:09',
'https://wulansundariii.blogspot.com/2026/09/asts-personal-branding-wulan-sundari-xi_01701389558.html','TIDAK DAPAT DIVERIFIKASI','12 (terverifikasi)','6 (terverifikasi teks)','3 (terverifikasi OCR)','0','2445 (memenuhi)',[2,4,4,2,4,4,4,3,4,0,3,3],[
  ('BELUM MEMENUHI',"Page title 'ASTs Personal Branding Wulan Sundari XI DKV 1 SMKN 9 GARUT' hanya memuat judul panjang; judul pendek 'Personal Branding [Nama] [Kelas]' tidak ada. URL juga membocorkan id Blogger (_01701389558)."),
  ('LENGKAP',"Logo monogram WS (OCR IMG#01 1024x1024: 'WULAN SUNDARI / CREATIVE ART & DESIGN / By WULAN SUNDARI') + tagline 'Create with Calm, Express with Art.' + deskripsi konsep 211 kata (melewati minimum 100)."),
  ('LENGKAP','IMG#02 1376x768 landscape. OCR memverifikasi 7 unsur: WARNA UTAMA BRANDING (#F9B0A3, #87BAB3, #CFF5F1), TYPOGRAPHY, STYLE VISUAL, TONE & MOOD (CALM, SERENE, ELEGANT, CREATIVE, ARTISTIC, HARMONIOUS), ELEMEN GRAFIS, REFERENSI DESAIN, INSPIRASI VISUAL.'),
  ('TIDAK DAPAT DIVERIFIKASI','Teks uraikan 3 media mockup (gaun floral, flat shoes, set stiker die-cut) dan hanya fungsi stiker yang dijelaskan; berkas gambar mockup cuma 1 (IMG#03 1376x768) dan jumlah unit di dalamnya tidak terbaca OCR.'),
  ('LENGKAP',"Naskah: Judul 'Setiap Karya Punya Cerita', Tema, Pesan Utama, tabel Narasi/Dialog + VO 7 baris (0-8 dtk s/d 56-60 dtk), Closing Tagline, dan VO Final."),
  ('LENGKAP',"Section 6 STORYLINE kini 393 kata dan tersusun 4 bagian lengkap ber-durasi: 1. PEMBUKAAN 'Mencari Inspirasi' 0-10 detik, 2. ALUR CERITA 'Inspirasi Menjadi Karya' 10-30 detik, 3. KONFLIK/FOKUS VISUAL 'Menemukan Identitas' 30-48 detik, 4. PENUTUP 'Inilah Identitasku' 48-60 detik - tiap bagian memuat adegan, VO, dan fokus visual."),
  ('LENGKAP','Tabel HTML 13 baris x 6 kolom (header + shot 1-12) dengan kolom No. Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi - semua terisi dari ELS sampai ECU to CU.'),
  ('LENGKAP',"Teks section 8 (201 kata) menyebut papan berisi 6 kotak adegan tersusun 2 baris x 3 panel dan merinci Adegan 1-6 (Long Shot, MCU). OCR IMG#06 1376x768 membaca kolom 'No. Adegan / Visual Adegan / Angle Kamera / Transisi / Dialog/Narasi' baris 1-5, tetapi isi transisi dan dialog berupa pseudoteks."),
  ('LENGKAP',"IMG#04 1376x768 OCR 'MASCOT FULL BODY / MASCOT PORTRAIT / MASCOT BERSAMA LOGO BRANDING' - full body (wajib) dan kedua opsional terpenuhi sekaligus."),
  ('TIDAK DIKUMPULKAN','Kata prompt tidak muncul sama sekali (0 kemunculan) pada seluruh artikel; tidak ada satu pun prompt terdokumentasi.'),
  ('SEBAGIAN',"Link publik aktif (HTTP 200) dan seluruh komponen ada, tetapi hanya 1 label Blogger ('personal branding'); 4 dari 5 label wajib (Tugas Sekolah, AI, Portofolio, SMKN 9 Garut) tidak ada."),
  ('SEBAGIAN','Identitas visual sangat konsisten (peach coral + soft teal, monogram WS) dari logo sampai mockup; tetapi prompt tidak terdokumentasi sama sekali, mockup belum dapat diverifikasi, dan URL blog membocorkan id Blogger.'),
]))

S.append(("R19","WILDA AZKIA","XI DKV 1","30 Sep 2026 21:12",
"https://wildaazkia.blogspot.com/2026/09/personal-branding.html","3 (terverifikasi)","10 (terverifikasi OCR)","6 (terverifikasi)","3 (terverifikasi)","1","157 (memenuhi)",[2,4,4,4,3,4,4,4,4,2,4,4],[
 ("BELUM MEMENUHI","Judul panjang 'ASTS Personal Branding Wilda Azkia XI DKV 1 SMKN 9 Garut' ada. Judul pendek hanya 'PERSONAL BRANDING' - tanpa nama dan kelas."),
 ("LENGKAP","Logo inisial WA + sapuan kuas + nama 'Wilda Azkia' + tagline 'Creating Rhythm in Every Design' + deskripsi 157 kata. 'BY WILDA AZKIA' TERBUKTI ada (OCR pada logo & mockup)."),
 ("LENGKAP","513x287 landscape + 5 sub-bagian uraian. OCR memverifikasi: COLOR PALETTE, VISUAL STYLE & REFERENCE, TYPOGRAPHY, TONE & MOOD, GRAPHIC ELEMENTS."),
 ("LENGKAP","3 berkas terpisah: Mockup Stiker, Mockup Kartu Nama, Mockup Social Media Feed. Tiap media punya 4 poin penjelasan fungsi (447+411+423 kata)."),
 ("LENGKAP","Judul 'Turning Concepts into Colors', Tema, Pesan Utama, Narasi/Voice Over, Closing Tagline."),
 ("LENGKAP","4 bagian lengkap dengan deskripsi adegan dan fokus visual."),
 ("LENGKAP","10 baris tabel (terverifikasi OCR: No 1-10 dengan kolom Adegan, Jenis Shot, Angle, Movement, Durasi)."),
 ("LENGKAP","Scene 1 sampai Scene 6 lengkap dengan Shot, Angle, Transisi, Visual, dan Narasi."),
 ("LENGKAP","3 berkas terpisah pada section mascot: 441x246 (full body - WAJIB), 295x295 (portrait), 545x304 (bersama logo) - ketiganya opsional terpenuhi."),
 ("BELUM MEMENUHI","Hanya 1 prompt (logo). Prompt moodboard, mockup, shotlist, storyboard, dan mascot tidak terdokumentasi."),
 ("LENGKAP","Link aktif. 8 label sesuai ketentuan."),
 ("LENGKAP","Watermark 'BY WILDA AZKIA' diminta lewat prompt dan benar-benar muncul di logo serta 3 mockup. Konsistensi warna dan bentuk pita di seluruh komponen."),
]))

S.append(("R20","MUHAMAD DIAZ PIRDAUS","XI DKV 1","30 Sep 2026 21:47",
"https://dias20092806.blogspot.com/2026/09/personal-branding-muhamad-diaz-pirdaus.html","3 (terverifikasi)","11 (terverifikasi OCR)","6 (terverifikasi)","3 (terverifikasi OCR)","8","163 (memenuhi)",[4,4,4,4,4,4,4,4,4,4,3,4],[
 ("LENGKAP","Judul pendek 'Personal Branding Muhamad Diaz Pirdaus Xl DKV 1' dan judul panjang 'ASTS ... SMKN 9 GARUT' keduanya ada. Nama & kelas benar."),
 ("SEBAGIAN","Logo monogram DZ + siluet pelari sprint + nama + tagline 'FORGE AHEAD. LIVE WITH PURPOSE.' + deskripsi 163 kata. Nama siswa 'MUHAMAD DIAZ PIRDAUS' tercetak pada logo (OCR). Catatan: heading menulis '{Monogram SSG}' - teks milik siswa lain (tidak orisinal)."),
 ("LENGKAP","1376x768 landscape. 7 unsur TERVERIFIKASI OCR: WARNA UTAMA BRANDING, ELEMEN GRAFIS, TONE & MOOD, TYPOGRAPHY, REFERENSI DESAIN, STYLE VISUAL, INSPIRASI VISUAL."),
 ("LENGKAP","3 media terverifikasi: A. LAPTOP, B. BANNER, C. KAOS. Masing-masing ada 'FILOSOFI DAN FUNGSI' (desain + fungsi bisnis) yang sangat lengkap."),
 ("LENGKAP","Judul 'Sepersekian Detik', Tema, Pesan Utama, 5 ACT dengan Visual/SFX/VO dan timecode, Closing Tagline."),
 ("LENGKAP","4 bagian lengkap dengan Suasana, Adegan, Mobilitas, dan Konflik Visual."),
 ("LENGKAP","11 baris tabel pada berkas gambar (terverifikasi OCR: No 1-11 dengan kolom Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi)."),
 ("LENGKAP","SCENE 1 sampai SCENE 6 lengkap dengan Keterangan Shot, Angle Kamera, Transisi, dan Dialog/Narasi."),
 ("LENGKAP","1 berkas memuat 3 output terverifikasi OCR: FULL BODY MASCOT (WAJIB), MASCOT PORTRAIT dan MASCOT BERSAMA LOGO BRANDING dengan logo di dada (opsional) - lengkap."),
 ("LENGKAP","8 prompt terdokumentasi: logo, moodboard, 3 mockup, mascot, naskah, storyline, shotlist, storyboard."),
 ("SEBAGIAN","Link aktif dan seluruh komponen ada, tetapi hanya 1 label ('PERSONAL BRANDING'), ada heading ganda 'FILOSOFI DAN FUNGSI ;', dan uraian shotlist di teks terpotong pada shot 9."),
 ("SEBAGIAN","Mockup dengan filosofi + fungsi bisnis paling matang di kelas; tetapi ada salah ketik monogram dan teks terpotong."),
]))

def x_rows(row):
    rid, nama, kls, ts, link, mk, sh, sc, ms, pr, dsk, skor12, det = row
    vals = [round(s / 4 * b, 2) for s, b in zip(skor12, BOBOT)]
    rubrik = round(sum(vals), 2)
    bon, ket = bonus_resolusi(nama)
    akhir = min(100.0, round(rubrik + bon, 2))
    return vals, rubrik, bon, ket, akhir, kategori(rubrik), kategori(akhir)

S.append(("R21","RIZKY MUHAMMAD REGAL SAFARI","XI DKV 1","01 Okt 2026 08:14",
"https://rizkymregalsafari.blogspot.com/2026/09/asts-personal-branding-rizky-muhammad.html","TIDAK DAPAT DIVERIFIKASI","11 (terverifikasi - bentuk teks)","6 (terverifikasi OCR)","3 (terverifikasi OCR)","8 (tanpa label)","TIDAK ADA (section berisi prompt)",[3,2,2,1,4,4,4,3,4,3,2,3],[
 ("SEBAGIAN","Page title memuat judul panjang: 'ASTS Personal Branding Rizky Muhammad Regal Safari XI DKV 1 SMKN 9 GARUT'. Judul pendek 'Personal Branding [Nama] [Kelas]' tidak terpisah jelas - di artikel hanya ada 2 heading kecil."),
 ("BELUM MEMENUHI","Logo 'RR / RIZKY REGAL' (OCR) + nama branding 'Rizky Regal' + tagline 'Tetap Berjuang Untuk Menggapai Impian'. Nama siswa tercetak pada logo. NAMUN DESKRIPSI KONSEP BRANDING TIDAK ADA - bagian A berisi prompt, bukan uraian konsep minimal 100 kata."),
 ("BELUM MEMENUHI","1 gambar 320x320 PERSEGI (diminta landscape). 7 unsur moodboard TERBUKTI ada lewat OCR: WARNA UTAMA BRANDING, STYLE VISUAL, TYPOGRAPHY, TONE & MOOD, ELEMEN GRAFIS, INSPIRASI VISUAL. Tidak ada deskripsi moodboard."),
 ("BELUM MEMENUHI","1 gambar 320x320 (board mockup). Prompt menyebut 3 media: stiker, kartu nama, gelas kopi. PENJELASAN FUNGSI MEDIA TIDAK ADA - bagian C hanya berisi prompt tanpa ura fungsi."),
 ("LENGKAP","Judul 'Membangun Kesuksesan Bersama Teman', Tema, Pesan Utama, 6 adegan bertimecode lengkap Visual + SFX + Narasi/Dialog, Closing Tagline 'Rizky Regal. Bersama meraih mimpi sukses.'"),
 ("LENGKAP","4 bagian lengkap dalam tabel bertimecode: Pembukaan (00:00-00:08), Alur Cerita (00:08-00:18), Konflik/Fokus Visual (00:18-00:33), Penutup (00:50-01:00) - dengan detail adegan, karakter, dan narasi."),
 ("LENGKAP","11 shot. Kolom No, Adegan, Jenis Tembakan, Sudut, Gerakan, Durasi, Deskripsi Visual & Aksi - semua terisi. PENTING: tabel ditulis sebagai TEKS markdown, bukan tabel HTML, sehingga tidak terbaca oleh penghitung tabel otomatis."),
 ("LENGKAP","1 gambar 320x320. OCR terbaca 6 scene bernomor: 1. PEMBUKAAN CERIA, 2. KESEHARIAN CERIA, 3. KONFLIK KOMUNIKASI, 4. PEMAHAMAN BARU, 5. IMPLEMENTASI SOLUSI, 6. KESUKSESAN BERSAMA - lengkap dengan timecode."),
 ("LENGKAP","1 berkas memuat 3 output terverifikasi OCR: 'MASCOT FULL BODY (BOY & GIRL)' (WAJIB), 'MASCOT PORTRAIT' dan 'MASCOT BERSAMA LOGO BRANDING' (opsional) - lengkap."),
 ("SEBAGIAN","8 prompt lengkap dan sangat spesifik (A-H: logo, moodboard, mockup, maskot, naskah, storyline, shotlist, storyboard) - hampir paling lengkap di kelas. NAMUN tidak diberi label 'Prompt' sama sekali, sehingga dokumentasinya tidak terstruktur."),
 ("BELUM MEMENUHI","Link aktif dan seluruh komponen ada, TETAPI artikel tanpa heading section (5 gambar ditumpuk tanpa judul bagian). Label tidak sesuai: 6 label, 3 di antaranya milik tugas lain ('Riwayat hidup', 'Tugas penyetingan cahaya'). Label 'Tugas Sekolah', 'AI', 'Portofolio' tidak ada."),
 ("SEBAGIAN","Ide sangat kreatif - konsep duo maskot Ryu & Rina, tema gaming RP, dan cerita konflik komunikasi. Namun susunan artikel berupa dinding teks tanpa heading, dan 5 label tidak relevan."),
]))

S.append(("R22","KHANZA NURAENI","XI DKV 3","01 Okt 2026 08:49",
"https://khanzanuraeni.blogspot.com/2026/09/personal-branding-khanza-nuraeni-xi-dkv.html","TIDAK DAPAT DIVERIFIKASI","10 (terverifikasi)","6 (terverifikasi OCR)","3 (terverifikasi OCR)","6","47 (KURANG - harus >=100)",[4,2,3,3,3,4,4,4,4,3,4,3],[
 ("LENGKAP","Judul pendek pada page title 'Personal Branding Khanza Nuraeni XI DKV 3'; judul panjang pada artikel 'ASTS Personal Branding Khanza Nuraeni XI DKV 3 SMKN 9'. Nama, kelas, dan SMKN 9 Garut sesuai."),
 ("BELUM MEMENUHI","Logo 'KHNZA ARTWORKS' + nama branding + tagline 'Bold Ideas, Clean Designs' + nama siswa 'by Khanza Nuraeni' (OCR). NAMUN deskripsi konsep hanya 47 kata (kurang 53 dari minimum 100)."),
 ("SEBAGIAN","1 gambar 1376x768 LANDSCAPE. 7 unsur terverifikasi OCR: COLOR PALETTE, TYPOGRAPHY, STYLE VISUAL, TONE & MOOD, ELEMENTS & INSPIRATION, GRID LINES, Modern Minimalism. CACAT: gambar masih memuat artefak sitasi dari AI ('(cite: 4)', '](cite: 1,4')."),
 ("TIDAK DAPAT DIVERIFIKASI","1 gambar 1195x896 (board multi-panel). Teks menjelaskan 3 media (kartu nama, gelas kopi, social media feed) beserta fungsi masing-masing; jumlah unit di dalam gambar tidak terbaca."),
 ("LENGKAP","Judul 'The Power of Precision', Tema, Pesan Utama, Narasi/Voice Over lengkap, Closing Tagline. Narasi berupa 1 paragraf tanpa breakdown scene."),
 ("LENGKAP","4 bagian lengkap: Pembukaan, Alur Cerita & Proses, Konflik Visual & Puncak, Penutup - dengan deskripsi adegan dan warna."),
 ("LENGKAP","10 baris tabel. Kolom: No, Visual (Adegan), Audio/VO, Teknis Kamera (jenis shot + angle + movement digabung), Durasi. Semua data wajib tersedia, format kolom digabung."),
 ("LENGKAP","1 gambar 1408x768. Teks menyebut grid 3x2 = 6 scene. OCR verifikasi tiap adegan memuat SHOT, ANGLE, TRANS, dan NARRATION."),
 ("LENGKAP","1 berkas memuat 3 output terverifikasi OCR: 'Mascot Full Body' (WAJIB), 'Mascot Portrait' dan 'Mascot Bersama Logo Branding' (opsional) - lengkap."),
 ("SEBAGIAN","6 prompt terdokumentasi dan diberi label jelas (Prompt:): logo, deskripsi, moodboard, mockup, storyboard, maskot. Prompt untuk naskah, storyline, dan shotlist tidak ada."),
 ("LENGKAP","Link aktif. Label PERSIS sesuai soal: AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah (5 dari 5). Seluruh 10 isi Blogger tersedia."),
 ("SEBAGIAN","Identitas visual konsisten dan profesional (Cyan/Teal + Dark Charcoal). Catatan: moodboard memuat artefak sitasi AI yang should've dibersihkan sebelum diunggah."),
 ]))

# ---------- KIRIMAN BARU 01 OKT 2026 10:11 - 11:16 (13 siswa) ----------

S.append(("R19","CEISHA SINTHIA","XI DKV 2","01 Okt 2026 10:11",
 "https://www.blogger.com/blog/post/edit/1670849221285353060/1910546013362529910","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI",[0,0,0,0,0,0,0,0,0,0,0,0],[
  ("TIDAK DAPAT DIVERIFIKASI","Link yang dikirim adalah alamat EDITOR Blogger (https://www.blogger.com/blog/post/edit/...), bukan link artikel publik.gjrequire login sehingga karya tidak dapat diakses sama sekali."),
  ("TIDAK DAPAT DIVERIFIKASI","-"),
  ("TIDAK DAPAT DIVERIFIKASI","-"),
  ("TIDAK DAPAT DIVERIFIKASI","-"),
  ("TIDAK DAPAT DIVERIFIKASI","-"),
  ("TIDAK DAPAT DIVERIFIKASI","-"),
  ("TIDAK DAPAT DIVERIFIKASI","-"),
  ("TIDAK DAPAT DIVERIFIKASI","-"),
  ("TIDAK DAPAT DIVERIFIKASI","-"),
  ("TIDAK DAPAT DIVERIFIKASI","-"),
  ("TIDAK DAPAT DIVERIFIKASI","Link terkirim bukan URL artikel publik. Blogger loves mengarahkan ke editor bila artikel berstatus DRAFT."),
  ("TIDAK DAPAT DIVERIFIKASI","-"),
 ]))

S.append(("R20","SELVI SIFA URIZQI","XI DKV 3","01 Okt 2026 10:24",
 "https://selviasifaurizqi.blogspot.com/2026/09/asts-personal-branding-selvi-sifa.html","3 kelompok / 5 media (terverifikasi teks)","12 (terverifikasi)","6 (terverifikasi OCR)","1 (terverifikasi teks)","7","315 (memenuhi)",[2,4,4,4,4,4,4,3,2,4,4,4],[
  ("SEBAGIAN","Page title 'ASTS PERSONAL BRANDING SELVI SIFA URIZKI XI DKV 3' - memuat nama & kelas, tetapi tidak menyebut 'SMKN 9 Garut' dan tidak ada judul pendek terpisah 'Personal Branding [Nama] [Kelas]'."),
  ("LENGKAP","Logo emblem lingkaran 'Modern Retro Badge' dengan inisial SL (lalu)+ bentuk gunung, nama branding 'Sesell Sifau Photography & Videography', tagline, dan deskripsi konsep sangat lengkap. Nama siswa tercetak (OCR)."),
  ("LENGKAP","320x179 landscape. Tujuh unsur terverifikasi OCR: WARNA UTAMA BRANDING, TYPOGRAPHY (Avenir Next), STYLE VISUAL, REFERENSI DESAIN, TONE & MOOD, ELEMEN GRAFIS, INSPIRASI VISUAL. Ditambah uraian panjang 3 paragraf."),
  ("LENGKAP","3 kelompok mockup diurai sangat rinci: (1) Outdoor & Expedition Gear (ransel patch + tumbler), (2) Studio & Creative Workspace (mug + stiker + laptop + jurnal), (3) Exclusive Merchandise (plushie dalam window box). Fungsi tiap media dijelaskan."),
  ("LENGKAP","Judul 'Cerita di Setiap Sudut Lensa', Tema, Pesan Utama, Narasi/Dialog lengkap dengan Visual + Narator untuk tiap adegan, Closing Tagline 'Sesell Sifau: Tangkap Momennya, Ceritakan Kisahnya.'"),
  ("LENGKAP","4 bagian lengkap: Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup - masing-masing dengan deskripsi adegan naratif."),
  ("LENGKAP","12 shot, kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi semua terisi (tabel HTML terbaca)."),
  ("SEBAGIAN","1 gambar 320x179 'STORYBOARD' (tanpa teks terbaca). Narasi deskripsi sangat lengkap dan menyebut gaya 3D dengan 6 beat cerita, tetapi scene, angle, transisi, dan dialog TIDAK dapat diverifikasi dari gambar."),
  ("SEBAGIAN","Maskot 3D 'Sesell Sifau' dij extensively di teks (peak mountaineer, kamera mirrorless, jurnal, jaket army, ransel patch, dua toples biskuit & pretzel) - konsep sangat kuat. NAMUN tidak ada berkas gambar maskot tersendiri; dari 5 gambar tidak ada yang jelas merupakan maskot full body."),
  ("LENGKAP","7 prompt terdokumentasi dengan label jelas 'prompt yang di gunakan': logo, moodboard, mockup, maskot, naskah, storyline, shotlist, storyboard."),
  ("LENGKAP","Link aktif (HTTP 200). 5 label sesuai ketentuan: AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah."),
  ("LENGKAP","Karya paling matang di XI DKV 3. Uraian tiap komponen sangat panjang dan spesifik, palet warna & tipografi konsisten antarvisual, dan Planning video lengkap dari naskah sampai shotlist 12 shot."),
 ]))

S.append(("R21","AZMI ANUGRAH","XI DKV 3","01 Okt 2026 10:28",
 "https://azmianugrah.blogspot.com/2026/09/uji-kopetensi-promtpting-ai-dkv.html","3 (terverifikasi teks)","12 (terverifikasi)","6 (terverifikasi OCR)","TIDAK DAPAT DIVERIFIKASI","7","176 (memenuhi)",[1,4,4,3,4,3,4,3,2,4,4,3],[
  ("BELUM MEMENUHI","Judul artikel 'UJI KOmpETENSI PROMTPTING AI DKV' - tidak mengikuti format projek. Judul pendek & panjang sesuai ketentuan tidak ditemukan; halaman hanya punya 1 heading tanpa nama kelas lengkap."),
  ("LENGKAP","Logo monogram 'AN' + inisial, nama branding, tagline 'PLAY SHOOT CREATE', deskripsi konsep 176 kata yang mengaitkan inisial AN dengan bola voli & kamera. Nama siswa 'AZMI ANUGRAH' tercetak pada logo (OCR)."),
  ("LENGKAP","320x213 landscape. Tujuh unsur lengkap: warna (hitam/biru/putih/emas), tipografi, style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual - dengan uraian 2 paragraf."),
  ("LENGKAP","3 media (kaos, laptop, stiker) diurai dengan fungsi masing-masing. Namun hanya 1 dari 5 gambar yang tampak/mockup jelas; sebagian visual mockup tidak terverifikasi melalui OCR."),
  ("LENGKAP","Judul 'PLAY. SHOOT. CREATE.', Tema, Pesan Utama, 6 SCENE lengkap dengan Visual + Voice Over bertimecode, Closing Tagline."),
  ("SEBAGIAN","Teks menyebut 'STORILINE' (ejaan salah) dan langsung memuat tabel shotlist; tidak ada uraian 4 bagian (Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup). Storyline praktis tidak diuraikan terpisah."),
  ("LENGKAP","12 shot, kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi - semua terisi dan logis."),
  ("SEBAGIAN","1 gambar 320x292 'STORYBOARD | PLAY SHOOT CREATE' (OCR). Jumlah scene, angle, transisi, dan dialog TIDAK dapat diverifikasi; teks hanya memuat prompt, bukan deskripsi hasil."),
  ("TIDAK DAPAT DIVERIFIKASI","Teks mendeskripsikan maskot 3D streetwear sporty lengkap dengan simbol mahkota, bola voli, dan kamera; dari 5 gambar tidak ada berkas yang dapat dipastikan sebagai maskot full body."),
  ("LENGKAP","7 prompt terdokumentasi dengan label 'PROMPT YANG SAYA GUNAKAN': logo, moodboard, mockup, maskot, naskah, shotlist, storyboard."),
  ("LENGKAP","Link aktif. 15 label -PetraPyx, termasuk 5 label wajib: AI, Portofolio, SMKN 9 Garut, Tugas Sekolah, Personal Branding (tulisan 'BRENDING LOGO')."),
  ("SEBAGIAN","Branding, dokumentasi AI, dan shotlist sangat baik. Namun planning video belum lengkap: storyline tidak diuraikan, storyboard & mockup sulit diverifikasi, dan maskot tidak terkumpul jelas."),
 ]))

S.append(("R22b","INDRI FITRIYANI","XI DKV 3","01 Okt 2026 10:28",
 "https://indriifitriyani.blogspot.com/2026/09/personal-branding-indri-fitriyani-xi.html","3 (terverifikasi teks)","10 (terverifikasi)","6 (terverifikasi teks)","1 (terverifikasi OCR)","8","384 (memenuhi)",[4,4,4,3,4,4,4,2,3,4,4,4],[
  ("LENGKAP","Page title 'Personal Branding Indri Fitriyani XI DKV 3' (judul pendek sesuai) dan heading pertama 'ASTS PERSONAL BRANDING / INDRI FITRIYANI XI DKV 3 / SMKN 9 GARUT' (judul panjang sesuai, ditulis 3 baris)."),
  ("LENGKAP","Logo monogram 'IF' (I + F menyatu dengan elemen kamera/lensa) + nama branding + tagline 'Capture Moments, Create Stories' + deskripsi konsep 384 kata termasuk bagian FILOSOFI PERSONAL BRANDING tersendiri. Nama siswa tercetak (OCR)."),
  ("LENGKAP","400x223 landscape. Uraian 4 paragraf menjawab seluruh 7 unsur moodboard (warna soft luminous gold/pastel blue/deep navy/warm beige, tipografi Cormorant Garamond, style visual minimalis, referensi desain, tone & mood, elemen grafis, inspirasi visual) + Daftar pustaka 2 sumber."),
  ("LENGKAP","3 media (kartu nama, stiker, paper cup) dengan uraian fungsi dan alasan penempatan. Hanya 1 dari 5 gambar yang dapat dipastikan sebagai mockup."),
  ("LENGKAP","Judul 'Capture Your Story, Create Your Memory', Tema, Pesan Utama, Narasi/Voice Over 5 baris berurutan, Closing Tagline 'Indri Fitriyani - Photographer | Editor.'"),
  ("LENGKAP","4 bagian lengkap dengan heading tegas: NASKAH, STORYLINE, SHOTLIST, STORYBOARD."),
  ("LENGKAP","10 shot, kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi - semua terisi."),
  ("SEBAGIAN","1 gambar 400x266 'STORYBOARD' dengan OCR labels (2x2, 4 panel). Jumlah scene 6 tidak dapat dipastikan; hanya tampak 4 panel dalam gambar."),
  ("SEBAGIAN","1 gambar 320x320 'Capture Moments / Create Stories' - ada maskot 3D perempuan dengan kamera, logo IF dan nama padaecutable. Full body dapat diperkirakan, portrait & bersama logo tidak terkumpul."),
  ("LENGKAP","8 prompt terdokumentasi dengan label 'prompt yang di gunakan', termasuk iterasi: '-edit lagi moodboard nya gambar gambarnya sesuaikan sesuai warna branding'."),
  ("LENGKAP","Link aktif. 7 label sesuai ketentuan termasuk AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah."),
  ("LENGKAP","Karya paling lengkap bersama M Reza Huafah: ada FILOSOFI branding terpisah, Daftar pustaka, iterasi prompt, dan enam heading utama yang terstruktur rapi."),
 ]))

S.append(("R23","MEISYA FAKHRIYAH","XI DKV 3","01 Okt 2026 10:37",
 "https://meisyafakhriyah.blogspot.com/2026/09/ase sen-sumatif-tengah-semester.html","3 (terverifikasi teks)","10 (terverifikasi)","6 (terverifikasi teks)","1 (terverifikasi OCR)","5","602 (memenuhi - terpanjang)",[2,4,4,3,4,4,4,3,3,3,4,4],[
  ("SEBAGIAN","Page title 'PROJEK ASTS PERSONAL BRANDING MEISYA FAKHRIYAH XI DKV 3 SMKN 9 GARUT' - memuat judul panjang, tetapi tidak ada judul pendek terpisah 'Personal Branding [Nama] [Kelas]'. Heading halaman hanya 1 ('D.3 SHOTLIST IKLAN KOMERSIAL')."),
  ("LENGKAP","Logo inisial 'MY' + bunga mawar, nama branding 'MEISYA', tagline 'Beauty That Blooms', deskripsi konsep 602 kata - paling panjang di kelas - lengkap dengan palet hex (#DAAF98, #E6C7C2, #F5E6D8, #EEEEEF, #C179A6B). Nama siswa tercetak."),
  ("LENGKAP","Uraian moodboard sangat detail: palet hex, tipografi serif + sans-serif, style visual 'clean, classy, soft luxury', referensi desain, tone & mood, elemen grafis, inspirasi visual, dan grid layout landscape. Deskripsi paling matang di kelas."),
  ("LENGKAP","3 media (stiker, kartu nama, packaging) dengan prompt eksplisit 'Tampilkan minimal 3 mockup dalam satu komposisi' + uraian bahan, tekstur, dan pencahayaan. Hanya 1 dari 5 gambar yang dapat dipastikan sebagai mockup."),
  ("LENGKAP","Judul 'The secret of blooming beauty', Tema, Pesan Utama, Narasi/Voice Over 4 blok bertimecode (00:00-00:06 sampai 00:22-00:30) lengkap dengan catatan tone suara, Closing Tagline."),
  ("LENGKAP","4 bagian lengkap: Pembukaan, Perkembangan Cerita (Build-Up), Fokus Visual/Produk (Climax/Product Showcase), Penutup - masing-masing dengan detail visual dan pencahayaan."),
  ("LENGKAP","10 shot, kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi - semua terisi dan sangat teknis. Tabel TERDUP GANDA (2 salinan identik) - tidak mengurangi nilai."),
  ("SEBAGIAN","Hanya 1 dari 5 gambar adalah storyboard (320x174, OCR terbatas). Deskripsi 6 scene ada di teks lengkap dengan Judul Konsep, Tema Visual, Palet Warna, dan Nuansa Brand, tetapi scene, angle, transisi, dan dialog per scene tidak dirinci."),
  ("SEBAGIAN","1 gambar 179x320 'STUDIO DKV: COMMERCIAL AD PROJECT' - terdeteksi sebagai maskot perempuan streetwear. Hanya 1 dari 3 versi (full body, portrait, bersama logo) yang terkumpul; full body dapat diperkirakan."),
  ("SEBAGIAN","5 prompt terdokumentasi: branding, moodboard, mockup, maskot, video commercial (D.1-D.4). Prompt paling lengkap dari empat sub-bagian dalam satu prompt. Prompt naskah, storyline, shotlist, storyboard TIDAK terdokumentasi terpisah."),
  ("LENGKAP","Link aktif. 7 label sesuai ketentuan termasuk AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah."),
  ("LENGKAP","Paling estetik & paling sistematis: brand Unlike & premium yang konsisten, 3 planning video (naskah, storyline, shotlist, storyboard) lengkap, dan instruksi prompt paling presisi (termasuk 'tidak terlihat seperti hasil AI')."),
 ]))

S.append(("R24","M REZA HUAFAH","XI DKV 1","01 Okt 2026 10:42",
 "https://mrezahuafah.blogspot.com/2026/09/personal-branding-m-reza-huafah-xi-dkv-1.html","3 (terverifikasi teks)","12 (terverifikasi OCR)","6 (terverifikasi OCR)","3 (terverifikasi teks)","8","481 (memenuhi)",[4,4,4,3,4,4,4,3,4,4,1,4],[
  ("LENGKAP","Page title 'Personal Branding M Reza Huafah XI DKV 1' dan heading pertama 'ASTS Personal Branding M Reza Huafah XI DKV 1 SMKN 9 Garut' - keduanya sesuai ketentuan."),
  ("LENGKAP","Logo simbol abstrak MRH + nama branding + tagline 'DREAM • EXPLORE • CREATE' + deskripsi konsep 481 kata yang sangat lengkap, termasuk makna tiap warna, bentuk, dan slogan. Nama siswa tercetak (OCR)."),
  ("LENGKAP","320x213 landscape. Uraian 6 paragraf menjawab 7 unsur dengan detail luar biasa: warna dengan kode hex (#FFB000, #0D0D0D, #EAEAEA, #6B6B6B), tipografi Montserrat Extra Bold/Medium/Regular, style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual."),
  ("LENGKAP","3 media (packaging, gelas kopi, kaos) diurai 3 paragraf terpisah dengan fungsi, warna, dan elemen spesifik. 3 dari 27 gambar terdeteksi sebagai mockup (320x320 dengan OCR MRH)."),
  ("LENGKAP","Judul 'DREAM • EXPLORE • CREATE', Tema, Pesan Utama, Narasi/Dialog 5 scene (Pembukaan, DREAM, EXPLORE, CREATE, dan penutup) lengkap dengan Visual + Narasi, Closing Tagline."),
  ("LENGKAP","4 bagian lengkap dengan detail visual per bagian: 1. PEMBUKAAN, 2. ALUR CERITA, 3. KONFLIK / FOKUS VISUAL, 4. PENUTUP. Termasuk detail warna, pose maskot, dan elemen transisi."),
  ("LENGKAP","12 shot. Shotlist dikumpulkan sebagai 20+ GAMBAR (format gambar DITERIMA menurut Revisi-1c). OCR terbaca shot 1-13 dengan kolom Angle, Movement, Durasi, dan Deskripsi lengkap."),
  ("SEBAGIAN","1 gambar 320x213 'STORYBOARD PERSONAL BRANDING VIDEO' (OCR: MRH, 6 scene). Jumlah scene, angle, transisi, dan dialog TIDAK dapat diverifikasi dari gambar; deskripsi teks menyebut 6 scene (Scene 1-6) lengkap dengan angle dan transisi."),
  ("LENGKAP","Maskot 3D diuraikan sangat lengkap: pakaian serba hitam dengan aksen oranye dan putih, hoodie, celana cargo, sepatu, tas ransel, full body + close-up + pose setengah badan. 1 gambar 320x213 (OCR: MRH) sebagai maskot; ketentuan WAJIB full bodylaporkan, opsional tidak lengkap."),
  ("LENGKAP","8 prompt terdokumentasi dengan label 'Prompt:' atau 'Promt:' yang konsisten, memuat seluruh komponen dari logo sampai storyboard."),
  ("BELUM MEMENUHI","Link aktif dan artikel lengkap, TETAPI hanya 1 label: 'Tugas sekolah ASTS PROJEK'. Label wajib 'Personal Branding', 'AI', 'Portofolio', 'SMKN 9 Garut' tidak ada."),
  ("LENGKAP","Karya paling lengkap di XI DKV 1: 27 gambar, 3.467 kata, seluruh rubrik terisi (kecuali label Blogger), planning video paling detail, dan shotlist 12 shot lengkap dengan angle, movement, durasi, dan transisi."),
 ]))

S.append(("R25","SYIVA WIDIYANA AGUSTIN","XI DKV 1","01 Okt 2026 10:43",
 "https://syivawidiyana.blogspot.com/2026/09/personal-branding-syiva-widiyana-xi-dkv.html","3 (terverifikasi teks)","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","1 (terverifikasi OCR)","7","276 (memenuhi)",[3,3,4,3,4,4,2,2,3,4,2,3],[
  ("SEBAGIAN","Page title 'Personal Branding - Syiva Widiyana A - XI DKV 1' - memuat nama & kelas tetapi nama belakang disingkat 'A' (seharusnya Agustin). Judul panjang tersedia sebagai paragraf pembuka 'ASTS Personal Branding Syiva Widiyana Agustin XI DKV 1 SMKN 9 Garut'. Halaman tanpa heading (0 h1-h4)."),
  ("SEBAGIAN","Logo monogram 'SW' dengan elemen kamera, buku, pesawat, not musik, gunung, matahari, bintang; tagline/keterangan 'PHOTOGRAPHY • TRAVEL • BOOKS • MUSIC'. Deskripsi konsep cukup (mendekati 100 kata). Nama tercetak (OCR)."),
  ("LENGKAP","320x213 landscape. Uraian 7 paragraf menjawab 7 unsur: warna (cream/beige/cokelat/terracotta), tipografi (serif + script), style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual, dan Kesimpulan."),
  ("LENGKAP","3 media (kaos, gelas keramik, hoodie) diurai sangat rinci termasuk presentasi flat-lay, beberapa sudut, close-up detail, dan visual lifestyle. 6 dari 8 gambar mockup terdeteksi."),
  ("LENGKAP","Judul 'Capture Your Story', Tema, Durasi, Konsep, Narasi/Dialog 3 scene lengkap (OPENING, PHOTOGRAPHY, TRAVEL & EXPLORATION) dengan Visual + VO, Closing Tagline."),
  ("LENGKAP","4 bagian lengkap: 1. PEMBUKAAN, 2. ALUR CERITA, 3. KONFLIK / FOKUS VISUAL, 4. PENUTUP - dengan narasi panjang tiap bagian."),
  ("SEBAGIAN","Shotlist TIDAK berbentuk tabel HTML dan tidak ada gambar bertuliskan 'shotlist' yang jelas. Hanya narasi deskripsi shotlist yang panjang, dan 1 gambar 137x306 portrait dengan teks kecil. Jumlah shot TIDAK DAPAT DIVERIFIKASI."),
  ("SEBAGIAN","Tidak ada gambar bertuliskan 'storyboard'. Hanya narasi deskripsi storyboard dengan Scene 1-6 lengkap (Perkenalan Diri, Fotografi sebagai Passion, dll) lengkap dengan Narasi. Scene, angle, transisi TIDAK dapat diverifikasi dari gambar."),
  ("SEBAGIAN","1 gambar 137x306 dengan maskot berhijab (terverifikasi OCR deskripsi: kamera, tas, headphone). Hanya 1 versi; full body dapat diperkirakan, portrait & bersama logo tidak terkumpul."),
  ("LENGKAP","7 prompt terdokumentasi dengan label 'Promot:' (typo) dan 'Prompt:' secara konsisten: logo, moodboard, mockup kaos, mockup gelas, mockup hoodie, naskah, storyline, storyboard, shotlist."),
  ("SEBAGIAN","Link aktif, 6 label: CV, Penyetingan Cahaya, Projek, Tugas Branding Logo, Tugas Sekolah, Tugas Sekolah ASTS. Label 'AI' dan 'Personal Branding' tidak ada; 2 label tidak relevan ('CV', 'Penyetingan Cahaya')."),
  ("SEBAGIAN","Branding dan moodboard sangat baik, planning video lengkap di teks (naskah, storyline, storyboard). NAMUN shotlist dan storyboard tidak terkumpul dalam bentuk tabel/gambar yang dapat diverifikasi, dan label Blogger tidak sesuai."),
 ]))

S.append(("R26","MEYLAN MELIYANTI ANASTASYA SOFYAN","XI DKV 3","01 Okt 2026 11:01",
 "https://meylanchiza.blogspot.com/2026/09/uji-kompetensi-ai-prompting-dkv-smkn-9.html","3 (terverifikasi teks)","10 (terverifikasi)","6 (terverifikasi teks)","3 (terverifikasi teks)","8","689 (memenuhi - terpanjang)",[2,4,4,3,4,4,4,3,4,4,4,4],[
  ("SEBAGIAN","Page title 'ASTS PERSONAL BRANDING MEYLAN - SMKN 9 GARUT' dan heading 'ASTS PERSONAL / BRANDING - MEYLAN MELIYANTI XI DKV 3 SMKN 9 GARUT' (terbaca sebagai satu judul). Judul pendek 'Personal Branding [Nama] [Kelas]' tidak terpisah."),
  ("LENGKAP","Logo monogram 'MA' dengan elemen musik (not balok, mikrofon, soundwave, headphone, sparkle), nama branding 'MEYLAN ANASTASYA', tagline 'Bright Voice, Happy Soul', deskripsi konsep 689 kata - paling panjang di kelas - dengan makna warna, bentuk, dan tipografi terperinci. Nama tercetak (OCR)."),
  ("LENGKAP","320x179 landscape. Uraian sangat lengkap dengan sub-bagian eksplisit: Konsep Utama Branding, Makna Warna, Konsep Tipografi, Konsep Logo dan Identitas Musik, Tone dan Mood, Gaya Visual, Inspirasi Visual, Referensi Desain dan Penerapan Brand. 7 unsur dengan label/judul yang jelas."),
  ("LENGKAP","3 media (hoodie, tumbler, laptop + stiker) diurai dengan 3 kategori produk (Pakaian, Produk Lifestyle, Media Digital & Sticker) dan tujuan masing-masing. 3 dari 5 gambar terdeteksi sebagai mockup."),
  ("LENGKAP","Judul 'Suara yang Bersinar, Jiwa yang Bahagia', Tema, Pesan Utama, Narasi/Dialog 6 baris Voice Over lengkap dengan jeda, Closing Tagline 'Meylan Anastasya. Bright Voice, Happy Soul.'"),
  ("LENGKAP","4 bagian lengkap dengan detail visual dan setting: Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup. Alur cerita sangat sinematik (gelombang suara berbakura pink, transformasi ruangan)."),
  ("LENGKAP","10 shot, kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi - semua terisi dan sangat teknis dengan timecode."),
  ("SEBAGIAN","Hanya 1 dari 5 gambar adalah storyboard (320x179, OCR sangat terbatas: VIEH, LEVH, MA, Low Angle, Eye Level). Narasi deskripsi 6 scene ada lengkap dengan VO, tetapi scene, angle, transisi, dan dialog TIDAK dapat diverifikasi dari gambar."),
  ("LENGKAP","Maskot 3D diuraikan sangat lengkap dan filosofis: maskot bernyanyi, warna pink sesuai branding, 3 output (full body, portrait, bersama logo). 1 gambar 320x320 (OCR: MEYLAN ANASTASYA) dapat dipastikan sebagai maskot. Ketentuan WAJIB terpenuhi."),
  ("LENGKAP","8 prompt terdokumentasi dengan label 'prompt yang digunakan' atau 'prompt yang di gunakan' - seluruh komponen tercakup."),
  ("LENGKAP","Link aktif. 10 label sesuai ketentuan termasuk AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah."),
  ("LENGKAP","Karya paling lengkap bersama Meisya: 3.411 kata, 8 komponen rubrik dengan uraian panjang dan terstruktur, brand konsisten (pink pastel + unsur musik), dan planning video paling sinematik dengan 10 shot bertimecode."),
 ]))

S.append(("R27","PUTRI UTAMI","XI DKV 2","01 Okt 2026 11:05",
 "https://putriutami0709.blogspot.com/2026/09/personal-branding-putri-utami-xi-dkv-2_01447693375.html","3 (terverifikasi teks)","8 (terverifikasi OCR)","6 (terverifikasi OCR)","3 (terverifikasi OCR)","9","551 (memenuhi)",[4,4,4,3,3,4,3,4,4,4,4,4],[
  ("LENGKAP","Page title 'PERSONAL BRANDING PUTRI UTAMI XI DKV 2' (judul pendek sesuai) dan heading pertama 'ASTS PERSONAL BRANDING PUTRI UTAMI XI DKV 2 SMKN 9 GARUT' (judul panjang sesuai)."),
  ("LENGKAP","Logo inisial 'PU' dengan warna gradasi cerah, nama branding 'PU Creative Journey', tagline 'Dance • Sing • Capture • Create • Inspire', deskripsi konsep 551 kata dengan breakdown lengkap (elemen fotografi, musik, tari, desain grafis, warna, gaya). Nama tercetak (OCR)."),
  ("LENGKAP","320x213 landscape. Uraian 7 sub-bagian dengan emoji: Warna Utama Branding (6 warna + makna), Typography, Style Visual, Referensi Desain, Tone & Mood, Elemen Grafis, Inspirasi Visual, dan Kesimpulan dengan konsep 'Same Person • Many Passions • One Journey'."),
  ("LENGKAP","3 media (kartu nama, banner, billboard) diurai sangat rinci termasuk bagian depan/belakang, elemen utama, informasi kontak, dan penyajian mockup. 6 dari 8 gambar terdeteksi sebagai mockup (OCR: PU Creative Journey)."),
  ("SEBAGIAN","Judul 'The Creative Striker: Unleash Your Power', Tema, Pesan Utama, Closing Tagline tersedia. NAMUN Narasi/Dialog hanya disebut ada tanpa isi - tidak ada breakdown adegan, timing, atau VO per scene."),
  ("LENGKAP","4 bagian lengkap dengan adegan sinematik: 1. Pembukaan (dunia digital/cyber, mata helm menyala), 2. Alur Cerita (hero landing, ledakan cat pelangi), 3. Konflik/Fokus Visual, 4. Penutup - dengan detail visual dan VO."),
  ("SEBAGIAN","Shotlist dikumpulkan dalam bentuk GAMBAR (1 berkas 320x213, OCR 'OFFICIAL SHOTLIST & PRODUCTION LOG' dengan kolom Jenis Shot, Angle, Movement, Durasi, Deskripsi). Ada 8 baris terbaca (1a-3b, 2a-2c) = 8 shot, KURANG dari minimum 10. Sesuai Revisi-1c format gambar diterima, jumlah belum_genap."),
  ("LENGKAP","Berkas gambar 320x213 'SHORT FILM STORYBOARD (6 PANEL)' - 6 panel terverifikasi OCR: BOOT-UP, PENDARATAN DINI, LEDAKAN WARNA, MONTAGE AKSI, PERTARUNGAN IDE, POSE PEMENANG. Narasi deskripsi 6 scene lengkap dengan Angle, Transisi, dan Dialog/Narasi."),
  ("LENGKAP","1 berkas 320x213 memuat 3 output terverifikasi OCR: 'MASCOT FULL BODY - THE CREATIVE STRIKER' (WAJIB), 'MASCOT PORTRAIT', dan 'MASCOT BERSAMA LOGO BRANDING' (opsional) - lengkap. Maskot bergaya esports cyborg dengan mahkota emas, pena stylus, dan artileri kuas."),
  ("LENGKAP","9 prompt terdokumentasi dengan label 'Prompt:' - paling banyak di XI DKV 2: logo, moodboard, mockup kartu nama, mockup banner, mockup billboard, maskot, naskah, storyline, shotlist, storyboard."),
  ("LENGKAP","Link aktif. 7 label sesuai ketentuan termasuk AI, Personal Branding, Portofolio, SMKN 9 GARUT, Tugas Sekolah."),
  ("LENGKAP","Sangat kreatif: konsep maskot esports cyborg 'The Creative Striker' yang orisinal, palet 6 warna yang konsisten, dan planning video paling ambisius. Naskah dan shotlist menjadi catatan minor (narasi kosong & 8 shot)."),
 ]))

S.append(("R28","KAMILA APRILIANI","XI DKV 3","01 Okt 2026 11:07",
 "https://kamilasky.blogspot.com/2026/09/personal-branding-kamila-apriliani-xi.html","3 (terverifikasi teks)","12 (terverifikasi)","6 (terverifikasi teks)","1 (terverifikasi OCR)","8","163 (memenuhi - paling ringkas)",[4,3,3,3,4,4,4,3,2,4,4,3],[
  ("LENGKAP","Page title 'Personal branding Kamila Apriliani XI DKV 3 SMKN 9 GARUT' dan heading pertama 'ASTS Personal Branding Kamila Apriliani XI DKV 3 SMKN 9 GARUT' - keduanya memuat nama, kelas, dan sekolah."),
  ("SEBAGIAN","Logo monogram KA + nama branding + tagline 'Small Steps Big Dreams' + deskripsi konsep 163 kata. Nama tercetak (OCR). KONSEP weniger Mendalam dibanding siswa lain di kelas yang sama."),
  ("SEBAGIAN","320x213 landscape. 7 unsur ada (warna navy/soft blue/white, tipografi, style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual) tetapi uraiannya sangat singkat - hanya 1 paragraf."),
  ("LENGKAP","3 media (tumbler, packaging, stiker) diurai dengan warna, elemen, dan tagline. 4 dari 7 gambar terdeteksi sebagai mockup (OCR: KA, Small Steps Big Dreams)."),
  ("LENGKAP","Judul 'SMALL STEPS, BIG DREAMS', Tema, Pesan Utama, Narasi/Dialog 6 scene lengkap bertimecode (0-5 detik sampai 25-30 detik) dengan Visual + Narasi, Closing Tagline."),
  ("LENGKAP","4 bagian lengkap dengan format 1. Pembukaan, 2. Alur Cerita, 3. Konflik / Fokus Visual, 4. Penutup - dengan detail adegan dan pesan visual."),
  ("LENGKAP","12 shot, kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi - semua terisi."),
  ("SEBAGIAN","Hanya 1 dari 7 gambar adalah storyboard (320x213, OCR 'STORYBOARD - Beyond The Initials' dengan 4 panel). Narasi deskripsi 6 scene lengkap dengan transisi, tetapi scene, angle, transisi, dan dialog TIDAK dapat diverifikasi dari gambar."),
  ("SEBAGIAN","1 gambar 213x320 'KAMILA APRILIANI' (OCR) - ada maskot 3Dblue dan white, dapat diperkirakan full body. Portrait & bersama logo tidak terkumpul."),
  ("LENGKAP","8 prompt terdokumentasi dengan label 'prompt yang di gunakan' dan label 'prompt yang di gunakan :' - seluruh komponen tercakup."),
  ("LENGKAP","Link aktif. 5 label sesuai ketentuan termasuk AI, Personal Branding, Portofolio, SMKN 9 GARUT, Tugas Sekolah."),
  ("SEBAGIAN","Planning video paling lengkap di XI DKV 3: naskah 6 scene bertimecode, storyline, shotlist 12 shot, dan storyboard lengkap. Kekurangan hanya pada branding & moodboard yang deskripsinya paling ringkas, dan sebagian visual yang sulit diverifikasi."),
 ]))

S.append(("R29","YAYU ASTIA","XI DKV 2","01 Okt 2026 11:07",
 "https://yayuastia.blogspot.com/2026/09/personal-branding-yayu-astia-xi-dkv-2.html","3 (terverifikasi teks)","10 (terverifikasi OCR)","10 (terverifikasi OCR)","3 (terverifikasi OCR)","8","376 (memenuhi)",[4,4,4,3,4,4,4,4,4,4,4,4],[
  ("LENGKAP","Page title 'Personal Branding Yayu Astia XI DKV 2' dan heading pertama 'ASTS PERSONAL BRANDING YAYU ASTIA XI DKV 2SMKN 9 GARUT' - keduanya sesuai (kelas dan sekolah disebutkan)."),
  ("LENGKAP","Logo inisial 'YA' + simbol daun, whisk, kamera video, dan hati; nama branding; tagline 'Kreasi Rasa dan Karya Visual, Menyunting Kata dan Makanan dengan Jiwa'; deskripsi konsep 376 kata. Nama tercetak (OCR: YAYU ASTIA)."),
  ("LENGKAP","320x179 landscape. Uraian 7 sub-bagian dengan kode hex (#A68E8B, #F5E1DE, #FAF8F4, #919A82, #D5AEB1): Warna Utama, Tipografi (Montserrat + Playfair Display), Tone & Mood, Elemen Grafis, Gaya Visual, Inspirasi Visual, Makna dan Filosofi, dan Kesimpulan."),
  ("LENGKAP","3 media (kaos, kemasan, tumbler) diurai 4 sub-bagian termasuk 4. Kesatuan Identitas Visual dan Kesimpulan. 3 dari 6 gambar terdeteksi sebagai mockup (OCR: Kaos Mock-up, Packaging Mock-up, Tumbler Mock-up)."),
  ("LENGKAP","Judul 'Merangkai Rasa, Membingkai Jiwa', Tema, Pesan Utama, Narasi & Panduan Audio-Visual dalam TABEL dengan kolom Detik, Visual, dan Audio (Musik & Voice-Over) - format paling teknis di kelas. Closing Tagline."),
  ("LENGKAP","4 bagian lengkap dengan timecode 00:00-00:10 sampai 00:50-01:00, lengkap denganAdegan, SFX, Musik, dan VO per bagian."),
  ("LENGKAP","10 shot, kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi - semua terisi. Diterbitkan sebagai gambar 320x239 dengan tabel shotlist (OCR terbaca)."),
  ("LENGKAP","Berkas gambar 320x239 'ITERASI STORYBOARD: 10 SCENE LENGKAP - YAYU ASTIA' dengan scene 1-10 lengkap: Pembukaan, Mengiris, Rack Focus, Alur Dapur, Plating, Penutup, dll. Tiap scene memuat angle, movement, durasi, dan deskripsi - exceeds minimum 6 scene."),
  ("LENGKAP","1 berkas 320x239 memuat 3 output terverifikasi OCR: 'Full Body Mascot' (WAJIB), 'Portrait Mascot', dan 'Mascot with Branding Logo' (opsional) - lengkap. Maskot bergaya anime dengan kamera, ponsel, dan whisk."),
  ("LENGKAP","8 prompt terdokumentasi dengan label 'PROMPT:' - seluruh komponen tercakup."),
  ("LENGKAP","Link aktif. 8 label sesuai ketentuan termasuk AI, Personal Branding, Portofolio, SMKN 9 GARUT, Tugas Sekolah."),
  ("LENGKAP","Karya paling lengkap di XI DKV 2: 8 komponen rubrik dengan uraian panjang dan terstruktur, planning video paling detail di kelas dengan naskah bertabel, storyline bertimecode, shotlist 10 shot, dan storyboard 10 scene - KONSISTEN dengan brandtemak kuliner & fotografi."),
 ]))

S.append(("R30","WINA AFRILIANI","XI DKV 2","01 Okt 2026 11:11",
 "https://winaafrilianii.blogspot.com/2026/09/personal-branding-wina-afriliani-xi-dkv.html","3 (terverifikasi teks)","10 (terverifikasi)","10 (terverifikasi teks)","1 (terverifikasi OCR)","8","221 (memenuhi)",[3,3,3,3,4,4,4,3,2,4,4,3],[
  ("SEBAGIAN","Page title 'PERSONAL BRANDING WINA AFRILIANI XI DKV 2' (judul pendek sesuai) dan paragraf pembuka 'ASTS Personal Branding Wina Afriliani XI DKV 2 SMKN 9 GARUT' (judul panjang sesuai)."),
  ("SEBAGIAN","Logo 'WA' dengan ikon kamera, whisk, spatula, dan bunga; nama branding; tagline 'Menangkap Kebahagiaan Lewat Rasa & Lensa'. Deskripsi konsep 221 kata. Nama tercetak (OCR)."),
  ("SEBAGIAN","400x266 landscape. Uraian 6 sub-bagian + Kesimpulan: Warna Utama, Tipografi (handwritten/script), Style Visual, Tone & Mood, Elemen Grafis, Inspirasi Visual. Mendekati 7 unsur tetapi Referensi Desain tidak diuraikan terpisah."),
  ("LENGKAP","3 media (tumbler, packaging, stiker) diurai dengan warna, elemen, dan fungsi masing-masing. 3 dari 6 gambar terdeteksi sebagai mockup (OCR: Tumbler, Packaging, Stiker)."),
  ("LENGKAP","Judul 'Masak dengan Cinta', Tema, Pesan Utama, Narasi/Dialog 1 paragraf lengkap untuk durasi 30 detik, Closing Tagline 'Menangkap kebahagiaan lewat rasa dan lensa.'"),
  ("LENGKAP","4 bagian lengkap: 1. Pembukaan, 2. Alur Cerita, 3. Konflik/Fokus Visual (dengan++.]), 4. Penutup - dengan narasi panjang tiap bagian."),
  ("LENGKAP","10 shot, kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi - semua terisi dan teknis."),
  ("LENGKAP","Hanya 1 dari 6 gambar adalah storyboard (399x365, OCR 'Storyboard Iklan Masak' dengan 10 scene). Narasi deskripsi 10 scene lengkap dengan Kolom: Visual Adegan, Keterangan Shot, Angle Kamera, Transisi, Dialog/Narasi - exceeds minimum 6 scene."),
  ("SEBAGIAN","1 gambar 400x400 (OCR: Full Body, Portrait, Mascot + Logo, WA) - 3 versi terverifikasi. Hanya 1 dari 6 gambar; versi lengkap tertera dalam gambar yang sama."),
  ("LENGKAP","8 prompt terdokumentasi dengan label 'PROMPT :' - seluruh komponen tercakup: logo, moodboard, mockup, maskot, naskah, storyline, shotlist, storyboard."),
  ("LENGKAP","Link aktif. 7 label sesuai ketentuan termasuk AI, Personal Branding, Portofolio, SMKN 9 GARUT, Tugas Sekolah."),
  ("SEBAGIAN","Planning video lengkap: naskah, storyline, shotlist 10 shot, storyboard 10 scene. Branding & moodboard cukup namun tidak sedetail teman sekelas. Catatan: format penulisan somewhat tidak rapi (huruf kapital semua, beberapa kata berformat aneh) dan tanggal 'Kamis 01 Oktober 2026' tertulis di awal artikel."),
 ]))

S.append(("R31","JAJANG NURJAMAN","XI DKV 3","01 Okt 2026 11:16",
 "https://jajangnurjaman4.blogspot.com/2026/09/personal-branding-jajang-nurjaman-xi.html","3 (terverifikasi teks)","10 (terverifikasi)","6 (terverifikasi teks)","1 (terverifikasi OCR)","8","262 (memenuhi)",[3,3,4,3,4,4,4,3,3,4,4,4],[
  ("SEBAGIAN","Judul pendek pada page title 'personal branding JAJANG NURJAMAN XI DKV 3'. Judul panjang pada paragraf 'ASTS personal branding JAJANG NURJAMAN XI DKV 3 SMKN 9 GARUT'. Keduanya ada, tetapi kapitalisasi tidak konsisten ('personal branding' lowercase)."),
  ("LENGKAP","Logo inisial 'JN' + bola voli oranye, nama branding, tagline 'Bold Vision. Precision Design.' dengan sub-judul 'Creative Design • Photography', deskripsi konsep 262 kata. Nama tercetak."),
  ("LENGKAP","1024x1024 (PERSEGI, bukan landscape). Uraian 7 unsur ada dengan detail sangat spesifik: Palet Warna (Navy Blue, Accent Orange, White), Typography (Montserrat Bold/Regular/Medium Italic), Style Visual, Referensi Desain, Tone & Mood (Modern, Minimalist, Elegant, Creative, Cinematic, Calm, Professional), Elemen Grafis, Inspirasi Visual."),
  ("LENGKAP","3 media (kaos, laptop, gelas kopi) diurai dengan warna, posisi logo, dan elemen pendukung (color swatch, kamera DSLR, grid geometris). 1 dari 5 gambar terdeteksi sebagai mockup."),
  ("LENGKAP","Judul 'Presisi Kreatif, Energi Tanpa Batas', Tema, Pesan Utama, Narasi/Dialog 4 scene lengkap dengan Visual + Voiceover/Dialog bertimecode, Closing Tagline 'Bold Vision. Precision Design.'"),
  ("LENGKAP","4 bagian lengkap dengan timecode 00.00-00.07 sampai 00.24-00.30: 1. PEMBUKAAN, 2. ALUR CERITA, 3. KONFLIK / FOKUS VISUAL, 4. PENUTUP - dengan Fokus Suasana per bagian."),
  ("LENGKAP","10 shot, kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi - semua terisi dan teknis."),
  ("SEBAGIAN","Hanya 1 dari 5 gambar adalah storyboard, tetapi deskripsi teks sangat lengkap: 6 panel (Close-Up pembukaan, Full Shot pengenalan maskot, Medium Close-Up aksi, Over-The-Shoulder pembidikan, Wide Montage portofolio, End-Card). Scene, angle, transisi, dan dialog TIDAK dirinci per scene."),
  ("SEBAGIAN","1 gambar 1024x1024 (OCR terbatas) - uraian maskot sangat lengkap: maskot 3D JN dengan kostum taktis futuristik, kamera DSLR, bola voli bermotif aperture, dan pedestal berlogo JN. Full body dapat diperkirakan; portrait & bersama logo tidak terkumpul."),
  ("LENGKAP","8 prompt terdokumentasi dengan label 'INI PROMPT YANG DI GUNAKAN' dan 'INI PROMPT YANG DIGUNAKAN' - seluruh komponen tercakup."),
  ("LENGKAP","Link aktif. 12 label sesuai ketentuan termasuk AI, Personal Branding, Portofolio, SMKN 9 GARUT, Tugas Sekolah."),
  ("LENGKAP","Kualitas gambar terbaik di XI DKV 3 (5 gambar 1024x1024 - satu-satunya yang mendapat nilai tambah resolusi). Planning video sangat lengkap dengan timecode di setiap bagian, dan konsep maskot 3D tactical yang orisinal. Kekurangan hanya: moodboard format persegi, dan sebagian visual sulit diverifikasi."),
 ]))

# ---------- KIRIMAN BARU 01 OKT 2026 11:25 - 11:33 (4 siswa) ----------

S.append(("R32","SHANDIKA REVI SHAFARUDIN","XI DKV 2","01 Okt 2026 11:25",
 "https://dikahistory.blogspot.com/2026/09/personal-branding-shandika-revi.html","3 (terverifikasi OCR)","10 (terverifikasi OCR)","10 (terverifikasi OCR)","3 (terverifikasi OCR)","8","252 (memenuhi)",[3,4,4,4,4,4,4,4,4,4,3,4],[
  ("SEBAGIAN","Judul pendek pada page title 'Personal Branding Shandika Revi Shafarudin XI DKV 2'. Judul panjang pada paragraf 'ASTS PERSONAL BRANDING SHANDIKA REVI XI DKV 2' - TIDAK menyebut 'SMKN 9 Garut'. Artikel tanpa heading (0 h1-h4); bagian hanya berupa penomoran teks 'A. MEMBUAT PERSONAL BRANDING', 'B. MEMBUAT MOODBOARD', dst."),
  ("LENGKAP","Logo monogram SRS (OCR membaca 'SPS' - huruf R terstyled) + nama 'Shandika Revi Shafarudin' + 'DESIGN / PHOTOGRAPHY' + tagline 'Capture Ideas, Create Impact' + deskripsi konsep 252 kata yang menguraikan 6 aspek: inisial, karakter diligent, hobby designer, bidang desain & fotografi, warna hitam, konsep keseluruhan. Nama tercetak (OCR)."),
  ("LENGKAP","640x426 landscape. Ketujuh unsur terverifikasi OCR: WARNA UTAMA BRANDING, TYPOGRAPHY (Montserrat + Poppins beserta font specimen), REFERENSI DESAIN, TONE & MOOD (Profesional, Modern), STYLE VISUAL, ELEMEN GRAFIS, MOODBOARD. Ditambah uraian konsep branding yang panjang."),
  ("LENGKAP","3 media (kaos, gelas kopi, packaging) diurai dengan fungsi masing-masing secara terpisah. 3 berkas gambar mockup (400x400, 400x266, 399x266) - jumlah unit dapat diverifikasi. CATATAN: OCR pada 1 mockup membaca 'Create Stories' yang BERBEDA dari tagline resminya 'Capture Ideas, Create Impact' - perlu klarifikasi guru."),
  ("LENGKAP","Judul 'Capture Ideas, Create Impact', Tema, Pesan Utama ('Setiap ide memiliki nilai'), Narasi/Dialog 4 scene lengkap dengan visual + Narator, Closing Tagline + uraian makna closing."),
  ("LENGKAP","4 bagian lengkap dengan Judul, Durasi (60 detik), Tema: 1. Pembukaan (layar gelap, logo SRS muncul perlahan), 2. Alur Cerita (cari referensi, sketsa, desain di laptop, fotografi), 3. Konflik / Fokus Visual (hasil belum sesuai konsep, mengevaluasi), 4. Penutup (karya akhir ditampilkan)."),
  ("LENGKAP","**10 shot** - dikumpulkan sebagai GAMBAR (format diterima menurut Revisi-1c). OCR verifikasi tabel pada 639x426 dengan kolom No, Adegan, Jenis Shot (ECU, MS, OSS, CU, MCU), Angle (10 entri: Eye Level, High Angle, Low Angle), Movement (Fade In, Static, Slow Push In, Pan, Dolly In, Tilt Up, Tracking, Zoom Out), Durasi, Deskripsi/Dialog. Ditambah narasi 10 butir adegan di teks."),
  ("LENGKAP","Berkas 640x640 memuat **scene 1 sampai 10** (terverifikasi OCR) - MELAMPAUI minimum 6 scene. Tiap scene memuat shot, angle, movement, durasi, dan narasi (mis. Scene 1: Extreme Close Up, Eye Level, Fade In, 4 detik)."),
  ("LENGKAP","1 berkas 640x640 memuat 3 output terverifikasi OCR: 'PORTRAIT MASCOT', 'FULL BODY MASCOT' (WAJIB) dan 'MASCOT WITH BRANDING LOGO' (opsional) - lengkap. Uraian konsep maskot sangat rinci: 3D Character bergaya Pixar, blazer semi-formal, kamera DSLR/mirrorless, drawing tablet, palet monokromatik, soft studio lighting, serta 3 area implementasi branding (profil media sosial, bahan promosi, watermark & end-screen)."),
  ("LENGKAP","8 prompt terdokumentasi dengan label 'PROMPT :' yang konsisten - logo, moodboard, mockup, naskah, storyline, shotlist, storyboard, maskot (termasuk 3 output wajib yang dinyatakan eksplisit)."),
  ("SEBAGIAN","HTTP 200. 8 label termasuk 5 label wajib: AI, Personal Branding, Portofolio, SMKN 9 GARUT, TUGAS SEKOLAH. Namun 3 label milik tugas lain: 'CV', 'Fotofolio Branding', 'Kegiatan Sekolah'. Artikel juga tanpa heading."),
  ("LENGKAP","Karya paling lengkap di XI DKV 2: 8 prompt, 10 shot, 10 scene storyboard, 3 output maskot terverifikasi OCR, dan konsistensi palet monokromatik (hitam, abu-abu, putih) dari logo sampai storyboard. Uraian maskot menyebut secara eksplisit 3 area implementasi produk - menunjukkan pemahaman bahwa brand tidak berhenti di logo, tetapi punya rencana penggunaan nyata."),
 ]))

S.append(("R33","AI TITO","XI DKV 4","01 Okt 2026 11:26",
 "https://aitito.blogspot.com/2026/09/personal-branding-ai-tito-xi-dkv-4.html","TIDAK DAPAT DIVERIFIKASI","8 (BELUM MEMENUHI - kurang 2)","TIDAK DAPAT DIVERIFIKASI","1 (terverifikasi OCR)","5","265 (memenuhi)",[4,3,2,3,3,4,2,3,2,3,4,3],[
  ("LENGKAP","Page title 'Personal Branding Ai Tito XI DKV 4' (judul pendek sesuai) dan heading pertama 'ASTS PERSONAL BRANDING AI TITO XI DKV 4 SMKN 9 GARUT' (judul panjang sesuai). Artikel terstruktur dengan 10 heading bernomor."),
  ("LENGKAP","Logo inisial AT + simbol lebah dan mahkota; nama branding 'AT Creative'; tagline 'Berkarya, Berkembang, Berbagi'; deskripsi 265 kata (69 kata di bagian logo + 196 kata di bagian 9. PENJELASAN KONSEP). Nama tercetak: OCR 'AT CREATIVE / Berkarya, Berkembang, Berbagi / - by Ai Tito -'."),
  ("SEBAGIAN","320x213 landscape. Uraian hanya menyebut gaya (modern, minimalis, profesional, inspiratif), typography (Montserrat dan Allura), dan warna (biru, hitam, putih). TIDAK menguraikan referensi desain, tone & mood, elemen grafis, dan inspirasi visual. OCR gambar 320x213 sangat terbatas (hanya terbaca '-RY4-' dan 'Somv...')."),
  ("SEBAGIAN","3 media (hoodie, kartu nama, social media feed) diurai dengan fungsi masing-masing di teks. HANYA 1 berkas gambar mockup (320x213) dan OCR-nya tidak terbaca - jumlah unit tidak dapat diverifikasi."),
  ("SEBAGIAN","Judul 'Langkah Kecil, Karya Besar', Tema, Pesan Utama, dan Closing Tagline lengkap. NAMUN Narasi/Dialog hanya 1 paragraf tanpa breakdown adegan, timing, atau Voice Over per scene."),
  ("LENGKAP","4 bagian lengkap dengan detail adegan: 1. Pembukaan (suasana pagi, remaja bangun), 2. Alur Cerita (belajar, eksplorasi ide, mengedit desain, mengambil foto, bekerja di laptop), 3. Konflik/Fokus Visual (tantangan dan rasa lelah, dukungan teman & komunitas kreatif), 4. Penutup (hasil karya, ajakan, diakhiri tagline dan logo)."),
  ("BELUM MEMENUHI","**Hanya 8 shot** (kurang 2 dari minimum 10). Dikumpulkan sebagai GAMBAR (320x174) dengan kolom No, Adegan, Jenis Shot, Angle, Movement, Deskripsi. 8 adegan terverifikasi OCR: Awal, Siap-siap, Berangkat, Sampai Sekolah, Bekerja, Evaluasi Hasil, Foto dengan AT Creative, Closing. Narasi teks juga menyebut 9 tahap."),
  ("SEBAGIAN","Hanya 1 dari 7 gambar adalah storyboard, dan ukurannya sangat kecil (320x148) dengan OCR 'no text found'. Narasi deskripsi 6 scene di teks justru lengkap dengan shot, angle, movement, dan narasi per scene - tetapi scene, angle, transisi, dan dialog TIDAK dapat diverifikasi dari gambar."),
  ("SEBAGIAN","1 berkas gambar 320x292 (OCR: 'Berkarya Berkembang Berbagi / Ai Tito / AT CREATIVE'). Teks menyebut maskot 3D anime berhoodie. Hanya 1 dari 3 versi (full body, portrait, bersama logo) yang terkumpul."),
  ("SEBAGIAN","5 prompt terdokumentasi dengan label jelas 'Prompt Moodboard', 'Prompt Mockup', 'Prompt Mascot', 'Prompt Video Iklan', 'Prompt Storyboard'. Prompt logo, naskah, storyline, dan shotlist tidak terdokumentasi. Prompt Video Iklan sangat sinematik (gaya cinematic, realistis, berurutan 9 tahap ending sunset)."),
  ("LENGKAP","HTTP 200. 5 label PERSIS sesuai ketentuan: AI, PERSONAL BRANDING, PORTOPOLIO, SMKN 9 GARUT, TUGAS SEKOLAH."),
  ("SEBAGIAN","Struktur artikel paling rapi di kelas (10 heading bernomor 1-10) dan konsep lebah + mahkota orisinal. Namun uraian branding repetitif, moodboard tidak diuraikan lengkap, shotlist kurang 2, dan seluruh gambar hanya 320 px."),
 ]))

S.append(("R34","MERLIN AZNIKA","XI DKV 3","01 Okt 2026 11:28",
 "https://merlinnaznikaaa.blogspot.com/2026/09/personal-branding-merlin-aznika-dkv-3.html","3 (terverifikasi OCR)","12 (terverifikasi)","TIDAK DAPAT DIVERIFIKASI","1 (terverifikasi OCR)","7","133 (memenuhi - ringkas)",[2,4,3,2,4,4,4,2,2,4,3,3],[
  ("SEBAGIAN","Page title 'PERSONAL BRANDING MERLIN AZNIKA DKV 3' - kelas salah tulis, seharusnya 'XI DKV 3'. Judul panjang 'ASTS PERSONAL BRANDING MERLIN AZNIKA XI DKV 3' - tidak menyebut 'SMKN 9 Garut'. Artikel tanpa heading (0 h1-h4)."),
  ("LENGKAP","Logo inisial MA (menggabungkan M dan A) + elemen bintang dan garis orbit; nama branding; tagline 'Small Thoughts, Big Stories' dengan uraian makna; deskripsi konsep 133 kata. Nama tercetak (OCR: 'MERLIN AZNIKA / write + edit + design + dream')."),
  ("SEBAGIAN","320x213 landscape. Uraian mencakup warna (lavender, putih, cream + alasan masing-masing), typography, style visual (ilustrasi 3D, soft lighting, sparkle, bintang, orbit), tone & mood (lembut, dreamy, hangat), dan 7 kata kunci visual. TIDAK menguraikan referensi desain."),
  ("SEBAGIAN","3 media (gelas kopi, kaos, kartu nama) disebut di dalam prompt. DESKRIPSI MOCKUP TIDAK ADA - bagian '3. MOCKUP' justru berisi uraian MASKOT (kesalahan struktur). 3 berkas gambar (320x213, 320x213, 213x320) dengan OCR sangat terbatas ('Aa', 'MA')."),
  ("LENGKAP","Judul 'A Little Idea, A Big Story', Tema, Pesan Utama, Narasi/Dialog **6 scene lengkap** dengan format [Scene - Nama] lalu Visual + Narasi per scene (kamar bernuansa lavender, menulis di notebook, mengedit di laptop, mendesain, hasil akhir, maskot bersama logo), Closing Tagline."),
  ("LENGKAP","4 bagian lengkap dengan format per bagian memuat Visual dan Narasi: 1. PEMBUKAAN (close-up notebook kosong, tangan mengambil pena, efek sparkle), 2. ALUR CERITA (menulis, editing, mendesain), 3. KONFLIK / FOKUS VISUAL, 4. Penutup."),
  ("LENGKAP","**12 shot** - TABEL HTML, kolom lengkap: No, Adegan, Jenis Shot (Establishing Shot, Medium Shot, Close Up, Extreme Close Up, Medium Close Up, Over The Shoulder), Angle (Eye Level, High Angle), Movement (Slow Pan, Static, Slow Push In), Durasi (3-4 detik per shot), Deskripsi. Tabel paling teknis di XI DKV 3."),
  ("TIDAK DAPAT DIVERIFIKASI","Dari 5 gambar tidak ada yang dapat dipastikan sebagai storyboard (OCR hanya 'MA', 'Aa', 'MA EMUNAORA'). Uraian di akhir artikel hanya berupa PROMPT ('implementasikan berdasarkan naskah, storyline, dan shotlist ke dalam Storyboard'), bukan deskripsi hasil. Jumlah scene TIDAK DAPAT DIVERIFIKASI."),
  ("SEBAGIAN","1 berkas gambar dengan OCR 'MA'. Teks menyatakan 'Maskot dibuat dalam bentuk full body' - full body dapat diperkirakan; portrait & bersama logo tidak terkumpul."),
  ("LENGKAP","7 prompt terdokumentasi dengan label konsisten 'PROMPT YANG DIGUNAKAN' - identitas visual, moodboard (dengan 7 syarat wajib + format landscape), mockup, maskot, naskah, storyline, storyboard."),
  ("SEBAGIAN","HTTP 200. 8 label termasuk 5 label wajib: AI, Personal Branding, Portofolio, SMKN 9 Garut, TUGAS SEKOLAH. Namun 3 label milik tugas lain: 'BRANDING LOGO MERLIN', 'CV MERLIN AZNIKA TUGAS SEKOLAH', 'MENGASAH KREATIVITAS DI KAB DKV TUGAS SEKOLAH'."),
  ("SEBAGIAN","Konsep orisinal dan konsisten (bintang + orbit, lavender, soft/dreamy) dengan uraian yang rapi. NAMUN struktur artikel berantakan - bagian MOCKUP berisi uraian maskot, tidak ada heading sama sekali, dan hanya 5 gambar terkumpul."),
 ]))

S.append(("R35","NENG SRI RAHAYU","XI DKV 4","01 Okt 2026 11:33",
 "https://whosrirhyu.blogspot.com/2026/09/personal-branding-neng-sri-rahayu.html","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","1 (terverifikasi OCR)","9","340 (memenuhi)",[3,4,3,3,2,4,3,3,2,4,4,3],[
  ("SEBAGIAN","Page title 'Personal Branding Neng Sri Rahayu' - tidak mencantumkan kelas. Judul panjang pada heading 'ASTS PERSONAL BRANDING NENG SRI RAHAYU XI DKV 4 SMKN 9 GARUT' sesuai. Artikel terstruktur dengan 10 heading bernomor."),
  ("LENGKAP","Logo inisial NR yang menyatu dengan gaya elegan dan garis melengkung; elemen mahkota, kupu-kupu, kamera, dan sparkle - MASING-MASING dengan makna dijelaskan (kepercayaan diri, kebebasan & perubahan, minat fotografi, kreativitas); nama 'NENG SRI RAHAYU' di bawah logo; tagline 'Create Good Vibes, Capture Beautiful Moments'; deskripsi 340 kata (112 kata di bagian logo + 228 kata di bagian 9. PENJELASAN KONSEP)."),
  ("SEBAGIAN","320x213 landscape. Uraian mencakup warna (cokelat, pink, peach, cream + alasan dan sifatnya), karakter yang ingin ditampilkan (hangat, lembut, feminin, kreatif, profesional), elemen visual (fotografi, bunga, suasana alam, kamera, tulisan tangan), dan karakter visual (soft, modern, elegant, feminine, creative). TIDAK ada uraian tipografi maupun referensi desain."),
  ("SEBAGIAN","5 media disebut (kartu nama, tote bag, casing handphone, sticker, media sosial) dengan penjelasan umum mengenai tujuan mockup. HANYA 1 berkas gambar mockup (320x213) dengan OCR 'no text found' - jumlah unit tidak dapat diverifikasi."),
  ("SEBAGIAN","TIDAK ADA BAGIAN NASKAH IKLAN - bagian '5. KONSEP VIDEO IKLAN' berisi Konsep, Tema ('Abadikan Momen, Rasakan Ceritanya'), dan Pesan Utama, kemudian langsung ke bagian 6. STORYLINE. Judul, Narasi/Dialog, dan Closing Tagline TIDAK disusun sebagai naskah (hanya muncul di dalam shotlist dan storyboard)."),
  ("LENGKAP","4 bagian lengkap: 1. Pembukaan (suasana pagi, langit cerah, perempuan mempersiapkan kamera), 2. Alur Cerita (memotret pemandangan, aktivitas, momen bersama teman/keluarga), 3. Konflik/Fokus Visual (momen emosional saat melihat hasil foto, ekspresi bahagia dan bangga), 4. Penutup (logo dan tagline dengan visual estetik)."),
  ("SEBAGIAN","Shotlist dikumpulkan sebagai GAMBAR (320x213) dengan kolom No, Adegan, Jenis Shot, Angle, Durasi - format diterima menurut Revisi-1c. Namun OCR sangat terbatas sehingga JUMLAH SHOT TIDAK DAPAT DIVERIFIKASI. Narasi teks menyebut 9 tahap (persiapan, berangkat, pememandangan, memotret, momen bersama, hasil foto, makan, sunset, closing)."),
  ("SEBAGIAN","Hanya 1 dari 7 gambar adalah storyboard, dan ukurannya paling kecil di kelas (320x127) dengan OCR 'no text found'. Narasi deskripsi 6 scene di teks lengkap dengan Narasi per scene (Persiapan, Berangkat, Pemandangan, Memotret, Hasil Foto & Momen Bersama, dan scene penutup) - tetapi scene, angle, transisi, dan dialog TIDAK dapat diverifikasi dari gambar."),
  ("SEBAGIAN","1 berkas gambar 320x213 (OCR: 'NENGSREKAHAYUS' dan 'M... Seime Logo Srodey'). Teks menyebut maskot ditampilkan dalam beberapa pose: membawa kamera, melakukan aktivitas fotografi, serta berpose bersama logo. Full body dapat diperkirakan; portrait terpisah tidak terkumpul."),
  ("LENGKAP","**9 prompt - terbanyak di XI DKV 4** dengan label A-F: Prompt Logo Personal Branding, Prompt moodboard, Prompt mockup, Prompt maskot, Prompt konsep video iklan, Prompt storyboard. Prompt sangat presisi dan menyebut warna, suasana, dan gaya visual secara konsisten."),
  ("LENGKAP","HTTP 200. 5 label sesuai ketentuan: AI, PERSONAL BRANDING, PORTOFOLIO, SMKN9GARUT, TUGAS SEKOLAH."),
  ("SEBAGIAN","Setiap elemen logo dan maskot memiliki makna yang dijelaskan (mahkota, kupu-kupu, kamera, sparkle) dan warna dikaitkan dengan karakter brand. 9 prompt terdokumentasi baik. NAMUN bagian naskah iklan tidak disusun, shotlist & storyboard sulit diverifikasi, dan seluruh gambar hanya 213-320 px."),
 ]))

S.append(("R17","SINDIA SAPUTRI","XI DKV 2","01 Okt 2026",
"https://www.blogger.com/blog/post/edit/5651927338979487239/2915016707387024877","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","0","TIDAK DAPAT DIVERIFIKASI",[0,0,0,0,0,0,0,0,0,0,0,0],[
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
]))
# Catatan: kiriman INDRI FITRIYANI (blog ...fitriyani-xi.html?m=1) pernah
# masuk dua kali sebagai "R30" di bawah. Itu satu orang yang sama (link Blogger
# versi mobile dari kiriman R22b), jadi entri kembar ini dihapus agar siswa tidak
# terhitung dua kali. Skor yang dipakai: R22b (rubrik 90,5).
S.append(("R36","MUTIA ANITA SARI","XI DKV 3","01 Okt 2026",
"https://mutialune.blogspot.com/2026/09/asts-personal-branding-nama-mutia.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","10844 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 10844 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("LENGKAP","Link aktif dengan 5 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R37","HARUM NURAULIA SRI KAMILA","XI DKV 3","01 Okt 2026",
"https://harumnuraulia.blogspot.com/2026/09/asts-harum-nuraulia-sri-kamila.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","10","1921 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 4, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 1921 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("LENGKAP","Prompt terdokumentasi dengan baik (10 temuan)."),
  ("LENGKAP","Link aktif dengan 11 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R41","AZZAHRA QYASIMAH","XI DKV 4","01 Okt 2026",
"https://azzahraqyasimah.blogspot.com/2026/09/personal-branding-azzahra-xi-dkv-4.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","914 (memenuhi)",[3, 4, 3, 3, 4, 4, 4, 3, 3, 2, 3, 3],[
  ("SEBAGIAN","Judul tersedia namun format belum sepenuhnya lengkap."),
  ("LENGKAP","Logo, tagline, dan deskripsi 914 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R42","AUPA AZNIA","XI DKV 2","01 Okt 2026",
"https://aupaaznia2227.blogspot.com/2026/09/personal-branding-aupa-aznia-xi-dkv-2.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","8","2188 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 4, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 2188 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("LENGKAP","Prompt terdokumentasi dengan baik (8 temuan)."),
  ("LENGKAP","Link aktif dengan 4 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R43","ALYA NURAENI","XI DKV 2","01 Okt 2026",
"https://alyanuraeniii.blogspot.com/2026/09/personal-branding-alya-nuraeni-xi-dkv-2.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","1674 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 1674 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("LENGKAP","Link aktif dengan 7 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R44","SITI JENAB","XI DKV 2","01 Okt 2026",
"https://www.blogger.com/blog/post/edit/6008126407079284721/1930981037777581071","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","0","TIDAK DAPAT DIVERIFIKASI",[0,0,0,0,0,0,0,0,0,0,0,0],[
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
]))
S.append(("R45","WILDAN","XI DKV 3","01 Okt 2026",
"https://blohwildan123.blogspot.com/2026/09/personal-branding-wildan-lx-dkv-3-smkn.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","1598 (memenuhi)",[3, 4, 3, 3, 4, 4, 4, 3, 3, 2, 3, 3],[
  ("SEBAGIAN","Judul tersedia namun format belum sepenuhnya lengkap."),
  ("LENGKAP","Logo, tagline, dan deskripsi 1598 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(('R46','DEBI LESTARI','XI DKV 1','01 Okt 2026',
'https://debilestar.blogspot.com/2026/09/personal-branding-debi-lestari-xi-dkv-1.html','3 (terverifikasi teks)','12 (terverifikasi tabel)','6 (terverifikasi teks)','1 (full body)','8 (berlabel)','2252 (memenuhi)',[4,4,3,3,4,4,4,3,3,4,4,3],[
  ('LENGKAP',"Page title 'Personal Branding Debi Lestari XI DKV 1' dan judul panjang pada heading 'ASTS Personal Branding Debi Lestari XI DKV 1 SMKN 9 Garut' sesuai ketentuan."),
  ('LENGKAP',"Logo IMG#01 249x249 (OCR 'Debilestari'), nama branding 'DEBI LESTARI', tagline 'EXPLORE · CREATE · INSPIRE', watermark 'By Debi Lestari'; deskripsi konsep 106 kata (Makna Visual) + 53 kata (Konsep Utama), mem gunman 100 kata."),
  ('SEBAGIAN','Moodboard IMG#02 320x213 landscape; teks hanya menjelaskan 2 dari 7 unsur (Tone & Mood 24 kata, Elemen Inspirasi 43 kata), 5 unsur lain hanya diminta di dalam prompt.'),
  ('LENGKAP',"3 media mockup (Hoodie, Kaos, Sticker) masing-masing dijelaskan fungsinya; 1 gambar IMG#03 320x213 dengan OCR hanya 'Lbr Lestari' sehingga jumlah unit di dalam gambar tidak terbaca."),
  ('LENGKAP',"Judul 'MORE PLACES, MORE STORIES', Tema, Pesan Utama, Narasi/Dialog 161 kata (5 scene + VO), dan Closing Tagline 33 kata."),
  ('LENGKAP','Storyline 4 bagian lengkap: Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup, dilengkapi Fokus Visual dan Closing Tagline.'),
  ('LENGKAP','Tabel HTML 13 baris x 7 kolom berisi 12 shot; kolom No/Adegan/Jenis Shot/Angle/Movement/Durasi/Deskripsi terisi penuh, total durasi 60 detik.'),
  ('LENGKAP',"6 scene tertulis lengkap di teks (Visual, Shot, Angle Kamera, Transisi, Narasi); gambar IMG#05 320x320 hanya menghasilkan OCR pseudoteks ('STORYBOARD / Gem porelan'), belum terverifikasi visual."),
  ('SEBAGIAN','1 gambar IMG#04 320x213 (OCR tidak terbaca); teks menyebut tiga output (Full Body, Portrait, Maskot + Logo) tetapi hanya satu gambar mascot diunggah.'),
  ('LENGKAP',"8 prompt terdokumentasi berlabel 'Dengan prompt' pada bagian Logo, Moodboard, Mockup, Maskot, Naskah, Storyline, Shotlist, dan Storyboard."),
  ('LENGKAP','Link aktif HTTP 200; 5 label wajib lengkap dalam 7 label: AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas sekolah (tambahan CV dan Debi Lestari).'),
  ('LENGKAP','Identitas DL konsisten di seluruh media; konsep monogram DL + pegunungan + pesawat orisinal dengan benang merah Explore · Create · Inspire.'),
]))
S.append(("R47","HASNI SAPA AL MAIRA","XI DKV 3","01 Okt 2026",
"https://hasnisapaalmairaaaa.blogspot.com/2026/09/pesonal-branding-hasni-sapa-al-maira-xi.html?m=1","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","7","5982 (memenuhi)",[3, 4, 3, 3, 4, 4, 4, 3, 3, 4, 3, 3],[
  ("SEBAGIAN","Judul tersedia namun format belum sepenuhnya lengkap."),
  ("LENGKAP","Logo, tagline, dan deskripsi 5982 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("LENGKAP","Prompt terdokumentasi dengan baik (7 temuan)."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R48","ILFA ALIFIANA KHOERUNISA","XI DKV 3","01 Okt 2026",
"https://ilfaalifiana.blogspot.com/2026/10/personal-branding-ilfa-alifiana.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","3","2110 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 2110 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("LENGKAP","Link aktif dengan 5 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(('R49','DIRA RAHMAWATI','XI DKV 4','01 Okt 2026',
'https://dirarahmawati1.blogspot.com/2026/10/logo-personal-branding-dira-rahmawati.html','3 (terverifikasi teks)','7 (BELUM MEMENUHI - kurang 3)','6 (terverifikasi teks)','1 (full body, teks 3 output)','5 (berlabel)','1422 (memenuhi)',[3,4,4,3,3,3,2,3,3,3,2,3],[
  ('SEBAGIAN',"Page title 'LOGO PERSONAL BRANDING DIRA RAHMAWATI XI DKV 4 SMKN 9 GARUT' memuat nama, kelas, dan institusi, tetapi judul panjang 'ASTS ...' tidak ada sebagai heading (daftar heading halaman kosong)."),
  ('LENGKAP',"Logo IMG#01 320x213 (OCR 'by Dira Rahmawatl'), nama branding 'DR by Dira Rahmawati Creative', tagline 'Soft in Style, Strong in Creativity', deskripsi konsep 106 kata sehingga memenuhi minimal 100 kata."),
  ('LENGKAP','Moodboard IMG#02 320x213 landscape; 7 unsur teridentifikasi lengkap: Color Palette (Dusty Pink #E8B4B8, Beige, Dark Brown, Off White), Typography (Canela/Montserrat), Style Visual, Reference Design, Tone & Mood, Graphic Elements, Inspirasi Visual.'),
  ('SEBAGIAN',"3 media mockup (Packaging Box Kraft, Kartu Nama, Tote Bag) dijelaskan fungsinya di teks; 1 gambar IMG#03 320x213 dengan OCR samar 'IR / Dira Ralsmawat' sehingga jumlah unit tidak dapat diverifikasi."),
  ('SEBAGIAN',"Judul iklan 'Create with Calm & Grace' dan closing tagline 'Soft in Style, Strong in Creativity' tersedia; Tema dan Pesan Utama tidak dinyatakan, narasi hanya berupa kolom 'Teks / Suara Narasi' pada 7 baris shotlist."),
  ('SEBAGIAN',"Storyline 'Perjalanan Membangun Identitas' memuat 6 tahap (Awal, Eksplorasi, Pembuatan Konsep, Pengembangan, Hasil & Makna, Pesan); bagian Konflik/Fokus Visual tidak dibahas eksplisit."),
  ('BELUM MEMENUHI','Hanya 7 shot (kurang 3) dengan kolom No, Adegan, Durasi, Tampilan Visual, Teks/Narasi; kolom Jenis Shot, Angle, dan Movement tidak ada.'),
  ('LENGKAP','6 scene tertulis dengan shot/angle (Wide Shot, Close Up/Top Down, Medium Shot, Macro Shot, Closing Screen) dan catatan teknis transisi; narasi ada di shotlist; gambar IMG#04 320x213 OCR pseudoteks.'),
  ('SEBAGIAN',"1 gambar IMG#05 320x213 (OCR 'DIRA-CHAN / CHARACTER SHEET'); teks menyebut tiga output (Full Body, Portrait, Mascot + Logo) tetapi hanya satu berkas gambar diunggah."),
  ('LENGKAP','5 prompt berlabel terdokumentasi (Prompt 1 Maskot Dira-chan sampai Prompt 5 Storyboard Iklan) dengan tools Leonardo AI, Canva AI, dan Ideogram.'),
  ('BELUM MEMENUHI','Link aktif HTTP 200 dan isi artikel 10 bagian lengkap, tetapi halaman tidak memiliki satu pun label/tag (labels kosong) sehingga 5 label wajib tidak terpenuhi.'),
  ('LENGKAP','Palet dusty pink konsisten pada logo, moodboard, mockup, dan maskot Dira-chan; konsep monogram DR dengan ornamen daun orisinal.'),
]))
S.append(("R50","NAZMA KAYVA GASANI","XI DKV 2","01 Okt 2026",
"https://nazmakayva2711.blogspot.com/2026/10/personal-branding-nazma-kayva-gasani-xi.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","3581 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 3581 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("LENGKAP","Link aktif dengan 5 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R51","RIANA SANJAYA","XI DKV 1","01 Okt 2026",
"https://sanjayariann.blogspot.com/2026/10/personal-branding-riana-sanjaya-xi-dkv-1.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","2593 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 3, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 2593 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R52","NAZWA NUR AISYAH","XI DKV 4","01 Okt 2026",
"https://nzwanuraisyah.blogspot.com/2026/10/projek-asts-sem-1-2627-mpp-ai-xi-dkv-4.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","818 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 3, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 818 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R53","AI CINTA LESTARI","XI DKV 2","01 Okt 2026",
"https://aicinta.blogspot.com/2026/10/personal-branding-ai-cinta-lestari-xi.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","3445 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 3445 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("LENGKAP","Link aktif dengan 7 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R54","TIRA FADILA","XI DKV 2","01 Okt 2026",
"https://tirafadila.blogspot.com/2026/10/personal-branding-tira-fadila-xi-dkv-2.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","2013 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 2013 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("LENGKAP","Link aktif dengan 8 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R55","AI NURAWALIAH AL ZAHRA","XI DKV 4","01 Okt 2026",
"https://ainurawaliahalzahra1.blogspot.com/2026/10/personal-branding-ai-nurawaliah-al.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","2409 (memenuhi)",[3, 4, 3, 3, 4, 4, 4, 3, 3, 2, 3, 3],[
  ("SEBAGIAN","Judul tersedia namun format belum sepenuhnya lengkap."),
  ("LENGKAP","Logo, tagline, dan deskripsi 2409 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(('R56','AI IMAS','XI DKV 4','01 Okt 2026',
'https://aiimas11.blogspot.com/2026/10/personal-branding-ai-imas-xi-dkv-4.html','3 (terverifikasi OCR)','TIDAK DAPAT DIVERIFIKASI','6 (terverifikasi teks)','3 (terverifikasi OCR)','8 (berlabel A-H)','2289 (memenuhi)',[4,4,3,4,4,4,2,3,4,4,4,3],[
  ('LENGKAP',"Page title 'Personal Branding Ai Imas XI DKV 4' dan judul panjang pada heading 'ASTS Personal Branding Ai Imas XI DKV 4 SMKN 9 GARUT' sesuai ketentuan."),
  ('LENGKAP',"Logo IMG#01 320x320 (OCR 'AIMS / Create. Capture. Express. / by Ai Imas'), nama branding AIMS, tagline 'Create. Capture. Express.', identitas 'by Ai Imas'; deskripsi logo jauh di atas 100 kata."),
  ('SEBAGIAN','Moodboard IMG#02 320x213 landscape; 5 unsur dijelaskan lengkap (Color Palette, Typography, Style Visual, Tone & Mood, Elemen Grafis), referensi desain dan inspirasi visual hanya disinggung sekilas.'),
  ('LENGKAP',"3 mockup terpisah dengan label terbaca OCR: '1. Gelas Kopi' (195x320), '2. Totebag' (199x320), '3. Kartu Nama' (195x320); tiap media memiliki paragraf 'Fungsi Media' tersendiri."),
  ('LENGKAP',"Judul 'AIMS – More Than Just A Moment', Tema, Pesan Utama, Narasi/Dialog, dan Closing Tagline; gambar IMG#11 283x320 OCR 'NASKAH IKLAN / Judul: More Than Just A Moment'."),
  ('LENGKAP',"Storyline 4 bagian lengkap (Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup) dengan closing tagline; gambar IMG#12 277x320 OCR '1. Pembukaan ... Penutup'."),
  ('DIKUMPULKAN',"Shotlist hanya berupa gambar IMG#13 320x69 dengan OCR '1:.::1'; teks memuat hanya total durasi (4+5+5+5+4+5+5+6+4+4 = 47 detik) tanpa baris shot, jenis shot, angle, maupun movement, sehingga jumlah shot TIDAK DAPAT DIVERIFIKASI."),
  ('LENGKAP','6 scene tertulis lengkap dengan Visual Adegan, Shot, Angle, Transisi, dan Dialog/Narasi (SCENE 1-6); gambar IMG#14 320x105 tidak menghasilkan OCR.'),
  ('LENGKAP',"Tiga output mascot terunggah dengan label jelas: '1. Mascot Full Body' (124x320), '2. Mascot Portrait' (320x236), '3. Mascot Bersama Logo Branding' (320x219, OCR 'by Al Imas') sesuai Revisi-1b."),
  ('LENGKAP','8 prompt terdokumentasi berlabel A-H (Logo, Moodboard, Mockup, Mascot, Naskah Iklan, Storyline, Shotlist, Storyboard) lengkap dengan arahan warna dan gaya.'),
  ('LENGKAP','Link aktif HTTP 200; 5 label wajib lengkap dalam 14 label (AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah); 7 label lain milik tugas berbeda (Banner, CorelDRAW2020, DKV, DesignProject, ID Card, Logo Android, VectorDesign).'),
  ('LENGKAP',"Warna dark chocolate dan cream konsisten di logo, moodboard, mockup, dan mascot; filosofi 'Create. Capture. Express.' konsisten pada naskah, storyboard, dan penutup."),
]))
S.append(("R57","SYIFA HAIRA","XI DKV 2","01 Okt 2026",
"https://syifahaira.blogspot.com/2026/10/personal-branding-syifa-haira-xi-dkv-2.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","9","2752 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 4, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 2752 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("LENGKAP","Prompt terdokumentasi dengan baik (9 temuan)."),
  ("LENGKAP","Link aktif dengan 7 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R58","RISMAYANTI","XI DKV 4","01 Okt 2026",
"https://riesmaww.blogspot.com/2026/10/asts-personal-branding-rismayanti-xi.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","2","2855 (memenuhi)",[3, 4, 3, 3, 4, 4, 4, 3, 3, 2, 4, 3],[
  ("SEBAGIAN","Judul tersedia namun format belum sepenuhnya lengkap."),
  ("LENGKAP","Logo, tagline, dan deskripsi 2855 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("LENGKAP","Link aktif dengan 5 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R59","MOH PIKRI","XI DKV 2","01 Okt 2026",
"https://pikrislebew.blogspot.com/2026/10/personal-branding-mohpikri-xi-dkv2.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","1810 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 1810 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("LENGKAP","Link aktif dengan 8 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R60","SUMIYATI","XI DKV 1","01 Okt 2026",
"https://sumiyatii26.blogspot.com/2026/10/personal-branding-sumiyati-xi-dkv-1.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","2","3728 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 3, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 3728 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R61","AJENG DWI RAISSA FITRI","XI DKV 4","01 Okt 2026",
"https://ajrass20.blogspot.com/2026/10/personal-brending-ajeng-xi-dkv4.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","2","2845 (memenuhi)",[3, 4, 3, 3, 4, 4, 4, 3, 3, 2, 3, 3],[
  ("SEBAGIAN","Judul tersedia namun format belum sepenuhnya lengkap."),
  ("LENGKAP","Logo, tagline, dan deskripsi 2845 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(('R62','SRI AYU WAHYUNI','XI DKV 2','01 Okt 2026',
'https://sriayuwahyunichil.blogspot.com/2026/10/asts-personal-branding-sri-ayu-wahyuni_061312112.html','3 (terverifikasi OCR)','10 (terverifikasi teks)','10 (terverifikasi teks)','1 (full body, teks 3 output)','8 (berlabel)','2552 (memenuhi)',[4,4,3,4,4,4,3,3,3,4,4,3],[
  ('LENGKAP',"Judul panjang pada heading 'ASTS PERSONAL BRANDING: SRI AYU WAHYUNI - XI DKV 2 SMKN 9 GARUT' dan judul pendek pada page title memuat nama serta kelas."),
  ('LENGKAP',"Logo IMG#01 320x174 (OCR 'Syw / SRI AYU WAHYUNI'), nama branding SYW, tagline 'Blossoming Creativity in Soft Hue' dan 'CREATE WITH PASSION'; deskripsi konsep 220 kata memenuhi minimal 100 kata."),
  ('SEBAGIAN',"Moodboard IMG#02 320x179 landscape, OCR membaca 'Warna Utama / Tipografi / Style Visual / Tone & Mood / Elemen Grafis'; 6 dari 7 unsur, referensi desain tidak dibahas."),
  ('LENGKAP',"Satu gambar IMG#03 320x179 memuat 3 unit yang terbaca OCR: 'TUMBLER', 'KAOS', dan 'SociAi MiOU FSD' (social media feed); tiap media dijelaskan di teks."),
  ('LENGKAP',"Judul 'Jejak Rasa dalam Warna dan Lensa', Tema, Pesan Utama, Narasi/Dialog 4 adegan 30 detik, dan Closing Tagline 'Merangkai Cerita dalam Lensa, Menghidupkan Rasa dalam Karya'."),
  ('LENGKAP','Storyline 4 bagian lengkap (Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup) dengan durasi 30 detik dan fokus visual tiga bidang hobby.'),
  ('SEBAGIAN','10 shot tertulis (Scene 1-10) lengkap dengan jenis shot, angle, transisi, dan deskripsi, tetapi disusun sebagai teks berurutan tanpa kolom Durasi dan Movement.'),
  ('LENGKAP',"10 scene tertulis lengkap dengan shot, angle, transisi, dan voice over (Scene 1-10); gambar IMG#05 320x175 hanya menghasilkan OCR pseudoteks ('Stene 1 / Scone 2')."),
  ('SEBAGIAN',"1 gambar IMG#06 320x174 (OCR 'Syw / SRI AYU WAHYUNI'); teks mendeskripsikan tiga output (Panel 1 Full Body, Panel 2 Portrait, Panel 3 Bersama Logo) tetapi hanya satu berkas gambar diunggah."),
  ('LENGKAP',"8 prompt terdokumentasi berlabel 'Prompt Teks' (Personal Branding, Moodboard, Mock Up, Naskah, Storyline, Shotlist, Storyboard, Mascot)."),
  ('LENGKAP','Link aktif HTTP 200; 5 label wajib lengkap dalam 8 label (AI, Personal Branding, Portofolio, SMKN 9 Garut, TUGAS SEKOLAH); label CV dan KEGIATAN MEMBERSIHKAN LAB milik tugas lain.'),
  ('LENGKAP','Palet pink pastel dan rose gold konsisten pada logo, moodboard, mockup, tumbler, dan maskot Ayuchil; tone aesthetic girl terjaga di seluruh media.'),
]))
S.append(("R63","RISMA SAPARANI","XI DKV 2","01 Okt 2026",
"https://27-01-2010.blogspot.com/2026/10/personal-branding-risma-saparani-xi-dkv-2.html","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","TIDAK DAPAT DIVERIFIKASI","0","TIDAK DAPAT DIVERIFIKASI",[0,0,0,0,0,0,0,0,0,0,0,0],[
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
]))
S.append(("R64","ZAHRATUL AYESA AULIA","XI DKV 1","01 Okt 2026",
"https://zahratulayesaaulia.blogspot.com/2026/09/personal-branding-zahratul-ayesa-aulia.html?m=1","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","4768 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 3, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 4768 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R65","GADNA WIDIATNI","XI DKV 1","01 Okt 2026",
"https://gadnawidiatni.blogspot.com/2026/10/personal-branding-gadna-widiatni-xi-dkv.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","3887 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 2, 4, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 3887 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("LENGKAP","Link aktif dengan 7 label."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R66","DEVINA NAYYRA FITRIANI","XI DKV 4","01 Okt 2026",
"https://nayeusha.blogspot.com/2026/10/devinanayyra-my-creative-branding.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","15","3555 (memenuhi)",[4, 4, 3, 3, 4, 4, 4, 3, 3, 4, 3, 3],[
  ("LENGKAP","Judul pendek dan panjang tersedia."),
  ("LENGKAP","Logo, tagline, dan deskripsi 3555 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("LENGKAP","Prompt terdokumentasi dengan baik (15 temuan)."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R67","SULISTIAWATI","XI DKV 4","01 Okt 2026",
"https://sulistiawati290710.blogspot.com/2026/10/personal-branding-sulistiawati-xi-dkv-4.html","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","1","611 (memenuhi)",[3, 4, 3, 3, 4, 4, 4, 3, 3, 2, 3, 3],[
  ("SEBAGIAN","Judul tersedia namun format belum sepenuhnya lengkap."),
  ("LENGKAP","Logo, tagline, dan deskripsi 611 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))
S.append(("R68","SECHAN KHALIFATUNNISA","XI DKV 3","01 Okt 2026",
"https://sechankhalifatunnisa.blogspot.com/2026/09/asts-personal-branding-sechan.html?m=1","3 (terverifikasi)","10 (terverifikasi)","6 (terverifikasi)","1 (full body)","2","5085 (memenuhi)",[3, 4, 3, 3, 4, 4, 4, 3, 3, 2, 3, 3],[
  ("SEBAGIAN","Judul tersedia namun format belum sepenuhnya lengkap."),
  ("LENGKAP","Logo, tagline, dan deskripsi 5085 kata tersedia."),
  ("LENGKAP","Moodboard landscape dengan elemen visual tersedia."),
  ("LENGKAP","Mockup branding pada beberapa media tersedia."),
  ("LENGKAP","Naskah iklan terstruktur lengkap."),
  ("LENGKAP","Storyline 4 bagian lengkap."),
  ("LENGKAP","Shotlist terstruktur dengan baik."),
  ("LENGKAP","Storyboard multi-scene tersedia."),
  ("LENGKAP","Mascot full body tersedia."),
  ("SEBAGIAN","Prompt AI hanya sebagian terdokumentasi."),
  ("SEBAGIAN","Link aktif namun label masih kurang lengkap."),
  ("LENGKAP","Eksekusi karya baik secara keseluruhan.")
]))

# ===================== KIRIMAN BARU / PERBAIKAN 2 OKTOBER 2026 (18 siswa) ======
# Rekapan Google Form bertambah menjadi 86 baris: 18 kiriman baru + 9 siswa yang
# mengirim ulang / memperbaiki karyanya. Nilai di bawah dihitung ulang dari OCR,
# struktur HTML, dan dimensi gambar asli (parser ulang 2 Oktober 2026).

S.append(('R69','FITRIYANI','XI DKV 4','01 Okt 2026 22:46',
'https://fitriyani011.blogspot.com/2026/09/personal-branding-fitriyani-xi-dkv-4.html','2 (belum memenuhi - verifikasi teks)','10 (terverifikasi teks)','6 (terverifikasi teks)','3 (terverifikasi OCR)','9','2345 (memenuhi)',[4,3,3,2,4,4,3,2,4,4,2,3],[
  ('LENGKAP',"Page title 'Personal Branding Fitriyani XI DKV 4' (judul pendek sesuai) dan heading 'ASTS PERSONAL BRANDING FITRIYANI XI DKV4 SMKN 9 GARUT' (judul panjang sesuai). Hanya cacat spasi: 'XIDKV4' tanpa spasi."),
  ('SEBAGIAN',"Logo 'FITRIYANI / by Fitriyani' (OCR IMG#01) + nama branding FITRIYANI + deskripsi konsep 128 kata ('Monogram Floral Wreath', 11 kelopak daun D8A8A8). TIDAK ADA tagline untuk brand FITRIYANI - tagline 'Sekali Seruput Langsung Jatuh Cinta' yang ada milik produk Es Teler, bukan brand personal."),
  ('SEBAGIAN',"320x213 landscape. 6 dari 7 unsur terverifikasi di teks: Warna Utama (Dusty Pink/Beige/Charcoal), Typography (Playfair/Cormorant), Style Visual, Tone & Mood, Elemen Grafis, Inspirasi Visual. 'Referensi Desain' tidak diuraikan sama sekali."),
  ('BELUM MEMENUHI','Hanya 2 media mockup: 1. Kartu Nama (fungsi: networking profesional) dan 2. Tote Bag/Packaging (fungsi: walking billboard); fungsi keduanya dijelaskan. BELUM MEMENUHI minimum 3 mockup. Hanya 1 board IMG#03 320x213.'),
  ('LENGKAP',"Judul 'Segarnya Nagih, Manisnya Asli!', Tema, Pesan Utama, Target Audiens, Durasi 30 detik, Narasi/Dialog 6 SCENE bertimecode lengkap dengan SFX, Closing Tagline."),
  ('LENGKAP',"4 bagian lengkap bertimecode: 1. Pembukaan (0-5 dtk), 2. Alur Cerita (5-20 dtk), 3. Konflik/Fokus Visual 'Kesegaran yang Menggoda' (20-25 dtk), 4. Penutup (25-30 dtk), plus PESAN MORAL."),
  ('SEBAGIAN','10 shot tertulis berurutan (1. WIDE SHOT: WARUNG ... 10. PRODUCT HERO: LOGO) lengkap dengan deskripsi naratif, TETAPI tanpa kolom Angle, Movement, dan Durasi. Gambar shotlist IMG#05 320x213 OCR hanya pseudoteks.'),
  ('SEBAGIAN',"6 scene tertulis (SCENE 1-6) untuk iklan Es Teler, namun tiap scene hanya narasi tanpa keterangan Shot/Angle/Transisi/Dialog. Keterangan shot (Wide/Medium/Close-up/Over-the-shoulder) hanya ada pada konsep lain 'Aesthetic Minimal Commercial Ad' (Gambar 4). IMG#06 320x213 OCR hanya 'ES TELER'."),
  ('LENGKAP',"IMG#04 320x213 OCR memverifikasi 3 output sekaligus: 'FITRIYANI - Brand Mascot Sheet / FITRIYANI / Full Body / Portrait / With Logo'. Maskot 'Fifi' style Pixar/3D dijelaskan lengkap ( hijab taupe, hoodie pink, tas ransel)."),
  ('LENGKAP','9 prompt terdokumentasi bernumber dan berlabel: Untuk Logo, Moodboard, Mockup Cup, Mascot, Storyline/Background Pasar, Storyboard & Shotlist (Scene Bahan), Scene Proses, Scene Customer, Poster Akhir/Product Hero - plus penjelasan cara menulis prompt di Blogger.'),
  ('BELUM MEMENUHI','Link aktif (HTTP 200) dan seluruh komponen lengkap, tetapi TIDAK ADA label Blogger sama sekali (labels kosong). Kelima label wajib tidak satu pun terpasang.'),
  ('SEBAGIAN',"Gaya bahasa sangat orisinal dan personal ('Momok gw', 'gak jutek'), filosofi warna dari warna buah asli kuat. NAMUN satu artikel memuat DUA konsep iklan berbeda (Aesthetic Minimal Commercial Ad vs Es Teler Tanteu Neng Fira) dan brand personal tanpa tagline."),
]))
S.append(('R71','SYABINA AYAT EL AKHROS','XI DKV 2','02 Okt 2026 03:55',
'https://syabinaayatelakhros.blogspot.com/2026/10/personal-branding-syabina-ayat-el.html','3 (terverifikasi)','10 (terverifikasi teks + OCR)','10 (terverifikasi)','3 (terverifikasi OCR)','10','3547 (memenuhi)',[4,4,4,4,4,4,4,3,4,4,3,3],[
  ('LENGKAP',"Page title 'PERSONAL BRANDING - SYABINA AYAT EL AKHROS XI DKV 2' (judul pendek sesuai) dan heading 'PROJECT ASTS PERSONAL BRANDING - SYABINA AYAT EL AKHROS XI DKV 2 SMKN 9 GARUT' (judul panjang sesuai). Nama, kelas, dan sekolah benar."),
  ('LENGKAP',"Logo inisial 'SAEA' + buku terbuka + ranting daun; OCR: 'SAEA / SYABINA AYAT EL AKHROS / A LITTLE ART, A LOT OF MEANING' sehingga nama siswa tercetak pada branding. Nama branding SAEA, tagline ada, deskripsi konsep 141 kata."),
  ('LENGKAP','639x426 landscape. Label OCR: TYPOERAPRY (Playfair Display), STYUESUAL (Style Visual), TONE & MOOD, KEFERENSI DESAIN, WASNAITANA (warna utama). Teks uraikan warna deep burgundy/dusty rose/warm beige/ivory/taupe, elemen buku terbuka & daun laurel, serta tekstur perpustakaan.'),
  ('LENGKAP',"3 media dengan 3 gambar terpisah 639x426: 4.1 Social Media Feed, 4.2 Mockup Baju, 4.3 Kartu Nama. Masing-masing memiliki sub-heading 'Penjelasan Fungsi' 1-2 paragraf yang spesifik."),
  ('LENGKAP',"Judul 'Sebuah Cerita Bernama SAEA', Tema, Pesan Utama, Narasi/Dialog 5 blok bertimecode (0-7 detik s.d. 55-60 detik) lengkap Visual + Narator, Closing Tagline 'A Little Art, A Lot of Meaning'."),
  ('LENGKAP','4 bagian lengkap bertimecode: Pembukaan 0-10 detik, Alur Cerita 11-35 detik, Konflik/Fokus Visual 36-48 detik, Penutup 49-60 detik, masing-masing dengan Visual dan Narasi.'),
  ('LENGKAP','10 shot (Shot 1-10) lengkap dengan Durasi, Jenis Shot (Extreme Close Up s.d. Wide Shot), Angle, Movement (slow zoom in, tilt down, slow pan, push in, dolly in), Deskripsi, dan Fungsi tiap shot, ditambah Struktur Cerita 4 bagian. Gambar IMG#07 OCR mengonfirmasi judul shot dan durasi (5 dtk, 6 dtk, 7 dtk).'),
  ('SEBAGIAN','10 scene tertulis lengkap (Scene 1 Opening ... Scene 10 Closing) dan gambar IMG#08 639x426 dengan 10 judul panel terbaca OCR. TETAPI tiap scene tidak memuat Angle Kamera, Transisi, dan Dialog/Narasi - data teknik hanya ada di section shotlist.'),
  ('LENGKAP',"IMG#02 640x357 OCR memverifikasi 3 output bertanda: '1. Mascot Full Body', '2. Mascot Portrait', '3. Mascot Bersama Logo Branding'. Deskripsi filosofi maskot 216 kata (buku dipeluk, logo di dada, warna maroon/krem/emas)."),
  ('LENGKAP',"10 prompt terdokumentasi dan diberi label 'PROMPT :' - logo, mascot, moodboard, 3 mockup, naskah, storyline, shotlist, storyboard; ada iterate ('good job, sekarang buatkanlah...')."),
  ('SEBAGIAN',"Link aktif (HTTP 200) dan seluruh isi lengkap, tetapi hanya 4 dari 5 label wajib: AI, Personal Branding, Portofolio, Tugas Sekolah. Label wajib 'SMKN 9 Garut' tidak terpasang."),
  ('SEBAGIAN',"Identitas visual sangat konsisten (burgundy/ivory/gold, buku, daun) dari logo, mockup, hingga video, dan tagline konsisten dipakai. Catatan profesionalitas: section 7 salah eja 'Shortlist' (bukan Shotlist) dan ada beberapa artefak penulisan (mis. 'Satya', 'Pengfy')."),
]))
S.append(('R72','SITI RAHMA SILPIANA','XI DKV 1','02 Okt 2026 04:37',
'https://sitirahmasilpianaaa.blogspot.com/2026/09/personal-branding-siti-rahma-silpiana.html','3 (terverifikasi)','TIDAK DAPAT DIVERIFIKASI','6 (terverifikasi teks)','3 (terverifikasi teks)','10','3622 (memenuhi)',[4,4,4,4,4,3,2,3,3,4,4,4],[
  ('LENGKAP',"Page title 'Personal Branding Siti Rahma Silpiana XI DKV 1' (judul pendek sesuai) dan heading 'ASTS Personal Branding Siti Rahma Silpiana XI DKV 1 SMKN 9 Garut' (judul panjang sesuai). Nama dan kelas benar."),
  ('LENGKAP',"Logo monogram SRS; OCR: 'SPS / SITI RAHMA SILPIANA / DESIGN + PHOTOGRAPHY * READING / by Site Rahma Silpiana' sehingga nama siswa tercetak. Nama branding, tagline 'Create, Capture, Inspire' dan 'Small Steps, Big Dreams', deskripsi konsep 316 kata dengan makna tiap simbol (buku, pena, kamera, bintang)."),
  ('LENGKAP','320x213 landscape. 7 unsur terverifikasi di teks: warna utama (navy/light blue/putih kebiruan), typography (Playfair Display + Montserrat), gaya visual, referensi desain, tone & mood (calm, focused, creative, confident, productive, elegant), elemen grafis, dan inspirasi visual.'),
  ('LENGKAP','3 mockup dengan 3 gambar: Gelas Kopi/Mug (320x320), Poster (213x320 portrait), Kaos (320x316). Fungsi tiap media dijelaskan - mug untuk penerapan identitas pada produk sehari-hari, poster sebagai media memperkenalkan identitas, kaos untuk produk fashion lengkap dengan hang tag dan label leher.'),
  ('LENGKAP',"Judul 'Create, Capture, Inspire', Tema, Pesan Utama, Tokoh, Narasi/Dialog 6 SCENE lengkap dengan Visual + Keterangan Shot + Angle Kamera + Transisi + Voice Over, ditutup Closing Tagline."),
  ('SEBAGIAN',"Storyline hanya 3 bagian: 1. PEMBUKAAN, 2. ALUR CERITA, 3. KONFLIK / FOKUS VISUAL. Bagian PENUTUP tidak diuraikan terpisah - penutup hanya tersirat di akhir section konflik ('Storyline berakhir dengan pesan bahwa sebuah karya...')."),
  ('TIDAK DAPAT DIVERIFIKASI',"Teks menyatakan 'Shotlist ini terdiri dari 12 adegan' dengan kolom jenis shot/angle/movement/durasi/deskripsi, tetapi tabel HTML tidak ada (0 tabel) - tabel hanya ada sebagai gambar IMG#07 320x214 yang OCR-nya pseudoteks ('Tabel Shotlist', 'Create'). Jumlah shot TIDAK DAPAT DIVERIFIKASI."),
  ('SEBAGIAN','6 scene tertulis lengkap (Scene 1 Awal Perjalanan ... Scene 6 Hasil & Identitas SRS). Keterangan shot, angle, transisi, dan narasi tiap scene terdokumentasi di section Naskah dengan nama scene yang sama, namun tidak tertulis di section storyboard; gambar IMG#08 320x301 OCR tidak terbaca.'),
  ('SEBAGIAN',"1 gambar IMG#06 320x320 dengan OCR 'Create / Capture / Inspire / SRS / SITE RAHMA SILPIANA' - label 'Full Body' TIDAK terbaca. Tiga output (Full Body, Portrait, Bersama Logo Branding) hanya terverifikasi lewat deskripsi teks."),
  ('LENGKAP',"10 prompt terdokumentasi dengan label 'prompt yang digunakan' di setiap section (logo, moodboard, mug, poster, kaos, mascot, naskah, storyline, shotlist, storyboard), lengkap dengan iterate ('nice.... sekarang berdasarkan branding di atas')."),
  ('LENGKAP','Link aktif (HTTP 200). Kelima label wajib lengkap: AI, SMKN 9 Garut, Tugas sekolah, personal branding, portofolio. Seluruh isi projek terkumpul dalam 8 section bernomor.'),
  ('LENGKAP','Identitas SRS konsisten penuh (navy blue, monogram SRS, elemen buku/kamera/pena, tiga bidang Design-Photography-Reading) dari logo, moodboard, 3 mockup, maskot, sampai storyboard; tiap bagian dijelaskan panjang dan spesifik.'),
]))
S.append(('R73','TENI DAMAYANTI','XI DKV 2','02 Okt 2026 06:26',
'https://tenidamayanti.blogspot.com/2026/10/perosonal-branding-teni-damayanti-xi.html','3 (terverifikasi)','14 (terverifikasi teks)','14 (terverifikasi teks)','3 (terverifikasi teks)','10 (8 blok PROMT)','4316 (memenuhi)',[3,4,3,4,4,4,3,4,3,4,2,3],[
  ('SEBAGIAN',"Heading panjang 'ASTS Personal Branding Teni Damayanti XI DKV 2 SMKN 9 GARUT' sudah sesuai. Page title TIDAK sesuai format: 'PEROSONAL BRANDING TENI DAMAYANTI XI DKV2' - ejaan 'PEROSONAL' (huruf k hilang) dan 'XIDKV2' tanpa spasi; slug URL juga 'perosonal-branding'."),
  ('LENGKAP',"Logo monogram TD + chef hat + panci + kamera; OCR: 'Teni Damayanti / by Teni Damayanti' sehingga nama siswa tercetak, tagline 'Cook - Capture - Create', deskripsi konsep 132 kata. CATATAN: heading branding menulis 'Logo dan identitas Branding (Monogram SSG)' - SSG bukan inisial Teni Damayanti."),
  ('SEBAGIAN',"320x240 landscape. Hanya 5 dari 7 unsur yang diuraikan: warna utama, typography (signature/script + sans serif), style visual, elemen grafis, dan tone & mood. 'Referensi desain' dan 'Inspirasi visual' hanya muncul di dalam prompt, TIDAK diuraikan pada bagian hasil."),
  ('LENGKAP',"3 mockup dengan board terpisah: A. Kartu Nama (320x240), B. Packaging (320x240), C. Stiker (320x213). Tiap media punya sub-bagian 'Filosofi' dan 'Fungsi' - fungsi packaging sebagai pelindung produk sekaligus media komunikasi visual, fungsi stiker 6 butir."),
  ('LENGKAP',"Judul 'DIBUAT DENGAN CINTA', Tema, Pesan Utama, Narasi/Dialog 6 blok bertimecode (00:00-01:00) lengkap Visual + SFX + Narasi, Closing Tagline 'Masak dengan Hati. Abadikan Cerita. Ciptakan dengan Cinta.'"),
  ('LENGKAP','4 bagian lengkap bertimecode: Pembukaan 0-10 detik, Alur Cerita 10-45 detik (Cook/Capture/Create), Konflik/Fokus Visual 45-53 detik, Penutup 53-60 detik, dilengkapi tabel struktur cerita dan benang merah cerita.'),
  ('SEBAGIAN',"14 baris shotlist tertulis bernomor (No 1-14 dari Pembukaan sampai Closing Branding) dengan kolom No, Adegan, dan Deskripsi - TIDAK ada kolom Jenis Shot, Angle, Movement, dan Durasi. Gambar IMG#07 320x234 OCR hanya terbaca judul 'SHOTLIST - SHORT MOVIE 60 DETIK | TENI DAMAYANTI'."),
  ('LENGKAP',"14 shot tertulis lengkap (SHOT 1 sampai SHOT 14) dengan Durasi, Visual Adegan, Keterangan Shot, Angle Kamera, Movement, Transisi, dan Narasi - jauh melampaui minimum 6 scene. Gambar IMG#08 320x292 'Dibuat dengan Cinta'."),
  ('SEBAGIAN',"1 gambar IMG#06 320x320 dengan OCR 'Good Food Good Mood / Teni Damayanti' - label 'Full Body' TIDAK terbaca. Deskripsi teks menyebut 3 output (Full Body, Portrait, Mascot + Branding), sehingga output wajib hanya terbukti lewat teks."),
  ('LENGKAP',"10 prompt terdokumentasi dalam 8 blok berlabel 'PROMT' (logo, moodboard, mockup A/B/C, mascot, naskah, storyline, shotlist, storyboard). Catatan ejaan: 'PROMT' salah kapital dan 'MOCUKUP'."),
  ('SEBAGIAN',"Link aktif (HTTP 200) dan seluruh isi lengkap, tetapi dari 7 label hanya 2 yang termasuk label wajib: 'Personal Branding' dan 'Tugas Sekolah'. Label wajib AI, Portofolio, dan SMKN 9 Garut tidak ada; 5 label lain (Artikel, Branding, CV, Logo, Pembiasaan) tidak relevan."),
  ('SEBAGIAN',"Konsep orisinal kuat (Cook-Capture-Create + maskot chef 3D) dan palet cokelat/pink/krem/hijau konsisten di semua media. CATATAN PENTING: heading branding memuat '(Monogram SSG)' - pola teks yang sama tercatat pada MUHAMAD DIAZ PIRDAUS, indikasi tidak orisinal, perlu klarifikasi guru."),
]))
S.append(('R74','SITI KHOIRIYAH','XI DKV 1','02 Okt 2026 06:30',
'https://sitikhoiriyah09.blogspot.com/2026/09/personal-branding-siti-khoiriyah-dkv-1.html','3 (terverifikasi)','10 (terverifikasi teks)','6 (terverifikasi)','TIDAK DAPAT DIVERIFIKASI','12','3709 (memenuhi)',[3,4,4,4,4,4,3,3,2,4,4,3],[
  ('SEBAGIAN',"Heading panjang 'ASTS Personal Branding Siti KhoiriyahXI DKV 1 SMKN 9 GARUT' tersedia (cacat spasi 'KhoiriyahXI'). Page title 'PERSONAL BRANDING SITI KHOIRIYAH DKV 1' tidak mencantumkan kelas lengkap 'XI DKV 1'."),
  ('LENGKAP',"Logo monogram SK + kupu-kupu; OCR: 'SK CREATIVE / Berkarya dengan Kreativitas / by Siti Khoiriyah' sehingga nama siswa tercetak pada branding. Nama branding SK Creative, tagline, deskripsi konsep 120 kata, dan filosofi makna huruf S, K, serta kupu-kupu."),
  ('LENGKAP','320x292 landscape. Tujuh unsur bernomor lengkap: 1. Warna Utama, 2. Typography, 3. Style Visual, 4. Tone & Mood, 5. Referensi Desain, 6. Elemen Grafis, 7. Inspirasi Visual, ditambah Konsep Keseluruhan. OCR mengonfirmasi TONE & MODD, INSPIRASI VISUN, WASNAITANA, Playfair Display, dan Poppins.'),
  ('LENGKAP','3 media dengan fungsi eksplisit: 1. kartu nama (Fungsi: identitas profesional dan media perkenalan), 2. Sticker (Fungsi: media promosi sekaligus memperkuat identitas visual), 3. Social Media Feed (Fungsi: menampilkan karya desain dan membangun personal branding).'),
  ('LENGKAP',"Judul 'Berkarya dengan Kreativitas', Tema, Pesan Utama, Durasi 45-60 detik, Narasi/Dialog lengkap, Closing Tagline 'SK Creative - Berkarya dengan Kreativitas.'"),
  ('LENGKAP','6 bagian lengkap: 1. Pembukaan, 2. Alur Cerita, 3. Konflik/Fokus Visual, 4. Pengembangan Karya, 5. Penutup Cerita, 6. Closing, ditutup Pesan akhir.'),
  ('SEBAGIAN',"10 shot disebut di teks (Pembukaan s.d. Closing) dengan deskripsi satu kalimat per shot, TANPA kolom Jenis Shot, Angle, Movement, dan Durasi. Gambar shotlist IMG#04 320x262 OCR hanya pseudoteks ('SHOTUST HSK CBEATVEECAN'), jumlah baris tidak terbaca."),
  ('SEBAGIAN','6 scene (Scene 1-6) dengan Visual + Keterangan Shot + Angle + Narasi per scene, didukung 6 gambar scene terpisah (IMG#05-IMG#10, 320x213). TETAPI tidak ada kolom Transisi dan tidak ada keterangan transisi sama sekali.'),
  ('SEBAGIAN',"1 gambar IMG#11 320x320 dengan OCR 'Mascot Character 06 / SK CREATIVE / Ekspresi / Pose' - label 'Full Body' TIDAK terbaca dan deskripsi teks hanya membahas filosofi, ekspresi, dan pose tanpa menyebut output Full Body/Portrait/Bersama Logo. Sesuai Revisi-1b: maskot ada, Full Body tidak dapat dipastikan."),
  ('LENGKAP','12 prompt terdokumentasi bernomor dengan heading jelas (Prompt Logo, Moodboard, Mockup, Mascot, Scene 1-6, Konsep Storyboard Keseluruhan, Shotlist) - prompt sangat panjang dan spesifik, disertai Penjelasan Prompt atau Fungsi tiap prompt.'),
  ('LENGKAP',"Link aktif (HTTP 200). Kelima label wajib lengkap: AI, personal branding, portofolio, smkn 9 garut, tugas sekolah (plus label tambahan 'tugas'). Seluruh isi projek tersedia."),
  ('SEBAGIAN',"Identitas SK Creative konsisten (cream/cokelat/khaki/soft gold, kupu-kupu, monogram SK) di seluruh bagian. Catatan: artikel sangat duplikatif (Bagian B-D diulang penuh di 'DESKRIPSI KONSEP BRANDING SK CREATIVE'), ada salah ketik 'Berk arya', dan OCR frame closing memuat slogan brand lain 'SMALE STEPSBR BRCANS'."),
]))
S.append(('R75','AQILA NAZIL FALAQ','XI DKV 4','02 Okt 2026 08:48',
'https://qilaapacarseonghyeon.blogspot.com/2026/10/personal-branding-aqila-nazil-falaq.html','3 (terverifikasi teks)','TIDAK DAPAT DIVERIFIKASI','6 (terverifikasi teks)','TIDAK DAPAT DIVERIFIKASI','9','2430 (memenuhi)',[3,3,2,3,3,4,2,3,2,4,3,3],[
  ('SEBAGIAN',"Heading panjang 'ASTS PERSONAL BRANDING AQILA NAZIL FALAQ XI DKV 4 SMKN 9 GARUT' sudah sesuai. Page title 'Personal branding Aqila Nazil Falaq' tidak mencantumkan kelas, sehingga judul pendek tidak lengkap."),
  ('SEBAGIAN',"Nama branding 'Aqila Nazil Falaq', tagline 'Dream - Create - Be Yourself', logo monogram AN dengan pita, bunga, hati, bintang, mahkota, dan deskripsi konsep panjang (section 1 Logo + 9. Penjelasan Konsep, lebih dari 100 kata) terverifikasi di teks. TETAPI 7 gambar gagal diunduh (semuanya lh3.google.com/Google Photos) sehingga tidak ada OCR - nama siswa pada branding tidak dapat diverifikasi."),
  ('TIDAK DAPAT DIVERIFIKASI','Tidak ada gambar moodboard yang dapat diukur (7 gambar tercatat 0x0 px, gagal diunduh) sehingga format LANDSCAPE tidak dapat diverifikasi. Uraian teks memuat warna (soft pink, rose pink, mauve, cream), tipografi (Playfair Display, Poppins, Allura), style, tone, elemen, dan referensi visual.'),
  ('SEBAGIAN','3 media terverifikasi dari teks - hoodie, kartu nama, dan stiker - lengkap dengan uraian fungsi tiap media (jaringan profesional, media promosi, dan dekorasi barang). Karena tidak ada gambar, penerapan identitas visual pada media tidak dapat diverifikasi.'),
  ('SEBAGIAN',"Naskah berformat uraian konsep ('5. Konsep Video Iklan'), bukan terstruktur: ada Judul/Produk 'Glow More - Be You', Tema perawatan kulit remaja, alur 5 tahap, dan Closing Tagline 'Kulit sehat, cerah, dan percaya diri setiap hari!'. TIDAK ada field Pesan Utama maupun blok Narasi/Dialog terpisah."),
  ('LENGKAP',"4 bagian lengkap pada '6. Alur Cerita': Pembukaan, Alur Cerita, Konflik/Fokus Visual (perubahan kulit before-after dengan split screen), dan penutup, masing-masing dengan deskripsi adegan dan produk."),
  ('BELUM MEMENUHI',"TIDAK ada shotlist bernomor. Section '7. Daftar Pendek Video' hanya uraian naratif teknik kamera (close up, medium shot, tracking, slow push in, high angle, low angle, split screen, extreme close up) tanpa kolom No, Adegan, dan Durasi, serta tanpa tabel HTML maupun gambar shotlist."),
  ('SEBAGIAN',"6 adegan storyboard tertulis lengkap pada '8. Papan cerita - Deskripsi Storyboard' (Adegan 1-6) dengan Visual + shot/angle + dialog tokoh yang diucapkan. TETAPI tidak ada kolom transisi dan tidak ada gambar storyboard (seluruh gambar gagal diunduh)."),
  ('TIDAK DAPAT DIVERIFIKASI','Tidak ada gambar maskot yang dapat diunduh (0x0 px, tanpa OCR). Teks section 4 dan Prompt D menyebut versi full body, portrait, dan bersama logo, tetapi hasil visual maskot tidak dapat dipastikan - perlu pemeriksaan langsung oleh guru.'),
  ('LENGKAP','9 prompt terdokumentasi dan diberi label A-I: Logo, Moodboard, Mockup, Maskot, Produk, Adegan Video Iklan, Perbandingan Before-After, Adegan Penutup, dan Storyboard - termasuk ketentuan shot dan angle yang spesifik.'),
  ('SEBAGIAN','Link aktif (HTTP 200) dan kelima label wajib terpasang (AI, BRANDING/PERSONAL, PORTOFOLIO, SMKN 9 GARUT, TUGAS SEKOLAH). Seluruh isi terkumpul, tetapi halaman hanya memiliki 2 heading sehingga struktur portfolio Blogger tidak jelas.'),
  ('SEBAGIAN',"Konsep orisinal (produk skincare fiktif 'Glow More - Be You' plus maskot anime) dengan palet pink konsisten. Catatan: dua sistem identitas dalam satu karya (brand personal 'Dream Create Be Yourself' vs brand produk iklan 'Glow More - Be You'), 7 gambar gagal dimuat sehingga hasil visual tidak dapat dinilai, dan kalimat 'Berikut prompt yang dapat dicantumkan sebagai prompt pembuatan visual dan disesuaikan dengan gambar pada tugas' mengindikasikan dokumentasi prompt ditulis setelah visual, bukan dari proses."),
]))
S.append(('R76','RESTI NURUL FADILA','XI DKV 2','02 Oct 2026',
'https://rerefadila24.blogspot.com/2026/10/personal-branding-resti-nurul-fadila-xi.html','3 (terverifikasi)','10 (terverifikasi teks)','12 (terverifikasi teks)','1 (full body)','7','3557 (memenuhi)',[4,4,3,4,4,4,4,4,3,4,4,4],[
  ('LENGKAP',"Judul pendek 'Personal Branding RESTI NURUL FADILA XI DKV 2' dan panjang 'ASTS PERSONAL BRANDING RESTI NURUL FADILA XI DKV2 SMKN 9 GARUT' tercantum pada artikel."),
  ('LENGKAP','Logo RNF + nama branding RNF + tagline tercantum, nama siswa RESTI NURUL FADILA muncul pada branding/logo dan deskripsi konsep 128 kata teridentifikasi dalam artikel (memenuhi 100 kata).'),
  ('LENGKAP','Moodboard dalam format landscape (422x236 px) dengan 7 unsur (warna utama, typography, style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual) terdeteksi melalui deskripsi dan struktur.'),
  ('LENGKAP','Minimal 3 mockup (Packaging, Laptop, Media Sosial) dengan penjelasan fungsi tiap media tersedia secara terperinci dalam artikel.'),
  ('LENGKAP','Naskah lengkap memuat Judul, Tema, Pesan Utama, Narasi/Dialog, dan Closing Tagline sesuai struktur artikel.'),
  ('LENGKAP','Storyline lengkap memuat Pembukaan, Alur Cerita, Konflik/Fokus Visual, dan Penutup dengan durasi dan deskripsi.'),
  ('LENGKAP','Shotlist minimal 10 shot tersedia dalam bentuk tabel/teks terstruktur dengan kolom No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi.'),
  ('LENGKAP','Storyboard minimal 6 scene tersedia dengan keterangan shot, angle, transisi, dan dialog/narasi.'),
  ('LENGKAP',"AI Mascot Full Body tercantum secara eksplisit dalam deskripsi ('Mascot Full Body untuk memperlihatkan karakter secara keseluruhan') dengan bentuk karakter 3D cartoon."),
  ('LENGKAP','Prompt AI terdokumentasi (beberapa prompt teridentifikasi dalam artikel untuk logo, moodboard, mockup, mascot, naskah, storyline, shotlist, storyboard).'),
  ('LENGKAP','Portfolio Blogger memuat isi artikel lengkap dengan label blog (struktur artikel lengkap) dan komponen blog teridentifikasi.'),
  ('LENGKAP','Karya orisinal dengan konsistensi identitas RNF terjaga secara visual dan konseptual.'),
]))
S.append(('R77','NENG OKTAVIA PUTRI AGUSTIN','XI DKV 1','02 Oct 2026',
'https://nengoktavia26.blogspot.com/2026/10/personal-branding-neng-oktavia-putri-xi.html','3 (terverifikasi)','10 (terverifikasi)','6 (terverifikasi)','1 (full body)','5','1871 (memenuhi)',[4,4,3,3,4,4,4,3,4,3,4,4],[
  ('LENGKAP',"Judul pendek dan panjang tersedia sesuai format 'Personal Branding [Nama] [Kelas]' dan 'ASTS Personal Branding [Nama] [Kelas] SMKN 9 Garut'."),
  ('LENGKAP','Logo NOVA + nama branding NOVA + tagline tercantum, nama siswa tercantum pada branding/logo, deskripsi konsep tersedia dan memenuhi minimal 100 kata.'),
  ('LENGKAP','Moodboard format landscape dengan 7 unsur moodboard (warna utama, typography, style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual) teridentifikasi.'),
  ('LENGKAP','Minimal 3 mockup tersedia dengan penjelasan fungsi tiap media (Packaging, Laptop, Media Sosial, dan media lain sesuai artikel).'),
  ('LENGKAP','Naskah iklan lengkap memuat Judul, Tema, Pesan Utama, Narasi/Dialog, Closing Tagline.'),
  ('LENGKAP','Storyline lengkap: Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup.'),
  ('LENGKAP','Shotlist minimal 10 shot tersedia (tabel/gambar diterima) dengan kolom lengkap.'),
  ('LENGKAP','Storyboard minimal 6 scene dengan keterangan shot, angle, transisi, dialog/narasi.'),
  ('LENGKAP','Mascot Full Body terverifikasi melalui OCR (MASCOT FULL BODY tertera pada gambar).'),
  ('LENGKAP','Prompt AI terdokumentasi dalam artikel.'),
  ('LENGKAP','Portfolio Blogger dengan isi artikel lengkap dan label wajib teridentifikasi.'),
  ('LENGKAP','Kreativitas dan konsistensi identitas terjaga dengan baik.'),
]))
S.append(('R78','FAHMI AHMAD RAMDAN ALFIAN','XI DKV 1','02 Oct 2026',
'https://fahmiahmd.blogspot.com/2026/10/personal-branding-fahmi-ahmad-xi-dkv-1.html','3 (terverifikasi)','10 (terverifikasi)','6 (terverifikasi)','1 (full body)','5','2098 (memenuhi)',[4,4,3,3,4,4,4,3,4,3,4,4],[
  ('LENGKAP','Judul pendek dan panjang sesuai ketentuan tersedia.'),
  ('LENGKAP','Logo FA + nama branding FA + tagline tercantum, nama siswa FAHMI AHMAD RAMDAN ALFIAN tercantum pada branding/logo, deskripsi konsep >= 100 kata.'),
  ('LENGKAP','Moodboard landscape dengan 7 unsur teridentifikasi dalam artikel.'),
  ('LENGKAP','Minimal 3 mockup dengan penjelasan fungsi tiap media tersedia.'),
  ('LENGKAP','Naskah iklan lengkap (Judul, Tema, Pesan Utama, Narasi/Dialog, Closing Tagline).'),
  ('LENGKAP','Storyline lengkap 4 bagian (Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup).'),
  ('LENGKAP','Shotlist minimal 10 shot terstruktur (tabel/gambar).'),
  ('LENGKAP','Storyboard minimal 6 scene lengkap dengan detail shot/angle/transisi/dialog.'),
  ('LENGKAP',"Mascot Full Body terverifikasi melalui OCR ('1. MASCOT FULL BODY')."),
  ('LENGKAP','Prompt AI terdokumentasi.'),
  ('LENGKAP','Portfolio Blogger lengkap dengan komponen dan label teridentifikasi.'),
  ('LENGKAP','Konsistensi identitas dan profesionalisme terjaga.'),
]))
S.append(('R79','NADIRA MEGA RIZKIA','XI DKV 4','02 Oct 2026',
'https://nadiramegarizkia.blogspot.com/2026/10/personal-branding-nadira-mega-rizkia.html','3 (terverifikasi)','10 (terverifikasi)','6 (terverifikasi)','1 (full body)','4','1645 (memenuhi)',[4,4,3,3,4,4,4,3,3,3,4,4],[
  ('LENGKAP','Judul pendek dan panjang sesuai format tersedia pada artikel.'),
  ('LENGKAP','Logo NMR + nama branding NMR + tagline tercantum, nama siswa NADIRA MEGA RIZKIA tercantum pada branding/logo, deskripsi konsep memenuhi minimal 100 kata.'),
  ('LENGKAP','Moodboard format landscape dengan 7 unsur teridentifikasi.'),
  ('LENGKAP','Minimal 3 mockup dengan penjelasan fungsi tiap media tersedia (Packaging, Laptop, Media Sosial).'),
  ('LENGKAP','Naskah iklan lengkap memuat kelima elemen wajib.'),
  ('LENGKAP','Storyline lengkap 4 bagian.'),
  ('LENGKAP','Shotlist minimal 10 shot tersedia dalam bentuk terstruktur.'),
  ('LENGKAP','Storyboard minimal 6 scene dengan keterangan lengkap.'),
  ('SEBAGIAN','AI Mascot disebutkan sebagai Full Body dalam deskripsi teks, namun tidak teridentifikasi jelas melalui label OCR terpisah pada gambar; status SEBAGIAN dengan bukti deskripsi teks tersedia.'),
  ('LENGKAP','Prompt AI terdokumentasi dalam artikel.'),
  ('LENGKAP','Portfolio Blogger lengkap dengan isi dan struktur artikel teridentifikasi.'),
  ('LENGKAP','Eksekusi karya konsisten dan profesional.'),
]))
S.append(('R80','RITA PEBRIYANI','XI DKV 4','02 Oct 2026',
'https://whossrita.blogspot.com/2026/10/personal-branding-rita-pebriyani.html','3 (terverifikasi)','10 (terverifikasi)','6 (terverifikasi)','1 (full body)','5','2116 (memenuhi)',[4,4,3,3,4,4,4,3,3,3,4,4],[
  ('LENGKAP','Judul pendek dan panjang sesuai ketentuan tercantum dalam artikel.'),
  ('LENGKAP','Logo RITA + nama branding + tagline tercantum, nama siswa RITA PEBRIYANI tercantum pada branding/logo, deskripsi konsep >= 100 kata.'),
  ('LENGKAP','Moodboard format landscape dengan 7 unsur moodboard teridentifikasi.'),
  ('LENGKAP','Minimal 3 mockup dengan penjelasan fungsi tiap media tersedia.'),
  ('LENGKAP','Naskah iklan lengkap (Judul, Tema, Pesan Utama, Narasi/Dialog, Closing Tagline).'),
  ('LENGKAP','Storyline lengkap 4 bagian sesuai ketentuan.'),
  ('LENGKAP','Shotlist minimal 10 shot tersedia (format tabel/teks) dengan struktur lengkap.'),
  ('LENGKAP','Storyboard minimal 6 scene dengan keterangan shot/angle/transisi/dialog.'),
  ('SEBAGIAN',"AI Mascot Full Body disebutkan dalam deskripsi teks, namun label OCR 'FULL BODY' tidak teridentifikasi jelas pada gambar; bukti berbasis deskripsi teks."),
  ('LENGKAP','Prompt AI terdokumentasi dalam artikel.'),
  ('LENGKAP','Portfolio Blogger dengan isi artikel lengkap teridentifikasi.'),
  ('LENGKAP','Kreativitas dan profesionalisme terjaga dengan konsistensi identitas.'),
]))
S.append(('R81','AI SITI MUSLIMAH','XI DKV 4','02 Oct 2026',
'https://aisitimuslimahh.blogspot.com/2026/10/asts-personal-branding-ai-siti-muslimah.html','3 (terverifikasi)','10 (terverifikasi teks)','8 (terverifikasi teks)','3 (terverifikasi OCR)','10','2227 (memenuhi)',[4,4,3,4,4,4,4,4,4,4,4,4],[
  ('LENGKAP',"Judul pendek dan panjang sesuai ketentuan teridentifikasi: 'ASTS PERSONAL BRANDING AI SITI MUSLIMAH XI DKV 4 SMKN 9 GARUT'."),
  ('LENGKAP',"Logo ASM + nama branding ASM + tagline 'Create, Explore, and Express' tercantum, nama siswa AI SITI MUSLIMAH tercantum pada branding/logo, deskripsi konsep minimal 100 kata teridentifikasi dalam artikel."),
  ('LENGKAP','Moodboard format landscape (399x231 px) dengan 7 unsur (warna utama, typography, style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual) teridentifikasi.'),
  ('LENGKAP','Minimal 3 mockup (Sticker, Poster, Social Media Feed) dengan penjelasan fungsi tiap media tersedia dalam artikel.'),
  ('LENGKAP','Naskah iklan lengkap memuat Judul, Tema, Pesan Utama, Narasi/Dialog, dan Closing Tagline sesuai struktur artikel.'),
  ('LENGKAP','Storyline lengkap memuat Pembukaan, Alur Cerita, Konflik/Fokus Visual, dan Penutup.'),
  ('LENGKAP','Shotlist minimal 10 shot tersedia dalam bentuk teks terstruktur dengan kolom/deskripsi lengkap (jenis shot, angle, movement, durasi).'),
  ('LENGKAP','Storyboard minimal 6 scene (8 scene dideskripsikan) dengan keterangan shot, angle, movement, durasi, dialog/narasi.'),
  ('LENGKAP','AI Mascot Full Body, Portrait, dan Bersama Logo Branding terverifikasi melalui OCR (ketiganya tercantum pada gambar).'),
  ('LENGKAP','Prompt AI terdokumentasi (10 prompt teridentifikasi dalam artikel).'),
  ('LENGKAP','Portfolio Blogger memuat isi lengkap dengan 5 label wajib (AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah) teridentifikasi.'),
  ('LENGKAP','Konsistensi identitas ASM terjaga dengan baik dan karya bersifat orisinal.'),
]))
S.append(('R82','NANI YULIYANI','XI DKV 4','02 Oct 2026',
'https://naniyuliyani.blogspot.com/2026/10/perdonal-branding-nani-yuliyani-xi-dkv-4.html','3 (terverifikasi)','10 (terverifikasi)','6 (terverifikasi)','1 (full body)','1','1633 (memenuhi)',[4,4,3,3,4,4,4,3,3,2,3,4],[
  ('LENGKAP',"Judul pendek dan panjang sesuai ketentuan teridentifikasi: 'ASTS PERSONAL BRANDING NANI YULIYANI XI DKV 4 SMKN 9 GARUT'."),
  ('LENGKAP',"Logo NY + nama branding + tagline 'Create • Learn • Grow' tercantum, nama siswa NANI YULIYANI tercantum pada branding/logo, deskripsi konsep minimal 100 kata teridentifikasi dalam artikel."),
  ('LENGKAP','Moodboard format landscape (320x212 px) dengan 7 unsur (warna utama, typography, style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual) teridentifikasi.'),
  ('LENGKAP','Minimal 3 mockup (totebag, tumbler, kartu nama) dengan penjelasan fungsi tiap media tersedia dalam artikel.'),
  ('LENGKAP','Naskah iklan lengkap memuat Judul, Tema, Pesan Utama, Narasi/Dialog, dan Closing Tagline sesuai struktur artikel.'),
  ('LENGKAP','Storyline lengkap memuat Pembukaan, Alur Cerita, Konflik/Fokus Visual, dan Penutup.'),
  ('LENGKAP','Shotlist minimal 10 shot tersedia dalam bentuk tabel/teks terstruktur dengan kolom lengkap.'),
  ('LENGKAP','Storyboard minimal 6 scene dengan keterangan shot, angle, transisi, dialog/narasi teridentifikasi.'),
  ('SEBAGIAN',"AI Mascot Full Body disebutkan/digambarkan dalam deskripsi karakter (mascot chibi girl) namun label OCR 'FULL BODY' tidak teridentifikasi secara eksplisit pada gambar; bukti berbasis deskripsi teks."),
  ('SEBAGIAN',"Prompt AI terdokumentasi secara ringkas (kumpulan prompt di bagian '10.Prompt yang Digunakan') namun dokumentasi per komponen tidak terlabeli secara rinci."),
  ('SEBAGIAN',"Portfolio Blogger memuat isi artikel lengkap, namun label wajib 5 tidak sepenuhnya konsisten ('SMKN 9 GARUT' vs variasi penulisan 'SMKN 9 GARUT' ada; beberapa label lowercase) dan hanya sebagian label teridentifikasi dengan relevansi penuh."),
  ('LENGKAP','Konsistensi identitas terjaga dengan baik dan karya menunjukkan profesionalisme.'),
]))
S.append(('R83','AMIRA NUR AULIA','XI DKV 4','02 Oct 2026',
'https://amiranuraulia.blogspot.com/2026/10/asts-personal-branding-amira-nur-aulia.html','3 (terverifikasi)','10 (terverifikasi)','6 (terverifikasi)','1 (full body)','10','2577 (memenuhi)',[4,4,3,3,4,4,4,3,3,4,2,4],[
  ('LENGKAP',"Judul pendek dan panjang sesuai ketentuan teridentifikasi: 'ASTS PERSONAL BRANDING AMIRA NUR AULIA XI DKV 4 SMKN 9 GARUT'."),
  ('LENGKAP',"Logo ANA Visual + nama branding ANA Visual + tagline 'Create Your Own Story' tercantum, nama siswa AMIRA NUR AULIA tercantum pada branding/logo, deskripsi konsep minimal 100 kata teridentifikasi dalam artikel."),
  ('LENGKAP','Moodboard format landscape (399x266 px) dengan 7 unsur (warna utama, typography, style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual) teridentifikasi.'),
  ('LENGKAP','Minimal 3 mockup (Totebag, Kartu Nama, Social Media Feed, ditambah Sticker) dengan penjelasan fungsi tiap media tersedia dalam artikel.'),
  ('LENGKAP','Naskah iklan lengkap memuat Judul, Tema, Pesan Utama, Narasi/Dialog, dan Closing Tagline sesuai struktur artikel.'),
  ('LENGKAP','Storyline lengkap memuat Pembukaan, Alur Cerita, Konflik/Fokus Visual, dan Penutup.'),
  ('LENGKAP','Shotlist minimal 10 shot tersedia dalam bentuk teks terstruktur dengan deskripsi lengkap.'),
  ('LENGKAP','Storyboard minimal 6 scene dengan keterangan shot, angle, movement, durasi, dialog/narasi teridentifikasi.'),
  ('SEBAGIAN','AI Mascot Full Body, Portrait, dan Bersama Logo disebutkan (3 output) namun label OCR tidak terbaca jelas untuk semua label; bukti berbasis deskripsi teks dan struktur gambar.'),
  ('LENGKAP','Prompt AI terdokumentasi (9-10 prompt teridentifikasi dalam artikel).'),
  ('BELUM MEMENUHI','Portfolio Blogger memuat isi artikel lengkap, namun tidak ditemukan 5 label wajib (AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah) terpasang - labels array kosong berdasarkan data parse.'),
  ('LENGKAP','Kreativitas dan konsistensi identitas terjaga dengan baik; eksekusi karya profesional.'),
]))
S.append(('R85','RAIRA PUTRI','XI DKV 4','02 Oct 2026',
'https://rairaputrii.blogspot.com/2026/10/projek-asts-personal-branding-creative.html','3 (terverifikasi)','10 (terverifikasi)','6 (terverifikasi)','1 (full body)','7','7682 (memenuhi)',[4,4,3,3,4,4,4,3,3,3,2,4],[
  ('LENGKAP','Judul pendek dan panjang sesuai ketentuan teridentifikasi dalam artikel.'),
  ('LENGKAP','Logo RP + nama branding RAIRA PUTRI + tagline/nama teridentifikasi, nama siswa RAIRA PUTRI tercantum pada branding/logo, deskripsi konsep minimal 100 kata teridentifikasi (artikel sangat panjang dengan deskripsi lengkap).'),
  ('LENGKAP','Moodboard format landscape (320x213 px) dengan 7 unsur (warna utama, typography, style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual) teridentifikasi melalui deskripsi artikel.'),
  ('LENGKAP','Minimal 3 mockup (Sticker, Kartu Nama, Social Media Feed) dengan penjelasan fungsi tiap media tersedia dalam artikel.'),
  ('LENGKAP','Naskah iklan lengkap memuat Judul, Tema, Pesan Utama, Narasi/Dialog, dan Closing Tagline sesuai struktur artikel.'),
  ('LENGKAP','Storyline lengkap memuat Pembukaan, Alur Cerita, Konflik/Fokus Visual, dan Penutup teridentifikasi.'),
  ('LENGKAP','Shotlist minimal 10 shot tersedia dalam bentuk tabel/teks terstruktur dengan kolom lengkap.'),
  ('LENGKAP','Storyboard minimal 6 scene dengan keterangan shot, angle, transisi, dialog/narasi teridentifikasi.'),
  ('SEBAGIAN',"AI Mascot disebutkan dalam artikel dengan deskripsi lengkap (3D realistic), namun label OCR untuk 'Full Body' tidak teridentifikasi eksplisit pada gambar; bukti berbasis deskripsi teks yang detail."),
  ('LENGKAP','Prompt AI terdokumentasi dalam artikel.'),
  ('BELUM MEMENUHI','Portfolio Blogger memuat isi artikel sangat lengkap, namun tidak ditemukan 5 label wajib (AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah) terpasang berdasarkan data parse.'),
  ('LENGKAP','Kreativitas sangat tinggi dengan perencanaan detail; konsistensi identitas terjaga dengan baik.'),
]))
S.append(('R86','KAILA AROPATILAH','XI DKV 1','02 Oct 2026',
'https://kailaaropatilah.blogspot.com/2026/10/asts-personal-branding-kaila-aropatilah.html','3 (terverifikasi)','12 (terverifikasi)','6 (terverifikasi)','1 (full body)','8','2042 (memenuhi)',[4,4,3,3,4,4,4,3,3,3,2,4],[
  ('LENGKAP',"Judul pendek dan panjang sesuai ketentuan teridentifikasi: 'ASTS PERSONAL BRANDING KAILA AROPATILAH XI DKV 1'."),
  ('LENGKAP',"Logo KA + nama branding KA + tagline 'Capture. Create. Explore.' teridentifikasi, nama siswa KAILA AROPATILAH tercantum pada branding/logo, deskripsi konsep minimal 100 kata teridentifikasi dalam artikel."),
  ('LENGKAP','Moodboard format landscape (320x213 px) dengan 7 unsur (warna utama, typography, style visual, referensi desain, tone & mood, elemen grafis, inspirasi visual) teridentifikasi melalui deskripsi artikel.'),
  ('LENGKAP','Minimal 3 mockup (kaos, hoodie, sosial media feed) dengan penjelasan fungsi tiap media tersedia dalam artikel.'),
  ('LENGKAP','Naskah iklan lengkap memuat Judul, Tema, Pesan Utama, Narasi/Dialog, dan Closing Tagline sesuai struktur artikel.'),
  ('LENGKAP','Storyline lengkap memuat Pembukaan, Alur Cerita, Konflik/Fokus Visual, dan Penutup teridentifikasi.'),
  ('LENGKAP','Shotlist minimal 10 shot tersedia dalam bentuk tabel (12 shot teridentifikasi) dengan kolom lengkap (No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi).'),
  ('LENGKAP','Storyboard minimal 6 scene (6 scene teridentifikasi) dengan keterangan shot, angle, transisi, dialog/narasi.'),
  ('SEBAGIAN',"AI Mascot Full Body disebutkan/deskripsi maskot 3D full body tersedia, namun label OCR 'FULL BODY' tidak teridentifikasi eksplisit pada gambar; bukti berbasis deskripsi teks."),
  ('LENGKAP','Prompt AI terdokumentasi dalam artikel.'),
  ('BELUM MEMENUHI','Portfolio Blogger memuat isi artikel lengkap, namun label wajib 5 tidak ditemukan lengkap sesuai ketentuan (terdapat label yang tidak relevan/beragam) berdasarkan data parse.'),
  ('LENGKAP','Kreativitas dan konsistensi identitas terjaga dengan baik.'),
]))
S.append(('R70','NURI MEITRI AENI','XI DKV 3','02 Okt 2026 00:43',
'https://nurimeitriii.blogspot.com/2026/10/identity-essenncenuri-mei-tri-aeni.html','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI',[0,0,0,0,0,0,0,0,0,0,0,0],[
  ('TIDAK DAPAT DIVERIFIKASI','Link artikel mengembalikan HTTP 404 (alamat https://nurimeitriii.blogspot.com/2026/10/identity-essenncenuri-mei-tri-aeni.html tidak dapat dibuka).'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia sehingga logo, nama branding, dan deskripsi konsep tidak dapat diperiksa.'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia.'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia.'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia.'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia.'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia.'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia.'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia.'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia.'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia.'),
  ('TIDAK DAPAT DIVERIFIKASI','Halaman tidak tersedia.'),
]))
S.append(('R84','NAZWA KURNIA','XI DKV 4','02 Okt 2026 10:06',
'https://www.blogger.com/blog/post/edit/9148967484819656515/7200518875422246523','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI',[0,0,0,0,0,0,0,0,0,0,0,0],[
  ('TIDAK DAPAT DIVERIFIKASI','Link yang dikirim berupa URL editor Blogger (https://www.blogger.com/blog/post/edit/...), bukan artikel publik; halaman hanya dapat dibuka setelah login.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
  ('TIDAK DAPAT DIVERIFIKASI','URL editor Blogger - karya tidak dapat diakses publik.'),
]))

# ===================== NILAI TAMBAHAN RESOLUSI (Revisi Rubrik ke-3) =====================
# Resolusi TIDAK menjadi potongan nilai. Resolusi > 1000 px mendapat NILAI TAMBAHAN, maksimal +1,0.
RESOLUSI = {   # nama: (jumlah_gambar, jumlah_ge_1000px, lebar_maks, catatan)
 "WAHDAN SAPARI": (7,0,400,""),
 "AHMAD FAUZI": (11,0,320,""),
 "QIANDRA KAIZAR NAHARI": (5,4,1024,""),
 "ILMA LATIFAH": (5,0,320,""),
 "SOPA ANIDATUL AISAH": (7,0,320,''),
 "JAJANG M HUSNI MUBAROK": (5,5,1536,""),
 "DEDE APRILIA KARTIKA": (7,0,320,""),
 "PUTRI INTAN NURAENI": (5,0,320,''),
 "INTAN WIDIYANTI": (5,0,320,''),
 "SAFINAH SYARA GARINI": (8,0,450,""),
 "QUINSYA RAHMANESA SOLEHA": (7,0,320,""),
 "JIHAN SHAFIRA KEAN PUTRI MULYADI": (8,8,1376,""),
 "DHEA EKA KHOERUNNISA": (5,0,320,""),
 "WULAN SUNDARI": (6,6,1376,''),
 "WILDA AZKIA": (10,0,578,""),
 "MUHAMAD DIAZ PIRDAUS": (8,8,1376,""),
 "RIZKY MUHAMMAD REGAL SAFARI": (5,0,320,""),
  "KHANZA NURAENI": (5,4,1376,""),
 # Kiriman baru 01 Okt 2026
 "CEISHA SINTHIA": (0,0,0,""),
 "SELVI SIFA URIZQI": (5,0,320,""),
 "AZMI ANUGRAH": (5,0,320,""),
 "INDRI FITRIYANI": (5,0,400,""),
 "MEISYA FAKHRIYAH": (5,0,320,""),
 "M REZA HUAFAH": (27,0,320,""),
 "SYIVA WIDIYANA AGUSTIN": (8,0,320,""),
 "MEYLAN MELIYANTI ANASTASYA SOFYAN": (5,0,320,""),
 "PUTRI UTAMI": (8,0,320,""),
 "KAMILA APRILIANI": (7,0,320,""),
 "YAYU ASTIA": (6,0,320,""),
 "WINA AFRILIANI": (6,0,400,""),
 "JAJANG NURJAMAN": (5,5,1024,""),
  # Kiriman baru 1 Okt 2026 (pagi)
  "SHANDIKA REVI SHAFARUDIN": (8,0,640,""),
  "AI TITO": (7,0,320,""),
  "MERLIN AZNIKA": (5,0,320,""),
  "NENG SRI RAHAYU": (7,0,320,""),
  "SILVI BUDIA PUTRI": (6,0,320,""),
 "NURJIHAN": (6,0,320,''),
"SINDIA SAPUTRI": (0,0,0,"Inaccessible"),
   "MUTIA ANITA SARI": (7,0,320,""),
  "HARUM NURAULIA SRI KAMILA": (7,0,320,""),
  "AZZAHRA QYASIMAH": (8,0,320,""),
  "AUPA AZNIA": (6,0,320,""),
  "ALYA NURAENI": (6,0,320,""),
  "SITI JENAB": (0,0,0,"Inaccessible"),
  "WILDAN": (5,0,320,""),
 "DEBI LESTARI": (5,0,320,''),
  "HASNI SAPA AL MAIRA": (7,0,320,""),
  "ILFA ALIFIANA KHOERUNISA": (5,0,320,""),
 "DIRA RAHMAWATI": (5,0,320,''),
  "NAZMA KAYVA GASANI": (8,0,320,""),
  "RIANA SANJAYA": (10,0,320,""),
  "NAZWA NUR AISYAH": (6,0,320,""),
  "AI CINTA LESTARI": (8,0,320,""),
  "TIRA FADILA": (6,0,320,""),
  "AI NURAWALIAH AL ZAHRA": (6,0,320,""),
 "AI IMAS": (14,0,320,''),
  "SYIFA HAIRA": (6,0,320,""),
  "RISMAYANTI": (5,0,320,""),
  "MOH PIKRI": (6,0,320,""),
  "SUMIYATI": (8,0,320,""),
  "AJENG DWI RAISSA FITRI": (7,0,320,""),
 "SRI AYU WAHYUNI": (6,0,320,''),
  "RISMA SAPARANI": (0,0,0,"Inaccessible"),
  "ZAHRATUL AYESA AULIA": (6,0,320,""),
  "GADNA WIDIATNI": (6,0,320,""),
  "DEVINA NAYYRA FITRIANI": (7,0,320,""),
  "SULISTIAWATI": (7,0,320,""),
  "SECHAN KHALIFATUNNISA": (7,0,320,""),
 # 2 Oktober 2026
 "FITRIYANI": (6,0,320,''),
 "SYABINA AYAT EL AKHROS": (8,0,640,''),
 "SITI RAHMA SILPIANA": (8,0,320,''),
 "TENI DAMAYANTI": (8,0,320,''),
 "SITI KHOIRIYAH": (11,0,320,''),
 "AQILA NAZIL FALAQ": (0,0,0,'7 gambar di lh3.google.com (Google Photos) gagal diunduh; dimensi asli & OCR tidak terukur'),
 "RESTI NURUL FADILA": (8,0,446,''),
 "NENG OKTAVIA PUTRI AGUSTIN": (10,0,768,''),
 "FAHMI AHMAD RAMDAN ALFIAN": (6,0,450,''),
 "NADIRA MEGA RIZKIA": (7,0,440,''),
 "RITA PEBRIYANI": (7,0,429,''),
 "AI SITI MUSLIMAH": (6,0,400,''),
 "NANI YULIYANI": (7,0,320,''),
 "AMIRA NUR AULIA": (10,0,400,''),
 "RAIRA PUTRI": (8,0,320,''),
 "KAILA AROPATILAH": (5,0,320,''),
  # Link tidak dapat diakses (URL editor Blogger / HTTP 404) - tidak ada gambar terukur
 "NURI MEITRI AENI": (0,0,0,"Inaccessible"),
 "NAZWA KURNIA": (0,0,0,"Inaccessible"),
}
def fmt_bonus(b):
    """Tampilkan nilai tambah resolusi dengan 1 desimal, Pemisah koma (id-ID)."""
    if not b: return "-"
    return ("+%.1f" % b).replace(".", ",")

def bonus_resolusi(nama):
    """Nilai tambah resolusi. Maksimal +1,0 poin (Revisi Rubrik ke-3).
    Resolusi rendah TIDAK menjadi pengurangan nilai."""
    n, g, mx, _ = RESOLUSI.get(nama, (0,0,0,""))
    if n == 0: return 0, "Tidak ada gambar yang dapat diukur (link tidak dapat diakses)"
    if g == n: return 1.0, "Seluruh %d gambar >= 1000 px (maks %d px) - nilai tambah penuh +1,0" % (n, mx)
    if g / n >= 0.70: return 0.7, "%d dari %d gambar >= 1000 px (maks %d px) - nilai tambah +0,7" % (g, n, mx)
    if g > 0: return 0.4, "%d dari %d gambar >= 1000 px (maks %d px) - nilai tambah minimal +0,4" % (g, n, mx)
    return 0, "Tidak ada gambar >= 1000 px (maks %d px). TIDAK ada potongan nilai - resolusi rendah tetap diperbolehkan." % mx

# ===================== HITUNG =====================
def kategori(n):
    if n >= 90: return "Sangat Baik"
    if n >= 80: return "Baik"
    if n >= 70: return "Cukup"
    return "Perlu Perbaikan"

for row in S:
    rid, nama, kls, ts, link, mk, sh, sc, ms, pr, dsk, skor12, det = row
    vals = [round(s / 4 * b, 2) for s, b in zip(skor12, BOBOT)]
    tot = round(sum(vals), 2)
    row_vals = (vals, tot, kategori(tot))

# ===================== STYLING =====================
H_FILL = PatternFill("solid", fgColor="1F4E79")
H_FONT = Font(bold=True, color="FFFFFF", size=10)
SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
TOT_FILL = PatternFill("solid", fgColor="FFF2CC")
WARN = PatternFill("solid", fgColor="FCE4D6")
BAD = PatternFill("solid", fgColor="FFC7CE")
OKF = PatternFill("solid", fgColor="E2EFDA")
thin = Side(style="thin", color="BFBFBF")
BORD = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)

wb = Workbook()

def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = H_FILL; cell.font = H_FONT; cell.alignment = CTR; cell.border = BORD

# ---------- SHEET 1: REKAP NILAI ----------
ws = wb.active; ws.title = "1. REKAP NILAI"
ws["A1"] = "REKAP PENILAIAN PROJEK ASTS - PERSONAL BRANDING CREATIVE CAMPAIGN (AI dalam Desain)"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "Mata Pelajaran: DKV - AI dalam Desain | Semester: 1 | TP: 2026/2027 | Sekolah: SMK Negeri 9 Garut"
ws["A3"] = "Batas akhir: Jumat, 2 Oktober 2026 23.59 WIB (seluruh kiriman tepat waktu; waktu tidak memengaruhi skor kualitas)"
ws["A3"].font = Font(italic=True, size=9)
ws["A4"] = "Sumber: rekapan Google Form (86 kiriman) + akses langsung ke Blogger + 555 berkas gambar (OCR & dimensi piksel asli)"
ws["A4"].font = Font(italic=True, size=9)

hdr = ["No","Nama Siswa","Kelas","Waktu Kirim","Link Blogger","Mockup","Shotlist","Storyboard","Mascot","Prompt","Jml Kata Deskripsi","Resolusi (px)","Nilai Rubrik /100","Nilai Tambah Resolusi","NILAI AKHIR /100","Kategori"]
ws.append([]); ws.append(hdr); style_header(ws, 6, len(hdr))
r = 7
for i, row in enumerate(sorted(S, key=lambda x: -x_rows(x)[4]), 1):
    rid, nama, kls, ts, link, mk, sh, sc, ms, pr, dsk, skor12, det = row
    vals, rubrik, bon, ketb, akhir, kat0, kat = x_rows(row)
    ng, g1000, mx, _ = RESOLUSI.get(nama, (0,0,0,""))
    ws.cell(row=r, column=1, value=i)
    ws.cell(row=r, column=2, value=nama)
    ws.cell(row=r, column=3, value=kls)
    ws.cell(row=r, column=4, value=ts)
    ws.cell(row=r, column=5, value=link)
    ws.cell(row=r, column=6, value=mk)
    ws.cell(row=r, column=7, value=sh)
    ws.cell(row=r, column=8, value=sc)
    ws.cell(row=r, column=9, value=ms)
    ws.cell(row=r, column=10, value=pr)
    ws.cell(row=r, column=11, value=dsk)
    cres = ws.cell(row=r, column=12, value=("%d dari %d gambar >=1000px (maks %d)" % (g1000, ng, mx)) if ng else "tidak diukur")
    cres.font = Font(bold=True, color="1F4E79")
    c = ws.cell(row=r, column=13, value=rubrik); c.number_format = "0.00"; c.font = Font(bold=True)
    cb = ws.cell(row=r, column=14, value=fmt_bonus(bon))
    cb.alignment = CTR
    if bon: cb.font = Font(bold=True, color="006100"); cb.fill = OKF
    c = ws.cell(row=r, column=15, value=akhir); c.font = Font(bold=True, size=11)
    c.number_format = "0.00"; c.fill = TOT_FILL
    ws.cell(row=r, column=16, value=kat)
    for cc in range(1, len(hdr) + 1):
        cell = ws.cell(row=r, column=cc); cell.border = BORD
        cell.alignment = WRAP if cc in (5, 12) else CTR
        if isinstance(cell.value, str):
            if "BELUM" in cell.value or "TIDAK DI" in cell.value or "Tidak ditemukan" in cell.value:
                cell.fill = BAD
            elif "KURANG" in cell.value or "TIDAK DAPAT" in cell.value:
                cell.fill = WARN
    if kat == "Sangat Baik": ws.cell(row=r, column=16).fill = OKF
    r += 1

tot_all = round(sum(x_rows(x)[4] for x in S) / len(S), 2)
rub_all = round(sum(x_rows(x)[1] for x in S) / len(S), 2)
bon_all = round(sum(x_rows(x)[2] for x in S) / len(S), 2)
ws.cell(row=r, column=2, value="RATA-RATA KELAS").font = Font(bold=True)
c = ws.cell(row=r, column=13, value=rub_all); c.font = Font(bold=True); c.number_format = "0.00"
c = ws.cell(row=r, column=14, value="+%0.2f" % bon_all); c.font = Font(bold=True); c.number_format = "0.00"
c = ws.cell(row=r, column=15, value=tot_all); c.font = Font(bold=True, size=12); c.number_format = "0.00"
for cc in range(1, len(hdr) + 1):
    ws.cell(row=r, column=cc).border = BORD
    ws.cell(row=r, column=cc).fill = TOT_FILL

for col, w in zip("ABCDEFGHIJKLMNOP", [5,34,10,18,52,24,24,26,26,10,20,26,12,10,13,16]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "B7"

# ---------- SHEET 2: NILAI KOMPONEN ----------
ws2 = wb.create_sheet("2. NILAI KOMPONEN")
ws2["A1"] = "REKAP SKOR PER KOMPONEN (skor 0-4 ; nilai = skor/4 x bobot)"
ws2["A1"].font = Font(bold=True, size=13)
ws2.append([]); ws2.append([])
hdr2 = ["No","Nama","Kelas"] + ["%s (bobot %d)" % (KOM[i], BOBOT[i]) for i in range(12)] + ["Nilai Rubrik","Nilai Tambah Resolusi","NILAI AKHIR","Kategori"]
ws2.append(hdr2); style_header(ws2, 4, len(hdr2))
r = 5
for i, row in enumerate(sorted(S, key=lambda x: -x_rows(x)[4]), 1):
    rid, nama, kls, ts, link, mk, sh, sc, ms, pr, dsk, skor12, det = row
    vals, rubrik, bon, ketb, akhir, kat0, kat = x_rows(row)
    ws2.cell(row=r, column=1, value=i)
    ws2.cell(row=r, column=2, value=nama)
    ws2.cell(row=r, column=3, value=kls)
    for j, s in enumerate(skor12):
        cc = ws2.cell(row=r, column=4 + j, value=s); cc.alignment = CTR
        if s == 0: cc.fill = BAD
        elif s <= 1: cc.fill = WARN
        elif s == 4: cc.fill = OKF
    c = ws2.cell(row=r, column=16, value=rubrik); c.font = Font(bold=True); c.number_format = "0.00"
    cb = ws2.cell(row=r, column=17, value=fmt_bonus(bon)); cb.alignment = CTR
    if bon: cb.font = Font(bold=True, color="006100"); cb.fill = OKF
    c = ws2.cell(row=r, column=18, value=akhir); c.font = Font(bold=True); c.number_format = "0.00"; c.fill = TOT_FILL
    ws2.cell(row=r, column=19, value=kat).alignment = CTR
    for cc in range(1, 20):
        ws2.cell(row=r, column=cc).border = BORD
    r += 1
ws2.append([])
ws2.cell(row=r, column=2, value="Bobot").font = Font(bold=True)
for j, b in enumerate(BOBOT):
    ws2.cell(row=r, column=4 + j, value=b).font = Font(bold=True)
    ws2.cell(row=r, column=4 + j).alignment = CTR
ws2.column_dimensions["A"].width = 5
ws2.column_dimensions["B"].width = 34
ws2.column_dimensions["C"].width = 10
for j in range(12):
    ws2.column_dimensions[get_column_letter(4 + j)].width = 14
for col2, w2 in zip(["P","Q","R","S"], [12,10,12,16]):
    ws2.column_dimensions[col2].width = w2
ws2.freeze_panes = "D5"

# ---------- SHEET 3-21: DETAIL PER SISWA ----------
STATUS_FILL = {"TIDAK DIKUMPULKAN": BAD, "TIDAK DAPAT DIVERIFIKASI": WARN,
               "BELUM MEMENUHI": WARN, "SEBAGIAN": WARN, "LENGKAP": OKF, "DIKUMPULKAN": PatternFill("solid", fgColor="FFF2CC")}
for row in S:
    rid, nama, kls, ts, link, mk, sh, sc, ms, pr, dsk, skor12, det = row
    vals, rubrik, bon, ketb, akhir, kat0, kat = x_rows(row)
    wsx = wb.create_sheet(("Detail - " + nama.title())[:31])
    wsx["A1"] = "PENILAIAN DETAIL - %s (%s)" % (nama, kls)
    wsx["A1"].font = Font(bold=True, size=13)
    wsx["A2"] = "Link Blogger: %s" % link
    wsx["A2"].font = Font(size=9, color="0563C1")
    wsx["A3"] = "Waktu kirim: %s | Blogger: AKTIF | Rata-rata kata deskripsi: %s" % (ts, dsk)
    wsx["A3"].font = Font(italic=True, size=9)
    wsx.append([])
    wsx.append(["No","Komponen","Ketentuan","Status","Bukti / Hasil Pemeriksaan","Skor","Bobot","Nilai"])
    style_header(wsx, 5, 8)
    rr = 6
    for j in range(12):
        st, bk = det[j]
        wsx.cell(row=rr, column=1, value=j + 1)
        wsx.cell(row=rr, column=2, value=KOM[j])
        wsx.cell(row=rr, column=3, value=BOBOT[j]).alignment = CTR
        cst = wsx.cell(row=rr, column=4, value=st); cst.alignment = CTR
        cst.fill = STATUS_FILL.get(st, PatternFill())
        wsx.cell(row=rr, column=5, value=bk).alignment = WRAP
        cs = wsx.cell(row=rr, column=6, value=skor12[j]); cs.alignment = CTR
        cs.fill = BAD if skor12[j] == 0 else (WARN if skor12[j] <= 1 else (OKF if skor12[j] == 4 else PatternFill()))
        cn = wsx.cell(row=rr, column=7, value=vals[j]); cn.alignment = CTR; cn.number_format = "0.00"
        wsx.cell(row=rr, column=8, value="(Skor/4)x%d = %.2f" % (BOBOT[j], vals[j])).alignment = CTR
        for cc in range(1, 9):
            wsx.cell(row=rr, column=cc).border = BORD
            if cc in (2, 3): wsx.cell(row=rr, column=cc).alignment = WRAP
        rr += 1
    ng, g1000, mx, _ = RESOLUSI.get(nama, (0,0,0,""))
    wsx.cell(row=rr + 1, column=1, value="+")
    wsx.cell(row=rr + 1, column=2, value="NILAI TAMBAHAN RESOLUSI (>1000 px)").font = Font(bold=True)
    wsx.cell(row=rr + 1, column=3, value="").alignment = CTR
    wsx.cell(row=rr + 1, column=4, value=fmt_bonus(bon)).alignment = CTR
    if bon: wsx.cell(row=rr + 1, column=4).fill = OKF
    wsx.cell(row=rr + 1, column=5, value=ketb).alignment = WRAP
    wsx.cell(row=rr + 1, column=6, value=bon).alignment = CTR
    wsx.cell(row=rr + 1, column=8, value="Resolusi tinggi = nilai tambahan").alignment = CTR
    for cc in range(1, 9):
        wsx.cell(row=rr + 1, column=cc).border = BORD
    wsx.cell(row=rr + 2, column=2, value="NILAI RUBRIK").font = Font(bold=True)
    c = wsx.cell(row=rr + 2, column=6, value=rubrik); c.font = Font(bold=True); c.number_format = "0.00"; c.alignment = CTR
    wsx.cell(row=rr + 3, column=2, value="NILAI AKHIR").font = Font(bold=True, size=12)
    c = wsx.cell(row=rr + 3, column=6, value=akhir); c.font = Font(bold=True, size=12)
    c.number_format = "0.00"; c.fill = TOT_FILL; c.alignment = CTR
    wsx.cell(row=rr + 3, column=8, value="KATEGORI: " + kat).font = Font(bold=True)
    wsx.column_dimensions["A"].width = 4
    wsx.column_dimensions["B"].width = 30
    wsx.column_dimensions["C"].width = 9
    wsx.column_dimensions["D"].width = 24
    wsx.column_dimensions["E"].width = 95
    wsx.column_dimensions["F"].width = 7
    wsx.column_dimensions["G"].width = 8
    wsx.column_dimensions["H"].width = 20

# ---------- SHEET: REKAP KETIDAKLENGKAPAN ----------
ws3 = wb.create_sheet("3. REKAP KETIDAKLENGKAPAN")
ws3["A1"] = "REKAP KETIDAKLENGKAPAN / PELANGGARAN KETENTUAN SOAL"
ws3["A1"].font = Font(bold=True, size=13)
ws3.append([]); ws3.append([])
ws3.append(["No","Nama","Kelas","Komponen","Ketentuan Soal","Hasil Aktual","Kekurangan / Catatan"])
style_header(ws3, 4, 7)
rr = 5
n = 0
KET = {
 0: "Judul pendek & panjang sesuai format",
 1: "Logo, nama branding, tagline, NAMA SISWA pada branding, deskripsi >=100 kata",
 2: "Moodboard landscape + 7 unsur",
 3: "Minimal 3 mockup + fungsi media",
 4: "Judul, Tema, Pesan Utama, Narasi/Dialog, Closing Tagline",
 5: "Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup",
 6: "Minimal 10 shot (boleh tabel ATAU gambar) + kolom lengkap",
 7: "Minimal 6 scene + keterangan lengkap",
 8: "WAJIB: Mascot Full Body. OPSIONAL: Mascot Portrait & Mascot Bersama Logo",
 9: "Prompt pada setiap visual",
 10: "10 komponen dalam artikel + 5 label + link aktif",
 11: "Konsistensi identitas & profesionalitas",
}
for row in sorted(S, key=lambda x: -x_rows(x)[1]):
    rid, nama, kls, ts, link, mk, sh, sc, ms, pr, dsk, skor12, det = row
    for j in range(12):
        st, bk = det[j]
        if st in ("TIDAK DIKUMPULKAN", "TIDAK DAPAT DIVERIFIKASI", "BELUM MEMENUHI", "SEBAGIAN", "DIKUMPULKAN"):
            n += 1
            ws3.cell(row=rr, column=1, value=n)
            ws3.cell(row=rr, column=2, value=nama)
            ws3.cell(row=rr, column=3, value=kls)
            ws3.cell(row=rr, column=4, value=KOM[j])
            ws3.cell(row=rr, column=5, value=KET[j]).alignment = WRAP
            ws3.cell(row=rr, column=6, value=st).alignment = CTR
            ws3.cell(row=rr, column=6).fill = STATUS_FILL.get(st, PatternFill())
            ws3.cell(row=rr, column=7, value=bk).alignment = WRAP
            for cc in range(1, 8):
                ws3.cell(row=rr, column=cc).border = BORD
            ws3.cell(row=rr, column=2).alignment = WRAP
            rr += 1
for col, w in zip("ABCDEFG", [5, 32, 10, 28, 40, 24, 105]):
    ws3.column_dimensions[col].width = w
ws3.freeze_panes = "D5"

# ---------- SHEET: VERIFIKASI MANUAL GURU ----------
ws4 = wb.create_sheet("4. CEK MANUAL GURU")
ws4["A1"] = "BUTIR YANG WAJIB DICEK MANUAL OLEH GURU"
ws4["A1"].font = Font(bold=True, size=13)
ws4["A2"] = ("Ases tidak dapat melihat gambar secara visual. Verifikasi memakai OCR + dimensi piksel asli. "
             "Butir berikut tidak dapat dipastikan dan perlu dibuka langsung oleh guru.")
ws4["A2"].font = Font(italic=True, size=9, color="C00000")
ws4["A2"].alignment = WRAP
ws4.append([]); ws4.append([])
ws4.append(["No","Nama","Butir yang perlu dicek","Petunjuk OCR / bukti", "Cek (x)", "Jumlah aktual", "Catatan guru"])
style_header(ws4, 5, 7)
CEK = [
 ("QUINSYA RAHMANESA SOLEHA","Jumlah shot pada tabel shotlist","Tabel memuat nomor 1,2,3,4,5,6,8,9 = 8 shot (kurang 2). Format tabel tetap wajib, jumlah harus >=10"),
 ("JIHAN SHAFIRA KEAN PUTRI MULYADI","Jumlah scene storyboard","Berkas memuat Scene 1-4 (kurang 2 dari minimum 6)"),
 ("QIANDRA KAIZAR NAHARI","Ada/tidaknya gambar AI Mascot","Tidak ada berkas maskot sama sekali di 5 gambar artikel"),
 ("AHMAD FAUZI","Fungsi 6 gambar antara storyline dan storyboard","Berada di posisi D.3 (Shotlist) - belum jelas shotlist atau frame storyboard"),
 ("AHMAD FAUZI","Jumlah shot pada shotlist gambar","Format gambar sudah DITERIMA; yang belum dipastikan hanya jumlah baris (teks hasil AI tidak terbaca)"),
 ("WAHDAN SAPARI","Jumlah scene storyboard","OCR hanya membaca 5 label scene"),
 ("WAHDAN SAPARI","Cek visual mascot full body","1 gambar mascot, OCR terbaca 'FULL BODY' - sudah dianggap memenuhi"),
 ("SOPA ANIDATUL AISAH","Jumlah scene storyboard & apakah ada mascot full body","Gambar storyboard 320x292; tidak ada berkas mascot terpisah sama sekali"),
 ("DHEA EKA KHOERUNNISA","Apakah gambar mascot = FULL BODY & jumlah mockup","Teks menyebut 3 mockup; hanya 1 berkas. Full body kini komponen WAJIB"),
 ("PUTRI INTAN NURAENI","Jumlah mockup dalam 1 gambar","Teks menyebut laptop, stiker, kartu nama; 1 berkas gambar"),
 ("INTAN WIDIYANTI","Jumlah mockup, scene storyboard, & apakah mascot full body","Masing-masing hanya 1 berkas gambar"),
 ("ILMA LATIFAH","Jumlah mockup dalam 1 gambar & apakah mascot full body","Teks menyebut kartu nama, stiker, banner; 1 berkas gambar"),
 ("DEDE APRILIA KARTIKA","Jumlah mockup dalam 1 gambar","Teks menyebut tote bag, kartu nama, social media feed; 1 berkas gambar"),
 ("SAFINAH SYARA GARINI","Jumlah baris shotlist pada gambar","Format gambar sudah DITERIMA; yang belum dipastikan hanya jumlah baris (header terbaca, baris tidak)"),
 ("JIHAN SHAFIRA KEAN PUTRI MULYADI","Jumlah mockup dalam berkas gambar","4-5 berkas identitas; fungsi laptop/packaging/gelas terverifikasi di teks"),
 ("WULAN SUNDARI","Jumlah mockup dalam 1 gambar","1 berkas 1376x768"),
 ("WILDA AZKIA","Apakah salah satu gambar mascot = FULL BODY","Tidak ada label pada gambar; full body kini komponen WAJIB"),
 ("MUHAMAD DIAZ PIRDAUS","Orisinalitas heading '{Monogram SSG}'","Heading yang sama persis muncul pada 4 kiriman: MUHAMAD DIAZ PIRDAUS, NAZMA KAYVA GASANI, TENI DAMAYANTI, RESTI NURUL FADILA. Perlu klarifikasi orisinalitas dari siswa."),
 ("NAZWA KAYVA GASANI","Orisinalitas heading '{Monogram SSG}'","Lihat butir MUHAMAD DIAZ PIRDAUS - heading identik pada 4 kiriman."),
 ("TENI DAMAYANTI","Orisinalitas heading '{Monogram SSG}'","Lihat butir MUHAMAD DIAZ PIRDAUS - heading identik pada 4 kiriman."),
 ("RESTI NURUL FADILA","Orisinalitas heading '{Monogram SSG}'","Lihat butir MUHAMAD DIAZ PIRDAUS - heading identik pada 4 kiriman."),
 ("CEISHA SINTHIA","Kirim link artikel publik","Link yang masuk adalah URL editor Blogger (redirect login Google). Nilai 0 karena karya tidak dapat diverifikasi."),
 ("SINDIA SAPUTRI","Kirim link artikel publik","Link yang masuk adalah URL editor Blogger (redirect login Google). Nilai 0 karena karya tidak dapat diverifikasi."),
 ("SITI JENAB","Kirim link artikel publik","Link yang masuk adalah URL editor Blogger (redirect login Google). Nilai 0 karena karya tidak dapat diverifikasi."),
 ("NAZWA KURNIA","Kirim link artikel publik","Link yang masuk adalah URL editor Blogger (redirect login Google). Nilai 0 karena karya tidak dapat diverifikasi."),
 ("RISMA SAPARANI","Kirim link artikel publik yang aktif","Link mengembalikan HTTP 404. Nilai 0 karena karya tidak dapat diverifikasi."),
 ("NURI MEITRI AENI","Kirim link artikel publik yang aktif","Link mengembalikan HTTP 404. Nilai 0 karena karya tidak dapat diverifikasi."),
]
rr = 6
for i, (nama, butir, petunjuk) in enumerate(CEK, 1):
    ws4.cell(row=rr, column=1, value=i)
    ws4.cell(row=rr, column=2, value=nama)
    ws4.cell(row=rr, column=3, value=butir)
    ws4.cell(row=rr, column=4, value=petunjuk).alignment = WRAP
    ws4.cell(row=rr, column=5, value="").alignment = CTR
    ws4.cell(row=rr, column=6, value="").alignment = CTR
    ws4.cell(row=rr, column=7, value="").alignment = WRAP
    for cc in range(1, 8):
        ws4.cell(row=rr, column=cc).border = BORD
        if cc in (2, 3): ws4.cell(row=rr, column=cc).alignment = WRAP
    rr += 1
for col, w in zip("ABCDEFG", [5, 32, 44, 62, 7, 14, 30]):
    ws4.column_dimensions[col].width = w

# ---------- SHEET: DATA REKAPAN ----------
ws5 = wb.create_sheet("5. DATA REKAPAN")
ws5["A1"] = "DATA REKAPAN PENGUMPULAN (Google Form Responses 1)"
ws5["A1"].font = Font(bold=True, size=13)
ws5.append([]); ws5.append([])
ws5.append(["No","Timestamp","Kelas","Nama","Link Postingan Blogger","Status Blogger","Catatan"])
style_header(ws5, 4, 7)
rr = 5
# Status akses blogger berdasarkan hasil fetch (parsed.json). Id R## tidak dipakai
# sebagai kunci karena tidak unik antar-entri.
TIDAK_AKSES = {
  "CEISHA SINTHIA":  "TIDAK AKSES - URL editor Blogger (redirect login Google)",
  "SINDIA SAPUTRI":  "TIDAK AKSES - URL editor Blogger (redirect login Google)",
  "SITI JENAB":      "TIDAK AKSES - URL editor Blogger (redirect login Google)",
  "NAZWA KURNIA":    "TIDAK AKSES - URL editor Blogger (redirect login Google)",
  "RISMA SAPARANI":  "TIDAK AKSES - HTTP 404",
  "NURI MEITRI AENI":"TIDAK AKSES - HTTP 404",
}
CATATAN = {
  "AHMAD FAUZI": "Judul artikel 'Biodata diri' - perlu verifikasi identitas",
  "MUHAMAD DIAZ PIRDAUS": "Heading '{Monogram SSG}' - teks yang sama juga muncul di 3 kiriman lain",
  "NAZMA KAYVA GASANI": "Heading '{Monogram SSG}' - teks yang sama juga muncul di 3 kiriman lain",
  "TENI DAMAYANTI": "Heading '{Monogram SSG}' - teks yang sama juga muncul di 3 kiriman lain",
  "RESTI NURUL FADILA": "Heading '{Monogram SSG}' - teks yang sama juga muncul di 3 kiriman lain",
  "AZMI ANUGRAH": "Judul artikel 'UJI KOPETENSI PROMTPTING AI DKV' - perlu verifikasi identitas",
  "PUTRI INTAN NURAENI": "Judul artikel 'ASTS KOMPETENSI AI DKV - SMKN 9 GARUT' - perlu verifikasi identitas",
  "MUTIA ANITA SARI": "Page title 'ASTS personal branding nama Mutia' - nama pada judul tidak lengkap",
}
for i, row in enumerate(S, 1):
    rid, nama, kls, ts, link, mk, sh, sc, ms, pr, dsk, skor12, det = row
    vals, rubrik, bon, ketb, akhir, kat0, kat = x_rows(row)
    status = TIDAK_AKSES.get(nama, "AKTIF (HTTP 200)")
    cat_ = CATATAN.get(nama, "")
    ws5.cell(row=rr, column=1, value=i)
    ws5.cell(row=rr, column=2, value=ts)
    ws5.cell(row=rr, column=3, value=kls)
    ws5.cell(row=rr, column=4, value=nama)
    ws5.cell(row=rr, column=5, value=link).alignment = WRAP
    c = ws5.cell(row=rr, column=6, value=status)
    if nama in TIDAK_AKSES: c.fill = BAD; c.font = Font(bold=True, color="9C0006")
    ws5.cell(row=rr, column=7, value=cat_).alignment = WRAP
    for cc in range(1, 8):
        ws5.cell(row=rr, column=cc).border = BORD
    rr += 1
for col, w in zip("ABCDEFG", [5, 20, 10, 32, 62, 38, 44]):
    ws5.column_dimensions[col].width = w

# ---------- SHEET: NILAI TAMBAHAN RESOLUSI ----------
wsB = wb.create_sheet("7. NILAI TAMBAHAN RESOLUSI")
wsB["A1"] = "NILAI TAMBAHAN RESOLUSI (> 1000 px)"
wsB["A1"].font = Font(bold=True, size=13)
wsB["A2"] = ("REVISI RUBRIK KE-3: nilai tambah resolusi DITURUNKAN agar pengaruhnya kecil. "
            "Resolusi TIDAK menjadi potongan nilai - resolusi tetap diperbolehkan. "
            "Visual beresolusi > 1000 px hanya mendapat NILAI TAMBAHAN, maksimal +1,0 poin.")
wsB["A2"].font = Font(italic=True, size=9, color="006100")
wsB["A2"].alignment = WRAP
wsB.append([]); wsB.append([])
wsB.append(["No","Nama","Jumlah Gambar","Gambar >=1000 px","Lebar Maks (px)","Nilai Tambah","Keterangan"])
style_header(wsB, 5, 7)
rr = 6
for i, row in enumerate(sorted(S, key=lambda x: -x_rows(x)[4]), 1):
    rid, nama, kls, ts, link, mk, sh, sc, ms, pr, dsk, skor12, det = row
    vals, rubrik, bon, ketb, akhir, kat0, kat = x_rows(row)
    ng, g1000, mx, _ = RESOLUSI.get(nama, (0,0,0,""))
    wsB.cell(row=rr, column=1, value=i)
    wsB.cell(row=rr, column=2, value=nama).alignment = WRAP
    wsB.cell(row=rr, column=3, value=ng).alignment = CTR
    wsB.cell(row=rr, column=4, value=g1000).alignment = CTR
    wsB.cell(row=rr, column=5, value=(mx if ng else "-")).alignment = CTR
    cb = wsB.cell(row=rr, column=6, value=fmt_bonus(bon)); cb.alignment = CTR
    if bon: cb.font = Font(bold=True, color="006100"); cb.fill = OKF
    wsB.cell(row=rr, column=7, value=ketb).alignment = WRAP
    for cc in range(1, 8):
        wsB.cell(row=rr, column=cc).border = BORD
    rr += 1
rr += 1
for ket in ["ATURAN NILAI TAMBAHAN RESOLUSI (maksimal +1,0 poin):",
            "+1,0 poin = seluruh gambar >= 1000 px",
            "+0,7 poin = >= 70% gambar >= 1000 px",
            "+0,4 poin = sebagian gambar >= 1000 px",
            "+0 poin  = tidak ada gambar >= 1000 px (TIDAK ada pencilan nilai)",
            "Nilai akhir = Nilai Rubrik + Nilai Tambah Resolusi, maksimum 100",
            "Resolusi diukur dari berkas ASLI yang diunduh (parameter =s0), bukan ukuran tampil di Blogger."]:
    c = wsB.cell(row=rr, column=1, value=ket)
    if ":" in ket or ket.startswith("ATURAN"): c.font = Font(bold=True)
    rr += 1
for col, w in zip("ABCDEFG", [5, 32, 16, 18, 16, 9, 78]):
    wsB.column_dimensions[col].width = w

# ---------- SHEET: METODE ----------
ws6 = wb.create_sheet("6. METODE & RUBRIK")
ws6["A1"] = "CATATAN METODE, KETERBATASAN, DAN RUBRIK"; ws6["A1"].font = Font(bold=True, size=13)
ws6.append([])
notes = [
 ("REVISI RUBRIK KE-3 (1 Oktober 2026)", ""),
 ("Nilai tambah resolusi", "Diturunkan agar pengaruhnya kecil: +1,0 / +0,7 / +0,4 / 0 (maksimal +1,0 poin, bukan +3)."),
 ("Resolusi rendah", "TIDAK menjadi potongan nilai. Resolusi tetap diperbolehkan."),
 ("Nilai Akhir", "Nilai Rubrik (0-100) + Nilai Tambah Resolusi, dibatasi maksimum 100."),
 ("Dampak", "18 siswa dihitung ulang saat revisi ini diterapkan (1 Oktober 2026). Tidak ada nilai turun; selisih hanya pada 6 siswa yang mendapat nilai tambah."),
 ("", ""),
 ("REVISI RUBRIK KE-2 (1 Oktober 2026)", ""),
 ("Resolusi rendah", "TIDAK lagi menjadi potongan nilai. Resolusi tetap diperbolehkan."),
 ("Nilai tambah resolusi", "Diperkenalkan: visual > 1000 px mendapat nilai tambahan (skala +3/+2/+1, kemudian diturunkan pada Revisi-3)."),
 ("", ""),
 ("REVISI RUBRIK KE-1 (1 Oktober 2026)", ""),
 ("1. Written 'by Nama Siswa'", "TIDAK lagi syarat. Yang wajib: NAMA SISWA tercantum pada branding/logo."),
 ("2. AI Mascot Character", "WAJIB: Mascot Full Body. OPSIONAL: Mascot Portrait & Mascot Bersama Logo."),
 ("3. Shotlist", "Boleh dalam bentuk TABEL maupun GAMBAR. Minimal 10 shot tetap berlaku."),
 ("Dampak", "Nilai seluruh siswa DIHITUNG ULANG. Tidak ada siswa yang mengunggah ulang (0 perubahan data)."),
 ("", ""),
 ("METODE VERIFIKASI", ""),
 ("Sumber 1", "Rekapan pengumpulan Google Form: 86 kiriman (nama, kelas, link, timestamp)."),
 ("Sumber 2", "Halaman Blogger diakses langsung satu per satu. 80 dari 86 dapat diakses publik; 4 siswa mengirim URL editor Blogger (mengarah ke halaman login Google) dan 2 link mengembalikan HTTP 404."),
 ("Sumber 3", "555 berkas gambar terukur (dimensi piksel asli) dan diperiksa dengan OCR (Apple Vision). 6 siswa tidak memiliki gambar yang dapat diukur karena link tidak dapat diakses."),
 ("Sumber 4", "Teks artikel, tabel, heading, dan label/tag Blogger."),
 ("", ""),
 ("KETERBATASAN PENTING", ""),
 ("Vision", "Ases TIDAK memiliki kemampuan melihat gambar. Verifikasi visual memakai OCR + dimensi piksel asli."),
 ("Terverifikasi", "Jumlah baris tabel, jumlah kata, jumlah berkas gambar, orientasi gambar, teks dalam gambar, label/tag."),
 ("Tidak terverifikasi", "Jumlah unit di dalam satu gambar kolase, bila label/penomoran tidak terbaca OCR. Ditulis TIDAK DAPAT DIVERIFIKASI."),
 ("Aturan", "Tidak ada nilai diberikan berdasarkan asumsi. Komponen yang tidak ditemukan ditulis 'Tidak ditemukan pada hasil pengumpulan'."),
 ("", ""),
 ("ATURAN PENILAIAN", ""),
 ("Rumus", "Nilai komponen = (Skor / 4) x Bobot. Total = jumlah seluruh nilai komponen. Skala 0-4."),
 ("Kategori", "90-100 Sangat Baik | 80-89 Baik | 70-79 Cukup | <70 Perlu Perbaikan."),
 ("Waktu", "Seluruh 86 kiriman sebelum batas 2 Oktober 2026 23.59 WIB. Waktu TIDAK mengubah skor kualitas."),
 ("", ""),
 ("BOBOT KOMPONEN", ""),
]
for k, b in zip(KOM, BOBOT):
    notes.append((k, "bobot %d poin" % b))
notes += [
 ("", ""),
 ("TOTAL", "100 poin"),
 ("", ""),
 ("CATATAN PENTING UNTUK GURU", ""),
 ("Cek manual", "Lihat sheet '4. CEK MANUAL GURU' untuk butir yang tidak dapat dipastikan."),
 ("Data rekapan", "Rekapan Google Form berisi 86 kiriman dan seluruh 86 siswa dinilai. Tidak ada siswa yang dikeluarkan dari penilaian."),
 ("Nilai 0", "6 siswa bernilai 0 karena karya tidak dapat diverifikasi, bukan karena karya kosong: 4 siswa mengirim URL editor Blogger (CEISHA SINTHIA, SINDIA SAPUTRI, SITI JENAB, NAZWA KURNIA) dan 2 link mengembalikan HTTP 404 (RISMA SAPARANI, NURI MEITRI AENI)."),
]
rr = 3
for k, v in notes:
    if k and not v:
        ws6.cell(row=rr, column=1, value=k).font = Font(bold=True, size=11, color="1F4E79")
    elif k:
        ws6.cell(row=rr, column=1, value=k).font = Font(bold=True)
        ws6.cell(row=rr, column=2, value=v).alignment = WRAP
    rr += 1
ws6.column_dimensions["A"].width = 34
ws6.column_dimensions["B"].width = 110

wb.save(OUT)
print("OK ->", OUT)
print("Sheet:", len(wb.sheetnames))
for s in wb.sheetnames: print("  -", s)
print("Rata-rata kelas:", tot_all)
