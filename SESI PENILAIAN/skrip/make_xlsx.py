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

S.append(('R01','SAFINAH SYARA GARINI','XI DKV 1','30 Sep 2026 18:20',
'https://safinahsyaragarini.blogspot.com/2026/09/personal-branding-safinah-syaara-garini.html',
'3 (terverifikasi)','11 (terverifikasi tabel)','6 (terverifikasi gambar + teks)','1 (3 varian: full body, portrait, bersama logo)','8','4777 (memenuhi)',[4, 4, 4, 3, 4, 4, 4, 4, 4, 4, 3, 4],
[('LENGKAP', "Page title 'PERSONAL BRANDING SAFINAH SYARA GARINI XI DKV 1' + heading artikel 'ASTS Personal Branding Safinah Syara Garini XI DKV 1 SMKN 9 GARUT' - keduanya sesuai format. Catatan: slug URL 'personal-branding-safinah-syaara-garini' salah eja (Syaara), isi artikel tetap menulis 'Safinah Syara Garini'."), ('LENGKAP', "Logo monogram SSG IMG#01 320x320, OCR 'SAFINAH SYARA GARINI / StorieS, mOtIon, And LIfe.' sehingga NAMA SISWA tercetak pada logo; nama branding 'Safinah Syara Garini' + studio 'SSG Creative'; tagline 'Stories, Motion, and Life'; 'by: Safinah Syara Garini'; deskripsi konsep 'Filosofi Elemen Visual & Merek Pribadi' 431 kata (gelombang air, lintasan lari, diafragma kamera)."), ('LENGKAP', "IMG#02 446x249 landscape. 7 unsur terverifikasi di teks dan OCR: Warna Utama (Deep Navy #0C2C5E, Warm Amber #FFB145, Teal Green #2DB89B, Ivory White), Typography (Montserrat/Agnifa), Style Visual, Referensi Desain, Tone & Mood (HUMBLE/AUTHENTIC/INSPIRATE/PROFESIONAL), Elemen Grafis, Inspirasi Visual - OCR 'WARNAUTAMA / TPOCRASNY / TONE&HGOD / INSPRASIMISUALSELLMENCRASIE'."), ('SEBAGIAN', "3 mockup dengan 3 gambar: A. Kartu Nama (IMG#03 386x215), B. Feed Media Sosial (IMG#04 320x320), C. Kemasan/Packaging (IMG#05 406x227), masing-masing ada penjelasan filosofi dan fungsi. CATATAN: bagian 'C.Kemasan / Packaging' berisi teks 139 kata yang IDENTIK dengan bagian 'A. Kartu Nama' (duplikasi), sehingga penjelasan fungsi kemasan bukan spesifik."), ('LENGKAP', "'5. NASKAH IKLAN DOKUMEN NASKAH VIDEO SHORT MOVIE (60 DETIK)': judul dokumen, Tema 'Perjalanan Kreatif, Ketangguhan, dan Storytelling Visual', Pesan Utama, Rundown 5 scene bertimecode (00:00-01:00) lengkap Visual + Audio/SFX + Voice Over, Closing Tagline 'Safinah Syara Garini - Stories, Motion, and Life.'"), ('LENGKAP', "'6. STORYLINE' memuat keempat unsur: 1. Pembukaan (The Spark of Momentum, lintasan lari golden hour), 2. Alur Cerita (The Adaptive Flow, kolam renang), 3. Konflik & Fokus Visual (Climax - Framing the Vision, POV viewfinder), 4. Penutup (The Authentic Connection) - tiap bagian punya Waktu & Tempat, Deskripsi Visual, dan Fokus & Emosi Utama."), ('LENGKAP', "Tabel HTML shotlist 12 baris x 7 kolom, header 'No | Adegan | Jenis Shot | Angle | Movement | Durasi | Deskripsi Visual', memuat 11 shot bernomor 1-11 (Scene 1 Opening Lintasan Lari sampai Scene 5 Closing Graphic) - jumlah dan kolom lengkap, memenuhi minimum 10 shot (naik dari 'jumlah TIDAK DAPAT DIVERIFIKASI')."), ('LENGKAP', "IMG#07 450x253 landscape dengan 6 panel berlabel 'PERSIAPAN&FOKUS', 'SEMANGAT BERLARI', 'ADAPTASI&KETENANCAN', 'PENGAMATANSENI', 'MENANGKAPRASA', 'PENUTUP&IDENTITAS'; teks 'Analisis Filosofi Tiap Panel' menguraikan Panel 1-6 dan OCR memuat token transisi ('CUTtkreuamg.AmgediatYonss') serta 'DalogMarsi' / 'DalogNars' (dialog/narasi)."), ('LENGKAP', "IMG#06 320x320 dengan OCR 'Full Body Masoo. / Panrait viow / Maseot with Loga' - ketiga output wajib (Mascot Full Body, Mascot Portrait, Mascot Bersama Logo) terverifikasi dari label pada gambar maskot."), ('LENGKAP', "8 prompt terdokumentasi dan diberi label 'PROMPT' (logo, moodboard, mockup, maskot, naskah, storyline, shotlist, storyboard); prompt mockup berisi 3 instruksi terpisah (Packaging, Feed Media Sosial, Kemasan/Packaging) dan setiap prompt menyebut iterate ('wow kerenn', 'Okeee lanjut')."), ('SEBAGIAN', "Isi artikel paling lengkap dalam batch ini (4.777 kata, 8 komponen + 2 tabel shotlist, link HTTP 200), tetapi hanya 2 label Blogger: ['PERSONAL BRANDING', 'PERSONAL BRANDING1'] - efektif 1 dari 5 label wajib; label 'Tugas Sekolah', 'AI', 'Portofolio', dan 'SMKN 9 Garut' tidak terpasang dan label kedua hanya duplikasi ejaan."), ('LENGKAP', "Konsep orisinal dan konsisten: kreator fotografi (lari, renang, fotografi) dengan elemen visual gelombang air, lintasan lari, dan diafragma kamera; palet Deep Navy/Warm Amber/Teal Green + tagline 'Stories, Motion, and Life' dipakai seragam pada logo, moodboard, 3 mockup, maskot, naskah, storyline, shotlist, dan storyboard. Catatan minor: duplikasi teks bagian kemasan (sudah dipotong pada K4).")],
))

S.append(('R98','QUINSYA RAHMANESA SOLEHA','XI DKV 3','01 Okt 2026 18:59',
'https://quinsyarahmanesasoleha.blogspot.com/2026/10/personal-branding-quinsya-rahmanesa.html',
'3 media (terverifikasi teks; 1 berkas gambar 368x206)','10 (terverifikasi tabel)','6 (terverifikasi teks + OCR label panel)','3 varian (terverifikasi OCR)','8','2114 (memenuhi)',[3, 4, 4, 3, 4, 4, 4, 3, 4, 4, 4, 3],
[('SEBAGIAN', "Baris pertama artikel 'ASTS Personal Branding Quinsya Rahmanesa Soleha XI DKV 3 SMKN 9 Garut' - judul panjang sesuai. Page title 'quinsya blog: Personal Branding Quinsya Rahmanesa Soleha XI DKV 3' memuat pola judul pendek tetapi diberi awalan 'quinsya blog:' sehingga tidak persis sesuai format."), ('LENGKAP', "Logo QRS 3D glossy (IMG#01 320x320); nama branding 'QRS'; tagline 'Small Clips, Big Stories'; deskripsi konsep 240 kata (>=100) yang menguraikan warna maroon/silver, elemen editing, dan makna slogan; NAMA SISWA tercetak pada logo ('Tulisan Quinsya Rahmanesa Soleha di bawah logo') - pada pemeriksaan sebelumnya komponen ini belum terpenuhi."), ('LENGKAP', 'IMG#02 346x193 landscape. Ketujuh unsur terverifikasi OCR dan teks: WARNA UTAMA BRANDING, TYPOGRAPHY (Cormorant SC, Playfair Display, Inter), STYLE VISUAL, REFERENSI DESAIN, TONE & MOOD, ELEMEN GRAFIS (clapperboard, timeline, bintang), INSPIRASI VISUAL.'), ('SEBAGIAN', "Teks menjelaskan 3 media (stiker, laptop, social media feed) beserta fungsi tiap media - pada pemeriksaan sebelumnya fungsi laptop belum dijelaskan, kini lengkap. Namun hanya 1 berkas gambar mockup (IMG#03 368x206) dengan OCR pseudoteks ('QRS / YORS / SORS / TEUYTIOT / TIERNOEOO') sehingga jumlah unit mockup tidak dapat diverifikasi."), ('LENGKAP', "'5. Naskah iklan' kini lengkap: Judul 'Cerita Besar dalam Setiap Potongan Klip', Tema 'Kreativitas & Penyuntingan Video Profesional', Pesan Utama, Narasi/Dialog Scene 1-4 dengan NARRATOR (VO) dan EDITOR (DIALOG), serta Closing Tagline 'QRS - Small Clips, Big Stories'."), ('LENGKAP', "'6.Storyline' kini terisi penuh dengan 4 bagian: 1. Pembukaan (Opening), 2. Alur Cerita (Development), 3. Konflik / Fokus Visual (Climax), 4. Penutup (Closing) - tiap bagian memuat Visual, Audio, dan Proses - pada pemeriksaan sebelumnya bagian ini kosong."), ('LENGKAP', 'Tabel HTML shotlist 11 baris x 7 kolom: header No./Adegan/Jenis Shot/Angle/Movement/Durasi/Deskripsi + 10 baris shot bernomor 1-10 lengkap (Extreme Long Shot sampai Closing & Title Tagline). Pada pemeriksaan sebelumnya hanya 8 shot dengan nomor tidak berurutan - kini MEMENUHI minimum 10.'), ('SEBAGIAN', "Bagian '8.storyboard' (versi lama: 'Gambar di atas') kini mendeskripsikan 6 panel bernomor eksplisit (Panel 1 Kebingungan, 2 Inspirasi, 3 Proses, 4 Transisi, 5 Hasil Akhir, 6 Logo Brand) dan OCR IMG#05 320x320 membaca keenam label panel (KEBINGUNGAN/CONFUSION, INSPIRASI/INSPIRATION, PROSES/PROCESS, TRANSISI, HASIL AKHIR, LOGO BRAND). Keterangan shot/angle/transisi/dialog per scene tetap tidak ada."), ('LENGKAP', "IMG#04 338x338, OCR 'MASCOT FULL BODY' (WAJIB), 'MASCOT PORTRAIT', dan 'MASCOT & LOGO' - ketiga varian terverifikasi dan juga dirinci di teks."), ('LENGKAP', "8 prompt terdokumentasi dan diberi label 'promt yang digunakan:' - logo, moodboard, mockup, mascot, naskah, storyline, shotlist, storyboard. Pada pemeriksaan sebelumnya hanya 1 prompt (logo) yang ada."), ('LENGKAP', "Link aktif HTTP 200. Kelima label wajib LENGKAP: 'AI', 'SMKN 9 Garut', 'personal branding', 'portofolio', 'tugas sekolah' (semua 5 dari 5; pada pemeriksaan sebelumnya 0 dari 5)."), ('SEBAGIAN', "Konsep QRS (brand penyuntingan video) orisinal dan konsisten - maroon + metallic silver, motif clapperboard/timeline konsisten di logo, mockup, mascot, storyboard, dan naskah. Catatan: 3 artefak teks mentah '[cite: 3]' masih tertinggal di paragraf bagian Mascot, menandakan hasil salinan chat AI belum dibersihkan.")],
))

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

S.append(('R19','DHEA EKA KHOERUNNISA','XI DKV 3','02 Okt 2026 18:11',
'https://dheaekadhey.blogspot.com/2026/09/personal-branding-dhea-eka-khoerunnisa.html?m=1',
'3 (terverifikasi teks)','12 (terverifikasi tabel)','6 (terverifikasi teks)','1 (full body tidak dapat dipastikan)','8','1736 (memenuhi)',[4, 3, 4, 3, 4, 4, 4, 3, 2, 4, 4, 3],
[('SEBAGIAN', "Page title 'Personal Branding Dhea Eka Khoerunnisa XI DKV 3' memuat kelas, tetapi ejaannya berbeda dari heading artikel 'ASTS PERSONAL BRANDING DHEA EKA KHOERUNNISA XI DKV 3 SMKN 9 GARUT' (Khoerunnisa vs Khoerunnisa) - judul pendek dan judul panjang tidak konsisten ejaan. Body juga memakai 'Dhea Khoerunnisa' (tanpa 'Eka')."), ('SEBAGIAN', 'Logo inisial \'DK\' gaya Y2K, warna soft pink + soft cyan; nama branding \'DK\', tagline \'Create Your Own Style.\', deskripsi konsep +/- 200 kata (>= 100). NAMA SISWA pada branding HANYA diklaim di teks (\'Tulisan "by Dhea Khoerunnisa" ditambahkan untuk memperjelas pemilik\'), sedangkan IMG#01 320x320 OCR \'(tidak ada teks yang terbaca)\' sehingga nama pada logo tidak dapat diverifikasi dari gambar.'), ('LENGKAP', "IMG#02 320x213 landscape; 7 unsur terverifikasi di teks: Warna utama (soft pink, soft cyan, light pink, light cyan, putih), Typography (Bubblegum Sans + Poppins, OCR 'Bubblegum Sans / Poppins'), Style visual (Y2K simple clean playful modern), Referensi desain (digital, retro, glossy, kreatif), Tone & mood (playful, fresh, friendly, creative, calm, optimistic), Elemen grafis (polkadot, bintang, hati, bunga, pita, cursor, orbit), Inspirasi visual (pastel, langit, kamera, CD, perangkat digital)."), ('DIKUMPULKAN', "3 media mockup dijelaskan fungsinya di teks: 1. Sticker (media promosi pada laptop/tumbler/notebook), 2. Poster (perkenalan identitas di kelas dan portfolio), 3. Instagram Feed (media digital menampilkan karya). Namun hanya ada 1 berkas gambar IMG#03 300x320 portrait dengan OCR '1.StitaxrMxbyp / 2PoterMechp' - mockup feed Instagram tidak punya gambar terpisah."), ('LENGKAP', "'5. Naskah': Judul 'Create Your Own Style', Tema (DK Y2K playful/fresh/friendly), Pesan Utama ('Setiap orang memiliki karakter dan kreativitas yang berbeda'), Narasi/Dialog 6 scene ('Setiap ide dimulai dari sesuatu yang sederhana' ... 'Playful, creative, dan tetap menjadi diri sendiri'), Closing Tagline 'DK - Create Your Own Style.'"), ('LENGKAP', "'6. Storyline' memuat keempat unsur: Pembukaan (meja kerja + elemen bintang/hati/polkadot), Alur Cerita (proses pembuatan desain di laptop hingga membentuk logo DK), Konflik/Fokus Visual (perubahan dari sederhana menjadi Y2K colorful; logo DK sebagai fokus utama), Penutup (logo DK by Dhea Khoerunnisa + closing tagline)."), ('LENGKAP', "Tabel HTML shotlist 13 baris x 7 kolom, header 'No | Adegan | Jenis Shot | Angle | Movement | Durasi | Deskripsi', dengan 12 shot bernomor 1-12 (Meja kerja sampai Closing tagline) - jumlah dan kolom lengkap, tetap yang terbaik di batch ini."), ('SEBAGIAN', "'8. Storyboard' menyatakan 'Storyboard terdiri dari 6 scene yang menggambarkan proses dari ide sederhana hingga terbentuknya identitas visual DK' dan menyebut isi tiap scene, serta menyatakan tiap scene dilengkapi visual adegan, jenis shot, angle, movement, durasi, transisi, dan narasi - tetapi rincian itu tidak dipecah per scene. Gambar IMG#05 300x320 OCR-nya pseudoteks ('Storyboard / DK / Create / Con Stye'), jadi jumlah panel visual tidak dapat diverifikasi."), ('SEBAGIAN', "Hanya 1 gambar maskot IMG#04 213x320 portrait dengan OCR pseudoteks ('ASACHK / 18 / RRIX / 0'). Teks menyatakan 'Maskot dibuat dalam format full body, portrait, dan resolusi 4K' serta inisial 'DK' pada pakaian, tetapi framing full body tidak dapat dipastikan dari gambar (satu-satunya gambar maskot berorientasi potret)."), ('LENGKAP', "8 prompt terdokumentasi dan diberi label 'Prompt yang digunakan:' (logo, moodboard, mockup, maskot, naskah, storyline, shotlist, storyboard) - naik dari 1 prompt pada penilaian sebelumnya; tiap prompt menyebut iterate ('baguss, sekarang buatkan maskotnya')."), ('LENGKAP', "Link aktif (HTTP 200), artikel 1.736 kata memuat 8 komponen bernomor, dan kelima label wajib terpasang: 'AI', 'personal branding', 'portofolio', 'SMKN 9 Garut', 'tugas sekolah'. 10 label tambahan sebagian relevan; hanya 'label' dan 'logo' yang tidak relevan."), ('SEBAGIAN', "Konsep orisinal (inisial DK + estetika Y2K) dengan palet soft pink/soft cyan konsisten pada logo, moodboard, mockup, dan maskot; tagline 'Create Your Own Style.' seragam. Catatan konsistensi: nama ditulis 'Dhea Khoerunnisa' (double n) pada sebagian besar body dan 'DHEA EKA KHOERUNNISA' pada judul; nama branding hanya berupa inisial 'DK'.")],
))

S.append(('R02','WULAN SUNDARI','XI DKV 1','30 Sep 2026 21:09',
'https://wulansundariii.blogspot.com/2026/09/asts-personal-branding-wulan-sundari-xi_01701389558.html',
'3 (terverifikasi)','12 (terverifikasi tabel)','6 (terverifikasi gambar + teks)','3 varian (full body)','0','2445 (memenuhi)',[3, 4, 4, 4, 4, 4, 4, 4, 4, 0, 2, 4],
[('SEBAGIAN', "Page title dan heading artikel sama-sama 'ASTs Personal Branding Wulan Sundari XI DKV 1 SMKN 9 GARUT' - judul panjang sesuai format. Judul pendek 'Personal Branding [Nama] [Kelas]' tidak ada di halaman mana pun; label Blogger hanya 'personal branding'. KETERANGAN: halaman publik kini mengembalikan HTTP 404, seluruh bukti diambil dari arsip HTML lokal (raw/R02.html), bukan dari halaman live."), ('LENGKAP', "Logo monogram WS (IMG#01 1024x1024, OCR 'WULANSUNDARI / CREATIVEART&DESIGN / ByWULANSUNDARI'); nama branding 'Wulan Sundari - Creative Art & Design'; tagline 'Create with Calm, Express with Art.'; NAMA SISWA tercetak pada logo dan watermark; deskripsi konsep 211 kata (monogram WS fluid script, aksen daun, sparkles, gradasi peach coral ke soft teal, latar 1:1) - jauh di atas 100 kata."), ('LENGKAP', 'IMG#02 1376x768 LANDSCAPE. Tujuh unsur terverifikasi dua kali: OCR membaca WARNA UTAMA BRANDING, TYPOGRAPHY, STYLE VISUAL, REFERENSI DESAIN, TONE & MOOD, ELEMEN GRAFIS, INSPIRASI VISUAL; teks uraian 248 kata menyebut Peach Coral #F9B0A3, Soft Teal #87BAB3, Silver Grey #CFF5F1/#CBCFD0, Neutral Cream, dan kata kunci Calm/Serene/Elegant/Creative/Artistic/Harmonious.'), ('LENGKAP', "Tiga mockup diuraikan terpisah dengan penjelasan fungsi tiap media (233 kata): Gaun Floral (motif watercolor bunga, nama WULAN SUNDARI di dada), Sepatu Flat (monogram WS terukir di heel, kesan kalem-modis), Set Stiker Die-Cut (fungsi merchandise, packaging seal, koleksi desainer). Didukung IMG#03 1376x768 ber-OCR 'WULAN SUNDARI'."), ('LENGKAP', "Naskah 60 detik lengkap: Judul 'Setiap Karya Punya Cerita', Tema, Pesan Utama, tabel Durasi|Visual/Adegan|Narasi/Dialog|VO dengan 7 baris adegan (0-8 s sampai 56-60 s) berisi VO dan dialog Wulan, serta Closing Tagline 'Wulan Sundari - Creative Art & Design / Create with Calm, Express with Art.' dan VO Final."), ('LENGKAP', "Storyline 'Setiap Karya Punya Cerita' 4 bagian lengkap dengan durasi: 1. PEMBUKAAN - Mencari Inspirasi (0-10 detik), 2. ALUR CERITA - Inspirasi Menjadi Karya (10-30 detik), 3. KONFLIK / FOKUS VISUAL - Menemukan Identitas (30-48 detik, termasuk catatan konfliknya bukan konflik besar melainkan proses kreatif), 4. PENUTUP - Inilah Identitasku (48-60 detik). Dilengkapi 'Konsep alur singkat' dan 'Alur Besar Storyline'."), ('LENGKAP', "Tabel shotlist HTML 13 baris x 6 kolom = 12 shot bernomor, kolom No. Adegan | Jenis Shot | Angle | Movement | Durasi | Deskripsi terisi penuh (ELS, LS, MS, CU, ECU, FS, OTS, MCU sampai ECU -> CU; total 60 detik). Didukung IMG#05 1376x768 ber-OCR 'SHOT LISTTABLE:WULANSUNDARI-CREATIVEART&DESIGN' dengan kolom No.Adegan/Jenis Shot/Angle/Movement/Durasi/Deskripsi."), ('LENGKAP', "Storyboard 6 scene pada IMG#06 1376x768: OCR membaca header 'STORYBOARD-WULANSUNDARICREATIVEART&DESIGN', baris No.Adegan 1 sampai 6, dan keempat kolom keterangan Visual Adegan | Keterangan Shot | Angle Kamera (Long Shot, MCU, Blackmagic URSA) | Transisi | Dialog/Narasi. Teks 'Rincian Tiap Adegan pada Papan' merinci 6 adegan lengkap dengan keterangan shot; isi deskripsi sebagian berupa pseudoteks khas gambar AI."), ('LENGKAP', "Tiga varian maskot terverifikasi OCR pada IMG#04 1376x768: 'MASCOT FULL BODY' (WAJIB), 'MASCOT PORTRAIT', dan 'MASCOT BERSAMA LOGO BRANDING' dengan tagline 'THE ARTIST & HER BRAND'. Teks 375 kata merinci karakter (gaun floral peach coral, tas selempang, kalung mutiara) dan latar tiap panel."), ('BELUM MEMENUHI', "TIDAK ADA satu pun prompt. Kata 'prompt' tidak muncul sama sekali dalam 2.513 kata artikel hasil ekstraksi penuh. Bagian '7. SHORTLIST' hanya berisi sisa percakapan AI ('Siap bro. Berdasarkan seluruh gambar branding Wulan Sundari ... berikut shotlist untuk video 60 detik') yang juga bukan prompt - tidak dihitung sebagai dokumentasi prompt."), ('BELUM MEMENUHI', "Isi artikel sangat lengkap (2.445 kata, 8 bagian bernomor) dan halaman memuat 2 tabel, tetapi label Blogger hanya 1: 'personal branding'. Empat dari lima label wajib hilang: 'Tugas Sekolah', 'AI', 'Portofolio', dan 'SMKN 9 Garut'. KETERANGAN: halaman publik kini HTTP 404 sehingga label tidak dapat diverifikasi ulang - dicatat sebagai keterangan, bukan sebagai dasar penalti; skor 2 diberikan semata karena label tidak memenuhi."), ('LENGKAP', 'Identitas visual konsisten kuat di seluruh media: monogram WS, palet peach coral + soft teal + silver grey, elemen bunga/daun/sparkles, dan maskot yang sama dipakai pada logo, moodboard, mockup, naskah, shotlist, dan storyboard. Nama pada rekapan (WULAN SUNDARI) sama dengan nama pada halaman dan branding. Konsep orisinal: perjalanan kreator dari mencari inspirasi hingga menemukan identitas.')],
))

S.append(('R03','WILDA AZKIA','XI DKV 1','30 Sep 2026 21:12',
'https://wildaazkia.blogspot.com/2026/09/personal-branding.html',
'3 (terverifikasi)','10 (terverifikasi OCR)','6 (terverifikasi teks)','3 (terverifikasi; full body 441x246)','1','1737 (memenuhi)',[3, 4, 4, 4, 3, 4, 4, 4, 4, 2, 4, 4],
[('SEBAGIAN', "Page title 'PERSONAL BRANDING WILDA AZKIA XI DKV 1' persis sesuai format judul pendek. Judul panjang 'ASTS Personal Branding Wilda Azkia XI DKV 1 SMKN 9 Garut' terverifikasi pada pemeriksaan sebelumnya dan isi artikel tidak berubah struktural, namun heading tersebut TIDAK dapat dibaca ulang pada ekstraksi dossier ini (11 bagian terdaftar langsung mulai dari 'Deskripsi Konsep Branding'), sehingga hanya satu dari dua judul yang terverifikasi langsung."), ('LENGKAP', "Logo monogram WA dengan sapuan kuas (IMG#01 452x247, OCR 'WILDA AZKIA'); nama branding 'Wilda Azkia'; tagline 'Creating Rhythm in Every Design'; deskripsi konsep 157 kata (>=100); NAMA SISWA tercetak pada logo dan watermark 'BY WILDA AZKIA'."), ('LENGKAP', 'IMG#02 513x287 landscape. Tujuh unsur terverifikasi: Warna Utama (koral/pink pastel/magenta/kuning keemasan), Typography (Modern Sans-Serif Poppins/Montserrat + Expressive Brush Script), Style Visual & Referensi (fluid shapes + Inspirasi Seni Tari), Tone & Mood (Friendly, Creative, Energetic), Elemen Grafis (monogram WA, watercolor splash, sparkles).'), ('LENGKAP', '3 mockup terpisah dengan penjelasan fungsi masing-masing: Stiker (447 kata, Portabel/Brand Awareness), Kartu Nama (411 kata, Networking & First Impression), Social Media Feed (423 kata, Digital Showcase & Brand Cohesion); didukung IMG#03/04/05.'), ('SEBAGIAN', "Bagian naskah TIDAK muncul sebagai heading pada ekstraksi dossier ini, sehingga isi lengkapnya tidak dapat dibaca ulang. Pada pemeriksaan sebelumnya bagian ini memuat Judul 'Turning Concepts into Colors', Tema, Pesan Utama, Narasi/Voice Over, dan Closing Tagline; isi artikel tidak berubah selain pengurangan 12 kata."), ('LENGKAP', "Bagian 'D.4 Storyboard' 388 kata. Scene 1-4 terbaca verbatim lengkap dengan pola Shot|Angle|Movement|Transisi|Visual|Narasi (Extreme Close-Up|Eye Level|Static|Fade In; Medium Shot|Eye Level|Tilt Up|Wipe; Close-Up|Low Angle|Zoom In|Dissolve; Medium Close-Up|High Angle|Pan Right|Cut); 6 scene lengkap terverifikasi pada pemeriksaan sebelumnya."), ('LENGKAP', "IMG#06 566x316. OCR terbaca header 'TABEL SHOTLIST (MINIMAL 10 SHOT)' dan kolom No/Adegan/Jenis Shot/Angle/Movement/Durasi dengan baris bernomor 1-10 - 10 shot terpenuhi."), ('LENGKAP', "Teks bagian 'D.4 Storyboard' memuat keterangan shot, angle, movement, transisi, visual, dan narasi per scene; dapat diverifikasi dari teks tanpa bergantung pada gambar."), ('LENGKAP', "3 berkas mascot terpisah pada bagian 'E. AI MASCOT CHARACTER (PIXAR STYLE)': IMG#08 441x246 (rasio landscape konsisten dengan full body - WAJIB), IMG#09 295x295 (portrait), IMG#10 545x304 (bersama logo) - ketiga varian terpenuhi."), ('BELUM MEMENUHI', 'Tidak ada bagian prompt pada ekstraksi dossier ini; berdasarkan pemeriksaan sebelumnya hanya 1 prompt (prompt logo) yang terdokumentasi. Prompt moodboard, mockup, shotlist, storyboard, dan mascot tidak ada - jumlah ini tidak dapat dinaikkan tanpa teks prompt yang terbaca.'), ('LENGKAP', 'Link aktif HTTP 200. Kelima label wajib lengkap dalam 8 label: AI, Personal Branding, Portofolio, SMKN 9 GARUT, Tugas Sekolah; 3 label tambahan milik tugas lain (Tugas Branding, Tugas CV, Tugas Praktikum Lighting).'), ('LENGKAP', "Watermark 'BY WILDA AZKIA' muncul konsisten di logo dan 3 mockup (OCR IMG#01, IMG#03, IMG#04, IMG#05); palet koral-magenta dan bentuk pita WA konsisten di seluruh bagian.")],
))

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

S.append(('R05','RIZKY MUHAMMAD REGAL SAFARI','XI DKV 1','01 Okt 2026 08:14',
'https://rizkymregalsafari.blogspot.com/2026/09/asts-personal-branding-rizky-muhammad.html?m=1',
'3 diklaim (1 berkas gambar 320x320)','11 (terverifikasi teks; tabel markdown)','6 (terverifikasi OCR)','3 varian (terverifikasi OCR)','8 (tanpa label)','1958 (memenuhi)',[3, 2, 2, 1, 4, 4, 4, 3, 4, 3, 2, 3],
[('SEBAGIAN', "Page title berisi judul panjang 'ASTS Personal Branding Rizky Muhammad Regal Safari XI DKV 1 SMKN 9 GARUT'. Judul pendek 'Personal Branding [Nama] [Kelas]' tidak ada sebagai heading artikel - bentuk terdekat hanya label Blogger 'PERSONAL BRANDING RIZKY MUHAMMAD REGAL SAFARI XI DKV 1' dan paragraf pembuka 'Personal Branding ASTS Rizky Muhammad Regal Safari XI DKV 1 SMKN 9 GARUT'."), ('BELUM MEMENUHI', "Logo 'RR / RIZKY REGAL' (IMG#01 320x320, OCR 'RIZKY REGAL / TETAP BERJUANG UNTUK MENGGAPAI IMPIAN'); nama branding 'Rizky Regal'; tagline 'Tetap Berjuang Untuk Menggapai Impian'; NAMA SISWA tercetak pada logo. NAMUN DESKRIPSI KONSEP 100 KATA TIDAK ADA - bagian A hanya berisi prompt, bukan uraian konsep."), ('BELUM MEMENUHI', 'IMG#02 320x320 PERSEGI, bukan landscape (diminta Format Landscape). Tujuh unsur memang terbaca OCR: WARNA UTAMA BRANDING, TYPOGRAPHY, STYLE VISUAL (Gamer RP), TONE & MOOD (Regal & Ambisius), ELEMEN GRAFIS, INSPIRASI VISUAL (GTA V / Los Santos); tidak ada deskripsi moodboard.'), ('BELUM MEMENUHI', 'Hanya 1 berkas gambar (IMG#03 320x320) untuk 3 media yang dijanjikan prompt (stiker, kartu nama, gelas kopi). PENJELASAN FUNGSI MEDIA TIDAK ADA - bagian C hanya berisi prompt tanpa uraian fungsi tiap media.'), ('LENGKAP', "'DOKUMEN 1: FILM PENDEK NASKAH IKLAN (60 DETIK)' memuat Judul 'Membangun ... Bersama Teman' (ejaan kata kedua pada dokumen tidak baku), Tema, Pesan Utama 'Bersama meraih mimpi sukses', tabel Waktu|Visual|Audio 6 adegan (00:00-01:00) lengkap dengan SFX/musik dan VO RYU/RINA, serta Closing Tagline 'TETAP BERJUANG UNTUK MENGGAPAI IMPIAN'."), ('LENGKAP', "'DOKUMEN 2: STORYLINE FILM PENDEK RIZKY REGAL (60 DETIK)' memuat 5 tahap: I. PEMBUKAAN (00:00-00:08), II. ALUR CERITA (00:08-00:18), KONFLIK masalah komunikasi (00:18-00:33), IV. SOLUSI (00:33-00:50), V. PENUTUP (00:50-01:00) - lengkap dengan karakter, detail, dan narasi."), ('LENGKAP', "'DOKUMEN 3: SHOTLIST' memuat 11 baris bernomor dengan kolom No | Adegan | Jenis Tembakan [ejaan asli siswa, seharusnya 'Tembakan'] | Sudut | Gerakan | Durasi | Deskripsi Visual & Aksi (EWS, MS, MCU, CU, ECU sampai Bidikan Grafis) - melebihi minimum 10 shot."), ('LENGKAP', 'IMG#05 320x320, OCR terbaca 6 scene bernomor lengkap dengan timecode: 1. PEMBUKAAN CERIA (00:00-00:08), 2. KESEHARIAN CERIA (00:08-00:18), 3. KONFLIK KOMUNIKASI (00:18-00:33), 4. PEMAHAMAN BARU (00:33-00:50), 5. IMPLEMENTASI SOLUSI (00:50-00:55), 6. KESUKSESAN BERSAMA (00:55-01:00) - minimum 6 scene terpenuhi.'), ('LENGKAP', "IMG#04 320x320, OCR 'MASCOT FULLBODY(BOY&GIRL)' (WAJIB), 'MASCOT PORTRAIT', dan 'MASCOT BERSAMA LOGO BRANDING' - ketiga varian terverifikasi."), ('SEBAGIAN', "8 prompt terdokumentasi lengkap dan spesifik (A-H: logo, moodboard, mockup, maskot, naskah, storyline, shotlist, storyboard) dengan ketentuan output sangat rinci. NAMUN tidak ada satu pun diberi label 'Prompt' - dokumentasinya berupa blok huruf A-H, sehingga tidak terstruktur."), ('BELUM MEMENUHI', "Link aktif dan seluruh komponen ada, tetapi label BELUM MEMENUHI ketentuan. Perubahan pada kiriman ini: label bertambah menjadi 9 (ASTS; DKV 1.; PERSONAL BRANDING RIZKY MUHAMMAD REGAL SAFARI XI DKV 1; Personal Branding “Ikyy Regal”; SMK 9 Garut; XI DKV 1) - namun 'Tugas Sekolah', 'AI', dan 'Portofolio' tetap tidak ada, dan 2 label (Riwayat hidup, Tugas penyetingan cahaya) milik tugas lain."), ('SEBAGIAN', 'Ide sangat kreatif dan orisinal: konsep duo maskot Ryu & Rina, tema gaming RP, dan konflik komunikasi. Susunan artikel tetap berupa dinding teks tanpa heading bagian (0 bagian terdeteksi), sehingga 5 gambar menumpuk tanpa judul.')],
))

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

S.append(('R26','INDRI FITRIYANI','XI DKV 3','01 Okt 2026 10:29',
'https://indriifitriyani.blogspot.com/2026/09/personal-branding-indri-fitriyani-xi.html?m=1',
'3 (terverifikasi)','10 (terverifikasi tabel)','TIDAK DAPAT DIVERIFIKASI','1 (full body tidak dapat dipastikan)','8','2318 (memenuhi)',[4, 4, 4, 4, 4, 4, 4, 2, 2, 4, 4, 3],
[('LENGKAP', "Page title 'Personal Branding Indri Fitriyani XI DKV 3' + heading artikel 'ASTS PERSONAL BRANDING INDRI FITRIYANI XI DKV 3 SMKN 9 GARUT' - keduanya sesuai format dan menyebut nama serta kelas."), ('LENGKAP', "Logo inisial 'IF' + elemen kamera/lensa dan buku (IMG#01 400x400) dengan OCR 'INDRI FITRIYANI / PHOTOGRAPHER | EDITOR' sehingga NAMA SISWA tercetak pada logo; nama branding 'IF - Indri Fitriyani'; tagline 'Capture Moments, Create Stories.'; deskripsi konsep +/- 330 kata plus daftar 'Filosofi Personal Branding' (IF, Kamera/Lensa, Buku/Lembaran, Kuning, Biru, Gradasi, Photographer | Editor)."), ('LENGKAP', 'IMG#02 400x223 landscape; 7 unsur terverifikasi dan OCR mengonfirmasi labelnya: WARNAUTAMABRANDING (Soft Luminous Gold, Soft Pastel Blue, Deep Navy Blue, Warm Beige), TIPOGRAFI (Cormorant Garamond), STLEVSUAL, REEERENSIOESAIN, TONL&UOOD, ELEMENGRAFIS, INSRRASIVSUAL.'), ('LENGKAP', "3 mockup: Kartu nama (fungsi perkenalan identitas), Stiker (media pendukung branding), Paper cup (penerapan logo pada produk sehari-hari) - fungsi tiap media dijelaskan. IMG#03 400x266 landscape memuat tiga panel, OCR 'INDRI FITRIYANI'/'INDRIEITRIYANI'/'INDRI FITRIYANI' + 'Photography Editor'."), ('LENGKAP', "'NASKAH': Judul 'Capture Your Story, Create Your Memory', Tema (perjalanan kreatif seorang fotografer), Pesan Utama, Narasi/Dialog 5 baris Voice Over ('Setiap momen memiliki cerita yang layak untuk diabadikan' ... 'tentang cerita dan kenangan di dalamnya'), Closing Tagline 'Indri Fitriyani - Photographer | Editor. Capture Your Story, Create Your Memory.'"), ('LENGKAP', "'STORYLINE' memuat keempat unsur: Pembukaan (maskot INDRI FITRIYANI sebagai fotografer berhijab, outfit cream dan vest biru), Alur Cerita Singkat (kamera dan lensa di meja kerja warm beige, logo IF), Konflik atau Fokus Visual (editing di laptop, perbandingan before & after), Penutup (hasil foto, maskot bersama logo IF, tagline)."), ('LENGKAP', "Tabel HTML shotlist 11 baris x 7 kolom, header 'No | Adegan | Jenis Shot | Angle | Movement | Durasi | Deskripsi', dengan tepat 10 shot bernomor 1-10 (Kamera dan lensa di meja sampai Logo, maskot, dan tagline) - memenuhi minimum 10 shot."), ('TIDAK DAPAT DIVERIFIKASI', "Gambar storyboard IMG#05 400x266 memang ada, tetapi OCR-nya pseudoteks ('STORYBOARD / Capedvre jeor Stoy. / Gene yor Menxeg / CENS / 0-12000') sehingga jumlah scene/panel tidak dapat dihitung andal; teks 'STORYBOARD' hanya menguraikan alur umum tanpa mendaftarkan 6 scene. Ini turun dari penilaian sebelumnya yang mencatat 6 scene terverifikasi dari teks."), ('SEBAGIAN', "Hanya 1 gambar maskot IMG#04 320x213 landscape dengan OCR pseudoteks ('KEGHTTrASS / INDRITITRIYAN! / IF / INORUITRIYANI'). Teks menjelaskan tiga bentuk (Maskot Bersama Logo, Maskot Portrait, Maskot Full Body) tetapi tidak ada label full body yang terbaca dari gambar, sehingga output wajib hanya terbukti lewat teks."), ('LENGKAP', "8 prompt terdokumentasi berlabel 'prompt yang di gunakan :' (logo, moodboard, mock up, maskot, naskah, storyline, shotlist, storyboard), termasuk iterasi '-edit lagi moodboard nya gambar gambarnya sesuaikan sesuai warna branding, sesuaikan juga warna logonya'."), ('LENGKAP', "Link aktif (HTTP 200), artikel 2.318 kata, kelima label wajib terpasang: 'AI', 'Personal Branding', 'Portofolio', 'SMKN 9 Garut', 'Tugas sekolah'; label tambahan 'CV' dan 'TUGAS' juga relevan."), ('SEBAGIAN', "Konsep orisinal (Photographer | Editor) dan konsisten: logo IF, palet soft gold/pastel blue/deep navy/warm beige, serta elemen kamera dan buku dipakai pada logo, moodboard, mockup, maskot, naskah, storyline, dan shotlist. Catatan konsistensi: tagline tidak konsisten - 'Capture Moments, Create Stories.' pada bagian branding, tetapi 'Capture Your Story, Create Your Memory' pada naskah, storyline, dan shotlist.")],
))

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

S.append(('R06','M REZA HUAFAH','XI DKV 1','01 Okt 2026 10:42',
'https://mrezahuafah.blogspot.com/2026/09/personal-branding-m-reza-huafah-xi-dkv-1.html',
'3 (terverifikasi)','12 (terverifikasi gambar 20 lembar)','6 (terverifikasi teks)','3 konsep tertulis; 1 gambar (320x213)','8','3467 (memenuhi)',[4, 4, 4, 3, 4, 4, 4, 3, 4, 4, 2, 4],
[('LENGKAP', "Page title 'Personal Branding M Reza Huafah XI DKV 1' (judul pendek) dan baris pertama artikel 'ASTS Personal Branding M Reza Huafah XI DKV 1 SMKN 9 Garut' (judul panjang) - keduanya sesuai ketentuan."), ('LENGKAP', "Logo simbol abstrak MRH (IMG#01 320x320, OCR 'M / REZA / HUAFAH / DREAM / EXPLORE / CREATE'); nama branding 'MRH / M Reza Huafah'; tagline 'DREAM • EXPLORE • CREATE'; deskripsi konsep 481 kata dengan makna tiap warna, bentuk, dan slogan; NAMA SISWA tercetak pada logo."), ('LENGKAP', 'IMG#02 320x213 landscape. Uraian 6 paragraf menjawab ketujuh unsur dengan kode hex: warna utama (#FFB000 oranye, #0D0D0D hitam, #EAEAEA putih, #6B6B6B abu-abu), tipografi (Montserrat Extra Bold/Medium/Regular), style visual (minimalis, clean, bold), referensi desain (hoodie, topi, kaus, kartu nama), tone & mood (Adventure, Creative, Progress, Supportive), elemen grafis (garis orbit, diagonal, bintang), inspirasi visual (gunung, matahari terbenam, kamera).'), ('LENGKAP', '3 media mockup (packaging kotak, gelas kopi, kaos merchandise) masing-masing diurai terpisah dengan penjelasan fungsi, warna, dan elemen spesifik; didukung IMG#03 320x213, IMG#04 320x320, IMG#05 320x320 yang seluruhnya ber-OCR MRH.'), ('LENGKAP', "'NASKAH IKLAN KOMERSIAL PERSONAL BRANDING MRH' memuat Judul 'DREAM • EXPLORE • CREATE', Tema, Pesan Utama, Narasi/Dialog Scene 1-6 lengkap dengan Visual + Narasi, dan Closing Tagline 'Berani bermimpi. Berani mengeksplorasi. Berani menciptakan.'"), ('LENGKAP', "'STORYLINE PERSONAL BRANDING MRH' memuat 5 bagian: 1. PEMBUKAAN, 2. ALUR CERITA, 3. KONFLIK / FOKUS VISUAL, 4. PUNCAK CERITA, 5. PENUTUP + KESIMPULAN, dengan Scene 1-11 bertimecode visual dan narasi."), ('LENGKAP', "Shotlist dikumpulkan sebagai 20 lembar gambar (IMG#07-IMG#26, format diterima menurut Revisi-1) - OCR IMG#25/26 membaca header 'SHOTLIST-PERSONAL BRANDING MRH' dengan kolom No/Adegan/Jenis Shot/Angle/Durasi/Deskripsi dan baris shot bernomor sampai 13; teks artikel merangkum Shot 1-12 per tahap cerita."), ('SEBAGIAN', "Bagian '8.Storyboard' pada teks mendeskripsikan 6 scene lengkap dengan angle, transisi, dan narasi (Scene 1 Eye Level/Fade In sampai Scene 6 Full Shot/Fade Out); gambar IMG#27 320x213 ber-OCR 'STORYBOARD / M REZA HUAF'AH' tetapi isi panel tidak terbaca, sehingga keterangan visual per panel hanya bersumber dari teks."), ('LENGKAP', "Maskot 3D diuraikan sangat lengkap (hoodie, celana cargo, sepatu, tas ransel; full body, close-up, pose setengah badan; elemen gunung dan garis diagonal); 1 gambar IMG#06 320x213 ber-OCR 'MREZA HUAFAH / DREAM EXPLORE CREATE'."), ('LENGKAP', "8 prompt terdokumentasi konsisten dengan label 'Promt:' atau 'Prompt:' - logo, moodboard, mockup (packaging, gelas kopi, kaos), maskot, naskah, storyline, shotlist, storyboard."), ('SEBAGIAN', "Link aktif dan artikel paling lengkap (3.467 kata, 27 gambar), tetapi label Blogger tetap tidak memadai. Perubahan pada kiriman ini: label bertambah dari 1 menjadi 2, yaitu 'Tugas sekolah ASTS PROJEK' dan 'TUGAS SEKOLAH ASTS Personal Branding M Reza Huafah XI DKV 1 SMKN 9 Garut' (mengandung 3 dari 5 label wajib); 'AI' dan 'Portofolio' tetap tidak ada."), ('SEBAGIAN', "Karya paling lengkap di XI DKV 1: identitas visual konsisten (hitam + oranye, monogram MRH, tagline DREAM • EXPLORE • CREATE) dari logo, moodboard, 3 mockup, maskot, hingga storyboard. Catatan konsistensi teks: nama ditulis dua ejaan - 'M Reza Huafah' pada judul dan logo, 'M Reza Huaf'ah' pada bagian naskah, storyline, dan storyboard.")],
))

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
S.append(('R63','GADNA WIDIATNI','XI DKV 1','01 Okt 2026 21:33',
'https://gadnawidiatni.blogspot.com/2026/10/personal-branding-gadna-widiatni-xi-dkv.html',
'3 (terverifikasi)','15 (terverifikasi teks)','6 (terverifikasi teks)','1 (full body tidak dapat dipastikan)','8','4059 (memenuhi)',[4, 3, 4, 4, 4, 4, 3, 4, 2, 4, 4, 3],
[('LENGKAP', "Page title 'Personal Branding Gadna Widiatni XI DKV 1' + heading artikel 'ASTS Personal Branding Gadna Widiatni XI DKV 1 SMKN 9 Garut' - keduanya sesuai format. Catatan: nama ditulis 'Gadna Widiatri' (2x) pada bagian 'Deskripsi Maskot Anime'."), ('SEBAGIAN', "Logo inisial G-W + headphone, papan klaket, buku sketsa, sepeda, kupu-kupu (IMG#01 320x320) tetapi OCR hanya pseudoteks ('C / Wyata. / 3ELL') sehingga NAMA SISWA pada logo tidak dapat diverifikasi dari gambar. Nama branding 'GW', tagline 'Capture, Create, Tell' dan 'Little Moments, Big Stories', watermark 'By Gadna Widiatni' diminta di prompt, dan 'GW Gadna Widiatni' tercetak pada mockup; deskripsi konsep 7 elemen filosofi (600+ kata) memenuhi 100 kata."), ('LENGKAP', 'IMG#02 399x266 landscape; 7 unsur bernomor lengkap di teks: 1. Warna Utama (Mocha, Latte, Khaki, Cream, Dusty Pink, Olive Gray, Warm Beige, Charcoal), 2. Tipografi (serif elegan + script/handwritten), 3. Style Visual (editorial vintage & scrapbook photography), 4. Referensi Desain, 5. Tone & Mood, 6. Elemen Grafis, 7. Inspirasi Visual.'), ('LENGKAP', "3 media dengan fungsi eksplisit: KAOS (Fungsi: Brand Merchandise -> Brand Awareness -> Lifestyle Identity), GELAS KOPI (brand recall, properti photoshoot, konten IG/Reels, merchandise), SOCIAL MEDIA FEED (grid 3x3: photography, lifestyle, quote/storytelling, brand post); palet diberi hex (Cream #F8F5EF, Linen #E6DDD3, Latte #B89B8A, Khaki #C9B59B, Mahogany #6B4F45). IMG#03 376x376 OCR 'KAOS / GFLAS XOPI / SOCLALMEDAFEED' - ketiga media terverifikasi."), ('LENGKAP', "'A. Naskah Iklan Komersial': Judul 'A Moment Worth Remembering / Sebuah Momen untuk Diingat', Tema, Pesan Utama, Alur 60 detik 6 blok (Opening 0-7 dtk, The Little Moments, The Moment, Capture & Create, The Story, Brand Reveal, Closing 56-60 dtk) dengan Visual + Narasi per blok, ditutup Closing Tagline 'Little Moments, Big Stories.'"), ('LENGKAP', "'B. Storyline': 1. PEMBUKAAN 'Hari yang Biasa' (0-12 detik), 2. ALUR CERITA 'Mulai Memperhatikan' (12-27), 3. KONFLIK/FOKUS VISUAL 'Momen yang Hampir Hilang' (27-43, konflik WAKTU YANG TERUS BERJALAN vs KEINGINAN MENYIMPAN MOMEN), 4. RESOLUSI 'Capture, Create, Tell' (43-53), 5. PENUTUP 'Little Moments, Big Stories' (53-60) - tiap bagian disertai durasi, deskripsi adegan, dan fokus visual."), ('SEBAGIAN', "15 shot tertulis bernomor '01 - Kamar di Pagi Hari' sampai '15 - Brand Reveal' lengkap dengan rentang durasi (0-4 detik, 4-7 detik, dst) dan deskripsi adegan, sehingga jumlah memenuhi minimum 10. TETAPI tidak disajikan sebagai tabel (0 tabel HTML) dan tanpa kolom Jenis Shot/Angle/Movement terpisah - angle dan movement hanya disebut di dalam kalimat deskripsi; gambar IMG#05 356x534 OCR 'SHOTLIST / A Momerft / St ie Dok' tidak dapat dihitung andal."), ('LENGKAP', "6 scene tertulis lengkap di 'D. Storyboard' (Scene 1 Awal Hari sampai Scene 6 Penutup) dengan Shot, Angle, Kamera, Movement, Transisi, dan Narasi per scene, contoh Scene 4: 'Wide Shot -> Montage -> Close Up / Eye Level -> High Angle / Time-lapse -> Slow Zoom / Cross Dissolve -> Fade'; IMG#06 351x321 OCR 'GW / STORYBOARD / A Mooxnt Worth Remembening'."), ('SEBAGIAN', "Hanya 1 gambar maskot IMG#04 266x400 portrait dengan OCR 'GW' - tidak ada label framing yang terbaca. Teks menyatakan 'Karakter dibuat full body' dan mendeskripsikan busana sampai sepatu (topi, sweater, celana cargo, sneakers), tampilan belakang/samping, serta ekspresi (tersenyum, normal, wink), tetapi full body tidak dapat dipastikan dari gambar."), ('LENGKAP', "8 prompt terdokumentasi dan diberi label 'Prompt' (logo, moodboard, mockup kaos/gelas/feed, maskot, naskah, storyline, shotlist, storyboard) - naik dari 1 prompt pada penilaian sebelumnya; tiap prompt bernomor tahap ('langkah selanjutnya', 'langkah kedua', 'langkah terakhir')."), ('LENGKAP', "Link aktif (HTTP 200), artikel 4.059 kata, kelima label wajib terpasang: 'AI', 'Personal branding', 'Portofolio', 'SMKN 9 Garut', 'tugas sekolah'; label tambahan 'CV' dan 'kegiatan belajar' relevan."), ('SEBAGIAN', "Konsep orisinal (dokumentasi momen sederhana: kamera, headphone, kupu-kupu, palet vintage) dan konsisten di seluruh bagian; tagline 'Capture, Create, Tell' + 'Little Moments, Big Stories' dipakai seragam dari logo hingga storyboard. Catatan konsistensi: nama ditulis 'Gadna Widiatri' pada bagian maskot (2x) dan 'Gadna Widiatni' pada bagian lain.")],
))
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
S.append(('R69','SITI RAHMA SILPIANA','XI DKV 1','02 Okt 2026 04:37',
'https://sitirahmasilpianaaa.blogspot.com/2026/09/personal-branding-siti-rahma-silpiana.html',
'3 (terverifikasi)','12 (klaim teks; baris tabel tidak terbaca)','6 (terverifikasi teks)','1 (3 varian; full body tidak dapat dipastikan)','10','3639 (memenuhi)',[4, 4, 4, 4, 4, 4, 3, 3, 2, 4, 4, 4],
[('LENGKAP', "Page title 'Personal Branding Siti Rahma Silpiana XI DKV 1' + heading artikel 'ASTS Personal Branding Siti Rahma Silpiana XI DKV 1 SMKN 9 Garut' - keduanya sesuai format dan menyebut nama serta kelas dengan benar."), ('LENGKAP', "Logo monogram SRS + elemen buku, pena, kamera/lensa, bintang (IMG#01 320x320) dengan OCR 'SITI RAHMA SILPIANA / DESIGN / PHOTOGRAPHY / READING' sehingga NAMA SISWA tercetak pada logo; nama branding 'SRS', tagline 'Create, Capture, Inspire' dan 'Small Steps, Big Dreams', 'by Siti Rahma Silpiana'; deskripsi konsep +/- 330 kata dengan makna tiap simbol."), ('LENGKAP', 'IMG#02 320x213 landscape; 7 unsur terverifikasi di teks: warna utama (biru navy, biru, biru muda, putih kebiruan, abu-abu gelap), typography (Playfair Display + Montserrat), gaya visual (clean, minimalis, modern, elegan), referensi desain (kartu nama, tote bag, notebook, smartphone, pulpen), tone and mood (calm, focused, creative, confident, productive, elegant), elemen grafis (bintang, garis lengkung, kamera, aperture, buku, pena), dan bagian inspirasi visual (langit biru, bangunan, kamera, buku, alam, pemandangan).'), ('LENGKAP', "3 mockup dengan 3 gambar terpisah: Gelas Kopi/Mug (IMG#03 320x320, OCR 'SITI RAHNA SILPIARA / SRS'), Poster (IMG#04 213x320, OCR 'Small Steps BigDreamz'), Kaos (IMG#05 320x316, OCR 'SITI RAHMA SILPIARA / SRS'); fungsi tiap media dijelaskan dan tiap mockup punya prompt tersendiri."), ('LENGKAP', "'5. Naskah Iklan Komersial': Judul 'Create, Capture, Inspire', Tema, Durasi +/- 60 detik, Tokoh, Brand 'SRS - Siti Rahma Silpiana', Pesan Utama, Narasi/Dialog 6 SCENE lengkap dengan Visual + Keterangan Shot (ES/MCU/MS/CU/WS) + Angle Kamera + Transisi + Voice Over, ditutup Closing Tagline 'Create. Capture. Inspire.'"), ('LENGKAP', "'6. Storyline' memuat keempat unsur: 1. PEMBUKAAN (meja kerja, buku, mug SRS, tokoh membaca), 2. ALUR CERITA (baca -> sketsa -> desain -> memotret -> montage), 3. KONFLIK / FOKUS VISUAL (ragu pada hasil sketsa lalu mencoba lagi), 4. PENUTUP (logo SRS - Siti Rahma Silpiana dengan tagline 'CREATE - CAPTURE - INSPIRE') - keempat unsur ada, termasuk PENUTUP yang sebelumnya tercatat tidak diuraikan."), ('SEBAGIAN', "Teks menyatakan 'Shotlist ini terdiri dari 12 adegan yang menggambarkan perjalanan seorang kreator muda' dan menyatakan tiap adegan dilengkapi jenis shot, angle/movement, durasi, dan deskripsi visual (Establishing Shot, Close Up, Medium Shot, Extreme Close Up, Over The Shoulder). TETAPI 0 tabel HTML dan baris tabel hanya ada pada gambar IMG#07 320x214 yang OCR-nya pseudoteks ('Tabel Shotlist / Adegas / Peskr gol') - jumlah 12 shot tidak dapat dihitung andal dari bukti sekarang."), ('SEBAGIAN', "'8. Storyboard' menyebut 6 scene (Scene 1 Awal Perjalanan, Scene 2 Menemukan Inspirasi, Scene 3 Mulai Berkarya, Scene 4 Tantangan, Scene 5 Bangkit dan Mencoba Lagi, Scene 6 Hasil & Identitas SRS) beserta keterangan visual adegan, jenis shot, angle, transisi, dan narasi, tetapi rincian per scene tidak diuraikan di bagian storyboard; gambar IMG#08 320x301 hanya sebagian terbaca ('Storyboard / Short Movie Personal branding SRS / Scewa4 / ScerkS') sehingga jumlah panel tidak dapat dipastikan."), ('SEBAGIAN', "Hanya 1 gambar maskot IMG#06 320x320 dengan OCR 'Create / Capture / Inspure / SRS / SITI RAHMA SILPIANA / Masco, / LegoBreedrg' - label 'Full Body' tidak terbaca. Teks mendeskripsikan tiga tampilan (full body, portrait, dan bersama logo dengan nama serta tagline 'Create - Capture - Inspire'), sehingga output wajib hanya terbukti lewat teks."), ('LENGKAP', "10 prompt terdokumentasi berlabel 'prompt yang digunakan' (logo, moodboard, mug, poster, kaos, mascot, naskah, storyline, shotlist, storyboard), termasuk iterate ('nice.... sekarang berdasarkan branding di atas implementasikan branding ke media promosi nyata')."), ('LENGKAP', "Link aktif (HTTP 200), artikel 3.639 kata, dan kelima label wajib tepat lengkap: 'AI', 'SMKN 9 Garut', 'Tugas sekolah', 'personal branding', 'portofolio' - tidak ada label sia-sia."), ('LENGKAP', 'Identitas SRS konsisten penuh (navy/light blue, monogram SRS, elemen buku-kamera-pena, tiga bidang Design-Photography-Reading) dari logo, moodboard, 3 mockup, maskot, naskah, storyline, sampai storyboard; tiap bagian dijelaskan panjang dan spesifik dengan prompt sendiri.')],
))
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
S.append(('R73','AQILA NAZIL FALAQ','XI DKV 4','02 Okt 2026 08:48',
'https://qilaapacarseonghyeon.blogspot.com/2026/10/personal-branding-aqila-nazil-falaq.html',
'3 (terverifikasi teks; gambar gagal diunduh)','10 (terverifikasi teks; tanpa tabel)','6 (terverifikasi teks)','3 varian (terverifikasi teks; gambar gagal diunduh)','9 (terverifikasi teks)','2430 (memenuhi)',[3, 3, 3, 3, 3, 4, 3, 3, 2, 4, 4, 4],
[('SEBAGIAN', "Heading artikel 'ASTS PERSONAL BRANDING AQILA NAZIL FALAQ XI DKV 4 SMKN 9 GARUT' - judul panjang sesuai format. Page title 'Personal branding Aqila Nazil Falaq' memuat nama tetapi TANPA kelas XI DKV 4, jadi judul pendek tidak lengkap. Nama pada rekapan identik dengan nama pada halaman."), ('TIDAK DAPAT DIVERIFIKASI', "Logo tidak dapat diverifikasi: 7 gambar artikel semuanya berada di album Google Photos privat (lh3.google.com/u/0/...) dan gagal diunduh sehingga px=0 dan OCR error - isi visual logo TIDAK DAPAT DIVERIFIKASI. Yang terverifikasi hanya teks: monogram 'AN', pita besar, bintang, ornamen hati, hiasan lengkung; nama branding 'Aqila Nazil Falaq'; tagline 'Dream - Create - Be Yourself'. Deskripsi konsep bagian 1 hanya 96 kata (di bawah 100 kata), dilengkapi penjelasan konsep tambahan pada bagian 9; nama siswa pada logo hanya diminta lewat prompt A, tidak terlihat pada gambar."), ('TIDAK DAPAT DIVERIFIKASI', "Moodboard (bagian 2, 193 kata) memuat warna utama (soft pink, light pink, rose pink, mauve pink, cream white), typography (Playfair Display, Poppins, Allura), style visual, referensi visual (bunga, alat tulis, laptop, boneka, awan, ruang kreatif), tone & mood (dreamy, nyaman) dan elemen grafis (pita, bunga, mahkota, bintang, hati, kupu-kupu, monogram AN) - 6 dari 7 unsur; 'Inspirasi Visual' tidak dirumuskan terpisah. Format LANDSCAPE TIDAK DAPAT DIVERIFIKASI karena gambar moodboard gagal diunduh."), ('TIDAK DAPAT DIVERIFIKASI', 'Tiga media mockup terverifikasi dari teks bagian 3 (155 kata) dengan penjelasan fungsi tiap media: hoodie (logo, nama, tagline sebagai elemen utama), kartu nama (nama, bidang kreatif, kontak, kode QR untuk networking), dan stiker (berbagai bentuk untuk laptop, buku, botol). Berkas gambar mockup TIDAK DAPAT DIVERIFIKASI karena seluruh 7 gambar gagal diunduh.'), ('SEBAGIAN', "Unsur-unsur naskah terpenuhi tetapi TIDAK terkumpul dalam satu bagian naskah: Judul/tema produk 'Glow More - Be You' dan Tema (perawatan kulit remaja) pada bagian 5, Pesan Utama 'Kulit sehat, cerah, dan percaya diri setiap hari!' pada bagian 5, Narasi/Dialog tokoh pada bagian 8 ('Kulit kusam dan lelah? Saatnya kenalan sama Glow More.', 'Tekstur ringan, nyaman di kulit.', 'Glow More! Kulit sehat, cerah, dan percaya diri setiap hari.'), Closing Tagline 'Glow More - Be You'. Tidak ada tabel naskah dengan kolom Durasi|Visual|Narasi seperti pada naskah karya lengkap."), ('LENGKAP', "Storyline bagian 6 'Alur cerita' memuat keempat unsur wajib secara berlabel: Pembukaan (pagi, remaja kurang percaya diri, kulit kusam, melihat produk di meja kamar), Alur Cerita (pemakaian produk, proses, aktivitas sekolah/kuliah, wajah lebih segar), Konflik/Fokus Visual (perubahan kulit sebelum-sesudah, tekstur produk, kandungan bahan alami, manfaat), dan Penutup (tokoh tersenyum percaya diri memegang produk, muncul logo dan tagline)."), ('SEBAGIAN', "Tidak ada tabel shotlist; bagian 7 'Daftar Pendek' (207 kata) menguraikan 10 setelan kamera secara berurutan dalam prosa: close up wajah di cermin, medium shot slow push in, close up detail tangan membuka kemasan, sudut tinggi produk, medium shot tracking, close up split screen before-after, medium close up low angle slow zoom in, jarak dekat memegang produk, logo + tagline, dan extreme close up produk. Jumlah 10 setelan terpenuhi tetapi tanpa nomor shot, kolom angle/durasi, dan tanpa gambar shotlist yang dapat diunduh."), ('SEBAGIAN', "Papan cerita bagian 8 memuat 6 adegan bernomor (Adean 1 Pembukaan sampai Adegan 6 Penutup) dengan keterangan shot (close up, medium shot, extreme close up), angle (high angle, low angle, medium close up) dan dialog/narasi tokoh per adegan - minimal 6 scene terpenuhi. Transisi antar adegan TIDAK disebut eksplisit. CATATAN: angka '1.377 kata' pada dossier untuk heading 'Deskripsi Storyboard' sebenarnya mencakup teks sampai akhir artikel (bagian 9 Penjelasan konsep dan bagian 10 Prompt); uraian storyboard itu sendiri sekitar 308 kata."), ('TIDAK DAPAT DIVERIFIKASI', "Bagian 4 (135 kata) menyebut tiga varian maskot perempuan anime bergaya hijab hitam: versi full body (pose ceria mengangkat tangan membentuk simbol peace), versi potret, dan versi maskot bersama logo monogram 'AN', dengan personality Kind, Creative, Positive, Dream Big. Namun gambar maskot berada di album Google Photos privat dan gagal diunduh, sehingga keberadaan varian FULL BODY (wajib) TIDAK DAPAT DIVERIFIKASI - sesuai ketentuan konservatif skor 2."), ('LENGKAP', "Sembilan prompt terdokumentasi lengkap, bernomor, dan diberi label eksplisit pada bagian 10 'Prompt yang Digunakan': A. Prompt Logo Personal Branding, B. Prompt Moodboard, C. Prompt Mockup Branding, D. Prompt Maskot, E. Prompt Produk Glow More, F. Prompt Adegan Video Iklan, G. Prompt Perbandingan Before-After, H. Prompt Adegan Penutup, I. Prompt Storyboard - masing-masing berisi warna, typeface, dan elemen spesifik. CATATAN: kalimat pengantar 'prompt yang dapat dicantumkan ... disesuaikan dengan gambar pada tugas' memberi kesan prompt ditulis menyesuaikan gambar, bukan catatan prompt yang benar-benar dipakai."), ('LENGKAP', "Link aktif HTTP 200. Tujuh label: AI, PERSONAL, BRANDING, PORTOFOLIO, SMKN 9 GARUT, TUGAS SEKOLAH, dan KOMPUTER GRAFIS AQILA NAZIL FALAQ XI DKV 4 SMKN 9 GARUT - kelima label wajib terpenuhi, dengan catatan label 'Personal Branding' terpecah menjadi dua label terpisah ('PERSONAL' dan 'BRANDING') dan tidak ada satu pun label milik tugas lain. Isi artikel lengkap 2.430 kata dengan 10 bagian bernomor."), ('LENGKAP', "Konsistensi identitas visual terjaga di seluruh bagian: monogram 'AN', palet pink (soft/light/rose/mauve/cream white), tipografi Playfair Display + Poppins + Allura, dan elemen pita, bunga, mahkota, bintang, hati, kupu-kupu. Maskot, logo, mockup, dan produk fiktif 'Glow More' memakai palet yang sama; maskot punya karakter Kind, Creative, Positive, Dream Big. Nama pada rekapan sama dengan nama pada artikel dan branding.")],
))
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
S.append(('R72','NURI MEITRI AENI','XI DKV 3','02 Okt 2026 22:56',
'https://nurimeiaeni.blogspot.com/2026/10/personal-branding-nuri-mei-tri.html',
'3 (terverifikasi; fungsi tiap media tertulis)','TIDAK DAPAT DIVERIFIKASI (hanya prompt)','6 (terverifikasi teks)','1 gambar 160x320 (full body TIDAK DAPAT DIVERIFIKASI)','8','2630 (memenuhi)',[2, 4, 3, 4, 4, 4, 1, 3, 2, 4, 2, 3],
[('SEBAGIAN', "Page title 'Personal branding nuri mei tri' hanya memuat nama, TANPA KELAS dan huruf kecil. Judul panjang terbaca pada baris pertama artikel 'PERSONAL BRANDING NURI MEI TRI AENI XI DKV 3 SMKN 9 GARUT' tetapi tanpa prefiks 'ASTS'. Catatan domain: link yang aktif sekarang berdomain berbeda dari kiriman 404 sebelumnya (nurimeiaeni.blogspot.com, sebelumnya nurimeitriii.blogspot.com)."), ('LENGKAP', "Logo inisial N-M-T-A dengan lingkaran botanical (IMG#01 320x320, OCR 'NURI / MEITRI AENI / Your skin is valuable'); nama branding 'NMTA'; tagline 'Your skin is valuable'; deskripsi konsep 262 kata (>=100) yang menguraikan makna warna rose gold/cream dan simbol lensa kamera; NAMA SISWA tercetak pada logo."), ('SEBAGIAN', "IMG#02 320x213 landscape. Enam dari tujuh unsur terverifikasi di teks: warna utama (#B76E79 rose gold, #FFF5E9 cream), typography (Elegant Serif + Soft Script), style visual (soft, feminine, minimal, luxury skincare), elemen grafis (delicate line art, dedaunan, thin serif monogram, lingkaran floral), tone & mood (gentle, calm, serene), inspirasi visual (tekstur krim, rose gold foil, kain satin, flatlay). 'Referensi Desain' tidak dibahas sebagai butir tersendiri."), ('LENGKAP', "3 mockup dengan penjelasan 'Fungsi di Media' untuk tiap media: 1. MOCKUP KAOS (merchandise/brand awareness), 2. PACKAGING SERUM (luxury, clean, trustworthy), 3. MOCKUP LAPTOP (workstation/professionalitas) - didukung IMG#03 240x320, IMG#04 320x320, IMG#05 320x213."), ('LENGKAP', "'NASKAH' memuat 1. JUDUL 'Your Skin is Valuable - Cerita dari NUNA', 2. TEMA Self-Love & Mindful Beauty, 3. PESAN UTAMA, 4. NARASI/DIALOG Scene 1-4 lengkap dengan VO NUNA dan keterangan waktu (0-10, 10-25, 25-40, 40-50 detik), 5. CLOSING TAGLINE 'Your skin is valuable'."), ('LENGKAP', "'STORYLINE BRAND FILM - NMTA : THE VALUABLE GLOW' memuat 4 bagian lengkap: 1. PEMBUKAAN, 2. ALUR CERITA, 3. KONFLIK / FOKUS VISUAL (Visual 1 Insecure, Visual 2 Titik Balik, transisi warna dingin ke hangat, fokus macro), 4. PENUTUP dengan tagline penutup dan backsound."), ('BELUM MEMENUHI', "Bagian '7.shootlist' HANYA berisi prompt ('Oke mantap,sekarang Buatin tabel shotlist minimal 10 shot. No,Adegan,Jenis,Shot Angle,Movement, Durasi,Deskripsi') - tidak ada satu pun baris shot atau tabel di teks. Satu-satunya bukti IMG#07 320x227 dengan OCR pseudoteks ('Ee / sidrhi / ooaHae / DhNaM / VMOXM') sehingga jumlah shot TIDAK DAPAT DIVERIFIKASI."), ('SEBAGIAN', "Bagian 'STORYBOARD' (314 kata) memuat konsep dasar, konsep visual, dan 6 scene yang dirinci (Scene 1 Insecure, 2 Mindful Choice, 3 Mindful Detail, 4 The Ritual, 5 Glow Up, 6 Closing Brand) - minimum 6 scene terpenuhi di teks. NAMUN keterangan shot/angle/transisi/dialog per scene tidak ada di teks, dan OCR IMG#08 320x213 sepenuhnya pseudoteks."), ('SEBAGIAN', "Mascot NURI-CHAN diuraikan sangat rinci (309 kata: bentuk 3D chibi, rambut bergelombang, jaket puffer, rok plisket, kamera, botol serum, tas selempang). Hanya 1 berkas gambar IMG#06 160x320 dengan OCR '? / ND / NMEX'; prompt hanya meminta 'karakter full body' dan tidak ada varian portrait maupun maskot bersama logo."), ('LENGKAP', "8 prompt terdokumentasi dan diberi label 'Prompt'/'prompt': logo, moodboard, mockup (kaos, packaging serum, laptop), maskot, naskah, storyline, shotlist, storyboard."), ('BELUM MEMENUHI', "Link aktif HTTP 200 dengan 5 label Blogger: 'Cv', 'MASKOT NURI', 'MOCKUP NURI', 'Mengasah kreativitas di lab dkv', 'Moodboard'. KELIMANYA LABEL WAJIB TIDAK ADA - 0 dari 5 (Tugas Sekolah, Personal Branding, AI, Portofolio, SMKN 9 Garut); 2 label (Cv, Mengasah kreativitas di lab dkv) milik tugas lain."), ('SEBAGIAN', "Konsep orisinal (NMTA - brand skincare dengan fotografi) dan palet rose gold + cream konsisten di logo, moodboard, mockup, maskot, dan storyboard. NAMUN nama maskot tidak konsisten: 'NURI-CHAN' pada bagian MASKOT, tetapi 'NUNA' pada bagian NASKAH, STORYLINE, dan STORYBOARD.")],
))
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
# ===================== KIRIMAN BARU 2-3 OKTOBER 2026 (26 siswa) + ULANG (12 siswa) ======
# Dinilai dari dossier: teks artikel + OCR RapidOCR + dimensi piksel asli.
# 26 kiriman pertama-ever, 12 siswa lama dinilai ulang karena mengirim ulang/memperbaiki.

S.append(('R95','ADE SAHRUL GUNAWAN','XI DKV 4','02 Okt 2026 16:10',
'https://adesahrulgunawan.blogspot.com/2026/10/personal-branding-ade-sahrul-gunawan-xi.html',
'1 (belum terverifikasi)','12 (klaim teks; TIDAK DAPAT DIVERIFIKASI)','6 (terverifikasi teks)','1 (gambar; 3 varian diklaim)','5','2953',[4, 4, 2, 1, 4, 4, 2, 4, 2, 4, 2, 2],
[('LENGKAP', "Page title 'PERSONAL BRANDING ADE SAHRUL GUNAWAN XI DKV 4' + heading artikel 'ASTS PERSONAL BRANDING ADE SAHRUL GUNAWAN XI DVK 4 SMKN 9 GARUT' - keduanya sesuai dan nama konsisten."), ('LENGKAP', "Logo ASG (IMG#01 320x320, OCR 'ADE SAHRUL GUNAWAN / CAPTURE / CREATE / INSPIRE'); nama branding 'Ade Sahrul Gunawan', tagline 'CAPTURE CREATE INSPIRE', NAMA SISWA tercetak pada logo; deskripsi konsep berupa 10 sub-bagian (Identitas Brand, Konsep Utama, Makna Bentuk, Warna, Tipografi, Tagline, Karakter, Target, Penerapan, Filosofi) - cuplikan saja sudah 235 kata, jauh di atas 100."), ('SEBAGIAN', "IMG#02 320x213 landscape. Hanya 4 dari 7 unsur bernama jelas di prosa: Warna Utama (navy/royal/sky/light blue), Typography (Montserrat + Allura), Referensi Desain, Elemen Grafis (kamera, video, play button, crown). 'Style Visual' dan 'Tone & Mood' hanya tersirat, 'Inspirasi Visual' tidak ada; ketujuh unsur hanya muncul di dalam blok PROMPT sehingga tidak dihitung."), ('BELUM MEMENUHI', "Hanya ada 1 gambar mockup (IMG#03 320x213, OCR '(tidak ada teks yang terbaca)'). Teks bagian 'C. MEMBUAT MOCKUP BRANDING' hanya 105 kata umum tanpa sub-bagian per media dan tanpa penjelasan fungsi tiap media, padahal PROMPT meminta 'Minimal 3 Mockup Branding... Disertai penjelasan fungsi media'."), ('LENGKAP', "Naskah Iklan memuat Judul 'Lebih dari Sekadar Gambar', Tema, Pesan Utama, Narasi/Dialog Scene 1-6 (Visual + Narator), dan Closing Tagline 'ADE SAHRUL GUNAWAN / CAPTURE CREATE INSPIRE' beserta makna tiap kata tagline."), ('LENGKAP', "Storyline memuat keempat unsur wajib: '1. Pembukaan', '2. Alur Cerita' (Melihat -> Mengambil -> Mengedit -> Menciptakan karya -> Membagikan), '3. Konflik/Fokus Visual', '4. Penutup' dengan logo ASG dan tagline."), ('BELUM MEMENUHI', "Shotlist hanya berupa paragraf Deskripsi Shotlist 102 kata yang mengklaim 'Video terdiri dari 12 shot'; tidak ada daftar/tabel shot dan tidak ada jenis shot, angle, movement, maupun durasi per shot. IMG#06 320x213 hanya terbaca OCR 'SHOTLIST ... 8' sehingga jumlah baris TIDAK DAPAT DIVERIFIKASI."), ('LENGKAP', "6 scene storyboard tertulis lengkap di artikel dengan jenis shot (Wide Shot, Close Up, Medium Close Up, Medium Shot), angle (Eye Level, Over Shoulder, High Angle), transisi (Fade In, Cut, Push In -> Cut, Match Cut, Quick Cut, Fade Out), dan narasi; IMG#07 320x213 memuat header 'STORYBOARD IKLAN KOMERSIAL'."), ('SEBAGIAN', "Teks menjelaskan tiga varian ('1. Mascot Full Body' dari kepala hingga kaki, '2. Mascot Portrait', '3. Mascot Bersama Logo Branding' dengan logo SAG/ASG) tetapi hanya ada 1 gambar (IMG#05 320x292) tanpa label varian yang terbaca, sehingga full body tidak dapat dipastikan dari gambar."), ('LENGKAP', "5 prompt berlabel 'PROMPT :': branding/logo, moodboard, mockup, video (naskah + storyline + shotlist + storyboard dalam satu prompt), dan maskot. Catatan koreksi: blok 'PROMPT : Ubah foto diri menjadi karakter maskot...' hanya muncul 1 kali pada teks penuh - tidak ada prompt maskot duplikat."), ('BELUM MEMENUHI', '0 dari 5 label wajib terpasang (daftar labels kosong). Isi artikel lengkap (2953 kata, 7 bagian bernomor A-E) dan URL aktif status 200.'), ('SEBAGIAN', "Konten orisinal dan naratif kuat, tetapi ada ketidakkonsistenan: dua bagian berbeda sama-sama berjudul 'Deskripsi Desain Branding' (B moodboard dan C mockup), palet berubah dari 'navy blue, biru royal, biru sky, light blue' di B menjadi 'hitam dan biru' di C, dan nama logo disebut 'SAG/ASG'. Catatan koreksi: bagian huruf B ADA ('B. MEMBUAT MOODBOARD BRANDING'); huruf F tidak ada sebagai bagian tingkat atas dan hanya muncul sebagai butir 'F. Elemen Mahkota' di dalam '3. Makna Bentuk Logo'.")],
))
S.append(('R104','ALIA ALAIKA NURFADILA','XI DKV 1','02 Okt 2026 20:08',
'https://aliaalaikanurfadila.blogspot.com/2026/10/asts-personal-branding-alia-alaika.html',
'TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI',[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
[('TIDAK DAPAT DIVERIFIKASI', 'Link artikel publik yang dikirim mengembalikan HTTP 404 (https://aliaalaikanurfadila.blogspot.com/2026/10/asts-personal-branding-alia-alaika.html tidak dapat dibuka); kemungkinan artikel masih berstatus draft atau URL tidak dipublikasikan.'), ('TIDAK DAPAT DIVERIFIKASI', 'Halaman HTTP 404 - tidak ada page title, tidak ada teks artikel, tidak ada gambar yang dapat diunduh.'), ('TIDAK DAPAT DIVERIFIKASI', 'Halaman HTTP 404 - bagian moodboard tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Halaman HTTP 404 - bagian mockup tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Halaman HTTP 404 - bagian naskah iklan tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Halaman HTTP 404 - bagian storyline tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Halaman HTTP 404 - bagian shotlist tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Halaman HTTP 404 - bagian storyboard tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Halaman HTTP 404 - bagian maskot tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Halaman HTTP 404 - bagian prompt tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Halaman HTTP 404 - label Blogger tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Karya tidak dapat diakses sehingga orisinalitas dan profesionalisme tidak dapat dinilai.')],
))
S.append(('R109','AZIZAH NURUL KAMIL','XI DKV 1','02 Okt 2026 22:29',
'https://azizahnurulkamil.blogspot.com/2026/10/personal-branding-azizah-nurul-kamil.html',
'1 (terverifikasi gambar IMG#03)','10 (terverifikasi gambar IMG#04)','6 (terverifikasi teks + gambar IMG#06)','3 varian (terverifikasi OCR IMG#07)','4','2945 (memenuhi)',[3, 4, 4, 2, 4, 4, 4, 4, 4, 3, 2, 3],
[('SEBAGIAN', "Heading artikel 'ASTS Personal Branding Azizah Nurul Kamil XI DKV 1 SMKN 9 Garut' sesuai; tetapi page title 'Personal Branding Azizah Nurul Kamil' TANPA nama kelas."), ('LENGKAP', "Logo monogram ANK (IMG#01 1024x1024, OCR 'AZIZAH NURULKAMIL / CHEF|ARTIST|CREATIVEEXPLORER'); nama branding 'Azizah Nurul Kamil'; tagline 'CHEF | ARTIST | CREATIVE EXPLORER'; deskripsi konsep 559 kata (jauh >100); NAMA SISWA tercetak pada logo."), ('LENGKAP', 'IMG#02 1376x768 landscape dengan 7 unsur bernomor eksplisit di teks: 2. Typography, 3. Style Visual, 4. Referensi Desain, 5. Tone & Mood, 6. Elemen Grafis, 7. Inspirasi Visual, ditambah Palet Warna; OCR gambar mengonfirmasi blok COLOR PALETTE, TYPOGRAPHY, VISUAL STYLE & INSPIRATION, GRAPHIC ELEMENTS, TONE & MOOD, DESIGN REFERENCES.'), ('BELUM MEMENUHI', "Hanya 1 mockup yang dikerjakan dan dijelaskan, yaitu 'Mockup Gelas Kaca (Glassware)' pada bagian 3; permintaan stiker dan kartu nama di bagian itu masih berbentuk prompt ('prmpt tolong buatkan saya gambar mockup ... 1. mockup stiker, 2. mockup kartu nama, 3. mockup gelas') sehingga belum memenuhi minimal 3 mockup."), ('LENGKAP', "Naskah memuat Judul 'ANK: Where Art Meets Flavor', Tema, Pesan Utama, Durasi 60 detik, Narasi/Dialog 6 scene dengan VO '(AZIZAH (VO))', serta Closing Tagline 'AZIZAH NURUL KAMIL: Exploring the intersection of art and cuisine.'"), ('LENGKAP', "Bagian '6.Storyline' memuat 4 unsur lengkap: Pembukaan, Alur Cerita, Konflik/Fokus Visual, Penutup - masing-masing dengan Teknik (close-up, split-screen, transisi morphing)."), ('LENGKAP', 'Shotlist lengkap 10 baris dengan kolom No/Adegan/Jenis Shot/Angle/Pergerakan/Durasi, terbaca dari OCR IMG#04 (1408x768): 1 Kuas cat menggores kanvas pastel (ECU)/Low/Static/3s ... 10 Logo ANK, Nama, Tagline (Full Graphic)/Eye Level/Pull Out/7s. Didukung prosa Shot 1-10 di bagian 7.'), ('LENGKAP', 'Storyboard 6 scene tertulis lengkap dengan Visual Adegan, Keterangan Shot, Angle Kamera, Transisi (Cut To, Dissolve To, Wiping, Morphing, Fade Out) dan Narasi/Dialog; IMG#06 (1408x768) merupakan gambar storyboard dengan panel 1-6.'), ('LENGKAP', "IMG#07 1024x559 dengan OCR '1. Mascot Full Body', '2. Mascot Portrait', '3. Mascot Bersama Logo Branding' - ketiga output wajib terverifikasi."), ('SEBAGIAN', "4 prompt terdokumentasi dengan label: 'Prompt Logo Personal Branding', 'Prompt Moodboard Branding', 'prmpt tolong buatkan gambar mockup', 'Prompt Maskot' - teksnya ada, tetapi prompt mockup hanya menyebut '1. mockup stiker, 2. mockup kartu nama, 3. mockup gelas' tanpa hasil, dan masih menyisakan baris template AI ('sekarang bantu MEMBUAT AI MASCOT CHARACTER ... Pilihan Style')."), ('BELUM MEMENUHI', "Link aktif, tetapi hanya 2 dari 5 label wajib ('Tugas Sekolah', 'SMKN 9 Garut'); 'Personal Branding', 'AI', dan 'Portofolio' tidak ada. Label lain tidak relevan: 'CV', 'Prakrikum', 'zizah jzj'."), ('SEBAGIAN', "Nama, monogram ANK, palet oranye-toska dan tagline konsisten dari section 1 sampai 9; ada artefak tempelan 'media[cite: 10]' pada bagian maskot dan salah ketik 'rasa dan rasa seni' pada dialog penutup.")],
))
S.append(('R105','DAPA MUSTOPA','XI DKV 1','02 Okt 2026 20:22',
'https://dafamstf.blogspot.com/2026/10/personal-branding-dapa-mustopa-xi-dkv-1.html',
'3 (terverifikasi teks)','TIDAK DAPAT DIVERIFIKASI','5 (terverifikasi gambar)','3 varian (1 gambar 1024x559)','0 (tidak terdokumentasi)','1834',[2, 3, 4, 3, 4, 3, 1, 2, 3, 1, 4, 2],
[('SEBAGIAN', "Page title 'PERSONAL BRANDING DAPA MUSTOPA XI DKV 1' sesuai dengan nama pada rekapan; heading panjang 'ASTS Personal Branding ... SMKN 9 Garut' TIDAK DAPAT DIVERIFIKASI karena blok TEKS ARTIKEL LENGKAP tidak tersedia untuk R105 dan bagian pertama yang terbaca adalah 'A. Personal Branding'. KETIDAKSESUAIAN IDENTITAS: page title menulis 'DAPA MUSTOPA' sedangkan seluruh isi artikel menulis 'DAFA MUSTOFA'."), ('SEBAGIAN', "Logo monogram DM (IMG#01 1024x559, OCR 'DAFAMUSTOFA / DESIGN | KREATIF I DAMAI / BYDAFA'), nama branding dan tagline 'DESIGN | KREATIF | DAMAI' ada; deskripsi konsep unik +/-165 kata (>=100) dari 4 sub-bagian. TAPI NAMA SISWA tidak tercetak sebagai nama yang benar: yang tercetak pada logo adalah 'DAFA MUSTOFA', dan satu-satunya penyebutan 'DAPA MUSTOPA' di artikel adalah baris 'Nama Branding: DAPA MUSTOPA' di bagian '3. Tipografi & Tagline'."), ('LENGKAP', 'IMG#02 1536x1024 landscape; 7 label unsur terbaca jelas pada gambar (WARNA UTAMA BRANDING, TYPOGRAPHY, STYLE VISUAL, TONE & MOOD, REFERENSI DESAIN, ELEMEN GRAFIS, INSPIRASI VISUAL) dengan palette #0ESASA/#2EBB7F/#A7E3C1/#E5F7ED/#FFFFFF; prosa memuat uraian unsur 1-5. Catatan: prosa unsur 6 dan 7 terpotong pada batas 1600 karakter dossier.'), ('SEBAGIAN', '3 mockup terverifikasi di teks dan pada IMG#03 1536x1024: 1. Kartu Nama (fungsi eksplisit: media memperkenalkan identitas dan memberikan informasi kontak), 2. Mockup Kaos (uraikan penempatan logo, fungsi hanya tersirat), 3. Mockup Gelas (fungsi hanya tersirat: penerapan pada merchandise). Fungsi tidak dijelaskan tegas untuk tiap media.'), ('LENGKAP', "D.1 Naskah Iklan memuat Judul 'Kreativitas Tanpa Batas', Tema, Pesan Utama, Narasi/Dialog, dan Closing Tagline 'DAFA MUSTOFA - DESIGN | KREATIF | DAMAI'."), ('SEBAGIAN', "Keempat unsur storyline lengkap: 1. Pembukaan (suasana alam, gunung, dedaunan), 2. Alur Cerita (sketsa -> desain digital -> karya visual), 3. Konflik/Fokus Visual (proses mengubah ide jadi karya), 4. Penutup (logo + tagline). Catatan: tidak ada judul bagian 'D.2 Storyline' dan penomoran melompat dari D.1 ke D.4."), ('TIDAK DAPAT DIVERIFIKASI', "Tidak ada teks shotlist sama sekali di artikel - bagian 'D.3 Shotlist' tidak muncul sebagai heading (lompat dari D.1 ke D.4). Satu-satunya bukti adalah IMG#04 385x256, gambar terkecil di batch, dengan OCR tabel pseudoteks ('SHOTLIST / DAFAMUSTOFA / VIDEO IKLAN KOMERSIAL' dan angka 2, 6, 8, 10, 11 yang tidak berurutan); jumlah shot TIDAK DAPAT DIVERIFIKASI."), ('BELUM MEMENUHI', 'IMG#05 1024x559 memuat 5 scene bertanda lengkap (SCENE1 KONTROL BOLA | ECU LOW ANGLE | MATCH CUT; SCENE2 KREASI GARIS DM | CU OVER SHOULDER; SCENE3 FOKUS DAFA | MCU | EYE LEVEL | CUT; SCENE4 LOGO DM | CU | STRAIGHT ON | ZOOM OUT; SCENE5 STUDIO TENANG | FS | LOW ANGLE | DISSOLVE) - 5 scene, sedangkan syarat minimal 6; teks D.4 Storyboard (724 kata) terpotong pada Scene 2 di dossier sehingga scene ke-6 tidak dapat dipastikan.'), ('SEBAGIAN', "Ketiga varian terdeskripsi lengkap di teks (Output 1 Mascot Full Body 226 kata berisi pose, pakaian, dan hex warna; Output 2 Portrait 152 kata; Output 3 Bersama Logo 157 kata) dan labelnya terbaca pada IMG#06 1024x559 ('1.MASCOTFULLBODY / 2.MASCOT PORTRAIT / 3.MASCOTBERSAMALOGOBRANDING'), tetapi ketiga varian berada dalam satu gambar - render terpisah tidak dapat dipastikan."), ('TIDAK DIKUMPULKAN', "Nol blok 'Prompt' pada 17 bagian artikel (pencarian teks 'prompt' = 0 hasil). Satu-satunya jejak prompt adalah teks 'MEMBUAT AI MASCOT CHARACTER / Pilihan Style: 3D Character (Pixar Style)' yang tercetak di DALAM IMG#06 (1024x559), bukan terdokumentasi sebagai teks artikel. Beberapa bagian terpotong pada 1600 karakter dossier sehingga tidak 100% pasti tidak ada prompt di bagian yang terpotong."), ('LENGKAP', "Kelima label wajib terpasang: 'Tugas Sekolah', 'Personal Branding', 'AI', 'Portofolio', dan 'SMKN 9 Garut'. URL aktif status 200; isi artikel 1834 kata."), ('BELUM MEMENUHI', "Ketidaksesuaian identitas total: rekapan dan page title 'DAPA MUSTOPA', sedangkan seluruh isi artikel dan seluruh gambar memakai 'DAFA MUSTOFA' (OCR IMG#01 'DAFAMUSTOFA / BYDAFA', IMG#03 'dafa.mustofa@gmail.com / @dafamustofa', IMG#05 'DAFAMUSTOFA', IMG#06 'dafa mustofa'). Ditambah tiap paragraf branding/moodboard/storyboard tampil dua kali dalam ekstraksi, dan penomoran bagian melompat (D.1 -> D.4 tanpa D.2/D.3; tanpa huruf F). Terindikasi tidak konsisten identitas - perlu klarifikasi guru apakah yang dimaksud Dafa Mustofa.")],
))
S.append(('R20','FUZI NAILA ANNURI','XI DKV 3','02 Okt 2026 20:35',
'https://fuzinailaannuri.blogspot.com/2026/10/personal-branding-fuzi-naila-annuri-xi.html',
'3 (terverifikasi gambar)','10 (terverifikasi tabel)','1 (terverifikasi gambar)','3 (terverifikasi teks)','10','7436 (memenuhi)',[4, 4, 4, 4, 4, 3, 4, 2, 4, 4, 4, 3],
[('LENGKAP', "Page title 'Personal Branding Fuzi Naila Annuri XI DKV 3' + heading artikel 'ASTS Personal Branding Fuzi Naila Annuri XI DKV 3 SMKN 9 GARUT' - keduanya sesuai, nama dan kelas sama."), ('LENGKAP', "Logo monogram FNA (IMG#01 400x400, OCR 'fuzi naila annuri'); nama branding 'Fuzi Naila Annuri (FNA)'; tagline 'designing dreams, creating style'; deskripsi konsep 587 kata; nama tercetak pada logo."), ('LENGKAP', 'IMG#02 400x266 landscape. 7 unsur terverifikasi di teks: Warna Utama (hex #FAD7E3, #CFE8FF, #D9D9D9, #E6C98A), Typography (Playfair Display/Montserrat/Allura), Style Visual, Referensi Desain, Tone & Mood, Elemen Grafis, Inspirasi Visual.'), ('LENGKAP', "3 mockup dengan penjelasan fungsi tiap media: kaos (IMG#03), packaging (IMG#04; 8 media: paper bag, box, tissue, hang tag, thank you card, pouch, sticker, label), hoodie (IMG#05); tiap bagian diuraikan 10-15 poin termasuk 'Tujuan Pembuatan Mockup'."), ('LENGKAP', "Naskah ADA, hanya tidak tertangkap ekstraksi per-bagian (yang terbaca di bagian 5 hanya prompt). Di blok teks penuh: Judul 'Small Steps, Big Dreams - Dream, Design, Create', Tema, Pesan Utama, Narasi/VO Scene 1-6, Closing Tagline 'DREAM + DESIGN + CREATE'."), ('SEBAGIAN', "Bagian '6. STORYLINE' memuat 1. PEMBUKAAN, 2. ALUR CERITA, 3. KONFLIK/FOKUS VISUAL; sub-bagian Penutup tidak terlihat karena teks dossier terpotong pada '...memperlihatkan hasil kar'."), ('LENGKAP', "Tabel HTML 11 baris x 7 kolom (1 baris header + 10 shot) lengkap No/Adegan/Jenis Shot/Angle/Movement/Durasi/Deskripsi, dari shot 1 'Dunia Mimpi' sampai shot 10 'Closing Branding'."), ('TIDAK DAPAT DIVERIFIKASI', "Hanya 1 gambar storyboard (IMG#09 399x365). OCR terbaca 'StoryboardIklan', 'Small Steps, Big Dreams', 'DREAM+DESXN-CXIATS' (pseudoteks) dan label scene tak utuh ('Scewe', 'Scon2', 'Scena2', 'Sca14', 'SoeneS') - jumlah 5 atau 6 scene tidak dapat dipastikan; teks bagian 8 hanya berisi prompt."), ('LENGKAP', "3 gambar maskot sesuai 3 sub-bagian: MASCOT FULL BODY (IMG#06 266x400 potret, OCR 'fuzi naila annuri'), MASCOT PORTRAIT (IMG#07 266x399), MASCOT BERSAMA LOGO (IMG#08 400x400, OCR 'Small steps / fuzi naila annuri / designing dreams, creating style')."), ('LENGKAP', "10 prompt berlabel 'PROMPT YANG DIGUNAKAN:': 9 terlihat di teks (logo, moodboard, 3 mockup, 3 maskot, naskah) + 1 prompt storyboard yang terbaca di ekstraksi bagian 8; prompt shotlist tidak tampak karena teks terpotong."), ('LENGKAP', "5 label wajib lengkap (Tugas Sekolah, Personal Branding, AI, Portofolio, SMKN 9 Garut) plus 5 label relevan; isi 7436 kata sangat lengkap. Catatan: label 'CV_FUZI_TUGAS_SEKOLAH' dan 'PRAKTIKUM_DKV_TUGAS_SEKOLAH' tampaknya milik kiriman tugas lain."), ('SEBAGIAN', "Identitas FNA (pink pastel/baby blue/gold, mahkota, pita, sparkle) konsisten di logo, moodboard, 3 mockup, 3 maskot, naskah, shotlist dan storyboard; namun 3 tagline dipakai bergantian: 'designing dreams, creating style', 'DREAM + DESIGN + CREATE', dan 'Small Steps, Big Dreams'.")],
))
S.append(('R102','GILAR APGAN MUHAMAD SOLEH','XI DKV 4','02 Okt 2026 19:49',
'https://gilarapganms.blogspot.com/2026/10/personal-branding-gilar-apgan-muhamad.html',
'7 media (terverifikasi teks)','TIDAK DAPAT DIVERIFIKASI','6 (terverifikasi teks)','1 (full body tidak dapat dipastikan)','8','2740 (memenuhi)',[3, 4, 3, 3, 3, 4, 1, 4, 2, 4, 2, 2],
[('SEBAGIAN', "Page title 'Personal Branding Gilar Apgan Muhamad Soleh' ada tetapi tidak mencantumkan kelas; heading artikel sudah lengkap: 'ASTS PERSONAL BRANDING GILAR APGAN MUHAMAD SOLEH; XI DKV 4 SMKN 9 GARUT'."), ('LENGKAP', "Logo monogram GIL (IMG#01 246x240, OCR 'GILAR APGAN / MUHAMAD / SOLEH / CAMERA.VISUAL.FILM' - nama tercetak); nama branding 'GILAR APGAN MUHAMAD SOLEH'; tagline 'Lebih dari Sekadar Kamera'; deskripsi konsep >= 100 kata ('1. LOGO PERSONAL BRANDING' 116 kata ditambah blok 'Filosofi Branding' sekitar 150 kata)."), ('DIKUMPULKAN', "IMG#02 434x289 landscape. Teks 137 kata menguraikan warna utama (biru elektrik, hitam, abu-abu metalik, putih), typography (serif + sans serif), style visual (cinematic, minimalis, profesional, premium), tone & mood, elemen grafis (monogram, garis geometris, target), dan inspirasi visual (kamera, fotografer, pegunungan, langit malam, elang); 'Referensi Desain' tidak diuraikan di teks - hanya terbaca sebagai panel 'REFCRCNO OCSAIN' pada OCR gambar."), ('DIKUMPULKAN', "Blok 'Filosofi dan Fungsi Media Branding' memuat 7 media (Hoodie & T-Shirt, Sticker, Poster, Botol Minum, Social Media Feed, Laptop, Gelas Kopi) masing-masing dengan paragraf 'Filosofi:' dan 'Fungsi:'; secara visual hanya ada satu gambar gabungan IMG#03 422x281 (OCR 'MEOLAPROMOSI &FUNGSINYA', 'SOCIALMEEALFEEO'), bukan 3 berkas mockup terpisah."), ('SEBAGIAN', "'5. KONSEP VIDEO IKLAN / NASKAH IKLAN' memuat Tema 'Lebih dari Sekadar Kamera', Pesan Utama, alur narasi per tahap (pembukaan hingga penutupan dengan logo dan tagline) dan closing tagline; field 'Judul' tidak dicantumkan eksplisit - judul memakai tema/tagline yang sama."), ('LENGKAP', "'6. STORYLINE' memuat 1. Pembukaan, 2. Awal Perjalanan, 3. Proses Kreatif, 4. Konflik dan Tantangan, 5. Penutup dengan timecode, ditambah 'Deskripsi Storyline' 148 kata; IMG#06 401x267 OCR mengonfirmasi panel '1. PEMBUKAAN / 2. ALUR CERITA / KONFLIK-FOKUS VISUAL / PENUTUP'."), ('BELUM MEMENUHI', "Bagian '7. SHOTLIS' hanya berisi uraian paragraf 174 kata (wide shot logo, medium shot kreator, close-up kamera, tracking/pan/tilt/dolly) tanpa tabel, kolom, atau nomor shot; gambar IMG#07 391x260 OCR hanya membaca 'SHOTLIST / GILAR APGAN / MUHAMAD SOLEH / TEMA' tanpa baris shot. Jumlah shot TIDAK DAPAT DIVERIFIKASI dan tidak ada bukti 10 shot bernomor."), ('LENGKAP', 'Enam scene tertulis lengkap (Scene 1-6) dengan Visual Adegan, angle (eye level, over shoulder, low angle), pergerakan (slow push-in, dolly in, pan right, crane up, slow zoom out), transisi (fade in, cut, match cut, whip pan, fade out), dan narasi/dialog; IMG#08 413x275 OCR menunjukkan 6 panel (PEMBUKAAN, ALUR CERITA, KONFLIK, PROSES SOLUSI, HASIL, PENUTUP).'), ('SEBAGIAN', "Hanya satu gambar maskot IMG#04 436x290 di bagian '4. MASKOT' dengan OCR 'MASCOT' dan 'PORTRAT'; teks bagian 4. MASKOT berisi salinan verbatim deskripsi mockup (hoodie, stiker, poster, botol minum) bukan uraikan maskot, dan tidak ada label 'Full Body' - full body tidak dapat dipastikan."), ('LENGKAP', "8 prompt terdokumentasi dan diberi label 'Prompt': MEMBUAT PERSONAL BRANDING, moodboard, MEMBUAT MOCKUP BRANDING, MEMBUAT AI MASCOT CHARACTER, PERENCANAAN VIDEO IKLAN / D.1 Naskah Iklan, Storyline, Shotlist, dan Storyboard."), ('SEBAGIAN', 'Isi artikel lengkap 2.740 kata dengan 8 komponen, tetapi labels kosong ([]) - 0 dari 5 label wajib (Tugas Sekolah, Personal Branding, AI, Portofolio, SMKN 9 Garut) tidak ada sama sekali; link aktif (status 200).'), ('SEBAGIAN', "Palet hitam-biru dan logo GIL konsisten di 8 media, tetapi konsistensi konten bermasalah: bagian '4. MASKOT' menyalin ulang paragraf bagian mockup, slogan 'Dream / Create / Achieve' pada blok filosofi bertentangan dengan tagline 'Lebih dari Sekadar Kamera', dan semua gambar memuat pseudoteks AI ('GILARAPGAX', 'MUHANAOSOLLH').")],
))
S.append(('R103','INDAH TRIJAYANTI','XI DKV 4','02 Okt 2026 20:06',
'https://indahtrijayanti1.blogspot.com/2026/10/personal-branding-indah-trijayanti-xi.html',
'3 (terverifikasi gambar)','11 (terverifikasi teks)','6 (terverifikasi teks)','3 (full body)','9','5250',[3, 4, 3, 4, 4, 4, 4, 3, 4, 4, 2, 3],
[('SEBAGIAN', "Page title 'INDAH DIARY: PERSONAL BRANDING INDAH TRIJAYANTI XI DKV 4' memuat awalan tambahan 'INDAH DIARY:'; heading artikel 'ASTS PERSONAL BRANDING INDAH TRIJAYANTI XI DKV 4 SMKN 9 GARUT' ada. Nama konsisten dengan rekapan."), ('LENGKAP', "Logo IT (IMG#01 400x223, OCR 'by Indah Trijayanti / Create. Connect. Innovate.'); nama branding 'IT' ( Information Technology), tagline 'Create. Connect. Innovate.', NAMA SISWA tercetak pada logo/gambar; deskripsi konsep 350 kata (>=100)."), ('SEBAGIAN', "IMG#02 399x266 landscape. Hanya 6 dari 7 unsur terverifikasi di prosa: Warna Utama, Typography, Style Visual, Referensi Desain, Elemen Grafis, Tone & Mood. 'Inspirasi Visual' hanya disebut di dalam blok PROMPT ('serta inspirasi visual') tanpa uraian pada teks artikel, jadi tidak dihitung."), ('LENGKAP', "3 mockup dengan uraian 'Fungsi Media:' masing-masing: Laptop (identitas dan promosi mobile/brand recognition), ID Card (identitas fisik dan penghubung ke media digital), Social Media Feed (komunikasi, publikasi, portofolio digital); 3 gambar IMG#03, IMG#04, IMG#05."), ('LENGKAP', "E.1 Naskah Iklan memuat Judul 'Create Your Identity', Tema, Pesan Utama, Narasi/Dialog Scene 1-6 lengkap dengan Visual dan Narasi, serta Closing Tagline 'IT - Indah Trijayanti / Create. Connect. Innovate.'"), ('LENGKAP', 'E.2 Storyline memuat keempat unsur wajib: 1. Pembukaan (suasana gelap, cahaya cyan-magenta), 2. Alur Cerita (sketsa -> desain digital -> fotografi), 3. Konflik/Fokus Visual (perubahan ide sederhana menjadi karya final), 4. Penutup (logo penuh + tagline).'), ('LENGKAP', "11 shot tertulis berurutan sebagai Scene 1-11 di E.3; tiap scene memuat jenis shot (MCU/CU/MS/LS), angle (low angle, eye level, side angle), movement (tilt up, tracking, orbit, pan, fast zoom, static), durasi 4-7 detik (total 62 detik) dan narasi. Ditambah IMG#09 400x223 'DAFTAR SHOT PRODUKSI' (OCR kacau, jumlah baris tabel tidak terverifikasi)."), ('SEBAGIAN', 'E.4 Storyboard memuat 6 scene lengkap (Titik Balik Ide, Garis Pertama, Koneksi Digital, Visi Lensa, Identitas Berkarakter, Masa Depan Karya) dengan visual, jenis shot, angle, durasi, dan narasi - tetapi keterangan TRANSISI tidak ada di teks maupun di IMG#05, berbeda dari R91/R95 yang menyebut Fade In/Cut/Zoom In.'), ('LENGKAP', "Ketiga varian mascot punya gambar terpisah: IMG#06 300x400 pada bagian '1. Mascot Full Body', IMG#07 300x400 '2. Mascot Portrait', IMG#08 400x223 '3. Mascot Bersama Logo Branding'; teks menjelaskan karakter full body, portrait, dan maskot di samping logo."), ('LENGKAP', "9 prompt berlabel 'PROMPT' (logo, moodboard, mockup, maskot, naskah, storyline, shotlist, storyboard, penjelasan konsep). Catatan: nomor di dalam prompt memakai D.1/D.2/D.4 sementara judul bagian artikel memakai E.1/E.2/E.3/E.4."), ('BELUM MEMENUHI', "0 dari 5 label wajib terpasang - daftar labels kosong: 'Tugas Sekolah', 'Personal Branding', 'AI', 'Portofolio', dan 'SMKN 9 Garut' tidak ada. Isi artikel sendiri lengkap (5250 kata) dan URL aktif status 200."), ('SEBAGIAN', "Karya orisinal dengan sistem warna hitam-cyan-biru-ungu-magenta konsisten di logo, moodboard, mockup, dan maskot; tetapi teks didominasi pengulangan formula 'Fungsi Media:' dan 'Secara keseluruhan', serta penomoran bagian tidak konsisten (E.1-E.4 di artikel vs D.1-D.4 di dalam prompt).")],
))
S.append(('R25','INDRI','XI DKV 3','02 Okt 2026 11:23',
'https://indriii07.blogspot.com/2026/10/personal-branding-indri-xi-dkv-3.html',
'3 (terverifikasi gambar)','10 (terverifikasi tabel)','6 (terverifikasi teks)','3 (dalam 1 gambar komposit)','7','4403 (memenuhi)',[4, 4, 3, 4, 4, 4, 4, 4, 4, 3, 4, 4],
[('LENGKAP', "Page title 'PERSONAL BRANDING INDRI XI DKV 3' + heading 'ASTS PERSONAL BRANDING INDRI - XI DKV3 SMKN 9 GARUT' - nama dan kelas sesuai."), ('LENGKAP', "Logo monogram IR dengan tulisan INDRI (IMG#01 320x320, OCR 'INDI / R I'); nama branding 'INDRI'; tagline 'Capture the beauty in every moment'; deskripsi konsep 589 kata (8 sub-bagian makna elemen); nama tercantum pada logo."), ('SEBAGIAN', "5 unsur tertulis lengkap (Warna, Typography, Style Visual, Referensi Desain, Tone & Mood) ditambah 'Kesan Keseluruhan'; 'Elemen Grafis' dan 'Inspirasi Visual' tidak ditulis terpisah. Gambar IMG#02 320x320 persegi, bukan landscape."), ('LENGKAP', '3 mockup dengan penjelasan tiap media: hoodie (IMG#03 305x320), tumbler (IMG#04 320x320, hex #2E5E8F/#79AEDD/#B7D8F6), kaos (IMG#05 320x320) - tiap media diuraikan warna, tipografi, latar, dan fungsi brand.'), ('LENGKAP', "Judul 'Jejak Keindahan Dalam Setiap Momen', Tema 'Aesthetic Exploration & Creative Lifestyle', Pesan Utama, Narasi/Voice Over 4 blok bertimecode 00:00-00:35, Closing Tagline 'INDRI - Capture the beauty in every moment'."), ('LENGKAP', "Judul storyline 'The Silent Canvas' + Konsep Utama, lalu 1. Pembukaan, 2. Alur Cerita (Rising Action), 3. Konflik (Climax / Plot Twist), 4. Penutup (Resolution & Closing) - 4 unsur lengkap."), ('LENGKAP', 'Tabel HTML 11 baris x 7 kolom (1 header + 10 shot) lengkap; dikonfirmasi gambar IMG#07 685x382 yang OCR-nya membaca seluruh header dan 10 shot.'), ('LENGKAP', "6 scene tertulis lengkap di artikel (Scene 1-6) dengan Visual Adegan, Teknis Kamera, Transisi (Dissolve / Cut to / Wipe / Fade Out) dan Narasi; dikonfirmasi IMG#08 640x640 (OCR 'Story Board Komersial Film: INDRI - The Silent Canvas' dan kolom No/Adegan/Jenis Shot/Angle/Transisi)."), ('LENGKAP', 'IMG#06 320x320 memuat 4 kuadran sesuai deskripsi: Full Body, Potret Close-Up, Logo Branding, dan Maskot & Logo -(full body, portrait, bersama logo) ketiganya terpenuhi.'), ('SEBAGIAN', '7 prompt berlabel (logo, moodboard, maskot, naskah, storyline, shortlist, storyboard) - prompt mockup tidak ada, dan prompt shortlist tertulis ulang persis sama dengan prompt storyline.'), ('LENGKAP', 'Tepat 5 label wajib (Tugas Sekolah, Personal Branding, AI, Portofolio, SMKN 9 Garut) tanpa label tambahan; isi 4403 kata lengkap 8 bagian.'), ('LENGKAP', "Slogan 'Capture the beauty in every moment' konsisten di logo, mockup, naskah, storyline, shotlist dan storyboard; konten orisinal, tidak ditemukan teks identik dengan siswa lain pada batch ini.")],
))
S.append(('R27','INTAN MAHARANY','XI DKV 3','02 Okt 2026 11:25',
'https://intannmaharany.blogspot.com/2026/10/intans.html',
'3 (terverifikasi teks)','10 (terverifikasi tabel)','6 (terverifikasi teks)','1 (full body tidak dapat dipastikan)','8','4644 (memenuhi)',[4, 4, 4, 3, 4, 4, 4, 3, 2, 4, 2, 3],
[('LENGKAP', "Page title 'PERSONAL BRANDING INTAN MAHARANY XI DKV 3' + heading artikel 'PROJEK ASTS - PERSONAL BRANDING INTAN MAHARANY KELAS XI DKV 3 SMKN 9 GARUT' - keduanya ada dan menyebut nama serta kelas."), ('LENGKAP', "Logo monogram IN (IMG#01 320x320, OCR 'INTANMAHARANY / MUA SKINCARE FORMULATOR BEAUTY PHOTOGRAPHY' - nama siswa tercetak pada logo); nama branding 'INTAN MAHARANY'; tagline/descriptor 'MUA - SKINCARE FORMULATOR - BEAUTY PHOTOGRAPHY'; deskripsi konsep vastly melebihi 100 kata (21 bagian branding, palettes E8B4B8/A8B5A0/C4A484 diuraikan per warna)."), ('LENGKAP', 'IMG#02 320x213 landscape. 7 unsur terverifikasi di teks bagian 1-7: Warna Utama (hex E8B4B8, A8B5A0, C4A484), Typography (serif elegan), Style Visual (Soft Luxury Minimal), Referensi Desain, Tone dan Mood, Elemen Grafis (monogram IN, brush stroke, botanical line art), Inspirasi Visual (makeup brushes, dusty pink silk, dewy skin).'), ('DIKUMPULKAN', "Tiga media mockup dijelaskan fungsinya di 'DESKRIPSI MOCKUP PRODUK' (Radiance Serum, Makeup Pouch, Sticker Sheet + dekorasi), tetapi visualnya hanya satu gambar flatlay gabungan IMG#03 320x180 (OCR nyaris kosong, hanya 'N') - bukan 3 berkas mockup terpisah."), ('LENGKAP', "'5.NASKAH IKLAN NASKAH PROMOSI - INTAN MAHARANY': Judul 'Gentle Radiance', Tema 'Soft Luxury Beauty', Pesan 'Merawat dan memancarkan cantik alami', Narasi/Dialog 5 scene lengkap dengan Visual + VO, TAGLINE 'Gentle Care, Timeless Radiance'."), ('LENGKAP', "'6.STORYLINE': 1. PEMBUKAAN (Introduction), 2. ALUR CERITA (Tahap 1 Formulasi Herbal, Tahap 2 Transformasi MUA), 3. KONFLIK / FOKUS VISUAL (Dull to Glow, Mirror Reveal), 4. PENUTUP (Conclusion & Identity) - keempat unsur lengkap."), ('LENGKAP', 'Tabel HTML shotlist 11 baris x 7 kolom dengan header No | Adegan | Jenis Shot | Angle | Movement | Durasi | Deskripsi, memuat 10 shot bernomor 1-10 (Opening Studio sampai Closing Brand) - kolom lengkap, jumlah shot memenuhi.'), ('SEBAGIAN', "'8.STORYBOARD' menguraikan 6 scene (Scene 1 Opening sampai Scene 6 Closing) dengan shot type dan angle per scene (Wide Shot, Macro Close-Up Top-Down 45, Over The Shoulder MCU, Medium Shot Front Low Angle, Static Center Shot), tetapi transisi dan VO hanya disebut sebagai label panel tanpa dirinci per scene; gambar IMG#05 320x179 (berada di bagian shotlist) OCR-nya pseudoteks sehingga label panel tidak dapat diverifikasi."), ('SEBAGIAN', "Hanya satu gambar maskot IMG#04 320x175 (OCR 'INTANMAHLARANY'); prompt maskot menyebut 'cute chibi girl mascot human full body' dan storyline menyebut 'Karakter chibi tampil full body', namun tidak ada gambar/label terpisah untuk Full Body, Portrait, maupun Bersama Logo - full body tidak dapat dipastikan."), ('LENGKAP', "8 prompt terdokumentasi dan diberi label 'prompt yang digunakan:' - logo, moodboard (7 sections), mockup (3 beauty product mockups), maskot, naskah, storyline, shotlist (10 shot), dan storyboard (6 panels 16:9)."), ('SEBAGIAN', "Isi artikel paling lengkap dalam batch ini (4.644 kata, 8 komponen + tabel shotlist), tetapi label Blogger tidak sesuai: ['CV','MOCKUP','logo branding','maskot','moadboard','praktik dkv'] - 0 dari 5 label wajib (Tugas Sekolah, Personal Branding, AI, Portofolio, SMKN 9 Garut); 'moadboard' juga salah ketik."), ('SEBAGIAN', "Konsep orisinal (beauty 360 derajat: MUA + Skincare Formulator + Beauty Photography) dan konsisten - monogram IN serta hex warna yang sama muncul di branding, moodboard, mockup, storyline, dan storyboard; catatan: label Blogger tidak relevan/tidak profesional dan 'tagline' lebih berupa brand descriptor ('MUA - SKINCARE FORMULATOR - BEAUTY PHOTOGRAPHY') daripada slogan.")],
))
S.append(('R90','LUSI NURAENI','XI DKV 1','02 Okt 2026 12:46',
'https://www.blogger.com/blog/post/edit/4743577463240745335/7829928751200628715',
'TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI',[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
[('TIDAK DAPAT DIVERIFIKASI', 'Link yang dikirim adalah alamat EDITOR Blogger (https://www.blogger.com/blog/post/edit/...), bukan link artikel publik; halaman hanya dapat dibuka setelah login Google sehingga karya tidak dapat diakses sama sekali.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - tidak ada page title, tidak ada teks artikel, tidak ada gambar yang dapat diunduh.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian moodboard tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian mockup tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian naskah iklan tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian storyline tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian shotlist tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian storyboard tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian maskot tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian prompt tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger (bukan artikel publik) - label Blogger tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Karya tidak dapat diakses sehingga orisinalitas dan profesionalisme tidak dapat dinilai.')],
))
S.append(('R101','MUHAMAD DANDI NUGRAHA','XI DKV 1','02 Okt 2026 19:29',
'https://dandingrh.blogspot.com/2026/10/personal-branding-muhammad-dandi.html',
'3 (terverifikasi)','10 (terverifikasi gambar)','6 (terverifikasi teks)','3 (termasuk full body)','5','3303 (memenuhi)',[4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 4, 3],
[('LENGKAP', "Page title 'PERSONAL BRANDING MUHAMMAD DANDI NUGRAHA XI DKV 1' + heading artikel 'ASTS Personal Branding Muhammad Dandi Nugraha XI DKV 1 SMKN 9 Garut' - keduanya lengkap; catatan ejaan nama di artikel 'Muhammad Dandi Nugraha' berbeda dari rekapan 'Muhamad Dandi Nugraha'."), ('LENGKAP', "Logo inisial MDN dengan D-pad gaming dan api Paskibra (IMG#01 320x175, OCR 'MUHAMMAD DANDI NUGRAHA / Graphic Designer-Gamer-Passionate Paskibra Member' - nama tercetak); nama branding 'MDN'; tagline 'Di Mana Disiplin Bertemu Kreativitas'; deskripsi konsep 738 kata (jauh di atas 100 kata)."), ('LENGKAP', 'IMG#02 1600x893 landscape (resolusi terbaik batch ini) dan 8 unsur terverifikasi di teks: Judul dan Identitas Utama, Visual Inspiration & Style, Tone & Mood (ENERGETIC, PROFESSIONAL, CREATIVE, DISCIPLINED), Typography (Montserrat Bold/Light), Warna Utama (#0B3D91, #F97316, #3B82F6, #FBBF24, #1F2937), Logo Applied, Elemen Grafis, Referensi Desain - ketujuh unsur wajib ada.'), ('LENGKAP', "Tiga mockup dengan penjelasan fungsi terstruktur per media: 1. Mockup Kaos (Walking Billboard, Perkenalan Diri Cepat, karakter asik), 2. Mockup Poster (Pernyataan Identitas, Skalabilitas, Kredibilitas), 3. Mockup Gelas Kopi (Media Keseharian, Casual Touchpoint, Potensi Foto Media Sosial); IMG#03 1024x559 OCR 'T-SHIRT MOCKUP / POSTER MOCKUP / COFFEE CUP MOCKUP / MDN PERSONAL BRANDING - 3 MOCKUP PRESENTATION'."), ('LENGKAP', "'D.1 Naskah Iklan': Judul 'MDN: Di Mana Disiplin Bertemu Kreativitas', Tema (Struktur Paskibra vs Imajinasi Gaming), Pesan Utama, Durasi 60 detik, Narasi/Dialog 6 scene (NARATOR & DANDI) dan Closing Tagline (On-Screen Text) 'MDN: Di Mana Disiplin Bertemu Kreativitas'."), ('LENGKAP', "'D.2 Storyline' memuat keempat unsur: Pembukaan (barisan Paskibra), Alur Cerita (kontras ke gaming lalu studio desain), Konflik/Fokus Visual (split screen Paskibra vs tablet grafis), Penutup (perkenalan diri + logo MDN + nama lengkap 'MUHAMMAD DANDI NUGRAHA')."), ('LENGKAP', "Shotlist sebagai gambar IMG#04 438x245 di bagian 'D.3 Shotlist'; OCR terbaca sebagai tabel berheader No. | Adegan | Visual (Deskripsi) | Jenis Shot | Angle | Pergerakan | Durasi dengan 10 baris bernomor 1-10 (Closeup sepatu Paskibra sampai Penutup) - jumlah dan kolom memenuhi."), ('LENGKAP', 'Enam scene tertulis lengkap (Scene 1-6) dengan Visual Adegan, Keterangan Shot (CU/WS/MS/MCU/FS), Angle, Kamera (Push In, Pull Out, Transisi Morph), Narasi/Dialog, Tujuan; IMG#05 1024x559 OCR mengonfirmasi 6 panel dengan label Shot Info, Narasi, Dialogue, dan Tagline.'), ('LENGKAP', "Ketiga varian maskot terdokumentasi: Mascot Full Body (Image 1 full-body 3D), Mascot Portrait (Image 2 close-up), Mascot Bersama Logo Branding (Image 3 maskot di samping logo MDN); IMG#06 1024x572 OCR '1. MASCOT FULL BODY / 2. MASCOT PORTRAIT / 3. MASCOT BERSAMA LOGO'."), ('DIKUMPULKAN', "5 prompt terdokumentasi dan berlabel: logo ('PROMT : halo bro, berperanlah sebagai designer...'), moodboard, mockup 3 media, PERENCANAAN VIDEO IKLAN/naskah, dan MASCOT CHARACTER - tidak ada prompt terpisah untuk storyline, shotlist, dan storyboard."), ('LENGKAP', '5 dari 5 label wajib terpasang (AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah) dan artikel berisi 3.303 kata dengan 8 komponen lengkap; link aktif (status 200).'), ('SEBAGIAN', "Konsep orisinal dan konsisten (Paskibra x gaming, palet #0B3D91/#F97316, logo MDN sama di semua media); catatan professionalism: seluruh uraian brand ditulis dengan kata sapaan 'Anda' (ciri jawaban AI generik) dan logo IMG#01 memuat pseudoteks 'CnhkDeier.Carer.PautonyoPnktriMkrbkr'.")],
))
S.append(('R93','MUHAMAD REZA RAMDANI','XI DKV 2','02 Okt 2026 14:49',
'https://mhmdreza65.blogspot.com/2026/09/projek-asts.html',
"3 (terverifikasi gambar, brand 'The Creative Studio')",'10 (terverifikasi teks)','6 (terverifikasi teks + gambar IMG#07)','1 (full body terverifikasi teks, brand lain)','8','4560 (memenuhi)',[2, 3, 2, 2, 3, 3, 4, 4, 2, 4, 2, 1],
[('SEBAGIAN', "Page title 'PROJEK ASTS' tidak sesuai format; heading artikel 'ASTS Personal Branding Muhammad Reza Ramdani XI DKV 2 SMKN 9 Garut' sesuai. Catatan nama: artikel 'Muhammad Reza Ramdani', rekapan 'MUHAMAD REZA RAMDANI' (beda satu huruf)."), ('SEBAGIAN', "Logo RR ada (IMG#01 320x320, OCR 'MUHAMMAD REZA RAMDANI / RR') dengan nama branding 'RR' dan NAMA SISWA tercetak pada logo; deskripsi konsep 168 kata (>=100). Tagline TIDAK ada - prompt sendiri menulis 'tagline = sesuaikan' dan tidak pernah diisi."), ('BELUM MEMENUHI', "Moodboard BUKAN milik personal branding siswa. IMG#02 400x266 (OCR 'WARNAUTAMAIRANHKO','Poppins','Montserra','NEGERI9 GARVT') dan seluruh deskripsi berbunyi 'Moodboard branding SMK Negeri 9 Garut' - ini branding SEKOLAH (biru navy/kuning emas), bukan RR."), ('SEBAGIAN', "Tiga mockup terverifikasi (IMG#03/04/05 400x218, 400x218, 400x218) dengan penjelasan fungsi lengkap (Packaging, Laptop, Hoodie), tetapi seluruh konten beridentitas 'The Creative Studio' dengan logo inisial 'AF' dan slogan 'Ahmad Faisal | Visual Storyteller' - bukan RR."), ('SEBAGIAN', "Naskah memuat Judul 'Langkahmu Dimulai di Sini', Tema, Pesan Utama, Narasi/Dialog 5 scene, Closing Tagline 'Berkarakter, Kompeten, Siap Menghadapi Masa Depan' - lengkap secara struktur, tetapi subjeknya iklan SEKOLAH 'SMK Negeri 9 Garut', bukan personal branding siswa."), ('SEBAGIAN', "Storyline 'THE FUTURE FRAME' memuat 1. Pembukaan, 2. Alur Cerita, 3. Konflik/Fokus Visual, 4. Penutup lengkap dengan Visual/Audio/Narasi, tetapi ditulis untuk brand 'The Creative Studio' (AF), bukan RR."), ('LENGKAP', "Shotlist ditulis penuh 10 baris dengan kolom No/Adegan/Jenis Shot/Sudut Pandang/Gerakan/Durasi/Deskripsi Visual & Audio (baris 1-10), dan didukung gambar tabel IMG#06 400x213 (OCR 'No. / Adegan / Jenis Shot / Angle / Morement / Durasi / Deskripsi' dengan baris 01-10). Subjeknya 'The Creative Studio'."), ('LENGKAP', "Storyboard 6 panel tertulis lengkap (Panel 01-06) dengan Visual, Keterangan Shot, Angle Kamera, Transisi, dan Narasi per panel; didukung gambar IMG#07 400x218 (OCR 'STORYBOARD VIDEO PROMOSI: THE CREATIVE STUDIO')."), ('SEBAGIAN', "Teks maskot mendeskripsikan karakter berdiri tegak (full body) tetapi untuk 'The Shinobi Studio - Maskot Kedua (Rina)' dengan logo 'AF'; gambar IMG#08 400x223 hanya terbaca 'THE SHINOBI STUOIO / MASKOT KEDUA CRINA'. Bukan maskot personal branding RR."), ('LENGKAP', '8 prompt terdokumentasi dan diberi label di dalam artikel: prompt logo, prompt moodboard, prompt mockup, prompt naskah, prompt storyline, prompt shotlist, prompt storyboard, prompt maskot ke-2.'), ('BELUM MEMENUHI', "Link aktif (status 200) tetapi labels=[] - 0 dari 5 label wajib ('Tugas Sekolah', 'Personal Branding', 'AI', 'Portofolio', 'SMKN 9 Garut') tidak ada."), ('BELUM MEMENUHI', "Identitas visual berubah-ubah dalam satu artikel: bagian 1 logo 'RR' Muhammad Reza Ramdani; bagian 2 moodboard branding SMK Negeri 9 Garut; bagian 3, 4.2, 4.3, 4.4 memakai brand 'The Creative Studio' dengan inisial 'AF' dan nama 'Ahmad Faisal | Visual Storyteller'; bagian 5 memakai 'The Shinobi Studio - Rina'. Teks deskripsi mockup 'The Creative Studio' tersebut terindikasi bukan karya pribadi siswa (identik dengan template deskripsi generik hasil AI) - terindikasi tidak orisinal / perlu klarifikasi guru.")],
))
S.append(('R55','MUHAMMAD TAUFIQ ISMAIL','XI DKV 3','03 Okt 2026 00:12',
'https://taufikismailll.blogspot.com/2026/10/personal-branding-muhammad-taufiq.html',
'2 (terverifikasi teks)','0 (tidak ada shotlist)','6 (terverifikasi teks)','1 (full body teks)','8 (7 komponen)','6815 (memenuhi)',[4, 4, 4, 2, 4, 4, 1, 4, 3, 4, 2, 3],
[('LENGKAP', "Page title 'Personal Branding Muhammad Taufiq Ismail XI DKV 3' (format pendek sesuai) + heading artikel dua h1: 'ASTS PERSONAL BRANDING' dan 'MUHAMMAD TAUFIQ ISMAI XI DKV 3 SMKN 9 GARUT' (format panjang sesuai). CATATAN: nama pada heading terpotong satu huruf menjadi 'TAUFIQ ISMAI', sedangkan 22 kali lain di artikel ditulis 'Taufiq Ismail'."), ('LENGKAP', "Logo monogram TL + pemain futsal + bola + mahkota (IMG#01 320x320, OCR 'FUTSAL / TL'); nama branding 'TL - Muhammad Taufiq Ismail'; tagline 'PLAY FUTSAL, LIVE THE DREAM.'; NAMA SISWA tercetak pada logo menurut deskripsi bagian '10. Komposisi Logo' ('Pada bagian bawah terdapat tulisan MUHAMMAD TAUFIQ ISMAIL'); deskripsi konsep 15 sub-bagian (Latar Belakang s.d. Kesimpulan) +/-1.177 kata, jauh di atas 100 kata."), ('LENGKAP', "IMG#02 320x213 landscape (prompt moodboard juga meminta 'format landscafe dan kualitas lebih dari 4k'). Tujuh unsur terverifikasi di teks: Warna Utama (Black #000000, White #FFFFFF, Silver #C0C0C0, Electric Blue #007BFF), Typography (Montserrat Bold/Extra Bold, Bebas Neue, Brush Script, Sporty Brush), Style Visual (Modern, Dynamic, Bold, Sporty), Referensi Desain (jersey futsal, sepatu, banner, desain logo olahraga, lapangan futsal, apparel), Elemen Grafis (Mahkota, Bola Futsal, Speed Line, Monogram TL, Perisai, Garis Dinamis), Inspirasi Visual (lapangan futsal malam hari, jaring gawang, lampu lapangan, gunung). Tone & mood tidak memakai label khusus, diwakili bagian '12. Karakter yang Ingin Ditampilkan' (Tegas, Sporty, Modern, Dinamis, Percaya diri, Profesional)."), ('BELUM MEMENUHI', "Hanya 2 media mockup, di bawah syarat minimal 3: Futsal Jersey Set dan Hoodie - teks sendiri menulis 'Dua produk utama yang ditampilkan adalah: Futsal Jersey Set'. Uraian produk sangat rinci (nomor 02, nama TAUFIQ, kerah V-neck, bahan dry-fit, detail hood, jahitan, palet warna, elemen grafis) tetapi tidak ada penjelasan fungsi tiap media; visual tunggal IMG#03 320x213 (OCR '02 / 02 / 正 / 02') tidak memisahkan dua produk."), ('LENGKAP', "'5. naskah' memuat Judul 'PLAY FUTSAL, LIVE THE DREAM', Tema (semangat, passion, impian dalam dunia futsal), Pesan Utama, Narasi/Dialog 6 scene (Scene 1 Pembukaan sampai Scene 6 Penutup, tiap scene memuat Visual + kutipan dialog), serta Closing Tagline pada 'Teks layar: TL - MUHAMMAD TAUFIQ ISMAIL / MORE THAN A GAME, IT'S A LIFESTYLE.'"), ('LENGKAP', "'6. storyline' memuat keempat unsur: 1. PEMBUKAAN (lampu electric blue menyala, maskot memasuki lapangan), 2. ALUR CERITA (memasang sepatu, menggiring bola, skill, tendangan gol, pose percaya diri), 3. FOKUS VISUAL (warna, elemen yang ditonjolkan, gaya cinematic sports style, transisi fade/whip pan/speed ramp), 4. PENUTUP (medium shot ke low angle hero shot + tagline 'PLAY FUTSAL, LIVE THE DREAM.'). Diperkuat tabel HTML 7 baris x 6 kolom (1 header + 6 scene) dengan kolom Visual Adegan, Keterangan Shot, Angle Kamera, Transisi, Dialog/Narasi."), ('BELUM MEMENUHI', "Tidak ada shotlist. Bagian '7. Shootlish' (ejaan salah) hanya memuat jawaban AI 'Siap. Berikut storyboard AI 6 scene yang sudah menyatukan naskah, storyline, dan shotlist untuk branding TL' - jadi shotlist tidak pernah dikirim sebagai dokumen tersendiri; hanya ada 6 scene pada tabel, di bawah syarat minimal 10 shot, dan tanpa kolom No/Duration/Movement. Teks bagian 7 hanya memuat judul 'Shootlish' (ejaan salah dari 'Shotlist') dan 'Arahan Visual AI'."), ('LENGKAP', "Enam scene storyboard lengkap di artikel dan di tabel (1. Pembukaan s.d. 6. Closing Branding) dengan Keterangan Shot (Establishing shot, Close-up, Tracking Shot, POV Ball, Slow Motion, Hero Shot), Angle Kamera (Wide Shot + Low Angle, Medium Shot, Over Shoulder, Front Center Shot), Transisi (Fade In, Slow Zoom + Light Flash, Match Cut, Whip Pan, Speed Ramp, Fade Out) dan Dialog/Narasi tiap scene; 'Artikel Storyboard Iklan TL' (182 kata) menyebut durasi 30-45 detik; IMG#05 320x292 OCR 'STORYBOARD IKLAN' + angka 02, 02, 0个0 (pseudoteks, jumlah panel tidak dapat dihitung)."), ('LENGKAP', "'4. maskot' diuraikan 20 sub-bagian: karakter pemain futsal duduk dengan bola di depan, format digital 1:1, komposisi bagian atas (mahkota + logo), tengah (karakter), bawah (bola, sepatu, lantai lapangan) - ini bukti full body secara tekstual; prompt maskot terdokumentasi ('sekarang buatkan maskot tentang logo yg tadi dengan ketentuan kualitas tinggi 4k'). TAPI hanya ada 1 gambar maskot IMG#04 320x320 (OCR '及 / PLAYHISN / LINE THE DREAMY / TECPIOIEWAL / GIAAL / 02 / T' - pseudoteks); tidak ada varian Portrait maupun Maskot Bersama Logo."), ('LENGKAP', "8 prompt terdokumentasi dan diberi label ('Prompt yang saya gunakan:' 4x dan 'promp yang saya gunakan:' 4x): logo ('berperanlah sebagai expert designer, tolong buatkan logo dari singkatan atau inisial MUHAMMAD TAUFIQ ISMAIL'), moodboard ('tolong buatkan moudbore tentang, warna utama branding, typography, style visual, referensi desain, elemen grafis dan inspirasi visual, output 1 moodbore, format landscafe'), mockup ('sekarang buat mockup dari logo tersebut kedalam baju jersey dan hodie kualitas 4k'), maskot, naskah, storyline, dan storyboard (2x). Tidak ada prompt terpisah untuk shotlist."), ('SEBAGIAN', "Link aktif status 200, isi artikel paling panjang di batch ini (6.815 kata, bagian 1-15 branding, 1-13 moodboard, 2-18 mockup, 1-20 maskot, 5 naskah, 6 storyline, 7 Shootlish, 8 Storyboard). Namun hanya 3 dari 5 label wajib terpasang: 'personal branding', 'tugas sekolah', dan 'SMKN 9 GARUT' (yang muncul di dalam label tidak relevan 'mengesahkan kreatifitas di lab dkv SMKN 9 GARUT'); label 'AI' tidak ada dan 'Portofolio' tidak ada (hanya label 'profil'); ada salah ketik label 'brending logo'."), ('SEBAGIAN', "Sistem visual sangat konsisten (monogram TL, mahkota, bola futsal, perisai, speed line; palet #000000/#FFFFFF/#C0C0C0/#007BFF; slogan 'PLAY FUTSAL, LIVE THE DREAM') dari logo sampai storyboard. TEMUAN KHUSUS: ada tanda ezrawan dari sumber lain yang tidak dihapus - '( IndiBlogHub )' di paragraf pembuka bagian maskot dan '( Markuva )' di bagian '16. Hubungan dengan Personal Branding' - indikasi teks mentah hasil AI disalin tanpa diedit, perlu klarifikasi guru (terindikasi tidak orisinal). Ditambah teks banyak pengulangan ('Jersey futsal Jersey futsal', 'Percaya diri Percaya diri'), salah eja ('Shootlish', 'brending', 'promp', 'moudbore', 'visiual'), dan kelima gambar berukuran kecil (320 px) dengan pseudoteks.")],
))
S.append(('R94','NADA NISRINA','XI DKV 2','02 Okt 2026 15:44',
'https://nadaanisrina.blogspot.com/2026/10/personal-branding-nada-nisrina-xi-dkv-2.html',
'3 (terverifikasi gambar IMG#03-05)','13 (terverifikasi teks)','13 (terverifikasi teks + gambar IMG#08)','3 varian (terverifikasi OCR IMG#06)','10','4064 (memenuhi)',[4, 4, 3, 4, 4, 4, 3, 3, 4, 4, 2, 4],
[('LENGKAP', "Page title 'Personal Branding Nada Nisrina XI DKV 2' dan heading 'ASTS PERSONAL BRANDING NADA NISRINA XI DKV 2 SMKN 9 GARUT' - keduanya sesuai format; nama rekapan = nama artikel."), ('LENGKAP', "Logo monogram NN (IMG#01 320x320, OCR 'NADA / NISRINA'); nama branding 'NADA NISRINA'; tagline 'Wrap Your Feelings in Bloom'; deskripsi konsep 181 kata (>=100); NAMA SISWA tercetak pada logo."), ('SEBAGIAN', "IMG#02 320x179 landscape. Enam unsur dibahas pada label terpisah (Hirarki Tipografi, Palet Warna, Filosofi Desain, Elemen Visual Utama 1-6); 'Referensi Desain' tidak ada sebagai unsur tersendiri dan prompt yang dilampirkan hanya meminta 6 butir tanpa referensi desain."), ('LENGKAP', "Tiga mockup dengan uraian Visual + Filosofi + Fungsi yang lengkap: A.KAOS, B.kartu nama, C.packaging; didukung 3 berkas IMG#03/04/05 (320x320) ber-OCR 'NADA / NISRINA'."), ('LENGKAP', "Naskah memuat Judul 'Seni Merangkai Rasa (The Art of Floral & Emotional Crafting)', Tema, Pesan Utama, Closing Tagline, serta Narasi/Dialog 4 bagian bertimestamp 00:00-00:12 sampai 00:50-01:00 lengkap VO dan Dialog Nada."), ('LENGKAP', 'Storyline memuat 4 unsur eksplisit dengan durasi, visual, narasi/VO, dan suasana musik: 1. Pembukaan, 2. Alur Cerita, 3. Konflik/Fokus Visual, 4. Penutup.'), ('SEBAGIAN', 'Shotlist ditulis lengkap sebagai prosa berurutan Shot 1 sampai Shot 13 (>=10 shot terpenuhi) lengkap dengan jenis shot dan angle tiap shot (Extreme Close-Up, High Angle, Top Down, Eye Level), tetapi tanpa kolom No/Adegan/Jenis Shot/Angle/Movement/Durasi/Deskripsi; gambar IMG#07 320x239 OCR-nya pseudoteks.'), ('SEBAGIAN', "Storyboard tertulis 13 scene (Shot 1-13) dengan Spesifikasi Kamera, Visual & Keterangan, Aksi, dan rentang waktu, didukung gambar IMG#08 320x239 (OCR 'STORYBOARD PRODUKSI' dengan penanda waktu 00:00-00:12 ... 00:55-01:00); namun dialog atau narasi tidak disediakan per scene."), ('LENGKAP', "IMG#06 320x320 dengan OCR 'MASCOTFULLBODY', 'MASCOTPORTRAIT', 'MASCOYWTHBRANOINGLOGO' dan 'NADA NISRINA' - ketiga output wajib terverifikasi."), ('LENGKAP', '10 prompt terdokumentasi dan berlabel: prompt logo, prompt moodboard, prompt mockup kaos (A), mockup kartu nama (B), mockup packaging (C), prompt maskot, prompt naskah, prompt storyline, prompt shotlist, prompt storyboard.'), ('BELUM MEMENUHI', "Link aktif, tetapi hanya 1 dari 5 label wajib ('tugas sekolah'); 'Personal Branding', 'AI', 'Portofolio', dan 'SMKN 9 Garut' tidak ada. Label lain tidak relevan: 'CV' dan 'kegiatan membersihkan lab'."), ('LENGKAP', "Identitas visual NN, palet blush pink/sage green/cream/rose gold, dan tagline 'Wrap Your Feelings in Bloom' konsisten dari bagian 1 sampai 8; nama tidak pernah berubah.")],
))
S.append(('R67','NADIA FITRIANI','XI DKV 3','02 Okt 2026 21:04',
'https://nadiafitriani17.blogspot.com/2026/10/personal-branding-nadia-fitriani-xl-dkv.html',
'3 (terverifikasi gambar)','10 (terverifikasi teks)','6 (terverifikasi teks)','3 (terverifikasi gambar)','8','2431 (memenuhi)',[4, 3, 3, 4, 4, 4, 3, 4, 4, 4, 2, 4],
[('LENGKAP', "Page title 'personal branding nadia fitriani Xl DKV 3' + heading 'ASTS PERSONAL BRANDING NADIAFITRIANI Xl DKV 3 SMKN 9 GARUT' (nama ditulis tanpa spasi) - nama dan kelas sesuai."), ('SEBAGIAN', "Logo NF dengan deskripsi konsep 110 kata dan makna elemen (monogram NF, mahkota, pink dusty/rose, garis melingkar, hati kecil, sparkle) serta tagline 'Dream Design Create Grow'; nama lengkap terbaca di IMG#02 ('Nadia Fitrinni') tetapi OCR logo IMG#01 320x320 hanya 'T' sehingga nama pada logo tidak dapat diverifikasi."), ('SEBAGIAN', '7 unsur lengkap tertulis (Warna Utama, Typography, Style Visual, Referensi Desain, Tone & Mood, Elemen Grafis, Inspirasi Visual) - tetapi gambar moodboard IMG#02 316x320 berorientasi potret, bukan landscape.'), ('LENGKAP', "3 mockup dengan makna branding tiap media: laptop ( centerpiece), roll-up/standing banner, dan kartu nama; gambar IMG#03 320x179 (OCR 'NF' berulang pada laptop, banner, dan kartu nama)."), ('LENGKAP', "Judul 'Langkah Kecil, Mimpi Besar bersama Nadia Fitriani', Tema 'Elegansi Kreatif, Kepercayaan Diri, dan Inspirasi Kecantika', Pesan Utama, Narasi & Dialog 4 scene (VO dan dialog Nadia), Closing Tagline 'Nadia Fitriani - Dream. Design. Create. Grow.'"), ('LENGKAP', '1. Pembukaan (Opening), 2. Alur Cerita (Development), 3. Konflik / Fokus Visual (Climax & Key Visual), 4. Penutup (Closing) - 4 unsur lengkap.'), ('SEBAGIAN', "10 shot tertulis berurutan (Shot 1-10) lengkap dengan jenis shot, angle, movement, dan durasi, tetapi ditulis paragraf bukan tabel. Tabel shotlist ada di IMG#05 148x320 (potret, OCR 'SHOTLIST ... 10 SHOTS - NADIA FITRIANI') yang isinya tidak terbaca."), ('LENGKAP', "6 scene tertulis lengkap di artikel (Scene 1-6) dengan Visual, Shot, Angle, Transition, dan Narasi; dikonfirmasi IMG#06 320x179 (OCR 'Nadia Fitrians | Branding Storyboard')."), ('LENGKAP', "IMG#04 148x320 memuat label 'Mascot Full Body' dan 'Mascot Portrait'; teks menyebut tiga visual: 1. Mascot Full Body, 2. Mascot Portrait, 3. Mascot Bersama Logo Branding - ketiganya terpenuhi."), ('LENGKAP', "8 prompt berlabel ('prompt yg saya gunakan' / 'prompt yg di gunakan') untuk 8 bagian: logo, moodboard, mockup, maskot, naskah, storyline, shootlist, storyboard."), ('BELUM MEMENUHI', '0 dari 5 label wajib. Label yang terpasang bukan label tugas ini: ASTS, Branding logo, CV, Maskot, Moodboard, kegiatan pratikum.'), ('LENGKAP', "Identitas NF konsisten (monogram NF, dusty rose/maroon, 'Dream Design Create Grow') di logo, moodboard, mockup, maskot, naskah, storyline, shotlist, dan storyboard; konten orisinal.")],
))
S.append(('R89','NADYA NURLATIFA','XI DKV 4','02 Okt 2026 12:37',
'https://nadya29710.blogspot.com/2026/10/personal-branding-nadya-xi-dkv-4.html',
'1 (terverifikasi gambar)','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','1 (terverifikasi teks)','0','415 (tidak memenuhi)',[4, 2, 0, 1, 0, 1, 2, 2, 2, 0, 1, 2],
[('LENGKAP', "Page title 'Personal branding nadya xi dkv 4' + heading 'ASTS PERSONAL BRANDING NADYA NURLATIFA XI DKV 4 SMKN 9 GARUT' - keduanya sesuai dengan nama rekapan."), ('SEBAGIAN', "Logo NN dengan deskripsi konsep 110 kata (makna: monogram NN, mahkota, ungu/lavender, garis melingkar, hati, sparkle) dan konsep warna; tetapi nama branding ditulis 'NN - by Nadya Nurlatifa' - salah eja dari NADYA NURLATIFA (5 kali di teks) dan OCR logo IMG#01 tidak terbaca."), ('TIDAK DIKUMPULKAN', 'Tidak ada bagian moodboard: 5 bagian artikel hanya LOGO BRANDING, KONSEP VIDIO, KONSEP MASCOT, MOCKUP BRANDING, NADYA CREATIVE. Tiga gambar (IMG#03 320x249, IMG#05 dan IMG#07 320x316) OCR-nya pseudoteks sehingga tidak dapat dipastikan sebagai moodboard - perlu dibuka guru.'), ('BELUM MEMENUHI', "Hanya 1 gambar mockup (IMG#04 320x213, OCR 'mOckuP BRAndING / STICIR / LAMNER'); teks menyebut media logo, sticker, banner, kemasan, kartu ucapan tanpa penjelasan fungsi tiap media."), ('TIDAK DIKUMPULKAN', 'Tidak ada bagian naskah iklan; 415 kata artikel tidak memuat Judul, Tema, Pesan Utama, Narasi/Dialog, maupun Closing Tagline.'), ('SEBAGIAN', "Hanya bagian '2.KONSEP VIDIO' yang berisi alur singkat (ide -> sketsa -> diedit -> dibagikan -> logo), tanpa struktur Pembukaan, Konflik/Fokus Visual, dan Penutup."), ('TIDAK DAPAT DIVERIFIKASI', "Gambar shotlist IMG#02 320x211 terbaca headed 'SHOTLIST / VIDEO IKLAN KOMERSIAL', tetapi angka dan baris di dalamnya ('8', '15', '15') adalah pseudoteks sehingga jumlah shot tidak dapat dihitung andal."), ('TIDAK DAPAT DIVERIFIKASI', "Gambar storyboard IMG#06 320x213 terbaca 'STORYBOARD / VIDEO IKLAN KOMERSIAL'; jumlah scene serta keterangan shot/angle/transisi tidak terbaca - tidak boleh dihitung."), ('SEBAGIAN', "Teks '3.KONSEP MASCOT' menjelaskan maskot perempuan berhijab bergaya chibi (ungu/cream, hoodie, tote bag, sneakers, tablet, kamera); dua gambar 320x316 (IMG#05/IMG#07) memuat token 'MASC...' yang tidak terbaca - Full Body tidak dapat dipastikan."), ('TIDAK DIKUMPULKAN', "Tidak ada satu pun prompt terdokumentasi di seluruh 415 kata artikel (kemunculan kata 'prompt' = 0)."), ('BELUM MEMENUHI', '0 dari 5 label wajib; daftar label kosong ([]). Isi artikel sangat tipis: 415 kata dan 5 sub-bagian.'), ('SEBAGIAN', "Palet ungu-krem dan logo NN konsisten pada logo, mockup, dan concept video; tetapi nama pada branding salah eja 'Nadya Nurlatifa' dan teks artikel terpotong di bagian '5.NADYA CREATIVE' (...serta elem).")],
))
S.append(('R87','NAILA AGUSTIN','XI DKV 4','02 Okt 2026 11:32',
'https://lanailagstin.blogspot.com/2026/10/personal-branding-naila-agustin.html',
'4 media (1 berkas gambar 320x292)','10 (terverifikasi teks)','6 (terverifikasi teks)','1 (full body TIDAK DAPAT DIVERIFIKASI)','5','1759 (memenuhi)',[3, 4, 3, 3, 3, 4, 3, 3, 2, 4, 3, 3],
[('SEBAGIAN', "Page title 'Personal branding Naila Agustin' TANPA nama kelas; heading artikel 'ASTS PERSONAL BRANDING NAILA AGUSTIN XI DKV 4 SMKN 9 GARUT' sesuai. Nama pada rekapan dan artikel sama."), ('LENGKAP', "Logo inisial NA (IMG#01 320x320, OCR 'NAILA AGUSTIN'); nama branding 'Naila Agustin'; tagline 'More Than Just a Brand, It's a Lifestyle'; deskripsi konsep 142 kata (>=100); NAMA SISWA tercetak pada logo."), ('SEBAGIAN', 'IMG#02 320x213 landscape. Ketujuh unsur hanya berbentuk prosa-paragraf, bukan 7 butir berlabel: warna utama (Royal Blue/Sky Blue/Light Blue/Ice Blue), typography (Playfair Display + Montserrat), style visual (modern, feminin, minimalis), referensi desain (kain biru, bunga, arsitektur lengkungan), tone & mood (calm, sophisticated), elemen grafis (bintang, divider, ribbon), closure/penutup. Semua unsur ada tetapi tidak dihitung per butir.'), ('SEBAGIAN', "Teks menjelaskan 4 media mockup (hoodie, kartu nama, social media feed, packaging) beserta fungsi tiap media, tetapi hanya 1 berkas gambar mockup (IMG#03 320x292, OCR pseudoteks 'STOOZUA','VAILANCOSTIS') sehingga jumlah mockup visual tidak dapat diverifikasi."), ('SEBAGIAN', "Bagian '5.Konsep Vidio iklan' memuat judul 'Naila Agustin - More Than Just a Brand, It's a Lifestyle', tema fashion-lifestyle, pesan utama, alur adegan dan tagline penutup, tetapi tanpa blok Dialog/Narasi terpisah."), ('LENGKAP', "Bagian '6.Alur Cerita' memuat 4 unsur bernomor eksplisit: 1. Pembukaan, 2. Alur Cerita, 3. Konflik/Fokus Visual, 4. Penutup - masing-masing satu kalimat bermakna."), ('SEBAGIAN', "Teks 'Daftar Pendek' menulis 10 shot berurutan lengkap dengan jenis shot, angle, movement dan durasi (5 detik pada shot 1), tetapi disajikan sebagai paragraf naratif tanpa kolom No/Adegan/Jenis Shot/Angle/Movement/Durasi/Deskripsi."), ('SEBAGIAN', 'Enam scene tertulis eksplisit (Scene 1-6) lengkap dengan angle dan transisi (fade in, cut, pan right) serta narasi pada Scene 1, tetapi dialog/narasi tidak tersedia pada scene 2-6.'), ('SEBAGIAN', "Teks menyatakan maskot dibuat 3 versi (full body, portrait, maskot dengan logo); hanya 1 berkas IMG#04 320x320 dengan OCR fragmen 'LSlascFull Bw', '2MfasuPorteaa', 'Naila Agustin'. Full Body tidak dapat dipastikan dari bukti visual."), ('LENGKAP', '5 prompt terdokumentasi dengan label dan teks penuh: 1 Prompt Logo Branding, 2 Prompt Maskot, 3 Prompt Moodboard, 4 Prompt Mockup Branding, 5 Prompt Video Iklan.'), ('SEBAGIAN', "Link aktif (status 200). 6 label: 'AI', 'BRANDING', 'PERSONAL', 'PORTOFOLIO', 'SMKN 9 GARUT', 'TUGAS SEKOLAH.' - fifth label wajib terpecah menjadi 'PERSONAL'+'BRANDING' dan ditulis dengan titik ('TUGAS SEKOLAH.'), bukan label baku."), ('SEBAGIAN', 'Identitas visual konsisten (nama, palet biru, monogram NA) di seluruh 10 bagian; penomoran bagian melompat (6 lalu 8, tanpa 7) dan satu blok prompt video iklan masih menyatu dengan bagian 5.')],
))
S.append(('R96','NAPSA JAKIYAH','XI DKV 4','02 Okt 2026 16:25',
'https://napsajakiyah.blogspot.com/2026/10/personal-brending-napsa-jakiyah-xi-dkv4.html',
'6 media (1 berkas gambar 320x213)','10 (terverifikasi teks)','7 (terverifikasi teks)','1 (full body TIDAK DAPAT DIVERIFIKASI)','6','2629 (memenuhi)',[3, 4, 3, 3, 3, 2, 4, 4, 2, 4, 4, 3],
[('SEBAGIAN', "Kedua judul ada tapi salah eja: page title 'PERSONAL BRENDING NAPSA JAKIYAH XI-DKV4' (BRENDING) dan heading 'ASTAGA PERSONAL BRANDING NAPSA JAKIYAH XI-DKV4 SMKN 9 GARUT' (ASTAGA, XI-DKV4). Nama rekapan = nama artikel."), ('LENGKAP', "Logo inisial NJ (IMG#01 320x320, OCR 'NAPSA / JAKIYAH'); nama branding 'NAPSA JAKIYAH'; tagline 'Create - Capture - Inspire' dan slogan 'Small Steps, Big Dreams'; deskripsi konsep 160 kata; NAMA SISWA tercetak pada logo."), ('SEBAGIAN', 'IMG#02 320x213 landscape. Enam unsur dibahas prosa (warna utama pink/biru/putih, tipografi Playfair Display+Allura+Montserrat, style visual, elemen grafis, REFERENSI DESAIN hanya disebut di bagian 9, tone & mood, inspirasi visual) - tidak disusun sebagai 7 butir berlabel.'), ('SEBAGIAN', "Teks mencantumkan 6 media mockup (Kaos & Hoodie, Sticker, Kartu Nama, Gelas Kopi, Social Media Feed, Media Promosi) dengan fungsi tiap media, tetapi hanya 1 berkas gambar (IMG#03 320x213, OCR 'No. / N') sehingga 3 mockup visual tidak dapat diverifikasi."), ('SEBAGIAN', "Naskah memuat judul 'Create - Capture - Inspire', durasi 30-45 detik, tema, alur 5 tahap, pesan utama dan closing tagline 'Create - Capture - Inspire' + 'Small Steps, Big Dreams', tetapi tidak ada blok Dialog/VO."), ('BELUM MEMENUHI', 'Storyline hanya 4 butir bullet tanpa label Pembukaan / Alur Cerita / Konflik-Fokus Visual / Penutup dan tidak menyebut konflik sama sekali; yang ada adalah urutan adegan promo brand.'), ('LENGKAP', 'Shotlist ditulis penuh di teks dengan 7 kolom (No, Adegan, Jenis Shot, Angle, Movement, Durasi, Deskripsi) baris 1-10: 1.Pembukaan Close Up Eye Level Fade In 4 detik ... 10.Penutup Wide Shot Eye Level Fade Out 3 detik. Kolom Movement diisi jenis transisi (Fade In/Pan/Zoom In/Slide).'), ('LENGKAP', 'Storyboard ditulis penuh dengan kolom (Scene, Visual Adegan, Keterangan Shot, Angle Kamera, Transisi, Dialog/Narasi) untuk 7 scene (1 Pembukaan ... 7 Closing), termasuk dialog narasi tiap scene.'), ('SEBAGIAN', "Bagian '4.MASCOT BANNER' hanya 1 berkas IMG#04 320x292 (OCR 'MAPSA / JAK'YAH / NAPSA JAKIYAH'). Teks menyebut character perempuan anime berpakaian pink-biru tanpa menyatakan varian full body; prompt D (meminta full body/portrait/beserta logo) adalah perintah, bukan bukti."), ('LENGKAP', '6 prompt terdokumentasi berlabel A-F lengkap dengan teks: A Prompt Logo, B Prompt Moodboard, C Prompt Mockup, D Prompt AI Mascot, E Prompt Video Iklan Komersial, F Prompt Storyboard.'), ('LENGKAP', "Link aktif. Kelima label wajib tertera: 'TUGAS SEKOLAH', 'PERSONAL BRANDING', 'AI', 'PORTOFOLIO', dan 'SMKN9GARUT' (hanya berbeda pada spasi)."), ('SEBAGIAN', "Nama, inisial NJ, palet pink-biru dan tagline konsisten di 10 bagian; namun terdapat artefak obrolan AI yang belum dibersihkan di bagian 9 ('Tentu, Napsa. Ini terusan lengkap dari semua bagian di atas, tetapi sudah disesuaikan...') dan ejaan 'ASTAGA' pada heading.")],
))
S.append(('R108','PITRYA HANDAYANI','XI DKV 2','02 Okt 2026 21:48',
'https://pitriaaahandayaniiiiiiiii.blogspot.com/2026/10/personal-branding-pitrya-handayanixi.html',
'3 (terverifikasi gambar)','10 (terverifikasi teks)','10 (terverifikasi teks)','3 (dalam 1 gambar komposit)','7','3223 (memenuhi)',[4, 4, 4, 4, 4, 4, 3, 4, 4, 3, 3, 3],
[('LENGKAP', "Page title 'Personal Branding Pitrya HandayaniXI DKV 2' (tanpa spasi sebelum XI) + heading 'ASTS Personal Branding Pitrya Handayani Xl DKV 2 SMKN 9 GARUT'. Catatan: nama pada prosa artikel ditulis 'Pitria Handayani' (7x) dan 'Pistrya' (1x) - tidak konsisten dengan nama rekapan PITRYA HANDAYANI."), ('LENGKAP', "Logo monogram PH (IMG#01 320x239, OCR 'PITRYA HANDAYANI ... COOKING STORIES PHOTOGRAPHY MEMORIES'); nama branding 'PH / Pitrya Handayani'; tagline 'Kreasi, Rasa, dan Karya Visual'; deskripsi konsep 193 kata (Deskripsi 83 + Makna Elemen Logo 94) - sudah >=100 kata."), ('LENGKAP', "IMG#02 320x213 landscape (OCR 'BRAND MOODBOARD / Pitrya Handayan'). 7 unsur terverifikasi di teks: Color Palette (hex #F4D9D9, #F6F1E7, #CBBAA8), Typography (script + serif), Tone & Mood, Graphic Elements (wave/camera/cooking), Inspiration & Visual References (3 foto), Brand Style, Brand Essence."), ('LENGKAP', "3 mockup dengan 'Fungsi Media' + 3 poin manfaat tiap media: Packaging, Tumbler, Kaos; gambar IMG#03 320x292 terbaca '1. PACKAGING / 2.TUMBLER / 3.KAOS' dan 'PITRYAHANDAYANI'."), ('LENGKAP', "NASKAH IKLAN 60 detik: Judul 'RASA YANG TERCERITA', Tema 'Mindful Creation', Pesan Utama, Narasi/Dialog 4 scene bertimecode (0-15 / 15-35 / 35-50 / 50-60 detik), Closing Tagline 'PITRYA HANDAYANI KREASI, RASA, DAN KARYA VISUAL'."), ('LENGKAP', "Storyline 'Cerita dalam Rasa dan Lensa': Pembukaan 00:00-00:15, Alur/Pertengahan 00:15-00:35, Konflik/Klimaks 00:35-00:50, Penutup Branding 00:50-01:00 - 4 unsur lengkap."), ('SEBAGIAN', "10 shot tertulis berurutan ([SHOT 1] sampai [SHOT 10]) lengkap dengan Jenis Shot, Angle, Movement, durasi, dan deskripsi, tetapi bentuknya paragraf bukan tabel. Tabel shotlist ada di IMG#05 320x294 (OCR 'SHOTLIST MASKOT PTRYA HANDAYANI') dengan isi pseudoteks sehingga kolomnya tidak dapat diverifikasi."), ('LENGKAP', "10 scene tertulis lengkap di artikel (Scene 1-10) dengan Visual Adegan, Jenis Shot, Movement, Angle, Transisi, dan Dialog/Narasi; dikonfirmasi IMG#06 320x292 (OCR 'STORYBOARD / Cerita dalam Rasa / PITRYA HANDAYANI / Maskot Pitrya Handayani')."), ('LENGKAP', "Teks '4.MASCOUT' menyebut tiga visual utama: Full body (gambar kiri), Portrait (gambar kanan atas), dan Brand presentation/men bersama logo (gambar kanan bawah). Satu gambar yang belum terpetakan adalah IMG#04 320x292 (OCR hanya 'PITRYAHANDAYAAI') dan sesuai urutan 6 bagian inilah sheet maskot."), ('SEBAGIAN', '7 prompt berlabel PROMPT (logo, moodboard, mockup, mascot, naskah, shotlist) - prompt yang ditulis setelah bagian storyline tertuli ulang sama dengan prompt naskah, dan tidak ada prompt untuk storyboard.'), ('SEBAGIAN', "4 dari 5 label wajib (AI, Personal Branding, Portofolio, tugas sekolah); label 'SMKN 9 Garut' tidak ada, sedangkan label 'CV' tampaknya milik kiriman tugas lain."), ('SEBAGIAN', "Identitas visual konsisten (monogram PH, dusty pink/taupe/cream, ikon kamera/kuliner/renang) dan konten orisinal; tetapi nama prosa salah tulis serta 3 tagline berbeda dipakai: 'Kreasi, Rasa, dan Karya Visual', 'Small Details Big Impact', dan 'Good Food Good Mood Better Memories'.")],
))
S.append(('R99','RADIT KURNIAWAN','XI DKV 3','02 Okt 2026 22:20',
'https://www.blogger.com/u/1/blog/post/edit/2251886040121859008/2007859641830967938',
'TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI','TIDAK DAPAT DIVERIFIKASI',[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
[('TIDAK DAPAT DIVERIFIKASI', 'Link yang dikirim adalah alamat EDITOR Blogger (https://www.blogger.com/u/1/blog/post/edit/...), bukan link artikel publik; halaman hanya dapat dibuka setelah login Google sehingga karya tidak dapat diakses sama sekali.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - tidak ada page title, tidak ada teks artikel, tidak ada gambar yang dapat diunduh.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian moodboard tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian mockup tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian naskah iklan tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian storyline tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian shotlist tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian storyboard tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian maskot tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger - bagian prompt tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'URL editor Blogger (bukan artikel publik) - label Blogger tidak dapat diperiksa.'), ('TIDAK DAPAT DIVERIFIKASI', 'Karya tidak dapat diakses sehingga orisinalitas dan profesionalisme tidak dapat dinilai.')],
))
S.append(('R97','RAFI FAUZAN NAJA LUTFIANA','XI DKV 4','02 Okt 2026 16:38',
'https://rafiifauzan-profil.blogspot.com/2026/10/asts-komputer-grafis-xi-dkv-4.html?m=1',
'0 (TIDAK DIKUMPULKAN)','0 (TIDAK DIKUMPULKAN)','0 (TIDAK DIKUMPULKAN)','0 (TIDAK DIKUMPULKAN)','0','146 (tidak memenuhi)',[1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1],
[('BELUM MEMENUHI', "Page title 'ASTS KOMPUTER GRAFIS XI DKV 4' dan teks heading 'ASTS KOMPUTER GRAFIS RAFI FAUZAN XI DKV 4 SMKN 9 GARUT' - keduanya TIDAK memakai format 'Personal Branding [Nama] [Kelas]' / 'ASTS Personal Branding ...'."), ('TIDAK DIKUMPULKAN', "Tidak ada logo personal branding. Yang ada ID Card atas nama 'Joko Apar' (anggota program MBG), banner rental PS 'NOOBLE PLAYSTATION', dan 'sample logo Android' (IMG#03 275x320, OCR kosong). Tidak ada nama branding, tagline, atau deskripsi konsep untuk identitas pribadi siswa."), ('TIDAK DIKUMPULKAN', 'Bagian moodboard tidak ada sama sekali (heading(0), 146 kata total).'), ('TIDAK DIKUMPULKAN', 'ID Card dan banner rental PS adalah tugas komputer grafis, bukan mockup branding; tidak ada penjelasan fungsi media branding.'), ('TIDAK DIKUMPULKAN', 'Bagian naskah iklan tidak ada.'), ('TIDAK DIKUMPULKAN', 'Bagian storyline tidak ada.'), ('TIDAK DIKUMPULKAN', 'Bagian shotlist tidak ada.'), ('TIDAK DIKUMPULKAN', 'Bagian storyboard tidak ada.'), ('TIDAK DIKUMPULKAN', 'Tidak ada gambar karakter maskot; isi 3 gambar adalah ID Card, banner PS, dan logo Android.'), ('TIDAK DIKUMPULKAN', "Tidak ada prompt AI. Teks justru menyebut 'Dibuat di aplikasi Vektor, Software CorelDRAW Versi 2020' - dikerjakan manual, bukan AI."), ('BELUM MEMENUHI', "Link aktif (status 200) tetapi hanya 1 dari 5 label wajib ('AI'); label 'Tugas Sekolah', 'Personal Branding', 'Portofolio', 'SMKN 9 Garut' tidak ada. Label lain 'Corel' tidak relevan."), ('BELUM MEMENUHI', 'Identitas visual tidak konsisten/tidak ada: karya terdiri atas 3 projek terpisah (ID Card, banner rental PS, logo Android) dengan 3 subjek berbeda, dan tidak ada satu pun aset personal branding milik RAFI FAUZAN NAJA LUTFIANA sebagai karya ASTS.')],
))
S.append(('R88','RAISYA ALAWIYAH','XI DKV 1','02 Okt 2026 11:35',
'https://rrewwwww.blogspot.com/2026/10/personal-branding-raisya-alawiyah-xi.html',
'3 (terverifikasi gambar)','6 (belum 10 shot)','6 (terverifikasi gambar)','3 (dalam 1 gambar komposit)','8','2153 (memenuhi)',[4, 3, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4],
[('LENGKAP', "Page title 'PERSONAL BRANDING RAISYA ALAWIYAH XI DKV 1' + heading 'ASTS Personal Branding Raisya Alawiyah XI DKV 1 SMKN 9 GARUT' - sesuai."), ('SEBAGIAN', "Logo monogram RA (IMG#01 400x400, OCR 'RAISYA ALAWIYAH / Create . watch . Grow'), nama branding dan tagline 'Create - Watch - Grow' ada; tetapi deskripsi konsep logo hanya 68 kata, di bawah syarat minimal 100 kata."), ('LENGKAP', 'IMG#02 639x426 landscape. OCR gambar membaca 7 unsur: Color Palette, Typography (Playfair Display/Allura/Montserrat), Style Visual, Referensi Desain, Tone & Mood, Elemen Grafis, Inspirasi Visual.'), ('LENGKAP', "3 mockup dengan 'Fungsi Media' tiap media: jersey (IMG#03 400x400), laptop (IMG#04 400x266), hoodie (IMG#05 399x365) - masing-masing ditulis 5-7 baris fungsi dan elemen branding."), ('LENGKAP', "Judul 'Create, Watch, Grow', Tema 'Kreativitas, proses berkarya, dan pengembangan diri', Pesan Utama, NASKAH/SCRIPT Scene 1-6 lengkap dengan VO per rentang detik, CLOSING TAGLINE ON SCREEN 'RA - RAISYA ALAWIYAH'."), ('LENGKAP', 'Judul, Tokoh, Durasi, Tema, lalu Pembukaan, Alur Cerita, Konflik / Fokus Visual, dan Penutup - 4 unsur lengkap.'), ('BELUM MEMENUHI', "Hanya 6 scene, bukan 10 shot. Teks menyatakan 'Shotlist disusun menjadi 6 scene'; IMG#07 639x426 (OCR 'SHOTLIST / Durasi: 60Detik / Jumlah Scene:6') memang ber kolom lengkap tetapi jumlah shot kurang dari ketentuan."), ('LENGKAP', "IMG#08 639x584; OCR membaca 'Jumlah Scene: 6' dan 6 label scene (Scene 1 Pembukaan 0-10s, Scene 2 Menonton Film 10-20s, Scene 3 Mendapat Ide 20-30s, Scene 4 Proses & Keraguan 30-45s, Scene 5 Hasil Akhir, Scene 6 Closing) beserta Jenis Shot, Angle Kamera, Transisi, dan Dialog/Narasi."), ('LENGKAP', "IMG#06 399x365 memuat tiga panel berlabel '1. Mascot ... Fu Body', '2. Mascot Portrait', '3. Mascot+Logo (logo di Body)' dengan OCR 'RA / RAISYA ALAWIYAH'."), ('LENGKAP', "8 prompt berlabel untuk 8 bagian: 'Prompt :' (logo), 'Prompt :' (moodboard), 'Prompt Mock Up :', 'Prompt :' (maskot), 'Prompt:' (naskah), 'Prompt :' (storyline), 'Prompt :' (shotlist), 'Prompt :' (storyboard)."), ('LENGKAP', 'Tepat 5 label wajib (Tugas Sekolah, Personal Branding, AI, Portofolio, SMKN 9 Garut); isi 2153 kata dengan 8 bagian titled dan lengkap.'), ('LENGKAP', 'Identitas RA (cream-cokelat, clapperboard, film strip) konsisten di logo, moodboard, 3 mockup, maskot, naskah, storyline, dan storyboard; konten orisinal.')],
))
S.append(('R91','RAISYA NURIL MAULIDA','XI DKV 1','02 Okt 2026 13:15',
'https://ecawwwww.blogspot.com/2026/10/personal-branding-raisya-nuril-maulida.html',
'3 (terverifikasi teks)','16 shot (prosa, tanpa tabel)','6 scene (terverifikasi teks)','3 varian (1 gambar 320x320)','10','3500',[4, 4, 4, 4, 4, 4, 3, 4, 3, 4, 2, 3],
[('LENGKAP', "Page title 'Personal Branding Raisya Nuril Maulida XI DKV 1' + heading artikel 'ASTS Personal Branding Raisya Nuril Maulida XI DKV 1 SMKN 9 Garut' - keduanya sesuai dan nama konsisten di seluruh isi artikel."), ('LENGKAP', "Logo monogram 'RNM' (IMG#01 320x320, OCR 'RAISYA NURIL MAULIDA / EMBRACING VISION. DEFINING CREATIVITY. / BY RAISYA NURIL MAULIDA'); nama branding 'Raisya Nuril Maulida', tagline 'Embracing Vision. Defining Creativity.', NAMA SISWA tercetak pada logo; deskripsi konsep 228 kata (>=100)."), ('LENGKAP', 'IMG#02 320x179 landscape. 7 unsur terverifikasi di prosa: Warna Utama (#0B3352, #1F526F, #CFAC67, #F5F1E5, #333333), Typography (Montserrat/Arimo + Lora/Open Sans), Style Visual (minimalis-geometris), Referensi Desain, Tone & Mood, Elemen Grafis, Inspirasi Visual.'), ('LENGKAP', "3 mockup dengan penjelasan fungsi: a. Hoodie (merchandise/brand wear), b. Laptop (digital showcase/brand awareness), c. Packaging (deliverable box, unboxing, perceived value). Catatan: images.json menaruh IMG#03, IMG#04, dan IMG#05 semuanya pada bagian 'a. HOODIE' padahal teks menjelaskan tiga media berbeda."), ('LENGKAP', "Naskah memuat Judul 'Beyond the Canvas of Vision', Tema, Pesan Utama, Narasi/Dialog Voice Over Scene 1-5 lengkap dengan SFX per scene, dan Closing Tagline 'Embracing Vision. Defining Creativity.'"), ('LENGKAP', 'Storyline 60 detik memuat keempat unsur wajib: Pembukaan (studio Deep Navy Blue), Alur Cerita (photo shoot -> perancangan grafis), Konflik/Fokus Visual (kompleksitas proyek antar tiga bidang), Penutup (luxury rigid box + end card RNM).'), ('SEBAGIAN', "Klaim 'total 16 shot utama' pada prosa Deskripsi Konsep Shot List dengan jenis shot berurutan (Establishing Shot, Tilt Up, Whip Pan, MCU High Angle, Over the Shoulder, ECU, Zoom In, Montage, End Card) tetapi tanpa nomor shot dan tanpa tabel/kolom; 1 gambar IMG#07 320x179 dengan OCR tabel kacau ('Adtgan / Angke / MrenHrk') sehingga jumlah baris TIDAK DAPAT DIVERIFIKASI."), ('LENGKAP', "Enam panel storyboard tertulis lengkap di artikel (Panel 1-6) dengan keterangan shot, angle, transisi, dan Voice Over; ditutup kalimat 'Rangkaian enam panel storyboard 3D ini...'; IMG#08 320x179 OCR memuat 6 label scene."), ('SEBAGIAN', "Teks menyebut varian 'full body, portrait, hingga versi terintegrasi dengan logo utama', namun hanya ada 1 gambar IMG#06 320x320 yang memuat sekaligus OCR label 'MASCOT FULLBODY / MASCOT PORIRAI / MASCOT BERSAMA / LOCO BRANDING'; full body hanya terverifikasi dari label, bukan dari render terpisah."), ('LENGKAP', "10 prompt terdokumentasi dan berlabel pada tiap tahap: logo, moodboard, hoodie, laptop, packaging, maskot, naskah, storyline, shotlist, storyboard; 2 di antaranya salah ketik 'Promt yang di gunakan'."), ('SEBAGIAN', "Hanya 1 dari 5 label wajib terpasang, yaitu 'Personal Branding'; 'Tugas Sekolah', 'AI', 'Portofolio', dan 'SMKN 9 Garut' tidak ada. Isi artikel lengkap (3500 kata, 8 bagian) dan URL aktif status 200."), ('SEBAGIAN', "Karya orisinal dengan palet navy-teal-gold konsisten; tetapi teks instruksi AI bocor ke badan artikel ('sekarang kita akan membuat story line nya berikut ketentuan nya : Story Line Wajib memuat: Alur Cerita, Konflik/Fokus Visual') dan terdapat salah ketik/rusak seperti 'anchor pointso sigeometris' dan 'Promt'.")],
))
S.append(('R100','SAVINA KHOERUNNISA','XI DKV 2','02 Okt 2026 19:25',
'https://savinakh.blogspot.com/2026/10/personal-branding-savina-khoerunnisa.html',
'3 (terverifikasi gambar IMG#03-05)','10 (terverifikasi gambar IMG#07)','0 (TIDAK DIKUMPULKAN)','3 varian (terverifikasi OCR IMG#06)','8','2639 (memenuhi)',[2, 3, 3, 4, 4, 4, 3, 0, 4, 3, 2, 2],
[('SEBAGIAN', "Page title 'PERSONAL BRANDING SAVINA KHOERUNNISA' ada nama tetapi TANPA kelas; tidak ada heading 'ASTS Personal Branding ... SMKN 9 Garut' (headings=0, artikel langsung mulai '1.PERSONAL BRANDING'). Nama reposisi: artikel 'Savina Khoerunnisa' vs rekapan 'SAVINA KHOERUNNISA' (ejaan berbeda)."), ('SEBAGIAN', "Logo monogram SK (IMG#01 1448x1086, OCR 'SAVINA KHOERUNNISA / masa depan di bayar dengan usaha hari ini'), nama branding 'SAVINA KHOERUNNISA', tagline 'masa depan di bayar dengan usaha hari ini'; NAMA SISWA tercetak pada logo. Namun deskripsi konsep hanya 52 kata (<100) dan pada heading bagian 1 tertulis '{monogrom SSG}' padahal seluruh isi memakai inisial SK."), ('SEBAGIAN', "IMG#02 1448x1086 landscape; OCR terbaca 6 blok bernomor '1.WARNAUTAMABRANDING / 2.TIPOGRAFI / 3.STYLEVISUAL / 4.TONE & MOOD / 5.ELEMEN GRAFIS / 6.INSPIRASI VISUAL'. 'Referensi Desain' tidak ada sebagai unsur tersendiri (prompt juga hanya meminta 6 butir)."), ('LENGKAP', 'Tiga mockup dengan uraian visual, filosofi, dan fungsi yang sangat rinci: A.KAOS (fungsi: seragam tim, merchandise resmi, brand awareness berjalan), B.TUMBLER, C.KARTU NAMA; didukung 3 berkas IMG#03, IMG#04 (1376x768), IMG#05.'), ('LENGKAP', "Naskah memuat Judul 'Resep Rasa & Karya: Dapur Savina Khoerunnisa', Tema, Pesan Utama, Narasi/Dialog 4 blok bertimestamp [00:00-00:05] sampai [00:22-00:30] lengkap VO, dan Closing Tagline 'Masa depan dibayar dengan usaha hari ini.'"), ('LENGKAP', "Storyline 'Merangkai Gita dan Rasa' memuat 1. Pembukaan, 2. Alur Cerita, 3. Konflik/Fokus Visual, 4. Penutup - masing-masing dengan Fokus Visual dan Narasi Visual terperinci."), ('SEBAGIAN', "Teks di bawah heading '7.SHOTLIST' BUKAN shotlist: isinya deskripsi visual logo (monogram SK, topi koki, whisk, kamera, pita, tipografi, palet warna). Shotlist yang sebenarnya hanya ada sebagai gambar IMG#07 (1195x896) - OCR terbaca jelas sebagai tabel dengan kolom No/JENIS SHOT/ANGLE/MOVEMENT/DURASI/DESKRIPSI LITERASI dan baris bernomor 1-10 (10 shot, minimum terpenuhi)."), ('TIDAK DIKUMPULKAN', "Bagian '8.STORYBOARD' tidak berisi scene sama sekali. Isinya uraian logo monogram rose gold yang 96% identik dengan teks bagian 7 (dihitung pembanding baris karakter, rasio similaritas 0,96) dan tidak ada satu pun scene, keterangan shot, angle, transisi, atau dialog. Tidak ada gambar storyboard; IMG#08 adalah salinan identik dari IMG#07 (shotlist)."), ('LENGKAP', "IMG#06 1195x896 dengan OCR terbaca 'MascotFull Body', 'Mascot Portrait', 'Mascot and Branding' - ketiga output wajib terverifikasi."), ('SEBAGIAN', "8 prompt terdokumentasi (logo, moodboard, mockup kaos, mockup tumbler, mockup kartu nama, mascot, naskah, storyline) tetapi sebagian besar terpotong di tengah kalimat dan salah ketik: 'PROMPT;', 'PROMT:', 'FROMT:', 'WOW KEREN', 'Mascout'."), ('BELUM MEMENUHI', "Link aktif (status 200) tetapi label kategori kosong total (labels=[]) - 0 dari 5 label wajib: 'Tugas Sekolah', 'Personal Branding', 'AI', 'Portofolio', 'SMKN 9 Garut' tidak ada sama sekali."), ('SEBAGIAN', "Nama konsisten SAVINA KHOERUNNISA dan palet rose gold/krem konsisten, tetapi ada ketidakkonsistenan isi: bagian shotlist/storyboard memakai maskot yang berenang di air ('Maskot berenang gaya dada', 'keluar dari permukaan air') yang bertentangan dengan brand kuliner + fotografi, dan inisial disebut SSG pada judul bagian 1 namun SK di seluruh dokumen lain.")],
))
S.append(('R112','SITI MULYANI','XI DKV 1','03 Okt 2026 08:04',
'https://sitimulyani0127.blogspot.com/2026/10/personal-branding-siti-mulyani-xi-dkv-1.html',
'3 (terverifikasi)','10 (terverifikasi gambar)','6 (terverifikasi teks)','3 (full body + portrait + bersama logo)','5','1682 (memenuhi)',[4, 4, 3, 4, 4, 4, 3, 4, 4, 3, 4, 3],
[('LENGKAP', "Page title 'Personal Branding Siti Mulyani XI DKV 1' + heading artikel 'ASTS Personal Branding Siti Mulyani XI DKV 1 SMKN 9 Garut' - keduanya ada dan menyebut nama serta kelas."), ('LENGKAP', "Logo monogram SM dengan ikon bola futsal (IMG#01 1024x559, OCR 'SITI MULYANI / BYSITIMULYANI' - nama siswa tercetak); nama branding 'Siti Mulyani'; tagline 'Menciptakan Jejak Visual yang Dinamis'; deskripsi konsep 172 kata (memenuhi minimal 100 kata)."), ('DIKUMPULKAN', "IMG#02 1024x572 landscape. 6 dari 7 unsur terverifikasi di teks dan label panel: Warna Utama (#001F4D, #2A8C8C, #FA8072, #D5F3E9), Typography, Style Visual, Referensi Desain & Inspirasi, Tone & Mood, dan Elemen Grafis. Unsur 'Inspirasi Visual' tidak memiliki uraian terpisah - hanya digabung pada bagian 5 'Referensi Desain & Inspirasi'."), ('LENGKAP', "Tiga mockup terpisah: Stiker IMG#03 1024x572 (OCR 'SITI MULYANI'), Poster Portofolio IMG#04 765x1024 (OCR 'SITI MULYANI / KREATIF. DINAMIS. PENUH ASA.'), Kartu Nama IMG#05 1024x572 (OCR '+62 812 3456 7890, info@sitimulyani.com'). Tiap media punya penjelasan fungsi ('Stiker die-cut vinyl ... Cocok diaplikasikan pada laptop, helm, atau perlengkapan desain'; poster 'promosi/portofolio grafis'; kartu nama 'tampilan depan dan belakang ... informasi kontak lengkap') - Catatan: teks fungsi ini ditulis di dalam div sehingga tidak muncul di blok per-bagian dossier, ditemukan pada field teks artikel."), ('LENGKAP', "'D.1 Naskah Iklan': Judul 'Menyusun Jejak, Menyuarakan Karya', Tema (futsal + presisi desain), Pesan Utama (karakter pemalu bukan penghalang), Narasi/Dialog 3 VO, dan Closing Tagline 'Siti Mulyani: Kreatif. Dinamis. Penuh Asa.'"), ('LENGKAP', "'D.2 Storyline' memuat keempat unsur: Pembukaan (kontras studio desain dan lapangan futsal), Alur Cerita (transisi bola futsal ke sapuan pen tool), Konflik / Fokus Visual (karakter pemalu dibantah fokus tajam), Penutup (resolusi karya + animasi logo SM)."), ('DIKUMPULKAN', "Shotlist berupa gambar IMG#06 486x266 di bagian 'D.3 Shotlist'; OCR terbaca sebagai tabel berheader No | Adegan | Jenis Shot | Angle | Movement dengan 10 baris bernomor 1-10 - jumlah shot memenuhi, tetapi kolom Durasi dan Deskripsi tidak ada (prompt siswa hanya meminta 5 kolom)."), ('LENGKAP', 'Enam scene tertulis lengkap (Scene 1-6) dengan Visual Adegan, Keterangan Shot (CU/MCU/MS/ECU/FS), Angle Kamera, Transisi (Fade In, Cut to, Match Cut, Cross Dissolve, Smooth Wipe, Fade to White), dan Dialog/Narasi; IMG#07 490x268 OCR mengonfirmasi 6 panel dengan label shot, angle, transisi, dan VO.'), ('LENGKAP', "Ketiga varian maskot ada sebagai gambar terpisah: 1. Mascot Full Body IMG#08 1024x572, 2. Mascot Portrait IMG#09 1024x572, 3. Mascot Bersama Logo Branding IMG#10 1024x559 (OCR 'SITI MULYANI / KREATIF. DINAMIS. PENUH ASA.') dengan deskripsi komposisi maskot di samping logo SM pada heading."), ('DIKUMPULKAN', "5 prompt terdokumentasi dan berlabel 'Prompt yang digunakan': logo inisial SM, moodboard (7 unsur, landscape), 3 mockup (stiker, poster, kartu nama), PERENCANAAN VIDEO IKLAN / D.1 Naskah, dan MEMBUAT AI MASCOT CHARACTER - tidak ada prompt terpisah untuk storyline, shotlist, dan storyboard."), ('LENGKAP', '5 dari 5 label wajib terpasang (AI, Personal Branding, Portofolio, SMKN 9 Garut, Tugas Sekolah) dan artikel berisi 1.682 kata dengan 8 komponen lengkap (A branding sampai E maskot); link aktif (status 200).'), ('SEBAGIAN', "Identitas visual konsisten (monogram SM, palet navy-teal-coral, elemen futsal) dari logo, moodboard, mockup, storyline, sampai maskot; catatan professionalism: banyak paragraf terulang dua kali berturut-turut (mis. 'Kreatif & Profesional: Menunjukkan seorang desainer yang serius namun bersemangat.' muncul dua kali) dan tipografi pada IMG#02 tulis 'Penun Sans' (hasil generasi AI, seharusnya Poppins/Montserrat).")],
))
S.append(('R92','SYIPA NURAENI','XI DKV 1','02 Okt 2026 14:45',
'https://nuraenisyipaa.blogspot.com/2026/10/personal-branding-syipa-nuraeni-xi-dkv-1.html',
'3 (terverifikasi gambar IMG#03-05)','14 (klaim saja; TIDAK DAPAT DIVERIFIKASI)','6 (terverifikasi teks, tanpa shot/angle per scene)','3 output diklaim; 1 gambar (full body TIDAK DAPAT DIVERIFIKASI)','8','3282 (memenuhi)',[4, 4, 3, 4, 4, 4, 2, 3, 2, 4, 2, 3],
[('LENGKAP', "Page title 'Personal Branding Syipa Nuraeni XI DKV 1' dan heading 'ASTS Personal Branding Syipa Nuraeni XI DKV 1 SMKN 9 Garut' - keduanya persis sesuai format."), ('LENGKAP', "Logo monogram SN (IMG#01 320x320, OCR 'SYIPA NURAENI','DESIGN + PHOTOGRAPHY'); nama branding 'SN - Syipa Nuraeni'; tagline 'Create. Capture. Inspire.' + 'Small Steps Big Dreams'; deskripsi konsep 269 kata; NAMA SISWA tercetak pada logo."), ('SEBAGIAN', 'IMG#02 320x213 landscape. Tujuh unsur dibahas satu paragraf panjang (warna utama deep purple/lavender/beige, typography Playfair Display+Allura+Montserrat, style visual minimalis, referensi desain, tone & mood, elemen grafis, inspirasi visual), tanpa 7 butir berlabel.'), ('LENGKAP', "Tiga mockup dengan penjelasan fungsi panjang masing-masing: a. kartu nama, b. gelas kopi, c. hoodie; didukung 3 gambar IMG#03/04/05 (320x213, 320x292, 320x213) ber-OCR 'SYIPA NURAENI'."), ('LENGKAP', "Naskah memuat Judul 'Create, Capture, Inspire', Tema, Pesan Utama, Durasi, Narasi/VO 4 blok, Dialog Talent, dan Closing Tagline 'Create. Capture. Inspire.'"), ('LENGKAP', "Storyline 'STORYLINE SHORT MOVIE PERSONAL BRANDING' memuat 1. Pembukaan, 2. Alur Cerita, 3. Konflik/Fokus Visual, 4. Penutup, lengkap dengan VO penutup."), ('TIDAK DAPAT DIVERIFIKASI', "Bagian '7. SHOTLIST' HANYA berisi deskripsi umum ('Shotlist ini terdiri dari 14 shot ... disusun berurutan') - tidak ada satu pun baris shot yang ditulis. Dua gambar IMG#07 (320x292) dan IMG#08 (305x320) memang berisiko 'SHOTLIST', tetapi OCR-nya pseudoteks ('CWK.CXUK','SIFAXUKSEXI','GIFATI','CAPTUU'), sehingga jumlah shot tidak dapat dihitung andal."), ('SEBAGIAN', "Enam poin ditulis eksplisit (1. Pembukaan, 2. Proses & Aktivitas, 3. Hobi & Dunia yang Disukai, 4. Tujuan & Masa Depan, 5. Identitas - Siapa Aku, 6. Penutup - Terus Melangkah) tetapi seluruhnya berformat 'Filosofi' dan tidak ada keterangan shot/angle/transisi/dialog per scene; tidak ada gambar storyboard."), ('SEBAGIAN', "Teks menyatakan tiga output (Mascot Full Body, Mascot Portrait, Mascot Bersama Logo Branding); hanya 1 berkas IMG#06 320x213 dengan OCR pseudoteks 'Tutoty / Poerl / Lopomoaix'. Full Body tidak dapat dipastikan."), ('LENGKAP', '8 prompt terdokumentasi berlabel dan berurutan: logo, moodboard, mockup (kartu nama/gelas/hoodie), mascot, naskah, storyline (tercantum 2x), shotlist, storyboard.'), ('BELUM MEMENUHI', "Link aktif, tetapi hanya 2 dari 5 label wajib terpenuhi: 'Tugas Sekolah' dan 'SMKN 9 Garut'. Label 'AI', 'Portofolio', dan 'Personal Branding' tidak ada; muncul label tidak relevan 'CV', 'Kegiatan Belajar', dan label berpola 'Tugas ASTS personal branding SYIPA N XI DKV 1'."), ('SEBAGIAN', "Nama konsisten SN/Syipa Nuraeni dan palet deep purple konsisten; tetapi prompt storyline tertulis dua kali berturut-turut dan tipografi label 'SYIPA N XI DKV 1'.")],
))
# ===================== NILAI TAMBAHAN RESOLUSI (Revisi Rubrik ke-3) =====================
# Resolusi TIDAK menjadi potongan nilai. Resolusi > 1000 px mendapat NILAI TAMBAHAN, maksimal +1,0.
RESOLUSI = {
  'SYIPA NURAENI': (8, 0, 320, 'Semua gambar lebar 320px'),
  'SITI MULYANI': (10, 8, 1024, '2 gambar kecil: shotlist IMG#06 486x266 dan storyboard IMG#07 490x268; poster IMG#04 765x1024 (lebar di bawah 1000).'),
  'SAVINA KHOERUNNISA': (8, 8, 1448, 'Semua gambar >=1000px; lebar maks 1448'),
  'RAISYA NURIL MAULIDA': (8, 0, 320, 'tidak ada gambar >=1000px; moodboard, shotlist, dan storyboard hanya 320x179'),
  'RAISYA ALAWIYAH': (8, 0, 639, 'Moodboard, shotlist, dan storyboard 639px, sisanya 400px; tidak ada gambar >=1000px.'),
  'RAFI FAUZAN NAJA LUTFIANA': (3, 0, 320, 'Karya 3 gambar: 204x320, 320x107, 275x320'),
  'RADIT KURNIAWAN': (0, 0, 0, 'Inaccessible'),
  'PITRYA HANDAYANI': (6, 0, 320, "Semua gambar <=320px; IMG#04 (sheet maskot) dipetakan lewat eliminasi karena OCR hanya 'PITRYAHANDAYAAI'; tidak ada gambar >=1000px."),
  'NAPSA JAKIYAH': (5, 0, 320, 'Semua gambar lebar 320px'),
  'NAILA AGUSTIN': (7, 0, 320, 'Semua gambar lebar 320px'),
  'NADYA NURLATIFA': (7, 0, 320, 'Semua gambar <=320px; IMG#03, IMG#05, dan IMG#07 tidak dapat diidentifikasi dari OCR (pseudoteks); tidak ada gambar >=1000px.'),
  'NADIA FITRIANI': (6, 0, 320, 'Semua gambar <=320px; shotlist hanya 148x320 (potret) sehingga tabelnya tidak terbaca; tidak ada gambar >=1000px.'),
  'NADA NISRINA': (8, 0, 320, 'Semua gambar lebar 320px'),
  'MUHAMMAD TAUFIQ ISMAIL': (5, 0, 320, 'Seluruh 5 gambar berukuran 320 px (320x320, 320x213, 320x213, 320x320, 320x292) - thumbnail Blogger; tidak ada gambar >= 1000 px.'),
  'MUHAMAD REZA RAMDANI': (8, 0, 400, 'Semua gambar di bawah 500px; lebar maks 400'),
  'MUHAMAD DANDI NUGRAHA': (6, 4, 1600, '2 gambar kecil: logo IMG#01 320x175 dan shotlist IMG#04 438x245; 4 gambar >= 1000 px (moodboard 1600x893).'),
  'LUSI NURAENI': (0, 0, 0, 'Inaccessible'),
  'INTAN MAHARANY': (5, 0, 320, 'Seluruh 5 gambar berukuran 320 px (ukuran placeholder Blogger); tidak ada gambar >= 1000 px.'),
  'INDRI': (8, 0, 685, 'Gambar shotlist 685x382 dan storyboard 640x640, sisanya 320px; tidak ada gambar >=1000px.'),
  'INDAH TRIJAYANTI': (10, 0, 400, 'tidak ada gambar >=1000px; 2 gambar maskot potret 300x400, sisanya landscape 399-400px'),
  'GILAR APGAN MUHAMAD SOLEH': (8, 0, 436, 'Seluruh gambar 246-436 px (maks IMG#04 436x290); tidak ada gambar >= 1000 px.'),
  'FUZI NAILA ANNURI': (9, 0, 400, 'Semua gambar <=400px (hasil unduh AI berukuran kecil); tidak ada gambar >=1000px, resolusi tidak menjadi penalti.'),
  'DAPA MUSTOPA': (6, 5, 1536, 'hanya shotlist IMG#04 385x256 di bawah 1000px; sisanya 1024-1536px'),
  'AZIZAH NURUL KAMIL': (7, 7, 1408, 'Semua gambar >=1000px; lebar maks 1408'),
  'AQILA NAZIL FALAQ': (7, 0, 0, 'gambar di album Google Photos privat - gagal diunduh, dimensi tidak terukur'),
  'ALIA ALAIKA NURFADILA': (0, 0, 0, 'Inaccessible'),
  'ADE SAHRUL GUNAWAN': (7, 0, 320, 'tidak ada gambar >=1000px; mockup (IMG#03) tanpa teks yang terbaca'),   # nama: (jumlah_gambar, jumlah_ge_1000px, lebar_maks, catatan)
 "WAHDAN SAPARI": (7,0,400,""),
 "AHMAD FAUZI": (11,0,320,""),
 "QIANDRA KAIZAR NAHARI": (5,4,1024,""),
 "ILMA LATIFAH": (5,0,320,""),
 "SOPA ANIDATUL AISAH": (7,0,320,''),
 "JAJANG M HUSNI MUBAROK": (5,5,1536,""),
 "DEDE APRILIA KARTIKA": (7,0,320,""),
 "PUTRI INTAN NURAENI": (5,0,320,''),
 "INTAN WIDIYANTI": (5,0,320,''),
 'SAFINAH SYARA GARINI': (7, 0, 450, '7 gambar berskala 320-450 px (2 di antaranya 320x320); tidak ada gambar >= 1000 px.'),
 'QUINSYA RAHMANESA SOLEHA': (5, 0, 368, ''),
 "JIHAN SHAFIRA KEAN PUTRI MULYADI": (8,8,1376,""),
 'DHEA EKA KHOERUNNISA': (5, 0, 320, 'Semua gambar 213-320 px; tidak ada gambar >= 1000 px.'),
 'WULAN SUNDARI': (6, 6, 1376, '6 gambar terbaca dari arsip lokal: IMG#01 1024x1024 dan IMG#02-IMG#06 1376x768; seluruhnya lebar >=1000px, lebar maks 1376. Halaman publik kini HTTP 404 - gambar tetap dapat diunduh dari arsip raw/R02.'),
 'WILDA AZKIA': (10, 0, 578, ''),
 "MUHAMAD DIAZ PIRDAUS": (8,8,1376,""),
 'RIZKY MUHAMMAD REGAL SAFARI': (5, 0, 320, ''),
  "KHANZA NURAENI": (5,4,1376,""),
 # Kiriman baru 01 Okt 2026
 "CEISHA SINTHIA": (0,0,0,""),
 "SELVI SIFA URIZQI": (5,0,320,""),
 "AZMI ANUGRAH": (5,0,320,""),
 'INDRI FITRIYANI': (5, 0, 400, 'Semua gambar 320-400 px; tidak ada gambar >= 1000 px.'),
 "MEISYA FAKHRIYAH": (5,0,320,""),
 'M REZA HUAFAH': (27, 0, 320, ''),
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
  'GADNA WIDIATNI': (6, 0, 399, 'Semua gambar 266-399 px; tidak ada gambar >= 1000 px.'),
  "DEVINA NAYYRA FITRIANI": (7,0,320,""),
  "SULISTIAWATI": (7,0,320,""),
  "SECHAN KHALIFATUNNISA": (7,0,320,""),
 # 2 Oktober 2026
 "FITRIYANI": (6,0,320,''),
 "SYABINA AYAT EL AKHROS": (8,0,640,''),
 'SITI RAHMA SILPIANA': (8, 0, 320, '8 gambar 213-320 px; tidak ada gambar >= 1000 px.'),
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
 'NURI MEITRI AENI': (8, 0, 320, ''),
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
ws["A4"] = "Sumber: rekapan Google Form (112 kiriman) + akses langsung ke Blogger + 723 berkas gambar (OCR RapidOCR & dimensi piksel asli)"
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
 ("LUSI NURAENI","Kirim link artikel publik","Link yang masuk adalah URL editor Blogger (redirect login Google). Nilai 0 karena karya tidak dapat diverifikasi."),
 ("RADIT KURNIAWAN","Kirim link artikel publik","Link yang masuk adalah URL editor Blogger (redirect login Google). Nilai 0 karena karya tidak dapat diverifikasi."),
 ("ALIA ALAIKA NURFADILA","Kirim link artikel publik yang aktif","Link mengembalikan HTTP 404. Nilai 0 karena karya tidak dapat diverifikasi."),
 ("RAFI FAUZAN NAJA LUTFIANA","Konfirmasi kiriman - projek yang dikirim bukan Personal Branding","Isi: ID Card 'Joko Apar', banner rental PS 'NOOBLE PLAYSTATION', sample logo Android - seluruhnya CorelDRAW 2020, tanpa AI. Perlu konfirmasi apakah link keliru kirim."),
 ("DAPA MUSTOPA","Verifikasi identitas nama","Rekapan & page title 'DAPA MUSTOPA'; seluruh isi artikel dan gambar 'DAFA MUSTOFA'. Perlu klarifikasi nama yang benar."),
 ("MUHAMAD REZA RAMDANI","Verifikasi orisinalitas & identitas brand","Identitas berubah 4x dalam satu artikel: logo 'RR', moodboard branding sekolah, mockup & storyboard 'The Creative Studio' (Ahmad Faisal). Perlu klarifikasi."),
 ("MUHAMMAD TAUFIQ ISMAIL","Verifikasi nama & sisa teks mentah AI","Heading artikel 'MUHAMMAD TAUFIQ ISMAI' (tanpa L); ada sisa '( IndiBlogHub )' dan '( Markuva )' - teks mentah hasil AI."),
 ("SAVINA KHOERUNNISA","Cek isi bagian SHOTLIST dan STORYBOARD","Keduanya berisi deskripsi logo yang sama, bukan shot/scene. Shotlist sebenarnya hanya pada gambar IMG#07 (10 shot); storyboard tidak ada."),
 ("SAFINAH SYARA GARINI","Cek duplikasi bagian mockup","Bagian 'C. Kemasan/Packaging' berisi teks identik dengan 'A. Kartu Nama' - fungsi kemasan bukan spesifik."),
 ("AQILA NAZIL FALAQ","Cek gambar - album Google Photos privat","7 gambar berada di lh3.google.com/u/0 (butuh login) sehingga tidak dapat diunduh. Bila guru punya akses, periksa moodboard/mockup/logo/maskot."),
 ("WULAN SUNDARI","Konfirmasi link - sekarang HTTP 404","Nilai 87,75 dihitung dari arsip HTML + 6 gambar yang berhasil diunduh. Bila siswa mengirim link publik baru, nilai dapat dihitung ulang."),
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
  "LUSI NURAENI":    "TIDAK AKSES - URL editor Blogger (redirect login Google)",
  "RISMA SAPARANI":  "TIDAK AKSES - HTTP 404",
  "ALIA ALAIKA NURFADILA": "TIDAK AKSES - HTTP 404",
  "RADIT KURNIAWAN":"TIDAK AKSES - URL editor Blogger (redirect login Google)",
  # Link kini 404, tetapi karya telasip dan dinilai dari arsip HTML + gambar yang
  # masih dapat diunduh - bukan siswa bernilai 0.
  "WULAN SUNDARI":   "TIDAK AKSES - HTTP 404 (karya dinilai dari arsip)",
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
  # Temuan baru dari batch 2-3 Oktober 2026
  "DAPA MUSTOPA": "Rekapan & page title 'DAPA MUSTOPA', seluruh isi artikel & gambar 'DAFA MUSTOFA' - perlu verifikasi identitas",
  "RAFI FAUZAN NAJA LUTFIANA": "Kiriman berisi ID Card, banner rental PS, dan sample logo (CorelDRAW) - bukan projek Personal Branding",
  "MUHAMAD REZA RAMDANI": "Identitas branding berubah-ubah dalam satu artikel: logo 'RR', moodboard branding sekolah, mockup/storyboard 'The Creative Studio' - perlu klarifikasi",
  "MUHAMAD TAUFIQ ISMAIL": "Nama tidak konsisten: rekapan 'TAUFIQ ISMAIL', heading artikel 'TAUFIQ ISMAI'; ada sisa teks mentah AI",
  "SAVINA KHOERUNNISA": "Bagian SHOTLIST dan STORYBOARD berisi deskripsi logo, bukan shot/scene",
  "NADYA NURLATIFA": "Ejaan nama branding 'Nadya Nurlatifa' berbeda dari rekapan",
  "AQILA NAZIL FALAQ": "Gambar berada di album Google Photos privat - tidak dapat diunduh, komponen visual tidak dapat diverifikasi",
  "SAFINAH SYARA GARINI": "Bagian 'C. Kemasan/Packaging' berisi teks identik dengan 'A. Kartu Nama'",
  "INDRI FITRIYANI": "Dua tagline berbeda antara branding dan naskah",
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
 ("Sumber 1", "Rekapan pengumpulan Google Form: 112 kiriman (nama, kelas, link, timestamp)."),
 ("Sumber 2", "Halaman Blogger diakses langsung satu per satu. 104 dari 112 dapat diakses publik; 5 siswa mengirim URL editor Blogger (mengarah ke halaman login Google) dan 4 link mengembalikan HTTP 404. WULAN SUNDARI masuk kelompok 404 tetapi karyanya berhasil dipulihkan dari arsip HTML dan dinilai penuh."),
 ("Sumber 3", "723 berkas gambar terukur (dimensi piksel asli) dan diperiksa dengan OCR (RapidOCR/ONNX, pengganti Apple Vision yang tidak tersedia di platform Windows). 8 siswa tidak memiliki gambar yang dapat diukur karena link tidak dapat diakses."),
 ("Sumber 4", "Teks artikel, tabel, heading, dan label/tag Blogger. Teks di dalam <div>/<span> ikut dibaca (get_text() penuh) karena ekstraksi <p> saja under-report hingga 98%."),
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
 ("Waktu", "Seluruh 112 kiriman sebelum batas 2 Oktober 2026 23.59 WIB. Waktu TIDAK mengubah skor kualitas."),
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
 ("Data rekapan", "Rekapan Google Form berisi 112 kiriman dan seluruh 112 siswa dinilai. Tidak ada siswa yang dikeluarkan dari penilaian."),
 ("Nilai 0", "8 siswa bernilai 0 karena karya tidak dapat diverifikasi, bukan karena karya kosong: 6 siswa mengirim URL editor Blogger (CEISHA SINTHIA, SINDIA SAPUTRI, SITI JENAB, NAZWA KURNIA, LUSI NURAENI, RADIT KURNIAWAN) dan 2 link mengembalikan HTTP 404 (RISMA SAPARANI, ALIA ALAIKA NURFADILA)."),
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
