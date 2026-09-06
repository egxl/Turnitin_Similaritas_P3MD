# 🔍 Pemeriksa Similaritas Turnitin-Style Dokumen P3MD

Aplikasi pendeteksi similaritas dokumen berbasis standar Turnitin resmi yang dioptimalkan untuk menangani ratusan dokumen tugas (seperti laporan P3MD) secara cepat, inkremental, dan dilengkapi antarmuka web (Web UI) gratis tanpa kartu kredit.

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/egxl/Turnitin_Similaritas_P3MD/blob/main/Turnitin_Similaritas_P3MD.ipynb)

---

## ⚡ Fitur Utama & Peningkatan Performa

1. **Pemrosesan Inkremental (SQLite Cache)**:
   - **Tidak perlu hitung ulang:** Dokumen lama yang sudah pernah diproses disimpan ke database SQLite (`similarity_cache.db`).
   - Saat ada **1 dokumen baru ditambahkan**, sistem **hanya membandingkan dokumen baru tersebut** terhadap dokumen yang sudah ada ($N$ pasangan, bukan $\approx 80.000$ pasangan!). Menghemat 99% waktu komputasi.
   - Pre-computed K-Gram indexing dilakukan sekali per dokumen untuk eksekusi ribuan kali lebih cepat.

2. **Bypass Batas 50 File & Auto-Sync Cache ke Google Drive**:
   - Menghilangkan batasan unduhan 50 file dari `gdown` atau Google Drive web view.
   - Menggunakan Google Drive API v3 dengan pagination resmi (`pageSize=1000` & loop `nextPageToken`).
   - **Auto-Sync Database Cache (`similarity_cache.db`)**: Otomatis mengunduh cache database dari Google Drive saat awal sesi, dan mengunggah kembali cache terbaru serta laporan Excel ke Google Drive setelah analisis. Data komputasi lama **tidak pernah hilang** meski runtime Google Colab di-restart!
   - Pemeriksaan ukuran file lokal: file yang sudah diunduh otomatis dilewati (**Smart Sync**), sehingga saat ada file baru di Drive, hanya file baru tersebut yang diunduh.
   - Tanpa perlu pengaturan folder lokal manual atau mount Google Drive yang rumit.

3. **Antarmuka Web Modern (Gradio - Uptime 72 Jam)**:
   - Dilengkapi dashboard web interaktif modern yang dapat diakses langsung dari browser komputer maupun smartphone.
   - Menghasilkan tautan publik (`https://xxxx.gradio.live`) gratis selama 72 jam per sesi tanpa perlu registrasi akun atau kartu kredit.
   - Fitur Web UI:
     - **Tampilan Langsung Data Terbaru (Tanpa Layar Kosong)**: Saat pertama kali membuka tautan Web UI, data hasil analisis terbaru, ringkasan KPI, dan file unduhan Excel langsung tampil secara instan.
     - **Kolom Pencarian Dokumen / Peserta (*Live Search Bar*)**: Memudahkan peserta mencari nama file atau namanya secara instan tanpa perlu mencari satu per satu.
     - **Tabel Rekap Berukuran Penuh (*Full-Page Table*)**: Tabel Leaderboard dirancang dengan tinggi layar penuh (~750px) agar nyaman dibaca dan memuat puluhan baris sekaligus.
     - **Folder Google Drive Terintegrasi & Clickable**: Tombol langsung untuk membuka folder Google Drive tugas P3MD tempat peserta mengunggah dokumen.
     - **Panduan Alur Kerja Jelas**: Kartu instruksi langkah-demi-langkah (Upload $\rightarrow$ Cek $\rightarrow$ Pantau $\rightarrow$ Unduh) terpampang langsung di antarmuka.
     - **Bilah Kemajuan (*Live Progress Bar*)**: Memantau setiap tahapan (sinkronisasi Drive, ekstraksi teks dokumen, perbandingan inkremental pasangan, dan pembuatan laporan Excel) secara visual dan realtime.
     - **Unduh Laporan Excel Resmi**: Unduh laporan 7-Sheet lengkap dalam format `.xlsx` dengan satu klik.

4. **Standar Parameter Resmi Turnitin**:
   - **Metode String Matching & K-Gram Shingling:** Mencocokkan deretan 6–8 kata berurutan secara verbatim.
   - **Word-Level Similarity Index:** Menghitung rasio kata cocok terhadap total kata dokumen.
   - **Eksklusi Turnitin:** Mengabaikan kutipan dalam tanda petik (`"..."` / `“...”`) dan bagian Daftar Pustaka (*Bibliography*).
   - **Skema Warna Turnitin:**
     - 🔵 Blue (0%)
     - 🟢 Green (1–24%)
     - 🟡 Yellow (25–49%)
     - 🟠 Orange (50–74%)
     - 🔴 Red (75–100%)

---

## 🚀 Cara Menjalankan

### Cara 1: Menggunakan Google Colab / Web UI (Direkomendasikan)
1. Buka notebook di [Google Colab](https://colab.research.google.com/github/egxl/Turnitin_Similaritas_P3MD/blob/main/Turnitin_Similaritas_P3MD.ipynb) (atau jalankan `python app.py` secara lokal/server).
2. Klik menu **Runtime > Run all** (`Ctrl + F9`).
3. Pada **Sel 5: Antarmuka Web Interaktif**, buka tautan publik yang muncul (contoh: `https://xxxxxxxx.gradio.live`).
4. **Alur Pengguna:**
   - Klik tombol **"📂 Buka Folder Google Drive P3MD"** untuk mengunggah dokumen tugas (`.docx`, `.pdf`, atau `.txt`).
   - Klik tombol **"🚀 Mulai Analisis Similaritas / Cek Dokumen Baru"**.
   - Pantau bilah kemajuan (*progress bar*) yang berjalan secara realtime.
   - Lihat hasil peringkat di tabel **Rekap Per Peserta (*Leaderboard*)** dan unduh laporan Excel resmi.

### Cara 2: Eksekusi Langsung Tanpa Web UI (Batch Mode)
Jika hanya ingin menjalankan analisis langsung di Colab tanpa membuka antarmuka web:
1. Pada **Sel 6: Alternatif: Eksekusi Langsung Tanpa Web UI (Batch Mode)**, centang opsi `jalankan_batch_mode`.
2. Masukkan link folder Google Drive pada kolom `direct_drive_url`.
3. Klik tombol Run pada sel tersebut. Dokumen akan langsung dianalisis dan laporan Excel akan diunduh secara otomatis.

---

## 📁 Struktur Laporan Excel Eksekutif (Executive 7-Sheet Edition)

Untuk mengatasi kebingungan saat membandingkan ratusan dokumen (misal **398 dokumen** yang menghasilkan **78.803 baris pasangan**), file Excel laporan (`Turnitin_Similarity_Report_P3MD.xlsx`) disusun ke dalam **7 lembar kerja terstruktur**:

1. **📊 Dashboard Eksekutif**:
   - Ringkasan KPI cohort (*Total Dokumen, % Lulus, % Melebihi Batas, Total Pasangan Diuji, Pasangan Gagal*).
   - Tabel distribusi 5 Tier Warna Turnitin (🔵 Blue, 🟢 Green, 🟡 Yellow, 🟠 Orange, 🔴 Red).
   - Panduan cepat membaca laporan bagi peserta maupun evaluator.

2. **👤 Rekap Per Peserta (Leaderboard)** *(Solusi Utama Masalah 397 Dokumen Lawan)*:
   - **Tepat 1 baris per peserta/dokumen** (hanya 398 baris, bukan 78.803 baris!).
   - Peserta cukup mencari namanya (`Ctrl + F`) untuk langsung melihat:
     - Status Kelulusan (`PASS` / `FAIL`).
     - **Skor Tertinggi (%)** yang didapatkan dokumennya terhadap seluruh cohort.
     - **Dokumen Lawan Paling Mirip** (*Top Matched Partner* & *2nd Match*).
     - **Rata-rata Similaritas Cohort** (membedakan kemiripan alami akibat template tugas ~2% vs plagiasi).
     - Total kata dokumen.

3. **⚠️ Investigasi Plagiasi (Flagged Pairs)**:
   - **Hanya menyaring pasangan yang MELEBIHI batas toleransi** (Skor > 15%).
   - Memungkinkan Komandan/Evaluator langsung fokus pada 5–20 kasus yang benar-benar bermasalah tanpa terganggu puluhan ribu pasangan yang bersih.
   - Dilengkapi analisis arah dugaan salinan berdasarkan perbandingan rasio kata kedua dokumen.

4. **🔍 Cek Dokumen Individu**:
   - Tampilan komparasi dua arah. Memudahkan peserta atau pembimbing memfilter 1 dokumen target di kolom A untuk melihat peringkat kemiripannya dengan seluruh dokumen lain dari tertinggi ke terendah.

5. **📝 Detail Kalimat Identik**:
   - Bukti potongan kalimat identik secara *verbatim*.
   - Pasangan yang berstatus `FAIL` diprioritaskan di baris teratas dengan penanda visual khusus.

6. **📋 Semua Pasangan (Arsip)**:
   - Dataset lengkap seluruh kombinasi perbandingan berpasangan ($N(N-1)/2$).
   - Dilengkapi *AutoFilter* aktif, *Freeze Panes*, zebra row striping, dan format persentase rapi `0.00%`.

7. **⚙️ Parameter Analisis**:
   - Catatan parameter pengujian resmi (Ambang batas Komandan 17%, Safety Cushion 15%, k-gram shingling, status filter kutipan & daftar pustaka, waktu analisis).
