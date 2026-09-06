"""
app.py
Aplikasi Web Antarmuka Turnitin Similarity Checker (P3MD) berbasis Gradio.
Dapat dijalankan secara mandiri via `python app.py` atau di-deploy ke Hugging Face Spaces / server.
"""

import os
import sys
import re
import io
import time
import json
import sqlite3
import pandas as pd
from itertools import combinations

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    import gradio as gr
except ImportError:
    gr = None

from report_generator import generate_excel_report, compute_leaderboard

# Konfigurasi Standar Turnitin Resmi
COMMANDER_THRESHOLD = 17.0
PASS_THRESHOLD = 15.0
PUBLIC_DRIVE_URL = "https://drive.google.com/drive/u/0/folders/1YbKgSou6XhmCr1CLD2dRWy_ahzQ3HFDi"

def turnitin_badge(score):
    if score == 0: return "🔵 Blue (0%)"
    elif score < 25: return "🟢 Green (1-24%)"
    elif score < 50: return "🟡 Yellow (25-49%)"
    elif score < 75: return "🟠 Orange (50-74%)"
    else: return "🔴 Red (75-100%)"

def extract_raw_text(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    text = ""
    try:
        if ext == ".docx":
            if Document is None:
                raise ImportError("Library 'python-docx' belum terinstal. Silakan jalankan: pip install python-docx")
            doc = Document(file_path)
            text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        elif ext == ".pdf":
            if PdfReader is None:
                raise ImportError("Library 'pypdf' belum terinstal. Silakan jalankan: pip install pypdf")
            reader = PdfReader(file_path)
            text = "\n".join([page.extract_text() or "" for page in reader.pages])
        elif ext == ".txt":
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                text = f.read()
    except Exception as e:
        print(f"⚠️ Gagal mengekstrak {os.path.basename(file_path)}: {e}")
    return text

def apply_turnitin_exclusions(text, drop_quotes=True, drop_bib=True):
    """Menerapkan filter eksklusi Turnitin: kutipan dan daftar pustaka."""
    if drop_bib:
        bib_pattern = r"\n\s*(?:daftar\s+pustaka|references|bibliography|rujukan)\s*[:\n]"
        parts = re.split(bib_pattern, text, flags=re.IGNORECASE)
        if len(parts) > 1:
            text = parts[0]

    if drop_quotes:
        text = re.sub(r'"[^"]*"|[“”][^“”]*[“”]', ' ', text)

    return text

def tokenize_words(text):
    return re.findall(r"\b\w+\b", text.lower())

def build_kgram_map(words, k=6):
    """Membangun map dari k-gram ke daftar index posisi munculnya."""
    kgram_map = {}
    for i in range(len(words) - k + 1):
        gram = " ".join(words[i:i+k])
        if gram not in kgram_map:
            kgram_map[gram] = []
        kgram_map[gram].append(i)
    return kgram_map

def calculate_turnitin_similarity(doc_a_words, doc_b_words, map_a, map_b, k=6):
    """Kalkulasi similaritas Turnitin menggunakan precomputed k-gram map."""
    if len(doc_a_words) < k or len(doc_b_words) < k:
        return 0.0, 0.0, []

    common_grams = set(map_a.keys()) & set(map_b.keys())
    if not common_grams:
        return 0.0, 0.0, []

    matched_indices_a = set()
    for gram in common_grams:
        for start_idx in map_a[gram]:
            for offset in range(k):
                matched_indices_a.add(start_idx + offset)

    matched_indices_b = set()
    for gram in common_grams:
        for start_idx in map_b[gram]:
            for offset in range(k):
                matched_indices_b.add(start_idx + offset)

    score_a = (len(matched_indices_a) / len(doc_a_words) * 100) if doc_a_words else 0.0
    score_b = (len(matched_indices_b) / len(doc_b_words) * 100) if doc_b_words else 0.0

    passages = []
    if matched_indices_a:
        sorted_indices = sorted(matched_indices_a)
        current_passage = [doc_a_words[sorted_indices[0]]]
        for prev_idx, curr_idx in zip(sorted_indices[:-1], sorted_indices[1:]):
            if curr_idx == prev_idx + 1:
                current_passage.append(doc_a_words[curr_idx])
            else:
                passages.append(" ".join(current_passage))
                current_passage = [doc_a_words[curr_idx]]
        passages.append(" ".join(current_passage))

    return score_a, score_b, passages

class TurnitinDBCache:
    """Manajer SQLite Cache untuk pemrosesan inkremental Turnitin."""
    def __init__(self, db_path="similarity_cache.db"):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        conn = sqlite3.connect(self.db_path)
        try:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS documents (
                    filename TEXT PRIMARY KEY,
                    file_mtime REAL,
                    file_size INTEGER,
                    word_count INTEGER,
                    words_json TEXT
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS pairs (
                    doc_a TEXT,
                    doc_b TEXT,
                    k_val INTEGER,
                    max_score REAL,
                    score_a REAL,
                    score_b REAL,
                    matches INTEGER,
                    passages_json TEXT,
                    badge TEXT,
                    status TEXT,
                    PRIMARY KEY (doc_a, doc_b, k_val)
                )
            """)
            conn.commit()
        finally:
            conn.close()

    def sync_documents(self, file_paths, drop_quotes=True, drop_bib=True, progress_callback=None):
        """Memproses hanya dokumen yang baru atau berubah."""
        doc_database = {}
        new_or_updated = 0
        loaded_from_cache = 0
        total_files = len(file_paths)

        conn = sqlite3.connect(self.db_path)
        try:
            cur = conn.cursor()
            for idx, fp in enumerate(file_paths):
                name = os.path.basename(fp)
                if progress_callback:
                    progress_callback(idx + 1, total_files, name)

                mtime = os.path.getmtime(fp)
                size = os.path.getsize(fp)

                cur.execute("SELECT file_mtime, file_size, words_json FROM documents WHERE filename = ?", (name,))
                row = cur.fetchone()

                if row and row[0] == mtime and row[1] == size:
                    words = json.loads(row[2])
                    doc_database[name] = words
                    loaded_from_cache += 1
                else:
                    raw = extract_raw_text(fp)
                    filtered = apply_turnitin_exclusions(raw, drop_quotes=drop_quotes, drop_bib=drop_bib)
                    words = tokenize_words(filtered)
                    if len(words) > 0:
                        doc_database[name] = words
                        cur.execute("""
                            INSERT OR REPLACE INTO documents (filename, file_mtime, file_size, word_count, words_json)
                            VALUES (?, ?, ?, ?, ?)
                        """, (name, mtime, size, len(words), json.dumps(words)))
                        new_or_updated += 1
            conn.commit()
        finally:
            conn.close()

        return doc_database, loaded_from_cache, new_or_updated

    def run_incremental_comparisons(self, doc_database, k_val=6, threshold=PASS_THRESHOLD, max_passages=5, progress_callback=None):
        """Menghitung hanya pasangan dokumen yang belum pernah dianalisis."""
        names = sorted(list(doc_database.keys()))
        if len(names) < 2:
            return 0, 0, 0

        # Precompute k-gram maps (dilakukan 1 kali per dokumen, menghemat ratusan ribu operasi)
        kgram_maps = {}
        for name in names:
            kgram_maps[name] = build_kgram_map(doc_database[name], k=k_val)

        # Cari pasangan yang belum ada di SQLite
        missing_pairs = []
        conn = sqlite3.connect(self.db_path)
        try:
            cur = conn.cursor()
            cur.execute("SELECT doc_a, doc_b FROM pairs WHERE k_val = ?", (k_val,))
            cached_pairs_set = set(cur.fetchall())

            all_pairs = list(combinations(names, 2))
            total_pairs = len(all_pairs)

            for da, db in all_pairs:
                pair_key = (da, db) if da < db else (db, da)
                if pair_key not in cached_pairs_set:
                    missing_pairs.append(pair_key)
        finally:
            conn.close()

        cached_count = total_pairs - len(missing_pairs)
        new_computed = 0

        if missing_pairs:
            batch = []
            conn = sqlite3.connect(self.db_path)
            total_missing = len(missing_pairs)
            step = max(1, total_missing // 50)
            try:
                for idx, (da, db) in enumerate(missing_pairs):
                    w_a, w_b = doc_database[da], doc_database[db]
                    m_a, m_b = kgram_maps[da], kgram_maps[db]
                    s_a, s_b, passages = calculate_turnitin_similarity(w_a, w_b, m_a, m_b, k=k_val)
                    max_s = max(s_a, s_b)
                    status = "PASS" if max_s <= threshold else "FAIL"
                    badge = turnitin_badge(max_s)

                    batch.append((
                        da, db, k_val, round(max_s, 2), round(s_a, 2), round(s_b, 2),
                        len(passages), json.dumps(passages[:max_passages]), badge, status
                    ))
                    new_computed += 1

                    if len(batch) >= 500:
                        conn.executemany("""
                            INSERT OR REPLACE INTO pairs
                            (doc_a, doc_b, k_val, max_score, score_a, score_b, matches, passages_json, badge, status)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, batch)
                        conn.commit()
                        batch = []

                    if progress_callback and ((idx + 1) % step == 0 or (idx + 1) == total_missing):
                        progress_callback(idx + 1, total_missing)

                if batch:
                    conn.executemany("""
                        INSERT OR REPLACE INTO pairs
                        (doc_a, doc_b, k_val, max_score, score_a, score_b, matches, passages_json, badge, status)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, batch)
                    conn.commit()
            finally:
                conn.close()

        return total_pairs, cached_count, new_computed

    def get_results_dataframe(self, active_names, k_val=6):
        """Mengambil seluruh hasil analisis yang diurutkan dari skor tertinggi."""
        conn = sqlite3.connect(self.db_path)
        try:
            cur = conn.cursor()
            cur.execute("""
                SELECT doc_a, doc_b, max_score, status, badge, matches, passages_json, score_a, score_b
                FROM pairs
                WHERE k_val = ?
                ORDER BY max_score DESC
            """, (k_val,))
            rows = cur.fetchall()

            cur.execute("SELECT filename, word_count FROM documents")
            doc_words = dict(cur.fetchall())
        finally:
            conn.close()

        names_set = set(active_names)
        filtered_rows = []
        for r in rows:
            if r[0] in names_set and r[1] in names_set:
                passages = json.loads(r[6])
                filtered_rows.append({
                    "Dokumen 1": r[0],
                    "Dokumen 2": r[1],
                    "Turnitin Max Score (%)": r[2],
                    "Status Kelulusan": r[3],
                    "Kategori Turnitin": r[4],
                    "Doc 1 Cocok di Doc 2 (%)": r[7],
                    "Doc 2 Cocok di Doc 1 (%)": r[8],
                    "Total Kata Doc 1": doc_words.get(r[0], 0),
                    "Total Kata Doc 2": doc_words.get(r[1], 0),
                    "Jumlah Blok Teks Cocok": r[5],
                    "Matches": r[5],
                    "Passages": passages,
                    "Score A": r[7],
                    "Score B": r[8]
                })

        return pd.DataFrame(filtered_rows)

def extract_folder_id(url):
    match = re.search(r"folders/([a-zA-Z0-9_-]+)", url)
    if match: return match.group(1)
    match = re.search(r"id=([a-zA-Z0-9_-]+)", url)
    if match: return match.group(1)
    return url.strip()

def upload_file_to_drive(file_path, folder_id, drive_service=None, log_callback=print):
    """Mengunggah atau memperbarui file ke folder Google Drive."""
    if not os.path.exists(file_path):
        return False
    file_name = os.path.basename(file_path)
    if not drive_service:
        try:
            from google.colab import auth
            auth.authenticate_user()
            from googleapiclient.discovery import build
            drive_service = build('drive', 'v3')
        except Exception:
            try:
                from googleapiclient.discovery import build
                drive_service = build('drive', 'v3')
            except Exception as e:
                log_callback(f"ℹ️ Google Drive upload dilewati: {e}")
                return False

    try:
        from googleapiclient.http import MediaFileUpload
        query = f"'{folder_id}' in parents and trashed = false and name = '{file_name}'"
        res = drive_service.files().list(q=query, fields="files(id, name)").execute()
        existing_files = res.get('files', [])
        media = MediaFileUpload(file_path, resumable=True)

        if existing_files:
            file_id = existing_files[0]['id']
            drive_service.files().update(
                fileId=file_id,
                media_body=media
            ).execute()
            log_callback(f"☁️ Berhasil memperbarui file di Google Drive: {file_name}")
        else:
            metadata = {
                'name': file_name,
                'parents': [folder_id]
            }
            drive_service.files().create(
                body=metadata,
                media_body=media,
                fields='id'
            ).execute()
            log_callback(f"☁️ Berhasil mengunggah file baru ke Google Drive: {file_name}")
        return True
    except Exception as e:
        log_callback(f"⚠️ Gagal mengunggah {file_name} ke Google Drive: {e}")
        return False

def sync_drive_folder(folder_url_or_id, destination="./dokumen_tugas_p3md", log_callback=print, progress_callback=None):
    """
    Sinkronisasi Google Drive dengan pagination (pageSize=1000) dan skip file lokal.
    Otomatis mendownload similarity_cache.db dan dokumen naskah tugas baru.
    """
    os.makedirs(destination, exist_ok=True)
    folder_id = extract_folder_id(folder_url_or_id)
    if not folder_id:
        log_callback("⚠️ URL Google Drive tidak valid!")
        return 0, 0

    drive_service = None
    try:
        from google.colab import auth
        auth.authenticate_user()
        from googleapiclient.discovery import build
        drive_service = build('drive', 'v3')
    except Exception as e:
        try:
            from googleapiclient.discovery import build
            drive_service = build('drive', 'v3')
        except Exception as e2:
            log_callback(f"⚠️ Info: Koneksi API Google Drive lokal/eksternal: {e} / {e2}")
            return 0, 0

    if progress_callback:
        progress_callback(0, 1, "Membaca daftar file dari Google Drive...")
    log_callback("🔍 Mengambil daftar seluruh dokumen dari Google Drive...")
    query = f"'{folder_id}' in parents and trashed = false and mimeType != 'application/vnd.google-apps.folder'"
    items = []
    page_token = None

    # Pagination loop untuk bypass limit 50 / 100 file
    while True:
        try:
            results = drive_service.files().list(
                q=query,
                pageSize=1000,
                fields="nextPageToken, files(id, name, mimeType, size)",
                pageToken=page_token
            ).execute()
            items.extend(results.get('files', []))
            page_token = results.get('nextPageToken', None)
            if not page_token:
                break
        except Exception as e:
            log_callback(f"⚠️ Gagal mendapatkan daftar file: {e}")
            break

    log_callback(f"📂 Ditemukan {len(items)} file di folder Google Drive.")

    from googleapiclient.http import MediaIoBaseDownload

    # 1. Cek & sinkronisasi file cache database atau laporan Excel dari Google Drive
    for sync_fname in ["similarity_cache.db", "Turnitin_Similarity_Report_P3MD.xlsx"]:
        remote_file = next((it for it in items if it['name'] == sync_fname), None)
        if remote_file:
            local_fpath = os.path.join(destination, sync_fname)
            rem_size = int(remote_file.get('size', 0))
            if not os.path.exists(local_fpath) or (rem_size > 0 and os.path.getsize(local_fpath) != rem_size):
                log_callback(f"📦 Mengunduh {sync_fname} dari Google Drive...")
                try:
                    req = drive_service.files().get_media(fileId=remote_file['id'])
                    with io.FileIO(local_fpath, 'wb') as fh:
                        dl_tool = MediaIoBaseDownload(fh, req)
                        d_flag = False
                        while d_flag is False:
                            _, d_flag = dl_tool.next_chunk()
                    log_callback(f"✅ {sync_fname} berhasil disinkronkan dari Google Drive!")
                except Exception as e:
                    log_callback(f"⚠️ Gagal mengunduh {sync_fname}: {e}")

    downloaded = 0
    skipped = 0
    supported_exts = {".docx", ".pdf", ".txt"}
    eligible_items = [
        item for item in items 
        if os.path.splitext(item['name'])[1].lower() in supported_exts and not item['name'].startswith("~")
    ]
    total_eligible = len(eligible_items)

    for idx, item in enumerate(eligible_items):
        file_id = item['id']
        file_name = item['name']
        local_path = os.path.join(destination, file_name)
        remote_size = int(item.get('size', 0))

        # Cek apakah file sudah ada secara lokal dengan ukuran sama
        if os.path.exists(local_path) and remote_size > 0:
            if os.path.getsize(local_path) == remote_size:
                skipped += 1
                if progress_callback:
                    progress_callback(idx + 1, total_eligible, f"File sudah ada (dilewati): {file_name}")
                continue

        if progress_callback:
            progress_callback(idx + 1, total_eligible, f"Mengunduh ({downloaded + 1}): {file_name}")

        # Unduh file baru/berubah dengan retry
        success = False
        for attempt in range(3):
            try:
                request = drive_service.files().get_media(fileId=file_id)
                with io.FileIO(local_path, 'wb') as fh:
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while done is False:
                        status, done = downloader.next_chunk()
                downloaded += 1
                log_callback(f"  📥 Diunduh ({downloaded}): {file_name}")
                success = True
                break
            except Exception as e:
                time.sleep(1 + attempt * 2)

        if not success:
            log_callback(f"  ⚠️ Gagal mengunduh: {file_name}")

    log_callback(f"✅ Sinkronisasi selesai: {downloaded} file baru diunduh, {skipped} file sudah ada dilewati.")
    return downloaded, skipped

def load_latest_leaderboard(target_dir="./dokumen_tugas_p3md"):
    """
    Memuat data laporan terbaru saat aplikasi pertama kali dibuka.
    Mengecek direktori target, direktori saat ini, atau mengunduh baseline jika belum ada.
    """
    excel_candidates = [
        os.path.join(target_dir, "Turnitin_Similarity_Report_P3MD.xlsx"),
        "Turnitin_Similarity_Report_P3MD.xlsx"
    ]
    
    excel_path = None
    for p in excel_candidates:
        if os.path.exists(p):
            excel_path = p
            break

    # Jika file belum ada, coba unduh baseline dari GitHub raw
    if not excel_path:
        raw_url = "https://raw.githubusercontent.com/egxl/Turnitin_Similaritas_P3MD/main/Turnitin_Similarity_Report_P3MD.xlsx"
        try:
            import urllib.request
            target_p = os.path.join(target_dir, "Turnitin_Similarity_Report_P3MD.xlsx")
            os.makedirs(target_dir, exist_ok=True)
            urllib.request.urlretrieve(raw_url, target_p)
            if os.path.exists(target_p):
                excel_path = target_p
        except Exception:
            pass

    if excel_path and os.path.exists(excel_path):
        try:
            df = pd.read_excel(excel_path, sheet_name="👤 Rekap Per Peserta")
            failed = sum(1 for s in df["Status Kelulusan"] if str(s).upper() == "FAIL")
            passed = len(df) - failed
            total_pairs = (len(df) * (len(df) - 1)) // 2
            summary_md = f"""### 📊 Ringkasan Eksekutif Similaritas Cohort P3MD (Data Terbaru)
- **Total Dokumen Peserta:** {len(df)} file
- **Kelulusan Cohort:** ✅ **{passed} LULUS** ({passed/len(df)*100:.1f}%) | ❌ **{failed} MELEBIHI BATAS** ({failed/len(df)*100:.1f}%)
- **Total Pasangan Diuji:** {total_pairs:,} pasang
- ℹ️ *Data di bawah adalah hasil analisis tersimpan terbaru. Unggah file tugas baru ke Google Drive lalu klik tombol **"🚀 Mulai Analisis Similaritas / Cek Dokumen Baru"** untuk memperbarui data.*
"""
            return summary_md, df, excel_path
        except Exception:
            pass

    # Fallback kosong jika benar-benar belum ada data sama sekali
    empty_df = pd.DataFrame(columns=[
        "Rank", "Nama Dokumen (Peserta)", "Status Kelulusan", "Skor Tertinggi (%)", 
        "Kategori Turnitin", "Pasangan Paling Mirip (Top Match)", "Skor Match #1 (%)", 
        "Pasangan Match #2", "Skor Match #2 (%)", "Rata-rata Similaritas Cohort (%)", 
        "Jumlah Pasangan > Batas", "Total Kata"
    ])
    default_md = """### 📊 Ringkasan Eksekutif Similaritas Cohort P3MD
Belum ada data analisis tersimpan. Silakan unggah file tugas ke Google Drive lalu klik tombol **"🚀 Mulai Analisis Similaritas / Cek Dokumen Baru"** di atas.
"""
    return default_md, empty_df, None

def run_analysis_pipeline(drive_url, min_words, pass_thresh, drop_quotes, drop_bib, force_recompute, progress=None):
    if progress is None and gr is not None:
        progress = gr.Progress()
    if progress is None:
        progress = lambda f, desc="": None
    target_dir = "./dokumen_tugas_p3md"
    os.makedirs(target_dir, exist_ok=True)
    db_path = os.path.join(target_dir, "similarity_cache.db")
    
    if force_recompute and os.path.exists(db_path):
        try:
            os.remove(db_path)
        except:
            pass

    progress(0.02, desc="Menyiapkan sistem analisis...")
    log_messages = []
    def log(msg):
        log_messages.append(msg)
        print(msg)

    # 1. Sinkronisasi Dokumen Google Drive (0.05 -> 0.25)
    active_drive_url = (drive_url or "").strip() or PUBLIC_DRIVE_URL
    progress(0.05, desc="Menghubungi Google Drive...")
    
    def drive_prog_cb(curr, total, desc_text):
        frac = 0.05 + 0.20 * (curr / max(total, 1))
        progress(frac, desc=f"Google Drive [{curr}/{total}]: {desc_text[:40]}")

    dl, sk = sync_drive_folder(active_drive_url, target_dir, log_callback=log, progress_callback=drive_prog_cb)
    
    # 2. Pindai Dokumen Lokal
    supported_exts = {".docx", ".pdf", ".txt"}
    file_paths = []
    for root, _, f_list in os.walk(target_dir):
        for f in f_list:
            if os.path.splitext(f)[1].lower() in supported_exts and not f.startswith("~"):
                file_paths.append(os.path.join(root, f))

    if len(file_paths) < 2:
        return (
            f"⚠️ Ditemukan {len(file_paths)} dokumen di folder tugas. Minimal diperlukan 2 dokumen untuk analisis.\n\n"
            f"Silakan unggah dokumen ke folder Google Drive terlebih dahulu:\n{active_drive_url}",
            pd.DataFrame(),
            None,
            pd.DataFrame(),
            ""
        )

    # 3. Sinkronisasi SQLite Dokumen & Ekstraksi Teks (0.25 -> 0.50)
    progress(0.25, desc="Mempersiapkan database cache SQLite...")
    cache = TurnitinDBCache(db_path=db_path)
    
    def doc_prog_cb(curr, total, name):
        frac = 0.25 + 0.25 * (curr / max(total, 1))
        progress(frac, desc=f"Ekstraksi teks [{curr}/{total}]: {name[:35]}")

    doc_db, loaded_c, new_c = cache.sync_documents(
        file_paths, drop_quotes=drop_quotes, drop_bib=drop_bib, progress_callback=doc_prog_cb
    )

    # 4. Hitung Inkremental Pasangan (0.50 -> 0.85)
    progress(0.50, desc="Kalkulasi similaritas inkremental Turnitin...")
    def pair_prog_cb(curr, total):
        frac = 0.50 + 0.35 * (curr / max(total, 1))
        progress(frac, desc=f"Menghitung pasangan baru [{curr}/{total}]...")

    total_pairs, cached_pairs, new_pairs = cache.run_incremental_comparisons(
        doc_db, k_val=int(min_words), threshold=float(pass_thresh), progress_callback=pair_prog_cb
    )

    # 5. Buat Laporan Excel Komprehensif (0.85 -> 0.95)
    progress(0.85, desc="Menyusun data rekapitulasi...")
    df_results = cache.get_results_dataframe(list(doc_db.keys()), k_val=int(min_words))
    
    # Kumpulkan matched passages untuk detail bukti teks
    all_passages = []
    for _, r in df_results.iterrows():
        if r.get("Matches", 0) > 0 and isinstance(r.get("Passages"), list):
            s_val = float(r.get("Turnitin Max Score (%)", 0.0))
            p_status = "FAIL" if s_val > float(pass_thresh) else "PASS"
            for p_text in r["Passages"]:
                all_passages.append({
                    "doc1": r["Dokumen 1"],
                    "doc2": r["Dokumen 2"],
                    "score": s_val,
                    "status": p_status,
                    "text": p_text
                })

    progress(0.90, desc="Menyusun workbook Excel 7-Sheet...")
    excel_path = os.path.join(target_dir, "Turnitin_Similarity_Report_P3MD.xlsx")
    generate_excel_report(
        df_results, len(doc_db), min_words=int(min_words), 
        commander_threshold=COMMANDER_THRESHOLD, pass_threshold=float(pass_thresh),
        matched_passages=all_passages, output_file=excel_path
    )

    # Hitung Leaderboard 1 Baris Per Peserta untuk Tampilan Web
    progress(0.95, desc="Menghasilkan Leaderboard per peserta...")
    leaderboard = compute_leaderboard(df_results, threshold=float(pass_thresh))
    for i, item in enumerate(leaderboard, 1):
        item["Rank"] = i
        item["Nama Dokumen (Peserta)"] = item["Dokumen"]
        item["Pasangan Paling Mirip (Top Match)"] = item["Top Matched Document"]
        item["Skor Match #1 (%)"] = item["Top Match Score (%)"]
        item["Pasangan Match #2"] = item["2nd Matched Document"]
        item["Skor Match #2 (%)"] = item["2nd Match Score (%)"]

    df_display = pd.DataFrame(leaderboard)[[
        "Rank", "Nama Dokumen (Peserta)", "Status Kelulusan", "Skor Tertinggi (%)", 
        "Kategori Turnitin", "Pasangan Paling Mirip (Top Match)", "Skor Match #1 (%)", 
        "Pasangan Match #2", "Skor Match #2 (%)", "Rata-rata Similaritas Cohort (%)", 
        "Jumlah Pasangan > Batas", "Total Kata"
    ]]

    # 6. Otomatis Unggah Cache Database dan Laporan Excel ke Google Drive (0.95 -> 0.99)
    folder_id = extract_folder_id(active_drive_url)
    drive_upload_success = False
    if folder_id:
        progress(0.97, desc="Mengunggah cache database ke Google Drive...")
        up_db = upload_file_to_drive(db_path, folder_id, log_callback=log)
        up_xl = upload_file_to_drive(excel_path, folder_id, log_callback=log)
        drive_upload_success = up_db or up_xl

    failed_docs_count = sum(1 for d in leaderboard if d["Status Kelulusan"] == "FAIL")
    passed_docs_count = len(leaderboard) - failed_docs_count
    fail_pairs_count = len(df_results[df_results["Turnitin Max Score (%)"] > float(pass_thresh)])

    sync_note = "☁️ **Cache database & Excel tersinkron ke Google Drive.**" if drive_upload_success else "💾 **Cache tersimpan secara lokal.**"

    summary_md = f"""### 📊 Ringkasan Eksekutif Similaritas Cohort P3MD
- **Total Dokumen Peserta:** {len(doc_db)} file (📦 Dari Cache: {loaded_c}, 🆕 Baru Diunduh/Diproses: {new_c})
- **Kelulusan Cohort:** ✅ **{passed_docs_count} LULUS** ({passed_docs_count/len(leaderboard)*100:.1f}%) | ❌ **{failed_docs_count} MELEBIHI BATAS** ({failed_docs_count/len(leaderboard)*100:.1f}%)
- **Total Pasangan Diuji:** {total_pairs:,} pasang (⚡ Dari Cache: {cached_pairs:,}, 🔍 Baru Dihitung: {new_pairs:,})
- **Pasangan Melanggar Batas ({pass_thresh}%):** {fail_pairs_count} pasang
- **Waktu Hemat:** ~{round((cached_pairs * 0.005) / 60, 1)} menit berkat SQLite Cache!
- {sync_note}
"""
    progress(1.0, desc="Selesai!")
    return summary_md, df_display, excel_path, df_display, ""


# Setup Antarmuka Gradio
def build_gradio_app():
    if gr is None:
        raise ImportError("Gradio belum terpasang. Silakan jalankan: pip install gradio")

    init_summary, init_df, init_excel = load_latest_leaderboard()

    with gr.Blocks(title="Turnitin Document Similarity - P3MD", theme=gr.themes.Soft(), css=".dataframe-table { font-size: 13.5px !important; }") as demo:
        gr.Markdown("""
        # 🔍 Turnitin Document Similarity Checker (P3MD)
        **Sistem Deteksi Similaritas Dokumen Tugas Cohort P3MD Berbasis Standar Turnitin Resmi**
        """)

        with gr.Group():
            gr.Markdown(f"""
            ### 📋 Alur Kerja Pengumpulan Dokumen & Pemeriksaan
            Ikuti 4 langkah mudah berikut untuk memeriksa dokumen tugas Anda:

            1. **Unggah File Tugas ke Google Drive:**  
               Klik tombol **"📂 Buka Folder Google Drive P3MD"** di bawah untuk membuka folder pengumpulan tugas. Masukkan naskah tugas Anda (format `.docx`, `.pdf`, atau `.txt`) langsung ke dalam folder tersebut.
            2. **Jalankan Analisis Similaritas:**  
               Setelah file berhasil diunggah ke Google Drive, klik tombol **"🚀 Mulai Analisis Similaritas / Cek Dokumen Baru"**. Sistem akan otomatis mendeteksi dan mengunduh file baru Anda tanpa mengulang unduhan file lama.
            3. **Pantau Kemajuan (*Progress Bar*):**  
               Bilah kemajuan di bagian atas akan menampilkan progres secara realtime mulai dari sinkronisasi Google Drive, ekstraksi teks dokumen, kalkulasi pasangan Turnitin, hingga auto-upload cache database.
            4. **Lihat Hasil & Unduh Laporan Excel:**  
               Gunakan fitur **Cari Dokumen** di bawah untuk menemukan nama dokumen Anda pada tabel **Rekap Per Peserta (*Leaderboard*)**, dan klik tombol unduh untuk mengunduh laporan resmi Excel 7-Sheet lengkap.
            """)

            with gr.Row():
                drive_link_btn = gr.Button(
                    "📂 Buka Folder Google Drive P3MD (Upload Dokumen Di Sini) ↗",
                    variant="secondary",
                    size="lg",
                    link=PUBLIC_DRIVE_URL
                )
            gr.Markdown(f"🔗 *Tautan Alternatif Folder Drive:* [{PUBLIC_DRIVE_URL}]({PUBLIC_DRIVE_URL})")

        with gr.Row():
            run_btn = gr.Button("🚀 Mulai Analisis Similaritas / Cek Dokumen Baru", variant="primary", size="lg")

        with gr.Accordion("⚙️ Parameter Analisis & Pengaturan Lanjutan (Opsional)", open=False):
            drive_input = gr.Textbox(
                label="📁 Link Folder Google Drive (Default: Folder Publik Tugas P3MD)",
                value=PUBLIC_DRIVE_URL,
                placeholder="https://drive.google.com/drive/folders/..."
            )
            with gr.Row():
                min_words_slider = gr.Slider(minimum=4, maximum=12, value=6, step=1, label="Min Consecutive Words (Standar Turnitin: 6)")
                thresh_slider = gr.Slider(minimum=5.0, maximum=50.0, value=15.0, step=1.0, label="Batas Toleransi Kelulusan (%)")
            with gr.Row():
                quotes_cb = gr.Checkbox(value=True, label="Abaikan Kutipan (\" \")")
                bib_cb = gr.Checkbox(value=True, label="Abaikan Daftar Pustaka")
                force_cb = gr.Checkbox(value=False, label="Paksa Hitung Ulang Semua (Reset Cache)")

        status_output = gr.Markdown(value=init_summary)
        with gr.Row():
            download_btn = gr.File(
                value=init_excel,
                label="📥 Unduh Laporan Excel Resmi (Turnitin_Similarity_Report_P3MD.xlsx)"
            )
        
        # Search Bar & Filter Controls
        with gr.Row():
            search_input = gr.Textbox(
                label="🔍 Cari Dokumen / Nama Peserta",
                placeholder="Ketik nama file atau peserta untuk menyaring tabel (contoh: Fauzi, Tariq, Rayga, atau status: PASS/FAIL)...",
                scale=5
            )
            reset_search_btn = gr.Button("🔄 Reset Pencarian", scale=1, variant="secondary")

        # Full Page Recap Table
        table_output = gr.Dataframe(
            value=init_df,
            label="👤 Rekap Hasil Per Peserta (Leaderboard 1 Baris Per Dokumen - Diurutkan dari Skor Tertinggi)",
            interactive=False,
            wrap=True,
            height=750
        )

        current_df_state = gr.State(value=init_df)

        def filter_table(query, full_df):
            if full_df is None or len(full_df) == 0:
                return full_df
            if not query or not query.strip():
                return full_df
            q = query.strip().lower()
            mask = full_df.astype(str).apply(lambda row: row.str.lower().str.contains(q, regex=False).any(), axis=1)
            return full_df[mask]

        def reset_search(full_df):
            return "", full_df

        search_input.change(fn=filter_table, inputs=[search_input, current_df_state], outputs=[table_output])
        reset_search_btn.click(fn=reset_search, inputs=[current_df_state], outputs=[search_input, table_output])

        run_btn.click(
            fn=run_analysis_pipeline,
            inputs=[drive_input, min_words_slider, thresh_slider, quotes_cb, bib_cb, force_cb],
            outputs=[status_output, table_output, download_btn, current_df_state, search_input]
        )

    return demo

if __name__ == "__main__":
    if gr is None:
        print("⚠️ Gradio belum terpasang. Jalankan: pip install gradio")
        sys.exit(1)
    demo = build_gradio_app()
    demo.queue().launch(share=True, debug=False)
