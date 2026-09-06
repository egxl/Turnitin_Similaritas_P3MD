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
import threading
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

from report_generator import generate_excel_report, compute_leaderboard, get_turnitin_tier

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
        elif ext in [".txt", ".md", ".text"]:
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
                        if row is not None:
                            # Dokumen yang sudah ada diperbarui: hapus data pasangan lama agar dihitung ulang otomatis
                            cur.execute("DELETE FROM pairs WHERE doc_a = ? OR doc_b = ?", (name, name))
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

    def cleanup_deleted_documents(self, active_filenames):
        """Menghapus dokumen dan pasangannya dari cache jika file sudah tidak ada di disk."""
        conn = sqlite3.connect(self.db_path)
        try:
            cur = conn.cursor()
            cur.execute("SELECT filename FROM documents")
            cached_docs = [r[0] for r in cur.fetchall()]
            active_set = set(active_filenames)
            pruned_count = 0
            for d in cached_docs:
                if d not in active_set:
                    cur.execute("DELETE FROM documents WHERE filename = ?", (d,))
                    cur.execute("DELETE FROM pairs WHERE doc_a = ? OR doc_b = ?", (d, d))
                    pruned_count += 1
            conn.commit()
            return pruned_count
        finally:
            conn.close()

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
            if not os.path.exists(local_fpath):
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

    # 3. Prune file lokal lama yang sudah dihapus/direname di Google Drive (mencegah duplikat tugas)
    remote_names = {it['name'] for it in eligible_items}
    pruned = 0
    if os.path.exists(destination):
        for local_f in os.listdir(destination):
            ext = os.path.splitext(local_f)[1].lower()
            if ext in supported_exts and not local_f.startswith("~"):
                if local_f not in remote_names:
                    del_path = os.path.join(destination, local_f)
                    try:
                        os.remove(del_path)
                        pruned += 1
                        log_callback(f"  🗑️ Menghapus file lokal lama yang sudah dihapus di Drive: {local_f}")
                    except Exception:
                        pass
        if pruned > 0:
            log_callback(f"🧹 Membersihkan {pruned} file lokal yang tidak ada lagi di Google Drive.")

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

pipeline_lock = threading.Lock()
last_pipeline_run_time = 0.0
pipeline_start_time = 0.0
active_jobs_count = 0
jobs_lock = threading.Lock()

def get_current_queue_status():
    """
    Mengembalikan status antrean server secara realtime tanpa memicu antrean analisis.
    """
    global active_jobs_count, last_pipeline_run_time, pipeline_start_time
    
    target_dir = "./dokumen_tugas_p3md"
    db_candidates = [
        os.path.join(target_dir, "similarity_cache.db"),
        "similarity_cache.db"
    ]
    doc_count = 0
    pair_count = 0
    for p in db_candidates:
        if os.path.exists(p):
            try:
                conn = sqlite3.connect(p)
                cur = conn.cursor()
                cur.execute("SELECT count(*) FROM documents")
                doc_count = cur.fetchone()[0]
                cur.execute("SELECT count(*) FROM pairs")
                pair_count = cur.fetchone()[0]
                conn.close()
                break
            except Exception:
                pass

    with jobs_lock:
        current_queued = active_jobs_count

    now = time.time()
    
    if current_queued > 0 or pipeline_lock.locked():
        elapsed = int(now - pipeline_start_time) if pipeline_start_time > 0 else 0
        waiting = max(0, current_queued - 1)
        status_html = f"""<div style="padding: 10px 14px; border-radius: 8px; background: #fff3cd; border: 1.5px solid #ffeeba; color: #856404; margin-bottom: 12px;">
    <div style="font-size: 14.5px; font-weight: bold; display: flex; align-items: center; justify-content: space-between;">
        <span>⏳ <b>Status Server: SEDANG MEMPROSES ANALISIS</b></span>
        <span style="font-size: 12px; background: #ffe8a1; padding: 2px 8px; border-radius: 10px; color: #664d03;">{current_queued} Tugas Aktif</span>
    </div>
    <div style="font-size: 12.5px; margin-top: 4px; color: #664d03;">
        • <b>Proses Berjalan:</b> 1 analisis sedang aktif ({elapsed}s berjalan) &nbsp;|&nbsp; <b>Menunggu di Antrean:</b> {waiting} tugas<br/>
        • <b>Database Saat Ini:</b> {doc_count} dokumen tersimpan ({pair_count:,} pasangan teranalisis)
    </div>
</div>"""
    else:
        last_updated_str = "Belum pernah dijalankan"
        if last_pipeline_run_time > 0:
            diff_m = int((now - last_pipeline_run_time) / 60)
            if diff_m < 1:
                last_updated_str = "Baru saja (< 1 menit yang lalu)"
            elif diff_m < 60:
                last_updated_str = f"{diff_m} menit yang lalu"
            else:
                last_updated_str = f"{diff_m // 60} jam yang lalu"

        status_html = f"""<div style="padding: 10px 14px; border-radius: 8px; background: #d4edda; border: 1.5px solid #c3e6cb; color: #155724; margin-bottom: 12px;">
    <div style="font-size: 14.5px; font-weight: bold; display: flex; align-items: center; justify-content: space-between;">
        <span>🟢 <b>Status Server: KOSONG & SIAP (IDLE)</b></span>
        <span style="font-size: 12px; background: #c3e6cb; padding: 2px 8px; border-radius: 10px; color: #0f5132;">0 Antrean Menunggu</span>
    </div>
    <div style="font-size: 12.5px; margin-top: 4px; color: #155724;">
        • <b>Antrean Bebas:</b> Tidak ada proses berjalan. Analisis baru dapat langsung dimulai tanpa menunggu.<br/>
        • <b>Terakhir Disinkronkan:</b> {last_updated_str} &nbsp;|&nbsp; <b>Database:</b> {doc_count} dokumen tersimpan ({pair_count:,} pasangan)
    </div>
</div>"""

    return status_html

def run_analysis_pipeline(drive_url, min_words, pass_thresh, drop_quotes, drop_bib, force_recompute, progress=None):
    global last_pipeline_run_time, pipeline_start_time, active_jobs_count
    if progress is None and gr is not None:
        progress = gr.Progress()
    if progress is None:
        progress = lambda f, desc="": None
    target_dir = "./dokumen_tugas_p3md"
    os.makedirs(target_dir, exist_ok=True)
    db_path = os.path.join(target_dir, "similarity_cache.db")
    excel_path = os.path.join(target_dir, "Turnitin_Similarity_Report_P3MD.xlsx")

    with jobs_lock:
        active_jobs_count += 1

    try:
        with pipeline_lock:
            pipeline_start_time = time.time()
        now = time.time()
        # Pintasan Cerdas: Jika analisis baru selesai < 30 detik lalu dan bukan force_recompute
        if not force_recompute and (now - last_pipeline_run_time < 30) and os.path.exists(excel_path):
            progress(1.0, desc="Data baru saja diperbarui oleh antrean sebelumnya...")
            init_summary, init_df, init_excel = load_latest_leaderboard(target_dir=target_dir)
            fast_summary = "⚡ **Data sudah mutakhir.** Analisis cohort baru saja selesai diproses oleh antrean sebelumnya.\n\n" + init_summary
            return fast_summary, init_df, init_excel, init_df, ""

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
        
        # Bersihkan dokumen cache yang file fisiknya sudah dihapus dari disk
        active_basenames = [os.path.basename(p) for p in file_paths]
        pruned_c = cache.cleanup_deleted_documents(active_basenames)
        if pruned_c > 0:
            log(f"🧹 Menghapus {pruned_c} dokumen usang dari database cache SQLite.")
        
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

        # 5. Rekapitulasi Data & Pembuatan Laporan Excel
        progress(0.85, desc="Menyusun data rekapitulasi...")
        df_results = cache.get_results_dataframe(list(doc_db.keys()), k_val=int(min_words))
        
        # Cek apakah pembuatan ulang Excel dan upload Drive bisa dilewati
        skip_excel_and_upload = (new_c == 0 and new_pairs == 0 and not force_recompute and os.path.exists(excel_path))

        if not skip_excel_and_upload:
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
            generate_excel_report(
                df_results, len(doc_db), min_words=int(min_words), 
                commander_threshold=COMMANDER_THRESHOLD, pass_threshold=float(pass_thresh),
                matched_passages=all_passages, output_file=excel_path
            )
        else:
            log("⚡ Tidak ada dokumen/pasangan baru. Melewati pembuatan ulang Excel 7-sheet & upload Google Drive.")
            progress(0.92, desc="Data mutakhir, menggunakan laporan Excel yang sudah ada...")

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
        if folder_id and not skip_excel_and_upload:
            progress(0.97, desc="Mengunggah cache database ke Google Drive...")
            up_db = upload_file_to_drive(db_path, folder_id, log_callback=log)
            up_xl = upload_file_to_drive(excel_path, folder_id, log_callback=log)
            drive_upload_success = up_db or up_xl
        elif skip_excel_and_upload:
            drive_upload_success = True

        failed_docs_count = sum(1 for d in leaderboard if d["Status Kelulusan"] == "FAIL")
        passed_docs_count = len(leaderboard) - failed_docs_count
        fail_pairs_count = len(df_results[df_results["Turnitin Max Score (%)"] > float(pass_thresh)])

        sync_note = "☁️ **Cache database & Excel tersinkron ke Google Drive.**" if drive_upload_success else "💾 **Cache tersimpan secara lokal.**"
        recalc_note = "⚡ **Hemat Waktu:** Tidak ada dokumen baru, pembuatan Excel dilewati." if skip_excel_and_upload else f"⏱️ **Waktu Hemat:** ~{round((cached_pairs * 0.005) / 60, 1)} menit berkat SQLite Cache!"

        summary_md = f"""### 📊 Ringkasan Eksekutif Similaritas Cohort P3MD
- **Total Dokumen Peserta:** {len(doc_db)} file (📦 Dari Cache: {loaded_c}, 🆕 Baru Diunduh/Diproses: {new_c})
- **Kelulusan Cohort:** ✅ **{passed_docs_count} LULUS** ({passed_docs_count/len(leaderboard)*100:.1f}%) | ❌ **{failed_docs_count} MELEBIHI BATAS** ({failed_docs_count/len(leaderboard)*100:.1f}%)
- **Total Pasangan Diuji:** {total_pairs:,} pasang (⚡ Dari Cache: {cached_pairs:,}, 🔍 Baru Dihitung: {new_pairs:,})
- **Pasangan Melanggar Batas ({pass_thresh}%):** {fail_pairs_count} pasang
- {recalc_note}
- {sync_note}
"""
        last_pipeline_run_time = time.time()
        progress(1.0, desc="Selesai!")
        return summary_md, df_display, excel_path, df_display, ""
    finally:
        with jobs_lock:
            active_jobs_count = max(0, active_jobs_count - 1)

def check_single_document(file_obj, min_words=6, pass_thresh=PASS_THRESHOLD, drop_quotes=True, drop_bib=True, progress=None):
    """
    Memeriksa similaritas 1 dokumen yang diunggah langsung ke Gradio terhadap seluruh
    dokumen cohort yang tersimpan di cache SQLite (selesai dalam ~1-2 detik).
    """
    if progress is None and gr is not None:
        progress = gr.Progress()
    if progress is None:
        progress = lambda f, desc="": None

    if file_obj is None:
        return "⚠️ Silakan pilih/unggah file dokumen (.docx, .pdf, atau .txt) terlebih dahulu.", pd.DataFrame(), ""

    fp = file_obj if isinstance(file_obj, str) else getattr(file_obj, "name", str(file_obj))
    uploaded_name = getattr(file_obj, "orig_name", None) or os.path.basename(fp)
    progress(0.1, desc=f"Membaca file: {uploaded_name}...")

    raw_text = extract_raw_text(fp)
    if not raw_text.strip():
        return (
            f"⚠️ Dokumen **{uploaded_name}** kosong atau teks tidak dapat diekstrak dari format tersebut.",
            pd.DataFrame(),
            ""
        )

    filtered = apply_turnitin_exclusions(raw_text, drop_quotes=drop_quotes, drop_bib=drop_bib)
    words = tokenize_words(filtered)
    k_val = int(min_words)
    threshold = float(pass_thresh)

    if len(words) < k_val:
        return (
            f"⚠️ Dokumen hanya memiliki {len(words)} kata setelah difilter (minimal {k_val} kata).",
            pd.DataFrame(),
            ""
        )

    progress(0.3, desc="Menghubungkan ke database cache cohort...")
    target_dir = "./dokumen_tugas_p3md"
    db_candidates = [
        os.path.join(target_dir, "similarity_cache.db"),
        "similarity_cache.db"
    ]
    db_path = None
    for p in db_candidates:
        if os.path.exists(p):
            db_path = p
            break

    if not db_path:
        return (
            "⚠️ Basis data cache cohort belum ditemukan. Silakan jalankan sinkronisasi pada tab Rekapitulasi Cohort terlebih dahulu.",
            pd.DataFrame(),
            ""
        )

    progress(0.4, desc="Membangun indeks k-gram dokumen unggahan...")
    map_uploaded = build_kgram_map(words, k=k_val)

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT filename, words_json FROM documents")
        cohort_docs = cur.fetchall()
    finally:
        conn.close()

    if not cohort_docs:
        return (
            "⚠️ Basis data cache cohort belum memiliki dokumen tersimpan.",
            pd.DataFrame(),
            ""
        )

    progress(0.5, desc=f"Membandingkan terhadap {len(cohort_docs)} dokumen cohort...")
    comparison_results = []
    clean_uploaded_name = re.sub(r"^[a-f0-9]{16,}_", "", uploaded_name.lower())

    for idx, (doc_name, w_json) in enumerate(cohort_docs):
        clean_cohort_name = re.sub(r"^[a-f0-9]{16,}_", "", doc_name.lower())
        # Hindari membandingkan dokumen dengan dirinya sendiri jika nama filenya identik
        if clean_cohort_name == clean_uploaded_name:
            continue

        c_words = json.loads(w_json)
        c_map = build_kgram_map(c_words, k=k_val)
        score_up, score_c, passages = calculate_turnitin_similarity(words, c_words, map_uploaded, c_map, k=k_val)
        badge_name, badge_label = get_turnitin_tier(score_up)
        status_val = "PASS" if score_up <= threshold else "FAIL"

        comparison_results.append({
            "Dokumen Pembanding": doc_name,
            "Similaritas Naskah Anda (%)": round(score_up, 2),
            "Similaritas Naskah Pembanding (%)": round(score_c, 2),
            "Status": status_val,
            "Kategori Turnitin": badge_label,
            "Jumlah Blok Teks Cocok": len(passages),
            "Passages": passages
        })

    if not comparison_results:
        return (
            "Dokumen pembanding tidak ditemukan dalam basis data.",
            pd.DataFrame(),
            ""
        )

    progress(0.85, desc="Menyusun kesimpulan analisis dokumen...")
    comparison_results.sort(key=lambda x: x["Similaritas Naskah Anda (%)"], reverse=True)
    max_score = comparison_results[0]["Similaritas Naskah Anda (%)"]
    top_doc = comparison_results[0]["Dokumen Pembanding"]
    badge_name, badge_label = get_turnitin_tier(max_score)
    is_pass = (max_score <= threshold)
    status_label = "✅ LULUS (PASS)" if is_pass else "❌ MELEBIHI BATAS (FAIL)"
    status_color = "#155724" if is_pass else "#721C24"
    bg_color = "#D4EDDA" if is_pass else "#F8D7DA"
    border_color = "#C3E6CB" if is_pass else "#F5C6CB"

    summary_html = f"""<div style="background-color: {bg_color}; border: 1px solid {border_color}; border-radius: 8px; padding: 14px; margin-bottom: 12px; color: {status_color};">
    <div style="font-size: 18px; font-weight: bold; margin-bottom: 6px;">
        Status: {status_label} &nbsp;|&nbsp; Skor Similaritas Tertinggi: {max_score:.2f}% ({badge_label})
    </div>
    <div style="font-size: 13.5px;">
        <b>Nama Dokumen:</b> <code>{uploaded_name}</code> ({len(words):,} kata) &nbsp;|&nbsp; 
        Dibandingkan terhadap: <b>{len(comparison_results)}</b> dokumen cohort &nbsp;|&nbsp; 
        Batas Toleransi: <b>{threshold:.1f}%</b><br/>
        <b>Pasangan Paling Mirip (Top Match):</b> <code>{top_doc}</code> (<b>{max_score:.2f}%</b>)
    </div>
</div>"""

    df_display = pd.DataFrame([{
        "Peringkat": i,
        "Dokumen Pembanding": r["Dokumen Pembanding"],
        "Similaritas (%)": r["Similaritas Naskah Anda (%)"],
        "Status": r["Status"],
        "Kategori Turnitin": r["Kategori Turnitin"],
        "Blok Teks Cocok": r["Jumlah Blok Teks Cocok"]
    } for i, r in enumerate(comparison_results[:20], 1)])

    passages_md = ""
    top_with_passages = [r for r in comparison_results if r["Passages"]][:3]
    if top_with_passages:
        passages_md += "#### 📝 Bukti Cuplikan Teks yang Terdeteksi Mirip (Top 3 Dokumen):\n"
        for r in top_with_passages:
            passages_md += f"\n**🔹 Dokumen Pembanding: `{r['Dokumen Pembanding']}` (Similaritas: {r['Similaritas Naskah Anda (%)']}%)**\n"
            for p_idx, p_text in enumerate(r["Passages"][:3], 1):
                passages_md += f"> *Blok #{p_idx}:* \"...{p_text}...\"\n\n"
    else:
        passages_md = "*(Tidak ditemukan blok teks kembar berturut-turut yang melebihi batas)*"

    progress(1.0, desc="Selesai!")
    return summary_html, df_display, passages_md


# Setup Antarmuka Gradio
def build_gradio_app():
    if gr is None:
        raise ImportError("Gradio belum terpasang. Silakan jalankan: pip install gradio")

    init_summary, init_df, init_excel = load_latest_leaderboard()

    import inspect
    blocks_kwargs = {"title": "Turnitin Document Similarity - P3MD"}
    blocks_params = inspect.signature(gr.Blocks.__init__).parameters
    if "theme" in blocks_params:
        blocks_kwargs["theme"] = gr.themes.Soft()
    if "css" in blocks_params:
        blocks_kwargs["css"] = ".dataframe-table { font-size: 13.5px !important; }"

    with gr.Blocks(**blocks_kwargs) as demo:
        gr.Markdown("""
        # 🔍 Turnitin Document Similarity Checker (P3MD)
        **Sistem Deteksi Similaritas Dokumen Tugas Cohort P3MD Berbasis Standar Turnitin Resmi**
        """)

        # Live Queue & Server Status Monitor (Bebas Antrean - queue=False)
        with gr.Row():
            with gr.Column(scale=5):
                queue_status_display = gr.HTML(value=get_current_queue_status)
            with gr.Column(scale=1):
                refresh_queue_btn = gr.Button("🔄 Cek Antrean", variant="secondary", size="sm")

        refresh_queue_btn.click(
            fn=get_current_queue_status,
            outputs=[queue_status_display],
            queue=False
        )

        with gr.Tabs():
            # ==========================================
            # TAB 1: CEK MANDIRI DOKUMEN (INSTAN ~2 DETIK)
            # ==========================================
            with gr.TabItem("⚡ Cek Mandiri Dokumen (Instan ~2 Detik)", id="tab_single"):
                gr.Markdown("""
                ### ⚡ Pemeriksaan Mandiri & Instan Naskah Tugas
                Unggah draf naskah tugas Anda langsung di sini untuk memeriksa tingkat kemiripan Turnitin terhadap seluruh basis data cohort P3MD dalam hitungan detik, **tanpa perlu menunggu antrean sinkronisasi seluruh angkatan**.
                """)
                with gr.Row():
                    with gr.Column(scale=4):
                        single_file_input = gr.File(
                            label="📄 Unggah File Naskah Tugas (.docx, .pdf, .txt)",
                            file_types=[".docx", ".pdf", ".txt"],
                            type="filepath"
                        )
                    with gr.Column(scale=1):
                        single_run_btn = gr.Button("🔍 Cek Similaritas Naskah", variant="primary", size="lg")

                with gr.Accordion("⚙️ Pengaturan Cek Mandiri (Opsional)", open=False):
                    with gr.Row():
                        s_min_words = gr.Slider(minimum=4, maximum=12, value=6, step=1, label="Min Consecutive Words (Standar Turnitin: 6)")
                        s_thresh = gr.Slider(minimum=5.0, maximum=50.0, value=15.0, step=1.0, label="Batas Toleransi Kelulusan (%)")
                    with gr.Row():
                        s_quotes = gr.Checkbox(value=True, label="Abaikan Kutipan (\" \")")
                        s_bib = gr.Checkbox(value=True, label="Abaikan Daftar Pustaka")

                single_result_html = gr.HTML()
                single_matches_df = gr.Dataframe(
                    label="📋 Dokumen Cohort yang Paling Mirip (Top Matches)",
                    interactive=False,
                    wrap=True
                )
                single_passages_md = gr.Markdown()

                single_run_btn.click(
                    fn=check_single_document,
                    inputs=[single_file_input, s_min_words, s_thresh, s_quotes, s_bib],
                    outputs=[single_result_html, single_matches_df, single_passages_md]
                )

            # ==========================================
            # TAB 2: REKAPITULASI COHORT LENGKAP & EXCEL
            # ==========================================
            with gr.TabItem("📊 Rekapitulasi Cohort P3MD (Leaderboard & Laporan Excel)", id="tab_cohort"):
                with gr.Group():
                    gr.Markdown(f"""
                    ### 📋 Alur Kerja Pengumpulan Dokumen & Sinkronisasi Cohort
                    Ikuti langkah berikut untuk menyinkronkan seluruh tugas cohort:

                    1. **Unggah File Tugas ke Google Drive:**  
                       Klik tombol **"📂 Buka Folder Google Drive P3MD"** di bawah untuk membuka folder pengumpulan tugas. Masukkan naskah tugas Anda (format `.docx`, `.pdf`, atau `.txt`) langsung ke dalam folder tersebut.
                    2. **Jalankan Analisis Similaritas:**  
                       Klik tombol **"🚀 Mulai Sinkronisasi & Analisis Lengkap Cohort"**. Sistem akan otomatis mendeteksi file baru tanpa mengulang unduhan file lama.
                    3. **Pantau Kemajuan (*Progress Bar*):**  
                       Bilah kemajuan menampilkan progres secara realtime mulai dari sinkronisasi Google Drive, kalkulasi pasangan Turnitin, hingga auto-upload cache database.
                    4. **Lihat Hasil & Unduh Laporan Excel:**  
                       Gunakan kolom pencarian di bawah untuk mencari nama peserta secara instan, atau unduh laporan resmi Excel 7-Sheet lengkap.
                    5. **💡 Pembaruan / Revisi Naskah Tugas:**  
                       Jika memperbarui naskah, **timpa file lama dengan nama file yang sama** ATAU **hapus file lama di Google Drive**. Sistem akan otomatis mengabaikan versi lama Anda sendiri!
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
                    run_btn = gr.Button("🚀 Mulai Sinkronisasi & Analisis Lengkap Cohort", variant="primary", size="lg")

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
                
                # Search Bar & Filter Controls (Instant Memory Search - Non-Queued)
                with gr.Row():
                    search_input = gr.Textbox(
                        label="🔍 Cari Dokumen / Nama Peserta (Hasil Instan)",
                        placeholder="Ketik nama file atau peserta untuk menyaring tabel (contoh: Fauzi, Tariq, Rayga, atau status: PASS/FAIL)...",
                        scale=5
                    )
                    reset_search_btn = gr.Button("🔄 Reset Pencarian", scale=1, variant="secondary")

                # Full Page Recap Table (Kompatibel Gradio 4, 5, dan 6)
                df_kwargs = {
                    "value": init_df,
                    "label": "👤 Rekap Hasil Per Peserta (Leaderboard 1 Baris Per Dokumen - Diurutkan dari Skor Tertinggi)",
                    "interactive": False,
                    "wrap": True,
                }
                df_params = inspect.signature(gr.Dataframe.__init__).parameters
                if "max_height" in df_params:
                    df_kwargs["max_height"] = 800
                elif "height" in df_params:
                    df_kwargs["height"] = 800

                table_output = gr.Dataframe(**df_kwargs)
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

                # queue=False memastikan pencarian tabel berjalan seketika di memori tanpa antrean
                search_input.change(fn=filter_table, inputs=[search_input, current_df_state], outputs=[table_output], queue=False)
                reset_search_btn.click(fn=reset_search, inputs=[current_df_state], outputs=[search_input, table_output], queue=False)

                run_btn.click(
                    fn=run_analysis_pipeline,
                    inputs=[drive_input, min_words_slider, thresh_slider, quotes_cb, bib_cb, force_cb],
                    outputs=[status_output, table_output, download_btn, current_df_state, search_input]
                )

        demo.load(fn=get_current_queue_status, outputs=[queue_status_display], queue=False)
        if hasattr(gr, "Timer"):
            try:
                timer = gr.Timer(value=10)
                timer.tick(fn=get_current_queue_status, outputs=[queue_status_display], queue=False)
            except Exception:
                pass

    return demo

if __name__ == "__main__":
    if gr is None:
        print("⚠️ Gradio belum terpasang. Jalankan: pip install gradio")
        sys.exit(1)
    demo = build_gradio_app()
    import inspect
    launch_kwargs = {"share": True, "debug": False}
    launch_params = inspect.signature(demo.launch).parameters
    if "theme" in launch_params and "theme" not in inspect.signature(gr.Blocks.__init__).parameters:
        launch_kwargs["theme"] = gr.themes.Soft()
    if "css" in launch_params and "css" not in inspect.signature(gr.Blocks.__init__).parameters:
        launch_kwargs["css"] = ".dataframe-table { font-size: 13.5px !important; }"

    queue_kwargs = {}
    queue_params = inspect.signature(demo.queue).parameters
    if "default_concurrency_limit" in queue_params:
        queue_kwargs["default_concurrency_limit"] = 1
    elif "concurrency_count" in queue_params:
        queue_kwargs["concurrency_count"] = 1
    if "max_size" in queue_params:
        queue_kwargs["max_size"] = 20

    demo.queue(**queue_kwargs).launch(**launch_kwargs)
