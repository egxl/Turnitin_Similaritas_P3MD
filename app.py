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
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None

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
            if fitz is not None:
                try:
                    doc = fitz.open(file_path)
                    text = "\n".join([page.get_text() or "" for page in doc if page.get_text().strip()])
                except Exception:
                    text = ""
            if not text and PdfReader is not None:
                reader = PdfReader(file_path)
                text = "\n".join([page.extract_text() or "" for page in reader.pages])
            elif not text and fitz is None and PdfReader is None:
                raise ImportError("Library 'pypdf' atau 'pymupdf' belum terinstal. Silakan jalankan: pip install pypdf pymupdf")
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

class CohortMemoryCache:
    """
    Cache memori global (in-memory) untuk menyimpan tokens dan inverted index k-gram dokumen cohort.
    Mengeliminasi lag disk SQLite dan JSON deserialization pada setiap pemanggilan Cek Mandiri.
    Pencarian kemiripan berjalan dalam < 30ms langsung di RAM.
    """
    def __init__(self):
        self.db_path = None
        self.db_mtime = None
        self.doc_words = {}            # filename -> list[str] (tokens)
        self.kgram_sets = {}           # (k_val, filename) -> set[str]
        self.inverted_index = {}       # k_val -> dict[str, list[(filename, int)]]
        self.lock = threading.Lock()

    def invalidate(self):
        with self.lock:
            self.doc_words.clear()
            self.kgram_sets.clear()
            self.inverted_index.clear()
            self.db_mtime = None

    def ensure_loaded(self, db_path, k_val=6, target_dir="./dokumen_tugas_p3md", drop_quotes=True, drop_bib=True):
        with self.lock:
            current_mtime = os.path.getmtime(db_path) if os.path.exists(db_path) else 0.0

            # Jika database belum termuat atau berkas SQLite diperbarui:
            if (self.db_path != db_path or self.db_mtime != current_mtime or not self.doc_words):
                self.doc_words.clear()
                self.kgram_sets.clear()
                self.inverted_index.clear()
                self.db_path = db_path
                self.db_mtime = current_mtime

                # 1. Coba baca dari SQLite
                if os.path.exists(db_path):
                    conn = sqlite3.connect(db_path, timeout=15)
                    try:
                        cur = conn.cursor()
                        cur.execute("SELECT filename, words_json FROM documents")
                        for fname, w_json in cur.fetchall():
                            if w_json and w_json != "[]":
                                try:
                                    w_list = json.loads(w_json)
                                    if w_list:
                                        self.doc_words[fname] = w_list
                                except Exception:
                                    pass
                    finally:
                        conn.close()

                # 2. Jika words_json masih kosong di database, coba ekstrak otomatis dari berkas lokal
                if not self.doc_words and os.path.exists(target_dir):
                    supported_exts = {".docx", ".pdf", ".txt"}
                    local_files = []
                    for root, _, f_list in os.walk(target_dir):
                        for f in f_list:
                            if os.path.splitext(f)[1].lower() in supported_exts and not f.startswith("~"):
                                local_files.append(os.path.join(root, f))

                    if local_files:
                        db_conn = sqlite3.connect(db_path, timeout=15) if os.path.exists(db_path) else None
                        try:
                            for fp in local_files:
                                bname = os.path.basename(fp)
                                raw = extract_raw_text(fp)
                                if raw.strip():
                                    filt = apply_turnitin_exclusions(raw, drop_quotes=drop_quotes, drop_bib=drop_bib)
                                    toks = tokenize_words(filt)
                                    if len(toks) > 0:
                                        self.doc_words[bname] = toks
                                        if db_conn:
                                            mtime = os.path.getmtime(fp)
                                            sz = os.path.getsize(fp)
                                            db_conn.execute("""
                                                INSERT OR REPLACE INTO documents (filename, file_mtime, file_size, word_count, words_json)
                                                VALUES (?, ?, ?, ?, ?)
                                            """, (bname, mtime, sz, len(toks), json.dumps(toks)))
                            if db_conn:
                                db_conn.commit()
                                self.db_mtime = os.path.getmtime(db_path)
                        finally:
                            if db_conn:
                                db_conn.close()

            # 3. Bangun Inverted Index & K-gram sets untuk k_val tertentu di RAM jika belum ada
            if k_val not in self.inverted_index and self.doc_words:
                inv_idx = {}
                for fname, words in self.doc_words.items():
                    kset = set()
                    for i in range(len(words) - k_val + 1):
                        gram = " ".join(words[i:i+k_val])
                        kset.add(gram)
                        if gram not in inv_idx:
                            inv_idx[gram] = []
                        inv_idx[gram].append((fname, i))
                    self.kgram_sets[(k_val, fname)] = kset
                self.inverted_index[k_val] = inv_idx

GLOBAL_COHORT_CACHE = CohortMemoryCache()

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
            if new_or_updated > 0:
                GLOBAL_COHORT_CACHE.invalidate()
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
            if pruned_count > 0:
                GLOBAL_COHORT_CACHE.invalidate()
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
    Memprioritaskan data langsung dari database cache SQLite agar selalu mutakhir dan instan.
    Jika database belum tersedia, menggunakan cadangan file Excel atau mengunduh baseline.
    """
    db_candidates = [
        os.path.join(target_dir, "similarity_cache.db"),
        "similarity_cache.db"
    ]
    excel_candidates = [
        os.path.join(target_dir, "Turnitin_Similarity_Report_P3MD.xlsx"),
        "Turnitin_Similarity_Report_P3MD.xlsx"
    ]

    excel_path = None
    for p in excel_candidates:
        if os.path.exists(p):
            excel_path = p
            break

    # 1. Prioritaskan pembacaan langsung dari database cache SQLite (Sangat cepat ~15ms & anti-stuck)
    for db_p in db_candidates:
        if os.path.exists(db_p):
            try:
                conn = sqlite3.connect(db_p)
                cur = conn.cursor()
                cur.execute("SELECT filename FROM documents")
                docs = [r[0] for r in cur.fetchall()]
                conn.close()
                if len(docs) >= 2:
                    cache = TurnitinDBCache(db_path=db_p)
                    df_results = cache.get_results_dataframe(docs, k_val=6)
                    if len(df_results) > 0:
                        leaderboard = compute_leaderboard(df_results, threshold=PASS_THRESHOLD)
                        for i, item in enumerate(leaderboard, 1):
                            item["Rank"] = i
                            item["Nama Dokumen (Peserta)"] = item["Dokumen"]
                            item["Pasangan Paling Mirip (Top Match)"] = item["Top Matched Document"]
                            item["Skor Match #1 (%)"] = item["Top Match Score (%)"]
                            item["Pasangan Match #2"] = item["2nd Matched Document"]
                            item["Skor Match #2 (%)"] = item["2nd Match Score (%)"]

                        df = pd.DataFrame(leaderboard)[[
                            "Rank", "Nama Dokumen (Peserta)", "Status Kelulusan", "Skor Tertinggi (%)", 
                            "Kategori Turnitin", "Pasangan Paling Mirip (Top Match)", "Skor Match #1 (%)", 
                            "Pasangan Match #2", "Skor Match #2 (%)", "Rata-rata Similaritas Cohort (%)", 
                            "Jumlah Pasangan > Batas", "Total Kata"
                        ]]
                        failed = sum(1 for d in leaderboard if d["Status Kelulusan"] == "FAIL")
                        passed = len(leaderboard) - failed
                        total_pairs = len(df_results)
                        summary_md = f"""### 📊 Ringkasan Eksekutif Similaritas Cohort P3MD (Data Database Terkini)
- **Total Dokumen Peserta:** {len(leaderboard)} file
- **Kelulusan Cohort:** ✅ **{passed} LULUS** ({passed/len(leaderboard)*100:.1f}%) | ❌ **{failed} MELEBIHI BATAS** ({failed/len(leaderboard)*100:.1f}%)
- **Total Pasangan Diuji:** {total_pairs:,} pasang
- ℹ️ *Data di bawah disajikan langsung dari basis data cache SQLite terbaru. Klik tombol **"🚀 Mulai Sinkronisasi & Analisis Lengkap Cohort"** untuk menyinkronkan tugas baru.*
"""
                        return summary_md, df, excel_path
            except Exception as e:
                print(f"⚠️ Info: Gagal memuat dari database cache: {e}, mencoba cadangan Excel...")

    # 2. Cadangan: Muat dari file Excel jika ada
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
            summary_md = f"""### 📊 Ringkasan Eksekutif Similaritas Cohort P3MD (Data Cadangan Excel)
- **Total Dokumen Peserta:** {len(df)} file
- **Kelulusan Cohort:** ✅ **{passed} LULUS** ({passed/len(df)*100:.1f}%) | ❌ **{failed} MELEBIHI BATAS** ({failed/len(df)*100:.1f}%)
- **Total Pasangan Diuji:** {total_pairs:,} pasang
- ℹ️ *Data di bawah adalah hasil analisis tersimpan dari laporan Excel.*
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
Belum ada data analisis tersimpan. Silakan unggah file tugas ke Google Drive lalu klik tombol **"🚀 Mulai Sinkronisasi & Analisis Lengkap Cohort"** di atas.
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
        status_html = f"""<div style="padding: 12px 16px; border-radius: 4px; background: #161920; border: 1px solid #F59E0B; color: #F3F4F6; margin-bottom: 12px; font-family: -apple-system, BlinkMacSystemFont, 'Geist', sans-serif;">
    <div style="font-size: 13.5px; font-weight: 700; display: flex; align-items: center; justify-content: space-between; font-family: 'Geist Mono', monospace;">
        <span style="color: #FCD34D;">⏳ <b>SERVER // SEDANG MEMPROSES ANALISIS</b></span>
        <span style="font-size: 11px; background: rgba(245, 158, 11, 0.15); border: 1px solid rgba(245, 158, 11, 0.4); padding: 2px 8px; border-radius: 2px; color: #FCD34D;">{current_queued} TUGAS AKTIF</span>
    </div>
    <div style="font-size: 12px; margin-top: 6px; color: #94A3B8; font-family: 'Geist Mono', monospace; line-height: 1.6;">
        • <b>Proses:</b> 1 analisis sedang aktif ({elapsed}s berjalan) &nbsp;|&nbsp; <b>Menunggu di Antrean:</b> {waiting} tugas<br/>
        • <b>Database Cache:</b> {doc_count} dokumen tersimpan ({pair_count:,} pasangan teranalisis)
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

        status_html = f"""<div style="padding: 12px 16px; border-radius: 4px; background: #111317; border: 1px solid #222631; color: #F3F4F6; margin-bottom: 12px; font-family: -apple-system, BlinkMacSystemFont, 'Geist', sans-serif;">
    <div style="font-size: 13.5px; font-weight: 700; display: flex; align-items: center; justify-content: space-between; font-family: 'Geist Mono', monospace;">
        <span style="color: #4ADE80; display: flex; align-items: center; gap: 8px;">
            <span style="width: 8px; height: 8px; border-radius: 50%; background: #22C55E; box-shadow: 0 0 8px #22C55E; display: inline-block;"></span>
            <b>SERVER // SIAP & BEBAS ANTREAN (IDLE)</b>
        </span>
        <span style="font-size: 11px; background: rgba(34, 197, 94, 0.1); border: 1px solid rgba(34, 197, 94, 0.3); padding: 2px 8px; border-radius: 2px; color: #4ADE80;">0 ANTREAN MENUNGGU</span>
    </div>
    <div style="font-size: 12px; margin-top: 6px; color: #94A3B8; font-family: 'Geist Mono', monospace; line-height: 1.6;">
        • <b>Status Antrean:</b> Bebas. Analisis mandiri (~2s) &amp; batch cohort dapat langsung dieksekusi.<br/>
        • <b>Sinkronisasi Terakhir:</b> {last_updated_str} &nbsp;|&nbsp; <b>Basis Data:</b> {doc_count} dokumen ({pair_count:,} pasangan teranalisis)
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
        
        # 2. Pindai Dokumen Lokal & Siapkan Database Cache
        supported_exts = {".docx", ".pdf", ".txt"}
        file_paths = []
        for root, _, f_list in os.walk(target_dir):
            for f in f_list:
                if os.path.splitext(f)[1].lower() in supported_exts and not f.startswith("~"):
                    file_paths.append(os.path.join(root, f))

        cache = TurnitinDBCache(db_path=db_path)

        if len(file_paths) < 2:
            # Periksa apakah database cache SQLite sudah memiliki dokumen tersimpan
            conn = sqlite3.connect(db_path)
            try:
                cur = conn.cursor()
                cur.execute("SELECT filename, words_json FROM documents")
                cached_rows = cur.fetchall()
            except Exception:
                cached_rows = []
            finally:
                conn.close()

            if len(cached_rows) >= 2:
                log(f"ℹ️ Menggunakan {len(cached_rows)} dokumen dari basis data cache SQLite.")
                doc_db = {r[0]: json.loads(r[1]) for r in cached_rows}
                loaded_c = len(cached_rows)
                new_c = 0
            else:
                return (
                    f"⚠️ Ditemukan {len(file_paths)} dokumen di folder tugas dan belum ada cache tersimpan. Minimal diperlukan 2 dokumen untuk analisis.\n\n"
                    f"Silakan unggah dokumen ke folder Google Drive terlebih dahulu:\n{active_drive_url}",
                    pd.DataFrame(),
                    None,
                    pd.DataFrame(),
                    ""
                )
        else:
            # 3. Sinkronisasi SQLite Dokumen & Ekstraksi Teks (0.25 -> 0.50)
            progress(0.25, desc="Mempersiapkan database cache SQLite...")
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

        # 5. Rekapitulasi Data & Hitung Leaderboard untuk Tampilan Web SEGERA (0.85 -> 0.90)
        progress(0.85, desc="Menyusun data rekapitulasi...")
        df_results = cache.get_results_dataframe(list(doc_db.keys()), k_val=int(min_words))

        progress(0.88, desc="Menghasilkan Leaderboard per peserta...")
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

        failed_docs_count = sum(1 for d in leaderboard if d["Status Kelulusan"] == "FAIL")
        passed_docs_count = len(leaderboard) - failed_docs_count
        fail_pairs_count = len(df_results[df_results["Turnitin Max Score (%)"] > float(pass_thresh)])

        # 6. Pembuatan Laporan Excel (Non-Blocking & Aman terhadap Kunci Berkas)
        skip_excel = (new_c == 0 and new_pairs == 0 and not force_recompute and os.path.exists(excel_path))
        excel_out_path = excel_path if os.path.exists(excel_path) else None

        if not skip_excel:
            progress(0.90, desc="Menyusun workbook Excel...")
            try:
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
                            if len(all_passages) >= 500:
                                break
                    if len(all_passages) >= 500:
                        break

                excel_out_path = generate_excel_report(
                    df_results, len(doc_db), min_words=int(min_words), 
                    commander_threshold=COMMANDER_THRESHOLD, pass_threshold=float(pass_thresh),
                    matched_passages=all_passages, output_file=excel_path
                )
                log("✅ Laporan Excel resmi berhasil diperbarui.")
            except Exception as e_xl:
                log(f"⚠️ Info Excel: {e_xl}. Tabel web tetap disajikan mutakhir!")
        else:
            log("⚡ Tidak ada dokumen/pasangan baru. Menggunakan laporan Excel yang sudah ada.")
            progress(0.92, desc="Menggunakan laporan Excel yang sudah ada...")

        # 7. Otomatis Unggah Cache Database dan Laporan Excel ke Google Drive (0.95 -> 0.99)
        folder_id = extract_folder_id(active_drive_url)
        drive_upload_success = False
        if folder_id and not skip_excel:
            progress(0.96, desc="Sinkronisasi cache ke Google Drive...")
            try:
                up_db = upload_file_to_drive(db_path, folder_id, log_callback=log)
                up_xl = upload_file_to_drive(excel_out_path, folder_id, log_callback=log) if excel_out_path else False
                drive_upload_success = up_db or up_xl
            except Exception as e_drv:
                log(f"⚠️ Info upload Drive: {e_drv}")
        elif skip_excel:
            drive_upload_success = True

        sync_note = "☁️ **Cache database & Excel tersinkron ke Google Drive.**" if drive_upload_success else "💾 **Cache tersimpan secara lokal.**"
        recalc_note = "⚡ **Hemat Waktu:** Tidak ada dokumen baru, pembuatan Excel dilewati." if skip_excel else f"⏱️ **Waktu Hemat:** ~{round((cached_pairs * 0.005) / 60, 1)} menit berkat SQLite Cache!"

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
        return summary_md, df_display, excel_out_path, df_display, ""
    finally:
        with jobs_lock:
            active_jobs_count = max(0, active_jobs_count - 1)

def check_single_document(file_obj, min_words=6, pass_thresh=PASS_THRESHOLD, drop_quotes=True, drop_bib=True, progress=None):
    """
    Memeriksa similaritas 1 dokumen yang diunggah langsung ke Gradio terhadap seluruh
    dokumen cohort menggunakan in-memory inverted index (selesai dalam hitungan milidetik).
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

    progress(0.3, desc="Menghubungkan ke cache memori cohort...")
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

    # Pastikan data cohort termuat di RAM
    GLOBAL_COHORT_CACHE.ensure_loaded(db_path, k_val=k_val, target_dir=target_dir, drop_quotes=drop_quotes, drop_bib=drop_bib)

    if not GLOBAL_COHORT_CACHE.doc_words:
        return (
            "⚠️ Basis data cache cohort belum memiliki teks dokumen tersimpan. Silakan klik tombol 'Mulai Sinkronisasi & Analisis Lengkap Cohort' pada tab Rekapitulasi untuk memproses dokumen.",
            pd.DataFrame(),
            ""
        )

    progress(0.6, desc="Kalkulasi similaritas Turnitin via Inverted Index...")
    map_uploaded = build_kgram_map(words, k=k_val)
    inv_index = GLOBAL_COHORT_CACHE.inverted_index.get(k_val, {})

    # Pembersihan nama dokumen yang diunggah untuk deteksi kesamaan nama
    clean_uploaded_name = re.sub(r"^[a-f0-9]{16,}_", "", uploaded_name.lower())
    clean_up_base = os.path.splitext(clean_uploaded_name)[0]

    # Inisialisasi pelacakan indeks kecocokan per dokumen cohort
    matched_indices_by_doc = {fname: set() for fname in GLOBAL_COHORT_CACHE.doc_words.keys()}

    # 1-Pass Inverted Index Lookup: O(N_uploaded)
    for gram, start_indices in map_uploaded.items():
        if gram in inv_index:
            for fname, c_offset in inv_index[gram]:
                clean_cohort_name = re.sub(r"^[a-f0-9]{16,}_", "", fname.lower())
                clean_c_base = os.path.splitext(clean_cohort_name)[0]

                # Lewati perbandingan jika file identik persis
                if clean_cohort_name == clean_uploaded_name or clean_c_base == clean_up_base:
                    continue

                for s_idx in start_indices:
                    for offset in range(k_val):
                        matched_indices_by_doc[fname].add(s_idx + offset)

    comparison_results = []
    self_matches = []

    for fname, matched_set in matched_indices_by_doc.items():
        clean_cohort_name = re.sub(r"^[a-f0-9]{16,}_", "", fname.lower())
        clean_c_base = os.path.splitext(clean_cohort_name)[0]
        if clean_cohort_name == clean_uploaded_name or clean_c_base == clean_up_base:
            continue

        score_up = (len(matched_set) / len(words) * 100) if words else 0.0

        # Deteksi Smart Self-Match:
        # Jika kemiripan >= 85% dan nama file memiliki kemiripan kata kunci (misal revisi atau nama peserta yang sama)
        is_potential_self = False
        if score_up >= 85.0:
            name_words_up = set(re.findall(r"\w+", clean_up_base))
            name_words_c = set(re.findall(r"\w+", clean_c_base))
            common_stopwords = {"ujian", "tahap", "p3md", "tugas", "danbatch", "jawaban", "lembar", "komprehensif", "uk1", "g1", "g2", "g3", "g4", "e1", "e2", "e3", "pdf", "docx", "txt", "test", "tes"}
            name_words_up -= common_stopwords
            name_words_c -= common_stopwords
            if name_words_up and name_words_c and (name_words_up & name_words_c):
                is_potential_self = True

        badge_name, badge_label = get_turnitin_tier(score_up)
        status_val = "PASS" if score_up <= threshold else "FAIL"

        item = {
            "Dokumen Pembanding": fname,
            "Similaritas Naskah Anda (%)": round(score_up, 2),
            "Status": status_val,
            "Kategori Turnitin": badge_label,
            "Jumlah Blok Teks Cocok": 0,
            "MatchedSet": matched_set,
            "IsSelfMatch": is_potential_self
        }

        if is_potential_self:
            self_matches.append(item)
        else:
            comparison_results.append(item)

    if not comparison_results and not self_matches:
        return (
            "Dokumen pembanding tidak ditemukan dalam basis data.",
            pd.DataFrame(),
            ""
        )

    # Urutkan berdasarkan skor tertinggi naskah Anda
    comparison_results.sort(key=lambda x: x["Similaritas Naskah Anda (%)"], reverse=True)

    # Hitung Cumulative Union Turnitin Score (Resmi Standar Turnitin)
    valid_matched_indices = set()
    for r in comparison_results:
        valid_matched_indices.update(r["MatchedSet"])

    cumulative_score = (len(valid_matched_indices) / len(words) * 100) if words else 0.0
    cum_badge_name, cum_badge_label = get_turnitin_tier(cumulative_score)

    top_doc = comparison_results[0]["Dokumen Pembanding"] if comparison_results else "-"
    top_single_score = comparison_results[0]["Similaritas Naskah Anda (%)"] if comparison_results else 0.0
    top_badge_name, top_badge_label = get_turnitin_tier(top_single_score)

    # Standar Penilaian Resmi & Pencegahan False Positive (Selaras dengan Rekapitulasi Cohort Excel):
    # Kelulusan ditentukan oleh sumber tunggal terbesar (Top Match). Jika naskah tidak menyalin > batas dari
    # salah satu rekan, maka dokumen dinyatakan LULUS. Ini mencegah false positive akibat akumulasi template soal/UU
    # yang tersebar di ratusan dokumen peserta.
    is_pass = (top_single_score <= threshold)

    status_label = "✅ LULUS (PASS)" if is_pass else "❌ MELEBIHI BATAS (FAIL)"
    status_color = "#155724" if is_pass else "#721C24"
    bg_color = "#D4EDDA" if is_pass else "#F8D7DA"
    border_color = "#C3E6CB" if is_pass else "#F5C6CB"

    # Lazy Passage Construction: Hanya untuk TOP 3 dokumen!
    for r in comparison_results[:3]:
        m_set = r["MatchedSet"]
        if m_set:
            sorted_indices = sorted(m_set)
            passages = []
            curr_p = [words[sorted_indices[0]]]
            for prev_idx, curr_idx in zip(sorted_indices[:-1], sorted_indices[1:]):
                if curr_idx == prev_idx + 1:
                    curr_p.append(words[curr_idx])
                else:
                    if len(curr_p) >= k_val:
                        passages.append(" ".join(curr_p))
                    curr_p = [words[curr_idx]]
            if len(curr_p) >= k_val:
                passages.append(" ".join(curr_p))
            r["Passages"] = passages
            r["Jumlah Blok Teks Cocok"] = len(passages)
        else:
            r["Passages"] = []
            r["Jumlah Blok Teks Cocok"] = 0

    self_match_notice = ""
    if self_matches:
        s_names = ", ".join([f"<code style='color:#38BDF8;'>{sm['Dokumen Pembanding']}</code> ({sm['Similaritas Naskah Anda (%)']}%)" for sm in self_matches])
        self_match_notice = f"""<div style="margin-top: 12px; padding: 10px 14px; background: #161920; border: 1px solid #343B4D; border-radius: 4px; font-size: 12.5px; color: #94A3B8; font-family: 'Geist Mono', monospace;">
        ℹ️ <b>Draf / Revisi Sebelumnya Terdeteksi:</b> {s_names}<br/>
        <i>Sistem otomatis memfilter draf lama Anda agar tidak dianggap sebagai plagiasi terhadap diri sendiri. Skor perbandingan murni terhadap naskah rekan cohort lainnya.</i>
        </div>"""

    # Kotak Transparansi Metodologi Penilaian (Anti False-Positive)
    if is_pass:
        if cumulative_score > threshold:
            transparency_box = f"""<div style="margin-top: 14px; padding: 12px 16px; background: rgba(56, 189, 248, 0.06); border-left: 3px solid #38BDF8; border-radius: 0 4px 4px 0; font-size: 13px; color: #BAE6FD; line-height: 1.6; font-family: -apple-system, BlinkMacSystemFont, 'Geist', sans-serif;">
                💡 <b>Transparansi Metodologi Resmi (Pencegahan False Positive):</b><br/>
                • <b>Penentu Kelulusan (Standar Laporan Excel):</b> Status diukur dari <b>Sumber Tunggal Terbesar (Top Match)</b> terhadap satu rekan (<b>{top_single_score:.2f}%</b> &le; {threshold:.1f}%). Dokumen Anda dinyatakan <b>LULUS</b> karena tidak terindikasi menyalin naskah rekan tertentu.<br/>
                • <b>Tentang Skor Kumulatif ({cumulative_score:.2f}%):</b> Angka ini adalah total gabungan kemiripan terhadap seluruh {len(comparison_results)} dokumen cohort. Dalam ujian bersama, skor kumulatif wajar terakumulasi dari template soal ujian, rujukan UU/peraturan desa, dan terminologi baku yang digunakan banyak peserta, bukan plagiasi individu.
            </div>"""
        else:
            transparency_box = f"""<div style="margin-top: 14px; padding: 12px 16px; background: rgba(34, 197, 94, 0.06); border-left: 3px solid #22C55E; border-radius: 0 4px 4px 0; font-size: 13px; color: #BBF7D0; line-height: 1.6; font-family: -apple-system, BlinkMacSystemFont, 'Geist', sans-serif;">
                💡 <b>Transparansi Metodologi:</b> Baik kemiripan sumber tunggal terbesar (<b>{top_single_score:.2f}%</b>) maupun skor kumulatif cohort (<b>{cumulative_score:.2f}%</b>) berada di bawah ambang batas toleransi {threshold:.1f}%. Dokumen sepenuhnya bersih dari indikasi plagiasi.
            </div>"""
    else:
        transparency_box = f"""<div style="margin-top: 14px; padding: 12px 16px; background: rgba(239, 68, 68, 0.08); border-left: 3px solid #EF4444; border-radius: 0 4px 4px 0; font-size: 13px; color: #FECACA; line-height: 1.6; font-family: -apple-system, BlinkMacSystemFont, 'Geist', sans-serif;">
            ⚠️ <b>Kemiripan Melebihi Batas Toleransi:</b><br/>
            Terdeteksi kemiripan tinggi sebesar <b>{top_single_score:.2f}%</b> terhadap dokumen rekan <code>{top_doc}</code> (batas toleransi: {threshold:.1f}%). Silakan periksa cuplikan teks identik pada bagian <i>Bukti Cuplikan Teks</i> di bawah untuk direvisi atau diparafrase.
        </div>"""

    status_badge_html = f"""<span style="font-size: 13px; font-weight: 700; font-family: 'Geist Mono', monospace; background: {'rgba(34, 197, 94, 0.12)' if is_pass else 'rgba(239, 68, 68, 0.15)'}; color: {'#4ADE80' if is_pass else '#F87171'}; border: 1px solid {'rgba(34, 197, 94, 0.3)' if is_pass else 'rgba(239, 68, 68, 0.35)'}; padding: 4px 12px; border-radius: 2px;">
        {status_label}
    </span>"""

    summary_html = f"""<div style="background-color: #111317; border: 1px solid #222631; border-radius: 4px; padding: 20px 22px; margin-bottom: 16px; color: #F3F4F6; font-family: -apple-system, BlinkMacSystemFont, 'Geist', sans-serif;">
    <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; flex-wrap: wrap; gap: 12px; border-bottom: 1px solid #222631; padding-bottom: 14px;">
        <div>
            <div style="font-size: 19px; font-weight: 700; letter-spacing: -0.02em; color: #FFFFFF;">
                {uploaded_name}
            </div>
            <div style="font-size: 12px; color: #94A3B8; font-family: 'Geist Mono', monospace; margin-top: 4px;">
                {len(words):,} KATA &nbsp;|&nbsp; DIBANDINGKAN TERHADAP {len(comparison_results)} DOKUMEN COHORT &nbsp;|&nbsp; TOLERANSI: {threshold:.1f}%
            </div>
        </div>
        <div>
            {status_badge_html}
        </div>
    </div>
    
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-bottom: 16px;">
        <div style="background: #090A0C; border: 1px solid #222631; padding: 12px 16px; border-radius: 2px;">
            <div style="font-size: 11px; font-family: 'Geist Mono', monospace; color: #64748B; text-transform: uppercase;">Top Match (Sumber Tunggal)</div>
            <div style="font-size: 24px; font-weight: 700; font-family: 'Geist Mono', monospace; color: {'#4ADE80' if is_pass else '#FCD34D'}; margin: 4px 0 2px;">
                {top_single_score:.2f}%
            </div>
            <div style="font-size: 11.5px; color: #94A3B8; font-family: 'Geist Mono', monospace;">
                {top_badge_label} &bull; vs <code>{top_doc}</code>
            </div>
        </div>
        <div style="background: #090A0C; border: 1px solid #222631; padding: 12px 16px; border-radius: 2px;">
            <div style="font-size: 11px; font-family: 'Geist Mono', monospace; color: #64748B; text-transform: uppercase;">Kumulatif Seluruh Cohort</div>
            <div style="font-size: 24px; font-weight: 700; font-family: 'Geist Mono', monospace; color: #38BDF8; margin: 4px 0 2px;">
                {cumulative_score:.2f}%
            </div>
            <div style="font-size: 11.5px; color: #94A3B8; font-family: 'Geist Mono', monospace;">
                {cum_badge_label} &bull; Total Gabungan Cohort
            </div>
        </div>
    </div>
    {transparency_box}
    {self_match_notice}
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
    top_with_passages = [r for r in comparison_results if r.get("Passages")][:3]
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


SWISS_MONOCHROME_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Geist+Mono:wght@400;500;600;700&family=Geist:wght@300;400;500;600;700&display=swap');

:root, body, .gradio-container, .gradio-container.light, .gradio-container.dark, gradio-app {
    --bg-dark: #090A0C !important;
    --surface-dark: #111317 !important;
    --surface-card: #161920 !important;
    --border-dark: #222631 !important;
    --border-strong: #343B4D !important;
    --text-primary: #F3F4F6 !important;
    --text-muted: #94A3B8 !important;
    --text-dim: #64748B !important;
    --accent-cyan: #38BDF8 !important;
    --body-background-fill: #090A0C !important;
    --background-fill-primary: #111317 !important;
    --background-fill-secondary: #161920 !important;
    --block-background-fill: #111317 !important;
    --block-border-color: #222631 !important;
    --border-color-primary: #222631 !important;
    --body-text-color: #F3F4F6 !important;
    --block-label-text-color: #94A3B8 !important;
    --input-background-fill: #090A0C !important;
    --input-border-color: #222631 !important;
    --input-text-color: #F3F4F6 !important;
    --table-odd-background-fill: #111317 !important;
    --table-even-background-fill: #161920 !important;
    --panel-background-fill: #111317 !important;
    --panel-border-color: #222631 !important;
    background-color: #090A0C !important;
    color: #F3F4F6 !important;
    font-family: 'Geist', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
}

body, html, .gradio-container, .gradio-container.light, gradio-app {
    background: #090A0C !important;
    color: #F3F4F6 !important;
}

.gradio-container .block,
.gradio-container .panel,
.gradio-container .form,
.gradio-container fieldset {
    background: #111317 !important;
    border: 1px solid #222631 !important;
    border-radius: 4px !important;
    color: #F3F4F6 !important;
}

.swiss-masthead {
    background: #111317;
    border: 1px solid #222631;
    border-radius: 4px;
    padding: 18px 22px;
    margin-bottom: 16px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 16px;
}

.swiss-brand-title {
    font-size: 17px;
    font-weight: 700;
    letter-spacing: -0.02em;
    color: #F3F4F6;
    display: flex;
    align-items: center;
    gap: 12px;
}

.swiss-brand-sub {
    font-size: 12px;
    color: #94A3B8;
    margin-top: 3px;
    font-family: 'Geist Mono', monospace;
}

.swiss-specs {
    display: flex;
    align-items: center;
    gap: 14px;
    font-family: 'Geist Mono', monospace;
    font-size: 11px;
    color: #64748B;
}

.swiss-chip {
    background: #161920;
    border: 1px solid #222631;
    padding: 3px 8px;
    border-radius: 2px;
    color: #38BDF8;
}

.tabs > .tab-nav {
    border-bottom: 1px solid #222631 !important;
    background: transparent !important;
    gap: 4px !important;
}

.tabs > .tab-nav > button {
    font-family: 'Geist Mono', monospace !important;
    font-size: 12.5px !important;
    font-weight: 600 !important;
    color: #64748B !important;
    border: 1px solid transparent !important;
    border-bottom: 2px solid transparent !important;
    background: transparent !important;
    border-radius: 4px 4px 0 0 !important;
    padding: 10px 18px !important;
    transition: all 0.15s ease !important;
}

.tabs > .tab-nav > button:hover {
    color: #F3F4F6 !important;
}

.tabs > .tab-nav > button.selected {
    color: #F3F4F6 !important;
    background: #111317 !important;
    border-color: #222631 !important;
    border-bottom: 2px solid #38BDF8 !important;
}

button.primary, .btn-primary {
    background: #F3F4F6 !important;
    color: #090A0C !important;
    font-weight: 700 !important;
    font-size: 13.5px !important;
    border: none !important;
    border-radius: 4px !important;
    font-family: 'Geist', sans-serif !important;
    letter-spacing: -0.01em !important;
    padding: 10px 18px !important;
    transition: all 0.15s ease !important;
    cursor: pointer !important;
}

button.primary:hover, .btn-primary:hover {
    background: #FFFFFF !important;
    box-shadow: 0 0 16px rgba(255, 255, 255, 0.2) !important;
    transform: translateY(-1px) !important;
}

button.secondary, .btn-secondary {
    background: #161920 !important;
    color: #F3F4F6 !important;
    border: 1px solid #343B4D !important;
    font-family: 'Geist Mono', monospace !important;
    font-size: 12px !important;
    border-radius: 4px !important;
    transition: all 0.15s ease !important;
    cursor: pointer !important;
}

button.secondary:hover, .btn-secondary:hover {
    border-color: #38BDF8 !important;
    color: #38BDF8 !important;
}

.dataframe-table, table, .table-wrap, [data-testid="dataframe"] {
    background-color: #111317 !important;
    color: #F3F4F6 !important;
    border-collapse: collapse !important;
    font-size: 13px !important;
    border: 1px solid #222631 !important;
}

th, .dataframe-table th {
    background-color: #161920 !important;
    color: #94A3B8 !important;
    font-family: 'Geist Mono', monospace !important;
    font-size: 11px !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.06em !important;
    border-bottom: 1px solid #222631 !important;
    padding: 12px 16px !important;
}

td, .dataframe-table td {
    border-bottom: 1px solid #222631 !important;
    padding: 11px 16px !important;
    color: #CBD5E1 !important;
}

tr:hover td, .dataframe-table tr:hover td {
    background-color: rgba(255, 255, 255, 0.025) !important;
    color: #F3F4F6 !important;
}

input[type="text"], textarea, .textbox input {
    background: #090A0C !important;
    color: #F3F4F6 !important;
    border: 1px solid #222631 !important;
    border-radius: 4px !important;
    font-family: 'Geist', sans-serif !important;
    font-size: 13px !important;
}

input[type="text"]:focus, textarea:focus {
    border-color: #38BDF8 !important;
    box-shadow: 0 0 0 1px #38BDF8 !important;
}

.file-upload, [data-testid="file-upload"], .dropzone {
    background: #090A0C !important;
    border: 1px dashed #343B4D !important;
    border-radius: 4px !important;
    transition: all 0.15s ease !important;
}

.file-upload:hover, [data-testid="file-upload"]:hover {
    border-color: #38BDF8 !important;
    background: rgba(56, 189, 248, 0.02) !important;
}

.accordion {
    border: 1px solid #222631 !important;
    border-radius: 4px !important;
    background: #111317 !important;
}

.prose blockquote {
    background: #090A0C !important;
    border-left: 3px solid #EAB308 !important;
    color: #CBD5E1 !important;
    padding: 10px 16px !important;
    border-radius: 0 4px 4px 0 !important;
    font-family: 'Geist Mono', monospace !important;
    font-size: 12.5px !important;
    line-height: 1.7 !important;
}
"""

# Setup Antarmuka Gradio
def build_gradio_app():
    if gr is None:
        raise ImportError("Gradio belum terpasang. Silakan jalankan: pip install gradio")

    init_summary, init_df, init_excel = load_latest_leaderboard()

    import inspect
    blocks_kwargs = {"title": "Turnitin Document Similarity - P3MD"}
    blocks_params = inspect.signature(gr.Blocks.__init__).parameters
    if "theme" in blocks_params:
        try:
            blocks_kwargs["theme"] = gr.themes.Monochrome(
                primary_hue="neutral",
                secondary_hue="slate",
                neutral_hue="zinc",
                font=[gr.themes.GoogleFont("Geist"), "sans-serif"],
                font_mono=[gr.themes.GoogleFont("Geist Mono"), "monospace"]
            )
        except Exception:
            blocks_kwargs["theme"] = gr.themes.Base()
    if "css" in blocks_params:
        blocks_kwargs["css"] = SWISS_MONOCHROME_CSS
    if "js" in blocks_params:
        blocks_kwargs["js"] = "() => { document.documentElement.classList.add('dark'); document.body.classList.add('dark'); }"

    with gr.Blocks(**blocks_kwargs) as demo:
        gr.HTML(f"""
        <style>
        {SWISS_MONOCHROME_CSS}
        </style>
        <script>
        document.documentElement.classList.add('dark');
        document.body.classList.add('dark');
        </script>
        <div class="swiss-masthead">
            <div class="swiss-brand">
                <div class="swiss-brand-title">
                    <span style="display:inline-flex; align-items:center; justify-content:center; width:28px; height:28px; background:#161920; border:1px solid #343B4D; border-radius:2px;">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="#38BDF8"><path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"/></svg>
                    </span>
                    <span>TURNITIN // P3MD COHORT SIMILARITY ENGINE</span>
                </div>
                <div class="swiss-brand-sub">Sistem Deteksi Similaritas Dokumen Tugas Cohort P3MD Berbasis Standar Turnitin Resmi</div>
            </div>
            <div class="swiss-specs">
                <span>SQLITE RAM CACHE: AKTIF (~15ms)</span>
                <span>STANDAR RESMI: VERBATIM 6-GRAM</span>
                <span class="swiss-chip">🔵-🔴 5 TIERS</span>
            </div>
        </div>
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

                > 💡 **Standar Kelulusan Bebas False Positive:** Selaras dengan Laporan Resmi Excel, status kelulusan dinilai dari **Sumber Tunggal Terbesar (Top Match)**. Hal ini memastikan naskah Anda tidak dinyatakan gagal secara keliru (*false positive*) hanya karena akumulasi template soal, sitasi UU/regulasi desa, atau format baku yang tersebar di ratusan peserta cohort.
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
                        s_thresh = gr.Slider(minimum=5.0, maximum=50.0, value=15.0, step=1.0, label="Batas Toleransi Sumber Tunggal (%) (Standar Excel: 15.0%)")
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
                single_click_kwargs = {}
                btn_click_params = inspect.signature(single_run_btn.click).parameters
                if "concurrency_limit" in btn_click_params:
                    single_click_kwargs["concurrency_limit"] = 10
                elif "concurrency_id" in btn_click_params:
                    single_click_kwargs["concurrency_id"] = "instant_check"

                single_run_btn.click(
                    fn=check_single_document,
                    inputs=[single_file_input, s_min_words, s_thresh, s_quotes, s_bib],
                    outputs=[single_result_html, single_matches_df, single_passages_md],
                    **single_click_kwargs
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

                run_click_kwargs = {}
                if "concurrency_limit" in btn_click_params:
                    run_click_kwargs["concurrency_limit"] = 1
                elif "concurrency_id" in btn_click_params:
                    run_click_kwargs["concurrency_id"] = "cohort_pipeline"

                run_btn.click(
                    fn=run_analysis_pipeline,
                    inputs=[drive_input, min_words_slider, thresh_slider, quotes_cb, bib_cb, force_cb],
                    outputs=[status_output, table_output, download_btn, current_df_state, search_input],
                    **run_click_kwargs
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
        try:
            launch_kwargs["theme"] = gr.themes.Monochrome()
        except Exception:
            pass
    if "css" in launch_params and "css" not in inspect.signature(gr.Blocks.__init__).parameters:
        launch_kwargs["css"] = SWISS_MONOCHROME_CSS

    queue_kwargs = {}
    queue_params = inspect.signature(demo.queue).parameters
    if "default_concurrency_limit" in queue_params:
        queue_kwargs["default_concurrency_limit"] = 10
    elif "concurrency_count" in queue_params:
        queue_kwargs["concurrency_count"] = 10
    if "max_size" in queue_params:
        queue_kwargs["max_size"] = 50

    demo.queue(**queue_kwargs).launch(**launch_kwargs)
