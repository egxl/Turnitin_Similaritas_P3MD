"""
report_generator.py
Generator Laporan Excel Resmi & Komprehensif Turnitin P3MD.
Mengubah perbandingan pasangan (hingga puluhan ribu baris) menjadi sajian data
yang sangat mudah dipahami oleh peserta individu maupun tim penilai/komandan.
"""

import os
import sys
import re
import time
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- PALET WARNA & STYLE STANDAR ---
NAVY_HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
TEAL_HEADER_FILL = PatternFill(start_color="2C6B6F", end_color="2C6B6F", fill_type="solid")
DARK_RED_HEADER_FILL = PatternFill(start_color="842029", end_color="842029", fill_type="solid")
HEADER_FONT = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

CARD_HEADER_FILL = PatternFill(start_color="D9E1E8", end_color="D9E1E8", fill_type="solid")
CARD_HEADER_FONT = Font(name="Segoe UI", size=10, bold=True, color="1F4E79")
CARD_VALUE_FONT = Font(name="Segoe UI", size=18, bold=True, color="1F4E79")
CARD_SUB_FONT = Font(name="Segoe UI", size=9, italic=True, color="555555")

# Status Kelulusan
PASS_FILL = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
PASS_FONT = Font(name="Segoe UI", size=10, bold=True, color="155724")
FAIL_FILL = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
FAIL_FONT = Font(name="Segoe UI", size=10, bold=True, color="721C24")

# Zebra striping
ZEBRA_FILL = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

# Borders
THIN_GRAY = Side(style="thin", color="D3D3D3")
BORDER_BOX = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
BORDER_TOP_BOTTOM = Border(top=THIN_GRAY, bottom=THIN_GRAY)
DOUBLE_BOTTOM = Side(style="double", color="1F4E79")
BORDER_CARD_BOTTOM = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=DOUBLE_BOTTOM)

# Badges Turnitin
BADGE_STYLES = {
    "Blue": (PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid"), Font(name="Segoe UI", size=10, bold=True, color="004085")),
    "Green": (PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid"), Font(name="Segoe UI", size=10, bold=True, color="155724")),
    "Yellow": (PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"), Font(name="Segoe UI", size=10, bold=True, color="856404")),
    "Orange": (PatternFill(start_color="FFE5D0", end_color="FFE5D0", fill_type="solid"), Font(name="Segoe UI", size=10, bold=True, color="A04000")),
    "Red": (PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid"), Font(name="Segoe UI", size=10, bold=True, color="721C24")),
}

def get_turnitin_tier(score):
    if score == 0:
        return "Blue", "🔵 Blue (0%)"
    elif score < 25:
        return "Green", "🟢 Green (1-24%)"
    elif score < 50:
        return "Yellow", "🟡 Yellow (25-49%)"
    elif score < 75:
        return "Orange", "🟠 Orange (50-74%)"
    else:
        return "Red", "🔴 Red (75-100%)"

def style_header_row(ws, row_idx, num_cols, fill=NAVY_HEADER_FILL, font=HEADER_FONT):
    for c in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_BOX
    ws.row_dimensions[row_idx].height = 28

def auto_fit_columns(ws, min_w=12, max_w=65):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or "")
            if "\n" in val_str:
                lines = val_str.split("\n")
                max_len = max(max_len, max(len(l) for l in lines))
            else:
                max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, min_w), max_w)


def compute_leaderboard(df_pairs, threshold=15.0):
    """
    Menghitung rekapitulasi per dokumen (1 baris per dokumen).
    Memproses setiap dokumen terhadap seluruh pasangannya dalam cohort.
    """
    all_docs = set()
    for _, row in df_pairs.iterrows():
        all_docs.add(row["Dokumen 1"])
        all_docs.add(row["Dokumen 2"])

    leaderboard = []
    # Build pairwise lookup map
    doc_matches = {doc: [] for doc in all_docs}
    doc_word_counts = {}

    for _, r in df_pairs.iterrows():
        d1 = r["Dokumen 1"]
        d2 = r["Dokumen 2"]
        max_s = float(r.get("Turnitin Max Score (%)", r.get("Max Score", 0.0)))
        status = r.get("Status Kelulusan", r.get("Status", "PASS"))
        blocks = r.get("Jumlah Blok Teks Cocok", r.get("Matches", 0))

        # Doc 1 stats
        w1 = r.get("Total Kata Doc 1", r.get("Words A", 0))
        if w1 and d1 not in doc_word_counts:
            doc_word_counts[d1] = w1

        # Doc 2 stats
        w2 = r.get("Total Kata Doc 2", r.get("Words B", 0))
        if w2 and d2 not in doc_word_counts:
            doc_word_counts[d2] = w2

        # Record match from d1 perspective
        doc_matches[d1].append({
            "partner": d2,
            "score": max_s,
            "status": status,
            "blocks": blocks
        })
        # Record match from d2 perspective
        doc_matches[d2].append({
            "partner": d1,
            "score": max_s,
            "status": status,
            "blocks": blocks
        })

    for doc, matches in doc_matches.items():
        if not matches:
            continue
        matches_sorted = sorted(matches, key=lambda x: x["score"], reverse=True)
        top_match = matches_sorted[0]
        second_match = matches_sorted[1] if len(matches_sorted) > 1 else None

        worst_score = top_match["score"]
        overall_status = "FAIL" if worst_score > threshold else "PASS"
        fail_partners_count = sum(1 for m in matches if m["score"] > threshold)
        avg_score = sum(m["score"] for m in matches) / len(matches)
        tier_key, tier_label = get_turnitin_tier(worst_score)

        leaderboard.append({
            "Dokumen": doc,
            "Status Kelulusan": overall_status,
            "Skor Tertinggi (%)": worst_score,
            "Kategori Turnitin": tier_label,
            "Tier Key": tier_key,
            "Top Matched Document": top_match["partner"],
            "Top Match Score (%)": top_match["score"],
            "Top Match Blocks": top_match["blocks"],
            "2nd Matched Document": second_match["partner"] if second_match else "-",
            "2nd Match Score (%)": second_match["score"] if second_match else 0.0,
            "Rata-rata Similaritas Cohort (%)": round(avg_score, 2),
            "Jumlah Pasangan > Batas": fail_partners_count,
            "Total Kata": doc_word_counts.get(doc, "-")
        })

    # Urutkan: FAIL lebih dulu, lalu skor tertinggi descending
    leaderboard.sort(key=lambda x: (0 if x["Status Kelulusan"] == "FAIL" else 1, -x["Skor Tertinggi (%)"]))
    return leaderboard


def build_two_way_comparisons(df_pairs):
    """
    Membangun tabel komparasi dua arah (Target vs Pembanding).
    Memudahkan peserta memfilter 1 dokumen target untuk melihat seluruh 397 lawannya.
    """
    two_way = []
    for _, r in df_pairs.iterrows():
        d1 = r["Dokumen 1"]
        d2 = r["Dokumen 2"]
        max_s = float(r.get("Turnitin Max Score (%)", r.get("Max Score", 0.0)))
        status = r.get("Status Kelulusan", r.get("Status", "PASS"))
        badge = r.get("Kategori Turnitin", r.get("Badge", "-"))
        blocks = r.get("Jumlah Blok Teks Cocok", r.get("Matches", 0))

        s1 = r.get("Doc 1 Cocok di Doc 2 (%)", r.get("Score A", max_s))
        s2 = r.get("Doc 2 Cocok di Doc 1 (%)", r.get("Score B", max_s))

        # Baris dari perspektif Dokumen 1
        two_way.append({
            "Dokumen Target": d1,
            "Dokumen Pembanding": d2,
            "Turnitin Max Score (%)": max_s,
            "Status Kelulusan": status,
            "Kategori Turnitin": badge,
            "Kemiripan Target di Pembanding (%)": s1,
            "Kemiripan Pembanding di Target (%)": s2,
            "Jumlah Blok Cocok": blocks
        })
        # Baris dari perspektif Dokumen 2
        two_way.append({
            "Dokumen Target": d2,
            "Dokumen Pembanding": d1,
            "Turnitin Max Score (%)": max_s,
            "Status Kelulusan": status,
            "Kategori Turnitin": badge,
            "Kemiripan Target di Pembanding (%)": s2,
            "Kemiripan Pembanding di Target (%)": s1,
            "Jumlah Blok Cocok": blocks
        })

    two_way.sort(key=lambda x: (x["Dokumen Target"], -x["Turnitin Max Score (%)"]))
    return two_way


def generate_excel_report(df_results, total_docs, min_words=6, commander_threshold=17.0, pass_threshold=15.0, 
                          matched_passages=None, output_file="Turnitin_Similarity_Report_P3MD.xlsx"):
    """
    Fungsi utama pembuatan file Excel 6-Sheet Executive Report.
    """
    wb = openpyxl.Workbook()

    # 1. Hitung Leaderboard Per Peserta
    leaderboard = compute_leaderboard(df_results, threshold=pass_threshold)
    actual_total_docs = len(leaderboard) if leaderboard else total_docs
    total_pairs = len(df_results)
    
    passed_docs = sum(1 for d in leaderboard if d["Status Kelulusan"] == "PASS")
    failed_docs = sum(1 for d in leaderboard if d["Status Kelulusan"] == "FAIL")
    pass_rate = (passed_docs / actual_total_docs * 100) if actual_total_docs > 0 else 0.0

    # Pasangan yang melebihi batas (FAIL)
    score_col = "Turnitin Max Score (%)" if "Turnitin Max Score (%)" in df_results.columns else "Max Score"
    df_flagged = df_results[df_results[score_col].astype(float) > pass_threshold]
    flagged_pairs_count = len(df_flagged)

    # -------------------------------------------------------------------------
    # SHEET 1: 📊 Dashboard Eksekutif
    # -------------------------------------------------------------------------
    ws_dash = wb.active
    ws_dash.title = "📊 Dashboard Eksekutif"
    ws_dash.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_dash.merge_cells("A1:G1")
    title_cell = ws_dash["A1"]
    title_cell.value = "🔍 LAPORAN EKSEKUTIF SIMILARITAS TURNITIN - DOKUMEN P3MD"
    title_cell.fill = NAVY_HEADER_FILL
    title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[1].height = 36

    # Subtitle / Info Bar
    ws_dash.merge_cells("A2:G2")
    sub_cell = ws_dash["A2"]
    sub_cell.value = f"Waktu Pembuatan: {time.strftime('%d %B %Y, %H:%M:%S')}  |  Ambang Batas Kelulusan: {pass_threshold}% (Safety Cushion)  |  Batas Resmi Komandan: {commander_threshold}%  |  Min Kata (k-gram): {min_words}"
    sub_cell.fill = CARD_HEADER_FILL
    sub_cell.font = Font(name="Segoe UI", size=9, bold=True, color="1F4E79")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[2].height = 22

    # KPI Summary Cards (Row 4 to Row 6)
    kpis = [
        ("A", "B", "TOTAL DOKUMEN", f"{actual_total_docs} File", f"Dievaluasi dalam cohort", "1F4E79"),
        ("C", "C", "DOKUMEN LULUS", f"{passed_docs} ({pass_rate:.1f}%)", f"Skor maks <= {pass_threshold}%", "155724"),
        ("D", "D", "MELEBIHI BATAS", f"{failed_docs} File", f"Perlu perbaikan / remedial", "721C24"),
        ("E", "F", "TOTAL PASANGAN DIUJI", f"{total_pairs:,} Pasang", f"Kombinasi n*(n-1)/2", "1F4E79"),
        ("G", "G", "PASANGAN GAGAL", f"{flagged_pairs_count} Pasang", f"Skor > {pass_threshold}%", "721C24")
    ]

    for start_col, end_col, card_title, card_val, card_sub, color_code in kpis:
        c_top = ws_dash[f"{start_col}4"]
        c_val = ws_dash[f"{start_col}5"]
        c_sub = ws_dash[f"{start_col}6"]

        if start_col != end_col:
            ws_dash.merge_cells(f"{start_col}4:{end_col}4")
            ws_dash.merge_cells(f"{start_col}5:{end_col}5")
            ws_dash.merge_cells(f"{start_col}6:{end_col}6")

        c_top.value = card_title
        c_top.fill = CARD_HEADER_FILL
        c_top.font = CARD_HEADER_FONT
        c_top.alignment = Alignment(horizontal="center", vertical="center")
        c_top.border = BORDER_BOX

        c_val.value = card_val
        c_val.font = Font(name="Segoe UI", size=16, bold=True, color=color_code)
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.border = BORDER_BOX

        c_sub.value = card_sub
        c_sub.font = CARD_SUB_FONT
        c_sub.alignment = Alignment(horizontal="center", vertical="center")
        c_sub.border = BORDER_CARD_BOTTOM

    ws_dash.row_dimensions[4].height = 20
    ws_dash.row_dimensions[5].height = 32
    ws_dash.row_dimensions[6].height = 18

    # Tier Breakdown Table (Row 8 to Row 15)
    ws_dash.cell(row=8, column=1, value="📊 DISTRIBUSI TIER WARNA STANDAR TURNITIN COHORT").font = Font(name="Segoe UI", size=11, bold=True, color="1F4E79")
    
    tier_headers = ["Kategori / Tier Turnitin", "Rentang Skor", "Jumlah Dokumen", "Persentase Dokumen", "Status Standar", "Tindakan Disarankan"]
    for idx, h in enumerate(tier_headers, 1):
        cell = ws_dash.cell(row=9, column=idx, value=h)
        cell.fill = NAVY_HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER_BOX
    ws_dash.row_dimensions[9].height = 24

    # Count distribution per tier based on leaderboard max scores
    tier_counts = {"Blue": 0, "Green": 0, "Yellow": 0, "Orange": 0, "Red": 0}
    for item in leaderboard:
        t_key = item["Tier Key"]
        tier_counts[t_key] = tier_counts.get(t_key, 0) + 1

    tier_rows_meta = [
        ("🔵 Blue (0%)", "0%", tier_counts["Blue"], "Aman Sempurna", "Lulus murni tanpa kesamaan verbatim.", "Blue"),
        ("🟢 Green (1-24%)", "1% - 24%", tier_counts["Green"], "Sangat Baik / Wajar", f"Lulus jika <= {pass_threshold}%. Kemiripan wajar format/istilah umum.", "Green"),
        ("🟡 Yellow (25-49%)", "25% - 49%", tier_counts["Yellow"], "Waspada / Mencurigakan", "Tidak Lulus. Terdapat bagian teks panjang yang sama.", "Yellow"),
        ("🟠 Orange (50-74%)", "50% - 74%", tier_counts["Orange"], "Risiko Tinggi", "Tidak Lulus. Mayoritas isi dokumen identik dengan peserta lain.", "Orange"),
        ("🔴 Red (75-100%)", "75% - 100%", tier_counts["Red"], "Plagiasi Berat", "Tidak Lulus. Dokumen duplikasi langsung / copy-paste penuh.", "Red")
    ]

    for r_idx, (t_name, t_range, t_count, t_status, t_action, t_key) in enumerate(tier_rows_meta, 10):
        pct = (t_count / actual_total_docs * 100) if actual_total_docs > 0 else 0.0
        ws_dash.cell(row=r_idx, column=1, value=t_name).font = Font(name="Segoe UI", size=10, bold=True)
        ws_dash.cell(row=r_idx, column=2, value=t_range).alignment = Alignment(horizontal="center")
        ws_dash.cell(row=r_idx, column=3, value=t_count).alignment = Alignment(horizontal="center")
        ws_dash.cell(row=r_idx, column=4, value=f"{pct:.1f}%").alignment = Alignment(horizontal="center")
        
        status_cell = ws_dash.cell(row=r_idx, column=5, value=t_status)
        status_cell.alignment = Alignment(horizontal="center")
        fill_badge, font_badge = BADGE_STYLES[t_key]
        status_cell.fill = fill_badge
        status_cell.font = font_badge

        ws_dash.cell(row=r_idx, column=6, value=t_action).font = Font(name="Segoe UI", size=9)

        for c in range(1, 7):
            ws_dash.cell(row=r_idx, column=c).border = BORDER_BOX
        ws_dash.row_dimensions[r_idx].height = 22

    # Panduan Penggunaan Workbook (Row 17 onwards)
    guide_row = 16
    ws_dash.cell(row=guide_row, column=1, value="📖 PANDUAN MEMBACA & MENGGUNAKAN LAPORAN INI").font = Font(name="Segoe UI", size=11, bold=True, color="1F4E79")
    
    guides = [
        ("1. Bagi Peserta Tugas / Penulis Dokumen:", 
         "Buka sheet '👤 Rekap Per Peserta'. Cukup cari nama Anda (Ctrl+F). Anda akan langsung melihat: Status Lulus/Gagal, Nilai Similaritas Tertinggi Anda, Dokumen Lawan yang paling mirip, dan Rata-rata Similaritas Cohort. Anda TIDAK PERLU memeriksa ribuan baris pasangan."),
        ("2. Memahami Nilai Similaritas:",
         f"Sistem Turnitin membandingkan naskah Anda terhadap seluruh peserta. Jika dokumen Anda memiliki kemiripan 35% dengan Si B, maka skor resmi Anda adalah 35% (status FAIL karena melebihi batas {pass_threshold}%). Rata-rata cohort (~2%) mencerminkan kemiripan template soal."),
        ("3. Bagi Komandan / Tim Penilai:",
         "Buka sheet '⚠️ Investigasi Plagiasi' untuk melihat langsung daftar pasangan yang melanggar batas kelulusan tanpa terganggu ribuan pasangan yang bersih. Buka sheet '📝 Detail Kalimat Identik' untuk memeriksa bukti kutipan kalimat verbatim yang sama."),
        ("4. Analisis Detail Pasangan Tertentu:",
         "Buka sheet '🔍 Cek Dokumen Individu' jika Anda ingin melihat peringkat perbandingan lengkap dari 1 dokumen spesifik terhadap seluruh 397 dokumen lainnya secara terurut.")
    ]

    for g_idx, (g_title, g_desc) in enumerate(guides, guide_row + 1):
        ws_dash.cell(row=g_idx, column=1, value=g_title).font = Font(name="Segoe UI", size=10, bold=True, color="1F4E79")
        ws_dash.merge_cells(f"B{g_idx}:G{g_idx}")
        desc_cell = ws_dash[f"B{g_idx}"]
        desc_cell.value = g_desc
        desc_cell.font = Font(name="Segoe UI", size=9)
        desc_cell.alignment = Alignment(wrap_text=True)
        ws_dash.row_dimensions[g_idx].height = 36

    auto_fit_columns(ws_dash, min_w=14, max_w=70)


    # -------------------------------------------------------------------------
    # SHEET 2: 👤 Rekap Per Peserta (Leaderboard) - 1 Baris Per Dokumen
    # -------------------------------------------------------------------------
    ws_lead = wb.create_sheet("👤 Rekap Per Peserta")
    ws_lead.views.sheetView[0].showGridLines = True
    
    lead_headers = [
        "Rank", "Nama Dokumen (Peserta)", "Status Kelulusan", "Skor Tertinggi (%)", 
        "Kategori Turnitin", "Pasangan Paling Mirip (Top Match)", "Skor Match #1 (%)", 
        "Pasangan Match #2", "Skor Match #2 (%)", "Rata-rata Similaritas Cohort (%)", 
        "Jumlah Pasangan > Batas", "Total Kata"
    ]
    ws_lead.append(lead_headers)
    style_header_row(ws_lead, 1, len(lead_headers), fill=NAVY_HEADER_FILL)

    for rank, item in enumerate(leaderboard, 1):
        r_num = ws_lead.max_row + 1
        ws_lead.append([
            rank,
            item["Dokumen"],
            item["Status Kelulusan"],
            item["Skor Tertinggi (%)"],
            item["Kategori Turnitin"],
            item["Top Matched Document"],
            item["Top Match Score (%)"],
            item["2nd Matched Document"],
            item["2nd Match Score (%)"],
            item["Rata-rata Similaritas Cohort (%)"],
            item["Jumlah Pasangan > Batas"],
            item["Total Kata"]
        ])

        # Style row
        row_cells = [ws_lead.cell(row=r_num, column=c) for c in range(1, len(lead_headers) + 1)]
        
        # Rank & numbers alignment
        row_cells[0].alignment = Alignment(horizontal="center")
        row_cells[2].alignment = Alignment(horizontal="center")
        row_cells[3].alignment = Alignment(horizontal="right")
        row_cells[4].alignment = Alignment(horizontal="center")
        row_cells[6].alignment = Alignment(horizontal="right")
        row_cells[8].alignment = Alignment(horizontal="right")
        row_cells[9].alignment = Alignment(horizontal="right")
        row_cells[10].alignment = Alignment(horizontal="center")
        row_cells[11].alignment = Alignment(horizontal="right")

        # Number formatting
        row_cells[3].number_format = '0.00"%"'
        row_cells[6].number_format = '0.00"%"'
        row_cells[8].number_format = '0.00"%"'
        row_cells[9].number_format = '0.00"%"'
        if isinstance(item["Total Kata"], (int, float)):
            row_cells[11].number_format = '#,##0'

        # Status styling
        if item["Status Kelulusan"] == "PASS":
            row_cells[2].fill = PASS_FILL
            row_cells[2].font = PASS_FONT
        else:
            row_cells[2].fill = FAIL_FILL
            row_cells[2].font = FAIL_FONT

        # Badge styling
        t_key = item["Tier Key"]
        f_badge, fnt_badge = BADGE_STYLES.get(t_key, (None, None))
        if f_badge:
            row_cells[4].fill = f_badge
            row_cells[4].font = fnt_badge

        # Zebra striping for neutral columns
        if rank % 2 == 0 and item["Status Kelulusan"] == "PASS":
            for c_idx in [1, 5, 7]:
                row_cells[c_idx].fill = ZEBRA_FILL

        for cell in row_cells:
            cell.border = BORDER_BOX
        ws_lead.row_dimensions[r_num].height = 20

    ws_lead.freeze_panes = "A2"
    ws_lead.auto_filter.ref = ws_lead.dimensions
    auto_fit_columns(ws_lead, min_w=10, max_w=55)


    # -------------------------------------------------------------------------
    # SHEET 3: ⚠️ Investigasi Plagiasi (Hanya Pasangan Melebihi Batas)
    # -------------------------------------------------------------------------
    ws_flag = wb.create_sheet("⚠️ Investigasi Plagiasi")
    ws_flag.views.sheetView[0].showGridLines = True

    flag_headers = [
        "No.", "Dokumen 1", "Dokumen 2", "Turnitin Max Score (%)", "Status", 
        "Kategori Turnitin", "Doc 1 Cocok di Doc 2 (%)", "Doc 2 Cocok di Doc 1 (%)", 
        "Total Kata Doc 1", "Total Kata Doc 2", "Jumlah Blok Cocok", "Indikasi & Dugaan Hubungan"
    ]
    ws_flag.append(flag_headers)
    style_header_row(ws_flag, 1, len(flag_headers), fill=DARK_RED_HEADER_FILL)

    flag_idx = 0
    for _, r in df_results.iterrows():
        max_s = float(r.get("Turnitin Max Score (%)", r.get("Max Score", 0.0)))
        if max_s <= pass_threshold:
            continue

        flag_idx += 1
        r_num = ws_flag.max_row + 1

        d1 = r["Dokumen 1"]
        d2 = r["Dokumen 2"]
        status = r.get("Status Kelulusan", r.get("Status", "FAIL"))
        badge = r.get("Kategori Turnitin", r.get("Badge", "-"))
        s1 = float(r.get("Doc 1 Cocok di Doc 2 (%)", r.get("Score A", max_s)))
        s2 = float(r.get("Doc 2 Cocok di Doc 1 (%)", r.get("Score B", max_s)))
        w1 = r.get("Total Kata Doc 1", r.get("Words A", 0))
        w2 = r.get("Total Kata Doc 2", r.get("Words B", 0))
        blocks = r.get("Jumlah Blok Teks Cocok", r.get("Matches", 0))

        # Analisis rasio & arah dugaan
        if abs(s1 - s2) > 10 and w1 and w2:
            if s1 > s2:
                indikasi = f"Dokumen 1 ({w1:,} kata) memuat porsi besar teks dari Dokumen 2 ({w2:,} kata). Indikasi Dokumen 1 menyalin sebagian Dokumen 2."
            else:
                indikasi = f"Dokumen 2 ({w2:,} kata) memuat porsi besar teks dari Dokumen 1 ({w1:,} kata). Indikasi Dokumen 2 menyalin sebagian Dokumen 1."
        else:
            indikasi = "Kemiripan simetris dua arah dalam skala signifikan. Indikasi pengerjaan bersama atau sumber bahan yang sama."

        ws_flag.append([
            flag_idx, d1, d2, max_s, status, badge, s1, s2, w1, w2, blocks, indikasi
        ])

        # Styling
        row_cells = [ws_flag.cell(row=r_num, column=c) for c in range(1, len(flag_headers) + 1)]
        row_cells[0].alignment = Alignment(horizontal="center")
        row_cells[3].alignment = Alignment(horizontal="right")
        row_cells[3].number_format = '0.00"%"'
        row_cells[4].alignment = Alignment(horizontal="center")
        row_cells[4].fill = FAIL_FILL
        row_cells[4].font = FAIL_FONT
        row_cells[5].alignment = Alignment(horizontal="center")
        row_cells[6].alignment = Alignment(horizontal="right")
        row_cells[6].number_format = '0.00"%"'
        row_cells[7].alignment = Alignment(horizontal="right")
        row_cells[7].number_format = '0.00"%"'
        row_cells[8].alignment = Alignment(horizontal="right")
        row_cells[9].alignment = Alignment(horizontal="right")
        row_cells[10].alignment = Alignment(horizontal="center")
        row_cells[11].alignment = Alignment(vertical="center", wrap_text=True)

        for cell in row_cells:
            cell.border = BORDER_BOX
        ws_flag.row_dimensions[r_num].height = 30

    if flag_idx == 0:
        ws_flag.append(["-", "Selamat! Tidak ada pasangan dokumen yang melebihi batas kelulusan.", "", "", "ALL PASS", "", "", "", "", "", "", ""])
        ws_flag.cell(row=2, column=2).font = Font(name="Segoe UI", size=11, bold=True, color="155724")

    ws_flag.freeze_panes = "A2"
    ws_flag.auto_filter.ref = ws_flag.dimensions
    auto_fit_columns(ws_flag, min_w=10, max_w=65)


    # -------------------------------------------------------------------------
    # SHEET 4: 🔍 Cek Dokumen Individu (Two-Way Inspector View)
    # -------------------------------------------------------------------------
    ws_inspect = wb.create_sheet("🔍 Cek Dokumen Individu")
    ws_inspect.views.sheetView[0].showGridLines = True

    inspect_headers = [
        "Dokumen Target (Filter Nama Anda Di Sini)", "Dokumen Pembanding", 
        "Turnitin Max Score (%)", "Status Kelulusan", "Kategori Turnitin", 
        "Kemiripan Target di Pembanding (%)", "Kemiripan Pembanding di Target (%)", "Jumlah Blok Cocok"
    ]
    ws_inspect.append(inspect_headers)
    style_header_row(ws_inspect, 1, len(inspect_headers), fill=TEAL_HEADER_FILL)

    two_way_data = build_two_way_comparisons(df_results)
    for r_data in two_way_data:
        r_num = ws_inspect.max_row + 1
        ws_inspect.append([
            r_data["Dokumen Target"],
            r_data["Dokumen Pembanding"],
            r_data["Turnitin Max Score (%)"],
            r_data["Status Kelulusan"],
            r_data["Kategori Turnitin"],
            r_data["Kemiripan Target di Pembanding (%)"],
            r_data["Kemiripan Pembanding di Target (%)"],
            r_data["Jumlah Blok Cocok"]
        ])

        row_cells = [ws_inspect.cell(row=r_num, column=c) for c in range(1, len(inspect_headers) + 1)]
        row_cells[2].alignment = Alignment(horizontal="right")
        row_cells[2].number_format = '0.00"%"'
        row_cells[3].alignment = Alignment(horizontal="center")
        row_cells[4].alignment = Alignment(horizontal="center")
        row_cells[5].alignment = Alignment(horizontal="right")
        row_cells[5].number_format = '0.00"%"'
        row_cells[6].alignment = Alignment(horizontal="right")
        row_cells[6].number_format = '0.00"%"'
        row_cells[7].alignment = Alignment(horizontal="center")

        if r_data["Status Kelulusan"] == "FAIL":
            row_cells[3].fill = FAIL_FILL
            row_cells[3].font = FAIL_FONT
        else:
            row_cells[3].fill = PASS_FILL
            row_cells[3].font = PASS_FONT

        for cell in row_cells:
            cell.border = BORDER_BOX
        ws_inspect.row_dimensions[r_num].height = 19

    ws_inspect.freeze_panes = "A2"
    ws_inspect.auto_filter.ref = ws_inspect.dimensions
    auto_fit_columns(ws_inspect, min_w=12, max_w=55)


    # -------------------------------------------------------------------------
    # SHEET 5: 📝 Detail Kalimat Identik
    # -------------------------------------------------------------------------
    ws_pass = wb.create_sheet("📝 Detail Kalimat Identik")
    ws_pass.views.sheetView[0].showGridLines = True

    pass_headers = [
        "Dokumen 1", "Dokumen 2", "Status Pasangan", "Skor Max (%)", 
        "Perkiraan Jumlah Kata Cocok", "Potongan Kalimat Identik (Verbatim)"
    ]
    ws_pass.append(pass_headers)
    style_header_row(ws_pass, 1, len(pass_headers), fill=NAVY_HEADER_FILL)

    # Populate matched passages with priority given to FAIL pairs
    if matched_passages:
        # Sort matched passages so FAIL pairs come first
        passages_sorted = sorted(
            matched_passages, 
            key=lambda x: (0 if x.get("status") == "FAIL" else 1, -float(str(x.get("score", 0)).replace("%", "")))
        )
        for p_item in passages_sorted:
            r_num = ws_pass.max_row + 1
            text_val = p_item.get("text", "")
            word_cnt = len(text_val.split())
            score_num = float(str(p_item.get("score", 0)).replace("%", ""))

            ws_pass.append([
                p_item.get("doc1", "-"),
                p_item.get("doc2", "-"),
                p_item.get("status", "PASS"),
                score_num,
                word_cnt,
                text_val
            ])

            row_cells = [ws_pass.cell(row=r_num, column=c) for c in range(1, len(pass_headers) + 1)]
            row_cells[2].alignment = Alignment(horizontal="center")
            row_cells[3].alignment = Alignment(horizontal="right")
            row_cells[3].number_format = '0.00"%"'
            row_cells[4].alignment = Alignment(horizontal="center")
            row_cells[5].alignment = Alignment(vertical="center", wrap_text=True)

            if p_item.get("status") == "FAIL":
                row_cells[2].fill = FAIL_FILL
                row_cells[2].font = FAIL_FONT
            else:
                row_cells[2].fill = PASS_FILL
                row_cells[2].font = PASS_FONT

            for cell in row_cells:
                cell.border = BORDER_BOX
            ws_pass.row_dimensions[r_num].height = 24

    ws_pass.freeze_panes = "A2"
    ws_pass.auto_filter.ref = ws_pass.dimensions
    auto_fit_columns(ws_pass, min_w=12, max_w=75)


    # -------------------------------------------------------------------------
    # SHEET 6: 📋 Semua Pasangan (Arsip)
    # -------------------------------------------------------------------------
    ws_all = wb.create_sheet("📋 Semua Pasangan (Arsip)")
    ws_all.views.sheetView[0].showGridLines = True

    all_headers = [
        "No.", "Dokumen 1", "Dokumen 2", "Turnitin Max Score (%)", "Status Kelulusan", 
        "Kategori Turnitin", "Doc 1 Cocok di Doc 2 (%)", "Doc 2 Cocok di Doc 1 (%)", 
        "Total Kata Doc 1", "Total Kata Doc 2", "Jumlah Blok Teks Cocok"
    ]
    ws_all.append(all_headers)
    style_header_row(ws_all, 1, len(all_headers), fill=NAVY_HEADER_FILL)

    for idx, (_, r) in enumerate(df_results.iterrows(), 1):
        r_num = ws_all.max_row + 1
        max_s = float(r.get("Turnitin Max Score (%)", r.get("Max Score", 0.0)))
        status = r.get("Status Kelulusan", r.get("Status", "PASS"))
        badge = r.get("Kategori Turnitin", r.get("Badge", "-"))
        s1 = float(r.get("Doc 1 Cocok di Doc 2 (%)", r.get("Score A", max_s)))
        s2 = float(r.get("Doc 2 Cocok di Doc 1 (%)", r.get("Score B", max_s)))
        w1 = r.get("Total Kata Doc 1", r.get("Words A", "-"))
        w2 = r.get("Total Kata Doc 2", r.get("Words B", "-"))
        blocks = r.get("Jumlah Blok Teks Cocok", r.get("Matches", 0))

        ws_all.append([
            idx, r["Dokumen 1"], r["Dokumen 2"], max_s, status, badge, s1, s2, w1, w2, blocks
        ])

        row_cells = [ws_all.cell(row=r_num, column=c) for c in range(1, len(all_headers) + 1)]
        row_cells[0].alignment = Alignment(horizontal="center")
        row_cells[3].alignment = Alignment(horizontal="right")
        row_cells[3].number_format = '0.00"%"'
        row_cells[4].alignment = Alignment(horizontal="center")
        row_cells[5].alignment = Alignment(horizontal="center")
        row_cells[6].alignment = Alignment(horizontal="right")
        row_cells[6].number_format = '0.00"%"'
        row_cells[7].alignment = Alignment(horizontal="right")
        row_cells[7].number_format = '0.00"%"'
        row_cells[8].alignment = Alignment(horizontal="right")
        row_cells[9].alignment = Alignment(horizontal="right")
        row_cells[10].alignment = Alignment(horizontal="center")

        if status == "FAIL":
            row_cells[4].fill = FAIL_FILL
            row_cells[4].font = FAIL_FONT
        else:
            row_cells[4].fill = PASS_FILL
            row_cells[4].font = PASS_FONT

        if idx % 2 == 0 and status == "PASS":
            for c_i in [1, 2]:
                row_cells[c_i].fill = ZEBRA_FILL

        for cell in row_cells:
            cell.border = BORDER_BOX
        ws_all.row_dimensions[r_num].height = 19

    ws_all.freeze_panes = "A2"
    ws_all.auto_filter.ref = ws_all.dimensions
    auto_fit_columns(ws_all, min_w=10, max_w=50)


    # -------------------------------------------------------------------------
    # SHEET 7: ⚙️ Parameter Analisis
    # -------------------------------------------------------------------------
    ws_meta = wb.create_sheet("⚙️ Parameter Analisis")
    ws_meta.views.sheetView[0].showGridLines = True

    meta_headers = ["Parameter Sistem", "Nilai Konfigurasi", "Keterangan Standar Turnitin"]
    ws_meta.append(meta_headers)
    style_header_row(ws_meta, 1, len(meta_headers), fill=NAVY_HEADER_FILL)

    metadata_rows = [
        ("Ambang Batas Resmi Komandan (Official Threshold)", f"{commander_threshold}%", "Batas maksimal toleransi similaritas yang diizinkan komandan."),
        ("Ambang Batas Pengujian (Safety Cushion Threshold)", f"{pass_threshold}%", "Batas aman sistem (cushioning) untuk mendeteksi potensi pelanggaran."),
        ("Minimum Consecutive Words (k-gram Shingling)", f"{min_words} kata berurutan", "Standar resmi Turnitin untuk mendeteksi kesamaan frasa verbatim."),
        ("Eksklusi Kutipan (Filter Quotes)", "True (Aktif)", "Mengabaikan teks di dalam tanda petik ganda (\"...\" / “...”)."),
        ("Eksklusi Daftar Pustaka (Filter Bibliography)", "True (Aktif)", "Mengabaikan bab Daftar Pustaka, References, dan Rujukan."),
        ("Total Dokumen Dianalisis", f"{actual_total_docs} dokumen", "Jumlah total dokumen naskah yang berhasil diekstraksi."),
        ("Total Pasangan Dianalisis", f"{total_pairs:,} pasangan", "Jumlah seluruh kombinasi perbandingan berpasangan n*(n-1)/2."),
        ("Dokumen Memenuhi Syarat (PASS)", f"{passed_docs} dokumen ({pass_rate:.1f}%)", "Dokumen dengan skor tertinggi <= batas aman kelulusan."),
        ("Dokumen Melebihi Batas (FAIL)", f"{failed_docs} dokumen ({100 - pass_rate:.1f}%)", "Dokumen yang memerlukan revisi atau pemeriksaan komprehensif."),
        ("Waktu Pembuatan Laporan", time.strftime("%Y-%m-%d %H:%M:%S"), "Waktu komputasi dan penulisan file laporan Excel selesai.")
    ]

    for m_param, m_val, m_desc in metadata_rows:
        r_num = ws_meta.max_row + 1
        ws_meta.append([m_param, m_val, m_desc])
        row_cells = [ws_meta.cell(row=r_num, column=c) for c in range(1, len(meta_headers) + 1)]
        row_cells[0].font = Font(name="Segoe UI", size=10, bold=True, color="1F4E79")
        row_cells[1].font = Font(name="Segoe UI", size=10, bold=True)
        row_cells[1].alignment = Alignment(horizontal="center")
        row_cells[2].font = Font(name="Segoe UI", size=9)
        for cell in row_cells:
            cell.border = BORDER_BOX
        ws_meta.row_dimensions[r_num].height = 22

    ws_meta.freeze_panes = "A2"
    auto_fit_columns(ws_meta, min_w=15, max_w=65)

    # Simpan file
    wb.save(output_file)
    return output_file


if __name__ == "__main__":
    import pandas as pd

    # Script testing / direct upgrade on existing file
    excel_input = "Turnitin_Similarity_Report_P3MD.xlsx"
    if os.path.exists(excel_input):
        print(f"📖 Membaca file contoh: {excel_input}...")
        wb_in = openpyxl.load_workbook(excel_input, data_only=True)
        ws_summary = wb_in["Similarity Summary"]
        data = list(ws_summary.iter_rows(values_only=True))
        df_in = pd.DataFrame(data[1:], columns=data[0])

        passages_in = []
        if "Matched Passages Detail" in wb_in.sheetnames:
            ws_p = wb_in["Matched Passages Detail"]
            p_data = list(ws_p.iter_rows(values_only=True))
            if len(p_data) > 1:
                p_headers = p_data[0]
                for r in p_data[1:]:
                    pair_str = str(r[0] or "")
                    score_str = str(r[1] or "0%")
                    passage_text = str(r[2] or "")
                    parts = re.split(r"\s*<-->\s*|\s+vs\s+", pair_str)
                    d1 = parts[0].strip() if len(parts) > 0 else "-"
                    d2 = parts[1].strip() if len(parts) > 1 else "-"
                    
                    try:
                        s_val = float(score_str.replace("%", "").strip())
                    except:
                        s_val = 0.0

                    status = "FAIL" if s_val > 15.0 else "PASS"
                    passages_in.append({
                        "doc1": d1,
                        "doc2": d2,
                        "score": s_val,
                        "status": status,
                        "text": passage_text
                    })

        print(f"⚙️ Memproses ulang {len(df_in)} baris pasangan & {len(passages_in)} kutipan kalimat...")
        out = generate_excel_report(
            df_results=df_in,
            total_docs=len(set(df_in["Dokumen 1"]).union(set(df_in["Dokumen 2"]))),
            min_words=6,
            commander_threshold=17.0,
            pass_threshold=15.0,
            matched_passages=passages_in,
            output_file="Turnitin_Similarity_Report_P3MD.xlsx"
        )
        print(f"✅ Selesai! File laporan baru tersimpan di: {out}")
