#!/usr/bin/env python3
"""PDF generator untuk laporan audit Buku Kurikulum.

Fokus laporan:
- sederhana dan rapi
- status memakai PATUH / KURANG PATUH / TIDAK PATUH berbasis skor
- tabel per bagian utama / parent
- temuan tambahan masuk ke bagian terkait
- bagian opsional yang tidak ditemukan tidak ditampilkan
- tidak memakai hardcode nama bagian dokumen
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


STATUS_PATUH = "PATUH"
STATUS_KURANG_PATUH = "KURANG PATUH"
STATUS_TIDAK_PATUH = "TIDAK PATUH"

TELKOM_RED = "#E30613"
TELKOM_NAVY = "#0A2C70"
TELKOM_LIGHT = "#F3F6FB"
TELKOM_BORDER = "#D6DEEE"
TELKOM_TEXT = "#1F2937"
TELKOM_MUTED = "#6B7280"
TELKOM_GREEN = "#16823A"
TELKOM_ORANGE = "#D58A00"
TELKOM_DANGER = "#B00020"
TELKOM_SOFT_RED = "#FFF0F2"
TELKOM_SOFT_GREEN = "#F0FFF4"
TELKOM_SOFT_ORANGE = "#FFF8E8"

STATUS_COLORS = {
    STATUS_PATUH: TELKOM_GREEN,
    STATUS_KURANG_PATUH: TELKOM_ORANGE,
    STATUS_TIDAK_PATUH: TELKOM_DANGER,
    "ADA": TELKOM_GREEN,
    "TIDAK ADA": TELKOM_DANGER,
    "TAMBAHAN": TELKOM_NAVY,
    "OPSIONAL": TELKOM_MUTED,
}


# =========================================================
# Basic helpers
# =========================================================


#INI FUNGSI/BAGIAN UNTUK (merapikan teks: hilangkan newline dan spasi berlebih)
def clean_text(value: Any) -> str:
    if value is None:
        return "-"
    text = str(value).replace("\r", " ").replace("\n", " ")
    text = " ".join(text.split()).strip()
    return text or "-"


#INI FUNGSI/BAGIAN UNTUK (mengubah nilai apa pun menjadi teks aman untuk PDF)
def safe_text(value: Any) -> str:
    return escape(clean_text(value))


#INI FUNGSI/BAGIAN UNTUK (menormalkan teks (huruf kecil, rapikan spasi))
def normalize(value: Any) -> str:
    return clean_text(value).casefold()


#INI FUNGSI/BAGIAN UNTUK (memilih warna sesuai status)
def color_for(status: str) -> str:
    return STATUS_COLORS.get(clean_text(status).upper(), TELKOM_NAVY)


#INI FUNGSI/BAGIAN UNTUK (mengambil nama file sumber dari hasil)
def source_file_name(result: Dict[str, Any]) -> str:
    file_path = clean_text(result.get("file"))
    if not file_path or file_path == "-":
        return "Dokumen"
    name = Path(str(file_path)).name
    name = re.sub(r"^\d{10,}[-_\s]+", "", name).strip()
    return name or "Dokumen"


#INI FUNGSI/BAGIAN UNTUK (menampilkan nama file saja tanpa folder dan tanpa angka acak)
def display_name(value: Any) -> str:
    # Ambil nama file saja (tanpa folder) dan buang prefix angka acak/timestamp agar tampil apa adanya.
    text = clean_text(value)
    if not text or text == "-":
        return "-"
    name = Path(str(text)).name
    name = re.sub(r"^\d{10,}[-_\s]+", "", name).strip()
    return name or "-"


#INI FUNGSI/BAGIAN UNTUK (mengambil teks/keterangan halaman dari sebuah item)
def page_text(item: Dict[str, Any]) -> str:
    page = clean_text(item.get("page") or item.get("page_ref") or item.get("toc_page"))
    if page == "-":
        return ""
    return f"hal. {safe_text(page)}"


#INI FUNGSI/BAGIAN UNTUK (menggabungkan nomor dan judul bagian)
def title_with_number(item: Dict[str, Any], title_key: str = "field", number_key: str = "matched_number") -> str:
    title = clean_text(item.get(title_key))
    number = clean_text(item.get(number_key))
    if number != "-" and title != "-":
        return f"{number} {title}"
    return title


#INI FUNGSI/BAGIAN UNTUK (mengambil nama dasar file input untuk nama laporan)
def report_input_stem(result: Dict[str, Any]) -> str:
    file_path = clean_text(result.get("file"))
    if file_path == "-":
        return "dokumen"
    stem = Path(file_path).stem
    return stem or "dokumen"


#INI FUNGSI/BAGIAN UNTUK (menentukan path output PDF laporan)
def resolve_report_output_path(output_path: str, result: Dict[str, Any], default_names: Tuple[str, ...]) -> str:
    output = Path(output_path)
    output_dir = output.parent
    output_name = output.name

    if output.suffix.lower() != ".pdf":
        output_dir = output
        output_name = ""

    default_set = {name.casefold() for name in default_names}
    if not output_name or output_name.casefold() in default_set:
        output_name = f"laporan_{report_input_stem(result)}.pdf"

    return str(output_dir / output_name)


# =========================================================
# Styles
# =========================================================


#INI FUNGSI/BAGIAN UNTUK (menyiapkan kumpulan gaya teks/paragraf untuk PDF)
def build_styles() -> Dict[str, ParagraphStyle]:
    styles = getSampleStyleSheet()
    return {
        "base": styles,
        "cover_title": ParagraphStyle(
            "cover_title",
            parent=styles["Title"],
            fontSize=22,
            leading=27,
            alignment=1,
            textColor=colors.HexColor(TELKOM_RED),
            spaceAfter=6,
        ),
        "cover_subtitle": ParagraphStyle(
            "cover_subtitle",
            parent=styles["Heading2"],
            fontSize=14,
            leading=18,
            alignment=1,
            textColor=colors.HexColor(TELKOM_NAVY),
            spaceAfter=8,
        ),
        "heading": ParagraphStyle(
            "heading",
            parent=styles["Heading2"],
            fontSize=13,
            leading=16,
            alignment=1,
            textColor=colors.HexColor(TELKOM_NAVY),
            spaceAfter=6,
        ),
        "section_heading": ParagraphStyle(
            "section_heading",
            parent=styles["Heading2"],
            fontSize=14,
            leading=17,
            textColor=colors.HexColor(TELKOM_NAVY),
            spaceBefore=4,
            spaceAfter=8,
        ),
        "body": ParagraphStyle(
            "body",
            parent=styles["BodyText"],
            fontSize=9.2,
            leading=12,
            textColor=colors.HexColor(TELKOM_TEXT),
        ),
        "cell": ParagraphStyle(
            "cell",
            parent=styles["BodyText"],
            fontSize=8.4,
            leading=10.7,
            textColor=colors.HexColor(TELKOM_TEXT),
            splitLongWords=False,
        ),
        "cell_small": ParagraphStyle(
            "cell_small",
            parent=styles["BodyText"],
            fontSize=8.0,
            leading=10.2,
            textColor=colors.HexColor(TELKOM_TEXT),
            splitLongWords=False,
        ),
        "cell_center": ParagraphStyle(
            "cell_center",
            parent=styles["BodyText"],
            fontSize=8.3,
            leading=10.5,
            alignment=1,
            textColor=colors.HexColor(TELKOM_TEXT),
            splitLongWords=False,
        ),
        "muted": ParagraphStyle(
            "muted",
            parent=styles["BodyText"],
            fontSize=8.7,
            leading=11.4,
            textColor=colors.HexColor(TELKOM_MUTED),
        ),
        "small_center": ParagraphStyle(
            "small_center",
            parent=styles["Normal"],
            fontSize=9,
            leading=12,
            alignment=1,
            textColor=colors.grey,
        ),
        "box_title": ParagraphStyle(
            "box_title",
            parent=styles["Heading3"],
            fontSize=10.2,
            leading=13,
            textColor=colors.HexColor(TELKOM_NAVY),
            spaceAfter=3,
        ),
        "subtitle": ParagraphStyle(
            "subtitle",
            parent=styles["BodyText"],
            fontSize=10,
            leading=13,
            textColor=colors.HexColor(TELKOM_NAVY),
            alignment=1,
        ),
    }


# =========================================================
# Status and scoring
# =========================================================


#INI FUNGSI/BAGIAN UNTUK (menentukan status satu aturan/bagian)
def rule_status(item: Dict[str, Any]) -> str:
    """Status sederhana untuk satu rule.

    - Ada -> PATUH
    - Tidak ada -> TIDAK PATUH

    Catatan: selama bagian ditemukan di dokumen, bagian dianggap PATUH.
    Perbedaan posisi parent/induk tidak lagi menurunkan status menjadi kurang patuh.
    """
    if item.get("found"):
        return STATUS_PATUH
    return STATUS_TIDAK_PATUH


#INI FUNGSI/BAGIAN UNTUK (menghitung skor satu aturan/bagian)
def rule_score(item: Dict[str, Any]) -> float:
    status = rule_status(item)
    if status == STATUS_PATUH:
        return 1.0
    if status == STATUS_KURANG_PATUH:
        return 0.5
    return 0.0


#INI FUNGSI/BAGIAN UNTUK (menyaring aturan yang perlu ditampilkan)
def visible_rule_rows(result: Dict[str, Any]) -> List[Dict[str, Any]]:
    visible: List[Dict[str, Any]] = []
    for item in result.get("validation_results", []):
        # Bagian opsional yang tidak ditemukan tidak ditampilkan.
        if not item.get("required") and not item.get("found"):
            continue
        visible.append(item)
    return visible


#INI FUNGSI/BAGIAN UNTUK (menghitung ringkasan jumlah bagian per status)
def count_summary(result: Dict[str, Any]) -> Dict[str, int]:
    rows = result.get("validation_results", [])
    extras = result.get("extra_sections", [])

    required_rows = [item for item in rows if item.get("required")]
    optional_rows = [item for item in rows if not item.get("required")]

    required_ok = sum(1 for item in required_rows if rule_status(item) == STATUS_PATUH)
    required_partial = sum(1 for item in required_rows if rule_status(item) == STATUS_KURANG_PATUH)
    required_missing = sum(1 for item in required_rows if rule_status(item) == STATUS_TIDAK_PATUH)
    optional_found = sum(1 for item in optional_rows if item.get("found"))

    return {
        "rules_total": len(rows),
        "required_total": len(required_rows),
        "required_ok": required_ok,
        "required_partial": required_partial,
        "required_missing": required_missing,
        "optional_found": optional_found,
        "extra_total": len(extras),
    }


#INI FUNGSI/BAGIAN UNTUK (menghitung skor rata-rata seluruh komponen)
def overall_score(result: Dict[str, Any]) -> float:
    required_rows = [item for item in result.get("validation_results", []) if item.get("required")]
    if not required_rows:
        return 0.0
    total = sum(rule_score(item) for item in required_rows)
    return round((total / len(required_rows)) * 100, 2)


#INI FUNGSI/BAGIAN UNTUK (mengubah skor angka menjadi status PATUH/KURANG/TIDAK)
def score_to_status(score: float) -> str:
    """Konversi nilai 0-100 menjadi status, disamakan dengan aturan RPS."""
    if score >= 85:
        return STATUS_PATUH
    if score >= 60:
        return STATUS_KURANG_PATUH
    return STATUS_TIDAK_PATUH


#INI FUNGSI/BAGIAN UNTUK (keseluruhan status)
def overall_status(result: Dict[str, Any]) -> str:
    """Status keseluruhan berbasis skor, sama seperti RPS.

    - skor >= 85 -> PATUH
    - skor 60-84 -> KURANG PATUH
    - skor < 60 -> TIDAK PATUH
    """
    return score_to_status(overall_score(result))


# =========================================================
# Summary and metadata tables
# =========================================================


#INI FUNGSI/BAGIAN UNTUK (membuat tabel metadata dokumen)
def build_metadata_table(result: Dict[str, Any]) -> Table:
    rows = [
        ["Nama Dokumen", source_file_name(result)],
    ]
    table = Table(rows, colWidths=[3.5 * cm, 11.0 * cm])
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.40, colors.HexColor(TELKOM_BORDER)),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor(TELKOM_LIGHT)),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor(TELKOM_NAVY)),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8.4),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return table


#INI FUNGSI/BAGIAN UNTUK (membuat tabel ringkasan hasil audit)
def build_summary_table(result: Dict[str, Any]) -> Table:
    summary = count_summary(result)
    status = overall_status(result)
    score = overall_score(result)

    rows = [
        ["Status Keseluruhan", status],
        ["Nilai Kepatuhan", f"{score}/100"],
        ["Aturan Status", ">=85 PATUH | 60-84 KURANG PATUH | <60 TIDAK PATUH"],
        ["Bagian Wajib Patuh", f"{summary['required_ok']}/{summary['required_total']}"],
        ["Bagian Wajib Tidak Patuh", str(summary["required_missing"])],
        ["Opsional yang Ada", str(summary["optional_found"])],
        ["Temuan Tambahan", str(summary["extra_total"])],
        ["Daftar Isi", "Ditemukan" if result.get("toc_found") else "Tidak ditemukan"],
    ]

    table = Table(rows, colWidths=[7.3 * cm, 6.7 * cm])
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor(TELKOM_BORDER)),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(TELKOM_LIGHT)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(TELKOM_NAVY)),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9.6),
                ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFCFF")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return table


# =========================================================
# Grouping rules and extras
# =========================================================


#INI FUNGSI/BAGIAN UNTUK (memetakan setiap aturan ke induknya)
def rule_parent_map(result: Dict[str, Any]) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for item in result.get("validation_results", []):
        field = clean_text(item.get("field"))
        parent = clean_text(item.get("parent"))
        if field != "-" and parent != "-":
            mapping[normalize(field)] = parent
    return mapping


#INI FUNGSI/BAGIAN UNTUK (menentukan bagian induk (root) dari sebuah aturan)
def root_for_rule(item: Dict[str, Any], parent_map: Dict[str, str]) -> str:
    """Find the top parent using only Parent values from RULES.xlsx."""
    field = clean_text(item.get("field"))
    parent = clean_text(item.get("parent"))

    if parent == "-":
        return field

    current = parent
    safety = 0
    while normalize(current) in parent_map and safety < 25:
        current = parent_map[normalize(current)]
        safety += 1
    return current


#INI FUNGSI/BAGIAN UNTUK (mengambil nomor bagian induk)
def root_number(value: Any) -> str:
    """Ambil nomor bab teratas dari nomor heading.

    Contoh: 2.1 -> 2, 3.2.1 -> 3.
    Kalau tidak ada nomor, hasilnya "-".
    """
    number = clean_text(value)
    if number == "-":
        return "-"
    first = number.split(".")[0].strip()
    return first if first.isdigit() else "-"


#INI FUNGSI/BAGIAN UNTUK (menentukan bagian induk untuk temuan tambahan)
def root_for_extra(
    extra: Dict[str, Any],
    toc_entries: List[Dict[str, Any]],
    number_to_root: Dict[str, str] | None = None,
) -> str:
    """Cari parent utama untuk temuan tambahan.

    Prioritas:
    1. nomor heading tambahan, misalnya 2.1 -> bab 2;
    2. mapping nomor bab yang diinfer dari rules dan hasil daftar isi;
    3. judul parent dari parser daftar isi jika ada;
    4. fallback terakhir: Temuan Tambahan.

    Dengan cara ini, temuan tambahan tidak dipisah di grup sendiri kalau nomor
    heading-nya masih bisa dikaitkan dengan bagian utama di dokumen.
    """
    number = clean_text(extra.get("number"))
    number_to_root = number_to_root or {}

    rn = root_number(number)
    if rn != "-" and rn in number_to_root:
        return number_to_root[rn]

    if number != "-":
        root_num = number.split(".")[0]
        for entry in toc_entries:
            if clean_text(entry.get("number")) == root_num:
                return clean_text(entry.get("title"))

    parent_title = clean_text(extra.get("parent_title"))
    if parent_title != "-":
        return parent_title

    return "Temuan Tambahan"


#INI FUNGSI/BAGIAN UNTUK (memetakan nomor bagian ke induknya)
def build_number_to_root_map(result: Dict[str, Any], groups: List[Dict[str, Any]]) -> Dict[str, str]:
    """Bangun mapping nomor bab -> nama bagian utama secara dinamis.

    Beberapa PDF hanya menulis sub-header di Daftar Isi, misalnya 2.1, 2.2,
    tetapi judul bab 2 tidak ikut terbaca sebagai item TOC. Agar temuan tambahan
    seperti 2.1 tetap masuk ke parent yang benar, generator menginfer urutan bab
    dari rules dan nomor yang berhasil ditemukan.

    Tidak ada nama bagian yang di-hardcode di sini. Semua mengikuti urutan rules
    dan nomor heading yang muncul di hasil validasi.
    """
    parent_map = rule_parent_map(result)
    name_to_known_number: Dict[str, int] = {}

    for item in visible_rule_rows(result):
        rn = root_number(item.get("matched_number"))
        if rn == "-":
            continue
        root = root_for_rule(item, parent_map)
        try:
            name_to_known_number[normalize(root)] = int(rn)
        except ValueError:
            continue

    number_to_root: Dict[str, str] = {}

    # 1) Mapping langsung dari root yang punya nomor terdeteksi.
    for group in groups:
        key = normalize(group.get("name"))
        if key in name_to_known_number:
            number_to_root[str(name_to_known_number[key])] = clean_text(group.get("name"))

    if not name_to_known_number:
        return number_to_root

    # 2) Mapping inferensi dari anchor pertama yang diketahui.
    # Misal Pendahuluan terdeteksi dari 1.1, maka grup setelahnya dianggap bab 2, 3, dst.
    anchors: List[Tuple[int, int]] = []
    for idx, group in enumerate(groups):
        key = normalize(group.get("name"))
        if key in name_to_known_number:
            anchors.append((idx, name_to_known_number[key]))

    if not anchors:
        return number_to_root

    anchor_idx, anchor_number = sorted(anchors, key=lambda item: item[1])[0]
    for idx in range(anchor_idx, len(groups)):
        inferred_number = anchor_number + (idx - anchor_idx)
        if inferred_number <= 0:
            continue
        number_to_root.setdefault(str(inferred_number), clean_text(groups[idx].get("name")))

    # 3) Anchor lain yang diketahui tetap menang kalau ada.
    for idx, known_number in anchors:
        number_to_root[str(known_number)] = clean_text(groups[idx].get("name"))

    return number_to_root


#INI FUNGSI/BAGIAN UNTUK (mengelompokkan aturan ke dalam bagian utama)
def build_grouped_sections(result: Dict[str, Any]) -> List[Dict[str, Any]]:
    parent_map = rule_parent_map(result)
    groups: List[Dict[str, Any]] = []
    group_index: Dict[str, Dict[str, Any]] = {}

    #INI FUNGSI/BAGIAN UNTUK (memastikan sebuah grup bagian sudah ada)
    def ensure_group(name: str, order: int = 10_000) -> Dict[str, Any]:
        key = normalize(name)
        if key not in group_index:
            group = {"name": name, "rules": [], "extras": [], "order": order}
            group_index[key] = group
            groups.append(group)
        else:
            group = group_index[key]
            group["order"] = min(group.get("order", order), order)
        return group

    for order, item in enumerate(visible_rule_rows(result)):
        root = root_for_rule(item, parent_map)
        ensure_group(root, order)["rules"].append(item)

    # Urutkan dulu supaya inferensi nomor bab mengikuti urutan rules/laporan.
    groups.sort(key=lambda group: group.get("order", 10_000))
    number_to_root = build_number_to_root_map(result, groups)

    toc_entries = result.get("toc_entries", [])
    for extra in result.get("extra_sections", []):
        root = root_for_extra(extra, toc_entries, number_to_root)
        ensure_group(root)["extras"].append(extra)

    groups.sort(key=lambda group: group.get("order", 10_000))
    return groups


#INI FUNGSI/BAGIAN UNTUK (menghitung skor sebuah grup bagian)
def group_score(group: Dict[str, Any]) -> float:
    required_rules = [item for item in group.get("rules", []) if item.get("required")]
    if not required_rules:
        return 100.0
    total = sum(rule_score(item) for item in required_rules)
    return round((total / len(required_rules)) * 100, 2)


#INI FUNGSI/BAGIAN UNTUK (menentukan status sebuah grup bagian)
def group_status(group: Dict[str, Any]) -> str:
    return score_to_status(group_score(group))


#INI FUNGSI/BAGIAN UNTUK (membuat ringkasan sebuah grup bagian)
def group_summary(group: Dict[str, Any]) -> str:
    rules = group.get("rules", [])
    extras = group.get("extras", [])
    required = [item for item in rules if item.get("required")]
    optional = [item for item in rules if not item.get("required")]

    required_ok = sum(1 for item in required if rule_status(item) == STATUS_PATUH)
    required_partial = sum(1 for item in required if rule_status(item) == STATUS_KURANG_PATUH)
    required_missing = sum(1 for item in required if rule_status(item) == STATUS_TIDAK_PATUH)
    optional_found = sum(1 for item in optional if item.get("found"))

    parts = []
    if required:
        parts.append(f"patuh {required_ok}/{len(required)}")
    if required_missing:
        parts.append(f"tidak patuh {required_missing}")
    if optional_found:
        parts.append(f"opsional ada {optional_found}")
    if extras:
        parts.append(f"tambahan {len(extras)}")
    return ", ".join(parts) if parts else "-"


#INI FUNGSI/BAGIAN UNTUK (memisahkan item aturan per status)
def split_rule_items(group: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
    patuh: List[Dict[str, Any]] = []
    kurang_patuh: List[Dict[str, Any]] = []
    tidak_patuh: List[Dict[str, Any]] = []
    opsional_ada: List[Dict[str, Any]] = []

    for item in group.get("rules", []):
        status = rule_status(item)
        if item.get("required") and status == STATUS_PATUH:
            patuh.append(item)
        elif item.get("required") and status == STATUS_KURANG_PATUH:
            kurang_patuh.append(item)
        elif item.get("required") and status == STATUS_TIDAK_PATUH:
            tidak_patuh.append(item)
        elif not item.get("required") and item.get("found"):
            opsional_ada.append(item)

    return patuh, kurang_patuh, tidak_patuh, opsional_ada


# =========================================================
# Table details
# =========================================================


#INI FUNGSI/BAGIAN UNTUK (memformat satu item aturan menjadi teks)
def format_rule_item(item: Dict[str, Any]) -> str:
    title = title_with_number(item)
    page = page_text(item)
    if page:
        return f"- {safe_text(title)} <font color='{TELKOM_MUTED}'>({page})</font>"
    return f"- {safe_text(title)}"


#INI FUNGSI/BAGIAN UNTUK (memformat satu temuan tambahan menjadi teks)
def format_extra_item(extra: Dict[str, Any]) -> str:
    title = title_with_number(extra, title_key="title", number_key="number")
    page = page_text(extra)
    if page:
        return f"- {safe_text(title)} <font color='{TELKOM_MUTED}'>({page})</font>"
    return f"- {safe_text(title)}"


#INI FUNGSI/BAGIAN UNTUK (membuat blok teks satu bagian berwarna)
def section_block(label: str, color: str, lines: List[str]) -> str:
    if not lines:
        return ""
    header = f"<font color='{color}'><b>{safe_text(label)}</b></font>"
    return header + "<br/>" + "<br/>".join(lines)


#INI FUNGSI/BAGIAN UNTUK (menyusun detail grup dalam bentuk HTML)
def build_group_detail_html(group: Dict[str, Any]) -> str:
    patuh, kurang_patuh, tidak_patuh, opsional_ada = split_rule_items(group)
    extras = group.get("extras", [])

    blocks = [f"<b>Ringkasan:</b> {safe_text(group_summary(group))}"]

    patuh_lines = [format_rule_item(item) for item in patuh]
    tidak_lines = [format_rule_item(item) for item in tidak_patuh]
    optional_lines = [format_rule_item(item) for item in opsional_ada]
    extra_lines = [format_extra_item(extra) for extra in extras]

    if patuh_lines:
        blocks.append(section_block("Bagian wajib patuh", TELKOM_GREEN, patuh_lines))
    if tidak_lines:
        blocks.append(section_block("Bagian wajib tidak patuh", TELKOM_DANGER, tidak_lines))
    if optional_lines:
        blocks.append(section_block("Bagian opsional ditemukan", TELKOM_MUTED, optional_lines))
    if extra_lines:
        blocks.append(section_block("Temuan tambahan", TELKOM_NAVY, extra_lines))

    return "<br/><br/>".join(block for block in blocks if block)


#INI FUNGSI/BAGIAN UNTUK (membuat tabel aturan yang sudah dikelompokkan)
def build_grouped_rules_table(result: Dict[str, Any], styles: Dict[str, ParagraphStyle]) -> Table:
    groups = build_grouped_sections(result)
    rows: List[List[Any]] = [["No", "Bagian Utama", "Hasil Pengecekan", "Nilai", "Status"]]

    for index, group in enumerate(groups, start=1):
        score = group_score(group)
        status = group_status(group)
        color = color_for(status)
        detail_html = build_group_detail_html(group)

        rows.append(
            [
                Paragraph(str(index), styles["cell_center"]),
                Paragraph(f"<b>{safe_text(group.get('name'))}</b>", styles["cell"]),
                Paragraph(detail_html, styles["cell_small"]),
                Paragraph(f"{score}/100", styles["cell_center"]),
                Paragraph(f"<font color='{color}'><b>{safe_text(status)}</b></font>", styles["cell_center"]),
            ]
        )

    if len(rows) == 1:
        rows.append(["-", "Tidak ada data", "-", "-", "-"])

    table = Table(rows, colWidths=[0.75 * cm, 3.45 * cm, 10.65 * cm, 1.45 * cm, 1.9 * cm], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(TELKOM_NAVY)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 8.8),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor(TELKOM_BORDER)),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor(TELKOM_LIGHT), colors.white]),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return table


# =========================================================
# Conclusion
# =========================================================


#INI FUNGSI/BAGIAN UNTUK (membuat kotak kesimpulan audit)
def build_conclusion_box(result: Dict[str, Any], styles: Dict[str, ParagraphStyle]) -> Table:
    status = overall_status(result)
    summary = count_summary(result)
    score = overall_score(result)

    if status == STATUS_PATUH:
        message = (
            "Dokumen dinyatakan patuh karena nilai kepatuhan mencapai batas PATUH. "
            f"Nilai kepatuhan: {score}/100. Bagian opsional yang tidak ada tidak dihitung sebagai masalah."
        )
        background = TELKOM_SOFT_GREEN
    elif status == STATUS_KURANG_PATUH:
        message = (
            "Dokumen dinyatakan kurang patuh karena nilai kepatuhan berada pada rentang 60 sampai 84. "
            f"Nilai kepatuhan: {score}/100."
        )
        background = TELKOM_SOFT_ORANGE
    else:
        message = (
            "Dokumen dinyatakan tidak patuh karena nilai kepatuhan berada di bawah 60. "
            f"Bagian wajib tidak patuh: {summary['required_missing']}. Nilai kepatuhan: {score}/100."
        )
        background = TELKOM_SOFT_RED

    table = Table(
        [[Paragraph("<b>Kesimpulan</b>", styles["box_title"])], [Paragraph(safe_text(message), styles["body"])]],
        colWidths=[14.2 * cm],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(background)),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor(TELKOM_BORDER)),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return table


# =========================================================
# Main generator
# =========================================================


#INI FUNGSI/BAGIAN UNTUK (fungsi utama: membuat PDF laporan Buku Kurikulum)
def generate_buku_kurikulum_pdf(result: Dict[str, Any], output_path: str) -> str:
    output_path = resolve_report_output_path(output_path, result, default_names=("laporan.pdf",))
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    styles = build_styles()
    base_styles = styles["base"]
    document_name = source_file_name(result)    
    status = overall_status(result)
    color = color_for(status)
    generated_at = datetime.now().strftime("%d-%m-%Y %H:%M")

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=1.35 * cm,
        rightMargin=1.35 * cm,
        topMargin=2.2 * cm,
        bottomMargin=1.6 * cm,
        title="Laporan Audit Buku Kurikulum",
        author="Telkom University",
    )

    story = [
    Paragraph("TELKOM UNIVERSITY", styles["cover_subtitle"]),
    Paragraph("Sistem Pengecekan Otomatis", styles["subtitle"]),
    Spacer(1, 10),

    Paragraph("LAPORAN AUDIT BUKU KURIKULUM", styles["cover_title"]),
    Paragraph("Pengecekan Kelengkapan Dokumen Buku Kurikulum", styles["heading"]),
    Spacer(1, 8),

    Paragraph("Dokumen", styles["subtitle"]),
    Paragraph(safe_text(document_name), base_styles["Heading3"]),
    Spacer(1, 4),

    Paragraph(f"File: {safe_text(display_name(result.get('file')))}", styles["muted"]),
    Paragraph(f"Rules: {safe_text(display_name(result.get('rules_path')))}", styles["muted"]),

    Spacer(1, 16),
    build_summary_table(result),
    Spacer(1, 12),

    Paragraph(f"Status: <font color='{color}'><b>{safe_text(status)}</b></font>", styles["heading"]),
    Spacer(1, 8),

    build_conclusion_box(result, styles),
    Spacer(1, 10),

    Paragraph(
        "Catatan: Laporan ini hanya mengecek keberadaan bagian. Sistem tidak menilai benar atau salahnya isi dokumen.",
        styles["small_center"],
    ),

    PageBreak(),

    Paragraph("HASIL PENGECEKAN BAGIAN", styles["section_heading"]),
    Paragraph(
        "Tabel dibuat per bagian utama. Bagian opsional yang tidak ditemukan tidak ditampilkan. Header atau sub-header tambahan dari Daftar Isi dimasukkan pada bagian terkait sebagai informasi tambahan.",
        styles["muted"],
    ),
    Spacer(1, 6),
    build_grouped_rules_table(result, styles),
]

    #INI FUNGSI/BAGIAN UNTUK (menggambar header dan footer di setiap halaman PDF)
    def on_page(canvas, doc_obj):
        width, height = A4
        canvas.saveState()
        canvas.setFillColor(colors.HexColor(TELKOM_NAVY))
        canvas.rect(
            doc.leftMargin,
            height - 1.42 * cm,
            width - doc.leftMargin - doc.rightMargin,
            0.03 * cm,
            fill=True,
            stroke=False,
        )
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(colors.HexColor(TELKOM_NAVY))
        canvas.drawString(doc.leftMargin, height - 1.2 * cm, "Sistem Audit Buku Kurikulum - Telkom University")
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawString(doc.leftMargin, 0.8 * cm, f"Generated: {generated_at}")
        canvas.drawRightString(width - doc.rightMargin, 0.8 * cm, f"Halaman {doc_obj.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return output_path


# Backward-compatible alias, supaya aman kalau ada pemanggilan lama.
#INI FUNGSI/BAGIAN UNTUK (fungsi utama: membuat PDF laporan Laporan Penelitian)
def generate_laporan_penelitian_pdf(result: Dict[str, Any], output_path: str) -> str:
    return generate_buku_kurikulum_pdf(result, output_path)