#!/usr/bin/env python3

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    KeepTogether,
)


STATUS_PATUH = "PATUH"
STATUS_KURANG_PATUH = "KURANG PATUH"
STATUS_TIDAK_PATUH = "TIDAK PATUH"

TELKOM_RED = "#E30613"
TELKOM_NAVY = "#0A2C70"
TELKOM_LIGHT = "#F3F6FB"
TELKOM_BORDER = "#D6DEEE"
TELKOM_TEXT = "#1F2937"
TELKOM_MUTED = "#6B7280"
TELKOM_SOFT_RED = "#FFF0F2"

STATUS_COLORS = {
    STATUS_PATUH: "#16823A",
    STATUS_KURANG_PATUH: "#D58A00",
    STATUS_TIDAK_PATUH: "#B00020",
    "SESUAI": "#16823A",
    "KURANG SESUAI": "#D58A00",
    "TIDAK SESUAI": "#B00020",
}


HTML_BREAK = "<br/>"


#INI FUNGSI/BAGIAN UNTUK (merapikan teks: hilangkan newline dan spasi berlebih)
def clean_text(value: Any) -> Any:
    if isinstance(value, str):
        value = value.replace("\r\n", " ")
        value = value.replace("\n", " ")
        value = value.replace("\r", " ")
        value = re.sub(r"\s+", " ", value)
        return value.strip()
    return value


#INI FUNGSI/BAGIAN UNTUK (membersihkan seluruh teks di dalam struktur data secara rekursif)
def clean_data(data: Any) -> Any:
    if isinstance(data, dict):
        return {key: clean_data(value) for key, value in data.items()}
    if isinstance(data, list):
        return [clean_data(item) for item in data]
    return clean_text(data)


#INI FUNGSI/BAGIAN UNTUK (mengubah nilai apa pun menjadi teks aman untuk PDF)
def safe_text(value: Any) -> str:
    if value is None:
        return "-"
    value = clean_text(value)
    if value is None or value == "":
        return "-"
    return escape(str(value))


#INI FUNGSI/BAGIAN UNTUK (memendekkan teks panjang dengan tanda titik-titik)
def shorten(value: Any, limit: int = 170) -> str:
    text = str(clean_text(value) or "-")
    if len(text) <= limit:
        return text
    return text[:limit].rstrip() + "..."


#INI FUNGSI/BAGIAN UNTUK (menampilkan nama file saja tanpa folder dan tanpa angka acak)
def display_name(value: Any) -> str:
    # Ambil nama file saja (tanpa folder) dan buang prefix angka acak/timestamp agar tampil apa adanya.
    text = clean_text(value)
    if not text or text == "-":
        return "-"
    name = Path(str(text)).name
    name = re.sub(r"^\d{10,}[-_\s]+", "", name).strip()
    return name or "-"


#INI FUNGSI/BAGIAN UNTUK (mengambil nama dasar file input untuk nama laporan)
def report_input_stem(result: Dict[str, Any]) -> str:
    file_path = clean_text(result.get("file"))
    if not file_path or file_path == "-":
        return "dokumen"
    stem = Path(str(file_path)).stem
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


#INI FUNGSI/BAGIAN UNTUK (memilih warna sesuai status kepatuhan)
def get_status_color(status: str) -> str:
    return STATUS_COLORS.get(status, TELKOM_NAVY)


#INI FUNGSI/BAGIAN UNTUK (memilih warna sesuai tingkat prioritas)
def get_priority_color(priority: str) -> str:
    priority = str(priority or "").upper()
    if "TINGGI" in priority:
        return "#B00020"
    if "MENENGAH" in priority:
        return "#D58A00"
    return TELKOM_NAVY


#INI FUNGSI/BAGIAN UNTUK (mengubah skor angka menjadi status PATUH/KURANG/TIDAK)
def score_to_status(score: float) -> str:
    if score >= 85:
        return STATUS_PATUH
    if score >= 60:
        return STATUS_KURANG_PATUH
    return STATUS_TIDAK_PATUH


#INI FUNGSI/BAGIAN UNTUK (menghitung skor satu header (ditemukan dan terisi))
def score_header(header: Dict[str, Any]) -> int:
    found = bool(header.get("found"))
    has_value = bool(header.get("has_value"))
    value_required = bool(header.get("value_required", True))
    if not found:
        return 0
    if has_value:
        return 100
    # Field ketemu tapi value kosong.
    # Optional (value_required=False): value tidak dinilai -> presence sudah cukup -> patuh penuh.
    if not value_required:
        return 100
    return 50


#INI FUNGSI/BAGIAN UNTUK (menghitung skor satu komponen/kategori)
def score_category(category: Dict[str, Any]) -> float:
    if category.get("has_header_rules"):
        headers = category.get("headers", [])
        if not headers:
            return 0
        scores = [score_header(header) for header in headers]
        return round(sum(scores) / len(scores), 2) if scores else 0
    category_found = bool(category.get("category_found"))
    has_value = bool(category.get("has_value"))
    value_required = bool(category.get("value_required", True))
    if not category_found:
        return 0
    if has_value:
        return 100
    # Kategori tanpa field, ketemu tapi value kosong.
    # Optional (value_required=False): presence kategori sudah cukup -> patuh penuh.
    if not value_required:
        return 100
    return 50


#INI FUNGSI/BAGIAN UNTUK (menghitung skor rata-rata seluruh komponen)
def overall_score(categories: List[Dict[str, Any]]) -> float:
    if not categories:
        return 0
    scores = [score_category(category) for category in categories]
    return round(sum(scores) / len(scores), 2)


#INI FUNGSI/BAGIAN UNTUK (menentukan status keseluruhan dari skor)
def get_overall_status(score: float) -> str:
    return score_to_status(score)


#INI FUNGSI/BAGIAN UNTUK (mengambil tanggal dan jam saat ini sebagai teks)
def get_now_text() -> str:
    return datetime.now().strftime("%d-%m-%Y %H:%M")


#INI FUNGSI/BAGIAN UNTUK (mengambil nama mata kuliah dari hasil validasi)
def get_course_name(result: Dict[str, Any]) -> str:
    for category in result.get("categories", []):
        if category.get("category") != "Identitas Mata Kuliah":
            continue
        for header in category.get("headers", []):
            if str(header.get("header", "")).lower() == "nama mk":
                values = header.get("values", [])
                if values and values[0]:
                    return str(values[0])
    return "Dokumen RPS"


#INI FUNGSI/BAGIAN UNTUK (menghitung jumlah header yang ditemukan dan terisi)
def count_headers(categories: List[Dict[str, Any]]) -> Dict[str, int]:
    total = 0
    found = 0
    value_filled = 0
    for category in categories:
        if not category.get("has_header_rules"):
            continue
        for header in category.get("headers", []):
            total += 1
            if header.get("found"):
                found += 1
            if header.get("has_value"):
                value_filled += 1
    return {"total": total, "found": found, "value_filled": value_filled}


#INI FUNGSI/BAGIAN UNTUK (mengubah boolean menjadi label Ada/Tidak Ada)
def bool_label(value: bool) -> str:
    return "Ada" if value else "Tidak Ada"


#INI FUNGSI/BAGIAN UNTUK (menentukan status keberadaan header (Ditemukan/Tidak))
def header_presence_status(header: Dict[str, Any]) -> Tuple[str, str]:
    """Status keberadaan header secara eksak: (label, warna)."""
    if header.get("found"):
        return "Ditemukan", STATUS_COLORS[STATUS_PATUH]
    return "Tidak Ditemukan", STATUS_COLORS[STATUS_TIDAK_PATUH]


#INI FUNGSI/BAGIAN UNTUK (menentukan status keterisian value header (Terisi/Kosong))
def header_value_status(header: Dict[str, Any]) -> Tuple[str, str]:
    """Status keterisian value secara eksak: (label, warna)."""
    if not header.get("found"):
        return "-", STATUS_COLORS[STATUS_TIDAK_PATUH]
    if header.get("has_value"):
        return "Terisi", STATUS_COLORS[STATUS_PATUH]
    return "Kosong", STATUS_COLORS[STATUS_KURANG_PATUH]


#INI FUNGSI/BAGIAN UNTUK (membuat ringkasan singkat status sebuah komponen)
def category_summary_text(category: Dict[str, Any]) -> str:
    if category.get("has_header_rules"):
        headers = category.get("headers", [])
        total = len(headers)
        found = sum(1 for header in headers if header.get("found"))
        value_filled = sum(1 for header in headers if header.get("has_value"))
        return f"Header ditemukan {found}/{total}, value terisi {value_filled}/{total}"
    category_found = bool_label(bool(category.get("category_found")))
    has_value = bool_label(bool(category.get("has_value")))
    return f"Komponen {category_found}, value {has_value}"


#INI FUNGSI/BAGIAN UNTUK (menyusun temuan detail per komponen (header/value yang kurang))
def category_finding(category: Dict[str, Any]) -> str:
    if category.get("has_header_rules"):
        headers = category.get("headers", [])
        missing_headers = [str(header.get("header")) for header in headers if not header.get("found")]
        empty_values = [str(header.get("header")) for header in headers if header.get("found") and not header.get("has_value")]
        findings = []
        if missing_headers:
            findings.append("Header belum ditemukan: " + ", ".join(missing_headers))
        if empty_values:
            findings.append("Value belum terisi pada header: " + ", ".join(empty_values))
        if not findings:
            return "Seluruh header ditemukan dan seluruh value yang dibutuhkan sudah terisi."
        return ". ".join(findings) + "."
    if category.get("category_found") and category.get("has_value"):
        return "Komponen ditemukan dan value sudah terisi."
    if category.get("category_found") and not category.get("has_value"):
        return "Komponen ditemukan, tetapi value belum terisi."
    return "Komponen tidak ditemukan pada dokumen."


#INI FUNGSI/BAGIAN UNTUK (menyusun rekomendasi tindak lanjut per komponen)
def category_action(category: Dict[str, Any]) -> str:
    score = score_category(category)
    if score >= 85:
        return "Tidak diperlukan tindak lanjut."
    if category.get("has_header_rules"):
        return "Lengkapi header atau value yang belum sesuai. Apabila template berubah, sesuaikan rules agar header dapat dikenali oleh sistem."
    return "Lengkapi value pada komponen terkait. Jika komponen tersebut tidak wajib pada template terbaru, sesuaikan rules."


#INI FUNGSI/BAGIAN UNTUK (menentukan prioritas perbaikan komponen dari skor)
def category_priority(category: Dict[str, Any]) -> str:
    score = score_category(category)
    if score < 60:
        return "TINGGI"
    if score < 85:
        return "MENENGAH"
    return "RENDAH"


#INI FUNGSI/BAGIAN UNTUK (menyiapkan kumpulan gaya teks/paragraf untuk PDF)
def build_styles():
    styles = getSampleStyleSheet()
    cover_title = ParagraphStyle(
        "cover_title",
        parent=styles["Title"],
        fontSize=22,
        leading=27,
        alignment=1,
        textColor=colors.HexColor(TELKOM_RED),
        spaceAfter=6,
    )
    cover_subtitle = ParagraphStyle(
        "cover_subtitle",
        parent=styles["Heading2"],
        fontSize=14,
        leading=18,
        alignment=1,
        textColor=colors.HexColor(TELKOM_NAVY),
        spaceAfter=8,
    )
    subtitle_style = ParagraphStyle(
        "subtitle_style",
        parent=styles["BodyText"],
        fontSize=10,
        leading=13,
        textColor=colors.HexColor(TELKOM_NAVY),
        alignment=1,
    )
    heading_style = ParagraphStyle(
        "heading_style",
        parent=styles["Heading2"],
        fontSize=13,
        leading=16,
        textColor=colors.HexColor(TELKOM_NAVY),
        spaceAfter=6,
        alignment=1,
    )
    section_heading = ParagraphStyle(
        "section_heading",
        parent=styles["Heading2"],
        fontSize=14,
        leading=17,
        textColor=colors.HexColor(TELKOM_NAVY),
        spaceBefore=4,
        spaceAfter=8,
    )
    body_style = ParagraphStyle(
        "body_style",
        parent=styles["BodyText"],
        fontSize=9,
        leading=12,
        textColor=colors.HexColor(TELKOM_TEXT),
    )
    cell_style = ParagraphStyle(
        "cell_style",
        parent=styles["BodyText"],
        fontSize=8.7,
        leading=10.6,
        textColor=colors.HexColor(TELKOM_TEXT),
    )
    small_center = ParagraphStyle(
        "small_center",
        parent=styles["Normal"],
        fontSize=9,
        leading=12,
        alignment=1,
        textColor=colors.grey,
    )
    analysis_label = ParagraphStyle(
        "analysis_label",
        parent=styles["BodyText"],
        fontSize=8.8,
        leading=10.2,
        textColor=colors.HexColor(TELKOM_NAVY),
    )
    analysis_text = ParagraphStyle(
        "analysis_text",
        parent=styles["BodyText"],
        fontSize=8.6,
        leading=11,
        textColor=colors.HexColor(TELKOM_TEXT),
    )
    muted_text = ParagraphStyle(
        "muted_text",
        parent=styles["BodyText"],
        fontSize=8.6,
        leading=11.4,
        textColor=colors.HexColor(TELKOM_MUTED),
    )
    box_title = ParagraphStyle(
        "box_title",
        parent=styles["Heading3"],
        fontSize=10,
        leading=13,
        textColor=colors.HexColor(TELKOM_NAVY),
        spaceAfter=3,
    )
    return {
        "base": styles,
        "cover_title": cover_title,
        "cover_subtitle": cover_subtitle,
        "subtitle": subtitle_style,
        "heading": heading_style,
        "section_heading": section_heading,
        "body": body_style,
        "cell": cell_style,
        "small_center": small_center,
        "analysis_label": analysis_label,
        "analysis_text": analysis_text,
        "muted": muted_text,
        "box_title": box_title,
    }


#INI FUNGSI/BAGIAN UNTUK (membuat tabel ringkasan hasil audit)
def build_summary_table(result: Dict[str, Any], categories: List[Dict[str, Any]]) -> Table:
    summary = result.get("summary", {})
    score = overall_score(categories)
    status = get_overall_status(score)
    headers_count = count_headers(categories)
    total_components = summary.get("total_categories", len(categories))
    valid_components = summary.get("valid_categories", 0)
    invalid_components = summary.get("invalid_categories", total_components - valid_components)
    alignment_total = len(result.get("alignment_results", []))
    alignment_ok = sum(1 for item in result.get("alignment_results", []) if item.get("status") == "SESUAI")
    table = Table(
        [
            ["Status Keseluruhan", status],
            ["Nilai Audit", f"{score}/100"],
            ["Komponen Patuh", f"{valid_components} komponen"],
            ["Komponen Kurang/Tidak Patuh", f"{invalid_components} komponen"],
            ["Header Ditemukan", f"{headers_count['found']}/{headers_count['total']} header"],
            ["Value Header Terisi", f"{headers_count['value_filled']}/{headers_count['total']} header"],
            ["Alignment Sesuai", f"{alignment_ok}/{alignment_total} baris" if alignment_total else "-"],
            ["Aturan Status", ">=85 PATUH | 60-84 KURANG PATUH | <60 TIDAK PATUH"],
            ["Total Komponen", f"{total_components} komponen"],
        ],
        colWidths=[7.2 * cm, 6.8 * cm],
    )
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor(TELKOM_BORDER)),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(TELKOM_LIGHT)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(TELKOM_NAVY)),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9.7),
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


#INI FUNGSI/BAGIAN UNTUK (menyusun baris-baris tabel komponen)
def build_component_rows(categories: List[Dict[str, Any]], styles: Dict[str, ParagraphStyle]) -> List[List[Any]]:
    rows = []
    for index, category in enumerate(categories, start=1):
        score = score_category(category)
        status = score_to_status(score)
        status_color = get_status_color(status)
        rows.append(
            [
                str(index),
                Paragraph(safe_text(category.get("category")), styles["body"]),
                Paragraph(category_summary_text(category), styles["cell"]),
                Paragraph(f"<font color='{status_color}'><b>{status}</b></font>", styles["body"]),
                Paragraph(f"{score}/100", styles["body"]),
            ]
        )
    return rows


#INI FUNGSI/BAGIAN UNTUK (membuat tabel status per komponen)
def build_component_table(categories: List[Dict[str, Any]], styles: Dict[str, ParagraphStyle]) -> Table:
    table = Table(
        [["No", "Komponen", "Ringkasan", "Status", "Nilai"]] + build_component_rows(categories, styles),
        colWidths=[1.0 * cm, 5.1 * cm, 6.2 * cm, 2.4 * cm, 1.8 * cm],
        repeatRows=1,
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(TELKOM_NAVY)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9.7),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor(TELKOM_BORDER)),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor(TELKOM_LIGHT), colors.white]),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return table


#INI FUNGSI/BAGIAN UNTUK (membuat tabel rincian status header dan value)
def build_header_detail_table(category: Dict[str, Any], styles: Dict[str, ParagraphStyle]) -> Table:
    headers = category.get("headers", [])
    rows = [["No", "Header", "Keberadaan", "Value"]]
    for index, header in enumerate(headers, start=1):
        presence_text, presence_color = header_presence_status(header)
        value_text, value_color = header_value_status(header)
        rows.append(
            [
                str(index),
                Paragraph(safe_text(header.get("header")), styles["cell"]),
                Paragraph(f"<font color='{presence_color}'><b>{presence_text}</b></font>", styles["cell"]),
                Paragraph(f"<font color='{value_color}'><b>{value_text}</b></font>", styles["cell"]),
            ]
        )
    table = Table(
        rows,
        colWidths=[0.9 * cm, 8.2 * cm, 3.4 * cm, 3.4 * cm],
        repeatRows=1,
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(TELKOM_LIGHT)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(TELKOM_NAVY)),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 8.6),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("ALIGN", (2, 0), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor(TELKOM_BORDER)),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFCFF")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table


#INI FUNGSI/BAGIAN UNTUK (menyusun bagian analisis per komponen)
def build_analysis_story(categories: List[Dict[str, Any]], styles: Dict[str, ParagraphStyle]) -> List[Any]:
    story = [
        Paragraph("ANALISIS PER KOMPONEN", styles["section_heading"]),
        Spacer(1, 4),
        Paragraph(
            "Analisis berikut menjelaskan temuan dan tindak lanjut berdasarkan keberadaan komponen/header serta keterisian value. Isi value tidak ditampilkan pada laporan audit.",
            styles["muted"],
        ),
        Spacer(1, 6),
    ]
    for index, category in enumerate(categories, start=1):
        score = score_category(category)
        status = score_to_status(score)
        status_color = get_status_color(status)
        priority = category_priority(category)
        priority_color = get_priority_color(priority)
        section_title = Paragraph(
            f"{index}. <b>{safe_text(category.get('category'))}</b> - <font color='{status_color}'><b>{status}</b></font> ({score}/100)",
            styles["body"],
        )
        analysis_table = Table(
            [
                [Paragraph("Temuan", styles["analysis_label"]), Paragraph(category_finding(category), styles["analysis_text"])],
                [Paragraph("Tindak Lanjut", styles["analysis_label"]), Paragraph(category_action(category), styles["analysis_text"])],
                [Paragraph("Prioritas", styles["analysis_label"]), Paragraph(f"<font color='{priority_color}'><b>{priority}</b></font>", styles["analysis_text"])],
            ],
            colWidths=[3.2 * cm, 12.7 * cm],
        )
        analysis_table.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor(TELKOM_BORDER)),
                    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor(TELKOM_LIGHT)),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.extend([Spacer(1, 6), section_title, Spacer(1, 2), KeepTogether([analysis_table])])
        if category.get("has_header_rules") and category.get("headers"):
            story.append(Spacer(1, 4))
            story.append(Paragraph("Rincian status header dan value:", styles["muted"]))
            story.append(Spacer(1, 2))
            story.append(build_header_detail_table(category, styles))
    return story


#INI FUNGSI/BAGIAN UNTUK (membuat paragraf status kesesuaian berwarna)
def alignment_status_paragraph(status: str, styles: Dict[str, ParagraphStyle]) -> Paragraph:
    color = get_status_color(status)
    return Paragraph(f"<font color='{color}'><b>{safe_text(status)}</b></font>", styles["cell"])


#INI FUNGSI/BAGIAN UNTUK (membuat tabel analisis kesesuaian metode pembelajaran)
def build_alignment_table(alignment_results: List[Dict[str, Any]], styles: Dict[str, ParagraphStyle]) -> Table:
    if not alignment_results:
        return Table([["Tidak ada data alignment."]])

    first_item = alignment_results[0]

    description_header = (
        first_item.get("description_header")
        or "Deskripsi / Capaian"
    )

    method_header = (
        first_item.get("method_header")
        or "Metode Pembelajaran"
    )

    rows = [
        [
            "No",
            safe_text(description_header),
            safe_text(method_header),
            "Kesesuaian",
            "Analisis",
        ]
    ]

    for index, item in enumerate(alignment_results, start=1):
        description_value = item.get("description_value")
        method_value = item.get("method_value")

        status = item.get("status", "-")
        analysis = item.get("analysis", "-")
        expected = item.get("expected_methods") or []

        if status != "SESUAI" and expected:
            analysis = (
                f"{analysis} "
                f"Metode yang disarankan: {', '.join(expected)}."
            )

        rows.append(
            [
                str(index),
                Paragraph(safe_text(description_value), styles["cell"]),
                Paragraph(safe_text(method_value), styles["cell"]),
                alignment_status_paragraph(status, styles),
                Paragraph(safe_text(analysis), styles["analysis_text"]),
            ]
        )

    table = Table(
        rows,
        colWidths=[
            0.8 * cm,
            5.0 * cm,
            3.8 * cm,
            2.2 * cm,
            5.0 * cm,
        ],
        repeatRows=1,
    )

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
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )

    return table


#INI FUNGSI/BAGIAN UNTUK (menyusun bagian analisis kesesuaian metode)
def build_alignment_story(result: Dict[str, Any], styles: Dict[str, ParagraphStyle]) -> List[Any]:
    alignment_results = result.get("alignment_results", [])

    if not alignment_results:
        return []

    first_item = alignment_results[0]

    description_header = (
        first_item.get("description_header")
        or "Deskripsi / Capaian"
    )

    method_header = (
        first_item.get("method_header")
        or "Metode Pembelajaran"
    )

    alignment_summary = result.get("alignment_summary", {}) or {}

    total_rows = alignment_summary.get("total_rows", len(alignment_results))
    sesuai = alignment_summary.get("sesuai", 0)
    kurang_sesuai = alignment_summary.get("kurang_sesuai", 0)
    tidak_sesuai = alignment_summary.get("tidak_sesuai", 0)
    avg_score = alignment_summary.get("avg_score", "-")

    summary_table = Table(
        [
            ["Header 1", safe_text(description_header)],
            ["Header 2", safe_text(method_header)],
            ["Total Data", f"{total_rows} Minggu"],
            ["Sesuai", f"{sesuai} Minggu"],
            ["Kurang Sesuai", f"{kurang_sesuai} Minggu"],
            ["Tidak Sesuai", f"{tidak_sesuai} Minggu"],
            ["Nilai Rata-rata", f"{avg_score}/100" if avg_score != "-" else "-"],
        ],
        colWidths=[4.2 * cm, 10.2 * cm],
    )

    summary_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor(TELKOM_BORDER)),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor(TELKOM_LIGHT)),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )

    story = [
        PageBreak(),
        Paragraph("ANALISIS KESESUAIAN METODE PEMBELAJARAN", styles["section_heading"]),
        Spacer(1, 4),
        Paragraph(
            (
                "Bagian ini membandingkan value dari dua field yang ditandai compare = yes pada rules.xlsx. "
                "Nama field dan data yang ditampilkan mengikuti field yang ditandai pada rules."
            ),
            styles["muted"],
        ),
        Spacer(1, 6),
        summary_table,
        Spacer(1, 10),
        build_alignment_table(alignment_results, styles),
    ]

    return story


#INI FUNGSI/BAGIAN UNTUK (membuat kotak kesimpulan audit)
def build_conclusion_box(categories: List[Dict[str, Any]], styles: Dict[str, ParagraphStyle]) -> Table:
    score = overall_score(categories)
    status = get_overall_status(score)
    if status == STATUS_PATUH:
        message = "Dokumen RPS secara umum telah memenuhi kelengkapan komponen yang diperiksa. Perbaikan hanya diperlukan apabila terdapat catatan minor pada bagian analisis."
    elif status == STATUS_KURANG_PATUH:
        message = "Dokumen RPS sudah memiliki sebagian besar komponen, tetapi masih terdapat beberapa bagian yang perlu dilengkapi agar sesuai dengan rules."
    else:
        message = "Dokumen RPS memerlukan perbaikan signifikan karena beberapa komponen/header belum ditemukan atau value penting belum terisi."
    table = Table(
        [[Paragraph("<b>Kesimpulan Audit</b>", styles["box_title"])], [Paragraph(message, styles["body"])]],
        colWidths=[14 * cm],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(TELKOM_SOFT_RED)),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor(TELKOM_BORDER)),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return table


#INI FUNGSI/BAGIAN UNTUK (fungsi utama: membuat PDF laporan audit RPS)
def generate_audit_pdf(result: Dict[str, Any], output_path: str = "laporan_audit.pdf") -> str:
    result = clean_data(result)
    output_path = resolve_report_output_path(output_path, result, default_names=("laporan_audit.pdf", "laporan.pdf"))
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    categories = result.get("categories", [])
    score = overall_score(categories)
    status = get_overall_status(score)
    status_color = get_status_color(status)
    course_name = get_course_name(result)
    generated_at = get_now_text()
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=1.6 * cm,
        rightMargin=1.6 * cm,
        topMargin=2.2 * cm,
        bottomMargin=1.6 * cm,
        title="Laporan Audit Dokumen RPS",
        author="Telkom University",
    )
    styles = build_styles()
    base_styles = styles["base"]
    summary_table = build_summary_table(result, categories)
    status_line = Paragraph(f"Status: <font color='{status_color}'><b>{status}</b></font>", styles["heading"])
    cover_story = [
        Paragraph("TELKOM UNIVERSITY", styles["cover_subtitle"]),
        Paragraph("Sistem Pengecekan Otomatis", styles["subtitle"]),
        Spacer(1, 10),
        Paragraph("LAPORAN AUDIT DOKUMEN RPS", styles["cover_title"]),
        Paragraph("Pengecekan Kelengkapan Dokumen RPS", styles["heading"]),
        Spacer(1, 8),
        Paragraph("Mata Kuliah / Dokumen", styles["subtitle"]),
        Paragraph(safe_text(course_name), base_styles["Heading3"]),
        Spacer(1, 4),
        Paragraph(f"File: {safe_text(display_name(result.get('file')))}", styles["muted"]),
        Paragraph(f"Rules: {safe_text(display_name(result.get('rules_path') or result.get('settings', {}).get('rules_path')))}", styles["muted"]),
        Spacer(1, 16),
        summary_table,
        Spacer(1, 12),
        status_line,
        Spacer(1, 8),
        build_conclusion_box(categories, styles),
        Spacer(1, 10),
        Paragraph(
            "Catatan: Laporan ini dihasilkan secara otomatis untuk membantu proses audit dokumen RPS. Hasil tetap dapat diverifikasi kembali oleh tim akademik.",
            styles["small_center"],
        ),
        PageBreak(),
    ]
    components_story = [
        Paragraph("HASIL PENGECEKAN PER KOMPONEN", base_styles["Heading2"]),
        Spacer(1, 6),
        Paragraph(
            "Ringkasan ini menampilkan status komponen dan header berdasarkan rules. Isi value tidak ditampilkan agar laporan tetap ringkas.",
            styles["subtitle"],
        ),
        Spacer(1, 6),
        build_component_table(categories, styles),
        Spacer(1, 10),
    ]
    analysis_story = build_analysis_story(categories, styles)
    alignment_story = build_alignment_story(result, styles)
    story = cover_story + components_story + analysis_story + alignment_story

    #INI FUNGSI/BAGIAN UNTUK (menggambar header dan footer di setiap halaman PDF)
    def on_page(canvas, doc_obj):
        width, height = A4
        canvas.saveState()
        canvas.setFillColor(colors.HexColor(TELKOM_NAVY))
        canvas.rect(doc.leftMargin, height - 1.42 * cm, width - doc.leftMargin - doc.rightMargin, 0.03 * cm, fill=True, stroke=False)
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(colors.HexColor(TELKOM_NAVY))
        canvas.drawString(doc.leftMargin, height - 1.2 * cm, "Sistem Validasi RPS - Telkom University")
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawString(doc.leftMargin, 0.8 * cm, f"Generated: {generated_at}")
        canvas.drawRightString(width - doc.rightMargin, 0.8 * cm, f"Halaman {doc_obj.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return output_path