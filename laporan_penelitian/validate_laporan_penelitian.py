#!/usr/bin/env python3
"""
Sistem Audit Laporan Penelitian PDF
Menggabungkan ekstraksi Daftar Isi (TOC) yang presisi dengan 
aturan dinamis berdasarkan warna dan Parent dari RULES.xlsx.
"""

import argparse
import json
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import fitz  # PyMuPDF
from openpyxl import load_workbook

# Import PDF generator.
# Mendukung dua mode pemakaian:
#  - sebagai bagian package, dipanggil dari main.py (laporan_penelitian.validate_laporan_penelitian)
#  - dijalankan langsung sebagai script di dalam folder laporan_penelitian
try:
    from .PDF.pdfgenerator import generate_laporan_penelitian_pdf
except ImportError:
    from PDF.pdfgenerator import generate_laporan_penelitian_pdf


# =========================================================
# Basic project helpers
# =========================================================

#INI FUNGSI/BAGIAN UNTUK (mengambil folder root project)
def project_root() -> Path:
    return Path(__file__).resolve().parent

#INI FUNGSI/BAGIAN UNTUK (memastikan folder kerja sudah dibuat)
def ensure_dirs(root: Path) -> None:
    for name in ["config", "debug", "input", "output", "PDF"]:
        (root / name).mkdir(parents=True, exist_ok=True)

#INI FUNGSI/BAGIAN UNTUK (merapikan spasi pada teks)
def clean_space(text: Any) -> str:
    if text is None:
        return ""
    text = str(text).replace("\u00a0", " ")
    text = text.replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", text).strip()

#INI FUNGSI/BAGIAN UNTUK (membuang nomor di depan teks (mis. '1.2 Judul'))
def strip_number_prefix(text: str) -> str:
    return re.sub(r"^\s*\d+(?:\.\d+)*\s*", "", text).strip()

#INI FUNGSI/BAGIAN UNTUK (menormalkan teks (huruf kecil, rapikan spasi))
def normalize(text: Any) -> str:
    text = clean_space(strip_number_prefix(str(text or ""))).casefold()
    text = re.sub(r"[^a-z0-9À-ÿ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()

#INI FUNGSI/BAGIAN UNTUK (memecah teks menjadi token kata)
def tokens(text: Any) -> List[str]:
    return [tok for tok in normalize(text).split() if len(tok) > 1]

#INI FUNGSI/BAGIAN UNTUK (menghitung kemiripan dua teks)
def similarity(a: Any, b: Any) -> float:
    na, nb = normalize(a), normalize(b)
    if not na or not nb: return 0.0
    if na == nb: return 1.0

    if na in nb or nb in na:
        short_len = min(len(na), len(nb))
        long_len = max(len(na), len(nb))
        ratio = short_len / long_len if long_len else 0
        if short_len >= 5 and ratio >= 0.45:
            return 0.80 + (0.15 * ratio)
        return 0.70

    ta, tb = set(tokens(a)), set(tokens(b))
    token_score = len(ta & tb) / len(ta | tb) if ta and tb else 0.0
    seq_score = SequenceMatcher(None, na, nb).ratio()

    last_bonus = 0.0
    tok_a, tok_b = tokens(a), tokens(b)
    if tok_a and tok_b and tok_a[-1] == tok_b[-1]:
        last_bonus = 0.05

    return min(1.0, max(seq_score, token_score) + last_bonus)

#INI FUNGSI/BAGIAN UNTUK (mengecek apakah dua teks dianggap cocok)
def is_match(expected: Any, actual: Any, parent_context: bool = False) -> bool:
    score = similarity(expected, actual)
    if score >= 0.80: return True
    if parent_context and score >= 0.72: return True
    return False

# =========================================================
# Requirement reading from RULES.xlsx (kolom Requirement: Mandatory / Optional)
# =========================================================

#INI FUNGSI/BAGIAN UNTUK (menentukan jenis kewajiban dari kolom Requirement)
def classify_requirement(value: Any) -> str:
    text = normalize(value)
    if text in {"optional", "opsional", "opt", "tidak wajib", "tidakwajib"}:
        return "optional"
    if text in {"mandatory", "wajib", "required", "mandatori", "harus"}:
        return "required"
    # Default: mandatory (bila sel Requirement kosong/tidak dikenal).
    return "required"

#INI FUNGSI/BAGIAN UNTUK (membaca rules dari file)
def read_rules(rules_path: Path) -> List[Dict[str, Any]]:
    wb = load_workbook(rules_path, data_only=True)
    ws = wb.active

    header_row = field_col = parent_col = requirement_col = None
    for row in range(1, min(ws.max_row, 30) + 1):
        for col in range(1, ws.max_column + 1):
            value = normalize(ws.cell(row, col).value)
            if value == "field":
                header_row, field_col = row, col
            elif value == "parent":
                parent_col = col
            elif value in ("requirement", "requirements", "persyaratan", "kebutuhan"):
                requirement_col = col
        if header_row and field_col: break

    if not header_row or not field_col: raise ValueError("Kolom FIELD tidak ditemukan")
    if not parent_col: raise ValueError("Kolom Parent tidak ditemukan")
    if not requirement_col: raise ValueError("Kolom Requirement tidak ditemukan")

    rules = []
    for row in range(header_row + 1, ws.max_row + 1):
        field = clean_space(ws.cell(row, field_col).value)
        parent = clean_space(ws.cell(row, parent_col).value)
        if not field: continue
        requirement = classify_requirement(ws.cell(row, requirement_col).value)
        rules.append({
            "row": row, "field": field, "parent": parent or None,
            "required": requirement == "required",
            "requirement": "WAJIB" if requirement == "required" else "OPSIONAL",
        })
    return rules

# =========================================================
# PDF extraction & TOC Logic (DARI KODINGAN ANDA YANG RAPI)
# =========================================================

#INI FUNGSI/BAGIAN UNTUK (mengekstrak teks tiap halaman PDF)
def extract_pdf_pages(pdf_path: Path) -> Tuple[List[Dict[str, Any]], Dict[int, str], int]:
    doc = fitz.open(str(pdf_path))
    pages_list = []
    pages_dict = {}
    for i, page in enumerate(doc, start=1):
        text = page.get_text("text") or ""
        pages_list.append({"page": i, "text": text, "normalized": normalize(text)})
        pages_dict[i] = text
    count = len(doc)
    doc.close()
    return pages_list, pages_dict, count

#INI FUNGSI/BAGIAN UNTUK (mencari kemunculan sebuah kata kunci)
def find_occurrences(pages_dict, keyword):
    keyword = keyword.upper()
    return [page_num for page_num, text in pages_dict.items() if keyword in text.upper()]

#INI FUNGSI/BAGIAN UNTUK (menentukan rentang halaman Daftar Isi)
def find_toc_range(pages_dict, rules):
    rule_fields = [str(r["field"]).upper() for r in rules]
    toc_pages = find_occurrences(pages_dict, "DAFTAR ISI")
    
    if not toc_pages: 
        return None
        
    start_page = toc_pages[0]
    
    try:
        # Cari next section setelah DAFTAR ISI di rules
        if "DAFTAR ISI" in rule_fields:
            toc_idx = rule_fields.index("DAFTAR ISI")
            next_section = rule_fields[toc_idx + 1]
            next_hits = find_occurrences(pages_dict, next_section)
            real_next = next((p for p in next_hits if p > start_page), None)
            end_page = real_next - 1 if real_next else len(pages_dict)
        else:
            end_page = len(pages_dict)
    except (ValueError, IndexError):
        end_page = len(pages_dict)
        
    return start_page, end_page

#INI FUNGSI/BAGIAN UNTUK (menggabungkan baris-baris Daftar Isi yang terpisah)
def merge_toc_lines(lines):
    merged, buffer = [], ""
    for line in lines:
        line = line.strip()
        if not line or line.upper() == "DAFTAR ISI": continue
        if re.match(r"^(?:\.{2,}\s*)?([IVXLCDMivxlcdm]+|\d+)$", line):
            if buffer: merged.append(buffer + " " + line); buffer = ""
            continue
        is_new_start = re.match(r"^(BAB\s+[IVX]+|\d+(\.\d+)*\.?|[A-Z]\.)\s+", line, re.IGNORECASE)
        is_end_of_entry = re.search(r"\.{2,}\s*([IVXLCDMivxlcdm]+|\d+)?$", line)
        if is_new_start:
            if buffer: merged.append(buffer)
            buffer = line
        else:
            buffer = buffer + " " + line if buffer else line
        if is_end_of_entry:
            merged.append(buffer); buffer = ""
    if buffer: merged.append(buffer)
    return merged

#INI FUNGSI/BAGIAN UNTUK (mem-parsing satu baris entri Daftar Isi)
def parse_entry(line):
    line = re.sub(r"\s+", " ", line).strip()
    if not line: return None
    match_dots = re.search(r"^(.*?)\.{2,}\s*([IVXLCDMivxlcdm]+|\d+)?$", line, re.IGNORECASE)
    if match_dots: return {"title": match_dots.group(1).strip(), "page": match_dots.group(2).strip() if match_dots.group(2) else "Tidak Ada"}
    match_spaces = re.search(r"^(.*?)\s+([IVXLCDMivxlcdm]+|\d+)$", line)
    if match_spaces:
        title = match_spaces.group(1).strip()
        if title.upper() in ["BAB", "BAGIAN", "CHAPTER"]: return {"title": line, "page": "Tidak Ada"}
        return {"title": title, "page": match_spaces.group(2).strip()}
    return {"title": line, "page": "Tidak Ada"}

#INI FUNGSI/BAGIAN UNTUK (mem-parsing seluruh entri Daftar Isi)
def parse_toc_entries_hybrid(pages_dict: Dict[int, str], rules: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], bool, List[int]]:
    result = find_toc_range(pages_dict, rules)
    if not result:
        return [], False, []
        
    start_page, end_page = result
    toc_text = "\n".join([pages_dict[p] for p in range(start_page, end_page + 1)])
    merged_lines = merge_toc_lines(toc_text.splitlines())
    
    entries = []
    stack = {}
    
    for line in merged_lines:
        parsed = parse_entry(line)
        if not parsed: continue
        
        raw_title = parsed["title"]
        page_ref = parsed["page"] if parsed["page"] != "Tidak Ada" else None
        
        # Pisahkan Nomor dan Judul Bab secara cerdas
        bab_match = re.match(r"^BAB\s+([IVXLCDMivxlcdm\d]+)(?:\.)?\s+(.+)$", raw_title, re.IGNORECASE)
        if bab_match:
            number = bab_match.group(1).strip()
            title = bab_match.group(2).strip()
            level = 1
        else:
            num_match = re.match(r"^(\d+(?:\.\d+)*)\.?\s*(.*)$", raw_title)
            if num_match:
                number = num_match.group(1).strip()
                title = num_match.group(2).strip()
                level = len(number.split("."))
            else:
                number = None
                title = raw_title
                level = 1
                
        if not title: title = raw_title
        parent = stack.get(level - 1) if level > 1 else None
        
        item = {
            "index": len(entries),
            "number": number,
            "title": title,
            "raw_title": raw_title,
            "level": level,
            "parent_number": parent.get("number") if parent else None,
            "parent_title": parent.get("title") if parent else None,
            "toc_page": start_page,
            "page_ref": page_ref,
        }
        entries.append(item)
        stack[level] = item
        
        for k in list(stack.keys()):
            if k > level: del stack[k]
                
    return entries, True, list(range(start_page, end_page + 1))


# =========================================================
# Matching and validation (Logika Buku Kurikulum Baru)
# =========================================================

#INI FUNGSI/BAGIAN UNTUK (mencari kandidat paling cocok)
def best_match(candidates: Iterable[Dict[str, Any]], expected: str, parent_context: bool = False) -> Optional[Dict[str, Any]]:
    best = best_score = None
    for cand in candidates:
        score = similarity(expected, cand.get("title"))
        if best_score is None or score > best_score:
            best_score = score
            best = cand

    if best and is_match(expected, best.get("title"), parent_context=parent_context):
        copied = dict(best)
        copied["match_score"] = round(best_score, 3)
        return copied
    return None

#INI FUNGSI/BAGIAN UNTUK (mencari entri yang cocok)
def find_any_entry(entries: List[Dict[str, Any]], expected: str) -> Optional[Dict[str, Any]]:
    return best_match(entries, expected, parent_context=False)

#INI FUNGSI/BAGIAN UNTUK (mencari entri yang cocok, utamakan level 1)
def find_any_entry_prefer_level1(entries: List[Dict[str, Any]], expected: str) -> Optional[Dict[str, Any]]:
    level1 = [entry for entry in entries if entry.get("level") == 1]
    matched = best_match(level1, expected, parent_context=False)
    if matched: return matched
    return best_match(entries, expected, parent_context=False)

#INI FUNGSI/BAGIAN UNTUK (mencari entri yang cocok secara ketat)
def find_any_entry_strict(entries: List[Dict[str, Any]], expected: str) -> Optional[Dict[str, Any]]:
    matched = best_match(entries, expected, parent_context=False)
    if matched and float(matched.get("match_score") or 0) >= 0.90: return matched
    return None

#INI FUNGSI/BAGIAN UNTUK (mencari sub-entri (anak) dari sebuah entri)
def find_children(entries: List[Dict[str, Any]], parent_entry: Dict[str, Any]) -> List[Dict[str, Any]]:
    parent_number = parent_entry.get("number")
    parent_level = parent_entry.get("level")
    return [e for e in entries if e.get("parent_number") == parent_number and e.get("level") == parent_level + 1]

#INI FUNGSI/BAGIAN UNTUK (mengecek apakah halaman memuat sebuah field)
def page_contains_field(page_text: str, field: str) -> bool:
    n_page, n_field = normalize(page_text), normalize(field)
    if not n_field: return False
    if n_field in n_page: return True
    return similarity(field, page_text) >= 0.90

#INI FUNGSI/BAGIAN UNTUK (mencari field di seluruh teks dokumen)
def find_in_full_text(field: str, pages: List[Dict[str, Any]]) -> Dict[str, Any]:
    for page in pages:
        if page_contains_field(page.get("text", ""), field):
            return {"found": True, "matched_title": field, "page": page.get("page"), "source": "full_text"}
    return {"found": False, "matched_title": None, "page": None, "source": "full_text"}

#INI FUNGSI/BAGIAN UNTUK (mengambil indeks aturan root pertama)
def first_root_rule_index(rules: List[Dict[str, Any]]) -> Optional[int]:
    for idx, rule in enumerate(rules):
        if not rule.get("parent"): return idx
    return None

#INI FUNGSI/BAGIAN UNTUK (memvalidasi seluruh rules terhadap dokumen)
def validate_rules(rules, entries, pages, page_count):
    results = []
    matched_entry_indices = set()
    rule_match_by_field = {}
    first_root_idx = first_root_rule_index(rules)

    for rule_idx, rule in enumerate(rules):
        field, parent = rule["field"], rule.get("parent")
        base_result = {
            **rule, "found": False, "parent_ok": True,
            "status": "BELUM TERPENUHI" if rule.get("required") else "OPSIONAL TIDAK ADA",
            "matched_title": None, "matched_number": None, "matched_parent": None,
            "toc_page": None, "page": None, "match_score": None,
            "source": "daftar_isi" if parent else "daftar_isi_or_full_text",
        }

        matched = None
        parent_ok = True

        if parent:
            parent_entry = rule_match_by_field.get(normalize(parent)) or find_any_entry(entries, parent)
            if parent_entry:
                children = find_children(entries, parent_entry)
                matched = best_match(children, field, parent_context=True)
            if not matched:
                any_match = find_any_entry_strict(entries, field)
                if any_match:
                    matched, parent_ok = any_match, False
        else:
            matched = find_any_entry_prefer_level1(entries, field)
            if not matched and rule_idx == first_root_idx and page_count > 0:
                base_result.update({"found": True, "matched_title": "Halaman pertama dokumen", "page": 1, "source": "first_page_by_rule_order"})
                base_result["status"] = "TERPENUHI" if rule.get("required") else "OPSIONAL ADA"
                results.append(base_result)
                rule_match_by_field[normalize(field)] = None
                continue

            if not matched:
                text_match = find_in_full_text(field, pages)
                if text_match.get("found"):
                    base_result.update(text_match)
                    base_result["status"] = "TERPENUHI" if rule.get("required") else "OPSIONAL ADA"
                    results.append(base_result)
                    rule_match_by_field[normalize(field)] = None
                    continue

        if matched:
            base_result.update({
                "found": True, "parent_ok": parent_ok,
                "matched_title": matched.get("title"), "matched_number": matched.get("number"),
                "matched_parent": matched.get("parent_title"), "toc_page": matched.get("toc_page"),
                "page": matched.get("page_ref"), "match_score": matched.get("match_score"),
                "source": "daftar_isi",
            })
            matched_entry_indices.add(matched.get("index"))
            rule_match_by_field[normalize(field)] = matched
            base_result["status"] = "TERPENUHI" if rule.get("required") else "OPSIONAL ADA" if parent_ok else "PARENT BERBEDA"
        else:
            rule_match_by_field[normalize(field)] = None

        results.append(base_result)

    extras = [
        {"number": e.get("number"), "title": e.get("title"), "level": e.get("level"), "parent_number": e.get("parent_number"), "parent_title": e.get("parent_title"), "toc_page": e.get("toc_page"), "page_ref": e.get("page_ref"), "status": "TAMBAHAN"}
        for e in entries if e.get("index") not in matched_entry_indices
    ]
    return results, extras

#INI FUNGSI/BAGIAN UNTUK (menyimpulkan keberadaan bagian root dari posisi)
def apply_root_positional_inference(results: List[Dict[str, Any]], page_count: int) -> None:
    root_positions = [idx for idx, item in enumerate(results) if not item.get("parent")]
    if not root_positions: return

    #INI FUNGSI/BAGIAN UNTUK (halaman nomor)
    def page_number(item):
        try: return int(item.get("page") or item.get("page_ref") or item.get("toc_page"))
        except: return None

    for pos_idx, result_idx in enumerate(root_positions):
        if results[result_idx].get("found"): continue

        prev_pos_idx = next_pos_idx = None
        for i in range(pos_idx - 1, -1, -1):
            if results[root_positions[i]].get("found") and page_number(results[root_positions[i]]) is not None:
                prev_pos_idx = i; break
        for i in range(pos_idx + 1, len(root_positions)):
            if results[root_positions[i]].get("found") and page_number(results[root_positions[i]]) is not None:
                next_pos_idx = i; break

        if prev_pos_idx is None or next_pos_idx is None: continue

        prev_page = page_number(results[root_positions[prev_pos_idx]])
        next_page = page_number(results[root_positions[next_pos_idx]])
        
        if prev_page is None or next_page is None or next_page <= prev_page: continue

        missing_between = [root_positions[i] for i in range(prev_pos_idx + 1, next_pos_idx) if not results[root_positions[i]].get("found")]
        available_pages = list(range(prev_page + 1, next_page))
        if len(available_pages) < len(missing_between): continue

        for idx, inferred_page in zip(missing_between, available_pages):
            inferred = results[idx]
            if inferred.get("found") or inferred_page < 1 or inferred_page > page_count: continue
            inferred.update({
                "found": True, "parent_ok": True, "matched_title": "Diduga berdasarkan urutan rules",
                "page": inferred_page, "source": "posisi_berdasarkan_urutan_rules", "match_score": None,
            })
            inferred["status"] = "TERPENUHI" if inferred.get("required") else "OPSIONAL ADA"

#INI FUNGSI/BAGIAN UNTUK (menyusun hasil validasi akhir)
def build_result(pdf_path: Path, rules_path: Path, debug_dir: Path, output_pdf: Path) -> Dict[str, Any]:
    rules = read_rules(rules_path)
    pages_list, pages_dict, page_count = extract_pdf_pages(pdf_path)
    
    # Memanggil Eksekutor Daftar Isi Hybrid Kita!
    toc_entries, toc_found, toc_pages = parse_toc_entries_hybrid(pages_dict, rules)

    if not toc_found:
        validation_results = []
        extra_sections = []
        for rule in rules:
            validation_results.append({
                **rule, "found": False, "parent_ok": True, "status": "DAFTAR ISI TIDAK DITEMUKAN",
                "matched_title": None, "matched_number": None, "matched_parent": None,
                "toc_page": None, "page": None, "match_score": None, "source": "daftar_isi",
            })
    else:
        validation_results, extra_sections = validate_rules(rules, toc_entries, pages_list, page_count)

    apply_root_positional_inference(validation_results, page_count)

    result = {
        "file": str(pdf_path),
        "rules_path": str(rules_path),
        "output_pdf": str(output_pdf),
        "page_count": page_count,
        "toc_found": toc_found,
        "toc_pages": toc_pages,
        "rules": rules,
        "toc_entries": toc_entries,
        "validation_results": validation_results,
        "extra_sections": extra_sections,
    }

    debug_dir.mkdir(parents=True, exist_ok=True)
    (debug_dir / "extracted_toc.json").write_text(json.dumps(toc_entries, ensure_ascii=False, indent=2), encoding="utf-8")
    (debug_dir / "rules_loaded.json").write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
    (debug_dir / "validation_result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result

# =========================================================
# CLI
# =========================================================

#INI FUNGSI/BAGIAN UNTUK (membaca argumen command line)
def parse_args() -> argparse.Namespace:
    root = project_root()
    parser = argparse.ArgumentParser(description="Audit kelengkapan Laporan Penelitian berdasarkan RULES.xlsx")
    parser.add_argument("--pdf", default=str(root / "input" / "input.pdf"), help="Path PDF input")
    parser.add_argument("--rules", default=str(root / "config" / "RULES.xlsx"), help="Path RULES.xlsx")
    parser.add_argument("--output", "-o", default=str(root / "output" / "laporan.pdf"), help="Path laporan PDF")
    parser.add_argument("--debug", default=str(root / "debug"), help="Folder debug JSON")
    return parser.parse_args()

#INI FUNGSI/BAGIAN UNTUK (entry point untuk menjalankan dari command line)
def main() -> int:
    args = parse_args()
    root = project_root()
    ensure_dirs(root)

    pdf_path = Path(args.pdf)
    rules_path = Path(args.rules)
    output_path = Path(args.output)
    debug_dir = Path(args.debug)

    if not pdf_path.exists():
        print(f"PDF input tidak ditemukan: {pdf_path}", file=sys.stderr)
        return 1
    if not rules_path.exists():
        print(f"Rules tidak ditemukan: {rules_path}", file=sys.stderr)
        return 1

    print("Membaca PDF dan mengekstrak aturan...")
    result = build_result(pdf_path, rules_path, debug_dir, output_path)
    
    print("Menghasilkan laporan PDF Laporan Penelitian...")
    generate_laporan_penelitian_pdf(result, str(output_path))

    print("\n✅ Audit selesai.")
    print(f"Daftar Isi ditemukan: {'Ya' if result.get('toc_found') else 'Tidak'}")
    print(f"Total rules: {len(result.get('validation_results', []))}")

    missing_required = [item for item in result.get("validation_results", []) if item.get("required") and not item.get("found")]
    parent_mismatch = [item for item in result.get("validation_results", []) if item.get("required") and item.get("found") and not item.get("parent_ok", True)]

    print(f"Wajib belum ditemukan: {len(missing_required)}")
    print(f"Wajib parent berbeda: {len(parent_mismatch)}")
    print(f"Header/sub-header tambahan: {len(result.get('extra_sections', []))}")
    print(f"Laporan berhasil dibuat di: {output_path}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())