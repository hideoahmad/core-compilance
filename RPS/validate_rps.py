

import argparse
import json
import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from RPS.PDF.pdfgenerator import generate_audit_pdf


# =============================================================================
# Path Configuration
# =============================================================================

#INI FUNGSI/BAGIAN UNTUK (mengambil folder root project)
def get_project_root() -> Path:
    return Path(__file__).parent

#INI FUNGSI/BAGIAN UNTUK (mengambil daftar subfolder kerja RPS)
def get_rps_subfolders():
    root = get_project_root()
    return {
        "config": root / "Config",
        "input": root / "INPUT",
        "output": root / "OUTPUT",
        "debug": root / "DEBUG",
        "pdf": root / "PDF",
    }

#INI FUNGSI/BAGIAN UNTUK (menentukan path file input)
def resolve_input_path(filename: str) -> Path:
    subfolders = get_rps_subfolders()
    
    input_file = subfolders["input"] / filename
    if input_file.exists():
        return input_file
    
    config_file = subfolders["config"] / filename
    if config_file.exists():
        return config_file
    return Path(filename)

#INI FUNGSI/BAGIAN UNTUK (menentukan path file config/rules)
def resolve_config_path(filename: str) -> Optional[Path]:
    subfolders = get_rps_subfolders()
    config_dir = subfolders["config"]

    config_file = config_dir / filename
    if config_file.exists():
        return config_file
    
    cwd_file = Path(filename)
    if cwd_file.exists():
        return cwd_file
    
    return None

#INI FUNGSI/BAGIAN UNTUK (menentukan path file output)
def resolve_output_path(filename: str, create_dir: bool = True) -> Path:

    subfolders = get_rps_subfolders()
    
    file_ext = Path(filename).suffix.lower()
    
    if file_ext == ".json":
        output_dir = subfolders["debug"]
    else:
        output_dir = subfolders["output"]
    
    if create_dir:
        output_dir.mkdir(parents=True, exist_ok=True)
    
    return output_dir / filename


#INI FUNGSI/BAGIAN UNTUK (data hasil pencocokan sebuah sel)
@dataclass
class CellHit:
    row: int
    col: int
    value: Any
    text: str
    location: str
    score: float
    bounds: Tuple[int, int, int, int]


#INI FUNGSI/BAGIAN UNTUK (data hasil pencocokan sebuah header)
@dataclass
class HeaderHit:
    header: str
    row: int
    col: int
    value: Any
    text: str
    location: str
    score: float
    bounds: Tuple[int, int, int, int]


#INI FUNGSI/BAGIAN UNTUK (data kandidat section pada sheet)
@dataclass
class SectionCandidate:
    category: str
    category_index: int
    has_header_rules: bool
    detection_method: str
    category_label_found: bool
    category_inferred_from_headers: bool
    category_is_explicit: bool
    category_anchor: Optional[CellHit]
    inferred_headers: Dict[str, HeaderHit]
    raw_row_min: int
    raw_row_max: int
    raw_col_min: int
    raw_col_max: int
    row_min: int = 0
    row_max: int = 0
    col_min: int = 0
    col_max: int = 0


#INI FUNGSI/BAGIAN UNTUK (menormalkan teks untuk pembandingan)
def normalize_text(value: Any) -> str:
    if value is None:
        return ""

    text = str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()

    text = text.replace("&", " dan ")
    text = text.replace("/", " ")
    text = text.replace("\\", " ")

    text = re.sub(r"[\n\r\t]+", " ", text)
    text = re.sub(r"[^a-z0-9%]+", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip()


#INI FUNGSI/BAGIAN UNTUK (membersihkan teks pada output)
def clean_output_text(value: Any) -> Any:
    if isinstance(value, str):
        value = value.replace("\r\n", " ")
        value = value.replace("\n", " ")
        value = value.replace("\r", " ")
        value = re.sub(r"\s+", " ", value)
        return value.strip()

    return value


#INI FUNGSI/BAGIAN UNTUK (membersihkan seluruh teks output secara rekursif)
def clean_output_data(data: Any) -> Any:
    if isinstance(data, dict):
        return {
            key: clean_output_data(value)
            for key, value in data.items()
        }

    if isinstance(data, list):
        return [
            clean_output_data(item)
            for item in data
        ]

    return clean_output_text(data)


#INI FUNGSI/BAGIAN UNTUK (mengecek apakah sebuah nilai kosong)
def is_empty(value: Any) -> bool:
    return value is None or str(value).strip() == ""


#INI FUNGSI/BAGIAN UNTUK (menyusun lokasi sel (mis. B6) dari baris dan kolom)
def cell_location(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


#INI FUNGSI/BAGIAN UNTUK (menyusun lokasi area dari batas baris/kolom)
def boundary_location(row_min: int, col_min: int, row_max: int, col_max: int) -> str:
    return f"{cell_location(row_min, col_min)}:{cell_location(row_max, col_max)}"


#INI FUNGSI/BAGIAN UNTUK (mengecek apakah sebuah area valid)
def is_valid_region(region: Optional[Tuple[int, int, int, int]]) -> bool:
    if region is None:
        return False

    row_min, col_min, row_max, col_max = region
    return row_min <= row_max and col_min <= col_max


#INI FUNGSI/BAGIAN UNTUK (mengubah batas area menjadi dictionary)
def bounds_to_dict(bounds: Optional[Tuple[int, int, int, int]]) -> Optional[Dict[str, Any]]:
    if not is_valid_region(bounds):
        return None

    row_min, col_min, row_max, col_max = bounds

    return {
        "row_min": row_min,
        "row_max": row_max,
        "col_min": col_min,
        "col_max": col_max,
        "range": boundary_location(row_min, col_min, row_max, col_max),
    }


#INI FUNGSI/BAGIAN UNTUK (mengecek apakah dua area saling beririsan)
def bounds_intersect(
    a: Tuple[int, int, int, int],
    b: Tuple[int, int, int, int],
) -> bool:
    a_row_min, a_col_min, a_row_max, a_col_max = a
    b_row_min, b_col_min, b_row_max, b_col_max = b

    return not (
        a_row_max < b_row_min
        or b_row_max < a_row_min
        or a_col_max < b_col_min
        or b_col_max < a_col_min
    )


#INI FUNGSI/BAGIAN UNTUK (menghitung kemiripan antara dua teks)
def match_score(source: Any, target: Any) -> float:
    source_norm = normalize_text(source)
    target_norm = normalize_text(target)

    if not source_norm or not target_norm:
        return 0.0

    if source_norm == target_norm:
        return 1.0

    if target_norm in source_norm:
        return 0.96

    if (
        source_norm in target_norm
        and len(source_norm) >= 4
        and not source_norm.isdigit()
    ):
        return 0.90

    return SequenceMatcher(None, source_norm, target_norm).ratio()


#INI FUNGSI/BAGIAN UNTUK (menormalkan kode warna isi sel)
def normalize_fill_rgb(cell) -> str:
    fg = cell.fill.fgColor

    if fg is None:
        return ""

    if fg.type == "rgb" and fg.rgb:
        return str(fg.rgb).upper()

    if fg.type == "indexed" and fg.indexed is not None:
        return str(fg.indexed)

    return ""


#INI FUNGSI/BAGIAN UNTUK (mengecek apakah sel berwarna hijau)
def is_green_fill(cell) -> bool:
    rgb = normalize_fill_rgb(cell)

    if not rgb:
        return False
    
    if rgb.endswith("00B050"):
        return True

    if rgb.endswith("92D050"):
        return True

    if rgb.endswith("00FF00"):
        return True

    return False


#INI FUNGSI/BAGIAN UNTUK (mengubah teks menjadi nilai boolean)
def parse_bool(value: Any, default: bool = True) -> bool:
    if value is None or str(value).strip() == "":
        return default

    text = normalize_text(value)

    if text in {"true", "1", "yes", "ya", "y", "wajib"}:
        return True

    if text in {"false", "0", "no", "tidak", "n", "opsional", "optional"}:
        return False

    return default


#INI FUNGSI/BAGIAN UNTUK (membaca kolom requirement: mandatory / optional)
def parse_requirement(value: Any, default: str = "mandatory") -> str:

    if value is None or str(value).strip() == "":
        return default

    text = normalize_text(value)

    if text in {"optional", "opsional", "opt", "tidak wajib", "tidakwajib"}:
        return "optional"

    if text in {"mandatory", "wajib", "required", "mandatori", "harus"}:
        return "mandatory"

    return default


#INI FUNGSI/BAGIAN UNTUK (menurunkan required & valueRequired dari requirement)
def requirement_to_flags(requirement: str) -> Tuple[bool, bool]:

    if requirement == "optional":
        return True, False
    return True, True


#INI FUNGSI/BAGIAN UNTUK (membaca kolom compare: yes berarti ikut dibandingkan)
def parse_compare_flag(value: Any) -> bool:

    return parse_bool(value, default=False)


#INI FUNGSI/BAGIAN UNTUK (mencari indeks kolom berdasarkan nama header)
def find_column_index(header_map: Dict[str, int], *names: str) -> Optional[int]:
    normalized_lookup = {normalize_text(name): col for name, col in header_map.items()}

    for name in names:
        key = normalize_text(name)

        if key in normalized_lookup:
            return normalized_lookup[key]

    return None


#INI FUNGSI/BAGIAN UNTUK (memuat dan mem-parsing rules dari file .xlsx)
def load_rules_xlsx(path: str) -> Dict[str, Any]:
    workbook = load_workbook(path, data_only=True)

    if "rules" in workbook.sheetnames:
        ws = workbook["rules"]
    else:
        ws = workbook.active

    header_map = {}

    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col).value

        if value is None:
            continue

        header_map[str(value).strip()] = col

    category_col = find_column_index(header_map, "category", "kategori") or 1

    field_col = find_column_index(header_map, "field", "fields", "header", "headers") or 2
    section_col = find_column_index(header_map, "section", "bagian")


    requirement_col = find_column_index(header_map, "requirement", "requirements", "persyaratan", "kebutuhan")
    compare_col = find_column_index(header_map, "compare", "bandingkan", "dibandingkan", "perbandingan")

    legacy_required_col = find_column_index(header_map, "required", "header required", "wajib header")
    legacy_value_required_col = find_column_index(header_map, "value required", "value_required", "valueRequired", "wajib value")

    categories = []
    category_lookup: Dict[str, Dict[str, Any]] = {}
    current_category = None
    current_section = None
    compare_markers = []

    for row in range(2, ws.max_row + 1):
        category_value = ws.cell(row=row, column=category_col).value
        section_value = ws.cell(row=row, column=section_col).value if section_col else None
        header_cell = ws.cell(row=row, column=field_col)
        header_value = header_cell.value

        if category_value is not None and str(category_value).strip() != "":
            current_category = str(category_value).strip()

        if section_value is not None and str(section_value).strip() != "":
            current_section = str(section_value).strip()

        if current_category is None and current_section is None and is_empty(header_value):
            continue

        category_name = current_category or current_section or f"Header Group {len(categories) + 1}"
        category_is_explicit = current_category is not None

        if requirement_col:
            row_requirement = parse_requirement(
                ws.cell(row=row, column=requirement_col).value, default="mandatory"
            )
            row_required, row_value_required = requirement_to_flags(row_requirement)
        else:
            row_required = (
                parse_bool(ws.cell(row=row, column=legacy_required_col).value, True)
                if legacy_required_col
                else True
            )
            row_value_required = (
                parse_bool(ws.cell(row=row, column=legacy_value_required_col).value, True)
                if legacy_value_required_col
                else True
            )
            row_requirement = "mandatory" if row_value_required else "optional"

        if category_name not in category_lookup:
            category_item = {
                "category": category_name,
                "section": current_section,
                "category_is_explicit": category_is_explicit,
                "requirement": row_requirement,
                "required": row_required,
                "valueRequired": row_value_required,
                "headers": [],
            }
            category_lookup[category_name] = category_item
            categories.append(category_item)

        if is_empty(header_value):
            continue

        field_name = str(header_value).strip()

        header_rule = {
            "name": field_name,
            "required": row_required,
            "valueRequired": row_value_required,
            "requirement": row_requirement,
        }

        category_lookup[category_name]["headers"].append(header_rule)

        if compare_col:
            is_compared = parse_compare_flag(ws.cell(row=row, column=compare_col).value)
            compare_source = "compare column"
        else:
            is_compared = is_green_fill(header_cell)
            compare_source = "legacy green fill"

        if is_compared:
            compare_markers.append(
                {
                    "category": category_name,
                    "header": field_name,
                    "row": row,
                    "column": field_col,
                    "cell": cell_location(row, field_col),
                    "source": compare_source,
                }
            )

    alignment_config = {
        "enabled": len(compare_markers) >= 2,
        "source": "rules.xlsx kolom compare = yes",
        "markers": compare_markers,
        "description_header": compare_markers[0]["header"] if len(compare_markers) >= 1 else None,
        "method_header": compare_markers[1]["header"] if len(compare_markers) >= 2 else None,
        "category": compare_markers[0]["category"] if len(compare_markers) >= 1 else None,
        "note": "Alignment memakai 2 field pertama yang ditandai compare = yes pada rules.xlsx: field pertama sebagai deskripsi/capaian, field kedua sebagai metode pembelajaran.",
    }

    return {
        "categories": categories,
        "alignment_config": alignment_config,
    }


#INI FUNGSI/BAGIAN UNTUK (memuat rules dari file)
def load_rules(path: str) -> Dict[str, Any]:
    suffix = Path(path).suffix.lower()

    if suffix != ".xlsx":
        raise ValueError("Rules hanya menerima file .xlsx")

    if not Path(path).exists():
        raise FileNotFoundError(f"File rules tidak ditemukan: {path}")

    return load_rules_xlsx(path)


#INI FUNGSI/BAGIAN UNTUK (menyeragamkan struktur rules)
def normalize_rules_structure(rules: Dict[str, Any]) -> Dict[str, Any]:
    normalized_categories = []
    raw_categories = rules.get("categories", [])

    if not raw_categories and rules.get("headers"):
        raw_categories = [
            {
                "section": "Header Group 1",
                "headers": rules.get("headers", []),
            }
        ]

    for index, item in enumerate(raw_categories):
        headers = item.get("headers", [])

        category_name = (
            item.get("category")
            or item.get("section")
            or item.get("name")
            or f"Header Group {index + 1}"
        )

        if "category_is_explicit" in item:
            category_is_explicit = bool(item.get("category_is_explicit"))
        else:
            category_is_explicit = bool(item.get("category"))

        normalized_categories.append(
            {
                **item,
                "category": category_name,
                "category_is_explicit": category_is_explicit,
                "headers": headers,
            }
        )

    return {
        **rules,
        "categories": normalized_categories,
    }


#INI FUNGSI/BAGIAN UNTUK (memetakan sel-sel yang digabung (merge))
def get_merged_cell_maps(ws) -> Tuple[
    Dict[Tuple[int, int], Tuple[int, int]],
    Dict[Tuple[int, int], Tuple[int, int, int, int]],
]:
    merged_origin_map = {}
    merged_bounds_map = {}

    for merged_range in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        origin = (min_row, min_col)
        bounds = (min_row, min_col, max_row, max_col)

        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                merged_origin_map[(row, col)] = origin
                merged_bounds_map[(row, col)] = bounds

    return merged_origin_map, merged_bounds_map


#INI FUNGSI/BAGIAN UNTUK (mengambil informasi sebuah sel)
def get_cell_info(
    ws,
    row: int,
    col: int,
    merged_origin_map: Dict[Tuple[int, int], Tuple[int, int]],
    merged_bounds_map: Dict[Tuple[int, int], Tuple[int, int, int, int]],
) -> Dict[str, Any]:
    origin_row, origin_col = merged_origin_map.get((row, col), (row, col))
    value = ws.cell(row=origin_row, column=origin_col).value

    bounds = merged_bounds_map.get(
        (row, col),
        (row, col, row, col),
    )

    return {
        "value": value,
        "origin_row": origin_row,
        "origin_col": origin_col,
        "origin_location": cell_location(origin_row, origin_col),
        "bounds": bounds,
        "bounds_range": boundary_location(*bounds),
        "is_merged": bounds != (row, col, row, col),
    }


#INI FUNGSI/BAGIAN UNTUK (membangun representasi seluruh sel pada sheet)
def build_cells(
    ws,
    merged_origin_map: Dict[Tuple[int, int], Tuple[int, int]],
    merged_bounds_map: Dict[Tuple[int, int], Tuple[int, int, int, int]],
) -> List[Dict[str, Any]]:
    cells = []
    seen_origins = set()

    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            info = get_cell_info(ws, row, col, merged_origin_map, merged_bounds_map)
            origin = (info["origin_row"], info["origin_col"])

            if origin in seen_origins:
                continue

            seen_origins.add(origin)

            value = info["value"]

            if is_empty(value):
                continue

            cells.append(
                {
                    "row": info["origin_row"],
                    "col": info["origin_col"],
                    "value": value,
                    "text": str(value).strip(),
                    "norm": normalize_text(value),
                    "location": info["origin_location"],
                    "bounds": info["bounds"],
                    "bounds_range": info["bounds_range"],
                    "is_merged": info["is_merged"],
                }
            )

    return cells


#INI FUNGSI/BAGIAN UNTUK (menyusun daftar istilah dari rules)
def make_rule_terms(rules: Dict[str, Any]) -> List[str]:
    terms = []

    for category in rules.get("categories", []):
        category_name = category.get("category", "")

        if category.get("category_is_explicit") and category_name:
            terms.append(category_name)

        for header in category.get("headers", []):
            if isinstance(header, dict):
                header_name = header.get("name", "")
            else:
                header_name = str(header)

            if header_name:
                terms.append(header_name)

    return list(dict.fromkeys(terms))


#INI FUNGSI/BAGIAN UNTUK (menebak apakah teks adalah label aturan)
def is_probably_rule_label(value: Any, rule_terms: List[str], threshold: float = 0.92) -> bool:
    value_norm = normalize_text(value)

    if not value_norm:
        return True

    for term in rule_terms:
        term_norm = normalize_text(term)

        if not term_norm:
            continue

        if value_norm == term_norm:
            return True

        if SequenceMatcher(None, value_norm, term_norm).ratio() >= threshold:
            return True

    return False


#INI FUNGSI/BAGIAN UNTUK (mencari sel yang teksnya cocok)
def find_text_matches(
    cells: List[Dict[str, Any]],
    target: str,
    threshold: float,
    row_min: Optional[int] = None,
    row_max: Optional[int] = None,
    col_min: Optional[int] = None,
    col_max: Optional[int] = None,
) -> List[CellHit]:
    hits = []
    seen_origins = set()

    for cell in cells:
        row = cell["row"]
        col = cell["col"]
        origin = (row, col)

        if origin in seen_origins:
            continue

        seen_origins.add(origin)

        if row_min is not None and row < row_min:
            continue

        if row_max is not None and row > row_max:
            continue

        if col_min is not None and col < col_min:
            continue

        if col_max is not None and col > col_max:
            continue

        score = match_score(cell["text"], target)

        if score >= threshold:
            hits.append(
                CellHit(
                    row=row,
                    col=col,
                    value=cell["value"],
                    text=cell["text"],
                    location=cell["location"],
                    score=round(score, 4),
                    bounds=cell["bounds"],
                )
            )

    hits.sort(key=lambda item: (-item.score, item.row, item.col))
    return hits


#INI FUNGSI/BAGIAN UNTUK (mencari label kategori paling cocok)
def find_best_category_label(
    cells: List[Dict[str, Any]],
    category_name: str,
    threshold: float,
) -> Optional[CellHit]:
    matches = find_text_matches(cells, category_name, threshold)
    return matches[0] if matches else None


#INI FUNGSI/BAGIAN UNTUK (mengambil nama header dari aturan)
def parse_header_name(header_rule: Any) -> str:
    if isinstance(header_rule, dict):
        return str(header_rule.get("name", "")).strip()

    return str(header_rule).strip()


#INI FUNGSI/BAGIAN UNTUK (mengecek apakah header wajib ada)
def parse_header_required(header_rule: Any) -> bool:
    if isinstance(header_rule, dict):
        return bool(header_rule.get("required", True))

    return True


#INI FUNGSI/BAGIAN UNTUK (mengecek apakah value header wajib terisi)
def parse_value_required(header_rule: Any) -> bool:
    if isinstance(header_rule, dict):
        return bool(header_rule.get("valueRequired", True))

    return True


#INI FUNGSI/BAGIAN UNTUK (mengumpulkan kandidat posisi header)
def collect_header_candidates(
    cells: List[Dict[str, Any]],
    headers: List[Any],
    header_threshold: float,
    row_min: Optional[int] = None,
    row_max: Optional[int] = None,
    col_min: Optional[int] = None,
    col_max: Optional[int] = None,
) -> Dict[str, List[HeaderHit]]:
    result = {}

    for header_rule in headers:
        header_name = parse_header_name(header_rule)

        if not header_name:
            continue

        matches = find_text_matches(
            cells=cells,
            target=header_name,
            threshold=header_threshold,
            row_min=row_min,
            row_max=row_max,
            col_min=col_min,
            col_max=col_max,
        )

        result[header_name] = [
            HeaderHit(
                header=header_name,
                row=match.row,
                col=match.col,
                value=match.value,
                text=match.text,
                location=match.location,
                score=match.score,
                bounds=match.bounds,
            )
            for match in matches
        ]

    return result


#INI FUNGSI/BAGIAN UNTUK (menyimpulkan kelompok header yang berdekatan)
def infer_header_cluster(
    header_candidates: Dict[str, List[HeaderHit]],
    cluster_row_window: int,
) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int], Dict[str, HeaderHit]]:
    flattened = []

    for _, candidates in header_candidates.items():
        flattened.extend(candidates)

    if not flattened:
        return None, None, None, None, {}

    best_score = -1.0
    best_selected = {}

    for anchor in flattened:
        row_min = max(1, anchor.row - cluster_row_window)
        row_max = anchor.row + cluster_row_window

        selected = {}
        total_match_score = 0.0

        for header_name, candidates in header_candidates.items():
            inside = [
                candidate
                for candidate in candidates
                if row_min <= candidate.row <= row_max
            ]

            if not inside:
                continue

            inside.sort(
                key=lambda item: (
                    -item.score,
                    abs(item.row - anchor.row),
                    abs(item.col - anchor.col),
                    item.row,
                    item.col,
                )
            )

            selected[header_name] = inside[0]
            total_match_score += inside[0].score

        if not selected:
            continue

        selected_rows = []
        selected_cols = []

        for item in selected.values():
            row_min_b, col_min_b, row_max_b, col_max_b = item.bounds
            selected_rows.extend([row_min_b, row_max_b])
            selected_cols.extend([col_min_b, col_max_b])

        row_span = max(selected_rows) - min(selected_rows)
        col_span = max(selected_cols) - min(selected_cols)

        score = (
            len(selected) * 100
            + total_match_score * 10
            - row_span * 1.5
            - col_span * 0.05
        )

        if score > best_score:
            best_score = score
            best_selected = selected

    if not best_selected:
        return None, None, None, None, {}

    rows = []
    cols = []

    for hit in best_selected.values():
        row_min, col_min, row_max, col_max = hit.bounds
        rows.extend([row_min, row_max])
        cols.extend([col_min, col_max])

    return min(rows), max(rows), min(cols), max(cols), best_selected


#INI FUNGSI/BAGIAN UNTUK (memilih header terdekat dari anchor kategori)
def choose_header_hits_near_anchor(
    header_candidates: Dict[str, List[HeaderHit]],
    anchor: CellHit,
) -> Dict[str, HeaderHit]:
    selected = {}

    for header_name, candidates in header_candidates.items():
        if not candidates:
            continue

        candidates.sort(
            key=lambda item: (
                -item.score,
                abs(item.row - anchor.row),
                abs(item.col - anchor.col),
                item.row,
                item.col,
            )
        )

        selected[header_name] = candidates[0]

    return selected


#INI FUNGSI/BAGIAN UNTUK (mendeteksi kandidat section pada sheet)
def detect_section_candidates(
    ws,
    cells: List[Dict[str, Any]],
    rules: Dict[str, Any],
    category_threshold: float,
    header_threshold: float,
    explicit_category_search_window: int,
    inferred_cluster_row_window: int,
    value_lookahead: int,
) -> List[SectionCandidate]:
    candidates = []

    for index, category_rule in enumerate(rules.get("categories", [])):
        category_name = category_rule.get("category", "")
        headers = category_rule.get("headers", [])
        category_is_explicit = bool(category_rule.get("category_is_explicit"))

        if category_is_explicit:
            category_anchor = find_best_category_label(
                cells=cells,
                category_name=category_name,
                threshold=category_threshold,
            )
        else:
            category_anchor = None

        if not headers:
            if category_anchor is None:
                continue

            cat_row_min, cat_col_min, cat_row_max, cat_col_max = category_anchor.bounds

            candidates.append(
                SectionCandidate(
                    category=category_name,
                    category_index=index,
                    has_header_rules=False,
                    detection_method="category_label",
                    category_label_found=True,
                    category_inferred_from_headers=False,
                    category_is_explicit=category_is_explicit,
                    category_anchor=category_anchor,
                    inferred_headers={},
                    raw_row_min=cat_row_min,
                    raw_row_max=cat_row_max,
                    raw_col_min=cat_col_min,
                    raw_col_max=ws.max_column,
                )
            )

            continue

        if category_anchor:
            cat_row_min, cat_col_min, cat_row_max, cat_col_max = category_anchor.bounds

            provisional_row_min = cat_row_min
            provisional_row_max = min(ws.max_row, cat_row_max + explicit_category_search_window)

            header_candidates = collect_header_candidates(
                cells=cells,
                headers=headers,
                header_threshold=header_threshold,
                row_min=provisional_row_min,
                row_max=provisional_row_max,
            )

            selected_headers = choose_header_hits_near_anchor(
                header_candidates=header_candidates,
                anchor=category_anchor,
            )

            rows = [cat_row_min, cat_row_max]
            cols = [cat_col_min, cat_col_max]

            for hit in selected_headers.values():
                h_row_min, h_col_min, h_row_max, h_col_max = hit.bounds
                rows.extend([h_row_min, h_row_max])
                cols.extend([h_col_min, h_col_max])

            candidates.append(
                SectionCandidate(
                    category=category_name,
                    category_index=index,
                    has_header_rules=True,
                    detection_method="category_label",
                    category_label_found=True,
                    category_inferred_from_headers=False,
                    category_is_explicit=category_is_explicit,
                    category_anchor=category_anchor,
                    inferred_headers=selected_headers,
                    raw_row_min=min(rows),
                    raw_row_max=max(rows),
                    raw_col_min=min(cols),
                    raw_col_max=min(ws.max_column, max(cols) + value_lookahead),
                )
            )

        else:
            header_candidates = collect_header_candidates(
                cells=cells,
                headers=headers,
                header_threshold=header_threshold,
            )

            _, _, _, _, selected_headers = infer_header_cluster(
                header_candidates=header_candidates,
                cluster_row_window=inferred_cluster_row_window,
            )

            if not selected_headers:
                continue

            rows = []
            cols = []

            for hit in selected_headers.values():
                h_row_min, h_col_min, h_row_max, h_col_max = hit.bounds
                rows.extend([h_row_min, h_row_max])
                cols.extend([h_col_min, h_col_max])

            candidates.append(
                SectionCandidate(
                    category=category_name,
                    category_index=index,
                    has_header_rules=True,
                    detection_method="inferred_from_headers",
                    category_label_found=False,
                    category_inferred_from_headers=True,
                    category_is_explicit=category_is_explicit,
                    category_anchor=None,
                    inferred_headers=selected_headers,
                    raw_row_min=min(rows),
                    raw_row_max=max(rows),
                    raw_col_min=min(cols),
                    raw_col_max=min(ws.max_column, max(cols) + value_lookahead),
                )
            )

    return candidates


#INI FUNGSI/BAGIAN UNTUK (menetapkan batas-batas tiap section)
def assign_section_boundaries(
    ws,
    section_candidates: List[SectionCandidate],
    value_lookahead: int,
) -> Dict[str, SectionCandidate]:
    detected = sorted(
        section_candidates,
        key=lambda item: (item.raw_row_min, item.raw_col_min, item.category_index),
    )

    for current in detected:
        current.row_min = max(1, current.raw_row_min)
        current.col_min = max(1, current.raw_col_min)

        lower_sections = [
            other
            for other in detected
            if other.raw_row_min > current.raw_row_min
        ]

        if lower_sections:
            next_lower_row = min(other.raw_row_min for other in lower_sections)
            current.row_max = max(current.row_min, next_lower_row - 1)
        else:
            current.row_max = ws.max_row

        current.col_max = min(
            ws.max_column,
            max(current.raw_col_max, current.raw_col_min + value_lookahead),
        )

    groups_by_start_row: Dict[int, List[SectionCandidate]] = {}

    for item in detected:
        groups_by_start_row.setdefault(item.raw_row_min, []).append(item)

    for _, group in groups_by_start_row.items():
        if len(group) <= 1:
            continue

        group.sort(key=lambda item: (item.raw_col_min, item.category_index))

        for i, current in enumerate(group):
            if i + 1 >= len(group):
                continue

            next_same_row = group[i + 1]

            if next_same_row.raw_col_min > current.raw_col_min:
                current.col_max = min(current.col_max, next_same_row.raw_col_min - 1)

    return {item.category: item for item in detected}


#INI FUNGSI/BAGIAN UNTUK (mengubah batas section menjadi dictionary)
def get_section_boundary_dict(section: Optional[SectionCandidate]) -> Optional[Dict[str, Any]]:
    if section is None:
        return None

    return {
        "row_min": section.row_min,
        "row_max": section.row_max,
        "col_min": section.col_min,
        "col_max": section.col_max,
        "range": boundary_location(section.row_min, section.col_min, section.row_max, section.col_max),
        "raw_detected_range": boundary_location(
            section.raw_row_min,
            section.raw_col_min,
            section.raw_row_max,
            section.raw_col_max,
        ),
    }


#INI FUNGSI/BAGIAN UNTUK (menyiapkan batas area yang dipesan)
def make_reserved_bounds(
    section: Optional[SectionCandidate],
    selected_headers: Optional[Dict[str, HeaderHit]] = None,
) -> List[Tuple[int, int, int, int]]:
    reserved = []

    if section and section.category_anchor:
        reserved.append(section.category_anchor.bounds)

    if selected_headers:
        for hit in selected_headers.values():
            reserved.append(hit.bounds)

    return reserved


#INI FUNGSI/BAGIAN UNTUK (menyimpulkan tata letak header (baris/kolom))
def infer_header_layout(selected_headers: Dict[str, HeaderHit]) -> str:
    if len(selected_headers) <= 1:
        return "single"

    rows = []
    cols = []

    for hit in selected_headers.values():
        row_min, col_min, row_max, col_max = hit.bounds
        rows.append((row_min + row_max) / 2)
        cols.append((col_min + col_max) / 2)

    row_span = max(rows) - min(rows)
    col_span = max(cols) - min(cols)

    if row_span <= 2 and col_span > 1:
        return "horizontal_table"

    if col_span <= 2 and row_span > 1:
        return "vertical_key_value"

    return "mixed_or_stacked"


#INI FUNGSI/BAGIAN UNTUK (mencari header berikutnya di baris bawah)
def get_next_header_row_below(
    current_header: HeaderHit,
    selected_headers: Dict[str, HeaderHit],
) -> Optional[int]:
    _, _, h_row_max, _ = current_header.bounds
    next_rows = []

    for hit in selected_headers.values():
        if hit.location == current_header.location:
            continue

        other_row_min, _, _, _ = hit.bounds

        if other_row_min <= h_row_max:
            continue

        next_rows.append(other_row_min)

    if not next_rows:
        return None

    return min(next_rows)


#INI FUNGSI/BAGIAN UNTUK (mencari header berikutnya di kolom kanan)
def get_next_header_col_right_on_same_row(
    current_header: HeaderHit,
    selected_headers: Dict[str, HeaderHit],
) -> Optional[int]:
    h_row_min, _, h_row_max, h_col_max = current_header.bounds
    next_cols = []

    for hit in selected_headers.values():
        if hit.location == current_header.location:
            continue

        other_row_min, other_col_min, other_row_max, _ = hit.bounds
        rows_overlap = not (other_row_max < h_row_min or h_row_max < other_row_min)

        if rows_overlap and other_col_min > h_col_max:
            next_cols.append(other_col_min)

    if not next_cols:
        return None

    return min(next_cols)


#INI FUNGSI/BAGIAN UNTUK (mengumpulkan value pada sebuah area)
def collect_values_in_region(
    ws,
    merged_origin_map: Dict[Tuple[int, int], Tuple[int, int]],
    merged_bounds_map: Dict[Tuple[int, int], Tuple[int, int, int, int]],
    region: Optional[Tuple[int, int, int, int]],
    rule_terms: List[str],
    reserved_bounds: List[Tuple[int, int, int, int]],
    section: Optional[SectionCandidate] = None,
) -> List[Dict[str, Any]]:
    if not is_valid_region(region):
        return []

    row_min, col_min, row_max, col_max = region
    values = []
    seen_origins: Set[Tuple[int, int]] = set()

    row_min = max(1, row_min)
    col_min = max(1, col_min)
    row_max = min(ws.max_row, row_max)
    col_max = min(ws.max_column, col_max)

    if row_min > row_max or col_min > col_max:
        return []

    for row in range(row_min, row_max + 1):
        if section is not None and row_is_helper_marker_row(
            ws=ws,
            row=row,
            section=section,
            merged_origin_map=merged_origin_map,
            merged_bounds_map=merged_bounds_map,
        ):
            continue

        for col in range(col_min, col_max + 1):
            info = get_cell_info(ws, row, col, merged_origin_map, merged_bounds_map)
            origin = (info["origin_row"], info["origin_col"])
            value = info["value"]

            if origin in seen_origins:
                continue

            seen_origins.add(origin)

            if is_empty(value):
                continue

            cell_bounds = info["bounds"]

            if any(bounds_intersect(cell_bounds, reserved) for reserved in reserved_bounds):
                continue

            if is_probably_rule_label(value, rule_terms):
                continue

            values.append(
                {
                    "value": value,
                    "location": info["origin_location"],
                    "cell_boundary": bounds_to_dict(cell_bounds),
                    "inside_value_boundary": True,
                }
            )

    return values




#INI FUNGSI/BAGIAN UNTUK (mengecek apakah nilai hanya penanda bantu)
def is_helper_marker_value(value: Any) -> bool:
    """Return True for internal mapping markers like -1, -2, etc."""
    if value is None or isinstance(value, bool):
        return False

    if isinstance(value, int):
        return value < 0

    if isinstance(value, float):
        return value < 0 and value.is_integer()

    return bool(re.fullmatch(r"-\d+", str(value).strip()))


#INI FUNGSI/BAGIAN UNTUK (mengecek apakah baris hanya berisi penanda bantu)
def row_is_helper_marker_row(
    ws,
    row: int,
    section: SectionCandidate,
    merged_origin_map: Dict[Tuple[int, int], Tuple[int, int]],
    merged_bounds_map: Dict[Tuple[int, int], Tuple[int, int, int, int]],
) -> bool:
    """Detect hidden/helper rows such as -1, -2, ... without hardcoding row numbers."""
    values = []
    seen_origins: Set[Tuple[int, int]] = set()

    for col in range(section.col_min, section.col_max + 1):
        info = get_cell_info(ws, row, col, merged_origin_map, merged_bounds_map)
        origin = (info["origin_row"], info["origin_col"])

        if origin in seen_origins:
            continue

        seen_origins.add(origin)
        value = info["value"]

        if is_empty(value):
            continue

        values.append(value)

    # A real data row can contain one negative number, but a helper row usually
    # contains several marker cells like -1, -2, -3 across the table.
    return len(values) >= 2 and all(is_helper_marker_value(value) for value in values)


#INI FUNGSI/BAGIAN UNTUK (mengecek apakah baris berisi data pada section)
def row_has_any_data_in_section(
    ws,
    row: int,
    section: SectionCandidate,
    merged_origin_map: Dict[Tuple[int, int], Tuple[int, int]],
    merged_bounds_map: Dict[Tuple[int, int], Tuple[int, int, int, int]],
    rule_terms: List[str],
    reserved_bounds: List[Tuple[int, int, int, int]],
) -> bool:
    if row_is_helper_marker_row(
        ws=ws,
        row=row,
        section=section,
        merged_origin_map=merged_origin_map,
        merged_bounds_map=merged_bounds_map,
    ):
        return False

    seen_origins = set()

    for col in range(section.col_min, section.col_max + 1):
        info = get_cell_info(ws, row, col, merged_origin_map, merged_bounds_map)
        origin = (info["origin_row"], info["origin_col"])

        if origin in seen_origins:
            continue

        seen_origins.add(origin)

        value = info["value"]

        if is_empty(value):
            continue

        cell_bounds = info["bounds"]

        if any(bounds_intersect(cell_bounds, reserved) for reserved in reserved_bounds):
            continue

        if is_probably_rule_label(value, rule_terms):
            continue

        return True

    return False


#INI FUNGSI/BAGIAN UNTUK (mengumpulkan value sambil mempertahankan sel kosong)
def collect_values_preserve_nulls(
    ws,
    merged_origin_map: Dict[Tuple[int, int], Tuple[int, int]],
    merged_bounds_map: Dict[Tuple[int, int], Tuple[int, int, int, int]],
    region: Optional[Tuple[int, int, int, int]],
    direction: Optional[str],
    section: SectionCandidate,
    rule_terms: List[str],
    reserved_bounds: List[Tuple[int, int, int, int]],
) -> List[Any]:
    if not is_valid_region(region):
        return []

    row_min, col_min, row_max, col_max = region

    row_min = max(1, row_min)
    col_min = max(1, col_min)
    row_max = min(ws.max_row, row_max)
    col_max = min(ws.max_column, col_max)

    if row_min > row_max or col_min > col_max:
        return []

    values = []

    if direction == "down":
        seen_origins: Set[Tuple[int, int]] = set()

        for row in range(row_min, row_max + 1):
            active_row = row_has_any_data_in_section(
                ws=ws,
                row=row,
                section=section,
                merged_origin_map=merged_origin_map,
                merged_bounds_map=merged_bounds_map,
                rule_terms=rule_terms,
                reserved_bounds=reserved_bounds,
            )

            if not active_row:
                continue

            row_values = []
            row_seen_origins: Set[Tuple[int, int]] = set()
            only_seen_duplicate_values = False

            for col in range(col_min, col_max + 1):
                info = get_cell_info(ws, row, col, merged_origin_map, merged_bounds_map)
                origin = (info["origin_row"], info["origin_col"])

                if origin in row_seen_origins:
                    continue

                row_seen_origins.add(origin)

                value = info["value"]
                cell_bounds = info["bounds"]

                if any(bounds_intersect(cell_bounds, reserved) for reserved in reserved_bounds):
                    continue

                if is_empty(value):
                    continue

                if is_probably_rule_label(value, rule_terms):
                    continue

                if origin in seen_origins:
                    only_seen_duplicate_values = True
                    continue

                seen_origins.add(origin)
                row_values.append(value)

            if not row_values and only_seen_duplicate_values:
                # This physical row is only the continuation of a vertically merged cell.
                # Do not append duplicate values or artificial nulls.
                continue

            if not row_values:
                values.append(None)
            elif len(row_values) == 1:
                values.append(row_values[0])
            else:
                values.append(row_values)

        return values

    if direction == "right":
        row_values = []
        seen_origins = set()

        for row in range(row_min, row_max + 1):
            for col in range(col_min, col_max + 1):
                info = get_cell_info(ws, row, col, merged_origin_map, merged_bounds_map)
                origin = (info["origin_row"], info["origin_col"])

                if origin in seen_origins:
                    continue

                seen_origins.add(origin)

                value = info["value"]
                cell_bounds = info["bounds"]

                if any(bounds_intersect(cell_bounds, reserved) for reserved in reserved_bounds):
                    continue

                if is_empty(value):
                    continue

                if is_probably_rule_label(value, rule_terms):
                    continue

                row_values.append(value)

        return row_values if row_values else [None]

    return []


#INI FUNGSI/BAGIAN UNTUK (memilih value terbaik dari kandidat)
def pick_best_value(candidate_values: List[Dict[str, Any]]) -> Tuple[Optional[Any], Optional[Dict[str, Any]]]:
    if not candidate_values:
        return None, None

    best = candidate_values[0]

    return best["value"], {
        "value": best["value"],
        "location": best["location"],
        "cell_boundary": best["cell_boundary"],
    }


#INI FUNGSI/BAGIAN UNTUK (menentukan area value untuk tiap header)
def build_header_value_regions(
    header_hit: HeaderHit,
    selected_headers: Dict[str, HeaderHit],
    section: SectionCandidate,
) -> Dict[str, Optional[Tuple[int, int, int, int]]]:
    h_row_min, h_col_min, h_row_max, h_col_max = header_hit.bounds

    next_header_row_below = get_next_header_row_below(
        current_header=header_hit,
        selected_headers=selected_headers,
    )

    next_header_col_right = get_next_header_col_right_on_same_row(
        current_header=header_hit,
        selected_headers=selected_headers,
    )

    down_row_min = h_row_max + 1

    if next_header_row_below is not None:
        down_row_max = min(section.row_max, next_header_row_below - 1)
    else:
        down_row_max = section.row_max

    down_region = (
        down_row_min,
        h_col_min,
        down_row_max,
        h_col_max,
    )

    right_col_min = h_col_max + 1

    if next_header_col_right is not None:
        right_col_max = min(section.col_max, next_header_col_right - 1)
    else:
        right_col_max = section.col_max

    right_region = (
        h_row_min,
        right_col_min,
        h_row_max,
        right_col_max,
    )

    return {
        "right": right_region if is_valid_region(right_region) else None,
        "down": down_region if is_valid_region(down_region) else None,
    }


#INI FUNGSI/BAGIAN UNTUK (mendeteksi area value di bawah/sebelah header)
def detect_header_value_area(
    ws,
    merged_origin_map: Dict[Tuple[int, int], Tuple[int, int]],
    merged_bounds_map: Dict[Tuple[int, int], Tuple[int, int, int, int]],
    header_hit: HeaderHit,
    selected_headers: Dict[str, HeaderHit],
    section: SectionCandidate,
    rule_terms: List[str],
) -> Dict[str, Any]:
    layout = infer_header_layout(selected_headers)

    regions = build_header_value_regions(
        header_hit=header_hit,
        selected_headers=selected_headers,
        section=section,
    )

    reserved_bounds = make_reserved_bounds(
        section=section,
        selected_headers=selected_headers,
    )

    right_values = collect_values_in_region(
        ws=ws,
        merged_origin_map=merged_origin_map,
        merged_bounds_map=merged_bounds_map,
        region=regions["right"],
        rule_terms=rule_terms,
        reserved_bounds=reserved_bounds,
        section=section,
    )

    down_values = collect_values_in_region(
        ws=ws,
        merged_origin_map=merged_origin_map,
        merged_bounds_map=merged_bounds_map,
        region=regions["down"],
        rule_terms=rule_terms,
        reserved_bounds=reserved_bounds,
        section=section,
    )

    h_row_min, h_col_min, h_row_max, h_col_max = header_hit.bounds
    header_col_span = h_col_max - h_col_min + 1

    if layout == "vertical_key_value":
        preferred_order = ["right", "down"]
    elif layout == "horizontal_table":
        preferred_order = ["down", "right"]
    elif layout == "mixed_or_stacked":
        if header_col_span > 1:
            preferred_order = ["down", "right"]
        else:
            preferred_order = ["right", "down"]
    else:
        preferred_order = ["right", "down"]

    value_sets = {
        "right": right_values,
        "down": down_values,
    }

    chosen_direction = None
    chosen_region = None
    candidates = []

    for direction in preferred_order:
        if regions[direction] is not None and value_sets[direction]:
            chosen_direction = direction
            chosen_region = regions[direction]
            candidates = value_sets[direction]
            break

    if chosen_direction is None:
        for direction in preferred_order:
            if regions[direction] is not None:
                chosen_direction = direction
                chosen_region = regions[direction]
                candidates = value_sets[direction]
                break

    value, value_source = pick_best_value(candidates)

    values = collect_values_preserve_nulls(
        ws=ws,
        merged_origin_map=merged_origin_map,
        merged_bounds_map=merged_bounds_map,
        region=chosen_region,
        direction=chosen_direction,
        section=section,
        rule_terms=rule_terms,
        reserved_bounds=reserved_bounds,
    )

    has_value = any(item is not None for item in values)

    return {
        "header_boundary": bounds_to_dict(header_hit.bounds),
        "value_boundary": bounds_to_dict(chosen_region),
        "value_direction": chosen_direction,
        "debug_alternative_value_boundaries": {
            "right": bounds_to_dict(regions["right"]),
            "down": bounds_to_dict(regions["down"]),
        },
        "has_value": has_value,
        "value": value,
        "values": values,
        "value_source": value_source,
        "candidate_values": candidates,
    }


#INI FUNGSI/BAGIAN UNTUK (mendeteksi area value untuk kategori tanpa header)
def detect_category_value_area_without_headers(
    ws,
    merged_origin_map: Dict[Tuple[int, int], Tuple[int, int]],
    merged_bounds_map: Dict[Tuple[int, int], Tuple[int, int, int, int]],
    section: SectionCandidate,
    rule_terms: List[str],
) -> Dict[str, Any]:
    anchor = section.category_anchor

    if anchor is None:
        return {
            "category_boundary": None,
            "value_boundary": None,
            "value_direction": None,
            "has_value": False,
            "value": None,
            "values": [],
            "value_source": None,
            "candidate_values": [],
        }

    c_row_min, c_col_min, c_row_max, c_col_max = anchor.bounds

    right_region = (
        section.row_min,
        c_col_max + 1,
        section.row_max,
        section.col_max,
    )

    down_region = (
        c_row_max + 1,
        section.col_min,
        section.row_max,
        section.col_max,
    )

    reserved_bounds = make_reserved_bounds(section=section, selected_headers=None)

    #INI FUNGSI/BAGIAN UNTUK (mengumpulkan value langsung dari sebuah area)
    def collect_direct_values(region: Optional[Tuple[int, int, int, int]]) -> List[Dict[str, Any]]:
        if not is_valid_region(region):
            return []

        row_min, col_min, row_max, col_max = region
        values = []
        seen_origins = set()

        for row in range(row_min, row_max + 1):
            for col in range(col_min, col_max + 1):
                info = get_cell_info(ws, row, col, merged_origin_map, merged_bounds_map)
                origin = (info["origin_row"], info["origin_col"])

                if origin in seen_origins:
                    continue

                seen_origins.add(origin)

                value = info["value"]
                cell_bounds = info["bounds"]

                if any(bounds_intersect(cell_bounds, reserved) for reserved in reserved_bounds):
                    continue

                if is_empty(value):
                    continue

                values.append(
                    {
                        "value": value,
                        "location": info["origin_location"],
                        "cell_boundary": bounds_to_dict(cell_bounds),
                        "inside_value_boundary": True,
                    }
                )

        return values

    right_values = collect_direct_values(right_region)

    if right_values:
        chosen_direction = "right"
        chosen_region = right_region
        candidates = right_values
    else:
        down_values = collect_direct_values(down_region)

        chosen_direction = "down"
        chosen_region = down_region
        candidates = down_values

    values = [item["value"] for item in candidates]
    value = values[0] if values else None

    value_source = None

    if candidates:
        first = candidates[0]
        value_source = {
            "value": first["value"],
            "location": first["location"],
            "cell_boundary": first["cell_boundary"],
        }

    return {
        "category_boundary": bounds_to_dict(anchor.bounds),
        "value_boundary": bounds_to_dict(chosen_region),
        "value_direction": chosen_direction,
        "debug_alternative_value_boundaries": {
            "right": bounds_to_dict(right_region),
            "down": bounds_to_dict(down_region),
        },
        "has_value": bool(values),
        "value": value,
        "values": values,
        "value_source": value_source,
        "candidate_values": candidates,
    }


#INI FUNGSI/BAGIAN UNTUK (memvalidasi kategori yang tidak punya header)
def validate_category_without_headers(
    ws,
    merged_origin_map: Dict[Tuple[int, int], Tuple[int, int]],
    merged_bounds_map: Dict[Tuple[int, int], Tuple[int, int, int, int]],
    category_rule: Dict[str, Any],
    rule_terms: List[str],
    section: Optional[SectionCandidate],
) -> Dict[str, Any]:
    category_name = category_rule.get("category", "")

    result = {
        "category": category_name,
        "has_header_rules": False,
        "category_label_found": section.category_label_found if section else False,
        "category_inferred_from_headers": False,
        "category_is_explicit": section.category_is_explicit if section else category_rule.get("category_is_explicit", False),
        "category_found": section is not None,
        "detection_method": section.detection_method if section else "not_found",
        "section_boundary": get_section_boundary_dict(section),
        "category_location": section.category_anchor.location if section and section.category_anchor else None,
        "matched_text": section.category_anchor.text if section and section.category_anchor else None,
        "match_score": section.category_anchor.score if section and section.category_anchor else None,
        "category_boundary": None,
        "value_boundary": None,
        "value_direction": None,
        "has_value": False,
        "value": None,
        "values": [],
        "value_source": None,
        "candidate_values": [],
        "note": None,
    }

    if section is None or section.category_anchor is None:
        result["note"] = (
            "Kategori ini tidak punya daftar header di rules.xlsx. "
            "Kalau label kategori tidak ada di file, sistem tidak punya anchor untuk menentukan batas section."
        )
        result["status"] = "missing_category_label"
        return result

    area = detect_category_value_area_without_headers(
        ws=ws,
        merged_origin_map=merged_origin_map,
        merged_bounds_map=merged_bounds_map,
        section=section,
        rule_terms=rule_terms,
    )

    result.update(area)

    # Hormati requirement level kategori: optional -> value boleh kosong tetap valid,
    # selama label/section kategorinya tetap ketemu (cek di atas).
    category_value_required = bool(category_rule.get("valueRequired", True))
    result["requirement"] = category_rule.get("requirement", "mandatory")
    result["value_required"] = category_value_required
    result["status"] = (
        "valid"
        if (area["has_value"] or not category_value_required)
        else "category_found_but_value_empty"
    )

    return result


#INI FUNGSI/BAGIAN UNTUK (memvalidasi kategori yang punya header)
def validate_category_with_headers(
    ws,
    cells: List[Dict[str, Any]],
    merged_origin_map: Dict[Tuple[int, int], Tuple[int, int]],
    merged_bounds_map: Dict[Tuple[int, int], Tuple[int, int, int, int]],
    category_rule: Dict[str, Any],
    rule_terms: List[str],
    section: Optional[SectionCandidate],
    header_threshold: float,
) -> Dict[str, Any]:
    category_name = category_rule.get("category", "")
    headers = category_rule.get("headers", [])

    if section is None:
        header_results = []

        for header_rule in headers:
            header_name = parse_header_name(header_rule)
            header_required = parse_header_required(header_rule)
            value_required = parse_value_required(header_rule)

            header_results.append(
                {
                    "header": header_name,
                    "required": header_required,
                    "value_required": value_required,
                    "found": False,
                    "location": None,
                    "matched_text": None,
                    "match_score": None,
                    "header_boundary": None,
                    "value_boundary": None,
                    "value_direction": None,
                    "has_value": False,
                    "value": None,
                    "values": [],
                    "value_source": None,
                    "candidate_values": [],
                    "status": "missing_header" if header_required else "optional_header_not_found",
                }
            )

        return {
            "category": category_name,
            "has_header_rules": True,
            "category_label_found": False,
            "category_inferred_from_headers": False,
            "category_is_explicit": category_rule.get("category_is_explicit", False),
            "category_found": False,
            "detection_method": "not_found",
            "section_boundary": None,
            "category_location": None,
            "matched_text": None,
            "match_score": None,
            "required_headers_found": 0,
            "required_headers_total": sum(1 for item in headers if parse_header_required(item)),
            "headers": header_results,
            "status": "not_found",
        }

    header_candidates = collect_header_candidates(
        cells=cells,
        headers=headers,
        header_threshold=header_threshold,
        row_min=section.row_min,
        row_max=section.row_max,
        col_min=section.col_min,
        col_max=section.col_max,
    )

    if section.category_anchor:
        selected_headers = choose_header_hits_near_anchor(
            header_candidates=header_candidates,
            anchor=section.category_anchor,
        )
    else:
        _, _, _, _, selected_headers = infer_header_cluster(
            header_candidates=header_candidates,
            cluster_row_window=max(1, section.row_max - section.row_min),
        )

    header_results = []
    found_required_headers = 0
    total_required_headers = 0
    value_required_missing = 0
    header_required_missing = 0

    for header_rule in headers:
        header_name = parse_header_name(header_rule)
        header_required = parse_header_required(header_rule)
        value_required = parse_value_required(header_rule)

        if not header_name:
            continue

        if header_required:
            total_required_headers += 1

        hit = selected_headers.get(header_name)

        if not hit:
            if header_required:
                header_required_missing += 1

            header_results.append(
                {
                    "header": header_name,
                    "required": header_required,
                    "value_required": value_required,
                    "found": False,
                    "location": None,
                    "matched_text": None,
                    "match_score": None,
                    "header_boundary": None,
                    "value_boundary": None,
                    "value_direction": None,
                    "has_value": False,
                    "value": None,
                    "values": [],
                    "value_source": None,
                    "candidate_values": [],
                    "status": "missing_header" if header_required else "optional_header_not_found",
                }
            )

            continue

        if header_required:
            found_required_headers += 1

        area = detect_header_value_area(
            ws=ws,
            merged_origin_map=merged_origin_map,
            merged_bounds_map=merged_bounds_map,
            header_hit=hit,
            selected_headers=selected_headers,
            section=section,
            rule_terms=rule_terms,
        )

        has_value = area["has_value"]

        if value_required and not has_value:
            value_required_missing += 1

        header_results.append(
            {
                "header": header_name,
                "required": header_required,
                "value_required": value_required,
                "found": True,
                "location": hit.location,
                "matched_text": hit.text,
                "match_score": hit.score,
                "header_boundary": area["header_boundary"],
                "value_boundary": area["value_boundary"],
                "value_direction": area["value_direction"],
                "debug_alternative_value_boundaries": area["debug_alternative_value_boundaries"],
                "has_value": has_value,
                "value": area["value"],
                "values": area["values"],
                "value_source": area["value_source"],
                "candidate_values": area["candidate_values"],
                "status": "valid" if has_value or not value_required else "header_found_but_value_empty",
            }
        )

    if header_required_missing > 0:
        status = "headers_incomplete"
    elif value_required_missing > 0:
        status = "values_incomplete"
    else:
        status = "valid"

    return {
        "category": category_name,
        "has_header_rules": True,
        "category_label_found": section.category_label_found,
        "category_inferred_from_headers": section.category_inferred_from_headers,
        "category_is_explicit": section.category_is_explicit,
        "category_found": True,
        "detection_method": section.detection_method,
        "section_boundary": get_section_boundary_dict(section),
        "category_location": section.category_anchor.location if section.category_anchor else None,
        "matched_text": section.category_anchor.text if section.category_anchor else None,
        "match_score": section.category_anchor.score if section.category_anchor else None,
        "header_layout": infer_header_layout(selected_headers),
        "required_headers_found": found_required_headers,
        "required_headers_total": total_required_headers,
        "headers": header_results,
        "status": status,
    }


#INI FUNGSI/BAGIAN UNTUK (menyusun ringkasan hasil yang mudah dibaca)
def build_readable_summary(category_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    ringkasan = {}

    for category in category_results:
        category_name = category.get("category")

        if not category_name:
            continue

        if category.get("has_header_rules"):
            ringkasan[category_name] = {}

            for header in category.get("headers", []):
                header_name = header.get("header")

                if not header_name:
                    continue

                ringkasan[category_name][header_name] = header.get("values", [])

        else:
            values = category.get("values", [])
            ringkasan[category_name] = values[0] if len(values) == 1 else values

    return ringkasan

#INI FUNGSI/BAGIAN UNTUK (memuat mapping kesesuaian metode)
def load_alignment_mapping(mapping_path: Optional[str]) -> Dict[str, Any]:
    if not mapping_path:
        raise FileNotFoundError(
            "File alignment_mapping.json tidak ditemukan. "
            "Letakkan file di folder Config/ atau gunakan argumen --alignment-map."
        )

    path = Path(mapping_path)

    if not path.exists():
        raise FileNotFoundError(f"File alignment_mapping.json tidak ditemukan: {path}")

    with open(path, "r", encoding="utf-8") as file:
        mapping = json.load(file)

    rules = mapping.get("rules")

    if not isinstance(rules, list):
        raise ValueError("Format alignment_mapping.json tidak valid. Key 'rules' harus berupa list.")

    for index, rule in enumerate(rules, start=1):
        if not isinstance(rule, dict):
            raise ValueError(f"Rule alignment ke-{index} harus berupa object/dict.")

        if not isinstance(rule.get("outcome_keywords"), list):
            raise ValueError(f"Rule alignment ke-{index} wajib punya outcome_keywords berupa list.")

        if not isinstance(rule.get("recommended_methods"), list):
            raise ValueError(f"Rule alignment ke-{index} wajib punya recommended_methods berupa list.")

    return mapping


#INI FUNGSI/BAGIAN UNTUK (menentukan path mapping kesesuaian)
def resolve_alignment_mapping_path(explicit_path: Optional[str], rules_path: str) -> Optional[str]:
    if explicit_path:
        return explicit_path

    # Cari di Config/ terlebih dahulu karena ini folder konfigurasi utama sistem.
    config_path = resolve_config_path("alignment_mapping.json")
    if config_path:
        return str(config_path)

    # Lalu cari di folder yang sama dengan file rules.xlsx.
    candidate = Path(rules_path).with_name("alignment_mapping.json")
    if candidate.exists():
        return str(candidate)

    # Terakhir cari di folder kerja saat ini.
    candidate = Path("alignment_mapping.json")
    if candidate.exists():
        return str(candidate)

    return None

#INI FUNGSI/BAGIAN UNTUK (mengecek apakah value kosong)
def value_is_blank(value: Any) -> bool:
    if value is None:
        return True

    if isinstance(value, list):
        return all(value_is_blank(item) for item in value)

    return str(value).strip() == ""


#INI FUNGSI/BAGIAN UNTUK (mengubah value menjadi teks)
def value_to_text(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, list):
        return " ".join(value_to_text(item) for item in value if not value_is_blank(item))

    return str(value).strip()


#INI FUNGSI/BAGIAN UNTUK (mencari hasil sebuah kategori)
def find_category_result(category_results: List[Dict[str, Any]], category_name: Optional[str]) -> Optional[Dict[str, Any]]:
    if not category_name:
        return None

    category_norm = normalize_text(category_name)

    for category in category_results:
        if normalize_text(category.get("category")) == category_norm:
            return category

    return None


#INI FUNGSI/BAGIAN UNTUK (mencari hasil sebuah header)
def find_header_result(category: Dict[str, Any], header_name: Optional[str]) -> Optional[Dict[str, Any]]:
    if not category or not header_name:
        return None

    header_norm = normalize_text(header_name)

    for header in category.get("headers", []):
        if normalize_text(header.get("header")) == header_norm:
            return header

    return None


#INI FUNGSI/BAGIAN UNTUK (mendeteksi aturan kesesuaian dari teks deskripsi)
def detect_alignment_rule(description_text: str, mapping: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    description_norm = normalize_text(description_text)

    if not description_norm:
        return None

    for rule in mapping.get("rules", []):
        for keyword in rule.get("outcome_keywords", []):
            keyword_norm = normalize_text(keyword)

            if keyword_norm and keyword_norm in description_norm:
                return rule

    return None


#INI FUNGSI/BAGIAN UNTUK (mencocokkan metode dengan rekomendasi)
def method_matches_recommendation(method_text: str, recommended_methods: List[str]) -> List[str]:
    method_norm = normalize_text(method_text)
    matched = []

    for method in recommended_methods:
        method_norm_rule = normalize_text(method)

        if method_norm_rule and method_norm_rule in method_norm:
            matched.append(method)

    return matched


#INI FUNGSI/BAGIAN UNTUK (menilai kesesuaian pasangan deskripsi dan metode)
def evaluate_alignment_pair(description_value: Any, method_value: Any, mapping: Dict[str, Any]) -> Dict[str, Any]:
    description_text = value_to_text(description_value)
    method_text = value_to_text(method_value)

    if value_is_blank(description_value) and value_is_blank(method_value):
        return {
            "status": "TIDAK SESUAI",
            "score": 0,
            "matched_keywords": [],
            "matched_methods": [],
            "expected_methods": [],
            "analysis": "Deskripsi capaian dan metode pembelajaran sama-sama belum terisi, sehingga kesesuaian tidak dapat dibuktikan.",
        }

    if value_is_blank(description_value):
        return {
            "status": "TIDAK SESUAI",
            "score": 0,
            "matched_keywords": [],
            "matched_methods": [],
            "expected_methods": [],
            "analysis": "Deskripsi capaian belum terisi, sehingga metode pembelajaran tidak memiliki acuan yang dapat dibandingkan.",
        }

    if value_is_blank(method_value):
        return {
            "status": "TIDAK SESUAI",
            "score": 0,
            "matched_keywords": [],
            "matched_methods": [],
            "expected_methods": [],
            "analysis": "Metode pembelajaran belum terisi, sehingga belum mendukung deskripsi capaian yang ada.",
        }

    rule = detect_alignment_rule(description_text, mapping)

    if rule is None:
        return {
            "status": "KURANG SESUAI",
            "score": 50,
            "matched_keywords": [],
            "matched_methods": [],
            "expected_methods": [],
            "analysis": "Sistem belum menemukan kata kunci capaian pada mapping. Tambahkan kata kunci capaian dan metode rekomendasi pada alignment_mapping.json jika pola ini memang valid.",
        }

    recommended = rule.get("recommended_methods", [])
    matched_methods = method_matches_recommendation(method_text, recommended)
    matched_keywords = [kw for kw in rule.get("outcome_keywords", []) if normalize_text(kw) in normalize_text(description_text)]

    if matched_methods:
        return {
            "status": "SESUAI",
            "score": 100,
            "matched_keywords": matched_keywords,
            "matched_methods": matched_methods,
            "expected_methods": recommended,
            "analysis": f"Metode pembelajaran sesuai karena memuat {', '.join(matched_methods)} yang mendukung capaian pada deskripsi. {rule.get('reason', '')}".strip(),
        }

    return {
        "status": "KURANG SESUAI",
        "score": 50,
        "matched_keywords": matched_keywords,
        "matched_methods": [],
        "expected_methods": recommended,
        "analysis": f"Metode pembelajaran belum menunjukkan dukungan yang kuat terhadap capaian pada deskripsi. Sebaiknya metode memuat salah satu pendekatan berikut: {', '.join(recommended)}.",
    }


#INI FUNGSI/BAGIAN UNTUK (menyusun analisis kesesuaian metode pembelajaran)
def build_alignment_analysis(
    category_results: List[Dict[str, Any]],
    rules: Dict[str, Any],
    mapping: Dict[str, Any],
) -> Dict[str, Any]:
    config = rules.get("alignment_config", {}) or {}

    if not config.get("enabled"):
        return {
            "alignment_pairs": [],
            "alignment_results": [],
            "alignment_summary": {
                "enabled": False,
                "message": "Alignment tidak dijalankan karena rules tidak memiliki minimal 2 header berwarna hijau.",
            },
        }

    category_name = config.get("category")
    description_header_name = config.get("description_header")
    method_header_name = config.get("method_header")

    category = find_category_result(category_results, category_name)

    if category is None:
        return {
            "alignment_pairs": [],
            "alignment_results": [],
            "alignment_summary": {
                "enabled": False,
                "message": f"Alignment tidak dijalankan karena kategori '{category_name}' tidak ditemukan pada hasil validasi.",
            },
        }

    description_header = find_header_result(category, description_header_name)
    method_header = find_header_result(category, method_header_name)

    if description_header is None or method_header is None:
        return {
            "alignment_pairs": [],
            "alignment_results": [],
            "alignment_summary": {
                "enabled": False,
                "message": "Alignment tidak dijalankan karena salah satu header hijau tidak ditemukan pada hasil validasi.",
                "description_header": description_header_name,
                "method_header": method_header_name,
            },
        }

    description_values = description_header.get("values", []) or []
    method_values = method_header.get("values", []) or []
    total_rows = max(len(description_values), len(method_values))
    results = []

    for index in range(total_rows):
        description_value = description_values[index] if index < len(description_values) else None
        method_value = method_values[index] if index < len(method_values) else None
        evaluation = evaluate_alignment_pair(description_value, method_value, mapping)

        results.append(
            {
                "row_index": index + 1,
                "category": category_name,
                "checked_headers": [description_header_name, method_header_name],
                "description_header": description_header_name,
                "method_header": method_header_name,
                "description_value": description_value,
                "method_value": method_value,
                **evaluation,
            }
        )

    total = len(results)
    sesuai = sum(1 for item in results if item.get("status") == "SESUAI")
    kurang_sesuai = sum(1 for item in results if item.get("status") == "KURANG SESUAI")
    tidak_sesuai = sum(1 for item in results if item.get("status") == "TIDAK SESUAI")
    avg_score = round(sum(item.get("score", 0) for item in results) / total, 2) if total else 0

    return {
        "alignment_pairs": [
            {
                "category": category_name,
                "headers": [description_header_name, method_header_name],
                "description_header": description_header_name,
                "method_header": method_header_name,
                "source": config.get("source"),
                "markers": config.get("markers", []),
            }
        ],
        "alignment_results": results,
        "alignment_summary": {
            "enabled": True,
            "category": category_name,
            "description_header": description_header_name,
            "method_header": method_header_name,
            "total_rows": total,
            "sesuai": sesuai,
            "kurang_sesuai": kurang_sesuai,
            "tidak_sesuai": tidak_sesuai,
            "score": avg_score,
        },
    }


#INI FUNGSI/BAGIAN UNTUK (fungsi utama: menjalankan seluruh validasi RPS)
def validate_rps(
    xlsx_path: str,
    rules_path: str,
    category_threshold: float = 0.86,
    header_threshold: float = 0.82,
    value_lookahead: int = 8,
    explicit_category_search_window: int = 80,
    inferred_cluster_row_window: int = 14,
    alignment_mapping_path: Optional[str] = None,
) -> Dict[str, Any]:
    rules = load_rules(rules_path)
    rules = normalize_rules_structure(rules)

    workbook = load_workbook(xlsx_path, data_only=True)
    ws = workbook.active

    merged_origin_map, merged_bounds_map = get_merged_cell_maps(ws)
    cells = build_cells(ws, merged_origin_map, merged_bounds_map)
    rule_terms = make_rule_terms(rules)

    section_candidates = detect_section_candidates(
        ws=ws,
        cells=cells,
        rules=rules,
        category_threshold=category_threshold,
        header_threshold=header_threshold,
        explicit_category_search_window=explicit_category_search_window,
        inferred_cluster_row_window=inferred_cluster_row_window,
        value_lookahead=value_lookahead,
    )

    section_map = assign_section_boundaries(
        ws=ws,
        section_candidates=section_candidates,
        value_lookahead=value_lookahead,
    )

    category_results = []

    for category_rule in rules.get("categories", []):
        category_name = category_rule.get("category", "")
        headers = category_rule.get("headers", [])
        section = section_map.get(category_name)

        if headers:
            result = validate_category_with_headers(
                ws=ws,
                cells=cells,
                merged_origin_map=merged_origin_map,
                merged_bounds_map=merged_bounds_map,
                category_rule=category_rule,
                rule_terms=rule_terms,
                section=section,
                header_threshold=header_threshold,
            )
        else:
            result = validate_category_without_headers(
                ws=ws,
                merged_origin_map=merged_origin_map,
                merged_bounds_map=merged_bounds_map,
                category_rule=category_rule,
                rule_terms=rule_terms,
                section=section,
            )

        category_results.append(result)

    detected_sections = []

    for section in sorted(section_map.values(), key=lambda item: (item.row_min, item.col_min, item.category_index)):
        detected_sections.append(
            {
                "category": section.category,
                "detection_method": section.detection_method,
                "category_label_found": section.category_label_found,
                "category_inferred_from_headers": section.category_inferred_from_headers,
                "category_is_explicit": section.category_is_explicit,
                "section_boundary": get_section_boundary_dict(section),
                "category_location": section.category_anchor.location if section.category_anchor else None,
                "category_boundary": bounds_to_dict(section.category_anchor.bounds)
                if section.category_anchor
                else None,
                "inferred_header_locations": {
                    header_name: hit.location
                    for header_name, hit in section.inferred_headers.items()
                },
                "inferred_header_boundaries": {
                    header_name: bounds_to_dict(hit.bounds)
                    for header_name, hit in section.inferred_headers.items()
                },
            }
        )

    mapping_path = resolve_alignment_mapping_path(alignment_mapping_path, rules_path)
    alignment_mapping = load_alignment_mapping(mapping_path)
    alignment_data = build_alignment_analysis(
        category_results=category_results,
        rules=rules,
        mapping=alignment_mapping,
    )

    total_categories = len(category_results)
    valid_categories = sum(1 for item in category_results if item.get("status") == "valid")
    found_categories = sum(1 for item in category_results if item.get("category_found"))
    inferred_categories = sum(1 for item in category_results if item.get("category_inferred_from_headers"))

    missing_or_invalid = [
        item.get("category")
        for item in category_results
        if item.get("status") != "valid"
    ]

    return {
        "file": str(Path(xlsx_path).name),
        "sheet": ws.title,
        "summary": {
            "total_categories": total_categories,
            "found_categories": found_categories,
            "valid_categories": valid_categories,
            "inferred_categories_from_headers": inferred_categories,
            "invalid_categories": total_categories - valid_categories,
            "is_valid": valid_categories == total_categories,
            "missing_or_invalid_categories": missing_or_invalid,
        },
        "settings": {
            "rules_path": str(Path(rules_path).name),
            "category_threshold": category_threshold,
            "header_threshold": header_threshold,
            "value_lookahead": value_lookahead,
            "explicit_category_search_window": explicit_category_search_window,
            "inferred_cluster_row_window": inferred_cluster_row_window,
            "alignment_mapping_path": str(Path(mapping_path).name) if mapping_path else None,
            "note": (
                "Header adalah anchor utama. Category bersifat opsional dan hanya membantu memisahkan section. "
                "Untuk kategori tanpa header, value langsung diambil dari kanan kategori sampai kolom terakhir sheet, "
                "termasuk merged cell seperti B6:J6."
            ),
        },
        "detected_sections": detected_sections,
        "alignment_mapping_path": mapping_path,
        "alignment_pairs": alignment_data.get("alignment_pairs", []),
        "alignment_results": alignment_data.get("alignment_results", []),
        "alignment_summary": alignment_data.get("alignment_summary", {}),
        "categories": category_results,
    }


#INI FUNGSI/BAGIAN UNTUK (entry point untuk menjalankan dari command line)
def main():
    parser = argparse.ArgumentParser(
        description="Validator dinamis dokumen RPS XLSX berdasarkan rules.xlsx."
    )

    parser.add_argument(
        "--xlsx",
        required=True,
        help="Path file RPS .xlsx. Cari di INPUT/ folder, atau path absolute/relative.",
    )

    parser.add_argument(
        "--rules",
        required=True,
        help="Path rules.xlsx. Cari di Config/ atau INPUT/ folder, atau path absolute/relative.",
    )

    parser.add_argument(
        "--out",
        default="result.json",
        help="Path output JSON debug. Default: result.json di DEBUG/ folder.",
    )

    parser.add_argument(
        "--summary-out",
        default=None,
        help="Path output JSON ringkasan. Default: ringkasan.json di DEBUG/ folder.",
    )

    parser.add_argument(
        "--pdf-out",
        default=None,
        help="Path output laporan audit PDF. Default: laporan_audit.pdf di OUTPUT/ folder.",
    )

    parser.add_argument(
        "--alignment-map",
        default=None,
        help="Path alignment_mapping.json. Default: cari di Config/, folder rules, lalu folder kerja saat ini.",
    )

    parser.add_argument(
        "--pretty",
        action="store_true",
        help="Output JSON dibuat rapi dengan indentasi.",
    )

    parser.add_argument(
        "--category-threshold",
        type=float,
        default=0.86,
        help="Batas kemiripan untuk mendeteksi kategori. Default: 0.86",
    )

    parser.add_argument(
        "--header-threshold",
        type=float,
        default=0.82,
        help="Batas kemiripan untuk mendeteksi header. Default: 0.82",
    )

    parser.add_argument(
        "--value-lookahead",
        type=int,
        default=8,
        help="Lebar bantuan pencarian value ke kanan. Default: 8",
    )

    parser.add_argument(
        "--explicit-category-search-window",
        type=int,
        default=80,
        help="Jumlah baris pencarian awal setelah kategori eksplisit ditemukan. Default: 80",
    )

    parser.add_argument(
        "--inferred-cluster-row-window",
        type=int,
        default=14,
        help="Jarak cluster baris untuk infer kategori dari kumpulan header. Default: 14",
    )

    args = parser.parse_args()

    # Resolve input file paths
    xlsx_path = str(resolve_input_path(args.xlsx))
    rules_path = str(resolve_input_path(args.rules))

    result = validate_rps(
        xlsx_path=xlsx_path,
        rules_path=rules_path,
        category_threshold=args.category_threshold,
        header_threshold=args.header_threshold,
        value_lookahead=args.value_lookahead,
        explicit_category_search_window=args.explicit_category_search_window,
        inferred_cluster_row_window=args.inferred_cluster_row_window,
        alignment_mapping_path=args.alignment_map,
    )

    result = clean_output_data(result)
    ringkasan = build_readable_summary(result.get("categories", []))

    indent = 2 if args.pretty else None

    result_json = json.dumps(
        result,
        ensure_ascii=False,
        indent=indent,
        default=str,
    )

    ringkasan_json = json.dumps(
        ringkasan,
        ensure_ascii=False,
        indent=indent,
        default=str,
    )

    if args.out:
        result_path = resolve_output_path(args.out)

        with open(result_path, "w", encoding="utf-8") as file:
            file.write(result_json)

        if args.summary_out:
            ringkasan_path = resolve_output_path(args.summary_out)
        else:
            ringkasan_path = resolve_output_path("ringkasan.json")

        with open(ringkasan_path, "w", encoding="utf-8") as file:
            file.write(ringkasan_json)

        if args.pdf_out:
            pdf_path = resolve_output_path(args.pdf_out)
        else:
            pdf_path = resolve_output_path("laporan_audit.pdf")

        generate_audit_pdf(result, str(pdf_path))

        print(f"Debug result saved to: {result_path}")
        print(f"Readable summary saved to: {ringkasan_path}")
        print(f"Audit PDF saved to: {pdf_path}")

    else:
        print(result_json)

        ringkasan_path = resolve_output_path(args.summary_out or "ringkasan.json")

        with open(ringkasan_path, "w", encoding="utf-8") as file:
            file.write(ringkasan_json)

        if args.pdf_out:
            pdf_path = resolve_output_path(args.pdf_out)
        else:
            pdf_path = resolve_output_path("laporan_audit.pdf")

        generate_audit_pdf(result, str(pdf_path))

        print(f"Readable summary saved to: {ringkasan_path}")
        print(f"Audit PDF saved to: {pdf_path}")


if __name__ == "__main__":
    main()