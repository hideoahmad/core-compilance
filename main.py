from __future__ import annotations

import mimetypes
import os
import re
import shutil
import sys
from io import BytesIO
from pathlib import Path
from threading import Lock
from typing import Optional

from flask import Flask, jsonify, request, send_file

ROOT = Path(__file__).resolve().parent

# Cukup root project saja.
# Jangan masukkan subfolder Buku_Kurikulum dan laporan_penelitian ke sys.path
# karena keduanya punya struktur modul yang bisa bentrok, misalnya PDF/pdfgenerator.py.
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RPS.PDF.pdfgenerator import (
    generate_audit_pdf,
    get_overall_status as get_rps_overall_status,
    overall_score as rps_overall_score,
)
from RPS.validate_rps import (
    clean_output_data,
    validate_rps,
)

from Buku_Kurikulum.validate_buku_kurikulum import (
    build_result as build_buku_result,
    ensure_dirs as ensure_buku_dirs,
    generate_buku_kurikulum_pdf,
)
from Buku_Kurikulum.PDF.pdfgenerator import (
    overall_status as get_buku_overall_status,
)

from laporan_penelitian.validate_laporan_penelitian import (
    build_result as build_penelitian_result,
    ensure_dirs as ensure_penelitian_dirs,
    generate_laporan_penelitian_pdf,
)
from laporan_penelitian.PDF.pdfgenerator import (
    overall_status as get_penelitian_overall_status,
)

app = Flask(__name__)
app.config["JSON_AS_ASCII"] = False

RPS_ROOT = ROOT / "RPS"
BUKU_ROOT = ROOT / "Buku_Kurikulum"
PENELITIAN_ROOT = ROOT / "laporan_penelitian"

ALLOWED_RPS_EXT = {".xlsx", ".xlsm", ".xls"}
ALLOWED_PDF_EXT = {".pdf"}
ALLOWED_RULES_EXT = {".xlsx"}

# Sistem antrean.
# Selama 1 request belum selesai bikin dan kirim PDF,
# request berikutnya tidak akan masuk ke proses validasi.
PROCESS_LOCK = Lock()

#INI FUNGSI/BAGIAN UNTUK (membuat response error JSON yang seragam)
def json_error(message: str, status: int = 400):
    return jsonify({"status": "error", "message": message}), status

#INI FUNGSI/BAGIAN UNTUK (menormalkan nilai tipe dokumen (rps/buku/penelitian))
def normalize_doc_type(value: Optional[str]) -> str:
    norm = (value or "").strip().lower()

    if norm in {"laporan_penelitian", "laporan-penelitian", "penelitian"}:
        return "penelitian"

    if norm in {"buku_kurikulum", "buku-kurikulum", "buku"}:
        return "buku"

    return norm

#INI FUNGSI/BAGIAN UNTUK (mengubah teks menjadi nilai boolean)
def parse_bool(value: Optional[str]) -> bool:
    return str(value or "").strip().lower() in {
        "1",
        "true",
        "yes",
        "ya",
        "y",
        "on",
    }

#INI FUNGSI/BAGIAN UNTUK (mengambil ekstensi file dengan huruf kecil secara aman)
def safe_suffix(filename: str) -> str:
    return Path(filename or "").suffix.lower()

#INI FUNGSI/BAGIAN UNTUK (membuang prefix angka/timestamp acak di depan nama file)
def strip_random_prefix(stem: str) -> str:
    # Buang prefix angka/timestamp acak dari frontend, contoh "1778927999882-Aturan" -> "Aturan".
    return re.sub(r"^\d{10,}[-_\s]+", "", stem).strip()


#INI FUNGSI/BAGIAN UNTUK (membersihkan nama dasar file dari karakter ilegal dan angka acak)
def clean_filename_stem(filename: str, fallback: str = "Dokumen") -> str:
    stem = Path(filename or fallback).stem.strip()

    # Hapus karakter yang tidak aman untuk nama file Windows/Linux.
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()

    # Buang prefix angka acak (timestamp) supaya nama file tampil apa adanya sesuai inputan.
    stem = strip_random_prefix(stem)
    stem = stem.strip(" .")

    return stem or fallback

#INI FUNGSI/BAGIAN UNTUK (menyusun nama file upload yang bersih (nama bersih + ekstensi))
def clean_uploaded_filename(filename: str, fallback_stem: str = "File") -> str:
    original_name = Path(filename or "").name.strip()
    suffix = Path(original_name).suffix.lower()

    stem = clean_filename_stem(original_name, fallback=fallback_stem)

    if not suffix:
        suffix = ".xlsx"

    return f"{stem}{suffix}"

#INI FUNGSI/BAGIAN UNTUK (menyusun nama file PDF laporan sesuai tipe dan nama input)
def make_report_filename(doc_type: str, input_filename: str) -> str:
    input_name = clean_filename_stem(input_filename)

    if doc_type == "rps":
        return f"Laporan Validasi RPS - {input_name}.pdf"

    if doc_type == "buku":
        return f"Laporan Validasi Buku Kurikulum - {input_name}.pdf"

    if doc_type == "penelitian":
        return f"Laporan Validasi Laporan Penelitian - {input_name}.pdf"

    return f"Laporan Validasi - {input_name}.pdf"

#INI FUNGSI/BAGIAN UNTUK (mengosongkan seluruh isi sebuah folder)
def clear_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)

    for item in path.iterdir():
        try:
            if item.is_file() or item.is_symlink():
                item.unlink()
            elif item.is_dir():
                shutil.rmtree(item)
        except FileNotFoundError:
            pass

#INI FUNGSI/BAGIAN UNTUK (menghapus semua file rules .xlsx lama di folder config)
def clear_rules_files(config_dir: Path) -> None:
    config_dir.mkdir(parents=True, exist_ok=True)

    for item in config_dir.iterdir():
        try:
            if item.is_file() and item.suffix.lower() == ".xlsx":
                item.unlink()
        except FileNotFoundError:
            pass


#INI FUNGSI/BAGIAN UNTUK (menghapus satu file jika ada)
def cleanup_file(path: Optional[Path]) -> None:
    if not path:
        return

    try:
        if path.exists() and path.is_file():
            path.unlink()
    except OSError:
        pass


#INI FUNGSI/BAGIAN UNTUK (membersihkan beberapa folder kerja sekaligus)
def cleanup_work_dirs(*paths: Path) -> None:
    for path in paths:
        try:
            clear_dir(path)
        except OSError:
            pass

#INI FUNGSI/BAGIAN UNTUK (mengirim PDF ke client lalu menghapus file fisiknya)
def send_pdf_and_delete(path: Path, download_name: str, compliance_status: Optional[str] = None):
    if not path.exists():
        return json_error("laporan belum tersedia", 404)

    data = path.read_bytes()

    # Setelah PDF masuk memory, file fisik langsung dihapus.
    cleanup_file(path)

    mime, _ = mimetypes.guess_type(download_name)

    response = send_file(
        BytesIO(data),
        as_attachment=True,
        download_name=download_name,
        mimetype=mime or "application/pdf",
    )

    # Body response tetap PDF, status dikirim lewat header agar laporan tetap langsung ter-download.
    if compliance_status:
        response.headers["X-Compliance-Status"] = compliance_status
        response.headers["Access-Control-Expose-Headers"] = "X-Compliance-Status"

    return response




#INI FUNGSI/BAGIAN UNTUK (menghitung status kepatuhan keseluruhan RPS)
def get_rps_compliance_status(result: dict) -> str:
    categories = result.get("categories", [])
    score = rps_overall_score(categories)
    return get_rps_overall_status(score)

#INI FUNGSI/BAGIAN UNTUK (menyimpan file upload ke folder tujuan dengan validasi ekstensi)
def save_uploaded_file(
    file_storage,
    dest_dir: Path,
    allowed_exts: set,
    fallback_stem: str = "File",
) -> Path:
    suffix = safe_suffix(file_storage.filename)

    if suffix not in allowed_exts:
        raise ValueError(f"Ekstensi tidak didukung: {suffix}")

    dest_dir.mkdir(parents=True, exist_ok=True)

    target_name = clean_uploaded_filename(
        file_storage.filename,
        fallback_stem=fallback_stem,
    )

    target = dest_dir / target_name
    file_storage.save(str(target))

    return target


#INI FUNGSI/BAGIAN UNTUK (mencari file yang sudah ada berdasarkan nama)
def pick_existing_file(folder: Path, name: Optional[str]) -> Optional[Path]:
    if not name:
        return None

    safe_name = Path(name).name
    candidate = folder / safe_name

    if candidate.exists():
        return candidate

    return None


#INI FUNGSI/BAGIAN UNTUK (mengambil file .xlsx paling baru di sebuah folder)
def pick_latest_xlsx(folder: Path) -> Optional[Path]:
    if not folder.exists():
        return None

    candidates = [
        item
        for item in folder.iterdir()
        if item.is_file() and item.suffix.lower() == ".xlsx"
    ]

    if not candidates:
        return None

    candidates.sort(key=lambda item: item.stat().st_mtime, reverse=True)
    return candidates[0]


#INI FUNGSI/BAGIAN UNTUK (mencari file rules dari request upload)
def find_uploaded_rules_file(include_file_key: bool = False):
    preferred_keys = ["rules", "rules_file", "rulesFile"]

    for key in preferred_keys:
        item = request.files.get(key)
        if item and item.filename:
            return item

    for key, item in request.files.items():
        if key == "file" and not include_file_key:
            continue

        if safe_suffix(item.filename) == ".xlsx":
            return item

    return None


#INI FUNGSI/BAGIAN UNTUK (menentukan path rules RPS dari upload atau config tersimpan)
def get_rps_rules_path(rules_file, rules_name: Optional[str]) -> Optional[Path]:
    config_dir = RPS_ROOT / "Config"

    if rules_file:
        clear_rules_files(config_dir)
        return save_uploaded_file(
            rules_file,
            config_dir,
            ALLOWED_RULES_EXT,
            fallback_stem="Rules RPS",
        )

    by_name = pick_existing_file(config_dir, rules_name)
    if by_name:
        return by_name

    latest = pick_latest_xlsx(config_dir)
    if latest:
        return latest

    return None


#INI FUNGSI/BAGIAN UNTUK (menentukan path rules Buku Kurikulum)
def get_buku_rules_path(rules_file, rules_name: Optional[str]) -> Optional[Path]:
    config_dir = BUKU_ROOT / "config"

    if rules_file:
        clear_rules_files(config_dir)
        return save_uploaded_file(
            rules_file,
            config_dir,
            ALLOWED_RULES_EXT,
            fallback_stem="Rules Buku Kurikulum",
        )

    by_name = pick_existing_file(config_dir, rules_name)
    if by_name:
        return by_name

    latest = pick_latest_xlsx(config_dir)
    if latest:
        return latest

    return None


#INI FUNGSI/BAGIAN UNTUK (menentukan path rules Laporan Penelitian)
def get_penelitian_rules_path(rules_file, rules_name: Optional[str]) -> Optional[Path]:
    config_dir = PENELITIAN_ROOT / "config"

    if rules_file:
        clear_rules_files(config_dir)
        return save_uploaded_file(
            rules_file,
            config_dir,
            ALLOWED_RULES_EXT,
            fallback_stem="Rules Laporan Penelitian",
        )

    by_name = pick_existing_file(config_dir, rules_name)
    if by_name:
        return by_name

    latest = pick_latest_xlsx(config_dir)
    if latest:
        return latest

    return None


#INI FUNGSI/BAGIAN UNTUK (endpoint cek kesehatan server)
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


#INI FUNGSI/BAGIAN UNTUK (endpoint upload rules (dibungkus antrean lock))
@app.route("/api/rules", methods=["POST"])
def upload_rules():
    with PROCESS_LOCK:
        return upload_rules_locked()


#INI FUNGSI/BAGIAN UNTUK (proses inti upload rules)
def upload_rules_locked():
    doc_type = normalize_doc_type(request.form.get("type"))

    if doc_type not in {"rps", "buku", "penelitian"}:
        return json_error("type harus rps, buku_kurikulum, atau laporan_penelitian")

    rules_file = find_uploaded_rules_file(include_file_key=True)

    if not rules_file:
        return json_error("rules wajib diupload")

    rules_name = request.form.get("rules_name")

    try:
        if doc_type == "rps":
            rules_path = get_rps_rules_path(rules_file, rules_name)
        elif doc_type == "penelitian":
            rules_path = get_penelitian_rules_path(rules_file, rules_name)
        else:
            rules_path = get_buku_rules_path(rules_file, rules_name)
    except ValueError as exc:
        return json_error(str(exc))
    except OSError as exc:
        return json_error(f"gagal menyimpan rules: {exc}", 500)

    if not rules_path:
        return json_error("rules tidak ditemukan")

    return jsonify(
        {
            "status": "ok",
            "type": doc_type,
            "rules": {
                "name": rules_path.name,
            },
        }
    )


#INI FUNGSI/BAGIAN UNTUK (endpoint validasi dokumen (dibungkus antrean lock))
@app.route("/api/validate", methods=["POST"])
def validate():
    # Semua request validasi masuk antrean di sini.
    # Request berikutnya baru jalan setelah request sebelumnya selesai total.
    with PROCESS_LOCK:
        return validate_locked()


#INI FUNGSI/BAGIAN UNTUK (proses inti validasi dan pemilihan handler sesuai tipe)
def validate_locked():
    doc_type = normalize_doc_type(request.form.get("type"))

    if doc_type not in {"rps", "buku", "penelitian"}:
        return json_error("type harus rps, buku_kurikulum, atau laporan_penelitian")

    extra_files = [key for key in request.files.keys() if key != "file"]
    if extra_files:
        return json_error("upload input hanya boleh file utama saja tanpa file lain")

    uploaded = request.files.get("file")

    if not uploaded or not uploaded.filename:
        return json_error("file wajib diupload")

    if doc_type == "rps":
        return handle_rps(uploaded)

    if doc_type == "penelitian":
        return handle_penelitian(uploaded)

    return handle_buku(uploaded)


#INI FUNGSI/BAGIAN UNTUK (alur validasi RPS: simpan input, validasi, buat PDF, kirim)
def handle_rps(uploaded):
    suffix = safe_suffix(uploaded.filename)

    if suffix not in ALLOWED_RPS_EXT:
        return json_error("file RPS harus .xlsx/.xlsm/.xls")

    report_name = make_report_filename("rps", uploaded.filename)

    input_dir = RPS_ROOT / "INPUT"
    output_dir = RPS_ROOT / "OUTPUT"

    input_path = input_dir / f"input{suffix}"
    output_pdf = output_dir / report_name

    try:
        clear_dir(input_dir)
        clear_dir(output_dir)

        uploaded.save(str(input_path))

        rules_file = find_uploaded_rules_file()
        rules_name = request.form.get("rules_name")

        rules_path = get_rps_rules_path(rules_file, rules_name)

        if not rules_path:
            return json_error("rules tidak ditemukan. Upload dulu via /api/rules")

        result = validate_rps(
            xlsx_path=str(input_path),
            rules_path=str(rules_path),
            alignment_mapping_path=None,
        )

        result = clean_output_data(result)

        # Pakai nama input asli + nama rules apa adanya (tanpa path, tanpa angka acak) untuk ditampilkan di PDF.
        result["file"] = clean_uploaded_filename(uploaded.filename, fallback_stem="Dokumen")
        result["rules_path"] = rules_path.name

        generate_audit_pdf(result, str(output_pdf))

        if not output_pdf.exists():
            return json_error("PDF RPS gagal dibuat", 500)

        compliance_status = get_rps_compliance_status(result)

        return send_pdf_and_delete(
            output_pdf,
            download_name=report_name,
            compliance_status=compliance_status,
        )

    except ValueError as exc:
        return json_error(str(exc))

    except Exception as exc:
        return json_error(f"gagal validasi RPS: {exc}", 500)

    finally:
        cleanup_work_dirs(input_dir, output_dir)


#INI FUNGSI/BAGIAN UNTUK (alur validasi Buku Kurikulum)
def handle_buku(uploaded):
    suffix = safe_suffix(uploaded.filename)

    if suffix not in ALLOWED_PDF_EXT:
        return json_error("file Buku Kurikulum harus .pdf")

    report_name = make_report_filename("buku", uploaded.filename)

    input_dir = BUKU_ROOT / "input"
    output_dir = BUKU_ROOT / "output"
    debug_dir = BUKU_ROOT / "debug"

    input_path = input_dir / "input.pdf"
    output_pdf = output_dir / report_name

    try:
        ensure_buku_dirs(BUKU_ROOT)

        clear_dir(input_dir)
        clear_dir(output_dir)
        clear_dir(debug_dir)

        uploaded.save(str(input_path))

        rules_file = find_uploaded_rules_file()
        rules_name = request.form.get("rules_name")

        rules_path = get_buku_rules_path(rules_file, rules_name)

        if not rules_path:
            return json_error("rules tidak ditemukan. Upload dulu via /api/rules")

        result = build_buku_result(
            input_path,
            rules_path,
            debug_dir,
            output_pdf,
        )

        # Pakai nama input asli + nama rules apa adanya (tanpa path, tanpa angka acak) untuk ditampilkan di PDF.
        result["file"] = clean_uploaded_filename(uploaded.filename, fallback_stem="Dokumen")
        result["rules_path"] = rules_path.name

        generate_buku_kurikulum_pdf(result, str(output_pdf))

        if not output_pdf.exists():
            return json_error("PDF Buku Kurikulum gagal dibuat", 500)

        compliance_status = get_buku_overall_status(result)

        return send_pdf_and_delete(
            output_pdf,
            download_name=report_name,
            compliance_status=compliance_status,
        )

    except ValueError as exc:
        return json_error(str(exc))

    except Exception as exc:
        return json_error(f"gagal validasi Buku Kurikulum: {exc}", 500)

    finally:
        cleanup_work_dirs(input_dir, output_dir, debug_dir)


#INI FUNGSI/BAGIAN UNTUK (alur validasi Laporan Penelitian)
def handle_penelitian(uploaded):
    suffix = safe_suffix(uploaded.filename)

    if suffix not in ALLOWED_PDF_EXT:
        return json_error("file Laporan Penelitian harus .pdf")

    report_name = make_report_filename("penelitian", uploaded.filename)

    input_dir = PENELITIAN_ROOT / "input"
    output_dir = PENELITIAN_ROOT / "output"
    debug_dir = PENELITIAN_ROOT / "debug"

    input_path = input_dir / "input.pdf"
    output_pdf = output_dir / report_name

    try:
        ensure_penelitian_dirs(PENELITIAN_ROOT)

        clear_dir(input_dir)
        clear_dir(output_dir)
        clear_dir(debug_dir)

        uploaded.save(str(input_path))

        rules_file = find_uploaded_rules_file()
        rules_name = request.form.get("rules_name")

        rules_path = get_penelitian_rules_path(rules_file, rules_name)

        if not rules_path:
            return json_error("rules tidak ditemukan. Upload dulu via /api/rules")

        result = build_penelitian_result(
            input_path,
            rules_path,
            debug_dir,
            output_pdf,
        )

        # Pakai nama input asli + nama rules apa adanya (tanpa path, tanpa angka acak) untuk ditampilkan di PDF.
        result["file"] = clean_uploaded_filename(uploaded.filename, fallback_stem="Dokumen")
        result["rules_path"] = rules_path.name

        generate_laporan_penelitian_pdf(result, str(output_pdf))

        if not output_pdf.exists():
            return json_error("PDF Laporan Penelitian gagal dibuat", 500)

        compliance_status = get_penelitian_overall_status(result)

        return send_pdf_and_delete(
            output_pdf,
            download_name=report_name,
            compliance_status=compliance_status,
        )

    except ValueError as exc:
        return json_error(str(exc))

    except Exception as exc:
        return json_error(f"gagal validasi Laporan Penelitian: {exc}", 500)

    finally:
        cleanup_work_dirs(input_dir, output_dir, debug_dir)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    debug = parse_bool(os.environ.get("FLASK_DEBUG"))

    app.run(
        host="0.0.0.0",
        port=port,
        debug=debug,
        threaded=True,
    )