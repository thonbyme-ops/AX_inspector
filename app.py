import json
import os
import time
import uuid

from flask import Flask, render_template, request, jsonify, abort
from werkzeug.utils import secure_filename

from extractor import extract, row_definitions, COLUMNS
from evidence_extractor import extract_labor_cost_rows, extract_retirement_fund_rows

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
DATA_DIR = os.path.join(BASE_DIR, "data")
VERIFICATIONS_PATH = os.path.join(DATA_DIR, "verifications.json")
ALLOWED_EXT = {"pdf", "xlsx", "xlsm", "xls"}
MAX_CONTENT_LENGTH = 60 * 1024 * 1024

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
app.config["MAX_FORM_MEMORY_SIZE"] = MAX_CONTENT_LENGTH

# 진행 중인 비교 결과를 메모리에 잠시 보관 (검증 버튼 클릭 시 조회용)
_COMPARISONS = {}

META_LABELS = [
    ("project_name", "공사명"),
    ("contract_no", "계약번호"),
    ("calc_period", "기성고 산정기간"),
    ("project_period", "공사기간"),
    ("contractor", "계약자"),
    ("supply_amount", "공급가액"),
    ("vat", "부가세"),
    ("contract_date", "계약일"),
]


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[-1].lower() in ALLOWED_EXT


def save_upload(file_storage, slot):
    filename = secure_filename(file_storage.filename)
    if not filename or not allowed_file(filename):
        raise ValueError(f"허용되지 않는 파일 형식입니다: {file_storage.filename}")
    unique_name = f"{uuid.uuid4().hex}_{slot}_{filename}"
    path = os.path.join(UPLOAD_DIR, unique_name)
    file_storage.save(path)
    return path, filename


def build_comparison(result_a, result_b):
    rows = []
    for key, label, _keyword, is_key_field in row_definitions():
        item_a = result_a["items"].get(key, {})
        item_b = result_b["items"].get(key, {})
        cols = {}
        row_has_diff = False
        for col in COLUMNS:
            val_a = item_a.get(col)
            val_b = item_b.get(col)
            present_a = key in result_a["items"]
            present_b = key in result_b["items"]
            is_diff = present_a and present_b and val_a != val_b
            if is_diff:
                row_has_diff = True
            cols[col] = {"a": val_a, "b": val_b, "diff": is_diff}
        rows.append(
            {
                "key": key,
                "label": label,
                "is_key_field": is_key_field,
                "missing_a": key not in result_a["items"],
                "missing_b": key not in result_b["items"],
                "has_diff": row_has_diff,
                "cols": cols,
            }
        )

    meta_rows = []
    meta_has_diff = False
    for key, label in META_LABELS:
        val_a = result_a["meta"].get(key)
        val_b = result_b["meta"].get(key)
        is_diff = bool(val_a) and bool(val_b) and val_a != val_b
        if is_diff:
            meta_has_diff = True
        meta_rows.append({"key": key, "label": label, "a": val_a, "b": val_b, "diff": is_diff})

    any_diff = meta_has_diff or any(r["has_diff"] for r in rows)
    return {
        "meta_rows": meta_rows,
        "rows": rows,
        "any_diff": any_diff,
        "source_a": result_a["source"],
        "source_b": result_b["source"],
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/compare", methods=["POST"])
def api_compare():
    file_a = request.files.get("file_a")
    file_b = request.files.get("file_b")
    if not file_a or not file_b:
        return jsonify({"error": "두 개의 파일을 모두 업로드해주세요."}), 400

    saved_paths = []
    try:
        path_a, name_a = save_upload(file_a, "a")
        path_b, name_b = save_upload(file_b, "b")
        saved_paths = [path_a, path_b]

        result_a = extract(path_a, name_a)
        result_b = extract(path_b, name_b)
    except ValueError as e:
        for p in saved_paths:
            _safe_remove(p)
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        for p in saved_paths:
            _safe_remove(p)
        return jsonify({"error": f"파일 처리 중 오류가 발생했습니다: {e}"}), 500
    finally:
        for p in saved_paths:
            _safe_remove(p)

    comparison = build_comparison(result_a, result_b)
    token = uuid.uuid4().hex
    _COMPARISONS[token] = {
        "file_a": name_a,
        "file_b": name_b,
        "comparison": comparison,
        "created_at": time.time(),
    }

    html = render_template(
        "_result.html",
        token=token,
        name_a=name_a,
        name_b=name_b,
        comparison=comparison,
    )
    return jsonify({"html": html, "token": token})


EVIDENCE_MAX_PAGES = 15


@app.route("/api/compare_evidence", methods=["POST"])
def api_compare_evidence():
    file_a = request.files.get("file_a")  # 노무비 지급 내역서
    file_b = request.files.get("file_b")  # 퇴직공제부금 납부 신고 내역
    if not file_a or not file_b:
        return jsonify({"error": "두 개의 파일을 모두 업로드해주세요."}), 400

    saved_paths = []
    try:
        path_a, name_a = save_upload(file_a, "a")
        path_b, name_b = save_upload(file_b, "b")
        saved_paths = [path_a, path_b]

        if not name_a.lower().endswith(".pdf") or not name_b.lower().endswith(".pdf"):
            raise ValueError("이 비교 모드는 스캔 PDF만 지원합니다.")

        labor = extract_labor_cost_rows(path_a, max_pages=EVIDENCE_MAX_PAGES)
        retirement = extract_retirement_fund_rows(path_b, max_pages=EVIDENCE_MAX_PAGES)
    except ValueError as e:
        for p in saved_paths:
            _safe_remove(p)
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        for p in saved_paths:
            _safe_remove(p)
        return jsonify({"error": f"파일 처리 중 오류가 발생했습니다: {e}"}), 500
    finally:
        for p in saved_paths:
            _safe_remove(p)

    needs_review_count = sum(1 for r in labor["rows"] if r["needs_review"]) + sum(
        1 for r in retirement["rows"] if r["needs_review"]
    )
    correction_count = sum(1 for r in retirement["rows"] if r.get("has_correction"))

    token = uuid.uuid4().hex
    _COMPARISONS[token] = {
        "file_a": name_a,
        "file_b": name_b,
        "comparison": {"any_diff": needs_review_count > 0},
        "created_at": time.time(),
    }

    html = render_template(
        "_evidence_result.html",
        token=token,
        name_a=name_a,
        name_b=name_b,
        labor=labor,
        retirement=retirement,
        needs_review_count=needs_review_count,
        correction_count=correction_count,
        max_pages=EVIDENCE_MAX_PAGES,
    )
    return jsonify({"html": html, "token": token})


@app.route("/api/verify", methods=["POST"])
def api_verify():
    data = request.get_json(silent=True) or {}
    token = data.get("token")
    decision = data.get("decision")
    supervisor = (data.get("supervisor") or "").strip()

    if token not in _COMPARISONS:
        abort(400, "비교 결과를 찾을 수 없습니다. 다시 업로드해주세요.")
    if decision not in ("yes", "no"):
        abort(400, "검증 결과는 yes 또는 no 여야 합니다.")
    if not supervisor:
        abort(400, "검증자 성명을 입력해주세요.")

    record = _COMPARISONS[token]
    entry = {
        "token": token,
        "file_a": record["file_a"],
        "file_b": record["file_b"],
        "decision": decision,
        "supervisor": supervisor,
        "any_diff": record["comparison"]["any_diff"],
        "verified_at": time.time(),
    }

    existing = []
    if os.path.exists(VERIFICATIONS_PATH):
        with open(VERIFICATIONS_PATH, "r", encoding="utf-8") as f:
            existing = json.load(f)
    existing.append(entry)
    with open(VERIFICATIONS_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    del _COMPARISONS[token]
    return jsonify({"ok": True, "entry": entry})


def _safe_remove(path):
    try:
        os.remove(path)
    except OSError:
        pass


if __name__ == "__main__":
    app.run(debug=True, port=5000, use_reloader=False)
