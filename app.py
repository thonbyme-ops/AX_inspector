import io
import json
import os
import secrets
import time
import uuid
from functools import wraps

from dotenv import load_dotenv
from flask import Flask, render_template, request, jsonify, abort, send_file, session, redirect, url_for
from werkzeug.utils import secure_filename

load_dotenv()

import db
from extractor import extract, row_definitions, COLUMNS
from evidence_extractor import extract_labor_cost_rows, extract_retirement_fund_rows
from hr_cost_extractor import (
    extract_hr_costs,
    aggregate_by_company,
    aggregate_by_person,
    aggregate_by_month,
    build_export_workbook,
    CATEGORY_ORDER,
)
from settlement_pdf_extractor import extract_settlement_pdf

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
DATA_DIR = os.path.join(BASE_DIR, "data")
VERIFICATIONS_PATH = os.path.join(DATA_DIR, "verifications.json")
SECRET_KEY_PATH = os.path.join(DATA_DIR, ".secret_key")
ALLOWED_EXT = {"pdf", "xlsx", "xlsm", "xls"}
MAX_CONTENT_LENGTH = 60 * 1024 * 1024

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)


def _load_or_create_secret_key():
    env_key = os.environ.get("FLASK_SECRET_KEY")
    if env_key:
        return env_key
    if os.path.exists(SECRET_KEY_PATH):
        with open(SECRET_KEY_PATH, "r", encoding="utf-8") as f:
            return f.read().strip()
    key = secrets.token_hex(32)
    with open(SECRET_KEY_PATH, "w", encoding="utf-8") as f:
        f.write(key)
    return key


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
app.config["MAX_FORM_MEMORY_SIZE"] = MAX_CONTENT_LENGTH
app.secret_key = _load_or_create_secret_key()
db.init_db()


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "account_id" not in session:
            if request.path.startswith("/api/"):
                return jsonify({"error": "로그인이 필요합니다."}), 401
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def current_account():
    return {
        "id": session.get("account_id"),
        "username": session.get("username"),
        "display_name": session.get("display_name"),
        "company": session.get("company"),
    }

# 진행 중인 비교 결과를 메모리에 잠시 보관 (검증 버튼 클릭 시 조회용)
_COMPARISONS = {}

# 이슈#2: 인건비(건강/연금/퇴직공제) 추출 결과를 내보내기 전까지 메모리에 보관
_HR_COST_RESULTS = {}
HR_COST_ALLOWED_EXT = {"xlsx", "xlsm", "xls", "pdf"}

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


def allowed_hr_cost_file(filename):
    return "." in filename and filename.rsplit(".", 1)[-1].lower() in HR_COST_ALLOWED_EXT


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
    unread_ocr_count = 0
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
            # 경비 항목 등 실제 값이 0/빈칸일 수 있어, OCR로 열 자체를 읽지 못한
            # 경우(행은 인식했지만 숫자를 못 찾음)를 '값 없음'과 구분해 보여준다.
            unread_a = present_a and col != "비고" and val_a is None and result_a["source"] == "ocr"
            unread_b = present_b and col != "비고" and val_b is None and result_b["source"] == "ocr"
            if unread_a:
                unread_ocr_count += 1
            if unread_b:
                unread_ocr_count += 1
            cols[col] = {"a": val_a, "b": val_b, "diff": is_diff, "unread_a": unread_a, "unread_b": unread_b}
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
        "unread_ocr_count": unread_ocr_count,
    }


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return render_template("login.html", error=None)

    username = (request.form.get("username") or "").strip()
    password = request.form.get("password") or ""
    account = db.verify_login(username, password)
    if not account:
        return render_template("login.html", error="아이디 또는 비밀번호가 올바르지 않습니다."), 401

    session["account_id"] = account["id"]
    session["username"] = account["username"]
    session["display_name"] = account["display_name"]
    session["company"] = account["company"]
    next_url = request.form.get("next") or request.args.get("next") or url_for("index")
    return redirect(next_url)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def index():
    return render_template("index.html")


@app.route("/api/compare", methods=["POST"])
@login_required
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
@login_required
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


@app.route("/api/hr_cost/extract", methods=["POST"])
@login_required
def api_hr_cost_extract():
    file_a = request.files.get("file_a")
    file_b = request.files.get("file_b")
    if not file_a and not file_b:
        return jsonify({"error": "최소 한 개의 엑셀 파일을 업로드해주세요."}), 400

    saved_paths = []
    records = []
    pdf_stats = []
    filenames = []
    try:
        for f in (file_a, file_b):
            if not f or not f.filename:
                continue
            filename = secure_filename(f.filename)
            if not filename or not allowed_hr_cost_file(filename):
                raise ValueError(f"엑셀(.xlsx/.xlsm/.xls) 또는 PDF 파일만 업로드할 수 있습니다: {f.filename}")
            filenames.append(filename)
            path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}_{filename}")
            f.save(path)
            saved_paths.append(path)
            ext = filename.lower().rsplit(".", 1)[-1]
            if ext == "pdf":
                pdf_records, stats = extract_settlement_pdf(path, filename)
                records.extend(pdf_records)
                pdf_stats.append({"filename": filename, **stats})
            else:
                records.extend(extract_hr_costs(path, filename))
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"파일 처리 중 오류가 발생했습니다: {e}"}), 500
    finally:
        for p in saved_paths:
            _safe_remove(p)

    if not records:
        return jsonify({"error": "추출된 데이터가 없습니다. 지원하는 양식인지 확인해주세요."}), 400

    token = uuid.uuid4().hex
    _HR_COST_RESULTS[token] = {"records": records, "filenames": filenames, "created_at": time.time()}

    html = render_template(
        "_hr_cost_result.html",
        token=token,
        by_company=aggregate_by_company(records),
        by_person=aggregate_by_person(records),
        by_month=aggregate_by_month(records),
        person_count=len({(r["company"], r["person"]) for r in records}),
        company_count=len({r["company"] for r in records}),
        record_count=len(records),
        pdf_stats=pdf_stats,
    )
    return jsonify({"html": html, "token": token})


@app.route("/api/hr_cost/export/<token>")
@login_required
def api_hr_cost_export(token):
    record = _HR_COST_RESULTS.get(token)
    if not record:
        abort(404, "추출 결과를 찾을 수 없습니다. 다시 업로드해주세요.")
    wb = build_export_workbook(record["records"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="인건비_추출결과.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/hr_cost/save/<token>", methods=["POST"])
@login_required
def api_hr_cost_save(token):
    record = _HR_COST_RESULTS.get(token)
    if not record:
        return jsonify({"error": "추출 결과를 찾을 수 없습니다. 다시 업로드해주세요."}), 404
    try:
        upload_id = db.save_upload(current_account(), record.get("filenames", []), record["records"])
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"ok": True, "upload_id": upload_id})


@app.route("/api/history/companies")
@login_required
def api_history_companies():
    return jsonify({"companies": db.list_companies(current_account())})


@app.route("/api/history/months")
@login_required
def api_history_months():
    company = request.args.get("company", "").strip()
    if not company:
        return jsonify({"error": "업체를 선택해주세요."}), 400
    try:
        months = db.list_year_months(current_account(), company)
    except PermissionError as e:
        return jsonify({"error": str(e)}), 403
    return jsonify({"months": months})


@app.route("/api/history/compare", methods=["POST"])
@login_required
def api_history_compare():
    data = request.get_json(silent=True) or {}
    company = (data.get("company") or "").strip()
    months = [m for m in (data.get("months") or []) if m]
    if not company or not months:
        return jsonify({"error": "업체와 연월을 선택해주세요."}), 400

    account = current_account()
    try:
        rows = db.get_latest_by_month(account, company, months)
        history_by_month = {ym: db.get_upload_history(account, company, ym) for ym in months}
    except PermissionError as e:
        return jsonify({"error": str(e)}), 403

    by_month = {}
    for r in rows:
        by_month.setdefault(r["year_month"], {c: 0 for c in CATEGORY_ORDER})[r["category"]] = r["amount"]
    compare_rows = []
    for ym in sorted(by_month):
        cats = by_month[ym]
        compare_rows.append({"year_month": ym, **cats, "total": sum(cats.values())})

    html = render_template(
        "_history_result.html",
        company=company,
        compare_rows=compare_rows,
        history_by_month=history_by_month,
    )
    return jsonify({"html": html})


@app.route("/api/verify", methods=["POST"])
@login_required
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
    # PDF OCR 추출은 페이지 수에 따라 수십 분 이상 걸릴 수 있어, threaded=True로
    # 그 동안 다른 요청(홈페이지 등)까지 완전히 멈추지 않게 한다.
    app.run(debug=True, port=5000, use_reloader=False, threaded=True)
