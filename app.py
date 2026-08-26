import io
import json
import os
import re
import secrets
import time
import uuid
from datetime import datetime
from functools import wraps

from dotenv import load_dotenv
from flask import Flask, render_template, request, jsonify, abort, send_file, session, redirect, url_for

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
    CATEGORY_LABELS,
)
from attendance_cross_check import CROSS_HEADERS, build_cross_check_workbook, cross_check
from labor_ledger_extractor import build_labor_template_workbook, parse_labor_ledger
from settlement_pdf_extractor import extract_settlement_pdf

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
DATA_DIR = os.path.join(BASE_DIR, "data")
VERIFICATIONS_PATH = os.path.join(DATA_DIR, "verifications.json")
SECRET_KEY_PATH = os.path.join(DATA_DIR, ".secret_key")
ALLOWED_EXT = {"pdf", "xlsx", "xlsm", "xls"}
# 실제 업무 파일이 크다 -- 실측: "기성고 검사 요청서.pdf" 47.5MB,
# "06-03. 제20회 기성 실적정산(건강요양).pdf" 41.5MB, "06-04.(국민연금)" 29.1MB.
# 두 개만 올려도 예전 상한 60MB를 넘어 413이 났고, 그때 Werkzeug는 본문을 받지 않고
# 연결을 끊어서 브라우저에서는 "업로드 중"에서 그대로 멈춘 것처럼 보였다(실측 재현).
MAX_CONTENT_LENGTH = int(os.environ.get("MAX_UPLOAD_MB", "300")) * 1024 * 1024

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


@app.errorhandler(413)
def handle_too_large(_e):
    """업로드 용량 초과를 JSON으로 알려준다.

    기본 413 응답은 HTML이라 프론트엔드의 `res.json()`이 예외를 던지고, 그 결과
    실제 원인("파일이 너무 큼") 대신 "네트워크 오류"만 보였다. 게다가 Werkzeug는
    상한을 넘으면 본문을 읽지 않고 끊어서 브라우저가 매달리기도 한다 -- 그래서
    클라이언트에서 보내기 전에 크기를 먼저 확인하고(static/script.js), 여기서는
    그래도 서버까지 닿은 경우에 대비한다.
    """
    limit_mb = MAX_CONTENT_LENGTH // (1024 * 1024)
    message = (
        f"업로드 용량 상한({limit_mb}MB)을 넘었습니다. 파일을 나눠 올리거나, "
        f"서버를 MAX_UPLOAD_MB 환경변수로 더 크게 설정해주세요."
    )
    if request.path.startswith("/api/"):
        return jsonify({"error": message}), 413
    return message, 413


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

# 이슈#5-3: 하도급사별 노무비 지급 명세서를 공통 스키마로 정규화한 결과를
# 표준 템플릿으로 내보내기 전까지 메모리에 보관
_LEDGER_RESULTS = {}
# 노무비 명세서는 엑셀만이지만, 같이 올리는 대조 자료로는 공단 발급 PDF도 받는다
# (이슈 #5-1의 정밀 추출 경로 -- 생년월일이 있어 동명이인을 갈라 대조할 수 있다).
LEDGER_ALLOWED_EXT = {"xlsx", "xlsm", "pdf"}

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


def allowed_ledger_file(filename):
    return "." in filename and filename.rsplit(".", 1)[-1].lower() in LEDGER_ALLOWED_EXT


def safe_original_filename(filename):
    """경로 조작·제어문자만 제거하고 한글 등 비-ASCII 문자는 그대로 보존한다.

    werkzeug의 secure_filename()은 한글을 통째로 지워버려서, PDF 문서 종류를
    파일명의 한글 키워드(건강/연금/퇴직 등)로 판별하는 로직이 항상 실패하는
    문제가 있었다 (예: "06-05. 제20회 기성 실적정산(퇴직공제).pdf" -> "06-05._20_.pdf").
    """
    name = os.path.basename((filename or "").replace("\\", "/"))
    name = name.replace("\x00", "")
    name = re.sub(r'[<>:"/\\|?*\r\n\t]', "_", name)
    return name.strip(" .")


def save_upload(file_storage, slot):
    filename = safe_original_filename(file_storage.filename)
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


def _load_verifications():
    if not os.path.exists(VERIFICATIONS_PATH):
        return []
    with open(VERIFICATIONS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


@app.route("/")
@login_required
def index():
    verifications = _load_verifications()
    stats = {
        "total": len(verifications),
        "match_count": sum(1 for v in verifications if v.get("decision") == "yes"),
        "mismatch_count": sum(1 for v in verifications if v.get("decision") == "no"),
    }
    recent = sorted(verifications, key=lambda v: v.get("verified_at", 0), reverse=True)[:10]
    for v in recent:
        v["verified_at_display"] = datetime.fromtimestamp(v["verified_at"]).strftime("%Y-%m-%d %H:%M")
    return render_template(
        "index.html", stats=stats, recent=recent, max_upload_bytes=MAX_CONTENT_LENGTH
    )


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
            filename = safe_original_filename(f.filename)
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
        category_order=CATEGORY_ORDER,
        category_labels=CATEGORY_LABELS,
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
        download_name="데이터_추출결과.xlsx",
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


@app.route("/api/labor_ledger/extract", methods=["POST"])
@login_required
def api_labor_ledger_extract():
    """하도급사별 노무비 지급 명세서(양식 상이)를 공통 스키마로 정규화한다 (이슈 #5-3)."""
    uploads = [
        f
        for key in ("file_a", "file_b")
        for f in request.files.getlist(key)
        if f and f.filename
    ]
    if not uploads:
        return jsonify({"error": "노무비 지급 명세서 엑셀을 최소 한 개 업로드해주세요."}), 400

    saved_paths = []
    attendance, summaries, filenames = [], [], []
    premium, premium_files, premium_skipped = [], [], []
    try:
        for f in uploads:
            filename = safe_original_filename(f.filename)
            if not filename or not allowed_ledger_file(filename):
                # 실제 제출본에 구형 .xls도 섞여 있는데(실측: NSC 26.03) openpyxl이
                # 읽지 못하므로, 그냥 거부하지 않고 변환 방법을 알려준다.
                extra = (
                    " 엑셀에서 열어 '다른 이름으로 저장 > Excel 통합 문서(.xlsx)'로 바꿔 올려주세요."
                    if filename.lower().endswith(".xls")
                    else ""
                )
                raise ValueError(
                    f"노무비 명세서는 엑셀(.xlsx/.xlsm), 대조 자료는 엑셀 또는 공단 발급 PDF만 "
                    f"업로드할 수 있습니다: {f.filename}.{extra}"
                )
            path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}_{filename}")
            f.save(path)
            saved_paths.append(path)

            # 공단 발급 PDF는 좌표 기반 정밀 추출 경로로 보낸다 -- 성명·생년월일까지
            # 나오므로 대조에서 동명이인을 가를 수 있다.
            if filename.lower().endswith(".pdf"):
                records, _stats = extract_settlement_pdf(path, filename)
                precise = [r for r in records if _usable_for_cross_check(r)]
                if precise:
                    premium.extend(precise)
                    premium_files.append(filename)
                if len(precise) < len(records):
                    premium_skipped.append(
                        {"filename": filename, "kept": len(precise), "dropped": len(records) - len(precise)}
                    )
                continue

            # 엑셀은 파일명이 아니라 내용으로 가른다 -- 노무비 명세서는 성명+일자별
            # 공수 표가 있고, 보험료 원장/퇴직공제 신고는 그런 표가 없는 대신 기존
            # 인건비 추출 양식에 걸린다.
            rows, people = parse_labor_ledger(path, filename)
            if people:
                filenames.append(filename)
                attendance.extend(rows)
                summaries.extend(people)
                continue
            try:
                records = extract_hr_costs(path, filename)
            except Exception:
                records = []
            if records:
                premium.extend(records)
                premium_files.append(filename)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"파일 처리 중 오류가 발생했습니다: {e}"}), 500
    finally:
        for p in saved_paths:
            _safe_remove(p)

    if not summaries:
        return jsonify({
            "error": "출역 데이터를 찾지 못했습니다. 성명과 일자별 공수가 있는 "
                     "'노무비 지급 명세서' 시트가 포함된 파일인지 확인해주세요."
        }), 400

    cross_rows, cross_reverse, cross_summary = ([], [], [])
    if premium:
        cross_rows, cross_reverse, cross_summary = cross_check(summaries, premium)

    token = uuid.uuid4().hex
    _LEDGER_RESULTS[token] = {
        "attendance": attendance,
        "summaries": summaries,
        "cross": (cross_rows, cross_reverse, cross_summary),
        "created_at": time.time(),
    }
    html = render_template(
        "_labor_ledger_result.html",
        token=token,
        filenames=filenames,
        premium_files=premium_files,
        premium_skipped=premium_skipped,
        summaries=sorted(summaries, key=lambda r: (r["company"] or "", r["year_month"] or "", r["person"] or "")),
        by_company=_ledger_by_company(summaries),
        attendance_count=len(attendance),
        needs_review_count=sum(1 for s in summaries if s["needs_review"]),
        cross_rows=sorted(cross_rows, key=lambda r: (not r["needs_review"], r["company"] or "", r["year_month"] or "", r["person"] or "")),
        cross_reverse=cross_reverse,
        cross_summary=cross_summary,
        cross_review_count=sum(1 for r in cross_rows if r["needs_review"]),
        cross_headers=CROSS_HEADERS,
    )
    return jsonify({"html": html, "token": token})


def _usable_for_cross_check(record):
    """이 PDF 레코드를 대조에 써도 되는지 판단한다.

    스캔본 PDF는 좌표 기반 정밀 경로가 서식을 못 잡으면 페이지 전체 OCR로 떨어지는데,
    그 결과는 대조에 넣으면 오히려 해롭다 -- 실측(신한ACT 26.03 건강, 순수 스캔본
    4페이지): 성명이 전부 "확인필요_..."로 나오고 귀속월도 2026-01/2026-04로 잘못
    읽혔다(실제 2026-03). 매칭이 안 될 뿐 아니라 엉뚱한 달에 금액을 붙인다.
    그래서 정밀 경로 산출물(`detail`이 있는 레코드)만 대조 소스로 받는다.
    """
    if not record.get("detail"):
        return False
    return not str(record.get("person") or "").startswith("확인필요")


def _ledger_by_company(summaries):
    totals = {}
    for s in summaries:
        key = (s["company"], s["year_month"])
        agg = totals.setdefault(key, {"people": 0, "days": 0, "manday": 0.0, "paid": 0})
        agg["people"] += 1
        agg["days"] += s["출역일수_계산"]
        agg["manday"] += s["총공수_계산"]
        agg["paid"] += s.get("실지급액") or 0
    return [
        {"company": company, "year_month": year_month, **agg, "manday": round(agg["manday"], 2)}
        for (company, year_month), agg in sorted(totals.items(), key=lambda kv: (kv[0][0] or "", kv[0][1] or ""))
    ]


@app.route("/api/labor_ledger/template/<token>")
@login_required
def api_labor_ledger_template(token):
    record = _LEDGER_RESULTS.get(token)
    if not record:
        abort(404, "정규화 결과를 찾을 수 없습니다. 다시 업로드해주세요.")
    wb = build_labor_template_workbook(record["attendance"], record["summaries"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="노무비_출역_표준양식.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/labor_ledger/crosscheck/<token>")
@login_required
def api_labor_ledger_crosscheck(token):
    record = _LEDGER_RESULTS.get(token)
    if not record or not record.get("cross") or not record["cross"][0]:
        abort(404, "대조 결과를 찾을 수 없습니다. 보험료 원장을 함께 올려 다시 시도해주세요.")
    wb = build_cross_check_workbook(*record["cross"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="출역_보험료_대조결과.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


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
        category_order=CATEGORY_ORDER,
        category_labels=CATEGORY_LABELS,
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
