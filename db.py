"""이슈 #3: 추출된 인건비 집계를 SQLite에 저장하고, 계정별로 스코핑해서 조회한다.

성명(개인정보)은 저장하지 않는다 - hr_cost_extractor.aggregate_by_company_month()로
업체+연월+구분 집계만 만들어서 insert-only로 쌓는다. 같은 업체+연월을 다시 저장해도
새 upload로 쌓이므로 이 자체가 월별 데이터 이력이 되고, "현재 값" 조회는 각
(업체, 연월, 구분)별로 가장 최근 upload의 값을 사용한다.
"""
import json
import os
import sqlite3
from datetime import datetime, timezone

from werkzeug.security import check_password_hash, generate_password_hash

from hr_cost_extractor import aggregate_by_company_month

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "ax_inspector.db")

os.makedirs(DATA_DIR, exist_ok=True)


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_connection()
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS accounts (
                id INTEGER PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                display_name TEXT,
                company TEXT,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS uploads (
                id INTEGER PRIMARY KEY,
                account_id INTEGER NOT NULL REFERENCES accounts(id),
                filenames TEXT NOT NULL,
                uploaded_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS hr_cost_aggregates (
                id INTEGER PRIMARY KEY,
                upload_id INTEGER NOT NULL REFERENCES uploads(id),
                company TEXT NOT NULL,
                year_month TEXT NOT NULL,
                category TEXT NOT NULL,
                amount INTEGER NOT NULL
            );
            """
        )
        conn.commit()
    finally:
        conn.close()


def _now():
    return datetime.now(timezone.utc).isoformat()


def _scoped_company(account):
    """계정에 소속 업체가 지정돼 있으면 그 업체명, 없으면(관리자) None."""
    return account.get("company")


def _assert_company_access(account, company):
    scoped = _scoped_company(account)
    if scoped and company != scoped:
        raise PermissionError("해당 업체 데이터에 접근할 권한이 없습니다.")


def create_account(username, password, display_name=None, company=None):
    conn = get_connection()
    try:
        conn.execute(
            "INSERT INTO accounts (username, password_hash, display_name, company, created_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (username, generate_password_hash(password), display_name, company, _now()),
        )
        conn.commit()
    finally:
        conn.close()


def list_accounts():
    conn = get_connection()
    try:
        rows = conn.execute(
            "SELECT id, username, display_name, company, created_at FROM accounts ORDER BY id"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def verify_login(username, password):
    conn = get_connection()
    try:
        row = conn.execute("SELECT * FROM accounts WHERE username = ?", (username,)).fetchone()
        if row and check_password_hash(row["password_hash"], password):
            return dict(row)
        return None
    finally:
        conn.close()


def save_upload(account, filenames, records):
    """records: hr_cost_extractor 스타일 레코드(company/person/year_month/category/amount) 목록.
    성명은 버리고 업체+연월+구분 집계만 저장한다. 계정에 소속 업체가 있으면 그 업체 데이터만
    저장하고 나머지는 조용히 걸러낸다."""
    scoped = _scoped_company(account)
    agg_rows = aggregate_by_company_month(records)
    if scoped:
        agg_rows = [r for r in agg_rows if r["company"] == scoped]
    if not agg_rows:
        raise ValueError("저장할 데이터가 없습니다. (소속 업체 데이터와 일치하지 않을 수 있습니다.)")

    conn = get_connection()
    try:
        cur = conn.execute(
            "INSERT INTO uploads (account_id, filenames, uploaded_at) VALUES (?, ?, ?)",
            (account["id"], json.dumps(filenames, ensure_ascii=False), _now()),
        )
        upload_id = cur.lastrowid
        conn.executemany(
            "INSERT INTO hr_cost_aggregates (upload_id, company, year_month, category, amount) "
            "VALUES (?, ?, ?, ?, ?)",
            [(upload_id, r["company"], r["year_month"], r["category"], r["amount"]) for r in agg_rows],
        )
        conn.commit()
        return upload_id
    finally:
        conn.close()


def list_companies(account):
    scoped = _scoped_company(account)
    if scoped:
        return [scoped]
    conn = get_connection()
    try:
        rows = conn.execute("SELECT DISTINCT company FROM hr_cost_aggregates ORDER BY company").fetchall()
        return [r["company"] for r in rows]
    finally:
        conn.close()


def list_year_months(account, company):
    _assert_company_access(account, company)
    conn = get_connection()
    try:
        rows = conn.execute(
            "SELECT DISTINCT year_month FROM hr_cost_aggregates WHERE company = ? ORDER BY year_month",
            (company,),
        ).fetchall()
        return [r["year_month"] for r in rows]
    finally:
        conn.close()


def get_latest_by_month(account, company, months):
    """각 (업체, 연월, 구분)별로 가장 최근 upload에 저장된 값만 반환한다."""
    _assert_company_access(account, company)
    if not months:
        return []
    conn = get_connection()
    try:
        placeholders = ",".join("?" for _ in months)
        rows = conn.execute(
            f"""
            SELECT a.year_month, a.category, a.amount
            FROM hr_cost_aggregates a
            WHERE a.company = ?
              AND a.year_month IN ({placeholders})
              AND a.upload_id = (
                  SELECT MAX(a2.upload_id) FROM hr_cost_aggregates a2
                  WHERE a2.company = a.company AND a2.year_month = a.year_month AND a2.category = a.category
              )
            """,
            (company, *months),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_upload_history(account, company, year_month):
    """해당 업체+연월의 모든 업로드 이력을 최신순으로 반환한다 (감사 추적용)."""
    _assert_company_access(account, company)
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT u.id AS upload_id, u.filenames, u.uploaded_at, SUM(a.amount) AS total_amount
            FROM uploads u
            JOIN hr_cost_aggregates a ON a.upload_id = u.id
            WHERE a.company = ? AND a.year_month = ?
            GROUP BY u.id
            ORDER BY u.uploaded_at DESC
            """,
            (company, year_month),
        ).fetchall()
        result = []
        for r in rows:
            d = dict(r)
            d["filenames"] = json.loads(d["filenames"])
            result.append(d)
        return result
    finally:
        conn.close()
