"""실적정산 증빙 PDF(건강·장기요양/국민연금/퇴직공제)에서 인당 납부 레코드를 추출한다 (이슈 #2, PDF 확장).

이 PDF들은 텍스트 레이어가 없는 스캔본이지만, 기존 evidence_extractor.py가 다루던
손글씨 서명 스캔과 달리 컴퓨터로 출력된 정갈한 표라 Tesseract 한글 OCR 정확도가
꽤 높다.

건강·장기요양/국민연금 PDF는 한 페이지에 6명짜리 회사도 40명짜리 회사도 있어
표 밀도 편차가 크다. 페이지 전체를 한 번에 OCR(--psm 4)하면 Tesseract 자체 줄
분리가 밀도와 무관하게 실사용에 충분히 정확해(검증됨) 이 방식을 쓴다. 반면
퇴직공제 PDF는 한 페이지에 30~40명이 훨씬 촘촘한 여러 열로 들어가 있어 페이지
전체 OCR로는 여러 행이 한 줄로 뭉개진다. 그래서 이 문서만 가로 격자선 투영으로
행을 먼저 잘라 행 단위로 OCR한다(evidence_extractor.py와 동일한 기법).

성명·업체명 등 한글은 OCR 결과를 그대로 신뢰하되, 업체명은 오탈자가 잦아
(예: "거명이앤씨" -> "거명이엔써") 알려진 업체명 목록에 대해 유사도 매칭으로
보정한다.

**지원하지 않는 서식(의도적 스킵, skipped_pages로 개수 노출)**: 건강·연금 PDF에는
회사가 자체 제출하는 "납부 내역서"(순번/성명/금액이 한 행, 이 모듈이 지원하는
서식) 외에 국민건강보험공단/국민연금공단이 발급하는 "납부확인서"가 섞여 있다.
이 확인서는 (1) 배경에 발급 확인용 워터마크·직인이 성명 칸을 그대로 덮고 있어
Tesseract가 이름을 아예 인식하지 못하고, (2) 인당 데이터가 2행(가입자부담/
사용자부담)으로 나뉘어 있으며, (3) 페이지 하단에 해당 사업장의 합계(고지총액/
납부총액)만은 워터마크에서 벗어나 있어 깨끗하게 읽히지만, 그 위치가 그 페이지에
표시된 인원 수에 따라 달라져(고정 좌표로 자를 수 없음) 자동으로 찾으려면
단어 위치 기반 탐색이 추가로 필요하다. 이 서식이 담고 있는 것과 동일한 데이터가
이미 신뢰도 높은 엑셀 원본(`hr_cost_extractor.py`)으로 완전히 확보되어 있어,
투자 대비 효용이 낮다고 판단해 이번 버전에서는 지원하지 않는다.
"""
import difflib
import io
import re
import shutil

import numpy as np
import pymupdf
import pytesseract
from PIL import Image

from extractor import clean_amount
from hr_cost_extractor import _normalize_company

TESSERACT_CMD = shutil.which("tesseract") or r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

RENDER_DPI = 300
DARK_THRESHOLD = 200
ROW_LINE_DENSITY = 0.08
ROW_CLUSTER_GAP = 8
MIN_ROW_HEIGHT = 20

NAME_RE = re.compile(r"[가-힣]{2,5}")
# OCR이 가끔 "434, 260"처럼 콤마 뒤에 공백을 끼워 넣으므로 선택적 공백을 허용한다.
AMOUNT_RE = re.compile(r"\d{1,3}(?:,\s?\d{3})+")
TITLE_YM_RE = re.compile(r"(\d{2,4})\s*년\s*(\d{1,2})\s*월")

# 엑셀 추출 결과에서 이미 확인된 업체명(정규화 후) — OCR 오탈자 보정용 기준 목록
KNOWN_COMPANY_CORES = [
    "동남건설산업개발", "스틸뱅크", "신보", "신한에이씨티", "아이스기술",
    "엔에스컴퍼니", "이엔티파워", "지크아이디", "포유이엔지", "화랑이엔씨",
    "거명이앤씨", "거성건설", "광영건설", "금호건설", "대아이앤씨",
    "배가건설", "세안이엔씨", "삼우이앤아이", "신한ACT",
]


def _render_page(doc, page_index, dpi=RENDER_DPI):
    page = doc[page_index]
    pix = page.get_pixmap(dpi=dpi)
    return Image.open(io.BytesIO(pix.tobytes("png")))


def _cluster(values, gap=ROW_CLUSTER_GAP):
    if not values:
        return []
    groups = [[values[0]]]
    for v in values[1:]:
        if v - groups[-1][-1] <= gap:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [int(sum(g) / len(g)) for g in groups]


def _detect_row_lines(gray_arr, dark_threshold=DARK_THRESHOLD, density=ROW_LINE_DENSITY):
    dark = gray_arr < dark_threshold
    row_frac = dark.mean(axis=1)
    h = gray_arr.shape[0]
    lines = [y for y in range(h) if row_frac[y] > density]
    return _cluster(lines)


def _ocr_line(img, psm=7):
    return pytesseract.image_to_string(img, lang="kor+eng", config=f"--psm {psm}").strip()


def _year_month_from_text(text):
    m = TITLE_YM_RE.search(text)
    if not m:
        return None
    yy = int(m.group(1))
    if yy < 100:
        yy += 2000
    mm = int(m.group(2))
    if not (1 <= mm <= 12):
        return None
    return f"{yy:04d}-{mm:02d}"


def _fuzzy_company(text, min_ratio=0.6):
    norm_text = re.sub(r"\s+", "", text or "")
    if not norm_text:
        return None
    best, best_ratio = None, 0.0
    for core in KNOWN_COMPANY_CORES:
        length = len(core)
        for start in range(0, max(1, len(norm_text) - length + 1) + 3):
            window = norm_text[start:start + length + 2]
            ratio = difflib.SequenceMatcher(None, window, core).ratio()
            if ratio > best_ratio:
                best_ratio, best = ratio, core
    if best_ratio >= min_ratio:
        return _normalize_company(best)
    return None


def _page_header_text(img, top_frac=0.20):
    crop = img.crop((0, 0, img.width, int(img.height * top_frac)))
    return _ocr_line(crop, psm=6)


def _company_from_header(header_text):
    """헤더 블록에는 제목/공사명/업체명 줄이 섞여 있고, 스캔 잡음으로 앞에 가짜 줄이
    끼어들기도 해 줄 번호로는 위치를 못 믿는다. "업체명" 라벨 자체도 OCR로 자주
    깨지므로(예: "업 제 명"), 대신 각 줄의 값 부분이 알려진 업체명과 매칭되는지로
    찾는다 (공사명 등 다른 줄은 매칭될 리 없다)."""
    for line in header_text.splitlines():
        if not line.strip():
            continue
        m = re.search(r"[:：]", line)
        value = line[m.end():] if m else line
        company = _fuzzy_company(value)
        if company:
            return company
    return None


# ---------------------------------------------------------------------------
# 건강·장기요양 / 국민연금 실적정산 (인당 1행, "납부 내역서" 서식만 지원)
# ---------------------------------------------------------------------------

NON_PERSON_LINE_RE = re.compile(r"^(계|구\s*분|순\s*번|성\s*명)")


def _parse_person_lines(text):
    """OCR된 표 본문에서 (성명 or None, 첫 번째 금액) 목록을 뽑는다.
    금액이 없는 줄(빈 칸/헤더/합계 줄)은 자연히 제외된다."""
    results = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or NON_PERSON_LINE_RE.match(stripped):
            continue
        amounts = AMOUNT_RE.findall(line)
        if not amounts:
            continue
        amount = clean_amount(amounts[0])
        if amount is None:
            continue
        names = NAME_RE.findall(line)
        results.append((names[0] if names else None, amount))
    return results


def _drop_trailing_total_row(rows, tolerance=0.01):
    """"계"(합계) 표시가 OCR로 소실된 줄을 잡아낸다: 마지막 줄의 금액이 그 앞
    줄들의 합과 거의 같으면 합계 줄로 보고 제외한다 (그대로 두면 인당 금액과
    합계가 이중 집계된다)."""
    if len(rows) < 2:
        return rows
    total_candidate = rows[-1][1]
    prior_sum = sum(amount for _, amount in rows[:-1])
    if prior_sum > 0 and abs(total_candidate - prior_sum) <= prior_sum * tolerance:
        return rows[:-1]
    return rows


def extract_premium_settlement(pdf_path, category, max_pages=None):
    """category: "health" 또는 "pension". 반환은 hr_cost_extractor와 동일한 레코드 스키마.

    한 페이지에 6명짜리 회사도, 40명짜리 회사도 있어(회사 규모 편차가 큼) 행을
    직접 잘라 OCR하면 표 내부 격자선 밝기가 일정하지 않아 행 경계 검출이
    흔들린다. 대신 페이지 전체를 한 번에 OCR(--psm 4)하면 Tesseract 자체
    줄 분리가 6명이든 40명이든 실사용에 충분히 정확해(검증됨), 훨씬 빠르고
    안정적이다.
    """
    doc = pymupdf.open(pdf_path)
    n_pages = len(doc) if max_pages is None else min(len(doc), max_pages)
    records = []
    skipped_pages = 0
    processed_pages = 0

    for i in range(n_pages):
        try:
            img = _render_page(doc, i)
            text = pytesseract.image_to_string(img, lang="kor+eng", config="--psm 4")
            if "납부내역서" not in re.sub(r"\s+", "", text):
                skipped_pages += 1
                continue
            year_month = _year_month_from_text(text)
            company = _company_from_header(text)
            if not year_month or not company:
                skipped_pages += 1
                continue
            processed_pages += 1

            page_rows = _parse_person_lines(text)
            page_rows = _drop_trailing_total_row(page_rows)
            for idx, (name, amount) in enumerate(page_rows):
                records.append(
                    {
                        "company": company,
                        "person": name or f"확인필요_{i}_{idx}",
                        "year_month": year_month,
                        "category": category,
                        "amount": amount,
                    }
                )
        except Exception:
            skipped_pages += 1
            continue

    doc.close()
    return records, {"processed_pages": processed_pages, "skipped_pages": skipped_pages, "total_pages": n_pages}


# ---------------------------------------------------------------------------
# 퇴직공제부금 실적정산 (한 페이지에 여러 업체가 섞인 월별 표)
# ---------------------------------------------------------------------------

def extract_retirement_settlement(pdf_path, max_pages=None):
    doc = pymupdf.open(pdf_path)
    n_pages = len(doc) if max_pages is None else min(len(doc), max_pages)
    records = []
    skipped_pages = 0
    processed_pages = 0

    for i in range(n_pages):
        try:
            img = _render_page(doc, i)
            header_text = _page_header_text(img, top_frac=0.10)
            year_month = _year_month_from_text(header_text)
            gray = np.array(img.convert("L"))
            # 퇴직공제 표는 굵은 격자선이 매 행마다 있어 더 엄격한 기준으로도 잘 잡힌다.
            row_lines = _detect_row_lines(gray, dark_threshold=180, density=0.3)
            if not year_month or len(row_lines) < 3:
                skipped_pages += 1
                continue
            processed_pages += 1

            for j in range(1, len(row_lines) - 1):
                y0, y1 = row_lines[j], row_lines[j + 1]
                if y1 - y0 < MIN_ROW_HEIGHT:
                    continue
                row_img = img.crop((0, y0, img.width, y1))
                line = _ocr_line(row_img)
                if not line:
                    continue
                amounts = AMOUNT_RE.findall(line)
                if not amounts:
                    continue
                amount = clean_amount(amounts[-1])
                if amount is None:
                    continue
                name_match = NAME_RE.search(line)
                person = name_match.group() if name_match else f"확인필요_{i}_{j}"
                business_zone = line[name_match.end():] if name_match else line
                company = _fuzzy_company(business_zone) or "미상"
                records.append(
                    {
                        "company": company,
                        "person": person,
                        "year_month": year_month,
                        "category": "retirement",
                        "amount": amount,
                    }
                )
        except Exception:
            skipped_pages += 1
            continue

    doc.close()
    return records, {"processed_pages": processed_pages, "skipped_pages": skipped_pages, "total_pages": n_pages}


def extract_settlement_pdf(path, filename):
    """파일명 힌트로 어떤 실적정산 PDF인지 판별해서 추출한다."""
    name = filename or ""
    if "건강" in name:
        return extract_premium_settlement(path, "health")
    if "연금" in name:
        return extract_premium_settlement(path, "pension")
    if "퇴직" in name:
        return extract_retirement_settlement(path)
    raise ValueError(
        f"파일명으로 문서 종류를 판별할 수 없습니다: {filename} "
        "(파일명에 '건강', '연금', '퇴직' 중 하나가 포함되어야 합니다.)"
    )
