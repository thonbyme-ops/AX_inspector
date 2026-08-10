"""노무비 지급 내역서 / 퇴직공제부금(4대보험) 납부 신고 내역 스캔 PDF에서
행 단위 데이터를 추출한다.

이 문서들은 스캔 이미지라 텍스트 레이어가 없고, Tesseract의 한글 인식은
이 문서 폰트에서 성명 같은 짧은 한글 단어를 신뢰할 수 없는 수준으로 틀린다
(예: "차덕환" -> "자녁환"). 반면 숫자는 매우 정확하게 인식된다.

그래서 사람 이름 등 한글 필드는 OCR 텍스트로 변환하지 않고 행 이미지를
그대로 잘라 썸네일로 보여주고(감독자가 눈으로 대조), 숫자 필드(금액, 날짜,
근로일수 등)만 자동 추출한다. 표의 세로 격자선은 스캔 기울기 때문에
픽셀 투영으로는 신뢰도가 낮아, 가로 격자선(행 경계)만 자동 검출하고
각 행은 전체 폭을 그대로 잘라 숫자 화이트리스트 OCR을 돌린다.
"""
import base64
import io
import re
import shutil

import numpy as np
import pymupdf
import pytesseract
from PIL import Image

TESSERACT_CMD = shutil.which("tesseract") or r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

RENDER_DPI = 300
DARK_THRESHOLD = 180
ROW_LINE_DENSITY = 0.3
ROW_CLUSTER_GAP = 5
MAX_PAGES = 30


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


def _detect_row_lines(gray_arr):
    dark = gray_arr < DARK_THRESHOLD
    row_frac = dark.mean(axis=1)
    h = gray_arr.shape[0]
    lines = [y for y in range(h) if row_frac[y] > ROW_LINE_DENSITY]
    return _cluster(lines)


def _render_page(doc, page_index, dpi=RENDER_DPI):
    page = doc[page_index]
    pix = page.get_pixmap(dpi=dpi)
    return Image.open(io.BytesIO(pix.tobytes("png")))


THUMBNAIL_SCALE = 0.5
THUMBNAIL_JPEG_QUALITY = 62


def _row_thumbnail_base64(row_img):
    """행 이미지를 축소한 흑백 JPEG로 인코딩한다 (OCR용 원본 대신 사람이 보는 썸네일).
    300dpi 컬러 PNG 그대로 쓰면 행마다 수십 KB씩 붙어 응답이 수십 MB로 부풀어
    브라우저가 무거워지므로, 사람이 성명을 읽는 데 지장 없는 선에서 압축한다.
    색상(빨간 정정 표시)은 has_correction 플래그/배지로 이미 전달하므로
    썸네일 자체는 흑백으로 저장해 용량을 더 줄인다."""
    w, h = row_img.size
    small = row_img.convert("L").resize(
        (max(1, int(w * THUMBNAIL_SCALE)), max(1, int(h * THUMBNAIL_SCALE))),
        Image.LANCZOS,
    )
    buf = io.BytesIO()
    small.save(buf, format="JPEG", quality=THUMBNAIL_JPEG_QUALITY, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


AMOUNT_TOKEN_RE = re.compile(r"\d{1,3}(?:,\d{3})+")


def _ocr_digits_raw(img, psm, whitelist="0123456789.,-"):
    return pytesseract.image_to_string(
        img, config=f"--psm {psm} -c tessedit_char_whitelist={whitelist}"
    ).strip()


def _ocr_digits(img, whitelist="0123456789.,-"):
    """행 전체 폭 OCR. 내용이 희소한 행(합계·총계 행 등)에서는 --psm 7이
    아무 텍스트도 못 찾는 경우가 있어(자동 레이아웃 분석 실패),
    --psm 3/12로 순서대로 재시도한다."""
    for psm in (7, 3, 12):
        text = _ocr_digits_raw(img, psm, whitelist)
        if text:
            return text
    return ""


def _to_int(s):
    s = s.replace(",", "").strip()
    if not s:
        return None
    try:
        return int(s)
    except ValueError:
        return None


def _page_header_text(img, top_frac=0.12):
    h = img.height
    crop = img.crop((0, 0, img.width, int(h * top_frac)))
    text = pytesseract.image_to_string(crop, lang="kor+eng")
    return re.sub(r"\s+", " ", text).strip()


def extract_labor_cost_rows(pdf_path, max_pages=MAX_PAGES):
    """2026년 0X월분 노무비 지급 내역서: 순번/성명/은행명/계좌번호/예금주/지급금액/비고.

    성명은 신뢰할 수 없어 OCR하지 않고 행 이미지를 그대로 반환한다.
    """
    doc = pymupdf.open(pdf_path)
    rows = []
    total = None
    n_pages = min(len(doc), max_pages)

    for page_index in range(n_pages):
        img = _render_page(doc, page_index)
        gray = np.array(img.convert("L"))
        row_lines = _detect_row_lines(gray)
        if len(row_lines) < 3:
            continue  # 표가 없는 페이지(표지 등)

        for i in range(1, len(row_lines) - 1):  # i=0은 항상 헤더 행이므로 건너뜀
            y0, y1 = row_lines[i], row_lines[i + 1]
            if y1 - y0 < 20:
                continue
            row_img = img.crop((0, y0, img.width, y1))
            digits_text = _ocr_digits(row_img)

            comma_amounts = AMOUNT_TOKEN_RE.findall(digits_text)
            seq_no = None
            amount = None
            if comma_amounts:
                amount_str = max(comma_amounts, key=lambda s: _to_int(s) or 0)
                amount = _to_int(amount_str)
                prefix = digits_text.split(amount_str)[0]
                seq_m = re.search(r"\d{1,3}", prefix)
                seq_no = seq_m.group() if seq_m else None
            else:
                bare_nums = re.findall(r"\d+", digits_text)
                if bare_nums:
                    if len(bare_nums) >= 2:
                        seq_no = bare_nums[0]
                        amount = _to_int(bare_nums[-1])
                    else:
                        amount = _to_int(bare_nums[-1])

            is_last_band = i == len(row_lines) - 2
            is_total_row = is_last_band
            needs_review = amount is None
            if seq_no is None and not is_total_row and not needs_review:
                prev_seq = next(
                    (r["seq_no"] for r in reversed(rows) if r["page"] == page_index and r["seq_no"]),
                    None,
                )
                seq_no = str(int(prev_seq) + 1) if prev_seq else None
            rows.append(
                {
                    "page": page_index,
                    "seq_no": seq_no,
                    "amount": amount,
                    "row_image": _row_thumbnail_base64(row_img),
                    "is_total": is_total_row,
                    "needs_review": needs_review,
                }
            )

    doc.close()

    data_rows = [r for r in rows if not r["is_total"]]
    total_rows = [r for r in rows if r["is_total"]]
    if total_rows and total_rows[-1]["amount"] is not None:
        total = total_rows[-1]["amount"]
    else:
        total = sum(r["amount"] for r in data_rows if r["amount"] is not None)

    return {"rows": data_rows, "total": total, "pages_processed": n_pages}


BIRTHDATE_RE = re.compile(r"\d{2}\.\d{2}\.\d{2}")
YEARMONTH_RE = re.compile(r"\b\d{2}-\d{2}\b")
AMOUNT_RE = re.compile(r"\d{1,3}(?:,\d{3})+")


def _has_red_marking(img, frac_threshold=0.003):
    """원본 표에서 정정(빨간 글씨) 표시가 있는 행인지 픽셀 색상으로 판별한다.
    OCR로 부호(-)를 읽는 것은 신뢰할 수 없어, 금액 부호는 추정하지 않고
    이 플래그로 감독자가 원본 이미지를 다시 보도록 안내한다."""
    arr = np.array(img.convert("RGB")).astype(int)
    r, g, b = arr[..., 0], arr[..., 1], arr[..., 2]
    red_mask = (r > 120) & ((r - g) > 40) & ((r - b) > 40)
    return bool(red_mask.mean() > frac_threshold)


def extract_retirement_fund_rows(pdf_path, max_pages=MAX_PAGES):
    """월별 퇴직공제부금(또는 건강보험/국민연금) 납부 신고 내역:
    NO/성명/생년월일/근로년월/근로일수/직종/업체명/단위공사명/납부액.

    성명·직종·업체명은 OCR하지 않고 행 이미지를 그대로 반환한다.
    생년월일/근로년월/근로일수/납부액만 정규식으로 회수한다(최선 노력치).
    행에서 아무 숫자도 못 찾더라도 행 자체는 버리지 않고 이미지만 담아
    "확인 필요"로 남긴다 (자동 인식 실패가 곧 데이터 유실이 되지 않도록).
    금액의 +/- 부호는 OCR로 추정하지 않고, 빨간 정정 표시 유무만 플래그로 남긴다.
    """
    doc = pymupdf.open(pdf_path)
    rows = []
    n_pages = min(len(doc), max_pages)

    for page_index in range(n_pages):
        img = _render_page(doc, page_index)
        gray = np.array(img.convert("L"))
        row_lines = _detect_row_lines(gray)
        if len(row_lines) < 3:
            continue

        for i in range(1, len(row_lines) - 1):  # i=0은 헤더 행
            y0, y1 = row_lines[i], row_lines[i + 1]
            if y1 - y0 < 20:
                continue
            row_img = img.crop((0, y0, img.width, y1))
            digits_text = _ocr_digits(row_img)
            has_correction = _has_red_marking(row_img)

            birthdate_m = BIRTHDATE_RE.search(digits_text)
            amount_m = AMOUNT_RE.findall(digits_text)
            if not birthdate_m or not amount_m:
                rows.append(
                    {
                        "page": page_index,
                        "birthdate": birthdate_m.group() if birthdate_m else None,
                        "year_month": None,
                        "work_days": None,
                        "amount": None,
                        "has_correction": has_correction,
                        "needs_review": True,
                        "row_image": _row_thumbnail_base64(row_img),
                    }
                )
                continue

            remainder = digits_text[birthdate_m.end():]
            yearmonth_m = YEARMONTH_RE.search(remainder)
            work_days = None
            if yearmonth_m:
                after_ym = remainder[yearmonth_m.end():].strip()
                wd_m = re.match(r"-?\s*(\d{1,3})", after_ym)
                if wd_m:
                    work_days = _to_int(wd_m.group(1))

            amount = _to_int(amount_m[-1])
            rows.append(
                {
                    "page": page_index,
                    "birthdate": birthdate_m.group(),
                    "year_month": yearmonth_m.group() if yearmonth_m else None,
                    "work_days": work_days,
                    "amount": amount,
                    "has_correction": has_correction,
                    "needs_review": False,
                    "row_image": _row_thumbnail_base64(row_img),
                }
            )

    doc.close()
    total = sum(r["amount"] for r in rows if r["amount"] and not r["needs_review"])
    return {"rows": rows, "total": total, "pages_processed": n_pages}
