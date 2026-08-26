"""공단 발급 "납부확인서/결정내역서" PDF에서 인당 레코드를 좌표 기반으로 추출한다 (이슈 #5-1).

`settlement_pdf_extractor.py`는 회사가 자체 제출하는 "납부 내역서"(순번/성명/금액이
한 행) 서식을 대상으로 "한 줄에서 n번째 금액을 집는" 방식으로 읽는다. 그런데 같은
PDF 묶음에 섞여 있는 **공단 발급 공식 서식**은 구조가 근본적으로 달라서 그 방식으로
읽으면 조용히 틀린 값이 나온다 -- 실측한 서식 구조(값은 예시로 바꿈):

    OOO       가입자부담   200,000원    26,000원
    1                     400,000원    52,000원   <- 납부보험료(총액)
    YYYY.MM.DD. 사용자부담 200,000원    26,000원

한 사람이 3줄에 걸쳐 있어서 줄 단위 파서는 이 사람을 3개 레코드로 쪼개고, 금액은
"첫 번째 금액"인 가입자부담분(200,000)만 집는다(실제 납부액은 그 2배). 장기요양보험료
컬럼은 통째로 사라진다. 국민연금 "결정내역서"도 마찬가지로 첫 금액인 **기준소득월액**
을 연금보험료로 잘못 집는다.

반면 `extract_compare.py`에는 이 공식 서식 전용의 **컬럼 x좌표 기반 정밀 엔진**이 이미
있고(대아이앤씨 166행 100% 일치로 검증됨, NOTES.md 참고), 다만 CLI 대조 전용이라
금액 컬럼만 읽고 성명/생년월일은 엑셀 쪽에서 가져다 쓴다. 이 모듈은 그 엔진에
`text_columns`(성명/생년월일)만 얹어서 **PDF 하나만으로 엑셀을 만들 수 있게** 한다.

좌표는 실측으로 확인했다(`test/` 네이티브 PDF 1페이지 기준):
  건강  성명 108.6~135.1 / 생년월일 96.7~146.8  -> 같은 열에 위아래로 놓임 (95, 152)
  연금  성명 70.8~100.8 (63,120) / 생년월일 127.9~182.9 (122,200)
        기준소득월액 230.8~275.8 (205,290)  <- 기존 CLI 설정에는 아예 없던 컬럼
"""
import re

import pdfplumber

from extract_compare import (
    HEALTH_CONFIG,
    PENSION_CONFIG,
    _detect_needs_noise_filter,
    _extract_group_records,
    _group_target_pages,
)
from hr_cost_extractor import _normalize_company, _yy_mm_to_iso

# 대조(CLI)용 설정을 그대로 두고 웹 추출에 필요한 컬럼만 얹은 사본을 쓴다.
# CLI 결과 엑셀 포맷이 바뀌지 않도록 원본 dict는 건드리지 않는다.
HEALTH_WEB_CONFIG = {
    **HEALTH_CONFIG,
    # 성명과 생년월일이 같은 x열에 위아래로 놓여 있어 한 칸으로 읽고 내용으로 가른다.
    "text_columns": {"성명_생년월일": (95, 152)},
}

PENSION_WEB_CONFIG = {
    **PENSION_CONFIG,
    "column_x_ranges": {**PENSION_CONFIG["column_x_ranges"], "기준소득월액": (205, 290)},
    "value_columns": ["기준소득월액", *PENSION_CONFIG["value_columns"]],
    "text_columns": {"성명": (63, 120), "생년월일": (122, 200)},
}

WEB_CONFIGS = {"health": HEALTH_WEB_CONFIG, "pension": PENSION_WEB_CONFIG}

# 한 사람에게서 만들어질 (구분, 금액 컬럼) 목록. 건강 PDF 한 장에서 건강보험료와
# 장기요양보험료 두 비목이 나온다 -- 기존 파이프라인은 이 중 건강만 담고 있었다.
CATEGORY_VALUE_COLUMNS = {
    "health": [("health", "건강보험료_납부"), ("longterm", "장기요양보험료_납부")],
    "pension": [("pension", "연금보험료")],
}

# 스캔본은 페이지마다 전체 OCR을 돌려야 문서 그룹을 판별할 수 있어 매우 느리다
# (실측: `sample/06-03. 제20회 기성 실적정산(건강요양).pdf` 349페이지 기준 그룹
# 판별에만 페이지당 6.05초 -> 약 35분). 동기 웹 요청에서는 감당할 수 없으므로,
# 텍스트 레이어가 없는 PDF는 이 정밀 경로를 이 페이지 수까지만 시도한다.
# 네이티브 PDF(공단 EDI 전자발급 원본)는 pdfplumber로 바로 읽어 제한이 없다.
SCAN_PAGE_BUDGET = 15

BIRTHDATE_RE = re.compile(r"(\d{4})\s*\.\s*(\d{1,2})\s*\.\s*(\d{1,2})")
NAME_RE = re.compile(r"[가-힣]{2,5}")
# "대아이앤씨(주)/일용/공주금호가스배관공사" 처럼 상호/고용형태/공사명이 슬래시로
# 이어진 사업장명칭 표기에서 맨 앞 상호만 뽑는다.
WORKPLACE_RE = re.compile(r"([^\s/]{2,})\s*/\s*[^\s/]*\s*/\s*[^\s]*")


def _parse_birthdate(text):
    m = BIRTHDATE_RE.search(text or "")
    if not m:
        return None
    year, month, day = m.groups()
    return f"{year}-{int(month):02d}-{int(day):02d}"


def _birth6(birthdate):
    """"1958-06-08" -> "580608" (주민번호 앞 6자리와 같은 형식)."""
    if not birthdate:
        return None
    year, month, day = birthdate.split("-")
    return f"{year[2:]}{month}{day}"


def _split_name_birthdate(cell_text):
    """건강보험 서식처럼 성명과 생년월일이 한 칸에 같이 잡힌 경우를 가른다."""
    birthdate = _parse_birthdate(cell_text)
    without_date = BIRTHDATE_RE.sub(" ", cell_text or "")
    name_match = NAME_RE.search(without_date)
    return (name_match.group() if name_match else None), birthdate


def _company_from_group_text(text):
    """문서 그룹 첫 페이지 텍스트에서 사업장(협력업체) 상호를 뽑는다."""
    m = WORKPLACE_RE.search(text or "")
    if m:
        return _normalize_company(m.group(1))
    return None


def _record_person(record, category):
    if category == "health":
        return _split_name_birthdate(record.get("성명_생년월일", ""))
    name = (record.get("성명") or "").replace(" ", "") or None
    return name, _parse_birthdate(record.get("생년월일", ""))


def _is_native(pdf, sample_pages=5):
    """텍스트 레이어가 있는 PDF인지(=OCR 없이 읽을 수 있는지) 판단한다."""
    return any(len(page.chars) > 0 for page in pdf.pages[:sample_pages])


def extract_confirmation(pdf_path, category, max_pages=None):
    """공단 발급 확인서 서식 페이지만 골라 인당 레코드를 뽑는다.

    반환: (records, consumed_page_indices, stats)
    이 서식이 아닌 페이지는 손대지 않고 그대로 남겨 호출부가 기존 파서로
    넘길 수 있게 한다(한 PDF에 두 서식이 섞여 있는 실적정산철 대응).
    """
    cfg = WEB_CONFIGS[category]
    records = []
    consumed = set()
    group_count = 0

    with pdfplumber.open(pdf_path) as pdf:
        if not _is_native(pdf) and len(pdf.pages) > SCAN_PAGE_BUDGET:
            # 큰 스캔본은 기존 OCR 파서에 맡긴다 (SCAN_PAGE_BUDGET 주석 참고).
            # 정밀 대조가 필요하면 CLI(`python extract_compare.py`)로 돌린다.
            return [], set(), {"confirmation_skipped": "스캔본(페이지 초과)"}

        apply_noise_filter = _detect_needs_noise_filter(pdf)
        groups = _group_target_pages(pdf, cfg)
        for group in groups:
            pages = group["page_indices"]
            if max_pages is not None:
                pages = [i for i in pages if i < max_pages]
            if not pages:
                continue
            group_count += 1
            consumed.update(pages)

            company = _company_from_group_text(group.get("company_text", ""))
            year_month = _yy_mm_to_iso(group["month"]) if group.get("month") else None
            raw = _extract_group_records(
                pdf, pages, str(pdf_path), cfg, apply_noise_filter
            )

            for row in raw:
                person, birthdate = _record_person(row, category)
                detail = {
                    key: row.get(key)
                    for key in cfg["value_columns"]
                    if row.get(key) is not None
                }
                for cat, value_column in CATEGORY_VALUE_COLUMNS[category]:
                    amount = row.get(value_column)
                    if amount is None:
                        continue
                    records.append(
                        {
                            "company": company or "미상",
                            "person": person or f"확인필요_{row.get('순번_pdf')}",
                            "birthdate": birthdate,
                            # 노무비 명세서/퇴직공제 신고에서 쓰는 주민번호 앞 6자리와
                            # 같은 형식으로도 함께 담는다 -- 이게 있으면 동명이인을
                            # 갈라서 대조할 수 있다(이슈 #5-5).
                            "birth6": _birth6(birthdate),
                            "year_month": year_month,
                            "category": cat,
                            "amount": amount,
                            "seq": row.get("순번_pdf"),
                            "needs_review": bool(row.get("needs_review")),
                            "detail": detail,
                        }
                    )

    stats = {
        "confirmation_groups": group_count,
        "confirmation_pages": len(consumed),
        "confirmation_records": len(records),
    }
    return records, consumed, stats
