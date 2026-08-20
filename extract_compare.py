"""보험료 대조 파이프라인 (이슈 #2, #3) -- 통합 버전 (v3 + v4 + 국민연금 신규 추가).

기존 v3(Acrobat OCR 스캔본)와 v4(네이티브 PDF)는 로직이 거의 동일하고 상수만
달랐다. 이 파일은 그 공통 엔진(회사/월 자동 매칭, 엑셀 읽기, 비교, 결과 저장,
숨은 폰트 필터링)을 하나로 합치고, 문서 종류별로 다른 부분(표 판별 키워드,
컬럼 좌표, 사람당 줄 수, 엑셀 컬럼 매핑)만 `DOCUMENT_TYPES` 설정으로 뽑아냈다.

실행:
    python extract_compare.py --doc-type health --sheet-filter 대아
    python extract_compare.py --doc-type pension --pdf "대아이앤씨 국민연금 납부확인서 26.04.pdf"
    python extract_compare.py --doc-type all              # 문서 종류 전부 순서대로 처리 + 결과 합침

숨은 폰트 필터링에 대해 (NOTES.md 참고, 2026-08-14 발견):
같은 "Hidden" 계열 폰트라도 문서 종류에 따라 정반대 의미다 -- 스캔본(Acrobat
OCR)은 Hidden 텍스트가 유일한 본문(78%)이라 걸러내면 안 되고, 네이티브 PDF는
Hidden이 눈에 보이는 진짜 텍스트 위에 겹친 중복 오염 레이어라 걸러내야 한다.
그래서 필터 적용 여부도 문서 종류 설정(`apply_noise_font_filter`)으로 분기한다
-- 스캔본 계열 문서 종류를 추가할 땐 반드시 False로 둘 것.

국민연금 관련 알려진 한계: 국민연금 엑셀 시트("_연금(YY.MM)")는 건강보험
시트와 달리 2행짜리 그룹 헤더에, 심지어 회사가 여러 하위 공사(예: 대아이앤씨의
"플랜트배관공사"/"주기기설치공사")를 진행하면 "사용자부담금" 열 자체가
공사별로 나뉘어 있다(사람마다 그중 한 열에만 값이 있고 나머지는 빈칸). 이
파일의 `_extract_excel_pension`은 그 여러 "사용자부담금" 열을 라벨로 찾아
비어있지 않은 값을 취하는 방식으로 대응했지만, 대아 시트 하나로만 검증했다
-- 다른 회사 시트에서는 레이아웃이 다를 수 있으니 실제 비교 전 반드시
샘플 검증할 것.
"""
import argparse
import re
import statistics
import time
from pathlib import Path

import pandas as pd
import pdfplumber
from openpyxl import load_workbook

from extract_compare_v2 import (
    _find_header_idx,
    _find_single_file,
    _first_amount,
    _first_grouped_amount,
    clean_amount,
    step5_save_result,
)

BASE_DIR = Path(__file__).resolve().parent
RAW_DIR = BASE_DIR / "raw"
OCR_DONE_DIR = BASE_DIR / "ocr_done"
RESULT_DIR = BASE_DIR / "result"


# ---------------------------------------------------------------------------
# 공통: 숨은/오염 폰트 필터링 (NOTES.md 참고 -- 스캔본에는 절대 적용 금지)
# ---------------------------------------------------------------------------

NOISE_FONTS = {"Helvetica", "Times-Roman", "Courier", "Symbol", "ZapfDingbats"}


def _is_real_content_char(obj):
    """실측(page.chars) 결과: 정상 데이터는 서브셋 임베드 폰트(예: 'LKBAPO+Dotum')로
    일관되게 나오고, 오염된 글자들은 'Helvetica'/'Times-Roman'/'HiddenHorzOCR' 같은
    PDF 표준 범용 폴백 폰트나 이름에 "Hidden"이 들어간 폰트로 나온다. 이 폰트들은
    파일마다 달라질 수 있는 서브셋 폰트명이 아니라 표준 폰트라 이름으로 안전하게
    골라낼 수 있다.

    스캔본(Acrobat OCR)에서는 이 필터를 쓰면 안 된다 -- 그 문서는 Hidden 폰트가
    유일한 본문이라 필터링하면 실제 데이터 대부분이 사라진다(NOTES.md 실측:
    78%). 적용 여부는 문서 종류가 아니라 `_detect_needs_noise_filter`로 PDF
    마다 실제 폰트 구성을 보고 자동으로 정한다(아래 참고) -- 필터가 필요
    없다고 판단된 PDF에는 이 함수를 호출하지 않는다.
    """
    if obj.get("object_type") != "char":
        return True
    fontname = obj.get("fontname", "")
    if "Hidden" in fontname:
        return False
    if fontname in NOISE_FONTS:
        return False
    return True


def _detect_needs_noise_filter(pdf, sample_pages=3):
    """페이지 몇 개의 page.chars 폰트 분포를 보고 Hidden/노이즈 폰트를
    걸러내야 하는지 PDF마다 자동으로 판단한다(NOTES.md 체크리스트의 "1. 폰트
    분포부터 찍어본다"를 코드로 자동화).

    이걸 "문서 종류"(health/pension) 단위 고정값으로 두면 안 된다는 걸 실측
    으로 확인했다 -- 같은 "health" 종류라도 Acrobat OCR 스캔본은 Hidden이
    유일한 본문이고, 네이티브 발급 PDF는 Hidden이 진짜 텍스트 위에 겹친
    중복 오염층이라 정반대 처리가 필요하다(문서 종류가 아니라 "이 PDF가
    어떻게 만들어졌는가"에 달린 문제). 그래서 doc_type 설정이 아니라 실제
    로드된 PDF 각각에 대해 이 함수로 판단한다.

    실측: 스캔본은 서브셋 임베드 폰트("XXXXXX+폰트명" 형태)가 전혀 없이
    Hidden 폰트만 있었고(embedded_other 비율 0%), 네이티브 오염 문서는
    Hidden과 별개로 진짜 임베드 폰트가 이미 압도적 비중(88%)이었다 -- 10%를
    기준으로 삼으면 두 경우 다 안전하게 갈린다.
    """
    hidden_count = 0
    embedded_other_count = 0
    total = 0
    for page in pdf.pages[:sample_pages]:
        for c in page.chars:
            total += 1
            fontname = c.get("fontname", "")
            if "Hidden" in fontname:
                hidden_count += 1
            elif "+" in fontname:
                embedded_other_count += 1
    if total == 0 or hidden_count == 0:
        return False  # Hidden 폰트 자체가 없으면 필터링할 게 없음
    return (embedded_other_count / total) > 0.10


# ---------------------------------------------------------------------------
# 공통: 발급번호/월 파싱, 페이지 그룹핑 (문서 종류 무관)
# ---------------------------------------------------------------------------

MONTH_RE = re.compile(r"(\d{4})년\s*(\d{1,2})월")
ISSUE_NO_RE = re.compile(r"발\s*급\s*번?\s*호\s*:?\s*([A-Za-z0-9\-‐-―一]{5,})")


def _normalize_month(text):
    m = MONTH_RE.search(text)
    if not m:
        return None
    year, month = int(m.group(1)), int(m.group(2))
    return f"{year % 100:02d}.{int(month):02d}"


def _group_target_pages(pdf, cfg):
    """문서 종류의 `is_target_page`를 통과하는 페이지를 순서대로 훑어
    "문서 그룹"(회사x월 하나)으로 묶는다 -- v3의 _group_format_b_pages와 동일한
    알고리즘이되 페이지 판별 조건만 설정에서 받는다.

    같은 문서인데 페이지마다 발급번호를 반복 인쇄하는 경우(실측 확인)와,
    첫 페이지에만 있고 이후 페이지는 이어지는 경우가 섞여 있어서 발급번호는
    "숫자만 남긴 값"으로 비교하고(대시 문자가 페이지마다 "-"/"–"/"—"/"一"로
    제각각 인식됨), 짧게 잘린 저신뢰 값은 새 문서로 믿지 않고 직전 그룹에
    이어붙인다.
    """
    groups = []
    prev_was_target = False
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        if not cfg["is_target_page"](text):
            prev_was_target = False
            continue

        issue_match = ISSUE_NO_RE.search(text)
        issue_no = issue_match.group(1) if issue_match else None
        issue_digits = re.sub(r"\D", "", issue_no) if issue_no else None
        has_header = issue_no is not None
        issue_reliable = issue_digits is not None and len(issue_digits) >= 10

        same_doc_as_prev = (
            prev_was_target
            and groups
            and ((not issue_reliable) or (issue_digits == groups[-1]["issue_digits"]))
        )
        if same_doc_as_prev:
            groups[-1]["page_indices"].append(i)
        else:
            groups.append(
                {
                    "company_text": text if has_header else "",
                    "month": _normalize_month(text) if has_header else None,
                    "issue_no": issue_no,
                    "issue_digits": issue_digits,
                    "page_indices": [i],
                    "needs_review": not has_header,
                }
            )
        prev_was_target = True
    return groups


# ---------------------------------------------------------------------------
# 공통: 좌표 기반 블록 추출 (컬럼/줄 구조는 doc_type 설정에서 받음)
# ---------------------------------------------------------------------------


def _cluster_lines(words, gap):
    ws = sorted(words, key=lambda w: w["top"])
    clusters = []
    for w in ws:
        if clusters and w["top"] - clusters[-1][-1]["top"] <= gap:
            clusters[-1].append(w)
        else:
            clusters.append([w])
    return clusters


def _find_table_body_bounds(words, page_height, cfg):
    header_word = next((w for w in words if w["text"] == cfg["header_word"]), None)
    top = (header_word["top"] + cfg["header_top_margin"]) if header_word else 0

    footer_markers = cfg.get("footer_markers") or ()
    footer_candidates = [
        w for w in words if w["top"] > top and any(m in w["text"] for m in footer_markers)
    ]
    footer_word = min(footer_candidates, key=lambda w: w["top"], default=None)

    if footer_word is not None:
        return top, footer_word["top"] - cfg["footer_bottom_margin"], True
    return top, page_height, False


def _typical_row_gap(clusters, cfg):
    tops = sorted(min(w["top"] for w in c) for c in clusters)
    gaps = [b - a for a, b in zip(tops, tops[1:])]
    lo, hi = cfg["min_plausible_subrow_gap"], cfg["max_plausible_subrow_gap"]
    plausible = [g for g in gaps if lo <= g <= hi]
    return statistics.median(plausible) if plausible else cfg["default_subrow_height"]


def _first_amount_safe(text):
    """v2의 _first_grouped_amount(쉼표 3자리 그룹만 매치)에 "0원"(쉼표 없어
    원래 정규식이 못 잡는 값) 예외만 추가한다. 크롭 OCR이 아니라 이미 인식된
    깨끗한 단어 텍스트를 읽는 것이라 "0"/"0원" 토큰만 좁게 허용해도 안전하다
    (v4에서 실측: 이걸 안 하면 0원인 사람이 needs_review 없이 통째로 사라짐)."""
    grouped = _first_grouped_amount(text)
    if grouped is not None:
        return grouped
    for token in text.split():
        if token.rstrip("원") == "0":
            return 0
    return None


def _words_in_box(words, x0, x1, y0, y1):
    return [w for w in words if x0 <= w["x0"] < x1 and y0 <= w["top"] < y1]


def _extract_group_records(pdf, page_indices, source_label, cfg, apply_noise_filter):
    """한 문서 그룹(같은 회사x월, 페이지 1개 이상)에서 사람별 레코드를 뽑는다.

    v3/v4의 _extract_group_records와 같은 알고리즘(동적 줄 클러스터링 -> 블록
    경계 계산 -> 블록별 순번/금액 파싱)이되, 컬럼 x범위·줄 간격·블록 높이
    배수 등을 전부 doc_type 설정에서 받아 문서 종류에 상관없이 동작한다.
    "block_height_multiplier"가 그 차이를 흡수한다 -- 건강보험처럼 사람당
    여러 줄(가입자부담/사용자부담)이 겹쳐 있으면 typical 줄 간격의 2배가
    한 블록 높이지만, 국민연금처럼 사람당 정확히 한 줄이면 배수가 1이라
    typical 줄 간격이 곧 블록 높이다.

    apply_noise_filter는 doc_type 설정이 아니라 이 PDF에 대해
    _detect_needs_noise_filter가 실제로 판단한 값을 호출부(run_doc_type)가
    넘겨준다 -- 문서 종류로 고정하면 안 되는 이유는 _is_real_content_char
    docstring 참고.
    """
    x_ranges = cfg["column_x_ranges"]
    value_columns = cfg["value_columns"]

    records = []
    prev_seq = None
    for page_index in page_indices:
        page = pdf.pages[page_index]
        if apply_noise_filter:
            page = page.filter(_is_real_content_char)
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        if not words:
            continue

        body_top, body_bottom, footer_confident = _find_table_body_bounds(
            words, page.height, cfg
        )
        body_words = [w for w in words if body_top <= w["top"] <= body_bottom]
        if not body_words:
            continue

        clusters = _cluster_lines(body_words, cfg["line_cluster_gap"])
        if not clusters:
            continue

        block_height = cfg["block_height_multiplier"] * _typical_row_gap(clusters, cfg)
        cluster_tops = sorted(min(w["top"] for w in c) for c in clusters)

        block_starts = [cluster_tops[0]]
        y = cluster_tops[0]
        while True:
            next_tops = [t for t in cluster_tops if t > y + block_height * 0.8]
            if not next_tops:
                break
            y = next_tops[0]
            block_starts.append(y)
        block_starts.append(body_bottom)

        consecutive_empty = 0
        block_count = 0
        max_blocks_no_footer = cfg["max_blocks_no_footer"]
        for i in range(len(block_starts) - 1):
            if block_starts[i] >= body_bottom or consecutive_empty >= cfg["max_consecutive_empty_blocks"]:
                break
            if not footer_confident and block_count >= max_blocks_no_footer:
                break
            block_count += 1

            y0 = block_starts[i] - cfg["block_pad_above"]
            y1 = block_starts[i + 1]

            seq_words = _words_in_box(
                words, *x_ranges["순번_pdf"], y0 - cfg["seq_extra_pad"], y1 + cfg["seq_extra_pad"]
            )
            seq_text = " ".join(w["text"] for w in seq_words)
            seq_no = _first_amount(seq_text)

            values = {}
            for key in value_columns:
                cell_words = _words_in_box(words, *x_ranges[key], y0, y1)
                cell_text = " ".join(w["text"] for w in cell_words)
                values[key] = _first_amount_safe(cell_text)

            all_values_empty = all(v is None for v in values.values())
            if seq_no is None and all_values_empty:
                consecutive_empty += 1
                continue
            consecutive_empty = 0
            if seq_no is not None and all_values_empty:
                continue  # 요약/합계 섹션에서 새어든 순번 잡음으로 추정 -- 기록 안 함

            amount_consistent = True
            for a_key, b_key in cfg.get("amount_consistency_pairs", []):
                a, b = values.get(a_key), values.get(b_key)
                if a is not None and b is not None and b != 2 * a:
                    amount_consistent = False

            needs_review = (
                seq_no is None
                or any(v is None for v in values.values())
                or (prev_seq is not None and seq_no != prev_seq + 1)
                or not amount_consistent
            )
            if seq_no is not None:
                prev_seq = seq_no
            records.append(
                {
                    "순번_pdf": seq_no,
                    **values,
                    "needs_review": needs_review,
                    "출처파일": source_label,
                }
            )
    return records


def _dedupe_by_seq(df, value_columns):
    if df.empty:
        return df
    with_seq = df[df["순번_pdf"].notna()].copy()
    without_seq = df[df["순번_pdf"].isna()]
    if not with_seq.empty:
        with_seq["_missing"] = with_seq[value_columns].isna().sum(axis=1)
        with_seq = with_seq.sort_values("_missing").drop_duplicates("순번_pdf", keep="first")
        with_seq = with_seq.drop(columns="_missing")
    return pd.concat([with_seq, without_seq], ignore_index=True)


# ---------------------------------------------------------------------------
# 건강·장기요양보험료 문서 종류
# ---------------------------------------------------------------------------


def _is_health_target_page(text):
    """공단 발급 공식 "납부확인서"(가입자부담/사용자부담 2줄 구조) 페이지인지
    판별. OCR 노이즈로 "가입자부담"이 "?f입자부담" 등으로 깨져도 "입자부담"
    부분은 349페이지 전수 조사에서 항상 살아남았다."""
    return ("입자부담" in text) and ("용자부담" in text)


def _extract_excel_health(wb, sheet_name):
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = _find_header_idx(rows)
    if header_idx is None:
        raise ValueError(f"'순번/성명' 헤더 행을 찾을 수 없습니다: (시트: {sheet_name})")
    records = []
    for row in rows[header_idx + 1 :]:
        seq_no, name = row[0], row[1]
        if not isinstance(seq_no, (int, float)) or not name:
            break
        records.append(
            {
                "순번_엑셀": int(seq_no),
                "성명": str(name).strip(),
                "건강보험료_엑셀": clean_amount(row[2]),
                "장기요양보험료_엑셀": clean_amount(row[3]),
            }
        )
    return pd.DataFrame(records)


HEALTH_CONFIG = {
    "label": "건강·장기요양보험료",
    "sheet_marker": "_건강(",
    "is_target_page": _is_health_target_page,
    "header_word": "순번",
    "header_top_margin": 20,
    "footer_bottom_margin": 10,
    # "사용목적"이 Acrobat OCR에서 한 토큰으로 깨끗하게 인식돼(예전 Tesseract는
    # 음절 단위로 쪼갰음) 1순위 마커로 쓴다. 나머지는 폴백.
    "footer_markers": ("사용목적", "납부확인용", "합계", "고지", "인원"),
    "max_blocks_no_footer": 60,
    "line_cluster_gap": 10,
    "block_pad_above": 8,
    "seq_extra_pad": 3,
    "block_height_multiplier": 2,  # 가입자부담/사용자부담 두 줄이 한 블록
    "min_plausible_subrow_gap": 3,
    "max_plausible_subrow_gap": 20,
    "default_subrow_height": 13,
    "max_consecutive_empty_blocks": 2,
    "column_x_ranges": {
        "순번_pdf": (50, 95),
        "건강보험료_고지": (230, 310),
        "장기요양보험료_고지": (300, 390),
        "건강보험료_납부": (395, 465),
        "장기요양보험료_납부": (465, 545),
    },
    "value_columns": ["건강보험료_고지", "장기요양보험료_고지", "건강보험료_납부", "장기요양보험료_납부"],
    # "납부"(가운데 줄, 합계)는 항상 "고지"(각 절반)의 정확히 2배다 -- 실측
    # 결과 Acrobat/네이티브 OCR이 드물게 큰 금액의 앞자리를 통째로 빠뜨리는
    # 경우가 있었는데(예: "1,108,680원"->",108,680원"), 이 내부 일관성이
    # 깨지면 needs_review로 잡는다.
    "amount_consistency_pairs": [
        ("건강보험료_고지", "건강보험료_납부"),
        ("장기요양보험료_고지", "장기요양보험료_납부"),
    ],
    "extract_excel": _extract_excel_health,
    "compare_pairs": [
        ("건강보험료_엑셀", "건강보험료_납부", "건강보험료"),
        ("장기요양보험료_엑셀", "장기요양보험료_납부", "장기요양보험료"),
    ],
    # 숨은 폰트 필터 적용 여부는 doc_type 고정값이 아니라 PDF마다
    # _detect_needs_noise_filter로 자동 판단한다(run_doc_type 참고) --
    # 이 종류에 스캔본과 네이티브 PDF가 둘 다 올 수 있어서다.
}


# ---------------------------------------------------------------------------
# 국민연금 문서 종류 (신규)
# ---------------------------------------------------------------------------


def _is_pension_target_page(text):
    """국민연금보험료 결정내역서의 "연금보험료 대상자" 표 페이지 판별.
    실측(대아이앤씨 26.04, 5페이지): 첫 페이지에만 발급번호/집계 섹션이
    있고 이후 페이지는 표 헤더("순번 성명 생년월일 기준소득월액 연금보험료
    근로자기여금 사용자부담금")만 반복하며 이어진다 -- "연금보험료"와
    "기준소득월액"이 둘 다 있으면 대상자 표가 있는 페이지로 본다."""
    return ("연금보험료" in text) and ("기준소득월액" in text)


def _extract_excel_pension(wb, sheet_name):
    """국민연금 엑셀 시트를 읽는다.

    건강보험 시트와 달리 2행짜리 그룹 헤더고(예: "납부보험료"가 상위 그룹,
    그 아래 실제 "사용자부담금"/"노무비 지급내역서 No." 열들이 옴), 회사가
    여러 하위 공사를 진행하면 "사용자부담금" 열 자체가 공사별로 나뉘어
    있다(대아이앤씨 26.04 실측: "플랜트배관공사"/"주기기설치공사" 각각의
    사용자부담금 열이 따로 있고, 사람마다 그중 한 열에만 값이 있음).

    그래서 "사용자부담금"이라는 라벨을 가진 열을 전부 찾아서, 각 행마다 그중
    비어있지 않은 첫 값을 취한다(한 사람이 동시에 여러 공사에서 일하지
    않는다고 전제). "연금보험료"(총액) 열은 헤더 행에 직접 라벨이 없고 한
    행 위의 그룹 헤더("납부보험료")에 있어서, 순번/성명 바로 다음(성명 뒤
    첫 번째, "사용자부담금" 라벨이 아닌) 열로 찾는다.

    주의: 대아이앤씨 시트 하나로만 검증했다 -- 다른 회사 시트는 하위 공사가
    하나뿐이라 레이아웃이 더 단순하거나, 반대로 또 다르게 생겼을 수 있다.
    실제 비교 전 반드시 대상 시트로 샘플 검증할 것.
    """
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = _find_header_idx(rows)
    if header_idx is None:
        raise ValueError(f"'순번/성명' 헤더 행을 찾을 수 없습니다: (시트: {sheet_name})")

    header_row = rows[header_idx]
    burden_cols = [i for i, v in enumerate(header_row) if isinstance(v, str) and "사용자부담금" in v]
    if not burden_cols:
        raise ValueError(f"'사용자부담금' 열을 찾을 수 없습니다: (시트: {sheet_name})")
    # "연금보험료"(총액) 열: 성명(1) 바로 다음 칸부터 "사용자부담금"이 아닌 첫 열
    pension_col = next((i for i in range(2, max(burden_cols) + 1) if i not in burden_cols), 2)

    records = []
    for row in rows[header_idx + 1 :]:
        seq_no, name = row[0], row[1]
        if not isinstance(seq_no, (int, float)) or not name:
            break
        burden = next((clean_amount(row[i]) for i in burden_cols if row[i] not in (None, "")), None)
        records.append(
            {
                "순번_엑셀": int(seq_no),
                "성명": str(name).strip(),
                "연금보험료_엑셀": clean_amount(row[pension_col]) if pension_col < len(row) else None,
                "사용자부담금_엑셀": burden,
            }
        )
    return pd.DataFrame(records)


PENSION_CONFIG = {
    "label": "국민연금",
    "sheet_marker": "_연금(",
    "is_target_page": _is_pension_target_page,
    "header_word": "순번",
    "header_top_margin": 15,
    "footer_bottom_margin": 10,
    # 실측 결과 대상자 표 페이지 자체에는 뚜렷한 마커 텍스트가 안 보였다
    # (건강보험과 달리 요약 섹션이 첫 페이지에만 따로 있음) -- 표시 못 찾으면
    # max_blocks_no_footer 폴백에 의존한다. 한 페이지 최대 인원 관측치(약
    # 24명)보다 넉넉하게 잡는다.
    "footer_markers": (),
    "max_blocks_no_footer": 40,
    "line_cluster_gap": 10,
    "block_pad_above": 3,
    "seq_extra_pad": 3,
    "block_height_multiplier": 1,  # 사람당 정확히 한 줄
    "min_plausible_subrow_gap": 15,
    "max_plausible_subrow_gap": 35,
    "default_subrow_height": 24,
    "max_consecutive_empty_blocks": 2,
    "column_x_ranges": {
        "순번_pdf": (35, 62),
        "연금보험료": (295, 365),
        "근로자기여금": (380, 460),
        "사용자부담금": (478, 560),
    },
    "value_columns": ["연금보험료", "근로자기여금", "사용자부담금"],
    "amount_consistency_pairs": [],  # 근로자기여금==사용자부담금이지만 필수 관계로 강제하지 않음
    "extract_excel": _extract_excel_pension,
    "compare_pairs": [
        ("연금보험료_엑셀", "연금보험료", "연금보험료"),
        ("사용자부담금_엑셀", "사용자부담금", "사용자부담금"),
    ],
    # 숨은 폰트 필터는 여기서도 doc_type 고정값이 아니라 PDF마다 자동 판단
    # (HEALTH_CONFIG 주석 참고).
}


DOCUMENT_TYPES = {
    "health": HEALTH_CONFIG,
    "pension": PENSION_CONFIG,
}


# ---------------------------------------------------------------------------
# 공통: 회사·월 자동 매칭 + 대조 + 결과 저장
# ---------------------------------------------------------------------------

SHEET_RE_TEMPLATE = r"^(.+?)_{marker}\((\d{{2}}\.\d{{2}})\)$"


def _find_matching_groups(company_abbrev, month, groups):
    return [g for g in groups if g["month"] == month and company_abbrev in g["company_text"]]


def _compare_generic(excel_df, pdf_df, compare_pairs):
    """문서 종류의 compare_pairs(엑셀컬럼, pdf컬럼, 표시이름) 목록을 기반으로
    순번 매칭 + 항목별 일치 판정을 한다 -- v2의 step4_compare을 문서 종류
    무관하게 일반화한 버전."""
    pdf_df = pdf_df.drop(columns=["성명"], errors="ignore")
    merged = pd.merge(excel_df, pdf_df, left_on="순번_엑셀", right_on="순번_pdf", how="outer")

    columns = ["성명", "순번_엑셀", "순번_pdf"]
    for excel_col, pdf_col, label in compare_pairs:
        diff_col = f"{label}_차이"
        merged[diff_col] = merged[excel_col] - merged[pdf_col]
        columns += [excel_col, pdf_col, diff_col]

    def _judge(row):
        has_excel = pd.notna(row[compare_pairs[0][0]])
        has_pdf = pd.notna(row[compare_pairs[0][1]])
        if not has_excel:
            return "엑셀_누락"
        if not has_pdf:
            return "PDF_누락"
        if row.get("needs_review"):
            return "확인필요"
        all_ok = all(row[excel_col] == row[pdf_col] for excel_col, pdf_col, _ in compare_pairs)
        return "일치" if all_ok else "불일치"

    merged["판정"] = merged.apply(_judge, axis=1)
    columns += ["판정", "needs_review", "출처파일"]
    for col in columns:
        if col not in merged.columns:
            merged[col] = None
    return merged[columns].sort_values(["순번_엑셀", "순번_pdf"]).reset_index(drop=True)


def run_doc_type(doc_type, excel_path, pdf_path, sheet_filter=None):
    """한 문서 종류(health/pension)에 대해 PDF 그룹 빌드 -> 엑셀 시트 매칭 ->
    대조까지 전부 수행하고, 합쳐진 결과 DataFrame과 매칭 요약을 돌려준다."""
    cfg = DOCUMENT_TYPES[doc_type]
    print(f"[{doc_type}] PDF 문서 그룹 빌드 중: {pdf_path.name}", flush=True)
    t0 = time.time()
    with pdfplumber.open(pdf_path) as pdf:
        apply_noise_filter = _detect_needs_noise_filter(pdf)
        print(
            f"      숨은/노이즈 폰트 필터: {'적용' if apply_noise_filter else '미적용'} "
            f"(이 PDF의 실제 폰트 구성으로 자동 판단)",
            flush=True,
        )
        groups = _group_target_pages(pdf, cfg)
        print(f"      {cfg['label']} 문서 그룹 {len(groups)}개 발견 ({time.time() - t0:.1f}초)", flush=True)
        for g in groups:
            g["records"] = _extract_group_records(
                pdf, g["page_indices"], pdf_path.name, cfg, apply_noise_filter
            )

    wb = load_workbook(excel_path, data_only=True, read_only=True)
    sheet_names = [s for s in wb.sheetnames if cfg["sheet_marker"] in s]
    if sheet_filter:
        sheet_names = [s for s in sheet_names if sheet_filter in s]
    print(f"      대상 시트 {len(sheet_names)}개 ('{cfg['sheet_marker']}' 포함)", flush=True)

    sheet_re = re.compile(SHEET_RE_TEMPLATE.format(marker=cfg["sheet_marker"].strip("_(")))
    value_columns = cfg["value_columns"]

    all_results = []
    matched_count = 0
    unmatched_sheets = []
    for sheet_name in sheet_names:
        m = sheet_re.match(sheet_name)
        if not m:
            continue
        company_abbrev, month = m.group(1), m.group(2)
        # 국민연금 시트는 하위 공사 구분이 회사약칭 뒤에 괄호로 붙는다(예:
        # "대아(주기기,배관)_연금(26.04)") -- 그 괄호 부분은 엑셀 내부
        # 표기일 뿐 PDF 회사명 텍스트에는 나오지 않으므로, 매칭에는 괄호
        # 앞의 순수 약칭만 쓴다.
        company_key = company_abbrev.split("(")[0]

        matched_groups = _find_matching_groups(company_key, month, groups)
        if not matched_groups:
            unmatched_sheets.append(sheet_name)
            continue
        matched_count += 1

        excel_df = cfg["extract_excel"](wb, sheet_name)
        all_records = []
        for g in matched_groups:
            for r in g["records"]:
                if g["needs_review"]:
                    r = {**r, "needs_review": True}
                all_records.append(r)
        pdf_df = pd.DataFrame(all_records, columns=["순번_pdf", *value_columns, "needs_review", "출처파일"])
        pdf_df = _dedupe_by_seq(pdf_df, value_columns) if not pdf_df.empty else pdf_df

        result = _compare_generic(excel_df, pdf_df, cfg["compare_pairs"])
        result.insert(0, "문서종류", cfg["label"])
        result.insert(1, "시트명", sheet_name)
        all_results.append(result)

    print(
        f"[{doc_type}] 완료: {matched_count}개 시트 매칭, {len(unmatched_sheets)}개 시트 PDF_문서없음 "
        f"({time.time() - t0:.1f}초)",
        flush=True,
    )
    if unmatched_sheets:
        print(f"      PDF 매칭 실패 시트: {unmatched_sheets}", flush=True)

    combined = pd.concat(all_results, ignore_index=True) if all_results else pd.DataFrame()
    return combined


def run(doc_types, excel_path, pdf_paths, result_dir, sheet_filter=None):
    """doc_types 목록(예: ["health"] 또는 ["health","pension"])을 순서대로
    처리해 결과를 하나로 합쳐 저장한다. pdf_paths는 {doc_type: Path} 매핑."""
    all_combined = []
    for doc_type in doc_types:
        pdf_path = pdf_paths[doc_type]
        combined = run_doc_type(doc_type, excel_path, pdf_path, sheet_filter=sheet_filter)
        if not combined.empty:
            all_combined.append(combined)
        print(flush=True)

    final = pd.concat(all_combined, ignore_index=True) if all_combined else pd.DataFrame()
    if final.empty:
        print("경고: 매칭된 시트가 하나도 없어 저장할 결과가 없습니다 (result 파일을 만들지 않음).", flush=True)
        return final

    print("결과 저장 중...", flush=True)
    paths = step5_save_result(final, result_dir)
    print(f"저장 완료: {paths}", flush=True)

    review_ratio = final["needs_review"].fillna(False).mean()
    print(f"\n요약: 총 {final.shape[0]}행, needs_review 비율 {review_ratio:.1%}", flush=True)
    print(final.groupby("문서종류")["판정"].value_counts().to_string(), flush=True)

    return final


def _parse_args():
    parser = argparse.ArgumentParser(description="보험료 대조 파이프라인 (통합, 문서 종류별 --doc-type)")
    parser.add_argument(
        "--doc-type",
        required=True,
        choices=[*DOCUMENT_TYPES.keys(), "all"],
        help="처리할 문서 종류. 'all'이면 전부 순서대로 처리 후 결과를 합침",
    )
    parser.add_argument("--excel", default=None)
    parser.add_argument(
        "--pdf",
        default=None,
        help="단일 --doc-type일 때만 사용. 생략하면 raw 폴더에서 해당 문서 종류 PDF를 자동 탐색",
    )
    parser.add_argument("--raw-dir", default=str(RAW_DIR))
    parser.add_argument("--result-dir", default=str(RESULT_DIR))
    parser.add_argument(
        "--sheet-filter",
        default=None,
        help="디버깅용: 시트명에 이 문자열이 포함된 시트만 처리(예: 대아)",
    )
    return parser.parse_args()


def _find_pdf_for_doc_type(raw_dir, doc_type):
    """raw 폴더에서 해당 문서 종류로 보이는 PDF를 이름 키워드로 자동 탐색한다."""
    keyword = {"health": "건강", "pension": "국민연금"}[doc_type]
    candidates = [p for p in Path(raw_dir).glob("*.pdf") if keyword in p.name]
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) == 0:
        raise FileNotFoundError(f"raw 폴더에서 '{keyword}'가 포함된 PDF를 찾을 수 없습니다.")
    raise ValueError(f"raw 폴더에 '{keyword}' PDF가 {len(candidates)}개 있어 자동 선택 불가: {[c.name for c in candidates]}")


if __name__ == "__main__":
    args = _parse_args()
    raw_dir = Path(args.raw_dir)
    result_dir = Path(args.result_dir)

    excel_path = Path(args.excel) if args.excel else _find_single_file(raw_dir, "*.xlsx", "엑셀")

    doc_types = list(DOCUMENT_TYPES.keys()) if args.doc_type == "all" else [args.doc_type]

    if args.pdf and len(doc_types) > 1:
        raise SystemExit("--pdf는 --doc-type이 단일 문서 종류일 때만 지정할 수 있습니다.")

    pdf_paths = {}
    for dt in doc_types:
        pdf_paths[dt] = Path(args.pdf) if args.pdf else _find_pdf_for_doc_type(raw_dir, dt)

    run(doc_types, excel_path, pdf_paths, result_dir, sheet_filter=args.sheet_filter)
    print("\n완료! result 폴더 확인하세요.")
