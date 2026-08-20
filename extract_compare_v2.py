"""extract_compare.py / extract_compare_v4.py가 공유하는 소규모 헬퍼 모음.

원본 v2 파일이 git에 커밋되지 않은 채로 사라져(2026-08-14 브랜치 작업 당시
로컬에만 있던 파일로 추정) extract_compare.py/v4.py가 import에 실패하는
상태였다. 여기서는 두 파일의 실제 호출 방식(인자/반환값)을 근거로 동작을
재구성했다.
"""
import re
from pathlib import Path

import pandas as pd

AMOUNT_CLEAN_RE = re.compile(r"[^\d.-]")
GROUPED_AMOUNT_RE = re.compile(r"\d{1,3}(?:,\d{3})+")


def clean_amount(value):
    """엑셀 셀 값을 정수 금액으로 정리한다. 빈 값/파싱 불가 값은 None."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return int(round(value))
    s = AMOUNT_CLEAN_RE.sub("", str(value))
    if not s:
        return None
    try:
        return int(round(float(s)))
    except ValueError:
        return None


def _find_header_idx(rows, max_scan=10):
    """'순번'(0번째 칸)과 '성명'(1번째 칸)이 함께 있는 헤더 행의 인덱스를 찾는다."""
    for idx, row in enumerate(rows[:max_scan]):
        if len(row) < 2:
            continue
        col0 = re.sub(r"\s+", "", str(row[0] or ""))
        col1 = re.sub(r"\s+", "", str(row[1] or ""))
        if "순번" in col0 and "성명" in col1:
            return idx
    return None


def _find_single_file(dir_path, pattern, label):
    """dir_path에서 pattern(glob)에 맞는 파일을 정확히 1개 찾는다."""
    candidates = list(Path(dir_path).glob(pattern))
    if len(candidates) == 1:
        return candidates[0]
    if not candidates:
        raise FileNotFoundError(f"{dir_path}에서 {label} 파일을 찾을 수 없습니다 (패턴: {pattern}).")
    raise ValueError(
        f"{dir_path}에 {label} 파일이 {len(candidates)}개 있어 자동 선택할 수 없습니다: "
        f"{[c.name for c in candidates]}"
    )


def _first_grouped_amount(text):
    """쉼표 3자리 그룹으로 묶인 금액(예: '1,234,560원')만 매치한다 -- 순번처럼
    작은 숫자나 페이지 번호 같은 잡음을 금액으로 오인하지 않기 위함."""
    m = GROUPED_AMOUNT_RE.search(text)
    return clean_amount(m.group()) if m else None


def _first_amount(text):
    """쉼표 그룹 여부와 무관하게 첫 숫자 토큰을 금액으로 본다 (순번처럼 작은
    수를 읽을 때 씀)."""
    m = re.search(r"\d[\d,]*", text)
    return clean_amount(m.group()) if m else None


HEALTH_COMPARE_PAIRS = [
    ("건강보험료_엑셀", "건강보험료_납부", "건강보험료"),
    ("장기요양보험료_엑셀", "장기요양보험료_납부", "장기요양보험료"),
]


def step4_compare(excel_df, pdf_df):
    """건강·장기요양보험료 전용 대조(v4). extract_compare.py의 `_compare_generic`과
    같은 알고리즘이되 건강보험료 compare_pairs로 고정되어 있다. pdf_df의 '성명'은
    항상 None(네이티브 PDF에서는 성명을 안 읽음)이라 병합 전에 버린다."""
    pdf_df = pdf_df.drop(columns=["성명"], errors="ignore")
    merged = pd.merge(excel_df, pdf_df, left_on="순번_엑셀", right_on="순번_pdf", how="outer")

    columns = ["성명", "순번_엑셀", "순번_pdf"]
    for excel_col, pdf_col, label in HEALTH_COMPARE_PAIRS:
        diff_col = f"{label}_차이"
        merged[diff_col] = merged[excel_col] - merged[pdf_col]
        columns += [excel_col, pdf_col, diff_col]

    def _judge(row):
        has_excel = pd.notna(row[HEALTH_COMPARE_PAIRS[0][0]])
        has_pdf = pd.notna(row[HEALTH_COMPARE_PAIRS[0][1]])
        if not has_excel:
            return "엑셀_누락"
        if not has_pdf:
            return "PDF_누락"
        if row.get("needs_review"):
            return "확인필요"
        all_ok = all(row[e] == row[p] for e, p, _ in HEALTH_COMPARE_PAIRS)
        return "일치" if all_ok else "불일치"

    merged["판정"] = merged.apply(_judge, axis=1)
    columns += ["판정", "needs_review", "출처파일"]
    for col in columns:
        if col not in merged.columns:
            merged[col] = None
    return merged[columns].sort_values(["순번_엑셀", "순번_pdf"]).reset_index(drop=True)


def step5_save_result(df, result_dir):
    """대조 결과 DataFrame을 result_dir에 csv/xlsx로 저장하고 저장된 경로를 돌려준다."""
    result_dir = Path(result_dir)
    result_dir.mkdir(parents=True, exist_ok=True)
    csv_path = result_dir / "비교결과.csv"
    xlsx_path = result_dir / "비교결과.xlsx"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    df.to_excel(xlsx_path, index=False)
    return {"csv": csv_path, "xlsx": xlsx_path}
