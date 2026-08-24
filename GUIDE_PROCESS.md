# [VSCode 개발 가이드] 보험료 증빙 vs 단위공사별 납부내역 비교 모니터링 시스템

본 문서는 **VSCode 환경에서 건강보험, 국민연금, 퇴직공제부금 증빙 PDF와 단위공사별 엑셀 납부원장을 대조하고 모니터링하는 검증 시스템**을 순차적으로 개발·테스트·배포할 수 있도록 구성된 표준 실행 프로세스 가이드입니다.

---

## 📌 목차
1. [개발 환경 준비 (Environment Setup)](#1-개발-환경-준비-environment-setup)
2. [VSCode 권장 확장 및 디버깅 설정](#2-vscode-권장-확장-및-디버깅-설정)
3. [전체 시스템 아키텍처 및 3자 대조 흐름](#3-전체-시스템-아키텍처-및-3자-대조-흐름)
4. [단계별 개발 및 실행 프로세스 (Step-by-Step)](#4-단계별-개발-및-실행-프로세스-step-by-step)
   - [Step 1: 입력 데이터 구조 진단](#step-1-입력-데이터-구조-진단)
   - [Step 2: 추출 엔진 범용화 (PDF & 엑셀)](#step-2-추출-엔진-범용화-pdf--엑셀)
   - [Step 3: 3-Way 대조 엔진 및 위험 감지 룰셋 구현](#step-3-3-way-대조-엔진-및-위험-감지-룰셋-구현)
   - [Step 4: Flask 웹 대시보드 및 리포트 내보내기 연동](#step-4-flask-웹-대시보드-및-리포트-내보내기-연동)
5. [테스트 및 검증 명령어 모음](#5-테스트-및-검증-명령어-모음)
6. [트러블슈팅 가이드](#6-트러블슈팅-가이드)

---

## 1. 개발 환경 준비 (Environment Setup)

VSCode 내장 터미널(`Ctrl + ~` 또는 ``Ctrl + ` ``)에서 아래 명령어를 실행하여 파이썬 가상환경을 활성화하고 의존성을 확인합니다.

```powershell
# 1. 가상환경 활성화 (PowerShell 기준)
.\venv\Scripts\Activate.ps1

# 2. 필수 라이브러리 설치/확인
pip install -r requirements.txt

# 필수 패키지 구성:
# - flask, openpyxl, pandas, pdfplumber, python-dotenv, werkzeug
```

---

## 2. VSCode 권장 확장 및 디버깅 설정

### 1) `.vscode/launch.json` 설정 (F5 디버깅용)
VSCode에서 `F5` 키를 눌러 손쉽게 디버그할 수 있도록 `.vscode/launch.json`에 다음 구성을 추가합니다:

```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "name": "Flask App (Web Dashboard)",
      "type": "debugpy",
      "request": "launch",
      "module": "flask",
      "env": {
        "FLASK_APP": "app.py",
        "FLASK_ENV": "development",
        "FLASK_DEBUG": "1"
      },
      "args": ["run", "--port", "5000"]
    },
    {
      "name": "대조 CLI: 대아이앤씨 (Health & Pension)",
      "type": "debugpy",
      "request": "launch",
      "program": "${workspaceFolder}/extract_compare.py",
      "args": ["--doc-type", "all", "--sheet-filter", "대아"]
    },
    {
      "name": "대조 CLI: 퇴직공제 검증",
      "type": "debugpy",
      "request": "launch",
      "program": "${workspaceFolder}/extract_compare.py",
      "args": ["--doc-type", "retirement"]
    }
  ]
}
```

---

## 3. 전체 시스템 아키텍처 및 3자 대조 흐름

```
[1. 공단/공제회 전자발급 PDF] ──┐
                               ├──> [정규화 & 3-Way 대조 엔진] ──> [5대 위험 분류] ──> [웹 대시보드 / 엑셀 리포트]
[2. 현장 실적정산철 (06-03~05)] ──┤    (extract_compare.py)       - MATCH (일치)
                               │                               - UNDERPAID (과소납)
[3. 단위공사별 납부원장 엑셀] ────┘                               - OVERPAID (과오납)
                                                               - GHOST_WORKER (인력 누락)
                                                               - REVIEW_NEEDED (확인 필요)
```

---

## 4. 단계별 개발 및 실행 프로세스 (Step-by-Step)

### Step 1: 입력 데이터 구조 진단
다양한 협력업체(대아, 신보, 신한, NSC 등)의 시트 구조와 PDF 포맷을 자동 분석합니다.

* **실행 명령**:
  ```powershell
  python -c "
  import openpyxl
  wb = openpyxl.load_workbook('sample/보험료 납부_단위공사별(2026.05.26).xlsx', read_only=True)
  print('총 시트 수:', len(wb.sheetnames))
  print('26.04월 관련 시트:', [s for s in wb.sheetnames if '26.04' in s][:10])
  "
  ```
* **점검 포인트**:
  1. PDF가 네이티브 텍스트인지 스캔본(OCR 필요)인지 폰트 분석
  2. 엑셀 헤더가 1행인지 2행 복합 그룹 헤더인지 확인

---

### Step 2: 추출 엔진 범용화 (PDF & 엑셀)

`extract_compare.py`의 핵심 파서를 고도화합니다.

1. **퇴직공제부금(`retirement`) 파서 추가**:
   - `06-05. 제20회 기성 실적정산(퇴직공제).pdf` 및 `퇴직공제부금 납부 신고 내역.xlsx` 대응
   - 한 파일 내 다수 협력업체가 혼재된 경우 업체별 자동 필터링 및 공제부금 일수 추출
2. **엑셀 복합 하위공사 파서 일반화**:
   - 대아처럼 1개 업체 내 복수 하위공사(예: 배관공사, 주기기설치공사 등)로 '사용자부담금' 열이 나뉜 경우 자동 탐색 및 병합
3. **노이즈 폰트 자동 감지 모듈 공용화**:
   - `_detect_needs_noise_filter()`를 통해 네이티브 PDF의 중복 레이어 제거

---

### Step 3: 3-Way 대조 엔진 및 위험 감지 룰셋 구현

추출된 데이터 간의 무결성을 검증하고 5대 이상 징후를 판별합니다.

* **매칭 키(Matching Key)**: `(성명, 생년월일 앞자리, 협력업체명, 귀속월)`
* **비교 필드**:
  - 건강보험: `건강보험료_고지`, `장기요양보험료_고지`, `건강/요양 납부액` vs 엑셀 납부액
  - 국민연금: `기준소득월액`, `연금보험료`, `사용자부담금` vs 엑셀 사용자부담금
  - 퇴직공제: `신고공수(일수)`, `공제부금액` vs 엑셀 공제부금
* **결과 판정 로직**:
  - 금액 100% 일치 ➔ `MATCH` (정상)
  - 증빙 > 엑셀 ➔ `OVERPAID` (초과 청구 / 과오납 위험)
  - 증빙 < 엑셀 ➔ `UNDERPAID` (과소납 / 정산 누락 위험)
  - 한쪽에만 존재 ➔ `GHOST_WORKER` (인력 누락 / 유령 인력)
  - 폰트/OCR 왜곡 ➔ `REVIEW_NEEDED` (육안 확인 필요)

---

### Step 4: Flask 웹 대시보드 및 리포트 내보내기 연동

1. **Flask 서버 기동**:
   ```powershell
   python app.py
   ```
2. **웹 브라우저 접속**: `http://localhost:5000`
3. **주요 기능**:
   - 기성회차별 실적정산철 PDF 및 단위공사별 엑셀 일괄 업로드
   - 비목별(건강, 연금, 퇴직) 검증 현황 요약 카드 (총 인원, 일치율, 이상 건수)
   - 불일치 항목 필터링 및 원본 대조 Side-by-Side 뷰어
   - 검증 결과 엑셀 리포트 다운로드

---

## 5. 테스트 및 검증 명령어 모음

VSCode 터미널에서 아래 테스트를 순차적으로 실행하여 개발 상태를 검증합니다.

```powershell
# 1. 건강보험 파이프라인 검증 (대아이앤씨)
python extract_compare.py --doc-type health --sheet-filter 대아

# 2. 국민연금 파이프라인 검증 (대아이앤씨)
python extract_compare.py --doc-type pension --sheet-filter 대아

# 3. 전체 비목 통합 실행 (건강 + 연금 + 퇴직)
python extract_compare.py --doc-type all --sheet-filter 대아

# 4. 단위 테스트 스크립트 실행
python -m unittest discover -s test -p "test_*.py"
```

---

## 6. 트러블슈팅 가이드

* **Q1. PDF 텍스트가 추출되지 않고 빈 값으로 나올 때**
  * *원인*: PDF가 텍스트 레이어가 없는 순수 스캔 이미지인 경우입니다.
  * *해결*: OCR을 사전 적용(`ocr_done/`)하거나, 공단 웹사이트에서 직접 다운로드한 전자발급 원본 PDF(네이티브)를 사용합니다.
* **Q2. 엑셀 헤더 위치를 찾지 못할 때 (`ValueError: 헤더 행을 찾을 수 없습니다`)**
  * *원인*: 신규 협력업체의 헤더 표기가 다르거나 그룹 헤더 구조가 상이한 경우입니다.
  * *해결*: `_find_header_idx()`에 해당 업체의 헤더 키워드(`성명`, `이름`, `순번` 등)를 추가 등록합니다.
* **Q3. 금액 앞자리가 잘려서 인식될 때**
  * *해결*: `extract_compare.py` 내의 `amount_consistency_pairs` (고지액 2배 = 납부액) 검증으로 자동 플래그 처리됩니다.
