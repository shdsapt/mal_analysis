# 🛡️ Automated File Analysis & AI Report Tool

이 프로젝트는 의심스러운 파일(PDF, Excel, Image, EXE, PPT)을 다각도로 분석하고, 그 결과를 AI(Gemini)와 연동하여 상세한 보안 분석 보고서를 자동 생성하는 도구입니다.

> **Windows/Linux 모두 지원** — Python 내장 라이브러리 기반으로 OS 독립적으로 동작합니다.

## 📂 주요 기능 및 업데이트

1.  **`file_analysis.py`**: 정적 분석 도구
    *   **다양한 포맷 지원**: PDF, Excel, Image(PNG/JPG), EXE(PE), PPT 파일 분석.
    *   **심층 분석 옵션 (`--deep`)**: Excel 파일(xlsx)의 압축을 풀고 내부 구조(매크로, 외부 링크, 스크립트 등)를 정밀 검사.
    *   **자동 감지**: 파일 확장자에 따른 자동 분석 모드 지원.
    *   **OS 독립적**: `strings`, `grep`, `sha256sum` 등 Linux 셸 명령어를 Python 내장 라이브러리로 대체.

2.  **`ai_analysis.py`**: AI 자동화 래퍼 (Batch Processing)
    *   **비동기 일괄 처리**: Python `asyncio`를 사용하여 폴더 내의 **모든 파일**을 동시에 빠르게 분석.
    *   **AI 리포트 생성**: 분석된 기술적 데이터를 바탕으로 사람이 읽기 쉬운 한글 요약 보고서 생성.
    *   **자동 파일 정리**: 분석이 완료된 원본 파일은 `analyzed` 폴더로 자동 이동.

---

## 🛠️ 사전 요구사항 (Prerequisites)

### Python 패키지 설치

```bash
pip install -r requirements.txt
```

**필수 패키지:**
*   `pefile` — PE(EXE/DLL) 파일 헤더/섹션 분석
*   `oletools` — Office VBA 매크로 분석 (`olevba`)
*   `python-magic-bin` — 파일 타입 감지 (Windows 호환)
*   `pdfid` — PDF 구조 분석
*   `pikepdf` — PDF 객체/스트림 심층 분석

### 선택적 외부 도구
*   `exiftool` — [다운로드](https://exiftool.org/) 후 PATH에 추가 (이미지 메타데이터 분석)
*   `vt` (VirusTotal CLI) — [다운로드](https://github.com/VirusTotal/vt-cli) 후 PATH에 추가
*   `gemini-cli` — AI 분석용 (ai_analysis.py 사용 시 필요)

---

## 🚀 사용 방법 (Usage)

### 1. 개별 파일 상세 분석 (`file_analysis.py`)

단일 파일에 대해 기술적인 분석을 수행하고 결과를 출력합니다.

```bash
# 기본 사용법 (자동 파일 형식 감지)
python file_analysis.py -file <파일명>

# 엑셀 파일 심층 분석 (Unzip & Recursive search)
python file_analysis.py -file <파일명.xlsx> --deep

# 특정 형식 지정
python file_analysis.py -pdf <파일명.pdf>
python file_analysis.py -exe <파일명.exe>
```

**주요 옵션:**
*   `-file`: 파일 형식을 자동 감지하여 분석합니다.
*   `-pdf`, `-xls`, `-ppt`, `-img`, `-exe`: 특정 형식을 명시적으로 지정합니다.
*   `--deep`, `-d`: (엑셀 전용) 압축 해제 후 내부 XML 및 파일들에 대해 정밀 검색을 수행합니다.
*   `-out`: 로그 저장 경로 지정. (기본값: `./instruction_output/`)

---

### 2. AI 일괄/단일 분석 (`ai_analysis.py`)

여러 파일을 한 번에 처리하거나, AI 분석 리포트를 받아보고 싶을 때 사용합니다.

#### A. 폴더 내 모든 파일 일괄 분석 (Batch Mode)
별도의 인자 없이 실행하면, **현재 디렉토리**에 있는 분석 대상 파일들을 자동으로 찾아 **병렬(비동기) 분석**을 시작합니다.

```bash
python ai_analysis.py
```

*   **동작 흐름**:
    1.  현재 폴더 스캔 (스크립트 제외)
    2.  `file_analysis.py` 실행 (엑셀일 경우 `--deep` 모드 자동 적용 등)
    3.  `gemini-cli`로 리포트 생성
    4.  결과물은 `analysis_result` 폴더(또는 지정 경로)에 저장
    5.  **완료된 원본 파일은 `./analyzed` 폴더로 이동**

#### B. 단일 파일 분석 (Single Mode)
특정 파일 하나만 분석하고 싶을 때 사용합니다.

```bash
python ai_analysis.py -file malicious_doc.xlsx
```

---

## 📝 산출물 예시

*   **분석 리포트 (`analysis_result/`)**: `[날짜]_[파일명]_analyis_result.md`
    *   포함 내용: 파일 정보, 감지된 위협(해시, 매크로, 문자열), AI 종합 의견, 원본 로그.
*   **이동된 원본 (`analyzed/`)**: 분석이 끝난 파일들이 안전하게 격리/이동됨.

## ⚠️ 주의사항

*   일괄 분석(`ai_analysis.py`) 실행 시, 현재 폴더의 파일 위치가 변경(이동)되므로 테스트 시 유의하세요.
*   `--deep` 옵션은 압축을 해제하므로 디스크 공간을 일시적으로 사용하며, 완료 후 자동 삭제됩니다.
*   `exiftool`과 `vt` CLI는 별도 설치 후 PATH에 추가해야 합니다. 미설치 시 해당 분석 단계만 건너뜁니다.
