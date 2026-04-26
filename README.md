# smart-doc-extraction

PDF, PPTX, XLSX, DOCX 같은 비정형 문서에서 **구조 보존된 텍스트 + 표 + 이미지 + VLM 캡션** 을
한 번에 뽑아내는 추출 파이프라인.

오픈소스 라이브러리 (PyMuPDF, pdfplumber, python-pptx, openpyxl, LibreOffice) +
[IBM watsonx Text Extraction V2](https://www.ibm.com/products/watsonx-ai) +
watsonx VLM 을 조합해서 슬라이드/페이지/시트 단위로 정렬된 마크다운을 만들어줌.

---

## 지원하는 파일 타입

| 입력 | 결과 폴더 | 핵심 산출물 |
|---|---|---|
| **PDF** (`.pdf`) | `<output>/<stem>_pdf_analysis/` | `comprehensive_labeling_with_vlm.md`, `comprehensive_analysis_complete.json`, `visual_captures/*.png` |
| **PPTX** (`.pptx`) | `<output>/<stem>_analysis/` | `comprehensive_labeling_with_vlm.md`, `comprehensive_analysis_complete.json`, `visual_captures/*.png` |
| **XLSX / XLSM / XLS** | `<output>/<stem>/` | `assembly.md` (시트 합본), `sheets/<sheet>/`, `excel_extracted.jsonl` |
| **DOCX** (`.docx`) | `<output>/<stem>/` | watsonx 의 `assembly.md` + 임베디드 이미지 |

지원 안 하는 포맷: `.csv`, `.doc`, `.ppt` (legacy 바이너리). DOCX/PPTX 로 미리 변환 후 입력.

---

## 빠른 시작

### 1. 설치

```bash
git clone https://github.com/valentinazie/smart-doc-extraction.git
cd smart-doc-extraction
python -m venv venv
source venv/bin/activate            # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 2. 시스템 의존성 (LibreOffice)

XLSX → PDF 변환을 위해 LibreOffice 가 필요해 (PPTX 도 일부 경로에서 사용):

```bash
# macOS
brew install --cask libreoffice

# Ubuntu / Debian
sudo apt-get install -y libreoffice

# 설치 확인
soffice --version
```

### 3. `.env` 작성

저장소 루트에 `.env` 파일을 만들고 본인 watsonx 자격증명 채우기:

```bash
cp .env.example .env
# 그 다음 .env 열어서 실제 값으로 채움
```

`.env` 에 들어갈 키들 (자세한 건 `.env.example` 참고):

```bash
WATSONX_URL=https://us-south.ml.cloud.ibm.com
WATSONX_APIKEY=...                       # IBM Cloud API key
WATSONX_PROJECT_ID=...                   # PDF VLM 캡션용 (project-scoped)
SPACE_ID=...                             # watsonx Text Extraction 용 (space-scoped)
COS_BUCKET_NAME=...                      # space 와 연결된 COS 버킷
# (선택) 이미지를 별도 COS 로 업로드하고 싶을 때만:
MASTER_COS_ENDPOINT=...
MASTER_COS_ACCESS_KEY=...
MASTER_COS_SECRET_KEY=...
```

> **PDF 만 처리할 거면** `SPACE_ID`/`COS_BUCKET_NAME` 없어도 동작 (PDF 는 watsonx Text Extraction 사용 안 함). 단 VLM 캡션 받으려면 `WATSONX_APIKEY`, `WATSONX_PROJECT_ID` 는 필수.

---

## 사용법

### 한 파일 처리

```bash
# 자동 라우팅 — 확장자 보고 알맞은 파이프라인으로 디스패치
python extract.py path/to/report.pdf --output ./output

python extract.py path/to/slides.pptx --output ./output

python extract.py path/to/data.xlsx --output ./output

python extract.py path/to/memo.docx --output ./output
```

### 폴더 통째로

```bash
# 폴더 안의 모든 지원 파일을 재귀적으로 처리
python extract.py --folder ./my_docs --output ./output

# 한 파일이 실패해도 나머지는 계속 진행
python extract.py --folder ./my_docs --output ./output --continue-on-error
```

### 옵션

```text
python extract.py [path | --folder DIR] [--output DIR] [--dpi 200] [--continue-on-error]

  path                  단일 파일 경로
  --folder, -f DIR      폴더 안 모든 지원 문서를 재귀 처리
  --output, -o DIR      산출물 저장 루트 (기본: ./output)
  --dpi N               PDF 페이지 렌더링 DPI (PDF 만 적용, 기본 200)
  --continue-on-error   폴더 모드에서 한 파일 실패해도 다음으로 진행
```

### 개별 파이프라인 직접 호출 (옵션이 더 많이 필요할 때)

```bash
# PDF
python pdf_extraction/comprehensive_pdf_analyzer.py path/to/file.pdf --output ./output

# PPTX
python ppt_extraction/comprehensive_presentation_analyzer.py path/to/file.pptx --output ./output

# XLSX (per-sheet vs whole-workbook 옵션, 이미지 모드 옵션 등)
python excel_extraction/excel_to_jsonl_pipeline.py path/to/file.xlsx --output-dir ./output

# DOCX (folder 도 받음, --reprocess 로 재추출 강제)
python watsonx_text_extraction/text_extraction.py path/to/file.docx --output ./output
```

---

## 산출물 살펴보기

### PDF / PPTX 결과 폴더

```
<output>/<stem>_pdf_analysis/        (PDF) 또는 <stem>_analysis/ (PPTX)
├── comprehensive_labeling_with_vlm.md   ← 사람이 읽기 좋은 결과 (페이지/슬라이드별
│                                          텍스트 + 표 + 이미지 + VLM 캡션)
├── comprehensive_analysis_complete.json ← 모든 위치/타입 정보 포함된 풀 덤프
├── comprehensive_summary.json           ← 통계 요약
└── visual_captures/                     ← 잘라낸 이미지 PNG 들
    ├── slide_01_picture_03_visual.png
    └── ...
```

`comprehensive_labeling_with_vlm.md` 가 보통 후속 RAG / LLM 입력으로 쓰는 메인 파일.

### XLSX 결과 폴더

```
<output>/<stem>/
├── assembly.md                          ← 모든 시트의 마크다운을 하나로 합친 파일
├── excel_extracted.jsonl                ← 시트별 메타+텍스트가 한 줄씩 들어있는 JSONL
└── sheets/
    ├── 01_품질지표/
    │   ├── pdf/                         ← 시트당 PDF 변환 결과
    │   └── watsonx/                     ← watsonx 가 뽑은 마크다운 + 임베디드 이미지
    └── ...
```

### DOCX 결과 폴더

```
<output>/<stem>/
├── assembly.md                          ← watsonx 가 만든 마크다운
└── (임베디드 이미지가 있다면 이미지 파일들)
```

---

## 빠른 트러블슈팅

| 증상 | 원인 / 대처 |
|---|---|
| `LibreOffice not found` | 위 "시스템 의존성" 섹션대로 `soffice` 설치 |
| Excel/DOCX 가 `InvalidAccessKeyId` 로 죽음 | watsonx Space 의 COS HMAC 자격증명이 만료/회전됨. IBM Cloud 콘솔에서 Space 의 storage credential 갱신 |
| PPTX 결과에 `**VLM Status:** Not Available` 라고 뜸 | watsonx 호출이 실패해서 VLM 단계로 못 감. `.env` 의 `WATSONX_APIKEY` / `SPACE_ID` 확인 |
| `KeyError: 'WATSONX_APIKEY'` | `.env` 가 저장소 루트에 있는지, 키 이름 오타 없는지 확인 |
| 폴더 모드로 돌렸는데 한 파일에서 멈춤 | `--continue-on-error` 추가 |
| 같은 파일 다시 처리하고 싶음 (DOCX) | `python watsonx_text_extraction/text_extraction.py FILE --output DIR --reprocess` |

---

## 디렉토리 한눈에 보기

```
.
├── extract.py                           # 통합 진입점 (확장자 자동 라우팅)
├── pdf_extraction/                      # PDF 파이프라인
├── ppt_extraction/                      # PPTX 파이프라인
├── excel_extraction/                    # XLSX 파이프라인
├── watsonx_text_extraction/             # watsonx Text Extraction V2 래퍼 (DOCX 메인)
├── common/                              # .env 로딩, watsonx/COS 클라이언트 헬퍼
├── utils/                               # geometry 등 공유 유틸
├── requirements.txt
└── .env.example
```

---

## 라이선스 / 책임

본 저장소는 데모/PoC 목적입니다. watsonx 호출에는 IBM Cloud 의 사용량이 발생합니다.
민감 데이터를 처리할 때는 `.env`, 처리 결과 폴더(`output/`)가 절대 git 에 커밋되지
않도록 `.gitignore` 를 확인하세요.
