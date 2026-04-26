# Unstructured Data Extraction Pipeline (`src/`)

삼성 실패사례 원본 문서 (`.pdf` / `.pptx` / `.docx` / `.xlsx`) 를 파싱해서
텍스트·이미지·표·레이아웃까지 구조화된 Markdown / JSON 으로 뽑아내는
추출 파이프라인입니다.

검색/추론을 담당하는 `agent-system/` 과 달리, 여기서는 "문서 → 구조화된 데이터"
변환만 책임집니다. (에이전트가 읽는 통합 manifest 는 별도로 구축되며
`data/tbl_failure_joined.with_extracted.json` 에 위치합니다 — git 에서는 제외.)

---

## 빠른 시작 — 통합 진입점 `src/extract.py`

파일 타입을 자동으로 감지해서 해당 포맷의 파이프라인으로 라우팅해 줍니다.
`.pdf` / `.pptx` / `.xlsx` (`.xlsm` / `.xls` 포함) / `.docx` 지원, `.csv` 등은 건너뜁니다.

```bash
source venv/bin/activate

# 단일 파일 — 확장자 보고 알아서 분기
python src/extract.py path/to/file.pdf
python src/extract.py path/to/slides.pptx  --output ./output
python src/extract.py path/to/book.xlsx    --output ./output

# 폴더 전체 재귀 처리 (지원 안 되는 파일은 자동 skip)
python src/extract.py --folder client_data/data --output ./output --continue-on-error
```

Python 모듈에서 직접 호출하려면:

```python
from src.extract import extract
extract("docs/report.pdf", output="./output")     # PDF → pdf_extraction
extract("docs/deck.pptx",  output="./output")     # PPTX → ppt_extraction
extract("docs/table.xlsx", output="./output")     # XLSX → excel_extraction
extract("docs/memo.docx",  output="./output")     # DOCX → watsonx_text_extraction
```

확장자 → 엔진 라우팅 규칙:

| 확장자 | 엔진 |
|--------|------|
| `.pdf` | `pdf_extraction/comprehensive_pdf_analyzer.py` |
| `.pptx` | `ppt_extraction/comprehensive_presentation_analyzer.py` |
| `.xlsx` / `.xlsm` / `.xls` | `excel_extraction/excel_to_jsonl_pipeline.py` |
| `.docx` | `watsonx_text_extraction/text_extraction.py` |
| 그 외 (`.csv`, `.doc`, `.ppt`, …) | ❌ skip |

내부적으로는 각 파이프라인의 `main()` 을 `sys.argv` 만 갈아 끼워서 in-process
로 호출합니다. 따라서 각 파이프라인의 기존 CLI / 옵션은 그대로 살아 있고, 더
세밀한 제어가 필요하면 각 `*.py` 를 직접 실행해도 됩니다.

---

## 전체 분석 아키텍처

```
                        📄 문서 입력
                             │
                             ▼
                ┌──── 확장자 기반 라우팅 ────┐
                │                             │
        ┌───────┼───────┬──────────┬─────────┐
        ▼       ▼       ▼          ▼
      .pptx   .pdf    .xlsx      .docx
        │       │       │          │
━━━━━━━━┿━━━━━━━┿━━━━━━━┿━━━━━━━━━━┿━━━━━━━━━━━━━━━━━━━━━━━━
        │       │       │          │
        ▼       ▼       ▼          │
┌──────────────────────────────┐   │
│ 🟢 오픈소스                   │   │
│ (python-pptx / PyMuPDF /      │   │
│  pdfplumber / openpyxl /      │   │
│  LibreOffice / Pillow)        │   │
├──────────────────────────────┤   │
│  · 문서 구조 파악              │   │
│    (페이지·시트·슬라이드 분할) │   │
│  · 좌표 / bbox 추출           │   │
│  · 이미지 영역 crop            │   │
│  · (PPTX / XLSX) → PDF 변환   │   │
└────────────┬─────────────────┘   │
             │                      │
             ▼                      ▼
     ┌───────────────────────────────────────┐
     │ 🔵 IBM watsonx Text Extraction V2      │
     ├───────────────────────────────────────┤
     │  · 텍스트 본문 OCR                     │
     │    (high_quality, 한국어+영어)         │
     │  · 표 구조 / 셀 추출                   │
     │                                         │
     │  ※ PDF 는 V2 사용 안 함                │
     │     (OSS 만으로 텍스트+표 추출)        │
     └────────────────┬──────────────────────┘
                      │
                      ▼
     ┌───────────────────────────────────────┐
     │ 🔵 watsonx.ai VLM                      │
     │    (Mistral-Small 3.1 24B)             │
     ├───────────────────────────────────────┤
     │  · 이미지 캡션 생성 (전 포맷 공통)     │
     └────────────────┬──────────────────────┘
                      │
                      ▼
     ┌───────────────────────────────────────┐
     │ 🟢 Labeling (순수 Python)              │
     │    텍스트 + 표 + 이미지 캡션을         │
     │    페이지/시트 단위 Markdown 으로 병합 │
     │    → comprehensive_labeling_with_vlm.md│
     │      또는 assembly.md                  │
     └────────────────┬──────────────────────┘
                      │
                      ▼
           output/<stem>_analysis/   ← 분석 산출물 (per-document)
```

### 한 줄 요약

| 포맷  | 오픈소스                             | IBM Text Extraction V2       | watsonx VLM |
|-------|--------------------------------------|------------------------------|-------------|
| PPTX  | 구조 파악 · 이미지 crop · PDF 변환   | 텍스트 · 표                  | 이미지 캡션 |
| PDF   | 구조 파악 · 텍스트 · 표 · 이미지 crop | ❌                           | 이미지 캡션 |
| XLSX  | 시트 분할 · 이미지 crop · PDF 변환   | 텍스트 · 표                  | 이미지 캡션 |
| DOCX  | —                                    | 텍스트 · 표 · 문서 전체       | 이미지 캡션 |

- **오픈소스** = 문서를 뜯어서 조각내는 역할 (페이지/시트/슬라이드로 나누고 이미지 영역을 잘라냄)
- **IBM Text Extraction V2** = 텍스트·표 추출 전담 (PDF 제외)
- **watsonx VLM** = 잘라낸 이미지에 한국어 캡션 달아주는 역할

---

## 포맷별 파이프라인 상세

각 pipeline 이 내부적으로 무엇을 실행하는지입니다. 공통 원칙: **포맷이 가진
구조(슬라이드/페이지/시트)를 최대한 보존**하면서 텍스트·이미지·표를 한꺼번에 뽑는
것. 그래야 나중에 에이전트가 "2페이지 두 번째 표" 같은 질문에 답할 수 있음.

### 🟦 PDF — `src/pdf_extraction/comprehensive_pdf_analyzer.py`

```
┌─────────────────────────────────── PDF 파이프라인 ──────────────────────────────────┐
│ 입력: *.pdf                                                                          │
│                                                                                      │
│   ┌─ PDFConversionMixin ───────────────────────────────────────────────┐           │
│   │  PyMuPDF 로 각 페이지를 고해상도 PNG 렌더 (render_dpi=200)          │           │
│   │  → pdf_pages/page_N.png (VLM 입력 + 시각 검증용)                    │           │
│   └────────────────────────────────────────────────────────────────────┘           │
│                                   │                                                  │
│   ┌─ PDFTextExtractionMixin ──────▼────────────────────────────────────┐           │
│   │  PyMuPDF(fitz) 로 페이지별 텍스트 블록 + 좌표(bbox) 추출            │           │
│   │  → text_structure/*.json   (글자 단위 x,y,w,h, 폰트, 줄단위)        │           │
│   │  ※ watsonx Text Extraction V2 는 PDF 경로에서 사용하지 않음         │           │
│   └────────────────────────────────────────────────────────────────────┘           │
│                                   │                                                  │
│   ┌─ PDFSpatialMixin ─────────────▼────────────────────────────────────┐           │
│   │  텍스트 블록을 페이지 좌표계에서 공간 분석 — 컬럼/행/헤더 추정      │           │
│   │  → spatial_analysis/*.json                                          │           │
│   └────────────────────────────────────────────────────────────────────┘           │
│                                   │                                                  │
│   ┌─ PDFTablesMixin ──────────────▼────────────────────────────────────┐           │
│   │  pdfplumber.extract_tables + 격자선 hinting → 2D 배열               │           │
│   │  → table_extractions/*.json (셀 단위 텍스트 + rowspan/colspan)      │           │
│   └────────────────────────────────────────────────────────────────────┘           │
│                                   │                                                  │
│   ┌─ PDFVisualCaptureMixin ───────▼────────────────────────────────────┐           │
│   │  본문 내 이미지(그림/스캔/도표) 영역을 PNG 로 크롭                  │           │
│   │  → visual_captures/shape_*.png                                      │           │
│   └────────────────────────────────────────────────────────────────────┘           │
│                                   │                                                  │
│   ┌─ SmartGroupingMixin + ReadingOrderMixin  (↙ ppt_extraction 에서 공유) ┐        │
│   │  공간 근접성 + 라벨 매칭으로 텍스트 + 이미지 + 표를 "의미 단위 그룹"으로│        │
│   │  묶고, 각 페이지의 독서 순서(reading order) 를 재구성              │        │
│   │  → smart_groups/ · reading_order_groups/                            │        │
│   └────────────────────────────────────────────────────────────────────┘           │
│                                   │                                                  │
│   ┌─ VLMMixin ────────────────────▼────────────────────────────────────┐           │
│   │  각 visual_captures/*.png → watsonx.ai VLM (Mistral-Small 3.1 24B)  │           │
│   │  한국어 caption 생성 ("세탁기 flange shaft 부식 장면")              │           │
│   │  → watsonx_raw_outputs/*.json                                       │           │
│   └────────────────────────────────────────────────────────────────────┘           │
│                                   │                                                  │
│   ┌─ LabelingMixin ───────────────▼────────────────────────────────────┐           │
│   │  위의 텍스트·표·이미지 caption 을 Markdown 한 파일로 합침 (페이지별) │           │
│   │  → comprehensive_labeling_with_vlm.md                               │           │
│   └────────────────────────────────────────────────────────────────────┘           │
│                                                                                      │
│   산출물 (analyzed_output/pdf/<stem>_pdf_analysis/):                                 │
│     comprehensive_analysis_complete.json    (구조화된 모든 분석)                     │
│     comprehensive_summary.json              (요약)                                   │
│     comprehensive_labeling_with_vlm.md      (사람/LLM 이 읽을 최종 md)              │
│     visual_captures/*.png                   (그림 크롭)                              │
│     pdf_pages/*.png                         (페이지 원본 렌더)                       │
└──────────────────────────────────────────────────────────────────────────────────────┘
```

**핵심 포인트**
- ppt_extraction 의 `SmartGrouping`/`ReadingOrder`/`VLM`/`Labeling` Mixin 을 그대로 상속해 로직 재사용.
- 원래 MS 포맷을 거치지 않기 때문에 PPTX 전환 과정에서 깨지던 레이아웃 손실 없음.

### 🟧 PPTX — `src/ppt_extraction/comprehensive_presentation_analyzer.py`

```
┌─────────────────────────────────── PPTX 파이프라인 ─────────────────────────────────┐
│ 입력: *.pptx                                                                         │
│                                                                                      │
│   ① python-pptx 로 각 슬라이드 shape 트리 파싱 (OSS)                                 │
│       → SpatialMixin     : shape 좌표 / bbox                                         │
│       → TablesMixin      : 표 cell 그리드 (python-pptx native)                       │
│                                                                                      │
│   ② LibreOffice → PDF → pdf2image 로 슬라이드별 PNG 렌더링 (OSS)                     │
│       → pdf_pages/slide_NN.png  (레이아웃 기준 이미지)                              │
│       ※ <stem>.pdf 가 이미 있으면 LibreOffice 건너뜀                                │
│                                                                                      │
│   ③ VisualCaptureMixin : shape 단위로 이미지 영역 크롭 (OSS, Pillow)                 │
│       → visual_captures/slide_NN_shape_*.png                                         │
│                                                                                      │
│   ④ TextExtractionMixin — watsonx Text Extraction V2                                 │
│       ②의 PDF 를 COS 업로드 → high_quality + OCR + ko/en 로 텍스트 본문 OCR         │
│       ①의 python-pptx 좌표 정보와 결합 (combine_watsonx_text_with_spatial)           │
│       ※ V2 가 실패/미사용이면 python-pptx native 로 fallback                         │
│                                                                                      │
│   ⑤ SmartGroupingMixin + ReadingOrderMixin (OSS)                                     │
│       텍스트 + 표 + 이미지 shape 를 의미 그룹으로 묶고 독서 순서 재구성              │
│                                                                                      │
│   ⑥ VLMMixin — watsonx.ai Mistral-Small 3.1 24B 로 한국어 이미지 caption             │
│                                                                                      │
│   ⑦ LabelingMixin — 슬라이드별 Markdown 합성 (OSS)                                   │
│                                                                                      │
│   산출물 (analyzed_output/ppt/<stem>_analysis/):                                     │
│     comprehensive_analysis_complete.json                                             │
│     comprehensive_summary.json                                                       │
│     comprehensive_labeling_with_vlm.md      ← 슬라이드 단위 split                    │
│     visual_captures/*.png                                                            │
└──────────────────────────────────────────────────────────────────────────────────────┘
```

**핵심 포인트**
- 10 개 Mixin (`conversion`, `text_extraction`, `spatial`, `tables`, `visual_capture`,
  `smart_grouping`, `reading_order`, `vlm`, `labeling`, `visualization`) 으로 기능 분해.
- VLM 호출엔 `WATSONX_PROJECT_ID` 필수. 없으면 caption 영역에 `"VLM Caption Failed"` 로 표시되고 나머지 단계는 정상 진행.

### 🟩 XLSX — `src/excel_extraction/excel_to_jsonl_pipeline.py`

```
┌─────────────────────────────────── XLSX 파이프라인 ─────────────────────────────────┐
│ 입력: *.xlsx | *.xlsm | *.xls                                                        │
│                                                                                      │
│   ┌─ 시트 분할 ────────────────────────────────────────────────────────┐            │
│   │  통합 문서를 시트별로 쪼개서 각 시트를 독립 xlsx 로 저장             │            │
│   │  → sheets/NN_<sheet_name>.xlsx                                      │            │
│   └────────────────────────────────────────────────────────────────────┘            │
│                                   │                                                  │
│   ┌─ 네이티브 이미지 추출 ────────▼────────────────────────────────────┐            │
│   │  xlsx zip 안의 xl/media/* 를 직접 언패킹 (LibreOffice 렌더 중       │            │
│   │  떨어지는 floating shape 손실 방지)                                 │            │
│   │  → sheets/NN_*/images/*.png                                         │            │
│   └────────────────────────────────────────────────────────────────────┘            │
│                                   │                                                  │
│   ┌─ LibreOffice 변환 ────────────▼────────────────────────────────────┐            │
│   │  각 시트 xlsx → PDF (headless)                                      │            │
│   │  → sheets/NN_*/*.pdf                                                │            │
│   └────────────────────────────────────────────────────────────────────┘            │
│                                   │                                                  │
│   ┌─ watsonx Text Extraction V2 ──▼────────────────────────────────────┐            │
│   │  PDF → COS 업로드 → TextExtractionsV2 호출                          │            │
│   │  (CREATE_EMBEDDED_IMAGES = enabled_verbatim  ← 이미지 인라인 포함)  │            │
│   │  결과물 다운로드 → assembly.md + image files                        │            │
│   │  → sheets/NN_*/watsonx/                                             │            │
│   └────────────────────────────────────────────────────────────────────┘            │
│                                   │                                                  │
│   ┌─ 시트 병합 ───────────────────▼────────────────────────────────────┐            │
│   │  시트별 assembly.md 를 `## Sheet: <name>` 헤더와 함께 합침          │            │
│   │  → <stem>/assembly.md                                               │            │
│   └────────────────────────────────────────────────────────────────────┘            │
│                                   │                                                  │
│   ┌─ JSONL append ────────────────▼────────────────────────────────────┐            │
│   │  이 파일 처리 결과 한 줄을 excel_extracted.jsonl 에 추가            │            │
│   │  → excel_extraction/output/excel_extracted.jsonl                    │            │
│   └────────────────────────────────────────────────────────────────────┘            │
└──────────────────────────────────────────────────────────────────────────────────────┘
```

**핵심 포인트**
- **시트 = 페이지** 가 정확히 1:1 로 보존됨. 30 건의 failure 중 엑셀은 시트별로 `pages[]` 에 분리되어 들어감.
- 이미지 2중 안전장치: (a) xlsx zip 네이티브 추출 + (b) watsonx verbatim mode. 어느 한쪽이 놓치면 다른 쪽이 커버.

### 🟨 DOCX — `src/watsonx_text_extraction/text_extraction.py`

```
┌─────────────────────────────────── DOCX 파이프라인 ─────────────────────────────────┐
│ 입력: *.docx                                                                         │
│                                                                                      │
│   docx 는 내부 구조가 단순하고 슬라이드/페이지 분할이 의미가 없어서                  │
│   별도 파이프라인 없이 **watsonx Text Extraction V2 한 방** 으로 끝냄               │
│                                                                                      │
│   ┌─ text_extraction.py  (single file OR folder, optional --output) ┐            │
│   │  ① 파일을 COS 버킷에 업로드 (DataConnection + S3Location)          │            │
│   │  ② TextExtractionsV2 job 제출                                      │            │
│   │     - AUTO_ROTATION_CORRECTION = True                              │            │
│   │     - CREATE_EMBEDDED_IMAGES  = enabled_placeholder                │            │
│   │     - OUTPUT_DPI              = 150                                │            │
│   │     - OUTPUT_TOKENS_AND_BBOX  = True                               │            │
│   │  ③ 완료 poll → 결과 COS 에 저장                                    │            │
│   │     text_extraction_results/<stem>_<job_id>/assembly.md            │            │
│   │     text_extraction_results/<stem>_<job_id>/embedded_images_       │            │
│   │       assembly/*.png                                               │            │
│   │  ④ idempotent: 같은 stem prefix 에 assembly.md 이미 있으면 skip   │            │
│   └────────────────────────────────────────────────────────────────────┘            │
│                                   │                                                  │
│   ┌─ download_cos_results.py (별도 스크립트) ───────────────────────────┐           │
│   │  COS prefix → 로컬 analyzed_output/docx/<stem>_<ts>/ 로 pull        │            │
│   └────────────────────────────────────────────────────────────────────┘            │
│                                                                                      │
│   batch 모드는 폴더를 재귀 탐색해서 PDF/PPTX/XLSX 도 같은 엔진으로 뽑을 수 있는데,   │
│   실제 운영에선 PDF/PPTX/XLSX 는 위의 전용 파이프라인 을 사용 (레이아웃/이미지 퀄리티 ↑) │
└──────────────────────────────────────────────────────────────────────────────────────┘
```

**핵심 포인트**
- DOCX 는 "한 덩어리 문서" 로 간주하고 page 분할 안 함 → `pages[0]` 하나로 저장.
- `.doc` (구형 바이너리) / `.ppt` 는 지원 안 됨 → `.docx` / `.pptx` 로 먼저 변환해야.
- `.xls` / `.xlsm` 은 자동으로 `.xlsx` 로 변환 후 처리.

### ❌ CSV — 지원 안 함

- CSV 는 레이아웃/이미지가 없기 때문에 전용 파이프라인 자체가 불필요.
- 필요하면 사전에 JSON/JSONL 로 변환해서 외부에서 주입하세요.

---

## 포맷 × 도구 매트릭스 — 요약표

| 확장자 | 파이프라인 | 구조 분석 (OSS) | 텍스트·표 | 이미지 캡션 | 분할 단위 | 출력 md |
|--------|-----------|------------------|-----------|-------------|-----------|---------|
| `.pdf`  | `pdf_extraction/comprehensive_pdf_analyzer.py` | PyMuPDF + pdfplumber + Pillow | **OSS** (PyMuPDF + pdfplumber) | watsonx VLM | 페이지 | `comprehensive_labeling_with_vlm.md` |
| `.pptx` | `ppt_extraction/comprehensive_presentation_analyzer.py` | python-pptx + LibreOffice + Pillow | **watsonx Text Extraction V2** (fallback: python-pptx) | watsonx VLM | 슬라이드 | `comprehensive_labeling_with_vlm.md` |
| `.xlsx` / `.xlsm` / `.xls` | `excel_extraction/excel_to_jsonl_pipeline.py` | openpyxl + zipfile + LibreOffice | **watsonx Text Extraction V2** | watsonx VLM | 시트 | `assembly.md` (시트 헤더로 split) |
| `.docx` | `watsonx_text_extraction/text_extraction.py` (파일 OR 폴더) | — (원본 업로드) | **watsonx Text Extraction V2** | watsonx VLM | 없음 (전체 1 page) | `assembly.md` |
| `.doc` / `.ppt` | — | — | — | — | — | ❌ 지원 안함 (신포맷으로 변환 필요) |

---

## 공유 코드 (`common/`, `utils/`)

파이프라인 간 중복을 줄이기 위해 자주 반복되던 보일러플레이트를 한곳에 모았다.
각 파이프라인은 이제 env 로딩·watsonx 인증·COS 클라이언트·LibreOffice 변환을
직접 구현하지 않고 이 두 패키지를 import 해서 쓴다.

### `common/config.py`

```python
from common.config import (
    load_env,               # .env 여러 후보 경로 탐색 후 로드
    get_watsonx_credentials,   # Credentials(url, api_key)
    get_api_client,            # APIClient(space_id=..., project_id=...)
    get_space_cos_client,      # (cos_client, bucket) — watsonx space 버킷
    get_master_cos_resource,   # HMAC 기반 failure-case-images 버킷
)
```

이전에는 `load_dotenv(...)` + `Credentials(url=os.environ[...], api_key=...)`
+ `ibm_boto3.client(service_name="s3", ...)` 같은 10~20줄의 동일 코드가
**6개 파일**에 흩어져 있었다. 이제 위 한 줄로 대체.

### `common/libreoffice.py`

```python
from common.libreoffice import convert_to_pdf, find_libreoffice

pdf_path = convert_to_pdf(
    src=Path("report.pptx"),
    out_dir=Path("./pdfs"),
    pdf_filter=None,        # 또는 writer_pdf_Export:{...} 필터 문자열
    timeout=180,
)
```

PPT 파이프라인(mixins/conversion.py), Excel 파이프라인, watsonx_text_extraction
파이프라인에 각각 "soffice/libreoffice 탐색 → subprocess 실행 → 에러 처리" 가
**3번 중복**되어 있던 걸 제거하고 이 함수로 일원화.

### `utils/geometry.py`

PDF / PPTX 공간 분석에서 공용으로 쓰는 bbox overlap / containment / bounds
계산 함수들. 예전엔 `pdf_extraction/utils/` 와 `ppt_extraction/utils/` 에
**동일 파일이 두 벌** 있었음.

### 공유하지 않은 건?

- **`TextExtractionsV2` 호출** — PPT / Excel / watsonx_text_extraction 세
  파이프라인이 다 다른 파라미터와 후처리를 쓴다 (embedded images 모드,
  언어 힌트, OCR 모드 등). 억지로 묶으면 분기 로직이 공유 함수에 다 들어가서
  오히려 읽기 어려워짐.
- **VLM 캡션** — 이미 `ppt_extraction/mixins/vlm.py` 를 PDF analyzer 가
  상속해서 쓰고 있음. 추가로 옮길 필요 없음.
- **각 CLI argparse** — 파이프라인별 옵션이 서로 다름. `extract.py` 가
  외부 래퍼 역할을 하므로 굳이 argparse 스키마까지 통일하진 않음.

---

## 디렉토리 레이아웃

```
src/
├── extract.py                          # ✨ 통합 진입점 (확장자 기반 라우팅)
│
├── common/                             # 🔗 전 파이프라인 공유 헬퍼
│   ├── config.py                       #    load_env + watsonx Credentials + COS 클라이언트
│   └── libreoffice.py                  #    LibreOffice → PDF 변환 단일 진입점
│
├── utils/
│   └── geometry.py                     # PDF/PPTX 공용 기하 헬퍼 (overlap, bounds…)
│
├── pdf_extraction/
│   ├── comprehensive_pdf_analyzer.py   # 엔트리 — ComprehensivePDFAnalyzer
│   └── mixins/
│       ├── pdf_conversion.py           # PyMuPDF 로 페이지 PNG 렌더
│       ├── pdf_text_extraction.py      # PyMuPDF 텍스트 + bbox
│       ├── pdf_spatial.py              # 공간 구조 분석
│       ├── pdf_tables.py               # pdfplumber 표 추출
│       └── pdf_visual_capture.py       # 이미지 영역 크롭
│
├── ppt_extraction/
│   ├── comprehensive_presentation_analyzer.py  # 엔트리
│   ├── mixins/
│   │   ├── conversion.py               # LibreOffice → PDF 변환
│   │   ├── text_extraction.py          # python-pptx + watsonx V2 텍스트
│   │   ├── spatial.py                  # shape 공간 분석
│   │   ├── tables.py                   # 표 추출
│   │   ├── visual_capture.py           # shape → PNG
│   │   ├── smart_grouping.py           # 의미 단위 그룹핑   ← PDF 에서 공유
│   │   ├── reading_order.py            # 독서 순서 재구성    ← PDF 에서 공유
│   │   ├── vlm.py                      # Mistral-Small 3.1 24B caption ← PDF 에서 공유
│   │   ├── labeling.py                 # 최종 md 조립        ← PDF 에서 공유
│   │   └── visualization.py            # 디버그 시각화
│   └── output_sample/                  # 커밋된 참고용 샘플 결과
│       └── F250110-4649_…_analysis/
│           ├── comprehensive_analysis_complete.json
│           ├── comprehensive_labeling_with_vlm.md
│           ├── comprehensive_summary.json
│           └── visual_captures/
│
├── excel_extraction/
│   └── excel_to_jsonl_pipeline.py      # 엔트리 — 시트 분할 + LibreOffice + watsonx V2
│
└── watsonx_text_extraction/            # DOCX 주력 + 공용 watsonx V2 래퍼
    ├── text_extraction.py              # 엔트리 — 파일 OR 폴더, --output 로 로컬 다운로드
    ├── download_cos_results.py         # (독립) COS → 로컬 결과 pull
    ├── delete_cos_results.py           # (독립) 오래된 job 결과 정리
    └── cos_results_utils.py            # (공유) paginator + filter + size 포맷
```

---

## 실행 순서

```bash
source venv/bin/activate

# 모든 원본 문서를 한 번에 처리 — 확장자별로 알맞은 파이프라인으로 라우팅
python src/extract.py --folder client_data/data --output ./output --continue-on-error

# (포맷별로 직접 돌리고 싶다면)
# python src/pdf_extraction/comprehensive_pdf_analyzer.py       --folder client_data/data/
# python src/ppt_extraction/comprehensive_presentation_analyzer.py --folder client_data/data/
# python src/excel_extraction/excel_to_jsonl_pipeline.py        client_data/data/
# python src/watsonx_text_extraction/text_extraction.py         client_data/data/ --output ./output/docx_results

# DOCX 결과를 나중에 따로 받고 싶으면
python src/watsonx_text_extraction/download_cos_results.py --download
```

각 파이프라인은 `--output` (또는 `--output-dir`) 로 지정한 폴더 아래 `<stem>_analysis/`
(PDF / PPTX) 또는 `<stem>/` (Excel, DOCX) 형태로 산출물을 쓴다. 예시:

```
output/
├── F200204-0531_…_analysis/
│   ├── comprehensive_analysis_complete.json
│   ├── comprehensive_labeling_with_vlm.md
│   ├── comprehensive_summary.json
│   └── visual_captures/
├── F200609-0791_…/
│   ├── assembly.md
│   ├── pdf/
│   └── sheets/
└── …
```

---

## 관련 문서

- 에이전트/검색 파이프라인: [`agent-system/README.md`](../agent-system/README.md)
