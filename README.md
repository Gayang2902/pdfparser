# pdfparser

LLM 에이전트를 위한 PDF / PPT / PPTX 파서입니다.
문서에서 텍스트와 이미지를 분리 추출하여 토큰을 최소화하고 에이전트가 읽기 쉬운 마크다운으로 변환합니다.

## 개요

LLM 에이전트에 문서를 통째로 넘기면 이미지와 텍스트가 뒤섞여 토큰 낭비가 심합니다.
이 도구는 텍스트는 마크다운으로, 이미지는 티어로 분류하여 최소한의 파일만 에이전트에 전달합니다.

### 이미지 티어 분류

| 티어 | 조건 | 처리 방식 |
|------|------|-----------|
| **Tier 2** | OCR로 텍스트 추출 가능 (코드 스크린샷 등) | 마크다운 코드블록 인라인 삽입, 이미지 파일 미저장 |
| **Tier 3** | 시각 해석 필요 (차트, 플로우차트, 다이어그램 등) | `images/` 저장 후 `[IMG: path]` 태그 참조 |

## 설치

```bash
bash setup.sh
```

tesseract, pymupdf, python-pptx 등 의존성을 자동 설치하고 `parse` 명령어를 등록합니다.

## 사용법

```bash
# 기본 사용
parse document.pdf

# 출력 디렉터리 지정
parse document.pdf -o ./output

# 최소 이미지 면적 비율 조정 (기본 2%)
parse document.pdf --min-img-ratio 0.05

# OCR 경고 활성화
parse document.pdf --ocr
```

PPT 파일은 LibreOffice가 필요합니다 (`brew install --cask libreoffice`).

## 출력 구조

```
output/
├── GUIDE.md        # 에이전트 독서 가이드 (먼저 읽을 것)
├── content.md      # 전체 텍스트 (메인 읽기 대상)
├── images/         # Tier3 이미지 (시각 확인 필요한 것만)
└── manifest.json   # 구조 메타데이터
```

에이전트는 `GUIDE.md` → `content.md` 순으로 읽고, `[IMG: ...]` 태그를 만나면 해당 이미지를 로드합니다.

## 지원 형식

- PDF
- PPTX
- PPT (LibreOffice 필요)
