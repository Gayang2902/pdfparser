#!/usr/bin/env bash
set -euo pipefail

REAL_SCRIPT="$(python3 -c "import os,sys; print(os.path.realpath(sys.argv[1]))" "${BASH_SOURCE[0]}")"
SCRIPT_DIR="$(dirname "$REAL_SCRIPT")"
EXTRACT="$SCRIPT_DIR/extract.py"

usage() {
    cat <<EOF
사용법: parse.sh [옵션] <파일...>

지원 형식: .pdf  .pptx  .ppt  .docx

옵션:
  -o <디렉터리>    출력 디렉터리 (기본값: <파일명>_output/)
  --ocr            이미지 안의 텍스트 OCR 추출
  --ratio <숫자>   PDF 이미지 최소 크기 필터, 0~1 (기본값: 0.02)
  -q, --quiet      진행률·결과 출력 억제
  -h, --help       이 도움말 출력

예시:
  parse.sh 기획서.pdf
  parse.sh 기획서.pdf -o ./결과
  parse.sh 기획서.pdf --ocr
  parse.sh 기획서.pptx -o ./출력
  parse.sh 기획서.pdf --ratio 0.05 --ocr
  parse.sh *.pdf -o ./전체결과
  parse.sh 기획서.docx
  parse.sh 기획서.pdf 기획서.docx -q

OCR 의존성 설치:
  brew install tesseract tesseract-lang
  pip install pytesseract
EOF
}

FILES=()
OUTPUT_ARGS=()
OCR_FLAG=""
RATIO_ARGS=()
QUIET_FLAG=""

while [[ $# -gt 0 ]]; do
    case "$1" in
        -h|--help)
            usage; exit 0 ;;
        -o)
            OUTPUT_ARGS=("-o" "$2"); shift 2 ;;
        --ocr)
            OCR_FLAG="--ocr"; shift ;;
        --ratio)
            RATIO_ARGS=("--min-img-ratio" "$2"); shift 2 ;;
        -q|--quiet)
            QUIET_FLAG="--quiet"; shift ;;
        -*)
            echo "알 수 없는 옵션: $1" >&2; usage; exit 1 ;;
        *)
            FILES+=("$1"); shift ;;
    esac
done

if [[ ${#FILES[@]} -eq 0 ]]; then
    usage; exit 1
fi

for f in "${FILES[@]}"; do
    if [[ ! -f "$f" ]]; then
        echo "오류: 파일을 찾을 수 없습니다 — $f" >&2
        exit 1
    fi
done

ARGS=("${FILES[@]}")
[[ ${#OUTPUT_ARGS[@]} -gt 0 ]] && ARGS+=("${OUTPUT_ARGS[@]}")
[[ ${#RATIO_ARGS[@]} -gt 0 ]] && ARGS+=("${RATIO_ARGS[@]}")
[[ -n "$OCR_FLAG" ]] && ARGS+=("$OCR_FLAG")
[[ -n "$QUIET_FLAG" ]] && ARGS+=("$QUIET_FLAG")

python3 "$EXTRACT" "${ARGS[@]}"
