#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PARSE_SH="$SCRIPT_DIR/parse.sh"
LINK_TARGET="/usr/local/bin/parse"

echo "=== 기획서 파서 설치 ==="

# Python 확인
if ! command -v python3 &>/dev/null; then
    echo "오류: python3가 없습니다. https://www.python.org 에서 설치하세요." >&2
    exit 1
fi

# Homebrew 확인
if ! command -v brew &>/dev/null; then
    echo "오류: Homebrew가 없습니다. https://brew.sh 에서 설치하세요." >&2
    exit 1
fi

# Python 패키지
echo ""
echo "[1/3] Python 패키지 설치 중..."
pip install -q -r "$SCRIPT_DIR/requirements.txt"
pip install -q pytesseract

# Tesseract
echo ""
echo "[2/3] Tesseract 설치 중..."
if command -v tesseract &>/dev/null; then
    echo "  이미 설치됨: $(tesseract --version 2>&1 | head -1)"
else
    brew install tesseract tesseract-lang
fi

# parse 명령어 등록
echo ""
echo "[3/3] parse 명령어 등록 중..."
chmod +x "$PARSE_SH"

if [[ -L "$LINK_TARGET" ]]; then
    rm "$LINK_TARGET"
fi

if ln -sf "$PARSE_SH" "$LINK_TARGET" 2>/dev/null; then
    echo "  등록 완료: parse → $PARSE_SH"
else
    # /usr/local/bin 권한 없으면 sudo 시도
    if sudo ln -sf "$PARSE_SH" "$LINK_TARGET" 2>/dev/null; then
        echo "  등록 완료 (sudo): parse → $PARSE_SH"
    else
        # 대안: ~/.local/bin
        mkdir -p "$HOME/.local/bin"
        ln -sf "$PARSE_SH" "$HOME/.local/bin/parse"
        echo "  등록 완료: $HOME/.local/bin/parse"
        echo ""
        echo "  PATH에 추가하려면 ~/.zshrc 또는 ~/.bashrc에 다음 줄을 추가하세요:"
        echo "    export PATH=\"\$HOME/.local/bin:\$PATH\""
    fi
fi

echo ""
echo "=== 설치 완료 ==="
echo ""
echo "사용법:"
echo "  parse 기획서.pdf"
echo "  parse 기획서.pdf --ocr"
echo "  parse 기획서.pptx -o ./결과"
echo "  parse --help"
