#!/usr/bin/env python3
"""
기획서 파서: PPT/PPTX/PDF에서 텍스트와 이미지를 분리 추출합니다.
에이전트 친화적(토큰 최소화) 형식으로 출력합니다.

이미지 티어:
  Tier 2 - 텍스트 추출 성공 → 코드블록 인라인, 이미지 파일 미저장
  Tier 3 - 시각 해석 필요 → images/ 저장, [IMG: path] 태그로 참조
"""

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


_BULLETS = {"•", "·", "▪", "▸", "►", "–", "○", "●"}


def _join_bullets(texts: list) -> list:
    """PDF에서 '•'가 단독 줄/블록으로 추출될 때 다음 텍스트와 합침"""
    # 블록 내 줄바꿈 전개
    lines = []
    for text in texts:
        for line in text.split("\n"):
            line = line.strip()
            if line:
                lines.append(line)

    result = []
    i = 0
    while i < len(lines):
        if lines[i] in _BULLETS and i + 1 < len(lines):
            result.append(f"- {lines[i + 1]}")
            i += 2
        else:
            result.append(lines[i])
            i += 1
    return result


def _is_meaningful_text(text: str) -> bool:
    """다이어그램 라벨·OCR 노이즈를 걸러내고 실제 읽을 수 있는 텍스트인지 판별.

    다이어그램: 짧은 라벨 다수, 평균 줄 길이 짧음, OCR 쓰레기 문자 다량.
    의미 있는 텍스트: 충분한 단어 수, 평균 줄 길이 ≥ 10, 깨진 문자 비율 < 35%.
    """
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    if not lines:
        return False
    total_words = sum(len(l.split()) for l in lines)
    if total_words < 6:
        return False
    avg_line_len = sum(len(l) for l in lines) / len(lines)
    if avg_line_len < 10:
        return False
    body = text.replace(" ", "").replace("\n", "").replace("\r", "")
    if body:
        _ALLOWED = set('.,;:()[]{}\'\"!?-_=+*/\\@#%^&<>·•▪')
        clean = sum(1 for c in body if c.isalnum() or c in _ALLOWED)
        if clean / len(body) < 0.65:
            return False
    return True


def _is_duplicate_of_page(text: str, page_words: set, threshold: float = 0.6) -> bool:
    words = set(text.split())
    if len(words) < 5:
        return False
    return len(words & page_words) / len(words) >= threshold


def _ocr_image(img_path: Path, warn: bool = False) -> str:
    try:
        import pytesseract
        from PIL import Image, ImageOps, ImageStat
    except ImportError:
        if warn:
            print("  [경고] OCR 불가: pip install pytesseract Pillow && brew install tesseract tesseract-lang")
        return ""
    try:
        img = Image.open(img_path).convert("RGB")
        gray = img.convert("L")
        if ImageStat.Stat(gray).mean[0] < 128:
            img = ImageOps.invert(img)
        w, h = img.size
        img = img.resize((w * 2, h * 2), Image.LANCZOS)
        return pytesseract.image_to_string(img, lang="kor+eng").strip()
    except Exception:
        return ""


def _md_header(path: Path, file_type: str, n_sections: int, tier3_images: list) -> str:
    lines = [
        f"# 문서: {path.name}",
        f"- 형식: {file_type.upper()} | {'슬라이드' if file_type == 'pptx' else '페이지'}: {n_sections}",
        f"- Tier3 이미지 (에이전트 확인 필요): {len(tier3_images)}개",
    ]
    if tier3_images:
        for p in tier3_images:
            lines.append(f"  - {p}")
    return "\n".join(lines)


# ─── PPTX ───────────────────────────────────────────────────────────────────

def extract_pptx(path: Path, out_dir: Path, use_ocr: bool = False) -> dict:
    try:
        from pptx import Presentation
    except ImportError:
        sys.exit("python-pptx 미설치: pip install python-pptx")

    img_dir = out_dir / "images"
    img_dir.mkdir(parents=True, exist_ok=True)

    prs = Presentation(path)
    slides_md = []
    tier3_images = []
    manifest = {"source": str(path), "type": "pptx", "slides": []}

    for slide_idx, slide in enumerate(prs.slides, 1):
        slide_info = {"slide": slide_idx, "title": None, "images": []}

        title = ""
        if slide.shapes.title and slide.shapes.title.text.strip():
            title = slide.shapes.title.text.strip()
            slide_info["title"] = title

        lines = [f"## [S{slide_idx}]{' ' + title if title else ''}"]

        # 텍스트 프레임
        for shape in slide.shapes:
            if shape.has_text_frame and shape != slide.shapes.title:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if not text:
                        continue
                    prefix = "  " * para.level + "-" if para.level > 0 else "-"
                    lines.append(f"{prefix} {text}")

        page_words = set(" ".join(lines).split())

        # 이미지 (relationship 기반)
        img_count = 0
        seen_parts = set()
        for rel in slide.part.rels.values():
            if "image" not in rel.reltype.lower():
                continue
            try:
                img_part = rel.target_part
            except Exception:
                continue
            part_id = id(img_part)
            if part_id in seen_parts:
                continue
            seen_parts.add(part_id)

            ext = img_part.partname.suffix.lstrip(".")
            if ext.lower() in ("emf", "wmf"):
                continue

            img_name = f"slide_{slide_idx}_img_{img_count}.{ext}"
            img_path = img_dir / img_name
            img_path.write_bytes(img_part.blob)

            tier = 3
            extracted_text = ""

            if use_ocr:
                ocr = _ocr_image(img_path, warn=True)
                if ocr and not _is_duplicate_of_page(ocr, page_words):
                    extracted_text = ocr
                    tier = 2

            if tier == 2:
                lines.append(f"\n```\n{extracted_text}\n```")
                img_path.unlink()
                slide_info["images"].append({"id": img_name, "tier": 2})
            else:
                lines.append(f"\n[IMG: images/{img_name}]")
                tier3_images.append(f"images/{img_name}")
                slide_info["images"].append({"id": img_name, "tier": 3, "path": f"images/{img_name}"})

            img_count += 1

        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                lines.append(f"\n> 노트: {notes}")

        slides_md.append("\n".join(lines))
        manifest["slides"].append(slide_info)

    header = _md_header(path, "pptx", len(prs.slides), tier3_images)
    manifest["summary"] = {
        "total_slides": len(prs.slides),
        "tier3_count": len(tier3_images),
        "tier3_images": tier3_images,
    }
    full_md = header + "\n\n---\n\n" + "\n\n---\n\n".join(slides_md)
    return {"markdown": full_md, "manifest": manifest}


# ─── PDF ────────────────────────────────────────────────────────────────────

def extract_pdf(
    path: Path,
    out_dir: Path,
    min_area_ratio: float = 0.02,
    render_scale: float = 2.0,
    use_ocr: bool = False,
) -> dict:
    try:
        import fitz
    except ImportError:
        sys.exit("pymupdf 미설치: pip install pymupdf")

    img_dir = out_dir / "images"
    img_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(path))
    pages_md = []
    tier3_images = []
    manifest = {"source": str(path), "type": "pdf", "pages": []}

    for page_idx, page in enumerate(doc, 1):
        page_info = {"page": page_idx, "images": []}

        page_rect = page.rect
        page_area = page_rect.width * page_rect.height

        # 텍스트 추출 + 불릿 정리
        raw = [b[4].strip() for b in page.get_text("blocks", sort=True) if b[6] == 0 and b[4].strip()]
        cleaned = _join_bullets(raw)
        page_words = set(" ".join(cleaned).split())

        lines = [f"## [P{page_idx}]"]
        lines.extend(cleaned)

        # 이미지 정보
        try:
            img_infos = page.get_image_info(xrefs=True)
        except TypeError:
            img_infos = page.get_image_info()

        seen_xrefs: set = set()
        seen_hashes: set = set()
        mat = fitz.Matrix(render_scale, render_scale)
        img_count = 0

        for info in img_infos:
            xref = info.get("xref", 0)
            if xref in seen_xrefs:
                continue

            bbox = fitz.Rect(info["bbox"])
            if bbox.width * bbox.height < page_area * min_area_ratio:
                continue
            bbox = bbox & page_rect
            if bbox.is_empty:
                continue
            seen_xrefs.add(xref)

            pix = page.get_pixmap(matrix=mat, clip=bbox)
            img_hash = hashlib.md5(pix.tobytes()).hexdigest()
            if img_hash in seen_hashes:
                continue
            seen_hashes.add(img_hash)

            # 텍스트 추출 시도
            clip_text = page.get_text("text", clip=bbox).strip()
            tier = 3
            extracted_text = ""

            if clip_text and _is_meaningful_text(clip_text) and not _is_duplicate_of_page(clip_text, page_words):
                extracted_text = clip_text
                tier = 2
            else:
                ocr = _ocr_image(
                    _save_temp_pix(pix),
                    warn=use_ocr,
                )
                if ocr and _is_meaningful_text(ocr) and not _is_duplicate_of_page(ocr, page_words):
                    extracted_text = ocr
                    tier = 2

            if tier == 2:
                lines.append(f"\n```\n{extracted_text}\n```")
                page_info["images"].append({"id": f"page_{page_idx}_img_{img_count}", "tier": 2})
            else:
                img_name = f"page_{page_idx}_img_{img_count}.png"
                img_path = img_dir / img_name
                pix.save(str(img_path))
                lines.append(f"\n[IMG: images/{img_name}]")
                tier3_images.append(f"images/{img_name}")
                page_info["images"].append({"id": img_name, "tier": 3, "path": f"images/{img_name}"})

            img_count += 1

        pages_md.append("\n".join(lines))
        manifest["pages"].append(page_info)

    doc.close()

    total_imgs = sum(len(p["images"]) for p in manifest["pages"])
    header = _md_header(path, "pdf", len(manifest["pages"]), tier3_images)
    manifest["summary"] = {
        "total_pages": len(manifest["pages"]),
        "total_images": total_imgs,
        "tier3_count": len(tier3_images),
        "tier3_images": tier3_images,
    }
    full_md = header + "\n\n---\n\n" + "\n\n---\n\n".join(pages_md)
    return {"markdown": full_md, "manifest": manifest}


def _build_guide(source: Path, section: str, n: int, tier3_images: list) -> str:
    if tier3_images:
        img_list = "\n".join(f"- {p}" for p in tier3_images)
        img_section = f"\n시각 확인 필요 이미지:\n{img_list}"
    else:
        img_section = "\n이미지 없음 — content.md만 읽으면 됩니다."

    return (
        f"원본: {source.name} | {section}: {n}\n"
        f"읽기: content.md → `[IMG: path]` 태그 위치에서 해당 이미지 로드"
        f"{img_section}"
    )


def _save_temp_pix(pix) -> Path:
    """OCR용 임시 파일 저장 후 경로 반환"""
    tmp = Path(tempfile.mktemp(suffix=".png"))
    pix.save(str(tmp))
    return tmp


# ─── PPT 변환 ────────────────────────────────────────────────────────────────

def convert_ppt_to_pptx(ppt_path: Path) -> Path:
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not libreoffice:
        sys.exit(".ppt 변환에 LibreOffice 필요\n설치: brew install --cask libreoffice")

    tmp_dir = Path(tempfile.mkdtemp())
    result = subprocess.run(
        [libreoffice, "--headless", "--convert-to", "pptx", "--outdir", str(tmp_dir), str(ppt_path)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        sys.exit(f"PPT 변환 실패:\n{result.stderr}")
    converted = list(tmp_dir.glob("*.pptx"))
    if not converted:
        sys.exit("변환된 파일을 찾을 수 없습니다.")
    return converted[0]


# ─── 진입점 ──────────────────────────────────────────────────────────────────

def run(input_path: Path, out_dir: Path, min_area_ratio: float, use_ocr: bool) -> None:
    suffix = input_path.suffix.lower()

    if suffix == ".pptx":
        result = extract_pptx(input_path, out_dir, use_ocr=use_ocr)
        key = "slides"
    elif suffix == ".ppt":
        print("PPT → PPTX 변환 중...")
        result = extract_pptx(convert_ppt_to_pptx(input_path), out_dir, use_ocr=use_ocr)
        key = "slides"
    elif suffix == ".pdf":
        result = extract_pdf(input_path, out_dir, min_area_ratio=min_area_ratio, use_ocr=use_ocr)
        key = "pages"
    else:
        sys.exit(f"지원하지 않는 형식: {suffix}")

    md_path = out_dir / "content.md"
    md_path.write_text(result["markdown"], encoding="utf-8")

    manifest = result["manifest"]
    manifest_path = out_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    s = manifest["summary"]
    tier3 = s.get("tier3_images", [])
    n = s.get(f"total_{key}", s.get("total_pages", 0))
    section = "슬라이드" if key == "slides" else "페이지"

    guide = _build_guide(input_path, section, n, tier3)
    (out_dir / "GUIDE.md").write_text(guide, encoding="utf-8")

    print(f"\n추출 완료: {out_dir}")
    print(f"  {section}: {n}개")
    print(f"  Tier3 이미지 (에이전트 확인): {s['tier3_count']}개")
    print(f"  텍스트: {md_path.name}")
    print(f"  매니페스트: {manifest_path.name}")


def main():
    parser = argparse.ArgumentParser(description="PPT/PPTX/PDF → 에이전트 친화적 마크다운 추출")
    parser.add_argument("file", type=Path)
    parser.add_argument("-o", "--output", type=Path, default=None)
    parser.add_argument("--min-img-ratio", type=float, default=0.02, metavar="RATIO")
    parser.add_argument("--ocr", action="store_true", help="OCR 경고 메시지 활성화")
    args = parser.parse_args()

    input_path = args.file.resolve()
    if not input_path.exists():
        sys.exit(f"파일 없음: {input_path}")

    out_dir = args.output or input_path.parent / f"{input_path.stem}_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    run(input_path, out_dir, min_area_ratio=args.min_img_ratio, use_ocr=args.ocr)


if __name__ == "__main__":
    main()
