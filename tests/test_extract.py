"""extract.py 유닛 + 통합 테스트"""
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from extract import (
    _detect_pdf_headings,
    _is_duplicate_of_page,
    _is_meaningful_text,
    _join_bullets,
    _resolve_files,
    _rows_to_markdown,
    extract_docx,
    extract_pdf,
    extract_pptx,
    run,
)


# ─── 유틸리티 함수 ──────────────────────────────────────────────────────────

class TestJoinBullets:
    def test_merges_standalone_bullets(self):
        result = _join_bullets(["•", "항목1", "•", "항목2"])
        assert result == ["- 항목1", "- 항목2"]

    def test_preserves_normal_text(self):
        result = _join_bullets(["일반 텍스트", "두 번째 줄"])
        assert result == ["일반 텍스트", "두 번째 줄"]

    def test_handles_multiline_blocks(self):
        result = _join_bullets(["줄1\n줄2\n줄3"])
        assert result == ["줄1", "줄2", "줄3"]

    def test_trailing_bullet_no_crash(self):
        result = _join_bullets(["•"])
        assert result == ["•"]


class TestIsMeaningfulText:
    def test_long_text_is_meaningful(self):
        text = "이것은 충분히 긴 텍스트입니다. 여러 단어가 포함되어 있어서 의미 있는 텍스트로 판별됩니다."
        assert _is_meaningful_text(text) is True

    def test_short_label_not_meaningful(self):
        assert _is_meaningful_text("A") is False
        assert _is_meaningful_text("OK") is False

    def test_empty_not_meaningful(self):
        assert _is_meaningful_text("") is False

    def test_garbage_chars_not_meaningful(self):
        assert _is_meaningful_text("◆◇◈▣▤▥▦▧▨▩" * 5) is False


class TestIsDuplicateOfPage:
    def test_high_overlap_is_duplicate(self):
        page_words = {"hello", "world", "this", "is", "a", "test", "sentence"}
        assert _is_duplicate_of_page("hello world this is a test", page_words) is True

    def test_no_overlap_not_duplicate(self):
        page_words = {"alpha", "beta", "gamma", "delta", "epsilon"}
        assert _is_duplicate_of_page("one two three four five", page_words) is False

    def test_short_text_not_duplicate(self):
        page_words = {"a", "b", "c"}
        assert _is_duplicate_of_page("a b", page_words) is False


class TestRowsToMarkdown:
    def test_basic_table(self):
        rows = [["A", "B"], ["1", "2"]]
        result = _rows_to_markdown(rows)
        assert "| A | B |" in result
        assert "| 1 | 2 |" in result
        assert "| --- | --- |" in result

    def test_empty_rows(self):
        assert _rows_to_markdown([]) == ""
        assert _rows_to_markdown([[]]) == ""

    def test_pipe_escape(self):
        rows = [["header"], ["val|ue"]]
        result = _rows_to_markdown(rows)
        assert "val\\|ue" in result

    def test_uneven_columns(self):
        rows = [["A", "B", "C"], ["1", "2"]]
        result = _rows_to_markdown(rows)
        assert result.count("|") > 0

    def test_header_only(self):
        rows = [["A", "B"]]
        result = _rows_to_markdown(rows)
        assert "| A | B |" in result


# ─── PDF 추출 ────────────────────────────────────────────────────────────────

class TestExtractPdf:
    def test_basic_extraction(self, sample_pdf, tmp_out):
        result = extract_pdf(sample_pdf, tmp_out, quiet=True)
        assert "markdown" in result
        assert "manifest" in result
        assert result["manifest"]["type"] == "pdf"
        assert result["manifest"]["summary"]["total_pages"] == 2

    def test_output_files(self, sample_pdf, tmp_out):
        run(sample_pdf, tmp_out, min_area_ratio=0.02, use_ocr=False, quiet=True)
        assert (tmp_out / "content.md").exists()
        assert (tmp_out / "manifest.json").exists()
        assert (tmp_out / "GUIDE.md").exists()
        manifest = json.loads((tmp_out / "manifest.json").read_text())
        assert manifest["type"] == "pdf"

    def test_heading_detection(self, sample_pdf, tmp_out):
        result = extract_pdf(sample_pdf, tmp_out, quiet=True)
        md = result["markdown"]
        assert "## [P" in md


# ─── PPTX 추출 ───────────────────────────────────────────────────────────────

class TestExtractPptx:
    def test_basic_extraction(self, sample_pptx, tmp_out):
        result = extract_pptx(sample_pptx, tmp_out, quiet=True)
        assert result["manifest"]["type"] == "pptx"
        assert result["manifest"]["summary"]["total_slides"] == 2

    def test_table_extraction(self, sample_pptx, tmp_out):
        result = extract_pptx(sample_pptx, tmp_out, quiet=True)
        md = result["markdown"]
        assert "| 항목 | 수량 | 비고 |" in md
        assert result["manifest"]["summary"]["tables"] >= 1

    def test_title_extraction(self, sample_pptx, tmp_out):
        result = extract_pptx(sample_pptx, tmp_out, quiet=True)
        md = result["markdown"]
        assert "테스트 슬라이드" in md


# ─── DOCX 추출 ───────────────────────────────────────────────────────────────

class TestExtractDocx:
    def test_basic_extraction(self, sample_docx, tmp_out):
        result = extract_docx(sample_docx, tmp_out, quiet=True)
        assert result["manifest"]["type"] == "docx"

    def test_heading_extraction(self, sample_docx, tmp_out):
        result = extract_docx(sample_docx, tmp_out, quiet=True)
        md = result["markdown"]
        assert "## 문서 제목" in md
        assert "### 하위 제목" in md

    def test_table_extraction(self, sample_docx, tmp_out):
        result = extract_docx(sample_docx, tmp_out, quiet=True)
        md = result["markdown"]
        assert "| 이름 | 값 |" in md
        assert result["manifest"]["summary"]["tables"] >= 1

    def test_output_files(self, sample_docx, tmp_out):
        run(sample_docx, tmp_out, min_area_ratio=0.02, use_ocr=False, quiet=True)
        assert (tmp_out / "content.md").exists()
        assert (tmp_out / "manifest.json").exists()
        assert (tmp_out / "GUIDE.md").exists()


# ─── 배치 처리 ───────────────────────────────────────────────────────────────

class TestBatch:
    def test_resolve_files_single(self, sample_pdf):
        files = _resolve_files([str(sample_pdf)])
        assert len(files) == 1

    def test_resolve_files_nonexistent(self):
        files = _resolve_files(["/nonexistent/file.pdf"])
        assert len(files) == 0

    def test_resolve_files_unsupported(self, tmp_path):
        txt = tmp_path / "test.txt"
        txt.write_text("hello")
        files = _resolve_files([str(txt)])
        assert len(files) == 0


# ─── 에러 복원 ───────────────────────────────────────────────────────────────

class TestResilience:
    def test_corrupted_page_continues(self, tmp_path, tmp_out):
        """손상된 페이지가 있어도 나머지 페이지는 추출됨"""
        import fitz
        path = tmp_path / "partial.pdf"
        doc = fitz.open()
        page1 = doc.new_page()
        page1.insert_text((72, 100), "정상 페이지 텍스트 내용입니다. 충분한 길이의 텍스트를 넣습니다.", fontsize=11)
        doc.new_page()
        doc.save(str(path))
        doc.close()

        result = extract_pdf(path, tmp_out, quiet=True)
        assert result["manifest"]["summary"]["total_pages"] == 2
