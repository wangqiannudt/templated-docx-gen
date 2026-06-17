from docx import Document
from docx.shared import Twips
import build_docx


def test_detect_body_width_a4_default():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Twips(11906)
    sec.left_margin = Twips(1440)
    sec.right_margin = Twips(1440)
    assert build_docx.detect_page_body_width(doc) == 11906 - 1440 - 1440


def test_detect_body_width_custom():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Twips(10000)
    sec.left_margin = Twips(800)
    sec.right_margin = Twips(800)
    assert build_docx.detect_page_body_width(doc) == 8400


def test_detect_body_width_fallback_when_missing():
    doc = Document()
    w = build_docx.detect_page_body_width(doc)
    assert isinstance(w, int) and w > 0
