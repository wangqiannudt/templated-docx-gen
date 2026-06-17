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


def test_find_style_by_candidate():
    doc = Document()
    from docx.enum.style import WD_STYLE_TYPE
    doc.styles.add_style('报告标题', WD_STYLE_TYPE.PARAGRAPH)
    found = build_docx.find_style(doc, ['封皮标题', '报告标题', '主标题'])
    assert found == '报告标题'


def test_find_style_none_when_absent():
    doc = Document()
    assert build_docx.find_style(doc, ['不存在的样式']) is None


def test_find_style_heading_default():
    doc = Document()
    assert build_docx.find_style(doc, ['Heading 1']) == 'Heading 1'
