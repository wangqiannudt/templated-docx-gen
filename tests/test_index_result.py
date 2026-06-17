import build_docx


def test_toc_style_is_index():
    assert build_docx.is_index_result('toc 1', '任意') is True
    assert build_docx.is_index_result('TOC Heading', 'x') is True


def test_table_of_figures_style_is_index():
    assert build_docx.is_index_result('table of figures', 'x') is True


def test_no_entries_found_text_is_index():
    assert build_docx.is_index_result('Normal', 'No table of figures entries found.') is True


def test_normal_body_not_index():
    assert build_docx.is_index_result('Normal', '正文段落') is False
