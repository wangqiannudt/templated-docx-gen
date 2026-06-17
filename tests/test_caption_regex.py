import build_docx


def test_figure_caption_with_space():
    assert build_docx.match_caption("图1 示意图") == ('图', '示意图')


def test_figure_caption_with_dot():
    assert build_docx.match_caption("图2.原理框图") == ('图', '原理框图')


def test_table_caption():
    assert build_docx.match_caption("表 2 测试结果") == ('表', '测试结果')


def test_table_caption_chinese_sep():
    assert build_docx.match_caption("表3、指标对比") == ('表', '指标对比')


def test_plain_word_not_caption():
    assert build_docx.match_caption("图片说明如下") is None
    assert build_docx.match_caption("图形结构") is None
    assert build_docx.match_caption("表格内容") is None


def test_caption_without_desc_not_caption():
    assert build_docx.match_caption("图1") is None
