"""钉住 build_docx.parse_md / strip_title_num 现有正确行为（回归护栏）。"""
import build_docx


def test_strip_title_num_arabic():
    assert build_docx.strip_title_num("1.1 概述") == "概述"
    assert build_docx.strip_title_num("2 试验目的") == "试验目的"


def test_strip_title_num_keeps_chinese_numbering():
    assert build_docx.strip_title_num("一、概述") == "一、概述"
    assert build_docx.strip_title_num("（一）背景") == "（一）背景"


def test_strip_title_num_no_prefix():
    assert build_docx.strip_title_num("背景介绍") == "背景介绍"


def test_parse_md_headings():
    items = build_docx.parse_md("# 一级\n## 二级\n### 三级\n正文")
    assert ('h1', '一级') in items
    assert ('h2', '二级') in items
    assert ('h3', '三级') in items
    assert ('p', '正文') in items


def test_parse_md_image():
    items = build_docx.parse_md("![图1 示意图](assets/a.png)")
    assert items == [('image', ('图1 示意图', 'assets/a.png'))]


def test_parse_md_table():
    md = "| 序号 | 名称 |\n|---|---|\n| 1 | foo |\n"
    items = build_docx.parse_md(md)
    tbl = [c for t, c in items if t == 'table'][0]
    assert tbl[0] == ['序号', '名称']
    assert tbl[1] == ['1', 'foo']


def test_parse_md_list_items():
    items = build_docx.parse_md("- 项A\n1. 项B")
    assert ('li', '项A') in items
    assert ('li', '项B') in items


def test_parse_md_skips_blank_and_newpage():
    items = build_docx.parse_md("正文A\n\n\\newpage\n正文B")
    assert ('p', '正文A') in items
    assert ('p', '正文B') in items
    assert all(t != '' for t, _ in items)
