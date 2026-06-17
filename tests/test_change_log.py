from docx import Document
import build_docx


def _make_table(ncols):
    doc = Document()
    t = doc.add_table(rows=3, cols=ncols)
    for ci, h in enumerate(['版本号', '日期', '编制', '类型', '说明'][:ncols]):
        t.rows[0].cells[ci].text = h
    t.rows[1].cells[0].text = 'OLD'
    return doc, t


def test_apply_change_log_clears_and_rebuilds():
    doc, t = _make_table(5)
    build_docx.apply_change_log(t, [['V1.0', '2025.12', '编制单位', '增加', '初建']])
    assert t.rows[0].cells[0].text == '版本号'
    assert t.rows[1].cells[0].text == 'V1.0'
    assert len(t.rows) == 2


def test_apply_change_log_multi_rows():
    doc, t = _make_table(5)
    rows = [
        ['V0.0', '2025.11', '甲', '增加', '初建'],
        ['V1.0', '2025.12', '乙', '修订', '补充'],
    ]
    build_docx.apply_change_log(t, rows)
    assert len(t.rows) == 3
    assert t.rows[1].cells[0].text == 'V0.0'
    assert t.rows[2].cells[4].text == '补充'


def test_apply_change_log_aligns_fewer_cols():
    doc, t = _make_table(5)
    build_docx.apply_change_log(t, [['V1.0', '2025.12', '编制单位']])
    assert t.rows[1].cells[0].text == 'V1.0'
    assert t.rows[1].cells[2].text == '编制单位'
    assert t.rows[1].cells[3].text == ''
    assert t.rows[1].cells[4].text == ''


def test_apply_change_log_empty_rows_leaves_table():
    doc, t = _make_table(5)
    before = len(t.rows)
    build_docx.apply_change_log(t, [])
    assert len(t.rows) == before
    assert t.rows[1].cells[0].text == 'OLD'
