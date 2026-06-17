import build_docx


def test_narrow_cols_get_fixed_width():
    header = ['序号', '名称', '数量']
    w = build_docx.compute_col_widths(header, ncols=3, narrow_keys=['序号', '数量'],
                                      total_w=9000, narrow_w=1000)
    assert w[0] == 1000
    assert w[2] == 1000
    assert w[1] == 7000


def test_all_wide_even_split():
    header = ['名称', '说明']
    w = build_docx.compute_col_widths(header, ncols=2, narrow_keys=[],
                                      total_w=8000, narrow_w=1000)
    assert w == [4000, 4000]


def test_sum_near_total():
    header = ['序号', '名称', '时间', '备注']
    w = build_docx.compute_col_widths(header, ncols=4, narrow_keys=['序号', '时间'],
                                      total_w=8505, narrow_w=1000)
    assert abs(sum(w) - 8505) <= 3


def test_short_header_pad_safe():
    header = ['序号']
    w = build_docx.compute_col_widths(header, ncols=3, narrow_keys=['序号'],
                                      total_w=9000, narrow_w=1000)
    assert len(w) == 3
