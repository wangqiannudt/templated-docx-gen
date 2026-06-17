#!/usr/bin/env python3
"""基于模板docx生成正式文档：保留封面/变更记录/节结构，正文用模板样式（标题自动编号）"""
import sys, re
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

def strip_title_num(text):
    """去掉阿拉伯数字编号前缀（1/1.1，试验类Heading自动编号）；保留中文编号（一、/（一），总结报告Heading无自动编号需文字带编号）"""
    text = re.sub(r'^\d+(\.\d+)*\s+', '', text)
    return text.strip()

def add_runs(paragraph, text):
    clean = text.replace('**', '').replace('`', '')
    paragraph.add_run(clean)

def add_seq_field(paragraph, label, result_text='1'):
    """添加SEQ自动编号域（fldChar标准方式），Word打开时随updateFields自动编号"""
    from docx.oxml import OxmlElement
    def mk(fieldtype=None, instr=None, text=None):
        r = OxmlElement('w:r')
        if fieldtype:
            fc = OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'), fieldtype); r.append(fc)
        elif instr:
            it = OxmlElement('w:instrText'); it.text = instr; r.append(it)
        elif text is not None:
            t = OxmlElement('w:t'); t.text = text; r.append(t)
        return r
    _begin = mk('begin')
    _begin.find(qn('w:fldChar')).set(qn('w:dirty'), 'true')
    paragraph._p.append(_begin)
    paragraph._p.append(mk(instr=' SEQ ' + label + ' \\* ARABIC '))
    paragraph._p.append(mk('separate'))
    paragraph._p.append(mk(text=str(result_text)))
    paragraph._p.append(mk('end'))

def clear_frontmatter_field_results(doc):
    """清掉模板目录/图目录/表目录的旧显示值，保留域结构，打开文档更新域后重建。"""
    for p in doc.paragraphs:
        try:
            style_name = p.style.name.lower()
        except Exception:
            style_name = ''
        if p.style.name == 'Heading 1':
            break
        text = p.text.strip()
        is_index_result = (
            style_name.startswith('toc')
            or 'table of figures' in style_name
            or text == 'No table of figures entries found.'
        )
        if not is_index_result:
            continue
        for r in p._p.findall(qn('w:r')):
            if r.find(qn('w:instrText')) is not None:
                continue
            for t in r.findall(qn('w:t')):
                t.text = ''

def parse_md(md_text):
    items = []
    lines = md_text.split('\n')
    start = 0
    for idx, line in enumerate(lines):
        if line.startswith('# '):
            start = idx
            break
    lines = lines[start:]
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        img_m = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)\s*$', line)
        if img_m:
            items.append(('image', (img_m.group(1), img_m.group(2))))
        elif line.startswith('# '):
            items.append(('h1', strip_title_num(line[2:])))
        elif line.startswith('## '):
            items.append(('h2', strip_title_num(line[3:])))
        elif line.startswith('### '):
            items.append(('h3', strip_title_num(line[4:])))
        elif line.startswith('|') and '---' not in line and '文档编号' not in line and '版本号' not in line[:20]:
            tbl = []
            while i < len(lines) and lines[i].startswith('|'):
                if '---' not in lines[i]:
                    cells = [c.strip() for c in lines[i].strip('|').split('|')]
                    tbl.append(cells)
                i += 1
            if tbl:
                items.append(('table', tbl))
            continue
        elif line.startswith('- '):
            items.append(('li', line[2:].strip()))
        elif re.match(r'^\d+\.\s', line):
            items.append(('li', re.sub(r'^\d+\.\s*', '', line).strip()))
        elif line.strip() == '' or line.strip() == '\\newpage':
            pass
        else:
            items.append(('p', line.strip()))
        i += 1
    return items

def build(template, md_path, out, cover_title, cover_date, change_row, strip_heading_num=False):
    doc = Document(template)
    body = doc.element.body

    # 1. 修改封面主标题
    title_changed = False
    for p in doc.paragraphs:
        if p.style.name == '封皮标题' and p.text.strip() and not title_changed:
            if p.runs:
                p.runs[0].text = cover_title
                for r in p.runs[1:]:
                    r.text = ''
            title_changed = True
            break

    # 2. 修改封面日期
    for p in doc.paragraphs:
        if p.style.name == '封皮单位' and ('年' in p.text or re.search(r'\d{4}', p.text)):
            if p.runs:
                p.runs[0].text = cover_date
                for r in p.runs[1:]:
                    r.text = ''
            break

    # 3. 变更记录表：清空数据行，重建（初建何姗 + V1.0王倩，2025.11）
    if doc.tables:
        t = doc.tables[0]
        first_row_text = ''.join(c.text for c in t.rows[0].cells)
        if '版本号' in first_row_text and len(t.rows) >= 2:
            while len(t.rows) > 1:
                tr = t.rows[-1]._tr
                tr.getparent().remove(tr)
            rows_data = [
                ['V0.0', '2025.11', '何姗', '增加', '初建'],
                ['V1.0', '2025.11', '王倩', '增加', '结合课题研究成果形成正式版本'],
            ]
            for rd in rows_data:
                new_tr = deepcopy(t.rows[0]._tr)
                t._tbl.append(new_tr)
                new_row = t.rows[-1]
                for ci, cell in enumerate(new_row.cells):
                    if ci < len(rd):
                        for para in cell.paragraphs:
                            for r in para.runs:
                                r.text = ''
                        if cell.paragraphs and cell.paragraphs[0].runs:
                            cell.paragraphs[0].runs[0].text = rd[ci]
                        else:
                            cell.paragraphs[0].text = rd[ci]

    # 4. 删除原正文区域（第一个Heading1到结尾）
    # 动态找Heading1的styleId（不同模板可能是Heading1/1/11等）
    style_id_to_name = {}
    h1_sid = None
    for s in doc.styles:
        try:
            style_id_to_name[s.style_id] = s.name
            if s.name == 'Heading 1':
                h1_sid = s.style_id
        except Exception:
            pass
    h1_sid = h1_sid or 'Heading1'
    to_remove = []
    found_h1 = False
    for child in list(body.iterchildren()):
        if child.tag == qn('w:p'):
            pPr = child.find(qn('w:pPr'))
            style_val = ''
            if pPr is not None:
                pStyle = pPr.find(qn('w:pStyle'))
                if pStyle is not None:
                    style_val = pStyle.get(qn('w:val')) or ''
            style_name = style_id_to_name.get(style_val, '')
            if style_val in (h1_sid, 'Heading1', '1') or style_name in ('Heading 1', '标题 1'):
                found_h1 = True
        if found_h1:
            to_remove.append(child)
    for elem in to_remove:
        body.remove(elem)

    # 5. 解析md，添加正文
    with open(md_path, 'r', encoding='utf-8') as f:
        md_text = f.read()
    items = parse_md(md_text)

    tbl_style = None
    for s in doc.styles:
        if 'Table Grid' in s.name or '网格' in s.name:
            tbl_style = s.name
            break

    seq_counters = {'图': 0, '表': 0}
    for typ, content in items:
        if typ == 'h1':
            p = doc.add_paragraph(style='Heading 1')
            add_runs(p, content)
        elif typ == 'h2':
            p = doc.add_paragraph(style='Heading 2')
            add_runs(p, content)
        elif typ == 'h3':
            p = doc.add_paragraph(style='Heading 3')
            add_runs(p, content)
        elif typ == 'p':
            clean = content.replace('**', '').replace('`', '')
            table_m = re.match(r'^(表)\s*\d+\s*(.*)', clean)
            fig_m = re.match(r'^(图)\s*\d*\s*(\S.*)', clean)
            if table_m:
                seq_counters['表'] += 1
                p = doc.add_paragraph()
                try: p.style = doc.styles['表题注']
                except Exception: pass
                p.add_run(table_m.group(1) + ' ')
                add_seq_field(p, '表', seq_counters['表'])
                p.add_run(' ' + table_m.group(2))
            elif fig_m:
                seq_counters['图'] += 1
                p = doc.add_paragraph()
                try: p.style = doc.styles['图题注']
                except Exception: pass
                p.add_run(fig_m.group(1) + ' ')
                add_seq_field(p, '图', seq_counters['图'])
                p.add_run(' ' + fig_m.group(2))
            else:
                p = doc.add_paragraph(style='Normal')
                add_runs(p, content)
        elif typ == 'image':
            caption, img_path = content
            pic_p = doc.add_paragraph()
            pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pic_run = pic_p.add_run()
            import os
            _full = img_path if os.path.isabs(img_path) else os.path.join(os.path.dirname(md_path), img_path)
            try:
                pic_run.add_picture(_full, width=Cm(13))
            except Exception as e:
                print('图片插入失败', _full, e)
            cap_p = doc.add_paragraph()
            try: cap_p.style = doc.styles['图题注']
            except Exception: pass
            cm = re.match(r'^(图)\s*\d*\s*(\S.*)', caption)
            if cm:
                seq_counters['图'] += 1
                cap_p.add_run(cm.group(1)+' ')
                add_seq_field(cap_p, '图', seq_counters['图'])
                cap_p.add_run(' '+cm.group(2))
            else:
                cap_p.add_run(caption)
        elif typ == 'li':
            try:
                p = doc.add_paragraph(style='List Paragraph')
            except Exception:
                p = doc.add_paragraph(style='Normal')
                p.paragraph_format.left_indent = Cm(0.74)
            add_runs(p, content)
        elif typ == 'table':
            rows = content
            if not rows:
                continue
            ncols = max(len(r) for r in rows)
            tbl = doc.add_table(rows=len(rows), cols=ncols)
            try:
                tbl.style = tbl_style or 'Table Grid'
            except Exception:
                pass
            from docx.oxml import OxmlElement
            from docx.shared import Twips
            tblPr_el = tbl._element.tblPr
            if tblPr_el is None:
                tblPr_el = OxmlElement('w:tblPr')
                tbl._element.insert(0, tblPr_el)
            # 边框
            tblBorders = OxmlElement('w:tblBorders')
            for edge in ('top','left','bottom','right','insideH','insideV'):
                el = OxmlElement('w:' + edge)
                el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), '4'); el.set(qn('w:space'), '0'); el.set(qn('w:color'), '000000')
                tblBorders.append(el)
            tblPr_el.append(tblBorders)
            # 总宽度(版心8505 dxa≈15cm) + fixed布局 + tblInd=0
            for tag, attrs in (('w:tblW', [('w:w','8505'),('w:type','dxa')]),
                               ('w:tblLayout', [('w:type','fixed')]),
                               ('w:tblInd', [('w:w','0'),('w:type','dxa')])):
                el = OxmlElement(tag)
                for k,v in attrs: el.set(qn(k), v)
                tblPr_el.append(el)
            # 单元格内边距
            tblMar = OxmlElement('w:tblCellMar')
            for edge, w in (('top','60'),('bottom','60'),('left','108'),('right','108')):
                el = OxmlElement('w:'+edge); el.set(qn('w:w'), w); el.set(qn('w:type'),'dxa'); tblMar.append(el)
            tblPr_el.append(tblMar)
            tbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # 列宽：短列(序号/数量/时间/阶段/年度/是否/编号/级别/日期/方式/用途/类型)窄，其余均分宽
            narrow_keys = ('序号','数量','时间','阶段','年度','是否','编号','级别','日期','方式','用途','类型')
            header = rows[0] if rows else []
            total_w = 8505
            narrow_w = 1000
            col_w = [narrow_w if any(k in (header[ci] if ci < len(header) else '') for k in narrow_keys) else 0 for ci in range(ncols)]
            wide_n = col_w.count(0)
            if wide_n > 0:
                wide_w = (total_w - sum(col_w)) // wide_n
                col_w = [w if w > 0 else wide_w for w in col_w]
            # 关键：设置 tblGrid 的 gridCol（fixed布局下真正决定列宽的是 gridCol）
            tbl_el = tbl._element
            # 移除旧 grid
            old_grid = tbl_el.find(qn('w:tblGrid'))
            if old_grid is not None:
                tbl_el.remove(old_grid)
            # 构造新 grid，插到 tblPr 之后
            new_grid = OxmlElement('w:tblGrid')
            for w in col_w:
                gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); new_grid.append(gc)
            tblPr_el.addnext(new_grid)
            for ci, w in enumerate(col_w):
                tbl.columns[ci].width = Twips(w)
            # 填充单元格
            for ri, row in enumerate(rows):
                for ci in range(ncols):
                    cell_text = row[ci] if ci < len(row) else ''
                    cell = tbl.rows[ri].cells[ci]
                    cell.width = Twips(col_w[ci])
                    tcPr = cell._tc.get_or_add_tcPr()
                    vA = OxmlElement('w:vAlign'); vA.set(qn('w:val'), 'center'); tcPr.append(vA)
                    cell.text = cell_text.replace('**','').replace('`','')
                    for para in cell.paragraphs:
                        # 清除首行缩进（Normal样式带2字符缩进，表格内不要）
                        pPr = para._p.get_or_add_pPr()
                        ind = pPr.find(qn('w:ind'))
                        if ind is None:
                            ind = OxmlElement('w:ind')
                            pPr.append(ind)
                        ind.set(qn('w:firstLine'), '0')
                        ind.set(qn('w:firstLineChars'), '0')
                        ind.set(qn('w:left'), '0')
                        ind.set(qn('w:leftChars'), '0')
                        para.paragraph_format.space_before = Pt(0)
                        para.paragraph_format.space_after = Pt(0)
                        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if ri == 0:
                    trPr = tbl.rows[ri]._tr.get_or_add_trPr()
                    th = OxmlElement('w:tblHeader'); trPr.append(th)
                    for ci in range(ncols):
                        for para in tbl.rows[0].cells[ci].paragraphs:
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for r in para.runs:
                                r.bold = True

    # 可选：去掉Heading 1/2/3自动编号（用于总结报告等文字自带中文编号、而模板自动编号不每章重置的情况）
    if strip_heading_num:
        for s in doc.styles:
            try:
                if s.name in ('Heading 1','heading 1','Heading 2','heading 2','Heading 3','heading 3'):
                    pPr = s.element.find(qn('w:pPr'))
                    if pPr is not None:
                        numPr = pPr.find(qn('w:numPr'))
                        if numPr is not None:
                            pPr.remove(numPr)
            except Exception:
                pass
    clear_frontmatter_field_results(doc)
    doc.save(out)
    print(f"已生成: {out}")

if __name__ == '__main__':
    strip_num = len(sys.argv) > 7 and sys.argv[7] == 'strip_num'
    build(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6].split('|'), strip_num)
