import sys, re, zipfile
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

for path in sys.argv[1:]:
    with zipfile.ZipFile(path) as z:
        doc_xml = z.read('word/document.xml').decode('utf-8')
    has_toc = re.search(r'<w:instrText[^>]*>\s*TOC', doc_xml) is not None
    # 加 updateFields
    doc = Document(path)
    settings = doc.settings.element
    existing = settings.find(qn('w:updateFields'))
    if existing is None:
        uf = OxmlElement('w:updateFields')
        uf.set(qn('w:val'), 'true')
        settings.insert(0, uf)
        doc.save(path)
        flag = '已加updateFields'
    else:
        flag = '已有updateFields'
    name = path.split('/')[-1]
    print(f'{name}: 目录TOC域={"在" if has_toc else "不在"} | {flag}')
