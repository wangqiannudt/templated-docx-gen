import re

with open('/tmp/v0unpacked/word/styles.xml','r',encoding='utf-8') as f:
    styles = f.read()
with open('/tmp/v0unpacked/word/numbering.xml','r',encoding='utf-8') as f:
    numbering = f.read()

print("=== Heading相关样式 ===")
for m in re.finditer(r'<w:style\b[^>]*>.*?</w:style>', styles, re.S):
    body = m.group(0)
    sid_m = re.search(r'w:styleId="([^"]+)"', body)
    name_m = re.search(r'<w:name w:val="([^"]+)"', body)
    if not sid_m: continue
    sid = sid_m.group(1)
    name = name_m.group(1) if name_m else ''
    if sid in ('1','2','3') or 'Heading' in name or '标题' in name:
        numPr = '<w:numPr>' in body
        numId = re.search(r'<w:numId w:val="([^"]+)"', body)
        ilvl = re.search(r'<w:ilvl w:val="([^"]+)"', body)
        print(f'styleId={sid} name="{name}" 自动编号={"有" if numPr else "无"} numId={numId.group(1) if numId else "-"} ilvl={ilvl.group(1) if ilvl else "-"}')

print("\n=== numbering.xml abstractNum 定义（前3个）===")
for m in re.finditer(r'<w:abstractNum\b[^>]*w:abstractNumId="([^"]+)"[^>]*>(.*?)</w:abstractNum>', numbering, re.S):
    aid = m.group(1)
    body = m.group(2)
    lvls = re.findall(r'<w:lvl\b[^>]*w:ilvl="(\d+)"[^>]*>.*?<w:numFmt w:val="([^"]+)"[^/]*/>\s*<w:lvlText w:val="([^"]*)"', body, re.S)
    if lvls:
        print(f'abstractNum {aid}:')
        for il, fmt, txt in lvls[:5]:
            print(f'    lvl{il}: fmt={fmt} text="{txt}"')

print("\n=== 正文标题段落（Heading样式）的pPr ===")
with open('/tmp/v0unpacked/word/document.xml','r',encoding='utf-8') as f:
    doc = f.read()
# 找含pStyle val=1/2/3的段落
for m in re.finditer(r'<w:p\b[^>]*>(.*?)</w:p>', doc, re.S):
    body = m.group(1)
    pstyle = re.search(r'<w:pStyle w:val="([^"]+)"', body)
    if pstyle and pstyle.group(1) in ('1','2','3'):
        text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', body))
        has_numPr = '<w:numPr>' in body
        print(f'pStyle={pstyle.group(1)} text="{text[:30]}" 段内numPr={"有" if has_numPr else "无"}')
