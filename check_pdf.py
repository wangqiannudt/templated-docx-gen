import pdfplumber, sys
for pdf_path in sys.argv[1:]:
    with pdfplumber.open(pdf_path) as pdf:
        pages = len(pdf.pages)
        # 抽样首页和中间页的文本量
        sample_pages = [0, pages//2, pages-1] if pages > 2 else [0]
        total_chars = 0
        for idx in sample_pages:
            t = pdf.pages[idx].extract_text() or ""
            total_chars += len(t)
        avg = total_chars / len(sample_pages)
        kind = "文本型" if avg > 100 else "扫描件/图片型"
        print(f"{pdf_path.split('/')[-1]}: {pages} 页, 抽样均字符 {avg:.0f} -> {kind}")
