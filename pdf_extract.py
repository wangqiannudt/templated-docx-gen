import pdfplumber, sys
pdf_path = sys.argv[1]
out_path = sys.argv[2]
with pdfplumber.open(pdf_path) as pdf:
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(f"=== 总页数: {len(pdf.pages)} ===\n\n")
        for i, page in enumerate(pdf.pages):
            f.write(f"\n========== 第 {i+1} 页 ==========\n")
            text = page.extract_text() or ""
            f.write(text)
            f.write("\n")
print(f"提取完成: {out_path}, 共 {len(pdf.pages)} 页")
