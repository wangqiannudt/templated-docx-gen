import fitz, sys
pdf_path, start, end, out_prefix = sys.argv[1], int(sys.argv[2]), int(sys.argv[3]), sys.argv[4]
doc = fitz.open(pdf_path)
print(f"总页数: {len(doc)}")
for i in range(start-1, min(end, len(doc))):
    page = doc[i]
    pix = page.get_pixmap(dpi=150)
    pix.save(f"{out_prefix}_p{i+1}.png")
    print(f"渲染第{i+1}页 -> {out_prefix}_p{i+1}.png")
