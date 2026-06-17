import fitz, pytesseract, io, sys
from PIL import Image
pdf_path, out_path = sys.argv[1], sys.argv[2]
doc = fitz.open(pdf_path)
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(f"=== 总页数: {len(doc)} ===\n")
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=200)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        text = pytesseract.image_to_string(img, lang='chi_sim+eng')
        f.write(f"\n========== 第 {i+1} 页 ==========\n{text}\n")
print(f"OCR完成: {out_path}, 共{len(doc)}页")
