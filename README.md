# 项目材料 docx 生成工具

生成「试验」课题归档材料（试验方案/试验大纲/试验报告/研究总结报告）的 docx，基于项目提供的模板（v0.0 草稿 / 研究总结报告模板），保留封面、变更记录、节结构、自定义样式，正文用模板 Heading 样式，表格做几何控制，图片走 assets 相对路径。

## 环境准备

```bash
cd ~/dev/项目材料docx工具
python3 -m venv docenv
source docenv/bin/activate            # Windows: docenv\Scripts\activate
pip install -r requirements.txt
```

系统依赖（OCR 扫描件 PDF 用）：`brew install tesseract tesseract-lang`、`brew install poppler`（PDF 渲染）。

## 核心脚本

- **build_docx.py**：基于模板生成 docx。保留封面/变更记录/节结构；正文用模板 Heading 1/2/3（标题自动编号）；表题注/图题注用模板样式 + SEQ 域（fldChar + dirty，打开自动编号）；表格几何控制（tblGrid 列宽、fixed 布局、tblW 版心宽度、表头跨页重复、单元格垂直居中、内容统一左对齐）；图片按 **md 文件目录**解析 `assets/` 相对路径嵌入。
- **enable_toc_update.py**：给 docx 写 `updateFields=true`，Word/WPS 打开时自动更新正文目录 + 表目录 + 图目录。

## 辅助脚本（读素材 / 分析，非必需）

- analyze_docx.py / compare_headings.py / check_numbering.py：分析模板结构、对比标题、查自动编号
- check_pdf.py / pdf_extract.py / render_pdf.py / ocr_pdf.py：PDF 文本提取 / 渲染 / OCR
- extract_pptx.py：提取 PPT 文本

## 用法

工作目录切到「交付定稿」（md + assets 所在目录），图片相对路径才能解析：

```bash
cd ~/Desktop/项目材料各类参考

# 试验类三件（v0.0 模板）
DOCENV=~/dev/项目材料docx工具/docenv/bin/python
TPL=可用素材/课题交付物-拟制/试验方案-v0.0.docx

$DOCENV ~/dev/项目材料docx工具/build_docx.py "$TPL" \
  交付定稿/试验方案.md \
  交付定稿/试验方案.docx \
  "试验方案" "2025年12月" \
  "V1.0|2025.12|编制单位|增加|结合仿真与数实结合验证成果形成正式版本"

# 研究总结报告（总结模板 + strip_num：去 Heading 自动编号、用文字编号）
$DOCENV ~/dev/项目材料docx工具/build_docx.py \
  可用素材/研究总结报告模版/研究总结报告（大纲目录）.docx \
  交付定稿/研究总结报告.md \
  交付定稿/研究总结报告.docx \
  "研究总结报告" "2025年12月" \
  "V1.0|2025.12|编制单位|编制|结合课题研究成果与验收要求编制" strip_num

# 加目录自动更新标记
$DOCENV ~/dev/项目材料docx工具/enable_toc_update.py \
  交付定稿/试验方案.docx \
  交付定稿/试验试验报告.docx \
  交付定稿/研究总结报告.docx
```

打开 docx 时 Word 会弹"是否更新域"→ 点"是"，正文目录 / 表目录 / 图目录 / SEQ 编号全部自动刷新。

## 输出逻辑

**md（文字）+ assets/（图片）→ docx**

- md 里图片用 `assets/xxx.png` 相对路径
- build_docx.py 按 md 文件所在目录解析图片、嵌入 docx
- 自包含、可移植：交付定稿文件夹（md + assets + docx）整体拷走即可
