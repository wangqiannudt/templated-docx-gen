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

- **build_docx.py**：基于模板生成 docx。保留封面/变更记录/节结构；正文用模板 Heading 1/2/3；表题注/图题注用模板样式 + SEQ 域（fldChar + dirty，打开自动编号）；表格几何控制（tblGrid 列宽、fixed 布局、tblW 版心宽度、表头跨页重复、单元格垂直居中）；图片按 **md 文件目录**解析 `assets/` 相对路径嵌入。CLI 命名参数 + manifest。
- **analyze_template.py**：探测模板结构（封面/正文/题注样式名、变更记录表、版心宽）→ 生成 **manifest 草稿 YAML**，供确认后存盘复用。换新模板时先跑它。
- **enable_toc_update.py**：给 docx 写 `updateFields=true`，Word 打开时自动更新正文目录 + 表目录 + 图目录 + SEQ。

## 三层配置

| 层 | 谁不同 | 承载 | 例子 |
|---|---|---|---|
| 模板级 | 换模板才变 | `templates/<模板名>.manifest.yaml` | `cover_title_style: 封皮标题`、`strip_heading_num` |
| 文档级 | 每文档不同 | CLI 参数 | `--cover-title`、`--change-log` |
| 可机读 | 模板自带 | 自动探测 | 版心宽(sectPr)、变更记录表实际列数 |

## 辅助脚本（读素材 / 分析，非必需）

- analyze_docx.py / compare_headings.py / check_numbering.py：分析模板结构、对比标题、查自动编号
- check_pdf.py / pdf_extract.py / render_pdf.py / ocr_pdf.py：PDF 文本提取 / 渲染 / OCR
- extract_pptx.py：提取 PPT 文本

## 用法

工作目录切到 md + assets 所在目录（图片相对 md 目录解析）。

**1. 首次用某模板：探测生成 manifest（确认后存盘复用）**

```bash
cd ~/dev/项目材料docx工具
docenv/bin/python analyze_template.py "<模板.docx>"                       # 打印草稿
# 核对草稿（封面/题注/strip_heading_num）无误后存盘；有误则手编 yaml，勿用 -o 固化：
docenv/bin/python analyze_template.py "<模板.docx>" -o templates/<模板名>.manifest.yaml
```

草稿必查：`cover_title_style`（TOC 在前的模板会误判成"目录标题"，无封皮段改 null）；`caption_styles`（应干净无"字符"后缀）；`strip_heading_num`（用 `check_numbering.py <模板>` 看 numId，无 numId 或不每章重置→true）。

**2. 生成 docx（CLI 文档级参数 + manifest 模板级配置）**

```bash
cd ~/Desktop/项目材料各类参考
PY=~/dev/项目材料docx工具/docenv/bin/python

# 试验类（v0.0 模板，已有 manifest）
$PY ~/dev/项目材料docx工具/build_docx.py \
  可用素材/课题交付物-拟制/试验方案-v0.0.docx \
  交付定稿/试验方案.md \
  交付定稿/试验方案.docx \
  --manifest ~/dev/项目材料docx工具/templates/试验方案-v0.0.manifest.yaml \
  --cover-title "试验方案" --cover-date "2025年12月" \
  --change-log "V1.0|2025.12|编制单位|增加|结合仿真与数实结合验证成果形成正式版本"

# 研究总结报告（总结模板，封面 null + strip_heading_num 在 manifest 内为 true）
$PY ~/dev/项目材料docx工具/build_docx.py \
  可用素材/研究总结报告模版/研究总结报告（大纲目录）.docx \
  交付定稿/研究总结报告.md \
  交付定稿/研究总结报告.docx \
  --manifest ~/dev/项目材料docx工具/templates/研究总结报告.manifest.yaml \
  --change-log "V1.0|2025.12|编制单位|编制|结合课题研究成果与验收要求编制"
```

CLI 参数：`--cover-title/--cover-date`（封面 null 模板可不传）、`--change-log`（多次=多行，列用 `|`）、`--image-width-cm`（默认13）、`--narrow-keys`（逗号分隔）、`--strip-heading-num`（CLI 覆盖 manifest）、`--table-total-width`（覆盖自动探测版心宽）。

**3. 加目录自动更新标记**

```bash
$PY ~/dev/项目材料docx工具/enable_toc_update.py 交付定稿/*.docx
```

打开 docx 时 Word（不要用 WPS/LibreOffice）会弹"是否更新域"→ 点"是"，正文目录 / 表目录 / 图目录 / SEQ 编号全部自动刷新。

## 校验

```bash
cd ~/dev/项目材料docx工具 && docenv/bin/python -m pytest -q    # 44 测试，纯函数+端到端冲烟
```

## 输出逻辑

**md（文字）+ assets/（图片）→ docx**

- md 里图片用 `assets/xxx.png` 相对路径
- build_docx.py 按 md 文件所在目录解析图片、嵌入 docx
- 自包含、可移植：交付定稿文件夹（md + assets + docx）整体拷走即可
