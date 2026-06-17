# templated-docx-gen

> Generate specification‑compliant Chinese archiving **.docx** from a **Word template + Markdown** — preserves cover page / change‑log table / section structure / custom styles, with SEQ auto‑numbered captions, multi‑level TOC fields, and controlled table geometry. Also ships as a **Claude Code plugin**.

从 **Word 模板 + Markdown 正文** 生成可直接归档的规范 docx：保留模板的封面、变更记录表、节结构、自定义样式，正文走模板 Heading 样式，图/表题注用 SEQ 自动编号，表格做几何控制，目录用域自动更新。**给任意模板都能用**——已有配置直接生成，没有就探测结构、生成配置草稿、确认后复用。

面向中国相关/科研院所归档材料（试验方案、试验大纲、试验报告、研究/技术总结报告、验证报告等）。

---

## 特点

- **基于模板**：`python-docx` 加载模板，保留封面/变更记录/节结构/自定义样式，只填正文。
- **Markdown 驱动**：`#`/`##`/`###` → Heading 1/2/3；`![题注](assets/x.png)` 按相对路径嵌入。
- **SEQ 自动题注**：图/表题注用 `fldChar` 域 + `dirty`，Word 打开自动编号，可被图表目录 `TOC \c` 收集。
- **表格几何控制**：`tblGrid` 列宽 + `fixed` 布局 + `tblW` 版心宽 + 跨页表头重复 + 单元格垂直居中。
- **多级目录域**：`updateFields=true`，Word 打开一键更新正文/表/图目录与 SEQ。
- **三层配置、零硬编码**：模板级 manifest / 文档级 CLI / 可机读自动探测。
- **任意模板可用**：探测 + 人工确认 + 存盘复用，换模板不改代码。
- **可作 Claude Code 插件分发**：plugin + marketplace 二合一，路径经 `${CLAUDE_PLUGIN_ROOT}` 解析。

---

## 目录

- [快速开始](#快速开始)
- [工作流：探测 → 生成 → 目录更新](#工作流探测--生成--目录更新)
- [CLI 参数](#cli-参数)
- [三层配置](#三层配置)
- [内置模板 manifest](#内置模板-manifest)
- [辅助脚本](#辅助脚本)
- [校验](#校验)
- [作为 Claude Code 插件使用](#作为-claude-code-插件使用)
- [常见问题 / 踩坑](#常见问题--踩坑)
- [许可证](#许可证)

---

## 快速开始

### 方式 A：命令行直接用

```bash
# 1. 克隆
git clone https://github.com/wangqiannudt/templated-docx-gen.git
cd templated-docx-gen

# 2. 建虚拟环境并装核心依赖（python-docx / PyYAML / pillow）
python3 -m venv docenv
source docenv/bin/activate            # Windows: docenv\Scripts\activate
pip install -r requirements.txt       # 核心
# pip install -r requirements-optional.txt   # 辅助脚本（PDF/PPT/OCR），按需

# 3. 生成 docx（已有模板 manifest 时）
docenv/bin/python build_docx.py \
  "你的模板.docx" "正文.md" "输出.docx" \
  --manifest templates/研究总结报告.manifest.yaml \
  --change-log "V1.0|2025.12|编制单位|编制|初版"

# 4. 写目录自动更新标记
docenv/bin/python enable_toc_update.py "输出.docx"
```

用 **Word**（不要用 WPS / LibreOffice）打开生成的 docx → 弹“是否更新域” → 点“是”，正文目录 / 表目录 / 图目录 / SEQ 编号全部自动刷新。

> 也可以用 `CLAUDE_PLUGIN_ROOT=. sh setup.sh` 一键建 venv + 装核心依赖（等价于上面的第 2 步）。

### 方式 B：作为 Claude Code 插件

见 [作为 Claude Code 插件使用](#作为-claude-code-插件使用)。

---

## 工作流：探测 → 生成 → 目录更新

### 1. 首次用某模板：探测生成 manifest（确认后存盘复用）

```bash
docenv/bin/python analyze_template.py "你的模板.docx"                       # 打印草稿到终端
# 草稿核对无误后存盘；若有误判则手编 yaml，不要用 -o 一键固化错误：
docenv/bin/python analyze_template.py "你的模板.docx" -o templates/<模板名>.manifest.yaml
```

**草稿必查**（探测是启发式，可能误判）：

- `cover_title_style`：TOC 在前的模板会被误判成“目录标题”——按实际封面样式名改；模板无封皮段则改 `null`。
- `caption_styles`：应是干净的 `表题注`/`图题注`（带“字符”后缀是字符样式误匹配，已修，仍留心）。
- `heading_styles`：探测输出 `Heading 1/2/3`；若模板用中文名（如 `标题 1` / `GF一级标题`），手编改对。
- `strip_heading_num`：用 `check_numbering.py "你的模板.docx"` 看 Heading 的 numId。**无 numId**（编号写在文字里如“一、/（一）”）→ `true`；**有 numId 且不每章重置**（会累加）→ `true`；**有 numId 且每章重置** → `false`。
- `compute_heading_num`：模板 Heading 自动编号在 Word 里若 H2/H3 不按上级重置，设 `true`——build 自己算阿拉伯多级编号（1 / 1.1 / 1.1.1）拼进标题文字，确定性可靠。详见 [踩坑](#常见问题--踩坑)。

### 2. 生成 docx（CLI 文档级参数 + manifest 模板级配置）

工作目录切到 **md + assets 所在目录**（图片按 md 文件目录解析）：

```bash
cd <你的工作目录>     # 存放 正文.md 与 assets/

<repo>/docenv/bin/python <repo>/build_docx.py \
  "模板.docx" "正文.md" "输出.docx" \
  --manifest <repo>/templates/<模板名>.manifest.yaml \
  --cover-title "文档标题" --cover-date "2025年12月" \
  --change-log "V1.0|2025.12|编制单位|增加|说明" \
  --change-log "V2.0|2026.03|编制单位|修订|补充"   # 多行 = 多次 --change-log
```

`<repo>` = 本仓库克隆后的根目录路径。封面样式为 `null` 的模板（如大纲目录版）可不传 `--cover-title`/`--cover-date`。

### 3. 写目录自动更新标记

```bash
<repo>/docenv/bin/python <repo>/enable_toc_update.py "输出.docx"
```

---

## CLI 参数

`build_docx.py <模板.docx> <正文.md> <输出.docx> [选项]`

| 参数 | 说明 | 默认 |
|---|---|---|
| `--manifest` | 模板级配置 YAML（可省，省则探测/兜底） | 无 |
| `--cover-title` / `--cover-date` | 封面标题/日期（模板有封面样式时填；`null` 封面模板可不传） | None |
| `--change-log` | 变更记录行，`action=append` 可重复；列用 `\|` 分隔 | 无（不传则不动变更记录表） |
| `--image-width-cm` | 图片宽度（cm） | 13 |
| `--narrow-keys` | 窄列关键词（逗号分隔） | 序号,数量,时间,阶段,年度,是否,编号,级别,日期,方式,用途,类型 |
| `--strip-heading-num` | 去 Heading 自动编号，用 md 自带文字编号（CLI 覆盖 manifest） | manifest 值 |
| `--compute-heading-num` | strip 自动编号 + build 自算阿拉伯多级编号 (1/1.1/1.1.1) 拼进标题（**模板自动编号不可靠时用**） | manifest 值 |
| `--table-total-width` | 表格总宽（dxa，覆盖自动探测版心宽） | 自动探测 |

---

## 三层配置

| 层 | 谁不同 | 承载 | 例子 |
|---|---|---|---|
| 模板级 | 换模板才变 | `templates/<模板名>.manifest.yaml` | `cover_title_style`、`strip_heading_num` |
| 文档级 | 每文档不同 | CLI 参数 | `--cover-title`、`--change-log` |
| 可机读 | 模板自带 | 自动探测 | 版心宽（sectPr）、变更记录表实际列数 |

---

## 内置模板 manifest

| 文件 | 适用 | 要点 |
|---|---|---|
| `templates/试验方案-v0.0.manifest.yaml` | 试验类（有封面） | 封皮标题/单位 + Heading1-3 + 表/图题注，`strip=false`，`compute_heading_num=true` |
| `templates/研究总结报告.manifest.yaml` | 总结类（大纲目录版） | `cover=null`（无封皮段）+ `strip=true`（Heading 无 numId，用 md 文字编号） |
| `templates/技术总结报告（大纲目录）.manifest.yaml` | 技术总结（GF 样式） | GF 一/二/三级标题 + GF 表/图题，`compute_heading_num=true` |

换新模板时，跑 `analyze_template.py` 探测 + 确认后新增一份 manifest 即可。

---

## 辅助脚本

读素材 / 分析用，非生成必需：

- `analyze_template.py` — 探测模板结构 → manifest 草稿（换模板第一步）
- `analyze_docx.py` / `compare_headings.py` / `check_numbering.py` — 分析模板结构、对比标题、查自动编号 numId
- `enable_toc_update.py` — 写 `updateFields=true`
- `check_pdf.py` / `pdf_extract.py` / `render_pdf.py` / `ocr_pdf.py` — PDF 文本提取 / 渲染 / OCR（依赖 `requirements-optional.txt`）
- `extract_pptx.py` — 提取 PPT 文本

OCR 扫描件 PDF 的系统依赖：`brew install tesseract tesseract-lang`、`brew install poppler`。

---

## 校验

```bash
docenv/bin/python -m pytest -q    # 单元 + 端到端冲烟，全绿 = 工具健康
```

---

## 作为 Claude Code 插件使用

本仓库同时是一个 Claude Code 插件（plugin + marketplace 二合一）。技能 `templated-docx-gen` 随插件分发，脚本路径通过 `${CLAUDE_PLUGIN_ROOT}` 解析，不依赖写死的本机路径。

```bash
# 1. 添加本仓库为插件市场（本地路径或远程 git URL）
/plugin marketplace add wangqiannudt/templated-docx-gen
#   本地开发：/plugin marketplace add /path/to/templated-docx-gen

# 2. 安装插件
/plugin install templated-docx-gen@docx-template-tools

# 3. 首次使用前准备 Python 环境（在插件目录建 venv + 装核心依赖）
#    （Claude Code 会在 skill 子进程注入 CLAUDE_PLUGIN_ROOT）
CLAUDE_PLUGIN_ROOT=<插件目录> sh setup.sh
```

或用非交互 CLI：

```bash
claude plugin marketplace add wangqiannudt/templated-docx-gen
claude plugin install templated-docx-gen@docx-template-tools
```

技能触发后：venv = `${CLAUDE_PLUGIN_ROOT}/docenv/bin/python`，模板 manifest 在 `${CLAUDE_PLUGIN_ROOT}/templates/`。完整用法见 [`skills/templated-docx-gen/SKILL.md`](skills/templated-docx-gen/SKILL.md)。

> **关于 `CLAUDE_PLUGIN_ROOT`**：这个变量只在 skill 运行的子进程里由 Claude Code 注入；在普通终端 `echo $CLAUDE_PLUGIN_ROOT` 必然为空——这是预期行为，不是故障。手动跑 `setup.sh` 时请显式 `export CLAUDE_PLUGIN_ROOT=<插件目录>`，或让 skill 自己跑。

---

## 常见问题 / 踩坑

- **插件 venv 损坏（`bin/python` 不存在 / exit 127）**：从本地目录安装插件时，venv 会被整目录复制进缓存，但 `bin/python` 等符号链接可能丢失 → 半残。`setup.sh` 已内置自愈：检测到 `bin/python` 不可执行会自动 `rm -rf` + 重建。重装/升级后若异常，重跑一次 `CLAUDE_PLUGIN_ROOT=<插件目录> sh setup.sh` 即可。
- **模板 Heading 自动编号在 Word 里不可靠**：即便 `numbering.xml` 的 `lvlRestart` 写对，Word 仍可能让 H2/H3 全局累加（实测出 3.4 / 5.6 / 5.10）。**不要赌 Word 引擎**——在 manifest 开 `compute_heading_num: true`，build 自己算阿拉伯多级编号拼进标题文字。生成后务必在 Word 里验收 H2/H3 重置。
- **不要用 LibreOffice `--convert-to` 固化域值**：会破坏 SEQ `fldChar` 结构，导致 Word 表/图目录收集失败。目录更新只用 `updateFields` + Word 打开。
- **图片路径相对 md 文件目录**解析，不是当前工作目录。
- **变更记录表按实际列数对齐**：模板几列就填几列，多余留空；不传 `--change-log` 则保留模板原表不动。

---

## 许可证

[MIT](LICENSE) © 2026 用户
