---
name: templated-docx-gen
description: Use when generating formal Chinese archiving/规范 docx (试验方案, 试验大纲, 试验报告, 研究总结报告, 归档材料, 验证报告) from a Word template + markdown; when a docx needs cover page / change-log table / multi-level TOC fields / SEQ auto-numbered figure & table captions / controlled table column widths; or when handed a NEW Word template (e.g. 技术总结报告模板) whose structure must be probed before generating. Covers 规范归档/归档/规范格式文档.
---

# 模板化归档 docx 生成

## Overview

从**项目 Word 模板 + markdown 正文**生成符合规范的归档 docx，保留模板的封面/变更记录/节结构/自定义样式，正文用模板 Heading 样式，图表用 SEQ 自动编号题注，表格做几何控制，目录用域自动更新。**给任意模板都能用**：已有 manifest 直接用，没有就探测+用户确认后存盘复用。

工具位置：`${CLAUDE_PLUGIN_ROOT}/`，venv：`${CLAUDE_PLUGIN_ROOT}/docenv/bin/python`（首次用先跑 `${CLAUDE_PLUGIN_ROOT}/setup.sh` 建环境）。

## When to Use

- 用户给一个 Word 模板 + markdown，要生成规范归档 docx。
- 需要 SEQ 自动编号的图/表题注、多级目录域、变更记录表、封面填入。
- 换了新模板（结构/样式名不同），要先探测再生成。

**不要用于**：从零创建无模板的 docx（用 document-skills:docx）；改 PDF/Excel/PPT。

## 三层配置

| 层 | 谁不同 | 承载 | 例子 |
|---|---|---|---|
| 模板级 | 换模板才变 | `templates/<模板名>.manifest.yaml`（封面/正文/题注样式名、strip_num） | `cover_title_style: 封皮标题` |
| 文档级 | 每文档不同 | CLI 参数 | `--cover-title`、`--change-log` |
| 可机读 | 模板自带 | 自动探测 | 版心宽(sectPr)、变更记录表实际列数 |

## 工作流（给任意模板就能用）

```
用户给：模板路径 + md + 文档参数
  │
  ├─ templates/<模板名>.manifest.yaml 存在？
  │    ├─ 是 → 直接用它（要求1：已有模板直接用）
  │    └─ 否 → 跑探测 → manifest 草稿 → 用户确认 → 存盘（要求2：新模板检测确认后生成）
  └─ build_docx → enable_toc_update → 提示用户 Word 打开点"是"更新域
```

**新模板探测**（manifest 不存在时）：

```bash
cd ${CLAUDE_PLUGIN_ROOT}
${CLAUDE_PLUGIN_ROOT}/docenv/bin/python ${CLAUDE_PLUGIN_ROOT}/analyze_template.py "<模板.docx>"          # 打印草稿给用户看
```

草稿**经用户核对无误**后才存盘。**草稿若有误判（见下），手编 yaml 落盘——不要用 `-o` 一键存盘**，否则会把错误固化：
```bash
# 仅当草稿完全正确时才用 -o：
${CLAUDE_PLUGIN_ROOT}/docenv/bin/python ${CLAUDE_PLUGIN_ROOT}/analyze_template.py "<模板.docx>" -o templates/<模板名>.manifest.yaml
# 草稿有误 → 手编 templates/<模板名>.manifest.yaml 后保存
```

**用户确认 manifest 草稿时必查**（探测是启发式，可能误判）：
- `cover_title_style`：TOC 在前的模板会被误判成"目录标题"——按模板实际封面样式名改；模板无封皮段则改 `null`。
- `caption_styles`：应是干净的 `表题注`/`图题注`（带"字符"后缀是字符样式误匹配，已修，仍留心）。
- `heading_styles`：探测输出 `Heading 1/2/3`；若模板用中文名 `标题 1/2/3`，手编改对。
- `strip_heading_num`：用 `${CLAUDE_PLUGIN_ROOT}/docenv/bin/python ${CLAUDE_PLUGIN_ROOT}/check_numbering.py "<模板.docx>"` 看 Heading 样式的 numId。**无 numId**（编号写在正文文字里，如"一、/（一）"）→ `true`（用 md 自带文字编号，strip 无害更稳）；**有 numId 且不每章重置**（会累加成"（五）（六）"）→ `true`；**有 numId 且每章重置** → `false`。
- `compute_heading_num`（**重要，见下"标题编号"陷阱**）：模板 Heading 自动编号在 Word 里若 H2/H3 不按上级重置（实测会全局累加成 3.4/5.6），就设 `true`——build 自己算阿拉伯多级编号拼进文字，确定性可靠。**优先级最高**：开了它就不用纠结 strip_heading_num。生成后务必让用户在 Word 里验收 H2/H3 是否正确重置。
- `cover_*_style: null` = 显式禁用该字段填入（manifest 写 null 与不写 key 语义不同）。

## 生成 + 目录更新

```bash
cd ~/Desktop/<md与assets所在目录>     # 图片相对md目录解析，须在md所在目录跑

${CLAUDE_PLUGIN_ROOT}/docenv/bin/python ${CLAUDE_PLUGIN_ROOT}/build_docx.py \
  "<模板.docx>" "<正文.md>" "<输出.docx>" \
  --manifest ${CLAUDE_PLUGIN_ROOT}/templates/<模板名>.manifest.yaml \
  --cover-title "标题" --cover-date "2025年12月" \
  --change-log "V1.0|2025.12|编制单位|增加|说明" \
  --change-log "V2.0|2026.03|编制单位|修订|补充"   # 多行=多次 --change-log
  # 封面 null 的模板（如研究总结报告大纲目录版）可不传 --cover-title/--cover-date
  # [--strip-heading-num]  覆盖 manifest 的 strip（manifest 已配则不用传）
  # [--image-width-cm 13]  [--narrow-keys 序号,数量,时间]  [--table-total-width 8505]

${CLAUDE_PLUGIN_ROOT}/docenv/bin/python ${CLAUDE_PLUGIN_ROOT}/enable_toc_update.py "<输出.docx>"
```

生成后告诉用户：**用 Word（不是 WPS/LibreOffice）打开 → 弹"是否更新域"→ 点"是"** → 正文目录/表目录/图目录/SEQ 编号全部刷新。

交付时可直接粘贴给用户：
```
docx 已生成：<输出.docx>
请用 Word 打开 → 弹"是否更新域"选"是"，目录/图表题注编号会自动刷新。
```

## CLI 参数速查

| 参数 | 说明 | 默认 |
|---|---|---|
| `--manifest` | 模板级配置 YAML（可省，省则探测/兜底） | 无 |
| `--cover-title` / `--cover-date` | 封面标题/日期（模板有封面样式时填；null 封面模板可不传） | None |
| `--change-log` | 变更记录行，`action=append` 可重复；列用 `\|` 分隔 | 无（不传则不动变更记录表） |
| `--image-width-cm` | 图片宽度 cm | 13 |
| `--narrow-keys` | 窄列关键词（逗号分隔） | 序号,数量,时间,阶段,年度,是否,编号,级别,日期,方式,用途,类型 |
| `--strip-heading-num` | 去 Heading 自动编号（CLI 覆盖 manifest） | manifest 值 |
| `--compute-heading-num` | strip 自动编号 + build 自己算阿拉伯多级编号(1/1.1/1.1.1)拼进标题文字（**模板自动编号不可靠时用**） | manifest 值 |
| `--table-total-width` | 表格总宽 dxa（覆盖自动探测版心宽） | 自动探测 |

## 方法论（为何这么做）

1. **基于模板**：python-docx 加载模板，保留封面/变更记录/节结构/自定义样式，只填内容。
2. **markdown 驱动**：`#`→Heading 1/2/3；`![caption](assets/x.png)` 按 **md 文件目录**解析嵌入。
3. **SEQ 题注**：`fldChar(begin/instr/separate/result/end)` + `dirty="true"`，Word 打开随 updateFields 自动编号，被表/图目录 TOC `\c` 收集。
4. **表格几何**：`tblGrid gridCol` 列宽 + `fixed` 布局 + `tblW` 版心宽 + `tblHeader` 跨页重复 + `vAlign` 垂直居中；窄列定宽其余均分。
5. **目录更新**：`enable_toc_update.py` 写 `updateFields=true`，Word 打开更新。

## 关键陷阱（踩过的坑）

- **❌ 绝不用 LibreOffice `--convert-to` 固化域值**：会破坏 SEQ fldChar 结构，导致 Word 表/图目录收集失败。目录更新只用 `updateFields` + Word 打开。
- **⚠️ 模板 Heading 自动编号在 Word 里不可靠**：即便 numbering.xml 的 `lvlRestart` 写对了、也补全了，Word 仍可能让 H2/H3 计数器**全局累加不按上级重置**（实测 V0.0 模板出 3.4/5.6/5.10）。靠模板自动编号 = 赌 Word 引擎，**不要赌**。凡模板带 H2/H3 自动编号的，直接在 manifest 开 `compute_heading_num: true`：build 自己算 1/1.1/1.1.1 拼进标题文字，确定性可靠（md 标题文字本身不带阿拉伯编号时用这个；若 md 已带"一、/（一）"文字编号则用 `strip_heading_num`）。**生成后必须让用户在 Word 验收 H2/H3 重置**。
- **图片路径相对 md 文件目录**，不是工作目录。
- **变更记录表按实际列数对齐**：模板几列就填几列，多余留空；不传 `--change-log` 则保留模板原变更记录不动。
- **`strip_heading_num` 场景**：研究总结报告类模板 Heading 自动编号不每章重置 → 必须 strip，用 md 标题自带的文字编号（一、/（一））；试验类模板每章重置 → 不 strip。
- **版心宽自动探测**：从 sectPr 读 `pgSz.w − pgMar(left+right)`，不再写死；删正文区时已保留末尾 sectPr。

## 校验

```bash
cd ${CLAUDE_PLUGIN_ROOT} && ${CLAUDE_PLUGIN_ROOT}/docenv/bin/python -m pytest -q    # 全绿=工具健康
```

生成后用 python-docx 抽查：SEQ+dirty 在、tblGrid 列宽和≈版心宽、封面/变更记录已填、图片已嵌入。
