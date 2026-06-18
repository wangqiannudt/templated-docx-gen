# 更新日志

所有重要变更记录于此文件。格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本号遵循[语义化版本](https://semver.org/lang/zh-CN/)。

## [1.1.0] - 2026-06-18

### 新增
- README 重写为产品介绍页：痛点钩子、价值主张、30 秒上手命令、亮点速览
- 效果图 `docs/img/hero.png`（Markdown → 规范 docx 对比）、`docs/img/features.png`（四项核心亮点）
- 顶部 shields 徽章（license / python / platform / claude-code）
- 本 CHANGELOG，开始版本化管理
- 版本同步脚本 `scripts/sync-version.py`（plugin.json 为版本唯一来源）
- `CONTRIBUTING.md`（提交规范与发版流程）

### 变更
- 模板 `试验方案-v0.0.manifest.yaml` → `质量体系-v0.0.manifest.yaml`（文件名、注释、README 模板表同步更新）
- 插件版本 `1.0.0` → `1.1.0`

## [1.0.0] - 2026-06-17

### 初始公开版本
- 模板 + Markdown → 规范归档 docx：保留封面 / 变更记录 / 节结构，SEQ 自动题注，多级目录域，表格几何控制
- 三层配置（模板级 manifest / 文档级 CLI / 可机读自动探测），任意模板可探测 + 生成
- 内置三套 manifest（质量体系、研究总结报告、技术总结报告）
- Claude Code 插件（plugin + marketplace 二合一），脚本路径经 `${CLAUDE_PLUGIN_ROOT}` 解析
- 脱敏后公开至 GitHub
