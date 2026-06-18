# 贡献指南

感谢参与。本仓库遵循以下约定。

## 提交规范

使用 [Conventional Commits](https://www.conventionalcommits.org/zh-hans/)：

    type(scope): 简述

    可选正文，说明动机与影响面。

    Co-Authored-By: Claude <noreply@anthropic.com>

| type | 用途 |
|---|---|
| `feat` | 新功能 |
| `fix` | 修复缺陷 |
| `docs` | 文档 |
| `refactor` | 重构（不改对外行为） |
| `chore` | 构建 / 工具 / 杂务 |
| `test` | 测试 |
| `perf` | 性能 |

常用 scope：`readme`、`templates`、`plugin`、`setup`、`skill`、`deps`。示例见 `git log --oneline`。AI 协作的提交末尾带 `Co-Authored-By`。

## 发版流程

版本号遵循[语义化版本](https://semver.org/lang/zh-CN/)，所有变更记录在 [CHANGELOG.md](CHANGELOG.md)。

1. 定版本号 `X.Y.Z`：不兼容变更升主版本、新增功能升次版本、修复 / 文档升修订
2. 改 `.claude-plugin/plugin.json` 的 `version`（**唯一来源**）
3. 同步：`python scripts/sync-version.py`（写入 `marketplace.json`）
4. 在 `CHANGELOG.md` 顶部加 `## [X.Y.Z] - YYYY-MM-DD` 段
5. 提交：`git commit -m "chore(plugin): 发版 vX.Y.Z"`
6. 打标签：`git tag -a vX.Y.Z -m "vX.Y.Z"`
7. 推送：`git push && git push --tags`
8. 发 Release：`gh release create vX.Y.Z --notes-file <提取出的 CHANGELOG 段>`（或网页粘贴）

## 模板与脱敏

- 新模板 manifest 先跑 `analyze_template.py` 探测 + 人工核对必查项，再存盘复用
- **严禁**把真实课题名 / 单位 / 人名 / 本机绝对路径写进仓库（历史脱敏见 `git log`）
- 示例与效果图一律使用虚构内容
