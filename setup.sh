#!/usr/bin/env sh
# templated-docx-gen 插件环境准备：在插件根目录建 venv 并装核心依赖。
# venv 建在 ${CLAUDE_PLUGIN_ROOT}/docenv：
#   - 本地 marketplace add（文件复制）会把已有 docenv 一起带进 cache，但 bin/python 等
#     符号链接可能在复制时丢失 → venv 半残；本脚本会检测并自动重建。
#   - git 远程分发时，clone 下来无 docenv，由本脚本现场建。
# 可选依赖（PDF/PPT/OCR 辅助脚本）见 requirements-optional.txt，按需安装。
# 系统依赖（OCR 扫描件 PDF）：brew install tesseract tesseract-lang poppler

set -e

: "${CLAUDE_PLUGIN_ROOT:?未设置 CLAUDE_PLUGIN_ROOT，请在插件上下文运行，或手动 export CLAUDE_PLUGIN_ROOT=<插件根目录>}"
ROOT="${CLAUDE_PLUGIN_ROOT}"
VENV="${ROOT}/docenv"

# 建/复用/重建 venv：venv 目录在但 bin/python 跑不起来（本地复制常见：symlink 丢失）
# 视为损坏，删后重建；否则复用；目录不在则新建。
if [ -d "$VENV" ]; then
  if "$VENV/bin/python" -c 'import sys' >/dev/null 2>&1; then
    echo "→ venv 已存在且可用: $VENV（跳过创建）"
  else
    echo "→ venv 已存在但损坏（bin/python 不可执行，常见于本地复制丢失 symlink），重建: $VENV"
    rm -rf "$VENV"
    python3 -m venv "$VENV"
  fi
else
  echo "→ 建虚拟环境: $VENV"
  python3 -m venv "$VENV"
fi

echo "→ 装核心依赖"
"$VENV/bin/pip" install --upgrade pip
"$VENV/bin/pip" install -r "$ROOT/requirements.txt"

echo "→ 验证核心导入"
"$VENV/bin/python" -c "import docx, yaml; print('OK: python-docx / PyYAML 就绪')"

echo "完成。辅助脚本（PDF/PPT/OCR）需要时执行："
echo "  $VENV/bin/pip install -r $ROOT/requirements-optional.txt"
