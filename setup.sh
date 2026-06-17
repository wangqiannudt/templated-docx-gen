#!/usr/bin/env sh
# templated-docx-gen 插件环境准备：在插件根目录建 venv 并装核心依赖。
# venv 建在 ${CLAUDE_PLUGIN_ROOT}/docenv：
#   - 本地 marketplace add（文件复制）会把已有 docenv 一起带进 cache，可直接用；
#   - git 远程分发时，clone 下来无 docenv，由本脚本现场建。
# 可选依赖（PDF/PPT/OCR 辅助脚本）见 requirements-optional.txt，按需安装。
# 系统依赖（OCR 扫描件 PDF）：brew install tesseract tesseract-lang poppler

set -e

: "${CLAUDE_PLUGIN_ROOT:?未设置 CLAUDE_PLUGIN_ROOT，请在插件上下文运行，或手动 export CLAUDE_PLUGIN_ROOT=<插件根目录>}"
ROOT="${CLAUDE_PLUGIN_ROOT}"
VENV="${ROOT}/docenv"

if [ -d "$VENV" ]; then
  echo "→ venv 已存在: $VENV（跳过创建；如需重建请先删除该目录）"
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
