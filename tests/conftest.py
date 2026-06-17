"""公共 fixtures 与 helper（后续 task 扩充）。"""
import os

VENV_PY = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "docenv", "bin", "python",
)
