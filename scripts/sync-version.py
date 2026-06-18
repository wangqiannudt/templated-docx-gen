#!/usr/bin/env python3
"""版本单一来源同步：以 plugin.json 的 version 为准，写入 marketplace.json 的 plugin 条目。

用法：python scripts/sync-version.py
幂等：版本已一致则不改文件。

设计：plugin.json 是版本唯一来源（source of truth），marketplace.json 的 plugin.version
由本脚本同步，避免两处手改不同步。详见 CONTRIBUTING.md「发版流程」。
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLUGIN = Path(ROOT / ".claude-plugin" / "plugin.json")
MARKET = Path(ROOT / ".claude-plugin" / "marketplace.json")


def main():
    plugin = json.loads(PLUGIN.read_text(encoding="utf-8"))
    version = plugin.get("version")
    if not version:
        sys.exit("错误：plugin.json 没有 version 字段")

    market = json.loads(MARKET.read_text(encoding="utf-8"))
    changed = False
    for p in market.get("plugins", []):
        if p.get("name") == plugin["name"] and p.get("version") != version:
            p["version"] = version
            changed = True

    if changed:
        MARKET.write_text(
            json.dumps(market, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print(f"已同步 marketplace.json → version {version}")
    else:
        print(f"版本已一致（{version}），无需改动")


if __name__ == "__main__":
    main()
