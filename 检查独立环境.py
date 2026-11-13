# -*- coding: utf-8 -*-
"""检查项目专属 .venv 的依赖。由安装脚本和启动菜单调用。"""
import importlib
import sys
from pathlib import Path

# 依赖名称与导入名称不同的项目：python-docx -> docx
DEPENDENCIES = ("requests", "docx", "openpyxl")


def main():
    print("当前 Python：", sys.executable, flush=True)
    failures = []
    for name in DEPENDENCIES:
        try:
            importlib.import_module(name)
            print("  [OK]", name, flush=True)
        except Exception as exc:
            failures.append(name)
            print(f"  [FAIL] {name}: {type(exc).__name__}: {exc}", flush=True)
    if failures:
        print("依赖检查失败。请使用本项目的 .venv 重新安装缺少的依赖。", flush=True)
        return 1
    print("依赖检查通过：独立环境已具备所需的 Python 包。", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
