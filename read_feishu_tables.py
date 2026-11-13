"""为现有 PowerShell/CMD 菜单读取飞书真实表格名称，不读取或写入业务记录。"""
import argparse
import json
import sys
from pathlib import Path

from common import clean_app_token, require_config, resolve_feishu_tables


def main():
    parser = argparse.ArgumentParser(description="读取多维表格名称和子表清单，供原有菜单及业务脚本使用")
    source = parser.add_mutually_exclusive_group()
    source.add_argument("--config", help="已有配置文件；可只填写 app_token，tables/name 可省略")
    source.add_argument("--app-token", help="多维表格 app_token 或包含 /base/ 的完整链接")
    parser.add_argument("--output", required=True, help="本次会话配置的输出路径（不包含应用凭据）")
    args = parser.parse_args()
    if args.app_token is not None:
        config = {"app_token": clean_app_token(args.app_token), "tables": []}
    else:
        config = require_config(args.config or "config.json", allow_discovery=True)
    resolved = resolve_feishu_tables(config)
    Path(args.output).write_text(json.dumps(resolved, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"读取飞书表格清单失败：{exc}", file=sys.stderr)
        raise SystemExit(1)
