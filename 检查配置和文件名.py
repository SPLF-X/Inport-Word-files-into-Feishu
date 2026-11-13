# -*- coding: utf-8 -*-
"""离线核对 config.json 的 app_token / table_id 格式，不接触飞书 API，也不写入。"""
from pathlib import Path
from common import require_config, filename_code


def main():
    root = Path(__file__).resolve().parent
    config_file = root / "config.json"
    print("本次仅检查本地配置与文件名规则，不访问飞书、不修改任何表格。")
    config = require_config(config_file)
    print("\n多维表格 app_token：", config["app_token"])
    for entry in config["tables"]:
        print(f"子表 {entry['name']} -> table_id：{entry['table_id']}")
    print("\n文件名提取示例：")
    for filename in ("C001_候选人姓名.docx", "C062_张三.docx", "C125.docx"):
        print(f"  {filename} -> 候选编号 {filename_code(filename)}（姓名不用于匹配）")
    print("\n本地配置格式检查通过。注意：这不代表已验证真实飞书表格的访问权限或子表 ID。")


if __name__ == "__main__":
    main()
