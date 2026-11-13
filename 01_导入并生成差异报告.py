# ================== table_id 获取方法（请先阅读） ==================
# 1. 在浏览器中打开飞书多维表格。
# 2. 点击左侧的「T1 mapping」子表。
# 3. 查看浏览器地址栏，找到 table= 后面的内容。
#    例如：https://xxx.feishu.cn/base/xxxxx?table=tblAbc123
#    则 T1 的 table_id 为：tblAbc123
# 4. 点击「T2 mapping」子表，按相同方法获取 T2 的 table_id。
# 5. 将 config.json 中的示例值替换为实际获取的 ID。
#
# 注意：
# - 每个子表有独立的 table_id。
# - table_id 与 App ID、App Secret、app_token 不同。
# - 不要填写子表名称，必须填写对应的真实 ID。
# - 本工具不是在 Python 代码里直接填写 ID，而是在 config.json 中填写。
# - 如果网址中没有 table=，可以通过飞书「列出数据表」API 查找。
# - Word 文件名示例 C001_候选人姓名.docx：只用 C001 对应「候选编号」匹配飞书，不使用文件名中的姓名。
# - config.json 填纯 ID；base/MNCR... 会自动提取 MNCR...，tbl...&view=... 会只保留 tbl...。
# ================================================================

"""阶段一：多文件夹 Word -> 多维表格。默认仅预览，--apply-blanks 只填空白。"""
import argparse
import json
from collections import defaultdict
from pathlib import Path
from tkinter import Tk, filedialog

from common import (
    FIELD_NAMES, FeishuClient, add_sheet, create_workbook, filename_code,
    index_records, is_blank, make_digest, normalize_for_compare,
    read_docx, require_config, selected_table_entries, text_value, timestamp,
)


def choose_tables(config, no_dialog):
    root = None
    selections = []
    used_folders = set()
    for entry in config["tables"]:
        folder = (entry.get("folder") or "").strip()
        if not folder and not no_dialog:
            if root is None:
                root = Tk()
                root.withdraw()
            folder = filedialog.askdirectory(title=f"选择要导入『{entry['name']}』的 Word 文件夹；取消可跳过")
        if not folder:
            print(f"跳过子表 {entry['name']}：没有选择文件夹")
            continue
        p = Path(folder).expanduser().resolve()
        if not p.is_dir():
            raise ValueError(f"文件夹不存在：{p}")
        if str(p).lower() in used_folders:
            raise ValueError(f"不同子表选中了同一个文件夹：{p}，为安全起见已终止")
        used_folders.add(str(p).lower())
        selections.append((entry["name"], entry["table_id"], p))
    if root:
        root.destroy()
    return selections


def scan_table(client, name, table_id, folder):
    records = client.list_records(table_id)
    idx = index_records(records)
    files = sorted(folder.iterdir())
    groups = defaultdict(list)
    exceptions = []
    detail = []
    diffs = []
    candidates = []

    for path in files:
        if not path.is_file() or path.name.startswith("~$"):
            continue
        if path.suffix.lower() == ".doc":
            exceptions.append([name, path.name, "", "旧版 .doc 不支持，请先转换为 .docx"])
            continue
        if path.suffix.lower() != ".docx":
            continue
        try:
            code = filename_code(path)
            groups[code].append(path)
        except Exception as exc:
            exceptions.append([name, path.name, "", str(exc)])

    for code, paths in groups.items():
        if len(paths) > 1:
            for path in paths:
                exceptions.append([name, path.name, code, "同一子表的文件夹内候选编号重复，全部跳过"])
            continue
        path = paths[0]
        matches = idx.get(code, [])
        if len(matches) != 1:
            exceptions.append([name, path.name, code, "飞书无匹配记录" if not matches else "飞书候选编号重复"])
            continue
        record = matches[0]
        record_id = record["record_id"]
        try:
            word = read_docx(path)
            previous = {f: text_value((record.get("fields") or {}).get(f)) for f in FIELD_NAMES}
        except Exception as exc:
            exceptions.append([name, path.name, code, str(exc)])
            continue

        blanks = {}
        for field in FIELD_NAMES:
            old, new = previous[field], word[field]
            if is_blank(old):
                blanks[field] = new
                status = "预览：将填写空白"
            elif old == new:
                status = "一致：无需修改"
            else:
                status = "待人工审核"
                kind = "仅空白字符/排版差异" if normalize_for_compare(old) == normalize_for_compare(new) else "文字内容有差异"
                digest = make_digest(client.app_token, table_id, record_id, code, field, old, new)
                # 审批状态、子表、table_id、record_id、编号、字段、旧值、新值、差异类型、文件名、校验标识、备注
                diffs.append(["待确认", name, table_id, record_id, code, field, old, new, kind, path.name, digest, ""])
            detail.append([name, path.name, code, record_id, field, old, new, status])
        if blanks:
            candidates.append({"table_name": name, "table_id": table_id, "code": code,
                               "record_id": record_id, "fields": blanks, "path": path.name})
    return candidates, diffs, detail, exceptions


def apply_blank_fills(client, candidates, detail, exceptions):
    """写前重读：若字段不再为空就跳过，不覆盖。每批 50 个不同记录。"""
    by_table = defaultdict(list)
    for c in candidates:
        by_table[(c["table_name"], c["table_id"])].append(c)
    for (table_name, table_id), rows in by_table.items():
        for start in range(0, len(rows), 50):
            batch_items = rows[start:start + 50]
            validated = []
            for item in batch_items:
                try:
                    current = client.get_record(table_id, item["record_id"])
                    current_code = text_value((current.get("fields") or {}).get("候选编号")).strip().upper()
                    if current_code != item["code"]:
                        raise ValueError("写前校验失败：候选编号已变化")
                    new_fields = {}
                    for field, new in item["fields"].items():
                        old = text_value((current.get("fields") or {}).get(field))
                        if is_blank(old):
                            new_fields[field] = new
                        else:
                            exceptions.append([table_name, item["path"], item["code"], f"写前发现『{field}』已有内容，不覆盖"])
                            mark_detail(detail, table_name, item["record_id"], field, "写前已有内容：跳过")
                    if new_fields:
                        validated.append((item, new_fields))
                except Exception as exc:
                    exceptions.append([table_name, item["path"], item["code"], f"写前校验失败：{exc}"])
                    for field in item["fields"]:
                        mark_detail(detail, table_name, item["record_id"], field, "校验失败：跳过")
            if not validated:
                continue
            payload = [{"record_id": item["record_id"], "fields": fields} for item, fields in validated]
            try:
                client.batch_update(table_id, payload)
                for item, fields in validated:
                    for field in fields:
                        mark_detail(detail, table_name, item["record_id"], field, "已填写空白")
                print(f"  {table_name} 已填写 {len(validated)} 条记录")
            except Exception as exc:
                # 批量请求失败时不自动重试，避免未知写入状态下重复操作
                for item, fields in validated:
                    exceptions.append([table_name, item["path"], item["code"], f"批量写入失败，请人工查询飞书确认：{exc}"])
                    for field in fields:
                        mark_detail(detail, table_name, item["record_id"], field, "写入失败/状态待核实")


def mark_detail(detail, table_name, record_id, field, status):
    for row in detail:
        if row[0] == table_name and row[3] == record_id and row[4] == field:
            row[-1] = status


def save_report(run_id, config, mode, diffs, detail, exceptions, selected_tables=None):
    out = Path("报告")
    out.mkdir(exist_ok=True)
    path = out / f"差异审核_{run_id}.xlsx"
    wb = create_workbook()
    add_sheet(wb, "差异审核", ["审批状态", "目标子表", "table_id", "record_id", "候选编号", "字段名",
              "飞书现值", "Word新值", "差异类型", "来源文件", "校验标识", "备注"], diffs, review=True)
    add_sheet(wb, "字段处理明细", ["目标子表", "来源文件", "候选编号", "record_id", "字段名", "飞书现值", "Word新值", "处理状态"], detail)
    add_sheet(wb, "异常与跳过", ["目标子表", "来源文件", "候选编号", "错误描述"], exceptions)
    add_sheet(wb, "使用说明", ["项目", "说明"], [
        ["本次运行", run_id],
        ["本次模式", "仅预览" if mode == "preview" else "只填写空白字段（已执行）"],
        ["app_token", config["app_token"]],
        ["本次选择子表", "、".join(x["name"] for x in selected_tables) if selected_tables else "（无）"],
        ["审核方法", "只修改『差异审核』工作表 A 列：选择『覆盖』『保留』『待确认』"],
        ["审核提示", "只有『覆盖』会被脚本二考虑；默认『待确认』不执行。"],
        ["禁止修改", "不要改动目标子表、table_id、record_id、候选编号、字段名、新旧内容、校验标识等列。"],
        ["保存方法", "Excel 保存为 .xlsx 后，再运行 02 脚本预检和执行。"],
        ["安全说明", "审核期间如有人改动飞书旧值，脚本二会校验失败并拒绝覆盖。"],
        ["数据保护", "报告含候选人邮件内容，存本机受控目录，不要公开分享。"],
    ])
    wb.save(path)
    return path


def main():
    p = argparse.ArgumentParser(description="Word -> 飞书多维表格：预览、填空和差异报告")
    p.add_argument("--config", default="config.json", help="配置文件路径")
    p.add_argument("--apply-blanks", action="store_true", help="确认只填写飞书空白字段（默认仅预览）")
    p.add_argument("--no-dialog", action="store_true", help="不弹文件夹窗口，仅使用配置中的 folder")
    p.add_argument("--table-id", action="append", default=None, help="本次操作的子表 ID；可多次传入；未传时按配置处理全部子表")
    args = p.parse_args()
    config = require_config(args.config)
    target_tables = selected_table_entries(config, args.table_id)
    if not target_tables:
        print("未选择任何子表：不读取 Word、不请求飞书、不生成写入任务。")
        return
    print("本次选中子表：" + "、".join(x["name"] for x in target_tables))
    scoped_config = dict(config, tables=target_tables)
    selections = choose_tables(scoped_config, args.no_dialog)
    if not selections:
        print("没有选择任何子表/文件夹；未进行任何写入。")
        return
    client = FeishuClient(config["app_token"])
    run_id = timestamp()
    all_candidates, all_diffs, all_details, all_exceptions = [], [], [], []
    for name, table_id, folder in selections:
        print(f"\n读取 {name} ← {folder}")
        candidates, diffs, detail, exceptions = scan_table(client, name, table_id, folder)
        all_candidates.extend(candidates)
        all_diffs.extend(diffs)
        all_details.extend(detail)
        all_exceptions.extend(exceptions)
        print(f"  预定填空 {len(candidates)} 条，差异字段 {len(diffs)} 个，异常 {len(exceptions)} 条")
    if args.apply_blanks and all_candidates:
        counts = defaultdict(int)
        for item in all_candidates:
            counts[item["table_name"]] += 1
        print("即将写入的子表与记录数（只填空白）：")
        for name, count in counts.items():
            print(f"  {name}：{count} 条记录")
        answer = input("请再次确认本次选中的子表；只填空白，不覆盖旧值。输入全部大写 YES 确认（小写 yes 不执行）：").strip()
        if answer == "YES":
            # 正式填写空白前，先在本机留下可追溯的处理计划。
            backup_dir = Path("备份")
            backup_dir.mkdir(exist_ok=True)
            plan_path = backup_dir / f"填写空白前计划_{run_id}.json"
            plan_path.write_text(json.dumps(all_candidates, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"填空前计划备份：{plan_path.resolve()}")
            apply_blank_fills(client, all_candidates, all_details, all_exceptions)
            mode = "fill"
        else:
            print("未确认写入，本次仍只生成预览。")
            mode = "preview"
    else:
        mode = "preview"
    report = save_report(run_id, config, mode, all_diffs, all_details, all_exceptions, target_tables)
    print(f"\n生成报告：{report.resolve()}")
    print(f"差异待审核：{len(all_diffs)} 项；异常/跳过：{len(all_exceptions)} 项")
    if mode == "preview":
        print("本次未修改飞书。确认预览无误后重新运行，加 --apply-blanks。")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"运行已终止，未保证本轮之前的写入会回滚：{exc}")
        raise
