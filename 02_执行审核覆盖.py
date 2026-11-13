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

"""阶段二：人工把 Excel 的『审批状态』改为『覆盖』后，重新比对飞书原值，再选择性覆盖。"""
import argparse
import json
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook
from common import (
    FIELD_NAMES, FeishuClient, add_sheet, create_workbook, make_digest,
    require_config, selected_table_entries, text_value, timestamp,
)

COLS = ["审批状态", "目标子表", "table_id", "record_id", "候选编号", "字段名",
        "飞书现值", "Word新值", "差异类型", "来源文件", "校验标识", "备注"]


def read_approvals(report, config, allowed_table_ids=None):
    wb = load_workbook(report, read_only=True, data_only=True)
    if "差异审核" not in wb or "使用说明" not in wb:
        raise ValueError("不是脚本一生成的审核报告（缺少『差异审核』或『使用说明』）")
    meta = {str(row[0]): str(row[1]) for row in wb["使用说明"].iter_rows(min_row=2, values_only=True)
            if len(row) >= 2 and row[0] is not None and row[1] is not None}
    if meta.get("app_token") != config["app_token"]:
        raise ValueError("报告来自另一个多维表格 app_token，已终止")
    sh = wb["差异审核"]
    headers = [c.value for c in sh[1]]
    if headers[:len(COLS)] != COLS:
        raise ValueError("审核表列结构已改变，请勿删除、重排或重命名列")

    table_map = {t["table_id"]: t["name"] for t in config["tables"]}
    approvals, issues = [], []
    seen = set()
    duplicate_keys = set()
    total_reviewed = 0
    for excel_row, values in enumerate(sh.iter_rows(min_row=2, max_col=len(COLS), values_only=True), 2):
        if not values or all(v is None for v in values):
            continue
        status = str(values[0] or "待确认").strip()
        if status not in {"待确认", "保留", "覆盖"}:
            issues.append([excel_row, "", "审核状态无效，必须为待确认/保留/覆盖"])
            continue
        if status != "覆盖":
            continue
        total_reviewed += 1
        _, name, table_id, record_id, code, field, old, new, kind, filename, digest, note = values
        if allowed_table_ids is not None and table_id not in allowed_table_ids:
            issues.append([excel_row, f"{name or ''} {code or ''} {field or ''}",
                           "本次未选中该子表：审批项目已跳过，不读取或写入飞书"])
            continue
        key = (table_id, record_id, field)
        try:
            if not all(isinstance(x, str) and x for x in (name, table_id, record_id, code, field, digest)):
                raise ValueError("关键定位字段为空或格式错误")
            if table_id not in table_map or table_map[table_id] != name:
                raise ValueError("报告中的目标子表与配置不一致")
            if field not in FIELD_NAMES:
                raise ValueError("审核字段不是『邮件主题』或『邮件正文』")
            if not isinstance(old, str) or not isinstance(new, str) or not old.strip() or not new.strip():
                raise ValueError("审核旧值/新值为空或不是文本")
            if key in seen:
                duplicate_keys.add(key)
                raise ValueError("同一记录的同一字段被重复审批，所有重复项都将跳过")
            seen.add(key)
            expected = make_digest(config["app_token"], table_id, record_id, code, field, old, new)
            if expected != digest:
                raise ValueError("报告内容发生修改，校验标识不匹配；请重新生成报告")
            approvals.append({"excel_row": excel_row, "name": name, "table_id": table_id,
                              "record_id": record_id, "code": code, "field": field,
                              "old": old, "new": new, "source": filename})
        except Exception as exc:
            issues.append([excel_row, f"{name or ''} {code or ''} {field or ''}", str(exc)])
    if duplicate_keys:
        kept = []
        for item in approvals:
            key = (item["table_id"], item["record_id"], item["field"])
            if key in duplicate_keys:
                issues.append([item["excel_row"], f"{item['name']} {item['code']} {item['field']}", "审批字段重复：为安全起见全部跳过"])
            else:
                kept.append(item)
        approvals = kept
    wb.close()
    return approvals, issues, total_reviewed


def verify_against_feishu(client, approvals):
    """逐条重读候选编号及原值，不满足旧值相等则阻止写入。"""
    good, rejected = [], []
    cache = {}
    for item in approvals:
        key = (item["table_id"], item["record_id"])
        try:
            if key not in cache:
                cache[key] = client.get_record(*key)
            current_fields = cache[key].get("fields") or {}
            current_code = text_value(current_fields.get("候选编号")).strip().upper()
            if current_code != item["code"]:
                raise ValueError(f"候选编号变化：现在为 {current_code}")
            current = text_value(current_fields.get(item["field"]))
            if current != item["old"]:
                raise ValueError("飞书当前内容已与审核报告中的旧值不同：拒绝覆盖，请重新审核")
            good.append(item)
        except Exception as exc:
            rejected.append([item["excel_row"], f"{item['name']} {item['code']} {item['field']}", str(exc)])
    return good, rejected


def save_result(run_id, mode, good, rejected, succeeded, failures, backup_path=None):
    out = Path("报告")
    out.mkdir(exist_ok=True)
    path = out / f"审核覆盖执行结果_{run_id}.xlsx"
    wb = create_workbook()
    add_sheet(wb, "审核通过明细", ["目标子表", "候选编号", "字段名", "旧值", "新值", "Excel行号"],
              [[x["name"], x["code"], x["field"], x["old"], x["new"], x["excel_row"]] for x in good])
    add_sheet(wb, "拒绝和异常", ["Excel行号", "项目", "错误描述"], rejected)
    add_sheet(wb, "写入结果", ["目标子表", "候选编号", "字段名", "执行状态"],
              [[x["name"], x["code"], x["field"], "已覆盖"] for x in succeeded] +
              [[x["name"], x["code"], x["field"], msg] for x, msg in failures])
    add_sheet(wb, "说明", ["项目", "说明"], [
        ["执行方式", "正式执行（实际写入结果见下方）" if mode == "execute" else "仅预检，未写入"],
        ["通过预检字段数", len(good)],
        ["拒绝/异常字段数", len(rejected)],
        ["已覆盖字段数", len(succeeded)],
        ["写入失败字段数", len(failures)],
        ["覆盖前备份", str(backup_path or "无：预检模式没有写入")],
        ["注意", "如审核报告旧值与飞书当前值不符，一律拒绝覆盖。"]
    ])
    wb.save(path)
    return path


def run_write(client, good, run_id):
    # 备份旧内容，先保存到本地，再开始任何覆盖。
    backups = [{"table_id": x["table_id"], "record_id": x["record_id"],
                "候选编号": x["code"], "字段名": x["field"], "旧值": x["old"], "拟写入值": x["new"]} for x in good]
    out = Path("备份")
    out.mkdir(exist_ok=True)
    backup_path = out / f"覆盖前原值_{run_id}.json"
    backup_path.write_text(json.dumps(backups, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已保存覆盖前备份：{backup_path.resolve()}")

    grouped = defaultdict(lambda: defaultdict(dict))
    items_by_key = defaultdict(list)
    for x in good:
        key = (x["table_id"], x["record_id"])
        grouped[x["table_id"]][x["record_id"]][x["field"]] = x["new"]
        items_by_key[key].append(x)

    succeeded, failures = [], []
    for table_id, records in grouped.items():
        keys = list(records.keys())
        for i in range(0, len(keys), 50):
            ids = keys[i:i + 50]
            batch = [{"record_id": rid, "fields": records[rid]} for rid in ids]
            try:
                # 说明：飞书 API 的读取+更新并非数据库原子 CAS；请运行期间避免同事同时修改。
                client.batch_update(table_id, batch)
                for rid in ids:
                    succeeded.extend(items_by_key[(table_id, rid)])
            except Exception as exc:
                for rid in ids:
                    for item in items_by_key[(table_id, rid)]:
                        failures.append((item, f"写入失败或结果待人工核实：{exc}"))
                print(f"  子表 {table_id} 批次失败，请在飞书中核对：{exc}")
    return succeeded, failures, backup_path


def main():
    parser = argparse.ArgumentParser(description="根据人工审核 Excel 选择性覆盖飞书字段")
    parser.add_argument("--report", required=True, help="01 脚本生成并人工审核的 .xlsx 文件")
    parser.add_argument("--config", default="config.json", help="配置文件路径")
    parser.add_argument("--execute", action="store_true", help="执行覆盖；默认仅预检")
    parser.add_argument("--table-id", action="append", default=None, help="本次允许审核/覆盖的子表 ID；可多次传入")
    args = parser.parse_args()
    config = require_config(args.config)
    selected_tables = selected_table_entries(config, args.table_id)
    if not selected_tables:
        print("未选择任何子表：未读取审核报告，未请求飞书，也不会写入。")
        return
    print("本次选中子表：" + "、".join(t["name"] for t in selected_tables))
    selected_ids = {t["table_id"] for t in selected_tables}
    report = Path(args.report)
    if not report.is_file():
        raise FileNotFoundError(report)
    approvals, issues, chosen = read_approvals(report, config, allowed_table_ids=selected_ids)
    skipped = sum("本次未选中该子表" in item[2] for item in issues)
    print(f"Excel 中勾选覆盖：{chosen} 项；未选中子表已跳过：{skipped} 项；"
          f"本次范围内基础校验通过：{len(approvals)} 项；异常/跳过：{len(issues)} 项")
    if not approvals:
        run_id = timestamp()
        path = save_result(run_id, "execute" if args.execute else "preview", [], issues, [], [])
        print(f"当前选中的子表没有可执行覆盖的项目，未请求飞书、未写入。报告：{path.resolve()}")
        return
    client = FeishuClient(config["app_token"])
    good, rejected = verify_against_feishu(client, approvals)
    rejected = issues + rejected
    print(f"飞书旧值二次核对通过：{len(good)} 项；拒绝：{len(rejected)} 项")
    run_id = timestamp()

    if not args.execute:
        path = save_result(run_id, "preview", good, rejected, [], [])
        print(f"预检报告：{path.resolve()}；未修改飞书。")
        return
    if not good:
        path = save_result(run_id, "execute", [], rejected, [], [])
        print(f"没有可安全覆盖的记录。已生成：{path.resolve()}")
        return
    counts = defaultdict(int)
    for item in good:
        counts[item["name"]] += 1
    print("即将覆盖的子表与字段数（仅包含本次选中范围）：")
    for name, count in counts.items():
        print(f"  {name}：{count} 个字段")
    answer = input(f"即将覆盖 {len(good)} 个字段。请确认以上目标子表后，输入全部大写 YES（小写 yes 不执行）：").strip()
    if answer != "YES":
        print("未确认，已取消，无写入。")
        return
    # 确认后再次从飞书获取当前值，降低审核到写入期间的变更风险
    good, rejected_after_confirm = verify_against_feishu(client, good)
    rejected.extend(rejected_after_confirm)
    succeeded, failures, backup = ([], [], None)
    if good:
        succeeded, failures, backup = run_write(client, good, run_id)
    path = save_result(run_id, "execute", good, rejected, succeeded, failures, backup)
    print(f"覆盖成功字段：{len(succeeded)}；拒绝：{len(rejected)}；失败/待核实：{len(failures)}")
    print(f"执行结果：{path.resolve()}")


if __name__ == "__main__":
    main()
