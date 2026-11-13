"""V3.1 作用范围隔离测试（全程模拟飞书，不发真实请求）。"""
import importlib.util
import json
import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from common import selected_table_entries
from docx import Document
from openpyxl import load_workbook


def load(name, filename):
    spec = importlib.util.spec_from_file_location(name, ROOT / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

s1 = load("v31_stage1", "01_导入并生成差异报告.py")
s2 = load("v31_stage2", "02_执行审核覆盖.py")


class FakeClient:
    app_token = "appTEST"

    def __init__(self):
        self.data = {
            "tblT1": [{"record_id": "recT1", "fields": {
                "候选编号": "C001", "候选人": "张三", "邮件主题": "T1旧标题", "邮件正文": ""}}],
            "tblT2": [{"record_id": "recT2", "fields": {
                "候选编号": "C001", "候选人": "其他姓名", "邮件主题": "T2旧标题", "邮件正文": ""}}],
        }
        self.list_calls = []
        self.get_calls = []
        self.update_calls = []

    def list_records(self, table_id):
        self.list_calls.append(table_id)
        return self.data[table_id]

    def get_record(self, table_id, record_id):
        self.get_calls.append(table_id)
        return next(x for x in self.data[table_id] if x["record_id"] == record_id)

    def batch_update(self, table_id, records):
        self.update_calls.append(table_id)
        for update in records:
            self.get_record(table_id, update["record_id"])["fields"].update(update["fields"])


def word(path):
    doc = Document()
    doc.add_paragraph("主题：Word新标题")
    doc.add_paragraph("邮件正文")
    doc.save(path)


def run():
    config = {"app_token": "appTEST", "tables": [
        {"name": "T1 mapping", "table_id": "tblT1"},
        {"name": "T2 mapping", "table_id": "tblT2"},
    ]}
    assert len(selected_table_entries(config, ["tblT1"])) == 1
    assert len(selected_table_entries(config, ["tblT1", "tblT2"])) == 2
    assert selected_table_entries(config, []) == []
    assert len(selected_table_entries(config, None)) == 2  # 仅支持直接命令行旧行为
    for items in (["tblBad"], ["tblT1", "tblT1"]):
        try:
            selected_table_entries(config, items)
        except ValueError:
            pass
        else:
            raise AssertionError("未知/重复子表没有拒绝")
    with tempfile.TemporaryDirectory() as tmp:
        os.chdir(tmp)
        t1 = Path(tmp) / "T1"
        t2 = Path(tmp) / "T2"
        t1.mkdir()
        t2.mkdir()
        word(t1 / "C001_张三.docx")
        word(t2 / "C001_和飞书姓名不一致.docx")
        config["tables"][0]["folder"] = str(t1)
        config["tables"][1]["folder"] = str(t2)
        configpath = Path(tmp) / "config.json"
        configpath.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
        client = FakeClient()

        # 只选 T1 的菜单1预览：不得读取 T2 记录或文件夹
        with patch.object(s1, "FeishuClient", lambda token: client), patch.object(sys, "argv", [
            "stage1", "--config", str(configpath), "--no-dialog", "--table-id", "tblT1"
        ]):
            s1.main()
        assert client.list_calls == ["tblT1"], client.list_calls
        reports = list(Path(tmp, "报告").glob("差异审核_*.xlsx"))
        assert len(reports) == 1
        wb = load_workbook(reports[0], data_only=True)
        assert wb["字段处理明细"]["A2"].value == "T1 mapping"
        assert wb["使用说明"]["B5"].value == "T1 mapping"  # 第四行为本轮子表范围
        wb.close()
        assert not client.update_calls

        # 菜单2只选 T1：只能补充 T1 正文，不能调用 T2 写入。
        with patch.object(s1, "FeishuClient", lambda token: client), patch.object(sys, "argv", [
            "stage1", "--config", str(configpath), "--no-dialog", "--apply-blanks", "--table-id", "tblT1"
        ]), patch("builtins.input", lambda prompt: "YES"):
            s1.main()
        assert client.update_calls == ["tblT1"], client.update_calls
        assert client.data["tblT1"][0]["fields"]["邮件正文"].startswith("邮件正文")
        assert client.data["tblT2"][0]["fields"]["邮件正文"] == ""
        client.update_calls.clear()

        # 制造一份 T1+T2 均有待审批项的 Excel
        c1, d1, det1, e1 = s1.scan_table(client, "T1 mapping", "tblT1", t1)
        c2, d2, det2, e2 = s1.scan_table(client, "T2 mapping", "tblT2", t2)
        assert len(d1) == len(d2) == 1
        report = s1.save_report("v31_review", config, "preview", d1 + d2, det1 + det2, e1 + e2, config["tables"])
        wb = load_workbook(report)
        wb["差异审核"]["A2"] = "覆盖"
        wb["差异审核"]["A3"] = "覆盖"
        wb.save(report)
        approvals, issues, total = s2.read_approvals(report, config, {"tblT1"})
        assert total == 2 and len(approvals) == 1 and len(issues) == 1
        assert approvals[0]["table_id"] == "tblT1"
        assert "本次未选中该子表" in issues[0][2]
        # 菜单4只允许 T1，实际模拟覆盖后 T2 不变，也不请求 T2 API
        client.get_calls.clear()
        client.update_calls.clear()
        with patch.object(s2, "FeishuClient", lambda token: client), patch.object(sys, "argv", [
            "stage2", "--config", str(configpath), "--report", str(report), "--execute", "--table-id", "tblT1"
        ]), patch("builtins.input", lambda prompt: "YES"):
            s2.main()
        assert client.update_calls == ["tblT1"], client.update_calls
        assert set(client.get_calls) == {"tblT1"}, client.get_calls
        assert client.data["tblT1"][0]["fields"]["邮件主题"] == "Word新标题"
        assert client.data["tblT2"][0]["fields"]["邮件主题"] == "T2旧标题"
        resultfiles = list(Path(tmp, "报告").glob("审核覆盖执行结果_*.xlsx"))
        assert resultfiles
        result = load_workbook(resultfiles[-1], data_only=True)
        skipped = [c.value for row in result["拒绝和异常"].iter_rows() for c in row]
        assert any(isinstance(x, str) and "本次未选中该子表" in x for x in skipped)
        result.close()

    print("V3.1 子表选择离线测试通过：单选/多选/不选、范围预览、跨表审批过滤及 T2 不写入")


if __name__ == "__main__":
    run()
