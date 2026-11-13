"""表格名称 API、配置兼容及后续预览/审批流程测试，全程模拟飞书。"""
import importlib.util
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from common import FeishuClient, require_config, resolve_feishu_tables
from docx import Document
from openpyxl import load_workbook
import read_feishu_tables


def load_script(name, filename):
    spec = importlib.util.spec_from_file_location(name, ROOT / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


stage1 = load_script("catalog_stage1", "01_导入并生成差异报告.py")
stage2 = load_script("catalog_stage2", "02_执行审核覆盖.py")


def metadata_client():
    client = object.__new__(FeishuClient)
    client.app_token = "BaseTEST"
    return client


class CatalogTests(unittest.TestCase):
    def test_metadata_paths_and_pagination(self):
        client = metadata_client()
        with patch.object(client, "request", side_effect=[
            {"app": {"name": "招聘管理（海外）"}},
            {"items": [{"table_id": "tblTech", "name": "技术 & 产品"}],
             "has_more": True, "page_token": "page2"},
            {"items": [{"table_id": "tblOps", "name": "运营团队"}], "has_more": False},
        ]) as request:
            config = resolve_feishu_tables({"app_token": "BaseTEST"}, client)
        self.assertEqual(config["app_name"], "招聘管理（海外）")
        self.assertEqual([t["name"] for t in config["tables"]], ["技术 & 产品", "运营团队"])
        self.assertEqual(request.call_args_list[0].args, ("GET", "/bitable/v1/apps/BaseTEST"))
        self.assertEqual(request.call_args_list[1].args, ("GET", "/bitable/v1/apps/BaseTEST/tables"))
        self.assertEqual(request.call_args_list[2].kwargs["params"],
                         {"page_size": 100, "page_token": "page2"})

    def test_configured_scope_order_and_external_folders_are_preserved(self):
        config = {"app_token": "BaseTEST", "tables": [
            {"name": "原运营名称", "table_id": "tblOps", "folder": "D:\\任意目录\\Word 文件"},
            {"name": "原技术名称", "table_id": "tblTech", "folder": ""},
        ]}
        original = json.loads(json.dumps(config))
        client = metadata_client()
        with patch.object(client, "get_app", return_value={"name": "真实表格"}), \
                patch.object(client, "list_tables", return_value=[
                    {"table_id": "tblTech", "name": "技术岗位"},
                    {"table_id": "tblExtra", "name": "未配置子表"},
                    {"table_id": "tblOps", "name": "运营岗位"},
                ]):
            result = resolve_feishu_tables(config, client)
        self.assertEqual(config, original)
        self.assertEqual([t["table_id"] for t in result["tables"]], ["tblOps", "tblTech"])
        self.assertEqual(result["tables"][0]["folder"], "D:\\任意目录\\Word 文件")
        self.assertEqual(result["tables"][0]["name"], "运营岗位")
        self.assertEqual(result["tables"][0]["configured_name"], "原运营名称")

    def test_optional_names_and_minimal_config_remain_offline(self):
        with tempfile.TemporaryDirectory() as td:
            config_path = Path(td) / "config.json"
            config_path.write_text(json.dumps({"app_token": "https://example.feishu.cn/base/BaseTEST?table=tblTech"}),
                                   encoding="utf-8")
            with patch("common.FeishuClient", side_effect=AssertionError("不应联网")):
                config = require_config(config_path, allow_discovery=True)
            self.assertEqual(config["app_token"], "BaseTEST")
            self.assertEqual(config["tables"], [])
            config_path.write_text(json.dumps({"app_token": "BaseTEST", "tables": [{"table_id": "tblTech"}]}),
                                   encoding="utf-8")
            config = require_config(config_path, allow_discovery=True)
            self.assertEqual(config["tables"][0]["table_id"], "tblTech")
            with self.assertRaises(ValueError):
                require_config(config_path)  # 原业务入口的配置校验仍保留。

    def test_invalid_metadata_and_pagination_stop(self):
        client = metadata_client()
        with patch.object(client, "request", return_value={"app": {}}):
            with self.assertRaisesRegex(RuntimeError, "多维表格名称"):
                client.get_app()
        for responses in [
            [{"items": [], "has_more": True}],
            [{"items": [], "has_more": True, "page_token": "A"},
             {"items": [], "has_more": True, "page_token": "B"},
             {"items": [], "has_more": True, "page_token": "A"}],
            [{"items": [{"table_id": "tblTech", "name": "技术"},
                        {"table_id": "tblTech", "name": "重复"}]}],
            [{"items": [{"table_id": "tblTech", "name": ""}]}],
        ]:
            with self.subTest(responses=responses), patch.object(client, "request", side_effect=responses):
                with self.assertRaises(RuntimeError):
                    client.list_tables()
        for tables in [[], [{"table_id": "tblOther", "name": "另一子表"}]]:
            with patch.object(client, "get_app", return_value={"name": "表格"}), \
                    patch.object(client, "list_tables", return_value=tables):
                with self.assertRaises(ValueError):
                    resolve_feishu_tables({"app_token": "BaseTEST", "tables": [
                        {"table_id": "tblMissing", "name": "失效子表"}]}, client)

    def test_failed_catalog_does_not_publish_session_config(self):
        with tempfile.TemporaryDirectory() as td:
            output = Path(td) / "session.json"
            args = ["catalog", "--app-token", "BaseTEST", "--output", str(output)]
            with patch.object(sys, "argv", args), \
                    patch.object(read_feishu_tables, "resolve_feishu_tables", side_effect=RuntimeError("无权限")):
                with self.assertRaisesRegex(RuntimeError, "无权限"):
                    read_feishu_tables.main()
            self.assertFalse(output.exists())

    def test_link_only_startup_discovers_all_tables_without_saving_credentials(self):
        with tempfile.TemporaryDirectory() as td:
            output = Path(td) / "session.json"
            args = ["catalog", "--app-token", "https://example.feishu.cn/base/BaseTEST?table=tblTech&view=vewTEST",
                    "--output", str(output)]
            with patch.object(sys, "argv", args), \
                    patch.dict(os.environ, {"FEISHU_APP_ID": "dummy-id", "FEISHU_APP_SECRET": "dummy-secret"}), \
                    patch.object(FeishuClient, "request", side_effect=[
                        {"app": {"name": "招聘管理"}},
                        {"items": [{"table_id": "tblTech", "name": "技术岗位"},
                                   {"table_id": "tblOps", "name": "运营岗位"}]},
                    ]):
                read_feishu_tables.main()
            result = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual(result["app_token"], "BaseTEST")
            self.assertEqual([t["table_id"] for t in result["tables"]], ["tblTech", "tblOps"])
            self.assertTrue(all(t["folder"] == "" for t in result["tables"]))
            self.assertNotIn("dummy-id", output.read_text(encoding="utf-8"))
            self.assertNotIn("dummy-secret", output.read_text(encoding="utf-8"))

    def test_session_config_flows_into_existing_preview_and_approval(self):
        previous_cwd = Path.cwd()
        try:
            with tempfile.TemporaryDirectory() as td, tempfile.TemporaryDirectory() as external:
                os.chdir(td)
                folder = Path(external)  # Word 目录可以位于工具工作目录之外。
                doc = Document()
                doc.add_paragraph("主题：新的主题")
                doc.add_paragraph("新的邮件正文")
                doc.save(folder / "C001_姓名.docx")
                source = Path(td) / "config.json"
                source.write_text(json.dumps({"app_token": "BaseTEST", "tables": [
                    {"name": "旧显示名", "table_id": "tblTech", "folder": str(folder)}]}), encoding="utf-8")
                original_bytes = source.read_bytes()
                output = Path(td) / "session.json"
                args = ["catalog", "--config", str(source), "--output", str(output)]
                with patch.object(sys, "argv", args), \
                        patch.dict(os.environ, {"FEISHU_APP_ID": "dummy-id", "FEISHU_APP_SECRET": "dummy-secret"}), \
                        patch.object(FeishuClient, "request", side_effect=[
                            {"app": {"name": "真实多维表格"}},
                            {"items": [{"table_id": "tblTech", "name": "技术 & 产品（海外）"}]},
                        ]) as request:
                    read_feishu_tables.main()
                self.assertTrue(all(call.args[0] == "GET" for call in request.call_args_list))
                self.assertNotIn("dummy-secret", output.read_text(encoding="utf-8"))
                self.assertEqual(source.read_bytes(), original_bytes)
                resolved = require_config(output)
                self.assertEqual(resolved["app_name"], "真实多维表格")

                class Records:
                    app_token = "BaseTEST"

                    def __init__(self):
                        self.list_calls = []

                    def list_records(self, table_id):
                        self.list_calls.append(table_id)
                        return [{"record_id": "recTech", "fields": {
                            "候选编号": "C001", "邮件主题": "旧主题", "邮件正文": "新的邮件正文"}}]

                records = Records()
                with patch.object(sys, "argv", ["stage1", "--config", str(output), "--no-dialog", "--table-id", "tblTech"]), \
                        patch.object(stage1, "FeishuClient", return_value=records):
                    stage1.main()
                self.assertEqual(records.list_calls, ["tblTech"])
                report = next(Path("报告").glob("差异审核_*.xlsx"))
                wb = load_workbook(report)
                self.assertEqual(wb["差异审核"]["B2"].value, "技术 & 产品（海外）")
                metadata = {row[0]: row[1] for row in wb["使用说明"].iter_rows(values_only=True)}
                self.assertEqual(metadata["多维表格名称"], "真实多维表格")
                wb["差异审核"]["A2"] = "覆盖"
                wb.save(report)
                wb.close()
                good, issues, _ = stage2.read_approvals(report, resolved, {"tblTech"})
                self.assertFalse(issues)
                self.assertEqual(good[0]["name"], "技术 & 产品（海外）")
                wb = load_workbook(report)
                wb["差异审核"]["B2"] = "旧显示名"
                wb.save(report)
                wb.close()
                good, issues, _ = stage2.read_approvals(report, resolved, {"tblTech"})
                self.assertFalse(issues)
                self.assertEqual(good[0]["name"], "技术 & 产品（海外）")
                wb = load_workbook(report)
                wb["差异审核"]["B2"] = "其他名称"
                wb.save(report)
                wb.close()
                good, issues, _ = stage2.read_approvals(report, resolved, {"tblTech"})
                self.assertFalse(good)
                self.assertIn("目标子表与配置不一致", issues[0][2])
        finally:
            os.chdir(previous_cwd)


if __name__ == "__main__":
    unittest.main()
