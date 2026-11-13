"""离线集成测试：不访问真实飞书。运行 python tests/offline_test.py"""
import importlib.util
import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, ROOT / path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

s1 = load("stage1", "01_导入并生成差异报告.py")
s2 = load("stage2", "02_执行审核覆盖.py")

from common import read_docx
from docx import Document
from openpyxl import load_workbook


class FakeClient:
    app_token = "appTEST"
    def __init__(self):
        self.data = {
            "tblT1": [{"record_id": "recT1", "fields": {"候选编号": "C062", "候选人": "王同学", "邮件主题": "旧主题", "邮件正文": ""}}],
            "tblT2": [{"record_id": "recT2", "fields": {"候选编号": "C062", "候选人": "孔同学", "邮件主题": "", "邮件正文": ""}}],
        }
    def list_records(self, table_id):
        return self.data[table_id]
    def get_record(self, table_id, record_id):
        return next(x for x in self.data[table_id] if x["record_id"] == record_id)
    def batch_update(self, table_id, records):
        for update in records:
            target = self.get_record(table_id, update["record_id"])
            target["fields"].update(update["fields"])


def create_docx(path):
    doc = Document()
    doc.add_paragraph("主题：新标题邀请")
    doc.add_paragraph("[候选人姓名]同学，你好：")
    doc.add_paragraph("")
    doc.add_paragraph("你好，这里是正文。")
    doc.save(path)


def test():
    with tempfile.TemporaryDirectory() as td:
        os.chdir(td)
        t1 = Path(td) / "T1"
        t2 = Path(td) / "T2"
        t1.mkdir(); t2.mkdir()
        create_docx(t1 / "C062_祝金国.docx")
        create_docx(t2 / "C062_候选人姓名不相同.docx")
        assert read_docx(t1 / "C062_祝金国.docx")["邮件正文"].startswith("[候选人姓名]同学")
        client = FakeClient()
        c1,d1,a1,e1 = s1.scan_table(client,"T1 mapping","tblT1",t1)
        c2,d2,a2,e2 = s1.scan_table(client,"T2 mapping","tblT2",t2)
        assert len(c1)==1 and len(c2)==1
        assert len(d1)==1 and len(d2)==0, (d1,d2)
        assert c1[0]["fields"].keys()=={"邮件正文"}
        assert c2[0]["fields"].keys()=={"邮件主题","邮件正文"}
        details = a1+a2; exceptions=e1+e2
        s1.apply_blank_fills(client,c1+c2,details,exceptions)
        assert client.data["tblT1"][0]["fields"]["邮件主题"] == "旧主题"
        assert client.data["tblT1"][0]["fields"]["邮件正文"].startswith("[候选人姓名]")
        assert client.data["tblT2"][0]["fields"]["邮件主题"] == "新标题邀请"
        assert client.data["tblT1"][0]["fields"]["候选人"] == "王同学"
        config={"app_token":"appTEST","tables":[{"name":"T1 mapping","table_id":"tblT1"},{"name":"T2 mapping","table_id":"tblT2"}]}
        report=s1.save_report("test",config,"fill",d1+d2,details,exceptions)
        wb=load_workbook(report)
        assert wb["差异审核"]["B2"].value=="T1 mapping"
        wb["差异审核"]["A2"]="覆盖"
        wb.save(report)
        approved,issues,n=s2.read_approvals(report,config)
        assert len(approved)==1 and not issues
        good,rejected=s2.verify_against_feishu(client,approved)
        assert len(good)==1 and not rejected
        # 审核期间有人改动内容，必须拒绝
        client.data["tblT1"][0]["fields"]["邮件主题"]="人工修改过的标题"
        good,rejected=s2.verify_against_feishu(client,approved)
        assert len(good)==0 and len(rejected)==1
        client.data["tblT1"][0]["fields"]["邮件主题"]="旧主题"
        good,rejected=s2.verify_against_feishu(client,approved)
        succeeded,failures,backup=s2.run_write(client,good,"test")
        assert len(succeeded)==1 and not failures and backup.exists()
        assert client.data["tblT1"][0]["fields"]["邮件主题"]=="新标题邀请"
        assert client.data["tblT2"][0]["fields"]["邮件主题"]=="新标题邀请"
        assert client.data["tblT1"][0]["fields"]["候选人"]=="王同学"
        print("离线测试全部通过：多子表独立匹配、空白填写、差异审核、校验拒绝、审批覆盖及备份")


if __name__ == "__main__":
    test()
