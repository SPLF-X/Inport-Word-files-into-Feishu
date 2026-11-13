"""纯本地配置解析、ID 路径、文件名编码测试。"""
import json
import sys
import tempfile
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from common import clean_app_token, clean_table_id, filename_code, require_config, FeishuClient


def test():
    app_token = "BaseDemoToken123"
    table = "tblDemoT1ABC"
    assert clean_app_token("base/" + app_token) == app_token
    assert clean_app_token("https://example.feishu.cn/base/" + app_token + "?table=" + table + "&view=vewDemo123") == app_token
    assert clean_app_token(app_token) == app_token
    assert clean_table_id(table + "&view=vewDemo123") == table
    assert clean_table_id("https://example.feishu.cn/base/" + app_token + "?table=" + table + "&view=vewDemo123") == table
    assert clean_table_id(table) == table
    assert filename_code("C001_候选人姓名.docx") == "C001"
    assert filename_code("c062_张三.docx") == "C062"
    assert filename_code("C125.docx") == "C125"
    for invalid in ("01_祝金国.docx", "C001C002_姓名.docx", "C001_姓名_C002.docx"):
        try:
            filename_code(invalid)
        except ValueError:
            pass
        else:
            raise AssertionError("未能拒绝不明确文件名: " + invalid)
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp)/"config.json"
        path.write_text(json.dumps({"app_token":"base/"+app_token,
            "tables":[{"name":"T1 mapping","table_id":table+"&view=vewDemo123"},
                      {"name":"T2 mapping","table_id":"tblT2TEST"}]},ensure_ascii=False),encoding="utf-8")
        cfg = require_config(path)
        assert cfg["app_token"] == app_token
        assert cfg["tables"][0]["table_id"] == table
        # 不需要真实凭证：只验证最终拼出来的 URL 路径无 base/ 和 view。
        client = object.__new__(FeishuClient)
        client.app_token = cfg["app_token"]
        assert client.base_path(cfg["tables"][0]["table_id"]) == f"/bitable/v1/apps/{app_token}/tables/{table}/records"
    print("离线 ID 解析与候选编号测试通过")


if __name__ == "__main__":
    test()
