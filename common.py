"""飞书 Word 批量导入：公共逻辑。仅支持文本型候选编号、邮件主题、邮件正文。"""
from __future__ import annotations

import hashlib
import json
import os
import re
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit, parse_qs

import requests
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

BASE = "https://open.feishu.cn/open-apis"
OAUTH_URL = "https://accounts.feishu.cn/oauth/v3/token"
FIELD_NAMES = ("邮件主题", "邮件正文")
CODE_PATTERN = re.compile(r"^(C\d+)(?=$|[_\-\s])", re.I)
MAX_EXCEL_CHARS = 32767


def timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


# 飞书网址示例：
# https://xxx.feishu.cn/base/MNCRxxxx?table=tblABC123&view=vewXXXX
# app_token 只能填 MNCRxxxx；table_id 只能填 tblABC123，不能带 base/ 或 &view=...
# 下方函数也兼容用户误复制路径或完整网址，自动提取并在运行前提示。
_APP_TOKEN_RE = re.compile(r"^[A-Za-z0-9_-]+$")
_TABLE_ID_RE = re.compile(r"^tbl[A-Za-z0-9_-]+$")


def clean_app_token(value):
    """解析纯 app_token、base/xxx 或 https://.../base/xxx?... 。"""
    if not isinstance(value, str):
        raise ValueError("app_token 必须是字符串")
    raw = value.strip().strip('"').strip("'")
    if not raw or "填写" in raw or "示例" in raw:
        raise ValueError("app_token 尚未填写真实值")
    if raw.lower().startswith(("https://", "http://")):
        parts = urlsplit(raw)
        match = re.search(r"(?:^|/)base/([A-Za-z0-9_-]+)(?:/|$)", parts.path)
        if not match:
            raise ValueError("网址未包含 /base/多维表格token；如为 /wiki/ 链接，请先获取真正的 app_token")
        token = match.group(1)
    elif raw.startswith(("base/", "/base/")):
        match = re.fullmatch(r"/?base/([A-Za-z0-9_-]+)(?:\?.*)?", raw)
        if not match:
            raise ValueError("app_token 格式不正确：base/ 后只能是 token，不要粘贴额外路径")
        token = match.group(1)
    else:
        token = raw
    if not _APP_TOKEN_RE.fullmatch(token) or token.startswith("cli_"):
        raise ValueError("app_token 必须是多维表格的纯 ID（不是 App ID），不能包含 base/、?table= 或斜杠")
    return token


def clean_table_id(value):
    """解析纯 table_id、tblxxx&view=xxx 或完整飞书多维表格网址。"""
    if not isinstance(value, str):
        raise ValueError("table_id 必须是字符串")
    raw = value.strip().strip('"').strip("'")
    if not raw or "填写" in raw or "示例" in raw:
        raise ValueError("table_id 尚未填写真实值")
    if raw.lower().startswith(("https://", "http://")):
        query = parse_qs(urlsplit(raw).query)
        values = query.get("table", [])
        if len(values) != 1:
            raise ValueError("子表网址缺少 ?table=tbl...，无法提取 table_id")
        table_id = values[0]
    elif raw.startswith("?table=") or raw.startswith("table="):
        values = parse_qs(raw.lstrip("?")).get("table", [])
        if len(values) != 1:
            raise ValueError("不能从 table= 参数中唯一提取 table_id")
        table_id = values[0]
    else:
        # 兼容错误粘贴的 tblxxx&view=vewxxx；view_id 不是 table_id。
        table_id = re.split(r"[?&#]", raw, maxsplit=1)[0]
    if not _TABLE_ID_RE.fullmatch(table_id):
        raise ValueError("table_id 应为 tbl 开头的纯 ID，例如 tblAbc123；不能带 &view=、路径或子表名称")
    return table_id


def require_config(path):
    p = Path(path)
    if not p.exists():
        raise ValueError(f"配置文件不存在：{p}。请复制 config.example.json 为 config.json，并填写真实值。")
    config = json.loads(p.read_text(encoding="utf-8-sig"))
    if not isinstance(config, dict):
        raise ValueError("config.json 必须是 JSON 对象")
    original_token = config.get("app_token", "")
    token = clean_app_token(original_token)
    if token != original_token:
        print(f"[配置自动修正] app_token 已提取纯 ID：{token}（建议同步修改 config.json）")
    config["app_token"] = token
    tables = config.get("tables")
    if not isinstance(tables, list) or not tables:
        raise ValueError("config.json 的 tables 必须是非空列表")
    seen = set()
    for t in tables:
        if not isinstance(t, dict):
            raise ValueError("tables 中每个项目都必须是对象")
        name, original_id = t.get("name", ""), t.get("table_id", "")
        if not isinstance(name, str) or not name.strip():
            raise ValueError("tables 中有未命名子表")
        try:
            table_id = clean_table_id(original_id)
        except ValueError as exc:
            raise ValueError(f"子表『{name}』的 table_id 错误：{exc}") from exc
        if table_id != original_id:
            print(f"[配置自动修正] {name} 的 table_id 已提取纯 ID：{table_id}（已忽略 view 视图参数）")
        if table_id in seen:
            raise ValueError(f"配置中 table_id 重复：{table_id}；不同子表必须使用不同 table_id")
        seen.add(table_id)
        t["table_id"] = table_id
    return config



def selected_table_entries(config, selected_ids=None):
    """按真实 table_id 限定本轮作用范围，拒绝未配置或重复的 ID。

    selected_ids=None 表示命令行直接调用时沿用全部配置子表；
    启动菜单总会显式传入选中的 ID，没有选择时禁止调用业务脚本。
    """
    tables = config["tables"]
    if selected_ids is None:
        return list(tables)
    if not selected_ids:
        return []
    ids = [clean_table_id(x) for x in selected_ids]
    if len(ids) != len(set(ids)):
        raise ValueError("选择了重复的子表 ID，请重新选择")
    available = {t["table_id"]: t for t in tables}
    unknown = [id_ for id_ in ids if id_ not in available]
    if unknown:
        raise ValueError(f"所选子表不属于当前 config.json：{', '.join(unknown)}；已停止")
    return [available[id_] for id_ in ids]

def text_value(value):
    """把多维表格文本字段读取结果转为纯文本；拒绝无法判断的非文本类型。"""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, list):
        if all(isinstance(x, dict) and isinstance(x.get("text"), str) for x in value):
            return "".join(x["text"] for x in value)
    raise ValueError(f"字段不是可处理的文本类型（{type(value).__name__}），已安全跳过")


def is_blank(s):
    return not s or not s.strip()


def normalize_for_compare(s):
    return re.sub(r"\s+", "", s)


def filename_code(path):
    """只按文件名开头的候选编号匹配；姓名仅供人工核对，不参与飞书匹配。

    如 C001_候选人姓名.docx / C062_张三.docx / C125.docx -> C001 / C062 / C125。
    要求编号在最开头，后面为空或以 _、-、空格分隔，避免误取文件名其他数字。
    """
    stem = Path(path).stem.strip()
    match = CODE_PATTERN.match(stem)
    if not match:
        raise ValueError("文件名应为『C001_候选人姓名.docx』，且以候选编号开头")
    code = match.group(1).upper()
    rest = stem[match.end():]
    # 发现第二个明确的候选编号时不猜测，交给人工处理。
    if re.search(r"(?<![A-Z0-9])C\d+(?![A-Z0-9])", rest, re.I):
        raise ValueError("文件名中疑似包含多个候选编号，已跳过")
    return code


def read_docx(path):
    """提取 Word 正文顺序的段落与表格文本；首条非空行必须为主题。"""
    doc = Document(str(path))
    lines = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            para = Paragraph(child, doc)
            lines.extend(para.text.split("\n"))
        elif child.tag == qn("w:tbl"):
            table = Table(child, doc)
            for row in table.rows:
                lines.append("\t".join(cell.text for cell in row.cells))
    first = next((i for i, s in enumerate(lines) if s.strip()), None)
    if first is None:
        raise ValueError("Word 文档为空")
    match = re.match(r"^主题\s*[：:]\s*(.+?)\s*$", lines[first].strip(), re.S)
    if not match:
        raise ValueError("首条非空行必须是『主题：...』")
    subject = match.group(1).strip()
    body = "\n".join(lines[first + 1:])
    if not subject or is_blank(body):
        raise ValueError("邮件主题或正文为空")
    if any(len(s) > MAX_EXCEL_CHARS for s in (subject, body)):
        raise ValueError("单字段内容超过 Excel 单元格 32767 字上限，跳过避免审核报告丢失文本")
    return {"邮件主题": subject, "邮件正文": body}


def make_digest(app_token, table_id, record_id, code, field, old, new):
    payload = [app_token, table_id, record_id, code, field, old, new]
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")).hexdigest()


def index_records(records):
    idx = defaultdict(list)
    for record in records:
        raw = (record.get("fields") or {}).get("候选编号")
        try:
            code = text_value(raw).strip().upper()
        except ValueError:
            continue
        if code:
            idx[code].append(record)
    return idx


class FeishuClient:
    def __init__(self, app_token):
        self.app_token = clean_app_token(app_token)
        self.app_id = os.getenv("FEISHU_APP_ID", "")
        self.app_secret = os.getenv("FEISHU_APP_SECRET", "")
        if not self.app_id or not self.app_secret:
            raise ValueError("缺少环境变量 FEISHU_APP_ID / FEISHU_APP_SECRET；请参照 README 配置。")
        self.session = requests.Session()
        self.token = None
        self.expire_at = 0
        self.last_api_call_at = 0.0

    def _refresh(self):
        # 新版 OAuth 2.0 客户端凭据模式（2026 年文档）
        r = self.session.post(OAUTH_URL, json={
            "grant_type": "client_credentials",
            "client_id": self.app_id,
            "client_secret": self.app_secret,
        }, timeout=30)
        r.raise_for_status()
        data = r.json()
        access = data.get("access_token")
        if not access:
            raise RuntimeError(f"获取飞书访问凭证失败：{data.get('error', data.get('code', '未知错误'))}")
        self.token = access
        self.expire_at = time.monotonic() + max(60, int(data.get("expires_in", 3600)) - 120)

    def request(self, method, path, *, params=None, payload=None):
        if self.token is None or time.monotonic() >= self.expire_at:
            self._refresh()
        # 保守控制 API 调用节奏：单客户端约 8 QPS，低于常见的 10 QPS 额度。
        pause = 0.13 - (time.monotonic() - self.last_api_call_at)
        if pause > 0:
            time.sleep(pause)
        self.last_api_call_at = time.monotonic()
        r = self.session.request(method, BASE + path,
            headers={"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"},
            params=params, json=payload, timeout=60)
        try:
            r.raise_for_status()
        except requests.HTTPError as exc:
            if r.status_code == 404 and "/bitable/v1/apps/" in path:
                raise RuntimeError(
                    "飞书返回 HTTP 404：请核对 app_token、table_id 是否属于同一多维表格，"
                    "以及应用是否有访问该表格的权限。config.json 只能填写纯 ID；"
                    "app_token 不带 base/，table_id 不带 &view=。"
                ) from exc
            raise
        result = r.json()
        if result.get("code") != 0:
            raise RuntimeError(f"飞书 API 错误 code={result.get('code')} msg={result.get('msg')} logid={r.headers.get('X-Tt-Logid', '')}")
        return result.get("data") or {}

    def base_path(self, table_id):
        safe_table_id = clean_table_id(table_id)
        return f"/bitable/v1/apps/{self.app_token}/tables/{safe_table_id}/records"

    def list_records(self, table_id):
        all_records, page_token = [], None
        while True:
            params = {"page_size": 500}
            if page_token:
                params["page_token"] = page_token
            data = self.request("GET", self.base_path(table_id), params=params)
            all_records.extend(data.get("items") or [])
            if not data.get("has_more"):
                break
            if not data.get("page_token") or data.get("page_token") == page_token:
                raise RuntimeError("飞书分页返回异常，已停止以避免遗漏记录")
            page_token = data["page_token"]
        return all_records

    def get_record(self, table_id, record_id):
        data = self.request("GET", self.base_path(table_id) + "/" + record_id)
        rec = data.get("record")
        if not rec:
            raise RuntimeError(f"无法读取记录：{record_id}")
        return rec

    def batch_update(self, table_id, records):
        if not records:
            return
        if len(records) > 50:
            raise ValueError("每次提交最多 50 条（工具的保守上限）")
        self.request("POST", self.base_path(table_id) + "/batch_update", payload={"records": records})


def create_workbook():
    """运行端通过 openpyxl 生成可审核 Excel；所有文本强制按字符串保存。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    wb.remove(wb.active)
    return wb


def add_sheet(wb, name, headers, data, *, review=False):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    sh = wb.create_sheet(name)
    sh.append(headers)
    for row in data:
        sh.append(["" if v is None else v for v in row])
    for row in sh.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"  # 防止 Word 正文含 '=' 时被 Excel 当作公式
    sh.freeze_panes = "C2" if review else "A2"
    sh.auto_filter.ref = f"A1:{sh.cell(max(1, sh.max_row), len(headers)).coordinate}"
    for c in sh[1]:
        c.fill = PatternFill("solid", fgColor="17365D")
        c.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        c.alignment = Alignment(vertical="center")
    sh.row_dimensions[1].height = 28
    for j, name_ in enumerate(headers, 1):
        width = 18
        if name_ in ("飞书现值", "Word新值", "旧值", "新值"):
            width = 58
        elif name_ in ("来源文件", "说明", "错误描述"):
            width = 42
        elif name_ in ("校验标识", "table_id", "record_id"):
            width = 24
        sh.column_dimensions[__import__("openpyxl").utils.get_column_letter(j)].width = width
    for row in sh.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
        sh.row_dimensions[row[0].row].height = 60 if review else 32
    if review and sh.max_row >= 2:
        dv = DataValidation(type="list", formula1='"待确认,覆盖,保留"', allow_blank=False)
        dv.error = "只能选择：待确认、覆盖、保留"
        dv.errorTitle = "审批状态无效"
        dv.showErrorMessage = True
        sh.add_data_validation(dv)
        dv.add(f"A2:A{sh.max_row}")
        for c in sh["A"][1:]:
            c.fill = PatternFill("solid", fgColor="FFF2CC")
    return sh
