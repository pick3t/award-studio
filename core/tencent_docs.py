"""腾讯文档成绩表读取：把在线表格拉下来，落成一份本地 xlsx。

设计要点
--------
* **两种通道，同一套接口**（`Transport`）：
  - `McpTransport`：走 `docs.qq.com/api/v6/sheet/mcp`，票据由宿主注入。
    本机开发/联调用它，不做分发。
  - `OpenApiTransport`：走 `/openapi/spreadsheet/*`，用本应用自己的
    OAuth 三元组（`Client-Id` / `Open-Id` / `Access-Token`）。正式分发用这个。
* **只负责取数，不碰解析**：拉下来写成 xlsx，交给 `core.excel_reader`
  走原来的表头识别/多评委求和那一套，云表和本地文件的行为因此完全一致。
* **单次读单元格上限 20000 格**（腾讯文档的硬限制），所以按页读、
  遇到整页空白就停，避免把 200 行 × 26 列的空网格全拉一遍。

对外主入口：`sync_to_cache(url, cache_dir)` → `SyncResult`
"""

from __future__ import annotations

import base64
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

# 单次请求允许的最大单元格数（腾讯文档硬限制 20000，留一点余量）
MAX_CELLS_PER_REQUEST = 18000
# 最多读多少行/列（与 excel_reader 的 60 列上限对齐）
DEFAULT_MAX_ROWS = 2000
DEFAULT_MAX_COLS = 60

MCP_BASE = "https://docs.qq.com"
OPENAPI_BASE = "https://docs.qq.com"


# ---------------------------------------------------------------- 异常
class TencentDocsError(Exception):
    """腾讯文档相关的错误基类；`hint` 给界面直接展示的处置建议。"""

    hint = ""

    def __init__(self, message: str, hint: str = "") -> None:
        super().__init__(message)
        if hint:
            self.hint = hint


class BadLink(TencentDocsError):
    hint = "链接应形如 https://docs.qq.com/sheet/XXXXXXXX"


class NotSpreadsheet(TencentDocsError):
    hint = "目前只支持「在线表格」，其他类型（文档/幻灯片/智能表格）暂不支持"


class AuthRequired(TencentDocsError):
    hint = "请先在软件里连接腾讯文档完成授权"


class PermissionDenied(TencentDocsError):
    hint = "当前账号没有这份文档的查看权限，请联系文档所有者开通"


class RateLimited(TencentDocsError):
    hint = "调用太频繁了，等一分钟再点刷新"


class DocNotFound(TencentDocsError):
    hint = "文档不存在或已被删除，请确认链接"


class NetworkError(TencentDocsError):
    hint = "网络不通或腾讯文档服务异常，稍后再试"


class ApiError(TencentDocsError):
    pass


# ---------------------------------------------------------------- 链接解析
_LINK_RE = re.compile(
    r"docs\.qq\.com/(?:[a-z]{2}/)?(?P<kind>sheet|smartsheet|doc|slide|form|mind)/(?P<id>[A-Za-z0-9_\-$%]+)",
    re.IGNORECASE,
)
_BARE_ID_RE = re.compile(r"^[A-Za-z0-9_\-$%]{6,64}$")


@dataclass
class LinkInfo:
    """一条腾讯文档链接解析后的结果。"""

    kind: str          # sheet / smartsheet / doc / slide …
    url: str           # 规范化后的链接（去掉 query、hash）
    file_id: str       # 文档 ID（Open API 用这个）


def _decode_file_id(url_id: str) -> str:
    """URL 里的 id → file_id。

    腾讯文档的 C 端链接是 `D` + base64(file_id)，例如
    `DS29TWUFoYnNmbUxH` → `KoSYAhbsfmLG`。解不出来就原样返回。
    """
    text = url_id.strip()
    if not text:
        return ""
    if text.startswith("D") and len(text) > 8 and "$" not in text:
        body = text[1:]
        body += "=" * (-len(body) % 4)
        try:
            decoded = base64.b64decode(body).decode("utf-8")
        except Exception:  # noqa: BLE001 - 不是 base64 就按原样用
            decoded = ""
        if decoded and re.fullmatch(r"[A-Za-z0-9_\-$]{6,64}", decoded):
            return decoded
    return text


def parse_link(raw: str) -> LinkInfo:
    """校验并解析用户粘贴的链接（也接受纯 file_id）。"""
    text = (raw or "").strip().strip("<>\"'")
    if not text:
        raise BadLink("链接是空的")
    match = _LINK_RE.search(text)
    if match:
        kind = match.group("kind").lower()
        url_id = match.group("id")
        if kind != "sheet":
            raise NotSpreadsheet(f"这是一份 {kind}，不是在线表格")
        clean = f"https://docs.qq.com/sheet/{url_id}"
        return LinkInfo(kind=kind, url=clean, file_id=_decode_file_id(url_id))
    if "docs.qq.com" in text:
        raise BadLink("没能从链接里认出文档 ID，请复制浏览器地址栏里的完整链接")
    if _BARE_ID_RE.match(text):
        return LinkInfo(kind="sheet", url="", file_id=_decode_file_id(text))
    raise BadLink("不像腾讯文档链接，请复制 docs.qq.com/sheet/… 这样的地址")


# ---------------------------------------------------------------- 数据结构
@dataclass
class SheetMeta:
    """一个子表的元信息。"""

    sheet_id: str
    title: str
    rows: int = 0
    cols: int = 0
    hidden: bool = False


@dataclass
class SyncResult:
    """一次同步的结果。"""

    xlsx_path: Path
    file_id: str
    url: str
    fetched_at: float
    sheets: list[str] = field(default_factory=list)
    row_counts: dict[str, int] = field(default_factory=dict)

    @property
    def total_rows(self) -> int:
        return sum(self.row_counts.values())

    def describe(self) -> str:
        stamp = time.strftime("%H:%M:%S", time.localtime(self.fetched_at))
        sheet_text = "、".join(self.sheets[:3]) + ("…" if len(self.sheets) > 3 else "")
        return f"腾讯文档 · {len(self.sheets)} 个工作表（{sheet_text}）· {self.total_rows} 行 · {stamp} 刷新"


# ---------------------------------------------------------------- 票据
@dataclass
class McpTokens:
    """MCP 通道用的两种票据，二选一即可。"""

    oauth: str = ""
    oneid: str = ""
    api_base: str = ""
    # 个人 Token（文档列表页 → 更多操作 → 开放平台 拿到的那个）
    # 官方示例是直接放在 Authorization 里，不带 Bearer 前缀
    raw_auth: bool = False

    @property
    def present(self) -> bool:
        return bool(self.oauth or self.oneid)


class PersonalTokenProvider:
    """个人开发者 Token：只读本机设置，不需要应用凭证。

    个人主体在开放平台拿不到 Client ID，只能拿一个长期 Token
    （docs.qq.com → 更多操作 → 开放平台 → 龙虾专用入口）。
    """

    def __init__(self, token: str = "") -> None:
        self._token = (token or "").strip()
        if not self._token:
            from . import settings as prefs

            self._token = str(prefs.load().get("tdoc_token") or "").strip()

    def tokens(self) -> McpTokens:
        if not self._token:
            raise AuthRequired("还没有填写个人 Token",
                               "点「连接腾讯文档…」粘贴 Token 即可")
        return McpTokens(oauth=self._token, raw_auth=True)


class EnvTokenProvider:
    """从环境变量取票据（宿主注入时就是这条路）。"""

    def __init__(self) -> None:
        self._cache: tuple[float, McpTokens] | None = None

    def tokens(self) -> McpTokens:
        return McpTokens(
            oauth=os.environ.get("TDOC_OAUTH_ACCESS_TOKEN", ""),
            oneid=os.environ.get("TDOC_ONEID_ACCESS_TOKEN", ""),
            api_base=os.environ.get("TDOC_API_BASE_URL", ""),
        )


class HostTokenProvider:
    """向 WorkBuddy 宿主的 MCP 网关要票据（只用于本机联调，不分发）。

    宿主把 `CODEBUDDY_MCP_CONFIG` 塞进子进程环境，里面是本机 MCP 网关的地址
    和一组必需的认证头；网关的 `/internal/tencent-docs/tokens` 会把当前用户
    的腾讯文档票据发下来。票据只在内存里放几分钟，不落盘。

    打包成 .app 后从 Finder 启动是没有这个环境变量的，这时读同目录的
    `tdoc_host.json`（由 `tools/make_host_bridge.py` 生成），仅用于本机测试。
    """

    TTL = 240  # 秒

    def __init__(self, config_path: str | Path | None = None) -> None:
        self._cache: tuple[float, McpTokens] | None = None
        self.config_path = Path(config_path) if config_path else bridge_path()

    def _read_config(self) -> dict:
        """优先宿主注入的环境变量，其次本机联调文件。"""
        raw = os.environ.get("CODEBUDDY_MCP_CONFIG", "")
        if not raw:
            path = self.config_path
            if path is None or not path.exists():
                return {}
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                return {}
            return data if isinstance(data, dict) else {}
        try:
            data = json.loads(raw)
        except ValueError:
            return {}
        return data if isinstance(data, dict) else {}

    def _gateway(self) -> tuple[str, dict]:
        data = self._read_config()
        servers = data.get("mcpServers") or {}
        server = servers.get("connector-proxy") or servers.get("workbuddy") or {}
        if not server.get("url") and data.get("gateway_url"):
            server = {"url": data["gateway_url"], "headers": data.get("headers") or {}}
        gateway = str(server.get("url") or "")
        headers = {name: value
                   for name, value in (server.get("headers") or {}).items()
                   if isinstance(name, str) and isinstance(value, str) and value}
        if not gateway.endswith("/mcp") or not any(
                name.lower() == "authorization" for name in headers):
            raise AuthRequired(
                "没有可用的腾讯文档凭证",
                "请在软件里点「连接腾讯文档」完成授权；"
                "本机联调可运行 tools/make_host_bridge.py 借用 WorkBuddy 的票据",
            )
        return gateway, headers

    def tokens(self) -> McpTokens:
        if self._cache and time.time() - self._cache[0] < self.TTL:
            return self._cache[1]

        gateway, headers = self._gateway()
        request = urllib.request.Request(
            gateway + "/internal/tencent-docs/tokens", headers=headers, method="GET")
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        try:
            with opener.open(request, timeout=10) as response:
                data = json.loads(response.read().decode("utf-8"))
        except Exception as exc:  # noqa: BLE001
            raise AuthRequired(f"取票据失败：{exc}",
                               "请在软件里重新连接腾讯文档，或重跑 make_host_bridge.py") from exc

        personal = data.get("personal") or {}
        enterprise = data.get("enterprise") or {}
        tokens = McpTokens(
            oauth=str(personal.get("token") or "") if personal.get("available") else "",
            oneid=str(enterprise.get("token") or "") if enterprise.get("available") else "",
            api_base=str(data.get("apiBase") or ""),
        )
        if not tokens.present:
            reason = personal.get("reason") or enterprise.get("reason") or "token_unavailable"
            raise AuthRequired(f"宿主没有可用的腾讯文档票据（{reason}）",
                               "请在 WorkBuddy 里重新授权腾讯文档")
        self._cache = (time.time(), tokens)
        return tokens


def bridge_path() -> Path:
    """本机联调用的网关凭据文件（由 tools/make_host_bridge.py 生成，不入库）。"""
    from . import settings as prefs

    return prefs.settings_path().parent / "tdoc_host.json"


# ---------------------------------------------------------------- 通道
class Transport:
    """取数通道的统一接口。"""

    def sheets(self, file_id: str) -> list[SheetMeta]:
        raise NotImplementedError

    def read_range(self, file_id: str, sheet_id: str,
                   start_row: int, start_col: int,
                   end_row: int, end_col: int) -> dict[tuple[int, int], Any]:
        """读一块区域，返回 `{(行, 列): 值}`，行列都是 0 基的绝对坐标。"""
        raise NotImplementedError


def _cell_value(cell: dict) -> Any:
    """把 MCP 的结构化单元格简化成 Python 值。"""
    value_type = str(cell.get("value_type") or "").upper()
    if value_type == "NUMBER":
        number = cell.get("number_value")
        try:
            return float(number)
        except (TypeError, ValueError):
            return cell.get("string_value") or ""
    if value_type == "BOOL":
        return bool(cell.get("bool_value"))
    text = cell.get("string_value")
    if text in (None, ""):
        # 有些单元格只有公式，没有字符串值
        text = cell.get("formula") or ""
    return str(text)


class McpTransport(Transport):
    """走 MCP 端点：一份 JSON-RPC 请求读一块区域。"""

    PATH = "/api/v6/sheet/mcp"

    def __init__(self, provider: EnvTokenProvider | HostTokenProvider,
                 base: str = MCP_BASE, timeout: int = 30) -> None:
        self.provider = provider
        self.base = (base or MCP_BASE).rstrip("/")
        self.timeout = timeout
        self._request_id = 0

    # -- 底层 ---------------------------------------------------------
    def _url(self) -> str:
        tokens = self.provider.tokens()
        base = (tokens.api_base or self.base).rstrip("/")
        return base + self.PATH

    def _call(self, tool: str, arguments: dict) -> dict:
        tokens = self.provider.tokens()
        self._request_id += 1
        payload = {
            "jsonrpc": "2.0",
            "id": self._request_id,
            "method": "tools/call",
            "params": {"name": tool, "arguments": arguments},
        }
        # Authorization 的写法有两种流派：宿主的 OAuth 票是 `Bearer xxx`，
        # 个人 Token 在官方示例里是直接放 token。先按提供方的约定来，
        # 401 再换另一种试一次（只多一次请求，能自愈）。
        styles = [tokens.raw_auth]
        if tokens.oauth:
            styles.append(not tokens.raw_auth)
        last_error: TencentDocsError | None = None
        for raw in styles:
            try:
                return self._post(payload, tokens, raw)
            except AuthRequired as exc:
                last_error = exc
                continue
        raise last_error or AuthRequired("授权无效")

    def _post(self, payload: dict, tokens: McpTokens, raw: bool) -> dict:
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
            "User-Agent": "AwardStudio",
        }
        if tokens.oauth:
            headers["Authorization"] = tokens.oauth if raw else f"Bearer {tokens.oauth}"
        if tokens.oneid:
            headers["X-Oneid-Access-Token"] = tokens.oneid

        request = urllib.request.Request(
            self._url(), data=json.dumps(payload).encode("utf-8"),
            headers=headers, method="POST")
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                raw_text = response.read().decode("utf-8", errors="replace")
                content_type = response.headers.get("Content-Type", "")
        except urllib.error.HTTPError as exc:
            body = ""
            try:
                body = exc.read().decode("utf-8", errors="replace")[:300]
            except Exception:  # noqa: BLE001
                pass
            self._raise_http(exc.code, body)
        except urllib.error.URLError as exc:
            raise NetworkError(f"连不上腾讯文档：{exc.reason}") from exc
        except OSError as exc:
            raise NetworkError(f"连不上腾讯文档：{exc}") from exc

        if "text/event-stream" in content_type:
            raw_text = _extract_sse(raw_text)
        return _parse_rpc(raw_text)

    @staticmethod
    def _raise_http(status: int, body: str) -> None:
        if status in (401, 403):
            raise AuthRequired(f"票据无效（HTTP {status}）")
        if status == 429:
            raise RateLimited("请求太频繁（HTTP 429）")
        if status >= 500:
            raise NetworkError(f"腾讯文档服务异常（HTTP {status}）")
        raise ApiError(f"请求被拒绝（HTTP {status}）{body}")

    # -- 接口 ---------------------------------------------------------
    def sheets(self, file_id: str) -> list[SheetMeta]:
        data = self._call("get_sheet_info", {"file_id": file_id})
        result: list[SheetMeta] = []
        for item in data.get("sheets") or []:
            result.append(SheetMeta(
                sheet_id=str(item.get("sheet_id") or ""),
                title=str(item.get("sheet_name") or ""),
                rows=int(item.get("row_count") or 0),
                cols=int(item.get("col_count") or 0),
                hidden=bool(item.get("hidden")),
            ))
        return [sheet for sheet in result if sheet.sheet_id]

    def read_range(self, file_id: str, sheet_id: str,
                   start_row: int, start_col: int,
                   end_row: int, end_col: int) -> dict[tuple[int, int], Any]:
        data = self._call("get_cell_data", {
            "file_id": file_id,
            "sheet_id": sheet_id,
            "start_row": start_row,
            "start_col": start_col,
            "end_row": end_row,
            "end_col": end_col,
            "return_csv": False,
        })
        cells: dict[tuple[int, int], Any] = {}
        for cell in data.get("cells") or []:
            try:
                key = (int(cell.get("row")), int(cell.get("col")))
            except (TypeError, ValueError):
                continue
            value = _cell_value(cell)
            if value not in (None, ""):
                cells[key] = value
        return cells


class OpenApiTransport(Transport):
    """走 Open API：正式的 OAuth 通道，请求头带 Client-Id / Open-Id / Access-Token。

    文档：
      * 子表信息 GET /openapi/spreadsheet/v3/files/{fileId}
      * 区域取值 GET /openapi/spreadsheet/v2/files/{fileId}/sheets/{sheetId}/values?range=A1:D10
    """

    def __init__(self, client_id: str, open_id: str, access_token: str,
                 base: str = OPENAPI_BASE, timeout: int = 30) -> None:
        self.client_id = client_id
        self.open_id = open_id
        self.access_token = access_token
        self.base = (base or OPENAPI_BASE).rstrip("/")
        self.timeout = timeout

    def _get(self, path: str, params: dict | None = None) -> dict:
        url = self.base + path
        if params:
            url += ("&" if "?" in url else "?") + urllib.parse.urlencode(params)
        request = urllib.request.Request(url, method="GET", headers={
            "Accept": "application/json",
            "User-Agent": "AwardStudio",
            "Access-Token": self.access_token,
            "Client-Id": self.client_id,
            "Open-Id": self.open_id,
        })
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                raw = response.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as exc:
            body = ""
            try:
                body = exc.read().decode("utf-8", errors="replace")[:300]
            except Exception:  # noqa: BLE001
                pass
            self._raise(exc.code, body)
        except urllib.error.URLError as exc:
            raise NetworkError(f"连不上腾讯文档：{exc.reason}") from exc
        except OSError as exc:
            raise NetworkError(f"连不上腾讯文档：{exc}") from exc
        try:
            return json.loads(raw)
        except ValueError as exc:
            raise ApiError(f"返回内容无法解析：{raw[:120]}") from exc

    @staticmethod
    def _raise(status: int, body: str) -> None:
        code = 0
        try:
            code = int(json.loads(body).get("code") or 0)
        except Exception:  # noqa: BLE001
            pass
        if code == 400006 or status in (401, 403):
            raise AuthRequired("授权已过期或未授权", "请重新连接腾讯文档")
        if code == 400007 or status == 429:
            raise RateLimited("请求太频繁，稍后再试")
        if code in (400001, 10313):
            raise ApiError(f"请求参数被拒绝：{body}")
        if status >= 500:
            raise NetworkError(f"腾讯文档服务异常（HTTP {status}）")
        raise ApiError(f"请求失败（HTTP {status}）{body}")

    def sheets(self, file_id: str) -> list[SheetMeta]:
        data = self._get(f"/openapi/spreadsheet/v3/files/{urllib.parse.quote(file_id)}",
                         {"concise": 1})
        if data.get("code") not in (None, 0):
            raise ApiError(f"读取子表失败：{data.get('message') or data.get('code')}")
        result: list[SheetMeta] = []
        for item in data.get("properties") or []:
            result.append(SheetMeta(
                sheet_id=str(item.get("sheetId") or item.get("sheetID") or ""),
                title=str(item.get("title") or ""),
                rows=int(item.get("rowCount") or 0),
                cols=int(item.get("columnCount") or 0),
            ))
        return [sheet for sheet in result if sheet.sheet_id]

    def read_range(self, file_id: str, sheet_id: str,
                   start_row: int, start_col: int,
                   end_row: int, end_col: int) -> dict[tuple[int, int], Any]:
        cell_range = f"{_a1(start_row, start_col)}:{_a1(max(start_row, end_row - 1), max(start_col, end_col - 1))}"
        data = self._get(
            f"/openapi/spreadsheet/v2/files/{urllib.parse.quote(file_id)}"
            f"/sheets/{urllib.parse.quote(sheet_id)}/values",
            {"range": cell_range})
        if data.get("code") not in (None, 0):
            raise ApiError(f"读取单元格失败：{data.get('message') or data.get('code')}")
        values = ((data.get("data") or {}).get("values")
                  or data.get("values") or [])
        cells: dict[tuple[int, int], Any] = {}
        for row_offset, row in enumerate(values):
            if not isinstance(row, list):
                continue
            for col_offset, cell in enumerate(row):
                value = cell
                if isinstance(cell, dict):
                    value = (cell.get("stringValue") or cell.get("string_value")
                             or cell.get("numberValue") or cell.get("number_value"))
                    if isinstance(value, str) and value == "":
                        continue
                if value in (None, ""):
                    continue
                if isinstance(value, str):
                    try:
                        value = float(value)
                    except ValueError:
                        pass
                cells[(start_row + row_offset, start_col + col_offset)] = value
        return cells


def _a1(row: int, col: int) -> str:
    """0 基行列 → A1 记法（列号可能超过 26）。"""
    letters = ""
    number = col + 1
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return f"{letters}{row + 1}"


def _extract_sse(raw: str) -> str:
    """从 SSE 报文里抽出最后一条能解析的 data 负载。"""
    lines = [line[len("data:"):].strip() for line in raw.splitlines()
             if line.startswith("data:")]
    for chunk in reversed(lines):
        try:
            json.loads(chunk)
            return chunk
        except ValueError:
            continue
    return lines[-1] if lines else raw


def _parse_rpc(raw: str) -> dict:
    """解析 JSON-RPC 响应，返回 MCP 工具的结果字典。"""
    try:
        envelope = json.loads(raw)
    except ValueError as exc:
        raise ApiError(f"返回内容无法解析：{raw[:120]}") from exc

    error = envelope.get("error")
    if error:
        message = str(error.get("message") or error)
        raise _error_from_message(message)

    result = envelope.get("result") or {}
    if result.get("isError"):
        texts = [item.get("text", "") for item in result.get("content") or []
                 if isinstance(item, dict)]
        raise _error_from_message(" ".join(texts) or "工具调用失败")

    structured = result.get("structuredContent")
    if isinstance(structured, dict) and structured:
        return structured
    for item in result.get("content") or []:
        if isinstance(item, dict) and item.get("type") == "text":
            try:
                parsed = json.loads(item.get("text") or "{}")
            except ValueError:
                continue
            if isinstance(parsed, dict):
                return parsed
    return {}


def _error_from_message(message: str) -> TencentDocsError:
    """把腾讯文档的业务错误码翻译成有处置建议的异常。"""
    text = message or ""
    if "400006" in text or "未授权" in text or "鉴权失败" in text:
        return AuthRequired(f"授权无效：{_short(text)}")
    if "400007" in text or "VIP" in text:
        return PermissionDenied(f"需要更高权限：{_short(text)}", "该文档可能需要文档所有者开通对应权限")
    if "400016" in text:
        return NotSpreadsheet(f"文档类型不匹配：{_short(text)}")
    if "限频" in text or "429" in text:
        return RateLimited(f"请求太频繁：{_short(text)}")
    if any(token in text for token in ("not found", "不存在", "无权限", "not exist")):
        return DocNotFound(f"读不到这份文档：{_short(text)}")
    return ApiError(f"腾讯文档返回错误：{_short(text)}")


def _short(text: str, limit: int = 160) -> str:
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


# ---------------------------------------------------------------- 读取与落盘
class TencentSheet:
    """一份在线表格：列子表、按页读满、转成二维数组。"""

    PAGE_ROWS = 200

    def __init__(self, transport: Transport, link: LinkInfo) -> None:
        self.transport = transport
        self.link = link

    def sheet_list(self) -> list[SheetMeta]:
        sheets = self.transport.sheets(self.link.file_id)
        visible = [sheet for sheet in sheets if not sheet.hidden]
        return visible or sheets

    def read_sheet(self, sheet: SheetMeta,
                   max_rows: int = DEFAULT_MAX_ROWS,
                   max_cols: int = DEFAULT_MAX_COLS) -> list[list[Any]]:
        """把一个子表读成稠密二维数组（尾部空行/空列会去掉）。

        `sheet.rows` 是表格的网格行数（新建的空表也报 200），比实际数据多，
        所以拿它当上限既安全又省请求；再配合「这一页没填满就不再往下读」，
        正常情况下只多花一次请求就能确认到底。
        """
        ceiling = min(max_rows, sheet.rows) if sheet.rows > 0 else max_rows
        total_cols = min(max_cols, sheet.cols) if sheet.cols else max_cols
        total_cols = max(1, total_cols)
        page = max(1, min(self.PAGE_ROWS, MAX_CELLS_PER_REQUEST // total_cols))

        cells: dict[tuple[int, int], Any] = {}
        row = 0
        while row < ceiling:
            end_row = min(row + page, ceiling)
            block = self.transport.read_range(
                self.link.file_id, sheet.sheet_id, row, 0, end_row, total_cols)
            if not block:
                break  # 整页空白：后面不会再有数据了
            cells.update(block)
            last_row = max(r for r, _c in block)
            row = end_row
            if last_row < end_row - 1:
                break  # 这一页最后一行是空的，说明数据到这为止

        if not cells:
            return []
        last_row = max(r for r, _c in cells)
        last_col = max(c for _r, c in cells)
        grid = [[None] * (last_col + 1) for _ in range(last_row + 1)]
        for (r, c), value in cells.items():
            grid[r][c] = value
        # 去掉尾部整行空
        while grid and all(v in (None, "") for v in grid[-1]):
            grid.pop()
        return grid


def _safe_sheet_title(title: str, used: set[str]) -> str:
    """Excel 工作表名限制：≤31 字符、不能含 []:*?/\\ 。"""
    text = re.sub(r"[\[\]:*?/\\]", "_", str(title or "工作表")).strip() or "工作表"
    text = text[:31]
    candidate, index = text, 2
    while candidate in used:
        suffix = f"({index})"
        candidate = text[: 31 - len(suffix)] + suffix
        index += 1
    used.add(candidate)
    return candidate


def write_xlsx(path: str | Path, sheets: list[tuple[str, list[list[Any]]]]) -> Path:
    """把若干 (标题, 二维数组) 写成一个 xlsx。数字保持数字类型。"""
    from openpyxl import Workbook

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    used: set[str] = set()
    first = True
    for title, rows in sheets:
        name = _safe_sheet_title(title, used)
        sheet = book.active if first else book.create_sheet()
        first = False
        sheet.title = name
        for row_index, row in enumerate(rows, start=1):
            for col_index, value in enumerate(row, start=1):
                if value in (None, ""):
                    continue
                sheet.cell(row=row_index, column=col_index, value=value)
    if first:  # 一个子表都没写，仍然要留一张空表
        book.active.title = "工作表1"
    temp = target.with_suffix(".tmp.xlsx")
    book.save(str(temp))
    os.replace(temp, target)
    return target


def cache_paths(cache_dir: str | Path, file_id: str) -> tuple[Path, Path]:
    """缓存文件的固定位置：`<file_id>.xlsx` 与同名 `.json` 元信息。"""
    safe = re.sub(r"[^A-Za-z0-9_\-$]", "_", file_id) or "tencent"
    root = Path(cache_dir)
    return root / f"{safe}.xlsx", root / f"{safe}.json"


def default_cache_dir() -> Path:
    """缓存目录：Windows 用 %LOCALAPPDATA%，其余用 ~/Library/Caches 或 ~/.cache。"""
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or "."
        return Path(base) / "AwardStudio" / "tencent"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Caches" / "AwardStudio" / "tencent"
    return Path.home() / ".cache" / "AwardStudio" / "tencent"


def sync_to_cache(url: str, cache_dir: str | Path | None = None,
                  transport: Transport | None = None) -> SyncResult:
    """把一份在线表格同步成本地 xlsx，返回结果（含行列统计与刷新时间）。

    `transport` 传 None 时按「先 OAuth、后宿主票据」的顺序自动选：
    正式分发只有 OAuth；本机开发时宿主票据兜底。
    """
    link = parse_link(url)
    active = transport or build_transport()
    sheet_reader = TencentSheet(active, link)

    metas = sheet_reader.sheet_list()
    if not metas:
        raise DocNotFound("这份表格里没有可读的工作表")

    collected: list[tuple[str, list[list[Any]]]] = []
    row_counts: dict[str, int] = {}
    for meta in metas:
        rows = sheet_reader.read_sheet(meta)
        collected.append((meta.title or meta.sheet_id, rows))
        row_counts[meta.title or meta.sheet_id] = len(rows)

    if not any(rows for _title, rows in collected):
        raise DocNotFound("表格是空的（没有读到任何单元格）", "确认链接指向的是成绩表，且表格里已有数据")

    target = cache_dir or default_cache_dir()
    xlsx_path, meta_path = cache_paths(target, link.file_id)
    write_xlsx(xlsx_path, collected)

    result = SyncResult(
        xlsx_path=xlsx_path,
        file_id=link.file_id,
        url=link.url or url,
        fetched_at=time.time(),
        sheets=[title for title, _rows in collected],
        row_counts=row_counts,
    )
    try:
        meta_path.write_text(json.dumps({
            "url": result.url,
            "file_id": result.file_id,
            "fetched_at": result.fetched_at,
            "sheets": result.sheets,
            "row_counts": result.row_counts,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass  # 元信息写不进去不影响主流程
    return result


def build_transport() -> Transport:
    """按优先级选通道：应用自己的 OAuth → 宿主票据 → 环境变量。

    正式分发的软件只有第一条；后两条是给本机开发和自测用的，
    所以在「已配置凭证但还没授权」时也会先试一下宿主票据，方便调试。
    """
    from . import tdoc_auth

    credentials = tdoc_auth.AppCredentials.load()
    store = tdoc_auth.TokenStore()
    tokens = store.load()
    if credentials.configured and tokens.access_token:
        fresh = store.ensure_fresh(credentials, tokens)
        return OpenApiTransport(
            client_id=credentials.client_id,
            open_id=fresh.open_id,
            access_token=fresh.access_token,
        )

    # 个人开发者：只有 Token，没有 Client ID（个人主体拿不到第三方应用）
    personal = PersonalTokenProvider()
    try:
        if personal.tokens().present:
            return McpTransport(personal)
    except AuthRequired:
        pass

    host = HostTokenProvider()
    try:
        if host.tokens().present:
            return McpTransport(host)
    except AuthRequired:
        pass
    env = EnvTokenProvider()
    if env.tokens().present:
        return McpTransport(env)

    if credentials.configured:
        raise AuthRequired("还没有授权腾讯文档", "请点「连接腾讯文档」用你的账号扫码授权")
    raise AuthRequired("没有可用的腾讯文档凭证",
                       "请点「连接腾讯文档」完成授权；"
                       "若软件未内置应用凭证，需要先在设置里填 Client ID / Secret；"
                       "本机联调可运行 tools/make_host_bridge.py")
