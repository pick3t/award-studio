"""腾讯文档 OAuth 授权（自建应用，可分发给别人用）。

流程：软件内置本应用的 Client ID / Secret → 用户点「连接腾讯文档」→ 浏览器打开
腾讯的授权页 → 用户用自己的账号同意 → 腾讯带 `code` 跳到回调地址 → 软件换
Access Token（30 天）+ Refresh Token（1 年）→ 存本地，之后用 Open API 读表。

腾讯要求回调地址**必须是 HTTPS**，所以这里给了两种收码方式：

1. `LocalCallbackServer`：软件内起一个 `https://127.0.0.1:端口/callback`（自签
   证书，首次运行用 openssl 生成）。体验最好，但要确认开放平台是否接受这种
   回调地址。
2. 粘贴模式：把 `code`（或整条回调链接）粘回软件。用固定的 HTTPS 静态页做
   回调时走这条，零服务器也能跑。

凭证与令牌都不进仓库：凭证放设置文件或环境变量，令牌存设置文件。
"""

from __future__ import annotations

import json
import os
import re
import ssl
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, replace
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from . import settings as prefs

AUTHORIZE_URL = "https://docs.qq.com/oauth/v2/authorize"
TOKEN_URL = "https://docs.qq.com/oauth/v2/token"
DEFAULT_SCOPE = "all"
DEFAULT_REDIRECT = "https://127.0.0.1:17801/callback"
CALLBACK_PORT = 17801

# 令牌提前 3 天续期，避免"正好在裁判会场上过期"
REFRESH_AHEAD = 3 * 24 * 3600


# ---------------------------------------------------------------- 应用凭证
@dataclass
class AppCredentials:
    """本应用在腾讯文档开放平台登记的身份。"""

    client_id: str = ""
    client_secret: str = ""
    redirect_uri: str = DEFAULT_REDIRECT
    scope: str = DEFAULT_SCOPE

    @property
    def configured(self) -> bool:
        return bool(self.client_id and self.client_secret)

    @property
    def https_redirect(self) -> bool:
        return self.redirect_uri.lower().startswith("https://")

    @classmethod
    def load(cls) -> "AppCredentials":
        """优先环境变量，其次设置文件（便于 CI 或临时覆盖）。"""
        saved = prefs.load()
        return cls(
            client_id=os.environ.get("AWARD_TDOC_CLIENT_ID")
            or str(saved.get("tdoc_client_id") or ""),
            client_secret=os.environ.get("AWARD_TDOC_CLIENT_SECRET")
            or str(saved.get("tdoc_client_secret") or ""),
            redirect_uri=os.environ.get("AWARD_TDOC_REDIRECT_URI")
            or str(saved.get("tdoc_redirect_uri") or DEFAULT_REDIRECT),
            scope=os.environ.get("AWARD_TDOC_SCOPE")
            or str(saved.get("tdoc_scope") or DEFAULT_SCOPE),
        )

    @classmethod
    def save(cls, client_id: str, client_secret: str,
             redirect_uri: str = DEFAULT_REDIRECT, scope: str = DEFAULT_SCOPE) -> None:
        data = prefs.load()
        data.update({
            "tdoc_client_id": client_id.strip(),
            "tdoc_client_secret": client_secret.strip(),
            "tdoc_redirect_uri": redirect_uri.strip() or DEFAULT_REDIRECT,
            "tdoc_scope": scope.strip() or DEFAULT_SCOPE,
        })
        prefs.save(data)


# ---------------------------------------------------------------- 令牌
@dataclass
class Tokens:
    """一次授权拿到的令牌集合。"""

    access_token: str = ""
    refresh_token: str = ""
    open_id: str = ""
    scope: str = ""
    expires_at: float = 0.0

    @property
    def present(self) -> bool:
        return bool(self.access_token)

    @property
    def expires_in(self) -> float:
        return self.expires_at - time.time() if self.expires_at else 0.0

    @property
    def expired(self) -> bool:
        return self.expires_at > 0 and self.expires_in <= 0

    @property
    def needs_refresh(self) -> bool:
        return self.expires_at > 0 and self.expires_in < REFRESH_AHEAD

    def describe(self) -> str:
        if not self.present:
            return "未连接"
        if self.expired:
            return "授权已过期，请重新连接"
        days = max(0, int(self.expires_in // 86400))
        return f"已连接（授权剩余 {days} 天）"

    def to_dict(self) -> dict:
        return {
            "access_token": self.access_token,
            "refresh_token": self.refresh_token,
            "open_id": self.open_id,
            "scope": self.scope,
            "expires_at": self.expires_at,
        }

    @classmethod
    def from_dict(cls, data: dict | None) -> "Tokens":
        data = data or {}
        try:
            expires_at = float(data.get("expires_at") or 0)
        except (TypeError, ValueError):
            expires_at = 0.0
        return cls(
            access_token=str(data.get("access_token") or ""),
            refresh_token=str(data.get("refresh_token") or ""),
            open_id=str(data.get("open_id") or ""),
            scope=str(data.get("scope") or ""),
            expires_at=expires_at,
        )


class TokenStore:
    """令牌存在设置文件里（和字号、奖项方案同一份）。"""

    KEY = "tdoc_tokens"

    def load(self) -> Tokens:
        return Tokens.from_dict(prefs.load().get(self.KEY))

    def save(self, tokens: Tokens) -> None:
        data = prefs.load()
        data[self.KEY] = tokens.to_dict()
        prefs.save(data)

    def clear(self) -> None:
        data = prefs.load()
        data.pop(self.KEY, None)
        prefs.save(data)

    def ensure_fresh(self, credentials: AppCredentials, tokens: Tokens) -> Tokens:
        """快过期就自动续期；续不上就原样返回，让调用方报"需要重新授权"。"""
        if not tokens.present or not tokens.needs_refresh:
            return tokens
        if not tokens.refresh_token or not credentials.configured:
            return replace(tokens, expires_at=0.0) if tokens.expired else tokens
        try:
            fresh = refresh_access_token(credentials, tokens.refresh_token)
        except TencentAuthError:
            return replace(tokens, expires_at=0.0) if tokens.expired else tokens
        merged = replace(
            tokens,
            access_token=fresh.access_token or tokens.access_token,
            refresh_token=fresh.refresh_token or tokens.refresh_token,
            scope=fresh.scope or tokens.scope,
            open_id=fresh.open_id or tokens.open_id,
            expires_at=fresh.expires_at,
        )
        self.save(merged)
        return merged


# ---------------------------------------------------------------- 错误
class TencentAuthError(Exception):
    """授权过程中的错误。"""

    def __init__(self, message: str, hint: str = "") -> None:
        super().__init__(message)
        self.hint = hint


# ---------------------------------------------------------------- 接口调用
def _request_json(url: str, timeout: int = 20) -> dict:
    request = urllib.request.Request(url, headers={
        "Accept": "application/json",
        "User-Agent": "AwardStudio",
    })
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        body = ""
        try:
            body = exc.read().decode("utf-8", errors="replace")[:200]
        except Exception:  # noqa: BLE001
            pass
        if exc.code in (401, 403):
            raise TencentAuthError(
                "授权码无效或已过期（5 分钟内有效且只能用一次）",
                "请重新点「连接腾讯文档」再授权一次",
            ) from exc
        raise TencentAuthError(f"腾讯文档返回 HTTP {exc.code}：{body}") from exc
    except urllib.error.URLError as exc:
        raise TencentAuthError(f"连不上腾讯文档：{exc.reason}", "检查网络后重试") from exc
    except OSError as exc:
        raise TencentAuthError(f"连不上腾讯文档：{exc}") from exc
    try:
        return json.loads(raw)
    except ValueError as exc:
        raise TencentAuthError(f"返回内容无法解析：{raw[:120]}") from exc


def _tokens_from_payload(payload: dict) -> Tokens:
    if payload.get("error") or payload.get("error_code"):
        message = payload.get("error_description") or payload.get("error") or payload.get("msg")
        raise TencentAuthError(f"授权失败：{message}")
    access_token = str(payload.get("access_token") or "")
    if not access_token:
        raise TencentAuthError(f"授权失败：{payload.get('msg') or payload}")
    try:
        expires_in = int(payload.get("expires_in") or 2592000)
    except (TypeError, ValueError):
        expires_in = 2592000
    return Tokens(
        access_token=access_token,
        refresh_token=str(payload.get("refresh_token") or ""),
        open_id=str(payload.get("user_id") or payload.get("open_id") or ""),
        scope=str(payload.get("scope") or ""),
        expires_at=time.time() + expires_in,
    )


def authorize_url(credentials: AppCredentials, state: str,
                  force_confirm: bool = False) -> str:
    """拼出授权页地址（用户在浏览器里打开/扫码）。"""
    if not credentials.configured:
        raise TencentAuthError("还没有填写应用凭证", "请先在设置里填入 Client ID / Secret")
    params = {
        "client_id": credentials.client_id,
        "redirect_uri": credentials.redirect_uri,
        "response_type": "code",
        "scope": credentials.scope or DEFAULT_SCOPE,
        "state": state,
    }
    if force_confirm:
        params["force_login"] = "1"
    return AUTHORIZE_URL + "?" + urllib.parse.urlencode(params)


def exchange_code(credentials: AppCredentials, code: str) -> Tokens:
    """用授权码换令牌。"""
    params = {
        "client_id": credentials.client_id,
        "client_secret": credentials.client_secret,
        "redirect_uri": credentials.redirect_uri,
        "grant_type": "authorization_code",
        "code": code,
    }
    payload = _request_json(TOKEN_URL + "?" + urllib.parse.urlencode(params))
    return _tokens_from_payload(payload)


def refresh_access_token(credentials: AppCredentials, refresh_token: str) -> Tokens:
    """用 Refresh Token 续期（Refresh Token 有效期 1 年）。"""
    params = {
        "client_id": credentials.client_id,
        "client_secret": credentials.client_secret,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }
    payload = _request_json(TOKEN_URL + "?" + urllib.parse.urlencode(params))
    return _tokens_from_payload(payload)


# ---------------------------------------------------------------- 收码
_CODE_RE = re.compile(r"(?:^|[?&#])code=([^&#\s]+)")
_STATE_RE = re.compile(r"(?:^|[?&#])state=([^&#\s]+)")
_ERROR_RE = re.compile(r"(?:^|[?&#])error=([^&#\s]+)")


def parse_callback(raw: str) -> tuple[str, str]:
    """从回调链接或裸 code 里取出 `(code, state)`。"""
    text = (raw or "").strip()
    if not text:
        raise TencentAuthError("粘贴的内容是空的", "请把浏览器地址栏里的整条回调链接复制过来")
    error = _ERROR_RE.search(text)
    if error:
        raise TencentAuthError(f"授权被拒绝：{error.group(1)}", "重新发起一次授权")
    code_match = _CODE_RE.search(text)
    if code_match:
        state_match = _STATE_RE.search(text)
        return urllib.parse.unquote(code_match.group(1)), urllib.parse.unquote(
            state_match.group(1)) if state_match else ""
    # 用户可能只复制了 code 本身
    if re.fullmatch(r"[A-Za-z0-9_\-]{8,}", text):
        return text, ""
    raise TencentAuthError(
        "没能从粘贴内容里找到 code",
        "请复制整条回调链接（含 code=… 的那一截）",
    )


def cert_paths(folder: str | Path) -> tuple[Path, Path]:
    root = Path(folder)
    return root / "callback_cert.pem", root / "callback_key.pem"


def ensure_self_signed_cert(folder: str | Path) -> tuple[Path, Path]:
    """生成（或复用）127.0.0.1 的自签证书。没有 openssl 就抛错让调用方走粘贴模式。"""
    cert, key = cert_paths(folder)
    if cert.exists() and key.exists():
        return cert, key
    cert.parent.mkdir(parents=True, exist_ok=True)
    command = [
        "openssl", "req", "-x509", "-newkey", "rsa:2048", "-nodes",
        "-keyout", str(key), "-out", str(cert), "-days", "3650",
        "-subj", "/CN=127.0.0.1",
        "-addext", "subjectAltName=IP:127.0.0.1,DNS:localhost",
    ]
    try:
        subprocess.run(command, check=True, capture_output=True, timeout=30)
    except FileNotFoundError as exc:
        raise TencentAuthError(
            "本机没有 openssl，无法启用本地回调",
            "改用「粘贴授权码」方式，或安装 openssl 后重试",
        ) from exc
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or b"").decode("utf-8", errors="replace")[:160]
        raise TencentAuthError(f"生成自签证书失败：{detail}") from exc
    return cert, key


class LocalCallbackServer:
    """在本机起一个 HTTPS 端点接收授权码（自签证书）。

    用法：
        server = LocalCallbackServer(cert_dir)
        server.start()
        webbrowser.open(authorize_url(...))
        code, state = server.wait(120)
    """

    def __init__(self, cert_dir: str | Path, port: int = CALLBACK_PORT) -> None:
        self.cert_dir = Path(cert_dir)
        self.port = port
        self._server: ThreadingHTTPServer | None = None
        self._thread: threading.Thread | None = None
        self.result: tuple[str, str] | None = None
        self.error: str = ""
        self._done = threading.Event()

    def start(self) -> str:
        cert, key = ensure_self_signed_cert(self.cert_dir)
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(certfile=str(cert), keyfile=str(key))

        outer = self

        class Handler(BaseHTTPRequestHandler):
            protocol_version = "HTTP/1.1"

            def log_message(self, *_args) -> None:  # 静音
                return

            def do_GET(self) -> None:  # noqa: N802 - BaseHTTPRequestHandler 约定
                parsed = urllib.parse.urlparse(self.path)
                if parsed.path.rstrip("/") not in ("/callback", ""):
                    self._respond(404, "找不到页面")
                    return
                query = urllib.parse.parse_qs(parsed.query)
                if query.get("error"):
                    outer.error = str(query["error"][0])
                    self._respond(400, "授权失败，请回到软件重试。")
                elif query.get("code"):
                    outer.result = (str(query["code"][0]),
                                    str(query.get("state", [""])[0]))
                    self._respond(200, "授权成功，可以关闭这个页面回到 Award Studio 了。")
                else:
                    outer.error = "回调里没有 code"
                    self._respond(400, "回调参数不完整。")
                outer._done.set()

            def _respond(self, status: int, text: str) -> None:
                body = (
                    "<!doctype html><meta charset='utf-8'>"
                    "<style>body{font:16px/1.7 -apple-system,'PingFang SC',sans-serif;"
                    "display:flex;align-items:center;justify-content:center;height:100vh;"
                    "margin:0;background:#F3F5FA;color:#111827}"
                    "div{background:#fff;padding:32px 40px;border-radius:16px;"
                    "box-shadow:0 8px 32px rgba(17,24,39,.08);max-width:420px;text-align:center}"
                    "</style><div>" + text + "</div>"
                ).encode("utf-8")
                self.send_response(status)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

        server = ThreadingHTTPServer(("127.0.0.1", self.port), Handler)
        server.socket = context.wrap_socket(server.socket, server_side=True)
        self._server = server
        self._thread = threading.Thread(target=server.serve_forever, daemon=True)
        self._thread.start()
        return f"https://127.0.0.1:{self.port}/callback"

    def wait(self, timeout: float = 180.0) -> tuple[str, str]:
        ok = self._done.wait(timeout)
        self.stop()
        if not ok:
            raise TencentAuthError("等待授权超时", "请重新点「连接腾讯文档」")
        if self.error and not self.result:
            raise TencentAuthError(f"授权失败：{self.error}")
        if not self.result:
            raise TencentAuthError("没有收到授权码")
        return self.result

    def stop(self) -> None:
        if self._server is not None:
            try:
                self._server.shutdown()
                self._server.server_close()
            except Exception:  # noqa: BLE001
                pass
            self._server = None


def open_in_browser(url: str) -> None:
    """用系统浏览器打开授权页。"""
    import webbrowser

    try:
        webbrowser.open(url, new=1)
    except Exception:  # noqa: BLE001
        if sys.platform == "darwin":
            subprocess.Popen(["open", url])
        elif sys.platform == "win32":
            os.startfile(url)  # noqa: S606
        else:
            subprocess.Popen(["xdg-open", url])
