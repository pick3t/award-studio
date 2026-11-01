"""自测：腾讯文档 OAuth 里不依赖真实凭证的部分。

    python tools/check_tdoc_auth.py

覆盖：
1. 授权链接的参数拼装（client_id / redirect_uri / scope / state）
2. 回调解析（整条链接、只有 code、出错、垃圾输入）
3. 本地自签 HTTPS 回调服务真的能收码（起服务 + 跳过证书校验去请求）
4. 令牌存取、过期判断、续期触发时机

真正「拿 code 换 token」需要开放平台的应用凭证，不在自测范围内。
"""

from __future__ import annotations

import ssl
import sys
import tempfile
import threading
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core import tdoc_auth  # noqa: E402

failures: list[str] = []


def check(ok: bool, message: str) -> None:
    print(("  [OK] " if ok else "  [!!] ") + message)
    if not ok:
        failures.append(message)


def test_authorize_url() -> None:
    print("1) 授权链接")
    credentials = tdoc_auth.AppCredentials(
        client_id="CID123", client_secret="SEC",
        redirect_uri="https://127.0.0.1:17801/callback")
    url = tdoc_auth.authorize_url(credentials, "STATE456")
    check(url.startswith(tdoc_auth.AUTHORIZE_URL + "?"), f"指向授权页：{url[:56]}…")
    check("client_id=CID123" in url and "state=STATE456" in url, "带上 client_id 与 state")
    check("response_type=code" in url, "使用授权码模式")
    check("redirect_uri=https%3A%2F%2F127.0.0.1%3A17801%2Fcallback" in url,
          "回调地址正确编码")
    try:
        tdoc_auth.authorize_url(tdoc_auth.AppCredentials(), "x")
    except tdoc_auth.TencentAuthError:
        check(True, "没配凭证时拒绝发起授权")
    else:
        check(False, "没配凭证本应报错")


def test_parse_callback() -> None:
    print("2) 回调解析")
    for raw, expected in (
        ("https://127.0.0.1:17801/callback?code=ABC123&state=S1", ("ABC123", "S1")),
        ("https://example.com/cb?state=S1&code=ABC123", ("ABC123", "S1")),
        ("code=ABC123", ("ABC123", "")),
        ("ABC123def", ("ABC123def", "")),
    ):
        got = tdoc_auth.parse_callback(raw)
        check(got == expected, f"{raw[:40]} → {got}")
    for raw in ("", "https://x/callback?error=access_denied", "乱七八糟"):
        try:
            tdoc_auth.parse_callback(raw)
        except tdoc_auth.TencentAuthError:
            check(True, f"「{raw or '空'}」被正确拒绝")
        else:
            check(False, f"「{raw}」本应报错")


def test_local_callback(work: Path) -> None:
    print("3) 本地自签 HTTPS 回调")
    server = tdoc_auth.LocalCallbackServer(work, port=17819)
    callback = server.start()
    try:
        check(callback == "https://127.0.0.1:17819/callback", f"监听地址：{callback}")
        cert, key = tdoc_auth.cert_paths(work)
        check(cert.exists() and key.exists(), "自签证书已生成")

        def fire() -> None:
            time.sleep(0.6)
            context = ssl.create_default_context()
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE  # 自签证书，测试里跳过校验
            try:
                urllib.request.urlopen(callback + "?code=CODE9&state=ST",
                                       context=context, timeout=8).read()
            except Exception:  # noqa: BLE001 - 只关心服务端有没有收到
                pass

        threading.Thread(target=fire, daemon=True).start()
        code, state = server.wait(10)
        check((code, state) == ("CODE9", "ST"), f"收到 code={code} state={state}")
    finally:
        server.stop()


def test_token_store(tmp: Path) -> None:
    print("4) 令牌存储")
    settings_file = tmp / "settings.json"
    original = tdoc_auth.prefs.settings_path
    tdoc_auth.prefs.settings_path = lambda: settings_file  # type: ignore[assignment]
    try:
        store = tdoc_auth.TokenStore()
        check(store.load().present is False, "初始为空")
        store.save(tdoc_auth.Tokens(access_token="AT", refresh_token="RT", open_id="OID",
                                    expires_at=time.time() + 10 * 86400))
        loaded = store.load()
        check(loaded.access_token == "AT" and loaded.open_id == "OID", "存取一致")
        check(not loaded.needs_refresh, "10 天后过期 → 暂不续期")
        check("已连接" in loaded.describe(), f"文案：{loaded.describe()}")

        store.save(tdoc_auth.Tokens(access_token="AT", expires_at=time.time() + 3600))
        check(store.load().needs_refresh, "1 小时后过期 → 触发续期")
        store.save(tdoc_auth.Tokens(access_token="AT", expires_at=time.time() - 10))
        check(store.load().expired, "已过期能识别")
        store.clear()
        check(not store.load().present, "清空生效")
    finally:
        tdoc_auth.prefs.settings_path = original  # type: ignore[assignment]


def main() -> int:
    work = Path(tempfile.mkdtemp(prefix="tdoc_auth_"))
    test_authorize_url()
    test_parse_callback()
    test_local_callback(work)
    test_token_store(work)

    print()
    if failures:
        print(f"失败 {len(failures)} 项：")
        for item in failures:
            print(" - " + item)
        return 1
    print("全部通过 ✓")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
