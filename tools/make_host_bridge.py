"""把当前 WorkBuddy 会话里的腾讯文档网关凭据写到本机联调文件。

    python tools/make_host_bridge.py            # 生成/刷新
    python tools/make_host_bridge.py --show     # 只看位置与状态，不写

用途：打包出来的 .app / .exe 从 Finder 双击启动时，拿不到宿主注入的
`CODEBUDDY_MCP_CONFIG` 环境变量，于是「拉取腾讯文档」会提示没有凭证。
这个脚本把网关地址和认证头存到 `~/.config/AwardStudio/tdoc_host.json`，
软件在缺少 OAuth 凭证时会退而用它（仅本机联调，不要随包分发）。

注意：文件里是**当前会话**的凭据，WorkBuddy 重启或退出登录后会失效，
重新跑一次本脚本即可。
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.tencent_docs import bridge_path  # noqa: E402


def main() -> int:
    target = bridge_path()
    if "--show" in sys.argv:
        if not target.exists():
            print(f"未生成：{target}")
            return 0
        data = json.loads(target.read_text(encoding="utf-8"))
        server = (data.get("mcpServers") or {}).get("connector-proxy") or {}
        print(f"位置：{target}")
        print(f"网关：{server.get('url') or '—'}")
        print(f"认证头：{', '.join(sorted((server.get('headers') or {}).keys())) or '—'}")
        return 0

    raw = os.environ.get("CODEBUDDY_MCP_CONFIG", "")
    if not raw:
        print("没有 CODEBUDDY_MCP_CONFIG 环境变量：请在 WorkBuddy 的终端里运行本脚本。")
        return 2
    try:
        config = json.loads(raw)
    except ValueError:
        print("CODEBUDDY_MCP_CONFIG 不是合法 JSON")
        return 2

    servers = config.get("mcpServers") or {}
    server = servers.get("connector-proxy") or servers.get("workbuddy") or {}
    url = str(server.get("url") or "")
    headers = {k: v for k, v in (server.get("headers") or {}).items()
               if isinstance(k, str) and isinstance(v, str) and v}
    if not url.endswith("/mcp") or not any(k.lower() == "authorization" for k in headers):
        print("环境变量里没有可用的 MCP 网关条目，无法生成")
        return 2

    payload = {"_note": "本机联调凭据，WorkBuddy 重启后需重新生成",
               "mcpServers": {"connector-proxy": {"url": url, "headers": headers}}}
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        target.chmod(0o600)
    except OSError:
        pass
    print(f"已写入：{target}")
    print(f"网关：{url}")
    print(f"认证头：{', '.join(sorted(headers))}")
    print("现在启动打包版软件，点「载入」就能拉到表了（凭证随会话失效，失效后重跑本脚本）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
