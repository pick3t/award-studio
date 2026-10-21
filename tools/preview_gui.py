"""界面截图工具（开发/文档用）：启动界面、截图、自动退出。

    python tools/preview_gui.py [输出png路径]
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from PIL import ImageGrab  # noqa: E402

import app as gui  # noqa: E402


def main() -> None:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "docs" / "screenshot.png"
    out.parent.mkdir(parents=True, exist_ok=True)

    window = gui.AwardStudio(demo=True)
    window.attributes("-topmost", True)
    window.lift()

    def capture() -> None:
        window.update_idletasks()
        # winfo_* 已经是物理像素，直接用
        x, y = window.winfo_rootx(), window.winfo_rooty()
        w, h = window.winfo_width(), window.winfo_height()
        image = ImageGrab.grab(bbox=(x, y, x + w, y + h))
        image.save(out)
        print(f"saved {out} ({image.size[0]}x{image.size[1]})")
        window.destroy()

    window.after(1800, capture)
    window.mainloop()


if __name__ == "__main__":
    main()
