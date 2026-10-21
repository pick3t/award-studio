"""Award Studio · 比赛评分排名与奖状批量生成

运行：
    python app.py            正常启动
    python app.py --demo     启动并自动载入 samples 里的测试数据
"""

from __future__ import annotations

import calendar
import os
import queue
import re
import sys
import threading
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

if sys.platform == "win32":
    # 让 Windows 按系统缩放渲染，避免高分屏下界面被截断
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        pass

import customtkinter as ctk

BASE_DIR = Path(__file__).resolve().parent
# 打包成 exe 后，示例数据放在 exe 同目录
DATA_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else BASE_DIR
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from core.certificate import (  # noqa: E402
    CertificateContext,
    CertificateError,
    certificate_filename,
    generate_certificate,
    inspect_template,
    merge_documents,
)
from core import fields as field_rules  # noqa: E402
from core import settings as prefs  # noqa: E402
from core.excel_reader import ExcelParseError, load_players  # noqa: E402
from core.ranking import award_conflicts, rank_players, winners_of  # noqa: E402

# ------------------------------------------------------------------ 设计令牌
BG = "#F3F5FA"
CARD = "#FFFFFF"
INK = "#111827"
MUTED = "#6B7280"
LINE = "#E8EBF2"
PRIMARY = "#4F46E5"
PRIMARY_HOVER = "#4338CA"
PRIMARY_SOFT = "#EEF0FF"
SUCCESS = "#059669"
WARNING = "#D97706"
DANGER = "#DC2626"
GOLD = "#B8860B"

FONT = "Microsoft YaHei UI"

# 界面字号（磅）：以正文 12pt 为基准，整体等比缩放
WINDOW_W, WINDOW_H = 1300, 860

FONT_SIZE_MIN = 10
FONT_SIZE_MAX = 28
FONT_SIZE_DEFAULT = 12

# 首次启动的奖项方案
DEFAULT_AWARDS = [
    {"name": "一等奖", "count": 1},
    {"name": "二等奖", "count": 2},
    {"name": "三等奖", "count": 3},
]


def short(text: str, limit: int = 20) -> str:
    return text if len(text) <= limit else text[: limit - 1] + "…"


class DatePicker(ctk.CTkToplevel):
    """一个轻量的日期选择弹窗（周一开始）。"""

    def __init__(self, master, initial=None, on_pick=None, font=None):
        super().__init__(master)
        self.on_pick = on_pick
        today = date.today()
        self.current = initial or today
        self.selected = initial
        self.today = today

        self.title("选择日期")
        self.resizable(False, False)
        self.transient(master)
        self.configure(fg_color=CARD)
        self._font = font or ctk.CTkFont(FONT, 12)

        self._build()
        self._render()
        self.geometry("300x360")
        self.after(60, self._center_on)
        self.after(120, self._grab)

    def _grab(self) -> None:
        try:
            self.grab_set()
        except Exception:  # noqa: BLE001
            pass

    def _center_on(self) -> None:
        try:
            master = self.master
            x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
            y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 3
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
            self.focus_force()
        except Exception:  # noqa: BLE001
            pass

    def _build(self) -> None:
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", padx=12, pady=(12, 6))
        ctk.CTkButton(head, text="‹", width=30, height=28, corner_radius=6,
                      fg_color="#F3F5FA", hover_color="#E6EAF3", text_color=INK,
                      font=self._font, command=lambda: self._shift_month(-1)).pack(side="left")
        self.month_label = ctk.CTkLabel(head, text="", font=self._font, text_color=INK)
        self.month_label.pack(side="left", expand=True)
        ctk.CTkButton(head, text="›", width=30, height=28, corner_radius=6,
                      fg_color="#F3F5FA", hover_color="#E6EAF3", text_color=INK,
                      font=self._font, command=lambda: self._shift_month(1)).pack(side="right")

        weekdays = ctk.CTkFrame(self, fg_color="transparent")
        weekdays.pack(fill="x", padx=12)
        for index, name in enumerate(("一", "二", "三", "四", "五", "六", "日")):
            weekdays.grid_columnconfigure(index, weight=1)
            ctk.CTkLabel(weekdays, text=name, font=self._font, text_color=MUTED,
                         width=36).grid(row=0, column=index)

        self.grid_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.grid_frame.pack(fill="both", expand=True, padx=12, pady=(2, 6))
        for index in range(7):
            self.grid_frame.grid_columnconfigure(index, weight=1)

        foot = ctk.CTkFrame(self, fg_color="transparent")
        foot.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(foot, text="今天", height=28, corner_radius=7, fg_color="#F3F5FA",
                      hover_color="#E6EAF3", text_color=INK, font=self._font,
                      command=self._pick_today).pack(side="left")
        ctk.CTkButton(foot, text="取消", height=28, corner_radius=7, fg_color="#F3F5FA",
                      hover_color="#E6EAF3", text_color=MUTED, font=self._font,
                      command=self.destroy).pack(side="right")

    def _shift_month(self, delta: int) -> None:
        month = self.current.month - 1 + delta
        year = self.current.year + month // 12
        month = month % 12 + 1
        self.current = date(year, month, 1)
        self._render()

    def _render(self) -> None:
        for child in self.grid_frame.winfo_children():
            child.destroy()
        self.month_label.configure(text=f"{self.current.year} 年 {self.current.month} 月")

        weeks = calendar.Calendar(firstweekday=0).monthdayscalendar(
            self.current.year, self.current.month)
        for row, week in enumerate(weeks):
            for col, day in enumerate(week):
                if day == 0:
                    ctk.CTkLabel(self.grid_frame, text="", width=36, height=30).grid(
                        row=row, column=col, padx=1, pady=1)
                    continue
                value = date(self.current.year, self.current.month, day)
                is_today = value == self.today
                is_selected = value == self.selected
                fg = PRIMARY if is_selected else ("#EEF0FF" if is_today else "transparent")
                color = "#FFFFFF" if is_selected else (PRIMARY if is_today else INK)
                ctk.CTkButton(
                    self.grid_frame, text=str(day), width=36, height=30, corner_radius=7,
                    fg_color=fg, hover_color="#E6EAF3", text_color=color, font=self._font,
                    command=lambda v=value: self._pick(v),
                ).grid(row=row, column=col, padx=1, pady=1)

    def _pick(self, value: date) -> None:
        if self.on_pick:
            self.on_pick(value)
        self.destroy()

    def _pick_today(self) -> None:
        self._pick(self.today)


class AwardStudio(ctk.CTk):
    def __init__(self, demo: bool = False) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        if getattr(sys, "frozen", False):
            # 打包后的进程里 Tk 与系统缩放换算不一致，界面会被裁切；
            # 这里锁定为 1:1 缩放，保证窗口尺寸和内容尺寸一致。
            for kwargs in (
                {"widget_scaling": 1.0, "window_scaling": 1.0},
                {"widget_scaling": 1.0},
                {},
            ):
                try:
                    self._set_scaling(**kwargs)
                    break
                except TypeError:
                    continue
                except Exception:  # noqa: BLE001
                    break
            try:
                self.tk.call("tk", "scaling", 1.33)
            except Exception:  # noqa: BLE001
                pass
        self.title("Award Studio · 比赛评分与奖状生成")
        icon = BASE_DIR / "assets" / "icon.ico"
        if icon.exists():
            try:
                self.iconbitmap(str(icon))
            except Exception:  # noqa: BLE001
                pass
        self.configure(fg_color=BG)

        # 状态
        self.excel_path: Path | None = None
        self.template_path: Path | None = None
        self.output_dir: Path | None = None
        self.players = []
        self.entries = []
        self.winners = 5
        self.quota = 5
        self.max_winners = 20
        self.excel_title = ""
        self.ui_queue: queue.Queue = queue.Queue()
        self.generated: list[Path] = []
        self.busy = False

        saved = prefs.load()
        try:
            size = int(saved.get("font_size", FONT_SIZE_DEFAULT))
        except (TypeError, ValueError):
            # 兼容早期只存百分比的版本
            legacy = (0.85, 1.0, 1.12, 1.25)
            try:
                size = round(FONT_SIZE_DEFAULT * legacy[int(saved.get("text_scale_index", 1))])
            except (TypeError, ValueError, IndexError):
                size = FONT_SIZE_DEFAULT
        self.font_size = max(FONT_SIZE_MIN, min(FONT_SIZE_MAX, size))
        self.text_scale = self.font_size / FONT_SIZE_DEFAULT
        self._fonts: dict[str, tuple] = {}

        self._step_labels: list[tuple] = []
        self._award_rows: list[tuple] = []
        self._scale_job = None
        self._drag_item = None
        self._drag_y = 0
        self._drag_moved = False
        self.field_rules: list[field_rules.FieldRule] = [
            field_rules.FieldRule.from_dict(item)
            for item in saved.get("field_rules", [])
            if isinstance(item, dict) and item.get("key")
        ]
        self.template_keys: list[str] = []
        self.column_names: list[str] = []

        self.awards = self._parse_awards(saved.get("awards"))
        self.award_mode = bool(saved.get("award_mode", True))

        # 窗口尺寸与字号解耦：始终按固定基准开窗，字号只影响内容
        self._win_size = (WINDOW_W, WINDOW_H)
        self._center(*self._win_size)
        self.after(120, self._capture_window_size)
        self.minsize(1000, 620)

        self.font_h1 = self._mkfont("h1", 19, "bold")
        self.font_sub = self._mkfont("sub", 11)
        self.font_step = self._mkfont("step", 12, "bold")
        self.font_body = self._mkfont("body", 12)
        self.font_small = self._mkfont("small", 11)
        self.font_num = self._mkfont("num", 22, "bold", "Segoe UI", cap=32)

        self._setup_style()
        self._build()
        self._refresh_table()
        if self.text_scale != 1.0:
            self._apply_text_scale()
        self._log("就绪 · 选择成绩表后会自动按分数排名", "muted")
        self._sync_buttons()

        self._enable_file_drop()

        if demo:
            self.after(300, self._load_demo)
        if "--diag" in sys.argv:
            self.after(1200, self._write_diag)

    def _write_diag(self) -> None:
        lines = []
        probes = [
            ("frozen", lambda: getattr(sys, "frozen", False)),
            ("fpixels", lambda: self.winfo_fpixels("1i")),
            ("tk_scaling", lambda: self.tk.call("tk", "scaling")),
            ("screen", lambda: f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}"),
            ("window", lambda: f"{self.winfo_width()}x{self.winfo_height()}"),
            ("window_scaling", lambda: self._get_window_scaling()),
            ("widget_scaling", lambda: self._get_widget_scaling()),
            ("ctk_version", lambda: ctk.__version__),
        ]
        for name, fn in probes:
            try:
                lines.append(f"{name}={fn()}")
            except Exception as exc:  # noqa: BLE001
                lines.append(f"{name}=ERR {exc}")
        try:
            (DATA_DIR / "diag.txt").write_text("\n".join(lines), encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass

    # -------------------------------------------------------------- 布局
    def _center(self, width: int, height: int) -> None:
        """按逻辑尺寸开窗并居中，内部换算成物理像素。"""
        self.update_idletasks()
        try:
            scale = self._get_window_scaling() or 1.0
        except Exception:  # noqa: BLE001
            scale = 1.0
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        # 留出任务栏和边距
        width = max(800, min(width, int((screen_w - 80) / scale)))
        height = max(560, min(height, int((screen_h - 120) / scale)))

        phys_w, phys_h = int(width * scale), int(height * scale)
        x = max(0, (screen_w - phys_w) // 2)
        y = max(0, (screen_h - phys_h) // 2 - 20)

        if getattr(sys, "frozen", False):
            # 打包后 CTk 不换算窗口尺寸，位置反而会换算，这里直接给物理值
            self.tk.call("wm", "geometry", self._w,
                         f"{phys_w}x{phys_h}+{int(x * scale)}+{int(y * scale)}")
        else:
            self.geometry(f"{width}x{height}+{x}+{y}")

    def _setup_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Award.Treeview",
            background=CARD,
            fieldbackground=CARD,
            foreground=INK,
            rowheight=40,
            borderwidth=0,
            relief="flat",
        )
        style.configure(
            "Award.Treeview.Heading",
            background="#F8FAFD",
            foreground=MUTED,
            relief="flat",
            borderwidth=0,
            padding=(10, 10),
        )
        style.map(
            "Award.Treeview",
            background=[("selected", PRIMARY_SOFT)],
            foreground=[("selected", "#312E81")],
        )
        style.map("Award.Treeview.Heading", background=[("active", "#F1F4FA")])
        try:
            style.layout("Award.Treeview", [("Award.Treeview.treearea", {"sticky": "nswe"})])
        except Exception:  # noqa: BLE001
            pass
        self._restyle_table()

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_header()

        self.main_frame = main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 14))
        main.grid_columnconfigure(0, weight=0, minsize=352)
        main.grid_columnconfigure(1, weight=1)
        main.grid_rowconfigure(0, weight=1)

        self._build_left(main)
        self._build_right(main)

    def _build_header(self) -> None:
        self.header_frame = header = ctk.CTkFrame(self, fg_color="transparent", height=72)
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(12, 8))
        header.grid_columnconfigure(1, weight=1)

        self.badge = badge = ctk.CTkFrame(header, width=44, height=44, corner_radius=13, fg_color=PRIMARY)
        badge.grid(row=0, column=0, rowspan=2, sticky="w")
        badge.grid_propagate(False)
        ctk.CTkLabel(badge, text="A", font=self._mkfont("badge", 22, "bold", "Segoe UI"),
                     text_color="#FFFFFF").place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(header, text="Award Studio", font=self.font_h1, text_color=INK,
                     anchor="w").grid(row=0, column=1, sticky="w", padx=(14, 0))
        ctk.CTkLabel(header, text="比赛评分排名 · 奖状批量生成", font=self.font_sub,
                     text_color=MUTED, anchor="w").grid(row=1, column=1, sticky="w", padx=(14, 0))

        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.grid(row=0, column=2, rowspan=2, sticky="e")

        # 字号调节
        font_box = ctk.CTkFrame(actions, fg_color="#F8FAFD", corner_radius=9,
                                border_width=1, border_color=LINE)
        font_box.grid(row=0, column=0, padx=(0, 10))
        ctk.CTkButton(font_box, text="A-", width=30, height=30, corner_radius=7,
                      fg_color="transparent", hover_color="#E9EEF7", text_color=INK,
                      font=self.font_small, command=lambda: self._change_scale(-1)
                      ).pack(side="left", padx=(3, 0), pady=2)
        self.scale_var = ctk.StringVar(value=str(self.font_size))
        self.scale_entry = ctk.CTkEntry(
            font_box, textvariable=self.scale_var, width=46, height=26, corner_radius=6,
            justify="center", font=self.font_small,
            fg_color="#FFFFFF", border_width=1, border_color=LINE,
        )
        self.scale_entry.pack(side="left", padx=(4, 2), pady=3)
        self.scale_entry.bind("<Return>", self._on_scale_entry)
        self.scale_entry.bind("<KP_Enter>", self._on_scale_entry)
        self.scale_entry.bind("<FocusOut>", self._on_scale_entry)
        self.scale_entry.bind("<Up>", lambda _e: self._change_scale(1))
        self.scale_entry.bind("<Down>", lambda _e: self._change_scale(-1))
        ctk.CTkLabel(font_box, text="pt", font=self.font_small, text_color=MUTED).pack(
            side="left", padx=(0, 4))
        ctk.CTkButton(font_box, text="A+", width=30, height=30, corner_radius=7,
                      fg_color="transparent", hover_color="#E9EEF7", text_color=INK,
                      font=self.font_small, command=lambda: self._change_scale(1)
                      ).pack(side="left", padx=(0, 3), pady=2)

        ctk.CTkButton(actions, text="使用说明", width=92, height=34, corner_radius=9,
                      fg_color=CARD, hover_color="#F1F4FA", text_color=INK,
                      border_width=1, border_color=LINE, font=self.font_small,
                      command=self._show_help).grid(row=0, column=1, padx=(0, 8))
        ctk.CTkButton(actions, text="打开输出目录", width=112, height=34, corner_radius=9,
                      fg_color=CARD, hover_color="#F1F4FA", text_color=INK,
                      border_width=1, border_color=LINE, font=self.font_small,
                      command=self._open_output).grid(row=0, column=2)

    # -------------------------------------------------------------- 字号
    def _mkfont(self, key: str, size: int, weight: str = "normal", family: str = FONT,
                cap: int | None = None):
        """按当前字号倍率创建字体，并登记以便统一调整。

        同一个 key 只创建一次，后续调用直接复用，否则新字体对象不在登记表里，
        字号变化时那些控件不会跟着变。
        """
        if key in self._fonts:
            return self._fonts[key][0]
        font = ctk.CTkFont(family=family, size=max(8, self._font_px(size, cap)),
                           weight=weight)
        self._fonts[key] = (font, size, weight, family, cap)
        return font

    def _font_px(self, size: int, cap: int | None = None) -> int:
        value = max(8, round(size * self.text_scale))
        return min(value, cap) if cap else value

    def _px(self, value: float) -> int:
        """把设计尺寸换算成当前字号下的实际像素。"""
        return max(1, int(round(value * self.text_scale)))

    def _apply_text_scale(self) -> None:
        """分批应用字号：CTk 每改一个字体都要重绘所有相关控件，一次性做完会冻结窗口。"""
        self._apply_token = getattr(self, "_apply_token", 0) + 1
        token = self._apply_token
        queue = list(self._fonts.values())

        def table_step():
            if token != self._apply_token:
                return
            self._restyle_table()
            self.after(1, font_step)

        def font_step():
            if token != self._apply_token:
                return
            batch, rest = queue[:4], queue[4:]
            queue[:] = rest
            for font, size, weight, family, cap in batch:
                try:
                    font.configure(size=self._font_px(size, cap))
                except Exception:  # noqa: BLE001
                    pass
            if queue:
                self.after(1, font_step)
            else:
                self.after(1, layout_step)

        def layout_step():
            if token != self._apply_token:
                return
            self._resize_layout()
            # 控件变大后 Tk 会把窗口一起撑大，这里锁回原来的尺寸
            self.after(30, self._restore_window_size)

        self.after(0, table_step)

    def _restyle_table(self) -> None:
        style = ttk.Style()
        style.configure(
            "Award.Treeview",
            font=(FONT, max(8, round(10 * self.text_scale))),
            rowheight=max(24, round(40 * self.text_scale)),
        )
        style.configure(
            "Award.Treeview.Heading",
            font=(FONT, max(8, round(9 * self.text_scale)), "bold"),
        )

    def _resize_layout(self) -> None:
        """字号变化后同步控件尺寸。窗口本身保持不变，避免界面整体膨胀。"""
        px = self._px
        # 左栏按比例变宽，但设上限，防止挤占右侧表格
        panel_w = min(px(352), 560)
        narrow = min(self.text_scale, 1.4)
        try:
            self.left_panel.configure(width=panel_w)
            self.left_scroll.configure(width=max(200, panel_w - 30))
            self.header_frame.configure(height=px(72))
            self.badge.configure(width=px(44), height=px(44))
            for button, height in (
                (self.excel_btn, 40), (self.template_btn, 40), (self.output_btn, 32),
                (self.add_award_btn, 28), (self.generate_btn, 46),
            ):
                button.configure(height=px(height))
            self.log_card.configure(height=px(140))
            self.log_box.configure(height=px(100))
            for title_label, title_h, sub_label, sub_h in self._step_labels:
                title_label.configure(height=px(title_h))
                if sub_label is not None:
                    sub_label.configure(height=px(sub_h))
            self._fit_columns()
            wrap = max(160, panel_w - 60)
            self.excel_hint.configure(wraplength=wrap)
            self.template_hint.configure(wraplength=wrap)
            self._fit_award_rows()
            if self.field_rules and self.template_keys:
                self._render_mappings()
        except Exception:  # noqa: BLE001 - 控件还没建好时跳过
            return

    def _is_maximized(self) -> bool:
        try:
            import ctypes

            return bool(ctypes.windll.user32.IsZoomed(int(self.frame(), 16)))
        except Exception:  # noqa: BLE001
            return False

    def _capture_window_size(self) -> None:
        """记住当前窗口尺寸（手动拉大过也会记住；最大化时不记，避免把屏幕尺寸当窗口尺寸）"""
        if self._is_maximized():
            return
        try:
            self.update_idletasks()
            scale = self._get_window_scaling() or 1.0
            width, height = self.winfo_width(), self.winfo_height()
            # winfo 与 geometry 的单位差一个缩放系数，源码与打包环境都一样换算
            width, height = int(width / scale), int(height / scale)
            if width > 200 and height > 200:
                self._win_size = (width, height)
        except Exception:  # noqa: BLE001
            pass

    def _on_window_resize(self, event) -> None:
        if event.widget is not self:
            return

        maximized = self._is_maximized()
        was = getattr(self, "_was_maximized", False)
        self._was_maximized = maximized
        if was and not maximized:
            # 刚从最大化还原：锁回到最大化之前的窗口尺寸
            self.after(90, self._restore_window_size)
        elif not maximized:
            # 用户拖动改变尺寸后记住它，下次切字号或还原时用这个大小
            job = getattr(self, "_remember_job", None)
            if job is not None:
                try:
                    self.after_cancel(job)
                except Exception:  # noqa: BLE001
                    pass
            self._remember_job = self.after(600, self._capture_window_size)

        job = getattr(self, "_resize_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except Exception:  # noqa: BLE001
                pass
        self._resize_job = self.after(220, self._fit_columns)

    def _fit_columns(self) -> None:
        self._resize_job = None
        columns = getattr(self, "_table_columns", None)
        if not columns:
            return
        try:
            base_total = sum(width for _key, _text, width, _anchor in columns)
            available = max(420, self.tree.winfo_width() - 20)
            factor = min(self.text_scale, 1.5, max(1.0, available / base_total))
            for key, _text, width, _anchor in columns:
                self.tree.column(key, width=max(48, int(width * factor)))
        except Exception:  # noqa: BLE001
            pass

    def _restore_window_size(self) -> None:
        size = getattr(self, "_win_size", None)
        if not size or self._is_maximized():
            return
        try:
            if getattr(sys, "frozen", False):
                self.tk.call("wm", "geometry", self._w, f"{size[0]}x{size[1]}")
            else:
                self.geometry(f"{size[0]}x{size[1]}")
        except Exception:  # noqa: BLE001
            pass

    def _fit_award_rows(self) -> None:
        """只调整已有奖项行的尺寸，不销毁重建（重建要 150ms 且会打断编辑）。"""
        px = self._px
        for item_row in self._award_rows:
            if len(item_row) < 3:
                continue
            row, _item, widgets = item_row
            handle, name_entry, count_entry, del_btn = widgets
            handle.configure(width=px(20), height=px(30))
            name_entry.configure(height=px(30))
            count_entry.configure(width=px(46), height=px(30))
            del_btn.configure(width=px(26), height=px(28))

    def _change_scale(self, delta: int) -> None:
        self._set_font_size(self.font_size + delta)

    def _on_scale_entry(self, event=None) -> None:
        raw = self.scale_var.get().strip().lower().replace("pt", "").strip()
        try:
            size = int(float(raw))
        except ValueError:
            self.scale_var.set(str(self.font_size))
            return
        if size == self.font_size:
            self.scale_var.set(str(self.font_size))
            return
        self._set_font_size(size)

    def _set_font_size(self, size: int) -> None:
        size = max(FONT_SIZE_MIN, min(FONT_SIZE_MAX, int(size)))
        self.scale_var.set(str(size))
        if size == self.font_size:
            return
        if self._scale_job is None:
            self._capture_window_size()
        self.font_size = size
        self.text_scale = size / FONT_SIZE_DEFAULT
        # 连续点击/输入时只渲染最后一次，避免每次都全量重排
        if self._scale_job is not None:
            try:
                self.after_cancel(self._scale_job)
            except Exception:  # noqa: BLE001
                pass
        self._scale_job = self.after(140, self._apply_scale_now)

    def _apply_scale_now(self) -> None:
        self._scale_job = None
        self._apply_text_scale()
        self._save_settings()
        self._log(f"界面字号：{self.font_size} pt", "muted")

    def _save_settings(self) -> None:
        data = prefs.load()
        data["font_size"] = self.font_size
        data["award_mode"] = bool(self.award_switch.get())
        data["field_rules"] = [rule.to_dict() for rule in self.field_rules]
        data["awards"] = [{"name": item["name"], "count": item["count"]} for item in self.awards]
        prefs.save(data)

    def _build_left(self, parent) -> None:
        panel = ctk.CTkFrame(parent, width=352, corner_radius=18, fg_color=CARD,
                             border_width=1, border_color=LINE)
        panel.grid(row=0, column=0, sticky="nsw")
        panel.grid_propagate(False)
        panel.configure(width=352)
        self.left_panel = panel

        # 底部固定操作区：无论窗口多高，按钮始终可见
        footer = ctk.CTkFrame(panel, fg_color="transparent")
        footer.pack(side="bottom", fill="x", padx=18, pady=(4, 14))
        ctk.CTkFrame(footer, height=1, fg_color=LINE).pack(fill="x", pady=(0, 12))
        self.generate_btn = ctk.CTkButton(
            footer, text="生成奖状", height=46, corner_radius=11,
            fg_color=PRIMARY, hover_color=PRIMARY_HOVER, text_color="#FFFFFF",
            font=self._mkfont("cta", 14, "bold"), command=self._start_generate,
        )
        self.generate_btn.pack(fill="x")
        self.status_label = ctk.CTkLabel(footer, text="", font=self.font_small,
                                         text_color=MUTED, anchor="w")
        self.status_label.pack(fill="x", pady=(6, 0))

        # 窗口高度不够时可以滚动，避免中间步骤被裁掉
        self.left_scroll = ctk.CTkScrollableFrame(
            panel, width=322, fg_color="transparent", corner_radius=0,
            scrollbar_fg_color="transparent", scrollbar_button_color="#E3E7F0",
            scrollbar_button_hover_color="#C9CFDD",
        )
        self.left_scroll.pack(fill="both", expand=True, padx=(6, 0), pady=(6, 0))

        wrap = ctk.CTkFrame(self.left_scroll, fg_color="transparent")
        wrap.pack(fill="both", expand=True, padx=12, pady=8)

        # ① 成绩表
        self._step_title(wrap, "①", "成绩表", "读取选手与分数，自动识别列")
        self.excel_btn = ctk.CTkButton(
            wrap, text="选择 Excel 成绩表", height=40, corner_radius=10,
            fg_color="#F8FAFD", hover_color="#EFF3FA", text_color=INK,
            border_width=1, border_color=LINE, font=self.font_body,
            command=self._choose_excel,
        )
        self.excel_btn.pack(fill="x", pady=(8, 3))
        self.excel_hint = ctk.CTkLabel(wrap, text="支持 .xlsx / .xlsm，也可直接把文件拖进窗口", font=self.font_small,
                                       text_color=MUTED, anchor="w", justify="left")
        self.excel_hint.pack(fill="x")

        # ② 奖状模板
        self._step_title(wrap, "②", "奖状模板", "Word 模板中用 {{姓名}}、{{名次}} 占位")
        self.template_btn = ctk.CTkButton(
            wrap, text="选择 Word 奖状模板", height=40, corner_radius=10,
            fg_color="#F8FAFD", hover_color="#EFF3FA", text_color=INK,
            border_width=1, border_color=LINE, font=self.font_body,
            command=self._choose_template,
        )
        self.template_btn.pack(fill="x", pady=(8, 3))
        self.template_hint = ctk.CTkLabel(wrap, text="支持 .docx，也可直接把文件拖进窗口", font=self.font_small,
                                          text_color=MUTED, anchor="w", justify="left",
                                          wraplength=286)
        self.template_hint.pack(fill="x")

        # ③ 字段映射
        self.mapping_anchor = self._step_title(wrap, "③", "字段映射", "指定占位符取什么值")
        # 固定容器：内部控件按需显示，容器本身始终占位以保持顺序（内容为空时不占高度）
        self.mapping_section = ctk.CTkFrame(wrap, fg_color="transparent")
        self.mapping_box = ctk.CTkFrame(self.mapping_section, fg_color="transparent")
        self.mapping_hint = ctk.CTkLabel(
            self.mapping_section, text="选择奖状模板后，这里会列出模板里的占位符",
            font=self.font_small, text_color=MUTED, anchor="w",
            justify="left", wraplength=286,
        )
        self.mapping_toggle = ctk.CTkButton(
            self.mapping_section, text="收起映射", height=22, fg_color="transparent",
            hover_color="#F1F4FA", text_color=MUTED, font=self.font_small,
            anchor="w", command=self._toggle_mappings,
        )
        self.mapping_open = False

        # ④ 奖项
        self._step_title(wrap, "④", "奖项设置", "自定义名称与人数，按住 ≡ 可拖动排序")

        self.award_switch = ctk.CTkSwitch(
            wrap, text="设置奖项等级", font=self.font_small,
            text_color=MUTED, progress_color=PRIMARY, command=self._on_award_toggle,
        )
        if self.award_mode:
            self.award_switch.select()
        else:
            self.award_switch.deselect()
        self.award_switch.pack(anchor="w", pady=(8, 2))

        self.award_area = ctk.CTkFrame(wrap, fg_color="transparent")
        self.award_area.pack(fill="x")

        # A) 奖项列表（开启「设置奖项等级」时显示）
        self.awards_panel = ctk.CTkFrame(self.award_area, fg_color="transparent")
        self.awards_box = ctk.CTkFrame(self.awards_panel, fg_color="transparent")
        self.awards_box.pack(fill="x", pady=(4, 0))
        self.add_award_btn = ctk.CTkButton(
            self.awards_panel, text="＋ 添加奖项", height=28, corner_radius=8,
            fg_color="#F8FAFD", hover_color="#EFF3FA", text_color=PRIMARY,
            border_width=1, border_color=LINE, font=self.font_small,
            command=self._add_award,
        )
        self.add_award_btn.pack(fill="x", pady=(6, 0))
        self.award_total = ctk.CTkLabel(
            self.awards_panel, text="", font=self.font_small, text_color=MUTED, anchor="w",
        )
        self.award_total.pack(fill="x", pady=(4, 0))
        self.award_warn = ctk.CTkLabel(
            self.awards_panel, text="", font=self.font_small, text_color=WARNING,
            anchor="w", justify="left", wraplength=286,
        )

        # B) 仅排名次（关闭奖项时显示）
        self.qty_panel = ctk.CTkFrame(self.award_area, fg_color="transparent")
        qty = ctk.CTkFrame(self.qty_panel, fg_color="transparent")
        qty.pack(fill="x", pady=(6, 0))
        self.qty_label = ctk.CTkLabel(qty, text="5",
                                      font=self._mkfont("qty", 20, "bold", "Segoe UI"),
                                      text_color=PRIMARY, anchor="w")
        self.qty_label.pack(side="left")
        ctk.CTkLabel(qty, text="名", font=self.font_small, text_color=MUTED).pack(
            side="left", padx=(6, 12))
        self.slider = ctk.CTkSlider(
            qty, from_=1, to=20, number_of_steps=19, height=18,
            fg_color="#E6E9F2", progress_color=PRIMARY, button_color=PRIMARY,
            button_hover_color=PRIMARY_HOVER, command=self._on_slider,
        )
        self.slider.set(5)
        self.slider.pack(side="left", fill="x", expand=True)
        self.qty_hint = ctk.CTkLabel(
            self.qty_panel, text="只排名次，奖状里不填奖项等级",
            font=self.font_small, text_color=MUTED, anchor="w", justify="left",
            wraplength=286,
        )
        self.qty_hint.pack(fill="x", pady=(4, 0))

        self._sync_award_ui()

        # ⑤ 输出
        self._step_title(wrap, "⑤", "输出位置", "")
        output_row = ctk.CTkFrame(wrap, fg_color="transparent")
        output_row.pack(fill="x", pady=(6, 0))
        self.output_btn = ctk.CTkButton(
            output_row, text="选择输出文件夹", height=32, corner_radius=9,
            fg_color="#F8FAFD", hover_color="#EFF3FA", text_color=INK,
            border_width=1, border_color=LINE, font=self.font_small,
            command=self._choose_output,
        )
        self.output_btn.pack(side="left", fill="x", expand=True)
        self.merge_check = ctk.CTkCheckBox(
            output_row, text="合并版", font=self.font_small,
            text_color=MUTED, checkbox_width=17, checkbox_height=17, corner_radius=5,
            fg_color=PRIMARY, hover_color=PRIMARY_HOVER, border_color="#C7CCDA",
        )
        self.merge_check.pack(side="left", padx=(6, 0))
        ctk.CTkFrame(wrap, fg_color="transparent", height=6).pack(fill="x")

    def _step_title(self, parent, badge: str, title: str, subtitle: str) -> None:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(11, 0))
        ctk.CTkLabel(row, text=badge, font=self.font_step, text_color=PRIMARY,
                     width=18).pack(side="left", anchor="n")
        text = ctk.CTkFrame(row, fg_color="transparent")
        text.pack(side="left", fill="x", expand=True, padx=(6, 0))
        title_label = ctk.CTkLabel(text, text=title, font=self._mkfont("step_title", 12, "bold"),
                                   text_color=INK, height=self._px(19), anchor="w")
        title_label.pack(fill="x")
        sub_label = None
        if subtitle:
            sub_label = ctk.CTkLabel(text, text=subtitle, font=self._mkfont("step_sub", 10),
                                     text_color=MUTED, height=self._px(16), anchor="w")
            sub_label.pack(fill="x")
        self._step_labels.append((title_label, 19, sub_label, 16))
        return row

    def _build_right(self, parent) -> None:
        right = ctk.CTkFrame(parent, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=(18, 0))

        # 统计卡片
        stats = ctk.CTkFrame(right, fg_color="transparent")
        stats.pack(fill="x")
        for i in range(4):
            stats.grid_columnconfigure(i, weight=1, uniform="stat")
        self.stat_values = {}
        for index, (key, label) in enumerate(
            (("total", "参赛人数"), ("max", "最高分"), ("avg", "平均分"), ("awards", "奖状份数"))
        ):
            card = ctk.CTkFrame(stats, corner_radius=16, fg_color=CARD,
                                border_width=1, border_color=LINE)
            card.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 10, 0))
            value = ctk.CTkLabel(card, text="—", font=self.font_num, text_color=INK, anchor="w")
            value.pack(fill="x", padx=16, pady=(12, 0))
            ctk.CTkLabel(card, text=label, font=self.font_small, text_color=MUTED,
                         anchor="w").pack(fill="x", padx=16, pady=(0, 12))
            self.stat_values[key] = value

        # 排名表
        table_card = ctk.CTkFrame(right, corner_radius=18, fg_color=CARD,
                                  border_width=1, border_color=LINE)
        table_card.pack(fill="both", expand=True, pady=(12, 0))
        table_card.grid_columnconfigure(0, weight=1)
        table_card.grid_rowconfigure(2, weight=1)

        head = ctk.CTkFrame(table_card, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 0))
        head.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(head, text="排名总览", font=self._mkfont("card_title", 14, "bold"),
                     text_color=INK, anchor="w").grid(row=0, column=0, sticky="w")
        self.table_note = ctk.CTkLabel(head, text="同分并列名次，前三名高亮",
                                       font=self.font_small, text_color=MUTED, anchor="w")
        self.table_note.grid(row=1, column=0, sticky="w")
        self.merge_hint = ctk.CTkLabel(head, text="",
                                       font=self._mkfont("merge_hint", 11, cap=16),
                                       text_color=PRIMARY, anchor="e")
        self.merge_hint.grid(row=0, column=1, rowspan=2, sticky="e")

        ctk.CTkFrame(table_card, height=1, fg_color=LINE).grid(row=1, column=0, sticky="ew",
                                                              padx=18, pady=(12, 0))

        self.table_holder = holder = ctk.CTkFrame(table_card, fg_color="transparent")
        holder.grid(row=2, column=0, sticky="nsew", padx=(14, 8), pady=(6, 12))
        holder.grid_columnconfigure(0, weight=1)
        holder.grid_rowconfigure(0, weight=1)

        # 没有数据时显示占位提示，而不是一张空表格
        self.empty_panel = ctk.CTkFrame(holder, fg_color="transparent")
        self.empty_panel.grid(row=0, column=0, columnspan=2, sticky="nsew")
        empty_inner = ctk.CTkFrame(self.empty_panel, fg_color="transparent")
        empty_inner.place(relx=0.5, rely=0.45, anchor="center")
        ctk.CTkLabel(empty_inner, text="▤", font=self._mkfont("empty_icon", 30),
                     text_color="#CBD2E0").pack()
        ctk.CTkLabel(empty_inner, text="还没有成绩数据",
                     font=self._mkfont("empty_title", 13), text_color="#8A93A6").pack(pady=(4, 2))
        self.empty_hint = ctk.CTkLabel(
            empty_inner, text="选择左侧的 Excel 成绩表，或直接把文件拖进窗口",
            font=self.font_small, text_color="#A7AFC0", justify="center")
        self.empty_hint.pack()

        self._table_columns: list[tuple] = []
        self._build_table()
        self.bind("<Configure>", self._on_window_resize)

        # 日志
        self.log_card = log_card = ctk.CTkFrame(right, corner_radius=16, fg_color=CARD,
                                border_width=1, border_color=LINE, height=140)
        log_card.pack(fill="x", pady=(12, 0))
        log_card.grid_propagate(False)
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(log_card, text="运行日志", font=self._mkfont("log_title", 11, "bold"),
                     text_color=MUTED, anchor="w").grid(row=0, column=0, sticky="w",
                                                        padx=18, pady=(10, 2))
        self.log_box = ctk.CTkTextbox(log_card, fg_color="transparent", text_color=INK,
                                      font=self._mkfont("log", 11), border_width=0,
                                      wrap="word", height=100)
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))
        self.log_box.configure(state="disabled")
        for tag, color in (("info", INK), ("muted", MUTED), ("ok", SUCCESS),
                           ("warn", WARNING), ("err", DANGER)):
            try:
                self.log_box._textbox.tag_config(tag, foreground=color)
            except Exception:  # noqa: BLE001
                pass

    # -------------------------------------------------------------- 交互
    def _choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="选择成绩表",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
            initialdir=str(self.excel_path.parent) if self.excel_path else str(BASE_DIR),
        )
        if path:
            self._load_excel(Path(path))

    def _load_excel(self, path: Path) -> None:
        try:
            players, warnings, report = load_players(path)
        except ExcelParseError as exc:
            messagebox.showerror("读取失败", str(exc))
            self._log(f"读取失败：{exc}", "err")
            return

        self.excel_path = path
        self.players = players
        self.excel_title = report.title
        self.column_names = list(players[0].fields.keys()) if players else []
        self._build_table()
        self.excel_btn.configure(text=short(path.name, 24), text_color=INK)
        self.excel_hint.configure(
            text=f"{report.sheet} · 表头第 {report.header_row} 行 · {report.describe()}",
            text_color=MUTED,
        )
        self._log(f"已读取 {len(players)} 名选手：{path.name}", "ok")
        for warning in warnings:
            self._log("  " + warning, "warn")

        if self.output_dir is None:
            self.output_dir = path.parent / "output" / "奖状"
        self._update_output_hint()

        # 名额上限跟着人数走，表格里有多少人就能选多少
        total = max(1, len(players))
        self.max_winners = total
        self.slider.configure(to=total, number_of_steps=max(1, total - 1))
        self.winners = min(self.winners, self.max_winners)
        self.slider.set(self.winners)
        self.qty_label.configure(text=str(self.winners))
        self._refresh_field_rules()
        self._recalculate()

    def _choose_template(self) -> None:
        path = filedialog.askopenfilename(
            title="选择奖状模板",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
            initialdir=str(self.template_path.parent) if self.template_path else str(BASE_DIR),
        )
        if path:
            self._load_template(Path(path))

    def _load_template(self, template: Path) -> None:
        try:
            keys = inspect_template(template)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("模板无法读取", f"{exc}\n\n请确认是 .docx 文件且没有加密。")
            return
        self.template_path = template
        self.template_keys = list(keys)
        self.template_btn.configure(text=short(template.name, 24), text_color=INK)
        if keys:
            shown = "、".join(sorted(keys))
            self.template_hint.configure(text=f"占位符：{shown}", text_color=MUTED)
            self._log(f"模板已载入，可用字段：{shown}", "ok")
        else:
            self.template_hint.configure(
                text="模板里没有占位符，请在需要填字的位置写 {{姓名}}、{{名次}}",
                text_color=WARNING,
            )
            self._log("模板里没有找到 {{ }} 占位符，生成时会失败", "warn")
        self.mapping_open = True
        self._refresh_field_rules()
        self._sync_buttons()

    # -------------------------------------------------------------- 文件拖放
    def _enable_file_drop(self) -> None:
        """接收从资源管理器拖进窗口的文件。

        tkinter 自身不支持文件拖放，这里用 Windows 原生的 DragAcceptFiles + 窗口过程
        替换来处理 WM_DROPFILES。任何一步失败都静默降级，用户仍可点按钮选文件。
        """
        if sys.platform != "win32":
            return
        try:
            import ctypes
            from ctypes import wintypes

            user32 = ctypes.windll.user32
            shell32 = ctypes.windll.shell32
            hwnd = int(self.frame(), 16)

            WM_DROPFILES = 0x0233
            GWL_WNDPROC = -4

            user32.GetWindowLongPtrW.restype = ctypes.c_void_p
            user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
            user32.SetWindowLongPtrW.restype = ctypes.c_void_p
            user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
            user32.CallWindowProcW.restype = ctypes.c_ssize_t
            user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT,
                                               wintypes.WPARAM, wintypes.LPARAM]
            shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
            shell32.DragQueryFileW.restype = wintypes.UINT
            shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT,
                                               wintypes.LPWSTR, wintypes.UINT]
            shell32.DragFinish.argtypes = [wintypes.HANDLE]

            self._drop_hwnd = hwnd
            self._user32 = user32
            self._old_wndproc = user32.GetWindowLongPtrW(hwnd, GWL_WNDPROC)
            shell32.DragAcceptFiles(hwnd, True)

            callback_type = ctypes.WINFUNCTYPE(
                ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
            )

            def wndproc(h, msg, wparam, lparam):
                if msg == WM_DROPFILES:
                    try:
                        files = self._read_dropped_files(wparam)
                    finally:
                        shell32.DragFinish(wparam)
                    if files:
                        self.after(10, lambda: self._handle_dropped_files(files))
                    return 0
                return user32.CallWindowProcW(self._old_wndproc, h, msg, wparam, lparam)

            self._wndproc = callback_type(wndproc)
            user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC, self._wndproc)
        except Exception:  # noqa: BLE001
            self._old_wndproc = None

    @staticmethod
    def _read_dropped_files(hdrop) -> list[str]:
        import ctypes

        shell32 = ctypes.windll.shell32
        total = shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
        files: list[str] = []
        for index in range(total):
            length = shell32.DragQueryFileW(hdrop, index, None, 0)
            buffer = ctypes.create_unicode_buffer(length + 1)
            shell32.DragQueryFileW(hdrop, index, buffer, length + 1)
            if buffer.value:
                files.append(buffer.value)
        return files

    def _handle_dropped_files(self, files: list[str]) -> None:
        for raw in files:
            path = Path(raw)
            suffix = path.suffix.lower()
            if suffix in {".xlsx", ".xlsm"}:
                self._load_excel(path)
            elif suffix == ".docx":
                self._load_template(path)
            else:
                self._log(f"忽略不支持的文件：{path.name}（需要 .xlsx 或 .docx）", "warn")

    def destroy(self) -> None:
        # 还原窗口过程，避免退出时回调已销毁的 Tk 对象
        try:
            if getattr(self, "_old_wndproc", None):
                self._user32.SetWindowLongPtrW(self._drop_hwnd, -4, self._old_wndproc)
                self._old_wndproc = None
        except Exception:  # noqa: BLE001
            pass
        super().destroy()

    def _choose_output(self) -> None:
        initial = str(self.output_dir) if self.output_dir else str(BASE_DIR)
        path = filedialog.askdirectory(title="选择输出文件夹", initialdir=initial)
        if path:
            self.output_dir = Path(path)
            self._update_output_hint()
            self._sync_buttons()

    def _update_output_hint(self) -> None:
        if self.output_dir:
            try:
                shown = self.output_dir.relative_to(BASE_DIR)
                text = f"输出到：{shown}"
            except ValueError:
                text = f"输出到：{short(str(self.output_dir), 40)}"
            self.status_label.configure(text=text, text_color=MUTED)

    def _on_slider(self, value: float) -> None:
        self.winners = int(round(value))
        self.qty_label.configure(text=str(self.winners))
        if self.players:
            self._recalculate(keep_log=True)

    # -------------------------------------------------------------- 字段映射
    def _toggle_mappings(self) -> None:
        self.mapping_open = not getattr(self, "mapping_open", False)
        self._sync_mapping_visibility()

    def _sync_mapping_visibility(self) -> None:
        """没有模板或没有占位符时整块移除，避免 CTkFrame 空置时留下 200px 空白。"""
        has_fields = bool(self.field_rules) and bool(self.template_keys)
        if not has_fields:
            self.mapping_box.pack_forget()
            self.mapping_hint.pack_forget()
            self.mapping_toggle.pack_forget()
            self.mapping_section.pack_forget()
            return
        if not self.mapping_section.winfo_ismapped():
            self.mapping_section.pack(fill="x", after=self.mapping_anchor)
        self.mapping_hint.pack_forget()
        # 先按钮后列表，保证容器内顺序稳定
        if not self.mapping_toggle.winfo_ismapped():
            self.mapping_toggle.pack(fill="x", pady=(2, 0))
        if self.mapping_open:
            if not self.mapping_box.winfo_ismapped():
                self.mapping_box.pack(fill="x", before=self.mapping_toggle)
            self.mapping_toggle.configure(text="收起映射")
        else:
            self.mapping_box.pack_forget()
            self.mapping_toggle.configure(
                text=f"展开映射（{len(self.field_rules)} 个字段）")

    def _column_names(self) -> list[str]:
        if self.players:
            return list(self.players[0].fields.keys())
        return list(self.column_names)

    def _refresh_field_rules(self) -> None:
        """按模板占位符重建映射列表，保留用户已配置过的项。"""
        keys = list(self.template_keys)
        if not keys:
            self.field_rules = []
            self._render_mappings()
            return
        existing = {rule.key: rule for rule in self.field_rules}
        context = CertificateContext(contest=self._guess_contest(), issuer="大赛组委会", date="")
        self.field_rules = field_rules.build_rules(
            keys, self._column_names(), context, existing)
        self._render_mappings()
        if not getattr(self, "mapping_open", True):
            self.mapping_toggle.configure(text=f"展开映射（{len(self.field_rules)} 个字段）")

    def _render_mappings(self) -> None:
        for child in self.mapping_box.winfo_children():
            child.destroy()
        if not self.field_rules or not self.template_keys:
            self._sync_mapping_visibility()
            return
        for rule in self.field_rules:
            row = ctk.CTkFrame(self.mapping_box, fg_color="transparent")
            row.pack(fill="x", pady=1)

            ctk.CTkLabel(row, text=rule.key, width=self._px(52), anchor="w",
                         font=self.font_small, text_color=INK).pack(side="left")
            menu = ctk.CTkOptionMenu(
                row, values=list(field_rules.SOURCE_LABELS.values()),
                width=self._px(76), height=self._px(26), font=self.font_small,
                dropdown_font=self.font_small, corner_radius=7,
                fg_color="#F8FAFD", button_color="#E8ECF5",
                button_hover_color="#DDE3EF", text_color=INK,
                command=lambda value, r=rule: self._on_source_change(r, value),
            )
            menu.set(rule.source_label)
            menu.pack(side="left", padx=(2, 4))

            holder = ctk.CTkFrame(row, fg_color="transparent")
            holder.pack(side="left", fill="x", expand=True)
            self._build_mapping_value(holder, rule)
        self._sync_mapping_visibility()

    def _build_mapping_value(self, holder, rule) -> None:
        height = self._px(26)
        if rule.source == "auto":
            ctk.CTkLabel(holder, text=field_rules.AUTO_FIELDS.get(rule.key, "自动取值"),
                         font=self.font_small, text_color=MUTED, anchor="w",
                         height=height).pack(fill="x")
            return

        if rule.source == "column":
            names = self._column_names() or ["（先选择成绩表）"]
            menu = ctk.CTkOptionMenu(
                holder, values=names, font=self.font_small, dropdown_font=self.font_small,
                height=height, corner_radius=7, fg_color="#F8FAFD",
                button_color="#E8ECF5", button_hover_color="#DDE3EF", text_color=INK,
                command=lambda value, r=rule: self._set_rule_value(r, value),
            )
            if rule.value not in names:
                rule.value = names[0]
            menu.set(rule.value)
            menu.pack(fill="x")
            return

        if rule.source == "system_date":
            ctk.CTkLabel(holder, text="每次生成使用当天", font=self.font_small,
                         text_color=MUTED, anchor="w", height=height).pack(fill="x")
            return

        if rule.source == "fixed_date":
            box = ctk.CTkFrame(holder, fg_color="transparent")
            box.pack(fill="x")
            var = ctk.StringVar(value=rule.value or date.today().isoformat())
            entry = ctk.CTkEntry(box, textvariable=var, height=height,
                                 font=self.font_small, corner_radius=7,
                                 fg_color="#F8FAFD", border_width=1, border_color=LINE)
            entry.pack(side="left", fill="x", expand=True)
            ctk.CTkButton(
                box, text="📅", width=self._px(28), height=height, corner_radius=7,
                fg_color="#F3F5FA", hover_color="#E6EAF3", text_color=INK,
                font=self.font_small,
                command=lambda r=rule, v=var: self._pick_date(r, v),
            ).pack(side="left", padx=(3, 0))
            var.trace_add("write", lambda *_a, r=rule, v=var: self._set_rule_value(r, v.get()))
            return

        # 固定值
        var = ctk.StringVar(value=rule.value)
        entry = ctk.CTkEntry(holder, textvariable=var, height=height,
                             font=self.font_small, corner_radius=7,
                             fg_color="#F8FAFD", border_width=1, border_color=LINE,
                             placeholder_text="填写固定内容")
        entry.pack(fill="x")
        var.trace_add("write", lambda *_a, r=rule, v=var: self._set_rule_value(r, v.get()))

    def _on_source_change(self, rule, label: str) -> None:
        rule.source = field_rules.LABEL_TO_SOURCE.get(label, "auto")
        if rule.source == "column":
            names = self._column_names()
            if rule.value not in names:
                rule.value = names[0] if names else ""
        elif rule.source == "fixed_date" and not rule.value:
            rule.value = date.today().isoformat()
        elif rule.source == "auto":
            rule.value = ""
        self._render_mappings()
        self._save_settings()

    def _set_rule_value(self, rule, value: str) -> None:
        rule.value = str(value)
        self._save_settings()

    def _pick_date(self, rule, var) -> None:
        initial = date.today()
        try:
            initial = datetime.strptime(var.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            pass

        def apply(value):
            var.set(value.isoformat())

        DatePicker(self, initial=initial, on_pick=apply, font=self.font_small)

    # -------------------------------------------------------------- 奖项
    @staticmethod
    def _parse_awards(raw) -> list[dict]:
        result: list[dict] = []
        if isinstance(raw, list):
            for item in raw:
                if not isinstance(item, dict):
                    continue
                name = str(item.get("name", "")).strip()
                if not name:
                    continue
                try:
                    count = int(item.get("count", 1))
                except (TypeError, ValueError):
                    count = 1
                result.append({"name": name, "count": max(1, count)})
        return result or [dict(item) for item in DEFAULT_AWARDS]

    def _award_total_count(self) -> int:
        return sum(max(0, int(item["count"])) for item in self.awards)

    def _sync_award_ui(self) -> None:
        """按开关状态切换「奖项列表」与「仅排名次」两套控件。"""
        on = bool(self.award_switch.get())
        self.awards_panel.pack_forget()
        self.qty_panel.pack_forget()
        if on:
            self.awards_panel.pack(fill="x")
            self._render_awards()
        else:
            self.qty_panel.pack(fill="x")
        self._update_award_total()

    def _update_award_total(self) -> None:
        on = bool(self.award_switch.get())
        if not on:
            self.award_total.configure(text="")
            return
        total = self._award_total_count()
        text = f"共 {total} 份奖状"
        if self.players and total > len(self.players):
            text += f"（实际只有 {len(self.players)} 人，超出部分自动忽略）"
        self.award_total.configure(text=text,
                                   text_color=WARNING if self.players and total > len(self.players)
                                   else MUTED)

    def _render_awards(self) -> None:
        self._drag_item = None
        for child in self.awards_box.winfo_children():
            child.destroy()
        self._award_rows: list[tuple] = []
        self._scale_job = None

        for item in self.awards:
            row = ctk.CTkFrame(self.awards_box, fg_color="transparent")
            row.pack(fill="x", pady=2)

            handle = ctk.CTkLabel(
                row, text="≡", width=self._px(20), height=self._px(30),
                font=self.font_body, text_color="#A9B0C2", cursor="fleur",
            )
            handle.pack(side="left")
            handle.bind("<ButtonPress-1>", lambda event, it=item: self._begin_drag(event, it))

            name_var = ctk.StringVar(value=item["name"])
            count_var = ctk.StringVar(value=str(item["count"]))
            entry = ctk.CTkEntry(
                row, textvariable=name_var, height=self._px(30), corner_radius=8,
                fg_color="#F8FAFD", border_width=1, border_color=LINE,
                font=self.font_small, placeholder_text="奖项名称",
            )
            entry.pack(side="left", fill="x", expand=True, padx=(4, 0))
            count_entry = ctk.CTkEntry(
                row, textvariable=count_var, width=self._px(46), height=self._px(30),
                corner_radius=8, fg_color="#F8FAFD", border_width=1, border_color=LINE,
                font=self.font_small, justify="center",
            )
            count_entry.pack(side="left", padx=(6, 4))
            del_btn_widget = ctk.CTkButton(
                row, text="×", width=self._px(26), height=self._px(28), corner_radius=8,
                fg_color="transparent", hover_color="#FBE9E9", text_color="#B91C1C",
                font=self.font_small, command=lambda it=item: self._remove_award(it),
            )
            del_btn_widget.pack(side="left")

            name_var.trace_add(
                "write",
                lambda *_a, it=item, var=name_var: self._update_award(it, "name", var.get()),
            )
            count_var.trace_add(
                "write",
                lambda *_a, it=item, var=count_var: self._update_award(it, "count", var.get()),
            )
            self._award_rows.append((row, item, (handle, entry, count_entry, del_btn_widget)))

    # ---------------------------------------------------------- 拖动排序
    def _begin_drag(self, event, item: dict) -> None:
        if len(self.awards) < 2:
            return
        self._drag_item = item
        self._drag_y = event.y_root
        self._drag_moved = False
        for row, target, _widgets in self._award_rows:
            if target is item:
                row.configure(fg_color="#EEF0FF")
        self.bind_all("<B1-Motion>", self._on_drag_move)
        self.bind_all("<ButtonRelease-1>", self._end_drag)

    def _row_height(self) -> int:
        if self._award_rows:
            return max(self._px(34), self._award_rows[0][0].winfo_height() + 4)
        return self._px(34)

    def _on_drag_move(self, event) -> None:
        if self._drag_item is None:
            return
        steps = int((event.y_root - self._drag_y) / self._row_height())
        if steps and self._shift_award(self._drag_item, steps):
            self._drag_y += steps * self._row_height()
            self._drag_moved = True

    def _end_drag(self, event=None) -> None:
        try:
            self.unbind_all("<B1-Motion>")
            self.unbind_all("<ButtonRelease-1>")
        except Exception:  # noqa: BLE001
            pass
        for row, _target, _widgets in self._award_rows:
            try:
                row.configure(fg_color="transparent")
            except Exception:  # noqa: BLE001
                pass
        moved = getattr(self, "_drag_moved", False)
        self._drag_item = None
        if moved:
            self._save_settings()
            if self.players:
                self._recalculate(keep_log=True)
            self._log("奖项顺序已更新：" + " → ".join(it["name"] for it in self.awards), "muted")

    def _shift_award(self, item: dict, steps: int) -> bool:
        """把某个奖项上移/下移若干位。"""
        index = next((i for i, it in enumerate(self.awards) if it is item), None)
        if index is None:
            return False
        target = max(0, min(len(self.awards) - 1, index + steps))
        if target == index:
            return False
        self.awards.insert(target, self.awards.pop(index))
        self._apply_award_order()
        self._update_award_total()
        if self.players:
            self._recalculate(keep_log=True)
        return True

    def _apply_award_order(self) -> None:
        """只调整行控件顺序，不重建，避免拖动时输入框失焦。"""
        frames = {id(item): (frame, widgets)
                  for frame, item, widgets in self._award_rows}
        for frame, _item, _widgets in self._award_rows:
            frame.pack_forget()
        ordered = []
        for item in self.awards:
            entry = frames.get(id(item))
            if entry is None:
                continue
            frame, widgets = entry
            frame.pack(fill="x", pady=2)
            ordered.append((frame, item, widgets))
        if ordered:
            self._award_rows = ordered

    def _update_award(self, item: dict, field: str, value: str) -> None:
        if not any(it is item for it in self.awards):
            return
        if field == "name":
            item["name"] = value
        else:
            try:
                item["count"] = max(0, int(value))
            except ValueError:
                item["count"] = 0
        self._update_award_total()
        self._save_settings()
        if self.players:
            self._recalculate(keep_log=True)

    def _add_award(self) -> None:
        self.awards.append({"name": f"奖项 {len(self.awards) + 1}", "count": 1})
        self.award_switch.select()
        self.award_mode = True
        self._sync_award_ui()
        self._save_settings()
        if self.players:
            self._recalculate(keep_log=True)

    def _remove_award(self, item: dict) -> None:
        self.awards = [it for it in self.awards if it is not item]
        self._sync_award_ui()
        self._save_settings()
        if self.players:
            self._recalculate(keep_log=True)

    def _on_award_toggle(self) -> None:
        self.award_mode = bool(self.award_switch.get())
        self._sync_award_ui()
        self._save_settings()
        if self.players:
            self._recalculate(keep_log=True)
        self._log("奖项等级：开启" if self.award_mode else "奖项等级：关闭，只填名次", "muted")

    def _recalculate(self, keep_log: bool = False) -> None:
        if not self.players:
            return
        tiers: list[tuple[str, int]] = []
        if bool(self.award_switch.get()):
            tiers = [(item["name"].strip() or "奖项", max(0, int(item["count"])))
                     for item in self.awards]
        quota = sum(count for _name, count in tiers) if tiers else self.winners
        self.quota = quota
        self.entries = rank_players(
            self.players,
            winners=quota,
            use_tiers=bool(tiers),
            award_tiers=tiers,
        )
        self._refresh_table()
        self._refresh_stats()
        self._update_award_warning()
        if not keep_log:
            top = self.entries[0] if self.entries else None
            if top:
                self._log(f"最高分 {top.player.score:g} · {top.player.name}", "info")

    def _column_has_values(self, key: str) -> bool:
        for player in self.players:
            if str(player.fields.get(key, "")).strip():
                return True
        return False

    def _table_column_specs(self) -> list[tuple]:
        """列由成绩表的实际字段决定：名次 + 原始各列 + 总分 + 奖项。"""
        specs: list[tuple] = [("rank", "名次", 64, "center")]
        if self.players:
            for key in self.players[0].fields:
                if key == "分数":
                    continue
                # 整列都没有值的（比如表里没有编号）就不显示
                if key != "姓名" and not self._column_has_values(key):
                    continue
                width = {"姓名": 110, "编号": 90}.get(key, 130)
                if key.startswith("评委"):
                    width = 84
                anchor = "center" if key in ("编号",) or key.startswith("评委") else "w"
                specs.append((key, key, width, anchor))
            specs.append(("score", "总分", 84, "center"))
        else:
            specs += [("name", "姓名", 110, "w"), ("no", "编号", 90, "center"),
                      ("score", "总分", 84, "center")]
        specs.append(("award", "奖项", 96, "center"))
        return specs

    def _build_table(self) -> None:
        """按当前字段重建排名表（列会随成绩表变化）。"""
        holder = self.table_holder
        for child in list(holder.winfo_children()):
            if child is self.empty_panel:
                continue
            child.destroy()

        self._table_columns = self._table_column_specs()
        keys = [spec[0] for spec in self._table_columns]
        self.tree = ttk.Treeview(holder, columns=keys, show="headings", height=6,
                                 style="Award.Treeview", selectmode="browse")
        for key, text, width, anchor in self._table_columns:
            self.tree.heading(key, text=text, anchor=anchor)
            # 文本类列与奖项列承担富余宽度，数值列保持紧凑
            stretch = anchor == "w" or key == "award"
            self.tree.column(key, width=width, minwidth=48, anchor=anchor, stretch=stretch)
        self.tree.grid(row=0, column=0, sticky="nsew")

        self.table_scrollbar = ctk.CTkScrollbar(
            holder, command=self.tree.yview, fg_color="transparent",
            button_color="#D4D9E4", button_hover_color="#B9C0D0", width=10)
        self.table_scrollbar.grid(row=0, column=1, sticky="ns", padx=(2, 4))
        self.tree.configure(yscrollcommand=self.table_scrollbar.set)

        self.tree.tag_configure("first", background="#FFF8E6", foreground="#8A5A00")
        self.tree.tag_configure("second", background="#F4F7FB", foreground="#3A4A5F")
        self.tree.tag_configure("third", background="#FDF2EC", foreground="#8A4A1F")
        self.tree.tag_configure("winner", background="#F7F5FF", foreground="#312E81")
        self.tree.tag_configure("normal", background=CARD, foreground=MUTED)
        self.tree.bind("<Double-1>", self._open_certificate)
        self._fit_columns()

    def _show_empty_state(self, empty: bool) -> None:
        """没有数据时用占位提示替代空表格，避免看起来像残留。"""
        try:
            if empty:
                self.tree.grid_remove()
                self.table_scrollbar.grid_remove()
                self.empty_panel.grid()
            else:
                self.empty_panel.grid_remove()
                self.tree.grid()
                self.table_scrollbar.grid()
        except Exception:  # noqa: BLE001
            pass

    def _refresh_table(self) -> None:
        if not self.entries:
            self._show_empty_state(True)
            return
        self._show_empty_state(False)
        self.tree.delete(*self.tree.get_children())
        for index, entry in enumerate(self.entries):
            fields = entry.player.fields
            award_text = (entry.award or "—") if entry.is_winner else "—"
            values = [entry.rank]
            for key, _text, _width, _anchor in self._table_columns[1:-1]:
                if key == "score":
                    values.append(f"{entry.player.score:g}")
                else:
                    values.append(fields.get(key, "") or "—")
            values.append(award_text)
            if entry.is_winner and index == 0:
                tag = "first"
            elif entry.is_winner and index == 1:
                tag = "second"
            elif entry.is_winner and index == 2:
                tag = "third"
            elif entry.is_winner:
                tag = "winner"
            else:
                tag = "normal"
            self.tree.insert("", "end", values=values, tags=(tag,))

    def _update_award_warning(self) -> None:
        """同分并列导致实际发奖人数与设置不一致时提醒用户改人数。"""
        warn = getattr(self, "award_warn", None)
        if warn is None:
            return
        if not self.award_switch.get() or not self.entries:
            warn.pack_forget()
            warn.configure(text="")
            self._last_conflict = ""
            return
        tiers = [(item["name"], item["count"]) for item in self.awards]
        conflicts = award_conflicts(self.entries, tiers)
        if not conflicts:
            warn.pack_forget()
            warn.configure(text="")
            self._last_conflict = ""
            return

        parts = []
        for item in conflicts:
            name, configured, actual = item["name"], item["configured"], item["actual"]
            if actual > configured:
                parts.append(f"「{name}」设了 {configured} 名，但有 {actual} 人同分并列")
            else:
                parts.append(f"「{name}」设了 {configured} 名，实际只能发 {actual} 名")
        warn.configure(text="；".join(parts) + "。建议按实际人数调整上面的奖项设置。",
                       wraplength=self._px(252))
        if not warn.winfo_ismapped():
            warn.pack(fill="x", pady=(6, 0))

        signature = "|".join(f"{c['name']}:{c['actual']}" for c in conflicts)
        if signature != getattr(self, "_last_conflict", ""):
            self._last_conflict = signature
            self._log("；".join(parts) + "，建议调整奖项人数", "warn")

    def _refresh_stats(self) -> None:
        if not self.players:
            for label in self.stat_values.values():
                label.configure(text="—")
            return
        scores = [p.score for p in self.players]
        winners = winners_of(self.entries)
        self.stat_values["total"].configure(text=str(len(self.players)))
        self.stat_values["max"].configure(text=f"{max(scores):g}")
        self.stat_values["avg"].configure(text=f"{sum(scores) / len(scores):.1f}")
        self.stat_values["awards"].configure(text=str(len(winners)))
        if bool(self.award_switch.get()):
            tiers: dict[str, int] = {}
            for entry in winners:
                key = entry.award or "未定奖项"
                tiers[key] = tiers.get(key, 0) + 1
            parts = " · ".join(f"{name} {count} 人" for name, count in tiers.items())
            self.merge_hint.configure(text=parts)
        else:
            self.merge_hint.configure(text=f"仅排名次 · 前 {len(winners)} 名")

    def _sync_buttons(self) -> None:
        ready = bool(self.excel_path and self.template_path and not self.busy)
        self.generate_btn.configure(state="normal" if ready else "disabled",
                                    fg_color=PRIMARY if ready else "#C7CBDA")

    # -------------------------------------------------------------- 生成
    def _start_generate(self) -> None:
        if self.busy:
            return
        if not self.excel_path or not self.players:
            messagebox.showinfo("还差一步", "请先选择 Excel 成绩表。")
            return
        if not self.template_path:
            messagebox.showinfo("还差一步", "请先选择 Word 奖状模板。")
            return
        winners = winners_of(self.entries)
        if not winners:
            messagebox.showinfo("没有获奖者", "请设置大于 0 的获奖名额。")
            return

        context = CertificateContext(
            contest=self._guess_contest(),
            issuer="大赛组委会",
            date=date.today().strftime("%Y 年 %m 月 %d 日"),
            extra={},
        )
        self.busy = True
        self.generate_btn.configure(state="disabled", text="正在生成…", fg_color="#9CA3AF")
        self.status_label.configure(text=f"正在生成 {len(winners)} 份奖状…", text_color=PRIMARY)
        self._log(f"开始生成 {len(winners)} 份奖状 → {self.output_dir}", "info")

        threading.Thread(
            target=self._worker,
            args=(self.template_path, self.output_dir, winners, context,
                  bool(self.merge_check.get()), list(self.field_rules)),
            daemon=True,
        ).start()
        self.after(80, self._poll)

    def _guess_contest(self) -> str:
        if self.excel_title:
            head = re.split(r"[·•|｜\-—_:：]", self.excel_title)[0].strip()
            if len(head) >= 4:
                return head
        if self.excel_path:
            stem = self.excel_path.stem
            for suffix in ("决赛评分表", "评分表", "成绩表", "打分表"):
                if stem.endswith(suffix):
                    cleaned = stem[: -len(suffix)].strip(" -_·")
                    if cleaned:
                        return cleaned
            if stem:
                return stem
        return "技能竞赛"

    def _worker(self, template: Path, out_dir: Path, winners, context, merge: bool,
                rules=None) -> None:
        created: list[Path] = []
        failures: list[str] = []
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            for entry in winners:
                path = out_dir / certificate_filename(entry)
                try:
                    generate_certificate(template, path, entry, context, rules=rules)
                    created.append(path)
                    self.ui_queue.put(("log", f"{path.name} 已生成", "ok"))
                except CertificateError as exc:
                    failures.append(f"{entry.player.name}：{exc}")
                    self.ui_queue.put(("log", f"{entry.player.name} 生成失败：{exc}", "err"))

            if merge and created:
                merged = out_dir / "全部奖状_合并版.docx"
                try:
                    merge_documents(created, merged)
                    created.append(merged)
                    self.ui_queue.put(("log", "合并版已生成", "ok"))
                except Exception as exc:  # noqa: BLE001
                    self.ui_queue.put(("log", f"合并版失败（单人文件正常）：{exc}", "warn"))

            self.ui_queue.put(("done", created, failures, out_dir))
        except Exception as exc:  # noqa: BLE001
            self.ui_queue.put(("fatal", str(exc)))

    def _poll(self) -> None:
        try:
            while True:
                message = self.ui_queue.get_nowait()
                if message[0] == "log":
                    self._log(message[1], message[2])
                elif message[0] == "done":
                    _, created, failures, out_dir = message
                    self._finish(created, failures, out_dir)
                    return
                elif message[0] == "fatal":
                    self.busy = False
                    self.generate_btn.configure(text="生成奖状", fg_color=PRIMARY)
                    self._sync_buttons()
                    self.status_label.configure(text="生成失败", text_color=DANGER)
                    self._log(f"生成失败：{message[1]}", "err")
                    messagebox.showerror("生成失败", message[1])
                    return
        except queue.Empty:
            pass
        self.after(80, self._poll)

    def _finish(self, created: list[Path], failures: list[str], out_dir: Path) -> None:
        self.busy = False
        self.generated = created
        self.generate_btn.configure(text="生成奖状", fg_color=PRIMARY)
        self._sync_buttons()
        singles = [p for p in created if "合并" not in p.name]
        if failures:
            self.status_label.configure(text=f"完成 {len(singles)} 份，{len(failures)} 份失败",
                                        text_color=WARNING)
            self._log(f"完成：{len(singles)} 份成功，{len(failures)} 份失败", "warn")
            messagebox.showwarning("部分失败", "\n".join(failures[:6]))
        else:
            self.status_label.configure(text=f"已完成 {len(singles)} 份奖状", text_color=SUCCESS)
            self._log(f"全部完成：{len(singles)} 份奖状 → {out_dir}", "ok")
        if created and messagebox.askyesno("生成完成", f"共生成 {len(created)} 个文件。\n现在打开输出文件夹？"):
            self._open_output()

    def _open_output(self) -> None:
        target = self.output_dir or (self.generated[0].parent if self.generated else None)
        if not target or not Path(target).exists():
            messagebox.showinfo("暂无输出", "还没有生成过奖状，先选择文件并生成吧。")
            return
        try:
            os.startfile(str(target))  # noqa: S606
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("打不开文件夹", str(exc))

    def _open_certificate(self, event) -> None:
        row = self.tree.identify_row(event.y)
        if not row:
            return
        values = self.tree.item(row, "values")
        if not values or not self.generated:
            return
        name = values[1]
        for path in self.generated:
            if name in path.name:
                try:
                    os.startfile(str(path))  # noqa: S606
                except Exception as exc:  # noqa: BLE001
                    messagebox.showerror("打不开文件", str(exc))
                return
        messagebox.showinfo("还没有奖状", f"「{name}」的奖状尚未生成。")

    # -------------------------------------------------------------- 其它
    def _log(self, text: str, level: str = "info") -> None:
        self.log_box.configure(state="normal")
        self.log_box._textbox.insert("end", text + "\n", level)
        self.log_box._textbox.see("end")
        self.log_box.configure(state="disabled")

    def _show_help(self) -> None:
        text = (
            "使用步骤\n"
            "1. 选择 Excel 成绩表：软件自动找表头、识别姓名列与分数列，并立即按分数排序。\n"
            "2. 选择 Word 奖状模板：在需要填字的位置写 {{姓名}}、{{名次}}，可加 {{奖项}}、"
            "{{比赛名称}}、{{日期}}、{{编号}}，以及表格里的任意列（如 {{班级}}）。\n"
            "3. 字段映射：模板里每个占位符都会列出来，逐行指定取值方式——\n"
            "   · 自动：姓名、名次、奖项、编号由软件算好\n"
            "   · 表格列：从成绩表里挑一列（如把 {{班级}} 映射到「班级」列）\n"
            "   · 固定值：自己填，比如主办方名称、比赛全称\n"
            "   · 日期：用当天系统日期，或点日历选一个具体日期\n"
            "4. 设置奖项：可自定义奖项名称和人数（如「特等奖 1 名、一等奖 3 名」），"
            "点「＋ 添加奖项」继续加，按住行首的 ≡ 可以拖动排序；"
            "关掉「设置奖项等级」则只排名次，奖状里不填奖项。\n"
            "5. 选择输出目录，默认放在成绩表同级的 output/奖状，可勾选额外生成合并版。\n"
            "6. 点「生成奖状」，每人一份 Word，文件名形如 01_张三_一等奖.docx。\n\n"
            "界面字号\n"
            "右上角中间的数字可以直接改：点进去输入 10–28 之间的数值，回车生效；"
            "也可以点 A- / A+ 逐级调整，或用键盘上下键微调。\n"
            "调字号不会改变窗口大小，超出空间的文字会省略、左侧设置区可滚动；"
            "字号较大时建议先最大化窗口，表格列会自动铺开。\n\n"
            "小技巧\n"
            "· 成绩表和奖状模板都可以直接从资源管理器拖进窗口，不必点按钮选路径。\n"
            "· 表格前几行有标题、说明没关系，表头会在前 20 行内自动定位。\n"
            "· 没有「总分」列但有多位评委列时，分数自动相加。\n"
            "· 双击排名表里的某一行，可以直接打开已生成的那份奖状。\n"
            "· 排名表的列跟着成绩表走：表里有什么列就显示什么，整列没值的字段会自动省略。\n"
            "· 同分的选手一定拿同一个奖。如果同分人数超过某个奖项的名额，整组都会按该奖项发放，"
            "并在奖项设置下方提示你按实际人数改一下名额。\n"
            "· 映射面板占地方，点「收起映射」可以折起来，需要时再展开。"
        )
        messagebox.showinfo("使用说明", text)

    def _load_demo(self) -> None:
        excel = DATA_DIR / "samples" / "决赛评分表.xlsx"
        template = DATA_DIR / "samples" / "奖状模板.docx"
        if excel.exists():
            self._load_excel(excel)
            self.output_dir = DATA_DIR / "output" / "奖状"
            self._update_output_hint()
        if template.exists():
            self._load_template(template)
        self._sync_buttons()


def main() -> None:
    demo = "--demo" in sys.argv or "-d" in sys.argv
    app = AwardStudio(demo=demo)
    app.mainloop()


if __name__ == "__main__":
    main()
