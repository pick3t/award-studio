"""Award Studio · 比赛评分排名与奖状批量生成

运行：
    python app.py            正常启动
    python app.py --demo     启动并自动载入 samples 里的测试数据
"""

from __future__ import annotations

import os
import queue
import re
import sys
import threading
from datetime import date
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
from core import settings as prefs  # noqa: E402
from core.excel_reader import ExcelParseError, load_players  # noqa: E402
from core.ranking import rank_players, winners_of  # noqa: E402

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
        self.awards = self._parse_awards(saved.get("awards"))
        self.award_mode = bool(saved.get("award_mode", True))

        self._center(int(WINDOW_W * self.text_scale), int(WINDOW_H * self.text_scale))
        self.minsize(1000, 620)

        self.font_h1 = self._mkfont("h1", 19, "bold")
        self.font_sub = self._mkfont("sub", 11)
        self.font_step = self._mkfont("step", 12, "bold")
        self.font_body = self._mkfont("body", 12)
        self.font_small = self._mkfont("small", 11)
        self.font_num = self._mkfont("num", 22, "bold", "Segoe UI")

        self._setup_style()
        self._build()
        if self.text_scale != 1.0:
            self._apply_text_scale()
        self._log("就绪 · 选择成绩表后会自动按分数排名", "muted")
        self._sync_buttons()

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
        self.scale_label = ctk.CTkLabel(font_box, text=f"{self.font_size} pt",
                                        font=self.font_small, text_color=MUTED, width=52)
        self.scale_label.pack(side="left")
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
    def _mkfont(self, key: str, size: int, weight: str = "normal", family: str = FONT):
        """按当前字号倍率创建字体，并登记以便统一调整。"""
        font = ctk.CTkFont(family=family, size=max(8, round(size * self.text_scale)),
                           weight=weight)
        self._fonts[key] = (font, size, weight, family)
        return font

    def _px(self, value: float) -> int:
        """把设计尺寸换算成当前字号下的实际像素。"""
        return max(1, int(round(value * self.text_scale)))

    def _apply_text_scale(self) -> None:
        for font, size, weight, family in self._fonts.values():
            try:
                font.configure(size=max(8, round(size * self.text_scale)))
            except Exception:  # noqa: BLE001
                pass
        self._restyle_table()
        self._resize_layout()

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
        """字号变化后同步左栏宽度、控件高度与窗口尺寸。"""
        px = self._px
        try:
            self.left_panel.configure(width=px(352))
            self.left_scroll.configure(width=px(322))
            self.main_frame.grid_columnconfigure(0, minsize=px(352))
            self.header_frame.configure(height=px(72))
            self.badge.configure(width=px(44), height=px(44))
            for button, height in (
                (self.excel_btn, 40), (self.template_btn, 40), (self.output_btn, 38),
                (self.add_award_btn, 28), (self.generate_btn, 46),
            ):
                button.configure(height=px(height))
            self.log_card.configure(height=px(140))
            self.log_box.configure(height=px(100))
            for title_label, title_h, sub_label, sub_h in self._step_labels:
                title_label.configure(height=px(title_h))
                sub_label.configure(height=px(sub_h))
            for key, base in (
                ("rank", 70), ("name", 110), ("no", 90), ("class", 120),
                ("project", 220), ("score", 90), ("award", 90),
            ):
                self.tree.column(key, width=px(base))
        except Exception:  # noqa: BLE001 - 控件还没建好时跳过
            return
        if self.award_switch.get():
            self._render_awards()

    def _change_scale(self, delta: int) -> None:
        size = max(FONT_SIZE_MIN, min(FONT_SIZE_MAX, self.font_size + delta))
        if size == self.font_size:
            return
        self.font_size = size
        self.text_scale = size / FONT_SIZE_DEFAULT
        self._apply_text_scale()
        self.scale_label.configure(text=f"{size} pt")
        self._save_settings()
        self._center(int(WINDOW_W * self.text_scale), int(WINDOW_H * self.text_scale))
        self._log(f"界面字号：{size} pt（可在右上角 A- / A+ 继续调整）", "muted")

    def _save_settings(self) -> None:
        data = prefs.load()
        data["font_size"] = self.font_size
        data["award_mode"] = bool(self.award_switch.get())
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
        self.excel_hint = ctk.CTkLabel(wrap, text="支持 .xlsx / .xlsm", font=self.font_small,
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
        self.template_hint = ctk.CTkLabel(wrap, text="尚未选择模板", font=self.font_small,
                                          text_color=MUTED, anchor="w", justify="left",
                                          wraplength=286)
        self.template_hint.pack(fill="x")

        # ③ 奖项
        self._step_title(wrap, "③", "奖项设置", "自定义奖项名称与人数")

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

        # ④ 输出
        self._step_title(wrap, "④", "输出位置", "每人生成一份，可选合并版")
        self.output_btn = ctk.CTkButton(
            wrap, text="选择输出文件夹", height=38, corner_radius=10,
            fg_color="#F8FAFD", hover_color="#EFF3FA", text_color=INK,
            border_width=1, border_color=LINE, font=self.font_small,
            command=self._choose_output,
        )
        self.output_btn.pack(fill="x", pady=(8, 4))
        self.merge_check = ctk.CTkCheckBox(
            wrap, text="同时生成合并版（一份文件多页）", font=self.font_small,
            text_color=MUTED, checkbox_width=18, checkbox_height=18, corner_radius=5,
            fg_color=PRIMARY, hover_color=PRIMARY_HOVER, border_color="#C7CCDA",
        )
        self.merge_check.pack(anchor="w")
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
        sub_label = ctk.CTkLabel(text, text=subtitle, font=self._mkfont("step_sub", 10),
                                 text_color=MUTED, height=self._px(16), anchor="w")
        sub_label.pack(fill="x")
        self._step_labels.append((title_label, 19, sub_label, 16))

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
        self.merge_hint = ctk.CTkLabel(head, text="", font=self.font_small, text_color=PRIMARY,
                                       anchor="e")
        self.merge_hint.grid(row=0, column=1, rowspan=2, sticky="e")

        ctk.CTkFrame(table_card, height=1, fg_color=LINE).grid(row=1, column=0, sticky="ew",
                                                              padx=18, pady=(12, 0))

        holder = ctk.CTkFrame(table_card, fg_color="transparent")
        holder.grid(row=2, column=0, sticky="nsew", padx=(14, 8), pady=(6, 12))
        holder.grid_columnconfigure(0, weight=1)
        holder.grid_rowconfigure(0, weight=1)

        columns = ("rank", "name", "no", "class", "project", "score", "award")
        self.tree = ttk.Treeview(holder, columns=columns, show="headings", height=6,
                                 style="Award.Treeview", selectmode="browse")
        headings = {
            "rank": ("名次", 70, "center"),
            "name": ("姓名", 110, "w"),
            "no": ("编号", 90, "center"),
            "class": ("班级", 120, "w"),
            "project": ("作品名称", 220, "w"),
            "score": ("总分", 90, "center"),
            "award": ("奖项", 90, "center"),
        }
        for key, (text, width, anchor) in headings.items():
            self.tree.heading(key, text=text, anchor=anchor)
            self.tree.column(key, width=width, anchor=anchor,
                             stretch=key in ("project", "class"))
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ctk.CTkScrollbar(holder, command=self.tree.yview,
                                     fg_color="transparent", button_color="#D4D9E4",
                                     button_hover_color="#B9C0D0", width=10)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(2, 4))
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.tag_configure("first", background="#FFF8E6", foreground="#8A5A00")
        self.tree.tag_configure("second", background="#F4F7FB", foreground="#3A4A5F")
        self.tree.tag_configure("third", background="#FDF2EC", foreground="#8A4A1F")
        self.tree.tag_configure("winner", background="#F7F5FF", foreground="#312E81")
        self.tree.tag_configure("normal", background=CARD, foreground=MUTED)
        self.tree.bind("<Double-1>", self._open_certificate)

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

        self.max_winners = max(1, min(len(players), 50))
        self.slider.configure(to=max(1, min(20, self.max_winners)),
                              number_of_steps=max(1, min(20, self.max_winners) - 1))
        self.winners = min(self.winners, self.max_winners)
        self.slider.set(self.winners)
        self.qty_label.configure(text=str(self.winners))
        self._recalculate()

    def _choose_template(self) -> None:
        path = filedialog.askopenfilename(
            title="选择奖状模板",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
            initialdir=str(self.template_path.parent) if self.template_path else str(BASE_DIR),
        )
        if not path:
            return
        template = Path(path)
        try:
            keys = inspect_template(template)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("模板无法读取", f"{exc}\n\n请确认是 .docx 文件且没有加密。")
            return
        self.template_path = template
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
        self._sync_buttons()

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
        for child in self.awards_box.winfo_children():
            child.destroy()
        for index, item in enumerate(self.awards):
            row = ctk.CTkFrame(self.awards_box, fg_color="transparent")
            row.pack(fill="x", pady=2)

            name_var = ctk.StringVar(value=item["name"])
            count_var = ctk.StringVar(value=str(item["count"]))
            entry = ctk.CTkEntry(
                row, textvariable=name_var, height=self._px(30), corner_radius=8,
                fg_color="#F8FAFD", border_width=1, border_color=LINE,
                font=self.font_small, placeholder_text="奖项名称",
            )
            entry.pack(side="left", fill="x", expand=True)
            count_entry = ctk.CTkEntry(
                row, textvariable=count_var, width=self._px(46), height=self._px(30),
                corner_radius=8, fg_color="#F8FAFD", border_width=1, border_color=LINE,
                font=self.font_small, justify="center",
            )
            count_entry.pack(side="left", padx=(6, 4))
            ctk.CTkButton(
                row, text="×", width=self._px(26), height=self._px(28), corner_radius=8,
                fg_color="transparent", hover_color="#FBE9E9", text_color="#B91C1C",
                font=self.font_small, command=lambda i=index: self._remove_award(i),
            ).pack(side="left")

            name_var.trace_add(
                "write",
                lambda *_a, i=index, var=name_var: self._update_award(i, "name", var.get()),
            )
            count_var.trace_add(
                "write",
                lambda *_a, i=index, var=count_var: self._update_award(i, "count", var.get()),
            )

    def _update_award(self, index: int, field: str, value: str) -> None:
        if index >= len(self.awards):
            return
        if field == "name":
            self.awards[index]["name"] = value
        else:
            try:
                number = int(value)
            except ValueError:
                number = 0
            self.awards[index]["count"] = max(0, number)
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

    def _remove_award(self, index: int) -> None:
        if 0 <= index < len(self.awards):
            self.awards.pop(index)
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
        if not keep_log:
            top = self.entries[0] if self.entries else None
            if top:
                self._log(f"最高分 {top.player.score:g} · {top.player.name}", "info")

    def _refresh_table(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for index, entry in enumerate(self.entries):
            extras = entry.player.extras
            award_text = (entry.award or "—") if entry.is_winner else "—"
            values = (
                entry.rank,
                entry.player.name,
                entry.player.no or "—",
                extras.get("班级", "—"),
                extras.get("作品名称", "—"),
                f"{entry.player.score:g}",
                award_text,
            )
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
                  bool(self.merge_check.get())),
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

    def _worker(self, template: Path, out_dir: Path, winners, context, merge: bool) -> None:
        created: list[Path] = []
        failures: list[str] = []
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            for entry in winners:
                path = out_dir / certificate_filename(entry)
                try:
                    generate_certificate(template, path, entry, context)
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
            "3. 设置奖项：可自定义奖项名称和人数（如「特等奖 1 名、一等奖 3 名」），"
            "点「＋ 添加奖项」继续加；关掉「设置奖项等级」则只排名次，奖状里不填奖项。\n"
            "4. 点「生成奖状」，每人一份 Word，文件名形如 01_张三_一等奖.docx。\n\n"
            "界面字号\n"
            "右上角的 A- / A+ 调整界面字号，范围 10–28 pt（默认 12 pt）。"
            "窗口和控件会跟着一起放大，超出屏幕时自动贴合屏幕，左侧设置区可上下滚动。"
            "选择会自动记住，下次打开保持这个字号。\n\n"
            "小技巧\n"
            "· 表格前几行有标题、说明没关系，表头会在前 20 行内自动定位。\n"
            "· 没有「总分」列但有多位评委列时，分数自动相加。\n"
            "· 双击排名表里的某一行，可以直接打开已生成的那份奖状。\n"
            "· 奖项按人数依次往下分，不按名次号，所以并列时也不会多发。"
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
            self.template_path = template
            self.template_btn.configure(text=short(template.name, 24))
            try:
                keys = inspect_template(template)
                self.template_hint.configure(text="占位符：" + "、".join(sorted(keys)))
                self._log(f"模板已载入，可用字段：{'、'.join(sorted(keys))}", "ok")
            except Exception:  # noqa: BLE001
                pass
        self._sync_buttons()


def main() -> None:
    demo = "--demo" in sys.argv or "-d" in sys.argv
    app = AwardStudio(demo=demo)
    app.mainloop()


if __name__ == "__main__":
    main()
