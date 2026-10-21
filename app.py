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
        self._center(1300, 860)
        self.minsize(1100, 640)

        # 状态
        self.excel_path: Path | None = None
        self.template_path: Path | None = None
        self.output_dir: Path | None = None
        self.players = []
        self.entries = []
        self.winners = 5
        self.max_winners = 20
        self.excel_title = ""
        self.ui_queue: queue.Queue = queue.Queue()
        self.generated: list[Path] = []
        self.busy = False

        self.font_h1 = ctk.CTkFont(FONT, 19, "bold")
        self.font_sub = ctk.CTkFont(FONT, 11)
        self.font_step = ctk.CTkFont(FONT, 12, "bold")
        self.font_body = ctk.CTkFont(FONT, 12)
        self.font_small = ctk.CTkFont(FONT, 11)
        self.font_num = ctk.CTkFont("Segoe UI", 22, "bold")

        self._setup_style()
        self._build()
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
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(width, screen_w - 80)
        height = min(height, screen_h - 100)
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2 - 20)
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
            font=(FONT, 10),
        )
        style.configure(
            "Award.Treeview.Heading",
            background="#F8FAFD",
            foreground=MUTED,
            font=(FONT, 9, "bold"),
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

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_header()

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 14))
        main.grid_columnconfigure(0, weight=0, minsize=352)
        main.grid_columnconfigure(1, weight=1)
        main.grid_rowconfigure(0, weight=1)

        self._build_left(main)
        self._build_right(main)

    def _build_header(self) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent", height=72)
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(12, 8))
        header.grid_columnconfigure(1, weight=1)

        badge = ctk.CTkFrame(header, width=44, height=44, corner_radius=13, fg_color=PRIMARY)
        badge.grid(row=0, column=0, rowspan=2, sticky="w")
        badge.grid_propagate(False)
        ctk.CTkLabel(badge, text="A", font=ctk.CTkFont("Segoe UI", 22, "bold"),
                     text_color="#FFFFFF").place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(header, text="Award Studio", font=self.font_h1, text_color=INK,
                     anchor="w").grid(row=0, column=1, sticky="w", padx=(14, 0))
        ctk.CTkLabel(header, text="比赛评分排名 · 奖状批量生成", font=self.font_sub,
                     text_color=MUTED, anchor="w").grid(row=1, column=1, sticky="w", padx=(14, 0))

        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.grid(row=0, column=2, rowspan=2, sticky="e")
        ctk.CTkButton(actions, text="使用说明", width=92, height=34, corner_radius=9,
                      fg_color=CARD, hover_color="#F1F4FA", text_color=INK,
                      border_width=1, border_color=LINE, font=self.font_small,
                      command=self._show_help).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(actions, text="打开输出目录", width=112, height=34, corner_radius=9,
                      fg_color=CARD, hover_color="#F1F4FA", text_color=INK,
                      border_width=1, border_color=LINE, font=self.font_small,
                      command=self._open_output).grid(row=0, column=1)

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
            font=ctk.CTkFont(FONT, 14, "bold"), command=self._start_generate,
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

        # ③ 名额
        self._step_title(wrap, "③", "获奖名额", "前 N 名自动生成奖状")
        qty = ctk.CTkFrame(wrap, fg_color="transparent")
        qty.pack(fill="x", pady=(8, 0))
        self.qty_label = ctk.CTkLabel(qty, text="5", font=ctk.CTkFont("Segoe UI", 20, "bold"),
                                      text_color=PRIMARY, width=46, anchor="w")
        self.qty_label.pack(side="left")
        ctk.CTkLabel(qty, text="名", font=self.font_small, text_color=MUTED).pack(side="left",
                                                                                padx=(0, 8))
        self.slider = ctk.CTkSlider(
            qty, from_=1, to=20, number_of_steps=19, height=18,
            fg_color="#E6E9F2", progress_color=PRIMARY, button_color=PRIMARY,
            button_hover_color=PRIMARY_HOVER, command=self._on_slider,
        )
        self.slider.set(5)
        self.slider.pack(side="left", fill="x", expand=True)

        self.tier_switch = ctk.CTkSwitch(
            wrap, text="按名次自动划分一 / 二 / 三等奖", font=self.font_small,
            text_color=MUTED, progress_color=PRIMARY, command=self._on_tier_toggle,
        )
        self.tier_switch.select()
        self.tier_switch.pack(anchor="w", pady=(4, 0))

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
        ctk.CTkLabel(text, text=title, font=ctk.CTkFont(FONT, 12, "bold"), text_color=INK,
                     anchor="w").pack(fill="x")
        ctk.CTkLabel(text, text=subtitle, font=ctk.CTkFont(FONT, 10), text_color=MUTED,
                     anchor="w").pack(fill="x")

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
        ctk.CTkLabel(head, text="排名总览", font=ctk.CTkFont(FONT, 14, "bold"),
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
        log_card = ctk.CTkFrame(right, corner_radius=16, fg_color=CARD,
                                border_width=1, border_color=LINE, height=140)
        log_card.pack(fill="x", pady=(12, 0))
        log_card.grid_propagate(False)
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(log_card, text="运行日志", font=ctk.CTkFont(FONT, 11, "bold"),
                     text_color=MUTED, anchor="w").grid(row=0, column=0, sticky="w",
                                                        padx=18, pady=(10, 2))
        self.log_box = ctk.CTkTextbox(log_card, fg_color="transparent", text_color=INK,
                                      font=ctk.CTkFont(FONT, 11), border_width=0,
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

    def _on_tier_toggle(self) -> None:
        if self.players:
            self._recalculate(keep_log=True)

    def _recalculate(self, keep_log: bool = False) -> None:
        if not self.players:
            return
        self.entries = rank_players(self.players, winners=self.winners,
                                    use_tiers=bool(self.tier_switch.get()))
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
            award_text = entry.award if entry.is_winner else "—"
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
        winners = winners_of(self.entries, self.winners)
        self.stat_values["total"].configure(text=str(len(self.players)))
        self.stat_values["max"].configure(text=f"{max(scores):g}")
        self.stat_values["avg"].configure(text=f"{sum(scores) / len(scores):.1f}")
        self.stat_values["awards"].configure(text=str(len(winners)))
        tiers = {}
        for entry in winners:
            tiers[entry.award] = tiers.get(entry.award, 0) + 1
        parts = " · ".join(f"{k} {v} 人" for k, v in tiers.items())
        self.merge_hint.configure(text=parts if self.tier_switch.get() else "")

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
        winners = winners_of(self.entries, self.winners)
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
            "3. 拖动滑块设置前 N 名，开启分档后自动划分一 / 二 / 三等奖。\n"
            "4. 点「生成奖状」，每人一份 Word，文件名形如 01_张三_一等奖.docx。\n\n"
            "小技巧\n"
            "· 表格前几行有标题、说明没关系，表头会在前 20 行内自动定位。\n"
            "· 没有「总分」列但有多位评委列时，分数自动相加。\n"
            "· 双击排名表里的某一行，可以直接打开已生成的那份奖状。"
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
