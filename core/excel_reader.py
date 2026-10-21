"""成绩表解析：自动识别表头、姓名列、分数列。

支持：
* 表头不在第一行（前面有标题、说明行）
* 姓名/分数列名多种写法（姓名、选手、参赛者、总分、成绩、得分……）
* 多个评委打分列（无"总分"列时自动求和）
* 分数带单位（"95.5 分"）、百分制、空单元格
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .formula import FormulaBook


class ExcelParseError(Exception):
    """成绩表无法解析时抛出。"""


# ---------------------------------------------------------------- 列名别名
NAME_ALIASES = (
    "姓名", "名字", "选手", "选手姓名", "参赛者", "参赛选手", "参赛人员",
    "学生姓名", "队员", "队员姓名", "人员", "人员姓名", "姓名name", "name",
)
SCORE_ALIASES = (
    "总分", "总得分", "最终得分", "综合得分", "合计得分", "分数", "得分",
    "成绩", "评分", "总成绩", "总分值", "score", "totalscore", "marks",
    # 只有评委列、没有总分的表：评委一/评委二/评委A 等也当作分数列
    "评委", "打分", "得分率", "petunjuk",
)
NO_ALIASES = (
    "编号", "选手编号", "参赛编号", "序号", "学号", "参赛号", "选手号",
    "号码", "no", "id", "number",
)
EXTRA_ALIASES = (
    "班级", "单位", "部门", "学校", "年级", "专业", "队名", "团队",
    "指导老师", "指导教师", "备注", "说明",
)

ROLE_LABELS = {
    "name": "姓名",
    "score": "分数",
    "no": "编号",
    "extra": "附加",
}


@dataclass
class Player:
    """一名选手。"""

    name: str
    score: float
    no: str = ""
    extras: dict[str, str] = field(default_factory=dict)
    row: int = 0

    @property
    def fields(self) -> dict[str, str]:
        """姓名/编号/分数 + 表格里的其他列，供字段映射按列名取值。"""
        data = {"姓名": self.name, "编号": self.no, "分数": f"{self.score:g}"}
        data.update(self.extras)
        return data


@dataclass
class ParseReport:
    """解析结果摘要，用于界面展示与排查。"""

    path: str
    sheet: str
    header_row: int
    headers: list[str]
    columns: dict[str, list[int]] = field(default_factory=dict)
    total_rows: int = 0
    title: str = ""

    def describe(self) -> str:
        parts = []
        for role in ("name", "no", "score"):
            idx = self.columns.get(role)
            if idx:
                names = "、".join(self.headers[i] for i in idx if i < len(self.headers))
                parts.append(f"{ROLE_LABELS[role]}：{names}")
        return " | ".join(parts)


# ---------------------------------------------------------------- 工具函数
def _norm(value: Any) -> str:
    """归一化文本：全角转半角、去空格与常见标点、转小写。"""
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[\s_/\\()\[\]【】{}<>《》:：,，.。、'\"-]", "", text)
    return text.strip().lower()


def _role_score(header: Any, aliases: tuple[str, ...]) -> int:
    """表头与某角色别名的匹配得分：完全相等 > 前后缀 > 包含。"""
    text = _norm(header)
    if not text:
        return 0
    best = 0
    for alias in aliases:
        alias_norm = _norm(alias)
        if not alias_norm:
            continue
        if text == alias_norm:
            best = max(best, 100)
        elif text.startswith(alias_norm) or text.endswith(alias_norm):
            best = max(best, 80)
        elif alias_norm in text:
            best = max(best, 60)
    return best


def parse_score(value: Any) -> float | None:
    """把单元格里的分数解析成 float，失败返回 None。"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = unicodedata.normalize("NFKC", str(value)).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("，", "").replace(" ", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return float(match.group())
    except ValueError:
        return None


def _detect_header(rows: list[tuple[int, list[Any]]]) -> int | None:
    """在前若干行里找最像表头的一行，返回其行号（1 基）。"""
    best_row, best_total = None, 0
    for row_no, values in rows:
        scores = {"name": 0, "score": 0, "no": 0}
        for value in values:
            for role, aliases in (
                ("name", NAME_ALIASES),
                ("score", SCORE_ALIASES),
                ("no", NO_ALIASES),
            ):
                scores[role] = max(scores[role], _role_score(value, aliases))
        # 姓名列和分数列必须同时出现，否则不认为是表头
        if scores["name"] < 60 or scores["score"] < 60:
            continue
        total = scores["name"] + scores["score"] + scores["no"] * 0.5
        if total > best_total:
            best_row, best_total = row_no, total
    return best_row


def _detect_title(rows: list[tuple[int, list[Any]]], header_row: int) -> str:
    """表头上方第一行非空文字，当成比赛名称的候选。"""
    for row_no, values in rows:
        if row_no >= header_row:
            break
        for value in values:
            text = _cell_text(value)
            if text and len(text) > 2:
                return text
    return ""


def _map_columns(header_values: list[Any]) -> dict[str, list[int]]:
    """把每一列归类到 name / score / no / extra。"""
    columns: dict[str, list[int]] = {"name": [], "score": [], "no": [], "extra": []}
    best_scores: dict[str, list[int]] = {"name": [], "score": [], "no": [], "extra": []}
    for idx, value in enumerate(header_values):
        candidate = None
        for role, aliases in (
            ("name", NAME_ALIASES),
            ("score", SCORE_ALIASES),
            ("no", NO_ALIASES),
            ("extra", EXTRA_ALIASES),
        ):
            score = _role_score(value, aliases)
            if score and (candidate is None or score > candidate[1]):
                candidate = (role, score)
        if candidate:
            role, score = candidate
            columns[role].append(idx)
            best_scores[role].append(score)

    # 每个角色只保留得分最高的列，score 列允许多个（多评委）
    for role in ("name", "no"):
        if columns[role]:
            top = max(best_scores[role])
            keep = [c for c, s in zip(columns[role], best_scores[role]) if s == top]
            columns[role] = [keep[0]]

    if columns["score"]:
        norms = [_norm(header_values[i]) for i in columns["score"]]
        total_like = [
            i for i, n in zip(columns["score"], norms)
            if any(k in n for k in ("总", "最终", "综合", "合计"))
        ]
        columns["score"] = total_like or columns["score"]
    return columns


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


class _CellReader:
    """读单元格：优先缓存值，为空时尝试用公式兜底计算。"""

    def __init__(self, path: Path, sheet) -> None:
        self.path = path
        self.sheet = sheet
        self._book: FormulaBook | None = None
        self._book_wb = None

    def get(self, row: int, col: int) -> Any:
        value = self.sheet.cell(row=row, column=col).value
        if value is not None and not (isinstance(value, str) and value.startswith("=")):
            return value
        if self._book is None:
            try:
                self._book_wb = load_workbook(self.path, data_only=False)
                self._book = FormulaBook(self.sheet, self._book_wb[self.sheet.title])
            except Exception:  # noqa: BLE001 - 读不到就当作空值
                self._book = FormulaBook(self.sheet, self.sheet)
        return self._book.value(row, col)

    def close(self) -> None:
        if self._book_wb is not None:
            try:
                self._book_wb.close()
            except Exception:  # noqa: BLE001
                pass


# ---------------------------------------------------------------- 主入口
def load_players(path: str | Path, sheet_name: str | None = None) -> tuple[list[Player], list[str], ParseReport]:
    """读取成绩表，返回 (选手列表, 警告列表, 解析报告)。"""
    file_path = Path(path)
    if not file_path.exists():
        raise ExcelParseError(f"文件不存在：{file_path}")
    if file_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ExcelParseError("只支持 .xlsx / .xlsm 格式，请先用 Excel 另存为 xlsx")

    try:
        workbook = load_workbook(file_path, data_only=True)
    except Exception as exc:  # noqa: BLE001
        raise ExcelParseError(f"无法打开表格：{exc}") from exc

    try:
        worksheet = workbook[sheet_name] if sheet_name else workbook.worksheets[0]
    except KeyError as exc:
        raise ExcelParseError(f"找不到工作表：{sheet_name}") from exc

    reader = _CellReader(file_path, worksheet)
    max_col = min(worksheet.max_column, 60)
    rows: list[tuple[int, list[Any]]] = []
    raw_rows: dict[int, list[Any]] = {}
    max_scan = min(worksheet.max_row, 20)
    for row_no in range(1, max_scan + 1):
        values = [reader.get(row_no, c) for c in range(1, max_col + 1)]
        raw_rows[row_no] = values
        if any(v not in (None, "") for v in values):
            rows.append((row_no, values))

    header_row = _detect_header(rows)
    if header_row is None:
        workbook.close()
        raise ExcelParseError(
            "没找到表头：需要同时包含「姓名」（或 选手/参赛者）和「分数」（或 总分/成绩/得分）两列"
        )

    headers = [_cell_text(v) for v in raw_rows[header_row]]
    columns = _map_columns(raw_rows[header_row])
    if not columns["name"] or not columns["score"]:
        workbook.close()
        raise ExcelParseError("识别到表头，但姓名列或分数列定位失败")

    name_col = columns["name"][0]
    no_col = columns["no"][0] if columns["no"] else None
    score_cols = columns["score"]
    # 其余有表头的列都作为附加字段（班级、作品名称、指导老师……），可用于奖状占位符
    used_cols = {name_col, *score_cols}
    if no_col is not None:
        used_cols.add(no_col)
    extra_cols = [
        i for i in range(len(raw_rows[header_row]))
        if i not in used_cols and _cell_text(raw_rows[header_row][i])
    ]

    players: list[Player] = []
    warnings: list[str] = []
    header_texts = {_norm(headers[i]) for i in range(len(headers))}

    for row_no in range(header_row + 1, worksheet.max_row + 1):
        values = [reader.get(row_no, c) for c in range(1, max_col + 1)]
        if all(v in (None, "") for v in values):
            continue
        name = _cell_text(values[name_col]) if name_col < len(values) else ""
        # 跳过重复表头、纯说明行
        if not name or _norm(name) in header_texts:
            continue

        scores = [parse_score(values[i]) for i in score_cols if i < len(values)]
        valid = [s for s in scores if s is not None]
        if not valid:
            warnings.append(f"第 {row_no} 行「{name}」分数无效，已跳过")
            continue
        total = float(sum(valid))
        if len(valid) > 1:
            total = round(total, 4)

        extras = {}
        for i in extra_cols:
            if i < len(values):
                text = _cell_text(values[i])
                if text:
                    extras[headers[i] or f"列{i + 1}"] = text

        players.append(
            Player(
                name=name,
                score=total,
                no=_cell_text(values[no_col]) if no_col is not None and no_col < len(values) else "",
                extras=extras,
                row=row_no,
            )
        )

    report = ParseReport(
        path=str(file_path),
        sheet=worksheet.title,
        header_row=header_row,
        headers=headers,
        columns=columns,
        total_rows=len(players),
        title=_detect_title(rows, header_row),
    )
    workbook.close()
    reader.close()

    if not players:
        raise ExcelParseError("表格里没有读到有效的选手数据")
    return players, warnings, report


def list_sheets(path: str | Path) -> list[str]:
    """列出工作簿里的所有工作表名。"""
    workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()
