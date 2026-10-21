"""轻量公式兜底求值。

openpyxl 只能读到 Excel 保存过的缓存值。如果用户的成绩表是别的程序生成的、
公式还没被 Excel 计算过，分数列会读成空。这里对常见的 SUM / AVERAGE / 四则运算
做一次兜底计算，覆盖「总分=SUM(评委列)」这类真实场景。
"""

from __future__ import annotations

import ast
import re

from openpyxl.utils import column_index_from_string

RANGE_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d+)\s*:\s*\$?([A-Za-z]{1,3})\$?(\d+)")
CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
FUNC_RE = re.compile(r"\b([A-Za-z]+)\s*\(([^()]*)\)")

ALLOWED_NODES = (
    ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Add, ast.Sub,
    ast.Mult, ast.Div, ast.Pow, ast.Mod, ast.FloorDiv, ast.USub, ast.UAdd,
)


def _safe_eval(expr: str) -> float | None:
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError:
        return None
    for node in ast.walk(tree):
        if not isinstance(node, ALLOWED_NODES):
            return None
    try:
        result = eval(compile(tree, "<formula>", "eval"), {"__builtins__": {}}, {})  # noqa: S307
    except Exception:  # noqa: BLE001 - 除零、类型错误等一律视为无法求值
        return None
    return float(result) if isinstance(result, (int, float)) else None


class FormulaBook:
    """在一个工作簿内解析公式单元格。"""

    def __init__(self, values_sheet, formulas_sheet):
        self.values_sheet = values_sheet
        self.formulas_sheet = formulas_sheet
        self.cache: dict[tuple[int, int], float | str | None] = {}

    def value(self, row: int, col: int, depth: int = 0):
        key = (row, col)
        if key in self.cache:
            return self.cache[key]
        if depth > 8:
            return None

        cached = self.values_sheet.cell(row=row, column=col).value
        if cached is not None and not (isinstance(cached, str) and cached.startswith("=")):
            self.cache[key] = cached
            return cached

        raw = self.formulas_sheet.cell(row=row, column=col).value
        if isinstance(raw, str) and raw.startswith("="):
            result = self._eval(raw[1:], depth)
            self.cache[key] = result
            return result
        self.cache[key] = cached
        return cached

    # ------------------------------------------------------------ 内部
    def _numbers(self, args: str, depth: int) -> list[float]:
        numbers: list[float] = []
        for part in args.split(","):
            part = part.strip()
            if not part:
                continue
            match = RANGE_RE.fullmatch(part.replace(" ", ""))
            if match:
                col1, row1 = column_index_from_string(match.group(1).upper()), int(match.group(2))
                col2, row2 = column_index_from_string(match.group(3).upper()), int(match.group(4))
                for r in range(min(row1, row2), max(row1, row2) + 1):
                    for c in range(min(col1, col2), max(col1, col2) + 1):
                        v = self.value(r, c, depth + 1)
                        if isinstance(v, (int, float)) and not isinstance(v, bool):
                            numbers.append(float(v))
                continue
            match = CELL_RE.fullmatch(part.replace(" ", ""))
            if match:
                v = self.value(int(match.group(2)), column_index_from_string(match.group(1).upper()), depth + 1)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    numbers.append(float(v))
                continue
            try:
                numbers.append(float(part))
            except ValueError:
                continue
        return numbers

    def _eval(self, expression: str, depth: int) -> float | None:
        if "!" in expression or "[" in expression:  # 跨表 / 外部引用不处理
            return None

        def replace(match: re.Match) -> str:
            name = match.group(1).upper()
            numbers = self._numbers(match.group(2), depth)
            if name in {"SUM", "PRODUCT"}:
                if not numbers:
                    return "0"
                result = 0.0 if name == "SUM" else 1.0
                for n in numbers:
                    result = result + n if name == "SUM" else result * n
                return repr(result)
            if not numbers:
                return "0"
            if name in {"AVERAGE", "AVG"}:
                return repr(sum(numbers) / len(numbers))
            if name == "MIN":
                return repr(min(numbers))
            if name == "MAX":
                return repr(max(numbers))
            if name == "COUNT":
                return repr(float(len(numbers)))
            return match.group(0)

        expanded = FUNC_RE.sub(replace, expression)
        if re.search(r"[A-Za-z]", expanded):  # 还有没展开的函数名，放弃
            return None
        return _safe_eval(expanded)
