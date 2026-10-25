"""界面逻辑自测：没有任何评判列时，表格必须完全不排序。

    python tools/check_ranking.py

验证点：
1. 载入成绩表后，评判列为空（程序里不能有任何默认评判列）
2. 表格列完全来自成绩表，不凭空出现「总分 / 合计 / 平均」
3. 名次列全部显示「—」，行顺序 = Excel 原始行顺序
4. 统计区最高分/平均分/获奖人数显示「—」，生成按钮不可用
5. 勾选一列后才排序，取消勾选后回到原始顺序（可逆）
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import app as gui  # noqa: E402

SAMPLES = ROOT / "samples"
CASES = [
    ("成绩表_多评委无总分.xlsx", "评委一"),
    ("成绩表_同分并列.xlsx", None),
    ("决赛评分表.xlsx", "分数"),
    ("成绩表_两列最简.xlsx", None),
]

failures: list[str] = []


def check(ok: bool, message: str) -> None:
    print(("  [OK] " if ok else "  [!!] ") + message)
    if not ok:
        failures.append(message)


def cells(tree) -> list[list[str]]:
    return [list(tree.item(item, "values")) for item in tree.get_children()]


def run(window, filename: str, numeric_key: str | None) -> None:
    path = SAMPLES / filename
    print(f"\n== {filename}")
    if not path.exists():
        check(False, f"缺少样本 {filename}")
        return

    window._load_excel(path)
    window.update_idletasks()
    order = [p.name for p in window.players]

    cols = [spec[1] for spec in window._table_columns]
    labels = set(window.column_names) | set(window.score_column_names)
    stray = [c for c in cols if c not in labels and c not in ("选", "名次", "奖项")]
    check(window.score_columns == [], f"载入后没有任何默认评判列（{window.score_columns}）")
    check(not stray, f"列全部来自成绩表，没有计算列：{cols}")
    rows = cells(window.tree)
    rank_idx = cols.index("名次")
    check(all(row[rank_idx] == "—" for row in rows), "名次列全部为「—」")
    check([row[cols.index("姓名")] for row in rows] == order, "行顺序 = Excel 原始行顺序")
    stats = [label.cget("text") for label in window.stat_values.values()]
    check(stats[1] == "—" and stats[2] == "—" and stats[3] == "—", f"统计区不显示分数/奖项：{stats}")
    window._sync_buttons()
    check(str(window.generate_btn.cget("state")) == "disabled", "未选评判列时生成按钮不可用")

    if numeric_key is None:
        numeric_key = next((c for c in cols if c.startswith("评委")) or "分数", None)
    if numeric_key and numeric_key in cols:
        window._on_heading_click(numeric_key)
        window.update_idletasks()
        ranked = cells(window.tree)
        ranks = [row[cols.index("名次")] for row in ranked]
        check(window.score_columns == [numeric_key], f"勾选「{numeric_key}」后作为唯一评判列")
        check(ranks[0] == "1", f"勾选后开始排名（首行名次 {ranks[0]}）")
        check(any(r != "—" for r in ranks), "勾选后名次列有实际名次")
        window._sync_buttons()
        check(str(window.generate_btn.cget("state")) == "normal", "勾选后生成按钮可用")

        window._on_heading_click(numeric_key)
        window.update_idletasks()
        back = cells(window.tree)
        check(window.score_columns == [], "再次点击可取消勾选")
        check(all(row[cols.index("名次")] == "—" for row in back),
              "取消勾选后名次恢复为「—」")
        check([row[cols.index("姓名")] for row in back] == order, "取消勾选后回到 Excel 原始顺序")


def main() -> int:
    window = gui.AwardStudio()
    window.withdraw()
    # 直接给模板路径，让「生成按钮是否可用」只取决于有没有勾选评判列（不走模板解析，避免改配置）
    window.template_path = ROOT / "samples" / "奖状模板.docx"
    try:
        for filename, key in CASES:
            run(window, filename, key)
    finally:
        window.destroy()

    print()
    if failures:
        print(f"失败 {len(failures)} 项：")
        for item in failures:
            print(" - " + item)
        return 1
    print("全部通过 ✓ 未选评判列时完全不排序")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
