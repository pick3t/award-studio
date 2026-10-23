"""多工作表（一个项目一个 sheet）自测。

    python tools/check_sheets.py

用 samples/真实场地 下的模拟分数表验证：
1. 能列出全部工作表，并自动选中第一个有数据的
2. 两行合并表头（「最后」+「得分」）的项目也能读
3. 集体项目用「代表队」作为名称
4. 切换工作表后人数、表头、输出子目录都跟着变
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import app as gui  # noqa: E402
from core.certificate import certificate_filename, inspect_template  # noqa: E402
from core.excel_reader import ExcelParseError, list_sheets, load_players  # noqa: E402
from core.ranking import (  # noqa: E402
    certificate_total,
    expand_entries,
    rank_players,
    split_members,
    winners_of,
)

SAMPLE = ROOT / "samples" / "真实场地" / "一号场地成绩单_模拟分数.xlsx"

failures: list[str] = []


def check(ok: bool, message: str) -> None:
    print(("  [OK] " if ok else "  [!!] ") + message)
    if not ok:
        failures.append(message)


def main() -> int:
    if not SAMPLE.exists():
        print(f"缺少样本：{SAMPLE}")
        print("先运行 python tools/fill_scores.py 生成模拟分数表")
        return 2

    sheets = list_sheets(SAMPLE)
    print(f"工作表（{len(sheets)}）：{'、'.join(sheets)}")
    check(len(sheets) == 7, f"列出 7 个工作表（实际 {len(sheets)}）")

    print("\n1) 直接解析每个工作表")
    counts = {}
    for name in sheets:
        try:
            players, _warns, report = load_players(SAMPLE, sheet_name=name)
            counts[name] = len(players)
            check(bool(players), f"「{name}」读到 {len(players)} 条 · 表头第 {report.header_row} 行")
        except ExcelParseError as exc:
            check(False, f"「{name}」解析失败：{exc}")

    print("\n2) 两行表头与集体项目")
    players, _w, report = load_players(SAMPLE, sheet_name="一场地4组")
    check("最后得分" in report.headers, f"两行表头被拼成「最后得分」：{report.headers}")
    check(len(players) == 11, f"一场地4组 11 人（实际 {len(players)}）")

    teams, _w, _r = load_players(SAMPLE, sheet_name="一场地1组集体")
    check(all(" " not in p.name for p in teams), "代表队名称里没有换行/空格")
    check(any("太极辅导站" in p.name for p in teams), "集体项目用代表队作为名称")
    check("参赛选手" in teams[0].extras, "参赛选手保留为附加字段，可用于奖状占位符")

    print("\n3) 集体项目按人发奖状")
    entries = rank_players(teams, winners=2)
    winners = winners_of(entries, 2)
    people = expand_entries(winners)
    members_total = sum(len(split_members(e.player.extras["参赛选手"])) for e in winners)
    check(certificate_total(winners) == members_total, f"奖状份数 = 队员总数（{members_total}）")
    check(len(people) == members_total, f"展开成 {len(people)} 条个人记录")
    check(all(e.player.name != e.player.extras.get("代表队") for e in people), "姓名换成了个人姓名")
    check(all(e.player.extras.get("代表队") for e in people), "每人仍保留所属代表队")
    check(all(" " not in e.player.name for e in people), "个人姓名没有残留空格")
    first = people[0]
    name = certificate_filename(first)
    check(first.player.extras["代表队"] in name, f"文件名带上代表队：{name}")

    print("\n4) 带项目组的模板")
    template = ROOT / "samples" / "奖状模板_项目组.docx"
    if template.exists():
        keys = inspect_template(template)
        check({"项目组", "参赛项目"} <= set(keys), f"模板包含项目组/参赛项目占位符：{keys}")
        check({"姓名", "奖项", "名次"} <= set(keys), "模板包含姓名/奖项/名次")
    else:
        check(False, "缺少 samples/奖状模板_项目组.docx（运行 tools/make_template.py）")

    print("\n5) 界面里切换工作表")
    window = gui.AwardStudio()
    window.withdraw()
    try:
        window._load_excel(SAMPLE)
        window.update_idletasks()
        check(window.current_sheet == sheets[0], f"默认选中第一个表：{window.current_sheet}")
        check(len(window.players) == counts[sheets[0]], "人数与直接解析一致")
        check(window.sheet_menu.cget("values") == sheets, "下拉里列出全部工作表")

        window._on_sheet_change("一场地2组")
        window.update_idletasks()
        check(window.current_sheet == "一场地2组", "切换到一场地2组")
        check(len(window.players) == counts["一场地2组"], f"人数跟着变成 {counts['一场地2组']}")
        expected = SAMPLE.parent / "output" / "奖状" / "一场地2组"
        check(window._target_dir() == expected, f"输出子目录：{window._target_dir().name}")

        window._on_sheet_change("一场地1组集体")
        window.update_idletasks()
        check(window._target_dir().name == "一场地1组集体", "集体项目也有独立输出目录")
        check(window.sheet_names == sheets, "切换后工作表列表保持不变")

        window._on_heading_click("分数")
        window.update_idletasks()
        winners = winners_of(window.entries)
        expected_count = certificate_total(winners)
        shown = window.stat_values["awards"].cget("text")
        check(shown == str(expected_count),
              f"界面上「奖状份数」按人算：{shown}（队数 {len(winners)}）")
    finally:
        window.destroy()

    print()
    if failures:
        print(f"失败 {len(failures)} 项：")
        for item in failures:
            print(" - " + item)
        return 1
    print("全部通过 ✓ 一个工作簿里的多个项目可以分别读、分别排名、分别输出")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
