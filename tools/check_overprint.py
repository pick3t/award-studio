"""自测：套印模板（文字在文本框里）也能识别占位符并正确替换。

    python tools/check_overprint.py

验证点：
1. 文本框里的 {{占位符}} 能被识别
2. 生成后文本框里的占位符全部替换成真实值（不是空白）
3. 替换只改文字，run 的格式（字体/字号/加粗）保持原样
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.certificate import (  # noqa: E402
    CertificateContext,
    generate_certificate,
    inspect_template,
)
from core.excel_reader import load_players  # noqa: E402
from core.ranking import expand_entries, rank_players, winners_of  # noqa: E402

SAMPLES = ROOT / "samples"
OVERPRINT = SAMPLES / "套印"
SHEET = Path.home() / "Desktop" / "一号场地成绩单_模拟分数.xlsx"
OUT = Path(r"C:\Users\admin\AppData\Local\Temp\_overprint_check")

failures: list[str] = []


def check(ok: bool, message: str) -> None:
    print(("  [OK] " if ok else "  [!!] ") + message)
    if not ok:
        failures.append(message)


def box_texts(path: Path) -> list[str]:
    document = Document(str(path))
    return ["".join(t.text or "" for t in box.iter(qn("w:t")))
            for box in document.element.body.iter(qn("w:txbxContent"))]


def first_run_props(path: Path) -> str | None:
    document = Document(str(path))
    box = next(document.element.body.iter(qn("w:txbxContent")))
    run = next(box.iter(qn("w:r")))
    props = run.find(qn("w:rPr"))
    return None if props is None else props.xml


def main() -> int:
    person = OVERPRINT / "奖状套印模板（个人）_占位符.docx"
    team = OVERPRINT / "奖状套印模板（集体）_占位符.docx"
    if not person.exists() or not team.exists():
        print(f"缺少套印模板：{OVERPRINT}")
        return 2
    if not SHEET.exists():
        print(f"缺少成绩表：{SHEET}（先运行 tools/fill_scores.py）")
        return 2

    print("1) 识别文本框里的占位符")
    keys = inspect_template(person)
    check(set(keys) == {"分数", "组别", "姓名", "代表队", "参赛项目", "名次中文"},
          f"个人模板识别到 {keys}")
    keys = inspect_template(team)
    check(set(keys) == {"代表队", "参赛项目", "奖项", "分数", "姓名"},
          f"集体模板识别到 {keys}")

    print("2) 生成并检查替换结果")
    OUT.mkdir(parents=True, exist_ok=True)
    context = CertificateContext(contest="自测比赛", issuer="组委会", date="2026 年 10 月 26 日")

    players, _, _ = load_players(SHEET, sheet_name="一场地4组")
    entry = winners_of(rank_players(players, winners=1))[0]
    out = OUT / "个人.docx"
    generate_certificate(person, out, entry, context)
    texts = [t for t in box_texts(out) if t]
    check(all("{{" not in t and "}}" not in t for t in texts), "个人模板：没有残留占位符")
    check(any(t == entry.player.name for t in texts), f"姓名已填入：{entry.player.name}")
    check(any(t == f"{entry.player.score:g}" for t in texts), f"分数已填入：{entry.player.score:g}")
    check(any("武术辅导站" in t or "太极" in t or "队" in t for t in texts), "代表队/单位已填入")
    check(first_run_props(person) == first_run_props(out), "个人模板：字体字号等格式未被改动")

    teams, _, _ = load_players(SHEET, sheet_name="一场地1组集体")
    team_entry = expand_entries(winners_of(rank_players(teams, winners=1)), per_person=False)[0]
    out = OUT / "集体.docx"
    generate_certificate(team, out, team_entry, context)
    texts = [t for t in box_texts(out) if t]
    check(all("{{" not in t and "}}" not in t for t in texts), "集体模板：没有残留占位符")
    check(any(t == team_entry.player.extras.get("代表队") for t in texts), "代表队已填入")
    roster = team_entry.player.extras.get("参赛选手", "").split()
    lines = []
    for text in texts:
        names = text.split()
        if names and all(name in roster for name in names) and text not in lines:
            lines.append(text)
    check(len(lines) >= 2, f"姓名按两行分开放：{lines[:2]}")
    if len(lines) >= 2:
        first, second = lines[0].split(), lines[1].split()
        check(len(first) == len(roster) // 2 and len(second) == len(roster) - len(first),
              f"两行人数 {len(first)}+{len(second)} = {len(roster)}")
        check(set(first) | set(second) == set(roster), "两行合起来正好是全体队员")
        check(not (set(first) & set(second)), "两行之间没有重复的人")
    check(first_run_props(team) == first_run_props(out), "集体模板：格式未被改动")

    # 勾「每人一份」时只填一个人的名字，第二行留空（从原始队记录展开）
    person = expand_entries(winners_of(rank_players(teams, winners=1)), per_person=True)[0]
    out_person = OUT / "集体_每人.docx"
    generate_certificate(team, out_person, person, context)
    texts = [t for t in box_texts(out_person) if t]
    check(person.player.name in texts, f"每人一份时填个人姓名：{person.player.name}")
    others = [t for t in texts if " " in t and len(t) > 6]
    check(not others, f"每人一份时不再拼接全队：{others[:1]}")

    print("3) 合并版：文本框形状 id 必须唯一")
    merged = OUT / "合并版.docx"
    from core.certificate import merge_documents
    singles = sorted(OUT.glob("*.docx"))
    singles = [p for p in singles if p.name in ("个人.docx", "集体.docx")]
    if len(singles) >= 2:
        merge_documents(singles, merged)
        document = Document(str(merged))
        ids = [element.get("id") for element in
               document.element.body.iter(
                   "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr")]
        check(len(ids) == len(set(ids)), f"合并后 {len(ids)} 个形状 id 全部唯一")
        # 文本框数量不写死：随模板结构变化（是否带 VML 兼容副本）自动跟随，
        # 只保证「合并后不丢形状」这一条真正的不变量。
        counts = {path.name: len(box_texts(path)) for path in singles}
        expected_boxes = sum(counts.values())
        actual_boxes = len(box_texts(merged))
        detail = " + ".join(f"{name.split('.')[0]} {count}" for name, count in counts.items())
        check(actual_boxes == expected_boxes,
              f"合并版文本框数量 {actual_boxes} = {detail}")

    print()
    if failures:
        print(f"失败 {len(failures)} 项：")
        for item in failures:
            print(" - " + item)
        return 1
    print("全部通过 ✓ 文本框里的占位符可识别、可替换，格式不变")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
