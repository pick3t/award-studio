"""界面自测：字段映射能显示、能改，生成的奖状内容正确。

    python tools/check_ui.py

验证点：
1. 载入带文本框的套印模板后，映射区列出全部占位符（行数 = 占位符数，且真的显示出来）
2. 每行「自动」旁边都是可选的下拉，不是死文字
3. 把「名次中文」改成「奖项」后生成，四份奖状依次是 金/银/银/铜
4. 合并版四页的内容各不相同（不是每页都显示最后一页）
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import app as gui  # noqa: E402
import customtkinter as ctk  # noqa: E402
from core.certificate import CertificateContext  # noqa: E402

SHEET = Path.home() / "Desktop" / "一号场地成绩单_模拟分数.xlsx"
TEMPLATE = ROOT / "samples" / "套印" / "奖状套印模板（个人）_占位符.docx"
OUT = Path(r"C:\Users\admin\AppData\Local\Temp\_ui_check")

failures: list[str] = []


def check(ok: bool, message: str) -> None:
    print(("  [OK] " if ok else "  [!!] ") + message)
    if not ok:
        failures.append(message)


def box_texts(path: Path) -> list[str]:
    document = Document(str(path))
    return ["".join(t.text or "" for t in box.iter(qn("w:t")))
            for box in document.element.body.iter(qn("w:txbxContent"))]


def main() -> int:
    if not SHEET.exists() or not TEMPLATE.exists():
        print("缺少样本（成绩表或套印模板）")
        return 2

    window = gui.AwardStudio()
    window.withdraw()
    try:
        window._load_excel(SHEET)
        window._apply_sheet("一场地1组", keep_columns=True)
        window._on_heading_click("分数")
        window._load_template(TEMPLATE)
        window.update_idletasks()

        # 1) 映射区显示
        rows = window.mapping_box.winfo_children()
        check(len(window.field_rules) == 6, f"模板识别出 6 个占位符（{len(window.field_rules)}）")
        check(len(rows) == len(window.field_rules),
              f"映射行数 = 占位符数（{len(rows)} 行）")
        check(bool(window.mapping_box.winfo_ismapped()), "映射列表真的显示出来了")

        # 2) 每行都有下拉
        def walk(widget):
            for child in widget.winfo_children():
                yield child
                yield from walk(child)

        menus = [w for row in rows for w in walk(row) if isinstance(w, ctk.CTkOptionMenu)]
        check(len(menus) >= len(rows),
              f"每行都有来源下拉（共 {len(menus)} 个下拉）")
        labels = list(gui.field_rules.SOURCE_LABELS.values())
        check(labels == ["表格列", "固定值", "系统日期", "指定日期", "奖项", "名次"],
              f"来源只有 6 种：{labels}")
        check(all(r.source in gui.field_rules.SOURCE_LABELS for r in window.field_rules),
              "所有规则都用了新的 6 种来源之一")

        # 3) 名次中文 → 奖项，生成四份
        window.award_switch.select()
        window.awards = [{"name": "金", "count": 1}, {"name": "银", "count": 2},
                         {"name": "铜", "count": 1}]
        window._recalculate(keep_log=True)
        winners = [e for e in window.entries if e.is_winner]
        check(len(winners) == 4, f"4 人获奖（{len(winners)}）")

        rule = next(r for r in window.field_rules if r.key == "名次中文")
        window._on_source_change(rule, "名次")
        context = CertificateContext(contest="自测", issuer="组委会", date="2026 年 10 月 26 日")
        OUT.mkdir(parents=True, exist_ok=True)
        for old in OUT.glob("*.docx"):
            old.unlink()
        window._worker(TEMPLATE, OUT, winners, context, True, list(window.field_rules), False)

        singles = sorted(p for p in OUT.glob("0*.docx"))
        check(len(singles) == 4, f"生成 4 份单份奖状（{len(singles)}）")
        awards = []
        for path in singles:
            texts = box_texts(path)
            awards.append(texts[-2] if len(texts) >= 2 else "")
        check(awards == ["第一名", "第二名", "第三名", "第四名"],
              f"名次来源给出第X名：{awards}")

        # 改成「奖项」来源
        window._on_source_change(rule, "奖项")
        for old in OUT.glob("*.docx"):
            old.unlink()
        window._worker(TEMPLATE, OUT, winners, context, True, list(window.field_rules), False)
        singles = sorted(p for p in OUT.glob("0*.docx"))
        awards = [box_texts(p)[-2] for p in singles]
        check(awards == ["金", "银", "银", "铜"], f"奖项来源给出奖项：{awards}")

        # 4) 合并版分页内容
        merged = OUT / "全部奖状_合并版.docx"
        check(merged.exists(), "生成了合并版")
        if merged.exists():
            texts = box_texts(merged)
            pages = [[texts[i * 12 + j] for j in range(0, 12, 2)]
                     for i in range(len(texts) // 12)]
            page_awards = [page[-1] for page in pages]
            check(page_awards == ["金", "银", "银", "铜"],
                  f"合并版四页奖项各不相同：{page_awards}")
        # 5) 切换工作表后，列、表头、映射都要跟着变
        print("4) 切换工作表后不残留")
        window._on_sheet_change("一场地1组集体")
        window.update_idletasks()
        columns = [c[1] for c in window._table_columns]
        check("代表队" in columns, f"集体表出现「代表队」列：{columns}")
        check(columns.count("代表队") == 1, "「代表队」列没有重复")
        check(columns[2] == "代表队", f"名称列按原表头显示（{columns[2]}）")
        options = window._column_names()
        check("代表队" in options, f"表格列下拉里能看到代表队：{options}")
        rule = next(r for r in window.field_rules if r.key == "代表队")
        check(rule.source == "column" and rule.value == "代表队",
              f"代表队的映射自动指向代表队列：{rule.source}/{rule.value}")
        rule = next(r for r in window.field_rules if r.key == "组别")
        check(rule.source == "text",
              f"新表没有组别列时改成固定值：{rule.source}/{rule.value!r}")
        row = window.tree.item(window.tree.get_children()[0], "values")
        check(row[2] == window.entries[0].player.name,
              f"首行是当前表排名第一的人/队：{row[2]}")

        window._on_sheet_change("一场地1组")
        window.update_idletasks()
        columns = [c[1] for c in window._table_columns]
        check(columns[2] == "姓名", "切回个人表恢复「姓名」表头")
        check("序号" in columns and "编号" not in columns,
              f"编号列按原表头显示成「序号」：{columns}")
    finally:
        window.destroy()

    print()
    if failures:
        print(f"失败 {len(failures)} 项：")
        for item in failures:
            print(" - " + item)
        return 1
    print("全部通过 ✓ 六种来源可选，奖项/名次/合并版都对")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
