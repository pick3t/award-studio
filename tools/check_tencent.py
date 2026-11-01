"""自测：腾讯文档链接 → 拉表 → 落 xlsx → 走原有解析/排名链路。

    python tools/check_tencent.py                  # 离线（假数据，不联网）
    python tools/check_tencent.py --live <链接>     # 实网拉一份真实表格

离线部分用假 Transport 模拟接口，重点验证：
1. 链接解析（带 ?tab= / #hash / 纯 file_id / 非表格类型 / 垃圾输入）
2. 分页读取（超过单页行数时能读全，遇整页空白提前停）
3. 尾部空行裁剪
4. 落盘 xlsx 的表头/类型能被 core.excel_reader 正确识别
5. 空表格要报错，而不是生成一份空奖状
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.excel_reader import load_players  # noqa: E402
from core.ranking import rank_players, winners_of  # noqa: E402
from core.tencent_docs import (  # noqa: E402
    BadLink,
    DocNotFound,
    NotSpreadsheet,
    SheetMeta,
    TencentSheet,
    Transport,
    parse_link,
    sync_to_cache,
    write_xlsx,
)

failures: list[str] = []


def check(ok: bool, message: str) -> None:
    print(("  [OK] " if ok else "  [!!] ") + message)
    if not ok:
        failures.append(message)


class FakeTransport(Transport):
    """内存里的假表格：cells 里只存有值的格子，和真实接口一样是稀疏的。"""

    def __init__(self, cells: dict[tuple[int, int], object],
                 rows: int = 200, cols: int = 26,
                 sheets: list[SheetMeta] | None = None) -> None:
        self.cells = cells
        self.rows = rows
        self.cols = cols
        self._sheets = sheets or [SheetMeta("S1", "工作表1", rows, cols)]
        self.reads: list[tuple[int, int, int, int]] = []

    def sheets(self, file_id: str) -> list[SheetMeta]:
        return self._sheets

    def read_range(self, file_id, sheet_id, start_row, start_col, end_row, end_col):
        self.reads.append((start_row, start_col, end_row, end_col))
        return {
            (r, c): v for (r, c), v in self.cells.items()
            if start_row <= r < end_row and start_col <= c < end_col
        }


def test_links() -> None:
    print("1) 链接解析")
    cases = [
        ("https://docs.qq.com/sheet/DS29TWUFoYnNmbUxH", "sheet", "KoSYAhbsfmLG"),
        ("https://docs.qq.com/sheet/DS29TWUFoYnNmbUxH?tab=BB08J2", "sheet", "KoSYAhbsfmLG"),
        ("https://docs.qq.com/sheet/DS29TWUFoYnNmbUxH#slide=1", "sheet", "KoSYAhbsfmLG"),
        ("  docs.qq.com/sheet/DS29TWUFoYnNmbUxH  ", "sheet", "KoSYAhbsfmLG"),
        ("KoSYAhbsfmLG", "sheet", "KoSYAhbsfmLG"),
    ]
    for raw, kind, file_id in cases:
        info = parse_link(raw)
        check(info.kind == kind and info.file_id == file_id,
              f"{raw} → {info.kind}/{info.file_id}")

    for raw, error in (("https://docs.qq.com/smartsheet/AAAA", NotSpreadsheet),
                       ("https://docs.qq.com/doc/DUXh", NotSpreadsheet),
                       ("https://docs.qq.com/sheet/", BadLink),
                       ("随便一段文字", BadLink),
                       ("", BadLink)):
        try:
            parse_link(raw)
        except error:
            check(True, f"「{raw or '空字符串'}」被正确拒绝（{error.__name__}）")
        except Exception as exc:  # noqa: BLE001
            check(False, f"「{raw}」应报 {error.__name__}，实际 {type(exc).__name__}")
        else:
            check(False, f"「{raw}」本应报错却通过了")


def build_score_cells(rows: int) -> dict[tuple[int, int], object]:
    """造一张带标题行的成绩表：第 1 行标题、第 2 行表头、之后是选手。"""
    cells: dict[tuple[int, int], object] = {
        (0, 0): "2026 校园科技创新大赛 决赛评分表",
        (1, 0): "编号", (1, 1): "姓名", (1, 2): "班级", (1, 3): "评委一", (1, 4): "评委二",
    }
    for index in range(rows):
        row = index + 2
        cells[(row, 0)] = f"C{index + 1:03d}"
        cells[(row, 1)] = f"选手{index + 1}"
        cells[(row, 2)] = "高二(1)班"
        cells[(row, 3)] = 90.0 + index * 0.1
        cells[(row, 4)] = 91.0 + index * 0.1
    return cells


def test_paging() -> None:
    print("2) 分页读取")
    cells = build_score_cells(450)
    # 真实接口的 row_count 是网格行数（比数据多），这里也用 600 模拟
    transport = FakeTransport(cells, rows=600, cols=26)
    reader = TencentSheet(transport, parse_link("https://docs.qq.com/sheet/DS29TWUFoYnNmbUxH"))
    rows = reader.read_sheet(SheetMeta("S1", "工作表1", 600, 26))

    check(len(rows) == 452, f"450 人 + 标题 + 表头 = 452 行（实际 {len(rows)}）")
    check(len(transport.reads) >= 3, f"至少分了 3 页读（实际 {len(transport.reads)} 次）")
    biggest = max((end - start) * (end_col - start_col)
                  for start, start_col, end, end_col in transport.reads)
    check(biggest <= 18000, f"单次请求最多 {biggest} 格，未超过 20000 上限")

    rows_empty = TencentSheet(FakeTransport({}), parse_link("DS29TWUFoYnNmbUxH"))
    check(rows_empty.read_sheet(SheetMeta("S1", "空表", 200, 26)) == [],
          "空表格返回空列表，不会硬拉 200 行")


def test_write_and_parse() -> None:
    print("3) 落盘 + 原有解析链路")
    work = Path(tempfile.mkdtemp(prefix="tdoc_check_"))
    try:
        cells = build_score_cells(3)
        transport = FakeTransport(cells)
        result = sync_to_cache("https://docs.qq.com/sheet/DS29TWUFoYnNmbUxH",
                               cache_dir=work, transport=transport)
        check(result.xlsx_path.exists(), f"生成了 {result.xlsx_path.name}")
        check(result.row_counts.get("工作表1") == 5,
              f"行数统计为 5（实际 {result.row_counts.get('工作表1')}）")
        check("腾讯文档" in result.describe() and "刷新" in result.describe(),
              f"状态文案可用：{result.describe()}")

        players, warnings, report = load_players(result.xlsx_path)
        check(len(players) == 3, f"解析出 3 名选手（实际 {len(players)}）")
        check(report.header_row == 2, f"表头定位到第 2 行（实际第 {report.header_row} 行）")
        check(players[0].name == "选手1" and abs(players[0].score - 181.0) < 1e-6,
              f"第一名分数 = 181.0（实际 {players[0].score}）")
        check(not warnings, f"没有缺分警告（实际 {warnings}）")

        entries = rank_players(players, winners=2)
        top = winners_of(entries)
        check(len(top) == 2 and top[0].player.name == "选手3",
              f"排名可用：第 1 名 {top[0].player.name if top else '—'}")

        # 刷新：同一份表格再同步一次，覆盖同一个缓存文件
        result2 = sync_to_cache("https://docs.qq.com/sheet/DS29TWUFoYnNmbUxH",
                                cache_dir=work, transport=transport)
        check(result2.xlsx_path == result.xlsx_path, "刷新写回同一个缓存文件（不会堆一堆）")
    finally:
        shutil.rmtree(work, ignore_errors=True)


def test_empty_doc() -> None:
    print("4) 空表格")
    work = Path(tempfile.mkdtemp(prefix="tdoc_empty_"))
    try:
        transport = FakeTransport({})
        try:
            sync_to_cache("DS29TWUFoYnNmbUxH", cache_dir=work, transport=transport)
        except DocNotFound:
            check(True, "空表格报 DocNotFound，不会生成空奖状")
        else:
            check(False, "空表格本应报错")
    finally:
        shutil.rmtree(work, ignore_errors=True)


def test_trailing_blanks() -> None:
    print("5) 尾部空行裁剪")
    cells = {(0, 0): "编号", (0, 1): "姓名", (0, 2): "总分",
             (1, 0): "C001", (1, 1): "张三", (1, 2): 95}
    work = Path(tempfile.mkdtemp(prefix="tdoc_trim_"))
    try:
        result = sync_to_cache("DS29TWUFoYnNmbUxH", cache_dir=work,
                               transport=FakeTransport(cells))
        check(result.row_counts.get("工作表1") == 2,
              f"空白行没被当成数据（统计 {result.row_counts.get('工作表1')} 行）")
    finally:
        shutil.rmtree(work, ignore_errors=True)


def test_write_xlsx_titles() -> None:
    print("6) 工作表名规整")
    work = Path(tempfile.mkdtemp(prefix="tdoc_title_"))
    try:
        path = work / "t.xlsx"
        write_xlsx(path, [("项目一/第一组", [["a"]]), ("项目一/第一组", [["b"]])])
        from openpyxl import load_workbook

        names = load_workbook(path).sheetnames
        check(len(names) == 2 and names[0] != names[1],
              f"非法字符被替换且重名自动加后缀：{names}")
    finally:
        shutil.rmtree(work, ignore_errors=True)


def run_live(url: str) -> int:
    print(f"实网拉取：{url}")
    try:
        result = sync_to_cache(url)
    except Exception as exc:  # noqa: BLE001
        print(f"  [!!] 失败：{type(exc).__name__}: {exc}")
        hint = getattr(exc, "hint", "")
        if hint:
            print(f"       建议：{hint}")
        return 1
    print(f"  [OK] {result.describe()}")
    print(f"       缓存：{result.xlsx_path}")
    players, warnings, report = load_players(result.xlsx_path)
    print(f"  [OK] 解析出 {len(players)} 行 · {report.describe()}")
    for warning in warnings[:3]:
        print(f"       ! {warning}")
    return 0


def main() -> int:
    if "--live" in sys.argv:
        index = sys.argv.index("--live")
        try:
            url = sys.argv[index + 1]
        except IndexError:
            print("用法：python tools/check_tencent.py --live <腾讯文档链接>")
            return 2
        return run_live(url)

    test_links()
    test_paging()
    test_write_and_parse()
    test_empty_doc()
    test_trailing_blanks()
    test_write_xlsx_titles()

    print()
    if failures:
        print(f"失败 {len(failures)} 项：")
        for item in failures:
            print(" - " + item)
        return 1
    print("全部通过 ✓")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
