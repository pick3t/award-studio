"""生成一批覆盖不同场景的测试数据，放在 samples/ 下。

用法：
    python tools/make_more_samples.py [输出目录]

生成 6 张成绩表和 4 份奖状模板，用来验证表格识别、并列名次、大名单、
以及不同版式的模板填充。
"""

from __future__ import annotations

import random
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

from make_samples import CLASSES, PROJECTS, SURNAMES, GIVEN, _set_font  # noqa: E402

THIN = Side(style="thin", color="C9CFDD")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="EEF1F8")


def _names(count: int, seed: int) -> list[str]:
    rng = random.Random(seed)
    used: set[str] = set()
    result: list[str] = []
    while len(result) < count:
        name = rng.choice(SURNAMES) + rng.choice(GIVEN)
        if name not in used:
            used.add(name)
            result.append(name)
    return result


def _write_sheet(path: Path, headers: list[str], rows: list[list], *,
                 title_lines: list[str] | None = None,
                 widths: list[int] | None = None,
                 number_format: str | None = None) -> Path:
    """写一张带样式的成绩表；title_lines 会占用表头之前的行。"""
    book = Workbook()
    sheet = book.active
    sheet.title = "成绩"

    row_index = 1
    for line in title_lines or []:
        cell = sheet.cell(row=row_index, column=1, value=line)
        cell.font = Font(size=14, bold=(row_index == 1))
        sheet.merge_cells(start_row=row_index, start_column=1,
                          end_row=row_index, end_column=max(2, len(headers)))
        cell.alignment = Alignment(horizontal="center", vertical="center")
        sheet.row_dimensions[row_index].height = 26
        row_index += 1

    header_row = row_index
    for col, name in enumerate(headers, start=1):
        cell = sheet.cell(row=header_row, column=col, value=name)
        cell.font = Font(bold=True, size=11)
        cell.fill = HEAD_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[header_row].height = 24

    for offset, values in enumerate(rows, start=1):
        for col, value in enumerate(values, start=1):
            cell = sheet.cell(row=header_row + offset, column=col, value=value)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if number_format and isinstance(value, (int, float)):
                cell.number_format = number_format

    for col in range(1, len(headers) + 1):
        width = (widths or [])[col - 1] if widths and col <= len(widths) else 12
        sheet.column_dimensions[get_column_letter(col)].width = width

    sheet.freeze_panes = sheet.cell(row=header_row + 1, column=1)
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)
    return path


# ---------------------------------------------------------------- 六张成绩表
def make_two_column(path: Path) -> Path:
    """最简：只有姓名和总分，没有标题行。"""
    names = _names(10, 101)
    rng = random.Random(101)
    rows = [[name, round(rng.uniform(62, 98), 1)] for name in names]
    return _write_sheet(path, ["姓名", "总分"], rows, widths=[16, 12], number_format="0.0")


def make_with_notes(path: Path) -> Path:
    """表头在第四行，前面有标题和说明。"""
    names = _names(20, 202)
    rng = random.Random(202)
    rows = []
    for index, name in enumerate(names, start=1):
        rows.append([f"S{index:02d}", name, rng.choice(CLASSES),
                     round(rng.uniform(70, 97), 1)])
    return _write_sheet(
        path, ["序号", "姓名", "班级", "总分"], rows,
        title_lines=["2026 年校园技能节 · 决赛成绩表",
                     "（本表由评委现场打分汇总，如有疑问请于赛后三日内反馈）",
                     "制表：教务处"],
        widths=[10, 16, 18, 12], number_format="0.0",
    )


def make_judges_only(path: Path) -> Path:
    """没有总分组，只有三位评委，需要自动求和。"""
    names = _names(15, 303)
    rng = random.Random(303)
    rows = []
    for index, name in enumerate(names, start=1):
        # 基准分限制在 62–93，避免加上评委偏差后被 100 截断成一模一样的满分
        base = min(93.0, max(62.0, rng.gauss(84, 7)))
        judges = [round(min(99.5, max(55, base + rng.gauss(0, 3))), 1) for _ in range(3)]
        rows.append([f"J{index:02d}", name, *judges])
    return _write_sheet(
        path, ["编号", "姓名", "评委一", "评委二", "评委三"], rows,
        widths=[10, 16, 12, 12, 12], number_format="0.0",
    )


def make_big_roster(path: Path) -> Path:
    """100 人，用来测试大名单与名额滑块。"""
    names = _names(100, 404)
    rng = random.Random(404)
    rows = []
    for index, name in enumerate(names, start=1):
        rows.append([f"B{index:03d}", name, rng.choice(CLASSES),
                     rng.choice(PROJECTS), round(rng.uniform(60, 99), 1)])
    return _write_sheet(
        path, ["编号", "姓名", "班级", "作品名称", "总分"], rows,
        title_lines=["2026 全校创新大赛 · 初赛成绩"],
        widths=[10, 16, 18, 26, 12], number_format="0.0",
    )


def make_tied_scores(path: Path) -> Path:
    """2–4 名刻意同分，用来测试并列名次。"""
    names = _names(10, 505)
    scores = [97.5, 95.0, 95.0, 95.0, 92.4, 90.1, 88.8, 88.8, 85.2, 81.6]
    rows = [[f"T{index:02d}", name, score] for index, (name, score) in
            enumerate(zip(names, scores), start=1)]
    return _write_sheet(
        path, ["编号", "姓名", "总分"], rows, widths=[10, 16, 12],
        number_format="0.0",
    )


def make_alias_columns(path: Path) -> Path:
    """列名用别名，分数带「分」字单位。"""
    names = _names(8, 606)
    rng = random.Random(606)
    rows = []
    for index, name in enumerate(names, start=1):
        rows.append([index, name, f"{round(rng.uniform(70, 96), 1)} 分",
                     rng.choice(CLASSES)])
    return _write_sheet(
        path, ["序号", "选手", "得分", "班级"], rows,
        title_lines=["（本表列名与常规表不同：选手 / 得分）"],
        widths=[10, 16, 14, 18],
    )


# ---------------------------------------------------------------- 四份模板
def _setup_section(document: Document, landscape: bool = True) -> None:
    section = document.sections[0]
    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = Cm(29.7), Cm(21.0)
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width, section.page_height = Cm(21.0), Cm(29.7)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)


def _border(section, color: str, size: int = 12) -> None:
    sect_pr = section._sectPr
    borders = OxmlElement("w:pgBorders")
    borders.set(qn("w:offsetFrom"), "page")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), str(size))
        element.set(qn("w:space"), "24")
        element.set(qn("w:color"), color)
        borders.append(element)
    sect_pr.append(borders)


def _para(document, text: str, size: float, *, align=WD_ALIGN_PARAGRAPH.CENTER,
          bold=False, color="1A1A1A", before=0, after=0, font="宋体"):
    paragraph = document.add_paragraph()
    paragraph.alignment = align
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    run = paragraph.add_run(text)
    _set_font(run, font if font != "宋体" else "Times New Roman", size, bold=bold,
              color=color, eastasia=font)
    return paragraph


def make_portrait_template(path: Path) -> Path:
    """竖版 A4，细线边框，适合证书。"""
    document = Document()
    _setup_section(document, landscape=False)
    _border(document.sections[0], "9C2B2B", 8)

    _para(document, "", 10, after=30)
    _para(document, "荣 誉 证 书", 34, bold=True, color="9C2B2B", after=6)
    _para(document, "CERTIFICATE OF HONOR", 10, color="B08D57", after=44)
    _para(document, "兹 有", 16, font="楷体", after=10)
    _para(document, "{{姓名}}", 26, bold=True, color="9C2B2B", after=10)
    _para(document, "同 学", 16, font="楷体", after=18)
    _para(document, "在「{{比赛名称}}」中荣获", 16, font="楷体", after=10)
    _para(document, "{{奖项}}", 22, bold=True, color="9C2B2B", after=26)
    _para(document, "特发此证，以资鼓励", 16, font="楷体", after=54)
    _para(document, "{{颁发单位}}", 14, align=WD_ALIGN_PARAGRAPH.RIGHT,
          font="楷体", after=4)
    _para(document, "{{日期}}", 14, align=WD_ALIGN_PARAGRAPH.RIGHT, font="楷体")

    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(path)
    return path


def _badge_image(size: int = 320) -> Path:
    """画一个圆形徽章，用作模板里的校徽占位。"""
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.ellipse([6, 6, size - 6, size - 6], outline=(156, 43, 43), width=10)
    draw.ellipse([34, 34, size - 34, size - 34], outline=(176, 141, 87), width=4)
    font = None
    for candidate in ("C:/Windows/Fonts/simhei.ttf", "C:/Windows/Fonts/msyh.ttc"):
        try:
            font = ImageFont.truetype(candidate, 120)
            break
        except OSError:
            continue
    if font is None:
        font = ImageFont.load_default()
    draw.text((size / 2, size / 2), "奖", font=font, fill=(156, 43, 43), anchor="mm")

    temp = Path(tempfile.gettempdir()) / "award_badge.png"
    image.save(temp)
    return temp


def make_badge_template(path: Path) -> Path:
    """横版，左上角带徽章，落款用中文日期与颁发单位。"""
    document = Document()
    _setup_section(document, landscape=True)
    _border(document.sections[0], "B08D57", 18)

    badge = _badge_image()
    header = document.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header.add_run().add_picture(str(badge), width=Cm(2.4))
    header.paragraph_format.space_after = Pt(6)

    _para(document, "荣 誉 证 书", 36, bold=True, color="9C2B2B", after=30)
    _para(document, "{{姓名}}  同学：", 18, align=WD_ALIGN_PARAGRAPH.LEFT,
          font="楷体", after=12)
    body = document.add_paragraph()
    body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.first_line_indent = Pt(36)
    body.paragraph_format.line_spacing = 2.2
    for text, size, bold, color, eastasia in (
        ("你在「", 17, False, "1A1A1A", "楷体"),
        ("{{比赛名称}}", 18, True, "1A1A1A", "黑体"),
        ("」中表现突出，经评审委员会评定，荣获 ", 17, False, "1A1A1A", "楷体"),
        ("{{奖项}}", 20, True, "9C2B2B", "黑体"),
        ("（总排名第 ", 17, False, "1A1A1A", "楷体"),
        ("{{名次}}", 20, True, "9C2B2B", "黑体"),
        (" 名）。", 17, False, "1A1A1A", "楷体"),
    ):
        run = body.add_run(text)
        _set_font(run, "Times New Roman" if eastasia != "楷体" else "楷体",
                  size, bold=bold, color=color, eastasia=eastasia)

    _para(document, "", 12, after=24)
    _para(document, "{{颁发单位}}", 15, align=WD_ALIGN_PARAGRAPH.RIGHT,
          font="楷体", after=4)
    _para(document, "{{日期}}", 15, align=WD_ALIGN_PARAGRAPH.RIGHT, font="楷体")

    try:
        badge.unlink()
    except OSError:
        pass
    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(path)
    return path


def make_bracket_template(path: Path) -> Path:
    """方括号占位符（【姓名】），无边框居中排版。"""
    document = Document()
    _setup_section(document, landscape=True)
    section = document.sections[0]
    section.top_margin = Cm(2.6)
    section.bottom_margin = Cm(2.4)

    _para(document, "", 12, after=36)
    _para(document, "获  奖  证  书", 40, bold=True, after=40)
    _para(document, "【姓名】", 24, bold=True, after=16)
    _para(document, "在【比赛名称】中获得", 17, font="楷体", after=12)
    _para(document, "【奖项】", 24, bold=True, color="9C2B2B", after=40)
    _para(document, "特发此证", 17, font="楷体", after=46)
    _para(document, "【颁发单位】　　　【日期】", 15, font="楷体")

    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(path)
    return path


def _shade(cell, color: str) -> None:
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), color)
    cell._tc.get_or_add_tcPr().append(shading)


def make_table_template(path: Path) -> Path:
    """用表格排版的证书，验证占位符落在表格单元格里也能替换。"""
    document = Document()
    _setup_section(document, landscape=True)
    section = document.sections[0]
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.0)

    _para(document, "获奖证书", 30, bold=True, color="1F3C88", after=18)

    table = document.add_table(rows=4, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    labels = ["姓　　名", "班　　级", "参赛项目", "获奖等级"]
    values = ["{{姓名}}", "{{班级}}", "{{作品名称}}", "{{奖项}}"]
    for row, (label, value) in enumerate(zip(labels, values)):
        left, right = table.rows[row].cells
        left.text = ""
        right.text = ""
        left_run = left.paragraphs[0].add_run(label)
        _set_font(left_run, "Times New Roman", 13, bold=True, eastasia="黑体")
        left.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _shade(left, "F2F5FB")
        right_run = right.paragraphs[0].add_run(value)
        _set_font(right_run, "Times New Roman", 13, eastasia="宋体")
        right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    _para(document, "", 12, after=20)
    _para(document, "{{比赛名称}}　{{颁发单位}}", 14, font="楷体", after=4)
    _para(document, "{{日期}}", 14, font="楷体")

    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(path)
    return path


def main() -> None:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "samples"
    out.mkdir(parents=True, exist_ok=True)

    sheets = [
        ("成绩表_两列最简.xlsx", make_two_column),
        ("成绩表_带说明行.xlsx", make_with_notes),
        ("成绩表_多评委无总分.xlsx", make_judges_only),
        ("成绩表_100人大名单.xlsx", make_big_roster),
        ("成绩表_同分并列.xlsx", make_tied_scores),
        ("成绩表_列名别名.xlsx", make_alias_columns),
    ]
    for name, builder in sheets:
        path = builder(out / name)
        print(f"表格：{path.name}  ({path.stat().st_size / 1024:.1f} KB)")

    templates = [
        ("模板_竖版简约.docx", make_portrait_template),
        ("模板_横版带徽章.docx", make_badge_template),
        ("模板_方括号占位.docx", make_bracket_template),
        ("模板_表格排版.docx", make_table_template),
    ]
    for name, builder in templates:
        path = builder(out / name)
        print(f"模板：{path.name}  ({path.stat().st_size / 1024:.1f} KB)")


if __name__ == "__main__":
    main()
