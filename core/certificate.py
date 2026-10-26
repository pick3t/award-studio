"""奖状生成：把名次、姓名写进 Word 模板。

模板占位符写法（任选其一，支持中英文冒号无关）：
    {{姓名}}    【姓名】     {{名次}}    【奖项】

支持的字段见 build_mapping()，另可用表格里的任意附加列（如 {{班级}}）。
占位符被 Word 拆成多个 run（格式片段）时也能正确替换。
"""

from __future__ import annotations

import re
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from . import fields as field_rules
from .ranking import RankedEntry, expand_entries, safe_filename

PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]{1,40}?)\s*\}\}|【\s*([^【】]{1,20}?)\s*】")


class CertificateError(Exception):
    """奖状生成失败。"""


@dataclass
class CertificateContext:
    """奖状里除选手信息外的公共内容。"""

    contest: str = "技能竞赛"
    issuer: str = ""
    date: str = ""
    extra: dict[str, str] = field(default_factory=dict)


def _lookup(mapping: dict[str, str], key: str) -> str | None:
    """按忽略大小写/空格的方式查字段。"""
    normalized = re.sub(r"[\s:：]", "", key).lower()
    for name, value in mapping.items():
        if re.sub(r"[\s:：]", "", name).lower() == normalized:
            return value
    return None


def _apply_value(runs, start: int, end: int, value: str) -> None:
    """把 [start, end) 区间的文本替换为 value，保留起始 run 的格式。"""
    bounds = []
    cursor = 0
    for run in runs:
        length = len(run.text)
        bounds.append((cursor, cursor + length))
        cursor += length

    start_idx = next((i for i, (a, b) in enumerate(bounds) if a <= start < b), None)
    end_idx = next((i for i, (a, b) in enumerate(bounds) if a < end <= b), None)
    if start_idx is None or end_idx is None:
        return

    prefix = runs[start_idx].text[: start - bounds[start_idx][0]]
    suffix = runs[end_idx].text[end - bounds[end_idx][0] :]

    if start_idx == end_idx:
        runs[start_idx].text = prefix + value + suffix
        return

    runs[start_idx].text = prefix + value
    for middle in range(start_idx + 1, end_idx):
        runs[middle].text = ""
    runs[end_idx].text = suffix


def _replace_in_paragraph(paragraph, mapping: dict[str, str]) -> int:
    """替换一个段落里的占位符，返回替换次数（可跨 run）。"""
    count = 0
    for _ in range(400):
        runs = [r for r in paragraph.runs if r._element.getparent() is not None]
        if not runs:
            break
        text = "".join(r.text for r in runs)
        if not text:
            break
        replaced = False
        for match in PLACEHOLDER_RE.finditer(text):
            key = match.group(1) or match.group(2) or ""
            value = _lookup(mapping, key)
            # 映射里没有的占位符按空值处理，避免生成的证书里留着花括号
            _apply_value(runs, match.start(), match.end(), "" if value is None else str(value))
            count += 1
            replaced = True
            break
        if not replaced:
            break
    return count


def _iter_paragraphs(document):
    """遍历正文、表格（含嵌套）、页眉页脚的段落。"""
    def walk_tables(tables):
        for table in tables:
            for row in table.rows:
                seen = set()
                for cell in row.cells:
                    if id(cell._tc) in seen:
                        continue
                    seen.add(id(cell._tc))
                    yield from cell.paragraphs
                    yield from walk_tables(cell.tables)

    yield from document.paragraphs
    yield from walk_tables(document.tables)
    for section in document.sections:
        for part in (section.header, section.footer, section.first_page_header,
                     section.first_page_footer, section.even_page_header, section.even_page_footer):
            try:
                yield from part.paragraphs
                yield from walk_tables(part.tables)
            except Exception:  # noqa: BLE001 - 某些 section 没有对应页眉页脚
                continue


def inspect_template(path: str | Path) -> list[str]:
    """读出模板里用到的占位符，按在文档中出现的先后顺序返回。"""
    document = Document(str(path))
    keys: list[str] = []
    for paragraph in _iter_paragraphs(document):
        text = "".join(run.text for run in paragraph.runs)
        for match in PLACEHOLDER_RE.finditer(text):
            key = (match.group(1) or match.group(2) or "").strip()
            if key and key not in keys:
                keys.append(key)
    return keys


def build_mapping(entry: RankedEntry, context: CertificateContext) -> dict[str, str]:
    """一名选手对应的全部可填字段。"""
    player = entry.player
    mapping = {
        "姓名": player.name,
        "name": player.name,
        "选手": player.name,
        "名次": str(entry.rank),
        "rank": str(entry.rank),
        "名次中文": entry.rank_cn,
        "第几名": entry.rank_cn,
        # 不设奖项时用它填「第一名」这样的名次说法，模板里写「荣获 {{奖项}}」两种模式都通顺
        "奖项": entry.award or entry.rank_cn,
        "award": entry.award or entry.rank_cn,
        "编号": player.no,
        "no": player.no,
        "比赛名称": context.contest,
        "赛事名称": context.contest,
        "contest": context.contest,
        "颁发单位": context.issuer,
        "issuer": context.issuer,
        "日期": context.date,
        "date": context.date,
    }
    mapping.update(context.extra)
    mapping.update(player.extras)
    # 模板里可能写 {{项目组}} / {{参赛项目}}，而表头是「组别」「项目名称」或反过来
    for alias, keys in (
        ("项目组", ("组别", "项目组", "组")),
        ("参赛项目", ("项目名称", "参赛项目", "项目")),
        ("代表队", ("代表队", "单位")),
    ):
        if mapping.get(alias):
            continue
        for key in keys:
            value = mapping.get(key)
            if value:
                mapping[alias] = value
                break
    return mapping


def generate_certificate(
    template: str | Path,
    output: str | Path,
    entry: RankedEntry,
    context: CertificateContext,
    rules: list | None = None,
) -> int:
    """用模板生成一份奖状，返回替换的占位符数量。"""
    template_path = Path(template)
    output_path = Path(output)
    if not template_path.exists():
        raise CertificateError(f"奖状模板不存在：{template_path}")

    try:
        document = Document(str(template_path))
    except Exception as exc:  # noqa: BLE001
        raise CertificateError(f"无法打开奖状模板：{exc}") from exc

    mapping = build_mapping(entry, context)
    if rules:
        # 用户配置的字段映射优先于内置规则
        mapping.update(field_rules.resolve_all(rules, entry, context))
    replaced = 0
    for paragraph in _iter_paragraphs(document):
        replaced += _replace_in_paragraph(paragraph, mapping)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        document.save(str(output_path))
    except PermissionError as exc:
        raise CertificateError(f"无法写入 {output_path.name}，请先关闭已打开的同名文件") from exc
    if replaced == 0:
        raise CertificateError(
            f"模板里没有发现可替换的占位符（请在模板中写 {{姓名}}、{{名次}} 等）"
        )
    return replaced


def certificate_filename(entry: RankedEntry) -> str:
    """奖状文件名：01_张三_一等奖.docx

    集体项目拆成每人一份时带上代表队：01_一队_张三_一等奖.docx
    （不设奖项时省略奖项部分）
    """
    parts = [f"{entry.rank:02d}"]
    team = entry.player.extras.get("代表队", "")
    name = entry.player.name
    if team and " " in name:
        # 一队一份：姓名栏是全体队员，文件名只用队名，避免名字过长
        parts.append(safe_filename(team))
    elif team and team != name:
        parts.append(safe_filename(team))
        parts.append(safe_filename(name))
    else:
        parts.append(safe_filename(name))
    if entry.award:
        parts.append(safe_filename(entry.award))
    return "_".join(parts) + ".docx"


def merge_documents(paths: list[Path], output: str | Path) -> None:
    """把多份奖状合并成一个文件（每份一页，保留图片）。"""
    if not paths:
        raise CertificateError("没有可合并的奖状")
    master = Document(str(paths[0]))
    body = master.element.body
    sect_pr = body.find(qn("w:sectPr"))

    def append(element) -> None:
        if sect_pr is not None:
            sect_pr.addprevious(element)
        else:
            body.append(element)

    for path in paths[1:]:
        source = Document(str(path))
        page_break = OxmlElement("w:p")
        run = OxmlElement("w:r")
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run.append(br)
        page_break.append(run)
        append(page_break)

        for child in list(source.element.body):
            if child.tag == qn("w:sectPr"):
                continue
            new_child = deepcopy(child)
            for blip in new_child.iter(qn("a:blip")):
                rid = blip.get(qn("r:embed"))
                if rid and rid in source.part.related_parts:
                    image_part = source.part.related_parts[rid]
                    try:
                        new_rid = master.part.relate_to(image_part, RT.IMAGE)
                        blip.set(qn("r:embed"), new_rid)
                    except Exception:  # noqa: BLE001 - 图片复制失败时保留空引用
                        pass
            append(new_child)

    master.save(str(output))


def generate_all(
    template: str | Path,
    output_dir: str | Path,
    winners: list[RankedEntry],
    context: CertificateContext,
    merge: bool = False,
    merged_name: str = "全部奖状_合并版.docx",
    rules: list | None = None,
    per_person: bool = True,
) -> list[Path]:
    """批量生成奖状，返回生成的文件列表。"""
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    created: list[Path] = []
    # 集体项目：勾选「每人一份」时拆开，否则一队一份（姓名栏放全体队员）
    for entry in expand_entries(winners, per_person):
        path = out_dir / certificate_filename(entry)
        if path.exists():
            stem, suffix, index = path.stem, path.suffix, 2
            while path.exists():
                path = out_dir / f"{stem}_{index}{suffix}"
                index += 1
        generate_certificate(template, path, entry, context, rules=rules)
        created.append(path)

    if merge and created:
        merged = out_dir / merged_name
        try:
            merge_documents(created, merged)
            created.append(merged)
        except Exception as exc:  # noqa: BLE001 - 合并失败不影响单人文件
            raise CertificateError(f"合并版生成失败（单人奖状已生成）：{exc}") from exc
    return created
