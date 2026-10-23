"""占位符字段映射：决定奖状里每个占位符最终填什么。

每个占位符对应一条规则，来源有四种：
* auto         程序内置计算（姓名、名次、奖项……）
* column       取成绩表里的某一列
* text         用户填写的固定文字
* system_date  当前系统日期
* fixed_date   用户指定日期
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime

# 占位符 -> 内置取值的说明
AUTO_FIELDS: dict[str, str] = {
    "姓名": "选手姓名",
    "name": "选手姓名",
    "选手": "选手姓名",
    "名次": "名次数字",
    "rank": "名次数字",
    "名次中文": "中文名次",
    "第几名": "中文名次",
    "奖项": "奖项等级（不设奖项时为名次）",
    "award": "奖项等级",
    "编号": "选手编号",
    "no": "选手编号",
    "比赛名称": "比赛名称",
    "赛事名称": "比赛名称",
    "contest": "比赛名称",
    "颁发单位": "颁发单位",
    "issuer": "颁发单位",
    "日期": "日期",
    "date": "日期",
    # 成绩单里叫法不同：个人表是「组别 + 项目名称」，集体表是「参赛项目」
    "项目组": "项目组 / 组别",
    "组别": "项目组 / 组别",
    "参赛项目": "参赛项目名称",
    "项目名称": "参赛项目名称",
    "代表队": "代表队 / 单位",
}

SOURCE_LABELS = {
    "auto": "自动",
    "column": "表格列",
    "text": "固定值",
    "system_date": "系统日期",
    "fixed_date": "指定日期",
}
LABEL_TO_SOURCE = {label: key for key, label in SOURCE_LABELS.items()}

DATE_FORMAT = "%Y 年 %m 月 %d 日"


@dataclass
class FieldRule:
    """一个占位符的取值规则。"""

    key: str
    source: str = "auto"
    value: str = ""

    @property
    def source_label(self) -> str:
        return SOURCE_LABELS.get(self.source, "自动")

    def to_dict(self) -> dict:
        return {"key": self.key, "source": self.source, "value": self.value}

    @classmethod
    def from_dict(cls, data: dict) -> "FieldRule":
        source = str(data.get("source", "auto"))
        if source not in SOURCE_LABELS:
            source = "auto"
        return cls(key=str(data.get("key", "")), source=source,
                   value=str(data.get("value", "")))


def format_date(value: str | date | datetime) -> str:
    """把日期统一成「2026 年 10 月 21 日」。"""
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return value.strftime(DATE_FORMAT)
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d", DATE_FORMAT):
        try:
            return datetime.strptime(text, fmt).strftime(DATE_FORMAT)
        except ValueError:
            continue
    return text


def default_rule(key: str, columns: list[str] | None = None,
                 context=None) -> FieldRule:
    """给一个占位符挑一个合理的默认来源。"""
    if key in AUTO_FIELDS:
        if key in ("日期", "date"):
            return FieldRule(key, "system_date", "")
        if key in ("颁发单位", "issuer"):
            return FieldRule(key, "text", getattr(context, "issuer", "") or "")
        if key in ("比赛名称", "赛事名称", "contest"):
            return FieldRule(key, "text", getattr(context, "contest", "") or "")
        return FieldRule(key, "auto", "")
    if columns and key in columns:
        return FieldRule(key, "column", key)
    return FieldRule(key, "text", "")


def build_rules(keys: list[str], columns: list[str] | None = None,
                context=None, existing: dict[str, FieldRule] | None = None
                ) -> list[FieldRule]:
    """按模板里的占位符生成规则列表，已有的配置优先保留。"""
    existing = existing or {}
    rules: list[FieldRule] = []
    for key in keys:
        if key in existing:
            rules.append(existing[key])
        else:
            rules.append(default_rule(key, columns, context))
    return rules


def auto_value(key: str, entry, context) -> str:
    """内置字段的取值。"""
    player = entry.player
    normalized = key.strip().lower()
    if normalized in ("姓名", "name", "选手"):
        return player.name
    if normalized in ("名次", "rank"):
        return str(entry.rank)
    if normalized in ("名次中文", "第几名"):
        return entry.rank_cn
    if normalized in ("奖项", "award"):
        return entry.award or entry.rank_cn
    if normalized in ("编号", "no"):
        return player.no
    if normalized in ("比赛名称", "赛事名称", "contest"):
        return context.contest
    if normalized in ("颁发单位", "issuer"):
        return context.issuer
    if normalized in ("日期", "date"):
        return context.date
    if normalized in ("项目组", "组"):
        extra = getattr(context, "extra", None) or {}
        return str(extra.get("项目组") or player.extras.get("项目组") or "")
    if normalized == "组别":
        return str(player.extras.get("组别") or "")
    if normalized in ("参赛项目", "项目名称", "项目"):
        return str(player.extras.get("项目名称") or player.extras.get("参赛项目") or "")
    if normalized in ("代表队", "参赛队", "队名", "单位"):
        # 个人项目没有代表队，用「单位」顶上
        return str(player.extras.get("代表队") or player.extras.get("单位") or "")
    return ""


def resolve(rule: FieldRule, entry, context, today: date | None = None) -> str:
    """按规则算出这个占位符在某个选手身上应该填什么。"""
    if rule.source == "auto":
        return auto_value(rule.key, entry, context)
    if rule.source == "column":
        return str(entry.player.fields.get(rule.value, ""))
    if rule.source == "text":
        return rule.value
    if rule.source == "system_date":
        return format_date(today or date.today())
    if rule.source == "fixed_date":
        return format_date(rule.value) if rule.value else format_date(today or date.today())
    return ""


def resolve_all(rules: list[FieldRule], entry, context,
                today: date | None = None) -> dict[str, str]:
    return {rule.key: resolve(rule, entry, context, today) for rule in rules}
