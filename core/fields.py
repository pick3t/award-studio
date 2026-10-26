"""占位符字段映射：决定奖状里每个占位符最终填什么。

来源一共六种，界面里也是这六个选项：
* column       取成绩表里的某一列（也可以取工作表标题里的「项目组」）
* text         用户填写的固定文字
* system_date  当前系统日期
* fixed_date   用户指定日期
* award        奖项设置里用户自定义的奖项
* rank         按分数排出来的名次（第X名）
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime

# 占位符里常见的叫法与成绩表列名的对应关系（模板写「参赛项目」，表里可能叫「项目名称」）
COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "姓名": ("姓名",),
    "编号": ("编号", "序号"),
    "分数": ("分数",),
    "成绩": ("分数",),
    "组别": ("组别",),
    "项目名称": ("项目名称", "参赛项目"),
    "参赛项目": ("项目名称", "参赛项目"),
    "代表队": ("代表队", "单位"),
    "单位": ("单位", "代表队"),
    "项目组": ("项目组",),
}

SOURCE_LABELS = {
    "column": "表格列",
    "text": "固定值",
    "system_date": "系统日期",
    "fixed_date": "指定日期",
    "award": "奖项",
    "rank": "名次",
}
LABEL_TO_SOURCE = {label: key for key, label in SOURCE_LABELS.items()}

DATE_FORMAT = "%Y 年 %m 月 %d 日"

# 旧版本里「自动」的取值 -> 新来源
_AUTO_AWARD = {"奖项", "award"}
_AUTO_RANK = {"名次", "rank", "名次中文", "第几名"}
_AUTO_DATE = {"日期", "date"}
_AUTO_TEXT = {"比赛名称", "赛事名称", "contest", "颁发单位", "issuer"}


@dataclass
class FieldRule:
    """一个占位符的取值规则。"""

    key: str
    source: str = "column"
    value: str = ""

    @property
    def source_label(self) -> str:
        return SOURCE_LABELS.get(self.source, "表格列")

    def to_dict(self) -> dict:
        return {"key": self.key, "source": self.source, "value": self.value}

    @classmethod
    def from_dict(cls, data: dict) -> "FieldRule":
        source = str(data.get("source", "column"))
        if source not in SOURCE_LABELS:
            source = "text"
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


def column_for(key: str, columns: list[str] | None) -> str | None:
    """占位符名字该取哪一列：先找同名列，再按别名找。"""
    if not columns:
        return None
    if key in columns:
        return key
    for alias in COLUMN_ALIASES.get(key, ()):
        if alias in columns:
            return alias
    return None


def default_rule(key: str, columns: list[str] | None = None,
                 context=None) -> FieldRule:
    """给一个占位符挑一个合理的默认来源。"""
    if key in ("日期", "date"):
        return FieldRule(key, "system_date", "")
    if key in ("奖项", "award"):
        return FieldRule(key, "award", "")
    if key in ("名次", "rank", "名次中文", "第几名"):
        return FieldRule(key, "rank", "")
    if key in ("比赛名称", "赛事名称", "contest"):
        return FieldRule(key, "text", getattr(context, "contest", "") or "")
    if key in ("颁发单位", "issuer"):
        return FieldRule(key, "text", getattr(context, "issuer", "") or "")
    hit = column_for(key, columns)
    if hit:
        return FieldRule(key, "column", hit)
    return FieldRule(key, "text", "")


def _migrate_auto(key: str, choice: str, columns: list[str] | None) -> FieldRule:
    """把旧配置里的「自动」换算成新的六种来源之一。"""
    choice = (choice or "").strip() or key
    if choice in _AUTO_AWARD:
        return FieldRule(key, "award", "")
    if choice in _AUTO_RANK:
        return FieldRule(key, "rank", "")
    if choice in _AUTO_DATE:
        return FieldRule(key, "system_date", "")
    if choice in _AUTO_TEXT:
        return FieldRule(key, "text", "")
    hit = column_for(choice, columns)
    if hit:
        return FieldRule(key, "column", hit)
    return default_rule(key, columns)


def build_rules(keys: list[str], columns: list[str] | None = None,
                context=None, existing: dict[str, FieldRule] | None = None
                ) -> list[FieldRule]:
    """按模板里的占位符生成规则列表，已有的配置优先保留。"""
    existing = existing or {}
    rules: list[FieldRule] = []
    for key in keys:
        rule = existing.get(key)
        if rule is None:
            rules.append(default_rule(key, columns, context))
            continue
        if rule.source == "auto":  # 旧版本遗留
            rules.append(_migrate_auto(key, rule.value, columns))
            continue
        if rule.source == "column":
            # 换了工作表后原来的列名可能不存在了，按新表重新对号
            if columns and rule.value.strip() not in columns:
                hit = column_for(key, columns)
                if hit:
                    rule.value = hit
                else:
                    fresh = default_rule(key, columns, context)
                    rule.source, rule.value = fresh.source, fresh.value
            elif not rule.value.strip():
                rule.value = column_for(key, columns) or ""
        rules.append(rule)
    return rules


def resolve(rule: FieldRule, entry, context, today: date | None = None) -> str:
    """按规则算出这个占位符在某个选手身上应该填什么。"""
    if rule.source == "column":
        name = rule.value.strip()
        if name == "项目组":
            extra = getattr(context, "extra", None) or {}
            return str(extra.get("项目组") or entry.player.extras.get("项目组") or "")
        return str(entry.player.fields.get(name, ""))
    if rule.source == "text":
        return rule.value
    if rule.source == "system_date":
        return format_date(today or date.today())
    if rule.source == "fixed_date":
        return format_date(rule.value) if rule.value else format_date(today or date.today())
    if rule.source == "award":
        # 奖项设置里自定义的奖项；没设奖项时退回名次说法
        return entry.award or entry.rank_cn
    if rule.source == "rank":
        return entry.rank_cn
    return ""


def resolve_all(rules: list[FieldRule], entry, context,
                today: date | None = None) -> dict[str, str]:
    return {rule.key: resolve(rule, entry, context, today) for rule in rules}
