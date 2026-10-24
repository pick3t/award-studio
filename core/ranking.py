"""排名与奖项计算。"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, replace

from .excel_reader import Player

CN_DIGITS = "零一二三四五六七八九"


def cn_number(num: int) -> str:
    """1 -> 一, 12 -> 十二, 25 -> 二十五, 105 -> 105。"""
    if num < 0:
        return str(num)
    if num > 99:
        return str(num)
    if num < 10:
        return CN_DIGITS[num]
    tens, ones = divmod(num, 10)
    head = "十" if tens == 1 else f"{CN_DIGITS[tens]}十"
    return head if ones == 0 else head + CN_DIGITS[ones]


@dataclass
class RankedEntry:
    """排好名次的一名选手。"""

    rank: int
    player: Player
    award: str
    is_winner: bool = False

    @property
    def rank_cn(self) -> str:
        return f"第{cn_number(self.rank)}名"


def award_for_rank(rank: int, winners: int, use_tiers: bool = True) -> str:
    """默认的三等分规则：一等奖 / 二等奖 / 三等奖。"""
    if not use_tiers or winners < 3:
        return "一等奖"
    first_cut = max(1, math.ceil(winners / 3))
    second_cut = max(first_cut + 1, math.ceil(winners * 2 / 3))
    if rank <= first_cut:
        return "一等奖"
    if rank <= second_cut:
        return "二等奖"
    return "三等奖"


def build_slots(winners: int, use_tiers: bool, award_tiers=None) -> list[str]:
    """按奖项设置生成「每个名额对应什么奖」的列表。

    award_tiers 形如 [("特等奖", 1), ("一等奖", 3)]，按顺序占名额。
    不设奖项时返回空列表，此时名额由 winners 决定。
    """
    if not use_tiers:
        return []
    if award_tiers:
        slots: list[str] = []
        for name, count in award_tiers:
            try:
                number = int(count)
            except (TypeError, ValueError):
                continue
            name = str(name).strip() or "奖项"
            slots.extend([name] * max(0, number))
        return slots
    return [award_for_rank(position, winners, True) for position in range(1, max(0, winners) + 1)]


def rank_players(
    players: list[Player],
    winners: int = 5,
    use_tiers: bool = True,
    award_tiers=None,
    ranked: bool = True,
) -> list[RankedEntry]:
    """按分数降序排名，同分并列（1、2、2、4）。

    award_tiers 为自定义奖项（名称 + 人数），按排序位置依次分配；
    传 None 时沿用三等分；use_tiers=False 时只排名次、不留奖项。
    ranked=False 时不排序、不分奖，按表格原始行顺序返回。
    """
    if not ranked:
        return [RankedEntry(rank=0, player=p, award="", is_winner=False) for p in players]

    ordered = sorted(players, key=lambda p: (-p.score, p.name))
    slots = build_slots(winners, use_tiers, award_tiers)
    quota = len(slots) if slots else max(0, winners)

    entries: list[RankedEntry] = []
    previous_score: float | None = None
    previous_rank = 0

    for index, player in enumerate(ordered, start=1):
        if previous_score is not None and abs(player.score - previous_score) < 1e-9:
            rank = previous_rank
        else:
            rank = index
        previous_score, previous_rank = player.score, rank
        entries.append(
            RankedEntry(rank=rank, player=player, award="", is_winner=False)
        )

    # 同分必须同奖：按分数分组整体分配，避免同分的人被拆到不同奖项
    index = 0
    while index < len(ordered):
        end = index + 1
        while end < len(ordered) and abs(ordered[end].score - ordered[index].score) < 1e-9:
            end += 1
        in_quota = index < quota
        award = slots[index] if in_quota and index < len(slots) else ""
        for position in range(index, end):
            entries[position].award = award
            entries[position].is_winner = in_quota
        index = end
    return entries


def award_conflicts(entries: list[RankedEntry], award_tiers) -> list[dict]:
    """实际发放人数与配置不一致的奖项（多因同分并列导致）。"""
    counts: dict[str, int] = {}
    for entry in entries:
        if entry.is_winner and entry.award:
            counts[entry.award] = counts.get(entry.award, 0) + 1

    conflicts: list[dict] = []
    for item in award_tiers or []:
        try:
            name, configured = str(item[0]), int(item[1])
        except (TypeError, ValueError, IndexError):
            continue
        actual = counts.get(name, 0)
        if actual != configured:
            conflicts.append({"name": name, "configured": configured, "actual": actual})
    return conflicts


def winners_of(entries: list[RankedEntry], winners: int | None = None) -> list[RankedEntry]:
    """取获奖者（按排序位置，而不是名次号，避免并列导致超额）。

    winners 为 None 时直接返回全部获奖者，不额外截断。
    """
    result = [e for e in entries if e.is_winner]
    if winners is None:
        return result
    return result[:winners]


MEMBER_KEYS = ("参赛选手", "队员", "成员", "选手名单")


def split_members(text: str) -> list[str]:
    """把「马兴丽 胡菊园 张锐」拆成多人。

    表里用空格分隔队员，但两字姓名中间也常多打空格（张  锐），
    而分隔符也有时是两个空格（翁素维  张月珍），单看空格数无法区分。
    所以按任意空白切成 token 后，只把「单字 + 单字」粘起来：
    中文姓名至少两个字，单字 token 必定是被空格劈开的两字名。
    """
    if not text:
        return []
    cleaned = str(text).replace("\u3000", " ").replace("\n", " ").replace("\r", " ")
    tokens = [token for token in cleaned.split(" ") if token]
    merged: list[str] = []
    index = 0
    while index < len(tokens):
        token = tokens[index]
        if len(token) == 1 and index + 1 < len(tokens) and len(tokens[index + 1]) == 1:
            merged.append(token + tokens[index + 1])
            index += 2
            continue
        merged.append(token)
        index += 1
    return merged


def members_of(entry: RankedEntry) -> list[str]:
    """这一条记录里包含的队员名单（个人项目返回空列表）。"""
    for key in MEMBER_KEYS:
        value = entry.player.extras.get(key, "")
        members = split_members(value)
        if len(members) > 1:
            return members
    return []


def expand_entries(entries: list[RankedEntry]) -> list[RankedEntry]:
    """集体项目：一队多人时拆成每人一条，保证一人一份奖状。"""
    expanded: list[RankedEntry] = []
    for entry in entries:
        members = members_of(entry)
        if not members:
            expanded.append(entry)
            continue
        team = entry.player.name
        roster = " ".join(members)
        for position, name in enumerate(members, start=1):
            extras = dict(entry.player.extras)
            extras["代表队"] = team
            extras["参赛选手"] = roster
            # 一队多人时证书编号加后缀，避免每份都写同一个号
            no = entry.player.no
            number = f"{no}-{position}" if no else str(position)
            player = replace(entry.player, name=name, no=number, extras=extras)
            expanded.append(replace(entry, player=player))
    return expanded


def certificate_total(entries: list[RankedEntry]) -> int:
    """这些获奖记录最终要生成多少份奖状（集体项目按人算）。"""
    return len(expand_entries(entries))


def safe_filename(text: str, fallback: str = "certificate") -> str:
    """把姓名等文本变成安全的文件名。"""
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(text)).strip().strip(".")
    return cleaned or fallback
