"""排名与奖项计算。"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass

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
    """按名次给出奖项名称。

    前 n 名三等分：一等奖 / 二等奖 / 三等奖。n < 3 或关闭分档时统一为一等奖。
    """
    if not use_tiers or winners < 3:
        return "一等奖"
    first_cut = max(1, math.ceil(winners / 3))
    second_cut = max(first_cut + 1, math.ceil(winners * 2 / 3))
    if rank <= first_cut:
        return "一等奖"
    if rank <= second_cut:
        return "二等奖"
    return "三等奖"


def rank_players(players: list[Player], winners: int = 5, use_tiers: bool = True) -> list[RankedEntry]:
    """按分数降序排名，同分并列（1、2、2、4），并标记前 n 名为获奖者。"""
    ordered = sorted(players, key=lambda p: (-p.score, p.name))
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
            RankedEntry(
                rank=rank,
                player=player,
                award=award_for_rank(rank, winners, use_tiers),
                is_winner=index <= winners,
            )
        )
    return entries


def winners_of(entries: list[RankedEntry], winners: int) -> list[RankedEntry]:
    """取前 n 名（按排序位置，而不是名次号，避免并列导致超额）。"""
    return [e for e in entries if e.is_winner][:winners]


def safe_filename(text: str, fallback: str = "certificate") -> str:
    """把姓名等文本变成安全的文件名。"""
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(text)).strip().strip(".")
    return cleaned or fallback
