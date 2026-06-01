"""Schedule auto-balance rules and scoring (no UI dependencies)."""

import datetime
import sys

import numpy as np
import pandas as pd

SHIFT_INFO = {
    'Trike': {'color': '#9999FF', 'isHour': False},
    'Gallery': {'color': '#5eb91e', 'isHour': False},
    'Back': {'color': '#E1EC24', 'isHour': False},
    'Front': {'color': '#9BC2E6', 'isHour': False},
    'Float': {'color': '#ffffff', 'isHour': False},
    'ENCA': {'color': '#FFD966', 'isHour': False},
    'Head': {'color': '#ffd221', 'isHour': False},
    'MOT': {'color': '#ffd221', 'isHour': False},
    'DIAR': {'color': '#b50934', 'isHour': False},
    'Lunch': {'color': '#7d7f7c', 'isHour': False},
    'Dinner': {'color': '#7d7f7c', 'isHour': False},
    'Security': {'color': '#00b1d2', 'isHour': True},
    'Tickets': {'color': '#f6f9d4', 'isHour': True},
    'FLOAT': {'color': '#D3D3D3', 'isHour': False},
    'Badges': {'color': '#f6f9d4', 'isHour': True},
    'Project': {'color': '#f83dda', 'isHour': False},
    'GRGR': {'color': '#cf6498', 'isHour': True},
    'Manager': {'color': '#ACB9CA', 'isHour': True},
    'Securager': {'color': '#00b1d2', 'isHour': True},
    'STST': {'color': '#bf94e4', 'isHour': False},
    'MP': {'color': '#FFA500', 'isHour': False},
    'Museum Project': {'color': '#FFA500', 'isHour': True},
    'Training': {'color': '#FFD580', 'isHour': True},
    'Camp': {'color': '#c7ea46', 'isHour': True},
    'Retail': {'color': '#ffccd4', 'isHour': True},
    'Float 0': {'color': '#BED785', 'isHour': False},
    'Float 1': {'color': '#90E4C1', 'isHour': False},
    'CORO': {'color': '#f4b183', 'isHour': False},
    'Pizza': {'color': '#f73939', 'isHour': True},
    'Zoom': {'color': '#ffffff', 'isHour': False},
    '': {'color': '#D3D3D3', 'isHour': False},
    None: {'color': '#D3D3D3', 'isHour': False},
}

STANDARD_FLOOR_SHIFTS = [
    'Security', 'Trike', 'CORO', 'Gallery', 'Front', 'Back', 'Float 0', 'Float 1', 'ENCA',
]
NONSTANDARD_SHIFTS = [
    'Manager', 'GRGR', 'Project', 'MP', 'Pizza', 'Retail', 'Zoom'
]
SWAPPABLE_FLOOR_SHIFTS = [
    'Trike', 'CORO', 'Gallery', 'Front', 'Back', 'Float 0', 'Float 1', 'ENCA',
]

if sys.platform == 'win32':
    RES_FILE_NAME = f'momath_schedule_{datetime.date.today()}.xlsx'
else:
    RES_FILE_NAME = f'/tmp/momath_schedule_{datetime.date.today()}.xlsx'

BEFORE_1PM_CUTOFF = '01:00'


class ShiftBalanceRule:
    """One schedulable constraint for auto_balance_shifts."""

    name = 'base'
    description = ''

    def __init__(self, enabled=True):
        self.enabled = enabled

    def count(self, col_series):
        if not self.enabled:
            return 0
        return self._count(col_series)

    def _count(self, col_series):
        raise NotImplementedError


class NoDuplicateConsecutiveRule(ShiftBalanceRule):
    name = 'no_duplicate_consecutive'
    description = 'No duplicate consecutive shifts.'

    def _count(self, col_series):
        values = col_series.tolist()
        violations = 0
        for i in range(len(values) - 1):
            val, next_val = values[i], values[i + 1]
            if val in SWAPPABLE_FLOOR_SHIFTS and val == next_val:
                violations += 1
        return violations


class NoTrikeCoroAdjacencyRule(ShiftBalanceRule):
    name = 'no_trike_coro_adjacency'
    description = 'No consecutive Trike/CORO.'

    def _count(self, col_series):
        values = col_series.tolist()
        violations = 0
        for i in range(len(values) - 1):
            val, next_val = values[i], values[i + 1]
            if {val, next_val} == {'Trike', 'CORO'}:
                violations += 1
        return violations


class TrikeCoroBefore1pmRule(ShiftBalanceRule):
    name = 'trike_coro_before_1pm'
    description = 'Balance Trike/CORO before 1 PM.'

    def _count(self, col_series):
        index = col_series.index.tolist()
        try:
            cutoff_pos = index.index(BEFORE_1PM_CUTOFF)
        except ValueError:
            cutoff_pos = len(index)

        trike_coro_before_1pm = sum(
            1 for i, val in enumerate(col_series.tolist())
            if i < cutoff_pos and val in ('Trike', 'CORO')
        )
        if trike_coro_before_1pm > 1:
            return trike_coro_before_1pm - 1
        return 0


class NoTrikeAdjacentLunchRule(ShiftBalanceRule):
    name = 'no_trike_adjacent_lunch'
    description = 'Trike not adjacent to Lunch.'

    def _count(self, col_series):
        values = col_series.tolist()
        violations = 0
        for i in range(len(values) - 1):
            if {values[i], values[i + 1]} == {'Trike', 'Lunch'}:
                violations += 1
        return violations


def default_balance_rules():
    return [
        NoDuplicateConsecutiveRule(),
        NoTrikeCoroAdjacencyRule(),
        TrikeCoroBefore1pmRule(),
        NoTrikeAdjacentLunchRule(),
    ]


def set_balance_rule_enabled(rules, name, enabled):
    for rule in rules:
        if rule.name == name:
            rule.enabled = enabled
            return
    raise ValueError(f'Unknown balance rule: {name!r}')


def count_column_violations(col_series, rules):
    return tuple(rule.count(col_series) for rule in rules)


def total_violations(df, rules):
    totals = [0] * len(rules)
    for col in df.columns:
        for bucket, count in enumerate(count_column_violations(df[col], rules)):
            totals[bucket] += count
    return tuple(totals)


def introduces_no_new_violations(df, row_label, col_name, new_value, rules):
    current_score = count_column_violations(df[col_name], rules)
    df_copy = df.copy()
    df_copy.at[row_label, col_name] = new_value
    new_score = count_column_violations(df_copy[col_name], rules)
    return new_score <= current_score


def format_balance_rule_line(index, rule):
    flag = '[on]' if rule.enabled else '[off]'
    return f'{flag}  {index + 1}.  {rule.description}'


def shift_background_color(shift) -> str | None:
    """Return QColor-ready hex for a shift, or None for empty/default cells."""
    if pd.isna(shift) or shift == '':
        return None
    info = SHIFT_INFO.get(shift)
    if not info:
        return '#FFFFFF'
    color = info['color']
    if not color.startswith('#'):
        color = f'#{color}'
    return color
