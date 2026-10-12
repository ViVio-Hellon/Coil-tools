# -*- coding: utf-8 -*-
"""マスタの CSV を読む（資材マスタ CSV・型番マスタ CSV）。

UTF-8（BOM あり・なし）で読み、読めなければ Shift_JIS（CP932）で読み直す。
「手本を出す」で出した CSV を Excel で直して「CSV（コンマ区切り）」で保存すると
Shift_JIS になるため（統合 1.2.8）。どちらでも読めないときは UnicodeDecodeError。
"""
from __future__ import annotations

import csv
import io
from typing import List

ENCODINGS = ("utf-8-sig", "cp932")


def read_text(path: str) -> str:
    """CSV の中身を文字で返す。"""
    with open(path, "rb") as f:
        raw = f.read()
    last: UnicodeDecodeError = None
    for enc in ENCODINGS:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError as exc:
            last = exc
    raise last


def read_rows(path: str) -> List[List[str]]:
    """CSV の行（空行は除く）。"""
    reader = csv.reader(io.StringIO(read_text(path), newline=""))
    return [r for r in reader if any((c or "").strip() for c in r)]


def read_dicts(path: str) -> List[dict]:
    """1 行目を見出しにした CSV の行。"""
    return list(csv.DictReader(io.StringIO(read_text(path), newline="")))
