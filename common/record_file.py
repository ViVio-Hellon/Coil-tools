# -*- coding: utf-8 -*-
"""作業の記録を、月ごとのファイルに 1 件ずつ足す(ペナラベル・梱包明細で共通。統合 1.2.44)

現場の指摘: .jsonl は普通は開けない。CSV で出すか、選べるようにしてほしい。

形式(設定「記録ファイルの形式」。PC ごと):

- ``csv``(既定): **Excel でそのまま開ける CSV**。UTF-8(BOM つき)・行末 CRLF。
  列は ``CSV_COLUMNS``。決まった列(日時・種類・内容・モード・PC名・ライン・利用者・版)のあとに、
  記録の全部を JSON で入れた「詳しい中身」の列を置く(CSV にしても何も落とさない)
- ``jsonl``: 1 行 1 件の JSON(機械で読むとき)。1.2.43 までの形
- ``both``: 両方

**Excel で開いている間も記録を落とさない。** Excel は開いた CSV を書けないように押さえるので、
書けないときは同じフォルダーの「<名前>.csv.追記待ち」に足しておき、次に書けたときに本体へまとめて
書き足して消す。どちらにも書けないときだけ、書けなかった理由を返す(計算・印刷は止めない)。

ファイルは足すだけ(消さない・書き換えない)。見るときは写しを開くか、読み取り専用で開く
(Excel で上書き保存すると形式が変わることがある)。
"""
from __future__ import annotations

import csv
import datetime
import io
import json
import os
from typing import Any, Dict, Iterable, List, Mapping, Optional, Tuple

FORMAT_CSV = "csv"
FORMAT_JSONL = "jsonl"
FORMAT_BOTH = "both"
FORMATS: Tuple[str, ...] = (FORMAT_CSV, FORMAT_JSONL, FORMAT_BOTH)
DEFAULT_FORMAT = FORMAT_CSV
#: 設定画面の選択肢(値, 見出し)
FORMAT_CHOICES: Tuple[Tuple[str, str], ...] = (
    (FORMAT_CSV, "CSV(Excel で開ける)"),
    (FORMAT_JSONL, "JSONL(1 行 1 件の JSON)"),
    (FORMAT_BOTH, "CSV と JSONL の両方"),
)

CSV_COLUMNS: Tuple[str, ...] = (
    "日時", "種類", "内容", "モード", "PC名", "ライン", "利用者",
    "統合の版", "機能の版", "中身の指紋", "詳しい中身",
)
PENDING_SUFFIX = ".追記待ち"
BOM = "﻿"

#: 「内容」に出すときの項目名(無いものはそのままの名前で出す)
DEFAULT_LABELS: Dict[str, str] = {
    "title": "画面", "kensaNo": "検査番号", "lot": "ロット", "no": "No", "nos": "No",
    "path": "場所", "query": "条件", "width": "巾", "tip": "チップボール",
    "material": "資材重量", "label": "取得元", "asOf": "時点", "stale": "古い内容",
    "file": "ファイル", "what": "項目", "value": "値", "reason": "理由",
    "changes": "変わったもの", "packages": "梱包", "notOnLabel": "ラベルに載らない",
}
#: 決まった列に入れるので「内容」には出さないもの
_HEAD_KEYS = ("at", "kind", "version", "who")


def normalize_format(value: Any) -> str:
    """設定の値を形式へ。知らない値・空は既定(CSV)。"""
    text = str(value or "").strip().lower()
    return text if text in FORMATS else DEFAULT_FORMAT


def format_label(value: Any) -> str:
    fmt = normalize_format(value)
    return dict(FORMAT_CHOICES)[fmt]


def file_names(prefix: str, when: datetime.datetime, fmt: Any) -> List[str]:
    """その月に書くファイルの名前(形式の順)。"""
    fmt = normalize_format(fmt)
    month = when.strftime("%Y-%m")
    out = []
    if fmt in (FORMAT_CSV, FORMAT_BOTH):
        out.append(f"{prefix}_{month}.csv")
    if fmt in (FORMAT_JSONL, FORMAT_BOTH):
        out.append(f"{prefix}_{month}.jsonl")
    return out


# ------------------------------------------------------------------ 1 件 → 1 行
def _text(value: Any, labels: Mapping[str, str]) -> str:
    if isinstance(value, bool):
        return "はい" if value else "いいえ"
    if value is None:
        return ""
    if isinstance(value, Mapping):
        parts = [f"{labels.get(str(k), str(k))}={_text(v, labels)}"
                 for k, v in value.items() if v not in (None, "", [], {})]
        return "(" + "・".join(parts) + ")" if parts else ""
    if isinstance(value, (list, tuple)):
        return "・".join(_text(v, labels) for v in value)
    return str(value)


def summary(entry: Mapping[str, Any], labels: Optional[Mapping[str, str]] = None) -> str:
    """「内容」の列: 決まった列に入れたもの以外を「項目: 値」で並べる。"""
    names = dict(DEFAULT_LABELS)
    names.update(labels or {})
    parts = []
    for key, value in entry.items():
        if key in _HEAD_KEYS or value in (None, "", [], {}):
            continue
        parts.append(f"{names.get(key, key)}: {_text(value, names)}")
    return " / ".join(parts)


def csv_row(entry: Mapping[str, Any], feature_version_key: str,
            labels: Optional[Mapping[str, str]] = None) -> List[str]:
    """1 件を CSV の 1 行に(列は ``CSV_COLUMNS``)。"""
    who = entry.get("who") or {}
    ver = entry.get("version") or {}
    return [
        str(entry.get("at", "")), str(entry.get("kind", "")), summary(entry, labels),
        str(who.get("mode", "")), str(who.get("pc", "")), str(who.get("line", "")),
        str(who.get("user", "")),
        str(ver.get("app", "")), str(ver.get(feature_version_key, "")), str(ver.get("stamp", "")),
        json.dumps(entry, ensure_ascii=False),
    ]


def _csv_text(rows: Iterable[Iterable[str]]) -> str:
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n")
    for row in rows:
        w.writerow(row)
    return buf.getvalue()


# ------------------------------------------------------------------ 書く
def _append_text(path: str, text: str, *, header: str = "") -> None:
    """足す。新しいファイル(空)なら先に header を書く。"""
    new = not os.path.exists(path) or os.path.getsize(path) == 0
    with open(path, "a", encoding="utf-8", newline="") as f:
        if new and header:
            f.write(header)
        f.write(text)


def _write_one(path: str, text: str, header: str) -> str:
    """本体へ足す(前に書けなかった「追記待ち」があれば先に書き足して消す)。

    本体が押さえられて書けなければ「追記待ち」へ足す。戻り値は知らせる文(書けたら空)。
    """
    pending = path + PENDING_SUFFIX
    try:
        waiting = ""
        if os.path.exists(pending):
            with open(pending, encoding="utf-8", newline="") as f:
                waiting = f.read()
        _append_text(path, waiting + text, header=header)
        if waiting:
            try:
                os.remove(pending)
            except OSError:
                # 消せなくても次に 2 重に書かないよう、空にしておく
                open(pending, "w", encoding="utf-8").close()
        return ""
    except PermissionError as exc:
        try:
            _append_text(pending, text)
        except OSError as exc2:
            return f"{path} に書けません({exc2.strerror or exc2})"
        return (f"{os.path.basename(path)} が開かれていて書けません(Excel などで開いているときは閉じてください)。"
                f"いまは {os.path.basename(pending)} に足しておき、閉じたあと次の記録で書き足します"
                f"({exc.strerror or exc})")


def append(folder: str, prefix: str, entry: Mapping[str, Any], fmt: Any, *,
           feature_version_key: str, labels: Optional[Mapping[str, str]] = None,
           when: Optional[datetime.datetime] = None) -> str:
    """1 件をそのフォルダーの月ごとのファイルへ足す。

    戻り値: 知らせる文(すべて本体に書けたら空)。フォルダーを作れないなどは OSError をそのまま上げる。
    """
    when = when or datetime.datetime.now()
    os.makedirs(folder, exist_ok=True)
    problems = []
    for name in file_names(prefix, when, fmt):
        path = os.path.join(folder, name)
        if name.endswith(".csv"):
            text = _csv_text([csv_row(entry, feature_version_key, labels)])
            header = BOM + _csv_text([CSV_COLUMNS])
        else:
            text = json.dumps(entry, ensure_ascii=False) + "\n"
            header = ""
        msg = _write_one(path, text, header)
        if msg:
            problems.append(msg)
    return "。".join(problems)


def _entries_from_csv_text(text: str, has_header: bool) -> List[Dict[str, Any]]:
    out = []
    rows = csv.reader(io.StringIO(text.lstrip(BOM)))
    if has_header:
        next(rows, None)
    detail = CSV_COLUMNS.index("詳しい中身")
    for row in rows:
        if len(row) <= detail:
            continue                                    # 書きかけ・壊れた行は飛ばす
        try:
            item = json.loads(row[detail])
        except ValueError:
            continue
        if isinstance(item, dict):
            out.append(item)
    return out


def _entries_from_jsonl_text(text: str) -> List[Dict[str, Any]]:
    out = []
    for line in text.splitlines():
        if not line.strip():
            continue
        try:
            item = json.loads(line)
        except ValueError:
            continue                                    # 書きかけの行は飛ばす
        if isinstance(item, dict):
            out.append(item)
    return out


def read_month(folder: str, prefix: str, month: datetime.date) -> List[Dict[str, Any]]:
    """その月の記録を全部(CSV・JSONL・追記待ちのどれからでも)。古い順。

    形式を月の途中で変えたとき・両方のときも 1 件は 1 回だけ(CSV と JSONL の両方にある同じ中身は 1 つにまとめる)。
    読めないファイルは OSError をそのまま上げる(見つからないものは飛ばす)。
    """
    base = os.path.join(folder, f"{prefix}_{month:%Y-%m}")
    families: List[List[Dict[str, Any]]] = []
    for parts in (((base + ".csv", "csv", True), (base + ".csv" + PENDING_SUFFIX, "csv", False)),
                  ((base + ".jsonl", "jsonl", False), (base + ".jsonl" + PENDING_SUFFIX, "jsonl", False))):
        items: List[Dict[str, Any]] = []
        for path, kind, header in parts:
            try:
                with open(path, encoding="utf-8", newline="") as f:
                    text = f.read()
            except FileNotFoundError:
                continue
            items += (_entries_from_csv_text(text, header) if kind == "csv"
                      else _entries_from_jsonl_text(text))
        families.append(items)
    # CSV と JSONL の両方に書いた分は 1 回だけ(同じファイルの中の同じ中身は別の記録なので残す)
    out = list(families[0])
    left: Dict[str, int] = {}
    for item in families[0]:
        key = json.dumps(item, ensure_ascii=False, sort_keys=True)
        left[key] = left.get(key, 0) + 1
    for item in families[1]:
        key = json.dumps(item, ensure_ascii=False, sort_keys=True)
        if left.get(key, 0) > 0:
            left[key] -= 1
            continue
        out.append(item)
    out.sort(key=lambda x: str(x.get("at", "")))        # 同じ時刻は読んだ順のまま
    return out


def read_csv(path: str) -> List[Dict[str, str]]:
    """書いた CSV を読む(試験・確かめ用)。"""
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))
