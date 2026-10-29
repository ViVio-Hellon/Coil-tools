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
import time
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
    """足す。ファイルが空なら先に header を書く。

    空かどうかは**開いたあとの位置**で見る(統合 1.2.61)。以前は `os.path.exists` で見ていて、共有で一瞬
    確かめられないと「無い」とみなし、ファイルの途中に見出しをもう一度書くことがあった。
    """
    with open(path, "a", encoding="utf-8", newline="") as f:
        if header and f.tell() == 0:
            f.write(header)
        f.write(text)


#: 追記待ちを書き足すあいだの名前(「<名前>.追記待ち.<番号>.処理中」)。この名前のファイルは二度と書き足さない
CLAIM_SUFFIX = ".処理中"


def _read_text(path: str) -> str:
    with open(path, encoding="utf-8", newline="") as f:
        return f.read()


def _claim_pending(pending: str) -> Tuple[Optional[str], str]:
    """追記待ちを**名前を変えて引き取る**。戻り値は (引き取ったファイル, 中身)。無い・引き取れないなら (None, "")。

    統合 1.2.61。以前は追記待ちを読んで本体へ書き足してから消していた。消せない(ウイルス対策・バックアップ・
    Excel で開いている・消す権限が無い)と追記待ちが残り、**記録のたびに同じ行が本体へ何度も書き足された**
    (別のツールで、読めなかったのを取り違えて送り直し、重複を大量に作った事案と同じ形)。
    いまは書き足す前に「…処理中」へ名前を変える。本体へ書けたら、その名前のファイルは二度と書き足さない
    (消せなくても重ならない)。本体へ書けなければ追記待ちへ戻す。
    """
    if not os.path.exists(pending):
        return None, ""
    claimed = f"{pending}.{os.getpid()}.{time.time_ns()}{CLAIM_SUFFIX}"
    try:
        os.replace(pending, claimed)
    except OSError:
        return None, ""                 # 引き取れなければ今回は混ぜない(次に書けたときに書き足す)
    try:
        return claimed, _read_text(claimed)
    except OSError:
        _give_back(claimed, pending)
        return None, ""


def _give_back(claimed: str, pending: str) -> bool:
    """引き取った追記待ちを戻す(本体へ書けなかったとき)。戻せたら True。"""
    try:
        if not os.path.exists(pending):
            os.replace(claimed, pending)
            return True
        # 引き取ったあとに新しい追記待ちができていた。後ろへ足してから消す(古い順を保つため前へ)
        old, new = _read_text(claimed), _read_text(pending)
        tmp = pending + CLAIM_SUFFIX + ".戻し"
        with open(tmp, "w", encoding="utf-8", newline="") as f:
            f.write(old + new)
        os.replace(tmp, pending)
        os.remove(claimed)
        return True
    except OSError:
        return False


def _write_one(path: str, text: str, header: str) -> str:
    """本体へ足す(前に書けなかった「追記待ち」があれば先に書き足す)。

    本体に書けなければ「追記待ち」へ足す。戻り値は知らせる文(書けたら空)。
    **本体に書けたあとは、追記待ちへ足さない・「書けません」と言わない**(統合 1.2.61。以前は追記待ちを消せない
    ことを「本体が開かれている」と取り違え、今回の行を追記待ちにも足して、次に 2 重に書いていた)。
    """
    pending = path + PENDING_SUFFIX
    claimed, waiting = _claim_pending(pending)
    try:
        _append_text(path, waiting + text, header=header)
    except OSError as exc:
        kept = "" if claimed is None or _give_back(claimed, pending) else (
            f"。前に書けなかった分は {os.path.basename(claimed)} に残っています")
        try:
            _append_text(pending, text)
        except OSError as exc2:
            return f"{path} に書けません({exc2.strerror or exc2}){kept}"
        if isinstance(exc, PermissionError):
            head = (f"{os.path.basename(path)} が開かれていて書けません(Excel などで開いているときは閉じてください)。"
                    f"いまは {os.path.basename(pending)} に足しておき、閉じたあと次の記録で書き足します")
        else:
            head = (f"{os.path.basename(path)} に書けません。いまは {os.path.basename(pending)} に足しておき、"
                    "次の記録で書き足します")
        return f"{head}({getattr(exc, 'strerror', None) or exc}){kept}"
    if claimed is not None:
        try:
            os.remove(claimed)
        except OSError:
            pass                        # 名前が「…処理中」なので、残っても二度と書き足さない
    return ""


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
