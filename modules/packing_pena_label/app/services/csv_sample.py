# -*- coding: utf-8 -*-
"""設定の「手本を出す」（資材マスタ CSV・型番マスタ CSV。統合 1.2.8）。

欄に入れたパスへ、列の見出しと例の行（同梱の CSV と同じ中身）を書く。
現場はそれを Excel で開いて直し、そのパスのまま使う。

- 出す先は**欄に入れたパス**（保存していなくてよい）。空なら断る
- フォルダーを入れたときは、その中に同梱と同じ名前（資材重量.csv / 型番.csv）で出す
- **すでにあるファイルは上書きしない**（使っているマスタを手本で消さない）
- 出す先のフォルダーが無ければ断る（打ち間違いで知らない場所にフォルダーを作らない）
- 文字は UTF-8（BOM 付き。Excel でそのまま開ける）・改行は CRLF。
  Excel で「CSV（コンマ区切り）」に保存し直した Shift_JIS でも読める（csv_text）
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Callable, Dict

from ..config import Config
from ..repositories import csv_text


@dataclass(frozen=True)
class Sample:
    label: str
    #: 同梱の CSV（手本の中身）
    bundled: Callable[[Config], str]
    #: 読む順の説明（画面に出す）
    order: str


SAMPLES: Dict[str, Sample] = {
    "material_csv_file": Sample(
        "資材マスタ CSV", lambda cfg: cfg.default_material_csv_path,
        "梱包資材マスタ（SQLite）→ 旧 Access（「旧 Access も試す」がオンのとき）→ この CSV"),
    "kataban_csv_file": Sample(
        "型番マスタ CSV", lambda cfg: cfg.default_kataban_csv_path,
        "梱包資材マスタ（SQLite）の表「型番」→ この CSV"),
}


def _clean(raw) -> str:
    """欄の値。前後の空白と、Windows の「パスのコピー」で付く " を取る。"""
    v = str(raw or "").strip()
    if len(v) >= 2 and v[0] == v[-1] == '"':
        v = v[1:-1].strip()
    return v


def write_sample(cfg: Config, key: str, raw_path) -> dict:
    sample = SAMPLES.get(key)
    if sample is None:
        return {"ok": False, "level": "warn", "message": "手本を出せる欄ではありません"}
    path = _clean(raw_path)
    if not path:
        return {"ok": False, "level": "warn",
                "message": "%s の欄に、手本を出す先のパスを入れてください" % sample.label}
    bundled = sample.bundled(cfg)
    if os.path.isdir(path):
        path = os.path.join(path, os.path.basename(bundled))
    folder = os.path.dirname(os.path.abspath(path))
    if not os.path.isdir(folder):
        return {"ok": False, "level": "warn",
                "message": "出す先のフォルダーがありません: %s" % folder}
    if os.path.exists(path):
        return {"ok": False, "level": "warn", "path": path,
                "message": ("すでにファイルがあります。上書きしません: %s\n"
                            "（手本が要るときは、別の名前を入れてから押してください）") % path}
    try:
        text = csv_text.read_text(bundled)
        lines = [ln for ln in text.splitlines()]
        with open(path, "x", encoding="utf-8-sig", newline="") as f:
            f.write("\r\n".join(lines) + "\r\n")
    except FileExistsError:
        return {"ok": False, "level": "warn", "path": path,
                "message": "すでにファイルがあります。上書きしません: %s" % path}
    except OSError as exc:
        return {"ok": False, "level": "error", "path": path,
                "message": "手本を書けませんでした: %s（%s）" % (path, exc)}
    rows = max(len([ln for ln in lines if ln.strip()]) - 1, 0)
    return {"ok": True, "level": "info", "path": path, "rows": rows,
            "message": ("%s の手本を出しました（見出しと例 %d 行）: %s\n"
                        "Excel で開いて直してください。読む順: %s。"
                        "この欄が保存されていなければ「この欄を保存」も押してください。")
                       % (sample.label, rows, path, sample.order)}
