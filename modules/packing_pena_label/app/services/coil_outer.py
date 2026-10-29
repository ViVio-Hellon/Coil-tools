# -*- coding: utf-8 -*-
"""コイル外径から、近いチップボールを選ぶ(統合 1.2.35・現場の指示)

資材計算でできるようになった外径の出し方(統合 1.2.29)を、ペナラベルでも使う。

    コイル外径 = BOX実績_ｺｲﾙ内径 + 2 × BOX実績_ｺｲﾙ肉厚

- 引く鍵は**検査番号**(現場の確認: 検査番号 = 仕掛ファイルのロット番号)
- 読むファイルは**資材計算の設定(仕掛パス)と同じ**: 第1パス(梱包課共有のフォルダ・名前)→ 第2パス。
  名前が LS4LOT のままなら読まない(内径・肉厚が無い形式なので)。決まりは
  `common/ls4_source.py`、読み方は資材計算の `coil_tool/coil_results.py` をそのまま使う
  (いまの年月の行だけ・同じロットは作業開始がいちばん新しい行・共有のファイルは手元へ写してから)
- 第1パスの行に内径・肉厚が無ければ、第2パスの同じロットの行も探す(統合 1.2.45・現場の指摘)
- 読んだ内容はファイルの更新日時と大きさで覚えておき、変わったときだけ読み直す

近いチップボール: チップボールごとに風袋計算が使う外径(TIP1000 = 1100mm・TIP950 = 1000mm・
TIP820 = 900mm。`models.TIP_DIAMETER`)のうち、コイル外径にいちばん近いもの(同じ近さなら大きいほう)。
選んだら、どの外径からどれを選んだかを知らせる(違うときは選び直してもらう)。
"""
from __future__ import annotations

import logging
import os
import threading
import time
from typing import Dict, Optional, Tuple

from ..models import TIP_DIAMETER

log = logging.getLogger(__name__)

_lock = threading.Lock()
#: (ファイル, 更新日時, 大きさ) → {ロット: 行}
_cache: Dict[Tuple[str, float, int], Dict[str, dict]] = {}
#: 仕掛ファイルを探した結果(届かない共有を重量反映のたびに探して待たせない)
SOURCES_TTL_SEC = 60.0
_sources: Tuple[float, list] = (0.0, [])


def _material():
    """資材計算の部品(仕掛パスの設定と、実績の形式の読み方)。無ければ None(統合版でないとき)。"""
    try:
        from modules.packing_material_calculation.coil_tool import coil_results
        from common import ls4_source
    except Exception:                                   # noqa: BLE001 - 単体のペナラベルでは無い
        return None
    return coil_results, ls4_source


def enabled() -> bool:
    """外径を引く設定になっているか(資材計算の仕掛パスの名前が LS4LOT 以外)。"""
    parts = _material()
    if parts is None:
        return False
    try:
        return bool(parts[0].wanted())
    except Exception:                                   # noqa: BLE001
        log.warning("仕掛パスの設定を読めません", exc_info=True)
        return False


def _rows_of(path) -> Dict[str, dict]:
    coil_results, ls4_source = _material()
    st = os.stat(str(path))
    key = (str(path), st.st_mtime, st.st_size)
    with _lock:
        hit = _cache.get(key)
    if hit is not None:
        return hit
    rows, _total = coil_results.read_rows(path)
    by_lot = {str(r.get(ls4_source.COL_LOT) or "").strip(): r for r in rows}
    with _lock:
        for old in [k for k in _cache if k[0] == key[0]]:
            del _cache[old]
        _cache[key] = by_lot
    return by_lot


def _find_sources(coil_results) -> list:
    global _sources
    with _lock:
        at, files = _sources
    if at and time.monotonic() - at < SOURCES_TTL_SEC:
        return files
    files = list(coil_results.sources())
    with _lock:
        _sources = (time.monotonic(), files)
    return files


def _details():
    """梱包明細のロット引き(`meisai/lot_lookup.py`)。無ければ None(統合版でないとき)。"""
    try:
        from modules.packing_details.meisai import lot_lookup
    except Exception:                                   # noqa: BLE001 - 単体のペナラベルでは無い
        return None
    return lot_lookup


def lookup_lot(lot: str) -> dict:
    """LOTNO から、梱包明細が読むファイルの検査番号(仕掛ロットなど)とコイル外径(LS4LOT)(統合 1.2.57・1.2.58)。

    梱包明細の部品が無ければ ``{"found": False, "off": True}``。
    """
    parts = _details()
    if parts is None:
        return {"found": False, "off": True, "notes": [], "kensa": "", "outer": None}
    try:
        return parts.lookup(lot)
    except Exception as exc:                            # noqa: BLE001 - 引けなくても画面は止めない
        log.warning("LOTNO を引けませんでした", exc_info=True)
        return {"found": False, "reason": f"仕掛のファイルを読めません: {exc}", "notes": [],
                "kensa": "", "outer": None}


def lookup(lot: str) -> dict:
    """そのロットのコイル外径。

    先に梱包明細が読む仕掛(LS4LOT。**名前に関わらず列で読む**。統合 1.2.57)を見て、外径が出せなければ
    資材計算の設定(実績の形式のファイル)を見る。
    戻り値: ``{"found": bool, "outer": float|None, "inner": .., "thick": .., "file": ..,
    "reason": 見つからない・読めないときの理由}``。どちらの部品も無ければ ``{"found": False, "off": True}``。

    以前(1.2.35〜1.2.56)は資材計算の仕掛パスの名前が LS4LOT 以外のときだけ読み、LS4LOT のままなら
    **何も言わずに**選ばなかった(現場の指摘「外径計算してチップボールを選んでいない」)。
    """
    lot = str(lot or "").strip()
    if not lot:
        return {"found": False, "off": True}
    first = lookup_lot(lot)
    if first.get("outer") is not None:
        return {"found": True, "outer": first["outer"], "inner": first.get("inner"),
                "thick": first.get("thick"), "file": first.get("file", "")}
    if not enabled():
        if first.get("off"):
            return {"found": False, "off": True}
        reason = first.get("reason") or "・".join(first.get("notes") or []) or "外径を出せません"
        return {"found": False, "reason": reason}
    coil_results, ls4_source = _material()
    reasons = []
    try:
        files = _find_sources(coil_results)
    except Exception as exc:                            # noqa: BLE001
        return {"found": False, "reason": f"仕掛ファイルを探せません: {exc}"}
    if not files:
        return {"found": False, "reason": "仕掛ファイル(実績の形式)が見つかりません"}
    empty_in = ""
    for path in files:
        try:
            row = _rows_of(path).get(lot)
        except Exception as exc:                        # noqa: BLE001 - 読めないファイルは次へ
            log.warning("仕掛ファイルを読めません: %s", path, exc_info=True)
            reasons.append(f"{os.path.basename(str(path))} を読めません({exc})")
            continue
        if row is None:
            continue
        outer = ls4_source.outer_diameter(row)
        if outer is None:
            # 第1パスの行に内径・肉厚が無ければ、第2パスの同じロットも探す(統合 1.2.45・現場の指摘)
            empty_in = empty_in or str(path)
            continue
        return {"found": True, "outer": outer,
                "inner": row.get(ls4_source.COL_INNER), "thick": row.get(ls4_source.COL_THICK),
                "file": str(path)}
    if empty_in:
        return {"found": False, "file": empty_in,
                "reason": f"ロット {lot} の内径・肉厚が入っていません"
                          + ("(第2パスにもありません)" if len(files) > 1 else "")}
    return {"found": False,
            "reason": "・".join(reasons) or f"仕掛ファイルの今月の行にロット {lot} がありません"}


def nearest_tip(outer: float) -> str:
    """コイル外径にいちばん近いチップボール(同じ近さなら大きいほう)。"""
    return min(TIP_DIAMETER, key=lambda t: (abs(TIP_DIAMETER[t][0] - outer), -TIP_DIAMETER[t][0]))


def tip_label(tip: str) -> str:
    return tip.replace("TIP", "チップボール ")


def choose(lot: str, current_tip: str) -> Tuple[Optional[str], str]:
    """ロット番号から近いチップボールを選ぶ(LOTNO を入れていなければ検査番号をロット番号として)。

    戻り値: (選んだチップボール(変えないなら None), 知らせる文(出さないなら ""))。
    """
    hit = lookup(lot)
    if hit.get("off"):
        return None, ""
    now = tip_label(current_tip) if current_tip in TIP_DIAMETER else "未選択"
    if not hit.get("found"):
        return None, (f"チップボールは自動で選んでいません({hit.get('reason')})。"
                      f"いまの選択: {now}")
    outer = hit["outer"]
    tip = nearest_tip(outer)
    basis = (f"コイル外径 {outer:g}mm(内径 {hit.get('inner')} + 2 × 肉厚 {hit.get('thick')})"
             f" → いちばん近いのは {tip_label(tip)}(外径 {TIP_DIAMETER[tip][0]}mm として計算)")
    if tip == current_tip:
        return None, f"{basis}。選んであるとおりです"
    return tip, f"{basis}。{tip_label(tip)}を選びました(前: {now})。違うときは選び直してください"


def forget() -> None:
    """覚えた内容を捨てる(試験用)。"""
    global _sources
    with _lock:
        _cache.clear()
        _sources = (0.0, [])
