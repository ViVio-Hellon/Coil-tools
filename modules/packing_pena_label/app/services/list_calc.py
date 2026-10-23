# -*- coding: utf-8 -*-
"""羅列計算（1〜50本の一覧表）。VBA の ``羅列計算`` に相当する画面。

== 2026.09 方針変更 ==

VBA の ``羅列計算`` は ``DB_Materials`` とは **別の計算をしていた**。

    1. 風袋に EX-DRY を含めない（8 項目合計。DB_Materials は 9 項目）
    2. 梱包高さに ハードボード 2.35mm×2 を加えない
    3. HB を ``CalcHBForList``（HBFlag を使わず巾と本数だけで判定）で出す

これは「片側だけ直して、もう片側が残った」ものと確認できたため、
**一覧表は参考値であり ``CalculateMaterials`` を正とする** という判断に従い、
本モジュールは自前の計算をやめ、**本数 1〜50 について
``tare_calc.calculate_materials`` を呼ぶだけ**に変えた。

これにより
    ・風袋 / GW / 梱包高さ / HB が 風袋計算画面と必ず一致する
    ・資材の計算式が 1 箇所（tare_calc）だけになる
    ・EX-DRY が一覧にも現れる（列を 1 つ追加した）
"""

from __future__ import annotations

import logging
from typing import Dict, List

from ..repositories import material_repo as M
from . import tare_calc as TC
from .vba_compat import cstr, fmt, single, val, ws_roundup

log = logging.getLogger(__name__)

MAX_COILS = 50

#: 一覧の見出し（VBA の `50まで` シート 1 行目に相当。EX-DRY を追加）
HEADERS = [
    "本数",
    "ﾋﾟﾝ\n(kg)", "ﾁｯﾌﾟﾎﾞｰﾙ\n(kg)", "ｽﾄﾚｯﾁﾌｨﾙﾑ\n(kg)", "ｴｻﾌｫｰﾑ\n(kg)",
    "ﾎﾟﾘｼｰﾄ\n(kg)", "ﾊｰﾄﾞﾎﾞｰﾄﾞ\n(kg)", "樹脂ﾊﾟﾚｯﾄ\n(kg)", "ﾍﾟｯﾄﾊﾞﾝﾄﾞ\n(kg)",
    "EX-DRY\n(kg)",
    "NW\n(kg)", "風袋\n(kg)", "GW\n(kg)",
    "ｺｲﾙ高さ\n(mm)", "梱包高さ\n(mm)", "ﾁｯﾌﾟﾎﾞｰﾙ高さ\n(mm)",
]

#: 列キー（HEADERS と同順）
COLUMN_KEYS = ["count", "pin", "tip", "sut", "esa", "por", "hb", "par", "pet",
               "dry", "nw", "hu", "gw", "coilHeight", "packHeight", "tipHeight"]


def build_list(table: M.MaterialTable, width_mm: float, tip: str,
               t1: float, t2: float,
               take1_visible: bool, take2_visible: bool,
               cb_key: str = "", ob_idx: int = 0,
               ) -> Dict:
    """1〜50 本の一覧を作る。各行は ``calculate_materials`` の結果そのもの。

    NW の丈の選び方は VBA と同じ
    （丈1 が表示中なら丈1 の 1 条重量、丈2 だけなら丈2 の 1 条重量）。
    ``calculate_materials`` は coilNo 1,2 で T1・3,4 で T2 を使うため、
    丈1 が見えていれば 1、そうでなければ 3 を渡す。
    """
    coil_no = 1 if take1_visible else 3

    rows: List[dict] = []
    missing: List[str] = []

    for q in range(1, MAX_COILS + 1):
        mat = TC.calculate_materials(
            table, q, coil_no, width_mm, tip, cb_key, ob_idx,
            t1, t2, take1_visible, take2_visible,
            all_coil_empty=False,
            )

        for name in mat.missing:
            if name not in missing:
                missing.append(name)

        disp = TC.display_values(mat)
        rows.append({
            "count": str(q),
            "pin": disp["PI"],
            "tip": disp["TIP"],
            "sut": disp["SUT"],
            "esa": disp["ESA"],
            "por": disp["POR"],
            "hb": disp["HB"],
            "par": disp["PAR"],
            "pet": disp["PET"],
            "dry": disp["DRY"],
            "nw": disp["NW"],
            "hu": disp["HU"],
            "gw": disp["GW"],
            # 高さは VBA のセル書式 0.0_ と同じ小数1桁(風袋計算の画面とも同じ。統合 1.2.12。以前は "0" で 596.9 → 597)
            "coilHeight": fmt(val(mat.TA * 1000), "0.0"),
            # 風袋計算画面の高さ表示と同じ（樹脂パレット上下 + ハードボード上下、
            # 10mm 単位で切り上げ）
            "packHeight": fmt(ws_roundup(mat.KTA, -1), "0.0"),
            "tipHeight": fmt(val(mat.TIPT * 1000), "0.0"),
            "raw": {
                "pin": mat.PIN, "tip": mat.TIP, "sut": mat.SUT, "esa": mat.ESA,
                "por": mat.POR, "hb": mat.HB, "par": mat.PAR, "pet": mat.PET,
                "dry": mat.DRY, "nw": mat.NW, "hu": mat.HU, "gw": mat.GW,
                "ta": mat.TA, "kta": mat.KTA, "tipt": mat.TIPT,
            },
        })

    return {"headers": HEADERS, "columnKeys": COLUMN_KEYS,
            "rows": rows, "missing": missing}


def quick_height(coil_count: float, width_mm: float) -> Dict[str, str]:
    """積み高さ計算(参考。VBA ``CommandButton13``)。**風袋計算の本体と同じ式**で出す(統合 1.2.47)。

        チップボール = (本数 + 1) × 0.7mm
        コイル       = 巾 × 本数
        コイル高さ   = コイル + チップボール                       (`tare_calc.get_coil_ta`)
        梱包高さ     = コイル高さ + 樹脂パレット 150mm × 2 + ハードボード 2.35mm × 2
                       → 10mm 単位で切り上げ                     (`tare_calc.set_coil_dimensions` の KTA)

    VBA の誤り(現場の指摘で直した): VBA は ``TA = 1 * col / 10^3 + TIPT`` と **本数を掛けず**
    (「1条分」= 巾 + チップボール。11 本でも 1 本分の高さ)、``KTA`` にハードボード 2.35 × 2 も
    足していなかった。1.2.46 までは「1 条分の参考表示」として VBA のまま出していたが、聞いた本数の
    積み高さになっていなかった。巾は他処理と同じ 53.5(VBA はここだけ 43.5。解析書 13章 #2)。
    表示は小数 1 桁(風袋計算の画面の高さと同じ)。
    """
    tipt = single(val((coil_count + TC.CHIP_EXTRA_SHEETS) * TC.CHIP_THICKNESS_MM / 10 ** 3))
    coils = val(coil_count * width_mm / 10 ** 3)
    ta = single(TC.get_coil_ta(coil_count, width_mm))          # 風袋計算の mat.TA と同じ
    kta = single(val(ta * 1000 + TC.PALLET_HEIGHT_MM * 2) + (TC.HB_THICKNESS_MM * 2))
    pack = ws_roundup(kta, -1)
    n = fmt(coil_count, "0")
    w = fmt(width_mm, "0.0")
    out = {
        "count": n, "width": w,
        "tip": fmt(tipt * 1000, "0.0"),
        "coils": fmt(coils * 1000, "0.0"),
        "coilHeight": fmt(ta * 1000, "0.0"),
        "packHeight": fmt(pack, "0.0"),
        "packRaw": fmt(kta, "0.0"),
        "tipFormula": f"({n}+{TC.CHIP_EXTRA_SHEETS})枚 × {TC.CHIP_THICKNESS_MM}",
        "coilsFormula": f"{w} × {n}本",
        "coilHeightFormula": "コイル + チップボール",
        "packFormula": (f"コイル高さ + パレット {TC.PALLET_HEIGHT_MM}×2 + ハードボード {TC.HB_THICKNESS_MM}×2"
                        f" = {fmt(kta, '0.0')} → 10mm 切り上げ"),
    }
    out["message"] = (f"積み高さ({w}mm × {n}本): チップボール {out['tip']} / コイル {out['coils']} / "
                      f"コイル高さ {out['coilHeight']} / 梱包高さ {out['packHeight']}mm")
    return out
