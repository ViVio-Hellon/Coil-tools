# -*- coding: utf-8 -*-
"""風袋計算の図 ── GW をどこから出しているかを絵で見せる(統合 1.2.21・現場の求め)

風袋計算の表(資材 9 項目・NW・風袋・GW と計算式)は、数と式しか出ていない。
**どの資材を・どの寸法で数えているか**を、1 梱包ずつ 2 枚の図で出す。

    断面図   梱包を横から見た模式図。高さは寸法どおりの縮尺(樹脂パレット・コイル・
             チップボール・八角ハードボード)。薄い物(チップボール 0.7mm・ハードボード 2.35mm)は
             見えないので太く描き、寸法は数で添える。資材ごとに色を決め、重さを引き出して書く
    内訳     GW = NW + 風袋 の帯と、風袋 9 項目の帯(どれが重いか)。式は計算式の欄と同じもの

**計算はしない。** 描くのは風袋計算の結果(`Workflow._tare_payload`)そのもの。
数は画面・印刷の表と同じ文字を使う(図と表で数が食い違わない)。
図は画面だけに出す(印刷の紙面は VBA のシートと同じ表のまま)。
"""

from __future__ import annotations

import html
import re
from typing import Dict, List, Optional, Tuple

from . import tare_calc as TC

#: 資材の色と呼び名。断面図・内訳の帯・凡例で同じ色を使う
#: (キー: 風袋計算の表示値 `display` のキー)
ITEMS: List[Tuple[str, str, str, str]] = [
    # key,  呼び名,               式のキー,   色
    ("PAR", "樹脂パレット",       "PARsiki", "#4f7cac"),
    ("TIP", "チップボール",       "TIPsiki", "#c39a4e"),
    ("HB",  "ハードボード",       "HBsiki",  "#8b5a2b"),
    ("PI",  "ピン",               "PIsiki",  "#d1495b"),
    ("PET", "PETバンド",          "PETsiki", "#2e9d6a"),
    ("SUT", "ストレッチフィルム", "SUTsiki", "#3fa7d6"),
    ("ESA", "エサフォーム",       "ESAsiki", "#e0a526"),
    ("POR", "ポリシート",         "PORsiki", "#8e6cc0"),
    ("DRY", "乾燥剤(EX-DRY)",     "DRYsiki", "#7d8790"),
]
COLOR = {k: c for k, _, _, c in ITEMS}
NAME = {k: n for k, n, _, _ in ITEMS}
COIL = "#a7b1bc"
COIL_ALT = "#c3cbd4"

#: 断面図で、薄くて見えない物の最低の太さ(px)
MIN_PX = 2.2


def _num(text, default: float = 0.0) -> float:
    try:
        return float(str(text).replace(",", ""))
    except (TypeError, ValueError):
        return default


def _esc(text) -> str:
    return html.escape(str(text), quote=True)


def _lines(formula: str) -> Tuple[str, str]:
    """「【名】説明\\n式」を (説明, 式) に。"""
    text = str(formula or "")
    head, _, body = text.partition("\n")
    head = re.sub(r"^【[^】]*】", "", head).strip()
    return head, body.strip()


def _pct(v: float, whole: float) -> str:
    """割合。1% に満たないものも 0% と出さない(小さくても数えていることが分かるように)。"""
    r = v / whole * 100 if whole else 0.0
    if r >= 1:
        return f"{r:.0f}%"
    return f"{r:.1f}%" if r >= 0.1 else "0.1%未満"


def raw_totals(p: dict) -> Tuple[float, float]:
    """丸める前の 風袋・GW。保存した値が無い(前の版で計算した)ときは表示値から足す。"""
    hu = _num(p.get("huRaw"), -1.0)
    gw = _num(p.get("gwRaw"), -1.0)
    disp = p.get("display") or {}
    if hu < 0:
        hu = sum(_num(disp.get(k)) for k, _, _, _ in ITEMS)
    if gw < 0:
        gw = hu + _num(disp.get("NW"))
    return hu, gw


def hb_kind(p: dict) -> str:
    """ハードボードの数え方(固定物 HB590 / HB530 か、外周×積高の標準か)。"""
    f = str((p.get("breakdown") or {}).get("HBsiki", ""))
    for kind in ("HB590", "HB530"):
        if kind in f:
            return f"固定物 {kind}"
    return "標準(外周+700mm × 積高)"


# ====================================================================== 断面図
def section_svg(p: dict) -> str:
    """梱包を横から見た模式図。高さは寸法どおり(薄い物は太く描く)。"""
    disp = p.get("display") or {}
    n = int(round(_num(p.get("coilText"))))
    width = _num(p.get("width"))
    ta = _num(p.get("coilHeightMm"))
    diam = _num(p.get("tempDiameter"), 1000.0) or 1000.0
    kta_raw = ta + TC.PALLET_HEIGHT_MM * 2 + TC.HB_THICKNESS_MM * 2
    kta_disp = str(p.get("taDisplay", ""))
    pallet = TC.PALLET_HEIGHT_MM
    chip = TC.CHIP_THICKNESS_MM
    hb_t = TC.HB_THICKNESS_MM
    wraps = 2 if ta / 1000.0 > TC.ESA_DOUBLE_LIMIT_M else 1

    W, H = 600, 486
    top, bottom = 52, H - 58
    avail = bottom - top
    # 薄い物は最低の太さで描く。そのぶんを除いた高さで縮尺を決める
    n_chip = max(n + TC.CHIP_EXTRA_SHEETS, 0)
    thin_px = (n_chip + 2) * MIN_PX
    scale = max((avail - thin_px) / max(kta_raw - n_chip * chip - 2 * hb_t, 1.0), 0.05)
    body_w = min(diam * scale, 230.0)
    cx = 168
    x0, x1 = cx - body_w / 2, cx + body_w / 2

    out: List[str] = []
    add = out.append
    add(f'<svg class="tfig-sec" viewBox="0 0 {W} {H}" role="img" '
        f'aria-label="梱包の断面図: {n}本・梱包高さ{_esc(kta_disp)}mm">')
    add('<defs><pattern id="tfFilm" width="6" height="6" patternUnits="userSpaceOnUse" '
        'patternTransform="rotate(45)"><line x1="0" y1="0" x2="0" y2="6" '
        f'stroke="{COLOR["SUT"]}" stroke-width="1.2" opacity=".55"/></pattern></defs>')

    y = float(bottom)
    labels: List[Tuple[float, str, str, str]] = []   # (y, key, 名, 値)

    def rect(x, yy, w, h, fill, extra=""):
        add(f'<rect x="{x:.1f}" y="{yy:.1f}" width="{w:.1f}" height="{max(h, 0.6):.1f}" '
            f'fill="{fill}" {extra}/>')

    # --- 下の樹脂パレット(脚つき) ---
    ph = pallet * scale
    pw = body_w + 26
    px = cx - pw / 2
    for i in range(3):
        fx = px + 6 + i * (pw - 30) / 2
        rect(fx, y - ph * 0.68, 18, ph * 0.68, COLOR["PAR"], 'rx="1.5" opacity=".85"')
    rect(px, y - ph, pw, ph * 0.32, COLOR["PAR"], 'rx="2"')
    deck_bottom = y - ph
    labels.append((y - ph * 0.5, "PAR", "樹脂パレット 上下2台", disp.get("PAR", "")))
    y -= ph
    # --- 下の八角ハードボード ---
    hpx = max(hb_t * scale, MIN_PX)
    rect(x0 - 6, y - hpx, body_w + 12, hpx, COLOR["HB"])
    y -= hpx
    coil_bottom = y
    # --- コイルとチップボール(本数 + 1 枚) ---
    cpx = max(chip * scale, MIN_PX)
    coil_px = width * scale
    for i in range(n):
        rect(x0, y - cpx, body_w, cpx, COLOR["TIP"])
        y -= cpx
        fill = COIL if i % 2 == 0 else COIL_ALT
        rect(x0, y - coil_px, body_w, coil_px, fill, 'stroke="#56616c" stroke-width=".6"')
        # コイルの穴(内径)を模式で
        add(f'<line x1="{cx:.1f}" y1="{y - coil_px + 1:.1f}" x2="{cx:.1f}" y2="{y - 1:.1f}" '
            'stroke="#56616c" stroke-width=".5" stroke-dasharray="2 2" opacity=".6"/>')
        y -= coil_px
    for _ in range(max(n_chip - n, 0)):
        rect(x0, y - cpx, body_w, cpx, COLOR["TIP"])
        y -= cpx
    coil_top = y
    # --- 上の八角ハードボード・樹脂パレット ---
    rect(x0 - 6, y - hpx, body_w + 12, hpx, COLOR["HB"])
    y -= hpx
    for i in range(3):
        fx = px + 6 + i * (pw - 30) / 2
        rect(fx, y - ph, 18, ph * 0.68, COLOR["PAR"], 'rx="1.5" opacity=".85"')
    rect(px, y - ph * 0.32, pw, ph * 0.32, COLOR["PAR"], 'rx="2"')
    rect(px, y - ph, pw, ph * 0.32, COLOR["PAR"], 'rx="2"')
    top_pallet_top = y - ph
    y -= ph

    # --- 胴巻きハードボード(コイルの両脇) ---
    band_h = coil_bottom - coil_top            # 上ほど y が小さい
    rect(x0 - 5, coil_top, 4, coil_bottom - coil_top, COLOR["HB"], 'opacity=".9"')
    rect(x1 + 1, coil_top, 4, coil_bottom - coil_top, COLOR["HB"], 'opacity=".9"')
    # --- ポリシート・エサフォーム・ストレッチ(外側を包む。模式) ---
    add(f'<rect x="{x0 - 9:.1f}" y="{coil_top - 6:.1f}" width="{body_w + 18:.1f}" '
        f'height="{band_h + 12:.1f}" fill="none" stroke="{COLOR["POR"]}" stroke-width="1.6" '
        'stroke-dasharray="5 3" rx="3"/>')
    for k in range(wraps):
        add(f'<rect x="{x0 - 13 - k * 3:.1f}" y="{coil_top + band_h * 0.18:.1f}" '
            f'width="{body_w + 26 + k * 6:.1f}" height="{band_h * 0.64:.1f}" fill="none" '
            f'stroke="{COLOR["ESA"]}" stroke-width="2.2" rx="3" opacity=".9"/>')
    add(f'<rect x="{px - 4:.1f}" y="{top_pallet_top + ph * 0.4:.1f}" width="{pw + 8:.1f}" '
        f'height="{deck_bottom + ph * 0.6 - top_pallet_top - ph * 0.4:.1f}" fill="url(#tfFilm)" '
        f'stroke="{COLOR["SUT"]}" stroke-width="1" rx="4" opacity=".75"/>')
    # --- PET バンド: 縦 4 本 ×2(見えるのは 2 本)・胴巻き 4 本 ---
    for fx in (cx - body_w * 0.28, cx + body_w * 0.28):
        rect(fx - 2, top_pallet_top - 2, 4, deck_bottom - top_pallet_top + ph + 4, COLOR["PET"],
             'opacity=".85"')
    for i in range(4):
        by = coil_top + band_h * (0.14 + 0.24 * i)
        rect(px - 6, by - 1.6, pw + 12, 3.2, COLOR["PET"], 'opacity=".85"')
    # --- ピン(パレットの角。4 本) ---
    for fx in (px + 3, px + pw - 7):
        for fy in (top_pallet_top + ph * 0.5, deck_bottom + ph * 0.18):
            rect(fx, fy - 5, 4, 10, COLOR["PI"], 'rx="1"')
    # --- 乾燥剤(上の段に 1 つ。模式) ---
    rect(cx - 9, coil_top - hpx - 12, 18, 9, COLOR["DRY"], 'rx="2"')

    # --- 引き出し線と重さ(右側) ---
    labels += [
        (coil_top - 6, "DRY", "乾燥剤", disp.get("DRY", "")),
        (top_pallet_top + ph * 0.5 + 0.1, "PI", "ピン 4本", disp.get("PI", "")),
        (coil_top + band_h * 0.14, "PET", "PETバンド", disp.get("PET", "")),
        (coil_top + band_h * 0.30, "SUT", "ストレッチフィルム 5周", disp.get("SUT", "")),
        (coil_top + band_h * 0.46, "ESA", f"エサフォーム {wraps}周", disp.get("ESA", "")),
        (coil_top + band_h * 0.62, "POR", "ポリシート", disp.get("POR", "")),
        (coil_top + band_h * 0.78, "HB", "ハードボード(胴巻き+八角)", disp.get("HB", "")),
        (coil_bottom - cpx * 0.5 - 0.1, "TIP", f"チップボール {n_chip}枚", disp.get("TIP", "")),
    ]
    _leaders(add, labels, x_from=px + pw + 8, x_text=W - 222, top=top, bottom=bottom)

    # --- 寸法(左側): コイル積高・梱包高さ ---
    _dim(add, 38, coil_top, coil_bottom, f"{ta:.1f}", "コイル積高")
    _dim(add, 14, top_pallet_top, float(bottom), f"{kta_disp}", "梱包高さ")
    add(f'<text x="{cx:.1f}" y="{coil_top + band_h / 2:.1f}" class="tf-in" text-anchor="middle">'
        f'コイル {n}本</text>')
    add(f'<text x="{cx:.1f}" y="{coil_top + band_h / 2 + 15:.1f}" class="tf-in tf-sm" '
        f'text-anchor="middle">NW {_esc(disp.get("NW", ""))}kg</text>')
    # 床と外径
    add(f'<line x1="{px - 14:.1f}" y1="{bottom + .5}" x2="{px + pw + 14:.1f}" y2="{bottom + .5}" '
        'class="tf-floor"/>')
    add(f'<text x="{cx:.1f}" y="{bottom + 18}" class="tf-sm" text-anchor="middle">'
        f'外径 Φ{_esc(p.get("tempDiameter", ""))}mm ・ 外周 {_esc(p.get("gai", ""))}m</text>')
    add(f'<text x="{W / 2:.0f}" y="{H - 10}" class="tf-note" text-anchor="middle">'
        '模式図。高さは寸法どおり(チップボール・ハードボードは見えるよう太く描いています)。'
        '包む物の位置は目安</text>')
    add(f'<text x="8" y="16" class="tf-sm">コイル積高 = {n}本 × {width:g}mm + '
        f'チップボール {n_chip}枚 × {chip:g}mm = {ta:.1f}mm</text>')
    add(f'<text x="8" y="32" class="tf-sm">梱包高さ = 積高 + 樹脂パレット {pallet:g}×2 + '
        f'八角HB {hb_t:g}×2 = {kta_raw:.1f} → {_esc(kta_disp)}mm(10mm 単位で切り上げ)</text>')
    add('</svg>')
    return "".join(out)


def _dim(add, x: float, y_top: float, y_bot: float, value: str, label: str) -> None:
    add(f'<g class="tf-dim"><line x1="{x}" y1="{y_top:.1f}" x2="{x}" y2="{y_bot:.1f}"/>'
        f'<line x1="{x - 4}" y1="{y_top:.1f}" x2="{x + 4}" y2="{y_top:.1f}"/>'
        f'<line x1="{x - 4}" y1="{y_bot:.1f}" x2="{x + 4}" y2="{y_bot:.1f}"/>'
        f'<text transform="translate({x - 3},{(y_top + y_bot) / 2:.1f}) rotate(-90)" '
        f'text-anchor="middle" class="tf-sm">{_esc(label)} {_esc(value)}mm</text></g>')


def _leaders(add, labels, *, x_from: float, x_text: float, top: float, bottom: float) -> None:
    """重さの札を右に縦に並べ、描いた物から線で結ぶ(札どうしが重ならないよう間を空ける)。"""
    items = sorted(labels, key=lambda t: t[0])
    if not items:
        return
    lo, hi = top + 4, bottom + 10
    gap = min(22.0, (hi - lo) / max(len(items) - 1, 1))
    slots = [min(max(want, lo), hi) for want, *_ in items]
    for i in range(1, len(slots)):                  # 上から: 前の札と間を空ける
        slots[i] = max(slots[i], slots[i - 1] + gap)
    if slots[-1] > hi:                              # 下へはみ出したら、下から押し上げる
        slots[-1] = hi
        for i in range(len(slots) - 2, -1, -1):
            slots[i] = min(slots[i], slots[i + 1] - gap)
    for (want, key, name, value), y in zip(items, slots):
        color = COLOR.get(key, "#888")
        add(f'<polyline points="{x_from:.1f},{want:.1f} {x_text - 14:.1f},{y:.1f} {x_text - 6:.1f},{y:.1f}" '
            f'fill="none" stroke="{color}" stroke-width="1.1"/>')
        add(f'<rect x="{x_text - 4:.1f}" y="{y - 6:.1f}" width="10" height="10" rx="2" fill="{color}"/>')
        val = f' <tspan class="tf-v">{_esc(value)}kg</tspan>' if value else ""
        add(f'<text x="{x_text + 11:.1f}" y="{y + 3.5:.1f}" class="tf-lab">{_esc(name)}{val}</text>')


# ====================================================================== 内訳
def breakdown_svg(p: dict) -> str:
    """GW = NW + 風袋 の帯と、風袋 9 項目の帯。"""
    disp = p.get("display") or {}
    formulas = p.get("breakdown") or {}
    nw = _num(disp.get("NW"))
    hu, gw = raw_totals(p)
    vals = [(k, _num(disp.get(k))) for k, _, _, _ in ITEMS]
    W = 600
    row_h = 34
    H = 150 + row_h * len(ITEMS) + 40
    out: List[str] = []
    add = out.append
    add(f'<svg class="tfig-bar" viewBox="0 0 {W} {H}" role="img" '
        f'aria-label="GW の内訳: NW {_esc(disp.get("NW", ""))}kg + 風袋 {hu:.3f}kg">')
    # --- GW = NW + 風袋 ---
    add(f'<text x="10" y="22" class="tf-h">GW = NW + 風袋</text>')
    add(f'<text x="10" y="42" class="tf-sm">{nw:.1f} + {hu:.3f} = {gw:.3f} '
        f'→ GW {_esc(disp.get("GW", ""))}kg(小数第1位を四捨五入して整数に)</text>')
    bx, bw, by, bh = 10.0, W - 20.0, 54.0, 30.0
    total = max(gw, 1e-9)
    nw_w = bw * nw / total
    add(f'<rect x="{bx}" y="{by}" width="{nw_w:.1f}" height="{bh}" fill="{COIL}" rx="3"/>')
    add(f'<text x="{bx + 8}" y="{by + 20}" class="tf-in">NW(コイル) {nw:.1f}kg '
        f'{nw / total * 100:.0f}%</text>')
    add(f'<text x="{bx}" y="{by + bh + 16}" class="tf-sm">NW = 1条重量 × 本数: '
        f'{_esc(formulas.get("NWsiki", ""))}</text>')
    x = bx + nw_w
    for key, v in vals:
        w = bw * v / total
        add(f'<rect x="{x:.1f}" y="{by}" width="{max(w, 0):.2f}" height="{bh}" fill="{COLOR[key]}"/>')
        x += w
    add(f'<text x="{bx + bw:.1f}" y="{by + bh + 16}" class="tf-sm" text-anchor="end">'
        f'風袋 {hu:.3f}kg({hu / total * 100:.0f}%) → 表示 {_esc(disp.get("HU", ""))}kg</text>')

    # --- 風袋 9 項目 ---
    y0 = 130.0
    add(f'<text x="10" y="{y0 - 8}" class="tf-h">風袋の内訳(9 項目)</text>')
    biggest = max([v for _, v in vals] + [1e-9])
    bar_x = 160.0
    bar_w = W - bar_x - 140
    for i, (key, v) in enumerate(vals):
        y = y0 + i * row_h
        head, body = _lines(formulas.get(dict((k, f) for k, _, f, _ in ITEMS)[key], ""))
        add(f'<rect x="10" y="{y + 3}" width="10" height="10" rx="2" fill="{COLOR[key]}"/>')
        add(f'<text x="26" y="{y + 12}" class="tf-lab">{_esc(NAME[key])}</text>')
        add(f'<rect x="{bar_x}" y="{y + 2}" width="{max(bar_w * v / biggest, 1.0):.1f}" height="12" '
            f'fill="{COLOR[key]}" rx="2"/>')
        add(f'<text x="{W - 10}" y="{y + 12}" class="tf-v" text-anchor="end">'
            f'{_esc(disp.get(key, ""))}kg ・ {_pct(v, hu)}</text>')
        # ポリシートの式の「5502」は 550 の 2 乗(計算式の欄の書き方のまま)。図では ² で出す
        tip = re.sub(r"×(\d+)2 ", r"×\1² ", body or head)
        add(f'<text x="26" y="{y + 27}" class="tf-f">{_esc(tip)}</text>')
    yb = y0 + len(ITEMS) * row_h + 14
    add(f'<text x="10" y="{yb}" class="tf-sm">ハードボード: {_esc(hb_kind(p))}'
        f' ・ エサフォーム: 積高 {_num(p.get("coilHeightMm")):.1f}mm '
        f'{"> 820mm なので 2 周" if _num(p.get("coilHeightMm")) > TC.ESA_DOUBLE_LIMIT_M * 1000 else "≦ 820mm なので 1 周"}'
        '</text>')
    add('</svg>')
    return "".join(out)


def figure_html(p: dict) -> str:
    """1 梱包ぶんの図(断面図と内訳)。風袋計算の画面のブロックの下に置く。"""
    return ('<details class="tfig no-print" open>'
            '<summary>図で見る(GW をどこから出しているか)</summary>'
            '<div class="tfig-row">'
            f'<figure>{section_svg(p)}<figcaption>断面図(横から見た模式図)</figcaption></figure>'
            f'<figure>{breakdown_svg(p)}<figcaption>重さの内訳</figcaption></figure>'
            '</div></details>')
