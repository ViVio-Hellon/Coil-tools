# -*- coding: utf-8 -*-
"""風袋計算の図 ── GW をどこから出しているかを絵で見せる(統合 1.2.21・現場の求め)

風袋計算の表(資材 9 項目・NW・風袋・GW と計算式)は、数と式しか出ていない。
**どの資材を・どの寸法で数えているか**を、1 梱包ずつ 2 枚の図で出す。

    断面図   梱包を横から見た模式図。左半分は外から見た所、右半分は断面。包装仕様書
             (7C1175 4. 荷姿の「内装」「前」「横」)に合わせ、ピン(樹脂製丸棒)はハードボードの外に
             立て、PET バンドは胴巻き 4 本と縦(十文字)に描く(統合 1.2.34 で描き直し)。
             高さは寸法どおりの縮尺。薄い物(チップボール 0.7mm・ハードボード 2.35mm)は
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
#: 胴巻き PET バンドの高さ(コイル積高の下からの割合)。包装仕様書 7C1175 の「前」「横」の図に合わせ、
#: 下に 2 本・中ほどに 2 本。上の 3 割ほどはラベルを貼る所なのでバンドを掛けない
#: (「全てのラベルは胴巻きペットバンドの上にのらないように貼ること」)
BAND_LEVELS = (0.12, 0.26, 0.51, 0.64)
#: 縦の PET バンド(十文字に 2 本ずつ)の位置。中心からコイル外径の何割か(仕様書の図に合わせる)
TATE_OFFSET = 0.15


def section_svg(p: dict) -> str:
    """梱包を横から見た模式図。左半分は外から見た所、右半分は断面(中の重なり)。

    包装仕様書(7C1175「包装仕様書」4. 荷姿)の「内装」「前」「横」に合わせて描く。

    - 樹脂パレット: 上下 2 台(上は裏返し)。フォーク爪の差し込み口がある
    - 八角形緩衝材(八角ハードボード): 下のパレットの上と、上のパレットの下
    - チップボール: いちばん下・コイルの間・いちばん上
    - ストレッチフィルム → ポリシート → エサフォーム(発泡ポリシート)→ ハードボード の順に外周を包む
    - ピン(樹脂製丸棒)4 本: ハードボードの外に立て、下のパレットの穴に差す
      (フォーク爪の差し込み口より下に出さない・上のパレットより上に突き出さない)
    - PET バンド: 胴巻き 4 本(ハードボード用・丸棒用)と、上下のパレットを十文字に 2 本ずつ締める縦 4 本
    - 乾燥剤: コイルの内径部に入れる

    高さは寸法どおりの縮尺。薄い物(チップボール 0.7mm・ハードボード 2.35mm)は見えないので太く描く。
    """
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

    W, H = 640, 500
    top, bottom = 56, H - 62
    avail = bottom - top
    # 薄い物は最低の太さで描く。そのぶんを除いた高さで縮尺を決める
    n_chip = max(n + TC.CHIP_EXTRA_SHEETS, 0)
    thin_px = (n_chip + 2) * MIN_PX
    scale = max((avail - thin_px) / max(kta_raw - n_chip * chip - 2 * hb_t, 1.0), 0.05)
    body_w = min(diam * scale, 220.0)
    cx = 190.0
    x0, x1 = cx - body_w / 2, cx + body_w / 2
    # 外周を包む物の厚み(模式): フィルム・ポリシート・エサフォーム(周回ぶん)・ハードボード
    sut_x, por_x, esa_x = 1.0, 3.5, 6.0
    hb_x = esa_x + 3.0 * wraps + 1.0
    hb_w = 4.0
    side = hb_x + hb_w                 # コイルの縁からハードボードの外側まで
    bar_w = 6.0
    bar_x = side + 1.0                 # 丸棒の内側
    band_x = bar_x + bar_w             # 胴巻きバンドの外側(丸棒を締める)
    pw = body_w + 2 * (band_x + 12)
    px = cx - pw / 2

    out: List[str] = []
    add = out.append
    add(f'<svg class="tfig-sec" viewBox="0 0 {W} {H}" role="img" '
        f'aria-label="梱包の断面図: {n}本・梱包高さ{_esc(kta_disp)}mm">')

    def rect(x, yy, w, h, fill, extra=""):
        add(f'<rect x="{x:.1f}" y="{yy:.1f}" width="{max(w, 0.6):.1f}" height="{max(h, 0.6):.1f}" '
            f'fill="{fill}" {extra}/>')

    def pallet_at(y_top: float, h: float) -> None:
        """樹脂パレット 1 台(横から見た所)。フォーク爪の差し込み口が 2 つ。"""
        rect(px, y_top, pw, h, COLOR["PAR"], 'rx="2"')
        for f in (0.17, 0.61):
            rect(px + pw * f, y_top + h * 0.3, pw * 0.22, h * 0.42, "currentColor",
                 'class="tf-hole" rx="1"')

    # ------------------------------------------------ 高さ(下から積む)
    ph = pallet * scale
    hpx = max(hb_t * scale, MIN_PX)
    cpx = max(chip * scale, MIN_PX)
    coil_px = width * scale
    deck_top = bottom - ph                       # 下のパレットの上面
    coil_bottom = deck_top - hpx                 # 下の八角の上 = 積みの下端
    coil_top = coil_bottom - (n * coil_px + n_chip * cpx)
    top_pallet_bottom = coil_top - hpx
    top_pallet_top = top_pallet_bottom - ph
    band_h = coil_bottom - coil_top

    # ------------------------------------------------ パレット・八角形緩衝材
    pallet_at(deck_top, ph)
    pallet_at(top_pallet_top, ph)
    for yy in (deck_top - hpx, coil_top - hpx):
        rect(px + 6, yy, pw - 12, hpx, COLOR["HB"])

    # ------------------------------------------------ 右半分: 断面
    # コイルとチップボール(いちばん下・コイルの間・いちばん上)。内径部は空けて乾燥剤を置く
    hole = body_w * 0.16
    y = coil_bottom
    chips_left = n_chip
    for i in range(n):
        rect(cx, y - cpx, body_w / 2, cpx, COLOR["TIP"])
        chips_left -= 1
        y -= cpx
        fill = COIL if i % 2 == 0 else COIL_ALT
        rect(cx + hole, y - coil_px, body_w / 2 - hole, coil_px, fill,
             'stroke="#56616c" stroke-width=".6"')
        y -= coil_px
    for _ in range(max(chips_left, 0)):
        rect(cx, y - cpx, body_w / 2, cpx, COLOR["TIP"])
        y -= cpx
    add(f'<line x1="{cx + hole:.1f}" y1="{coil_top:.1f}" x2="{cx + hole:.1f}" y2="{coil_bottom:.1f}" '
        'stroke="#56616c" stroke-width=".6" stroke-dasharray="2 2"/>')
    dry_y = coil_bottom - band_h * 0.5
    rect(cx + 2, dry_y - 6, hole - 4, 12, COLOR["DRY"], 'rx="2"')
    # 外周を包む物(右の縁)
    film = f'stroke="{COLOR["SUT"]}" stroke-width="1.6"'
    add(f'<line x1="{x1 + sut_x:.1f}" y1="{coil_top:.1f}" x2="{x1 + sut_x:.1f}" y2="{coil_bottom:.1f}" {film}/>')
    # ポリシートは上下も包む
    add(f'<polyline points="{cx:.1f},{coil_top - 1.2:.1f} {x1 + por_x:.1f},{coil_top - 1.2:.1f} '
        f'{x1 + por_x:.1f},{coil_bottom + 1.2:.1f} {cx:.1f},{coil_bottom + 1.2:.1f}" fill="none" '
        f'stroke="{COLOR["POR"]}" stroke-width="1.4" stroke-dasharray="4 2"/>')
    for k in range(wraps):
        xe = x1 + esa_x + 3.0 * k
        add(f'<line x1="{xe:.1f}" y1="{coil_top:.1f}" x2="{xe:.1f}" y2="{coil_bottom:.1f}" '
            f'stroke="{COLOR["ESA"]}" stroke-width="2.2"/>')
    rect(x1 + hb_x, coil_top, hb_w, band_h, COLOR["HB"])
    # 丸棒(右)。下のパレットの穴に差す(差し込み口より下に出さない)・上のパレットの下まで
    bar_bottom = deck_top + ph * 0.3
    rect(x1 + bar_x, top_pallet_bottom, bar_w, bar_bottom - top_pallet_bottom, COLOR["PI"], 'rx="1.5"')
    # 胴巻きバンドの切り口(丸棒の外)
    band_ys = [coil_bottom - band_h * f for f in BAND_LEVELS]
    for by in band_ys:
        rect(x1 + band_x, by - 1.6, 3, 3.2, COLOR["PET"])

    # ------------------------------------------------ 左半分: 外から見た所
    rect(x0 - side, coil_top, cx - x0 + side, band_h, COLOR["HB"], 'opacity=".72"')
    rect(x0 - bar_x - bar_w, top_pallet_bottom, bar_w, bar_bottom - top_pallet_bottom, COLOR["PI"],
         'rx="1.5"')
    for by in band_ys:
        rect(x0 - band_x - 3, by - 1.6, cx - x0 + band_x + 3, 3.2, COLOR["PET"])
    # 縦のバンド(上下のパレットを十文字に締める。フォーク爪の差し込み口の内側を通す)
    xt = cx - body_w * TATE_OFFSET
    rect(xt - 1.6, top_pallet_top - 1.6, 3.2, bottom - top_pallet_top + 1.6, COLOR["PET"])
    rect(xt, top_pallet_top - 1.6, cx - xt + body_w * TATE_OFFSET, 1.6, COLOR["PET"])
    # 真ん中の線(左右の境)
    add(f'<line x1="{cx:.1f}" y1="{top_pallet_top - 8:.1f}" x2="{cx:.1f}" y2="{bottom + 4:.1f}" '
        'class="tf-center"/>')
    add(f'<text x="{(px + cx) / 2:.1f}" y="{top_pallet_top - 6:.1f}" class="tf-sm" '
        'text-anchor="middle">外から見た所</text>')
    add(f'<text x="{(cx + px + pw) / 2:.1f}" y="{top_pallet_top - 6:.1f}" class="tf-sm" '
        'text-anchor="middle">断面</text>')

    # --- 引き出し線と重さ(右側)。線は描いた物から出す ---
    side_y = lambda f: coil_bottom - band_h * f   # noqa: E731
    labels = [
        (deck_top + ph * 0.5, "PAR", "樹脂パレット 上下2台", disp.get("PAR", ""), px + pw),
        (coil_bottom - cpx * 0.5 - 0.1, "TIP", f"チップボール {n_chip}枚", disp.get("TIP", ""),
         x1 - 4),
        (side_y(0.92), "HB", "ハードボード(胴巻き+八角)", disp.get("HB", ""), x1 + hb_x + hb_w / 2),
        (side_y(0.84), "ESA", f"エサフォーム {wraps}周", disp.get("ESA", ""), x1 + esa_x),
        (side_y(0.76), "POR", "ポリシート", disp.get("POR", ""), x1 + por_x),
        (side_y(0.70), "SUT", "ストレッチフィルム 5周", disp.get("SUT", ""), x1 + sut_x),
        (band_ys[2], "PET", "PETバンド 胴巻き4・縦4本", disp.get("PET", ""), x1 + band_x + 3),
        (side_y(0.38), "PI", "ピン(樹脂製丸棒) 4本", disp.get("PI", ""), x1 + bar_x + bar_w),
        (dry_y, "DRY", "乾燥剤(内径部)", disp.get("DRY", ""), cx + hole - 2),
    ]
    _leaders(add, labels, x_out=px + pw + 8, x_text=W - 226, top=top, bottom=bottom)

    # --- 寸法(左側): コイル積高・梱包高さ ---
    _dim(add, 38, coil_top, coil_bottom, f"{ta:.1f}", "コイル積高")
    _dim(add, 14, top_pallet_top, float(bottom), f"{kta_disp}", "梱包高さ")
    # 本数と NW は外から見た所の上の方(ラベルを貼る所。バンドの無い所)に札で書く
    tx = (x0 - side + xt) / 2
    ty = coil_bottom - band_h * 0.84
    rect(tx - 44, ty - 13, 88, 32, "#fffdf6", 'stroke="#56616c" stroke-width=".6" rx="2"')
    add(f'<text x="{tx:.1f}" y="{ty + 1:.1f}" class="tf-in" text-anchor="middle">コイル {n}本</text>')
    add(f'<text x="{tx:.1f}" y="{ty + 14:.1f}" class="tf-in tf-sm" text-anchor="middle">'
        f'NW {_esc(disp.get("NW", ""))}kg</text>')
    # 床と外径
    add(f'<line x1="{px - 14:.1f}" y1="{bottom + .5}" x2="{px + pw + 14:.1f}" y2="{bottom + .5}" '
        'class="tf-floor"/>')
    add(f'<text x="{cx:.1f}" y="{bottom + 17}" class="tf-sm" text-anchor="middle">'
        f'外径 Φ{_esc(p.get("tempDiameter", ""))}mm ・ 外周 {_esc(p.get("gai", ""))}m</text>')
    add(f'<text x="{W / 2:.0f}" y="{H - 24}" class="tf-note" text-anchor="middle">'
        '模式図。左半分は外から見た所・右半分は断面。高さは寸法どおり(薄い物は太く描く)</text>')
    add(f'<text x="{W / 2:.0f}" y="{H - 10}" class="tf-note" text-anchor="middle">'
        '丸棒・バンドの位置は包装仕様書 7C1175 の図に合わせた目安(上の方はラベルの場所)</text>')
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


def _leaders(add, labels, *, x_out: float, x_text: float, top: float, bottom: float) -> None:
    """重さの札を右に縦に並べ、描いた物から線で結ぶ(札どうしが重ならないよう間を空ける)。

    labels: (y, キー, 名, 値, 線を出す x)。線は描いた物の点(小さな丸)から右へ出て、札へ折れる。
    """
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
    for (want, key, name, value, x_from), y in zip(items, slots):
        color = COLOR.get(key, "#888")
        add(f'<circle cx="{x_from:.1f}" cy="{want:.1f}" r="1.8" fill="{color}" class="tf-pt"/>')
        add(f'<polyline points="{x_from:.1f},{want:.1f} {x_out:.1f},{want:.1f} '
            f'{x_text - 14:.1f},{y:.1f} {x_text - 6:.1f},{y:.1f}" '
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
        # 2 乗は計算式の欄では「^2」。図では ² で出す(ペナラベル 1.5.51 まで VBA のまま「5502」と出ていた)
        tip = (body or head).replace("^2", "²")
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
