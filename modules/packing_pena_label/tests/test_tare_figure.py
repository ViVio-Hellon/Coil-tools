# -*- coding: utf-8 -*-
"""風袋計算の図(断面図・重さの内訳。統合 1.2.21・現場の求め)

図は計算しない。風袋計算の結果(`Workflow._tare_payload`)をそのまま描く。ここで見ること:

- 9 資材・NW・風袋・GW の数が、表と同じ文字で出ている(図と表で食い違わない)
- 丸める前の 風袋・GW と、表示の値(四捨五入)の両方が出ている
- 高さの式(積高・梱包高さ・10mm 切り上げ)とエサフォームの周回(積高 820mm 超で 2 周)
- ハードボードの数え方(固定物 HB590 / HB530 か、標準か)
- 画面の風袋計算にだけ出て、印刷ビューには出ない
- SVG として読める(文字の逃がし漏れが無い)
"""
import re
import unittest
import xml.etree.ElementTree as ET

from modules.packing_pena_label.app.services import tare_calc as TC
from modules.packing_pena_label.app.services import tare_figure as TF
from modules.packing_pena_label.app.services.workflow import Workflow

from .test_new_size_calc import _table


def _payload(n, width=53.5, tip="TIP1000", ob=9, t=15.0):
    mat = TC.calculate_materials(_table(), float(n), 1, width, tip, "", ob, t, t, True, False)
    return Workflow._tare_payload(None, 1, float(n), mat, t, t), mat


def _svgs(html):
    return re.findall(r"<svg.*?</svg>", html, re.S)


class TareFigureTest(unittest.TestCase):
    def test_numbers_are_the_same_text_as_the_table(self):
        p, _ = _payload(11)
        html = TF.figure_html(p)
        for key, *_ in TF.ITEMS:
            self.assertIn(f'{p["display"][key]}kg', html, key)
        self.assertIn(f'NW {p["display"]["NW"]}kg', html)
        self.assertIn(f'GW {p["display"]["GW"]}kg', html)
        self.assertIn(f'表示 {p["display"]["HU"]}kg', html)

    def test_raw_and_rounded_totals(self):
        p, mat = _payload(11)
        self.assertEqual(p["huRaw"], "%.3f" % mat.HU)
        self.assertEqual(p["gwRaw"], "%.3f" % mat.GW)
        html = TF.breakdown_svg(p)
        self.assertIn("165.0 + %.3f = %.3f" % (mat.HU, mat.GW), html)
        self.assertIn(f'→ GW {p["display"]["GW"]}kg', html)

    def test_old_payload_without_raw_totals_falls_back_to_the_items(self):
        p, mat = _payload(11)
        p.pop("huRaw"), p.pop("gwRaw")
        hu, gw = TF.raw_totals(p)
        self.assertAlmostEqual(hu, mat.HU, delta=0.01)
        self.assertAlmostEqual(gw, mat.GW, delta=0.01)

    def test_height_formula(self):
        p, mat = _payload(11)
        html = TF.section_svg(p)
        self.assertIn("コイル積高 = 11本 × 53.5mm + チップボール 12枚 × 0.7mm = 596.9mm", html)
        self.assertIn("= 901.6 → 910.0mm(10mm 単位で切り上げ)", html)
        self.assertEqual(p["taDisplay"], "910.0")

    def test_esa_wraps_follow_the_stack_height(self):
        low, _ = _payload(11)      # 596.9mm
        high, _ = _payload(20)     # 1084.7mm
        self.assertIn("エサフォーム 1周", TF.section_svg(low))
        self.assertIn("≦ 820mm なので 1 周", TF.breakdown_svg(low))
        self.assertIn("エサフォーム 2周", TF.section_svg(high))
        self.assertIn("> 820mm なので 2 周", TF.breakdown_svg(high))

    def test_hb_kind(self):
        p, _ = _payload(11)
        self.assertEqual(TF.hb_kind(p), "標準(外周+700mm × 積高)")
        for kind in ("HB590", "HB530"):
            q = dict(p, breakdown=dict(p["breakdown"], HBsiki=f"【ハードボード】{kind}*2000 上下2枚"))
            self.assertEqual(TF.hb_kind(q), f"固定物 {kind}")

    def test_svg_is_well_formed_for_every_tip_and_count(self):
        for tip in ("TIP1000", "TIP950", "TIP820"):
            for n in (1, 4, 11, 20):
                p, _ = _payload(n, tip=tip)
                for svg in _svgs(TF.figure_html(p)):
                    ET.fromstring(svg)          # 読めなければ例外

    def test_labels_stay_inside_the_drawing(self):
        for n in (1, 11, 20):
            p, _ = _payload(n)
            svg = TF.section_svg(p)
            height = float(re.search(r'viewBox="0 0 \d+ (\d+)"', svg).group(1))
            ys = [float(y) for y in re.findall(r'<text x="[\d.]+" y="([-\d.]+)" class="tf-lab"', svg)]
            self.assertEqual(len(ys), 9, n)
            self.assertTrue(all(0 < y < height for y in ys), (n, ys))
            gaps = [b - a for a, b in zip(sorted(ys), sorted(ys)[1:])]
            self.assertTrue(all(g >= 12 for g in gaps), (n, gaps))

    def test_pins_and_bands_follow_the_packing_spec(self):
        """包装仕様書 7C1175 の図のとおり(1.5.35): ピン(樹脂製丸棒)はハードボードの外に立って
        上下のパレットの間に渡り、胴巻きバンドは 4 本で上の方(ラベルの場所)には掛けない。"""
        self.assertEqual(len(TF.BAND_LEVELS), 4)
        self.assertTrue(all(0 < f <= 0.7 for f in TF.BAND_LEVELS), TF.BAND_LEVELS)
        for n in (1, 11, 20):
            p, _ = _payload(n)
            svg = TF.section_svg(p)
            self.assertIn("外から見た所", svg)
            self.assertIn("断面", svg)
            root = ET.fromstring(svg)
            rects = [r for r in root.iter("{http://www.w3.org/2000/svg}rect") if r.get("fill")]
            rects += [r for r in root.iter("rect") if r.get("fill")]
            box = lambda r: (float(r.get("x")), float(r.get("y")),            # noqa: E731
                             float(r.get("width")), float(r.get("height")))
            pallets = sorted((box(r) for r in rects if r.get("fill") == TF.COLOR["PAR"] and box(r)[2] > 50),
                             key=lambda b: b[1])
            top_pallet, bottom_pallet = pallets[0], pallets[-1]
            # 札の色見本(10×10)を除いた、縦長のピン
            pins = [box(r) for r in rects if r.get("fill") == TF.COLOR["PI"] and box(r)[3] > 12]
            self.assertEqual(len(pins), 2, n)          # 左右に見える 2 本(4 本のうち手前)
            for x, y, w, h in pins:
                self.assertAlmostEqual(y, top_pallet[1] + top_pallet[3], delta=0.2)   # 上のパレットの下から
                self.assertGreater(y + h, bottom_pallet[1])                            # 下のパレットの穴へ差す
                self.assertLess(y + h, bottom_pallet[1] + bottom_pallet[3])            # 下には突き抜けない
                self.assertTrue(top_pallet[0] < x < top_pallet[0] + top_pallet[2], (n, x))
            # 胴巻きバンド(外から見た所の横長の帯)は 4 本
            wide = [box(r) for r in rects if r.get("fill") == TF.COLOR["PET"] and box(r)[2] > 20
                    and box(r)[3] < 4]
            self.assertEqual(len([b for b in wide if b[1] > top_pallet[1] + 2]), 4, n)

    def test_small_shares_are_not_shown_as_zero(self):
        self.assertEqual(TF._pct(0.5, 100), "0.5%")
        self.assertEqual(TF._pct(0.02, 100), "0.1%未満")
        self.assertEqual(TF._pct(42.4, 100), "42%")


class TarePageTest(unittest.TestCase):
    """画面の風袋計算にだけ出て、印刷ビューには出ない。"""

    def test_screen_has_figures_print_view_does_not(self):
        from modules.packing_pena_label.app.routes.pages import PageRoutes as Pages

        p, _ = _payload(11)

        class Store:
            def all_tare(self):
                return {1: p}

            # 統合 1.2.49: 風袋計算の画面は、計算したときの資材重量の素性(画面だけ)も控えから読む
            def get_kv(self, key, default=None):
                return default

        class WF:
            store = Store()
            _TARE_MATERIAL_KEY = "tare_material"

            def load_state(self):
                from modules.packing_pena_label.app.models import FormState as WorkState
                return WorkState()

            def tare_is_stale(self, state=None):
                return False

            # 統合 1.2.49: 止める理由(入力が変わった・資材重量が変わった)を分けて聞くようになった
            def tare_stale_info(self, state=None):
                return {"stale": False, "input": False, "material": ""}

            @staticmethod
            def visible_slots(state):
                return [1, 2, 3, 4]

        pages = Pages.__new__(Pages)
        pages.wf = WF()
        pages.cfg = type("Cfg", (), {"url_prefix": ""})()
        pages._shell = lambda title, body, active, script="", print_only=False, **kw: body
        screen = pages.tare(print_view=False)
        printed = pages.tare(print_view=True)
        self.assertIn('class="tfig no-print"', screen)
        self.assertIn("blocks--fig", screen)
        self.assertNotIn("tfig", printed)


if __name__ == "__main__":
    unittest.main()


class BreakdownCoilTest(unittest.TestCase):
    """計算内容の梱包ごとのボタン(?coil=N)で、その梱包が出る(統合 1.2.22)。"""

    def _pages(self, data):
        from modules.packing_pena_label.app.routes.pages import PageRoutes

        class Store:
            def all_tare(self):
                return data

            # 統合 1.2.49: 計算内容の画面は、計算したときの資材マスタの時点を控えから読む
            def get_kv(self, key, default=None):
                return default

        class WF:
            store = Store()
            _TARE_INPUT_KEY = "tare_result_inputs"
            _TARE_MATERIAL_KEY = "tare_material"

        pages = PageRoutes.__new__(PageRoutes)
        pages.wf = WF()
        pages.cfg = type("Cfg", (), {"url_prefix": ""})()
        pages._shell = lambda title, body, active, script="", print_only=False: body
        return pages

    def test_selected_coil_is_shown(self):
        p1, _ = _payload(11)
        p3, _ = _payload(4, width=82.5, tip="TIP820", ob=7)
        p3 = dict(p3, coilNo=3, title="丈2_1梱包目:4本")
        pages = self._pages({1: p1, 3: p3})
        first = pages.breakdown()
        third = pages.breakdown(3)
        self.assertIn(p1["display"]["GW"], first)
        self.assertIn(p3["breakdown"]["NWsiki"], third)
        self.assertNotIn(p1["breakdown"]["NWsiki"], third)
        self.assertIn('aria-current="page" href="/breakdown?coil=3"', third)
        # 計算していない番号なら、今までどおり最小番号
        self.assertEqual(pages.breakdown(4), first)
