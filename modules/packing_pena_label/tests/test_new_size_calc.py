# -*- coding: utf-8 -*-
"""足したサイズで GW・HB・バーコードが生きるか(統合 1.2.16・現場の求め)

確かめ方
    1. **双子で突き合わせる。** 同梱の 7 サイズそれぞれに、同じ巾で厚みだけ違うサイズを足し、
       1〜20 本 × TIP 3 種 × 丈1・丈2・同時 で、資材 9 項目・HB・NW・風袋・GW・高さ・式の文字まで
       同梱のサイズと**全部同じ**になることを見る(計算に使うのは巾と本数。厚みは使わない)。
       同梱のサイズは VBA と照らし済み(test_tare_calc ほか)なので、足したサイズも VBA と同じになる
    2. **新しい巾**(45.0mm)は、同じ計算関数に巾を渡したものと一致し、HB は標準の式
       (外周 + 700mm × 積高 × HB 単重 + 八角)になることを見る
    3. **画面と同じ道で**(API: サイズを選ぶ → 重量反映 → 重量計算_DB)、同梱と双子の結果・
       ラベルへ写る NW/GW が同じになることを見る
    4. **バーコードを読み戻す。** 印刷ビューの SVG のバーを測って Code 39 を読み、型番と一致するか。
       型番を直したら、刷るバーコードも変わるか。型番マスタに型番を足したら、
       再起動せずに全サイズの一覧に出て、刷るバーコードがその型番になるか
"""
import csv
import json
import os
import re
import shutil
import sqlite3
import tempfile
import threading
import unittest
import urllib.error
import urllib.parse
import urllib.request

from modules.packing_pena_label import server
from modules.packing_pena_label.app.config import Config
from modules.packing_pena_label.app.models import TIP_CHOICES
from modules.packing_pena_label.app.repositories import csv_text
from modules.packing_pena_label.app.repositories.material_repo import MaterialRepository
from modules.packing_pena_label.app.services import barcode39 as BC
from modules.packing_pena_label.app.services import settings as S
from modules.packing_pena_label.app.services import size_master as SM
from modules.packing_pena_label.app.services import tare_calc as TC
from modules.packing_pena_label.app.services.size_master import SizeMaster

from .test_master_admin import free_port
from ._material import business_csv  # noqa: E402

T1, T2 = 101.35, 99.8          # 1条重量(小数2桁: 丸めの境目も通る)


def _table():
    repo = MaterialRepository(accdb_dir="", csv_path=Config().default_material_csv_path,
                              prefer_access=False)
    return repo.load(use_cache=False)


def _twin_master():
    """同梱の 7 サイズに、同じ巾・厚み 9.x の双子を足したマスタ。双子の名前 → 同梱の名前。"""
    m = SizeMaster()
    firsts = {}
    for c in m.all():
        firsts.setdefault(c.base_name, c)
    rows, twins, ob = [], {}, SM.FIRST_EXTRA_OB
    for i, (base, c) in enumerate(sorted(firsts.items())):
        name = SM.normalize_size_name("9.%dmm×%smm" % (i + 1, ("%g" % c.width_mm)))
        rows.append({"obIdx": ob, "baseName": name, "widthMm": c.width_mm,
                     "kataban": "TWIN%02d" % i, "symbol": "", "note": ""})
        twins[name] = base
        ob += 2
    m.apply_extra(rows)
    return m, twins


def _selections(m: SizeMaster, base: str):
    """そのサイズの (呼び名, cb_key, ob_idx, coil_no, take1_visible, take2_visible) 3 通り。

    ワークフローの ``_selection_keys`` と同じ形(同時は CheckBox 名、丈1/丈2 は台紙番号)。
    """
    cfgs = sorted((c for c in m.all() if c.base_name == base), key=lambda c: c.ob_idx)
    c1, c2 = cfgs[0], cfgs[1]
    cb = c1.cb_key if c1.cb_key != SM.NAMED_CB_KEY else SM.NAMED_CB_KEY
    return [("丈1", "", c1.ob_idx, 1, True, False),
            ("丈2", "", c2.ob_idx, 3, False, True),
            ("同時", cb, 0, 1, True, True),
            ("同時(丈2側)", cb, 0, 3, True, True)]


def _calc(table, n, width, tip, sel):
    _, cb, ob, coil_no, v1, v2 = sel
    mat = TC.calculate_materials(table, float(n), coil_no, width, tip, cb, ob,
                                 T1, T2, v1, v2, all_coil_empty=False)
    d = {k: v for k, v in vars(mat).items() if k != "missing"}
    # `_HBFlag` は VBA の旗そのもの(計算内容の画面に出す控え)。同梱のサイズは VBA の名前で決まり、
    # 足したサイズは巾で決まるので、**結果に効かない本数**(40mm の 14 本など)では旗だけ違うことがある。
    # 旗を除いた全部(HB の値・式の文字を含む)が同じことを見る。旗の違いは別の試験で数える
    d["formulas"] = {k: v for k, v in d["formulas"].items() if k != "_HBFlag"}
    return d


class TwinSizeCalcTest(unittest.TestCase):
    """1. 同じ巾の双子は、同梱のサイズと全部同じ結果になる。"""

    @classmethod
    def setUpClass(cls):
        cls.table = _table()
        cls.master, cls.twins = _twin_master()

    def test_双子は全サイズ_1から20本_TIP3種_丈1丈2同時で同じ(self):
        compared = 0
        for twin, base in self.twins.items():
            width = self.master.width_of(min(c.ob_idx for c in self.master.all()
                                             if c.base_name == base))
            for (label, *_), sel_b, sel_t in zip(_selections(self.master, base),
                                                  _selections(self.master, base),
                                                  _selections(self.master, twin)):
                self.assertTrue(SM.is_extra_selection(sel_t[1], sel_t[2]), "双子は足したサイズとして選ばれる")
                for tip in TIP_CHOICES:
                    for n in range(1, 21):
                        a = _calc(self.table, n, width, tip, sel_b)
                        b = _calc(self.table, n, width, tip, sel_t)
                        if a != b:
                            diff = {k: (a[k], b[k]) for k in a if a[k] != b[k]}
                            self.fail("%s(%s) %s %s %d本 で違う: %s" % (twin, base, label, tip, n, diff))
                        compared += 1
        self.assertEqual(compared, 7 * 4 * 3 * 20)

    def test_旗が違うのは結果に効かない本数だけ(self):
        """同梱のサイズで VBA の旗が立つのに足したサイズで立たないのは、HB の値が変わらない本数だけ。"""
        differ = []
        for twin, base in self.twins.items():
            for sel_b, sel_t in zip(_selections(self.master, base), _selections(self.master, twin)):
                w = self.master.width_of(min(c.ob_idx for c in self.master.all() if c.base_name == base))
                for n in range(1, 21):
                    fb = TC.judge_hb_flag(float(n), sel_b[1], sel_b[2])
                    ft = TC.judge_hb_flag_by_width(float(n), w)
                    if fb != ft:
                        hb_b = TC.calc_hb(self.table, float(n), w, "TIP1000", sel_b[1], sel_b[2]).HBValue
                        hb_t = TC.calc_hb(self.table, float(n), w, "TIP1000", sel_t[1], sel_t[2]).HBValue
                        self.assertEqual(hb_b, hb_t, (base, sel_b[0], n))
                        self.assertNotEqual(w, 53.5, "53.5mm では旗も同じ")
                        differ.append((base, n))
        self.assertTrue(differ, "VBA の旗は 40mm 14 本など結果に効かない本数でも立つ")

    def test_HB590とHB530は巾535の10本11本で効く(self):
        """53.5mm の双子(足したサイズ)でも、10 本・11 本で HB590 / HB530 の式になる。"""
        twin = next(t for t, b in self.twins.items() if b == "1.0mm×53.5mm")
        for sel in _selections(self.master, twin):
            for n, board in ((11, "HB590"), (10, "HB530"), (9, "外周")):
                with self.subTest(sel=sel[0], n=n):
                    d = _calc(self.table, n, 53.5, "TIP1000", sel)
                    self.assertIn(board, d["formulas"]["HBsiki"])
                    self.assertGreater(d["HB"], 0)

    def test_GWは風袋とNWの和_NWは1条重量と本数(self):
        twin = next(t for t, b in self.twins.items() if b == "1.0mm×40.0mm")
        for sel in _selections(self.master, twin):
            d = _calc(self.table, 14, 40.0, "TIP950", sel)
            t = T1 if sel[3] in (1, 2) else T2
            self.assertAlmostEqual(d["NW"], round(14 * t, 1), places=9)
            parts = ("PIN", "TIP", "SUT", "ESA", "POR", "HB", "PET", "PAR", "DRY")
            self.assertAlmostEqual(d["HU"], sum(d[k] for k in parts), places=9)
            self.assertAlmostEqual(d["GW"], d["HU"] + d["NW"], places=9)
            self.assertGreater(d["GW"], d["NW"])


class NewWidthCalcTest(unittest.TestCase):
    """2. 新しい巾(45.0mm)は、巾を渡した計算そのものと同じ。HB は標準の式。"""

    def test_新しい巾(self):
        table = _table()
        m = SizeMaster()
        m.apply_extra([{"obIdx": 15, "baseName": "1.2mm×45mm", "widthMm": 45.0,
                        "kataban": "BJB7699900QR", "symbol": "", "note": ""}])
        for sel in _selections(m, "1.2mm×45.0mm"):
            for n in (1, 10, 11, 13, 20):
                with self.subTest(sel=sel[0], n=n):
                    d = _calc(table, n, 45.0, "TIP1000", sel)
                    # 「同梱でない・HB の本数指定なし」の計算(cb_key・台紙番号を渡さない)と同じ
                    plain = _calc(table, n, 45.0, "TIP1000", ("", "", 0, sel[3], sel[4], sel[5]))
                    self.assertEqual(d, plain)
                    self.assertIn("外周", d["formulas"]["HBsiki"], "45mm は標準の式")
                    self.assertAlmostEqual(d["TA"], TC.get_coil_ta(float(n), 45.0), places=6)
                    self.assertGreater(d["GW"], d["NW"])


# ======================================================================
# 3・4. 画面と同じ道で(API)
# ======================================================================
def _master_db(path: str) -> None:
    """同梱の資材重量・型番を入れた共有マスタ(SQLite)。"""
    cfg = Config()
    conn = sqlite3.connect(path)
    conn.execute('CREATE TABLE "資材重量" ("管理番号" TEXT, "梱包資材名" TEXT, "単位質量" TEXT, "係数" TEXT)')
    for r in csv_text.read_dicts(cfg.default_material_csv_path):
        conn.execute('INSERT INTO "資材重量" VALUES (?,?,?,?)',
                     (r.get("管理番号"), r.get("梱包資材名"), r.get("単位質量"), r.get("係数")))
    conn.execute('CREATE TABLE "型番" ("サイズ名" TEXT, "記号" TEXT, "型番" TEXT, "備考" TEXT)')
    for r in csv_text.read_dicts(cfg.default_kataban_csv_path):
        conn.execute('INSERT INTO "型番" VALUES (?,?,?,?)',
                     (r.get("サイズ名"), r.get("記号"), r.get("型番"), r.get("備考")))
    conn.commit()
    conn.close()


def read_code39(svg: str) -> str:
    """印刷ビューの SVG のバーを測って Code 39 を読む(描いた側の文字列は使わない)。

    細と太はバー・すき間の幅で分け(最小幅の 1.6 倍より太いもの)、1 文字 = バー 5 本 +
    すき間 4 つ(うち太が 3 つ)、文字のあいだに細いすき間 1 つ、を確かめながら読む。
    """
    rects = [(float(x), float(w)) for x, w in
             re.findall(r'<rect x="([0-9.]+)" y="0" width="([0-9.]+)" height="[0-9.]+" fill="#000"/>', svg)]
    if not rects:
        raise AssertionError("バーがありません")
    narrow = min(w for _, w in rects)
    spaces = [rects[i + 1][0] - (rects[i][0] + rects[i][1]) for i in range(len(rects) - 1)]
    narrow = min([narrow] + spaces)

    def nw(width):
        return "w" if width > narrow * 1.6 else "n"
    inverse = {v: k for k, v in BC._PATTERNS.items()}
    if len(rects) % 5:
        raise AssertionError("バーの本数が 5 の倍数ではありません: %d" % len(rects))
    out = []
    for k in range(len(rects) // 5):
        pat = ""
        for j in range(5):
            i = k * 5 + j
            pat += nw(rects[i][1])
            if j < 4:
                pat += nw(spaces[i])
        if pat.count("w") != 3:
            raise AssertionError("1 文字に太が 3 つではありません: %s" % pat)
        if k * 5 + 4 < len(spaces) and nw(spaces[k * 5 + 4]) != "n":
            raise AssertionError("文字のあいだのすき間が細くありません")
        out.append(inverse[pat])
    text = "".join(out)
    if not (text.startswith("*") and text.endswith("*")):
        raise AssertionError("スタート/ストップ(*)がありません: %s" % text)
    return text[1:-1]


def barcodes(html: str):
    """印刷ビューのバーコード(読んだ中身, aria-label, 下の文字)。"""
    out = []
    for m in re.finditer(r'<svg class="bc39".*?</svg>', html, re.S):
        svg = m.group(0)
        label = re.search(r'aria-label="([^"]*)"', svg).group(1)
        text = re.search(r'<text [^>]*>([^<]*)</text>', svg)
        out.append((read_code39(svg), label, text.group(1) if text else ""))
    return out


class ApiFlowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._local_dir = Config.__dict__["local_dir"]
        cls.tmp = tempfile.mkdtemp()
        cls.db = os.path.join(cls.tmp, "梱包資材マスタ.sqlite3")
        _master_db(cls.db)
        cfg = Config()
        cfg.port = free_port()
        cfg.prefer_access = False
        cfg.material_csv_file = business_csv()   # 同梱サンプルの写しを指定した CSV として
        cfg.material_db_file = cls.db
        cfg.barcode_mode = "svg"
        cfg.__class__.local_dir = property(lambda self, d=cls.tmp: d)
        cls.cfg = cfg
        cls.httpd, cls.ctx = server.create_server(cfg)
        cls.thread = threading.Thread(target=cls.httpd.serve_forever,
                                      kwargs={"poll_interval": 0.1}, daemon=True)
        cls.thread.start()
        cls.base = "http://127.0.0.1:%d" % cfg.port
        for t, w, k in (("1.2", "53.5", "BJB7699500QR"), ("1.2", "45", "BJB7699900QR")):
            st, j = cls.post(cls, "/api/settings/size-add",
                             {"thickness": t, "width": w, "kataban": k, "password": S.PATH_PASSWORD})
            assert j["ok"], j

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()
        cls.thread.join(timeout=5)
        cls.ctx.store.close()
        Config.local_dir = cls._local_dir
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def post(self, path, payload):
        req = urllib.request.Request(self.base + path, data=json.dumps(payload).encode("utf-8"),
                                     method="POST", headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return r.status, json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            return exc.code, json.loads(exc.read().decode("utf-8"))

    def get(self, path):
        with urllib.request.urlopen(self.base + path, timeout=30) as r:
            return r.read().decode("utf-8")

    # ------------------------------------------------------------
    def _run(self, base_name: str, mode: str, coils, tip="TIP1000"):
        """画面と同じ順に API を呼ぶ。返すのは 重量計算_DB の結果とラベルへ写った値。"""
        master = self.ctx.wf.master
        cfgs = sorted((c for c in master.all() if c.base_name == base_name), key=lambda c: c.ob_idx)
        cur = {"kensaNo": "W111111", "weight1": str(T1), "weight2": str(T2), "tip": tip,
               "coilH": {}}
        if mode == "同時":
            if cfgs[0].cb_key == SM.NAMED_CB_KEY:
                _, j = self.post("/api/select-checkbox", {"named": True})
                cur["namedCb"] = True
            else:
                cb = int(cfgs[0].cb_key.replace("CheckBox", ""))
                _, j = self.post("/api/select-checkbox", {"cbIdx": cb})
                cur["selectedCb"] = cb
            obs = (cfgs[0].ob_idx, cfgs[1].ob_idx)
        else:
            ob = cfgs[0].ob_idx if mode == "丈1" else cfgs[1].ob_idx
            _, j = self.post("/api/select-size", {"obIdx": ob})
            cur["selectedOb"] = ob
            obs = (ob,)
        self.assertTrue(j["ok"], j)
        _, j = self.post("/api/apply-weight", {"current": cur})
        self.assertTrue(j["ok"], j)
        cur["coilH"] = {str(k): str(v) for k, v in coils.items()}
        _, j = self.post("/api/calc-tare", {"current": cur})
        self.assertTrue(j["ok"], j)
        self.assertEqual(j["missing"], [], "資材マスタに無い項目が無い")
        tare = {k: [(r["name"], r["value"], r["formula"]) for r in v["rows"]]
                for k, v in j["tare"].items()}
        labels = {}
        for ob in obs:
            view = self.ctx.wf.label_view(ob)
            labels[ob % 2] = {k: view["printHeader"][k] for k in
                              ("coilH1", "ta1", "nw1", "gw1", "coilH2", "ta2", "nw2", "gw2")}
        return tare, labels

    def test_3_画面の道で双子は同梱と同じ_GWとHBも(self):
        cases = [("丈1", {1: 11, 2: 10}), ("丈2", {3: 11, 4: 9}), ("同時", {1: 11, 2: 10, 3: 10, 4: 7})]
        for mode, coils in cases:
            with self.subTest(mode=mode):
                a = self._run("1.0mm×53.5mm", mode, coils)
                b = self._run("1.2mm×53.5mm", mode, coils)
                self.assertEqual(a, b)
                tare, labels = b
                hb = {k: dict((n, f) for n, _, f in rows) for k, rows in tare.items()}
                for k, rows in hb.items():
                    n = coils[int(k)]
                    want = {11: "HB590", 10: "HB530"}.get(n, "外周")
                    self.assertTrue(any(want in f for f in rows.values()), (mode, k, n))
                gw = [labels[o]["gw1"] for o in labels]
                self.assertTrue(all(g not in ("", "0", "0.0") for g in gw), labels)

    def test_3_新しい巾45も画面の道で計算できる(self):
        tare, labels = self._run("1.2mm×45.0mm", "同時", {1: 13, 2: 12, 3: 13, 4: 12})
        for k, rows in tare.items():
            vals = {n: v for n, v, _ in rows}
            self.assertTrue(vals, k)
        self.assertTrue(all(labels[o]["gw1"] for o in labels), labels)

    def test_4_バーコードは型番を表す_足したサイズも同梱も(self):
        for base, kata in (("1.2mm×45.0mm", "BJB7699900QR"), ("1.0mm×53.5mm", "BJB7606500QR")):
            with self.subTest(base=base):
                self._run(base, "同時", {1: 11, 2: 10, 3: 11, 4: 10})
                obs = [c.ob_idx for c in self.ctx.wf.master.all() if c.base_name == base]
                html = self.get("/labels/print?ob=%s" % ",".join(map(str, obs)))
                codes = barcodes(html)
                kata_codes = [c for c in codes if c[0] == kata]
                self.assertTrue(kata_codes, "型番のバーコードが無い: %s" % codes[:3])
                for read, label, text in kata_codes:
                    self.assertEqual(label, kata)
                    self.assertEqual(text, "*%s*" % kata, "下の文字は * で挟む")
                kensa = [c for c in codes if c[0].startswith("W111111")]
                self.assertTrue(kensa, "検査番号のバーコードも読める")

    def test_4_型番を直すと刷るバーコードも変わる(self):
        # マスタ管理で、足したサイズの型番を直す(合言葉付き)
        _, j = self.post("/api/master/rows", {"table": "ペナラベル追加サイズ", "keyword": "45.0"})
        row = j["rows"][0]
        _, j = self.post("/api/master/save", {"table": "ペナラベル追加サイズ", "rowId": row["__行"],
                                              "values": {"型番": "BJB7699901QR"},
                                              "password": S.PATH_PASSWORD})
        self.assertTrue(j["ok"], j)
        try:
            self._run("1.2mm×45.0mm", "丈1", {1: 13, 2: 12})
            ob = min(c.ob_idx for c in self.ctx.wf.master.all() if c.base_name == "1.2mm×45.0mm")
            codes = [c[0] for c in barcodes(self.get("/labels/print?ob=%d" % ob))]
            self.assertIn("BJB7699901QR", codes)
            self.assertNotIn("BJB7699900QR", codes)
        finally:
            self.post("/api/master/save", {"table": "ペナラベル追加サイズ", "rowId": row["__行"],
                                           "values": {"型番": "BJB7699900QR"},
                                           "password": S.PATH_PASSWORD})

    def test_4_型番マスタに足した型番が再起動なしで全サイズに出てバーコードになる(self):
        name = "1.2×45.0×Coil"
        self.assertNotIn(name, self.get("/all-size"))
        conn = sqlite3.connect(self.db)
        conn.execute('INSERT INTO "型番" VALUES (?,?,?,?)', (name, "h", "BJB7699977QR", ""))
        conn.commit()
        conn.close()
        os.utime(self.db, None)
        page = self.get("/all-size")
        self.assertIn(name, page, "足した型番が一覧に出る(出すサイズにも入っている)")
        _, j = self.post("/api/all-size/prepare", {"comboText": name})
        self.assertTrue(j["ok"], j)
        _, j = self.post("/api/all-size/submit", {"comboText": name, "kensaNo": "W222222",
                                                 "threeDigit": False, "w1": "12.5", "w2": "11.5"})
        self.assertTrue(j["ok"], j)
        html = self.get("/all-size/print?combo=" + urllib.parse.quote(name))
        codes = [c[0] for c in barcodes(html)]
        self.assertIn("BJB7699977QR", codes)
        self.assertTrue(any(c.startswith("W222222") for c in codes), codes[:4])


class Code39ReaderTest(unittest.TestCase):
    """読み戻しの道具そのものを確かめる(規格の 0 と * の並び)。"""

    def test_規格の並び(self):
        # Code 39 の規格: 0 = 101001101101, * = 100101101101(1=黒, 0=白)
        self.assertEqual(BC._PATTERNS["0"], "nnnwwnwnn")
        self.assertEqual(BC._PATTERNS["*"], "nwnnwnwnn")
        self.assertTrue(all(p.count("w") == 3 for p in BC._PATTERNS.values()))
        self.assertEqual(len(set(BC._PATTERNS.values())), len(BC._PATTERNS), "並びに重なりが無い")

    def test_読み戻せる(self):
        for text in ("BJB7699900QR", "W111111-0101 10", "BJB7610400QR"):
            with self.subTest(text=text):
                self.assertEqual(read_code39(BC.svg(text, 40, 8)), text)


if __name__ == "__main__":
    unittest.main()
