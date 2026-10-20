# -*- coding: utf-8 -*-
"""風袋計算の結果を固定する(統合 1.2.26・構造点検の見張り)

同梱のサイズ × 丈1・丈2・同時(2通り) × TIP 3 種 × 1〜50 本 の **全部** について、
資材 9 項目・HB・NW・風袋・GW・高さ・式の文字・画面に出す値(`Workflow._tare_payload`)を
まとめた指紋を `golden/tare_golden.json` と突き合わせる。

- 計算は VBA と照らし済み(test_tare_calc ほか)。ここは **いまの結果が変わっていないこと** だけを見る。
  整理(リファクタリング)や不具合の修正で、意図せず数・丸め・式の文字が変わったら落ちる
- 意図して変えたときだけ、作り直す:
      python -m modules.packing_pena_label.tests.test_tare_golden --regen
  (作り直した差分は、変えた理由と一緒に変更履歴へ書く)
"""
import hashlib
import json
import sys
import unittest
from pathlib import Path

from modules.packing_pena_label.app.models import TIP_CHOICES
from modules.packing_pena_label.app.services import size_master as SM
from modules.packing_pena_label.app.services import tare_calc as TC
from modules.packing_pena_label.app.services.size_master import SizeMaster
from modules.packing_pena_label.app.services.workflow import Workflow

GOLDEN = Path(__file__).with_name("golden") / "tare_golden.json"
T1, T2 = 101.35, 99.8          # 1条重量(小数2桁: 丸めの境目も通る)
MAX_COILS = 50


def _table():
    from modules.packing_pena_label.app.config import Config
    from modules.packing_pena_label.app.repositories.material_repo import MaterialRepository
    repo = MaterialRepository(accdb_dir="", csv_path=Config().default_material_csv_path,
                              prefer_access=False)
    return repo.load(use_cache=False)


def _selections(master: SizeMaster, base: str):
    """(呼び名, cb_key, ob_idx, coil_no, 丈1を使う, 丈2を使う)。画面の選び方と同じ4通り。"""
    cfgs = sorted((c for c in master.all() if c.base_name == base), key=lambda c: c.ob_idx)
    c1, c2 = cfgs[0], cfgs[1]
    return [("丈1", "", c1.ob_idx, 1, True, False),
            ("丈2", "", c2.ob_idx, 3, False, True),
            ("同時", c1.cb_key, 0, 1, True, True),
            ("同時2", c1.cb_key, 0, 3, True, True)]


def _digest(obj) -> str:
    text = json.dumps(obj, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:12]


def compute() -> dict:
    """いまのコードでの指紋。キーは「サイズ|選び方|TIP」、値は本数 1..50 の指紋の並び。"""
    table = _table()
    master = SizeMaster()
    bases = sorted({c.base_name for c in master.all()})
    out = {}
    for base in bases:
        width = master.width_of(min(c.ob_idx for c in master.all() if c.base_name == base))
        for label, cb, ob, coil_no, v1, v2 in _selections(master, base):
            for tip in TIP_CHOICES:
                prints = []
                for n in range(1, MAX_COILS + 1):
                    mat = TC.calculate_materials(table, float(n), coil_no, width, tip, cb, ob,
                                                 T1, T2, v1, v2, all_coil_empty=False)
                    payload = Workflow._tare_payload(None, coil_no, float(n), mat, T1, T2)
                    raw = {k: v for k, v in vars(mat).items()}
                    prints.append(_digest({"mat": raw, "payload": payload}))
                out[f"{base}|{label}|{tip}"] = prints
    return out


class TareGoldenTest(unittest.TestCase):
    def test_風袋計算の結果が固定した値から変わっていない(self):
        expected = json.loads(GOLDEN.read_text(encoding="utf-8"))
        actual = compute()
        self.assertEqual(sorted(actual), sorted(expected), "組み合わせの数が変わった")
        diffs = []
        for key, prints in expected.items():
            for n, (want, got) in enumerate(zip(prints, actual[key]), start=1):
                if want != got:
                    diffs.append(f"{key} {n}本")
        self.assertFalse(diffs, "結果が変わった組み合わせ(先頭20件): %s" % diffs[:20])

    def test_組み合わせを全部数えている(self):
        expected = json.loads(GOLDEN.read_text(encoding="utf-8"))
        bases = {k.split("|")[0] for k in expected}
        self.assertEqual(len(bases), len({c.base_name for c in SizeMaster().all()}))
        self.assertEqual(sum(len(v) for v in expected.values()),
                         len(bases) * 4 * len(TIP_CHOICES) * MAX_COILS)


if __name__ == "__main__":
    if "--regen" in sys.argv:
        GOLDEN.parent.mkdir(exist_ok=True)
        GOLDEN.write_text(json.dumps(compute(), ensure_ascii=False, indent=0, sort_keys=True) + "\n",
                          encoding="utf-8")
        print("作り直しました:", GOLDEN)
    else:
        unittest.main()
