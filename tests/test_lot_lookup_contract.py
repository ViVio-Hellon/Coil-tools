"""梱包明細のロット引き(`meisai/lot_lookup.lookup`)の返す形(統合 1.2.73・構造点検の改善案6)

ペナラベル(LOTNO から検査番号・外径・チップボール)と資材計算(コイル外径)が、梱包明細の `lookup` の返す dict を
直接読む。どちらも失敗を飲み込んで画面を止めないので、形が変わると**画面は動いたまま結果だけ変わる**。
返す鍵を `RESULT_KEYS` に1か所で書き、どの返し方でもそろっていること・2つの機能が読む鍵がその中にあることを見る。
"""
from __future__ import annotations

import ast
import unittest
from pathlib import Path
from unittest import mock

from modules.packing_details.meisai import lot_lookup

ROOT = Path(__file__).resolve().parent.parent
#: 2つの機能の、`lookup` の返事を読むところ
READERS = (ROOT / "modules/packing_pena_label/app/services/coil_outer.py",
           ROOT / "modules/packing_pena_label/app/services/workflow.py",
           ROOT / "modules/packing_material_calculation/coil_tool/lot_service.py")
#: 読む側が自分で足す・ペナラベルの部品が付ける印(`lookup` の鍵ではない)
EXTRA = {"off", "source", "kind", "chip", "level", "message", "state", "detail"}


def _keys_read(path: Path, var_names=("hit", "first", "found")) -> set[str]:
    """`hit.get("x")` / `hit["x"]` で読んでいる鍵。"""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    out = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and node.func.attr == "get" \
                and isinstance(node.func.value, ast.Name) and node.func.value.id in var_names \
                and node.args and isinstance(node.args[0], ast.Constant):
            out.add(node.args[0].value)
        if isinstance(node, ast.Subscript) and isinstance(node.value, ast.Name) and node.value.id in var_names \
                and isinstance(node.slice, ast.Constant) and isinstance(node.slice.value, str):
            out.add(node.slice.value)
    return out


class LotLookupContractTest(unittest.TestCase):
    def test_空のLOTでも見つからなくても_鍵はそろう(self):
        self.assertEqual(set(lot_lookup.lookup("")), set(lot_lookup.RESULT_KEYS))
        with mock.patch.object(lot_lookup, "sources", return_value=[]):
            got = lot_lookup.lookup("Z999999")
        self.assertFalse(got["found"])
        self.assertTrue(set(lot_lookup.RESULT_KEYS) <= set(got), set(lot_lookup.RESULT_KEYS) - set(got))

    def test_2つの機能が読む鍵は_返す鍵の中にある(self):
        for path in READERS:
            with self.subTest(path=path.name):
                unknown = _keys_read(path) - set(lot_lookup.RESULT_KEYS) - EXTRA
                self.assertEqual(unknown, set(), f"{path.name} が lookup の返さない鍵を読んでいる")


if __name__ == "__main__":
    unittest.main()
