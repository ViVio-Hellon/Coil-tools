"""3つの app_config.py の中身がずれていかないことを見張る(統合 1.2.73・構造点検の改善案5)

統合アプリ(`common/app_config.py`)と、梱包明細・資材計算(`meisai/app_config.py`・`coil_tool/app_config.py`)に
同じ関数が3つずつある。違うのは**値**(設定ファイル・置き場所の名前・環境変数・モード)だけ。

1つにまとめるには「同じモジュールの関数どうしの呼び出しを試験で差し替える」書き方を崩す細工が要り、読みにくく
なるので、まとめずに**関数の中身が同じであること**をここで確かめる。片方だけ直すと、この試験が落ちる
(直すときは3つとも直す。値の違いは下の `ALLOWED` に書いたものだけ)。
"""
from __future__ import annotations

import ast
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FILES = {
    "統合": ROOT / "common" / "app_config.py",
    "梱包明細": ROOT / "modules" / "packing_details" / "meisai" / "app_config.py",
    "資材計算": ROOT / "modules" / "packing_material_calculation" / "coil_tool" / "app_config.py",
}
#: 値だけが違う(中身は比べない)関数。理由つき
ALLOWED = {
    "local_root": "置き場所を差し替える環境変数の名前が機能ごとに違う",
    "port": "既定のモードの名前が機能ごとに違う(modes.DEFAULT / modes.MAIN)",
    "port_candidates": "同上",
    "describe": "モードの持ち方が機能ごとに違う(名前の並び / モードの組)",
}


def _functions(path: Path) -> dict[str, str]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    out = {}
    for node in tree.body:
        if isinstance(node, ast.FunctionDef):
            body = node.body
            if body and isinstance(body[0], ast.Expr) and isinstance(getattr(body[0], "value", None), ast.Constant) \
                    and isinstance(body[0].value.value, str):
                body = body[1:]                  # 説明の文は比べない
            out[node.name] = ast.dump(ast.Module(body=body, type_ignores=[]))
    return out


class AppConfigCopiesTest(unittest.TestCase):
    def test_3つに共通の関数は中身が同じ(self):
        found = {name: _functions(path) for name, path in FILES.items()}
        common = set.intersection(*(set(f) for f in found.values()))
        self.assertGreater(len(common), 15, "比べる関数が少なすぎる(作りが変わった)")
        for func in sorted(common - set(ALLOWED)):
            with self.subTest(func=func):
                bodies = {name: f[func] for name, f in found.items()}
                self.assertEqual(len(set(bodies.values())), 1,
                                 f"{func} の中身が3つで違います。直すときは3つとも直してください")

    def test_機能の写しにある関数は統合にもある(self):
        main = set(_functions(FILES["統合"]))
        for name in ("梱包明細", "資材計算"):
            with self.subTest(name=name):
                self.assertEqual(set(_functions(FILES[name])) - main, set())


if __name__ == "__main__":
    unittest.main()
