"""デスクトップ版の起動確認(scripts/desktop_smoke.py)の子孫の数え方(統合 1.2.30)

親の番号は親が終わってもそのまま残る。終わった親の番号が使い回されて exe と同じになっても、
親より前に起動したもの(Windows の wininit → services → svchost …)を exe の子孫と読み違えない。
"""
from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location("desktop_smoke", ROOT / "scripts" / "desktop_smoke.py")
smoke = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(smoke)


class ChildrenTest(unittest.TestCase):
    def test_番号を使い回した昔の子は数えない(self):
        procs = {
            1340: (500, "200", "CoilPackingTools.exe"),
            840: (1340, "100", "wininit.exe"),        # 昔の 1340(終わった)の子
            936: (840, "110", "services.exe"),
            2000: (1340, "210", "python.exe"),
            2100: (2000, "220", "python.exe"),
        }
        self.assertEqual(sorted(smoke.children(1340, procs)), [2000, 2100])

    def test_起動した時刻が分からなければ親の番号だけで数える(self):
        procs = {10: (1, "0", "exe"), 11: (10, "", "child"), 12: (11, "0", "grandchild")}
        self.assertEqual(sorted(smoke.children(10, procs)), [11, 12])


if __name__ == "__main__":
    unittest.main()
