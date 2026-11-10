"""ツールのフォルダの中に書かない(統合 1.2.64・業務ツール統合ランチャー 1.8.9「更新のため全 PC で止める」)

ランチャーは更新の前に全 PC でツールを止め、共有サーバーのツールのフォルダを**丸ごと上書き**する。
データベース・現場で変えた設定・ログ・書き出したファイルをフォルダの中に書くと、更新で消える・
書きかけのファイルが上書きを止める・共有では何台もの PC が同じファイルに書く。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from common import app_config, logging_utils

ROOT = Path(__file__).resolve().parent.parent


class InsideAppTest(unittest.TestCase):
    def test_ツールのフォルダの中かを見分ける(self):
        self.assertTrue(app_config.inside_app(app_config.APP_ROOT))
        self.assertTrue(app_config.inside_app(app_config.APP_ROOT / "export" / "x"))
        self.assertFalse(app_config.inside_app(Path(tempfile.gettempdir()) / "x"))
        self.assertFalse(app_config.inside_app(str(app_config.APP_ROOT) + "_old"), "隣の似た名前は外")

    @unittest.skipIf(os.name == "nt", "リンクを作るには権限が要る(Windows)")
    def test_別の名前で届く道でもツールのフォルダの中だと見抜く(self):
        """統合 1.2.69。割り当てたドライブ・短い名前・リンクで同じフォルダに届くと、外に見えていた。"""
        tmp = Path(tempfile.mkdtemp(prefix="cpt-alias-"))
        self.addCleanup(shutil.rmtree, tmp, True)
        alias = tmp / "別名"
        os.symlink(app_config.APP_ROOT, alias)
        self.assertTrue(app_config.inside_app(alias / "logs"))
        self.assertTrue(app_config.inside_app(alias))
        self.assertFalse(app_config.inside_app(tmp / "外"))

    def test_ログの出力先にツールのフォルダの中は書けないと言う(self):
        problem = logging_utils.writable_problem(app_config.APP_ROOT / "logs_here")
        self.assertIn("ツールのフォルダの中", problem)
        self.assertFalse((app_config.APP_ROOT / "logs_here").exists(), "確かめるために作らない")


class DetailsExportTest(unittest.TestCase):
    """梱包明細の CSV の出力先の既定は、この PC の領域(1.2.63 まではツールのフォルダの中)。"""

    def test_既定はツールのフォルダの外(self):
        from modules.packing_details.meisai import config
        if os.environ.get("PACKING_DETAILS_EXPORT_DIR"):
            self.skipTest("環境変数で変えてある")
        self.assertFalse(app_config.inside_app(config.EXPORT_DIR), config.EXPORT_DIR)

    def test_前の場所のCSVを初回に写す(self):
        from modules.packing_details.meisai import config
        tmp = Path(tempfile.mkdtemp(prefix="cpt-export-"))
        self.addCleanup(shutil.rmtree, tmp, True)
        old, new = tmp / "old", tmp / "new"
        old.mkdir()
        (old / "a.csv").write_text("x", encoding="utf-8")
        new.mkdir()
        (new / "b.csv").write_text("keep", encoding="utf-8")
        (old / "b.csv").write_text("old", encoding="utf-8")
        with mock.patch.object(config, "LEGACY_EXPORT_DIR", old), \
                mock.patch.object(config, "EXPORT_DIR", new), \
                mock.patch.object(config, "_legacy_checked", False):
            self.assertEqual(config.carry_legacy_exports(), ["a.csv"])
            self.assertEqual(config.carry_legacy_exports(), [], "1度だけ")
        self.assertEqual((new / "a.csv").read_text(encoding="utf-8"), "x")
        self.assertEqual((new / "b.csv").read_text(encoding="utf-8"), "keep", "同じ名前は上書きしない")
        self.assertTrue((old / "a.csv").exists(), "前の場所のものは消さない")

    def test_資材計算の書き出し先の既定も外(self):
        from modules.packing_material_calculation.coil_tool import config
        if os.environ.get("COIL_TOOL_EXPORT_DIR"):
            self.skipTest("環境変数で変えてある")
        self.assertFalse(app_config.inside_app(config.DEFAULT_EXPORT_DIR))


class PenaSettingsTest(unittest.TestCase):
    def test_記録の保存先にツールのフォルダの中を入れると保存しない(self):
        from modules.packing_pena_label.app.services import settings as S
        errors = S.validate_all({"record_dir": str(app_config.APP_ROOT / "records")})
        self.assertIn("ツールのフォルダの中", errors.get("record_dir", ""))
        self.assertEqual(S.validate_all({"record_dir": tempfile.gettempdir()}), {})


class PycacheTest(unittest.TestCase):
    """デスクトップ版の入口(bridge.py)を読んでも、ツールのフォルダに __pycache__ を作らない。"""

    def test_bridgeは読み込む前に書かないと決める(self):
        text = (ROOT / "bridge.py").read_text(encoding="utf-8")
        head = text[:text.index("import threading")]
        self.assertIn("sys.dont_write_bytecode = True", head)

    def test_bridgeを読み込んでもツールのフォルダにpycacheができない(self):
        tmp = Path(tempfile.mkdtemp(prefix="cpt-pyc-"))
        self.addCleanup(shutil.rmtree, tmp, True)
        copy = tmp / "app"
        shutil.copytree(ROOT, copy, ignore=shutil.ignore_patterns(
            "__pycache__", ".git", "node_modules", "target", "tests", "tools", "*.png"))
        env = {k: v for k, v in os.environ.items() if k not in ("PYTHONDONTWRITEBYTECODE", "PYTHONPYCACHEPREFIX")}
        env["COIL_PACKING_TOOLS_LOCAL_DIR"] = str(tmp / "local")
        # デスクトップ版は `python bridge.py`(台本として実行)。台本そのものは __pycache__ を作らないので、
        # 同じく台本として読み(__main__ ではない名前で。待ち受けは始めない)、そこから先の読み込みを見る
        # main() が最初に読み込むもの(start_app・common)まで読む。main の待ち受けは始めない
        code = ("import runpy; runpy.run_path('bridge.py', run_name='bridge_check'); "
                "import start_app; from common import app_config, logging_utils")
        subprocess.run([sys.executable, "-c", code], cwd=str(copy), env=env,
                       check=True, timeout=120, capture_output=True)
        made = [str(p.relative_to(copy)) for p in copy.rglob("__pycache__")]
        self.assertEqual(made, [])

    def test_pycacheの置き場所を作れなければ書かない(self):
        import start_app
        with mock.patch.object(sys, "pycache_prefix", None), \
                mock.patch.object(sys, "dont_write_bytecode", False), \
                mock.patch.object(Path, "mkdir", side_effect=OSError(13, "denied")):
            start_app.redirect_pycache(Path(tempfile.gettempdir()) / "nope")
            self.assertTrue(sys.dont_write_bytecode, "ツールのフォルダへ書かない")


if __name__ == "__main__":
    unittest.main()
