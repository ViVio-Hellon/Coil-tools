"""直下の README の「文書の置き場所」(統合 1.2.30)

機能の README に書いてある `docs/…` は機能のフォルダから見た場所で、直下からは探しにくかった。
直下の README から全部の文書へたどれること・リンクが切れていないことを確かめる。
"""
from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
README = (ROOT / "README.md").read_text(encoding="utf-8")
LINKS = {m for m in re.findall(r"\]\(([^)#\s]+)\)", README) if not m.startswith(("http:", "https:"))}


class ReadmeDocsTest(unittest.TestCase):
    def test_リンクが切れていない(self):
        missing = sorted(link for link in LINKS if not (ROOT / link).exists())
        self.assertEqual(missing, [])

    def test_文書はどれも直下のREADMEからたどれる(self):
        books = [*(ROOT / "docs").glob("*.md"), ROOT / "docs" / "ツール文書" / "00_文書一覧.md",
                 *(ROOT / "modules").glob("*/README.md"), *(ROOT / "modules").glob("*/docs/**/*.md")]
        self.assertTrue(books)
        unlisted = sorted(str(p.relative_to(ROOT)) for p in books
                          if str(p.relative_to(ROOT)).replace("\\", "/") not in LINKS)
        self.assertEqual(unlisted, [], "直下の README の「文書の置き場所」に足してください")


if __name__ == "__main__":
    unittest.main()
