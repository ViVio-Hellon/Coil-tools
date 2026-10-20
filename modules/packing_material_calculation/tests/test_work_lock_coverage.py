"""作業状態の順番待ちが、作業状態に触る口を全部覆っているか(統合 1.2.26・構造点検の見張り)

作業状態(計算の入力と結果・チェックリスト・発注票の控え)はプロセスに1組で、書く要求は
`app/__init__.py` の `WORK_PREFIXES` で1つずつ通す(統合 1.2.20)。画面の側も同じ一覧
(`static/js/core.js` の `WORK_PATHS`)で、押した順に1つずつ送る。

**一覧は2か所に手で書いてある。** 作業状態に触る口を新しく足したのに一覧へ入れ忘れると、
二度押しで空の結果が返る不具合が戻る。ここで見張る:
  1. 作業状態(`session` / 発注票の控え `_pending`)に触る**書き込みの口**は、全部一覧の中にある
  2. 画面の一覧とサーバの一覧が同じ
"""
import re
from pathlib import Path

from modules.packing_material_calculation import app as app_module
from modules.packing_material_calculation.tests._web import client, sandbox   # noqa: F401

ROOT = Path(app_module.__file__).resolve().parent


def _touches_work_state(view) -> bool:
    """その口が作業状態に触るか。関数の中で使う名前(`session` / `_pending`)で見る。

    `admin_session`(マスタ管理の合言葉)は別の名前なので数えない。"""
    code = getattr(view, "__code__", None)
    if code is None:
        return False
    names = set(code.co_names)
    return "_pending" in names or "session" in names


def test_作業状態に触る書き込みの口は全部順番待ちの一覧にある(client):
    flask_app = client.application
    missing, covered = [], []
    for rule in flask_app.url_map.iter_rules():
        if not ({"POST", "PUT", "DELETE"} & set(rule.methods or ())):
            continue
        view = flask_app.view_functions.get(rule.endpoint)
        if not _touches_work_state(view):
            continue
        path = re.sub(r"^/[^/]+(?=/api/)", "", rule.rule)     # 統合の前置き(/material)を外す
        (covered if path.startswith(app_module.WORK_PREFIXES) else missing).append(rule.rule)
    assert covered, "作業状態に触る口が1つも見つからない(見分け方が壊れている)"
    assert not missing, f"順番待ちの一覧(WORK_PREFIXES)に無い: {missing}"


def test_画面の一覧とサーバの一覧が同じ():
    js = (ROOT / "static" / "js" / "core.js").read_text(encoding="utf-8")
    m = re.search(r"const WORK_PATHS = \[([^\]]*)\]", js)
    assert m, "core.js の WORK_PATHS が見つからない"
    paths = tuple(re.findall(r"'([^']+)'", m.group(1)))
    assert paths == tuple(app_module.WORK_PREFIXES)
