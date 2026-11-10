"""この PC の設定が読めないとき、管理者パスワードを既定の値で通さない(統合 1.2.69・構造点検)

「読めない」を「無い」と取り違えない(統合 1.2.61 の決まり)。資材計算の管理者パスワードは、変えた値を
この PC の設定ファイルに持つ。設定ファイルが読めない(壊れた・掴まれている)と「変えていない」とみなして
**既定のパスワードで通していた**。梱包明細は 1.2.61 で「読めなければ断る」に直していた。
"""
from __future__ import annotations

from modules.packing_material_calculation.coil_tool import admin_password, config, user_settings
from modules.packing_material_calculation.tests._web import client, sandbox   # noqa: F401

NEW = "kawatta-pass-1"


def _changed():
    result = admin_password.change(config.ADMIN_PASSWORD, NEW, NEW)
    assert result.ok, result.message
    assert admin_password.verify(NEW)


def test_設定が読めないときは既定のパスワードで通さない(sandbox):
    _changed()
    config.USER_CONFIG_PATH.write_text("{壊れている", encoding="utf-8")
    assert not admin_password.verify(config.ADMIN_PASSWORD), "既定のパスワードで通った"
    assert not admin_password.verify(NEW), "読めないあいだは、どの値でも通さない"


def test_設定が読めないあいだは変えられないし既定にも戻せない(sandbox):
    _changed()
    config.USER_CONFIG_PATH.write_text("{壊れている", encoding="utf-8")
    assert not admin_password.change(config.ADMIN_PASSWORD, "another-pass", "another-pass").ok
    assert not admin_password.reset(config.ADMIN_PASSWORD).ok


def test_設定ファイルが無ければこれまでどおり既定のパスワードで通る(sandbox):
    if config.USER_CONFIG_PATH.exists():
        config.USER_CONFIG_PATH.unlink()
    assert admin_password.verify(config.ADMIN_PASSWORD)


def test_読めるようになれば変えた値で通る(sandbox):
    _changed()
    saved = config.USER_CONFIG_PATH.read_text(encoding="utf-8")
    config.USER_CONFIG_PATH.write_text("{壊れている", encoding="utf-8")
    assert not admin_password.verify(NEW)
    config.USER_CONFIG_PATH.write_text(saved, encoding="utf-8")
    assert admin_password.verify(NEW)
    assert not admin_password.verify(config.ADMIN_PASSWORD)
