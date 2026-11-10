"""管理者パスワードは梱包明細の共有のもの1つで確かめる(統合 1.2.75・構造点検の改善案4)

以前は3系統あった: 梱包明細(全ラインで共有)・資材計算(この PC の設定)・ペナラベル(決まった値)。
梱包明細で変えても資材計算は古いままで、資材計算は PC ごとに違う値になりえた。
"""
from __future__ import annotations

import json

from modules.packing_details.meisai import admin_password as shared
from modules.packing_details.meisai import shared_settings
from modules.packing_material_calculation.coil_tool import admin_password, config, user_settings
from modules.packing_material_calculation.tests._web import client, sandbox   # noqa: F401

NEW = "kawatta-pass-1"


def test_変えていなければ既定のパスワードで通る(sandbox):
    assert admin_password.verify(config.ADMIN_PASSWORD)
    assert not admin_password.verify(config.ADMIN_PASSWORD + "x")
    assert not admin_password.verify("")
    assert not admin_password.is_custom()


def test_資材計算で変えると梱包明細でも新しいほうだけ通る(sandbox):
    assert admin_password.change(config.ADMIN_PASSWORD, NEW, NEW).ok
    assert shared.verify(NEW) and admin_password.verify(NEW)
    assert not shared.verify(config.ADMIN_PASSWORD)
    assert not admin_password.verify(config.ADMIN_PASSWORD)
    assert admin_password.is_custom()


def test_梱包明細で変えると資材計算でも新しいほうだけ通る(sandbox):
    assert shared.change(config.ADMIN_PASSWORD, NEW, NEW).ok
    assert admin_password.verify(NEW)
    assert not admin_password.verify(config.ADMIN_PASSWORD)


def test_この_PC_の設定に前の値が残っていても読まない(sandbox):
    """切り替える前に PC ごとに持っていた値(撹拌済み)。もう照合に使わない。"""
    user_settings.save(admin_password.KEY, "pbkdf2$1$AAAA$BBBB")
    assert admin_password.verify(config.ADMIN_PASSWORD)
    assert not admin_password.is_custom()


def test_共有の控えが読めなければどの値でも通さない(sandbox):
    assert admin_password.change(config.ADMIN_PASSWORD, NEW, NEW).ok
    shared_settings.shared_path().write_text("{壊れている", encoding="utf-8")
    shared_settings.reset_for_tests()
    assert not admin_password.verify(config.ADMIN_PASSWORD), "既定のパスワードで通った"


def test_平文で持たない(sandbox):
    assert admin_password.change(config.ADMIN_PASSWORD, NEW, NEW).ok
    raw = shared_settings.shared_path().read_text(encoding="utf-8")
    assert NEW not in raw
    assert json.loads(raw)[shared.KEY].startswith("pbkdf2$")
    if config.USER_CONFIG_PATH.exists():
        assert NEW not in config.USER_CONFIG_PATH.read_text(encoding="utf-8")
