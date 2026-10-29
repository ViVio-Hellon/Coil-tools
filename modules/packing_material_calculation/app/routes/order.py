"""発注票 ── チェックリストから倉庫へ出す票を作る

【重複は止めずに知らせる】
同じロットが2週間以内に出ていたら確認を出す。続けるかどうかは人が決める
(VBA も「このまま重複して発注しますか？」と聞いていた)。
"""
from __future__ import annotations

from flask import Blueprint, current_app, jsonify, render_template, request

from modules.packing_material_calculation.coil_tool import (checklist_service, config, order_history,
                                                            order_sheet_service, reports, user_settings)
from modules.packing_material_calculation.coil_tool.logging_utils import get_logger

from .. import base, conf, get_db, shell

log = get_logger("app.routes.order")

bp = Blueprint("order", __name__)

# 組み立てた発注票を、確認のあいだだけ持っておく。
# **プロセスに1つ**(このアプリは1台を1人が使う)
# `committed` … いまの発注票を「発注済みとして記録」したか。**同じ票を2度記録しない**
# (統合 1.2.11。以前は押した後も画面が変わらず、もう一度押すと同じ発注が2重に残った)
# `checklist` … 票を作ったときのチェックリストの指紋(`checklist_service.fingerprint`)。
# 印刷・記録のときに今のチェックリストと照らす(統合 1.2.49・現場の指摘。以前は票を作ったあとに
# チェックリストを変えても、古い票をそのまま印刷・記録できた)
_pending: dict = {"sheets": [], "worker": "", "committed": 0, "checklist": ""}

#: 票を作ったあとにチェックリストが変わった(統合 1.2.49)
REASON_CHECKLIST_CHANGED = "checklist_changed"
CHECKLIST_CHANGED_MESSAGE = ("発注票を作ったあとにチェックリストが変わりました。"
                             "「発注票を作る」で作り直してください")


def discard_pending() -> None:
    """作ってある発注票を捨てる。チェックリストを積む・消す・直すときに呼ぶ(統合 1.2.49)。"""
    _pending.update(sheets=[], worker="", committed=0, checklist="")


def _outdated(conn) -> bool:
    """作ってある票が、今のチェックリストから作ったものでないか。そうなら票を捨てて True。"""
    if not _pending["sheets"]:
        return False
    if _pending["checklist"] == checklist_service.fingerprint(conn):
        return False
    log.info("発注票を作ったあとにチェックリストが変わったので、票を捨てました")
    discard_pending()
    return True


def _sheet_view(sheet) -> dict:
    return {
        "種類": sheet.種類, "角サイズ": sheet.角サイズ, "ロット": sheet.ロット,
        "提出日付": sheet.提出日付, "依頼者": sheet.依頼者,
        "rows": [{"長さ": r.長さ, "LotNo": r.LotNo, "数量": r.数量}
                 for r in sheet.rows],
    }


@bp.get("/order")
def page():
    conn = get_db()
    return render_template(
        "order.html",
        page=shell.page("order"),
        pages=shell.PAGES,
        ribbon=shell.ribbon(conn),
        token=conf()["TOKEN"],
    )


@bp.get("/api/order")
def current():
    notice = CHECKLIST_CHANGED_MESSAGE if _outdated(get_db()) else ""
    return jsonify(ok=True, sheets=[_sheet_view(s) for s in _pending["sheets"]],
                   worker=_pending["worker"], committed=_pending["committed"], notice=notice,
                   by_lot=by_lot_setting())


def by_lot_setting() -> bool:
    """「LOTで分ける」を最後に選んだか(資材計算 0.2.35)。"""
    return user_settings.get(config.KEY_ORDER_BY_LOT) is True


_MESSAGES = {
    order_sheet_service.REASON_EMPTY:
        "チェックリストに、発注票にできる行がありません"
        "(コイル間がリプラ / LVS の行だけが対象です)",
    order_sheet_service.REASON_TOO_MANY_SIZES:
        "発注リストにないサイズが多すぎます(書き足せるのは4件までです)",
}


@bp.post("/api/order/build")
def build():
    """チェックリストから発注票を組む。**履歴の重複は知らせるだけ。**"""
    conn = get_db()
    worker = user_settings.get_worker()
    if not worker:
        return jsonify(ok=False, reason="no_worker",
                       message="担当者を選んでください(帯の担当者欄)"), 422

    payload = request.get_json(silent=True) or {}
    check = not bool(payload.get("確認済み"))
    # 「LOTで分ける」(資材計算 0.2.35)。送られてこなければ前に選んだほう
    by_lot = bool(payload["LOTで分ける"]) if "LOTで分ける" in payload else by_lot_setting()
    if by_lot != by_lot_setting():
        user_settings.save(config.KEY_ORDER_BY_LOT, by_lot)

    stamp = checklist_service.fingerprint(conn)
    result = order_sheet_service.build(conn, worker=worker, check_history=check, by_lot=by_lot)
    if not result.ok:
        return jsonify(ok=False, reason=result.reason,
                       message=_MESSAGES.get(result.reason, "発注票を作れません"),
                       extra_lengths=result.extra_lengths), 422

    _pending["sheets"] = result.sheets
    _pending["worker"] = worker
    _pending["committed"] = 0
    _pending["checklist"] = stamp
    return jsonify(
        ok=True,
        sheets=[_sheet_view(s) for s in result.sheets],
        extra_lengths=result.extra_lengths,
        by_lot=by_lot,
        # 発注履歴を読めず、重複を確かめていない(統合 1.2.61)。画面が知らせる
        history_unchecked=result.history_unchecked,
        history_message=order_history.UNCHECKED_MESSAGE if result.history_unchecked else "",
        duplicates=[{
            "LotNo": d.LotNo, "処理日": d.処理日, "担当者": d.担当者,
            "サイズ": d.サイズ, "数量": d.数量, "日数表示": d.日数表示,
            "message": d.message(),
        } for d in result.duplicates],
    )


@bp.post("/api/order/commit")
def commit():
    """出した内容を履歴へ残す。これを押すまでは履歴に入らない。"""
    conn = get_db()
    if _outdated(conn):
        return jsonify(ok=False, reason=REASON_CHECKLIST_CHANGED,
                       message=CHECKLIST_CHANGED_MESSAGE), 409
    if not _pending["sheets"]:
        return jsonify(ok=False, reason="nothing_pending",
                       message="先に発注票を作ってください"), 422
    if _pending["committed"]:
        return jsonify(ok=False, reason="already_committed", committed=_pending["committed"],
                       message=f"この発注票は記録済みです({_pending['committed']} 件)。"
                               "もう一度記録するときは、発注票を作り直してください"), 409
    n = order_sheet_service.commit(conn, _pending["sheets"],
                                   worker=_pending["worker"])
    _pending["committed"] = n or -1
    return jsonify(ok=True, saved=n, committed=_pending["committed"])


@bp.get("/api/order/history")
def history():
    return jsonify(ok=True, rows=order_history.recent(get_db()))


@bp.get("/report/order")
def report():
    """印刷用の紙面。**1枚 = 1ページ。**"""
    from modules.packing_material_calculation.coil_tool import printing

    if _outdated(get_db()):
        return f"<p>{CHECKLIST_CHANGED_MESSAGE}</p>", 409
    if not _pending["sheets"]:
        return "<p>発注票がまだ作られていません。</p>", 404
    rep = reports.build_order_sheet_report(_pending["sheets"])
    # プレビューの上に「印刷する」(0.2.9。現場の指摘)
    return printing.render_html(rep, print_button=True)
