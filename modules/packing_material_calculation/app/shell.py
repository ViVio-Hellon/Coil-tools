"""画面の外枠 ── どの画面にも共通の帯とレール

画面ごとの中身はテンプレートが持ち、ここは**どの画面が在るか**と、
帯に出す事実(版・取り込み日時・担当者)だけを組む。
"""
from __future__ import annotations

from dataclasses import dataclass

from modules.packing_material_calculation.coil_tool import app_config, config, data_sync, staff, user_settings


@dataclass(frozen=True)
class Page:
    key: str
    path: str
    label: str
    detail: str


# レールに並ぶ画面。**この並びが業務の順番**でもある
PAGES: tuple[Page, ...] = (
    Page("calc", "/calc", "資材計算",
         "ロットから資材の員数を出す"),
    Page("checklist", "/checklist", "チェックリスト",
         "計算結果を15行に積む"),
    Page("order", "/order", "発注票",
         "チェックリストから倉庫へ出す票を作る"),
    Page("settings", "/settings", "設定",
         "取り込み元の置き場所・マスタ管理・ライン・担当者"),
)


def page(key: str) -> Page:
    for p in PAGES:
        if p.key == key:
            return p
    raise ValueError(f"未知の画面: {key!r}")


def ribbon(conn) -> dict:
    """帯に出す事実。

    **取り込み日時をいつも出す。** マスタは取り込んだ時点の写しなので、
    古いまま計算していると気づけない(基盤仕様書の言う「古いデータを
    保持したままにならないか」への答え)。

    【開くたびに取り込み元を確かめる(統合 1.2.49・現場の指摘)】
    帯を描くとき(画面を開いたとき・帯の問い合わせ)に、取り込み元の目印を軽く確かめる
    (`data_sync.freshness`。ファイルの大きさ・更新時刻を見るだけ)。梱包資材マスタ・コイル実績が
    変わっていればその場で取り込み直し、仕掛台帳が変わっていれば「いま取り込む」を出す。
    以前は起動時に1回だけで、動かしている間に共有のマスタが直っても計算は古いままだった。
    **担当者の名簿より先に確かめる**(班員名簿も取り込み直すことがあるため)。
    """
    fresh = data_sync.freshness(conn)
    imported = data_sync.imported_at(conn)
    worker = user_settings.get_worker()
    return {
        "display_name": app_config.display_name(),
        "version_label": app_config.version_label(),
        "line": user_settings.get_line(),
        "lines": list(config.LINES),
        "worker": worker,
        # 担当者は梱包資材マスタの「班員名簿」から、**担当ラインが
        # 機側の人だけ**。名簿が取り込めていなければ予備の一覧へ落ちる
        # (担当者が無いとチェックリストへ積めないので、名簿が読めない
        # だけで仕事を止めない)
        "workers": staff.choices(conn, current=worker),
        # 写した日時(いちばん古い取り込み)と、取り込み元を確かめた日時は別に出す(統合 1.2.49)
        "imported_at": imported,
        "checked_at": data_sync.checked_at(conn),
        # 「古い」は取り込み元が変わったか(目印が一致しないか)で見る。届かず確かめられない
        # ときだけ日付で見る(統合 1.2.49)。取り込み中は古いとは言わない(別の札を出す)
        "stale": fresh.reason not in (data_sync.STALE_NONE, data_sync.STALE_IMPORTING),
        # **なぜ古いのか**も出す。20分前に取り込んだ人が日付を
        # またいだだけで「古い」と言われると、壊れたように見える
        "stale_why": fresh.message if fresh.reason != data_sync.STALE_IMPORTING else "",
        # 取り込み中(起動時の自動取り込みが長引いたときなど)。計算Start は断る(統合 1.2.49)
        "importing": fresh.importing,
        # 共有の仕掛台帳が変わった(大きいので自動では取り込まない。帯の「いま取り込む」で)
        "lot_changed": fresh.lot_changed,
        "lot_note": LOT_CHANGED_NOTE if fresh.lot_changed else "",
        # その場で取り込み直したこと(梱包資材マスタ・コイル実績)
        "refreshed": fresh.refreshed_note,
    }


#: 帯に出す、仕掛台帳が変わったときの一言(統合 1.2.49)
LOT_CHANGED_NOTE = "共有の仕掛台帳が更新されています"
