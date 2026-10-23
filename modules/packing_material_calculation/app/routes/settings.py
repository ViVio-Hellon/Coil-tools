"""設定 ── 取り込み元の置き場所、ライン、担当者

VBA はパスを標準モジュールの `Public Const` に直接書いていた。置き場所が
変わるとブックを開いてコードを直す必要があり、端末ごとに写しがずれる
原因になっていた。ここでは**画面から変えて、その場で届くか確かめられる**。

    打つ → 「確かめる」で到達を見る → 「保存」 → 「取り込み」

保存と取り込みを分けてあるのは、**打ち間違いで前の写しを壊さない**ため。
保存しただけでは手元のデータは変わらない。
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Optional

from flask import Blueprint, current_app, jsonify, render_template, request

from common import app_config as _integrated_app_config
from common import app_config as _common_app_config
from common import logging_utils as _common_logging
from common import storage_places
from common import ls4_source
from modules.packing_material_calculation.coil_tool import (admin_session, app_config, coil_results, config, db,
                       data_sync, distribution, source_db, staff, user_settings)
from modules.packing_material_calculation.coil_tool.presenters import master as master_presenter
from modules.packing_material_calculation.coil_tool.logging_utils import get_logger

from .. import base, conf, get_db, shell

log = get_logger("app.routes.settings")

bp = Blueprint("settings", __name__)

# 画面に出す欄。**この表が唯一の出どころ** ── 欄を足すときはここに1行。
# `key` は `user_settings` の鍵、`resolve` はいま実際に使っている場所
FIELDS: tuple[dict[str, Any], ...] = (
    {
        "key": config.KEY_MASTER_DB_DIR,
        "label": "梱包資材マスタのフォルダ",
        "detail": "包装仕様・パレット・リプラサイズを読む先",
        "vba": "PATH_AIM_参照",
        "expect": config.MATERIAL_DB_NAME,
        "default": str(config.DEFAULT_MASTER_DB_DIR),
    },
    {
        "key": config.KEY_LOT_DB_DIR,
        "label": "第1パス: 仕掛台帳のフォルダ",
        "detail": "ロット検索と受注情報を読む先",
        "vba": "PATH_仕掛_参照",
        "expect": ", ".join(config.LOT_DB_FILES.values()),
        "default": str(config.DEFAULT_LOT_DB_DIR),
    },
    {
        # 仕掛台帳の2つ目(統合 1.2.7 で名前も設定できるようにした。python-web-tools と同じ決まり)
        "key": config.KEY_LOT_DB_DIR_2,
        "label": "第2パス: 仕掛台帳のフォルダ",
        "detail": ("第1パスで見つからないときに見る先。第1パスに無いファイルは第2パスから読み、"
                   "第1パスに無いロット・受注は第2パスから足す(両方にあれば第1パス)。空でよい"),
        "vba": "PATH_梱包_仕掛",
        "expect": "下の「第2パスの名前」(空なら 第1パスと同じ "
                  + ", ".join(config.LOT_DB_FILES[t] for t in config.KEY_LOT_DB2_FILES) + ")",
        "default": "",
        "optional": True,
    },
    *({
        "key": key,
        "label": f"第2パスの名前: {Path(config.LOT_DB_FILES[table]).stem} の代わり",
        "detail": ("第2パスのフォルダで探すファイル名(ファイル名だけ。拡張子を省けば .sqlite3 / .db)。"
                   "空なら第1パスと同じ名前。名前は変えてよいが、中身の要る列はそろえる約束"),
        "vba": "",
        "expect": "",
        "default": config.LOT_DB_FILES[table],
        "optional": True,
        "kind": "file",
        "table": table,
    } for table, key in config.KEY_LOT_DB2_FILES.items()),
    {
        # 梱包課共有の仕掛(第1パスの LS4LOT)。統合 1.2.29。資材計算は**名前が LS4LOT 以外**
        # (実績の形式)のときだけ読み、コイル外径 = 内径 + 2 × 肉厚 に使う
        "key": config.KEY_KONPO_DB_DIR,
        "label": "第1パス: 梱包課共有の仕掛フォルダ(LS4LOT)",
        "detail": ("名前が LS4LOT 以外(実績の形式)のときだけ読み、コイル外径 = 内径 + 2 × 肉厚 にする。"
                   "LS4LOT のままなら読まない(受注の外径 MAX を使う)"),
        "vba": "",
        "expect": "下の「LS4LOT の名前」",
        "default": str(config.DEFAULT_KONPO_DB_DIR),
    },
    {
        "key": config.KEY_KONPO_DB_FILE,
        "label": "第1パス: LS4LOT の名前",
        "detail": ("梱包課共有の仕掛フォルダで探すファイル名。LS4LOT 以外の名前(例: 実績_LS4)なら実績の形式で読む"
                   "(いまの年月の行だけ。仕様は docs/仕掛ファイルの仕様.md)"),
        "vba": "",
        "expect": "",
        "default": config.KONPO_DB_FILE,
        "optional": True,
        "kind": "file",
        "table": "仕掛当工程",
        "where": "konpo",
    },
    {
        "key": config.KEY_LS4_DB2_FILE,
        "label": "第2パスの名前: LS4LOT の代わり",
        "detail": ("第2パス(仕掛台帳の第2パスのフォルダ)で探す名前。空なら第1パスと同じ名前。"
                   "LS4LOT 以外なら実績の形式で読み、第1パスに無いロットの外径を足す"),
        "vba": "",
        "expect": "",
        "default": "",
        "optional": True,
        "kind": "file",
        "table": "仕掛当工程",
    },
)
FIELD_BY_KEY = {f["key"]: f for f in FIELDS}

# 置き場所の**枠**(統合 1.2.30・現場の指摘: 同じ枠で考えるものを別々の欄で設定していて間違えやすい。
# 梱包明細の「置き場所・取り込み」と同じ形)。枠ごとに確かめる・保存して取り込み・既定に戻す。
# `tables` は保存したあとに取り込む表(その枠の分だけ)
GROUPS: tuple[dict[str, Any], ...] = (
    {"id": "master", "label": "梱包資材マスタのフォルダ",
     "lead": "包装仕様・パレット・リプラサイズ・班員名簿を読む先(VBA PATH_AIM_参照)。",
     "keys": (config.KEY_MASTER_DB_DIR,),
     "tables": ("包装仕様", "パレット", "リプラサイズ", "班員名簿")},
    {"id": "lot", "label": "第1パス: 仕掛台帳のフォルダ",
     "lead": "ロット検索と受注情報を読む先(VBA PATH_仕掛_参照)。ファイル名は決まっています。",
     "keys": (config.KEY_LOT_DB_DIR,),
     "tables": ("仕掛引当", "仕掛受注")},
    {"id": "konpo", "label": "第1パス: 梱包課共有の仕掛フォルダ(LS4LOT)",
     "lead": ("名前が LS4LOT 以外(実績の形式。例: 実績_LS4)のときだけ読み、コイル外径(内径 + 2 × 肉厚)と"
              "検入数(BOX実績_枚本数)に使います。LS4LOT のままなら読みません(仕様は docs/仕掛ファイルの仕様.md)。"),
     "keys": (config.KEY_KONPO_DB_DIR, config.KEY_KONPO_DB_FILE),
     "tables": ("コイル実績",)},
    {"id": "second", "label": "第2パス: 仕掛ファイルのフォルダ",
     "lead": ("第1パスで見つからないときに見る先(空なら使いません)。第1パスに無いファイルは第2パスから読み、"
              "第1パスに無いロット・受注は第2パスから足します(両方にあれば第1パス)。名前が空なら第1パスと同じ名前。"),
     "keys": (config.KEY_LOT_DB_DIR_2, *config.KEY_LOT_DB2_FILES.values(), config.KEY_LS4_DB2_FILE),
     "tables": ("仕掛引当", "仕掛受注", "コイル実績")},
)
GROUP_BY_ID = {g["id"]: g for g in GROUPS}


def _group_files(group_id: str) -> str:
    """その枠で探すファイル(画面に出す。**何を置けばよいか**が分からないと直せない)。"""
    if group_id == "master":
        return config.MATERIAL_DB_NAME
    if group_id == "lot":
        return " / ".join(config.LOT_DB_FILES[t] for t in config.KEY_LOT_DB2_FILES)
    if group_id == "konpo":
        return config.konpo_db_file()
    names = config.lot_db2_files()
    return " / ".join([names[t] for t in config.KEY_LOT_DB2_FILES] + [config.ls4_db2_file()])


def _group_view(group: dict[str, Any], imported: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": group["id"], "label": group["label"], "lead": group["lead"],
        "files": _group_files(group["id"]),
        "fields": [_field_view(FIELD_BY_KEY[k]) for k in group["keys"]],
        "tables": [{"table": t,
                    "at": (imported.get(t) or {}).get("取り込み日時", ""),
                    "rows": (imported.get(t) or {}).get("件数", ""),
                    "source": (imported.get(t) or {}).get("元ファイル", "")}
                   for t in group["tables"]],
    }


# 「書き出し先」(CSV を出す場所)の欄は外した(統合 1.2.22)。移植元から欄だけがあって、
# そこへ書き出す処理がどこにも無かった(文書を作るときに見つかった)。前の設定に値が残っていても使わない

# **置き場所はパスワードが要る。** ライン・担当者・自動取り込みは
# 使う人ごとの持ち物なので要らない。
#
# 分けるのは、効く範囲が違うから ── 置き場所を間違えると
# **その端末が計算できなくなる**(マスタを読めない)。担当者を間違えても
# 出てくる票の名前が違うだけで、気づいたら選び直せばよい。
PATH_KEYS = frozenset({
    config.KEY_MASTER_DB_DIR, config.KEY_LOT_DB_DIR,
    config.KEY_LOT_DB_DIR_2,
    *config.KEY_LOT_DB2_FILES.values(),
    config.KEY_KONPO_DB_DIR, config.KEY_KONPO_DB_FILE, config.KEY_LS4_DB2_FILE,
})

REFUSE_NEEDS_PASSWORD = "needs_password"
NEEDS_PASSWORD_MESSAGE = (
    "置き場所を変えるにはパスワードが要ります。"
    "設定のいちばん上でパスワードを入れてください。")


def _path_gate(keys) -> bool:
    """置き場所を触ってよいか。触らないなら聞かない。"""
    if not (PATH_KEYS & set(keys)):
        return True
    return admin_session.is_open()


_RESOLVERS = {
    config.KEY_MASTER_DB_DIR: config.master_db_dir,
    config.KEY_LOT_DB_DIR: config.lot_db_dir,
    config.KEY_LOT_DB_DIR_2: lambda: config.lot_db_dir_fallback() or "",
}


def _second_file_place(table: str) -> str:
    """第2パスのファイル名の欄の「いま見に行く場所」。"""
    if len(config.lot_db_dirs()) < 2:
        return "(第2パスのフォルダが空なので使いません)"
    found = data_sync.find_second_lot(table)
    if found is None:
        return f"{config.lot_db2_files()[table]}(第2パスのフォルダに見つかりません)"
    return str(found)


for _table, _key in config.KEY_LOT_DB2_FILES.items():
    _RESOLVERS[_key] = (lambda t: lambda: _second_file_place(t))(_table)


def _ls4_place(first: bool) -> str:
    """LS4LOT の名前の欄の「いま見に行く場所」(統合 1.2.29)。"""
    name = config.konpo_db_file() if first else config.ls4_db2_file()
    if not ls4_source.is_results_name(name):
        return f"{name}(名前が LS4LOT なので読みません。外径は受注の外径 MAX)"
    if first:
        folder = config.konpo_db_dir()
    else:
        dirs = config.lot_db_dirs()
        if len(dirs) < 2:
            return "(第2パスのフォルダが空なので使いません)"
        folder = dirs[1]
    found = source_db.find_named(folder, name)
    return str(found) if found else f"{name}({folder} に見つかりません)"


_RESOLVERS[config.KEY_KONPO_DB_DIR] = config.konpo_db_dir
_RESOLVERS[config.KEY_KONPO_DB_FILE] = lambda: _ls4_place(True)
_RESOLVERS[config.KEY_LS4_DB2_FILE] = lambda: _ls4_place(False)


# ==================================================================
# 画面
# ==================================================================
@bp.get("/settings")
def page():
    conn = get_db()
    return render_template(
        "settings.html",
        page=shell.page("settings"),
        pages=shell.PAGES,
        ribbon=shell.ribbon(conn),
        # マスタ管理の面。**共有フォルダには触らない分だけ**渡して、
        # 中身は画面が開いてから読む
        master=master_presenter.to_dict(master_presenter.frame(conn)),
        token=conf()["TOKEN"],
    )


# ==================================================================
# いまの設定
# ==================================================================
def _field_view(field: dict[str, Any]) -> dict[str, Any]:
    """1欄ぶんの表示用データ。

    **「打った値」と「実際に見に行く場所」を両方出す。** 相対で打たれた
    ときにどこを指すのかが分からないと、届かない理由を追えない。
    """
    key = field["key"]
    raw = user_settings.get(key)
    raw = raw if isinstance(raw, str) else ""
    resolved = _RESOLVERS[key]()
    return {
        "key": key,
        "label": field["label"],
        "detail": field["detail"],
        "vba": field.get("vba", ""),
        "expect": field.get("expect", ""),
        "optional": bool(field.get("optional")),
        "value": raw,
        "default": field["default"],
        "using_default": not raw,
        "resolved": str(resolved),
        "relative": config.is_relative_setting(raw),
    }


def _about() -> dict[str, Any]:
    """このアプリ自身の素性。

    **版は `config/app.json` の `version` ただ1つが出どころ**で、帯の
    バッジも `/api/health` もここも同じものを読む。二重に持たない。

    番号だけでなく、置き場所・ポート・Python まで出すのは、現場から
    「動かない」と言われたときに最初に聞くことがこの一式だから。
    画面で読めれば、端末に入って調べなくても電話で確かめられる。
    """
    # 【統合版】ポートとアプリ本体は**統合アプリのもの**を出す。機能の config/app.json の
    # ポートは使われない(統合アプリが決める)ので、出すと「8740 (設定は 8733)」と
    # 繰り上がったように見えていた。版・アプリID・設定ファイルは機能のもの
    mode = str(conf().get("MODE", ""))
    port = int(conf().get("PORT", 0) or 0)
    return {
        "display_name": app_config.display_name(),
        "version": app_config.version(),
        "version_label": app_config.version_label(),
        "app_id": app_config.app_id(),
        "mode": mode,
        "port": port,
        "actual_port": int(conf().get("PORT", 0) or 0),
        "python": sys.version.split()[0],
        "python_exe": sys.executable,
        "app_root": str(_integrated_app_config.APP_ROOT),
        "module_root": str(app_config.APP_ROOT),
        "app_config_path": str(app_config.CONFIG_PATH),
        "local_root": str(app_config.local_root()),
        # 版の書き方が壊れていても起動は止めない。**ここに出して気づかせる**
        "version_problem": app_config.version_problem(),
        "config_error": app_config.load_error(),
    }


def _view(conn) -> dict[str, Any]:
    worker = user_settings.get_worker()
    return {
        "about": _about(),
        "fields": [_field_view(f) for f in FIELDS],
        # 置き場所の枠(統合 1.2.30)。画面はこれで描く
        "groups": [_group_view(g, data_sync.last_import(conn)) for g in GROUPS],
        "line": user_settings.get_line(),
        "lines": list(config.LINES),
        "worker": worker,
        # 梱包資材マスタの「班員名簿」から、担当ラインが機側の人だけ。
        # 班ごとに束ねて返す
        "workers": staff.choices(conn, current=worker),
        "auto_import": user_settings.get_auto_import(),
        "imported": data_sync.last_import(conn),
        "imported_at": data_sync.imported_at(conn),
        "stale": data_sync.is_stale(conn),
        "stale_why": data_sync.stale_message(conn),
        # 置き場所を触れるか。**画面はこれを見て、押す前に断る**
        "can_edit_paths": admin_session.peek(),
        # 配布設定(`coil_tool/distribution.py`)。パスワードの値は出さない
        "distribution": distribution.summary(),
        # 画面に出す場所は実際に置かれている場所(Microsoft Store の Python。統合 1.2.2)
        "config_path": _common_app_config.real_location(config.USER_CONFIG_PATH),
        "db_path": _common_app_config.real_location(config.DB_PATH),
        # いま書いている場所(上の帯の「ログ」で変えられる。統合 1.0.12)
        "log_dir": _common_app_config.real_location(_common_logging.log_dir()),
        "storage": _storage(),
    }


def _storage() -> dict[str, Any]:
    """設定・データの保存先。**このPCに残るもの**と**複数のPCで共有するもの**を分けて出す。

    現場の指摘: このPCで引き継いで使うものと、複数のPCで共有するものは違う。
    設定部にそういうファイルがあることを明記してほしい。
    **パスワードはこのPCのもの**(梱包明細は全ラインで共有)。取り違えないよう、ここでも言う。
    """
    P = storage_places.Place
    config_path = config.USER_CONFIG_PATH
    lot_dir = config.lot_db_dir()
    return storage_places.Places(
        local=[
            P("設定ファイル", str(config_path),
              "取り込み元の置き場所（梱包資材マスタ・仕掛台帳・予備をどこにするか）・"
              "ライン・担当者・起動時の自動取り込み・パスワード（撹拌した値。このPCだけのもの）",
              "" if config_path.exists() else "まだありません（設定を保存すると作ります）"),
            P("作業用DB", str(config.DB_PATH),
              "取り込んだマスタと仕掛台帳の写し（包装仕様・パレット・リプラサイズ・班員名簿・"
              "仕掛引当・仕掛受注）・発注チェックリスト・発注履歴・取り込みの記録"),
            storage_places.log_place(),
        ],
        shared=[
            P("梱包資材マスタ", str(config.master_db_dir() / config.MATERIAL_DB_NAME),
              "取り込み元。「マスタ管理」で直すとここへ書きます（手元の写しではなく元のファイル）"),
            P("仕掛台帳", str(lot_dir),
              "取り込み元（" + " / ".join(config.LOT_DB_FILES.values()) + "）",
              "読むだけ（このツールは書き換えません）"),
        ],
        dist=[
            P("配布設定", str(distribution.settings_path()),
              "「配布設定」の面で書き出した値（"
              + "・".join(label for _, label, _ in distribution.ITEMS) + "）"),
            storage_places.log_dist_place(),
        ],
        shared_note=("取り込み元の場所（どこを見るか）は、このPCの設定ファイルに入っています。"
                     "ここで変えても、ほかのPCは変わりません（ほかのPCもそろえるときは配布設定）。"),
    ).to_dict()


@bp.get("/api/settings")
def get_settings():
    return jsonify(ok=True, view=_view(get_db()))


@bp.get("/api/ribbon")
def get_ribbon():
    """上の帯の値(取り込み日時・古い・ライン・担当者)。帯を読み直さずに描き直すため(統合 1.2.11)。"""
    return jsonify(ok=True, ribbon=shell.ribbon(get_db()))


@bp.post("/api/ribbon/import-lot")
def import_lot_now():
    """帯の「いま取り込む」── 共有の仕掛台帳だけを取り込み直す(統合 1.2.49・現場の指摘)。

    仕掛台帳は大きいので、変わっていても自動では取り込まない(帯で知らせる)。
    覚えている写しを捨ててから読む(`data_sync.reimport_lot`)。
    """
    conn = get_db()
    if data_sync.is_importing():
        return jsonify(ok=False, reason="importing", message="取り込み中です。終わってからもう一度押してください",
                       ribbon=shell.ribbon(conn)), 409
    db.apply_schema(conn)
    tables = data_sync.reimport_lot(conn)
    failed = [t for t in tables if not t.ok]
    return jsonify(ok=not failed, tables=_table_results(tables),
                   message=" / ".join(f"{t.table}: {t.error}" for t in failed),
                   ribbon=shell.ribbon(conn))


# ==================================================================
# 確かめる ── 保存する前に、そこに届くか見る
# ==================================================================
def _typed_dir(typed: dict[str, Any], key: str) -> Optional[Path]:
    """同じ枠で**いま打ってある**フォルダ(保存前)。打っていなければ None(保存してある場所を見る)。"""
    if key not in typed:
        return None
    text = str(typed.get(key) or "").strip()
    if not text:
        return None
    try:
        return config.resolve_dir(text)
    except ValueError:
        return None


def _probe_dir(field: dict[str, Any], text: str,
               typed: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    """打たれた場所を見に行く。**保存はしない。**"""
    out: dict[str, Any] = {"key": field["key"], "ok": False,
                           "resolved": "", "message": "", "found": []}
    text = (text or "").strip()
    if field.get("kind") == "file":
        return _probe_file(field, text, out, typed or {})
    if not text:
        if field.get("optional"):
            out.update(ok=True, message="空のままにします(使いません)")
        else:
            out.update(ok=True, resolved=field["default"],
                       message="空なので既定の場所を使います")
        return out

    try:
        path = config.resolve_dir(text)
    except ValueError as exc:
        out["message"] = f"読めない書き方です: {exc}"
        return out
    out["resolved"] = str(path)

    if not path.exists():
        out["message"] = "そこにフォルダがありません(綴りと共有への接続を確認してください)"
        return out
    if not path.is_dir():
        out["message"] = "フォルダではなくファイルを指しています"
        return out

    if field.get("writable"):
        # 書く場所は**書けることまで**確かめる。読めるだけでは足りない
        probe = path / ".coil_tool_write_test"
        try:
            probe.write_text("", encoding="utf-8")
            probe.unlink()
        except OSError as exc:
            out["message"] = f"フォルダはありますが書けません: {exc}"
            return out
        out.update(ok=True, message="書けます")
        return out

    # 読む場所は**期待するファイルが在るか**まで見る。フォルダだけ在って
    # 中身が無い、が現場でいちばん多い
    try:
        files = [p.name for p in source_db.list_source_files(path)]
    except OSError as exc:
        out["message"] = f"フォルダを見られません: {exc}"
        return out

    out["found"] = files
    if not files:
        out["message"] = "フォルダはありますが、sqlite3 のファイルが見つかりません"
        return out
    out.update(ok=True, message=f"{len(files)} 件見つかりました")
    return out


def _probe_file(field: dict[str, Any], text: str, out: dict[str, Any],
                typed: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    """仕掛台帳の第2パスのファイル名を確かめる: 使える名前か・第2パスのフォルダにあるか・要る列がそろっているか。"""
    name = text or field["default"]
    if field.get("table") == "仕掛当工程" and not name:
        name = config.konpo_db_file()                 # 第2パスの名前が空なら第1パスと同じ
    problem = config.file_name_problem(name)
    if problem:
        out["message"] = problem
        return out
    if field.get("table") == "仕掛当工程":
        return _probe_ls4(field, name, out, typed or {})
    dirs = config.lot_db_dirs()
    second_typed = _typed_dir(typed or {}, config.KEY_LOT_DB_DIR_2)
    if second_typed is not None:
        dirs = [dirs[0], second_typed]
    if len(dirs) < 2:
        out.update(ok=True, message="第2パスのフォルダが空なので、この名前は使いません")
        return out
    path = source_db.find_named(dirs[1], name)
    if path is None:
        out["message"] = f"第2パスのフォルダ({dirs[1]})に {name} が見つかりません"
        return out
    out["resolved"] = str(path)
    missing = data_sync.second_missing_columns(path, field["table"])
    if missing:
        out["message"] = "見つかりましたが、足りない列があります: " + "・".join(missing)
        return out
    out.update(ok=True, message="見つかりました・要る列はそろっています")
    return out


def _probe_ls4(field: dict[str, Any], name: str, out: dict[str, Any],
               typed: Optional[dict[str, Any]] = None) -> dict[str, Any]:
    """LS4LOT の名前を確かめる(統合 1.2.29): LS4LOT なら読まない・それ以外は実績の形式の列。"""
    if not ls4_source.is_results_name(name):
        out.update(ok=True, message="名前が LS4LOT なので資材計算は読みません(外径は受注の外径 MAX)")
        return out
    if field.get("where") == "konpo":
        folder = _typed_dir(typed or {}, config.KEY_KONPO_DB_DIR) or config.konpo_db_dir()
    else:
        dirs = config.lot_db_dirs()
        second_typed = _typed_dir(typed or {}, config.KEY_LOT_DB_DIR_2)
        if second_typed is not None:
            dirs = [dirs[0], second_typed]
        if len(dirs) < 2:
            out.update(ok=True, message="第2パスのフォルダが空なので、この名前は使いません")
            return out
        folder = dirs[1]
    path = source_db.find_named(folder, name)
    if path is None:
        out["message"] = f"{folder} に {name} が見つかりません"
        return out
    out["resolved"] = str(path)
    missing = coil_results.missing_columns(path)
    if missing:
        out["message"] = "見つかりましたが、実績の形式の列が足りません: " + "・".join(missing)
        return out
    out.update(ok=True, message="見つかりました・実績の形式の列はそろっています(外径 = 内径 + 2 × 肉厚)")
    return out


@bp.post("/api/settings/probe")
def probe():
    """打たれた場所に届くかを確かめる。保存はしない。"""
    payload = request.get_json(silent=True) or {}
    results = []
    for field in FIELDS:
        if field["key"] in payload:
            results.append(_probe_dir(field, str(payload[field["key"]]), payload))
    if not results:
        return jsonify(ok=False, reason="no_field",
                       message="確かめる欄がありません"), 400
    return jsonify(ok=True, results=results)


# ==================================================================
# 保存
# ==================================================================
@bp.post("/api/settings")
def save_settings():
    """設定を保存する。**まとめて1回で書く**(半分だけ変わらないように)。"""
    payload = request.get_json(silent=True) or {}
    if not _path_gate(payload):
        # 403。**入力は正しく、いまこの操作をする資格が無い**という断り方
        return jsonify(ok=False, reason=REFUSE_NEEDS_PASSWORD,
                       message=NEEDS_PASSWORD_MESSAGE), 403

    values: dict[str, Any] = {}

    for field in FIELDS:
        key = field["key"]
        if key in payload:
            values[key] = str(payload[key] or "").strip()
            if field.get("kind") == "file":
                problem = config.file_name_problem(values[key])
                if problem:
                    return jsonify(ok=False, reason="bad_name", message=problem), 422

    if "line" in payload:
        line = str(payload["line"] or "")
        if line not in config.LINES:
            return jsonify(ok=False, reason="bad_line",
                           message=f"ラインは {' / '.join(config.LINES)} です"), 422
        values[config.KEY_LINE] = line

    if "worker" in payload:
        worker = str(payload["worker"] or "")
        # 名簿(担当ラインが機側の人)と突き合わせる。
        # **いま選ばれている名前は通す** ── 異動で名簿から消えた人や
        # 機側から外れた人を選んだままの端末が、設定を開いただけで
        # 弾かれないようにする
        if not staff.is_known(get_db(), worker,
                              current=user_settings.get_worker()):
            return jsonify(ok=False, reason="bad_worker",
                           message="担当者に選べない名前です"
                                   "(名簿にない/担当ラインが機側ではない)"), 422
        values[config.KEY_WORKER] = worker

    if "auto_import" in payload:
        values[config.KEY_AUTO_IMPORT] = bool(payload["auto_import"])

    if not values:
        return jsonify(ok=False, reason="nothing_to_save",
                       message="保存するものがありません"), 400

    if not user_settings.save_many(values):
        return jsonify(ok=False, reason="write_failed",
                       message=f"設定ファイルに書けません: {config.USER_CONFIG_PATH}"), 500

    return jsonify(ok=True, view=_view(get_db()))


def _table_results(tables) -> list[dict[str, Any]]:
    return [{"table": t.table, "ok": t.ok, "rows": t.rows, "source": t.source, "error": t.error,
             "filled": t.filled, "duplicates": t.duplicates, "note": t.note} for t in tables]


@bp.post("/api/settings/group")
def save_group():
    """1つの枠(フォルダとファイル名)をまとめて保存し、**その枠の分だけ取り込む**(統合 1.2.30)。

    梱包明細の「保存して取り込み」と同じ形。置き場所を変えただけでは手元は変わらないので、
    保存したらすぐ取り込む。
    """
    payload = request.get_json(silent=True) or {}
    group = GROUP_BY_ID.get(str(payload.get("group", "")))
    if group is None:
        return jsonify(ok=False, reason="unknown_group", message="知らない枠です"), 400
    given = payload.get("values") or {}
    if not isinstance(given, dict):
        return jsonify(ok=False, reason="bad_values", message="値の形が違います"), 400
    if not _path_gate(group["keys"]):
        return jsonify(ok=False, reason=REFUSE_NEEDS_PASSWORD,
                       message=NEEDS_PASSWORD_MESSAGE), 403
    values: dict[str, Any] = {}
    for key in group["keys"]:
        if key not in given:
            continue
        value = str(given.get(key) or "").strip()
        field = FIELD_BY_KEY[key]
        if field.get("kind") == "file":
            problem = config.file_name_problem(value)
            if problem:
                return jsonify(ok=False, reason="bad_name", key=key,
                               message=f"{field['label']}: {problem}"), 422
        elif value:
            try:
                config.resolve_dir(value)
            except ValueError as exc:
                return jsonify(ok=False, reason="bad_path", key=key,
                               message=f"{field['label']}: 読めない書き方です: {exc}"), 422
        values[key] = value
    if values and not user_settings.save_many(values):
        return jsonify(ok=False, reason="write_failed",
                       message=f"設定ファイルに書けません: {config.USER_CONFIG_PATH}"), 500

    conn = get_db()
    db.apply_schema(conn)
    # 押したときは、覚えている写しを捨ててから読む(統合 1.2.49)
    data_sync.forget_copies(data_sync.source_paths())
    tables = []
    if group["id"] == "master":
        tables += data_sync.import_master(conn)
    if group["id"] in ("lot", "second"):
        tables += data_sync.import_lot(conn)
    if group["id"] in ("lot", "konpo", "second"):
        coil = coil_results.import_results(conn)
        if coil is not None:
            tables.append(coil)
    return jsonify(ok=all(t.ok for t in tables), group=group["id"],
                   tables=_table_results(tables), view=_view(conn))


@bp.post("/api/settings/group/reset")
def reset_group():
    """1つの枠をまとめて既定に戻す(取り込みはしない。戻したあとに「保存して取り込み」)。"""
    payload = request.get_json(silent=True) or {}
    group = GROUP_BY_ID.get(str(payload.get("group", "")))
    if group is None:
        return jsonify(ok=False, reason="unknown_group", message="知らない枠です"), 400
    if not _path_gate(group["keys"]):
        return jsonify(ok=False, reason=REFUSE_NEEDS_PASSWORD,
                       message=NEEDS_PASSWORD_MESSAGE), 403
    for key in group["keys"]:
        user_settings.clear(key)
    return jsonify(ok=True, view=_view(get_db()))


@bp.post("/api/settings/reset")
def reset_field():
    """1つの欄を既定へ戻す。"""
    payload = request.get_json(silent=True) or {}
    key = str(payload.get("key", ""))
    if key not in {f["key"] for f in FIELDS}:
        return jsonify(ok=False, reason="unknown_key",
                       message="知らない設定です"), 400
    # 既定へ戻すのも**置き場所を変えること**。同じ関門を通す
    if not _path_gate([key]):
        return jsonify(ok=False, reason=REFUSE_NEEDS_PASSWORD,
                       message=NEEDS_PASSWORD_MESSAGE), 403
    user_settings.clear(key)
    return jsonify(ok=True, view=_view(get_db()))


# ==================================================================
# 取り込み
# ==================================================================
@bp.post("/api/settings/import")
def run_import():
    """共有フォルダから手元へ写す。

    **1つ失敗しても残りは続ける。** 仕掛台帳に届かない端末でも、
    マスタさえ読めれば包装仕様の閲覧はできる。
    """
    conn = get_db()
    result = data_sync.import_all(conn)
    return jsonify(
        ok=result.ok,
        total=result.total_rows,
        # 鍵が重なっていて飛ばした行(黙って捨てない)・第2パスから足した件数なども入る
        tables=_table_results(result.tables),
        view=_view(conn),
    )


# ==================================================================
# 配布設定(`coil_tool/distribution.py`)
# ==================================================================
# 書き出す・読み込み直す・消すは、どれも**置き場所を変えるのと同じ関門**
# (設定のいちばん上の認証)を通す。書き出すものには置き場所と
# 管理者パスワードが入り、読み込み直すとこの端末の置き場所が変わるため
DIST_NEEDS_PASSWORD_MESSAGE = (
    "配布設定を扱うにはパスワードが要ります。"
    "設定のいちばん上でパスワードを入れてください。")


def _distribution_reply(result):
    if not result.ok:
        status = 500 if result.reason == distribution.REFUSE_FAILED else 400
        return jsonify(ok=False, reason=result.reason or "bad_input",
                       message=result.message), status
    return jsonify(ok=True, message=result.message, view=_view(get_db()))


def _distribution_gate():
    """認証が開いていなければ断りの応答、開いていれば None。"""
    if admin_session.is_open():
        return None
    return jsonify(ok=False, reason=REFUSE_NEEDS_PASSWORD,
                   message=DIST_NEEDS_PASSWORD_MESSAGE), 403


@bp.post("/api/settings/distribution/export")
def export_distribution():
    """この端末のいまの設定を、配布設定として書き出す。"""
    refused = _distribution_gate()
    if refused:
        return refused
    items = (request.get_json(silent=True) or {}).get("items")
    if not isinstance(items, list) or not all(isinstance(x, str) for x in items):
        return jsonify(ok=False, reason="bad_input",
                       message="入れる項目の形が違います。"), 400
    return _distribution_reply(distribution.export(items))


@bp.post("/api/settings/distribution/reapply")
def reapply_distribution():
    """置いてある配布設定を読み込み直す(この端末の値も上書き)。"""
    refused = _distribution_gate()
    if refused:
        return refused
    return _distribution_reply(distribution.reapply())


@bp.post("/api/settings/distribution/remove")
def remove_distribution():
    refused = _distribution_gate()
    if refused:
        return refused
    return _distribution_reply(distribution.remove())
