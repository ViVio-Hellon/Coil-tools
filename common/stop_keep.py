r"""強制で止める前に、保存していない入力を控えに残し、次の起動で知らせる(統合 1.2.67。管理者以外停止モード)

現場の声:「保存していない入力・処理中の PC は重要な内容と考えている。止まらなければ5分たったら強制で止める。
保存の持ち方は考えてください」。

【持ち方】
- **業務のデータ(明細・記録・発注・チェックリスト)には書かない。** 「送る・登録するまで書かない」という
  約束を守る(勝手に送ると、捨てるつもりだった入力が入ってしまう)
- 代わりに、この PC の手元(`%LOCALAPPDATA%\CoilPackingTools\data\停止控え\`)へ**控え**として書く:

      画面の入力        統合画面が5秒ごとに答える「保存していない入力の名前」と「打っていた値」
                        (変えた欄だけ。パスワードの欄は持たない)。値はブラウザ・窓の中にしか無いので、
                        次の起動で**画面に並べて見せる**(打ち直してもらう。勝手に欄へ戻さない ──
                        LOT や重量を黙って入れ直すと、別のコイルの値で送ってしまうおそれがある)
      資材計算の作業    計算の作業状態(LOT・受注・計算の結果)はプロセスのメモリにしか無い。控えに書き、
                        次の起動で**そのまま戻す**(まだチェックリストへ積んでいない状態)。
                        発注票の下書きは、チェックリスト(DB に残る)からいつでも作り直せるので持たない

- 次の起動で、統合画面に「戻しました/消えた入力はこれです」を出す。**利用者が閉じるまで出し続ける**
- 知らせた控えは `停止控え\戻した\` へ移して残す(もう一度知らせない。念のため消さない)
- **控えを読めなければ知らせず、控えも残す**(壊れていれば記録に出す)
"""
from __future__ import annotations

import dataclasses
import json
import os
import time
from pathlib import Path
from typing import Any, Optional

from common.logging_utils import get_logger

log = get_logger("coil_packing_tools", "stop_keep")

DIR_NAME = "停止控え"
DONE_NAME = "戻した"
PREFIX = "控え-"
MATERIAL_LABEL = "資材計算の作業"


def keep_dir() -> Path:
    from common import app_config
    return app_config.local_dir("data") / DIR_NAME


def _write(path: Path, data: dict) -> bool:
    tmp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, path)
        return True
    except OSError as exc:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        log.error("保存していない入力の控えを書けませんでした(%s): %s", path, exc)
        return False


# ------------------------------------------------------------------
# 資材計算の作業状態
# ------------------------------------------------------------------
def _material_state() -> Optional[dict]:
    """資材計算の作業状態(LOT が入っているときだけ)。読み込んでいなければ None(作らない)。"""
    import sys
    session = sys.modules.get("modules.packing_material_calculation.app.session")
    if session is None:
        return None
    state = session.get()
    if not str(getattr(state, "LOT", "") or "").strip():
        return None
    data = dataclasses.asdict(state)
    data["flags"] = sorted(data.get("flags") or [])
    return data


def _material_label(data: dict) -> str:
    lot = str(data.get("LOT") or "").strip()
    order = str((data.get("order") or {}).get("受注番号") or "").strip()
    extra = "・".join(x for x in (f"LOT {lot}" if lot else "", f"受注 {order}" if order else "") if x)
    return f"{MATERIAL_LABEL}({extra})" if extra else MATERIAL_LABEL


def build_material_state(data: dict) -> Any:
    """控えから資材計算の作業状態を作り直す。**知らない欄は捨てる**(版が変わっても読める)。"""
    from modules.packing_material_calculation.coil_tool.models import CalcState, OrderInfo
    names = {f.name for f in dataclasses.fields(CalcState)}
    order_names = {f.name for f in dataclasses.fields(OrderInfo)}
    values = {k: v for k, v in data.items() if k in names and k not in ("order", "flags")}
    order = OrderInfo(**{k: str(v) for k, v in (data.get("order") or {}).items() if k in order_names})
    state = CalcState(**values)
    state.order = order
    state.flags = set(str(x) for x in (data.get("flags") or []))
    return state


# ------------------------------------------------------------------
# 書く(強制で止める前)
# ------------------------------------------------------------------
def save_before_force() -> list[str]:
    """保存していない入力を控えに書く。書けたものの呼び名を返す。"""
    from common import remote_stop
    names, values = remote_stop.screen_values()
    names = [n for n in names if n != remote_stop.WAITING_SCREEN]
    try:
        material = _material_state()
    except Exception:                         # noqa: BLE001 - 片方が駄目でももう片方は残す
        log.exception("資材計算の作業状態を読めませんでした")
        material = None
    if not names and not values and material is None:
        return []
    data = {"saved_at": time.time(), "screens": names, "fields": values, "material": material}
    path = keep_dir() / f"{PREFIX}{time.strftime('%Y%m%d-%H%M%S')}.json"
    if not _write(path, data):
        return []
    saved = list(names) or (["画面の入力"] if values else [])
    if material is not None:
        saved.append(_material_label(material))
    log.warning("保存していない入力を控えに残しました(管理者以外停止で強制で止める前): %s (%s)",
                "、".join(saved), path)
    return saved


# ------------------------------------------------------------------
# 戻す(次の起動)
# ------------------------------------------------------------------
def _pending() -> list[Path]:
    try:
        return sorted(keep_dir().glob(f"{PREFIX}*.json"))
    except OSError:
        return []


def _done(path: Path) -> None:
    target = keep_dir() / DONE_NAME
    try:
        target.mkdir(parents=True, exist_ok=True)
        os.replace(path, target / path.name)
    except OSError as exc:
        log.warning("知らせた控えを移せませんでした(次の起動でもう一度知らせます): %s", exc)


def _field_text(fields: list) -> str:
    parts = []
    for item in fields[:30]:
        if not isinstance(item, dict):
            continue
        screen = str(item.get("screen") or "").strip()
        label = str(item.get("label") or "").strip() or "(名前の無い欄)"
        value = str(item.get("value") or "")
        parts.append(f"{screen + ' ' if screen else ''}{label}「{value}」")
    more = len(fields) - 30
    return "、".join(parts) + (f" ほか {more} 欄" if more > 0 else "")


def restore() -> list[str]:
    """控えがあれば、資材計算の作業を戻し、消えた画面の入力を知らせる。知らせる一言を返す。"""
    notes: list[str] = []
    for path in _pending():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("中身が JSON のまとまりではありません")
        except (OSError, ValueError) as exc:
            log.warning("停止の控えを読めないので知らせませんでした(控えは残しています: %s): %s", path, exc)
            continue
        try:
            when = time.strftime("%m/%d %H:%M", time.localtime(float(data.get("saved_at") or 0)))
        except (TypeError, ValueError, OverflowError, OSError):
            when = "前回"
        found: list[str] = []
        material = data.get("material")
        if isinstance(material, dict):
            found.append(_restore_material(material, when))
        screens = [str(x) for x in (data.get("screens") or [])]
        fields = [x for x in (data.get("fields") or []) if isinstance(x, dict)]
        if screens or fields:
            text = f"{when} に管理者の停止で保存できずに消えた入力があります"
            if screens:
                text += f": {'、'.join(screens)}"
            text += "。"
            if fields:
                text += f"打っていた値: {_field_text(fields)}。"
            text += f"確かめて打ち直してください(控え: {path.name})"
            found.append(text)
        for note in found:
            log.warning("%s", note)
        notes += found
        _done(path)
    return notes


def _restore_material(data: dict, when: str) -> str:
    from modules.packing_material_calculation.app import session
    label = _material_label(data)
    current = session.get()
    if str(getattr(current, "LOT", "") or "").strip():
        return (f"{when} に管理者の停止で保存できなかった{label}の控えがありますが、いま別の作業があるので"
                "戻しませんでした。")
    try:
        state = build_material_state(data)
    except (TypeError, ValueError) as exc:
        log.warning("資材計算の作業の控えを戻せませんでした: %s", exc)
        return f"{when} に管理者の停止で保存できなかった{label}を戻せませんでした(控えが読めません)。"
    with session._lock:
        session._state = state
    return (f"{when} に管理者の停止で保存できなかった{label}を戻しました(まだチェックリストへ積んでいません)。"
            "資材計算の画面で確かめてください。要らなければ「はじめから」で消してください。")
