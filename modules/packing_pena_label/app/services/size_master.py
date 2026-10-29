# -*- coding: utf-8 -*-
"""サイズ構成マスタ（VBA modTeslaConfig 相当）。

VBA では ``GetSizeConfig`` に Select Case でハードコードされていたが、
サイズ追加のたびにコードを触らずに済むよう ``data/size_master.json`` へ
外出しした。JSON が無い場合は VBA と同一の内蔵定義にフォールバックする。

obIdx の体系（VBA と完全に同じ）
    1..12  : OptionButton1..12
    13, 14 : lblSpec_08_53_1 / _2（名前付きコントロール）
    奇数 = 丈1 / 偶数 = 丈2
"""

from __future__ import annotations

import json
import logging
import os
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .vba_compat import val

log = logging.getLogger(__name__)

# VBA: BuildSheetName = cfg(0) & " 　丈" & GetTakeType(obIdx)
SHEET_NAME_JOINER = " 　丈"

#: サイズ名の中の「数値 + mm」
_SIZE_NUM_RE = re.compile(r"(\d+(?:\.\d+)?)mm")


def normalize_size_name(name: str) -> str:
    """サイズ名の表記ゆれを揃える（2026.09 決定）。

    VBA のサイズ名は当時の**シート名がそのまま入っている**ため、
    ``1.0mm×73mm`` だけ小数点が無い、といった揺れが残っていた。
    Excel 側はシート名を色々な設定の基準にしてしまっていて
    今さら直せないが、このツールはシート名に依存していないので、
    **見える文字だけを統一する**。

        1.0mm×73mm    -> 1.0mm×73.0mm
        0.6mm×82.5mm  -> 0.6mm×82.5mm （変化なし）

    やることは「整数だった値に ``.0`` を足す」だけで、
    **桁を削る・丸めることはしない**（``1.25mm`` は ``1.25mm`` のまま）。
    巾の取り出し（``width_from_sheet_name``）は ``Val()`` 経由なので
    ``73`` でも ``73.0`` でも同じ 73.0 になり、計算結果は変わらない。
    """
    def one(m):
        num = m.group(1)
        return (num if "." in num else num + ".0") + "mm"
    return _SIZE_NUM_RE.sub(one, str(name))

# VBA GetLabelBaseRows() と同一
LABEL_BASE_ROWS = [8, 18, 28, 38, 48, 57, 66, 75,
                   97, 106, 115, 124, 133, 142, 151, 160]

# --- ラベル1枚内の行オフセット（VBA Const と同一）---
OFS_BC_KATA = -1
OFS_COIL = 0
OFS_KENSA = 0
OFS_KATABAN = 1
OFS_WEIGHT = 2
OFS_BC_KEN = 3

# --- 左面 / 右面の列（VBA Const と同一）---
COL_COIL_L, COL_COIL_R = 7, 39
COL_KEN_L, COL_KEN_R = 4, 36
COL_KATA_L, COL_KATA_R = 4, 36
COL_WT_L, COL_WT_R = 4, 36
COL_BC_KATA_L, COL_BC_KATA_R = 9, 41
COL_BC_KEN_L, COL_BC_KEN_R = 2, 34

# --- ヘッダー部 ---
HDR_KEN_ROW, HDR_KEN_COL = 3, 22
HDR_KEN2_ROW, HDR_KEN2_COL = 92, 22
HDR_WT1_ROW, HDR_WT1_COL = 3, 37
HDR_WT2_ROW, HDR_WT2_COL = 92, 37

# --- 本数 / 高さ / NW / GW の書込列 ---
COL_COILH = 43
COL_TA = 48
COL_NW = 53
COL_GW = 59

# VBA GetCBtoPair と同一（cbIdx -> (ob1, ob2)）
CB_TO_PAIR: Dict[int, tuple] = {
    1: (7, 8),    # 1.0mm×40.0mm
    2: (9, 10),   # 1.0mm×33.0mm
    3: (1, 2),    # 1.0mm×73mm
    4: (3, 4),    # 0.6mm×82.5mm
    5: (5, 6),    # 1.0mm×53.5mm
    6: (11, 12),  # 1.0mm×63.0mm
}
#: 名前付き CheckBox（ck_08_53）。VBA では isNamed:=True で 13/14 固定
NAMED_CB_KEY = "ck_08_53"
NAMED_CB_PAIR = (13, 14)


def take_of(ob_idx) -> int:
    """その台紙番号が丈1か丈2か(VBA ``GetTakeType``: 奇数=丈1 / 偶数=丈2)。

    **丈の決まりはここ1か所。** 同梱のサイズ(1〜14)も、足したサイズ(15〜)も同じ。
    以前は重量の未入力チェックが「11 以下の奇数か 13」「12 以下の偶数か 14」と番号を
    決め打ちしていて、足したサイズは重量が空でも重量反映が通った(統合 1.2.26・構造点検)。
    """
    return 1 if int(ob_idx) % 2 == 1 else 2


@dataclass
class SizeConfig:
    """VBA ``GetSizeConfig`` の戻り値 Array(0..4) に対応。"""
    ob_idx: int
    base_name: str          # cfg(0) シート名ベース  例 "1.0mm×53.5mm"
                            # 表記は normalize_size_name() で統一済み
    kataban: str            # cfg(1) 型番            例 "BJB7606500QR"
    symbol: str             # cfg(2) 記号            例 "d"
    data_row: int           # cfg(3) データ行        3 または 92
    extra_row: int          # cfg(4) 重量追加行      0 または 92
    width_mm: float         # 巾（SetCoilDimensions のハードコード値）
    cb_key: str             # 対応する CheckBox 名
    ob_key: str             # 対応する OptionButton 名
    size_cell_label: str    # Cells(dataRow, 2) の内容（シート上の固定文字）

    @property
    def take_type(self) -> int:
        """VBA GetTakeType: 奇数=丈1 / 偶数=丈2。"""
        return take_of(self.ob_idx)

    @property
    def clears_coil_on_select(self) -> bool:
        """選んだときに、その丈のコイル高さ・NW・GW・TA を消すか。

        VBA は 53.5mm の 2 サイズ(OB5/6・lblSpec_08_53_1/2)を選んだときだけ消す。足したサイズも
        巾 53.5mm なら同じ扱いにする(統合 1.2.27・現場の判断)。番号(5,6,13,14)で決め打ちしない。
        """
        return abs(self.width_mm - CLEAR_ON_SELECT_WIDTH_MM) < 1e-6

    @property
    def sheet_name(self) -> str:
        """VBA BuildSheetName。"""
        return f"{self.base_name}{SHEET_NAME_JOINER}{self.take_type}"

    def to_dict(self) -> dict:
        d = {
            "obIdx": self.ob_idx, "baseName": self.base_name,
            "kataban": self.kataban, "symbol": self.symbol,
            "dataRow": self.data_row, "extraRow": self.extra_row,
            "widthMm": self.width_mm, "cbKey": self.cb_key,
            "obKey": self.ob_key, "takeType": self.take_type,
            "sheetName": self.sheet_name,
            "sizeCellLabel": self.size_cell_label,
        }
        return d


#: この巾のサイズは、選んだときにコイル高さ・NW・GW・TA を消す(`SizeConfig.clears_coil_on_select`)
CLEAR_ON_SELECT_WIDTH_MM = 53.5

#: VBA GetSizeConfig の内蔵定義（JSON が無い場合のフォールバック）
#: (baseName, kataban, symbol, dataRow, extraRow, widthMm, cbKey)
_BUILTIN = {
    (1, 2):   ("1.0mm×73mm",   "BJB7604000QR", "c", 3,  0,  73.0,  "CheckBox3"),
    (3, 4):   ("0.6mm×82.5mm", "BJB7604200QR", "e", 3,  0,  82.5,  "CheckBox4"),
    (5, 6):   ("1.0mm×53.5mm", "BJB7606500QR", "d", 3,  92, 53.5,  "CheckBox5"),
    (7, 8):   ("1.0mm×40.0mm", "BJB7604300QR", "b", 3,  92, 40.0,  "CheckBox1"),
    (9, 10):  ("1.0mm×33.0mm", "BJB7603900QR", "a", 3,  92, 33.0,  "CheckBox2"),
    (11, 12): ("1.0mm×63.0mm", "BJB7605900QR", "f", 92, 0,  63.0,  "CheckBox6"),
    (13, 14): ("0.8mm×53.5mm", "BJB7610400QR", "g", 3,  92, 53.5,  NAMED_CB_KEY),
}

_OB_KEYS = {i: f"OptionButton{i}" for i in range(1, 13)}
_OB_KEYS[13] = "lblSpec_08_53_1"
_OB_KEYS[14] = "lblSpec_08_53_2"


def _file_stamp(path: Optional[str]):
    """ファイルの印 (パス, 大きさ, 更新日時ns)。無ければ None。**読む前**に取る。"""
    if not path:
        return None
    try:
        st = os.stat(path)
    except OSError:
        return None
    return (os.path.abspath(path), st.st_size, st.st_mtime_ns)


class SizeMaster:
    """サイズ構成マスタ。"""

    def __init__(self, json_path: Optional[str] = None):
        self._by_ob: Dict[int, SizeConfig] = {}
        #: あとから足したサイズ(共有マスタの表。`size_extra`)。読み直しても足し直す
        self._extra_rows: List[dict] = []
        self._extra_obs: set = set()
        self._notes: Dict[str, str] = {}
        #: ラベル台紙一覧で上書きする前の値 ``{obIdx: (シート名, 型番)}``(JSON・内蔵定義・足したサイズの元の値)。
        #: 上書きの前に毎回ここへ戻す(統合 1.2.49。空にした・行を消したのに前の上書きが残っていた)
        self._base_labels: Dict[int, Tuple[str, str]] = {}
        #: 最後に重ねたラベル台紙一覧(読み直したときに掛け直す)
        self._overlay_rows: Dict[int, tuple] = {}
        #: 読んだ JSON と、その印(読む前に取る。外で直されたら読み直す。統合 1.2.49)
        self._json_path = json_path or ""
        self._json_stamp = _file_stamp(json_path)
        loaded = False
        if json_path and os.path.exists(json_path):
            try:
                self._load_json(json_path)
                loaded = True
            except Exception:
                loaded = False
        if not loaded:
            self._load_builtin()

    # ------------------------------------------------------------ 読込
    def _load_builtin(self) -> None:
        for (o1, o2), (raw, kata, sym, drow, erow, w, cb) in _BUILTIN.items():
            base = normalize_size_name(raw)
            for ob in (o1, o2):
                take = 1 if ob % 2 == 1 else 2
                self._by_ob[ob] = SizeConfig(
                    ob_idx=ob, base_name=base, kataban=kata, symbol=sym,
                    data_row=drow, extra_row=erow, width_mm=w,
                    cb_key=cb, ob_key=_OB_KEYS[ob],
                    size_cell_label=f"{base}{SHEET_NAME_JOINER}{take}",
                )
                self._keep_base(ob)

    def _load_json(self, path: str) -> None:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        for row in data["sizes"]:
            base = normalize_size_name(row["baseName"])
            for ob in row["obIdx"]:
                take = 1 if ob % 2 == 1 else 2
                default_label = f"{base}{SHEET_NAME_JOINER}{take}"
                labels = row.get("sizeCellLabel") or {}
                self._by_ob[ob] = SizeConfig(
                    ob_idx=ob,
                    base_name=base,
                    kataban=row["kataban"],
                    symbol=row["symbol"],
                    data_row=int(row["dataRow"]),
                    extra_row=int(row["extraRow"]),
                    width_mm=float(row["widthMm"]),
                    cb_key=row["cbKey"],
                    ob_key=_OB_KEYS.get(ob, f"OptionButton{ob}"),
                    size_cell_label=labels.get(str(ob), default_label),
                )
                self._keep_base(ob)
        if not self._by_ob:
            raise ValueError("size_master.json にサイズ定義がありません")

    def apply_extra(self, rows) -> int:
        """あとから足したサイズ(``size_extra``)を足す。前に足した分は入れ替える。

        ``rows`` は ``size_extra.read_rows`` の形。同梱(JSON)にある台紙番号・寸法は足さない
        (同梱が先。共有マスタの表で同梱のサイズを上書きさせない)。
        """
        for ob in self._extra_obs:
            self._by_ob.pop(ob, None)
        self._extra_obs = set()
        self._extra_rows = [dict(r) for r in (rows or [])]
        names = {c.base_name for c in self._by_ob.values()}
        n = 0
        for r in self._extra_rows:
            ob1 = int(r["obIdx"])
            base = normalize_size_name(r["baseName"])
            if ob1 in self._by_ob or (ob1 + 1) in self._by_ob or base in names:
                continue
            for ob in (ob1, ob1 + 1):
                take = 1 if ob % 2 == 1 else 2
                self._by_ob[ob] = SizeConfig(
                    ob_idx=ob, base_name=base, kataban=r.get("kataban", ""),
                    symbol=r.get("symbol", ""),
                    # 台紙の上下とも使う(1.0mm×53.5mm などと同じ)
                    data_row=3, extra_row=92, width_mm=float(r["widthMm"]),
                    cb_key="CheckBox%d" % extra_cb_index(ob1),
                    ob_key="OptionButton%d" % ob,
                    size_cell_label=f"{base}{SHEET_NAME_JOINER}{take}",
                )
                self._keep_base(ob)
                self._extra_obs.add(ob)
            self._notes[base] = r.get("note", "")
            names.add(base)
            n += 1
        # 足し直したサイズにも、前に重ねたラベル台紙一覧を掛け直す
        if self._overlay_rows:
            self.apply_label_overlay(self._overlay_rows)
        return n

    def _keep_base(self, ob: int) -> None:
        """作ったばかりの台紙の シート名・型番 を、上書きの前の値として控える。"""
        cfg = self._by_ob[ob]
        self._base_labels[ob] = (cfg.size_cell_label, cfg.kataban)

    def is_extra(self, ob_idx: int) -> bool:
        """あとから足したサイズか。"""
        return int(ob_idx) in self._extra_obs

    def note_of(self, base_name: str) -> str:
        """足したサイズの備考(指定サイズの画面に赤字で出す)。"""
        return self._notes.get(base_name, "")

    def apply_label_overlay(self, rows) -> int:
        """共有マスタの ``ラベル台紙一覧`` でシート名・型番を上書きする。

        JSON は土台、マスタは上書き。
        表が無い端末では何も起きず、JSON の値がそのまま使われる。

        ``rows`` は ``{obIdx: (シート名, 型番)}``。
        空欄は「指定なし」として **上書きしない**
        （うっかり空にして型番が消えるのを避けるため）。

        **重ねる前に、毎回 JSON の値(上書きの前の値)へ戻す**(統合 1.2.49・現場の指摘)。
        以前はその場で上書きするだけだったので、型番・シート名を空にした・行を消した・表が読めなく
        なった、のどれでも、動いたままの PC には前の上書きが残った(再起動した PC は JSON の値になり、
        PC によって刷る型番が違った)。``rows`` が空なら JSON の値に戻るだけ。
        """
        self._overlay_rows = {int(ob): tuple(pair) for ob, pair in (rows or {}).items()}
        for ob, cfg in self._by_ob.items():
            base = self._base_labels.get(ob)
            if base is not None:
                cfg.size_cell_label, cfg.kataban = base
        n = 0
        for ob, pair in self._overlay_rows.items():
            cfg = self._by_ob.get(int(ob))
            if cfg is None:
                continue
            sheet, kataban = (list(pair) + ["", ""])[:2]
            sheet = str(sheet or "").strip()
            kataban = str(kataban or "").strip()
            changed = False
            if sheet and sheet != cfg.size_cell_label:
                cfg.size_cell_label = sheet
                changed = True
            if kataban and kataban != cfg.kataban:
                cfg.kataban = kataban
                changed = True
            if changed:
                n += 1
        if n:
            log.info("ラベル台紙一覧で %d 件を上書きしました", n)
        return n

    def reload(self, json_path: Optional[str] = None) -> bool:
        """マスタを読み直す（設定でファイルを差し替えたとき）。

        読み込みに失敗した場合は **今の内容を保持** して False を返す。
        壊れた JSON を掴んで業務が止まる方が困るため。
        """
        prev = dict(self._by_ob)
        prev_obs = set(self._extra_obs)
        prev_base = dict(self._base_labels)
        self._json_path = json_path or ""
        self._json_stamp = _file_stamp(json_path)          # 読む前に取る
        try:
            self._by_ob = {}
            self._extra_obs = set()
            self._base_labels = {}
            if json_path:
                if not os.path.exists(json_path):
                    # 設定したファイルが見えない(共有に届かない・差し替え中)。同梱の定義へ黙って戻さない
                    # (統合 1.2.61。戻すと、選んでいたサイズが外れ、違う型番で刷ることがある)。今の内容のまま
                    raise FileNotFoundError("サイズ構成マスタが見えません: %s" % json_path)
                self._load_json(json_path)
            else:
                self._load_builtin()            # 設定していないときだけ同梱の定義
            # 足したサイズを足し直し、ラベル台紙一覧の上書きも掛け直す(apply_extra の中で)
            self.apply_extra(self._extra_rows)
            return True
        except Exception as exc:
            self._by_ob = prev
            self._extra_obs = prev_obs
            self._base_labels = prev_base
            log.warning("サイズ構成マスタを読み直せません（従来の内容を維持）: %s", exc)
            return False

    def refresh_if_changed(self) -> bool:
        """size_master.json が外で直されていたら読み直す(統合 1.2.49・現場の指摘)。

        以前は起動時と設定の保存時にしか読まず、JSON を直しても再起動するまで前の型番・シート名で
        刷れた。ほかのマスタと同じく、読む前に取った印(大きさ・更新日時)と比べる。
        読み直したら、ラベル台紙一覧の上書きも掛け直す。戻り値は「読み直したか」。
        """
        if not self._json_path:
            return False
        if _file_stamp(self._json_path) == self._json_stamp:
            return False
        ok = self.reload(self._json_path)
        if ok:
            log.info("size_master.json が直されたので読み直しました: %s", self._json_path)
        return ok

    # ------------------------------------------------------------ 参照
    def get(self, ob_idx: int) -> Optional[SizeConfig]:
        """VBA GetSizeConfig。未定義は None（VBA は Array("","","",0,0)）。"""
        return self._by_ob.get(int(ob_idx))

    def all(self) -> List[SizeConfig]:
        return [self._by_ob[k] for k in sorted(self._by_ob)]

    def ob_indexes(self) -> List[int]:
        return sorted(self._by_ob)

    def display_name(self, ob_idx: int) -> str:
        """VBA GetSizeDisplayName。"""
        cfg = self.get(ob_idx)
        return cfg.base_name if cfg else ""

    def sheet_name(self, ob_idx: int) -> str:
        """VBA BuildSheetName。"""
        cfg = self.get(ob_idx)
        return cfg.sheet_name if cfg else ""

    def take_type(self, ob_idx: int) -> int:
        """VBA GetTakeType。cfg が無くても奇偶で決まる。"""
        return take_of(ob_idx)

    # ------------------------------------------------------------ ペア
    def cb_to_pair(self, cb_idx: int) -> tuple:
        """VBA GetCBtoPair。該当なしは (0, 0)。足したサイズは自分の CheckBox 名で探す。"""
        cb_idx = int(cb_idx)
        if cb_idx in CB_TO_PAIR:
            return CB_TO_PAIR[cb_idx]
        obs = sorted(c.ob_idx for c in self._by_ob.values()
                     if c.cb_key == "CheckBox%d" % cb_idx)
        return (obs[0], obs[1]) if len(obs) >= 2 else (0, 0)

    def pair_for_named_cb(self) -> tuple:
        return NAMED_CB_PAIR

    def cb_pairs(self) -> Dict[int, tuple]:
        return dict(CB_TO_PAIR)

    # ------------------------------------------------------------ 巾
    def width_of(self, ob_idx: int) -> float:
        """VBA SetCoilDimensions / GetHBCoilDimensions の巾（ハードコード側）。"""
        cfg = self.get(ob_idx)
        return cfg.width_mm if cfg else 0.0

    def width_from_sheet_name(self, ob_idx: int) -> float:
        """VBA GetSelectedCoilWidth（ExtractWidthFromName）。

        シート名ベースを "×" で分割して後半から mm を取る。
        ``ck_08_53`` 選択時は VBA が 53.5 固定を返すため呼び出し側で処理。
        """
        cfg = self.get(ob_idx)
        if not cfg:
            return 0.0
        parts = cfg.base_name.split("×")
        if len(parts) >= 2:
            return val(parts[1].replace("mm", ""))
        return 0.0


#: 足したサイズの台紙番号はここから(同梱の 1..14 と重ねない。`size_extra`)
FIRST_EXTRA_OB = 15
#: 足したサイズの「丈1・丈2 同時」の番号はここから(同梱の CheckBox1..6 と重ねない)
EXTRA_CB_BASE = 100


def extra_cb_index(ob1: int) -> int:
    """足したサイズの「丈1・丈2 同時」の番号(CheckBox の番号)。同梱の 1..6 と重ねない。"""
    return EXTRA_CB_BASE + (int(ob1) + 1) // 2


def is_extra_selection(cb_key: str, ob_idx: int) -> bool:
    """選んでいるのが足したサイズか(丈1/丈2 なら台紙番号、同時なら CheckBox の番号で見る)。"""
    if int(ob_idx or 0) >= FIRST_EXTRA_OB:
        return True
    m = re.fullmatch(r"CheckBox(\d+)", str(cb_key or ""))
    return bool(m and int(m.group(1)) >= EXTRA_CB_BASE)


_default: Optional[SizeMaster] = None


def get_size_master(json_path: Optional[str] = None) -> SizeMaster:
    """既定のサイズマスタ（プロセス内キャッシュ）。"""
    global _default
    if _default is None or json_path is not None:
        m = SizeMaster(json_path)
        if json_path is None or _default is None:
            _default = m
        return m
    return _default
