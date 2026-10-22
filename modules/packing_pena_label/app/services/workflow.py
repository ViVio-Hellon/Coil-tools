# -*- coding: utf-8 -*-
"""業務フロー。``テスラ指定サイズ`` UserForm の各ボタン処理を移植したもの。

VBA -> Python 対応

    OptionButtonN_Click / SelectSize     -> select_size()
    CheckBoxN_Click / SelectCheckBox     -> select_checkbox()
    CommandButton1 (重量反映)            -> apply_weight()
    CommandButton16 (重量計算_DB)        -> calc_tare()
    CommandButton2  (小ラベル印刷)       -> print_targets() + clear_all()
    CommandButton9/11 (羅列計算)          -> calc_list()
    CommandButton13 (積み高さ)           -> quick_height()
    CommandButton8  (文字分け)           -> 移植しない(VBA では要ったが、今は不要。現場の判断)
    クリア()                              -> clear_all()
    ラベル転記OP / ラベル転記CK           -> _transfer_coil_info()
"""

from __future__ import annotations

import logging
import threading
import time
from typing import Dict, List, Optional, Tuple

from ..models import FormState, TIP_CHOICES
from ..repositories import material_repo as M
from ..repositories.access_bridge import AccessError
from . import coil_outer
from . import work_record as WR
from . import label_layout as LL
from . import list_calc as LC
from . import size_master as SM
from . import tare_calc as TC
from . import validation as V
from . import label_sheet_master as LSM
from . import size_extra as SX
from .all_size import AllSizeService
from .vba_compat import fmt, val, ws_round, ws_roundup

log = logging.getLogger(__name__)

#: 風袋計算シートの開始行（VBA WriteToSheet の Select Case）
TARE_START_ROW = {1: 1, 2: 15, 3: 29, 4: 43}
#: 同 見出し
TARE_TITLE = {1: "丈1_1梱包目:", 2: "丈1_2梱包目:",
              3: "丈2_1梱包目:", 4: "丈2_2梱包目:"}
#: 風袋計算シートの項目名（VBA WriteToSheet の A 列）
TARE_ITEMS = ["ﾋﾟﾝ", "ﾁｯﾌﾟﾎﾞｰﾙ", "ｽﾄﾚｯﾁﾌｨﾙﾑ", "ｴｻﾌｫｰﾑ", "ﾎﾟﾘｼｰﾄ",
              "ﾊｰﾄﾞﾎﾞｰﾄﾞ", "樹脂ﾊﾟﾚｯﾄ", "ﾍﾟｯﾄﾊﾞﾝﾄﾞ", "EX-DRY",
              "NW", "風袋", "GW"]


class Result:
    """処理結果。VBA の MsgBox / 画面更新に相当する情報をまとめて返す。"""

    def __init__(self, ok: bool = True, message: str = "",
                 level: str = "info", data: Optional[dict] = None):
        self.ok = ok
        self.message = message
        self.level = level          # info / warn / error
        self.data = data or {}

    def to_dict(self) -> dict:
        return {"ok": self.ok, "message": self.message,
                "level": self.level, **self.data}


class Progress:
    """時間のかかる処理の進み具合(重量計算_DB)。画面が `/api/progress` で聞く。

    **何段目まで済んだか**を持つだけ。資材マスタの確かめ(Access だと数秒かかる)・
    梱包ごとの計算・保存の段ごとに進める。画面は押してから答えが返るまで、
    これを見て進み具合の棒を出す(現場の指摘: 重いのはいいが進み具合を出すこと)。
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._now: dict = {"active": False}

    def begin(self, title: str, total: int) -> None:
        with self._lock:
            self._now = {"active": True, "title": title, "step": 0,
                         "total": max(1, int(total)), "text": "", "startedAt": time.time()}

    def step(self, text: str) -> None:
        with self._lock:
            if self._now.get("active"):
                self._now["step"] = min(self._now["step"] + 1, self._now["total"])
                self._now["text"] = text

    def end(self) -> None:
        with self._lock:
            if self._now.get("active"):
                self._now.update(active=False, step=self._now["total"], text="済みました")

    def snapshot(self) -> dict:
        with self._lock:
            out = dict(self._now)
        if out.get("startedAt"):
            out["elapsed"] = round(time.time() - out.pop("startedAt"), 1)
        return out


class Workflow:
    """業務フローの入口。routes からはこのクラスだけを呼ぶ。"""

    def __init__(self, store, master: SM.SizeMaster,
                 materials: M.MaterialRepository, cfg):
        self.store = store
        self.master = master
        self.materials = materials
        self.cfg = cfg
        #: 正常終了を要求されたときに呼ぶ関数（server.py が差し込む）。
        #: 基盤仕様書 2.8「まずアプリ自身へ正常終了を要求する」経路。
        self.shutdown_hook = None
        #: 重量計算_DB の進み具合(画面が聞く)
        self.progress = Progress()
        #: テスラ全サイズ画面の業務処理
        self.all_size = AllSizeService(store, cfg.kataban_csv_path, cfg.material_db_file)
        #: 共有マスタの ペナラベル追加サイズ（あとから足したサイズ。統合 1.2.16）
        self.size_extra = SX.SizeExtra()
        #: 共有マスタの ラベル台紙一覧（シート名・型番）
        self.label_sheets = LSM.LabelSheetMaster()
        self.refresh_label_sheets()

    def refresh_sizes(self, force: bool = False) -> None:
        """足したサイズを読み直す(別の PC で足された場合も、更新日時が変われば追いつく)。"""
        try:
            rows = self.size_extra.load(self.materials.db_path, force=force)
            # 変わったときだけ入れ替える(入れ替えるとラベル台紙一覧の上書きが外れるので、掛け直す)
            if rows != self.master._extra_rows:
                self.master.apply_extra(rows)
                self.label_sheets.invalidate()
                self.label_sheets.apply_to(self.master, self.materials.db_path, force=True)
                self._drop_vanished_selection()
        except Exception as exc:                      # noqa: BLE001
            log.warning("足したサイズを読み込めません: %s", exc)

    def _drop_vanished_selection(self) -> None:
        """選んでいたサイズが消えた(マスタ管理で消した足したサイズ)なら、選択と現状値を外す(統合 1.2.38)。"""
        st = self.load_state()
        gone_ob = bool(st.selected_ob) and self.master.get(st.selected_ob) is None
        gone_cb = bool(st.selected_cb) and not any(
            c.cb_key == "CheckBox%d" % st.selected_cb for c in self.master.all())
        if not (gone_ob or gone_cb):
            return
        if gone_ob:
            st.selected_ob = 0
        if gone_cb:
            st.selected_cb = 0
        st.lbl_kensa_no = st.lbl_size1 = st.lbl_weight1 = st.lbl_size2 = st.lbl_weight2 = ""
        self.save_state(st)
        log.info("選んでいたサイズが消えたので、選択を外しました")

    def refresh_label_sheets(self, force: bool = False) -> int:
        """共有マスタの ラベル台紙一覧 を読み直して上書きする。

        ラベル台紙の画面を開いたとき・マスタを書き換えたときに呼ぶ。
        別の PC で直された場合も、更新日時が変われば追いつく。
        足したサイズを先に読む(足したサイズの型番もラベル台紙一覧で直せる)。
        """
        self.refresh_sizes(force=force)
        try:
            return self.label_sheets.apply_to(
                self.master, self.materials.db_path, force=force)
        except Exception as exc:                      # noqa: BLE001
            log.warning("ラベル台紙一覧を反映できません: %s", exc)
            return 0

    # ============================================================ 状態
    def load_state(self) -> FormState:
        return FormState.from_dict(self.store.get_kv("form_state", {}))

    def save_state(self, state: FormState) -> None:
        self.store.set_kv("form_state", state.to_dict())

    # ------------------------------------------------------------
    def _size_label(self, ob_idx: int) -> str:
        """VBA ``Cells(dataRow, 2)``（台紙の固定文字）。"""
        cfg = self.master.get(ob_idx)
        return cfg.size_cell_label if cfg else ""

    def _read_sheet_header(self, ob_idx: int) -> Tuple[str, str]:
        """台紙から (検査番号, 重量) を読む。VBA の ``Cells(dataRow, 22/37)``。"""
        cfg = self.master.get(ob_idx)
        if not cfg:
            return "", ""
        ken = self.store.get_cell(ob_idx, cfg.data_row, SM.HDR_KEN_COL)
        wt = self.store.get_cell(ob_idx, cfg.data_row, SM.HDR_WT1_COL)
        return ken, wt

    # ============================================================ 選択
    def select_size(self, state: FormState, ob_idx: int) -> Result:
        """VBA ``SelectSize`` + ``OptionButtonN_Click``。"""
        cfg = self.master.get(ob_idx)
        if not cfg:
            return Result(False, "サイズ定義が見つかりません（obIdx=%s）" % ob_idx,
                          "error")

        take = cfg.take_type

        # --- CheckBox 全解除（VBA と同じ）---
        state.selected_cb = 0
        state.named_cb = False
        state.selected_ob = ob_idx

        # --- Take / 重量入力欄の表示切替 ---
        if take == 1:
            state.weight2 = ""
        else:
            state.weight1 = ""

        # --- ラベル初期化 ---
        state.lbl_kensa_no = ""
        state.lbl_size1 = ""
        state.lbl_weight1 = ""
        state.lbl_size2 = ""
        state.lbl_weight2 = ""

        # --- 台紙から読込 ---
        ken, wt = self._read_sheet_header(ob_idx)
        state.lbl_kensa_no = ken
        # B-3: 重量が空のとき、VBA は "kg" だけを表示していた
        #（`Cells(dataRow, 37) & "kg"` が空セルでも "kg" になる）。
        # `テスラ全サイズ` 側には `If Label17 = "kg" Then Label17 = ""` という
        # 後始末があるのに、`テスラ指定サイズ` 側に無いだけの取りこぼし。
        # 空欄にそろえる。**計算には影響しない**
        #（`parse_label_weight` は "" も "kg" も 0.0 として扱う）。
        if take == 1:
            state.lbl_weight1 = (wt + "kg") if wt != "" else ""
            state.lbl_size1 = self._size_label(ob_idx)
        else:
            state.lbl_weight2 = (wt + "kg") if wt != "" else ""
            state.lbl_size2 = self._size_label(ob_idx)

        # --- VBA: OB5/6 と lblSpec_08_53_1/2 のみ CoilH/NW/GW/TA をクリア ---
        # 53.5mm のサイズ。足したサイズも巾 53.5mm なら同じ(統合 1.2.27。サイズの性質で決める)
        if cfg.clears_coil_on_select:
            slots = (1, 2) if take == 1 else (3, 4)
            for i in slots:
                state.coil_h[i] = ""
                state.nw[i] = ""
                state.gw[i] = ""
                state.ta[i] = ""

        self.save_state(state)
        return Result(True, "", "info", {"state": state.to_dict()})

    def select_checkbox(self, state: FormState, cb_idx: int,
                        named: bool = False) -> Result:
        """VBA ``SelectCheckBox`` + ``CheckBoxN_Click``。"""
        if named:
            ob1, ob2 = self.master.pair_for_named_cb()
            state.selected_cb = 0
            state.named_cb = True
        else:
            ob1, ob2 = self.master.cb_to_pair(cb_idx)
            if ob1 == 0:
                return Result(False, "CheckBox 定義が見つかりません（cbIdx=%s）"
                              % cb_idx, "error")
            state.selected_cb = cb_idx
            state.named_cb = False

        # --- OptionButton 全解除 ---
        state.selected_ob = 0

        # --- ラベル初期化 ---
        state.lbl_kensa_no = ""
        state.lbl_size1 = ""
        state.lbl_weight1 = ""
        state.lbl_size2 = ""
        state.lbl_weight2 = ""

        ken1, wt1 = self._read_sheet_header(ob1)
        state.lbl_kensa_no = ken1
        # B-3: 空なら "kg" だけを出さない（上記と同じ）
        state.lbl_weight1 = (wt1 + "kg") if wt1 != "" else ""
        state.lbl_size1 = self._size_label(ob1)

        _ken2, wt2 = self._read_sheet_header(ob2)
        state.lbl_weight2 = (wt2 + "kg") if wt2 != "" else ""
        state.lbl_size2 = self._size_label(ob2)

        log.info("[SelectCheckBox] ob1=%s ob2=%s", ob1, ob2)
        self.save_state(state)
        return Result(True, "", "info",
                      {"state": state.to_dict(), "ob1": ob1, "ob2": ob2})

    # ============================================================ 重量反映
    def apply_weight(self, state: FormState) -> Result:
        """VBA ``CommandButton1_Click``（重量反映）。

        判定順も VBA のまま:
            検査番号 -> InpCheck(1) -> InpCheck(2) -> InpCheck(3)
            -> チェック入力（CBモード） -> ExecuteWrite（OBモード）
        """
        ken = V.normalize_kensa_no(state.kensa_no)

        msg = V.validate_kensa_no(ken)
        if msg:
            return Result(False, msg, "warn")

        msg = V.check_weight_inputs(state)
        if msg:
            return Result(False, msg, "warn")

        # 次のロット(前に計算したときと違う検番)の重量反映なら、前のロットの NW・GW・高さと
        # 台紙の本数欄を消す(統合 1.2.35・現場の指示)。小ラベルを刷ったあとは消さずに残し
        # (刷り直せるように)、消すのは次のロットの処理が走ったとき。
        # VBA は刷ったあと「いいえ」でも必ずクリアしていた
        # 消すのは台紙へ書くと決まってから(サイズ未選択などで書かずに戻るときは何も消さない)
        saved = self.store.get_kv(self._CALC_INPUT_KEY, None)
        lot_changed = bool(saved and saved.get("kensaNo") and saved.get("kensaNo") != ken)

        written: List[dict] = []

        # --- CB モード（VBA チェック入力）---
        if state.is_cb_mode:
            if state.named_cb:
                ob1, ob2 = self.master.pair_for_named_cb()
            else:
                ob1, ob2 = self.master.cb_to_pair(state.selected_cb)
            log.info("[チェック入力] ob1=%s ob2=%s", ob1, ob2)
            if lot_changed:
                self.clear_all(state, save=False)

            written.append(self._write_label_sheet(ob1, ken, state.weight1, 1))
            cfg1 = self.master.get(ob1)
            state.lbl_kensa_no = self.store.get_cell(ob1, cfg1.data_row,
                                                    SM.HDR_KEN_COL)
            w1 = self.store.get_cell(ob1, cfg1.data_row, SM.HDR_WT1_COL)
            state.lbl_weight1 = w1 + "kg"
            state.lbl_size1 = self._size_label(ob1)

            written.append(self._write_label_sheet(ob2, ken, state.weight2, 2))
            cfg2 = self.master.get(ob2)
            w2 = self.store.get_cell(ob2, cfg2.data_row, SM.HDR_WT1_COL)
            state.lbl_weight2 = w2 + "kg"
            state.lbl_size2 = self._size_label(ob2)

        # --- OB モード（VBA ExecuteWrite）---
        elif state.is_ob_mode:
            ob = state.selected_ob
            cfg = self.master.get(ob)
            if cfg is None:
                # 足したサイズの表が読めない・知らない番号が来た。500 で落とさずに理由を返す
                return Result(False, "選んだサイズの定義が見つかりません。サイズを選び直してください",
                              "warn")
            take = cfg.take_type
            wt = state.weight1 if take == 1 else state.weight2
            if lot_changed:
                self.clear_all(state, save=False)

            written.append(self._write_label_sheet(ob, ken, wt, take))

            state.lbl_kensa_no = self.store.get_cell(ob, cfg.data_row,
                                                    SM.HDR_KEN_COL)
            w = self.store.get_cell(ob, cfg.data_row, SM.HDR_WT1_COL)
            if take == 1:
                state.lbl_weight1 = w + "kg"
                state.lbl_size1 = self._size_label(ob)
            else:
                state.lbl_weight2 = w + "kg"
                state.lbl_size2 = self._size_label(ob)
        else:
            # VBA も何も起きない（DetectSelectedOB=0 かつ CB 未選択）
            return Result(True, "サイズが選択されていません（転記なし）", "warn",
                          {"state": state.to_dict(), "written": []})

        state.kensa_no = ken
        # 検査番号(= ロット番号)のコイル外径から、近いチップボールを選ぶ(統合 1.2.35・現場の指示)。
        # 資材計算の仕掛パスが実績の形式のときだけ(`services/coil_outer.py`)
        try:
            tip, note = coil_outer.choose(ken, state.tip)
        except Exception:                               # noqa: BLE001 - 選べなくても反映は止めない
            log.warning("コイル外径からチップボールを選べませんでした", exc_info=True)
            tip, note = None, ""
        if tip:
            state.tip = tip
        state.tip_note = note
        self.save_state(state)
        message = "重量を反映しました"
        if lot_changed:
            message += "\n検査番号が変わったので、前のロットの NW・GW・高さと台紙の本数欄を消しました"
        if note:
            message += "\n" + note
        return Result(True, message, "info",
                      {"state": state.to_dict(), "written": written})

    def _write_label_sheet(self, ob_idx: int, ken: str, wt: str,
                           take_type: int) -> dict:
        """VBA ``WriteAllToSheet``。台紙 1 枚分を DB へ書く。"""
        cfg = self.master.get(ob_idx)
        cells = LL.build_label_cells(ken, wt, cfg.kataban, take_type)
        self.store.put_cells(ob_idx, cells)
        self.store.set_sheet_header(ob_idx, ken, wt, cfg.kataban)
        log.info("[WriteAllToSheet] sheet=%s ken=%s wt=%s kataban=%s take=%s",
                 cfg.sheet_name, ken, wt, cfg.kataban, take_type)
        return {"obIdx": ob_idx, "sheetName": cfg.sheet_name,
                "kensaNo": ken, "weight": wt, "kataban": cfg.kataban,
                "takeType": take_type, "cellCount": len(cells)}

    # ============================================================ 風袋計算
    def calc_tare(self, state: FormState) -> Result:
        """VBA ``CommandButton16_Click``（重量計算_DB）。

        段ごとに `self.progress` を進める(画面の進み具合の棒)。計算の中身は変えない。
        """
        coils = [n for n in (4, 3, 2, 1) if str(state.coil_h.get(n, "") or "").strip()]
        # 資材マスタの確かめ + 梱包ごと + ラベルへ写す + 保存
        self.progress.begin("重量計算", 3 + len(coils))
        try:
            return self._calc_tare(state)
        finally:
            self.progress.end()

    def _calc_tare(self, state: FormState) -> Result:
        # VBA: IsAllCoilEmpty なら何もしない
        if state.all_coil_empty():
            return Result(True, "本数が入力されていません", "warn",
                          {"state": state.to_dict()})

        # 本数を入れた丈の重量(重量反映でラベルへ写した値)が数でなければ止める(統合 1.2.31)。
        # 「-」だけで重量反映すると NW=0 のまま計算・印刷まで通った(VBA は CDbl で止まった)。
        # 消す前に確かめる(止めたときに、前の計算結果を消さない)
        msg = self._label_weight_problem(state)
        if msg:
            return Result(False, msg, "warn", {"state": state.to_dict()})

        # 設定(チップボール・丈・サイズ)と資材マスタを、**消す前に**確かめる(統合 1.2.35)。
        # VBA は クリア → 風袋シートの作り直し → IsValidSettings の順で、止めると台紙の本数欄だけ
        # 消えて、画面に残した NW・GW・高さと食い違った(開き直すと古い GW が出た)。
        # 止めるときは何も消さない(前の計算結果のまま。入力が変わっていれば印刷は止まる)
        msg = V.validate_settings(state)
        if msg is not None:
            return Result(False, msg, "warn", {"state": state.to_dict()})

        self.progress.step("資材マスタを確かめています")
        try:
            # 「計算」を押した瞬間は **必ずマスタを確かめる**。
            # 鮮度チェックの間隔（既定 300 秒）に隠れて、
            # 別の PC で直した単重が反映されないまま刷られるのを防ぐ。
            table = self.materials.load(check_now=True)
        except AccessError as exc:
            # VBA: MsgBox "資材重量なし", vbCritical
            return Result(False, "資材重量なし\n%s" % exc, "error",
                          {"state": state.to_dict()})
        if table.is_empty():
            return Result(False, "資材重量なし", "error",
                          {"state": state.to_dict()})

        width = self._selected_width(state)
        cb_key, ob_idx = self._selection_keys(state)
        t1 = TC.parse_label_weight(state.lbl_weight1)
        t2 = TC.parse_label_weight(state.lbl_weight2)

        missing: List[str] = []
        computed = []

        # VBA は 4 -> 3 -> 2 -> 1 の順に処理する。先に全部計算してから(書くのは確かめたあと)。
        # **隠れている丈の梱包は計算しない**(統合 1.2.35)。丈を切り替えても本数の欄は値が残るので、
        # 丈1・丈2 同時 → 丈1 だけ に切り替えると、丈2 の梱包(NW=0)が風袋の紙に出ていた
        # (VBA は印刷範囲で表示中の丈だけを刷っていた)
        for coil_no in self.visible_slots(state)[::-1]:
            raw = str(state.coil_h.get(coil_no, "") or "").strip()
            if raw == "":
                continue
            coil_text = val(raw)
            self.progress.step(f"{TARE_TITLE[coil_no].rstrip(':')} を計算しています")

            mat = TC.calculate_materials(
                table, coil_text, coil_no, width, state.tip, cb_key, ob_idx,
                t1, t2, state.take1_visible, state.take2_visible,
                all_coil_empty=False)
            for name in mat.missing:
                if name not in missing:
                    missing.append(name)
            computed.append((coil_no, coil_text, mat))

        if not computed:
            return Result(False, "表示中の丈に本数が入力されていません", "warn",
                          {"state": state.to_dict()})

        # 資材マスタに無い資材・単重が数でない資材があれば止める(統合 1.2.35。VBA も止まった。
        # 以前は 0kg として計算を続け、GW が軽いラベルが刷れた)
        if missing:
            return Result(False, self.missing_message(missing, self.materials.describe()),
                          "error", {"state": state.to_dict(), "missing": missing})

        # VBA: クリア()
        self.clear_all(state, save=False)
        # VBA: RecreateSheet("風袋計算")
        self.store.clear_tare()

        results = {}
        for coil_no, coil_text, mat in computed:
            # --- 画面へ（VBA SetCoilFormValues / SetCoilDimensions）---
            nw_s, gw_s = TC.form_values(mat)
            state.nw[coil_no] = nw_s
            state.gw[coil_no] = gw_s
            state.ta[coil_no] = fmt(ws_roundup(mat.KTA, -1), "0.0")

            payload = self._tare_payload(coil_no, coil_text, mat, t1, t2)
            self.store.set_tare(coil_no, payload)
            results[coil_no] = payload

        # VBA: ラベル転記CK / ラベル転記OP
        self.progress.step("ラベルへ写しています")
        transferred = self._transfer_coil_info(state)

        # この計算に使った入力を控えておく（印刷時の食い違い検出用）
        self.progress.step("保存しています")
        self.store.set_kv(self._CALC_INPUT_KEY, self._calc_fingerprint(state))
        # 風袋計算の結果(画面・印刷ビュー)がどの入力のものか。**クリアでは消さない** ──
        # 風袋計算の結果はクリアでも残るので、クリアのあと別の検番を反映して刷ると
        # 新しい見出しに前の結果が並ぶ(統合 1.2.26・構造点検で見つかった)
        self.store.set_kv(self._TARE_INPUT_KEY, self._calc_fingerprint(state))
        # どの取得元の、いつ時点の資材重量で計算したか(統合 1.2.31・現場の指示)。
        # 印刷の記録にもこれを写す(刷ったときではなく、計算したときの資材重量で刷るため)
        material = self.materials.describe()
        material["calcAt"] = time.strftime("%Y-%m-%d %H:%M:%S")
        self.store.set_kv(self._TARE_MATERIAL_KEY, material)

        self.save_state(state)
        coils, off_label = self._split_label_coils(
            {str(k): {"本数": str(state.coil_h.get(k, "") or ""),
                      "NW": state.nw.get(k, ""), "GW": state.gw.get(k, "")}
             for k in (1, 2, 3, 4) if str(state.coil_h.get(k, "") or "").strip()},
            self.visible_slots(state))
        self._log(self.KIND_CALC, "風袋計算", {
            "kensaNo": state.lbl_kensa_no,
            "sizes": [x for x in (state.lbl_size1, state.lbl_size2) if x],
            "weights": [x for x in (state.lbl_weight1, state.lbl_weight2) if x],
            "coils": coils, "notOnLabel": off_label,
            "material": material})

        return Result(True, "計算しました", "info", {
            "state": state.to_dict(),
            "tare": {str(k): v for k, v in results.items()},
            "transferred": transferred,
            "missing": missing,
            "materialSource": self.materials.last_source,
        })

    @staticmethod
    def visible_slots(state: FormState) -> List[int]:
        """表示中の丈の梱包の番号(丈1 → 1・2、丈2 → 3・4)。"""
        out: List[int] = []
        if state.take1_visible:
            out += [1, 2]
        if state.take2_visible:
            out += [3, 4]
        return out

    @staticmethod
    def missing_message(missing: List[str], material: dict) -> str:
        """資材マスタに無い・数でない資材で止めたときの文(どこを直せばよいかまで)。"""
        where = material.get("label") or material.get("source") or ""
        path = material.get("path") or ""
        return ("資材重量なし: 資材マスタに無い資材、または単重・係数が数でない資材があるので、"
                "計算を止めました。\n" + "、".join(missing)
                + (f"\n読んだ資材マスタ: {where}" if where else "")
                + (f"({path})" if path else "")
                + "\n資材マスタを直してから、もう一度計算してください")

    @staticmethod
    def _label_weight_problem(state: FormState) -> Optional[str]:
        """本数を入れた丈の、ラベルの重量(重量反映で写した値)が数でなければ理由。"""
        for take, coils, visible, label in (
                (1, (1, 2), state.take1_visible, state.lbl_weight1),
                (2, (3, 4), state.take2_visible, state.lbl_weight2)):
            if not visible:
                continue
            if not any(str(state.coil_h.get(n, "") or "").strip() for n in coils):
                continue
            if V.weight_number(label) is None:
                shown = str(label or "").replace("kg", "").strip() or "空"
                return (f"丈{take}の重量が数になっていません(「{shown}」)。"
                        "重量を数で入れて「重量反映」をやり直してください")
        return None

    def _tare_payload(self, coil_no: int, coil_text: float, mat, t1, t2) -> dict:
        """風袋計算シート 1 ブロック分（VBA WriteToSheet 相当）。"""
        disp = TC.display_values(mat)
        formulas = TC.sheet_formulas(mat, coil_text, coil_no, t1, t2)
        start = TARE_START_ROW[coil_no]
        values = [disp["PI"], disp["TIP"], disp["SUT"], disp["ESA"], disp["POR"],
                  disp["HB"], disp["PAR"], disp["PET"], disp["DRY"],
                  disp["NW"], disp["HU"], disp["GW"]]
        keys = ["PIN", "TIP", "SUT", "ESA", "POR", "HB", "PAR", "PET", "DRY",
                "NW", "HU", "GW"]
        rows = [{"row": start + 1 + i, "name": TARE_ITEMS[i],
                 "value": values[i], "formula": formulas[keys[i]]}
                for i in range(len(TARE_ITEMS))]
        return {
            "coilNo": coil_no,
            "startRow": start,
            "title": f"{TARE_TITLE[coil_no]}{fmt(coil_text, '0')}本",
            "coilText": fmt(coil_text, "0"),
            "rows": rows,
            "display": disp,
            "breakdown": mat.formulas,
            # 丸める前の 風袋・GW(図で「36.948 → 37.0」と丸めの前後を出す。統合 1.2.21)
            "huRaw": fmt(mat.HU, "0.000"),
            "gwRaw": fmt(mat.GW, "0.000"),
            "taDisplay": fmt(ws_roundup(mat.KTA, -1), "0.0"),
            "coilHeightMm": fmt(mat.TA * 1000, "0.0"),
            "gai": fmt(mat.GAI, "0.000"),
            "tempDiameter": mat.TempDiameter,
            "halfDiameter": mat.HalfDiameter,
            "width": fmt(mat.col, "0.0"),
            "hbFlag": mat.formulas.get("_HBFlag", ""),
            "hbTa": mat.formulas.get("_HB_TA", ""),
            # 胴巻きハードボードの高さ。締結代（HB_BAND_CLEARANCE_M）を
            # 入れた場合に積高と差が出るので、同じ単位で差まで出す。
            # 丸めた表示値から引くと差がずれるので、生の値から求める。
            "coilHeightHbMm": fmt(mat.HB_TA * 1000, "0.0"),
            "coilHeightDiffMm": fmt((mat.TA - mat.HB_TA) * 1000, "0.0"),
            "missing": mat.missing,
        }

    # ============================================================ ラベル転記
    def _transfer_coil_info(self, state: FormState) -> List[dict]:
        """VBA ``ラベル転記CK`` / ``ラベル転記OP``。本数/高さ/NW/GW を台紙へ。"""
        out = []

        if state.is_cb_mode:
            if state.named_cb:
                ob1, ob2 = self.master.pair_for_named_cb()
            else:
                ob1, ob2 = self.master.cb_to_pair(state.selected_cb)
            if ob1 == 0:
                return out
            if state.take1_visible:
                out.append(self._put_coil_info(ob1, state, 1, 2))
            if state.take2_visible:
                out.append(self._put_coil_info(ob2, state, 3, 4))
            return out

        if state.is_ob_mode:
            ob = state.selected_ob
            take = self.master.take_type(ob)
            s1, s2 = (1, 2) if take == 1 else (3, 4)
            out.append(self._put_coil_info(ob, state, s1, s2))
        return out

    def _put_coil_info(self, ob_idx: int, state: FormState,
                       s1: int, s2: int) -> dict:
        cfg = self.master.get(ob_idx)
        cells = LL.build_coil_info_cells(
            cfg.data_row, cfg.extra_row,
            state.coil_h.get(s1, ""), state.coil_h.get(s2, ""),
            state.ta.get(s1, ""), state.ta.get(s2, ""),
            state.nw.get(s1, ""), state.nw.get(s2, ""),
            state.gw.get(s1, ""), state.gw.get(s2, ""))
        self.store.put_cells(ob_idx, cells)
        return {"obIdx": ob_idx, "sheetName": cfg.sheet_name,
                "slots": [s1, s2], "cellCount": len(cells)}

    # ============================================================ クリア
    def clear_all(self, state: FormState, save: bool = True) -> Result:
        """VBA ``クリア()``。フォームの NW/GW/TA と 全台紙の本数列をクリア。"""
        for i in (1, 2, 3, 4):
            state.nw[i] = ""
            state.gw[i] = ""
            state.ta[i] = ""

        for ob in self.master.ob_indexes():
            cfg = self.master.get(ob)
            targets = LL.clear_coil_info_cells(cfg.data_row, cfg.extra_row)
            self.store.clear_cells(ob, targets)

        # 本数・高さ・NW・GW を消したので、計算済みの控えも消す
        self.store.set_kv(self._CALC_INPUT_KEY, None)

        if save:
            self.save_state(state)
        return Result(True, "クリアしました", "info", {"state": state.to_dict()})

    # ============================================================ 羅列計算
    def calc_list(self, state: FormState) -> Result:
        """VBA ``羅列計算``（1〜50本の一覧）。"""
        if not state.any_size_selected():
            return Result(False, "コイルサイズが選択されていません", "warn")
        if state.tip not in TIP_CHOICES:
            return Result(False, "チップボールチェックがありません", "warn")
        # 表示中の丈のラベルの重量が数でなければ止める(統合 1.2.35。風袋計算と同じ。
        # 羅列は 1〜50 本の NW・GW を並べるので、重量が数でないと NW=0 の一覧になる)
        msg = self._list_weight_problem(state)
        if msg:
            return Result(False, msg, "warn")

        try:
            # 「計算」を押した瞬間は **必ずマスタを確かめる**。
            # 鮮度チェックの間隔（既定 300 秒）に隠れて、
            # 別の PC で直した単重が反映されないまま刷られるのを防ぐ。
            table = self.materials.load(check_now=True)
        except AccessError as exc:
            return Result(False, "資材重量なし\n%s" % exc, "error")
        if table.is_empty():
            return Result(False, "資材重量なし", "error")

        # VBA GetSelectedCoilWidth はシート名から巾を取る（ck_08_53 は 53.5 固定）
        width = self._list_width(state)
        t1 = TC.parse_label_weight(state.lbl_weight1)
        t2 = TC.parse_label_weight(state.lbl_weight2)

        cb_key, ob_idx = self._selection_keys(state)
        data = LC.build_list(table, width, state.tip, t1, t2,
                             state.take1_visible, state.take2_visible,
                             cb_key=cb_key, ob_idx=ob_idx)
        # 資材マスタに無い資材・単重が数でない資材があれば止める(統合 1.2.35。風袋計算と同じ)
        if data["missing"]:
            return Result(False, self.missing_message(data["missing"], self.materials.describe()),
                          "error", {"missing": data["missing"]})
        data["width"] = fmt(width, "0.0")
        data["footer"] = self._list_footer(state)
        # どの取得元の、いつ時点の資材重量で計算したか(統合 1.2.31)
        data["material"] = dict(self.materials.describe(),
                                calcAt=time.strftime("%Y-%m-%d %H:%M:%S"))
        self._log(self.KIND_CALC, "羅列計算", {"width": data["width"], "material": data["material"]})
        self.store.set_kv("list_result", data)

        return Result(True, "羅列計算を作成しました", "info",
                      {"list": data, "materialSource": self.materials.last_source})

    @staticmethod
    def _list_weight_problem(state: FormState) -> Optional[str]:
        """羅列に使う丈(表示中の丈)の、ラベルの重量が数でなければ理由。"""
        for take, visible, label in ((1, state.take1_visible, state.lbl_weight1),
                                     (2, state.take2_visible, state.lbl_weight2)):
            if visible and V.weight_number(label) is None:
                shown = str(label or "").replace("kg", "").strip() or "空"
                return (f"丈{take}の重量が数になっていません(「{shown}」)。"
                        "重量を数で入れて「重量反映」をやり直してください")
        return None

    def _list_footer(self, state: FormState) -> str:
        """VBA ``FormatListSheet`` の左フッター。"""
        s1, s2 = state.lbl_size1, state.lbl_size2
        if s1 and s2:
            return f"＊検査番号：{state.lbl_kensa_no}＊ｺｲﾙｻｲｽﾞ：{s1} {s2}"
        if s1:
            return f"＊検査番号：{state.lbl_kensa_no}＊ｺｲﾙｻｲｽﾞ：{s1}"
        if s2:
            return f"＊検査番号：{state.lbl_kensa_no}＊ｺｲﾙｻｲｽﾞ：{s2}"
        return ""

    # ============================================================ 巾
    def _selected_width(self, state: FormState) -> float:
        """VBA ``SetCoilDimensions`` の巾決定（ハードコード表）。

        VBA は CB 側 Select Case のあとに OB 側 Select Case を評価するため、
        OB が選択されていれば OB が優先される。同じ順序で判定する。
        """
        width = 0.0
        if state.selected_cb:
            for ob in self.master.ob_indexes():
                cfg = self.master.get(ob)
                if cfg.cb_key == f"CheckBox{state.selected_cb}":
                    width = cfg.width_mm
                    break
        elif state.named_cb:
            ob1, _ = self.master.pair_for_named_cb()
            width = self.master.width_of(ob1)
        if state.selected_ob:
            width = self.master.width_of(state.selected_ob)
        return width

    def _list_width(self, state: FormState) -> float:
        """VBA ``GetSelectedCoilWidth``（シート名から抽出、ck_08_53 は 53.5 固定）。"""
        if state.selected_ob:
            return self.master.width_from_sheet_name(state.selected_ob)
        if state.selected_cb:
            ob1, _ = self.master.cb_to_pair(state.selected_cb)
            return self.master.width_from_sheet_name(ob1)
        if state.named_cb:
            return 53.5          # VBA のハードコード
        return 0.0

    def _selection_keys(self, state: FormState) -> Tuple[str, int]:
        """``judge_hb_flag`` へ渡す (CheckBox名, obIdx)。"""
        if state.named_cb:
            return SM.NAMED_CB_KEY, 0
        if state.selected_cb:
            return f"CheckBox{state.selected_cb}", 0
        return "", state.selected_ob

    # ============================================================ 印刷対象
    def print_targets(self, state: FormState) -> Result:
        """VBA ``CommandButton2_Click``。印刷対象の台紙を決める。"""
        if state.is_ob_mode:
            # 確認の文言も、ラベルに刷られる名前（マスタで直せる方）で出す
            cfg = self.master.get(state.selected_ob)
            name = cfg.size_cell_label if cfg else ""
            return Result(True, f"{name} を印刷しますか？", "info",
                          {"targets": [state.selected_ob], "displayName": name})

        if state.is_cb_mode:
            if state.named_cb:
                ob1, ob2 = self.master.pair_for_named_cb()
                name = self.master.display_name(ob1) + " 　丈1,2"
            else:
                ob1, ob2 = self.master.cb_to_pair(state.selected_cb)
                name = self.master.display_name(ob1) + " 　丈1,2"
            return Result(True, f"{name} を印刷しますか？", "info",
                          {"targets": [ob1, ob2], "displayName": name})

        return Result(False, "サイズが選択されていません", "warn", {"targets": []})

    #: 本数・高さ・NW・GW に影響する入力（これが変わったら再計算が必要）
    _CALC_INPUT_KEY = "tare_inputs"
    #: 風袋計算の結果を出したときの入力(クリアでは消さない。`tare_is_stale` が使う)
    _TARE_INPUT_KEY = "tare_result_inputs"
    #: 風袋計算(とラベルの本数・高さ・NW・GW)を出したときの資材重量の素性(統合 1.2.31)
    _TARE_MATERIAL_KEY = "tare_material"

    # ============================================================ 計算・印刷の記録
    def _log(self, kind: str, title: str, payload: dict) -> None:
        """記録を残す。残せなくても計算・印刷は止めない(ログに理由を出す)。"""
        try:
            # どの版で・どの PC(ライン)で、を添える(統合 1.2.39。切り替えの前か後か・どの版で刷ったかを
            # 記録から言えるように)
            payload = dict(payload, version=WR.versions(), who=WR.who(self.cfg))
            self.store.add_log(kind, title, payload)
            # 消えない控え(この PC の月ごとのファイルと、共有ログの出力先)。書けなくても止めない
            WR.append(self.cfg, {"at": time.strftime("%Y-%m-%d %H:%M:%S"), "kind": kind,
                                 "title": title, **payload})
            m = payload.get("material") or {}
            log.info("[記録] %s %s 検番=%s 資材重量=%s %s時点%s", kind, title,
                     payload.get("kensaNo", ""), m.get("label", "-"), m.get("asOf", ""),
                     "(古い内容)" if m.get("stale") else "")
        except Exception:                                   # noqa: BLE001
            log.exception("計算・印刷の記録を残せませんでした: %s %s", kind, title)

    #: 印刷の記録を残す画面(入口から見た道 → 呼び名)
    PRINT_KINDS = {
        "/labels/print": "小ラベル", "/labels/sheet": "ラベル台紙",
        "/tare/print": "風袋計算", "/tare": "風袋計算",
        "/list/print": "羅列計算", "/list": "羅列計算",
        "/all-size/print": "全サイズ",
    }

    #: 記録の種類(統合 1.2.39 で名前を正した)
    KIND_CALC = "計算"
    #: 印刷の画面(ブラウザの印刷ダイアログ)を開いて閉じた。**刷った証拠にはならない**
    #: (ブラウザは、刷ったか・やめたかを教えない)。以前は「印刷」と書いていた
    KIND_PRINT = "印刷画面"
    #: 入力が変わったまま(計算し直していない)で、印刷を止めた。混ざったラベルを止めた証拠
    KIND_STOPPED = "止めた印刷"

    def record_print(self, path: str, query: Dict[str, List[str]],
                     stopped: bool = False) -> Result:
        """印刷の画面を開いたこと・止めた印刷を残す(統合 1.2.31・現場の指示。1.2.39 で止めた印刷も)。

        印刷の画面(ブラウザの印刷ダイアログ)を閉じたときに、画面から知らせてくる。
        ブラウザは「刷ったか・やめたか」を教えないので、**印刷の画面を開いたこと**を残す
        (刷った証拠にはならない)。``stopped`` は、入力が変わったまま印刷ビューを開き、
        「印刷（停止中）」で断ったとき(サーバが印刷ビューを描くときに残す)。
        資材重量は、その紙の値を計算したときのもの(計算の記録と同じ)を写す。
        """
        title = self.PRINT_KINDS.get(path)
        if not title:
            return Result(False, "記録しない画面です", "info")
        payload: dict = {"path": path}
        if stopped:
            payload["stopReason"] = ("入力が変わったまま「計算」し直していないので、印刷を止めた"
                                     "(前の計算の NW・GW と新しい検番が混ざるのを防いだ)")
        if path.startswith("/all-size"):
            combo = (query.get("combo") or [""])[0]
            st = self.all_size.load_state()
            payload.update(combo=combo, kensaNo=st.get("kensaNo", ""),
                           weights=[w for w in (st.get("w1"), st.get("w2"),
                                                st.get("w4"), st.get("w5")) if w],
                           material={"label": "使わない(重量は打った値)"})
        elif path.startswith("/list"):
            data = self.store.get_kv("list_result") or {}
            payload.update(width=data.get("width", ""), material=data.get("material") or {})
        else:
            state = self.load_state()
            payload["kensaNo"] = state.lbl_kensa_no
            payload["material"] = self.store.get_kv(self._TARE_MATERIAL_KEY, {}) or {}
            if path.startswith("/labels"):
                obs = [int(x) for x in ",".join(query.get("ob") or []).split(",")
                       if x.strip().isdigit()]
                sheets = []
                for ob in obs:
                    v = self.label_view(ob)
                    if not v:
                        continue
                    ph = v.get("printHeader") or {}
                    # 本数の入った梱包だけ(本数が空の段は「NW 0 の梱包」に見えるので入れない。統合 1.2.39)
                    infos = [c for c in v.get("coilInfo") or [] if str(c.get("coilH", "") or "").strip()]
                    sheets.append({"size": ph.get("sizeName", ""), "kensaNo": ph.get("kensaNo", ""),
                                   "weight": ph.get("weight", ""), "labels": len(v.get("labels") or []),
                                   "NW": [c.get("nw", "") for c in infos],
                                   "GW": [c.get("gw", "") for c in infos]})
                payload["sheets"] = sheets
                payload["inputChanged"] = self.calc_is_stale()
            else:
                coils = {str(k): {"本数": v.get("coilText", ""),
                                  "NW": (v.get("display") or {}).get("NW", ""),
                                  "GW": (v.get("display") or {}).get("GW", "")}
                         for k, v in (self.store.all_tare() or {}).items()}
                payload["coils"], payload["notOnLabel"] = self._split_label_coils(
                    coils, self.visible_slots(state))
                payload["inputChanged"] = self.tare_is_stale(state)
        self._log(self.KIND_STOPPED if stopped else self.KIND_PRINT, title, payload)
        return Result(True, "記録しました", "info")

    @staticmethod
    def _split_label_coils(coils: Dict[str, dict], visible: List[int]):
        """ラベルに載る梱包と載らない梱包に分ける(統合 1.2.39・現場の指摘)。

        載らない梱包(隠れている丈・NW が 0 か空)を記録に混ぜると「NW 0 の梱包」に見えて誤解される。
        """
        on, off = {}, {}
        for k, c in coils.items():
            nw = V.weight_number(str(c.get("NW", "") or ""))
            (on if int(k) in visible and nw is not None else off)[k] = c
        return on, off

    def recent_log(self, limit: int = 300) -> List[dict]:
        return self.store.recent_log(limit)

    @staticmethod
    def _calc_fingerprint(state: FormState) -> dict:
        """風袋計算の結果を左右する入力だけを取り出す。

        ラベルの **本数・高さ・NW・GW** は「計算」を押したときにしか
        書き換わらない（VBA も同じ）。一方 **検番・重量** は「重量反映」で
        書き換わる。そのため入力を変えて反映だけして印刷すると、
        **新しい検番と古い NW/GW が混ざったラベル**が出てしまう。
        値そのものは VBA どおりのままにし、**食い違いを知らせる**ために使う。
        """
        return {
            "ob": state.selected_ob,
            "cb": state.selected_cb,
            "namedCb": bool(state.named_cb),
            "tip": state.tip,
            "kensaNo": V.normalize_kensa_no(state.kensa_no),
            "weight1": state.weight1 or "",
            "weight2": state.weight2 or "",
            "coilH": {str(i): str(state.coil_h.get(i, "") or "").strip()
                      for i in (1, 2, 3, 4)},
        }

    def calc_is_stale(self, state: Optional[FormState] = None) -> bool:
        """いまの入力が、最後に計算したときと違うか。

        一度も計算していない場合は False（まだ古い値が載っていないため）。
        """
        saved = self.store.get_kv(self._CALC_INPUT_KEY, None)
        if not saved:
            return False
        st = state if state is not None else self.load_state()
        return self._calc_fingerprint(st) != saved

    def tare_is_stale(self, state: Optional[FormState] = None) -> bool:
        """風袋計算の結果が、いまの入力のものでないか(統合 1.2.26)。

        風袋計算の画面・印刷ビューは、見出し(検番・サイズ)をいまの入力から、中身を前の
        計算から出す。入力を変えて重量反映だけすると、**新しい検番に前の風袋・GW が並ぶ**。
        ラベルの印刷(`calc_is_stale`)と同じ決まりで止める。

        ラベル側の控えはクリアで消えるが、風袋計算の結果はクリアでも残るので、
        こちらは消さない控え(`_TARE_INPUT_KEY`)と比べる。控えが無い(この版より前に
        計算した)ときはラベル側の控えで判断する。
        """
        saved = self.store.get_kv(self._TARE_INPUT_KEY, None)
        if not saved:
            return self.calc_is_stale(state)
        st = state if state is not None else self.load_state()
        return self._calc_fingerprint(st) != saved

    def label_view(self, ob_idx: int) -> dict:
        # ラベル台紙を開いたときは **必ず読み直す**。
        #
        # 間隔を置いてキャッシュすると、別の PC で直した内容が
        # そのぶん遅れて出る。刷ってから気付くのでは遅いので、
        # ここは取りこぼしの無さを優先する。
        # 読むのは 14 行の小さな表で、開くのも印刷のときだけ。
        self.refresh_label_sheets(force=True)
        """台紙 1 枚分の表示データ。印刷ビュー・確認画面で使う。"""
        cfg = self.master.get(ob_idx)
        if not cfg:
            return {}
        header = self.store.get_sheet_header(ob_idx)
        cells = self.store.get_cells(ob_idx)
        ken = header["kensaNo"]
        wt = header["weight"]
        # 型番は **いまのマスタ** を優先する。
        # header["kataban"] は重量反映のときに控えた値で、
        # そのあとマスタを直しても古いままになってしまう。
        rows = LL.build_label_rows(ken, wt, cfg.kataban or header["kataban"],
                                   cfg.take_type) if ken or wt else []
        coil_info = []
        for r in (cfg.data_row, cfg.data_row + 1):
            coil_info.append({
                "row": r,
                "coilH": cells.get((r, SM.COL_COILH), ""),
                "ta": cells.get((r, SM.COL_TA), ""),
                "nw": cells.get((r, SM.COL_NW), ""),
                "gw": cells.get((r, SM.COL_GW), ""),
            })
        # 印刷ヘッダー帯（VBA 台紙の行3 / 行92 の内容）
        take = cfg.take_type
        print_header = {
            "sizeName": cfg.size_cell_label,
            "kensaNo": ken,
            "weight": LL.weight_text(wt),          # 台紙のセル書式 0.0(統合 1.2.6)
            "weightLabel": "重量 丈%d" % take,
            "kataban": cfg.kataban or header["kataban"],
        }
        for n, ci in enumerate(coil_info, start=1):
            print_header["coilH%d" % n] = ci["coilH"]
            print_header["ta%d" % n] = ci["ta"]
            print_header["nw%d" % n] = ci["nw"]
            print_header["gw%d" % n] = ci["gw"]

        return {
            "config": cfg.to_dict(),
            "header": header,
            "printHeader": print_header,
            "labels": rows,
            "coilInfo": coil_info,
            "cellCount": len(cells),
            # 本数・高さ・NW・GW が現在の入力と食い違っていないか
            "stale": self.calc_is_stale(),
        }

    # ============================================================ その他
    def quick_height(self, state: FormState, coil_count: float) -> Result:
        """VBA ``CommandButton13_Click``（積み高さクイック計算）。"""
        width = self._selected_width(state)
        override = (self.cfg.quick_height_width_override or {}).get(str(width))
        if override:
            width = float(override)
        if width <= 0:
            return Result(False, "コイルサイズが選択されていません", "warn")
        data = LC.quick_height(coil_count, width)
        data["width"] = fmt(width, "0.0")
        return Result(True, data["message"], "info", {"quick": data})
