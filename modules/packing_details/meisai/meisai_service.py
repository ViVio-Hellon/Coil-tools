"""出力の業務ロジック (VBA `cmdKettei_Click` の移植)

紙を1枚出すまでにVBAがやっていたことを、順番どおりに並べ直したもの。

    1. 重量を読み直す            ReadWeights
    2. 積み上げ順に条番号を並べる   slotKeys
    3. 副番の重複を確かめる        IsDuplicate
    4. Noを決める                IncrementSeq
    5. 明細を作る                CreateMeisaiSheet + WriteData
    6. 副番を登録する             RegisterFuban
    7. 印刷済フラグをOFFにする     SetPrinted False
    8. 積み上げだけリセットする     ClearGimmickC + BuildGimmickC

1 はルート(`routes/meisai.output`)が、押したときの画面の重量欄の値でする
(`read_weights`)。ここに来るときは `board.weights` が読み直したあとの値。

【No指定は外した(VER 0.13.23・現場の判断)】
出力済みを直したいときは、出力済みにチェックを入れて「出力済み修正」で積み直す
(`begin_edit` → `output_edit`)。同じNoで出し直し、前の紙の履歴には「作り直し」の
印を付ける。番号を手で打つ No指定は、打ち間違いで別の紙を上書きしうるので使わない。

【VBAが確認を挟む場所は、ここでは「断り」として返す】
VBA は `MsgBox ... vbYesNo` で利用者に尋ね、「はい」なら続けていた。
Webでは1往復で決められないので、**まず確認事項を返し、利用者が
了解したら `confirm=True` でもう一度呼ぶ**形にする。判断そのもの
(何を尋ねるか・尋ねずに進めてよいか)はVBAと同じ。
"""
from __future__ import annotations

import re
import sqlite3
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Iterable, Optional

from . import config, db, history_repo, lot_repo, slip_history, strand_service
from .logging_utils import get_logger
from .strand_service import Board

log = get_logger("meisai_service")

# 確認の種類。**文言から推し量らない**
CONFIRM_DUPLICATE = "duplicate"      # 副番がすでに出力済み
CONFIRM_EDIT = "edit"                # 出力済みのNoを積み直した並びで出し直す(VER 0.13.23)

# 断りの種類
REFUSE_INCOMPLETE = "incomplete"     # スロットが埋まっていない
REFUSE_NO_WEIGHT = "no_weight"       # 重量が未入力
REFUSE_TOO_MANY = "too_many"         # 紙の行数を超える
REFUSE_NOT_EDITING = "not_editing"   # 修正中ではない
REFUSE_NO_OUTPUT = "no_output"       # そのNoの出力は無い
REFUSE_USED_ELSEWHERE = "used_elsewhere"  # ほかのNoで使っているコイルを積んでいる


@dataclass
class Confirm:
    """利用者に尋ねること (VBA の `MsgBox ... vbYesNo` 1つ分)。"""

    kind: str
    message: str
    detail: list[str] = field(default_factory=list)


@dataclass
class OutputResult:
    """出力の結果。"""

    ok: bool = False
    seq_no: int = 0
    confirms: list[Confirm] = field(default_factory=list)
    refuse: str = ""
    message: str = ""

    @property
    def needs_confirm(self) -> bool:
        return bool(self.confirms)


# ==================================================================
# 重量 (RULE-06)
# ==================================================================
def read_weights(values: list[object], jou_su: int) -> list[int]:
    """重量欄を読む (VBA `ReadWeights`)。

    VBA は `CLng(Val(ctl.Value))` で読み、数値でなければ
    「丈Nの重量が未入力です。」で止める。入力欄が数字以外を弾く
    (`clsWeightCtrl.Txt_KeyPress`)ので、**小数点も入らない＝整数**。

    足りない欄・空欄は**未入力として断る**。0を入れて進めると、
    紙に `0.0Kg` と印字されたまま現場に出る。
    """
    out: list[int] = []
    for index in range(jou_su):
        raw = values[index] if index < len(values) else None
        # 全角の数字は半角にしてから読む(画面でも直している)。数字だけを整数として読み、
        # float を通さない(統合 1.2.12。以前は "2.5e3"・"1_000"・"2502." も通っていた)
        text = unicodedata.normalize("NFKC", str(raw if raw is not None else "")).strip()
        if text == "":
            raise strand_service.RefusedError(
                REFUSE_NO_WEIGHT, f"丈{index + 1}の重量が未入力です。")
        # 小数点以下が 0 だけ("250.0")なら同じ数なので通す
        whole = re.fullmatch(r"([0-9]+)(?:\.0*)?", text)
        if not whole:
            # **小数を黙って切り捨てない**(統合 1.2.6)。VBA の入力欄は数字以外を弾くので
            # 小数は入らない。ツールでも日本語入力(IME)からは小数点が入りえて、以前は
            # 2502.7 を 2502 にして紙・履歴(共有)へ書いていた。違う数を書くより断る
            raise strand_service.RefusedError(
                REFUSE_NO_WEIGHT,
                f"丈{index + 1}の重量は整数(数字だけ)で入力してください: {text}")
        number = int(whole.group(1))
        if number <= 0:
            # 0 のまま進めると紙に 0.0Kg と刷られて現場に出る
            raise strand_service.RefusedError(
                REFUSE_NO_WEIGHT, f"丈{index + 1}の重量が0です。")
        out.append(number)
    return out


# ==================================================================
# 出力 (RULE-07 / RULE-08)
# ==================================================================
def output(conn: sqlite3.Connection, board: Board, *,
           confirm: bool = False,
           lot: Any = None, orders: Iterable[Any] = ()) -> OutputResult:
    """明細を1枚出す (VBA `cmdKettei_Click`)。次の連番で出す。

    修正中(`board.editing`)は呼ばない ── `output_edit` を使う(呼び出し側が分ける)。

    `confirm=False` のうちは確認事項を返すだけで**何も書かない**。
    利用者が了解したら `confirm=True` で呼び直す。

    `lot` / `orders` は画面に出ているロット情報。**明細の履歴**に残す
    (`slip_history`。全ライン・3年。あとから「このコイルはどの紙か」を
    追う・数えるため)。
    """
    result = OutputResult()

    if not board.is_complete:
        result.refuse = REFUSE_INCOMPLETE
        result.message = "積み上げが埋まっていません。"
        return result

    keys = board.stacked_keys()
    if len(keys) > config.SHEET_ROWS:
        # VBA `WriteData` は16本目以降を黙って捨てるが、こちらは断る。
        # 積み条数の上限が15なのでふつうは起きない
        result.refuse = REFUSE_TOO_MANY
        result.message = (f"1枚に載るのは{config.SHEET_ROWS}行までです"
                          f"（いま{len(keys)}行）。")
        return result

    confirms: list[Confirm] = []

    # --- 副番の重複 ---
    dup = history_repo.duplicates(conn, board.lot_no, keys)
    if dup:
        confirms.append(Confirm(
            CONFIRM_DUPLICATE,
            "以下の副番はすでに出力済みです。このまま出力しますか？",
            dup))

    if confirms and not confirm:
        result.confirms = confirms
        result.message = "確認してください。"
        return result

    # --- ここから書く ---
    #
    # **ひとまとまりにする。** 文ごとに確定していくと、
    #   ・明細は残ったが副番が登録されず、次に同じコイルを出しても
    #     重複の警告が出ない
    #   ・連番だけ進んで明細が無い
    # といった半端な状態がありうる。まとめれば、全部書けるか
    # 1行も書かないかのどちらかになる。
    #
    # **書けなければ `db.WriteError` が飛ぶ。** 以前は書き込みの
    # 結果を捨てていたので、1行も書けていないのに画面には
    # 「出力しました」と出ていた。
    with db.transaction(conn, name="明細の出力"):
        seq_no = history_repo.next_seq(conn)

        # **履歴にも同じまとまりで書く。** 明細だけ残って履歴が無い(あるいは
        # その逆)を作らない
        history_id = slip_history.record_output(
            conn, lot_no=board.lot_no, seq_no=seq_no, keys=keys,
            weights=board.weights, lot=lot, orders=orders)
        db.write(
            conn,
            "INSERT INTO 明細出力 (ロット番号, No, 副番の並び, 重量, 出力日時, 履歴ID)"
            " VALUES (?, ?, ?, ?, ?, ?)",
            (board.lot_no, seq_no, ",".join(keys),
             ",".join(str(w) for w in board.weights), db.now_db_string(), history_id),
            name="明細の出力")

        # --- 副番を登録 (VBA: すでにあるものは足さない) ---
        for key in keys:
            full = history_repo.fuban_key(board.lot_no, key)
            if not history_repo.is_duplicate(conn, full):
                history_repo.register_fuban(conn, full)

        # --- 印刷済フラグをOFF ---
        history_repo.set_printed(conn, False)

    log.info("output: %s-No%s を出力しました (%s行)", board.lot_no, seq_no, len(keys))
    result.ok = True
    result.seq_no = seq_no
    result.message = f"出力しました。No{seq_no}を確認してください。"
    return result


# ==================================================================
# 出力済みの修正 (VER 0.13.23・現場の求め)
# ==================================================================
# 出力済みにチェックを入れて「出力済み修正」を押すと、そのNoのコイルを積み上げへ戻して
# 積み直せる。**ほかのNoで使っているコイルは使えない**(No1〜No3 があってNo1を直すとき、
# No2・No3 のコイルは積めない)。「修正して出力」で同じNoのまま出し直し、前の紙の履歴には
# 「作り直し」の印を付ける(消さない)。
def output_keys(conn: sqlite3.Connection, lot_no: str, *,
                exclude_no: int = 0) -> set[str]:
    """そのロットの出力済みに載っているコイル(`exclude_no` のNoを除く)。"""
    return {k for o in list_outputs(conn, lot_no) if o.seq_no != exclude_no for k in o.keys}


def begin_edit(conn: sqlite3.Connection, board: Board, seq_no: int) -> Output:
    """出力済みの1枚を積み直しに入る。断るときは `RefusedError`。"""
    found = find_output(conn, board.lot_no, seq_no)
    if found is None:
        raise strand_service.RefusedError(REFUSE_NO_OUTPUT, f"No{seq_no}の出力はありません。")
    if not board.strands:
        raise strand_service.RefusedError(
            REFUSE_NO_OUTPUT, "条番号が表示されていません。先に「条番号表示」を押してください。")
    board.begin_edit(found.seq_no, found.keys, found.weights)
    # ほかのNoのコイルは使用済みのまま(途中経過から戻した直後でも確実に)
    for key in output_keys(conn, board.lot_no, exclude_no=found.seq_no):
        if board.knows(key) and key not in board.restored_used and key not in board.slots:
            board.restored_used.append(key)
    return found


def output_edit(conn: sqlite3.Connection, board: Board, *, confirm: bool = False,
                lot: Any = None, orders: Iterable[Any] = ()) -> OutputResult:
    """修正中のNoを、積み直した並びで出し直す。**Noは変えない。**

    - 埋まっていなければ断る / ほかのNoのコイルが入っていれば断る(念のための最後の砦。
      積むときにも断っている)
    - いつも確かめる(前の紙は使えなくなる。刷り直しが要る)
    - 前の紙の履歴は消さずに「作り直し」の印。書き足したサイズ・LOTNO は引き継ぐ
    - 副番履歴は、このNoから外れたコイルを外し、加わったコイルを足す
    """
    result = OutputResult()
    if not board.editing:
        result.refuse = REFUSE_NOT_EDITING
        result.message = "修正中ではありません。"
        return result
    seq_no = board.edit_no
    old = find_output(conn, board.lot_no, seq_no)
    if old is None:
        result.refuse = REFUSE_NO_OUTPUT
        result.message = f"No{seq_no}の出力はもうありません。修正をやめてください。"
        return result
    if not board.is_complete:
        result.refuse = REFUSE_INCOMPLETE
        result.message = "積み上げが埋まっていません。"
        return result
    keys = board.stacked_keys()
    if len(keys) > config.SHEET_ROWS:
        result.refuse = REFUSE_TOO_MANY
        result.message = (f"1枚に載るのは{config.SHEET_ROWS}行までです"
                          f"（いま{len(keys)}行）。")
        return result
    others = output_keys(conn, board.lot_no, exclude_no=seq_no)
    clash = [k for k in keys if k in others]
    if clash:
        result.refuse = REFUSE_USED_ELSEWHERE
        result.message = ("ほかのNoで使っているコイルは積めません: " + "、".join(clash))
        return result

    if not confirm:
        removed = [k for k in old.keys if k not in keys]
        added = [k for k in keys if k not in old.keys]
        detail = [f"前: {' '.join(old.keys)}", f"後: {' '.join(keys)}"]
        if removed:
            detail.append("外すコイル: " + "、".join(removed))
        if added:
            detail.append("加えるコイル: " + "、".join(added))
        result.confirms = [Confirm(
            CONFIRM_EDIT,
            f"No{seq_no}を積み直した並びで出し直します。前の No{seq_no} の紙は使えなくなります"
            "(刷り直してください)。よろしいですか？", detail)]
        result.message = "確認してください。"
        return result

    with db.transaction(conn, name="出力済みの修正"):
        # 前の紙の履歴は**消さずに**「作り直し」の印を付ける
        if old.history_id:
            slip_history.mark_replaced(conn, old.history_id)
        history_id = slip_history.record_output(
            conn, lot_no=board.lot_no, seq_no=seq_no, keys=keys, weights=board.weights,
            lot=lot, orders=orders, hand_size=old.hand_size, hand_lotno=old.hand_lotno)
        # 行は作り直す(`管理番号` が変わる)。前の紙面を開いたままでも、その書き足しが
        # 新しい明細へ紛れ込まない(`save_hand_edits` が断る)
        db.write(conn, "DELETE FROM 明細出力 WHERE ロット番号 = ? AND No = ?",
                 (board.lot_no, seq_no), name="出力済みの修正")
        db.write(
            conn,
            "INSERT INTO 明細出力 (ロット番号, No, 副番の並び, 重量, 出力日時, 履歴ID,"
            " 手入力サイズ, 手入力ロット番号) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (board.lot_no, seq_no, ",".join(keys),
             ",".join(str(w) for w in board.weights), db.now_db_string(), history_id,
             old.hand_size, old.hand_lotno),
            name="出力済みの修正")
        # 副番履歴: このNoから外れたコイルを外し(ほかのNoに無ければ)、加わったコイルを足す
        for key in old.keys:
            if key not in keys and key not in others:
                history_repo.unregister_fuban(conn, history_repo.fuban_key(board.lot_no, key))
        for key in keys:
            full = history_repo.fuban_key(board.lot_no, key)
            if not history_repo.is_duplicate(conn, full):
                history_repo.register_fuban(conn, full)
        history_repo.set_printed(conn, False)

    log.info("output_edit: %s-No%s を積み直しました (%s行)", board.lot_no, seq_no, len(keys))
    result.ok = True
    result.seq_no = seq_no
    result.message = f"No{seq_no}を修正しました。刷り直してください(印刷)。"
    return result


# ==================================================================
# 全クリア (VER 0.13.23・現場の求め)
# ==================================================================
def cancel_lot_history(conn: sqlite3.Connection, lot_no: str) -> int:
    """そのロットの出力済みの紙を、明細の履歴で「取り消し」にする。戻り値は枚数。

    全クリアは出力済み・副番履歴・連番・途中経過を消して初めからやり直す。紙の履歴まで
    消すと、取り消す前に刷って現場へ出た紙が追えなくなるので、**消さずに印を付ける**。
    """
    ids = [o.history_id for o in list_outputs(conn, lot_no) if o.history_id]
    return slip_history.mark_cancelled(conn, ids) if ids else 0


def reset_stack(board: Board) -> None:
    """出力のあと積み上げだけ空にする (VBA `cmdKettei_Click` の末尾)。

    **出したコイルは使用済みのまま残す。** VBA は
    `ClearGimmickC` → `BuildGimmickC` しか呼ばず、ギミックBの
    `IsUsed` はそのまま ── 「Cスロットのみ。Bは出力後もそのまま残す」
    とコメントしている。**同じコイルを次の梱包へ二重に積ませない**
    ための仕掛けなので、ここを落とすと現物と紙が食い違う。

    スロットから出すだけだと `used` はスロット由来なので消えてしまう。
    復元と同じ「使用済みだがスロットには無い」側へ移す。
    """
    carried = [k for k in board.slots if k]
    board.clear_slots()
    for key in carried:
        if key not in board.restored_used:
            board.restored_used.append(key)


# ==================================================================
# 出力の一覧と読み出し
# ==================================================================
@dataclass
class Output:
    """出力済みの1枚。"""

    lot_no: str
    seq_no: int
    keys: list[str]
    weights: list[int]
    printed_at: str = ""
    # 行の番号(`管理番号`)。**書き足しを返すときの目印。**
    # 出力済み修正で出し直すと行が作り直されて番号が変わるので、古い紙面
    # からの書き足しが新しい明細へ紛れ込まない
    row_id: int = 0
    # 印刷の前に紙面で書き足したもの。空なら何も印字しない
    hand_size: str = ""
    hand_lotno: str = ""
    # 明細の履歴(`slip_history`)の送信ID。紙面を開いたことを履歴に映すため
    history_id: str = ""

    @property
    def sheet_name(self) -> str:
        """VBA のシート名と同じ形。`N3250Q0-No1`。"""
        return f"{self.lot_no}-No{self.seq_no}"

    def weight_of(self, key: str) -> int:
        """行の重量 (VBA `WriteData` の `weights(jouPart)`)。"""
        try:
            jou, _ = strand_service.parse_key(key)
        except ValueError:
            return 0
        return self.weights[jou - 1] if 1 <= jou <= len(self.weights) else 0

    def rows(self) -> list[tuple[str, str]]:
        """紙に出る行。`(副番, 重量の表示)`。

        VBA `WriteData` は不正なキー(`-` が無い/先頭が空)を飛ばしつつ
        **行だけは進める**。同じ見え方にするため、読めないキーは
        重量を空にして行を残す。
        """
        out: list[tuple[str, str]] = []
        for key in self.keys:
            try:
                strand_service.parse_key(key)
            except ValueError:
                log.debug("rows: 不正キー key=%s", key)
                out.append((key, ""))
                continue
            out.append((key, f"{lot_repo.fmt(self.weight_of(key), 1)}Kg"))
        return out


def _to_output(row: sqlite3.Row) -> Output:
    return Output(
        lot_no=row["ロット番号"],
        seq_no=int(row["No"]),
        keys=[k for k in str(row["副番の並び"] or "").split(",") if k],
        weights=[int(float(w)) for w in str(row["重量"] or "").split(",") if w.strip()],
        printed_at=str(row["出力日時"] or ""),
        row_id=int(row["管理番号"] or 0),
        hand_size=str(row["手入力サイズ"] or ""),
        hand_lotno=str(row["手入力ロット番号"] or ""),
        history_id=str(row["履歴ID"] or "") if "履歴ID" in row.keys() else "",
    )


# ==================================================================
# 印刷の前の書き足し
# ==================================================================
# 紙面の「寸法」「検番」の行へ、刷る前に人が書き足すもの。
#
# **VBA はシートを直してから刷れた。** 紙がExcelシートだったので、
# 気に入らなければセルへ打ち込めた。Web版は紙面がHTMLなので、同じ
# ことを紙面の上でできるようにする(流用元の `printing.editable`)。
HAND_FIELDS = {
    "size": ("手入力サイズ", "サイズ"),
    "lotno": ("手入力ロット番号", "LOTNO"),
}

# 1欄に入れられる文字数。**紙の上で入り切る長さ。**
# 欄の幅は B〜C列(約47mm)で、9pt の太字なら半角20字ほど。それを
# 超えると紙の上で切れる ── 切れて刷られるより、入れる前に断る
HAND_MAX_LEN = 20


class HandEditError(ValueError):
    """書き足しを受け取れない。`status` は画面へ返すHTTPの番号。"""

    def __init__(self, message: str, status: int = 400) -> None:
        super().__init__(message)
        self.status = status


def save_hand_edits(conn: sqlite3.Connection, lot_no: str,
                    edits: dict[str, object]) -> int:
    """紙面で書き足した内容を覚える。直した明細の枚数を返す。

    `edits` は紙面の欄の名前と中身。名前は `<管理番号>.<欄>`
    (例 `12.size` / `12.lotno`)。紙面は開くたびに全欄を送ってくる
    ので、**1枚ぶんは丸ごと入れ替える**。

    **1枚でも書けなければ、1枚も書かない。** まとめて開いた紙面で
    一部だけ残ると、刷った紙と記録が食い違う。

    **作り直された明細には書かない。** 出力済み修正で出し直すと行が
    作り直されて `管理番号` が変わる。古い紙面を開いたままだと、
    前の明細への書き足しが新しい明細へ紛れ込むので、断る。
    """
    if not isinstance(edits, dict):
        raise HandEditError("書き足した内容が読めませんでした。")

    per_row: dict[int, dict[str, str]] = {}
    for name, raw in edits.items():
        row_text, _, field_name = str(name).partition(".")
        if not row_text.isdigit() or field_name not in HAND_FIELDS:
            raise HandEditError(f"知らない欄です: {name}")
        value = db.sanitize_for_db(str(raw if raw is not None else ""))
        label = HAND_FIELDS[field_name][1]
        if len(value) > HAND_MAX_LEN:
            raise HandEditError(
                f"{label}は{HAND_MAX_LEN}文字までです(いま{len(value)}文字)。"
                "長いと紙の上で切れて刷られます。")
        per_row.setdefault(int(row_text), {})[field_name] = value

    with db.transaction(conn, name="紙面の書き足し"):
        for row_id, fields in per_row.items():
            sets = ", ".join(f'"{HAND_FIELDS[f][0]}" = ?' for f in fields)
            count = db.write(
                conn,
                f"UPDATE 明細出力 SET {sets} WHERE 管理番号 = ? AND ロット番号 = ?",
                (*fields.values(), row_id, lot_no), name="紙面の書き足し")
            if count != 1:
                # 例外で抜ければ、まとまりごと戻る(ほかの面も書かない)
                raise HandEditError(
                    "この明細は作り直されたか、もうありません。"
                    "紙面を開き直してください。", status=409)
            # 履歴にも映す(**紙に出る中身**なので、あとから追えるように)
            row = conn.execute(
                "SELECT 履歴ID, 手入力サイズ, 手入力ロット番号 FROM 明細出力"
                " WHERE 管理番号 = ?", (row_id,)).fetchone()
            if row is not None and row["履歴ID"]:
                slip_history.update_hand(conn, row["履歴ID"], row["手入力サイズ"],
                                         row["手入力ロット番号"])
    log.info("save_hand_edits: %s %s枚", lot_no, len(per_row))
    return len(per_row)


def list_outputs(conn: sqlite3.Connection, lot_no: str) -> list[Output]:
    """そのロットの出力一覧 (VBA `RefreshSheetList`)。No順。"""
    rows = db.fetch_all_strict(conn,
                               "SELECT * FROM 明細出力 WHERE ロット番号 = ? ORDER BY No",
                               (lot_no,), caller_name="meisai_service.list_outputs")
    return [_to_output(r) for r in rows]


def find_output(conn: sqlite3.Connection, lot_no: str,
                seq_no: int) -> Optional[Output]:
    """1枚を読み出す。帳票を組み立てるときに使う。"""
    row = db.fetch_one(conn,
                       "SELECT * FROM 明細出力 WHERE ロット番号 = ? AND No = ?",
                       (lot_no, seq_no), caller_name="meisai_service.find_output")
    return _to_output(row) if row else None


# ==================================================================
# 開いたままの紙面がまだ使えるか (統合 1.2.49・現場の指摘)
# ==================================================================
def slip_identity(output: Output) -> str:
    """紙面を作ったときの身元。`<管理番号>:<履歴ID>`。行が無い(試験の作り物)なら空。

    出力済み修正で出し直すと行が作り直されて `管理番号` も `履歴ID` も変わる。全クリア・
    ロットの切り替えでは行が消える。どちらも、開いたままの紙面の身元が見つからなくなる。
    """
    if not output.row_id:
        return ""
    return f"{output.row_id}:{output.history_id}"


def check_slips(conn: sqlite3.Connection, identities) -> dict:
    """開いたままの紙面(`slip_identity` の並び)が、まだそのまま刷ってよいか。

    以前は、No1 の紙面のタブを残したまま「出力済み修正」で積み直す・全クリア・ロットの
    切り替えをしたあと、古いタブで刷ると、もう使わない積み順・重量の紙が出た(紙面のタブは
    右上の文字しか聞き直さなかった)。**手元の明細出力を引くだけ**(共有は見ない)。

    戻り値 `{"valid": bool, "title": 大きく出す一言, "detail": どの紙か}`。
    """
    from .slip_history import STATE_CANCELLED, STATE_REPLACED

    gone: list[tuple[str, str]] = []
    for item in identities or []:
        row_text, _, history_id = str(item).strip().partition(":")
        if not row_text.isdigit():
            continue
        row = conn.execute("SELECT 履歴ID FROM 明細出力 WHERE 管理番号 = ?",
                           (int(row_text),)).fetchone()
        if row is not None and (not history_id or str(row[0] or "") == history_id):
            continue
        name, state = "", ""
        if history_id:
            found = conn.execute("SELECT ロット番号, No, 状態 FROM 明細履歴 WHERE 送信ID = ?",
                                 (history_id,)).fetchone()
            if found is not None:
                name, state = f"{found[0]}-No{found[1]}", str(found[2] or "")
        gone.append((name, state))
    if not gone:
        return {"valid": True, "title": "", "detail": ""}
    names = "・".join(n for n, _ in gone if n) or "この紙"
    states = {st for _, st in gone}
    if states == {STATE_REPLACED}:
        title = "この紙は作り直されました。開き直してください"
        detail = (f"{names} は「出力済み修正」で積み直されました。この画面は前の中身なので"
                  "刷れません。出力済みの一覧から選んで「印刷」で開き直してください。")
    elif states == {STATE_CANCELLED}:
        title = "この紙は取り消されました"
        detail = f"{names} は「全クリア」で取り消されました。この紙は刷りません。"
    else:
        title = "この紙は作り直されました(または取り消されました)。開き直してください"
        detail = (f"{names} は作り直し・取り消し・ロットの切り替えで、もうありません。"
                  "この画面は前の中身なので刷れません。")
    log.info("check_slips: 古い紙面を止めました (%s)", names)
    return {"valid": False, "title": title, "detail": detail}


def delete_lot_outputs(conn: sqlite3.Connection, lot_no: str) -> int:
    """そのロットの出力を全部消す (VBA `DeleteLotSheets`)。

    VBA はロットを切り替えるときに前ロットのシートを消していた。
    戻り値は消した枚数。
    """
    count = db.write(conn, "DELETE FROM 明細出力 WHERE ロット番号 = ?",
                     (lot_no,), name="前ロットの出力の削除")
    log.info("delete_lot_outputs: %s の出力を%s件消しました", lot_no, count)
    return count
