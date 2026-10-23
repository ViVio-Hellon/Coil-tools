# -*- coding: utf-8 -*-
"""梱包資材マスタ（Access ``梱包資材マスタ.accdb`` / テーブル ``資材重量``）。

VBA の該当処理
    GetRecordsArr / GetFieldsArr        -> load()
    GetFieldIndices                     -> MaterialTable.field_index
    GetUnitMassAndCoefficient           -> MaterialTable.unit_of()

参照される資材名は VBA にハードコードされていたものをそのまま定数化している。
マスタに存在しない資材名を引いた場合、VBA は ``Array("", "")`` を返し
``Val("") * Val("")`` = 0 として **静かに 0kg** で計算を続ける。
移植でも結果（0）は同じにするが、**どの資材が見つからなかったかを記録**して
画面とログに残す（VBA では気付けなかった点の改善）。
"""

from __future__ import annotations

import csv
import hashlib
import json
import logging
import os
import sqlite3
import threading
import time
from typing import Dict, List, Optional, Tuple

from . import csv_text
from .access_bridge import AccessBridge, AccessError

log = logging.getLogger(__name__)

#: VBA が参照する資材名（ハードコードされていたキー）
MAT_PIN = "テスラピン"
MAT_TIP_1000 = "チップボール1000ф"
MAT_TIP_950 = "チップボール950ф"
MAT_TIP_820 = "チップボール820ф"
MAT_SUT = "ストレッチフィルム"
MAT_ESA = "エサフォーム"
MAT_POR = "ポリシート"
MAT_PET = "PETバンド"
MAT_PAR = "樹脂パレット"
MAT_DRY = "EX-DRY"
MAT_HB = "ハードボード"
MAT_HB_OCT = "八角ハードボード"
MAT_HB590 = "HB590*2000"
MAT_HB530 = "HB530*2000"
MAT_HB580 = "HB580*2000"   # 羅列計算で取得のみ・未使用（VBA と同じく保持）
MAT_HB500 = "HB500*2000"   # 同上

#: チップボール選択 -> マスタ資材名
TIP_MATERIAL = {
    "TIP1000": MAT_TIP_1000,
    "TIP950": MAT_TIP_950,
    "TIP820": MAT_TIP_820,
}

#: VBA GetFieldIndices が探すフィールド名
FLD_KANRI = "管理番号"
FLD_NAME = "梱包資材名"
FLD_MASS = "単位質量"
FLD_COEF = "係数"

TABLE_NAME = "資材重量"
ACCDB_NAME = "梱包資材マスタ.accdb"


class MaterialTable:
    """``資材重量`` テーブル 1 枚分。"""

    def __init__(self, fields: List[str], records: List[List[Optional[str]]],
                 source: str = ""):
        self.fields = fields
        self.records = records
        self.source = source
        #: VBA GetFieldIndices 相当（見つからないフィールドは None）
        self.field_index: Dict[str, Optional[int]] = {
            name: (fields.index(name) if name in fields else None)
            for name in (FLD_KANRI, FLD_NAME, FLD_MASS, FLD_COEF)
        }
        self._by_name: Dict[str, Tuple[str, str]] = {}
        i_name = self.field_index[FLD_NAME]
        i_mass = self.field_index[FLD_MASS]
        i_coef = self.field_index[FLD_COEF]
        if i_name is not None:
            for row in records:
                if i_name >= len(row):
                    continue
                key = "" if row[i_name] is None else str(row[i_name])
                mass = row[i_mass] if (i_mass is not None and i_mass < len(row)) else ""
                coef = row[i_coef] if (i_coef is not None and i_coef < len(row)) else ""
                # VBA は最初に一致した行を返す（Exit Function）
                if key not in self._by_name:
                    self._by_name[key] = ("" if mass is None else str(mass),
                                          "" if coef is None else str(coef))

    # ------------------------------------------------------------
    def coefficients(self) -> List[Tuple[str, str]]:
        """``(資材名, 係数)`` の一覧（マスタの文字列のまま）。

        係数が何を表すかはコードから読み取れない（解析書 C-3）。
        全件 ``1`` なら実質未使用と分かるので、診断画面で見えるようにする。
        """
        return [(name, coef) for name, (_m, coef) in self._by_name.items()]

    def unit_of(self, material_name: str) -> Tuple[str, str]:
        """VBA ``GetUnitMassAndCoefficient``。

        見つからない場合は VBA と同じく ``("", "")`` を返す。
        （呼び出し側で ``val("") * val("")`` = 0 となる）
        """
        return self._by_name.get(material_name, ("", ""))

    def has(self, material_name: str) -> bool:
        return material_name in self._by_name

    def names(self) -> List[str]:
        return list(self._by_name.keys())

    def as_rows(self) -> List[Dict[str, str]]:
        """表記用フォームの ListView（梱包資材名 / 単位質量）相当。"""
        i_name = self.field_index[FLD_NAME]
        i_mass = self.field_index[FLD_MASS]
        out = []
        for row in self.records:
            name = row[i_name] if (i_name is not None and i_name < len(row)) else ""
            mass = row[i_mass] if (i_mass is not None and i_mass < len(row)) else ""
            out.append({"name": "" if name is None else str(name),
                        "mass": "" if mass is None else str(mass)})
        return out

    def is_empty(self) -> bool:
        """VBA の ``IsEmpty(TempHiki)`` 相当。"""
        return not self.records or not self.fields

    def values(self) -> Dict[str, List[str]]:
        """``{資材名: [単位質量, 係数]}``(計算に使う中身だけ。統合 1.2.49)。"""
        return {name: [mass, coef] for name, (mass, coef) in self._by_name.items()}

    def content_hash(self) -> str:
        """計算に使う中身(資材名・単位質量・係数)の指紋(統合 1.2.49・現場の指摘)。

        計算したときの指紋を控えておき、印刷のときにいまの指紋と比べる。違えば、資材重量が
        計算のあとに変わった(前の NW・GW のまま刷ろうとしている)と分かる。
        """
        text = json.dumps(sorted(self.values().items()), ensure_ascii=False)
        return hashlib.sha256(text.encode("utf-8")).hexdigest()[:16]


#: 取得元の呼び名(画面右上の札・記録と同じ)
SOURCE_LABELS = {"sqlite": "SQLite", "access": "Access(ACE)", "csv": "CSV"}


class MaterialRepository:
    """資材マスタの取得元を切り替えるリポジトリ。

    読む順(統合 1.2.31。現場の指摘で直した):

    1. SQLite(``db_path`` が設定されていて実在するとき)
    2. Access(Windows かつ accdb が設定されているとき。cscript + ADODB + ACE)
    3. **設定で指定した CSV**(``csv_path``。空なら読まない)
    4. 直前に読んだ内容(取得元へ届かないとき。``serving_stale = True``。2026.09 決定 D-8)
    5. どれも無ければ止める(``AccessError``。VBA と同じ「資材重量なし」)。**何が読めなかったか・
       どこを見ればよいか**を文にして返す

    **同梱の data/資材重量.csv は読まない。** 試験用のサンプル値で、業務値ではない。以前は 3 の位置で
    必ず読まれ、朝から共有フォルダへ届かない・作業の途中で瞬断した、だけでサンプルの単重に入れ替わり、
    間違った NW/GW のラベルが黙って刷れた(画面右上のバッジが変わるだけ)。
    どれを使ったかは ``last_source`` で判別できる。
    """

    def __init__(self, accdb_dir: str, csv_path: str = "",
                 prefer_access: bool = True, timeout_sec: int = 60,
                 refresh_sec: int = 300, db_path: str = ""):
        self.accdb_dir = accdb_dir or ""
        self.csv_path = csv_path or ""
        #: SQLite 版の梱包資材マスタ（設定されていれば最優先で読む）
        self.db_path = db_path or ""
        self.prefer_access = prefer_access
        self.bridge = AccessBridge(timeout_sec=timeout_sec)
        self.last_source = ""
        self.last_error = ""
        #: 取得元へ到達できず、直前に読んだ内容で動いているか
        self.serving_stale = False
        #: 鮮度チェックの間隔（秒）。0 で「毎回チェックする」
        self.refresh_sec = int(refresh_sec)
        self._cache: Optional[MaterialTable] = None
        #: キャッシュ時点の取得元の印(`_stamp`。**読む前**に取ったもの)
        self._cache_stamp: Optional[tuple] = None
        self._cache_at = 0.0
        self._cache_at_wall = 0.0
        self._lock = threading.Lock()

    # ------------------------------------------------------------
    @property
    def accdb_path(self) -> str:
        """VBA: ``PATH_AIM_参照 & "梱包資材マスタ.accdb"``"""
        return os.path.join(self.accdb_dir, ACCDB_NAME) if self.accdb_dir else ""

    # ------------------------------------------------------------
    @staticmethod
    def _stamp(path: str) -> Optional[tuple]:
        """取得元ファイルの印 (パス, 更新日時, サイズ, 更新日時ns, -wal の (サイズ, 更新日時ns))。

        取れなければ None。**読む前に取る**(統合 1.2.49)。読んだあとに取ると、読んでいる途中に
        書かれた変更の更新日時を「読んだ内容の印」として控えてしまい、その変更を読み直さなかった。
        SQLite は WAL のとき変更が ``-wal`` にだけ書かれ、本体の更新日時・大きさが変わらないので、
        ``-wal`` も印に入れる(無ければ None)。比べるのは ``!=``(どこか 1 つでも違えば読み直す)。
        """
        try:
            st = os.stat(path)
        except OSError:
            return None
        try:
            ws = os.stat(path + "-wal")
            wal = (ws.st_size, ws.st_mtime_ns)
        except OSError:
            wal = None
        return (os.path.abspath(path), st.st_mtime, st.st_size, st.st_mtime_ns, wal)

    def _source_path(self) -> str:
        """今まさに読むべき取得元のパス。

        優先順は SQLite -> Access -> CSV。
        SQLite は設定されていてファイルが実在するときだけ使う。
        """
        if self.db_path and os.path.exists(self.db_path):
            return self.db_path
        if self.prefer_access and AccessBridge.available() and self.accdb_path:
            return self.accdb_path
        return self.csv_path

    def _source_kind(self) -> str:
        """`_source_path` の種類("sqlite" / "access" / "csv" / "")。"""
        path = self._source_path()
        if not path:
            return ""
        if path == self.db_path:
            return "sqlite"
        if path == self.accdb_path:
            return "access"
        return "csv"

    def preferred_source(self) -> str:
        """設定でいちばん先に読むことになっている取得元("sqlite" / "access" / "csv" / "")。

        これ以外で計算したら(SQLite に届かず CSV に落ちた等)、計算の返事で知らせる(統合 1.2.49)。
        """
        if self.db_path:
            return "sqlite"
        if self.prefer_access and AccessBridge.available() and self.accdb_path:
            return "access"
        if self.csv_path:
            return "csv"
        return ""

    def is_stale(self, check_now: bool = False) -> bool:
        """マスタが更新された可能性があるか。

        業務中にマスタが差し替わっても古い単重で計算し続けないよう、
        **更新日時とサイズ** を見て判定する。
        ``refresh_sec`` の間は毎回 stat しない（共有フォルダーへの負荷を避ける）。

        ``check_now`` を立てると、その間隔を飛ばして**必ず確かめる**。
        「計算」を押した瞬間のように、**古い単重で計算しては困る**場面で使う。
        SQLite と CSV は **毎回読み直す**(統合 1.2.49)。更新日時と大きさだけでは、値を直しても
        大きさの変わらない SQLite・更新日時の刻みが粗い共有・WAL(``-wal`` にだけ書かれる)の
        変更を見逃した。資材重量は十数行の小さな表なので、読み直しても重くない。
        Access(cscript 経由で数秒かかる)は今までどおり印(stat 1 回)で比べる。
        """
        if self._cache is None:
            return True
        # 経過時間は **進むだけの時計** で測る（time.time() は時刻変更で飛ぶ）。
        if not check_now and self.refresh_sec > 0 \
                and (time.monotonic() - self._cache_at) < self.refresh_sec:
            return False
        # 届かずに直前の内容で動いているあいだは、確かめる時が来るたびに読み直しを試す
        # (届くようになったら「古い内容」を外す)
        if self.serving_stale:
            return True
        if check_now and self._source_kind() in ("sqlite", "csv"):
            return True
        return self._stamp(self._source_path()) != self._cache_stamp

    def load(self, use_cache: bool = True,
             check_now: bool = False) -> MaterialTable:
        """資材マスタを読む。失敗時は ``AccessError``。

        キャッシュは持つが、**取得元の更新日時・サイズが変わっていれば
        自動で読み直す**（ツールを起動したまま放置し、その間に
        マスタが更新された場合の取りこぼしを防ぐ）。

        ``check_now`` … 鮮度チェックの間隔を飛ばして必ず確かめる。
        計算のように、古い単重を使っては困る場面で立てる。
        """
        with self._lock:
            if use_cache and self._cache is not None \
                    and not self.is_stale(check_now=check_now):
                return self._cache
            before = self._cache
            table = self._load_locked()
            # 「計算」「印刷」のたびに読み直すので、記録に残すのは中身が変わったときだけ
            if before is not None and table is not before \
                    and before.values() != table.values():
                log.info("資材マスタの中身が変わっていたので読み直しました: %s",
                         self._source_path())
            return table

    def _load_locked(self) -> MaterialTable:
        errors = []

        # 1) SQLite（設定されていれば最優先）
        if self.db_path:
            if os.path.exists(self.db_path):
                try:
                    stamp = self._stamp(self.db_path)          # 読む前に取る(統合 1.2.49)
                    table = self._load_sqlite(self.db_path)
                    self.last_source = "sqlite"
                    self.last_error = ""
                    self._remember(table, stamp)
                    return table
                except sqlite3.Error as exc:
                    errors.append("SQLite: %s" % exc)
                    log.warning("SQLite を読めません、次の取得元を試します: %s", exc)
            else:
                errors.append("梱包資材マスタ(SQLite): %s が見つかりません" % self.db_path)

        # 2) Access
        if self.prefer_access and AccessBridge.available() and self.accdb_path:
            try:
                stamp = self._stamp(self.accdb_path)           # 読む前に取る(統合 1.2.49)
                fields, records = self.bridge.fetch(self.accdb_path, TABLE_NAME, "")
                table = MaterialTable(fields, records, source=self.accdb_path)
                self.last_source = "access"
                self.last_error = ""
                self._remember(table, stamp)
                return table
            except AccessError as exc:
                errors.append("Access: %s を読めません（%s）" % (self.accdb_path, exc))
                log.warning("Access を読めません、次の取得元を試します: %s", exc)

        # 3) 設定で指定した CSV(空なら読まない。同梱のサンプルは読まない)
        if self.csv_path:
            if os.path.exists(self.csv_path):
                try:
                    stamp = self._stamp(self.csv_path)         # 読む前に取る(統合 1.2.49)
                    table = self._load_csv(self.csv_path)
                except (OSError, UnicodeDecodeError, csv.Error) as exc:
                    errors.append("CSV: %s を読めません（%s）" % (self.csv_path, exc))
                else:
                    self.last_source = "csv"
                    self.last_error = "; ".join(errors)
                    self._remember(table, stamp)
                    return table
            else:
                errors.append("CSV: %s が見つかりません" % self.csv_path)
        self.last_error = "; ".join(errors)

        # 4) 取得元へ一時的に到達できないだけなら、直前のマスタで業務を続けさせる。
        # （ライン作業の途中でネットワークが瞬断しても止めないため。2026.09 決定 D-8）
        # 取得元の表示は変えず、"古い内容で動いている" ことだけを別に立てる。
        if self._cache is not None:
            log.warning("資材マスタを読み直せないため、直前の内容で継続します: %s",
                        self.last_error)
            self._cache_at = time.monotonic()      # 次に試すのは鮮度チェックの間隔のあと
            self.serving_stale = True
            return self._cache

        # 5) 止める。何が読めなかったか・どこを見ればよいかを言う
        self.last_source = ""
        raise AccessError(self.why_unreadable(errors))

    def why_unreadable(self, errors: List[str]) -> str:
        """読めなかったときの説明(画面にそのまま出す)。"""
        lines = ["資材マスタを読めないため、計算を止めました。"
                 "このまま計算すると、ラベルの NW・GW が正しくなりません。"]
        tried = [e for e in errors if e]
        if not (self.db_path or (self.prefer_access and self.accdb_path) or self.csv_path):
            lines.append("・資材マスタの参照先が設定されていません"
                         "(設定の「梱包資材マスタ(SQLite)」または「資材マスタの参照先フォルダー」)。")
        elif self.prefer_access and self.accdb_path and not AccessBridge.available() \
                and not self.db_path and not self.csv_path:
            lines.append("・Access の参照先 %s は、この PC では読めません"
                         "(Windows の cscript が使えません)。" % self.accdb_path)
        for item in tried:
            lines.append("・" + item)
        lines.append("確かめること: 共有フォルダに届くか(ネットワーク・サーバーが止まっていないか)、"
                     "設定の資材マスタの参照先が合っているか。")
        lines.append("届くようになったら、もう一度「計算」を押してください。")
        return "\n".join(lines)

    def _remember(self, table: MaterialTable, stamp: Optional[tuple]) -> None:
        """読んだ内容を控える。``stamp`` は**読む前**に取った取得元の印。"""
        self._cache = table
        self.serving_stale = False
        self._cache_stamp = stamp
        self._cache_at = time.monotonic()     # 経過判定用
        self._cache_at_wall = time.time()     # 表示用

    def describe(self) -> dict:
        """いま使っている資材重量の素性(計算・印刷の記録に残す。統合 1.2.31)。

        取得元・ファイル・**そのファイルの更新日時(いつ時点のマスタか)**・読んだ時刻・
        届かずに前の内容で動いているか。
        """
        stamp = self._cache_stamp
        fmt = "%Y-%m-%d %H:%M:%S"
        preferred = self.preferred_source()
        return {
            "source": self.last_source or "",
            "label": SOURCE_LABELS.get(self.last_source, "未読込"),
            "stale": bool(self.serving_stale),
            # 設定でいちばん先に読む取得元と、それ以外で読んだか(統合 1.2.49)
            "preferred": preferred,
            "preferredLabel": SOURCE_LABELS.get(preferred, ""),
            "fallback": bool(preferred and self.last_source and self.last_source != preferred),
            "path": stamp[0] if stamp else "",
            "asOf": time.strftime(fmt, time.localtime(stamp[1])) if stamp else "",
            "readAt": (time.strftime(fmt, time.localtime(self._cache_at_wall))
                       if self._cache_at_wall else ""),
        }

    def cache_info(self) -> dict:
        """/diag 画面と運用確認用。"""
        return {
            "cached": self._cache is not None,
            "source": self.last_source,
            "sourcePath": self._source_path(),
            "loadedAt": (time.strftime("%Y-%m-%d %H:%M:%S",
                                       time.localtime(self._cache_at_wall))
                         if self._cache_at_wall else ""),
            "ageSec": (round(time.monotonic() - self._cache_at, 1)
                       if self._cache_at else None),
            "refreshSec": self.refresh_sec,
            "lastError": self.last_error,
            "servingStale": self.serving_stale,
        }

    # ------------------------------------------------------------
    @staticmethod
    def _load_sqlite(path: str) -> MaterialTable:
        """SQLite 版の梱包資材マスタから ``資材重量`` を読む。

        **読取専用で開く**（``mode=ro``）。マスタは共有されている想定で、
        こちらから書いたりジャーナルを作ったりしないため。
        """
        uri = "file:%s?mode=ro" % path.replace("?", "%3f").replace("#", "%23")
        conn = sqlite3.connect(uri, uri=True, timeout=10)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.execute('SELECT * FROM "%s"' % TABLE_NAME)
            fields = [d[0] for d in cur.description]
            records = [[None if v is None else str(v) for v in tuple(r)]
                       for r in cur.fetchall()]
        finally:
            conn.close()
        # 「計算」「印刷」のたびに読むので、ふだんの読みは debug に(統合 1.2.49)
        log.debug("SQLite から資材マスタを読みました: %s (%d 件)", path, len(records))
        return MaterialTable(fields, records, source=path)

    @staticmethod
    def _load_csv(path: str) -> MaterialTable:
        # UTF-8 でも Shift_JIS（Excel で保存した CSV）でも読む（統合 1.2.8）
        rows = csv_text.read_rows(path)
        if not rows:
            return MaterialTable([], [], source=path)
        return MaterialTable([c.strip() for c in rows[0]], rows[1:], source=path)

    def reconfigure(self, accdb_dir: str, csv_path: str, prefer_access: bool,
                    timeout_sec: int, refresh_sec: int,
                    db_path: str = "") -> bool:
        """設定変更を再起動せずに反映する。

        参照先が変わった場合はキャッシュを捨てて次回に読み直す。
        戻り値は「参照先が実際に変わったか」。
        """
        with self._lock:
            changed = (
                (accdb_dir or "") != self.accdb_dir
                or (csv_path or "") != self.csv_path
                or (db_path or "") != self.db_path
                or bool(prefer_access) != self.prefer_access
            )
            self.accdb_dir = accdb_dir or ""
            self.csv_path = csv_path or ""
            self.db_path = db_path or ""
            self.prefer_access = bool(prefer_access)
            self.bridge.timeout_sec = int(timeout_sec)
            self.refresh_sec = int(refresh_sec)
            if changed:
                self._cache = None
                self._cache_stamp = None
                self._cache_at = 0.0
                self._cache_at_wall = 0.0
                self.serving_stale = False
                log.info("資材マスタの参照先を変更しました: "
                         "sqlite=%r accdb=%r csv=%r access=%s",
                         self.db_path, self.accdb_dir, self.csv_path,
                         self.prefer_access)
        return changed

    def clear_cache(self) -> None:
        with self._lock:
            self._cache = None
            self._cache_stamp = None
            self._cache_at = 0.0
            self._cache_at_wall = 0.0
            self.serving_stale = False
