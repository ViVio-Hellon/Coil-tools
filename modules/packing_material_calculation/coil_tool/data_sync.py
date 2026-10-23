"""取り込み ── 共有フォルダのマスタを手元へ写す

VBA は計算のたびに共有フォルダの .accdb を開いて読んでいた
(`GetRecordsArr(filePath, tableName, SQLfilter)`)。1回の計算で
同じファイルを何度も開くので、共有が重い日は目に見えて遅かった。

ここでは**一度写して、手元の SQLite を引く**。

    共有フォルダ (読むだけ)  ──取り込み──▶  手元の作業用DB  ──▶ 画面
       source_db(mode=ro)                     db.py

【取り込み中も画面を止めない】
1テーブルぶんを「DELETE してから全行 INSERT」する。既定の
rollback journal だとその間ずっと読み手が止まるので、WAL が使えるなら
使う(`sqlite_toolkit.enable_wal`)。共有フォルダでは WAL を張れない
ことがあるが、**手元の DB はローカル**なので効く。

【取り込み元が変わったかは「目印」で見る(統合 1.2.49・現場の指摘)】
取り込み履歴に、取り込み元の目印(本体と `-wal`・`-journal` の大きさ・更新時刻)を
**読む前に**取って残す。前と**一致するか**だけを見る(新しいか古いかは見ない)。
以前は「更新日時が取り込み日時より新しいか」で見ていて、次を取りこぼした:

    WAL で `-wal` にだけ書かれた / 更新日時を保ったまま差し替えた(エクスプローラのコピー)/
    古いバックアップに戻した / 写したあと・記録する前に共有が変わった / サーバと PC の時計のずれ

【動かしている間も確かめる(統合 1.2.49・現場の指摘)】
起動時の自動取り込み(設定で切れる)のほか、画面を開いたとき(上の帯)と「計算Start」のときに
目印を軽く確かめる(ファイルの大きさ・更新時刻を見るだけ。中身は読まない)。
- 梱包資材マスタ(包装仕様・パレット・リプラサイズ・班員名簿)・コイル実績 … 小さいので、
  変わっていればその場で取り込み直す
- 仕掛台帳 … 大きいので、帯に「共有の仕掛台帳が更新されています [いま取り込む]」を出す
以前は起動時に1回だけで、マスタ管理の画面(共有を直に読む)は新しいのに、計算は古い写しのまま、が起きた。

【任意の表「パレット種類_外径条件」(統合 1.2.50・現場の指示)】
梱包資材マスタの1表として、上の目印・確かめ・取り込み直し・世代にそのまま乗せる。ただし
**共有にまだ無いことがある**ので、無くても取り込みの誤りにせず(結果に「まだ作っていません」)、
帯の「古い」の対象にもしない。手元の表は共有に表があれば作り、無ければ消す(`_import_optional`)。
手元に表があるか無いかで、1C1103 のパレット種類をマスタで決めるかプログラムで決めるかが変わる
(`outer_condition.py`)。
"""
from __future__ import annotations

import hashlib
import itertools
import os
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Optional

from common import source_copy

from . import coil_results, config, db, import_specs, outer_condition, source_db
from .logging_utils import get_logger

log = get_logger("data_sync")


@dataclass
class TableResult:
    """1テーブルの取り込み結果。"""
    table: str
    ok: bool
    rows: int = 0
    source: str = ""
    error: str = ""
    # 取り込み元に無かった列(既定値で埋めたもの)。画面に出して気づけるように
    filled: list[str] = field(default_factory=list)
    # 鍵が重なっていて飛ばした行の鍵。**黙って捨てない** ──
    # 元のマスタに同じ番号の行が2つある、ということなので画面に出す
    duplicates: list[str] = field(default_factory=list)
    # 知らせ(失敗ではない)。仕掛台帳の第2パスから足した件数など(統合 1.2.7)
    note: str = ""


@dataclass
class SyncResult:
    """取り込み全体の結果。"""
    tables: list[TableResult] = field(default_factory=list)
    started_at: str = ""
    finished_at: str = ""

    @property
    def ok(self) -> bool:
        return bool(self.tables) and all(t.ok for t in self.tables)

    @property
    def total_rows(self) -> int:
        return sum(t.rows for t in self.tables)

    @property
    def failures(self) -> list[TableResult]:
        return [t for t in self.tables if not t.ok]


# ==================================================================
# 取り込み元の目印(統合 1.2.49・現場の指摘)
# ==================================================================
#: 目印に入れるファイル(本体・`-wal`・`-journal`)。`common/source_copy` が写すものと同じ。
#: WAL のファイルは書いても本体が変わらず `-wal` だけが変わるので、付属も見る
STAMP_PARTS = ("", "-wal", "-journal")


def source_stamp(path: Optional[Path]) -> str:
    """取り込み元の目印(本体と `-wal`・`-journal` の大きさ・更新時刻 ns)。見られなければ空文字。

    **ファイルの大きさと更新時刻を見るだけ**(中身は読まない)。共有が遠い日でも軽い。
    比べるときは**一致するか**だけを見る(`_same_source`)。新しいか古いかで見ると、
    更新日時を保ったままの差し替え・古いバックアップへの戻し・時計のずれを取りこぼす。
    """
    if path is None:
        return ""
    parts: list[str] = []
    for extra in STAMP_PARTS:
        try:
            st = os.stat(str(path) + extra)
        except FileNotFoundError:
            if not extra:
                return ""
            parts.append("-")
            continue
        except OSError:
            return ""
        parts.append(f"{st.st_size}:{st.st_mtime_ns}")
    return "|".join(parts)


def _same_source(record: Optional[dict], path: Optional[Path]) -> bool:
    """前に取り込んだときと同じ中身か。**場所が同じで、目印が一致する**ときだけ真。

    目印の無い記録(前の版で取り込んだもの)は「同じ」と言わない ── 黙って新しい扱いにしない。
    """
    if not record or path is None or record.get("元ファイル") != str(path):
        return False
    recorded = str(record.get("元の目印") or "")
    return bool(recorded) and recorded == source_stamp(path)


# ==================================================================
# 取り込み中(統合 1.2.49・現場の指摘)
# ==================================================================
# 取り込みは1表ずつ入れ替えるので、**途中は表の間で新旧が混ざる**(包装仕様は新しく、
# パレットはまだ前、など)。その間は「計算Start」を断り、帯に「取り込み中」を出す
_busy_lock = threading.Lock()
_busy_count = 0
# 動かしている間の取り込み直し(帯・計算Start)を2つ同時に走らせない
_refresh_lock = threading.Lock()
#: 取り込み直しの終わりを待つ上限(秒)。過ぎたら「取り込み中」として扱う
REFRESH_WAIT_SEC = 60


@contextmanager
def _importing():
    """この間は「取り込み中」。入れ子にしてよい。"""
    global _busy_count
    with _busy_lock:
        _busy_count += 1
    try:
        yield
    finally:
        with _busy_lock:
            _busy_count -= 1


def is_importing() -> bool:
    """いま取り込んでいるか(起動時の自動取り込み・いま取り込む・取り込み直し)。"""
    with _busy_lock:
        if _busy_count > 0:
            return True
    try:
        from . import jobs                     # 循環を避ける
        job = jobs.get_registry().running()
    except Exception:                          # noqa: BLE001 - 見られなければ取り込み中ではない扱い
        return False
    return job is not None and job.kind == "import"


# ==================================================================
# 取り込みの世代(統合 1.2.49・現場の指摘)
# ==================================================================
# 取り込むたびに進む数。同じ秒に2回取り込んでも、世代は変わる
_changes = itertools.count(1)
_last_change = 0


def _bump() -> None:
    global _last_change
    _last_change = next(_changes)


def generation(conn: Any) -> str:
    """いまの手元の写しの「世代」。取り込み・マスタ管理の書き込みのたびに変わる。

    計算した時点の世代を作業状態に持たせ(`CalcState.世代`)、チェックリストへ積むときに
    いまの世代と比べる。違えば、計算のあとにマスタ・台帳が変わっている。
    """
    try:
        rows = conn.execute(
            "SELECT テーブル名, 取り込み日時, 元ファイル, 件数, 元の目印 FROM 取り込み履歴"
            " ORDER BY テーブル名").fetchall()
    except Exception:                          # noqa: BLE001 - 読めなければ数だけで
        rows = []
    text = repr((_last_change, [tuple(r) for r in rows]))
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:16]


def forget_copies(paths: Iterable[Optional[Path]]) -> None:
    """覚えている写しを捨てる(次に読むときは必ず写し直す。統合 1.2.49)。

    「いま取り込む」・マスタ管理の取り込み直しのときに呼ぶ。目印が同じでも中身だけ
    変わっていることがある(更新時刻の刻みが粗い共有など)ので、押したときは必ず読み直す。
    """
    for path in paths:
        if path is not None:
            source_copy.forget(Path(path))


def source_paths() -> list[Path]:
    """いまの設定で読む取り込み元ぜんぶ(梱包資材マスタ・仕掛台帳の1つ目と2つ目・コイル実績)。"""
    out: list[Path] = []
    master = find_master_db()
    if master is not None:
        out.append(master)
    for table in import_specs.LOT_IMPORT_SPECS:
        for path in (find_lot_db(table), find_second_lot(table)):
            if path is not None and path not in out:
                out.append(path)
    if coil_results.wanted():
        out += [p for p in coil_results.sources() if p not in out]
    return out


# ==================================================================
# 取り込み元を探す
# ==================================================================
def find_master_db() -> Optional[Path]:
    """梱包資材マスタのファイルを探す。

    名前で見つからなければフォルダの中の sqlite3 を拾う。ただし
    **仕掛台帳(`SIKA` で始まる)は除く** ── 同じフォルダに置かれて
    いることがあり、名前の緩い一致で拾うと取り違える。
    """
    directory = config.master_db_dir()
    found = source_db.find(directory, config.MATERIAL_DB_NAME)
    if found is not None:
        return found
    try:
        candidates = source_db.list_source_files(directory)
    except OSError as exc:
        log.warning("マスタのフォルダを見られませんでした: %s (%s)", directory, exc)
        return None
    for path in candidates:
        if path.stem.upper().startswith("SIKA"):
            continue
        return path
    return None


def find_lot_db(table: str) -> Optional[Path]:
    """仕掛台帳の1ファイルを探す。1つ目 → 2つ目の順(VBA の2系統)。

    1つ目は決まった名前(`LOT_DB_FILES`)、2つ目は**第2パスのファイル名**で探す
    (統合 1.2.6。python-web-tools と同じ決まり)。
    """
    name = config.LOT_DB_FILES.get(table)
    if not name:
        return None
    found = source_db.find(config.lot_db_dir(), name)
    if found is not None:
        return found
    return find_second_lot(table)


def find_second_lot(table: str) -> Optional[Path]:
    """仕掛台帳の**2つ目**のその表のファイル。2つ目を設定していなければ None。

    第2パスのフォルダでは**第2パスのファイル名だけ**を探す。名前を変えていれば、2つ目に
    SIKAHIKI.sqlite3 があっても読まない(取り違えを防ぐ)。
    """
    dirs = config.lot_db_dirs()
    if len(dirs) < 2:
        return None
    name = config.lot_db2_files().get(table)
    return source_db.find_named(dirs[1], name) if name else None


def second_missing_columns(path: Path, table: str) -> list[str]:
    """第2パスのファイルに**足りない列**。現場の約束「名前は変えても、中身の要る列は
    必ずそろえる」を設定画面で確かめるため。読めなければ空。"""
    try:
        available = set(source_db.columns(Path(path), import_specs.LOT_SOURCE_TABLE))
    except (source_db.SourceError, OSError):
        return []
    if not available:
        return []
    return import_specs.missing_columns(table, available)


#: 2つ目の取り込みを覚える名前(取り込み履歴。表の記録とは別に持つ)
SECOND_NOTE = "{table}(2つ目)"
#: 第2パスから足すときの「同じもの」の決め方。この鍵が第1パスに無いものだけ足す
LOT_MERGE_KEYS = {"仕掛引当": "ロット番号", "仕掛受注": "受注番号"}


def import_lot_table(conn: Any, table: str, path: Path) -> TableResult:
    """仕掛台帳の1表を写し、2つ目にだけあるロット・受注を足す(統合 1.2.7)。

    1つ目にファイルはあるが、目当てのロット(受注)が入っていないとき、2つ目にあれば
    それを使う(python-web-tools と同じ)。両方にあれば第1パスを正とする。
    """
    result = _import_table(conn, dest_table=table, source_path=path,
                           source_table=import_specs.LOT_SOURCE_TABLE)
    second = find_second_lot(table)
    if result.ok and second is not None and second != path:
        _merge_second(conn, table, second, result)
    return result


def _merge_second(conn: Any, table: str, path: Path, result: TableResult) -> int:
    """第2パスのファイルから、**第1パスに無いもの**だけを手元に足す。足した行数。"""
    stamp = source_stamp(path)                  # 読む前に取る(統合 1.2.49)
    missing = second_missing_columns(path, table)
    if missing:
        result.note = (f"第2パスのファイル {path.name} に列がありません: {', '.join(missing)}"
                       "(第2パスからは足していません)")
        # 見たことは残す。残さないと、変わっていないのに毎回「更新されています」と出る。
        # ファイルが直れば目印が変わるので、そのとき読み直す(統合 1.2.49)
        _record_import(conn, SECOND_NOTE.format(table=table), path, 0, stamp)
        return 0
    try:
        rows = source_db.read_table(path, import_specs.LOT_SOURCE_TABLE)
    except source_db.SourceError as exc:
        result.note = f"第2パスのファイルを読めませんでした({exc})"
        return 0
    key = LOT_MERGE_KEYS[table]
    specs = import_specs.LOT_IMPORT_SPECS[table]
    columns = [dest for dest, _, _ in specs]
    quoted_table = source_db.quote_identifier(table)
    have = {str(r[0]) for r in conn.execute(
        f"SELECT DISTINCT {source_db.quote_identifier(key)} FROM {quoted_table}")}
    adding = []
    keys: set[str] = set()
    for row in rows:
        converted = import_specs.convert_row(table, row)
        if converted is None:
            continue
        value = converted.get(key)
        if value is None or str(value).strip() == "" or str(value) in have:
            continue
        adding.append([converted[c] for c in columns])
        keys.add(str(value))
    placeholders = ", ".join("?" for _ in columns)
    quoted = ", ".join(source_db.quote_identifier(c) for c in columns)
    try:
        with conn:
            conn.executemany(f"INSERT OR IGNORE INTO {quoted_table} ({quoted})"
                             f" VALUES ({placeholders})", adding)
    except Exception as exc:  # noqa: BLE001 - 1つ目の分は入っている
        result.note = f"第2パスから足せませんでした({exc})"
        log.warning("第2パスから足せませんでした %s: %s", table, exc)
        return 0
    _record_import(conn, SECOND_NOTE.format(table=table), path, len(adding), stamp)
    if adding:
        result.rows += len(adding)
        what = "受注" if key == "受注番号" else "ロット"
        result.note = (f"第1パスに無い{what} {len(keys):,}件({len(adding):,}行)を"
                       f"第2パス({path.name})から足しました")
    log.info("取り込み %s: 2つ目(%s)から %d 行足しました", table, path, len(adding))
    return len(adding)


# ==================================================================
# 1テーブルを写す
# ==================================================================
def _import_table(
    conn: Any, *, dest_table: str, source_path: Path, source_table: str,
) -> TableResult:
    """取り込み元の1テーブルを、手元の1テーブルへ写す。

    **全入れ替え**(DELETE → INSERT)。差分更新にしないのは、上流が
    行を消したときに手元へ残り続けるのを避けるため。
    """
    result = TableResult(table=dest_table, ok=False, source=str(source_path))
    # 目印は**読む前に**取る(統合 1.2.49)。読んだあとに取ると、写したあと・記録する前に
    # 共有が変わったとき、新しい目印に古い中身を結びつけてしまい、次の取り込みで飛ばされる
    stamp = source_stamp(source_path)

    try:
        available = set(source_db.columns(source_path, source_table))
    except source_db.SourceError as exc:
        result.error = f"取り込み元を読めません: {exc}"
        return result

    if not available:
        result.error = f"取り込み元にテーブル「{source_table}」がありません"
        return result

    missing = import_specs.missing_columns(dest_table, available)
    if missing:
        result.error = "取り込み元に列がありません: " + ", ".join(missing)
        return result

    specs = (import_specs.master_specs(dest_table)
             or import_specs.LOT_IMPORT_SPECS[dest_table])
    result.filled = [src for _, src, _ in specs if src not in available]

    try:
        rows = source_db.read_table(source_path, source_table)
    except source_db.SourceError as exc:
        result.error = f"取り込み元を読めません: {exc}"
        return result

    converted = [r for r in (import_specs.convert_row(dest_table, row) for row in rows)
                 if r is not None]

    # 鍵が重なる行は**先に出てきたほうを残す**(VBA は鍵で引いて
    # 最初の1件しか見ない)。あとの行で上書きすると、VBA と違う行で
    # 計算することになる
    converted, result.duplicates = import_specs.drop_duplicates(
        dest_table, converted)
    if result.duplicates:
        log.warning("取り込み %s: 鍵が重なる行を %d 件飛ばしました"
                    "(先の行を残しました): %s", dest_table,
                    len(result.duplicates), ", ".join(result.duplicates[:5]))

    columns = [dest for dest, _, _ in specs]
    placeholders = ", ".join("?" for _ in columns)
    quoted = ", ".join(source_db.quote_identifier(c) for c in columns)
    sql = (f"INSERT OR REPLACE INTO {source_db.quote_identifier(dest_table)}"
           f" ({quoted}) VALUES ({placeholders})")

    # 1テーブルを1トランザクションで入れ替える。途中で落ちたら丸ごと戻す
    # (半分だけ新しい写し、が残らない)
    try:
        with conn:
            conn.execute(f"DELETE FROM {source_db.quote_identifier(dest_table)}")
            conn.executemany(sql, [[row[c] for c in columns] for row in converted])
    except Exception as exc:  # noqa: BLE001 - 理由を残して次のテーブルへ進む
        result.error = f"手元のDBへ書けませんでした: {exc}"
        log.warning("取り込み失敗 %s: %s", dest_table, exc)
        return result

    result.ok = True
    result.rows = len(converted)
    _record_import(conn, dest_table, source_path, len(converted), stamp)
    log.info("取り込み %s: %d 件 (%s)", dest_table, len(converted), source_path.name)
    return result


def _record_import(conn: Any, table: str, source_path: Path, rows: int,
                   stamp: str = "") -> None:
    """「いつ・どこから・何件・どの中身から(目印)」を残す。画面に出して古さを判断させる。

    `stamp` は**読む前に**取った目印(`source_stamp`)。確かめた日時も今にする。
    """
    now = db.now_db_string()
    with conn:
        conn.execute(
            "INSERT OR REPLACE INTO 取り込み履歴"
            " (テーブル名, 取り込み日時, 元ファイル, 件数, 元の目印, 確かめた日時)"
            " VALUES (?, ?, ?, ?, ?, ?)",
            (table, now, str(source_path), rows, stamp, now))
    _bump()


# ==================================================================
# 任意の表を写す(統合 1.2.50・現場の指示)
# ==================================================================
#: 共有にまだ任意の表が無いときの知らせ(取り込みの誤りにはしない)
OPTIONAL_ABSENT_NOTE = ("まだ作っていません(マスタ管理で作れます。作るまでは 1C1103 のパレット種類を"
                        "プログラムの規則で決めます)")


def _import_optional(conn: Any, table: str, path: Path) -> TableResult:
    """梱包資材マスタの**任意の表**を写す。共有に無ければ手元の表を消し、誤りにしない。

    - 共有に表がある … 手元の表を作って写す(写せなければ誤り。前から手元にあった表はそのまま残す)
    - 共有に表が無い … 手元の表を消す(= プログラムの規則に戻る)。結果は ok で「まだ作っていません」

    どちらでも取り込み履歴に目印を残す ── 残さないと、表が無いあいだ毎回読みに行き、
    共有で表が作られたときに「変わった」と気づけない(`check_sources`)。
    """
    stamp = source_stamp(path)                  # 読む前に取る(統合 1.2.49)
    try:
        # 共有が読めるかを先に確かめる。読めないのに「表が無い」と取り違えて手元の表を消さない
        source_db.read_query(path, "SELECT 1 FROM sqlite_master LIMIT 1")
    except source_db.SourceError as exc:
        return TableResult(table=table, ok=False, source=str(path),
                           error=f"取り込み元を読めません: {exc}")
    available = set(source_db.columns(path, table))
    if not available:
        try:
            if table == outer_condition.TABLE:
                outer_condition.drop_local(conn)
        except Exception as exc:  # noqa: BLE001 - 理由を残して次へ
            return TableResult(table=table, ok=False, source=str(path),
                               error=f"手元のDBへ書けませんでした: {exc}")
        _record_import(conn, table, path, 0, stamp)
        return TableResult(table=table, ok=True, rows=0, source=str(path),
                           note=OPTIONAL_ABSENT_NOTE)
    missing = import_specs.missing_columns(table, available)
    if missing:
        # 作る前に断る。空の表を作ってしまうと「表はあるが行が無い」= マスタの値のまま、に化ける
        return TableResult(table=table, ok=False, source=str(path),
                           error="取り込み元に列がありません: " + ", ".join(missing))
    existed = outer_condition.present(conn)
    outer_condition.ensure_local(conn)
    result = _import_table(conn, dest_table=table, source_path=path, source_table=table)
    if not result.ok and not existed:
        outer_condition.drop_local(conn)
    return result


def _import_master_table(conn: Any, table: str, path: Path) -> TableResult:
    """梱包資材マスタの1表を写す(任意の表も含む)。"""
    if table in import_specs.OPTIONAL_MASTER_SPECS:
        return _import_optional(conn, table, path)
    return _import_table(conn, dest_table=table, source_path=path, source_table=table)


def master_tables() -> list[str]:
    """梱包資材マスタから写す表ぜんぶ(任意の表も含む。統合 1.2.50)。"""
    return list(import_specs.MASTER_IMPORT_SPECS) + list(import_specs.OPTIONAL_MASTER_SPECS)


# ==================================================================
# まとめて取り込む
# ==================================================================
def import_master(conn: Any) -> list[TableResult]:
    """梱包資材マスタの表(包装仕様・パレット・リプラサイズ・班員名簿と、任意の表)を写す。

    マスタが見つからないときの誤りは、そろっていないと困る4表ぶんだけ出す(任意の表は出さない)。
    """
    path = find_master_db()
    if path is None:
        directory = config.master_db_dir()
        return [TableResult(table=t, ok=False,
                            error=f"梱包資材マスタが見つかりません: {directory}")
                for t in import_specs.MASTER_IMPORT_SPECS]
    with _importing():
        return [_import_master_table(conn, t, path) for t in master_tables()]


def import_lot(conn: Any) -> list[TableResult]:
    """仕掛台帳の2テーブルを写す。"""
    results: list[TableResult] = []
    with _importing():
        for table in import_specs.LOT_IMPORT_SPECS:
            results.append(_import_lot_one(conn, table))
    return results


def _import_lot_one(conn: Any, table: str) -> TableResult:
    """仕掛台帳の1表を写す(見つからなければ理由)。"""
    path = find_lot_db(table)
    if path is None:
        dirs = " / ".join(str(d) for d in config.lot_db_dirs())
        return TableResult(
            table=table, ok=False,
            error=f"{config.LOT_DB_FILES[table]} が見つかりません: {dirs}"
              + (f"(第2パスは {config.lot_db2_files()[table]} で探しました)"
                 if len(config.lot_db_dirs()) > 1 else ""))
    return import_lot_table(conn, table, path)


def import_all(conn: Optional[Any] = None) -> SyncResult:
    """マスタと仕掛台帳をまとめて写す。

    **1つ失敗しても残りは続ける。** 仕掛台帳に届かない端末でも、
    マスタさえ読めれば包装仕様の閲覧はできる。
    """
    result = SyncResult(started_at=db.now_db_string())

    def run(target) -> list[TableResult]:
        tables = import_master(target) + import_lot(target)
        coil = coil_results.import_results(target)      # 実績の形式を読む設定のときだけ(統合 1.2.29)
        return tables + ([coil] if coil is not None else [])

    # 押したときは、覚えている写しを捨ててから読む(統合 1.2.49)
    forget_copies(source_paths())
    with _importing():
        if conn is None:
            with db.connect() as own:
                db.apply_schema(own)
                result.tables = run(own)
        else:
            result.tables = run(conn)
    result.finished_at = db.now_db_string()

    if result.ok:
        log.info("取り込み完了: %d 件", result.total_rows)
    else:
        for failure in result.failures:
            log.warning("取り込めませんでした %s: %s", failure.table, failure.error)
    return result


# ==================================================================
# 写しの古さ
# ==================================================================
def last_import(conn: Any) -> dict[str, dict[str, Any]]:
    """テーブルごとの「いつの写しか」。画面に出す。"""
    rows = db.fetch_all(
        conn, "SELECT テーブル名, 取り込み日時, 元ファイル, 件数, 元の目印, 確かめた日時"
              " FROM 取り込み履歴") or []
    return {r["テーブル名"]: dict(r) for r in rows}


def expected_tables() -> list[str]:
    """写しがそろっているかを見る表。コイル実績は読む設定のときだけ入れる(統合 1.2.49)。"""
    tables = list(import_specs.MASTER_IMPORT_SPECS) + list(import_specs.LOT_IMPORT_SPECS)
    if coil_results.wanted():
        tables.append(coil_results.TABLE)
    return tables


def imported_at(conn: Any) -> str:
    """いちばん古い取り込み日時(写した日時)。1つでも未取り込みなら空文字。

    **いちばん古いほうを出す。** 「マスタは今朝、仕掛は3日前」の
    ときに新しいほうを出すと、古さに気づけない。
    コイル実績も、読む設定のときは入れる(統合 1.2.49。以前は入れておらず、実績が起動時の写しのままでも気づけなかった)。
    """
    expected = expected_tables()
    records = last_import(conn)
    if not all(t in records for t in expected):
        return ""
    return min(records[t]["取り込み日時"] for t in expected)


def checked_at(conn: Any) -> str:
    """取り込み元を最後に確かめた日時(いちばん古いもの)。写した日時とは別(統合 1.2.49)。

    目印が同じと確かめられた表は、確かめた日時だけが進む(写し直さない)。
    """
    expected = expected_tables()
    records = last_import(conn)
    if not all(t in records for t in expected):
        return ""
    return min((records[t].get("確かめた日時") or records[t]["取り込み日時"]) for t in expected)


# ------------------------------------------------------------------
# 取り込み元を軽く確かめる(統合 1.2.49・現場の指摘)
# ------------------------------------------------------------------
@dataclass
class SourceCheck:
    """取り込み元を確かめた結果。**ファイルの大きさ・更新時刻を見ただけ**(中身は読まない)。"""
    # 取り込み元が変わっている表(手元はまだ前の写し)
    master: list[str] = field(default_factory=list)
    lot: list[str] = field(default_factory=list)
    coil: bool = False
    # 取り込み元が変わっている任意の表(統合 1.2.50)。取り込み直すが、「古い」には数えない
    optional: list[str] = field(default_factory=list)
    # 確かめられなかった表(取り込み元が見つからない・届かない)
    unverified: list[str] = field(default_factory=list)
    master_path: Optional[Path] = None

    @property
    def changed(self) -> bool:
        return bool(self.master or self.lot or self.coil)


def check_sources(conn: Any) -> SourceCheck:
    """取り込み元の目印を、手元に残した目印と比べる。**一致しなければ「変わった」**。

    一致した表は「確かめた日時」を今にする(写した日時はそのまま)。まだ取り込んでいない表は
    ここでは数えない(`stale_reason` が「まだ取り込んでいません」と言う。取り込めない表を
    帯を開くたびに読みに行かないため)。
    """
    known = last_import(conn)
    check = SourceCheck()
    same: list[str] = []

    def judge(table: str, ok: bool, bucket: list[str]) -> None:
        (same if ok else bucket).append(table)

    master = find_master_db()
    check.master_path = master
    for table in import_specs.MASTER_IMPORT_SPECS:
        if table not in known:
            continue
        if master is None:
            check.unverified.append(table)
        else:
            judge(table, _same_source(known.get(table), master), check.master)
    # 任意の表(統合 1.2.50)。**まだ見たことが無くても**取り込みに回す(前の版の手元DB・
    # 自動取り込みを切っている端末でも、1回で表の有る無しが分かる)。届かなければ何もしない
    if master is not None:
        for table in import_specs.OPTIONAL_MASTER_SPECS:
            judge(table, _same_source(known.get(table), master), check.optional)

    for table in import_specs.LOT_IMPORT_SPECS:
        if table not in known:
            continue
        path = find_lot_db(table)
        if path is None:
            check.unverified.append(table)
            continue
        second = find_second_lot(table)
        second = second if second is not None and second != path else None
        ok = _same_source(known.get(table), path) and (
            second is None or _same_source(known.get(SECOND_NOTE.format(table=table)), second))
        judge(table, ok, check.lot)

    if coil_results.wanted() and coil_results.TABLE in known:
        files = coil_results.sources()
        if not files:
            check.unverified.append(coil_results.TABLE)
        else:
            ok = _same_source(known.get(coil_results.TABLE), files[0]) and (
                len(files) < 2 or _same_source(known.get(coil_results.SECOND_NOTE), files[1]))
            if ok:
                same.append(coil_results.TABLE)
            else:
                check.coil = True

    _touch_checked(conn, same)
    return check


def _touch_checked(conn: Any, tables: list[str]) -> None:
    """目印が同じと確かめた表の「確かめた日時」を今にする。"""
    if not tables:
        return
    try:
        with conn:
            conn.executemany("UPDATE 取り込み履歴 SET 確かめた日時 = ? WHERE テーブル名 = ?",
                             [(db.now_db_string(), t) for t in tables])
    except Exception:                          # noqa: BLE001 - 記録できなくても確かめた結果は使える
        log.warning("確かめた日時を残せませんでした", exc_info=True)


@dataclass
class Freshness:
    """帯・計算Start に出す「写しの新しさ」。"""
    importing: bool = False
    check: SourceCheck = field(default_factory=SourceCheck)
    # その場で取り込み直した表
    refreshed: list[str] = field(default_factory=list)
    # 取り込み直そうとして失敗したもの(理由)
    errors: list[str] = field(default_factory=list)
    reason: str = ""

    @property
    def lot_changed(self) -> bool:
        return bool(self.check.lot)

    @property
    def message(self) -> str:
        return STALE_MESSAGES.get(self.reason, "")

    @property
    def refreshed_note(self) -> str:
        """その場で取り込み直したことを人の言葉で。無ければ空文字。"""
        if not self.refreshed:
            return ""
        return ("共有の取り込み元が変わっていたので、手元へ取り込み直しました("
                + "・".join(self.refreshed) + ")")


def freshness(conn: Any, *, refresh: bool = True) -> Freshness:
    """取り込み元を軽く確かめ、**小さい表は変わっていればその場で取り込み直す**(統合 1.2.49)。

    - 梱包資材マスタの表(包装仕様・パレット・リプラサイズ・班員名簿)・コイル実績
      … 小さいので取り込み直す(マスタ管理の画面は共有を直に読むので、ここを飛ばすと
      「画面は新しいのに計算は古い」が起きる)
    - 仕掛台帳 … 大きいので取り込み直さず、帯で知らせる(`Freshness.lot_changed`)

    取り込み中(起動時の自動取り込みなど)は何もしない(`importing`)。
    """
    out = Freshness()
    # ほかの要求が取り込み直している最中なら、終わるのを待ってから確かめる(すぐ終わる。
    # 待たずに「取り込み中」と言うと、帯を開いただけで計算Start が断られることがある)
    if is_importing() or not _refresh_lock.acquire(timeout=REFRESH_WAIT_SEC):
        out.importing = True
        out.reason = STALE_IMPORTING
        return out
    try:
        out.check = check_sources(conn)
        if refresh and (out.check.master or out.check.coil or out.check.optional):
            _refresh_small(conn, out)
            out.check = check_sources(conn)
        out.reason = stale_reason(conn, check=out.check)
    finally:
        _refresh_lock.release()
    return out


def _refresh_small(conn: Any, out: Freshness) -> None:
    """マスタの表・コイル実績を取り込み直す(写しを捨ててから読む)。

    `_refresh_lock` を持って呼ぶ。計算Start は同じ鍵を待ってから確かめるので、ここは
    「取り込み中」にしない(1表ずつでも、計算は取り込み直しが終わってから始まる)。
    """
    tables = out.check.master + out.check.optional
    if tables and out.check.master_path is not None:
        forget_copies([out.check.master_path])
        for table in tables:                        # 変わった表だけ(取り込めない表を毎回読み直さない)
            done = _import_master_table(conn, table, out.check.master_path)
            if done.ok:
                out.refreshed.append(table)
            else:
                out.errors.append(f"{table}: {done.error}")
    if out.check.coil:
        forget_copies(coil_results.sources())
        done = coil_results.import_results(conn)
        if done is not None and done.ok:
            out.refreshed.append(coil_results.TABLE)
        elif done is not None:
            out.errors.append(f"{coil_results.TABLE}: {done.error}")
    if out.refreshed:
        log.info("取り込み元が変わっていたので取り込み直しました: %s", "・".join(out.refreshed))
    for error in out.errors:
        log.warning("取り込み直せませんでした %s", error)


def reimport_lot(conn: Any) -> list[TableResult]:
    """仕掛台帳だけを取り込み直す(帯の「いま取り込む」。写しを捨ててから読む。統合 1.2.49)。"""
    with _importing():
        results = []
        for table in import_specs.LOT_IMPORT_SPECS:
            forget_copies([find_lot_db(table), find_second_lot(table)])
            results.append(_import_lot_one(conn, table))
        return results


# 「古い」の理由。画面は文言ではなくこれを見る
STALE_NONE = ""
STALE_NOT_IMPORTED = "not_imported"
STALE_DATE_CHANGED = "date_changed"
STALE_SOURCE_CHANGED = "source_changed"
STALE_IMPORTING = "importing"

STALE_MESSAGES = {
    STALE_NOT_IMPORTED: "まだ取り込んでいません",
    # 取り込み元に届かず、目印で確かめられないときだけ日付で見る(統合 1.2.49)
    STALE_DATE_CHANGED: "取り込み元を確かめられないまま、取り込んだ日から日付が変わりました",
    STALE_SOURCE_CHANGED: "共有の取り込み元が変わりました(手元はまだ前の写しです)",
    STALE_IMPORTING: "取り込み中です",
}


def stale_reason(conn: Any, *, check: Optional[SourceCheck] = None) -> str:
    """写しが古いか、その理由。古くなければ空文字。

    **取り込み元が変わったか(目印が一致するか)で見る**(統合 1.2.49・現場の指摘)。
    以前は「取り込んだ日と今日が違えば古い」で見ていて、当てにならなかった:
    - 出るべきときに出ない … 朝に取り込んだあと、昼に共有のマスタが直っても「古い」が出ない
    - 出なくてよいときに出る … マスタが何日も変わらないのに、毎朝「古い」が出続ける

    取り込み元に届かず目印で確かめられない表だけは、いままでどおり日付で見る
    (確かめた日時、無ければ取り込んだ日時が今日でなければ古い。経過時間では見ない ──
    23:50 に取り込んで翌 00:10 に使ったとき、経過20分でも前日の写しなので)。
    """
    stamp = imported_at(conn)
    if not stamp:
        return STALE_NOT_IMPORTED
    try:
        datetime.strptime(stamp, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        # 読めない日時は、古いものとして扱う(黙って新しい扱いにしない)
        return STALE_NOT_IMPORTED
    check = check if check is not None else check_sources(conn)
    if check.changed:
        return STALE_SOURCE_CHANGED
    if check.unverified:
        records = last_import(conn)
        seen = [records[t].get("確かめた日時") or records[t]["取り込み日時"]
                for t in check.unverified if t in records]
        try:
            oldest = min(datetime.strptime(v, "%Y-%m-%d %H:%M:%S") for v in seen) if seen else None
        except ValueError:
            return STALE_NOT_IMPORTED
        if oldest is not None and oldest.date() != datetime.now().date():
            return STALE_DATE_CHANGED
    return STALE_NONE


def stale_message(conn: Any) -> str:
    """「古い」の理由を人の言葉で。古くなければ空文字。

    **なぜ古いのかを出す。** 20分前に取り込んだ人が日付をまたいだ
    だけで「古い」と言われると、壊れたように見える。
    """
    return STALE_MESSAGES.get(stale_reason(conn), "")


def is_stale(conn: Any) -> bool:
    """写しが古いか。未取り込みも「古い」に含める。"""
    return stale_reason(conn) != STALE_NONE


# ==================================================================
# 起動時の自動取り込み
# ==================================================================
def find_lot_dbs() -> dict[str, Path]:
    """仕掛台帳のファイルを、見つかったぶんだけ返す。診断とログに使う。"""
    found: dict[str, Path] = {}
    for table, name in config.LOT_DB_FILES.items():
        path = find_lot_db(table)
        if path is not None:
            found[name] = path
    return found


def _unchanged(record: Optional[dict], path: Path) -> bool:
    """前に取り込んだときから、そのファイルが変わっていないか(同じ場所・目印が**一致**)。

    以前は「更新時刻が取り込み日時より前か」で見ていた(統合 1.2.49 で一致に改めた)。
    """
    return _same_source(record, path)


def auto_import(conn: Any, *, progress=None) -> SyncResult:
    """**変わったものだけ**取り込む(起動時に走る)。

    毎回すべて読み直すと、共有が重い日は起動が数十秒になる。
    元ファイルの目印(大きさ・更新時刻)を覚えておいて、**一致する**ものは飛ばす。

    `progress` は `(done, total, label)` で呼ばれる。起動待機画面が
    これを読んで「いま何をしているか」を出す。
    """
    with _importing():
        return _auto_import(conn, progress=progress)


def _auto_import(conn: Any, *, progress=None) -> SyncResult:
    db.apply_schema(conn)

    plan: list[tuple[str, Path, str]] = []
    master = find_master_db()
    if master is not None:
        for table in master_tables():               # 任意の表も(統合 1.2.50)
            plan.append((table, master, table))
    for table in import_specs.LOT_IMPORT_SPECS:
        path = find_lot_db(table)
        if path is not None:
            plan.append((table, path, import_specs.LOT_SOURCE_TABLE))

    known = last_import(conn)
    result = SyncResult(started_at=db.now_db_string())
    total = len(plan)
    same: list[str] = []

    for i, (dest, path, _source_table) in enumerate(plan, start=1):
        if progress is not None:
            progress(i - 1, total, f"{dest} を取り込み中")

        # 元が変わっていなければ飛ばす。仕掛台帳は2つ目(足す分)も変わっていないときだけ
        record = known.get(dest)
        is_lot = dest in import_specs.LOT_IMPORT_SPECS
        second = find_second_lot(dest) if is_lot else None
        second = second if second is not None and second != path else None
        if _unchanged(record, path) and (
                second is None or _unchanged(known.get(SECOND_NOTE.format(table=dest)), second)):
            result.tables.append(TableResult(
                table=dest, ok=True, rows=int(record.get("件数", 0)), source=str(path)))
            same.append(dest)
            continue

        if is_lot:
            result.tables.append(import_lot_table(conn, dest, path))
        else:
            result.tables.append(_import_master_table(conn, dest, path))

    # コイル実績(実績の形式を読む設定のときだけ。統合 1.2.29)。元が変わっていなければ飛ばす
    if coil_results.wanted():
        files = coil_results.sources()
        record = known.get(coil_results.TABLE)
        unchanged = bool(files) and _unchanged(record, files[0]) and (
            len(files) < 2 or _unchanged(known.get(coil_results.SECOND_NOTE), files[1]))
        if unchanged:
            result.tables.append(TableResult(table=coil_results.TABLE, ok=True,
                                             rows=int(record.get("件数", 0)), source=str(files[0])))
            same.append(coil_results.TABLE)
        else:
            coil = coil_results.import_results(conn)
            if coil is not None:
                result.tables.append(coil)

    # 飛ばした表も「確かめた」ことは残す(写した日時はそのまま。統合 1.2.49)
    _touch_checked(conn, same)

    if not plan:
        log.warning("取り込み元が1つも見つかりません: マスタ=%s 台帳=%s",
                    config.master_db_dir(), config.lot_db_dir())

    if progress is not None:
        progress(total, total, "取り込み完了")
    result.finished_at = db.now_db_string()
    return result


# ==================================================================
# 1表だけ取り込み直す
# ==================================================================
def reimport_table(conn: Any, table: str, *,
                   path: Optional[Path] = None) -> TableResult:
    """その表だけを取り込み直して、手元を取り込み元に追いつかせる。

    マスタ管理が取り込み元へ書いたあとに呼ぶ(`master_admin._follow`)。
    ここを飛ばすと、**取り込み元は直っているのに画面の動きは変わらない**
    ── 直したのに効かない、がいちばんたちが悪い。

    まとめて取り込むのと違い、**触った1表だけ**にする。他の表まで
    読み直すと、1行直すたびに共有フォルダを5回叩くことになる。
    読む前に、覚えている写しを捨てる(統合 1.2.49。目印が同じでも中身だけ変わることがある)。
    """
    if import_specs.master_specs(table):           # 任意の表も(統合 1.2.50)
        found = path or find_master_db()
    elif table in import_specs.LOT_IMPORT_SPECS:
        found = path or find_lot_db(table)
    else:
        return TableResult(table=table, ok=False,
                           error=f"取り込みの定義がありません: {table}")

    if found is None:
        return TableResult(table=table, ok=False,
                           error="取り込み元が見つかりません")

    db.apply_schema(conn)
    with _importing():
        if table in import_specs.LOT_IMPORT_SPECS:
            forget_copies([found, find_second_lot(table)])
            return import_lot_table(conn, table, found)
        forget_copies([found])
        return _import_master_table(conn, table, found)


def catch_up_master(conn: Any, path: Path, *, skip: str = "") -> None:
    """同じ梱包資材マスタのほかの表も、目印が変わっていれば取り込み直す(統合 1.2.49)。

    マスタ管理で1表を直すと、ファイルの目印が変わる。ほかの表の記録は前の目印のままなので、
    放っておくと次に帯を開いたとき「変わっていた」と取り込み直し、直した本人に
    「共有の取り込み元が変わっていたので…」と出てしまう。写しは同じものを使うので軽い。
    """
    known = last_import(conn)
    with _importing():
        for table in master_tables():               # 任意の表も(統合 1.2.50)
            if table != skip and not _same_source(known.get(table), path):
                _import_master_table(conn, table, path)
