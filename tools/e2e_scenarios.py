#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""3機能を現場と同じように使う通し試験(手動・開発用の PC)

    python tools/e2e_scenarios.py                 # 一連の流れ → 交互に使う → 放置する
    python tools/e2e_scenarios.py --quick         # 放置を短くする(動きの確認用)
    python tools/e2e_scenarios.py --only flow     # flow / print / alternate / multitab / storage / master / trail / theme / manual / idle / rapid のどれかだけ

【何をするか】
本物の起動(`start_app.py`)で統合アプリを立て、実ブラウザ(Chromium)で統合画面を開き、
**画面を操作して**3機能を使う。取り込み元(仕掛台帳・梱包課共有・梱包資材マスタ)は
試験用に作る。現場と同じく、梱包明細と資材計算は**同じ台帳フォルダ**を読み、
梱包明細の共有の設定は**梱包資材マスタのフォルダ**に置く。

1. 一連の流れ
   梱包明細  ロット → 重量 → 条番号表示 → 積み → 出力 → 帳票(別タブ)→ 書き足し
   ペナラベル サイズ → 検査番号・重量 → 重量反映 → 積み本数 → 風袋計算 → 小ラベル印刷(別タブ)
             → 全サイズ(反映 → 印刷ビュー)
   資材計算  担当者 → ロット → 検入数 → 計算 → チェックリストへ → 印刷(別タブ)→ 発注票 → 履歴
2. 交互に使う
   3機能を途中まで入れてはタブを替え、戻って続きをする。値が消えない・混ざらない・
   断られないこと。タブの行き来を繰り返し、梱包明細と資材計算の取り込みを同時に走らせる
3. 放置する
   a. 表のまま放置(心拍だけが流れる)
   b. 裏に回して放置(ブラウザがタブを凍結したのと同じ ── 裏に回った合図のあと
      ページの JavaScript を止める)
   c. スリープ(サーバのプロセスも止め、ブラウザも止める)
   c2. スリープから戻ったが、サーバだけ先に動き、画面は 100 秒止まったまま(統合 1.2.56)
   c3. 確認の窓(confirm)を出したまま 100 秒(統合 1.2.56)
   d. 統合画面を閉じ、帳票の別タブだけを残して放置
   f. 本ツール以外のタブを開き、本ツールのタブが裏で捨てられる → 戻る(読み直し)
   e. 全部閉じて放置 → 自分で終わる → 翌朝と同じく開き直す(残るもの・消えるもの)
      (f で捨てられた前の画面が残っていても、自分で終わること)
   放置のあと、3機能とも「断られた」「接続なし」「開き直してください」が出ず、
   そのまま操作を続けられること
4. 保存先(`--only storage`)
   3機能の設定を開き、「このPCに保存(このPCで引き継ぐ)」と「複数のPCで共有」が
   ファイルの場所つきで出ていること(`--shots フォルダ` で画面の写しも残す)
5. マスタ管理(`--only master`)
   3機能のマスタ管理で、見出し(列名)を押すと昇順 → もう一度押すと降順に並ぶこと
6. エラーの後追い(`--only trail`)
   画面でわざとエラーを起こし、番号が出る → 「中身を見る」で記録(なぜなぜの欄つき)が開く。
   上の帯の「ログ」で出力先を変えると、すぐ PC の名前のフォルダに書かれる → 配布設定に
   書き出す・消す → 既定に戻す
7. 配色(`--only theme`)
   ブラウザの外観がダークのとき、3機能の背景が暗くそろう。印刷用のタブとバーコードは白地のまま。
   上の帯の「画面の色」でブラウザと逆を選んでも3機能と統合画面がそろい、読み直しても残る
9. 説明書(`--only manual`。統合 1.2.0)
   上の帯の「説明書」で、見ているタブのツールの説明書が開き、写真がすべて出て版が出る。
   ダイアログの中で冊を移れ、「別の窓で開く」でその冊が別のタブで開く。ダウンロードにならない
8. 印刷する(`--only print`。一連の流れのあと)
   3機能の印刷プレビュー11画面すべてで「印刷する」が1つあり、押すと印刷のダイアログが開き
   (window.print が呼ばれた回数を数える)、印刷のときには隠れる(紙に出ない)

10. 続けて印刷する・すぐ切り替えて計算する(`--only rapid`。統合 1.2.20)
   3機能の印刷を待たずに続けて押す(3回)→ 3枚とも開き中身が混ざらない・どれも1回ずつ刷れる。
   ペナラベルの印刷の確かめを6秒読んでから押しても開く(ポップアップのブロックは入れたまま)。
   計算を押してすぐ別のツールへ・2機能から同時に計算・入れてすぐ計算(欄を離れたときの送信を遅らせる)
   → どの結果も正しい。印刷の画面のあいだの心拍の合図・デスクトップ版の「印刷の画面を出せませんでした」

【要るもの】Playwright と Chromium(`tools/smoke_shell.py` と同じ)。現場の PC には要らない。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import signal
import socket
import sqlite3
import subprocess
import sys
import tempfile
import time
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
results: list = []


def pena_material_csv(home: Path) -> str:
    """ペナラベルの資材マスタ(試験用)。同梱のサンプルの写しを「設定で指定した CSV」として読ませる。

    統合 1.2.31 から、ペナラベルは同梱の data/資材重量.csv(試験用のサンプル値)を読まない
    (共有の資材マスタに届かないだけでサンプルに落ち、間違った NW/GW が刷れたため)。
    """
    dest = Path(home) / "pena_material" / "資材重量.csv"
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(ROOT / "modules" / "packing_pena_label" / "data" / "資材重量.csv", dest)
    return str(dest)


def default_password(rel: str, name: str) -> str:
    """既定のパスワードを、その機能の設定のファイルから読む(試験・道具に文字で書かない)。

    取り込むと起動時の処理が走るものもあるので、ファイルの字面から拾う。"""
    src = (ROOT / rel).read_text(encoding="utf-8")
    m = re.search(rf'^{name}\s*=\s*(?:os\.environ\.get\([^,]+,\s*)?"([^"]*)"', src, re.M)
    if not m:
        raise RuntimeError(f"{rel} に {name} が見つかりません")
    return m.group(1)


MATERIAL_PASSWORD = default_password("modules/packing_material_calculation/coil_tool/config.py",
                                     "ADMIN_PASSWORD")
PENA_PASSWORD = default_password("modules/packing_pena_label/app/services/settings.py", "PATH_PASSWORD")
DETAILS_PASSWORD = default_password("modules/packing_details/meisai/config.py", "ADMIN_PASSWORD")


def check(name, ok, detail=""):
    results.append((name, bool(ok), detail))
    print("  %s %-52s %s" % ("OK " if ok else "NG ", name, str(detail)[:110]), flush=True)
    return ok


def wait_until(fn, timeout=15.0, step=0.2):
    end = time.monotonic() + timeout
    last = None
    while time.monotonic() < end:
        try:
            last = fn()
            if last:
                return last
        except Exception as exc:                     # noqa: BLE001 - 待つあいだの失敗は続ける
            last = exc
        time.sleep(step)
    return None


# ======================================================================
# 取り込み元(現場の共有フォルダと同じ形)
# ======================================================================
def _source(path: Path, columns, rows, *, declare=True, created_at=None, table="仕掛"):
    conn = sqlite3.connect(str(path))
    decl = ", ".join(f'"{c}" TEXT' if declare else f'"{c}"' for c in columns)
    conn.execute(f'CREATE TABLE "{table}" ({decl})')
    marks = ", ".join("?" for _ in columns)
    cols = ", ".join(f'"{c}"' for c in columns)
    with conn:
        conn.executemany(f'INSERT INTO "{table}" ({cols}) VALUES ({marks})',
                         [[r.get(c, "") for c in columns] for r in rows])
        if created_at:
            conn.execute('CREATE TABLE "_更新情報" ("項目" TEXT, "値" TEXT)')
            conn.execute('INSERT INTO "_更新情報" VALUES (?, ?)', ("作成日時", created_at))
    conn.close()


LOTS = ("L5160Z0", "A123456")
ORDER_NO = "12345678"
SPEC_NO = "1C0001"


def make_sources(home: Path) -> None:
    """仕掛台帳(梱包明細・資材計算が同じフォルダを読む)・梱包課共有・梱包資材マスタ。"""
    lot, konpo, master = home / "lot", home / "konpo", home / "master"
    for d in (lot, konpo, master):
        d.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y/%m/%d %H:%M:%S")
    _source(lot / "SIKALOT.sqlite3",
            ("ﾛｯﾄ番号", "用途ｺｰﾄﾞ", "用途名", "製造材質", "製造調質", "製造板厚", "製造板幅",
             "ｵｰﾀﾞｰ板厚", "ｵｰﾀﾞｰ板幅", "ｵｰﾀﾞｰ板丈", "設計_設備ｺｰｽ", "実績_設備ｺｰｽ"),
            [{"ﾛｯﾄ番号": n, "用途ｺｰﾄﾞ": "R192", "用途名": "ｱﾝｾﾞﾝﾀｲｻﾝｺｰ", "製造材質": "52S",
              "製造調質": "H34", "製造板厚": "1.985", "製造板幅": "104.0", "ｵｰﾀﾞｰ板厚": "2.0",
              "ｵｰﾀﾞｰ板幅": "104.0", "ｵｰﾀﾞｰ板丈": "0.0", "設計_設備ｺｰｽ": "HOT L-1 LS4",
              "実績_設備ｺｰｽ": "HOT L-1"} for n in LOTS], created_at=stamp)
    # 引当・受注は2つのツールが読む列を合わせ持つ(本物の台帳は全部の列を持っている)
    _source(lot / "SIKAHIKI.sqlite3", ("ﾛｯﾄ番号", "引当番号", "受注番号", "出荷日"),
            [{"ﾛｯﾄ番号": n, "引当番号": "1", "受注番号": ORDER_NO, "出荷日": "03/28"} for n in LOTS],
            created_at=stamp)
    _source(lot / "SIKAODR.sqlite3",
            ("受注番号", "受注材質", "受注調質", "受注板厚", "受注板幅", "用途ｺｰﾄﾞ", "用途名",
             "包装仕様NO", "納入先名称", "取引先名称", "製品単重", "梱包単位_重量", "梱包単位_枚数",
             "ｺｲﾙ外径_MAX", "材質_比重", "コイル内径_目標", "営業納期"),
            [{"受注番号": ORDER_NO, "受注材質": "52S", "受注調質": "H34", "受注板厚": "1.0",
              "受注板幅": "100", "用途ｺｰﾄﾞ": "R192", "用途名": "ﾃｽﾄ用途", "包装仕様NO": SPEC_NO,
              "納入先名称": "納入先A", "取引先名称": "取引先A", "製品単重": "100",
              "梱包単位_重量": "", "梱包単位_枚数": "5", "ｺｲﾙ外径_MAX": "900", "材質_比重": "2.71",
              "コイル内径_目標": "508", "営業納期": "04/01"}], created_at=stamp)
    _source(konpo / "LS4LOT.sqlite3",
            ("ﾛｯﾄ番号", "当工程設計_設備名", "当工程設計_横割数", "当工程設計_縦割数",
             "当工程設計_枚本数", "製造板厚", "製造板幅", "製造板丈"),
            [{"ﾛｯﾄ番号": "L5160Z0", "当工程設計_設備名": "LS4", "当工程設計_横割数": "12",
              "当工程設計_縦割数": "2", "当工程設計_枚本数": "24", "製造板厚": "1.985",
              "製造板幅": "104.0", "製造板丈": "0.0"},
             {"ﾛｯﾄ番号": "A123456", "当工程設計_設備名": "LS4", "当工程設計_横割数": "6",
              "当工程設計_縦割数": "2", "当工程設計_枚本数": "12", "製造板厚": "1.985",
              "製造板幅": "104.0", "製造板丈": "0.0"}], declare=False)
    m = master / "梱包資材マスタ.sqlite3"
    spec = {"包装仕様NO": SPEC_NO, "納入先名称": "", "用途名": "", "パレット": "スカシ", "サイズ": "",
            "種類": "", "リプラサイズ": "30", "コイル間スペーサー": "リプラ",
            "最下部スペーサー": "リプラ", "下本数": "3", "緩衝材": "HB", "ｺｲﾙ間は間紙入": "",
            "梱包単位_重量": "", "重量範囲": "", "梱包単位_枚数": "", "枚数範囲": "以下",
            "梱包総高さ": "", "高さ範囲": ""}
    _source(m, tuple(spec), [spec], table="包装仕様")
    pallets = []
    for w, mark in ((800, "P08"), (900, "P09"), (1000, "P10"), (1100, "P11"),
                    (1200, "P12"), (1300, "P13"), (1350, "P135")):
        for kind, wide in (("スカシ", True), ("全面", True), ("EXPS", False), ("強度UP", None)):
            pallets.append({"種類": kind, "巾下限": "0",
                            "巾上限": str(w) if wide is not False else "9999",
                            "丈下限": "0", "丈上限": str(w) if wide else "9999",
                            "新記号": mark, "Ｗ": str(w)})
    _source(m, ("種類", "巾下限", "巾上限", "丈下限", "丈上限", "新記号", "Ｗ"), pallets,
            table="パレット")
    # ペナラベルのマスタ管理が見る表(並べ替えの試験で使う。ふだんのペナラベルは CSV で動く)
    _source(m, ("管理番号", "梱包資材名", "単位質量", "係数"),
            [{"管理番号": str(i), "梱包資材名": n, "単位質量": w, "係数": "1"}
             for i, (n, w) in enumerate((("テスラピン", "0.5"), ("ﾊｰﾄﾞﾎﾞｰﾄﾞ", "1.25"),
                                         ("ｽﾄﾚｯﾁ", "0.08"), ("アングル", "12.5")), 1)],
            table="資材重量")
    _source(m, ("リプラ長さ", "長さ"),
            [{"リプラ長さ": str(n), "長さ": str(n)} for n in (700, 800, 900, 950, 1000, 1050, 1100, 1150)],
            table="リプラサイズ")
    _source(m, ("管理番号", "苗字", "班", "名前", "読み", "担当ライン"),
            [{"管理番号": "001", "苗字": "高村", "班": "1班", "名前": "高村 敏幸", "読み": "たかむら",
              "担当ライン": "機側"},
             {"管理番号": "002", "苗字": "山田", "班": "2班", "名前": "山田 太郎", "読み": "やまだ",
              "担当ライン": "機側"}], table="班員名簿")


# ======================================================================
# 統合アプリ(本物の起動)
# ======================================================================
class App:
    def __init__(self, home: Path):
        self.home = home
        with socket.socket() as s:
            s.bind(("127.0.0.1", 0))
            self.port = s.getsockname()[1]
        self.conf = home / "app.json"
        self.conf.write_text(json.dumps({
            "app_id": "nlm.coil-packing-tools", "display_name": "コイル梱包ツール",
            # 版は本物のまま(画面の帯・撮った画面に出る版を、配る版と合わせる)
            "version": json.loads((ROOT / "config" / "app.json").read_text(encoding="utf-8"))["version"],
            "local_dir_name": "CoilPackingTools",
            "server": {"host": "127.0.0.1", "port_retry": 3, "roles": {"main": {"port": self.port}}},
        }, ensure_ascii=False), encoding="utf-8")
        self.local = home / "local"
        self.env = dict(os.environ,
                        COIL_PACKING_TOOLS_APP_CONFIG=str(self.conf),
                        COIL_PACKING_TOOLS_LOCAL_DIR=str(self.local),
                        XDG_DATA_HOME=str(home / "xdg-data"), XDG_STATE_HOME=str(home / "xdg-state"),
                        LOCALAPPDATA=str(home / "localappdata"),
                        PACKING_DETAILS_SHARED_DIR=str(home / "master"),
                        PACKING_DETAILS_LOT_DB_DIR=str(home / "lot"),
                        PACKING_DETAILS_KONPO_DB_DIR=str(home / "konpo"),
                        COIL_TOOL_MASTER_DB_DIR=str(home / "master"),
                        COIL_TOOL_LOT_DB_DIR=str(home / "lot"),
                        COIL_PACKING_TOOLS_DISTRIBUTION_DIR=str(self.dist_common(home)),
                        PPL_PREFER_ACCESS="0", PPL_MATERIAL_CSV=pena_material_csv(home))
        for name in ("PYTHONDONTWRITEBYTECODE",):
            self.env.pop(name, None)
        self.proc = None
        self.token = ""

    @staticmethod
    def dist_common(home: Path) -> Path:
        """配布設定(共通: ログ)の書き出し先。アプリのフォルダの 配布設定\\common を汚さない。"""
        return home / "配布設定" / "common"

    @property
    def base(self):
        return "http://127.0.0.1:%d" % self.port

    def start(self, timeout=150):
        self.proc = subprocess.Popen([sys.executable, str(ROOT / "start_app.py"), "--no-browser"],
                                     cwd=str(self.home), env=self.env, stdout=subprocess.DEVNULL,
                                     stderr=subprocess.DEVNULL)
        lock = self.local / "runtime" / "main.lock"
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            if self.proc.poll() is not None:
                return None
            try:
                info = json.loads(lock.read_text(encoding="utf-8"))
                if info.get("pid") == self.proc.pid:
                    self.port, self.token = info["port"], info["token"]
                    h = self.health()
                    if h and h.get("ready"):
                        return h
            except (OSError, ValueError):
                pass
            time.sleep(0.3)
        return None

    def health(self):
        try:
            op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            with op.open(self.base + "/api/health", timeout=3) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception:                            # noqa: BLE001
            return None

    def alive(self):
        return self.proc is not None and self.proc.poll() is None

    def logs(self):
        return "".join(p.read_text(encoding="utf-8", errors="replace")
                       for p in sorted((self.local / "logs").rglob("*.log")))   # logs\coilpacking\ の下(統合 1.2.53)

    def stop(self):
        if self.alive():
            subprocess.run([sys.executable, str(ROOT / "process_manager.py")], cwd=str(self.home),
                           env=self.env, capture_output=True, text=True, encoding="utf-8", errors="replace",
                           timeout=60)
            try:
                self.proc.wait(timeout=20)
            except subprocess.TimeoutExpired:
                self.proc.kill()


# ======================================================================
# 画面の操作
# ======================================================================
OVERLAYS = {
    None: ("#offline", "#restarted", "#stopped"),
    "details": ("#secondScreen", "#takenOver", "#restarted", "#offline"),
    "pena": ("#sgDeny", "#sgLost", "#sgWait"),
    "material": ("#offline",),
}

_VISIBLE = """s => { const e = document.querySelector(s); if (!e || e.hidden) return false;
  for (let n = e; n && n.nodeType === 1; n = n.parentElement) {
    if (n.hidden) return false;
    const st = getComputedStyle(n);
    if (st.display === 'none' || st.visibility === 'hidden') return false;
  }
  return true; }"""


def visible(fr, sel) -> bool:
    return bool(fr.evaluate(_VISIBLE, sel))


def text(fr, sel) -> str:
    return (fr.evaluate("s => { const e = document.querySelector(s); return e ? "
                        "(e.value !== undefined && e.tagName !== 'BUTTON' ? e.value : e.textContent) : null }",
                        sel) or "").strip()


class Shell:
    """統合画面(1枚のタブ)と、その中の3機能の iframe。"""

    def __init__(self, ctx, app: App):
        self.ctx, self.app = ctx, app
        self.page = ctx.new_page()
        self.page.goto(app.base + "/?t=" + app.token)
        wait_until(lambda: all(self.frame(k) for k in ("details", "pena", "material")), 20)

    def frame(self, key):
        return self.page.frame(name=f"cpt-{key}")

    def tab(self, key):
        self.page.click(f"#tab-{key}")
        self.page.wait_for_timeout(250)
        return self.frame(key)

    def trouble(self) -> list:
        """画面に出ている「断られた」「接続なし」などの覆い・帯。"""
        found = []
        for key, sels in OVERLAYS.items():
            fr = self.page if key is None else self.frame(key)
            for sel in sels:
                try:
                    if visible(fr, sel):
                        found.append(f"{key or '外枠'}{sel}")
                except Exception as exc:              # noqa: BLE001
                    found.append(f"{key or '外枠'}{sel}? {exc}")
        mat = self.frame("material")
        if mat is not None and visible(mat, "#gate"):
            found.append("material#gate:" + text(mat, "#gateTitle"))
        if text(self.page, "#conn") != "接続OK":
            found.append("外枠#conn:" + text(self.page, "#conn"))
        return found

    def ready(self, timeout=20):
        return wait_until(lambda: not self.trouble(), timeout)


def pena_toast(fr, want, timeout=10):
    return wait_until(lambda: want in text(fr, "#toast .toast-body")
                      and visible(fr, "#toast"), timeout)


def material_toast(fr, want, timeout=10):
    return wait_until(lambda: want in text(fr, "#toast") and visible(fr, "#toast"), timeout)


# ---- 梱包明細 --------------------------------------------------------
def details_open_lot(fr, lot):
    fr.fill("#lotNo", "")
    fr.type("#lotNo", lot, delay=20)
    # **そのロットが開いたのを待つ。** 前のロットの表示が出ているだけでは開いたことにしない
    # (統合 1.2.26。開く前に重量を打ち、前のロットの画面の値が開いたロットに入っていた)
    return wait_until(lambda: details_lot(fr) == lot.upper()
                      and visible(fr, "#lotInfo") and text(fr, "#warisu"), 10)


def details_lot(fr) -> str:
    """梱包明細の画面がいま描いているロット(サーバの返事で描いたもの)。"""
    return fr.evaluate("(window.APP && window.APP.state && window.APP.state.lot_no) || ''")


def details_slow_lot(fr, seconds: float):
    """画面からロットを開く要求だけ、送るのを遅らせる(台帳の読み込みが遅いときと同じ形)。

    Playwright の route は試験の側で待つので、打つ操作まで止まってしまう。画面の中で遅らせる。
    """
    fr.evaluate("""ms => {
        if (!window.__origFetch) window.__origFetch = window.fetch;
        const orig = window.__origFetch;
        window.fetch = (url, opts) => String(url).endsWith('/api/lot')
            ? new Promise(r => setTimeout(r, ms)).then(() => orig(url, opts))
            : orig(url, opts);
    }""", int(seconds * 1000))
    return lambda: fr.evaluate("() => { if (window.__origFetch) window.fetch = window.__origFetch; }")


def details_weights(fr, w1, w2):
    fr.fill("#weight1", str(w1))
    fr.fill("#weight2", str(w2))
    fr.press("#weight2", "Tab")
    fr.wait_for_timeout(400)


def details_weights_late_reply(ctx, fr):
    """丈1を送った返事が、丈2へ移ったあとに届いても、丈2に打つ数字を失わない(統合 1.2.26)。

    以前は返事のたびに重量の欄を作り直していた。丈1 → Tab → 丈2 と打つあいだに丈1の返事が
    届くと、丈2の欄が作り直されて入力の場所(フォーカス)が消え、続けて打った数字がどこにも
    入らなかった(打ちかけの途中なら、そこまでの数字が送られた)。共有フォルダが遅い・PC が
    重いと起きる。
    """
    fr.page.bring_to_front()               # 打っている窓が前にある
    before = (fr.input_value("#weight1"), fr.input_value("#weight2"))
    undo = _delay(ctx, "**/details/api/weights", 1.0)
    try:
        fr.fill("#weight1", "251")
        fr.click("#weight2")                # 丈1の change が飛ぶ(返事は1秒遅れる)
        fr.wait_for_timeout(1800)           # 遅れた返事が届いて画面が描き直されるまで
        fr.page.keyboard.press("Control+A")
        fr.page.keyboard.type("249", delay=20)
        typed = fr.input_value("#weight2")
    finally:
        undo()
    fr.press("#weight2", "Tab")
    fr.wait_for_timeout(500)
    check("梱包明細: 丈2へ移ったあとに丈1の返事が届いても、丈2に打った重量が入る", typed == "249", typed)
    details_weights(fr, *before)            # 元へ戻す(あとの試験が使う)


def details_slots(fr):
    return fr.evaluate("[...document.querySelectorAll('#stack button.slot')].map(b => b.textContent.trim())")


def details_stack_and_output(fr, keys, resize=False):
    """積んで出力する。`resize` なら積みかけで積み条数(積み本数)を増やして戻し、積んだものが残るか見る
    (梱包明細 VER 0.13.33・現場の求め。VBA は積み条数を変えると全部外れて積み直しになった)。"""
    fr.click("#btnStrands")
    if not wait_until(lambda: visible(fr, "#strandCard"), 8):
        return "条番号が出ない: " + text(fr, "#notice")
    fr.fill("#stackMax", str(len(keys)))
    fr.press("#stackMax", "Tab")
    wait_until(lambda: fr.evaluate("document.querySelectorAll('#stack button.slot').length") == len(keys), 5)
    for i, k in enumerate(keys):
        fr.click(f'#grid button.coil[data-key="{k}"]')
        fr.wait_for_timeout(150)
        if resize and i == 1:
            fr.click('[data-stack="1"]')               # 積み条数を1つ増やす(＋)
            grown = wait_until(lambda: len(details_slots(fr)) == len(keys) + 1, 5)
            kept = details_slots(fr)[:2] == list(keys[:2])
            fr.click('[data-stack="-1"]')              # 戻す(上の空きだけ取れる)
            back = wait_until(lambda: len(details_slots(fr)) == len(keys), 5)
            if not (grown and kept and back and details_slots(fr)[:2] == list(keys[:2])):
                return f"積み条数を変えたら積んだコイルが外れた: {details_slots(fr)}"
    if not wait_until(lambda: fr.evaluate("!document.querySelector('#btnOutput').disabled"), 5):
        return "出力が押せない"
    fr.click("#btnOutput")
    # 重なり・上書きなどの確かめが出たら「はい」
    fr.wait_for_timeout(500)
    if visible(fr, "#confirmDialog"):
        fr.click('#confirmDialog button[value="yes"]')
    ok = wait_until(lambda: "出力しました" in text(fr, "#notice"), 8)
    return "" if ok else "出力できない: " + text(fr, "#notice")


def details_edit_output(fr):
    """出力済み修正(梱包明細 VER 0.13.23)。No2 を出してから、No2 だけを積み直して同じ No で出し直す。

    No1 には触らない(帳票の別タブで開いたままの No1 を、あとの試験が使う)。
    """
    sh_page = fr.page
    sh_page.bring_to_front()
    for k in ("1-10", "1-9", "2-10", "2-9"):
        fr.click(f'#grid button.coil[data-key="{k}"]')
        fr.wait_for_timeout(150)
    wait_until(lambda: fr.evaluate("!document.querySelector('#btnOutput').disabled"), 5)
    fr.click("#btnOutput")
    fr.wait_for_timeout(500)
    if visible(fr, "#confirmDialog"):
        fr.click('#confirmDialog button[value="yes"]')
    check("梱包明細: 2枚目を出す", wait_until(lambda: fr.locator("#outputs .output").count() == 2, 8),
          text(fr, "#notice"))
    fr.check("#pick1")
    fr.check("#pick2")
    check("梱包明細: 出力済みに2つチェックすると「出力済み修正」は押せない",
          fr.evaluate("document.querySelector('#btnEditOutput').disabled")
          and "1枚ずつ" in text(fr, "#editNote"), text(fr, "#editNote"))
    fr.uncheck("#pick1")
    fr.click("#btnEditOutput")
    check("梱包明細: 出力済み修正で No2 のコイルが積み上げに戻る",
          wait_until(lambda: visible(fr, "#editBanner") and "修正して出力" in text(fr, "#btnOutput"), 8)
          and fr.evaluate("[...document.querySelectorAll('#stack button.slot')].map(b => b.textContent)")
          .count("2-9") == 1, text(fr, "#editBanner"))
    # 4段目(2-9)を外し、No1 のコイル(1-12)を押しても積まれない → 2-8 を積む
    fr.click('#stack button.slot >> nth=3')
    fr.wait_for_timeout(300)
    fr.click('#grid button.coil[data-key="1-12"]', force=True)
    fr.wait_for_timeout(300)
    check("梱包明細: 修正中、ほかの No のコイルは積めない",
          fr.evaluate("document.querySelector('#btnOutput').disabled"), text(fr, "#notice"))
    fr.click('#grid button.coil[data-key="2-8"]')
    wait_until(lambda: fr.evaluate("!document.querySelector('#btnOutput').disabled"), 5)
    fr.click("#btnOutput")
    ok = wait_until(lambda: visible(fr, "#confirmDialog") and "No2" in text(fr, "#confirmDialog"), 5)
    if ok:
        fr.click('#confirmDialog button[value="yes"]')
    check("梱包明細: 修正して出力 → 同じ No2 のまま(枚数は2のまま)",
          ok and wait_until(lambda: "No2を修正しました" in text(fr, "#notice"), 8)
          and fr.locator("#outputs .output").count() == 2 and not visible(fr, "#editBanner"),
          text(fr, "#notice"))
    fr.uncheck("#pick2")
    # 積み上げを始めたロットは丈数を変えられない(VER 0.13.24)
    check("梱包明細: 積み上げを始めたら丈数は変えられない(全クリアの案内が出る)",
          fr.evaluate("document.querySelector('#jouSu').disabled")
          and visible(fr, "#jouLockNote") and "全クリア" in text(fr, "#jouLockNote"),
          text(fr, "#jouLockNote"))
    check("梱包明細: 積み上げを始めたら前工程実績数も変えられない(VER 0.13.25)",
          fr.evaluate("document.querySelector('#btnEditZen').disabled") and visible(fr, "#zenLockNote"),
          text(fr, "#zenLockNote"))


# ---- ペナラベル ------------------------------------------------------
def pena_ready(fr, timeout=15):
    """画面の受付(使ってよい画面か)が済み、接続OK になるまで待つ。"""
    return wait_until(lambda: fr.evaluate("document.readyState") == "complete"
                      and not visible(fr, "#sgWait") and not visible(fr, "#sgDeny")
                      and text(fr, "#conn") == "接続OK", timeout)


PENA_EXPECT = {"#nw1": "110.0", "#gw1": "147.0", "#ta1": "910.0"}


def pena_weights(fr, kensa="w111111", w1="10", w2="12"):
    pena_ready(fr)
    if not fr.is_checked("#cbMode"):
        fr.check("#cbMode")
    fr.click('input[name=size][data-kind=cb][data-cb="5"]')
    wait_until(lambda: text(fr, "#lblSize1") and not fr.is_disabled("#weight1"), 8)
    fr.fill("#kensaNo", kensa)
    fr.fill("#weight1", w1)
    fr.fill("#weight2", w2)
    fr.click("button[data-act=applyWeight]")
    return pena_toast(fr, "重量を反映しました")


def pena_tare(fr, counts=("11", "11", "10", "10")):
    for i, n in enumerate(counts, 1):
        fr.fill(f"#coilH{i}", n)
    fr.click("button[data-act=calcTare]")
    if not pena_toast(fr, "計算しました"):
        return "計算できない: " + text(fr, "#toast .toast-body")
    got = {sel: text(fr, sel) for sel in PENA_EXPECT}
    return "" if got == PENA_EXPECT else f"結果が違う: {got}"


# ---- 資材計算 --------------------------------------------------------
def material_ready(fr, timeout=15):
    return wait_until(lambda: not visible(fr, "#gate") and visible(fr, ".shell"), timeout)


def material_worker(fr):
    opts = fr.evaluate("[...document.querySelectorAll('#rb-worker option')].map(o => o.value).filter(Boolean)")
    if not opts:
        return ""
    if text(fr, "#rb-worker") != opts[0]:
        fr.select_option("#rb-worker", opts[0])
        material_toast(fr, "担当者を")
    return opts[0]


def material_ins(fr, value: str) -> None:
    """検入数を入れて、サーバへ送られるまで待つ(`change` で送る。同じ値だと送られない)。"""
    if text(fr, "#ins") in (value, value + ".00"):
        tmp = "1" if value != "1" else "2"
        fr.fill("#ins", tmp)
        with fr.page.expect_response(lambda r: "/api/calc/input" in r.url):
            fr.press("#ins", "Tab")
    fr.fill("#ins", value)
    with fr.page.expect_response(lambda r: "/api/calc/input" in r.url):
        fr.press("#ins", "Tab")


def material_run_right_after_typing(app, fr):
    """検入数を打って Tab で抜けずに、人の速さ(押して 0.15 秒で離す)で「計算Start」を押す。
    以前は押した瞬間に送られた検入数の返事で上のお知らせ欄が動き、ボタンがずれて1回目が空振りした。"""
    page = fr.page
    results = []
    for value in ("12", "10"):
        fr.click("#ins")
        fr.fill("#ins", value)
        box = fr.locator("#run").bounding_box()
        n0 = app.logs().count("/api/calc/run")
        page.mouse.move(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        page.mouse.down()
        page.wait_for_timeout(150)
        page.mouse.up()
        ran = wait_until(lambda: app.logs().count("/api/calc/run") > n0, 5)
        results.append((value, bool(ran), text(fr, "#ins")))
        page.wait_for_timeout(500)
    return all(r[1] for r in results), results


def material_calc(fr, lot, ins="10"):
    fr.fill("#lot", "")
    fr.type("#lot", lot, delay=20)
    if not wait_until(lambda: text(fr, "#o-包装") == SPEC_NO, 10):
        return "ロットが引けない: " + text(fr, "#toast")
    material_ins(fr, ins)
    fr.click("#run")
    if not material_toast(fr, "台数"):
        return "計算できない: " + text(fr, "#toast")
    got = (text(fr, "#r-種類"), text(fr, "#r-台数"))
    return "" if got == ("スカシ", "2") else f"結果が違う: {got}"


# ======================================================================
# 1. 一連の流れ
# ======================================================================
def scenario_flow(app: App, ctx, sh: Shell, note: dict) -> None:
    print("■ 1. 一連の流れ")
    check("統合画面が開き、3機能とも使える状態", sh.ready(), sh.trouble())

    # ---- 梱包明細 ----
    fr = sh.tab("details")
    check("梱包明細: 取り込まれた台帳でロットが引ける", details_open_lot(fr, "L5160Z0"),
          text(fr, "#warisu") or text(fr, "#notice"))
    check("梱包明細: 割数(縦割2 × 横割12)", "縦割2" in text(fr, "#warisu") and "横割12" in text(fr, "#warisu"),
          text(fr, "#warisu"))
    details_weights(fr, 250, 248)
    details_weights_late_reply(ctx, fr)
    err = details_stack_and_output(fr, ["1-12", "1-11", "2-12", "2-11"], resize=True)
    check("梱包明細: 重量 → 条番号表示 → 積み → 出力", not err, err or text(fr, "#notice"))
    with ctx.expect_page() as info:
        fr.click("#btnPrint")
    report = info.value
    report.wait_for_load_state()
    note["report"] = report
    check("梱包明細: 帳票が別タブで開く", "L5160Z0-No1" in report.title(), report.url.split("?")[0])
    check("梱包明細: 帳票に「印刷する」と「このタブを閉じる」(統合 1.2.55)",
          report.locator(".editbar button.print").count() == 1
          and report.locator(".editbar button[data-close-tab]").count() == 1)
    rows = report.evaluate("[...document.querySelectorAll('tr.r-data td')].map(t => t.textContent.trim()).filter(Boolean)")
    check("梱包明細: 帳票に積んだ4本と重量", sum(1 for r in rows if r.endswith("Kg")) == 4, rows[:8])
    report.click('span.edit[data-placeholder="サイズ"]')
    report.keyboard.type("1.0×104")
    report.keyboard.press("Enter")
    check("梱包明細: 帳票の書き足しが保存される",
          wait_until(lambda: "保存しました" in text(report, "#editSaved"), 8), text(report, "#editSaved"))
    details_edit_output(fr)

    # ---- ペナラベル ----
    fr = sh.tab("pena")
    check("ペナラベル: 重量反映", pena_weights(fr), text(fr, "#toast .toast-body"))
    check("ペナラベル: 検査番号・重量が表に出る",
          text(fr, "#lblKensaNo") == "W111111" and text(fr, "#lblWeight1") == "10kg",
          (text(fr, "#lblKensaNo"), text(fr, "#lblWeight1")))
    err = pena_tare(fr)
    check("ペナラベル: 風袋計算(正味・総重量・高さ)", not err, err)
    with ctx.expect_page() as info:
        fr.click("button[data-act=printLabels]")
        _confirm_in_page(fr)
    labels = info.value
    labels.wait_for_load_state()
    check("ペナラベル: 小ラベルの印刷ビューが別タブで開く",
          "/pena/labels/print" in labels.url and labels.locator("section.sheet").count() >= 1, labels.url)
    labels.close()
    # 全サイズ(同じ iframe の中で開く)
    fr.goto(app.base + "/pena/all-size")
    fr.wait_for_load_state()
    combos = fr.evaluate("[...document.querySelectorAll('#asCombo option')].map(o => o.value).filter(Boolean)")
    if combos:
        note["combo"] = combos[0]
        fr.select_option("#asCombo", combos[0])
        fr.wait_for_load_state()
        wait_until(lambda: sh.frame("pena").evaluate("document.readyState") == "complete" and
                   visible(sh.frame("pena"), "#asKen"), 10)
        fr = sh.frame("pena")
        pena_ready(fr)
        fr.fill("#asKen", "W123456")
        fr.check('input[name=asDigit][value="2"]')
        fr.fill("#asW1", "10")
        fr.fill("#asW2", "12")
        fr.click("#asSubmit")
        fr.wait_for_load_state()
        ok = wait_until(lambda: "反映値" in sh.frame("pena").evaluate("document.body.innerText"), 10)
        check("ペナラベル: 全サイズの反映", ok, combos[0])
        fr = sh.frame("pena")
        pena_ready(fr)
        with ctx.expect_page() as info:
            fr.click("#asPrint")
            _confirm_in_page(fr)
        sheet = info.value
        sheet.wait_for_load_state()
        check("ペナラベル: 全サイズの印刷ビューが別タブで開く",
              "/pena/all-size/print" in sheet.url and sheet.locator(".lbl").count() > 0, sheet.url)
        sheet.close()
    else:
        check("ペナラベル: 全サイズの候補がある", False, "候補なし")
    fr.goto(app.base + "/pena/")
    fr.wait_for_load_state()
    pena_ready(sh.frame("pena"))
    check("ペナラベル: 戻ると入力と結果が残っている",
          wait_until(lambda: text(sh.frame("pena"), "#nw1") == "110.0", 8), text(sh.frame("pena"), "#nw1"))

    # ---- 資材計算 ----
    fr = sh.tab("material")
    check("資材計算: 画面が使える(受付を通る)", material_ready(fr), text(fr, "#gateTitle"))
    note["worker"] = material_worker(fr)
    check("資材計算: 担当者を選べる(名簿の機側の人)", note["worker"], note["worker"])
    err = material_calc(fr, "A123456")
    check("資材計算: ロット → 検入数 → 計算", not err, err)
    check("資材計算: 検入数を打ってすぐ「計算Start」を押しても1回で計算する(0.2.34)",
          *material_run_right_after_typing(app, fr))
    fr.check("#size-fixed")
    fr.click("#add")
    check("資材計算: チェックリストへ積む", material_toast(fr, "行目に積みました"), text(fr, "#toast"))
    # 面を移って戻っても、選んだ受注番号が残っている(戻ると「—」になっていた。統合 1.2.30)
    fr.click('nav.rail a.rail__item[href*="/checklist"]')
    fr.wait_for_load_state()
    fr = sh.frame("material")
    material_ready(fr)
    fr.click('nav.rail a.rail__item[href*="/calc"]')
    fr.wait_for_load_state()
    fr = sh.frame("material")
    material_ready(fr)
    check("資材計算: 面を移って戻っても受注番号が残っている",
          wait_until(lambda: text(fr, "#order-no") == ORDER_NO and text(fr, "#lot") == "A123456", 8),
          (text(fr, "#lot"), text(fr, "#order-no")))
    fr.click('nav.rail a.rail__item[href*="/checklist"]')
    fr.wait_for_load_state()
    fr = sh.frame("material")
    material_ready(fr)
    check("資材計算: チェックリストに1行", wait_until(lambda: text(fr, "#sub").startswith("1 /"), 8),
          text(fr, "#sub"))
    with ctx.expect_page() as info:
        fr.click("a#print")
    rep = info.value
    rep.wait_for_load_state()
    check("資材計算: チェックリストの印刷が別タブで開く",
          "A123456" in rep.content() and "チェックリスト" in rep.title() + rep.content(), rep.url.split("?")[0])
    rep.click(".printbar button[data-close-tab]")
    check("資材計算: チェックリストの印刷を「このタブを閉じる」で閉じられる(統合 1.2.55)",
          wait_until(lambda: rep.is_closed(), 5))
    if not rep.is_closed():
        rep.close()
    fr.click("a#to-order")
    fr.wait_for_load_state()
    fr = sh.frame("material")
    material_ready(fr)
    fr.click("#build")
    check("資材計算: 発注票を作る", material_toast(fr, "枚作りました"), text(fr, "#toast") or text(fr, "#note"))
    with ctx.expect_page() as info:
        fr.click("a#print")
    rep = info.value
    rep.wait_for_load_state()
    check("資材計算: 発注票の印刷が別タブで開く", "発注票" in rep.title(), rep.url.split("?")[0])
    # プレビューの「印刷する」で印刷のダイアログが開く(0.2.9。ダイアログの代わりに数える)
    rep.evaluate("() => { window.__printed = 0; window.print = () => { window.__printed++; }; }")
    ok = visible(rep, "#printNow")
    if ok:
        rep.click("#printNow")
    check("資材計算: 発注票のプレビューに「印刷する」", ok and rep.evaluate("window.__printed") == 1)
    rep.close()
    fr.click("#commit")
    check("資材計算: 発注を履歴に残す", material_toast(fr, "履歴に残しました"), text(fr, "#toast"))
    fr.click('nav.rail a.rail__item[href*="/calc"]')
    fr.wait_for_load_state()
    material_ready(sh.frame("material"))
    trouble = [] if sh.ready() else sh.trouble()
    check("一連の流れのあと、どの画面にも断り・接続なしが出ていない", not trouble, trouble)


# ======================================================================
# 説明書(統合 1.2.0): 上の帯の「説明書」で、見ているタブのツールの説明書が開く
# ======================================================================
def _manual_frame(sh):
    # 枠の一覧・URL は Playwright が知らせを受けて更新する。wait_until の time.sleep の間は
    # 受けないので、ここで一瞬受けてから見る(できたばかりの説明書の枠が見つからなかった)
    sh.page.wait_for_timeout(1)
    return next((f for f in sh.page.frames if "/manual/" in f.url), None)


def scenario_manual(app: App, ctx, sh: Shell, note: dict) -> None:
    print("■ 説明書(上の帯の「説明書」)")
    sh.ready()                                   # 統合画面の JS が動いてから押す
    downloads: list = []
    ctx.on("page", lambda p: p.on("download", lambda d: downloads.append(d.url)))
    sh.page.on("download", lambda d: downloads.append(d.url))
    for key, title in (("details", "梱包明細"), ("pena", "ペナラベル"), ("material", "資材計算")):
        sh.tab(key)
        sh.page.click("#manualBtn")
        ok = wait_until(lambda: (_manual_frame(sh) is not None
                                 and f"/manual/{key}" in _manual_frame(sh).url
                                 and title in text(_manual_frame(sh), "h1")), 10)
        fr = _manual_frame(sh)
        check(f"説明書: {title}のタブで押すと{title}の説明書が開く", ok, fr.url if fr else "開かない")
        if fr is not None:
            fr.wait_for_load_state()
            bad = fr.evaluate("[...document.images].filter(i => !i.complete || i.naturalWidth === 0)"
                              ".map(i => i.src)")
            count = fr.evaluate("document.images.length")
            check(f"説明書: {title}の写真がすべて出る", count > 3 and not bad, {"写真": count, "出ない": bad[:2]})
            ver = text(fr, ".vbadge")
            want = (app.health() or {}).get("versions", {}).get(key, "?")
            check(f"説明書: {title}の版が出る", ver == f"VER{want}", ver)
        sh.page.click("#manualClose")
        sh.page.wait_for_timeout(200)
    # ダイアログの中で別の冊へ移り、「別の窓で開く」でその冊が別のタブで開く
    sh.tab("details")
    sh.page.click("#manualBtn")
    wait_until(lambda: _manual_frame(sh) is not None and "/manual/details" in _manual_frame(sh).url
               and "梱包明細" in text(_manual_frame(sh), "h1"), 10)   # 読み終えてから押す
    _manual_frame(sh).click('.mnav a[href^="/manual/pena"]')
    moved = wait_until(lambda: "/manual/pena" in (_manual_frame(sh).url if _manual_frame(sh) else ""), 10)
    check("説明書: ダイアログの中で別の冊へ移れる(ダイアログの形のまま)",
          moved and "embed=1" in _manual_frame(sh).url, _manual_frame(sh).url if _manual_frame(sh) else "")
    with ctx.expect_page() as info:
        sh.page.click("#manualPop")
    pop = info.value
    pop.wait_for_load_state()
    check("説明書: 「別の窓で開く」で、いま出ている冊が別のタブで開く",
          "/manual/pena" in pop.url and "embed" not in pop.url
          and not visible(sh.page, "#manualDialog[open]"), pop.url)
    pop.evaluate("() => { window.__printed = 0; window.print = () => { window.__printed++; }; }")
    pop.click("#manualPrint")
    check("説明書: 「印刷する」で印刷の窓が開く", pop.evaluate("window.__printed") == 1)
    pop.close()
    check("説明書: ダウンロードにならない", not downloads, downloads[:2])
    trouble = [] if sh.ready() else sh.trouble()
    check("説明書を開いたあとも、どの画面にも断り・接続なしが出ていない", not trouble, trouble)


# ======================================================================
# 2. 交互に使う
# ======================================================================
def _api(fr, path, body=None, *, header="X-Tool-Token", screen_key=None):
    """その画面(iframe)から、画面の JS と同じ送り方で API を叩く。"""
    return fr.evaluate("""([path, body, header, screenKey]) => {
        const h = {'Content-Type': 'application/json'};
        const tok = (window.APP && window.APP.token) || window.APP_TOKEN || document.body.dataset.token || '';
        h[header] = tok;
        if (screenKey) h[screenKey[0]] = sessionStorage.getItem(screenKey[1]) || '';
        return fetch(path, {method: 'POST', headers: h, body: JSON.stringify(body || {})})
          .then(r => r.json().then(j => ({status: r.status, body: j})));
    }""", [path, body or {}, header, screen_key])


def scenario_alternate(app: App, ctx, sh: Shell, note: dict) -> None:
    print("■ 2. 交互に使う")
    # --- 途中まで入れてはタブを替える ---
    d = sh.tab("details")
    # 台帳の読み込みが遅く、ロットが開く前に重量を打った(前のロットの画面のまま)。
    # その重量は、開いたロットに入れずに断る(統合 1.2.26。以前は前のロットの丈2の重量が入った)
    old_lot = details_lot(d)
    undo = details_slow_lot(d, 1.5)
    try:
        d.fill("#lotNo", "")
        d.type("#lotNo", "A123456", delay=20)
        details_weights(d, 300, 302)
        opened = wait_until(lambda: details_lot(d) == "A123456", 10)
        d.wait_for_timeout(800)
    finally:
        undo()
    late = d.evaluate("(window.APP.state.weights || []).map(String)")
    check("交互: ロットが開く前に打った重量は、開いたロットに入れず知らせる",
          opened and all(w in ("", "0") for w in late) and "変わったので" in text(d, "#notice"),
          (old_lot, late, text(d, "#notice")[:60]))
    ok = details_open_lot(d, "A123456")
    details_weights(d, 300, 302)
    check("交互: 梱包明細を途中まで(別のロット・重量)", ok and "横割6" in text(d, "#warisu"), text(d, "#warisu"))
    p = sh.tab("pena")
    check("交互: ペナラベルで別の検査番号・重量を反映", pena_weights(p, "x222222", "20", "22"),
          text(p, "#toast .toast-body"))
    m = sh.tab("material")
    material_ready(m)
    m.fill("#lot", "")
    m.type("#lot", "L5160Z0", delay=20)
    wait_until(lambda: text(m, "#o-包装") == SPEC_NO, 10)
    material_ins(m, "10")
    d = sh.tab("details")
    check("交互: 梱包明細に戻るとロット・重量が残っている",
          text(d, "#lotNo") == "A123456" and text(d, "#weight1") == "300" and text(d, "#weight2") == "302",
          (text(d, "#lotNo"), text(d, "#weight1"), text(d, "#weight2")))
    err = details_stack_and_output(d, ["1-6", "2-6"])
    check("交互: 梱包明細の続き(条番号表示 → 積み → 出力)", not err, err or text(d, "#notice"))
    m = sh.tab("material")
    check("交互: 資材計算に戻るとロット・検入数が残っている",
          text(m, "#lot") == "L5160Z0" and text(m, "#ins").startswith("10"), (text(m, "#lot"), text(m, "#ins")))
    m.click("#run")
    check("交互: 資材計算の続き(計算)", material_toast(m, "台数") and text(m, "#r-台数") == "2",
          (text(m, "#r-種類"), text(m, "#r-台数")))
    p = sh.tab("pena")
    check("交互: ペナラベルに戻ると反映した検査番号が残っている", text(p, "#lblKensaNo") == "X222222",
          text(p, "#lblKensaNo"))
    for i, n in enumerate(("11", "11", "10", "10"), 1):
        p.fill(f"#coilH{i}", n)
    p.click("button[data-act=calcTare]")
    check("交互: ペナラベルの続き(風袋計算。重量 20kg × 11本 = 220.0)",
          pena_toast(p, "計算しました") and text(p, "#nw1") == "220.0", text(p, "#nw1"))
    note["pena"] = ("X222222", "220.0")          # 翌朝も残っているはずの値
    d = sh.tab("details")
    with ctx.expect_page() as info:
        d.click("#btnPrint")
    r2 = info.value
    r2.wait_for_load_state()
    check("交互: 梱包明細の帳票(2つ目のロット)", "A123456-No1" in r2.title(), r2.title())
    m = sh.tab("material")
    m.check("#size-fixed")
    m.click("#add")
    check("交互: 資材計算のチェックリストに2行目", material_toast(m, "2 行目に積みました"), text(m, "#toast"))

    # --- 帳票の別タブと統合画面を行き来する ---
    r2.bring_to_front()
    r2.click('span.edit[data-placeholder="LOTNO"]')
    r2.keyboard.type("LT-001")
    r2.keyboard.press("Enter")
    check("交互: 帳票の別タブに行って書き足し(LOTNO)",
          wait_until(lambda: "保存しました" in text(r2, "#editSaved"), 8), text(r2, "#editSaved"))
    # 1つ目のロットの帳票: 別のロットを開くと前のロットの明細は片付けられる
    # (移植元・VBA の `txtLotNo_Change` と同じ)。書き足しは**はっきり断られる**
    old = note.get("report")
    if old is not None and not old.is_closed():
        old.bring_to_front()
        old.click('span.edit[data-placeholder="LOTNO"]')
        old.keyboard.type("X")
        old.keyboard.press("Enter")
        check("交互: 前のロットの帳票の書き足しは「開き直して」と断られる(移植元どおり)",
              wait_until(lambda: "開き直してください" in text(old, "#editSaved"), 8), text(old, "#editSaved"))
        old.close()
    note["report"] = r2
    sh.page.bring_to_front()

    # --- タブをすばやく何度も行き来する ---
    for i in range(36):
        sh.page.click("#tab-" + ("details", "pena", "material")[i % 3])
        sh.page.wait_for_timeout(80)
    sh.page.wait_for_timeout(1500)
    trouble = [] if sh.ready(10) else sh.trouble()
    check("交互: タブを36回行き来しても断り・接続なしが出ない", not trouble, trouble)

    # --- 取り込みを同時に走らせる(同じ台帳フォルダを2つのツールが読む) ---
    now = time.time()
    for f in (app.home / "lot").glob("*.sqlite3"):
        os.utime(f, (now, now))
    p = sh.tab("pena")
    d, m = sh.frame("details"), sh.frame("material")
    d.evaluate("""() => { window.__imp = fetch(APP.base + '/api/import', {method: 'POST',
        headers: {'Content-Type': 'application/json', 'X-Tool-Token': APP.token,
                  'X-Tool-Screen': APP.screen}, body: JSON.stringify({force: true})})
        .then(r => r.json().then(j => ({status: r.status, body: j}))); }""")
    m.evaluate("""() => { window.__imp = fetch(APP_BASE + '/api/settings/import', {method: 'POST',
        headers: {'Content-Type': 'application/json', 'X-App-Token': APP_TOKEN,
                  'X-App-Screen': sessionStorage.getItem('coil.screen') || ''}, body: '{}'})
        .then(r => r.json().then(j => ({status: r.status, body: j}))); }""")
    p.click("button[data-act=calcTare]")
    pena_ok = pena_toast(p, "計算しました")
    di = d.evaluate("window.__imp")
    mi = m.evaluate("window.__imp")
    check("交互: 梱包明細の取り込み(資材計算と同時)", di["status"] == 200 and di["body"].get("ok"),
          (di["body"].get("summary") or di["body"])) 
    check("交互: 資材計算の取り込み(梱包明細と同時)", mi["status"] == 200 and mi["body"].get("ok"),
          mi["body"].get("total", mi["body"]))
    check("交互: 取り込みの最中もペナラベルの計算ができる", pena_ok, text(p, "#toast .toast-body"))
    d = sh.tab("details")
    check("交互: 取り込みのあとも梱包明細の開いているロットはそのまま",
          text(d, "#lotNo") == "A123456" and visible(d, "#outputCard"), text(d, "#lotNo"))

    # --- 統合画面をもう1枚開いてしまった ---
    sh2 = Shell(ctx, app)
    sh2.page.wait_for_timeout(3000)
    second = [k for k, sel in (("details", "#secondScreen"), ("pena", "#sgDeny"))
              if visible(sh2.frame(k), sel)]
    if visible(sh2.frame("material"), "#gate") and "別の画面" in text(sh2.frame("material"), "#gateTitle"):
        second.append("material")
    check("交互: 統合画面の2枚目では3機能とも「別の画面で開いています」", len(second) == 3, second)
    sh2.page.close(run_before_unload=True)
    sh.page.bring_to_front()
    sh.page.wait_for_timeout(2500)
    trouble = [] if sh.ready(10) else sh.trouble()
    check("交互: 2枚目を閉じても、1枚目はそのまま使える", not trouble, trouble)
    p = sh.tab("pena")
    p.click("button[data-act=calcTare]")
    check("交互: 1枚目のペナラベルで計算できる(断られない)", pena_toast(p, "計算しました"),
          text(p, "#toast .toast-body"))


# ======================================================================
# 2b. タブを何枚も開いたまま使う(送信・受信)
# ======================================================================
def _print_tab_state(pg) -> dict:
    """印刷用のタブ: メニュー・受付の覆いが出ていないか。"""
    return pg.evaluate("""() => {
      const shown = s => { const e = document.querySelector(s); return !!e && !e.hidden &&
        getComputedStyle(e).display !== 'none'; };
      return {menu: shown('.appbar-nav'), guard: shown('#screenguard'), bar: shown('.printbar')}; }""")


# ======================================================================
# 印刷する: すべての印刷プレビューに「印刷する」があり、押すと印刷が開く(統合 1.0.15)
# ======================================================================
_PRINT_STUB = "() => { window.__printed = 0; window.print = () => { window.__printed++; }; }"


def _print_preview(pg, label: str, sel: str, shots: str = "") -> None:
    """その画面に「印刷する」が1つあり、押すと印刷のダイアログ(window.print)が開き、
    紙(印刷のとき)には出ない。ダイアログの代わりに呼ばれた回数を数える。"""
    pg.wait_for_load_state()
    pg.wait_for_timeout(300)
    # 開いたとき自分で印刷を出す画面(ペナラベルの印刷ビュー)もある。それは数えない
    pg.evaluate(_PRINT_STUB)
    btn = pg.locator(sel)
    found = btn.count()
    name = (btn.first.text_content() or "").strip() if found else ""
    shown = found == 1 and btn.is_visible()
    called = 0
    if shown:
        btn.click()
        called = pg.evaluate("window.__printed")
    if shots:
        Path(shots).mkdir(parents=True, exist_ok=True)
        pg.screenshot(path=str(Path(shots) / f"print_{len(list(Path(shots).glob('print_*')))+1:02d}.png"))
    pg.emulate_media(media="print")
    on_paper = found == 1 and btn.is_visible()
    pg.emulate_media(media="screen")
    check(f"印刷する: {label}", shown and name == "印刷する" and called == 1 and not on_paper,
          {"数": found, "名前": name, "印刷が開いた回数": called, "紙に出る": on_paper,
           "画面": urllib.parse.urlparse(pg.url).path})


def _details_history_id(app: "App") -> str:
    """出力した明細の履歴の番号(梱包明細の手元のDBから)。"""
    for db in app.home.rglob("packing_details.db"):
        with sqlite3.connect(str(db)) as conn:
            row = conn.execute("SELECT 履歴ID FROM 明細出力 WHERE 履歴ID != '' "
                               "ORDER BY 出力日時 LIMIT 1").fetchone()
        if row:
            return row[0]
    return ""


def scenario_print(app: App, ctx, sh: Shell, note: dict, shots: str = "") -> None:
    """3機能のすべての印刷プレビュー(印刷用のタブ)を開いて「印刷する」を押す。
    一連の流れのあと(明細・ラベル・風袋・全サイズ・チェックリスト・発注票ができている)。"""
    print("■ 印刷する(すべての印刷プレビュー)")
    t = "t=" + app.token
    hid = _details_history_id(app)
    pages = [
        ("梱包明細 明細表(1枚)", f"/details/report/L5160Z0/1?{t}", "p.editbar button.print"),
        ("梱包明細 明細表(まとめて)", f"/details/report/L5160Z0?nos=1&{t}", "p.editbar button.print"),
        ("梱包明細 履歴から作り直した紙面", f"/details/report/history?ids={hid}&{t}",
         "p.editbar button.print"),
        ("ペナラベル 風袋計算の印刷ビュー", "/pena/tare/print", ".printbar [data-print-now]"),
        ("ペナラベル 羅列計算の印刷ビュー", "/pena/list/print", ".printbar [data-print-now]"),
        ("ペナラベル 小ラベルの印刷(実寸)", "/pena/labels/print?ob=5", ".printbar [data-print-now]"),
        ("ペナラベル ラベル台紙の印刷ビュー", "/pena/labels/sheet?ob=5", ".printbar [data-print-now]"),
        ("ペナラベル 位置合わせ(試し刷り)", "/pena/labels/calibration", ".printbar [data-print-now]"),
        ("ペナラベル 全サイズの印刷ビュー",
         "/pena/all-size/print?combo=" + urllib.parse.quote(note.get("combo", "")),
         ".printbar [data-print-now]"),
        ("資材計算 チェックリスト", f"/material/report/checklist?{t}", "#printNow"),
        ("資材計算 発注票", f"/material/report/order?{t}", "#printNow"),
    ]
    check("印刷する: 梱包明細の履歴の番号が取れる", bool(hid), hid)
    for label, path, sel in pages:
        pg = ctx.new_page()
        try:
            pg.goto(app.base + path)
            _print_preview(pg, label, sel, shots)
        finally:
            pg.close()
    # ペナラベルの小ラベルの印刷には、帯とは別に案内の中にも「印刷する」がある
    pg = ctx.new_page()
    try:
        pg.goto(app.base + "/pena/labels/print?ob=5")
        _print_preview(pg, "ペナラベル 小ラベルの印刷(案内の中のボタン)",
                       '.preview-note button[onclick="window.print()"]')
    finally:
        pg.close()
    check("印刷する: 見たあとも3機能とも使える", sh.ready(10), sh.trouble())


def scenario_multitab(app: App, ctx, sh: Shell, note: dict) -> None:
    print("■ 2b. タブを何枚も開いたまま使う(送信・受信)")
    tabs = []
    other = ctx.new_page()                          # 本ツール以外のサイト
    other.set_content("<h1>社内ポータル</h1>")
    # --- ペナラベル: 小ラベルの印刷を3回(タブが増える)・風袋計算の印刷ビュー ---
    p = sh.tab("pena")
    pena_ready(p)
    for _ in range(3):
        with ctx.expect_page() as info:
            p.click("button[data-act=printLabels]")
            _confirm_in_page(p)
        tabs.append(info.value)
    p.click('a[data-pane="tare"]')
    wait_until(lambda: p.locator('a[href$="/tare/print"]').count() > 0, 8)
    with ctx.expect_page() as info:
        p.click('a[href$="/tare/print"]')
    tabs.append(info.value)
    p.click('a[data-pane="home"]')
    # --- 梱包明細: 帳票を2回 ---
    d = sh.tab("details")
    for _ in range(2):
        with ctx.expect_page() as info:
            d.click("#btnPrint")
        tabs.append(info.value)
    # --- 資材計算: チェックリストと発注票 ---
    m = sh.tab("material")
    material_ready(m)
    for path in ("/material/report/checklist", "/material/report/order"):
        pg = ctx.new_page()
        pg.goto(app.base + path + "?t=" + app.token)
        tabs.append(pg)
    for pg in tabs:
        pg.wait_for_load_state()
    time.sleep(2)
    pena_tabs = [pg for pg in tabs if "/pena/" in pg.url]
    states = [_print_tab_state(pg) for pg in pena_tabs]
    check("複数タブ: ペナラベルの印刷のタブにメニューも受付の覆いも出ない",
          all(not st["menu"] and not st["guard"] and st["bar"] for st in states), states[:2])
    check("複数タブ: 印刷のタブを %d 枚開いても、統合画面に断り・接続なしが出ない" % len(tabs),
          not sh.trouble(), sh.trouble())

    # --- 開いたまま3機能で送る・受ける ---
    p = sh.tab("pena")
    check("複数タブ: ペナラベルで重量反映", pena_weights(p, "y333333", "30", "32"),
          text(p, "#toast .toast-body"))
    for i, n in enumerate(("11", "11", "10", "10"), 1):
        p.fill(f"#coilH{i}", n)
    p.click("button[data-act=calcTare]")
    check("複数タブ: ペナラベルで風袋計算(30kg × 11本 = 330.0)",
          pena_toast(p, "計算しました") and text(p, "#nw1") == "330.0", text(p, "#nw1"))
    note["pena"] = ("Y333333", "330.0")
    labels = pena_tabs[0]
    labels.reload()
    labels.wait_for_load_state()
    check("複数タブ: 印刷のタブを読み直すと新しい計算が出る(受信)",
          "Y333333" in labels.content() or "330" in labels.content(), labels.url.split("?")[0])
    d = sh.tab("details")
    err = details_stack_and_output(d, ["1-5", "2-5"])
    check("複数タブ: 梱包明細で出力(送信)", not err, err or text(d, "#notice"))
    m = sh.tab("material")
    err = material_calc(m, "A123456")
    check("複数タブ: 資材計算で計算(送信)", not err, err)

    # --- 同時に送る: 帳票の書き足し × 2・ペナラベルの計算・資材計算の計算 ---
    reports = [pg for pg in tabs if "/details/report/" in pg.url]
    for i, rp in enumerate(reports):
        rp.click('span.edit[data-placeholder="サイズ"]')
        rp.keyboard.press("End")
        rp.keyboard.type(f"-{i}")
    p = sh.frame("pena")
    m = sh.frame("material")
    p.evaluate("() => document.querySelector('button[data-act=calcTare]').click()")
    m.evaluate("() => document.querySelector('#run').click()")
    for rp in reports:
        rp.keyboard.press("Enter")
    saved = [wait_until(lambda rp=rp: "保存しました" in text(rp, "#editSaved")
                        or "開き直して" in text(rp, "#editSaved"), 8) for rp in reports]
    check("複数タブ: 同時に送っても帳票の書き足しが届く", all(saved), [text(rp, "#editSaved") for rp in reports])
    check("複数タブ: 同時に送ってもペナラベル・資材計算の計算が返る",
          pena_toast(p, "計算しました") and material_toast(m, "台数"),
          (text(p, "#toast .toast-body"), text(m, "#toast")))

    # --- 統合画面を読み直す(印刷のタブは開いたまま) ---
    sh.page.reload()
    wait_until(lambda: all(sh.frame(k) for k in ("details", "pena", "material")), 20)
    sh.page.wait_for_timeout(3000)
    trouble = [] if sh.ready(15) else sh.trouble()
    check("複数タブ: 印刷のタブを開いたまま統合画面を読み直しても断られない", not trouble, trouble)
    p = sh.tab("pena")
    p.click("button[data-act=calcTare]")
    check("複数タブ: 読み直したあとも計算できる", pena_toast(p, "計算しました"), text(p, "#toast .toast-body"))

    # --- 印刷のタブを順ばらばらに閉じる ---
    for pg in (tabs[3], tabs[0], tabs[6], tabs[1], tabs[5], tabs[2], tabs[7], tabs[4]):
        if not pg.is_closed():
            pg.close(run_before_unload=True)
        time.sleep(0.4)
    other.close()
    time.sleep(2)
    trouble = [] if sh.ready(10) else sh.trouble()
    check("複数タブ: 印刷のタブを全部閉じても統合画面はそのまま", not trouble and app.alive(), trouble)
    logs = app.logs()
    opened = logs.count("別タブ tab-")
    check("複数タブ: 印刷のタブの心拍がサーバに届いていた", opened >= len(tabs), f"{opened} 回")


# ======================================================================
# 10. 続けて別のツールで印刷する・すぐ切り替えて計算する(`--only rapid`。統合 1.2.20)
# ======================================================================
#: 確認ダイアログ(window.confirm)を押すまでの秒数。人が読んでから押すのを真似る
DIALOG_WAIT = {"sec": 0.0}

_COUNT_PRINT = """() => { if (window.__printCounted) return; window.__printCounted = true;
  if (typeof window.__printed !== 'number') window.__printed = 0;
  window.print = function () { window.__printed++; }; }"""


def _on_dialog(dialog) -> None:
    if DIALOG_WAIT["sec"]:
        time.sleep(DIALOG_WAIT["sec"])
    dialog.accept()


def _confirm_in_page(fr, wait: float = 0.0) -> bool:
    """画面の中の確かめ(ペナラベルの「印刷します」)が出ていれば、wait 秒おいて OK を押す。"""
    sel = "dialog.pplconfirm[open] button[data-ok]"
    if not wait_until(lambda: visible(fr, sel), 10, 0.05):
        return False
    if wait:
        fr.wait_for_timeout(int(wait * 1000))
    fr.click(sel)
    return True


def _pump_until(page, fn, timeout=15.0) -> bool:
    """Playwright の知らせ(新しいタブなど)を受け取りながら待つ。`time.sleep` の間は届かない。"""
    def ok():
        try:
            return fn()
        except Exception:                       # noqa: BLE001 - 面を移っている最中は読めない
            return False

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if ok():
            return True
        page.wait_for_timeout(100)
    return bool(ok())


def _kind(url: str) -> str:
    path = urllib.parse.urlparse(url).path
    for key, part in (("details", "/details/report/"), ("pena", "/pena/labels/print"),
                      ("material", "/material/report/checklist")):
        if part in path:
            return key
    return ""


def _delay(ctx, pattern: str, seconds: float):
    """その要求だけ、サーバへ届くのを遅らせる(共有フォルダが遅い・PC が重いときと同じ形)。"""
    def handler(route):
        time.sleep(seconds)
        route.continue_()
    ctx.route(pattern, handler)
    return lambda: ctx.unroute(pattern, handler)


def scenario_rapid(app: App, ctx, sh: Shell, note: dict) -> None:
    print("■ 10. 続けて別のツールで印刷する・すぐ切り替えて計算する")
    opened: list = []
    ctx.on("page", lambda pg: opened.append(pg))
    page = sh.page

    # ---- 用意: 3機能とも刷れる状態 ----
    d = sh.tab("details")
    lot = text(d, "#lotNo")
    p = sh.tab("pena")
    pena_ready(p)
    ok = pena_weights(p, "r333333", "15", "17")
    for i, n in enumerate(("11", "11", "10", "10"), 1):
        p.fill(f"#coilH{i}", n)
    p.click("button[data-act=calcTare]")
    # 計算しないと印刷ビューの「印刷する」は止まる(入力が変わったまま刷らない。移植元どおり)
    check("連続: 用意(ペナラベルの重量反映 R333333 15kg → 計算 165.0)",
          ok and pena_toast(p, "計算しました") and text(p, "#nw1") == "165.0", text(p, "#nw1"))
    m = sh.tab("material")
    material_ready(m)
    # チェックリストは一連の流れで積んだ行(A123456)を使う。ここで積み足すと、あとの試験
    # (交互に使う・放置)が数える行数が変わる。面を移り終えるまで待つ(押した直後は前の面のまま)
    m.click('nav.rail a.rail__item[href*="/checklist"]')
    _pump_until(page, lambda: "/checklist" in (sh.frame("material").url or "")
                and sh.frame("material").evaluate("document.readyState") == "complete", 15)
    m = sh.frame("material")
    material_ready(m)
    _pump_until(page, lambda: m.evaluate("document.querySelectorAll('#rows tr.on').length") >= 1, 10)
    rows = m.evaluate("document.querySelectorAll('#rows tr.on').length")
    check("連続: 用意(資材計算のチェックリストに行がある)", rows >= 1 and "A123456" in text(m, "#rows"), rows)
    d = sh.frame("details")
    check("連続: 用意(梱包明細の印刷が押せる)", lot and not d.is_disabled("#btnPrint"), lot)

    # ---- A. 3機能の印刷を待たずに続けて押す(3回) ----
    for rnd in range(1, 4):
        before = len(opened)
        t0 = time.monotonic()
        page.click("#tab-details")
        d.click("#btnPrint", no_wait_after=True)
        page.click("#tab-pena")
        p.click("button[data-act=printLabels]", no_wait_after=True)
        _confirm_in_page(p)
        page.click("#tab-material")
        m.click("a#print", no_wait_after=True)
        took = time.monotonic() - t0
        _pump_until(page, lambda: len(opened) - before >= 3, 15)
        page.wait_for_timeout(500)
        new = opened[before:]
        for pg in new:
            pg.wait_for_load_state()
        kinds = sorted(_kind(pg.url) for pg in new)
        check(f"連続 {rnd}回目: 3機能の印刷を {took:.1f}秒 で続けて押して、3枚とも開く",
              kinds == ["details", "material", "pena"], kinds)
        got = {}
        for pg in new:
            k = _kind(pg.url)
            body = pg.evaluate("document.body.innerText")
            if k == "details":
                got[k] = lot in pg.title()
            elif k == "pena":
                got[k] = "R333333" in body and pg.locator("section.sheet").count() >= 1
                # 小ラベルの印刷は開いただけでは刷らない(移植元どおり。風袋・台紙・全サイズは刷る)
                got["pena 開いたとき刷った回数"] = pg.evaluate("window.__printed")
            elif k == "material":
                got[k] = "A123456" in body
        check(f"連続 {rnd}回目: どの紙も自分のツールの中身(混ざらない)",
              got.get("details") and got.get("pena") and got.get("material")
              and got.get("pena 開いたとき刷った回数") == 0, got)
        # 開いた3枚で、続けて「印刷する」
        presses = {}
        for pg in new:
            k = _kind(pg.url)
            if not k:
                continue
            pg.evaluate(_COUNT_PRINT)
            start = pg.evaluate("window.__printed")
            sel = {"details": "p.editbar button.print", "pena": ".printbar [data-print-now]",
                   "material": "#printNow"}[k]
            try:
                pg.bring_to_front()
                pg.click(sel, timeout=8000)
                presses[k] = pg.evaluate("window.__printed") - start
            except Exception as exc:              # noqa: BLE001 - 押せなかったことを残す
                presses[k] = "押せない"
                print("    押せない(%s): %s" % (k, " | ".join(str(exc).splitlines()[:14])[:1500]))
                print("    その画面:", pg.url, pg.evaluate("document.body.innerHTML.slice(0, 700)"))
        check(f"連続 {rnd}回目: 3枚で続けて「印刷する」を押すと、それぞれ1回ずつ刷る",
              presses == {"details": 1, "pena": 1, "material": 1}, presses)
        for pg in new:
            pg.close()
    page.bring_to_front()
    check("連続: 印刷を続けたあとも3機能とも使える", sh.ready(10), sh.trouble())

    # ---- A2. 確かめを読んでから押す(6秒)。ブラウザのポップアップのブロックに止められない ----
    p = sh.tab("pena")
    before = len(opened)
    DIALOG_WAIT["sec"] = 6.0
    try:
        p.click("button[data-act=printLabels]", no_wait_after=True)
        _confirm_in_page(p, wait=6.0)
        ok = _pump_until(page, lambda: any(_kind(x.url) == "pena" for x in opened[before:]), 15)
    finally:
        DIALOG_WAIT["sec"] = 0.0
    check("連続: ペナラベルの印刷で、確かめを6秒読んでから押しても印刷ビューが開く", ok,
          [x.url for x in opened[before:]])
    for pg in opened[before:]:
        pg.close()

    # ---- B. 計算を押してすぐ別のツールへ ----
    m = sh.tab("material")
    m.click('nav.rail a.rail__item[href*="/calc"]')
    m.wait_for_load_state()
    m = sh.frame("material")
    material_ready(m)
    m.fill("#lot", "")
    m.type("#lot", "A123456", delay=20)
    wait_until(lambda: text(m, "#o-包装") == SPEC_NO, 10)
    material_ins(m, "10")
    p = sh.frame("pena")
    for _ in range(5):
        page.click("#tab-material")
        m.click("#run")
        page.click("#tab-pena")
        p.click("button[data-act=calcTare]")
        page.click("#tab-details")
    page.wait_for_timeout(2500)
    m, p = sh.frame("material"), sh.frame("pena")
    check("すぐ切替: 資材計算で計算 → すぐペナラベルで計算 → すぐ梱包明細(5回)。資材計算の結果",
          (text(m, "#r-種類"), text(m, "#r-台数")) == ("スカシ", "2"), (text(m, "#r-種類"), text(m, "#r-台数")))
    check("すぐ切替: 同じく、ペナラベルの結果(15kg × 11本 = 165.0・検査番号 R333333)",
          text(p, "#nw1") == "165.0" and text(p, "#lblKensaNo") == "R333333",
          (text(p, "#nw1"), text(p, "#lblKensaNo")))
    bad = [t for t in (text(p, "#toast .toast-body"), text(m, "#toast")) if re.search("失敗|エラー|できません", t)]
    check("すぐ切替: どちらにも失敗の知らせが出ていない", not bad, bad)
    check("すぐ切替: 断り・接続なしが出ていない", sh.ready(10), sh.trouble())

    # ---- B2. 2機能から同時に計算を投げる(10回ずつ) ----
    p.evaluate("""() => { window.__burst = Promise.all(Array.from({length: 10}, () =>
        pplPost('/api/calc-tare', {current: {}}).then(j => (j && j.ok ? 'ok ' : 'NG ')
          + JSON.stringify(j && j.state).length + ' ' + (j && j.message)))); }""")
    m.evaluate("""() => { window.__burst = Promise.all(Array.from({length: 10}, () =>
        fetch(APP_BASE + '/api/calc/run', {method: 'POST', headers: {'Content-Type': 'application/json',
          'X-App-Token': APP_TOKEN, 'X-App-Screen': sessionStorage.getItem('coil.screen') || ''},
          body: '{}'}).then(r => r.json()).then(j => j.ok ? j.view.result.台数 : (j.message || '?')))); }""")
    pena_got = p.evaluate("window.__burst")
    mat_got = m.evaluate("window.__burst")
    check("同時: ペナラベルの計算10回がどれも同じ答え", len(set(pena_got)) == 1, sorted(set(pena_got))[:3])
    check("同時: 資材計算の計算10回がどれも台数2", set(map(str, mat_got)) == {"2"},
          sorted(set(map(str, mat_got)))[:3])

    # ---- C. 入れてすぐ計算(欄を離れたときの送信と、計算が行き違わない) ----
    # 検入数で台数が変わる2つの値を、ゆっくり(1つずつ待って)測っておく
    m = sh.tab("material")
    want = {}
    for ins in ("10", "400"):
        material_ins(m, ins)
        m.click("#run")
        material_toast(m, "台数")
        page.wait_for_timeout(300)
        want[ins] = text(m, "#r-台数")
    check("入れてすぐ計算: 検入数で台数が変わる2つの値(ゆっくり測る)", want["10"] != want["400"], want)
    # 欄を離れたときの送信が少し遅れても(PC が重いときと同じ形)、計算は入れた値で行う
    undo = _delay(ctx, "**/api/calc/input", 0.8)
    try:
        for ins in ("10", "400", "10"):
            m.fill("#ins", ins)
            m.click("#run")              # Tab を押さずに、すぐ「計算」
            page.wait_for_timeout(2000)
            check(f"入れてすぐ計算: 検入数 {ins} を入れてすぐ「計算」→ 台数 {want[ins]}",
                  text(m, "#r-台数") == want[ins], {"台数": text(m, "#r-台数"), "検入数": text(m, "#ins")})
    finally:
        undo()
    # 梱包明細: 重量を入れてすぐ別のツールへ → 戻る(送ったはずの重量が残っている)
    d = sh.tab("details")
    undo = _delay(ctx, "**/api/weights", 0.8)
    try:
        w1 = text(d, "#weight1")
        new_w = "251" if w1 != "251" else "252"
        d.fill("#weight1", new_w)
        page.click("#tab-pena")          # 欄を離れる(送信) → すぐ切り替える
        page.wait_for_timeout(200)
        page.click("#tab-details")
        page.wait_for_timeout(2000)
        saved = d.evaluate("(window.APP.state.weights || [])[0]")
        check("すぐ切替: 梱包明細で重量を入れてすぐ切り替えても、入れた重量が残る(画面とサーバ)",
              text(d, "#weight1") == new_w and str(saved).rstrip("0").rstrip(".") == new_w,
              {"画面": text(d, "#weight1"), "サーバ": saved})
    finally:
        undo()
    check("入れてすぐ計算: 断り・接続なしが出ていない", sh.ready(10), sh.trouble())

    # ---- D. 印刷の画面が開いているあいだ(ブラウザ版): 心拍が止まっても終わらない ----
    # 印刷の画面のあいだは画面の JavaScript が止まる。その前に「裏に回った」と同じ合図が出るか
    p = sh.tab("pena")
    said0, back0 = app.logs().count("心拍: print (hidden)"), app.logs().count("心拍: afterprint (visible)")
    p.evaluate("() => { window.dispatchEvent(new Event('beforeprint')); }")
    said = _pump_until(page, lambda: app.logs().count("心拍: print (hidden)") > said0, 5)
    p.evaluate("() => { window.dispatchEvent(new Event('afterprint')); }")
    back = _pump_until(page, lambda: app.logs().count("心拍: afterprint (visible)") > back0, 5)
    check("印刷中: 機能の画面で刷る前に「裏」の合図、済んだら「前」の合図(統合画面から)",
          said and back, {"刷る前": said, "済んだら": back})
    tab = ctx.new_page()
    tab.goto(app.base + "/pena/tare/print")
    tab.wait_for_load_state()
    said0, back0 = app.logs().count("心拍: print (hidden)"), app.logs().count("心拍: afterprint (visible)")
    tab.evaluate("() => { window.dispatchEvent(new Event('beforeprint')); }")
    said = _pump_until(page, lambda: app.logs().count("心拍: print (hidden)") > said0, 5)
    tab.evaluate("() => { window.dispatchEvent(new Event('afterprint')); }")
    back = _pump_until(page, lambda: app.logs().count("心拍: afterprint (visible)") > back0, 5)
    check("印刷中: 印刷用のタブでも、刷る前に「裏」・済んだら「前」の合図", said and back,
          {"刷る前": said, "済んだら": back})
    tab.close()

    # ---- E. デスクトップ版: 印刷の画面が出なかったら知らせる(出たときは知らせない) ----
    pg = ctx.new_page()
    pg.goto(app.base + "/api/health")
    pg.set_content("<!doctype html><html><head></head><body><p>印刷ビュー</p></body></html>")
    pg.evaluate("""() => { window.__TAURI__ = {core: {invoke: () => Promise.resolve()}};
        window.__mode = 'busy';
        window.print = function () {
          if (window.__mode === 'shown') { window.dispatchEvent(new Event('beforeprint')); }
          if (window.__mode === 'slow') { const t = Date.now(); while (Date.now() - t < 700) {} }
        }; }""")
    pg.add_script_tag(url=app.base + "/static/js/desktop.js")
    seen = {}
    for mode in ("busy", "shown", "slow"):
        pg.evaluate("""m => { const b = document.getElementById('cptPrintBusy'); if (b) b.remove();
            window.__mode = m; window.print(); }""", mode)
        seen[mode] = pg.evaluate("!!document.getElementById('cptPrintBusy')")
    pg.emulate_media(media="print")
    on_paper = pg.evaluate("""() => { window.__mode = 'busy'; window.print();
        const b = document.getElementById('cptPrintBusy'); return !!b && getComputedStyle(b).display !== 'none'; }""")
    pg.emulate_media(media="screen")
    check("デスクトップ版: 印刷の画面が出ずにすぐ戻ったら知らせる(出たとき・待ったときは出さない・紙に出ない)",
          seen == {"busy": True, "shown": False, "slow": False} and not on_paper, {**seen, "紙に出る": on_paper})
    pg.close()


# ======================================================================
# 保存先(このPC／複数のPCで共有) ── 3機能の設定部に明記してあるか(統合 1.0.10)
# ======================================================================
def _store_rows(fr) -> dict:
    """描かれた保存先の表。{種類: [名前・場所…の文字]}"""
    return fr.evaluate("""() => Object.fromEntries([...document.querySelectorAll('[data-store]')]
        .map(b => [b.dataset.store, [...b.querySelectorAll('tbody tr')].map(r => r.innerText)]))""")


def scenario_storage(app: App, ctx, sh: Shell, note: dict, shots: str = "") -> None:
    """3機能の設定を開き、**このPCに保存**と**全ラインで共有**が分けて出ていること。

    現場の指摘: このPCで引き継いで使うものと、複数のPCで共有するものは違う。設定部に
    そういうファイルがあることを明記してほしい。画面の JavaScript が描くところまで見る。
    """
    def shot(page_or_frame, name):
        if shots:
            Path(shots).mkdir(parents=True, exist_ok=True)
            sh.page.screenshot(path=str(Path(shots) / f"{name}.png"), full_page=False)

    # --- 梱包明細: 設定 → 保存先(このPC／共有)
    d = sh.tab("details")
    d.click("#btnSettings")
    d.click("#tabStore")
    ok = wait_until(lambda: visible(d, "#panelStore")
                    and d.locator("#storeGroups [data-store]").count() == 3, 10)
    rows = _store_rows(d) if ok else {}
    local = "\n".join(rows.get("local", []))
    shared = "\n".join(rows.get("shared", []))
    check("保存先: 梱包明細 このPCに保存(設定ファイル・手元のDB)",
          "user_config.json" in local and "packing_details.db" in local, local[:200])
    check("保存先: 梱包明細 全ラインで共有(右上の文字・管理者パスワード・明細の履歴)",
          all(n in shared for n in ("梱包資材マスタ.sqlite3", "梱包明細打ち出し.json",
                                    "梱包明細履歴.sqlite3")), shared[:200])
    shot(d, "details_storage")
    d.click("#tabPlace")
    badges = d.locator("#panelPlace .where-local").count()
    check("保存先: 梱包明細 置き場所の欄ごとに「このPCに保存」の札", badges == 12, badges)  # 仕掛台帳(2つ目)で 7、起動時の自動取り込み(1.2.22)で 8、記録の保存先・ライン名・テストモード(1.2.42)で 11、記録ファイルの形式(1.2.44)で 12
    shot(d, "details_place")
    d.click("#btnSettingsClose")

    # --- 資材計算: 設定 → 保存先(このPC／共有)
    m = sh.tab("material")
    m.click('.rail__item[href*="/settings"]')
    material_ready(m)
    m.click("#tab-storage")
    # 面が**開いて見えている**ところまで(描いてあっても隠れていれば読めない)
    ok = wait_until(lambda: visible(m, "#panel-storage") and not visible(m, "#panel-source")
                    and m.locator("#storeGroups [data-store]").count() == 3, 10)
    rows = _store_rows(m) if ok else {}
    local = "\n".join(rows.get("local", []))
    check("保存先: 資材計算 このPCに保存(設定ファイル・作業用DB・このPCだけのパスワード)",
          "user_config.json" in local and "coil_tool.db" in local and "このPCだけ" in local,
          local[:200])
    check("保存先: 資材計算 全ラインで共有(マスタ管理の書き先)",
          any("梱包資材マスタ.sqlite3" in r and "マスタ管理" in r for r in rows.get("shared", [])),
          rows.get("shared"))
    shot(m, "material_storage")
    m.click("#tab-master")
    check("保存先: 資材計算 パスワードは「このPCだけ」と書いてある",
          "このPCだけのパスワードです" in text(m, "#panel-master"))
    m.locator(".rail__item").first.click()            # 最初の画面へ戻す
    material_ready(m)

    # --- ペナラベル: 設定(パス設定の面の「設定の保存先」)
    p = sh.tab("pena")
    p.click('.appbar-nav a[href$="/settings"]')
    ok = wait_until(lambda: visible(p, "#storage"), 10)
    rows = _store_rows(p) if ok else {}
    local = "\n".join(rows.get("local", []))
    check("保存先: ペナラベル このPCに保存(local.json・状態DB)",
          "local.json" in local and "state.sqlite3" in local, local[:200])
    p.locator("#storage").scroll_into_view_if_needed()
    shot(p, "pena_storage")
    p.click('.appbar-nav a[data-pane="home"]')
    pena_ready(p)
    check("保存先: 見たあとも3機能とも使える", sh.ready(10), sh.trouble())


# ======================================================================
# マスタ管理: 見出しを押して並べ替える(3機能。統合 1.0.11)
# ======================================================================
_GRID = """([head, body]) => {
  const ths = [...document.querySelectorAll(head + ' th')];
  const names = ths.map(th => (th.querySelector('button.sortbtn')?.firstChild?.textContent
                               ?? th.textContent).trim());
  const sorted = ths.map(th => th.getAttribute('aria-sort') || '');
  const rows = [...document.querySelectorAll(body + ' tr')]
      .map(tr => [...tr.children].map(td => td.textContent.trim()));
  return {names, sorted, rows};
}"""


def _sort_check(fr, label: str, head: str, body: str) -> None:
    """見出しを押して昇順 → もう一度押して降順。並びと ▲▼(aria-sort)を確かめる。

    **文字の列**と**数字の列**を1つずつ選んで押す。数字の列は数の大きさで並ぶこと
    (取り込み元は数字も文字で持つので、文字の並びだと 800 が 1350 の後ろへ来ていた)。
    """
    ok = wait_until(lambda: fr.locator(head + " button.sortbtn").count() > 0
                    and fr.locator(body + " tr").count() > 1, 15)
    check(f"マスタ: {label} 見出しが押せる(並べ替えのボタン)", ok)
    if not ok:
        return
    grid = fr.evaluate(_GRID, [head, body])

    def num(v):
        return v.replace(".", "", 1).lstrip("-").isdigit()

    def column(col):
        return [r[col] for r in grid["rows"] if col < len(r)]

    def is_text(col):
        values = column(col)
        return len(set(values)) > 1 and all(values) and not all(num(v) for v in values)

    def is_number(col):
        values = column(col)
        # 桁の違う数字が混ざる列(文字の並びと数の並びが食い違う)を選ぶ
        return (len(set(values)) > 1 and all(values) and all(num(v) for v in values)
                and len({len(v.split(".")[0].lstrip("-")) for v in values}) > 1)

    picks = [("文字", next((i for i in range(len(grid["names"])) if is_text(i)), None),
              lambda vs, rev: sorted(vs, reverse=rev)),
             ("数字", next((i for i in range(len(grid["names"])) if is_number(i)), None),
              lambda vs, rev: sorted(vs, key=float, reverse=rev))]
    for kind, col, expect in picks:
        check(f"マスタ: {label} {kind}の列がある", col is not None, grid["names"])
        if col is None:
            continue
        name = grid["names"][col]
        for want, rev in (("ascending", False), ("descending", True)):
            fr.locator(head + " th").nth(col).locator("button.sortbtn").click()
            ok = wait_until(
                lambda: fr.evaluate(_GRID, [head, body])["sorted"][col] == want, 10)
            grid = fr.evaluate(_GRID, [head, body])
            values = column(col)
            check(f"マスタ: {label} 「{name}」({kind})で{'降順' if rev else '昇順'}に並ぶ",
                  ok and values == expect(values, rev), values[:8])


def scenario_master(app: App, ctx, sh: Shell, note: dict) -> None:
    """3機能のマスタ管理で、見出し(列名)を押すと並べ替わること(現場の指摘)。"""
    # --- 梱包明細: マスタ管理(見るだけ)
    d = sh.tab("details")
    d.click("#btnMaster")
    _sort_check(d, "梱包明細", "#mHead", "#mRows")
    d.click("#btnMasterClose")

    # --- 資材計算: 設定 → マスタ管理
    m = sh.tab("material")
    m.click('.rail__item[href*="/settings"]')
    material_ready(m)
    m.click("#tab-master")
    m.locator("#mTables button", has_text="パレット").first.click()
    _sort_check(m, "資材計算", "#mHead", "#mRows")
    m.locator(".rail__item").first.click()
    material_ready(m)

    # --- ペナラベル: 設定 → マスタ管理(SQLite の資材マスタを指したときだけ見られる)
    p = sh.tab("pena")
    master_db = str(app.home / "master" / "梱包資材マスタ.sqlite3")
    res = _api(p, "/pena/api/settings/save",
               {"values": {"material_db_file": master_db}, "password": PENA_PASSWORD})
    check("マスタ: ペナラベル 資材マスタ(SQLite)を指す", res["status"] == 200, res["body"])
    p.click('.appbar-nav a[href$="/settings"]')
    wait_until(lambda: visible(p, '.subtab[data-sub="master"]'), 10)
    p.click('.subtab[data-sub="master"]')
    _sort_check(p, "ペナラベル", "#mTbl thead", "#mTbl tbody")
    # 元へ戻す(ほかの段はペナラベルを CSV で動かしている)
    res = _api(p, "/pena/api/settings/save",
               {"values": {"material_db_file": ""}, "password": PENA_PASSWORD})
    check("マスタ: ペナラベル 資材マスタの設定を元へ戻す", res["status"] == 200, res["body"])
    p.click('.appbar-nav a[data-pane="home"]')
    pena_ready(p)
    check("マスタ: 見たあとも3機能とも使える", sh.ready(10), sh.trouble())


# ======================================================================
# エラーの後追い: エラーの記録・ログの出力先(統合 1.0.12)
# ======================================================================
#: わざと起こすエラーの印(最後の「エラーが無い」の確かめから外す)
TRAIL_MARK = "通し試験のわざとのエラー"


def scenario_trail(app: App, ctx, sh: Shell, note: dict) -> None:
    """画面でエラーが起きたら番号が出て、上の帯の「ログ」で中身(なぜなぜの欄つき)が見られる。
    ログの出力先を画面で変えると、すぐその場所(PC の名前のフォルダ)に書かれる。"""
    # --- 資材計算の画面で、わざと JavaScript のエラーを起こす
    m = sh.tab("material")
    m.evaluate("(mark) => setTimeout(() => { throw new Error(mark); }, 0)", TRAIL_MARK)
    ok = wait_until(lambda: visible(m, "#cpt-error-notice"), 10)
    shown = text(m, "#cpt-error-notice") if ok else ""
    found = re.search(r"E\d{8}-\d{6}-[0-9A-F]{4}", shown)
    check("後追い: 画面のエラーで番号が出る", bool(found), shown[:120])
    if not found:
        return
    eid = found.group(0)
    # 「中身を見る」→ 統合画面の「ログ」が開いて、その記録が出る
    m.locator("#cpt-error-notice button", has_text="中身を見る").click()
    t0 = time.monotonic()
    ok = wait_until(lambda: sh.page.locator("#logDialog[open]").count() == 1, 20)
    t_dialog = time.monotonic() - t0
    lf = None
    if ok:
        # ログ画面の枠は統合画面の #logDialog の中だけを探す(ほかの /log? の画面と取り違えない)
        def log_frame():
            el = sh.page.query_selector("#logDialog iframe")
            fr = el.content_frame() if el else None
            return fr if fr is not None and "/log?" in (fr.url or "") else None
        wait_until(lambda: log_frame() is not None, 20)
        lf = log_frame()
    shown_text = ""
    if lf is not None:
        wait_until(lambda: TRAIL_MARK in (lf.locator("#incText").text_content() or ""), 20)
        shown_text = lf.locator("#incText").text_content() or ""
    opened = (eid in shown_text and TRAIL_MARK in shown_text and "なぜ1" in shown_text
              and "資材計算" in shown_text)
    # 落ちたときは様子を残す(1 回だけ落ちて理由が分からなかったため)
    diag = "" if opened else (f"dialog={ok}({t_dialog:.1f}秒) frame={'あり' if lf else 'なし'} "
                              f"frames={[f.url.split('?')[0] for f in sh.page.frames]} "
                              f"text={shown_text[:80]!r}")
    check("後追い: 「中身を見る」で記録が開く(なぜなぜの欄つき)", opened, shown_text[:200] if opened else diag)
    check("後追い: 記録がファイルに残る",
          bool(list((app.local / "logs" / "coilpacking" / "incidents").glob(f"coilpacking_なぜなぜ_*_{eid}.md"))))
    if lf is None:
        # ログ画面が開いたまま残ると、あとの段の操作がすべて止まる。閉じてから次へ
        if sh.page.locator("#logDialog[open]").count():
            sh.page.click("#logClose")
        return
    # --- 出力先を変える(共有フォルダのつもり)→ <出力パス>\coilpacking\ に書かれる(統合 1.2.53。日報複合ツールの決まり)
    lf.click('.tab[data-tab="settings"]')
    wait_until(lambda: (lf.locator("#nowDir").text_content() or "") != "", 10)
    share = app.home / "shared-logs"
    lf.fill("#dirInput", str(share))
    lf.click("#btnCheck")
    wait_until(lambda: "書けます" in (lf.locator("#result").text_content() or ""), 10)
    lf.click("#btnSave")
    ok = wait_until(lambda: "保存しました" in (lf.locator("#result").text_content() or ""), 10)
    now = lf.locator("#nowDir").text_content() or ""
    check("後追い: 出力先を保存するとすぐ切り替わる(<出力パス>\\coilpacking)",
          ok and now == str(share / "coilpacking"), now)
    # 「ログ」の画面が開いているあいだは、ほかの画面は押せない(ダイアログ)。
    # 切り替えたこと自体と、「ログ」の画面の問い合わせが新しい出力先に書かれる
    lf.click('.tab[data-tab="today"]')
    wait_until(lambda: list(share.glob("coilpacking/coilpacking_*_*.log")), 10)
    files = list(share.glob("coilpacking/coilpacking_*_*.log"))
    check("後追い: 新しい出力先にログが書かれる(名前はツールIDと PC 名で始まる・直下に置かない)", bool(files)
          and not list(share.glob("*.log")) and any(
        "ログの出力先" in f.read_text(encoding="utf-8") for f in files), [str(f) for f in files])
    # --- 配布設定(共通)に書き出す → 消す(統合 1.0.14。管理者パスワードは梱包明細と同じ)
    lf.click('.tab[data-tab="settings"]')
    wait_until(lambda: (lf.locator("#distNow").text_content() or "") != "読み込み中…", 10)
    def dist_said(word):
        return wait_until(lambda: word in (lf.locator("#distResult").text_content() or ""), 10)
    lf.fill("#distPassword", "ちがう")
    lf.click("#btnDistExport")
    check("後追い: 配布設定に書き出すには管理者パスワードが要る", dist_said("違います"),
          lf.locator("#distResult").text_content())
    admin = os.environ.get("PACKING_DETAILS_ADMIN_PASSWORD", DETAILS_PASSWORD)
    lf.fill("#distPassword", admin)
    lf.click("#btnDistExport")
    dist_file = App.dist_common(app.home) / "設定.json"
    ok = dist_said("書き出しました")
    written = (json.loads(dist_file.read_text(encoding="utf-8")).get("settings", {})
               if dist_file.exists() else {})
    check("後追い: ログの出力先を配布設定に書き出す",
          ok and written.get("log_dir") == str(share), written)
    check("後追い: 置いてある配布設定が画面に出る",
          str(share) in (lf.locator("#distNow").text_content() or ""),
          lf.locator("#distNow").text_content())
    # 「消します」の確かめは、どのページの窓も「OK」で閉じる決まり(main の ctx.on("page"))
    lf.fill("#distPassword", admin)
    lf.click("#btnDistRemove")
    check("後追い: 配布設定を消す", dist_said("消しました") and not dist_file.exists(),
          lf.locator("#distResult").text_content())
    # 元へ戻す(ほかの段はこのPCの既定の場所を見る)
    lf.click("#btnReset")
    ok = wait_until(lambda: "保存しました" in (lf.locator("#result").text_content() or "")
                    and (lf.locator("#nowDir").text_content() or "").startswith(
                        str(app.local / "logs" / "coilpacking")), 10)
    check("後追い: 既定に戻す", ok, lf.locator("#nowDir").text_content())
    sh.page.click("#logClose")
    check("後追い: 見たあとも3機能とも使える", sh.ready(10), sh.trouble())


# ======================================================================
# 配色: ブラウザの「全体的な外観」がダークのとき、3機能の背景がそろう(統合 1.0.13)
# ======================================================================
_BG = """() => {
  const c = getComputedStyle(document.body).backgroundColor.match(/\\d+(\\.\\d+)?/g).map(Number);
  return (0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]) / 255;
}"""


def scenario_theme(app: App, browser, note: dict) -> None:
    """ダークの外観で、3機能の画面の背景が暗い(ペナラベルだけ明るい、が無い)。
    印刷用のタブ(紙)とバーコードは白地のまま。"""
    ctx = browser.new_context(viewport={"width": 1400, "height": 900}, color_scheme="dark")
    try:
        sh = Shell(ctx, app)
        sh.ready(20)
        lum = {}
        for key in ("details", "pena", "material"):
            fr = sh.tab(key)
            lum[key] = round(fr.evaluate(_BG), 2)
        check("配色: ダークの外観で3機能とも背景が暗い", all(v < 0.2 for v in lum.values()), lum)
        pg = ctx.new_page()
        for path in ("/pena/labels/print?ob=5", "/pena/tare/print"):
            pg.goto(app.base + path)
            pg.wait_for_timeout(500)
            check(f"配色: 印刷用のタブは白地のまま({path.split('?')[0]})",
                  pg.evaluate(_BG) > 0.95, round(pg.evaluate(_BG), 2))
        pg.goto(app.base + "/pena/labels/print?ob=5")
        fills = pg.evaluate("""() => [...document.querySelectorAll('svg rect')].slice(0, 3)
            .map(r => r.getAttribute('fill'))""")
        check("配色: バーコードは白地に黒のまま", not fills or (fills[0] == "#fff"
              and "#000" in fills), fills)
        pg.close()
    finally:
        ctx.close()
    _theme_switch(app, browser)


def _theme_lums(sh) -> dict:
    lum = {k: round(sh.tab(k).evaluate(_BG), 2) for k in ("details", "pena", "material")}
    lum["統合画面"] = round(sh.page.evaluate(_BG), 2)
    return lum


def _theme_switch(app: App, browser) -> None:
    """上の帯の「画面の色」(統合 1.0.16)。ブラウザの外観と逆を選んでも、3機能と統合画面が
    そろって変わる・読み直しても残る・紙は白地のまま・自動に戻すとブラウザに合わせる。"""
    for scheme, pick, dark in (("light", "dark", True), ("dark", "light", False)):
        ctx = browser.new_context(viewport={"width": 1400, "height": 900}, color_scheme=scheme)
        try:
            sh = Shell(ctx, app)
            sh.ready(20)
            sh.page.click(f'[data-theme-set="{pick}"]')
            sh.page.wait_for_timeout(500)
            want = (lambda v: v < 0.2) if dark else (lambda v: v > 0.8)
            lum = _theme_lums(sh)
            name = {"dark": "ダーク", "light": "ライト"}[pick]
            check(f"画面の色: ブラウザが{scheme}でも「{name}」を押すと3機能と統合画面がそろう",
                  all(want(v) for v in lum.values()), lum)
            sh.page.reload()
            sh.ready(20)
            lum = _theme_lums(sh)
            pressed = sh.page.get_attribute(f'[data-theme-set="{pick}"]', "aria-pressed")
            check(f"画面の色: 読み直しても「{name}」のまま(このPCに保存)",
                  all(want(v) for v in lum.values()) and pressed == "true", (lum, pressed))
            if dark:
                pg = ctx.new_page()
                pg.goto(app.base + "/pena/labels/print?ob=5")
                pg.wait_for_timeout(500)
                check("画面の色: ダークを選んでもラベルの印刷用のタブは白地のまま",
                      pg.evaluate(_BG) > 0.95, round(pg.evaluate(_BG), 2))
                pg.close()
            sh.page.click('[data-theme-set="auto"]')
            sh.page.wait_for_timeout(500)
            lum = _theme_lums(sh)
            back = (lambda v: v > 0.8) if scheme == "light" else (lambda v: v < 0.2)
            check(f"画面の色: 「自動」でブラウザ({scheme})に合わせる", all(back(v) for v in lum.values()), lum)
        finally:
            ctx.close()


# ======================================================================
# 3. 放置する
# ======================================================================
_HIDE = """() => {
  for (const [k, v] of [['visibilityState', 'hidden'], ['hidden', true]])
    Object.defineProperty(document, k, {configurable: true, get: () => v});
  document.dispatchEvent(new Event('visibilitychange'));
  document.dispatchEvent(new Event('freeze'));
}"""
_SHOW = """() => {
  // 書き換えを外して、ブラウザ本来の値に戻す(残すと、閉じるときの「裏に回った」まで
  // 「前に戻った」として送られてしまう)
  delete document.visibilityState;
  delete document.hidden;
  document.dispatchEvent(new Event('resume'));
  document.dispatchEvent(new Event('visibilitychange'));
  window.dispatchEvent(new Event('focus'));
}"""


class Frozen:
    """ページの JavaScript を止める(ブラウザがタブを凍結したのと同じ)。

    ブラウザは裏に回したタブのタイマーを間引き、長く裏にあるタブは凍結する。
    ここでは `hide=True` なら先に「裏に回った」合図(visibilitychange・freeze)を出してから、
    デバッガでページの実行を止める。止めているあいだは心拍も何も出ない。
    """

    def __init__(self, pages, *, hide: bool):
        self.pages, self.hide, self.sessions = list(pages), hide, []

    def __enter__(self):
        for pg in self.pages:
            if self.hide:
                for fr in pg.frames:
                    try:
                        fr.evaluate(_HIDE)
                    except Exception:            # noqa: BLE001 - 読み込み中の枠は飛ばす
                        pass
            cdp = pg.context.new_cdp_session(pg)
            cdp.send("Debugger.enable")
            cdp.send("Debugger.pause")
            self.sessions.append(cdp)
        return self

    def __exit__(self, *exc):
        for pg, cdp in zip(self.pages, self.sessions):
            cdp.send("Debugger.resume")
            cdp.send("Debugger.disable")
            cdp.detach()
            if self.hide:
                for fr in pg.frames:
                    try:
                        fr.evaluate(_SHOW)
                    except Exception:            # noqa: BLE001
                        pass
        return False


def _hold(app: App, seconds: float, label: str) -> bool:
    """放置する。サーバが途中で終わらないかを見ながら待つ。"""
    print(f"   … {label} {seconds:.0f}秒", flush=True)
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        if not app.alive():
            return False
        time.sleep(min(5.0, max(0.1, end - time.monotonic())))
    return app.alive()


def _after_idle(sh: Shell, label: str) -> None:
    """放置のあと: 断り・接続なしが出ず、3機能とも続けて操作できる。"""
    sh.page.wait_for_timeout(3000)
    trouble = [] if sh.ready(15) else sh.trouble()
    check(f"{label}: 断り・接続なし・開き直しが出ない", not trouble, trouble)
    p = sh.tab("pena")
    p.click("button[data-act=calcTare]")
    check(f"{label}: ペナラベルで計算できる", pena_toast(p, "計算しました"), text(p, "#toast .toast-body"))
    m = sh.tab("material")
    m.click("#run")
    check(f"{label}: 資材計算で計算できる", material_toast(m, "台数"), text(m, "#toast"))
    d = sh.tab("details")
    codes = []
    for _ in range(2):                          # 反転して、元に戻す
        with d.page.expect_response(lambda r: "/details/api/reverse" in r.url) as info:
            d.click("#btnReverse")
        codes.append(info.value.status)
    check(f"{label}: 梱包明細の画面を操作できる(断られない)",
          codes == [200, 200] and not visible(d, "#takenOver"), codes)


def scenario_idle(app: App, ctx, sh: Shell, note: dict, minutes: dict) -> None:
    print("■ 3. 放置する")
    pid = (app.health() or {}).get("pid")

    ok = _hold(app, minutes["visible"], "a. 表のまま放置")
    check("放置a(表のまま): サーバが終わらない", ok, "")
    _after_idle(sh, "放置a(表のまま)")

    pages = [p for p in ctx.pages if not p.is_closed()]
    with Frozen(pages, hide=True):
        ok = _hold(app, minutes["hidden"], "b. 裏に回して凍結")
    check("放置b(裏に回して凍結): サーバが終わらない", ok, f"{len(pages)} 枚を凍結")
    _after_idle(sh, "放置b(裏→凍結→戻る)")

    with Frozen(pages, hide=False):
        os.kill(app.proc.pid, signal.SIGSTOP)
        try:
            time.sleep(minutes["sleep"])
        finally:
            os.kill(app.proc.pid, signal.SIGCONT)
    ok = _hold(app, 20, "c. スリープから戻った直後")
    check("放置c(スリープ): 戻ってもサーバが終わらない", ok, "")
    _after_idle(sh, "放置c(スリープから戻る)")
    # c2. スリープから戻ったが、サーバだけ先に動き出し、画面は 100 秒止まったまま(統合 1.2.56)。
    #     以前は戻ってから数え直した 90 秒しか待たず、ブラウザのタブが休止したまま・戻りが遅いと終わった
    with Frozen(pages, hide=False):
        os.kill(app.proc.pid, signal.SIGSTOP)
        try:
            time.sleep(20)
        finally:
            os.kill(app.proc.pid, signal.SIGCONT)
        ok = _hold(app, 100, "c2. 戻ったがサーバだけ先に動き、画面は止まったまま")
    check("放置c2(スリープから戻り画面だけ遅れる): 画面が止まっていても終わらない", ok, "")
    _after_idle(sh, "放置c2(画面が動き出す)")
    # c3. 確認の窓(confirm)を出したまま 100 秒(統合 1.2.56)。窓のあいだは画面の JavaScript が止まる
    m = sh.tab("material")
    material_ready(m)
    n0 = app.logs().count("(dialog)") + app.logs().count(": dialog ")
    DIALOG_WAIT["sec"] = 100.0
    try:
        m.evaluate("() => setTimeout(() => document.querySelector('#weight').click(), 50)")
        print("   … c3. 確認の窓を出したまま 100秒", flush=True)
        sh.page.wait_for_timeout(1000)            # ここで窓が出て、押すまで 100 秒待つ(_on_dialog)
    finally:
        DIALOG_WAIT["sec"] = 0.0
    signalled = (app.logs().count("(dialog)") + app.logs().count(": dialog ")) > n0
    check("放置c3(確認の窓を出したまま 100 秒): サーバが終わらない", app.alive(),
          "窓の前の合図あり" if signalled else "窓の前の合図がログに無い")
    _after_idle(sh, "放置c3(窓を閉じたあと)")
    check("放置a〜c: サーバは同じプロセスのまま(開き直しになっていない)",
          (app.health() or {}).get("pid") == pid, pid)

    report = note.get("report")
    if report is not None and not report.is_closed():
        # 帳票の別タブを開いたまま統合画面を閉じてしまい、すぐ開き直す
        sh.page.close(run_before_unload=True)
        time.sleep(1.5)
        sh.page = ctx.new_page()
        sh.page.goto(app.base + "/?t=" + app.token)
        wait_until(lambda: all(sh.frame(k) for k in ("details", "pena", "material")), 20)
        sh.page.wait_for_timeout(4000)
        trouble = sh.trouble()
        check("統合画面を閉じてすぐ開き直しても、3機能とも断られない", not trouble, trouble)
        sh.page.close(run_before_unload=True)
        ok = _hold(app, minutes["report"], "d. 統合画面を閉じ、帳票の別タブだけで放置")
        check("放置d(帳票の別タブだけ): サーバが終わらない", ok, "")
        report.bring_to_front()
        report.click('span.edit[data-placeholder="サイズ"]')
        report.keyboard.press("End")
        report.keyboard.type("-2")
        report.keyboard.press("Enter")
        check("放置d: 放置のあと帳票の書き足しが保存できる",
              wait_until(lambda: "保存しました" in text(report, "#editSaved"), 8), text(report, "#editSaved"))
        sh.page = ctx.new_page()
        sh.page.goto(app.base + "/?t=" + app.token)
        wait_until(lambda: all(sh.frame(k) for k in ("details", "pena", "material")), 20)
        sh.page.wait_for_timeout(2500)
        trouble = [] if sh.ready(15) else sh.trouble()
        check("放置d: 統合画面を開き直すと3機能とも使える", not trouble, trouble)
        d = sh.tab("details")
        check("放置d: 梱包明細は開いていたロットのまま", text(d, "#lotNo") == "A123456", text(d, "#lotNo"))

    # f. 本ツール以外のタブを開いて、本ツールのタブを裏へ。長く裏にあるとブラウザは
    #    タブを**捨てる**(Chrome のメモリセーバー・Edge のタブのスリープ)。捨てるときは
    #    何の合図も出さない。戻ると読み直される
    other = ctx.new_page()
    other.set_content("<h1>別のサイト</h1><script>setInterval(() => { let x = 0; "
                      "for (let i = 0; i < 2e5; i++) x += i; }, 200)</script>")
    for fr in sh.page.frames:
        try:
            fr.evaluate(_HIDE)
        except Exception:                          # noqa: BLE001
            pass
    time.sleep(1.5)
    for fr in sh.page.frames:                      # 捨てる = もう何も送れない
        try:
            fr.evaluate("() => { navigator.sendBeacon = () => true;"
                        " window.fetch = () => new Promise(() => {}); }")
        except Exception:                          # noqa: BLE001
            pass
    with Frozen([sh.page], hide=False):
        ok = _hold(app, minutes["hidden"], "f. 本ツールのタブが裏で捨てられる")
    check("放置f(ほかのタブを開き、裏のタブが捨てられる): サーバが終わらない", ok, "")
    sh.page.reload()
    wait_until(lambda: all(sh.frame(k) for k in ("details", "pena", "material")), 20)
    sh.page.wait_for_timeout(3000)
    trouble = [] if sh.ready(15) else sh.trouble()
    check("放置f: 戻る(読み直される)と3機能とも続けて使える", not trouble, trouble)
    d = sh.tab("details")
    check("放置f: 梱包明細は開いていたロットのまま", text(d, "#lotNo") == "A123456", text(d, "#lotNo"))
    other.close()

    # e. 全部閉じて放置 → 自分で終わる → 翌朝と同じく開き直す
    for pg in list(ctx.pages):
        pg.close(run_before_unload=True)
    try:
        app.proc.wait(timeout=40)
        ended = True
    except subprocess.TimeoutExpired:
        ended = False
    check("放置e: 全部閉じたら自分で終わる", ended, "")
    check("放置e: 終わったあと多重起動の印(ロック)が残らない",
          not (app.local / "runtime" / "main.lock").exists(), "")
    h = app.start()
    check("放置e: 翌朝と同じく起動し直せる", h, (h or {}).get("pid"))
    sh.page = ctx.new_page()
    sh.page.goto(app.base + "/?t=" + app.token)
    wait_until(lambda: all(sh.frame(k) for k in ("details", "pena", "material")), 20)
    sh.page.wait_for_timeout(2500)
    trouble = [] if sh.ready(15) else sh.trouble()
    check("放置e: 開き直した統合画面で3機能とも使える", not trouble, trouble)
    p = sh.tab("pena")
    want = note.get("pena", ("X222222", "220.0"))   # 最後に入れた値
    check("放置e(残るもの): ペナラベルの検査番号と風袋計算の結果",
          (text(p, "#lblKensaNo"), text(p, "#nw1")) == want,
          (text(p, "#lblKensaNo"), text(p, "#nw1")))
    m = sh.tab("material")
    material_ready(m)
    m.click('nav.rail a.rail__item[href*="/checklist"]')
    m.wait_for_load_state()
    m = sh.frame("material")
    material_ready(m)
    check("放置e(残るもの): 資材計算のチェックリスト2行", wait_until(lambda: text(m, "#sub").startswith("2 /"), 8),
          text(m, "#sub"))
    d = sh.tab("details")
    details_open_lot(d, "A123456")
    check("放置e(残るもの): 梱包明細の出力済みの明細(No1)",
          wait_until(lambda: visible(d, "#outputCard") and "A123456-No1" in text(d, "#outputCard"), 8),
          text(d, "#outputCard")[:40])


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--only", choices=("flow", "print", "alternate", "multitab", "storage", "master",
                                       "trail", "theme", "manual", "idle", "rapid"))
    ap.add_argument("--shots", default="", help="保存先の画面の写しを置くフォルダ")
    ap.add_argument("--quick", action="store_true", help="放置を短くする")
    ap.add_argument("--keep", action="store_true", help="試験用の置き場所を残す")
    args = ap.parse_args(argv)
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright がありません(tools/smoke_shell.py と同じものが要ります)")
        return 2

    home = Path(tempfile.mkdtemp(prefix="cpt-e2e-"))
    make_sources(home)
    app = App(home)
    print("=" * 80)
    print("置き場所:", home)
    h = app.start()
    check("起動して取り込みまで済む(本物の start_app)", h, (h or {}).get("version_set", "起動しない"))
    if not h:
        return 1
    note: dict = {}
    errors: list = []
    with sync_playwright() as pw:
        exe = os.environ.get("PLAYWRIGHT_CHROMIUM", "/opt/pw-browsers/chromium")
        # ポップアップのブロックは**現場のブラウザと同じく入れたまま**(Playwright は既定で切る)。
        # 押してから時間がたって開く窓は、現場では止められる(統合 1.2.20 で見つかった)
        br = pw.chromium.launch(executable_path=exe if os.path.exists(exe) else None,
                                ignore_default_args=["--disable-popup-blocking"])
        ctx = br.new_context(viewport={"width": 1400, "height": 900})
        ctx.on("page", lambda p: p.on("pageerror", lambda e: errors.append(f"{p.url.split('?')[0]}: {e}")))
        ctx.on("page", lambda p: p.on("dialog", _on_dialog))
        # 開いたときに自分で刷る画面(ペナラベルの印刷ビュー)が何回刷ったかを数える
        ctx.add_init_script("window.__printed = 0; window.print = function () { window.__printed++; };")
        sh = Shell(ctx, app)
        if args.only in (None, "flow", "print", "alternate", "multitab", "idle", "rapid"):
            scenario_flow(app, ctx, sh, note)
        if args.only in (None, "print"):
            scenario_print(app, ctx, sh, note, args.shots)
        if args.only in (None, "rapid"):
            scenario_rapid(app, ctx, sh, note)
        if args.only in (None, "alternate", "multitab", "idle"):
            scenario_alternate(app, ctx, sh, note)
        if args.only in (None, "multitab"):
            scenario_multitab(app, ctx, sh, note)
        if args.only in (None, "storage"):
            scenario_storage(app, ctx, sh, note, args.shots)
        if args.only in (None, "master"):
            scenario_master(app, ctx, sh, note)
        if args.only in (None, "trail"):
            scenario_trail(app, ctx, sh, note)
        if args.only in (None, "theme"):
            scenario_theme(app, br, note)
        if args.only in (None, "manual"):
            scenario_manual(app, ctx, sh, note)
        if args.only in (None, "idle"):
            minutes = ({"visible": 125, "hidden": 125, "sleep": 30, "report": 110} if args.quick else
                       {"visible": 360, "hidden": 300, "sleep": 120, "report": 150})
            scenario_idle(app, ctx, sh, note, minutes)
        # わざと起こしたエラー(「エラーの記録」の段)は数えない
        errors = [e for e in errors if TRAIL_MARK not in e]
        check("画面の JavaScript のエラーが無い", not errors, errors[:3])
        br.close()
    app.stop()
    logs = app.logs()
    bad = [l for l in logs.splitlines()
           if ("| ERROR |" in l or "Traceback" in l) and TRAIL_MARK not in l]
    check("ログに ERROR・Traceback が無い", not bad, bad[:2])
    print("=" * 80)
    ng = [r for r in results if not r[1]]
    print("結果: %d 項目中 %d 件 NG" % (len(results), len(ng)))
    if not args.keep and not ng:
        import shutil
        shutil.rmtree(home, ignore_errors=True)
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
