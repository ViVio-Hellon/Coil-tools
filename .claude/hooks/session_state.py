"""作業状態ファイル(.claude/SESSION_STATE.md)の見張り役 — Claude Code の hook から呼ぶ

  precompact   … 圧縮の直前。状態ファイルを state_backups/ に控え、末尾の「自動記録」欄だけを
                 いまの Git の状態に書き換える(欄は毎回置き換えるので溜まらない)。
  sessionstart … 起動・再開・圧縮のあと。状態ファイルの中身を Claude に読ませる。

決まりごと
  - 標準ライブラリだけ。Windows(Git Bash / python)でも Linux でも動く。
  - 何が起きても終了コード 0。圧縮やセッション開始を止めない。
  - プロジェクトの場所は CLAUDE_PROJECT_DIR、無ければこのファイルの 2 つ上。
"""
from __future__ import annotations

import datetime as _dt
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

KEEP_BACKUPS = 20          # 控えは新しいものから 20 個まで
MAX_CONTEXT_CHARS = 15000  # 読ませる量の上限(長すぎると文脈を食う)
GIT_TIMEOUT_SEC = 5
AUTO_START = "<!-- AUTO:GIT START (圧縮のたびに hook が書き換える。手で書かない) -->"
AUTO_END = "<!-- AUTO:GIT END -->"

TEMPLATE = """# SESSION_STATE — 作業の引き継ぎメモ

> 圧縮・再開のあとに Claude が最初に読むファイル。会話の全文は書かない。
> 完了した細かい話は消して、決定事項と未完了を優先して残す。

## 現在の作業目的

## 完了した作業(直近のみ・古いものは消す)

## 未完了の作業

## 重要な仕様・設計上の決定

## 発生した不具合と解決状況

## 変更したファイル(直近の作業分)

## 次に実行する作業

"""


def project_dir() -> Path:
    env = os.environ.get("CLAUDE_PROJECT_DIR")
    if env and Path(env).is_dir():
        return Path(env)
    return Path(__file__).resolve().parents[2]


def paths():
    root = project_dir()
    state = root / ".claude" / "SESSION_STATE.md"
    return root, state, root / ".claude" / "state_backups"


def read_stdin() -> dict:
    try:
        data = sys.stdin.read()
        return json.loads(data) if data.strip() else {}
    except Exception:
        return {}


def now() -> _dt.datetime:
    return _dt.datetime.now()


def git(root: Path, *args: str) -> str:
    try:
        out = subprocess.run(["git", *args], cwd=str(root), capture_output=True,
                             timeout=GIT_TIMEOUT_SEC)
        if out.returncode != 0:
            return ""
        return out.stdout.decode("utf-8", "replace").strip()
    except Exception:
        return ""


def git_block(root: Path, trigger: str) -> str:
    branch = git(root, "rev-parse", "--abbrev-ref", "HEAD") or "(不明)"
    status = git(root, "status", "--short") or "(変更なし)"
    lines = status.splitlines()
    if len(lines) > 40:
        status = "\n".join(lines[:40] + [f"... ほか {len(lines) - 40} 件"])
    log = git(root, "log", "--oneline", "-8") or "(取れず)"
    return "\n".join([
        AUTO_START,
        "## 自動記録(最後の圧縮の直前)",
        f"- 日時: {now():%Y-%m-%d %H:%M:%S} / 圧縮: {trigger or '不明'} / ブランチ: {branch}",
        "",
        "未コミットの変更:",
        "```",
        status,
        "```",
        "直近のコミット:",
        "```",
        log,
        "```",
        AUTO_END,
    ])


def replace_auto(text: str, block: str) -> str:
    """自動記録欄を置き換える。無ければ末尾に足す(何度圧縮しても 1 つだけ)。"""
    s, e = text.find(AUTO_START), text.find(AUTO_END)
    if s != -1 and e != -1 and e > s:
        return text[:s] + block + text[e + len(AUTO_END):]
    return text.rstrip("\n") + "\n\n" + block + "\n"


def write_atomic(path: Path, text: str) -> None:
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    os.replace(tmp, path)


def backup(state: Path, backups: Path, trigger: str) -> Path | None:
    if not state.is_file():
        return None
    backups.mkdir(parents=True, exist_ok=True)
    dest = backups / f"SESSION_STATE_{now():%Y%m%d-%H%M%S}_{trigger or 'unknown'}.md"
    n = 1
    while dest.exists():
        dest = dest.with_name(f"{dest.stem}_{n}.md")
        n += 1
    shutil.copy2(state, dest)
    old = sorted(backups.glob("SESSION_STATE_*.md"), key=lambda p: p.name)
    for p in old[:-KEEP_BACKUPS]:
        try:
            p.unlink()
        except OSError:
            pass
    return dest


def precompact(payload: dict) -> dict:
    root, state, backups = paths()
    trigger = str(payload.get("trigger") or "")
    saved = backup(state, backups, trigger)
    text = state.read_text(encoding="utf-8") if state.is_file() else TEMPLATE
    write_atomic(state, replace_auto(text, git_block(root, trigger)))
    where = saved.name if saved else "(初回のため控えなし・ひな形を作成)"
    return {"systemMessage": f"SESSION_STATE.md を控えました: {where}"}


def sessionstart(payload: dict) -> dict:
    _, state, _ = paths()
    source = str(payload.get("source") or "")
    if not state.is_file():
        ctx = ("[SESSION_STATE] .claude/SESSION_STATE.md はまだありません。"
               "作業を始めたら、ひな形の見出しに沿って作ってください。")
    else:
        body = state.read_text(encoding="utf-8")
        if len(body) > MAX_CONTEXT_CHARS:
            body = body[:MAX_CONTEXT_CHARS] + (
                f"\n\n...(長すぎるため {MAX_CONTEXT_CHARS} 文字で切りました。"
                "完了済みの古い項目を消して短くしてください)")
        ctx = "\n".join([
            f"[SESSION_STATE] セッション開始({source or '不明'})。"
            "前回までの作業状態です。作業目的・未完了・次の作業を確かめてから続けてください。"
            "作業が進んだら .claude/SESSION_STATE.md を更新してください(完了した細目は消す)。",
            "",
            body,
        ])
    return {"hookSpecificOutput": {"hookEventName": "SessionStart",
                                   "additionalContext": ctx}}


def main(argv) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    mode = argv[1] if len(argv) > 1 else ""
    payload = read_stdin()
    try:
        if mode == "precompact":
            out = precompact(payload)
        elif mode == "sessionstart":
            out = sessionstart(payload)
        else:
            return 0
        print(json.dumps(out, ensure_ascii=False))
    except Exception as exc:  # 何があっても止めない
        try:
            sys.stderr.write(f"session_state hook: {exc}\n")
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
