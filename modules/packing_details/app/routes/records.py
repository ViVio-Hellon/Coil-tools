"""作業の記録の画面 (VER 0.13.30。ペナラベルの「計算・印刷の記録」と同じ考え)

    GET /report/records     作業の記録(この PC の新しい順に 300 件)と、資材マスタの編集履歴

`/report/` の下に置く ── 紙面と同じく**トークンは要る・画面の持ち主でなくてよい**
(設定から別のタブで開くため)。読むだけ。
"""
from __future__ import annotations

from typing import Any

from flask import Blueprint, Response

from modules.packing_details.meisai import edit_history, printing, work_record

bp = Blueprint("records", __name__)

LIMIT = 300

esc = printing.escape

_CSS = """
body{margin:0 auto; max-width:90rem; padding:1.25rem 16px 3rem;
  font-family:"Meiryo UI",Meiryo,system-ui,sans-serif; font-size:13px; line-height:1.7;
  color:#13232b; background:#fff}
h1{font-size:1.35rem; border-bottom:2px solid #0e6a72; padding-bottom:.3rem}
h2{font-size:1.05rem; margin-top:2rem; border-left:4px solid #0e6a72; padding-left:.5rem}
.hint{color:#3b4a52}
.bad{color:#b3261e; font-weight:bold}
.scroll{overflow-x:auto}
table{border-collapse:collapse; width:100%; font-size:12px}
th,td{border:1px solid #cfdadf; padding:.3rem .45rem; text-align:left; vertical-align:top}
th{background:#eef2f4; position:sticky; top:0}
code{background:#eef2f4; padding:.05rem .3rem; border-radius:3px; font-family:Consolas,monospace}
.nowrap{white-space:nowrap}
.sub{font-size:11px; opacity:.75}
.old{opacity:.6}
.test{color:#fff; background:#b3261e; padding:0 6px; border-radius:8px; font-size:11px;
  font-weight:bold}
.kind-out{font-weight:bold}
.kind-clear{font-weight:bold; color:#b3261e}
.empty{padding:1rem; color:#5b6b73}
@media (prefers-color-scheme:dark){
  body{background:#0f1a1e; color:#e3ecef}
  .hint{color:#b8c7cd}
  th{background:#1d2b30}
  th,td{border-color:#2f4249}
  code{background:#1d2b30}
}
"""

_KIND_CLASS = {"出力": "kind-out", "出力し直し": "kind-out", "全クリア": "kind-clear"}


def _join(values: Any, sep: str = " ") -> str:
    if isinstance(values, (list, tuple)):
        return sep.join(str(v) for v in values)
    return "" if values is None else str(values)


def _nos(item: dict) -> str:
    if item.get("no") not in (None, ""):
        return f"No{item['no']}"
    nos = item.get("nos")
    if isinstance(nos, list) and nos:
        return "・".join(f"No{n}" for n in nos)
    return ""


def _detail(item: dict) -> str:
    """中身(種類ごとに要るものだけ)。"""
    parts = []
    if item.get("coils"):
        parts.append(f"副番 {esc(_join(item['coils']))}")
    if item.get("weights"):
        parts.append(f"重量(丈ごと) {esc(_join(item['weights'], ' / '))}")
    if item.get("historyId"):
        parts.append(f'<span class="sub">履歴ID {esc(item["historyId"])}</span>')
    if item.get("qa"):
        parts.append(f"右上の文字 {esc(item['qa'])}")
    if "cancelled" in item:
        parts.append(f"明細の履歴を「取り消し」にした紙 {esc(item['cancelled'])}枚")
    if item.get("slips"):
        parts.append(f"紙 {esc(_join(item['slips'], '・'))}")
    if item.get("what"):
        parts.append(f"{esc(item['what'])}: {esc(item.get('value', '')) or '(空)'}")
    return "<br>".join(parts)


def _who(item: dict) -> str:
    w = item.get("who") or {}
    if not w:
        return '<span class="old">(記録なし)</span>'
    out = esc(w.get("pc", ""))
    if w.get("line"):
        out += f'<br><span class="sub">{esc(w["line"])}</span>'
    if w.get("user"):
        out += f'<br><span class="sub">{esc(w["user"])}</span>'
    if w.get("mode"):
        out += f'<br><span class="test">{esc(w["mode"])}</span>'
    return out


def _version(item: dict) -> str:
    v = item.get("version") or {}
    text = " / ".join(x for x in (("統合 " + v["app"]) if v.get("app") else "",
                                  ("梱包明細 " + v["details"]) if v.get("details") else "") if x)
    return esc(text) if text else '<span class="old">(記録なし)</span>'


def _records_card() -> str:
    rows = []
    for item in work_record.recent(LIMIT):
        kind = str(item.get("kind", ""))
        rows.append(
            f'<tr><td class="nowrap">{esc(item.get("at", ""))}</td>'
            f'<td class="nowrap"><span class="{_KIND_CLASS.get(kind, "")}">{esc(kind)}</span></td>'
            f'<td class="nowrap">{esc(item.get("lot", ""))}</td>'
            f'<td>{esc(_nos(item))}</td>'
            f'<td class="nowrap">{_who(item)}</td>'
            f'<td class="nowrap">{_version(item)}</td>'
            f'<td>{_detail(item)}</td></tr>')
    st = work_record.status()
    keep = ('<p class="hint"><b>消えない控え</b>: 1 件ごとに月ごとのファイルへ足しています（消さない・書き換えない）。'
            f'形式: <b>{esc(st["format_label"])}</b>（設定「記録ファイルの形式」）。今月のファイル: '
            + "・".join(f'<code>{esc(n)}</code>' for n in st["files"])
            + '。CSV は Excel でそのまま開けます（見るときは写しを開くか、読み取り専用で開いてください。'
              '開いている間に書けなかった分は「追記待ち」に足し、閉じたあと次の記録で書き足します）。'
            f'<br>この PC: <code>{esc(st["local"])}</code>')
    if st["local_problem"]:
        keep += f'<br><span class="bad">書けていません: {esc(st["local_problem"])}</span>'
    if st["shared"]:
        keep += f'<br>設定の保存先: <code>{esc(st["shared"])}</code>'
    else:
        keep += ('<br><span class="bad">この PC の中にしかありません。</span>設定の「置き場所・取り込み」の'
                 '「記録の保存先フォルダー」に共有フォルダーを入れると、そこへも PC ごとに書きます'
                 '（PC の入れ替えでも残るように）。')
    if st["shared_problem"]:
        keep += f'<br><span class="bad">設定の保存先に書けていません: {esc(st["shared_problem"])}</span>'
    keep += "</p>"
    table = ('<div class="scroll"><table><thead><tr><th>時刻</th><th>種類</th><th>ロット</th>'
             '<th>No</th><th>PC・ライン</th><th>版</th><th>中身</th></tr></thead><tbody>'
             + "".join(rows) + "</tbody></table></div>"
             if rows else '<div class="empty">まだ記録がありません（今月と先月の分を見ています）。</div>')
    return ('<h2>作業の記録（この PC）</h2>'
            f'<p class="hint">出力・出力し直し・修正・全クリア・印刷画面・紙面を開いた、を<b>どの版で・'
            f'どの PC（ライン）で</b>したかを残しています（新しい順に {LIMIT} 件。今月と先月の分）。<br>'
            '<b>印刷画面</b>は「印刷」を押して紙面を開いた記録で、<b>刷った証拠にはなりません</b>'
            '（ブラウザは、刷ったか・やめたかを教えません）。'
            '<b>紙面を開いた</b>は紙面のタブを開いた（開き直した）記録です。'
            '<b>テストモード</b>の札が付いた行は、実機の試験中の記録です。</p>'
            + keep + table)


def _history_card() -> str:
    hist = edit_history.read(LIMIT)
    rows = "".join(
        f'<tr><td class="nowrap">{esc(h.get("日時", ""))}</td>'
        f'<td class="nowrap">{esc(h.get("PC", ""))}'
        + (f'<br><span class="sub">{esc(h.get("ライン", ""))}</span>' if h.get("ライン") else "")
        + (f'<br><span class="test">{esc(h.get("モード", ""))}</span>' if h.get("モード") else "")
        + f'</td><td>{esc(h.get("利用者", ""))}</td><td class="nowrap">{esc(h.get("表", ""))}</td>'
        f'<td>{esc(h.get("行の名前", "") or h.get("行", ""))}</td>'
        f'<td class="nowrap">{esc(h.get("操作", ""))}</td>'
        f'<td>{esc(h.get("列", ""))}</td><td>{esc(h.get("前", ""))}</td>'
        f'<td>{esc(h.get("後", ""))}</td></tr>'
        for h in hist["rows"])
    return ('<h2>資材マスタの編集履歴（全ライン）</h2>'
            '<p class="hint">紙面の右上の文字・管理者パスワードを変えたときに、<b>誰が（PC・ライン・'
            'ログインID）・いつ・何を何から何に</b>変えたかを、梱包資材マスタの中（表「'
            + esc(edit_history.TABLE) + '」）に残しています（新しい順に '
            f'{LIMIT} 件。ペナラベルがマスタを直した行も同じ表に入ります）。'
            '管理者パスワードは変えたことだけを残し、値は残しません（前・後は「'
            + esc(edit_history.HIDDEN) + '」）。このツールの外（Access など）で直したものは残りません。'
            + (f'<br>梱包資材マスタ: <code>{esc(hist["path"])}</code>' if hist.get("path") else "")
            + "</p>"
            + (f'<p class="bad">{esc(hist["note"])}</p>' if hist.get("note") else "")
            + ('<div class="scroll"><table><thead><tr><th>日時</th><th>PC・ライン</th><th>利用者</th>'
               '<th>表</th><th>行</th><th>操作</th><th>列</th><th>前</th><th>後</th></tr></thead>'
               '<tbody>' + rows + '</tbody></table></div>'
               if rows else '<div class="empty">まだ編集履歴がありません。</div>'))


@bp.get("/report/records")
def records_page() -> Response:
    """作業の記録と、資材マスタの編集履歴。読むだけ。"""
    badge = ('<p><span class="test">テストモード</span> この PC はいまテストモードです'
             '（試験が終わったら設定で切ってください）。</p>' if work_record.test_mode() else "")
    body = (f"<h1>作業の記録</h1>{badge}" + _records_card() + _history_card())
    return Response(
        '<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1">'
        "<title>作業の記録</title>"
        f"<style>{_CSS}</style></head><body>{body}</body></html>",
        mimetype="text/html")
