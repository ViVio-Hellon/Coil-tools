/*
  tab_alive.js — 別のタブで開いた画面の心拍(統合アプリが差し込む)

  【なぜ要るのか】
  3機能とも**印刷は別のタブ**で開く。

    梱包明細  帳票(`/details/report/…`)。刷る前に紙面で書き足したサイズ・
              LOTNO を**サーバへ保存する**。右上の文字も開き直すたびに聞き直す
    ペナラベル 印刷ビュー(`/pena/…/print`)。バーコードのフォントもサーバから読む
    資材計算  チェックリスト・発注票(`/material/report/…`)。発注票は
              **サーバが覚えている中身**から組む(サーバが止まると出せない)

  統合アプリは、画面が全部居なくなると自分で終わる(common/idle_exit.py)。
  統合画面のタブを閉じても、開いたままの帳票・印刷ビューがあるあいだは
  終わってはいけない ── そこで**別のタブで開いた画面も自分で心拍を送る**。
  見張りは画面の名乗りごとに数えるので、どれか1枚でも開いていれば終わらない。

  【統合画面の中(iframe)では何もしない】
  外枠(shell.js)が心拍を送っているので、ここで送ると二重になる。

  【各機能のコードは触らない】
  このスクリプトは統合アプリが HTML の末尾に差し込む(app.py)。画面にも
  紙面にも何も描かないので、印刷の見た目は変わらない。
*/
(function () {
  "use strict";
  if (window.top !== window) { return; }          // 統合画面の中

  var me = document.currentScript;
  function attr(name, fallback) {
    var v = me && me.getAttribute(name);
    return v ? v : fallback;
  }
  var ALIVE_MS = parseInt(attr("data-alive-ms", ""), 10) || 20000;
  var URL_ALIVE = attr("data-url", "/api/alive");
  // 名乗りは**開くたびに新しい名前**(タブを複製しても別の画面として数える)
  var CLIENT = "tab-" + attr("data-key", "page") + "-" + Date.now().toString(36) + "-" +
               Math.random().toString(36).slice(2, 10);

  function state() { return document.visibilityState === "hidden" ? "hidden" : "visible"; }
  function body(extra) {
    var b = { client: CLIENT, state: state() };
    Object.keys(extra || {}).forEach(function (k) { b[k] = extra[k]; });
    return JSON.stringify(b);
  }

  // ふつうの心拍
  function alive(reason, extra) {
    var b = { reason: reason };
    Object.keys(extra || {}).forEach(function (k) { b[k] = extra[k]; });
    try {
      fetch(URL_ALIVE, {
        method: "POST", cache: "no-store", body: body(b),
        headers: { "Content-Type": "application/json" }
      }).catch(function () { /* 届かなければ次の心拍で */ });
    } catch (e) { /* 同上 */ }
  }

  // 閉じる・裏に回る瞬間でも届く送り方
  function signal(extra) {
    var text = body(extra);
    try {
      if (navigator.sendBeacon &&
          navigator.sendBeacon(URL_ALIVE, new Blob([text], { type: "application/json" }))) {
        return;
      }
    } catch (e) { /* 下の fetch へ */ }
    try {
      fetch(URL_ALIVE, { method: "POST", keepalive: true, body: text,
                         headers: { "Content-Type": "application/json" } });
    } catch (e) { /* 届かなくても次の心拍で直る */ }
  }

  // 閉じた合図は**2つの道で**送る(統合 1.2.79)。sendBeacon 1回だけだと、閉じた瞬間に PC が重いと
  // 届かないことがあり(通し試験の CI で再現)、その画面が「裏に回ったまま」と数えられ続けて、全部閉じても
  // 終わらなかった。重なって届いても害は無い(閉じるのをやめれば、次の心拍が取り消す)
  var leftAt = 0;
  function leaving(reason) {
    var now = Date.now();
    if (now - leftAt < 300) { return; }          // 同じ閉じ方で続けて来たら1回ぶん
    leftAt = now;
    var text = body({ leaving: true, reason: reason });
    try {
      if (navigator.sendBeacon) {
        navigator.sendBeacon(URL_ALIVE, new Blob([text], { type: "application/json" }));
      }
    } catch (e) { /* 下の fetch で */ }
    try {
      fetch(URL_ALIVE, { method: "POST", keepalive: true, body: text,
                         headers: { "Content-Type": "application/json" } }).catch(function () {});
    } catch (e) { /* 届かなければそれまで */ }
  }

  // 同じタブの前の名乗り(ブラウザが裏のタブを捨てて読み直したとき。shell.js と同じ)
  var PREV_KEY = "cpt.tab." + attr("data-key", "page") + ".client", previous = "";
  try { previous = sessionStorage.getItem(PREV_KEY) || ""; sessionStorage.setItem(PREV_KEY, CLIENT); }
  catch (e) { previous = ""; }
  alive("open", { replaces: previous });
  setInterval(function () { alive("timer"); }, ALIVE_MS);
  document.addEventListener("visibilitychange", function () {
    if (state() === "visible") { alive("foreground"); } else { signal({ reason: "hidden" }); }
  });
  // 印刷の画面が開いているあいだは JavaScript が止まり心拍が途切れる。用紙を替えに行くなどで
  // 長くなっても終わらないよう、開く前に「裏に回った」と同じ合図を送る(統合 1.2.20)
  window.addEventListener("beforeprint", function () { signal({ reason: "print", state: "hidden" }); });
  window.addEventListener("afterprint", function () { alive("afterprint"); });
  // 確認の窓(confirm / alert / prompt)のあいだも JavaScript が止まる。押さずに 90 秒を超えても
  // 終わらないよう、印刷と同じく出す前に合図を送り、閉じたら戻す(統合 1.2.56)
  ["confirm", "alert", "prompt"].forEach(function (name) {
    var original = window[name];
    if (typeof original !== "function" || window.__cptDialogs) { return; }
    window[name] = function () {
      signal({ reason: "dialog", state: "hidden" });
      try { return original.apply(window, arguments); }
      finally { alive("dialog-closed"); }
    };
  });
  window.__cptDialogs = true;
  document.addEventListener("freeze", function () { signal({ reason: "freeze" }); });
  document.addEventListener("resume", function () { alive("resume"); });
  // 閉じる直前にも言っておく(pagehide まで JavaScript が動かないことがある)。閉じるのを
  // やめた・読み直しなら、次の心拍が取り消す
  window.addEventListener("beforeunload", function () { leaving("beforeunload"); });
  window.addEventListener("pagehide", function (e) {
    // 保存状態(bfcache)へ入るだけなら「裏に回った」、本当に閉じたら「閉じた」
    if (e.persisted) { signal({ reason: "hidden" }); }
    else { leftAt = 0; leaving("close"); }
  });
  window.addEventListener("pageshow", function (e) {
    if (e.persisted) { alive("restore"); }
  });
})();
