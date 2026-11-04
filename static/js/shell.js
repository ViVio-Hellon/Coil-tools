/*
  shell.js — 統合画面の外枠のふるまい

  【ここがすること】
    ・タブの切り替え(iframe を見せる/隠す。**外さない**)
    ・アプリ全体の心拍(`/api/alive`)と生存確認(`/api/health`)
    ・終了ボタン
    ・Ctrl+P で、いま見せている機能の画面を刷る(外枠を刷らない)

  【ここがしないこと】
    3機能の画面の中のこと。中の画面は自分の心拍・画面の見張り・
    入力チェックをそれぞれ持っていて、iframe の中でそのまま動く。

  【裏に回ったタブ】
  ブラウザは裏に回ったタブのタイマーを間引く(Chrome/Edge は5分後から
  1分に1回まで、Edge のスリープタブは止める)。裏に回るときに合図
  (`state: hidden`)を送っておけば、サーバは心拍の途切れで終わらない。
  中の画面も同じことをそれぞれ言うので、どちらか片方が届けば足りる。
*/
(function () {
  "use strict";

  var S = window.SHELL || {};
  var TOKEN = S.token || "";
  var TAB_KEY = "cpt.tab";
  var MISSES_BEFORE_OFFLINE = 2;
  // デスクトップ版(窓の中)。外からの停止は窓の × で来るので、入口にたずねに行かない
  var DESKTOP = !!(window.__TAURI__ && window.__TAURI__.core);
  var CLOCK_TICK_MS = 5000;
  var CLOCK_JUMP_MS = 30000;

  function $(id) { return document.getElementById(id); }

  // ---------------------------------------------------------- タブ
  var tabs = Array.prototype.slice.call(document.querySelectorAll(".tab"));
  var panes = Array.prototype.slice.call(document.querySelectorAll(".pane"));

  function show(key) {
    var found = false;
    panes.forEach(function (p) {
      var on = (p.id === "pane-" + key);
      p.hidden = !on;
      found = found || on;
    });
    if (!found) return false;
    tabs.forEach(function (t) {
      t.setAttribute("aria-selected", t.dataset.key === key ? "true" : "false");
    });
    try { sessionStorage.setItem(TAB_KEY, key); } catch (e) { /* 私用ウィンドウ等 */ }
    if (location.hash !== "#" + key) history.replaceState(null, "", "#" + key);
    // 見せた画面へ焦点を移す(キーボードで操作している人のため)
    var frame = document.querySelector("#pane-" + key + " iframe");
    if (frame) { try { frame.focus(); } catch (e) { /* 無視 */ } }
    return true;
  }

  tabs.forEach(function (t) {
    t.addEventListener("click", function () { show(t.dataset.key); });
  });
  // ← → でタブを移る
  document.querySelector(".tabs").addEventListener("keydown", function (e) {
    if (e.key !== "ArrowLeft" && e.key !== "ArrowRight") return;
    var i = tabs.findIndex(function (t) { return t.getAttribute("aria-selected") === "true"; });
    var next = tabs[(i + (e.key === "ArrowRight" ? 1 : tabs.length - 1)) % tabs.length];
    next.focus();
    show(next.dataset.key);
  });

  // 最初に見せるタブ: URL の #key → 前回のタブ → 先頭
  (function () {
    var want = (location.hash || "").slice(1);
    if (!want) { try { want = sessionStorage.getItem(TAB_KEY) || ""; } catch (e) { want = ""; } }
    if (!want || !show(want)) show(tabs.length ? tabs[0].dataset.key : "");
  })();
  window.addEventListener("hashchange", function () {
    var key = (location.hash || "").slice(1);
    if (key) show(key);
  });

  // ---------------------------------------------------------- 版の一覧
  // 統合ツールの版と3機能の版は分けて持つ。帯の版を押すと一覧が出る
  (function () {
    var btn = $("verBtn"), panel = $("verPanel");
    if (!btn || !panel) return;
    function set(open) {
      panel.hidden = !open;
      btn.setAttribute("aria-expanded", open ? "true" : "false");
    }
    btn.addEventListener("click", function (e) { e.stopPropagation(); set(panel.hidden); });
    document.addEventListener("click", function (e) {
      if (!panel.hidden && !panel.contains(e.target)) set(false);
    });
    document.addEventListener("keydown", function (e) { if (e.key === "Escape") set(false); });
    // 機能の画面(iframe)を押すと外枠の文書には click が来ないので、窓から焦点が外れたら閉じる
    window.addEventListener("blur", function () { set(false); });
  })();

  // ---------------------------------------------------------- 画面の色(統合 1.0.16)
  // 上の帯の「自動 / ライト / ダーク」。サーバに保存(このPC)し、この画面と、開いている
  // 画面すべて(3機能の iframe・ログ・別のタブの帳票)をその場で切り替える。
  // 各画面は差し込まれた theme.js が BroadcastChannel を受けて data-theme を付け替える。
  // 使えないブラウザに備えて、統合画面の中の iframe は直接も付け替える(同じ origin)。
  (function () {
    var box = $("themeSwitch");
    if (!box) return;
    var buttons = Array.prototype.slice.call(box.querySelectorAll("[data-theme-set]"));
    var channel = null;
    try { channel = new BroadcastChannel("cpt-theme"); } catch (e) { channel = null; }
    function applyTo(doc, theme) {
      try {
        var root = doc.documentElement;
        if (theme === "light" || theme === "dark") root.setAttribute("data-theme", theme);
        else root.removeAttribute("data-theme");
      } catch (e) { /* 読み込み中の枠は飛ばす(開いたときにサーバが付ける) */ }
    }
    function show(theme) {
      buttons.forEach(function (b) {
        b.setAttribute("aria-pressed", b.dataset.themeSet === theme ? "true" : "false");
      });
    }
    function apply(theme) {
      applyTo(document, theme);
      Array.prototype.forEach.call(document.querySelectorAll("iframe"), function (f) {
        try { applyTo(f.contentDocument, theme); } catch (e) { /* 別の origin は触らない */ }
      });
      if (channel) channel.postMessage({ type: "theme", theme: theme });
      show(theme);
    }
    buttons.forEach(function (b) {
      b.addEventListener("click", function () {
        var theme = b.dataset.themeSet;
        var before = (buttons.filter(function (x) {
          return x.getAttribute("aria-pressed") === "true"; })[0] || {}).dataset;
        apply(theme);                                   // 押したらすぐ変える
        fetch("/api/theme", {
          method: "POST", cache: "no-store",
          headers: { "Content-Type": "application/json", "X-Tool-Token": TOKEN },
          body: JSON.stringify({ theme: theme })
        }).then(function (r) {
          if (!r.ok) throw new Error(String(r.status));
        }).catch(function () {
          // 保存できなければ元へ戻す(次に開いた画面と食い違わせない)
          if (before && before.themeSet) apply(before.themeSet);
          window.alert("画面の色を保存できませんでした。もう一度押してください。");
        });
      });
    });
  })();

  // ---------------------------------------------------------- ログとエラーの記録
  // 上の帯の「ログ」。中身は /log(出力先の設定・エラーの一覧)。開くたびに読み直す。
  // 機能の画面の隅に出たエラー番号の「中身を見る」からも開く(postMessage。同じ origin だけ)
  (function () {
    var btn = $("logBtn"), dlg = $("logDialog"), body = $("logBody"), close = $("logClose");
    if (!btn || !dlg || !body) return;
    var frame = null;
    function open(id) {
      if (!frame) {
        frame = document.createElement("iframe");
        frame.className = "logdlg__frame";
        frame.title = "ログとエラーの記録";
        // ログの画面の確認の窓もくるむ(読み込むたびに文書が替わる。統合 1.2.56)
        frame.addEventListener("load", function () {
          if (window.cptWrapDialogs) window.cptWrapDialogs(frame.contentWindow);
        });
        body.appendChild(frame);
      }
      frame.src = "/log?embed=1&t=" + encodeURIComponent(window.SHELL.token) +
                  (id ? "&id=" + encodeURIComponent(id) : "");
      if (!dlg.open) { try { dlg.showModal(); } catch (e) { dlg.setAttribute("open", ""); } }
    }
    btn.addEventListener("click", function () { open(""); });
    if (close) close.addEventListener("click", function () { dlg.close(); });
    window.addEventListener("message", function (e) {
      if (e.origin !== location.origin) return;
      var d = e.data || {};
      if (d.type === "cpt-open-log") open(String(d.id || ""));
    });
  })();

  // ---------------------------------------------------------- 操作説明書(統合 1.2.0)
  // 上の帯の「説明書」。**いま見ているタブのツールの説明書**(写真入り)を開く。
  // 中身は /manual/<機能>(ダウンロードにしない。このまま読む)。説明書どうしは中で行き来できる。
  // 「別の窓で開く」は、いまダイアログに出ている説明書を別の窓(デスクトップ版)・別のタブ
  // (ブラウザ版)で開き直す。画面と並べて読むため。
  (function () {
    var btn = $("manualBtn"), dlg = $("manualDialog"), body = $("manualBody");
    var pop = $("manualPop"), close = $("manualClose");
    if (!btn || !dlg || !body) return;
    var frame = null;
    function currentKey() {
      var pane = panes.filter(function (p) { return !p.hidden; })[0];
      return pane ? pane.id.replace(/^pane-/, "") : "";
    }
    function open() {
      if (!frame) {
        frame = document.createElement("iframe");
        frame.className = "logdlg__frame";
        frame.title = "操作説明書";
        body.appendChild(frame);
      }
      frame.src = "/manual/" + encodeURIComponent(currentKey()) + "?embed=1";
      if (!dlg.open) { try { dlg.showModal(); } catch (e) { dlg.setAttribute("open", ""); } }
    }
    function shownPath() {
      // ダイアログの中で別の説明書へ移っていれば、そちらを開く(同じ origin なので読める)
      try {
        var loc = frame.contentWindow.location;
        if (loc.pathname.indexOf("/manual/") === 0) return loc.pathname + loc.hash;
      } catch (e) { /* 読めなければ、いまのタブの説明書 */ }
      return "/manual/" + encodeURIComponent(currentKey());
    }
    btn.addEventListener("click", open);
    if (close) close.addEventListener("click", function () { dlg.close(); });
    if (pop) pop.addEventListener("click", function () {
      var path = frame ? shownPath() : "/manual/" + encodeURIComponent(currentKey());
      window.open(path, "_blank");
      dlg.close();
    });
  })();

  // ---------------------------------------------------------- 印刷(Ctrl+P)
  // 3機能の画面は iframe の中にある。ブラウザの Ctrl+P は**一番外の文書**
  // (この外枠)を刷るので、そのままでは見出しとタブと、iframe の見えている
  // 部分だけが紙に出る。**Ctrl+P は、いま見せている機能の画面を刷る**
  // (移植元で Ctrl+P を押したときと同じ紙面)。機能の画面の中で押しても同じ。
  // 各機能の「印刷」ボタン(帳票を別のタブに出す・その画面で window.print)は
  // これまでどおりで、ここは関わらない。
  function currentFrame() {
    var pane = panes.filter(function (p) { return !p.hidden; })[0];
    return pane ? pane.querySelector("iframe") : null;
  }
  function isPrintKey(e) {
    return (e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey &&
           (e.key === "p" || e.key === "P");
  }
  function printCurrent(e) {
    if (!isPrintKey(e)) return;
    var frame = currentFrame();
    var win = frame && frame.contentWindow;
    if (!win) return;                          // 画面が無ければブラウザに任せる
    e.preventDefault();
    e.stopPropagation();
    try { win.focus(); win.print(); } catch (err) { window.print(); }
  }
  document.addEventListener("keydown", printCurrent, true);
  // 機能の画面の中で押されたときも拾う(同じアプリなので中の文書に触れる)。
  // 画面を移る(iframe の中で別のページを開く)たびに文書が替わるので、読み込むたびに付け直す
  panes.forEach(function (p) {
    var frame = p.querySelector("iframe");
    if (!frame) return;
    function attach() {
      try {
        frame.contentWindow.document.addEventListener("keydown", printCurrent, true);
        // 機能の画面で刷る(「このまま印刷」・Ctrl+P)あいだも心拍が止まる。外枠の合図で知らせる
        frame.contentWindow.addEventListener("beforeprint", function () {
          if (window.cptPrinting) window.cptPrinting();
        });
        frame.contentWindow.addEventListener("afterprint", function () {
          if (window.cptPrinted) window.cptPrinted();
        });
        if (window.cptWrapDialogs) window.cptWrapDialogs(frame.contentWindow);
      }
      catch (err) { /* 触れない文書(起きないはず)はブラウザに任せる */ }
    }
    frame.addEventListener("load", attach);
    attach();
  });

  // ---------------------------------------------------------- 生存確認
  var misses = 0;
  var restarted = false;
  var stopped = false;

  function setOffline(on) { var el = $("offline"); if (el) el.hidden = !on || stopped; }
  function setConn(state, text) {
    var el = $("conn");
    if (!el) return;
    el.textContent = text;
    el.className = "conn " + state;
  }

  function checkRestart(pid) {
    if (!pid || !S.serverPid || pid === S.serverPid || restarted) return restarted;
    restarted = true;
    setOffline(false);
    var box = $("restarted");
    if (box) box.hidden = false;
    var reload = $("btnReload");
    if (reload) reload.addEventListener("click", function () { location.reload(); }, { once: true });
    return true;
  }

  function health() {
    return fetch("/api/health", { cache: "no-store" })
      .then(function (r) { if (!r.ok) throw new Error(String(r.status)); return r.json(); })
      .then(function (j) {
        // 止めた直後は、まだ答えが返ることがある(止めるのは少し遅れる)。
        // 「終了しました」を「接続OK」で上書きしない
        if (stopped) return true;
        if (checkRestart(j.pid)) return true;
        misses = 0;
        setOffline(false);
        setConn("ok", "接続OK");
        return true;
      })
      .catch(function () {
        if (stopped) return false;
        if (++misses >= MISSES_BEFORE_OFFLINE) { setOffline(true); setConn("ng", "接続なし"); }
        return false;
      });
  }

  // ---------------------------------------------------------- 心拍
  function visibility() { return document.visibilityState === "hidden" ? "hidden" : "visible"; }

  // この画面の名乗り。**開くたびに新しい名前**にする(タブを複製すると
  // sessionStorage も写るので、そこには置かない)。見張りは名乗りごとに
  // 生き死にを持つので、別のタブ(ペナラベルの印刷ビュー)が開いていれば
  // この画面を閉じてもサーバは終わらない
  var CLIENT = "shell-" + Date.now().toString(36) + "-" +
               Math.random().toString(36).slice(2, 10);

  function payload(extra) {
    var body = { state: visibility(), client: CLIENT };
    // 外からの停止のとき、入口が「閉じる前の頼み」を出す相手(ブラウザ版だけ。統合 1.2.5)
    if (!DESKTOP) body.page = CLIENT;
    Object.keys(extra || {}).forEach(function (k) { body[k] = extra[k]; });
    return JSON.stringify(body);
  }

  function beacon(extra) {
    var body = new Blob([payload(extra)], { type: "application/json" });
    if (navigator.sendBeacon && navigator.sendBeacon("/api/alive", body)) return;
    fetch("/api/alive", { method: "POST", keepalive: true, cache: "no-store",
                          headers: { "Content-Type": "application/json" },
                          body: payload(extra) }).catch(function () {});
  }

  function alive(extra) {
    return fetch("/api/alive", { method: "POST", cache: "no-store",
                                 headers: { "Content-Type": "application/json" },
                                 body: payload(extra) })
      .then(function (r) { return r.json(); })
      .then(function (j) { checkRestart(j && j.pid); return j; })
      .catch(function () { return null; });
  }

  var resuming = 0;
  function resume(reason) {
    var now = Date.now();
    if (now - resuming < 1000) return;
    resuming = now;
    alive({ reason: reason }).then(function () { health(); });
  }

  document.addEventListener("visibilitychange", function () {
    if (visibility() === "hidden") beacon({ reason: "hidden" });
    else resume("foreground");
  });

  // 印刷の画面が開いているあいだは、この画面の JavaScript が止まり心拍が途切れる。
  // 用紙を替えに行くなどで 90 秒を超えると、誰も見ていないとみなされて終わっていた
  // (統合 1.2.20)。開く前に「裏に回った」と同じ合図を送り、閉じたら戻る。
  // 3機能の画面(iframe)の中で刷ったときも、その画面の合図を拾う(下の attach)
  function printing() { beacon({ reason: "print", state: "hidden" }); }
  // 済んだら**必ず**「前」を送る(resume は1秒以内の重なりをまとめるので、直前にタブを
  // 押していると送られず、次の心拍まで「裏」のままになる)
  function printed() { alive({ reason: "afterprint" }).then(function () { health(); }); }
  window.addEventListener("beforeprint", printing);
  window.addEventListener("afterprint", printed);
  window.cptPrinting = printing;
  window.cptPrinted = printed;

  // 確認の窓(confirm / alert / prompt)が出ているあいだも、この画面の JavaScript は止まり
  // 心拍が途切れる。押さずに 90 秒を超えると、誰も見ていないとみなされて終わっていた
  // (統合 1.2.56。実ブラウザで再現)。印刷と同じく、出す前に「裏に回った」と同じ合図を送り、
  // 閉じたら戻す。統合画面・3機能の画面(iframe)・ログの画面の窓をくるむ(同じアプリなので触れる)
  function wrapDialogs(win) {
    try {
      if (!win || win.__cptDialogs) return;
      win.__cptDialogs = true;
      ["confirm", "alert", "prompt"].forEach(function (name) {
        var original = win[name];
        if (typeof original !== "function") return;
        win[name] = function () {
          beacon({ reason: "dialog", state: "hidden" });
          try { return original.apply(win, arguments); }
          finally { alive({ reason: "dialog-closed" }); }
        };
      });
    } catch (e) { /* 触れない文書はくるまない */ }
  }
  window.cptWrapDialogs = wrapDialogs;
  wrapDialogs(window);
  panes.forEach(function (p) {                 // すでに読み込み済みの画面(読み込むたびにも付け直す)
    var f = p.querySelector("iframe");
    if (f) wrapDialogs(f.contentWindow);
  });
  document.addEventListener("freeze", function () { beacon({ reason: "freeze" }); });
  document.addEventListener("resume", function () { resume("resume"); });
  window.addEventListener("focus", function () { resume("focus"); });
  window.addEventListener("online", function () { resume("online"); });
  window.addEventListener("pagehide", function (event) {
    if (stopped) return;
    if (event.persisted) beacon({ reason: "hidden" });
    else beacon({ leaving: true, reason: "close" });
  });
  window.addEventListener("pageshow", function (event) {
    if (event.persisted) resume("restore");
  });

  // スリープからの戻りは時計の飛びで気づく
  var lastTick = Date.now();
  setInterval(function () {
    var now = Date.now();
    var jumped = now - lastTick > CLOCK_TICK_MS + CLOCK_JUMP_MS;
    lastTick = now;
    if (jumped) resume("sleep");
  }, CLOCK_TICK_MS);

  // ---------------------------------------------------------- 閉じる前(統合 1.2.5)
  // 3機能の画面(と開いているログ)に「まだ保存していない入力」を訊く。各画面は
  // `window.cptUnsaved()` で名前を返す(同じ origin なので直接呼べる)。あれば1つの確認に
  // まとめ、「閉じない」ならその画面を見せて止めない。all-tools の prepareFrames と同じ役目。
  var discarded = null;     // 「閉じてよい」と答えた中身と時刻(続けて2度は訊かない)
  function askFrame(frame) {
    try {
      var fn = frame && frame.contentWindow && frame.contentWindow.cptUnsaved;
      return typeof fn === "function" ? (fn() || []).map(String) : [];
    } catch (e) {
      return [];            // 読めない画面(読み込み中・別の origin)は数えない
    }
  }
  function prepareFrames() {
    var names = [], firstKey = "";
    panes.forEach(function (p) {
      var found = askFrame(p.querySelector("iframe"));
      if (found.length && !firstKey) firstKey = p.id.replace(/^pane-/, "");
      names = names.concat(found);
    });
    var logDlg = $("logDialog");
    if (logDlg && logDlg.open) names = names.concat(askFrame(logDlg.querySelector("iframe")));
    if (!names.length) return true;
    var stamp = JSON.stringify(names);
    if (discarded && discarded.stamp === stamp && Date.now() - discarded.at < 120000) return true;
    if (!window.confirm("まだ保存していない入力があります:\n\n" +
                        names.map(function (n) { return "・" + n; }).join("\n") +
                        "\n\n閉じると、この入力は消えます。閉じますか?")) {
      if (firstKey) show(firstKey);
      return false;
    }
    discarded = { stamp: stamp, at: Date.now() };
    return true;
  }

  function post(path, body) {
    return fetch(path, { method: "POST", cache: "no-store",
                         headers: { "Content-Type": "application/json", "X-Tool-Token": TOKEN },
                         body: JSON.stringify(body || {}) })
      .then(function (r) {
        return r.json().catch(function () { return {}; })
          .then(function (j) { return { status: r.status, body: j || {} }; });
      });
  }

  // ---------------------------------------------------------- 終了
  function markStopped(text) {
    stopped = true;
    var label = $("stoppedText");
    if (label && text) label.textContent = text;
    var band = $("rsBanner"); if (band) band.hidden = true;
    setConn("", "終了しました");
    var box = $("stopped"); if (box) box.hidden = false;
    panes.forEach(function (p) { p.hidden = true; });
    if (quit) quit.disabled = true;
  }
  // 機能の画面の「終了」(資材計算の帯)で止まったときも、同じ表示にする。
  // 知らせが無いと、外枠は「バックエンドと通信できません」を出していた
  window.addEventListener("message", function (e) {
    if (e.origin !== location.origin || !e.data) return;
    if (e.data.type === "cpt:stopped") markStopped("");
    // 機能の画面の「終了」(資材計算の帯)は、統合画面の「終了」と同じ流れにする(統合 1.2.5)
    if (e.data.type === "cpt:quit") requestQuit();
  });

  /**
   * 統合画面の「終了」。窓の × (デスクトップ版)もここを通る。
   *
   * 「終了します。よろしいですか」は訊かない(統合 1.2.5。all-tools と同じ)。訊くのは
   *   ・保存していない入力があるとき(閉じると消える)
   *   ・取り込みなど途中の処理があるとき(「中断して終了しますか」→ はい なら中断して止める)
   * だけ。ランチャーが窓を閉じる(× と同じ)ときに、毎回確認で止まらないため。
   */
  function requestQuit() {
    if (!quit || quit.disabled || stopped) return;
    if (!prepareFrames()) return;
    quit.disabled = true;
    // 打ちかけはいま確かめた。入口に「画面へ訊く」をもう一度させない
    post("/api/shutdown", { screens_ready: true })
      .then(function (res) {
        if (res.body.stopped) { markStopped(""); return; }
        if (res.status === 409 && res.body.reason === "busy") {
          var text = (res.body.message || "実行中の処理があります") + "\n\n中断して終了しますか?";
          if (!window.confirm(text)) { quit.disabled = false; return; }
          return post("/api/shutdown", { force: true }).then(function (again) {
            if (again.body.stopped) { markStopped(""); return; }
            quit.disabled = false;
            window.alert(again.body.message || ("終了できませんでした (HTTP " + again.status + ")"));
          });
        }
        quit.disabled = false;
        window.alert(res.body.message || ("終了できませんでした (HTTP " + res.status + ")"));
      })
      .catch(function () { quit.disabled = false; window.alert("終了の要求を送れませんでした。"); });
  }

  var quit = $("quit");
  if (quit) quit.addEventListener("click", requestQuit);

  // ---------------------------------------------------------- 外からの「閉じて」(ブラウザ版)
  // ランチャー・stop.bat が止めるとき(`/api/shutdown`)、入口は開いている画面に
  // 「閉じる前の頼み」を出して少し待つ(`/api/close-ask`)。画面は毎秒たずね、頼まれたら
  // 先に「確かめています」(`working`)を返してから、「終了」と同じく保存していない入力を訊く。
  // 閉じてよければ `ok`(入口が止まる)、「閉じない」なら `refused`(入口は止まらない)。
  // デスクトップ版は外から止める口が窓の × だけなので、たずねない。
  var closeAsking = 0, closeAnswered = 0;
  function watchCloseAsk() {
    if (DESKTOP || stopped || closeAsking) return;
    fetch("/api/close-ask?page=" + encodeURIComponent(CLIENT),
          { cache: "no-store", headers: { "X-Tool-Token": TOKEN } })
      .then(function (r) { return r.ok ? r.json() : {}; })
      .then(function (j) {
        var seq = Number((j && j.seq) || 0);
        if (!seq || seq === closeAnswered || closeAsking || stopped) return;
        closeAsking = seq;
        return post("/api/close-answer", { page: CLIENT, seq: seq, state: "working" })
          .catch(function () { return null; })
          .then(function () {
            var ok = true;
            try { ok = prepareFrames(); } catch (e) { ok = true; }
            return post("/api/close-answer", { page: CLIENT, seq: seq, state: ok ? "ok" : "refused" });
          })
          .then(function (res) {
            closeAnswered = seq;
            if (res && res.body && res.body.stopping) markStopped("");
          });
      })
      .catch(function () { /* 入口に届かない。次の回に */ })
      .then(function () { closeAsking = 0; });
  }

  // ---------------------------------------------------------- 管理者以外停止(統合 1.2.67)
  // 5秒ごとに入口へたずねる(入口はメモリだけを見る。共有を読むのは入口の見張り)。お願いが出ていれば
  // 帯を出し、3機能の画面に「保存していない入力」を訊いて答える(何も無ければ入口がまもなく止める)。
  // 答えには、強制で止まったときの控えに使う「打っていた値」(保存していない入力がある画面の、変えた欄だけ。
  // パスワードは持たない)を付ける。控えはこの PC の手元に残り、次の起動で知らせる。
  var RS_POLL_MS = 5000, RS_MISSES = 3;   // 3回(15秒)続けて届かなければ止まったとみなす
  var rsActive = null, rsDeadline = 0, rsMiss = 0, rsNotesShown = [];
  function changedField(el) {
    var type = (el.type || "").toLowerCase();
    if (el.readOnly || el.disabled) return false;
    if (["password", "hidden", "button", "submit", "reset", "file", "image"].indexOf(type) >= 0) return false;
    if (type === "checkbox" || type === "radio") return el.checked !== el.defaultChecked;
    if (el.tagName === "SELECT") {
      if (el.multiple) return false;
      var def = Array.prototype.findIndex.call(el.options, function (o) { return o.defaultSelected; });
      return el.selectedIndex !== (def < 0 ? 0 : def);
    }
    return el.value !== el.defaultValue;
  }
  function fieldLabel(el) {
    var text = "";
    try { if (el.labels && el.labels[0]) text = el.labels[0].textContent; } catch (e) { text = ""; }
    text = (text || el.getAttribute("aria-label") || el.getAttribute("title") || el.placeholder ||
            el.name || el.id || "").replace(/\s+/g, " ").trim();
    return text.slice(0, 60);
  }
  function fieldValue(el) {
    var type = (el.type || "").toLowerCase();
    if (type === "checkbox" || type === "radio") return el.checked ? "入" : "切";
    if (el.tagName === "SELECT") {
      var opt = el.options[el.selectedIndex];
      return opt ? (opt.textContent || opt.value).trim() : "";
    }
    return String(el.value || "");
  }
  function screenFields(frame, screen) {
    var out = [];
    try {
      var doc = frame && frame.contentDocument;
      if (!doc) return out;
      Array.prototype.forEach.call(doc.querySelectorAll("input, select, textarea"), function (el) {
        if (out.length < 80 && changedField(el)) {
          out.push({ screen: screen, label: fieldLabel(el), value: fieldValue(el).slice(0, 300) });
        }
      });
    } catch (e) { /* 読めない画面は数えない */ }
    return out;
  }
  function screensNow() {
    var labels = [], fields = [];
    panes.forEach(function (p) {
      var frame = p.querySelector("iframe");
      var found = askFrame(frame);
      if (!found.length) return;
      labels = labels.concat(found);
      var tab = $("tab-" + p.id.replace(/^pane-/, ""));
      var name = tab ? (tab.querySelector(".tab__label") || tab).textContent.trim() : "";
      fields = fields.concat(screenFields(frame, name));
    });
    var logDlg = $("logDialog");
    if (logDlg && logDlg.open) labels = labels.concat(askFrame(logDlg.querySelector("iframe")));
    return { labels: labels, fields: fields.slice(0, 200) };
  }
  function rsShowBand(req, localLabels) {
    var band = $("rsBanner");
    if (!band) return;
    if (!req || stopped) { band.hidden = true; return; }
    band.hidden = false;
    $("rsBy").textContent = req.text || "";
    var left = (req.left && req.left.length) ? req.left : localLabels;
    $("rsLeft").textContent = left.length
      ? "残っているもの: " + left.join("、") + "。保存・送信してから「確かめて終了」を押してください(残りの時間が来ると強制で止まります)"
      : "保存していない入力・処理中はありません。まもなく自動で終了します";
    rsDeadline = Date.now() + Math.max(0, Number(req.deadline_ms || 0) - Number(req.now_ms || 0));
    rsCountdown();
  }
  function rsCountdown() {
    var el = $("rsCount");
    if (!el || !rsActive) return;
    var sec = Math.max(0, Math.round((rsDeadline - Date.now()) / 1000));
    el.textContent = Math.floor(sec / 60) + ":" + ("0" + (sec % 60)).slice(-2);
  }
  function rsShowNotes(notes) {
    var box = $("rsNotes"), list = $("rsNotesList");
    if (!box || !list) return;
    notes = notes || [];
    if (JSON.stringify(notes) === JSON.stringify(rsNotesShown) && !box.hidden === !!notes.length) return;
    rsNotesShown = notes.slice();
    list.innerHTML = "";
    notes.forEach(function (n) { var li = document.createElement("li"); li.textContent = n; list.appendChild(li); });
    box.hidden = !notes.length || stopped;
  }
  function rsPoll() {
    if (stopped) return;
    fetch("/api/remote-stop/state?page=" + encodeURIComponent(CLIENT),
          { cache: "no-store", headers: { "X-Tool-Token": TOKEN } })
      .then(function (r) { if (!r.ok) throw new Error(String(r.status)); return r.json(); })
      .then(function (j) {
        rsMiss = 0;
        rsShowNotes(j.notes);
        var req = j.request || null;
        if (req && req.stopping) {
          markStopped("管理者の停止で終了しました。更新が終わってから開き直してください。");
          return;
        }
        rsActive = req;
        if (!req) { rsShowBand(null, []); return; }
        var now = screensNow();
        rsShowBand(req, now.labels);
        return post("/api/remote-stop/report", { page: CLIENT, labels: now.labels, fields: now.fields })
          .then(function (res) { if (res.body && res.body.request) rsShowBand(res.body.request, now.labels); })
          .catch(function () {});
      })
      .catch(function () {
        // お願いを受けていたあとで入口に届かなくなった = 止まった
        rsMiss += 1;
        if (rsActive && rsMiss >= RS_MISSES) {
          markStopped("管理者の停止で終了しました。更新が終わってから開き直してください。");
        }
      });
  }
  (function () {
    var quitBtn = $("rsQuit");
    if (quitBtn) quitBtn.addEventListener("click", requestQuit);
    var close = $("rsNotesClose");
    if (close) close.addEventListener("click", function () {
      var shown = rsNotesShown.slice();
      $("rsNotes").hidden = true;
      post("/api/remote-stop/notes-seen", { notes: shown })
        .then(function (res) { rsShowNotes((res.body && res.body.notes) || []); })
        .catch(function () {});
    });
  })();

  // 管理者の画面(上の帯の「管理者停止」)。開いているあいだだけ、5秒ごとに各 PC の様子を読む
  (function () {
    var btn = $("rsBtn"), dlg = $("rsDialog");
    if (!btn || !dlg) return;
    var timer = 0;
    function msg(text, ng) {
      var el = $("rsMsg"); if (!el) return;
      el.textContent = text || ""; el.className = "rsdlg__msg" + (ng ? " ng" : "");
    }
    function cell(tr, text) { var td = document.createElement("td"); td.textContent = text || ""; tr.appendChild(td); }
    function render(j) {
      var st = $("rsStatus");
      st.className = "rsdlg__status";
      if (!j.available) {
        st.textContent = j.why || "共有の梱包資材マスタのフォルダに届きません";
      } else if (j.unreadable) {
        st.textContent = j.unreadable + "(読めるようになるまで、出す・解除するはできません)";
      } else if (j.request) {
        st.className = "rsdlg__status on";
        st.textContent = "お願いが出ています: " + (j.request_text || "");
      } else {
        st.textContent = "お願いは出ていません(この端末: " + (j.me_pc || "?") + " / " + (j.me_login || "?") + ")";
      }
      var body = $("rsPcs");
      body.innerHTML = "";
      (j.pcs || []).forEach(function (pc) {
        var tr = document.createElement("tr");
        if (pc.stale) tr.className = "stale";
        cell(tr, pc.pc); cell(tr, pc.login); cell(tr, pc.label); cell(tr, pc.detail); cell(tr, pc.beat);
        body.appendChild(tr);
      });
      if (!(j.pcs || []).length) {
        var tr = document.createElement("tr"); var td = document.createElement("td");
        td.colSpan = 5; td.textContent = "動いている PC の知らせはありません"; tr.appendChild(td); body.appendChild(tr);
      }
    }
    function load() {
      fetch("/api/remote-stop/overview", { cache: "no-store", headers: { "X-Tool-Token": TOKEN } })
        .then(function (r) { return r.json(); })
        .then(render)
        .catch(function () { $("rsStatus").textContent = "様子を読めませんでした"; });
    }
    function act(path, body, okText) {
      var pw = $("rsPassword");
      body.password = pw ? pw.value : "";
      msg("確かめています…", false);
      return post(path, body).then(function (res) {
        if (pw) pw.value = "";
        msg(res.body.message || (res.body.ok ? okText : "できませんでした (HTTP " + res.status + ")"), !res.body.ok);
        load();
      }).catch(function () { msg("入口に届きませんでした", true); });
    }
    btn.addEventListener("click", function () {
      msg("", false);
      if (!dlg.open) { try { dlg.showModal(); } catch (e) { dlg.setAttribute("open", ""); } }
      load();
      clearInterval(timer);
      timer = setInterval(function () { if (dlg.open) load(); else clearInterval(timer); }, 5000);
    });
    $("rsClose").addEventListener("click", function () { dlg.close(); });
    dlg.addEventListener("close", function () { clearInterval(timer); });
    $("rsIssue").addEventListener("click", function () {
      if (!window.confirm("ほかの PC のコイル梱包ツールを止めます。\n" +
                          "保存していない入力が残っている PC も、5分たつと強制で止まります。よろしいですか?")) return;
      act("/api/remote-stop/issue", { reason: ($("rsReason").value || "").trim() }, "お願いを出しました");
    });
    $("rsClear").addEventListener("click", function () {
      act("/api/remote-stop/clear", {}, "解除しました");
    });
  })();

  // ---------------------------------------------------------- 開始
  health();
  setInterval(health, S.healthPollMs || 15000);
  // 同じタブの前の名乗り。ブラウザが裏のタブを捨てて読み直したとき、前の名乗りは
  // 閉じた合図を出せないまま「裏に回ったまま」で残る(全部閉じても終わらなくなる)。
  // 開いたときに言っておけば、サーバが忘れる(`IdleWatch.forget`)
  var PREV_KEY = "cpt.shell.client", previous = "";
  try { previous = sessionStorage.getItem(PREV_KEY) || ""; sessionStorage.setItem(PREV_KEY, CLIENT); }
  catch (e) { previous = ""; }
  alive({ reason: "open", replaces: previous });
  setInterval(function () { alive({ reason: "timer" }); }, S.alivePollMs || 20000);
  if (!DESKTOP) setInterval(watchCloseAsk, 1000);
  rsPoll();
  setInterval(rsPoll, RS_POLL_MS);
  setInterval(rsCountdown, 1000);
})();
