/* 梱包ペナ ラベル・風袋計算 — 画面スクリプト
   追加ライブラリなし（素の DOM API のみ）。 */
(function () {
  "use strict";

  // ---------------------------------------------------------- 入口とトークン(統合版)
  // この機能の入口(統合版では `/pena`)と起動トークンは、外枠(_layout.html)の
  // body に載っている。**経路はここで1度だけ付け、トークンはヘッダに載せる**
  var BASE = (document.body && document.body.getAttribute("data-base")) || "";
  var TOKEN = (document.body && document.body.getAttribute("data-token")) || "";
  function url(path) { return BASE + path; }
  function headers(extra) {
    var h = { "X-Tool-Token": TOKEN };
    Object.keys(extra || {}).forEach(function (k) { h[k] = extra[k]; });
    return h;
  }
  window.pplUrl = url;
  window.pplHeaders = headers;

  // ---------------------------------------------------------- 共通
  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

  var toastTimer = null;
  function hideToast() {
    var el = $("#toast");
    if (el) { el.hidden = true; }
    if (toastTimer) { clearTimeout(toastTimer); toastTimer = null; }
  }
  function toast(msg, level) {
    var el = $("#toast");
    if (!el) { return; }
    if (!msg) { hideToast(); return; }
    // 本文と閉じるボタンを分けて入れる（本文は textContent で安全に）
    el.textContent = "";
    var body = document.createElement("span");
    body.className = "toast-body";
    body.textContent = msg;
    var x = document.createElement("button");
    x.type = "button";
    x.className = "toast-x";
    x.title = "消す";
    x.textContent = "×";
    el.appendChild(body);
    el.appendChild(x);
    el.className = "toast no-print " + (level || "info");
    el.hidden = false;
    if (toastTimer) { clearTimeout(toastTimer); toastTimer = null; }
    // 止めた知らせ(エラー)は、読んで閉じるまで出したままにする(統合 1.2.31。資材マスタに届かず
    // 計算を止めたとき、理由と確かめることが 7 秒で消えて分からなくなるため)
    if (level !== "error") { toastTimer = setTimeout(hideToast, 7000); }
  }
  window.pplToast = toast;

  // どこを押しても消せる（待たずに片付けたいときのため）
  (function () {
    var el = $("#toast");
    if (el) { el.addEventListener("click", hideToast); }
  })();

  // 操作は**押した順に1つずつ**送る(統合 1.2.20)。続けて押すと要求が同時に出て、返事の順が
  // 入れ替わると、あとの結果を前の返事の画面で上書きする(資材計算で再現した形)。
  // 画面の持ち主の知らせ(/api/screen/…)は待たせない(心拍が遅れると使用権を失う)
  var chain = Promise.resolve();
  function post(path, body) {
    if (String(path).indexOf("/api/screen/") === 0) { return send(path, body); }
    var run = chain.then(function () { return send(path, body); });
    chain = run.catch(function () { /* 失敗しても次は送る */ });
    return run;
  }

  function send(path, body) {
    var payload = body || {};
    // どの画面からの操作かを必ず名乗る（入力画面は同時に1つだけ）
    if (screenId && !payload.screenId) { payload.screenId = screenId; }
    return fetch(url(path), {
      method: "POST",
      headers: headers({ "Content-Type": "application/json" }),
      body: JSON.stringify(payload)
    }).then(function (r) {
      return r.json().catch(function () {
        throw new Error("サーバー応答を解釈できません (HTTP " + r.status + ")");
      });
    }).then(function (j) {
      // 使用権を失っていた（別の画面で「この画面で使う」が押された）
      if (j && j.screenTaken) { showLost(); }
      return j;
    });
  }
  window.pplPost = post;

  // ---------------------------------------------------------- その場で表示を新しくする
  // 操作のあとは**ページを読み直さずに**状態を出す(統合 1.2.11。現場の指摘: どのツールでも
  // リロードを挟まずに反映させる)。同じページを裏で取り直し、`data-live` の付いた要素だけを
  // id で対応させて差し替える。打ちかけの入力・スクロール位置・開いている面はそのまま。
  // `data-live` の中のボタンは document で受ける(差し替えても効くように)。
  //   opts.values: 欄(data-key)の値も新しくする key の配列。"all" ならすべて
  function refresh(opts) {
    opts = opts || {};
    return fetch(location.href, { cache: "no-store", headers: headers() }).then(function (r) {
      if (!r.ok) { throw new Error("画面を取り直せません (HTTP " + r.status + ")"); }
      return r.text();
    }).then(function (html) {
      var doc = new DOMParser().parseFromString(html, "text/html");
      $$("[data-live]").forEach(function (el) {
        var fresh = el.id ? doc.getElementById(el.id) : null;
        if (fresh) { el.replaceWith(document.importNode(fresh, true)); }
      });
      var keys = opts.values;
      if (keys) {
        $$("[data-key]").forEach(function (el) {
          var k = el.getAttribute("data-key");
          if (keys !== "all" && keys.indexOf(k) < 0) { return; }
          var fresh = doc.querySelector('[data-key="' + k + '"]');
          if (!fresh) { return; }
          // 保存した値が「描いたときの値」になる(閉じる前の「保存していない欄」に数えない)
          if (el.type === "checkbox") { el.checked = el.defaultChecked = fresh.hasAttribute("checked"); }
          else { el.value = el.defaultValue = fresh.getAttribute("value") || ""; }
        });
      }
      document.dispatchEvent(new CustomEvent("ppl:refreshed"));
      return doc;
    });
  }
  window.pplRefresh = refresh;

  // ---------------------------------------------------------- 合言葉を聞く
  // **打った文字を見せない**(●で出す)。以前はブラウザの `prompt()` で聞いていて、
  // 合言葉がそのまま画面に見えていた(現場の指摘)。`prompt()` は伏せ字にできない。
  // 戻り値は Promise: 入れた合言葉、やめたら null。
  function askPassword(message) {
    return new Promise(function (resolve) {
      var dlg = document.createElement("dialog");
      dlg.className = "pwdialog";
      var p = document.createElement("p");
      p.className = "pwdialog-msg";
      p.textContent = message || "合言葉を入力してください";
      var box = document.createElement("input");
      box.type = "password";
      box.autocomplete = "off";
      box.className = "pwdialog-box";
      box.setAttribute("aria-label", "合言葉");
      var row = document.createElement("div");
      row.className = "btn-row";
      var ok = document.createElement("button");
      ok.type = "button"; ok.className = "btn btn-primary"; ok.textContent = "OK";
      var cancel = document.createElement("button");
      cancel.type = "button"; cancel.className = "btn"; cancel.textContent = "やめる";
      row.appendChild(ok); row.appendChild(cancel);
      dlg.appendChild(p); dlg.appendChild(box); dlg.appendChild(row);
      document.body.appendChild(dlg);
      var done = false;
      function finish(value) {
        if (done) { return; }
        done = true;
        try { dlg.close(); } catch (e) { /* 閉じていても構わない */ }
        dlg.remove();
        resolve(value);
      }
      ok.addEventListener("click", function () { finish(box.value); });
      cancel.addEventListener("click", function () { finish(null); });
      dlg.addEventListener("cancel", function (ev) { ev.preventDefault(); finish(null); });
      box.addEventListener("keydown", function (ev) {
        if (ev.key === "Enter") { ev.preventDefault(); finish(box.value); }
      });
      dlg.showModal();
      box.focus();
    });
  }
  window.pplAskPassword = askPassword;

  // ---------------------------------------------------------- 確かめてから印刷用のタブを開く
  // 以前は `window.confirm` で確かめてから `window.open` していた。**確かめを読んで OK を
  // 押すまでに 5 秒ほどかかると、ブラウザのポップアップのブロックに止められて何も開かなかった**
  // (統合 1.2.20。ブロックを入れたままの Chromium で再現。押した操作から時間がたった
  // `window.open` はポップアップとみなされる)。そこで画面の中で確かめ、**OK を押したその
  // 操作で**開く(ブロックされない)。返り値は Promise: 開いたら true、やめたら false
  function confirmOpen(message, href, okLabel) {
    return new Promise(function (resolve) {
      var dlg = document.createElement("dialog");
      dlg.className = "pwdialog pplconfirm";
      var p = document.createElement("p");
      p.className = "pwdialog-msg";
      p.style.whiteSpace = "pre-line";
      p.textContent = message || "印刷用のタブを開きます。";
      var row = document.createElement("div");
      row.className = "btn-row";
      var ok = document.createElement("button");
      ok.type = "button"; ok.className = "btn btn-primary"; ok.setAttribute("data-ok", "");
      ok.textContent = okLabel || "OK";
      var cancel = document.createElement("button");
      cancel.type = "button"; cancel.className = "btn"; cancel.textContent = "キャンセル";
      row.appendChild(ok); row.appendChild(cancel);
      dlg.appendChild(p); dlg.appendChild(row);
      document.body.appendChild(dlg);
      var done = false;
      function finish(opened) {
        if (done) { return; }
        done = true;
        try { dlg.close(); } catch (e) { /* 閉じていても構わない */ }
        dlg.remove();
        resolve(opened);
      }
      ok.addEventListener("click", function () {
        window.open(href, "_blank", "noopener");     // この押した操作の中で開く
        finish(true);
      });
      cancel.addEventListener("click", function () { finish(false); });
      dlg.addEventListener("cancel", function (ev) { ev.preventDefault(); finish(false); });
      dlg.showModal();
      ok.focus();
    });
  }
  window.pplConfirmOpen = confirmOpen;

  // ---------------------------------------------------------- 印刷の記録(統合 1.2.31)
  // 印刷の画面(ブラウザの印刷ダイアログ)を閉じたら、どの画面を刷ったかをサーバへ知らせる。
  // サーバは「どの取得元の、いつ時点の資材重量で計算した紙か」と一緒に残す(/print-log で見る)。
  // ブラウザは刷ったか・やめたかを教えないので、記録は「印刷の画面を開いた」こと
  window.addEventListener("afterprint", function () {
    var path = location.pathname;
    if (BASE && path.indexOf(BASE) === 0) { path = path.slice(BASE.length); }
    send("/api/print-log", { path: path || "/", search: location.search })
      .catch(function () { /* 残せなくても印刷は止めない */ });
  });

  // ---------------------------------------------------------- 印刷用のタブ
  // 帯の「印刷」「このタブを閉じる」(`pages._printbar`)。閉じられないブラウザ・
  // 開き方のときは、× で閉じるよう案内する(本ツールのタブはそのまま残っている)
  document.addEventListener("click", function (e) {
    var t = e.target;
    if (!t || !t.closest) { return; }
    if (t.closest("[data-print-now]")) { window.print(); return; }
    if (t.closest("[data-close-tab]")) {
      window.close();
      setTimeout(function () {
        if (!window.closed) {
          toast("このタブはブラウザの × で閉じてください。本ツールのタブはそのまま開いています。", "info");
        }
      }, 300);
    }
  });

  // ---------------------------------------------------------- 画面の多重起動
  // 作業状態（検査番号・重量・本数・計算結果）はサーバー側に1組しか無い。
  // 入力できる画面を2枚開くと両方が同じ1組を書き換え、
  // 画面に出ている検査番号と刷られる検査番号が食い違う。
  // そのため入力画面は同時に1つだけにし、2枚目は開くこと自体を断る。
  var guardEl = $("#screenguard");
  var guarded = !!(guardEl && guardEl.getAttribute("data-guard") === "1");
  var screenId = "";
  var beatTimer = null;

  function newScreenId() {
    // タブ単位で一意ならよい。暗号用途ではない。
    var rnd = "";
    try {
      var a = new Uint8Array(8);
      (window.crypto || window.msCrypto).getRandomValues(a);
      for (var i = 0; i < a.length; i++) {
        rnd += ("0" + a[i].toString(16)).slice(-2);
      }
    } catch (e) {
      rnd = String(Math.random()).slice(2);
    }
    return "s" + Date.now().toString(36) + rnd;
  }

  // sessionStorage はタブごとに分かれる（再読込では同じ画面のまま）。
  // localStorage だとタブ間で共有されてしまうので使わない。
  try {
    screenId = sessionStorage.getItem("pplScreenId") || "";
    if (!screenId) {
      screenId = newScreenId();
      sessionStorage.setItem("pplScreenId", screenId);
    }
  } catch (e) {
    screenId = newScreenId();       // プライベートウィンドウなど
  }

  function sgShow(which) {
    if (!guardEl) { return; }
    ["sgWait", "sgDeny", "sgLost"].forEach(function (id) {
      var el = $("#" + id);
      if (el) { el.hidden = (id !== which); }
    });
    guardEl.hidden = false;
  }

  function stopBeat() {
    if (beatTimer) { clearInterval(beatTimer); beatTimer = null; }
  }

  function showLost() {
    if (!guarded) { return; }
    lost = true;
    holding = false;
    stopBeat();
    sgShow("sgLost");
  }

  // 表に出ているか。裏のタブ・最小化したウィンドウでは false。
  function isVisible() {
    return !(document.hidden || document.visibilityState === "hidden");
  }

  var lost = false;                   // 別の画面に取られた（戻さない）
  var holding = false;                // この画面が使用権を持っている

  // 生存通知・表裏の知らせの返事を見て、画面を合わせる
  function onBeatReply(j) {
    if (!j || j.granted !== false) { return; }
    if (j.heldByOther) {
      showLost();                     // 別の画面に取られた
    } else {
      // 誰も使っていない（期限切れ・サーバーの再起動）。
      // 警告を出さずに黙って取り直す。
      claim(false);
    }
  }

  function beat() {
    post("/api/screen/ping", { visible: isVisible() })
      .then(onBeatReply)
      .catch(function () { /* 接続監視の担当 */ });
  }

  function startBeat() {
    stopBeat();
    beatTimer = setInterval(beat, 5000);
  }

  function claim(force) {
    sgShow("sgWait");
    return post("/api/screen/claim", { force: !!force, visible: isVisible() }).then(function (j) {
      if (j && j.granted) {
        lost = false;
        holding = true;
        guardEl.hidden = true;
        startBeat();
        return true;
      }
      var since = $("#sgSince");
      if (since) { since.textContent = (j && j.heldSince) || "-"; }
      // 相手の画面が裏に回っているなら、そう伝える（タブを探してもらう）
      var bg = $("#sgBg"), bgSince = $("#sgBgSince");
      if (bg) { bg.hidden = !(j && j.holderHidden); }
      if (bgSince) { bgSince.textContent = (j && j.hiddenSince) || "-"; }
      holding = false;
      sgShow("sgDeny");
      return false;
    }).catch(function () {
      // サーバーへ届かないなら接続監視に任せる（画面は塞いだまま）
      sgShow("sgWait");
      return false;
    });
  }

  if (guarded) {
    ["sgTake", "sgBack"].forEach(function (id) {
      var b = $("#" + id);
      if (b) { b.addEventListener("click", function () { claim(true); }); }
    });
    var retry = $("#sgRetry");
    if (retry) { retry.addEventListener("click", function () { claim(false); }); }

    claim(false);

    // ページが止められる直前でも届く送り方（sendBeacon → keepalive fetch）
    function beacon(path, obj) {
      obj.screenId = screenId;
      var body = JSON.stringify(obj);
      try {
        if (navigator.sendBeacon &&
            navigator.sendBeacon(url(path), new Blob([body], { type: "application/json" }))) {
          return;
        }
      } catch (e) { /* 下の fetch へ */ }
      try {
        fetch(url(path), {
          method: "POST", body: body, keepalive: true,
          headers: { "Content-Type": "application/json" }
        });
      } catch (e) { /* 届かなくても生存通知に表裏を載せている */ }
    }

    // ---- 裏に回る ----
    // ブラウザーは裏のタブのタイマーを間引く（Chrome / Edge は 5 分ほどで
    // 1 分に 1 回まで。眠ったタブ・PC のスリープでは止まる）。生存通知が
    // 途切れても空きとみなさないよう、裏に回る瞬間にサーバーへ知らせる。
    // 断られている画面・取られた画面は知らせない（使用中の画面だけ）
    function goingAway(reason) {
      if (!holding) { return; }
      beacon("/api/screen/state", { visible: false, reason: reason });
    }

    // ---- 表に戻る・凍結やスリープから戻る ----
    // タイマーを待たずに、すぐ生存を知らせて使用権と接続を確かめる。
    var lastWake = 0;
    function wake(reason) {
      if (!holding || !isVisible()) { return; }
      var now = Date.now();
      if (now - lastWake < 1000) { return; }     // 同時に来るイベントをまとめる
      lastWake = now;
      post("/api/screen/state", { visible: true, reason: reason })
        .then(onBeatReply)
        .catch(function () { /* 接続監視の担当 */ });
      startBeat();                                // 次の通知を今から数え直す
      if (typeof ping === "function") { ping(); } // 接続表示も今すぐ更新
    }

    document.addEventListener("visibilitychange", function () {
      if (isVisible()) { wake("visible"); } else { goingAway("hidden"); }
    });
    // Page Lifecycle（Chrome / Edge）: 眠ったタブの凍結と復帰
    document.addEventListener("freeze", function () { goingAway("freeze"); });
    document.addEventListener("resume", function () { wake("resume"); });
    window.addEventListener("focus", function () { wake("focus"); });
    window.addEventListener("online", function () { wake("online"); });

    // 戻る／進むで保存状態（bfcache）から戻った: 閉じるときに手放しているので取り直す
    window.addEventListener("pageshow", function (e) {
      if (e.persisted && !lost) { claim(false); }
    });

    // ---- スリープ復帰の検知 ----
    // PC がスリープするとタイマーごと止まる。表に出ているのに時計が
    // 大きく飛んでいたら、スリープ（または凍結）から戻ったとみなす。
    var lastTick = Date.now();
    setInterval(function () {
      var now = Date.now();
      var gap = now - lastTick;
      lastTick = now;
      if (gap > 30000 && isVisible()) { wake("sleep"); }
    }, 5000);

    // 閉じるときは手放す。fetch は間に合わないので sendBeacon を使う。
    window.addEventListener("pagehide", function () {
      if (!holding) { return; }
      beacon("/api/screen/release", {});
    });
  }

  // ---------------------------------------------------------- 画面の切替
  // 表示だけの面（風袋計算・羅列・計算内容・ラベル台紙）は、
  // 別タブを開かずに起動ページの中で差し替える。
  // 印刷ビューだけは今までどおり別タブ（用紙を替える作業があるため）。
  (function () {
    var home = $("#paneHome");
    var view = $("#paneView");
    if (!home || !view) { return; }            // 起動ページ以外では何もしない

    var links = $$(".appbar-nav a[data-pane]");

    function mark(name) {
      links.forEach(function (a) {
        a.classList.toggle("on", a.getAttribute("data-pane") === name);
      });
    }

    function runScripts(root) {
      // innerHTML で入れた <script> は動かないので入れ直す
      $$("script", root).forEach(function (old) {
        var s = document.createElement("script");
        if (old.src) { s.src = old.src; } else { s.textContent = old.textContent; }
        old.parentNode.replaceChild(s, old);
      });
    }

    function showHome() {
      view.hidden = true;
      view.innerHTML = "";
      home.hidden = false;
      mark("home");
      history.replaceState(null, "", BASE + "/");
    }

    function showPane(name, href) {
      home.hidden = true;
      view.hidden = false;
      view.innerHTML = '<div class="card"><div class="empty">読み込んでいます…</div></div>';
      mark(name);
      fetch(href + (href.indexOf("?") < 0 ? "?" : "&") + "pane=1",
            { cache: "no-store" })
        .then(function (r) { return r.text(); })
        .then(function (html) {
          view.innerHTML = html;
          runScripts(view);
          history.replaceState(null, "", href);
          window.scrollTo(0, 0);
        })
        .catch(function (err) {
          view.innerHTML = '<div class="card"><div class="empty">'
            + "読み込めませんでした: " + String(err) + "</div></div>";
        });
    }

    links.forEach(function (a) {
      a.addEventListener("click", function (e) {
        if (e.metaKey || e.ctrlKey || e.shiftKey || e.button) { return; }
        e.preventDefault();
        var name = a.getAttribute("data-pane");
        if (name === "home") { showHome(); } else { showPane(name, a.getAttribute("href")); }
      });
    });
  })();

  // ---------------------------------------------------------- 接続監視
  // 基盤仕様書 2.9: ブラウザーとバックエンドの状態を分けて扱う（監視レベル1）
  var connEl = $("#conn");
  function ping() {
    if (!connEl) { return; }
    fetch(url("/api/health"), { cache: "no-store" })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (j && j.appId) {
          connEl.textContent = "接続OK";
          connEl.className = "conn ok";
        } else {
          throw new Error("別アプリ応答");
        }
      })
      .catch(function () {
        connEl.textContent = "接続なし — 再読込してください";
        connEl.className = "conn ng";
      });
  }
  ping();
  setInterval(ping, 15000);

  // ---------------------------------------------------------- 入力制限
  // VBA の KeyPress / Change 制限を再現
  function applyFilter(input) {
    var kind = input.getAttribute("data-filter");
    var v = input.value;
    var out = v;
    if (kind === "alnum") {
      // 0-9 A-Z a-z のみ。小文字は大文字へ（VBA KeyAscii - 32）
      out = v.replace(/[^0-9A-Za-z]/g, "").toUpperCase();
      // 桁の上限(検査番号は 7 桁。data-max)。超える入力は**入れない**(前の値に戻して知らせる)。
      // 以前は maxlength で、8 桁を貼り付けると黙って 7 桁に切られ、違う番号のまま進めた(統合 1.2.36)
      var max = parseInt(input.getAttribute("data-max") || "0", 10);
      if (max && out.length > max) {
        var prev = input.getAttribute("data-prev") || "";
        toast("検査番号は" + max + "桁までです。" + out.length + "桁の入力は入れていません（前の「"
              + prev + "」のまま）", "warn");
        out = prev;
      }
      if (max) { input.setAttribute("data-prev", out); }
    } else if (kind === "num") {
      // VBA: If Not IsNumeric(text) Then text = ""（"-" 単独は許容）
      if (v !== "-" && v !== "" && !/^[+-]?(\d+\.?\d*|\.\d+)$/.test(v)) {
        out = v.slice(0, -1);
        if (out !== "-" && out !== "" && !/^[+-]?(\d+\.?\d*|\.\d+)$/.test(out)) { out = ""; }
      }
    } else if (kind === "coil") {
      // CoilH1..4: 数字だけ(VBA の KeyPress は 0-9 しか通さない。「.」「-」は打てない。統合 1.2.35)。
      // 以前は「11.5」が打てて、サーバが黙って空にしていた
      if (v !== "" && !/^\d+$/.test(v)) {
        out = v.replace(/[^0-9]/g, "");
      }
    }
    if (out !== v) { input.value = out; }
  }
  // 桁の上限のある欄は、書き換わる直前の値を控える(超えたらこの値に戻す。統合 1.2.36)
  document.addEventListener("beforeinput", function (e) {
    if (e.target && e.target.hasAttribute && e.target.hasAttribute("data-max")) {
      e.target.setAttribute("data-prev", e.target.value);
    }
  });
  document.addEventListener("input", function (e) {
    if (e.target && e.target.hasAttribute && e.target.hasAttribute("data-filter")) {
      applyFilter(e.target);
    }
  });

  // ---------------------------------------------------------- 保存していない入力
  // 統合画面の「終了」・窓の × ・外からの停止(ランチャー・stop.bat)の前に、統合画面が
  // `window.cptUnsaved()` を呼んで訊く(統合 1.2.5)。名前を返せば「閉じると消えます。
  // 閉じますか?」が出る。`data-unsaved` の枠の中の欄は、描いたときの値(defaultValue)と比べる
  // (保存したあとは pplRefresh が描いたときの値も新しくするので、保存したものは数えない)。
  function changedField(el) {
    if (el.readOnly || el.disabled || el.type === "password" || el.type === "button") { return false; }
    if (el.type === "checkbox" || el.type === "radio") { return el.checked !== el.defaultChecked; }
    if (el.tagName === "SELECT") {
      if (el.multiple) { return false; }
      var def = Array.prototype.findIndex.call(el.options, function (o) { return o.defaultSelected; });
      return el.selectedIndex !== (def < 0 ? 0 : def);
    }
    return el.value !== el.defaultValue;
  }
  var unsavedChecks = [];
  window.cptUnsaved = function () {
    var out = [];
    $$("[data-unsaved]").forEach(function (box) {
      if (box.hidden || (box.parentElement && box.parentElement.closest("[hidden]"))) { return; }
      if ($$("input, select, textarea", box).some(changedField)) { out.push(box.getAttribute("data-unsaved")); }
    });
    unsavedChecks.forEach(function (check) {
      try { out = out.concat(check() || []); } catch (e) { /* 訊けなくても閉じられる */ }
    });
    return out.map(function (name) { return "ペナラベル: " + name; });
  };

  // ---------------------------------------------------------- メイン画面
  var form = $("#mainForm");
  if (!form) { return; }

  function collect() {
    var sel = $$("input[name=size]:checked")[0];
    var st = {
      selectedOb: 0, selectedCb: 0, namedCb: false,
      // 選んでいなければ空のまま送る(1000 で埋めない。サーバが「チップボールチェックがありません」で
      // 止める。既定の 1000 は画面を開いたときに選んである。統合 1.2.32・現場の指示)
      tip: ($$("input[name=tip]:checked")[0] || {}).value || "",
      kensaNo: $("#kensaNo").value,
      weight1: $("#weight1").value,
      weight2: $("#weight2").value,
      coilH: {}
    };
    if (sel) {
      if (sel.dataset.kind === "ob") {
        st.selectedOb = parseInt(sel.dataset.ob, 10);
      } else if (sel.dataset.kind === "cbNamed") {
        st.namedCb = true;
      } else {
        st.selectedCb = parseInt(sel.dataset.cb, 10);
      }
    }
    [1, 2, 3, 4].forEach(function (i) {
      var el = $("#coilH" + i);
      st.coilH[String(i)] = el ? el.value : "";
    });
    return st;
  }

  // 最後にサーバへ渡した入力(ボタン・サイズ選択のたびに `collect()` を渡し、答えで描き直す)
  var sentForm = null;
  function formKey() {
    var st = collect();
    return JSON.stringify([st.kensaNo, st.weight1, st.weight2, st.coilH, st.tip]);
  }
  unsavedChecks.push(function () {
    return (sentForm !== null && formKey() !== sentForm)
      ? ["新検番・重量・高さ・チップボール(ボタンを押していない入力)"] : [];
  });

  function paint(state) {
    if (!state) { return; }
    // 描き終えたら、いまの入力を「渡したもの」として覚える
    setTimeout(function () { sentForm = formKey(); }, 0);
    $("#lblKensaNo").textContent = state.lblKensaNo || "";
    $("#lblSize1").textContent = state.lblSize1 || "";
    $("#lblWeight1").textContent = state.lblWeight1 || "";
    $("#lblSize2").textContent = state.lblSize2 || "";
    $("#lblWeight2").textContent = state.lblWeight2 || "";
    [1, 2, 3, 4].forEach(function (i) {
      var nw = $("#nw" + i), gw = $("#gw" + i), ta = $("#ta" + i), ch = $("#coilH" + i);
      if (nw) { nw.textContent = (state.nw && state.nw[i]) || ""; }
      if (gw) { gw.textContent = (state.gw && state.gw[i]) || ""; }
      if (ta) { ta.textContent = (state.ta && state.ta[i]) || ""; }
      if (ch && state.coilH && typeof state.coilH[i] === "string") { ch.value = state.coilH[i]; }
    });
    if (typeof state.kensaNo === "string") { $("#kensaNo").value = state.kensaNo; }
    if (typeof state.lotNo === "string" && $("#lotNo") && document.activeElement !== $("#lotNo")) {
      $("#lotNo").value = state.lotNo;
    }
    // チップボール: サーバがコイル外径から選び直すことがある(統合 1.2.35)。選んだわけを下に出す
    if (typeof state.tip === "string") {
      $$("input[name=tip]").forEach(function (r) { r.checked = r.value === state.tip; });
    }
    var tipNote = $("#tipNote");
    if (tipNote && typeof state.tipNote === "string") { tipNote.textContent = state.tipNote; }
    applyTakeVisibility(state);
  }

  // VBA: Take1/Take2 の Visible と txtWeight1/2 の表示切替
  function applyTakeVisibility(state) {
    var t1 = state ? state.take1Visible : false;
    var t2 = state ? state.take2Visible : false;
    $("#take1").classList.toggle("off", !t1);
    $("#take2").classList.toggle("off", !t2);
    $("#w1wrap").style.opacity = t1 ? "" : ".4";
    $("#w2wrap").style.opacity = t2 ? "" : ".4";
    $("#weight1").disabled = !t1;
    $("#weight2").disabled = !t2;
    [1, 2].forEach(function (i) { var e = $("#coilH" + i); if (e) { e.disabled = !t1; } });
    [3, 4].forEach(function (i) { var e = $("#coilH" + i); if (e) { e.disabled = !t2; } });
  }

  function markSelectedRow() {
    $$(".size-row").forEach(function (row) {
      row.classList.toggle("on", !!$("input[name=size]:checked", row));
    });
  }

  // --- サイズ選択（VBA OptionButtonN_Click / CheckBoxN_Click）---
  form.addEventListener("change", function (e) {
    var t = e.target;
    if (t.name === "size") {
      markSelectedRow();
      var body;
      if (t.dataset.kind === "ob") {
        body = { obIdx: parseInt(t.dataset.ob, 10) };
      } else if (t.dataset.kind === "cbNamed") {
        body = { named: true };
      } else {
        body = { cbIdx: parseInt(t.dataset.cb, 10) };
      }
      body.current = collect();
      var url = (t.dataset.kind === "ob") ? "/api/select-size" : "/api/select-checkbox";
      post(url, body).then(function (j) {
        if (!j.ok) { toast(j.message, j.level || "error"); }
        paint(j.state);
      }).catch(function (err) { toast(String(err), "error"); });
    }
    if (t.id === "cbMode") {
      // 表示モードを切り替え、選択をいったん外す（VBA の排他と同じ考え方）
      document.body.classList.toggle("mode-cb", t.checked);
      $$("input[name=size]").forEach(function (r) {
        var isCb = (r.dataset.kind !== "ob");
        r.closest("label").hidden = (t.checked !== isCb);
        if (r.checked && (t.checked !== isCb)) { r.checked = false; }
      });
      $("#modeHint").textContent = t.checked
        ? "丈1・丈2 を同時に処理します（VBA の CheckBox モード）"
        : "丈1 または 丈2 を個別に処理します（VBA の OptionButton モード）";
      markSelectedRow();
    }
  });

  // --- 進み具合の棒(重量計算_DB) ---
  // 押してから答えが返るまで、サーバの進み具合(`/api/progress`)を 0.25 秒ごとに聞いて
  // 「何段目まで済んだか」を棒と文で出す(現場の指摘: 重いのはいいが進み具合を出すこと)。
  // すぐ終わるときに一瞬だけ出てちらつかないよう、0.3 秒たってから出す。
  function progressBar(title) {
    var box = null, timer = null, shown = null, stopped = false;
    function ensure() {
      if (box) { return box; }
      box = document.createElement("div");
      box.className = "pbar no-print";
      box.setAttribute("role", "progressbar");
      box.innerHTML = '<div class="pbar-card"><b class="pbar-title"></b>' +
        '<div class="pbar-track"><div class="pbar-fill"></div></div>' +
        '<div class="pbar-text"></div></div>';
      box.querySelector(".pbar-title").textContent = title + "中…";
      document.body.appendChild(box);
      return box;
    }
    function draw(p) {
      var el = ensure();
      var total = Math.max(1, p.total || 1), step = Math.min(p.step || 0, total);
      var pct = Math.round(step * 100 / total);
      el.querySelector(".pbar-fill").style.width = pct + "%";
      el.setAttribute("aria-valuenow", String(pct));
      el.querySelector(".pbar-text").textContent =
        (p.text || "始めています") + "(" + step + " / " + total + ")" +
        (p.elapsed ? "  " + p.elapsed + "秒" : "");
    }
    function poll() {
      if (stopped) { return; }
      fetch(url("/api/progress"), { cache: "no-store", headers: headers() })
        .then(function (r) { return r.json(); })
        .then(function (p) { if (!stopped && p && p.active) { draw(p); } })
        .catch(function () { /* 聞けなくても計算は続く */ })
        .then(function () { if (!stopped) { timer = setTimeout(poll, 250); } });
    }
    shown = setTimeout(function () { if (!stopped) { draw({ step: 0, total: 1 }); poll(); } }, 300);
    return {
      stop: function () {
        stopped = true;
        clearTimeout(shown); clearTimeout(timer);
        if (box) { box.remove(); box = null; }
      }
    };
  }

  // --- ボタン ---
  var ACTIONS = {
    lotLookup: function () {
      var lot = ($("#lotNo") || {}).value || "";
      return post("/api/lot-lookup", { current: collect(), lotNo: lot }).then(function (j) {
        toast(j.message, j.ok ? (j.level || "info") : (j.level || "warn"));
        paint(j.state);
        if (j.ok && $("#weight1")) { $("#weight1").focus(); }
      });
    },
    applyWeight: function () {
      return post("/api/apply-weight", { current: collect() }).then(function (j) {
        toast(j.message, j.ok ? (j.level || "info") : (j.level || "warn"));
        paint(j.state);
      });
    },
    calcTare: function () {
      var bar = progressBar("重量計算_DB");
      return post("/api/calc-tare", { current: collect() }).then(function (j) {
        bar.stop();
        toast(j.message, j.ok ? (j.level || "info") : (j.level || "warn"));
        paint(j.state);
      }, function (e) { bar.stop(); throw e; });
    },
    clearAll: function () {
      if (!window.confirm("入力内容（NW・GW・高さ）と台紙の本数欄をクリアします。よろしいですか？")) {
        return Promise.resolve();
      }
      return post("/api/clear", { current: collect() }).then(function (j) {
        toast(j.message, "info");
        paint(j.state);
      });
    },
    printLabels: function () {
      return post("/api/print-targets", { current: collect() }).then(function (j) {
        if (!j.ok) { toast(j.message, "warn"); return; }
        return confirmOpen(j.message, url("/labels/print?ob=" + j.targets.join(",")));
      });
    },
    quickHeight: function () {
      var v = window.prompt("コイル積み数", "10");
      if (v === null || v === "") { return Promise.resolve(); }
      return post("/api/quick-height", { current: collect(), coilCount: v })
        .then(function (j) {
          if (!j.ok || !j.quick) { toast(j.message, "warn"); return; }
          showQuick(j.quick, v);
        });
    }
  };

  // 積み高さの結果は、消すまで出したままにする（作業中に見ていたい参考値）
  function showQuick(q, count) {
    var card = $("#quickCard");
    if (!card) { toast(q.message || "", "info"); return; }
    var set = function (id, text) {
      var el = $(id);
      if (el) { el.textContent = text; }
    };
    // 風袋計算の本体と同じ式(統合 1.2.47。以前は VBA のまま 1 本分で「1条分」「ﾊﾟﾚｯﾄ+分」と出していた)
    set("#qWidth", (q.width || "") + "mm × " + (q.count || count) + "本");
    set("#qTip", q.tip);
    set("#qTipF", q.tipFormula);
    set("#qCoils", q.coils);
    set("#qCoilsF", q.coilsFormula);
    set("#qCoilH", q.coilHeight);
    set("#qCoilHF", q.coilHeightFormula);
    set("#qPack", q.packHeight);
    set("#qPackF", q.packFormula);
    card.hidden = false;
  }
  (function () {
    var x = $("#quickClose");
    if (x) {
      x.addEventListener("click", function () { $("#quickCard").hidden = true; });
    }
  })();

  // LOTNO で Enter → 検査番号を引く(統合 1.2.57)
  var lotBox = $("#lotNo");
  if (lotBox) {
    // ロット番号の形(7 桁)がそろったら、押さなくても自動で引く(統合 1.2.63・現場の指示)。
    // 形はサーバの `validation.LOT_PATTERN`(画面へ渡す)。読み取り機の速い入力も、そろってから少し待って1回だけ
    var lotRe = null;
    try { lotRe = new RegExp(lotBox.getAttribute("data-lot-pattern") || "^$"); } catch (err) { lotRe = null; }
    var lastAutoLot = "";
    var autoLotTimer = null;
    lotBox.addEventListener("input", function () {
      clearTimeout(autoLotTimer);
      autoLotTimer = setTimeout(function () {
        var v = String(lotBox.value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
        if (!lotRe || !lotRe.test(v)) { lastAutoLot = ""; return; }   // 打ち直したら、同じ番号でもまた引く
        if (v === lastAutoLot) { return; }
        lastAutoLot = v;
        var b = form.querySelector('[data-act="lotLookup"]');
        if (b && !b.disabled) { b.click(); }
      }, 150);
    });
    lotBox.addEventListener("keydown", function (ev) {
      if (ev.key !== "Enter") { return; }
      ev.preventDefault();
      var b = form.querySelector('[data-act="lotLookup"]');
      if (b && !b.disabled) { b.click(); }
    });
  }

  form.addEventListener("click", function (e) {
    var btn = e.target.closest("[data-act]");
    if (!btn) { return; }
    e.preventDefault();
    var fn = ACTIONS[btn.getAttribute("data-act")];
    if (!fn) { return; }
    btn.disabled = true;
    Promise.resolve(fn())
      .catch(function (err) { toast(String(err), "error"); })
      .then(function () { btn.disabled = false; });
  });

  // --- 初期表示 ---
  var boot = window.__PPL_STATE__ || null;
  if (boot) {
    document.body.classList.toggle("mode-cb", !!(boot.selectedCb || boot.namedCb));
    var cb = $("#cbMode");
    if (cb) {
      cb.checked = !!(boot.selectedCb || boot.namedCb);
      $$("input[name=size]").forEach(function (r) {
        var isCb = (r.dataset.kind !== "ob");
        r.closest("label").hidden = (cb.checked !== isCb);
      });
    }
    paint(boot);
  }
  markSelectedRow();
  setTimeout(function () { if (sentForm === null) { sentForm = formKey(); } }, 0);
})();
