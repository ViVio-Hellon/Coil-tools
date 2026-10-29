/*
  meisai.js — 梱包明細の画面

  VBA `frmCoilPacking` のイベントに対応する。

  【状態はサーバが持つ】
  操作のたびにサーバが「いま画面がどう見えるか」を丸ごと返し、ここは
  それを描くだけ。**画面側に状態を持たない** ── 2か所に持つと、
  どちらが正しいかを決める規則が要り、ずれたときに直す先が分からなくなる。

  ここが自前で持つのは、サーバに送るまでの**入力途中の値**だけ
  (重量の入力欄・廃棄モードの on/off)。
*/
import { api, ApiError, BASE } from "./api.js";
import * as health from "./health.js";
import * as tab from "./screen.js";
import * as master from "./master.js";
import * as settings from "./settings.js";
import { importWithProgress, showNote, waiting, wireBusy } from "./ui.js";

const $ = (id) => document.getElementById(id);

// 画面だけが持つもの
let haikiMode = false;         // 廃棄選択中か (VBA `mHaikiMode`)
let editingZen = false;        // 前工程実績数の修正中か (VBA ギミックA)
let selectedOutput = null;     // 一覧で選んでいる出力のNo
let selectedOrder = null;      // 選んでいる受注番号
// まとめ刷りで選んだ出力のNo。**1枚ずつ別の用紙**に出る(A4 1枚に
// 2枚並べる運用はしていない)。
const checkedOutputs = new Set();

// ------------------------------------------------------------------
// 通知
// ------------------------------------------------------------------
let noticeTimer = null;
function notify(message, kind = "") {
  const el = $("notice");
  el.textContent = message;
  el.className = "banner" + (kind ? ` ${kind}` : "");
  el.hidden = !message;
  clearTimeout(noticeTimer);
  if (message && kind === "ok") {
    noticeTimer = setTimeout(() => { el.hidden = true; }, 6000);
  }
}

/** 引き継がれた画面を畳む。**触らせない**のが要点。

    ここに出ている値(ロット番号・条番号・積み上げ)は、もう別の画面が
    開いているロットのもの。そのまま押せると、現物と紙が食い違う。 */
function lockScreen() {
  const box = document.getElementById("takenOver");
  if (box) box.hidden = false;
  // 畳まれた側は、以後あとから開くタブに「居る」と返さない ──
  // 返すと、次に開いたタブまで2枚目扱いになる
  tab.stopAnswering();
  // 下の画面を押せなくする。**隠すのではなく止める** ── 消してしまうと
  // 「何が起きたのか」が分からない
  document.querySelector(".wrap")?.setAttribute("inert", "");
}

/** 断りを画面に出す。**文言はサーバが持っている**ので、そのまま見せる。 */
function showError(err) {
  if (err instanceof ApiError) {
    // **心拍を待たずに畳む。** 20秒のあいだ押し続けられると、
    // 押すたびに赤帯が出るだけで理由が伝わらない
    if (err.code === "screen_taken") { health.screenLost(); return; }
    notify(err.message, "danger");
  } else {
    notify("処理できませんでした。もう一度お試しください。", "danger");
    console.error(err);
  }
}

// ------------------------------------------------------------------
// 確認ダイアログ (VBA `MsgBox ... vbYesNo`)
// ------------------------------------------------------------------
function confirmAsk(message, detail = []) {
  return new Promise((resolve) => {
    const dialog = $("confirmDialog");
    $("confirmBody").textContent = message;
    if (detail.length) {
      const list = document.createElement("ul");
      for (const line of detail) {
        const item = document.createElement("li");
        item.textContent = line;
        list.appendChild(item);
      }
      $("confirmBody").appendChild(list);
    }
    dialog.addEventListener("close", function once() {
      dialog.removeEventListener("close", once);
      resolve(dialog.returnValue === "yes");
    });
    dialog.showModal();
  });
}

// ------------------------------------------------------------------
// 描画
// ------------------------------------------------------------------
function render(state) {
  window.APP.state = state;

  $("lotNo").value = state.lot_no || "";
  $("zenKotei").value = state.zen_kotei || "";
  $("seqNote").textContent = `現在の連番 No${state.current_seq}`;

  renderLot(state);
  renderWeights(state);
  renderGrid(state);
  renderStack(state);
  renderOutputs(state);
}

function renderLot(state) {
  const found = state.found && state.lot;
  const hasStrands = Boolean(found && state.strands.length);
  $("lotInfo").hidden = !found;
  $("orderRow").hidden = !found;
  $("weightCard").hidden = !found;
  // 条番号と積み上げは**別々に隠す**。横に並んでいるので、
  // まとめて隠すと片方だけ出したいときに効かない
  $("strandCard").hidden = !hasStrands;
  $("stackCard").hidden = !hasStrands;
  $("outputCard").hidden = !(found && state.outputs.length);

  // 丈数が入らなかった理由と、途中経過の値が台帳と違うこと(統合 1.2.49)
  const note = $("lotNote");
  const lotNotes = found ? [state.lot.jou_su_note, state.lot.ledger_note].filter(Boolean) : [];
  note.hidden = !lotNotes.length;
  if (lotNotes.length) note.textContent = lotNotes.join("\n");

  if (!found) {
    $("warisu").textContent = "";
    return;
  }

  const lot = state.lot;
  for (const key of ["yoto_code", "yoto_name", "zaishitsu", "choshitsu",
                     "odr_thickness", "odr_width", "seizou_thickness",
                     "seizou_width", "sekkei_course", "jisseki_course"]) {
    $(`f_${key}`).textContent = lot[key] || "";
  }
  $("warisu").textContent = `縦割${lot.tate_wari} × 横割${lot.yoko_wari}`;
  $("jouSu").value = state.jou_su || "";

  // 受注番号。押すと包装仕様NOが切り替わる (VBA `lstJuchuList_Click`)
  const orders = $("orders");
  orders.textContent = "";
  $("orderCount").textContent = `${state.orders.length}件`;
  if (selectedOrder === null && state.orders.length) {
    selectedOrder = state.orders[0].order_no;
  }
  for (const order of state.orders) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "order";
    chip.textContent = order.order_no;
    chip.setAttribute("aria-pressed", String(order.order_no === selectedOrder));
    chip.addEventListener("click", () => {
      selectedOrder = order.order_no;
      renderLot(window.APP.state);
    });
    orders.appendChild(chip);
  }
  renderSpec(state);
}

function renderSpec(state) {
  const spec = $("specNo");
  spec.textContent = "";
  const order = state.orders.find((o) => o.order_no === selectedOrder);
  if (!order || !order.spec_no) {
    // VBA: 包装仕様NOが空ならリンクを張らず、黒いまま
    spec.textContent = order ? "(未登録)" : "";
    return;
  }
  // VBA `lblHousouShiyou_Click`: クリップボードへ写して閲覧システムを開く
  const link = document.createElement("a");
  link.href = window.APP.hosoUrl;
  link.target = "_blank";
  link.rel = "noopener";
  link.textContent = order.spec_no;
  link.title = "クリックでコピーし、包装仕様書の画面を開きます";
  link.addEventListener("click", () => {
    navigator.clipboard?.writeText(order.spec_no)
      .then(() => notify(`包装仕様NO ${order.spec_no} をコピーしました`, "ok"))
      .catch(() => { /* コピーできなくても画面は開く */ });
  });
  spec.appendChild(link);
}

function renderWeights(state) {
  const box = $("weights");
  // 丈の数が同じなら欄は作り直さず、値だけ合わせる。**打っている最中の欄には触らない**
  // (統合 1.2.26。作り直すと、丈1を送った返事が、丈2へ打ちかけた値を前の値に戻していた)
  const inputs = box.querySelectorAll("input[id^='weight']");
  if (state.jou_su > 0 && inputs.length === state.jou_su) {
    inputs.forEach((input, index) => {
      if (input === document.activeElement && document.hasFocus()) return;
      input.value = state.weights[index] || "";
      input.classList.remove("is-bad");
    });
    return;
  }
  box.textContent = "";
  for (let jou = 1; jou <= state.jou_su; jou += 1) {
    const cell = document.createElement("div");
    cell.className = "weight";

    const lab = document.createElement("div");
    lab.className = "lab";
    lab.textContent = `丈${jou}`;
    cell.appendChild(lab);

    const input = document.createElement("input");
    input.type = "text";
    input.inputMode = "numeric";
    input.id = `weight${jou}`;
    input.value = state.weights[jou - 1] || "";
    // VBA `clsWeightCtrl.Txt_KeyPress`: 数字以外を拒否(小数点も不可)
    input.addEventListener("beforeinput", (event) => {
      if (event.data && !/^\d+$/.test(event.data)) event.preventDefault();
    });
    // 日本語入力(IME)の確定は beforeinput で止められない(統合 1.2.6)。全角の数字は半角へ
    // 直す(同じ数)。小数点などは**消さずに知らせる** ── 消すと 2502.7 が 25027 になる。
    // そのまま送ればサーバが断る(以前は黙って 2502 に切り捨てていた)
    input.addEventListener("input", () => {
      const half = input.value.replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
      if (half !== input.value) input.value = half;
      const bad = /[^0-9]/.test(half);
      input.classList.toggle("is-bad", bad);
      if (bad) notify(`丈${jou}の重量は整数で入力してください(小数点は使えません)`, "ng");
    });
    input.addEventListener("change", sendWeights);
    cell.appendChild(input);

    // ±1 / ±10 (VBA の spn1 / spn10)
    const spin = document.createElement("div");
    spin.className = "spin";
    for (const step of [-10, -1, 1, 10]) {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = step > 0 ? `+${step}` : String(step);
      button.addEventListener("click", () => {
        // 整数でない値(日本語入力で入った 2502.7 など)を parseInt で切り捨てて足さない
        // (統合 1.2.12。以前は 2502.7 → 2503 として送っていた)
        const now = input.value.trim();
        if (now !== "" && !/^[0-9]+$/.test(now)) {
          input.classList.add("is-bad");
          notify("重量は整数で入力してください(小数点は使えません)。", "danger");
          return;
        }
        const next = Math.max(0, (parseInt(now, 10) || 0) + step);
        input.value = String(next);
        sendWeights();
      });
      spin.appendChild(button);
    }
    cell.appendChild(spin);
    box.appendChild(cell);
  }
}

function renderGrid(state) {
  const grid = $("grid");
  grid.textContent = "";
  state.rows.forEach((row, index) => {
    const jou = index + 1;
    const block = document.createElement("div");
    block.className = "jou";

    const lab = document.createElement("div");
    lab.className = "lab";
    lab.textContent = `丈${jou} ／ ${state.strands[index]}条`;
    block.appendChild(lab);

    const coils = document.createElement("div");
    coils.className = "coils";
    for (const cell of row) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "coil";
      button.dataset.state = cell.state;
      button.dataset.key = cell.key;
      button.textContent = cell.key;
      // 使用済みは押しても何も起きない(VBA も早期 Exit)。
      // **廃棄済みは押せる** ── 廃棄モードで解除するため
      button.disabled = cell.state === "used" && !haikiMode;
      button.addEventListener("click", () => clickCoil(cell.key));
      coils.appendChild(button);
    }
    block.appendChild(coils);
    grid.appendChild(block);
  });

  $("strandHint").textContent = haikiMode
    ? "廃棄選択中：クリックで廃棄の付け外し"
    : "クリックで積み上げへ";
}

function renderStack(state) {
  const box = $("stack");
  box.textContent = "";
  $("stackMax").value = state.stack_max || "";
  for (let i = 0; i < state.slots.length; i += 1) {
    const key = state.slots[i];
    const button = document.createElement("button");
    button.type = "button";
    button.className = key ? "slot" : "slot empty";
    button.textContent = key || `${i + 1}`;
    button.disabled = !key;
    button.addEventListener("click", () => unstack(i + 1));
    box.appendChild(button);
  }
  $("btnOutput").disabled = !state.complete;
  renderEdit(state);
}

/**
 * 出力済みの修正中か(VER 0.13.23)。修正中は「出力」が「修正して出力」になり、
 * 条番号の再表示・丈数の変更は止める(サーバも断る)。
 */
function renderEdit(state) {
  const edit = state.edit;
  $("editBanner").hidden = !edit;
  $("btnOutput").textContent = edit ? "修正して出力" : "出力";
  // いま断られる操作は**サーバが決めて渡す**(`state.locks`。統合 1.2.26)。画面で決め直さない ──
  // 以前は画面とサーバに別々に書いていて、条番号表示がサーバの一覧から漏れた
  const locks = state.locks || {};
  const lockedBy = (action, code) => Boolean(locks[action]) && locks[action].code === code;
  $("btnStrands").disabled = Boolean(locks.strands);
  $("btnStrands").title = locks.strands ? locks.strands.message : "";
  $("strandsLockNote").hidden = !lockedBy("strands", "strands_locked") || lockedBy("jou_su", "jou_su_locked");
  $("strandsLockNote").textContent = locks.strands ? locks.strands.message : "";
  // 丈数は、修正中か、積み上げを始めたら変えられない(VER 0.13.24。変えるときは全クリア)
  const jouLocked = Boolean(locks.jou_su);
  $("jouSu").disabled = jouLocked;
  document.querySelectorAll("[data-jou]").forEach((b) => { b.disabled = jouLocked; });
  $("jouLockNote").hidden = !lockedBy("jou_su", "jou_su_locked");
  // 重量も、積み上げを始めたら変えられない(VER 0.13.35。−/＋ と欄を止める。サーバも断る)
  const weightLocked = Boolean(locks.weights);
  document.querySelectorAll("#weights input, #weights button").forEach((el) => {
    el.disabled = weightLocked;
    el.title = weightLocked ? locks.weights.message : "";
  });
  // 前工程実績数も、修正中か積み上げを始めたら変えられない(VER 0.13.25)
  const zenLocked = Boolean(locks.zen_kotei);
  $("btnEditZen").disabled = zenLocked;
  $("zenLockNote").hidden = !lockedBy("zen_kotei", "zen_kotei_locked");
  if (zenLocked && editingZen) {
    // 打ちかけの「確定する」は取りやめる(サーバも断る)
    editingZen = false;
    $("zenKotei").readOnly = true;
    $("zenKotei").value = state.zen_kotei || "";
    $("btnEditZen").textContent = "修正する";
  }
  if (!edit) return;
  $("editTitle").textContent = `No${edit.no}を修正中`;
  $("editBefore").textContent = `修正前: ${edit.keys.join(" ")}`;
}

function renderOutputs(state) {
  const box = $("outputs");
  box.textContent = "";
  if (!state.outputs.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = "まだ出力していません。";
    box.appendChild(empty);
    checkedOutputs.clear();
    updatePrintLabel();
    return;
  }
  // 最新を選んでおく (VBA `RefreshSheetList` は末尾を自動選択する)
  if (selectedOutput === null
      || !state.outputs.some((o) => o.no === selectedOutput)) {
    selectedOutput = state.outputs[state.outputs.length - 1].no;
  }
  // 消えた出力のチェックは落とす
  const alive = new Set(state.outputs.map((o) => o.no));
  for (const no of [...checkedOutputs]) {
    if (!alive.has(no)) checkedOutputs.delete(no);
  }

  const editingNo = state.edit ? state.edit.no : 0;
  for (const output of state.outputs) {
    const row = document.createElement("div");
    row.className = output.no === editingNo ? "output editing" : "output";
    row.setAttribute("aria-pressed", String(output.no === selectedOutput));

    // まとめて刷れるように選べるようにする(1枚ずつ別の用紙に出る)
    const check = document.createElement("input");
    check.type = "checkbox";
    check.id = `pick${output.no}`;
    check.checked = checkedOutputs.has(output.no);
    check.addEventListener("change", () => {
      if (check.checked) checkedOutputs.add(output.no);
      else checkedOutputs.delete(output.no);
      updatePrintLabel();
      updateEditButton();
    });

    const name = document.createElement("label");
    name.className = "n";
    name.htmlFor = check.id;
    name.textContent = output.name;

    const open = document.createElement("button");
    open.type = "button";
    open.className = "pick";
    open.textContent = `${output.count}本`;
    open.title = "この1枚を選ぶ";
    open.addEventListener("click", () => {
      selectedOutput = output.no;
      renderOutputs(window.APP.state);
      $("btnPrint").disabled = false;
    });

    row.append(check, name);
    if (output.no === editingNo) {
      const tag = document.createElement("span");
      tag.className = "tag";
      tag.textContent = "修正中";
      row.appendChild(tag);
    }
    row.append(open);
    box.appendChild(row);
  }
  updatePrintLabel();
  updateEditButton();
}

/**
 * 「出力済み修正」は**チェック1つのときだけ**押せる(VER 0.13.23・複数は不可)。
 * 押せない理由はボタンの横に出す。
 */
function updateEditButton() {
  const state = window.APP.state || {};
  const button = $("btnEditOutput");
  const note = $("editNote");
  const n = checkedOutputs.size;
  let reason = "";
  if (state.edit) reason = `No${state.edit.no}を修正中です`;
  else if (!(state.strands || []).length) reason = "先に「条番号表示」を押してください";
  else if (n === 0) reason = "直す1枚にだけチェックを入れて押します";
  else if (n > 1) reason = `修正は1枚ずつです（いま${n}枚にチェック）`;
  else if ((state.slots || []).some(Boolean)) reason = "積み上げ中のコイルを出力するか外してから押します";
  button.disabled = Boolean(reason);
  note.textContent = reason || `No${[...checkedOutputs][0]}を積み直します`;
  note.classList.toggle("warn", n > 1);
}

/** 印刷ボタンに「いま何枚刷るか」を出す。**紙の枚数も添える。** */
function updatePrintLabel() {
  const nos = printTargets();
  const button = $("btnPrint");
  button.disabled = nos.length === 0;
  // **開いた紙面で書き足せる**ことをここで言う。押す前に見る場所なので、
  // 紙面を開いてから気づくより早い
  const hint = "開いた紙面でサイズ・LOTNOを書き足せます";
  // 用紙はA4横。**左半分に明細、右半分は空ける**(現場で切り取る)
  if (nos.length <= 1) {
    button.textContent = "印刷";
    $("printNote").textContent = nos.length
      ? `A4横の左半分に出ます（右半分は空白）。${hint}` : "";
    return;
  }
  button.textContent = `印刷 (${nos.length}枚)`;
  $("printNote").textContent =
    `${nos.length}枚 → 用紙${nos.length}枚（1枚ずつ左半分に）。${hint}`;
}

/** 刷る対象。チェックがあればそれ、無ければ選択中の1枚。 */
function printTargets() {
  if (checkedOutputs.size) return [...checkedOutputs].sort((a, b) => a - b);
  return selectedOutput === null ? [] : [selectedOutput];
}

// ------------------------------------------------------------------
// 操作
// ------------------------------------------------------------------
/*
  作業の操作は**押した順に1つずつ**送る(統合 1.2.20)。重量は欄を離れたとき(`change`)に
  送るので、入れてすぐ「条番号表示」などを押すと2つの要求が同時に出る。サーバは書き込みを
  1つずつ通すが、**どちらが先かは決まらない**(資材計算では前の値で計算した結果が出た)。
  前の返事を待ってから次を送れば、押した順に処理されて、押した順に描かれる。
*/
let sendChain = Promise.resolve();
function queued(path, data) {
  // **押したときの画面のロットを付ける**(統合 1.2.26)。並んで待つあいだにロットが変わったら、
  // サーバは使わずに断る(`lot_changed`)。前のロットの画面で打った重量を、開いたロットへ入れない
  const lot = (window.APP.state || {}).lot_no || "";
  const run = sendChain.then(() => api.post(path, data, { "X-Screen-Lot": lot }));
  sendChain = run.catch(() => {});
  return run;
}

/** サーバを呼んで、返ってきた画面を描く。 */
async function send(path, data, onDone) {
  try {
    const body = await queued(path, data);
    if (body.confirm) return body;          // 確認が要る。呼び元が扱う
    if (body.state) render(body.state);
    if (body.message) notify(body.message, "ok");
    if (onDone) onDone(body);
    return body;
  } catch (err) {
    showError(err);
    return null;
  }
}

async function openLot(lotNo, confirm = false, cancelEdit = false) {
  const edit = (window.APP.state || {}).edit;
  if (edit && !cancelEdit) {
    if (!await confirmAsk(
        `No${edit.no}を修正中です。修正をやめてロットを開き直しますか？（紙は前のままです）`)) {
      $("lotNo").value = window.APP.state.lot_no || "";
      return;
    }
    cancelEdit = true;      // サーバは了解(cancel_edit)が無いと修正中のロット切替を断る
  }
  const body = await send("/api/lot", { lot_no: lotNo, confirm, cancel_edit: cancelEdit });
  if (body && body.confirm === "unprinted") {
    if (await confirmAsk(body.message)) openLot(lotNo, true, cancelEdit);
    else $("lotNo").value = window.APP.state.lot_no || "";
  }
  if (body && body.ok) checkNewLedger();
}

/** いま画面の重量欄に入っている値(表示している丈の数だけ。打ったまま)。 */
function screenWeights() {
  const state = window.APP.state || {};
  const values = [];
  for (let jou = 1; jou <= (state.jou_su || 0); jou += 1) {
    const input = $(`weight${jou}`);
    values.push(input ? input.value : "");
  }
  return values;
}

function sendWeights() {
  const values = screenWeights();
  // 未入力があるうちは送らない(断られて赤帯が出続けるだけ)
  if (values.some((v) => String(v).trim() === "")) return;
  send("/api/weights", { weights: values });
}

/*
  「条番号表示」「出力」は、押したときの重量欄の値を添えて送る(VBA は押すたびに画面の欄を
  読み直していた)。欄を離れたときの送信(`sendWeights`)は空欄なら送らず、断られてもサーバは
  前の値のまま ── 添えないと、画面は空なのに前の重量が紙に載った。断られたら赤帯で理由が出て、
  画面は描き直さないので、欄は打った値のまま残る。
*/
async function buildStrands(confirm = false) {
  const body = await send("/api/strands", { confirm, weights: screenWeights() });
  if (body && body.confirm === "overwrite_restored") {
    if (await confirmAsk(body.message)) buildStrands(true);
  }
}

function clickCoil(key) {
  if (haikiMode) send("/api/discard", { key });
  else send("/api/stack", { key });
}

function unstack(slotNo) {
  send("/api/unstack", { slot_no: slotNo });
}

async function doOutput(confirm = false) {
  // 修正中なら、サーバがそのNoを積み直した並びで出し直す(確認は kind "edit")
  const editing = Boolean((window.APP.state || {}).edit);
  const body = await send("/api/output", { confirm, weights: screenWeights() });
  if (body && body.confirm === "output") {
    const first = body.confirms[0];
    const detail = body.confirms.flatMap((c) => c.detail || []);
    const message = body.confirms.map((c) => c.message).join("\n");
    if (await confirmAsk(message, detail)) doOutput(true);
    else if (!first) notify("");
  } else if (body && body.ok) {
    selectedOutput = body.seq_no;
    // 直した1枚をそのまま刷れるように、チェックは外して選んでおく
    if (editing) checkedOutputs.clear();
    renderOutputs(window.APP.state);
  }
}

/** 出力済みの1枚を積み直しに入る(チェック1つ)。 */
function beginEdit() {
  if (checkedOutputs.size !== 1) return;
  const no = [...checkedOutputs][0];
  send("/api/edit-output", { no }, () => {
    selectedOutput = no;
    renderOutputs(window.APP.state);
  });
}

async function cancelEdit() {
  const edit = (window.APP.state || {}).edit;
  if (!edit) return;
  if (!await confirmAsk(`No${edit.no}の修正をやめますか？（紙は前のままです）`)) return;
  send("/api/edit-cancel", {});
}

async function doPrint() {
  const nos = printTargets();
  if (!nos.length) return;
  const lotNo = window.APP.state.lot_no;
  // 先に印刷済みを記録してから帳票を開く。ブラウザの印刷は
  // 押されたかどうかを確実には拾えない
  // 作業の記録に「印刷画面」と残す(どのロットのどの No か。VER 0.13.30)
  await send("/api/printed", { lot: lotNo, nos });
  const token = encodeURIComponent(window.APP.token);
  const url = nos.length === 1
    ? `${BASE}/report/${encodeURIComponent(lotNo)}/${nos[0]}?t=${token}`
    : `${BASE}/report/${encodeURIComponent(lotNo)}?nos=${nos.join(",")}&t=${token}`;
  window.open(url, "_blank", "noopener");
}

/**
 * 全クリア(VER 0.13.23)。そのロットの出力済み・副番履歴・連番・条番号と廃棄を消して、
 * 初めからやり直す。サーバが確認の文言(何枚消えるか)を返すので、それを訊いてから送る。
 */
async function clearAll(confirm = false) {
  const body = await send("/api/clear-all", { confirm });
  if (body && body.confirm === "clear_all") {
    if (await confirmAsk(body.message)) clearAll(true);
  } else if (body && body.ok) {
    checkedOutputs.clear();
    selectedOutput = null;
    renderOutputs(window.APP.state);
  }
}

// ------------------------------------------------------------------
// 取り込み元の鮮度(バージョン情報に出す。設定画面は `settings.js`)
// ------------------------------------------------------------------
function renderStamps(stamps, boxId) {
  const box = $(boxId);
  box.textContent = "";
  for (const s of stamps) {
    const line = document.createElement("div");
    // **届くかどうかと、取り込んであるかどうかは別。** ここは届くかを
    // 見に行かない(`found` が null)ので、いつ時点のものかだけを出す
    const reach = s.found === null || s.found === undefined ? ""
      : (s.found ? "届く  " : "届かない  ");
    line.className = s.found === false ? "miss" : "ok";
    line.textContent =
      `${s.table}  ${reach}${s.as_of || "まだ取り込んでいません"}  ${s.count.toLocaleString()}件`;
    box.appendChild(line);
  }
}

// ------------------------------------------------------------------
// 明細の履歴 (全ライン・3年)
// ------------------------------------------------------------------
const pad2 = (n) => String(n).padStart(2, "0");
const ymd = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;

function renderHistory(body) {
  const box = $("hisState");
  box.textContent = "";
  const line = (text) => {
    const div = document.createElement("div");
    div.textContent = text;
    box.appendChild(div);
  };
  line(`共有: ${body.shared_path}`);
  if (body.shared_count !== null && body.shared_count !== undefined && body.shared_count > 0) {
    line(`共有の履歴: ${body.shared_count.toLocaleString()}枚・${body.shared_coils.toLocaleString()}本`
      + `（${body.shared_oldest.slice(0, 10)} 〜 ${body.shared_newest.slice(0, 10)}）`);
  } else if (body.shared_count === 0) {
    line("共有の履歴: まだ0枚です（明細を出力すると送られます）");
  }
  if (body.shared_problem) line(`※ ${body.shared_problem}`);
  line(body.pending
    ? `この PC の送り残し: ${body.pending}枚（届けば自動で送ります）`
    : "この PC の送り残し: なし");
  if (body.last_error) line(`送れなかった理由: ${body.last_error}`);
  else if (body.last_sent_at) line(`最後に送った: ${body.last_sent_at.slice(0, 16)}`);
  line(`${body.keep_years}年より前（${body.cutoff} より前）は自動で整理します`);
  const bad = Boolean(body.last_error) || (body.shared_problem && body.shared_count === null);
  box.className = bad ? "set-share bad" : "set-share";
}

async function openHistory() {
  const today = new Date();
  if (!$("hisFrom").value) $("hisFrom").value = ymd(new Date(today.getFullYear(), today.getMonth(), 1));
  if (!$("hisTo").value) $("hisTo").value = ymd(today);
  showNote("hisSendNote", "");
  showNote("hisFindNote", "");
  // **先に開く。** 共有の様子は後から埋める(共有の応えしだいで数秒かかる)
  const box = $("hisState");
  box.className = "set-share";
  const stop = waiting(box, "共有の履歴を確かめています…");
  $("historyDialog").showModal();
  try {
    const body = await api.get("/api/history/status");
    stop();
    box.hidden = false;
    renderHistory(body);
  } catch (err) {
    stop();
    box.hidden = false;
    if (err instanceof ApiError && err.code === "screen_taken") {
      $("historyDialog").close();
      showError(err);
      return;
    }
    box.className = "set-share bad";
    box.textContent = `共有の履歴の様子を読めませんでした: ${err.message}`;
  }
}

async function sendHistory() {
  $("btnHisSend").disabled = true;
  try {
    const body = await api.post("/api/history/send", {});
    renderHistory(body);
    showNote("hisSendNote", body.message, Boolean(body.last_error));
    if (body.sending) watchSending();
  } catch (err) {
    showNote("hisSendNote", err.message, true);
  } finally {
    $("btnHisSend").disabled = false;
  }
}

/**
 * 裏で送っているあいだは、送り終わるまで様子(送り残し・最後に送った)を取り直す。
 * 以前は「裏で続けます」と出たあと、閉じて開き直すまで古い数のままだった(統合 1.2.11)
 */
let sendingTimer = null;
function watchSending(tries = 40) {
  clearTimeout(sendingTimer);
  sendingTimer = setTimeout(async () => {
    if (!$("historyDialog").open) return;
    try {
      const body = await api.get("/api/history/status");
      renderHistory(body);
      if (body.sending && tries > 1) { watchSending(tries - 1); return; }
      showNote("hisSendNote", body.last_error
        ? `送れませんでした(あとで送り直します): ${body.last_error}`
        : "送り終わりました。", Boolean(body.last_error));
    } catch { /* 開き直せば分かる */ }
  }, 3000);
}

function historyFilter() {
  return { from: $("hisFrom").value, to: $("hisTo").value,
           lot: $("hisLot").value, fuban: $("hisFuban").value };
}

let hisLoadLimit = 20;

function updatePicked() {
  const picked = [...document.querySelectorAll("#hisRows input:checked")];
  $("btnHisReprint").disabled = !picked.length || picked.length > hisLoadLimit;
  $("hisPicked").textContent = picked.length > hisLoadLimit
    ? `一度に開けるのは${hisLoadLimit}枚までです（いま${picked.length}枚）`
    : (picked.length ? `${picked.length}枚選んでいます` : "紙を選んでください（行を押すと選べます）");
}

function renderHistoryRows(body) {
  hisLoadLimit = body.load_limit || hisLoadLimit;
  const tbody = $("hisRows");
  tbody.textContent = "";
  for (const row of body.rows) {
    const tr = document.createElement("tr");
    // 作り直し(出力済み修正で出し直された)・取り消し(全クリア)は、使わない紙
    if (row["状態"] === "作り直し") tr.className = "replaced";
    else if (row["状態"] === "取り消し") tr.className = "replaced cancelled";
    const cell = (text, cls = "") => {
      const td = document.createElement("td");
      if (cls) td.className = cls;
      td.textContent = text;
      tr.appendChild(td);
      return td;
    };
    const pick = document.createElement("input");
    pick.type = "checkbox";
    pick.value = row["送信ID"];
    pick.setAttribute("aria-label", `${row["ロット番号"]}-No${row["No"]} を選ぶ`);
    pick.addEventListener("change", updatePicked);
    const first = document.createElement("td");
    first.appendChild(pick);
    tr.appendChild(first);
    cell(String(row["出力日時"]).slice(0, 16));
    cell(`${row["ロット番号"]}-No${row["No"]}`);
    cell(String(row["本数"]), "num");
    cell(row["副番の並び"] || "", "fuban").title = row["副番の並び"] || "";
    cell(row["状態"], "state");
    const pc = cell(row["出力PC"]);
    // テストモードで出した紙(VER 0.13.30)。列は増やさず PC の欄に札を付ける
    if (row["モード"]) {
      const mark = document.createElement("span");
      mark.className = "testmode";
      mark.textContent = row["モード"];
      pc.appendChild(mark);
    }
    // 行のどこを押しても選べる(小さな四角だけを狙わせない)
    tr.addEventListener("click", (e) => {
      if (e.target !== pick) { pick.checked = !pick.checked; updatePicked(); }
    });
    tbody.appendChild(tr);
  }
  $("hisResult").hidden = false;
  updatePicked();
}

async function searchHistory() {
  $("btnHisSearch").disabled = true;
  $("hisFiles").hidden = true;
  try {
    const body = await api.post("/api/history/search", historyFilter());
    renderHistoryRows(body);
    const where = body.source === "local" ? "この PC の分から" : "全ラインから";
    showNote("hisFindNote", body.total
      ? `${where}${body.total.toLocaleString()}枚見つかりました。${body.note || ""}`
      : `見つかりませんでした。${body.note || ""}`, body.source === "local");
  } catch (err) {
    $("hisResult").hidden = true;
    showNote("hisFindNote", err.message, true);
  } finally {
    $("btnHisSearch").disabled = false;
  }
}

/** 選んだ紙を、履歴の中身のとおりに作り直して別のタブで開く(読むだけ)。 */
function reprintHistory() {
  const ids = [...document.querySelectorAll("#hisRows input:checked")].map((i) => i.value);
  if (!ids.length) return;
  const token = encodeURIComponent(window.APP.token);
  window.open(`${BASE}/report/history?ids=${ids.map(encodeURIComponent).join(",")}&t=${token}`,
              "_blank", "noopener");
}

async function exportHistory() {
  $("btnHisExport").disabled = true;
  $("hisFiles").hidden = true;
  showNote("hisFindNote", "書き出しています…");
  try {
    const body = await api.post("/api/history/export", historyFilter());
    $("hisDetailName").textContent = body.detail_file;
    $("hisSummaryName").textContent = body.summary_file;
    $("hisFolder").textContent = body.folder;
    $("hisFiles").hidden = false;
    showNote("hisFindNote", body.note
      || `${body.slips.toLocaleString()}枚・${body.detail_rows.toLocaleString()}本を書き出しました。`);
    renderHistory(await api.get("/api/history/status"));
  } catch (err) {
    showNote("hisFindNote", err.message, true);
  } finally {
    $("btnHisExport").disabled = false;
  }
}

// ------------------------------------------------------------------
// バージョン情報
// ------------------------------------------------------------------
/** 秒を「2時間14分」の形に。**秒まで出しても読む人が困るだけ。** */
function since(sec) {
  const s = Math.max(0, Math.round(Number(sec) || 0));
  const d = Math.floor(s / 86400);
  const h = Math.floor((s % 86400) / 3600);
  const m = Math.floor((s % 3600) / 60);
  if (d) return `${d}日${h}時間`;
  if (h) return `${h}時間${m}分`;
  if (m) return `${m}分`;
  return `${s}秒`;
}

async function openAbout() {
  // **健康確認と設定の両方から集める。** どちらか片方が落ちても
  // 出せるところまでは出す ── 版を知りたいときはたいてい何かが
  // おかしいときで、そこで画面ごと出ないのがいちばん困る
  const [health, settings] = await Promise.all([
    fetch(BASE + "/api/health", { cache: "no-store" })
      .then((r) => r.json()).catch(() => null),
    api.get("/api/settings").catch(() => null),
  ]);

  if (health) {
    $("abVer").textContent = `VER ${health.version}`;
    $("abRoot").textContent = health.app_root || "—";
    $("abAppId").textContent = health.app_id || "—";
    $("abPort").textContent = `${health.port || "なし（デスクトップ版）"} / ${health.pid}`;
    $("abUptime").textContent = since(health.uptime_sec);
    $("abPython").textContent = health.python
      ? `${health.python}　${health.python_exe || ""}`.trim() : "—";
    // 版の書き方が壊れていても、手元DBが共有に置かれていても、
    // 起動は止めない。**気づける場所に出す**
    const bad = $("abProblem");
    const problems = [health.version_problem, health.db_path_problem]
      .filter(Boolean);
    bad.textContent = problems.join("\n");
    bad.hidden = !problems.length;
  }
  if (settings) {
    $("abDb").textContent = settings.db_path;
    $("abLog").textContent = settings.log_dir;
    renderStamps(settings.stamps, "abStamps");
  }
  $("aboutDialog").showModal();
}

/** 上の帯の「いつ時点の台帳か」を描き直す。**いつ時点かはサーバが決める**(`as_of`)。 */
function renderTopStamps(stamps) {
  $("stamp").textContent = stamps.map((s) => `${s.table} ${s.as_of_short}`).join("　");
}

/** 取り込んだあと。帯と「まだ取り込まれていません」の案内を、読み直さずに新しくする(統合 1.2.11) */
function afterImport(body) {
  renderTopStamps(body.stamps);
  const banner = document.getElementById("noDataBanner");
  if (banner && body.ok) banner.hidden = true;
  // 「新しい台帳があります」も聞き直す(一部だけ取り込めたときは残る。統合 1.2.49)
  $("newLedgerBanner").hidden = true;
  checkNewLedger();
}

/*
  新しい台帳があるか(統合 1.2.49・現場の指摘)。以前は動いている間は仕掛データを読み直さず、
  上の帯も取り込んだ時点のままで「新しい台帳がある」とは言わなかった。画面を開いたとき・
  ロットを入れたとき・前に戻ったとき・数分ごとに聞く。サーバはファイルの大きさと更新時刻を
  見るだけ(中身は読まない)。確かめきれなかったときは帯をそのままにする。
*/
const LEDGER_POLL_MS = 3 * 60 * 1000;
let ledgerChecking = false;
async function checkNewLedger() {
  if (ledgerChecking) return;
  ledgerChecking = true;
  try {
    const body = await api.get("/api/import/check");
    if (!body || !body.checked) return;
    const newer = body.newer || [];
    if (newer.length) {
      $("newLedgerText").textContent =
        `${body.message}取り込むと、次にロットを入れたときから新しい台帳で引けます。`;
    }
    $("newLedgerBanner").hidden = !newer.length;
  } catch { /* 次に聞く */ } finally {
    ledgerChecking = false;
  }
}

async function runImport() {
  $("btnImport").disabled = true;
  $("btnNewLedgerImport").disabled = true;
  try {
    // **進み具合の窓を出す。** 何を読んでいるか・何秒たったかが見えないと、
    // 押せていないと思って押し直す(現場の指摘: 反応が鈍い)
    const body = await importWithProgress({ force: true }, "取り込み中");
    // **開いているロットが古いままなら、そう言う。** 黙って
    // 読み直すと、手で直した前工程実績数が消える
    // 知らせ(元が空なので前の中身のまま など。統合 1.2.49)も出す
    notify([body.summary, ...(body.notes || []), body.note, ...body.errors]
      .filter(Boolean).join("\n"), body.ok ? "ok" : "danger");
    afterImport(body);
  } catch (err) {
    showError(err);
  } finally {
    $("btnImport").disabled = false;
    $("btnNewLedgerImport").disabled = false;
  }
}

// ------------------------------------------------------------------
// 組み立て
// ------------------------------------------------------------------
function wire() {
  // --- ロット番号。英数字だけ・大文字・7桁で自動検索 ---
  const lotNo = $("lotNo");
  lotNo.addEventListener("input", () => {
    const cleaned = lotNo.value.replace(/[^A-Za-z0-9]/g, "").toUpperCase();
    if (cleaned !== lotNo.value) lotNo.value = cleaned;
    if (cleaned.length === window.APP.lotNoLength) openLot(cleaned);
  });

  // --- 前工程実績数の修正 (VBA ギミックA) ---
  $("btnEditZen").addEventListener("click", () => {
    const input = $("zenKotei");
    if (!editingZen) {
      editingZen = true;
      input.readOnly = false;
      input.focus();
      input.select();
      $("btnEditZen").textContent = "確定する";
      return;
    }
    send("/api/zen-kotei", { value: input.value }, () => {
      editingZen = false;
      input.readOnly = true;
      $("btnEditZen").textContent = "修正する";
    });
  });

  // --- 丈数 ---
  $("jouSu").addEventListener("change", () => {
    send("/api/jou-su", { value: $("jouSu").value });
  });
  for (const button of document.querySelectorAll("[data-jou]")) {
    button.addEventListener("click", () => {
      const step = Number(button.dataset.jou);
      const next = Math.min(window.APP.jousuMax,
                            Math.max(1, (parseInt($("jouSu").value, 10) || 0) + step));
      $("jouSu").value = String(next);
      send("/api/jou-su", { value: next });
    });
  }

  $("btnStrands").addEventListener("click", () => buildStrands());
  $("btnReverse").addEventListener("click", () => send("/api/reverse", {}));

  // --- 廃棄モード (VBA `cmdHaiki_Click`) ---
  $("btnHaiki").addEventListener("click", () => {
    haikiMode = !haikiMode;
    $("btnHaiki").textContent = haikiMode ? "廃棄選択中" : "廃棄選択";
    $("btnHaiki").classList.toggle("on", haikiMode);
    renderGrid(window.APP.state);
  });

  // --- 積み条数 ---
  // 積んだコイルは残したまま変えられる(VER 0.13.33)。断られたら(減らすと積んだコイルの
  // スロットが無くなる)欄をいまの値へ戻す
  const sendStackMax = async (value) => {
    const body = await send("/api/stack-max", { value });
    if (!body) $("stackMax").value = (window.APP.state || {}).stack_max || "";
  };
  $("stackMax").addEventListener("change", () => sendStackMax($("stackMax").value));
  // −/＋ でも入れられる(丈数と同じ形)。1〜最大の範囲に収める
  document.querySelectorAll("[data-stack]").forEach((button) => {
    button.addEventListener("click", () => {
      const box = $("stackMax");
      const max = parseInt(box.dataset.max, 10) || 15;
      const now = parseInt(box.value, 10) || 0;
      const next = Math.min(max, Math.max(1, now + parseInt(button.dataset.stack, 10)));
      if (String(next) === box.value) return;
      box.value = String(next);
      sendStackMax(box.value);
    });
  });

  // --- 出力・印刷 ---
  $("btnOutput").addEventListener("click", () => doOutput());
  $("btnPrint").addEventListener("click", doPrint);
  $("btnEditOutput").addEventListener("click", beginEdit);
  $("btnEditCancel").addEventListener("click", cancelEdit);
  $("btnClearAll").addEventListener("click", () => clearAll());
  $("btnImport").addEventListener("click", runImport);
  // 「新しい台帳があります」の帯から取り込む(左の「取り込み」と同じ。統合 1.2.49)
  $("btnNewLedgerImport").addEventListener("click", runImport);
  wireBusy();
  settings.wire({ showError, afterImport });
  $("btnHistory").addEventListener("click", openHistory);
  $("btnHisClose").addEventListener("click", () => $("historyDialog").close());
  $("btnHisSend").addEventListener("click", sendHistory);
  $("btnHisExport").addEventListener("click", exportHistory);
  $("btnHisSearch").addEventListener("click", searchHistory);
  $("btnHisReprint").addEventListener("click", reprintHistory);
  // 場所を写すだけ。**こちらからフォルダや Excel は開かない**
  $("btnHisCopyFolder").addEventListener("click", () => {
    navigator.clipboard?.writeText($("hisFolder").textContent)
      .then(() => showNote("hisFindNote", "場所をコピーしました。"))
      .catch(() => {});
  });
  // マスタ管理の「置き場所を変える」。マスタ管理を閉じて、設定の
  // 梱包資材マスタのフォルダの欄を開く(**変えたらすぐ、マスタ管理もその場所を見る**)
  master.wire({ openPlace: () => settings.open({ tab: "place", focus: "setShareDir" }) });

  // 引き継ぐ。**手放させてから開き直す** ── 開き直すだけだと、
  // いま持っている画面に断られて2度手間になる
  for (const bid of ["btnRetake", "btnSecondTake"]) {
    $(bid).addEventListener("click", () => takeOverScreen($(bid)));
  }
  $("btnSecondRetry").addEventListener("click", () => location.reload());

  $("btnAbout").addEventListener("click", () => openAbout().catch(showError));
  $("btnAboutClose").addEventListener("click",
    () => $("aboutDialog").close());
}

// ------------------------------------------------------------------
// 起動
// ------------------------------------------------------------------
/** 2枚目として開いたことを告げる。**業務の画面は出さない。** */
function showSecondScreen(info) {
  const box = $("secondScreen");
  const note = $("secondNote");
  if (info && info.idle_sec !== undefined) {
    note.textContent = `前の画面からの応答: ${info.idle_sec}秒前`
      + `　（${info.grace_sec}秒とだえると、自動で空きます）`;
  } else {
    note.textContent = "別のタブが開いています。";
  }
  box.hidden = false;
  // 下の画面は**消さずに止める**。消すと「何が起きたのか」が分からない
  document.querySelector(".wrap")?.setAttribute("inert", "");
  watchFirstScreen();
}

/**
 * 前の画面の応答を数え直す(統合 1.2.11)。以前は開いたときの秒数のままで、
 * 前の画面が閉じられても分からなかった。空いたら「もう一度ためす」で続けられると出す
 */
let firstScreenTimer = null;
function watchFirstScreen() {
  clearInterval(firstScreenTimer);
  firstScreenTimer = setInterval(async () => {
    if ($("secondScreen").hidden) { clearInterval(firstScreenTimer); return; }
    try {
      const st = await api.get("/api/screen/state");
      $("secondNote").textContent = st.busy
        ? `前の画面からの応答: ${st.idle_sec}秒前　（${st.grace_sec}秒とだえると、自動で空きます）`
        : "前の画面は閉じられました。「もう一度ためす」を押すと、この画面で続けられます。";
    } catch { /* 次に数え直す */ }
  }, 3000);
}

/** 前の画面に手放させて、開き直す。**2枚にするのではなく、移す。** */
async function takeOverScreen(button) {
  button.disabled = true;
  // **先に知らせる。** 開き直してから知らせると、開き直した先では
  // まだ前のタブが「居る」と返してくるので、いつまでも入れない
  tab.announceTakeOver();
  try { await api.post("/api/screen/take-over", {}); } catch { /* 開き直せば分かる */ }
  location.href = BASE + "/meisai";
}

/**
 * まだサーバへ送っていない入力の名前(統合 1.2.5)。統合画面の「終了」・窓の × ・
 * 外からの停止(ランチャー・stop.bat)の前に、統合画面が訊く。空なら閉じてよい。
 */
function unsaved() {
  const state = window.APP.state || {};
  const out = [];
  const lot = $("lotNo").value || "";
  if (lot && lot !== (state.lot_no || "")) out.push("LOT No(7桁になっていない入力)");
  if (editingZen) out.push("前工程実績数(「確定する」を押していない)");
  const weights = [];
  for (let jou = 1; jou <= (state.jou_su || 0); jou += 1) {
    const input = $(`weight${jou}`);
    if (input) weights.push([input.value, (state.weights || [])[jou - 1] || ""]);
  }
  if (weights.some(([v]) => String(v).trim() === "")
      && weights.some(([v, sent]) => String(v).trim() !== "" && String(v) !== String(sent))) {
    out.push("重量(空の丈があるので、まだ送っていない)");
  }
  if (state.edit) out.push(`No${state.edit.no}の出力済み修正(「修正して出力」を押していない)`);
  return out.concat(settings.unsaved()).map((name) => `梱包明細: ${name}`);
}
window.cptUnsaved = unsaved;

async function boot() {
  window.APP.screen = tab.tabId();
  wire();

  // **まず他のタブに直接きく。** タブの複製は名前ごと写されるので、
  // 名前の比較だけでは見分けられない
  if (await tab.anotherTabIsOpen()) { showSecondScreen(null); return; }

  try {
    await api.post("/api/screen/claim", { screen: window.APP.screen });
  } catch (err) {
    if (err instanceof ApiError && err.code === "screen_busy") {
      showSecondScreen(err.body || null);
      return;
    }
    showError(err);
    return;
  }

  tab.answerPings(lockScreen);
  render(window.APP.state);
  health.start({ onLost: lockScreen });
  // 新しい台帳があるか(統合 1.2.49)。開いたとき・前に戻ったとき・数分ごと
  checkNewLedger();
  setInterval(checkNewLedger, LEDGER_POLL_MS);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") checkNewLedger();
  });
}

boot();
