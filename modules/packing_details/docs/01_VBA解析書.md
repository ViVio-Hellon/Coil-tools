# VBA 解析書（移植前 現行仕様の確定）

対象の版: 梱包明細 VER 0.13.29(コイル梱包ツール 1.2.32)

対象: コイル梱包明細打ち出しシステム（Excel VBA。フォーム `frmCoilPacking`）

> **この文書の作り方（先に読む）**
>
> **VBA の元ソースはこのリポジトリに無い。** 移植元の単体リポジトリ（`packing-details-python-web`）にも無い。
> この文書は、移植したコードの注記（`VBA \`cmdKettei_Click\`` のように書いた箇所）・
> `docs/変更履歴.md`・`README.md`・試験から組み立てた。
> 変更履歴 VER 0.13.23 にも「VBA のソースはこの版を作るときに見ていない」とある。
>
> - 事実ごとに、根拠のコードの場所を書く。
> - 根拠の無いことは書かない。推測で埋めない。
> - 分からないところは `docs/03_不明点と要確認事項.md` に「要確認」として回した（本文では「→ 03 C-n」と書く）。
>
> 根拠の書き方: パスは `modules/packing_details/` からの相対。`strand_service.distribute` は
> `meisai/strand_service.py` の関数 `distribute`。`routes/meisai.py` は `app/routes/meisai.py`。
> 「移植元 0.13.1」は単体リポジトリ `packing-details-python-web` の VER 0.13.1
> （No指定・副番履歴クリアの注記は、そちらにだけ残っている）。

---

## 1. このシステムは何をしているのか

アルミコイルを梱包するときに貼る **梱包明細表** を作る、ライン端末用の Excel ツール（`README.md`「これは何をするものか」）。

- 1 本の母材ロットは丈（たて）方向に切り分けられ、さらに横方向に条へ割られる。
  1 ロットから `縦割数 × 横割数` 本のコイルができる。
- 各コイルは `丈番号-条番号`（例 `2-15`）で呼ぶ。ロット番号を前に付けたものが **副番**（例 `L5160Z0-2-15`）。
- 作業者は現物を積む順にコイルをクリックし、**その並びをそのまま紙に出す**。

```
ロット番号(7桁)を打つ
  → 仕掛台帳から ロット情報 / 受注番号 / 包装仕様NO を引く
  → 前工程実績数(= 縦割数 × 横割数)を丈へ配分して条番号を作る
  → 丈ごとの重量を入れる
  → 積む順にクリックする
  → 梱包明細表(Excel のシート)を作って印刷する
```

---

## 2. モジュール構成

### 2.1 規模

| 項目 | 内容 | 根拠 |
|---|---|---|
| 構成 | `frmCoilPacking` ＋ 標準 3・クラス 3 モジュール | 変更履歴 VER 0.1.0 |
| 行数 | 3,235 行 | 変更履歴 VER 0.1.0 |

### 2.2 フォーム

| 名前 | 役割 | 根拠 |
|---|---|---|
| `frmCoilPacking` | 画面のすべて。ロット・丈数・重量・条番号・積み上げ・出力・印刷 | `routes/meisai.py` 冒頭、`app/static/js/meisai.js` 冒頭 |

画面の部品は「ギミック」と呼ばれていた（注記の呼び名）。

| 呼び名 | 中身 | 根拠 |
|---|---|---|
| ギミックA | 前工程実績数の「修正する → 確定する」 | `routes/meisai.py` `set_zen_kotei`、`meisai.js` の `editingZen` |
| ギミックB | 条番号のラベル（丈ごとの並び） | `strand_service.py` 冒頭 |
| ギミックC | 積み上げのスロット | `strand_service.py` 冒頭、`routes/meisai.py`「積み上げ (VBA ギミックC)」 |

### 2.3 標準モジュール（注記に名前が出るもの）

注記に出てくる標準モジュール名は 3 つ。変更履歴の「標準 3」と数は合う。

| モジュール | 注記に出てくる手続き・定数 | 根拠 |
|---|---|---|
| `mDB` | `GetLotInfo` `GetWariSuFromShikakari` `GetJuchuBangouList` `GetOdrInfo` `SearchLotAll` `OpenAdo` `SafeStr`、定数 `PATH_仕掛_台帳` `PATH_梱包課共有仕掛` `DB_KONPO_SIKAKARI` | `meisai/lot_repo.py` 冒頭、`meisai/config.py` |
| `mSheet` | `RegisterFuban` `IsDuplicate` `ClearFubanList` `SaveSnapshot` `LoadSnapshot` `DeleteSnapshot` `ClearAllSnapshots` `Get/SetCurrentLotNo` `Get/Set/IncrementSeq` `IsPrinted` `SetPrinted` `BuildOneSheet` `AdjustLayout` `WriteData` | `meisai/history_repo.py` 冒頭、`meisai/report.py` 冒頭 |
| `mUtil` | 定数 `URL_HOUSOU` | `meisai/config.py` |

`mDB` / `mUtil` は定数を `Public Const` で持っていた（`meisai/config.py` 冒頭）。

### 2.4 クラスモジュール・型

| 名前 | 中身 | 根拠 |
|---|---|---|
| `clsWeightCtrl` | 重量欄の入力制御（`Txt_KeyPress` で数字以外を拒否） | `meisai.js`、`meisai_service.read_weights` |
| `clsOdrInfo` | 受注 1 件（受注番号・包装仕様NO ほか） | `lot_repo.OrderInfo` |
| （3 つ目） | **名前が注記に無い** → 03 C-1 | ― |
| `TLotInfo` | ロット情報（クラスか型かは注記に無い） | `lot_repo.LotInfo` |

### 2.5 フォームが持っていた状態

| VBA の変数 | 中身 | 根拠 |
|---|---|---|
| `mHandlersB` | 条番号ラベルの Collection | `strand_service.Board` |
| `mHandlersC` | 積み上げスロットの Collection | `strand_service.Board` |
| `mStrandsPerJou(i)` | 丈ごとの条数 | `strand_service.keys_for_jou` |
| `IsUsed` | 条番号ラベルの「使用済み」フラグ | `strand_service.Board.used` |
| `mHaikiMode` | 廃棄選択中か | `meisai.js` の `haikiMode` |
| `mRestoredSnapshot` | 途中経過から復元した直後か | `session.py` の `_restored` |
| `slotKeys` | 積み上げ順の条番号 | `strand_service.Board.stacked_keys` |

### 2.6 ワークブックのシート

| シート | 中身 | 根拠 |
|---|---|---|
| 隠しシート「副番履歴」 | 出力した副番・現在ロット番号（D1）・現在連番（E1）・印刷済（F1）・途中経過（G〜K 列。直近 5 ロット） | `meisai/schema.sql`「作業表」「明細出力状態」「明細スナップショット」 |
| 明細シート `<ロット番号>-No<N>`（例 `N3250Q0-No1`） | 出力 1 件 = 1 シート。シート名が重複しないことで、同じロットの同じ No を 1 枚に保っていた（同名なら上書き確認） | `meisai_service.Output.sheet_name`、`schema.sql` の `ux_shutsuryoku` の注記 |

副番の記録をシートのどの列に置いたかは注記に無い → 03 C-1。

---

## 3. 外部依存

### 3.1 仕掛台帳（Access ＋ ADO ＋ ACE）

VBA は取り込み元（Access）を **直接開いて** 引いていた（`lot_repo.py` 冒頭、`README.md`「VBA版との主な違い」）。

| 定数・設定 | 中身 | 根拠 |
|---|---|---|
| `mDB.PATH_仕掛_台帳` | `\\nlmsrvngy03\Read\【New】仕掛\台帳\` | `config.LOT_DB_DIR` の注記「VBA と同じ場所」 |
| `mDB.PATH_梱包課共有仕掛` | `\\nlmfangyshrd\各課共有\0130_日軽稲沢\梱包課\梱包課共有\仕掛` | `config.KONPO_DB_DIR` の注記「VBA と同じ場所」 |
| `DB_KONPO_SIKAKARI` | 梱包課共有の 1 ファイル（縦割数・横割数の出どころ） | `config.KONPO_DB_FILES` |
| 仕掛台帳のファイル | `SIKALOTNOW.accdb` 等（移植後は `NOW` の付かない `.sqlite3`） | `config.LOT_DB_FILES` の注記 |
| 接続の待ち | `ConnectionTimeout = 10秒` | `db.get_connection` の注記 |

`SIKALOTNOW.accdb` 以外のファイル名（受注・引当・LS4 の側）は注記に無い → 03 C-1。

### 3.2 包装仕様書の閲覧システム

| 定数 | 中身 | 根拠 |
|---|---|---|
| `mUtil.URL_HOUSOU` | 包装仕様書の閲覧画面の URL（移植後の既定は `config.URL_HOSO_SHIYOSHO`） | `config.py` |

VBA は包装仕様NO をクリップボードへ写してからこの URL を開き、利用者が貼り付けて検索していた（`config.py`、`meisai.js` の `lblHousouShiyou_Click` の注記）。

### 3.3 説明書

`cmdManual_Click` は共有フォルダの HTML（`\\nlmfangyshrd\…\説明書\梱包明細システム説明書.html`）を開いていた（`app/routes/settings.py`「説明書」）。

### 3.4 Excel（組版と印刷）

帳票は Excel のシートとして作り、`PrintOut` で刷っていた。Excel を「レイアウトエンジン＋プリンタドライバ」として使っていた（`meisai/report.py` 冒頭、`README.md`）。

---

## 4. 入力データ

### 4.1 画面で入れるもの

| 入力 | 決まり | 根拠 |
|---|---|---|
| ロット番号 | 7 桁（`txtLotNo.MaxLength = 7` / `If Len(...) <> 7`）。7 桁になった時点で自動検索 | `config.LOT_NO_LENGTH`、`lot_repo.search` |
| 前工程実績数 | ギミックA（修正する → 確定する）。確定時に数値かを確かめる | `routes/meisai.py` `set_zen_kotei` |
| 丈数 | 上限 `JOUSU_MAX` = 20 | `config.JOUSU_MAX` |
| 重量 | 丈ごとに 1 つ。入力欄が数字以外を拒否（小数点も不可）＝整数。±1・±10 のスピン（`spn1` / `spn10`） | `meisai.js`、`meisai_service.read_weights` |
| 積み条数 | 上限 `STACK_LIMIT` = 15 | `config.STACK_LIMIT` |
| No指定 | チェック（`chkSpecifyNo`）と番号。出力後にチェックを外す | 移植元 0.13.1 の `meisai.js`。VER 0.13.23 で外した |

### 4.2 仕掛台帳から引くもの（`mDB`）

| 手続き | 引き方 | 根拠 |
|---|---|---|
| `GetLotInfo` | `SELECT TOP 1 ... FROM 仕掛 WHERE ﾛｯﾄ番号 = ? AND ｵｰﾀﾞｰ板丈 = '0'`。ロット情報が取れたときだけ、中から `GetWariSuFromShikakari` を呼ぶ | `lot_repo.find_lot` |
| `GetWariSuFromShikakari` | 梱包課共有の仕掛から `当工程設計_縦割数` / `当工程設計_横割数`。エラー時・該当なし時は両方 0 | `lot_repo.find_warisu`、`schema.sql`「仕掛当工程」 |
| `GetJuchuBangouList` | 引当から受注番号を `Collection` へ順に足し、既にあれば飛ばす（初出順・重複なし）。並びは取り込み元の物理順 | `lot_repo.find_order_nos` |
| `GetOdrInfo` | 受注番号 → 包装仕様NO。`r.juchuBangou` を先に入れてから検索するので、見つからなくても受注番号は返り、包装仕様NO は空 | `lot_repo.find_spec_no` |
| `SearchLotAll` | ロット番号 1 本で上の 3 つをまとめて引く（3 つの戻り値） | `lot_repo.search`、`lot_repo.SearchResult` |

`GetLotInfo` が使う列は 10 列と抽出条件の 1 列（`import_specs.py` の `IMPORT_SPECS` の注記）。

| 列（取り込み元の名前） | 画面での使いみち |
|---|---|
| `用途ｺｰﾄﾞ` `用途名` | 用途 |
| `製造材質` `製造調質` | 材質・調質 |
| `製造板厚` `製造板幅` `ｵｰﾀﾞｰ板厚` `ｵｰﾀﾞｰ板幅` | 寸法（表示書式は §6.12） |
| `設計_設備ｺｰｽ` `実績_設備ｺｰｽ` | 設備コース |
| `ｵｰﾀﾞｰ板丈` | 抽出条件（`= '0'`） |

実データについて分かっていること（移植のときの調べ。`schema.sql` の注記）:

- 仕掛ロットのロット番号は一意ではない（13,624 行に対し 6,677 ロット）。
- オーダー板丈 0 のロット 3,113 件のうち 2,987 件は複数行あるが、上の 10 列の内容はどの行でも同じだった。どの行を採っても結果は変わらない。
- 引当側の受注番号のうち、仕掛受注にあるのは 81%。残りは包装仕様NO が空になる。
- 梱包課共有の仕掛（LS4LOT）は LS4 ラインの現在仕掛だけで、589 ロット。

---

## 5. 処理フロー（画面の流れ）

### 5.1 ロット番号を入れる（`txtLotNo_Change`）

根拠: `session.open_lot` / `session.close_lot` の注記、`routes/meisai.py` `open_lot`。

```
7桁に達した
  ├ 前のロットと違う
  │   ├ 直近の出力が未印刷なら確認(MsgBox vbYesNo)  ← 未印刷のまま切り替えると出力が消えるため
  │   └ 前ロットの片付け: DeleteLotSheets → ClearFubanList → ResetSeq → DeleteSnapshot
  └ 検索: SearchLotAll → DisplayLotInfo → RefreshSheetList → RestoreSnapshot
```

- 片付けの順序はこのとおり。`ClearFubanList` では連番は戻らないので、`ResetSeq` を明示的に呼ぶ（VBA のコメントが同じことを言っている。`session.close_lot`、`history_repo.clear_fuban`）。
- 確認に「いいえ」と答えたときの VBA の動きは注記に無い → 03 C-4。

### 5.2 ロット情報を出す（`DisplayLotInfo`）

- 丈数に縦割数を入れる。**1〜20 なら入れる、21 以上は警告のみ、0 以下は何もしない**（`session.open_lot` の注記）。
- 板厚・板幅を `Format` で `0.000` / `0.0` に整える（`lot_repo.LotInfo` の `*_text`）。
- 受注番号の一覧を出す。受注番号を押すと包装仕様NO が切り替わる（`lstJuchuList_Click`。`meisai.js`）。
- 包装仕様NO が空ならリンクを張らず、黒いまま（`meisai.js`）。

### 5.3 前工程実績数を直す（ギミックA `cmdToggleEditA_Click`）

「修正する → 確定する」のトグル。確定時に数値かどうかを確かめる（`routes/meisai.py` `set_zen_kotei`）。
LS4LOT に載っていないロットは前工程実績数が 0 になるので、手で入れるために使う（`README.md`「LS4LOT は代替できません」）。

### 5.4 丈数を変える（`txtJouSu_Change`）

丈数が変わると、**重量欄・条番号・積み上げ・廃棄をすべて作り直す**（廃棄は `ClearHaiki`）。根拠: `routes/meisai.py` `set_jou_su`。

### 5.5 条番号を出す（`cmdShowJouTai_Click`）

確かめる順（`routes/meisai.py` `build_strands` の注記「VBA が確かめていることを同じ順で確かめる」）:

1. 復元直後（`mRestoredSnapshot`）なら上書きしてよいか尋ねる
2. 丈数が入っているか
3. 前工程実績数が数値で、0 より大きいか
4. 重量がすべて入っているか

通ったら配分を計算し（§6.3）、`BuildGimmickB` でラベルを作り直す。ラベルを作り直すので **`IsUsed` は全部消える**。廃棄だけは `ReapplyHaiki` で塗り直す（`Board.build_strands`）。最後に `SaveCurrentSnapshot`。

### 5.6 反転（`cmdReverse_Click`）

ラベルの `Left` だけを入れ替える。条番号も積み上げ済みの状態もそのまま。現物の並び向きと画面の向きを合わせるための表示補助（`Board.reverse`）。

### 5.7 廃棄（`cmdHaiki_Click` ＋ `GimmickB_Click` の廃棄モード）

`cmdHaiki_Click` で `mHaikiMode` を切り替える（`meisai.js`）。廃棄モードで条番号を押したときの分岐（`Board.toggle_discard`）:

| 押したコイル | 結果 |
|---|---|
| すでに廃棄 | 解除する |
| 未使用 | 廃棄にする |
| 使用済み | 何もしない（使用済みは廃棄不可） |

廃棄の増減のたびに `SaveCurrentSnapshot`（`session.save`）。

### 5.8 積み条数（`txtTsumiJouSu_Change`）

`ResetGimmickBUsed` で **いま割り当たっているものだけ** を戻してから、スロットを作り直す（`Board.set_stack_max`）。

### 5.9 積む・外す（`GimmickB_Click` / `GimmickC_Click`）

| 操作 | VBA の動き | 根拠 |
|---|---|---|
| 条番号を押す（通常モード） | 最も若い空きスロットに入る。スロット 1 が最下段 | `Board.stack` |
| 使用済み・廃棄済みを押す | 黙って無視（早期 Exit） | `routes/meisai.py` `stack`、`meisai.js` |
| スロットを押す | そのスロットだけ空にする。**詰め直さない**（次に積むとそこへ入る） | `Board.unstack` |
| 空のスロットを押す | 何も起きない | `routes/meisai.py` `unstack` |
| 全スロットが埋まったか | `CheckAllFilled` で `cmdKettei.Enabled` を切り替える（埋まるまで出力できない） | `Board.is_complete` |

### 5.10 出力（`cmdKettei_Click`）

手順（`meisai/meisai_service.py` 冒頭。移植元 0.13.1 の版で並べ直したもの）:

| # | 手順 | VBA |
|---|---|---|
| 1 | 重量を読み直す | `ReadWeights` |
| 2 | 積み上げ順に条番号を並べる | `slotKeys` |
| 3 | 副番の重複を確かめる（No指定時はスキップ） | `IsDuplicate` |
| 4 | No を決める | `IncrementSeq`（No指定なら指定値） |
| 5 | 明細を作る | `CreateMeisaiSheet` ＋ `WriteData` |
| 6 | 副番を登録する（すでにあるものは足さない） | `If Not IsDuplicate Then RegisterFuban` |
| 7 | 印刷済フラグを OFF | `SetPrinted False` |
| 8 | 積み上げだけリセット | `ClearGimmickC` ＋ `BuildGimmickC` |

- 重複があれば「以下の副番はすでに出力済みです。このまま出力しますか？」を `MsgBox ... vbYesNo` で尋ね、「はい」なら続ける（`meisai_service.output`、`history_repo.duplicates`）。続けた場合、重複した副番は登録をスキップする（`schema.sql`「副番履歴」）。
- 手順 8 で **ギミックB の `IsUsed` はそのまま**。VBA のコメントは「Cスロットのみ。Bは出力後もそのまま残す」。同じコイルを次の梱包へ二重に積ませないための仕掛け（`meisai_service.reset_stack`）。
- 出力のあと `SaveCurrentSnapshot`（`session.save`）。

### 5.11 印刷（`cmdPrint_Click`）

`SetPrinted True` を立てる（`routes/meisai.py` `mark_printed`）。シートは `PrintOut` で刷る（`README.md`）。
**どのシートを何枚刷ったか**（一覧で選んだ 1 枚か、複数か）は注記に無い → 03 C-2。

### 5.12 出力済みの一覧（`lstSheetList` / `RefreshSheetList`）

そのロットの明細シートを並べる。**末尾を自動で選ぶ**（`meisai.js`「最新を選んでおく」、`meisai_service.list_outputs`）。

### 5.13 No指定（`chkSpecifyNo`）― VER 0.13.23 で外した

移植元 0.13.1 の実装と注記（`meisai_service.output`、`meisai.js`、`docs/説明書.md`）による。

- 番号を自分で決めて出す。刷り直しに使っていた。**副番の重複は確かめない**。
- 指定した No が既にあれば上書きの確認、現在の連番を超えていれば作成の確認を挟む（確認文言が VBA と同じかは注記に無い）。
- 指定した No が現在連番より大きければ、現在連番を引き上げる（`SetCurrentSeq`）。
- 出したら No指定のチェックを外す（`chkSpecifyNo.Value = False`）。

### 5.14 副番履歴クリア（`cmdClearDup_Click`）― VER 0.13.23 で外した

移植元 0.13.1 の `routes/meisai.py` `clear_history`・`strand_service.Board.release_all_used` の注記による。

- 副番履歴と途中経過（`ClearAllSnapshots`）を消す。
- 画面の使用済みと積み上げをほどく（`ResetAllGimmickBUsed`）。VBA のコメントは「条番号・積み上げをリセット（キャンバスは残す）」。廃棄は残る。
- 連番は戻さない。

### 5.15 出し直しの手段（VBA の場合）

VBA には出力済みを直す専用のボタンは無く、「No指定」（重複を確かめずに同じ番号へ上書き）で出し直していた。
出したコイルは使用済みのままほどけないので、積み直すには「副番履歴クリア」か「条番号表示」で使用済みをほどくしかなく、
そのとき **ほかの No のコイルまでほどけていた**（変更履歴 VER 0.13.23、`docs/説明書.md` §4）。

---

## 6. 判定ロジック・計算式（移植の最重要部）

`README.md` の RULE-01〜10 が「VBA と結果が変わってはいけない箇所」。

### 6.1 前工程実績数（RULE-01）

```
前工程実績数 = 縦割数 × 横割数      ' ZenKoteiJissekiSu
```

根拠: `lot_repo.LotInfo.zen_kotei_jisseki_su`。縦割数・横割数は LS4LOT（`当工程設計_*`）から採る。LS4LOT に無いロットは 0（§4.2）。

### 6.2 丈数の自動セット（RULE-02）

丈数（初期値）= 縦割数。1〜20 なら入れる、21 以上は警告のみ、0 以下は何もしない（`session.open_lot`）。

### 6.3 条数の配分（RULE-03）

`cmdShowJouTai_Click` の配分（`strand_service.distribute` の注記に書き写された形）:

```vb
base = total \ jou_su        ' 整数除算(切り捨て)
rem  = total Mod jou_su
For i = jou_su To 1 Step -1
    条数(i) = base
    If rem > 0 Then 条数(i) = 条数(i) + 1 : rem = rem - 1
Next
```

**余りは丈番号の大きいほうから配る。**

| 前工程実績数 / 丈数 | 配分（丈1, 丈2, …） |
|---|---|
| 30 / 2 | 15, 15 |
| 31 / 2 | 15, 16 |
| 5 / 3 | 1, 2, 2 |

- 試験は、この手順をそのまま書き下した参照実装と、0〜299 本 × 1〜20 丈の全組合せで突き合わせている（`tests/test_strand_service.py`）。
- 実データでは余りが出ない。前工程実績数 = 縦割数 × 横割数、丈数 = 縦割数なので必ず割り切れる。余りが出るのは丈数か前工程実績数を手で直したときだけ（変更履歴 VER 0.1.0）。

### 6.4 条番号（RULE-04）

- 条番号キーは `i & "-" & k`（例 `2-15`）（`strand_service.key_of`）。
- `BuildGimmickB` は `For k = mStrandsPerJou(i) To 1 Step -1` で左から置くので、**左端が最大の条番号**（`strand_service.keys_for_jou`）。
- 0 条の丈は `If jouCount < 1 Then GoTo ContinueJou` で飛ばす（`ReDim positions(1 To 0)` のエラーを避けるため）。

### 6.5 積み上げ（RULE-05）

最も若い空きスロットへ入れる。外しても詰め直さない（§5.9）。

### 6.6 重量（RULE-06）

- 読み方は `CLng(Val(ctl.Value))`。数値でなければ「丈Nの重量が未入力です。」で止める（`meisai_service.read_weights` の注記）。
- 重量は **丈単位**。行の重量は副番キーの丈番号で引く（`WriteData` の `weights(jouPart)`）。同じ丈のコイルはみな同じ値（`Board.weight_of`、`Output.weight_of`）。

### 6.7 副番と重複（RULE-07）

- 副番キーは `lotNo & "-" & slotKeys(i)`（例 `N3250Q0-2-15`）（`history_repo.fuban_key`）。
- 出力のたびに `IsDuplicate` で 1 件ずつ確かめ、警告文を組み立てる（`history_repo.duplicates`）。
- 登録は `If Not IsDuplicate Then RegisterFuban`。すでにあるものは足さない（`history_repo.register_fuban`）。
- 副番履歴はロットを切り替えると全部消える（`ClearFubanList`。§5.1）。

### 6.8 連番（RULE-08）

| 場面 | 連番 | 根拠 |
|---|---|---|
| 出力 | `IncrementSeq` で +1（`GetCurrentSeq` で読む） | `history_repo.next_seq` |
| No指定で出力 | 指定値。現在連番より大きければ `SetCurrentSeq` で引き上げる | 移植元 0.13.1 `meisai_service.output` |
| ロット切替 | `ResetSeq` で 0 に戻す。**ロット切替時だけ** | `history_repo.reset_seq` |
| 副番履歴クリア | 戻さない | `history_repo.clear_fuban`、移植元 0.13.1 `clear_history` |

### 6.9 ロット抽出（RULE-10）

`ﾛｯﾄ番号 = ? AND ｵｰﾀﾞｰ板丈 = '0'` の **先頭 1 行**（`SELECT TOP 1`）（`lot_repo.find_lot`）。

### 6.10 使用済み（`IsUsed`）の付き方・外れ方

| 場面 | `IsUsed` | 根拠 |
|---|---|---|
| 積む | 立つ | `Board.stack` |
| 出力 | 立ったまま（スロットだけ空く） | `meisai_service.reset_stack` |
| 積み条数を変える | いまスロットに割り当たっているものだけ戻る（`ResetGimmickBUsed`） | `Board.set_stack_max` |
| 条番号表示 | 全部消える（ラベルを作り直すため） | `Board.build_strands` |
| 途中経過から復元 | 立て直す。ただしスロットは作り直さない | `session._restore`、`Board.restored_used` の注記 |
| 副番履歴クリア | 全部ほどく（`ResetAllGimmickBUsed`） | 移植元 0.13.1 `Board.release_all_used` |

復元で立て直した使用済みは、スロットに無いので押しても反応しない。次の `txtTsumiJouSu_Change` で `ResetGimmickBUsed` が回っても、スロットが空なので戻らない。解除されるのは条番号の再表示のときだけ（`Board.restored_used` の注記）。

### 6.11 廃棄の決まり

- 使用済みは廃棄にできない（§5.7）。
- 条数が変わって存在しなくなった条番号の廃棄は落とす。ハンドラが取れないものを新しい Collection に入れ直さないため（`ReapplyHaiki`。`Board.reapply_discards`）。
- 丈数を変えると廃棄は全部消える（`ClearHaiki`。§5.4）。

### 6.12 値の読み方・表示書式

| VBA | 決まり | 根拠 |
|---|---|---|
| `Val()` | 先頭から数値として読める分だけ読む。`Val("12abc")` は 12、`Val("")` と `Val("abc")` は 0。全角数字は読めない（`Val("１２")` は 0） | `lot_repo.val`、`import_specs.vba_val` |
| `Format(x, "0.000")` / `"0.0"` | 板厚は小数 3 桁、板幅は小数 1 桁。半分は 0 から遠い側へ（104.25 → 104.3） | `lot_repo.fmt`、`LotInfo.*_text` |
| `CLng(Val(Trim(...)))` | 途中経過の数値の読み方。空や数字以外は 0 | `history_repo._split_ints` |
| `CLng(Val(ctl.Value))` | 重量の読み方 | `meisai_service.read_weights` |

`Format` の丸めを「半分は 0 から遠い側」とする根拠は、移植版の注記と試験（`tests/test_rounding_vba.py`）。VBA を実際に動かして比べた記録は無い。

### 6.13 範囲と上限（RULE-09）

| 項目 | 値 | VBA | 根拠 |
|---|---|---|---|
| ロット番号 | 7 桁 | `txtLotNo.MaxLength = 7` | `config.LOT_NO_LENGTH` |
| 丈数 | 1〜20 | `JOUSU_MAX` | `config.JOUSU_MAX` |
| 積み条数 | 1〜15 | `STACK_LIMIT` | `config.STACK_LIMIT` |
| 明細の行 | 15 行（A6:F20） | `For r = 6 To 20` | `config.SHEET_ROWS` |
| 途中経過 | 直近 5 ロット | `SNAP_MAX` | `config.SNAPSHOT_KEEP` |

`STACK_LIMIT` と紙の行数は同じ 15 だが、別の理由で決まっている（1 梱包に積める本数 / 紙に収まる行数。`config.py`）。

---

## 7. 状態の保存（隠しシート「副番履歴」）

### 7.1 副番履歴

| 手続き | 動き | 根拠 |
|---|---|---|
| `RegisterFuban` | 副番を 1 行足す | `history_repo.register_fuban` |
| `IsDuplicate` | その副番がもうあるか | `history_repo.is_duplicate` |
| `ClearFubanList` | 全件消す。連番は戻さない | `history_repo.clear_fuban` |

### 7.2 出力状態（D1 / E1 / F1）

| セル | 中身 | 使いみち・変わるとき | 根拠 |
|---|---|---|---|
| D1 | 現在ロット番号 | ロット切替の判定に使う | `schema.sql`「明細出力状態」 |
| E1 | 現在連番 | 出力のたびに +1。ロット切替で 0 | 同上、`history_repo.reset_seq` |
| F1 | 印刷済 | 出力で 0、印刷で 1 | 同上 |

### 7.3 途中経過（G〜K 列）

| 項目 | 内容 | 根拠 |
|---|---|---|
| 中身 | 前工程実績数・丈数・重量・条数・使用済・廃棄。重量・条数・使用済・廃棄はカンマ区切りの文字列 | `schema.sql`「明細スナップショット」 |
| 件数 | 直近 5 ロット（G〜K 列という物理的な制約） | `history_repo.save_snapshot` |
| 書く時点 | `SaveCurrentSnapshot` を呼ぶ 3 つ: 条番号表示・廃棄の増減・出力 | `session.save` |
| 手続き | `SaveSnapshot`（同じロットは上書き）・`LoadSnapshot`・`DeleteSnapshot`（1 ロット）・`ClearAllSnapshots`（全部。副番履歴クリアと連動） | `history_repo.py` |

### 7.4 復元（`RestoreSnapshot`）

確かめる順（`session._restore`）:

1. 条数の要素数と丈数が合わないなら **何もしない**（合わない状態で描くと、あるはずのコイルが出ない）
2. 前工程実績数は途中経過の値で **上書きする**（台帳が更新されていても、作業中の数を勝手に変えない）
3. 丈数が範囲外なら諦める

- 使用済みフラグだけを復元し、積み上げスロットは空のまま出す（積み条数は入れ直す）。
- 復元した直後は `mRestoredSnapshot` が立ち、条番号表示の前に確認を挟む（§5.5）。

---

## 8. 出力内容（明細シート）

### 8.1 紙面の構成（`BuildOneSheet` ＋ `WriteData`）

根拠: `meisai/report.py` 冒頭「紙面の構造（実ブックの A2:F21 = 1面ぶん）」と `build_slip`。

| 行 | セル | 中身 |
|---|---|---|
| 1 | A1:D1 | `<ロット番号>-No<N>  梱包明細表`（`"-No1" & "  梱包明細表"`。空白は 2 つ） |
| 1 | E1:F1 | `NLM.NAGOYA.QA`（赤。`MergePut ws, 1, 5, 1, 6, "NLM.NAGOYA.QA", True`） |
| 2〜4 | A2:A4 | 「種類質別 / 寸法 / 検番」（3 行を 1 つに結合して中央寄せ） |
| 2〜4 | B2:F4 | 罫線なし（小ラベル貼付エリア。右端だけ縦線を引く） |
| 3 | D3 | 「小ラベル添付」（`Cells(3, 4)`。橙） |
| 5 | A5:C5 / D5:F5 | 「コイル副番」 / 「重量(Kg)」 |
| 6〜20 | A〜C / D〜F | 積んだ順の副番 / 重量。15 行。空行も枠だけ引く |

- 種類質別・寸法・検番の値は印字しない。B2:F4 は現物のラベルを手で貼る場所。
- 副番も重量も `MergePut` で A〜C・D〜F を結合して `xlCenter` にしてから値を入れる（中央）。見出し・「小ラベル添付」・左の見出しも `xlCenter`（`tests/test_meisai_service.py` の注記）。
- 色は `RGB(255, 0, 0)`（E1:F1）と `RGB(255, 102, 0)`（D3）（`report.COLOR_QA` / `COLOR_SMALL_LABEL`）。
- 文字の大きさは、`report.py` 冒頭に D3 が 9pt とあるだけ。ほかの欄の大きさは注記に無い → 03 C-3。
- 重量の表示は、移植版では `<丈の重量を小数1桁>Kg`（例 `2502.0Kg`。`Output.rows`）。VBA の書式そのものは注記に無い → 03 C-3。

### 8.2 `WriteData` の細かい決まり

- 不正なキー（`-` が無い・先頭が空）は飛ばしつつ **行だけは進める**（`Output.rows`）。
- 16 本目以降は **黙って捨てる**（`meisai_service.output` の注記）。積み条数の上限が 15 なので、ふつうは起きない。

### 8.3 ページ設定（`PageSetup`）

根拠: `report.py` 冒頭、変更履歴 VER 0.5.0。

```
.Orientation = xlLandscape / .PaperSize = xlPaperA4 / 余白 0
.FitToPagesWide = 1 / .FitToPagesTall = 1 / .Zoom = False
(中央寄せの指定なし → 左に寄る)
```

- `FitToPages` は **縮めるだけで引き伸ばさない**。1 面（A1:F20）は等倍のまま A4 横の **左半分** に出て、右半分が空く。
- 現場は右半分を切り取って使う。A4 1 枚に 2 枚並べる運用はしていない（現場に確認済み。`report.py`）。

### 8.4 寸法（実ブックの値）

根拠: `report.py` の `_COL_UNITS` / `_ROW_*`、`README.md`。

```
行の高さ  23.25 + 42×3 + 18.75 + 28.5×15 = 595.5pt  (= A4 の高さ)
列の幅    10.375 + 12.375 + 12×4 = 70.75             (= A4 横の幅の半分)
```

`AdjustLayout` が何をしていたか（この寸法を当てていたのか）は注記に無い → 03 C-3。

> 過去の誤り: VER 0.2.0 では「VBA は A4 いっぱいに引き伸ばしていた」とし、実ブックが A〜F / H〜M の 2 面で保存されていたことから 2 面付けにしていた。**どちらも誤り**だった（変更履歴 VER 0.5.0）。

---

## 9. 確認（`MsgBox ... vbYesNo`）

注記から VBA が尋ねていたと分かるもの。

| 場面 | 何を尋ねるか | 根拠 |
|---|---|---|
| ロット切替 | 直近の出力が未印刷のとき、切り替えてよいか | `routes/meisai.py` `open_lot` |
| 出力 | 副番がすでに出力済みのとき、このまま出力するか | `meisai_service.output` |
| 条番号表示 | 復元直後のとき、上書きしてよいか | `session.py` の `_restored`、`routes/meisai.py` `build_strands` |
| No指定の出力 | 既にある No を上書きするか | 移植元 0.13.1 `meisai_service.output`、`schema.sql`「同名なら上書き確認」 |

---

## 10. 黙って何もしない操作・エラー処理

| 操作 | VBA | 根拠 |
|---|---|---|
| 使用済み・廃棄済みのコイルを積もうとする | 黙って無視 | `routes/meisai.py` `stack` |
| 空のスロットを押す | 何も起きない | `routes/meisai.py` `unstack` |
| 使用済みを廃棄にしようとする | 黙って無視 | `routes/meisai.py` `discard` |
| LS4LOT にロットが無い・読めない | 縦割数・横割数を 0 にして続ける | `lot_repo.find_warisu` |
| 受注が見つからない | 包装仕様NO を空で返す | `lot_repo.find_spec_no` |
| 途中経過の数値が読めない | 0 として読む | `history_repo._split_ints` |

ログは `DebugLog` に出していた（`README.md`、`meisai/logging_utils.py`）。

---

## 11. ハードコードの値・定数（まとめ）

| VBA | 値 | 移植先 |
|---|---|---|
| `txtLotNo.MaxLength` | 7 | `config.LOT_NO_LENGTH` |
| `JOUSU_MAX` | 20 | `config.JOUSU_MAX` |
| `STACK_LIMIT` | 15 | `config.STACK_LIMIT` |
| `For r = 6 To 20` | 15 行 | `config.SHEET_ROWS` |
| `SNAP_MAX` | 5 | `config.SNAPSHOT_KEEP` |
| E1:F1 の文字 | `NLM.NAGOYA.QA` | `qa_mark.DEFAULT`（設定で変えられるようにした） |
| `ConnectionTimeout` | 10 秒 | `db.py`（手元 DB は 4 秒、取り込み元は 10 秒に分けた） |
| `mDB.PATH_仕掛_台帳` ほか | §3.1 | `config.py`（設定で変えられるようにした） |
| `mUtil.URL_HOUSOU` | §3.2 | `config.URL_HOSO_SHIYOSHO`（環境変数で変えられる） |

**このアプリの VBA にはパスワードが無かった。** 管理者パスワードは Web 版で足したもの（`meisai/admin_password.py` 冒頭）。

---

## 12. 不明点・要確認事項

`docs/03_不明点と要確認事項.md` に一覧する。この文書から回したもの:

| # | 内容 |
|---|---|
| C-1 | クラスモジュールの 3 つ目の名前・副番履歴シートで副番を置いた列・仕掛台帳の Access ファイル名（`SIKALOTNOW.accdb` 以外） |
| C-2 | `cmdPrint_Click` がどのシートを何枚刷っていたか |
| C-3 | 文字の大きさ（D3 以外）・重量の書式・`AdjustLayout` の中身 |
| C-4 | ロット切替の確認に「いいえ」と答えたときの VBA の動き |

---

## 13. 移植方針（要約）

- **業務上の結果・判定・計算・出力内容は同じ。変えたのは実現方法だけ**（`README.md`）。
- VBA のボタン・イベントと API を 1 対 1 にする（`routes/meisai.py` 冒頭）。対応は `docs/02_移植対応表.md`。
- 確かめる順番も VBA のまま（条番号表示の 4 つ・出力の手順・ロット切替の片付け）。
- VBA が黙って無視する操作は、黙って無視する。
- VBA の確認（`MsgBox ... vbYesNo`）は、まず確認事項を返し、了解されたら `confirm` 付きで呼び直す形にする。**何を尋ねるかの判断は VBA と同じ**（`meisai_service.py` 冒頭）。
- VBA のままだと現場で困る所は、理由を書いて変えた（No指定・副番履歴クリアを外した、紙の端から 5mm、16 行目以降は断る など）。一覧は `docs/03_不明点と要確認事項.md` §0 と、リポジトリ直下の `docs/ツール文書/13_梱包明細_思想とモード.md` §1.1。
