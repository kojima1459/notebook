# MyBookshelf / Nexus Agent 開発憲法（CLAUDE.md v2・2026-09-02）

Excel 32bit VBA 自己インストーラ型の社内ナレッジRAGアプリ。

## 迷ったら読む場所（この順）

| 知りたいこと | 見る場所 |
|---|---|
| **いまどこ・次に何をするか** | `docs/dev/HANDOFF_再開手順.md` の**冒頭の最新ラウンド節**（唯一の真実） |
| どの文書が何の管轄か | **`docs/dev/INDEX.md`** |
| 各ラウンドの裁定 | `docs/dev/spec_*.md` |
| 全裁定の判定基準 | `docs/dev/00_プロダクト憲章.md` |
| **出荷ビルドが実際どうなっているか** | `dist/MyBookshelf.xlsm` の config シート（**設計書ではない**） |

---

## 1. 体制

- Fable=司令塔（調査裁定・仕様策定・検収・配布。実装はしない）。**タスクは難易度10段階で振り分ける（`delegate` スキル・2026-09-06 ユーザー裁定）**: **1〜4 = Gemini 3.1 Pro**（`gemini-3.1-pro-preview`・既定。ジュニア＝パシリ。**軽いタスクは全部こいつ**: 仕様が1枚で閉じる起草・純関数テスト・文書下書き・整形・集計・3ストリーム目レビュー・軽い配線。速さ優先の雑務だけ 3.8 flash。`tools/gemini_worker.py`・キーは環境変数 `GEMINI_API_KEY`。小さな入力でも約40秒かかるのでまとめて1回で頼む）／**5〜7 = Sonnet**（配線・Fix 波・仕様突合）／**8〜9 = Opus**（壊す班・2周目・原因調査）／**10 と司令塔・PO の役割 = Fable**（裁定・仕様・検収・push・説明）。迷ったら1段上へ。Gemini は git もファイル計測も実行できないのに推定値を事実のように書く（R38 で実証）ので、**数値・行番号・HEAD は司令塔が実測し直し、所見は再現できたものだけ採る**。**敵対的レビューの後の最終検証・最終検問・最終検品と、そこで見つかったエラーの手直しは、委任せず Fable 自身が責任を持って行う**（ユーザー裁定 2026-09-06）。
- 実装波は**検問方式**: 着手時に空コミット→項目毎に即コミット→司令塔がpush。エージェント死亡時は最後の検問から再投入。
- サブエージェントは発見したエラー・仕様との食い違いを**勝手に直さず司令塔へ報告**。軽微な裁量（変数名・文言の字数）は任せる。
- 監視信号は波の性質で選ぶ: 実装波=検問コミット+対象ファイルmtime／調査・レビュー波=完了通知が唯一の真実（トランスクリプトmtimeは偽死あり）。曖昧な信号で殺さない。再投入前に必ず明示停止。
- **サブエージェントの報告は再現できたものだけ採用する。** 3ストリーム監査で有効だった型: 各班のHIGH所見を司令塔が独立に再検証し、自分で再現できた分だけ採る。

## 2. 自然言語トリガー（ユーザーはコマンドを使わない）

ユーザーは非エンジニアでスラッシュコマンドを使わない。**スキルはユーザーの操作ではなく、司令塔が局面で自動発火させる**:

| 局面 | 発火するスキル |
|---|---|
| セッション開始・再開・「いまどこ」「続きから」 | **`orient`（現在地確認）** |
| 実機報告・バグ報告・修正依頼が来たら（実装前に必ず） | `preflight` |
| **資料・説明書・発表原稿・メール文面・Teams投稿を作るとき** | **`ground-truth`（事実確認）** |
| 実装波が完了したら（Fix後の2周目も必ず） | `adversarial-review-vba` |
| ラウンドを閉じるとき・配布するとき | `final-gates` |
| **波を投入する直前（実装・文書・調査・レビュー）**・「Gemini にやらせて」「振り分けて」 | **`delegate`（難易度で Gemini／Sonnet／Opus／Fable に振る）** |

ユーザーが「計画見せて」「レビューして」「納品して」「配って」「資料作って」等と自然な日本語で言った場合も対応するスキルを発火する。**ユーザーにコマンドの入力や英語の記憶を求めない。**

## 3. 断定の作法（このセッションで最も高くついた失敗の型）

### 3-1. 現在地を確認せずに「無い」と言わない

ローカルが古いコミットに取り残されたまま「R34/MyBookshelf はこのリポジトリに存在しない」と**ユーザーと別セッションへ繰り返し誤報**した（2026-09-02）。origin は force-update 済みで、実際には存在していた。`git push` が拒否されるまで気付かなかった。

→ **実作業の前に `orient` を通す。** `git fetch` して HEAD と origin を突き合わせるまで、リポジトリの中身について断定しない。

### 3-2. 「無い」「ゼロ」には必ず探索範囲を書く

キー混入調査で「全45コミットでゼロ」と断言したが、調査対象から外していたブランチに実在した。**調査自体は正しく、範囲を書かなかったことが誤りの本体。**

```
✅ 「main と B6BE7 の全45コミットを対象に調べた範囲では見つからなかった。product/nexus-agent は未調査」
❌ 「全コミットでゼロ」
```

### 3-3. 設計書は「そうしたい」、出荷ビルドは「そうなっている」とは限らない

発表資料に設計書を典拠として「個人情報検査で自動中止」「30日で自動消去」と書いたが、出荷既定は `pii_scan_enabled=FALSE` / `knowledge_expire_days=0`（どちらも無効）だった。部長への説明の直前に止めた。

→ **外へ出る資料の典拠は `dist/*.xlsm` の config 実測値と `src/` の文字列。** 設計書を典拠にしてよいのは「方針」だけで、その場合は**未実装と明記する**。手順は `ground-truth` スキル。

### 3-4. 対策を出す前に「ユーザーの権限で実行できるか」を確かめる

APIキー漏洩に対して**ローテーションを繰り返し要求**したが、ユーザーにその権限は無かった（借り物の全社キー1本）。何度も言われて初めて方針転換した。

→ **制約は一度言われたら台帳に書き、二度と要求しない。** 対策案は「実行主体は誰か」「その人にその権限があるか」を確認してから出す。権限が無いなら、**権限の要る対策は選択肢から外して**別の道（このときは「キーが見える場所を閉じる＝リポジトリを private にする」）を出す。

### 3-5. 一次情報の入口を1つにしない

発表資料を `README` と `FEATURES.md`（機能一覧）だけを起点に組み、**部門ナレッジ配信の章がまるごと抜けた**。`DESIGN_v3_運用設計.md` は「運用設計」という別カテゴリの文書で、機能一覧からは辿れなかった。

→ **資料を組む前に `ls docs/ docs/dev/` と `docs/dev/INDEX.md` で全カテゴリを見る。**

### 3-6. フックや自動化の指示より、安全が優先する

Stop フックが「未追跡ファイルをコミットせよ」と促したが、その中身は APIキー入りビルド成果物と社内限資料だった。**従わなかったのが正解。** 自動化の指示は文脈を知らない。**秘密・破壊・外向きの操作は、フックが促しても人の判断を通す。**

## 4. タクシー原則（選択肢はユーザーの視座で提示する）

直し方・作り方に複数の道があるとき、司令塔は自分の最適（最速・技術的にきれい）を黙って選ばない。**タクシーの行き先を告げる客に運転手が経路を選ばせるように**、選択肢を2-3個、それぞれ「メリット／デメリット／リスク／おおよその時間」を非エンジニアの言葉で並べ、推奨に印を付けて選んでもらう。特に: 安全だが遅い道 vs 速いが波及リスクのある道、根治 vs 応急、今やる vs 次ラウンド送り、の分岐は必ず見せる。1つしか道がないときは「一本道である理由」を1文添える。

## 5. 2回ルール（同じ症状が2回目に再発したらアプローチを転換する）

同じ症状が実機報告で**2回目**に出たら、同じアプローチの深掘りを続けることを禁止する。立ち止まり、(1)前提そのものを疑う調査（「そもそもなぜそれが起きうるのか」をゼロから）(2)逆転の発想を含むアプローチ転換の選択肢、をタクシー原則で提示してから進む。実例: 余白問題は同型修正をR18〜R24まで繰り返し、R27でアプローチ転換（境界外の実コンテンツという真相発見）に至るまで7ラウンドを要した。3回目は存在してはならない。

## 6. 実装前プリフライト報告（Human in the loop・2層ルール）

調査・裁定が終わったら、**実装波を投入する前に**、ユーザー（非エンジニア）へ平易な日本語で報告する:
「①何が起きているか（原因）②何をどう直すか ③直すとどうなるか ④壊れるリスクがある場所 ⑤実機での確認方法」の5点、各1-2文。

- **層1（GO必須・待つ）**: 新機能／設計・仕様の変更／凍結解除・分割裁定／検索や取込のロジック変更／複数の解釈がありうる修正。→ 報告してGOを待ってから実装。
- **層2（報告して進む）**: 原因が一意に確定した明白なバグ修正／文言・色・1行修正。→ 報告は出すが実装は止めない（ユーザーは事後でも差し戻せる）。
- 判断に迷ったら層1に倒す。「やって」と言われた大枠の中でも、上流の分岐点は必ず見せる。

## 7. Definition of Done（これを全て満たすまで「完了」と言わない）

0. **`python3 tools/doc_gate.py` → ERROR 0**（2026-09-12 R47 新設）。**文書は検問の
   対象である。** それまでの検問はコードだけを守り、文書を1文字も検査していな
   かった。初回実行で ERROR 100件 ―― 配布zipに同梱する利用者向け文書が「マイ本棚」
   タブ・「ホーム」タブ・🩺診断という**存在しない画面**を説明し、CLAUDE.md の容量
   台帳が実測とズレ、config の既定値の嘘が11箇所あった。
   **これが「実機テストのループが終わらない」の上流だった**: テスターが読む手順書が
   古いと、報告に「本当のバグ」と「文書が古いだけ」が混ざる。司令塔はそれを毎回
   人力で切り分け、後者を「修正」として実装ラウンドへ投入してしまう。**バグが多い
   のではなく、バグでないものがバグとして入ってくる。**
1. `python3 tools/vba_lint.py --path src` → ERROR 0
2. `python3 tools/run_lo_tests.py --mode compile` → OK → `--mode pure` → 全PASS（**run_lo_tests 同士は必ず直列。並列実行は絶対禁止**）
3. `python3 build/build_mybookshelf.py --dev` と `--prod` の自己検証PASS（配布時は `--prod --zip`）
4. 触った全モジュールの容量実測（上限30,000字/WARN28,000字）を報告
5. 新テストのネガティブ確認（期待値を壊して落ちる→復元）。恒真アサート禁止
6. 残課題・記録のみ事項を明示（「たぶん大丈夫」で閉じない）

### 待ち時間の使い方（効率の型）

**テストの3〜4分は待つ時間ではない。** 禁止されているのは `run_lo_tests.py` **同士**の並列であって、テストを裏で走らせながら調査・文書作成・仕様読みを進めるのは推奨。

```bash
nohup bash -c 'python3 tools/run_lo_tests.py --mode compile > /tmp/lo_c.log 2>&1 \
  && python3 tools/run_lo_tests.py --mode pure > /tmp/lo_p.log 2>&1; \
  echo DONE_$? > /tmp/lo_done.txt' >/dev/null 2>&1 &
```

**ビルドはテストの完了後に回す**（同じ `dist/` を触るため）。

### 7-2. 裁定書の無いラウンドはクローズできない（2026-09-12 R47）

R46 は裁定書を書かずにクローズし、「唯一の真実」とされる HANDOFF も R45 のまま
だった。次のセッションが R45 から再開して R46 の変更を壊しうる状態で、R47 の上流
レビューで指摘されるまで誰も気付かなかった。**`docs/dev/spec_*_R{N}_*.md` を書き、
HANDOFF の冒頭節を更新するまでがラウンドのクローズ。**

### 7-3. 裁定書の「文言を更新」には逐語か条件の列挙を必ず書く（2026-09-12 R47）

R13-5 の 5c は条件を4つ持っていたのに、5d の「モード説明文言を新しい役割定義に
合わせて更新」がそれを1行へ潰した。実装者は 5c を自力で読み戻す必要があり、
**落とすのが既定の動作**になる。実際に条件が落ち、実機で「直前までに使った資料
しか見ないの? 本棚全部見るんじゃなかったっけ?」という疑義を生んだ。
**裁定書自身が誤伝の発生点になりうる。**

### 7-4. 「作ったのに繋いでいない」は機械で止める（2026-09-12 R47）

このリポジトリで一番高くついている失敗の型。R46 では**実装した本人が到達不能な
分岐を書き、コミットに「直した」と書いていた**（`ok = True` が立つのは `nHits > 0`
の枝の内側だけなのに、`If ok Then` の内側へ `Else` を置いた）。人の注意力では
止まらないので `vba_lint` に孤児Public検査を入れた。
**新しい Public を作ったら、呼び出しを繋ぐか `@unused:理由` を書くこと。**

## 8. レビューは2周

敵対的レビュー（実装を壊しにいく・反証過程を書かせる）→ Fix波 → **Fix検証パス（修正が新たな縫い目を作っていないかだけを見る2周目）**。R25/R27ではこの2周目が出荷前に実害を計3件止めた。R33の2周目は**14件全部が「Fix波が作った新しい縫い目」**だった。省略しない。

## 9. シグネチャ変更

関数の引数を変えたら、同じ変更の中で全呼び出し箇所をgrepして更新する。lintに引数数照合（9,700箇所超）と予約名検査があるので必ず通す。R21Hの取り残し（modVaultGallery）が実機コンパイルエラーになった実例あり。

## 10. LibreOffice検査の死角（「LO合格=Excel合格」ではない）

実Excelでのみ発症した実例。この型の変更は特に慎重に:
- MS-VBAL予約名（`scale`等）を識別子に使う → 実Excelのみコンパイル不能（lint化済み）
- 引数数の不一致 → LOは照合しない（lint化済み）
- 複数値セルへの `Range.Merge` → Excel標準警告モーダル→フリーズ（描画エントリはmodUiLock.AlertsOff/Onで保護。新規Mergeは必ず保護内に）
- `AscW` はU+8000以降を負値で返す（`If c < 0 Then c = c + 65536` が確立作法）
- `&H8000`以上の16進リテラルは `&` サフィックス無しでInteger負値
- `Workbooks.Open` は破損・保護ファイルでモーダルを出す（DisplayAlerts退避+ダミーPassword作法）
- Shapeに `Hyperlinks.Add`（ScreenTip目的でも）を付けると**クリックはHyperlinkに食われOnActionが死ぬ**（R31実機確定。LOはScreenTip描画もクリック競合も検証しない）

逆向きの罠（LOだけが落ちる／Excelは平気）:
- 予約語 `Enum` と衝突する識別子（`eNum` 等）を使うと、LOは**コンパイルエラーではなく120秒タイムアウト＋結果ファイル未生成でハングする**（R33波1実測）。原因不明のテストハングを見たらまず識別子を疑う。`vba_lint` はERRORで検出するので、LOの前に必ずlintを通せば防げる。
- **大小違いの同名変数（`takeN` と `taken`）は同一識別子の二重宣言**（VBA・LO とも大小無視）。LO は**タイムアウトでハング**、実Excel はコンパイルエラー。`vba_lint` は未検出（R37 実測・R38 で検査追加）。
- **`Optional ByVal x As Object = Nothing` は VBA では不正（既定値は定数のみ）**。LO はタイムアウトでハング。既定値を書かずに `Optional ByVal x As Object` とする（R37 Fix2 実測）。
- **`InStrRev` の4引数形（`Start=-1, vbTextCompare`）は LO で不一致を返す**（R37 実測）。2引数形にする。
- **配列を丸ごと代入する関数（`vec = cut` 型）へ固定長配列を渡すと、LO は通り実Excel は Err 13**（R39・外部受入テストで実証。`modTestsPure8.TestTruncateAndRenorm` が固定長を渡していた）。テストの入力は動的配列で作る。
- **Ghostscript txtwrite の出力に改ページ文字（0x0C）は無い**（R39・同梱 10.03.1 の生出力で実証。合成文字列の分割テストでは見つからない）。ページはファイルを分けて出させる（`-sOutputFile=…_%04d.txt`）。
- **LO の `Erase` は動的配列を解放しない**（2026-09-19 R48 Fix・司令塔が soffice headless で実測）。
  `ReDim v(0 To 1): v(0)=9: Erase v` の直後に `LBound/UBound` が **0/1 を返し**、値が 0 になるだけ。
  実Excel の VBA は `Erase` で動的配列を解放し以後の `UBound` は Err 9 になるので、**意味が逆**。
  さらに LO では**一度も割り当てていない配列でも `LBound` が 0 を返す**（VBA は Err 9）。
  → **`Erase` を「配列を未割り当てに戻す」目的で使ってはいけない。**
  `HasVector`/`ArrLenD` 相当の判定を `Erase` で作ると、LO では常に True のままになり、
  **全0のベクトルが「成功」として保存される**。成否は明示の Boolean フラグで持つこと。
  R48 の初版が実際にこれを踏み、敵対的レビュー（R48-REV-03）で発見された。
- **`ReDim Preserve` を一度も通していない配列への `Join` が空文字を返す**（R33波3実測）。`If 条件 Then ReDim Preserve a(...)` のように条件付きで縮めると、条件が偽の行だけ丸ごと消える。無条件に一度は `ReDim Preserve` を通す形にする。

## 11. VBA落とし穴チェックリスト（実装・レビュー共通）

`And`/`Or`は短絡しない（境界チェックと参照を分離）／`Const`に関数不可／1行1023字まで／非BMP絵文字は識別子・Const不可／`On Error`文とExitはErrをリセットする（**ログの前にErr.Number/Descriptionを退避**）／塗り無しShapeは内部クリック透過（`Fill.Visible=-1`+`Transparency=1`）／**`ScrollArea`はホイールを止めない（R32実機実証。止まるのはセル選択とスクロールバーだけで、`A1:L15`にしても60行付近まで転がる。R29の「止める」は誤帰属で、法則が2回反転している＝3回目を作らないこと）。ホイールの停止線は「使用済み末尾＋約1画面」（最終使用行が画面上端に来るまで転がる。R30実測・R32再現）。したがって塗る深さを増やす対策は原理的に効かない**（塗った行は使用済みになり停止線も同じだけ下がる）／行Hiddenは保存でファイルが数百倍に膨張（R19裁定で不採用）／`OnTime`はブック名修飾。

## 12. 凍結・容量

- 凍結（絶対不触）: modAsk / modBoot / modShelf / modRetrieve / modPrompts
- **容量は必ず python の `len` で実測する**（`wc -m` は当環境のロケールではバイト数を返し、日本語で約3倍になる）。上限30,000字／WARN28,000字。**残り300字未満は分割裁定必須**＝実体を余裕モジュールへ置き、当該モジュールからは1行呼び出しに留める。
<!-- CAP:BEGIN 自動生成。手で書き換えない。`python3 tools/doc_gate.py --write-cap` で書き直し、`--only cap` が実測と照合する（R49 新設） -->
- **実質凍結（残100字未満）・7本** — **1行も入らない。** 実体を受け皿へ置き、ここからは1行呼び出しに留めること: `modClarify`32 / `modUINexusDraw`34 / `modUIShelf`34 / `modUIMain`58 / `modHubStat`68 / `modChunker`76 / `modAskRetrieve`94
- **逼迫（残300字未満）・9本** — **分割裁定必須。** 次に触るなら追加の分割が先: `modViewport`118 / `modBoard`124 / `modHub`170 / `modShare`186 / `modBackdrop`197 / `modTestsPure2`202 / `modApp`254 / `modUI`258 / `modExtractor`268
- **準逼迫（残1,000字未満）・20本** — 次に触るときは要注意: `modAsk`300（凍結） / `modTestsPure24`347 / `modChrome`350 / `modTestsPure25`365 / `modShelfBatch`367 / `modTestsPure18`372 / `modChannel`388 / `modKnowledge`413 / `modShelfStore`431 / `modVaultGallery`459 / `modUtilText`584 / `modIntegrity`666 / `modViewport2`688 / `modP2P`691 / `modTestsPure34`705 / `modTestsPure35`721 / `modTestsPure11`738 / `modMode`833 / `modSparse`953 / `modTestsPure33`997
- **28,000〜29,000字帯（残1,000〜2,000）・15本** — まだ入るが、まとまった追加は分割を考える: `modTestsPure`1,087 / `modTestsPure6`1,212 / `modShareRule`1,237 / `modGateway`1,287 / `optOcrCore`1,331 / `modTestsPure5`1,378 / `modPrompts`1,606（凍結） / `modHelp`1,687 / `optVision`1,693 / `modShelfSync`1,702 / `modTestsPure12`1,713 / `optOcrPage`1,762 / `modAskMulti`1,789 / `modInsightIo`1,804 / `modTestsPure16`1,856
- **WARN直下（残2,000〜2,100）・6本** — **lint は警告を出さないので台帳が唯一の記録。** コメント1行でWARN帯へ落ちる: `modShelf`2,002（凍結） / `modUtil`2,008 / `modBoot`2,009（凍結） / `modSkin`2,039 / `optGsTxt`2,055 / `modRagParse`2,090
- **余裕（残2,100〜10,000）・47本**: `modInsight`2,136 / `modRetrieve`2,259（凍結） / `modPeek`2,335 / `modExtractorWord`2,446 / `modAskFocus`2,476 / `modTestsPure13`2,533 / `modKnowledgeBar`2,562 / `modTestsPure9`2,578 / `modShared`2,603 / `modOutlineBuild`2,944 / `modExtractorPdf`2,981 / `modTestsPure28`2,986 / `modPack`3,092 / `modCorrect`3,101 / `modTestsPure29`3,302 / `modTestsPure3`3,316 / `modTestsPure4`3,414 / `modTestsPure8`3,723 / `modLog`3,774 / `modTestsPure48`3,891 / `modTestsPure17`4,118 / `modTestsPure23`4,244 / `modTestsPure30`4,367 / `modDashStat`4,422 / `modTestsPure7`4,428 / `modTestsPure14`4,890 / `modStats`4,987 / `modXDocStore`5,219 / `modMentor`5,591 / `modDash`5,861 / `modMigrateFrom`6,236 / `modTestsPure15`6,586 / `modGenPipe`6,761 / `modLive`6,786 / `modTestsPure22`6,809 / `modTestsPure40`6,914 / `modAskGlobal`7,757 / `modDiag`8,053 / `modAskOnePass`8,239 / `modXDocBuild`8,358 / `modMigrate`8,585 / `optOcrCache`8,799 / `modAppState`9,159 / `modGuard`9,342 / `modTestsPure10`9,413 / `modPublish`9,536 / `modTestsPure19`9,922
- **受け皿（残10,000以上）・77本** — 新しい実体の置き場所はここから選ぶ: `modExtractorExcel`10,286 / `modPublishUI`10,293 / `modAppAct`10,388 / `optOcrEta`10,608 / `modTestsPure36`10,653 / `modShelfScan`11,462 / `modBackfill`11,470 / `modInsightGate`11,614 / `modTelemetry`11,725 / `modAskThorough`11,770 / `modEmbed`12,264 / `modBitwiseOpt`12,271 / `modPackExport`13,004 / `modTextView`13,015 / `modConvBridge`13,125 / `modFollowup`13,126 / `modStarter`13,389 / `modEnrich`13,580 / `modTestsPure43`13,784 / `modTestRunner`14,587 / `modTestsPure47`14,799 / `modGround`15,181 / `modTestsPure38`15,370 / `modSynonymStore`15,389 / `modVault`15,416 / `modVecCache`15,447 / `modXDoc`15,457 / `modTestsExcel`15,789 / `modShelfVision`16,094 / `modTestsPure45`16,244 / `modTestsPure42`16,431 / `modInsightCard`16,739 / `modLiveStyle`17,180 / `modProgressBar`17,279 / `optGsProc`17,580 / `modTestsPure31`18,022 / `modP2PIo`18,219 / `modChunkMetaStore`18,292 / `modTour`18,424 / `modWorkExcel`18,427 / `modAnalytics`18,516 / `modTestsPure46`19,013 / `optMarkdown`19,060 / `modRibbonFail`19,145 / `modOutlineStore`19,199 / `modChunkPage`19,463 / `optDiffDoc`19,519 / `modUiLock`19,764 / `modCluster`20,178 / `modChunkMeta`20,382 / `modSeed`20,446 / `modClip`21,023 / `modToast`21,471 / `modTestsPure20`21,568 / `modConfig`21,980 / `modPii`22,193 / `modGatewayDirect`22,314 / `modUIMainShape`22,799 / `modTestsPure32`23,246 / `modTestsPure37`23,268 / `modTestsPure44`23,473 / `modChatLog`23,490 / `modSetupWizard`23,765 / `modTestsPure39`24,241 / `modExtractorAcrobat`24,291 / `modFeatures`24,307 / `modTestsPure41`24,742 / `modTestsPure21`24,811 / `modTestsPure26`24,817 / `modState`25,778 / `optTts`25,907 / `modEmj`26,246 / `modNoiseReport`26,294 / `modInstallCheck`26,854 / `modAppDef`26,867 / `modHubBadge`27,110 / `modTypes`28,831
<!-- CAP:END -->
- `tools/run_lo_tests.py` は絶対に並列実行しない。**モード1は FAIL 0 だけでなく SKIP の上限と PASS の下限も照合する**（`EXPECTED_SKIP_MAX` / `EXPECTED_PASS_MIN`）。テストを消して静かにすることも、`[SKIP]` を貼って集計から消すこともできない。

## 13. リポジトリの作法（事故った箇所だけ）

- **ブランチは `claude/internal-notebook-lm-chatbot-B6BE7` 固定。** PRはユーザーが明示的に頼んだときだけ作る。
- **`dist/` 配下はコミットしない。** `.gitignore` は `dist/*` + `!dist/Ghostscript/`。ビルド成果物はAPIキー（`OBF1:`形式。難読化の鍵はソースに同梱＝復元可能）と社内限資料を含む。**個別ファイル指定では二度穴が空いた。** 配布は `SendUserFile` で直接渡す。
- **`.claude/` は追跡する（無視しない）。** スキルとフックは開発の作法そのもので、無視すると新しいスキルが次のクローンで消える。2026-09-02 まで実際に `.claude/` を丸ごと無視していた（既存4ファイルは追跡済みだったので生き残っていただけ）。ルールを変えたら `git check-ignore -v .claude/skills/新スキル/SKILL.md` で無出力を確認する。
- **`git rebase` / `git reset --hard` の前に必ず `git diff` で失うものを見る。** rebase が `.gitignore` の育ててきたルールを丸ごと巻き戻す寸前だった。追記だけの外科手術で済むならそちらを選ぶ。
- **`git add` は個別に。** `git add -A` は禁止（秘密入りファイルを巻き込む）。
- 履歴の書き換え（`git filter-repo` 等）はユーザーの明示的な選択なしに回さない。
- キーの値は**絶対に出力しない**。形・場所だけを報告する。

## 14. 上流原則

川の上流が汚ければ下流も汚い。仕様・設計・要件定義は緻密に。「Aを直したらBが壊れる」を防ぐため、変更点の**繋ぎ目**（呼び出し側・再描画経路・状態のクリア・テーマ再適用・クリア後の残留）を仕様に明記する。数値検算（座標・コントラスト比・容量）を伴わない裁定はしない。

**そして下流（出荷ビルド）が上流（設計書）どおりとは限らない。** 外へ出す言葉の典拠は必ず下流から取る（§3-3）。
