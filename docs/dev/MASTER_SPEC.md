# マイ本棚AI — マスター仕様書(実装契約書) v1.0

> **この文書が唯一の正**。実装エージェントは本書の契約(シグネチャ・シート列・エラーコード)から逸脱してはならない。
> 逸脱が必要だと判断した場合は、実装せずに最終レポートで逸脱提案として報告すること。
> 製品構想の背景は `/home/user/notebook/docs/v3/00_マイ本棚_構想と設計書.md` を参照(ただし本書と矛盾する場合は本書が勝つ)。

---

## 1. ゴールと完成の定義(DoD)

**プロダクト**: 「マイ本棚AI」— 承認なしで自分の資料をぶちこんで、すぐAIに質問できる社内版NotebookLM。
Excel単体(.xlsm)で動作し、社内AIリボン「リボンちゃん」の `ChatGPT()` / `GetEmbeddings()` に相乗りする。
**V2チャットボットとは別プロダクト**(コードは流用するがファイル参照はしない。`mybookshelf/` 配下で完全自己完結)。

DoD:
1. `mybookshelf/build/build_mybookshelf.py` で `mybookshelf/dist/MyBookshelf.xlsm` が本環境でビルドできる
2. `mybookshelf/tools/vba_lint.py` が全モジュールで違反0
3. LibreOffice headless で全モジュールがコンパイル成功+純ロジックテスト全緑
4. コア機能(取込・同期・2速回答・パック・ダッシュボード)は確認済み関数のみで完結
5. opt機能はビルドリスト1行削除+configフラグOFFで撤去でき、コアは無傷
6. 非エンジニア向け使い方/運用保守/撤去手順ドキュメント完備
7. 指定ブランチにコミット済み・mainと未統合・`git subtree split` で切り出し可能

## 2. スコープ

**入る**: ファイル取込(ダイアログ+本棚フォルダ差分同期)、抽出(txt/md/csv/pdf/docx/doc/xlsx)、チャンク分割、
埋め込み(再開可能バッチ)、任意のバッチ富化、2速QA(⚡すぐ聞く/🔍しっかり調べる)+出典表示、
ナレッジパック書き出し/取込(PIIスキャン+重複排除)、個人ダッシュボード+バッジ、環境診断、
opt: 画像読み取り(Vision・スクショ取込含む)/Markdown表示・Wordで開く/約款差分。
(読み上げ(TTS)は公式回答で「音声合成は非公開」と確定したためスコープ外。optTts.basはソース保管のみで
ビルド対象外=modules.jsonから撤去。確定台帳 RIBBON_API_CONFIRMED.md 裁定D4/D10)

**入らない**: V2公式ナレッジベースとの結合(公式ナレッジは将来「公式パック」として取り込む設計で代替)、
Graph API・外部HTTP(リボン以外の外部依存ゼロ)、リアルタイムファイル監視(ポーリング差分で代替)、UFフォーム(シートUIのみ)。

## 3. アーキテクチャと依存ルール

```
[UI層]      modUIMain / modUIShelf / modUIDashboard / modBoot
[機能層]    modShelf / modShelfSync / modAsk / modPack / modStats / modDiag / modEnrich
[部品層]    modExtractor(+Word/Excel/Acrobat) / modChunker / modEmbed / modRetrieve / modPrompts / modPii
[基盤層]    modAppDef / modConfig / modLog / modUtil / modGateway / modTypes / modFeatures
[opt層]     optTts / optVision / optMarkdown / optDiffDoc   ←— コアからの参照は modFeatures 経由の遅延バインドのみ
```

**依存ルール(lintで機械検査する)**:
- R1: 下層は上層を呼ばない(基盤→機能はNG)。例外(UI通知コールバック): 機能層→`modUIMain.SetStage`(進捗実況)、
  `modUIMain.RenderSourcesPreview`(出典先出し)、`modUIMain.RenderAnswer`(回答描画)、`modUIShelf.RenderShelf`(カード再描画)のみ許可。
  (Wave2 PM裁定: §7.2/§7.3が命じる呼び出しと整合させるため例外を4件に拡張。これ以外のUI参照は引き続き禁止)
- R2: `opt*` モジュール名の直接参照(`optTts.` 等のトークン)はコアのどのモジュールにも書かない。
  呼び出しは必ず `modFeatures.InvokeFeature` / `Application.Run` の文字列経由。
- R3: リボン関数(`ChatGPT`/`GetEmbeddings`/その他)の `Application.Run` は modGateway 内のみ。opt層も自前でRunせず `modGateway.TryRibbonRun` を使う。
- R4: **純ロジックモジュール**(modUtil / modChunker / modPii / modTypes と各Pureテスト)は
  Excelオブジェクト(`Worksheets`/`Range`/`Application`/`ThisWorkbook`/`MsgBox`)に触れてはならない(LibreOffice実行互換のため)。
- R5: エラーを握りつぶさない。`On Error Resume Next` は1行スコープ(直後に `On Error GoTo 0`)のみ。失敗は modLog に必ず残す。

## 4. データモデル(シート定義)

シート名は全て英小文字。タブ表示ユーザー向けシートのみ日本語。行1=ヘッダ固定。

| シート | 可視性 | 役割 |
|---|---|---|
| `使い方` | 可視(先頭) | マクロ有効化手順+クイックスタート。マクロ無効でも読める救済ページ |
| `ホーム` | 可視 | 質問UI |
| `マイ本棚` | 可視 | 本棚UI |
| `ダッシュボード` | 可視 | 個人統計+バッジ |
| `config` | hidden | 設定(§5) |
| `my_knowledge` | veryHidden | チャンク本体 |
| `my_vectors` | veryHidden | ベクトル |
| `chunk_meta` | veryHidden | チャンクの構造メタ(R17 Phase1)。ビルドが headers-only で生成する(実行時の `Worksheets.Add` は壊れたブックの自己修復専用)。列 `chunk_id, section_path, refs_out`。**無くても全機能が従来どおり動く**フェイルセーフ前提のシート(下記) |
| `doc_outline` | veryHidden | 章単位要約(R17 Phase2)。ビルドが headers-only で生成する(実行時の `Worksheets.Add` は壊れたブックの自己修復専用)。列 `source, section_key, summary, keywords, chunk_n`。**無くても全機能が従来どおり動く**フェイルセーフ前提のシート(下記) |
| `synonyms` | veryHidden | 用語の表記ゆれ辞書(R17 Phase3)。ビルドが headers-only で生成する(実行時の `Worksheets.Add` は壊れたブックの自己修復専用)。列 `term, canonical`。**無くても全機能が従来どおり動く**フェイルセーフ前提のシート(下記) |
| `ocr_cache` | veryHidden | 画像PDF OCRの頁チェックポイント(R15-7d)。ビルドが headers-only で生成する(R15-FixB FB-2。実行時の `Worksheets.Add` は壊れたブックの自己修復専用)。列 `key, text, saved_at`。key=`Fnv1a64Hex(元フルパス)\|FileLen\|IsoDateTime(更新日時)\|p<頁>`、text は先頭に番兵1字 `t` を置いて書き読み出しで剥ぐ(数式誤解釈の防止と空頁の判別)。opt層(optOcrCache)だけが読み書きし、資料が本棚に `done` として並んだ時点で modShelf がその資料の行を削除、孤児行は起動時GCで2日超を削除する |
| `my_manifest` | hidden | 同期台帳 |
| `my_stats` | hidden | 統計カウンタ+バッジ取得日 |
| `usage_log` | hidden | 利用ログ(1行=1質問/1操作) |
| `err_log` | hidden | エラーログ |
| `ui_state` | veryHidden | UI内部状態(モード選択等) |
| `vba_src` | veryHidden | 自己インストーラ用ソース格納(ビルド時生成) |

**my_knowledge** 列: `chunk_id, source, origin, page, summary, keywords, full_text, added_at, embedded, norm_text`
- norm_text(10列目・2026-08-01 R12-4追加): 照合用の正規化済みテキスト
  (`modSparse.MatchDocText(summary, keywords, source, full_text)` の結果)。
  取込時に1回だけ作り、検索側は読むだけにする。空欄は「未計算」を意味し、
  検索時にその行だけ計算して書き戻す(遅延バックフィル)。見出しは
  `modShelfStore.EnsureKnowledgeSheet` が毎回・冪等に付ける(既存ブックの
  移行処理は不要。manifest の fail_count と同じ作法)。
- chunk_id 形式: `bs::<fnv64hex(full_text正規化後)>::p<page>::c<連番>`。fnvハッシュ部が重複排除キー。
- origin: `self` | `pack:<作成者名>`
- embedded: 0/1(1=my_vectorsに行がある)。再開可能バッチの走査キー。
- full_textは1セル32,000字を超えない(チャンカー保証)。

**my_vectors** 列: `chunk_id, vector_csv`(L2正規化済みDoubleのカンマ結合。次元はconfig `embed_dim`=1536)

**chunk_meta** 列(3列・2026-08-05 R17 Phase1): `chunk_id, section_path, refs_out`
- 目的: 規程・約款の「章>条」という構造ラベルと、本文中の明示参照(「第8条による」
  「別表2のとおり」)を取込時に保存し、検索が【点】ではなく【面】を組めるようにする
  (R17設計書 `docs/dev/design_20260805_R17_構造グラフ設計.md` §3 Phase1)。
- `section_path`: そのチャンクの見出し階層を `>` で連ねたもの(例 `第3章 総則>第12条(免責)`)。
  資料名は含めない(資料の同一性は my_knowledge.source が持つ)。取り出し元は
  modChunker が各チャンク先頭へ置く breadcrumb 行で、**ApplyCrumb を通す前の生チャンク**
  から採る(config `embed_prefix_breadcrumb` が FALSE だと保存本文からこの行が消えるため)。
  全角数字は `modSparse.NormalizeForSearch` で半角へ寄せる(質問側と同じ式=表記ゆれで外れない)。
- `refs_out`: 本文中の明示参照を `|` 区切りにしたもの(例 `第6条|第8条|別表2`)。拾うのは
  「第N条」「第N項」「第N章」「別表N」「様式N」で、1チャンクあたり最大24件。
  自分自身の見出し番号も入る(落とす判断は検索側 `modChunkMeta.RefLabelsFor` が持つ)。
- 抽出は `modChunkMeta`(純ロジック・InStr走査のみ)。**VBScript.RegExp / ScriptControl は
  使わない**(政策ブロックのリスクとLO実行テスト不能=回帰を機械で固定できないため)。
  書込み・掃除は `modChunkMetaStore`、検索への合流は `modAskFocus`。
- **フェイルセーフ(この節の要点)**: chunk_meta が無い/0行(=まだ取り込み直していない
  既存本棚)のとき、`modChunkMeta.GraphActive` が False を返し、参照展開も条番号の直接
  ヒット保証も**1行も足さずに戻る**。回答は R16 までと完全に同じになる。既存資料の移行
  処理は書かない(再取込で生成される)。config `graph_refs=off` でも同じく無操作。
- 掃除: 再取込(`modShelf.IngestFile` の手順7.5)と資料削除(`DeleteSource`)で、消える
  my_knowledge 行の chunk_id を引いて同じ行を落とす(`RemoveMetaForSource`)。
  この呼び出しは必ず `modShelfStore.RemoveKnowledgeAndVectorsForSource` の**前**に置く
  (後だと消えた行の chunk_id がどこにも残っていない)。書込み・掃除の失敗は取込を止めず
  `usage_log("chunk_meta_fail")` を1行残す(検索精度の上積みであってデータ保全ではない)。

**doc_outline** 列(5列・2026-08-05 R17 Phase2): `source, section_key, summary, keywords, chunk_n`
- 目的: 点検索(dense+sparseの上位k件)では原理的に答えられない**俯瞰質問**
  (「〜を全部教えて」「全体像は?」「どんな種類がある?」)へ答えるための材料。
  取込時に章ごと1回だけ要約を作っておき、質問時は**章の要約だけ**を読んで
  「どの章を読むか」を選ぶ(R17設計書 §3 Phase2。GraphRAG のコミュニティ要約を
  「マニュアルが既に持つ章」で代替する、という設計書の骨子そのもの)。
- `section_key`: 章キー = `section_path` の**第1要素をそのまま**
  (`modOutlineBuild.ChapterKeyOf`)。ここで追加の正規化はしない
  (section_path は取込時に `modSparse.NormalizeForSearch` を通っており、
  別の式を掛けると保存側と照合側で章キーが割れる)。**章見出しの無い資料では
  第1要素が条になる**(=条単位の要約)。無い章立てを推測するより外れ方が小さい、
  という保守的動作の明示的な採用。
- `summary` 200〜300字 / `keywords` は `|` 区切り(最大8語) / `chunk_n` はその章の
  チャンク数(人が後から見るための記録で、回答経路は読まない)。
- 生成: `modOutlineBuild.BuildOutlineFor`(取込の出口 `modShelf.IngestFile` の
  Finish から1行)。章ごとに `CallLLM(step="chapter_summary")` を1回=254頁の規程で
  20〜30回・取込+3〜8分。**R15の枠組みにそのまま乗る**: 章の境界で
  `modShelfBatch.CancelRequested` を見て中断なら**そこまでの章を保存して正常終了**、
  `StageBanner`「章の要約中… k/N章(資料名)」+`BlendPerItemMs` のETA、
  `SaveCheckpoint(1, 120)`。1章の本文は `max_context_chars` で打ち切り(`BudgetTake`)。
  失敗章は `(要約失敗)` の行として保存して続行する(行ごと落とすと「その章だけ
  要約が無い」ことが誰にも見えない)。
- **フェイルセーフ**: doc_outline が無い/0行(=まだ取り込み直していない既存本棚、
  `graph_outline=off` のまま使ってきた本棚)なら `modAskGlobal.OutlineActive` が
  False を返し、俯瞰は**一切動かず**回答は R16 までと完全に同じになる。章が1つも
  選ばれない・章のチャンクが引けない・回答生成が失敗、も同じく従来フローへ落ちる。
  既存資料の移行処理は書かない(再取込で生成される。chunk_meta と同じ判断)。
- 掃除: 再取込(`IngestFile` 手順7.5)・資料削除(`DeleteSource`)・章要約の作り直し
  (`BuildOutlineFor` の書込み直前)で `RemoveOutlineForSource`。doc_outline は
  source 列を自分で持つので、chunk_meta と違い
  `RemoveKnowledgeAndVectorsForSource` との前後関係の制約は無い。
  書込み・掃除の失敗は取込を止めず `usage_log("outline_fail")` を1行残す。
- `veryHidden` はビルドの焼き込みと `EnsureOutlineSheet` の自己設定の二重で守る。
  **`modBoot.HideInternalSheets` へは足していない**(残8字で1行も入らないため。
  憲章§4-6。次に modBoot を触る波が分割と同時に足すこと)。

**synonyms** 列(2列・2026-08-05 R17 Phase3): `term, canonical`
- 目的: 用語の表記ゆれ(「回収」⇔「リコール」等)を吸収する辞書。同じ意味の
  質問でも言葉が違うだけでヒットしないRAGの取りこぼしに対して、質問文へ
  同義語を追記してから検索する(docs/45 項目34)。取込末尾に1回だけ生成し、
  質問のたびには生成しない(LLM呼び出しは資料1本の取込につき最大1回)。
- 生成: `modSynonymStore.BuildSynonymsFor`(`modOutlineBuild.BuildOutlineFor` の
  末尾から1行)。当該資料の chunk_meta(section_path)・doc_outline(keywords)・
  my_knowledge(keywords列)から用語候補を集め(重複排除・最大200語)、
  `CallLLM(step="name_dedup")` で表記ゆれグループを1回取得して追記する。
  出力契約は `<syn>表記>正規形|表記>正規形</syn>`(`modRagParse.ParseSynResp`)。
- **前提となる配線(2026-08-05 R17H FB-4 / A-M6・記録+説明)**: 名寄せは
  【章要約の配線に相乗りしている】。`BuildSynonymsFor` の呼び出し口は
  `BuildOutlineFor` の末尾1箇所だけで、そこへ到達するには chunk_meta が
  非0行で、かつ章キー(`ChapterKeyOf`)が1つ以上取れる必要がある。したがって
  **`graph_outline=off` の本棚と、章・条の見出しが取れない資料では
  `graph_synonyms=on` でも synonyms は1行も増えない**。config の説明文
  (`build_config_rows`)にも同じ事実を書いてある。配線の分離
  (`modShelf.IngestFile` の Finish から直接呼ぶ)は modShelf 凍結解除後の次期。
- **既存termは上書き**: マージ規則の唯一の持ち主は純関数
  `modRagParse.MergeSynPairs(oldCsv, newCsv)`(2026-08-05 R17H FA-1 / A-H1)。
  新CSVと term が重なる旧行を落とし、残った旧行を元の順序のまま先に、その
  あとへ新行を並べる(順序安定・大小無視・同じtermは先勝ち・`">"`無しや
  どちらかが空の要素は捨てる)。`modSynonymStore.MergeAndSave` は
  `ReadMapCsv` → `MergeSynPairs` → `RemoveAll` → `WriteSynonymRows` と
  流すだけの薄い層で、規則そのものは持たない(1つの表に追記と上書きの2つの
  書き方を混在させない)。R17H 以前はここに規則があり、`ParseSynResp` が返す
  **0始まり**の配列を 1〜n で読んでいたため、実データでは毎回「添字が範囲外」で
  `BuildSynonymsFor` のハンドラに握り潰され、synonyms は永久に0行だった。
  書けた行数は `usage_log("synonyms_built", detail="source=… groups=N new=M")`
  に残す(N=シートへ書いた行数、M=AIが返したグループ数)。
- 読み出し: `modAskRetrieve.RunMultiRetrieve` の入口で、config `graph_synonyms`
  (既定on)がonかつ synonyms が非空のとき、質問文中の語に一致した同義語を
  最大3語・半角空白区切りで質問文へ追記してから既存の検索フローへ渡す
  (`modRetrieve`/`modSparse` のスコアリング本体は無改修)。一致判定は
  `modSparse.NormalizeForSearch` を両辺に通してから `InStr`(全角/半角の
  表記ゆれも吸収)。純関数 `modRagParse.ExpandQueryBySyn(q, mapCsv, maxAdd)`
  が展開そのものを行う(mapCsv は `ReadMapCsv` の戻り値そのものを渡す設計で、
  modSparse/modRagParse をシートI/O非依存のまま=PURE_LOGIC_MODULESに保つ)。
  `ReadMapCsv` は**1セッション1回だけ**呼びモジュール変数へ控える
  (取込・同期での更新は次にブックを開いたときから反映。docs/10に明記)。
- **フェイルセーフ**: synonyms が無い/0行(=まだ取り込み直していない既存本棚、
  `graph_synonyms=off` のまま使ってきた本棚)なら質問文は一切書き換えられず、
  検索は R17 Phase2 までと完全に同じになる。既存資料の移行処理は書かない
  (再取込で生成される。chunk_meta/doc_outline と同じ判断)。
- 掃除: 資料単位の掃除はしない(synonymsは資料ではなく用語の辞書であり、
  複数資料の候補が1つのtermへ合流し得るため)。書込み・掃除の失敗は取込を
  止めず `usage_log("synonyms_fail")` を1行残す。
- `veryHidden` はビルドの焼き込みと `EnsureSynonymSheet` の自己設定の二重で守る。
  **`modBoot.HideInternalSheets` へは足していない**(doc_outlineと同じ理由。
  次に modBoot を触る波が分割と同時に足すこと)。

**my_manifest** 列(10列): `file_path, file_name, modified_at, size, chunk_count, status, error_note, ingested_at, origin, fail_count`
- 10列目 `fail_count` は 2026-08-01(R12-3-3)で追加した連続失敗回数。`MAX_FAIL_STREAK`(3)回で status を `failed_permanent` へ倒し自動同期のスコープから外す。復帰は「資料を追加」での明示選択(`ResetFailCountForPath`)かファイル更新のみ。行の詰め直しは必ず全10列を運ぶ(9列で詰めると fail_count だけが別の行に残る)。
- status: `done` | `pending` | `partial`(埋め込み未了あり) | `failed` | `image_pdf` | `missing`(同期でファイル消失検知→削除待ち)
- ダイアログ取込のfile_pathは実パス。パック由来はmanifestに載せない(my_knowledge.originで管理)。
- **chunk_count の契約(2026-08-05 R18-2a)**: `modShelfStore.UpsertManifestRow` に
  chunk_count として **負値(-1)を渡すと既存行の値を保持**する(新規行なら0)。
  実データ(my_knowledge の行)を1行も消していない失敗経路は必ず -1 を渡す。
  0で上書きするとカードだけが「0件」に化け、資料が消えたように見えるため
  (実機第5報⑧の確定原因。`modShelf.IngestFile` の失敗3経路が該当)。
- **表示との整合(R18-2b)**: マイ本棚/ギャラリーのカードは chunk_count を表示する。
  `modShelf.SourceList` が毎回 my_knowledge の source別実行数と突合し、食い違えば
  **実行数を正**として manifest を自動修復し `usage_log("manifest_fix")` を1行残す
  (`modIntegrity.ReconcileChunkCount`)。カードは常に実データを映す。

**my_stats** 列: `key, value, updated_at`(key例: `ask_quick_total`, `ask_deep_total`, `ask_thorough_total`, `selfsolve_total`, `hint_total`, `fail_total`, `ingest_files_total`, `pack_export_total`, `pack_import_total`, `streak_days`, `last_used_date`, `badge:<id>`=取得日)

**usage_log** 列: `timestamp, event, mode, detail, latency_ms, hit_count`
(event: `ask` / `feedback_green` / `feedback_yellow` / `feedback_red` / `ingest` / `sync` / `pack_export` / `pack_import` / `diag`)

**err_log** 列: `timestamp, code, context, detail, version`

## 5. config キー台帳(config シート A=key, B=value, C=説明(日本語))

**注: configの正典は `build/build_mybookshelf.py` の `build_config_rows()` 関数(2026-08-03時点で約120キー)です。本表は概要であり、既定値・全キーはビルドスクリプトを参照してください。キー数は要件ごとに増えるため、この文に固定値は書かない。**

| key | 既定値 | 意味 |
|---|---|---|
| mock_llm | TRUE(開発ビルド)/FALSE(本番) | リボン無しモック動作 |
| recommended_model | gpt-5.5 | 精査モードモデル |
| quick_model | gpt-5.5 | 即答モードモデル(将来 gpt-5.4-nano 等に差し替え可) |
| quick_effort / quick_verbosity | low / low | 即答モードの reasoning_effort / verbosity |
| deep_draft_effort / deep_draft_verbosity | medium / high | 精査ドラフト |
| deep_verify_effort / deep_verify_verbosity | high / medium | 精査検証 |
| reasoning_tuning | TRUE | FALSEでeffort/verbosity引数を空送信 |
| llm_wait_sec | 1200 | ChatGPT() Wait |
| topk_quick / topk_deep | 6 / 12 | LLMに渡す上位件数 |
| max_context_chars | 40000 | プロンプトに載せる本文合計上限 |
| answer_language | 日本語 | 回答言語(プロンプトに指定を挿入) |
| embed_dim | 768 | ベクトル次元(パック互換性検査に使用。Plan B: 1536取得→先頭768切詰め+再正規化) |
| embed_sleep_ms | 0 | 埋め込み呼び出し間スロットリング(ミリ秒。0=待たない。レート制限時のみ50〜150へ) |
| shelf_max_chunks | 20500 | 本棚チャンク上限(超過時は取込拒否+整理案内) |
| binary_rag / binary_rag_auto | FALSE / TRUE | 粗選別(バイナリ量子化)の明示指定と自動有効化。binary_rag=FALSEでも binary_rag_auto=TRUE かつ件数>=binary_rag_min なら自動で有効(R12-4)。完全に止めるには binary_rag_auto=FALSE |
| shelf_folder | (空) | 本棚フォルダパス |
| sync_interval_min | 0 | OnTime自動同期間隔(0=off) |
| sync_on_open | TRUE | 起動時に差分同期 |
| enrich_mode | light | off/light/full: バッチ富化(§7.7)の強さ。2026-08-05 R17 Phase3で既定off→lightへ常時ON化(取込直後は少量・残りは同期のたびに少しずつ追いつく後追い方式。1回の処理は`EnrichPending`既定30チャンクぶんだけで長時間ブロックを作らない)。light/fullは現状処理内容の差は無い |
| max_pages_per_file | 300 | 抽出ページ上限(超過は打ち切り+partial) |
| ribbon_addin_name | リボンちゃん | AIリボンのアドイン検出名(RibbonAvailable用。裁定D2) |
| limit_check | TRUE | 起動時LimitCheck(期限・利用同意)。FALSEで無効化(裁定D3) |
| followup_max_pairs | 3 | 『続けて質問』で引き継ぐ履歴の最大ペア数。0以下で無効(裁定D11) |
| word_export_effort / word_export_verbosity | medium / medium | 『Wordで開く』の文書整形パラメータ(裁定D12) |
| feature_tts | FALSE | 読み上げ: 非公開確定のため提供不可・FALSE固定(裁定D4) |
| feature_vision / feature_markdown | TRUE | opt機能フラグ(公式仕様確定によりTRUE昇格。裁定D14) |
| feature_diffdoc | TRUE | 約款差分(確認済み関数のみ使用) |
| ghostscript_path | (空) | 画像PDFのOCRに使うgswin32c.exeのフルパス。空ならブックの隣の`Ghostscript\`を探す(R6) |
| vision_pdf_max_pages | 300 | 画像PDFを読み取る最大ページ数(1ページ=AI1回。20ページずつ画像化→OCR→画像削除を繰り返す。超過は打ち切りpartial。R14-4bで20→100、R15-7aで100→300。ハード上限 optOcrCore.PAGES_MAX も 200→300) |
| ocr_confirm_min_minutes | 5 | 画像PDFのOCRがこの分数以上かかる見込みのとき、取込前に確認ダイアログを出す(R15-7b。R18-1fで15→5=24頁8〜10分級でも確認と作業用Excel導線が出るように)。「はい」の直後に2段目「先に作業用Excelを開いてから開始しますか?」を出す(R18-1f)。見積もり=これから読むページ数×1ページあたりの実測(modState `ocr_avg_page_ms`。無ければ25秒/ページ)。前回の続きから復元できるページは数えない。0以下=確認しない。無人経路(silent同期)では出さない |
| vision_pdf_dpi | 150 | 画像PDFのページ画像化の解像度(公式帳票OCR版と同値。300は約2倍重い) |
| vision_pdf_timeout_sec | 120 | GS処理の**無進捗許容秒数(アイドル上限)**。進捗が観測できる場合はページが進む限り待ち続け、進まなくなってからこの秒数で失敗にする(R13-1cで意味変更) |
| gs_abs_timeout_sec | 1200 | GS待ちの**絶対上限**(秒)。進んでいても必ずここで打ち切り、1資料でExcelが何十分も戻らない事態を防ぐ(R13-1c)。**画像PDFのOCR経路(optOcrPage)のGS待ちはこの絶対上限だけを使い、20ページずつのバッチ全体で1資料あたりの累計として消費する**(各バッチには残り時間だけを渡し、最低10秒。使い切ったら中断としてpartialにする。R14-F6)。R15-7c: OCR経路に限り、総ページ数が判明していれば予算は `max(gs_abs_timeout_sec, 総ページ数×8秒)` へ自動で伸びる(254頁を13バッチ描くのに1200秒では後半が必ず時間切れになるため。他経路のこの設定の使い方は変えない) |
| deep_scope_subqueries | 6 | 「続けて質問」を『しっかり調べる』で行うとき、会話で引用済みの資料の中だけを掘るために作るサブクエリ数(0以下は6扱い。R13-5c) |
| decompose_mode | auto | 「入念に調べる」で複合質問を論点ごとに分けて調べるか(R16-3A)。auto=`decompose_min_chars`以上の質問だけ段0の判定を呼ぶ/always=長さを見ずに毎回呼ぶ/off=分解しない(従来どおり1本の質問として調べる)。知らない値はauto扱い |
| decompose_max_parts | 3 | 分ける論点数の上限(2〜5へ丸める)。1論点ごとにAI呼び出しが1回増える(質問全体で 論点数+6 回。3論点で9回) |
| decompose_min_chars | 25 | `decompose_mode=auto` のとき、この文字数以上の質問だけ段0の判定を行う。短い質問は割る論点が無く、判定の1回ぶんだけ遅くなるため |
| clarify_mode | auto | 読み方が定まらない質問に番号の選択肢で聞き返すか(R16-3B)。auto=選択肢が2件以上作れたときだけ聞き返す(保留はTTL30分)/off=聞き返さずそのまま回答を作る |
| deep_neighbor | 2 | 「入念に調べる」の精読半径。根拠チャンクの前後何個ぶんを一緒に読むか(0=off。R16-3C)。R16H FA-4で適用先を入念のみとし、深掘り(deep)には効かせない(戻り件数が「N件ヒット」バッジと直結するため) |
| graph_refs | on | 条文の参照関係を回答の材料に足すか(R17 Phase1)。on=根拠チャンクの `refs_out` を1ホップ展開して**同じ資料の中**から参照先(第8条・別表2 等)を精読束へ足し、質問が名指しした条番号のチャンクが1件も無ければ chunk_meta から引いて先頭へ入れる(最大2件)/off=検索ヒットだけで答える(R16までと同じ)。LLM呼び出しは1回も増えない。**chunk_meta が無い本棚では on でも従来動作** |
| graph_outline | on | 章単位要約(R17 Phase2)を取込時に作るか。on=章の数だけAIを呼んで doc_outline を作り(254頁の規程で+3〜8分・中断ボタンで途中まで保存)、入念モードの段0が `verdict=global` と判定した質問で章をまたいで答える(質問あたり+2回)/off=作らない・俯瞰質問も従来の検索で答える。**doc_outline が0行の本棚では on でも従来動作** |
| graph_synonyms | on | 用語の表記ゆれ辞書(R17 Phase3)を作り、質問に使うか。on=章要約のあとAIへ用語一覧を1回だけ渡し synonyms を更新(資料1本の取込につき+1回)、質問は一致した語の同義語を最大3語まで質問文に足してから検索する/off=辞書を作らず質問文もそのまま検索する(既に作った辞書は残るが読まれない)。**synonyms が0行の本棚では on でも従来動作**。synonyms は1セッション1回だけ読む(更新は次にブックを開いたときから反映) |
| freeze_keep_banner | TRUE | 長時間ブロック中のDWM「応答なし」白画面化を`user32.DisableProcessWindowsGhosting`で抑止する(R16-2b)。抑止中はウィンドウの移動・最小化・×閉じが効かない(公式の既知の制約)。FALSEで従来どおり白画面化。判定はプロセス中1回だけキャッシュされるため変更はExcel再起動で反映(R16H FB-1) |
| minutes_per_selfsolve | 15 | Hub「自分の節約時間/みんなの節約」の換算係数(自己解決1件=何分か)。modStats/modBoard/modDashStatの3重複定数をここへ統合(R13-7d) |
| pack_author | (空:初回起動で入力) | パック作成者名 |
| debug_mode | FALSE | ゲートウェイのプロンプト/応答ログ |

## 6. エラーコード台帳(modLog)

形式 `E<2桁カテゴリ><2桁連番>`。**ユーザー向けメッセージは必ず「何が起きたか+どうすればよいか」の2文構成**。
コードは小さく末尾表示(保守者がerr_logと突合するため)。

| コード | 意味 | ユーザー向け対処文の要点 |
|---|---|---|
| E0101 | configシート欠損/必須キー欠落 | 配布元にこのファイルの再入手を依頼 |
| E0102 | Scripting.Dictionary(Windows Script Runtime)が使えない端末(R11-D・起動時プローブ) | 管理者へ連絡(端末ポリシーの解除が必要) |
| E0201 | リボン未検出(ChatGPT呼び出し失敗) | AIリボン入りExcelで開き直す。診断ボタンで確認 |
| E0202 | LLM応答がエラー/空 | 時間を置いて再実行。続くなら診断結果を管理者へ |
| E0203 | 埋め込み取得失敗 | 同上 |
| E0204 | 利用上限(LimitCheck)らしき拒否 | 今日はここまで。明日「同期」を押せば続きから再開 |
| E0301 | 対応外の拡張子 | 対応形式一覧を提示 |
| E0302 | ファイルが開けない(ロック/権限) | ファイルを閉じてから再実行 |
| E0303 | 画像PDF(文字が取れない) | このPDFは画像のため取込不可(Vision対応は将来) |
| E0304 | 抽出結果が空 | 中身が空/保護されている可能性 |
| E0401 | チャンク0件 | 文字数が少なすぎる資料 |
| E0501 | 本棚上限超過 | 使っていない資料を削除してから追加 |
| E0502 | 同期フォルダが見つからない | フォルダ選び直しを案内 |
| E0503 | 取込中の再実行(再入禁止) | 処理完了を待つ |
| E0504 | 同名別ファイルの衝突(別フォルダの同名資料が既に登録済み) | ファイル名を変えるか元の資料を削除してから追加(Wave4追加。§7.2 modShelf.IngestFile参照) |
| E0601 | 検索0件 | 資料が本棚にあるか確認を案内 |
| E0602 | 回答生成失敗 | 再実行案内 |
| E0701 | パック形式不正/バージョン不一致 | 相手に再書き出しを依頼 |
| E0702 | ベクトル次元不一致 | リボンの版が違う可能性。管理者へ |
| E0703 | PII検知(書き出し中断可) | 内容確認を促す(警告であり停止ではない) |
| E0705 | 共有フォルダI/O失敗 | ネットワーク再接続・権限確認を案内。通常は次回同期で再試行 |
| E0801 | UI再構築失敗 | ブックを開き直す |
| E0805 | 読み取り専用で開かれている(取込結果が保存されない) | 別のExcelで開いていないか確認を案内。機能は止めない(R15-3b) |
| E0806 | チャンネル数上限超過 | 管理者へ連絡して上限引き上げを依頼 |
| E0807 | 発行ロック異常 | 共有フォルダの接続・書込権限・時計を確認 |
| E0808 | 発行処理異常 | 再実行案内。繰り返す場合は共有フォルダ接続と権限確認 |
| E0901 | 診断で異常検知 | 診断レポートの指示に従う |
| E0904 | 作業用Excelの起動失敗(WScript.Shell 不可等) | タスクバーのExcelをAlt押しながらクリックで代替可(R16-2a。取込中はStatusBar表示のみ) |
| E0905 | 社内ポータルを開けない(FollowHyperlink+WScript.Shell の両方が失敗) | URLをクリップボードへコピーして案内(R18-5b。Hub最下部のフッターからのみ発生) |

## 7. モジュール契約(公開API)

> 記載のない Public は作らない(Privateは自由)。引数・戻り値・失敗時挙動は厳守。
> 全モジュール: `Option Explicit` 必須 / `Attribute VB_Name` = ファイル名 / 1モジュール30,000字以内。

### 7.1 基盤層

**modAppDef.bas** — 定数のみ(ロジック禁止)
```vba
Public Const APP_NAME As String = "マイ本棚AI"
Public Const APP_VERSION As String = "0.1.0"   ' ビルド時にbuildスクリプトが検証表示
Public Const PACK_FORMAT_VERSION As Long = 1
Public Const SH_HOWTO = "使い方": Public Const SH_HOME = "ホーム"
Public Const SH_SHELF = "マイ本棚": Public Const SH_DASH = "ダッシュボード"
Public Const SH_CONFIG = "config": Public Const SH_KNOWLEDGE = "my_knowledge"
Public Const SH_VECTORS = "my_vectors": Public Const SH_MANIFEST = "my_manifest"
Public Const SH_STATS = "my_stats": Public Const SH_USAGE = "usage_log"
Public Const SH_ERRLOG = "err_log": Public Const SH_UISTATE = "ui_state"
' (実際は行を分けて Public Const … As String = … と正しいVBA構文で書くこと)
```

**modTypes.bas** — 型のみ
```vba
Public Type ExtractedPage
    page As Long
    Text As String
End Type
Public Type ShelfChunk
    chunk_id As String: source As String: origin As String
    page As Long: summary As String: keywords As String
    full_text As String
End Type
Public Type Hit
    chunk_id As String: score As Double: source As String
    page As Long: preview As String        ' 先頭120字(出典先出し表示用)
    origin As String
    full_text As String                    ' チャンク本文全体(回答生成の根拠。Wave3 PM裁定1)
End Type
```

**modConfig.bas** — V2 `src/chatbot_v2/modConfig.bas` を流用(SHEET_NAME="config")。
公開API: `EnsureLoaded` / `GetString` / `GetLong` / `GetDouble` / `GetBool` / `SetValue`(V2と同一シグネチャ)。

**modLog.bas**
```vba
Public Sub LogError(ByVal code As String, ByVal context As String, ByVal detail As String)
    ' err_logへ追記(失敗しても例外は出さない)。debug時はDebug.Printも。
Public Sub LogUsage(ByVal event_name As String, ByVal mode As String, ByVal detail As String, _
                    Optional ByVal latency_ms As Long = 0, Optional ByVal hit_count As Long = 0)
Public Function FriendlyMessage(ByVal code As String) As String  ' §6の対処文(2文構成)を返す
Public Sub ShowError(ByVal code As String, ByVal context As String, ByVal detail As String)
    ' LogError + MsgBox(FriendlyMessage & vbLf & "(コード: " & code & ")")
```

**modUtil.bas** — 純ロジック(R4遵守: Excelオブジェクト禁止)
```vba
Public Function Fnv1a64Hex(ByVal s As String) As String        ' 16桁hex。Currency/2xLong実装可。決定的であること
Public Function NormalizeForHash(ByVal s As String) As String  ' 空白圧縮+改行統一+Trim
Public Function VectorToCsv(vec() As Double) As String
Public Function CsvToVector(ByVal s As String, ByRef vec() As Double) As Boolean ' 不正時False
Public Function DotProduct(a() As Double, b() As Double) As Double ' 次元不一致は0を返しはしない→-2を返す規約は複雑なので: 次元不一致時は0返し+呼び出し側が次元検査
Public Function L2Normalize(ByRef vec() As Double) As Boolean  ' ゼロベクトルはFalse
Public Function HasVector(vec() As Double) As Boolean          ' 未初期化/空配列判定(On Error使用可)
Public Function SplitKeepNonEmpty(ByVal s As String, ByVal sep As String) As String()
Public Function HumanBytes(ByVal n As Double) As String        ' "1.2 MB"
Public Function HumanSeconds(ByVal sec As Double) As String    ' "約1分20秒"
Public Function SafeLeft(ByVal s As String, ByVal n As Long) As String
Public Function NowStamp() As String                            ' "yyyy-mm-dd hh:nn:ss"
Public Function FileNameOf(ByVal path As String) As String      ' 区切りは \ と / 両対応
Public Function ExtOf(ByVal path As String) As String           ' 小文字拡張子(ドットなし)
Public Function IsSameTimestamp(ByVal a As Date, ByVal b As Date) As Boolean ' 2秒丸め(FAT/OneDrive誤差)
Public Function EtaText(ByVal remainCount As Long, ByVal msPerItem As Double) As String
    ' R7 B-1: 残り件数×1件あたりミリ秒 → "残り約3分" / 60秒未満は "まもなく完了" /
    ' 見積り不能(msPerItem<=0)・残り0件は "" を返す(呼び出し側は件数だけ出す)
Public Function ProgressText(ByVal doneN As Long, ByVal totalN As Long, ByVal etaPart As String) As String
    ' R7 B-1: "12/88件・残り約3分"(etaPartが空なら "12/88件")
```
注: `NowStamp`はNow()を使う(Excel/LO両対応)。`FileNameOf`/`ExtOf`は文字列処理のみなのでR4適合。

**modGateway.bas** — リボン唯一の窓口。V2 `modRibbonGateway.bas` を基に拡張。
```vba
Public Function CallLLM(ByVal prompt As String, ByVal step_name As String, _
                        ByVal effort As String, ByVal verbosity As String, _
                        Optional ByVal model_override As String = "", _
                        Optional ByRef latency_ms As Long = 0, _
                        Optional ByVal prevU As String = "", _
                        Optional ByVal prevA As String = "") As String
    ' prevU/prevA: 会話継続用の履歴(確定台帳§1 #1 第7・8引数。新しい順;;;区切り。裁定D11)
    ' 成功: 応答文字列 / 失敗: "#ERR:E0202:..." で始まる文字列(例外は出さない)
    ' mock_llm=TRUE時: step_nameに応じた整形済みダミー(quick_draft/deep_draft/deep_verify/enrich/diff)
Public Function GetEmbedding(ByVal Text As String, Optional ByRef latency_ms As Long = 0) As Double()
    ' 成功: L2正規化済み配列 / 失敗: 空配列(modUtil.HasVector=False)。E0203をLogError
    ' mock時: テキストのFnvハッシュをシードにした決定的擬似ベクトル(L2正規化済み, embed_dim次元)
    '         → リボン無しでも取込→検索→回答の全画面フローが動く
Public Function RibbonAvailable() As Boolean   ' AddInsループでアドイン検出(裁定D2。結果をセッションキャッシュ)
Public Function RunLimitCheck() As Boolean     ' リボンLimitCheck()の唯一の呼び出し口(裁定D3。True=続行不可、エラー時False)
Public Function TryRibbonRun(ByVal funcName As String, ByVal args As Variant) As Variant
    ' Variant配列argsを展開してApplication.Run(要素数0〜6対応のSelect Case)。
    ' 失敗時: 文字列 "#ERR:E0202:<説明>" を返す。opt層はこれだけを使う
Public Function LooksLikeLimitError(ByVal response As String) As Boolean ' 上限/limit/回数の語を検知→E0204系判定
```
ChatGPT位置引数規約は確定済み(V2実証+確定台帳§0「互換後方追加」。裁定D1):
`Application.Run("ChatGPT", prompt, "", 0.4, 0, waitSec, model, prevU, prevA, "マイ本棚AI:" & step_name, effort, verbosity)`。
第9引数toolNは「マイ本棚AI:」ブランドを付けて管理側ログでツールを識別できるようにする。
GetEmbeddings の実引数規約は V2 `src/chatbot_v2/modEmbeddings.bas` の実装をそのまま踏襲すること。

**modFeatures.bas** — opt機能の分離実行
```vba
Public Function FeatureEnabled(ByVal featureId As String) As Boolean
    ' config "feature_<id>"=TRUE かつ ModulePresent
Public Function ModulePresent(ByVal featureId As String) As Boolean
    ' Application.Run "opt<Pascal>.Ping" をOn Errorで試す(結果キャッシュ)。opt側はPublic Function Ping() As Boolean を必ず持つ
Public Function InvokeFeature(ByVal featureId As String, ByVal procName As String, ByVal args As Variant) As Variant
    ' Application.Run("opt<Pascal>." & procName, ...) 遅延バインド。
    ' 無効/不在/失敗 → "#ERR:FEATURE_UNAVAILABLE" を返し、呼び出しUI側が
    ' 「この機能は現在利用できません(管理者が有効化すると使えます)」を表示
```
featureId→モジュール名対応表: tts→optTts, vision→optVision, markdown→optMarkdown, diffdoc→optDiffDoc(modAppDefではなくmodFeatures内Privateに保持)。

**modDiag.bas**
```vba
Public Sub RunDiagnostics()   ' 「🩺診断」ボタン。diag_reportシートを(再)生成して表示:
    ' バージョン/ビルド日、リボン検出、mock_llm状態、各シート存在+行数、config必須キー、
    ' 本棚統計(資料数/チャンク数/未埋め込み数)、opt機能の在否、直近エラー5件、
    ' 各行に ✅/⚠️ と対処文。最後に「この画面をスクリーンショットして管理者に送ってください」
Public Function QuickHealthCheck() As String  ' Boot時の軽量版: 問題なければ "" / あれば警告文
Public Sub WarnIfReadOnly()   ' R15-3b: 読み取り専用なら起動時に1回だけ伝える(E0805)
```

**modIntegrity.bas**(2026-08-05 R18-2b/2d 新設。データ整合性の観測点)
```vba
Public Function IndexOfName(ByRef names() As String, ByVal n As Long, _
                            ByVal target As String, ByVal hint As Long) As Long
    ' source別集計の位置引き(hintは速さだけの助言。正しさは線形探索が保証)
Public Function ReconcileStatText(ByVal statText As String, ByVal actualN As Long) As String
    ' "status|ingested_at|chunk_count|error_note|origin" の3つ目を実行数へ差し替え。
    ' 一致していれば同じ文字列を素通し(=呼び出し側の「何もしない」合図)
Public Function ReconcileChunkCount(ByVal sourceName As String, ByVal statText As String, _
                                    ByVal actualN As Long) As String
    ' 上記+manifest の5列目も直し usage_log("manifest_fix") を1行(modShelf.SourceListから)
Public Sub RecordSaveMark()   ' (my_knowledge行数, FullName) を ui_state へ控える
    ' 【R18H FA-3】呼ぶのは (a) ThisWorkbook.Save の【直前】(SaveCheckpoint)、
    ' (b) 意図した削除の直後(modShelf.DeleteSource / modShelfStore.
    ' RemoveRowsByOrigin(Prefix)。保存はしない)。(a)を保存の「後」に置くと
    ' ディスク上のマークが常に1世代古くなり、(b)が無いと利用者自身の削除が
    ' 次回起動で「資料が消えました」の虚偽警告になる
Public Function DataShrunk(ByVal prevRows As Long, ByVal curRows As Long) As Boolean
    ' prevRows<=0(記録なし)は判定しない=初回起動で根拠なく警告しない
Public Function IsVolatilePath(ByVal folderPath As String, _
                               Optional ByVal tempA As String = "", _
                               Optional ByVal tempB As String = "") As Boolean
    ' 【R18H FA-4】判定対象は ThisWorkbook.Path(フォルダ)。"temp1_" / ".zip\" の
    ' 部分一致か、tempA/tempB(=Environ("TEMP")/Environ("TMP") の実値)配下の
    ' 前方一致。"\temp\" の部分一致は廃止(D:\temp\ や \\share\Temp\ の誤検知)。
    ' 環境変数は呼び出し元(WarnAtStartup)が読む=この関数は純ロジックのまま
Public Function ShrinkWarnMsg(...) As String / VolatileWarnMsg(...) As String  ' 警告文(BMPのみ)
Public Sub WarnAtStartup()    ' 起動時の突合(modBootから1行)。1セッション1回まで
```
判定材料は my_knowledge / my_manifest / ui_state のシートだけで、上位層は呼ばない(R1)。
置き場が基盤層なのは modBoot / modShelf / modShelfStore がいずれも30,000字上限近くで
判定本体を置けないため(modDiag.WarnIfReadOnly と同型の判断。憲章§4-6)。

### 7.2 取込層

**modExtractor.bas / modExtractorWord.bas / modExtractorExcel.bas / modExtractorAcrobat.bas**
— `src/admin/` の同名モジュールをコピーし適応。公開契約:
```vba
' modExtractor
Public Function ExtractFile(ByVal path As String, ByRef pages() As ExtractedPage, _
                            ByRef errCode As String, ByRef errDetail As String) As Boolean
    ' 拡張子ディスパッチ。失敗時False+errCode(E0301/E0302/E0303/E0304)。例外は内部で捕捉
Public Function SupportedExts() As String   ' "txt,md,csv,pdf,docx,doc,xlsx,xls,xlsm"
```
適応要件: V1の`Err.Raise`スタイルはこの関数内で捕捉してBoolean+errCodeに変換。
画像PDF検知: 抽出結果の総文字数がページ数×10未満 → E0303。
`max_pages_per_file`超過は打ち切り、`errCode="PARTIAL_PAGES"`(Boolean=True扱い・呼び出し側がpartial記録)。

**modChunker.bas** — `src/admin/modChunker.bas` をコピー適応。純ロジック(R4)。
```vba
Public Function ChunkPages(pages() As ExtractedPage, ByVal targetChars As Long, _
                           ByVal overlapChars As Long, ByRef chunks() As ShelfChunk) As Long
    ' 戻り値=チャンク数(0可)。chunk_id/source/originは空のまま(modShelfが付番)。
    ' full_textは32000字を絶対に超えない。既定 targetChars=700 / overlap=150
```

**modEmbed.bas** — 再開可能バッチ埋め込み(V2 modVectorize のパターン移植)
```vba
Public Function EmbedPending(Optional ByVal maxCount As Long = -1) As Long
    ' my_knowledge の embedded=0 行を順に GetEmbedding → my_vectors 追記 → embedded=1
    ' 進捗を modUIMain.SetStage("📥 ベクトル化中 12/250 …") で実況。DoEvents。
    ' ESCで中断(Application.EnableCancelKey=xlErrorHandler, Err=18捕捉)→ そこまで保存し件数返す
    ' 3連続失敗 or LooksLikeLimitError → 中断し E0204 案内(「明日🔄同期で続きから」)
    ' 戻り値: 今回埋め込んだ件数。呼び出し間に embed_sleep_ms スリープ
Public Function PendingCount() As Long
```

**modShelf.bas** — 本棚中核
```vba
Public Sub AddFilesViaDialog()      ' 複数選択FileDialog(フィルタ=SupportedExts)→ 各IngestFile
Public Function IngestFile(ByVal path As String, ByVal origin As String) As String
    ' 戻り値=manifest status("done"/"partial"/"failed"/"image_pdf")
    ' 手順: 再入guard(E0503) → 同名衝突検査(E0504。Wave4追加: source識別がファイル名の
    '       みであるため、manifestに「同じfile_nameだが別のfile_path」の行が既にあれば、
    '       削除せず明示エラーで止める。誠実な失敗の方が黙ったデータ消失より安全という
    '       R5の判断。同一パスの再取込=置換はこの検査を通過する)
    '       → 上限検査(E0501)
    '       → ExtractFile → ChunkPages → chunk_id付番(bs::hash::pN::cN, ハッシュ重複はスキップ)
    '       → my_knowledge追記(embedded=0)
    '       → 【R18-2e】既存同名sourceの旧chunk/vectorはここで削除する(新しい行を
    '          書き終えた【後】。書込みがerr#7等で落ちても旧データを無傷で残すため。
    '          いま書いた行は keepFromRow で巻き添えにしない。同一sourceの新旧行が
    '          同時に存在する窓は取込中だけで、取込中は検索がUIロックで走らない)
    '       → 【R18H FA-5】acceptedCount=0(全チャンクが既存と同一ハッシュ)なら
    '          削除そのものをスキップし旧行を温存。status=failed / chunk_count=-1 /
    '          メモ=modShelfStore.MEMO_ALL_DUP、usage_log("ingest_all_dup")を1行
    '       → 【R18H FA-6】旧行削除が失敗(Err<>0)したら status=partial とし、メモ
    '          先頭へ modShelfStore.MEMO_REPLACE_NG を付ける(E0801ログは従来どおり)
    '       → manifest upsert(status=pending)
    '       → EmbedPending → manifest status確定 → 成功時(done/partial)のみ
    '       my_stats.ingest_files_totalをBump+usage_logに"ingest"イベントを記録
    '       → 【R18-2c】成功時(done/partial)のみ modShelfBatch.SaveCheckpoint(silent含む。
    '          【R18H FA-7】silent は throttleSec:=120 を渡す=無人同期で1件ごとに
    '          ブックを書き戻さない。同期末尾の1回(modShelfSync)は従来どおり)。
    '          中間保存の呼び出し点はこのFinish1箇所に集約する(以前は
    '          modShelfBatch.AddFilesResult のループにしか無く、スクショ取込・
    '          ナレッジ登録・修正・共有登録・パック取込・部門チャンネル取込の
    '          6経路が【セッション中に一度も保存されなかった】=憲章§4-5)
    '       → カード再描画(modUIShelf.RenderShelf)
Public Sub DeleteSource(ByVal sourceName As String)  ' knowledge/vectors/manifest から一括削除+再描画
Public Function SourceList(ByRef names() As String, ByRef stats() As String) As Long ' カード描画用
Public Function TotalChunks() As Long
Public Function IsBusy() As Boolean
    ' R7 B-2: 再入ガード mIngesting の公開。抽出ループのDoEventsで発火した
    ' クリックを画面遷移側(modUiLock.BlockIfIngesting)が受け流すために使う。
    ' 焼き付き対策としてGUARD_EXPIRY_MIN超過のガードはFalseを返す。
```
注(R7 B-2): 再入guard(E0503)は ShowError(モーダル)ではなく
`modLog.LogError` + `modUIMain.SetStage` のみで通知する(連打のたびに
ダイアログが積み上がると、それ自体が「応答なし」に見えるため)。
E0503 のログ記録は従来どおり残す。

**modShelfSync.bas**
```vba
Public Sub PickShelfFolder()      ' フォルダ選択→config shelf_folder 保存→即SyncNow
Public Sub SyncNow(Optional ByVal silent As Boolean = False)
    ' Dir()走査(サブフォルダ不要・第1階層のみ) vs my_manifest:
    '   新規→IngestFile / modified_at or size 変化(IsSameTimestamp)→置換取込 /
    '   消失→DeleteSource / pending・partial→EmbedPending再開
    ' フォルダ不存在→E0502。実行結果サマリをステータス表示+LogUsage("sync")
    ' silent(既定False。Wave4追加): TrueのときはE0502警告・完了サマリをMsgBoxで出さず
    ' SetStage(状態表示行+StatusBar)のみに留める。modBoot.Boot(sync_on_open)と
    ' AutoSyncTick(定期自動同期)はユーザー操作なしで走るバックグラウンド処理のため
    ' silent:=Trueで呼ぶ(対話ダイアログでフォーカスを奪わない。§8 UXレビュー対応:
    ' 配布直後にshelf_folder未設定のままE0502警告が必ず出る問題/定期MsgBoxが
    ' 作業を中断する問題への対応)。手動🔄ボタン・PickShelfFolder直後はsilent=False。
Public Function IsBusy() As Boolean
    ' R7 B-2: 再入ガード mSyncRunning の公開(modShelf.IsBusyと対)。
    ' 同期のファイルループも1ファイルごとにDoEventsを回すため必要。
Public Sub ScheduleAutoSync()     ' sync_interval_min>0ならApplication.OnTimeで次回予約(自己再帰)
Public Sub CancelAutoSync()       ' 予約解除(必ずOn Error握り: 予約なしでも安全)
Public Sub AutoSyncTick()         ' OnTimeコールバック本体(Wave4修正: 契約上Public必須。下記注参照)
Public Function ResolveDecision(ByVal decision As String, ByVal existsInManifest As Boolean, _
                                ByVal currentStatus As String) As String
    ' 純関数(§7.8): DiffDecisionの結果にmanifestの現在statusによる上書きルールを適用する。
    ' decision="keep"かつstatus="failed"/"missing" → "replace"(前回失敗の再試行/
    ' フォルダ復活時の復帰。Wave4修正: missing行が復活後もkeep固定のまま戻らない不具合の修正)
' 注意: 予約解除は modBoot.Auto_Close から必ず呼ぶ(OnTime残存→勝手にExcelが再起動する事故防止)
' 注意(Wave4修正・契約自体の欠陥だったための修正): Application.OnTimeのProcedure引数は
'   Application.Runと同じ遅延バインド(文字列からの実行時解決)であり、Private Subは
'   解決できない。予約自体は成立してしまうが、発火時に「マクロを実行できません」という
'   未処理エラーになり自動同期が機能しなくなる。そのためAutoSyncTickは
'   「OnTimeコールバックとして公開が必須」の契約とし、Public Subとして実装すること。
```

**modChunkMeta.bas**(2026-08-05 R17 Phase1) — 構造メタの抽出(純ロジック・R4準拠)
```vba
Public Function ExtractSectionPath(ByVal rawChunkText As String) As String
    ' チャンク1行目の breadcrumb「【〔資料〕 > 章 > 条】」から "章>条" を作る。
    ' 先頭要素(資料名)は必ず捨てる。breadcrumb 無し/閉じ括弧無し/中身が空は "" 。
Public Function ExtractRefs(ByVal bodyText As String) As String
    ' 本文の「第N条/第N項/第N章/別表N/様式N」を "|" 区切り・重複なし・最大24件で返す。
    ' InStr走査のみ(RegExp/ScriptControl不使用)。正規化は modSparse.NormalizeForSearch。
Public Sub MetaOf(ByVal rawChunkText As String, ByRef outPath As String, ByRef outRefs As String)
    ' 上2本を1回で呼ぶ入口(modShelf.IngestFile の残り字数のため。憲章§4-6)。
Public Function PathHasLabel(ByVal sectionPath As String, ByVal label As String) As Boolean
Public Function RefLabelsFor(ByVal refsOut As String, ByVal ownPath As String) As String
    ' 自分自身の見出し番号を落とした参照ラベル列(自己参照の無害化)。
Public Function GraphActive(ByVal metaCount As Long, ByVal nHits As Long) As Boolean
    ' 【フェイルセーフの単一情報源】chunk_meta 0行 or ヒット0件なら False=無操作。
```

**modChunkMetaStore.bas**(2026-08-05 R17波0→Phase1) — chunk_meta シートI/O
```vba
Public Function EnsureChunkMetaSheet() As Worksheet   ' EnsureKnowledgeSheetと同型(冪等)
Public Sub WriteMetaRows(ids, paths, refs, n)         ' 末尾へ1回のRange書込みで追記
Public Sub WriteMetaFromRows(srcRows, idCol, paths, refs, n)
    ' 取込ループが持つ my_knowledge の行列から chunk_id を取り出して追記する入口。
Public Function ReadAllMeta(outIds, outPaths, outRefs) As Long   ' 全行(0 To n-1)。無ければ0
Public Sub RemoveMetaForSource(ByVal sourceName As String, ByVal keepFromRow As Long)
    ' 消える my_knowledge 行(同名source・keepFromRowより前)の chunk_id を引いて
    ' chunk_meta の同じ行を落とす。必ず RemoveKnowledgeAndVectorsForSource の【前】。
' 失敗は全て握って usage_log("chunk_meta_fail") 1行。取込は止めない。
```

**modOutlineStore.bas**(2026-08-05 R17 Phase2) — doc_outline シートI/O
```vba
Public Function EnsureOutlineSheet() As Worksheet   ' EnsureChunkMetaSheetと同型(冪等)
    ' veryHidden をここで毎回・冪等に自己設定する(modBoot は残8字で足せない)。
Public Sub WriteOutlineRows(srcs, keys, sums, kws, ns, n)   ' 末尾へ1回のRange書込み
Public Function ReadOutline(outSrcs, outKeys, outSums, outKws) As Long
    ' 全行(0 To n-1)。シートが無い/0行なら 0(呼び出し元は必ず空配列を受け取れる)。
Public Sub RemoveOutlineForSource(ByVal sourceName As String)
    ' その資料の章要約を全部落とす(source列を自分で持つので前後関係の制約なし)。
' 失敗は全て握って usage_log("outline_fail") 1行。取込は止めない。
```

**modOutlineBuild.bas**(2026-08-05 R17 Phase2) — 章単位要約の作成
```vba
Public Sub BuildOutlineFor(ByVal sourceName As String)
    ' my_knowledge+chunk_meta から資料のチャンクを文書順に集め、章キーで束ね、
    ' 章ごとに CallLLM(step="chapter_summary") で要約+キーワードを作って保存。
    ' graph_outline=off / chunk_meta 0行 / 章キーが取れない なら完全に無操作。
    ' 中断は章の境界(CancelRequested / Err18)で拾い、そこまでの章を保存して正常終了。
    ' 呼び出しは modShelf.IngestFile の Finish(done/partial)から1行だけ。
Public Function ChapterKeyOf(ByVal sectionPath As String) As String   ' 第1要素(純ロジック)
Public Function BudgetTake(ByVal usedLen As Long, ByVal addLen As Long, ByVal cap As Long) As Long
    ' 1章の本文の打ち切り(純ロジック)。俯瞰側(modAskGlobal)の章ごと予算にも使う。
' プロンプトは modPrompts ではなく本モジュールの Private(modPrompts 残386字の容量裁定)。
' 出力契約: <summary>…</summary><keywords>a|b|c</keywords>(modRagParse.ParseOutlineResp)
```

**modEnrich.bas** — バッチ富化(summary/keywords付与)
```vba
Public Function EnrichPending(Optional ByVal maxCount As Long = 30) As Long
    ' enrich_mode=off なら即0。summary空のチャンクを10件/1回のCallLLM(JSON配列返し)で富化。
    ' JSONパースは軽量自前(期待形: [{"i":1,"summary":"…","keywords":"a,b"},…])。
    ' パース失敗はそのバッチをスキップして続行(富化は無くても検索は動く=非致命)
    ' maxCount既定30(2026-08-05 R17 Phase3): 常時ON化での小口バッチ上限。
    ' 呼び出し元(modShelf.IngestFile/modShelfSync.SyncNow)は無改修=引数無しで呼ぶ。
```

**modSynonymStore.bas**(2026-08-05 R17 Phase3) — 用語の表記ゆれ辞書(synonyms)I/O+名寄せバッチ
```vba
Public Function EnsureSynonymSheet() As Worksheet   ' EnsureOutlineSheetと同型(冪等)
Public Sub WriteSynonymRows(terms, canons, n)   ' 末尾へ1回のRange書込みで追記のみ(上書き判断なし)
Public Function ReadMapCsv() As String
    ' 全行を "term>canonical|term>canonical" の1文字列で返す(0行は空文字)。
    ' modAskRetrieveが1セッション1回だけ呼ぶ唯一の読み出し窓口。
Public Sub RemoveAll()   ' 全行消去(再構築用)
Public Sub BuildSynonymsFor(ByVal sourceName As String)
    ' 唯一の名寄せ入口。config graph_synonyms=off / 用語候補0件 / LLM応答が空
    ' のいずれかで完全に無操作。当該資料の chunk_meta(section_path)・
    ' doc_outline(keywords)・my_knowledge(keywords列)から用語候補を集め
    ' (重複排除・最大200語)、CallLLM(step="name_dedup")で表記ゆれグループを
    ' 取得しsynonymsへ追記する(既存termは上書き=ReadMapCsvの既存分から
    ' 新規termと重なる行を除きRemoveAll後に全件書き直す)。ゲート・失敗握り
    ' (usage_log "synonyms_fail")もこの層に閉じるので呼び出し元
    ' (modOutlineBuild.BuildOutlineFor)は1行。CancelRequestedも確認する。
' 出力契約: <syn>表記>正規形|表記>正規形</syn>(modRagParse.ParseSynResp)
' 質問文への展開そのものは純関数 modRagParse.ExpandQueryBySyn が持つ
' (modSparse/modRagParseはPURE_LOGIC_MODULESのためシートI/Oを持てない)。
```

### 7.3 QA層

**modVecCache**(2026-08-01 R12-4追加): セッション内ベクトルキャッシュと
埋め込み世代カウンタ。`Generation` / `BumpGeneration` / `ResetVecCache` /
`BuildFrom` / `PrepareVectors` / `DotAt` / `SlotOfRow` / `StampOf` /
`IsStale` / `Ready` ほか参照系。my_vectors の vector_csv を1セッション1回だけ
パースして Double 配列で保持する。世代カウンタは「取込(EmbedPendingの書込)・
削除(RemoveVectorsByIds)・再埋め込み(MarkAllForReembed)」の3つの書込点が
必ず進め、キャッシュと modBitwiseOpt の量子化コードの双方を無効化する
(件数・先頭/末尾idの印だけでは再埋め込みを検知できない)。構築失敗(err7)は
捕捉して従来経路へ自動フォールバックし、usage_log に `veccache_fallback` を残す。

**modAskGlobal.bas** — 俯瞰質問=疑似グローバル検索(R17 Phase2)
```vba
Public Function TryGlobal(ByVal q As String, ByRef hits() As Hit, ByRef nHits As Long, _
                          ByRef ok As Boolean, ByRef result As String) As Boolean
    ' 入念モードの段0が verdict=global と判定したときだけ modAskMulti から呼ばれる。
    ' 段1: doc_outline の章要約を1回のプロンプトへ載せ、読むべき章を最大4つ選ばせる
    '      (step="chapter_pick"。<pick>資料名::章キー|…</pick>。行頭の文字列をそのまま
    '       書き写させる=番号だと1つずれた瞬間に別の章を読み始めて誰も気付けない)
    ' 段2: 選ばれた章のチャンクを my_knowledge+chunk_meta から文書順に集める
    '      (資料と章キーの【両方】一致・章ごとに max_context_chars÷章数 の予算で打ち切り)
    ' 段3: 章をまたいだ回答を1回で作り(step="global_answer"・出典タグ必須)、
    '      modAskThorough.AnnotateAgainstHits で出典突合する
    ' 追加のLLM呼び出しは1質問あたり2回。返す hits の score は 0(検索スコアで
    ' 選んだ材料ではないので信頼度バッジを水増ししない=近傍と同じ判断)。
    ' False = 何もしていない(呼び出し元は従来の入念フローを実行する)。
Public Function OutlineActive(ByVal outlineN As Long) As Boolean
    ' 【フェイルセーフの単一情報源】doc_outline 0行なら False=俯瞰は一切動かない。
' 不発の理由は usage_log("global_zero" why=no_outline/pick_err/no_pick/no_chunk/
' answer_err/aborted/err)へ1行。中断(Err18)は段の境界で拾い、選んだ章の
' 【取込時に作った要約】だけをそうと明記して「※中断」つきで返す。
' プロンプトは本モジュールの Private。出典タグの書式は modPrompts.SourceTag を通す。
```

**modAskFocus.bas** — 精読(R16-3C)と構造グラフへの合流(R17 Phase1)
```vba
Public Sub NeighborExpand(hits(), nHits, ByVal radius As Long)
    ' 【物理近傍】根拠チャンクの前後 radius 個を文書順で末尾へ足す(config deep_neighbor)。
Public Sub RefsExpand(hits(), nHits, ByVal maxAdd As Long)
    ' 【参照エッジ】ヒットの refs_out を1ホップ展開し、同じ資料(source)の中で
    ' section_path にそのラベルを含むチャンクを末尾へ足す(maxAdd<=0 は既定8件)。
    ' 資料を跨がない: 「第8条」は資料ごとに別の条文で、跨ぐと無関係な規程が
    ' 出典タグ付きで根拠に混ざる(利用者が気付けない外し方)。score=0。
    ' 呼び出しは modAskThorough.RunThoroughFlow と modAskMulti.TryDecomposed の
    ' NeighborExpand 直後に各1行(deep_neighbor=0 でも効く=別軸の機能)。
Public Sub ArticleEnsure(ByVal query As String, hits(), nHits, ByVal maxIns As Long)
    ' 【条番号の直接ヒット保証】質問が名指しした条番号・別表・様式に一致する
    ' section_path のチャンクが hits に1件も無いときだけ、chunk_meta から引いて
    ' 先頭へ最大 maxIns 件入れる(score は先頭ヒットと同値。既にあれば無操作)。
    ' 呼び出しは modAskRetrieve.RunMultiRetrieve の検索直後(単段フォールバック側も)。
' 両方とも config graph_refs=off / chunk_meta 0行(GraphActive=False)なら無操作。
' ゲートの読みはこの層に閉じる(呼び出し元はどれも1行のまま)。
```

**modAskRetrieve.bas** — 検索して材料を揃える層(§C多段RAG)+語彙ズレの吸収(R17 Phase3)
```vba
Public Function RunMultiRetrieve(q, mdMode, topK, hits(), …) As Long
    ' 入口で config graph_synonyms=on かつ synonyms 非空のとき、質問文中の語に
    ' 一致した同義語を最大3語・半角空白区切りで q へ追記してから既存の多段RAG
    ' (拡張→マルチクエリ→再ランク)へ渡す。modRetrieve/modSparseのスコアリング
    ' 本体は無改修(質問文を書き換えるだけ)。synonymsは modSynonymStore.
    ' ReadMapCsv を1セッション1回だけ読み、モジュール変数(mSynMapCsv/
    ' mSynLoaded)へ控えて使い回す(取込・同期での更新は次セッションから反映)。
    ' 展開そのものは純関数 modRagParse.ExpandQueryBySyn。
```

**modRetrieve.bas**
```vba
Public Function Search(ByVal query As String, ByVal topK As Long, ByRef hits() As Hit) As Long
    ' GetEmbedding(query) → my_vectors全件Dot積(全てL2済みなのでコサイン同値)
    ' + キーワードボーナス: queryの語がsummary/keywords/sourceに含まれれば +0.05/語(上限0.15)
    ' 戻り値=件数(0可)。埋め込み失敗時は-1(呼び出し側がE0203表示)
    ' hits().full_text にはmy_knowledgeのfull_text列(チャンク本文全体)をそのまま格納する
    ' (Wave3 PM裁定1: previewだけでは回答生成の根拠として不十分なため)。
    ' hits().preview は引き続き先頭120字(出典先出し表示用)。
```

**modPrompts.bas** — プロンプト組み立て(純文字列。answer_language挿入)
```vba
Public Function BuildQuickPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long) As String
Public Function BuildDeepDraftPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long, ByVal history As String) As String
Public Function BuildDeepVerifyPrompt(ByVal q As String, ByVal draft As String, hits() As Hit, ByVal nHits As Long) As String
Public Function BuildEnrichPrompt(ByVal batchText As String) As String
```
出典指示: 回答内の引用は `[本棚:ファイル名 p.页]` 形式で、と明示。パック由来は `[パック(作成者):ファイル名]`。
「資料に無いことは『資料には見当たらない』と述べ、推測は推測と明示」を全モードで指示。
本棚抜粋ブロックの本文には `hits().full_text`(チャンク本文全体)を使う(Wave3 PM裁定1。
previewは出典先出し表示専用でプロンプト本文には使わない)。
本文組み込みは `max_context_chars` で打ち切り(切ったら「(一部省略)」を挿入)。

**modAsk.bas**
```vba
Public Sub AskFromUI()   ' ホームの質問セル+モード(ui_state)を読み、Answer実行→RenderAnswer
Public Function Answer(ByVal question As String, ByVal mode As String) As String
    ' mode "quick": SetStage(🔍検索中)→Search(topk_quick)→出典先出し表示→SetStage(✍️回答作成中)
    '               →CallLLM(quick)→回答
    ' mode "deep":  Search(topk_deep)→出典先出し→draft→SetStage(✅検証中)→verify→回答
    ' 検索0件: LLMを呼ばず「本棚に手がかりが見つからない」定型文+資料追加の案内(E0601はログのみ)
    ' ESC対応・所要秒を回答末尾に小さく表示。LogUsage("ask")
Public Sub FeedbackGreen() / FeedbackYellow() / FeedbackRed()  ' modStatsへ+お礼表示
Public Function CanFollowup() As Boolean       ' 続けて質問できる直近回答があるか(裁定D11)
Public Sub AskFollowup(ByVal followupText As String) ' 履歴付き追質問(検索も再実行→出典付き回答)
```
会話履歴(裁定D11): 直近 followup_max_pairs(既定3)往復を module 変数に保持し、CallLLM の
prevU/prevA(新しい順;;;区切り)へ渡す。履歴に積むのは深掘り候補ブロック除去後の本文のみ。
深掘り候補: modPrompts が最終応答系プロンプト(quick/deep_verify)に [[FOLLOWUP: 候補1 | 候補2]] の
出力指示を加え、modAsk がパース・除去して「🔎 深掘り候補(『続けて質問』でそのまま聞けます)」
ブロックを本文末尾に整形追記する(RenderAnswer契約は不変。マーカー無し応答でも壊れない)。

### 7.4 パック層

**modPii.bas** — `src/chatbot_v2/modPii.bas` コピー(純ロジックR4)。
`Public Function ScanText(ByVal s As String) As String`(検知種別のカンマ列挙 or 空)。

**modPack.bas**
```vba
Public Sub ExportPackDialog()
    ' 本棚全体 or アクティブ行の資料1件を選択(vbYesNoCancel)→保存先ダイアログ
    ' → 新規Workbook(.xlsx マクロ無し)に pack_meta/pack_chunks/pack_vectors を書いて保存
    ' → 書き出し前に modPii.ScanText を全チャンクに実行、検知時は件数+例を出して続行確認(E0703)
Public Sub ImportPackDialog()
    ' ファイル選択→ValidatePack→重複(fnvハッシュ)スキップしつつmy_knowledge/my_vectorsへ
    ' origin="pack:"&作成者。取込サマリ表示(「32件取込 / 8件は既にありました」)
Public Function ValidatePack(ByVal wb As Workbook, ByRef reason As String) As Boolean
    ' 3シート存在・pack_format_version一致・embed_dim一致(E0701/E0702)
```
**pack_meta** 列(key,value): pack_format_version, pack_name, author, created_at, description, embed_dim, app_version, chunk_count。
**pack_chunks** 列: chunk_id, source, page, summary, keywords, full_text。**pack_vectors** 列: chunk_id, vector_csv。

### 7.5 統計層

**modStats.bas**
```vba
Public Sub Bump(ByVal key As String, Optional ByVal delta As Long = 1)  ' my_stats upsert
Public Function GetStat(ByVal key As String) As Long
Public Sub TouchToday()            ' streak_days/last_used_date 更新(連続利用日数)
Public Sub EvaluateBadges()        ' §9のバッジ条件を検査、新規獲得は badge:<id> に日付+お祝いMsgBox(1回だけ)
Public Function SavedMinutesEstimate() As Long ' 自己解決1件=15分換算(config化不要・定数コメントで明示)
```

### 7.6 UI層(§8のレイアウト仕様に従う)

**modUIMain.bas**
```vba
Public Sub EnsureLayout()          ' ホームを冪等再構築(既存Shapes全削除→再生成)
Public Sub SetStage(ByVal msg As String)  ' ステータス行+Application.StatusBar 両方
Public Sub RenderAnswer(ByVal answerText As String, hits() As Hit, ByVal nHits As Long, ByVal mode As String, ByVal seconds As Long)
    ' mode=""(Wave4追加の目印): 空質問等で検索・回答生成を一切行わなかったことを示す。
    ' このときは所要秒・出典欄(前回質問の情報に見えてしまう)を付けず案内文だけを表示する。
Public Sub RenderSourcesPreview(hits() As Hit, ByVal nHits As Long)  ' 出典先出し(ドラフト生成前に呼ぶ)
Public Sub OnAskButton() / OnModeQuick() / OnModeDeep() / OnOpenHowto() / OnRunDiag()
Public Sub ShowTip()               ' 待ち時間豆知識(定型10本からRnd選択)
```

**modUIShelf.bas**
```vba
Public Sub EnsureLayout()
Public Sub RenderShelf()           ' manifest+SourceListから資料カード行を再描画(§8.2)
Public Sub OnAddFiles() / OnSyncNow() / OnPickFolder() / OnExportPack() / OnImportPack() / OnDeleteSource()
' OnDeleteSource: アクティブ行のsourceを特定(カード行範囲外なら案内)→「『x』を本棚から削除しますか?」確認
```

**modUIDashboard.bas**
```vba
Public Sub EnsureLayout()
Public Sub RenderDashboard()       ' 統計タイル+バッジ棚+REPT("■")棒グラフ。EvaluateBadges後に呼ばれる
```

**modBoot.bas + ThisWorkbook.cls**
```vba
Public Sub Boot()      ' V2パターン踏襲: config確認→first-run(pack_author入力)→3画面EnsureLayout
                       ' →QuickHealthCheck表示→AIリボン利用期限確認(modGateway.RunLimitCheck。
                       '   True=制限中でもvbInformation案内のみで起動は止めない。裁定D3)
                       ' →sync_on_openならSyncNow→ScheduleAutoSync→内部シート隠蔽
Public Sub Auto_Open() ' Boot呼び(ガード付き)
Public Sub Auto_Close()' CancelAutoSync(必須!)+Application.StatusBar=False
```
ThisWorkbook.cls は Workbook_Open→Boot / Workbook_BeforeClose→Auto_Close の薄い転送のみ
(ビルド版ではインストーラがThisWorkbookを占有するため、実行時はAuto_Open/Auto_Closeが本命。両方書く=V2実証済み二重化)。

**重要(2026-08-05 R18-1g・調査agent0で事実確認)**: `Workbook_BeforeClose` は
**開発構成にしか存在しない**。本番(prod)ビルドは自己インストーラが ThisWorkbook
ストリームを占有し `Workbook_Open` しか持たない(build/build_mybookshelf.py:1002-1010、
build/modules.json に明記)。`modBoot.Auto_Close` には Cancel 引数が無く閉鎖を止められない。
したがって **R11 C1 / R13-4d の「取込中の終了禁止ガード」(modUiLock.ConfirmCloseDuringIngest)は
本番では動いていない**。実機で×が効かないのはこのガードの働きではなく、
`DisableProcessWindowsGhosting`(config freeze_keep_banner 既定on)により
「応答なし」中の代行ウィンドウが作られないため(§本節の既知の制約)。
本番で保護されているのは🚪終了ボタン経路だけ。恒久対策は次期(HANDOFF 次期課題)。

### 7.7 opt層(全モジュール共通契約)

- 必ず `Public Function Ping() As Boolean`(True返すだけ)を持つ
- リボン呼び出しは `modGateway.TryRibbonRun` のみ。**リボン関数のシグネチャは確定台帳
  RIBBON_API_CONFIRMED.md §1 が唯一の根拠**(公開仕様の入手により、当初のSIGNATURE ASSUMPTION
  ブロックは「確定済み」の出典記載に置き換え済み)
- コアモジュールへの参照は基盤層+modUIMain.SetStageのみ可
- 失敗しても例外を外に出さない(文字列 "#ERR:..." 返し)

```vba
' optTts.bas    : ビルド対象外(音声合成は非公開と確定。裁定D4/D10。ソースのみ保管)
' optVision.bas : Public Function ExtractImagePdf(ByVal path As String, ByRef pages() As ExtractedPage) As Boolean
'                 ' 確定経路: Base64FromFile(path)→ChatGPTV(prompt, b64, "", "high", "マイ本棚AI:vision")(裁定D5)
'                 ' 対応形式はpng/jpg/jpegのみ(PDF直渡しは公式仕様上不可能と確定し撤去)
'                 Public Function ExtractImagePdfText(ByVal path As String) As String ' "#ERR:.." or 全文
'                 ' modShelfは E0303(画像PDF) または E0301+画像拡張子 のとき
'                 ' feature_vision有効なら InvokeFeature("vision","ExtractImagePdfText",path) を試す(裁定D13)
'                 Public Function HasClipboardImage() As Boolean   ' IsImageInCB確定関数(裁定D13)
'                 Public Function SaveClipboardImage() As String   ' Base64FromCB(Ptn=1)→Temp jpgパス/"#ERR:..."
' optMarkdown.bas: Public Function RenderMarkdownAt(ByVal sheetName As String, ByVal cellAddr As String, ByVal md As String) As String
'                 ' 確定方式: セルにSafeLeft(md,32000)を書き込んでから CellMarkDown(rng, False)(裁定D6)
'                 Public Function OpenAnswerInWord(ByVal md As String) As String  ' OpenWordMark確定関数(裁定D6)
'                 Public Function ExportAnswerAsDoc(ByVal answerText As String, ByVal instruction As String) As String
'                 ' 対話型文書生成(裁定D12): 指示文ありならCallLLM(step="word_export")で整形→OpenWordMark。
'                 ' 指示文空ならOpenAnswerInWordへ直行(そのまま転記)
' optDiffDoc.bas : Public Sub CompareTwoDocsDialog()  ' 新旧2ファイル→抽出→ChatGPT差分分析→diff_reportシート
'                  (確認済み関数のみ使用だがサブ機能なのでopt隔離)
```
UI側: opt機能のボタンは `FeatureEnabled` がTrueのときだけ生成(EnsureLayout内で分岐)。
ボタンOnActionは modUIMain/modUIShelf 内の `OnOpenWordButton`/`OnIngestScreenshot` 等の
ラッパー(そこから InvokeFeature)。ホーム31:32行は btn_followup「💬 続けて質問」(A31:D32・常時)+
btn_word「📝 Wordで開く」(E31:H32・feature_markdown時のみ)。本棚K1:L2は btn_screenshot
「📸 スクショ取込」(feature_vision時のみ)。読み上げボタンは存在しない(裁定D10)。

### 7.8 テストモジュール

**modTestRunner.bas**(R4準拠・ExcelでもLOでも動く中核)
```vba
Public Sub ResetTests()
Public Sub Check(ByVal name As String, ByVal cond As Boolean, Optional ByVal detail As String = "")
Public Function Failures() As Long
Public Function ReportText() As String     ' "PASS 42 / FAIL 0" + 失敗一覧
Public Sub RunAllPureTests()               ' 各modTests*のRun~を列挙呼び出し
```
**modTestsPure.bas**: modUtil(ハッシュ決定性・CSV往復・L2/Dot・タイムスタンプ丸め)、
modChunker(空/1文/長文/32000字上限/オーバーラップ検証/段落境界)、modPii、
modPrompts(出典形式・打ち切り)、manifest差分ロジック(modShelfSyncの差分判定をPrivateではなく
`Public Function DiffDecision(ByVal existsInManifest As Boolean, ByVal sizeChanged As Boolean, ByVal timeChanged As Boolean) As String`
のような純関数に切り出してテスト可能にする)、pack検証(ValidatePackの列検査部を純関数に切り出し)。
**Excel側のみ**: modTestsExcel.bas(シートI/O・ゲートウェイmock経由のE2E: mock取込→検索→回答)。診断から実行可能に。

## 8. UI仕様(非エンジニア向け・iPhone基準)

共通原則: ボタンは角丸四角Shape+OnAction、高さ2行以上・ラベルは動詞(「質問する」)。
選択状態は色で明示(選択=濃紺白字 / 非選択=薄灰黒字)。専門用語禁止(「ベクトル化」→「AIが読める形に変換」)。
全ての失敗メッセージは§6の2文構成。フォント: 游ゴシック11pt基準・見出し14pt太字。
列幅・行高はEnsureLayoutが設定(ユーザーが壊しても開き直せば直る=冪等)。

### 8.1 ホーム
```
[📚 マイ本棚AI  v0.1.0]                      [❓使い方] [🩺診断]
「自分で入れた資料に、すぐ聞ける」
┌───────────────┐ ┌───────────────┐
│ ⚡ すぐ聞く (10〜20秒) │ │ 🔍 しっかり調べる (1〜2分) │   ← トグル
└───────────────┘ └───────────────┘
質問をここに入力してください(例: 〇〇の手続きに必要な書類は?)
[                    大きな入力セル(結合・折返し)                    ]
              [ 💬  質 問 す る ]
状態: 🔍 検索中… / 📄 3件の資料がヒット / ✍️ 回答作成中…
─ 回答 ─────────────────────────
(回答本文: 結合セル・折返し・上詰め)
📖 この回答のもと: ファイルA p.3 / ファイルB p.12
この回答は役に立ちましたか?  [🟢 解決した!] [🟡 ヒントになった] [🔴 だめだった]
💡 豆知識: (待ち時間に表示)
```
- 本棚が空のとき: 回答エリアに「まず『マイ本棚』タブで資料を1つ追加してみましょう →」を常設表示。

### 8.2 マイ本棚
```
[＋ 資料を追加] [🔄 フォルダと同期] [📦 パックにして渡す] [📥 パックを取り込む] [🗑 選んだ資料を削除]
本棚フォルダ: C:\Users\...\OneDrive\マイ本棚  [📁 フォルダを選ぶ]  自動同期: 30分ごと/オフ
ここにファイルを入れておくと、自動で本棚に追加されます(消せば本棚からも消えます)
──資料カード(1行=1資料)──────────────────
 状態   資料名                    追加日     ページ数→チャンク数   メモ
 ✅     医療保険_引受ガイド.pdf   7/11      98p → 245件
 ⏳     新商品QA.docx            7/11      変換中… (34/120)
 ⚠️     約款_旧版.pdf            7/10      取込失敗: ファイルが開けません(E0302)
 🖼     スキャン申込書.pdf        7/10      画像PDFのため取込不可
```
- 削除はアクティブセルの行の資料を対象(ボタン押下時にファイル名を確認ダイアログで明示)。

### 8.3 ダッシュボード
統計タイル(今月の質問数/🟢自己解決数/節約した時間≒解決×15分/本棚の冊数)+
バッジ棚(獲得=カラー絵文字+日付、未獲得=灰色+条件文)+蔵書数の推移棒(REPT)。

## 9. バッジ定義
`first_ingest`初めての取込 / `shelf10`本棚10冊 / `shelf30`本棚30冊 / `first_pack_out`初パック共有 /
`first_pack_in`初パック取込 / `solve10`自己解決10件 / `solve50`自己解決50件 / `streak7`7日連続利用。

## 10. ビルド仕様(build_mybookshelf.py)

- `build/build_chatbot_v2.py` の実証機構(template_skeleton.xlsm + openpyxlシート生成 + vbaProject.bin の
  ThisWorkbook/dirストリーム外科パッチ + vba_src自己インストーラ)を **mybookshelf/build/ にコピーして独立させる**
  (make_xlsm.pyの圧縮関数含む。元ファイルは変更しない)。template_skeleton.xlsm もコピー。
- インストーラは `modBoot.Boot` を起動(cp932エンコード制約により installer 内文字列はASCIIのみ)。
- モジュールリストはマニフェスト(`build/modules.json`)で管理: `{"name","path","role":"core|opt|test"}`。
  **opt機能の撤去=このJSONから1行削除+configフラグFALSEでビルド**、を保証。
- `--dev`(mock_llm=TRUE, テストモジュール同梱)/`--prod`(mock_llm=FALSE, テスト同梱は維持=診断用)。
- シート生成: §4の全シート+ヘッダ+config初期値(C列に日本語説明)+使い方シート本文。
- ビルド後自己検証(同スクリプト内): 再オープンして全シート存在・vba_srcモジュール数一致・
  各ソースセル≤32,000字・olefileでThisWorkbookストリーム復元確認。検証失敗はexit 1。

## 11. テスト計画(3層)

1. **静的Lint(`tools/vba_lint.py`)** — 全.bas/.clsに対し:
   Option Explicit / VB_Name=ファイル名 / モジュール≤30,000字 /
   Public宣言が§7契約と一致(過不足検出: 契約表はlint内に埋め込む) /
   クロスモジュール参照 `modX.Proc` の実在 / R1〜R4依存ルール / `Application.Run`第1引数のホワイトリスト /
   `Dim a, b As Long` 型落ち警告 / `Integer`警告 / セル書込値に32,767字超リスクのSafeLeft無し警告(full_text系) /
   opt直接参照検出。**Exit code非0=違反あり**。
2. **LibreOffice実行(`tools/run_lo_tests.py`)** — ユーザープロファイルのStandardライブラリに
   R4準拠モジュール+modTestRunner+modTestsPureを.xba化して注入し、
   `soffice --headless "vnd.sun.star.script:...location=application"` でRunAllPureTests実行→
   結果を/tmpファイル経由で回収しexit code化。**さらに全モジュール(Excel依存含む)を構文コンパイルのみ**
   別ライブラリでロードして構文エラー検出(実行はしない)。Attribute行はローダが除去。
   VBA固有でLOが解釈できない構文が出た場合は、lint側で当該構文の代替を規約化する(勝手にテスト対象から外さない)。
3. **実機受入(`docs/40_受入チェックリスト15分.md`)** — 配布先PCで: 診断→mockでE2E→mock解除→
   リボン実呼び出し1回→txt1件取込→質問→パック書き出し/取込→終了時OnTime残存なし確認。

## 12. コーディング規約(VBA落とし穴・全員厳守)

- 文字列連結の長大化は `&` 連鎖でなく配列+Join(パフォーマンス)。ループ内Range直接アクセス禁止→配列一括読み書き。
- `Integer`禁止(Long)。`Dim a, b As Long`禁止(1行1変数か全てに型)。
- 日付比較は `IsSameTimestamp`(OneDrive/FATの2秒丸め)。
- セル書込は `SafeLeft(s, 32000)` を経由(full_text/回答/ログ)。
- 絵文字はセル値・Shape文字列では可、**MsgBox/InputBox文字列では使わない**(豆腐化)。VBA文字列リテラルには
  BMP外絵文字(🟢等)を書いてよい(V2実証済み)が、**ビルドスクリプトがUTF-8→セル格納**の経路のみ(cp932直埋め禁止)。
- `On Error Resume Next` の乱用禁止(R5)。エラーハンドラでは必ず `Err.Clear`/`On Error GoTo 0` を明示。
- `ActiveWorkbook`禁止→`ThisWorkbook`。`Select/Activate`はUI表示目的以外禁止。
- モジュール間の循環参照禁止(R1)。
- Shapes: 名前は `btn_`/`lbl_` プレフィクス+機械名。EnsureLayoutは既存の自前Shapesを全削除してから再生成(冪等)。
- OnTime予約は必ず変数に予約時刻を保持し、Cancelは同時刻指定(Excel仕様)。
- 純ロジックモジュール(R4)では `vbCrLf`ではなく`vbLf`基準(LO互換)。Excel層で表示時にvbLf使用可(セル内改行はvbLf)。

## 13. エッジケースカタログ(実装とレビューの検査表)

取込: 空ファイル/0字抽出/1文字/巨大(300p打切→partial表示)/画像PDF/開けない(ロック)/対応外拡張子/
日本語・スペース・長いパス/同名ファイル再取込(置換)/別フォルダの同名別ファイル(E0504で明示エラー・削除しない。Wave4追加)/
同一内容別名(ハッシュでスキップ)/上限5000超/
取込中にESC/取込中にもう一度ボタン(再入guard)/embedded=0が残った状態でExcel強制終了→次回同期で再開/
取込中に×で閉じる(**本番はBeforeCloseが無いため無防備**。R18-1g・上記§7.6の注記)。
同期: フォルダ未設定で同期ボタン/フォルダ削除・リネーム(E0502+missing)/ファイル差替え(更新日時2秒差)/
OneDriveオフライン(プレースホルダ: Dir()で見えるがOpenで失敗→E0302でfailed記録・次回再試行)/
自動同期中にユーザーが質問実行(guardで同期側をスキップ)。
QA: 本棚空で質問/検索0件/LLMがエラー文字列/LLMが出典形式を無視(そのまま表示・落ちない)/
質問が空/超長文質問(3000字で打切り確認)/連打(実行中フラグ)。
パック: 壊れたxlsx/シート欠落/次元不一致/自分のパックを自分で取込(全スキップ)/PII検知で中断選択/
書き出し先が読み取り専用。
環境: リボン無しPC(mock案内)/マクロ無効(使い方シートが説明)/VBAプロジェクトアクセス不可(インストーラがTrustメッセージ)/
Excel 32/64bit(Declare不使用で回避)/Mac(COM抽出不可→E0302系の丁寧な案内があること。コア動作はWindows前提と明記)。
終了: OnTime残存なし/StatusBar復元。

## 14. 実装ウェーブとファイル分担(エージェント間でファイル重複禁止)

| Wave | 担当 | 書くファイル |
|---|---|---|
| 1-A | foundation | src/core/{modAppDef,modTypes,modConfig,modLog,modUtil,modGateway,modFeatures,modDiag}.bas |
| 1-H | build | build/{build_mybookshelf.py, ovba.py, template_skeleton.xlsm(コピー), modules.json, README.md} |
| 1-I | testharness | tools/{vba_lint.py, run_lo_tests.py, README.md}, src/test/modTestRunner.bas |
| 2-B | extract | src/ingest/{modExtractor,modExtractorWord,modExtractorExcel,modExtractorAcrobat,modChunker}.bas |
| 2-C | shelf | src/ingest/{modEmbed,modShelf,modShelfSync,modEnrich}.bas |
| 2-D | qa | src/qa/{modRetrieve,modPrompts,modAsk}.bas |
| 2-E | pack | src/pack/{modPii,modPack}.bas, src/stats/modStats.bas |
| 2-F | ui | src/ui/{modUIMain,modUIShelf,modUIDashboard,modBoot}.bas, src/ui/ThisWorkbook.cls |
| 2-G | opt | src/opt/{optTts,optVision,optMarkdown,optDiffDoc}.bas |
| 2-X | tests | src/test/{modTestsPure,modTestsExcel}.bas(2-B〜2-Eの成果を読んで書く) |

Wave3: lint+LO全緑までの修正ループ / Wave4: Opus6レンズレビュー / Wave5: ビルド+ドキュメント+コミット。
