# マイ本棚AI — RAG全面再設計 設計書 v1.0(チャンク化・検索・回答生成)

> **位置づけ**: 本書は `MASTER_SPEC.md`(実装契約書)の **拡張提案**である。既存契約(シグネチャ・シート列・エラーコード・依存ルール R1〜R5)は原則不変とし、破壊的変更を避けて機能を積み増す。本書で新設・変更を要する公開API/configキー/CONTRACT(`tools/vba_lint.py`)は §G の各タスクで**明示**する。矛盾時は MASTER_SPEC が正。
> **前提**: 実行本線は会社PCのExcel VBAのみ(Python/PowerShell不可)。社内AIリボン確定API(`RIBBON_API_CONFIRMED.md`)に相乗り。ただしオーナー裁定で「Azure OpenAI直接HTTP(バッチ埋め込み)」を**設定切替のオプション**として両対応で抽象化する。コスト度外視でAPI呼び出し回数は増やしてよい。
> **実装者への約束**: 本書の全変更は **configフラグの既定値で従来動作を維持**し、フラグONで新動作に切り替わる(段階導入・即時ロールバック可能)。既存の純ロジックテスト84本 + lint + LibreOffice(LO)コンパイルゲートを**1タスクごとに全緑のまま**進める。

---

## A. 現状の品質ギャップ分析(NotebookLMとの差の構造的要因)

NotebookLM級の回答品質は「①文脈を保ったチャンク → ②高再現率・高精度の検索 → ③根拠に忠実な生成」の3段が噛み合って初めて出る。現行実装は各段に**構造的な**弱点を抱える。以下、コードの具体箇所で列挙する。

### A-1. チャンク化(`modChunker.bas`)— 文脈欠落が最上流の律速
1. **機械的スライディングウィンドウ**: `ChunkOnePage` は 700字固定 + 150字オーバーラップ。`FindSentenceBoundary` は境界を `nearPos`〜`nearPos + tgt\4`(=175字)前方の**句読点・改行**でしか探さない。見出し・条番号・箇条書き・表・段落の**意味境界を認識しない**ため、「第3条第2項」の途中でチャンクが分断される。
2. **メタデータ皆無**: チャンク本文(`full_text`)は素の抜粋のみ。**どの資料の・どの章・どの条**の抜粋かという文脈が本文にもベクトルにも入らない。埋め込み(`modEmbed` → `GetEmbedding(fullText)`)は文脈ゼロの700字だけをベクトル化するので、「該当しない」「同上」のような**それ単体では意味を持たない断片**が正しく検索できない。
3. **構造の破壊**: `NormalizeWhitespace` が3連続改行を2連続へ圧縮するのみで、見出し階層・箇条書き・表の行構造の情報は失われる。

### A-2. 検索(`modRetrieve.bas`)— 単段・単ベクトル・弱いレキシカル
4. **単一クエリ・単一ベクトル**: `Search` は質問文をそのまま1回 `GetEmbedding` し、全件内積 top-k を取るだけ。**クエリ拡張なし・マルチクエリなし・リランキングなし**。質問の言い換えや下位概念を拾えず**再現率が低い**。
5. **キーワードボーナスが実質死んでいる**: `TokenizeQuery` は半角/全角スペース分割。**日本語の質問はスペースが無く「質問文全体が1語」**になり、`summary`/`keywords`/`source` への部分一致がほぼ効かない。さらに `keywords`/`summary` は `enrich_mode=off`(既定)で**空**なので、レキシカル信号がそもそも存在しない。
6. **精度段が無い**: cosine上位をそのまま採用。近いが無関係なチャンク(数値の近さ)を弾く**再ランク段が無い**ため精度が上がらない。

### A-3. 生成(`modPrompts.bas` / `modAsk.bas`)— グラウンディングが緩い
7. **出典が任意・未検証**: `CitationInstruction` は出典形式を指示するが、**強制も検証もしない**。LLMが無視してもそのまま表示(§13で許容)。
8. **推論と回答が未分離**: `<thinking>`/`<answer>` のような構造が無く、思考の漏れ出しや冗長さを分離できない。「資料からは判断できません」の**明示的なリファューザル契約が弱い**(「見当たらない」の努力目標どまり)。
9. **quickは実質1発**: `RunQuickFlow` は検索→単発生成。deepも draft→verify の2段だが、いずれも**クエリ理解段・再ランク段が無い**。

### A-4. 取込パイプライン(`modEmbed.bas`)— 速度がスケールを縛る
10. **1チャンク=1API**: `GetEmbedding` を1件ずつ。リボンの `GetEmbeddings` はバッチ入力不可のため、数百〜数千チャンクの資料で**取込が長時間化**(実機で「遅すぎる」報告済み)。これが「本棚を大きくしたい」というオーナー要望の実質的な上限になっている。

### 総括(構造的因果)
> NotebookLMとの差は「モデルの賢さ」ではなく**RAGパイプラインの段数と各段の質**にある。本再設計は、**(B)文脈付きチャンク → (C)多段検索(拡張→マルチクエリ→再ランク)→ (D)構造化グラウンディング生成**へ作り替え、**(E)バッチ埋め込みで取込を高速化し、(F)ベクトル圧縮で本棚上限を引き上げる**ことで、この3+1段すべてを底上げする。

---

## B. チャンク化の再設計(構造認識 + メタデータ前置 + HyDE付与)

### B-0. 方針と互換性
- `modChunker.ChunkPages`(§7.2)の**公開シグネチャは不変**(`closed:True`・PURE_LOGIC・R4)。新ロジックは**モジュール内Privateの分岐**+新configキー `chunk_mode` で切替(既定は当面 `legacy` = 現行、検証後 `structure` へ昇格)。
- 純ロジック維持(R4): Excelオブジェクトトークン禁止。構造認識は文字列処理のみで実装可能。
- テスト可能化のため、境界判定・見出し判定の中核を **Privateから純関数へ切り出し**、テスト用に限定Public化する(§7.8が `DiffDecision`/`ResolveDecision` で採った手法と同じ。CONTRACT更新は §G-T1 で明示)。

### B-1. 構造認識チャンク化(`chunk_mode=structure`)
`ExtractedPage.Text` を行単位に分解し、各行を**構造ラベル**へ分類してから、意味単位(見出し配下のブロック)を尊重して詰める。

**行分類ヒューリスティック(純関数 `ClassifyLine(line) As Long`)**:
| ラベル | 判定(日本語社内文書想定) |
|---|---|
| `LT_HEADING_DOC` | Markdown `# `、全角『第N編/第N章』、`【…】`のみの行 |
| `LT_HEADING_SEC` | `第N条`、`第N節`、`N.` / `N.N` 番号見出し、`■`/`●`/`◆`始まりの短行(句点なし・40字未満) |
| `LT_ITEM` | `第N項`、`(N)`/`（N）`/`①〜⑳`、`・`/`- `始まりの箇条書き |
| `LT_TABLE` | タブ・複数連続スペース・`|` を含み数値列が並ぶ行 |
| `LT_BODY` | 上記以外の本文行 |

**分割規則(`ChunkOnePageStructured`)**:
1. `LT_HEADING_*` を**ブロック境界**とする。直近の見出し階層(資料名 > 章 > 条)を `outline` としてスタック保持する。
2. 見出し配下の本文を、既存の `targetChars`(既定1800)を上限に詰める。ただし**条文(`LT_HEADING_SEC`〜次見出しまで)は可能な限り1チャンクに収める**(`chunk_max_chars`=既定1800まではオーバーラップ分割せず原子的に保持。超過時のみ従来の文境界スライディング分割へフォールバック)。
3. `LT_TABLE` の連続行は行を跨いで分割しない(表の1行が途中で切れない)。
4. オーバーラップは**同一ブロック内でのみ**適用(見出しを跨いだ無意味な重なりを作らない)。
5. `full_text` は 32000字上限を従来どおり `MAX_CHUNK_CHARS` で絶対保証。

### B-2. メタデータ前置(breadcrumb を本文に物理結合)
埋め込み前に、各チャンクの `full_text` の**先頭に breadcrumb を物理的に前置**する:

```
【<資料名> > <章> > <条/節>】
<チャンク本文>
```

- 生成は `modChunker`(outlineスタックから breadcrumb 文字列を組む純関数 `BuildBreadcrumb(sourceName, outline) As String`)。ただし**資料名は付番時に確定**するため、`modShelf.IngestFile` の採番ループ(§7.2 手順6)で `ShelfChunk` に持たせた `outline`(新フィールド不要にするため後述)から breadcrumb を合成して `full_text` 先頭へ結合してから書込む。
  - **実装選択(破壊最小)**: `ShelfChunk` 型(§7.1)に列を足さず、`modChunker` が breadcrumb を **`full_text` 先頭に既に埋め込んだ状態**で返す設計にする(資料名部分だけプレースホルダ `〔資料〕` とし、`modShelf` が実ファイル名へ置換)。これにより `modTypes`/データモデル/pack形式は不変。
- **効果の連鎖**: breadcrumb入りの `full_text` が、①埋め込み入力(文脈付きで検索精度↑)、②プロンプト本文(`modPrompts.SourceBody` がそのまま使う=根拠に出典文脈が付く)、③`preview`(先頭120字=出典先出しに章・条が見える)すべてに効く。`chunk_id` ハッシュは `NormalizeForHash(full_text)` で決定的なまま(重複排除も従来どおり機能)。
- breadcrumb を検索本文から**除いた本文だけを表示したい**ニーズには、`modPrompts` 側で先頭の `【…】\n` 行を剥がす純関数を用意(表示整形のみ、任意)。

### B-3. 想定質問(HyDE)生成の付与
チャンクが「答えになりうる質問」を数個生成し、**埋め込み入力に混ぜる**ことで、質問クエリとのベクトル距離を縮める(HyDEの取込側適用)。

- 生成は `modEnrich` を拡張(§7.7骨格を最大流用)。`EnrichPending` のバッチLLM出力に `hyde` を追加:
  期待JSON `[{"i":1,"summary":"…","keywords":"a,b","hyde":"Q1 / Q2 / Q3"},…]`。
- 保存先: 既存列を汚さないため **`keywords` 列に統合**する(`keywords` は検索のレキシカル面でも使うため一石二鳥)。フォーマットは `keywords` 本体 + `\n#想定:` + HyDE。`modEnrich` の `KEYWORDS_MAX_CHARS`(300)を `enrich_mode=full` 時のみ 800 へ引き上げる(config `enrich_keywords_max`)。
- 生成本数: config `enrich_hyde_count`(既定3)。`enrich_mode=off` は従来どおり何もしない。`light`=summary+keywordsのみ、`full`=+HyDE。
- **埋め込み入力の組立**(§E で詳述): `embed_input=enriched` のとき `GetEmbedding` へ渡す文字列は `breadcrumb+本文+(想定質問)`。

### B-4. Bで新設する config キー
| key | 既定 | 意味 |
|---|---|---|
| `chunk_mode` | `legacy`(検証後 `structure`) | チャンク化方式 |
| `chunk_target_chars` | 700 | ブロック内詰め上限(現行ハードコードのconfig化) |
| `chunk_overlap_chars` | 150 | オーバーラップ |
| `chunk_max_chars` | 1800 | 条文を原子保持する上限(超で分割) |
| `embed_prefix_breadcrumb` | TRUE | breadcrumb前置ON/OFF |
| `enrich_hyde_count` | 3 | 想定質問生成本数 |
| `enrich_keywords_max` | 800 | full時のkeywords列上限 |

---

## C. 検索の再設計(多段RAG: 拡張 → マルチクエリ → 再ランク → 生成)

### C-0. 全体シーケンス(deep時の最大構成)
```
質問 + 会話履歴(prevU/prevA)
  └(1) クエリ拡張[API#1] : 独立クエリ化 + サブクエリN本 + HyDE仮回答
        └(2) マルチクエリ検索 : 各クエリをGetEmbedding→全件cosine→候補プール(union)
              └(3) 再ランク[API#2] : LLMが候補をリスト単位で関連度採点→上位Kへ
                    └(4) 最終生成[API#3] : グラウンディング生成(deepはdraft→verifyの2段)
```
quick時は軽量化(§D-3)。既存の cosine検索コア(`modRetrieve`)・出典先出し・FOLLOWUPは温存する。

### C-1. クエリ拡張(API#1)
- `modPrompts` に **新Public** `BuildExpandPrompt(q As String, history As String) As String`(純関数)。出力契約:
  ```
  <standalone>会話履歴を踏まえた独立質問</standalone>
  <subqueries>言い換え1 | 下位概念2 | 関連語3</subqueries>
  <hyde>この質問に理想的に答える1段落(検索用の仮回答)</hyde>
  ```
- 会話履歴は既存 `modAsk.HistoryBlock()` を流用(follow-up時は `prevU/prevA` もLLMに渡す既存経路)。**独立質問化がfollow-upの核**(「それの例外は?」→「〇〇制度の例外は?」)。
- パースは純関数(§H): `modRagParse.ParseExpand(resp, ByRef standalone, ByRef subs(), ByRef hyde)`(新pureモジュール、§G-T7)。マーカー欠落時は**質問文そのものを standalone・subs空・hyde空へ安全退化**(mock/形式崩れでも壊れない=`modAsk` のFOLLOWUP寛容実装と同じ流儀)。

### C-2. マルチクエリ検索(候補プール生成)
- `modRetrieve` に **新Public** `SearchExpanded(queries() As String, ByVal poolK As Long, ByRef hits() As Hit) As Long`。
  - 各 `queries(i)` を `GetEmbedding` し、**既存の全件内積スキャンを再利用**(現 `Search` のスコアリング本体を Private `ScoreAgainstStore(qv, ...)` へ切り出し、`Search`/`SearchExpanded` の双方から呼ぶ=**現行ロジックの流用**)。
  - 各クエリの上位を **chunk_id で union**(重複は最大スコア採用)。プール上限 `poolK`(config `multi_candidates` 既定40)。
  - キーワードボーナスは**改良**: `TokenizeQuery` に加え、日本語で語が取れないとき **breadcrumb/keywords との2-gram/部分一致** を補助信号にする(純関数化してテスト)。ボーナス上限は据置(0.15)。
- `queries` は `[standalone] + subs() + [hyde]`(hydeは仮回答なのでチャンクベクトルと同空間で強力)。
- 既存 `Search`(単段)は**残す**(quickの軽量経路・後方互換・既存テスト維持)。

### C-3. 再ランク(API#2)
- `modPrompts` に **新Public** `BuildRerankPrompt(q As String, hits() As Hit, nHits As Long) As String`(純関数)。候補に連番を振り、LLMへ「質問への関連度で並べ替え、上位を返せ」と指示。出力契約(パース容易・寛容):
  ```
  <rank>3,1,7,2</rank>   ' 関連度降順の連番。無関係は除外してよい
  ```
- パースは純関数 `modRagParse.ParseRankOrder(resp, ByVal nHits, ByRef order()) As Long`(範囲外・重複・欠番を無視、空なら**元順を維持**=退化安全)。
- オーケストレーションは `modAsk`(新Private `RunRerank`)。再ランク後、上位 `rerank_keep`(config、quick=`topk_quick`/deep=`topk_deep` に合わせる)を最終生成へ渡す。`rerank_enabled=FALSE` で丸ごとスキップ(候補プールをスコア順のまま採用)。
- モデル/パラメータ: config `rerank_model`(既定 `quick_model`)、`rerank_effort`/`rerank_verbosity`(既定 `low`/`low`。判定タスクは軽く速く)。

### C-4. 会話履歴(prevU/prevA)との統合・FOLLOWUP整合
- 履歴は**2箇所で使う**(役割分担を明確化):
  1. **クエリ拡張(C-1)**: `standalone` 化の材料。ここが follow-up 解決の主戦場。
  2. **リボン `CallLLM` の prevU/prevA**: 既存どおり生成段へ渡す(会話の口調・文脈連続性)。
- 既存の `modAsk.AskFollowup` / `CanFollowup` / `mPrevU`/`mPrevA` / `[[FOLLOWUP:…]]` は**シグネチャ不変で温存**。多段化は `AnswerWithContext` 内部の差し替えのみ(公開契約に影響しない)。
- 深掘り候補ブロックの生成指示(`FollowupInstruction`)は最終生成プロンプトに従来どおり付与。`<answer>` タグ導入(§D)との**併存**を明記(候補行は `<answer>` の外・末尾に出す指示にする)。

### C-5. Cで新設する config キー
| key | 既定 | 意味 |
|---|---|---|
| `retrieve_mode` | `single`(検証後 `multi`) | 検索方式(single=現行Search / multi=多段) |
| `expand_enabled` | TRUE | クエリ拡張ON |
| `expand_subqueries` | 3 | サブクエリ本数 |
| `expand_model`/`expand_effort`/`expand_verbosity` | quick_model/low/low | 拡張段パラメータ |
| `multi_candidates` | 40 | 候補プール上限(poolK) |
| `rerank_enabled` | TRUE | 再ランクON |
| `rerank_model`/`rerank_effort`/`rerank_verbosity` | quick_model/low/low | 再ランク段 |

---

## D. 回答生成の再設計(グラウンディング強化 + 2速割当)

### D-1. グラウンディング強化プロンプト
`modPrompts` の3本(`BuildQuickPrompt`/`BuildDeepDraftPrompt`/`BuildDeepVerifyPrompt`)を、共通の**強化System指示**へ寄せる(既存の `CitationInstruction`/`NotFoundInstruction` を強化):
1. **資料のみ**: 「本棚抜粋に書かれた情報**のみ**で回答。外部知識・推測での補完を禁止。抜粋に無ければ埋めない。」
2. **出典必須**: 各主張の直後に `[本棚:ファイル名 p.N]`(pack由来 `[パック(作成者):ファイル名]`)。**出典を1つも付けられない主張は書かない**。形式は §7.3 と breadcrumb を根拠に。
3. **構造化出力**: 
   ```
   <thinking>抜粋のどれが根拠か、矛盾はないかの検討(利用者非表示)</thinking>
   <answer>出典付きの最終回答本文</answer>
   ```
4. **リファューザル契約**: 「抜粋から判断できない場合は、無理に答えず **『資料からは判断できません』** とだけ述べ、何があれば答えられるか(どの資料を追加すべきか)を1行添える。」
5. 深掘り候補 `[[FOLLOWUP:…]]` は `<answer>` の**外**・末尾に(既存 §7.3/D11 と整合)。

### D-2. 応答パースと表示(RenderAnswer契約は不変)
- `modAsk` に純関数 `modRagParse.ExtractAnswer(resp, ByRef thinking, ByRef answer)`:
  - `<answer>…</answer>` があれば中身を採用、`<thinking>` は `debug_mode` 時のみ `usage_log`/`err_log` へ(利用者非表示)。
  - **タグ欠落時は応答全体を answer として採用**(mock・旧モデル・形式崩れでも壊れない)。
- 既存の `DecorateWithFollowups`/`SplitFollowupTrailer` は answer 抽出**後**に適用(順序: タグ抽出→FOLLOWUP分離→履歴保存は clean answer)。`RenderAnswer` シグネチャ不変。

### D-3. deep/quick への割当(段数の割当)
| 段 | quick(⚡すぐ聞く) | deep(🔍しっかり調べる) |
|---|---|---|
| クエリ拡張(API#1) | `expand_light`: standaloneのみ(subqueries=0・hyde省略可) | full(standalone+subs+hyde) |
| 検索 | multi(拡張が軽い分クエリ少) or single | multi(全クエリ) |
| 再ランク(API#2) | `rerank_enabled` 次第(既定ON・effort=low) | ON(effort=low〜medium) |
| 生成 | 単発 grounded(`quick_effort`/`quick_verbosity`) | draft(`deep_draft_*`)→verify(`deep_verify_*`) |
| 目安API回数 | 拡張1 + 埋め込み1〜2 + 再ランク1 + 生成1 ≒ **4〜5** | 拡張1 + 埋め込み3〜4 + 再ランク1 + draft1 + verify1 ≒ **7〜8** |

- quickは「体感10〜20秒」を守るため拡張・再ランクを軽く(effort=low、subqueries=0〜1)。オーナーが速度より品質を採るなら config で quick も full 相当へ引き上げ可能。
- 生成モデルは既存 config 駆動(`quick_model` 既定 `gpt-5.6-luna` / `recommended_model` 既定 `gpt-5.6-terra` / `thorough_model` 既定 `gpt-5.6-sol`、reasoning_effort/verbosity対応)を維持。
  **`expand_model`/`rerank_model` は空欄既定なので `quick_model` を指す。`quick_model` を変えるとクエリ拡張段と再ランク段も連動して変わる**(R48)。

### D-4. Dで新設する config キー
| key | 既定 | 意味 |
|---|---|---|
| `answer_tags` | TRUE | `<thinking>`/`<answer>` 構造出力 |
| `strict_grounding` | TRUE | 資料のみ・出典必須・リファューザルの強制文を挿入 |
| `quick_expand_light` | TRUE | quickでは拡張を軽量化 |

---

## E. 取込パイプラインの高速化

### E-1. バッチ埋め込み(直接API時)と2モード抽象化
オーナー裁定に従い、`modGateway` に埋め込み**バッチ関数を新設**し、実装は設定切替(裁定事項):

- **新Public** `modGateway.GetEmbeddingsBatch(texts() As String, ByRef outCsv() As String) As Long`
  - 戻り値=成功件数。`outCsv(i)` は `texts(i)` の **vector_csv(選択精度でシリアライズ済み・L2正規化済み)**。失敗要素は空文字。
  - `embed_transport=ribbon`(既定・安全): 内部で **`GetEmbedding` を1件ずつループ**(現行と完全同挙動。リボンのみ・キー非埋込)。
  - `embed_transport=direct`: **Azure OpenAI の embeddings エンドポイントへ配列POST**(1リクエストで最大 `embed_batch_size`(既定128)テキスト)。HTTPは `MSXML2.ServerXMLHTTP` / `WinHttp.WinHttpRequest.5.1` の遅延バインド(`CreateObject`)。R3の「リボンApplication.RunはmodGateway内のみ」に反しない(直接HTTPもmodGateway内に閉じ、opt/機能層は `GetEmbeddingsBatch` だけを見る)。
  - **キー管理リスク**: `direct` はエンドポイント/キーを config(`azure_embed_endpoint`/`azure_embed_key`/`azure_embed_deployment`/`azure_api_version`)へ保存=**ブックにキーが埋まる**。config シートは `veryHidden` ではなく `hidden`(§4)なので、`direct` 採用時は「ブック配布=キー配布」になる旨を **`使い方`/運用ドキュメントに警告** し、既定は `ribbon` に固定(オーナーが明示的にONにするまで直接HTTPは走らない)。
- `modEmbed.EmbedPending` の書込みループを、**未埋め込みを `embed_batch_size` 件ずつ集めて `GetEmbeddingsBatch` に渡す**形へ改修(再開可能性=都度 `embedded=1` 確定・chunk_id起点の書込み先再解決・ESC中断は**現行の安全策を維持**)。バッチ内の各成功要素をその場で `my_vectors` 追記+`embedded=1`。`ribbon` モードではバッチ=逐次なので現行と同じ堅牢性。

### E-2. リボン単発時の現実的な速度見積り(本線=ribbon)
| チャンク数 | 1件あたり往復(概算) | 逐次所要(概算) |
|---|---|---|
| 5,000 | 0.3s | **約25分** |
| 20,000 | 0.3s | 約100分(1.6時間) |
| 50,000 | 0.3s | 約250分(4時間) |

- 逐次は**再開可能バッチ**(ESC/上限検知で中断→翌日「🔄同期」で続き)なので、初回の大量取込は数日に分けられる設計。`embed_sleep_ms` は現行どおり既定0(mock時0)。
- **direct(バッチ)概算**: 128件/リクエスト・1リクエスト0.5s → 5,000で約40リクエスト≒**20〜40秒**、50,000で**3〜7分**。桁違いに速い。→ 大規模本棚を短時間で作りたいならオーナーが `direct` を選ぶ動機になる(トレードオフはキー管理)。

### E-3. enrich(要約・キーワード・想定質問)を走らせる段階
埋め込み入力に **breadcrumb + HyDE** を含めたい(§B-3)ため、原則 **enrich → embed の順**が最良。ただし enrich はLLMバッチで遅い。3運用を config `enrich_stage` で選べるようにする:
| `enrich_stage` | 流れ | 特徴 |
|---|---|---|
| `off` | enrichしない(現行) | 最速・HyDE無し |
| `pre`(推奨・品質優先) | 取込時: chunk→**enrich**→embed(enriched入力) | HyDE込みで検索最良。取込は遅い |
| `post`(速度優先) | 取込時: chunk→embed(breadcrumb+本文)→後で `EnrichPending`→**再embed** | 取込は速い。HyDE反映は再embed後 |
- `post` の再embed対象特定のため、`my_knowledge.embedded` を **`0=未/1=済/2=要再embed(enrich更新)`** の3値へ拡張(§4のembedded列意味を拡張。`modEmbed.PendingCount`/`EmbedPending` は `<>1` を対象とする現行判定と自然整合)。パック互換は `embedded` を出力しないため影響なし。
- 既定は破壊回避のため `off`(現行維持)。検証後 `pre` を推奨として昇格。

### E-4. Eで新設する config キー
| key | 既定 | 意味 |
|---|---|---|
| `embed_transport` | `ribbon` | ribbon(安全)/direct(バッチ・要キー) |
| `embed_batch_size` | 128 | direct時の1リクエスト件数 |
| `embed_input` | `fulltext`(検証後 `enriched`) | 埋め込み入力の組立 |
| `enrich_stage` | `off`(検証後 `pre`) | enrichを走らせる段 |
| `azure_embed_endpoint`/`azure_embed_key`/`azure_embed_deployment`/`azure_api_version` | (空) | direct時のみ使用 |

---

## F. 容量設計(ベクトル圧縮 × チャンク上限のマトリクス)

### F-1. 圧縮オプションと 1チャンクあたり `vector_csv` サイズ(概算)
現行: 1536次元 × `Str$` フル精度 ≒ **20〜27KB/チャンク**(オーナー実測、以下は中央値24KBで試算)。

| 圧縮オプション | 実装(`modUtil.VectorToCsv`) | 1チャンク | 精度影響 |
|---|---|---|---|
| **フル精度**(現行) | `Trim$(Str$(x))` | ~24KB | 基準 |
| **6桁丸め** | `Format$(x,"0.000000")`(ロケール非依存要注意→後述) | ~14KB(**-42%**) | cosine変化 < 1e-5。実質無影響 |
| **768次元(MRL切詰め)** | 先頭768成分を採用し**再L2正規化** | ~12KB(**-50%**) | ベンチで再現率 -1〜2%程度 |
| **768次元 + 6桁** | 上記2つの併用 | ~7KB(**-71%**) | 上記2要因の合算(実用上許容) |

- **Matryoshka(MRL)**: text-embedding-3-small は先頭次元だけを使っても意味が保たれるMRL特性を持つ。`direct` なら API の `dimensions=768` 指定が最良。`ribbon`(1536固定)では**取得後に先頭768へ切詰め+再正規化**で近似(`modUtil` に `TruncateAndRenorm(vec, 768)`)。
- **注意(6桁丸めのロケール)**: `Format$` はロケール依存の小数点になり得るため、`modUtil` 既存方針(`Str$`/`Val`)に倣い **`Str$` ベースで6桁に丸める純実装**にする(`Val` 復元はロケール非依存)。この丸め往復はテスト必須(§H)。
- 次元変更は `embed_dim`(§5)を768へ。**pack互換**は `embed_dim` 検査(E0702)で保護済み=次元が違うpackは弾かれるので安全。既存ベクトルは**全再embed**が必要(§G-T10)。

### F-2. マトリクス(合計サイズ/検索時間/品質 — 概算)
「合計サイズ」は `my_vectors` の**セル文字列合計(非圧縮=メモリ載せ替え時の実効量)**。.xlsm はzipで**ディスク上は概ね30〜50%**に縮む。32bit Excel の危険は**メモリ**(スキャン時に Variant 配列で二重に載る≒表の約2倍)。

| 精度＼上限 | 5,000 | 10,000 | 20,000 | 50,000 |
|---|---|---|---|---|
| フル精度(~24KB) | ~117MB | ~234MB | ~469MB ⚠️32bit危険 | ~1.17GB ❌不可 |
| 6桁(~14KB) | ~68MB | ~137MB | ~273MB ⚠️ | ~684MB ❌(要streaming) |
| 768次元(~12KB) | ~59MB | ~117MB | ~234MB | ~586MB ⚠️(要streaming) |
| 768+6桁(~7KB) | ~34MB | ~68MB | ~137MB | ~342MB ⚠️(要streaming) |

**検索時間(VBA全件内積・1クエリベクトルあたり概算)**: 5,000×1536 ≒ 0.5〜1.5s。線形にスケールし、50,000×1536 ≒ 5〜15s/ベクトル。マルチクエリ(4本)ではこれの約4倍→**大規模では致命的**。対策2つ:
1. **768次元**: 内積コストを半減(かつメモリ半減)。
2. **ストリーミング・スキャン**(`vector_scan_mode=streaming`): `my_vectors` を `vector_block_rows`(既定2000)行ずつ読み→スコア→破棄。**Variant二重載りを避け**、32bitのメモリ7エラーを回避。top-kはブロック跨ぎで維持(現行のストリーミングtop-kロジックをブロック境界へ拡張)。

### F-3. オーナー提示用の3案(推奨付き)
| | Plan A(安全・現状延長) | **Plan B(推奨・バランス)** | Plan C(大規模) |
|---|---|---|---|
| 精度 | フル精度 | **768次元 + 6桁丸め** | 768次元 + 6桁丸め |
| スキャン | メモリ | メモリ | **ストリーミング** |
| 上限(`shelf_max_chunks`) | 10,000 | **20,000** | 50,000 |
| メモリ実効 | ~234MB(64bit前提) | **~137MB(32bitでも安全域)** | ~342MB(streamingで安全) |
| 取込 | ribbon逐次 | ribbon逐次 or direct | **direct推奨** |
| 品質影響 | 最小 | 軽微(再現率-1〜2%を多段検索が相殺) | 軽微 |
| 再embed | 不要 | **必要(全件)** | 必要(全件) |
| 32bit適性 | ⚠️1万まで | **◎** | ◯(streaming前提) |

- **推奨=Plan B**: 32bit実機(メモリ7エラー既発)でも安全に**上限を現行5000→2万へ4倍**にでき、768+6桁で保存量を約1/3.4に圧縮、品質低下は多段検索(§C)が相殺する。direct導入で取込も高速化可能(任意)。
- Plan C は「限りなく増やしたい」を満たすがstreaming実装+direct取込が前提。Plan B を先行し、上限だけ config で引き上げてCへ移行する経路にする。

### F-4. Fで新設する config キー
| key | 既定 | 意味 |
|---|---|---|
| `vector_precision` | `full`(検証後 `d6`) | full / d6(6桁) |
| `embed_dim` | 1536(Plan B/Cで768) | 既存キー・次元 |
| `vector_scan_mode` | `memory` | memory / streaming |
| `vector_block_rows` | 2000 | streaming時のブロック行数 |
| `shelf_max_chunks` | 5000(案採用で10000/20000/50000) | 既存キー・上限 |

---

## G. 段階的実装計画(Sonnetワーカー割当タスク)

**共通の受入ゲート(全タスク)**: `tools/vba_lint.py` 違反0 / LO で全モジュール構文コンパイル成功 + `RunAllPureTests` 全緑(既存84本を含む) / 新config既定値で**従来動作が完全維持**されること(=各タスクは既定OFFで無影響)。新Public追加時は **`tools/vba_lint.py` の `CONTRACT`(§対象モジュール)** と **MASTER_SPEC §7 の該当契約** を同PRで更新する(`closed:True` モジュールは必須)。R4対象(`modChunker`/`modPrompts`/新pure)は禁止トークン検査を通すこと。

タスクは**依存順**に並べる(先行タスクの緑化を確認してから次へ)。

### G-T0 契約・設定の下地(先行必須)
- 対象: `MASTER_SPEC.md`(§5 configキー台帳・§7契約)、`tools/vba_lint.py`(CONTRACT/PURE_LOGIC_MODULES)、`build/modules.json`(新pureモジュール登録)、config シード(build)。
- 内容: §B〜F の新configキーを台帳へ追記(既定値=従来動作)。新Public予定(`GetEmbeddingsBatch`/`SearchExpanded`/`BuildExpandPrompt`/`BuildRerankPrompt`/`modRagParse.*`)をCONTRACTへ**プレースホルダ登録**。`modRagParse` を PURE_LOGIC_MODULES と modules.json(role:core)へ追加。
- 受入: lint/LO緑(コード変更なしで台帳・CONTRACT整合)。config既定で全既定OFF。

### G-T1 構造認識チャンク化(純ロジック)
- 対象: `src/ingest/modChunker.bas`。
- 内容: `chunk_mode` 分岐、`ClassifyLine`/`ChunkOnePageStructured`/`BuildBreadcrumb`(breadcrumbプレースホルダ前置)を追加。テスト用に境界判定の中核を限定Public化(CONTRACT追記)。`ChunkPages` シグネチャ不変。
- 受入: `chunk_mode=legacy` で現行と**バイト一致**の出力。`structure` で §H の新テスト緑。32000字上限保証を維持。

### G-T2 breadcrumb 実結合(取込)
- 対象: `src/ingest/modShelf.bas`(採番ループ 手順6)。
- 内容: `embed_prefix_breadcrumb=TRUE` 時、`full_text` 先頭プレースホルダ `〔資料〕` を実 `sourceName` へ置換して書込む。`chunk_id` ハッシュは結合後 `full_text` で計算(重複排除は従来どおり決定的)。
- 受入: 既定OFFで現行一致。ONで `my_knowledge.full_text` 先頭に breadcrumb 行が入り、`preview`/検索/プロンプトへ波及。IngestFileの再入guard・E0504・partial判定は不変(既存E2Eテスト緑)。

### G-T3 enrich拡張(HyDE/keywords)
- 対象: `src/ingest/modEnrich.bas`、`src/qa/modPrompts.bas`(`BuildEnrichPrompt`)。
- 内容: enrich出力JSONに `hyde` 追加、`enrich_mode=full` 時に `keywords` 列へ HyDE 統合、`enrich_stage` 導入。パーサ(`ParseEnrichJson`)へ `hyde` キー抽出を追加(現行の寛容実装を踏襲)。
- 受入: `enrich_mode=off` で完全無動作(現行)。`full` で keywords に想定質問が入る。既存 enrich テスト緑。

### G-T4 埋め込み入力組立 + embedded 3値
- 対象: `src/ingest/modEmbed.bas`。
- 内容: `embed_input=enriched` 時に breadcrumb+本文+(keywords内HyDE) を埋め込み入力へ。`embedded` を3値化(2=要再embed)。`enrich_stage=post` の再embed経路。
- 受入: 既定 `fulltext` で現行一致。再開可能性・ESC・chunk_id起点書込みの安全策維持。

### G-T5 Gateway バッチ埋め込み(2モード抽象化)
- 対象: `src/core/modGateway.bas`。
- 内容: `GetEmbeddingsBatch`(新Public)。`ribbon`=単発ループ / `direct`=HTTP配列POST(遅延バインド)。mock時は決定的擬似ベクトルをバッチ返し。`vector_precision`/`embed_dim` に応じたCSV生成は `modUtil` 経由。
- 受入: `embed_transport=ribbon`(既定)で現行と同結果。mockでバッチ経路が動く。direct はキー未設定なら安全にエラー(#ERR)へ退化。CONTRACT/MASTER_SPEC §7.1更新。

### G-T6 マルチクエリ検索
- 対象: `src/qa/modRetrieve.bas`。
- 内容: スコアリング本体を Private `ScoreAgainstStore` へ切出し、`Search`(現行・不変)と新Public `SearchExpanded` から共用。union候補プール、日本語向けキーワードボーナス改良、`vector_scan_mode`(streaming)対応。
- 受入: `retrieve_mode=single` で `Search` 挙動不変(既存テスト緑)。`SearchExpanded` の union/上限/退化(クエリ0本→Search同等)をテスト。E0702次元検査は維持。

### G-T7 純パースモジュール(新規)
- 対象: `src/test/`ではなく `src/qa/modRagParse.bas`(新pure・R4)、`src/test/modTestsPure2.bas`(テスト追記)。
- 内容: `ParseExpand`/`ParseRankOrder`/`ExtractAnswer`/`ParseSubqueries`。全て寛容退化(マーカー欠落=安全既定)。
- 受入: §H の純テスト緑。PURE_LOGIC禁止トークン0。LOのType配列制約を受けない(String配列のみ)ことを確認。

### G-T8 プロンプト新設(拡張・再ランク・生成強化)
- 対象: `src/qa/modPrompts.bas`。
- 内容: `BuildExpandPrompt`/`BuildRerankPrompt`(新Public・pure)、既存3本へ `strict_grounding`/`answer_tags` の指示を条件付き挿入(既定TRUEだが挿入文は純文字列)。FOLLOWUP指示と `<answer>` 外出しの整合。
- 受入: 既存 modPrompts テスト(出典形式・打ち切り)緑。新プロンプトの必須要素(タグ・出典・リファューザル文言)をテスト。CONTRACT更新。

### G-T9 modAsk 多段オーケストレーション
- 対象: `src/qa/modAsk.bas`。
- 内容: `AnswerWithContext` 内部を「拡張→SearchExpanded→再ランク→生成」へ差し替え(config段でON/OFF)。`ExtractAnswer`→`DecorateWithFollowups` の順。公開契約(`Answer`/`AskFollowup`/`CanFollowup`/Feedback系)不変。RenderSourcesPreview/RenderAnswer呼び出し不変。
- 受入: `retrieve_mode=single`+`expand_enabled=FALSE`+`rerank_enabled=FALSE`+`answer_tags=FALSE` で**現行フローに完全一致**。段を1つずつONにしても mock でE2Eが一巡(既存 modTestsExcel の mock取込→検索→回答が緑)。ESC/エラー/検索0件/空質問の既存挙動不変。

### G-T10 容量(圧縮・streaming・再embed運用)
- 対象: `src/core/modUtil.bas`(`VectorToCsv` 精度分岐・`TruncateAndRenorm`)、`src/qa/modRetrieve.bas`(streaming確定)、運用ドキュメント(再embed手順)。
- 内容: `vector_precision`/`embed_dim=768`/`vector_scan_mode`/`shelf_max_chunks` 引上げ。次元変更時の**全再embed導線**(診断ボタンから「ベクトル再構築」= `embedded` 全リセット→ `EmbedPending`)。
- 受入: 圧縮往復テスト(§H)緑。`memory`/`streaming` で同一top-k。次元不一致の残存ベクトルはE0702で検出(現行維持)。pack次元検査(E0702)で相互運用保護。

> **並行可能性**: T1/T3/T7 は相互独立で並行可。T5(Gateway)はT4の前提。T6→T9、T8→T9。T10は最後(または独立でPlan A維持のまま先行導入可)。各タスクは**単独マージで緑**を保つ粒度。

---

## H. 追加すべき純ロジックテスト一覧

`modTestsPure`/`modTestsPure2`(§7.8、LO+Excel両対応)へ追加。LOの「別モジュールPublic Type配列ReDim不可」制約(既知)に該当するもの(`ShelfChunk()`/`Hit()` を取る)は `CanUseTypeArrays()` ガード配下に置き、**String配列で完結するパーサ系は無条件で実行**する。

### H-1. modChunker(構造認識)
- 見出し検出: `第3条`/`■方針`/Markdown `# ` を `LT_HEADING_*` に正しく分類。
- 条文原子性: `chunk_max_chars` 未満の1条が**1チャンクに収まる**(途中分断されない)。
- 条文超過: `chunk_max_chars` 超で文境界スライディングへフォールバック分割。
- 段落境界優先: 見出しを跨ぐオーバーラップが発生しない。
- 表保持: `LT_TABLE` 連続行が分割されない。
- breadcrumb: `BuildBreadcrumb` が `【資料 > 章 > 条】` 形式・階層欠落時の縮約(章のみ等)を正しく生成。プレースホルダ `〔資料〕` の位置。
- 32000字上限: 巨大条文でも `full_text ≤ 32000`。
- legacy一致: `chunk_mode=legacy` が現行出力とバイト一致(回帰防止)。

### H-2. modUtil(容量・圧縮)
- `VectorToCsv` 6桁丸め: フル→6桁→`CsvToVector` 往復で各成分誤差 ≤ 5e-7、ロケール非依存(小数点`.`固定)。
- `TruncateAndRenorm(vec,768)`: 出力次元768・L2ノルム=1(±1e-9)・先頭768成分の方向保存。
- 圧縮後cosine保存: 同一ベクトル対のcosineが フル vs 6桁で差 ≤ 1e-5。

### H-3. modRagParse(新pure・パーサ寛容性)
- `ParseExpand`: 正常/マーカー一部欠落/全欠落(→standalone=元質問・subs空・hyde空)/`<subqueries>` の `|` 分割・空要素除去。
- `ParseRankOrder`: `<rank>3,1,7</rank>` を順序配列へ/範囲外・重複・非数値の無視/空→元順維持/`nHits` 境界。
- `ExtractAnswer`: `<answer>` 抽出・`<thinking>` 分離/タグ欠落→全体をanswer/`<answer>` 複数・入れ子の安全処理。
- `ParseSubqueries`: 区切り・トリム・上限本数。

### H-4. modPrompts(新プロンプト・強化)
- `BuildExpandPrompt`: `<standalone>/<subqueries>/<hyde>` 出力指示と history 埋込。
- `BuildRerankPrompt`: 候補連番の付与・`<rank>` 出力指示・`max_context_chars` 打ち切り。
- `BuildQuick/DeepVerifyPrompt`(強化): `strict_grounding=TRUE` で「資料のみ/出典必須/『資料からは判断できません』」の文言を含む・`answer_tags=TRUE` で `<answer>` 指示を含む・FOLLOWUPが `<answer>` 外指示になっている。既存の出典形式(`[本棚:… p.N]`/`[パック(…):…]`)・打ち切り(`(一部省略)`)テストは維持。

### H-5. modEnrich(パース)
- `ParseEnrichJson` の `hyde` キー抽出(`summary`/`keywords` 抽出は現行テスト維持)・`hyde` 欠落時の後方互換。

---

## 付録: 破壊的変更を避けるための不変条件チェックリスト(実装者用)
- [ ] `modChunker.ChunkPages` / `modRetrieve.Search` / `modPrompts.Build*`(既存4本) / `modAsk`公開7本 / `modEmbed.EmbedPending`,`PendingCount` の**シグネチャ不変**。
- [ ] 新config既定値で**現行と同一挙動**(全新機能OFF)。
- [ ] `my_knowledge`/`my_vectors`/`my_manifest` の**列構成不変**(embeddedの3値化は既存 `<>1` 判定と互換・pack出力に非影響)。
- [ ] `chunk_id` 生成(`bs::fnv64::pN::cN`)と重複排除の**決定性維持**。
- [ ] R1〜R5 遵守(特にR3: 直接HTTPも `modGateway` 内に閉じる/R4: `modChunker`・`modPrompts`・`modRagParse` は純ロジック)。
- [ ] 新Public追加のたびに `vba_lint.py` CONTRACT と MASTER_SPEC §7 を同時更新。
- [ ] `direct`(埋め込み直HTTP)は既定OFF・キー管理リスクを `使い方`/運用ドキュメントへ明記。
- [ ] pack互換: `embed_dim` 変更は E0702 検査で相互運用保護(次元混在を弾く)。
- [ ] 84本の純ロジックテスト + lint + LOコンパイルが**各タスク単独マージで全緑**。
</content>
</invoke>
