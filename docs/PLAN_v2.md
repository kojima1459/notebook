# Chatbot v2 設計プラン — 社内AIリボン経由・全社展開対応

**前提**：v1（Gemini直叩き / 19MBバイナリindex）は社内環境で動かないことが判明。
向こうの稼働中チャットボットの実装パターン（`Application.Run("ChatGPT", prompt)`）に乗り換える。

---

## 1. 設計の柱（ユーザー意思決定の反映）

| 決定事項 | 設計への落とし込み |
|---|---|
| **FB保存**：個人→共有→承認 | 3状態（pending/approved/rejected）+ SharePoint同期I/F |
| **モデル**：全部gpt-5.5、可換 | Excel設定シートで切替（VBAコード変更なし） |
| **全社展開＋部署別ナレッジ** | 部署IDでナレッジセット切替、同一xlsmで多部署対応 |
| **前処理MAX** | ファイル名＋業務領域タグ＋1行要約＋キーワード自動付与 |

## 2. 全体アーキテクチャ

```
[Chatbot_v2.xlsm]  (社内環境ネイティブ、Edge配布)
│
├ シート ★main           ← UI（入力/出力/送信/👍👎）
├ シート config          ← モデル設定、部署選択、動作モード
├ シート system-prompt    ← LLMへの指示（部署別）
├ シート manifest        ← マニュアル一覧（ファイル名・領域・バージョン）
├ シート knowledge_base   ← 全チャンク（id, source, summary, keywords, text）
├ シート feedback         ← 過去Q&A（pending/approved/rejected, dept）
├ シート usage_log        ← 質問履歴・所要時間・モデル
└ シート department       ← 部署マスタ（dept_id, name, knowledge_filter）
│
└ VBA ── Application.Run("ChatGPT", prompt)  ←社内AIリボン呼出
        （4ステップ・パイプライン、後述）
```

## 3. パイプライン（4ステップ）

```
質問入力
   │
   ▼
[Step 1: ルーター]
  入力：質問 + knowledge_base の「id + 1行要約 + キーワード」だけ
        （部署フィルタ適用後の対象チャンクのみ）
  出力：関連チャンクIDのリスト（4〜8件）
  目的：質問とナレッジ全体の意味的マッチングを軽量に
   │
   ▼
[Step 2: フィードバック検索]
  入力：質問
  処理：feedback シートの approved 状態の Q&A を全件走査し、
        類似度の高いものを最大3件抽出
  出力：採用Q&Aリスト（過去の認定回答）
   │
   ▼
[Step 3: ドラフト生成]
  入力：system-prompt + 採用Q&A + 関連チャンク全文 + 質問
  出力：構造化された回答（要点→結論→詳細→次のアクション）
        各主張に出典マーカー [#N] 付与
  目的：精度重視で答案作成
   │
   ▼
[Step 4: 自己検証]
  入力：system-prompt(verifier) + 同じチャンク + 質問 + ドラフト
  処理：ドラフトの各主張を引用元と照合、未根拠/反対意味/数値捏造を削除/修正
  出力：検証済み最終回答 + 「除外/修正した内容」セクション
   │
   ▼
回答表示 + 👍/👎 ボタン
```

**所要時間目安**：60〜90秒/問（精度優先）

## 3.4 意図解釈レイヤー / 自信度 / 教えてBOX（NotebookLM級への引き上げ）

精度と「営業の自己解決・本社負荷削減」を両立するため、4ステップの**前後**に以下を追加：

**Step 0: 意図解釈レイヤー（`modIntent`）** — 質問をそのまま検索しない。
- 照会の真意(intent)・前提(assumptions)を抽出し、約款を引くためのサブクエリに分解 → ルーターへ注入。
- **質問返し（切替式）**: `intent_mode`（main画面のボタンでユーザーが切替）
  - `smart`（既定）: 明確なら即答／回答が変わる前提が欠けている時だけ確認質問
  - `always`: 毎回まず前提を確認（誤解最小化）
- 確認質問は InputBox で受け、補足はドラフターへ最優先で反映。
- この層が失敗しても本体は止めない（graceful degradation）。

**自信度（`[[CONFIDENCE:高/中/低]]`）** — 検証パスが回答末尾に機械可読で出力し、VBAが解析・除去。
- 高=約款/ガイドラインに明確な根拠／中=事例・推論が主／低=該当規定が乏しい。
- 回答冒頭にバナー表示（🟢🟡🔴）。中/低では「教えてBOX」を強調。

**教えてBOX（`modInquiryBox`）** — 本社引受部門への照会パケットを自動生成。
- 照会者・質問・AIが解釈した意図/前提・参照した出典・暫定回答・自信度・迷っている点 を整形。
- Outlook があればメール下書きを自動生成（宛先=`config!hq_inquiry_email`）、無ければコピー用シート＋クリップボード。
- **目的**: 本社は一から聞き返す工程が消え、営業も論点が整理された状態で照会できる。

**深掘り候補（`[[FOLLOWUP:...]]`）** — 検証パスが次に聞くべき質問を提案し、回答末尾に表示。`続けて質問`へ誘導。

**隠しクレジット** — 「このボットは誰が作った？」系の質問は約款と無関係なので最初に横取りし、
開発者（`config!creator_group` / `creator_name`）を回答（`modPipeline.IsCreatorQuestion`）。

## 3.7 横展開（火災・自動車・生涯保険など他グループ）

v2は最初から**ナレッジ入れ替え可能**な設計：
- `knowledge_base`（chunk）＋`dept_scope`列＋`doc_type`駆動。他グループのチャンクを入れて
  `department`／`dept_scope`を振るだけで動く。VBA・プロンプトの汎用ロジックは変更不要。
- 3.4で追加した意図解釈・自信度・教えてBOX・隠しクレジットは**全て商品非依存**で、そのまま効く。
- **精度の天井**: 現行ルーターは「要約表を丸ごとLLMに見せてID選択」方式のため、1グループ単位
  （〜1000-1500チャンク）なら精度を保てる。火災+自動車+生涯+新種を**全部同時に**可視化する規模に
  なると表が大きくなりすぎ、ハイブリッド検索（キーワード一致＋意味検索＋リランキング）への移行が必要。
- 結論：**ナレッジ入れ替え＝即可 / グループ内追加＝精度維持 / 全社横断の巨大統合＝ハイブリッド検索が前提**。

## 3.5 商品部公式Q&A レイヤーと「事実／推論の分離」（精度・出典強化）

営業からの照会に対する**商品部の公式回答まとめ**（`build/source_data/shohinbu_qa.xlsx`）を、
`doc_type = "商品部公式Q&A"` として knowledge_base に取り込む（取り込みは
`build/add_shohinbu_qa_chunks.py`）。これは feedback シート（ユーザーの○×学習）とは別物で、
**最初から信頼できる公式の回答事例**として全社共通（dept_scope=common）で参照される。

**回答フロー（似た照会が来たとき）**
1. ルーターが、同じ商品・論点の「商品部公式Q&A」を**最優先で拾い**、さらにその**裏付けとなる
   約款・引受ガイドライン**もセットで選ぶ。
2. ドラフターは商品部Q&Aを**軸（結論の方向性）**に置きつつ、約款・ガイドラインで**裏付け**てから回答する。
   Q&Aは過去事例なので、前提条件が今回の照会と一致するか／約款と矛盾しないかを照合する。
3. 検証パスが、可否の結論が第一次根拠（約款/ガイドライン）または商品部Q&Aの出典に紐づいているかを確認する。

**根拠の3階層（商品部の懸念＝「AIが意図を汲み取って誤らないか」への対策）**

| 階層 | 中身 | 扱い |
|---|---|---|
| ① 第一次根拠 | 約款・引受ガイドライン・商品規定・重説 | 引受可否/補償可否/支払可否の**最終根拠**。必ず出典 [#N]。 |
| ② 商品部公式Q&A | 商品部の公式回答事例 | **強い参考**。軸にするが①で裏付ける。Q&A単独で断定しない。出典 [#N]。 |
| ③ 補足知識 | 一般的な損保実務知識 | 可否の根拠にしない。「実務上は」と明示・出典なし。 |

**出力での担保**
- **事実（資料に明記＝[#N]付き）** と **推論（実務上の解釈＝出典なし）** を、見出し（【根拠】／【解釈・当てはめ】）で
  物理的に分離。同じ文に混ぜない。
- 可否の結論は必ず①または②の [#N] に紐づける。該当規定・事例が無ければ
  「社内ナレッジ上、明確な規定・回答事例は確認できません」と正直に返す。
- 根拠は「これでもか」と網羅（該当する条文・事例すべてに [#N]）。

> 商品部Q&Aの更新手順：`build/source_data/shohinbu_qa.xlsx` を差し替え →
> `python build/add_shohinbu_qa_chunks.py` → `python build/build_chatbot_v2.py` で再ビルド。

## 4. フィードバック蓄積・承認フロー

```
ユーザーが回答に 👍 をクリック
   │
   ▼
ダイアログ「この回答のどこが良かった？(任意)」
   │
   ▼
feedback シートに行追加
  ・ status = "pending"   (個人レベルでは即時有効＝Step2で使用)
  ・ dept_id = ユーザーの部署
  ・ submitter = Windowsログイン名
  ・ approver = (空)
  ・ approved_at = (空)
   │
   ▼ (時々 / 月次)
管理者がmainブランチで feedback を眺める
  → 良いものは status = "approved" に変更
  → 全部署 or 同一部署に展開
  → 誤りは status = "rejected" + correction 列に正答記入
   │
   ▼
SharePoint連携（Phase 2）
  feedback シートを SharePoint 上の共有Excel／リスト と同期
  全社の承認済みナレッジが各営業PCに自動配信される
```

**👎 のフロー**:
```
👎 クリック → 「正しい回答は？」入力
  → feedback に rejected + correction として保存
  → 「同じ質問が来たら correction を優先採用」と Step3 に指示
```

## 5. 部署別ナレッジ切替

```
シート department (部署マスタ)
  dept_id | dept_name        | knowledge_filter (どのsourceを使うか)
  001     | 費用利益保険チーム  | sample_*.pdf
  002     | 火災保険チーム      | fire_*.pdf
  003     | 自動車保険チーム    | auto_*.pdf

初回起動時に部署IDを Application.InputBox で取得し、
%LOCALAPPDATA%\InternalNotebookLM\user_profile.ini に保存
```

ナレッジ切替の仕掛け：
- `knowledge_base` シートの全行に `source` 列がある
- `manifest` で各 source の所属部署を定義
- Step1 のルーターに渡す前に、ユーザー部署で使えるsourceに絞る
- 共通ナレッジは `dept = "common"` でタグ

## 6. モデル可換設計

```
シート config
  key                  | value           | コメント
  router_model         | gpt-5.5         | Step1（現状全部5.5）
  drafter_model        | gpt-5.5         | Step3
  verifier_model       | gpt-5.5         | Step4
  router_max_chunks    | 8               | Step1で何件まで絞るか
  drafter_temperature  | 0.15            | ドラフト時の温度
  verifier_enabled     | 1               | 検証パスON/OFF
  ...
```

**情シスから「コスト爆発」のクレームが来たら**：
1. config の `router_model` を `gpt-5.4-nano` に変える
2. `verifier_enabled` を `0` にする  
3. それだけで30%以上コスト減

「やめろ」と言われたら：
4. `verifier_model` も `gpt-4.1-mini` に
5. それでも厳しいなら、Step1で絞る件数を5→3に

**全部Excelセル変更だけ。コード変更なし**。

## 7. 前処理パイプライン（こちらでやる）

```
[16 PDF]
   ↓
extract: pypdf で本文抽出
   ↓
chunk: 既存の article-aware chunker で構造化（785チャンク）
   ↓
enrich: 各チャンクに対して
   ・LLMで1行要約生成（〜60字）
   ・LLMでキーワード抽出（5〜10語）
   ・PDF全体に業務領域タグ付与（手動で僕がレビュー）
   ・PDF全体に書類種別タグ付与（普通約款/特約/FAQ等）
   ↓
output: knowledge_base_seed.csv（Excelにそのまま貼り付け可能）
        manifest_seed.csv
```

**所要時間**：LLMコール ~785 × 1秒（並列）= 約15分 + 僕の目視チェック30分

## 8. Excelファイル構造（より具体的）

### シート main
- Title バー
- Department表示（読み取り専用）
- TextBox_UserInput（質問）
- 送信ボタン → `Call_GPT_Inquiry`
- TextBox_Status（処理中の状況）
- TextBox_AiOutput（回答）
- 👍 ボタン → `Submit_Feedback_Good`
- 👎 ボタン → `Submit_Feedback_Bad`
- 「テキスト出力」「HTML出力」ボタン

### シート config
- 名前付きセル（VBA から読みやすく）

### シート knowledge_base（核）
- 列：chunk_id | source | dept | section_header | summary | keywords | full_text
- 行：785（初期）、追加可

### シート feedback
- 列：fb_id | timestamp | dept | submitter | question | answer | status | approver | approved_at | correction | tags

### シート usage_log
- 列：timestamp | user | dept | question_hash | router_ms | draft_ms | verify_ms | total_ms | model_used | feedback_clicked

## 9. 実装フェーズ

### Phase 1: 自分1人で動作確認（MVP、今夜〜明日）
- 16PDFの前処理＋要約付与
- Chatbot_v2.xlsm を組む（VBAは新規）
- main / config / system-prompt / manifest / knowledge_base / feedback / department を作成
- 部署 = "費用利益保険チーム" 1個だけで動作確認
- 👍/👎 は個人xlsm内に蓄積のみ
- 配布物：Chatbot_v2.xlsm 1ファイル（数MB）

### Phase 2: チームパイロット（来週）
- 費用利益保険チーム 数名にパイロット配布
- 全員が個別にxlsm持つ。feedbackは各自蓄積
- 1〜2週間運用後、feedback内容を眺めて品質チューニング
- system-prompt の改善余地を発見

### Phase 3: SharePoint連携（2〜3週後）
- 承認済み feedback を SharePoint 共有Excel に同期
- 各営業PCは起動時に SharePoint から最新 feedback を取得
- 管理者用 Approval UI（feedback承認ボタン付きのxlsm）を別ファイルで配布

### Phase 4: 全社展開（要情シス調整）
- 部署マスタを完成
- 部署横断ナレッジ（共通約款等）を整理
- マニュアル追加フロー確立（こちら → ユーザー → 承認 → 配布）
- コスト・利用ログの集計ダッシュボード

## 10. リスクと対策

| リスク | 対策 |
|---|---|
| Step1 のルーターが誤判定し関連条文を見逃す | router_max_chunks を多めに（8）、近隣条文も自動追加 |
| Step4 の検証パスが過剰に削除する | system-prompt(verifier) を保守的に。デバッグモードで両方表示 |
| feedback が肥大化して Step2 が遅くなる | 全文検索でなくキーワード前提のフィルタを追加、approved上限300件 |
| 部署別ナレッジが交差して情報漏洩 | dept_id によるフィルタを最初のStep1段階で適用 |
| 社内AIリボンの仕様変更 | `Application.Run("ChatGPT", prompt)` 1点のみ依存。互換性を保ちやすい |

## 11. オープン疑問（実装時に確認）

- リボンの `ChatGPT` 関数は、戻り値にどんな形式で返してくるか（Markdown? プレーン?）
- system-prompt 相当を分離する仕組みはあるか（向こうは1本に連結している → こちらも踏襲）
- リボン経由でモデル選択をVBAから指定できるか、それともUI設定のみか
  → 後者なら、起動時に `current_model` を確認してconfigに記録
- レート制限（同時呼び出し）

これらは Chatbot_v2.xlsm の初版を作りつつ、実機で検証する。
