# デモ Quick Start (情シス向け)

これは **動くプロトタイプ** です。社内ナレッジQAボットの完成イメージを掴むためのもので、本番想定の Azure OpenAI ではなく、**個人の Gemini API キー** で動かしています。

所要 **5〜10分** で起動できます。

## このデモで体感してほしいこと

1. **Excel マクロだけで** 社内PCの制約下 (Python不可・ターミナル不可・生成AIサイトブラックリスト) でも RAG チャットが動くこと
2. **APIキーがなぜ必要か** - このxlsmは外部のAI(Gemini)にHTTPS POSTで質問を送って回答を得ます。APIキーは「誰のアカウントに課金するか」を識別するパスポート
3. **ナレッジが社内文書から検索される** こと - 16ファイル分の保険業務PDFを事前にindex化しています
4. 本番に向けて何を情シスに依頼する必要があるか → [docs/it-checklist.md](it-checklist.md)

## 必要な前提

- Windows 10 / 11
- Microsoft Excel 2019 以降 (デスクトップ版)
- インターネット接続 (`generativelanguage.googleapis.com` への HTTPS)
- 渡された xlsm 2ファイル と APIキー

## 手順

### 1. ファイルを配置

```
%LOCALAPPDATA%\InternalNotebookLM\
├── Chatbot.xlsm            (渡されたもの)
├── Admin_KnowledgeBuilder.xlsm  (渡されたもの - 任意、ナレッジ再構築用)
├── config.ini              (渡されたもの)
├── disclaimer.txt          (渡されたもの)
└── index\                  (渡されたもの - サンプル16PDFのindex)
    ├── manifest.json
    ├── embeddings.bin
    └── chunks.json
```

エクスプローラーのアドレスバーに `%LOCALAPPDATA%\InternalNotebookLM\` と打ち込めばその場所に飛びます。フォルダがなければ作成してください。

### 2. APIキーを config.ini に貼る

`config.ini` をメモ帳で開き、`api_key =` の行に渡されたキーを貼り付けます:

```ini
[api]
mode = P
provider = gemini
api_key = AQ.Ab8R...（ここに貼り付け）
```

保存して閉じる。

### 3. Chatbot.xlsm を開く

ダブルクリック → マクロ有効化 → 初回は部署と役割を聞かれるので適当に入れる → チャット画面が出る。

### 4. 質問してみる

サンプルPDFは保険業界の文書なので、たとえば:

- 「○○保険の補償範囲は？」
- 「△△の免責事項を教えて」
- 「□□のときの請求手続きは？」

数秒で **回答 + 出典 (どのファイルの何ページから来たか)** が表示されます。

## トラブルシュート

| 症状 | 原因 / 対処 |
|---|---|
| `HTTP 401` / `HTTP 403` | APIキー貼り間違い。`config.ini` を再確認。 |
| `HTTP 429` | Gemini 無料枠 (1分15回) 超過。1分待って再試行。 |
| `index読み込みに失敗` | `%LOCALAPPDATA%\InternalNotebookLM\index\` 配下のファイルが揃っているか確認。 |
| 起動が遅い | 初回は index 全件メモリロード。3000チャンクで5〜10秒。 |
| マクロブロック | Excel → ファイル → オプション → トラスト センター → マクロの設定 で「警告を表示してすべてのマクロを無効にする」以外に。 |

## このプロトタイプの限界 (本番では解決が必要)

- **APIキーが config.ini に平文** → 本番では **Mode B 難読化埋め込み** か **Mode A 中継サーバー** に切替予定。詳細は [docs/security.md](security.md)
- **個人の Gemini キー** → 本番は **社内 Azure OpenAI Service** にする。提供形態は情シス次第 → [docs/it-checklist.md](it-checklist.md)
- **無料枠の制限** → 本番はテナント単位の有料枠
- **ナレッジは管理者PC内のローカルindex** → 本番は **SharePoint上で集中管理 + 営業PCに自動同期**
- **PIIフィルタ・利用ログ集計** はコードに入っているが、本番運用では監査の仕組みを情シスと詰める必要あり

## 次の打ち合わせ案

1. このデモを情シスに見てもらう
2. [docs/it-checklist.md](it-checklist.md) を一緒に埋める (Azure OpenAI / SharePoint / プロキシ / 予算)
3. 埋まったら本番ビルド (Mode B、Azure OpenAI) に切り替えてパイロット5〜10名で配布
