# Architecture

## システム全体図

```
[管理者PC] Admin_KnowledgeBuilder.xlsm
   - 取り込み・index生成・公開・キー埋め込み・利用集計
                    |
                    | (SharePoint /knowledge_index/)
                    v
            manifest.json + embeddings.bin + chunks.json
                    |
                    | (Workbook_Open でローカルキャッシュへDL)
                    v
[営業PC] Chatbot.xlsm
   modBoot -> modIndexReader -> in-memory LoadedIndex
   frmChat -> modPiiGuard -> modRateLimiter -> modRagEngine
                                   -> modApiGateway (Embed -> TopK -> Chat)
                                   -> modUsageLogger -> SharePoint/{username}.csv
```

## モジュール責務

### shared
- `modConfig` - config.ini ローダー。`%ENV%` 展開。
- `modPaths` - LocalAppData 配下のパス決定。
- `modHttpClient` - WinHTTP ラッパ。プロキシ設定を `[proxy]` から適用。
- `modKeyVault` - Mode A/B/C の APIキー解決の抽象化。難読化 (XOR+Base64) を実装。
- `modApiGateway` - `Embed(texts)` / `Chat(sys, usr)` を提供。プロバイダ・モード切替を隠蔽。
- `JsonConverter` - VBA-JSON v2.3.1 (build時に注入)。

### admin
- `modExtractor` - 拡張子で抽出器を分岐 (Week1は .txt のみ)。
- `modChunker` - 段落配慮の overlap chunk + L2正規化ヘルパ。
- `modIndexWriter` - manifest を最後に差し替えるアトミック公開。
- `modKnowledgeBuilder` - ファイル選択→抽出→chunk→batch embed→書き出し。
- `modKeyEnroller` - 配布用xlsmへの難読化キー埋め込みUI。

### chatbot
- `modBoot` - Workbook_Open 起点。config → 初回設定 → index DL → load → frmChat。
- `modIndexReader` - manifest/embeddings/chunks を読み込み LoadedIndex に。
- `modSimilarity` - 内積 top-k (Variantゼロ)。
- `modRagEngine` - embed → topk → prompt assembly → chat。出典付き応答。
- `modPiiGuard` - 電話/メール/契約番号らしき正規表現で送信前警告。
- `modRateLimiter` - LocalAppData の Double配列 で 1h/1day 滑り窓カウント。
- `modUserProfile` - 部署/役割/同意フラグの永続化。
- `modUsageLogger` - UTF-8 BOM 付きCSV をローカル → SharePoint にフラッシュ。
- `frmChat` - 入力UI、出典表示、ステータス表示。

## データフロー

1. **管理者** が `BuildIndex` を実行 → 選択ファイルを抽出 → chunk → batch embed → L2正規化 → ローカル `IndexCacheDir` に書き出し → 必要なら SharePoint に publish。
2. **営業担当者PC** で起動 → SharePoint から差分DL → in-memory ロード → frmChat 表示。
3. **質問** → embed → top-k → プロンプト組立 → chat → 出典付きで描画 → 利用ログ追記。

## 設計判断

| 判断 | 理由 |
|---|---|
| Azure OpenAI 前提 | 社内情シスが想定する提供形態。OpenAI 直接の場合は config.ini の `provider` 切替で吸収可能。 |
| 2バイナリ分離 | 配布用に管理機能を持たせない・キー埋め込みを admin 側で完結。 |
| バイナリembeddings | JSON展開のメモリ・パース時間を回避。3000チャンクで~36MB。 |
| L2正規化を事前にやる | クエリ側で1回だけ正規化すれば top-k は単純内積。 |
| 出力ログを per-user CSV | SharePoint上の同時書き込み競合を回避。集計は Power Query で。 |
| Mode A/B/C 抽象化 | MVPは難読化埋め込み (B) で開始、運用が回り次第 A に移行する前提。 |
