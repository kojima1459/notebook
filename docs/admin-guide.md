# 管理者ガイド

## 必要なもの

- Windows PC、Excel 2016 以降 (デスクトップ版)
- Azure OpenAI Service のリソース情報
  - エンドポイントURL
  - chat デプロイメント名 (例: `gpt-4o`)
  - embedding デプロイメント名 (例: `text-embedding-3-small`)
  - APIキー
- SharePoint または共有ドライブ上のナレッジ公開フォルダ
- SharePoint 上の利用ログ集約フォルダ

## 初回セットアップ

1. `dist\Chatbot.xlsm` と `dist\Admin_KnowledgeBuilder.xlsm` をビルド (`build/build.ps1`)。
2. `config/config.sample.ini` を `%LOCALAPPDATA%\InternalNotebookLM\config.ini` にコピーし、Azure OpenAI 接続情報、SharePoint URL を埋める。
3. 同じ config.ini をパイロット配布対象のPCにも配置する。
4. Excel の **信頼できる場所** に `%LOCALAPPDATA%\InternalNotebookLM\` を追加 (情シスにGPO配信を依頼するのが本筋)。

## APIキーを Chatbot.xlsm に埋め込む

1. `Admin_KnowledgeBuilder.xlsm` を開き、開発タブ → `modKeyEnroller.EnrollKeyInteractive` を実行。
2. InputBox に Azure OpenAI のAPIキーを貼り付け。
3. ファイル選択で `Chatbot.xlsm` を指定。
4. 自動で難読化されたBlobが Chatbot.xlsm の隠し定義名に書き込まれ、保存される。
5. 元のAPIキー文字列は破棄される (シートにもログにも残らない)。

⚠️ Mode B は **暗号化ではなく難読化** に過ぎない。詳しくは [security.md](security.md)。

## ナレッジindexを作る

1. `Admin_KnowledgeBuilder.xlsm` を開き `modKnowledgeBuilder.BuildIndex` を実行。
2. FileDialog で PDF/Word/Excel ファイルを複数選択。
3. 抽出 → chunk → embed → 書き出しが順次走る (ステータスバーに進捗)。
4. `%LOCALAPPDATA%\InternalNotebookLM\index\` にローカル版が、`[index] remote_url` が設定されていれば SharePoint にも publish される。

## 利用集計

1. `Admin_KnowledgeBuilder.xlsm` の「利用集計」シートを開く。
2. Power Query クエリ `usage_all` をリフレッシュ (データ→クエリのリフレッシュ)。
3. SharePoint 上の全 `*.csv` が結合されてピボットテーブルに反映される。

## 月次タスク

- 新規ナレッジを取り込み → publish。
- 利用集計シートでコスト推定 (`prompt_tokens` + `completion_tokens` × 単価)。
- 上位N名の濫用チェック。
- PII警告フラグの件数を確認。
