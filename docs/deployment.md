# 配布とデプロイ

## 配布先パス

`%LOCALAPPDATA%\InternalNotebookLM\` 配下:

```
config.ini
disclaimer.txt
index\               (Workbook_Open でSharePointから同期)
  manifest.json
  embeddings.bin
  chunks.json
profile.dat          (初回設定で生成)
ratelimit.dat        (起動中に書かれる)
usage_buffer.csv     (リモート公開前のローカルバッファ)
```

xlsm 本体は別途置く (例: `C:\Program Files\InternalNotebookLM\Chatbot.xlsm` を IT管理として配置、デスクトップにショートカット)。

## Trust Center 設定

ユーザーPCで以下のどれかが必要:
1. **信頼できる場所** に xlsm の格納先を登録 (IT部門GPOで配布が望ましい)。
2. xlsm に **組織CAの署名** を施し、Trusted Publisher 経由で実行。
3. (非推奨) Trust Center で全マクロを有効化。

## 配布手順 (パイロット)

1. `dist\Chatbot.xlsm` をビルド (APIキー埋め込み済み)。
2. 配布対象PCに以下を配置:
   - `Chatbot.xlsm` → IT指定の信頼できる場所
   - `config.ini` → `%LOCALAPPDATA%\InternalNotebookLM\config.ini`
   - `disclaimer.txt` → 同上ディレクトリ
3. SharePoint のナレッジindex/利用ログフォルダのアクセス権を付与。
4. デスクトップショートカットを作る (任意)。
5. 初回起動: 部署/役割の入力ダイアログを通過させる。

## ロールバック

- 旧版を残してパスを差し替えるだけ。indexは後方互換 (`manifest.dim` で検証)。
- 重大な不具合時は config.ini で `[index] remote_url` を空にすればローカルキャッシュのみで動く。

## モニタリング

- 月1で利用集計シートを更新 → コスト・濫用・PII検知をチェック。
- アラート閾値:
  - コスト: 月予算の80%到達
  - 1ユーザー1日 > 200質問
  - PII警告 / 日 > 10件
