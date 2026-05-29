# プロトタイプ ビルド手順 (あなた専用)

このページは「情シスに渡す xlsm を、あなたの Windows PC で1回だけ組み立てる」ための手順です。本番版の [build-walkthrough.md](build-walkthrough.md) と内容は似ていますが、デモ専用の差分 (Gemini設定 / サンプルPDFのindex化) を盛り込んだ短縮版です。

所要 **30〜45分**。

## ステップ 1: ベースを揃える

[docs/build-walkthrough.md](build-walkthrough.md) のステップ 1〜6 を実施。終わると `dist\Chatbot.xlsm` と `dist\Admin_KnowledgeBuilder.xlsm` ができている状態。

> 本番版とコードは共通で、`config.ini` だけ Gemini 用に差し替えれば動きます。Gemini対応のコードはどちらの xlsm にも入っています。

## ステップ 2: APIキーを config.ini に入れる

```powershell
mkdir "$env:LOCALAPPDATA\InternalNotebookLM" -Force
copy config\config.sample.ini "$env:LOCALAPPDATA\InternalNotebookLM\config.ini"
copy config\disclaimer.txt    "$env:LOCALAPPDATA\InternalNotebookLM\"
notepad "$env:LOCALAPPDATA\InternalNotebookLM\config.ini"
```

`api_key =` の行に Claude から伝えられた Gemini キーを貼り付け、保存。

## ステップ 3: サンプルPDFから index を作る

リポジトリの `demo_data\` 配下に サンプルPDF 16ファイルが入っています。

1. `dist\Admin_KnowledgeBuilder.xlsm` を開く
2. 「開発」タブ → 「マクロ」 → `BuildIndex` を選択して実行
3. ファイル選択ダイアログで `demo_data\sample_01.pdf` 〜 `sample_16.pdf` を **全選択**
4. ステータスバーに「embed中: N / 全件」が表示されながら進行 (16ファイル ~15MB で **3〜10分** かかります - Gemini APIの batch embed を batch=16 で叩くため)
5. 完了すると「N chunks をindex化しました」が出る
6. `%LOCALAPPDATA%\InternalNotebookLM\index\` に `manifest.json`, `embeddings.bin`, `chunks.json` ができていることを確認

## ステップ 4: 自PCで動作確認

1. `dist\Chatbot.xlsm` をダブルクリック
2. 初回は部署・役割の入力ダイアログが2回出る (適当でOK)
3. チャット画面が出る
4. サンプルPDFに関する質問を入れて「送信」
5. 数秒〜10秒で 回答 + 出典 が表示されれば成功

ここで問題なく動けば、配布パッケージの中身は正しいです。

## ステップ 5: 情シスに渡すパッケージを作る

USB なり共有フォルダなりで以下を渡します。

```
demo_package\
├── Chatbot.xlsm                    (dist\ からコピー)
├── Admin_KnowledgeBuilder.xlsm     (dist\ からコピー - 任意)
├── config.ini                      (apikey入り)
├── disclaimer.txt
├── index\                          (ステップ3の結果丸ごと)
│   ├── manifest.json
│   ├── embeddings.bin
│   └── chunks.json
└── README.txt                      ← demo-quickstart.md をテキスト化して同梱
```

`README.txt` のかわりに [docs/demo-quickstart.md](demo-quickstart.md) を PDF や DOCX に変換して同梱するのが分かりやすいです。

## 渡した後のお作法

- **デモ終了後はGemini API キーを必ずローテート** (Google AI Studio で「キーを削除」→ 新規発行)
- 情シスにキーを抜かれて悪用される、というシナリオではなく、xlsm が他人の手に渡ったら漏洩前提で動く、という運用感覚を体感してもらうため
- これは demo-quickstart.md にも書いてあるので情シスにも共有
