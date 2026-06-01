# プロトタイプ ビルド手順 (あなた専用)

このページは「情シスに渡す xlsm を、あなたの Windows PC で1回だけ組み立てる」ための手順です。

**設計を作り直したので、フォームを手で配置する作業は不要になりました**。質問入力欄と送信ボタンを含むチャットUIは、xlsmを開いたときに自動でシート上に生成されます。

所要 **15〜30分**。

## 前提

- Windows 10 / 11
- Microsoft Excel 2019 以降 (デスクトップ版)
- インターネット接続

## ステップ 1: コードを自分のPCに持ってくる

GitHub のページを開いて、緑色の **`Code` ボタン → `Download ZIP`** を押すだけ。
zip を解凍して、たとえば `C:\notebook\` に置きます。

> Git をインストールしてある場合は `git clone -b claude/prototype-gemini-demo <URL>` でもOKですが、上記の ZIP の方が手間なし。

## ステップ 2: (不要・スキップ)

以前は VBA-JSON ライブラリを別途ダウンロードしてもらう必要がありましたが、リポジトリに同梱したのでこの手順は**不要になりました**。ZIPを解凍した時点で `build\vendor\JsonConverter.bas` に既に入っています。

## ステップ 3: Excel の設定を1つだけ変える

Excel を起動 → **ファイル > オプション > トラスト センター > トラスト センターの設定 > マクロの設定**

- **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」** にチェック
- OK を押して閉じる

これは「ビルドのときだけ」必要で、配布対象PCでは不要です。

## ステップ 4: ビルドを走らせる

PowerShell を開いて (スタート → `powershell` と入力):

```powershell
cd C:\notebook\build
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

これだけ。Excel が裏で動き、`C:\notebook\dist\` の中に **`Chatbot.xlsm`** と **`Admin_KnowledgeBuilder.xlsm`** が出来ます。

> エラーが出たら、メッセージごと僕に貼り付けてください。よくあるのはステップ 3 を忘れているか、ステップ 2 で別の場所にダウンロードしたパターン。

## ステップ 5: APIキーを設定する

```powershell
mkdir "$env:LOCALAPPDATA\InternalNotebookLM" -Force
copy C:\notebook\config\config.sample.ini "$env:LOCALAPPDATA\InternalNotebookLM\config.ini"
copy C:\notebook\config\disclaimer.txt    "$env:LOCALAPPDATA\InternalNotebookLM\"
notepad "$env:LOCALAPPDATA\InternalNotebookLM\config.ini"
```

メモ帳が開いたら、`api_key =` の行に Claude から伝えられたキーを貼り付け、保存。

## ステップ 6: サンプルPDFから index を作る

1. `C:\notebook\dist\Admin_KnowledgeBuilder.xlsm` をダブルクリックで開く
2. マクロ有効化のメッセージが出たら「有効にする」
3. **開発** タブ → **マクロ** → リストから `BuildIndex` を選択して **実行**
4. ファイル選択ダイアログが出るので `C:\notebook\demo_data\` に移動して、`sample_01.pdf` 〜 `sample_16.pdf` を **全選択** (Ctrl+A) → 開く
5. 画面下のステータスバーに「embed中: N / 全件」が出ながら進む (16ファイルで **3〜10分**)
6. 「N chunks をindex化しました」が出れば成功

## ステップ 7: 自PCで動作確認

1. `C:\notebook\dist\Chatbot.xlsm` をダブルクリックで開く
2. 初回は **部署と役割** を聞かれるので適当に入れる (例: 「営業1部」「営業担当」)
3. 「ChatUI」シートが自動で描かれて、青いタイトルバーと質問入力欄が出てくる
4. 質問欄に何か入れて **▶ 送信** をクリック
5. 数秒〜10秒で AI回答 + 出典が下に追加されれば成功

問題なく動けば、配布パッケージの中身は正しいです。

## ステップ 8: 情シスに渡すパッケージを作る

`C:\notebook\demo_package\` を作って、以下をコピー:

| コピー元 | 配布先 |
|---|---|
| `C:\notebook\dist\Chatbot.xlsm` | `demo_package\Chatbot.xlsm` |
| `C:\notebook\dist\Admin_KnowledgeBuilder.xlsm` | `demo_package\Admin_KnowledgeBuilder.xlsm` |
| `%LOCALAPPDATA%\InternalNotebookLM\config.ini` | `demo_package\config.ini` |
| `C:\notebook\config\disclaimer.txt` | `demo_package\disclaimer.txt` |
| `%LOCALAPPDATA%\InternalNotebookLM\index\` フォルダ | `demo_package\index\` |
| `C:\notebook\docs\demo-quickstart.md` | `demo_package\README.md` |

これを USB か社内共有フォルダで情シスに渡す。情シス側の手順は `demo-quickstart.md` の通り (約5分で起動)。

## 渡した後

- **デモが終わったら Gemini API キーを必ずローテート**
  - Google AI Studio (`https://aistudio.google.com/apikey`) で旧キーを削除して新規発行
  - 「xlsm が他人の手に渡る = キーが漏れる」前提なので、これは絶対やってください

## トラブル時の連絡

何か詰まったら、エラーメッセージのスクリーンショットと「どのステップでどう詰まったか」を僕に教えてください。
