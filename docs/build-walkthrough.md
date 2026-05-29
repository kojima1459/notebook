# Windows ビルド手順（はじめての1回）

このリポジトリには `.bas` / `.cls` / `.frm` テキストだけが入っていて、`.xlsm` は含まれていません。Windows + Excel で初回ビルドする手順を順に追ってください。**所要 30〜60 分**。

## 前提環境

- Windows 10 / 11
- Microsoft Excel 2019 以降 (デスクトップ版、Microsoft 365 でも可)
- Git for Windows
- PowerShell 5.1 以降 (デフォルトでOK)

## ステップ 1: リポジトリを取得

```powershell
cd C:\src
git clone <このリポジトリのURL> notebook
cd notebook
git checkout claude/internal-notebook-lm-chatbot-B6BE7
```

## ステップ 2: VBA-JSON ライブラリをダウンロード

`build\vendor\` を作って、VBA-JSON v2.3.1 の `JsonConverter.bas` を置きます。

```powershell
mkdir build\vendor
# ブラウザで https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.3.1 を開き、
# JsonConverter.bas を build\vendor\JsonConverter.bas に保存
```

## ステップ 3: Excel の Trust Center を一時的に開ける（ビルド時のみ）

Excel を起動 → **ファイル > オプション > トラスト センター > トラスト センターの設定 > マクロの設定** を開き、

- **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」** にチェック

これにチェックが入っていないと、`build.ps1` が xlsm に .bas を流し込めません。配布対象PCではこの設定は不要です（ビルド機のみ）。

## ステップ 4: template_chatbot.xlsm をひな型として作る

ここがコツが要るところ。「フォームのコントロール配置」は `.frm` テキストでは再現できないため、ひな型 xlsm に手作業で1回だけ frmChat を作ります。

1. Excel で空のブックを開く。
2. **名前を付けて保存 > Excel マクロ有効ブック (\*.xlsm)** で `C:\src\notebook\build\template_chatbot.xlsm` に保存。
3. **開発** タブ → **Visual Basic** で VBA エディタを開く。
4. **挿入 > ユーザーフォーム** で新規 UserForm を追加。プロパティウィンドウで `(Name)` を `frmChat` に変更。
5. フォーム上に以下のコントロールを順に配置（ツールボックスからドラッグ）:

   | コントロール種別 | (Name) | 主要プロパティ |
   |---|---|---|
   | Label | `lblConvLabel` | Caption: 会話 |
   | TextBox | `txtConversation` | MultiLine=True, Locked=True, ScrollBars=2-fmScrollBarsVertical, Height=300, Width=540 |
   | Label | `lblQuestLabel` | Caption: 質問 |
   | TextBox | `txtQuestion` | MultiLine=True, EnterKeyBehavior=False, Height=60, Width=540 |
   | CheckBox | `chkConsent` | Caption: ナレッジ改善に質問本文を提供する |
   | CommandButton | `btnSend` | Caption: 送信 |
   | Label | `lblStatus` | Caption: Ready |

   配置イメージ:
   ```
   [lblConvLabel]
   [txtConversation                                 ]
   [lblQuestLabel]
   [txtQuestion                                     ]
   [chkConsent       ]                    [btnSend]
   [lblStatus]
   ```

6. フォームを保存 (Ctrl+S)、VBA エディタを閉じる、xlsm を保存して閉じる。

## ステップ 5: template_admin.xlsm を作る

admin 側はフォーム不要なので簡単。

```powershell
# Excel で空のブックを開いて、template_admin.xlsm として保存するだけ
```

これで `build/` 配下は:
```
build/
  build.ps1
  README.md
  template_chatbot.xlsm  ← ステップ4の結果
  template_admin.xlsm    ← ステップ5の結果
  vendor/
    JsonConverter.bas    ← ステップ2の結果
```

## ステップ 6: ビルドを走らせる

```powershell
cd C:\src\notebook\build
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

成功すると `dist\Chatbot.xlsm` と `dist\Admin_KnowledgeBuilder.xlsm` が出来ます。

エラーが出た場合の見方:
- `Template is missing UserForm 'frmChat'` → ステップ 4 を忘れている。
- `インポート時にエラー` → Trust Center 設定 (ステップ 3) を確認。
- `モジュール ... が見つかりません` → `git checkout` で全ファイル取れているか確認。

## ステップ 7: config.ini を置く

```powershell
mkdir "$env:LOCALAPPDATA\InternalNotebookLM"
copy config\config.sample.ini "$env:LOCALAPPDATA\InternalNotebookLM\config.ini"
copy config\disclaimer.txt "$env:LOCALAPPDATA\InternalNotebookLM\"
```

`config.ini` を開いて、情シスから得た Azure OpenAI 接続情報を書き込みます ([docs/it-checklist.md](it-checklist.md) を参照)。

## ステップ 8: API キーを Chatbot.xlsm に埋め込む

1. `dist\Admin_KnowledgeBuilder.xlsm` を開く。
2. **開発 > マクロ** で `modKeyEnroller.EnrollKeyInteractive` を実行。
3. InputBox に Azure OpenAI のAPIキーを貼り付け、ファイル選択で `dist\Chatbot.xlsm` を指定。
4. 「埋め込みました」のメッセージが出れば成功。

## ステップ 9: 自PCで end-to-end テスト

1. ナレッジを作る:
   - `dist\Admin_KnowledgeBuilder.xlsm` を開いて、`modKnowledgeBuilder.BuildIndex` を実行。
   - 適当な `.txt` や `.pdf` を選択。「N chunks をindex化しました」が出れば成功。
   - `%LOCALAPPDATA%\InternalNotebookLM\index\` に manifest.json + embeddings.bin + chunks.json が出来ているはず。
2. チャット:
   - `dist\Chatbot.xlsm` を開く。初回は部署/役割を入力するダイアログが出る。
   - チャット画面が表示されたら何か質問を入れて「送信」。
   - 数秒で回答 + 出典が出れば成功。
3. ログ確認:
   - `%LOCALAPPDATA%\InternalNotebookLM\usage_buffer.csv` に1行追記されているか確認。

## ステップ 10: パイロット配布の前に

- `dist\Chatbot.xlsm` の VBA プロジェクトに **パスワード保護** をかける:
  VBAエディタ → ツール → VBAProject のプロパティ → 保護タブ → 「プロジェクトを表示用にロックする」+ パスワード。
- 信頼できる場所の登録 (情シスのGPO配信が理想)。
- `config.ini` をパイロット参加者のPCに配置 (キーは含まれない、エンドポイントとSharePoint URLのみ)。
- 配布フォーマットの最終確認 ([docs/deployment.md](deployment.md))。

## トラブルシュート

| 症状 | 対処 |
|---|---|
| `build.ps1`: `プログラムによる Visual Basic プロジェクトへのアクセスは信頼性に欠ける...` | ステップ 3 の Trust Center 設定。 |
| `frmChat` のコントロールが空 | ステップ 4 でコントロールの (Name) を正確に付けたか確認。 |
| `JsonConverter stub - run build.ps1...` がチャット時に出る | ステップ 2 で vendor の JsonConverter.bas を置いていない、もしくは古い。 |
| `embed失敗: HTTP 401` | config.ini の `endpoint` / `deployment` 名、もしくは ステップ 8 のキー埋め込みを確認。 |
| `HTTP 407` プロキシ認証 | `[proxy] mode = manual` + `proxy_url` を設定するか、情シスに NTLM 透過を依頼。 |
