# 【管理者用】社内NotebookLMチャットbot 導入ガイド

このガイドは **あなた自身（管理者・パッケージを配る人）** がWindowsの社内PCで初期設定・動作確認するためのものです。
営業担当者に配るときの手順は、別ファイル `02_distributor-package-guide.md` を参照してください。

所要時間：**約30〜45分**（一度やれば終わり）

---

## ステップ 0：用意するもの

| 必要なもの | どこから | 備考 |
|---|---|---|
| 社内 Windows PC | あなた自身のもの | デスクトップ版 Excel 2019 以降が入っていること |
| インターネット接続 | 社内LAN | Gemini API（`generativelanguage.googleapis.com`）に到達できる必要あり |
| Gemini APIキー | Google AI Studio | 既に発行済みのもの。漏らさないよう注意 |
| 配布パッケージ一式 | GitHub から ZIP ダウンロード | 後述 |

---

## ステップ 1：GitHubからファイル一式を落とす

> ⚠️ **注意**：GitHub のトップページから `Code → Download ZIP` を押すと **README.md だけ** しか入っていません（作業用ブランチがmainではないため）。**必ず下記の専用URLから落としてください。**

1. ブラウザ（Edge でOK）のアドレスバーに、以下を **コピペして Enter**：

   ```
   https://github.com/kojima1459/notebook/archive/refs/heads/claude/prototype-gemini-demo.zip
   ```

   すぐに ZIP のダウンロードが始まります。

2. ダウンロードフォルダに `notebook-claude-prototype-gemini-demo.zip` ができます。これをデスクトップにコピー → **右クリック → すべて展開** → 展開先は **デスクトップ** のままでOK。

3. 展開後、デスクトップに `notebook-claude-prototype-gemini-demo` フォルダが出来上がります。

   この中には：
   - `dist\Chatbot.xlsm` ← 営業担当者に配るやつ
   - `dist\Admin_KnowledgeBuilder.xlsm` ← あなた専用（PDF追加用）
   - `dist\index\` ← AI検索用データ（manifest.json / embeddings.bin / chunks.json）
   - `config\config.sample.ini` ← 設定ファイルのひな型
   - `config\disclaimer.txt` ← 免責事項
   - `demo_data\sample_01.pdf` 〜 `sample_16.pdf` ← デモ用のサンプルPDF

   もし README.md しか無い場合は、上記URLではなくGitHubトップから落としている可能性が高いので、URLを再確認してください。

---

## ステップ 2：あなたのPCに「設定フォルダ」を作る

1. キーボードの **Windowsキー + R** を同時押し → 「ファイル名を指定して実行」が開く。

2. そこに以下を貼り付け（コピペ）→ Enter：

   ```
   %LOCALAPPDATA%
   ```

3. エクスプローラーが `C:\Users\あなたのユーザー名\AppData\Local` を開きます。

4. このフォルダの **何もないところを右クリック → 新規作成 → フォルダー** → 名前を **`InternalNotebookLM`** にしてEnter。

5. これで `C:\Users\あなたのユーザー名\AppData\Local\InternalNotebookLM\` というフォルダが出来ました。これが **本ツールの設定置き場** です。

---

## ステップ 3：設定ファイルをコピー＆編集

1. デスクトップの `notebook-claude-prototype-gemini-demo` フォルダを開く。

2. `config\config.sample.ini` を **右クリック → コピー**。

3. ステップ2で作った `InternalNotebookLM` フォルダを開いて **右クリック → 貼り付け**。

4. 貼り付けたファイルを **右クリック → 名前の変更** で、ファイル名を `config.ini` に変更（`.sample` を消すだけ）。

5. `config.ini` を **右クリック → プログラムから開く → メモ帳** で開く。

6. 中の以下の行を探す：

   ```ini
   api_key =
   ```

7. イコールの右側に **半角スペース1個** をはさんで、Gemini APIキーを貼り付け：

   ```ini
   api_key = AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
   ```

8. **Ctrl + S** で保存して、メモ帳を閉じる。

> ⚠️ APIキーは誰にも見せない。USBやメールで送らない。今後ローテーション（差し替え）するときは、このファイルを直接書き換える。

---

## ステップ 4：免責事項ファイルとindexフォルダもコピー

同じ要領で、以下も `InternalNotebookLM` フォルダに入れます：

1. `notebook-claude-prototype-gemini-demo\config\disclaimer.txt` を `InternalNotebookLM` にコピー。

2. `notebook-claude-prototype-gemini-demo\dist\index\` **フォルダ丸ごと** を `InternalNotebookLM` にコピー。

完了後、`InternalNotebookLM` フォルダの中身は以下になります：

```
InternalNotebookLM\
  ├ config.ini             (APIキーを書いたやつ)
  ├ disclaimer.txt
  └ index\
      ├ manifest.json
      ├ embeddings.bin     (19MBくらい)
      └ chunks.json        (1.5MBくらい)
```

---

## ステップ 5：Excelのマクロ設定を確認（初回1回だけ）

1. Excel を起動。

2. **ファイル → オプション → トラスト センター → トラスト センターの設定 → マクロの設定**。

3. ラジオボタンが以下のどれかになっていればOK：
   - 「通知を有効にしてVBAマクロを無効にする」（推奨）
   - 「デジタル署名されたマクロを除き、すべてのVBAマクロを無効にする」

4. **OK → OK** で閉じる。

> 「すべてのマクロを有効にする」にする必要はありません。各ファイルを開くときに「コンテンツの有効化」を押せば動きます。

---

## ステップ 6：Chatbot.xlsm を開いて動作確認

1. `notebook-claude-prototype-gemini-demo\dist\Chatbot.xlsm` を **ダブルクリック**。

2. 上に黄色いバーで **「セキュリティの警告 マクロが無効にされました [コンテンツの有効化]」** が出る → **コンテンツの有効化** をクリック。

3. 数秒待つと、初回のみ **「部署名を入力してください」** ダイアログが出る。
   - 部署名：例 `引受第1部` と入力 → OK
   - 役割：例 `アンダーライター` と入力 → OK

4. シート「ChatUI」が自動で表示される。
   - 青いタイトルバー「社内ナレッジQAボット (デモ版)」
   - 黄色い質問入力欄
   - 「> 送信」ボタン
   - 下に免責事項

5. 質問欄に試しに入力：
   ```
   告知義務違反した場合の対応を教えてください
   ```

6. **「> 送信」** をクリック。

7. **30〜60秒** お待ちください（Gemini 2.5 Pro + 自己検証で時間がかかる）。

8. AI回答 + 出典 + 状態（"Ready (xxx/yyy tok, zzzz ms)"）が下に表示されれば成功。

---

## ステップ 7：問題が出たとき

| 症状 | 原因と対処 |
|---|---|
| 「ファイルが見つかりません」と出る | `InternalNotebookLM` フォルダ（ステップ2〜4）の場所が違う。`%LOCALAPPDATA%` の直下にあることを確認 |
| 「APIキーが無効です」「HTTP 401」 | config.ini の `api_key =` の右に正しいキーが貼られているか確認。`api_key=ABCD` のように前後にスペースが無いか確認 |
| 「HTTP 404」「model not found」 | config.ini の `chat_deployment = gemini-2.5-pro` と `embed_deployment = gemini-embedding-001` を確認 |
| 「HTTP 429」「rate limit」 | APIキーの利用上限。Google AI Studio で上限を上げるか、しばらく待つ |
| 30秒以上反応がない | 社内プロキシがブロックしている可能性。Edgeで `https://generativelanguage.googleapis.com/` にアクセスできるか確認 |
| マクロが動かない・ボタンが反応しない | 一度 Chatbot.xlsm を閉じる → 開き直して「コンテンツの有効化」を確実に押す |

エラーメッセージのスクリーンショットを撮っておくと、僕（Claude）に投げるときに早いです。

---

## ステップ 8：ナレッジを追加・削除したいとき

**現状の仕組み**：indexはあらかじめベクトル化したものを配るので、PDFを増減するときは **index を作り直して配布し直す** 必要があります。

### 楽な方法（推奨）

新しいPDF（または削除したいPDF名）を僕に渡してください：

1. 追加したいPDFを共有 → 僕が `demo_data\` に追加してindex再生成
2. 新しい `dist\index\` フォルダを受け取る
3. これを配布パッケージに上書きして配り直す（または各営業PCの `InternalNotebookLM\index\` を更新するだけ）

### 自分でやりたい場合（社内PCから Gemini APIが通れば可能）

`Admin_KnowledgeBuilder.xlsm` から **BuildIndex** マクロを呼び出すと、PDFを選択ダイアログで指定して再ベクトル化できます（ただし、ベクトル化はネットワーク次第で15〜30分かかります）。

---

## おまけ：ベクトル化の手作業手順（Pythonが入っているPCのみ）

社内PCではほぼ無理ですが、参考までに：

1. PowerShell で `pip install pypdf requests`
2. `set GEMINI_API_KEY=...`
3. `python build\build_index.py`

これで `dist\index\` に再生成されます。

---

## 次のステップ

動作確認ができたら、 **`02_distributor-package-guide.md`** を見て、営業担当者に配る USB/共有フォルダ用のパッケージを作ってください。
