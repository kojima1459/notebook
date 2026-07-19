# Windows実機テスト環境の作り方(あなたが手でテストしなくて済むようにする)

> **この文書は誰向けか**: 「Nexus Agent」の実機テストを、自分の手ではなく
> **Windows PC上のClaude Codeに自動でやらせたい**開発オーナー向け。

## 仕組み(3行)

1. 安いWindows PCにデスクトップ版Excelと[Claude Code](https://claude.com/claude-code)を入れる
2. そのPC上のClaudeがPowerShell経由でExcelを直接操縦(COM自動化)し、ビルド→自己インストール→マクロ実行→結果検証を全自動で回す
3. 本物の「リボンちゃん」が無い分は、確定シグネチャと同一のスタブ **「ニセリボンちゃん」**(このフォルダに同梱)で配管を全部検証する

これで**会社PCに残る確認は「本物のAIの回答品質」だけ**になる(docs/40 §13の実質1項目)。

## 買い物リスト

| もの | 目安 | 備考 |
|---|---|---|
| Windows PC | 中古2〜3万円で十分 | Windows 10/11 64bit・メモリ8GB以上。会社PCと同じWindows 11推奨 |
| Excel | Microsoft 365 Personal(月額) または Office Home & Business(買い切り) | **デスクトップ版必須**。Web版・モバイル版はVBAが動かないので不可 |

MacBook Airは不可(Mac版ExcelはVBAの互換性が別物で、本ツールはWindows専用設計)。

## 初期セットアップ(1回だけ・30分)

1. ExcelをインストールしてMicrosoftアカウントでライセンス認証する
2. [Git for Windows](https://gitforwindows.org/) と [Python 3](https://python.org)(ビルド用)を入れる
3. Claude Codeを入れる(PowerShellで `irm https://claude.ai/install.ps1 | iex` またはデスクトップアプリ)
4. このリポジトリをcloneする
5. **ニセリボンちゃんを作る**(5分・1回だけ):
   1. Excelで空のブックを開く → `Alt+F11` でVBEを開く
   2. メニュー「ファイル→ファイルのインポート」で `wintest/mock_ribbon/modMockRibbon.bas` を取り込む
   3. `F12`(名前を付けて保存)→ ファイルの種類「**Excelアドイン(*.xlam)**」→
      ファイル名 **`リボンちゃん(検証用).xlam`** で保存(※「リボンちゃん」という文字を必ず含めること。
      対象ツール側のアドイン検出が名前の部分一致のため)
   4. 保存先は既定のAddInsフォルダ(`%APPDATA%\Microsoft\AddIns`)のままでよい

## 実行(あなたは見てるだけ)

Claude Codeにこう頼む:

> wintest/README.md に従って実機テストを回して。失敗したら原因を調べて直して、全部緑になるまで繰り返して。

Claudeが実行する内容(手動でも実行可):

```powershell
# 1) ビルド(開発版=リボン無しでも全画面が動くmockモード)
python build\build_mybookshelf.py --dev

# 2) 実機テスト(開発版: 自己インストール+純ロジックテスト84本を実機VBAで実行)
powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1

# 3) 本番版+ニセリボンちゃんで、リボン依存の配管まで検証
python build\build_mybookshelf.py --prod
powershell -ExecutionPolicy Bypass -File wintest\run_excel_tests.ps1 -Target prod `
  -MockAddinPath "$env:APPDATA\Microsoft\AddIns\リボンちゃん(検証用).xlam"
```

3)のあとClaudeがCOMでExcelを操作し、docs/40 §13相当の動作(続けて質問/Wordで開く/
スクショ取込)をボタンのOnActionマクロ直接実行で確認できる。

## これでも検証できないもの(会社PCが必要)

- **本物のAIの回答品質**(ニセリボンは決め打ち応答)
- 本物のリボンちゃんの実挙動の最終確認(effort/verbosity引数・LimitCheckの真偽など、
  台帳 `docs/dev/RIBBON_API_CONFIRMED.md` §3 の未確定項目)

→ 会社PCでは docs/40_受入チェックリスト15分.md の §13 だけ実施すればよい。

## セキュリティ上の注意

- ニセリボンちゃんはAPIキーもエンドポイントも一切含まない(完全スタンドアロン)
- **会社PCにニセリボンちゃんを入れないこと**(本物と名前が部分一致するため誤検出の元)
- 個人PCに会社の実データ(顧客情報を含む資料)を持ち込まない。テストはダミー資料で行う
