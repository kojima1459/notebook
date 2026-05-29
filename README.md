# Internal NotebookLM (Excel + VBA / Azure OpenAI)

社内アンダーライターの問い合わせ対応を減らすため、支社営業担当者向けに配布する社内ナレッジQAチャットbot。Excel マクロ (.xlsm) で動作し、社内ナレッジ (PDF/Word/Excel) を出典付きで検索回答する。

## このリポジトリの構成

```
src/
  shared/      両xlsmで共有するモジュール (HTTP, ApiGateway, KeyVault, Config, ...)
  admin/       Admin_KnowledgeBuilder.xlsm 専用 (取り込み・index生成・キー埋め込み)
  chatbot/     Chatbot.xlsm 専用 (起動・チャットUI・利用ログ)
build/         build.ps1 と Excel テンプレート (Windowsで実行)
config/        config.sample.ini, disclaimer.txt
docs/          architecture.md, security.md, admin-guide.md, user-guide.md
dist/          ビルド成果物 (.gitignore)
```

## 設計の柱

- **2バイナリ構成**: 管理者用と配布用を分離。配布用は読み取り+利用のみ。
- **APIキー秘匿の3モード抽象化** (`modApiGateway`): Mode A 中継サーバー / **Mode B 難読化埋め込み (MVPデフォルト)** / Mode C Entra ID。
- **Azure OpenAI Service 前提**。エンドポイントとデプロイメント名は config.ini で外出し。
- **L2正規化済み embeddings をバイナリ保存** + 純VBA内積でtop-k。
- **PII検知 + レート制限 + 利用ログ** をクライアント側で持つ。
- **VBA保護の限界**を `docs/security.md` に明記。

## 実装ロードマップ

- Week 1 (この時点): 縦串通し。txt取り込み→Azure OpenAI→チャット応答。
- Week 2: PDF/Word/Excel抽出, プロキシ対応, SharePoint連携, ログ, レート制限, PII。
- Week 3: パイロット配布 (5〜10名), 集計シート, 個人ナレッジ。
- Phase 2: ModeA中継サーバー, ModeC Entra ID, 検索精度チューニング。

## 開発フロー

ソースは `src/` に .bas/.cls/.frm のテキストで管理する (xlsm はバイナリで diff 不能なため)。Windows + Excel 環境で `build/build.ps1` を走らせると `dist/*.xlsm` が生成される。

## ドキュメント

- [docs/it-checklist.md](docs/it-checklist.md) - **着手前に情シスへ持っていく1枚**
- [docs/build-walkthrough.md](docs/build-walkthrough.md) - **初回 Windows ビルド手順 (所要30〜60分)**
- [docs/architecture.md](docs/architecture.md) - 全体構成
- [docs/security.md](docs/security.md) - 秘匿の限界と移行計画
- [docs/admin-guide.md](docs/admin-guide.md) - 管理者手順
- [docs/user-guide.md](docs/user-guide.md) - 営業担当者向け
- [docs/deployment.md](docs/deployment.md) - 配布と信頼できる場所
