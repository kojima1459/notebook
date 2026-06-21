# ⚽ サンガキッズ — 京都サンガ＆サッカー 応援サイト

京都サンガと世界のサッカーが好きな子ども（おもに小学校高学年）向けの、モバイル対応の応援サイトです。
**本物のニュース・順位表・試合結果**を自動で取得し、技の練習メニューまで揃えた「サッカー専用ゾーン」。

🌐 公開URL: **https://notebook-phi-seven.vercel.app**

> 💡 **ねらい**：YouTubeでサッカーを見ていて、いつの間にか関係ない“消費系”動画に流れてしまう——
> その脱線だけを断つのが目的。サッカーは制限せず、動画・ニュース・順位・練習をこの中で完結させ、
> **見る → やってみる**へつなげます。

## できること

- 📰 **最新ニュース（自動）** — 京都サンガ・Jリーグ・海外サッカー・海外組（日本人選手）・速報。試合結果もここで（`news.html`）
- 📊 **順位表（自動・本物）** — J1全チームの最新順位をJリーグ公式から取得。サンガの順位・勝点・得失点をハイライト（`standings.html`）
- 🎬 **サッカー動画（自動）** — サッカー専門チャンネルの最新動画を自動再生。⭐お気に入り登録つき（`shorts.html`）
- 🔥 **練習＆技** — スター選手の技、毎日チャレンジ＆バッジ、レベル別プログラム、すきま時間トレ、強くなる栄養（`training.html`）
- 🎮 **サンガ＆世界サッカークイズ** — 全12問（`quiz.html`）

## アーキテクチャ（重要）

純粋な静的サイトでは外部データをリアルタイム取得できない（CORS）ため、**Vercel サーバーレス関数**で
サーバー側からデータを取得し、CDNキャッシュして返します。

| エンドポイント | 内容 | データ源 |
|---|---|---|
| `/api/standings` | J1順位表（EAST/WESTグループ） | jleague.jp 公式をサーバー側で取得・解析 |
| `/api/news?cat=...` | カテゴリ別ニュース | Google ニュース RSS をサーバー側で取得 |
| `/api/sanga` | サンガの次の試合・直近の結果・プロフィール | サンガ公式ニュースのタイトル解析 |
| `/api/roster` | 選手名鑑（主な選手） | TheSportsDB |
| `/api/chat` | サッカー博士チャットボット | Gemini（`GEMINI_API_KEY`）＋ Google検索グラウンディング |

`cat` = `sanga` / `jleague` / `official` / `overseas` / `japan` / `flash`。

### 💬 チャットボット（サッカー博士）

全ページ右下のウィジェット。質問すると、いま画面に出ている内容（ニュース見出し・順位表・試合情報）を
**テキストで自動収集**して `/api/chat` に送り、Gemini が回答します（画面共有は不要）。
最新の試合結果や得点者などは Google 検索グラウンディングで調べて答えます。

**セットアップ**：Vercel のプロジェクト環境変数に `GEMINI_API_KEY`（[Google AI Studio](https://aistudio.google.com/apikey) のキー）を設定してください。
キーはサーバー側だけで使われ、ブラウザには出ません。モデルは `gemini-flash-latest`。
レスポンスは `Cache-Control: s-maxage` でCDNキャッシュし、上流アクセスを最小化しています。

## 🔒 セキュリティ（一般公開むけ）

このサイトはインターネットに公開しても安全なように、次の対策をしています。

**APIキーの守り方（多層防御）**
- キーは **Vercelの環境変数 `GEMINI_API_KEY`** にだけ保存。コードにも、ブラウザにも、GitHubにも出ません（暗号化保存）。
- フロント（ブラウザ）からはキーを一切さわらず、`/api/chat`・`/api/quiz` の **サーバー関数の中だけ** で使用。
- リポジトリ・全コミット履歴を検査済み（キーのハードコードなし）。

**APIの濫用（課金あらし）対策** — `api/_guard.js`
- **レート制限**：同一IPから 1分あたり 15回まで（超えたら `429`）。
- **Originチェック**：このサイト（`*.vercel.app`）以外からの `/api/chat`・`/api/quiz` 呼び出しは `403` で拒否。
- **エラーの中身を隠す**：Geminiのエラー詳細はブラウザに返さず、サーバーログにだけ記録。
- 全APIに `X-Content-Type-Options` / `X-Frame-Options` / `Referrer-Policy` を付与。
- 外部取得系API（news/sanga/standings/roster）は **固定URLのみ取得**（ユーザー入力でURLを変えられない＝オープンプロキシにならない）。

**🛡️ おうちの人にやってほしい最重要設定（5分）**

万一キーが漏れても被害ゼロにするため、Google側でキーを「制限」してください：

1. [Google AI Studio / Google Cloud Console](https://console.cloud.google.com/apis/credentials) を開く
2. `GEMINI_API_KEY` を選ぶ → **「APIの制限」** で **Generative Language API だけ** に限定
3. [Google AI Studio](https://aistudio.google.com/) の **Billing/使用量上限** で、月の予算アラートや上限を設定
4. もしキーが漏れたかもと思ったら、すぐ **キーを削除して再発行** → Vercelの環境変数を入れ替えて再デプロイ

> これで「サーバー側だけで使う＋濫用制限＋Google側で用途と上限を固定」の三段構えになります。

## ファイル構成

```
.
├── index.html / news.html / standings.html / shorts.html / training.html / quiz.html
├── vercel.json              # cleanUrls 等
├── api/
│   ├── standings.js         # J1順位表（jleague.jp 解析）
│   └── news.js              # ニュース（GoogleニュースRSS）
├── data/channels.json       # 動画ゾーンのサッカーチャンネル設定
└── assets/
    ├── css/style.css
    └── js/
        ├── main.js          # ナビ・ホームのライブニュース
        ├── standings.js     # 順位表表示
        ├── news-feed.js     # ニュース表示（/api/news）
        ├── shorts.js        # 動画ゾーン＋お気に入り
        ├── training.js      # 技・チャレンジ・バッジ・プログラム
        └── quiz.js          # クイズ
```

## デプロイ

```bash
vercel deploy --prod
```

`api/*.js` は Vercel が自動的にサーバーレス関数として認識します（Node 18+ の global `fetch` を使用）。

## ご注意

ファンによる非公式サイトです。ニュース・順位表は Jリーグ公式 / Google ニュース等の外部コンテンツへリンク・取得しています。
公式の最新情報は **京都サンガF.C. / Jリーグ 公式サイト** をご確認ください。
2026シーズンは移行期の「百年構想リーグ」（EAST/WESTグループ制）です。
