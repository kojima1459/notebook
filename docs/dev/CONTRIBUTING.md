# Nexus Agent — 開発フロー・PRの出し方

このリポジトリへ変更を入れる同僚(高橋さん等)向けのガイド。
「何を読めばよいか」は `docs/dev/ARCHITECTURE.md` を、「何が実装済みか」は
`docs/dev/FEATURES.md` を、「何がまだか」は `docs/dev/TODO.md` を参照。
本書は**変更を安全に入れて、PRとして送る手順**に特化する。

## 1. 開発環境の準備

### 1.1 必須ツール
- Python 3.9+ (`build/`・`tools/`のスクリプトが利用)
- LibreOffice(構文コンパイル検証専用。**実行はしない**。インストール手順は
  `tools/README.md`参照)
- Windows版 Excel + 社内AIリボン「リボンちゃん」(実機受入テストのみ必須。
  日常のコード編集・静的検証はMac/Linux上でも完結する)

### 1.2 クローン後の初回確認
```bash
python3 tools/vba_lint.py                 # 0 ERROR になることを確認
python3 tools/run_lo_tests.py             # 全モジュールPASSになることを確認
python3 build/build_mybookshelf.py --dev  # devビルドが自己検証OKで通ることを確認
```
この3コマンドが全てグリーンの状態が「正常なスタート地点」。ここでエラーが
出る場合は環境構築(LibreOfficeのインストール等)を先に疑うこと。

## 2. 変更を入れる基本ルール

### 2.1 レイヤリング(R1)を破らない
`src/core/` ← `src/ingest,qa,pack,stats/` ← `src/ui/` の一方向のみ。
下位層から`ui/`を参照すると`vba_lint.py`が「R1違反」で止める。UI層の値を
下位層のロジックへ渡したい場合は、逆に**呼び出し元(ui層)が下位層の
Public関数を呼ぶ**形に設計すること(実例: `modBitwiseOpt`のログ文字列を
`modApp`が取り出してToast表示する。qa層からUI層へは呼ばない)。

### 2.2 closed契約モジュールへのPublic追加
`modStats`/`modAsk`/`modAppDef`/`modConfig`/`modUtil`等、
`tools/vba_lint.py`の`CONTRACT`辞書に`"closed": True`と書かれている
モジュールは、**Public メンバーの追加に契約更新が必須**。手順:

1. モジュールへPublic関数/Subを追加
2. `tools/vba_lint.py`の`CONTRACT`辞書、該当モジュールの`"required"`配列に
   関数名を追加
3. `python3 tools/vba_lint.py`を再実行し、「契約違反」エラーが消えることを
   確認

これを忘れると lint が「§7に無いPublicは作らない規約」違反として止める
(意図的な安全装置。既存の実例は`modAsk.LastHit*`アクセサ追加時の
コミット参照)。

### 2.3 新しい`.bas`/`.cls`モジュールを追加したら
`build/modules.json`に登録すること。登録し忘れると、ビルド時点で
モジュール集合の照合が落ち「見つからない」エラーになる
(R35で自動検出が強化されたため、手術漏れは即座に見える)。

```json
{
  "name": "modXxx",
  "path": "src/ui/modXxx.bas",
  "role": "core",
  "type": "std",
  "wave": "NEXUS-P12"
}
```

### 2.4 新機能は「疎結合プラグイン」方式で
既存の中核ロジック(RAG検索・回答生成・描画ループ)を直接改修するのではなく、
**新規モジュール + 呼び出し元への1行フック + モジュール内サーキット
ブレーカー(`On Error Resume Next`)**の形で追加することを強く推奨する。
理由・実例は `docs/dev/ARCHITECTURE.md` §4を参照。この方式を守れば、
新機能のバグが「その機能が動かないだけ」に留まり、中核機能を道連れに
しない。

### 2.5 1モジュールあたり約30,000字の上限
実機VBAの制約。`tools/vba_lint.py`が警告する。近づいてきたモジュール
(`modAsk`/`modUI`等。`docs/dev/TODO.md` D節参照)へ機能追加する場合は、
先に関連ロジックを別モジュールへ切り出すことを検討する。

## 3. 変更後に必ず通す3ゲート

```bash
python3 tools/vba_lint.py                 # 静的解析(契約・命名・レイヤリング・LO罠)
python3 tools/run_lo_tests.py             # 全モジュールの構文コンパイル検証
python3 build/build_mybookshelf.py --dev  # devビルド+自己検証(--prodも変更が大きければ)
```

いずれか1つでもエラーが出ている状態でPRを出さないこと。3つとも
グリーンであることがPRの最低条件(必須ではないが、実機テストが必要な
変更は`docs/40_受入チェックリスト15分.md`の該当箇所も自己チェックすると
レビューが速くなる)。

## 4. PRの出し方

1. 機能・修正ごとにブランチを切る(命名は自由。例:
   `feature/xxx` / `fix/xxx`)。
2. 上記3ゲートを実行し、全てグリーンであることを確認してからコミット。
3. コミットメッセージには「何を」「なぜ」を書く(実装の詳細は差分を見れば
   分かるため、意図・設計判断・見送った代替案があれば書く。既存コミット
   ログ(`git log --oneline`)がこのリポジトリの文体の参考になる)。
4. PRの説明には最低限、以下を含める:
   - 変更の目的
   - 触ったモジュール一覧
   - 3ゲートの実行結果(全てグリーンである旨)
   - 実機での動作確認有無(未確認なら明記。`docs/dev/TODO.md`のA節に
     追記すべき項目があれば追記する)
5. `docs/dev/CONTRACT`更新・`modules.json`登録が必要な変更は、その旨も
   PR説明に明記する(レビュアーが見落としやすいポイント)。

## 5. レビュー観点(セルフチェックにも使えるチェックリスト)

- [ ] レイヤリング(R1)を破っていないか
- [ ] closed契約モジュールへのPublic追加は`CONTRACT`辞書に反映したか
- [ ] 新規モジュールは`modules.json`に登録したか
- [ ] 新機能はサーキットブレーカー(`On Error Resume Next`)を持ち、失敗時に
      中核機能を止めない設計になっているか
- [ ] Shape名は既存の接頭辞と衝突しない命名になっているか
      (`docs/dev/ARCHITECTURE.md` §3.2の命名規約参照)
- [ ] 共有フォルダI/Oを追加した場合、リトライ(AVロック耐性)・
      collect-then-process(Dir列挙中にKillしない)を実装したか
      (`docs/dev/EDGE_CASES.md` §1.5参照)
- [ ] ファイル名にユーザー由来の文字列を使う場合、`Fnv1a64Hex`等で
      サニタイズ・短縮しているか(`docs/dev/EDGE_CASES.md` §1.4参照)
- [ ] EXP・称号・スキン等の加点系機能は、自己申告で不正加算できない
      設計になっているか(`docs/dev/EDGE_CASES.md` §3参照)
- [ ] `base`や`rem`を識別子に使っていないか、文字列リテラル末尾に`\`が
      無いか(LO構文チェッカーの罠。`docs/dev/EDGE_CASES.md` §1.3参照)

## 6. 迷ったときの判断基準

- 「これは仕様として正しいか」→ `docs/dev/ARCHITECTURE.md` → 無ければ
  `docs/dev/MASTER_SPEC.md`(V1・§3/4/6/12のみ現行有効)。
- 「これは前にやって却下されたアイデアか」→ `docs/dev/FUTURE_IDEAS.md`
  (理由付きで記録済み。同じ議論を繰り返さないための一次資料)。
- 「これは既知の罠か」→ `docs/dev/EDGE_CASES.md`。
- それでも判断がつかない場合は、実装前にPRのDraft/Issueで設計方針を
  共有してからコードを書くことを推奨する(このプロダクトは「新機能は
  疎結合プラグイン方式」という強い設計原則を持つため、大きな改修ほど
  事前合意のコストが低くなる)。
