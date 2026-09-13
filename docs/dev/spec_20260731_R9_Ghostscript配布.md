# 要件定義 R9: Ghostscriptの配布・自動検出（起草: Fable司令塔 / v2 利用者決定反映）

利用者決定(2026-07-31): **分離方式(zip一体配布)でOK**。ただし
「渡したファイルをそのままフォルダに閉じ込めて、すぐ配布できる状態」にする。
自分で Ghostscript を同封して圧縮する作業はしない。
**テスト段階の今でも「GitHubからダウンロード→解凍するだけ」で使えること。**
埋め込み方式(旧案B)は見送り(必要になれば再検討)。

## ゴール

1. 利用者(および配布先社員)は Ghostscript を別途DL・手動配置しない。
2. リポジトリ自体が「配布可能な形」を常に保持する:
   `dist/MyBookshelf.xlsm` の隣に `dist/Ghostscript/`(gswin32c.exe +
   gsdll32.dll + LICENSE + README)を**コミットして常置**。
   → GitHub の「Code → Download ZIP」(または
   `https://github.com/kojima1459/notebook/archive/refs/heads/<branch>.zip`)
   を落として解凍すれば、dist/ 内の .xlsm がそのまま動く。
3. 大規模配布用にワンコマンド梱包: `python3 build/build_mybookshelf.py --zip`
   で `dist/MyBookshelf_配布.zip`(xlsm + Ghostscript フォルダ)を生成。
   この zip は **gitignore する**(リポジトリ履歴の肥大防止。常置は生ファイルで
   足りるため)。

## 同梱物(雑務担当が取得済み・dist/Ghostscript/)

- Ghostscript 10.03.1 Windows 32bit (AGPL)。gswin32c.exe + gsdll32.dll。
  公式配布 (github.com/ArtifexSoftware/ghostpdl-downloads) の
  gs10031w32.exe から抽出。リソースはDLL内蔵ビルドのため2ファイルで動作。
- LICENSE.txt(AGPL明示) / README.txt(配置規約の説明)。

## VBA側の要件(実装=Sonnet)

R6 の optVision.ResolveGsExe を次の解決順に拡張する:
1. config `ghostscript_path`(明示フルパス)
2. `ThisWorkbook.Path & "\Ghostscript\gswin32c.exe"`(同梱・既定経路)
3. config `ghostscript_search_dirs`(セミコロン区切り。各ディレクトリに
   `\gswin32c.exe` を付けて探す。ITが公式OCRツールの標準配置先を焼き込む口)
4. 見つからない場合: **1回きりの案内カード**(赤エラーでなく)
   「画像PDFの読み取りには Ghostscript が必要です。配布zipの Ghostscript
   フォルダをこのファイルの隣に置くか、場所を指定してください」
   + [フォルダを選ぶ] → Application.FileDialog(msoFileDialogFolderPicker)。
   選択フォルダ直下(または bin\ 直下)に gswin32c.exe があれば
   `ghostscript_path` へ保存し即続行。無ければその旨を1行で。
- 解決結果はセッション内キャッシュ(モジュール変数)。ResetProbe 同様の
  再解決手段として、パス保存時にキャッシュを更新。
- パス合成・候補列挙の判定は純ロジックへ切り出し(optOcrCore に
  GsCandidatePaths(cfgPath, wbDir, searchDirs) 等)、境界値テストを追加。

## build 側の要件(実装=Sonnet・同一ラウンド)

- `--zip` フラグ: dist の完成 xlsm + dist/Ghostscript を
  `dist/MyBookshelf_配布.zip` に梱包(zipfile, ZIP_DEFLATED)。
  発行者用(--publisher)時は `MyBookshelf_発行者用_配布.zip`。
- .gitignore に `dist/*_配布.zip` を追加。
- 既定ビルド(--prod/--dev)の挙動は一切変えない。

## docs

- docs/43_画像PDFのOCR取込設定.md を全面改訂: 「リポジトリのZIPをDL→解凍→
  dist内のxlsmを開くだけ。Ghostscriptは同梱済み」。手動DL・配置の記述を削除
  (トラブルシュートとしてフォルダ選択カードの説明は残す)。
- docs/30 運用ガイドへ: ghostscript_search_dirs の焼き込み手順(IT向け)と
  --zip での大規模配布手順。

## 制約

- CP932外文字禁止 / Integer禁止 / 宣言先頭 / 既存規約に合わせる。
- dist/Ghostscript のバイナリはビルド対象外(build script が消さないこと)。
- R8完了後に着手(同一ファイル群の競合回避)。検証はlint/LOテスト/両ビルド。

## 受入条件

- クリーン環境相当で「リポジトリZIP解凍→dist/MyBookshelf.xlsm と
  dist/Ghostscript が並ぶ→ResolveGsExe が経路2で解決」の経路がコードで成立。
- 経路3(search_dirs)・経路4(案内カード→保存→即続行)の経路がコードで成立。
- --zip で配布zipが作られ、中身に xlsm と Ghostscript/ 一式が入る。
- lint ERROR 0 / LOテスト全PASS / 両ビルド自己検証OK。
