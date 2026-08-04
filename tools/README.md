# マイ本棚AI — テストハーネス(tools/)

このディレクトリは「マイ本棚AI」の品質保証3層のうち、機械的に自動で回せる
最初の2層(静的Lint・LibreOffice実行テスト)を担当する。3層目(実機受入)は
`docs/40_受入チェックリスト15分.md`(別Wave)を参照。

```
tools/
  vba_lint.py       静的Lint(実行しない。読むだけ)
  run_lo_tests.py   LibreOffice headlessでの実行テスト+コンパイルチェック
  README.md         このファイル
src/test/
  modTestRunner.bas 純ロジックテストの実行エンジン(R4準拠。ExcelでもLOでも動く)
```

呼び出し順の推奨: **vba_lint.py → run_lo_tests.py** (静的チェックの方が速く、
基本的な間違いはここで先に弾ける)。

---

## 1. vba_lint.py — 静的Lint

```bash
python3 tools/vba_lint.py
# 検査対象を変えたいとき(主にこのツール自体のテスト用):
python3 tools/vba_lint.py --path /path/to/some/src
```

`src/**/*.bas` と `*.cls` を対象に、実行せず読むだけで以下を検査する。

| # | 検査内容 | 根拠 |
|---|---|---|
| 1 | `Option Explicit` の有無 | §12 |
| 2 | `Attribute VB_Name` とファイル名の一致 | §7冒頭 |
| 3 | モジュール30,000字以内 | §7冒頭 |
| 4 | `Dim a, b As Long` 型落ち(先頭がVariantになる) | §12 |
| 5 | `As Integer` 禁止(Longを使う) | §12 |
| 6 | モジュール間参照 `modX.Y` の実在検証(Public一覧との突合) | §11.1 |
| 7 | 依存ルールR1(下層→上層参照禁止、例外は機能層→`modUIMain.SetStage`) | §3 |
| 8 | 依存ルールR2(opt直接参照禁止。`modFeatures.InvokeFeature`経由必須) | §3 |
| 9 | 依存ルールR3(`Application.Run`第1引数ホワイトリスト+呼び出し場所制限) | §3 |
| 10 | R4(純ロジックモジュールでのExcelオブジェクトトークン禁止) | §3 |
| 11 | §7公開契約との一致(過不足の両方をエラーに) | §7・§11.1 |
| 12 | (弱い警告)full_text系セル書込みでSafeLeft未経由の疑い | §11.1・§12 |
| 13 | (警告)`.OnAction = "modX.Y"` で配線されたPublic Subの先頭に再入の関所(`modUiLock.BlockIfIngesting` / `modUiLock.Enter`)があるか | R15-2b・実機第4報 RC8 |

検査13の例外は `vba_lint.py` の `ONACTION_GUARD_ALLOWLIST`(名前のリスト)に
理由コメントつきで登録する。取込中でも動くべきハンドラ(中断ボタン等)は
ここへ足す。文字列リテラルで配線された宛先だけを見る(変数経由・文字列連結の
配線は静的に宛先が定まらないため対象外)。

**exit code**: 0=違反なし / 1=ERRORが1件以上。WARN/SKIPはexit codeに影響しない。

### 契約表(CONTRACT辞書)について
`vba_lint.py` 冒頭の `CONTRACT` 辞書に、MASTER_SPEC §7の「モジュール→Public名
一覧」を人間が読んで書き起こしてある(MASTER_SPECの本文をパースしているわけ
ではない)。**§7の契約を変更したら、この辞書も手で追従させること。**
- `closed: True` のモジュールは「契約にあるPublicが全部揃っている」かつ
  「契約に無いPublicが無い」の両方を要求する(過不足ゼロ)。
- `closed: False` のモジュールは、MASTER_SPEC本文が「名前未確定の追加関数を
  切り出してよい」と書いている箇所(例: modPack, modUIMain/modUIShelfの
  opt機能ボタンラッパー, modTestsPureの個別Runサブ)。欠落だけ検出し、
  追加のPublicは許容する。
- 契約表に載っていないモジュール(modExtractorWord等、MASTER_SPECがPublic
  契約を明示していないもの)は対象外(自由)。
- ファイルがまだ存在しないモジュールは **SKIP** 表示になり、エラー扱いには
  ならない(Waveが進んで実装されたら自動的にチェック対象になる)。

### モジュール間参照チェックの仕組み
`modX.Y` のような「モジュール名っぽい識別子(`mod`または`opt`で始まる)+ドット
+識別子」だけを対象に、`modX` が実在するファイルなら `Y` がそのファイルの
Public一覧に無ければエラー、`modX` 自体がまだ存在しないファイルならSKIP扱い
にする。`ws.Cells` のような通常の変数呼び出しは対象にならないので誤検知しない。
**`ThisWorkbook.Worksheets` のようなExcel組み込みオブジェクトモデルの呼び出しは
意図的にチェック対象から除外している**(`ThisWorkbook`はモジュール名として
扱わない)。`ThisWorkbook.cls` 自体の契約(Publicは0件のはず)は別途、通常の
契約チェックで検査される。

### 依存層(R1)の判定基準
MASTER_SPEC §14のディレクトリ配置をそのまま使う:

| ディレクトリ | 層 |
|---|---|
| `src/core` | 基盤層 |
| `src/ingest`, `src/qa`, `src/pack`, `src/stats` | 中間層(§3の部品層+機能層をまとめて1段として扱う) |
| `src/ui` | UI層 |
| `src/opt` | opt層(独立。コアからの参照はR2で別途禁止) |
| `src/test` | テスト層(依存順序の対象外) |

§3の論理図では `modDiag` は機能層(中間層相当)だが、§14の配置では
`src/core`(基盤層)に置く指示になっている。本Lintはディレクトリ基準を
優先するため、`modDiag` は基盤層として扱われる(診断は他機能に依存しない
自己完結処理であるべきなので実用上の問題は無いはずだが、`modDiag`から
中間層/UI層への参照を書くとエラーになる点に注意)。

---

## 2. run_lo_tests.py — LibreOffice実行テスト

```bash
python3 tools/run_lo_tests.py                    # モード1+モード2 両方
python3 tools/run_lo_tests.py --mode pure         # モード1(純ロジック実行)のみ
python3 tools/run_lo_tests.py --mode compile      # モード2(全モジュール構文チェック)のみ
python3 tools/run_lo_tests.py --pure-timeout 120  # モード1のタイムアウト秒(既定120)
python3 tools/run_lo_tests.py --compile-timeout 15 # モード2の1モジュールあたりタイムアウト秒(既定15)
python3 tools/run_lo_tests.py --keep-profile      # 失敗時デバッグ用に一時プロファイルを残す
python3 tools/run_lo_tests.py --verbose
```

前提: `/usr/bin/soffice`(LibreOffice)がインストールされていること。
Pythonは標準ライブラリのみで動く(追加パッケージ不要)。

### モード1: 純ロジック実行
MASTER_SPEC §7.8の `modTestRunner` と、存在すれば `modTypes / modUtil /
modChunker / modPii / modPrompts / modTestsPure` を一時ライブラリへ注入し、
`modTestRunner.RunAllPureTests` → `modTestRunner.ReportText()` を実行して
結果テキストファイルへ書き出す。ファイル冒頭の "PASS n / FAIL m" 行を読み
`FAIL=0` かつ実行時エラー無しなら成功。**まだ実装されていないモジュールは
自動的にスキップして注入する**(欠けていてもエラーにはしない。その分の
テストが単に実行されないだけ)。

### モード2: 構文コンパイルチェック(実行はしない)
`src/` 配下の全 `.bas`/`.cls`(Excel依存モジュールを含む)を1本ずつ、
専用の一時ライブラリに隔離して読み込み、同居させたダミーの
`Chk_Driver.Probe()` だけを呼び出す。これによりライブラリ全体のコンパイルが
強制されるが、対象モジュールの中身(Worksheets/Range/Application/
ThisWorkbook/MsgBox等を触る部分)は実行しない。
`Chk_Driver.Probe()` が短時間(既定15秒)で成功すればPASS、タイムアウトすれば
「構文エラーの疑い」としてFAILにする。

### なぜ「1モジュール=1隔離ライブラリ」なのか
実験の結果、**StarBasicはライブラリ単位でまとめてコンパイルする**ことが
確認できた。1つのライブラリに複数モジュールを同居させると、そのうち1本でも
構文エラーがあると「そのライブラリ内のどのマクロを呼んでも」呼び出しが
ハングする(壊れていないモジュールまで巻き添えになる)。これでは「どの
モジュールが壊れているか」を特定できないため、モード2は必ず1モジュールを
専用ライブラリへ隔離してから検査する。

### なぜタイムアウトで「構文エラー」を判定するのか
LibreOffice Basicは、構文エラーのあるライブラリの中のマクロを呼び出すと
(ヘッドレスでも)応答不能になり、`soffice`プロセスが終了しないまま止まる
ことを実験で確認した(コンパイルエラーのダイアログが裏で出て、入力を待って
いると推測される)。正常にコンパイルできる場合は1秒程度で終わるため、
既定15秒のタイムアウトで十分に区別できる。プロセスの後始末は
coreutilsの `timeout --kill-after=N` に任せている(自前killより確実)。

---

## 3. 技術メモ: 実験して確定させたLibreOffice headlessの挙動

このツールを作る過程で、ドキュメント化されていない(見つけられなかった)
LibreOffice Basicの実際の挙動をいくつも実験で確認する必要があった。
将来このツールを直す人のために記録しておく。

1. **`vnd.sun.star.script:...&location=application` は、真っさらな
   UserInstallationプロファイルに対しては黙って何も実行しない**
   (exit codeは0のまま、指定した出力ファイルも作られない)。
   一度 `--terminate_after_init` で正常に起動を終えたプロファイルでないと
   動かない。本ツールはこのため「雛形プロファイル」を初回だけ作って
   使い回している(`/tmp/mybookshelf_lo_template_profile`)。
2. **モジュールをまたいでPublic Type(構造体)を参照すると、既定の
   StarBasic(非VBA互換モード)ではコンパイルがハングする。**
   各モジュール(注入用の一時コピー)の先頭に `Option VBASupport 1` を
   付けると解消する。本ツールは注入する全モジュールに機械的にこの1行を
   追加している(元ファイルは変更しない)。
3. **ライブラリ単位でコンパイルが行われる**(#2参照の「なぜ1モジュール
   =1隔離ライブラリなのか」を参照)。
4. **Excel固有オブジェクト(Worksheets/Range/Application/ThisWorkbook/
   MsgBox)は、実際にそのコードへ実行が到達しない限りコンパイルは通る。**
   未定義のグローバル識別子の解決は実行時まで遅延される。これを利用して、
   モード2は「対象モジュールの中身を呼ばずにライブラリ全体のコンパイルだけ
   強制する」設計にしている。
5. **`Public Function Foo(...) As String()` のように「配列を返す関数」の
   宣言は、`Option VBASupport 1` を付けてもコンパイルが通らない
   (ハングする)。** 配列を「引数」として受け取るのは問題なく、配列を
   `Variant`に包んで返すのも問題ない。関数の戻り値の型に配列型 `T()` を
   直接書いた場合だけ壊れる。これは実際に `modUtil.SplitKeepNonEmpty`
   (MASTER_SPEC §7.1の契約どおり `As String()`)がそのまま該当し、対処
   しないとモード1のテストが(このモジュールを含むライブラリ全体が)
   全滅する。MASTER_SPEC §11.2の「VBA固有でLOが解釈できない構文が出た
   場合は、lint側で当該構文の代替を規約化する(勝手にテスト対象から
   外さない)」という指示に従い、`run_lo_tests.py` は **.xbaへ変換する
   ときにだけ** `Function Foo(...) As T()` を `Function Foo(...) As
   Variant` へ機械的に書き換える(元の`.bas`/`.cls`ファイルは一切変更
   しない。実Excelでビルドされる成果物は元のままの `As String()` を使う)。
   関数本体が配列をそのまま戻り値へ代入している分にはVariant宣言でも
   実行時の挙動は同一であること(呼び出し側で `Dim r() As String: r =
   Foo(...)` と受けてUBound/LBound/添字アクセスできること)を実験で確認済み。
   **今後、配列を返すPublic関数を新しく書く場合、このルールを知っておくこと。**
   (契約シグネチャ自体は変更しない。あくまでLO実行テストの内部変換の話。)

---

## 4. 既知の限界: 「LOで通る」は「Excelで通る」を保証しない

このテストハーネスはコストゼロで多くのバグ(型落ち・依存違反・契約違反・
構文エラー)を機械的に検出できるが、**LibreOffice BasicとVBA(Excel)は
別の処理系であり、完全に同じ動作をするわけではない**。特に以下は
LibreOffice側のテストが緑でも実機(Excel)で確認すべき既知の差分:

- **セルの実際の書式・数式・条件付き書式・図形(Shapes)の挙動**は
  LO headlessでは(GUIを持たないため)実質検証できない。`modUIMain.
  EnsureLayout` 等のUI構築系は本ハーネスの対象外(§7.6のUI系モジュールは
  モード2の構文チェックは通るが、実際に正しいレイアウトになるかは
  §11.3の実機受入チェックリストで確認する)。
- **`Application.Run` によるリボン関数(`ChatGPT`/`GetEmbeddings`)呼び出し**
  はLO側では実行されない(そもそも存在しない関数なので、実行に到達すれば
  ランタイムエラーになる。本ハーネスはこれらの呼び出しを含むコードパスを
  実際には通らない`Chk_Driver.Probe()`だけを呼ぶ設計にしているため、
  この差分は問題にならない)。
- **日付/ロケール/文字コード周りの細かな挙動**(`Format$`、`CStr`のロケール
  依存表示、cp932とUTF-8の扱い等)はLOとExcelで一致しない場合がある。
  `modUtil`が`Str$`/`Val`を使ってロケール非依存にしているのはこの対策だが、
  全ての箇所が同じ配慮をしているとは限らない。
- **本ハーネスが `.xba` へ変換する際に行う書き換え**(`Option VBASupport 1`
  の付与・`Attribute`行の除去・配列返り値のVariant化)は、あくまでLO実行
  専用の一時コピーに対するものであり、実際にビルドされる`.xlsm`には一切
  影響しない。ただし「LOでは動いたのに変換の副作用でExcelと違う挙動を
  隠してしまっていないか」は常に頭の片隅に置くこと。
- **32,767字を超えるセル書込み**や**xlErrorHandler経由のESC中断**など、
  Excel APIに固有の実行時挙動はLO側では検証できない(§13のエッジケースは
  実機受入チェックリストでカバーする)。

**結論: `vba_lint.py` と `run_lo_tests.py` は「明らかな間違いを早期に、
無料で」弾くための道具であり、§11.3の実機受入チェックリストの代わりには
ならない。** 両方を必ず実施すること。
