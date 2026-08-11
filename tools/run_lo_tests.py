#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_lo_tests.py — LibreOffice headlessによるVBA実行テスト(MASTER_SPEC.md §11.2)
================================================================================
役割:
    vba_lint.py が「読むだけ」で見つけられる違反を検出するのに対し、本スクリプトは
    実際に LibreOffice(soffice --headless)にVBAソースを読み込ませ、
      モード1: 純ロジックモジュール一式を実行して modTestRunner.RunAllPureTests
               →ReportText の結果(PASS/FAIL)を回収する
      モード2: 全モジュール(Excel依存を含む)を1本ずつ隔離したライブラリに
               読み込み、実行はせず「コンパイルが通るか」だけを確認する
    の2段構えでテストする。

技術メモ(実験して確定させた挙動。ここが今回の技術リスクだった):
    1. soffice --headless で「vnd.sun.star.script:Lib.Mod.Sub?language=Basic&
       location=application」形式のURIを叩いてマクロを実行させるには、
       -env:UserInstallation=file://<profile> で指すプロファイルが
       「一度でも正常に起動を終えたことがある」状態でなければならない。
       手組みしただけの真っさらなプロファイル(script.xlc等だけ用意した状態)
       ではスクリプトが黙って実行されない(exit codeは0のまま何も起きない)。
       このため本スクリプトはまず --terminate_after_init で「一次起動」して
       プロファイルを初期化してから、そこへ Basic モジュールを注入する。
       (このテンプレートプロファイルは使い回して初期化コストを毎回払わない
       ようにしている。詳細は _ensure_template_profile を参照)
    2. モジュール間で Public Type(modTypes.ExtractedPage 等)を跨いで
       参照すると、既定(VBA非互換)のStarBasicコンパイラでは
       コンパイルがスタックし、呼び出しがハングする(タイムアウトでしか
       検知できない)。各モジュール先頭に "Option VBASupport 1" を追加すると
       この問題が解消することを実験で確認した。逆に、この1行が無いと
       modTypes を使う全モジュールのテストが原因不明のタイムアウトになる。
       そのため本スクリプトが注入する全モジュールの先頭に機械的に
       "Option VBASupport 1" を付与する(元の.bas/.clsファイルは変更しない。
       あくまでLO実行用に生成する一時コピーだけに付与する)。
    3. 1つのライブラリ内に構文エラーを含むモジュールが1つでもあると、
       そのライブラリ内の「どのマクロを呼んでも」呼び出しがハングする
       (ライブラリ単位でまとめてコンパイルされるため、壊れていない
       モジュールの実行も巻き添えを食う)。したがってモード2では
       モジュール1本ずつを専用ライブラリに隔離し、他モジュールの構文エラーに
       巻き込まれないようにしている。
    4. Excel固有オブジェクト(Worksheets/Range/Application/ThisWorkbook/
       MsgBox)は、それらに実際に「実行が到達」しない限りコンパイルは通る
       (未定義のグローバル識別子の解決は実行時に遅延される)。これを
       実験で確認済みなので、モード2は「対象モジュールのコンパイルが
       通るか」を、対象モジュール中の何かのPublicプロシージャを実際に
       呼ぶことはせず、同じライブラリに同居させたダミーの
       Chk_Driver.Probe() だけを呼ぶことで検査する(=ライブラリ全体の
       コンパイルを強制するが、対象モジュールの中身は実行しない)。
    5. タイムアウトの検知とプロセス後始末は、Pythonで自前のkill処理を
       書くより信頼できたため、coreutilsの `timeout --kill-after=N` に
       委譲している(実験でsoffice.binの完全終了を確認済み)。
    6. 【重要・既知のVBA/LO差異】 `Public Function Foo(...) As String()` の
       ように「配列を返す関数」の宣言は、Option VBASupport 1を付けても
       LibreOffice Basicではコンパイルが通らない(=ハングする)ことを実験で
       確認した(配列を「引数」として受け取るのは問題なく、配列を
       Variantに包んで返すのも問題ない。関数の戻り値型として配列型
       "T()" をそのまま書いた場合だけが壊れる)。これは実際に
       MASTER_SPEC §7.1 の modUtil.SplitKeepNonEmpty の契約シグネチャ
       ( `As String()` )がそのまま該当し、対処しないとLOテストが
       全滅する。§11.2の指示(「VBA固有でLOが解釈できない構文が出た場合は
       lint側で当該構文の代替を規約化する(勝手にテスト対象から外さない)」)
       に従い、本スクリプトは .xba へ変換する際にだけ
       `Function Foo(...) As T()` を `Function Foo(...) As Variant` へ
       機械的に書き換える(_fix_array_return_types)。元の.bas/.clsファイルは
       一切変更しない。関数本体が配列をそのままReturnValueに代入する分には
       Variant宣言でも実行時の挙動(呼び出し側で `Dim r() As String: r = ...`
       と受けてUBound/LBound/添字アクセスする)は同一であることを実証済み。
       実Excel(VBA)側は元の `As String()` のままビルドされるため、
       この書き換えはLO実行テストの内部実装だけの話であり、§7の公開契約
       (シグネチャ)そのものを変更するものではない。

使い方:
    python3 tools/run_lo_tests.py                  # モード1+モード2 両方
    python3 tools/run_lo_tests.py --mode pure       # モード1のみ
    python3 tools/run_lo_tests.py --mode compile    # モード2のみ
    python3 tools/run_lo_tests.py --keep-profile    # 一時プロファイルを残す(デバッグ用)
    exit code: 0 = 全テストPASS+全モジュールコンパイル成功 / 1 = いずれか失敗
================================================================================
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

TOOLS_DIR = Path(__file__).resolve().parent
MYBOOKSHELF_ROOT = TOOLS_DIR.parent
DEFAULT_SRC_ROOT = MYBOOKSHELF_ROOT / "src"

SOFFICE_CANDIDATES = ["/usr/bin/soffice", "soffice"]

# モード1(純ロジック実行)に含めるモジュール(存在するものだけを注入する)
# 2026-07-11 Wave3(テスト完成担当)で追加: modAppDef/modShelfSync/modPack。
# MASTER_SPEC §7.8はmodShelfSync.DiffDecisionとmodPack.ValidatePackMetaを
# 「純関数として切り出してmodTestsPureから直接検証する」ことを明示指示している
# (これらのモジュール全体がR4準拠というわけではなく、Excelに触れる他のSub/
# Functionと同居しているが、この2関数自体はExcelオブジェクトに触れない)。
# 追加前はこの2モジュールが未注入のため、modTestsPure.RunAll内の
# modShelfSync.DiffDecision呼び出しが実行時エラー12(Variable not defined)に
# なり、テストが「実際のロジックを検証しないまま失敗扱い」になっていた。
# modPack.ValidatePackMetaはmodAppDef.PACK_FORMAT_VERSIONも参照するため
# modAppDefも合わせて追加する。3モジュールとも「対象モジュールを1本だけ
# 隔離してコンパイル」するモード2(run_compile_mode)で既にコンパイル成功が
# 確認済み(Excel専用トークンはtechメモ4のとおり未実行なら未解決のままで
# 良い)。追加後、実際にモード1を実行してPASS/FAIL件数の悪化がないことを
# 確認済み(tools/README.mdまたはWave3完了報告のverification参照)。
# 2026-07-12 Wave3-Tで追加: modTestsPure2。modTestsPureが§7.1の
# 「1モジュール30,000字以内」を超過したため、modPrompts/modShelfSync/
# modPack関連のテスト(TestModPrompts/TestModShelfSync/TestModPack)を
# modTestsPure2.RunAll2へ分割した(src/test/modTestsPure2.bas冒頭コメント
# 参照)。modTestsPure.RunAllの末尾がmodTestsPure2.RunAll2を呼ぶため、
# この一時ライブラリに注入しないと実行時エラー(Variable not defined)に
# なり分割先のテストが「実行されないまま」になる。
PURE_ALLOWLIST = [
    "modTypes", "modUtil", "modChunker", "modPii", "modPrompts",
    "modTestRunner", "modTestsPure", "modTestsPure2",
    "modAppDef", "modShelfSync", "modPack",
    # modSparse: 日本語キーワード検索(文字bigram+BM25+完全一致)。
    # 純ロジックなので実行テストで検証できる。検索精度の要なので必ず載せる。
    "modSparse", "modMode",
    # modChannel: 部門チャンネル。Excel依存が多いが、origin タグの組み立て
    # (ChannelOriginTag)だけは純関数で、ここがズレると切替・更新配信が
    # まるごと空振りする(2026-07-28 レビュー C-1 の実バグ)。コンパイルごと
    # 載せてタグ規約を実行テストで固定する。
    "modChannel",
    # modClarify: 聞き返し。IsNumberChoiceOnly は純関数で、ここがゆるいと
    # 利用者が打った質問が黙って捨てられる(2026-07-28 レビュー H-12 の実バグ)。
    "modClarify",
    # modStats: バッジ表の単一情報源(BadgeCatalog)。判定と表示で表が
    # 二重化して「獲得しても見えないバッジ」が4種あった(解説書 §11-11)。
    # 表の整合(4配列の長さ一致・id重複なし)はここで固定する。
    "modStats",
    # 2026-07-30 R2要件B/C対応で追加。modShelfSync/modPackと同じ考え方
    # (モジュール全体がR4準拠というわけではないが、テストで実際に呼ぶ関数
    # 自体はExcel/COMオブジェクトに触れない)。
    #   modExtractor: GarbledRouteCode / BuildPagesFromGsText / PageArrayCount
    #     が純ロジックで、テストからそれらを呼ぶには本モジュール自体をこの
    #     一時ライブラリへ注入する必要がある。未注入のまま呼ぶと実行時エラー12
    #     (Variable not defined)になる(「対象モジュールを1本だけ隔離して
    #     コンパイル」するモード2では既にコンパイル成功を確認済み=Excel専用
    #     トークンは未実行なら未解決のままで良い、というtechメモ4のとおり)。
    #     2026-08-03 R14-F13: SharedCopyNextChunkLen は呼び出し元が無くなった
    #     ため本体ごと削除した(注入が要る理由は上の3関数に引き継がれている)。
    #   modTestsPure3: modTestsPure/modTestsPure2とも30,000字上限まで残りが
    #     少なく、要件Bの新規テストを追加する場所が無かったための分割先
    #     (src/test/modTestsPure3.bas冒頭コメント参照)。modTestsPure2.RunAll2
    #     の末尾がmodTestsPure3.RunAll3を呼ぶため、この一時ライブラリに
    #     注入しないと同じく実行時エラー12になり分割先のテストが
    #     「実行されないまま」になる(modTestsPure2追加時と同型の理由)。
    "modExtractor", "modTestsPure3",
    # modExtractorPdf(2026-08-03 R13 Phase 0): modExtractorの分割先。
    #   TempBaseNameFor(一時コピー名の導出)が純ロジックで、これを
    #   modTestsPure9 から呼ぶには本モジュール自体をこの一時ライブラリへ
    #   注入する必要がある(未注入だと実行時エラー12)。ファイルI/OやCOMを
    #   使う他の関数はテストから呼ばないので未解決のままでよい(techメモ4)。
    #   modExtractor 側の PageArrayCount / BuildPagesFromGsText を参照するが、
    #   modExtractor も同じ一時ライブラリに入っているため解決できる。
    "modExtractorPdf",
    # modUtilText(2026-07-31 R11-F2): ページ分割の添字計算 GsPageBounds /
    #   CleanTextLen / BlendPerItemMs / ElapsedMsSince を1本化した共通部品。
    #   optOcrCore.GsPageCount と modExtractor.BuildPagesFromGsText の両方が
    #   ここへ委譲するので、未注入だと両者を呼ぶテストが実行時エラー12になる。
    #   ADODB.Streamを使うのは Read/WriteTextFileUtf8 だけで、テストからは
    #   呼ばない(実行に到達しなければコンパイルは通る=techメモ4)。
    "modUtilText",
    # modChrome(2026-07-30 R4要件C/D): ツールバーとヘッダーピルの配置計算。
    # 「Wがいくつでも枠内に収まる/タイトルに重ならない」という保証は、
    # 実機で描いて目視するのではなくここで実行テストとして固定する。
    "modChrome",
    # 2026-07-31 R6(画像PDFのOCR取込)で追加。
    #   optOcrCore: Ghostscriptコマンド文字列の組み立てとページ上限の算数だけを
    #     持つ純ロジック(src/opt配下だが副作用ゼロ)。会社公式ツールが実際に
    #     踏んでいた「gsPathを引用符で囲み忘れる」バグを二度と出さないため、
    #     組み立て結果を1文字単位のゴールデンテストで固定する。
    #   modTestsPure4: modTestsPure3(24,111字)に要件R6のテストを足すと
    #     30,000字上限を超えるための分割先。modTestsPure3.RunAll3の末尾が
    #     modTestsPure4.RunAll4を呼ぶため、未注入だと実行時エラー12になり
    #     分割先のテストが「実行されないまま」になる(modTestsPure3と同型の理由)。
    #   optOcrEta(2026-08-04 R15-5b): optOcrCore から移設した進捗バナー・
    #     残り時間と終了目安・上限/中断/頁欠けのメモの純ロジック。現在時刻を
    #     引数で受け取る設計にしてあるので、終了目安まで含めて1文字単位の
    #     ゴールデンテストで固定できる(未注入だと modTestsPure11/12/13 が
    #     実行時エラー12で丸ごと落ちる)。
    "optOcrCore", "optOcrEta", "modTestsPure4",
    # 2026-07-31 R8(P2P/共有系の修正)で追加。
    #   modShareRule: 共有まわりの判定式だけを集めた純ロジック。感謝状の宛先
    #     解決(origin の名前空間)、到達性プローブ、端末失効、TTLキャッシュ、
    #     同期結果の文面。R8で見つかった実害級の不具合はどれもここの判定式の
    #     間違いで、しかも「Windows+共有フォルダ2台」でしか症状が出ない形で
    #     埋まっていた。実行テストで境界値を固定できる唯一の場所なので必ず載せる。
    #   modTestsPure5: modTestsPure4(17,713字)の分割先。modTestsPure4.RunAll4
    #     の末尾が modTestsPure5.RunAll5 を呼ぶため、未注入だと実行時エラー12に
    #     なり分割先のテストが「実行されないまま」になる(modTestsPure4と同型)。
    "modShareRule", "modTestsPure5",
    # modTestsPure6(2026-07-31 R8b): modTestsPure5(30,000字上限まで残り僅か)の
    #   分割先。敵対的レビューで見つかった「起きたら取り返しがつかないが実機では
    #   報告として上がりにくい」穴(B1 失効判定の戻り値/B7b 時計ズレ/B10 UNC共有
    #   ルート)を固定する。modTestsPure5.RunAll5 の末尾が RunAll6 を呼ぶため、
    #   未注入だと実行時エラー12になりテストが「実行されないまま」になる。
    "modTestsPure6",
    # modTestsPure7(2026-08-01 R12-3): 既存テストモジュールが軒並み30,000字
    #   上限に近く(憲章§4-6)、堅牢化の回帰テストを置くために新設した分割先。
    #   modTestsPure4.RunAll4 の末尾が RunAll7 を呼ぶため、未注入だと実行時
    #   エラー12になりテストが「実行されないまま」になる。
    "modTestsPure7",
    # 2026-08-01(R12-8 テスト補強): 監査「testmeta」Highで指摘された、
    # 純ロジック契約モジュールなのにテストが1件も無くPURE_ALLOWLIST未登録の
    # ため「テストを書いても実行時エラー12で走らない」状態だったモジュール。
    #   modRagParse: 多段RAGのLLM応答パーサ(ParseExpand/ParseSubqueries/
    #     ParseRankOrder/ExtractAnswer)。ExtractAnswerはM-2のプロンプト
    #     インジェクション対策(</thinking>後からの<answer>探索)の中枢で、
    #     回帰しても検知手段が無かった。vba_lint.py の PURE_LOGIC_MODULES には
    #     既に登録済み(R4準拠の実態はあった)だが、こちらの実行テスト用
    #     allowlistには無かったため、テストを書いても走らせられなかった。
    #   modBitwiseOpt / modFollowup: 同じく vba_lint.py の PURE_LOGIC_MODULES
    #     には登録済み(Excel専用トークン0件を実測済み)だが、実行テスト用
    #     allowlistが未登録だった。bit31境界(符号ビット)とマーカー
    #     パース/履歴縮約ループの停止性を実行テストで固定する。
    "modRagParse", "modBitwiseOpt", "modFollowup",
    # modVecCache(2026-08-01 R12-4): セッション内ベクトルキャッシュ。
    #   シートを読むのは PrepareVectors 1本だけで、キャッシュ構築(BuildFrom)・
    #   世代判定(IsStale/StampOf)・内積(DotAt)はいずれも配列だけで完結する
    #   (modShelfSync/modPack と同じ「モジュール全体はR4準拠ではないが、
    #   テストで呼ぶ関数自体はExcelに触れない」型)。等価テスト
    #   (キャッシュ経路 vs 直接パース経路のスコア一致)を走らせるために必須。
    "modVecCache",
    # modTestsPure8(2026-08-01 R12-8): modTestsPure7の容量逼迫による分割先。
    #   modTestsPure7.RunAll7 の末尾が RunAll8 を呼ぶため、未注入だと実行時
    #   エラー12になり分割先のテストが「実行されないまま」になる
    #   (modTestsPure7追加時と同型の理由)。
    "modTestsPure8",
    # modTestsPure9(2026-08-01 R12-4): modTestsPure8の容量逼迫による分割先。
    #   modTestsPure8.RunAll8 の末尾が RunAll9 を呼ぶため、未注入だと実行時
    #   エラー12になり分割先のテストが「実行されないまま」になる。
    "modTestsPure9",
    # modLog(2026-08-03 R13-3b): FriendlyMessage / FriendlyFailMsg は
    #   Select Case と文字列結合だけの純ロジック(シートを触るのは
    #   LogError/LogUsage/TrimLog だけで、テストからは呼ばない)。
    #   modShelfSync/modPack/modVecCache と同じ「モジュール全体がR4準拠では
    #   ないが、テストで呼ぶ関数自体はExcelに触れない」型。未注入のまま
    #   modTestsPure9 から呼ぶと実行時エラー12(Variable not defined)になり、
    #   RC6(案内文の上書き)の回帰テストが走らないまま全部PASSに見える。
    "modLog",
    # modP2PIo(2026-08-03 R13-7c): TeamCodeOf/DeptOf/BeaconDataText/
    #   BeaconTeamField は文字列処理だけの純ロジック(モジュール全体は
    #   Dir/Kill/MkDirを持つためPURE_LOGIC_MODULES非該当だが、テストが
    #   実際に呼ぶのはこの4関数だけで、いずれもExcel/COMに触れない。
    #   modShelfSync/modPack/modVecCache/modLogと同じ型)。未注入のまま
    #   modTestsPure10から呼ぶと実行時エラー12になる。
    # modTestsPure10: modTestsPure9の容量逼迫(WARN帯)による分割先。
    #   modTestsPure9.RunAll9 の末尾が RunAll10 を呼ぶため、未注入だと
    #   実行時エラー12になり分割先のテストが「実行されないまま」になる。
    "modP2PIo", "modTestsPure10",
    # modTestsPure11(2026-08-03 R14-3/R14-4): modTestsPure9の容量逼迫
    #   (残り578字)による分割先。modTestsPure10.RunAll10 の末尾が RunAll11 を
    #   呼ぶため、未注入だと実行時エラー12になり分割先のテストが
    #   「実行されないまま」になる(modTestsPure10追加時と同型の理由)。
    #   検証対象は modExtractorPdf(コピー失敗理由の判定)・optOcrCore
    #   (空入力の分類/バッチ境界/上限メモ/進捗バナー)・modLog(共有読み
    #   失敗の案内文が汎用文言に潰されないこと)で、いずれも既に注入済み。
    # modAskThorough(2026-08-03 R14-8a): 「入念に調べる」専用パイプラインの
    #   最後の段=出典の機械的突合(ExtractCiteTags/NormalizeCiteTag/TagIsKnown/
    #   AnnotateCitations)は LLM を使わない純ロジックで、ここがズレると
    #   「存在しない資料名やページを引用しても誰も気付けない」状態に戻る。
    #   RunThoroughFlow(CallLLM を呼ぶ段)はテストから呼ばないので、
    #   modShelfSync/modPack/modLog と同じ「モジュール全体はR4準拠ではないが、
    #   テストが呼ぶ関数自体はExcel/COMに触れない」型。未注入のまま
    #   modTestsPure11 から呼ぶと実行時エラー12になる。
    # modLive(2026-08-03 R14-8c): 回答本文の記法正規化(NormalizeAnswerText /
    #   AnswerParagraphs)と実況文の言い換え(Humanize)。どちらも文字列だけで
    #   完結する(Shapeを触るのは PaintStage/StyleFooter/StyleAnswerParas で、
    #   テストからは呼ばない)。■見出しの前の空行や【】変換は目で見るしか
    #   確認手段が無かった部分なので、ここで実行テストとして固定する。
    # modTestsPure12(2026-08-03 R14-8): modTestsPure11の容量逼迫による分割先。
    #   modTestsPure11.RunAll11 の末尾が RunAll12 を呼ぶため、未注入だと
    #   実行時エラー12になり分割先のテストが「実行されないまま」になる。
    # modTestsPure13(2026-08-04 R15波2): modTestsPure11(27,888字)/
    #   modTestsPure12(27,001字)のどちらもWARN帯直前で足せないための分割先。
    #   modTestsPure12.RunAll12 の末尾が RunAll13 を呼ぶため、未注入だと
    #   実行時エラー12になり分割先のテストが「実行されないまま」になる。
    # optOcrCache(2026-08-04 R15-7d): 頁チェックポイント。シートを触るのは
    #   保存/読み出し/掃除だけで、テストが呼ぶ鍵の組み立て(DocPrefixFor/
    #   CacheKeyFor/CacheTextFor)は modUtil の純関数しか使わない
    #   (modShelfSync/modPack/modLog と同じ「モジュール全体はR4準拠では
    #   ないが、テストで呼ぶ関数自体はExcelに触れない」型)。未注入のまま
    #   modTestsPure13 から呼ぶと実行時エラー12になり、差し替え検知
    #   (サイズ・更新日時の不一致でキャッシュを使わない)の回帰テストが
    #   走らないまま全部PASSに見える。
    "modTestsPure11", "modAskThorough", "modLive", "modTestsPure12",
    "modTestsPure13", "optOcrCache",
    # modTestsPure14(2026-08-04 R15-FixA): modTestsPure13 の分割先。
    #   modTestsPure13.RunAll13 の末尾が RunAll14 を呼ぶため、未注入だと
    #   実行時エラー12で Fix-A のテストが1件も走らない(13を足したときと同型)。
    "modTestsPure14",
    # modKnowledgeBar(2026-08-03 R14-2a): ToolbarContentRightはShapeを一切
    #   生成しない配置算数だけの関数(ToolbarSpec+modChrome.FlowLeft)。
    #   ToolbarSpecが呼ぶmodPublish.CanPublish/modFeatures.FeatureEnabledは
    #   On Error Resume Next配下のため、両モジュール未注入でも実行時エラー12が
    #   その場で握りつぶされ既定値(False)へ倒れるだけで、テストは壊れない
    #   (techメモ4と同型: 実行に到達した未解決識別子は実行時エラーとして
    #   遅延解決される)。DrawToolbar/ToolButton等Excelに触れる他の関数は
    #   テストから呼ばないので未解決のままでよい。
    # modShelf(2026-08-03 R14-5c): ChunkKeyOrderは文字列結合のみの純関数
    #   (baseKey & "_" & LCase(ext))。IngestFile等Excelに触れる他の関数は
    #   テストから呼ばない。未注入だと modShelf.ChunkKeyOrder 呼び出しが
    #   実行時エラー12になり、拡張子別チャンク設定のフォールバック順序の
    #   回帰テストが「実行されないまま」になる。
    "modKnowledgeBar", "modShelf",
    # modUIShelf(2026-08-04 R15-8b): ParseStatsをPublic化してmodVaultGalleryと
    #   共用した(単純Split(statLine,"|")の重複実装解消・実機第4報 RC1)。
    #   これをmodTestsPure13から直接呼ぶには本モジュール自体をこの一時
    #   ライブラリへ注入する必要がある(未注入だと実行時エラー12)。
    #   Shape/Rangeを触る他の関数(EnsureLayout/RenderShelf等)はテストから
    #   呼ばないので未解決のままでよい(techメモ4と同型)。
    "modUIShelf",
    # modAskMulti(2026-08-05 R16-3A): 複合質問の分解→論点ごとの調査→統合。
    #   モジュール全体はLLM呼び出しと実況を持つが、テストが呼ぶ3本
    #   (ShouldDecompose=発動条件の真理表 / PerPartTopK=論点あたりtopKの下限 /
    #   BuildPartSection=統合入力の1節と部分失敗の文言)はExcel/COMに触れない
    #   (modShelfSync/modPack/modLog/modAskThorough と同じ型)。BuildPartSection は
    #   modPrompts.PART_FAIL_TEXT を読むので modPrompts も要る(既に注入済み)。
    #   未注入のまま modTestsPure14 から呼ぶと実行時エラー12になり、分解ゲートと
    #   部分失敗の文言が「テストを書いても走らない」状態になる。
    "modAskMulti",
    # modAskFocus(2026-08-05 R16-3C): 精読(近傍チャンク束ね)。モジュール全体は
    #   my_knowledge を直接 Range 読みするが、テストが呼ぶ2本
    #   (ParseChunkKey=chunk_id からの (docKey, page, seq) 復元 /
    #    NeighborIdList=文書順に並べて前後 radius を取る)は Excel/COM に触れない
    #   (modShelfSync/modPack/modLog/modAskThorough/modAskMulti と同じ型)。
    #   未注入のまま modTestsPure15 から呼ぶと実行時エラー12になり、文書順復元と
    #   近傍取りが「テストを書いても走らない」状態になる。
    "modAskFocus",
    # modTestsPure15(2026-08-05 R16波3): modTestsPure14 の分割先。
    #   modTestsPure14.RunAll14 の末尾が RunAll15 を呼ぶため、未注入だと
    #   実行時エラー12で波3のテストが1件も走らない(14を足したときと同型)。
    "modTestsPure15",
    # modIntegrity(2026-08-05 R18-2b/2d): データ整合性の観測点。モジュール全体は
    #   my_knowledge/my_manifest/ui_state を読み書きするが、テストが呼ぶ5本
    #   (IndexOfName=source別集計の位置引き / ReconcileStatText=統計文字列の
    #    chunk_count 差し替え / DataShrunk=前回保存時からの減少判定 /
    #    IsVolatilePath=保存が次回に残らない場所の判定 / ShrinkWarnMsg・
    #    VolatileWarnMsg=警告文)はいずれも Excel/COM に触れない
    #   (modShelfSync/modPack/modLog と同じ「モジュール全体はR4準拠ではないが、
    #   テストが呼ぶ関数自体はExcelに触れない」型)。未注入のまま modTestsPure15
    #   から呼ぶと実行時エラー12になり、⑧永続化の判定が「テストを書いても
    #   走らない」状態になる。
    # modProgressBar(2026-08-05 R18-1a/1b): 進捗バナー。テストが呼ぶのは
    #   BarWidthFor(viewport幅→バナー幅の算数)1本だけで、Shape を触る
    #   PaintProgress/ClearProgress/SweepOrphans はテストから呼ばない。
    # modShelfScan(2026-08-05 R18-2f): EnumLooksFailed(列挙0件+台帳に資料あり
    #   =読めなかったと見なす判定)は Long 2つの比較だけの純関数。Dir$ を使う
    #   EnumFolderFiles 等はテストから呼ばない。
    # modTestsPure16(2026-08-05 R18): modTestsPure15(23,011字)にR18の真理表
    #   (約9,300字)を足すと30,000字上限を超えるための分割先。
    #   modTestsPure15.RunAll15 の末尾が RunAll16 を呼ぶため、未注入だと実行時
    #   エラー12でR18のテストが1件も走らない(15を足したときと同型)。
    # modViewport(2026-08-05 R18-3b): 画面ごとのScrollArea宣言。テストが呼ぶのは
    #   純関数(ColLetter/PadPtNeeded/PadUnitsRefine/RightEdgeAt/BoundBottomY)
    #   だけで、Worksheet を触る ApplyScrollBound/FitBandToViewport/BoundAddr は
    #   テストから呼ばない(modProgressBarと同型)。
    # modChunkMeta(2026-08-05 R17 Phase1): section_path/refs_out の抽出。
    #   モジュール全体が R4 純ロジック(PURE_LOGIC_MODULES にも登録)で、
    #   modSparse.NormalizeForSearch と modUtil.SafeLeft しか呼ばない
    #   (どちらも注入済み)。未注入のまま modTestsPure16 から呼ぶと実行時
    #   エラー12になり、構造抽出のゴールデンが「テストを書いても走らない」。
    # modTestsPure17(2026-08-05 R17 Phase1): modTestsPure16(20,739字)に構造
    #   メタの真理表(約7,000字)を足すとWARN帯へ入るための分割先。
    #   modTestsPure16.RunAll16 の末尾が RunAll17 を呼ぶため、未注入だと実行時
    #   エラー12でR17のテストが1件も走らない(16を足したときと同型)。
    # modOutlineBuild / modAskGlobal(2026-08-05 R17 Phase2): 章単位要約と
    #   俯瞰質問。モジュール全体は R4 非準拠(シートI/O・CallLLM を持つ)だが、
    #   ChapterKeyOf / BudgetTake / OutlineActive の3本は副作用ゼロの純関数で、
    #   ここが崩れると「章が別々のキーに割れて要約が章数ぶん増える」
    #   「本文が予算を超えて後ろの章が丸ごと落ちる」「doc_outline が0行でも
    #   俯瞰を試みて実行時エラー」のいずれかが【無音で】起きる。
    #   テストから呼ぶには本モジュールを注入する必要がある(modAskFocus /
    #   modAskMulti と同じ型。未注入だと実行時エラー12)。
    # modTestsPure18(2026-08-05 R17H): modTestsPure17(25,032字)に R17H の
    #   真理表(名寄せマージ・俯瞰シグナル・段0の門)を足すと WARN帯へ入る
    #   ための分割先。modTestsPure17.RunAll17 の末尾が RunAll18 を呼ぶため、
    #   未注入だと実行時エラー12で R17H のテストが1件も走らない(17と同型)。
    # modTestsPure19(2026-08-06 R20-1): modTestsPure16(27,170字)に実機第7報⑦
    #   (右・下余白の3層根治)の真理表を足すと30,000字上限を超えるための
    #   分割先。modTestsPure18.RunAll18 の末尾が RunAll19 を呼ぶため、未注入だと
    #   実行時エラー12で R20-1 のテストが1件も走らない(16〜18と同型)。
    # modDashStat(2026-08-06 R20-1b): テストが呼ぶのは CardWidthFor(帯幅→
    #   KPIカード幅の clamp)1本だけで、Shape とシートI/Oを触る DrawKpiRow /
    #   CountUsageEvent 等はテストから呼ばない(modProgressBar と同型)。
    "modIntegrity", "modProgressBar", "modShelfScan", "modTestsPure16",
    "modViewport", "modChunkMeta", "modTestsPure17",
    "modOutlineBuild", "modAskGlobal", "modTestsPure18",
    "modDashStat", "modTestsPure19",
    # modAppState / modAppAct(2026-08-06 R20-2b/2e・実機第7報①): テストが呼ぶのは
    #   ゲート判定の決定表(modAppAct.GateUsesGeneralHistory)と履歴クリアの
    #   境界(modAppState.ShouldClearGeneralHistory/ClearGeneralMemory/
    #   HasGeneralMemory)で、いずれもExcel/COMに実行が到達しても安全
    #   (modState経由でui_stateシートを読み書きするだけ。シートが無ければ
    #   既定値へ安全に倒れる契約)。
    # modState(同上): ClearGeneralMemory/HasGeneralMemoryが呼ぶ
    #   LoadState/SaveStateの実体。未注入のままだと「シートが無い→既定値」の
    #   安全側フォールバックへ到達する前に modState 自体が実行時エラー12に
    #   なる(=呼び出し元の On Error Resume Next が拾えない種類のエラーで
    #   グループ全体が失敗する)。modShelfSync/modPack/modLog と同じ
    #   「モジュール全体はR4準拠ではないが、テストが呼ぶ関数自体はExcelの
    #   実オブジェクトに触れても落ちない」型。
    "modAppState", "modAppAct", "modState",
    # modBackfill(2026-08-06 R20-3・実機第7報②): 再取込ゼロの「資料の仕上げ」。
    #   テストが呼ぶのは副作用ゼロの4本(ClassifyDoc/LooksLikeBreadcrumbLine/
    #   ConfirmText/ResultText)で、Worksheetsに触れるDetectLegacyDocs/
    #   BackfillOne/BackfillAllはテストから呼ばない(modOutlineBuild/
    #   modChunkMetaと同じ「モジュール全体はR4準拠ではないが、テストが呼ぶ
    #   関数自体はExcelに触れない」型)。未注入のまま modTestsPure19 から
    #   呼ぶと実行時エラー12になる。
    "modBackfill",
    # modTestsPure20(2026-08-06 R20-4/R20-7): modTestsPure19.RunAll19の末尾が
    #   RunAll20を呼ぶため、未注入だと実行時エラー12でR20-4/R20-7のテストが
    #   1件も走らない(16〜19と同型)。
    # modSetupWizard(R20-4b/R20H FA-9): テストが呼ぶのは DeptFromSelector
    #   (数値→部門名の純関数)とWizardShouldRun(保存済みフラグ文字列→
    #   出してよいかの純関数)の2本だけで、OnDeptSetup/RunFirstRunWizard
    #   (InputBox/MsgBoxを持つ)はテストから呼ばない(modOutlineBuild/
    #   modBackfillと同型)。
    # modSkin(R20-7a): テストが呼ぶのは ResolveColor/EffectiveSkin(配色の
    #   Select Caseだけの純関数)で、Shape/シートI/Oを触るBeautifyAll/
    #   ApplyTheme等はテストから呼ばない。EffectiveSkinがsakura/ocean/gold
    #   判定でmodStats.GetStatを呼ぶ経路もあるが、テストが渡すのは
    #   ""/"blueold"/"MSAD"/"light"/"dark"のみでその経路には到達しない
    #   (modDashStatと同型の「モジュール全体はR4準拠ではないが呼ぶ関数は
    #   安全」)。未注入だと実行時エラー12になる。
    "modTestsPure20", "modSetupWizard", "modSkin",
    # modTestsPure21(2026-08-06 R20H レビュー裁定Fix波): modTestsPure20.
    #   RunAll20の末尾がRunAll21を呼ぶため、未注入だと実行時エラー12で
    #   Fix波のテストが1件も走らない(19〜20と同型)。テストが呼ぶ
    #   modViewport.Busy3/modUIShelf.ClearAreaLastRow/
    #   modSetupWizard.WizardShouldRunはいずれも既存の同型注記のとおり
    #   Excel/COMに触れない純関数(各モジュールは既にこの一覧に登録済み)。
    "modTestsPure21",
    # modTestsPure22 / modViewport2(2026-08-07 R21-1 余白の構造完治):
    #   modTestsPure21.RunAll21 の末尾が RunAll22 を呼ぶため、未注入だと
    #   実行時エラー12で R21-1 のゴールデンが1件も走らない(19〜21と同型)。
    #   modViewport2 はテストが呼ぶ HScrollNeeded/SbWidthFrom/ViewMoved/
    #   CompressFactor/HubNeedY/BadgeRowsFor/GridColsFor/GridCardW/
    #   RightGapExceeds/ShelfPadCol がいずれも純関数(Excel/COMに触れない)。
    #   同モジュールの EnsureViewState など Excel 依存の口は、呼ばなければ
    #   未解決のままで良い(techメモ4・modViewport と同型の理由)。
    "modTestsPure22", "modViewport2",
    # modGateway(2026-08-07 R21-2 D1・実機第8報⑧): テストが呼ぶのは
    #   LooksLikeLimitError(モジュール全体はHTTP/COMを持つがこの関数は
    #   文字列比較だけ)。E0204誤爆(査読応答の誤検知)の回帰を固定するために
    #   注入する(modShelfSync/modPack/modLog/modAskThorough と同じ「モジュール
    #   全体はR4準拠ではないが、テストが呼ぶ関数自体はExcel/COMに触れない」型。
    #   未実行の他関数のCreateObject/WinHttp等はtechメモ4のとおり未解決のままで
    #   良い)。未注入のまま modTestsPure18 から呼ぶと実行時エラー12になる。
    "modGateway",
    # modTestsPure23(2026-08-07 R21-3・実機第8報②): 章検出の根治(俯瞰の復旧)。
    #   modTestsPure22.RunAll22 の末尾が RunAll23 を呼ぶため、未注入だと
    #   実行時エラー12でR21-3のゴールデンが1件も走らない(19〜22と同型)。
    #   テストが呼ぶのは modChunker.LooksLikeTocPage/ClassifyLine(既に注入済み
    #   のmodChunker本体)と modOutlineBuild.ChapterKeyOf/GroupChapters/
    #   ChapterKeyMatches/BudgetTake(既に注入済みのmodOutlineBuild本体。
    #   いずれもWorksheetに触れない純ロジック)。BackfillOne/DetectLegacyDocs
    #   等Worksheetに触れる関数はテストから呼ばない(modOutlineBuild/
    #   modBackfillの既存注記と同型)。
    "modTestsPure23",
    # modInstallCheck(2026-08-10 R23b・実機第9報①のMA-3): 自己インストーラの
    #   注入結果を行数で検算するモジュール。テストが呼ぶのは
    #   LineCountMismatch / ExpectedLineCount の2本(文字列とLongだけの純関数)
    #   で、VBProject/Worksheets/MsgBox に触れる VI() はテストから呼ばない
    #   (modOutlineBuild / modBackfill と同じ「モジュール全体はR4準拠では
    #   ないが、テストが呼ぶ関数自体はExcel/COMに触れない」型。未実行の
    #   ThisWorkbook.VBProject 等はtechメモ4のとおり未解決のままで良い)。
    #   未注入のまま modTestsPure23 から呼ぶと実行時エラー12になり、部分注入
    #   検出の境界(末尾空行1本の揺れは一致・1行不足は不一致)が
    #   「テストを書いても走らない」状態になる。
    "modInstallCheck",
    # modTestsPure24(2026-08-10 R25-3・実機第11報⑥): バッジ16枠化の
    #   純ロジック回帰。modTestsPure23.RunAll23 の末尾が RunAll24 を呼ぶため、
    #   未注入だと実行時エラー12でR25-3のゴールデンが1件も走らない
    #   (19〜23と同型)。テストが呼ぶのは modStats.BadgeCatalog のみで、
    #   modStats は既にこの一覧に登録済み(BadgeCatalog自体もWorksheetに
    #   触れない純関数)。
    "modTestsPure24",
    # modTestsPure25(2026-08-10 R27H・敵対的レビュー裁定Fix波): 多様性差し替え
    #   の最小介入(modSparse.DiversitySwapPick)。modTestsPure24.RunAll24 の
    #   末尾が RunAll25 を呼ぶため、未注入だと実行時エラー12でF1のゴールデンが
    #   1件も走らない(19〜24と同型)。テストが呼ぶのは modSparse の純関数
    #   1本だけで、modSparse は既にこの一覧に登録済み。
    "modTestsPure25",
    # modGenPipe(2026-08-11 R26-1): 一般アシスタント3段化の生成パイプライン。
    #   テストが呼ぶのは ParseVerdict(検証応答の判定)/ShouldRunVerifyLoop
    #   (周回上限)/PlanFor(モード分岐表)の3本で、いずれも文字列と数値だけの
    #   純関数(CallLLM・LogUsage・SetStage を持つ RunThorough はテストから
    #   呼ばない=modAskThorough と同じ「モジュール全体はR4準拠ではないが、
    #   テストが呼ぶ関数自体はExcel/COMに触れない」型)。未注入のまま
    #   modTestsPure25 から呼ぶと実行時エラー12になり、verdict のパースと
    #   ループ上限のゴールデンが「テストを書いても走らない」状態になる。
    #   ParseVerdict が呼ぶ modRagParse.IsErrorResponse と PlanFor が呼ぶ
    #   modMode.Normalize は、どちらも既にこの一覧に登録済み。
    "modGenPipe",
]

TEMPLATE_PROFILE_DIR = Path(tempfile.gettempdir()) / "mybookshelf_lo_template_profile"

ATTR_LINE_PATTERN = re.compile(r"^\s*Attribute\s+")
# "Function Foo(...) As T()" のみ(引数側の "name() As T" は対象外)を検出する。
# 理由・実証結果は本ファイル冒頭コメントの技術メモ6を参照。
ARRAY_RETURN_TYPE_PATTERN = re.compile(
    r"(Function\s+\w+\s*\((?:[^()]|\([^()]*\))*\)\s*)As\s+([A-Za-z_]\w*)\s*\(\s*\)",
    re.IGNORECASE | re.DOTALL,
)

XBA_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
    '<script:module xmlns:script="http://openoffice.org/2000/script" '
    'script:name="{name}" script:language="StarBasic">{body}\n'
    "</script:module>\n"
)

XLB_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
    '<library:library xmlns:library="http://openoffice.org/2000/library" '
    'library:name="{libname}" library:readonly="false" library:passwordprotected="false">\n'
    "{elements}\n"
    "</library:library>\n"
)

XLC_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">\n'
    '<library:libraries xmlns:library="http://openoffice.org/2000/library" '
    'xmlns:xlink="http://www.w3.org/1999/xlink">\n'
    "{libs}\n"
    "</library:libraries>\n"
)

CHK_DRIVER_SRC = (
    "Option Explicit\n\n"
    "Public Function Probe() As Boolean\n"
    "    Probe = True\n"
    "End Function\n"
)


def find_soffice() -> str:
    for cand in SOFFICE_CANDIDATES:
        p = shutil.which(cand) or (cand if Path(cand).exists() else None)
        if p:
            return p
    print("[run_lo_tests] soffice が見つかりません。LibreOfficeをインストールしてください。")
    sys.exit(2)


CLS_HEADER_PATTERN = re.compile(
    r"^\s*(VERSION\s+[\d.]+\s+CLASS|BEGIN|MultiUse\s*=.*|END)\s*$", re.IGNORECASE)


def strip_attributes(text: str) -> str:
    """Attribute行に加え、.clsファイル先頭のクラスヘッダブロック
    (VERSION 1.0 CLASS / BEGIN / MultiUse=... / END)も除去する。
    これらはVBEのエクスポート形式であってBasicソースではないため、
    残すとLO Basicが構文エラー(ハング)になる。"""
    lines = text.splitlines()
    out = []
    in_header = True
    for l in lines:
        if in_header:
            if CLS_HEADER_PATTERN.match(l) or ATTR_LINE_PATTERN.match(l) or l.strip() == "":
                continue
            in_header = False
        if not ATTR_LINE_PATTERN.match(l):
            out.append(l)
    return "\n".join(out)


def fix_array_return_types(text: str) -> str:
    """"Function Foo(...) As T()" を "Function Foo(...) As Variant" に書き換える
    (LO実行専用の一時変換。理由は本ファイル冒頭コメントの技術メモ6を参照)。"""
    return ARRAY_RETURN_TYPE_PATTERN.sub(lambda m: m.group(1) + "As Variant", text)


def to_module_body(source_text: str) -> str:
    """Attribute行を除去し、配列返り値の宣言をVariantへ書き換え、先頭に
    Option VBASupport 1 を付与する(理由は本ファイル冒頭コメント参照)。"""
    body = strip_attributes(source_text)
    body = fix_array_return_types(body)
    return "Option VBASupport 1\n" + body


def write_module_xba(lib_dir: Path, name: str, source_text: str) -> None:
    body = xml_escape(to_module_body(source_text))
    xba = XBA_TEMPLATE.format(name=name, body=body)
    (lib_dir / f"{name}.xba").write_text(xba, encoding="utf-8")


def write_library(profile_dir: Path, lib_name: str, modules: dict[str, str]) -> None:
    """modules: {モジュール名: 元の.bas/.clsソーステキスト}"""
    lib_dir = profile_dir / "user" / "basic" / lib_name
    lib_dir.mkdir(parents=True, exist_ok=True)
    elements = []
    for name, src in modules.items():
        write_module_xba(lib_dir, name, src)
        elements.append(f' <library:element library:name="{name}"/>')
    xlb = XLB_TEMPLATE.format(libname=lib_name, elements="\n".join(elements))
    (lib_dir / "script.xlb").write_text(xlb, encoding="utf-8")


def register_libraries(profile_dir: Path, lib_names: list[str]) -> None:
    libs = [' <library:library library:name="Standard" library:link="false"/>']
    for n in lib_names:
        libs.append(f' <library:library library:name="{n}" library:link="false"/>')
    xlc = XLC_TEMPLATE.format(libs="\n".join(libs))
    (profile_dir / "user" / "basic" / "script.xlc").write_text(xlc, encoding="utf-8")


def ensure_template_profile(soffice: str, verbose: bool) -> Path:
    """一度だけ soffice を「一次起動」させて雛形プロファイルを作る(§技術メモ1)。
    既にあれば使い回す(--fresh-template で強制作り直し可能)。"""
    marker = TEMPLATE_PROFILE_DIR / "user" / "basic" / "Standard" / "script.xlb"
    if marker.exists():
        return TEMPLATE_PROFILE_DIR
    if verbose:
        print(f"[run_lo_tests] 雛形プロファイルを初期化中: {TEMPLATE_PROFILE_DIR}")
    if TEMPLATE_PROFILE_DIR.exists():
        shutil.rmtree(TEMPLATE_PROFILE_DIR)
    cmd = [
        "timeout", "--kill-after=5", "60",
        soffice, "--headless", "--invisible", "--nologo", "--norestore",
        f"-env:UserInstallation=file://{TEMPLATE_PROFILE_DIR}",
        "--terminate_after_init",
    ]
    subprocess.run(cmd, capture_output=True, text=True)
    if not marker.exists():
        print("[run_lo_tests] 雛形プロファイルの初期化に失敗しました(soffice起動不可の可能性)。")
        sys.exit(2)
    return TEMPLATE_PROFILE_DIR


def fresh_profile_copy(template: Path, dest: Path) -> None:
    shutil.copytree(template, dest)


def run_uri(soffice: str, profile_dir: Path, uri: str, timeout_sec: int) -> tuple[int, str, str]:
    cmd = [
        "timeout", "--kill-after=5", str(timeout_sec),
        soffice, "--headless", "--invisible", "--nologo", "--norestore",
        f"-env:UserInstallation=file://{profile_dir}",
        uri,
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    return proc.returncode, proc.stdout, proc.stderr


def discover_modules(src_root: Path) -> list[tuple[str, Path]]:
    """(モジュール名, パス) のリスト。モジュール名はAttribute VB_Nameがあればそれ、
    無ければファイル名(拡張子抜き)。"""
    out = []
    for path in sorted(list(src_root.rglob("*.bas")) + list(src_root.rglob("*.cls"))):
        text = path.read_text(encoding="utf-8", errors="replace")
        name = path.stem
        m = re.search(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"', text, re.MULTILINE)
        if m:
            name = m.group(1)
        out.append((name, path))
    return out


# ==============================================================================
# モード1: 純ロジック実行(RunAllPureTests → ReportText)
# ==============================================================================
def run_pure_mode(soffice: str, template: Path, all_modules: dict[str, Path],
                   work_dir: Path, timeout_sec: int, verbose: bool) -> tuple[bool, str]:
    print("=" * 78)
    print("モード1: LibreOffice上で純ロジックテスト(modTestRunner.RunAllPureTests)を実行")
    print("=" * 78)

    pure_srcs: dict[str, str] = {}
    missing = []
    for name in PURE_ALLOWLIST:
        path = all_modules.get(name)
        if path is None:
            missing.append(name)
            continue
        pure_srcs[name] = path.read_text(encoding="utf-8", errors="replace")

    if missing:
        print(f"  (未実装のため注入をスキップ: {', '.join(missing)})")

    if "modTestRunner" not in pure_srcs:
        print("  modTestRunner.bas が見つからないためモード1を実行できません。")
        return False, "modTestRunner.bas not found"

    out_path = work_dir / "pure_result.txt"
    if out_path.exists():
        out_path.unlink()

    test_main_src = (
        "Option Explicit\n\n"
        "Sub Main\n"
        "    On Error Resume Next\n"
        "    Err.Clear\n"
        "    modTestRunner.RunAllPureTests\n"
        "    Dim runErr As String\n"
        "    If Err.Number <> 0 Then\n"
        '        runErr = "RUNNER_ERROR " & Err.Number & ": " & Err.Description\n'
        "        Err.Clear\n"
        "    End If\n"
        "    On Error GoTo 0\n\n"
        "    Dim iFile As Integer\n"
        "    iFile = FreeFile\n"
        f'    Open "{out_path.as_posix()}" For Output As #iFile\n'
        "    Print #iFile, modTestRunner.ReportText()\n"
        "    If Len(runErr) > 0 Then Print #iFile, runErr\n"
        "    Close #iFile\n"
        "End Sub\n"
    )

    profile_dir = work_dir / "profile_pure"
    fresh_profile_copy(template, profile_dir)
    modules = dict(pure_srcs)
    modules["TestMain"] = test_main_src
    write_library(profile_dir, "MbPureRun", modules)
    register_libraries(profile_dir, ["MbPureRun"])

    uri = "vnd.sun.star.script:MbPureRun.TestMain.Main?language=Basic&location=application"
    rc, out, err = run_uri(soffice, profile_dir, uri, timeout_sec)

    if not out_path.exists():
        msg = f"実行結果ファイルが生成されませんでした(soffice exit={rc}, timeout={timeout_sec}s)"
        print(f"  FAIL: {msg}")
        if verbose:
            print(f"    stdout: {out.strip()}")
            print(f"    stderr: {err.strip()}")
        return False, msg

    report = out_path.read_text(encoding="utf-8", errors="replace").strip()
    print(report if report else "(空の結果)")

    m = re.search(r"FAIL\s+(\d+)", report)
    fail_count = int(m.group(1)) if m else None
    has_runner_error = "RUNNER_ERROR" in report

    ok = (fail_count == 0) and not has_runner_error
    return ok, report


# ==============================================================================
# モード2: 全モジュールのコンパイルチェック(実行はしない)
# ==============================================================================
def safe_lib_name(module_name: str) -> str:
    return "Chk_" + re.sub(r"[^A-Za-z0-9_]", "_", module_name)


def run_compile_mode(soffice: str, template: Path, all_modules: dict[str, Path],
                      work_dir: Path, timeout_sec: int, verbose: bool) -> tuple[bool, list[tuple[str, bool, str]]]:
    print("\n" + "=" * 78)
    print("モード2: 全モジュールの構文コンパイルチェック(実行はしない)")
    print("=" * 78)

    modtypes_path = all_modules.get("modTypes")
    modtypes_src = modtypes_path.read_text(encoding="utf-8", errors="replace") if modtypes_path else None

    results: list[tuple[str, bool, str]] = []
    all_ok = True

    for name, path in sorted(all_modules.items()):
        target_src = path.read_text(encoding="utf-8", errors="replace")
        lib_name = safe_lib_name(name)

        modules_for_lib: dict[str, str] = {}
        if name != "modTypes" and modtypes_src is not None:
            modules_for_lib["modTypes"] = modtypes_src
        modules_for_lib[name] = target_src
        modules_for_lib["Chk_Driver"] = CHK_DRIVER_SRC

        profile_dir = work_dir / f"profile_{lib_name}"
        fresh_profile_copy(template, profile_dir)
        write_library(profile_dir, lib_name, modules_for_lib)
        register_libraries(profile_dir, [lib_name])

        uri = f"vnd.sun.star.script:{lib_name}.Chk_Driver.Probe?language=Basic&location=application"
        t0 = time.time()
        rc, out, err = run_uri(soffice, profile_dir, uri, timeout_sec)
        elapsed = time.time() - t0

        ok = (rc == 0)
        detail = f"exit={rc} ({elapsed:.1f}s)"
        if rc == 124:
            detail = f"タイムアウト({timeout_sec}s) — 構文エラーの疑い"
        results.append((name, ok, detail))
        all_ok = all_ok and ok

        status = "PASS" if ok else "FAIL"
        print(f"  {status:<4} {name:<24} {detail}")

        # 使い終わったプロファイルは都度削除してディスクを節約
        shutil.rmtree(profile_dir, ignore_errors=True)

    return all_ok, results


def main() -> int:
    parser = argparse.ArgumentParser(description="マイ本棚AI LibreOffice実行テスト")
    parser.add_argument("--path", type=str, default=str(DEFAULT_SRC_ROOT), help="対象src(既定: mybookshelf/src)")
    parser.add_argument("--mode", choices=["pure", "compile", "all"], default="all")
    parser.add_argument("--pure-timeout", type=int, default=120, help="モード1のタイムアウト秒(既定120)")
    parser.add_argument("--compile-timeout", type=int, default=15, help="モード2の1モジュールあたりタイムアウト秒(既定15)")
    parser.add_argument("--keep-profile", action="store_true", help="一時プロファイルを削除せず残す(デバッグ用)")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    src_root = Path(args.path).resolve()
    if not src_root.exists():
        print(f"[run_lo_tests] 対象ディレクトリが存在しません: {src_root}")
        return 2

    soffice = find_soffice()
    template = ensure_template_profile(soffice, args.verbose)

    modules_list = discover_modules(src_root)
    all_modules: dict[str, Path] = {name: path for name, path in modules_list}
    print(f"[run_lo_tests] soffice={soffice}")
    print(f"[run_lo_tests] 対象モジュール数: {len(all_modules)} (in {src_root})")

    work_dir = Path(tempfile.mkdtemp(prefix="mybookshelf_lo_run_"))
    overall_ok = True
    try:
        if args.mode in ("pure", "all"):
            ok, _report = run_pure_mode(soffice, template, all_modules, work_dir, args.pure_timeout, args.verbose)
            overall_ok = overall_ok and ok

        if args.mode in ("compile", "all"):
            ok, _results = run_compile_mode(soffice, template, all_modules, work_dir, args.compile_timeout, args.verbose)
            overall_ok = overall_ok and ok
    finally:
        if args.keep_profile:
            print(f"\n[run_lo_tests] --keep-profile 指定のため一時ディレクトリを残します: {work_dir}")
        else:
            shutil.rmtree(work_dir, ignore_errors=True)

    print("\n" + "-" * 78)
    print("結果: OK(exit code 0)" if overall_ok else "結果: NG(exit code 1)")
    print("-" * 78)
    return 0 if overall_ok else 1


if __name__ == "__main__":
    sys.exit(main())
