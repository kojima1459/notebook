#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
vba_lint.py — マイ本棚AI VBAソースの静的Lint(MASTER_SPEC.md §11.1)
================================================================================
役割:
    mybookshelf/src/ 配下の全 .bas / .cls を対象に、実行せずに読める範囲の
    契約違反・危険なコードパターンを機械的に検出する。「1-I テストハーネス」
    3層(静的Lint/LO実行/実機受入)のうち最初の層で、CIの門番として一番
    軽く・一番よく回すことを想定している。

設計判断(なぜこう作ったか):
    ・§7の公開契約(モジュール→Public名一覧)は、このファイル内に
      CONTRACT という辞書として「書き起こして埋め込む」よう指示されている
      (MASTER_SPEC本文をパースするのではなく、人間がMASTER_SPECを読んで
      転記する方式)。§7に無い/曖昧なモジュール(modExtractorWord等の
      アダプタ系、opt機能のUIラッパー等)はCONTRACTに載せない=チェック対象外
      とし、「まだ書かれていないモジュール」はSKIP(エラーにしない)。
    ・コメント剥がし: Lintの各種トークン検査(Worksheets禁止・opt直接参照
      禁止・Application.Runホワイトリスト等)は、ソース中の「日本語コメント
      で仕様を説明している行」を誤検知しないよう、必ず一度コメントを
      取り除いた「コード部分」に対して行う。実際、modTypes.bas 冒頭の
      設計判断コメントには「Worksheets/Range/Application/ThisWorkbook/
      MsgBox」という語がそのまま書かれており、コメント除去をサボると
      自分自身のドキュメントを違反として誤検知する。
    ・行連結(" _" 継続行)は1つの論理行にまとめてから解析する。契約シグ
      ネチャの多くは複数行にまたがる関数宣言なので、これをやらないと
      Public宣言の抽出やDim型落ち検出が正しく動かない。
    ・モジュール間参照(modX.Y)の実在検証は「modX/optX/ThisWorkbookの形を
      した識別子」だけを対象にする。ws.Cells や rs.Fields のような通常の
      オブジェクト変数はモジュール名の集合に含まれないため誤検知しない。
    ・依存層(R1)の判定はMASTER_SPEC §14のディレクトリ配置で行う指示なので、
      src/core=基盤層、src/{ingest,qa,pack,stats}=部品層+機能層をまとめた
      「中間層」、src/ui=UI層、src/opt=opt層、src/test=テスト層、とした。
      §3の論理図はmodDiagを機能層に置くが、実ファイルはsrc/core(基盤層)に
      置かれる指示(§14)なので、本Lintではディレクトリ基準を優先する
      (ディレクトリと論理層の食い違いはmodDiagの1件のみで、意味的にも
      「診断は他機能に依存しない自己完結処理」であるべきなので実害は無い)。
    ・SafeLeft未経由のfull_text系セル書込みチェックは§11.1に記載がある
      「警告」項目として実装したが、静的解析だけでは変数の由来を正確に
      追えないため誤検知が出やすい。exit codeには影響させない「弱い警告」
      として出すに留める(§12のSafeLeft原則をコードレビューで補完する
      前提)。

使い方:
    python3 tools/vba_lint.py                 # mybookshelf/src 配下を検査
    python3 tools/vba_lint.py --path <dir>     # 検査対象ディレクトリを変更(主にテスト用)
    exit code: 0 = 違反なし / 1 = 違反あり(ERRORが1件以上)
================================================================================
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
MYBOOKSHELF_ROOT = TOOLS_DIR.parent
DEFAULT_SRC_ROOT = MYBOOKSHELF_ROOT / "src"

MAX_MODULE_CHARS = 30000
# 残り2,000字を切ったら警告する(バグ修正1件ぶんの余裕がある状態を保つため)。
MODULE_WARN_CHARS = 28000

# ------------------------------------------------------------------------------
# §7 契約シグネチャ表: モジュール名 -> {"closed": bool, "required": [Public名...]}
#   closed=True  : requiredと実際のPublic集合が完全一致でなければERROR
#   closed=False : requiredは全部無いとERRORだが、追加のPublicがあっても
#                  ERRORにしない(MASTER_SPEC本文が「純関数に切り出す」等
#                  名前未確定の追加を許容する書き方をしている箇所)
#   ここに載っていないモジュール名は契約チェックの対象外(自由)。
# ------------------------------------------------------------------------------
CONTRACT: dict[str, dict] = {
    # ---- 7.1 基盤層 ----
    "modAppDef": {
        "closed": True,
        "required": [
            "APP_NAME", "APP_VERSION", "PACK_FORMAT_VERSION",
            "SH_HOWTO", "SH_HOME", "SH_SHELF", "SH_DASH", "SH_CONFIG",
            "SH_KNOWLEDGE", "SH_VECTORS", "SH_MANIFEST", "SH_STATS",
            "SH_USAGE", "SH_ERRLOG", "SH_UISTATE",
            "SH_NEXUS_DASH",   # Nexus専用ダッシュボード(Phase 3)のシート名
            # 2026-08-01(R12-1-8): 本棚チャンク上限の代替値。呼び出し側4箇所
            # (modShelf/modShelfSync/modShelfBatch/modDashStat)で 20,000 と
            # 10,000 に割れていたものを1箇所へ集約した。
            "DEFAULT_SHELF_MAX_CHUNKS",
            # 2026-08-05(R17波0): 構造メタデータ(section_path/refs_out)の
            # 格納先シート名。Phase1本体はまだ無い(波0は器のみ)。
            "SH_CHUNK_META",
            # 2026-08-05(R17 Phase2): 章単位要約(doc_outline)の格納先シート名。
            # 俯瞰質問(verdict=global)が読む唯一のシート。
            "SH_DOC_OUTLINE",
            # 2026-08-05(R17 Phase3): 用語の表記ゆれ辞書(synonyms)の格納先シート名。
            "SH_SYNONYMS",
        ],
    },
    "modTypes": {
        "closed": True,
        "required": ["ExtractedPage", "ShelfChunk", "Hit"],
    },
    "modConfig": {
        "closed": True,
        "required": ["EnsureLoaded", "GetString", "GetLong", "GetDouble", "GetBool", "SetValue"],
    },
    # FriendlyFailMsg(2026-08-03 R13-3b): 取込失敗を利用者へ伝える1文の決定を
    # 1箇所に集めたもの。errDetail に「Wordを開いたままに…」等の行動可能な
    # 案内があればそれを優先し、E0302の汎用文言は拡張子で分岐して docx/doc に
    # Ghostscript の話をしない(実機第2報 RC6: 正確な案内が汎用文言で
    # 上書きされ、利用者に一度も届いていなかった)。純粋な文字列処理なので
    # modTestsPure9 が分岐を固定する。
    # SharedReadFailMsg(2026-08-03 R14-3b): 共有のファイルを読み取れなかった
    # ときの1文。取込経路(modExtractorPdf)とOCR経路(optGsTxt)の両方が
    # 同じ文を出す必要があり、opt層はコア基盤層しか参照できない(§7.7)ため、
    # 文言の一次情報の置き場はここが唯一の交点になる(憲章§4-5)。
    # CopyFailMsgOf(2026-08-03 R14-F3/F11): 一時コピー失敗の【種類】ごとの
    # 利用者向けの1文(空・ロック・大きすぎる・一時名の枯渇・特殊文字)。
    # 全ての失敗に「ファイル名を変えてください」と言っていたのを種類別に
    # 分けたもので、置き場は SharedReadFailMsg と同じ理由でここになる。
    # ReadOnlyWarnMsg / SaveFailMsg(2026-08-04 R15-3・実機第4報 RC9): 読み取り
    # 専用で開かれた警告と、中間保存に失敗したときの1文。どちらも複数の
    # 呼び出し元(modDiag起動時案内/E0805のコード表/modShelfBatchの保存失敗)が
    # 同じ文を出す必要があるため、置き場は上の2関数と同じ理由でここになる。
    "modLog": {
        "closed": True,
        "required": ["LogError", "LogUsage", "FriendlyMessage", "ShowError",
                     "FriendlyFailMsg", "SharedReadFailMsg", "CopyFailMsgOf",
                     "ReadOnlyWarnMsg", "SaveFailMsg"],
    },
    # modChatLog: チャット履歴シート("チャット履歴")への質問/回答記録。
    # 公開APIはLogTurnのみ(書込失敗はDebug.Printのみ=modLogの「ログで死なない」方針踏襲)。
    "modChatLog": {
        "closed": True,
        "required": ["LogTurn"],
    },
    "modUtil": {
        "closed": True,
        "required": [
            "Fnv1a64Hex", "NormalizeForHash", "VectorToCsv", "CsvToVector",
            "DotProduct", "L2Normalize", "HasVector", "SplitKeepNonEmpty",
            "TruncateAndRenorm", "VectorToCsvPrec",
            "HumanBytes", "HumanSeconds", "SafeLeft", "NowStamp",
            "FileNameOf", "ExtOf", "IsSameTimestamp",
            # EtaText/ProgressText: バッチ処理の進捗実況(2026-07-31 R7 B-1)。
            # 「残り約N分」の算数と文言整形。modEmbed/modEnrichが計測値を
            # 渡すだけで済むよう純ロジック側に置き、modTestsPure4で固定する。
            "EtaText", "ProgressText",
            # JoinPagedText/SplitPagedText: 画像PDFのOCR(R6)で複数ページの
            # 結果を「文字列1本」の境界(modFeatures.InvokeFeature)越しに運ぶ
            # ための行マーカー符号化/復号。純文字列処理なのでmodUtilに置く。
            "JoinPagedText", "SplitPagedText",
            # DescribeComError(2026-07-31 R11-D / 監査2 指摘5): COM失敗の
            # 日本語案内。modExtractor/Word/Excel/Acrobat の4重実装を1本へ
            # まとめたもの。アプリ名を引数で差し替えるだけの純ロジック。
            "DescribeComError",
            # DeobfuscateSecret: build/build_mybookshelf.py obfuscate_secret() と対の
            # 復号(XOR+16進)。configシートに置くAPIキーを平文で持たないための軽い
            # 難読化解除。modGateway.DirectEmbedSliceのみが呼ぶ想定だが、R4純ロジック
            # (Excelオブジェクト非依存)のためmodUtilに置く。
            "DeobfuscateSecret",
        ],
    },
    "modGateway": {
        "closed": True,
        # RunLimitCheck: リボン公開API確定対応(裁定D3・RIBBON_API_CONFIRMED.md)。
        # 起動時のAIリボン利用期限チェック。True=続行不可(公式サンプルの解釈)。
        # GetEmbeddingsBatch: RAG再設計(RAG_OVERHAUL_DESIGN.md §E-1)。
        # ribbon(単発ループ)/direct(Azure配列POST)/mock を透過切替する唯一のバッチ窓口。
        # RibbonEmbedRange / SerializeVector: 2026-07-28 に direct 経路を
        # modGatewayDirect へ切り出した際、フォールバック先とベクトルCSV化を
        # 共有するため公開した(経路が変わってもCSVの形は変えない)。
        # ConsumeStepBuf(2026-08-03 R13 F8): 段ごとの所要時間は1段1行で
        # usage_log へ書くと1問10行前後になり、2,000行ローテーションが
        # 約180問で一周して feedback_green 等の履歴を押し出す。ここに貯め、
        # 質問の終わりに modAsk が1行(ask_steps)へまとめて書き出す。
        "required": ["CallLLM", "GetEmbedding", "RibbonAvailable", "TryRibbonRun",
                     "LooksLikeLimitError", "RunLimitCheck", "GetEmbeddingsBatch",
                     "RibbonEmbedRange", "SerializeVector", "ConsumeStepBuf"],
    },
    "modFeatures": {
        "closed": True,
        "required": ["FeatureEnabled", "ModulePresent", "InvokeFeature"],
    },
    "modDiag": {
        "closed": True,
        # WarnIfReadOnly(2026-08-04 R15-3b): 起動時の読み取り専用検知。
        # QuickHealthCheck と同じ「起動時の環境チェック」で、modBoot が
        # 30,000字上限まで残り僅かなため表示・記録までここで完結させる。
        "required": ["RunDiagnostics", "QuickHealthCheck", "RecentErrorsForClipboard",
                     "WarnIfReadOnly"],
    },
    # modIntegrity(2026-08-05 R18-2b/2d・実機第5報⑧): データ整合性の観測点。
    # 「昨日入れた資料が今日は0件」に気付く仕組みがアプリ側に1つも無かった
    # (表示は my_manifest.chunk_count をそのまま出すだけで実データと突き合わせず、
    # 保存が次回に残ったかを確かめる記録も無かった=調査agent1)。憲章§4-2。
    # ReconcileChunkCount: 台帳と my_knowledge の実行数の突合+自動修復
    #   (modShelf.SourceList から1行)。
    # RecordSaveMark: 保存成功時に (行数, FullName) を ui_state へ控える
    #   (modShelfBatch.SaveCheckpoint から1行)。
    # WarnAtStartup: 起動時の突合(modBoot から1行)。
    # ReconcileStatText / DataShrunk / IsVolatilePath / ShrinkWarnMsg /
    #   VolatileWarnMsg: Excelに触れない純ロジック(modTestsPure15 が固定)。
    # 置き場が基盤層なのは modBoot/modShelf/modShelfStore がいずれも30,000字上限
    # 近くで判定本体を置けないため(modDiag.WarnIfReadOnly と同型の判断)。
    "modIntegrity": {
        "closed": True,
        "required": ["IndexOfName", "ReconcileStatText", "ReconcileChunkCount",
                     "RecordSaveMark", "DataShrunk", "IsVolatilePath",
                     "ShrinkWarnMsg", "VolatileWarnMsg", "WarnAtStartup"],
    },
    # ---- 7.2 取込層 ----
    # modExtractorWord / modExtractorExcel / modExtractorAcrobat は
    # MASTER_SPECが個別のPublic契約を明示していないため対象外(自由)。
    "modExtractor": {
        "closed": True,
        # SharedCopyNextChunkLen(2026-08-03 R14-F13で削除): R14-3a で共有読み
        # コピーが ADODB.Stream 一本になった時点で呼び出し元が消え、契約と
        # テストだけが残っていた。R14-F3 でクラシックの1MB分割コピーを
        # 復活させたが、分割の算数は modExtractorPdf 側のループに3行で書く
        # (モジュールを跨いだPublicを1つ減らす)。
        # GarbledRouteCode: 2026-07-31 R6追補。「全ページ化け」のPDFをOCR経路
        # (E0303)へ回すか、従来どおり化けたまま続行するかの分岐だけを純関数に
        # 切り出したもの(modTestsPure4が検証する)。
        # BuildPagesFromGsText: 2026-07-31 R10c(L1)。Ghostscript(txtwrite)の
        # 出力を改ページ文字で割ってExtractedPage配列にする純ロジック。先頭/末尾の
        # 空ページの読み飛ばしがズレると出典ページ番号が丸ごと1つずれるため、
        # modTestsPure6 が境界を固定する。R2によりコア層からoptOcrCoreを呼べず、
        # ページ数の数え方が optOcrCore.GsPageCount と2箇所に分かれているので、
        # 両者の突き合わせもそちらのテストで担保している。
        # PageArrayCount: 2026-08-03 Phase 0(modExtractorPdf分割)。未初期化/空
        # 配列でも実行時エラー9を出さずに0件と数える唯一の実装。分割後は
        # modExtractorPdf.DropGarbledPages も同じ数え方を必要とするため、
        # 2箇所に同じ実装を置かない目的で公開した(憲章§4-5)。
        "required": ["ExtractFile", "SupportedExts",
                     "GarbledRouteCode", "BuildPagesFromGsText", "PageArrayCount"],
    },
    # modExtractorPdf(2026-08-03 R13 Phase 0): modExtractor が28,000字のWARN帯に
    # 達したための分割先。PDFの3経路フォールバック(GS→Word→Acrobat)、
    # ローカル一時コピー(共有読み)、文字化けページの除去を持つ。
    # TempBaseNameFor(R13-2): 一時コピー名 "mbtmp_<FNV-1a 64bit 16桁>.<元拡張子>"
    # の導出だけを切り出した純ロジック。元ファイル名をそのまま連結する旧方式は
    # CP932非対応文字(NFD分解濁点 U+3099 等)でディスク上の実名とGhostscriptへ
    # 渡す文字列が食い違う事故を起こしたため、境界を modTestsPure9 が固定する。
    # IsThinExtract / ThinExtractMemoFor(R13-3a): 「取り込めてはいるが本文が
    # 薄すぎる」の判定と、本棚カードへ出すメモ。実機第2報 RC1(44ページの約款が
    # chunks=1 / status=done で登録成功になった)への二段目の防衛で、判定式を
    # modShelf 側に散らさないためここへ置く。IsThinExtract は純ロジックなので
    # modTestsPure9 が閾値を固定する(閾値は仕様R13-3aから動かさないこと)。
    # IsUnreadableCopyReason / CopyFailMsgFor(2026-08-03 R14-3b): 一時コピーの
    # 失敗理由が「元ファイルを読めていない」ことを意味するかの判定と、その
    # ときの利用者向けの1文。読めていないのに元パスのまま続行すると、NFD分解名が
    # そのままGhostscriptへ渡って /undefinedfilename になり、利用者には
    # 「描画0枚」しか残らない(実機第3報 RC3)。どちらも純ロジックなので
    # modTestsPure11 が境界を固定する。
    "modExtractorPdf": {
        "closed": True,
        # CopyReasonKind(2026-08-03 R14-F3/F11): 理由の1語 → 文言の種類の
        # 対応表(純ロジック)。modLog.CopyFailMsgOf と対で使う。
        # GcOldTempCopies(R14-F11): %TEMP%\mbtmp_* の24時間より古い残骸の
        # 掃除。呼ぶのは modBoot の起動GCから1行だけで、本体をこちらに置くのは
        # modBoot が28,000字のWARN帯に近いのと、一時名の規約(TempBaseNameFor)を
        # 持つのがこのモジュールだから。
        "required": ["ExtractPdfWithFallback", "CopyToLocalTemp",
                     "DropGarbledPages", "TempBaseNameFor",
                     "IsThinExtract", "ThinExtractMemoFor",
                     "IsUnreadableCopyReason", "CopyFailMsgFor",
                     "CopyReasonKind", "GcOldTempCopies"],
    },
    "modMode": {
        "closed": True,
        # 回答モード(すぐ聞く/通常/入念)の方針。純ロジックなのでLOで検証する。
        # AskStage*(2026-08-03 R13-9b): 質問処理の段階ナレーションの算数。
        #   「(2/4) 資料を照合中…」の番号は、その回に実際に通す段の数から作る
        #   (拡張・再ランクは構成で有無が変わるため、総数を4に固定すると嘘になる)。
        #   表示そのものは modAskRetrieve.ShowAskStage が行い、ここは純ロジック。
        # ShouldEmitInsight(2026-08-03 R14-1b): 「解決した」で部内へ発信して
        #   よい回答かの判定(モード×出典件数の真理表)。一般モードの解禁と
        #   残留状態による誤爆(実機第3報 RC2)を1つの式で塞ぐ。
        # RerankEffort(2026-08-03 R14-8a): 入念モードだけ再ランクの effort を
        #   別設定(rerank_effort_thorough)にする分岐。
        "required": ["Normalize", "NextMode", "Caption", "Description", "TopK",
                     "UseExpand", "UseRerank", "UseVerify", "UseLightExpand",
                     "SubQueryCount", "ShouldEmitInsight", "RerankEffort",
                     "AskStageTotal", "AskStageIndex", "AskStageLabel", "AskStageText"],
    },
    "modSparse": {
        "closed": True,
        # 日本語キーワード検索(文字bigram + BM25 + 完全一致)。
        # 全て純ロジックなので modTestsPure2 から直接検証する。
        # 実測: 旧実装 R@1 32% → 本実装 84%(tools/bench_retrieval.py)。
        # HasAnyKey(2026-08-01 R12-3-7): binary_rag の粗選別が「効く語の完全
        # 一致」を持つ行を先に落とす件の救済union用。採点(KeyScore)とは別に、
        # 生テキストへの粗い包含判定を1本だけ公開する(判定規則が2箇所に
        # 分かれないよう、modRetrieve 側には条件を書かない)。純関数なので
        # modTestsPure4 が境界を固定する。
        # MatchDocText(2026-08-01 R12-4): 1チャンクの照合テキストを作る式を
        # 1本にした。取込時に my_knowledge の norm_text 列へ前計算して保存する
        # ようになったため、式が2箇所にあると保存済み行と未保存行でスコアが
        # 変わる(順位が静かに割れる)。
        "required": ["NormalizeForSearch", "Tokenize", "DistinctiveKeys",
                     "Bm25Score", "ExactHitCount", "CompactForMatch", "KeyScore",
                     "HasAnyKey", "MatchDocText"],
    },
    "modChunker": {
        "closed": True,
        # ChunkPagesEx/ClassifyLine/BuildBreadcrumb: 構造認識チャンク化(設計書§B)。
        # ChunkPagesは後方互換(legacy)のまま不変。
        # NormalizeForIngest/JoinSplitNumbers/IsPageNumberLine: 取込正規化
        #   (2026-07-27追加)。実物の約款PDFで条見出しの20.2%を取りこぼしていた
        #   ため、全取込経路が最初に通す正規化として追加。純ロジックなので
        #   Publicにして modTestsPure から直接検証する(回帰を二度と許さない)。
        # SkippedPageCount: 2026-07-29 実機事故対応。1ページの実行時エラーで
        # 資料まるごとが取込失敗になっていたため、ページ単位で飛ばせるように
        # した。何ページ飛ばしたかを取込側が利用者へ伝えるためのゲッター。
        "required": ["ChunkPages", "ChunkPagesEx", "ClassifyLine", "BuildBreadcrumb",
                     "NormalizeForIngest", "JoinSplitNumbers", "IsPageNumberLine",
                     "SkippedPageCount"],
    },
    "modEmbed": {
        "closed": True,
        # MarkAllForReembed: 圧縮/次元変更後の全再embed導線(設計書§G-T10)。
        "required": ["EmbedPending", "PendingCount", "MarkAllForReembed"],
    },
    "modShelf": {
        "closed": True,
        # IsBusy: 2026-07-31 R7 B-2。抽出ループのDoEventsで発火したクリックを
        # 画面遷移側(modUiLock.BlockIfIngesting)が受け流すための取込中フラグ。
        # AddFilesViaDialog / AddFilesResult は 2026-07-31 R10d で modShelfBatch へ
        # 移設した(modShelfが30,000字上限まで残り440字になったため)。
        # ChunkKeyOrder(2026-08-03 R14-5a/5c・実機第3報 RC7): 拡張子別チャンク
        # 設定キーの組み立て(純ロジック)。ChunkParamFor自体はmodConfig.GetLong
        # へ依存しExcel無しでは検証できないため、フォールバック順序の判定材料
        # (キー名)だけを公開してmodTestsPure11から固定する。
        "required": ["IngestFile", "DeleteSource",
                     "SourceList", "TotalChunks", "IsBusy", "ChunkKeyOrder"],
    },
    # modShelfBatch(2026-07-31 R10d): modShelfから切り出した一括取込の
    # オーケストレーション。AddFilesResult: 2026-07-30 R2要件E。
    # AddFilesViaDialog(OnAction互換)は内部でこれを呼ぶ薄いラッパー。
    # ok/ng/capped件数・新規追加チャンク数・失敗理由の内訳を戻り値の文字列で
    # 返し、showMsgBox:=False で呼び出し元(modApp.OnAddDocs等)が独自に
    # 表示を担えるようにする。
    # IsBatchBusy(2026-07-31 R10c H2): バッチ取込全体を覆う再入ガードの参照口。
    # modShelf.IsBusy が自分の mIngesting と OR で見るためだけに公開している。
    # StageBanner(2026-08-03 R13-4a): 取込の段階表示を進捗バナーへ流す薄いAPI。
    # 「バナーが既に出ているときだけ更新する」という規約を1箇所に閉じ込める
    # ためのPublicで、silent(自動同期)経路では何も表示しない。
    "modShelfBatch": {
        "closed": True,
        # TouchBusy / LastBeat / GuardExpiredNow(2026-08-04 R15-1a・実機第4報
        # RC5): 4本の再入ガードが共有する「最後に生きていた時刻」の置き場。
        # 開始から30分で失効させていたため、85〜127分かかる取込の途中で
        # ボタンが全面解放され二重取込が構造的に可能だった。
        # SaveCheckpoint(R15-3a・RC9): 1ファイル取込ごと/同期末尾の中間保存。
        # OnCancelIngest / CancelRequested / ResetCancel(2026-08-04 R15-6a・
        # RC3): 進捗バナー脇の中断ボタンと、その印の読み・下ろし。85〜127分の
        # 取込を止める手段が1つも無く、Excelの強制終了しか無かった。
        # 印の実体をここに置くのは、取込の入口(AddFilesResult)で必ず
        # リセットできる場所がここだけだから。
        # ShowIngestBanner(2026-08-04 R15-FixA FA-6): 【取込経路の】進捗バナー。
        # 中断ボタンは modSkin.PaintProgress が常に描いていたため、中断の仕組みが
        # 無い処理(質問の準備・部門更新・Q&A読込・ナレッジ登録)のバナーにも
        # 生えて「押しても何も起きないボタン」になっていた。ここを通る3経路
        # (AddFilesResult / StageBanner / modShelfSync のファイルループ)だけが
        # 中断できるバナーを出す。modUIMain.ShowProgress を分岐させないのは、
        # あちらが30,000字上限まで残り206字で引数1つ足す余地も無いため。
        "required": ["AddFilesViaDialog", "AddFilesResult", "IsBatchBusy",
                     "StageBanner", "ShowIngestBanner", "TouchBusy", "LastBeat",
                     "GuardExpiredNow",
                     "SaveCheckpoint", "OnCancelIngest", "CancelRequested",
                     "ResetCancel"],
    },
    # 2026-07-28 レビューI-2対応でmodShelfから切り出したシート行操作層。
    # 取込フロー以外(同期・失効ワイプ)からも呼ぶ共通処理のため open。
    "modShelfStore": {
        "closed": False,
        "required": [
            "EnsureKnowledgeSheet", "EnsureManifestSheet", "BuildExistingHashSet",
            "RemoveKnowledgeAndVectorsForSource", "RemoveVectorsByIds",
            "RemoveManifestRowForSource", "UpsertManifestRow", "SliceRows",
        ],
    },
    # modChunkMetaStore(2026-08-05 R17波0→波1): chunk_metaシート(chunk_id/
    # section_path/refs_out)のEnsure/バッチ書込み/全行読み層。modShelfStore.
    # EnsureKnowledgeSheetと同型(無ければ作成+ヘッダ+veryHidden、既存なら冪等)。
    # WriteMetaFromRows(Phase1): 取込ループが持つmy_knowledgeの行列からchunk_idを
    #   取り出して1回で追記する入口。modShelf.IngestFileは残り字数が数百字しか
    #   ないため、id配列の組み立てをこちらで引き受ける(憲章§4-6)。
    # RemoveMetaForSource(Phase1): 再取込・資料削除で消えるmy_knowledge行の
    #   chunk_idを引き、chunk_metaの同じ行を落とす。必ず
    #   modShelfStore.RemoveKnowledgeAndVectorsForSourceの【前】に呼ぶ
    #   (後だと消えた行のchunk_idがどこにも残っていない)。
    "modChunkMetaStore": {
        "closed": True,
        "required": ["EnsureChunkMetaSheet", "WriteMetaRows", "ReadAllMeta",
                     "WriteMetaFromRows", "RemoveMetaForSource"],
    },
    # modChunkMeta(2026-08-05 R17 Phase1): チャンクの構造ラベル(section_path)と
    #   明示参照(refs_out)の抽出。VBScript.RegExp/ScriptControlを使わず
    #   modSparse.DistinctiveKeysと同型のInStr走査だけで書いた純ロジックで、
    #   PURE_LOGIC_MODULES にも登録する(シートI/Oは modChunkMetaStore、
    #   検索への合流は modAskFocus が持つ)。
    #   MetaOf は ExtractSectionPath/ExtractRefs をまとめて呼ぶ入口で、
    #   modShelf.IngestFile の残り字数(憲章§4-6)のために置いてある。
    #   GraphActive は「chunk_meta が無い/空なら全機能が従来動作」という
    #   フェイルセーフの条件を1本にしたもの(条件が2箇所に割れるのを防ぐ)。
    "modChunkMeta": {
        "closed": True,
        "required": ["ExtractSectionPath", "ExtractRefs", "MetaOf",
                     "PathHasLabel", "RefLabelsFor", "GraphActive"],
    },
    # modOutlineStore(2026-08-05 R17 Phase2): doc_outlineシート(source/
    # section_key/summary/keywords/chunk_n)のEnsure/一括書込み/全行読み/
    # 資料単位の掃除。modChunkMetaStore と同型で、違いは source 列を自分で
    # 持つこと=掃除に my_knowledge を引かないので、
    # RemoveKnowledgeAndVectorsForSource との前後関係の制約が無い。
    # veryHidden は EnsureOutlineSheet が毎回・冪等に自己設定する
    # (modBoot.HideInternalSheets は残8字で1行も入らないため。憲章§4-6)。
    "modOutlineStore": {
        "closed": True,
        "required": ["EnsureOutlineSheet", "WriteOutlineRows", "ReadOutline",
                     "RemoveOutlineForSource"],
    },
    # modOutlineBuild(2026-08-05 R17 Phase2): 取込時の章単位要約。
    # BuildOutlineFor: 1資料ぶんの章要約を作って doc_outline へ保存する唯一の
    #   入口。config graph_outline=off / chunk_meta 0行 / 章キーが取れない
    #   なら完全に無操作(LLMも呼ばない)。中断は章の境界で拾い、そこまでの章を
    #   保存して正常終了する。失敗章は「(要約失敗)」の行として保存し続行。
    # ChapterKeyOf / BudgetTake: 章グルーピングキーと本文の打ち切りの純ロジック。
    #   PURE_LOGIC_MODULES には載せない(シートI/O・CallLLM を持つため)が、
    #   run_lo_tests の PURE_ALLOWLIST へは載せてこの2本だけ実行テストで固定する
    #   (modAskFocus/modAskMulti と同じ型)。
    # SetUnattended(2026-08-05 R17H FA-7 / A-M8): 無人実行(自動同期)の印。
    #   自動同期は OnTime で勝手に始まり、進捗バナーも中断ボタンも利用者の目の
    #   前には無い。そこで章数の多い資料の章要約(章数ぶんのLLM呼び出し)を
    #   始めると「何もしていないのにExcelが数分固まる」だけになる。上げ下げは
    #   modShelfSync.SyncNow の入口/Finish の2行だけで、判断そのもの
    #   (章数>12 なら先送りして usage_log("outline_deferred"))はこの層に閉じる。
    "modOutlineBuild": {
        "closed": True,
        "required": ["BuildOutlineFor", "ChapterKeyOf", "BudgetTake", "SetUnattended"],
    },
    # modSynonymStore(2026-08-05 R17 Phase3): 用語の表記ゆれ辞書(synonyms:
    #   term/canonical)のEnsure/一括書込み/読み/全消去/名寄せバッチ。
    # EnsureSynonymSheet: modOutlineStore.EnsureOutlineSheet と同型(冪等)。
    # WriteSynonymRows: 末尾へ1回のRange書込みで追記するだけ(上書きの判断は
    #   持たない=BuildSynonymsForがRemoveAll+全件書き直しで実現する)。
    # ReadMapCsv: 質問側(modAskRetrieve)が読む唯一の窓口。全行を
    #   "term>canonical|term>canonical" の1文字列へ畳んで返す(0行は空文字)。
    #   modSparseと同じくPURE_LOGIC_MODULESには置けない(シートI/O)ため、
    #   質問文への展開そのものは純関数 modRagParse.ExpandQueryBySyn へ渡す。
    # RemoveAll: 全行消去(再構築用)。BuildSynonymsForが「既存termは上書き」を
    #   実現するために、ReadMapCsvで読んだ既存分から新規termと重複する行を
    #   除いた集合をRemoveAll後に丸ごと書き直す。
    # BuildSynonymsFor: 1資料ぶんの名寄せバッチの唯一の入口。config
    #   graph_synonyms=off / 用語候補0件 / LLM応答が空 のいずれかで完全に
    #   無操作。ゲート・失敗握りもこの層に閉じるので呼び出し元
    #   (modOutlineBuild.BuildOutlineFor)は1行。
    "modSynonymStore": {
        "closed": True,
        "required": ["EnsureSynonymSheet", "WriteSynonymRows", "ReadMapCsv",
                     "RemoveAll", "BuildSynonymsFor"],
    },
    # modShelfVision(2026-07-31 R6): 取込失敗時のvisionフォールバック集約。
    # modShelfの取込フローからこの判断を丸ごと引き受ける(公開はTryVisionFallback
    # の1本だけ。増えるならまず「本当にIngestFileから見える必要があるか」を疑う)。
    "modShelfVision": {
        "closed": True,
        "required": ["TryVisionFallback"],
    },
    "modShelfSync": {
        "closed": True,
        # DiffDecision は §7.8 の記述により modShelfSync の差分判定を
        # テスト可能にするため追加で Public 化することが名指しされている。
        # AutoSyncTick(Wave4修正・MASTER_SPEC §7.2契約更新): Application.OnTime の
        # Procedure引数はApplication.Runと同じ遅延バインドでありPrivate Subを
        # 解決できないため、OnTimeコールバック本体はPublicが必須(VBAの実際の
        # 挙動に合わせた契約修正。当初「契約に無い名前だから」とPrivateにしたのは
        # 誤りだった)。ResolveDecision(同じくWave4修正・§7.8): DiffDecisionに
        # manifestの現在statusによる上書きルール(failed/missing→replace)を適用する
        # 純関数で、DiffDecisionと同じ理由でテスト可能にするため公開契約に追加した。
        # IsBusy(2026-07-31 R7 B-2): modShelf.IsBusy と対。同期のファイルループも
        # DoEvents を回すため、同期中かどうかを画面側へ公開する。
        "required": [
            "PickShelfFolder", "SyncNow", "ScheduleAutoSync", "CancelAutoSync",
            "DiffDecision", "AutoSyncTick", "ResolveDecision", "IsBusy",
        ],
    },
    "modEnrich": {
        "closed": True,
        "required": ["EnrichPending"],
    },
    # ---- 7.3 QA層 ----
    "modRetrieve": {
        "closed": True,
        # SearchExpanded: マルチクエリ検索(設計書§C-2)。Searchは不変。
        "required": ["Search", "SearchExpanded"],
    },
    # modVecCache(2026-08-01 R12-4): セッション内ベクトルキャッシュと埋め込み
    # 世代カウンタ。my_vectors の再パースを1セッション1回に畳む。open にして
    # あるのは、参照系アクセサ(SlotOfRow/DimOfRow/RowOfSlot/CachedDim…)が
    # 実装都合で増減するため。要件で名指しされている「世代」「解放」「構築」
    # 「等価な内積」の4点だけを契約として固定する。
    "modVecCache": {
        "closed": False,
        "required": [
            "Generation", "BumpGeneration", "ResetVecCache",
            "BuildFrom", "PrepareVectors", "DotAt", "SlotOfRow",
            "StampOf", "IsStale", "Ready",
        ],
    },
    "modPrompts": {
        "closed": True,
        # BuildExpandPrompt/BuildRerankPrompt: 多段RAGの拡張・再ランク段(設計書§C)。
        # BuildSourceDigestPrompt/BuildCritiquePrompt(2026-08-03 R14-8a):
        #   入念モードの(1)資料の要点整理と(3)自己批判。
        # SourceTag(同): 出典タグの形の単一情報源。LLMへ指示する形と
        #   modAskThorough が突合に使う形が別実装になると、検査が全件一致か
        #   全件不一致のどちらかへ静かに退化する(憲章§4-5)。
        # BuildQuestionsPrompt(2026-08-03 R14-7a): 質問例のオンデマンド生成。
        #   シード0でも資料があれば、資料名+冒頭抜粋から短い質問を5件作らせる。
        # BuildDecomposePrompt/BuildPartDraftPrompt/BuildMergePrompt
        #   (2026-08-05 R16-3A): 複合質問の分解(段0の判定)・論点1つぶんの副下書き・
        #   論点ごとの下書きの統合。PART_FAIL_TEXT は「その論点は資料を確認でき
        #   なかった」節の文言で、節を作る側(modAskMulti.BuildPartSection)と
        #   それを消させない側(BuildMergePrompt)と成否を数える側(modAskMulti)の
        #   単一情報源(2実装に分かれた瞬間、統合段が黙って言い換える)。
        "required": ["BuildQuickPrompt", "BuildDeepDraftPrompt", "BuildDeepVerifyPrompt", "BuildEnrichPrompt", "BuildExpandPrompt", "BuildRerankPrompt",
                     "BuildSourceDigestPrompt", "BuildCritiquePrompt", "SourceTag", "BuildQuestionsPrompt",
                     "BuildDecomposePrompt", "BuildPartDraftPrompt",
                     "BuildMergePrompt", "PART_FAIL_TEXT"],
    },
    "modRagParse": {
        "closed": True,
        # 多段RAGのLLM応答パーサ(設計書§C/§D/§G-T7)。全て寛容退化の純関数。
        # ParseQuestionLines(2026-08-03 R14-7a): 質問例オンデマンド生成の応答
        #   パーサ(1行1問・番号無しへ寛容退化)。RAG本体とは無関係だが、
        #   「LLM応答をパースする純関数」という性質はここと同じなので同居させる。
        # IsErrorResponse/BuildErrorAnswer(2026-08-03 R14-G1): modGateway.CallLLM の
        #   失敗は "#ERR:コード:説明" という応答文字列で返る契約なので、その判定と
        #   利用者向け文言化も「LLM応答のパース」。modAsk が30,000字上限まで残り44字に
        #   なったため(憲章§4-6)、3モジュール(modAsk/modAskRetrieve/modAskThorough)が
        #   共有していたこの3本を副作用ゼロのまま本モジュールへ移設した。
        # ParseDecomposeVerdict/ParseParts(2026-08-05 R16-3A): 複合質問の分解
        #   (段0)の応答パーサ。<verdict>single|parts|clarify</verdict> と
        #   <parts>a|b|c</parts> を読む。タグ欠落・不正な語・"#ERR:" はすべて
        #   single へ寛容退化し、呼び出し元が従来の入念フローへ無害に落ちる。
        # ParseOptions/ParseChoiceNumbers(2026-08-05 R16-3B): 逆質問=番号選択肢。
        #   ParseOptions は段0の <options> を選択肢配列へ(ParseParts と同型)。
        #   ParseChoiceNumbers は「1」「1と3」「①と③」のような返事から番号を
        #   読む唯一の場所で、数字と区切り以外の文字が1つでも混じれば空を返す
        #   (=書き直し扱い)。ここが甘いと、利用者が打ち直した質問文が黙って
        #   捨てられる(2026-07-28 レビュー H-12 と同型の事故)。
        # HasCompoundSignal(2026-08-05 R18-7a・実機第5報⑦): 質問文そのものの
        #   複合シグナル検知(？の2個以上出現/。区切りの非空節2個以上)。LLM応答の
        #   パースではないが「質問文の中身を見る文字列パターン検知」という性質は
        #   ParseChoiceNumbers等と同じで、modMode(モード名だけを見る)には置けない。
        # ParseOutlineResp/ParseChapterPick(2026-08-05 R17 Phase2): 章単位要約の
        #   応答パーサ。<summary>/<keywords> は取込時(章ごとに1回)、<pick> は
        #   俯瞰質問の章選択(質問ごとに1回)。どちらも読めなければ「その章は
        #   要約失敗」「章は選ばれなかった=従来フローへ」へ寛容退化する。
        # ParseSynResp(2026-08-05 R17 Phase3): 名寄せ(表記ゆれ)応答のパーサ。
        #   出力契約 <syn>表記>正規形|表記>正規形</syn>。">"の無い/どちらか空の
        #   ペアは1件ずつ破棄する(全滅ではなく読めた分だけ返す寛容退化)。
        # ExpandQueryBySyn(2026-08-05 R17 Phase3): synonymsシートの内容
        #   (modSynonymStore.ReadMapCsvが返す1文字列)を引数で受け、質問文中の
        #   語に一致した同義語を最大maxAdd件・半角空白区切りで追記する純関数。
        #   modSparseはPURE_LOGIC_MODULESのままシートI/Oを持てない(設計書§3
        #   Phase3・調査agent7 §3.3)ため、シートを読む側(modAskRetrieve)が
        #   文字列化した地図を渡す構成にした。一致判定はmodSparse.
        #   NormalizeForSearchを両辺に通すので全角/半角の表記ゆれも吸収する。
        # MergeSynPairs(2026-08-05 R17H FA-1 / A-H1): 旧synonyms CSV+新CSV→
        #   統合CSV のマージ規則。実体は modSynonymStore.MergeAndSave にあり、
        #   ParseSynResp が返す【0始まり】の配列を 1〜n で読んでいたため実データ
        #   では毎回「添字が範囲外」で握り潰され、synonyms が永久に0行だった。
        #   配列の起点に依存しない「文字列→文字列」の純関数へ切り出し、
        #   modTestsPure17 のゴールデンで固定する(Store側は薄く呼ぶだけ)。
        # HasGlobalSignal(2026-08-05 R17H FA-6 / A-M7・B-H2): 俯瞰(全体像・
        #   一覧)を求める語彙シグナル。「全体像は?」のような短文が段0判定を
        #   呼ばれず IsTooVague の聞き返しにも吸われて、俯瞰が一度も試されない
        #   狭間を塞ぐ。HasCompoundSignal と同じ「質問文の文字列パターン検知」。
        "required": ["ParseExpand", "ParseRankOrder", "ExtractAnswer", "ParseSubqueries",
                     "ParseQuestionLines", "IsErrorResponse", "BuildErrorAnswer",
                     "ParseDecomposeVerdict", "ParseParts",
                     "ParseOptions", "ParseChoiceNumbers", "HasCompoundSignal",
                     "ParseOutlineResp", "ParseChapterPick",
                     "ParseSynResp", "ExpandQueryBySyn",
                     "MergeSynPairs", "HasGlobalSignal"],
    },
    "modAsk": {
        "closed": True,
        # CanFollowup/AskFollowup: 続けて質問=深掘り機能(裁定D11・RIBBON_API_CONFIRMED.md §2b)
        # LastAnswerText: Nexus UI(modApp)が直近回答をバブル表示するためのゲッター。
        # LastHit*: Peek View(出典ポップアップ)用の読み取り専用アクセサ。直近回答が
        #   根拠にした出典(source/page/origin/本文)をUI層(modPeek)へ公開する。0始まり。
        # LastConfidence/LastConfidenceText: 回答の信頼度(2/1/0)とその説明文。
        #   検索スコアという機械側の情報を人間が判断に使える形でUI層へ渡す
        #   (2026-07-26: 利用者が回答の正誤を判断できずフィードバックが集まらない
        #   という課題への対処。共有知フライホイールの入口)。
        # FeedbackUnsure: 🤔微妙(入力を求めない1クリック評価)。
        "required": ["AskFromUI", "Answer", "FeedbackGreen", "FeedbackYellow", "FeedbackRed",
                     "FeedbackUnsure", "CanFollowup", "AskFollowup", "LastAnswerText",
                     "LastTopSource", "LastConfidence", "LastConfidenceText",
                     "LastHitCount", "LastHitSource", "LastHitPage", "LastHitOrigin", "LastHitPeek",
                     # 2026-07-28: modAskRetrieve への切り出し(レビューI-2)に伴い公開。
                     # HistoryBlock=拡張プロンプトに載せる直近履歴、
                     # IsErrorResponse=#ERR:応答を検索途中で捨てる判定。
                     # NoteGeneralAnswered(2026-08-03 R14-1b): 一般アシスタントで
                     #   答えたことを「直近の回答」として記録する唯一の窓口。
                     #   これが無いと一般モードの「解決した」が、前のRAG質問の
                     #   回答を解決したことにされる(実機第3報 RC2)。
                     # ApplyAnswerTags(2026-08-03 R14-8a): answer_tags時の本文抽出。
                     #   入念モードの各段(modAskThorough)も同じ規約で取り出す必要が
                     #   あるため公開した(2実装に分かれると片方だけ<thinking>が漏れる)。
                     # IsErrorResponse は 2026-08-03 R14-G1 で modRagParse へ移設
                     #   (#ERR: 応答のパースであり、容量的にもmodAskに置く理由が無い)。
                     # CanShareInsight(2026-08-03 R14-G1): 直近回答を部内へ発信して
                     #   よいか。判定材料(モード・出典件数)はmodAskのモジュール変数
                     #   にしか無いので、UI層(modAppAct)の訂正共有もこの窓口を通す。
                     # NoteAnswerFailed(2026-08-03 R14-G2): 回答を作れなかったターン。
                     #   これが無いと、一般モードの失敗直後の感想ボタンが【前の質問】
                     #   の状態で受理される。
                     "HistoryBlock", "CanShareInsight", "NoteAnswerFailed",
                     "NoteGeneralAnswered", "ApplyAnswerTags"],
    },
    # 2026-07-28 レビューI-2対応でmodAskから切り出した検索層。
    # modAskのモジュール変数を触らず、引数のhits()だけで完結する。
    # RunDeepScoped/RunUnscoped(2026-08-03 R13-5c): 深掘り(followup かつ deep)の
    #   スコープ内多段検索と、その退避先である従来経路。modAsk はこの2本しか
    #   呼ばない(retrieve_mode の判断を2箇所に書かない)。
    # PlanAskStages/ShowAskStage(R13-9b): 質問処理の段階ナレーション。
    #   何段通すかを先に決めてから「(2/4) 資料を照合中…」を出す。
    "modAskRetrieve": {
        "closed": True,
        "required": ["RunMultiRetrieve", "ApplyLowHitWarning", "IsTooVague", "HitSourceList",
                     "RunDeepScoped", "RunUnscoped", "PlanAskStages", "ShowAskStage"],
    },
    # modAskThorough(2026-08-03 R14-8a): 「入念に調べる」専用の生成パイプライン。
    #   実機第3報 RC8「deep と thorough が生成側で完全に同じ(下書き・検証の
    #   プロンプトもモデルも effort も共有)」への対処で、入念だけを
    #   要点整理→下書き→自己批判→検証→出典の機械的突合 の5段にする。
    #   quick/deep の経路は1行も変えない。
    #   RunThoroughFlow が唯一の入口(modAsk から1箇所だけ呼ばれる)。
    #   残りは最後の出典突合の部品で、LLMを一切使わない純ロジック。
    #   AnnotateAgainstHits がログ付きの本体、CiteIndexFrom/NormalizeCiteTag/
    #   ExtractCiteTags/IsCiteTag/TagIsKnown/AnnotateCitations は
    #   modTestsPure11 が真理表で固定する(run_lo_tests.py の PURE_ALLOWLIST
    #   にも登録済み。未登録だとテストが実行時エラー12で走らない)。
    #   UNVERIFIED_MARK は付記の文言で、テスト側と表示側の単一情報源。
    # modAskMulti(2026-08-05 R16-3A): 複合質問の分解 → 論点ごとの調査 → 統合。
    #   入念モードだけ段0(分解判定)を1回足し、verdict=parts のときに論点ごとの
    #   検索+副下書きを回して1本へ統合してから、自己点検・検証・出典突合へ渡す。
    #   TryDecomposed が唯一の入口(modAsk の thorough 分岐から1箇所だけ)で、
    #   False を返したら呼び出し元は従来どおり modAskThorough.RunThoroughFlow を
    #   実行する=分解は上積みであって置き換えではない。
    #   VerifyNote は modAsk が検証注記を取る【唯一の窓口】。分解経路のターンは
    #   自前の注記を、通らなかったターンは modAskThorough の注記を中継する
    #   (窓口が2つあると、分解したターンに前回の single 経路の注記が付く)。
    #   StepNote は usage_log の ask_steps へ足す "dec=論点数"(読んだら消える)。
    #   ShouldDecompose/PerPartTopK/BuildPartSection は発動条件・論点あたりtopK・
    #   統合入力の1節の組み立て(部分失敗の文言込み)という純ロジックで、
    #   modTestsPure14/15 が真理表で固定する(run_lo_tests.py の PURE_ALLOWLIST にも
    #   登録済み。未登録だとテストが実行時エラー12で走らない)。
    #   ShouldClarify/BuildClarifyAsk(2026-08-05 R16-3B): 読み方が定まらない質問へ
    #   返す番号選択肢の逆質問。発動条件(clarify_mode と選択肢2件以上)と本文の
    #   組み立てだけで、LLMは1回も呼ばない(段0の判定でもう材料が揃っている)。
    "modAskMulti": {
        "closed": True,
        # DecomposeGate(2026-08-05 R18H FB-2 / A-L9): 段0を呼ぶかどうかの最終
        #   判定(ShouldDecompose OR modRagParse.HasCompoundSignal)。実装と
        #   modTestsPure16 が同じ1本を呼ぶために Public 化した(従来はテストが
        #   同じ式を書き写しており、片方だけ直しても検知できなかった)。
        "required": ["TryDecomposed", "VerifyNote", "StepNote",
                     "ShouldDecompose", "DecomposeGate",
                     "PerPartTopK", "BuildPartSection",
                     "ShouldClarify", "BuildClarifyAsk"],
    },
    # modAskFocus(2026-08-05 R16-3C): 精読=根拠チャンクの前後を一緒に読む。
    #   NeighborExpand が唯一の入口(modAskThorough / modAskMulti の2箇所から、
    #   いずれも「材料が確定した直後に1回」だけ呼ばれる。R16H FA-4 で
    #   modAskRetrieve.FinishDeep からの呼び出しは外した=deep には効かない)。
    #   ParseChunkKey/NeighborIdList は chunk_id からの文書順復元と前後取りの
    #   純ロジックで、modTestsPure15 が真理表で固定する(run_lo_tests.py の
    #   PURE_ALLOWLIST にも登録済み。未登録だとテストが実行時エラー12で走らない)。
    #   PURE_LOGIC_MODULES には載せない(my_knowledge を直接 Range 読みするため)。
    #   RefsExpand / ArticleEnsure(2026-08-05 R17 Phase1): 物理近傍ではなく
    #   【意味の上で繋がっているチャンク】を足す2本。RefsExpand は根拠チャンクの
    #   refs_out を1ホップ展開して同じ資料の中から参照先を束ね(modAskThorough /
    #   modAskMulti の NeighborExpand 直後から1行ずつ)、ArticleEnsure は質問が
    #   名指しした条番号のチャンクが1件も無いときだけ先頭へ入れる
    #   (modAskRetrieve.RunMultiRetrieve から)。config graph_refs のゲートと
    #   「chunk_meta が無ければ無操作」の判定はこの層に閉じるので、呼び出し元は
    #   どれも1行のまま=検索の当て方は1文字も変わらない。
    "modAskFocus": {
        "closed": True,
        "required": ["NeighborExpand", "ParseChunkKey", "NeighborIdList",
                     "RefsExpand", "ArticleEnsure"],
    },
    # modAskGlobal(2026-08-05 R17 Phase2): 俯瞰質問(疑似グローバル検索)。
    # TryGlobal: doc_outline の章要約を1回のプロンプトへ載せて読むべき章を
    #   選ばせ(chapter_pick)、選ばれた章の本文を文書順に集めて1回で回答を
    #   作る(global_answer)。追加のLLM呼び出しは1質問あたり2回。
    #   不発(graph_outline=off / doc_outline 0行 / 章が選ばれない / 章の
    #   チャンクが引けない / 回答生成の失敗)は全て False で、呼び出し元
    #   (modAskMulti)は従来の入念フローへ落ちる=フェイルセーフはこの層に閉じる。
    # OutlineActive: そのフェイルセーフの単一情報源(0行なら俯瞰は動かない)。
    # WasGlobalTurn/ResetGlobalTurn(2026-08-05 R17H FA-2 / A-H2・B-H1):
    #   直近ターンが俯瞰だったかの1ビット。俯瞰の hits は score=0(章の要約から
    #   選んだもので検索スコアではない)ため、低関連度の警告と信頼度バッジが
    #   同時に付いて【正しく答えた回答】が二重に否定されて見えていた。点数は
    #   変えず表示だけを出自どおりに直すために、表示側(modAskRetrieve /
    #   modUINexusDraw)へこの1ビットだけを公開する。リセットは
    #   modAskMulti.TryDecomposed の入口1箇所(印を次のターンへ残さない)。
    "modAskGlobal": {
        "closed": True,
        "required": ["TryGlobal", "OutlineActive", "WasGlobalTurn", "ResetGlobalTurn"],
    },
    "modAskThorough": {
        "closed": True,
        # VerifyNote(2026-08-03 R14-G11): 検証段が落ちたターンの内部注記。
        #   本文へ混ぜると mLastCleanAnswer(=会話履歴と「解決済みQ&A」の部内
        #   共有)にまで注記が入るため、modAsk が整形の後に足せるよう別で返す
        #   (deep の RunDeepFlow と同じ並びに揃えた)。
        "required": ["UNVERIFIED_MARK", "RunThoroughFlow", "AnnotateAgainstHits",
                     "CiteIndexFrom", "NormalizeCiteTag", "ExtractCiteTags",
                     "IsCiteTag", "TagIsKnown", "AnnotateCitations", "VerifyNote"],
    },
    # ---- 7.4 パック層 ----
    "modPii": {
        "closed": True,
        "required": ["ScanText"],
    },
    "modPack": {
        # §7.8 が「ValidatePackの列検査部を純関数に切り出してテスト可能に
        # する」ことを求めているが名前が未確定のため open にして、名前不明の
        # 追加Publicまでは許容する(3つの契約関数の欠落だけは検出する)。
        "closed": False,
        "required": ["ImportPackDialog", "ValidatePack"],
    },
    # 2026-07-28 レビューI-2対応でmodPackから切り出した発行(書き出し)側。
    "modPackExport": {
        "closed": False,
        "required": ["ExportPackDialog", "ExportPackToFile"],
    },
    # ---- 7.5 統計層 ----
    "modStats": {
        "closed": True,
        # EXP/レベル(ゲーミフィケーション): AddExp=加点イベント口、ExpTotal/Level/
        # LevelProgress/ExpFloorForLevel はダッシュボードが読む集計ゲッター。
        # GetStatText/SetStatText: 部門チャンネルの取り込み済み版番号など、
        #   数値でない状態を持つための文字列アクセサ(2026-07-26)。
        "required": ["GetStatText", "SetStatText", "Bump", "GetStat", "TouchToday", "EvaluateBadges", "SavedMinutesEstimate",
                     "AddExp", "ExpTotal", "Level", "ExpFloorForLevel", "LevelProgress",
                     "ReportNoise", "NoiseThreshold", "ExcludedSources",
                     "MarkGlobalExcluded", "ResetGlobalExcluded",
                     "IsGloballyExcluded", "GlobalExcludedSources",
                     # 2026-07-28(解説書 §11-11): バッジ表の単一情報源。
                     # 判定(EvaluateBadges)と表示(modDash/modHub)で表が
                     # 二重化していたため、獲得しても表示されないバッジが
                     # 4種あった。判定を持つ側が名前も持ち、表示は読むだけにする。
                     "BadgeCatalog", "BadgeEarnedOn",
                     # 2026-07-31(R11-H Med3): 起動中のバッジ獲得告知を積んで
                     # 画面確定後にまとめて1本出すための遅延キュー。
                     "BeginDeferredBadges", "FlushBadgeToasts",
                     # AskTotalAll(2026-08-03 R14-1a): 質問回数の合算の単一情報源。
                     # quick+deep 決め打ちが3箇所にあり、入念モードの質問が
                     # どこにも出てこなかった(実機第3報 RC1)。
                     "AskTotalAll"],
    },
    # ---- 7.6 UI層 ----
    "modUIMain": {
        # opt機能ボタンのラッパー(OnTtsButton等、§7.7末尾)は名前未確定で
        # モジュール分担も未確定のため open。契約に載っている10個の欠落だけ
        # 検出する。
        "closed": False,
        "required": [
            "EnsureLayout", "SetStage", "RenderAnswer", "RenderSourcesPreview",
            "OnAskButton", "OnModeQuick", "OnModeDeep", "OnOpenHowto",
            "OnRunDiag", "ShowTip",
        ],
    },
    "modUIShelf": {
        "closed": False,  # 同上(opt機能ボタンラッパーを許容)
        # 2026-07-31(R11-E H-5): OnAddFiles/OnSyncNow/OnPickFolder/OnExportPack/
        # OnImportPack/OnBackToChat+AddButtonは、どのShapeのOnActionからも
        # 参照されない死にコードだったため削除した(実装は modKnowledge 側の
        # 同名+同等ハンドラに一本化済み)。契約の必須名からも外す。
        "required": [
            "EnsureLayout", "RenderShelf", "OnDeleteSource",
        ],
    },
    # modChrome(2026-07-30 R4要件C/D): ツールバーとヘッダーピルの配置計算だけを
    # 集めた純ロジック(R4準拠)。「ボタンが見切れる」「タイトルに重なる」は
    # どちらも算数の誤りだったので、算数だけを実行テストで固定できる形にした。
    # 表示側の都合で補助関数が増えうるため open。
    "modChrome": {
        "closed": False,
        "required": [
            "SumSpan", "FlowLeft", "FlowRight", "TitleReserve",
            "TextSpan", "ClipToWidth", "PillWidth",
        ],
    },
    # modShareRule(2026-07-31 R8): P2P/共有系の判定式だけを集めた純ロジック。
    # WIP段階のため open(補助関数が増えうる)。
    "modShareRule": {
        "closed": False,
        "required": [
            "OriginKind", "AuthorStatKey", "ProbeIsReachable",
            "ExpiryDecision", "StandardSubDirs",
        ],
    },
    "modBoot": {
        "closed": True,
        # RunFirstRunPromptEarly(2026-07-21追加): 自己インストーラの
        # Workbook_OpenがOnTime予約前に同期呼び出しする外部入口。Bootと同じ
        # 理由(自己インストーラ文字列からApplication.Runで呼ばれる)で必要。
        "required": ["Boot", "Auto_Open", "Auto_Close", "RunFirstRunPromptEarly"],
    },
    "ThisWorkbook": {
        # Workbook_Open/Workbook_BeforeClose は薄い転送のPrivateイベント
        # ハンドラのみで、Publicは想定されていない(§7.6末尾)。
        "closed": True,
        "required": [],
    },
    # ---- 7.7 opt層(全モジュール共通でPing必須) ----
    "optTts": {"closed": True, "required": ["Ping", "SpeakAnswer"]},
    # HasClipboardImage/SaveClipboardImage: スクショ取込(裁定D13・RIBBON_API_CONFIRMED.md §2b)
    # ExtractPdfOcrPagedText: 画像PDF(E0303)のOCR取込(2026-07-31 R6)。Ghostscriptで
    # ページ毎にJPEG化→1ページ=1回のChatGPTV→modUtil.JoinPagedTextで1本の文字列。
    # ResetGsGuidance(2026-07-31 R10-2): GS未検出の案内カード(セッション1回きり)
    # を再提示可能に戻す。modShelf.AddFilesResultからmodFeatures.InvokeFeature
    # 経由で呼ばれる(実機初報Aの観測性・堅牢化対応)。
    # ExtractPdfTextNoOcr(2026-07-31 R10-3): テキストPDFのCOM無し抽出の受け口。
    # 実体はoptGsTxt(optVisionに容量が無いための分割)で、ここは
    # modFeatures.InvokeFeature("vision", ...) の行き先がoptVision固定である
    # ことに合わせた薄い転送。
    # FindGsExeByCandidates/PathExists は optGsTxt へ貸すためのPublic
    # (R10-3 / R10-3b)。前者は案内カードを出さない「静かなGS解決」、後者は
    # Dir$による実在確認。コア側からは呼ばない(opt層内の参照はR2の対象外)。
    # OcrCapMemo(2026-08-03 R14-4c): OCRが上限ページで打ち切られたことを
    # 本棚カードのメモへ渡す唯一の経路(コア層はoptモジュール名を書けないため、
    # modFeatures.InvokeFeature("vision","OcrCapMemo") から呼ぶ受け口が要る)。
    # IsVisionError / SafeResultToString(R14-4a): ページOCRのループを
    # optOcrPage へ移した際、同じ判定を2箇所に持たないためPublic化した
    # (opt層内の参照なのでR2に触れない。憲章§4-5)。
    # OcrConfirmAsk / OcrDeclineMemo(2026-08-04 R15-7b): 何時間もかかる資料の
    # 事前確認。判断(頁数×レート≧ocr_confirm_min_minutes)と文面はopt側の
    # 純ロジックが持ち、聞く場所(MsgBox)はコア層(modShelfVision)にある。
    # コア層はoptモジュール名を書けない(R2)ので、この2本がその窓口になる。
    # OcrCacheGc(R15-7d): 頁キャッシュの孤児行の起動時GC。modBootの既存GC群
    # (nxocr_*/mbtmp_*)と同じ線で1回だけ呼ばれる。
    "optVision": {"closed": True, "required": ["Ping", "ExtractImagePdf", "ExtractImagePdfText",
                                               "ExtractPdfOcrPagedText", "OcrCapMemo",
                                               "HasClipboardImage", "SaveClipboardImage",
                                               "ResetGsGuidance", "ExtractPdfTextNoOcr",
                                               "FindGsExeByCandidates", "PathExists",
                                               "IsVisionError", "SafeResultToString",
                                               "OcrConfirmAsk", "OcrDeclineMemo",
                                               # OcrCachePurge(R15-FixB FB-1):
                                               # done で確定した資料の頁控えを
                                               # 消す受け口。コア層は opt名を
                                               # 書けない(R2)ので、InvokeFeature
                                               # から呼べる窓口がここに要る。
                                               "OcrCacheGc", "OcrCachePurge"]},
    # optOcrPage(2026-08-03 R14-4a): 画像PDFのページ描画とOCRの実行ループ。
    # 20ページずつGSを起動し、そのバッチをOCRし終えたら即座にJPEGを消す
    # (上限100ページでも一時領域は20枚ぶんで頭打ち)。各ページの
    # TryRibbonRun の前後で DoEvents を回し、「1ページ分より長くは固まらない」
    # 構造にする(同期Application.Runのハード中断は構造上不可能=制約)。
    # optVision が28,000字のWARN帯に達したための容量分割でもある(憲章§4-6)。
    # 純ロジックではない(Shell起動・待ち・ファイル削除・ログ)ので
    # PURE_LOGIC_MODULES には載せない。opt層に置く以上 Ping を持たせる。
    "optOcrPage": {"closed": True, "required": ["Ping", "OcrPdfByBatch"]},
    # optGsTxt(2026-07-31 R10-3 / R10-3b): Ghostscript実行の共通道具
    # (MakeOcrFolder/RunGsAsync/WaitForDoneFlag/CleanupOcrFolder。R10-3bで
    # optVisionから移設)と、txtwriteによるテキストPDF抽出(PDF本文抽出の
    # 第1選択。Word/AcrobatのOLE待ち回避)。opt層に置く以上、他のoptと同様に
    # Ping を持たせる。純ロジックではない(Shell起動・ファイルI/O・ログ)ので
    # PURE_LOGIC_MODULES には載せない。
    # GsFailureDetail(2026-07-31 R11-D・監査3 H-2): 完了フラグの中身(GSの
    # 終了コード)と gs_out.log の末尾を読み、err_log の detail にそのまま
    # 入れられる1本の文字列にする。optOcrPage(OCR経路)からも呼ぶ。
    # GsExitCode は 2026-08-03 R14-F13 で削除(終了コードは待ちループと
    # GsFailureDetail が読んでおり、外からの呼び出しが1件も無かった)。
    # LastGsTotalPages(2026-08-04 R15-5a・実機第4報 RC2): 直近の txtwrite 実行で
    # 判明した総ページ数。GSは最初に "Processing pages 1 through N" を出すので、
    # OCRへ回る資料の総頁は【OCRを始める前に】既に分かっている。従来これを
    # 誰も拾わず、OCR側は最終バッチに入るまで分母もETAも出せなかった。
    "optGsTxt": {"closed": True, "required": ["Ping", "ExtractPdfTextNoOcr",
                                              "MakeOcrFolder",
                                              "WaitForDoneFlag", "CleanupOcrFolder",
                                              "GsFailureDetail", "LastGsTotalPages"]},
    # optGsProc(2026-08-03 R13-F2): Ghostscriptプロセスの起動と停止だけを
    # optGsTxt から切り出したもの。PID再利用よけの本人確認(WMI Win32_Process
    # の名前照合)と rc=0/PID不明の扱いを足した結果 optGsTxt が28,000字の
    # WARN帯へ入ったための容量分割(憲章§4-6)。RunGsAsync は optGsTxt と
    # optVision の両方から、KillGsTree は optGsTxt の待ちループから呼ばれる。
    # 純ロジックではない(WMI・Shell起動・ログ)ので PURE_LOGIC_MODULES には
    # 載せない。opt層に置く以上、他のoptと同様に Ping を持たせる。
    "optGsProc": {"closed": True, "required": ["Ping", "RunGsAsync", "KillGsTree"]},
    # OpenAnswerInWord: 確定関数OpenWordMarkのラッパー(裁定D6)。
    # ExportAnswerAsDoc: 対話型Word文書生成(裁定D12・指示文→LLM整形→OpenWordMark)
    "optMarkdown": {"closed": True, "required": ["Ping", "RenderMarkdownAt", "OpenAnswerInWord",
                                                 "ExportAnswerAsDoc"]},
    "optDiffDoc": {"closed": True, "required": ["Ping", "CompareTwoDocsDialog"]},
    # optOcrCore(2026-07-31 R6): 画像PDFのOCR取込で使う「文字列の組み立てだけ」を
    # 集めた純ロジック(R4準拠)。GSコマンドの引用符の付け方が最大の地雷なので、
    # そこだけを副作用ゼロで切り出し modTestsPure4 のゴールデンテストで固定する。
    # modFeatures経由では呼ばれない(optVision専用の部品)が、opt層に置く以上
    # 他のopt同様に Ping を持たせる。
    "optOcrCore": {
        "closed": True,
        "required": [
            "Ping", "BuildGsCommand", "BuildRunCommand", "TempFolderFor",
            "OutPatternFor", "DoneFlagFor", "PageJpgName", "RenderCapFor",
            "IsTruncatedCount", "KeepPageCount", "SafeDpi", "SafeMaxPages",
            # R9: Ghostscript実行ファイルの解決候補列挙(配布・自動検出)。
            "GsCandidatePaths", "GsCandidatesForFolder",
            # R10-3: txtwriteによるテキストPDF抽出のコマンド組み立てと採否判定。
            # R10c: 採否判定を3値化(image/sparse/ok)し、その材料である
            # 「空白類を除いた文字数」と「実ページ数」も純ロジックとしてここへ集めた
            # (CleanTextLenはR10cでoptGsTxtから移設)。
            "BuildGsTextCommand", "GsTextVerdict",
            "CleanTextLen", "GsPageCount",
            # R11-D(監査3 H-2): GS実行の観測性。標準出力/標準エラーの落とし先
            # (GsLogFor)と、完了フラグに書かせた終了コードの読み取り
            # (GsExitCodeFromFlag)。どちらも純粋な文字列処理。
            "GsLogFor", "GsExitCodeFromFlag",
            # R13-1b/1c/1d(2026-08-03 実機第2報): GS待ちと失敗分類の判断材料を
            # 副作用ゼロで切り出したもの。ClassifyGsTextResult は「完了したのに
            # 本文が空」の本当の理由(image/gsfail/flagdelay)を決める真理表で、
            # ここを間違えるとスキャンPDFがWordのゴミ本文で「登録成功」になる
            # (RC1の事故そのもの)。GsTotalPagesFromLog/GsPagesFromLog は
            # gs_out.log から総ページ数と到達ページを読む進捗監視の目、
            # GsWaitBanner はその進捗の見せ方。全てLOテストで固定する。
            "ClassifyGsTextResult", "GsTotalPagesFromLog", "GsPagesFromLog",
            "GsWaitBanner",
            # R14-4a/4c(2026-08-03 実機第3報): ページ描画をバッチへ割る算数
            # (BatchCountFor/BatchBoundsFor)、OCR中の進捗バナーの文面
            # (OcrPageBanner)、上限打ち切りの正直なメモ(OcrCapMemoFor)。
            # 上限を100ページへ上げるにあたって足したものは全て副作用ゼロで、
            # 「20ページずつ描く」「総ページ数が不明なら数字を言わない」という
            # 判断をLOテストで固定する(RC4の嘘の案内を二度と作らないため)。
            # 2026-08-04(R15-5b): OcrPageBanner / OcrCapMemoFor /
            # OcrAbortMemoFor / RemainingWaitSec は optOcrEta へ移設した
            # (本モジュールが30,000字上限まで残り2,068字となり、R15-5のETAと
            # R15-6の中断メモが入らなくなったため。憲章§4-6)。ここは純減のみ。
            "BatchCountFor", "BatchBoundsFor",
        ],
    },
    # optOcrEta(2026-08-04 R15-5b): OCRの「進捗の見せ方」と「打ち切りの
    # 言い方」だけを集めた純ロジック。optOcrCore からの移設先で、あちらは
    # GSコマンドの組み立てとページ上限の算数(=Ghostscriptの都合)に専念する。
    # ETAと終了目安(RemainingText)、中断・頁欠けの正直なメモは、間違えると
    # 「いつ終わるか分からない」「頁が黙って欠ける」という実機第4報 RC2/RC6 の
    # 事故そのものへ戻るので、全てLOテスト(modTestsPure13)で固定する。
    # 現在時刻は引数(nowAt)で受け取る=この中で Now を呼ばないことが、
    # 終了目安のゴールデンテストを成立させている唯一の条件。
    # R15-7(2026-08-04 254頁対応)で追加した純ロジック:
    #   OcrEstMinutes/OcrConfirmAskFor/OcrDeclineMemoFor(7b 事前確認の見積もり
    #     と文面。これから読む頁数×レートで出し、キャッシュから復元できる頁は
    #     数えない=全頁揃っている資料では確認そのものが出ない)
    #   GsBudgetSec(7c 画像化待ち予算の頁数連動。max(config, 頁数×8秒))
    #   OcrResumeMemo(7d 復元があった取込のメモ冒頭)
    "optOcrEta": {
        "closed": True,
        "required": [
            "Ping", "BatchLabel", "OcrPageBanner", "RemainingText",
            "OcrCapMemoFor", "OcrAbortMemoFor", "OcrPartialMemoFor",
            "RemainingWaitSec",
            "OcrEstMinutes", "OcrConfirmAskFor", "OcrDeclineMemoFor",
            "GsBudgetSec", "OcrResumeMemo",
            # R15-FixA(2026-08-04 レビュー裁定 Fix-A)で追加した純ロジック:
            #   ComposeOcrMemo(FA-4): 本棚カードのメモの組み立てを1箇所へ。
            #     「復元の冒頭文は前置・置換禁止」「頁欠け/中断の理由と設定上限の
            #     説明は連結」という契約をここだけが持つ。従来は optOcrPage と
            #     optVision.OcrCapMemo が別々に組み立て、後から来たものが前を
            #     丸ごと置換していたため、事実が片方ずつ消えていた。
            #   BatchWaitSec(FA-5iii): 1バッチの画像化待ちの上限。資料あたりの
            #     残り予算をそのまま1バッチへ渡していたため、1バッチ目のハングが
            #     資料の予算を全部食い潰していた。
            #   RenderWaitBanner(FA-5i): 画像化待ちの実況文(経過秒つき)。
            "ComposeOcrMemo", "BatchWaitSec", "RenderWaitBanner",
            # ClampPageMs(2026-08-04 R15-FixB FB-4): ui_state に永続化される
            # 1頁あたり実績を常識の幅(3〜120秒)へ丸める純関数。異常値が1度
            # 混ざると以後ずっと嘘のETAを出し続けるので、読み出しで必ず通す。
            "ClampPageMs",
        ],
    },
    # optOcrCache(2026-08-04 R15-7d): 画像PDF OCRの頁チェックポイント。
    # 読めた頁を隠しシート "ocr_cache" へ控え、次の取込で読み直さない。
    # Vision(ChatGPTV)のハングはVBA側から制御できない(RIBBON_API_CONFIRMED
    # .md:29 に待ち秒数の引数が無い)ため、被害を限定する唯一の手段が
    # 「読めた頁を失わないこと」になる(実機第4報 RC7)。
    # opt層でシートを触る唯一のモジュール。作法(ThisWorkbook経由・
    # xlSheetVeryHidden)は modEmbed.EnsureVectorSheet を踏襲する。純ロジック
    # ではない(シートI/O・ログ)ので PURE_LOGIC_MODULES には載せないが、
    # 鍵の組み立て(DocPrefixFor/CacheKeyFor/CacheTextFor)だけは副作用ゼロで、
    # LO実行テスト(modTestsPure13)から直接呼んで差し替え検知を固定する。
    "optOcrCache": {
        "closed": True,
        "required": [
            "Ping", "DocPrefixFor", "CacheKeyFor", "CacheTextFor",
            "BeginDoc", "CachedText", "SaveRange", "HasSaved",
            # PurgeFor(2026-08-04 R15-FixB FB-1): 旧 PurgeDoc を置き換えた。
            # 控えを消してよいのは「OCRが読み切った」ときではなく「資料が
            # 本棚に done として並んだ」ときで、それを知っているのはコア層
            # (modShelf)だけ。パスを受け取り、その資料の行だけを消す。
            "PurgeFor", "FinishDoc", "GcOldRows",
            "ConfirmAskFor", "DeclineMemo",
        ],
    },
    # ---- R11-F1 分割(憲章§4-6の容量救済)。移設元と新設先を closed で固定し、
    #      「移したつもりで元にも残っている」「新設先にうっかり公開APIが増える」を機械で止める。
    # R11-F1: テーマ塊を modSkin へ移設した残り。UiColor/UiTheme/ToggleTheme は薄い委譲として残す。
    "modUI": {
        "closed": True,
        "required": [
            "InitUI", "AddChatBubble", "UpdateBubbleText", "ChatBottomFor",
            "BringFixedToFront", "ToggleTheme", "RestoreExcelUI", "EnsureAppView",
            # EnsureSessionResources(2026-08-01 R12-3-8): X閉じ→キャンセルで
            # Auto_Close だけが走った後、ホットキー3種と自動同期の予約を
            # 遷移時に冪等に戻す自己修復。EnsureAppView(表示の自己修復)と
            # 同じ場所・同じ考え方なので modUI に置き、modBoot.Boot もここを呼ぶ。
            "EnsureSessionResources",
            "GoToNexus", "GoToNativeSheet", "ActivateSheetRobust", "ParkFocus",
            "Repaint", "UiColor", "UiTheme", "FreezeShapePlacement",
            "RecalcChatBottom", "SettleChat", "ClearChat", "MarkActiveBubble",
            "BubbleTextOf", "LatestAiBubbleName",
        ],
    },
    # R11-F1: modUI からテーマ塊(CurrentTheme/SaveTheme/ThemeColor/ApplyTheme/PaintBubble/PaintActionButton/SetShapeTextColor/ThemeIcon)を受け入れた。
    "modSkin": {
        "closed": True,
        "required": [
            "BeautifyAll", "StyleShape", "ApplyHeaderDepth", "ApplyGradient",
            "ApplyLightShadow", "ApplyGreenDepth", "StyleBubble", "ApplySoftShadow",
            "EffectiveSkin", "ResolveColor", "CycleSkin", "ShowToast",
            "CurrentTheme", "SaveTheme",
            "ThemeColor", "ApplyTheme", "PaintBubble", "PaintActionButton",
            "SetShapeTextColor", "ThemeIcon",
        ],
    },
    # modProgressBar(2026-08-05 R18-1a): modSkin から移設した進捗バナー一式。
    # modSkin が27,990字(上限まで残り10字)で、実機第5報①の修正(幅の
    # viewport連動・ZOrder遮蔽の根治・砂時計の停止・孤児バナーの掃除)が
    # 1行も入らなかったための分割(憲章§4-6)。
    # PROGRESS_CANCEL_NAME / PROGRESS_WORK_NAME は Public Const:
    # Private Const はモジュールを跨いで参照できないため(実機VBA/LOとも)。
    # BarWidthFor は幅決定の純ロジック(modTestsPure15 が固定する)。
    # SweepOrphans は modHub.EnsureHubLayout からの孤児掃除(取込中は何もしない)。
    "modProgressBar": {
        "closed": True,
        "required": [
            "PROGRESS_CANCEL_NAME", "PROGRESS_WORK_NAME",
            # BarHeightFor(2026-08-05 R18H FA-2): 文面が本文可視幅に収まらない
            # ときバナーを2行ぶんへ広げる判定。幅(BarWidthFor)と同じ流儀の
            # 純ロジックで、modTestsPure16 が境界をゴールデンで固定する。
            "BarWidthFor", "BarHeightFor",
            "PaintProgress", "ClearProgress", "SweepOrphans",
        ],
    },
    # modViewport(2026-08-05 R18-3b / 2026-08-06 R19-1a で拡張): 画面の幾何を
    # 「見える範囲」に合わせる共通部品。実機第6報①で、ScrollArea はホイールを
    # 止めない(公式仕様=セル選択とスクロールバーの制限のみ)こと、右余白は
    # スクロールではなく列幅の【寸法】の問題であることが確定したため、
    # 5画面が同じ算数を書かずに済むよう共通部品を集約する(憲章§4-5)。
    #   ApplyScrollBound : ws.ScrollArea の設定(補助手段として維持)
    #   FitBandToViewport: 吸収列で列帯の合計幅を可視幅ぴったりに合わせる(右余白の唯一の解)
    #   ContentRight     : 帯・ピル・主ボタンが共有する右端の単一情報源
    #   BoundAddr        : 塗り/ScrollArea/Lockedが共有する実使用範囲の文字列
    #   BoundFor         : 列も実測で決める旧口(吸収列方式に移せない画面用)
    #   ViewportHeight   : 可視高(modUIMain.ViewportWidth の縦版)
    #   LogViewport      : 実機の可視幅×可視高の観測点(1画面1セッション1回)
    #   PadPtNeeded / RightEdgeAt / BoundBottomY / ColLetter は純ロジックで
    #   modTestsPure16 がゴールデンで固定する。
    "modViewport": {
        "closed": True,
        "required": [
            "ApplyScrollBound", "FitBandToViewport", "ContentRight",
            "BoundAddr", "BoundFor", "ViewportHeight", "LogViewport",
            "PadPtNeeded", "RightEdgeAt", "BoundBottomY", "ColLetter",
        ],
    },
    # R11-F1: 回答アクション系を modAppAct へ分離した残り(質問→回答/取込/ナビ/終了)。MAX_INPUT_CHARS は modAppAct と共有するため Public。
    "modApp": {
        "closed": True,
        "required": [
            "MAX_INPUT_CHARS", "LaunchNexus", "OnSend", "ModeCaption", "OnPeek",
            "OnPeekClose", "OnAddDocs", "OnNavChat", "OnNavHome", "OnNavShelf",
            "OnRefreshUI", "OnToggleMode", "OnToggleSpeed", "SpeedCaption",
            "OnLangCycle", "SummonNexus", "HotSend", "OnClearChat", "OnSaveAndExit",
        ],
    },
    # R11-F1: modApp から分離した回答の文脈アクション行(描画4本+ボタン6本のハンドラ)。
    # R13-6a(2026-08-03): 「続けて質問」の armed followup 一式。InputBoxを廃し、
    #   既存の広い入力欄+送信ボタンの1本道へ合流させるための印・チップ・解除。
    #   ArmFollowup は modUIMain.OnFollowupButton / OnActDrill の共通実体、
    #   ConsumeArmedFollowup は modApp.OnSend が1回だけ引く印、
    #   RedrawFollowupChip は再描画時の掃除、OnFollowupChipOff は Shape.OnAction。
    "modAppAct": {
        "closed": True,
        "required": [
            "DrawActions", "ClearConfidence", "DrawConfidence", "ClearActions",
            "OnActBad", "OnActDrill", "OnActResolve", "OnActUnsure", "OnActWord",
            "OnActCopy",
            "ArmFollowup", "ConsumeArmedFollowup", "RedrawFollowupChip",
            "OnFollowupChipOff",
        ],
    },
    # R11-F1: 受信箱(DrawInbox)を modHubStat へ移設した残り。
    "modHub": {
        "closed": True,
        "required": [
            # HUB_BOUND(2026-08-05 R18-3a/3b/5b): Hub画面の実使用範囲。書式の
            # 適用範囲・ScrollArea・フッターの境界チェック(modHubStat.DrawFooter)
            # が同じ1つの値を見るための Public Const。
            "HUB_BOUND",
            "EnsureHubLayout", "OnGoChat", "OnGoVault", "OnGoDash", "OnQuickAsk",
            "OnLangCycle", "OnThemeToggle", "OnHelp", "OnSaveAndExit",
            "OnCheckUpdates", "OnShareHelp", "OnOwnerReport", "OnAnonFeedback",
            "OnRedraw",
        ],
    },
    # R11-F1: Hubの数字とお知らせ。DrawInbox を modHub から受け入れた。
    "modHubStat": {
        "closed": True,
        "required": [
            "AllowShareQueries", "ShareQueriesAllowed", "PendingUpdatesCached",
            "InvalidatePending", "OnSyncPending", "PendingLabel", "RemoveHubShapes",
            "NumText", "SafeStat", "AskTotal", "SafeSavedMinutes", "SafeChunks",
            "ChunkUsage", "FmtMin", "DefaultTileValue", "OrgMin", "TilesHeight",
            # DrawFooter/OnFooterPortal(2026-08-05 R18-5b): Hub最下部の
            # 「(C) リスクコンサルティング支援部」と社内ポータルへの導線。
            # modHub が WARN帯まで残り401字だったため描画側もここへ置いた
            # (DrawInbox を R11-F1 でここへ移したのと同じ判断)。
            "DrawStatTiles", "DrawInbox", "DrawQuickAskCards",
            "DrawFooter", "OnFooterPortal",
        ],
    },
    # R11-F1: ツールバーを modKnowledgeBar へ分離した残り(ヘッダー/右肩ピル/モード管理/ハンドラ)。
    "modKnowledge": {
        "closed": True,
        "required": [
            # SHELF_BOUND(2026-08-05 R18-3a/3b): 「マイ本棚」シートの実使用範囲。
            # table/gallery/shared の3モジュールが書式適用範囲とScrollAreaの
            # 両方で参照するため、共通クロムを持つ modKnowledge が単一情報源。
            "CHROME_ROWS", "SHELF_BOUND", "DrawChrome", "PrepareScreenView", "IsTableMode",
            "ContentTop", "SearchCellAddress", "OnGoGallery", "OnGoShared",
            "OnGoTable", "OnBackHub", "OnHelp", "OnToChat", "OnSearch", "OnGapBoard",
            "OnChannels", "OnRegister", "OnAddFiles", "OnPackOut", "OnPackIn",
            "OnSync", "OnPickFolder", "OnDelete", "RefreshCurrent",
        ],
    },
    # R11-F1: modKnowledge から分離したツールバー(BAR_H は DrawChrome が行3の高さに使うため Public)。
    # ToolbarContentRight(2026-08-03 R14-2a・実機第3報 RC10): Shapeを作らず
    # DrawToolbarと同じ算数(ToolbarSpec+FlowLeft)だけを走らせ、ツールバーの
    # 実際の右端を返す。modKnowledge.DrawChromeが右肩ピルのアンカーに使う
    # (ピルがFlowRightで帯の右端へ密着する一方、ツールバーはFlowLeftで
    # 左詰めのため、ボタンが少ない端末で両者の右端が食い違って見えていた)。
    "modKnowledgeBar": {
        "closed": True,
        "required": [
            "BAR_H", "DrawToolbar", "ToolbarContentRight",
        ],
    },
    # R11-F1: ギャラリー系を modVaultGallery へ分離した残り(ナレッジ登録フォーム)。
    "modVault": {
        "closed": True,
        "required": [
            "ShowVaultInput", "OnVaultSubmit", "OnVaultCancel",
            "RegisterKnowledgeText",
        ],
    },
    # R11-F1: modVault から分離したナレッジ倉庫ギャラリー(カード一覧・検索・ページング)。
    "modVaultGallery": {
        "closed": True,
        "required": [
            "ShowVaultGallery", "OnVaultSearch", "OnVaultPrev", "OnVaultNext",
            "OnVaultBackToChat", "OnVaultCardClick",
        ],
    },
    # R11-F1: 本文の描画(KPI/経験値/バッジ/クラスタ地図)を modDashStat へ移設した残り。
    "modDash": {
        "closed": True,
        "required": [
            "ShowDashboard", "OnDashBackToChat", "OnHelp", "OnDashRefresh",
            "OnDashRestore",
        ],
    },
    # R11-F1: ダッシュボードの数値+本文描画。KPI_X0/ROW_WIDTH は modDash のヘッダー・管理者行と共有する版面基準のため Public。
    "modDashStat": {
        "closed": True,
        "required": [
            "KPI_CARD_W", "KPI_GAP", "KPI_X0", "ROW_WIDTH", "SavedTimeDeltaLabel", "CountUsageEvent",
            "IsThisMonthStamp", "IsLastMonthStamp", "UsageBarText", "FormatMinutes",
            "SafeGetStat", "SafeSavedMinutes", "SafeTotalChunks", "SafeShelfMax",
            "SafeLevel", "SafeExpTotal", "SafeExpFloorForLevel", "SafeLevelProgress",
            # DrawChartPlaceholder は 2026-08-05 R18-4(実機第5報③)で削除した
            # (ナレッジ地図の可視化を撤去。ChartNoteY はバッジ棚の下端=管理者
            #  セクションの起点として現役なので残す)。
            "DrawKpiRow", "DrawExpBar", "ChartNoteY", "DrawBadgeShelf",
        ],
    },
    # R11-F1: 発信/収集/GCと低水準I/Oを modInsightIo へ分離した残り(受信箱シートの参照・選択)。EnsureSheet は modInsightIo から呼ぶため Public。
    "modInsight": {
        "closed": True,
        "required": [
            "PendingQACount", "GapCount", "PendingQAAt", "IsSelected",
            "ToggleSelected", "SelectAllPending", "SelectedCount",
            "PendingRowsRanked", "RowField", "MarkQAConsumed", "SameQuestionCount",
            "QABodyText", "EnsureSheet", "GapListText",
        ],
    },
    # R11-F1: modInsight から分離した共有フォルダとのやり取り(発信/収集/GC)。
    "modInsightIo": {
        "closed": True,
        "required": [
            "EmitVerifiedQA", "EmitGap", "EmitCorrection", "CollectInsights",
        ],
    },
    # ---- R11-F2 で契約に載せたモジュール(監査4: 93モジュール中46しか契約が
    #      有効でなかった件の解消を、今回の変更が触った範囲から進める)。
    # modShare: 共有フォルダのベースパス解決と到達性プローブ(唯一の窓口)。
    "modShare": {
        "closed": True,
        "required": [
            "BasePath", "Reachable", "ProbePath", "ReportFailure",
            "ReportSuccess", "SubDir", "ResetProbe",
        ],
    },
    # modUiLock: 全ハンドラ共通の再入ロックと取込中の関所。
    # ConfirmCloseDuringIngest(2026-08-03 R13-4d): 取込中の終了要求を
    # 「無反応で拒否」から「事情を伝えて選ばせる」へ変えた確認口。
    # 呼び出し元は modApp.OnSaveAndExit と ThisWorkbook.Workbook_BeforeClose
    # (開発構成のみ。本番はThisWorkbookを自己インストーラが占有する)。
    "modUiLock": {
        "closed": True,
        # CancelCloseOk(2026-08-03 R13 F5): 「中断して終了」の承諾は、
        # 終了を取りやめた全経路で捨てる。残すと承諾から60秒のあいだ、
        # ウィンドウの×が確認なしで閉じる窓が開く。
        "required": ["Enter", "Leave", "IsBusy", "BlockIfIngesting",
                     "ConfirmCloseDuringIngest", "CancelCloseOk"],
    },
    # modAppState: Nexus画面の状態(モード/対象バブル/入力欄/ui_state)の唯一の窓口。
    "modAppState": {
        "closed": True,
        "required": [
            "MODE_KEY", "SetActiveBubble", "ShelfIsEmpty", "AnswerWithoutShelf",
            "HasTarget", "TargetBubbleName", "TargetText", "SharePath",
            "AskGeneral", "TrimPairs", "RagSpeed", "CurrentMode",
            "UpdateModeButton", "ReadInputCell", "RestoreInputCell",
            "ClearInputCell", "ReadUiState", "WriteUiState",
        ],
    },
    # modUtilText(2026-07-31 R11-F2 新設): UTF-8読み書き・経過ミリ秒・移動平均・
    # GSページ分割添字の共通部品。ここに何でも足されると「雑多置き場」になるので
    # closed で固定する。
    "modUtilText": {
        "closed": True,
        "required": [
            "ReadTextFileUtf8", "WriteTextFileUtf8", "ElapsedMsSince",
            "BlendPerItemMs", "GsPageIsBlank", "CleanTextLen", "GsPageBounds",
            # 2026-08-01(R12-1-4): カレンダー設定(和暦)非依存の日付文字列。
            # 日付を文字列で永続化・比較する箇所はここだけを通す。
            "IsoDate", "IsoDateTime", "NormalizeIsoDate",
            # 2026-08-04(R15-1a): 再入ガードの失効判定(最終ビート基準)。
            # 4本のガード(modShelf/modShelfBatch/modEmbed/modShelfSync)が
            # 同じ式を使うための純関数。時刻の算数はこのモジュールに集める。
            "GuardExpired",
            # 2026-08-01(R12-3-10): 区切り無しの統計キー用。同じ元号防御を
            # yyyymmdd/yyyymm/yyyy でも1箇所に集める(modBoard/modAsk/modHelp)。
            "IsoDateCompact", "IsoYm", "IsoYear",
            # 2026-08-01(R12-2-1): 数式インジェクション対策(数式注入)の
            # 共通化先。modChatLogの重複実装を吸収し、未信頼テキストを
            # セルへ書く全経路(modInsightIo/modPack/modChannel)から呼ぶ。
            "SanitizeForCell",
            # 2026-08-03(R13 F8): 段ごとの所要時間を1本の短い文字列へ畳む
            # 純ロジック(書式 "expand=1200;emb=6x830")。呼ぶのは
            # modGateway だけだが、中身は文字列処理なのでここに置き、
            # LOの実行テストで書式そのものを固定する。
            "AppendStepBuf",
            # 2026-08-04(R15-8a): usage_log "ingest" 行のdetailへ生成/重複の
            # 内訳を添える書式(実機第4報 RC1)。modShelfから1箇所だけ呼ぶが、
            # 文字列組み立てのみの純関数なのでここに置く。
            "IngestChunksDetail",
            # 2026-08-04(R15-FixB FB-5): 事前確認で「いいえ」を選んだ資料の
            # メモかどうか(先頭一致)。見送りは失敗ではないので一括取込の
            # 集計から分離する必要があり、判定は opt 層(文面を作る側)と
            # コア層(件数を数える側)の両方から呼ばれる。両方から見えるのは
            # 基盤層だけなので、先頭句の定数と判定をここに1つ置く。
            "DECLINE_MEMO_HEAD", "IsDeclineNote",
            # 2026-08-05(R16-2a): 「作業用Excelを開く」ボタン(modWorkExcel)が
            # 起動するコマンド行の組み立て。純粋な文字列処理(""囲み+
            # "\EXCEL.EXE"+" /x")なので基盤層に置く。modWorkExcelはこれを
            # 呼ぶだけで、パスの正規化ロジックを持たない。
            "BuildWorkExcelCmd",
        ],
    },
    # ---- 7.8 テストモジュール ----
    "modTestRunner": {
        "closed": True,
        "required": ["ResetTests", "Check", "Failures", "ReportText", "RunAllPureTests"],
    },
    "modTestsPure": {
        # 個別のRunXxxはmodTestsPure内部でいくつ増えてもよい。modTestRunner
        # から呼ばれる入口 RunAll だけが契約(本Lintツール実装者=1-Iが
        # 2-Xチームに要求する最小契約)。
        "closed": False,
        "required": ["RunAll"],
    },
    # modTestsPure2: 2026-07-12 Wave3-Tで追加。modTestsPureが§7.1の
    # 「1モジュール30,000字以内」を超えたため、modPrompts/modShelfSync/
    # modPack関連のテストを分割した先(src/test/modTestsPure2.bas冒頭コメント
    # 参照)。modTestsPure.RunAllの末尾から呼ばれる入口 RunAll2 だけが契約。
    "modTestsPure2": {
        "closed": False,
        "required": ["RunAll2"],
    },
    # modTestsPure3: 2026-07-30 R2要件Bで追加。modTestsPure(28,906字)/
    # modTestsPure2(29,170字)とも30,000字上限まで余裕が無く、要件Bの回帰
    # テスト(空ページ防御・化けページ圧縮後のチャンク数)を追加する場所が
    # 無かったための新規分割先。modTestsPure2.RunAll2の末尾から呼ばれる
    # 入口 RunAll3 だけが契約。
    "modTestsPure3": {
        "closed": False,
        "required": ["RunAll3"],
    },
    # modTestsPure4: 2026-07-31 R6で追加。modTestsPure3も上限に近づいたため、
    # 画像PDFのOCR取込(optOcrCoreのGSコマンド組み立て・modUtilのページ付き
    # テキスト・全ページ化けPDFの振り替え)のテストを分割した先。
    # modTestsPure3.RunAll3の末尾から呼ばれる入口 RunAll4 だけが契約。
    "modTestsPure4": {
        "closed": False,
        "required": ["RunAll4"],
    },
    # modTestsPure7: 2026-08-01 R12-3で追加。既存のテストモジュールがいずれも
    # 30,000字上限に余裕が無く(憲章§4-6)、堅牢化の回帰テストを置く先として
    # 新設した。modTestsPure4.RunAll4の末尾から呼ばれる入口 RunAll7 だけが契約。
    "modTestsPure7": {
        "closed": False,
        "required": ["RunAll7"],
    },
    # modTestsPure8: 2026-08-01 R12-8で追加した分割先。R12-Hの敵対的レビューで
    # 「自己ルール(分割先は入口だけを契約にする)の適用漏れ」として指摘された。
    # modTestsPure7.RunAll7 の末尾から呼ばれる入口 RunAll8 だけが契約。
    "modTestsPure8": {
        "closed": False,
        "required": ["RunAll8"],
    },
    # modTestsPure9: 2026-08-01 R12-4で追加。modTestsPure8が28,000字(WARN帯)に
    # 達したための分割先。modTestsPure8.RunAll8の末尾から呼ばれる入口 RunAll9
    # だけが契約。
    "modTestsPure9": {
        "closed": False,
        "required": ["RunAll9"],
    },
    # modTestsPure10: 2026-08-03 R13-7cで追加。modTestsPure9が28,000字の
    # WARN帯に触れるための分割先。modTestsPure9.RunAll9の末尾から呼ばれる
    # 入口 RunAll10 だけが契約。
    "modTestsPure10": {
        "closed": False,
        "required": ["RunAll10"],
    },
    # modTestsPure11: 2026-08-03 R14-3/R14-4で追加。modTestsPure9が28,000字の
    # WARN帯に触れる(残り578字)ための分割先。modTestsPure10.RunAll10の
    # 末尾から呼ばれる入口 RunAll11 だけが契約。
    "modTestsPure11": {
        "closed": False,
        "required": ["RunAll11"],
    },
    # modTestsPure12: 2026-08-03 R14-8で追加。modTestsPure11 に R14-8 の
    # テストを足すと30,000字上限を超えるための分割先。
    # modTestsPure11.RunAll11 の末尾から呼ばれる入口 RunAll12 だけが契約。
    "modTestsPure12": {
        "closed": False,
        "required": ["RunAll12"],
    },
    # modTestsPure13: 2026-08-04 R15波2で追加。modTestsPure11/12 はどちらも
    # 27,000字台で、R15-4/5/6 の真理表とゴールデン文字列を足すと上限を超える
    # ための分割先。modTestsPure12.RunAll12 の末尾から呼ばれる入口
    # RunAll13 だけが契約。
    "modTestsPure13": {
        "closed": False,
        "required": ["RunAll13"],
    },
    # modTestsPure14: 2026-08-04 R15-FixA で追加。modTestsPure13 に Fix-A の
    # 真理表とゴールデン文字列を足すと上限まで残り11字になるための分割先。
    # modTestsPure13.RunAll13 の末尾から呼ばれる入口 RunAll14 だけが契約。
    "modTestsPure14": {
        "closed": False,
        "required": ["RunAll14"],
    },
    # modTestsPure15(2026-08-05 R16波3): modTestsPure14 の容量逼迫による分割先。
    #   入口は RunAll15 の1本だけで、modTestsPure14.RunAll14 の末尾から呼ばれる。
    "modTestsPure15": {
        "closed": False,
        "required": ["RunAll15"],
    },
    # modTestsExcel はMASTER_SPECがPublic契約を明示していないため対象外。
}

# 純ロジックモジュール(R4): Excelオブジェクトトークン禁止
# modPrompts はMASTER_SPEC R4本文には明記されていないが、§7.3で
# 「純文字列」モジュールと明記されており、テストハーネス発注元の指示
# (1-Iタスク定義)により本Lintでは禁止トークン検査の対象に加える。
PURE_LOGIC_MODULES = {
    "modUtil", "modChunker", "modPii", "modTypes",
    "modTestRunner", "modTestsPure", "modTestsPure2", "modTestsPure3", "modPrompts",
    "modRagParse", "modSparse", "modMode",
    # modChunkMeta(2026-08-05 R17 Phase1): section_path/refs_out の抽出。
    # 副作用ゼロの文字列処理だけなので、LO実行テストから直接呼べる状態を
    # 機械で守る(シートI/Oは modChunkMetaStore が持つ)。
    "modChunkMeta",
    # modChrome(2026-07-30 R4): 配置計算だけを持つのでExcelオブジェクトは
    # 一切要らない。ここへ載せることで「うっかりRangeを触る」改修を機械で止める。
    "modChrome",
    # optOcrCore(2026-07-31 R6): GSコマンド文字列の組み立てとページ上限の算数
    # だけを持つ。opt層のモジュールだが副作用ゼロで、LO実行テストから直接
    # 呼べる状態を維持するために純ロジック検査の対象へ入れる。
    "optOcrCore",
    # optOcrEta(2026-08-04 R15-5b): optOcrCore から移設した進捗バナー・ETA・
    # 打ち切りメモの純ロジック。移設先でも副作用ゼロを機械で守る。
    "optOcrEta",
    # modTestsPure4: optOcrCore/modUtil(ページ付きテキスト)の純ロジックテスト。
    "modTestsPure4",
    # modShareRule(2026-07-31 R8): 共有系の判定式だけの純ロジック。
    "modShareRule",
    # modTestsPure5(2026-07-31 R8): modShareRule の境界値テスト。
    "modTestsPure5",
    # modTestsPure6(2026-07-31 R8b): 敵対的レビュー対応(B1/B7b/B10)のテスト。
    "modTestsPure6",
    # modTestsPure7(2026-08-01 R12-3): 堅牢化の純ロジックテスト。
    "modTestsPure7",
    # modTestsPure8(2026-08-01 R12-8): テスト補強(モジュール全滅解消・ハッシュ
    # 互換ゴールデン値)の純ロジックテスト。modTestsPure7の容量逼迫による分割先。
    "modTestsPure8",
    # modTestsPure9(2026-08-01 R12-4): 検索スケール恒久対策の純ロジックテスト。
    # modTestsPure8の容量逼迫(WARN帯)による分割先。
    "modTestsPure9",
    # modTestsPure10(2026-08-03 R13-7c): チーム/部・共有ビーコンのteam列の
    # 純ロジックテスト。modTestsPure9の容量逼迫(WARN帯)による分割先。
    "modTestsPure10",
    # modTestsPure11(2026-08-03 R14-3/R14-4): 共有読みコピーの失敗判定、
    # OCRのバッチ分割・上限メモの純ロジックテスト。modTestsPure9の容量逼迫に
    # よる分割先(modTestsPure10.RunAll10の末尾から呼ばれる)。
    "modTestsPure11",
    # modTestsPure12(2026-08-03 R14-8): 入念モードの段数・出典突合と、回答本文の
    # 記法正規化の純ロジックテスト。modTestsPure11の容量逼迫による分割先。
    "modTestsPure12",
    # modTestsPure13(2026-08-04 R15波2): OCRの正直さ(頁欠け・上限打ち切り)、
    # ETAと終了目安、中断メモの純ロジックテスト。modTestsPure11/12 の
    # 容量逼迫による分割先。
    "modTestsPure13",
    # modTestsPure15(2026-08-05 R16波3): 逆質問=番号選択肢のパースと合成、
    # 精読(chunk_idからの文書順復元と前後radius)、既出チャンク降格の
    # 純ロジックテスト。modTestsPure14(24,825字)の容量逼迫による分割先で、
    # modTestsPure14.RunAll14 の末尾から呼ばれる。
    "modTestsPure15",
    # 2026-07-31(R11-F2): qa層の3モジュールを追加。いずれも実測でExcel
    # オブジェクトトークン0件(Worksheets/Range(/Application./ThisWorkbook/
    # MsgBox/ActiveSheet が1つも無い)。純ロジックであることを規約として
    # 固定し、「ちょっとRangeを見たい」という改修を機械で止める。
    #   modBitwiseOpt: ビット演算による候補絞り込み。
    #   modFollowup  : 深掘り候補の抽出と本文からの除去。
    #   modClarify   : 聞き返し文の生成と番号選択の判定。
    "modBitwiseOpt", "modFollowup", "modClarify",
    # modUtilText(2026-08-04 R15-1a): 日付/経過時間の算数と文字列処理だけの
    # 共通部品。実測でExcelオブジェクトトークン0件(ADODB.Stream は
    # CreateObject 経由でExcel依存ではない)。R15-1a の GuardExpired を
    # ここへ置くにあたり、「ちょっとRangeを見たい」改修を機械で止める。
    "modUtilText",
    }

FORBIDDEN_TOKEN_PATTERNS = [
    (re.compile(r"\bWorksheets\b"), "Worksheets"),
    (re.compile(r"\bRange\s*\("), "Range("),
    (re.compile(r"\bApplication\."), "Application."),
    (re.compile(r"\bThisWorkbook\b"), "ThisWorkbook"),
    (re.compile(r"\bMsgBox\b"), "MsgBox"),
    (re.compile(r"\bActiveSheet\b"), "ActiveSheet"),
]

# R3: Application.Run 第1引数リテラルのホワイトリスト
RUN_LITERAL_WHITELIST_EXACT = {"ChatGPT", "GetEmbeddings", "LimitCheck", "modBoot.Boot"}
RUN_LITERAL_WHITELIST_PREFIX = re.compile(r"^opt[A-Za-z]\w*\.")
# ChatGPT/GetEmbeddings の直接Application.Runは modGateway 内のみ許可(R3本文)
# LimitCheck: 起動時利用期限チェック(裁定D3・RIBBON_API_CONFIRMED.md §1 #13)。
# modGateway.RunLimitCheck 内でのみ呼ぶ。
RUN_LITERAL_GATEWAY_ONLY = {"ChatGPT", "GetEmbeddings", "LimitCheck"}
# 変数経由(非リテラル)のApplication.Runはこの2ファイルのみ許可
RUN_VARIABLE_ALLOWED_MODULES = {"modGateway", "modFeatures"}

# R1例外のうち modSkin.ShowToast を呼んでよい機能層モジュール(R10c L5)。
# ShowToastは表示に1.1秒のブロッキング待ちを含むため、「完了を1回だけ知らせる」
# 用途に限る。取込(modShelf)と同期(modShelfSync)の完了通知だけが該当する。
# 2026-07-31(R11-E 監査3 M-5): modStatsのバッジ獲得通知も「一度きりの完了
# 告知」で、上の2モジュールと同じ性質(業務ロジックの継続に影響しない・
# 起きても稀)なので追加する。MsgBoxのままだと起動シーケンスの途中で
# モーダルが割り込み、利用者が「まだ動いているのか」判断できなくなる
# (憲章§3-4)。
R1_TOAST_ALLOWED_MODULES = {"modShelfBatch", "modShelfSync", "modStats"}

# R1例外(UIロックの状態問い合わせ。2026-08-01 R12-H-2a 裁定)。
# modUiLock.IsBusy は副作用ゼロの読み取り専用ゲッターで、UIを操作しない。
# 自動同期(OnTime)は「利用者が質問・取込をしている最中には始めない」判断が
# 要るが、その事実を持っているのはUI層のロックだけである。SetStage 等の
# 通知コールバック例外と同じ性質(向きは逆だが、状態を1つ読むだけで
# 機能層の処理がUIの都合に依存しない)として、同期モジュールにのみ許す。
# 広げるときは必ずここへ足す=どのモジュールがUIロックを見ているかが1箇所で分かる。
R1_UILOCK_ALLOWED_MODULES = {"modShelfSync"}

# R1例外(中断ボタン付き進捗バナー。2026-08-04 R15-FixA FA-6)。
# 2026-08-05 R18-1a: 実体は modSkin → modProgressBar へ移設した(名前だけの変更)。
# modProgressBar.PaintProgress は modUIMain.ShowProgress の後半そのもので、性質は
# 既存例外の ShowProgress と同じ(実況を伝えるだけ)。cancellable:=True を
# 渡せるのは取込のバナーを1本化した modShelfBatch.ShowIngestBanner だけに
# 限る(他所へ広がったらLintで止める=押しても止まらない中断ボタンを二度と
# 生やさない)。modShelfSync/AddFilesResult はその1本を経由して出す。
R1_PROGRESS_ALLOWED_MODULES = {"modShelfBatch"}

# R1例外(取込前確認からの作業用Excel起動。2026-08-05 R18-1f)。
# modWorkExcel.OpenWorkExcelNow は WScript.Shell で別プロセスの excel.exe /x を
# 起動するだけで、引数も戻り値も無く、業務ロジックを一切呼び返さない
# (性質は既存例外の ShowProgress/PaintProgress と同じ「通知して終わり」)。
# 呼べる場所は取込前確認の2段目だけに限る: この確認は「取込中このExcelは
# 操作できない」と伝えた直後の1点で、そこ以外から取込層がUIのプロセス起動を
# 始めてよい理由は無い。広げるときは必ずここへ足す=どのモジュールが別Excelを
# 開けるかが1箇所で分かる状態を保つ。
# 2026-08-05 R18H FA-1: 2段目確認の記憶(1セッション1回)・文面・起動を
# modWorkExcel.OfferBeforeIngest へ丸ごと寄せた。取込層から見えるのは
# 「聞く場所を1行呼ぶ」だけで、性質は OpenWorkExcelNow と変わらない。
R1_WORKEXCEL_ALLOWED_MODULES = {"modShelfVision"}
R1_WORKEXCEL_ALLOWED_MEMBERS = {"OpenWorkExcelNow", "OfferBeforeIngest"}

# R2: opt直接トークン参照禁止(src/opt以外)
OPT_TOKEN_PATTERN = re.compile(r"\bopt[A-Za-z]\w*\s*\.")

# 型落ち検出: Dim/Static文
DIM_STMT_PATTERN = re.compile(r"^(Dim|Static)\s+(.*)$", re.IGNORECASE)
AS_KEYWORD_PATTERN = re.compile(r"\bAs\b", re.IGNORECASE)
AS_INTEGER_PATTERN = re.compile(r"\bAs\s+Integer\b", re.IGNORECASE)

# ReDim x(0 To -1) / (5 To 2) 等の「上限が負」の負範囲(実機VBAで実行時エラー9)。
# 定数指定の負リテラル上限のみを対象にする(変数上限 n-1 は実行時にしか
# 分からないため対象外。それらはcount>0ガード付きの正当なパターン)。
NEGATIVE_REDIM_PATTERN = re.compile(
    r"\bReDim\b[^(]*\([^)]*\bTo\s+-\d+\s*\)", re.IGNORECASE
)

ATTRIBUTE_VBNAME_PATTERN = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"', re.IGNORECASE)
OPTION_EXPLICIT_PATTERN = re.compile(r"^\s*Option\s+Explicit\s*$", re.IGNORECASE)

PUB_SUB_PATTERN = re.compile(r"^Public\s+Sub\s+([A-Za-z_]\w*)", re.IGNORECASE)
PUB_FUNC_PATTERN = re.compile(r"^Public\s+Function\s+([A-Za-z_]\w*)", re.IGNORECASE)
PUB_CONST_PATTERN = re.compile(r"^Public\s+Const\s+([A-Za-z_]\w*)", re.IGNORECASE)
PUB_TYPE_PATTERN = re.compile(r"^Public\s+Type\s+([A-Za-z_]\w*)", re.IGNORECASE)

DOTTED_REF_PATTERN = re.compile(r"\b([A-Za-z_]\w*)\.([A-Za-z_]\w*)")
# 注意: ThisWorkbook はここに含めない。ThisWorkbook.cls の契約上Publicは
# 0件(Private イベント転送のみ)だが、実コードの "ThisWorkbook.Worksheets"
# 等はExcel組み込みオブジェクトモデルのプロパティであって、我々の
# ThisWorkbook.cls が自前定義したPublicメンバーではない。これをmodX.Y
# 実在検証の対象に含めると「ThisWorkbook.Worksheets は未定義」という
# 誤検知になるため、モジュール形状の判定対象は mod*/opt* のみに絞る。
MODULE_SHAPED_NAME = re.compile(r"^(mod[A-Z]\w*|opt[A-Z]\w*)$")

FULLTEXT_ASSIGN_PATTERN = re.compile(r"\.Value\s*=", re.IGNORECASE)
FULLTEXT_HINT_PATTERN = re.compile(r"full_?text|answer_?text", re.IGNORECASE)
SAFELEFT_CALL_PATTERN = re.compile(r"SafeLeft\s*\(", re.IGNORECASE)


# ==============================================================================
# レイヤー判定(MASTER_SPEC §14のディレクトリ配置基準)
# ==============================================================================
LAYER_FOUNDATION = 0
LAYER_MID = 1
LAYER_UI = 2
LAYER_OPT = "opt"
LAYER_TEST = "test"

LAYER_LABEL = {
    LAYER_FOUNDATION: "基盤層(src/core)",
    LAYER_MID: "部品層/機能層(src/ingest,qa,pack,stats)",
    LAYER_UI: "UI層(src/ui)",
    LAYER_OPT: "opt層(src/opt)",
    LAYER_TEST: "テスト層(src/test)",
}


def classify_layer(relpath: Path):
    parts = relpath.parts
    if not parts:
        return None
    top = parts[0]
    if top == "core":
        return LAYER_FOUNDATION
    if top in ("ingest", "qa", "pack", "stats"):
        return LAYER_MID
    if top == "ui":
        return LAYER_UI
    if top == "opt":
        return LAYER_OPT
    if top == "test":
        return LAYER_TEST
    return None


# ==============================================================================
# ソース前処理: 行連結・コメント除去・文分割
# ==============================================================================
def merge_continuations(raw_lines: list[str]) -> list[tuple[int, str]]:
    """" _" で終わる行を次行と結合し、(開始行番号, 論理行) のリストにする。"""
    out: list[tuple[int, str]] = []
    acc = ""
    acc_start = None
    for i, raw in enumerate(raw_lines, start=1):
        line = raw.rstrip("\n\r")
        if acc_start is None:
            acc_start = i
        rstripped = line.rstrip()
        cont = rstripped.endswith(" _") or rstripped == "_"
        line_wo_cont = rstripped[:-1].rstrip() if cont else line
        acc = f"{acc} {line_wo_cont}" if acc else line_wo_cont
        if not cont:
            out.append((acc_start, acc))
            acc = ""
            acc_start = None
    if acc:
        out.append((acc_start, acc))
    return out


def strip_comment(line: str) -> str:
    """文字列リテラル内の ' は無視して、行コメント(')以降を切り落とす。"""
    in_str = False
    for i, c in enumerate(line):
        if c == '"':
            in_str = not in_str
        elif c == "'" and not in_str:
            return line[:i]
    return line


def split_top_level(s: str, sep: str) -> list[str]:
    """文字列リテラル・カッコの中は無視してトップレベルの sep で分割する。"""
    parts: list[str] = []
    depth = 0
    in_str = False
    cur: list[str] = []
    for c in s:
        if c == '"':
            in_str = not in_str
            cur.append(c)
        elif not in_str and c == "(":
            depth += 1
            cur.append(c)
        elif not in_str and c == ")":
            depth = max(0, depth - 1)
            cur.append(c)
        elif not in_str and depth == 0 and c == sep:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(c)
    parts.append("".join(cur))
    return parts


def iter_statements(raw_lines: list[str]) -> list[tuple[int, str]]:
    """行連結→コメント除去→コロン分割まで済ませた (行番号, 文) の列。"""
    stmts: list[tuple[int, str]] = []
    for lineno, ltext in merge_continuations(raw_lines):
        code = strip_comment(ltext)
        for piece in split_top_level(code, ":"):
            piece = piece.strip()
            if piece:
                stmts.append((lineno, piece))
    return stmts


# ==============================================================================
# モジュール情報
# ==============================================================================
@dataclass
class Finding:
    level: str   # "ERROR" | "WARN" | "SKIP"
    line: int
    message: str


@dataclass
class ModuleInfo:
    path: Path
    relpath: Path
    raw_text: str
    vb_name: str
    filename_stem: str
    statements: list = field(default_factory=list)
    public_names: dict = field(default_factory=dict)  # name -> kind
    layer: object = None
    findings: list = field(default_factory=list)

    def add(self, level: str, line: int, message: str) -> None:
        self.findings.append(Finding(level, line, message))


def load_module(path: Path, src_root: Path) -> ModuleInfo:
    raw_text = path.read_text(encoding="utf-8", errors="replace")
    raw_lines = raw_text.splitlines()
    stmts = iter_statements(raw_lines)

    vb_name = ""
    for line in raw_lines[:10]:
        m = ATTRIBUTE_VBNAME_PATTERN.match(line)
        if m:
            vb_name = m.group(1)
            break

    filename_stem = path.stem
    relpath = path.relative_to(src_root)

    info = ModuleInfo(
        path=path,
        relpath=relpath,
        raw_text=raw_text,
        vb_name=vb_name,
        filename_stem=filename_stem,
        statements=stmts,
    )
    info.layer = classify_layer(relpath)

    for lineno, stmt in stmts:
        m = PUB_SUB_PATTERN.match(stmt)
        if m:
            info.public_names[m.group(1)] = "Sub"
            continue
        m = PUB_FUNC_PATTERN.match(stmt)
        if m:
            info.public_names[m.group(1)] = "Function"
            continue
        m = PUB_CONST_PATTERN.match(stmt)
        if m:
            info.public_names[m.group(1)] = "Const"
            continue
        m = PUB_TYPE_PATTERN.match(stmt)
        if m:
            info.public_names[m.group(1)] = "Type"
            continue

    return info


# ==============================================================================
# 個別チェック
# ==============================================================================
def check_basics(info: ModuleInfo) -> None:
    # Option Explicit
    if not any(OPTION_EXPLICIT_PATTERN.match(line) for line in info.raw_text.splitlines()):
        info.add("ERROR", 1, "Option Explicit がありません")

    # VB_Name = ファイル名一致
    if not info.vb_name:
        info.add("ERROR", 1, 'Attribute VB_Name が見つかりません')
    elif info.vb_name != info.filename_stem:
        info.add(
            "ERROR", 1,
            f'Attribute VB_Name="{info.vb_name}" がファイル名"{info.filename_stem}"と不一致',
        )

    # モジュール30,000字以内
    #
    # 2026-07-28(レビューI-2): 上限ちょうどまで使い切ったモジュールが6本あり、
    # 「バグを1行直すこともできない」状態になっていた。上限超過はビルドが
    # 落ちるので事故にはならないが、気付くのが遅すぎる。残りが少なくなった
    # 時点で警告し、切り出しを促す。
    n = len(info.raw_text)
    if n > MAX_MODULE_CHARS:
        info.add("ERROR", 1, f"モジュールが{n}字で上限{MAX_MODULE_CHARS}字を超過")
    elif n > MODULE_WARN_CHARS:
        info.add(
            "WARN", 1,
            f"モジュールが{n}字で上限{MAX_MODULE_CHARS}字まで残り{MAX_MODULE_CHARS - n}字。"
            f"次の修正が入らなくなる前に凝集した機能を新モジュールへ切り出すこと",
        )


def check_cp932_safe(info: ModuleInfo) -> None:
    """CP932に無い文字がソースに混ざっていないか。

    このブックは開くたび vba_src シートのソースを VBE へ注入するが、VBE は
    コードを CP932 で保持する。CP932 に無い文字はリテラル "?"(0x3F)として
    保存されるため、実行時の文字列がそのまま化ける。
    2026-07-28 のレビュー(H-16)では、全RAG回答の信頼度表示・常時見えている
    モードボタン・言語名・LLMへの指示文まで20箇所が化けていた。
    しかも vbaProject.bin を覗かないと気付けない(ソースは正しく見える)。

    規約(modUINexusDraw 冒頭)は「非ASCIIは ChrW で組む」だが、規約は破られる。
    ここで機械的に落とす。実行時文字列は ERROR、コメントは WARN
    (コメントの化けは動作に影響しないが、次に読む人が混乱する)。

    代表的な差し替え先:
        〜 U+301C  -> ～ U+FF5E   (見た目は同じ。CP932 の 0x8160 はこちら)
        —  U+2014  -> ― U+2015   (CP932 の 0x815C)
        絵文字・ベトナム語の声調記号 -> ChrW() で組み立てる
    """
    # 検査するのはコメントを除いた実行文だけ。コメント側の絵文字は
    # 「このChrWが何の字か」を示す注釈で、化けても動作に影響しない。
    # そこまで ERROR/WARN にすると警告が数十件常駐して、本当に直すべき
    # 実行時文字列の1件が埋もれる(それでは検査の意味が無い)。
    for lineno, stmt in info.statements:
        bad = sorted({ch for ch in stmt if not _cp932_encodable(ch)})
        if not bad:
            continue
        shown = " ".join(f"{ch}(U+{ord(ch):04X})" for ch in bad[:6])
        info.add(
            "ERROR", lineno,
            f"CP932に無い文字が実行文に含まれる: {shown} 。"
            f"VBEへの注入時に'?'へ化けます。ChrW()で組むかCP932内の字へ置き換えてください",
        )


def _cp932_encodable(ch: str) -> bool:
    # Pythonのcp932コーデックはWindowsより寛容な文字(U+301C等)があるため、
    # Windowsの CP932 表に合わせて明示的に除外する。
    if ch in _CP932_DENY:
        return False
    try:
        ch.encode("cp932")
        return True
    except UnicodeEncodeError:
        return False


# WindowsのCP932では0x8160はU+FF5Eであり、U+301C(WAVE DASH)には対応バイトが
# 無い。同様にU+2212/U+00A2等もWindows側では落ちる。実機で"?"化を確認した
# 文字を明示的に拒否する(Pythonのコーデックだけに任せると見逃す)。
_CP932_DENY = frozenset("〜‖−¢£¬")


MODULE_DECL_RE = re.compile(
    r"^(?:Public|Private|Global|Dim)\s+(?:Const\s+|WithEvents\s+)?(\w+)", re.IGNORECASE)
PROC_HEAD_RE = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property)\s+\w+",
    re.IGNORECASE)
PROC_TAIL_RE = re.compile(r"^End\s+(Sub|Function|Property)\s*$", re.IGNORECASE)
TYPE_DECL_RE = re.compile(r"^(?:Public|Private)\s+(?:Type|Enum)\b", re.IGNORECASE)


def collect_module_level_names(raw_text: str) -> set[str]:
    """モジュールレベルで宣言された定数・変数の名前を集める。

    プロシージャの中の Dim は対象外(ローカルなので他モジュールから見えない)。
    Public Type / Public Enum は型名であって変数ではないので除外する
    (modTypes.ExtractedPage 等は他モジュールから正しく参照される)。
    """
    names: set[str] = set()
    inside = False
    for raw in raw_text.split("\n"):
        s = strip_comment(raw).strip()
        if not s:
            continue
        if PROC_HEAD_RE.match(s):
            inside = True
        elif PROC_TAIL_RE.match(s):
            inside = False
        elif not inside:
            if TYPE_DECL_RE.match(s):
                continue
            m = MODULE_DECL_RE.match(s)
            if m:
                names.add(m.group(1))
    return names


# ==============================================================================
# OnAction配線ハンドラの再入保護(2026-08-04 R15-2b・実機第4報 RC8)
# ------------------------------------------------------------------------------
# Shape.OnAction に文字列で配線された Public Sub は、取込中の DoEvents で
# 発火したクリックから【取込の途中に入れ子で】呼び出され得る。入れ子で走った
# 画面遷移は、取込が掴んでいるシート状態やCOMオブジェクトと噛み合わずに
# 失敗する(実機のE0202・「取込中にボタンを押すとExcelが応答なし」の正体)。
# 保護は全ハンドラの先頭に1行入れるだけだが、40箇所超に手で入れている以上
# 必ず抜ける。実際 modHubStat.OnSyncPending と modApp.OnRefreshUI の2本が
# 抜けており、実機第4報まで誰も気付けなかった。機械で見張る。
#
# 判定は「配線先の先頭付近に modUiLock.BlockIfIngesting または
# modUiLock.Enter の呼び出しがあること」。ERRORではなくWARNにしているのは、
# 保護が要らない正当なハンドラ(下のALLOWLIST)が現に存在し、
# 「取込中でも動くべき中断ハンドラ」を将来足す余地を残すため。
#
# 収集するのは .OnAction = "modX.Y" の【文字列リテラル代入】だけ。
# 変数経由(btn.OnAction = action)や連結("modAppAct." & CStr(acts(i)))は
# 静的には宛先が定まらないので対象にしない(誤検知ゼロを優先する)。
ONACTION_LITERAL_PATTERN = re.compile(
    r"\.OnAction\s*=\s*\"(mod[A-Za-z]\w*\.[A-Za-z_]\w*)\"\s*$", re.IGNORECASE)

# ハンドラの「先頭付近」= 最初の実行文からこの本数まで。
# 既存の全ハンドラは1〜2文目までに関所を通しており、4文あれば
# 「Dim を数本置いてから通す」書き方(modVault.OnVaultSubmit 型)まで拾える。
ONACTION_GUARD_LOOKAHEAD = 4

# 保護を求めないハンドラの名簿(名前は "modX.Y" 形式)。
# 増やすときは必ず理由を1行書くこと。ここが「押しても何も起きない」を
# 生む場所になるので、黙って足せない形にしておく。
ONACTION_GUARD_ALLOWLIST = {
    # --- 恒久例外: 押しても業務ロジックを一切呼ばない、印/表示だけのハンドラ ---
    # 深掘りチップの[×]。モジュール変数を False にして印を消すだけ。
    "modAppAct.OnFollowupChipOff",
    # 質問例の「別の質問を見る」。表示位置(mOffset)を進めて描き直すだけ。
    "modStarter.OnMore",
    # 共有ダイアログのチェックボックス。選択状態を持ち替えて描き直すだけ。
    "modShared.OnToggle",
    # 登録フォームの[キャンセル]。入力欄を消してフォームを閉じるだけ。
    # 取込中でも「閉じられない」方が利用者を困らせる(§3-1)。
    "modVault.OnVaultCancel",
    # --- 恒久例外: 委譲先で必ず関所を通るハンドラ ---
    # 質問例のクリック。最後に modApp.OnSend を呼び、そちらが
    # BlockIfIngesting を先頭に持っている(二重に置く意味が無い)。
    "modStarter.OnPick",
    # --- 恒久例外: 取込中でも【必ず】動かなければ意味が無いハンドラ ---
    # 進捗バナー脇の中断ボタン(2026-08-04 R15-6a)。取込中にしか出ない
    # ボタンなので、BlockIfIngesting を付けたら永久に押せない。やることは
    # モジュール変数のフラグを立てて1行案内を出すだけで、業務ロジックを
    # 一切呼び返さない(再入の危険がそもそも無い)。
    "modShelfBatch.OnCancelIngest",
    # 作業用Excelボタン(2026-08-05 R16-2a)。取込バナー内(■中断の左隣)と
    # アイドル時のヘルプカードの両方から同じハンドラを指す。取込中こそ
    # 使いたい機能(別プロセスのExcelで他の仕事をする)なので、
    # BlockIfIngestingを付けたら本来の目的が果たせなくなる。やることは
    # WScript.Shellで別プロセスを起動しStatusBarへ短文を出すだけで、
    # 業務ロジック(取込・検索・回答生成)を一切呼び返さない。
    "modWorkExcel.OnOpenWorkExcel",
    # 2026-08-04(R15-2b)時点で現状追認としてここに載せていた5本
    # (modVaultGallery.OnVaultBackToChat/OnVaultCardClick/OnVaultNext/
    # OnVaultPrev, modPeek.OnOpenSource)は、R15波1の敵対的自己点検で
    # 司令塔裁定「保護する」を受け、R15波1bで各ハンドラの先頭に
    # BlockIfIngestingを追加し保護済み。登録は削除した。
}

ONACTION_GUARD_CALLS = (
    re.compile(r"modUiLock\s*\.\s*BlockIfIngesting", re.IGNORECASE),
    re.compile(r"modUiLock\s*\.\s*Enter", re.IGNORECASE),
)


def _proc_body_statements(info: ModuleInfo, proc_name: str) -> list:
    """指定Publicプロシージャの本体の実行文(コメント除去済み)を順に返す。"""
    body: list = []
    started = False
    head = re.compile(
        r"^Public\s+(?:Static\s+)?(?:Sub|Function)\s+%s\b" % re.escape(proc_name),
        re.IGNORECASE)
    end = re.compile(r"^End\s+(?:Sub|Function)\b", re.IGNORECASE)
    for lineno, stmt in info.statements:
        if not started:
            if head.match(stmt):
                started = True
            continue
        if end.match(stmt):
            break
        body.append((lineno, stmt))
    return body


def check_onaction_handler_guard(infos: list[ModuleInfo]) -> None:
    # 1) 配線先の収集(src/全体)。同じ宛先が複数箇所から配線されていてもよい。
    wired: dict[str, tuple[str, int]] = {}
    for info in infos:
        for lineno, stmt in info.statements:
            m = ONACTION_LITERAL_PATTERN.search(stmt)
            if m:
                wired.setdefault(m.group(1), (module_name_for_display(info), lineno))

    by_name = {module_name_for_display(i): i for i in infos}

    # 2) 収集した宛先ごとに、先頭付近の関所を確かめる。
    for target in sorted(wired):
        if target in ONACTION_GUARD_ALLOWLIST:
            continue
        mod_name, member = target.split(".", 1)
        owner = by_name.get(mod_name)
        wire_mod, wire_line = wired[target]
        wirer = by_name.get(wire_mod)
        if owner is None:
            if wirer is not None:
                wirer.add("WARN", wire_line,
                          f"OnAction配線先のモジュールが見つかりません: {target}")
            continue
        if member not in owner.public_names:
            if wirer is not None:
                wirer.add("WARN", wire_line,
                          f"OnAction配線先が見つかりません(Publicではない/改名?): {target}")
            continue

        body = _proc_body_statements(owner, member)
        if not body:
            continue
        head = body[:ONACTION_GUARD_LOOKAHEAD]
        if any(pat.search(stmt) for _, stmt in head for pat in ONACTION_GUARD_CALLS):
            continue
        owner.add(
            "WARN", head[0][0],
            f"OnActionで配線される {target} の先頭に再入の関所がありません"
            f"(modUiLock.BlockIfIngesting か modUiLock.Enter を先頭へ。"
            f"取込中でも動くべきハンドラなら vba_lint.py の "
            f"ONACTION_GUARD_ALLOWLIST に理由つきで登録すること)",
        )


# ==============================================================================
# MsgBox/InputBoxへの非BMP文字流出(2026-08-05 R18-6c・実機第5報⑤)
# ------------------------------------------------------------------------------
# ChrW(&HD8xx)+ChrW(&HDCxx〜DFxx)で組む非BMP絵文字(🗔🩺🔄等)は、Shape/セル値
# では正しく描けるが、ネイティブMsgBox/InputBoxではサロゲート1単位ごとに
# 「?」化ける(EDGE_CASES.md §1.3b)。check_cp932_safeとは別問題(こちらは
# 実行時の描画限界であり、ソース上の直書き文字ではなくChrWで組んだ文字列が
# 対象)。
#   (A) 同一文リテラル検出: MsgBox(...)/InputBox(...)と同じ実行文の中に
#       ChrW(&HD8xx)が直書きされていないか(将来の直書き回帰の防止線)。
#   (D) 「MsgBox到達関数」許可リスト方式: FriendlyMessageのように複数の
#       消費先を持つ共通関数は、【呼ばれ方に関わらず本体全体・全Case分岐】を
#       検査しないと(A)をすり抜ける(実機第5報⑤の実バグがこの型だった)。
# 呼び出しグラフ追跡はしない。Application.Run/InvokeFeature等の文字列
# ディスパッチが中核パターンのため静的な呼び出しグラフ自体が破綻する
# (check_onaction_handler_guardと同じ誤検知ゼロ優先の設計判断)。
# ==============================================================================
SURROGATE_HIGH_PATTERN = re.compile(r"ChrW\s*\(\s*&H[Dd][89ABab][0-9A-Fa-f]{2}\s*\)")
MSGBOX_CALL_PATTERN = re.compile(r"\b(?:MsgBox|InputBox)\b", re.IGNORECASE)

# (D) 本体全体(全Case分岐)を非BMP ChrWで検査する対象。増やすときは理由を
# 1行書くこと(R18-6c 初期登録分。agent6調査報告 §2.2参照)。
MSGBOX_REACH_ALLOWLIST = (
    "modLog.FriendlyMessage",       # E0xxxコード表→複数のMsgBoxが直接表示
    "modLog.FriendlyFailMsg",       # 取込失敗の1文→MsgBox/E0805表に流れる
    "modLog.ReadOnlyWarnMsg",       # 起動時案内+E0805表の両方がMsgBoxへ
    "optOcrEta.OcrConfirmAskFor",   # OCR事前確認→modShelfVisionのMsgBoxへ
    "optOcrCache.ConfirmAskFor",    # 同上の中継(OcrConfirmAskForを包む)
    "optVision.OcrConfirmAsk",      # 同上の中継(InvokeFeature窓口)
    # 2026-08-05 R18H FB-7(B-L7): 起動時の整合性警告。どちらも modIntegrity.
    # WarnAtStartup の MsgBox へ直行する戻り値で、性質は FriendlyMessage と
    # 同型(利用者が最初に見る画面なので化けた「?」の実害は最も大きい)。
    "modIntegrity.ShrinkWarnMsg",   # 資料が減ったときの警告文→WarnAtStartupのMsgBox
    "modIntegrity.VolatileWarnMsg",  # 一時フォルダで開いている警告文→同上
)


def check_msgbox_nonbmp(infos: list[ModuleInfo]) -> None:
    # (A) 同一文リテラル検出(全モジュール対象)。
    for info in infos:
        for lineno, stmt in info.statements:
            if MSGBOX_CALL_PATTERN.search(stmt) and SURROGATE_HIGH_PATTERN.search(stmt):
                info.add(
                    "ERROR", lineno,
                    "MsgBox/InputBoxと同じ文に非BMP絵文字(ChrWのサロゲートペア)が"
                    "直書きされています。ダイアログでは1単位ごとに'?'化けます。"
                    "絵文字を落として「」括弧表記等へ置き換えてください",
                )

    # (D) 許可リスト関数の本体全体(呼ばれ方に関わらず全Case分岐)。
    by_name = {module_name_for_display(i): i for i in infos}
    for target in MSGBOX_REACH_ALLOWLIST:
        mod_name, member = target.split(".", 1)
        owner = by_name.get(mod_name)
        if owner is None or member not in owner.public_names:
            continue
        for lineno, stmt in _proc_body_statements(owner, member):
            if SURROGATE_HIGH_PATTERN.search(stmt):
                owner.add(
                    "ERROR", lineno,
                    f"MsgBox到達関数 {target} の本体に非BMP絵文字(ChrWのサロゲート"
                    f"ペア)があります。呼ばれ方に関わらずMsgBoxで'?'化けるため、"
                    f"絵文字を落として「」括弧表記等へ置き換えてください",
                )


# ==============================================================================
# Split(/Filter(/Array( の具体配列型引数への直渡し(2026-08-06 R19-2b・実機第6報②)
# ------------------------------------------------------------------------------
# Split()/Filter()/Array() はコンパイラの型システム上ただの Variant を返す式。
# ByRef arr() As String のような【具体配列型】の仮引数へ実引数位置で直渡しすると、
# 実Excel VBAは「静的に配列型と確定できる式でなければならない」制約に違反して
# プロジェクト全体のコンパイルを拒否する(modSynonymStore.bas 3箇所の実バグ)。
# LibreOffice はこの型検証をしないため素通りし、compileモードも対象モジュールの
# 中身を実行しないため二重に見逃す(調査②班報告(3)章)。
#
# シグネチャテーブル(全モジュールのSub/Function宣言。Private含む)を構築し、
# 呼び出し文の実引数境界を split_top_level で解決した上で、呼び先の該当仮引数が
# 「Variant以外の具体配列型」の場合だけERRORにする。誤検知ゼロを最優先する
# 設計(検査13/14と同じ思想):
#   ・対象は「実引数の【全体】が Split(/Filter(/Array( ...) である」場合のみ
#     (Split(...)(0) のような即時インデックスは対象外)。
#   ・代入文(x = Split(...))・If/For等の制御構文・宣言文はそもそも対象にしない
#     (先頭が識別子の呼び出し形の文だけを見る。トップレベルの"="があれば
#     代入とみなして除外する。":="という名前付き引数は代入と誤認しない)。
#   ・呼び先シグネチャが解決できない場合(組込関数・シグネチャテーブルに
#     無い呼び先)は検査せず素通しする。InvokeFeature/TryRibbonRun のような
#     ByVal ... As Variant 受けは、解決した上で正しく「安全」と判定される
#     (Variant型は具体配列型ではないのでERROR対象にならない)。
# ==============================================================================
PROC_SIG_HEAD_RE = re.compile(
    r"^(?:Public|Private|Friend)?\s*(?:Static\s+)?(Sub|Function)\s+([A-Za-z_]\w*)\s*\(",
    re.IGNORECASE,
)
ARRAY_PARAM_TYPE_RE = re.compile(r"\(\s*\)\s*As\s+([A-Za-z_]\w*)", re.IGNORECASE)
ARRAY_ARG_CALL_HEAD_RE = re.compile(r"^([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?)")
ARRAY_ARG_EXPR_RE = re.compile(r"^\s*(Split|Filter|Array)\s*\(", re.IGNORECASE)

# 先頭トークンがこれらなら呼び出し文として扱わない(制御構文・宣言・代入系)。
ARRAY_ARG_CHECK_KEYWORDS = {
    "if", "elseif", "else", "end", "exit", "for", "each", "next", "while",
    "wend", "do", "loop", "until", "select", "case", "dim", "redim", "const",
    "static", "public", "private", "friend", "function", "sub", "property",
    "type", "enum", "with", "on", "resume", "goto", "gosub", "attribute",
    "option", "declare", "set", "let", "return", "stop", "erase",
    "randomize", "open", "close", "print", "input", "line", "width",
    "name", "kill", "mkdir", "rmdir", "chdir", "chdrive", "filecopy",
    "reset", "get", "put", "lock", "unlock", "debug", "err", "beep",
    "appactivate", "sendkeys", "wait", "implements", "event", "raiseevent",
    "rem", "true", "false", "nothing", "null", "me", "new",
}


def _find_matching_paren(s: str, open_idx: int) -> int:
    """s[open_idx] == '(' 前提。対応する ')' のインデックス(文字列リテラル考慮)。
    見つからなければ -1。"""
    depth = 0
    in_str = False
    for i in range(open_idx, len(s)):
        c = s[i]
        if c == '"':
            in_str = not in_str
        elif not in_str:
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
                if depth == 0:
                    return i
    return -1


def _has_toplevel_assignment(stmt: str) -> bool:
    """文字列・カッコの外にある単独の'=' (":="ではない)があれば代入とみなす。"""
    depth = 0
    in_str = False
    for i, c in enumerate(stmt):
        if c == '"':
            in_str = not in_str
        elif not in_str:
            if c == "(":
                depth += 1
            elif c == ")":
                depth = max(0, depth - 1)
            elif c == "=" and depth == 0:
                if i > 0 and stmt[i - 1] == ":":
                    continue
                return True
    return False


def _is_whole_array_expr_call(arg: str):
    """引数全体が Split(/Filter(/Array( ...) かどうか(直後の即時インデックス
    参照 Split(...)(0) は対象外)。マッチした関数名(先頭大文字化前)を返すか、
    該当しなければNone。"""
    m = ARRAY_ARG_EXPR_RE.match(arg)
    if not m:
        return None
    close_idx = _find_matching_paren(arg, m.end() - 1)
    if close_idx == -1:
        return None
    if arg[close_idx + 1:].strip() != "":
        return None
    return m.group(1)


def _build_array_arg_signatures(infos: list[ModuleInfo]) -> dict:
    """{(モジュール表示名, プロシージャ名小文字): [仮引数ごとに具体配列型か]}。"""
    sigs: dict = {}
    for info in infos:
        mod_name = module_name_for_display(info)
        for lineno, stmt in info.statements:
            m = PROC_SIG_HEAD_RE.match(stmt)
            if not m:
                continue
            proc_name = m.group(2)
            close_idx = _find_matching_paren(stmt, m.end() - 1)
            if close_idx == -1:
                continue
            param_str = stmt[m.end():close_idx]
            params = split_top_level(param_str, ",") if param_str.strip() else []
            flags = []
            for p in params:
                pm = ARRAY_PARAM_TYPE_RE.search(p)
                flags.append(bool(pm) and pm.group(1).strip().lower() != "variant")
            sigs[(mod_name, proc_name.lower())] = flags
    return sigs


def check_array_arg_variant_mismatch(infos: list[ModuleInfo]) -> None:
    sigs = _build_array_arg_signatures(infos)

    for info in infos:
        mod_name = module_name_for_display(info)
        for lineno, stmt in info.statements:
            head_m = ARRAY_ARG_CALL_HEAD_RE.match(stmt)
            if not head_m:
                continue
            head = head_m.group(1)
            if head.split(".", 1)[0].lower() in ARRAY_ARG_CHECK_KEYWORDS:
                continue
            if _has_toplevel_assignment(stmt):
                continue

            after = stmt[head_m.end():]
            after_lstrip = after.lstrip()
            if after_lstrip.startswith("("):
                open_idx = head_m.end() + (len(after) - len(after_lstrip))
                close_idx = _find_matching_paren(stmt, open_idx)
                if close_idx == -1:
                    continue
                if stmt[close_idx + 1:].strip() != "":
                    # Foo(...).Bar のような式の一部。呼び出し文全体ではない。
                    continue
                arg_str = stmt[open_idx + 1:close_idx]
            elif after_lstrip:
                arg_str = after_lstrip
            else:
                continue  # 引数無しの単独呼び出し

            args = split_top_level(arg_str, ",") if arg_str.strip() else []
            if not args:
                continue

            if "." in head:
                target_mod, proc_name = head.split(".", 1)
                sig = sigs.get((target_mod, proc_name.lower()))
            else:
                proc_name = head
                sig = sigs.get((mod_name, head.lower()))
                if sig is None:
                    candidates = []
                    for other in infos:
                        if other is info:
                            continue
                        if head in other.public_names and other.public_names[head] in ("Sub", "Function"):
                            s = sigs.get((module_name_for_display(other), head.lower()))
                            if s is not None:
                                candidates.append(s)
                    if len(candidates) == 1:
                        sig = candidates[0]

            if sig is None:
                continue  # 呼び先未解決(組込関数・外部)は素通し

            for idx, arg in enumerate(args):
                arg_s = arg.strip()
                if not arg_s or idx >= len(sig) or not sig[idx]:
                    continue
                fn = _is_whole_array_expr_call(arg_s)
                if fn is None:
                    continue
                info.add(
                    "ERROR", lineno,
                    f"{head} の第{idx + 1}引数が具体配列型(As T())なのに、"
                    f"Variant配列を返す {fn}() を実引数位置に直接渡しています"
                    f"(実Excelはコンパイル拒否。いったん Dim x() As T: x = {fn}(...) "
                    f"で受けてから渡すこと)",
                )


def check_module_level_refs(infos: list[ModuleInfo]) -> None:
    """他モジュールのモジュールレベル定数・変数を、宣言せずに参照していないか。

    2026-07-28 の実機事故: modUIMain から AddButton を modUIMainShape へ
    切り出したとき、そこで使っている定数 COLOR_UNSELECTED_BG を元のモジュール
    へ置いたままにした。実機Excelでは Option Explicit により
    「変数が定義されていません」というコンパイルエラーになり、
    アプリが起動しなくなる。

    ところが、これは vba_lint も LibreOffice のモジュール読み込みも素通りした
    (LibreOffice は Option VBASupport 下で未定義参照を実行時まで遅延する)。
    つまり【実機で開いて初めて全機能停止】という最悪の壊れ方になる。
    同じ切り出し作業で4モジュールが同時にこの状態だった。

    誤検出を避けるため、見るのは「モジュールレベルの宣言名」だけに限る。
    ローカル変数・引数・プロシージャ名は対象にしない(名前の重複が普通に
    起きるため。そこまで見ると警告だらけになって検査自体が無視される)。
    """
    owners: dict[str, set[str]] = {}
    declared: dict[str, set[str]] = {}
    for info in infos:
        names = collect_module_level_names(info.raw_text)
        declared[info.vb_name] = {n.lower() for n in names}
        for n in names:
            owners.setdefault(n, set()).add(info.vb_name)

    for info in infos:
        mine = declared.get(info.vb_name, set())
        for lineno, stmt in info.statements:
            for ident in set(re.findall(r"(?<![\w.])([A-Za-z_]\w*)", stmt)):
                own = owners.get(ident)
                if not own:
                    continue
                if ident.lower() in mine or info.vb_name in own:
                    continue
                info.add(
                    "ERROR", lineno,
                    f"「{ident}」は {'/'.join(sorted(own))} のモジュールレベル宣言で、"
                    f"{info.vb_name} では宣言されていません"
                    f"(実機Excelで『変数が定義されていません』のコンパイルエラーになります)",
                )


PROC_DEF_RE = re.compile(
    r"^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.IGNORECASE)
LABEL_DEF_RE = re.compile(r"^\s*([A-Za-z_]\w*):\s*$")
JUMP_RE = re.compile(r"\b(GoTo|GoSub|Resume)\s+\w+", re.IGNORECASE)


def _strip_strings(line: str) -> str:
    """文字列リテラルを空白へ潰す(コメントは strip_comment 済み前提)。

    "modApp.OnActBad" のような OnAction 文字列を識別子と誤認しないため。
    """
    out = []
    in_s = False
    for ch in line:
        if ch == '"':
            in_s = not in_s
            out.append(" ")
        else:
            out.append(" " if in_s else ch)
    return "".join(out)


def check_undefined_proc_refs(infos: list[ModuleInfo]) -> None:
    """他モジュールにしか定義が無いプロシージャを、修飾なしで呼んでいないか。

    2026-07-28 の実機事故(2件目): modShelf から modShelfStore を切り出したとき、
    切り出した側が使う GetSheet が範囲の外にあり、持ってくるのを忘れた。
    実機Excelでは「Sub または Function が定義されていません」になり、
    【資料の取込が全滅】した。modP2PIo・modAppState でも同じことが起きていた。

    モジュールを分けるとき、動かすのはコードだけでは足りない。
    そのコードが呼んでいるものも一緒に動かす必要がある。
    人間の目視で担保するのは無理なので機械で見る。

    誤検出を避けるための限定:
      ・見るのは「このプロジェクトのどこかのモジュールが定義している名前」だけ。
        VBA/Excel の組み込み関数は対象外(名前を列挙しきれないため)。
      ・行ラベル(Fail: / NextRow: 等)と GoTo/GoSub/Resume の飛び先は除外。
      ・文字列リテラルの中身は除外(OnAction に入れるマクロ名を拾わないため)。
      ・宣言行そのものは除外。
    """
    proc_owners: dict[str, set[str]] = {}
    proc_defs: dict[str, set[str]] = {}
    labels: dict[str, set[str]] = {}
    for info in infos:
        names, labs = set(), set()
        for raw in info.raw_text.split("\n"):
            m = PROC_DEF_RE.match(raw)
            if m:
                names.add(m.group(1))
            lm = LABEL_DEF_RE.match(strip_comment(raw))
            if lm:
                labs.add(lm.group(1))
        proc_defs[info.vb_name] = {n.lower() for n in names}
        labels[info.vb_name] = {n.lower() for n in labs}
        for n in names:
            proc_owners.setdefault(n, set()).add(info.vb_name)

    for info in infos:
        mine = proc_defs.get(info.vb_name, set())
        labs = labels.get(info.vb_name, set())
        for lineno, raw in merge_continuations(info.raw_text.split("\n")):
            code = _strip_strings(strip_comment(raw))
            if not code.strip():
                continue
            if PROC_DEF_RE.match(code):
                continue
            if LABEL_DEF_RE.match(code):
                continue
            code = JUMP_RE.sub(" ", code)
            for ident in set(re.findall(r"(?<![\w.])([A-Za-z_]\w*)", code)):
                own = proc_owners.get(ident)
                if not own:
                    continue
                if ident.lower() in mine or ident.lower() in labs:
                    continue
                if info.vb_name in own:
                    continue
                info.add(
                    "ERROR", lineno,
                    f"「{ident}」は {'/'.join(sorted(own))} でしか定義されていません。"
                    f"{info.vb_name} には無いので、修飾して呼ぶか、この モジュールへ持ってくること"
                    f"(実機Excelで『Sub または Function が定義されていません』になります)",
                )


ON_ERROR_GOTO_LABEL_RE = re.compile(r"^\s*On\s+Error\s+GoTo\s+([A-Za-z_]\w*)\s*$", re.IGNORECASE)
PROC_END_RE = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)
EXIT_PROC_RE = re.compile(r"^\s*(?:Exit\s+(?:Sub|Function|Property)|End)\s*$", re.IGNORECASE)
RESUME_RE = re.compile(r"^\s*Resume\b", re.IGNORECASE)
ERR_RAISE_RE = re.compile(r"\bErr\.Raise\b", re.IGNORECASE)


def check_handler_exit(info: ModuleInfo) -> None:
    """エラーハンドラから通常フローへ戻るとき Resume を使っているか。

    2026-07-29 に自分で踏んだ。VBAは「エラーハンドラ実行中」という状態を持ち、
    これを解除できるのは Resume と、プロシージャを抜けることだけである。
    On Error GoTo 0 はトラップの登録を消すだけで、この状態は解除しない。

    そのため

        For p = ...
            On Error GoTo SkipPage
            ...処理...
            GoTo NextPage
    SkipPage:
            skipped = skipped + 1
            On Error GoTo 0      ' ← これでは抜けられない
    NextPage:
        Next p

    と書くと、【1ページ目の失敗は拾えるが、2ページ目の失敗は SkipPage へ
    飛ばずに呼び出し元へ突き抜ける】。「1ページの失敗で資料全体を失わない」
    ための修正が、2ページ目以降では効かない。しかもテストでは
    1件しか壊さないので気づけない。正しくは `Resume NextPage`。

    検出方法: On Error GoTo <L> の飛び先ラベル <L> のブロック
    (ラベル行から、次のラベル行またはプロシージャ末尾まで)を見て、
    そのブロックがプロシージャ末尾に到達せず、かつ Resume / Exit /
    Err.Raise のいずれも含まないものを WARN にする。
    """
    lines = [(n, strip_comment(s)) for n, s in merge_continuations(info.raw_text.split("\n"))]

    # プロシージャ単位に切る
    proc_start = None
    for idx, (lineno, code) in enumerate(lines):
        if PROC_DEF_RE.match(code):
            proc_start = idx
            continue
        if PROC_END_RE.match(code) and proc_start is not None:
            _check_handler_exit_in_proc(info, lines, proc_start, idx)
            proc_start = None


def _check_handler_exit_in_proc(info: ModuleInfo, lines, start: int, end: int) -> None:
    targets = set()
    for _, code in lines[start:end]:
        m = ON_ERROR_GOTO_LABEL_RE.match(code)
        if m and m.group(1) != "0":
            targets.add(m.group(1).lower())
    if not targets:
        return

    label_at = {}
    for idx in range(start, end):
        m = LABEL_DEF_RE.match(lines[idx][1])
        if m:
            label_at[m.group(1).lower()] = idx

    for name in sorted(targets):
        idx = label_at.get(name)
        if idx is None:
            continue
        # ブロック = ラベル行の次から、次のラベル行 or プロシージャ末尾まで
        stop = end
        for j in range(idx + 1, end):
            if LABEL_DEF_RE.match(lines[j][1]):
                stop = j
                break
        if stop == end:
            continue  # プロシージャ末尾まで届く = 抜けるので問題ない
        body = [lines[j][1] for j in range(idx + 1, stop)]
        if any(RESUME_RE.match(c) or EXIT_PROC_RE.match(c) or ERR_RAISE_RE.search(c) for c in body):
            continue
        info.add(
            "WARN", lines[idx][0],
            f"エラーハンドラ「{name}:」が Resume / Exit を使わずに次のラベルへ落ちている。"
            f"VBAは On Error GoTo 0 ではハンドラ実行中の状態を解除しないため、"
            f"2回目以降のエラーが素通りする(`Resume <ラベル>` で抜けること)",
        )


ON_ERROR_RESUME_NEXT_RE = re.compile(r"^\s*On\s+Error\s+Resume\s+Next\s*$", re.IGNORECASE)
TERMINATOR_RE = re.compile(
    r"^\s*(Exit\s+(Sub|Function|Property)\b|GoTo\s+\w+\s*$|Resume\b|End\s*$)", re.IGNORECASE)


def check_resume_on_fallthrough_label(info: ModuleInfo) -> None:
    """正常系からも落ちてくるラベルの中で Resume を使っていないか。

    Resume は「エラーが起きている」ことが前提の命令で、エラーなしで実行すると
    実行時エラー20「Resume にエラーがありません」になる。

        Exit Sub は書かず、正常系がそのまま Finish: へ落ちる作り
    Finish:
        Resume FinishCleanup    ' ← 正常終了時にここを踏んで即クラッシュ

    ハンドラ兼後始末ラベル(Finish: / Done: / Cleanup: など、Exit を挟まずに
    正常系から流れ込む形)は本コードベースに多数あるため、
    「ハンドラを Resume で抜ける」修正を入れるときに必ず踏む罠。
    正しくはラベルの直前へ GoTo を置いて、正常系が Resume を跨ぐようにする。
    """
    lines = [(n, strip_comment(s)) for n, s in merge_continuations(info.raw_text.split("\n"))]

    proc_start = None
    for idx, (_lineno, code) in enumerate(lines):
        if PROC_DEF_RE.match(code):
            proc_start = idx
            continue
        if PROC_END_RE.match(code) and proc_start is not None:
            _check_fallthrough_resume_in_proc(info, lines, proc_start, idx)
            proc_start = None


def _check_fallthrough_resume_in_proc(info: ModuleInfo, lines, start: int, end: int) -> None:
    for idx in range(start + 1, end):
        if not LABEL_DEF_RE.match(lines[idx][1]):
            continue
        prev = next((lines[j][1] for j in range(idx - 1, start, -1) if lines[j][1].strip()), "")
        if TERMINATOR_RE.match(prev):
            continue          # 正常系はここへ落ちてこない
        for j in range(idx + 1, end):
            code = lines[j][1]
            if LABEL_DEF_RE.match(code):
                break
            if RESUME_RE.match(code):
                info.add(
                    "ERROR", lines[j][0],
                    f"ラベル「{LABEL_DEF_RE.match(lines[idx][1]).group(1)}:」は正常系からも"
                    f"落ちてくるのに Resume がある。エラーなしで Resume を実行すると"
                    f"実行時エラー20になる(ラベルの直前へ GoTo を置き、"
                    f"正常系が Resume を跨ぐようにすること)",
                )
                break


def check_resume_next_inside_handler(info: ModuleInfo) -> None:
    """稼働中のエラーハンドラの内側に On Error Resume Next を書いていないか。

    2026-07-30 実機: .doc/.docx/.pdf の取込が err#462 で全滅した件の真因。

    VBAには「有効(enabled)なハンドラ」と「稼働中(active)なハンドラ」の区別が
    ある。ハンドラへ飛んだ瞬間から Resume / Exit / プロシージャ終了までは
    「稼働中」で、この間に起きたエラーは【そのプロシージャでは一切捕まえられず、
    呼び出し元へ投げ返される】。On Error Resume Next を書いてあっても効かない。

    modExtractorWord.TryExtractOnce はこうなっていた。

        On Error GoTo Failed
        Set word = CreateObject("Word.Application")
        word.Visible = False          ' ← ここで実機は死んでいた
        ...
    Failed:
        errDetail = 実際の原因        ' ← 正しく作っている
        If Not word Is Nothing Then
            On Error Resume Next      ' ← 稼働中なので効かない
            word.Quit 0               ' ← 死んだWordを叩いて err#462
            On Error GoTo 0
        End If
        TryExtractOnce = False        ' ← ここへ到達しない

    結果として
      ・後始末の err#462 が呼び出し元(modExtractor.ExtractFile)まで飛び、
        本当の原因を上書きしたエラーだけがログに残る
      ・Extract の「開き方3通り」フォールバックが一度も走らない
      ・PDFの Acrobat フォールバックも一度も走らない
    つまり、ログに見えている 462 は原因ではなく【後始末の失敗】であり、
    フォールバックは全部ダメコードだった。ハンドラの中は素手で歩けない。

    正しい書き方はどちらか。
      ・後始末を別Subへ切り出す(呼ばれた側は新しいエラー文脈を持つので
        On Error Resume Next が効く)
      ・`Resume <ラベル>` でハンドラを抜けてから後始末する
    """
    lines = [(n, strip_comment(s)) for n, s in merge_continuations(info.raw_text.split("\n"))]

    proc_start = None
    for idx, (_lineno, code) in enumerate(lines):
        if PROC_DEF_RE.match(code):
            proc_start = idx
            continue
        if PROC_END_RE.match(code) and proc_start is not None:
            _check_resume_next_in_proc(info, lines, proc_start, idx)
            proc_start = None


def _check_resume_next_in_proc(info: ModuleInfo, lines, start: int, end: int) -> None:
    targets = set()
    for _, code in lines[start:end]:
        m = ON_ERROR_GOTO_LABEL_RE.match(code)
        if m and m.group(1) != "0":
            targets.add(m.group(1).lower())
    if not targets:
        return

    for idx in range(start, end):
        m = LABEL_DEF_RE.match(lines[idx][1])
        if not m or m.group(1).lower() not in targets:
            continue
        # ハンドラ稼働区間 = ラベル行の次から、Resume / Exit / プロシージャ末尾まで
        for j in range(idx + 1, end):
            code = lines[j][1]
            if RESUME_RE.match(code) or EXIT_PROC_RE.match(code):
                break
            if LABEL_DEF_RE.match(code):
                # 次のラベルへ落ちている場合は check_handler_exit が別途警告する。
                # 稼働中の状態は解除されないまま続くので、区間はここでは切らない。
                continue
            if ON_ERROR_RESUME_NEXT_RE.match(code):
                info.add(
                    "ERROR", lines[j][0],
                    f"エラーハンドラ「{m.group(1)}:」の稼働中に On Error Resume Next を書いている。"
                    f"VBAはハンドラ稼働中のエラーを同一プロシージャでは捕捉できないため無効で、"
                    f"ここで起きたエラーは呼び出し元へ飛び、本来の原因を上書きする"
                    f"(後始末は別Subへ切り出すか、Resume でハンドラを抜けてから行うこと)",
                )


def check_dim_type_drop_and_integer(info: ModuleInfo) -> None:
    for lineno, stmt in info.statements:
        if AS_INTEGER_PATTERN.search(stmt):
            info.add("ERROR", lineno, f"Integer型は禁止(Longを使う): 「{stmt.strip()[:80]}」")

        # 2026-07-16 実機事故の恒久再発防止: ReDim x(0 To -1) 等の「上限<下限」は
        # LibreOffice Basicでは0要素配列として通るが、実機Excel VBAでは実行時
        # エラー9「インデックスが有効範囲にありません」になる(VB.NETとの混同)。
        # LO実行テストはこの文を「環境差」と誤認しスキップしていたため、本番
        # コード20箇所以上に混入し、資料取込・同期・画面描画が実機で全滅した。
        # 0件は「count変数+ReDim(0 To 0)」か「Split(vbNullString)」で表現する。
        if NEGATIVE_REDIM_PATTERN.search(stmt):
            info.add(
                "ERROR", lineno,
                f"実機Excel VBAで実行時エラー9になる負範囲ReDim(To -1): 「{stmt.strip()[:80]}」",
            )

        m = DIM_STMT_PATTERN.match(stmt)
        if not m:
            continue
        rest = m.group(2)
        segments = split_top_level(rest, ",")
        if len(segments) < 2:
            continue
        for seg in segments[:-1]:
            if not AS_KEYWORD_PATTERN.search(seg):
                info.add(
                    "ERROR", lineno,
                    f"型落ち疑い(Dim a, b As T は先頭がVariantになる): 「{stmt.strip()[:80]}」",
                )
                break


PROC_DEF_PATTERN = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.IGNORECASE,
)
LOCAL_DECL_PATTERN = re.compile(r"\b(?:Dim|Static|Const)\s+(\w+)", re.IGNORECASE)
PROC_PARAM_PATTERN = re.compile(
    r"(?:ByVal\s+|ByRef\s+|Optional\s+(?:ByVal\s+|ByRef\s+)?)(\w+)", re.IGNORECASE
)


def check_name_shadowing(info: ModuleInfo) -> None:
    """変数/引数名が同一モジュール内の手続き名と衝突していないか。

    2026-07-16 実機事故の恒久再発防止: modEmbedのローカル変数sleepMsが
    同モジュールのSub SleepMsを隠し、「SleepMs sleepMs」の呼び出し行が
    実機Excel VBAでコンパイルエラー「Sub、Functionまたは Propertyが
    必要です」になった(LibreOffice Basicは同名解決を許すため、LOの
    コンパイル検査もLO実行テストも通過してしまう)。さらにVBAは遅延
    コンパイルのため、当該モジュールに実行が到達した瞬間に初めて爆発し、
    エラーハンドラも走らず再入ガードが焼き付く二次被害まで起きた。
    呼び出しの有無にかかわらず、同名宣言そのものを全面禁止にする。
    """
    procs: dict[str, int] = {}
    for lineno, stmt in info.statements:
        m = PROC_DEF_PATTERN.match(stmt)
        if m:
            procs[m.group(1).lower()] = lineno
    if not procs:
        return
    for lineno, stmt in info.statements:
        for m in LOCAL_DECL_PATTERN.finditer(stmt):
            name = m.group(1)
            if name.lower() in procs:
                info.add(
                    "ERROR", lineno,
                    f"変数名「{name}」が同一モジュール内の手続き名と衝突"
                    f"(実機VBAでコンパイルエラーの元): 「{stmt.strip()[:80]}」",
                )
        if PROC_DEF_PATTERN.match(stmt):
            for m in PROC_PARAM_PATTERN.finditer(stmt):
                name = m.group(1)
                if name.lower() in procs and procs[name.lower()] != lineno:
                    info.add(
                        "ERROR", lineno,
                        f"引数名「{name}」が同一モジュール内の手続き名と衝突"
                        f"(実機VBAでコンパイルエラーの元): 「{stmt.strip()[:80]}」",
                    )


MODULE_DECL_PATTERN = re.compile(
    r"^\s*(?:(?:Public|Private|Global)\s+)?(?:Const\s+\w|Declare\s)"
    r"|^\s*(?:Public|Private|Global|Dim)\s+\w+\s*(?:\(\s*\))?\s+As\s+",
    re.IGNORECASE,
)


def check_declaration_position(info: ModuleInfo) -> None:
    """モジュールレベル宣言(Const/変数/Declare)が最初のプロシージャ定義より
    後に無いか。VBAは宣言部→プロシージャ部の順序を強制し、違反すると
    「End Sub、End Function、または End Property の後には、コメントのみが
    記述できます」というコンパイルエラーになる(2026-07-16実機で発生。
    LibreOffice Basicは途中宣言を許容するためLOゲートでは検出不能)。"""
    depth = 0
    seen_proc = False
    for lineno, stmt in info.statements:
        if PROC_DEF_PATTERN.match(stmt):
            depth += 1
            seen_proc = True
            continue
        if re.match(r"^\s*End\s+(Sub|Function|Property)\b", stmt, re.IGNORECASE):
            depth -= 1
            continue
        if depth <= 0 and seen_proc and MODULE_DECL_PATTERN.match(stmt):
            info.add(
                "ERROR", lineno,
                f"モジュールレベル宣言がプロシージャ定義より後にある"
                f"(実機VBAでコンパイルエラー。宣言部はモジュール先頭へ): 「{stmt.strip()[:80]}」",
            )


RESERVED_IDENT_PATTERN = re.compile(
    r"\b(?:Dim|Const|Static|ByVal|ByRef)\s+base\b"
    r"|[(,]\s*base\s+As\b",
    re.IGNORECASE,
)


def check_reserved_identifiers(info: ModuleInfo) -> None:
    """StarBasic予約語(Option Base の 'Base')をVBAの識別子として使うと、
    LibreOffice構文チェック(LOゲート)がコンパイルダイアログでサイレントに
    ハングし、「タイムアウト=構文エラーの疑い」としてしか現れず原因特定に
    多大な時間を要する(2026-07-17 modP2P.ThanksDirで発生・二分探索で特定)。
    実機Excelでは 'base' は有効な変数名だが、ツールチェーンの沈黙ハングを
    防ぐため恒久ガードとして識別子 'base' の宣言/仮引数を禁止する
    (別名 basePath 等にする)。line/name は既存コードで実証上ハングしないため
    対象外(誤検知回避)。新たにハングする予約語が見つかったらここへ追記する。"""
    for lineno, stmt in info.statements:
        if RESERVED_IDENT_PATTERN.search(stmt):
            info.add(
                "ERROR", lineno,
                "StarBasic予約語 'base' を識別子に使用(LO構文チェックが沈黙ハング)。"
                f"別名(basePath等)にしてください: 「{stmt.strip()[:80]}」",
            )


# VBA_RESERVED_BLOCKLIST: 2026-07-21 実機事故の恒久再発防止。src/ui/modApp.bas の
# 「Dim fix As String」が実機Windows Excelで「コンパイルエラー: 構文エラー」に
# なり、プロジェクト全体が未コンパイル状態に陥った結果、起動時のShapes.AddShape/
# OnAction割当てが無関係な「実行時エラー1004」として表面化し、原因特定に
# 長時間を要した(実機写真で現物確認・IMG_4836)。LibreOffice Basicの構文
# チェック(LOゲート)は 'Fix' を予約語として扱わずコンパイルを通してしまうため、
# 本リポジトリの二段階検証(lint+LOテスト)を両方すり抜けていた。
# この事故クラスを二度と実機まで持ち越さないため、VBA文法キーワードに加えて
# VBA/Excelの組み込み関数名も総ざらいでブロックリスト化する(一部は実際には
# 識別子として使えるものも含むが、実機コンパイラでの可否を出力側からは検証
# できないため、コストゼロの防御としてまとめて禁止する)。
VBA_RESERVED_BLOCKLIST = {
    w.lower() for w in (
        # 文法キーワード(全ダイアレクトで確実に予約語)
        "Dim Static Const Public Private Friend As ByVal ByRef Optional ParamArray "
        "Sub Function Property Get Let Set End If Then Else ElseIf Select Case "
        "For Each In To Step Next Do While Wend Until Loop With Exit GoTo GoSub "
        "Return On Error Resume Call New Nothing Is Like Mod And Or Not Xor Eqv Imp "
        "True False Null Empty Me Option Explicit Compare Type Enum Declare "
        "Lib Alias Implements WithEvents Event RaiseEvent Class Attribute Rem "
        "ReDim Preserve Erase Stop Debug Variant Boolean Byte Integer Long "
        "LongLong LongPtr Single Double Currency Decimal Date String Object "
        # VBA/Excel組み込み関数名(識別子としての衝突が実機コンパイルエラーの
        # 原因になりうるため予防的に全面禁止。'base'は別枠でcheck_reserved_
        # identifiersが担当するためここには含めない)
        "Abs Array Asc AscB AscW Atn CBool CByte CCur CDate CDbl CDec Chr ChrB ChrW "
        "CInt CLng CLngLng CLngPtr Cos CSng CStr CurDir CVar CVDate CVErr "
        "DateAdd DateDiff DatePart DateSerial DateValue Day DDB Dir DoEvents Environ "
        "EOF Exp FileAttr FileDateTime FileLen Filter Fix Format "
        "FormatCurrency FormatDateTime FormatNumber FormatPercent FreeFile FV "
        "GetAllSettings GetAttr GetObject GetSetting Hex Hour IIf IMEStatus Input "
        "InputB InputBox InStr InStrB InStrRev Int IPmt IRR IsArray IsDate IsEmpty "
        "IsError IsMissing IsNull IsNumeric IsObject Join LBound LCase Left LeftB "
        "Len LenB LoadPicture Loc LOF Log LTrim Mid MidB Minute MIRR MkDir Month "
        "MonthName MsgBox Now NPer NPV Oct Partition Pmt PPmt PV QBColor Rate RGB "
        "Right RightB RmDir Rnd Round RTrim Second Seek Sgn Shell Sin SLN Space Spc "
        "Split Sqr Str StrComp StrConv StrReverse Switch SYD Tab Tan Time "
        "Timer TimeSerial TimeValue Trim TypeName UBound UCase Val VarType Weekday "
        "WeekdayName Year Name Kill Width Height Top"
    ).split()
}

# Dim/Private/Public/Static/ReDim(Preserve)の変数宣言、およびByVal/ByRefの
# 仮引数宣言を横断的に検出する(check_name_shadowingのLOCAL_DECL_PATTERN/
# PROC_PARAM_PATTERNは目的が異なる限定用途のため、ここでは独自に定義する)。
RESERVED_WORD_DECL_PATTERN = re.compile(
    r"\b(?:Dim|Private|Public|Static|ReDim(?:\s+Preserve)?)\s+([A-Za-z_]\w*)\s+As\b"
    r"|\b(?:ByVal|ByRef)\s+([A-Za-z_]\w*)\s+As\b",
    re.IGNORECASE,
)


def check_vba_reserved_words(info: ModuleInfo) -> None:
    """変数/引数名がVBA文法キーワードまたは組み込み関数名と衝突していないか
    (2026-07-21実機事故の恒久ガード。詳細はVBA_RESERVED_BLOCKLIST直上のコメント
    参照)。'base'単体は既存のcheck_reserved_identifiersが別途担当する。"""
    for lineno, stmt in info.statements:
        for m in RESERVED_WORD_DECL_PATTERN.finditer(stmt):
            name = m.group(1) or m.group(2)
            if name.lower() == "base":
                continue   # check_reserved_identifiers側で専用メッセージを出す
            if name.lower() in VBA_RESERVED_BLOCKLIST:
                info.add(
                    "ERROR", lineno,
                    f"識別子「{name}」がVBAの予約語/組み込み関数名と衝突"
                    f"(実機VBAでコンパイルエラー「構文エラー」の元。2026-07-21事故と同型): "
                    f"「{stmt.strip()[:80]}」",
                )


def module_name_for_display(info: ModuleInfo) -> str:
    return info.vb_name or info.filename_stem


def check_pure_logic_tokens(info: ModuleInfo) -> None:
    name = module_name_for_display(info)
    if name not in PURE_LOGIC_MODULES:
        return
    for lineno, stmt in info.statements:
        for pattern, label in FORBIDDEN_TOKEN_PATTERNS:
            if pattern.search(stmt):
                info.add(
                    "ERROR", lineno,
                    f"R4違反: 純ロジックモジュールで禁止トークン「{label}」使用: 「{stmt.strip()[:80]}」",
                )


def check_opt_token_reference(info: ModuleInfo) -> None:
    if info.layer == LAYER_OPT:
        return  # src/opt自身はoptトークンを書いてよい
    # テスト層も対象外(2026-07-31 R6)。R2の禁止対象は「コアのどのモジュール」で、
    # 目的は【opt機能を1行削除で撤去してもコアが壊れないこと】。テストモジュールは
    # 製品に同梱されるコアではなく、opt層の純ロジック(optOcrCoreのGSコマンド
    # 組み立て等)を直接検証できないと、一番壊れやすい引用符の付け方を実行テストで
    # 固定できない。同じ理由で check_layer_dependency も既にテスト層を除外している。
    if info.layer == LAYER_TEST:
        return
    for lineno, stmt in info.statements:
        m = OPT_TOKEN_PATTERN.search(stmt)
        if m:
            info.add(
                "ERROR", lineno,
                f"R2違反: opt直接参照「{m.group(0).strip()}」。"
                f"modFeatures.InvokeFeature経由にすること: 「{stmt.strip()[:80]}」",
            )


APPLICATION_RUN_PATTERN = re.compile(
    r"Application\s*\.\s*Run\s*\(?\s*(?:\"(?P<lit>[^\"]*)\"|(?P<var>[A-Za-z_]\w*))"
)


def check_application_run_whitelist(info: ModuleInfo) -> None:
    name = module_name_for_display(info)
    for lineno, stmt in info.statements:
        for m in APPLICATION_RUN_PATTERN.finditer(stmt):
            lit = m.group("lit")
            var = m.group("var")
            if lit is not None:
                allowed = (
                    lit in RUN_LITERAL_WHITELIST_EXACT
                    or RUN_LITERAL_WHITELIST_PREFIX.match(lit) is not None
                )
                if not allowed:
                    info.add(
                        "ERROR", lineno,
                        f'R3違反: Application.Run("{lit}", ...) はホワイトリスト外: 「{stmt.strip()[:80]}」',
                    )
                    continue
                if lit in RUN_LITERAL_GATEWAY_ONLY and name != "modGateway":
                    info.add(
                        "ERROR", lineno,
                        f'R3違反: Application.Run("{lit}", ...) はmodGateway内のみ許可'
                        f'(このモジュールは{name}): 「{stmt.strip()[:80]}」',
                    )
            else:
                # 変数/式経由のApplication.Run
                if name not in RUN_VARIABLE_ALLOWED_MODULES:
                    info.add(
                        "ERROR", lineno,
                        f"R3違反: 変数経由のApplication.Run({var}, ...)は"
                        f"modGateway/modFeatures以外で禁止(このモジュールは{name}): "
                        f"「{stmt.strip()[:80]}」",
                    )


def check_cross_module_references(info: ModuleInfo, known_modules: dict[str, ModuleInfo]) -> None:
    self_name = module_name_for_display(info)
    seen_skip: set[str] = set()
    for lineno, stmt in info.statements:
        for m in DOTTED_REF_PATTERN.finditer(stmt):
            prefix, member = m.group(1), m.group(2)
            if not MODULE_SHAPED_NAME.match(prefix):
                continue
            target = known_modules.get(prefix)
            if target is None:
                if prefix != self_name and prefix not in seen_skip:
                    seen_skip.add(prefix)
                    info.add(
                        "SKIP", lineno,
                        f"未実装モジュール参照のためスキップ: {prefix}.{member}",
                    )
                continue
            if member not in target.public_names:
                info.add(
                    "ERROR", lineno,
                    f"モジュール間参照エラー: {prefix}.{member} はPublicとして実在しない"
                    f"(現在の{prefix}のPublic一覧: {sorted(target.public_names) or 'なし'})",
                )


def check_layer_dependency(info: ModuleInfo, known_modules: dict[str, ModuleInfo]) -> None:
    cur_layer = info.layer
    if cur_layer is None or cur_layer == LAYER_TEST:
        return  # 分類不能・テスト層は依存順序の対象外

    self_name = module_name_for_display(info)

    for lineno, stmt in info.statements:
        for m in DOTTED_REF_PATTERN.finditer(stmt):
            prefix, member = m.group(1), m.group(2)
            if not MODULE_SHAPED_NAME.match(prefix):
                continue
            target = known_modules.get(prefix)
            if target is None or target.layer is None:
                continue
            if prefix == self_name:
                continue

            if cur_layer == LAYER_OPT:
                if target.layer == LAYER_FOUNDATION:
                    continue
                if prefix == "modUIMain" and member == "SetStage":
                    continue
                # R13-1d(2026-08-03): opt層からの段階バナー更新。
                # SetStage の出力先(状態行/StatusBar/チャットバブル)はNexus
                # 画面では実質不可視で、GSの本文抽出を数分待つあいだ利用者には
                # 何も見えなかった(実機第2報 RC5・憲章§3-2)。modShelfBatch.
                # StageBanner は「進捗バナーが既に出ているときだけ更新する」
                # 副作用の閉じた通知コールバックで、SetStage と同じ性質
                # (実況を伝えるだけ・業務ロジックを一切呼び返さない)。
                # 広げるときは必ずここへ足す=どのoptが画面へ触れるかが
                # 1箇所で分かる状態を保つ。
                if prefix == "modShelfBatch" and member == "StageBanner":
                    continue
                # R15-6b(2026-08-04): opt層からの中断の問い合わせ。
                # 進捗バナーの中断ボタンが立てる印を optOcrPage が頁境界・
                # バッチ境界で【読むだけ】の関数で、引数も戻り値も Boolean 1つ。
                # 業務ロジックを呼び返さず、状態を1ビットも書き換えない
                # (StageBanner よりさらに副作用が小さい)。印の実体を
                # 取込の入口(AddFilesResult)でリセットする都合上、置き場は
                # ingest 層である必要がある。広げるときは必ずここへ足す
                # =どのoptがコアへ触れるかが1箇所で分かる状態を保つ。
                if prefix == "modShelfBatch" and member == "CancelRequested":
                    continue
                # R15-FixA FA-2(2026-08-04 レビュー裁定 A-H2/B-H2): opt層からの
                # 中間保存。頁OCRの控え(optOcrCache)はシートへ書いた時点では
                # まだメモリ上のブックにしか無く、強制終了で丸ごと消える
                # =「続きから再開できる」という約束がその瞬間だけ嘘になる
                # (実機第4報で実際に起きた壊れ方)。保存の作法(ReadOnly判定・
                # save_fail の記録・トーストの1回きり・120秒スロットル)は
                # modShelfBatch.SaveCheckpoint が1箇所で持っており、opt層へ
                # ThisWorkbook.Save を書き写すのはその分散そのもの。
                # StageBanner/CancelRequested と同じく、広げるときはここへ足す。
                if prefix == "modShelfBatch" and member == "SaveCheckpoint":
                    continue
                if target.layer == LAYER_OPT:
                    continue
                info.add(
                    "ERROR", lineno,
                    f"opt層契約違反(§7.7): opt層からのコア参照は基盤層+"
                    f"modUIMain.SetStageのみ許可。{prefix}.{member} は"
                    f"{LAYER_LABEL.get(target.layer, target.layer)}: 「{stmt.strip()[:80]}」",
                )
                continue

            if isinstance(cur_layer, int) and isinstance(target.layer, int):
                if target.layer > cur_layer:
                    # R1例外(UI通知コールバック。MASTER_SPEC §3 R1 / Wave2 PM裁定):
                    # 機能層は処理の進捗・結果をUIへ通知するために以下のみ呼んでよい。
                    # ShowProgress/HideProgress(R10-5): 取込中の進捗バナー。
                    # SetStageと同じ「UIへ実況を伝えるだけ」の通知コールバックで、
                    # ShowProgress自体が内部でSetStageを呼ぶ薄いラッパーのため、
                    # 既存のSetStage例外と同列に扱う。
                    # SetStage は基盤層からも1箇所だけ許す(2026-07-31 R11-D /
                    # 監査3 M-7)。modGatewayDirect の埋め込みHTTPは【同期】
                    # Send で、NW瞬断時は最大60秒 Excel が完全に固まる。何の
                    # 表示も無く固まるのは憲章§3-2違反だが、待ちが発生する
                    # 場所そのものは基盤層にしかない。SetStage は opt層にも
                    # 同じ理由で既に例外が置かれている「実況を伝えるだけ」の
                    # 通知コールバックなので、同列に扱う。
                    if (cur_layer == LAYER_FOUNDATION
                            and (prefix, member) == ("modUIMain", "SetStage")
                            and self_name == "modGatewayDirect"):
                        continue
                    if cur_layer == LAYER_MID and (prefix, member) in (
                        ("modUIMain", "SetStage"),
                        ("modUIMain", "RenderSourcesPreview"),
                        ("modUIMain", "RenderAnswer"),
                        ("modUIMain", "ShowProgress"),
                        ("modUIMain", "HideProgress"),
                        ("modUIShelf", "RenderShelf"),
                    ):
                        continue
                    # ShowToast(R10-5、R10cのL5で絞り込み): 取込/同期の【完了
                    # 1回だけ】を知らせる非ブロッキング通知。ShowToastは表示に
                    # 1.1秒のブロッキング待ちを含むため、「お待ちください」等の
                    # 常用へ広がると待ち時間が積み上がる(R10cのM3で modUiLock
                    # から撤去したのがまさにその事故)。完了通知を出す2モジュール
                    # に限定し、他所へ広がったらLintで止める。
                    if (cur_layer == LAYER_MID
                            and (prefix, member) == ("modSkin", "ShowToast")
                            and self_name in R1_TOAST_ALLOWED_MODULES):
                        continue
                    # ShowToast(2026-08-05 R17H FB-7 / B-M): 基盤層からは
                    # modIntegrity.WarnAtStartup の1箇所だけ許す。起動時に
                    # 「my_knowledge はあるのに chunk_meta が0行」(=R17より前に
                    # 取り込んだ資料しか無い本棚)を1文だけ知らせるためのもので、
                    # 裁定でモーダル禁止=MsgBox を使えない。判定材料はシート2枚
                    # だけで上位層の状態を読まないため、向きは「基盤→UIへ通知」
                    # の一方通行(modGatewayDirect の SetStage 例外と同性質)。
                    # 起動時1回きり(mWarned)なので 1.1 秒の待ちも1回で済む。
                    # 広げるときは必ずここへ足す=どのモジュールが基盤層から
                    # トーストを出すかが1箇所で分かる状態を保つ。
                    if (cur_layer == LAYER_FOUNDATION
                            and (prefix, member) == ("modSkin", "ShowToast")
                            and self_name == "modIntegrity"):
                        continue
                    if (cur_layer == LAYER_MID
                            and (prefix, member) == ("modUiLock", "IsBusy")
                            and self_name in R1_UILOCK_ALLOWED_MODULES):
                        continue
                    # PaintProgress(2026-08-04 R15-FixA FA-6 / 2026-08-05 R18-1a
                    # で modSkin から modProgressBar へ移設): 中断ボタン付きの
                    # 進捗バナー。ShowProgress(=SetStage+PaintProgress)と同じ
                    # 「実況を伝えるだけ」の通知コールバックで、違いは中断ボタンを
                    # 添えるかどうかの1点だけ。modUIMain 側で分岐できれば
                    # ShowProgress の既存例外で済むが、あちらは30,000字上限まで
                    # 残り206字で引数1つ足す余地が無い(憲章§4-6)。取込の
                    # バナーを1本化する modShelfBatch.ShowIngestBanner だけに許す
                    # =どのモジュールが中断できるバナーを出せるかが1箇所で分かる。
                    if (cur_layer == LAYER_MID
                            and (prefix, member) == ("modProgressBar", "PaintProgress")
                            and self_name in R1_PROGRESS_ALLOWED_MODULES):
                        continue
                    # OpenWorkExcelNow(2026-08-05 R18-1f): 取込前確認の2段目から
                    # 作業用Excelを先に開く。理由は R1_WORKEXCEL_ALLOWED_MODULES の
                    # 注記を参照。
                    if (cur_layer == LAYER_MID
                            and prefix == "modWorkExcel"
                            and member in R1_WORKEXCEL_ALLOWED_MEMBERS
                            and self_name in R1_WORKEXCEL_ALLOWED_MODULES):
                        continue
                    info.add(
                        "ERROR", lineno,
                        f"R1違反: {LAYER_LABEL[cur_layer]}から上位の"
                        f"{LAYER_LABEL[target.layer]}を参照: {prefix}.{member}"
                        f"「{stmt.strip()[:80]}」",
                    )


def check_contract(info: ModuleInfo) -> None:
    name = module_name_for_display(info)
    contract = CONTRACT.get(name)
    if contract is None:
        return
    required = set(contract["required"])
    actual = set(info.public_names.keys())

    missing = required - actual
    for nm in sorted(missing):
        info.add("ERROR", 1, f"契約違反: Public {nm} が契約にあるが実装されていない")

    if contract["closed"]:
        extra = actual - required
        for nm in sorted(extra):
            info.add(
                "ERROR", 1,
                f"契約違反: Public {nm} は契約に無い(§7に無いPublicは作らない規約)",
            )


def check_safeleft_warning(info: ModuleInfo) -> None:
    for lineno, stmt in info.statements:
        if not FULLTEXT_ASSIGN_PATTERN.search(stmt):
            continue
        if not FULLTEXT_HINT_PATTERN.search(stmt):
            continue
        if SAFELEFT_CALL_PATTERN.search(stmt):
            continue
        info.add(
            "WARN", lineno,
            f"full_text系のセル書込みでSafeLeft経由が確認できません"
            f"(32,767字超で書込み失敗の恐れ・§12): 「{stmt.strip()[:80]}」",
        )


RAW_ACTIVATE_PATTERN = re.compile(r"\b([A-Za-z_]\w*)\.Activate\b")
# ThisWorkbook.Activate(複数ウィンドウの前面切替。シート遷移とは別物)と
# prevActive.Activate(EnsureLayout等の「元のシートへ戻す」復帰処理。戻す先の
# 失敗は非致命的で既存のOERNで許容されている)は素の.Activateチェックの
# 対象外にする(2026-07-31 R11-B)。
RAW_ACTIVATE_EXEMPT_PREFIXES = {"thisworkbook", "prevactive"}
# R11-C: 「Activate失敗を許容してログ(E0801)に残したうえで続行する」防御が
# 確立済みの既存4箇所だけ、行末にこのマーカーを付けて許可リスト化する
# (司令塔裁定・spec_20260731_R11 §5b)。新規追加箇所には付けないこと。
RAW_ACTIVATE_ALLOW_MARKER = "lint:allow-raw-activate"


def check_raw_activate(info: ModuleInfo) -> None:
    """modUI以外での素の.Activate使用をERRORにする(R11-B/C6・C系の再発防止)。

    2026-07-31実機「ギャラリー無反応」の真因は、ActivateSheetRobust(失敗を
    検知し、失敗時はRestoreExcelUIで脱出路を出す)を経由しない素の
    ws.Activateが複数箇所に残っていたこと。失敗しても利用者には何も伝わらず、
    画面が固まったまま戻る手段が無い。modUIはActivateSheetRobust自身の
    実装場所として唯一許可する(対象外)。

    R11-Cで全数を処理済み: modBoot.HideGuardSheetはActivateSheetRobust化、
    modDiag/optDiffDoc/modUIMain/modUIShelfの残り4箇所は「Activate失敗を
    E0801へログして続行する」防御が既に成立している(致命的にしない設計を
    裁定で受容)ため、RAW_ACTIVATE_ALLOW_MARKERを付けて許可リスト化した。
    以後の新規の素の.ActivateはERRORにして機械的に検出する。
    """
    name = module_name_for_display(info)
    if name == "modUI":
        return
    for lineno, raw in merge_continuations(info.raw_text.split("\n")):
        if RAW_ACTIVATE_ALLOW_MARKER in raw:
            continue
        code = _strip_strings(strip_comment(raw))
        for m in RAW_ACTIVATE_PATTERN.finditer(code):
            if m.group(1).lower() in RAW_ACTIVATE_EXEMPT_PREFIXES:
                continue
            info.add(
                "ERROR", lineno,
                f"素の.Activate使用(modUI.ActivateSheetRobust経由にすること。"
                f"失敗時の脱出路が無いと『ギャラリー無反応』と同型の無反応画面になる): "
                f"「{code.strip()[:80]}」",
            )


# ==============================================================================
# メイン
# ==============================================================================
def discover_module_files(src_root: Path) -> list[Path]:
    files = sorted(src_root.rglob("*.bas")) + sorted(src_root.rglob("*.cls"))
    return sorted(set(files))


def run_lint(src_root: Path) -> int:
    if not src_root.exists():
        print(f"[vba_lint] 対象ディレクトリが存在しません: {src_root}")
        return 1

    files = discover_module_files(src_root)
    modules: list[ModuleInfo] = []
    for f in files:
        modules.append(load_module(f, src_root))

    known_modules: dict[str, ModuleInfo] = {}
    for info in modules:
        known_modules[module_name_for_display(info)] = info

    for info in modules:
        check_basics(info)
        check_cp932_safe(info)
        check_dim_type_drop_and_integer(info)
        check_name_shadowing(info)
        check_declaration_position(info)
        check_handler_exit(info)
        check_resume_next_inside_handler(info)
        check_resume_on_fallthrough_label(info)
        check_reserved_identifiers(info)
        check_vba_reserved_words(info)
        check_pure_logic_tokens(info)
        check_opt_token_reference(info)
        check_application_run_whitelist(info)
        check_cross_module_references(info, known_modules)
        check_layer_dependency(info, known_modules)
        check_contract(info)
        check_safeleft_warning(info)
        check_raw_activate(info)

    # モジュールをまたいだモジュールレベル参照は、全モジュールの宣言を
    # 集め終わってからでないと判定できないので、ループの外で1回だけ行う。
    check_module_level_refs(modules)
    check_undefined_proc_refs(modules)
    check_onaction_handler_guard(modules)
    check_msgbox_nonbmp(modules)
    check_array_arg_variant_mismatch(modules)

    # 契約はあるがファイルがまだ存在しないモジュール -> SKIP表示
    implemented_names = set(known_modules.keys())
    not_yet: list[str] = sorted(set(CONTRACT.keys()) - implemented_names)

    total_error = 0
    total_warn = 0
    total_skip = 0

    print("=" * 78)
    print("vba_lint レポート — マイ本棚AI")
    print("=" * 78)

    for info in modules:
        if not info.findings:
            continue
        rel = info.relpath.as_posix()
        errors = [x for x in info.findings if x.level == "ERROR"]
        warns = [x for x in info.findings if x.level == "WARN"]
        skips = [x for x in info.findings if x.level == "SKIP"]
        total_error += len(errors)
        total_warn += len(warns)
        total_skip += len(skips)

        print(f"\n[{rel}]")
        for f in sorted(info.findings, key=lambda x: (x.level != "ERROR", x.level != "WARN", x.line)):
            print(f"  {f.level:<5} L{f.line}: {f.message}")

    if not_yet:
        print("\n[未実装モジュール(契約はあるがファイル無し) — SKIP]")
        for nm in not_yet:
            print(f"  SKIP  {nm}")
        total_skip += len(not_yet)

    print("\n" + "-" * 78)
    print(f"検査対象ファイル数: {len(modules)}")
    print(f"ERROR: {total_error} 件 / WARN: {total_warn} 件 / SKIP: {total_skip} 件")
    if total_error > 0:
        print("結果: NG(exit code 1) — ERRORを解消してください")
    else:
        print("結果: OK(exit code 0)")
    print("-" * 78)

    return 1 if total_error > 0 else 0


def main() -> int:
    parser = argparse.ArgumentParser(description="マイ本棚AI VBA静的Lint")
    parser.add_argument(
        "--path", type=str, default=str(DEFAULT_SRC_ROOT),
        help="検査対象ディレクトリ(既定: mybookshelf/src)",
    )
    args = parser.parse_args()
    src_root = Path(args.path).resolve()
    return run_lint(src_root)


if __name__ == "__main__":
    sys.exit(main())
