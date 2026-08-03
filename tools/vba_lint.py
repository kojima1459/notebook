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
    "modLog": {
        "closed": True,
        "required": ["LogError", "LogUsage", "FriendlyMessage", "ShowError",
                     "FriendlyFailMsg", "SharedReadFailMsg"],
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
        "required": ["RunDiagnostics", "QuickHealthCheck", "RecentErrorsForClipboard"],
    },
    # ---- 7.2 取込層 ----
    # modExtractorWord / modExtractorExcel / modExtractorAcrobat は
    # MASTER_SPECが個別のPublic契約を明示していないため対象外(自由)。
    "modExtractor": {
        "closed": True,
        # SharedCopyNextChunkLen: 2026-07-30 R2要件C(共有読みローカルコピー)の
        # 純ロジック部分。実ファイルI/Oを含むCopySharedRead自体はテストできない
        # ため、「次に読むべきバイト数」の境界計算だけを切り出してPublic化し、
        # modTestsPure3から検証できるようにした。
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
        "required": ["ExtractFile", "SupportedExts", "SharedCopyNextChunkLen",
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
        "required": ["ExtractPdfWithFallback", "CopyToLocalTemp",
                     "DropGarbledPages", "TempBaseNameFor",
                     "IsThinExtract", "ThinExtractMemoFor",
                     "IsUnreadableCopyReason", "CopyFailMsgFor"],
    },
    "modMode": {
        "closed": True,
        # 回答モード(すぐ聞く/通常/入念)の方針。純ロジックなのでLOで検証する。
        # AskStage*(2026-08-03 R13-9b): 質問処理の段階ナレーションの算数。
        #   「(2/4) 資料を照合中…」の番号は、その回に実際に通す段の数から作る
        #   (拡張・再ランクは構成で有無が変わるため、総数を4に固定すると嘘になる)。
        #   表示そのものは modAskRetrieve.ShowAskStage が行い、ここは純ロジック。
        "required": ["Normalize", "NextMode", "Caption", "Description", "TopK",
                     "UseExpand", "UseRerank", "UseVerify", "UseLightExpand",
                     "SubQueryCount",
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
        "required": ["IngestFile", "DeleteSource",
                     "SourceList", "TotalChunks", "IsBusy"],
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
        "required": ["AddFilesViaDialog", "AddFilesResult", "IsBatchBusy",
                     "StageBanner"],
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
        "required": ["BuildQuickPrompt", "BuildDeepDraftPrompt", "BuildDeepVerifyPrompt", "BuildEnrichPrompt", "BuildExpandPrompt", "BuildRerankPrompt"],
    },
    "modRagParse": {
        "closed": True,
        # 多段RAGのLLM応答パーサ(設計書§C/§D/§G-T7)。全て寛容退化の純関数。
        "required": ["ParseExpand", "ParseRankOrder", "ExtractAnswer", "ParseSubqueries"],
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
                     "HistoryBlock", "IsErrorResponse"],
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
                     "BeginDeferredBadges", "FlushBadgeToasts"],
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
    "optVision": {"closed": True, "required": ["Ping", "ExtractImagePdf", "ExtractImagePdfText",
                                               "ExtractPdfOcrPagedText", "OcrCapMemo",
                                               "HasClipboardImage", "SaveClipboardImage",
                                               "ResetGsGuidance", "ExtractPdfTextNoOcr",
                                               "FindGsExeByCandidates", "PathExists",
                                               "IsVisionError", "SafeResultToString"]},
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
    # GsFailureDetail / GsExitCode(2026-07-31 R11-D・監査3 H-2): 完了フラグの
    # 中身(GSの終了コード)と gs_out.log の先頭を読み、err_log の detail に
    # そのまま入れられる1本の文字列にする。optVision(OCR経路)からも呼ぶ。
    "optGsTxt": {"closed": True, "required": ["Ping", "ExtractPdfTextNoOcr",
                                              "MakeOcrFolder",
                                              "WaitForDoneFlag", "CleanupOcrFolder",
                                              "GsFailureDetail", "GsExitCode"]},
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
            "BatchCountFor", "BatchBoundsFor", "OcrPageBanner", "OcrCapMemoFor",
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
            "PaintProgress", "ClearProgress", "CurrentTheme", "SaveTheme",
            "ThemeColor", "ApplyTheme", "PaintBubble", "PaintActionButton",
            "SetShapeTextColor", "ThemeIcon",
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
            "DrawStatTiles", "DrawInbox",
        ],
    },
    # R11-F1: ツールバーを modKnowledgeBar へ分離した残り(ヘッダー/右肩ピル/モード管理/ハンドラ)。
    "modKnowledge": {
        "closed": True,
        "required": [
            "CHROME_ROWS", "DrawChrome", "PrepareScreenView", "IsTableMode",
            "ContentTop", "SearchCellAddress", "OnGoGallery", "OnGoShared",
            "OnGoTable", "OnBackHub", "OnHelp", "OnToChat", "OnSearch", "OnGapBoard",
            "OnChannels", "OnRegister", "OnAddFiles", "OnPackOut", "OnPackIn",
            "OnSync", "OnPickFolder", "OnDelete", "RefreshCurrent",
        ],
    },
    # R11-F1: modKnowledge から分離したツールバー(BAR_H は DrawChrome が行3の高さに使うため Public)。
    "modKnowledgeBar": {
        "closed": True,
        "required": [
            "BAR_H", "DrawToolbar",
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
            "DrawKpiRow", "DrawExpBar", "ChartNoteY", "DrawBadgeShelf",
            "DrawChartPlaceholder",
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
    # modChrome(2026-07-30 R4): 配置計算だけを持つのでExcelオブジェクトは
    # 一切要らない。ここへ載せることで「うっかりRangeを触る」改修を機械で止める。
    "modChrome",
    # optOcrCore(2026-07-31 R6): GSコマンド文字列の組み立てとページ上限の算数
    # だけを持つ。opt層のモジュールだが副作用ゼロで、LO実行テストから直接
    # 呼べる状態を維持するために純ロジック検査の対象へ入れる。
    "optOcrCore",
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
    # 2026-07-31(R11-F2): qa層の3モジュールを追加。いずれも実測でExcel
    # オブジェクトトークン0件(Worksheets/Range(/Application./ThisWorkbook/
    # MsgBox/ActiveSheet が1つも無い)。純ロジックであることを規約として
    # 固定し、「ちょっとRangeを見たい」という改修を機械で止める。
    #   modBitwiseOpt: ビット演算による候補絞り込み。
    #   modFollowup  : 深掘り候補の抽出と本文からの除去。
    #   modClarify   : 聞き返し文の生成と番号選択の判定。
    "modBitwiseOpt", "modFollowup", "modClarify",
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
                    if (cur_layer == LAYER_MID
                            and (prefix, member) == ("modUiLock", "IsBusy")
                            and self_name in R1_UILOCK_ALLOWED_MODULES):
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
