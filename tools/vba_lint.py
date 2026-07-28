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
    "modLog": {
        "closed": True,
        "required": ["LogError", "LogUsage", "FriendlyMessage", "ShowError"],
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
        "required": ["CallLLM", "GetEmbedding", "RibbonAvailable", "TryRibbonRun", "LooksLikeLimitError", "RunLimitCheck", "GetEmbeddingsBatch"],
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
        "required": ["ExtractFile", "SupportedExts"],
    },
    "modMode": {
        "closed": True,
        # 回答モード(すぐ聞く/通常/入念)の方針。純ロジックなのでLOで検証する。
        "required": ["Normalize", "NextMode", "Caption", "Description", "TopK",
                     "UseExpand", "UseRerank", "UseVerify", "UseLightExpand",
                     "SubQueryCount"],
    },
    "modSparse": {
        "closed": True,
        # 日本語キーワード検索(文字bigram + BM25 + 完全一致)。
        # 全て純ロジックなので modTestsPure2 から直接検証する。
        # 実測: 旧実装 R@1 32% → 本実装 84%(tools/bench_retrieval.py)。
        "required": ["NormalizeForSearch", "Tokenize", "DistinctiveKeys",
                     "Bm25Score", "ExactHitCount", "CompactForMatch", "KeyScore"],
    },
    "modChunker": {
        "closed": True,
        # ChunkPagesEx/ClassifyLine/BuildBreadcrumb: 構造認識チャンク化(設計書§B)。
        # ChunkPagesは後方互換(legacy)のまま不変。
        # NormalizeForIngest/JoinSplitNumbers/IsPageNumberLine: 取込正規化
        #   (2026-07-27追加)。実物の約款PDFで条見出しの20.2%を取りこぼしていた
        #   ため、全取込経路が最初に通す正規化として追加。純ロジックなので
        #   Publicにして modTestsPure から直接検証する(回帰を二度と許さない)。
        "required": ["ChunkPages", "ChunkPagesEx", "ClassifyLine", "BuildBreadcrumb",
                     "NormalizeForIngest", "JoinSplitNumbers", "IsPageNumberLine"],
    },
    "modEmbed": {
        "closed": True,
        # MarkAllForReembed: 圧縮/次元変更後の全再embed導線(設計書§G-T10)。
        "required": ["EmbedPending", "PendingCount", "MarkAllForReembed"],
    },
    "modShelf": {
        "closed": True,
        "required": ["AddFilesViaDialog", "IngestFile", "DeleteSource", "SourceList", "TotalChunks"],
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
        "required": [
            "PickShelfFolder", "SyncNow", "ScheduleAutoSync", "CancelAutoSync",
            "DiffDecision", "AutoSyncTick", "ResolveDecision",
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
    "modAskRetrieve": {
        "closed": True,
        "required": ["RunMultiRetrieve", "ApplyLowHitWarning", "IsTooVague", "HitSourceList"],
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
                     "BadgeCatalog", "BadgeEarnedOn"],
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
        "required": [
            "EnsureLayout", "RenderShelf", "OnAddFiles", "OnSyncNow",
            "OnPickFolder", "OnExportPack", "OnImportPack", "OnDeleteSource",
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
    "optVision": {"closed": True, "required": ["Ping", "ExtractImagePdf", "ExtractImagePdfText",
                                               "HasClipboardImage", "SaveClipboardImage"]},
    # OpenAnswerInWord: 確定関数OpenWordMarkのラッパー(裁定D6)。
    # ExportAnswerAsDoc: 対話型Word文書生成(裁定D12・指示文→LLM整形→OpenWordMark)
    "optMarkdown": {"closed": True, "required": ["Ping", "RenderMarkdownAt", "OpenAnswerInWord",
                                                 "ExportAnswerAsDoc"]},
    "optDiffDoc": {"closed": True, "required": ["Ping", "CompareTwoDocsDialog"]},
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
    # modTestsExcel はMASTER_SPECがPublic契約を明示していないため対象外。
}

# 純ロジックモジュール(R4): Excelオブジェクトトークン禁止
# modPrompts はMASTER_SPEC R4本文には明記されていないが、§7.3で
# 「純文字列」モジュールと明記されており、テストハーネス発注元の指示
# (1-Iタスク定義)により本Lintでは禁止トークン検査の対象に加える。
PURE_LOGIC_MODULES = {
    "modUtil", "modChunker", "modPii", "modTypes",
    "modTestRunner", "modTestsPure", "modTestsPure2", "modPrompts",
    "modRagParse", "modSparse", "modMode",
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

    for lineno, stmt in info.statements:
        for m in DOTTED_REF_PATTERN.finditer(stmt):
            prefix, member = m.group(1), m.group(2)
            if not MODULE_SHAPED_NAME.match(prefix):
                continue
            target = known_modules.get(prefix)
            if target is None or target.layer is None:
                continue
            if prefix == module_name_for_display(info):
                continue

            if cur_layer == LAYER_OPT:
                if target.layer == LAYER_FOUNDATION:
                    continue
                if prefix == "modUIMain" and member == "SetStage":
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
                    if cur_layer == LAYER_MID and (prefix, member) in (
                        ("modUIMain", "SetStage"),
                        ("modUIMain", "RenderSourcesPreview"),
                        ("modUIMain", "RenderAnswer"),
                        ("modUIShelf", "RenderShelf"),
                    ):
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
        check_reserved_identifiers(info)
        check_vba_reserved_words(info)
        check_pure_logic_tokens(info)
        check_opt_token_reference(info)
        check_application_run_whitelist(info)
        check_cross_module_references(info, known_modules)
        check_layer_dependency(info, known_modules)
        check_contract(info)
        check_safeleft_warning(info)

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
