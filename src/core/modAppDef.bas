Attribute VB_Name = "modAppDef"
Option Explicit

' ============================================================================
' modAppDef - アプリ全体の定数定義(ロジックなし)
' ----------------------------------------------------------------------------
' 役割:
'   バージョン番号・シート名など「あちこちのモジュールから参照される固定値」
'   を1箇所に集約する。値を変えたい時はこのファイルだけを直せばよい。
'
' 設計判断:
'   ・MASTER_SPEC §7.1の契約により、このモジュールは定数宣言のみを持ち、
'     Sub/Functionは一切置かない(ロジック禁止)。
'   ・シート名は全て英小文字(§4)。ユーザーが直接タブを見る4シート
'     (使い方/ホーム/マイ本棚/ダッシュボード)のみ日本語名。
'   ・MASTER_SPECの擬示コードは "Public Const A = 1: Public Const B = 2" の
'     ように「:」で複数宣言を1行に連結しているが、これは紙面節約のための
'     省略表記である。実装では正しいVBA構文どおり1行1宣言に展開する。
'   ・vba_src(ビルド時に自己インストーラ用ソースを格納するシート)は
'     ビルドスクリプトのみが触れる内部シートで、実行時のVBAコードからは
'     一切参照しないため、ここには定数を置かない。
' ============================================================================

' 2026-08-06 R20H FA-12: "Nexus Agent"→"MyBookshelf"。変更前に全参照を
' 走査し(grep結果は完了報告に添付)、MsgBox/InputBoxタイトル・ワークシート
' 見出し・診断文言などの表示専用用途しか無いことを確認済み(パス・キー・
' 識別子・OnTime修飾等の非表示用途は0件)。
Public Const APP_NAME As String = "MyBookshelf"
Public Const APP_VERSION As String = "0.1.0"   ' ビルド時にbuildスクリプトが検証表示
Public Const PACK_FORMAT_VERSION As Long = 1

' 本棚チャンク数上限の「configが読めなかったときの代替値」(2026-08-01 R12-1-8)。
' config シートの shelf_max_chunks が正であり、ここは読めなかった場合の保険。
' 従来はこの保険が呼び出し側4箇所に散らばり、20,000 / 10,000 の2種類に
' 割れていた(ビルドが書く既定は20,500)。同じ問いに3つの答えがある状態は
' 憲章§4-5違反であり、実害としても「取込は10,000で止まるのにダッシュボードの
' 残量ゲージは20,000で計算する」という食い違いになる。
' 値は build_mybookshelf.py が config へ書く既定値と必ず一致させること。
Public Const DEFAULT_SHELF_MAX_CHUNKS As Long = 20500

' ---- シート名(MASTER_SPEC §4) ----------------------------------------------
Public Const SH_HOWTO As String = "使い方"
Public Const SH_HOME As String = "ホーム"
Public Const SH_SHELF As String = "マイ本棚"
Public Const SH_DASH As String = "ダッシュボード"
Public Const SH_CONFIG As String = "config"
Public Const SH_KNOWLEDGE As String = "my_knowledge"
Public Const SH_VECTORS As String = "my_vectors"
Public Const SH_MANIFEST As String = "my_manifest"
Public Const SH_STATS As String = "my_stats"
Public Const SH_USAGE As String = "usage_log"
Public Const SH_ERRLOG As String = "err_log"
Public Const SH_UISTATE As String = "ui_state"
Public Const SH_NEXUS_DASH As String = "Dashboard"   ' Nexus専用ダッシュボード(旧SH_DASHとは別シート)
' R17波0: 構造メタデータ(section_path/refs_out)を格納する新シート(Phase1の
' 受け皿。ビルド時に焼き込み+実行時EnsureChunkMetaSheetで自己修復)。
Public Const SH_CHUNK_META As String = "chunk_meta"
' R17 Phase2: 章単位要約(source/section_key/summary/keywords/chunk_n)の格納先。
' 無い/0行なら俯瞰回答は必ず従来フローへ落ちる(modAskGlobal 側のフェイルセーフ)。
Public Const SH_DOC_OUTLINE As String = "doc_outline"
' R17 Phase3: 用語の表記ゆれ辞書(term, canonical)の格納先。
' 無い/0行なら質問文への同義語追記は無操作(modAskRetrieve 側のフェイルセーフ)。
Public Const SH_SYNONYMS As String = "synonyms"
