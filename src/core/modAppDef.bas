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

Public Const APP_NAME As String = "Nexus Agent"
Public Const APP_VERSION As String = "0.1.0"   ' ビルド時にbuildスクリプトが検証表示
Public Const PACK_FORMAT_VERSION As Long = 1

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
