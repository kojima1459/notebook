Attribute VB_Name = "modTypes"
Option Explicit

' ============================================================================
' modTypes - アプリ全体で共有するユーザー定義型(Type)のみを置くモジュール
' ----------------------------------------------------------------------------
' 役割:
'   モジュールをまたいで受け渡す構造体をここに集約する。
'   (例: 抽出したページ、本棚の1チャンク、検索ヒット1件)
'
' 設計判断:
'   ・MASTER_SPEC §7.1の契約により、Sub/Functionは一切置かない(型のみ)。
'   ・R4準拠: Excelオブジェクト(Worksheets/Range/Application/ThisWorkbook/
'     MsgBox)には一切触れない、純粋な型定義。LibreOffice実行環境でも
'     そのまま解釈できる。
' ============================================================================

Public Type ExtractedPage
    page As Long
    Text As String
End Type

Public Type ShelfChunk
    chunk_id As String
    source As String
    origin As String
    page As Long
    summary As String
    keywords As String
    full_text As String
End Type

Public Type Hit
    chunk_id As String
    score As Double
    source As String
    page As Long
    preview As String        ' 先頭120字(出典先出し表示用。UI表示にのみ使う)
    origin As String
    full_text As String      ' チャンク本文全体(最大32000字。回答生成の根拠として
                              ' modPromptsが使う。PM裁定1: previewだけでは根拠不足)
End Type
