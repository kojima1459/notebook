Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================================
' modConfig - configシート(A=key, B=value, C=説明)の読み書き
' ----------------------------------------------------------------------------
' 役割:
'   設定値は全てconfigシートに集約し、非エンジニアがVBAを開かずにExcelの
'   セルを書き換えるだけでチューニングできるようにする。キー文字列でA列を
'   検索し、対応するB列の値を返す(MASTER_SPEC §5のキー台帳を参照)。
'
' 設計判断:
'   ・V2 (src/chatbot_v2/modConfig.bas) をほぼそのまま流用(MASTER_SPEC §7.1
'     の指示どおり)。シート名だけ固定リテラルではなく modAppDef.SH_CONFIG
'     を参照するように変更した(依存は modAppDef のみで、循環参照なし)。
'   ・値のキャッシュは持たない。configは頻繁に読み書きする処理ではないため、
'     常にシートを直接読みに行くシンプルな実装を優先した。
'   ・EnsureLoadedはconfigシートの存在確認のみ行う。見つからない場合は
'     ここでは modLog を呼ばない(modLog.LogError は debug_mode 判定のため
'     modConfig を呼び返す設計になっており、ここで modLog を呼ぶと循環参照
'     になってしまう。MASTER_SPEC §12「モジュール間の循環参照禁止」に抵触
'     するため、致命的な欠落の検知と案内は modDiag 側に委ねる)。
'     ただしユーザーには即座に気づいてもらう必要があるため、V2と同様に
'     その場でMsgBoxを出す(コード E0101 を文面に含める)。
'   ・シートやキーが見つからない場合は例外を投げず既定値にフォールバックする。
' ============================================================================

Private mLoaded As Boolean

Public Sub EnsureLoaded()
    On Error GoTo NoSheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(modAppDef.SH_CONFIG)
    mLoaded = True
    Exit Sub
NoSheet:
    MsgBox "設定ファイル(configシート)に必要な項目が見つかりません。" & vbLf & _
           "配布元にこのファイルの再入手を依頼してください。" & vbLf & _
           "(コード: E0101)", vbCritical, modAppDef.APP_NAME
End Sub

Public Function GetString(ByVal key As String, ByVal defaultValue As String) As String
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Then
        GetString = defaultValue
    Else
        GetString = CStr(v)
    End If
End Function

Public Function GetLong(ByVal key As String, ByVal defaultValue As Long) As Long
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Or Not IsNumeric(v) Then
        GetLong = defaultValue
    Else
        GetLong = CLng(v)
    End If
End Function

Public Function GetDouble(ByVal key As String, ByVal defaultValue As Double) As Double
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Or Not IsNumeric(v) Then
        GetDouble = defaultValue
    Else
        GetDouble = CDbl(v)
    End If
End Function

Public Function GetBool(ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Then
        GetBool = defaultValue
    ElseIf VarType(v) = vbBoolean Then
        GetBool = CBool(v)
    ElseIf IsNumeric(v) Then
        GetBool = (CLng(v) <> 0)
    Else
        Dim s As String: s = LCase$(Trim$(CStr(v)))
        GetBool = (s = "true" Or s = "1" Or s = "yes" Or s = "on")
    End If
End Function

Public Sub SetValue(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim r As Long: r = FindKeyRow(ws, key)
    If r = 0 Then
        ' 新規キーは末尾に追記
        r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
        ws.Cells(r, 1).Value = key
    End If
    ws.Cells(r, 2).Value = value
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------
Private Function LookupValue(ByVal key As String) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim r As Long: r = FindKeyRow(ws, key)
    If r = 0 Then Exit Function
    LookupValue = ws.Cells(r, 2).Value
End Function

Private Function FindKeyRow(ByVal ws As Worksheet, ByVal key As String) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim i As Long
    For i = 2 To lastRow      ' 1行目はヘッダ
        If StrComp(CStr(ws.Cells(i, 1).Value), key, vbTextCompare) = 0 Then
            FindKeyRow = i
            Exit Function
        End If
    Next i
End Function
