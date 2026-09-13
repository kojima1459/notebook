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


Public Sub EnsureLoaded()
    On Error GoTo NoSheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(modAppDef.SH_CONFIG)
    Exit Sub
NoSheet:
    MsgBox "設定ファイル(configシート)に必要な項目が見つかりません。" & vbLf & _
           "配布元にこのファイルの再入手を依頼してください。" & vbLf & _
           "(コード: E0101)", vbCritical, modAppDef.APP_NAME
End Sub

' 2026-07-28(レビュー M-7): config は非エンジニアが直接編集する設計なので、
' セルにエラー値(#REF! 等)や桁あふれが混ざることは現実的に起きる。
' 従来は LookupValue が返した vbError を CStr/CLng に渡して型不一致や
' オーバーフローを起こし、【そのキーだけでなく、以降の全キーの取得が例外化】
' していた。しかも modBoot は中盤でハンドラを外していた(H-3)ため、
' そこで素のエラーダイアログが出て EnableEvents=False が焼き付く経路があった。
' 「既定値へ落ちる」という契約を、どんな入力でも守り切る。
Public Function GetString(ByVal key As String, ByVal defaultValue As String) As String
    On Error GoTo Fallback
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Or IsError(v) Then
        GetString = defaultValue
    Else
        GetString = CStr(v)
    End If
    Exit Function
Fallback:
    GetString = defaultValue
End Function

Public Function GetLong(ByVal key As String, ByVal defaultValue As Long) As Long
    On Error GoTo Fallback
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Or IsError(v) Then
        GetLong = defaultValue
    ElseIf Not IsNumeric(v) Then
        GetLong = defaultValue
    Else
        ' "99999999999" のような桁あふれは CLng がオーバーフローする。
        ' 既定値へ落とす(Long の範囲外は設定として意味を成さない)。
        Dim d As Double: d = CDbl(v)
        If d > 2147483647# Or d < -2147483648# Then
            GetLong = defaultValue
        Else
            GetLong = CLng(d)
        End If
    End If
    Exit Function
Fallback:
    GetLong = defaultValue
End Function

Public Function GetDouble(ByVal key As String, ByVal defaultValue As Double) As Double
    On Error GoTo Fallback
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Or IsError(v) Then
        GetDouble = defaultValue
    ElseIf Not IsNumeric(v) Then
        GetDouble = defaultValue
    Else
        GetDouble = CDbl(v)
    End If
    Exit Function
Fallback:
    GetDouble = defaultValue
End Function

Public Function GetBool(ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    On Error GoTo Fallback
    Dim v As Variant
    v = LookupValue(key)
    If IsEmpty(v) Or IsError(v) Then
        GetBool = defaultValue
    ElseIf VarType(v) = vbBoolean Then
        GetBool = CBool(v)
    ElseIf IsNumeric(v) Then
        GetBool = (CLng(v) <> 0)
    Else
        GetBool = ParseBoolText(CStr(v), defaultValue)
    End If
    Exit Function
Fallback:
    GetBool = defaultValue
End Function

' ----------------------------------------------------------------------------
' ParseBoolText - configのB列に入っていた【文字列】を真偽へ解釈する純関数。
' ----------------------------------------------------------------------------
'   【直したこと】(2026-08-16 R33 W4-1)
'   兄弟の GetLong / GetDouble は「解釈できない値は defaultValue へ落とす」
'   分岐(`ElseIf Not IsNumeric(v)`)を持ち、本ファイル45行の契約も
'   「『既定値へ落ちる』という契約を、どんな入力でも守り切る」と宣言して
'   いるのに、GetBool だけがその分岐を持たず
'   `GetBool = (s = "true" Or s = "1" Or …)` と書かれていた。つまり
'   【非空で解釈できない文字列】は defaultValue を無視して必ず False に
'   なっていた。configは非エンジニアが直接編集する設計なので、
'     ・値を消すつもりでスペースキーを押した(v=" ")
'     ・日本語IMEを全角のまま「ＴＲＵＥ」と打って確定した
'   のどちらも「TRUEに見える/空に見える」まま無条件 False になり、
'   既定TRUEのキー(insight_share_enabled / sync_on_open / conv_bridge 等)が
'   無言で全部オフに倒れる。エラーもログも出ず、modDiag はキーの【存在】
'   しか見ないので診断でも気付けない。同じ「空に見えるセル」でも Delete
'   なら既定値・スペース1個なら False と結果が反転するのは設計として
'   弁護できないため、取り残しと判断して3分岐へ揃えた。
'
'   【3分岐の意味】
'     真トークン("true"/"1"/"yes"/"on")   -> True
'     偽トークン("false"/"0"/"no"/"off")  -> False   ← 明示的なFALSEは尊重する
'     それ以外(空文字・空白のみ・誤記)    -> defaultValue(読めなかった)
'   偽トークンを独立させているのは、「FALSEと書いた」意図を既定TRUEの
'   キーで握り潰さないため。ここを2分岐(真トークン以外は既定値)にすると、
'   運用保守ガイドが指示する insight_share_enabled=FALSE の手入力が効かなく
'   なる。両方向のゴールデンを modTestsPure34 が固定している。
'
'   【Public である理由】VBA の Private プロシージャは他モジュールから
'   呼べず、modTestsPure34 からゴールデン固定できないため(§7契約表に登録)。
'   GetString / GetLong / GetDouble はこの関数を通らないので挙動は不変。
' ----------------------------------------------------------------------------
Public Function ParseBoolText(ByVal raw As String, ByVal defaultValue As Boolean) As Boolean
    Dim s As String
    s = LCase$(Trim$(NarrowAscii(raw)))
    Select Case s
        Case "true", "1", "yes", "on"
            ParseBoolText = True
        Case "false", "0", "no", "off"
            ParseBoolText = False
        Case Else
            ParseBoolText = defaultValue
    End Select
End Function

' ----------------------------------------------------------------------------
' NarrowAscii - 全角の英数字と全角スペースだけを半角へ均す(純関数)。
' ----------------------------------------------------------------------------
'   StrConv(s, vbNarrow) は日本語ロケールでしか意図どおり動かない(他ロケール
'   では変換されない/半角カナまで巻き込む)ので使わない。判定に要る文字は
'   トークン8種を構成する英数字だけなので、Replace の小さな表で足りる。
'   全角スペース(U+3000)を半角へ均すのは、Trim$ が U+3000 を落とさず
'   「全角スペースだけのセル」が解釈不能扱いから漏れるのを防ぐため。
'   非ASCIIは必ず ChrW で組む(CP932 へ潰れる文字をソースに直接置かない)。
'   modPii.NormalizeWidth と役割は似ているが、modPii は中間層(src/pack)で
'   基盤層(src/core)からは参照できない(lintの層依存検査でERROR)ため、
'   共有せずここに閉じている。あちらは PII 走査用にハイフン類・＠・
'   ピリオドまで均す別物で、統合すると片方の都合が他方を壊す。
' ----------------------------------------------------------------------------
Private Function NarrowAscii(ByVal s As String) As String
    Dim t As String
    t = Replace(s, ChrW(&H3000&), " ")

    Dim i As Long
    For i = 0 To 9
        t = Replace(t, ChrW(&HFF10& + i), Chr$(48 + i))
    Next i
    For i = 0 To 25
        t = Replace(t, ChrW(&HFF21& + i), Chr$(65 + i))
        t = Replace(t, ChrW(&HFF41& + i), Chr$(97 + i))
    Next i

    NarrowAscii = t
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
