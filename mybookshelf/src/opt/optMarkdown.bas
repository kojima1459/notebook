Attribute VB_Name = "optMarkdown"
Option Explicit

' ============================================================================
' optMarkdown - 回答等をMarkdown整形してセルに表示する(opt機能・MASTER_SPEC §7.7)
' ----------------------------------------------------------------------------
' 役割:
'   AIリボンのMarkdown整形表示関数(CellMarkDown仮定)へ、Markdown文字列と
'   表示先セルを渡す。この層が無くても製品は成立し、modules.jsonから1行
'   削除するだけで撤去できる(§7.7)。
'
' 設計判断:
'   ・リボン呼び出しは modGateway.TryRibbonRun のみを使う(R3)。
'   ・表示先の指定は「シート名+セル番地(文字列)」という疎結合な契約
'     (§7.7のシグネチャどおり)。Range自体は本モジュール内でシート名/
'     セル番地から解決する(呼び出し元がRangeオブジェクトを直接渡さずに
'     済むようにするための設計。UI層は文字列だけ渡せばよい)。
'   ・失敗しても例外を外に出さない。戻り値は "" (成功) または "#ERR:..." の
'     文字列(optTts.SpeakAnswerと同じ「空文字=成功」規約に揃えた。
'     §7.7はこの関数の戻り値の成功/失敗フォーマットまでは明記していないため、
'     opt層内で一貫させることを優先した実装判断)。
'   ・modUIMain.SetStage 以外のコア参照は行わない(基盤層のみ参照可・§7.7)。
'     シート/セルの解決自体はモジュール自身がExcelオブジェクトへ直接触れて
'     行う(他モジュール経由ではないため§7.7の「コアモジュールへの参照」
'     制限には抵触しない)。
' ============================================================================

' === SIGNATURE ASSUMPTION ===
'   仮定シグネチャ: CellMarkDown(md As String, targetRange As Range) -> Variant
'     成功時: 空文字列 または 何らかの確認用文字列を返す。
'     失敗時: "error" を含む文字列、または空でない失敗メッセージを返す。
'   根拠: MASTER_SPEC §7.7 optMarkdown.bas の記載「CellMarkDown仮定:
'     (md, targetRange)」のみが根拠。引数の順序(Markdown文字列が先か、
'     Rangeが先か)、Rangeを直接渡せるのか文字列アドレスで渡すべきかは未確認。
'     ここではRangeオブジェクトを直接渡す仮定を採用した
'     (modGateway.TryRibbonRunはVariant配列の要素をそのままApplication.Run
'     に渡すため、Rangeオブジェクトを含めても技術的には問題ない)。
'   仕様回答が来たら直す行:
'     ・引数の数/順序が違う場合(例: セル番地を文字列で渡す等)
'       -> BuildMarkdownArgs() の Array(...) 部分のみ修正。
'     ・戻り値の成功/失敗判定基準が違う場合 -> IsMarkdownError() の判定式のみ修正。
'     ・関数名自体が違う場合(CellMarkDown以外) -> Private Const RIBBON_FUNC_NAME
'       の値のみ修正。
' ================================================================================

Private Const RIBBON_FUNC_NAME As String = "CellMarkDown"

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' RenderMarkdownAt - sheetName!cellAddr にMarkdown文字列mdを整形表示する
'   (MASTER_SPEC §7.7)。戻り値: ""=成功 / "#ERR:..."=失敗(例外は出さない)。
' ----------------------------------------------------------------------------
Public Function RenderMarkdownAt(ByVal sheetName As String, ByVal cellAddr As String, ByVal md As String) As String
    On Error GoTo Fail

    If LenB(Trim$(md)) = 0 Then
        RenderMarkdownAt = "#ERR:表示するMarkdown文字列が空です"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = FindSheet(sheetName)
    If ws Is Nothing Then
        RenderMarkdownAt = "#ERR:シート「" & sheetName & "」が見つかりません"
        Exit Function
    End If

    Dim targetRange As Range
    Set targetRange = ResolveRange(ws, cellAddr)
    If targetRange Is Nothing Then
        RenderMarkdownAt = "#ERR:セル番地「" & cellAddr & "」を解決できません"
        Exit Function
    End If

    modUIMain.SetStage "📝 表示を整えています…"

    Dim result As Variant
    result = modGateway.TryRibbonRun(RIBBON_FUNC_NAME, BuildMarkdownArgs(md, targetRange))

    modUIMain.SetStage ""

    Dim s As String
    s = SafeResultToString(result)

    If Left$(s, 5) = "#ERR:" Then
        RenderMarkdownAt = s
        Exit Function
    End If
    If IsMarkdownError(s) Then
        RenderMarkdownAt = "#ERR:Markdown表示に失敗しました: " & modUtil.SafeLeft(s, 200)
        Exit Function
    End If

    RenderMarkdownAt = ""
    Exit Function

Fail:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    modUIMain.SetStage ""
    modLog.LogError "E0202", "optMarkdown.RenderMarkdownAt", errDesc
    RenderMarkdownAt = "#ERR:Markdown表示処理でエラーが発生しました: " & errDesc
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: optMarkdownの公開契約はPing/RenderMarkdownAtのみ)
' ----------------------------------------------------------------------------

Private Function FindSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set FindSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function ResolveRange(ByVal ws As Worksheet, ByVal cellAddr As String) As Range
    On Error Resume Next
    Set ResolveRange = ws.Range(cellAddr)
    On Error GoTo 0
End Function

' === SIGNATURE ASSUMPTION: 引数配列の組み立て。仕様が判明したらここだけ直す ===
Private Function BuildMarkdownArgs(ByVal md As String, ByVal targetRange As Range) As Variant
    BuildMarkdownArgs = Array(md, targetRange)
End Function

' === SIGNATURE ASSUMPTION: 失敗判定。仕様が判明したらここだけ直す ===
Private Function IsMarkdownError(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Then
        IsMarkdownError = False   ' 空文字は成功扱い(上部コメントの規約)
    ElseIf StrComp(Trim$(s), "error", vbTextCompare) = 0 Then
        IsMarkdownError = True
    Else
        IsMarkdownError = modGateway.LooksLikeLimitError(s)
    End If
End Function

Private Function SafeResultToString(ByVal result As Variant) As String
    On Error GoTo Empty0
    If IsError(result) Then Exit Function
    SafeResultToString = CStr(result)
    Exit Function
Empty0:
    SafeResultToString = ""
End Function
