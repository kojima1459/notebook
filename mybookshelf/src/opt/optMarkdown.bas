Attribute VB_Name = "optMarkdown"
Option Explicit

' ============================================================================
' optMarkdown - 回答等をMarkdown整形してセルに表示する(opt機能・MASTER_SPEC §7.7)
' ----------------------------------------------------------------------------
' 役割:
'   AIリボンのMarkdown整形表示関数(CellMarkDown・確定済み)へ表示先セルを
'   渡してセル内Markdownを装飾する。加えて、Markdown文字列をWordで開く
'   (OpenWordMark・確定済み)ラッパーを提供する。この層が無くても製品は
'   成立し、modules.jsonから1行削除するだけで撤去できる(§7.7)。
'
' 設計判断:
'   ・リボン呼び出しは modGateway.TryRibbonRun のみを使う(R3)。
'   ・表示先の指定は「シート名+セル番地(文字列)」という疎結合な契約
'     (§7.7のシグネチャどおり)。Range自体は本モジュール内でシート名/
'     セル番地から解決する(呼び出し元がRangeオブジェクトを直接渡さずに
'     済むようにするための設計。UI層は文字列だけ渡せばよい)。
'   ・失敗しても例外を外に出さない。戻り値は "" (成功) または "#ERR:..." の
'     文字列(opt層共通の「空文字=成功」規約。§7.7)。
'   ・mock_llm=TRUE のときはリボンを呼ばない:
'       RenderMarkdownAt … セルへの書き込みだけ行い装飾はスキップして成功扱い
'         (mock時も回答テキスト自体は表示される、という穏当なフォールバック)。
'       OpenAnswerInWord … Word連携は代替表示が無いため、親切メッセージの
'         "#ERR:..." を返す(modGatewayのmock時「ダミー応答」方針と同様、
'         実リボンには一切触れない)。
'   ・modUIMain.SetStage 以外のコア参照は基盤層(modConfig/modLog/modUtil/
'     modGateway)のみ(§7.7)。シート/セルの解決自体はモジュール自身が
'     Excelオブジェクトへ直接触れて行う(他モジュール経由ではないため
'     §7.7の「コアモジュールへの参照」制限には抵触しない)。
' ============================================================================

' === 確定済みシグネチャ(出典: RIBBON_API_CONFIRMED.md §1 #10, #12 / 裁定D6) ===
'   CellMarkDown(rng As Range, [isComment As Boolean = False]) -> (戻り値未使用)
'     「セルに既に入っているMarkdown」を書式付き表示に変換する。Markdown文字列を
'     引数で渡す方式ではないため、本モジュールは先に対象セルへmdを書き込んで
'     から CellMarkDown(rng, False) を呼ぶ(#12・裁定D6)。
'   OpenWordMark(Text As String) -> (戻り値未使用)
'     Markdown文字列をWordで開く(#10)。
'   戻り値の扱い: どちらも台帳上「(未使用)」のため、リボン側戻り値での
'     成否判定は行わない。失敗判定は TryRibbonRun 側の "#ERR:" 始まり
'     (Application.Run自体の失敗)のみとする。
' ============================================================================

Private Const RIBBON_FUNC_NAME As String = "CellMarkDown"
Private Const WORD_FUNC_NAME As String = "OpenWordMark"

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' RenderMarkdownAt - sheetName!cellAddr にMarkdown文字列mdを整形表示する
'   (MASTER_SPEC §7.7)。確定規約(裁定D6): 対象セルへmdを書き込んでから
'   CellMarkDown(rng, False) でセル内Markdownを装飾する。
'   戻り値: ""=成功 / "#ERR:..."=失敗(例外は出さない)。
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

    ' (1) 対象セルへMarkdown文字列を書き込む(セル上限32,767字対策で
    '     SafeLeft(…, 32000) を経由。規約§12)。
    targetRange.Value = modUtil.SafeLeft(md, 32000)

    ' mock_llm=TRUE のときはリボンを呼ばず、テキスト書き込みのみで成功扱い
    ' (装飾はスキップ。実リボン不在環境でもE2Eが回るようにするため)。
    If modConfig.GetBool("mock_llm", True) Then
        modUIMain.SetStage ""
        RenderMarkdownAt = ""
        Exit Function
    End If

    ' (2) セル内Markdownを装飾(確定: CellMarkDown(rng, False) / 台帳§1 #12・裁定D6)
    Dim result As Variant
    result = modGateway.TryRibbonRun(RIBBON_FUNC_NAME, Array(targetRange, False))

    modUIMain.SetStage ""

    Dim s As String
    s = SafeResultToString(result)

    ' 戻り値は台帳上「未使用」のため、TryRibbonRun側の失敗("#ERR:"始まり)のみ
    ' 失敗と判定する(確定仕様。上部コメントブロック参照)。
    If Left$(s, 5) = "#ERR:" Then
        RenderMarkdownAt = s
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
' OpenAnswerInWord - Markdown文字列mdをWordで開く(確定: OpenWordMark /
'   台帳§1 #10・裁定D6)。UI配線は別途(本モジュールはラッパーのみ提供)。
'   戻り値: ""=成功 / "#ERR:..."=失敗(例外は出さない)。
' ----------------------------------------------------------------------------
Public Function OpenAnswerInWord(ByVal md As String) As String
    On Error GoTo Fail

    If LenB(Trim$(md)) = 0 Then
        OpenAnswerInWord = "#ERR:Wordで開くMarkdown文字列が空です"
        Exit Function
    End If

    ' mock_llm=TRUE のときはリボンを呼ばず、親切メッセージで案内する
    ' (Word連携には代替表示が無いため。他のmock時挙動と同じくリボン非接触)。
    If modConfig.GetBool("mock_llm", True) Then
        OpenAnswerInWord = "#ERR:mockモード(mock_llm=TRUE)ではWord連携を利用できません。" & _
            "config の mock_llm を FALSE にすると実際にWordで開けるようになります。"
        Exit Function
    End If

    modUIMain.SetStage "📝 Wordで開いています…"

    Dim result As Variant
    result = modGateway.TryRibbonRun(WORD_FUNC_NAME, Array(md))

    modUIMain.SetStage ""

    Dim s As String
    s = SafeResultToString(result)

    ' 戻り値は台帳上「未使用」のため、TryRibbonRun側の失敗("#ERR:"始まり)のみ
    ' 失敗と判定する(確定仕様。上部コメントブロック参照)。
    If Left$(s, 5) = "#ERR:" Then
        OpenAnswerInWord = s
        Exit Function
    End If

    OpenAnswerInWord = ""
    Exit Function

Fail:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    modUIMain.SetStage ""
    modLog.LogError "E0202", "optMarkdown.OpenAnswerInWord", errDesc
    OpenAnswerInWord = "#ERR:Wordで開く処理でエラーが発生しました: " & errDesc
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: optMarkdownの公開契約はPing/RenderMarkdownAt/
' OpenAnswerInWordのみ)
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

Private Function SafeResultToString(ByVal result As Variant) As String
    On Error GoTo Empty0
    If IsError(result) Then Exit Function
    SafeResultToString = CStr(result)
    Exit Function
Empty0:
    SafeResultToString = ""
End Function
