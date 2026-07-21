Attribute VB_Name = "optMarkdown"
Option Explicit

' ============================================================================
' optMarkdown - 回答等をMarkdown整形してセルに表示する(opt機能・MASTER_SPEC §7.7)
' ----------------------------------------------------------------------------
' 役割:
'   AIリボンのMarkdown整形表示関数(CellMarkDown・確定済み)へ表示先セルを
'   渡してセル内Markdownを装飾する。加えて、Markdown文字列をWordで開く
'   (OpenWordMark・確定済み)ラッパーと、裁定D12の対話型文書生成
'   (ExportAnswerAsDoc: 指示文→LLM整形→Word起動)を提供する。この層が
'   無くても製品は成立し、modules.jsonから1行削除するだけで撤去できる(§7.7)。
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
'       ExportAnswerAsDoc … 同上(LLM整形のmock応答だけ返しても最後のWord
'         連携が成立しないため、LLMを呼ぶ前にOpenAnswerInWordと同じ親切
'         #ERR規約で案内する)。
'   ・ExportAnswerAsDoc(裁定D12)は「指示文があるときだけ」modGateway.CallLLM
'     (step_name="word_export")でMarkdown文書へ整形し、TryRibbonRun経由の
'     OpenWordMarkでWordを開く。指示文が空なら整形をスキップして
'     OpenAnswerInWord(直接転記)に委譲する(D12の空指示経路)。
'     引数の受け方はmodFeatures.InvokeFeatureの実装規約(Variant配列を
'     Application.Runの位置引数へ展開)に合わせた単純なByVal 2引数:
'       InvokeFeature("markdown", "ExportAnswerAsDoc", Array(回答本文, 指示文))
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

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCDD) & " 表示を整えています…"

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

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCDD) & " Wordで開いています…"

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
' ExportAnswerAsDoc - 回答本文を利用者の指示文に従いLLMでMarkdown文書化して
'   Wordで開く(裁定D12)。指示文が空なら整形をスキップして直接転記
'   (=OpenAnswerInWord相当)。UI側の想定フロー: ボタン押下→InputBoxで
'   指示文(例: お客様向けの回答文書風に/社内回覧用の要約に)→本関数。
'   引数はmodFeatures.InvokeFeatureの規約に合わせたByVal 2引数(上部の
'   設計判断コメント参照):
'     answerText  … 文書化する回答本文
'     instruction … 仕上げ方の指示文(空=整形なしで直接Wordへ)
'   戻り値: ""=成功 / "#ERR:..."=失敗(例外は出さない)。
' ----------------------------------------------------------------------------
Public Function ExportAnswerAsDoc(ByVal answerText As String, ByVal instruction As String) As String
    On Error GoTo Fail

    If LenB(Trim$(answerText)) = 0 Then
        ExportAnswerAsDoc = "#ERR:Wordに書き出す回答本文が空です"
        Exit Function
    End If

    ' 指示文が空なら整形をスキップして直接転記(裁定D12の空指示経路)。
    ' mock時の親切#ERRもOpenAnswerInWord側の既存実装がそのまま返す。
    If LenB(Trim$(instruction)) = 0 Then
        ExportAnswerAsDoc = OpenAnswerInWord(answerText)
        Exit Function
    End If

    ' mock_llm=TRUE のときはLLMを呼ぶ前に案内する(OpenAnswerInWordと同じ
    ' 親切#ERR規約。mockのLLM整形だけ成功させても最後のWord連携が成立
    ' しないため、経路の入口で止める)。
    If modConfig.GetBool("mock_llm", True) Then
        ExportAnswerAsDoc = "#ERR:mockモード(mock_llm=TRUE)ではWord連携を利用できません。" & _
            "config の mock_llm を FALSE にすると実際にWordで開けるようになります。"
        Exit Function
    End If

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCDD) & " 指示に沿って文書に整えています…"

    ' 整形プロンプト: 指示に従いMarkdown形式の文書に仕上げる。出典表記は保持
    ' (裁定D12)。effort/verbosityはconfigで調整可能(既定medium/medium)。
    ' モデルはmodGateway既定(recommended_model)に任せる。
    Dim eff As String: eff = modConfig.GetString("word_export_effort", "medium")
    Dim vrb As String: vrb = modConfig.GetString("word_export_verbosity", "medium")
    Dim latency As Long

    Dim formatted As String
    formatted = modGateway.CallLLM(BuildExportPrompt(answerText, instruction), _
                                   "word_export", eff, vrb, "", latency)

    ' CallLLMの失敗は"#ERR:E02xx:..."で返る(例外は出ない)ため、そのまま返す。
    If Left$(formatted, 5) = "#ERR:" Then
        modUIMain.SetStage ""
        ExportAnswerAsDoc = formatted
        Exit Function
    End If

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCDD) & " Wordで開いています…"

    Dim result As Variant
    result = modGateway.TryRibbonRun(WORD_FUNC_NAME, Array(formatted))

    modUIMain.SetStage ""

    Dim s As String
    s = SafeResultToString(result)

    ' 戻り値は台帳上「未使用」のため、TryRibbonRun側の失敗("#ERR:"始まり)のみ
    ' 失敗と判定する(確定仕様。上部コメントブロック参照)。
    If Left$(s, 5) = "#ERR:" Then
        ExportAnswerAsDoc = s
        Exit Function
    End If

    ExportAnswerAsDoc = ""
    Exit Function

Fail:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    modUIMain.SetStage ""
    modLog.LogError "E0202", "optMarkdown.ExportAnswerAsDoc", errDesc
    ExportAnswerAsDoc = "#ERR:Word文書化処理でエラーが発生しました: " & errDesc
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: optMarkdownの公開契約はPing/RenderMarkdownAt/
' OpenAnswerInWord/ExportAnswerAsDocのみ)
' ----------------------------------------------------------------------------

' 整形プロンプトの組み立て(裁定D12の趣旨: 「指示に従いMarkdown形式の文書に
' 仕上げる。出典表記は保持」)。回答に無い事実の追加を禁じ、出力をMarkdown
' 文書のみに限定する(前置きが付くとそのままWord文書の頭に載ってしまうため)。
Private Function BuildExportPrompt(ByVal answerText As String, ByVal instruction As String) As String
    Dim sb As String
    sb = "以下の回答本文を、利用者の指示に従ってMarkdown形式の文書に仕上げてください。" & vbLf
    sb = sb & "・見出し・箇条書き等のMarkdown記法で読みやすく整えること。" & vbLf
    sb = sb & "・本文中の出典表記([本棚:...]や[パック(...):...])は削除せずそのまま保持すること。" & vbLf
    sb = sb & "・回答本文に無い事実を追加しないこと。" & vbLf
    sb = sb & "・出力はMarkdown文書のみとし、前置きや説明文を付けないこと。" & vbLf & vbLf
    sb = sb & "## 利用者の指示" & vbLf & instruction & vbLf & vbLf
    sb = sb & "## 回答本文" & vbLf & answerText & vbLf
    BuildExportPrompt = sb
End Function

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
