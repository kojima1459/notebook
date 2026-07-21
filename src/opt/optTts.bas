Attribute VB_Name = "optTts"
Option Explicit

' ============================================================================
' optTts - 読み上げ(opt機能・MASTER_SPEC §7.7)
' ----------------------------------------------------------------------------
' 役割:
'   回答テキストをAIリボンの読み上げ関数(ttsSpeak仮定)に渡して読み上げる。
'   この層の存在意義は「リボン関数の仕様が入手できなかったらこのファイルを
'   modules.jsonから消すだけで製品が成立する」こと(§7.7)。
'
' 設計判断:
'   ・リボン呼び出しは modGateway.TryRibbonRun のみを使う(自前で
'     Application.Run しない。R3・lintが検査する)。
'   ・失敗しても例外を外に出さない。戻り値は "" (成功) または "#ERR:..." の
'     文字列(§7.7)。
'   ・長文は最初の800字に切り詰め、「以降は省略しました」を付記してから
'     読み上げに渡す(発注元の指示どおり。読み上げの実演時間を抑える狙い)。
'   ・modUIMain.SetStage 以外のコア参照は行わない(基盤層のみ参照可・§7.7)。
' ============================================================================

' === SIGNATURE ASSUMPTION ===
'   仮定シグネチャ: ttsSpeak(text As String) -> 戻り値は空文字("")で成功、
'                   もしくはエラー文字列("error"等)を返す想定。
'   根拠: MASTER_SPEC §7.7 optTts.bas の記載「ttsSpeak仮定: (text)」のみが
'     根拠であり、他に一次情報は無い。第2引数(音声種別/速度等)や戻り値の
'     成功時フォーマットは未確認。リボンちゃんの他の関数(ChatGPT/
'     GetEmbeddings)が「成功時は文字列、失敗時はエラーを示す文字列」を
'     返す設計であることに倣い、同型を仮定した。
'   仕様回答が来たら直す行:
'     ・引数の数/順序が違う場合 -> BuildTtsArgs() の Array(...) 部分のみ修正。
'     ・戻り値の成功判定基準が違う場合(例: 空文字ではなく"OK"を返す等)
'       -> IsTtsSuccess() の判定式のみ修正。
'     ・関数名自体が違う場合(ttsSpeak以外) -> Private Const RIBBON_FUNC_NAME
'       の値のみ修正。
' ================================================================================

Private Const RIBBON_FUNC_NAME As String = "ttsSpeak"
Private Const MAX_SPEAK_CHARS As Long = 800
Private Const TRUNCATE_NOTICE As String = "(以降は省略しました)"

Public Function Ping() As Boolean
    Ping = True
End Function

' ----------------------------------------------------------------------------
' SpeakAnswer - 回答テキストを読み上げる(MASTER_SPEC §7.7)。
'   長文は先頭800字+「以降は省略しました」に切り詰めてから渡す。
'   戻り値: ""=成功 / "#ERR:..."=失敗(例外は出さない)。
' ----------------------------------------------------------------------------
Public Function SpeakAnswer(ByVal Text As String) As String
    On Error GoTo Fail

    Dim t As String
    t = PrepareSpeechText(Text)
    If LenB(t) = 0 Then
        SpeakAnswer = "#ERR:読み上げる文章がありません"
        Exit Function
    End If

    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDD0A) & " 読み上げ中…"

    Dim result As Variant
    result = modGateway.TryRibbonRun(RIBBON_FUNC_NAME, BuildTtsArgs(t))

    modUIMain.SetStage ""

    If VarType(result) = vbString Then
        Dim s As String
        s = CStr(result)
        If Left$(s, 5) = "#ERR:" Then
            SpeakAnswer = s
            Exit Function
        End If
        If Not IsTtsSuccess(s) Then
            SpeakAnswer = "#ERR:読み上げに失敗しました: " & modUtil.SafeLeft(s, 200)
            Exit Function
        End If
    End If

    SpeakAnswer = ""
    Exit Function

Fail:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    modLog.LogError "E0202", "optTts.SpeakAnswer", errDesc
    SpeakAnswer = "#ERR:読み上げ処理でエラーが発生しました: " & errDesc
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: optTtsの公開契約はPing/SpeakAnswerのみ)
' ----------------------------------------------------------------------------

' 800字を超える場合は切り詰めて末尾に省略通知を付ける。
Private Function PrepareSpeechText(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If Len(t) <= MAX_SPEAK_CHARS Then
        PrepareSpeechText = t
    Else
        PrepareSpeechText = Left$(t, MAX_SPEAK_CHARS) & TRUNCATE_NOTICE
    End If
End Function

' === SIGNATURE ASSUMPTION: 引数配列の組み立て。仕様が判明したらここだけ直す ===
Private Function BuildTtsArgs(ByVal t As String) As Variant
    BuildTtsArgs = Array(t)
End Function

' === SIGNATURE ASSUMPTION: 成功判定。仕様が判明したらここだけ直す ===
' 現状の仮定: 空文字列は成功、"error"(大小文字無視)は失敗、それ以外の
' 非空文字列は(内容不明だが)成功扱いにする(リボン側が確認メッセージ的な
' 文字列を返す可能性を考慮した楽観判定)。
Private Function IsTtsSuccess(ByVal s As String) As Boolean
    If LenB(Trim$(s)) = 0 Then
        IsTtsSuccess = True
    ElseIf StrComp(Trim$(s), "error", vbTextCompare) = 0 Then
        IsTtsSuccess = False
    Else
        IsTtsSuccess = True
    End If
End Function
