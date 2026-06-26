Attribute VB_Name = "modIntent"
Option Explicit

' ============================================================================
' modIntent - Step 0: 照会の意図解釈レイヤー (NotebookLM級の本質)
' ----------------------------------------------------------------------------
' 質問をそのまま検索せず、まずLLMで「営業が本当に聞きたいこと」を解釈する:
'   - intent           : 照会の真意 (1文)
'   - assumptions      : 回答にあたり置いた前提
'   - needs_clarification + clarifying_questions : 曖昧な時の質問返し
'   - search_queries   : 約款/ガイドラインを引くためのサブクエリ分解
'
' モード (config: intent_mode, main画面で切替):
'   smart  : 明確なら即答 / 重要な前提が欠けている時だけ確認質問
'   always : 毎回まず前提を確認
'
' この層が失敗しても本体は止めない: OK=False を返し、呼び出し側は
' 確認質問なしで通常パイプラインに進む (graceful degradation)。
' ============================================================================

Public Type IntentResult
    OK As Boolean
    Intent As String
    Assumptions As String         ' " / " 区切り
    NeedsClarification As Boolean
    ClarifyingText As String      ' 表示用 (番号付き複数行)
    SearchQueries As String       ' vbLf 区切り (ルーター補強用)
    LatencyMs As Long
    ErrorMsg As String
End Type

Public Function AnalyzeIntent(ByVal question As String, _
                              ByVal mode As String, _
                              ByVal history As String) As IntentResult
    Dim res As IntentResult
    On Error GoTo Trap

    Dim p As String: p = modPrompts.GetIntent()
    If LenB(p) = 0 Then
        res.OK = False
        res.ErrorMsg = "intent プロンプト未設定"
        AnalyzeIntent = res
        Exit Function
    End If

    If LenB(history) = 0 Then history = "(履歴なし)"
    p = Replace(p, "{mode}", mode)
    p = Replace(p, "{history}", history)
    p = Replace(p, "{question}", question)

    Dim raw As String
    raw = modRibbonGateway.CallLLM(p, "intent", res.LatencyMs)
    modRibbonGateway.LogDebugCall "intent", p, raw, res.LatencyMs

    If Left$(raw, 12) = "#LLM_ERROR: " Then
        res.OK = False
        res.ErrorMsg = raw
        AnalyzeIntent = res
        Exit Function
    End If

    ' ---- Parse JSON (dependency-free) ----
    res.Intent = JsonStr(raw, "intent")
    Dim asum() As String: asum = JsonArr(raw, "assumptions")
    res.Assumptions = JoinArr(asum, " / ")
    res.NeedsClarification = JsonBool(raw, "needs_clarification")
    Dim sq() As String: sq = JsonArr(raw, "search_queries")
    res.SearchQueries = JoinArr(sq, vbLf)

    Dim cq() As String: cq = JsonArr(raw, "clarifying_questions")
    Dim ct As String, i As Long, n As Long
    For i = LBound(cq) To UBound(cq)
        If LenB(Trim$(cq(i))) > 0 Then
            n = n + 1
            If LenB(ct) > 0 Then ct = ct & vbLf
            ct = ct & "  " & n & ". " & cq(i)
        End If
    Next i
    res.ClarifyingText = ct
    ' If the model said "needs clarification" but produced no questions, treat as not-needed.
    If res.NeedsClarification And LenB(ct) = 0 Then res.NeedsClarification = False

    res.OK = True
    AnalyzeIntent = res
    Exit Function

Trap:
    res.OK = False
    res.ErrorMsg = "intent解析エラー: " & Err.Description
    AnalyzeIntent = res
End Function

' ----------------------------------------------------------------------------
' Lightweight JSON extractors. Pure VBA (no JsonConverter dependency) so they
' run on hardened corporate PCs and Mac Excel identically. Best-effort: handles
' the shapes the intent prompt emits ("key":"...", "key":true, "key":[...]).
' ----------------------------------------------------------------------------
Private Function FindKey(ByVal s As String, ByVal key As String) As Long
    ' Return position just AFTER the colon following "key", or 0.
    Dim needle As String: needle = """" & key & """"
    Dim k As Long: k = InStr(1, s, needle, vbTextCompare)
    If k = 0 Then Exit Function
    Dim c As Long: c = InStr(k + Len(needle), s, ":")
    If c = 0 Then Exit Function
    FindKey = c + 1
End Function

Private Function JsonStr(ByVal s As String, ByVal key As String) As String
    Dim pos As Long: pos = FindKey(s, key)
    If pos = 0 Then Exit Function
    ' skip whitespace to first quote
    Do While pos <= Len(s)
        Dim ch As String: ch = Mid$(s, pos, 1)
        If ch = """" Then Exit Do
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Function
        pos = pos + 1
    Loop
    pos = pos + 1
    Dim out As String
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        If ch = "\" And pos < Len(s) Then
            Dim nx As String: nx = Mid$(s, pos + 1, 1)
            Select Case nx
                Case """": out = out & """"
                Case "n": out = out & vbLf
                Case "t": out = out & vbTab
                Case "\": out = out & "\"
                Case "/": out = out & "/"
                Case Else: out = out & nx
            End Select
            pos = pos + 2
        ElseIf ch = """" Then
            Exit Do
        Else
            out = out & ch
            pos = pos + 1
        End If
    Loop
    JsonStr = Trim$(out)
End Function

Private Function JsonBool(ByVal s As String, ByVal key As String) As Boolean
    Dim pos As Long: pos = FindKey(s, key)
    If pos = 0 Then Exit Function
    Dim seg As String: seg = LCase$(Mid$(s, pos, 8))
    JsonBool = (InStr(seg, "true") > 0)
End Function

' Returns array elements (quoted strings) as a String() . Empty array -> 1 empty elem.
Private Function JsonArr(ByVal s As String, ByVal key As String) As String()
    Dim empty_(0 To 0) As String
    Dim pos As Long: pos = FindKey(s, key)
    If pos = 0 Then
        JsonArr = empty_
        Exit Function
    End If
    Dim ob As Long: ob = InStr(pos, s, "[")
    If ob = 0 Then
        JsonArr = empty_
        Exit Function
    End If
    Dim cb As Long: cb = InStr(ob, s, "]")
    If cb = 0 Then cb = Len(s)
    Dim body As String: body = Mid$(s, ob + 1, cb - ob - 1)

    Dim items() As String
    ReDim items(0 To 63)
    Dim cnt As Long, i As Long
    Dim inQ As Boolean, cur As String
    For i = 1 To Len(body)
        Dim ch As String: ch = Mid$(body, i, 1)
        If ch = """" Then
            If inQ Then
                items(cnt) = cur: cnt = cnt + 1: cur = ""
                If cnt > UBound(items) Then ReDim Preserve items(0 To UBound(items) + 64)
                inQ = False
            Else
                inQ = True: cur = ""
            End If
        ElseIf inQ Then
            If ch = "\" And i < Len(body) Then
                ' keep escaped char literally enough
                i = i + 1: cur = cur & Mid$(body, i, 1)
            Else
                cur = cur & ch
            End If
        End If
    Next i
    If cnt = 0 Then
        JsonArr = empty_
        Exit Function
    End If
    ReDim Preserve items(0 To cnt - 1)
    JsonArr = items
End Function

Private Function JoinArr(ByRef arr() As String, ByVal sep As String) As String
    Dim out As String, i As Long
    For i = LBound(arr) To UBound(arr)
        If LenB(Trim$(arr(i))) > 0 Then
            If LenB(out) > 0 Then out = out & sep
            out = out & Trim$(arr(i))
        End If
    Next i
    JoinArr = out
End Function
