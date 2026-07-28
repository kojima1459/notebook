Attribute VB_Name = "modRagParse"
Option Explicit

' ============================================================================
' modRagParse - 多段RAGのLLM応答パーサ(RAG_OVERHAUL_DESIGN.md §C/§D/§G-T7)
' ----------------------------------------------------------------------------
' 純ロジック(R4)。全関数は「寛容退化」契約: マーカー/タグが欠落・崩壊して
' いても例外を出さず、呼び出し側が安全に既定動作へ落ちられる値を返す
' (mock応答・旧モデル・形式無視の応答でもアプリを壊さないため)。
' タグ照合はすべて大文字小文字無視。
' ============================================================================

' ----------------------------------------------------------------------------
' ParseExpand - クエリ拡張応答から <standalone>/<subqueries>/<hyde> を抽出。
'   standaloneが取れたらTrue。取れなければ standalone="" でFalse
'   (呼び出し側が元質問へ退化する)。subsは空でも必ず初期化済みで返る。
' ----------------------------------------------------------------------------
Public Function ParseExpand(ByVal resp As String, ByRef standalone As String, _
                            ByRef subs() As String, ByRef hyde As String) As Boolean
    standalone = Trim$(TagInner(resp, "standalone"))
    hyde = Trim$(TagInner(resp, "hyde"))
    subs = ParseSubqueries(TagInner(resp, "subqueries"), 8)
    ParseExpand = (LenB(standalone) > 0)
End Function

' ----------------------------------------------------------------------------
' ParseSubqueries - 「a | b | c」形式を分割しTrim・空除去・先頭maxN件。
'   空入力は0要素配列(Split(vbNullString)。実機VBAで合法な唯一の空String配列)。
' ----------------------------------------------------------------------------
Public Function ParseSubqueries(ByVal s As String, ByVal maxN As Long) As String()
    Dim t As String: t = Trim$(s)
    If LenB(t) = 0 Or maxN < 1 Then
        ParseSubqueries = Split(vbNullString)
        Exit Function
    End If

    Dim raw() As String: raw = Split(t, "|")
    Dim outArr() As String
    ReDim outArr(0 To UBound(raw) - LBound(raw))
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = LBound(raw) To UBound(raw)
        Dim piece As String: piece = Trim$(raw(i))
        If LenB(piece) > 0 And cnt < maxN Then
            outArr(cnt) = piece
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then
        ParseSubqueries = Split(vbNullString)
    Else
        ReDim Preserve outArr(0 To cnt - 1)
        ParseSubqueries = outArr
    End If
End Function

' ----------------------------------------------------------------------------
' ParseRankOrder - <rank>3,1,7</rank> を候補番号列へ。1..nHits範囲外・重複・
'   非数値は無視。戻り値=有効件数(0なら呼び出し側が元順を維持)。
'   orderは0起点配列に1起点の候補番号を格納(0件時はダミー1要素で返す)。
' ----------------------------------------------------------------------------
Public Function ParseRankOrder(ByVal resp As String, ByVal nHits As Long, _
                               ByRef order() As Long) As Long
    ReDim order(0 To 0)
    ParseRankOrder = 0
    If nHits < 1 Then Exit Function

    Dim inner As String: inner = Trim$(TagInner(resp, "rank"))
    If LenB(inner) = 0 Then Exit Function

    Dim parts() As String: parts = Split(inner, ",")
    Dim seen() As Boolean: ReDim seen(1 To nHits)
    Dim tmp() As Long: ReDim tmp(0 To UBound(parts) - LBound(parts))
    Dim cnt As Long: cnt = 0

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim p As String: p = Trim$(parts(i))
        If LenB(p) > 0 Then
            If IsNumericDigits(p) Then
                Dim v As Long: v = CLng(Val(p))
                If v >= 1 And v <= nHits Then
                    If Not seen(v) Then
                        seen(v) = True
                        tmp(cnt) = v
                        cnt = cnt + 1
                    End If
                End If
            End If
        End If
    Next i

    If cnt = 0 Then Exit Function
    ReDim order(0 To cnt - 1)
    For i = 0 To cnt - 1
        order(i) = tmp(i)
    Next i
    ParseRankOrder = cnt
End Function

' ----------------------------------------------------------------------------
' ExtractAnswer - <answer>の中身をanswerへ、<thinking>をthinkingへ。
'   answerタグ欠落時は応答全体をanswerとしてFalse(寛容退化)。
'   閉じタグ欠落時は開始タグ以降すべてをanswerとする。
' ----------------------------------------------------------------------------
Public Function ExtractAnswer(ByVal resp As String, ByRef thinking As String, _
                              ByRef answer As String) As Boolean
    thinking = Trim$(TagInner(resp, "thinking"))

    ' 2026-07-28(レビュー M-2): <answer> は【thinking を閉じたあと】から探す。
    ' 資料本文に "<answer>…</answer>" という文字列が含まれていて、モデルが
    ' それを thinking の中で引用すると、先頭一致で拾ってしまい
    ' 【資料由来の偽の回答】が表示される(資料をそのまま信じるRAGでは、
    ' これはプロンプトインジェクションの経路そのものになる)。
    Dim searchFrom As Long: searchFrom = 1
    Dim pEndThink As Long
    pEndThink = InStr(1, resp, "</thinking>", vbTextCompare)
    If pEndThink > 0 Then searchFrom = pEndThink + Len("</thinking>")

    Dim body As String: body = TagInnerFrom(resp, "answer", searchFrom)
    If LenB(Trim$(body)) = 0 And searchFrom > 1 Then
        ' thinking の後に <answer> が無いモデルもある。その場合だけ
        ' 従来どおり全体から探す(退化はするが黙って空にはしない)。
        body = TagInner(resp, "answer")
    End If

    If LenB(Trim$(body)) > 0 Then
        answer = Trim$(body)
        ExtractAnswer = True
    Else
        ' <answer> が無いときのフォールバック。従来は応答全体をそのまま
        ' 表示していたため、thinking(モデルの思考過程)が利用者に見えていた。
        ' 思考は根拠ではないので、剥がしてから返す。
        answer = Trim$(StripThinking(resp))
        ExtractAnswer = False
    End If
End Function

' <thinking>…</thinking> を取り除く(閉じタグが無い場合は開始位置以降を捨てる)。
Private Function StripThinking(ByVal s As String) As String
    Dim p1 As Long: p1 = InStr(1, s, "<thinking>", vbTextCompare)
    If p1 = 0 Then
        StripThinking = s
        Exit Function
    End If
    Dim p2 As Long: p2 = InStr(p1, s, "</thinking>", vbTextCompare)
    If p2 = 0 Then
        StripThinking = Left$(s, p1 - 1)
    Else
        StripThinking = Left$(s, p1 - 1) & Mid$(s, p2 + Len("</thinking>"))
    End If
End Function

' TagInner の開始位置指定版。
Private Function TagInnerFrom(ByVal s As String, ByVal tagName As String, ByVal startAt As Long) As String
    If startAt < 1 Then startAt = 1
    If startAt > Len(s) Then Exit Function
    TagInnerFrom = TagInner(Mid$(s, startAt), tagName)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' <tag>…</tag> の中身を返す(大文字小文字無視・最初の1組)。開始タグのみで
' 閉じタグが無い場合は開始タグ以降の全部。タグが無ければ空文字列。
Private Function TagInner(ByVal s As String, ByVal tagName As String) As String
    Dim openTag As String: openTag = "<" & tagName & ">"
    Dim closeTag As String: closeTag = "</" & tagName & ">"

    Dim p1 As Long
    p1 = InStr(1, s, openTag, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(openTag)

    Dim p2 As Long
    p2 = InStr(p1, s, closeTag, vbTextCompare)
    If p2 = 0 Then
        TagInner = Mid$(s, p1)
    Else
        TagInner = Mid$(s, p1, p2 - p1)
    End If
End Function

' 半角数字のみで構成されているか(全角数字はTrim後のVal解釈に任せず不許可。
' LLM出力の連番は半角前提。全角が来た場合は安全側で無視する)。
Private Function IsNumericDigits(ByVal s As String) As Boolean
    If LenB(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsNumericDigits = True
End Function
