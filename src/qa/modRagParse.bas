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
' "#ERR:…" 応答の判定と文言化(2026-08-03 R14-G1: modAsk から移設)
' ----------------------------------------------------------------------------
' modGateway.CallLLM の失敗は "#ERR:コード:説明" という応答文字列で返る契約
' なので、その読み取りも「LLM応答のパース」であり、ここが正しい置き場所。
' modAsk / modAskRetrieve / modAskThorough の3モジュールが同じ判定を使う。
' 副作用は無い(modLog.FriendlyMessage は Select Case だけの純ロジック)。
Public Function IsErrorResponse(ByVal s As String) As Boolean
    IsErrorResponse = (Left$(s, 5) = "#ERR:")
End Function

' "#ERR:…" を利用者向けの1文へ。コードが読めないときは E0202(API失敗)。
Public Function BuildErrorAnswer(ByVal errResp As String) As String
    Dim code As String
    code = ExtractErrorCode(errResp)
    If LenB(code) = 0 Then code = "E0202"
    BuildErrorAnswer = modLog.FriendlyMessage(code) & vbLf & "(コード: " & code & ")"
End Function

' "#ERR:E0202:説明..." -> "E0202"
Private Function ExtractErrorCode(ByVal s As String) As String
    If Left$(s, 5) <> "#ERR:" Then Exit Function
    Dim rest As String
    rest = Mid$(s, 6)
    Dim p As Long
    p = InStr(rest, ":")
    If p = 0 Then
        ExtractErrorCode = rest
    Else
        ExtractErrorCode = Left$(rest, p - 1)
    End If
End Function

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
' 複合質問の分解(2026-08-05 R16-3A)。段0判定の応答パーサ。
'   出力契約: <verdict>single|parts|clarify</verdict>
'             <parts>副質問1 | 副質問2 | 副質問3</parts>
'             <options>読み方の候補1 | 候補2 | 候補3</options>
'
' ParseDecomposeVerdict - 判定を "single" / "parts" / "clarify" のいずれかへ。
'   タグ欠落・空・知らない語・"#ERR:" はすべて "single"(=従来の入念フローへ
'   無害フォールバック)。分解は「効けば速く正確になる」だけの上積みなので、
'   読めない応答で分解へ倒すより、確実に答えが出る側へ倒すのが正しい退化。
' ----------------------------------------------------------------------------
Public Function ParseDecomposeVerdict(ByVal resp As String) As String
    ParseDecomposeVerdict = "single"
    If IsErrorResponse(resp) Then Exit Function

    Dim v As String
    v = LCase$(Trim$(TagInner(resp, "verdict")))
    If v = "parts" Then
        ParseDecomposeVerdict = "parts"
    ElseIf v = "clarify" Then
        ParseDecomposeVerdict = "clarify"
    End If
End Function

' ----------------------------------------------------------------------------
' ParseParts - <parts>a | b | c</parts> を副質問の配列へ。戻り値=有効件数。
'   分割・Trim・空要素除去・maxN件での切り詰めは ParseSubqueries と同じ流儀
'   (区切りは "|" 1種。同じ書式を2実装に分けない)。
'   戻り値を配列型にしないのは LibreOffice の制約(tools/run_lo_tests.py
'   技術メモ6)。呼び出し側は ByRef の parts() で受ける。
' ----------------------------------------------------------------------------
Public Function ParseParts(ByVal resp As String, ByVal maxN As Long, _
                           ByRef parts() As String) As Long
    parts = ParseSubqueries(TagInner(resp, "parts"), maxN)

    Dim n As Long
    On Error Resume Next
    n = UBound(parts) - LBound(parts) + 1
    On Error GoTo 0
    If n < 0 Then n = 0
    ParseParts = n
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
' ParseQuestionLines - 質問例オンデマンド生成(2026-08-03 R14-7a)の応答パーサ。
'   1行1問想定のLLM応答をTrim・空行除去・先頭の番号/箇条書き記号の除去の
'   うえ、先頭maxN件だけ拾い"|"区切りへ畳む(modStarter.Drawが読む形式は
'   modSeed.SeedQuestionsと同じ"|"区切りのため、描画路を1本に保てる)。
'
' R14-G5: 質問文そのものに "|" が入っていると、区切り文字と衝突して1問が
'   2つのボタンへ割れる(後半は文の途中から始まる意味不明なボタンになる)。
'   区切りに使う以上、要素側からは必ず落とす。消すのではなく全角の "／" へ
'   置換して、元が並列の列挙だったことを読めるまま残す。
' ----------------------------------------------------------------------------
Public Function ParseQuestionLines(ByVal resp As String, ByVal maxN As Long) As String
    If maxN < 1 Then Exit Function

    Dim lines() As String
    lines = Split(Replace(resp, vbCr, vbLf), vbLf)

    Dim sb As String
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If cnt >= maxN Then Exit For
        Dim q As String: q = StripQuestionNumbering(Trim$(lines(i)))
        q = Trim$(Replace(q, "|", ChrW(&HFF0F)))
        If LenB(q) > 0 Then
            sb = sb & IIf(LenB(sb) > 0, "|", "") & q
            cnt = cnt + 1
        End If
    Next i
    ParseQuestionLines = sb
End Function

' 行頭の箇条書き記号("・""-""*")と連番("1.""1)""1、"等)を1回だけ剥がす。
Private Function StripQuestionNumbering(ByVal s As String) As String
    Dim t As String: t = s
    Dim lead As String: lead = Left$(t, 1)
    If lead = ChrW(&H30FB) Or lead = "-" Or lead = "*" Then
        t = Trim$(Mid$(t, 2))
    End If

    Dim i As Long: i = 1
    Do While i <= Len(t) And Mid$(t, i, 1) >= "0" And Mid$(t, i, 1) <= "9"
        i = i + 1
    Loop
    If i > 1 And i <= Len(t) Then
        Dim sep As String: sep = Mid$(t, i, 1)
        If sep = "." Or sep = ")" Or sep = " " Or sep = ChrW(&H3001) _
           Or sep = ChrW(&HFF0E) Or sep = ChrW(&HFF09) Then
            t = Trim$(Mid$(t, i + 1))
        End If
    End If
    StripQuestionNumbering = t
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
