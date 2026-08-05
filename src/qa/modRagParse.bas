Attribute VB_Name = "modRagParse"
Option Explicit

' ============================================================================
' modRagParse - 多段RAGのLLM応答パーサ(RAG_OVERHAUL_DESIGN.md §C/§D/§G-T7)
' ----------------------------------------------------------------------------
' 純ロジック(R4)。全関数は「寛容退化」契約: タグが欠落・崩壊しても例外を出さず、
' 呼び出し側が既定動作へ落ちられる値を返す。タグ照合は大小無視。
' ============================================================================

' 俯瞰質問(R17 Phase2)で選べる章数の絶対上限(ParseChapterPick が丸める)。
Private Const PICK_HARD_MAX As Long = 4

' ----------------------------------------------------------------------------
' "#ERR:…" 応答の判定と文言化(2026-08-03 R14-G1: modAsk から移設)
' ----------------------------------------------------------------------------
' modGateway.CallLLM の失敗は "#ERR:コード:説明" の応答文字列で返る契約なので、
' その読み取りも「LLM応答のパース」。3モジュールが同じ判定を使う。
Public Function IsErrorResponse(ByVal s As String) As Boolean
    IsErrorResponse = (Left$(s, 5) = "#ERR:")
End Function

' "#ERR:…" を利用者向けの1文へ。コードが読めないときは E0202(API失敗)。
' 2026-08-05(R16H FA-2): #ERR以外(逆質問等)は素通しして偽E0202にしない
' (答えを作らなかっただけで障害ではない。空文字は従来どおりE0202)。
Public Function BuildErrorAnswer(ByVal errResp As String) As String
    If LenB(Trim$(errResp)) > 0 And Not IsErrorResponse(errResp) Then
        BuildErrorAnswer = errResp
        Exit Function
    End If

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
'   standaloneが取れたらTrue(取れなければ空でFalse=呼び出し側は元質問へ退化)。
'   subsは空でも必ず初期化済みで返る。
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
'   出力契約: <verdict>single|parts|clarify|global</verdict> /
'             <parts>副質問1 | 副質問2</parts> / <options>候補1 | 候補2</options>
' ParseDecomposeVerdict - 判定を "single"/"parts"/"clarify"/"global" のどれかへ。
'   タグ欠落・空・知らない語・"#ERR:" はすべて "single"(=従来の入念フローへ
'   無害フォールバック)。読めない応答で分解へ倒すより確実に答えが出る側へ。
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
    ElseIf v = "global" Then
        ' R17 Phase2: 俯瞰(全体像・一覧・「全部教えて」型)。この語を知らない
        ' 旧プロンプト・旧モデルの応答は従来どおり single へ落ちる。
        ParseDecomposeVerdict = "global"
    End If
End Function

' ----------------------------------------------------------------------------
' ParseParts - <parts>a | b | c</parts> を副質問の配列へ。戻り値=有効件数。
'   切り詰めは ParseSubqueries と同じ流儀(区切りは "|" 1種)。戻り値を配列型に
'   しないのは LibreOffice の制約(tools/run_lo_tests.py 技術メモ6)。
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
' ParseOptions - <options>読み方1 | 読み方2</options> を選択肢の配列へ。
'   戻り値=有効件数。流儀は ParseParts と同じ。タグ欠落・空は0件で返り、
'   呼び出し元(modAskMulti)は逆質問を諦めて従来の入念フローへ落ちる
'   (「1件だけ」の扱いも呼び出し元が持つ。ここは読めたぶんだけ返す)。
' ----------------------------------------------------------------------------
Public Function ParseOptions(ByVal resp As String, ByVal maxN As Long, _
                             ByRef opts() As String) As Long
    opts = ParseSubqueries(TagInner(resp, "options"), maxN)

    Dim n As Long
    On Error Resume Next
    n = UBound(opts) - LBound(opts) + 1
    On Error GoTo 0
    If n < 0 Then n = 0
    ParseOptions = n
End Function

' ----------------------------------------------------------------------------
' 章単位要約(2026-08-05 R17 Phase2)。取込時と俯瞰質問の2つの応答パーサ。
'   出力契約: <summary>要約</summary><keywords>語1|語2</keywords> と
'             <pick>資料名::章キー|資料名::章キー</pick>
' ParseOutlineResp - 章要約の応答を summary / keywords へ。True=summary が1字
'   以上読めた(keywords は空でもよい)。False="#ERR:"/タグ欠落/中身が空で、
'   呼び出し元(modOutlineBuild)はその章を「(要約失敗)」として保存し次章へ進む
'   (1章の失敗で254頁ぶんを捨てない寛容退化。行ごと落とすと「聞いたのに答えが
'   無い章」に気付けない=憲章§3-3)。keywords は "|" 区切りへ正規化・上限8語。
' ----------------------------------------------------------------------------
Public Function ParseOutlineResp(ByVal resp As String, ByRef outSummary As String, _
                                 ByRef outKeywords As String) As Boolean
    outSummary = ""
    outKeywords = ""
    If IsErrorResponse(resp) Then Exit Function

    Dim s As String: s = Trim$(TagInner(resp, "summary"))
    If LenB(s) = 0 Then Exit Function

    outSummary = s
    outKeywords = NormalizeKeywordLine(TagInner(resp, "keywords"))
    ParseOutlineResp = True
End Function

' "a、b, c|d" のような列を "a|b|c|d" へ寄せる(空要素除去・上限8語)。
Private Function NormalizeKeywordLine(ByVal s As String) As String
    Dim t As String: t = Trim$(s)
    If LenB(t) = 0 Then Exit Function
    t = Replace(t, ChrW(&H3001), "|")     ' 読点
    t = Replace(t, ChrW(&HFF0C), "|")     ' 全角カンマ
    t = Replace(t, ",", "|")

    Dim arr() As String: arr = Split(t, "|")
    Dim outS As String
    Dim cnt As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Dim piece As String: piece = Trim$(arr(i))
        If LenB(piece) > 0 And cnt < 8 Then
            If LenB(outS) > 0 Then outS = outS & "|"
            outS = outS & piece
            cnt = cnt + 1
        End If
    Next i
    NormalizeKeywordLine = outS
End Function

' ----------------------------------------------------------------------------
' ParseChapterPick - <pick>資料名::章キー|…</pick> を選択章の配列へ。
'   戻り値=有効件数(0=1章も選ばれず=俯瞰は不発→従来フローへ落ちる)。
'   "::" を含まない要素は捨てる(資料名か章キーが欠けた指定は別の資料の同じ
'   章名に当たり得る)。maxN は 1..PICK_HARD_MAX(4)へ丸める(章が増えるほど
'   1章あたりの取り分が減り、どの章も途中で切れた抜粋になるため)。
' ----------------------------------------------------------------------------
Public Function ParseChapterPick(ByVal resp As String, ByVal maxN As Long, _
                                 ByRef picks() As String) As Long
    Dim lim As Long: lim = maxN
    If lim < 1 Then lim = PICK_HARD_MAX
    If lim > PICK_HARD_MAX Then lim = PICK_HARD_MAX

    picks = ParseSubqueries(TagInner(resp, "pick"), lim)

    Dim n As Long
    On Error Resume Next
    n = UBound(picks) - LBound(picks) + 1
    On Error GoTo 0
    If n < 1 Then Exit Function

    Dim lo As Long: lo = LBound(picks)
    Dim k As Long: k = 0
    Dim i As Long
    For i = lo To lo + n - 1
        Dim t As String: t = Trim$(picks(i))
        Dim p As Long: p = InStr(t, "::")
        If p > 1 And p + 2 <= Len(t) Then
            picks(lo + k) = t
            k = k + 1
        End If
    Next i
    ParseChapterPick = k
End Function

' ----------------------------------------------------------------------------
' 用語の名寄せ(R17 Phase3・設計書§3)。出力契約 <syn>表記>正規形|…</syn>。
' ParseSynResp - 応答を term/canonical の並行配列へ。戻り値=有効ペア数。
'   ">"無し・どちらか空の要素は1件ずつ破棄(読めた分だけ返す寛容退化)。
'   タグ欠落・"#ERR:"はすべて0件。配列は【0始まり】で返る。
' ----------------------------------------------------------------------------
Public Function ParseSynResp(ByVal resp As String, ByRef outTerms() As String, _
                             ByRef outCanons() As String) As Long
    outTerms = Split(vbNullString)
    outCanons = Split(vbNullString)
    If IsErrorResponse(resp) Then Exit Function

    Dim inner As String: inner = Trim$(TagInner(resp, "syn"))
    If LenB(inner) = 0 Then Exit Function

    Dim raw() As String: raw = Split(inner, "|")
    Dim tmpT() As String: ReDim tmpT(0 To UBound(raw) - LBound(raw))
    Dim tmpC() As String: ReDim tmpC(0 To UBound(raw) - LBound(raw))
    Dim cnt As Long
    Dim i As Long
    Dim t As String, c As String
    For i = LBound(raw) To UBound(raw)
        If SplitSynPair(raw(i), t, c) Then
            tmpT(cnt) = t
            tmpC(cnt) = c
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then Exit Function
    ReDim Preserve tmpT(0 To cnt - 1)
    ReDim Preserve tmpC(0 To cnt - 1)
    outTerms = tmpT
    outCanons = tmpC
    ParseSynResp = cnt
End Function

' ----------------------------------------------------------------------------
' MergeSynPairs - 旧CSVへ新CSVを合流させた統合CSVを返す(2026-08-05 R17H FA-1)。
'   どちらも "term>canonical|term>canonical"(=modSynonymStore.ReadMapCsv の形)。
'   マージ規則(ゴールデンで固定): 1.新CSVと term が重なる旧行は落とす=
'   【既存termは上書き】(大小無視) 2.残った旧行を元の順序のまま先に、その
'   あとへ新行を順に(順序安定) 3.同じtermは先勝ちで1回だけ 4.">"無し/
'   どちらか空の要素は捨てる 5.旧が空なら新だけ・両方空なら空文字。
'   純関数にした理由: 実体は modSynonymStore.MergeAndSave にあり、ParseSynResp
'   が返す【0始まり】の配列を 1〜n で読んでいたため実データでは毎回「添字が
'   範囲外」で握り潰され、synonyms が永久に0行だった(A-H1)。
' ----------------------------------------------------------------------------
Public Function MergeSynPairs(ByVal oldCsv As String, ByVal newCsv As String) As String
    Dim newT As String: newT = SynTermBox(newCsv)
    Dim outS As String
    Dim usedT As String: usedT = vbLf
    Dim parts() As String
    Dim i As Long
    Dim t As String, c As String

    If LenB(oldCsv) > 0 Then
        parts = Split(oldCsv, "|")
        For i = LBound(parts) To UBound(parts)
            If SplitSynPair(parts(i), t, c) Then
                If InStr(1, newT, vbLf & LCase$(t) & vbLf, vbBinaryCompare) = 0 Then
                    AppendSynPair outS, usedT, t, c
                End If
            End If
        Next i
    End If

    If LenB(newCsv) > 0 Then
        parts = Split(newCsv, "|")
        For i = LBound(parts) To UBound(parts)
            If SplitSynPair(parts(i), t, c) Then AppendSynPair outS, usedT, t, c
        Next i
    End If
    MergeSynPairs = outS
End Function

' "表記>正規形" を2つへ割る(">"無し・どちらか空は False=その要素は捨てる)。
Private Function SplitSynPair(ByVal piece As String, ByRef outTerm As String, _
                              ByRef outCanon As String) As Boolean
    outTerm = ""
    outCanon = ""
    Dim s As String: s = Trim$(piece)
    If LenB(s) = 0 Then Exit Function
    Dim p As Long: p = InStr(s, ">")
    If p < 2 Or p >= Len(s) Then Exit Function
    outTerm = Trim$(Left$(s, p - 1))
    outCanon = Trim$(Mid$(s, p + 1))
    SplitSynPair = (LenB(outTerm) > 0 And LenB(outCanon) > 0)
End Function

' CSV に含まれる term の一覧(vbLf区切り・小文字)。大小違いは同じtermとして扱う。
Private Function SynTermBox(ByVal csv As String) As String
    Dim box As String: box = vbLf
    SynTermBox = box
    If LenB(csv) = 0 Then Exit Function
    Dim arr() As String: arr = Split(csv, "|")
    Dim i As Long
    Dim t As String, c As String
    For i = LBound(arr) To UBound(arr)
        If SplitSynPair(arr(i), t, c) Then
            If InStr(1, box, vbLf & LCase$(t) & vbLf, vbBinaryCompare) = 0 Then
                box = box & LCase$(t) & vbLf
            End If
        End If
    Next i
    SynTermBox = box
End Function

' 統合CSVへ1組足す(同じtermは先勝ちで1回だけ)。
Private Sub AppendSynPair(ByRef outS As String, ByRef usedT As String, _
                          ByVal term As String, ByVal canon As String)
    Dim k As String: k = LCase$(term)
    If InStr(1, usedT, vbLf & k & vbLf, vbBinaryCompare) > 0 Then Exit Sub
    usedT = usedT & k & vbLf
    If LenB(outS) > 0 Then outS = outS & "|"
    outS = outS & term & ">" & canon
End Sub

' ----------------------------------------------------------------------------
' ExpandQueryBySyn - synonymsの地図(mapCsv=ReadMapCsvの戻り値そのもの)を使い、
'   質問文 q 中の語に一致した同義語を最大maxAdd件・半角空白区切りで末尾へ追記。
'   ・一致判定は modSparse.NormalizeForSearch を両辺に通してから InStr
'     (全角/半角・大小の表記ゆれを吸収。取込側section_pathと同じ式)。
'   ・双方向: 質問に term があれば canonical を、canonical があれば term を足す。
'   ・自己一致除外(正規化後に同じペアは何も足さない)。既にある語・追記済みの
'     語は二重に足さない。mapCsv が空・maxAdd<1 は無操作で q をそのまま返す。
' ----------------------------------------------------------------------------
Public Function ExpandQueryBySyn(ByVal q As String, ByVal mapCsv As String, _
                                 ByVal maxAdd As Long) As String
    ExpandQueryBySyn = q
    If LenB(Trim$(mapCsv)) = 0 Or maxAdd < 1 Then Exit Function

    Dim normQ As String: normQ = modSparse.NormalizeForSearch(q)
    If LenB(normQ) = 0 Then Exit Function

    Dim outQ As String: outQ = q
    Dim addedBox As String: addedBox = vbLf   ' 追記済み語(正規化後)の重複防止
    Dim cnt As Long

    Dim pairs() As String: pairs = Split(mapCsv, "|")
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs)
        If cnt >= maxAdd Then Exit For
        Dim term As String, canon As String
        If SplitSynPair(pairs(i), term, canon) Then
            Dim nTerm As String: nTerm = modSparse.NormalizeForSearch(term)
            Dim nCanon As String: nCanon = modSparse.NormalizeForSearch(canon)
            If StrComp(nTerm, nCanon, vbBinaryCompare) <> 0 Then   ' 自己一致除外
                If InStr(1, normQ, nTerm, vbBinaryCompare) > 0 Then
                    AppendSynWord outQ, addedBox, cnt, canon, nCanon, normQ, maxAdd
                ElseIf InStr(1, normQ, nCanon, vbBinaryCompare) > 0 Then
                    AppendSynWord outQ, addedBox, cnt, term, nTerm, normQ, maxAdd
                End If
            End If
        End If
    Next i
    ExpandQueryBySyn = outQ
End Function

' word をoutQへ半角空白区切りで追記(既にある語・追記済みの語は足さない)。
Private Sub AppendSynWord(ByRef outQ As String, ByRef addedBox As String, ByRef cnt As Long, _
                          ByVal word As String, ByVal normWord As String, _
                          ByVal normQ As String, ByVal maxAdd As Long)
    If cnt >= maxAdd Then Exit Sub
    If InStr(1, normQ, normWord, vbBinaryCompare) > 0 Then Exit Sub
    If InStr(1, addedBox, vbLf & normWord & vbLf, vbBinaryCompare) > 0 Then Exit Sub
    outQ = outQ & " " & word
    addedBox = addedBox & normWord & vbLf
    cnt = cnt + 1
End Sub

' ----------------------------------------------------------------------------
' ParseChoiceNumbers - 「番号で返信」への返事から選んだ番号を読み取る(R16-3B)。
' 戻り値 = 有効な番号を入力順・重複なしで "," 連結("1,3")。1つも無ければ空
'   文字列 = 「番号ではなく質問を書き直した」扱い。
' 受ける形: "1" / "1と3" / "①と③" / "1,3" / "2、3" / "1 3" / "１ ３" /
'           "1-3" / "2-②" / "(1)(3)" / "1ー3"
'   ・半角0-9・全角０-９は連続桁を1つの数("13"は13)。丸数字①〜⑳は1文字1数。
'   ・区切りとして認める文字は IsChoiceGap が唯一の持ち主。それ以外が1つでも
'     混じれば【書き直し】とみなして空を返す(「1番の話」を番号選択と誤読
'     すると打った質問が黙って捨てられる)。1..maxN の範囲外は無視。
' ----------------------------------------------------------------------------
Public Function ParseChoiceNumbers(ByVal s As String, ByVal maxN As Long) As String
    Dim t As String: t = Trim$(s)
    If LenB(t) = 0 Or maxN < 1 Then Exit Function

    Dim outS As String
    Dim cur As String                ' 読みかけの数字列(半角へ正規化済み)
    Dim i As Long
    For i = 1 To Len(t) + 1
        Dim cp As Long
        If i > Len(t) Then
            cp = 32                  ' 番兵: 末尾の読みかけを必ず確定させる
        Else
            cp = AscW(Mid$(t, i, 1))
            If cp < 0 Then cp = cp + 65536       ' AscWは符号付きで返る
        End If

        If cp >= 48 And cp <= 57 Then
            cur = cur & Chr$(cp)                          ' 0-9
        ElseIf cp >= 65296 And cp <= 65305 Then
            cur = cur & Chr$(cp - 65296 + 48)             ' ０-９
        Else
            If LenB(cur) > 0 Then
                AppendChoice outS, CLng(Val(cur)), maxN
                cur = ""
            End If
            If cp >= 9312 And cp <= 9331 Then
                AppendChoice outS, cp - 9312 + 1, maxN    ' ①-⑳
            ElseIf Not IsChoiceGap(cp) Then
                Exit Function                             ' 数字でも区切りでもない=書き直し
            End If
        End If
    Next i

    ParseChoiceNumbers = outS
End Function

' 番号の区切りとして認める文字か(ParseChoiceNumbers 専用)。
' 2026-08-05(R16H FA-7 / A-M7・B-M4): ハイフン類と括弧を足して
' modClarify.IsChoiceSeparator と揃えた。資料の聞き返しは「2-②」と打つよう
' 教えているのにこちらは "-" を知らず、同じ癖で「1-3」と打った人の返事が
' 【まるごと書き直し扱い】で捨てられていた。「1-3」は 1 と 3 の2つを選んだ意味。
' ----------------------------------------------------------------------------
Private Function IsChoiceGap(ByVal cp As Long) As Boolean
    Select Case cp
        Case 32, 9, 12288                        ' 半角スペース / タブ / 全角スペース
            IsChoiceGap = True
        Case 44, 46, 47                          ' , . /
            IsChoiceGap = True
        Case 65292, 65294, 65295                 ' ，．／
            IsChoiceGap = True
        Case 12289, 12290, 12539                 ' 、。・
            IsChoiceGap = True
        Case 12392                               ' と
            IsChoiceGap = True
        Case 45, 65293, 8722, 12540              ' - － −(マイナス) ー(長音)
            IsChoiceGap = True
        Case 40, 41, 65288, 65289                ' ( ) （ ）
            IsChoiceGap = True
    End Select
End Function

' 有効範囲内の番号だけを重複なしで "," 連結する。
Private Sub AppendChoice(ByRef outS As String, ByVal v As Long, ByVal maxN As Long)
    If v < 1 Or v > maxN Then Exit Sub
    If InStr(1, "," & outS & ",", "," & CStr(v) & ",") > 0 Then Exit Sub
    If LenB(outS) > 0 Then outS = outS & ","
    outS = outS & CStr(v)
End Sub

' ----------------------------------------------------------------------------
' HasCompoundSignal - 質問文が複合質問らしいかの軽量シグナル検知(R18-7a)。
' 実機第5報⑦: 「免責は?保険料は?」(9字)は decompose_min_chars(既定25)にも
' IsTooVague(≦10字かつ低スコア)にも掛からず、段0判定が一度も呼ばれない狭間に
' 落ちていた。長さだけを論点数のproxyにしていた ShouldDecompose のゲートを
' このシグナルとのORで迂回する(呼び出し側は modAskMulti.DecomposeGate)。
' S1=？/?が2個以上 / S2=。/．区切りの非空節が2個以上、のOR。「と/や」区切りの
' 名詞列は最頻出の助詞で誤検知率が高く不採用(agent4調査報告§3-3)。
' ----------------------------------------------------------------------------
Public Function HasCompoundSignal(ByVal q As String) As Boolean
    Dim t As String: t = Trim$(q)
    If LenB(t) = 0 Then Exit Function

    ' S1: ？/? の出現回数
    Dim qMarks As Long
    qMarks = (Len(t) - Len(Replace(t, "？", ""))) + (Len(t) - Len(Replace(t, "?", "")))
    If qMarks >= 2 Then HasCompoundSignal = True: Exit Function

    ' S2: 。/．区切りの非空セグメント数
    Dim parts() As String
    parts = Split(Replace(t, "．", "。"), "。")
    Dim i As Long, nSeg As Long
    For i = LBound(parts) To UBound(parts)
        If LenB(Trim$(parts(i))) > 0 Then nSeg = nSeg + 1
    Next i
    HasCompoundSignal = (nSeg >= 2)
End Function

' ----------------------------------------------------------------------------
' HasGlobalSignal - 質問文が俯瞰(全体像・一覧)を求めているかの語彙シグナル
'   (2026-08-05 R17H FA-6 / A-M7・B-H2)。InStr判定だけの純関数。
' 「全体像は?」(6字)のような俯瞰の質問は、短いというだけで段0判定
' (modAskMulti.DecomposeGate)を呼ばれず、さらに IsTooVague の聞き返しに
' 吸われて【俯瞰が一度も試されない】。長さは俯瞰の proxy にならない。
' 語彙は「その語が出たら全体を求めている」と言い切れるものだけを採る
' (「全て」のような単独で普通に使う短語は誤爆するので入れない)。誤発動の
' 実害は段0が1回増えることと聞き返しをしないことだけ(verdict=global を
' 出すのは後段のLLMで、ここは呼ぶかどうかの門にすぎない)。
' ----------------------------------------------------------------------------
Public Function HasGlobalSignal(ByVal q As String) As Boolean
    Dim t As String: t = Trim$(q)
    If LenB(t) = 0 Then Exit Function

    Dim words As Variant
    words = Array("全体像", "全体の流れ", "全体を通し", "全部教え", "ぜんぶ教え", _
                  "すべて教え", "全て教え", "一覧", "まとめて", "どんなこと", _
                  "どんな種類", "概要", "何が書いて", "どういう構成")
    Dim i As Long
    For i = LBound(words) To UBound(words)
        If InStr(1, t, CStr(words(i)), vbTextCompare) > 0 Then
            HasGlobalSignal = True
            Exit Function
        End If
    Next i
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
' ExtractAnswer - <answer>の中身をanswerへ、<thinking>をthinkingへ。タグ欠落時は
'   応答全体をanswerとしてFalse。閉じタグ欠落時は開始タグ以降すべてをanswerへ。
' ----------------------------------------------------------------------------
Public Function ExtractAnswer(ByVal resp As String, ByRef thinking As String, _
                              ByRef answer As String) As Boolean
    thinking = Trim$(TagInner(resp, "thinking"))

    ' 2026-07-28(レビュー M-2): <answer> は【thinking を閉じたあと】から探す。
    ' 資料本文の "<answer>…</answer>" をモデルが thinking 内で引用すると、
    ' 先頭一致で拾って【資料由来の偽の回答】が出る(RAGでのインジェクション)。
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
'   1行1問想定の応答をTrim・空行除去・先頭の番号/箇条書き記号の除去のうえ、
'   先頭maxN件だけ"|"区切りへ畳む(modSeed.SeedQuestionsと同じ形式)。
' R14-G5: 質問文中の "|" は区切りと衝突して1問が2つのボタンへ割れるので
'   全角 "／" へ置換する(消さずに元が列挙だったことを残す)。
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
