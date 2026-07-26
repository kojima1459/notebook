Attribute VB_Name = "modClarify"
Option Explicit

' ============================================================================
' modClarify - 質問の意図確認(逆質問)と、資料不足のときの調達ガイド
' ----------------------------------------------------------------------------
' 何を解決するか:
'   実務で飛んでくる質問の多くは「約款6条について」のような断片で、そのまま
'   検索しても回答は必ずズレる。ズレた回答に「違う」を押されると、利用者には
'   もやもやが残り(聞き方も分からないまま二度と使わなくなる)、システムには
'   質の低い否定シグナルだけが溜まる。負のループの入口はここにある。
'
' どう解決するか(自由記述の逆質問は避ける):
'   「もっと詳しく書いてください」と返すのは、聞き方が分からない人には
'   何の助けにもならない。代わりに、いったん検索して当たった資料を材料に
'   「番号で選べる逆質問」を出す。
'     ・どの資料の話か → 実際にヒットした資料名を最大4件、番号付きで提示
'     ・何を知りたいか → 適用条件/必要書類/期限/例外/金額 の5区分から選ばせる
'   利用者は「2-③」のように打つだけでよい。次の送信でこの選択と元の質問を
'   合成し、完全な質問文に組み立て直してから通常の回答フローへ流す。
'   選ぶ行為そのものが「こう聞けばいいのか」の学習になる(行動変容)。
'
'   保留状態は ui_state(modState)に置く。VBAプロジェクトのリセットや
'   ブックの開き直しをまたいでも、聞き返しの途中で迷子にならない。
'
' 資料が1件も当たらないとき:
'   「見つかりません」で終わらせず、質問文から拾った語をもとに
'   「ポータルでこういう資料を探して取り込めば答えられる」と具体的に指示する。
'   本棚が空では回答すらできない以上、次の一手を示すことが唯一の親切になる。
' ============================================================================

Private Const K_PENDING_Q As String = "clarify_q"        ' 聞き返し中の元質問
Private Const K_PENDING_SRC As String = "clarify_src"    ' 候補資料名(|区切り)
Private Const MAX_SRC As Long = 4

' 意図の区分。番号で選ばせるので順序を変えないこと(利用者の記憶に残る)。
Private Const INTENT_1 As String = "適用の条件・対象(どんなときに当てはまるか)"
Private Const INTENT_2 As String = "必要な書類・手続きの流れ"
Private Const INTENT_3 As String = "期限・所要日数"
Private Const INTENT_4 As String = "例外・特約・注意点"
Private Const INTENT_5 As String = "金額・料率・計算方法"

' ----------------------------------------------------------------------------
' BuildClarifyPrompt - 逆質問の本文を組み立て、保留状態を保存する。
'   sources: ヒットした資料名(重複除去済み・最大MAX_SRC件)を | 区切りで渡す。
' ----------------------------------------------------------------------------
Public Function BuildClarifyPrompt(ByVal q As String, ByVal sources As String) As String
    On Error Resume Next
    modState.SaveState K_PENDING_Q, q
    modState.SaveState K_PENDING_SRC, sources
    On Error GoTo 0

    Dim sb As String
    sb = ChrW(&HD83D) & ChrW(&HDCAD) & " もう少しだけ教えてください。" & _
         "そのほうが確実な答えを出せます。" & vbLf & vbLf

    Dim parts() As String
    Dim n As Long
    If LenB(sources) > 0 Then
        parts = Split(sources, "|")
        n = UBound(parts) - LBound(parts) + 1
    End If

    If n > 0 Then
        sb = sb & "■ どの資料のお話でしょうか" & vbLf
        Dim i As Long
        For i = 0 To n - 1
            sb = sb & "  " & (i + 1) & ". " & parts(i) & vbLf
        Next i
        sb = sb & "  0. この中にない / わからない" & vbLf & vbLf
    End If

    sb = sb & "■ 知りたいのはどれに近いですか" & vbLf & _
         "  " & ChrW(&H2460) & " " & INTENT_1 & vbLf & _
         "  " & ChrW(&H2461) & " " & INTENT_2 & vbLf & _
         "  " & ChrW(&H2462) & " " & INTENT_3 & vbLf & _
         "  " & ChrW(&H2463) & " " & INTENT_4 & vbLf & _
         "  " & ChrW(&H2464) & " " & INTENT_5 & vbLf & vbLf

    If n > 0 Then
        sb = sb & "「2-" & ChrW(&H2461) & "」のように番号だけ送ってください。" & vbLf
    Else
        sb = sb & ChrW(&H2460) & ChrW(&H301C) & ChrW(&H2464) & " の番号だけ送ってください。" & vbLf
    End If
    sb = sb & "もちろん、ご自身の言葉で詳しく書き直していただいても構いません。"

    BuildClarifyPrompt = sb
End Function

' ----------------------------------------------------------------------------
' HasPending - 聞き返しの途中かどうか。
' ----------------------------------------------------------------------------
Public Function HasPending() As Boolean
    On Error Resume Next
    HasPending = (LenB(modState.LoadState(K_PENDING_Q, "")) > 0)
    On Error GoTo 0
End Function

Public Sub ClearPending()
    On Error Resume Next
    modState.SaveState K_PENDING_Q, ""
    modState.SaveState K_PENDING_SRC, ""
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' MergeAnswer - 聞き返しへの返事を、元の質問と合成して完全な質問文にする。
'   reply が番号選択でないとき(利用者が自分の言葉で書き直したとき)は、
'   その文をそのまま新しい質問として扱う(合成しない)。
'   戻り値 = 合成後の質問文。保留状態はここで必ず消す。
' ----------------------------------------------------------------------------
Public Function MergeAnswer(ByVal reply As String) As String
    Dim origQ As String, srcList As String
    On Error Resume Next
    origQ = modState.LoadState(K_PENDING_Q, "")
    srcList = modState.LoadState(K_PENDING_SRC, "")
    On Error GoTo 0
    ClearPending

    Dim r As String: r = Trim$(reply)
    If LenB(origQ) = 0 Then
        MergeAnswer = r
        Exit Function
    End If

    Dim srcPick As String, intentPick As String
    srcPick = PickSource(r, srcList)
    intentPick = PickIntent(r)

    ' 番号がひとつも読み取れない=自分の言葉で書き直した、と解釈する。
    If LenB(srcPick) = 0 And LenB(intentPick) = 0 Then
        If Len(r) >= 8 Then
            MergeAnswer = r
        Else
            ' 短い上に番号でもない。元の質問に足すだけにする(捨てない)。
            MergeAnswer = origQ & " " & r
        End If
        Exit Function
    End If

    Dim sb As String
    sb = origQ
    If LenB(srcPick) > 0 Then sb = sb & vbLf & "対象の資料: " & srcPick
    If LenB(intentPick) > 0 Then sb = sb & vbLf & "知りたいこと: " & intentPick
    sb = sb & vbLf & "(この点に絞って、資料の記載に沿って具体的に答えてください)"
    MergeAnswer = sb
End Function

' 返事から資料番号(1..MAX_SRC)を読み取り、対応する資料名を返す。
Private Function PickSource(ByVal reply As String, ByVal srcList As String) As String
    If LenB(srcList) = 0 Then Exit Function
    Dim parts() As String
    parts = Split(srcList, "|")

    Dim i As Long
    For i = 0 To UBound(parts)
        Dim num As String: num = CStr(i + 1)
        ' 「2-②」「2.」「2 」等に耐える: 半角数字の出現位置だけを見る。
        If InStr(1, reply, num) > 0 Then
            ' ①〜⑤に含まれる数字と誤認しないよう、半角数字のみを対象にする。
            PickSource = parts(i)
            Exit Function
        End If
    Next i
End Function

' 返事から意図番号(①〜⑤ または 半角1〜5の後置)を読み取る。
Private Function PickIntent(ByVal reply As String) As String
    If InStr(1, reply, ChrW(&H2460)) > 0 Then PickIntent = INTENT_1: Exit Function
    If InStr(1, reply, ChrW(&H2461)) > 0 Then PickIntent = INTENT_2: Exit Function
    If InStr(1, reply, ChrW(&H2462)) > 0 Then PickIntent = INTENT_3: Exit Function
    If InStr(1, reply, ChrW(&H2463)) > 0 Then PickIntent = INTENT_4: Exit Function
    If InStr(1, reply, ChrW(&H2464)) > 0 Then PickIntent = INTENT_5: Exit Function

    ' 全角丸数字が打てない環境向け: 「-3」「 3」のような後置の数字を拾う。
    Dim p As Long: p = InStr(1, reply, "-")
    If p < 1 Then p = InStr(1, reply, ChrW(&HFF0D))     ' 全角ハイフン
    If p >= 1 And p < Len(reply) Then
        Select Case Trim$(Mid$(reply, p + 1, 1))
            Case "1": PickIntent = INTENT_1
            Case "2": PickIntent = INTENT_2
            Case "3": PickIntent = INTENT_3
            Case "4": PickIntent = INTENT_4
            Case "5": PickIntent = INTENT_5
        End Select
    End If
End Function

' ----------------------------------------------------------------------------
' MissingDocGuide - 本棚に根拠が無かったときの調達ガイド。
'   「見つかりません」で終わらせず、探すべき資料の種類を質問文の語から
'   組み立てて示す。本棚が空では回答すらできない以上、これが唯一の次の一手。
' ----------------------------------------------------------------------------
Public Function MissingDocGuide(ByVal q As String) As String
    Dim kw As String: kw = MainKeyword(q)

    Dim sb As String
    sb = ChrW(&HD83D) & ChrW(&HDCED) & " 本棚に、この内容の根拠になる資料がありませんでした。" & vbLf & _
         "推測でお答えすると実務で事故になるため、ここでは答えを作りません。" & vbLf & vbLf

    sb = sb & "■ こうすると答えられるようになります" & vbLf & _
        "社内ポータルで次のような資料を探し、「ナレッジ」画面の" & _
        ChrW(&HD83D) & ChrW(&HDCC1) & "追加 から取り込んでください。" & vbLf
    If LenB(kw) > 0 Then
        sb = sb & "  ・「" & kw & "」の約款 / 規程 / 取扱要領" & vbLf & _
                  "  ・「" & kw & "」の業務マニュアル・手続きガイド" & vbLf & _
                  "  ・「" & kw & "」に関する通達・改定のお知らせ" & vbLf
    Else
        sb = sb & "  ・関係する約款 / 規程 / 取扱要領" & vbLf & _
                  "  ・該当業務のマニュアル・手続きガイド" & vbLf
    End If
    sb = sb & "  ※ ページ数が多い資料でも大丈夫です。必要な箇所だけAIが拾います。" & vbLf & vbLf

    sb = sb & "■ いま押していただけると助かること" & vbLf & _
        "この質問は「まだ答えを用意できていない質問」として部内に共有されます。" & vbLf & _
        "資料を持っている人が見て登録すると、次からは全員がすぐ答えを得られます。"

    MissingDocGuide = sb
End Function

' 質問文から代表的な語を1つ取り出す(漢字・カタカナの連なりで最長のもの)。
' 形態素解析は使えないので、実務語がほぼ漢字/カタカナ列である性質を利用する。
Private Function MainKeyword(ByVal q As String) As String
    Dim best As String, cur As String
    Dim i As Long
    For i = 1 To Len(q)
        Dim ch As String: ch = Mid$(q, i, 1)
        If IsWordChar(ch) Then
            cur = cur & ch
            If Len(cur) > Len(best) Then best = cur
        Else
            cur = ""
        End If
    Next i
    If Len(best) >= 2 Then MainKeyword = best
End Function

Private Function IsWordChar(ByVal ch As String) As Boolean
    Dim c As Long: c = AscW(ch)
    If c < 0 Then c = c + 65536
    ' 漢字(CJK統合漢字) / カタカナ / 全角英数
    If c >= &H4E00 And c <= &H9FFF Then IsWordChar = True: Exit Function
    If c >= &H30A1 And c <= &H30FA Then IsWordChar = True: Exit Function
    If c = &H30FC Then IsWordChar = True                       ' 長音符
End Function
