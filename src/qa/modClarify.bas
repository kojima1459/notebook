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
Private Const K_PENDING_SRC As String = "clarify_src"    ' 候補資料名 or 読み方の選択肢(|区切り)
' 保留の種類(2026-08-05 R16-3B)。"source"=資料+意図の聞き返し(従来)、
' "topic"=質問の読み方を番号で選ばせる逆質問。どちらも同じ3キー
' (元質問・候補リスト・作成時刻)に載せ、後から作った方が前を上書きする
' (2つの聞き返しが同時に生きていると、返ってきた番号がどちらの表の番号か
'  決められない。保留は常に1つ、が唯一破綻しない持ち方)。
Private Const K_PENDING_KIND As String = "clarify_kind"
Private Const KIND_SOURCE As String = "source"
Private Const KIND_TOPIC As String = "topic"
' 番号だけの返答とみなす最大文字数(これを超えたら文章とみなす)。
Private Const MAX_CHOICE_CHARS As Long = 6
Private Const K_PENDING_AT As String = "clarify_at"      ' 保留を作った時刻(有効期限判定用)

' 保留の有効期限(分)。これを過ぎた聞き返しは無かったことにする。
' 2026-07-28(レビュー H-12): 保留はチャットをクリアするまでブックの再起動を
' またいで残るため、期限が無いと「翌日のまったく無関係な質問」が
' 前日の聞き返しに吸収されるという乗っ取りが起きる。
Private Const PENDING_TTL_MIN As Long = 30
Private Const MAX_SRC As Long = 4
Private Const CJK_LO As Long = 19968       ' &H4E00 CJK統合漢字 開始
Private Const CJK_HI As Long = 40959       ' &H9FFF CJK統合漢字 終端
Private Const KATA_LO As Long = 12449      ' &H30A1 ァ
Private Const KATA_HI As Long = 12538      ' &H30FA カタカナ末尾
Private Const KATA_PROLONG As Long = 12540 ' &H30FC ー


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
    SavePending q, sources, KIND_SOURCE

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
' SaveTopicPending - 「読み方を番号で選んでください」の保留を作る(R16-3B)。
'   options: 選択肢の文(| 区切り。modAskMulti が段0の <options> から作る)。
'   逆質問の本文そのものは modAskMulti が組み立てる(あちらが提示の主体で、
'   ここは保留状態の置き場を1つに保つ役割だけを持つ)。
' ----------------------------------------------------------------------------
Public Sub SaveTopicPending(ByVal q As String, ByVal optList As String)
    SavePending q, optList, KIND_TOPIC
End Sub

' 保留3キー+種類の保存。source/topic のどちらも必ずここを通す(片方だけ
' 種類を書き忘れると、前の種類が残って返事の解釈が入れ替わる)。
Private Sub SavePending(ByVal q As String, ByVal listLine As String, ByVal kind As String)
    On Error Resume Next
    modState.SaveState K_PENDING_Q, q
    modState.SaveState K_PENDING_SRC, listLine
    modState.SaveState K_PENDING_KIND, kind
    modState.SaveState K_PENDING_AT, modUtilText.IsoDateTime(Now)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' HasPending - 聞き返しの途中かどうか。
' ----------------------------------------------------------------------------
Public Function HasPending() As Boolean
    On Error Resume Next
    If LenB(modState.LoadState(K_PENDING_Q, "")) = 0 Then Exit Function
    ' 期限切れの保留は「無い」とみなし、その場で片付ける(レビュー H-12)。
    If PendingExpired() Then
        ClearPending
        Exit Function
    End If
    HasPending = True
    On Error GoTo 0
End Function

' 保留を作ってから PENDING_TTL_MIN 分を過ぎたか。時刻が読めないときは
' 期限切れ扱いにする(判断できない保留を残す方が危ない)。
Private Function PendingExpired() As Boolean
    On Error Resume Next
    PendingExpired = IsPendingExpired(modState.LoadState(K_PENDING_AT, ""), Now)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' IsPendingExpired - 保留の失効判定(純関数。savedIso は保存した時刻の文字列)。
'   空文字 = 旧データ(時刻なし)は従来どおり有効。読めない文字列は失効扱い
'   (判断できない保留を生かしておくと、翌日の無関係な質問が前日の聞き返しに
'    吸収される=レビュー H-12 の乗っ取り)。判定を純関数へ出したのは、
'   TTLの境界(30分ちょうどは有効・31分で失効)を実行テストで固定するため。
' ----------------------------------------------------------------------------
Public Function IsPendingExpired(ByVal savedIso As String, ByVal nowT As Date) As Boolean
    If LenB(Trim$(savedIso)) = 0 Then Exit Function
    On Error GoTo Expired
    Dim t0 As Date: t0 = CDate(savedIso)
    IsPendingExpired = (DateDiff("n", t0, nowT) > PENDING_TTL_MIN)
    Exit Function
Expired:
    IsPendingExpired = True
End Function

Public Sub ClearPending()
    On Error Resume Next
    modState.SaveState K_PENDING_Q, ""
    modState.SaveState K_PENDING_SRC, ""
    modState.SaveState K_PENDING_KIND, ""
    modState.SaveState K_PENDING_AT, ""
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' MergeAnswer - 聞き返しへの返事を、元の質問と合成して完全な質問文にする。
'   reply が番号選択でないとき(利用者が自分の言葉で書き直したとき)は、
'   その文をそのまま新しい質問として扱う(合成しない)。
'   戻り値 = 合成後の質問文。保留状態はここで必ず消す。
' ----------------------------------------------------------------------------
Public Function MergeAnswer(ByVal reply As String) As String
    Dim origQ As String, srcList As String, kind As String
    On Error Resume Next
    origQ = modState.LoadState(K_PENDING_Q, "")
    srcList = modState.LoadState(K_PENDING_SRC, "")
    kind = modState.LoadState(K_PENDING_KIND, "")
    On Error GoTo 0
    ClearPending

    Dim r As String: r = Trim$(reply)
    If LenB(origQ) = 0 Then
        MergeAnswer = r
        Exit Function
    End If

    ' R16-3B: 読み方の逆質問(topic)は番号の意味が違う。あちらの番号は
    ' 「調べる読み方」で複数選べるため、資料1つ+意図1つを組む下の合成には
    ' 乗らない(乗せると2つ選んだ人の2つ目が黙って消える)。
    If kind = KIND_TOPIC Then
        MergeAnswer = MergeTopicAnswer(origQ, srcList, r)
        Exit Function
    End If

    ' 2026-07-28(レビュー H-12): 番号選択とみなす条件を厳しくした。
    '
    ' 以前は「返答のどこかに半角1～4がある」だけで番号選択と判定し、
    ' 利用者が実際に打った文章を捨てて「前回の質問＋対象の資料: 候補N」を
    ' 送っていた。保険実務の質問は「第1条」「3日以内」「2026年度」のように
    ' 数字をほぼ確実に含むため、聞き返しのあと普通に質問を書き直した人は
    ' 自分の文章が消えたことに気付けない(何も言わずに別の質問が送られる)。
    '
    ' 番号選択は「番号だけを短く打った」ときに限る。それ以外は書き直しとして
    ' 打った文章をそのまま優先する。
    Dim srcPick As String, intentPick As String
    If IsNumberChoiceOnly(r) Then
        srcPick = PickSource(r, srcList)
        intentPick = PickIntent(r)
    End If

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

' ----------------------------------------------------------------------------
' MergeTopicAnswer - 読み方の逆質問(R16-3B)への返事を元質問と合成する(純関数)。
'   optList: 提示した選択肢(| 区切り)。reply: 利用者の返事。
'
'   番号が1つ以上読めたら「元質問 / 選んだ読み方 / …」を返す。複数選んだ場合は
'   そのまま並べる(次のターンの段0がこれを parts に割って論点ごとに調べる=
'   「1と3」が『AとBを両方調べる』として素直に成立する)。
'   番号が読めない返事は【書き直し】として扱い、打った文章を優先する。選択肢を
'   無視して質問を書き直す自由を必ず残す(番号を強制すると、聞き方が分からない
'   人ほど詰まる)。短すぎる返事だけは元質問へ足して捨てない。
' ----------------------------------------------------------------------------
Public Function MergeTopicAnswer(ByVal origQ As String, ByVal optList As String, _
                                 ByVal reply As String) As String
    Dim opts() As String
    Dim n As Long
    If LenB(optList) > 0 Then
        opts = Split(optList, "|")
        n = UBound(opts) - LBound(opts) + 1
    End If

    Dim picked As String
    If n > 0 Then picked = modRagParse.ParseChoiceNumbers(reply, n)

    If LenB(picked) = 0 Then
        If Len(Trim$(reply)) >= 8 Then
            MergeTopicAnswer = Trim$(reply)
        Else
            MergeTopicAnswer = Trim$(origQ & " " & reply)
        End If
        Exit Function
    End If

    Dim nums() As String: nums = Split(picked, ",")
    Dim sb As String: sb = origQ
    Dim i As Long
    For i = LBound(nums) To UBound(nums)
        Dim k As Long: k = CLng(Val(nums(i)))
        If k >= 1 And k <= n Then sb = sb & " / " & Trim$(opts(LBound(opts) + k - 1))
    Next i
    MergeTopicAnswer = sb
End Function

' 返事から資料番号(1..MAX_SRC)を読み取り、対応する資料名を返す。
' ----------------------------------------------------------------------------
' IsNumberChoiceOnly - 「番号だけを打った返答」かどうか(純関数)。
'
' True にする例: "2" / "２" / "1-3" / "① ③" / "3." / "(2)" / "2、4"
' False にする例: "第1条の適用範囲は?" / "3日以内に出す必要ある?" / "2026年度"
'
' 判定は「短い」かつ「数字とその周辺記号だけでできている」かつ
' 「数字を1つ以上含む」の3点。文字が1つでも混じったら書き直しとみなす。
' レビュー H-12 の恒久再発防止として、ここだけを見ればルールが分かるように
' 1つの純関数に閉じてある(テストは modTestsPure2 側)。
' ----------------------------------------------------------------------------
Public Function IsNumberChoiceOnly(ByVal s As String) As Boolean
    Dim r As String: r = Trim$(s)
    If LenB(r) = 0 Then Exit Function
    If Len(r) > MAX_CHOICE_CHARS Then Exit Function

    Dim hasDigit As Boolean
    Dim i As Long
    For i = 1 To Len(r)
        Dim cp As Long: cp = AscW(Mid$(r, i, 1))
        If cp < 0 Then cp = cp + 65536          ' AscWは符号付きで返る
        If IsChoiceDigit(cp) Then
            hasDigit = True
        ElseIf Not IsChoiceSeparator(cp) Then
            Exit Function                        ' 数字でも区切りでもない文字
        End If
    Next i
    IsNumberChoiceOnly = hasDigit
End Function

' 選択肢の番号として認める文字か。半角0-9 / 全角０-９ / 丸数字①-⑳。
Private Function IsChoiceDigit(ByVal cp As Long) As Boolean
    If cp >= 48 And cp <= 57 Then IsChoiceDigit = True: Exit Function        ' 0-9
    If cp >= 65296 And cp <= 65305 Then IsChoiceDigit = True: Exit Function  ' ０-９
    If cp >= 9312 And cp <= 9331 Then IsChoiceDigit = True                   ' ①-⑳
End Function

' 番号の周りに来てよい区切り文字か。ここに無い文字が1つでもあれば
' 「文章を書いた」とみなす。
Private Function IsChoiceSeparator(ByVal cp As Long) As Boolean
    Select Case cp
        Case 32, 9                      ' 半角スペース / タブ
            IsChoiceSeparator = True
        Case 12288                      ' 全角スペース
            IsChoiceSeparator = True
        Case 46, 44, 45, 40, 41         ' . , - ( )
            IsChoiceSeparator = True
        Case 65294, 65292, 65293, 65288, 65289   ' ．，－（）
            IsChoiceSeparator = True
        Case 12289, 12290               ' 、。
            IsChoiceSeparator = True
        Case 12540                      ' ー(長音。全角ハイフンの打ち間違い)
            IsChoiceSeparator = True
        Case 8722                       ' −(マイナス記号)
            IsChoiceSeparator = True
    End Select
End Function

Private Function PickSource(ByVal reply As String, ByVal srcList As String) As String
    If LenB(srcList) = 0 Then Exit Function
    Dim parts() As String
    parts = Split(srcList, "|")

    Dim i As Long
    For i = 0 To UBound(parts)
        Dim num As String: num = CStr(i + 1)
        ' 「2-②」「2.」「2 」等に耐える: 半角数字の出現位置だけを見る。
        If InStr(1, reply, num) > 0 Then
            ' ①～⑤に含まれる数字と誤認しないよう、半角数字のみを対象にする。
            PickSource = parts(i)
            Exit Function
        End If
    Next i
End Function

' 返事から意図番号(①～⑤ または 半角1～5の後置)を読み取る。
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

' 文字コードの定数は10進で書く。&H9FFF は 40959 だが、VBA/LO Basic は
' 0x8000 以上の &H リテラルを16bit整数として解釈するため -24577 になる。
' その結果 "c <= &H9FFF" は正の c に対して常に偽となり、
' 漢字が1文字も単語として認識されていなかった(2026-07-27発見)。
' カタカナ(&H30A1～&H30FA)は 0x8000 未満なので偶然動いていた。
' 影響: 質問からのキーワード抽出が漢字を拾えず、資料調達の案内が
' 的外れになっていた。損保の用語はほぼ漢字なので実質機能していない。

Private Function IsWordChar(ByVal ch As String) As Boolean
    If LenB(ch) = 0 Then Exit Function
    Dim c As Long: c = AscW(ch)
    If c < 0 Then c = c + 65536
    If c >= CJK_LO And c <= CJK_HI Then IsWordChar = True: Exit Function
    If c >= KATA_LO And c <= KATA_HI Then IsWordChar = True: Exit Function
    If c = KATA_PROLONG Then IsWordChar = True                 ' 長音符
End Function
