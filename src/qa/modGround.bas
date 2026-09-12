Attribute VB_Name = "modGround"
Option Explicit

' ============================================================================
' modGround - 回答本文に出てくる【数字が資料に実在するか】を確かめる(R46)
' ----------------------------------------------------------------------------
' なぜ要るのか
' ----------------------------------------------------------------------------
' R46 で作り直した信頼度バッジ(modMode.ConfidenceOf)が測っているのは
' 「検索が正しい材料を持ってこられたか」までで、【AIがその材料を正しく
' 読んで書いたか】は一度も見ていない。金額・期限・条文番号のような
' 「間違うと実害が出る値」を AI が補ってしまっても、バッジは緑のままになる。
'
' 損保の実務では、営業がこの数字をそのまま客へ伝える。1桁違えば事故になる。
' ここは「回答が上手いか」ではなく【誠実か】の問題なので、点数ではなく
' 事実で守る ―― 回答に出てきた数字を1つずつ、渡した資料の中に実在するか
' 文字列として照合する。
'
' 何を「主張を担う数字」とみなすか(誤検知を避けるための線引き)
' ----------------------------------------------------------------------------
'   ・第N条 / 第N項 / 別表N / 様式N      … 間違うと参照先ごと変わる
'   ・桁区切りのある数(1,000,000)         … 金額はほぼこの形
'   ・単位つきの数(50万円 / 3% / 30日 …) … 実務で効く値はだいたい単位を伴う
'   ・3桁以上の裸の数字                   … 証券番号・コード・年
' 逆に【数えないもの】:
'   ・自分で付けたタグの中の数字(出典タグ [本棚:… p.12] / セル番地 [B42] /
'     (B42 付近) / エラーコード (コード: E0203))。自分が書いた数字を
'     「AIが補った」と言うのは嘘になる。
'   ・行頭の箇条書き番号(1. / (2) / 3、)。文章の飾りであって主張ではない。
'   ・1〜2桁の裸の数字。「2つの方法」「3点」のような数え上げで頻出し、
'     拾うと警告が日常化して誰も読まなくなる(=検知しないのと同じ)。
'
' 照合の相手には【質問文も含める】。利用者が「500万円の車両保険は」と
' 聞いた回に、回答が 500万円 と書き返すのは引き写しであって捏造ではない。
'
' 実測(2026-09-11・教えてBOX 10,996かけら・本番と同じ条件)
' ----------------------------------------------------------------------------
' 質問250件について「検索の上位5件だけを根拠に、その質問の正解回答を照合」:
'   ・「資料に無い数字がある」と出た回答 13件(5.2%)
'     ―― 中身を見ると 080101 / 第8条 / 証券番号など、【上位5件に本当に
'        入っていない値】だった。本番では AI もその5件しか見ないので、
'        書いたとすれば補ったことになる = 指摘として正しい。
'   ・回答へ実在しない金額(7,654,321円)を1つ混ぜると 250/250 (100%) 検出。
'
' バッジは下げない(設計判断)
' ----------------------------------------------------------------------------
' 「検索が材料を取れたか」と「AIが材料どおり書いたか」は別の事実なので、
' 1つの点数に混ぜない。混ぜると、R44 で潰した「性質の違う信号を足して
' 絶対値のしきい値と比べる」のと同じ構造の誤りになる。
' ここは注記として別に出し、バッジは modMode.ConfidenceOf のままにする。
'
' 2026-09-12(R47) この配線で一度やらかした記録
' ----------------------------------------------------------------------------
' R46 で modAsk へ「nHits=0(本棚から1件も引けなかった回)にも常設ガードを付ける」
' という Else を足したが、【二重に誤り】だった。
'   (1) 到達不能だった。ok = True が立つのは modAsk.bas の2箇所だけで、どちらも
'       nHits>0 の枝の内側。nHits=0 では ok は False のままなので、
'       If ok Then の内側へ置いた Else は一度も実行されない死んだ分岐だった。
'   (2) 前提も違った。nHits=0 の枝は LLM を一度も呼ばず、
'       modClarify.MissingDocGuide の【自前の定型文】を返す。AI の回答でない
'       ものに「※ AIが資料から作った回答です」を付けるのはそれ自体が嘘になる。
' 「実装した本人がレビューでも気付けなかった」型なので、再発防止は人の注意では
' なく機械で行う ―― vba_lint の孤児Public検査(呼び出し0件をERROR)を R47 で
' 追加した。GuardOnly は一般アシスタント経路(modAppState)が唯一の呼び手。
'
' R4準拠(純ロジック): UngroundedNumbers / GroundNoteText は Excel オブジェクトに
'   触れないので LibreOffice でテストできる。Hit 配列を受ける AppendGroundNote
'   だけが配線用(テストからは呼ばない。modPeek の ShowPeek と同じ扱い)。
' ============================================================================

' 注記に並べる数字の上限。多すぎると読まれない(4件+「ほかN件」で足りる)。
Private Const GR_MAX_SHOW As Long = 4

' 単位つきの数として拾う単位。長いものを先に並べる(前方一致で先に当てるため)。
Private Const GR_UNITS As String = _
    "ヶ月|か月|カ月|箇月|万円|億円|千円|パーセント|％|%|割|円|日|年|歳|件|名|回|倍|点|口"

' 裸の数字をこの桁数以上なら主張とみなす。
Private Const GR_BARE_DIGITS As Long = 3

' ----------------------------------------------------------------------------
' UngroundedNumbers - 回答に出てくるが資料にも質問にも見つからない数字を
'   "|" 区切りで返す(純関数)。見つからなければ空文字。
'   maxN は返す最大件数(0以下なら GR_MAX_SHOW)。
' ----------------------------------------------------------------------------
Public Function UngroundedNumbers(ByVal ans As String, ByVal bodyText As String, _
                                  Optional ByVal maxN As Long = 0) As String
    If LenB(ans) = 0 Then Exit Function
    If maxN < 1 Then maxN = GR_MAX_SHOW

    Dim a As String: a = StripOwnTags(ans)
    a = modSparse.NormalizeForSearch(a)
    Dim body As String: body = modSparse.NormalizeForSearch(bodyText)
    If LenB(body) = 0 Then Exit Function

    Dim outS As String
    Dim cnt As Long
    Dim n As Long: n = Len(a)
    Dim i As Long: i = 1
    Do While i <= n
        Dim tok As String: tok = ClaimAt(a, i)
        If LenB(tok) = 0 Then
            i = i + 1
        Else
            ' 行頭の箇条書き番号は主張ではない
            If Not IsListHead(a, i, tok) Then
                If Not Grounded(tok, body) Then
                    If InStr(1, "|" & outS, "|" & tok & "|", vbBinaryCompare) = 0 Then
                        If cnt < maxN Then
                            outS = outS & tok & "|"
                            cnt = cnt + 1
                        End If
                    End If
                End If
            End If
            i = i + Len(tok)
        End If
    Loop

    If LenB(outS) > 0 Then UngroundedNumbers = Left$(outS, Len(outS) - 1)
End Function

' ----------------------------------------------------------------------------
' GroundNoteText - 注記の文面(純関数)。listPiped が空なら空文字。
'   totalN は実際に見つかった総数(上限で切った件数を「ほかN件」で出すため。
'   0 なら listPiped の件数と同じとみなす)。
'   文面は【断定しない】。書き写しの誤りかもしれないし、AI が補ったのかも
'   しれない。どちらであっても利用者がすべきことは同じ=出典で確かめる。
' ----------------------------------------------------------------------------
Public Function GroundNoteText(ByVal listPiped As String, _
                               Optional ByVal totalN As Long = 0) As String
    If LenB(listPiped) = 0 Then Exit Function
    Dim arr() As String: arr = Split(listPiped, "|")
    Dim shown As Long: shown = UBound(arr) - LBound(arr) + 1
    If totalN < shown Then totalN = shown

    Dim s As String
    s = ChrW(&H26A0) & ChrW(&HFE0F) & " 次の数字は出典の中に見つかりませんでした: "
    s = s & Replace(listPiped, "|", " / ")
    If totalN > shown Then s = s & " ほか" & (totalN - shown) & "件"
    s = s & vbLf & "書き写しの誤りか、AIが補った可能性があります。" & _
        "使う前に必ず出典で確かめてください。"
    GroundNoteText = s
End Function

' ----------------------------------------------------------------------------
' AppendGroundNote - 回答の末尾へ注記を足す配線(Hit配列を受けるので非純)。
'   modAsk から1行で呼ぶ。config ground_check(既定TRUE)で止められる。
'   注記を【常設ガードより前】に置くのは、ガードが最終段落である裁定
'   (modMode.DisplayNotes / R44)を崩さないため。呼び順は modAsk 側で
'   本関数 → ApplyLowHitWarning の順にしてある。
' ----------------------------------------------------------------------------
Public Function AppendGroundNote(ByVal ans As String, ByRef hits() As Hit, _
                                 ByVal n As Long, ByVal q As String) As String
    AppendGroundNote = ans
    If LenB(ans) = 0 Then Exit Function
    If n < 1 Then Exit Function

    On Error Resume Next
    If Not modConfig.GetBool("ground_check", True) Then Exit Function
    On Error GoTo 0

    Dim body As String
    Dim i As Long
    On Error Resume Next
    For i = 1 To n
        body = body & hits(i).full_text & vbLf
    Next i
    body = body & q
    On Error GoTo 0
    If LenB(body) = 0 Then Exit Function

    ' 総数は上限を外して数え、表示だけ上限で切る(「ほかN件」を正しく出す)。
    Dim allS As String: allS = UngroundedNumbers(ans, body, 99)
    If LenB(allS) = 0 Then Exit Function
    Dim total As Long: total = UBound(Split(allS, "|")) + 1

    Dim shown As String: shown = UngroundedNumbers(ans, body, GR_MAX_SHOW)
    If LenB(shown) = 0 Then Exit Function
    AppendGroundNote = ans & vbLf & vbLf & GroundNoteText(shown, total)
End Function

' ============================================================================
' 内部(純ロジック)
' ============================================================================

' StripOwnTags - 自分で付けたタグ・番地・コードを落とす。自分が書いた数字を
'   「AIが補った」と言わないため。落とすのは次の4種:
'     [本棚:… p.12] / [B42] / (B42 付近) / (コード: E0203)
Private Function StripOwnTags(ByVal s As String) As String
    Dim t As String: t = s
    t = DropBetween(t, "[本棚:", "]")
    t = DropBracketAddr(t)
    t = DropBetween(t, "(コード:", ")")
    t = DropBetween(t, "(コード：", ")")
    StripOwnTags = t
End Function

' DropBetween - head で始まり tail で終わる範囲を空白へ置き換える(全出現)。
'   tail が見つからない場合はそこで打ち切る(壊れた入力で無限ループしない)。
Private Function DropBetween(ByVal s As String, ByVal head As String, _
                             ByVal tail As String) As String
    Dim t As String: t = s
    Dim p As Long: p = InStr(1, t, head, vbBinaryCompare)
    Do While p > 0
        Dim q As Long: q = InStr(p + Len(head), t, tail, vbBinaryCompare)
        If q = 0 Then Exit Do
        t = Left$(t, p - 1) & " " & Mid$(t, q + Len(tail))
        p = InStr(1, t, head, vbBinaryCompare)
    Loop
    DropBetween = t
End Function

' DropBracketAddr - [A1] [AB123] のようなセル番地と、(B42 付近) を落とす。
'   英字1-3字+数字1-7桁だけを対象にする([本棚:…] は上で先に落ちている)。
Private Function DropBracketAddr(ByVal s As String) As String
    Dim t As String: t = s
    Dim outS As String
    Dim i As Long: i = 1
    Dim n As Long: n = Len(t)
    Do While i <= n
        Dim ch As String: ch = Mid$(t, i, 1)
        Dim consumed As Long: consumed = 0
        If ch = "[" Or ch = "(" Then
            consumed = AddrRunLen(t, i, IIf(ch = "[", "]", ")"))
        End If
        If consumed > 0 Then
            outS = outS & " "
            i = i + consumed
        Else
            outS = outS & ch
            i = i + 1
        End If
    Loop
    DropBracketAddr = outS
End Function

' AddrRunLen - 位置 i の開き括弧から「英字1-3+数字1-7(+ 付近)」+閉じ括弧まで
'   の長さ。当てはまらなければ0。
Private Function AddrRunLen(ByVal s As String, ByVal i As Long, _
                            ByVal closeCh As String) As Long
    Dim n As Long: n = Len(s)
    Dim p As Long: p = i + 1
    Dim alpha As Long
    Do While p <= n
        If Not IsAsciiAlpha(Mid$(s, p, 1)) Then Exit Do
        alpha = alpha + 1
        p = p + 1
    Loop
    If alpha < 1 Or alpha > 3 Then Exit Function
    Dim dig As Long
    Do While p <= n
        If Not IsAsciiDigit(Mid$(s, p, 1)) Then Exit Do
        dig = dig + 1
        p = p + 1
    Loop
    If dig < 1 Or dig > 7 Then Exit Function
    ' 「 付近」を挟む形も許す
    If Mid$(s, p, 3) = " 付近" Then p = p + 3
    If Mid$(s, p, 1) <> closeCh Then Exit Function
    AddrRunLen = p - i + 1
End Function

' ClaimAt - 位置 i から始まる「主張を担う数字」を返す。無ければ空文字。
Private Function ClaimAt(ByVal s As String, ByVal i As Long) As String
    Dim n As Long: n = Len(s)
    Dim ch As String: ch = Mid$(s, i, 1)

    ' 第N条 / 第N項 / 別表N / 様式N
    Dim lab As String: lab = LabeledAt(s, i)
    If LenB(lab) > 0 Then ClaimAt = lab: Exit Function

    If Not IsAsciiDigit(ch) Then Exit Function
    ' 直前が数字なら途中から拾わない
    If i > 1 Then
        If IsAsciiDigit(Mid$(s, i - 1, 1)) Then Exit Function
    End If

    ' 数字ラン(桁区切りのカンマと小数点を含めて読む)
    Dim p As Long: p = i
    Dim dig As Long
    Dim hasComma As Boolean
    Do While p <= n
        Dim c As String: c = Mid$(s, p, 1)
        If IsAsciiDigit(c) Then
            dig = dig + 1
            p = p + 1
        ElseIf c = "," Then
            If p + 1 <= n Then
                If IsAsciiDigit(Mid$(s, p + 1, 1)) Then
                    hasComma = True
                    p = p + 1
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        ElseIf c = "." Then
            If p + 1 <= n Then
                If IsAsciiDigit(Mid$(s, p + 1, 1)) Then
                    p = p + 1
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    Dim numLen As Long: numLen = p - i
    If numLen < 1 Then Exit Function

    ' 単位が続くか
    Dim u As String: u = UnitAt(s, p)
    If LenB(u) > 0 Then
        ClaimAt = Mid$(s, i, numLen) & u
        Exit Function
    End If
    If hasComma Then ClaimAt = Mid$(s, i, numLen): Exit Function
    If dig >= GR_BARE_DIGITS Then ClaimAt = Mid$(s, i, numLen)
End Function

' LabeledAt - 「第N条」「第N項」「別表N」「様式N」を返す。無ければ空文字。
Private Function LabeledAt(ByVal s As String, ByVal i As Long) As String
    Dim n As Long: n = Len(s)
    Dim head As String: head = Mid$(s, i, 1)
    Dim p As Long
    If head = "第" Then
        p = i + 1
        Dim d1 As Long
        Do While p <= n
            If Not IsAsciiDigit(Mid$(s, p, 1)) Then Exit Do
            d1 = d1 + 1
            p = p + 1
        Loop
        If d1 < 1 Then Exit Function
        Dim unit As String: unit = Mid$(s, p, 1)
        If unit = "条" Or unit = "項" Or unit = "章" Or unit = "号" Then
            LabeledAt = Mid$(s, i, p - i + 1)
        End If
        Exit Function
    End If
    If Mid$(s, i, 2) = "別表" Or Mid$(s, i, 2) = "様式" Then
        p = i + 2
        If Mid$(s, p, 1) = "第" Then p = p + 1
        Dim d2 As Long
        Do While p <= n
            If Not IsAsciiDigit(Mid$(s, p, 1)) Then Exit Do
            d2 = d2 + 1
            p = p + 1
        Loop
        If d2 >= 1 Then LabeledAt = Mid$(s, i, p - i)
    End If
End Function

' UnitAt - 位置 p から始まる単位を返す(GR_UNITS の並び順＝長いものが先)。
Private Function UnitAt(ByVal s As String, ByVal p As Long) As String
    If p > Len(s) Then Exit Function
    Dim arr() As String: arr = Split(GR_UNITS, "|")
    Dim k As Long
    For k = LBound(arr) To UBound(arr)
        If Len(arr(k)) > 0 Then
            If Mid$(s, p, Len(arr(k))) = arr(k) Then
                UnitAt = arr(k)
                Exit Function
            End If
        End If
    Next k
End Function

' IsListHead - その数字が行頭の箇条書き番号か。「1. 」「(2)」「3、」の形で、
'   かつ2桁以下のときだけ真(桁の多い数は番号ではない)。
Private Function IsListHead(ByVal s As String, ByVal i As Long, _
                            ByVal tok As String) As Boolean
    If Len(tok) > 2 Then Exit Function
    ' 行頭か(直前が改行か文頭。空白と開き括弧は跨いでよい)
    Dim p As Long: p = i - 1
    Do While p >= 1
        Dim c As String: c = Mid$(s, p, 1)
        If c = " " Or c = "(" Or c = "（" Then
            p = p - 1
        Else
            Exit Do
        End If
    Loop
    If p >= 1 Then
        Dim b As String: b = Mid$(s, p, 1)
        If b <> vbLf And b <> vbCr Then Exit Function
    End If
    ' 直後が区切り記号か
    Dim q As String: q = Mid$(s, i + Len(tok), 1)
    IsListHead = (q = "." Or q = "．" Or q = ")" Or q = "）" Or q = "、" Or q = "，")
End Function

' Grounded - その数字が本文に実在するか。表記ゆれ(単位の有無)を許すため、
'   丸ごと一致しなければ数字部分だけでも照合する。
Private Function Grounded(ByVal tok As String, ByVal body As String) As Boolean
    If InStr(1, body, tok, vbBinaryCompare) > 0 Then Grounded = True: Exit Function
    Dim d As String: d = DigitsOnly(tok)
    If Len(d) >= 2 Then
        If InStr(1, body, d, vbBinaryCompare) > 0 Then Grounded = True
    End If
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim outS As String
    Dim i As Long
    For i = 1 To Len(s)
        Dim c As String: c = Mid$(s, i, 1)
        If IsAsciiDigit(c) Or c = "," Or c = "." Then outS = outS & c
    Next i
    DigitsOnly = outS
End Function

Private Function IsAsciiDigit(ByVal c As String) As Boolean
    If LenB(c) = 0 Then Exit Function
    IsAsciiDigit = (c >= "0" And c <= "9")
End Function

Private Function IsAsciiAlpha(ByVal c As String) As Boolean
    If LenB(c) = 0 Then Exit Function
    If c >= "a" And c <= "z" Then IsAsciiAlpha = True: Exit Function
    If c >= "A" And c <= "Z" Then IsAsciiAlpha = True
End Function
