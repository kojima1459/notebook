Attribute VB_Name = "modSparse"
Option Explicit

' ============================================================================
' modSparse - 日本語のキーワード検索(文字bigram + BM25 + 完全一致)
' ----------------------------------------------------------------------------
' なぜ作り直したか(実測に基づく)
' ----------------------------------------------------------------------------
' 実物6資料180チャンク・31問で採点した結果(tools/bench_retrieval.py):
'
'     手法                          R@1    R@3    R@5    MRR
'     現行(空白分割+一律加点)         32%    32%    32%   0.34
'     文字bigram BM25                84%    97%   100%   0.91
'
' 旧実装は質問を空白で割って部分一致を見ていた。日本語は空白で区切らないので、
' 実質「質問文まるごと1語」を本文から探すことになり、まず当たらない。
' しかも当たっても一律 +0.05(上限0.15)なので、「保険」のような頻出語と
' 「瑕疵」のような希少語が同じ重みになる。3問に2問は正解資料が1位に来ない、
' つまり赤点だった。
'
' 設計
' ----------------------------------------------------------------------------
' 1. 正規化(NormalizeForSearch)
'    全角数字→半角、英大文字→小文字、全角空白→半角。
'    利用者は「第20条」「第２０条」「Recall」「RECALL」を区別せずに打つ。
'    ここで揃えないと、揺れた瞬間に0件になる。金融では最悪の事故。
'
' 2. 文字bigram(AppendBigrams)
'    形態素解析器が使えない環境での定石。
'    「保険金を支払わない」→ 保険/険金/金を/を支/支払/払わ/わな/ない
'    語の切れ目を知らなくても部分一致が確実に効き、表記の揺れにも強い。
'    英数字は語のまま(型番・コードは分割しないほうが効く)。
'
' 3. BM25(Bm25Score)
'    ・IDFで希少語を重くする(「保険」より「瑕疵」を効かせる)
'    ・文書長で正規化する(長いチャンクが不当に有利になるのを防ぐ)
'    ・上限で頭打ちにしない(強く一致したものが確実に上へ来る)
'
' 4. 完全一致の決定的な加点(ExactHitCount)
'    「第12条は?」と聞かれて第12条が1位に来ないのは、金融では事故。
'    ベクトルもBM25も確率的なので、条番号・型番だけは決定的に押し上げる。
'
' 速度について(VBAの現実)
' ----------------------------------------------------------------------------
' 2万チャンク全件に bigram BM25 を回すと、1チャンク900字×9000件で
' 800万回の文字走査になり数十秒かかる。実務では使えない。
' そこで modRetrieve は「候補生成 → 再ランク」にする:
'     ① dense(埋め込み)の上位K件            … 既存の内積計算。意味の近さ
'     ② 質問から抜いた"効く語"の全件部分一致 … InStrはネイティブで速い
'     ③ ①∪② の小さなプールにだけ BM25    … ここが本モジュール
' ②を入れるのは、denseが取りこぼした完全一致を救うため。
' 実測では②+③だけ(=denseが全く役に立たなかった最悪ケース)でも
' R@1 84% / R@5 100% を維持した。
'
' R4準拠: Excelオブジェクトに触れないので LibreOffice でテストできる。
' 回帰は modTestsPure2 が固定する。
' ============================================================================

Private Const FW_DIGIT_LO As Long = 65296   ' ０
Private Const FW_DIGIT_HI As Long = 65305   ' ９
Private Const FW_UPPER_LO As Long = 65313   ' Ａ
Private Const FW_UPPER_HI As Long = 65338   ' Ｚ
Private Const FW_LOWER_LO As Long = 65345   ' ａ
Private Const FW_LOWER_HI As Long = 65370   ' ｚ
Private Const IDEO_SPACE As Long = 12288    ' 全角スペース

Private Const BM25_K1 As Double = 1.2
Private Const BM25_B As Double = 0.75
Private Const MAX_KEYS As Long = 8

' ----------------------------------------------------------------------------
' NormalizeForSearch - 検索用の正規化。取込側・質問側の両方が必ず通す。
'   片方だけ通すと揺れが残るので、必ず両側で同じ関数を使うこと。
' ----------------------------------------------------------------------------
Public Function NormalizeForSearch(ByVal s As String) As String
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Function

    Dim buf() As String: ReDim buf(1 To n)
    Dim i As Long
    For i = 1 To n
        Dim ch As String: ch = Mid$(s, i, 1)
        Dim c As Long: c = AscW(ch)
        If c < 0 Then c = c + 65536

        If c >= FW_DIGIT_LO And c <= FW_DIGIT_HI Then
            buf(i) = Chr$(48 + (c - FW_DIGIT_LO))            ' ０-９ → 0-9
        ElseIf c >= FW_UPPER_LO And c <= FW_UPPER_HI Then
            buf(i) = Chr$(97 + (c - FW_UPPER_LO))            ' Ａ-Ｚ → a-z
        ElseIf c >= FW_LOWER_LO And c <= FW_LOWER_HI Then
            buf(i) = Chr$(97 + (c - FW_LOWER_LO))            ' ａ-ｚ → a-z
        ElseIf c >= 65 And c <= 90 Then
            buf(i) = Chr$(c + 32)                            ' A-Z → a-z
        ElseIf c = IDEO_SPACE Or c = 9 Then
            buf(i) = " "
        Else
            buf(i) = ch
        End If
    Next i
    NormalizeForSearch = Join(buf, "")
End Function

' ----------------------------------------------------------------------------
' Tokenize - 正規化済み文字列をトークン列(| 区切り)にする。
'   ・英数字は語のまま(型番・コードを割らない)
'   ・かな漢字は文字bigram
' ----------------------------------------------------------------------------
Public Function Tokenize(ByVal rawText As String) As String
    Dim s As String: s = NormalizeForSearch(rawText)
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Function

    Dim toks As String
    Dim jp As String
    Dim alnum As String

    Dim i As Long
    For i = 1 To n
        Dim ch As String: ch = Mid$(s, i, 1)
        If IsAlnumChar(ch) Then
            alnum = alnum & ch
        Else
            If LenB(alnum) > 0 Then
                toks = toks & alnum & "|"
                alnum = ""
            End If
            If IsJapaneseChar(ch) Then jp = jp & ch
        End If
    Next i
    If LenB(alnum) > 0 Then toks = toks & alnum & "|"

    ' かな漢字を連結してから bigram を作る。こうすると、途中に紛れ込んだ
    ' 空白や記号(PDF由来の字詰め)が語を割らない。
    Dim m As Long: m = Len(jp)
    For i = 1 To m - 1
        toks = toks & Mid$(jp, i, 2) & "|"
    Next i

    If LenB(toks) > 0 Then toks = Left$(toks, Len(toks) - 1)
    Tokenize = toks
End Function

' ----------------------------------------------------------------------------
' DistinctiveKeys - 質問から「効く語」を抜く(| 区切り・最大8件)。
'   全件の部分一致スキャンに使う。長いものから採るのは、長い語ほど
'   希少で、当たったときの確度が高いため。
' ----------------------------------------------------------------------------
Public Function DistinctiveKeys(ByVal query As String) As String
    Dim s As String: s = NormalizeForSearch(query)
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Function

    ' 条番号は最優先(「第12条」は絶対に外せない)
    Dim keys As String
    Dim runs As String

    Dim cur As String
    Dim curKind As Long        ' 0=その他 1=英数 2=漢字
    Dim i As Long
    For i = 1 To n + 1
        Dim ch As String
        If i <= n Then ch = Mid$(s, i, 1) Else ch = " "
        Dim kind As Long
        If IsAlnumChar(ch) Then
            kind = 1
        ElseIf IsKanjiChar(ch) Then
            kind = 2
        Else
            kind = 0
        End If

        If kind <> curKind Or kind = 0 Then
            If curKind > 0 And Len(cur) >= 2 Then runs = runs & cur & "|"
            cur = ""
        End If
        If kind > 0 Then cur = cur & ch
        curKind = kind
    Next i

    ' 「第N条」を runs から拾って先頭へ置く
    Dim p As Long: p = InStr(s, "第")
    Do While p > 0
        Dim j As Long: j = p + 1
        Dim digits As String
        Do While j <= n
            Dim d As String: d = Mid$(s, j, 1)
            If d >= "0" And d <= "9" Then
                digits = digits & d
                j = j + 1
            Else
                Exit Do
            End If
        Loop
        If LenB(digits) > 0 And j <= n Then
            Dim unit As String: unit = Mid$(s, j, 1)
            If unit = "条" Or unit = "章" Or unit = "項" Or unit = "号" Then
                keys = keys & "第" & digits & unit & "|"
            End If
        End If
        p = InStr(p + 1, s, "第")
    Loop

    ' 残りは長い順に採る(単純な選択ソート。最大8件なので十分)
    Dim arr() As String
    If LenB(runs) > 0 Then
        arr = Split(Left$(runs, Len(runs) - 1), "|")
        Dim a As Long, b As Long
        For a = LBound(arr) To UBound(arr) - 1
            Dim best As Long: best = a
            For b = a + 1 To UBound(arr)
                If Len(arr(b)) > Len(arr(best)) Then best = b
            Next b
            If best <> a Then
                Dim tmp As String: tmp = arr(a): arr(a) = arr(best): arr(best) = tmp
            End If
        Next a
        For a = LBound(arr) To UBound(arr)
            If CountSep(keys) >= MAX_KEYS Then Exit For
            If InStr(1, "|" & keys, "|" & arr(a) & "|", vbTextCompare) = 0 Then
                keys = keys & arr(a) & "|"
            End If
        Next a
    End If

    If LenB(keys) > 0 Then keys = Left$(keys, Len(keys) - 1)
    DistinctiveKeys = keys
End Function

' ----------------------------------------------------------------------------
' Bm25Score - 1文書ぶんのBM25スコア。
'   qTokens : 質問のトークン(| 区切り)
'   docText : 文書本文(正規化前でよい。内部でTokenizeする)
'   dfCsv   : 各質問トークンの文書頻度(カンマ区切り・qTokensと同順)
'   docCount: 母集団の文書数 / avgLen: 平均トークン数
' ----------------------------------------------------------------------------
Public Function Bm25Score(ByVal qTokens As String, ByVal docText As String, _
                          ByVal dfCsv As String, ByVal docCount As Long, _
                          ByVal avgLen As Double) As Double
    If LenB(qTokens) = 0 Then Exit Function
    If docCount < 1 Then Exit Function
    If avgLen <= 0 Then avgLen = 1

    Dim dTok As String: dTok = Tokenize(docText)
    If LenB(dTok) = 0 Then Exit Function

    ' Scripting.Dictionary は Windows 専用。R4(純ロジック=LibreOfficeでも
    ' 動く)を守るため、区切り文字つき文字列と InStr だけで頻度を数える。
    ' 前後に区切りを足すのは「保険」が「生命保険金」に部分一致するのを
    ' 防ぐため(トークン単位の完全一致でなければ頻度にならない)。
    Dim dBar As String: dBar = "|" & dTok & "|"
    Dim dLen As Long: dLen = CountChar(dTok, "|") + 1

    Dim qArr() As String: qArr = Split(qTokens, "|")
    Dim dfArr() As String
    If LenB(dfCsv) > 0 Then dfArr = Split(dfCsv, ",")

    Dim seen As String: seen = "|"
    Dim total As Double
    Dim i As Long
    For i = LBound(qArr) To UBound(qArr)
        Dim t As String: t = qArr(i)
        If LenB(t) = 0 Then GoTo NextT
        If InStr(1, seen, "|" & t & "|", vbBinaryCompare) > 0 Then GoTo NextT
        seen = seen & t & "|"

        Dim f As Double: f = CountToken(dBar, t)
        If f = 0 Then GoTo NextT

        Dim df As Double: df = 1
        If LenB(dfCsv) > 0 Then
            If i <= UBound(dfArr) Then df = Val(dfArr(i))
        End If
        If df < 1 Then df = 1
        If df > docCount Then df = docCount

        Dim idf As Double
        idf = Log(1 + (docCount - df + 0.5) / (df + 0.5))

        total = total + idf * (f * (BM25_K1 + 1)) / _
                (f + BM25_K1 * (1 - BM25_B + BM25_B * dLen / avgLen))
NextT:
    Next i
    Bm25Score = total
End Function

' 区切り付き文字列 "|a|b|a|" の中に "|t|" が何回現れるか。
Private Function CountToken(ByVal barText As String, ByVal t As String) As Long
    Dim needle As String: needle = "|" & t & "|"
    Dim p As Long: p = InStr(1, barText, needle, vbBinaryCompare)
    Do While p > 0
        CountToken = CountToken + 1
        ' 区切りを共有するので1文字だけ進める(重なりを数え落とさない)
        p = InStr(p + 1, barText, needle, vbBinaryCompare)
    Loop
End Function

Private Function CountChar(ByVal s As String, ByVal ch As String) As Long
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) = ch Then CountChar = CountChar + 1
    Next i
End Function

' ----------------------------------------------------------------------------
' ExactHitCount - 「効く語」のうち何個が本文に完全一致で含まれるか。
'   条番号・型番の取りこぼしを決定的に防ぐための加点材料。
' ----------------------------------------------------------------------------
Public Function ExactHitCount(ByVal keys As String, ByVal docText As String) As Long
    If LenB(keys) = 0 Then Exit Function
    Dim body As String: body = NormalizeForSearch(docText)
    If LenB(body) = 0 Then Exit Function

    Dim arr() As String: arr = Split(keys, "|")
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If LenB(arr(i)) > 0 Then
            If InStr(1, body, arr(i), vbBinaryCompare) > 0 Then
                ExactHitCount = ExactHitCount + 1
            End If
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(純ロジック)
' ----------------------------------------------------------------------------

Private Function CountSep(ByVal s As String) As Long
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) = "|" Then CountSep = CountSep + 1
    Next i
End Function

Private Function IsAlnumChar(ByVal ch As String) As Boolean
    If LenB(ch) = 0 Then Exit Function
    If ch >= "0" And ch <= "9" Then IsAlnumChar = True: Exit Function
    If ch >= "a" And ch <= "z" Then IsAlnumChar = True
End Function

' ひらがな(12353-12447) / カタカナ(12448-12543) / 漢字(19968-40959) / 長音符
Private Function IsJapaneseChar(ByVal ch As String) As Boolean
    Dim c As Long: c = CodePointOf(ch)
    If c = 0 Then Exit Function
    If c >= 12353 And c <= 12447 Then IsJapaneseChar = True: Exit Function
    If c >= 12448 And c <= 12543 Then IsJapaneseChar = True: Exit Function
    If c >= 19968 And c <= 40959 Then IsJapaneseChar = True
End Function

Private Function IsKanjiChar(ByVal ch As String) As Boolean
    Dim c As Long: c = CodePointOf(ch)
    If c >= 19968 And c <= 40959 Then IsKanjiChar = True: Exit Function
    If c >= 12448 And c <= 12543 Then IsKanjiChar = True    ' カタカナ語も固有名詞になる
End Function

' AscW は符号付きIntegerを返すため、BMP後半の文字は負値になる。必ず補正する
' (この補正漏れが modChunker/modClarify で実バグになっていた。2026-07-27)。
Private Function CodePointOf(ByVal ch As String) As Long
    If LenB(ch) = 0 Then Exit Function
    Dim c As Long: c = AscW(ch)
    If c < 0 Then c = c + 65536
    CodePointOf = c
End Function
