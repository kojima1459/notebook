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
' 速度と、なぜBM25を毎クエリ全件に掛けないか(実測に基づく)
' ----------------------------------------------------------------------------
' LibreOffice実測(閾値テストで固定):
'     Tokenize  … 900字1件で 225ms → 9000件で 2,025秒(約34分)
'     InStr     … 9000件×8キーで 0.5秒未満
' VBAはLO Basicより速いはずだが桁は変わらない。毎クエリ全件トークナイズは
' 選べない。実際、この構成を一度実装して34分かかることを実測で確認した。
'
' そこで出荷する経路(modRetrieve.Search)は:
'     ・チャンクごとにやるのは CompactForMatch + KeyScore(InStrのみ)
'     ・Tokenize / Bm25Score は毎クエリの全件走査では呼ばない
'       (パック作成時のオフライン処理や、将来の小プール再ランク用に残す)
'
' 精度はどうか(実物6資料180チャンク31問での採点):
'     文字bigram BM25(全件)      R@1 84% / R@3 97% / R@5 100% / ページ45%
'     ★採用: InStr・空白除去      R@1 84% / R@3 84% / R@5  90% / ページ81%
'   R@3/R@5 は BM25 が上だが R@1 は同じで、出典ページの正確さは採用案が
'   大きく上回る。「どの資料の何ページか」を出す製品なので後者を採る。
'   表記揺れ(全角/大文字/空白挿入/空白除去)は全て 84% を維持した。
'
' dense(埋め込み)との関係:
'   modRetrieve は全チャンクに対して dense(内積)と本モジュール(KeyScore)の
'   両方を必ず計算する。片方で候補を絞ってからもう片方を掛ける構成にはしない。
'   ベクトルで拾えない専門用語が一段目で足切りされる事故を、構造的に防ぐため。
'
' R4準拠: Excelオブジェクトにも Scripting.Dictionary にも触れないので
' LibreOffice でそのままテストできる。回帰は modTestsPure2 が固定する。
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

' 2026-08-10(R27 F1-1/F1-2): KeyScore の頭打ち。詳細は KeyScore の注記。
Private Const KEY_LEN_CLAMP As Long = 8          ' 語長重みの Len 上限
Private Const KEYSCORE_CAP_DEFAULT As Double = 10#
Private mCapLoaded As Boolean
Private mCapValue As Double

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

    ' 【重要・実測で判明】文字列を "toks = toks & x" で伸ばしてはいけない。
    ' VBAの文字列連結は毎回コピーが走るためO(n^2)になり、900字のチャンク1件に
    ' 200ms かかっていた(LibreOffice実測)。9000件なら1800秒=30分。
    ' 配列に詰めて最後に Join する(O(n))。これは規約§12にも明記がある。
    ' この1点だけで実用と非実用が分かれるので、絶対に戻さないこと。
    Dim buf() As String
    ReDim buf(1 To n * 2 + 2)
    Dim cnt As Long

    Dim jpBuf() As String
    ReDim jpBuf(1 To n)
    Dim jpN As Long

    Dim alnumBuf() As String
    ReDim alnumBuf(1 To n)
    Dim alnumN As Long

    Dim i As Long
    For i = 1 To n
        Dim ch As String: ch = Mid$(s, i, 1)
        If IsAlnumChar(ch) Then
            alnumN = alnumN + 1
            alnumBuf(alnumN) = ch
        Else
            If alnumN > 0 Then
                cnt = cnt + 1
                buf(cnt) = Join(SliceOf(alnumBuf, alnumN), "")
                alnumN = 0
            End If
            If IsJapaneseChar(ch) Then
                jpN = jpN + 1
                jpBuf(jpN) = ch
            End If
        End If
    Next i
    If alnumN > 0 Then
        cnt = cnt + 1
        buf(cnt) = Join(SliceOf(alnumBuf, alnumN), "")
    End If

    ' かな漢字を連結してから bigram を作る。こうすると、途中に紛れ込んだ
    ' 空白や記号(PDF由来の字詰め)が語を割らない。
    For i = 1 To jpN - 1
        cnt = cnt + 1
        buf(cnt) = jpBuf(i) & jpBuf(i + 1)
    Next i

    If cnt = 0 Then Exit Function
    Tokenize = Join(SliceOf(buf, cnt), "|")
End Function

' 配列の先頭n個だけを取り出す(Joinへ渡すため。ReDim Preserveより安い)。
Private Function SliceOf(ByRef src() As String, ByVal n As Long) As String()
    Dim out() As String
    ReDim out(1 To n)
    Dim i As Long
    For i = 1 To n
        out(i) = src(i)
    Next i
    SliceOf = out
End Function

' ----------------------------------------------------------------------------
' DistinctiveKeys - 質問から「効く語」を抜く(| 区切り・最大8件)。
'   全件の部分一致スキャンに使う。長いものから採るのは、長い語ほど
'   希少で、当たったときの確度が高いため。
' ----------------------------------------------------------------------------
Public Function DistinctiveKeys(ByVal query As String) As String
    ' 空白は先に全部落とす。落とさないと「保 険 金」が1文字断片に割れて
    ' 全部捨てられ、キーが1つも取れなくなる(実測: 空白を入れただけで
    ' R@1 84%→45% まで落ちた)。
    Dim s As String: s = Replace(NormalizeForSearch(query), " ", "")
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

    ' 「別表N」「様式N」も条番号と同格の最優先キーにする(2026-08-05 R17 Phase1)。
    ' 規程の答えが本文ではなく別表・様式にある質問(「別表2の料率は?」)で、
    ' 「別表」だけが2文字ランとして残り番号が落ちていた=どの別表でも同点になる。
    keys = keys & TableKeys(s, n, "別表", keys)
    keys = keys & TableKeys(s, n, "様式", keys)

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
' TableKeys - 「別表N」「様式N」を s から拾い "ラベル|" の並びで返す(R17 Phase1)。
'   上の「第」ループと同型(起点をInStrで送りながら数字ランを読む)。違いは
'   単位漢字が先頭側にあることと、「別表第2」のように間に「第」が入る書き方が
'   あること。ラベルは原文どおりに組む(同じ資料の中では表記が揃っているのが
'   普通で、勝手に「第」を落として寄せると本文側と一致しなくなる)。
'   already には既に採ったキー列("ラベル|"の並び)を渡す=重複を作らない。
' ----------------------------------------------------------------------------
Private Function TableKeys(ByVal s As String, ByVal n As Long, ByVal head As String, _
                           ByVal already As String) As String
    Dim outS As String
    Dim hl As Long: hl = Len(head)
    Dim p As Long: p = InStr(s, head)
    Do While p > 0
        Dim j As Long: j = p + hl
        Dim dai As String: dai = ""
        If j <= n Then
            If Mid$(s, j, 1) = "第" Then
                dai = "第"
                j = j + 1
            End If
        End If
        Dim dgt As String: dgt = ""
        Do While j <= n
            Dim dch As String: dch = Mid$(s, j, 1)
            If dch >= "0" And dch <= "9" Then
                dgt = dgt & dch
                j = j + 1
            Else
                Exit Do
            End If
        Loop
        If LenB(dgt) > 0 Then
            ' 2026-08-05(R17H FB-3 / A-L13): MAX_KEYS の枠はここでも尊重する。
            ' 別表・様式を無制限に足すと、DistinctiveKeys が8枠を超えて返り、
            ' 「残りは長い順に採る」側の打ち切り(CountSep>=MAX_KEYS)が意味を
            ' 失う=枠の意味が採る順序で変わってしまう。
            If CountSep(already & outS) >= MAX_KEYS Then Exit Do
            Dim lab As String: lab = head & dai & dgt
            If InStr(1, "|" & already & outS, "|" & lab & "|", vbBinaryCompare) = 0 Then
                outS = outS & lab & "|"
            End If
        End If
        p = InStr(p + 1, s, head)
    Loop
    TableKeys = outS
End Function

' ----------------------------------------------------------------------------
' CompactForMatch - 照合専用テキスト(正規化 + 空白を全除去)。
' ----------------------------------------------------------------------------
' InStr は連続一致なので、PDF由来の字詰めやユーザーが入れた余計な空白が
' 語の途中に入ると当たらなくなる。照合用に空白を落とした版を使えば
' 「保 険 金」と「保険金」が同じものとして一致する。
' 表示・出典に使うのは元の本文なので、見た目は一切変わらない。
Public Function CompactForMatch(ByVal s As String) As String
    CompactForMatch = Replace(NormalizeForSearch(s), " ", "")
End Function

' ----------------------------------------------------------------------------
' MatchDocText - 1チャンクの「照合用テキスト」を作る唯一の場所(2026-08-01 R12-4)。
' ----------------------------------------------------------------------------
' 要約・キーワード・資料名・本文をこの順で連結し、CompactForMatch を通す。
' R12-4 で my_knowledge の第10列(norm_text)へ取込時に前計算して保存するように
' したため、この式が2箇所にあると【保存済みの行と未保存の行でスコアが変わる】。
' 取込側(modShelf)・検索側(modRetrieve)・遅延バックフィルの三者が必ずここを
' 通ること。連結の順番や区切りを変えると保存済みデータと食い違うので変えない
' (空白は CompactForMatch が全て落とすので、区切りの見た目は結果に影響しない)。
Public Function MatchDocText(ByVal summary As String, ByVal keywords As String, _
                             ByVal source As String, ByVal fullText As String) As String
    MatchDocText = CompactForMatch(summary & " " & keywords & " " & source & " " & fullText)
End Function

' ----------------------------------------------------------------------------
' KeyScore - トークナイズせずに求めるキーワードスコア。
' ----------------------------------------------------------------------------
' なぜBM25を毎クエリ全件に掛けないか(実測に基づく設計判断):
'   Tokenize は1チャンク(900字)あたり 225ms かかる(LibreOffice実測)。
'   9000チャンクなら 2,025秒 = 34分。VBAはこれより速いはずだが、
'   桁が違う話ではないので、毎クエリ全件トークナイズは選べない。
'   一方 InStr はネイティブ実装で、9000件×8キーが 0.5秒以内に収まる
'   (同じく実測・閾値テストで固定済み)。
'
' 精度はどうか(実物6資料180チャンク31問での採点):
'     文字bigram BM25            R@1 84% / R@3 97% / R@5 100% / ページ45%
'     本方式(InStr・空白除去)     R@1 84% / R@3 84% / R@5  90% / ページ81%
'   R@3/R@5 は BM25 が上だが、R@1 は同じで、出典ページの正確さは本方式が
'   大きく上回る。「どの資料の何ページか」を出す製品なので後者を採る。
'   さらに本方式は全ての表記揺れ(全角/大文字/空白挿入/空白除去)で
'   84% を完全維持した。
'
'   keys       : DistinctiveKeys の結果(| 区切り)
'   compactDoc : CompactForMatch を通した本文
'
' 2026-08-10(R27 F1-1/F1-2・実機第12報②「特約が検索から消える」の根治):
'   実機ログの確定機序は【KeyScoreが無上限】だった。modRetrieve は
'   SPARSE_WEIGHT=0.06 を掛けて cos(-1〜1)と同じ土俵へ乗せる前提で書かれて
'   いるが、thorough の HyDE と会話履歴の独立質問化が「長い資料名」を丸ごと
'   キーへ注入するため、KeyScore が 330 まで伸びた(0.06×330 = 20.26)。
'   その結果 cos 成分は最終スコアの2〜5%しか占めず、実質キーワード検索に
'   なっていた。曖昧判定(0.6)・低ヒット警告(0.3)・確信度55 という3つの
'   安全装置は全て「スコアが 0〜1 付近」を前提にした閾値なので、同時に
'   無力化されていた(どれも鳴らないので誰も気付けない)。
'   ・F1-1 = 戻り値を cap で頭打ち(既定10 → 0.06×10 = 0.6 で cos と同格)。
'   ・F1-2 = 語長重み Len^1.5 の Len を8で clamp(16字の資料名キー1本で
'     64点入るのを止める。8^1.5 = 22.6 が上限)。
'   頭打ちは「上限へ張り付いた同点の山」を作るが、そこへ届くのは既に
'   決定的に強く一致したチャンクだけで、その先の順位は cos が付ける
'   (キーワードは候補を引き上げる役、順位を決めるのはベクトル、という
'   本来の設計に戻す)。
Public Function KeyScore(ByVal keys As String, ByVal compactDoc As String) As Double
    If LenB(keys) = 0 Then Exit Function
    If LenB(compactDoc) = 0 Then Exit Function

    Dim arr() As String: arr = Split(keys, "|")
    Dim total As Double
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Dim k As String: k = arr(i)
        If Len(k) >= 2 Then
            Dim c As Long: c = CountOccurrences(compactDoc, k)
            If c > 0 Then
                ' 長い語ほど希少 = IDFの安価な代用。^1.5 は実測で選んだ形。
                Dim w As Double: w = KeyLenWeight(Len(k))
                total = total + w * (1 + Log(CDbl(c)))
            End If
        End If
    Next i
    If total <= 0 Then Exit Function

    ' 文書長で正規化(長いチャンクが不当に有利になるのを防ぐ)
    KeyScore = CapKeyScore(total / (1 + Log(1 + Len(compactDoc) / 900#)), _
                           EffectiveKeyScoreCap())
End Function

' ----------------------------------------------------------------------------
' KeyLenWeight - キー1語ぶんの語長重み(2026-08-10 R27 F1-2)。純関数。
' ----------------------------------------------------------------------------
' Len^1.5 は「長い語ほど希少」の安価な代用だが、上限が無いと HyDE や資料名
' 由来の長語1本だけで全チャンクの順位が決まる。8字より長い語が更に希少に
' なるという根拠は無い(実測の採点は2〜8字のキーで取ったもの)ので、
' 8で寝かせる。7→8 は増え、8→9 以降は増えない、が仕様。
Public Function KeyLenWeight(ByVal keyLen As Long) As Double
    Dim L As Long: L = keyLen
    If L < 0 Then L = 0
    If L > KEY_LEN_CLAMP Then L = KEY_LEN_CLAMP
    KeyLenWeight = L ^ 1.5
End Function

' ----------------------------------------------------------------------------
' CapKeyScore - 生スコアを cap で頭打ちにする(2026-08-10 R27 F1-1)。純関数。
'   cap <= 0 は「頭打ちなし」= R27以前の挙動へ戻すエスケープハッチ
'   (config sparse_keyscore_cap に 0 を入れると旧来どおり暴れる)。
' ----------------------------------------------------------------------------
Public Function CapKeyScore(ByVal raw As Double, ByVal cap As Double) As Double
    CapKeyScore = raw
    If cap <= 0# Then Exit Function
    If raw > cap Then CapKeyScore = cap
End Function

' cap の実効値。KeyScore は「1質問あたり全チャンク分」呼ばれる(実機9,000件)
' のに対し modConfig はキャッシュを持たず毎回 config シートを線形走査する。
' そのままでは毎質問 9,000×250 セル読みになるので、1セッション1回だけ読んで
' 控える(modBitwiseOpt のセッションキャッシュと同型。config を書き換えた
' ときはブックを開き直すと反映される)。
' modConfig が解決できない実行環境(LO実行テスト。純ロジック一式しか注入
' しない)では On Error Resume Next で既定値のまま走り切る。
Private Function EffectiveKeyScoreCap() As Double
    If Not mCapLoaded Then
        Dim v As Double: v = KEYSCORE_CAP_DEFAULT
        On Error Resume Next
        v = CDbl(modConfig.GetLong("sparse_keyscore_cap", CLng(KEYSCORE_CAP_DEFAULT)))
        On Error GoTo 0
        mCapValue = v
        mCapLoaded = True
    End If
    EffectiveKeyScoreCap = mCapValue
End Function

' ----------------------------------------------------------------------------
' HasAnyKey - keys(DistinctiveKeysの結果)のどれか1つでも doc に含まれるか。
' ----------------------------------------------------------------------------
' 2026-08-01(R12-3-7): バイナリ粗選別(binary_rag)の「救済union」専用の
' 粗い包含判定。狙いは「第12条」のような決定的な語を持つ行を候補集合から
' 落とさないことで、拾いすぎ側の誤りは後段のFloatスコアと KeyScore が
' 正しく並べ替える(落としたら二度と戻らないが、拾いすぎは直せる)。
'
' 2026-08-01(R12-H-1 敵対的レビュー High-1/Med-4): doc は【正規化済み】の
' 照合テキスト(modSparse.MatchDocText の結果 = my_knowledge.norm_text)を
' 渡すこと。当初は速度を理由に生テキストへ vbTextCompare で当てていたが、
'   ・keys(DistinctiveKeys)は正規化済み(小文字・半角・空白除去)なので、
'     生テキスト相手では全角英数「ＡＢＣ」やPDF字詰め「保 険 金」を
'     取りこぼす。救済のはずが救済漏れを作っていた
'   ・R12-4 で norm_text を前計算したので、正規化のやり直しはもう発生しない
'     (速度を理由に生テキストを見る動機が消えた)
' 正規化済み同士なので比較は vbBinaryCompare で足りる(vbTextCompare は
' ロケール依存の照合表を引くぶん遅く、正規化済みなら結果も変わらない)。
Public Function HasAnyKey(ByVal keys As String, ByVal doc As String) As Boolean
    If LenB(keys) = 0 Then Exit Function
    If LenB(doc) = 0 Then Exit Function

    Dim arr() As String: arr = Split(keys, "|")
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Len(arr(i)) >= 2 Then                   ' 1文字キーは拾いすぎるので見ない
            If InStr(1, doc, arr(i), vbBinaryCompare) > 0 Then
                HasAnyKey = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function CountOccurrences(ByVal hay As String, ByVal needle As String) As Long
    If LenB(needle) = 0 Then Exit Function
    Dim p As Long: p = InStr(1, hay, needle, vbBinaryCompare)
    Do While p > 0
        CountOccurrences = CountOccurrences + 1
        p = InStr(p + Len(needle), hay, needle, vbBinaryCompare)
    Loop
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
