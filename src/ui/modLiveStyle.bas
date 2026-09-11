Attribute VB_Name = "modLiveStyle"
Option Explicit

' ============================================================================
' modLiveStyle - AI回答バブルの文字装飾一式(R43 波A・modLiveからの分割)。
' ----------------------------------------------------------------------------
' なぜ分割したか:
'   modLive は27,825字(WARN 28,000まで残り175字)で、新しい装飾(結論の強調・
'   セル番地の色)を1行足す余地も無かった。装飾一式(StyleAnswerParas/
'   CiteTagSpans/TagCloseOk・実測4,948字)をこちらへ移し、modLive には
'   同シグネチャの1行委譲だけを残す(modSkin→modToastと同型の分割。挙動は
'   一字も変えていない)。
'
'   外部からの呼び出し口は StyleAnswerParas 1本だけ(modApp.bas 2箇所から
'   modLive.StyleAnswerParas 経由。呼び出し側は1文字も変えていない)。
'   CiteTagSpans/TagCloseOk は modLive 経由の外部呼び出しが元々無かった
'   (modLive.StyleAnswerParas の内部専用+テストのみ)ため、委譲を残さず
'   丸ごと移設した。テスト呼び出し側(modTestsPure43)は modLiveStyle への
'   直接呼び出しへ書き替え済み(CLAUDE.md §9のシグネチャ変更時の作法)。
'
' 装飾の順序(唯一の落とし穴): 必ず「段落レベル(■見出しの太字・結論段落の
'   強調)→ 文字レベル(出典タグ・セル番地の8pt化)」の順で当てる。逆にすると
'   結論段落の中にある出典タグ・番地が12ptへ引きずられて潰れる
'   (StyleAnswerParas 内のコメント参照)。
'
' 位置の数え方: 生成直後に組み立てた文字列ではなく、Shapeへ代入した後の
'   TextFrame2.TextRange.Text を読み戻してから数える(vbLfが読み戻しで
'   vbCrになりうるため。読み戻し前の文字列で計算すると実機だけ位置が
'   ずれる。旧 modLive.bas:246 と同じ作法)。
' ============================================================================

' ----------------------------------------------------------------------------
' StyleAnswerParas - 「■」見出しの太字化・結論段落の強調・出典タグ/セル番地の
'   小型化を1回のシェイプ走査でまとめて行う(2026-08-03 R14-8c 発祥、
'   R40 F2で出典タグ、R43 波A で LeadParaIndex/CellRefSpans を追加)。
'   表示の失敗が回答を壊してはならないので全体を On Error Resume Next で
'   包む(装飾が付かないことはあっても、ここで質問が落ちることはない)。
' ----------------------------------------------------------------------------
Public Sub StyleAnswerParas(ByVal bubbleName As String)
    If LenB(bubbleName) = 0 Then Exit Sub
    On Error Resume Next
    Dim shp As Shape
    Set shp = ThisWorkbook.Worksheets("Nexus").Shapes(bubbleName)
    If shp Is Nothing Then Exit Sub
    Dim pcount As Long
    pcount = shp.TextFrame2.TextRange.Paragraphs.count
    If pcount < 1 Then Exit Sub

    ' ---- 段落レベル(■見出しの太字・結論段落の強調)を先に当てる -------------
    Dim i As Long
    For i = 1 To pcount
        Dim head As String
        head = Left$(LTrim$(shp.TextFrame2.TextRange.Paragraphs(i).Text), 1)
        If head = "■" Then
            shp.TextFrame2.TextRange.Paragraphs(i).Font.Bold = True
        ElseIf head = ChrW(&H203B) Then
            ' R44: 末尾の常設ガード(modMode.GuardNoteText)。毎回出るものなので
            ' 主役にはしない。
            ' R46: ただし 8pt/muted は落としすぎだった。直下のフッターが
            ' 8.5pt/同じ muted(modLive.StyleFooter)なので、ガードが【バブルの
            ' 中で一番小さく一番薄い行】になり、実機で「見にくい」と報告された。
            ' 9pt・太字・danger にする。サイズの序列は保つ:
            '   結論12pt太字 > 本文10.5pt > ガード9pt太字 > フッター8.5pt > 出典8pt
            ' 太字が要るのは見た目だけの理由ではない。次の質問を投げると
            ' modUI.MarkActiveBubble が非選択バブルへ modSkin.PaintBubble を掛け、
            ' その末尾の SetShapeTextColor が【バブル全体の文字色を一括上書き】
            ' するため、色だけの装飾は過去バブルから消える。Size と Bold は
            ' 上書きされないので、そちらが本体で色はおまけになる。
            ' 色は danger(テーマごとに明系/暗系を出し分け)を通す。赤のRGBを
            ' 直書きすると dark で 2.26:1 まで落ちて AA を割る。
            With shp.TextFrame2.TextRange.Paragraphs(i).Font
                .Size = 9
                .Bold = True
                .Fill.ForeColor.RGB = modUI.UiColor("danger")
            End With
        End If
    Next i

    Dim fullText As String
    fullText = shp.TextFrame2.TextRange.Text   ' 読み戻し(vbLfがvbCrになりうる)

    Dim leadIdx As Long: leadIdx = LeadParaIndex(fullText)
    If leadIdx >= 1 And leadIdx <= pcount Then
        With shp.TextFrame2.TextRange.Paragraphs(leadIdx).Font
            .Size = 12
            .Bold = True
        End With
    End If

    ' ---- 文字レベル(出典タグ・セル番地)は段落装飾の後 ------------------------
    Dim starts() As Long, lens() As Long
    Dim nTag As Long: nTag = CiteTagSpans(fullText, starts, lens)
    For i = 1 To nTag
        With shp.TextFrame2.TextRange.Characters(starts(i), lens(i)).Font
            .Size = 8
            .Fill.ForeColor.RGB = modUI.UiColor("primary")
        End With
    Next i

    Dim cStarts() As Long, cLens() As Long
    Dim nCell As Long: nCell = CellRefSpans(fullText, cStarts, cLens)
    For i = 1 To nCell
        With shp.TextFrame2.TextRange.Characters(cStarts(i), cLens(i)).Font
            .Size = 8
            .Fill.ForeColor.RGB = modUI.UiColor("primary")
        End With
    Next i

    ' ---- 高さの再フィットと下端の取り直し(レビュー R43 M2) -------------------
    ' 結論段落を 12pt にすると文字が箱より背が高くなりうる。AutoSize は一度
    ' 切って入れ直さないと再フィットしない(modUI.UpdateBubbleText:283-297 が
    ' 同じ2手を踏んでいる)。下端も取り直さないと、次のバブル(起動時の履歴
    ' 復元は1件ずつ装飾する)がこのバブルへ重なる。
    With shp.TextFrame2
        .WordWrap = -1        ' 手本(modUI.UpdateBubbleText)と同じ3手にする
        .AutoSize = 0
        .AutoSize = 1
    End With
    If shp.Height < 28 Then shp.Height = 28
    modUI.RecalcChatBottom shp.Parent
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' LeadParaIndex - 先頭段落を「結論」として強調してよいかどうか(純関数・
'   R43 波A 1-1追加2)。結論の強調は「1文目にもう答えが来ている」体験を
'   作るための書式なので、1文目が結論ではない回答にまで太字12ptを当てると
'   逆に読みにくくなる。次の2経路は先頭段落が結論ではないため対象外にする:
'
'   (1) ■ 見出しで始まる回答: NormalizeAnswerText が「■」見出しの前に
'       空行を入れるのは【本文の途中】の見出しに対してだけで(R14-G9)、
'       モデルが1文目から■見出しで書いた回答はそのまま先頭に来る
'       (modLive.NormalizeAnswerLine)。1文目が見出しなら、それは結論文
'       ではなく後に続く本文の表題であり、太字12ptを重ねると見出しの太字と
'       二重に浮いてしまう。
'   (2) 聞き返し・番号選択の回答: modClarify.BuildClarifyPrompt(src/qa/
'       modClarify.bas:98,101)と modAskMulti.BuildClarifyAsk(src/qa/
'       modAskMulti.bas:390)が組む「もう少しだけ教えてください」
'       「ご質問はいくつかの読み方ができます」は、どちらも文頭が
'       💭(ChrW(&HD83D)&ChrW(&HDCAD))で始まる固定書式。これは結論では
'       なく「これから番号で選んでもらう問い」なので、結論用の強調は
'       付けない(この絵文字で始まる回答は他に無く誤爆しない。
'       modApp.bas:141の「考えています…」プレースホルダも同じ絵文字を
'       使うが、StyleAnswerParasを通す前に削除される別バブルなので
'       ここには到達しない)。
'
'   戻り値は Shape.Paragraphs のインデックス(1始まり)。対象外なら 0。
'   text は段落区切りが vbCr または vbLf のどちらでも判定できるようにする
'   (Shapeからの読み戻しはvbCr、テストからの直接呼び出しはvbLfで組む
'   ことがあるため)。
' ----------------------------------------------------------------------------
Public Function LeadParaIndex(ByVal text As String) As Long
    If LenB(text) = 0 Then Exit Function

    Dim brk As Long: brk = InStr(text, vbCr)
    If brk = 0 Then brk = InStr(text, vbLf)

    Dim firstPara As String
    If brk > 0 Then
        firstPara = Left$(text, brk - 1)
    Else
        firstPara = text
    End If
    firstPara = LTrim$(firstPara)
    If LenB(firstPara) = 0 Then Exit Function

    If Left$(firstPara, 1) = "■" Then Exit Function
    If Left$(firstPara, 2) = ChrW(&HD83D) & ChrW(&HDCAD) Then Exit Function
    ' レビュー R43 M1: 先頭に前置されるだけで結論ではない2種。
    ' (3) ⚠ 始まりの前置2種: modAskRetrieve.ApplyLowHitWarning(関連が薄い)と
    '     modLive.EmptyShelfWrap(本棚が空)。どちらも結論ではない前置きで、
    '     強調すると警告文が画面で一番大きくなる。この2経路では本文側も
    '     強調しない(次の段落へずらさない割り切り)。
    ' (4) modApp.RestoreLastConversation が付ける「📜(前回の回答) 」。これが
    '     前置されると ■ も 💭 も貫通するので、復元履歴が全件12ptになる。
    If Left$(firstPara, 1) = ChrW(&H26A0) Then Exit Function
    If Left$(firstPara, 2) = ChrW(&HD83D) & ChrW(&HDCDC) Then Exit Function

    LeadParaIndex = 1
End Function

' ----------------------------------------------------------------------------
' CellRefSpans - 本文中の「(A6 付近)」型セル番地の位置と長さを列挙する
'   (純関数・R43 波A 1-3)。プロンプト側の指示(modPrompts.bas:499・
'   modAskGlobal.bas:540・modAskOnePass.bas:560)がすべて
'   "(<列letters><行数字> 付近)" の形で書かせているため、その形だけを
'   対象にする(前後の半角括弧を含めて1件)。
'   戻り値は件数、starts/lensは1始まり(TextRange.Charactersと同じ数え方)。
' ----------------------------------------------------------------------------
Public Function CellRefSpans(ByVal s As String, ByRef starts() As Long, _
                             ByRef lens() As Long) As Long
    Dim n As Long
    ReDim starts(1 To 1): ReDim lens(1 To 1)
    Dim p As Long: p = 1
    Do
        Dim op As Long: op = InStr(p, s, "(")
        If op = 0 Then Exit Do
        Dim closeAt As Long: closeAt = 0
        If CellRefCloseAt(s, op + 1, closeAt) Then
            n = n + 1
            ReDim Preserve starts(1 To n): ReDim Preserve lens(1 To n)
            starts(n) = op
            lens(n) = closeAt - op + 1
            p = closeAt + 1
        Else
            p = op + 1
        End If
    Loop
    CellRefSpans = n
End Function

' "(" の直後(startPos)から "<列letters(A-Z 1文字以上)><行数字(1文字以上)> 付近)"
' が続くかどうかを判定する(Private・純)。続けば closeAt に ")" の位置を
' 入れて True を返す。列・行のどちらかが1文字も読めない、" 付近" が続かない、
' 直後が ")" でない、のいずれかで False(そのままセル番地としては扱わない)。
Private Function CellRefCloseAt(ByVal s As String, ByVal startPos As Long, _
                                ByRef closeAt As Long) As Boolean
    Dim total As Long: total = Len(s)
    Dim i As Long: i = startPos

    Dim letStart As Long: letStart = i
    Do While i <= total
        If Mid$(s, i, 1) < "A" Or Mid$(s, i, 1) > "Z" Then Exit Do
        i = i + 1
    Loop
    If i = letStart Then Exit Function          ' 列の文字が1つも無い

    Dim digStart As Long: digStart = i
    Do While i <= total
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Do
        i = i + 1
    Loop
    If i = digStart Then Exit Function           ' 行の数字が1つも無い

    If Mid$(s, i, 3) <> " 付近" Then Exit Function
    i = i + 3
    If Mid$(s, i, 1) <> ")" Then Exit Function

    closeAt = i
    CellRefCloseAt = True
End Function

' ----------------------------------------------------------------------------
' CiteTagSpans - 本文中の出典タグの位置と長さを列挙する(純関数・R40 F2。
'   R43 波A で modLive から移設。ロジックは1字も変えていない)。
'   "[本棚:" または "[パック(" で始まり、同じ段落内の最初の "]" まで(段落を
'   跨ぐ・閉じ無しは除外)。戻り値は件数、starts/lensは1始まり
'   (TextRange.Characters と同じ数え方。段落区切りの vbCr も1字として
'   数える=Text の位置そのまま)。
' ----------------------------------------------------------------------------
Public Function CiteTagSpans(ByVal s As String, ByRef starts() As Long, _
                             ByRef lens() As Long) As Long
    Dim n As Long
    ReDim starts(1 To 1): ReDim lens(1 To 1)
    Dim p As Long: p = 1
    Do
        Dim a As Long: a = InStr(p, s, "[本棚:")
        Dim b As Long: b = InStr(p, s, "[パック(")
        If a = 0 And b = 0 Then Exit Do
        Dim st As Long
        If a = 0 Then
            st = b
        ElseIf b = 0 Then
            st = a
        ElseIf a < b Then
            st = a
        Else
            st = b
        End If
        Dim en As Long: en = InStr(st, s, "]")
        If en = 0 Then Exit Do
        Dim brk As Long: brk = InStr(st, s, vbCr)
        If brk = 0 Then brk = InStr(st, s, vbLf)
        ' 閉じ括弧より前に次のタグが始まる=この開始は閉じていない(レビュー R40 m1)。
        Dim nx As Long: nx = InStr(st + 1, s, "[本棚:")
        Dim nx2 As Long: nx2 = InStr(st + 1, s, "[パック(")
        If nx = 0 Or (nx2 > 0 And nx2 < nx) Then nx = nx2

        ' R41 §3 C2(レビュー両者 M4): "[本棚:" 形は資料名に "]" を含みうる
        ' (例: report[1].pdf)。最初の "]" の直前が「数字」で、その前が
        ' "p." または "シート" でなければ閉じ候補として認めず、同じ段落・
        ' 次のタグ開始より前に限って次の "]" を最大2回まで試す。該当が無ければ
        ' 従来どおり最初の "]"(空振りで本文を壊さない安全弁)。"[パック(" 形は
        ' 従来どおり(資料名にページ接尾辞が無く判定できないため §5 記録のみ)。
        If a > 0 And st = a Then
            If Not TagCloseOk(s, st, en) Then
                Dim tryFrom As Long: tryFrom = en
                Dim tries As Long
                For tries = 1 To 2
                    Dim cand As Long: cand = InStr(tryFrom + 1, s, "]")
                    If cand = 0 Then Exit For
                    If brk > 0 And brk < cand Then Exit For
                    If nx > 0 And nx < cand Then Exit For
                    If TagCloseOk(s, st, cand) Then
                        en = cand
                        Exit For
                    End If
                    tryFrom = cand
                Next tries
            End If
        End If

        If (brk > 0 And brk < en) Or (nx > 0 And nx < en) Then
            p = st + 1
        Else
            n = n + 1
            ReDim Preserve starts(1 To n): ReDim Preserve lens(1 To n)
            starts(n) = st
            lens(n) = en - st + 1
            p = en + 1
        End If
    Loop
    CiteTagSpans = n
End Function

' TagCloseOk - 位置 en の "]" が "[本棚:" 形タグの正しい閉じかどうか(純関数・
'   R41 §3 C2。R43 波A で modLive から移設)。en-1 から数字を1桁以上遡り、
'   その直前が "p." または "シート" ならOK。資料名に含まれる "]"
'   (例: report[1].pdf)は数字の直前がそのトークンにならないのでNGになり、
'   CiteTagSpans が次の "]" を試す。
Private Function TagCloseOk(ByVal s As String, ByVal st As Long, ByVal en As Long) As Boolean
    Dim i As Long: i = en - 1
    If i < st Then Exit Function
    Dim c As String: c = Mid$(s, i, 1)
    If c < "0" Or c > "9" Then Exit Function   ' "]" の直前が数字でなければNG
    Do While i > st
        c = Mid$(s, i - 1, 1)
        If c < "0" Or c > "9" Then Exit Do
        i = i - 1
    Loop
    ' i = 数字の先頭位置。その直前2文字が "p."、または直前3文字が "シート"。
    If i - 2 >= st Then
        If Mid$(s, i - 2, 2) = "p." Then
            TagCloseOk = True
            Exit Function
        End If
    End If
    If i - 3 >= st Then
        If Mid$(s, i - 3, 3) = "シート" Then
            TagCloseOk = True
        End If
    End If
End Function
