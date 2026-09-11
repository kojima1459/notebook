Attribute VB_Name = "modPeek"
Option Explicit

' ============================================================================
' modPeek - ワンクリック出典ポップアップ(Peek View)。UI層。
' ----------------------------------------------------------------------------
' 役割:
'   AI回答の直下に「出典チップ」(📄 資料名 p.N)を並べ、クリックすると元の
'   チャンク本文が小さな吹き出し(ツールチップ)としてフワッと浮かぶ。PDFを開かず
'   審査員が「AIの回答が本当に約款に書いてあるか」を秒で照合でき、ハルシネーションを
'   人間が即検知・是正できる、というデモ最大のアピール機能。
'
' 設計判断:
'   ・出典データは modAsk の読み取り専用アクセサ(LastHit*)から取得する(検索/回答
'     ロジックには一切触れない)。UI層→qa層の下向き参照でR1レイヤリング準拠。
'   ・modUI.basは字数上限間際のため一切変更しない。描画は本モジュールに集約。
'   ・チップは「最新の回答」の下にだけ出す(次の送信でHideCitationsして描き直す)。
'     modUIのmChatBottom(private)を触らずに済ませるための割り切り。過去回答の
'     チップは残さない(Peekは今の回答の照合が目的)。
'   ・チップ/ポップアップはShape。クリック時はOnActionでmodApp.OnPeek/OnPeekCloseへ。
'     Shape名 nx_cite_<hitIndex> に0始まりのヒット添字を埋め、Peek本文を引く。
'   ・ポップアップは nx_peek。次のPeek表示・送信・閉じるクリックで消える(孤児防止に
'     都度Deleteしてから描く)。柔らかい影は modSkin.ApplySoftShadow を共用。
'   ・LibreOfficeは静的コンパイルのみ(実行しない)。Shapeプロパティは標準VBA。
' ============================================================================

Private Const NEXUS_SHEET As String = "Nexus"
Private Const MAX_CHIPS As Long = 4          ' 出典チップの最大数(横並び)
Private Const CHIP_W As Double = 188
Private Const CHIP_H As Double = 22
Private Const PEEK_W As Double = 470
Private Const PEEK_BODY_MAX As Long = 600    ' ポップアップ本文の最大文字数
' R43 波A 1-2: 見出し/breadcrumbと本文、本文と閉じる案内の間を仕切る区切り行
' (半角ハイフン。太字・色つきの罫線をShapeのテキストで引く手段が無いため、
' 目に付く程度の長さの文字列で代用する)。
Private Const PEEK_DIVIDER As String = "----------------------------------------"

' R46 B-2: 表(Excel)由来の行のセル区切り。Const に ChrW は書けないので
' リテラルで持つ(U+2502。CP932 内なのでビルドの変換で化けない)。
Private Const PEEK_CELL_SEP As String = " │ "
' PeekSpans が「区切りの直後=行頭扱い」を見るときの2字(区切りの末尾)。
Private Const PEEK_CELL_SEP_TAIL As String = "│ "

' PeekBodyText - 吹き出し本文の整形(純関数・R46 B-2)。実機報告「Excelが出典元
'   の場合、なんかずれて表示されて、変な余白があったりと認知負荷が高く見づらい」
'   の対処。3つが重なっていた:
'   (1) タブが生のまま Shape へ届く。modChunker は「2セル以上=タブあり」を表行
'       (ClassifyLine=4)と見て RTrimOnly へ回し、タブを潰す CollapseSpaces は
'       本文行にしか掛けない。Excel は【全行が必ずタブ連結】なので 100% 表行に
'       落ち、唯一のタブ除去点を構造的に迂回する(PDF/Word では起きない)。
'       Office はタブを既定のタブ位置まで送るので、行ごとに値の位置が揃わない。
'       空セルは連続タブ(R33 W3-1 の「列位置を保つ」設計)なので、先に1個へ畳む。
'   (2) 改行の二重化。1行=1シート行の Excel は改行密度が桁違いで、二重化すると
'       本文の半分が空行になる。散文(PDF/Word)では読みやすさに効くので、
'       【表由来のときだけ】やめる。
'   (3) R45 で入れたシート見出し "# シート: 名前" の # が生で見えていた。
'       breadcrumb 側に同じシート名が既に出ているため二重でもある。
'       chapter 側は modChunker.StripHeadingMark が外すが、本文ブロックへは
'       Trim$ しただけで積まれる(modChunker.bas:539)。表示側で落とす。
Public Function PeekBodyText(ByVal bodyOnly As String) As String
    Dim s As String: s = bodyOnly
    Dim isTable As Boolean
    isTable = (InStr(1, s, vbTab, vbBinaryCompare) > 0)

    If Left$(s, 2) = "# " Then s = Mid$(s, 3)

    If isTable Then
        Do While InStr(1, s, vbTab & vbTab, vbBinaryCompare) > 0
            s = Replace(s, vbTab & vbTab, vbTab)
        Loop
        s = Replace(s, vbTab, PEEK_CELL_SEP)
    Else
        s = Replace(s, vbLf, vbLf & vbLf)
    End If
    PeekBodyText = s
End Function

' ChipDocLabel - 出典チップの資料名表示(純関数・R43 波A 1-3)。n字を超えた
'   ときだけ末尾に … (U+2026) を付ける。従来の modUtil.SafeLeft(src,16) は
'   16字ちょうどの資料名も16字未満の資料名も同じ見た目になり、切れたのか
'   どうか利用者に分からなかった(modChrome.ClipToWidthと同じ … の付け方)。
Public Function ChipDocLabel(ByVal s As String, ByVal n As Long) As String
    If Len(s) > n Then
        ChipDocLabel = modUtil.SafeLeft(s, n) & ChrW(&H2026)
    Else
        ChipDocLabel = s
    End If
End Function

' ChipOverflowLabel - MAX_CHIPS件を超えたときだけ「ほかN件」を返す(純関数・
'   R43 波A 1-3)。超えていなければ空文字(=ラベルを出さない)。
'   totalUnique: ユニーク化後の出典総数。shownCount: 実際に描いたチップ数。
Public Function ChipOverflowLabel(ByVal totalUnique As Long, ByVal shownCount As Long) As String
    If totalUnique > shownCount Then
        ChipOverflowLabel = "ほか " & (totalUnique - shownCount) & "件"
    End If
End Function

' ----------------------------------------------------------------------------
' RenderCitations - 直近RAG回答の出典チップを、指定バブルの直下に描画する。
'   bubbleName: 直前に追加したAIバブルのShape名(位置決めの基準)。
' ----------------------------------------------------------------------------
Public Sub RenderCitations(ByVal bubbleName As String)
    On Error GoTo Done
    Dim ws As Worksheet
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub

    HideCitations   ' 前回のチップ/ポップアップを消す(最新回答の下だけに出す)

    ' R33 W5-21 / R33H F7: 回答が成立しなかったターン(API失敗の #ERR・逆質問)
    ' では出典チップも出さない。「AIとの通信に失敗しました」の下に出典が並び、
    ' 押すと実文が開く ―― 根拠は揃っているのに答えが出ないのか、答えが出て
    ' いるのかを取り違える。
    ' F7: 判定は modMode.GroundingAllowed(信頼度バッジ側 modAppAct.DrawConfidence
    ' と同じ門)。W5-21 はこれを mLastMode="" で兼ねていたが、その印は
    ' modUIMain.RenderAnswer の【空質問ターン専用】分岐と衝突していた。
    ' LastConfidenceText の空判定は空質問ターン(mLastMode="")の抑止として残す。
    If Not modMode.GroundingAllowed() Then Exit Sub
    If LenB(modAsk.LastConfidenceText()) = 0 Then Exit Sub

    Dim n As Long: n = modAsk.LastHitCount()
    If n <= 0 Then Exit Sub

    Dim anchor As Shape
    On Error Resume Next
    Set anchor = ws.Shapes(bubbleName)
    On Error GoTo Done
    If anchor Is Nothing Then Exit Sub

    Dim baseL As Double: baseL = anchor.Left
    Dim baseY As Double: baseY = anchor.Top + anchor.Height + 6

    ' 2026-07-27: 積む順を「回答 → 信頼度 → 出典 → 評価」に変えた。
    ' 以前は評価ボタン(nx_act_)の下に出典を置いていたため、利用者は
    ' 542pt分のボタンを越えないと根拠にたどり着けなかった。つまり
    ' 「確かめる前に評価しろ」という並びになっていた。順序は主張なので、
    ' 根拠を先に、上に出す。評価ボタン側がこちらの下端を見て下がる。
    On Error Resume Next
    Dim confBottom As Double
    confBottom = modUINexusDraw.ConfidenceBottom(ws)
    If confBottom + 6 > baseY Then baseY = confBottom + 6
    On Error GoTo Done

    ' 見出しラベル
    Dim lbl As Shape
    Set lbl = ws.Shapes.AddShape(1, baseL, baseY, 260, 16)
    lbl.Name = "nx_cite_lbl"
    lbl.Fill.Visible = 0: lbl.Line.Visible = 0
    With lbl.TextFrame2
        .WordWrap = -1
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDD0E) & " 出典(クリックで原文を確認):"
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 8.5
        .MarginLeft = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    lbl.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
    lbl.Placement = 3

    ' 出典をソース名でユニーク化(先頭出現のヒット添字を保持)しつつチップ描画。
    ' R43 波A 1-3: 5件目以降を Exit For で黙って捨てると、捨てた総数が
    ' 分からず「ほかN件」が出せない。ループは最後まで回して総ユニーク数
    ' (uniqTotal)を数え続け、実際にチップを描くのは MAX_CHIPS 件までにする。
    Dim seen As String: seen = "|"
    Dim x As Double: x = baseL
    Dim y As Double: y = baseY + 18
    Dim drawn As Long: drawn = 0
    Dim uniqTotal As Long: uniqTotal = 0
    Dim i As Long
    For i = 0 To n - 1
        Dim src As String: src = modAsk.LastHitSource(i)
        If LenB(src) = 0 Then GoTo NextHit
        Dim key As String: key = "|" & LCase$(src) & "|"
        If InStr(seen, key) > 0 Then GoTo NextHit   ' 同じ資料は1チップに集約
        seen = seen & LCase$(src) & "|"
        uniqTotal = uniqTotal + 1

        If drawn < MAX_CHIPS Then
            If x + CHIP_W > baseL + 640 Then   ' チャット幅で折り返し
                x = baseL
                y = y + CHIP_H + 6
            End If
            DrawChip ws, i, x, y, src, modAsk.LastHitPage(i)
            x = x + CHIP_W + 8
            drawn = drawn + 1
        End If
NextHit:
    Next i

    Dim overflowCap As String: overflowCap = ChipOverflowLabel(uniqTotal, drawn)
    If LenB(overflowCap) > 0 Then
        If x + 90 > baseL + 640 Then
            x = baseL
            y = y + CHIP_H + 6
        End If
        DrawOverflowChip ws, x, y, overflowCap
    End If

    FreezeAndFront ws
Done:
End Sub

Private Sub DrawChip(ByVal ws As Worksheet, ByVal hitIdx As Long, ByVal x As Double, _
                     ByVal y As Double, ByVal src As String, ByVal page As Long)
    Dim chip As Shape
    Set chip = ws.Shapes.AddShape(5, x, y, CHIP_W, CHIP_H)   ' 5=角丸四角
    chip.Name = "nx_cite_" & CStr(hitIdx)
    chip.Adjustments(1) = 0.5
    chip.Line.Visible = -1
    chip.Line.Weight = 0.75
    chip.Line.ForeColor.RGB = modUI.UiColor("primary")
    chip.Fill.ForeColor.RGB = modUI.UiColor("surface")

    Dim cap As String
    cap = ChrW(&HD83D) & ChrW(&HDCC4) & " " & ChipDocLabel(src, 16)
    cap = cap & modLive.PageLabel(src, page)   ' R40 F3: Excelは「シートN」(0なら空)
    With chip.TextFrame2
        .WordWrap = -1
        .TextRange.Text = cap
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 8.5
        .TextRange.ParagraphFormat.Alignment = 1   ' 左
        .VerticalAnchor = 3
        .MarginLeft = 8: .MarginRight = 6: .MarginTop = 0: .MarginBottom = 0
    End With
    chip.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("primary")
    chip.OnAction = "modApp.OnPeek"
    chip.Placement = 3
End Sub

' DrawOverflowChip - MAX_CHIPS件を超えたぶんの「ほかN件」小ラベル(R43 波A
'   1-3)。チップと違いクリック非対応(OnAction無し)の素のテキストなので、
'   出典チップと誤認されないよう muted・小さめのサイズに留める。
'   名前は "nx_cite_" 始まりにして、HideCitations/FreezeAndFront の
'   既存の一括削除・整列ロジックへそのまま乗せる。
Private Sub DrawOverflowChip(ByVal ws As Worksheet, ByVal x As Double, _
                             ByVal y As Double, ByVal capText As String)
    Dim lbl2 As Shape
    Set lbl2 = ws.Shapes.AddShape(1, x, y, 84, CHIP_H)
    lbl2.Name = "nx_cite_more"
    lbl2.Fill.Visible = 0: lbl2.Line.Visible = 0
    With lbl2.TextFrame2
        .WordWrap = -1
        .TextRange.Text = capText
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 8.5
        .VerticalAnchor = 3
        .MarginLeft = 4: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    lbl2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
    lbl2.Placement = 3
End Sub

' ----------------------------------------------------------------------------
' ShowPeek - 出典チップのクリックで、そのチャンク本文をポップアップ表示する。
'   idx: modAskの0始まりヒット添字(チップ名 nx_cite_<idx> 由来)。
' ----------------------------------------------------------------------------
Public Sub ShowPeek(ByVal idx As Long)
    On Error GoTo Done
    Dim ws As Worksheet
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub

    HidePeek

    Dim body As String: body = modAsk.LastHitPeek(idx)
    If LenB(body) = 0 Then body = "(この出典の本文プレビューは取得できませんでした)"
    Dim src As String: src = modAsk.LastHitSource(idx)
    Dim page As Long: page = modAsk.LastHitPage(idx)

    Dim leftPos As Double: leftPos = 280
    Dim topPos As Double: topPos = 110
    On Error Resume Next
    leftPos = ActiveWindow.VisibleRange.Left + (ActiveWindow.VisibleRange.Width - PEEK_W) / 2
    topPos = ActiveWindow.VisibleRange.Top + 96
    On Error GoTo Done

    Dim head As String
    head = ChrW(&HD83D) & ChrW(&HDCC4) & " " & src
    ' R41 §1 A: Excel は「シートN」(modMode.PageTagPart。表示と出典タグの
    ' 単一情報源)。先頭空白を Trim$ で落として括弧の中へ入れる。
    If page > 0 Then head = head & "  (" & Trim$(modMode.PageTagPart(src, page)) & ")"

    ' R43 波A 1-2: 本文先頭のbreadcrumb行(【…】)を切り出して見出しの直下へ
    ' 別行に置き、見出し/本文/閉じる案内の間に区切り行、本文の行間には空行を
    ' 入れる(modTextView.StripBreadcrumbと同じ「先頭が【かつ改行あり」判定を
    ' 使うので、両者が食い違うことはない)。
    Dim rawBody As String: rawBody = modUtil.SafeLeft(body, PEEK_BODY_MAX)
    Dim crumb As String: crumb = ExtractBreadcrumb(rawBody)
    Dim bodyOnly As String: bodyOnly = PeekBodyText(modTextView.StripBreadcrumb(rawBody))

    Dim assembled As String
    assembled = head
    If LenB(crumb) > 0 Then assembled = assembled & vbLf & crumb
    assembled = assembled & vbLf & PEEK_DIVIDER & vbLf & bodyOnly & vbLf & PEEK_DIVIDER & vbLf & _
                ChrW(&H2715) & " クリックで閉じる"

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, leftPos, topPos, PEEK_W, 60)   ' 高さはAutoSizeで伸ばす
    shp.Name = "nx_peek"
    shp.Adjustments(1) = 0.06
    shp.Line.Visible = -1
    shp.Line.Weight = 1#
    shp.Line.ForeColor.RGB = modUI.UiColor("primary")
    shp.Fill.ForeColor.RGB = modUI.UiColor("surface")
    With shp.TextFrame2
        .WordWrap = -1
        .AutoSize = 1   ' msoAutoSizeShapeToFitText
        .MarginLeft = 14: .MarginRight = 14: .MarginTop = 10: .MarginBottom = 10
        .TextRange.Text = assembled
        .TextRange.Font.Name = "Yu Gothic UI"
        .TextRange.Font.Size = 10
        .TextRange.ParagraphFormat.Alignment = 1
    End With
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("text")

    ' R43 波A 1-2: 色分けは書き込んだ文字列(assembled)ではなく、Shapeへ
    ' 代入した後に読み戻したテキストで位置を数える(vbLfが読み戻しでvbCrに
    ' なりうるため。modLiveStyle.StyleAnswerParasと同じ作法)。
    ' 必ず床登録(SetOverlayFloor)・「原文を開く」ボタン配置より前に済ませる
    ' (AutoSizeでここから高さが変わるため)。
    On Error Resume Next
    Dim rt As String: rt = shp.TextFrame2.TextRange.Text
    Dim hS As Long, hL As Long, cS As Long, cL As Long, clS As Long, clL As Long
    Dim rStarts() As Long, rLens() As Long
    Dim nRef As Long: nRef = PeekSpans(rt, hS, hL, cS, cL, rStarts, rLens, clS, clL)
    If hL > 0 Then
        With shp.TextFrame2.TextRange.Characters(hS, hL).Font
            .Size = 11
            .Bold = True
            .Fill.ForeColor.RGB = modUI.UiColor("primary")
        End With
    End If
    If cL > 0 Then
        With shp.TextFrame2.TextRange.Characters(cS, cL).Font
            .Size = 8.5
            .Fill.ForeColor.RGB = modUI.UiColor("muted")
        End With
    End If
    Dim ri As Long
    For ri = 1 To nRef
        shp.TextFrame2.TextRange.Characters(rStarts(ri), rLens(ri)).Font.Fill.ForeColor.RGB = modUI.UiColor("primary")
    Next ri
    If clL > 0 Then
        ' R46: 実機報告「×で閉じるが見にくい」。原因はコントラスト不足ではない
        ' (muted は全テーマで AA を満たす)。8.5pt へ落として灰にしてある指定
        ' そのものと、同じ帯の右下に primary 塗りの「原文を開く」ボタンが並ぶ
        ' 構図で、目が右下へ行くこと。本文と同じ10ptへ戻し太字＋danger にする。
        ' danger はテーマごとに明系/暗系を出し分けるので、赤のRGB直書きと違い
        ' 全テーマで AA を満たす(直書きだと dark で 2.26:1 まで落ちる)。
        With shp.TextFrame2.TextRange.Characters(clS, clL).Font
            .Size = 10
            .Bold = True
            .Fill.ForeColor.RGB = modUI.UiColor("danger")
        End With
    End If
    On Error GoTo Done

    shp.OnAction = "modApp.OnPeekClose"
    shp.Placement = 3
    modSkin.ApplySoftShadow shp
    shp.ZOrder 0   ' msoBringToFront

    ' R27 F2-1b(実機第12報①): プレビューは modUI.RecalcChatBottom の許可
    ' リスト外(重ね表示)なので境界関所を通らない。AutoSizeで伸びた実下端で
    ' 関所を通す。「原文を開く」ボタンはこの吹き出しの内側に置くので、
    ' ここが常に最深(GoTo Doneで飛んだ経路でも取りこぼさない)。
    ' R27H F2(M-1裁定): 開いている間の下端を床として登録する(開いたまま
    ' 質問送信/トグルが走ると、別経路が会話の下端で境界を貼り直すため)。
    modSkin.SetOverlayFloor shp.Top + shp.Height + 12
    modSkin.ExtendChatBand ws, shp.Top + shp.Height + 12

    ' 「📂 原文を開く」。抜粋を読んで終わりではなく、実物の該当ページまで
    ' 連れて行く。ここまで来て初めて、人に見せられる根拠になる。
    ' 元ファイルの記録が無い資料(パック由来・手入力)には出さない。
    If LenB(SourcePath(src)) = 0 Then GoTo Done

    Dim openCap As String
    openCap = ChrW(&HD83D) & ChrW(&HDCC2) & " 原文を開く"
    If page > 0 Then openCap = openCap & "（" & Trim$(modMode.PageTagPart(src, page)) & "）"

    Dim btn As Shape
    On Error Resume Next
    Set btn = ws.Shapes.AddShape(5, shp.Left + PEEK_W - 158, shp.Top + shp.Height - 34, 142, 26)
    If Err.Number = 0 And Not btn Is Nothing Then
        btn.Name = "nx_peek_open_" & CStr(idx)
        btn.Adjustments(1) = 0.4
        btn.Line.Visible = 0
        btn.Fill.ForeColor.RGB = modUI.UiColor("primary")
        With btn.TextFrame2
            .WordWrap = -1
            .TextRange.Text = openCap
            .TextRange.Font.Name = "Yu Gothic UI"
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = -1
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 3
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
        End With
        btn.OnAction = "modPeek.OnOpenSource"
        btn.Placement = 3
        btn.ZOrder 0
    End If
    Set btn = Nothing
    Err.Clear
    On Error GoTo Done
Done:
End Sub

' ----------------------------------------------------------------------------
' PeekSpans - プレビュー本文(assembled、ShowPeekが組み立てた1本の文字列)の
'   中から「見出し／breadcrumb／行頭の[A6]・[シート:…]タグ／末尾の閉じる
'   案内」の位置と長さを返す(純関数・R43 波A 1-2)。改行は vbLf/vbCr の
'   どちらでも判定できるようにする(Shapeからの読み戻しはvbCr、テストからの
'   直接呼び出しはvbLfで組み、両方をここで吸収する)。
'
'   構造の前提(ShowPeekが必ずこの形で組む):
'     1行目            = 見出し(📄 資料名 (p.N) 等)
'     2行目(あれば)    = breadcrumb(先頭が「【」の行)
'     以降             = 区切り行・本文(行頭に[A6]等が来ることがある)・区切り行
'     最後の改行より後 = 閉じる案内(✕ クリックで閉じる)
'
'   見出し=1行目まるごと(最初の改行の直前まで)。breadcrumbは見出し直後の
'   行が「【」で始まるときだけそれを丸ごと。閉じる案内=文字列全体の最後の
'   改行より後ろ全部(区切り行を挟んでいても、最後の改行の直後から末尾までは
'   常に閉じる案内になるようShowPeek側が保証する)。行頭の"[...]"タグは
'   本文中に何個あってもよい(refStarts/refLensで列挙。戻り値はその件数)。
' ----------------------------------------------------------------------------
Public Function PeekSpans(ByVal s As String, _
        ByRef headStart As Long, ByRef headLen As Long, _
        ByRef crumbStart As Long, ByRef crumbLen As Long, _
        ByRef refStarts() As Long, ByRef refLens() As Long, _
        ByRef closeStart As Long, ByRef closeLen As Long) As Long
    headStart = 0: headLen = 0
    crumbStart = 0: crumbLen = 0
    closeStart = 0: closeLen = 0
    Dim nRef As Long
    ReDim refStarts(1 To 1): ReDim refLens(1 To 1)

    Dim total As Long: total = Len(s)
    If total = 0 Then
        PeekSpans = 0
        Exit Function
    End If

    ' 見出し行: 先頭から最初の改行の直前まで(改行が無ければ全体)。
    Dim firstBreak As Long: firstBreak = PeekNextBreak(s, 1)
    headStart = 1
    If firstBreak = 0 Then
        headLen = total
    Else
        headLen = firstBreak - 1
    End If

    ' breadcrumb行: 見出しの直後の行が「【」で始まればそれを丸ごと。
    If firstBreak > 0 And firstBreak < total Then
        Dim afterHead As Long: afterHead = firstBreak + 1
        If Mid$(s, afterHead, 1) = "【" Then
            Dim crumbBreak As Long: crumbBreak = PeekNextBreak(s, afterHead)
            crumbStart = afterHead
            If crumbBreak = 0 Then
                crumbLen = total - afterHead + 1
            Else
                crumbLen = crumbBreak - afterHead
            End If
        End If
    End If

    ' 閉じる案内行: 文字列全体の最後の改行より後ろ。改行が1つも無ければ0の
    ' まま(見出し行しか無い異常系。ShowPeekは必ず区切り行+閉じる案内を
    ' 付けるので実運用では到達しない)。
    Dim lastBreak As Long: lastBreak = PeekLastBreak(s)
    If lastBreak > 0 And lastBreak < total Then
        closeStart = lastBreak + 1
        closeLen = total - closeStart + 1
    End If

    ' 行頭の "[...]" タグ(セル番地・シート表記)。行頭=位置1、または直前が
    ' 改行の位置。見出し/breadcrumb/区切り行/閉じる案内はいずれも "[" で
    ' 始まらないため、範囲を区切らず全体を1回走査すればよい。
    Dim pos As Long: pos = 1
    Do While pos <= total
        Dim atLineStart As Boolean: atLineStart = (pos = 1)
        If Not atLineStart Then atLineStart = PeekIsBreak(Mid$(s, pos - 1, 1))
        ' R46 B-2: セル区切りの直後も行頭として扱う。R43-D で番地が【値の直前】
        ' へ移った結果、1行の2列目以降の [B1] は区切りの後ろに来て装飾から
        ' 漏れ、同じ種類の印が行の中で色違いになっていた。境界の判定と
        ' Mid$ の参照は分けて書く(And は短絡しない)。
        If Not atLineStart Then
            If pos >= 3 Then atLineStart = (Mid$(s, pos - 2, 2) = PEEK_CELL_SEP_TAIL)
        End If
        Dim matched As Boolean: matched = False
        If atLineStart And Mid$(s, pos, 1) = "[" Then
            Dim closeBr As Long: closeBr = InStr(pos, s, "]")
            Dim brk As Long: brk = PeekNextBreak(s, pos)
            If closeBr > 0 And (brk = 0 Or closeBr < brk) Then
                nRef = nRef + 1
                ReDim Preserve refStarts(1 To nRef): ReDim Preserve refLens(1 To nRef)
                refStarts(nRef) = pos
                refLens(nRef) = closeBr - pos + 1
                pos = closeBr + 1
                matched = True
            End If
        End If
        If Not matched Then pos = pos + 1
    Loop

    PeekSpans = nRef
End Function

Private Function PeekIsBreak(ByVal c As String) As Boolean
    PeekIsBreak = (c = vbLf Or c = vbCr)
End Function

Private Function PeekNextBreak(ByVal s As String, ByVal fromPos As Long) As Long
    Dim i As Long
    For i = fromPos To Len(s)
        If PeekIsBreak(Mid$(s, i, 1)) Then
            PeekNextBreak = i
            Exit Function
        End If
    Next i
End Function

Private Function PeekLastBreak(ByVal s As String) As Long
    Dim i As Long
    For i = Len(s) To 1 Step -1
        If PeekIsBreak(Mid$(s, i, 1)) Then
            PeekLastBreak = i
            Exit Function
        End If
    Next i
End Function

' ExtractBreadcrumb - 先頭の【…】breadcrumb行だけを取り出す(純関数・R43
'   波A 1-2)。modTextView.StripBreadcrumb(残す側を返す関数)と全く同じ条件
'   (先頭が「【」かつ改行がある)で判定するので、2つの関数の判定が食い違う
'   ことは無い(「【だけ」のような不完全な行は、どちらの関数も手を付けない)。
Private Function ExtractBreadcrumb(ByVal s As String) As String
    If Left$(s, 1) = "【" Then
        Dim lfPos As Long: lfPos = InStr(s, vbLf)
        If lfPos > 0 Then ExtractBreadcrumb = Left$(s, lfPos - 1)
    End If
End Function

Public Sub HidePeek()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    ws.Shapes("nx_peek").Delete
    ' 「原文を開く」ボタンはポップアップとは別Shapeなので、道連れにしないと
    ' 本文だけ消えてボタンだけが宙に浮いて残る。
    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 13) = "nx_peek_open_" Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        ws.Shapes(names(i)).Delete
    Next i
    ' R27 F2-1b: 伸ばした境界を会話の下端へ戻す(戻さないと閉じたあとも
    ' 境界だけが下に残り、白い余白として見える)。
    ' R27H F2: 床を先に下ろす(残っていると縮む方向が効かない)。
    modSkin.ClearOverlayFloor
    modSkin.ExtendChatBand ws, modUI.ChatBottomFor(ws)
    On Error GoTo 0
End Sub

' 出典チップとポップアップを一括削除(次の送信の先頭・画面リセット時に呼ぶ)。
Public Sub HideCitations()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) = "nx_cite_" Or shp.Name = "nx_peek" _
           Or Left$(shp.Name, 13) = "nx_peek_open_" Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        ws.Shapes(names(i)).Delete
    Next i
    ' R27H F2追加: nx_peekはここでも消えるため、ShowPeekが立てた床もここで下ろす
    ' (HidePeekと同じ理由。ClearOverlayFloorは冪等)。
    modSkin.ClearOverlayFloor
    On Error GoTo 0
End Sub

' 出典チップ群の下端(無ければ0)。評価ボタンがこの下へ回り込むために使う。
Public Function CitationsBottom(ByVal ws As Worksheet) As Double
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) = "nx_cite_" Then
            If shp.Top + shp.Height > CitationsBottom Then CitationsBottom = shp.Top + shp.Height
        End If
    Next shp
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' OnOpenSource - ポップアップの「📂 原文を開く」。
' ----------------------------------------------------------------------------
' ここが無いせいで、このアプリは「retrievalのデモ」で止まっていた。
' 損保の実務で意味を持つ瞬間は、上司やお客さまに約款の該当ページそのものを
' 見せるとき。灰色のテキストボックスで終わっていては、その場に持っていけない。
' file_pathは my_manifest に既にある。開くだけでよかった。
Public Sub OnOpenSource()
    ' R15波1b裁定: 外部アプリで原文を開く際、取込中は同じファイルを掴んでいる
    ' 可能性がある(既存40箇所超と同型。lint検査13で現状追認だったものを
    ' 保護へ切り替え)。
    If modUiLock.BlockIfIngesting() Then Exit Sub
    On Error Resume Next
    Dim caller As String
    caller = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(caller, 13) <> "nx_peek_open_" Then
        modLog.LogUsage "caller_mismatch", "modPeek.OnOpenSource", caller
        Exit Sub
    End If

    Dim idx As Long
    idx = CLng(Val(Mid$(caller, 14)))

    Dim src As String, page As Long
    On Error Resume Next
    src = modAsk.LastHitSource(idx)
    page = modAsk.LastHitPage(idx)
    On Error GoTo 0
    If LenB(src) = 0 Then Exit Sub

    Dim path As String
    path = SourcePath(src)

    If LenB(path) = 0 Then
        MsgBox "この資料の元ファイルの場所が記録されていません。" & vbCrLf & _
               "(パックで受け取った資料や、手入力で登録した内容には元ファイルがありません)", _
               vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' PDFはページ指定で開く。対応しないビューアでは先頭ページで開くだけで、
    ' 失敗はしない。開けなかったときはパスを見せる(手で辿れるようにする)。
    Dim target As String: target = path
    If page > 0 Then
        If LCase$(modUtil.ExtOf(path)) = "pdf" Then target = path & "#page=" & page
    End If

    Dim ok As Boolean: ok = True
    On Error Resume Next
    ' 2026-08-10(R27波3-14): waitless:=True。この一言は【この直後に出る
    ' Officeの確認画面】の予告であって、読ませるために待たせる文ではない。
    ' 既定のままだと 1.1 秒ぶん確認画面の到着が遅れ、その待ちの DoEvents 中に
    ' 別のクリックが割り込む窓も開く。描いたらすぐ開きにいく。
    modSkin.ShowToast "Officeの確認画面が出たら[はい]を押してください。", "info", True
    ThisWorkbook.FollowHyperlink target
    If Err.Number <> 0 Then ok = False
    Err.Clear
    On Error GoTo 0

    If Not ok Then
        On Error Resume Next
        ' R33 W4-7: コピーの成否を見ずに「コピーしました」と言わない。
        ' 失敗時は CopyOrGuide が手で控えるための案内へ差し替える。
        Dim clipMsg As String
        clipMsg = modClip.CopyOrGuide(path, "場所をクリップボードにコピーしました:" & vbCrLf & path)
        On Error GoTo 0
        MsgBox "元ファイルを開けませんでした(移動または削除された可能性があります)。" & vbCrLf & _
               clipMsg, vbExclamation, modAppDef.APP_NAME
    End If
End Sub

' ----------------------------------------------------------------------------
' SourcePath - 資料名(= my_manifest の file_name)から元ファイルの実パスを引く。
'   見つからない/実体が消えている場合は空文字を返し、呼び出し側はボタンを出さない。
' ----------------------------------------------------------------------------
' パスは取込時から manifest 1列目に入っていたのに、そこへ到達する道が
' どこにも無かった。パック由来・手入力のナレッジは manifest に行が無いので空。
Private Function SourcePath(ByVal sourceName As String) As String
    Dim nm As String: nm = Trim$(sourceName)
    If LenB(nm) = 0 Then Exit Function

    On Error Resume Next
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(modAppDef.SH_MANIFEST)
    On Error GoTo 0
    If wsM Is Nothing Then Exit Function

    Dim lastR As Long
    On Error Resume Next
    lastR = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    On Error GoTo 0
    If lastR < 2 Then Exit Function

    Dim r As Long
    On Error Resume Next
    For r = 2 To lastR
        If StrComp(Trim$(CStr(wsM.Cells(r, 2).Value)), nm, vbTextCompare) = 0 Then
            Dim p As String: p = Trim$(CStr(wsM.Cells(r, 1).Value))
            ' 消えたファイルのパスを返すと「開けません」で終わる。存在確認まで
            ' 済ませてから返し、無ければボタン自体を出さない。
            If LenB(p) > 0 Then
                If LenB(Dir(p)) > 0 Then SourcePath = p
            End If
            Exit For
        End If
    Next r
    On Error GoTo 0
End Function

Private Function GetSheet() As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(NEXUS_SHEET)
    On Error GoTo 0
End Function

Private Sub FreezeAndFront(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) = "nx_cite_" Then shp.Placement = 3
    Next shp
    On Error GoTo 0
End Sub
