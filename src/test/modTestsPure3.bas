Attribute VB_Name = "modTestsPure3"
Option Explicit

' ============================================================================
' modTestsPure3 - modTestsPure/modTestsPure2の分割先(2026-07-30 R2要件B追加)
' ----------------------------------------------------------------------------
' 役割:
'   modTestsPure(28,906字)・modTestsPure2(29,170字)とも§7.1の
'   「1モジュール30,000字以内」まで残りが少なく、要件Bで追加する回帰テストの
'   置き場が無かったための新規分割先。modExtractor.DropGarbledPagesの配列
'   圧縮(要件A)・modChunkerの空ページ多層防御(要件B)・modExtractorの
'   共有読みローカルコピーの純ロジック部分(要件C)を検証する。
'   入口は modTestsPure2.RunAll2 の末尾から呼ばれる Public Sub RunAll3()。
'   modTestRunner.RunAllPureTests は modTestsPure.RunAll だけを呼ぶ契約
'   (modTestRunner.bas §7.8。担当外につき変更しない)なので、本モジュールへの
'   導線は modTestsPure2.RunAll2 内に置く(modTestsPureがmodTestsPure2を
'   呼ぶのと同じ中継方式)。
'
' 設計判断(R4準拠・グループ単位の失敗隔離): modTestsPure/modTestsPure2側の
'   冒頭コメントと同じ方針(1グループの想定外エラーが他グループを道連れに
'   しない)。
'
' ■ CanUseTypeArraysの複製について(modTestsPure2冒頭コメントと同じ理由):
'   modTestsPureのPrivate Function CanUseTypeArraysは別モジュールのPrivateで
'   あるため、ここから呼べない。3行程度の軽量な実測プローブなので、モジュール
'   をまたいだ複製を許容する。挙動・コメントはmodTestsPure側のオリジナルと
'   同一にしてある。ExtractedPage()/ShelfChunk()を扱うテストグループだけを
'   これで守る(SharedCopyNextChunkLenはLong引数のみの純関数なので対象外)。
' ============================================================================

Private Function CanUseTypeArrays() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As ShelfChunk
    ReDim probe(0 To 0)
    CanUseTypeArrays = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' 要件B: 空ページの多層防御(2026-07-30)
' ----------------------------------------------------------------------------
' 背景: 化けページを配列から取り除く実装(要件A)以前は、DropGarbledPagesが
' 該当ページの.Textを""にするだけで配列を詰めておらず、空文字ページが後段の
' どこかでerr#9を出し、ページ単位スキップに化けていた(実機ログ:
' extract_garbled N件 と chunk_skipped_pages N件が全6ペアで完全一致)。
' 1ページしかない.docだと全文が消えてtotalChars=0→E0304で取込不能になって
' いた。NormalizeForIngest単体、および空ページ混在配列を渡したChunkPagesEx
' の両方でエラーが出ないこと・SkippedPageCountに数えられないことを固定する。
Private Sub TestNormalizeForIngestEmpty()
    Dim gotErr As Boolean: gotErr = False
    Dim errInfo As String: errInfo = ""
    Dim normResult As String
    On Error GoTo NormEmptyErr
    normResult = modChunker.NormalizeForIngest("")
    GoTo NormEmptyOk
NormEmptyErr:
    gotErr = True
    errInfo = "err#" & Err.Number & ": " & Err.Description
    Resume NormEmptyOk
NormEmptyOk:
    On Error GoTo 0
    modTestRunner.Check "NormalizeForIngest: 空文字でエラーを出さない", Not gotErr, errInfo
    modTestRunner.Check "NormalizeForIngest: 空文字は空文字のまま", normResult = "", "実際=" & normResult
End Sub

Private Sub TestChunkerEmptyPageDefense()
    If Not CanUseTypeArrays() Then
        modTestRunner.Check "modChunker(空ページ防御): LO環境の既知の制限によりスキップ", True, _
            "LibreOffice実行環境ではmodTypesのPublic Type配列(ExtractedPage()/" & _
            "ShelfChunk())をReDimすると実行時エラー420になることを確認済み" & _
            "(modTestsPure冒頭コメント参照。tools/run_lo_tests.pyは担当外のため変更不可)。" & _
            "Excel実機受入チェック(§11.3)で必ず再確認すること。"
        Exit Sub
    End If

    ' Text=""のページを含む配列でChunkPagesExがerrなく動き、
    ' SkippedPageCount()=0(空ページは「読めたのに落ちた」件数に数えない)。
    Dim pages2(0 To 2) As ExtractedPage
    pages2(0).page = 1: pages2(0).Text = "正常な本文です。十分な長さの文章を入れておきます。"
    pages2(1).page = 2: pages2(1).Text = ""
    pages2(2).page = 3: pages2(2).Text = "もう1件の正常な本文です。チャンク化の対象になります。"

    Dim chunksLegacy() As ShelfChunk
    Dim nLegacy As Long
    nLegacy = modChunker.ChunkPagesEx(pages2, 700, 150, 0, "legacy", chunksLegacy)
    modTestRunner.Check "ChunkPagesEx(legacy): 空ページ混在でもエラーにならない", (nLegacy >= 1), "n=" & nLegacy
    modTestRunner.Check "ChunkPagesEx(legacy): 空ページはSkippedPageCountに数えない", _
        modChunker.SkippedPageCount() = 0, "count=" & modChunker.SkippedPageCount()

    Dim chunksStruct() As ShelfChunk
    Dim nStruct As Long
    nStruct = modChunker.ChunkPagesEx(pages2, 700, 150, 1800, "structure", chunksStruct)
    modTestRunner.Check "ChunkPagesEx(structure): 空ページ混在でもエラーにならない", (nStruct >= 1), "n=" & nStruct
    modTestRunner.Check "ChunkPagesEx(structure): 空ページはSkippedPageCountに数えない", _
        modChunker.SkippedPageCount() = 0, "count=" & modChunker.SkippedPageCount()
End Sub

' ----------------------------------------------------------------------------
' 要件A: 化けページ圧縮後の配列でも総チャンク数が正しい(2026-07-30)
' ----------------------------------------------------------------------------
' DropGarbledPagesは化けページを取り除いて配列を詰めるため、残るページの
' .pageには欠番(このテストでは2)が生じる。欠番があってもチャンク数・
' 各チャンクの.page引継ぎが崩れないことを固定する(DropGarbledPages自体は
' modExtractorのPrivateだが、その出力に相当する「詰め済み配列」を直接
' 組み立ててmodChunker側の挙動を検証する)。
Private Sub TestChunkerCompactedPageArray()
    If Not CanUseTypeArrays() Then
        modTestRunner.Check "modChunker(圧縮配列): LO環境の既知の制限によりスキップ", True, _
            "LibreOffice実行環境ではmodTypesのPublic Type配列(ExtractedPage()/" & _
            "ShelfChunk())をReDimすると実行時エラー420になることを確認済み" & _
            "(modTestsPure冒頭コメント参照。tools/run_lo_tests.pyは担当外のため変更不可)。" & _
            "Excel実機受入チェック(§11.3)で必ず再確認すること。"
        Exit Sub
    End If

    Dim pages3(0 To 1) As ExtractedPage
    pages3(0).page = 1: pages3(0).Text = "第1条(目的)本規程は業務の適正な遂行を目的とする。以下略。"
    pages3(1).page = 3: pages3(1).Text = "第2条(適用範囲)本規程は全社員に適用する。以下略。"

    Dim chunks3() As ShelfChunk
    Dim n3 As Long
    n3 = modChunker.ChunkPages(pages3, 700, 150, chunks3)
    modTestRunner.Check "ChunkPages: 圧縮済み配列(ページ番号に欠番)でも正しいチャンク数", (n3 = 2), "n=" & n3
    If n3 = 2 Then
        modTestRunner.Check "ChunkPages: 圧縮後も元のページ番号を保持_1件目", chunks3(0).page = 1, _
            "page=" & chunks3(0).page
        modTestRunner.Check "ChunkPages: 圧縮後も元のページ番号を保持_2件目", chunks3(1).page = 3, _
            "page=" & chunks3(1).page
    End If
End Sub

' ----------------------------------------------------------------------------
' 要件C: 共有読みコピーの純ロジック部分(バッファ組み立ての境界計算)
' ----------------------------------------------------------------------------
' CopySharedRead自体は実ファイルI/Oのため純ロジックテストの対象外だが、
' 「pos(1始まり)から次に読むべきバイト数」を決めるSharedCopyNextChunkLenは
' 切り出し済みの純関数なので、端数処理・0バイト・pos超過の境界を固定する。
Private Sub TestSharedCopyNextChunkLen()
    modTestRunner.Check "共有読みコピー: ちょうど割り切れる場合はchunkBytes全部", _
        modExtractor.SharedCopyNextChunkLen(1, 2097152, 1048576) = 1048576, _
        "実際=" & modExtractor.SharedCopyNextChunkLen(1, 2097152, 1048576)
    modTestRunner.Check "共有読みコピー: 端数は残りぶんだけ", _
        modExtractor.SharedCopyNextChunkLen(1048577, 1500000, 1048576) = 451424, _
        "実際=" & modExtractor.SharedCopyNextChunkLen(1048577, 1500000, 1048576)
    modTestRunner.Check "共有読みコピー: 0バイトファイルは0を返す", _
        modExtractor.SharedCopyNextChunkLen(1, 0, 1048576) = 0, ""
    modTestRunner.Check "共有読みコピー: posがtotalLenを超えたら0", _
        modExtractor.SharedCopyNextChunkLen(500, 100, 1048576) = 0, ""
    modTestRunner.Check "共有読みコピー: chunkBytesが1未満なら0", _
        modExtractor.SharedCopyNextChunkLen(1, 100, 0) = 0, ""
    modTestRunner.Check "共有読みコピー: 総量ちょうどでも末尾を正しく返す", _
        modExtractor.SharedCopyNextChunkLen(100, 100, 1048576) = 1, ""
End Sub

' ----------------------------------------------------------------------------
' R3要件C(2026-07-30): welcomeバッジがBadgeCatalog(単一情報源)に含まれ、
' 件数が期待どおりであることを固定する(受入条件に明記された検証項目)。
' ----------------------------------------------------------------------------
Private Sub TestBadgeCatalogWelcome()
    Dim ids() As String, titles() As String, shorts() As String, conds() As String
    Dim n As Long
    n = modStats.BadgeCatalog(ids, titles, shorts, conds)

    ' 従来12種+welcome=13種。件数がズレたらここで気づける。
    modTestRunner.Check "R3バッジ表_件数は13", (n = 13), "n=" & n

    Dim found As Boolean, i As Long
    For i = LBound(ids) To UBound(ids)
        If StrComp(Trim$(ids(i)), "welcome", vbTextCompare) = 0 Then found = True
    Next i
    modTestRunner.Check "R3バッジ表_welcomeを含む", found

    ' EvaluateBadges側もwelcomeを先頭で判定する設計(要件C)。表の並びも
    ' 揃えてあることを固定する(必須ではないが、崩れたら気づけるように)。
    modTestRunner.Check "R3バッジ表_welcomeは先頭", _
        (n > 0) And (StrComp(Trim$(ids(LBound(ids))), "welcome", vbTextCompare) = 0)
End Sub

' ----------------------------------------------------------------------------
' R3要件D(2026-07-30): Hubタイルの既定値ロジック(空文字を絶対に返さない)。
' 実体はmodHubStat.DefaultTileValue(UI層・Shape/Worksheet依存の濃い
' modHub/modHubStatはPURE_ALLOWLISTへ注入できない)にあるため、ここでは
' 同じ選択規則を複製して固定する(modTestsPure/2でのCanUseTypeArrays複製と
' 同じ考え方。modHubStat.DefaultTileValueの実装を変えたらこちらも
' 合わせて直すこと)。
' ----------------------------------------------------------------------------
Private Function HubTileDefaultTextForTest(ByVal idx As Long) As String
    Select Case idx
        Case 2, 4, 5   ' 節約できた時間 / みんな(今日) / みんな(今月)
            HubTileDefaultTextForTest = "0分"
        Case 6         ' 連続ログイン
            HubTileDefaultTextForTest = "0日"
        Case 3         ' 本棚の使用量
            HubTileDefaultTextForTest = "0%"
        Case Else      ' 質問した回数 / 自己解決 / パック共有
            HubTileDefaultTextForTest = "0"
    End Select
End Function

Private Sub TestHubTileDefaultValue()
    modTestRunner.Check "R3タイル既定値_質問した回数は0", HubTileDefaultTextForTest(0) = "0"
    modTestRunner.Check "R3タイル既定値_自己解決は0", HubTileDefaultTextForTest(1) = "0"
    modTestRunner.Check "R3タイル既定値_節約できた時間は0分", HubTileDefaultTextForTest(2) = "0分"
    modTestRunner.Check "R3タイル既定値_本棚の使用量は0%", HubTileDefaultTextForTest(3) = "0%"
    modTestRunner.Check "R3タイル既定値_みんな今日は0分", HubTileDefaultTextForTest(4) = "0分"
    modTestRunner.Check "R3タイル既定値_みんな今月は0分", HubTileDefaultTextForTest(5) = "0分"
    modTestRunner.Check "R3タイル既定値_連続ログインは0日", HubTileDefaultTextForTest(6) = "0日"
    modTestRunner.Check "R3タイル既定値_パック共有は0", HubTileDefaultTextForTest(7) = "0"

    ' どのindexでも空文字を返さない(仕様上の絶対条件そのもの)。
    Dim i As Long, emptyN As Long
    For i = 0 To 7
        If LenB(HubTileDefaultTextForTest(i)) = 0 Then emptyN = emptyN + 1
    Next i
    modTestRunner.Check "R3タイル既定値_どのタイルも空文字を返さない", (emptyN = 0), "empty=" & emptyN
End Sub

' ----------------------------------------------------------------------------
' R4要件C: ツールバーの完全折り返し(2026-07-30)
' ----------------------------------------------------------------------------
' 背景: 折り返し判定が `rowIdx = 0` のときだけ発火していたため、2段目に
' 入った時点で幅を見なくなり、あふれたボタンは画面外へ描かれていた
' (実機写真「削除ボタンが『除』しか見えない」)。modChrome.FlowLeftは段数を
' 無制限にした流し込みで、「どの要素も帯の右端を超えない」ことを保証する。
' ここでは3段以上に折れるケースを作って、その保証を固定する。
Private Sub TestChromeFlowLeft()
    Dim widths() As Double
    Dim xs() As Double, rws() As Long, useW() As Double
    Dim hugeW() As Double
    Dim xs2() As Double, rws2() As Long, useW2() As Double
    Dim n As Long, i As Long, rowN As Long, rowN2 As Long
    Dim overflow As Long

    n = 11
    ReDim widths(0 To n - 1)
    For i = 0 To n - 1
        widths(i) = 100
    Next i

    rowN = modChrome.FlowLeft(widths, n, 0, 350, 5, xs, rws, useW)

    ' 100pt+間隔5pt なら1段に3個(x=0/105/210)。11個なら4段になる。
    modTestRunner.Check "FlowLeft: 11個100ptが幅350へ4段で収まる", rowN = 4, _
        "rowN=" & rowN
    modTestRunner.Check "FlowLeft: 4個目は2段目の左端へ戻る", _
        rws(3) = 1 And xs(3) = 0, "row=" & rws(3) & " x=" & xs(3)

    overflow = 0
    For i = 0 To n - 1
        If xs(i) + useW(i) > 350 Then overflow = overflow + 1
        If xs(i) < 0 Then overflow = overflow + 1
    Next i
    modTestRunner.Check "FlowLeft: 帯からはみ出す要素が1つも無い", overflow = 0, _
        "はみ出し=" & overflow & "件"

    ' 1個で帯幅を超える要素は帯幅へ丸める(丸めないと何段折ってもはみ出す)。
    ReDim hugeW(0 To 0)
    hugeW(0) = 1000
    rowN2 = modChrome.FlowLeft(hugeW, 1, 0, 350, 5, xs2, rws2, useW2)
    modTestRunner.Check "FlowLeft: 帯より広い1個は帯幅へ丸める", _
        rowN2 = 1 And useW2(0) = 350 And xs2(0) = 0, _
        "rowN=" & rowN2 & " w=" & useW2(0) & " x=" & xs2(0)
End Sub

' ----------------------------------------------------------------------------
' R4要件D: チャットヘッダーの幅予算(2026-07-30)
' ----------------------------------------------------------------------------
' 背景: 右端ピル8個の固定予約幅の合計598pt(+右余白8pt)が、ヘッダー実幅
' (約597pt)を一度も突き合わせられないまま右から積まれ、最内側のピルが必ず
' タイトルへ重なっていた。modChrome.FlowRightは「その段の左限界より左へは
' 絶対に置かない/段幅を超える要素は段幅へ丸める」の2点で、帯幅Wがどんな値でも
' タイトル領域へ入り込む配置を作れないようにしている。
' Wを極端な値まで振って、その不変条件が崩れないことを固定する。
Private Sub TestChromeHeaderBudget()
    Dim widths() As Double
    Dim xs() As Double, rws() As Long, useW() As Double
    Dim cases_ As Variant
    Dim n As Long, i As Long, c As Long, rowN As Long, bad As Long
    Dim barW As Double, titleMin As Double, limitX As Double
    Dim detail As String

    n = 8
    ReDim widths(0 To n - 1)
    widths(0) = 30: widths(1) = 62: widths(2) = 74: widths(3) = 30
    widths(4) = 26: widths(5) = 92: widths(6) = 118: widths(7) = 124

    ' 旧実装の固定予約幅の合計(間隔込み)は598pt。実幅597ptの帯には入らない。
    modTestRunner.Check "SumSpan: 旧予約幅の合計は598pt(実幅597ptを超える)", _
        modChrome.SumSpan(widths, n, 6) = 598, _
        "sum=" & modChrome.SumSpan(widths, n, 6)

    cases_ = Array(120, 200, 300, 400, 597, 640, 900, 1400)
    bad = 0
    detail = ""
    For c = LBound(cases_) To UBound(cases_)
        barW = CDbl(cases_(c))
        ' 確保幅は帯幅の半分が上限(TitleReserve)。極端に狭い帯でも
        ' 「確保しすぎてピルが1個も置けない」状態にはしない。
        titleMin = modChrome.TitleReserve(barW, 280)
        rowN = modChrome.FlowRight(widths, n, barW - 8, titleMin, 8, 6, xs, rws, useW)
        If rowN < 1 Then bad = bad + 1
        For i = 0 To n - 1
            If rws(i) = 0 Then
                limitX = titleMin
            Else
                limitX = 8
            End If
            If xs(i) < limitX Then
                bad = bad + 1
                detail = detail & " W=" & barW & "/i=" & i & "/x=" & xs(i)
            End If
            If xs(i) + useW(i) > barW - 8 Then
                bad = bad + 1
                detail = detail & " W=" & barW & "/i=" & i & "/右端超過"
            End If
        Next i
    Next c
    modTestRunner.Check "FlowRight: どの帯幅でもタイトル領域へ食い込まない", bad = 0, _
        "違反=" & bad & "件" & detail

    ' 広い帯なら1段に収まり、実機幅では旧予約幅のままだと段が増える
    ' (=旧実装が1段に詰め込んでいたのは物理的に不可能だった、という確認)。
    rowN = modChrome.FlowRight(widths, n, 1400 - 8, 280, 8, 6, xs, rws, useW)
    modTestRunner.Check "FlowRight: 1400pt幅なら1段で収まる", rowN = 1, "rowN=" & rowN
    rowN = modChrome.FlowRight(widths, n, 597 - 8, 280, 8, 6, xs, rws, useW)
    modTestRunner.Check "FlowRight: 597pt幅では旧予約幅は1段に入らない", rowN >= 2, _
        "rowN=" & rowN

    ' 確保幅の上限(TitleReserve)そのものの固定。
    modTestRunner.Check "TitleReserve: 広い帯では要求どおり280ptを確保", _
        modChrome.TitleReserve(900, 280) = 280, _
        "got=" & modChrome.TitleReserve(900, 280)
    modTestRunner.Check "TitleReserve: 狭い帯では帯幅の半分まで縮める", _
        modChrome.TitleReserve(120, 280) = 60, _
        "got=" & modChrome.TitleReserve(120, 280)
End Sub

' ----------------------------------------------------------------------------
' R4要件D: 部門ラベルの切り詰め(2026-07-30)
' ----------------------------------------------------------------------------
' ヘッダーのタイトルは「チャット + 部門ラベル」で、部門ラベルは
' modChannel.ActiveLabel()由来なので長さが読めない。利用可能幅で切り詰める。
' サロゲートペアの途中で切ると文字そのものが壊れるので、そこも固定する。
Private Sub TestChromeClipToWidth()
    Dim clipped As String, src As String, cut As String
    Dim lastUnit As Long

    modTestRunner.Check "ClipToWidth: 収まる文字列はそのまま返す", _
        modChrome.ClipToWidth("日本語", 100, 12) = "日本語", _
        "got=" & modChrome.ClipToWidth("日本語", 100, 12)

    clipped = modChrome.ClipToWidth("あいうえお", 36, 12)
    modTestRunner.Check "ClipToWidth: あふれたら省略記号を付けて縮める", _
        clipped = "あい" & ChrW(&H2026), "got=" & clipped

    modTestRunner.Check "ClipToWidth: 1字も置けない幅なら空を返す", _
        LenB(modChrome.ClipToWidth("あ", 6, 12)) = 0, _
        "got=" & modChrome.ClipToWidth("あ", 6, 12)

    ' U+1F4DA(サロゲートペア)+全角2字。24pt分だけ許すとペアで止まるはずで、
    ' 上位サロゲート単独で終わってはいけない。
    src = ChrW(&HD83D) & ChrW(&HDCDA) & "部門"
    cut = modChrome.ClipToWidth(src, 24, 12)
    lastUnit = 0
    If Len(cut) > 1 Then lastUnit = AscW(Mid$(cut, Len(cut) - 1, 1))
    If lastUnit < 0 Then lastUnit = lastUnit + 65536
    modTestRunner.Check "ClipToWidth: サロゲートペアの途中で切らない", _
        Left$(cut, 2) = Left$(src, 2) And lastUnit = &HDCDA& And Right$(cut, 1) = ChrW(&H2026), _
        "len=" & Len(cut) & " lastUnit=" & lastUnit

    ' 半角は全角の半分で数える(英語表記で極端に切られないため)。
    modTestRunner.Check "TextSpan: 半角は全角の半分", _
        modChrome.TextSpan("abcd", 12) = modChrome.TextSpan("あい", 12), _
        "half=" & modChrome.TextSpan("abcd", 12) & " full=" & modChrome.TextSpan("あい", 12)

    ' 幅はキャプションの実文字から出す(固定予約幅と実物の食い違いを無くす)。
    modTestRunner.Check "PillWidth: 短いキャプションでも最小幅を下回らない", _
        modChrome.PillWidth("A", 9, 14, 30) = 30, _
        "got=" & modChrome.PillWidth("A", 9, 14, 30)
    modTestRunner.Check "PillWidth: 長いキャプションは文字幅+余白になる", _
        modChrome.PillWidth("しっかり調べる", 9, 14, 30) = 7 * 9 + 14, _
        "got=" & modChrome.PillWidth("しっかり調べる", 9, 14, 30)
End Sub

Public Sub RunAll3()
    On Error GoTo NormEmptyGroupFail
    TestNormalizeForIngestEmpty
NextEmptyPage:
    On Error GoTo EmptyPageFail
    TestChunkerEmptyPageDefense
NextCompacted:
    On Error GoTo CompactedFail
    TestChunkerCompactedPageArray
NextCopyChunk:
    On Error GoTo CopyChunkFail
    TestSharedCopyNextChunkLen
NextBadgeCatalog:
    On Error GoTo BadgeCatalogFail
    TestBadgeCatalogWelcome
NextTileDefault:
    On Error GoTo TileDefaultFail
    TestHubTileDefaultValue
NextChromeFlow:
    On Error GoTo ChromeFlowFail
    TestChromeFlowLeft
NextChromeBudget:
    On Error GoTo ChromeBudgetFail
    TestChromeHeaderBudget
NextChromeClip:
    On Error GoTo ChromeClipFail
    TestChromeClipToWidth
NextDone:
    On Error GoTo 0
    Exit Sub

NormEmptyGroupFail:
    modTestRunner.Check "TestNormalizeForIngestEmpty(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextEmptyPage
EmptyPageFail:
    modTestRunner.Check "TestChunkerEmptyPageDefense(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCompacted
CompactedFail:
    modTestRunner.Check "TestChunkerCompactedPageArray(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCopyChunk
CopyChunkFail:
    modTestRunner.Check "TestSharedCopyNextChunkLen(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBadgeCatalog
BadgeCatalogFail:
    modTestRunner.Check "TestBadgeCatalogWelcome(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextTileDefault
TileDefaultFail:
    modTestRunner.Check "TestHubTileDefaultValue(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextChromeFlow
ChromeFlowFail:
    modTestRunner.Check "TestChromeFlowLeft(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextChromeBudget
ChromeBudgetFail:
    modTestRunner.Check "TestChromeHeaderBudget(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextChromeClip
ChromeClipFail:
    modTestRunner.Check "TestChromeClipToWidth(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone
End Sub
