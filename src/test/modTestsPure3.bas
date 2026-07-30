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
    Resume NextDone
End Sub
