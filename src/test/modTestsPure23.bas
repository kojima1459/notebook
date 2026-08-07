Attribute VB_Name = "modTestsPure23"
Option Explicit

' ============================================================================
' modTestsPure23 - R21-3(章検出の根治と俯瞰の復旧・実機第8報②)の純ロジック
'   ゴールデン。入口は modTestsPure22.RunAll22 の末尾から呼ばれる RunAll23。
' ----------------------------------------------------------------------------
' 仕様書(docs/dev/spec_20260807_R21_実機第8報.md)§E3の題材1〜4を実装する:
'   題材1: 目次+A/B/C章立て最小資料 → 目次行が独立章にならない/
'          本文が正しい章の下に入る(modChunker.ChunkPagesEx構造モード)。
'   題材2: 小見出し過多 → 上限到達時に縮退統合が働く(modOutlineBuild.
'          GroupChapters)。usage_log("outline_capped")はWorksheet依存の
'          modOutlineBuild.BuildOutlineFor内でしか呼べないため、ここでは
'          その発火条件と完全に同じ式(uniqRaw > MAX_CHAPTERS=60)が成立する
'          ことを確認する(単一情報源・BuildOutlineFor:129行目相当と対)。
'   題材3: 目次行vs本文見出しの表記ゆれ → ChapterKeyOfの正規化後に同一キーへ
'          統合される。
'   題材4: Backfill射程 → modChunkMeta.MetaOfは同一full_textから常に同一の
'          section_path/refs_outを再現する(BackfillOne.Phase1は保存済み
'          full_textをそのままMetaOfへ渡すだけ=決定論的な再計算であり、
'          breadcrumb(=my_knowledge.full_text)そのものへの書込みは無い)。
'          BackfillOne/DetectLegacyDocs自体はWorksheetに触れるためLO純粋
'          テストからは呼べない(modOutlineBuild/modBackfillの既存注記と
'          同型)。my_knowledgeへの書込みが無いことはコードレビューで確認済み
'          (modBackfill.bas: 書込みは modChunkMetaStore/modOutlineStore/
'          modSynonymStoreのみで、SH_KNOWLEDGEシートへの代入文は無い)。
'          このテストは「その再計算が決定論的である」という、breadcrumb
'          不変の再取込ゼロ設計が成立するための必要条件を固定するもの。
'
' ■ CanUseTypeArraysの複製について(modTestsPure2/3/4冒頭コメントと同じ理由):
'   modTestsPureのPrivate Function CanUseTypeArraysは別モジュールのPrivateで
'   あるため、ここから呼べない。ExtractedPage()/ShelfChunk()を扱う題材1の
'   テストだけをこれで守る(paths()=String配列だけの題材2〜4は対象外)。
' ============================================================================

' 章検出ロジックの上限(modOutlineBuild.Private定数の複製。テスト側で期待値を
' 手計算するための参照値。値がズレたらこのテストごと落ちるので複製リスクは無い)。
Private Const MAX_CHAPTERS23 As Long = 60

Private Function CanUseTypeArrays() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As ShelfChunk
    ReDim probe(0 To 0)
    CanUseTypeArrays = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ============================================================================
' 題材1: 目次+A/B/C章立て最小資料
' ============================================================================

' breadcrumb「【資料 > 章 > 節】」の「章」要素だけを取り出す(テスト専用の
' 薄い再実装。modChunkMeta.ExtractSectionPathと同じ冒頭ガードだが、モジュール
' を跨いだ複製は許容する=modBackfill.LooksLikeBreadcrumbLineと同じ判断)。
Private Function ChapterOfChunk(ByVal fullText As String) As String
    Dim lf As Long: lf = InStr(fullText, vbLf)
    Dim line1 As String
    If lf > 0 Then line1 = Left$(fullText, lf - 1) Else line1 = fullText
    If Left$(line1, 1) <> "【" Then Exit Function
    Dim cl As Long: cl = InStr(line1, "】")
    If cl < 3 Then Exit Function
    Dim parts() As String: parts = Split(Mid$(line1, 2, cl - 2), " > ")
    If UBound(parts) >= 1 Then ChapterOfChunk = Trim$(parts(1))
End Function

' "|"区切りの箱に値が(重複なく)含まれるかの簡易カウント。
Private Sub AddDistinct(ByRef box As String, ByVal v As String)
    If LenB(v) = 0 Then Exit Sub
    If InStr(1, "|" & box & "|", "|" & v & "|", vbBinaryCompare) > 0 Then Exit Sub
    If LenB(box) > 0 Then box = box & "|" & v Else box = v
End Sub

Private Function BoxCount(ByVal box As String) As Long
    If LenB(box) = 0 Then Exit Function
    BoxCount = Len(box) - Len(Replace(box, "|", "")) + 1
End Function

Private Sub TestTocPageSuppression()
    If Not CanUseTypeArrays() Then
        modTestRunner.Check "題材1: LO環境の既知の制限によりスキップ", True, _
            "modTypesのPublic Type配列(ExtractedPage()/ShelfChunk())のReDimが" & _
            "使えない環境。Excel実機受入で確認すること(modTestsPure冒頭コメント参照)。"
        Exit Sub
    End If

    ' page1=目次(3行とも「見出しっぽい#付き+リーダー記号+頁番号」という
    ' 最悪ケース。Visionが目次行にまで# を付けてしまっても、ページ密度判定
    ' (LooksLikeTocPage)が見出し化を止めることを確認する)。
    ' page2〜4=A/B/C章の本文(単一#が正しく章境界になることも合わせて確認)。
    Dim pages(0 To 3) As ExtractedPage
    pages(0).page = 1
    pages(0).Text = "# 第1章 総則．．．．．．12" & vbLf & _
                     "# 第2章 保険金の支払．．．．．．45" & vbLf & _
                     "# 第3章 雑則．．．．．．78"
    pages(1).page = 2
    pages(1).Text = "# 第1章 総則" & vbLf & "総則ページの本文です。目的と適用範囲を定める。"
    pages(2).page = 5
    pages(2).Text = "# 第2章 保険金の支払" & vbLf & "保険金の支払条件について定める本文です。"
    pages(3).page = 9
    pages(3).Text = "# 第3章 雑則" & vbLf & "雑則として届出義務等を定める本文です。"

    Dim chunks() As ShelfChunk
    Dim n As Long
    n = modChunker.ChunkPagesEx(pages, 700, 100, 0, "structure", chunks)
    modTestRunner.Check "題材1_前提_チャンクが作れる", (n >= 1), "n=" & n
    If n < 1 Then Exit Sub

    Dim distinctBox As String
    Dim hasLeaderInChapter As Boolean
    Dim tocTextAsBody As Boolean
    Dim ch1HasBody As Boolean
    Dim i As Long
    For i = 0 To n - 1
        Dim ch As String: ch = ChapterOfChunk(chunks(i).full_text)
        If LenB(ch) > 0 Then
            AddDistinct distinctBox, ch
            If InStr(ch, "．") > 0 Then hasLeaderInChapter = True
        Else
            ' 章が未確定(=目次ページの内容)の初期ブロックに、目次の生テキストが
            ' 「本文」としてそのまま残っていること(見出し化されず生き残る)。
            If InStr(chunks(i).full_text, "第1章 総則．．．．．．12") > 0 Then tocTextAsBody = True
        End If
        If ch = "第1章 総則" And InStr(chunks(i).full_text, "目的と適用範囲") > 0 Then ch1HasBody = True
    Next i

    modTestRunner.Check "題材1_目次行はどの章の名前にも現れない(リーダー記号混入なし)", _
        (Not hasLeaderInChapter), "distinct=" & distinctBox
    modTestRunner.Check "題材1_章はA/B/Cの3つだけ(目次行が4つ目の章にならない)", _
        (BoxCount(distinctBox) = 3), "distinct=" & distinctBox & " n=" & BoxCount(distinctBox)
    modTestRunner.Check "題材1_目次の生テキストは見出し化されず本文として残る", tocTextAsBody
    modTestRunner.Check "題材1_ChapterBody相当: 第1章の本文が正しく入る", ch1HasBody
End Sub

' 題材1(データモデル側の裏付け): 目次行由来のキーと本文見出しキーが
' ChapterKeyOf正規化で同一キーへ統合され、GroupChaptersが1章として束ねる
' (=modOutlineBuild.ChapterBodyは両方のチャンクを同じ章の本文として拾える)。
Private Sub TestChapterUnificationInGrouping()
    Dim paths(1 To 4) As String
    paths(1) = "第1章 総則……12"     ' 目次行由来(未正規化のsection_path想定)
    paths(2) = "第1章 総則"          ' 本文見出し由来
    paths(3) = "第1章 総則"
    paths(4) = "第2章 保険金の支払"

    Dim keys() As String, counts() As Long
    Dim uniqRaw As Long, mergedN As Long, droppedN As Long
    Dim nCh As Long
    nCh = modOutlineBuild.GroupChapters(paths, 4, keys, counts, uniqRaw, mergedN, droppedN)

    modTestRunner.Check "題材1_目次行キーと本文見出しキーが正規化で統合(章は2つ)", _
        (nCh = 2), "nCh=" & nCh
    modTestRunner.Check "題材1_統合後の章1は目次行+本文2件=3チャンク", _
        (nCh >= 1 And counts(1) = 3), "counts(1)=" & IIf(nCh >= 1, counts(1), -1)
    modTestRunner.Check "題材1_ChapterBody相当: 目次行キーの生値もChapterKeyMatchesで拾える", _
        (nCh >= 1 And modOutlineBuild.ChapterKeyMatches("第1章 総則", keys(1)))
End Sub

' ============================================================================
' 題材2: 小見出し過多 → 上限到達時に縮退統合(GroupChaptersの2パス縮退)
' ============================================================================

Private Sub BuildUniformChapters(ByVal headingN As Long, ByVal chunksPerHeading As Long, _
                                 ByRef outPaths() As String, ByRef outTotal As Long)
    Dim total As Long: total = headingN * chunksPerHeading
    ReDim outPaths(1 To total)
    Dim idx As Long, h As Long, r As Long
    For h = 1 To headingN
        Dim nm As String: nm = "小見出し" & Format(h, "0000")
        For r = 1 To chunksPerHeading
            idx = idx + 1
            outPaths(idx) = nm
        Next r
    Next h
    outTotal = total
End Sub

Private Sub TestGroupChaptersBoundaryAtCap()
    ' 境界: ちょうど60章(各3チャンク=最小サイズガードの閾値ちょうど)は
    ' 統合されず、無変更のまま60章として通る(旧実装でも切り捨てなかった域)。
    Dim paths() As String, total As Long
    BuildUniformChapters MAX_CHAPTERS23, 3, paths, total

    Dim keys() As String, counts() As Long
    Dim uniqRaw As Long, mergedN As Long, droppedN As Long
    Dim nCh As Long
    nCh = modOutlineBuild.GroupChapters(paths, total, keys, counts, uniqRaw, mergedN, droppedN)

    modTestRunner.Check "題材2_境界_ちょうど60章はそのまま60章(縮退なし)", _
        (nCh = 60), "nCh=" & nCh
    modTestRunner.Check "題材2_境界_uniqRaw=60はoutline_capped発火条件を満たさない", _
        Not (uniqRaw > MAX_CHAPTERS23), "uniqRaw=" & uniqRaw
    modTestRunner.Check "題材2_境界_統合0件", (mergedN = 0), "merged=" & mergedN
End Sub

Private Sub TestGroupChaptersDegenerate()
    ' 61章(各3チャンク=単独では最小サイズガードに掛からない大きさ)。
    ' MAX_CHAPTERS超過なので2パス目(粒度を上げた再集約)が働く。
    ' 手計算: total=183, minSize2=Ceil(183/60)=4。count=3の章はminSize2=4に
    ' 単独で届かないので2件ずつペアで統合(30ペア=60章分)+最後の1件が単独で
    ' 強制フラッシュ=最終31章。統合数=61-31=30。取りこぼしゼロ(sum=183)。
    Dim paths() As String, total As Long
    BuildUniformChapters MAX_CHAPTERS23 + 1, 3, paths, total

    Dim keys() As String, counts() As Long
    Dim uniqRaw As Long, mergedN As Long, droppedN As Long
    Dim nCh As Long
    nCh = modOutlineBuild.GroupChapters(paths, total, keys, counts, uniqRaw, mergedN, droppedN)

    modTestRunner.Check "題材2_uniqRaw=61(縮退前のユニーク章数)", (uniqRaw = 61), "uniqRaw=" & uniqRaw
    modTestRunner.Check "題材2_uniqRaw>60はoutline_capped発火条件と一致" & _
        "(BuildOutlineFor内のIf uniqRaw超過チェックと同じ式)", _
        (uniqRaw > MAX_CHAPTERS23)
    modTestRunner.Check "題材2_縮退統合後は上限以下(無言切り捨てせず61→31章)", _
        (nCh = 31 And nCh <= MAX_CHAPTERS23), "nCh=" & nCh
    modTestRunner.Check "題材2_統合数=61-31=30", (mergedN = 30), "merged=" & mergedN
    modTestRunner.Check "題材2_取りこぼしゼロ(dropped=0)", (droppedN = 0), "dropped=" & droppedN

    Dim sumC As Long, i As Long
    For i = 1 To nCh
        sumC = sumC + counts(i)
    Next i
    modTestRunner.Check "題材2_全183チャンクがどこかの章に属す(合計不変)", _
        (sumC = total), "sum=" & sumC & " total=" & total
End Sub

Private Sub TestGroupChaptersMinSizeGuard()
    ' MAX_CHAPTERS未満でも常時働く最小サイズガード(E1バックストップ)。
    ' 1チャンクの「見出し」5件(A,B,C,D,E)。閾値3で先頭から束ねると
    ' A+B+C=3で1件目確定、D+E=2は末尾強制フラッシュでもう1件=最終2章。
    Dim paths(1 To 5) As String
    paths(1) = "見出しA": paths(2) = "見出しB": paths(3) = "見出しC"
    paths(4) = "見出しD": paths(5) = "見出しE"

    Dim keys() As String, counts() As Long
    Dim uniqRaw As Long, mergedN As Long, droppedN As Long
    Dim nCh As Long
    nCh = modOutlineBuild.GroupChapters(paths, 5, keys, counts, uniqRaw, mergedN, droppedN)

    modTestRunner.Check "題材2_最小ガード_上限未満でも1-2チャンクの章は統合される", _
        (nCh = 2), "nCh=" & nCh
    modTestRunner.Check "題材2_最小ガード_1章目はA/B/Cの3件", _
        (nCh >= 1 And counts(1) = 3)
    modTestRunner.Check "題材2_最小ガード_2章目はD/Eの2件(末尾は強制フラッシュ)", _
        (nCh >= 2 And counts(2) = 2)
    modTestRunner.Check "題材2_最小ガード_ChapterKeyMatchesで隣接統合先を判定できる", _
        (nCh >= 1 And modOutlineBuild.ChapterKeyMatches("見出しB", keys(1)) And _
         Not modOutlineBuild.ChapterKeyMatches("見出しD", keys(1)))
End Sub

' ============================================================================
' 題材3: 目次行vs本文見出しの表記ゆれ → ChapterKeyOfの正規化
' ============================================================================

Private Sub TestChapterKeyOfNormalization()
    ' 前提: ChapterKeyOfへの入力はmodSparse.NormalizeForSearch済みのsection_path
    ' なので、頁番号は必ず半角数字("0"-"9")。ここもそれに合わせる。
    modTestRunner.Check "題材3_全角ピリオドのリーダー(6字)+半角頁番号を除去", _
        (modOutlineBuild.ChapterKeyOf("第3章 総則．．．．．．12") = "第3章 総則")
    modTestRunner.Check "題材3_半角ピリオドのリーダーを除去", _
        (modOutlineBuild.ChapterKeyOf("第3章 総則......12") = "第3章 総則")
    modTestRunner.Check "題材3_中黒(・)のリーダーを除去", _
        (modOutlineBuild.ChapterKeyOf("第3章 総則・・・・・・12") = "第3章 総則")
    modTestRunner.Check "題材3_リーダーの間に空白が混じっても除去", _
        (modOutlineBuild.ChapterKeyOf("第3章 総則 ． ． ． ． 12") = "第3章 総則")
    modTestRunner.Check "題材3_目次行キーと本文見出しキーが同一キーへ統合", _
        (modOutlineBuild.ChapterKeyOf("第3章 総則......12") = modOutlineBuild.ChapterKeyOf("第3章 総則"))
    modTestRunner.Check "題材3_>以降(条)は除いてから正規化", _
        (modOutlineBuild.ChapterKeyOf("第3章 総則．．．．12>第5条(免責)") = "第3章 総則")

    ' 誤爆防止(境界): リーダーが1字だけ(普通の句点相当)は削らない
    modTestRunner.Check "題材3_リーダー1字だけは誤爆防止で削らない(境界)", _
        (modOutlineBuild.ChapterKeyOf("第3章 総則.12") = "第3章 総則.12")
    ' 頁番号が無い(末尾が数字でない)場合は削らない
    modTestRunner.Check "題材3_末尾が数字でなければ削らない", _
        (modOutlineBuild.ChapterKeyOf("第3章 総則．．．．") = "第3章 総則．．．．")
    ' 全部がリーダー+数字だけ(章名が残らない)場合は安全側で削らない
    modTestRunner.Check "題材3_章名が消える場合は安全側で何もしない(境界)", _
        (modOutlineBuild.ChapterKeyOf("．．．．．．12") = "．．．．．．12")
    modTestRunner.Check "題材3_空/空白だけは空文字", (modOutlineBuild.ChapterKeyOf("   ") = "")
End Sub

' ============================================================================
' 題材4: Backfillの射程(chunk_meta/doc_outlineは再構築されるがbreadcrumbは不変)
' ============================================================================

Private Sub TestBackfillScopeIsDeterministic()
    ' modBackfill.BackfillOne の Phase1 は、保存済み my_knowledge.full_text
    ' (=breadcrumb焼き込み済みの本文)をそのまま modChunkMeta.MetaOf へ渡すだけ
    ' (modBackfill.bas: For i = 1 To nSrc: modChunkMeta.MetaOf texts(i), ...)。
    ' breadcrumbそのものへの書込みは無い(my_knowledgeへの代入文はゼロ。
    ' 書込み先はchunk_meta/doc_outline/synonymsのみ)。この不変を支えているのは
    ' 「同じ入力から常に同じsection_path/refs_outが再現される」という決定論的
    ' 性質そのものなので、それをここで固定する。
    Dim fullTextAtIngest As String
    fullTextAtIngest = "【〔資料〕 > 第3章 総則 > 第12条(免責)】" & vbLf & _
        "本文…第8条による免責が適用される。別表2のとおり。"

    Dim pathAtIngest As String, refsAtIngest As String
    modChunkMeta.MetaOf fullTextAtIngest, pathAtIngest, refsAtIngest

    ' 「後日のBackfill」を模す: 保存済みfull_textは不変のまま、もう一度
    ' MetaOfへ通す(BackfillOne.Phase1が実際にしていることと同じ操作)。
    Dim fullTextAtBackfill As String: fullTextAtBackfill = fullTextAtIngest
    Dim pathAtBackfill As String, refsAtBackfill As String
    modChunkMeta.MetaOf fullTextAtBackfill, pathAtBackfill, refsAtBackfill

    modTestRunner.Check "題材4_breadcrumb文字列そのものはBackfillで変化しない", _
        (fullTextAtIngest = fullTextAtBackfill)
    modTestRunner.Check "題材4_同一full_textから再計算してもsection_pathが再現する", _
        (pathAtIngest = pathAtBackfill), "A=" & pathAtIngest & " B=" & pathAtBackfill
    modTestRunner.Check "題材4_同一full_textから再計算してもrefs_outが再現する", _
        (refsAtIngest = refsAtBackfill), "A=" & refsAtIngest & " B=" & refsAtBackfill
    modTestRunner.Check "題材4_再計算後もChapterKeyOfの章キーが一致する(doc_outline再構築の入力が安定)", _
        (modOutlineBuild.ChapterKeyOf(pathAtIngest) = modOutlineBuild.ChapterKeyOf(pathAtBackfill))
End Sub

' ============================================================================
' 付随: E1(modChunker)の境界(ClassifyLine の#個数判定・LooksLikeTocPage)
' ============================================================================

Private Sub TestClassifyLineHashLevels()
    modTestRunner.Check "E1_単一#は章(lbl=1)", (modChunker.ClassifyLine("# 第1章 総則") = 1)
    modTestRunner.Check "E1_##は節(lbl=2)", (modChunker.ClassifyLine("## 第1条 目的") = 2)
    modTestRunner.Check "E1_###も節扱い(lbl=2。3階層目を新設しない)", _
        (modChunker.ClassifyLine("### さらに深い小見出し") = 2)
    modTestRunner.Check "E1_#の直後にスペースが無ければ見出しではない", _
        (modChunker.ClassifyLine("#第1章タグ的な文字列") = 0)
End Sub

Private Sub TestLooksLikeTocPage()
    Dim tocPage As String
    tocPage = "第1章 総則．．．．１２" & vbLf & "第2章 保険金の支払．．．．４５" & vbLf & _
              "第3章 雑則．．．．７８"
    modTestRunner.Check "E1_目次ページ(3/3行が目次形式)はTrue", _
        modChunker.LooksLikeTocPage(tocPage)

    Dim normalPage As String
    normalPage = "この資料は社内規程である。" & vbLf & "第1条は総則を定める。" & vbLf & _
                 "詳細は次項のとおり。" & vbLf & "以上。"
    modTestRunner.Check "E1_目次形式0件の通常ページはFalse", _
        (Not modChunker.LooksLikeTocPage(normalPage))

    Dim halfPage As String   ' 4行中ちょうど半分(2行)が目次形式=境界でTrue
    halfPage = "第1章 総則．．１２" & vbLf & "本文の説明が1行。" & vbLf & _
               "第2章 雑則．．３４" & vbLf & "本文の説明がもう1行。"
    modTestRunner.Check "E1_半数ちょうど(2/4行)は境界でTrue", _
        modChunker.LooksLikeTocPage(halfPage)

    Dim mostlyNormal As String  ' 4行中1行だけ目次形式=半数未満でFalse
    mostlyNormal = "第1章 総則．．１２" & vbLf & "本文の説明が1行。" & vbLf & _
                   "さらに本文が続く。" & vbLf & "もう1行の本文。"
    modTestRunner.Check "E1_半数未満(1/4行)はFalse(誤爆しない)", _
        (Not modChunker.LooksLikeTocPage(mostlyNormal))

    modTestRunner.Check "E1_空文字はFalse", (Not modChunker.LooksLikeTocPage(""))
End Sub

' ============================================================================
Public Sub RunAll23()
    On Error GoTo TocFail23
    TestTocPageSuppression
NextUnify23:
    On Error GoTo UnifyFail23
    TestChapterUnificationInGrouping
NextCapBound23:
    On Error GoTo CapBoundFail23
    TestGroupChaptersBoundaryAtCap
NextDegenerate23:
    On Error GoTo DegenerateFail23
    TestGroupChaptersDegenerate
NextMinGuard23:
    On Error GoTo MinGuardFail23
    TestGroupChaptersMinSizeGuard
NextKeyNorm23:
    On Error GoTo KeyNormFail23
    TestChapterKeyOfNormalization
NextBackfill23:
    On Error GoTo BackfillFail23
    TestBackfillScopeIsDeterministic
NextHash23:
    On Error GoTo HashFail23
    TestClassifyLineHashLevels
NextTocDet23:
    On Error GoTo TocDetFail23
    TestLooksLikeTocPage
NextDone23:
    On Error GoTo 0
    Exit Sub

TocFail23:
    modTestRunner.Check "TestTocPageSuppression(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextUnify23
UnifyFail23:
    modTestRunner.Check "TestChapterUnificationInGrouping(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCapBound23
CapBoundFail23:
    modTestRunner.Check "TestGroupChaptersBoundaryAtCap(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDegenerate23
DegenerateFail23:
    modTestRunner.Check "TestGroupChaptersDegenerate(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextMinGuard23
MinGuardFail23:
    modTestRunner.Check "TestGroupChaptersMinSizeGuard(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextKeyNorm23
KeyNormFail23:
    modTestRunner.Check "TestChapterKeyOfNormalization(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBackfill23
BackfillFail23:
    modTestRunner.Check "TestBackfillScopeIsDeterministic(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextHash23
HashFail23:
    modTestRunner.Check "TestClassifyLineHashLevels(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextTocDet23
TocDetFail23:
    modTestRunner.Check "TestLooksLikeTocPage(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone23
End Sub
