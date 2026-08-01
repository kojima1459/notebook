Attribute VB_Name = "modTestsPure7"
Option Explicit

' ============================================================================
' modTestsPure7 - R12-3(堅牢化)で入れた純ロジックの回帰テスト
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか:
'   modTestsPure(28,906字)/2(29,542)/5(28,622)/6(27,991)/4(24,989+本件分)は
'   いずれも30,000字上限に対する余裕が無く、憲章§4-6「WARN帯のモジュールに
'   機能を足さない。足す前に分割を裁定する」に従って分割先を新設した。
'   入口は modTestsPure4.RunAll4 の末尾から呼ばれる Public Sub RunAll7()。
'   modTestRunner.RunAllPureTests は modTestsPure.RunAll だけを呼ぶ契約なので、
'   ここへの導線は modTestsPure4 内の1行だけ。消すとテストが「全部PASS」の
'   まま実行されなくなる。
'
' 固定する事実:
'   (6) 出典ページ番号は物理ページ基準(画像だけの表紙があっても手前へずれない)
'   (7) 粗選別(binary_rag)の救済union判定 modSparse.HasAnyKey の境界
'   (9) チャンク本文の連結を配列+Joinにしても出力が1バイトも変わらないこと
'  (10) 統計キーの日付(yyyymmdd/yyyymm/yyyy)が元号に影響されないこと
'
' ■ CanUseTypeArraysの複製について(modTestsPure2/3/4の冒頭コメントと同じ理由):
'   他モジュールのPrivateは呼べないため、軽量な実測プローブを複製する。
'   ExtractedPage()配列を使うテストだけをこれで守る。
' ============================================================================

Private Function CanUseTypeArrays() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As ExtractedPage
    ReDim probe(0 To 0)
    CanUseTypeArrays = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' (6) 出典ページの物理番号(R12-3-6)
' ----------------------------------------------------------------------------
' 表紙が画像のみ(文字層ゼロ)のPDFをGhostscript(txtwrite)で取り込むと、
' 先頭要素が空になる。従来はそれを詰めて 1,2,3… と振り直していたため、
' 全チャンクのページ番号が表紙の枚数ぶん手前にずれ、「出典 p.5」を開くと
' 別のページが出た。GsPageBounds の firstIdx が「読み飛ばした枚数」であり、
' 採番が firstIdx+1 から始まることを固定する。
Private Sub TestGsPhysicalPageNumber()
    Dim ff As String: ff = Chr$(12)

    Dim f1 As Long, l1 As Long
    modTestRunner.Check "R12-3-6: 表紙1枚が空なら firstIdx=1(本文は物理2ページ目)", _
        (modUtilText.GsPageBounds(ff & "本文", f1, l1) = 1 And f1 = 1), _
        "first=" & f1 & " last=" & l1
    Dim f2 As Long, l2 As Long
    modTestRunner.Check "R12-3-6: 表紙2枚が空なら firstIdx=2(本文は物理3ページ目)", _
        (modUtilText.GsPageBounds(ff & ff & "本文", f2, l2) = 1 And f2 = 2), _
        "first=" & f2 & " last=" & l2
    Dim f3 As Long, l3 As Long
    modTestRunner.Check "R12-3-6: 先頭に中身があれば firstIdx=0(物理1ページ目)", _
        (modUtilText.GsPageBounds("本文" & ff, f3, l3) = 1 And f3 = 0), _
        "first=" & f3 & " last=" & l3
    Dim f4 As Long, l4 As Long
    modTestRunner.Check "R12-3-6: 改行だけの先頭ページも1枚として番号を消費する", _
        (modUtilText.GsPageBounds(vbCrLf & ff & "本文", f4, l4) = 1 And f4 = 1), _
        "first=" & f4 & " last=" & l4

    If Not CanUseTypeArrays() Then
        modTestRunner.Check "R12-3-6: BuildPagesFromGsTextはLO環境の制限によりスキップ", True, _
            "ExtractedPage()のReDimが420になる環境。番号規約は上のGsPageBoundsで固定済み。"
        Exit Sub
    End If

    Dim pgs() As ExtractedPage
    Dim tr As Boolean
    If modExtractor.BuildPagesFromGsText(ff & ff & "本文", 100, pgs, tr) Then
        modTestRunner.Check "R12-3-6: 画像表紙2枚のPDFで本文が p.3 になる", _
            (pgs(0).page = 3 And Trim$(pgs(0).Text) = "本文"), _
            "page=" & pgs(0).page & " text=[" & Trim$(pgs(0).Text) & "]"
    Else
        modTestRunner.Check "R12-3-6: 画像表紙2枚のPDFが1ページとして取れる", False, "戻り値=False"
    End If

    Dim pgs2() As ExtractedPage
    Dim tr2 As Boolean
    If modExtractor.BuildPagesFromGsText("1枚目" & ff & ff & "3枚目", 100, pgs2, tr2) Then
        modTestRunner.Check "R12-3-6: 中間の白紙ページも番号を消費する(1と3)", _
            (pgs2(0).page = 1 And pgs2(2).page = 3), _
            "p0=" & pgs2(0).page & " p2=" & pgs2(2).page
    End If
End Sub

' ----------------------------------------------------------------------------
' (7) 粗選別の救済union(R12-3-7)
' ----------------------------------------------------------------------------
' binary_rag の粗選別はハミング距離(ベクトル近似)だけで候補を選ぶため、
' 「第12条」のような決定的な語を持つ行が、加点を受ける前に落ちていた。
' 救済unionの判定規則(2文字以上・大小無視の部分一致)を固定する。
Private Sub TestHasAnyKey()
    Dim keys As String: keys = "第12条|保険金"

    modTestRunner.Check "R12-3-7: キーを含む本文はTrue", _
        modSparse.HasAnyKey(keys, "…第12条(保険金の支払)…"), ""
    modTestRunner.Check "R12-3-7: 2つ目のキーだけでもTrue", _
        modSparse.HasAnyKey(keys, "保険金の請求手続について"), ""
    modTestRunner.Check "R12-3-7: どのキーも含まなければFalse", _
        (modSparse.HasAnyKey(keys, "第11条の解約について") = False), ""
    modTestRunner.Check "R12-3-7: キーが空ならFalse(全件unionしない)", _
        (modSparse.HasAnyKey("", "第12条") = False), ""
    modTestRunner.Check "R12-3-7: 本文が空ならFalse", _
        (modSparse.HasAnyKey(keys, "") = False), ""
    modTestRunner.Check "R12-3-7: 1文字キーは拾わない(KeyScoreと同じ規則)", _
        (modSparse.HasAnyKey("第|", "第") = False), ""
    modTestRunner.Check "R12-3-7: 英字は大小を区別しない", _
        modSparse.HasAnyKey("abc", "XABCX"), ""
End Sub

' ----------------------------------------------------------------------------
' (9) チャンク連結の配列化で出力が変わらないこと(R12-3-9)
' ----------------------------------------------------------------------------
Private Sub TestChunkJoinUnchanged()
    modTestRunner.Check "R12-3-9: JoinSplitNumbers_空白入りの条番号を詰める", _
        (modChunker.JoinSplitNumbers("第 １ ０ 条 保険金") = "第10条 保険金"), _
        "実際=[" & modChunker.JoinSplitNumbers("第 １ ０ 条 保険金") & "]"
    modTestRunner.Check "R12-3-9: JoinSplitNumbers_「第」が無い行は素通し", _
        (modChunker.JoinSplitNumbers("  支払事由について  ") = "  支払事由について  "), _
        "実際=[" & modChunker.JoinSplitNumbers("  支払事由について  ") & "]"
    modTestRunner.Check "R12-3-9: JoinSplitNumbers_単位漢字が来なければ元のまま", _
        (modChunker.JoinSplitNumbers("第 12 回") = "第 12 回"), _
        "実際=[" & modChunker.JoinSplitNumbers("第 12 回") & "]"

    ' 2026-08-01(同梱-7): digitsがループ内Dimで、1行に「第」が複数回現れると
    ' 前回値を引き継いで「第1条と第2条」→「第12条」に化けていたバグの固定。
    modTestRunner.Check "同梱-7: 1行に第N条が2回でも化けない", _
        (modChunker.JoinSplitNumbers("第 1 条 と 第 2 条") = "第1条 と 第2条"), _
        "実際=[" & modChunker.JoinSplitNumbers("第 1 条 と 第 2 条") & "]"
    modTestRunner.Check "同梱-7: 1行に第N条が3回でも化けない", _
        (modChunker.JoinSplitNumbers("第 1 条 、 第 2 条 、 第 3 条") = "第1条 、 第2条 、 第3条"), _
        "実際=[" & modChunker.JoinSplitNumbers("第 1 条 、 第 2 条 、 第 3 条") & "]"
    modTestRunner.Check "同梱-7: 2件目が単位漢字無しでも1件目の連結が壊れない", _
        (modChunker.JoinSplitNumbers("第 1 条 と 第 2 回") = "第1条 と 第 2 回"), _
        "実際=[" & modChunker.JoinSplitNumbers("第 1 条 と 第 2 回") & "]"

    If Not CanUseTypeArrays() Then
        modTestRunner.Check "R12-3-9: 大入力の同一性検査はLO環境の制限によりスキップ", True, _
            "ExtractedPage()のReDimが420になる環境。"
        Exit Sub
    End If

    ' 見出しが1つも無い大きめの入力(旧実装でO(n²)を踏む形)。行の並びと
    ' 区切り文字が組み立て方を変えても保たれることを見る。
    Dim big As String, i As Long
    Dim rowsBuf() As String: ReDim rowsBuf(1 To 400)
    For i = 1 To 400
        rowsBuf(i) = "支払事由の詳細を記載した行です。"
    Next i
    big = Join(rowsBuf, vbLf)

    Dim pagesBig() As ExtractedPage
    ReDim pagesBig(0 To 0)
    pagesBig(0).page = 1
    pagesBig(0).Text = big

    Dim outC() As ShelfChunk
    Dim nBig As Long
    nBig = modChunker.ChunkPagesEx(pagesBig, 700, 150, 1800, "structure", outC)
    modTestRunner.Check "R12-3-9: 見出し無し大入力でもチャンクが生成される", _
        (nBig > 0), "件数=" & nBig
    If nBig < 1 Then Exit Sub

    modTestRunner.Check "R12-3-9: 連結の区切りは行区切り(vbLf)のまま", _
        (InStr(outC(0).full_text, "行です。" & vbLf & "支払事由") > 0), _
        "先頭120字=[" & Left$(outC(0).full_text, 120) & "]"
    modTestRunner.Check "R12-3-9: 末尾に余分な空行が付かない", _
        (Right$(outC(nBig - 1).full_text, 1) <> vbLf), _
        "末尾20字=[" & Right$(outC(nBig - 1).full_text, 20) & "]"

    ' 空行が混ざっても区切りが二重にならない(旧実装は空行を積まなかった)。
    Dim pages2() As ExtractedPage
    ReDim pages2(0 To 0)
    pages2(0).page = 1
    pages2(0).Text = "一行目" & vbLf & vbLf & "  " & vbLf & "二行目"
    Dim out2() As ShelfChunk
    If modChunker.ChunkPagesEx(pages2, 700, 150, 1800, "structure", out2) > 0 Then
        modTestRunner.Check "R12-3-9: 空行は連結に持ち込まない", _
            (InStr(out2(0).full_text, "一行目" & vbLf & "二行目") > 0), _
            "実際=[" & out2(0).full_text & "]"
    End If
End Sub

' ----------------------------------------------------------------------------
' (10) 統計キーの日付(R12-3-10)
' ----------------------------------------------------------------------------
' 和暦カレンダー端末では Format$(Date,"yyyymmdd") が元号年を返す。統計キーが
' "00080801" と "20260801" に割れると、同じ日の記録が二重化し、しかも誰も
' 気付けない(集計が静かに壊れる)。数値合成であることを桁まで固定する。
Private Sub TestIsoCompactKeys()
    Dim d1 As Date: d1 = DateSerial(2026, 1, 2)
    modTestRunner.Check "R12-3-10: IsoDateCompact_ゼロ詰め8桁", _
        (modUtilText.IsoDateCompact(d1) = "20260102"), _
        "実際=[" & modUtilText.IsoDateCompact(d1) & "]"
    modTestRunner.Check "R12-3-10: IsoYm_ゼロ詰め6桁", _
        (modUtilText.IsoYm(d1) = "202601"), "実際=[" & modUtilText.IsoYm(d1) & "]"
    modTestRunner.Check "R12-3-10: IsoYear_4桁", _
        (modUtilText.IsoYear(d1) = "2026"), "実際=[" & modUtilText.IsoYear(d1) & "]"
    modTestRunner.Check "R12-3-10: IsoDateCompactはIsoDateから区切りを抜いた形", _
        (modUtilText.IsoDateCompact(d1) = Replace(modUtilText.IsoDate(d1), "-", "")), ""
    modTestRunner.Check "R12-3-10: IsoYmはIsoDateCompactの先頭6字", _
        (modUtilText.IsoYm(d1) = Left$(modUtilText.IsoDateCompact(d1), 6)), ""
    modTestRunner.Check "R12-3-10: 日付順と文字列順が一致する(キーの昇順性)", _
        (modUtilText.IsoDateCompact(DateSerial(2026, 9, 9)) < _
         modUtilText.IsoDateCompact(DateSerial(2026, 10, 1))), ""
End Sub

' ----------------------------------------------------------------------------
' (3) 恒久失敗のバックオフ(R12-3-3)- 判定側の純ロジック
' ----------------------------------------------------------------------------
' 連続3回失敗した資料は status="failed_permanent" になり、以降の自動同期の
' 差分判定では拾い直さない(拾い直すと毎回同じ失敗を繰り返して同期時間と
' err_log を食い続ける)。ファイルが更新されたときは従来どおり replace。
Private Sub TestFailedPermanentScope()
    modTestRunner.Check "R12-3-3: failed_permanent は keep のまま(自動再試行しない)", _
        (modShelfSync.ResolveDecision("keep", True, "failed_permanent") = "keep"), _
        "実際=" & modShelfSync.ResolveDecision("keep", True, "failed_permanent")
    modTestRunner.Check "R12-3-3: failed は従来どおり replace(次回再試行)", _
        (modShelfSync.ResolveDecision("keep", True, "failed") = "replace"), ""
    modTestRunner.Check "R12-3-3: ファイルが更新されていれば permanent でも replace", _
        (modShelfSync.ResolveDecision("replace", True, "failed_permanent") = "replace"), ""
    modTestRunner.Check "R12-3-3: missing は従来どおり replace", _
        (modShelfSync.ResolveDecision("keep", True, "missing") = "replace"), ""
End Sub

' ----------------------------------------------------------------------------
' (1) 数式インジェクション対策の共通関数(R12-2-1)
' ----------------------------------------------------------------------------
' modChatLog.SanitizeForCellを modUtilText へ共通化した。未信頼テキストを
' セルへ書く3経路(modInsightIo.AppendRow/modPack.ImportChunksDedup/
' modChannel経由のmodStats.SetStatValue)がここへ揃うため、先頭 =/+/-/@ の
' 無害化と通常文字の素通しを固定する。
Private Sub TestSanitizeForCell()
    modTestRunner.Check "R12-2-1: 先頭=は'を前置", _
        (modUtilText.SanitizeForCell("=SUM(A1:A9)") = "'=SUM(A1:A9)"), _
        "実際=[" & modUtilText.SanitizeForCell("=SUM(A1:A9)") & "]"
    modTestRunner.Check "R12-2-1: 先頭+は'を前置", _
        (modUtilText.SanitizeForCell("+1+1") = "'+1+1"), ""
    modTestRunner.Check "R12-2-1: 先頭-は'を前置", _
        (modUtilText.SanitizeForCell("-1-1") = "'-1-1"), ""
    modTestRunner.Check "R12-2-1: 先頭@は'を前置", _
        (modUtilText.SanitizeForCell("@SUM(1)") = "'@SUM(1)"), ""
    modTestRunner.Check "R12-2-1: 通常文字は素通し", _
        (modUtilText.SanitizeForCell("第1条について") = "第1条について"), ""
    modTestRunner.Check "R12-2-1: 空文字は空文字のまま", _
        (modUtilText.SanitizeForCell("") = ""), ""
    modTestRunner.Check "R12-2-1: 危険文字が先頭以外にあっても無害", _
        (modUtilText.SanitizeForCell("保険金は=1000円です") = "保険金は=1000円です"), ""
End Sub

Public Sub RunAll7()
    On Error GoTo PageFail
    TestGsPhysicalPageNumber
NextKey:
    On Error GoTo KeyFail
    TestHasAnyKey
NextChunk:
    On Error GoTo ChunkFail
    TestChunkJoinUnchanged
NextIso:
    On Error GoTo IsoFail
    TestIsoCompactKeys
NextScope:
    On Error GoTo ScopeFail
    TestFailedPermanentScope
NextSanitize:
    On Error GoTo SanitizeFail
    TestSanitizeForCell
NextDone7:
    On Error GoTo 0
    Exit Sub

PageFail:
    modTestRunner.Check "TestGsPhysicalPageNumber(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextKey
KeyFail:
    modTestRunner.Check "TestHasAnyKey(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextChunk
ChunkFail:
    modTestRunner.Check "TestChunkJoinUnchanged(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextIso
IsoFail:
    modTestRunner.Check "TestIsoCompactKeys(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextScope
ScopeFail:
    modTestRunner.Check "TestFailedPermanentScope(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSanitize
SanitizeFail:
    modTestRunner.Check "TestSanitizeForCell(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone7
End Sub
