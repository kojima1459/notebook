Attribute VB_Name = "modTestsPure4"
Option Explicit

' ============================================================================
' modTestsPure4 - modTestsPure3の分割先(2026-07-31 R6: 画像PDFのOCR取込)
' ----------------------------------------------------------------------------
' 役割:
'   R6(画像PDFをGhostscriptでページ画像化してOCR取込する)で追加した純ロジック
'   を検証する。
'     ・optOcrCore: GSコマンド文字列の組み立て(引用符の位置まで固定)、
'       一時フォルダ・出力ファイル名の規約、ページ上限と切り詰め判定の算数
'     ・modUtil: ページ付きテキストの符号化/復号(JoinPagedText/SplitPagedText)
'     ・modExtractor: 「全ページ化け」のPDFをOCR経路へ回す分岐(R6追補)
'   入口は modTestsPure3.RunAll3 の末尾から呼ばれる Public Sub RunAll4()。
'   modTestRunner.RunAllPureTests は modTestsPure.RunAll だけを呼ぶ契約なので、
'   本モジュールへの導線は modTestsPure3 内に置く(3段目までと同じ中継方式)。
'
' なぜGSコマンドをゴールデンテストで固定するのか:
'   会社公式ツールの帳票OCR版は gsPath を引用符で囲み忘れており、ユーザー名に
'   空白が入る端末では起動すらしなかった(docs/dev/reference_公式OCRツール調査_
'   20260731.md §5)。引用符は目で見て確かめにくく、実機でしか症状が出ない。
'   だから「空白と日本語を含むパス」で1文字単位に固定する。
'
' ■ CanUseTypeArraysの複製について(modTestsPure2/3の冒頭コメントと同じ理由):
'   他モジュールのPrivateは呼べないため、3行程度の軽量な実測プローブを複製する。
'   ExtractedPage()配列を使うグループ(ページ付きテキスト)だけをこれで守る。
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

' 二重引用符で囲む(期待値の組み立て用。実装側のQuotedとは別物)。
Private Function Dq(ByVal s As String) As String
    Dq = Chr$(34) & s & Chr$(34)
End Function

' ----------------------------------------------------------------------------
' GSコマンドのゴールデンテスト
' ----------------------------------------------------------------------------
Private Sub TestBuildGsCommandGolden()
    Dim gsExe As String: gsExe = "C:\Program Files\Nexus Agent\Ghostscript\gswin32c.exe"
    Dim pdfPath As String: pdfPath = "C:\共有\契約 書類\画像PDF テスト.pdf"
    Dim outPattern As String: outPattern = "C:\Temp\nxocr_20260731_101112_4321\page_%03d.jpg"

    Dim expected As String
    expected = Dq(gsExe) & _
        " -dSAFER -dNOPAUSE -dBATCH" & _
        " -sDEVICE=jpeg -r150" & _
        " -dTextAlphaBits=4 -dGraphicsAlphaBits=4" & _
        " -dFirstPage=1 -dLastPage=21" & _
        " -sOutputFile=" & Dq(outPattern) & _
        " " & Dq(pdfPath)

    Dim actual As String
    actual = optOcrCore.BuildGsCommand(gsExe, pdfPath, outPattern, 150, 21)

    modTestRunner.Check "GSコマンド: 空白と日本語を含むパスでゴールデン一致", _
        (actual = expected), "実際=" & actual

    ' 公式ツールが踏んだバグ(gsPathの引用符漏れ)を機械で止める。
    modTestRunner.Check "GSコマンド: 実行ファイルのパスが引用符で始まる", _
        (Left$(actual, 1) = Chr$(34)), "実際=" & Left$(actual, 40)
    modTestRunner.Check "GSコマンド: PDFのパスが引用符で終わる", _
        (Right$(actual, 1) = Chr$(34)), "実際=" & Right$(actual, 40)
    modTestRunner.Check "GSコマンド: 必須オプションが全て入っている", _
        (InStr(actual, "-dSAFER") > 0 And InStr(actual, "-dNOPAUSE") > 0 And _
         InStr(actual, "-dBATCH") > 0 And InStr(actual, "-sDEVICE=jpeg") > 0 And _
         InStr(actual, "-dTextAlphaBits=4") > 0 And InStr(actual, "-dGraphicsAlphaBits=4") > 0 And _
         InStr(actual, "-dFirstPage=1") > 0 And InStr(actual, "-dLastPage=21") > 0), "実際=" & actual

    ' dpiは範囲外(0や10000)なら既定150へ丸める。lastPageは1未満なら1。
    modTestRunner.Check "GSコマンド: dpi=0は既定150へ丸める", _
        (InStr(optOcrCore.BuildGsCommand("g", "p", "o", 0, 5), " -r150 ") > 0), _
        "実際=" & optOcrCore.BuildGsCommand("g", "p", "o", 0, 5)
    modTestRunner.Check "GSコマンド: dpi=300はそのまま使う", _
        (InStr(optOcrCore.BuildGsCommand("g", "p", "o", 300, 5), " -r300 ") > 0), ""
    modTestRunner.Check "GSコマンド: lastPage=0は1へ丸める", _
        (InStr(optOcrCore.BuildGsCommand("g", "p", "o", 150, 0), "-dLastPage=1 ") > 0), ""
End Sub

Private Sub TestBuildRunCommandGolden()
    Dim gsCommand As String: gsCommand = Dq("C:\gs\gswin32c.exe") & " -dSAFER " & Dq("C:\a b\in.pdf")
    Dim flagPath As String: flagPath = "C:\Temp\nxocr_1\done.flag"

    Dim expected As String
    expected = "cmd.exe /s /c " & Chr$(34) & gsCommand & " & echo done>" & Dq(flagPath) & Chr$(34)

    Dim actual As String
    actual = optOcrCore.BuildRunCommand(gsCommand, flagPath)

    modTestRunner.Check "実行コマンド: cmd.exe /s /c ラップのゴールデン一致", _
        (actual = expected), "実際=" & actual
    modTestRunner.Check "実行コマンド: 完了フラグの作成が無条件(&)で連結されている", _
        (InStr(actual, " & echo done>") > 0), "実際=" & actual
    modTestRunner.Check "実行コマンド: 全体が二重引用符で閉じている", _
        (Right$(actual, 1) = Chr$(34)), "実際=" & Right$(actual, 30)
End Sub

' ----------------------------------------------------------------------------
' 一時フォルダ・出力ファイル名の規約
' ----------------------------------------------------------------------------
Private Sub TestOcrPathRules()
    modTestRunner.Check "一時フォルダ: TEMP配下にnxocr_接頭辞で作る", _
        (optOcrCore.TempFolderFor("C:\Users\x\Temp", "20260731_101112_4321") = _
         "C:\Users\x\Temp\nxocr_20260731_101112_4321"), _
        "実際=" & optOcrCore.TempFolderFor("C:\Users\x\Temp", "20260731_101112_4321")
    modTestRunner.Check "一時フォルダ: 末尾の区切り記号が重複しない", _
        (optOcrCore.TempFolderFor("C:\Users\x\Temp\", "a") = "C:\Users\x\Temp\nxocr_a"), _
        "実際=" & optOcrCore.TempFolderFor("C:\Users\x\Temp\", "a")
    modTestRunner.Check "出力パターン: page_%03d.jpg(公式ツールと同じ規約)", _
        (optOcrCore.OutPatternFor("C:\Temp\nxocr_a") = "C:\Temp\nxocr_a\page_%03d.jpg"), _
        "実際=" & optOcrCore.OutPatternFor("C:\Temp\nxocr_a")
    modTestRunner.Check "完了フラグ: フォルダ直下のdone.flag", _
        (optOcrCore.DoneFlagFor("C:\Temp\nxocr_a") = "C:\Temp\nxocr_a\done.flag"), _
        "実際=" & optOcrCore.DoneFlagFor("C:\Temp\nxocr_a")

    modTestRunner.Check "ページ画像名: 1枚目はpage_001.jpg", _
        (optOcrCore.PageJpgName(1) = "page_001.jpg"), "実際=" & optOcrCore.PageJpgName(1)
    modTestRunner.Check "ページ画像名: 10枚目はpage_010.jpg", _
        (optOcrCore.PageJpgName(10) = "page_010.jpg"), "実際=" & optOcrCore.PageJpgName(10)
    modTestRunner.Check "ページ画像名: 999枚目はpage_999.jpg", _
        (optOcrCore.PageJpgName(999) = "page_999.jpg"), "実際=" & optOcrCore.PageJpgName(999)
    modTestRunner.Check "ページ画像名: 1000枚目は桁が増える(%03dと同じ挙動)", _
        (optOcrCore.PageJpgName(1000) = "page_1000.jpg"), "実際=" & optOcrCore.PageJpgName(1000)
End Sub

' ----------------------------------------------------------------------------
' ページ上限と切り詰め判定の境界
'   総ページ数を知る手段がGS抜きには無いので、上限+1ページまで描かせて
'   「上限+1枚できていたら切り詰めあり」と判定する。境界を1枚間違えると
'   最後のページが黙って消えるか、余分な1枚をOCRして課金だけ増える。
' ----------------------------------------------------------------------------
Private Sub TestRenderCapAndTruncate()
    modTestRunner.Check "描画上限: 既定20なら21ページまで描く", _
        (optOcrCore.RenderCapFor(20) = 21), "実際=" & optOcrCore.RenderCapFor(20)
    modTestRunner.Check "描画上限: 0以下は1ページ扱い(=2まで描く)", _
        (optOcrCore.RenderCapFor(0) = 2), "実際=" & optOcrCore.RenderCapFor(0)
    modTestRunner.Check "描画上限: 極端な値は200で頭打ち", _
        (optOcrCore.RenderCapFor(99999) = 201), "実際=" & optOcrCore.RenderCapFor(99999)

    modTestRunner.Check "切り詰め判定: 21枚できたら切り詰めあり(上限20)", _
        (optOcrCore.IsTruncatedCount(21, 20) = True), ""
    modTestRunner.Check "切り詰め判定: ちょうど20枚なら切り詰めなし", _
        (optOcrCore.IsTruncatedCount(20, 20) = False), ""
    modTestRunner.Check "切り詰め判定: 19枚なら切り詰めなし", _
        (optOcrCore.IsTruncatedCount(19, 20) = False), ""

    modTestRunner.Check "OCR対象枚数: 21枚できても20枚で頭打ち(余分な1枚は捨てる)", _
        (optOcrCore.KeepPageCount(21, 20) = 20), "実際=" & optOcrCore.KeepPageCount(21, 20)
    modTestRunner.Check "OCR対象枚数: 5枚なら5枚", _
        (optOcrCore.KeepPageCount(5, 20) = 5), "実際=" & optOcrCore.KeepPageCount(5, 20)
    modTestRunner.Check "OCR対象枚数: 0枚は0枚", _
        (optOcrCore.KeepPageCount(0, 20) = 0), ""

    modTestRunner.Check "dpi丸め: 範囲外は既定150", _
        (optOcrCore.SafeDpi(0) = 150 And optOcrCore.SafeDpi(9999) = 150 And optOcrCore.SafeDpi(300) = 300), ""
    modTestRunner.Check "ページ上限丸め: 1未満は1・200超は200", _
        (optOcrCore.SafeMaxPages(0) = 1 And optOcrCore.SafeMaxPages(500) = 200 And _
         optOcrCore.SafeMaxPages(20) = 20), ""
End Sub

' ----------------------------------------------------------------------------
' ページ付きテキストの往復(0/1/Nページ・truncatedあり/なし・マーカー風の本文)
'   ここが壊れると、OCRできた本文が「1ページに全部くっついた」形で本棚に入り、
'   出典のページ番号が全部1になる(誰も気づけない壊れ方)。
' ----------------------------------------------------------------------------
Private Sub TestPagedTextRoundTrip()
    If Not CanUseTypeArrays() Then
        modTestRunner.Check "ページ付きテキスト: 実行環境の制限によりスキップ", True, _
            "Public Type配列(ExtractedPage())のReDimが使えない環境。Excel実機受入で確認すること。"
        Exit Sub
    End If

    ' --- Nページ + truncated あり ---
    Dim src(0 To 2) As ExtractedPage
    src(0).page = 1
    src(0).Text = "1ページ目の見出し" & vbLf & "本文A"
    src(1).page = 2
    src(1).Text = ""
    src(2).page = 7
    src(2).Text = "最後のページ"

    Dim encoded As String
    encoded = modUtil.JoinPagedText(src, True)

    modTestRunner.Check "符号化: 打ち切りマーカーが先頭行に付く", _
        (Left$(encoded, 19) = "@@NEXUS_TRUNCATED@@"), "実際=" & Left$(encoded, 30)
    modTestRunner.Check "符号化: 実ページ番号がマーカーに入る", _
        (InStr(encoded, "@@NEXUS_PAGE:7@@") > 0), "実際=" & encoded

    Dim back() As ExtractedPage
    Dim wasCut As Boolean
    Dim decoded As Boolean
    decoded = modUtil.SplitPagedText(encoded, back, wasCut)

    modTestRunner.Check "復号: 成功する", decoded, ""
    modTestRunner.Check "復号: truncatedがTrueで返る", wasCut, ""
    modTestRunner.Check "復号: ページ数が3", _
        (UBound(back) - LBound(back) + 1 = 3), "実際=" & (UBound(back) - LBound(back) + 1)
    modTestRunner.Check "復号: 1ページ目の本文が改行ごと一致", _
        (back(0).page = 1 And back(0).Text = "1ページ目の見出し" & vbLf & "本文A"), _
        "実際=" & back(0).page & "/" & back(0).Text
    modTestRunner.Check "復号: 空ページも欠けずに残る", _
        (back(1).page = 2 And back(1).Text = ""), "実際=" & back(1).page & "/" & back(1).Text
    modTestRunner.Check "復号: 飛び番のページ番号(7)が保たれる", _
        (back(2).page = 7 And back(2).Text = "最後のページ"), _
        "実際=" & back(2).page & "/" & back(2).Text

    ' --- 1ページ + truncated なし ---
    Dim one(0 To 0) As ExtractedPage
    one(0).page = 1
    one(0).Text = "ただ1ページ"
    Dim enc1 As String: enc1 = modUtil.JoinPagedText(one)
    Dim back1() As ExtractedPage
    Dim cut1 As Boolean
    modTestRunner.Check "1ページ往復: 復号できる", _
        modUtil.SplitPagedText(enc1, back1, cut1), "符号化=" & enc1
    modTestRunner.Check "1ページ往復: truncatedはFalse", (cut1 = False), ""
    modTestRunner.Check "1ページ往復: 本文が一致", _
        (back1(0).Text = "ただ1ページ" And back1(0).page = 1), "実際=" & back1(0).Text

    ' --- 0ページ ---
    Dim none() As ExtractedPage
    modTestRunner.Check "0ページ: 符号化は空文字", (modUtil.JoinPagedText(none) = ""), ""
    Dim back0() As ExtractedPage
    Dim cut0 As Boolean
    modTestRunner.Check "0ページ: 空文字の復号はFalse", _
        (modUtil.SplitPagedText("", back0, cut0) = False), ""
    modTestRunner.Check "ページ無しの普通の全文はFalse(呼び出し元が1ページ扱いにする)", _
        (modUtil.SplitPagedText("ページマーカーの無いただの本文", back0, cut0) = False), ""
End Sub

' 本文にマーカーそっくりの行が混ざっても壊れない(最悪でもページが割れるだけで
' 本文は1文字も失わない)。OCR結果は何が返ってくるか分からないので、ここは
' 「落ちない・消えない」ことを固定しておく。
Private Sub TestPagedTextMarkerLikeBody()
    If Not CanUseTypeArrays() Then
        modTestRunner.Check "マーカー風の本文: 実行環境の制限によりスキップ", True, ""
        Exit Sub
    End If

    Dim src(0 To 0) As ExtractedPage
    src(0).page = 1
    src(0).Text = "本文の途中に" & vbLf & "@@NEXUS_PAGE:9@@" & vbLf & "マーカー風の行がある"

    Dim back() As ExtractedPage
    Dim wasCut As Boolean
    Dim ok As Boolean
    ok = modUtil.SplitPagedText(modUtil.JoinPagedText(src), back, wasCut)

    modTestRunner.Check "マーカー風の本文: 復号は成功する(落ちない)", ok, ""

    Dim total As String
    Dim i As Long
    If ok Then
        For i = LBound(back) To UBound(back)
            total = total & back(i).Text
        Next i
    End If
    modTestRunner.Check "マーカー風の本文: 本文の文字が消えない", _
        (InStr(total, "本文の途中に") > 0 And InStr(total, "マーカー風の行がある") > 0), _
        "実際=" & total

    ' マーカーとして解釈されるのは「行全体が完全一致」した場合だけ。
    Dim src2(0 To 0) As ExtractedPage
    src2(0).page = 1
    src2(0).Text = "これは @@NEXUS_PAGE:3@@ を含むが行全体ではない"
    Dim back2() As ExtractedPage
    Dim cut2 As Boolean
    modTestRunner.Check "マーカー風の本文: 行の途中のマーカーでは割れない", _
        (modUtil.SplitPagedText(modUtil.JoinPagedText(src2), back2, cut2) And _
         (UBound(back2) - LBound(back2) + 1) = 1), ""
End Sub

' ----------------------------------------------------------------------------
' R6追補: 「全ページ化け」のPDFをOCR経路へ回す分岐
'   実機ログでは、画像PDFの典型症状は E0303(文字が少ない)ではなく
'   「WordリフローがゴミWord文字を返し、全ページが化け判定になる」形だった。
'   PDFだけをOCR経路へ回し、doc/docx等は従来どおり続行する。
' ----------------------------------------------------------------------------
Private Sub TestGarbledRouteCode()
    modTestRunner.Check "全ページ化け: pdfはE0303へ振り替えてOCR経路へ", _
        (modExtractor.GarbledRouteCode("pdf", True) = "E0303"), _
        "実際=" & modExtractor.GarbledRouteCode("pdf", True)
    modTestRunner.Check "全ページ化け: 拡張子の大文字小文字は問わない", _
        (modExtractor.GarbledRouteCode("PDF", True) = "E0303"), ""
    modTestRunner.Check "全ページ化け: docxは従来どおり続行(GSはPDF専用)", _
        (modExtractor.GarbledRouteCode("docx", True) = ""), _
        "実際=" & modExtractor.GarbledRouteCode("docx", True)
    modTestRunner.Check "一部だけ化け: 振り替えない(v1は除外のみ)", _
        (modExtractor.GarbledRouteCode("pdf", False) = ""), ""
End Sub

' ----------------------------------------------------------------------------
' R7 B-1: 進捗の目安時間(modUtil.EtaText / modUtil.ProgressText)
'   実機報告「チャンク化中の進捗に目安時間が無い」への対応。
'   時間の見積りは「当たらない予告を出さない」ことが要件なので、
'   ①まだ測れていないときは何も言わない ②1分未満は数字を出さない
'   の2点を回帰として固定する。
' ----------------------------------------------------------------------------
Private Sub TestEtaText()
    modTestRunner.Check "ETA: 実測前(msPerItem=0)は何も言わない", _
        (modUtil.EtaText(50, 0) = ""), "実際=[" & modUtil.EtaText(50, 0) & "]"
    modTestRunner.Check "ETA: 残り0件は何も言わない", _
        (modUtil.EtaText(0, 500) = ""), "実際=[" & modUtil.EtaText(0, 500) & "]"
    modTestRunner.Check "ETA: 残り件数が負でも何も言わない", _
        (modUtil.EtaText(-3, 500) = ""), ""
    ' 10件 × 1000ms = 10秒 → 秒数は出さない
    modTestRunner.Check "ETA: 60秒未満は「まもなく完了」", _
        (modUtil.EtaText(10, 1000) = "まもなく完了"), "実際=" & modUtil.EtaText(10, 1000)
    ' 76件 × 2000ms = 152秒 = 2.53分 → 四捨五入で3分
    modTestRunner.Check "ETA: 152秒は「残り約3分」(四捨五入)", _
        (modUtil.EtaText(76, 2000) = "残り約3分"), "実際=" & modUtil.EtaText(76, 2000)
    ' 60件 × 1000ms = 60秒ちょうど → 1分
    modTestRunner.Check "ETA: 60秒ちょうどは「残り約1分」", _
        (modUtil.EtaText(60, 1000) = "残り約1分"), "実際=" & modUtil.EtaText(60, 1000)
    ' 3600件 × 1000ms = 3600秒 = 60分 → 1時間(端数なしは「分」を付けない)
    modTestRunner.Check "ETA: ちょうど60分は「残り約1時間」", _
        (modUtil.EtaText(3600, 1000) = "残り約1時間"), "実際=" & modUtil.EtaText(3600, 1000)
    ' 4500件 × 1000ms = 75分 → 1時間15分
    modTestRunner.Check "ETA: 75分は「残り約1時間15分」", _
        (modUtil.EtaText(4500, 1000) = "残り約1時間15分"), "実際=" & modUtil.EtaText(4500, 1000)
End Sub

Private Sub TestProgressText()
    modTestRunner.Check "進捗: 見積り無しなら件数だけ", _
        (modUtil.ProgressText(12, 88, "") = "12/88件"), _
        "実際=" & modUtil.ProgressText(12, 88, "")
    modTestRunner.Check "進捗: 要件の書式「12/88件・残り約3分」", _
        (modUtil.ProgressText(12, 88, "残り約3分") = "12/88件・残り約3分"), _
        "実際=" & modUtil.ProgressText(12, 88, "残り約3分")
    ' modEmbedが実際に組み立てる形(残り76件・1件2秒)を通しで固定する。
    modTestRunner.Check "進捗: EtaTextとの結線(76件×2秒→3分)", _
        (modUtil.ProgressText(12, 88, modUtil.EtaText(76, 2000)) = "12/88件・残り約3分"), _
        "実際=" & modUtil.ProgressText(12, 88, modUtil.EtaText(76, 2000))
End Sub

Public Sub RunAll4()
    On Error GoTo GsGoldenFail
    TestBuildGsCommandGolden
NextRunCmd:
    On Error GoTo RunCmdFail
    TestBuildRunCommandGolden
NextPathRules:
    On Error GoTo PathRulesFail
    TestOcrPathRules
NextRenderCap:
    On Error GoTo RenderCapFail
    TestRenderCapAndTruncate
NextPagedRt:
    On Error GoTo PagedRtFail
    TestPagedTextRoundTrip
NextMarkerLike:
    On Error GoTo MarkerLikeFail
    TestPagedTextMarkerLikeBody
NextGarbledRoute:
    On Error GoTo GarbledRouteFail
    TestGarbledRouteCode
NextEta:
    On Error GoTo EtaFail
    TestEtaText
NextProgress:
    On Error GoTo ProgressFail
    TestProgressText
NextDone4:
    On Error GoTo 0
    Exit Sub

GsGoldenFail:
    modTestRunner.Check "TestBuildGsCommandGolden(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRunCmd
RunCmdFail:
    modTestRunner.Check "TestBuildRunCommandGolden(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPathRules
PathRulesFail:
    modTestRunner.Check "TestOcrPathRules(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRenderCap
RenderCapFail:
    modTestRunner.Check "TestRenderCapAndTruncate(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPagedRt
PagedRtFail:
    modTestRunner.Check "TestPagedTextRoundTrip(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextMarkerLike
MarkerLikeFail:
    modTestRunner.Check "TestPagedTextMarkerLikeBody(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGarbledRoute
GarbledRouteFail:
    modTestRunner.Check "TestGarbledRouteCode(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextEta
EtaFail:
    modTestRunner.Check "TestEtaText(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextProgress
ProgressFail:
    modTestRunner.Check "TestProgressText(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone4
End Sub
