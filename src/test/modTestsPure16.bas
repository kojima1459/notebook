Attribute VB_Name = "modTestsPure16"
Option Explicit

' ============================================================================
' modTestsPure16 - R18-1/R18-2/R18-6/R18-7(実機第5報①⑤⑦⑧⑪)の純ロジック回帰
' ----------------------------------------------------------------------------
' なぜ新設したか(憲章§4-6):
'   modTestsPure15 が23,011字で、ここの真理表(約9,300字)を足すと30,000字
'   上限を超える。15を新設したときと同じ線で分割する。
'   入口は modTestsPure15.RunAll15 の末尾から呼ばれる RunAll16 の1本だけ。
'
' ここで固定するもの:
'   ・modProgressBar.BarWidthFor(1b)/BarHeightFor(R18H FA-2): 進捗バナーの
'     幅の viewport 連動と、文面が収まらないときの2行化。
'   ・modIntegrity.ReconcileStatText / IndexOfName(2b): 台帳と実データの突合。
'   ・modIntegrity.DataShrunk / IsVolatilePath / 警告文(2d): 起動時の突合。
'   ・modShelfScan.EnumLooksFailed(2f): UNC列挙の途中切れを消失と誤判定しない。
'   ・MsgBox到達文言の非BMP絵文字除去(6a): OcrConfirmAskFor/FriendlyMessage
'     (E0201/E0204/E0705)の戻り値を通しで回帰確認(vba_lint検査14の実行時側)。
'   ・modRagParse.HasCompoundSignal(7a)と、modAskMulti.TryDecomposedの
'     ゲート合成(ShouldDecompose OR HasCompoundSignal・7b)。
'   ・modViewport.ColLetter(3b): ScrollAreaの範囲文字列を組む列番号→列名。
' ============================================================================

' ----------------------------------------------------------------------------
' R18-1b: 進捗バナーの幅(modProgressBar.BarWidthFor)。
' ----------------------------------------------------------------------------
' 従来は固定380ptで、本文の可視幅は 380-左16-右190 = 174pt しか無く、OCRの
' 実況文(約363pt)が折り返して高さ30ptのピルから溢れ、地色との対比1.05:1で
' 完全に読めなかった(実機第5報①)。viewport-16 と上限760 の小さい方にすれば
' 本文可視幅は 760-16-190 = 554pt 以上になり、現行の文面が1行に収まる。
' 狭い窓では viewport-16 まで縮み、■中断(右端 leftPos+barW-6)が画面外へ
' はみ出さない。下限300は「バナーの体をなす最小」。
Private Sub TestBarWidthFor()
    modTestRunner.Check "バナー幅_広い窓は上限760で頭打ち", _
        (modProgressBar.BarWidthFor(1600) = 760), _
        "実際=" & modProgressBar.BarWidthFor(1600)
    modTestRunner.Check "バナー幅_776でちょうど760(境界)", _
        (modProgressBar.BarWidthFor(776) = 760)
    modTestRunner.Check "バナー幅_775は759(上限直下はviewport連動)", _
        (modProgressBar.BarWidthFor(775) = 759)
    modTestRunner.Check "バナー幅_可視幅600なら584", _
        (modProgressBar.BarWidthFor(600) = 584)
    modTestRunner.Check "バナー幅_狭い窓(320)は304", _
        (modProgressBar.BarWidthFor(320) = 304)
    modTestRunner.Check "バナー幅_極端に狭くても300を割らない", _
        (modProgressBar.BarWidthFor(100) = 300)
    ' 本文可視幅(barW - 左余白16 - 右余白190)が554pt以上=写真の文面が1行。
    modTestRunner.Check "バナー幅_上限時の本文可視幅は554pt", _
        (modProgressBar.BarWidthFor(1366) - 16 - 190 = 554)
End Sub

' ----------------------------------------------------------------------------
' R18H FA-2: 進捗バナーの高さ(modProgressBar.BarHeightFor)。
' ----------------------------------------------------------------------------
' 幅だけを viewport 連動にしても、高さが1行(30pt)固定のままだと、狭い窓で
' 本文可視幅が縮んだ瞬間に文字が折り返してピルから溢れ、R18-1b 以前と同じ
' 「読めない実況」に戻る(B-H1/A-L11)。文字数×係数の近似で【過大側に】
' 見積もり、収まらないと分かった時点で2行(46pt)へ広げる。
' 可視幅554pt = BarWidthFor(1366) - 左16 - 右190(ボタン込み)。
Private Sub TestBarHeightFor()
    modTestRunner.Check "バナー高さ_空文字は1行", _
        (modProgressBar.BarHeightFor("", 554) = 30)
    modTestRunner.Check "バナー高さ_全角52字は1行に収まる(境界)", _
        (modProgressBar.BarHeightFor(String$(52, "あ"), 554) = 30)
    modTestRunner.Check "バナー高さ_全角53字で2行へ(境界)", _
        (modProgressBar.BarHeightFor(String$(53, "あ"), 554) = 46)
    modTestRunner.Check "バナー高さ_半角は全角より狭く数える", _
        (modProgressBar.BarHeightFor(String$(100, "a"), 554) = 30)
    modTestRunner.Check "バナー高さ_可視幅が狭ければ短文でも2行", _
        (modProgressBar.BarHeightFor(String$(20, "あ"), 94) = 46)
    modTestRunner.Check "バナー高さ_可視幅0以下は1行へ倒す(退化入力の防御)", _
        (modProgressBar.BarHeightFor(String$(20, "あ"), 0) = 30)
    ' 実機のOCR実況文(残り時間+終了目安)が上限幅では1行に収まること。
    modTestRunner.Check "バナー高さ_OCR実況文は上限幅で1行", _
        (modProgressBar.BarHeightFor("OCR中 12/254頁 残り約38分(終了目安 15:42)", _
            modProgressBar.BarWidthFor(1366) - 16 - 190) = 30)
End Sub

' ----------------------------------------------------------------------------
' R18-2b: 台帳の統計文字列の chunk_count 差し替え(modIntegrity)。
' ----------------------------------------------------------------------------
' statText は modShelf.SourceList が組み立てた
' "status|ingested_at|chunk_count|error_note|origin"。実行数と一致していれば
' 【同じ文字列をそのまま返す】ことが呼び出し側の「何もしない」合図なので、
' 一致時に素通しであることまで見る。error_note に "|" が混じる可能性が
' あるため、6つ以上に割れる入力でも3つ目だけが変わることを固定する。
Private Sub TestReconcileStatText()
    modTestRunner.Check "突合_一致なら素通し", _
        (modIntegrity.ReconcileStatText("done|2026-08-05|127||self", 127) = _
         "done|2026-08-05|127||self")
    modTestRunner.Check "突合_0上書きを実行数へ戻す", _
        (modIntegrity.ReconcileStatText("image_pdf|2026-08-05|0|中断しました|self", 127) = _
         "image_pdf|2026-08-05|127|中断しました|self")
    modTestRunner.Check "突合_実データが減っていれば減らす", _
        (modIntegrity.ReconcileStatText("done|2026-08-05|127||self", 0) = _
         "done|2026-08-05|0||self")
    modTestRunner.Check "突合_空欄のchunk_countも直す", _
        (modIntegrity.ReconcileStatText("failed|2026-08-05|||self", 5) = _
         "failed|2026-08-05|5||self")
    modTestRunner.Check "突合_メモに区切り文字が混じっても他節を壊さない", _
        (modIntegrity.ReconcileStatText("failed|t|0|a|b|self", 9) = "failed|t|9|a|b|self")
    modTestRunner.Check "突合_節が足りない壊れた文字列は触らない", _
        (modIntegrity.ReconcileStatText("done|2026-08-05|0", 7) = "done|2026-08-05|0")
    modTestRunner.Check "突合_空文字は触らない", _
        (modIntegrity.ReconcileStatText("", 3) = "")
End Sub

' ----------------------------------------------------------------------------
' R18-2b: source別集計の位置引き(modIntegrity.IndexOfName)。
' ----------------------------------------------------------------------------
' hint は「速さだけの助言」で、当たっても外れても結果は線形探索と同じでなければ
' ならない(ここがズレると集計が別の資料の行数を数え始めて台帳を壊す)。
Private Sub TestIndexOfName()
    Dim nm(0 To 2) As String
    nm(0) = "a.pdf": nm(1) = "b.pdf": nm(2) = "c.pdf"
    modTestRunner.Check "位置引き_先頭", (modIntegrity.IndexOfName(nm, 3, "a.pdf", -1) = 0)
    modTestRunner.Check "位置引き_末尾", (modIntegrity.IndexOfName(nm, 3, "c.pdf", -1) = 2)
    modTestRunner.Check "位置引き_不在は-1", (modIntegrity.IndexOfName(nm, 3, "z.pdf", -1) = -1)
    modTestRunner.Check "位置引き_大小無視", (modIntegrity.IndexOfName(nm, 3, "B.PDF", -1) = 1)
    modTestRunner.Check "位置引き_hint的中", (modIntegrity.IndexOfName(nm, 3, "b.pdf", 1) = 1)
    modTestRunner.Check "位置引き_hint外れでも正しい", _
        (modIntegrity.IndexOfName(nm, 3, "b.pdf", 0) = 1)
    modTestRunner.Check "位置引き_hintが範囲外でも壊れない", _
        (modIntegrity.IndexOfName(nm, 3, "b.pdf", 99) = 1)
    modTestRunner.Check "位置引き_空の目標は-1", (modIntegrity.IndexOfName(nm, 3, "", -1) = -1)
    modTestRunner.Check "位置引き_件数0は-1", (modIntegrity.IndexOfName(nm, 0, "a.pdf", -1) = -1)
    ' n はカードの件数(配列の実長ではない)。パック由来を突合しないための境界。
    modTestRunner.Check "位置引き_nより後ろは見ない", _
        (modIntegrity.IndexOfName(nm, 2, "c.pdf", -1) = -1)
End Sub

' ----------------------------------------------------------------------------
' R18-2d: 起動時突合の2判定(modIntegrity)。
' ----------------------------------------------------------------------------
' DataShrunk: 記録が無い(0以下)ときは判定しない。初回起動や旧ブックからの
'   移行で、根拠なく「資料が消えました」と言わないための最優先の性質。
' IsVolatilePath: zip をダブルクリックすると Windows は
'   %TEMP% の Temp1_<zip名>.zip フォルダへ展開してそこを開く。保存は成功し
'   ReadOnly でもないためアプリからは正常に見えるのに、次回は1件も残らない
'   (調査agent1 §4「検知ゼロ」)。判定材料はパス文字列だけ。
Private Sub TestIntegrityStartup()
    modTestRunner.Check "減少_127から0は警告", (modIntegrity.DataShrunk(127, 0) = True)
    modTestRunner.Check "減少_127から126でも警告", (modIntegrity.DataShrunk(127, 126) = True)
    modTestRunner.Check "減少_同数は警告しない", (modIntegrity.DataShrunk(127, 127) = False)
    modTestRunner.Check "減少_増えていれば警告しない", (modIntegrity.DataShrunk(127, 200) = False)
    modTestRunner.Check "減少_記録なし(0)は判定しない", (modIntegrity.DataShrunk(0, 0) = False)
    modTestRunner.Check "減少_記録が負なら判定しない", (modIntegrity.DataShrunk(-1, 0) = False)

    ' R18H FA-4: 判定対象は ThisWorkbook.Path(フォルダ)で、%TEMP%/%TMP% の
    ' 実値は引数で受け取る(この関数を純ロジックのまま保つため)。
    Dim tA As String: tA = EnvTempDir16()
    modTestRunner.Check "一時_Temp1_zip直開き", _
        (modIntegrity.IsVolatilePath(TmpZipPath16(), tA, "") = True)
    modTestRunner.Check "一時_環境変数TEMPの実値配下", _
        (modIntegrity.IsVolatilePath(TmpDirPath16(), tA, "") = True)
    modTestRunner.Check "一時_TMP側でも拾う", _
        (modIntegrity.IsVolatilePath(TmpDirPath16(), "", tA) = True)
    modTestRunner.Check "一時_大文字でも拾う", _
        (modIntegrity.IsVolatilePath(UCase$(TmpDirPath16()), tA, "") = True)
    modTestRunner.Check "一時_TEMPそのもの(末尾区切り無し)も拾う", _
        (modIntegrity.IsVolatilePath(tA, tA, "") = True)
    modTestRunner.Check "一時_デスクトップは正常", _
        (modIntegrity.IsVolatilePath(DesktopPath16(), tA, "") = False)
    modTestRunner.Check "一時_共有フォルダは正常", _
        (modIntegrity.IsVolatilePath(UncPath16(), tA, "") = False)
    modTestRunner.Check "一時_空パスは判定しない", (modIntegrity.IsVolatilePath("", tA, "") = False)
    ' 名前に templates を含むフォルダを誤検知しないこと(区切り込みで見るため)。
    modTestRunner.Check "一時_templatesフォルダは誤検知しない", _
        (modIntegrity.IsVolatilePath(TemplatesPath16(), tA, "") = False)
    ' R18H FA-4(A-M3/B-M5): "\temp\" の部分一致を廃止した回帰。正規の作業
    ' フォルダ D:\temp\ や共有 \\share\Temp\ を一時扱いにしてはならない
    ' (毎回の嘘の警告は、その機能を二度と使われなくする=憲章§3-3)。
    modTestRunner.Check "一時_D:\temp\の正規フォルダは誤検知しない", _
        (modIntegrity.IsVolatilePath(LocalTempWorkDir16(), tA, "") = False)
    modTestRunner.Check "一時_共有の\\share\Temp\は誤検知しない", _
        (modIntegrity.IsVolatilePath(UncTempDir16(), tA, "") = False)
    ' 1文字違いの隣フォルダ(...\Temp2)を前方一致で巻き込まないこと。
    modTestRunner.Check "一時_TEMPと1文字違いのフォルダは誤検知しない", _
        (modIntegrity.IsVolatilePath(tA & "2" & Sep16() & "a", tA, "") = False)

    ' 警告文はBMPの文字だけ(非BMP絵文字はCP932変換で化けてダイアログに出る)。
    Dim m As String
    m = modIntegrity.ShrinkWarnMsg(127, 0, "X:" & Sep16() & "a.xlsm", "Y:" & Sep16() & "b.xlsm")
    modTestRunner.Check "警告文_前回件数を言う", (InStr(m, "127") > 0)
    modTestRunner.Check "警告文_今の件数を言う", (InStr(m, "0 件") > 0)
    modTestRunner.Check "警告文_前回の場所を出す", (InStr(m, "X:") > 0)
    modTestRunner.Check "警告文_今の場所を出す", (InStr(m, "Y:") > 0)
    modTestRunner.Check "警告文_非BMPを含まない", (HasSurrogate16(m) = False)
    Dim v As String: v = modIntegrity.VolatileWarnMsg(TmpZipPath16())
    modTestRunner.Check "一時警告文_次の一手を言う", (InStr(v, "すべて展開") > 0)
    modTestRunner.Check "一時警告文_非BMPを含まない", (HasSurrogate16(v) = False)
End Sub

' ----------------------------------------------------------------------------
' R18-2f: UNC列挙の健全性(modShelfScan.EnumLooksFailed)。
' ----------------------------------------------------------------------------
' UNCでは Dir$ が途中で切れてもエラーを返さず短いリストや空を返すことがあり、
' その0件が「フォルダが空になった」として消失判定へ流れ、実在するファイルを
' DeleteSource していた(調査agent1 §3(c)・復旧不能)。
' 「0件かつ台帳に資料あり」だけを失敗扱いにする(diskN>0 まで倒すと、利用者が
' 本当に消したファイルの削除が永久に反映されなくなる)。
Private Sub TestEnumLooksFailed()
    modTestRunner.Check "列挙_0件で台帳に資料ありは失敗扱い", _
        (modShelfScan.EnumLooksFailed(0, 12) = True)
    modTestRunner.Check "列挙_0件で台帳も0件なら失敗扱いにしない", _
        (modShelfScan.EnumLooksFailed(0, 0) = False)
    modTestRunner.Check "列挙_1件でも読めていれば失敗扱いにしない", _
        (modShelfScan.EnumLooksFailed(1, 12) = False)
    modTestRunner.Check "列挙_台帳1件が境界", _
        (modShelfScan.EnumLooksFailed(0, 1) = True)
End Sub

' ----------------------------------------------------------------------------
' テスト用ヘルパー
' ----------------------------------------------------------------------------
' パス区切りは ChrW(&H5C) で組み立てる。リテラルの円記号を文字列の末尾に
' 置くとLOの構文チェッカーが継続行と誤読する既知の罠があり、途中でも
' CP932変換の往復で扱いが揺れるため、テストでは常に組み立てる
' (EDGE_CASES §1.3 と同じ考え方)。
Private Function Sep16() As String
    Sep16 = ChrW(&H5C)
End Function

' 以下はいずれも【フォルダ】のパス(R18H FA-4 で判定対象が ThisWorkbook.Path
' へ変わったため、ファイル名は付けない)。
Private Function EnvTempDir16() As String
    EnvTempDir16 = "C:" & Sep16() & "Users" & Sep16() & "x" & Sep16() & "AppData" & _
        Sep16() & "Local" & Sep16() & "Temp"
End Function

Private Function TmpZipPath16() As String
    TmpZipPath16 = EnvTempDir16() & Sep16() & "Temp1_MyBookshelf.zip"
End Function

Private Function TmpDirPath16() As String
    TmpDirPath16 = EnvTempDir16() & Sep16() & "MyBookshelf"
End Function

Private Function DesktopPath16() As String
    DesktopPath16 = "C:" & Sep16() & "Users" & Sep16() & "x" & Sep16() & "Desktop" & _
        Sep16() & "MyBookshelf"
End Function

Private Function UncPath16() As String
    UncPath16 = Sep16() & Sep16() & "pgiofs01" & Sep16() & "share"
End Function

Private Function TemplatesPath16() As String
    TemplatesPath16 = "C:" & Sep16() & "Users" & Sep16() & "x" & Sep16() & "templates"
End Function

' 正規の作業フォルダ(名前がたまたま temp)。誤検知してはならない。
Private Function LocalTempWorkDir16() As String
    LocalTempWorkDir16 = "D:" & Sep16() & "temp" & Sep16() & "約款"
End Function

Private Function UncTempDir16() As String
    UncTempDir16 = Sep16() & Sep16() & "share" & Sep16() & "Temp" & Sep16() & "本棚"
End Function

' 文字列にサロゲート(非BMP=CP932に無い絵文字)が含まれるか。
Private Function HasSurrogate16(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        If c >= &HD800 And c <= &HDFFF Then
            HasSurrogate16 = True
            Exit Function
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' R18-6a回帰: MsgBox到達文言から非BMP絵文字を除去(実機第5報⑤・
'   EDGE_CASES.md §1.3b)。直書きではなく関数の戻り値経由で紛れ込んでいた
'   ケースなので、実際の戻り値を通しで検証する(vba_lint検査14の静的検査を
'   実行時の観点から補完)。
' ----------------------------------------------------------------------------
Private Sub TestR18_6aNonBmpDialogText()
    Dim ask As String: ask = optOcrEta.OcrConfirmAskFor(254, 106, 15)
    modTestRunner.Check "確認文_非BMPを含まない(6a)", (HasSurrogate16(ask) = False)
    modTestRunner.Check "確認文_作業用Excelボタンの案内が残る(6a)", _
        (InStr(ask, "「作業用Excel」ボタン") > 0), "実際=" & ask

    Dim e0201 As String: e0201 = modLog.FriendlyMessage("E0201")
    modTestRunner.Check "E0201_非BMPを含まない(6a)", (HasSurrogate16(e0201) = False)
    modTestRunner.Check "E0201_診断ボタンの案内が残る(6a)", (InStr(e0201, "「診断」ボタン") > 0)

    Dim e0204 As String: e0204 = modLog.FriendlyMessage("E0204")
    modTestRunner.Check "E0204_非BMPを含まない(6a)", (HasSurrogate16(e0204) = False)
    modTestRunner.Check "E0204_同期の案内が残る(6a)", (InStr(e0204, "「同期」") > 0)

    Dim e0705 As String: e0705 = modLog.FriendlyMessage("E0705")
    modTestRunner.Check "E0705_非BMPを含まない(6a)", (HasSurrogate16(e0705) = False)
    modTestRunner.Check "E0705_同期の案内が残る(6a)", (InStr(e0705, "「同期」") > 0)
End Sub

' ----------------------------------------------------------------------------
' R18-7a: 複合質問シグナル検知(modRagParse.HasCompoundSignal)。実機第5報⑦。
'   agent4調査報告§4の表#1〜9(#10は次のゲート合成テストで扱う)。
' ----------------------------------------------------------------------------
Private Sub TestHasCompoundSignal()
    modTestRunner.Check "複合シグナル_全角?2個で検知(本件の再現ケース)", _
        (modRagParse.HasCompoundSignal("免責は？保険料は？") = True)
    modTestRunner.Check "複合シグナル_AとBの違いは未検知(既知のギャップ)", _
        (modRagParse.HasCompoundSignal("AとBの違いは") = False)
    modTestRunner.Check "複合シグナル_句点区切り3節(末尾空要素を除外)", _
        (modRagParse.HasCompoundSignal("保険料。免責。テロ。") = True)
    modTestRunner.Check "複合シグナル_?1個のみは未検知(真陰性)", _
        (modRagParse.HasCompoundSignal("これって対象?") = False)
    modTestRunner.Check "複合シグナル_空文字はFalse", (modRagParse.HasCompoundSignal("") = False)
    modTestRunner.Check "複合シグナル_?のみ3個でもTrue", _
        (modRagParse.HasCompoundSignal("？？？") = True)
    modTestRunner.Check "複合シグナル_半角?2個も検知(全角半角混在)", _
        (modRagParse.HasCompoundSignal("免責は?保険料は?") = True)
    modTestRunner.Check "複合シグナル_単一文+末尾句点1個は過検知しない", _
        (modRagParse.HasCompoundSignal("免責について教えてください。") = False)
    modTestRunner.Check "複合シグナル_読点は対象外(仕様)", _
        (modRagParse.HasCompoundSignal("免責、保険料について") = False)
End Sub

' ----------------------------------------------------------------------------
' R18-7b: modAskMulti.TryDecomposed のゲート合成(ShouldDecompose OR
'   HasCompoundSignal)。TryDecomposed自体はLLM呼び出し(段0)を伴い純関数の
'   枠外なので、実装と同じ式(modAskMulti.bas:96-107相当)を独立に組んで
'   検証する。agent4調査報告§4の表#10〜14。
' ----------------------------------------------------------------------------
Private Function DecGate16(ByVal modeCfg As String, ByVal q As String) As Boolean
    Dim g As Boolean
    g = modAskMulti.ShouldDecompose(modeCfg, Len(q), 25)
    If Not g And LCase$(Trim$(modeCfg)) <> "off" Then g = modRagParse.HasCompoundSignal(q)
    DecGate16 = g
End Function

Private Sub TestDecomposeGateOr()
    modTestRunner.Check "分解ゲート_auto/短い/複合ありは呼ぶ(修正の主眼)", _
        (DecGate16("auto", "免責は？保険料は？") = True)
    modTestRunner.Check "分解ゲート_auto/短い/複合なしは呼ばない(現状維持)", _
        (DecGate16("auto", "これって対象?") = False)
    modTestRunner.Check "分解ゲート_off/複合ありでも呼ばない(エスケープハッチ優先)", _
        (DecGate16("off", "免責は？保険料は？") = False)
    modTestRunner.Check "分解ゲート_always/複合なしでも呼ぶ(既存動作の非退行)", _
        (DecGate16("always", "これって対象?") = True)
    modTestRunner.Check "分解ゲート_25字以上は複合シグナルに関わらず呼ぶ(既存ゲート健全性)", _
        (DecGate16("auto", String$(25, "あ")) = True)
End Sub

' ----------------------------------------------------------------------------
' R18-3b: 列番号→列名(modViewport.ColLetter)。
' ----------------------------------------------------------------------------
' ScrollAreaの範囲文字列("A1:<列><行>")を組み立てる唯一の算数。ここを1つ
' 間違えると、範囲が意図より狭くなってボタンが境界の外へ取り残される
' (憲章§3-1「押せるものは必ず反応する」違反=見えない/押せない)。
' 26進の繰り上がり(Z→AA)は off-by-one を作りやすいので境界を固定する。
' 0以下は "A" に丸める(壊れた値でRangeを落とさない防御)。
Private Sub TestColLetter()
    modTestRunner.Check "列名_1はA", (modViewport.ColLetter(1) = "A")
    modTestRunner.Check "列名_12はL(Hubの右端)", (modViewport.ColLetter(12) = "L")
    modTestRunner.Check "列名_14はN(マイ本棚の右端)", (modViewport.ColLetter(14) = "N")
    modTestRunner.Check "列名_16はP(チャットの右端)", (modViewport.ColLetter(16) = "P")
    modTestRunner.Check "列名_20はT(ダッシュボードの右端)", (modViewport.ColLetter(20) = "T")
    modTestRunner.Check "列名_26はZ(繰り上がり直前)", (modViewport.ColLetter(26) = "Z")
    modTestRunner.Check "列名_27はAA(繰り上がり)", (modViewport.ColLetter(27) = "AA")
    modTestRunner.Check "列名_52はAZ", (modViewport.ColLetter(52) = "AZ")
    modTestRunner.Check "列名_53はBA", (modViewport.ColLetter(53) = "BA")
    modTestRunner.Check "列名_0以下はAへ丸める", (modViewport.ColLetter(0) = "A")
    modTestRunner.Check "列名_負値もAへ丸める", (modViewport.ColLetter(-5) = "A")
End Sub

Public Sub RunAll16()
    On Error GoTo BarWFail16
    TestBarWidthFor
NextBarH16:
    On Error GoTo BarHFail16
    TestBarHeightFor
NextRecon16:
    On Error GoTo ReconFail16
    TestReconcileStatText
NextIdx16:
    On Error GoTo IdxFail16
    TestIndexOfName
NextStart16:
    On Error GoTo StartFail16
    TestIntegrityStartup
NextEnum16:
    On Error GoTo EnumFail16
    TestEnumLooksFailed
NextNonBmp16:
    On Error GoTo NonBmpFail16
    TestR18_6aNonBmpDialogText
NextHcs16:
    On Error GoTo HcsFail16
    TestHasCompoundSignal
NextGate16:
    On Error GoTo GateFail16
    TestDecomposeGateOr
NextCol16:
    On Error GoTo ColFail16
    TestColLetter
NextDone16:
    On Error GoTo 0
    Exit Sub

BarWFail16:
    modTestRunner.Check "TestBarWidthFor(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBarH16
BarHFail16:
    modTestRunner.Check "TestBarHeightFor(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRecon16
ReconFail16:
    modTestRunner.Check "TestReconcileStatText(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextIdx16
IdxFail16:
    modTestRunner.Check "TestIndexOfName(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextStart16
StartFail16:
    modTestRunner.Check "TestIntegrityStartup(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextEnum16
EnumFail16:
    modTestRunner.Check "TestEnumLooksFailed(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextNonBmp16
NonBmpFail16:
    modTestRunner.Check "TestR18_6aNonBmpDialogText(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextHcs16
HcsFail16:
    modTestRunner.Check "TestHasCompoundSignal(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGate16
GateFail16:
    modTestRunner.Check "TestDecomposeGateOr(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCol16
ColFail16:
    modTestRunner.Check "TestColLetter(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone16
End Sub
