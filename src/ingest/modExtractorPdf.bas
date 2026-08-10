Attribute VB_Name = "modExtractorPdf"
Option Explicit

' ============================================================================
' modExtractorPdf - PDFのフォールバック抽出とローカル一時コピー
' ----------------------------------------------------------------------------
' 2026-08-03(R13 Phase 0): modExtractor が28,000字のWARN帯に達し、次の1行も
' 足せなくなったため、以下を純移設した(ロジックは変えていない)。
'   ・ExtractPdfWithFallback  PDF本文抽出の3経路フォールバック
'   ・CopyToLocalTemp / CopySharedRead  ローカル一時コピー(共有読み)
'   ・DropGarbledPages / GarbleRatio    文字化けページの除去
' 呼び出し元は modExtractor.ExtractFile のみ。ページ配列の組み立て
' (BuildPagesFromGsText)・件数(PageArrayCount)は modExtractor 側に残して
' ある(既存の契約と純ロジックテストの参照先を動かさないため)。
'
' R14-3a(2026-08-03): 「元サイズとコピー先サイズの突合」を必須にした。旧実装
' のクラシック Open(ANSI経路)は、CP932で表せない名前のときに存在しない別名の
' 0バイトファイルを【作って】掴み、それを「0バイトの正当なコピー」として True
' で返していた(実機第3報 RC3)。
' R14-F3(2026-08-03): その修正でコピーを ADODB.Stream 一本にしたため、今度は
' 「Excel/Acrobatで開いたままのファイル」が読めなくなった(LoadFromFile は
' 共有読みの意図を渡せない)。Dir$で見える普通の名前はクラシックの共有読み、
' 見えない名前(NFD分解等)と検証に落ちたときだけ Stream、の2段に戻した。
' 詳細は CopySharedRead のコメント。
'
' R13-2(2026-08-03): 一時コピー名を「mbtmp_<FNV-1a 64bit 16進16桁>.<元拡張子>」
' へ変更した。元のファイル名を連結する旧方式では、Mac由来のNFD分解濁点
' (U+3099)等のCP932非対応文字がクラシックOpen(ANSI経路)で別の字に化け、
' ディスク上の実名とGhostscriptへ渡す文字列が食い違って /undefinedfilename に
' なっていた(実機第2報 RC4)。元のファイル名は表示・ログ・メタデータ側だけで
' 使う。名前の導出は TempBaseNameFor(純ロジック)に切り出してテストで固定する。
' ============================================================================

' R10c(M1/M2): optGsTxtが「文字が薄すぎる」と判定した印。Word/Acrobatへは
' 通常のE0302と同様に譲り、そちらも失敗した場合の最終コードだけE0303にする。
Private Const GS_SPARSE_TOKEN As String = "#ERR:E0302:GS_SPARSE:"

' ----------------------------------------------------------------------------
' ExtractPdfWithFallback - PDFの本文抽出。3経路を順に試す。
'   (1) Ghostscript(txtwrite)  COM不要。同梱GSに文字を書き出させて読む
'   (2) Word(PDF Reflow)       COM
'   (3) Acrobat                 COM
'   outCode : "" のまま返れば呼び出し元が E0302 を採用する。"E0303" を入れて
'             返した場合は「画像PDF(文字が入っていない)」の印で、呼び出し元が
'             OCR経路(modShelfVision)へ回す。
'
' 2026-07-31(R10-3): 順序を「Word → Acrobat」から「GS → Word → Acrobat」へ
' 変えた。実機(管理端末)では、
'   ・Word/AcrobatのCreateObjectがポリシーで塞がれ、テキストPDFが1本も
'     取り込めない(実機初報B・E0302)。
'   ・塞がれていない端末でも、WordのPDF Reflowが「'Word'がOLE操作を完了する
'     のを待っています」ダイアログを頻発させ、数分たっても終わりが見えない
'     (追加報告)。
' という2つの症状が出ていた。GSは別プロセスでタイムアウト制御が効き、
' テキストPDFなら数秒で返るため、先に通せばOLE待ちダイアログごと消える。
' 画像PDFもGSが「文字がほとんど無い」と判定した時点でOCR経路へ回し、
' Wordには渡さない(必ず失敗するか極端に遅いだけなので待たせる意味が無い)。
' opt名は書かず modFeatures.InvokeFeature 経由で呼ぶ(R2)。vision機能が
' 無効なビルド構成では "#ERR:FEATURE_UNAVAILABLE" が返るだけなので、
' 従来どおり Word → Acrobat の連鎖へ静かに落ちる。
' ----------------------------------------------------------------------------
Public Function ExtractPdfWithFallback(ByVal path As String, ByVal maxPages As Long, _
                                       ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
                                       ByRef errDetail As String, ByRef outCode As String, _
                                       Optional ByVal origPath As String = "") As Boolean
    Dim gsText As String
    Dim gsErr As String
    Dim wordErr As String
    Dim acroErr As String
    Dim sparseSeen As Boolean

    outCode = ""
    sparseSeen = False

    gsText = VisionResultToText(modFeatures.InvokeFeature("vision", "ExtractPdfTextNoOcr", path))
    If Left$(gsText, 11) = "#ERR:E0202:" Then gsText = RetryGsAfterE0202(path)
    If LenB(gsText) > 0 And Left$(gsText, 5) <> "#ERR:" Then
        If modExtractor.BuildPagesFromGsText(gsText, maxPages, pages, truncated) Then
            ExtractPdfWithFallback = True
            Exit Function
        End If
        gsErr = "抽出テキストをページへ分解できませんでした"
    ElseIf Left$(gsText, 11) = "#ERR:E0303:" Then
        ' 文字層が1文字も無いPDF=スキャン確定。Wordに渡しても必ず空振りする
        ' (しかも遅い)ので、ここで即OCR経路へ回す。
        ' R13-3c: E0303は3箇所から同じコードで発報される。どこが出したのかを
        ' detailの先頭1語で言い切る(診断者の取り違えを消す。実機第2報 RC10)。
        outCode = "E0303"
        errDetail = modUtil.SafeLeft("gs_image: GS: " & gsText, 600)
        ExtractPdfWithFallback = False
        Exit Function
    Else
        gsErr = gsText
        If LenB(gsErr) = 0 Then gsErr = "(応答なし)"
        ' R10c(M1/M2): 「薄すぎる」は判断が割れる領域。即OCRへ回すと表紙だけの
        ' 短い正当なPDFを取りこぼすので、まずWord/Acrobatに任せる。
        If Left$(gsText, Len(GS_SPARSE_TOKEN)) = GS_SPARSE_TOKEN Then sparseSeen = True
    End If

    If modExtractorWord.Extract(path, maxPages, pages, truncated, wordErr, origPath) Then
        ExtractPdfWithFallback = True
        Exit Function
    End If

    If modExtractorAcrobat.Extract(path, maxPages, pages, truncated, acroErr) Then
        ExtractPdfWithFallback = True
        Exit Function
    End If

    ' R13-3c: 3経路が全滅し、なおかつGSが「薄すぎる」と言っていた場合のE0303。
    ' 上の gs_image(GSの即断)とは判断の根拠がまるで違うので、detailで区別する。
    Dim detailHead As String
    If sparseSeen Then
        outCode = "E0303"
        detailHead = "allfail: "
    End If

    errDetail = modUtil.SafeLeft(detailHead & "GS: " & gsErr & " / Word: " & wordErr & _
        " / Acrobat: " & acroErr, 600)
    ExtractPdfWithFallback = False
End Function

' ----------------------------------------------------------------------------
' RetryGsAfterE0202 - GS本文抽出が E0202 で落ちたときだけ、その場で1回やり直す
'   (2026-08-04 R15-2c・実機第4報 RC8)。
' ----------------------------------------------------------------------------
' E0202の実体はAIとの通信ではなく、Application.Run による自己呼び出し
' (optVision.ExtractPdfTextNoOcr)の失敗である。GSの完了待ちDoEvents中に
' 別の操作が入れ子で走ったときに起きる【一時的な】衝突で、資料の中身とは
' 何の関係も無い。にもかかわらず従来は即座に Word/Acrobat へ譲り、そちらも
' 塞がれている端末では取込ごと失敗して次回同期まで持ち越されていた。
' 1回だけやり直せば、入れ子の相手はもう終わっているのが普通なので通る。
' 2回目も失敗したら従来どおり(戻り値をそのまま返し、Word/Acrobatへ譲る)。
' 成否は usage_log に1行残す(効いているのかどうかを後から数えられるように)。
' R15-FixB(FB-9・レビューA-L3): 叩き直す前に一呼吸置く。E0202 の実体は
' 【入れ子の Application.Run が今まさに走っている】ことなので、間を空けずに
' 同じ呼び出しを繰り返せば同じ衝突をもう一度踏むだけ(1回きりの再試行の権利を
' 何もせずに捨てることになる)。DoEvents で相手に進む機会を渡し、0.5秒待つ。
' 0.5秒は「人が待たされたと感じない上限」と「入れ子の相手が終わる見込み」の
' 釣り合いで、失敗経路でしか通らないので通常の取込は1msも遅くならない。
Private Function RetryGsAfterE0202(ByVal path As String) As String
    DoEvents
    SleepMs 500

    Dim s As String
    s = VisionResultToText(modFeatures.InvokeFeature("vision", "ExtractPdfTextNoOcr", path))
    RetryGsAfterE0202 = s

    On Error Resume Next
    Dim outcome As String
    If Left$(s, 11) = "#ERR:E0202:" Then
        outcome = "fail"
    Else
        outcome = "ok"
    End If
    modLog.LogUsage "gs_e0202_retry", "", outcome & " " & _
        modUtil.SafeLeft(modUtil.FileNameOf(path), 100)
    On Error GoTo 0
End Function

' InvokeFeatureの戻り値(Variant)を安全に文字列化する
' (modShelfVision.ResultToText と同型。エラー値・非文字列は "" 扱い)。
Private Function VisionResultToText(ByVal result As Variant) As String
    On Error GoTo NotText
    If IsError(result) Then Exit Function
    If VarType(result) <> vbString Then Exit Function
    VisionResultToText = CStr(result)
    Exit Function
NotText:
    VisionResultToText = ""
End Function

' Declareを使わないスリープ(§13: 32/64bit互換のためDeclare不使用で回避)。
' DoEventsで応答性を保ちながらTimer基準で待つ。modEmbed に同型の Private が
' あるが、あちらは埋め込みのスロットリング専用で公開されておらず、
' ingest層の別モジュールから呼ぶために公開すると「待つ」という副作用が
' モジュール契約(§7)へ増える。5行の同型実装を許す方が影響が小さい
' (R15-FixB FB-9 の判断。共通化するなら両方を基盤層へ移すのが筋)。
Private Sub SleepMs(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do   ' 深夜0時のTimerロールオーバーガード
    Loop
End Sub

' ----------------------------------------------------------------------------
' TempBaseNameFor - 一時コピーのファイル名を決める(純ロジック・R13-2)。
'   "mbtmp_" & FNV-1a 64bit(元フルパスの16進16桁) & "." & 元拡張子(小文字)
'   seq>0 のときはハッシュの後ろに "_<seq>" を挟む(衝突時の連番)。
'   ハッシュ入力は【元のフルパス】。同名別フォルダのファイルを同時に取り込んでも
'   衝突しない。拡張子は抽出の振り分けに使うので必ず元のまま残す。
'   全ての名前が "mbtmp_" で始まるため、掃除のパターンは mbtmp_* のままでよい。
' ----------------------------------------------------------------------------
Public Function TempBaseNameFor(ByVal srcPath As String, Optional ByVal seq As Long = 0) As String
    Dim ext As String: ext = modUtil.ExtOf(srcPath)
    Dim nm As String: nm = "mbtmp_" & modUtil.Fnv1a64Hex(srcPath)
    If seq > 0 Then nm = nm & "_" & seq
    If LenB(ext) > 0 Then nm = nm & "." & ext
    TempBaseNameFor = nm
End Function

' path を %TEMP% 配下へコピーし、そのコピー先パスを返す(成功時)。
' コピーできなければ空文字列を返し、outReason に理由の1語を入れる。
' 拡張子は必ず元と同じにする(抽出は拡張子で振り分けるため)。
' Word/Excelの保護ビュー(ネットワーク/クラウド上のファイルで発動)を避け、
' かつ抽出中の元ファイルロック・ネットワーク瞬断の影響を受けないようにする。
'
' outReason(2026-08-03 R14-3a / R14-F3): "src_empty" / "size_mismatch …" /
'   "locked err#N" / "too_big err#N" / "load_fail err#N" / "no_temp" /
'   "name_busy"。呼び出し元はこれを見て「元パスのまま続行してよいか」を
'   決め(IsUnreadableCopyReason)、利用者への1文を選ぶ(CopyFailMsgFor)。
'   従来 localcopy_fail の usage_log は【例外が飛んだときだけ】書かれており、
'   「Falseで静かに帰ってきた」失敗は1行も残っていなかった(観測性ゼロ)。
'   ここでBoolean失敗も必ず1行残す(憲章§4-1)。
Public Function CopyToLocalTemp(ByVal path As String, _
                                Optional ByRef outReason As String) As String
    outReason = ""
    On Error GoTo Fail

    ' R13-4a: 数百MBのPDFではコピーだけで十数秒かかる。ここが無言だと
    ' 利用者には「押した直後から何も起きない」ようにしか見えない(憲章§3-2)。
    ' バナーが既に出ているときだけ更新される(silentは silent のまま)。
    modShelfBatch.StageBanner "コピー中…"

    Dim tempDir As String: tempDir = Environ$("TEMP")
    If LenB(tempDir) = 0 Then tempDir = Environ$("TMP")
    If LenB(tempDir) = 0 Then
        outReason = "no_temp"
        Exit Function
    End If
    If Right$(tempDir, 1) <> "\" Then tempDir = tempDir & "\"

    ' R13-2: 名前は元ファイル名ではなくフルパスのハッシュから作る(TempBaseNameFor)。
    ' 衝突回避のため連番を付ける(同時取込や前回の消し残りに備える)。
    Dim dest As String
    Dim n As Long: n = 0
    Do
        dest = tempDir & TempBaseNameFor(path, n)
        If LenB(Dir$(dest)) = 0 Then Exit Do
        n = n + 1
        If n > 500 Then
            outReason = "name_busy"
            Exit Function
        End If
    Loop

    Dim reason As String
    If Not CopySharedRead(path, dest, reason) Then
        outReason = reason
        On Error Resume Next
        modLog.LogUsage "localcopy_fail", "", modUtil.SafeLeft(path, 300) & _
            " : " & reason & " len=" & Len(path)
        On Error GoTo Fail
        Exit Function
    End If

    CopyToLocalTemp = dest
    Exit Function

Fail:
    ' R11-D(監査3 M-6): 従来ここは完全な無言だった。MAX_PATH超過・権限・
    ' 容量を切り分けるためErr情報とパス長を残す(LogUsageは自前で失敗を握る)。
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かない。Resumeで抜けてから記録する。
    Resume CopyFailCleanup
CopyFailCleanup:
    On Error Resume Next
    outReason = "load_fail err#" & failNum
    modLog.LogUsage "localcopy_fail", "", "err#" & failNum & " " & _
        modUtil.SafeLeft(failDesc, 120) & " len=" & Len(path)
    On Error GoTo 0
    CopyToLocalTemp = ""
End Function

' ----------------------------------------------------------------------------
' CopySharedRead - srcPath を destPath へバイナリコピーする(R14-3a / R14-F3)。
'   2段構えにする。どちらの段も「元サイズとコピー先サイズの突合」で締める。
'
'   (1) Dir$ で見えるファイル = 名前がANSI(CP932)で表せる普通のファイル。
'       クラシックの「Open For Binary Access Read Shared」+1MBずつの
'       Get/Put でコピーする。共有読みなので、Excel/Acrobat/Edge が開いた
'       ままのファイルでも【読むだけ】なら成立する(排他ロックは書き込み側の
'       意図が無い限りかからない)。R14-3a でここをADODB.Stream一本にした
'       ため、開いたままのxlsx/pdfが「読めない」に化けていた
'       (LoadFromFile は共有読みの意図を渡せない=ロックに弱い)。
'       Dir$ 自体は副作用が無く、ANSIで表せない名前なら「無い」と答えるだけ
'       なので、判別の道具として安全に使える。
'   (2) Dir$ で見えない(=NFD分解濁点 U+3099 等のCP932非対応名)、または
'       (1)が「0バイト/サイズ不一致」で終わったときは ADODB.Stream。
'       LoadFromFile はパスをBSTR(COM)で渡すのでUnicodeのまま届き、実体が
'       無ければその場で失敗する(クラシックOpenのように【作らない】)。
'       実機第3報 RC3 の0バイト嘘コピーはこれで塞いだままにする。
'
'   outReason(F3で細分化。呼び出し元はこれで利用者への1文を選ぶ):
'     "src_empty"        元が0バイト(空のPDFは取り込む中身が無い)
'     "size_mismatch …"  コピー後のサイズが元と違う
'     "locked err#N"     開けなかった(クラシック70 / Dir$可視でのStream 3002)
'     "too_big err#N"    大きすぎて載らない(Stream 7/3004/14・クラシック7)
'     "load_fail err#N"  それ以外(NFD名で実体に届かなかった場合もここ)
'   失敗時は書きかけのdestを必ず消す(%TEMP%に壊れた実体を残さない。
'   残すと次回の連番探索がその残骸を避けて番号を消費し続ける)。
' ----------------------------------------------------------------------------
Private Function CopySharedRead(ByVal srcPath As String, ByVal destPath As String, _
                                ByRef outReason As String) As Boolean
    outReason = ""

    Dim visible As Boolean: visible = PathIsAnsiVisible(srcPath)
    If visible Then
        Dim classicReason As String: classicReason = ""
        If CopyClassicShared(srcPath, destPath, classicReason) Then
            CopySharedRead = True
            Exit Function
        End If
        ' 中身が食い違う/空だったときだけ、Unicode経路でもう一度確かめる。
        ' ロック・容量の失敗はStreamでも同じ結末なので、そのまま返す。
        If Not RetryWithStream(classicReason) Then
            outReason = classicReason
            Exit Function
        End If
    End If

    CopySharedRead = CopyViaStream(srcPath, destPath, visible, outReason)
End Function

' クラシックの共有読みコピー(1MBずつ)。R13以前の実装をそのまま戻したもので、
' 「他のアプリで開いたままのファイルを取り込める」唯一の経路。
Private Function CopyClassicShared(ByVal srcPath As String, ByVal destPath As String, _
                                   ByRef outReason As String) As Boolean
    Const CHUNK_BYTES As Long = 1048576   ' 1MB単位。巨大PDFでも一括Getしない
    Dim srcNum As Long: srcNum = 0
    Dim dstNum As Long: dstNum = 0

    outReason = ""
    On Error GoTo ClassicFailed

    srcNum = FreeFile
    Open srcPath For Binary Access Read Shared As #srcNum

    ' R11-D(監査3 M-7): 2GB超はLongに収まらず代入自体がerr#6になる。
    ' Doubleで受けて範囲外は明示的に失敗させ、壊れた中途半端なコピーを作らない。
    Dim totalLenD As Double: totalLenD = LOF(srcNum)
    If totalLenD > 2147483647# Then Err.Raise 7
    Dim totalLen As Long: totalLen = CLng(totalLenD)

    If totalLen <= 0 Then
        Close #srcNum
        srcNum = 0
        outReason = "src_empty"
        Exit Function
    End If

    dstNum = FreeFile
    Open destPath For Binary Access Write As #dstNum

    Dim pos As Long: pos = 1
    Do While pos <= totalLen
        Dim thisLen As Long: thisLen = CHUNK_BYTES
        If pos + thisLen - 1 > totalLen Then thisLen = totalLen - pos + 1
        If thisLen < 1 Then Exit Do
        Dim buf() As Byte: ReDim buf(1 To thisLen)
        Get #srcNum, pos, buf
        Put #dstNum, pos, buf
        pos = pos + thisLen
    Loop

    Close #dstNum
    Close #srcNum
    dstNum = 0
    srcNum = 0

    ' R14-3a: サイズ突合まで通って初めて「コピーできた」と言う。
    Dim destSize As Double: destSize = CopiedFileSize(destPath)
    If destSize <> CDbl(totalLen) Then
        outReason = "size_mismatch src=" & CStr(totalLen) & " dest=" & CStr(destSize)
        KillCopyDest destPath
        Exit Function
    End If

    CopyClassicShared = True
    Exit Function

ClassicFailed:
    ' Err は Resume でクリアされるので先に控える。ハンドラ稼働中は
    ' On Error Resume Next が効かないため、後始末はハンドラを抜けてから行う。
    Dim failNum As Long: failNum = Err.Number
    Resume ClassicCleanup
ClassicCleanup:
    On Error Resume Next
    If srcNum <> 0 Then Close #srcNum
    If dstNum <> 0 Then Close #dstNum
    On Error GoTo 0
    outReason = ClassicFailReason(failNum)
    KillCopyDest destPath
    CopyClassicShared = False
End Function

' ADODB.Stream によるコピー(Unicodeパス経路)。visible はDir$で見えていたか
' (見えているのに開けない=ロック、見えていない=実体に届いていない)。
Private Function CopyViaStream(ByVal srcPath As String, ByVal destPath As String, _
                               ByVal visible As Boolean, ByRef outReason As String) As Boolean
    Dim st As Object
    Dim srcSize As Double
    Dim destSize As Double

    outReason = ""

    On Error GoTo StreamFailed
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                 ' adTypeBinary
    st.Open
    st.LoadFromFile srcPath     ' 実体が無ければここで失敗する(作らない)
    srcSize = CDbl(st.Size)
    If srcSize > 0 Then st.SaveToFile destPath, 2   ' adSaveCreateOverWrite
    st.Close
    Set st = Nothing            ' COM解放(正常パス)

    If srcSize <= 0 Then
        outReason = "src_empty"
        Exit Function
    End If

    destSize = CopiedFileSize(destPath)
    If destSize <> srcSize Then
        outReason = "size_mismatch src=" & CStr(srcSize) & " dest=" & CStr(destSize)
        KillCopyDest destPath
        Exit Function
    End If

    CopyViaStream = True
    Exit Function

StreamFailed:
    Dim failNum As Long: failNum = Err.Number
    Resume StreamCleanup
StreamCleanup:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing            ' COM解放(異常パス=半開きも確実に解放)
    On Error GoTo 0
    outReason = StreamFailReason(failNum, visible)
    KillCopyDest destPath
    CopyViaStream = False
End Function

' クラシックOpenの実行時エラー番号 → 理由の1語(F3)。
'   70 = 書き込み禁止(排他ロック中の共有読み拒否)
'   7  = メモリ不足 / 2GB超の明示的な失敗
Private Function ClassicFailReason(ByVal errNum As Long) As String
    Select Case errNum
        Case 70
            ClassicFailReason = "locked err#70"
        Case 7
            ClassicFailReason = "too_big err#7"
        Case Else
            ClassicFailReason = "load_fail err#" & errNum
    End Select
End Function

' ADODB.Streamの実行時エラー番号 → 理由の1語(F3)。
'   3002 = ファイルを開けない。Dir$で見えていたなら他アプリのロック、
'          見えていなかったなら「その名前の実体に届いていない」(NFD名)。
'   7/3004/14 = メモリ不足・読み取り位置超過・文字列が長すぎる(巨大ファイル)
Private Function StreamFailReason(ByVal errNum As Long, ByVal visible As Boolean) As String
    Select Case errNum
        Case 3002
            If visible Then
                StreamFailReason = "locked err#3002"
            Else
                StreamFailReason = "load_fail err#3002"
            End If
        Case 7, 3004, 14
            StreamFailReason = "too_big err#" & errNum
        Case Else
            StreamFailReason = "load_fail err#" & errNum
    End Select
End Function

' クラシック経路の失敗のうち、Stream経路で確かめ直す価値があるものだけTrue。
Private Function RetryWithStream(ByVal reason As String) As Boolean
    RetryWithStream = (Left$(reason, Len("src_empty")) = "src_empty" Or _
                       Left$(reason, Len("size_mismatch")) = "size_mismatch")
End Function

' Dir$ で見えるか(=名前がANSI/CP932で表せるか)。Dir$ は副作用が無く、
' 表せない名前には「無い」と答えるだけなので判別の道具に使える。
Private Function PathIsAnsiVisible(ByVal p As String) As Boolean
    On Error Resume Next
    PathIsAnsiVisible = (LenB(Dir$(p)) > 0)
    On Error GoTo 0
End Function

' コピー先の実サイズ(取れなければ -1 = 不一致として扱われる)。
' FileLen は実体が無いと実行時エラー53を出すので、検証専用に握る。
Private Function CopiedFileSize(ByVal p As String) As Double
    On Error Resume Next
    CopiedFileSize = -1
    CopiedFileSize = CDbl(FileLen(p))
    On Error GoTo 0
End Function

' 書きかけ・検証に落ちたコピー先を消す(失敗しても無視)。
Private Sub KillCopyDest(ByVal destPath As String)
    On Error Resume Next
    If LenB(destPath) > 0 Then
        If LenB(Dir$(destPath)) > 0 Then Kill destPath
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' IsUnreadableCopyReason / CopyFailMsgFor - コピー失敗の理由が「元ファイルの
'   中身を手にできていない」ことを意味するか、と、そのときの利用者向けの1文
'   (2026-08-03 R14-3b / R14-F3・F11。いずれも純ロジック)。
'   中身が無いのに元パスのまま処理を続けると、NFD分解名がそのまま
'   Ghostscriptへ渡って /undefinedfilename になり、利用者には「描画0枚」
'   としか残らない(RC3の本体)。そうなる前に正直に止めるための判定。
'   name_busy(一時ファイル名を500個試しても空きが無い)もここに含める
'   (R14-F11)。コピーが1つも作れていない以上、続行すれば同じ穴に落ちる。
'   文言の一次情報は modLog.CopyFailMsgOf(取込経路とOCR経路の両方から同じ
'   文を出す必要があり、opt層はコア基盤層しか参照できないため)。
'   CopyReasonKind - 理由の1語 → 文言の種類。ここだけが両者の対応表。
' ----------------------------------------------------------------------------
Public Function IsUnreadableCopyReason(ByVal reason As String) As Boolean
    Dim r As String: r = LCase$(Trim$(reason))
    If LenB(r) = 0 Then Exit Function

    IsUnreadableCopyReason = (StartsWith(r, "size_mismatch") Or _
                              StartsWith(r, "src_empty") Or _
                              StartsWith(r, "load_fail") Or _
                              StartsWith(r, "locked") Or _
                              StartsWith(r, "too_big") Or _
                              StartsWith(r, "name_busy"))
End Function

Public Function CopyFailMsgFor(ByVal reason As String) As String
    If Not IsUnreadableCopyReason(reason) Then Exit Function
    CopyFailMsgFor = modLog.CopyFailMsgOf(CopyReasonKind(reason)) & _
        " [" & Trim$(reason) & "]"
End Function

' 理由の1語("locked err#70" 等)から文言の種類を決める(純ロジック)。
' 既定("")は modLog 側で「ファイル名の特殊文字が原因の可能性」になる=
' 実体に届かなかった系(load_fail / size_mismatch)の文面。
Public Function CopyReasonKind(ByVal reason As String) As String
    Dim r As String: r = LCase$(Trim$(reason))
    If StartsWith(r, "src_empty") Then
        CopyReasonKind = "src_empty"
    ElseIf StartsWith(r, "locked") Then
        CopyReasonKind = "locked"
    ElseIf StartsWith(r, "too_big") Then
        CopyReasonKind = "too_big"
    ElseIf StartsWith(r, "name_busy") Then
        CopyReasonKind = "name_busy"
    End If
End Function

Private Function StartsWith(ByVal s As String, ByVal head As String) As Boolean
    StartsWith = (Left$(s, Len(head)) = head)
End Function

' ----------------------------------------------------------------------------
' GcOldTempCopies - %TEMP%\mbtmp_* のうち24時間より古いものを消す
'   (2026-08-03 R14-F11)。戻り値=消せた件数。
'   一時コピーは全ての出口で消す設計だが、Excelの強制終了・電源断・ロックの
'   タイミング次第で残る。残ると次回の連番探索(mbtmp_<hash>/_1/_2…)が
'   その残骸を避けて番号を消費し続け、最後は name_busy で取込が止まる
'   (500個で打ち切り)。掃除役はGCが既に1箇所ある起動時に相乗りさせる。
'   24時間という線は nxocr_* のGCと同じ理由(いま動いている取込を絶対に
'   壊さない)。呼び出しは modBoot から1行だけ。本体をこちらに置くのは
'   modBoot が28,000字のWARN帯に近いためで、命名規約(TempBaseNameFor)を
'   持つのがこのモジュールである以上、置き場としてもここが正しい。
'   列挙中に削除するとDirの列挙が壊れるので、先に名前だけ集める。
' ----------------------------------------------------------------------------
Public Function GcOldTempCopies() As Long
    Dim tempDir As String: tempDir = Environ$("TEMP")
    If LenB(tempDir) = 0 Then tempDir = Environ$("TMP")
    If LenB(tempDir) = 0 Then Exit Function
    If Right$(tempDir, 1) <> "\" Then tempDir = tempDir & "\"

    Dim names() As String: ReDim names(0 To 63)
    Dim n As Long: n = 0

    On Error Resume Next
    Dim nm As String: nm = Dir$(tempDir & "mbtmp_*")
    Do While LenB(nm) > 0
        If n > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 64)
        names(n) = nm
        n = n + 1
        If n >= 500 Then Exit Do        ' 暴走防止(1回の起動で見る上限)
        nm = Dir$()
    Loop
    Err.Clear

    Dim removed As Long: removed = 0
    Dim i As Long
    For i = 0 To n - 1
        Dim full As String: full = tempDir & names(i)
        Dim stamp As Date
        stamp = FileDateTime(full)
        If Err.Number = 0 Then
            If DateDiff("h", stamp, Now) >= 24 Then
                Kill full
                If Err.Number = 0 Then removed = removed + 1
            End If
        End If
        Err.Clear
    Next i

    If removed > 0 Then
        modLog.LogUsage "tempcopy_gc", "", CStr(removed) & _
            "件の古い一時コピー(" & tempDir & "mbtmp_*)を削除しました"
    End If
    On Error GoTo 0

    GcOldTempCopies = removed
End Function

' ----------------------------------------------------------------------------
' DropGarbledPages - 文字化けしたページを配列から取り除いて詰める
'   (戻り値=取り除いた数)。要件A(2026-07-30): 旧実装は.Textを""にする
'   だけで配列を詰めず、空文字ページが後段でerr#9を誘発していた。
' ----------------------------------------------------------------------------
' 判定は「日本語の業務文書には出ない文字」の比率。具体的には
'   ・制御文字(Chr 0～31。タブ/改行を除く)
'   ・キリル文字/ギリシャ文字のブロック
' これらが本文の2割を超えるページは、フォント由来の化けと見なして捨てる。
' 保険の約款・ガイドラインにギリシャ文字やキリル文字が2割入ることはない。
' 短いページ(50字未満)は判定しない(誤爆すると目次や章扉を落としてしまう)。
' 残すページの.pageは元のページ番号のまま(出典表示のため詰め直さない)。
' 全ページが化け判定になる場合は除去しない(戻り値0=何も取り除いていない)。
' フィルタの誤判定で文書を丸ごと失うより、化けた本文でも取り込む方がまし。
' 2026-07-31(R6追補): ただしPDFだけは例外で、呼び出し側が outAllGarbled=True を
' 見てOCR経路(E0303)へ回す。どちらの扱いになったかのusage_logは呼び出し側が
' 1本だけ残す(ここで出すと二重になり、ログを読む人が混乱するため)。
Public Function DropGarbledPages(ByRef pages() As ExtractedPage, _
                                 ByRef outAllGarbled As Boolean) As Long
    outAllGarbled = False
    Dim n As Long: n = modExtractor.PageArrayCount(pages)
    If n = 0 Then Exit Function

    Dim lo As Long: lo = LBound(pages)
    Dim isGarbled() As Boolean: ReDim isGarbled(0 To n - 1)
    Dim garbledCount As Long: garbledCount = 0

    Dim i As Long
    For i = 0 To n - 1
        Dim t As String: t = pages(lo + i).Text
        If Len(t) >= 50 Then
            If modExtractor.GarbleRatio(t) > 0.2 Then
                isGarbled(i) = True
                garbledCount = garbledCount + 1
            End If
        End If
    Next i

    If garbledCount = 0 Then Exit Function

    If garbledCount = n Then
        outAllGarbled = True
        Exit Function
    End If

    Dim kept() As ExtractedPage: ReDim kept(0 To n - garbledCount - 1)
    Dim k As Long: k = 0
    For i = 0 To n - 1
        If Not isGarbled(i) Then
            kept(k) = pages(lo + i)
            k = k + 1
        End If
    Next i
    pages = kept

    DropGarbledPages = garbledCount
End Function

' ----------------------------------------------------------------------------
' IsThinExtract - 「取り込めてはいるが、本文が薄すぎる」の判定(純ロジック)。
'   2026-08-03 R13-3a(二段目の防衛)。実機第2報 RC1 では、44ページの約款が
'   chunks=1 / status=done で「登録成功」になっていた。上流(R13-1b)で
'   文字層なしPDFはOCRへ回るようになったので、ここへ落ちるのは希少ケースだが、
'   「入っているのに中身が無い」を done と言い切るのは憲章§3-3/§4-1に反する。
'
'   条件(仕様R13-3aの値から動かさないこと):
'     ext = "pdf" かつ pageCount >= 10 かつ
'     (chunkN <= pageCount\20 または 総文字数 < pageCount*60)
'   10ページ未満を見ないのは、表紙+ポンチ絵のような正当に薄い資料を
'   partial と呼ばないため。整数除算(\)なのは「20ページで1チャンク以下」
'   という意味をそのまま式にしたもの。
'   R13-F7: PDF以外には当てない。docxの「ページ」もxlsxの「シート」も同じ
'   ExtractedPage配列に入るため、この関門はスライド資料や表計算にも当たって
'   いた。文言(画像中心のPDF・自動でOCR)がそもそも噛み合わないうえ、
'   1シートに要約表が1枚あるだけの正当なブックまで partial にしてしまう。
'   守りたい事故(RC1)はPDF経路でしか起きないので、判定ごとPDFに限定する。
' ----------------------------------------------------------------------------
Public Function IsThinExtract(ByVal ext As String, ByVal pageCount As Long, _
                              ByVal chunkN As Long, _
                              ByVal totalChars As Long) As Boolean
    If LCase$(Trim$(ext)) <> "pdf" Then Exit Function
    If pageCount < 10 Then Exit Function
    If chunkN <= (pageCount \ 20) Then
        IsThinExtract = True
        Exit Function
    End If
    IsThinExtract = (totalChars < pageCount * 60)
End Function

' ----------------------------------------------------------------------------
' ThinExtractMemoFor - 薄い抽出なら本棚カードに出すメモを、そうでなければ ""。
'   modShelf.IngestFile は「ページ数と総文字数を数える」処理を持たないので、
'   数える所からメモの文言までをここへ寄せる(modShelf は呼び出し1行で済む)。
'   文言は「何が起きたか+次に何が起きるか」の2点だけを言い、利用者に
'   操作を求めない(自動でOCRを試すのはこちらの仕事だから)。
'   ext(R13-F7): 呼び出し元が知っている拡張子。PDF以外は常に ""(判定しない)。
' ----------------------------------------------------------------------------
Public Function ThinExtractMemoFor(ByRef pages() As ExtractedPage, _
                                   ByVal chunkN As Long, _
                                   ByVal ext As String) As String
    If LCase$(Trim$(ext)) <> "pdf" Then Exit Function
    Dim n As Long: n = modExtractor.PageArrayCount(pages)
    If n = 0 Then Exit Function

    Dim lo As Long: lo = LBound(pages)
    Dim total As Long
    Dim i As Long
    For i = 0 To n - 1
        total = total + Len(pages(lo + i).Text)
    Next i

    If Not IsThinExtract(ext, n, chunkN, total) Then Exit Function
    ThinExtractMemoFor = "本文を十分に取り出せていない可能性があります" & _
        "(画像中心のPDFの場合は自動でOCRを試します)"
End Function

' 化け文字の比率(0.0〜1.0)は 2026-08-10(R27 F1-4)で modExtractor へ移設した。
' 本モジュールは30,000字上限まで残りが少なく、誤爆2件(AscWの符号・Word構造
' 制御文字)の是正とその説明が入らなかったため。呼び出しは上の
' DropGarbledPages 1箇所だけ。
