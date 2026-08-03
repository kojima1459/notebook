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
' (BuildPagesFromGsText)・件数(PageArrayCount)・分割サイズ計算
' (SharedCopyNextChunkLen)は modExtractor 側に残してある(既存の契約と
' 純ロジックテストの参照先を動かさないため)。
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
' コピーできなければ空文字列を返す(呼び出し側は元パスのまま処理を続ける)。
' 拡張子は必ず元と同じにする(抽出は拡張子で振り分けるため)。
' Word/Excelの保護ビュー(ネットワーク/クラウド上のファイルで発動)を避け、
' かつ抽出中の元ファイルロック・ネットワーク瞬断の影響を受けないようにする。
'
' 2026-07-30(要件C): 従来はFileCopy(排他オープン)を使っており、Word/Excel/
' Acrobatで開いているファイルはコピーに失敗していた(実機ログ: 悪天候の
' 定義.docx : localcopy=failed。Desktopのファイル。「Wordを開いたまま使う」
' 運用のため頻発)。共有読み(Open ... For Binary Access Read Shared)は、
' 他プロセスが開いていても読み出しだけなら成立する(排他ロックは書き込み側
' の意図が無い限りかからない)ため、これへ置き換える。
Public Function CopyToLocalTemp(ByVal path As String) As String
    On Error GoTo Fail

    ' R13-4a: 数百MBのPDFでは共有読みコピーだけで十数秒かかる。ここが無言だと
    ' 利用者には「押した直後から何も起きない」ようにしか見えない(憲章§3-2)。
    ' バナーが既に出ているときだけ更新される(silentは silent のまま)。
    modShelfBatch.StageBanner "コピー中…"

    Dim tempDir As String: tempDir = Environ$("TEMP")
    If LenB(tempDir) = 0 Then tempDir = Environ$("TMP")
    If LenB(tempDir) = 0 Then Exit Function
    If Right$(tempDir, 1) <> "\" Then tempDir = tempDir & "\"

    ' R13-2: 名前は元ファイル名ではなくフルパスのハッシュから作る(TempBaseNameFor)。
    ' 衝突回避のため連番を付ける(同時取込や前回の消し残りに備える)。
    Dim dest As String
    Dim n As Long: n = 0
    Do
        dest = tempDir & TempBaseNameFor(path, n)
        If LenB(Dir$(dest)) = 0 Then Exit Do
        n = n + 1
        If n > 500 Then Exit Function   ' 異常時の暴走防止
    Loop

    If Not CopySharedRead(path, dest) Then Exit Function
    CopyToLocalTemp = dest
    Exit Function

Fail:
    ' R11-D(監査3 M-6): 従来ここは完全な無言だった。MAX_PATH超過・権限・
    ' 容量を切り分けるためErr情報とパス長を残す(LogUsageは自前で失敗を握る)。
    modLog.LogUsage "localcopy_fail", "", "err#" & Err.Number & " " & _
        modUtil.SafeLeft(Err.Description, 120) & " len=" & Len(path)
    CopyToLocalTemp = ""
End Function

' 共有読み(Shared)でsrcPathをdestPathへバイナリコピーする。
' 0バイトファイルはGet/Putを一度も行わず、空ファイルのままコピー成立とする
' (要件C: 0バイトファイルの扱いに注意)。
Private Function CopySharedRead(ByVal srcPath As String, ByVal destPath As String) As Boolean
    Dim srcNum As Long: srcNum = 0
    Dim dstNum As Long: dstNum = 0
    On Error GoTo Failed

    srcNum = FreeFile
    Open srcPath For Binary Access Read Shared As #srcNum

    ' R11-D(監査3 M-7): 2GB超はLongに収まらず代入自体がerr#6になる。
    ' Doubleで受けて範囲外は明示的に失敗させ、壊れた中途半端なコピーを作らない。
    Dim totalLenD As Double: totalLenD = LOF(srcNum)
    If totalLenD > 2147483647# Then Err.Raise 6
    Dim totalLen As Long: totalLen = CLng(totalLenD)

    dstNum = FreeFile
    Open destPath For Binary Access Write As #dstNum

    Const CHUNK_BYTES As Long = 1048576   ' 1MB単位。巨大PDFでも一括Getせず段階的に読む
    Dim pos As Long: pos = 1
    Do While pos <= totalLen
        Dim thisLen As Long: thisLen = modExtractor.SharedCopyNextChunkLen(pos, totalLen, CHUNK_BYTES)
        If thisLen < 1 Then Exit Do
        Dim buf() As Byte: ReDim buf(1 To thisLen)
        Get #srcNum, pos, buf
        Put #dstNum, pos, buf
        pos = pos + thisLen
    Loop

    Close #dstNum
    Close #srcNum
    CopySharedRead = True
    Exit Function

Failed:
    ' ハンドラ稼働中はOn Error Resume Nextが効かないため、後始末は
    ' 別Subへ切り出す(R6)。開きかけたファイル番号を確実に閉じる。
    CloseCopyFileNumbers srcNum, dstNum, destPath
    CopySharedRead = False
End Function

' CopySharedRead失敗時の後始末専用(R6: 稼働中ハンドラの中では
' On Error Resume Nextが効かないため、新しいエラー文脈を持つ別Subへ切り出す)。
' 2026-07-30(レビュー4-E): 書きかけのコピー先も消す。途中まで書けた
' ファイルを残すと、%TEMP%に壊れたファイルが溜まるだけでなく、
' 次回の連番探索(mbtmp_<hash>/mbtmp_<hash>_1/...)がその残骸を避けて
' 番号を消費し続ける。
Private Sub CloseCopyFileNumbers(ByVal n1 As Long, ByVal n2 As Long, _
                                 ByVal destPath As String)
    On Error Resume Next
    If n1 <> 0 Then Close #n1
    If n2 <> 0 Then Close #n2
    If LenB(destPath) > 0 Then
        If LenB(Dir$(destPath)) > 0 Then Kill destPath
    End If
    On Error GoTo 0
End Sub

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
            If GarbleRatio(t) > 0.2 Then
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
'     pageCount >= 10 かつ (chunkN <= pageCount\20 または 総文字数 < pageCount*60)
'   10ページ未満を見ないのは、表紙+ポンチ絵のような正当に薄い資料を
'   partial と呼ばないため。整数除算(\)なのは「20ページで1チャンク以下」
'   という意味をそのまま式にしたもの。
' ----------------------------------------------------------------------------
Public Function IsThinExtract(ByVal pageCount As Long, ByVal chunkN As Long, _
                              ByVal totalChars As Long) As Boolean
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
' ----------------------------------------------------------------------------
Public Function ThinExtractMemoFor(ByRef pages() As ExtractedPage, _
                                   ByVal chunkN As Long) As String
    Dim n As Long: n = modExtractor.PageArrayCount(pages)
    If n = 0 Then Exit Function

    Dim lo As Long: lo = LBound(pages)
    Dim total As Long
    Dim i As Long
    For i = 0 To n - 1
        total = total + Len(pages(lo + i).Text)
    Next i

    If Not IsThinExtract(n, chunkN, total) Then Exit Function
    ThinExtractMemoFor = "本文を十分に取り出せていない可能性があります" & _
        "(画像中心のPDFの場合は自動でOCRを試します)"
End Function

' 化け文字の比率(0.0～1.0)。空白は数えない。
Private Function GarbleRatio(ByVal s As String) As Double
    Dim bad As Long, tot As Long
    Dim i As Long
    For i = 1 To Len(s)
        Dim c As Long: c = AscW(Mid$(s, i, 1))
        If c = 32 Or c = 9 Or c = 10 Or c = 13 Or c = &H3000 Then
            ' 空白類は分母に入れない
        Else
            tot = tot + 1
            If c < 32 Then
                bad = bad + 1                        ' 制御文字
            ElseIf c >= &H370 And c <= &H3FF Then
                bad = bad + 1                        ' ギリシャ文字
            ElseIf c >= &H400 And c <= &H52F Then
                bad = bad + 1                        ' キリル文字
            End If
        End If
    Next i
    If tot > 0 Then GarbleRatio = bad / tot
End Function
