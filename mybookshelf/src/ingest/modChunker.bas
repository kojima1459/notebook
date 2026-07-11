Attribute VB_Name = "modChunker"
Option Explicit

' ============================================================================
' modChunker - 抽出済みページ群をオーバーラップ付きチャンクへ分割する
' ----------------------------------------------------------------------------
' 役割:
'   段落・文境界をなるべく尊重しながら、targetChars字前後・overlapChars字の
'   重なりを持つスライディングウィンドウでテキストを分割する。
'
' 流用元: /home/user/notebook/src/admin/modChunker.bas(V2資産・コピー元。
'   変更禁止)。ChunkText(1ページ分の平文を分割)のアルゴリズムを
'   ChunkOnePageとして踏襲し、複数ページ(ExtractedPage配列)をまとめて
'   処理する ChunkPages でラップする。
'
' 適応要件(MASTER_SPEC §7.2):
'   ・R4準拠(純ロジック): Worksheets/Range/Application/ThisWorkbook/MsgBox/
'     ActiveSheetは一切使わない。LibreOffice headlessでもそのまま実行できる
'     (tools/run_lo_tests.py の RunAllPureTests 対象)。
'   ・target/overlapは引数で受け取るのみで、modConfig は読みに行かない
'     (config読取はmodShelfの責務。§7.2本文の指示どおり)。
'   ・chunk_id/source/originはShelfChunkの既定値(空文字列)のまま返す。
'     付番は modShelf.IngestFile が行う契約(bs::hash::pN::cN)。
'     summary/keywordsも同様に空のまま(modEnrichが後で埋める)。
'   ・full_textは32000字を絶対に超えない。呼び出し側が異常に大きな
'     targetChars を渡した場合でも安全なように、ウィンドウ確定後に
'     もう一段 MAX_CHUNK_CHARS でクランプする防御を入れている。
'   ・段落・文境界優先はV1のFindSentenceBoundary(句読点・改行を境界候補と
'     みなし、target/4字ぶんの余裕内で探す)をそのまま踏襲する。
'   ・オーバーラップが target 以上だと次チャンクの開始位置が前進しなくなり
'     無限ループになるため、ov >= tgt のときは ov = tgt \ 2 に丸める
'     (V1には無い追加の安全策。テスト「オーバーラップ検証」用の異常値にも
'     耐えるようにするため)。
' ============================================================================

Private Const MAX_CHUNK_CHARS As Long = 32000
Private Const DEFAULT_TARGET_CHARS As Long = 700

' ----------------------------------------------------------------------------
' ChunkPages - 複数ページをまとめてチャンク化する。
'   戻り値=チャンク数(0可)。chunksにはページごとの結果を連結して返す。
' ----------------------------------------------------------------------------
Public Function ChunkPages(pages() As ExtractedPage, ByVal targetChars As Long, _
                           ByVal overlapChars As Long, ByRef chunks() As ShelfChunk) As Long
    Dim tgt As Long: tgt = targetChars
    If tgt < 1 Then tgt = DEFAULT_TARGET_CHARS
    If tgt > MAX_CHUNK_CHARS Then tgt = MAX_CHUNK_CHARS

    Dim ov As Long: ov = overlapChars
    If ov < 0 Then ov = 0
    If ov >= tgt Then ov = tgt \ 2

    Dim outArr() As ShelfChunk: ReDim outArr(0 To 15)
    Dim outCount As Long: outCount = 0

    Dim pageCount As Long: pageCount = PageCountOf(pages)
    If pageCount > 0 Then
        Dim pLo As Long: pLo = LBound(pages)
        Dim p As Long
        For p = 0 To pageCount - 1
            ChunkOnePage pages(pLo + p).page, pages(pLo + p).Text, tgt, ov, outArr, outCount
        Next p
    End If

    If outCount = 0 Then
        ReDim chunks(0 To -1)
    Else
        ReDim Preserve outArr(0 To outCount - 1)
        chunks = outArr
    End If
    ChunkPages = outCount
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' 1ページ分の平文を、段落・文境界を優先しつつ target/overlap でスライディング
' ウィンドウ分割し、結果を outArr(ByRef)へ追記する(V1 ChunkText を踏襲)。
Private Sub ChunkOnePage(ByVal pageNum As Long, ByVal rawText As String, ByVal tgt As Long, _
                         ByVal ov As Long, ByRef outArr() As ShelfChunk, ByRef outCount As Long)
    Dim normalized As String: normalized = NormalizeWhitespace(rawText)
    Dim total As Long: total = Len(normalized)
    If total = 0 Then Exit Sub

    Dim startPos As Long: startPos = 1
    Do While startPos <= total
        Dim windowEnd As Long: windowEnd = startPos + tgt - 1
        If windowEnd > total Then windowEnd = total

        ' 文境界(句読点・改行)がtarget/4字の余裕内にあれば、そこへ寄せる
        If windowEnd < total Then
            Dim boundary As Long
            boundary = FindSentenceBoundary(normalized, windowEnd, tgt \ 4)
            If boundary > 0 Then windowEnd = boundary
        End If

        ' 32000字上限の絶対保証(targetCharsが異常に大きい場合の保険)
        If windowEnd - startPos + 1 > MAX_CHUNK_CHARS Then
            windowEnd = startPos + MAX_CHUNK_CHARS - 1
        End If

        Dim ck As ShelfChunk
        ck.page = pageNum
        ck.full_text = Mid$(normalized, startPos, windowEnd - startPos + 1)
        AppendChunk outArr, outCount, ck

        If windowEnd >= total Then Exit Do

        Dim nextStart As Long: nextStart = windowEnd - ov + 1
        If nextStart <= startPos Then nextStart = startPos + 1   ' 無限ループ防止
        startPos = nextStart
    Loop
End Sub

' 空白圧縮(半角空白・タブの連続を1個に)+改行統一+3連続以上の改行を2個へ
' 圧縮(段落区切りは残す)+Trim。modUtil.NormalizeForHashと同じ配列+Join方式
' (§12: &連鎖の長大化禁止)。
Private Function NormalizeWhitespace(ByVal s As String) As String
    Dim t As String: t = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
    Dim n As Long: n = Len(t)
    If n = 0 Then
        NormalizeWhitespace = ""
        Exit Function
    End If

    Dim outChars() As String: ReDim outChars(1 To n)
    Dim outCount As Long: outCount = 0
    Dim prevWasSpace As Boolean: prevWasSpace = False
    Dim i As Long
    For i = 1 To n
        Dim c As String: c = Mid$(t, i, 1)
        If c = " " Or c = vbTab Then
            If Not prevWasSpace Then
                outCount = outCount + 1
                outChars(outCount) = " "
            End If
            prevWasSpace = True
        Else
            outCount = outCount + 1
            outChars(outCount) = c
            prevWasSpace = False
        End If
    Next i

    Dim joined As String
    If outCount > 0 Then
        ReDim Preserve outChars(1 To outCount)
        joined = Join(outChars, "")
    End If

    Do While InStr(joined, vbLf & vbLf & vbLf) > 0
        joined = Replace(joined, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop

    NormalizeWhitespace = Trim$(joined)
End Function

' nearPos付近(+slack字まで)で文の区切り(句読点・改行)を探す。見つからなければ0。
Private Function FindSentenceBoundary(ByVal s As String, ByVal nearPos As Long, ByVal slack As Long) As Long
    Dim limit As Long: limit = nearPos + slack
    If limit > Len(s) Then limit = Len(s)
    Dim i As Long
    For i = nearPos To limit
        Dim ch As String: ch = Mid$(s, i, 1)
        If ch = "." Or ch = "!" Or ch = "?" Or ch = ChrW(&H3002) Or ch = ChrW(&HFF01) Or ch = ChrW(&HFF1F) Or ch = vbLf Then
            FindSentenceBoundary = i
            Exit Function
        End If
    Next i
    FindSentenceBoundary = 0
End Function

' 可変長配列outArrへckを追記する(容量不足時は倍々でReDim Preserve)。
Private Sub AppendChunk(ByRef outArr() As ShelfChunk, ByRef outCount As Long, ByRef ck As ShelfChunk)
    If outCount > UBound(outArr) Then
        ReDim Preserve outArr(0 To (UBound(outArr) + 1) * 2 - 1)
    End If
    outArr(outCount) = ck
    outCount = outCount + 1
End Sub

' 未初期化/空配列でも例外にせず0件として扱う(LBound/UBoundの実行時エラー9対策)。
Private Function PageCountOf(pages() As ExtractedPage) As Long
    On Error GoTo Empty0
    PageCountOf = UBound(pages) - LBound(pages) + 1
    Exit Function
Empty0:
    PageCountOf = 0
End Function
