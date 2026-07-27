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
' 構造チャンクの資料名プレースホルダ(modShelfが実ファイル名へ置換する)
Private Const CRUMB_PLACEHOLDER As String = "〔資料〕"
Private Const FW_ZERO As Long = 65296      ' ChrW(&HFF10) ０
Private Const FW_NINE As Long = 65305      ' ChrW(&HFF19) ９


' ----------------------------------------------------------------------------
' ChunkPages - 複数ページをまとめてチャンク化する(後方互換=legacy固定)。
'   戻り値=チャンク数(0可)。chunksにはページごとの結果を連結して返す。
' ----------------------------------------------------------------------------
Public Function ChunkPages(pages() As ExtractedPage, ByVal targetChars As Long, _
                           ByVal overlapChars As Long, ByRef chunks() As ShelfChunk) As Long
    ChunkPages = ChunkPagesEx(pages, targetChars, overlapChars, 0, "legacy", chunks)
End Function

' ----------------------------------------------------------------------------
' ChunkPagesEx - モード切替つきチャンク化(設計書§B-1)。
'   mode="structure": 見出し・条文・表の構造を認識し、意味単位で分割。
'     各チャンク先頭に breadcrumb「【〔資料〕 > 章 > 条】」を前置する
'     (〔資料〕は取込側=modShelfが実ファイル名へ置換するプレースホルダ)。
'   mode=それ以外: 従来の700字スライディング分割(legacy)と完全同一。
'   maxCharsは構造モードで「条文を分割せず1チャンクに収める上限」。
' ----------------------------------------------------------------------------
Public Function ChunkPagesEx(pages() As ExtractedPage, ByVal targetChars As Long, _
                             ByVal overlapChars As Long, ByVal maxChars As Long, _
                             ByVal mode As String, ByRef chunks() As ShelfChunk) As Long
    Dim tgt As Long: tgt = targetChars
    If tgt < 1 Then tgt = DEFAULT_TARGET_CHARS
    If tgt > MAX_CHUNK_CHARS Then tgt = MAX_CHUNK_CHARS

    Dim ov As Long: ov = overlapChars
    If ov < 0 Then ov = 0
    If ov >= tgt Then ov = tgt \ 2

    Dim mx As Long: mx = maxChars
    If mx < tgt Then mx = tgt
    If mx > MAX_CHUNK_CHARS Then mx = MAX_CHUNK_CHARS

    Dim useStructure As Boolean
    useStructure = (LCase$(Trim$(mode)) = "structure")

    Dim outArr() As ShelfChunk: ReDim outArr(0 To 15)
    Dim outCount As Long: outCount = 0

    Dim pageCount As Long: pageCount = PageCountOf(pages)
    If pageCount > 0 Then
        Dim pLo As Long: pLo = LBound(pages)
        Dim p As Long
        If useStructure Then
            ' 条文はページをまたぐのが普通。ページ単位で切ると、条の途中で
            ' 必ず割れる(実測: 実物4資料で151ブロックが分断されていた)。
            ' 割れた後半チャンクは見出しを持たないので、「第12条は?」に対して
            ' 条文の後半だけが当たる、という最悪の外し方をする。
            ' ページを跨いでブロックを組み立て、ページ番号は各ブロックの
            ' 先頭行が属するページを引き継ぐ(出典表示はそのままで良い)。
            ChunkAllPagesStructured pages, pLo, pageCount, tgt, ov, mx, outArr, outCount
        Else
            For p = 0 To pageCount - 1
                ChunkOnePage pages(pLo + p).page, pages(pLo + p).Text, tgt, ov, outArr, outCount
            Next p
        End If
    End If

    If outCount = 0 Then
        ReDim chunks(0 To 0)
    Else
        ReDim Preserve outArr(0 To outCount - 1)
        chunks = outArr
    End If
    ChunkPagesEx = outCount
End Function

' ============================================================================
' 取込の正規化(2026-07-27追加。実物の約款PDFで計測して判明した欠陥への対処)
' ============================================================================
' 実測(tools/audit_ingest.py・実物の約款2点):
'   ・条見出し 252行のうち 51行(20.2%)を見出しとして認識できていなかった
'     内訳はほぼ全てが2桁の条。PDFの字詰めで「第１ １条」「第 １ ０ 条」の
'     ように数字のあいだへ空白が入り、MatchesDaiN(数字の連続を要求)が外れる。
'     第1条〜第9条は通り、第10条以降が全滅する ―― 約款の中身の大半である。
'   ・「- 19 -」のようなページ番号行が本文に混じり、条文を分断していた。
'
' 見出しを取りこぼすと、その条は本文として隣の条にくっつき、
' 「第12条は?」に対して別の条の本文が返る。利用者からは、なぜ外したのか
' 永久に分からない。入力の欠陥は出力の全部に効くので、ここで必ず潰す。
'
' 純ロジック(R4)なので LibreOffice でそのままテストできる。

' NormalizeForIngest - 取込テキストの正規化。全ての取込経路が最初に通す。
Public Function NormalizeForIngest(ByVal s As String) As String
    Dim t As String
    t = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)

    Dim rows() As String: rows = Split(t, vbLf)
    Dim keep() As String: ReDim keep(0 To UBound(rows))
    Dim n As Long
    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        If Not IsPageNumberLine(rows(i)) Then
            keep(n) = JoinSplitNumbers(rows(i))
            n = n + 1
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve keep(0 To n - 1)
    NormalizeForIngest = Join(keep, vbLf)
End Function

' JoinSplitNumbers - 「第 １ ０ 条」→「第10条」。第〜条/章/節/項/号 の
'   あいだにある空白だけを取り除く(本文の空白には触れない)。
'   全角数字は半角へ寄せる(検索側も半角で来るため)。
Public Function JoinSplitNumbers(ByVal lineText As String) As String
    JoinSplitNumbers = lineText
    Dim p As Long: p = InStr(lineText, "第")
    If p = 0 Then Exit Function

    Dim out As String
    Dim i As Long: i = 1
    Dim ln As Long: ln = Len(lineText)
    Do While i <= ln
        Dim c As String: c = Mid$(lineText, i, 1)
        If c <> "第" Then
            out = out & c
            i = i + 1
        Else
            ' 「第」の後ろを走査: 数字と空白だけが続き、そのあと単位漢字が来るか
            Dim j As Long: j = i + 1
            Dim digits As String
            Do While j <= ln
                Dim d As String: d = Mid$(lineText, j, 1)
                If IsDigitChar(d) Then
                    digits = digits & HalfDigit(d)
                    j = j + 1
                ElseIf d = " " Or d = ChrW(&H3000) Or d = vbTab Then
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop
            Dim unit As String
            If j <= ln Then unit = Mid$(lineText, j, 1)
            If LenB(digits) > 0 And (unit = "条" Or unit = "章" Or unit = "節" _
                                     Or unit = "項" Or unit = "号" Or unit = "編") Then
                out = out & "第" & digits & unit
                i = j + 1
            Else
                out = out & c
                i = i + 1
            End If
        End If
    Loop
    JoinSplitNumbers = out
End Function

' IsPageNumberLine - 「- 19 -」「—19—」「‐ 3 ‐」等、ページ番号だけの行か。
Public Function IsPageNumberLine(ByVal lineText As String) As Boolean
    Dim t As String: t = Trim$(lineText)
    t = Replace(Replace(t, ChrW(&H3000), ""), " ", "")
    If Len(t) < 3 Or Len(t) > 8 Then Exit Function
    If Not IsDashChar(Left$(t, 1)) Then Exit Function
    If Not IsDashChar(Right$(t, 1)) Then Exit Function
    Dim mid_ As String: mid_ = Mid$(t, 2, Len(t) - 2)
    If LenB(mid_) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(mid_)
        If Not IsDigitChar(Mid$(mid_, i, 1)) Then Exit Function
    Next i
    IsPageNumberLine = True
End Function

Private Function IsDashChar(ByVal c As String) As Boolean
    IsDashChar = (c = "-" Or c = ChrW(&H2010) Or c = ChrW(&H2011) Or c = ChrW(&H2012) _
                  Or c = ChrW(&H2013) Or c = ChrW(&H2014) Or c = ChrW(&H2015) _
                  Or c = ChrW(&H30FC) Or c = ChrW(&HFF0D) Or c = ChrW(&H2212))
End Function

' 全角数字を半角へ寄せる(検索語は半角で来るため、本文側を揃える)。
Private Function HalfDigit(ByVal c As String) As String
    Dim code As Long: code = CodePointOf(c)
    If code >= FW_ZERO And code <= FW_NINE Then
        HalfDigit = Chr$(48 + (code - FW_ZERO))
    Else
        HalfDigit = c
    End If
End Function

' ----------------------------------------------------------------------------
' ClassifyLine - 行の構造ラベル分類(設計書§B-1。純関数・テスト用にPublic)。
'   0=本文 / 1=文書見出し(#・第N編/章・【…】のみの行) /
'   2=節見出し(第N条/節・番号見出し・■●◆短行) / 3=箇条書き / 4=表行
' ----------------------------------------------------------------------------
Public Function ClassifyLine(ByVal lineText As String) As Long
    ClassifyLine = 0
    Dim t As String: t = Trim$(lineText)
    If LenB(t) = 0 Then Exit Function

    ' 表行(タブ・3連続以上スペース・パイプ)
    If InStr(lineText, vbTab) > 0 Or InStr(lineText, "   ") > 0 Or InStr(lineText, "|") > 0 Then
        ClassifyLine = 4
        Exit Function
    End If

    ' 文書見出し
    If Left$(t, 2) = "# " Then ClassifyLine = 1: Exit Function
    If Left$(t, 1) = "【" And Right$(t, 1) = "】" And Len(t) <= 60 Then ClassifyLine = 1: Exit Function
    If MatchesDaiN(t, "編") Or MatchesDaiN(t, "章") Then ClassifyLine = 1: Exit Function

    ' 節見出し
    If MatchesDaiN(t, "条") Or MatchesDaiN(t, "節") Then ClassifyLine = 2: Exit Function
    If IsNumberHeading(t) Then ClassifyLine = 2: Exit Function
    Dim mark As String: mark = Left$(t, 1)
    If (mark = "■" Or mark = "●" Or mark = "◆") And Len(t) < 40 And InStr(t, "。") = 0 Then
        ClassifyLine = 2
        Exit Function
    End If

    ' 箇条書き
    If MatchesDaiN(t, "項") Then ClassifyLine = 3: Exit Function
    If IsItemMarker(t) Then ClassifyLine = 3: Exit Function
End Function

' ----------------------------------------------------------------------------
' BuildBreadcrumb - 「【資料 > 章 > 条】」のbreadcrumb文字列(設計書§B-2)。
'   空の階層は省略。各要素は80字で切る。
' ----------------------------------------------------------------------------
Public Function BuildBreadcrumb(ByVal sourceName As String, ByVal chapter As String, _
                                ByVal section As String) As String
    Dim joined As String
    joined = modUtil.SafeLeft(Trim$(sourceName), 80)
    If LenB(Trim$(chapter)) > 0 Then joined = joined & " > " & modUtil.SafeLeft(Trim$(chapter), 80)
    If LenB(Trim$(section)) > 0 Then joined = joined & " > " & modUtil.SafeLeft(Trim$(section), 80)
    BuildBreadcrumb = "【" & joined & "】"
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' 1ページ分の平文を、段落・文境界を優先しつつ target/overlap でスライディング
' ウィンドウ分割し、結果を outArr(ByRef)へ追記する(V1 ChunkText を踏襲)。
Private Sub ChunkOnePage(ByVal pageNum As Long, ByVal rawText As String, ByVal tgt As Long, _
                         ByVal ov As Long, ByRef outArr() As ShelfChunk, ByRef outCount As Long)
    Dim normalized As String: normalized = NormalizeWhitespace(NormalizeForIngest(rawText))
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

' ----------------------------------------------------------------------------
' ChunkAllPagesStructured - 全ページを1つの行列として扱い、見出し境界で割る。
'   ページ境界では絶対に切らない。行ごとに出所ページを持ち回り、ブロックの
'   ページ番号は「そのブロックの先頭行が載っていたページ」にする。
' ----------------------------------------------------------------------------
Private Sub ChunkAllPagesStructured(pages() As ExtractedPage, ByVal pLo As Long, _
                                    ByVal pageCount As Long, ByVal tgt As Long, _
                                    ByVal ov As Long, ByVal mx As Long, _
                                    ByRef outArr() As ShelfChunk, ByRef outCount As Long)
    Dim chapter As String, section As String
    Dim blockText As String
    Dim blockPage As Long

    Dim p As Long
    For p = 0 To pageCount - 1
        Dim pageNo As Long: pageNo = pages(pLo + p).page
        Dim raw As String
        raw = NormalizeForIngest(pages(pLo + p).Text)
        If LenB(Trim$(raw)) > 0 Then
            Dim rows() As String
            rows = Split(Replace(Replace(raw, vbCrLf, vbLf), vbCr, vbLf), vbLf)

            Dim i As Long
            For i = LBound(rows) To UBound(rows)
                Dim lbl As Long: lbl = ClassifyLine(rows(i))
                If lbl = 1 Or lbl = 2 Then
                    FlushBlock blockPage, blockText, chapter, section, tgt, ov, mx, outArr, outCount
                    If lbl = 1 Then
                        chapter = Trim$(StripHeadingMark(rows(i)))
                        section = ""
                    Else
                        section = Trim$(rows(i))
                    End If
                    blockText = Trim$(rows(i))
                    blockPage = pageNo
                Else
                    Dim rowText As String
                    If lbl = 4 Then
                        rowText = RTrimOnly(rows(i))
                    Else
                        rowText = CollapseSpaces(rows(i))
                    End If
                    If LenB(Trim$(rowText)) > 0 Then
                        If LenB(blockText) = 0 Then blockPage = pageNo
                        If LenB(blockText) > 0 Then blockText = blockText & vbLf
                        blockText = blockText & rowText
                    End If
                End If
            Next i
        End If
    Next p
    FlushBlock blockPage, blockText, chapter, section, tgt, ov, mx, outArr, outCount
End Sub

' ----------------------------------------------------------------------------
' 構造認識チャンク化の内部実装(設計書§B-1)
' ----------------------------------------------------------------------------

' 1ページを行分類→見出し境界のブロック単位で分割する。ブロックが
' maxChars以下なら1チャンクに原子保持、超過時のみ文境界スライディング
' (オーバーラップは同一ブロック内のみ)。表行はスペース構造を保持する。
Private Sub ChunkOnePageStructured(ByVal pageNum As Long, ByVal rawText As String, _
                                   ByVal tgt As Long, ByVal ov As Long, ByVal mx As Long, _
                                   ByRef outArr() As ShelfChunk, ByRef outCount As Long)
    Dim t As String: t = Replace(Replace(rawText, vbCrLf, vbLf), vbCr, vbLf)
    If LenB(Trim$(t)) = 0 Then Exit Sub

    Dim rows() As String: rows = Split(t, vbLf)

    Dim chapter As String: chapter = ""
    Dim section As String: section = ""
    Dim blockText As String: blockText = ""

    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        Dim lbl As Long: lbl = ClassifyLine(rows(i))

        If lbl = 1 Or lbl = 2 Then
            ' 見出し=ブロック境界。ここまでのブロックを確定し、階層を更新。
            FlushBlock pageNum, blockText, chapter, section, tgt, ov, mx, outArr, outCount
            If lbl = 1 Then
                chapter = Trim$(StripHeadingMark(rows(i)))
                section = ""
            Else
                section = Trim$(rows(i))
            End If
            ' 見出し行自体は直後ブロックの先頭に含める(見出しだけの空チャンクは作らない)
            blockText = Trim$(rows(i))
        Else
            Dim rowText As String
            If lbl = 4 Then
                rowText = RTrimOnly(rows(i))     ' 表行: スペース構造を保持
            Else
                rowText = CollapseSpaces(rows(i))
            End If
            If LenB(Trim$(rowText)) > 0 Then
                If LenB(blockText) > 0 Then blockText = blockText & vbLf
                blockText = blockText & rowText
            End If
        End If
    Next i
    FlushBlock pageNum, blockText, chapter, section, tgt, ov, mx, outArr, outCount
End Sub

' ブロック1つをチャンク列へ確定する。breadcrumb(プレースホルダ資料名)を
' 各チャンク先頭に前置。32000字上限は絶対保証。
Private Sub FlushBlock(ByVal pageNum As Long, ByVal blockText As String, _
                       ByVal chapter As String, ByVal section As String, _
                       ByVal tgt As Long, ByVal ov As Long, ByVal mx As Long, _
                       ByRef outArr() As ShelfChunk, ByRef outCount As Long)
    Dim body As String: body = Trim$(blockText)
    If LenB(body) = 0 Then Exit Sub

    Dim crumb As String
    crumb = BuildBreadcrumb(CRUMB_PLACEHOLDER, chapter, section) & vbLf

    Dim room As Long: room = MAX_CHUNK_CHARS - Len(crumb)
    If room < 1 Then room = 1
    Dim atomicLimit As Long: atomicLimit = mx - Len(crumb)
    If atomicLimit < 1 Then atomicLimit = 1

    If Len(body) <= atomicLimit Then
        AppendStructChunk pageNum, crumb, body, room, outArr, outCount
        Exit Sub
    End If

    ' 超過ブロックのみ文境界スライディング(改行も境界候補=表行を跨ぎにくい)
    Dim total As Long: total = Len(body)
    Dim startPos As Long: startPos = 1
    Do While startPos <= total
        Dim windowEnd As Long: windowEnd = startPos + tgt - 1
        If windowEnd > total Then windowEnd = total
        If windowEnd < total Then
            Dim boundary As Long
            boundary = FindSentenceBoundary(body, windowEnd, tgt \ 4)
            If boundary > 0 Then windowEnd = boundary
        End If
        If windowEnd - startPos + 1 > room Then windowEnd = startPos + room - 1

        AppendStructChunk pageNum, crumb, Mid$(body, startPos, windowEnd - startPos + 1), room, outArr, outCount

        If windowEnd >= total Then Exit Do
        Dim nextStart As Long: nextStart = windowEnd - ov + 1
        If nextStart <= startPos Then nextStart = startPos + 1
        startPos = nextStart
    Loop
End Sub

Private Sub AppendStructChunk(ByVal pageNum As Long, ByVal crumb As String, ByVal body As String, _
                              ByVal room As Long, ByRef outArr() As ShelfChunk, ByRef outCount As Long)
    Dim b As String: b = body
    If Len(b) > room Then b = Left$(b, room)
    Dim ck As ShelfChunk
    ck.page = pageNum
    ck.full_text = crumb & b
    AppendChunk outArr, outCount, ck
End Sub

' 「第N条」「第１２章」等: 先頭が「第」+数字(半角/全角)+指定漢字か。
Private Function MatchesDaiN(ByVal t As String, ByVal kanji As String) As Boolean
    If Left$(t, 1) <> "第" Then Exit Function
    Dim i As Long: i = 2
    Dim digitN As Long: digitN = 0
    Do While i <= Len(t)
        If IsDigitChar(Mid$(t, i, 1)) Then
            digitN = digitN + 1
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    If digitN = 0 Then Exit Function
    MatchesDaiN = (Mid$(t, i, Len(kanji)) = kanji)
End Function

' 「1.」「2.3」「３．」等の番号見出し(短い行のみ=通常の文中番号と区別)。
Private Function IsNumberHeading(ByVal t As String) As Boolean
    If Len(t) >= 50 Then Exit Function
    If Not IsDigitChar(Left$(t, 1)) Then Exit Function
    Dim i As Long: i = 1
    Do While i <= Len(t) And IsDigitChar(Mid$(t, i, 1))
        i = i + 1
    Loop
    If i > Len(t) Then Exit Function
    Dim sep As String: sep = Mid$(t, i, 1)
    IsNumberHeading = (sep = "." Or sep = "．")
End Function

' 箇条書きマーカー: ・ / - / (N) / （N） / ①〜⑳ 始まり。
Private Function IsItemMarker(ByVal t As String) As Boolean
    Dim h As String: h = Left$(t, 1)
    If h = "・" Then IsItemMarker = True: Exit Function
    If Left$(t, 2) = "- " Then IsItemMarker = True: Exit Function
    If h = "(" Or h = "（" Then
        IsItemMarker = IsDigitChar(Mid$(t, 2, 1))
        Exit Function
    End If
    Dim code As Long: code = AscW(h)
    If code < 0 Then code = code + 65536
    IsItemMarker = (code >= &H2460 And code <= &H2473)   ' ①〜⑳
End Function

' 半角/全角数字か。
' 全角数字の定数は必ず10進で書く。
' VBAもLibreOffice Basicも「&HFF10」を16bit整数として解釈するため、
' この値は -240 になる。したがって
'     code >= &HFF10 And code <= &HFF19
' は「65296 >= -240 かつ 65296 <= -231」= False となり、全角数字は
' 一度も真にならなかった(2026-07-27発見)。
' 影響: 「第１０条」のように全角数字で書かれた条見出しが、これまで一切
' 見出しとして認識されていなかった。約款は全角で組まれることが多く、
' その場合その資料は丸ごと構造を失って固定長分割に退化していた。

Private Function IsDigitChar(ByVal ch As String) As Boolean
    If LenB(ch) = 0 Then Exit Function
    If ch >= "0" And ch <= "9" Then IsDigitChar = True: Exit Function
    IsDigitChar = (CodePointOf(ch) >= FW_ZERO And CodePointOf(ch) <= FW_NINE)
End Function

' AscW は符号付きIntegerを返すため、BMP後半の文字は負値になる。必ず補正する。
Private Function CodePointOf(ByVal ch As String) As Long
    If LenB(ch) = 0 Then Exit Function
    Dim code As Long: code = AscW(ch)
    If code < 0 Then code = code + 65536
    CodePointOf = code
End Function

' Markdown見出しマーク「# 」を剥がす(それ以外はそのまま)。
Private Function StripHeadingMark(ByVal t As String) As String
    Dim s As String: s = Trim$(t)
    Do While Left$(s, 1) = "#"
        s = Mid$(s, 2)
    Loop
    StripHeadingMark = Trim$(s)
End Function

' 行内の半角空白・タブ連続を1個へ(表行以外の本文行整形)。
Private Function CollapseSpaces(ByVal s As String) As String
    Dim t As String: t = Replace(s, vbTab, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CollapseSpaces = Trim$(t)
End Function

' 右端の空白だけ落とす(表行用。左のインデント・桁揃えは保持)。
Private Function RTrimOnly(ByVal s As String) As String
    Dim i As Long: i = Len(s)
    Do While i > 0
        Dim ch As String: ch = Mid$(s, i, 1)
        If ch = " " Or ch = vbTab Then
            i = i - 1
        Else
            Exit Do
        End If
    Loop
    RTrimOnly = Left$(s, i)
End Function

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
