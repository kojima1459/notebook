Attribute VB_Name = "modChunkPage"
Option Explicit

' ============================================================================
' modChunkPage - 見出しの無い複数ページ資料の「かけらごとの物理ページ」計算
'   (R42 §1 A1/A2・受入FAIL PDF-F01/RT-27/28/37/スモーク5 対応)。
' ----------------------------------------------------------------------------
' 何が壊れていたか:
'   modChunker.ChunkAllPagesStructured は行の集まり(ブロック)を「見出し行」か
'   「ブロック先頭行」でしかページ番号を更新していなかった。見出しの無い資料は
'   全ページが1ブロックへ連結され、そのブロックをスライディング窓で割った
'   FlushBlock がどのかけらにも同じ pageNum(=ブロック先頭のページ)を付けて
'   いた。結果、後半ページだけの本文で質問しても出典が常に p.1 になる。
'
' 直し方:
'   ChunkAllPagesStructured は行ごとの物理ページを並走配列 blockPg() として
'   持ち回る(本体は modChunker 側。ここは器と純関数だけを持つ)。ブロックを
'   確定するとき、そのかけらの本文が実際に開始する物理ページを、行の開始位置
'   (LineStarts と同じ規則)と各行のページ(blockPg)から逆引きする
'   (PageAtPos)。見出し行から始まるかけらは見出し行のページ、オーバーラップで
'   前ページの末尾から始まるかけらは前ページになる(R42 spec §1-3)。
'
' なぜ modChunker 本体ではなく新モジュールか:
'   modChunker は残61字(CLAUDE.md §12・実質凍結)で1行も入らない。実体は
'   ここへ置き、modChunker 側は BlockAdd→BlockAddPg・FlushBlock→FlushPaged の
'   呼び替えだけに留める(§1-4)。
'
' 契約(tools/vba_lint.py の CONTRACT["modChunkPage"]で固定):
'   BlockAddPg / LineStarts / PageAtPos / FirstInkPos / PlanWindows /
'   FlushPaged / PhysicalKeep / StartPagesOf
'   (StartPagesOf は modTestsPure45 専用の検証口。5ページ・見出し無しの
'   ゴールデンを LO 上で固定するために置く)。
' ============================================================================

' BlockAddPg - BlockAdd(modChunker)と同じ伸長規則で、行(buf)と物理ページ
'   (pg)を並走させて積む。空文字は積まない(BlockAddと同一挙動)。
Public Sub BlockAddPg(ByRef buf() As String, ByRef pg() As Long, ByRef n As Long, _
                      ByVal s As String, ByVal page As Long)
    If LenB(s) = 0 Then Exit Sub
    If n = 0 Then
        ReDim buf(1 To 64)
        ReDim pg(1 To 64)
    End If
    If n >= UBound(buf) Then
        ReDim Preserve buf(1 To UBound(buf) * 2)
        ReDim Preserve pg(1 To UBound(pg) * 2)
    End If
    n = n + 1
    buf(n) = s
    pg(n) = page
End Sub

' LineStarts - 純関数の検証口。Join(buf(1..n), vbLf) における各行の開始位置
'   (1始まり)を "1,12,30" の形で返す(行iの開始 = 前行の開始 + Len(前行) + 1)。
'   FlushPaged 自体は文字列を経由せず同じ規則を Long 配列で計算する(下記)。
Public Function LineStarts(ByRef buf() As String, ByVal n As Long) As String
    If n < 1 Then Exit Function
    Dim result As String
    Dim pos As Long: pos = 1
    Dim i As Long
    For i = 1 To n
        If i > 1 Then result = result & ","
        result = result & CStr(pos)
        pos = pos + Len(buf(i)) + 1
    Next i
    LineStarts = result
End Function

' PageAtPos - starts(i) <= pos を満たす最大の i の pg(i)。pos が starts(1)より
'   小さければ pg(1)。n=0 は 0。starts は非減少(LineStarts と同じ規則で作った
'   もの)前提なので、超えた時点で走査を打ち切ってよい。
Public Function PageAtPos(ByRef starts() As Long, ByRef pg() As Long, _
                          ByVal n As Long, ByVal pos As Long) As Long
    If n < 1 Then Exit Function
    If pos < starts(1) Then
        PageAtPos = pg(1)
        Exit Function
    End If
    Dim i As Long, best As Long: best = 1
    For i = 1 To n
        If starts(i) <= pos Then
            best = i
        Else
            Exit For
        End If
    Next i
    PageAtPos = pg(best)
End Function

' FirstInkPos - startPos 以降で空白類(半角/全角スペース・vbLf・vbCr・vbTab)
'   でない最初の位置。無ければ startPos(全部空白類・startPosが末尾超え含む)。
Public Function FirstInkPos(ByVal body As String, ByVal startPos As Long) As Long
    Dim n As Long: n = Len(body)
    Dim i As Long, ch As String
    For i = startPos To n
        ch = Mid$(body, i, 1)
        If ch <> " " And ch <> ChrW$(&H3000) And ch <> vbLf And ch <> vbCr And ch <> vbTab Then
            FirstInkPos = i
            Exit Function
        End If
    Next i
    FirstInkPos = startPos
End Function

' PlanWindows - modChunker.FlushBlock:582-600 のスライディング規則を1文字も
'   変えずに移植した純関数版(FindSentenceBoundary を呼ぶため Public 化済み)。
'   本文の粒度はR42で変えない(受入報告§3の前提)。各窓を "s-e|s-e|…" で返す。
Public Function PlanWindows(ByVal total As Long, ByVal tgt As Long, ByVal ov As Long, _
                            ByVal room As Long, ByVal body As String) As String
    Dim result As String
    Dim startPos As Long: startPos = 1
    Do While startPos <= total
        Dim windowEnd As Long: windowEnd = startPos + tgt - 1
        If windowEnd > total Then windowEnd = total
        If windowEnd < total Then
            Dim boundary As Long
            boundary = modChunker.FindSentenceBoundary(body, windowEnd, tgt \ 4)
            If boundary > 0 Then windowEnd = boundary
        End If
        If windowEnd - startPos + 1 > room Then windowEnd = startPos + room - 1

        If LenB(result) > 0 Then result = result & "|"
        result = result & CStr(startPos) & "-" & CStr(windowEnd)

        If windowEnd >= total Then Exit Do
        Dim nextStart As Long: nextStart = windowEnd - ov + 1
        If nextStart <= startPos Then nextStart = startPos + 1
        startPos = nextStart
    Loop
    PlanWindows = result
End Function

' FlushPaged - modChunker.FlushBlock の後継(ページ持ち回り版)。ブロック
'   (buf/pg/n)を1つ確定し、原子保持または PlanWindows のスライディングで
'   outArr(ShelfChunk)へ追記する。呼び出し後 n=0(BlockTake と同じ契約)。
Public Sub FlushPaged(ByRef buf() As String, ByRef pg() As Long, ByRef n As Long, _
                      ByVal chapter As String, ByVal section As String, _
                      ByVal tgt As Long, ByVal ov As Long, ByVal mx As Long, _
                      ByRef outArr() As ShelfChunk, ByRef outCount As Long)
    If n < 1 Then Exit Sub

    ' レビュー R42 m8: 本文の組み立て・窓割り・ページ逆引きは StartPagesOf と
    ' 同じ Private 実装(PrepareBlock/PagesForPlan)を通す。LO が固定するのは
    ' StartPagesOf 側だが、経路が同一なのでここだけ将来ずれることがない。
    Dim body As String, lead As Long, starts() As Long
    Dim cnt As Long: cnt = n
    n = 0
    If Not PrepareBlock(buf, pg, cnt, body, lead, starts) Then Exit Sub

    ' crumb/room/atomicLimit は modChunker.FlushBlock:568-574 と同一
    ' (BuildBreadcrumb/CRUMB_PLACEHOLDER/MAX_CHUNK_CHARS は Public 化済み)。
    Dim crumb As String
    crumb = modChunker.BuildBreadcrumb(modChunker.CRUMB_PLACEHOLDER, chapter, section) & vbLf
    Dim room As Long: room = modChunker.MAX_CHUNK_CHARS - Len(crumb)
    If room < 1 Then room = 1
    Dim atomicLimit As Long: atomicLimit = mx - Len(crumb)
    If atomicLimit < 1 Then atomicLimit = 1

    Dim plan As String: plan = PlanFor(body, tgt, ov, room, atomicLimit)
    Dim pageList As String: pageList = PagesForPlan(plan, body, lead, starts, pg, cnt)

    Dim wins() As String: wins = Split(plan, "|")
    Dim pgs() As String: pgs = Split(pageList, ",")
    Dim w As Long
    For w = LBound(wins) To UBound(wins)
        Dim se() As String: se = Split(wins(w), "-")
        Dim s As Long: s = CLng(se(0))
        Dim e As Long: e = CLng(se(1))
        modChunker.AppendStructChunk CLng(pgs(w)), crumb, Mid$(body, s, e - s + 1), room, outArr, outCount
    Next w
End Sub

' PrepareBlock - ブロック(buf/pg/cnt)を本文へ組み立てる共通部(FlushPaged と
'   StartPagesOf の唯一の実装)。raw=Join(vbLf)、lead=先頭の半角空白数
'   (Trim$/LTrim$ は半角スペースだけを消すので body 座標 q ↔ raw 座標 q+lead
'   が厳密に成り立つ)、starts()=LineStarts と同じ規則の Long 配列。
'   本文が空なら False。
Private Function PrepareBlock(ByRef buf() As String, ByRef pg() As Long, ByVal cnt As Long, _
                              ByRef body As String, ByRef lead As Long, _
                              ByRef starts() As Long) As Boolean
    ReDim Preserve buf(1 To cnt)
    ReDim Preserve pg(1 To cnt)
    Dim raw As String: raw = Join(buf, vbLf)
    lead = Len(raw) - Len(LTrim$(raw))
    body = Trim$(raw)
    If LenB(body) = 0 Then Exit Function

    ReDim starts(1 To cnt)
    Dim pos As Long: pos = 1
    Dim i As Long
    For i = 1 To cnt
        starts(i) = pos
        pos = pos + Len(buf(i)) + 1
    Next i
    PrepareBlock = True
End Function

' PlanFor - 窓計画の唯一の実装(FlushPaged と StartPagesOf 共通・2周目 m9)。
'   本文が atomicLimit 以下なら原子保持=窓1つ "1-N"(FlushBlock:576 と同じ
'   判定)。atomicLimit<=0 は「常に分割」(StartPagesOf の従来挙動)。
Private Function PlanFor(ByVal body As String, ByVal tgt As Long, ByVal ov As Long, _
                         ByVal room As Long, ByVal atomicLimit As Long) As String
    If atomicLimit > 0 Then
        If Len(body) <= atomicLimit Then
            PlanFor = "1-" & CStr(Len(body))
            Exit Function
        End If
    End If
    PlanFor = PlanWindows(Len(body), tgt, ov, room, body)
End Function

' PagesForPlan - 窓計画("s-e|s-e|…")の各窓について、本文先頭(空白以外の最初
'   の文字)が載っている物理ページを "1,2,2,3,4" のカンマ区切りで返す
'   (FlushPaged と StartPagesOf の唯一の実装)。
Private Function PagesForPlan(ByVal plan As String, ByVal body As String, ByVal lead As Long, _
                              ByRef starts() As Long, ByRef pg() As Long, ByVal cnt As Long) As String
    Dim wins() As String: wins = Split(plan, "|")
    Dim result As String
    Dim w As Long
    For w = LBound(wins) To UBound(wins)
        Dim se() As String: se = Split(wins(w), "-")
        Dim s As Long: s = CLng(se(0))
        Dim pageN As Long: pageN = PageAtPos(starts, pg, cnt, FirstInkPos(body, s) + lead)
        If LenB(result) > 0 Then result = result & ","
        result = result & CStr(pageN)
    Next w
    PagesForPlan = result
End Function

' PhysicalKeep - 上限は物理ページ番号に当てる(R42 A2)。maxPages<=0は無制限
'   (truncated=False)。firstIdx+keptN が maxPages を超えるなら
'   truncated=True、戻りは maxPages-firstIdx(負は0=本文なし)。
Public Function PhysicalKeep(ByVal firstIdx As Long, ByVal keptN As Long, _
                             ByVal maxPages As Long, ByRef truncated As Boolean) As Long
    If maxPages <= 0 Then
        truncated = False
        PhysicalKeep = keptN
        Exit Function
    End If
    If firstIdx + keptN > maxPages Then
        truncated = True
        Dim r As Long: r = maxPages - firstIdx
        If r < 0 Then r = 0
        PhysicalKeep = r
        Exit Function
    End If
    truncated = False
    PhysicalKeep = keptN
End Function

' StartPagesOf - modTestsPure45 専用の検証口(§1-5 受入ゴールデン)。
'   BlockAddPg で組んだ buf/pg/n をそのまま PlanWindows へかけ、各窓の開始
'   ページ(PageAtPos+FirstInkPos)を "1,2,2,3,4" のようにカンマ区切りで返す。
'   FlushPaged と違い ShelfChunk を作らない(UDT を跨がずLO実行テストで
'   固定するための専用口)。atomicLimit は FlushPaged と同じ意味
'   (0=常に分割・>0 で本文がそれ以下なら窓1つ)。窓計画は PlanFor で共通。
Public Function StartPagesOf(ByRef buf() As String, ByRef pg() As Long, ByVal n As Long, _
                             ByVal tgt As Long, ByVal ov As Long, ByVal room As Long, _
                             Optional ByVal atomicLimit As Long = 0) As String
    If n < 1 Then Exit Function
    Dim body As String, lead As Long, starts() As Long
    If Not PrepareBlock(buf, pg, n, body, lead, starts) Then Exit Function
    Dim plan As String: plan = PlanFor(body, tgt, ov, room, atomicLimit)
    StartPagesOf = PagesForPlan(plan, body, lead, starts, pg, n)
End Function
