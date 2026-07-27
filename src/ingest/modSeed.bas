Attribute VB_Name = "modSeed"
Option Explicit

' ============================================================================
' modSeed - 初期ナレッジ(同梱シード)の展開と、初回の行き先づくり
' ----------------------------------------------------------------------------
' なぜ必要か:
'   配布した .xlsm を開いた瞬間、利用者は「で、何をすればいいの?」で止まる。
'   本棚は空、質問しても資料が無い、と言われて終わる。そこで離脱する。
'   「すごいですね」と言われて使われない反応の、いちばん手前の原因がこれ。
'
'   だから実物の社内資料を最初から積んで出す。ただし「それっぽい一般知識」は
'   絶対に入れない。損保で初手の回答が微妙にズレたら、その時点で信用は終わる。
'   入っているのは実物だけで、出典も実物のページを指す。
'
' 仕組み:
'   ビルド時に tools/make_seed_pack.py が作ったパックの中身を、
'   seed_meta / seed_chunks / seed_vectors の3シート(veryHidden)として
'   .xlsm に焼き込む。初回起動時にこのモジュールが my_knowledge /
'   my_vectors へ写すだけ。チャンク分割はビルド時に済んでいるので、
'   実行時に走るのは行のコピーだけ = 速い。
'
' ベクトルについて(ここが体験を分ける):
'   埋め込みは社内AIリボンでしか作れないため、ビルド環境では生成できない。
'     ・seed_vectors 入りで配布 → 初回起動は一瞬・API呼び出しゼロ(推奨)
'     ・空で配布           → 取り込みは一瞬だが検索に使えるまで
'                            埋め込みが要る。起動を止めず、次の同期で埋まる
'   どちらでも壊れないようにし、前者を既定の運用にする
'   (作り方は docs/60_初期ナレッジの作り方.md)。
'
' すぐ押せる質問:
'   資料名を並べて「聞いてみて」と言うだけでは、利用者はまだ
'   「何を聞けばいいか」を考えないといけない。課長の前で的外れな質問をする
'   リスクも負う。そこで、その資料が確実に答えられる質問をビルド時に生成して
'   同梱し、クリックだけで送信できるようにする。文面は実データ由来なので、
'   案内が嘘をつくことが構造的にありえない。
' ============================================================================

Private Const SH_SEED_META As String = "seed_meta"
Private Const SH_SEED_CHUNKS As String = "seed_chunks"
Private Const SH_SEED_VECTORS As String = "seed_vectors"
Private Const STAT_SEED_DONE As String = "seed_loaded"

' ----------------------------------------------------------------------------
' EnsureSeedLoaded - 初回だけ同梱シードを本棚へ展開する(2回目以降は即return)。
'   戻り値 = 取り込んだチャンク数(0 = 何もしなかった)
' ----------------------------------------------------------------------------
Public Function EnsureSeedLoaded() As Long
    On Error GoTo Done

    ' 済判定は my_stats。利用者がシード資料を自分で消した場合に
    ' 毎回復活すると鬱陶しいので、フラグは一度立てたら戻さない。
    If LenB(modStats.GetStatText(STAT_SEED_DONE)) > 0 Then Exit Function

    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = ThisWorkbook.Worksheets(SH_SEED_CHUNKS)
    On Error GoTo Done
    If wsC Is Nothing Then Exit Function          ' シード無しビルド = 何もしない

    Dim lastR As Long
    lastR = wsC.Cells(wsC.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then
        modStats.SetStatText STAT_SEED_DONE, "empty"
        Exit Function
    End If

    Dim wsK As Worksheet, wsV As Worksheet
    Set wsK = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    Set wsV = ThisWorkbook.Worksheets(modAppDef.SH_VECTORS)
    If wsK Is Nothing Then Exit Function

    Dim prevCalc As Long
    On Error Resume Next
    prevCalc = Application.Calculation
    Application.Calculation = -4135              ' xlCalculationManual
    Application.ScreenUpdating = False
    On Error GoTo Done

    ' seed_chunks: chunk_id, source, page, summary, keywords, full_text
    Dim src As Variant
    src = wsC.Range(wsC.Cells(2, 1), wsC.Cells(lastR, 6)).Value

    Dim n As Long: n = UBound(src, 1)
    Dim outK() As Variant
    ReDim outK(1 To n, 1 To 9)

    Dim vecMap As Object
    Set vecMap = LoadSeedVectors()

    Dim i As Long
    For i = 1 To n
        Dim cid As String: cid = CStr(src(i, 1))
        outK(i, 1) = cid
        outK(i, 2) = CStr(src(i, 2))             ' source(資料名)
        ' origin は "seed" 固定。あとから出所別に数えたり、まとめて外したり
        ' できるようにしておく(個人の資料と混ざったまま消せないのは困る)。
        outK(i, 3) = "seed"
        outK(i, 4) = src(i, 3)                   ' page
        outK(i, 5) = CStr(src(i, 4))             ' summary(第N条(…)等)
        outK(i, 6) = CStr(src(i, 5))             ' keywords
        outK(i, 7) = CStr(src(i, 6))             ' full_text
        outK(i, 8) = modUtil.NowStamp()
        outK(i, 9) = IIf(vecMap.Exists(cid), 1, 0)
    Next i

    Dim baseK As Long
    baseK = wsK.Cells(wsK.Rows.count, 1).End(xlUp).row + 1
    If baseK < 2 Then baseK = 2
    wsK.Range(wsK.Cells(baseK, 1), wsK.Cells(baseK + n - 1, 9)).Value = outK

    ' ベクトルは在るものだけ写す(無ければ embedded=0 のまま次の同期で埋まる)
    Dim wrote As Long
    If Not wsV Is Nothing And vecMap.count > 0 Then
        Dim outV() As Variant
        ReDim outV(1 To vecMap.count, 1 To 2)
        Dim j As Long
        For i = 1 To n
            Dim cid2 As String: cid2 = CStr(src(i, 1))
            If vecMap.Exists(cid2) Then
                j = j + 1
                outV(j, 1) = cid2
                outV(j, 2) = vecMap(cid2)
            End If
        Next i
        If j > 0 Then
            Dim baseV As Long
            baseV = wsV.Cells(wsV.Rows.count, 1).End(xlUp).row + 1
            If baseV < 2 Then baseV = 2
            wsV.Range(wsV.Cells(baseV, 1), wsV.Cells(baseV + j - 1, 2)).Value = outV
            wrote = j
        End If
    End If

    modStats.SetStatText STAT_SEED_DONE, modUtil.NowStamp()
    On Error Resume Next
    modLog.LogUsage "seed_load", "", "chunks=" & n & " vectors=" & wrote
    On Error GoTo Done

    EnsureSeedLoaded = n

Done:
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' SeedDocList - 同梱資料の名前(| 区切り)。案内文はここから組み立てる。
'   実データを読むので、案内が実際の中身とズレることが構造的に起きない。
' ----------------------------------------------------------------------------
Public Function SeedDocList() As String
    On Error Resume Next
    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Worksheets(SH_SEED_CHUNKS)
    If wsC Is Nothing Then Exit Function

    Dim lastR As Long
    lastR = wsC.Cells(wsC.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Function

    Dim arr As Variant
    arr = wsC.Range(wsC.Cells(2, 2), wsC.Cells(lastR, 2)).Value

    Dim seen As String, sb As String
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        Dim nm As String: nm = Trim$(CStr(arr(i, 1)))
        If LenB(nm) > 0 Then
            If InStr(1, "|" & seen & "|", "|" & nm & "|", vbTextCompare) = 0 Then
                seen = seen & IIf(LenB(seen) > 0, "|", "") & nm
                sb = sb & IIf(LenB(sb) > 0, "|", "") & nm
            End If
        End If
    Next i
    SeedDocList = sb
    On Error GoTo 0
End Function

Public Function SeedDocCount() As Long
    Dim l As String: l = SeedDocList()
    If LenB(l) = 0 Then Exit Function
    SeedDocCount = UBound(Split(l, "|")) + 1
End Function

' ----------------------------------------------------------------------------
' SeedQuestions - すぐ押せる質問(| 区切り)。ビルド時に資料から生成済み。
' ----------------------------------------------------------------------------
Public Function SeedQuestions() As String
    On Error Resume Next
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Worksheets(SH_SEED_META)
    If wsM Is Nothing Then Exit Function

    Dim lastR As Long
    lastR = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Function

    Dim arr As Variant
    arr = wsM.Range(wsM.Cells(2, 1), wsM.Cells(lastR, 2)).Value

    Dim sb As String
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        Dim k As String: k = LCase$(Trim$(CStr(arr(i, 1))))
        If Left$(k, 8) = "question" And Right$(k, 4) <> "_src" Then
            Dim q As String: q = Trim$(CStr(arr(i, 2)))
            If LenB(q) > 0 Then sb = sb & IIf(LenB(sb) > 0, "|", "") & q
        End If
    Next i
    SeedQuestions = sb
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' HasSeed - 同梱シードが本棚に入っているか(案内の出し分けに使う)。
' ----------------------------------------------------------------------------
' SeedNeedsEmbedding - 同梱シードにベクトルが無い(=検索に使えない)状態か。
'   ベクトル入りで配布されていれば常にFalse。案内文の出し分けに使う。
'   ここを黙っていると「資料は入っているのに答えられない」という、
'   いちばん信用を失う状態になる。
Public Function SeedNeedsEmbedding() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SEED_VECTORS)
    If ws Is Nothing Then Exit Function
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    SeedNeedsEmbedding = (lastR < 2) And (SeedDocCount() > 0)
    On Error GoTo 0
End Function

Public Function HasSeed() As Boolean
    HasSeed = (LenB(modStats.GetStatText(STAT_SEED_DONE)) > 0) And (SeedDocCount() > 0)
End Function

' ----------------------------------------------------------------------------
' 内部: seed_vectors を chunk_id -> vector_csv の辞書として読む。
' ----------------------------------------------------------------------------
Private Function LoadSeedVectors() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set LoadSeedVectors = d

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SEED_VECTORS)
    If ws Is Nothing Then Exit Function

    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 2)).Value
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        Dim cid As String: cid = Trim$(CStr(arr(i, 1)))
        Dim v As String: v = Trim$(CStr(arr(i, 2)))
        If LenB(cid) > 0 And LenB(v) > 0 Then d(cid) = v
    Next i
    On Error GoTo 0
End Function
