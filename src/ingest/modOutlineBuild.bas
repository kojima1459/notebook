Attribute VB_Name = "modOutlineBuild"
Option Explicit

' ============================================================================
' modOutlineBuild - 章単位要約の作成(R17 Phase2・設計書§3 Phase2)
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   検索は「点」しか拾えない。dense+sparse の上位k件は、質問の語に近い
'   チャンクを並べたものであって、「この資料には何が書いてあるか」という
'   問いには原理的に答えられない(俯瞰質問=「〜を全部教えて」「全体像は?」)。
'   規程・マニュアルは既に人の手で【章】という単位に割れているので、
'   GraphRAG のコミュニティ要約の代わりに、その章をそのまま要約単位にする。
'   取込時に章ごと1回だけ CallLLM して doc_outline へ貯め、質問時は
'   章の要約だけを読んで「どの章を読むか」を選ぶ(modAskGlobal)。
'
' 設計判断:
'   ・章キー = section_path の第1要素(ChapterKeyOf)。章見出しの無い資料では
'     これが「第12条(…)」のような条単位になるが、それが正しい保守的動作。
'     見出しの無い資料に勝手な章立てを与えるより、条ごとの要約が並ぶほうが
'     選択段の材料として正直で、外れても「条を1つ余計に読む」だけで済む。
'   ・LLM呼び出しは章数ぶん(254頁の規程で20〜30回=取込+3〜8分)。R15の枠組み
'     (中断ボタン・進捗バナー・ETA・中間保存)にそのまま乗せる。中断されたら
'     【そこまでの章を保存して正常終了】する: 20章のうち12章まで終えた成果を
'     捨てて0章に戻すのは、待った時間の全損になる(R15の一貫した判断)。
'   ・1章の失敗は「(要約失敗)」の行として保存して次の章へ進む。行ごと落とすと、
'     利用者にも次に触る人にも「その章だけ要約が無い」ことが見えない(憲章§3-3)。
'   ・1章の本文は max_context_chars で打ち切る(BudgetTake)。章はいくらでも
'     長くなり得るのに、1回のLLM呼び出しに載る量は決まっている。
'   ・失敗もゲート(config graph_outline)も【この層に閉じる】。呼び出し元
'     (modShelf.IngestFile)は残227字しかなく、判断を1つも持てない(憲章§4-6)。
'   ・プロンプトは modPrompts ではなくここの Private に置く(modPrompts は残386字。
'     章要約プロンプトは1本も入らない=R17 Phase2 の容量裁定)。
' ============================================================================

' my_knowledge の列(modShelfStore の定義と同じ並び。ここは読むだけ)。
Private Const COL_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_FULLTEXT As Long = 7

' 1資料あたりの章数の天井(=LLM呼び出し回数の天井)。254頁の規程で20〜30章。
' 見出しが取れない資料で条単位に割れたときの暴走(数百回の呼び出し)を止める。
Private Const MAX_CHAPTERS As Long = 60

' 1章の本文へ載せる最低限(max_context_chars が異常に小さい設定でも、
' 要約の材料が数十字しか渡らない状態にはしない)。
Private Const CAP_MIN As Long = 2000

Private Const SUMMARY_MAX As Long = 600
Private Const KEYWORDS_MAX As Long = 200
Private Const FAIL_TEXT As String = "(要約失敗)"

' 無人実行(自動同期)中は章要約を作らない資料の章数(2026-08-05 R17H FA-7)。
Private Const UNATTENDED_MAX_CH As Long = 12

' いま走っているのが「誰も見ていない処理」か(2026-08-05 R17H FA-7 / A-M8)。
' 自動同期(modShelfSync.SyncNow silent:=True)は OnTime で勝手に始まり、
' 進捗バナーも中断ボタンも利用者の目の前には無い。そこで章数の多い資料の
' 章要約(=章数ぶんのLLM呼び出し。254頁で20〜30回・数分〜十数分)を始めると、
' 「何もしていないのにExcelが数分固まる」だけの体験になる。手動取込・手動
' 同期(silent=False)では従来どおり全量作る。印の上げ下げは modShelfSync の
' 入口/Finish の2行だけで、こちらからは読むだけ。
Private mUnattended As Boolean

' ----------------------------------------------------------------------------
' SetUnattended - 無人実行(自動同期)の印を上げ下げする(R17H FA-7)。
'   呼ぶのは modShelfSync.SyncNow の入口(silent の値)と Finish(False)だけ。
' ----------------------------------------------------------------------------
Public Sub SetUnattended(ByVal b As Boolean)
    mUnattended = b
End Sub

' ----------------------------------------------------------------------------
' BuildOutlineFor - 1資料ぶんの章要約を作って doc_outline へ保存する。
'   config graph_outline=off / chunk_meta が0行 / その資料のチャンクが無い /
'   章キーが1つも取れない のいずれかなら【完全に無操作】(=LLMも呼ばない)。
'   呼び出しは modShelf.IngestFile の Finish(done/partial)から1行だけ。
' ----------------------------------------------------------------------------
Public Sub BuildOutlineFor(ByVal sourceName As String)
    If Not GateOn() Then Exit Sub
    If LenB(Trim$(sourceName)) = 0 Then Exit Sub

    On Error GoTo Quiet

    ' 1) 構造メタが無ければ章に割れない(=まだ取り込み直していない本棚)。
    Dim mIds() As String, mPaths() As String, mRefs() As String
    Dim metaN As Long
    metaN = modChunkMetaStore.ReadAllMeta(mIds, mPaths, mRefs)
    If metaN < 1 Then Exit Sub

    ' 2) その資料のチャンクを文書順(=シートの並び順)に集める。
    Dim rowIdx() As Long, ids() As String
    Dim nSrc As Long
    nSrc = CollectSourceRows(sourceName, rowIdx, ids)
    If nSrc < 1 Then Exit Sub

    ' 3) chunk_id → section_path を引き当てる(chunk_meta 側を1周するだけ)。
    Dim paths() As String: ReDim paths(1 To nSrc)
    MapPaths ids, nSrc, mIds, mPaths, metaN, paths

    ' 4) 章キーで束ねる(出現順=文書順)。
    Dim chKeys() As String, chCount() As Long
    Dim nCh As Long
    nCh = GroupChapters(paths, nSrc, chKeys, chCount)
    If nCh < 1 Then Exit Sub

    ' 無人実行(自動同期)で章数が多い資料は、章要約を先送りする(R17H FA-7)。
    ' 手動取込・手動同期(非silent)なら同じ資料でも全量作られる=先送りは
    ' 「作らない」ではなく「人が見ているときに作る」。名寄せ(name_dedup)は
    ' 1資料につき最大1回なので、無人でもそのまま続ける。
    If mUnattended And nCh > UNATTENDED_MAX_CH Then
        On Error Resume Next
        modLog.LogUsage "outline_deferred", "ingest", _
            "source=" & sourceName & " chapters=" & nCh & " (unattended)"
        On Error GoTo 0
        GoTo SynStep
    End If

    Dim cap As Long: cap = SafeCap()
    Dim sums() As String: ReDim sums(1 To nCh)
    Dim kws() As String: ReDim kws(1 To nCh)
    Dim srcs() As String: ReDim srcs(1 To nCh)

    Dim eff As String: eff = modConfig.GetString("quick_effort", "low")
    Dim vrb As String: vrb = modConfig.GetString("quick_verbosity", "low")
    Dim mdl As String: mdl = modConfig.GetString("quick_model", "gpt-5.5")

    Dim msPerItem As Double
    Dim doneN As Long
    Dim i As Long

    On Error GoTo ChapFail
    For i = 1 To nCh
        ' 中断は章の境界で拾う(LLM呼び出しの内側では modGateway が握り潰す)。
        If modShelfBatch.CancelRequested() Then Exit For
        Banner i, nCh, sourceName, msPerItem

        Dim t0 As Double: t0 = Timer
        srcs(i) = sourceName
        sums(i) = FAIL_TEXT      ' 先に失敗を書いておく=どこで落ちても正直に残る
        kws(i) = ""

        Dim body As String
        body = ChapterBody(rowIdx, paths, nSrc, chKeys(i), cap)
        If LenB(body) > 0 Then
            Dim lat As Long
            Dim resp As String
            resp = modGateway.CallLLM(BuildChapterPrompt(sourceName, chKeys(i), body), _
                                      "chapter_summary", eff, vrb, mdl, lat)
            Dim sm As String, kw As String
            If modRagParse.ParseOutlineResp(resp, sm, kw) Then
                sums(i) = modUtil.SafeLeft(sm, SUMMARY_MAX)
                kws(i) = modUtil.SafeLeft(kw, KEYWORDS_MAX)
            End If
        End If

        doneN = i
        msPerItem = modUtilText.BlendPerItemMs(msPerItem, t0, 1)
        modShelfBatch.SaveCheckpoint 1, 120
    Next i
    GoTo LoopDone

ChapFail:
    If Err.Number = 18 Then
        ' ESC。ここまでの章を保存して正常終了する(待った時間を捨てない)。
        Err.Clear
        Resume LoopDone
    End If
    ' 想定外の失敗はその章だけ「(要約失敗)」のまま次へ(1章の事故で全部を失わない)。
    Err.Clear
    Resume Next

LoopDone:
    On Error GoTo Quiet
    If doneN < 1 Then Exit Sub

    ' 作り直し: 同じ資料の古い章要約を落としてから今回ぶんを1回で書く
    ' (取込フック側でも掃除しているが、ここが唯一の書込み口なので二重に守る。
    '  同じ資料の行が2組並ぶと、章選択の段で同じ章が2回候補に出る)。
    modOutlineStore.RemoveOutlineForSource sourceName
    modOutlineStore.WriteOutlineRows srcs, chKeys, sums, kws, chCount, doneN

    On Error Resume Next
    modUIMain.SetStage ""
    modLog.LogUsage "outline_built", "ingest", _
        "source=" & sourceName & " chapters=" & doneN & "/" & nCh
    On Error GoTo 0

SynStep:
    ' R17 Phase3: 章要約に続けて名寄せ辞書のバッチ生成を1回だけ(ゲート・失敗握り
    ' ・usage_log("synonyms_fail")はmodSynonymStore側に閉じる。ここは1行)。
    On Error GoTo Quiet
    If Not modShelfBatch.CancelRequested() Then
        On Error Resume Next
        modSynonymStore.BuildSynonymsFor sourceName
        On Error GoTo 0
    End If
    Exit Sub

Quiet:
    ' ハンドラ稼働中は On Error Resume Next が効かないので、必ず Resume で
    ' 抜けてから記録する(2026-07-30 実機err#462 と同型の作法)。
    Resume QuietDone
QuietDone:
    Err.Clear
    On Error Resume Next
    modLog.LogUsage "outline_fail", "ingest", "build source=" & sourceName
    On Error GoTo 0
End Sub

' ============================================================================
' 純ロジック(LibreOfficeの実行テストで固定する)
' ============================================================================

' ----------------------------------------------------------------------------
' ChapterKeyOf - section_path("第3章 総則>第12条(免責)")から章キーを取る。
'   戻り値 = 第1要素をそのまま("第3章 総則")。正規化はしない: section_path は
'   取込時に modSparse.NormalizeForSearch を通っており、ここで別の式を掛けると
'   保存側(doc_outline)と照合側(modAskGlobal)で章キーが割れる(憲章§4-5)。
'   ・章見出しが無い資料では第1要素が条になる(=条単位の要約)。それが正しい
'     保守的動作で、無い章立てを推測するよりも外れ方が小さい。
'   ・空/">"だけ/先頭が空の path は空文字("章が分からない"という正しい答え。
'     呼び出し元はその行を束ねの対象から外す)。
' ----------------------------------------------------------------------------
Public Function ChapterKeyOf(ByVal sectionPath As String) As String
    Dim t As String: t = Trim$(sectionPath)
    If LenB(t) = 0 Then Exit Function
    Dim p As Long: p = InStr(t, ">")
    If p > 0 Then t = Left$(t, p - 1)
    ChapterKeyOf = Trim$(t)
End Function

' ----------------------------------------------------------------------------
' BudgetTake - 「あと何字入れられるか」の算数(1章の本文の打ち切り)。
'   usedLen: すでに積んだ字数 / addLen: 次に足したい字数 / cap: 上限
'   戻り値 0 = もう1字も入らない(呼び出し元はそこで打ち切る)。
'   cap を超えた分を黙って捨てるのではなく【手前で切る】ことにしているのは、
'   LLM側で切られると章の後半が理由も分からず消えるため(こちらで切れば
'   「どこまで読んだか」が呼び出し元の手に残る)。
' ----------------------------------------------------------------------------
Public Function BudgetTake(ByVal usedLen As Long, ByVal addLen As Long, _
                           ByVal cap As Long) As Long
    If cap <= 0 Or addLen <= 0 Then Exit Function
    Dim room As Long: room = cap - usedLen
    If room <= 0 Then Exit Function
    If addLen < room Then
        BudgetTake = addLen
    Else
        BudgetTake = room
    End If
End Function

' ============================================================================
' 内部ヘルパー
' ============================================================================

' config graph_outline(既定on)。off のときだけ切る(graph_refs と同じ作法)。
Private Function GateOn() As Boolean
    GateOn = (LCase$(Trim$(modConfig.GetString("graph_outline", "on"))) <> "off")
End Function

' 1章の本文へ載せる上限。LLM呼び出し1回ぶんの上限と同じ値を使う。
Private Function SafeCap() As Long
    Dim v As Long: v = 40000
    On Error Resume Next
    v = modConfig.GetLong("max_context_chars", 40000)
    On Error GoTo 0
    If v < CAP_MIN Then v = CAP_MIN
    SafeCap = v
End Function

' その資料の行(シート順=文書順)の行番号と chunk_id を集める。
' 読むのは chunk_id 列と source 列の2列だけ(本文まで全部読むと20,000件規模で
' 数百MBになる。modAskFocus と同じ作法)。
Private Function CollectSourceRows(ByVal sourceName As String, ByRef outRows() As Long, _
                                   ByRef outIds() As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function

    Dim lastK As Long
    lastK = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastK, COL_SOURCE)).Value

    ReDim outRows(1 To lastK - 1)
    ReDim outIds(1 To lastK - 1)
    Dim n As Long
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(Trim$(CStr(arr(i, COL_SOURCE))), sourceName, vbTextCompare) = 0 Then
            Dim id As String: id = Trim$(CStr(arr(i, COL_ID)))
            If LenB(id) > 0 Then
                n = n + 1
                outRows(n) = i + 1          ' arr は2行目起点
                outIds(n) = id
            End If
        End If
    Next i
    CollectSourceRows = n
End Function

' chunk_id → section_path の引き当て。chunk_meta 側を1周し、その資料の
' chunk_id の箱に入っているものだけ位置を割り出して書き込む(箱は数十KB、
' meta は本棚全体ぶん。modAskFocus.RefsExpand と同じ「箱へのInStr」流儀)。
Private Sub MapPaths(ByRef ids() As String, ByVal nSrc As Long, _
                     ByRef mIds() As String, ByRef mPaths() As String, _
                     ByVal metaN As Long, ByRef outPaths() As String)
    Dim box As String
    Dim i As Long
    For i = 1 To nSrc
        box = box & vbLf & ids(i)
    Next i
    box = box & vbLf

    For i = 0 To metaN - 1
        If LenB(mPaths(i)) > 0 Then
            Dim k As Long: k = IndexInBox(box, mIds(i))
            If k >= 1 And k <= nSrc Then outPaths(k) = mPaths(i)
        End If
    Next i
End Sub

' vbLf区切りの箱の中で id が何番目か(1始まり。無ければ0)。
' 位置は「一致した所より前にある vbLf の数」で決まる(箱は必ず vbLf で始まる)。
Private Function IndexInBox(ByVal box As String, ByVal id As String) As Long
    If LenB(id) = 0 Then Exit Function
    Dim p As Long
    p = InStr(1, box, vbLf & id & vbLf, vbBinaryCompare)
    If p = 0 Then Exit Function
    Dim head As String: head = Left$(box, p)
    IndexInBox = Len(head) - Len(Replace(head, vbLf, ""))
End Function

' 章キーで束ねる(出現順)。戻り値=章数。outCount にはその章のチャンク数。
Private Function GroupChapters(ByRef paths() As String, ByVal nSrc As Long, _
                               ByRef outKeys() As String, ByRef outCount() As Long) As Long
    ReDim outKeys(1 To nSrc)
    ReDim outCount(1 To nSrc)

    Dim n As Long
    Dim i As Long, j As Long
    For i = 1 To nSrc
        Dim k As String: k = ChapterKeyOf(paths(i))
        If LenB(k) > 0 Then
            Dim at As Long: at = 0
            For j = 1 To n
                If StrComp(outKeys(j), k, vbBinaryCompare) = 0 Then
                    at = j
                    Exit For
                End If
            Next j
            If at > 0 Then
                outCount(at) = outCount(at) + 1
            ElseIf n < MAX_CHAPTERS Then
                n = n + 1
                outKeys(n) = k
                outCount(n) = 1
            End If
        End If
    Next i
    GroupChapters = n
End Function

' 1章ぶんの本文(文書順に連結し cap で打ち切る)。本文は採用した行だけ
' セル単位で読む(全列を読み直すと本棚全体の本文をもう一度メモリへ載せる)。
Private Function ChapterBody(ByRef rowIdx() As Long, ByRef paths() As String, _
                             ByVal nSrc As Long, ByVal chKey As String, _
                             ByVal cap As Long) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function

    Dim sb As String
    Dim used As Long
    Dim i As Long
    For i = 1 To nSrc
        If StrComp(ChapterKeyOf(paths(i)), chKey, vbBinaryCompare) = 0 Then
            Dim t As String
            t = CStr(ws.Cells(rowIdx(i), COL_FULLTEXT).Value)
            Dim take As Long: take = BudgetTake(used, Len(t), cap)
            If take < 1 Then Exit For
            If take < Len(t) Then t = modUtil.SafeLeft(t, take)
            sb = sb & t & vbLf
            used = used + take
        End If
    Next i
    ChapterBody = sb
End Function

' 進捗(バナー+実況行)。ETAは1章あたりの移動平均から出す(modEnrich と同じ
' 部品 modUtilText.BlendPerItemMs / modUtil.EtaText を使う)。
Private Sub Banner(ByVal idx As Long, ByVal total As Long, ByVal sourceName As String, _
                   ByVal msPerItem As Double)
    On Error Resume Next
    Dim t As String
    t = "章の要約中… " & idx & "/" & total & "章(" & modUtil.SafeLeft(sourceName, 40) & ")"
    Dim eta As String: eta = modUtil.EtaText(total - idx + 1, msPerItem)
    If LenB(eta) > 0 Then t = t & " " & eta
    modShelfBatch.StageBanner t
    modUIMain.SetStage t
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' BuildChapterPrompt - 章要約のプロンプト(Private。置き場の理由は冒頭)。
'   出力契約はタグ形式(modRagParse.ParseOutlineResp が読む):
'     <summary>…</summary><keywords>a|b|c</keywords>
'   要約に「具体名を残す」と指示するのは、この要約が【章を選ぶための材料】
'   だから。「本章は手続について定める」とだけ書かれた要約が20本並ぶと、
'   どの章を読むべきかを選ぶ手掛かりが1つも無くなる。
' ----------------------------------------------------------------------------
Private Function BuildChapterPrompt(ByVal sourceName As String, ByVal chKey As String, _
                                    ByVal body As String) As String
    Dim sb As String
    sb = "あなたは社内資料の要約担当です。次に示すのは資料「" & sourceName & "」の" & _
         "「" & chKey & "」の本文です。後で【どの章を読むべきか】を選ぶための" & _
         "材料として、この範囲に何が書いてあるかを要約してください。" & vbLf
    sb = sb & "・要約は200～300字の日本語。対象・原則・例外の順に、具体名" & _
         "(条番号・書類名・期限・金額)を必ず残すこと。" & vbLf
    sb = sb & "・本文に書かれていないことは書かない(推測・一般論・前置き・感想は禁止)。" & vbLf
    sb = sb & "・キーワードは、この範囲を探すときに使われそうな語を最大8個。" & vbLf
    sb = sb & vbLf & "## 本文" & vbLf & body & vbLf & vbLf
    sb = sb & "## 出力形式(この形式のみで出力。説明文・前置きは一切禁止)" & vbLf
    sb = sb & "<summary>200～300字の要約</summary>" & vbLf
    sb = sb & "<keywords>語1|語2|語3</keywords>" & vbLf
    BuildChapterPrompt = sb
End Function
