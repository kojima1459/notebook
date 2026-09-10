Attribute VB_Name = "modAskGlobal"
Option Explicit

' ============================================================================
' modAskGlobal - 俯瞰質問(疑似グローバル検索) 2026-08-05 R17 Phase2
' ----------------------------------------------------------------------------
' なぜ要るのか(R17設計書§3 Phase2):
'   「この規程を全部教えて」「全体像は?」「どんな種類がある?」型の質問に、
'   点検索(dense+sparse の上位k件)は原理的に答えられない。質問の語に近い
'   チャンクを並べても、それは【資料のどこか一箇所】であって全体ではない。
'   利用者から見ると「一覧を聞いたのに1件だけ詳しく説明された」という、
'   合っているのに役に立たない外れ方になる。
'
'   そこで取込時に作っておいた章の要約(doc_outline)を全部1回のプロンプトへ
'   載せ、【どの章を読むか】をAIに選ばせる。選ばれた章の本文だけを文書順に
'   集めて、もう1回で回答を作る。1質問あたりの追加呼び出しは2回。
'
' 設計判断:
'   ・フェイルセーフはこの層に閉じる。config graph_outline=off / doc_outline が
'     0行(=まだ取り込み直していない本棚)/ 章が1つも選ばれない /
'     選んだ章のチャンクが1件も引けない / 回答生成が失敗、のどれでも
'     False を返し、呼び出し元(modAskMulti)は従来の入念フローへ落ちる。
'     「俯瞰に失敗したので何も答えません」は作らない(確実に答えが出る側へ
'     倒すのが正しい退化=modRagParse.ParseDecomposeVerdict と同じ思想)。
'   ・選ぶ章は最大4件。1回のプロンプトへ全部の本文を載せるので、章を増やすほど
'     1章あたりの取り分が減り、どの章も途中で切れた抜粋になる(俯瞰したのに
'     何も読めていない状態)。上限の実体は modRagParse.ParseChapterPick が持つ。
'   ・章の本文は【章ごとに】予算を切る(modOutlineBuild.BudgetTake)。先頭の章が
'     長いと後ろの章が丸ごと落ちるのを防ぐため、全体を等分してから積む。
'   ・返す hits の score は 0。俯瞰の材料は検索スコアで選んだものではなく、
'     章の要約から選んだものなので、点数を騙って信頼度バッジ(modAsk.
'     LastConfidence)を水増ししない(modAskFocus の近傍と同じ判断)。
'   ・中断(ESC=Err18)は段の境界で拾う。LLM呼び出しの内側は modGateway が
'     握り潰す既知の制約があり、そこは変えられない。拾ったら【選んだ章の
'     要約だけ】を並べて「※中断」を添えて返す(何も返さないより、どの章に
'     書いてありそうかが分かるほうが次の一手を選べる)。
'   ・プロンプトは modPrompts ではなくここの Private に置く(modPrompts は
'     残386字。R17 Phase2 の容量裁定)。出典タグの書式だけは
'     modPrompts.SourceTag(唯一の持ち主)を必ず通す。
' ============================================================================

' my_knowledge の列(modShelfStore の定義と同じ並び。ここは読むだけ)。
Private Const COL_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_ORIGIN As Long = 3
Private Const COL_PAGE As Long = 4
Private Const COL_FULLTEXT As Long = 7

Private Const PREVIEW_LEN As Long = 120

' 1質問で選べる章数(実体の上限は modRagParse.ParseChapterPick 側)。
Private Const MAX_PICK As Long = 4
' 章の本文として集めるチャンク数の天井(配列再確保とセル読みの回数の安全弁)。
Private Const MAX_HITS As Long = 120
' 1章あたりの本文の最低予算(max_context_chars が小さい設定でも、章の抜粋が
' 数十字しか渡らない状態にはしない)。
Private Const CAP_MIN As Long = 2000

' 章要約が作れなかった章に doc_outline へ入る文字列。modOutlineBuild.FAIL_TEXT と
' 同じ値を対で持つ(Private Const はモジュールを跨げない)。片方を変えるときは
' 必ず両方。中断時の部分回答(AbortDigest)でこの行だけ言い換える(R17H FB-8)。
Private Const OUTLINE_FAIL As String = "(要約失敗)"

' このターンが俯瞰(TryGlobal が回答を作った)だったか(2026-08-05 R17H FA-2)。
' 立てるのは TryGlobal の成功出口だけ、下ろすのは modAskMulti.TryDecomposed の
' 入口だけ。低関連度の警告(modAskRetrieve.ApplyLowHitWarning)と信頼度バッジ
' (modUINexusDraw.DrawConfidence)が「この回答の出自」を知るための1ビットで、
' score=0 のまま【表示だけ】を正しくするために使う(点数は騙さない)。
Private mWasGlobal As Boolean

' ----------------------------------------------------------------------------
' TryGlobal - 俯瞰質問の入口(modAskMulti の段0が verdict=global のときだけ)。
'   True  = この関数が回答を作った(result / ok / hits / nHits を書き換え済み)。
'   False = 何もしていない。呼び出し元は従来の入念フロー(single)を実行する。
'   ok=False で True を返すことは無い(答えを作れなかったときは必ず False で
'   戻し、従来フローに答えさせる)。
' ----------------------------------------------------------------------------
Public Function TryGlobal(ByVal q As String, ByRef hits() As Hit, ByRef nHits As Long, _
                          ByRef ok As Boolean, ByRef result As String) As Boolean
    If Not GateOn() Then Exit Function

    Dim srcs() As String, keys() As String, sums() As String, kws() As String
    Dim n As Long
    n = modOutlineStore.ReadOutline(srcs, keys, sums, kws)
    If Not OutlineActive(n) Then
        LogZero "no_outline"
        Exit Function
    End If

    Dim picks() As String
    Dim nPick As Long
    Dim gHits() As Hit
    Dim gN As Long
    Dim body As String
    Dim lat As Long

    On Error GoTo Bail

    ' 段1: 章の一覧(要約つき)から読むべき章を選ばせる。
    Stage 1, "資料の章立てから関連する章を選定中…"
    Dim resp As String
    resp = modGateway.CallLLM(BuildPickPrompt(q, srcs, keys, sums, n), "chapter_pick", _
        modConfig.GetString("expand_effort", "low"), _
        modConfig.GetString("expand_verbosity", "low"), PickModel(), lat)
    If modRagParse.IsErrorResponse(resp) Then
        LogZero "pick_err"
        Exit Function
    End If
    nPick = modRagParse.ParseChapterPick(resp, MAX_PICK, picks)
    If nPick < 1 Then
        LogZero "no_pick"
        Exit Function
    End If

    ' 段2: 選ばれた章の本文を文書順に集める(LLMは呼ばない)。
    DoEvents                       ' 段の境界。ESC の Err18 はこの窓から入る
    Stage 2, "選んだ" & nPick & "つの章を読み込み中…"
    gN = CollectChapterHits(picks, nPick, gHits)
    If gN < 1 Then
        LogZero "no_chunk"
        Exit Function
    End If

    ' 段3: 章をまたいだ回答を1回で作る。
    DoEvents
    Stage 3, "章をまたいで回答を作成中…"
    Dim ans As String
    ans = modGateway.CallLLM(BuildGlobalPrompt(q, gHits, gN), "global_answer", _
        modConfig.GetString("thorough_draft_effort", "high"), _
        modConfig.GetString("thorough_draft_verbosity", "high"), _
        modConfig.GetString("recommended_model", "gpt-5.5"), lat)
    If modRagParse.IsErrorResponse(ans) Then
        ' ここまでで2回呼んでいるが、答えが出る側へ倒す(従来フローが答える)。
        LogZero "answer_err"
        Exit Function
    End If

    ' 出典突合は分解経路と同じ1本(回答に書かれたタグが材料に在るかの機械照合)。
    body = modAskThorough.AnnotateAgainstHits(modAsk.ApplyAnswerTags(ans), gHits, gN)
    GoTo Done

Bail:
    If Err.Number = 18 Then
        Err.Clear
        Resume Aborted
    End If
    Err.Clear
    Resume Quiet

Quiet:
    LogZero "err"
    Exit Function

Aborted:
    ' 中断。選んだ章が分かっていれば、その要約だけを正直に並べて返す。
    ' 何も選べていないうちの中断は、素の中断として呼び出し元へ返す
    ' (modAskMulti が1論点も終わっていないときに Err.Raise 18 するのと同じ)。
    On Error GoTo 0
    body = AbortDigest(picks, nPick, srcs, keys, sums, n)
    If LenB(body) = 0 Then Err.Raise 18
    If gN > 0 Then
        hits = gHits
        nHits = gN
    End If
    result = body
    ok = True
    TryGlobal = True
    mWasGlobal = True
    LogZero "aborted"
    Exit Function

Done:
    On Error GoTo 0
    On Error Resume Next
    modLog.LogUsage "global_answer", "thorough", _
        "picks=" & nPick & " chunks=" & gN & " outline=" & n, 0, gN
    On Error GoTo 0

    hits = gHits
    nHits = gN
    result = body
    ok = True
    TryGlobal = True
    mWasGlobal = True
End Function

' ----------------------------------------------------------------------------
' WasGlobalTurn / ResetGlobalTurn - 直近ターンが俯瞰だったかの1ビット。
' ----------------------------------------------------------------------------
' 2026-08-05(R17H FA-2 / A-H2・B-H1): 俯瞰の回答は章の要約から章を選んで作る
' ので、返す hits の score は 0(モジュール冒頭の設計判断)。そのままだと
' 「⚠ 手元の資料との関連が薄い」と「🔴 根拠が乏しい」が同時に付き、
' 【章をまたいで正しく答えた回答】が二重に否定されて見える。点数は騙さず、
' 表示だけを出自どおりに直すために、この1ビットを表示側へ渡す。
' リセットは modAskMulti.TryDecomposed の入口(=入念ターンの最初)1箇所だけ。
' そこに置くのは、俯瞰を通らなかった次のターンへ印が残らない唯一の場所だから。
' ----------------------------------------------------------------------------
Public Function WasGlobalTurn() As Boolean
    WasGlobalTurn = mWasGlobal
End Function

Public Sub ResetGlobalTurn()
    mWasGlobal = False
End Sub

' ============================================================================
' 純ロジック(LibreOfficeの実行テストで固定する)
' ============================================================================

' ----------------------------------------------------------------------------
' OutlineActive - 俯瞰が働く条件(doc_outline に1行でも章要約があるか)。
'   【フェイルセーフの単一情報源】。0行(=章要約をまだ作っていない本棚、
'   config graph_outline=off のまま使ってきた本棚、R16以前から使っている
'   本棚)なら False を返し、俯瞰は一切動かず回答は従来と1文字も変わらない。
'   条件をここへ集めるのは、2箇所へ書き写すと片方だけ直したときに
'   「章要約が無いのに章を読もうとする」実行時エラーが出るため
'   (modChunkMeta.GraphActive と同じ考え方)。
' ----------------------------------------------------------------------------
Public Function OutlineActive(ByVal outlineN As Long) As Boolean
    OutlineActive = (outlineN > 0)
End Function

' ============================================================================
' 内部ヘルパー
' ============================================================================

' config graph_outline(既定on)。off のときだけ切る(取込側と同じキー=
' 「章要約を作らない設定なのに読もうとする」状態を作らない)。
Private Function GateOn() As Boolean
    GateOn = (LCase$(Trim$(modConfig.GetString("graph_outline", "on"))) <> "off")
End Function

' 章選択の段のモデル。判定だけの軽い段なので拡張段と同じ選び方。
Private Function PickModel() As String
    PickModel = modConfig.GetString("expand_model", "")
    If LenB(PickModel) = 0 Then PickModel = modConfig.GetString("quick_model", "gpt-5.5")
End Function

' 1回のプロンプトに載せる上限(章一覧・章の本文の両方で使う)。
Private Function SafeCap() As Long
    Dim v As Long: v = 40000
    On Error Resume Next
    v = modConfig.GetLong("max_context_chars", 40000)
    On Error GoTo 0
    If v < CAP_MIN Then v = CAP_MIN
    SafeCap = v
End Function

' 実況(R16-2c と同じ一文つき)。俯瞰は数分ブロックし、その間 Windows が
' 「応答なし」と出すのは正常だと先に言っておく。分母は常に3段。
Private Sub Stage(ByVal idx As Long, ByVal label As String)
    On Error Resume Next
    modUIMain.SetStage ChrW(&HD83D) & ChrW(&HDD2D) & " 俯瞰 " & idx & "/3段: " & _
        label & " ※応答なし表示でも処理中"
    On Error GoTo 0
End Sub

' 俯瞰が不発だったこと・中断したことの記録(R16H FA-1 の教訓: 効いていない
' 状態が無音で続くと、実データでは一度も動いていないことに誰も気付けない)。
'   no_outline = doc_outline が0行(取り込み直していない本棚)
'   pick_err   = 章選択のLLM呼び出しが失敗
'   no_pick    = 該当する章が無い(mock の <pick></pick> もここ)
'   no_chunk   = 選ばれた章のチャンクを1件も引けなかった
'   answer_err = 回答生成のLLM呼び出しが失敗
'   aborted    = 中断されたので章の要約だけを返した
'   err        = 想定外の失敗(従来フローへ落ちる)
Private Sub LogZero(ByVal reason As String)
    On Error Resume Next
    modLog.LogUsage "global_zero", "thorough", "why=" & reason
    On Error GoTo 0
End Sub

' "資料名::章キー" を2つへ割る(":: " が無ければ両方とも空のまま=採らない)。
Private Sub SplitPick(ByVal pick As String, ByRef outSrc As String, ByRef outKey As String)
    outSrc = ""
    outKey = ""
    Dim t As String: t = Trim$(pick)
    Dim p As Long: p = InStr(t, "::")
    If p < 2 Then Exit Sub
    outSrc = Trim$(Left$(t, p - 1))
    outKey = Trim$(Mid$(t, p + 2))
End Sub

' ----------------------------------------------------------------------------
' CollectChapterHits - 選ばれた章のチャンクを文書順に集めて hits を組む。
'   資料(source)と章キーの【両方】が一致した行だけを採る。章キーだけで採ると
'   「第1章 総則」のような、どの規程にもある章名で無関係な資料が混ざる
'   (出典タグは正しいので利用者は気付けない=一番危険な外し方。
'    modAskFocus.RefsExpand が資料を跨がないのと同じ理由)。
'
' 2026-09-06(R38): 1回読み(modAskOnePass)が章ごとに予算を渡すために
'   Public 化・末尾へ Optional capTotal/maxHits を追加した。0 は従来どおり
'   (SafeCap()/MAX_HITS を使う)。既存呼び出し(TryGlobal)は無改修。
' ----------------------------------------------------------------------------
Public Function CollectChapterHits(ByRef picks() As String, ByVal nPick As Long, _
                                    ByRef outHits() As Hit, _
                                    Optional ByVal capTotal As Long = 0, _
                                    Optional ByVal maxHits As Long = 0) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function

    Dim lastK As Long
    lastK = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastK < 2 Then Exit Function

    ' フェイルセーフの判定は共有の1本(R17H FB-1 / A-L11)。chunk_meta が0行の
    ' 本棚では章に割れないので、ここで静かに諦める=回答は R16 までと同じ。
    Dim mIds() As String, mPaths() As String, mRefs() As String
    Dim metaN As Long
    metaN = modChunkMetaStore.ReadAllMeta(mIds, mPaths, mRefs)
    If Not modChunkMeta.GraphActive(metaN, nPick) Then Exit Function

    ' 1) 選ばれた章に属する chunk_id を、章ごとの箱(vbLf区切り)へ集める。
    Dim pSrc() As String: ReDim pSrc(1 To nPick)
    Dim pKey() As String: ReDim pKey(1 To nPick)
    Dim boxes() As String: ReDim boxes(1 To nPick)
    Dim used() As Long: ReDim used(1 To nPick)

    Dim lo As Long: lo = LBound(picks)
    Dim i As Long, j As Long
    For i = 1 To nPick
        SplitPick picks(lo + i - 1), pSrc(i), pKey(i)
    Next i

    ' 2026-08-05(R17H FA-5 / A-M5): 章キーが一致する箱へは【全部】入れる。
    ' 最初の一致で Exit For していた頃は、同名の章キー(「第1章 総則」はどの
    ' 規程にもある)を持つ資料が2本選ばれると、どのチャンクも1つ目の箱にしか
    ' 入らず、2つ目の資料の章が丸ごと空になっていた。資料(source)の突き合わせは
    ' 下の 2) で行番号ごとに行うので、ここで両方の箱へ入れておけば
    ' 「src と 章キーの両方が一致した行だけ採る」が結果として成立する
    ' (chunk_meta には source 列が無く、ここで引くと本棚全行×meta行の走査に
    '  なるため、突合はシートを読む 2) 側へ置いたままにしてある)。
    ' 2026-08-07(R21-3 E2・実機第8報②): 章キーの縮退統合(modOutlineBuild.
    ' GroupChapters/CoalesceSmallChapters)により、pKey(j)は複数の生キーを
    ' CHAPTER_KEY_SEP連結した代表キーのことがある。単純なStrComp完全一致だと
    ' 統合された章のチャンクが1件も引けなくなる(=俯瞰が無音で死ぬ)ため、生キー1件が
    ' 代表キーへ属するかを判定する唯一の照合関数ChapterKeyMatchesを通す
    ' (単独章(非統合)は従来どおりの完全一致に等価。憲章§4-5)。
    ' 2026-08-07(R21H F5): 旧世代doc_outline(OUTLINE_LOGIC_VER<2で保存)の
    ' 章キーはStripTocTail正規化の導入前に作られているため、現行の
    ' ChapterKeyOfで再計算したキー(k)とは一致せず、俯瞰が無音全滅していた。
    ' 正規化前(kRaw)後(k)の2段で照合を試す(⚡仕上げを実行すれば新世代の
    ' キーで保存し直され、この段は自然に不要になる=橋渡しの措置)。
    For i = 0 To metaN - 1
        If LenB(mPaths(i)) > 0 Then
            Dim k As String: k = modOutlineBuild.ChapterKeyOf(mPaths(i))
            Dim kRaw As String: kRaw = modOutlineBuild.ChapterKeyOfRaw(mPaths(i))
            If LenB(k) > 0 Then
                For j = 1 To nPick
                    If LenB(pKey(j)) > 0 Then
                        If modOutlineBuild.ChapterKeyMatches(k, pKey(j)) Or _
                           (kRaw <> k And modOutlineBuild.ChapterKeyMatches(kRaw, pKey(j))) Then
                            boxes(j) = boxes(j) & vbLf & mIds(i)
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    For j = 1 To nPick
        If LenB(boxes(j)) > 0 Then boxes(j) = boxes(j) & vbLf
    Next j

    ' 2) my_knowledge を1回だけ読み(chunk_id と source の2列)、
    '    章の予算内で本文を採る(本文は採用した行だけセル単位で読む)。
    Dim idData As Variant
    idData = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastK, COL_SOURCE)).Value

    Dim capBase As Long
    If capTotal > 0 Then
        capBase = capTotal \ nPick
    Else
        capBase = SafeCap() \ nPick
    End If
    Dim capPer As Long: capPer = capBase
    If capPer < CAP_MIN Then capPer = CAP_MIN

    Dim hitCap As Long
    If maxHits > 0 Then
        hitCap = maxHits
    Else
        hitCap = MAX_HITS
    End If

    Dim n As Long
    For i = LBound(idData, 1) To UBound(idData, 1)
        If n >= hitCap Then Exit For
        Dim id As String: id = Trim$(CStr(idData(i, COL_ID)))
        Dim src As String: src = Trim$(CStr(idData(i, COL_SOURCE)))
        If LenB(id) > 0 And LenB(src) > 0 Then
            For j = 1 To nPick
                If StrComp(src, pSrc(j), vbTextCompare) = 0 Then
                    If InStr(1, boxes(j), vbLf & id & vbLf, vbBinaryCompare) > 0 Then
                        Dim txt As String
                        txt = CStr(ws.Cells(i + 1, COL_FULLTEXT).Value)
                        Dim take As Long
                        take = modOutlineBuild.BudgetTake(used(j), Len(txt), capPer)
                        If take > 0 Then
                            If take < Len(txt) Then txt = modUtil.SafeLeft(txt, take)
                            used(j) = used(j) + take
                            n = n + 1
                            If n = 1 Then
                                ReDim outHits(1 To 1)
                            Else
                                ReDim Preserve outHits(1 To n)
                            End If
                            FillHit ws, i + 1, txt, outHits(n)
                        End If
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
    CollectChapterHits = n
End Function

' 1行を Hit へ組む。score=0(モジュール冒頭の設計判断)。source / page /
' origin は【その行自身の値】を使う(出典タグと原文の整合)。
Private Sub FillHit(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String, _
                    ByRef h As Hit)
    h.chunk_id = CStr(ws.Cells(r, COL_ID).Value)
    h.source = CStr(ws.Cells(r, COL_SOURCE).Value)
    h.origin = CStr(ws.Cells(r, COL_ORIGIN).Value)
    h.page = CLng(Val(CStr(ws.Cells(r, COL_PAGE).Value)))
    h.full_text = txt
    h.preview = modUtil.SafeLeft(txt, PREVIEW_LEN)
    h.score = 0
End Sub

' ----------------------------------------------------------------------------
' AbortDigest - 中断されたときの部分回答(選んだ章の要約を並べたもの)。
'   これは【資料の引用ではなく取込時に作った章の要約】なので、そうと分かる
'   見出しを必ず付ける(憲章§3-3: 分からないことは分からないと言う。
'   要約を原文の引用に見せかけると、条文の数値を要約の言い換えで語ることになる)。
' ----------------------------------------------------------------------------
Private Function AbortDigest(ByRef picks() As String, ByVal nPick As Long, _
                             ByRef srcs() As String, ByRef keys() As String, _
                             ByRef sums() As String, ByVal n As Long) As String
    If nPick < 1 Then Exit Function

    Dim sb As String
    Dim lo As Long: lo = LBound(picks)
    Dim i As Long, j As Long
    For i = 1 To nPick
        Dim ps As String, pk As String
        SplitPick picks(lo + i - 1), ps, pk
        If LenB(pk) > 0 Then
            For j = 0 To n - 1
                If StrComp(Trim$(srcs(j)), ps, vbTextCompare) = 0 Then
                    If StrComp(Trim$(keys(j)), pk, vbTextCompare) = 0 Then
                        If LenB(sb) > 0 Then sb = sb & vbLf & vbLf
                        sb = sb & "■ " & ps & " / " & pk & vbLf & DigestLine(sums(j))
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
    If LenB(sb) = 0 Then Exit Function

    AbortDigest = "(中断されたため、関係しそうな章の【取込時に作った要約】" & _
        "だけを表示しています。原文は確認していません。)" & vbLf & vbLf & sb & vbLf & vbLf & _
        "※中断されたため章の本文は読んでいません。もう一度質問すると最初から調べ直します"
End Function

' 章の要約1行の表示。取込時に要約を作れなかった章(doc_outline に
' OUTLINE_FAIL が入っている)は、内部の失敗語ではなく利用者に分かる言い方へ
' 差し替える(2026-08-05 R17H FB-8 / B-L)。「(要約失敗)」だけが並ぶと、
' 何が壊れたのか・自分は何をすればよいのかが1文字も伝わらない。
Private Function DigestLine(ByVal summary As String) As String
    Dim s As String: s = Trim$(summary)
    If s = OUTLINE_FAIL Then s = "(この章の要約は未完了です)"
    DigestLine = s
End Function

' ----------------------------------------------------------------------------
' BuildPickPrompt - 段1のプロンプト(章の一覧から読むべき章を選ばせる)。
'   一覧の行頭を「資料名::章キー」に固定し、その文字列をそのまま書き写させる。
'   番号で選ばせないのは、番号だと1つずれた瞬間に【まったく別の章】を
'   読み始めるため(ずれても気付けない)。文字列なら、崩れた指定は
'   modRagParse.ParseChapterPick と CollectChapterHits の照合で単に落ちる。
' ----------------------------------------------------------------------------
Private Function BuildPickPrompt(ByVal q As String, ByRef srcs() As String, _
                                 ByRef keys() As String, ByRef sums() As String, _
                                 ByVal n As Long) As String
    Dim sb As String
    sb = "あなたは社内資料の司書です。利用者の質問に答えるために" & _
         "【どの章を読むべきか】を選んでください。" & vbLf
    sb = sb & "・選ぶのは最大" & MAX_PICK & "件。関係の薄い章は選ばないこと(1件でもよい)。" & vbLf
    sb = sb & "・該当しそうな章が1つも無ければ、何も選ばずに空で返すこと" & _
         "(無理に選ぶと関係の無い章の本文で回答が作られる)。" & vbLf
    sb = sb & "・章の指定は一覧の行頭にある【資料名::章キー】を1字も変えずに書き写すこと。" & vbLf
    sb = sb & vbLf & "## 利用者の質問" & vbLf & q & vbLf
    sb = sb & vbLf & "## 章の一覧(資料名::章キー = 取込時に作った章の要約)" & vbLf

    Dim cap As Long: cap = SafeCap()
    Dim used As Long
    Dim i As Long
    For i = 0 To n - 1
        Dim line As String
        line = Trim$(srcs(i)) & "::" & Trim$(keys(i)) & " = " & Trim$(sums(i)) & vbLf
        Dim take As Long
        take = modOutlineBuild.BudgetTake(used, Len(line), cap)
        If take < 1 Then
            sb = sb & "(以下省略)" & vbLf
            Exit For
        End If
        If take < Len(line) Then line = modUtil.SafeLeft(line, take)
        sb = sb & line
        used = used + take
    Next i

    sb = sb & vbLf & "## 出力形式(この形式のみで出力。説明文・前置きは一切禁止)" & vbLf
    sb = sb & "<pick>資料名::章キー|資料名::章キー</pick>" & vbLf
    BuildPickPrompt = sb
End Function

' ----------------------------------------------------------------------------
' BuildGlobalPrompt - 段3のプロンプト(選んだ章の本文で俯瞰的に答えさせる)。
'   出典タグの書式は modPrompts.SourceTag が唯一の持ち主(ここで別の形を
'   書くと、出典突合=modAskThorough.CiteIndexFrom が全件不一致になる)。
' ----------------------------------------------------------------------------
Private Function BuildGlobalPrompt(ByVal q As String, hits() As Hit, _
                                   ByVal nHits As Long) As String
    Dim sb As String
    sb = "あなたは社内資料の調査アシスタントです。利用者は資料の【全体像】を" & _
         "尋ねています。下の本棚抜粋(関連する章の本文)だけを根拠に、" & _
         modConfig.GetString("answer_language", "日本語") & "で答えてください。" & vbLf
    sb = sb & "・全体像が分かる順序で整理する(章ごとに「■ 見出し」を付け、" & _
         "各見出しの下は3行以内の箇条書き)。" & vbLf
    sb = sb & "・抜粋の順番をなぞるだけの写しにしない。共通する原則と、" & _
         "章ごとの違い・例外が分かるように書く。" & vbLf
    sb = sb & "・根拠にした箇所には必ず出典タグを文の直後に付ける" & _
         "(本棚の資料は [本棚:ファイル名 p.ページ番号]、" & _
         "受け取ったパック由来は [パック(作成者名):ファイル名])。" & vbLf
    ' R40 F3 / R41 §1 A / R44: 番地は R43 §4 で【値の直前】へ移った。ここは
    ' modPrompts.CitationInstruction の手書きの複製なので取り残されていた。
    sb = sb & "・出典タグは抜粋の形（Excel はシートN）をそのまま写す。値の直前にある [B42] の" & _
         "ようなセル番地を、根拠の文に「(B42 付近)」のように添える。" & vbLf
    sb = sb & "・抜粋に書かれていないことは書かない。" & _
         "抜粋は資料の【一部の章】なので、答え切れない部分は" & _
         "「この抜粋の範囲では確認できません」と正直に述べる。" & vbLf
    sb = sb & "・Markdown記号(#、**、表)は使わない(この画面では崩れて見える)。" & vbLf
    ' R44: 俯瞰の回答も章の本文から金額・期限・条文番号を引く。Quick と4段には
    ' 最初から入っていた金融・保険のガードレールが、この経路にも
    ' modAskOnePass にも入っていなかった(【数値の厳格性】【(要確認)】
    ' 【断定の禁止】)。文言は modPrompts の単一情報源をそのまま使う。
    sb = sb & modPrompts.DomainGuardInstruction() & vbLf
    If modConfig.GetBool("answer_tags", False) Then
        sb = sb & "・出力は<thinking>に検討、<answer>に利用者へ見せる回答、の構造にすること。" & vbLf
    End If

    sb = sb & vbLf & "## 利用者の質問" & vbLf & q & vbLf
    sb = sb & vbLf & "## 本棚抜粋(選んだ章の本文・文書順)" & vbLf

    Dim i As Long
    For i = 1 To nHits
        sb = sb & modPrompts.SourceTag(hits(i)) & vbLf & hits(i).full_text & vbLf & vbLf
    Next i
    BuildGlobalPrompt = sb
End Function
