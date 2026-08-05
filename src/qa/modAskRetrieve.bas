Attribute VB_Name = "modAskRetrieve"
Option Explicit

' ========================================
' modAskRetrieve - modAsk から切り出した「検索して材料を揃える」層
'
' 多段RAG(拡張→マルチクエリ→再ランク)と、ヒット結果の評価(低関連度の警告・
' 曖昧すぎる質問の判定・資料名の列挙)。いずれも modAsk のモジュール変数
' (mLastHits 等)を一切触らず、引数の hits() だけで完結する。
' そのため「答えが変」の原因が検索側か生成側かを、ここだけ読めば切り分けられる。
'
' 切り出しの理由(2026-07-28): modAsk が契約上限30,000字に対し残り88字で、
' 出典アクセサの添字修正(レビュー C-2)すら入らなかった(レビュー I-2)。
' ========================================

' 段階ナレーション(R13-9b)の計画。算数は modMode(純ロジック)が持つ。
Private mStgExpand As Boolean
Private mStgRerank As Boolean
Private mStgTotal As Long
' R14-8a: 入念モードは要点整理・自己批判の2段が増えて最大6段になる。
Private mStgThorough As Boolean

' 多段RAG(§C): 拡張→マルチクエリ→再ランク。失敗時は単段Searchへ退化。
' scopeSources(R13-5a/5c): 許可資料名のDictionary。Nothing=従来どおり本棚全体。
' subqOverride(R13-5c): >0 なら拡張のサブクエリ本数をこの値に固定し、
'   config/モードに関係なく拡張段を必ず通す(深掘りのスコープ内多段検索用)。
' skipExpand(R16-3A/裁定1で意味を拡張): True なら【拡張段と再ランク段の両方を
'   通さない】=副質問向けの軽量検索。複合質問を論点へ分解した後の検索で使う。
'   拡張を切る理由: 既に1論点まで割ってあるものを更にばらすと、論点の外の資料が
'   混ざって「その論点だけを詰める」という分解の目的が消える。
'   再ランクも切る理由: 論点あたりのtopKは6～8と小さく、候補プールとの差が
'   ほとんど無い(並べ替える余地が無い)。にもかかわらず論点数ぶんLLM呼び出しが
'   増え、3論点なら合計が6+3=9回ではなく12回になる。効かない段に待ち時間だけ
'   払う形なので、副質問では検索1本に絞る。
'   既定Falseで従来の呼び出しは1つも挙動が変わらない。subqOverrideとは併用しない。
Public Function RunMultiRetrieve(ByVal q As String, ByVal mdMode As String, _
                                  ByVal topK As Long, ByRef hits() As Hit, _
                                  Optional ByVal scopeSources As Object, _
                                  Optional ByVal subqOverride As Long = 0, _
                                  Optional ByVal skipExpand As Boolean = False) As Long
    On Error GoTo FallbackSingle

    ' 1) クエリ拡張
    Dim queries() As String
    Dim queryN As Long: queryN = 0
    ReDim queries(0 To 8)

    Dim standalone As String, hyde As String
    Dim subs() As String
    standalone = "": hyde = ""

    ' 各段を通すかはモードで決まる(modMode)。
    Dim useExpand As Boolean
    useExpand = modMode.UseExpand(mdMode, modConfig.GetBool("expand_enabled", False), _
                                  modConfig.GetBool("quick_expand", False))
    If subqOverride > 0 Then useExpand = True     ' R13-5c: スコープ内多段は必ず角度を作る
    If skipExpand Then useExpand = False           ' R16-3A: 分解済みの論点は更にばらさない

    If useExpand Then
        ShowAskStage "expand"
        Dim lightMode As Boolean
        lightMode = modMode.UseLightExpand(mdMode, modConfig.GetBool("quick_expand_light", True))

        Dim subN As Long
        subN = subqOverride
        If subN < 1 Then
            subN = modMode.SubQueryCount(mdMode, modConfig.GetLong("expand_subqueries", 3), _
                                         modConfig.GetLong("thorough_subqueries", 6))
        End If

        Dim exPrompt As String
        exPrompt = modPrompts.BuildExpandPrompt(q, modAsk.HistoryBlock(), subN, lightMode)

        Dim exModel As String: exModel = modConfig.GetString("expand_model", "")
        If LenB(exModel) = 0 Then exModel = modConfig.GetString("quick_model", "gpt-5.5")
        Dim exLat As Long
        Dim exResp As String
        exResp = modGateway.CallLLM(exPrompt, "expand", _
            modConfig.GetString("expand_effort", "low"), _
            modConfig.GetString("expand_verbosity", "low"), exModel, exLat)

        If Not modRagParse.IsErrorResponse(exResp) Then
            modRagParse.ParseExpand exResp, standalone, subs, hyde
        End If
    End If

    If LenB(Trim$(standalone)) = 0 Then standalone = q
    queries(queryN) = standalone: queryN = queryN + 1

    Dim si As Long
    On Error Resume Next
    For si = LBound(subs) To UBound(subs)
        If queryN <= 7 And LenB(Trim$(subs(si))) > 0 Then
            queries(queryN) = subs(si): queryN = queryN + 1
        End If
    Next si
    On Error GoTo FallbackSingle
    If LenB(Trim$(hyde)) > 0 And queryN <= 8 Then
        queries(queryN) = hyde: queryN = queryN + 1
    End If
    ReDim Preserve queries(0 To queryN - 1)

    ' 2) マルチクエリ検索(候補プール)
    Dim poolK As Long: poolK = modConfig.GetLong("multi_candidates", 40)
    If poolK < topK Then poolK = topK
    Dim poolHits() As Hit
    Dim poolN As Long
    poolN = modRetrieve.SearchExpanded(queries, poolK, poolHits, scopeSources)
    If poolN <= 0 Then GoTo FallbackSingle

    ' 3) 再ランク(候補がtopKより多いときだけ意味がある)
    Dim orderN As Long: orderN = 0
    Dim rankOrder() As Long
    Dim useRerank As Boolean
    useRerank = modMode.UseRerank(mdMode, modConfig.GetBool("rerank_enabled", False), _
                                  modConfig.GetBool("quick_rerank", False))
    If skipExpand Then useRerank = False           ' 裁定1: 副質問は軽量検索(再ランクも省く)
    If useRerank And poolN > topK Then
        ShowAskStage "rerank"
        Dim rkPrompt As String
        rkPrompt = modPrompts.BuildRerankPrompt(q, poolHits, poolN, _
            modConfig.GetLong("max_context_chars", 40000))
        Dim rkModel As String: rkModel = modConfig.GetString("rerank_model", "")
        If LenB(rkModel) = 0 Then rkModel = modConfig.GetString("quick_model", "gpt-5.5")
        Dim rkLat As Long
        Dim rkResp As String
        ' R14-8a: 入念モードだけ再ランクの effort を上げる(rerank_effort_thorough)。
        ' どの資料を根拠にするかを決める段なので、ここが雑だと後段の検証では直らない。
        rkResp = modGateway.CallLLM(rkPrompt, "rerank", _
            modMode.RerankEffort(mdMode, modConfig.GetString("rerank_effort", "low"), _
                                 modConfig.GetString("rerank_effort_thorough", "medium")), _
            modConfig.GetString("rerank_verbosity", "low"), rkModel, rkLat)
        If Not modRagParse.IsErrorResponse(rkResp) Then
            orderN = modRagParse.ParseRankOrder(rkResp, poolN, rankOrder)
        End If
    End If

    ' 4) 最終topKへ絞り込み(再ランク順 or スコア順)
    Dim outN As Long
    If orderN > 0 Then
        outN = orderN
    Else
        outN = poolN
    End If
    If outN > topK Then outN = topK

    ReDim hits(1 To outN)
    Dim oi As Long
    For oi = 1 To outN
        If orderN > 0 Then
            hits(oi) = poolHits(rankOrder(oi - 1))
        Else
            hits(oi) = poolHits(oi)
        End If
    Next oi
    RunMultiRetrieve = outN
    Exit Function

FallbackSingle:
    Err.Clear
    On Error GoTo 0
    RunMultiRetrieve = modRetrieve.Search(q, topK, hits, scopeSources)
End Function

' ----------------------------------------------------------------------------
' RunDeepScoped - 深掘り(followup かつ deep)の検索(2026-08-03 R13-5c)
' ----------------------------------------------------------------------------
' 「深掘り=会話の流れの中を深く掘る」を実装する唯一の場所。
'   (i)   会話出典メモリ(modFollowup)があればそれをスコープにする
'   (ii)  無ければ通常検索を1回だけ行い、上位ヒットの資料名をスコープにする
'   (iii) そのスコープの中で多段検索(サブクエリ deep_scope_subqueries 本)
'   (iv)  スコープ内のヒットが2件未満なら、無スコープで取り直す
'         (今日より悪くなる経路を作らない。落ちたことは usage_log に残す)
' 新規質問(非followup)の deep はこの関数を通らない=現行動作のまま。
Public Function RunDeepScoped(ByVal q As String, ByVal mdMode As String, _
                              ByVal topK As Long, ByRef hits() As Hit) As Long
    Dim scopeD As Object
    Set scopeD = modFollowup.CitedSourcesDict()

    If scopeD Is Nothing Then
        ' (ii) 会話出典がまだ無い(State Loss後など)。1回だけ通常検索して
        '      「いまの話題の資料」を自分で決める。ここで失敗したら
        '      スコープ無しへ静かに退化する(質問は必ず答える)。
        Dim seed() As Hit
        Dim seedN As Long
        seedN = modRetrieve.Search(q, topK, seed)
        If seedN = -1 Then
            RunDeepScoped = -1      ' 埋め込み失敗(E0203)。もう一度呼んでも同じなので繰り返さない
            Exit Function
        End If
        If seedN > 0 Then Set scopeD = modFollowup.ScopeDictFrom(HitSourceList(seed, seedN))
    End If

    If Not scopeD Is Nothing Then
        ' (iii) スコープ内で多段検索。拡張段は必ず通す(角度を作らないと
        '       「狭い中を深く」にならず、ただの絞り込みで終わる)。
        mStgExpand = True
        mStgTotal = modMode.AskStageTotal(mStgExpand, mStgRerank, True, mStgThorough)

        Dim subN As Long
        subN = modConfig.GetLong("deep_scope_subqueries", 6)
        If subN < 1 Then subN = 6

        Dim n As Long
        n = RunMultiRetrieve(q, mdMode, WideK(topK), hits, scopeD, subN)
        If n >= 2 Then
            On Error Resume Next
            modLog.LogUsage "deep_scoped", mdMode, "scope=" & scopeD.count & " subq=" & subN, 0, n
            On Error GoTo 0
            RunDeepScoped = FinishDeep(hits, n, topK)
            Exit Function
        End If
        If n = -1 Then
            RunDeepScoped = -1      ' 埋め込み失敗。広げ直しても同じ結果にしかならない
            Exit Function
        End If

        On Error Resume Next
        modLog.LogUsage "deep_scope_fallback", mdMode, _
            "会話の資料の中では" & n & "件しか見つからず、本棚全体へ広げ直しました(scope=" & scopeD.count & ")"
        On Error GoTo 0
        PlanAskStages mdMode          ' 番号計画を通常構成へ戻す
    End If

    ' (iv)/退化: 従来どおり本棚全体を検索する。
    Dim wideN As Long
    wideN = RunUnscoped(q, mdMode, WideK(topK), hits)
    If wideN < 1 Then
        RunDeepScoped = wideN          ' 0件・埋め込み失敗(-1)はそのまま返す
        Exit Function
    End If
    RunDeepScoped = FinishDeep(hits, wideN, topK)
End Function

' ----------------------------------------------------------------------------
' FinishDeep - 深掘りで材料が確定したあとの後処理(2026-08-05 R16-3C)。
'   戻り値 = 後処理後の件数。スコープ内で見つかった経路と、本棚全体へ広げ
'   直した経路の両方が必ずここを通る(片方だけに掛けると、同じ「続けて質問」が
'   本棚の当たり方で別の読み方をされる)。
'
'   精読(近傍チャンク束ね)は深掘りにも効かせる。深掘りは同じ資料の中を
'   もう一段深く読む操作なので、当たったチャンクの前後こそが一番読みたい場所
'   になる(表の続き・条文の但し書き)。
' ----------------------------------------------------------------------------
'   R16-3D: 検索は topK の2倍で取り、既出チャンク(前回までの深掘りで実際に
'   渡したもの)を後ろへ回してから topK へ切る。除外ではなく降格なので、
'   本棚が小さくて既出しか無いときも件数は減らない(modFollowup.DemoteUsed)。
'   順序: 降格 → 記憶 → 精読。記憶を降格の前に置くと、いま返すチャンクを
'   自分で既出扱いして全部後ろへ回す(1ターンで自滅する)。記憶を精読の後に
'   置くと、近傍(おまけ)まで既出になって次の深掘りで本命が沈む。
Private Function FinishDeep(ByRef hits() As Hit, ByVal n As Long, ByVal topK As Long) As Long
    Dim m As Long: m = n
    modFollowup.DemoteUsed hits, m, topK
    modFollowup.RememberUsedChunks hits, m
    modAskFocus.NeighborExpand hits, m, modConfig.GetLong("deep_neighbor", 2)
    FinishDeep = m
End Function

' 深掘りの検索で取る件数(降格ぶんの余裕を持たせた広めの取り方)。
' 2倍にするのは、topK 件すべてが既出でも「未出の候補が topK 件ある」状態を
' 作れる最小の倍率だから(多く取るほど再ランク段の入力も伸びて遅くなる)。
Private Function WideK(ByVal topK As Long) As Long
    WideK = topK * 2
    If WideK < 2 Then WideK = 2
End Function

' 従来の検索経路(retrieve_mode に従う)。modAsk と同じ判断を2箇所に書かない。
Public Function RunUnscoped(ByVal q As String, ByVal mdMode As String, _
                            ByVal topK As Long, ByRef hits() As Hit) As Long
    If LCase$(modConfig.GetString("retrieve_mode", "single")) = "multi" Then
        RunUnscoped = RunMultiRetrieve(q, mdMode, topK, hits)
    Else
        RunUnscoped = modRetrieve.Search(q, topK, hits)
    End If
End Function

' ----------------------------------------------------------------------------
' 段階ナレーション(R13-9b): この質問で何段通すかを先に決めてから実況する。
' ----------------------------------------------------------------------------
Public Sub PlanAskStages(ByVal mdMode As String)
    On Error Resume Next
    Dim isMulti As Boolean
    isMulti = (LCase$(modConfig.GetString("retrieve_mode", "single")) = "multi")
    mStgExpand = isMulti And modMode.UseExpand(mdMode, modConfig.GetBool("expand_enabled", False), _
                                               modConfig.GetBool("quick_expand", False))
    mStgRerank = isMulti And modMode.UseRerank(mdMode, modConfig.GetBool("rerank_enabled", False), _
                                               modConfig.GetBool("quick_rerank", False))
    mStgThorough = (modMode.Normalize(mdMode) = "thorough")
    mStgTotal = modMode.AskStageTotal(mStgExpand, mStgRerank, modMode.UseVerify(mdMode), mStgThorough)
    On Error GoTo 0
End Sub

' kind = "expand" / "rerank" / "digest" / "draft" / "critique" / "verify" / "quick"
Public Sub ShowAskStage(ByVal kind As String)
    On Error Resume Next
    modUIMain.SetStage modMode.AskStageText( _
        modMode.AskStageIndex(kind, mStgExpand, mStgRerank, mStgThorough), _
        mStgTotal, modMode.AskStageLabel(kind))
    On Error GoTo 0
End Sub

' 低関連度警告(表示専用): 最高スコアが config low_hit_warn_score(既定0.3)
' 未満なら注意書きを先頭に付ける。0以下は無効化扱い。
Public Function ApplyLowHitWarning(ByVal result As String, hits() As Hit, ByVal nHits As Long) As String
    ApplyLowHitWarning = result
    Dim threshold As Double
    threshold = modConfig.GetDouble("low_hit_warn_score", 0.3)
    If threshold <= 0 Then Exit Function

    Dim maxScore As Double
    maxScore = -1E+30
    Dim i As Long
    For i = 1 To nHits
        If hits(i).score > maxScore Then maxScore = hits(i).score
    Next i
    If maxScore >= threshold Then Exit Function

    ApplyLowHitWarning = ChrW(&H26A0) & ChrW(&HFE0F) & " 手元の資料との関連が薄い可能性があります。回答は参考程度にご覧ください。" & _
        vbLf & vbLf & result
End Function

' IsTooVague - 短すぎ かつ どの資料とも関連が薄い質問だけTrue。LLMを呼ばず
'   聞き方の例を返す。判定は全ヒットの最高スコア(1位だけ見ると誤発動する)。
Public Function IsTooVague(ByVal q As String, hits() As Hit, ByVal nHits As Long) As Boolean
    If nHits < 1 Then Exit Function

    Dim maxChars As Long, thr As Double
    On Error Resume Next
    maxChars = modConfig.GetLong("ambiguous_max_chars", 10)
    thr = CDbl(modConfig.GetLong("ambiguous_score_x100", 60)) / 100#
    On Error GoTo 0
    If thr <= 0# Then Exit Function              ' 0=無効化
    If maxChars < 1 Then maxChars = 10
    If Len(q) > maxChars Then Exit Function

    Dim best As Double, i As Long
    On Error Resume Next
    For i = LBound(hits) To UBound(hits)
        If i > nHits Then Exit For
        If hits(i).score > best Then best = hits(i).score
    Next i
    On Error GoTo 0

    IsTooVague = (best < thr)
End Function

' 逆質問の材料: ヒットした資料名を重複除去して最大4件、| 区切りで返す。
' R13 L-batch: 件数を呼び出し側で選べるようにした(既定4=従来と同一)。
' 会話の出典メモリ(modFollowup.RememberCitedSources)だけは8件で呼ぶ。
' 深掘りのスコープはその記憶が元になるため、4件で切ると「回答が引用した
' 資料なのに、次の深掘りでは対象外」という不可解な穴が空く。逆質問の
' 選択肢は多すぎると選べなくなるので4件のままにする。
Public Function HitSourceList(hits() As Hit, ByVal nHits As Long, _
                              Optional ByVal maxN As Long = 4) As String
    Dim sb As String, cnt As Long
    Dim lim As Long: lim = maxN
    If lim < 1 Then lim = 4
    Dim i As Long
    On Error Resume Next
    For i = 1 To nHits
        Dim nm As String: nm = Trim$(hits(i).source)
        If LenB(nm) > 0 Then
            If InStr(1, "|" & sb & "|", "|" & nm & "|", vbTextCompare) = 0 Then
                If LenB(sb) > 0 Then sb = sb & "|"
                sb = sb & nm
                cnt = cnt + 1
                If cnt >= lim Then Exit For
            End If
        End If
    Next i
    On Error GoTo 0
    HitSourceList = sb
End Function
