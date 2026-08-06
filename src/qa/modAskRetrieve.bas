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

' R17 Phase3(名寄せ): synonymsシートは質問1回の中で何度も読み直さない。
' 1セッション1回だけ modSynonymStore.ReadMapCsv を呼び、モジュール変数へ
' 控えて使い回す(取込・同期での更新は次セッションから反映=docs/10に明記)。
Private mSynMapCsv As String
Private mSynLoaded As Boolean

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

    ' 0) 語彙ズレの吸収(R17 Phase3): 質問文に一致した同義語を最大3語・
    ' 半角空白区切りで追記してから検索へ(modRetrieve/modSparseのスコアリング
    ' 本体は無改修=質問文への同義語追記だけで表記ゆれを吸収する)。
    If LCase$(Trim$(modConfig.GetString("graph_synonyms", "on"))) <> "off" Then
        If Not mSynLoaded Then
            mSynMapCsv = modSynonymStore.ReadMapCsv()
            mSynLoaded = True
        End If
        If LenB(mSynMapCsv) > 0 Then q = modRagParse.ExpandQueryBySyn(q, mSynMapCsv, 3)
    End If

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
    ' R17 Phase1: 質問が名指しした条番号(第5条/別表2 等)のチャンクが1件も
    ' 入っていなければ、chunk_meta から引いて先頭へ入れる(最大2件)。検索の
    ' 当て方は変えず、決定的なキーで最後に1回だけ確かめるだけ。
    ' R17H FA-3: スコープ指定つきの検索(深掘り)では、注入も【そのスコープの
    ' 中だけ】に限る。スコープ外の資料を注ぎ足すと、「会話の資料の中で見つけた
    ' 件数」を数える RunDeepScoped の n>=2 判定が、注入分だけで成立してしまう
    ' (スコープの外で答えたのに、スコープ内で答えたことになる)。
    modAskFocus.ArticleEnsure q, hits, outN, 2, ScopeBox(scopeSources)
    RunMultiRetrieve = EnsureArticleSeed(q, hits, outN)
    Exit Function

FallbackSingle:
    Err.Clear
    On Error GoTo 0
    outN = modRetrieve.Search(q, topK, hits, scopeSources)
    modAskFocus.ArticleEnsure q, hits, outN, 2, ScopeBox(scopeSources)
    RunMultiRetrieve = EnsureArticleSeed(q, hits, outN)
End Function

' ----------------------------------------------------------------------------
' EnsureArticleSeed - 検索が0件のとき、条文/別表/様式の直接キーで最大2件だけ
'   材料を用意する(2026-08-05 R17H FA-8 / A-M9)。戻り値=最終件数。
' ----------------------------------------------------------------------------
' 「第5条を見せて」に dense も sparse も1件も返さないことは実際に起きる
' (条番号は短くて特徴が薄い)。ところが ArticleEnsure は既存ヒットが1件以上
' あることを前提にしており、0件のときは何もしないまま「資料が見つかりません」
' で終わっていた。決定的なキー(section_path)は手元にあるのだから、最後に1回
' だけ引く。score は 0 のままなので低関連度の警告は付くが、それは正直な表示
' (検索では当たらなかった、という事実がそのまま出る)。
' n=-1(埋め込み失敗)には触らない。0件かどうかと理由が違う障害を混ぜない。
' ----------------------------------------------------------------------------
Private Function EnsureArticleSeed(ByVal q As String, ByRef hits() As Hit, _
                                   ByVal n As Long) As Long
    EnsureArticleSeed = n
    If n <> 0 Then Exit Function
    Dim m As Long: m = 0
    modAskFocus.ArticleEnsure q, hits, m, 2, "", True
    EnsureArticleSeed = m
End Function

' 許可資料名のDictionary(modRetrieve と同じ scopeSources)を、modAskFocus の
' AppendByIds が使う srcBox 形式("|資料名|" の連結)へ畳む。Nothing は空文字
' =スコープ無し(=従来動作)。判定そのものは modRetrieve の Exists と同型で、
' ここは【同じ許可集合を別の入口へ渡すための翻訳】だけを行う。
Private Function ScopeBox(ByVal scopeSources As Object) As String
    If scopeSources Is Nothing Then Exit Function
    Dim sb As String
    Dim v As Variant
    On Error Resume Next
    For Each v In scopeSources.Keys
        Dim nm As String: nm = Trim$(CStr(v))
        If LenB(nm) > 0 Then
            If InStr(1, sb, "|" & nm & "|", vbBinaryCompare) = 0 Then sb = sb & "|" & nm & "|"
        End If
    Next v
    On Error GoTo 0
    ScopeBox = sb
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
' FinishDeep - 深掘りで材料が確定したあとの後処理(2026-08-05 R16-3D)。
'   戻り値 = 後処理後の件数(=降格後 topK)。スコープ内で見つかった経路と、
'   本棚全体へ広げ直した経路の両方が必ずここを通る(片方だけに掛けると、同じ
'   「続けて質問」が本棚の当たり方で別の読み方をされる)。
'
'   R16-3D: 検索は topK の2倍で取り、既出チャンク(前回までの深掘りで実際に
'   渡したもの)を後ろへ回してから topK へ切る。除外ではなく降格なので、
'   本棚が小さくて既出しか無いときも件数は減らない(modFollowup.DemoteUsed)。
'   順序は 降格 → 記憶 の2段。記憶を降格の前に置くと、いま返すチャンクを
'   自分で既出扱いして全部後ろへ回す(1ターンで自滅する)。
'
'   2026-08-05(R16H FA-4 / A-M4): 精読(NeighborExpand)はここから外した。
'   modAsk 凍結下では deep の戻り件数がそのまま「N件ヒット」の件数バッジと
'   usage_log の hit_count になるため、近傍を混ぜた件数を返すと画面が嘘をつく
'   (「12件ヒット」と出るが検索で当たったのは4件)。精読は入念(thorough)の
'   役割とし、速度優先の deep には掛けない(仕様 §3C 改訂)。
' ----------------------------------------------------------------------------
Private Function FinishDeep(ByRef hits() As Hit, ByVal n As Long, ByVal topK As Long) As Long
    Dim m As Long: m = n
    modFollowup.DemoteUsed hits, m, topK
    modFollowup.RememberUsedChunks hits, m
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
' ----------------------------------------------------------------------------
' 2026-08-05(R16H FA-3 / A-M3・B-H2): 「続けて質問」のとき、モードに関係なく
'   既出チャンクの降格をここで掛ける。R16-3D は deep の RunDeepScoped だけに
'   降格を入れたが、深掘りは deep 専用の機能ではない。入念(thorough)や
'   すぐ聞く(quick)で「続けて質問」を押した人は RunUnscoped しか通らないため、
'   2回目の検索が1回目と1件も違わないチャンクを返し、同じ回答が返っていた
'   (深掘りループの残り半分)。判定は modFollowup.IsFollowupTurn() をここで読む
'   =modAsk を1行も触らずに全モードへ同じ規則を効かせられる。
'   精読(NeighborExpand)はここでは呼ばない。thorough の draft 直前で1回だけ
'   掛ける設計(仕様 §3C 改訂)で、ここで足すと deep の件数バッジが嘘になる。
'   RunDeepScoped のスコープ縮小フォールバックもここを通るため降格が2回走るが、
'   1回目で今回のチャンクを記憶済み=2回目は「全部既出」となり順序を変えずに
'   topK へ切るだけになる(結果は同じ。仕様どおり自動で効く側に倒す)。
' ----------------------------------------------------------------------------
Public Function RunUnscoped(ByVal q As String, ByVal mdMode As String, _
                            ByVal topK As Long, ByRef hits() As Hit) As Long
    If modFollowup.IsFollowupTurn() Then
        Dim n As Long
        n = RunMultiRetrieve(q, mdMode, WideK(topK), hits)
        If n < 1 Then
            RunUnscoped = n            ' 0件・埋め込み失敗(-1)はそのまま返す
            Exit Function
        End If
        modFollowup.DemoteUsed hits, n, topK
        modFollowup.RememberUsedChunks hits, n
        RunUnscoped = n
        Exit Function
    End If

    If LCase$(modConfig.GetString("retrieve_mode", "single")) = "multi" Then
        RunUnscoped = RunMultiRetrieve(q, mdMode, topK, hits)
    Else
        RunUnscoped = modRetrieve.Search(q, topK, hits)
    End If
End Function

' ----------------------------------------------------------------------------
' 段階ナレーション(R13-9b): この質問で何段通すかを先に決めてから実況する。
' ----------------------------------------------------------------------------
' 2026-08-05(R17H FA-2補 / 司令塔裁定): 俯瞰の印もここで下ろす。
' 裁定の当初位置 modAskMulti.TryDecomposed 入口は【入念モードでしか通らない】
' ため、俯瞰で答えた次のターンを ⚡すぐ聞く / 🔍しっかり調べる で質問すると、
' 印が1ターンぶん持ち越されて「🔭 章の要約に基づく回答(俯瞰)」のバッジと
' 低関連度警告の抑止が、俯瞰していない回答に付いていた。この Sub は
' modAsk.Answer の入口から【全モードで】必ず1回通るので、ここが全経路に
' 効く唯一の場所になる(TryDecomposed 側の既存リセットは残す=二重リセットは
' 無害。印を立てるのは TryGlobal の成功出口だけで、順序は
' PlanAskStages → TryDecomposed → TryGlobal のため表示までは必ず生き残る)。
' RunDeepScoped からの呼び直し(番号計画を通常構成へ戻す)もここを通るが、
' あちらは deep 専用の経路で TryGlobal を一度も通らない=消す印が無い。
' ----------------------------------------------------------------------------
Public Sub PlanAskStages(ByVal mdMode As String)
    On Error Resume Next
    modAskGlobal.ResetGlobalTurn
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

' R20-6a: 直近ターンで計画した段の総数(0=番号なし=すぐ聞く)。回答完了後に
' modLive.Footer が読み、フッターへ「(N段)」を付す。mStgTotalの寿命は
' PlanAskStages/RunDeepScoped が1ターンごとに立て直す既存の仕組みそのもの
' なので、ここは値を外へ見せるだけの薄い読み取り専用の窓口。
Public Function LastStageTotal() As Long
    LastStageTotal = mStgTotal
End Function

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
    ' 2026-08-05(R17H FA-2 / A-H2・B-H1): 俯瞰で答えたターンは無操作。
    ' 俯瞰の hits は章の要約から選んだもので score=0 が正しい値(検索スコアを
    ' 騙らない)。そこへこの警告を足すと、章をまたいで正しく答えた回答に
    ' 「手元の資料との関連が薄い」と書くことになる=表示が事実と逆になる。
    If modAskGlobal.WasGlobalTurn() Then Exit Function

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

' IsTooVague - 短すぎ かつ 資料が絞れない質問だけTrue。LLMを呼ばず
'   聞き方の例を返す。判定は全ヒットの最高スコア(1位だけ見ると誤発動する)。
' ----------------------------------------------------------------------------
' 2026-08-06(R19-4b・実機第6報④): 「絞れない」に2つ目の意味を足した。
'   (A) 低スコア: どの資料にも当たらない(従来)
'   (B) 資料分散: 局所的には当たるのに、当たり先が複数の資料に割れている
' 「免責は?」は(B)で、4字なので段0の分解ゲート(25字)にも掛からず、
' best>=0.6 なので(A)にも掛からない=どちらの網もすり抜けていた(調査④班1章)。
' 深掘りの「計算方法」も同じ経路(IsTooVague は modAsk.Answer から全モード・
' 新規/続けて質問の別なく必ず1回通る)で、同じ理由で同じ穴に落ちていた。
' 判定の本体は modClarify の純関数2本(HasScoreDispersion / MentionsSourceName)。
' ここは hits() を文字列へ畳んで渡すだけ=LOでテストできない Hit() の扱いを
' この1関数に閉じ込め、規則そのものはテストで固定できる側に置く。
' ----------------------------------------------------------------------------
Public Function IsTooVague(ByVal q As String, hits() As Hit, ByVal nHits As Long) As Boolean
    If nHits < 1 Then Exit Function
    ' 「全体像は?」のような俯瞰の短文は聞き返さない(2026-08-05 R17H FA-6)。
    ' 短くて低スコアなのは俯瞰質問の常態(点検索が当たらないから俯瞰なのに、
    ' 当たらないことを理由に聞き返すと、俯瞰は永久に一度も試されない)。
    If modRagParse.HasGlobalSignal(q) Then Exit Function

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

    ' (A) 低スコア。こちらが立つときは分散を見ない: 2つの理由が同時に成り立つ
    ' 場合の文面は「低スコア」用が正しい(「複数の資料に該当します」と言いながら
    ' どれも薄い、という辻褄の合わない聞き返しになる。調査④班6.1)。
    If best < thr Then
        IsTooVague = True
        Exit Function
    End If

    ' (B) 資料分散(R19-4b)。gap=0 で機能OFF(既存の閾値0と同じ思想)。
    ' R20-6d: 「入念に調べる」だけは時間より精度(modMode方針)なので、資料
    ' 確認を優先する姿勢へ倒し、閾値を広げる(既定quick/deep=10、thorough=25)。
    ' モードはmodAskを触らず、modAsk.ReadModeFromUiState(Private)と同型の
    ' 直読みをここに持つ(R1: qa層からui層のmodAppStateは参照できないため。
    ' AskFollowupが送信時点に同型の再読を行っている前提と整合)。
    Dim curMode As String: curMode = CurrentRagSpeedForDispersion()
    Dim gapX100 As Long
    On Error Resume Next
    gapX100 = DispersionThresholdFor(curMode, _
        modConfig.GetLong("ambiguous_dispersion_gap_x100", 10), _
        modConfig.GetLong("thorough_dispersion_gap_x100", 25))
    On Error GoTo 0
    If gapX100 <= 0 Then Exit Function

    ' R19-4d: 質問文が資料を名指ししているなら聞き返さない(誤発動対策)。
    Dim srcList As String: srcList = HitSourceList(hits, nHits)
    If modClarify.MentionsSourceName(q, srcList) Then Exit Function

    ' R19H FB-1(A-M⑧): 判定は「gapを返す純関数+ここでの比較」に分けた。
    ' 発動した回も【しなかった回も】gapの実値を1行残す ―― 既定0.10が高すぎるか
    ' 低すぎるかは、鳴らなかった側の分布を見ないと決められない(憲章§4-2)。
    ' 資料が2種類未満(gap=-1)は判定不能なので、その事実もそのまま記録する。
    Dim srcN As Long
    Dim gap As Long: gap = modClarify.DispersionGapX100(FoldSrcScoreLines(hits, nHits), srcN)
    On Error Resume Next
    ' R20-6b: mode列に資料数(srcN)が入っていた列崩れを是正。実モード文字列を
    ' mode列へ、資料数とgapはdetail列へ寄せる。
    modLog.LogUsage "dispersion", curMode, "src=" & srcN & " gap=" & gap
    On Error GoTo 0
    If gap >= 0 And gap < gapX100 Then
        ' 聞き返しの1行目だけを分散用に差し替えるための印(modClarify が使ったら
        ' その場で下ろす)。文面の本体=資料選択+意図5区分は従来のまま使う。
        modClarify.NoteDispersion
        IsTooVague = True
    End If
End Function

' R20-6d: ui_state "mode" キーの直読み(modAsk.ReadModeFromUiStateと同型)。
' modAskは触らず(凍結)、R1(qa層からui層のmodAppStateは参照不可)を守るため、
' 同じ形の読み取りをここに複製する(modAsk/modShelfSync/modStateと同じ
' 「ThisWorkbook.Worksheets(SH_UISTATE)を直読みする」既存の作法)。
Private Function CurrentRagSpeedForDispersion() As String
    Dim v As String: v = "quick"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_UISTATE)
    On Error GoTo 0
    If Not ws Is Nothing Then
        On Error Resume Next
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
        Dim i As Long
        For i = 1 To lastRow
            If StrComp(CStr(ws.Cells(i, 1).Value), "mode", vbTextCompare) = 0 Then
                v = LCase$(Trim$(CStr(ws.Cells(i, 2).Value)))
                Exit For
            End If
        Next i
        On Error GoTo 0
    End If
    CurrentRagSpeedForDispersion = modMode.Normalize(v)
End Function

' R20-6d: 分散閾値をモード別に選ぶ純関数(LOテスト対象)。「入念に調べる」
' だけ資料確認を優先する姿勢(=閾値を広げて聞き返しやすくする)にする。
' モードの読み出しと config の既定値は呼び出し側(IsTooVague)が持つ。
Public Function DispersionThresholdFor(ByVal mode As String, _
                                       ByVal baseGapX100 As Long, _
                                       ByVal thoroughGapX100 As Long) As Long
    If modMode.Normalize(mode) = "thorough" Then
        DispersionThresholdFor = thoroughGapX100
    Else
        DispersionThresholdFor = baseGapX100
    End If
End Function

' hits() を「資料名<TAB>スコア」の vbLf 連結へ畳む(R19-4b)。
' Hit() を受ける唯一の場所で、ここだけはLO実行テストの対象外になる
' (別モジュールの Public Type 配列はLOで ReDim できない既知制約)。
' 畳んだ後の判定は modClarify.HasScoreDispersion が担い、そちらはテストで固定する。
' スコアは Str$ で書く: CStr は端末の小数点記号(カンマ圏)に従うため、
' Val で読み直す受け側と食い違い得る。Str$ は常に "." で書く。
Private Function FoldSrcScoreLines(hits() As Hit, ByVal nHits As Long) As String
    Dim sb As String
    Dim i As Long
    On Error Resume Next
    For i = 1 To nHits
        Dim nm As String: nm = Trim$(hits(i).source)
        If LenB(nm) > 0 Then
            If LenB(sb) > 0 Then sb = sb & vbLf
            sb = sb & nm & vbTab & Trim$(Str$(hits(i).score))
        End If
    Next i
    On Error GoTo 0
    FoldSrcScoreLines = sb
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
