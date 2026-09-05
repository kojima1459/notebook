Attribute VB_Name = "modAskThorough"
Option Explicit

' ============================================================================
' modAskThorough - 「入念に調べる」専用の生成パイプライン(2026-08-03 R14-8a)
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか(実機第3報 RC8):
'   入念モードは、深掘り(RunDeepFlow)と【生成側が完全に同じ】だった。
'   下書きと検証のプロンプトも、モデルも、effort/verbosity も共有していて、
'   違うのは検索の幅(topK・サブクエリ本数)だけ。利用者から見れば
'   「同じことを、少しだけ多くの資料でやり直しているのに数分待たされる」
'   モードになっていた。しかも検証は1発生成で、回答に付いた出典タグが
'   本当に検索結果に在るかを機械的に確かめる仕組みは1つも無かった
'   (=存在しない資料名やページを堂々と引用しても誰も気付けない)。
'
'   そこで入念だけを別の道筋にする。段は6つ(番号は modMode.AskStageIndex の
'   実況番号と同じ並び。(1)は検索側=modAskRetrieve が担当する):
'     (1) 再ランク       … 入念だけ専用の effort で候補を並べ直す(modMode.RerankEffort)
'     (2) 資料の要点整理 … 抜粋を1回のLLM呼び出しで質問に関係する部分だけへ畳む
'     (3) 下書き         … 要点+原文のハイブリッドを根拠に書かせる(effort=high)
'     (4) 自己批判       … 自分の下書きの問題点【だけ】を挙げさせる(書き直させない)
'     (5) 検証           … 指摘を踏まえて本棚抜粋と照合し直させる(effort=high)
'     (6) 出典突合       … 回答中の出典タグを hits() と機械的に照合する(LLM不使用)
'   quick / deep は1行も変わらない。
'
' 設計判断:
'   ・(4)で書き直しまでさせない。「指摘と修正を1回でやれ」と言うと、モデルは
'     指摘を軽くして自分の文章を守る。指摘だけを吐かせてから、別の呼び出しに
'     直させるほうが、実際に消える断定の数が多い。
'   ・(2)と(4)は補助段。失敗しても回答自体は成立するので、握って先へ進み
'     usage_log に痕跡だけ残す(憲章§4-1: 無言の失敗禁止、ただし止めない)。
'     下書き(3)の失敗だけは #ERR: をそのまま返して呼び出し元に処理させる。
'   ・(6)は純ロジック。LLMに「出典が正しいか」を聞いても、聞かれたモデルは
'     自分が書いたタグを肯定する。文字列として在るか無いかを数えるだけの
'     検査は、遅くもなく、嘘もつかない。突合の基準になるタグの形は
'     modPrompts.SourceTag が単一情報源(指示する形と検査する形が別々に
'     なった瞬間、この検査は全件不一致か全件一致のどちらかに退化する)。
'   ・Excelオブジェクトには一切触れない。実況は modAskRetrieve.ShowAskStage
'     へ委ね、(6)の純関数群は LibreOffice の実行テストから直接呼ぶ
'     (tools/run_lo_tests.py の PURE_ALLOWLIST に登録済み)。
' ============================================================================

Private Const MODE_THOROUGH As String = "thorough"

' 出典が検索結果に見つからなかったタグへ付ける印。回答の見た目を壊さずに
' 「この1件だけは裏が取れていない」と言い切れる最小の書き方にする。
Public Const UNVERIFIED_MARK As String = "(出典確認できず)"

' 1回の回答から拾う出典タグの上限(異常な応答でループを長引かせないため)。
Private Const MAX_CITE_TAGS As Long = 64

' 資料名に "]" を含むファイル名(report[1].pdf 等)を救済するとき、
' タグの終端を次の "]" まで伸ばす回数の上限(2026-08-03 R14-G3)。
Private Const MAX_TAG_EXTEND As Long = 2

' 検証段が落ちたターンの内部注記(2026-08-03 R14-G11)。回答本文には混ぜず、
' 呼び出し元(modAsk)が深掘り候補の整形【後】に足す。本文へ混ぜると
' mLastCleanAnswer=会話履歴と「解決済みQ&A」の部内共有にまで注記が入り、
' 「人が確認した回答」として配られる文面が内部事情で汚れる(deepと同型)。
Private mVerifyNote As String

' VerifyNote - 直近の入念モード実行で検証段が落ちたときの注記(無ければ空)。
Public Function VerifyNote() As String
    VerifyNote = mVerifyNote
End Function

' ----------------------------------------------------------------------------
' RunThoroughFlow - 入念モードの本体。戻り値は最終回答本文、または "#ERR:…"。
'   ok = True のとき呼び出し元(modAsk)が深掘り候補の整形と履歴登録を行う。
' ----------------------------------------------------------------------------
Public Function RunThoroughFlow(ByVal q As String, hits() As Hit, ByVal nHits As Long, _
                                ByRef ok As Boolean, ByVal history As String, _
                                ByVal prevU As String, ByVal prevA As String) As String
    ok = False
    mVerifyNote = ""

    ' R38 §2: 1回読み(章丸ごと)。失敗時は黙って従来の4段へ落ちる。
    If modAskOnePass.Enabled() Then
        Dim opBody As String
        If modAskOnePass.TryOnePass(q, hits, nHits, history, prevU, prevA, opBody) Then
            ok = True
            RunThoroughFlow = opBody
            Exit Function
        End If
    End If

    Dim strictG As Boolean: strictG = modConfig.GetBool("strict_grounding", False)
    Dim ansTags As Boolean: ansTags = modConfig.GetBool("answer_tags", False)
    Dim mdl As String: mdl = modConfig.GetString("recommended_model", "gpt-5.5")
    Dim lat As Long

    ' --- (0) 精読: 根拠チャンクの前後を一緒に読む(2026-08-05 R16-3C) --------
    ' 当たったチャンクだけだと、表の見出しと値・条文と但し書きのように
    ' 意味の単位がチャンク境界をまたぐ資料で「書いてあるのに確認できない」に
    ' なる。件数だけをローカルで増やし、引数の nHits(=呼び出し元 modAsk の
    ' 出典チップ・信頼度の材料)は動かさない。近傍は根拠であって【ヒット】では
    ' ないので、検索結果の件数として数えると信頼度バッジが水増しされる。
    Dim nUse As Long: nUse = nHits
    modAskFocus.NeighborExpand hits, nUse, modConfig.GetLong("deep_neighbor", 2)
    ' R17 Phase1: 物理近傍のあとに【参照先】も束ねる(「第6条を読むなら、6条が
    ' 参照する第8条と別表2も同じ束で読む」)。deep_neighbor=0(精読off)でも
    ' 効かせる=別の軸の機能で、config graph_refs でだけ切れる。
    modAskFocus.RefsExpand hits, nUse, 0    ' 0=既定(最大8件)

    ' --- (2) 資料の要点整理 -------------------------------------------------
    modAskRetrieve.ShowAskStage "digest"
    Dim digest As String
    digest = modGateway.CallLLM(ThoroughDigestPrompt(modPrompts.BuildSourceDigestPrompt(q, hits, nUse)), _
        "thorough_digest", modConfig.GetString("thorough_digest_effort", "low"), _
        "medium", mdl, lat)
    If modRagParse.IsErrorResponse(digest) Then
        ' 要点整理は「あれば効く」補助段。ここで止めると、資料は揃っているのに
        ' 答えが出ないという一番惜しい失敗になる。原文だけで下書きへ進む。
        On Error Resume Next
        modLog.LogUsage "thorough_digest_skip", MODE_THOROUGH, modUtil.SafeLeft(digest, 120)
        On Error GoTo 0
        digest = ""
    Else
        digest = modAsk.ApplyAnswerTags(digest)
    End If

    ' --- (3) 下書き ---------------------------------------------------------
    modAskRetrieve.ShowAskStage "draft"
    Dim draft As String
    draft = modGateway.CallLLM( _
        ThoroughDraftPrompt(modPrompts.BuildDeepDraftPrompt(q, hits, nUse, history, strictG, ansTags, digest)), _
        "thorough_draft", modConfig.GetString("thorough_draft_effort", "high"), _
        modConfig.GetString("thorough_draft_verbosity", "high"), mdl, lat, prevU, prevA)
    If modRagParse.IsErrorResponse(draft) Then
        RunThoroughFlow = draft
        Exit Function
    End If

    Dim draftBody As String
    draftBody = modAsk.ApplyAnswerTags(draft)

    ' --- (4) 自己批判(指摘のみ。書き直しはさせない) -----------------------
    modAskRetrieve.ShowAskStage "critique"
    Dim critique As String
    ' 冗長性は最小で固定する。この段の出力は「番号付きの指摘」だけで、
    ' 長くなるほど次の検証段が読み落とす(configで伸ばせる意味が無い)。
    critique = modGateway.CallLLM(modPrompts.BuildCritiquePrompt(q, draftBody, hits, nUse), _
        "thorough_critique", modConfig.GetString("thorough_critique_effort", "medium"), _
        "low", mdl, lat)
    If modRagParse.IsErrorResponse(critique) Then
        On Error Resume Next
        modLog.LogUsage "thorough_critique_skip", MODE_THOROUGH, modUtil.SafeLeft(critique, 120)
        On Error GoTo 0
        critique = ""
    Else
        critique = modAsk.ApplyAnswerTags(critique)
    End If

    ' --- (5) 検証(指摘を踏まえて抜粋と照合し直す) -------------------------
    modAskRetrieve.ShowAskStage "verify"
    Dim verified As String
    verified = modGateway.CallLLM( _
        ThoroughVerifyPrompt(modPrompts.BuildDeepVerifyPrompt(q, draftBody, hits, nUse, strictG, ansTags, critique)), _
        "thorough_verify", modConfig.GetString("thorough_verify_effort", "high"), _
        modConfig.GetString("deep_verify_verbosity", "medium"), mdl, lat)

    ok = True
    Dim body As String
    If modRagParse.IsErrorResponse(verified) Then
        ' 注記は本文へ混ぜない(mVerifyNoteの宣言部を参照)。呼び出し元が
        ' 深掘り候補の整形後に足すので、履歴にも部内共有にも入らない。
        body = draftBody
        mVerifyNote = vbLf & vbLf & _
            "(注: 検証段階でエラーが発生したため、下書きの内容を表示しています。)"
    Else
        body = modAsk.ApplyAnswerTags(verified)
    End If

    ' --- (6) 出典の機械的突合 -----------------------------------------------
    RunThoroughFlow = AnnotateAgainstHits(body, hits, nUse)
End Function

' ----------------------------------------------------------------------------
' ThoroughVerifyPrompt - R20-6f(入念モードの文体差別化)。
'   modPrompts.BuildDeepVerifyPrompt(凍結)の戻り文字列に、入念モードだけの
'   文体指示を連結する。deep/quickは1文字も変えない(呼ばれるのはこの
'   関数だけ)。純関数なのでLibreOfficeの実行テストで固定できる。
' ----------------------------------------------------------------------------
Public Function ThoroughVerifyPrompt(ByVal basePrompt As String) As String
    ThoroughVerifyPrompt = basePrompt & vbLf & vbLf & ThoroughStyleAddendum()
End Function

Private Function ThoroughStyleAddendum() As String
    ThoroughStyleAddendum = _
        "■見出しで構造化し、各主張の直後に出典を明記。必要なら600" & ChrW(&H301C) & _
        "900字まで許容。断定できない点は「資料からは確認できません」と明示すること。"
End Function

' ----------------------------------------------------------------------------
' R21-2 D3(実機第8報⑧・逆質問の網羅性): 要点整理(digest)段・下書き(draft)段
' にも「免責事由」観点を連結する。ThoroughVerifyPrompt/ThoroughStyleAddendum
' と同型の連結機構(modPrompts本体=凍結は1文字も変えない)。deep/quickは
' この2関数を通らないので1文字も変わらない。
' ----------------------------------------------------------------------------
Public Function ThoroughDigestPrompt(ByVal basePrompt As String) As String
    ThoroughDigestPrompt = basePrompt & vbLf & vbLf & ExemptionCoverageAddendum()
End Function

Public Function ThoroughDraftPrompt(ByVal basePrompt As String) As String
    ThoroughDraftPrompt = basePrompt & vbLf & vbLf & ExemptionCoverageAddendum()
End Function

Private Function ExemptionCoverageAddendum() As String
    ExemptionCoverageAddendum = _
        "■支払う場合だけでなく、支払わない場合(免責事由・除外・適用除外)も" & _
        "対象に含め、資料にあれば必ず言及すること。"
End Function

' ----------------------------------------------------------------------------
' R30 W2-6(実機第15報D班・表記揺れ道1): 検索前クエリ拡張(modAskRetrieveの
' expand段)の指示文へ「他言語・別表記の言い換えも含めること」を追記する。
' modPromptsは凍結のため不触。ThoroughVerifyPrompt等と同型の連結機構だが、
' expand段は全モード共通(thorough限定ではない)ため接頭辞をThoroughにしない。
' quick/light系はそもそもexpand段自体が既定オフなので対象外(記録のみ)。
' ----------------------------------------------------------------------------
Public Function ExpandPromptWithSynonymHint(ByVal basePrompt As String) As String
    ExpandPromptWithSynonymHint = basePrompt & vbLf & vbLf & SynonymHintAddendum()
End Function

Private Function SynonymHintAddendum() As String
    SynonymHintAddendum = _
        "■言い換え候補には、他言語(英語等)・別表記(カナ/かな/ローマ字/漢字)の" & _
        "言い換えを1" & ChrW(&H301C) & "2語含めること。"
End Function

' ============================================================================
' (6) 出典突合 ― ここから下は純ロジック(LibreOfficeの実行テストで固定する)
' ============================================================================

' ----------------------------------------------------------------------------
' AnnotateAgainstHits - 回答中の出典タグを hits() と照合し、見つからないものへ
'   「(出典確認できず)」を付ける。不一致が1件でもあれば usage_log に残す。
' ----------------------------------------------------------------------------
Public Function AnnotateAgainstHits(ByVal answerText As String, hits() As Hit, _
                                    ByVal nHits As Long) As String
    Dim mismatchN As Long
    AnnotateAgainstHits = AnnotateCitations(answerText, CiteIndexFrom(hits, nHits), mismatchN)
    If mismatchN > 0 Then
        ' 出典が合わないことは、回答そのものより重い事実。件数を必ず残す
        ' (何件に1件ずれるのかが分からないと、プロンプト側の直し方が決まらない)。
        On Error Resume Next
        modLog.LogUsage "cite_mismatch", MODE_THOROUGH, _
            "検索結果に無い出典タグ " & mismatchN & " 件に「" & UNVERIFIED_MARK & "」を付記", _
            0, nHits
        On Error GoTo 0
    End If
End Function

' ----------------------------------------------------------------------------
' CiteIndexFrom - hits() から「正しい出典タグの集合」を作る。
'   形式: "|タグ1||タグ2||タグ3|"(前後を | で挟むので部分一致で誤判定しない)。
'   タグの文字列は modPrompts.SourceTag(LLMへ指示している形の単一情報源)。
' ----------------------------------------------------------------------------
Public Function CiteIndexFrom(hits() As Hit, ByVal nHits As Long) As String
    Dim sb As String
    Dim i As Long
    On Error Resume Next
    For i = 1 To nHits
        Dim k As String
        k = NormalizeCiteTag(modPrompts.SourceTag(hits(i)))
        If LenB(k) > 0 Then
            If InStr(1, sb, "|" & k & "|", vbTextCompare) = 0 Then sb = sb & "|" & k & "|"
        End If
    Next i
    On Error GoTo 0
    CiteIndexFrom = sb
End Function

' ----------------------------------------------------------------------------
' NormalizeCiteTag - 突合用の正規化。空白(半角/全角/タブ)を全部落とす。
'   「[本棚: 規約集 p. 12]」と「[本棚:規約集 p.12]」を同じものとして扱うため。
'   資料名に空白が含まれていても、突合の両側が同じ処理を通るので破綻しない。
' ----------------------------------------------------------------------------
Public Function NormalizeCiteTag(ByVal tag As String) As String
    Dim t As String
    t = Replace(tag, " ", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, ChrW(&H3000), "")
    NormalizeCiteTag = t
End Function

' ----------------------------------------------------------------------------
' ExtractCiteTags - 回答本文から出典タグを出現順に取り出す。
'   戻り値=件数。tags() はタグ文字列、tagEnd() はその "]" の位置(1起点)。
'   どちらも 0 始まりで、件数ぶんだけ意味がある。
'   [[FOLLOWUP: …]] のような別のマーカーは "[本棚:"/"[パック(" で始まらない
'   ので拾わない(拾うと深掘り候補に「(出典確認できず)」が付く)。
' ----------------------------------------------------------------------------
Public Function ExtractCiteTags(ByVal s As String, ByRef tags() As String, _
                                ByRef tagEnd() As Long) As Long
    ReDim tags(0 To MAX_CITE_TAGS - 1)
    ReDim tagEnd(0 To MAX_CITE_TAGS - 1)
    If LenB(s) = 0 Then Exit Function

    Dim n As Long: n = 0
    Dim i As Long: i = 1
    Dim total As Long: total = Len(s)

    Do While i <= total
        If n >= MAX_CITE_TAGS Then Exit Do
        Dim p As Long
        p = InStr(i, s, "[")
        If p = 0 Then Exit Do
        Dim e As Long
        e = InStr(p + 1, s, "]")
        If e = 0 Then Exit Do

        Dim cand As String
        cand = Mid$(s, p, e - p + 1)
        If IsCiteTag(cand) Then
            tags(n) = cand
            tagEnd(n) = e
            n = n + 1
        End If
        i = e + 1
    Loop

    ExtractCiteTags = n
End Function

' 出典タグらしさの判定(§7.3の2形式のみ)。
Public Function IsCiteTag(ByVal tag As String) As Boolean
    Dim t As String
    t = NormalizeCiteTag(tag)
    If Left$(t, 4) = "[本棚:" Then
        IsCiteTag = True
    ElseIf Left$(t, 5) = "[パック(" Then
        IsCiteTag = True
    End If
End Function

' ----------------------------------------------------------------------------
' TagIsKnown - そのタグが検索結果(citeIndex)に在るか。
'   ページ番号まで含めた完全一致(±0)で見る。ページがずれた引用は、原文を
'   開いた利用者が「書いていない」と感じる典型なので、寛容にしない。
' ----------------------------------------------------------------------------
Public Function TagIsKnown(ByVal tag As String, ByVal citeIndex As String) As Boolean
    If LenB(citeIndex) = 0 Then Exit Function
    Dim k As String
    k = NormalizeCiteTag(tag)
    If LenB(k) = 0 Then Exit Function
    TagIsKnown = (InStr(1, citeIndex, "|" & k & "|", vbTextCompare) > 0)
End Function

' ----------------------------------------------------------------------------
' AnnotateCitations - 本文中の出典タグのうち citeIndex に無いものの直後へ
'   「(出典確認できず)」を差し込み、その件数を mismatchN で返す(純関数)。
'
'   citeIndex が空(=突合の材料が無い)ときは何もしない。材料が無い状態で
'   全件に「確認できず」と付けるのは、検査したふりの中でいちばん質が悪い。
' ----------------------------------------------------------------------------
Public Function AnnotateCitations(ByVal answerText As String, ByVal citeIndex As String, _
                                  ByRef mismatchN As Long) As String
    mismatchN = 0
    AnnotateCitations = answerText
    If LenB(answerText) = 0 Then Exit Function
    If LenB(citeIndex) = 0 Then Exit Function

    Dim tags() As String
    Dim tagEnd() As Long
    Dim n As Long
    n = ExtractCiteTags(answerText, tags, tagEnd)
    If n = 0 Then Exit Function

    Dim outS As String
    Dim prev As Long: prev = 0
    Dim i As Long
    For i = 0 To n - 1
        ' 伸ばしたタグが次のタグを飲み込んだ場合は、その次のタグを飛ばす
        ' (同じ範囲を二度書き出さない。Mid$の長さが負になる事故も防ぐ)。
        If tagEnd(i) > prev Then
            Dim tg As String: tg = tags(i)
            Dim te As Long: te = tagEnd(i)
            Dim known As Boolean
            known = ResolveTag(answerText, citeIndex, tg, te)
            outS = outS & Mid$(answerText, prev + 1, te - prev)
            prev = te
            If Not known Then
                outS = outS & UNVERIFIED_MARK
                mismatchN = mismatchN + 1
            End If
        End If
    Next i
    outS = outS & Mid$(answerText, prev + 1)

    AnnotateCitations = outS
End Function

' ----------------------------------------------------------------------------
' ResolveTag - タグが突合表に在るかを判定し、必要なら終端を伸ばして再判定する
'   (2026-08-03 R14-G3)。
' ----------------------------------------------------------------------------
' 実機の資料には "report[1].pdf" のように名前へ "]" を含むものが普通にある
' (ブラウザやメールの重複ダウンロードが自動で付ける番号)。抽出は最初の "]"
' で切るため、そのままでは
'   ・突合表(正しい形は "[本棚:report[1].pdf p.3]")と一致せず全件が不一致になり
'   ・「(出典確認できず)」がタグの途中("report[1]" の直後)へ差し込まれて
'     残りの ".pdf p.3]" が本文に取り残される
' という二重の壊れ方をする。次の "]" まで最大MAX_TAG_EXTEND回だけ伸ばして
' 再照合し、一致したらタグ本体と終端位置の両方をその位置へ更新する。
' 伸ばして一致しなければ元のまま(=不一致として付記)へ戻る。
Private Function ResolveTag(ByVal s As String, ByVal citeIndex As String, _
                            ByRef tag As String, ByRef tagEndPos As Long) As Boolean
    ResolveTag = TagIsKnown(tag, citeIndex)
    If ResolveTag Then Exit Function

    Dim startPos As Long
    startPos = tagEndPos - Len(tag) + 1
    If startPos < 1 Then Exit Function

    Dim e As Long: e = tagEndPos
    Dim k As Long
    For k = 1 To MAX_TAG_EXTEND
        e = InStr(e + 1, s, "]")
        If e = 0 Then Exit Function
        Dim cand As String
        cand = Mid$(s, startPos, e - startPos + 1)
        If TagIsKnown(cand, citeIndex) Then
            tag = cand
            tagEndPos = e
            ResolveTag = True
            Exit Function
        End If
    Next k
End Function
