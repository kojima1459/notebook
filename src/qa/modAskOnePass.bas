Attribute VB_Name = "modAskOnePass"
Option Explicit

' ============================================================================
' modAskOnePass - 「入念に調べる」の1回読み(章丸ごと・メガコンテキスト) R38
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか(R38 仕様書 §0/§2):
'   現行の入念フローは 4段(要点整理→下書き→自己点検→検証)で構成され、
'   往復回数が多く応答完了まで数分を要していた。
'   髙橋情報および外部レビュー裏どりにより、リボン経由で 20 万トークン
'   (日本語約 30 万字)までの長文コンテキストが一括で通ることが判明したため、
'   検索で当たった資料の【章丸ごと】を文書順に束ねて 1 回の LLM 呼び出しで
'   高精度に回答を生成する「1回読み」経路を新設する。
'
' 設計判断:
'   ・「入念に調べる」の単発経路(複合質問の分解・全体像の俯瞰以外)のみを置換。
'     しっかり(deep)・すぐ聞く(quick)は1文字も変えない。
'   ・フェイルセーフ: 章も近傍も引けない・例外のときは静かに False を返し、
'     呼び出し元(modAskThorough)が従来の4段へ自動退化する(modAskGlobal.TryGlobal
'     と同等の思想。答えが出ない状態を作らない)。ただし LLM 呼び出し自体の
'     エラー(ESC中断・NW断・上限)は4段へ落とさず、4段と同じくエラー文をそのまま
'     返す(R38 Fix F3。止めたい利用者を4段でさらに数分待たせない)。
'   ・hits は必ずローカルへ複製して扱い、呼び出し元の hits/nHits は不変に保つ
'     (出典チップ・信頼度バッジは検索結果そのものを見せる原則)。
'   ・モジュール変数は mWasOnePass(フッターの段数表示用・R38 Fix F4)の1つだけ。
'   ・出典タグの書式は唯一の持ち主である modPrompts.SourceTag だけを通す。
' ============================================================================

' 1章あたりの最低文字予算(極端に短い切り捨てを防ぐ)
Private Const PER_CHAPTER_MIN_CAP As Long = 2000

' 重点抜粋(Lost-in-the-middle 対策)として載せる上位ヒットの天井
Private Const MAX_FOCUS_HITS As Long = 3

' 1回読みで集める章内チャンク件数の天井(CollectChapterHits に渡す安全弁)
Private Const MAX_CHAPTER_CHUNKS As Long = 300

' R38 Fix F4: 直近のターンが1回読みで完結したか(フッターの「(6段)」表示との
' 食い違いを消すため。modLive.Footer が ConsumeOnePassTurn で1回だけ消費する)。
Private mWasOnePass As Boolean

' ----------------------------------------------------------------------------
' Enabled - config thorough_onepass の有効判定。
' なぜそうするか:
'   設定シートの1行(thorough_onepass=off)で即座に従来の4段へ戻せる
'   切り戻し手段を保証するため。IsOn で安全に解釈する。
' ----------------------------------------------------------------------------
Public Function Enabled() As Boolean
    Dim valStr As String
    valStr = modConfig.GetString("thorough_onepass", "on")
    Enabled = IsOn(valStr)
End Function

' ----------------------------------------------------------------------------
' ConsumeOnePassTurn - R38 Fix F4。直近のターンが1回読みで完結したか(True=
' 成功・LLMエラーの2出口のいずれか)を返し、フラグを False へ戻す。
' なぜそうするか:
'   modLive.Footer の「(N段)」表示は modAskRetrieve.LastStageTotal() の
'   段数(4段=要点整理→下書き→自己点検→検証+再ランク等)をそのまま出すため、
'   1回読みに置き換わったターンでも従来のまま「(6段)」等と表示され実態と
'   食い違っていた(A-M4)。呼び出し元が1回だけ消費して差し引く。
' ----------------------------------------------------------------------------
Public Function ConsumeOnePassTurn() As Boolean
    ConsumeOnePassTurn = mWasOnePass
    mWasOnePass = False
End Function

' ----------------------------------------------------------------------------
' TryOnePass - 1回読みの主処理。
' なぜそうするか:
'   上位検索結果から章を特定し、章丸ごとの本文(引けなければ近傍展開)を
'   1つのメガコンテキストに束ねて回答を生成する。失敗時は即座に False で
'   戻り、従来の4段フローへ退化させる。
' ----------------------------------------------------------------------------
Public Function TryOnePass(ByVal q As String, ByRef hits() As Hit, ByVal nHits As Long, _
                           ByVal history As String, ByVal prevU As String, ByVal prevA As String, _
                           ByRef outBody As String) As Boolean
    outBody = ""
    mWasOnePass = False
    If nHits < 1 Then Exit Function

    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo CatchErr

    Dim maxChars As Long
    maxChars = modConfig.GetLong("onepass_max_chars", 450000)   ' R40 F5: 30万トークン≒45万字
    If maxChars < PER_CHAPTER_MIN_CAP Then maxChars = PER_CHAPTER_MIN_CAP

    Dim maxChapters As Long
    maxChapters = modConfig.GetLong("onepass_max_chapters", 8)
    If maxChapters < 1 Then maxChapters = 1

    Dim seedHitsLimit As Long
    seedHitsLimit = modConfig.GetLong("onepass_seed_hits", 6)
    If seedHitsLimit < 1 Then seedHitsLimit = 1
    ' R38 Fix F2: 章の候補にする上位件数は章数上限より小さくしない
    ' (seed 6 だと maxChapters=8 に届かず章数上限が到達不能だったため。A-M2)。
    If seedHitsLimit < maxChapters Then seedHitsLimit = maxChapters

    ' 1) 章の解決
    Dim mIds() As String, mPaths() As String, mRefs() As String
    Dim metaN As Long
    metaN = modChunkMetaStore.ReadAllMeta(mIds, mPaths, mRefs)

    Dim picksStr As String
    Dim nPick As Long
    Dim picks() As String

    If metaN > 0 Then
        Dim seedN As Long: seedN = nHits
        If seedN > seedHitsLimit Then seedN = seedHitsLimit

        Dim rawKeysBox As String
        Dim i As Long
        For i = 1 To seedN
            Dim cId As String: cId = hits(i).chunk_id
            Dim chapKey As String
            chapKey = modXDoc.ChapterOfChunk(cId, mIds, mPaths, metaN)
            If Len(Trim$(chapKey)) > 0 Then
                Dim srcName As String: srcName = Trim$(hits(i).source)
                If Len(srcName) > 0 Then
                    rawKeysBox = rawKeysBox & srcName & "::" & Trim$(chapKey) & vbLf
                End If
            End If
        Next i

        picksStr = DedupeKeys(rawKeysBox, maxChapters)
        nPick = CountLines(picksStr)
    End If

    ' 2) コンテキストヒットの収集(章丸ごと または 近傍代替)
    Dim ctxHits() As Hit
    Dim ctxN As Long
    Dim modeStr As String

    If nPick > 0 Then
        modeStr = "chapter"
        Dim splitLines() As String
        splitLines = Split(picksStr, vbLf)
        ReDim picks(1 To nPick)
        For i = 1 To nPick
            picks(i) = splitLines(i - 1)
        Next i

        Dim capPer As Long
        ' R38 Fix2(2周目 MAJOR-2): 章には予算の8割だけ渡し、残り2割を F1 で足す
        ' 元 hits のために空けておく。等分し切ると大きい規程(8章×37,500字)で
        ' 章本文だけが上限に達し、追加分が1件も載らない(extra=N は「足した数」
        ' であって「載った数」ではない)。
        capPer = PerChapterCap(maxChars - (maxChars \ 5), nPick, PER_CHAPTER_MIN_CAP)

        ' 章は【1章ずつ順位順】に読む(CollectChapterHits はシートを1回走査して
        ' 行順に返すので、まとめて渡すと章の順が取込順になり、最も関連の強い章が
        ' 真ん中へ埋もれる=Lost-in-the-middle)。picks は動的配列で渡す(固定長
        ' 配列を配列引数へ渡す型を避ける・SortPagesStable と同じ作法)。
        Dim pIdx As Long
        Dim singlePick() As String
        For pIdx = 1 To nPick
            ReDim singlePick(1 To 1)
            singlePick(1) = picks(pIdx)
            Dim chHits() As Hit
            Dim chN As Long
            chN = modAskGlobal.CollectChapterHits(singlePick, 1, chHits, capPer, MAX_CHAPTER_CHUNKS)
            If chN > 0 Then
                Dim oldCtxN As Long: oldCtxN = ctxN
                ctxN = ctxN + chN
                If oldCtxN = 0 Then
                    ReDim ctxHits(1 To ctxN)
                Else
                    ReDim Preserve ctxHits(1 To ctxN)
                End If
                Dim k As Long
                For k = 1 To chN
                    ctxHits(oldCtxN + k) = chHits(k)
                Next k
            End If
        Next pIdx
    End If

    ' 章が引けないかチャンクが0件なら近傍経路へ退化
    If ctxN < 1 Then
        modeStr = "neighbor"
        nPick = 0
        Dim opHits() As Hit
        opHits = hits
        Dim nUse As Long: nUse = nHits
        Dim neighborRad As Long
        neighborRad = modConfig.GetLong("onepass_neighbor", 6)
        modAskFocus.NeighborExpand opHits, nUse, neighborRad
        modAskFocus.RefsExpand opHits, nUse, 0

        ctxN = nUse
        ReDim ctxHits(1 To ctxN)
        For i = 1 To ctxN
            ctxHits(i) = opHits(i)
        Next i
    End If

    If ctxN < 1 Then
        GoTo FallbackZero
    End If

    ' R38 Fix F1: 章モードで、元のヒット(1..nHits)のうち ctxHits に1件も
    ' 含まれていないものを末尾へ追加する。検索7位以降の資料が章に含まれない
    ' 場合、章丸ごと読みでは一度もプロンプトに載らず精度が後退するため
    ' (レビューA-M1)。近傍モードは NeighborExpand/RefsExpand が元の hits を
    ' 既に土台にしているので対象外。
    Dim extra As Long: extra = 0
    If modeStr = "chapter" Then
        Dim addN As Long: addN = 0
        Dim toAdd() As Hit
        Dim j As Long
        For i = 1 To nHits
            Dim already As Boolean: already = False
            For j = 1 To ctxN
                If StrComp(Trim$(ctxHits(j).chunk_id), Trim$(hits(i).chunk_id), vbBinaryCompare) = 0 Then
                    already = True
                    Exit For
                End If
            Next j
            If Not already Then
                addN = addN + 1
                If addN = 1 Then
                    ReDim toAdd(1 To 1)
                Else
                    ReDim Preserve toAdd(1 To addN)
                End If
                toAdd(addN) = hits(i)
            End If
        Next i
        If addN > 0 Then
            Dim baseN As Long: baseN = ctxN
            ctxN = ctxN + addN
            ReDim Preserve ctxHits(1 To ctxN)
            For i = 1 To addN
                ctxHits(baseN + i) = toAdd(i)
            Next i
            extra = addN
        End If
    End If

    ' 3) 本文ブロックと重点抜粋ブロックの構築
    Dim ctxBlock As String
    ctxBlock = SourceBlockOf(ctxHits, ctxN, maxChars)

    Dim focusBlock As String
    Dim focusLimit As Long: focusLimit = nHits
    If focusLimit > MAX_FOCUS_HITS Then focusLimit = MAX_FOCUS_HITS
    Dim fHits() As Hit
    ReDim fHits(1 To focusLimit)
    For i = 1 To focusLimit
        fHits(i) = hits(i)
    Next i
    focusBlock = SourceBlockOf(fHits, focusLimit, maxChars)

    ' 4) 実況表示
    ' R38 Fix F5: neighbor モードで「0章」と出ると利用者に意味が伝わらない
    ' ため(A-m1)、モードごとに文言を分ける。
    Dim stageMsg As String
    Dim searchLensEmoji As String
    searchLensEmoji = ChrW(&HD83D) & ChrW(&HDD0E)
    If modeStr = "neighbor" Then
        stageMsg = searchLensEmoji & " 入念(1回読み): 当たった箇所の前後 " & CStr(ctxN) & _
                   "箇所を読んで回答を作成中" & ChrW(&H2026) & " ※応答なし表示でも処理中"
    Else
        stageMsg = searchLensEmoji & " 入念(1回読み): " & CStr(nPick) & "章・" & _
                   CStr(ctxN) & "箇所を読んで回答を作成中" & ChrW(&H2026) & " ※応答なし表示でも処理中"
    End If
    On Error Resume Next
    modUIMain.SetStage stageMsg
    On Error GoTo CatchErr

    ' 5) プロンプト構築と LLM 呼び出し
    Dim lang As String: lang = modConfig.GetString("answer_language", "日本語")
    Dim ansTags As Boolean: ansTags = modConfig.GetBool("answer_tags", False)
    Dim strictG As Boolean: strictG = modConfig.GetBool("strict_grounding", False)
    Dim prompt As String
    prompt = BuildOnePassPrompt(q, ctxBlock, focusBlock, history, lang, ansTags, strictG)

    Dim effortStr As String
    effortStr = modConfig.GetString("thorough_onepass_effort", "high")
    Dim verbosityStr As String
    verbosityStr = modConfig.GetString("thorough_draft_verbosity", "high")
    Dim modelStr As String
    modelStr = modConfig.GetString("recommended_model", "gpt-5.5")

    Dim latMs As Long
    Dim resp As String
    resp = modGateway.CallLLM(prompt, "thorough_onepass", effortStr, verbosityStr, _
                              modelStr, latMs, prevU, prevA)

    If modRagParse.IsErrorResponse(resp) Then
        ' R38 Fix F3: LLM呼び出しのエラー(ESC中断・NW断)は4段へ落とさず、
        ' そのままエラー表示として返す(4段と同じ「エラーをそのまま表示」に
        ' 揃える。エラーで改めて数分の4段を回さない。レビューA-M3)。
        On Error Resume Next
        modLog.LogUsage "onepass_llm_err", "thorough", "why=llm_err detail=" & modUtil.SafeLeft(resp, 120)
        On Error GoTo 0
        outBody = resp
        mWasOnePass = True
        TryOnePass = True
        Exit Function
    End If

    ' 6) 後処理(タグ適用・出典突合)
    ' 重点抜粋の出典も照合できるよう、ctxHits の後ろに元の hits を連結した全母集団を組む
    Dim allN As Long: allN = ctxN + nHits
    Dim allHits() As Hit
    ReDim allHits(1 To allN)
    For i = 1 To ctxN
        allHits(i) = ctxHits(i)
    Next i
    For i = 1 To nHits
        allHits(ctxN + i) = hits(i)
    Next i

    Dim rawBody As String
    rawBody = modAsk.ApplyAnswerTags(resp)

    outBody = modAskThorough.AnnotateAgainstHits(rawBody, allHits, allN)
    mWasOnePass = True
    TryOnePass = True

    ' 7) 利用ログ記録
    On Error Resume Next
    Dim logDetail As String
    logDetail = "mode=" & modeStr & " chapters=" & CStr(nPick) & _
                " chunks=" & CStr(ctxN) & " chars=" & CStr(Len(prompt)) & _
                " extra=" & CStr(extra)
    ' 秒数は書かない(段ごとの所要時間は CallLLM が ask_steps へ積む・R13 F8)。
    modLog.LogUsage "thorough_onepass", "thorough", logDetail, 0, ctxN
    On Error GoTo 0
    Exit Function

FallbackZero:
    On Error Resume Next
    modLog.LogUsage "onepass_fallback", "thorough", "why=no_chunks"
    On Error GoTo 0
    Exit Function

CatchErr:
    ' ハンドラ稼働中は On Error Resume Next が効かないので、Err を退避してから
    ' Resume で抜ける(modXDoc.Expand の Quiet と同じ作法・CLAUDE.md §11)。
    errNum = Err.Number
    errDesc = Err.Description
    Resume CatchDone
CatchDone:
    On Error Resume Next
    modLog.LogUsage "onepass_fallback", "thorough", "why=err num=" & CStr(errNum) & " desc=" & modUtil.SafeLeft(errDesc, 120)
    Err.Clear
    On Error GoTo 0
    TryOnePass = False
    ' R38 Fix F3: LLM 呼び出しの内側の ESC は modGateway が #ERR 文字列に
    ' 変えるので F3 の枝で返る。ここへ来る 18 はシート走査中の中断だけ。
    If errNum = 18 Then Err.Raise 18
End Function

' ----------------------------------------------------------------------------
' SourceBlockOf - ヒット配列から本文ブロック文字列を組む。
' なぜそうするか:
'   各チャンクの出典タグと全文を連結し、予算上限(BudgetTake)を超えた時点で
'   切り詰めて「(一部省略)」を付与する。
' ----------------------------------------------------------------------------
Private Function SourceBlockOf(ByRef srcHits() As Hit, ByVal n As Long, ByVal maxLen As Long) As String
    If n < 1 Then Exit Function

    Dim sb As String
    Dim usedLen As Long
    Dim i As Long

    For i = 1 To n
        Dim srcTag As String
        srcTag = modPrompts.SourceTag(srcHits(i))
        Dim itemText As String
        itemText = srcTag & vbLf & srcHits(i).full_text & vbLf & vbLf
        Dim itemLen As Long: itemLen = Len(itemText)
        Dim takeLen As Long
        takeLen = modOutlineBuild.BudgetTake(usedLen, itemLen, maxLen)
        If takeLen < 1 Then
            sb = sb & "(一部省略)" & vbLf
            Exit For
        End If
        If takeLen < itemLen Then
            ' R38 Fix F8: 残量が出典タグ+αにも満たないなら、タグの途中で
            ' 切れた文字列をモデルに見せず「(一部省略)」だけで打ち切る
            ' (レビューA-m6)。
            If takeLen < Len(srcTag) + 10 Then
                sb = sb & "(一部省略)" & vbLf
                Exit For
            End If
            itemText = modUtil.SafeLeft(itemText, takeLen)
            sb = sb & itemText & vbLf & "(一部省略)" & vbLf
            usedLen = usedLen + takeLen
            Exit For
        End If
        sb = sb & itemText
        usedLen = usedLen + itemLen
    Next i

    SourceBlockOf = sb
End Function

' ============================================================================
' 純ロジック関数群(Pure43 で直接回帰テスト可能・Hit型非依存)
' ============================================================================

' ----------------------------------------------------------------------------
' IsOn - 文字列の真偽判定。
' なぜそうするか:
'   config値が "on"/"true"/"1"/"yes" のとき有効、"off"/"false"/"0"/"no"
'   のとき無効と判定し、それ以外や空値は規定側(True)に倒す安全設計とする。
' ----------------------------------------------------------------------------
Public Function IsOn(ByVal s As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(s))
    If Len(t) = 0 Then
        IsOn = True
        Exit Function
    End If
    Select Case t
        Case "off", "false", "0", "no"
            IsOn = False
        Case "on", "true", "1", "yes"
            IsOn = True
        Case Else
            IsOn = True
    End Select
End Function

' ----------------------------------------------------------------------------
' DedupeKeys - vbLf 区切り文字列の順序保持・重複除外・件数制限。
' なぜそうするか:
'   上位ヒットから抽出された「資料名::章キー」の候補から、出現順序を維持し
'   大文字小文字・前後空白を無視して重複を排除し、指定件数以内で確定させる。
' ----------------------------------------------------------------------------
Public Function DedupeKeys(ByVal boxLf As String, ByVal maxN As Long) As String
    If maxN < 1 Then Exit Function
    If Len(Trim$(boxLf)) = 0 Then Exit Function

    Dim lines() As String
    lines = Split(boxLf, vbLf)

    Dim kept() As String
    Dim keptN As Long: keptN = 0

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineRaw As String
        lineRaw = Trim$(lines(i))
        If Len(lineRaw) > 0 Then
            Dim dup As Boolean: dup = False
            Dim j As Long
            If keptN > 0 Then
                For j = 1 To keptN
                    If StrComp(kept(j), lineRaw, vbTextCompare) = 0 Then
                        dup = True
                        Exit For
                    End If
                Next j
            End If
            If Not dup Then
                keptN = keptN + 1
                If keptN = 1 Then
                    ReDim kept(1 To 1)
                Else
                    ReDim Preserve kept(1 To keptN)
                End If
                kept(keptN) = lineRaw
                If keptN >= maxN Then Exit For
            End If
        End If
    Next i

    If keptN < 1 Then Exit Function

    Dim res As String
    res = kept(1)
    For i = 2 To keptN
        res = res & vbLf & kept(i)
    Next i
    DedupeKeys = res
End Function

' ----------------------------------------------------------------------------
' PerChapterCap - 章ごとの文字数予算の均等配分。
' なぜそうするか:
'   全体の文字上限を章数で等分し、極端な切り捨てを防ぐため最低保証額
'   (minCap)を下回らないように保護する。
' ----------------------------------------------------------------------------
Public Function PerChapterCap(ByVal total As Long, ByVal nChap As Long, ByVal minCap As Long) As Long
    If total < 1 Then
        PerChapterCap = minCap
        Exit Function
    End If
    If nChap < 1 Then
        PerChapterCap = total
        Exit Function
    End If
    Dim cap As Long
    cap = total \ nChap
    If cap < minCap Then cap = minCap
    PerChapterCap = cap
End Function

' ----------------------------------------------------------------------------
' CountLines - vbLf 区切り文字列の非空行数を数える。
' なぜそうするか:
'   DedupeKeys の結果行数を取得し、章数判定に用いるため。
' ----------------------------------------------------------------------------
Public Function CountLines(ByVal boxLf As String) As Long
    If Len(Trim$(boxLf)) = 0 Then
        CountLines = 0
        Exit Function
    End If
    Dim lines() As String
    lines = Split(boxLf, vbLf)
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            cnt = cnt + 1
        End If
    Next i
    CountLines = cnt
End Function

' ----------------------------------------------------------------------------
' BuildOnePassPrompt - 1回読み用の統合プロンプト文を構築する。
' なぜそうするか:
'   役割指示、思考手順(ansTags時)、厳格な出力規則、免責事項への言及、
'   会話履歴、本棚抜粋本文、重点抜粋、質問、深掘り候補指示を
'   固定された決定論的な順序で連結する。
' ----------------------------------------------------------------------------
Public Function BuildOnePassPrompt(ByVal q As String, ByVal ctxBlock As String, ByVal focusBlock As String, _
                                   ByVal history As String, ByVal lang As String, _
                                   ByVal ansTags As Boolean, ByVal strictG As Boolean) As String
    Dim sb As String

    ' 1. 役割
    sb = "あなたは社内資料の調査アシスタントです。『入念に調べる』モードです。" & _
         "下の本棚抜粋（関連する章の本文・文書順）だけを根拠に、質問へ" & lang & _
         "で丁寧に、根拠を示しながら回答してください。" & vbLf

    ' 2. 手順(ansTags のときだけ)
    If ansTags Then
        sb = sb & "出力は<thinking>に検討、<answer>に利用者へ見せる回答、の構造にすること。" & _
             "<thinking>では次の順で検討する: (1)質問の前提と論点を分解する " & _
             "(2)本文から関連する記述を漏れなく抜き出す（章をまたぐもの・表の見出しと値・但し書きを含む） " & _
             "(3)抜き出した記述同士の矛盾・例外・免責を確認する " & _
             "(4)本文に書かれていない推測を削る " & _
             "(5)各主張に出典を付けられるか確認する。" & _
             "[[FOLLOWUP:...]]の行は</answer>を閉じた後（タグの外）に置くこと。" & vbLf
    End If

    ' 3. 出力規則
    sb = sb & "・根拠にした箇所には必ず出典タグを文の直後に付ける（本棚の資料は [本棚:ファイル名 p.ページ番号]、受け取ったパック由来は [パック(作成者名):ファイル名]）。" & vbLf
    ' R40 F3 / R41 §1 A: Excel 由来の抜粋は行頭に [A6] のような番地が付いている。
    sb = sb & "・出典タグは抜粋の形（Excel はシートN）をそのまま写す。抜粋の行頭にある [A6] のようなセル番地を、根拠の文に「(A6 付近)」のように添える。" & vbLf
    ' R38 Fix F7: 同じ文言が modAskThorough と2箇所に分裂していたため、
    ' 単一情報源(modAskThorough.ThoroughStyleAddendum/ExemptionCoverageAddendum)
    ' を通す(先頭の■を落として・にする=出力文字列は1字も変わらない。A-m3)。
    ' R38 Fix2(2周目 MAJOR-1): ■ は落とさない。表示側(modLive.StyleAnswerParas/
    ' AnswerParagraphs)は行頭 ■ の段落だけを見出し扱いにし、同じプロンプトが
    ' Markdown の # を禁じているので、■ を消すと見出しの指示が1つも残らない。
    sb = sb & "・" & modAskThorough.ThoroughStyleAddendum() & vbLf
    sb = sb & "・" & Mid$(modAskThorough.ExemptionCoverageAddendum(), 2) & vbLf
    sb = sb & "・Markdown記号(#、**、表)は使わない(この画面では崩れて見える)。" & vbLf

    If strictG Then
        sb = sb & "【厳守】本棚抜粋に書かれた情報のみで回答し、外部知識や推測での補完は禁止。出典を付けられない主張は書かない。抜粋から判断できない場合は「資料からは判断できません」とだけ述べ、どんな資料を追加すれば答えられるかを1行添えること。" & vbLf
    End If

    ' 4. 会話履歴(非空時)
    If Len(Trim$(history)) > 0 Then
        sb = sb & vbLf & "## これまでの会話(参考。続きの質問なら踏まえて回答する)" & vbLf & history & vbLf
    End If

    ' 5. 本棚抜粋本文
    sb = sb & vbLf & "## 本棚抜粋(関連する章の本文・文書順)" & vbLf & ctxBlock

    ' 6. 重点抜粋(非空時)
    If Len(Trim$(focusBlock)) > 0 Then
        sb = sb & vbLf & "## 重点抜粋(検索で最も近かった箇所。上の本文と重複あり)" & vbLf & focusBlock
    End If

    ' 7. 質問
    sb = sb & vbLf & "## 質問" & vbLf & q & vbLf

    ' 8. 深掘り候補
    sb = sb & vbLf & "最後に、この回答をさらに深掘りするための質問候補を2つ考え、回答本文の一番最後に次の1行だけを追加してください(この行は利用者向け表示からは自動的に取り除かれます。良い候補が無ければ [[FOLLOWUP: なし]] と書く):" & _
         vbLf & "[[FOLLOWUP: 候補1 | 候補2]]"

    BuildOnePassPrompt = sb
End Function
