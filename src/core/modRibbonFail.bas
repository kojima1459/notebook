Attribute VB_Name = "modRibbonFail"
Option Explicit

' ============================================================================
' modRibbonFail - 社内AIリボンが返す「失敗の語彙」の単一情報源(R48)
' ----------------------------------------------------------------------------
' なぜ要るのか
' ----------------------------------------------------------------------------
' リボンちゃんのアドインは、通信に失敗しても【例外を出さない】。関数の戻り値
' として決まった文字列を返す。ver202609(UpDate=20260916_R)の log.bas を読んで
' 全数を確認した:
'
'   "(error:<HTTPステータス>)<メッセージ>"        … HTTP 200 以外
'   "接続切れ"                                    … status 12031
'   "content_filterに該当しました"                … HTTP 200 で返る
'   "レスポンスから当該テキストを抽出できません…"  … HTTP 200 でパース失敗
'   "レスポンスのJSON文字列が途中で終了しています。" … HTTP 200 でJSON途中終端
'
' R47 まで、これらは modGateway.CallLLM のどの分岐にも当たらず素通しして
' いた。素通しした文字列は modRagParse.ExtractAnswer がタグ無し応答として
' 全文を answer に採るため、【利用者には「AIの回答」として、確信度バッジと
' 出典つきで表示されていた】。憲章 §4-1(無言の失敗禁止)と §3-3(失敗は
' 日本語で「何が起きたか+次の一手」)の二重違反であり、本プロダクトの
' 価値基準(出力の品質・誠実さ)に正面から反する。
'
' なぜ modGateway ではなくここなのか
' ----------------------------------------------------------------------------
' modGateway は 29,875字(残125字)で、CLAUDE.md §12 の「残300字未満は分割裁定
' 必須」に該当する。1〜2行押し込むのは継ぎ足しで、次に触るときに詰む。
' また既存の LooksLikeLimitError / LooksLikeRealAnswer も「リボン・LLM が返す
' 失敗の語彙を判定する」という同じ関心事なので、憲章 §4-5(同型の問題は共通
' 部品で一度だけ)に従ってここへ集約した。前例は R47 の modEmj。
'
' 部分一致(InStr)を使わない理由
' ----------------------------------------------------------------------------
' このリポジトリは部分一致の誤爆を R21 D2 と R30 W2-2 で【2回】やっている
' (LooksLikeLimitError の注記が実例を残している)。"接続切れ" は4字で、通信
' 障害の手順書を引いた回の【正当な回答】に literal で現れうる。3度目を作らない。
' 判定は (a) Trim$ 後の完全一致 (b) 先頭一致+数字の確認 の2種類だけに限る。
' そして全判定の手前で LooksLikeRealAnswer を通し、出典タグ・構造タグを含む
' 応答は無条件に実回答として除外する。
'
' 機械で守る
' ----------------------------------------------------------------------------
' 判定は全て純関数なので modTestsPure47 が対で固定する。誤検知の反例
' (「接続切れの場合は…[本棚: 手順書.pdf p.3]」が失敗ではないこと)も同時に固定する。
' ============================================================================

' リボンが返す失敗文字列(ver202609 実測)。**完全一致で使う**。
Private Const RF_CONN As String = "接続切れ"
Private Const RF_FILTER As String = "content_filterに該当しました"
Private Const RF_JSONCUT As String = "レスポンスのJSON文字列が途中で終了しています。"
' パース失敗だけは後続の案内文が続くので先頭一致で見る。
Private Const RF_PARSE_HEAD As String = "レスポンスから当該テキストを抽出できません"
Private Const RF_HTTP_HEAD As String = "(error:"

' ----------------------------------------------------------------------------
' FailKind - 応答が「リボンの失敗文字列」かを判定し、種別を返す(純関数)。
'   戻り値: "" = 失敗ではない / "limit" / "http" / "conn" / "filter" / "parse"
' ----------------------------------------------------------------------------
Public Function FailKind(ByVal response As String) As String
    If LenB(response) = 0 Then Exit Function
    ' 出典タグ・構造タグを含む応答は、モデルが答えを組み立てた証拠。先に除外する。
    If LooksLikeRealAnswer(response) Then Exit Function

    If Left$(response, Len(RF_HTTP_HEAD)) = RF_HTTP_HEAD Then
        Dim st As Long: st = HttpStatusOf(response)
        ' 429 は「呼びすぎ」なので既存の利用上限(E0204)へ寄せる。
        ' modEnrich の取込打ち切りが E0204 系の判定に乗っているため、
        ' ここを http へ倒すと取込の安全装置が静かに効かなくなる。
        If st = 429 Then
            FailKind = "limit"
            Exit Function
        ElseIf st > 0 Then
            FailKind = "http"
            Exit Function
        End If
        ' 形が崩れていたら下の完全一致へ落とす(素通しはしない)。
    End If

    Dim t As String: t = Trim$(response)
    If t = RF_CONN Then
        FailKind = "conn"
        Exit Function
    End If
    If t = RF_FILTER Then
        FailKind = "filter"
        Exit Function
    End If
    If t = RF_JSONCUT Then
        FailKind = "parse"
        Exit Function
    End If
    If Left$(t, Len(RF_PARSE_HEAD)) = RF_PARSE_HEAD Then
        FailKind = "parse"
        Exit Function
    End If

    If LooksLikeLimitError(response) Then FailKind = "limit"
End Function

' ----------------------------------------------------------------------------
' CodeFor - 種別 → エラーコード(純関数)。
' ----------------------------------------------------------------------------
' E0202(混雑・操作の重なり)へ寄せない ―― 主因が違う。
' E0204(本日の上限。明日「同期」で再開)へ 401/403/500 を寄せない ―― 嘘になる。
' content_filter を通信失敗と混ぜない ―― 利用者がとるべき行動が違う
' (通信は「待つ」、フィルタは「表現を変える」)。
Public Function CodeFor(ByVal kind As String) As String
    Select Case kind
        Case "limit"
            CodeFor = "E0204"
        Case "filter"
            CodeFor = "E0206"
        Case "http", "conn", "parse"
            CodeFor = "E0205"
    End Select
End Function

' ----------------------------------------------------------------------------
' ErrFor - CallLLM から1行で呼ぶ窓口。
'   戻り値: "" = 失敗ではない(呼び出し元はそのまま応答として使う)
'           "#ERR:E02xx:<生応答>" = 失敗(呼び出し元は Exit する)
' ----------------------------------------------------------------------------
' 生応答は modRagParse.BuildErrorAnswer が捨てる(利用者には FriendlyMessage
' だけが出る)ので、ここで付けておくのは err_log と診断のため。
' 憲章 §4-2(観測性ファースト): リボンの文言が1文字変わって完全一致が外れても、
' err_log に生応答が残っていれば「なぜ効かなくなったか」が後から分かる。
Public Function ErrFor(ByVal response As String, ByVal step_name As String) As String
    Dim kind As String: kind = FailKind(response)
    If LenB(kind) = 0 Then Exit Function

    Dim code As String: code = CodeFor(kind)
    If LenB(code) = 0 Then Exit Function

    modLog.LogError code, "modGateway.CallLLM", _
        "kind=" & kind & " step=" & step_name & " resp=" & modUtil.SafeLeft(response, 200)
    ErrFor = "#ERR:" & code & ":" & response
End Function

' ----------------------------------------------------------------------------
' IsLimitErr - 「利用上限らしい」かを、#ERR 化の前後どちらでも判定する。
' ----------------------------------------------------------------------------
' modEnrich は CallLLM の戻り値へ LooksLikeLimitError を直接当てていたが、
' Azure の 429 本文は英語で 120字を超えるため Len>120 の早期 return に当たり、
' 【取込の打ち切りが一度も効いていなかった】。CallLLM が E0204 を付けるように
' なったので、コード側で判定できるようにする。
Public Function IsLimitErr(ByVal s As String) As Boolean
    If Left$(s, 11) = "#ERR:E0204:" Then
        IsLimitErr = True
        Exit Function
    End If
    IsLimitErr = LooksLikeLimitError(s)
End Function

' ----------------------------------------------------------------------------
' IsCancelled - 「利用者が ESC で止めた」応答か(純関数)。
' ----------------------------------------------------------------------------
' R48 Fix2(敵対的レビュー R48-REV-01 の2周目): E0207 を作っただけでは
' 「止めるときは ESC キー」の約束は果たせない。多段の経路は #ERR を受けても
' 【次の段へ進む】ように作られているため、ESC は1回の呼び出しぶんしか止まらず、
' 論点分解(最大11段)では論点の数だけ押す必要があった。
' 段をまたいで畳むために、各段の合流点でこれを見る。
Public Function IsCancelled(ByVal s As String) As Boolean
    IsCancelled = (Left$(s, 11) = "#ERR:E0207:")
End Function

' ----------------------------------------------------------------------------
' HttpStatusOf - "(error:429)…" → 429。形が違えば 0(純関数)。
' ----------------------------------------------------------------------------
' IsNumeric は "1e5" や "-1" や " 1 " を通すので使わない。
Public Function HttpStatusOf(ByVal response As String) As Long
    If Left$(response, Len(RF_HTTP_HEAD)) <> RF_HTTP_HEAD Then Exit Function
    Dim p As Long: p = InStr(Len(RF_HTTP_HEAD) + 1, response, ")")
    If p <= Len(RF_HTTP_HEAD) Then Exit Function

    Dim d As String: d = Mid$(response, Len(RF_HTTP_HEAD) + 1, p - Len(RF_HTTP_HEAD) - 1)
    If Len(d) < 1 Then Exit Function
    If Len(d) > 5 Then Exit Function

    Dim i As Long
    For i = 1 To Len(d)
        Dim c As String: c = Mid$(d, i, 1)
        If c < "0" Then Exit Function
        If c > "9" Then Exit Function
    Next i

    HttpStatusOf = CLng(d)
End Function

' ----------------------------------------------------------------------------
' LooksLikeLimitError - R48 で modGateway から【ロジックを1文字も変えずに】移設。
' ----------------------------------------------------------------------------
' LLMの応答文字列が「利用上限に達した」という定型拒否メッセージそのものらしいかを
' 検知する。実機報告(2026-07-22)「しっかり調べるモードだけ必ずE0204になる」で
' 確認された誤検知バグ: 保険約款は「上限」「回数」(支払限度額・請求回数等)を
' ごく普通に含むため、深掘りモードの長文で正当な分析結果がほぼ確実に誤爆して
' いた。定型拒否文は短いので、応答が短い場合に限って判定する(長い実回答は
' 対象外)。
Public Function LooksLikeLimitError(ByVal response As String) As Boolean
    If Len(response) > 120 Then Exit Function

    ' 2026-07-28(レビュー M-1): 120字以下で「上限」「回数」等を含むだけで
    ' 利用上限エラー扱いにしていたため、
    '   「請求回数の上限はありません。[本棚: 約款.pdf p.12]」
    ' のような【正当な短文回答】がまるごとエラーメッセージに差し替わっていた。
    ' 出典タグや構造タグを含む応答は、モデルが実際に答えを返した証拠なので
    ' 上限エラーではありえない。先に除外する。
    If LooksLikeRealAnswer(response) Then Exit Function

    Dim s As String: s = LCase$(response)
    LooksLikeLimitError = (InStr(s, "上限") > 0) Or (InStr(s, "limit") > 0) Or _
                           (InStr(s, "回数") > 0) Or (InStr(s, "rate") > 0) Or _
                           (InStr(s, "quota") > 0)
End Function

' 出典タグ・構造タグを含む=モデルが答えを組み立てている応答か。
' R48: modGateway から移設(ロジックは1文字も変えていない)。
Public Function LooksLikeRealAnswer(ByVal response As String) As Boolean
    If InStr(1, response, "[本棚:", vbTextCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    If InStr(1, response, "[出典", vbTextCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    If InStr(1, response, "<answer>", vbTextCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    If InStr(1, response, "<thinking>", vbTextCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    ' R30 W2-2(実機第15報C班・3件目): 分解段(modPrompts.BuildDecomposePrompt)の
    ' <verdict>single/parts/clarify/global</verdict> も構造タグの一種。短文の
    ' verdict応答が「上限」「回数」を含む論点(保険約款等)に触れただけでE0204
    ' 誤爆していた(R21 D2/R21H F4と同型の症状)。
    If InStr(1, response, "<verdict>", vbTextCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    ' R21-2 D2(実機第8報⑧): 査読(critique)「1. [観点] …」も救済(旧判定は
    ' 短い棄却指摘を誤爆させていた。タグはmodPrompts.BuildCritiquePromptと同一)
    ' R21H F4: 前置き判定が「1. [」の完全一致だけだったため「1.[」「1.  [」等
    ' 空白ゆれで誤爆していた。「1.」+任意空白+「[」へ緩和する。
    Dim t As String: t = Trim$(response)
    If Left$(t, 2) = "1." Then
        Dim p As Long: p = 3
        Do While p <= Len(t) And Mid$(t, p, 1) = " "
            p = p + 1
        Loop
        If Mid$(t, p, 1) = "[" Then LooksLikeRealAnswer = True: Exit Function
    End If
    If InStr(1, response, "[論点漏れ]", vbBinaryCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    If InStr(1, response, "[未検証の断定]", vbBinaryCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    ' R21H F4: 「[出典:」(実回答の出典表記)はあるが「[出典不備]」(査読の
    ' 指摘タグ)は判定漏れで、その指摘だけの棄却応答がE0204(上限超過)に
    ' 誤爆していた。
    If InStr(1, response, "[出典不備]", vbBinaryCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    If InStr(1, response, "[憶測]", vbBinaryCompare) > 0 Then LooksLikeRealAnswer = True
End Function
