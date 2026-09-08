Attribute VB_Name = "modCorrect"
Option Explicit

' ============================================================================
' modCorrect - 是正メモ(R36 §2「❌違う → 次から必ず使われる」)。
' ----------------------------------------------------------------------------
' 【なぜ要るか】R35 までの「❌違う」は、正しい内容を【修正ナレッジ】という
'   普通の資料として1冊増やすだけだった。元の質問文が本文に無いので、同じ
'   質問をもう一度打っても、回答抜粋に偶然含まれる語でしか当たらない。
'   トーストの「次から同じ質問をした人から、この内容で答えます」は、実装が
'   裏付けていない過大表現だった(R36 §2-1)。
'
' 【どう直すか】UI は1つも増やさない。2点だけ:
'   (1) 保存する本文に【元の質問文】を必ず入れる(BuildMemoBody)。資料名は
'       "是正メモ_<質問先頭24字>" に固定し、同じ質問への2回目は上書きする
'       (modAppAct.RecordCorrection が登録前に modShelf.DeleteSource する)。
'   (2) 回答生成の直前、modAskRetrieve が hits を返すその1行手前で、質問が
'       一致する是正メモを Hit として合成し【先頭へ】差し込む(InjectHits)。
'
' 【なぜ Hit の合成で足りるか】modAsk は hits の中身だけを使い my_knowledge を
'   引き直さない(modAsk.bas:296-297, 367-376)。modPrompts.BuildSourceBlock は
'   hits(i).full_text を先頭から順に [本棚:source p.N] タグ付きで載せる
'   (modPrompts.bas:560-617)。つまり Hit を1件先頭へ足すだけで、凍結モジュール
'   (modAsk/modPrompts/modRetrieve/modShelf/modBoot)を1文字も触らずに、本文にも
'   出典チップにも載る。本文冒頭の「【是正メモ】この質問には…で答えること」が
'   そのまま LLM への指示として働く。
'
' 【安全弁】この機能が壊れても質問には必ず答える。InjectHits は全体を
'   On Error のサーキットブレーカーで包み、失敗したら nHits をそのまま返す
'   (CONTRIBUTING §2.4「疎結合プラグイン方式」)。config correct_inject=off で
'   丸ごと止められる。nHits < 0(埋め込み失敗)は触らずそのまま返す。
'
' 【層】qa層。UI は呼ばない(Toast/MsgBox 禁止)。参照するのは core(modAppDef/
'   modConfig/modState/modUtil/modTypes)と同格の modSparse/modInsight、および
'   (R37 §2-1 CollectAnswerSources のみ)凍結 modAsk の公開読み取り専用
'   accessor(LastGenHitCount/LastHitSource/LastHitPage)。modAsk は modCorrect
'   を参照しないので循環は無い。
'   my_knowledge のシート直読みは同じ qa 層の既存前例に倣う
'   (modAskFocus.bas:80 / modAskGlobal.bas:294 / modRetrieve.bas:590)。
'   シート名は modAppDef.SH_KNOWLEDGE 経由で取る(リテラルを書かない)。
' ============================================================================

' 是正メモの資料名(source)の接頭辞。modAppAct と InjectHits の走査が同じ
' 文字列を見る単一情報源。ここを変えると過去の是正メモが拾えなくなる。
' R36 Fix A-M5/A-M8: Public。modAskRetrieve.HitSourceList(会話出典メモリの
' 8枠から除外)と modInsightIo.EmitVerifiedQA(共有本文の資料名の伏せ字)が
' 同じ接頭辞を見る ―― 書き写すと必ずズレる。
Public Const MEMO_PREFIX As String = "是正メモ_"

' 本文の「質問」行。組み立て(BuildMemoBody)は "質問: "、読み取り
' (ExtractQuestionLine)は "質問:" で照合する(全角スペース等のゆれを拾う)。
Private Const Q_HEAD_OUT As String = "質問: "
Private Const Q_HEAD_IN As String = "質問:"

' 本文の「誤答の根拠」行(R37 §2-1)。組み立て(BuildMemoBody)は "誤答の根拠: "、
' 読み取り(ExtractWrongSources/WrongSourceMatches)は "誤答の根拠:" で照合する。
' Q_HEAD_OUT/Q_HEAD_IN と同じ対の作法(片方を変えたら両方直す)。
Private Const WRONG_HEAD_OUT As String = "誤答の根拠: "
Private Const WRONG_HEAD_IN As String = "誤答の根拠:"

' 資料名に使う質問の先頭字数(R36 §2-2-1)。
Private Const DOCBASE_Q_LEN As Long = 24

' 語一致とみなす最低一致率(%)。config correct_key_min の既定値。
' 60% にした根拠: DistinctiveKeys は最大8件しか返さないので、5語の質問なら
' 3語一致で通り2語一致では通らない。実測の妥当性確認は modTestsPure40 の
' KeyMatchPct テスト(100% / 50% / 0% の3点)で固定している。
Private Const DEFAULT_KEY_MIN As Long = 60

' 是正メモの一致度に与えるスコア。低関連度警告は hits(i).score の最大値で
' 判定される(modAskRetrieve.bas:509-511)ので、是正が当たったターンで
' 「関連が薄い」と警告が出ないよう cos の上限側に置く。
Private Const SCORE_EXACT As Double = 1#
Private Const SCORE_KEY As Double = 0.9

' ----------------------------------------------------------------------------
' IsMemoSource - この資料名は是正メモか(R36 Fix A-M5/A-M8)。純関数。
'   接頭辞の照合を1箇所に閉じ込める。逼迫している呼び出し側
'   (modAskRetrieve 残217字 / modInsightIo)が1行で済むようにするためでも
'   ある ―― Left$/Len を各所へ書き写すと、接頭辞を変えたとき必ず取り残す。
' ----------------------------------------------------------------------------
Public Function IsMemoSource(ByVal src As String) As Boolean
    IsMemoSource = (Left$(Trim$(src), Len(MEMO_PREFIX)) = MEMO_PREFIX)
End Function

' ----------------------------------------------------------------------------
' SourceExists - my_knowledge にこの資料名(source)の行があるか(R36 Fix M1)。
'   modAppAct.RecordCorrection が「消す前に、消すものがあったか」を控えるため。
'   同名の旧是正メモを DeleteSource してから登録し、その登録が失敗すると
'   【前の是正メモだけが消えて何も残らない】(不可逆)。せめて何が起きたかを
'   利用者へ正しく言うために、消す前の有無をここで見る。
'   読むのは source 列だけ(1〜2列。1列だけの Range.Value はスカラーになるので
'   2列読む ―― InjectHits と同じ理由)。失敗したら False(=控えめな文言側)。
' ----------------------------------------------------------------------------
Public Function SourceExists(ByVal srcName As String) As Boolean
    Dim target As String: target = LCase$(Trim$(srcName))
    If LenB(target) = 0 Then Exit Function

    On Error GoTo Done

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function

    Dim lastK As Long
    lastK = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastK, 2)).Value

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If LCase$(Trim$(CStr(arr(i, 2)))) = target Then
            SourceExists = True
            Exit Function
        End If
    Next i
    Exit Function

Done:
    SourceExists = False
End Function

' ----------------------------------------------------------------------------
' BuildMemoBody - 是正メモの本文(R36 §2-2-1 の書式)。純関数。
'   modVault.RegisterKnowledgeText が先頭へ "【是正メモ】" & vbLf を足すので、
'   ここが返すのは2行目以降にあたる本体。1行目の指示文は【LLMへの指示】として
'   意図的に本文側にも置く(タイトル行はチャンク先頭の breadcrumb に埋もれる
'   ことがあり、指示が消える経路を作らないため)。
'   R37 §2-1: 第4引数 sourcesLine(空でもよい)が非空のとき、末尾へ
'   「誤答の根拠: <source> p.<page> | …」行を追加する。読み取りは
'   ExtractWrongSources。組み立てるのは modAppAct.RecordCorrection が
'   modCorrect.CollectAnswerSources() で作った値(その質問に答えたときの
'   出典・最大4件)。既存の3引数呼び出しは同一コミットで更新済み。
' ----------------------------------------------------------------------------
Public Function BuildMemoBody(ByVal question As String, ByVal fixText As String, _
                              ByVal answerExcerpt As String, ByVal sourcesLine As String) As String
    BuildMemoBody = "【是正メモ】この質問には、下の「正しい内容」で答えること。" & vbLf & _
                    Q_HEAD_OUT & question & vbLf & _
                    "正しい内容: " & fixText & vbLf & _
                    "(誤りだった回答の抜粋: " & answerExcerpt & ")"
    If LenB(Trim$(sourcesLine)) > 0 Then
        BuildMemoBody = BuildMemoBody & vbLf & WRONG_HEAD_OUT & sourcesLine
    End If
End Function

' ----------------------------------------------------------------------------
' CollectAnswerSources - 直近の回答が根拠にした資料名とページを最大4件、
'   "source p.N | source p.N | …" の1行にまとめる(R37 §2-1)。純関数ではない
'   (modAsk の状態を読む)が、書き込みは一切しない。
'   modAsk は凍結のため、公開済みの読み取り専用 accessor
'   (LastGenHitCount/LastHitSource/LastHitPage。いずれも0始まり)だけを使う。
'   是正メモ自身(source が MEMO_PREFIX で始まる)は除く ―― 「誤答の根拠」が
'   前の是正メモだった場合(前の是正が誤っていた場合)は記録しない、という
'   R37 §2-1 の裁定をここ1箇所で満たす。
'   呼び出しは modAppAct.RecordCorrection の是正メモ経路のみ(❌違うを押した
'   その場で、直前の回答の hits を読む=タイミングが命)。
' ----------------------------------------------------------------------------
Public Function CollectAnswerSources() As String
    Const MAX_SOURCES As Long = 4
    ' 【R37 Fix B-m7】聞き返し(逆質問)の途中で ❌違う を押された回は空を返す。
    '   そのターンの画面に出ているのは回答ではなく質問で、hits は「聞き返しを
    '   組むために引いた候補」でしかない。それを「誤答の根拠」として記録すると、
    '   悪くない資料が次から後ろへ下げられる(是正の誤爆はいちばん高くつく型)。
    On Error Resume Next
    Dim pending As Boolean: pending = modClarify.HasPending()
    Err.Clear
    On Error GoTo 0
    If pending Then Exit Function

    Dim buf As String
    Dim cnt As Long
    Dim n As Long: n = modAsk.LastGenHitCount()

    Dim i As Long
    For i = 0 To n - 1
        Dim src As String: src = Trim$(modAsk.LastHitSource(i))
        If LenB(src) > 0 Then
            If Not IsMemoSource(src) Then
                If cnt > 0 Then buf = buf & " | "
                ' R41 §1 A: Excel由来は「シートN」(modMode.PageTagPart。単一情報源)。
                buf = buf & src & modMode.PageTagPart(src, modAsk.LastHitPage(i))
                cnt = cnt + 1
                If cnt >= MAX_SOURCES Then Exit For
            End If
        End If
    Next i
    CollectAnswerSources = buf
End Function

' ----------------------------------------------------------------------------
' ExtractQuestionLine - 是正メモ本文から「質問:」行の中身を取り出す。純関数。
'   取込後の full_text は先頭に breadcrumb 行【資料名>章>条】が焼かれている
'   (modShelf.bas:615-630)が、行単位で探すので影響を受けない。
'   見つからなければ空文字(=是正メモではない/壊れている → 注入しない)。
' ----------------------------------------------------------------------------
Public Function ExtractQuestionLine(ByVal memoText As String) As String
    If LenB(memoText) = 0 Then Exit Function

    Dim seg() As String
    seg = Split(Replace(memoText, vbCr, ""), vbLf)

    Dim i As Long
    For i = LBound(seg) To UBound(seg)
        Dim ln As String: ln = seg(i)
        If Left$(ln, Len(Q_HEAD_IN)) = Q_HEAD_IN Then
            ExtractQuestionLine = Trim$(Mid$(ln, Len(Q_HEAD_IN) + 1))
            Exit Function
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' ExtractWrongSources - 是正メモ本文から「誤答の根拠:」行をそのまま取り出す。
'   純関数。ExtractQuestionLine とは違い、返すのは行の中身(コロンの後ろ)では
'   なく行全体(ヘッダー込み) ―― WrongSourceMatches がヘッダーの有無どちらでも
'   受け取れるように作ってあるので、ここでは剥がさない(R37 §2-2)。
'   見つからなければ空文字(=旧い是正メモ・sourcesLineが空だった是正メモ)。
' ----------------------------------------------------------------------------
Public Function ExtractWrongSources(ByVal memoText As String) As String
    If LenB(memoText) = 0 Then Exit Function

    Dim seg() As String
    seg = Split(Replace(memoText, vbCr, ""), vbLf)

    Dim i As Long
    For i = LBound(seg) To UBound(seg)
        Dim ln As String: ln = Trim$(seg(i))
        If Left$(ln, Len(WRONG_HEAD_IN)) = WRONG_HEAD_IN Then
            ExtractWrongSources = ln
            Exit Function
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' WrongSourceMatches - 「誤答の根拠:」行(ExtractWrongSources の戻り値、または
'   ヘッダーを含まない中身のどちらでもよい)に、この(source, page)が載っている
'   か。純関数。資料名は大小無視(vbTextCompare)・ページは数値一致。
'   "|" で区切った各項目を末尾から " p." で切り分ける(資料名にスペースが
'   含まれても、ページ番号は必ず項目の末尾にあるため崩れない)。
'   R41 §1 A: Excel由来の新しい是正メモは " シート" 接尾辞で書かれる
'   (modCorrect.CollectAnswerSources)。旧い是正メモは " p." のまま残る
'   (§5 記録のみ・表示は変えない)ので、両方を受ける。両方が項目内に
'   現れることは無いが、念のため後ろの位置(項目末尾に近い方)を採る。
' ----------------------------------------------------------------------------
Public Function WrongSourceMatches(ByVal line As String, ByVal source As String, _
                                   ByVal page As Long) As Boolean
    If LenB(Trim$(line)) = 0 Then Exit Function
    If LenB(Trim$(source)) = 0 Then Exit Function

    Dim content As String: content = Trim$(line)
    If Left$(content, Len(WRONG_HEAD_IN)) = WRONG_HEAD_IN Then
        content = Trim$(Mid$(content, Len(WRONG_HEAD_IN) + 1))
    End If
    If LenB(content) = 0 Then Exit Function

    Dim srcNorm As String: srcNorm = Trim$(source)
    Dim parts() As String: parts = Split(content, "|")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim ent As String: ent = Trim$(parts(i))
        If LenB(ent) > 0 Then
            ' LO は4引数形(Start=-1+Compare)で不一致を返す。" p."/" シート" は
            ' 大小無関係なので2引数形(CLAUDE.md §10)。
            Dim pPosP As Long: pPosP = InStrRev(ent, " p.")
            Dim pPosS As Long: pPosS = InStrRev(ent, " シート")
            Dim pPos As Long, sfxLen As Long
            If pPosP > 0 And pPosS > 0 Then
                If pPosS > pPosP Then
                    pPos = pPosS: sfxLen = 4   ' Len(" シート")
                Else
                    pPos = pPosP: sfxLen = 3   ' Len(" p.")
                End If
            ElseIf pPosS > 0 Then
                pPos = pPosS: sfxLen = 4
            ElseIf pPosP > 0 Then
                pPos = pPosP: sfxLen = 3
            End If
            If pPos > 0 Then
                Dim entSrc As String: entSrc = Trim$(Left$(ent, pPos - 1))
                Dim pageStr As String: pageStr = Mid$(ent, pPos + sfxLen)
                If IsNumeric(pageStr) Then
                    If StrComp(entSrc, srcNorm, vbTextCompare) = 0 Then
                        If CLng(pageStr) = page Then
                            WrongSourceMatches = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' MemoDocBase - 是正メモの資料名のもと("是正メモ_" + 質問先頭24字)。純関数。
' ----------------------------------------------------------------------------
' modVault.RegisterKnowledgeText は docBase を SanitizeName に通してから
' TEMP\<それ>.txt へ書き、source 名は拡張子込みのファイル名になる
' (modVault.bas:286-288 → modShelf.bas:72)。RecordCorrection は「同名の
' 既存資料を消してから登録し直す」ために、登録前に source 名を自分で組み立て
' なければならない ―― ところが SanitizeName は modVault の Private で呼べない。
' そこで【SanitizeName と同じ置換をここで先に済ませておく】。こうすると
' SanitizeName を通しても値が変わらない(冪等)ので、
'   source == MemoDocBase(q) & ".txt"
' が必ず成り立ち、消し忘れ=是正メモが増え続ける事故が起きない。
' 置換表は modVault.bas:371-372 と同一。片方を変えたら両方直すこと。
' (R36 §2-3 は「置換は SanitizeName に任せる」と書いているが、それだと
'  「?」を含む質問 ―― 日本語の質問では多数 ―― で source 名が食い違い、
'  §2-3 が同時に要求する DeleteSource の同名判定が空振りする。実装波の
'  判断で冪等化を選び、司令塔へ報告した。)
' R36 Fix M2 / A-M7: 先頭24字だけでは【別の質問が同じ資料名になる】。
' 「育児休業の申請期限について教えてください」と「育児休業の申請期限は
'  いつまでですか」は先頭24字が同じで、後から書いた是正メモが前のものを
' DeleteSource で消してしまう(2つ目の質問の是正が1つ目を殺す)。一方で
' 一致判定は NormKey(先頭40字)なので、物差しが24字と40字で食い違っていた。
' そこで末尾へ4桁の識別子を足す ―― NormKey(q) の全文字の AscW の和を
' 10000 で割った余り。AscW は U+8000 以降を負値で返すので +65536 で補正する
' (CLAUDE.md §10 の確立作法)。同じ質問なら必ず同じ4桁になる=冪等は維持。
Public Function MemoDocBase(ByVal question As String) As String
    Dim t As String: t = modUtil.SafeLeft(question, DOCBASE_Q_LEN)

    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i

    MemoDocBase = MEMO_PREFIX & Trim$(t) & "_" & QuestionTag4(question)
End Function

' 質問の4桁識別子(0000〜9999)。純関数。物差しは完全一致判定と同じ
' modInsight.NormKey(空白・句読点・?・括弧を落として先頭40字を小文字化)に
' 揃える ―― 「?」の有無だけが違う同じ質問で別の4桁になると、2回目の是正が
' 1つ目を上書きできず是正メモが増え続ける。
Public Function QuestionTag4(ByVal question As String) As String
    Dim k As String: k = modInsight.NormKey(question)
    Dim total As Long
    Dim i As Long
    For i = 1 To Len(k)
        Dim c As Long: c = AscW(Mid$(k, i, 1))
        If c < 0 Then c = c + 65536      ' U+8000以降は負値で返る(§10)
        total = (total + c) Mod 10000
    Next i
    QuestionTag4 = Right$("000" & CStr(total), 4)
End Function

' ----------------------------------------------------------------------------
' KeyMatchPct - 質問の「効く語」のうち何%が是正メモの質問に現れるか。純関数。
'   両引数とも modSparse.NormalizeForSearch 済みの文字列を渡すこと。
'   照合先を【是正メモの質問行だけ】に限るのは誤爆対策(R36 §7 リスク台帳)。
'   本文全体(正しい内容+誤答抜粋)を見ると、話題が近いだけの別の質問にも
'   古い是正が先頭に来る。
' ----------------------------------------------------------------------------
Public Function KeyMatchPct(ByVal qNorm As String, ByVal memoQNorm As String) As Long
    Dim keyText As String: keyText = modSparse.DistinctiveKeys(qNorm)
    If LenB(keyText) = 0 Then Exit Function

    ' DistinctiveKeys は空白を落とした上で語を採るので、照合先も落として揃える。
    Dim doc As String: doc = Replace(memoQNorm, " ", "")
    If LenB(doc) = 0 Then Exit Function

    Dim arr() As String: arr = Split(keyText, "|")
    Dim total As Long, okCnt As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Dim k As String: k = arr(i)
        If Len(k) >= 2 Then
            total = total + 1
            If InStr(1, doc, k, vbBinaryCompare) > 0 Then okCnt = okCnt + 1
        End If
    Next i
    ' R36 Fix N6: 効く語が1語しか取れない質問は、その1語が当たるだけで
    ' 一致率100%になる ―― 「有給は?」で登録した是正が「有給の繰越は?」にも
    ' 「有給の申請先は?」にも先頭で刺さる(誤爆の主要経路)。分母が2未満の
    ' ときは語一致を名乗らせず 0 にし、完全一致だけに任せる。
    If total < 2 Then Exit Function

    KeyMatchPct = Int((okCnt * 100#) / total)
End Function

' ----------------------------------------------------------------------------
' NormFull - 完全一致の物差し(R36 Fix A-M6)。純関数。
' ----------------------------------------------------------------------------
'   modInsight.NormKey と【同じ除去規則】(空白・全角空白・、。・半角/全角の
'   ?・全角括弧を落として小文字化)だが、先頭40字のクランプをしない。
'   NormKey をそのまま完全一致に使うと、41字目以降だけが違う別の質問
'   ―― 「育児休業を延長したいのですが手続きの流れを教えてください」と
'   「…手続きに必要な書類を教えてください」のような長い質問 ―― が同一視され、
'   score 1.0(=最強の一致)で先頭に差し込まれる。gap の重複判定は「似た質問を
'   まとめる」のが目的なので40字で正しいが、是正の完全一致は「同じ質問」でしか
'   成立してはならない。物差しが違うので関数を分ける。
'   modInsight(pack層・残2,606字)は触らず、qa層のここに置く(R36 §9-2 A-M6)。
'   除去する文字は modInsight.bas:417-425 と同一。片方を変えたら両方直すこと。
' ----------------------------------------------------------------------------
Public Function NormFull(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, " ", ""): t = Replace(t, ChrW(&H3000), "")
    t = Replace(t, ChrW(&H3001), ""): t = Replace(t, ChrW(&H3002), "")
    t = Replace(t, "?", ""): t = Replace(t, ChrW(&HFF1F), "")
    t = Replace(t, ChrW(&HFF08), ""): t = Replace(t, ChrW(&HFF09), "")
    NormFull = LCase$(t)
End Function

' ----------------------------------------------------------------------------
' MatchLevel - 質問と是正メモの一致度。2=完全一致 / 1=語一致 / 0=不一致。
'   qNorm は modSparse.NormalizeForSearch を通した質問。
'   完全一致の判定は NormFull(NormKey と同じ除去規則・40字クランプ無し)。
' ----------------------------------------------------------------------------
Public Function MatchLevel(ByVal qNorm As String, ByVal memoText As String) As Long
    If LenB(Trim$(qNorm)) = 0 Then Exit Function

    Dim memoQ As String: memoQ = ExtractQuestionLine(memoText)
    If LenB(Trim$(memoQ)) = 0 Then Exit Function

    Dim memoNorm As String: memoNorm = modSparse.NormalizeForSearch(memoQ)

    Dim kq As String: kq = NormFull(qNorm)
    Dim km As String: km = NormFull(memoNorm)
    If LenB(kq) = 0 Then Exit Function
    If LenB(km) = 0 Then Exit Function
    If StrComp(kq, km, vbBinaryCompare) = 0 Then
        MatchLevel = 2
        Exit Function
    End If

    ' R36 Fix A-m4: correct_key_min = 0 は「語一致を使わない(完全一致だけ)」
    ' という意思表示。従来はここで既定60へ黙って戻していたため、0 にした人の
    ' 設定が無視されていた(しかも 0 をそのまま閾値にすると全件一致になる)。
    Dim kmin As Long: kmin = KeyMinPct()
    If kmin < 1 Then Exit Function

    If KeyMatchPct(qNorm, memoNorm) >= kmin Then MatchLevel = 1
End Function

' 語一致の閾値(%)。config が読めない実行環境(LO実行テスト)では既定値のまま
' 走り切る ―― modSparse.EffectiveKeyScoreCap と同型。
' 0 = 語一致を使わない(完全一致だけ)。負値だけを「壊れた設定」として既定へ戻す。
Private Function KeyMinPct() As Long
    Dim v As Long: v = DEFAULT_KEY_MIN
    On Error Resume Next
    v = modConfig.GetLong("correct_key_min", DEFAULT_KEY_MIN)
    On Error GoTo 0
    If v < 0 Then v = DEFAULT_KEY_MIN
    If v > 100 Then v = 100
    KeyMinPct = v
End Function

' ----------------------------------------------------------------------------
' LastQuestionFromState - 直近の質問文。
' ----------------------------------------------------------------------------
' modAsk の mLastQuestion は Private で、modAsk は凍結のため accessor を
' 足せない。代わりに modAppState.SaveTurnForRestore(modAppState.bas:354-364)が
' 毎ターン ui_state の "nexus_hist_u" へ「新しい順・";;;"区切り・先頭300字」で
' 積んでいるので、その先頭セグメントを取る。呼び出しは modApp.OnSend の
' 1箇所(modApp.bas:233)で、4分岐すべての合流点にある。
' ※回答が例外で終わったターン(GoTo Fail)では保存されないため、その直後だけは
'   「1つ前の質問」が返る。ただし ❌違う は回答バブルが描けたターンにしか
'   出ないので、実運用の経路では起きない(R36報告済み)。
Public Function LastQuestionFromState() As String
    Dim s As String
    On Error Resume Next
    s = modState.LoadState("nexus_hist_u", "")
    On Error GoTo 0
    If LenB(s) = 0 Then Exit Function

    Dim p As Long: p = InStr(1, s, ";;;", vbBinaryCompare)
    If p > 0 Then s = Left$(s, p - 1)
    LastQuestionFromState = Trim$(s)
End Function

' ----------------------------------------------------------------------------
' PrependHit - hits の先頭へ1件差し込む(1始まり配列)。
'   同じ source が既にあれば重複させず、その1件を先頭へ移すだけ(件数不変)。
'   無ければ ReDim Preserve で1つ伸ばして先頭へ置く(件数 +1)。
'   戻り値は差し込んだ後の件数。
'   ReDim Preserve は【条件付きにしない】(CLAUDE.md §10 の R33波3 実測)。
'   nHits=0 は配列が未確保でありうるので Preserve なしの ReDim で確保する。
' ----------------------------------------------------------------------------
Public Function PrependHit(ByRef hits() As Hit, ByVal nHits As Long, ByRef newHit As Hit) As Long
    Dim n As Long: n = nHits
    If n < 0 Then n = 0

    Dim dupIdx As Long
    Dim i As Long
    For i = 1 To n
        If StrComp(hits(i).source, newHit.source, vbTextCompare) = 0 Then
            dupIdx = i
            Exit For
        End If
    Next i

    If dupIdx > 0 Then
        ' 既にある1件を先頭へ繰り上げる(間を1つずつ後ろへずらす)。
        For i = dupIdx To 2 Step -1
            hits(i) = hits(i - 1)
        Next i
        hits(1) = newHit
        PrependHit = n
        Exit Function
    End If

    If n = 0 Then
        ReDim hits(1 To 1)
    Else
        ReDim Preserve hits(1 To n + 1)
        For i = n + 1 To 2 Step -1
            hits(i) = hits(i - 1)
        Next i
    End If
    hits(1) = newHit
    PrependHit = n + 1
End Function

' ----------------------------------------------------------------------------
' DemoteWrongHits - 是正メモの「誤答の根拠:」行に載っている(source,page)と
'   一致する hit を末尾へ降格する(除外はしない)。先頭(=PrependHit が置いた
'   是正メモ自身)は絶対に動かさないので i=2 から見る。
'   PrependHit と同じ配列作法(hits はByRef・1始まり)。戻り値は降格した件数
'   (0なら何もしていない=呼び出し側のログや将来の可観測性のため)。
'   topK: 「これから残る枠」の外まで並べ替えても意味が無い(かつ枠外を動かすと
'   末尾側の走査が崩れる)ので、実効範囲を min(nHits, topK) に絞る。
'   topK < 1(上限なし)のときは nHits をそのまま使う。
' ----------------------------------------------------------------------------
Public Function DemoteWrongHits(ByVal memoText As String, ByRef hits() As Hit, _
                                ByVal nHits As Long, ByVal topK As Long) As Long
    DemoteWrongHits = 0
    Dim effN As Long: effN = nHits
    If topK >= 1 And topK < effN Then effN = topK
    If effN < 2 Then Exit Function   ' 先頭(是正メモ自身)しか無ければ動かす余地が無い

    Dim line As String: line = ExtractWrongSources(memoText)
    If LenB(line) = 0 Then Exit Function

    Dim demoted As Long
    Dim i As Long: i = 2
    Do While i <= effN - demoted
        If WrongSourceMatches(line, hits(i).source, hits(i).page) Then
            Dim moved As Hit: moved = hits(i)
            Dim j As Long
            For j = i To effN - demoted - 1
                hits(j) = hits(j + 1)
            Next j
            hits(effN - demoted) = moved
            demoted = demoted + 1
            ' i は据え置き ―― 繰り上がってきた次の要素を同じ位置で再検査する。
        Else
            i = i + 1
        End If
    Loop

    DemoteWrongHits = demoted
End Function

' ----------------------------------------------------------------------------
' InjectHits - 回答生成の直前に、質問が一致する是正メモを hits の先頭へ足す。
' ----------------------------------------------------------------------------
'   入口は modAskRetrieve.RunDeepScoped / RunUnscoped が値を返す直前の1行だけ
'   (modAskRetrieve は残1,092字のため、あちらには1行しか置けない)。
'   一般アシスタント(sendMode="normal")は modAsk 自体を通らないので、
'   ここに来ない=本棚を見ないモードへ是正が混ざることはない。
'   戻り値は注入後の件数。以下はいずれも nHits をそのまま返す:
'     ・nHits < 0(埋め込み失敗。もう一度呼んでも同じなので触らない)
'     ・config correct_inject = off
'     ・是正メモが1件も一致しない
'     ・途中で何か失敗した(サーキットブレーカー)
'   R36 Fix N4: 件数は topK を超えない。従来は nHits + 1 のまま返していたため、
'   topk_* が 6 の設定でも 7 件がプロンプトへ載り、設定した上限が守られて
'   いなかった(長い資料が7件並ぶと入念モードでトークン上限に触れる)。
'   落とすのは【末尾=最も弱いヒット】で、先頭へ入れた是正メモは必ず残る。
'   topK < 1(呼び出し側が上限を持たない)のときだけ切り詰めない。
Public Function InjectHits(ByVal q As String, ByRef hits() As Hit, ByVal nHits As Long, _
                           ByVal topK As Long) As Long
    InjectHits = nHits
    If nHits < 0 Then Exit Function
    If LenB(Trim$(q)) = 0 Then Exit Function

    On Error GoTo Done

    If LCase$(Trim$(modConfig.GetString("correct_inject", "on"))) = "off" Then Exit Function

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function

    Dim lastK As Long
    lastK = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function

    ' 【2段で読む理由】ここは質問1回ごとに必ず通る。my_knowledge は実機で
    ' 20,500行あり、full_text は1行が最大32,000字なので、
    ' modVaultGallery.EnsurePreviewIndex(:707)のように full_text 列まで
    ' 丸ごと配列にすると1質問あたり数十MBを読むことになる(あちらは画面操作の
    ' 1回きりでキャッシュも持つので成立している)。
    ' そこで (1) source 列だけを一括で読んで是正メモの行を絞り
    '        (2) 絞れた数行分だけ full_text を読む、の2段にする。
    ' 是正メモは「直した質問の数」しか無い(実機で数件〜数十件)。
    ' chunk_id(1列目)まで含めて【2列】読むのは、1列だけの Range.Value が
    ' 1行しか無いときに配列ではなくスカラーを返し、LBound で落ちるため
    ' (資料が1件だけの本棚で是正が丸ごと効かなくなる)。chunk_id は短いハッシュ
    ' 文字列なので、読んでも重くならない。
    Dim srcArr As Variant
    srcArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastK, 2)).Value

    Dim qNorm As String: qNorm = modSparse.NormalizeForSearch(q)

    Dim bestRow As Long, bestLv As Long
    Dim i As Long
    For i = LBound(srcArr, 1) To UBound(srcArr, 1)
        Dim nm As String: nm = Trim$(CStr(srcArr(i, 2)))
        If Left$(nm, Len(MEMO_PREFIX)) = MEMO_PREFIX Then
            ' srcArr の i 行目 = シートの i+1 行目(1行目は見出し)。
            Dim lv As Long: lv = MatchLevel(qNorm, CStr(ws.Cells(i + 1, 7).Value))
            ' R36 Fix A-m5: 同じレベルで複数当たったときは【後勝ち】(>=)。
            ' my_knowledge は取込順に積むので、後の行=新しい是正メモ。
            ' 従来の > は最も古い行を勝たせており、「直したのに古い答えが
            ' 出続ける」= 是正の仕組みそのものが利用者から見て壊れていた。
            ' lv >= 1 を先に見るのは、不一致(0)で bestRow を書かないため。
            If lv >= 1 And lv >= bestLv Then
                bestLv = lv
                bestRow = i + 1
            End If
        End If
        ' R36 Fix A-m5: 「完全一致を見つけたら打ち切る」は消した。打ち切ると
        ' 最初に見つかった=最も古い完全一致が勝ってしまい、後勝ちにならない。
        ' 走査するのは source 列に是正メモ接頭辞を持つ行だけ(実機で数件〜
        ' 数十件)なので、最後まで見ても質問1回あたりの負担は変わらない。
    Next i
    If bestLv < 1 Then Exit Function
    If bestRow < 2 Then Exit Function

    Dim memoHit As Hit
    memoHit.chunk_id = Trim$(CStr(ws.Cells(bestRow, 1).Value))
    memoHit.source = Trim$(CStr(ws.Cells(bestRow, 2).Value))
    memoHit.origin = "self"
    memoHit.page = 1
    memoHit.full_text = CStr(ws.Cells(bestRow, 7).Value)
    memoHit.preview = modUtil.SafeLeft(Replace(StripBreadcrumbLine(memoHit.full_text), vbLf, " "), 120)
    If bestLv >= 2 Then
        memoHit.score = SCORE_EXACT
    Else
        memoHit.score = SCORE_KEY
    End If

    Dim outN As Long: outN = PrependHit(hits, nHits, memoHit)

    ' R37 §2-1: 完全一致(bestLv=2)のときだけ、この是正メモが記録している
    ' 「誤答の根拠」に載っている hit を末尾へ降格する。語一致(=1)では動かさない
    ' (R36 §7 のリスク台帳「誤爆」を踏まえた裁定)。
    Dim demoted As Long
    If bestLv >= 2 Then demoted = DemoteWrongHits(memoHit.full_text, hits, outN, topK)

    ' R36 Fix N4: topk_* を超えたぶんは末尾(最も弱いヒット)を落とす。
    If topK >= 1 Then
        If outN > topK Then outN = topK
    End If
    InjectHits = outN

    On Error Resume Next
    ' 【R37 Fix B-m1】src= は必ず【行末】に置く。modAnalytics.ExtractSourceName は
    '   "src=" 以降を全部そのまま資料名として切り出す(区切りを見ない)ので、
    '   後ろに何か足すと分析CSVの資料名が "メモ… demoted=1" に化ける。
    modLog.LogUsage "correct_inject", "", _
        "level=" & bestLv & " demoted=" & demoted & " src=" & memoHit.source, 0, outN
    On Error GoTo 0
    Exit Function

Done:
    ' サーキットブレーカー: 是正が効かないだけで、質問には必ず答える。
    InjectHits = nHits
End Function

' 取込時に焼かれる breadcrumb 行【資料名>章>条】を preview から落とす。
' 剥がし方は modVaultGallery.MakePreviewText(:724-731)と同じ。
Private Function StripBreadcrumbLine(ByVal s As String) As String
    StripBreadcrumbLine = s
    If Left$(s, 1) <> "【" Then Exit Function
    Dim lfPos As Long: lfPos = InStr(1, s, vbLf, vbBinaryCompare)
    If lfPos > 0 Then StripBreadcrumbLine = Mid$(s, lfPos + 1)
End Function
