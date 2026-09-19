Attribute VB_Name = "modFollowup"
Option Explicit

' ============================================================================
' modFollowup - 「続けて質問」(裁定D11)の応答パース/履歴整形の純関数群
' ----------------------------------------------------------------------------
' modAsk が文字数上限(実機VBAの1モジュール上限)に迫ったため、状態を持たない
' 純粋ヘルパーを機能単位でここへ退避した(behavior-preserving リファクタ)。
' 区切り文字はmodAsk側の確定値(FOLLOWUP_PAIR_SEP=";;;")を引数で受け取り、
' 本モジュール自身は他モジュールへ一切依存しない(String操作のみ)。
'
' 2026-08-03(R13-5b): これに加えて「会話出典メモリ」(直近ターンで引用した
' 資料名の短いリスト)を持つ。詳細はモジュール末尾の該当ブロックを参照。
' ============================================================================

' 覚えておく資料名の上限(R13-5b)。会話1本の文脈として意味を保てる範囲。
' 増やすほど「深掘り」が「本棚全体」に近づき、モードの区別が消える。
Private Const CITED_MAX As Long = 8

' 会話出典メモリ(vbLf区切り・新しい順・重複なし)。State Lossで消えても
' 「スコープ無し=従来動作」へ静かに戻るだけなので、保存はしない。
Private mCited As String

' ----------------------------------------------------------------------------
' 既出チャンクメモリ(2026-08-05 R16-3D)
' ----------------------------------------------------------------------------
' 深掘り(続けて質問)を2回続けると、2回目が1回目とほぼ同じ回答になる。原因は
' 検索が毎回まったく同じ上位チャンクを返すこと(資料単位のスコープはあっても、
' チャンク単位の既出抑制はコードベースに1つも無かった=実機報告の「深掘り
' ループ」の構造原因)。同じ材料を渡して「もっと深く」と頼んでも、モデルには
' 前回と違うことを書く材料が無い。
'
' そこで直近の深掘りで【実際に渡したチャンク】を覚え、次の深掘りでは後ろへ
' 回す。除外ではなく降格にするのは、小さな本棚では既出しか無いことが普通に
' あるため(除外にすると2回目が0件になる)。
' 資料名メモリ(mCited)と同じくメモリのみ・保存しない。
' R28 W2-2(逆質問で選ばれた資料の検索スコープ。実体はモジュール末尾)。
' キー名と「広げ直す件数」の2つだけを宣言部に置く(実機VBAはモジュール
' レベル宣言がプロシージャより後ろにあるとコンパイルできない)。
' 2件は RunDeepScoped(modAskRetrieve.bas:338)の既存作法と同じ値。
Private Const K_CLARIFY_PICK As String = "clarify_pick_src"
Private Const MIN_SCOPE_HITS As Long = 2

Private Const USED_MAX As Long = 64
Private mUsedChunks As String

' このターンが「続けて質問(深掘り)」かどうか。modApp が送信のたびに
' SetFollowupTurn で入れる唯一の窓口で、modPrompts(下書き指示の出し分け)と
' 既出メモリの寿命の両方がこの1つの旗を見る。新規質問(False)で既出メモリを
' 捨てるのは、前の話題のチャンクを新しい質問で降格させないため。
Private mFollowupTurn As Boolean

' 応答aから [[FOLLOWUP: ...]] をパースし、body=マーカー行除去後の本文 /
' candidates=候補(vbLf区切り。無ければ空)に分離する。V2と同じ寛容実装:
' "]]"が見つからない・中身が空・「なし」の場合は候補なし扱いとし、
' マーカーを含む行だけを本文から取り除く(パース失敗でも例外は出さない)。
Public Sub SplitFollowupTrailer(ByVal a As String, ByRef body As String, ByRef candidates As String)
    body = a
    candidates = ""
    If LenB(a) = 0 Then Exit Sub

    Dim fp As Long
    fp = InStr(1, a, "[[FOLLOWUP:", vbTextCompare)
    If fp = 0 Then Exit Sub

    Dim fq As Long
    fq = InStr(fp, a, "]]")
    If fq > 0 Then
        Dim fval As String
        fval = Trim$(Mid$(a, fp + Len("[[FOLLOWUP:"), fq - fp - Len("[[FOLLOWUP:")))
        If StrComp(fval, "なし", vbTextCompare) <> 0 And LenB(fval) > 0 Then
            Dim parts() As String
            parts = Split(fval, "|")
            Dim out As String
            Dim i As Long
            For i = LBound(parts) To UBound(parts)
                Dim t As String
                t = Trim$(parts(i))
                If LenB(t) > 0 Then
                    If LenB(out) > 0 Then out = out & vbLf
                    out = out & t
                End If
            Next i
            candidates = out
        End If
    End If

    ' マーカーを含む行を本文から除去し、末尾の空行・空白を刈り込む
    ' (V2 modPipeline.ParseTrailersと同じ流儀)。
    Dim lines() As String
    lines = Split(a, vbLf)
    Dim keep As String
    Dim j As Long
    For j = LBound(lines) To UBound(lines)
        If InStr(lines(j), "[[FOLLOWUP:") = 0 Then
            If LenB(keep) > 0 Then keep = keep & vbLf
            keep = keep & lines(j)
        End If
    Next j
    Do While Len(keep) > 0 And (Right$(keep, 1) = vbLf Or Right$(keep, 1) = vbCr Or Right$(keep, 1) = " ")
        keep = Left$(keep, Len(keep) - 1)
    Loop
    body = keep
End Sub

' 「新しい順<sep>区切り」文字列の先頭からmaxPairs件だけを残す。
Public Function KeepNewestPairs(ByVal joined As String, ByVal maxPairs As Long, ByVal sep As String) As String
    Dim parts() As String
    parts = Split(joined, sep)
    Dim n As Long
    n = UBound(parts) - LBound(parts) + 1
    If n <= maxPairs Then
        KeepNewestPairs = joined
        Exit Function
    End If

    Dim out As String
    Dim i As Long
    For i = LBound(parts) To LBound(parts) + maxPairs - 1
        If LenB(out) > 0 Then out = out & sep
        out = out & parts(i)
    Next i
    KeepNewestPairs = out
End Function

' 履歴に積む文字列から区切り<sep>(";;;")を除去する。単純な1回のReplaceでは
' ";;;;;;"→";;;;"のように置換結果へ再び";;;"が現れ得るため、無くなるまで
' 繰り返す(各回で必ず短くなるので有限回で終わる)。
Public Function SanitizeForFollowupHistory(ByVal s As String, ByVal sep As String) As String
    Dim t As String
    t = s
    Do While InStr(t, sep) > 0
        t = Replace(t, sep, ";;")
    Loop
    SanitizeForFollowupHistory = t
End Function

' ============================================================================
' 会話出典メモリ(2026-08-03 R13-5b)
' ----------------------------------------------------------------------------
' 「深掘り=会話の流れ(引用済み資料)の中を深く掘る」を成立させるには、
' 直近のターンでどの資料を根拠にしたかを覚えておく場所が要る。それが無い
' ため、これまで deep と thorough はパラメータ量が違うだけの同じ検索だった
' (実機第2報 RC7)。
'
' ここに置く理由: 状態は「今の会話」に属し、ブックへ保存する類のものでは
' ない。VBAのState Lossで消えても「スコープ無し=従来動作」へ静かに戻るだけで
' 壊れない(消えて困る情報は持たない)。
' 保持形式は vbLf 区切りの資料名リスト(新しい順・重複なし・上限CITED_MAX件)。
' 宣言(CITED_MAX / mCited)はモジュール先頭にある。
' ============================================================================

' 直近ターンで引用した資料名を登録する。names は "|" 区切り(modAskRetrieve.
' HitSourceList の形式)でも改行区切りでも受ける。
' replaceAll=True(=followupでない新規質問)のときは覚え直す(上書き)。
Public Sub RememberCitedSources(ByVal names As String, Optional ByVal replaceAll As Boolean = False)
    If replaceAll Then mCited = ""
    mCited = MergeCitedSources(mCited, names, CITED_MAX)
End Sub

' 現在の会話出典(vbLf区切り)。空なら空文字列。
' @unused:読み出し口として公開しているが呼び出し元が1つも無い(R49 の監査 H-H-10
'   で、過去の監査報告に名前が出ているだけの救済が外れて発覚)。会話出典の書き込み
'   (RememberCitedSources)とクリア(ClearCitedSources)は使われている。削除可否は R50 で裁定する。
Public Function CitedSourcesLine() As String
    CitedSourcesLine = mCited
End Function

' 会話リセット(チャットのクリア)で消す。既出チャンクも同じ会話の持ち物
' なので一緒に捨てる(片方だけ残すと、消したはずの話題の材料が降格し続ける)。
Public Sub ClearCitedSources()
    mCited = ""
    mUsedChunks = ""
End Sub

' ----------------------------------------------------------------------------
' SetFollowupTurn / IsFollowupTurn - このターンが深掘りかどうか(R16-3D)。
'   modApp.OnSend が「続けて質問」の武装を消費した直後に1回だけ入れる。
'   False(=新規質問)を入れた時点で既出チャンクメモリを捨てる。前の話題で
'   読んだチャンクを新しい質問で後ろへ回すと、いちばん関係のある資料が
'   理由もなく沈む。
' ----------------------------------------------------------------------------
Public Sub SetFollowupTurn(ByVal b As Boolean)
    mFollowupTurn = b
    If Not b Then mUsedChunks = ""
End Sub

Public Function IsFollowupTurn() As Boolean
    IsFollowupTurn = mFollowupTurn
End Function

' ----------------------------------------------------------------------------
' RememberUsedChunks - このターンで根拠として渡したチャンクを覚える(R16-3D)。
'   新しい順・重複なし・上限USED_MAX件(資料名メモリと同じ畳み方を共用)。
'   chunk_id が空のヒット(旧データ・テストダブル)は覚えない=次回も降格
'   されない。判定材料が無いものを「既出」に倒すと、二度と出てこなくなる。
' ----------------------------------------------------------------------------
Public Sub RememberUsedChunks(hits() As Hit, ByVal nHits As Long)
    Dim ids As String
    Dim i As Long
    For i = 1 To nHits
        Dim id As String
        id = Trim$(hits(i).chunk_id)
        If LenB(id) > 0 Then
            If LenB(ids) > 0 Then ids = ids & vbLf
            ids = ids & id
        End If
    Next i
    RememberUsedIds ids
End Sub

' 上の文字列版(vbLf区切りの chunk_id 列)。実体はこちらで、Hit配列版は
' 詰め替えて呼ぶだけ。分けてあるのは LibreOffice の実行テストが
' 【別モジュールで定義された Type の配列】を扱えないため(modTestsPure 冒頭の
' 既知制約)。ここが文字列で受けられるので、既出降格の規則そのものは
' 実機を待たずに回帰テストで固定できる。
Public Sub RememberUsedIds(ByVal idsLine As String)
    If LenB(idsLine) = 0 Then Exit Sub
    mUsedChunks = MergeCitedSources(mUsedChunks, idsLine, USED_MAX)
End Sub

' ----------------------------------------------------------------------------
' DemoteUsed - 既出チャンクを後ろへ回して keepK 件へ切る(R16-3D)。
'   呼び出し側は topK の2倍を検索してからここへ渡す。未出が十分あれば
'   未出だけが残り、未出が足りなければ既出で埋め戻される(常に
'   min(nHits, keepK) 件が残る=深掘りが空振りする経路を作らない)。
'   並べ替えは安定(未出どうし・既出どうしの相対順=検索順位はそのまま)。
'   覚えている既出が無いターンは並べ替えず、件数を keepK へ切るだけ。
' ----------------------------------------------------------------------------
Public Sub DemoteUsed(ByRef hits() As Hit, ByRef nHits As Long, ByVal keepK As Long)
    If nHits < 1 Then Exit Sub

    Dim ids As String
    Dim i As Long
    For i = 1 To nHits
        If i > 1 Then ids = ids & vbLf
        ids = ids & Trim$(hits(i).chunk_id)
    Next i

    Dim ordLine As String
    ordLine = DemoteOrder(ids, keepK)
    If LenB(ordLine) = 0 Then Exit Sub

    Dim ord() As String
    ord = Split(ordLine, ",")
    Dim nOut As Long
    nOut = UBound(ord) - LBound(ord) + 1
    If nOut < 1 Or nOut > nHits Then Exit Sub

    ' 元の位置 → 現在の位置。並べ替えは「決めた順に前から引き抜いて差し込む」
    ' 挿入法で行う。作業用の Hit 配列を作らないのは、別モジュールで定義した
    ' Type の配列が LibreOffice の実行テスト環境で扱えないため(modTestsPure
    ' 冒頭の既知制約)。Long の配列なら両環境で問題なく使える。
    Dim cur() As Long
    ReDim cur(1 To nHits)
    For i = 1 To nHits
        cur(i) = i
    Next i

    Dim tmp As Hit
    Dim k As Long, j As Long, orig As Long, fromIx As Long
    For k = 1 To nOut
        orig = CLng(Val(ord(LBound(ord) + k - 1)))
        If orig >= 1 And orig <= nHits Then
            fromIx = cur(orig)
            If fromIx > k Then
                tmp = hits(fromIx)
                For j = fromIx To k + 1 Step -1
                    hits(j) = hits(j - 1)
                Next j
                hits(k) = tmp
                For j = 1 To nHits
                    If cur(j) >= k And cur(j) < fromIx Then cur(j) = cur(j) + 1
                Next j
                cur(orig) = k
            End If
        End If
    Next k

    nHits = nOut
End Sub

' ----------------------------------------------------------------------------
' DemoteOrder - 降格後に採る「元の位置」の並びを "," 連結で返す(純関数)。
'   idsLine: いまのヒットの chunk_id を vbLf 区切りで並べたもの(空要素可)。
'   戻り値 : 例 "2,4,1,3"(1起点・keepK 件で打ち切り)。
'
'   規則はこの1関数だけが持つ(DemoteUsed は決まった順に要素を入れ替える係)。
'     ・未出(=覚えていない chunk_id)を元の順序のまま前へ
'     ・そのあとに既出を元の順序のまま並べる(除外ではなく降格)
'     ・先頭 keepK 件で切る → 未出が足りなければ既出で埋まる=空にならない
'   chunk_id が空のヒットは「覚えようがない」ので常に未出として扱う。
'   判定材料の無いものを既出へ倒すと、二度と表に出てこなくなる。
' ----------------------------------------------------------------------------
Public Function DemoteOrder(ByVal idsLine As String, ByVal keepK As Long) As String
    If LenB(idsLine) = 0 Then Exit Function

    Dim raw() As String
    raw = Split(idsLine, vbLf)
    Dim lo As Long: lo = LBound(raw)
    Dim n As Long: n = UBound(raw) - lo + 1
    If n < 1 Then Exit Function

    Dim keep As Long
    keep = keepK
    If keep < 1 Or keep > n Then keep = n

    Dim outS As String
    Dim cnt As Long
    Dim pass As Long, i As Long
    For pass = 0 To 1
        For i = 1 To n
            If cnt >= keep Then Exit For
            Dim isUsed As Boolean
            isUsed = InCitedScope(mUsedChunks, Trim$(raw(lo + i - 1)))
            If (pass = 0) <> isUsed Then
                If LenB(outS) > 0 Then outS = outS & ","
                outS = outS & CStr(i)
                cnt = cnt + 1
            End If
        Next i
    Next pass
    DemoteOrder = outS
End Function

' 検索へ渡す許可Dictionary。1件も覚えていなければ Nothing(=スコープ無し)。
Public Function CitedSourcesDict() As Object
    If LenB(mCited) = 0 Then Exit Function
    Set CitedSourcesDict = ScopeDictFrom(mCited)
End Function

' 資料名リスト(vbLf または "|" 区切り)から許可Dictionaryを作る。
' 中身が1件も無い/Dictionaryを作れない端末では Nothing を返し、
' 呼び出し側は「スコープ無し=従来動作」へ静かに退化する(無言の劣化はしない。
' 呼び出し側が usage_log へ記録する)。
Public Function ScopeDictFrom(ByVal namesLine As String) As Object
    If LenB(namesLine) = 0 Then Exit Function

    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    If d Is Nothing Then Exit Function

    Dim parts() As String
    parts = Split(Replace(namesLine, "|", vbLf), vbLf)
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim nm As String
        nm = Trim$(parts(i))
        If LenB(nm) > 0 Then
            If Not d.Exists(nm) Then d.Add nm, True
        End If
    Next i
    If d.count = 0 Then Exit Function
    Set ScopeDictFrom = d
End Function

' 新しい順・重複なしで cap 件までに収めた資料名リスト(vbLf区切り)を返す。
' newNames が既存より前に来る(=最後に引用した資料が最優先で残る)。
Public Function MergeCitedSources(ByVal existing As String, ByVal newNames As String, _
                                  ByVal cap As Long) As String
    Dim lim As Long
    lim = cap
    If lim < 1 Then lim = 1

    Dim src As String
    src = Replace(newNames, "|", vbLf)
    If LenB(existing) > 0 Then
        If LenB(src) > 0 Then src = src & vbLf
        src = src & existing
    End If

    Dim parts() As String
    parts = Split(src, vbLf)
    Dim out As String
    Dim cnt As Long
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim nm As String
        nm = Trim$(parts(i))
        If LenB(nm) > 0 Then
            If Not InCitedScope(out, nm) Then
                If LenB(out) > 0 Then out = out & vbLf
                out = out & nm
                cnt = cnt + 1
                If cnt >= lim Then Exit For
            End If
        End If
    Next i
    MergeCitedSources = out
End Function

' 資料名がスコープ(vbLf区切りの許可リスト)に入っているか。
' 判定は【完全一致・大小区別あり】。検索側(modRetrieve)が使う
' Dictionary.Exists と同じ意味にすること。ここが非対称になると、
' 「スコープに入れたはずの資料が落ちる」「入れていない資料が混じる」の
' どちらかが静かに起きる。
Public Function InCitedScope(ByVal scopeLine As String, ByVal srcName As String) As Boolean
    If LenB(scopeLine) = 0 Then Exit Function
    If LenB(srcName) = 0 Then Exit Function
    InCitedScope = (InStr(1, vbLf & scopeLine & vbLf, vbLf & srcName & vbLf, vbBinaryCompare) > 0)
End Function

' ============================================================================
' 逆質問で選ばれた資料の検索スコープ(2026-08-12 R28 W2-2・実機第13報②)
' ============================================================================
' 逆質問で「1. 火災保険約款」を選んでも、これまで検索は本棚全体のままだった。
' modClarify.MergeAnswer が作る「対象の資料: ○○」はクエリ文字列のヒントに
' すぎず、modAskRetrieve.RunUnscoped は scope 引数を一度も渡していない
' (調査B班・modAskRetrieve.bas:423-427)。番号で答えたのに絞り込まれない、が
' 「逆質問に答えても無駄」という体験の正体で、逆質問そのものが死ぬ。
'
' 運び方は ui_state の1キー(clarify_pick_src)。modAsk/modApp を1行も触らずに
' modClarify(書く側)と modAskRetrieve(読む側)を繋げる唯一の道で、橋渡し
' (modConvBridge)が既に使っている作法と同じ。
' 【1ターン限り】が絶対条件: TakeClarifyScope は読んだ瞬間にキーを消す。
' 消し忘れると以後すべての質問がその1資料に閉じ込められ、本棚が壊れたように
' 見える(利用者からは原因が絶対に分からない種類の事故)。
'
' このブロックだけは modState/modLog に依存する(モジュール冒頭の「他モジュール
' へ一切依存しない」はここより上の純関数群についての記述)。スコープの構築は
' 既存の ScopeDictFrom を使い、許可判定の意味を1つに保つ。
' 逆質問で確定した資料名を、次の検索1回ぶんだけ預ける("" で取り消し)。
Public Sub NoteClarifyPick(ByVal srcName As String)
    On Error Resume Next
    modState.SaveState K_CLARIFY_PICK, srcName
    On Error GoTo 0
End Sub

' 預けた資料名を読み、その場でキーを消してから検索スコープを作る。
' 印が無ければ Nothing = 呼び出し側は従来どおり本棚全体を検索する。
Public Function TakeClarifyScope() As Object
    Dim nm As String
    On Error Resume Next
    nm = modState.LoadState(K_CLARIFY_PICK, "")
    If LenB(nm) > 0 Then modState.SaveState K_CLARIFY_PICK, ""
    On Error GoTo 0
    If LenB(Trim$(nm)) = 0 Then Exit Function
    Set TakeClarifyScope = ScopeDictFrom(nm)
End Function

' スコープ内検索の結果から「本棚全体へ広げ直すか」を決める(純関数)。
' n = -1 は埋め込み失敗で、広げ直しても同じ結果にしかならない(RunDeepScoped
' の :345-348 と同じ判断)ので広げない。
Public Function ScopeNeedsWiden(ByVal n As Long) As Boolean
    If n = -1 Then Exit Function
    ScopeNeedsWiden = (n < MIN_SCOPE_HITS)
End Function

' スコープ内検索の結果を採用してよいか。False=呼び出し側は本棚全体へ広げ直す。
' 広げ直す回だけ usage_log に clarify_scope_widen を残す(無言の劣化はしない)。
Public Function ClarifyScopeKept(ByVal scopeD As Object, ByVal n As Long) As Boolean
    If scopeD Is Nothing Then Exit Function
    If Not ScopeNeedsWiden(n) Then ClarifyScopeKept = True: Exit Function
    On Error Resume Next
    modLog.LogUsage "clarify_scope_widen", "", _
        "選ばれた資料の中では" & n & "件しか見つからず、本棚全体へ広げ直しました(scope=" & scopeD.count & ")"
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ClarifyLoneIntent - 逆質問の返事で「数を1つだけ打ったのに資料番号として
'   使えない」ときに、その数を意図番号(1..maxIntent)として読み直す(純関数)。
' ----------------------------------------------------------------------------
' R28H F7(m-1/m-2): modClarify.ParseClarifyReply は1個目の数を資料番号として
' しか見ておらず、1..nSrc の外なら黙って捨てていた。実害は2つ:
'   ・資料の選択肢を出していない逆質問(nSrc=0。意図だけを聞く型)へ「3」と
'     答えると、資料番号としても意図番号としても採られず何も選ばれない。
'   ・資料を2件しか出していないところへ「5」と答えると同じく何も選ばれない
'     (利用者は意図の⑤を半角で打っただけ)。
'   いずれも「番号で答えたのに、自分の言葉で書き直した」扱いへ落ちる。
' 採る条件は4つとも必須:
'   srcIdx=0  … 資料番号としては採れていない(採れていれば1個目は資料が正)
'   n2=0      … 2個目の数が無い(「2-3」の形なら従来の規則が正しい)
'   circled=0 … 丸数字も無い(丸数字は常に意図番号で、そちらが優先)
'   1<=n1<=maxIntent … 「0」は「この中にない/わからない」の意思表示なので採らない
' 採らないときは0を返す(呼び出し側の intentIdx を書き換えない)。
' 実体をここへ置いたのは modClarify が残518字のため(憲章§4-6)。
Public Function ClarifyLoneIntent(ByVal srcIdx As Long, ByVal n1 As Long, _
                                  ByVal n2 As Long, ByVal circled As Long, _
                                  ByVal maxIntent As Long) As Long
    If srcIdx <> 0 Then Exit Function
    If n2 <> 0 Then Exit Function
    If circled <> 0 Then Exit Function
    If n1 < 1 Or n1 > maxIntent Then Exit Function
    ClarifyLoneIntent = n1
End Function
