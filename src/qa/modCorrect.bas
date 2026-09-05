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
'   modConfig/modState/modUtil/modTypes)と同格の modSparse/modInsight だけ。
'   my_knowledge のシート直読みは同じ qa 層の既存前例に倣う
'   (modAskFocus.bas:80 / modAskGlobal.bas:294 / modRetrieve.bas:590)。
'   シート名は modAppDef.SH_KNOWLEDGE 経由で取る(リテラルを書かない)。
' ============================================================================

' 是正メモの資料名(source)の接頭辞。modAppAct と InjectHits の走査が同じ
' 文字列を見る単一情報源。ここを変えると過去の是正メモが拾えなくなる。
Private Const MEMO_PREFIX As String = "是正メモ_"

' 本文の「質問」行。組み立て(BuildMemoBody)は "質問: "、読み取り
' (ExtractQuestionLine)は "質問:" で照合する(全角スペース等のゆれを拾う)。
Private Const Q_HEAD_OUT As String = "質問: "
Private Const Q_HEAD_IN As String = "質問:"

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
' BuildMemoBody - 是正メモの本文(R36 §2-2-1 の書式)。純関数。
'   modVault.RegisterKnowledgeText が先頭へ "【是正メモ】" & vbLf を足すので、
'   ここが返すのは2行目以降にあたる本体。1行目の指示文は【LLMへの指示】として
'   意図的に本文側にも置く(タイトル行はチャンク先頭の breadcrumb に埋もれる
'   ことがあり、指示が消える経路を作らないため)。
' ----------------------------------------------------------------------------
Public Function BuildMemoBody(ByVal question As String, ByVal fixText As String, _
                              ByVal answerExcerpt As String) As String
    BuildMemoBody = "【是正メモ】この質問には、下の「正しい内容」で答えること。" & vbLf & _
                    Q_HEAD_OUT & question & vbLf & _
                    "正しい内容: " & fixText & vbLf & _
                    "(誤りだった回答の抜粋: " & answerExcerpt & ")"
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
Public Function MemoDocBase(ByVal question As String) As String
    Dim t As String: t = modUtil.SafeLeft(question, DOCBASE_Q_LEN)

    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i

    MemoDocBase = MEMO_PREFIX & Trim$(t)
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
    If total = 0 Then Exit Function

    KeyMatchPct = Int((okCnt * 100#) / total)
End Function

' ----------------------------------------------------------------------------
' MatchLevel - 質問と是正メモの一致度。2=完全一致 / 1=語一致 / 0=不一致。
'   qNorm は modSparse.NormalizeForSearch を通した質問。
'   完全一致の判定は modInsight.NormKey(空白・句読点・?・括弧を落として
'   先頭40字を小文字化。modInsight.bas:417-425)で、同じ質問の連投抑止が
'   使っているのと同じ物差しを使う(2箇所に書き写せば必ずズレる)。
' ----------------------------------------------------------------------------
Public Function MatchLevel(ByVal qNorm As String, ByVal memoText As String) As Long
    If LenB(Trim$(qNorm)) = 0 Then Exit Function

    Dim memoQ As String: memoQ = ExtractQuestionLine(memoText)
    If LenB(Trim$(memoQ)) = 0 Then Exit Function

    Dim memoNorm As String: memoNorm = modSparse.NormalizeForSearch(memoQ)

    Dim kq As String: kq = modInsight.NormKey(qNorm)
    Dim km As String: km = modInsight.NormKey(memoNorm)
    If LenB(kq) = 0 Then Exit Function
    If LenB(km) = 0 Then Exit Function
    If StrComp(kq, km, vbBinaryCompare) = 0 Then
        MatchLevel = 2
        Exit Function
    End If

    If KeyMatchPct(qNorm, memoNorm) >= KeyMinPct() Then MatchLevel = 1
End Function

' 語一致の閾値(%)。config が読めない実行環境(LO実行テスト)では既定値のまま
' 走り切る ―― modSparse.EffectiveKeyScoreCap と同型。
Private Function KeyMinPct() As Long
    Dim v As Long: v = DEFAULT_KEY_MIN
    On Error Resume Next
    v = modConfig.GetLong("correct_key_min", DEFAULT_KEY_MIN)
    On Error GoTo 0
    If v < 1 Then v = DEFAULT_KEY_MIN
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
'   件数は最大 nHits + 1(R36 §2-3)。topK での切り詰めはしない ―― この
'   シグネチャには topK が渡らず、切るなら呼び出し側の責務になるため。
Public Function InjectHits(ByVal q As String, ByRef hits() As Hit, ByVal nHits As Long) As Long
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
            If lv > bestLv Then
                bestLv = lv
                bestRow = i + 1
            End If
        End If
        If bestLv >= 2 Then Exit For      ' 完全一致より強い一致は無い
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

    InjectHits = PrependHit(hits, nHits, memoHit)

    On Error Resume Next
    modLog.LogUsage "correct_inject", "", "level=" & bestLv & " src=" & memoHit.source, 0, InjectHits
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
