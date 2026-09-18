Attribute VB_Name = "modMode"
Option Explicit

' ============================================================================
' modMode - 回答モードの方針(すぐ聞く / 通常 / 入念)
' ----------------------------------------------------------------------------
' なぜ3段階なのか(実測に基づく)
' ----------------------------------------------------------------------------
' 実物6資料180チャンク31問で採点した結果(tools/bench_retrieval.py):
'
'   手法                        R@1   R@3   R@5   R@10  R@20   ページ
'   単発検索(InStr・空白除去)     84%   84%   90%   90%   94%    81%
'   多クエリ→RRF統合            77%   84%   90%   97%   97%    71%
'
' ここから2つのことが分かる:
'   1. 「LLMに渡す件数」を増やしても効かない。R@5 90% → R@10 90% で頭打ち。
'      Qwen案の「quick 6→10件 / deep 12→24件」は、この実測では0%改善だった。
'      むしろ無関係なチャンクが増えて Lost in the middle を招く。
'   2. 「引き方(角度)」を増やすと効く。多クエリにすると R@10 が 90%→97%。
'      ただし1位の精度は落ちる(84%→77%)。角度を増やすほど、
'      本命以外も上位に混ざるため。
'
' つまりモードごとに最適な戦略が違う:
'   ・すぐ聞く … 少数しか渡さないので R@1 が命。単発検索が最良(84%)
'   ・通常     … 中間。単発+質問拡張
'   ・入念     … 多く渡してLLMと検証段に選ばせるので R@10 が命。
'                多クエリ(97%)+再ランク+下書き→照合
'
' 金融ドメインでは「ここ一番で外さない」ことが信用の全てなので、
' 入念モードは時間を一切気にせず精度側に全振りする。
'
' R4準拠(純ロジック): Excelオブジェクトに触れないので LibreOffice でテスト可能。
' ============================================================================

' ----------------------------------------------------------------------------
' 直近ターンで根拠表示(信頼度バッジ・出典チップ)を出してよいか(R33H F7)
' ----------------------------------------------------------------------------
' W5-21 は「回答が成立しなかったターンにバッジと出典を出さない」ために
' modAsk の mLastMode へ "" を入れた。ところが mLastMode="" は
' modUIMain.RenderAnswer が【空質問ターン専用】の目印として先に使っている印で、
' そこへ API失敗(#ERR)と逆質問が流れ込むと、
'   (a) 通信失敗の直後に状態セルが「状態: 準備できています」になる
'   (b) mLastAnswerText に到達せず、「Wordで開く」が【前のターンの回答】を出す
' という別の嘘が生まれた(nexus_ui=FALSE の端末は modDiag の必須キーで、
' 旧3画面は死んでいない)。
' 「検索も生成もしなかった」(空質問)と「生成が成立しなかった」(API失敗・
' 逆質問)は別の事実なので、別の印で表す。mLastMode は従来どおりモード名を
' 入れ、根拠表示の抑止だけをこの1ビットで行う。
' 置き場が modMode なのは modAsk が凍結モジュールだから(W5-21 と同じ理由。
' modAsk 側の変更は Done: の1行だけ)。作法は modAskGlobal.mWasGlobal と同型
' ―― 「このターンが何だったか」を1ビットだけ持ち、表示側が読む。
' 判定関数(Normalize / ShouldEmitInsight など)は従来どおり副作用ゼロのまま。
Private mGrounded As Boolean

' R34 B1: 直近ターンの回答モード名(quick/deep/thorough。正規化済みの生値)。
' NoteAnswered が控える。mGrounded と同じ「このターンが何だったか」の1ビット系で、
' 出典突合の後付け(AnnotateIfNeeded)がモードを見分けるためだけに使う。
' 一般アシスタント(modAsk.NoteGeneralAnswered)と空質問ターンはここを通らないため
' 前ターンの値が残るが、どちらも modApp 側の grounded ガードと
' modAsk.LastHitCount()=0 の二重で弾かれる(注釈は走らない)。
Private mLastAnsweredMode As String

' モード識別子。文字列はui_state/configと共有するので変更しないこと。
Private Const MODE_QUICK As String = "quick"
Private Const MODE_DEEP As String = "deep"
Private Const MODE_THOROUGH As String = "thorough"

' ----------------------------------------------------------------------------
' Normalize - 文字列を3つのモードのいずれかへ正規化する(既定=すぐ聞く)。
' ----------------------------------------------------------------------------
Public Function Normalize(ByVal mode As String) As String
    Dim m As String: m = LCase$(Trim$(mode))
    If m = MODE_THOROUGH Then
        Normalize = MODE_THOROUGH
    ElseIf m = MODE_DEEP Then
        Normalize = MODE_DEEP
    Else
        Normalize = MODE_QUICK
    End If
End Function

' ----------------------------------------------------------------------------
' NextMode - ボタンを押したときの巡回順(すぐ聞く → 通常 → 入念 → すぐ聞く)。
' ----------------------------------------------------------------------------
Public Function NextMode(ByVal mode As String) As String
    Select Case Normalize(mode)
        Case MODE_QUICK:    NextMode = MODE_DEEP
        Case MODE_DEEP:     NextMode = MODE_THOROUGH
        Case Else:          NextMode = MODE_QUICK
    End Select
End Function

' ----------------------------------------------------------------------------
' Caption - 画面に出す名前。何をするモードかが分かる言い方にする。
'   「速い/遅い」ではなく「何をするか」が違う、と伝えるのが目的。
' ----------------------------------------------------------------------------
Public Function Caption(ByVal mode As String) As String
    Select Case Normalize(mode)
        Case MODE_THOROUGH
            Caption = ChrW(&HD83D) & ChrW(&HDD2C) & " 入念に調べる"
        Case MODE_DEEP
            Caption = ChrW(&HD83D) & ChrW(&HDD0D) & " しっかり調べる"
        Case Else
            Caption = ChrW(&H26A1) & " すぐ聞く"
    End Select
End Function

' ----------------------------------------------------------------------------
' Description - 待ち時間の目安と、そのモードが何をするか(利用者向け)。
'   時間を隠さない。隠すと「固まった」と思われる。
' ----------------------------------------------------------------------------
' 2026-08-03(R13-5d): 3モードの役割を「速さの段階」から「探し方の違い」へ
' 書き直した。deep と thorough が「同じことをパラメータ違いでやる」状態
' (実機第2報 RC7)を直した以上、説明も役割で言い分ける。
'   すぐ聞く   … まず速く
'   しっかり   … 会話の流れ(引用済み資料)の中を深く
'   入念       … 本棚全体を広く
'
' 2026-08-03(R14-8b): 入念モードが実際に何をするかを書き直した。R14-8a で
' 「要点整理 → 下書き → 自己批判 → 検証 → 出典の突合」の5段になり、実測で
' 2～4分かかる。所要時間を「数分」とぼかすと、待っている人は止まったのか
' 判断できない(憲章§3-2)。刻んだ段と幅のある実測値をそのまま書く。
'
' 2026-08-11(R26-1): 一般アシスタントも3段化した(それまでは速さトグルが
' 社内ナレッジ検索にしか効かず、トーストに「一般アシスタントでは使われない」
' という注記を出していた)。同じトグルが2つの画面モードを支配する以上、
' 説明も両方を1文ずつ書く。書き分けないと、注記を外した瞬間に
' 「一般アシスタントでも本棚を広く調べる」と読めてしまう。
Public Function Description(ByVal mode As String) As String
    Select Case Normalize(mode)
        Case MODE_THOROUGH
            Description = "下書きを自己検証してから答えます。社内ナレッジ検索は本棚全体を" & _
                          "広く多段検証(要点整理 → 下書き → 自己批判 → 検証 → 出典の突合)、" & _
                          "一般アシスタントは別の視点で査読して書き直します。2～4分ほど。"
        Case MODE_DEEP
            ' R46: 実機で「直前までに使った資料しか見ないの? 本棚全部見るん
            ' じゃなかったっけ?」と疑義。実装は正しく、文言が誤っていた。
            ' 絞るのは modAsk.bas:190 の「isFollowup かつ MODE_DEEP」のときだけ
            ' で、しかも modAskRetrieve.DeepScopedCore はスコープ内が2件未満なら
            ' 本棚全体へ広げ直す。R13-5 の裁定書は 5c に条件付きで正しく書いて
            ' あるのに、5d の「文言を更新」でその条件が落ちていた(発生点)。
            ' 条件とフォールバックの両方を書く。
            Description = "会話の流れの中を深く。社内ナレッジ検索は本棚全体" & _
                          "(「続けて質問」のときだけ直前の資料へ絞り、足りなければ広げ直し)、" & _
                          "一般アシスタントは前提と論点を整理して注意点まで。1～2分ほど。"
        Case Else
            Description = "速さ優先。まず速く、そのまま答えます(社内ナレッジ検索は" & _
                          "そのまま検索)。数秒～20秒ほど。"
    End Select
End Function

' ----------------------------------------------------------------------------
' TopK - LLMへ渡すチャンク数。
'   実測で「増やしても効かない」ことが分かっているので闇雲に増やさない
'   (R@5 90% → R@10 90%)。入念だけ多クエリの成果を拾える幅を取る。
' ----------------------------------------------------------------------------
Public Function TopK(ByVal mode As String, ByVal quickK As Long, _
                     ByVal deepK As Long, ByVal thoroughK As Long) As Long
    Select Case Normalize(mode)
        Case MODE_THOROUGH: TopK = thoroughK
        Case MODE_DEEP:     TopK = deepK
        Case Else:          TopK = quickK
    End Select
    If TopK < 1 Then TopK = 6
End Function

' ----------------------------------------------------------------------------
' 各段を実行するか。入念は「時間より精度」なので全段を必ず通す。
' ----------------------------------------------------------------------------

' 質問拡張(1つの質問を複数の角度へ展開する)
Public Function UseExpand(ByVal mode As String, ByVal cfgEnabled As Boolean, _
                          ByVal quickOptIn As Boolean) As Boolean
    Select Case Normalize(mode)
        Case MODE_THOROUGH: UseExpand = True
        Case MODE_DEEP:     UseExpand = cfgEnabled
        Case Else:          UseExpand = cfgEnabled And quickOptIn
    End Select
End Function

' 再ランク(候補を関連度順に並べ直す)
Public Function UseRerank(ByVal mode As String, ByVal cfgEnabled As Boolean, _
                          ByVal quickOptIn As Boolean) As Boolean
    Select Case Normalize(mode)
        Case MODE_THOROUGH: UseRerank = True
        Case MODE_DEEP:     UseRerank = cfgEnabled
        Case Else:          UseRerank = cfgEnabled And quickOptIn
    End Select
End Function

' 検証段(下書きを資料と1行ずつ突き合わせる)。ここが「間違えない」の要。
' ----------------------------------------------------------------------------
' R48 警告: この関数は【入念(thorough)でも True を返す】。
' modAsk.AnswerWithContext で入念が「しっかり(RunDeepFlow)」へ落ちないのは、
' thorough の枝がこの UseVerify の枝より【前】に書かれているからだけである。
' 順序を入れ替える・thorough の枝の条件を緩める・前に別の枝を挟む、のどれを
' やっても、入念が静かに2段の RunDeepFlow(recommended_model=terra)へ降格する。
' 回答は出るので実機でも気付けない。modTestsPure47 がこの戻り値を固定している。
Public Function UseVerify(ByVal mode As String) As Boolean
    Dim m As String: m = Normalize(mode)
    UseVerify = (m = MODE_DEEP Or m = MODE_THOROUGH)
End Function

' 再ランク段の reasoning_effort(2026-08-03 R14-8a)。
'   入念モードだけ別の設定(rerank_effort_thorough、既定medium)を使う。
'   再ランクは「どの資料を根拠にするか」を決める段で、ここを間違えると
'   後段でいくら検証しても直らない。入念だけは他モードの速度事情に
'   引きずられないようにする。空文字が渡されたら従来値へ倒す。
Public Function RerankEffort(ByVal mode As String, ByVal baseEffort As String, _
                             ByVal thoroughEffort As String) As String
    RerankEffort = baseEffort
    If Normalize(mode) <> MODE_THOROUGH Then Exit Function
    If LenB(Trim$(thoroughEffort)) = 0 Then Exit Function
    RerankEffort = thoroughEffort
End Function

' 拡張の軽量版を使うか(すぐ聞くのときだけ。入念で軽くしては意味が無い)
Public Function UseLightExpand(ByVal mode As String, ByVal cfgLight As Boolean) As Boolean
    UseLightExpand = (Normalize(mode) = MODE_QUICK) And cfgLight
End Function

' ----------------------------------------------------------------------------
' 質問処理の段階ナレーション(2026-08-03 R13-9b)
' ----------------------------------------------------------------------------
' 「(1/4) 質問を分解中…」のような番号付き実況を出すための算数だけをここに置く
' (表示そのものは modAskRetrieve が SetStage へ流す)。
'
' 番号の誠実さについて:
'   拡張(expand)と再ランク(rerank)はモードとconfigで実行有無が変わる。
'   総数を常に4と書くと、2段しか通らない構成で「(3/4)で終わる」という嘘に
'   なる。そこで総数は【その回に通す段の数】から作る。
'   quick(検証段を通さない)では番号を出さない=AskStageTotal が0を返し、
'   AskStageText が素のラベルへ退化する。
'   なお rerank は「候補数がtopKより多いときだけ」実行されるため、計画には
'   入ったが実行されない回がある。その場合は番号が1つ飛ぶ(総数は嘘に
'   ならず、最後は必ず (N/N) で終わる)。
'
' 2026-08-03(R14-8a): 入念モードは要点整理(digest)と自己批判(critique)の
'   2段が増えて最大6段になった。hasThorough を足すのは Optional にしてある
'   ので、深掘り(4段まで)の番号は1つも動かない。
Public Function AskStageTotal(ByVal hasExpand As Boolean, ByVal hasRerank As Boolean, _
                              ByVal hasVerify As Boolean, _
                              Optional ByVal hasThorough As Boolean = False) As Long
    If Not hasVerify Then Exit Function      ' 0 = 番号を出さない(すぐ聞く)
    AskStageTotal = 2                        ' 下書き + 検証は必ず通る
    If hasExpand Then AskStageTotal = AskStageTotal + 1
    If hasRerank Then AskStageTotal = AskStageTotal + 1
    If hasThorough Then AskStageTotal = AskStageTotal + 2   ' 要点整理 + 自己批判
End Function

' 段の番号(1起点)。計画に入っていない段は0(=番号を出さない)。
Public Function AskStageIndex(ByVal kind As String, ByVal hasExpand As Boolean, _
                              ByVal hasRerank As Boolean, _
                              Optional ByVal hasThorough As Boolean = False) As Long
    Dim head As Long
    head = 0
    If hasExpand Then head = head + 1
    If hasRerank Then head = head + 1

    Select Case LCase$(Trim$(kind))
        Case "expand"
            If hasExpand Then AskStageIndex = 1
        Case "rerank"
            If hasRerank Then
                If hasExpand Then
                    AskStageIndex = 2
                Else
                    AskStageIndex = 1
                End If
            End If
        Case "digest"
            ' 入念モード以外にこの段は無い(番号を出さない)。
            If hasThorough Then AskStageIndex = head + 1
        Case "draft"
            If hasThorough Then
                AskStageIndex = head + 2
            Else
                AskStageIndex = head + 1
            End If
        Case "critique"
            If hasThorough Then AskStageIndex = head + 3
        Case "verify"
            If hasThorough Then
                AskStageIndex = head + 4
            Else
                AskStageIndex = head + 2
            End If
    End Select
End Function

' 段のラベル(利用者の言葉。専門用語を出さない)。
Public Function AskStageLabel(ByVal kind As String) As String
    Select Case LCase$(Trim$(kind))
        Case "expand":   AskStageLabel = "質問を分解中…"
        Case "rerank":   AskStageLabel = "資料を照合中…"
        Case "digest":   AskStageLabel = "資料の要点を整理中…"
        Case "draft":    AskStageLabel = "下書きを作成中…"
        Case "critique": AskStageLabel = "下書きを自己点検中…"
        Case "verify":   AskStageLabel = "検証中…"
        Case Else:       AskStageLabel = "回答を作成中…"
    End Select
End Function

' 実況テキスト。番号が付けられないときは素のラベルへ退化する(嘘の番号を
' 出すくらいなら番号を出さない)。
Public Function AskStageText(ByVal idx As Long, ByVal total As Long, ByVal label As String) As String
    If total > 0 And idx > 0 And idx <= total Then
        AskStageText = "(" & idx & "/" & total & ") " & label
    Else
        AskStageText = label
    End If
End Function

' ----------------------------------------------------------------------------
' ShouldEmitInsight - 「解決した」を押したとき、部内へ発信してよい回答か
'                     (2026-08-03 R14-1b / 実機第3報 RC2)
' ----------------------------------------------------------------------------
' 発信とは (a) 資料の作者への感謝状(modP2P.EmitThanksForLastAnswer)と
' (b) 解決済みQ&Aの部内共有(modInsightIo.EmitVerifiedQA)の2つ。どちらも
' 「本棚の資料を根拠に答えた」ことが前提で、根拠が無い回答で撃つと
'   ・一般アシスタントの雑談が「人が確認した社内Q&A」として配信される
'   ・直前まで残っていた別の質問の出典へ、無関係な感謝状が飛ぶ
' という取り返しのつかない誤爆になる(実機第3報 RC2 の危険)。
'
' 判定材料はモードと出典件数、そして直近ターンが成立したかの3つ:
'   ・mode="" … 検索も回答生成もしなかったターン(空質問)
'   ・mode="general" … 一般アシスタント(本棚を通っていない)
'   ・nHits=0 … 根拠となる資料が1件も無い
'   ・GroundingAllowed=False … 資料は引いたが回答が成立しなかったターン
'     (API失敗・聞き返し)。2026-08-16 R33H M2。
' 【この4つ目が要る理由】ここの契約は元々「聞き返しは mode="" で入って
' くる」と書いていたが、その遮断は W5-21 が mLastMode="" に頼っていたもので、
' F7 がモード名を戻した時点で外れた。以後 nHits>0 の聞き返し・API失敗ターンは
' 真理表を素通りし、🔴/🤔(modAsk:504/518)がその質問を部内へ発信していた
' (緑だけは mLastCleanAnswer 非空の併記があって助かっていた)。
' 窓口(modAsk.CanShareInsight)が呼ぶのは下の EmitInsightAllowed で、
' 真理表 ShouldEmitInsight そのものは【モードと件数だけの純関数】のまま残す
' ―― 状態を混ぜると、この関数を固定しているゴールデン(modTestsPure11)が
' 「直前にどのテストが走ったか」で答えを変えるようになるため。
' いずれか1つでも当てはまれば発信しない。個人統計(selfsolve_total・
' 節約時間・usage_log)は「解決した」という事実そのものなので、この判定とは
' 無関係に必ず加算する(呼び出し側の責任)。
' AnsweredMode(R33 W5-21)は R33H F7 の NoteAnswered / GroundingAllowed へ
'   置き換えられ、src からの呼び出し元が0件になったため R33H Fix波3 で撤去した。
'   production が呼ばない関数をテストが固定していると「回帰テストがある」と
'   見えるだけで、実装をどう変えても誰も困らない ―― このラウンドの主題
'   (テストが嘘をつく)そのものなので、対象ごと消す。
'
' NoteAnswered - R33H F7。直近ターンの成否を控えて、モード名をそのまま返す。
'   modAsk の Done: から1行で呼ぶための形(戻り値を mLastMode へ入れるので、
'   凍結モジュール側は代入1行のまま=W5-21 の手術跡の上に重ねる)。
Public Function NoteAnswered(ByVal okFlag As Boolean, ByVal modeName As String) As String
    mGrounded = okFlag
    mLastAnsweredMode = modeName   ' R34 B1: 出典突合の後付けがモードを見分ける材料
    NoteAnswered = modeName
End Function

' AnsweredModeName - 直近ターンの回答モード名(R34 B1)。表示には使わない
'   (表示側の単一情報源は従来どおり modAsk.mLastMode)。テストからの可観測化用。
Public Function AnsweredModeName() As String
    AnsweredModeName = mLastAnsweredMode
End Function

' GroundingAllowed - 直近ターンで根拠表示を出してよいか。
'   表示側(modAppAct.DrawConfidence / modPeek.RenderCitations)がこれを見る。
Public Function GroundingAllowed() As Boolean
    GroundingAllowed = mGrounded
End Function

' ============================================================================
' R34 B1: 機械的出典突合を ⚡すぐ聞く / 🔍しっかり調べる にも効かせる
' ============================================================================
' 出典タグの照合(回答本文に書かれた [本棚:…] が実際の検索結果に在るか)は
' 入念モードだけが通っていた(modAskThorough.AnnotateAgainstHits)。
' 幻覚の出典は「資料に書いてある」と読ませてしまうため、回答そのものより重い。
' quick/deep も同じ門をくぐらせる。
'
' 二重付与の禁止: thorough は modAskThorough/modAskMulti/modAskGlobal が
' 生成の内側で既に注記済みなので、ここでは絶対に走らせない(走らせると
' 「(出典確認できず)」が2つ並ぶ)。general/空質問は呼び出し側(modApp の
' grounded ガード)で来ないうえ、LastHitCount()=0 でも止まる。
'
' 置き場が modMode なのは modApp が容量逼迫(R34 B0 で作った余地は配線1行ぶん)
' で、かつ「このターンが何のモードだったか」を既に持っているのがここだから。
Public Function ShouldAnnotate(ByVal modeName As String) As Boolean
    ' Normalize は使わない("" や "general" まで quick へ丸めてしまうため)。
    Dim m As String: m = LCase$(Trim$(modeName))
    ShouldAnnotate = (m = MODE_QUICK Or m = MODE_DEEP)
End Function

' ----------------------------------------------------------------------------
' DisplayNotes - 回答に付ける【表示専用】の注意書きを1か所で組み立てる
'                (2026-09-10 R44。純関数=LibreOffice実行テストで組合せを固定)
' ----------------------------------------------------------------------------
' なぜ1か所に集めたか:
'   注意書きは2種類ある。(1)低関連度の⚠(条件つき)と(2)常設のガード(※)。
'   別々の場所で足すと「⚠の回だけ※が消える」「両方付いて前置きが二重になる」
'   といった組み合わせ事故が起きる。順序と重複の排除をここで確定させる。
'
' 引数は素の値だけにしてある(Hit()もconfigも見ない)。R4準拠にして実行テストで
' 固定するためで、maxScore の算出と config の読みは呼び出し側(modAskRetrieve)。
' modClarify.HasScoreDispersion を純関数側へ寄せたのと同じ理由。
'
' warnOn=False は「⚠を出さない」であって「ガードも出さない」ではない。俯瞰
' ターン(modAskGlobal.WasGlobalTurn)は score=0 が正しい値なので⚠は事実と逆に
' なるが、「出典を開いて確かめてから使う」はどのターンでも等しく正しい。
'
' ガードは【末尾】に置く。先頭に置くと LeadParaIndex(結論を12pt太字にする段落
' 判定)が毎回ガード文を掴み、画面で一番大きい文字が注意書きになる(R43 の⚠と
' 📜で実際に起きた型)。末尾なら結論強調の経路に一切触れない。
Public Function DisplayNotes(ByVal result As String, ByVal maxScore As Double, _
                             ByVal warnOn As Boolean, ByVal threshold As Double, _
                             ByVal guardOn As Boolean) As String
    Dim s As String: s = result
    If guardOn And LenB(Trim$(s)) > 0 Then
        ' 二重付与の防止。復元表示や再描画で同じ本文を通しても増えない。
        If InStr(1, s, GuardNoteText(), vbBinaryCompare) = 0 Then
            s = s & vbLf & vbLf & GuardNoteText()
        End If
    End If
    If warnOn And threshold > 0# Then
        If maxScore < threshold Then s = LowHitNoteText() & vbLf & vbLf & s
    End If
    DisplayNotes = s
End Function

' 常設ガードの文言(単一情報源)。装飾側 modLiveStyle は行頭の「※」で拾うので、
' 先頭文字を変えるときは modLiveStyle.StyleAnswerParas も同時に直すこと。
Public Function GuardNoteText() As String
    GuardNoteText = ChrW(&H203B) & " AIが資料から作った回答です。" & _
        "使う前に出典を開いて原文を確かめてください。"
End Function

' 低関連度の注意書き(単一情報源)。先頭の⚠は modLiveStyle.LeadParaIndex が
' 「結論ではない前置き」と判定する印を兼ねているので、先頭文字を変えない。
Public Function LowHitNoteText() As String
    LowHitNoteText = ChrW(&H26A0) & ChrW(&HFE0F) & _
        " 手元の資料との関連が薄い可能性があります。回答は参考程度にご覧ください。"
End Function

' 実況の末尾に付ける「待ってよいことの根拠」と、待てなくなったときの出口
' (単一情報源・R48)。
' ----------------------------------------------------------------------------
' R48 で判明したこと: リボンへの1回の呼び出しは config llm_wait_sec(既定1800)
' まで戻らず、こちら側に打ち切る手段が無い(Application.Run の内側に DoEvents が
' 無く ESC の窓が開かない。仮に届いても modGateway の ErrHandler が Err 18 を
' 握り潰す)。さらにリボンは wait 超過で同じ要求を【再送】し、回数上限が無い。
' つまり待ち続けても返らない場面が原理的に存在し、利用者に残る出口は Excel の
' 終了だけになる。憲章 §3-1「押せるものは必ず反応する」に照らして、無反応の
' まま放置するのではなく【出口を明示する】。打ち切り機構そのものの新設は
' 設計から作る話なので別ラウンド(裁定書 R48 §7)。
' modLive.TailPos が " ※" を末尾装飾の開始位置として拾うので、先頭は必ず ※。
Public Function WorkingNote() As String
    WorkingNote = " " & ChrW(&H203B) & "応答なし表示でも処理中" & _
        "(30分たっても変わらないときは Excel を終了して構いません)"
End Function

' PageTagPart - 出典タグの「p.N」部分(単一情報源・純関数・R41 §1 A)。
'   拡張子(modUtil.ExtOf の小文字)が xlsx/xlsm/xls/xlsb なら Excel 由来
'   (ページではなくシートの通し番号)なので " シートN"、それ以外は " p.N"。
'   page が 0 でも空にしない(CiteTagFrom("議事録.txt", 0, "") の既存挙動=
'   modTestsPure38 が固定する [本棚:議事録.txt p.0] を保つ)。page<=0 を
'   空文字にする判断は呼び出し側(modLive.PageLabel など表示専用の窓口)に
'   任せる。
Public Function PageTagPart(ByVal srcName As String, ByVal page As Long) As String
    Dim ext As String: ext = LCase$(modUtil.ExtOf(srcName))
    If ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Or ext = "xlsb" Then
        PageTagPart = " シート" & page
    Else
        PageTagPart = " p." & page
    End If
End Function

' CiteTagFrom - 出典タグ1件を組み立てる。書式は modPrompts.SourceTag と同一
'   ([本棚:資料名 p.N] / [パック(作成者名):資料名])。Hit 型ではなく素の3値から
'   作るのは、modAsk の読み取り専用アクセサ(LastHitSource/LastHitPage/
'   LastHitOrigin)が返すのがこの3つだけのため。書式を変えるときは
'   modPrompts.SourceTag と必ず同時に直す(modTestsPure38 が一致を固定する)。
'   R41 §1 A: Excel 由来は p.N ではなく シートN(PageTagPart 経由)。
Public Function CiteTagFrom(ByVal srcName As String, ByVal pageNo As Long, _
                            ByVal originTag As String) As String
    If LCase$(Left$(originTag, 5)) = "pack:" Then
        CiteTagFrom = "[パック(" & Mid$(originTag, 6) & "):" & srcName & "]"
    Else
        CiteTagFrom = "[本棚:" & srcName & PageTagPart(srcName, pageNo) & "]"
    End If
End Function

' CiteIndexAdd - 突合表("|タグ||タグ|" の連結)へ1件足す。重複は入れない。
'   modAskThorough.CiteIndexFrom の hits() 無し版(1件ずつ積む形)。正規化も
'   あちらと同じ modAskThorough.NormalizeCiteTag を通す(両側が同じ処理を
'   通らないと空白の有無だけで全件不一致になる)。
Public Function CiteIndexAdd(ByVal idx As String, ByVal citeTag As String) As String
    CiteIndexAdd = idx
    Dim k As String
    k = modAskThorough.NormalizeCiteTag(citeTag)
    If LenB(k) = 0 Then Exit Function
    If InStr(1, idx, "|" & k & "|", vbTextCompare) > 0 Then Exit Function
    CiteIndexAdd = idx & "|" & k & "|"
End Function

' AnnotateIfNeeded - 直近ターンが quick/deep のときだけ、回答本文の出典タグを
'   modAsk が持つ実際のヒットと突合して「(出典確認できず)」を付ける。
'   それ以外のモード・ヒット0件・タグ0件では ans をそのまま返す(無加工)。
'   mismatchN は戻り値の注記そのものが結果なので呼び出し側では使わない。
Public Function AnnotateIfNeeded(ByVal ans As String) As String
    AnnotateIfNeeded = ans
    If LenB(ans) = 0 Then Exit Function
    If Not ShouldAnnotate(mLastAnsweredMode) Then Exit Function

    Dim n As Long
    On Error Resume Next
    n = modAsk.LastHitCount()
    On Error GoTo 0
    If n < 1 Then Exit Function

    ' R34 F1: 突合表を作るときだけ【生成に渡した件数】まで回す。deep の精読
    ' (modAsk.RunDeepFlow が足す前後チャンク)は回答の根拠として実際に読ませて
    ' いるので、そのページを引いた出典は正当。実ヒットだけで表を作ると、
    ' 正しい引用へ「(出典確認できず)」を付けてしまう。入念モードも拡張後の
    ' 集合で突合している(modAskThorough:165)ので意味論は揃っている。
    ' 入口の門は LastHitCount のまま(一般モード・空質問ターンの残留を通さない)。
    Dim genN As Long
    On Error Resume Next
    genN = modAsk.LastGenHitCount()
    On Error GoTo 0
    If genN < n Then genN = n

    ' アクセサは0始まり(modAsk のコメント参照)。
    Dim idx As String
    Dim i As Long
    On Error Resume Next
    For i = 0 To genN - 1
        idx = CiteIndexAdd(idx, CiteTagFrom(modAsk.LastHitSource(i), _
                                            modAsk.LastHitPage(i), _
                                            modAsk.LastHitOrigin(i)))
    Next i
    On Error GoTo 0
    If LenB(idx) = 0 Then Exit Function

    Dim mismatchN As Long
    Dim outS As String
    outS = modAskThorough.AnnotateCitations(ans, idx, mismatchN)
    If LenB(outS) = 0 Then Exit Function   ' 想定外の空戻りで本文を消さない
    AnnotateIfNeeded = outS

    ' R41 §3 C3: 出典不一致の件数を捨てずに記録する(mismatchN>0 のときだけ)。
    ' ログ書き込みの失敗で回答そのものを壊さないよう OERN で包む。
    If mismatchN > 0 Then
        On Error Resume Next
        modLog.LogUsage "cite_mismatch", mLastAnsweredMode, "n=" & mismatchN
        On Error GoTo 0
    End If
End Function

' ----------------------------------------------------------------------------
' UseNeighborExpand - 前後チャンク結合(modAskFocus.NeighborExpand)を検索の
'   最後に掛けてよいモードか(R34 B3)。
' ----------------------------------------------------------------------------
' 入念は modAskThorough / modAskMulti が生成の内側で自前に呼ぶので、ここが
' True を返すと同じチャンクが二度足される。すぐ聞くは速さが目的なので対象外
' (GOの範囲外)。したがって deep だけ。Normalize を通すので、未知の文字列や
' 空文字は quick へ丸まって False になる。
Public Function UseNeighborExpand(ByVal mode As String) As Boolean
    UseNeighborExpand = (Normalize(mode) = MODE_DEEP)
End Function

' ----------------------------------------------------------------------------
' 旧UI(modUIMain.RenderAnswer)の出典欄と状態セルの文言(2026-08-16 R33H M1)
' ----------------------------------------------------------------------------
' Nexus側は3本(modAppAct/modPeek/modMentor)とも GroundingAllowed の門を
' くぐるようになったが、旧3画面(nexus_ui=FALSE)には門が1つも無く、API失敗
' (#ERR)・聞き返しのターンが【通常回答の枝】へ流れていた ―― 「AIとの通信に
' 失敗しました」の直下に「📖 この回答のもと: …」と「状態: 回答ができました。
' 出典もあわせてご確認ください。」が並ぶ(W5-21 以前より悪い)。
' 実体をここへ置くのは modUIMain が残150字で分岐を書けないため(憲章§4-6)。
' 門は引数で受け取る純関数にする(モジュール状態を読むと、この2本を固定する
' ゴールデンが「直前にどのテストが走ったか」に依存してしまう)。
'   allowed: 直近ターンで根拠表示を出してよいか(modMode.GroundingAllowed)。
Public Function AnswerSourcesText(ByVal allowed As Boolean, ByVal labels As String) As String
    If Not allowed Then Exit Function
    If LenB(Trim$(labels)) = 0 Then Exit Function
    AnswerSourcesText = ChrW(&HD83D) & ChrW(&HDCD6) & " この回答のもと: " & labels
End Function

' 不成立ターンの状態セル。API失敗と聞き返しの両方に当てはまる言い方にする
' (「回答を作れませんでした」は聞き返しのターンでは嘘になる)。
Public Function AnswerStatusText(ByVal allowed As Boolean) As String
    If allowed Then
        AnswerStatusText = "状態: 回答ができました。出典もあわせてご確認ください。"
    Else
        AnswerStatusText = "状態: 上のメッセージをご確認ください。"
    End If
End Function

Public Function ShouldEmitInsight(ByVal mode As String, ByVal nHits As Long) As Boolean
    Dim m As String: m = LCase$(Trim$(mode))
    If LenB(m) = 0 Then Exit Function
    If m = "general" Then Exit Function
    ShouldEmitInsight = (nHits > 0)
End Function

' EmitInsightAllowed - 発信してよいかの【窓口】(2026-08-16 R33H M2)。
'   真理表に F7 の門を畳んだもの。modAsk.CanShareInsight はこの1本だけを
'   呼ぶ(凍結モジュール側の変更を1行に留め、以後この門を動かすときに
'   modAsk を開けなくて済む形にする)。理由は上の見出しコメント。
Public Function EmitInsightAllowed(ByVal mode As String, ByVal nHits As Long) As Boolean
    If Not GroundingAllowed() Then Exit Function
    EmitInsightAllowed = ShouldEmitInsight(mode, nHits)
End Function

' サブクエリ数。入念は角度の数がそのまま精度になる(実測 R@10 90%→97%)。
Public Function SubQueryCount(ByVal mode As String, ByVal normalN As Long, _
                              ByVal thoroughN As Long) As Long
    If Normalize(mode) = MODE_THOROUGH Then
        SubQueryCount = thoroughN
    Else
        SubQueryCount = normalN
    End If
    If SubQueryCount < 1 Then SubQueryCount = 1
End Function

' ============================================================================
' 確信度(R46 A-4) - 検索スコアから切り離し、数えられる事実だけで決める
' ----------------------------------------------------------------------------
' 実機報告:「MSAD、三井住友の株価は？って聞いたら緑の〇印が出た、もちろん
' そんな資料は入れてない」「△マークはでない」「まるで高確率ガチャ」。
'
' 旧実装(modAsk.LastConfidence)は【しきい値1本】しか持っていなかった:
' 上位ヒットのうち score >= 0.55 が2件以上なら🟢、1件なら🟡、0件なら🔴。
' 3値に見えて実際に分けているのは「0.55 を超えた件数」だけで、似た文書は
' まとまって上位に来るため「ちょうど1件」はほとんど起きない=実質2値。
' 「△が出ない」はこの構造の帰結だった。
' さらに 0.55 は【埋め込みの絶対スケールに依存する数字】で、固有名詞が
' 一致するだけで超える。「三井住友の株価」は主語(三井住友)だけで 0.55 を
' 超え、述語(株価)が本棚に無いことを一切見ていなかった。
'
' R46 はスケール非依存の【比】だけで判定する。埋め込みが変わっても意味が
' 変わらないので、実機での再校正に依存しない(旧実装が壊れた理由そのもの):
'   coverage … 質問から抜いた「効く語」のうち、採用チャンクに実在する割合
'   flatness … 1位スコア ÷ 上位の平均。答えがあるときは1位が突出し、
'               無いときは全部が横並びの中くらいになる
'
' 【実測(2026-09-11)】教えてBOX 10,996かけら。「火災保険(法人)」440かけらを
' 本棚から丸ごと抜き、その分野の実際の質問87件を投げる=「いかにも社内資料に
' ありそうで、実は入っていない」という実務で一番危ない形で測った。
' 答えがある質問362件との比較:
'                coverage 中央   flatness 中央
'   答えがある     1.000          1.801
'   答えが無い     0.750          1.170
' 採用したしきい値での3値の出方:
'   答えがある  🟢69.6% 🟡29.0% 🔴 1.4%
'   答えが無い  🟢 5.7% 🟡12.6% 🔴81.6%
' 「三井住友の株価」は coverage 0.33(株価が当たらない)・flatness 1.05 で🔴、
' 「今日の食堂のメニュー」は coverage 0.67 で🔴 になることを実データで確認した。
'
' margin(1位÷2位)も測ったが flatness と相関していて分離に寄与しなかったため
' 採用しない(効かない設定を残さない)。
' ============================================================================

' CoverageOf - 質問の「効く語」が採用チャンクにどれだけ実在するか(0〜1)。
'   語の抽出と照合は modSparse が唯一の実装(同じ規則を2箇所に書かない)。
'   語が1つも取れない質問(記号だけ等)は判定材料が無いので 0 を返す。
Public Function CoverageOf(ByVal q As String, ByVal bodyText As String) As Double
    Dim keys As String: keys = modSparse.DistinctiveKeys(q)
    If LenB(keys) = 0 Then Exit Function
    Dim total As Long: total = UBound(Split(keys, "|")) + 1
    If total < 1 Then Exit Function
    CoverageOf = modSparse.ExactHitCount(keys, bodyText) / CDbl(total)
End Function

' FlatnessOf - 1位スコア ÷ 上位nの平均。1.0に近いほど「全部が同じくらい
'   似ている」=特定の資料が当たっていない。平均が0以下なら1(=平坦)を返す。
Public Function FlatnessOf(ByRef scores() As Double, ByVal n As Long) As Double
    FlatnessOf = 1#
    If n < 1 Then Exit Function
    ' top / sum は VBA の予約語・組み込みと衝突するので使わない(lint が捕捉)。
    Dim topScore As Double, acc As Double
    Dim i As Long
    For i = 1 To n
        If scores(i) > topScore Then topScore = scores(i)
        acc = acc + scores(i)
    Next i
    If acc <= 0# Then Exit Function
    Dim mean As Double: mean = acc / CDbl(n)
    If mean <= 0# Then Exit Function
    FlatnessOf = topScore / mean
End Function

' ConfidenceLevel - 2=根拠あり / 1=部分的 / 0=乏しい(純関数)。
'   しきい値は呼び出し側から渡す(config 読みを純関数に持ち込まない)。
Public Function ConfidenceLevel(ByVal cov As Double, ByVal flat As Double, _
                                ByVal covGreen As Double, ByVal flatGreen As Double, _
                                ByVal covAmber As Double) As Long
    If cov >= covGreen And flat >= flatGreen Then
        ConfidenceLevel = 2
    ElseIf cov >= covAmber Then
        ConfidenceLevel = 1
    End If
End Function

' ConfidenceOf - config を読んで ConfidenceLevel を呼ぶ窓口。modAsk(凍結・
'   残り僅か)から1行で呼べる形にしてある。
Public Function ConfidenceOf(ByVal q As String, ByVal bodyText As String, _
                             ByRef scores() As Double, ByVal n As Long) As Long
    Dim cg As Double, fg As Double, ca As Double
    cg = 1#: fg = 1.6: ca = 0.9
    On Error Resume Next
    cg = modConfig.GetLong("conf_cov_green_x100", 100) / 100#
    fg = modConfig.GetLong("conf_flat_green_x100", 160) / 100#
    ca = modConfig.GetLong("conf_cov_amber_x100", 90) / 100#
    On Error GoTo 0
    ConfidenceOf = ConfidenceLevel(CoverageOf(q, bodyText), FlatnessOf(scores, n), cg, fg, ca)
End Function

' GuardOnly - 常設ガード(※)だけを付ける窓口(R46)。検索が1件も引けなかった回と
'   一般アシスタントのように、低関連度の⚠を出す材料(スコア)が無い経路で使う。
'   config 読みをここに置くのは、凍結で残り僅かの modAsk から1行で呼ぶため。
Public Function GuardOnly(ByVal result As String) As String
    GuardOnly = DisplayNotes(result, 0#, False, 0#, _
        modConfig.GetBool("answer_guard_note", True))
End Function
