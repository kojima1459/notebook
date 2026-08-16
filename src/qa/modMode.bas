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
            Description = "会話の流れの中を深く。社内ナレッジ検索は直前までに使った資料へ" & _
                          "絞って掘り下げ、一般アシスタントは前提と論点を整理してから" & _
                          "判断基準・具体例・注意点まで踏み込みます。1～2分ほど。"
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
    NoteAnswered = modeName
End Function

' GroundingAllowed - 直近ターンで根拠表示を出してよいか。
'   表示側(modAppAct.DrawConfidence / modPeek.RenderCitations)がこれを見る。
Public Function GroundingAllowed() As Boolean
    GroundingAllowed = mGrounded
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
