Attribute VB_Name = "modConvBridge"
Option Explicit

' ============================================================================
' modConvBridge - 会話メモリの橋渡し(2026-08-11 R26-2 波B)
' ----------------------------------------------------------------------------
' 課題(調査確定済み・spec_20260810_R26 §2): RAG側の直前1往復は
' nexus_ask_prevu/preva、一般アシスタント側は nexus_gen_prevu/preva に別々の
' キーで永続化されており、モードを跨ぐと相手側からは見えない。「文脈が
' 切れる」の正体はここで、切替時に何かを破棄しているわけではなく、
' 単に別モードの記憶を読まないだけである。
'
' 方針: 切替元の直前1往復(質問+回答)だけを切替先のキーへコピーする
' 「橋渡し」。凍結モジュール(modAsk)には一切触れず、modState のキー経由の
' 読み書きだけで完結させる(§0設計原則: 実現可能性は調査済み)。
'   RAG→一般: 次の AskGeneral 呼び出し(follow-upに関係なく毎回)が
'             mGenPrevU/mGenPrevA を必ず読むため、確実に効く。
'   一般→RAG: modAsk.Answer(単発質問)は毎回 prevU/prevA に空文字列を
'             明示的に渡す実装のため、橋渡しした内容が効くのは
'             「続けて質問」を武装したときの AskFollowup 経由に限られる
'             (modAsk.CanFollowup が遅延ロードで nexus_ask_prevu/preva を
'             読み直す条件=モジュール変数がまだ空のときだけ)。凍結制約下の
'             現実解であり、この非対称は既知の限界として司令塔へ報告する。
'
' レイヤ配置: このモジュールは modState(基盤層)・modConfig(基盤層)しか
' 読み書きしない「計算」担当に徹する。modAppState への書き込み(UI層の
' モジュール変数 mGenPrevU/mGenPrevA)・modSkin.ShowToast(UI通知)は
' R1(下層は上層を呼ばない)に触れるため、呼び出し元の modAppState.
' BridgeConvMemory(UI層)側で行う。ここは ComputeBridge が計算した結果
' (ByRef の outQ/outA)を返すだけ。
' ============================================================================

Public Const CONFIG_KEY As String = "conv_bridge"

' 橋渡しする文字数の上限(既存記憶の上限作法に合わせる。超過分は先頭を
' 落とす=直近の内容を優先して残す)。1往復あたりの上限(R28 W3-6b以前から据置)。
Private Const MAX_CARRY_CHARS As Long = 4000

' R28 W3-6b: 運搬総量(Q/Aそれぞれの合計)の新クランプ。config
' followup_max_pairs が既定より大きく変更されても、1回の橋渡しで運ぶ量が
' 際限なく伸びない歯止め(既定5往復・1往復あたり最大4000字を目安にした値。
' 実際の会話はもっと短く、通常はここに触れずに5往復とも運ばれる)。
' 超過分は古い往復(配列の後方=";;;"区切りの後ろ)から切り捨てる。
Private Const MAX_CARRY_TOTAL_CHARS As Long = 20000

' 橋渡し済みの回答に付ける出所ヘッダーの接頭辞。この接頭辞で始まっていれば
' 「既に付与済み」とみなし、二重に付け足さない(冪等性)。
Private Const HEADER_PREFIX As String = "【直前の"

' ----------------------------------------------------------------------------
' BridgeEnabled - config conv_bridge(既定TRUE)。
' ----------------------------------------------------------------------------
Public Function BridgeEnabled() As Boolean
    BridgeEnabled = modConfig.GetBool(CONFIG_KEY, True)
End Function

' ----------------------------------------------------------------------------
' ComputeBridge - 切替元→切替先の橋渡し内容を計算する(状態読み取りあり)。
'   fromMode/toMode: "rag" / "normal"(modAppState.CurrentMode() と同じ語彙)。
'   outQ/outA: 橋渡しする質問・回答(切替先のキーへ書く値)。
'   戻り値: 橋渡しする内容があれば True(呼び出し元はこのときだけ書き込み・
'           通知を行う)。conv_bridge=off/同一モード/未知の組合せ/
'           切替元の記憶が空、のいずれも False。
'   実体は ComputeBridgeCore(状態を一切読まない純関数)へ委譲する。
'   config/modStateを読む本関数はLOのヘッドレステスト環境では固定できない
'   ため(configシートが無い環境ではconv_bridge=offを再現できない)、
'   conv_bridge=off等の境界はテストから ComputeBridgeCore を直接呼んで固定する。
' ----------------------------------------------------------------------------
Public Function ComputeBridge(ByVal fromMode As String, ByVal toMode As String, _
                              ByRef outQ As String, ByRef outA As String) As Boolean
    Dim srcQ As String, srcA As String
    Dim dstQ As String, dstA As String
    If fromMode = "rag" And toMode = "normal" Then
        srcQ = modState.LoadState("nexus_ask_prevu", "")
        srcA = modState.LoadState("nexus_ask_preva", "")
        dstQ = modState.LoadState("nexus_gen_prevu", "")
        dstA = modState.LoadState("nexus_gen_preva", "")
    ElseIf fromMode = "normal" And toMode = "rag" Then
        srcQ = modState.LoadState("nexus_gen_prevu", "")
        srcA = modState.LoadState("nexus_gen_preva", "")
        dstQ = modState.LoadState("nexus_ask_prevu", "")
        dstA = modState.LoadState("nexus_ask_preva", "")
    End If
    ComputeBridge = ComputeBridgeCore(fromMode, toMode, BridgeEnabled(), _
                                      srcQ, srcA, dstQ, dstA, CarryMaxPairs(), outQ, outA)
End Function

' 切替先に残す往復数の上限(会話履歴の既存作法と同じ config を使う。
' 橋渡しだけ別の上限を持つと「深掘りは3往復なのに切替後だけ1往復」に
' なり、利用者から見て記憶の深さが操作で変わる)。
' (名前が maxPairs 引数と衝突しないよう CarryMaxPairs。lintの手続き名衝突検査)
Private Function CarryMaxPairs() As Long
    CarryMaxPairs = modConfig.GetLong("followup_max_pairs", 3)
End Function

' ----------------------------------------------------------------------------
' ComputeBridgeCore - ComputeBridgeの純関数本体(modState/modConfigを一切
'   読まない)。enabled/srcQ/srcAを引数で受け取るだけなので、config/シートに
'   依存せず conv_bridge=off・記憶なし等の境界をテストで固定できる
'   (modGenPipe.ShouldRunVerifyLoop と同型の設計判断)。
' ----------------------------------------------------------------------------
' R26H F3(M-1): 旧実装は切替先の記憶を橋渡しの1往復で【置き換えて】いた。
'   一般アシスタントで3往復話したあと社内ナレッジ検索へ寄り道して戻ると、
'   その3往復が消えて1往復だけになる=「引き継ぎ」が実際には既存の記憶の
'   破棄になっていた。切替先の記憶は消さず、橋渡しの1往復を【先頭へ差し込む】
'   (dstQ/dstA が切替先の現在の記憶。maxPairs で上限まで丸める)。
'   同じ往復が既に先頭にあるとき(切替を往復させただけ)は差し込まず False を
'   返す=同じ内容が2件並ぶことも、そのたびにトーストが出ることも無い。
' R28 W3-6b: 直前1往復(FirstPair)だけの運搬から、保持している全往復
'   (最大maxPairs件・BuildCarryPairs)の運搬へ拡張。二重挿入防止の判定は
'   従来どおり「先頭往復の一致」で行い(往来を繰り返しただけでは増殖しない)、
'   一致していなければ全往復をまとめて先頭へ差し込む。
Public Function ComputeBridgeCore(ByVal fromMode As String, ByVal toMode As String, _
                                  ByVal enabled As Boolean, _
                                  ByVal srcQ As String, ByVal srcA As String, _
                                  ByVal dstQ As String, ByVal dstA As String, _
                                  ByVal maxPairs As Long, _
                                  ByRef outQ As String, ByRef outA As String) As Boolean
    outQ = "": outA = ""
    If fromMode = toMode Then Exit Function
    If Not enabled Then Exit Function

    Dim srcName As String
    If fromMode = "rag" And toMode = "normal" Then
        srcName = ModeDisplayName("rag")
    ElseIf fromMode = "normal" And toMode = "rag" Then
        srcName = ModeDisplayName("normal")
    Else
        Exit Function   ' 未知のモード値はフェイルセーフで何もしない
    End If

    If LenB(FirstPair(srcQ)) = 0 And LenB(FirstPair(srcA)) = 0 Then Exit Function   ' 引き継ぐ記憶が無い

    Dim carryQ As String, carryA As String
    carryQ = BuildCarryPairs(srcQ, maxPairs, False, "")
    carryA = BuildCarryPairs(srcA, maxPairs, True, srcName)
    If LenB(carryQ) = 0 And LenB(carryA) = 0 Then Exit Function

    If AlreadyAtHead(FirstPair(carryQ), FirstPair(carryA), dstQ, dstA) Then Exit Function

    outQ = InsertAtHead(carryQ, dstQ, maxPairs)
    outA = InsertAtHead(carryA, dstA, maxPairs)
    ComputeBridgeCore = True
End Function

' ----------------------------------------------------------------------------
' BuildCarryPairs - ";;;"区切りの往復履歴から先頭(最新)maxPairs件を取り出し、
'   1往復あたりMAX_CARRY_CHARS・運搬総量MAX_CARRY_TOTAL_CHARSの2段で切り詰めて
'   ";;;"で結び直す(純関数)。総量超過は古い往復(配列の後方)から切り捨てる。
'   withHeader: 回答側だけTrue(WithBridgeHeaderを各往復へ付ける)。
' ----------------------------------------------------------------------------
Private Function BuildCarryPairs(ByVal src As String, ByVal maxPairs As Long, _
                                 ByVal withHeader As Boolean, ByVal srcName As String) As String
    If LenB(src) = 0 Then Exit Function

    Dim mx As Long: mx = maxPairs
    If mx < 1 Then mx = 1

    Dim parts() As String
    parts = Split(src, ";;;")
    Dim n As Long: n = UBound(parts) - LBound(parts) + 1
    If n > mx Then n = mx

    Dim out As String, total As Long
    Dim i As Long
    For i = LBound(parts) To LBound(parts) + n - 1
        Dim item As String
        item = TruncateTail(CStr(parts(i)), MAX_CARRY_CHARS)
        If withHeader Then item = WithBridgeHeader(srcName, item)

        Dim addLen As Long: addLen = Len(item) + IIf(LenB(out) > 0, 3, 0)
        If LenB(out) > 0 And total + addLen > MAX_CARRY_TOTAL_CHARS Then Exit For

        If LenB(out) > 0 Then out = out & ";;;"
        out = out & item
        total = total + addLen
    Next i
    BuildCarryPairs = out
End Function

' ----------------------------------------------------------------------------
' AlreadyAtHead - 橋渡しする1往復が、切替先の記憶の先頭に既にあるか(純関数)。
'   質問と回答の【両方】が一致したときだけ True。片方だけの一致で止めると、
'   同じ質問を2モードで聞いたときに回答の引き継ぎだけが落ちる。
'   R28 W3-6b: 複数往復運搬でも判定は「先頭往復の一致」のまま
'   (呼び出し側がcarryQ/carryAの先頭要素=FirstPairを渡す)。
' ----------------------------------------------------------------------------
Public Function AlreadyAtHead(ByVal carryQ As String, ByVal carryA As String, _
                              ByVal dstQ As String, ByVal dstA As String) As Boolean
    If LenB(dstQ) = 0 And LenB(dstA) = 0 Then Exit Function
    AlreadyAtHead = (FirstPair(dstQ) = carryQ) And (FirstPair(dstA) = carryA)
End Function

' ----------------------------------------------------------------------------
' InsertAtHead - ";;;"区切り履歴の先頭へ差し込み、上限まで丸める(純関数)。
'   item は1件でも複数件(";;;"区切り)でもよい(R28 W3-6b: BuildCarryPairsが
'   複数往復を渡すようになったため、item側の件数を問わず動く実装のまま据置)。
'   丸めは modFollowup.KeepNewestPairs(既存の単一情報源)へ委譲する。
'   maxPairs<=0 は「履歴を持たない」設定なので、差し込んだ最新1件だけを残す
'   (橋渡しを押した本人の操作を黙って無効化しない)。
' ----------------------------------------------------------------------------
Public Function InsertAtHead(ByVal item As String, ByVal joined As String, _
                             ByVal maxPairs As Long) As String
    Dim mx As Long: mx = maxPairs
    If mx < 1 Then mx = 1
    If LenB(joined) = 0 Then
        InsertAtHead = item
        Exit Function
    End If
    InsertAtHead = modFollowup.KeepNewestPairs(item & ";;;" & joined, mx, ";;;")
End Function

' ----------------------------------------------------------------------------
' TruncateTail - 文字列を末尾から limitChars 文字だけ残す(超過分は先頭を
'   落とす=直近側を優先)。全角安全: 切り出した先頭が低位サロゲート
'   (&HDC00~&HDFFF)なら対になる高位サロゲートを道連れに1字余分へ落とす
'   (modUtil.SafeLeft の「末尾の片割れを落とす」発想を先頭側へ適用)。
' ----------------------------------------------------------------------------
Public Function TruncateTail(ByVal s As String, ByVal limitChars As Long) As String
    Dim lim As Long: lim = limitChars
    If lim < 0 Then lim = 0
    If Len(s) <= lim Then
        TruncateTail = s
        Exit Function
    End If
    If lim = 0 Then Exit Function   ' 戻り値は既定の空文字列のまま

    Dim startPos As Long: startPos = Len(s) - lim + 1
    Dim t As String: t = Mid$(s, startPos)
    Dim code As Long: code = AscW(Left$(t, 1))
    If code < 0 Then code = code + 65536
    If code >= &HDC00& And code <= &HDFFF& Then t = Mid$(t, 2)
    TruncateTail = t
End Function

' ----------------------------------------------------------------------------
' FirstPair - ";;;"区切り履歴の先頭要素(最新の1件)だけを取り出す。
' ----------------------------------------------------------------------------
Public Function FirstPair(ByVal pairsText As String) As String
    Dim p As Long: p = InStr(pairsText, ";;;")
    If p > 0 Then
        FirstPair = Left$(pairsText, p - 1)
    Else
        FirstPair = pairsText
    End If
End Function

' ----------------------------------------------------------------------------
' WithBridgeHeader - 回答テキストの先頭へ出所ヘッダー(【直前の○○での文脈】)
'   を付ける。既にこの種のヘッダーで始まっていれば付け直さない(冪等性:
'   何度橋渡ししてもヘッダーが積み重ならない)。
' ----------------------------------------------------------------------------
Public Function WithBridgeHeader(ByVal fromModeName As String, ByVal answerText As String) As String
    If Left$(answerText, Len(HEADER_PREFIX)) = HEADER_PREFIX Then
        WithBridgeHeader = answerText
        Exit Function
    End If
    WithBridgeHeader = HEADER_PREFIX & fromModeName & "での文脈】" & vbLf & answerText
End Function

' モード内部値("rag"/"normal")→表示名。modApp.ModeCaption の文言(絵文字抜き)
' と揃える(利用者がヘッダーの由来を見たとき、いつも見ている名前と一致させる)。
Private Function ModeDisplayName(ByVal modeKey As String) As String
    If modeKey = "normal" Then
        ModeDisplayName = "一般アシスタント"
    Else
        ModeDisplayName = "社内ナレッジ検索"
    End If
End Function
