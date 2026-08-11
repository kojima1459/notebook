Attribute VB_Name = "modGateway"
Option Explicit

' ============================================================================
' modGateway - 社内AIリボン「リボンちゃん」唯一の窓口
' ----------------------------------------------------------------------------
' 役割:
'   ChatGPT() / GetEmbeddings() へのApplication.Run呼び出しは、このモジュール
'   の中だけで行う(MASTER_SPEC §3 R3)。他のモジュールは絶対に直接
'   Application.Run("ChatGPT", ...) 等を書かない。opt層も自前でRunせず
'   TryRibbonRun経由にすることで、リボンの引数規約が変わってもここ1箇所の
'   修正で済むようにする。
'
' 設計判断:
'   ・呼び出し規約は確定済み: V2 src/chatbot_v2/modRibbonGateway.bas /
'     modEmbeddings.bas の実証実績に加え、担当部署の確定台帳
'     RIBBON_API_CONFIRMED.md §0(互換後方追加)で裏付けられた(下記
'     コメント参照)。
'   ・mock_llm=TRUE(configシート)のときはリボンを一切呼ばず、決定的な
'     ダミー応答/ダミーベクトルを返す。これにより取込→検索→回答の
'     画面フロー全体を、社内ネットワーク外(自宅PC/Mac等)でも
'     動作確認できる。
'   ・RibbonAvailableは公式のアドイン検出作法(RIBBON_API_CONFIRMED.md §1
'     末尾・裁定D2)に従い、Application.AddIns をループして
'     アドイン名の部分一致+Installed で判定する(API呼び出し不要・即時)。
'     旧実装の「実際に軽量なChatGPT呼び出しで確かめる」プローブ方式は
'     公式作法の確定に伴い廃止した。アドイン名は config
'     ribbon_addin_name(既定 "リボンちゃん")で可変。判定結果の
'     セッションキャッシュ(mRibbonChecked/mRibbonAvailable)は維持する。
'   ・RunLimitCheckはリボン公式のLimitCheck()(True=続行不可・裁定D3)の
'     唯一の呼び出し口。古いリボンにLimitCheckが無い場合でも利用者を
'     誤ブロックしないよう、エラー時はFalse(=続行可)へ倒す穏当運用。
'   ・mock埋め込みベクトルは modUtil.Fnv1a64Hex をシードにした線形合同法
'     (LCG)で生成する。同じテキスト(正規化後)からは常に同じベクトルが
'     生成される(決定的)。アルゴリズムの詳細は MockEmbedVector 直前の
'     コメントを参照。
' ============================================================================

Private mRibbonChecked As Boolean
Private mRibbonAvailable As Boolean

' R13 F8: 段(step)ごとの所要時間を貯めておくバッファ。1段1行で usage_log へ
' 書くと質問1回で10行前後が積み上がり、2,000行ローテーションを約180問で
' 一周してしまう。そこで貯めるだけ貯めて、質問の終わりに modAsk が1行
' (ask_steps)へまとめて書き出す。書式は STEPBUF_* を参照。
Private mStepBuf As String
' R26H F5: mock査読(入念モード)の周回カウンタ。奇数回=指摘/偶数回=PASS。
Private mMockVerifyRound As Long
' バッファ長の上限。usage_log の detail 1セルに収まり、かつ異常に長い
' step_name が来ても暴走しないための安全弁。
Private Const STEPBUF_MAX As Long = 400

' ----------------------------------------------------------------------------
' CallLLM - ChatGPT()の唯一の呼び出し口。
'   成功: 応答文字列 / 失敗: "#ERR:E02xx:<説明>" で始まる文字列(例外は出さない)
' ----------------------------------------------------------------------------
' ChatGPT位置引数規約は確定済み(V2実証+RIBBON_API_CONFIRMED.md §0:
' 「関数は互換性を保つよう維持される(引数は後方追加)」。公開ページの
' 9引数版に対し、第10・11引数effort/verbosityは互換後方追加):
'   Application.Run("ChatGPT", prompt, "", 0.4, 0, waitSec, model, prevU, prevA,
'                    toolN, effort, verbosity)
' 第9引数toolNには "マイ本棚AI:" & step_name を渡す(裁定D1。管理側ログで
' ツールを識別するため)。第7・8引数prevU/prevAは会話継続用の履歴(台帳§1 #1
' +裁定D11)。複数往復ぶんは「新しい順」に ";;;" 区切りで連結して渡す
' (履歴の保持と連結はmodAsk側の責務)。本関数のprevU/prevAはOptional末尾
' 追加で、既定""=履歴なし=従来と同じ呼び出し(既存呼び出し元は無改修)。
' arg3(Temperature)/arg4(MaxTokens)はDouble/Long型のため ""  を渡すと型不一致
' エラーになる。GPT-5系モデルではこの2つは無視され、代わりにeffort/verbosity
' (第10・11引数)が効く設計になっている(V2の実運用で確認済み)。
' ----------------------------------------------------------------------------
Public Function CallLLM(ByVal prompt As String, ByVal step_name As String, _
                        ByVal effort As String, ByVal verbosity As String, _
                        Optional ByVal model_override As String = "", _
                        Optional ByRef latency_ms As Long = 0, _
                        Optional ByVal prevU As String = "", _
                        Optional ByVal prevA As String = "") As String
    Dim t0 As Double: t0 = Timer
    On Error GoTo ErrHandler

    If modConfig.GetBool("mock_llm", True) Then
        CallLLM = MockLLMResponse(prompt, step_name)
        latency_ms = CLng(modUtilText.ElapsedMsSince(t0))
        LogStepLatency step_name, latency_ms
        Exit Function
    End If

    If Not RibbonAvailable() Then
        modLog.LogError "E0201", "modGateway.CallLLM", "step=" & step_name
        CallLLM = "#ERR:E0201:AIリボンが見つかりません"
        latency_ms = CLng(modUtilText.ElapsedMsSince(t0))
        LogStepLatency step_name, latency_ms
        Exit Function
    End If

    Dim mdl As String
    If LenB(model_override) > 0 Then
        mdl = model_override
    Else
        mdl = modConfig.GetString("recommended_model", "gpt-5.5")
    End If

    Dim eff As String: eff = effort
    Dim vrb As String: vrb = verbosity
    If Not modConfig.GetBool("reasoning_tuning", True) Then
        eff = "": vrb = ""    ' エスケープハッチ: GPT-5以外のモデル運用時などに空送信
    End If

    Dim waitSec As Long: waitSec = modConfig.GetLong("llm_wait_sec", 1200)

    '            text   roleSys Temp MaxTok Wait   model prevU  prevA  toolN                        effort verbosity
    Dim result As Variant
    result = Application.Run("ChatGPT", prompt, "", 0.4, 0, waitSec, mdl, prevU, prevA, "マイ本棚AI:" & step_name, eff, vrb)
    Dim s As String: s = CStr(result)
    latency_ms = CLng(modUtilText.ElapsedMsSince(t0))
    LogStepLatency step_name, latency_ms

    If LenB(s) = 0 Then
        modLog.LogError "E0202", "modGateway.CallLLM", "空応答 step=" & step_name
        CallLLM = "#ERR:E0202:応答が空でした"
        Exit Function
    End If
    If LooksLikeLimitError(s) Then
        modLog.LogError "E0204", "modGateway.CallLLM", "step=" & step_name & " resp=" & modUtil.SafeLeft(s, 200)
        CallLLM = "#ERR:E0204:" & s
        Exit Function
    End If

    CallLLM = s
    Exit Function

ErrHandler:
    latency_ms = CLng(modUtilText.ElapsedMsSince(t0))
    LogStepLatency step_name, latency_ms
    modLog.LogError "E0202", "modGateway.CallLLM", "step=" & step_name & " err=" & Err.Description, Err.Number
    CallLLM = "#ERR:E0202:" & Err.Description
End Function

' ----------------------------------------------------------------------------
' LogStepLatency - LLM 1段ぶんの所要時間を段バッファへ足す(R13-9a → R13 F8)
' ----------------------------------------------------------------------------
' 実機第2報 RC10「入念114秒の内訳が計測不能(CallLLMのlatency_msを誰も記録して
' いない)」への対処。どの段(expand/rerank/quick_draft/deep_draft/deep_verify/
' enrich…)が何ミリ秒かかったかが分からないと、遅さの相談に事実で答えられない。
' 成功・失敗・mock の全経路で必ず1件積む(失敗した段こそ時間を知りたい)。
'
' R13 F8: 初版は1段=1行を usage_log へ書いていたが、1問10行前後になり
' 2,000行で回る usage_log が約180問で一周して、feedback_green など月をまたぐ
' 履歴を押し出していた(modDashStat の前月比が静かに壊れる)。積むだけにして、
' 質問の終わりに modAsk が1行へまとめて書き出す。
' 記録の失敗が呼び出しを壊さないよう、全体を1行スコープの保護で囲む。
Private Sub LogStepLatency(ByVal step_name As String, ByVal ms As Long)
    On Error Resume Next
    mStepBuf = modUtilText.AppendStepBuf(mStepBuf, step_name, ms, STEPBUF_MAX)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ConsumeStepBuf - 段バッファを返して空にする(R13 F8。呼ぶのは modAsk)。
' ----------------------------------------------------------------------------
' 「読んだら消す」にしてあるので、質問と質問のあいだで持ち越さない。
' 取込だけのセッションのように誰も読みに来ない場合でも、STEPBUF_MAX で
' 頭打ちになるだけでメモリは伸び続けない。
Public Function ConsumeStepBuf() As String
    ConsumeStepBuf = mStepBuf
    mStepBuf = ""
End Function

' 埋め込み1バッチぶんの件数と所要時間(R13-9a)。件数は detail と hit_count の
' 両方に入れる(表計算で足し算しやすい形と、目で読める形の両方を残す)。
' 呼ぶのは「実際に外へ問い合わせた層」だけ。単発GetEmbeddingのdirect経路は
' 内部で GetEmbeddingsBatch を通るので、そちらに任せて二重計上しない。
' R13 F8: 取込のバッチは1バッチ=1行のままにする(何百チャンクを1行で表せる
' ので、そもそも usage_log を圧迫しない)。畳んだのは質問側の単発だけ。
' R27 F1-7(実機第12報④): n は【要求した件数】であって成功数ではない。
' 取込が「全件失敗しているのに embed_step だけは毎回同じ n で並ぶ」ため、
' ログからは埋め込みが効いているのか一度も分からなかった(観測不能)。
' 成功数(okCount)を detail へ足す。n>ok なら差が失敗数そのもの。
Private Sub LogEmbedStep(ByVal cnt As Long, ByVal ms As Long, ByVal okCount As Long)
    On Error Resume Next
    modLog.LogUsage "embed_step", "", "n=" & cnt & " ok=" & okCount, ms, cnt
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' GetEmbedding - GetEmbeddings()の唯一の呼び出し口。
'   成功: L2正規化済み配列 / 失敗: 空配列(modUtil.HasVector=False)。E0203を記録。
' ----------------------------------------------------------------------------
' 実引数規約はV2 src/chatbot_v2/modEmbeddings.bas の EmbedText をそのまま
' 踏襲する: GetEmbeddings(text) -> 成功時カンマ結合の数値文字列、
' 失敗時は文字列 "error"(例外は投げない仕様)。入力は EMB_MAX_CHARS=4000字
' で打ち切る(トークン上限に対する安全マージン)。
' ----------------------------------------------------------------------------
Public Function GetEmbedding(ByVal Text As String, Optional ByRef latency_ms As Long = 0) As Double()
    Dim t0 As Double: t0 = Timer
    Dim dim_ As Long: dim_ = modConfig.GetLong("embed_dim", 1536)
    If dim_ < 1 Then dim_ = 1536

    On Error GoTo ErrHandler

    Dim t As String: t = modUtil.NormalizeForHash(Text)
    If LenB(t) = 0 Then
        modLog.LogError "E0203", "modGateway.GetEmbedding", "空文字列は埋め込み不可"
        latency_ms = CLng(modUtilText.ElapsedMsSince(t0))
        Exit Function   ' 空配列を返す
    End If

    If modConfig.GetBool("mock_llm", True) Then
        GetEmbedding = MockEmbedVector(t, dim_)
        latency_ms = CLng(modUtilText.ElapsedMsSince(t0))
        LogStepLatency "emb", latency_ms
        Exit Function
    End If

    ' direct(Azure直接)モード: 単発もバッチ経路(1件)へ集約し、通信実装を
    ' 1本化する(設計書§E-1。検索クエリと保存ベクトルの次元・精度も一致)。
    If LCase$(modConfig.GetString("embed_transport", "ribbon")) = "direct" Then
        Dim one(0 To 0) As String
        one(0) = t
        Dim outCsv() As String
        If GetEmbeddingsBatch(one, outCsv) >= 1 Then
            If LenB(outCsv(0)) > 0 Then
                Dim dvec() As Double
                If modUtil.CsvToVector(outCsv(0), dvec) Then
                    GetEmbedding = dvec
                    latency_ms = CLng(modUtilText.ElapsedMsSince(t0))
                    Exit Function
                End If
            End If
        End If
        latency_ms = CLng(modUtilText.ElapsedMsSince(t0))
        Exit Function   ' 失敗は空配列(バッチ側でE0203記録済み)
    End If

    GetEmbedding = GetEmbeddingRibbonOnly(t, dim_)
    latency_ms = CLng(modUtilText.ElapsedMsSince(t0))
    ' R13 F8: 質問側の埋め込みは多段検索でクエリ本数ぶん走る。1本1行だと
    ' usage_log があっという間に埋まるので、段バッファへ "emb" として畳む
    ' (質問1回ぶんが "emb=6x830" の1トークンにまとまる)。
    LogStepLatency "emb", latency_ms
    Exit Function

ErrHandler:
    latency_ms = CLng(modUtilText.ElapsedMsSince(t0))
    LogStepLatency "emb", latency_ms
    modLog.LogError "E0203", "modGateway.GetEmbedding", "err=" & Err.Description, Err.Number
End Function

' ribbon経由の単発埋め込みの実体。GetEmbedding(公開API)と、direct失敗時の
' フォールバック経路(RibbonEmbedRange)の両方から呼ばれる。embed_transport設定は
' 一切見ない(常にribbonを叩く)ため、direct→ribbonフォールバック時にここへ来ても
' GetEmbeddingへ戻って再びdirect分岐へ入ることがなく、再帰しない。
Private Function GetEmbeddingRibbonOnly(ByVal t As String, ByVal dim_ As Long) As Double()
    If Not RibbonAvailable() Then
        modLog.LogError "E0203", "modGateway.GetEmbedding", "リボン未検出"
        Exit Function
    End If

    Dim raw As Variant
    raw = Application.Run("GetEmbeddings", modUtil.SafeLeft(t, 4000))
    Dim s As String: s = CStr(raw)

    If LenB(s) = 0 Then
        modLog.LogError "E0203", "modGateway.GetEmbedding", "空応答"
        Exit Function
    End If
    If StrComp(Trim$(s), "error", vbTextCompare) = 0 Then
        modLog.LogError "E0203", "modGateway.GetEmbedding", "ribbon returned error"
        Exit Function
    End If
    If InStr(s, ",") = 0 Then
        modLog.LogError "E0203", "modGateway.GetEmbedding", "応答がベクトル形式でない"
        Exit Function
    End If

    Dim vec() As Double
    If Not modUtil.CsvToVector(s, vec) Then
        modLog.LogError "E0203", "modGateway.GetEmbedding", "CSV解析失敗"
        Exit Function
    End If
    ' Plan B: 保存次元(embed_dim)へ切詰め+再正規化(リボンは1536固定のため)
    modUtil.TruncateAndRenorm vec, dim_
    GetEmbeddingRibbonOnly = vec
End Function

' ----------------------------------------------------------------------------
' GetEmbeddingsBatch - 埋め込みのバッチ窓口(設計書§E-1・裁定②)。
'   outCsv(i)=texts(i)のvector_csv(embed_dim切詰め+L2正規化+vector_precision
'   適用済み)。失敗要素は空文字。戻り値=成功件数。
'   mock→決定的擬似ベクトル / direct→Azure配列POST(embed_batch_sizeごと) /
'   ribbon→GetEmbedding単発ループ。directの通信失敗はribbonへフォールバック。
' ----------------------------------------------------------------------------
Public Function GetEmbeddingsBatch(texts() As String, ByRef outCsv() As String) As Long
    GetEmbeddingsBatch = 0
    Dim tB0 As Double: tB0 = Timer      ' R13-9a: 1バッチぶんの所要時間

    Dim lo As Long, hi As Long
    Dim badArr As Boolean: badArr = False
    On Error Resume Next
    lo = LBound(texts)
    hi = UBound(texts)
    badArr = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    If badArr Or hi < lo Then
        ReDim outCsv(0 To 0)
        Exit Function
    End If

    Dim n As Long: n = hi - lo + 1
    ReDim outCsv(0 To n - 1)

    Dim dim_ As Long: dim_ = modConfig.GetLong("embed_dim", 1536)
    If dim_ < 1 Then dim_ = 1536
    Dim prec As String: prec = LCase$(modConfig.GetString("vector_precision", "full"))

    Dim okCount As Long: okCount = 0
    Dim i As Long

    ' ---- mock: 決定的擬似ベクトル(LO/開発ビルド) ----
    If modConfig.GetBool("mock_llm", True) Then
        For i = 0 To n - 1
            Dim mt As String: mt = modUtil.NormalizeForHash(texts(lo + i))
            If LenB(mt) > 0 Then
                Dim mv() As Double
                mv = MockEmbedVector(mt, dim_)
                outCsv(i) = SerializeVector(mv, prec)
                okCount = okCount + 1
            End If
        Next i
        GetEmbeddingsBatch = okCount
        LogEmbedStep n, CLng(modUtilText.ElapsedMsSince(tB0)), okCount
        Exit Function
    End If

    ' ---- direct: Azure埋め込みエンドポイントへ配列POST ----
    If LCase$(modConfig.GetString("embed_transport", "ribbon")) = "direct" Then
        Dim batchSize As Long: batchSize = modConfig.GetLong("embed_batch_size", 128)
        If batchSize < 1 Then batchSize = 1
        If batchSize > 512 Then batchSize = 512

        Dim bStart As Long
        For bStart = 0 To n - 1 Step batchSize
            Dim bEnd As Long: bEnd = bStart + batchSize - 1
            If bEnd > n - 1 Then bEnd = n - 1
            okCount = okCount + modGatewayDirect.DirectEmbedSlice(texts, lo, bStart, bEnd, dim_, prec, outCsv)
        Next bStart
        GetEmbeddingsBatch = okCount
        LogEmbedStep n, CLng(modUtilText.ElapsedMsSince(tB0)), okCount
        Exit Function
    End If

    ' ---- ribbon: 既存の単発GetEmbeddingをループ(現行と同挙動・安全) ----
    okCount = RibbonEmbedRange(texts, lo, 0, n - 1, prec, outCsv)
    GetEmbeddingsBatch = okCount
    LogEmbedStep n, CLng(modUtilText.ElapsedMsSince(tB0)), okCount
End Function

' ----------------------------------------------------------------------------
' RibbonAvailable - AIリボンのアドイン存在確認(結果をセッションキャッシュ)
' ----------------------------------------------------------------------------
' 公式のアドイン検出作法(RIBBON_API_CONFIRMED.md §1末尾・裁定D2):
'   Application.AddIns をループし、アドイン名に config ribbon_addin_name
'   (既定 "リボンちゃん")を含み、かつ Installed=True のものがあれば True。
'   API呼び出し不要・即時で判定できる。
' AddInsコレクションへのアクセス自体が失敗した環境(LO等)では「検出失敗」
' としてFalseを返すが、エラー扱い(E0201)にはしない(利用者を誤ブロック
' しないよう、usage_logへの情報記録に留める)。E0201のLogErrorは実際に
' 呼び出しが必要になった各呼び出し口(CallLLM等)の責務。
' LO互換のためアドインは遅延バインド(Object型)で扱う。
' ----------------------------------------------------------------------------
Public Function RibbonAvailable() As Boolean
    If mRibbonChecked Then
        RibbonAvailable = mRibbonAvailable
        Exit Function
    End If
    mRibbonChecked = True
    mRibbonAvailable = False

    Dim addinName As String
    addinName = modConfig.GetString("ribbon_addin_name", "リボンちゃん")

    On Error GoTo DetectFail
    Dim ai As Object
    For Each ai In Application.AddIns
        If InStr(ai.Name, addinName) > 0 Then
            If ai.Installed Then
                mRibbonAvailable = True
                Exit For
            End If
        End If
    Next ai
    RibbonAvailable = mRibbonAvailable
    Exit Function

DetectFail:
    ' AddInsにアクセスできない環境(LO等)。エラーではなく検出失敗として
    ' 情報記録のみ残す(E0201は各呼び出し口が実呼び出し時に記録する)。
    modLog.LogUsage "ribbon_detect_fail", "gateway", "AddIns走査失敗: " & Err.Description
    mRibbonAvailable = False
    RibbonAvailable = False
End Function

' ----------------------------------------------------------------------------
' RunLimitCheck - リボン公式のLimitCheck()呼び出し口(裁定D3)
'   戻り値: True=続行不可(利用期限切れ等) / False=続行可
' ----------------------------------------------------------------------------
' RIBBON_API_CONFIRMED.md §1 #13: LimitCheck() は Boolean を返し True=続行不可
' (内部で日初の利用同意表示も行う)。mock_llm=TRUE / config limit_check=FALSE
' (エスケープハッチ)/ リボン未検出 / Application.Run が失敗(古いリボン)は
' すべて False(続行可)へ倒す。誤ブロック防止でエラー扱いにはせず
' usage_log への情報記録に留める(裁定D3の穏当運用)。
' ----------------------------------------------------------------------------
Public Function RunLimitCheck() As Boolean
    RunLimitCheck = False

    If modConfig.GetBool("mock_llm", True) Then Exit Function
    If Not modConfig.GetBool("limit_check", True) Then Exit Function
    If Not RibbonAvailable() Then Exit Function

    On Error GoTo CheckFail
    Dim res As Variant
    res = Application.Run("LimitCheck")
    RunLimitCheck = CBool(res)
    Exit Function

CheckFail:
    ' LimitCheck未実装の古いリボン等。誤ブロックしないようFalse(続行可)。
    modLog.LogUsage "limit_check_skip", "gateway", "LimitCheck呼び出し失敗: " & Err.Description
    RunLimitCheck = False
End Function

' ----------------------------------------------------------------------------
' TryRibbonRun - Variant配列argsを展開してApplication.Runを呼ぶ(要素数0～6)。
'   opt層はリボン呼び出しにこれだけを使う(R3)。失敗時は例外を出さず
'   "#ERR:E0202:<説明>" を返す。
' ----------------------------------------------------------------------------
Public Function TryRibbonRun(ByVal funcName As String, ByVal args As Variant) As Variant
    On Error GoTo ErrHandler

    Dim n As Long, lo As Long
    If IsArray(args) Then
        lo = LBound(args)
        n = UBound(args) - lo + 1
        If n < 0 Then n = 0
    ElseIf IsEmpty(args) Then
        n = 0
    Else
        n = 1
    End If

    Select Case n
        Case 0
            TryRibbonRun = Application.Run(funcName)
        Case 1
            If IsArray(args) Then
                TryRibbonRun = Application.Run(funcName, args(lo))
            Else
                TryRibbonRun = Application.Run(funcName, args)
            End If
        Case 2
            TryRibbonRun = Application.Run(funcName, args(lo), args(lo + 1))
        Case 3
            TryRibbonRun = Application.Run(funcName, args(lo), args(lo + 1), args(lo + 2))
        Case 4
            TryRibbonRun = Application.Run(funcName, args(lo), args(lo + 1), args(lo + 2), args(lo + 3))
        Case 5
            TryRibbonRun = Application.Run(funcName, args(lo), args(lo + 1), args(lo + 2), args(lo + 3), args(lo + 4))
        Case 6
            TryRibbonRun = Application.Run(funcName, args(lo), args(lo + 1), args(lo + 2), args(lo + 3), args(lo + 4), args(lo + 5))
        Case Else
            TryRibbonRun = "#ERR:E0202:引数の数が不正です(0～6のみ対応)"
    End Select
    Exit Function

ErrHandler:
    modLog.LogError "E0202", "modGateway.TryRibbonRun", funcName & " : " & Err.Description, Err.Number
    TryRibbonRun = "#ERR:E0202:" & Err.Description
End Function

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
Private Function LooksLikeRealAnswer(ByVal response As String) As Boolean
    If InStr(1, response, "[本棚:", vbTextCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    If InStr(1, response, "[出典", vbTextCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    If InStr(1, response, "<answer>", vbTextCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
    If InStr(1, response, "<thinking>", vbTextCompare) > 0 Then LooksLikeRealAnswer = True: Exit Function
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

' ----------------------------------------------------------------------------
' 内部ヘルパー: バッチ埋め込み(direct/ribbon)・シリアライズ
' ----------------------------------------------------------------------------

' ベクトルを保存精度でCSV化(d6=6桁丸め / full=フル精度)。
' modGatewayDirect からも使う(ベクトルのCSV化は経路で変えない)。
Public Function SerializeVector(ByRef vec() As Double, ByVal prec As String) As String
    If prec = "d6" Then
        SerializeVector = modUtil.VectorToCsvPrec(vec, 6)
    Else
        SerializeVector = modUtil.VectorToCsv(vec)
    End If
End Function

' texts(arrLo+i)(i=iFrom..iTo)をリボン単発でベクトル化しoutCsvへ。戻り値=成功数。
' modGatewayDirect からも使う(direct失敗時のフォールバック先)。
Public Function RibbonEmbedRange(texts() As String, ByVal arrLo As Long, _
                                  ByVal iFrom As Long, ByVal iTo As Long, _
                                  ByVal prec As String, ByRef outCsv() As String) As Long
    Dim dim_ As Long: dim_ = modConfig.GetLong("embed_dim", 1536)
    If dim_ < 1 Then dim_ = 1536

    Dim okCount As Long: okCount = 0
    Dim i As Long
    For i = iFrom To iTo
        Dim t As String: t = modUtil.NormalizeForHash(texts(arrLo + i))
        Dim v() As Double
        If LenB(t) > 0 Then v = GetEmbeddingRibbonOnly(t, dim_)
        If modUtil.HasVector(v) Then
            outCsv(i) = SerializeVector(v, prec)
            okCount = okCount + 1
        End If
    Next i
    RibbonEmbedRange = okCount
End Function


' ----------------------------------------------------------------------------
' 内部ヘルパー: mock応答/mockベクトル
' ----------------------------------------------------------------------------

' step_name別に出典形式[本棚:...]を含む整形済み日本語ダミーを返す。
' UIの全経路(⚡すぐ聞く/🔍しっかり調べる/富化/約款差分)がリボン無しでも
' 本物同様に動くことを保証するための固定応答。
' quick_draft/deep_verify には末尾に [[FOLLOWUP: 候補1 | 候補2]] マーカーを
' 含める(裁定D11。modAskのパース→除去→「深掘り候補」ブロック整形→
' 『続けて質問』のUXが、mock環境でも本物同様に一巡できるようにするため。
' V2 modRibbonGateway のmock検証応答と同じ流儀)。
Private Function MockLLMResponse(ByVal prompt As String, ByVal step_name As String) As String
    ' R26H F5(m-3): mock査読は「1回目=指摘1件 / 2回目=verdict:PASS」の交互。
    ' 常にPASSだと dev で起草→検証→改稿→再検証の4回経路が一度も通らず、常に
    ' 指摘だと早期終了が通らない。交互なら1ターンで両方を必ず通る。カウンタは
    ' 起草(=入念1ターンの開始)で0へ戻すので毎ターン同じ順で再現する。
    ' 実体をmodGenPipe(qa層)へ置く裁定だったが、基盤層→機能層はR1違反(lintが
    ' ERROR)のためカウンタごとここへ置いた(司令塔へ報告)。
    If LCase$(step_name) = "gen_thorough_draft" Then mMockVerifyRound = 0
    If InStr(prompt, "あなたは起草者とは別の査読者です") > 0 Then
        mMockVerifyRound = mMockVerifyRound + 1
        If (mMockVerifyRound Mod 2) = 1 Then
            MockLLMResponse = "・(モック査読)第2段落の断定に根拠が示されていない。" & _
                "適用条件を明示すること。"
        Else
            MockLLMResponse = "verdict:PASS"
        End If
        Exit Function
    End If
    Select Case LCase$(step_name)
        Case "quick_draft"
            MockLLMResponse = "【モック回答/すぐ聞く】" & vbLf & _
                "ご質問について、本棚の資料から関連しそうな箇所を確認しました。" & vbLf & _
                "・ここに実際の回答本文が入ります [本棚:サンプル資料.pdf p.1]" & vbLf & _
                "・資料に無い内容は「資料には見当たらない」と述べます。" & vbLf & vbLf & _
                "(mock_llm=TRUE のためこれはダミー応答です。config の mock_llm を FALSE にすると" & _
                "AIリボンへ実際に問い合わせます。)" & vbLf & _
                "[[FOLLOWUP: (モック)この手続きの必要書類は? | (モック)例外になるケースは?]]"
        Case "deep_draft"
            MockLLMResponse = "【モック回答/しっかり調べる・下書き】" & vbLf & _
                "■結論" & vbLf & _
                "・ここに下書き回答が入ります [本棚:サンプル資料.pdf p.3]" & vbLf & _
                "■根拠" & vbLf & _
                "・関連箇所の要約がここに入ります [パック(サンプル作成者):別資料.docx]" & vbLf & vbLf & _
                "(mock_llm=TRUE によるダミー下書きです。)"
        Case "deep_verify"
            MockLLMResponse = "【モック回答/しっかり調べる・検証済み】" & vbLf & _
                "・下書きの内容を確認し、出典 [本棚:サンプル資料.pdf p.3] と食い違いが無いことを確認しました。" & vbLf & _
                "・最終的な回答本文がここに入ります。" & vbLf & vbLf & _
                "(mock_llm=TRUE によるダミー検証結果です。)" & vbLf & _
                "[[FOLLOWUP: (モック)関連する規程はどれ? | (モック)適用開始日はいつから?]]"
        Case "enrich"
            MockLLMResponse = "[{""i"":1,""summary"":""(モック要約)この章の要点"",""keywords"":""キーワードA,キーワードB""}]"
        Case "decompose"
            ' R16H FB-2(A-L12/B-M6): 段0(論点分け・逆質問の判定)。Case Else の
            ' 汎用ダミーはタグが無く「読めない応答=single」へ退化していた
            ' (結果は正しいが【偶然】正しい)。mock で R16-3系が発火しない
            ' ことを明示的に固定する(docs/20 §3-1 にも1行記載)。
            MockLLMResponse = "<verdict>single</verdict>"
        Case "chapter_summary"
            ' R17 Phase2: 取込時の章単位要約。汎用ダミーはタグが無く
            ' ParseOutlineResp が読めないため【全章が「(要約失敗)」で保存】
            ' =mockでは doc_outline の中身を一度も確認できない(FB-2と同じ理由)。
            MockLLMResponse = "<summary>(モック章要約)この章の要点をここに200～300字で書きます。" & _
                "mock_llm=TRUE のためダミーです。</summary>" & vbLf & _
                "<keywords>キーワードA|キーワードB</keywords>"
        Case "chapter_pick"
            ' R17 Phase2: 俯瞰質問の章選択。空=0章選択で modAskGlobal は
            ' False を返し従来の入念フローへ落ちる(章の本文が無いまま
            ' もっともらしい俯瞰回答が出る状態を mock で作らない)。
            MockLLMResponse = "<pick></pick>"
        Case "global_answer"
            ' R17H FB-2(A-L12): 俯瞰の段3(章をまたいだ回答生成)。Case Else の
            ' 汎用ダミーは章の■見出しも複数の出典タグも持たないため、mockでは
            ' 出典突合(modAskThorough.AnnotateAgainstHits)が効いた形も、俯瞰
            ' らしい回答の並びも一度も確認できない(chapter_summary と同じ理由)。
            MockLLMResponse = "【モック回答/俯瞰】" & vbLf & _
                "■ 第1章 総則" & vbLf & _
                "・章をまたいだ回答がここに入ります [本棚:サンプル資料.pdf p.1]" & vbLf & _
                "■ 第2章 手続" & vbLf & _
                "・章ごとの違い・例外がここに入ります [本棚:サンプル資料.pdf p.3]" & vbLf & vbLf & _
                "(mock_llm=TRUE によるダミー俯瞰回答です。)"
        Case "name_dedup"
            ' R17 Phase3: 用語の名寄せ。空=0グループで modSynonymStore は
            ' 何も書かずに諦める(chapter_summary と同じ理由の明示)。
            MockLLMResponse = "<syn></syn>"
        Case "diff"
            MockLLMResponse = "【モック差分分析】" & vbLf & _
                "・新旧資料を比較しました [本棚:旧版.pdf p.1] → [本棚:新版.pdf p.1]" & vbLf & _
                "・変更点の要約がここに入ります。" & vbLf & vbLf & _
                "(mock_llm=TRUE によるダミー差分結果です。)"
        Case Else
            MockLLMResponse = "【モック応答】" & vbLf & _
                "step_name=""" & step_name & """ に対する整形済みダミー応答です。" & vbLf & _
                "実際の回答はAIリボン接続後にここへ表示されます [本棚:サンプル資料.pdf p.1]。"
    End Select
End Function

' ============================================================================
' MockEmbedVector - 決定的な擬似埋め込みベクトルを生成する(mock_llm=TRUE時)。
' ----------------------------------------------------------------------------
' ■ アルゴリズム(線形合同法 LCG)
'   1. 種(シード)は modUtil.Fnv1a64Hex(正規化後テキスト) の16進16桁の
'      うち先頭13桁(=52bit、Doubleが誤差なく扱える2^53未満)を10進数へ
'      変換した値。同一テキスト(正規化後)からは常に同一のシードになる。
'   2. 以後 dim 回、以下の漸化式でシードを更新しながら成分を1つずつ生成する
'      (定数はNumerical Recipes系でよく使われる値):
'        A = 1664525, C = 1013904223, M = 2^32
'        seed = (A * seed + C) Mod M
'        component = (seed / M) * 2 - 1     ' [-1, 1) の範囲へ写像
'      A*seed の最大値は約 1664525*(2^32-1) ≒ 7.15e15 で、Doubleが誤差なく
'      表現できる上限 2^53 ≒ 9.007e15 を下回るため、桁落ちは発生しない。
'   3. 生成した dim 次元のベクトルを modUtil.L2Normalize でL2正規化して返す。
' ============================================================================
Private Function MockEmbedVector(ByVal normalizedText As String, ByVal dimCount As Long) As Double()
    Dim d As Long: d = dimCount
    If d < 1 Then d = 1536

    Dim h As String: h = modUtil.Fnv1a64Hex(normalizedText)
    Dim seed As Double: seed = HexPrefixToDouble(h, 13)

    Const A As Double = 1664525#
    Const C As Double = 1013904223#
    Const M As Double = 4294967296#   ' 2^32

    Dim vec() As Double: ReDim vec(0 To d - 1)
    Dim i As Long
    For i = 0 To d - 1
        seed = A * seed + C
        seed = seed - Int(seed / M) * M     ' Mod M(Doubleで厳密な非負剰余)
        vec(i) = (seed / M) * 2# - 1#
    Next i

    modUtil.L2Normalize vec
    MockEmbedVector = vec
End Function

' 16進文字列の先頭nDigits桁を、符号なし整数値としてDoubleへ変換する。
' nDigitsは13以下を想定(52bit以内、Doubleで誤差なく表現できる範囲)。
Private Function HexPrefixToDouble(ByVal hexStr As String, ByVal nDigits As Long) As Double
    Dim v As Double: v = 0
    Dim n As Long: n = nDigits
    If n > Len(hexStr) Then n = Len(hexStr)
    Dim i As Long
    For i = 1 To n
        v = v * 16# + HexDigitValue(Mid$(hexStr, i, 1))
    Next i
    HexPrefixToDouble = v
End Function

Private Function HexDigitValue(ByVal c As String) As Long
    Dim u As String: u = UCase$(c)
    Select Case u
        Case "0" To "9": HexDigitValue = Asc(u) - Asc("0")
        Case "A" To "F": HexDigitValue = Asc(u) - Asc("A") + 10
        Case Else: HexDigitValue = 0
    End Select
End Function
