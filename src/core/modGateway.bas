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
' ツールを識別できるようにするため)。
' 第7・8引数prevU/prevAは会話継続用の履歴(確定根拠: 台帳§1 #1のChatGPT
' シグネチャ第7・8引数+裁定D11)。複数往復ぶんは「新しい順」に ";;;" 区切り
' で連結した文字列を渡す(prevU=過去の質問、prevA=過去の回答。履歴の保持と
' 連結はmodAsk側の責務)。本関数のprevU/prevAはOptional末尾追加のため、
' 既定""=履歴なしで従来と完全に同じ呼び出しになり、既存呼び出し元
' (modAsk/modEnrich/optDiffDoc等)は無改修で動く(後方互換)。
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
        latency_ms = CLng((Timer - t0) * 1000)
        Exit Function
    End If

    If Not RibbonAvailable() Then
        modLog.LogError "E0201", "modGateway.CallLLM", "step=" & step_name
        CallLLM = "#ERR:E0201:AIリボンが見つかりません"
        latency_ms = CLng((Timer - t0) * 1000)
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
    latency_ms = CLng((Timer - t0) * 1000)

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
    latency_ms = CLng((Timer - t0) * 1000)
    modLog.LogError "E0202", "modGateway.CallLLM", "step=" & step_name & " err=" & Err.Description
    CallLLM = "#ERR:E0202:" & Err.Description
End Function

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
        latency_ms = CLng((Timer - t0) * 1000)
        Exit Function   ' 空配列を返す
    End If

    If modConfig.GetBool("mock_llm", True) Then
        GetEmbedding = MockEmbedVector(t, dim_)
        latency_ms = CLng((Timer - t0) * 1000)
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
                    latency_ms = CLng((Timer - t0) * 1000)
                    Exit Function
                End If
            End If
        End If
        latency_ms = CLng((Timer - t0) * 1000)
        Exit Function   ' 失敗は空配列(バッチ側でE0203記録済み)
    End If

    GetEmbedding = GetEmbeddingRibbonOnly(t, dim_)
    latency_ms = CLng((Timer - t0) * 1000)
    Exit Function

ErrHandler:
    latency_ms = CLng((Timer - t0) * 1000)
    modLog.LogError "E0203", "modGateway.GetEmbedding", "err=" & Err.Description
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
            okCount = okCount + DirectEmbedSlice(texts, lo, bStart, bEnd, dim_, prec, outCsv)
        Next bStart
        GetEmbeddingsBatch = okCount
        Exit Function
    End If

    ' ---- ribbon: 既存の単発GetEmbeddingをループ(現行と同挙動・安全) ----
    okCount = RibbonEmbedRange(texts, lo, 0, n - 1, prec, outCsv)
    GetEmbeddingsBatch = okCount
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
' RIBBON_API_CONFIRMED.md §1 #13: LimitCheck() は Boolean を返し、
' True=続行不可(内部で日初の利用同意表示も行う)。
'   ・mock_llm=TRUE、または config limit_check=FALSE(エスケープハッチ)の
'     ときは呼ばずに即False(続行可)。
'   ・リボン未検出時もFalse(呼びようがない。E0201は実呼び出し時に記録)。
'   ・Application.Run("LimitCheck")がエラーになった場合(LimitCheckを持たない
'     古いリボン等)もFalseに倒す。誤ブロック防止のためエラー扱いにはせず
'     usage_logへの情報記録に留める(裁定D3の穏当運用)。
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
' TryRibbonRun - Variant配列argsを展開してApplication.Runを呼ぶ(要素数0〜6)。
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
            TryRibbonRun = "#ERR:E0202:引数の数が不正です(0〜6のみ対応)"
    End Select
    Exit Function

ErrHandler:
    modLog.LogError "E0202", "modGateway.TryRibbonRun", funcName & " : " & Err.Description
    TryRibbonRun = "#ERR:E0202:" & Err.Description
End Function

' LLMの応答文字列に「上限/limit/回数/rate/quota」の語が含まれるかを検知する。
' 確実な判定ではないが、E0204(利用上限)の可能性を利用者に知らせる簡易判定。
Public Function LooksLikeLimitError(ByVal response As String) As Boolean
    Dim s As String: s = LCase$(response)
    LooksLikeLimitError = (InStr(s, "上限") > 0) Or (InStr(s, "limit") > 0) Or _
                           (InStr(s, "回数") > 0) Or (InStr(s, "rate") > 0) Or _
                           (InStr(s, "quota") > 0)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー: バッチ埋め込み(direct/ribbon)・シリアライズ
' ----------------------------------------------------------------------------

' ベクトルを保存精度でCSV化(d6=6桁丸め / full=フル精度)。
Private Function SerializeVector(ByRef vec() As Double, ByVal prec As String) As String
    If prec = "d6" Then
        SerializeVector = modUtil.VectorToCsvPrec(vec, 6)
    Else
        SerializeVector = modUtil.VectorToCsv(vec)
    End If
End Function

' texts(arrLo+i)(i=iFrom..iTo)をリボン単発でベクトル化しoutCsvへ。戻り値=成功数。
Private Function RibbonEmbedRange(texts() As String, ByVal arrLo As Long, _
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

' 1バッチ分をAzureへ配列POSTする。失敗時はリボンへフォールバック。戻り値=成功数。
Private Function DirectEmbedSlice(texts() As String, ByVal arrLo As Long, _
                                  ByVal iFrom As Long, ByVal iTo As Long, _
                                  ByVal dims As Long, ByVal prec As String, _
                                  ByRef outCsv() As String) As Long
    Dim apiUrl As String: apiUrl = Trim$(modConfig.GetString("azure_embed_url", ""))
    Dim apiKey As String: apiKey = Trim$(modConfig.GetString("azure_embed_key", ""))
    If LenB(apiUrl) = 0 Or LenB(apiKey) = 0 Then
        modLog.LogError "E0203", "modGateway.GetEmbeddingsBatch", "azure_embed_url/keyが未設定(ribbonへフォールバック)"
        DirectEmbedSlice = RibbonEmbedRange(texts, arrLo, iFrom, iTo, prec, outCsv)
        Exit Function
    End If

    On Error GoTo HttpFail

    ' リクエストボディ {"input":["...",...]}(正規化+4000字打ち切りはribbon経路と同一)
    Dim bodyParts() As String
    ReDim bodyParts(0 To iTo - iFrom)
    Dim i As Long
    For i = iFrom To iTo
        Dim t As String: t = modUtil.NormalizeForHash(texts(arrLo + i))
        If LenB(t) = 0 Then t = " "   ' Azureは空文字を拒否するためダミー1字
        bodyParts(i - iFrom) = """" & EscapeJsonStr(modUtil.SafeLeft(t, 4000)) & """"
    Next i
    Dim body As String
    body = "{""input"":[" & Join(bodyParts, ",") & "]}"

    ' タイムアウト(ms)。NW瞬断でもExcelが無限フリーズしないよう明示設定する。
    Dim toMs As Long: toMs = modConfig.GetLong("azure_http_timeout_ms", 60000)
    If toMs < 1000 Then toMs = 1000

    Dim http As Object
    ' MSXML2.XMLHTTPはタイムアウトAPIを持たず瞬断で無限待ちになるため、
    ' setTimeoutsを持つ ServerXMLHTTP.6.0 を使う(resolve/connect/send/receive)。
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 10000, toMs, toMs
    http.Open "POST", apiUrl, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "api-key", apiKey
    http.Send body

    If CLng(http.Status) <> 200 Then
        modLog.LogError "E0203", "modGateway.GetEmbeddingsBatch", _
            "HTTP " & http.Status & ": " & modUtil.SafeLeft(CStr(http.responseText), 200) & "(ribbonへフォールバック)"
        DirectEmbedSlice = RibbonEmbedRange(texts, arrLo, iFrom, iTo, prec, outCsv)
        GoTo Cleanup
    End If

    ' レスポンスから "embedding":[...] を出現順に抽出(dataは入力順)
    Dim resp As String: resp = CStr(http.responseText)
    Dim okCount As Long: okCount = 0
    Dim searchPos As Long: searchPos = 1
    For i = iFrom To iTo
        Dim vecCsv As String
        vecCsv = NextEmbeddingArray(resp, searchPos)
        If LenB(vecCsv) > 0 Then
            Dim v() As Double
            If modUtil.CsvToVector(vecCsv, v) Then
                If modUtil.TruncateAndRenorm(v, dims) Then
                    outCsv(i) = SerializeVector(v, prec)
                    okCount = okCount + 1
                End If
            End If
        End If
    Next i

    If okCount = 0 Then
        modLog.LogError "E0203", "modGateway.GetEmbeddingsBatch", _
            "応答のembedding抽出0件(ribbonへフォールバック): " & modUtil.SafeLeft(resp, 200)
        DirectEmbedSlice = RibbonEmbedRange(texts, arrLo, iFrom, iTo, prec, outCsv)
        GoTo Cleanup
    End If

    DirectEmbedSlice = okCount

Cleanup:
    ' COM解放(正常・異常問わず必ず通る)。メモリリーク防止。
    On Error Resume Next
    Set http = Nothing
    On Error GoTo 0
    Exit Function

HttpFail:
    modLog.LogError "E0203", "modGateway.GetEmbeddingsBatch", _
        "通信エラー: " & Err.Description & "(ribbonへフォールバック)"
    Err.Clear
    On Error GoTo 0
    DirectEmbedSlice = RibbonEmbedRange(texts, arrLo, iFrom, iTo, prec, outCsv)
    On Error Resume Next
    Set http = Nothing
    On Error GoTo 0
End Function

' レスポンス文字列のsearchPos以降から次の "embedding":[数値,...] を探し、
' 中身(カンマ区切り数値)を返す。searchPosは次の検索開始位置へ進める。
' 見つからなければ""(呼び出し側が失敗扱い)。
Private Function NextEmbeddingArray(ByVal resp As String, ByRef searchPos As Long) As String
    Dim keyPos As Long
    keyPos = InStr(searchPos, resp, """embedding""", vbTextCompare)
    If keyPos = 0 Then Exit Function

    Dim openPos As Long
    openPos = InStr(keyPos, resp, "[")
    If openPos = 0 Then Exit Function

    Dim closePos As Long
    closePos = InStr(openPos, resp, "]")
    If closePos = 0 Then Exit Function

    searchPos = closePos + 1
    NextEmbeddingArray = Mid$(resp, openPos + 1, closePos - openPos - 1)
End Function

' JSON文字列エスケープ(direct用)。制御文字は\uXXXXへ。
Private Function EscapeJsonStr(ByVal s As String) As String
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Function

    Dim parts() As String: ReDim parts(1 To n)
    Dim i As Long
    For i = 1 To n
        Dim ch As String: ch = Mid$(s, i, 1)
        Select Case ch
            Case "\": parts(i) = "\\"
            Case """": parts(i) = "\"""
            Case vbLf: parts(i) = "\n"
            Case vbCr: parts(i) = "\r"
            Case vbTab: parts(i) = "\t"
            Case Else
                Dim code As Long: code = AscW(ch)
                If code < 0 Then code = code + 65536
                If code < 32 Then
                    parts(i) = "\u" & Right$("000" & Hex$(code), 4)
                Else
                    parts(i) = ch
                End If
        End Select
    Next i
    EscapeJsonStr = Join(parts, "")
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
