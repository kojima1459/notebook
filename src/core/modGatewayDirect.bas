Attribute VB_Name = "modGatewayDirect"
Option Explicit

' ========================================
' modGatewayDirect - 埋め込みの direct 経路(Azure へバッチ直接送信)
'
' modGateway から切り出した。埋め込みの窓口は modGateway 1本のままで、
' 「リボン経由」と「Azure直叩き」のうち後者の実装だけをここに置く。
' 経路の選択(embed_transport)とフォールバックの判断は modGateway が持つ。
'
' 既定は ribbon。本番ビルドは azure_embed_key を焼き込めない
' (ブック配布=キー配布になるため)ので、direct はキーを自分で入れる
' 管理端末で明示的に切り替えたときだけ使う経路になる。
'
' 切り出しの理由(2026-07-28): modGateway が契約上限30,000字に対し
' 残り292字となり、次の修正が入らなくなっていた。
' ========================================

' direct→ribbon フォールバックの案内を1セッション1回に抑えるフラグ。
Private mDirectFallbackLogged As Boolean

' 1バッチ分をAzureへ配列POSTする。失敗時はリボンへフォールバック。戻り値=成功数。
Public Function DirectEmbedSlice(texts() As String, ByVal arrLo As Long, _
                                  ByVal iFrom As Long, ByVal iTo As Long, _
                                  ByVal dims As Long, ByVal prec As String, _
                                  ByRef outCsv() As String) As Long
    Dim apiUrl As String: apiUrl = Trim$(modConfig.GetString("azure_embed_url", ""))
    ' azure_embed_keyはビルド時に軽く難読化(OBF1:接頭辞)されて格納されている
    ' 場合がある(build_mybookshelf.py obfuscate_secret)。未難読化の平文が
    ' 手動で入っている場合はDeobfuscateSecretがそのまま素通しする。
    Dim apiKey As String: apiKey = Trim$(modUtil.DeobfuscateSecret(modConfig.GetString("azure_embed_key", "")))
    If LenB(apiUrl) = 0 Or LenB(apiKey) = 0 Then
        ' 2026-07-28: ここは「障害」ではなく「そう設定されている」状態。
        ' 未設定のまま direct で走ると、スライスのたびに E0203 が積まれて
        ' err_log が埋まり、本当の障害が見えなくなる。1セッション1回だけ残す。
        If Not mDirectFallbackLogged Then
            mDirectFallbackLogged = True
            modLog.LogUsage "embed_fallback_ribbon", "", _
                "azure_embed_url/keyが未設定のため、このセッションはribbon経由で埋め込みます"
        End If
        DirectEmbedSlice = modGateway.RibbonEmbedRange(texts, arrLo, iFrom, iTo, prec, outCsv)
        Exit Function
    End If

    On Error GoTo HttpFail

    ' リクエストボディ {"input":["...",...]}(正規化+4000字打ち切りはribbon経路と同一)
    '
    ' 2026-07-28(レビュー L-4): 空チャンクをダミー1字 " " で埋めて
    ' 「ベクトル化成功」として保存していた。中身の無いチャンクに空白1文字の
    ' ベクトルが付き、検索の邪魔にしかならない。しかも ribbon/mock 経路は
    ' 空を失敗扱いにするため、経路によって結果が変わる非対称もあった。
    ' 空はリクエストから外し、結果も空のままにする。
    Dim sendIdx() As Long: ReDim sendIdx(0 To iTo - iFrom)
    Dim bodyParts() As String: ReDim bodyParts(0 To iTo - iFrom)
    Dim sendN As Long: sendN = 0
    Dim i As Long
    For i = iFrom To iTo
        Dim t As String: t = modUtil.NormalizeForHash(texts(arrLo + i))
        If LenB(t) > 0 Then
            sendIdx(sendN) = i
            bodyParts(sendN) = """" & EscapeJsonStr(modUtil.SafeLeft(t, 4000)) & """"
            sendN = sendN + 1
        End If
    Next i
    If sendN = 0 Then Exit Function      ' 送るものが無い(全部空)
    ReDim Preserve bodyParts(0 To sendN - 1)
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
        ' http_statusを構造化フィールドへ渡す(社内プロキシに阻まれた場合の407等、
        ' err_logを見た瞬間に原因が分かるようにするため)。
        modLog.LogError "E0203", "modGateway.GetEmbeddingsBatch", _
            "HTTP " & http.Status & ": " & modUtil.SafeLeft(CStr(http.responseText), 200) & "(ribbonへフォールバック)", _
            0, CLng(http.Status)
        DirectEmbedSlice = modGateway.RibbonEmbedRange(texts, arrLo, iFrom, iTo, prec, outCsv)
        GoTo Cleanup
    End If

    ' レスポンスから "embedding":[...] を出現順に抽出(dataは入力順)
    Dim resp As String: resp = CStr(http.responseText)
    Dim okCount As Long: okCount = 0
    Dim searchPos As Long: searchPos = 1
    Dim k As Long
    For k = 0 To sendN - 1
        Dim vecCsv As String
        vecCsv = NextEmbeddingArray(resp, searchPos)
        If LenB(vecCsv) > 0 Then
            Dim v() As Double
            If modUtil.CsvToVector(vecCsv, v) Then
                If modUtil.TruncateAndRenorm(v, dims) Then
                    ' 応答は入力順。空を外した分、書き戻し先は sendIdx で引く。
                    outCsv(sendIdx(k)) = modGateway.SerializeVector(v, prec)
                    okCount = okCount + 1
                End If
            End If
        End If
    Next k

    If okCount = 0 Then
        modLog.LogError "E0203", "modGateway.GetEmbeddingsBatch", _
            "応答のembedding抽出0件(ribbonへフォールバック): " & modUtil.SafeLeft(resp, 200)
        DirectEmbedSlice = modGateway.RibbonEmbedRange(texts, arrLo, iFrom, iTo, prec, outCsv)
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
    ' Err.Clearの前に必ず番号を退避する(社内プロキシ接続拒否等のCOM/WinHTTPエラー
    ' 番号を握りつぶさないため)。
    Dim httpFailNum As Long: httpFailNum = Err.Number
    modLog.LogError "E0203", "modGateway.GetEmbeddingsBatch", _
        "通信エラー: " & Err.Description & "(ribbonへフォールバック)", httpFailNum
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume HttpFailCleanup0
HttpFailCleanup0:
    Err.Clear
    On Error GoTo 0
    DirectEmbedSlice = modGateway.RibbonEmbedRange(texts, arrLo, iFrom, iTo, prec, outCsv)
    On Error Resume Next
    Set http = Nothing
    On Error GoTo 0
End Function

' レスポンス文字列のsearchPos以降から次の "embedding":[数値,...] を探し、
' 中身(カンマ区切り数値)を返す。searchPosは次の検索開始位置へ進める。
' 見つからなければ""(呼び出し側が失敗扱い)。
Public Function NextEmbeddingArray(ByVal resp As String, ByRef searchPos As Long) As String
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
Public Function EscapeJsonStr(ByVal s As String) As String
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
