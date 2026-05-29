Attribute VB_Name = "modApiGateway"
Option Explicit

' ============================================================================
' modApiGateway - Provider-agnostic Embed() / Chat() facade
' ----------------------------------------------------------------------------
' Hides the API mode (A/B/C) and provider (Azure OpenAI now, OpenAI/relay
' later) behind two functions. Callers only know about texts and messages.
' ============================================================================

Public Type ChatResult
    OK As Boolean
    Content As String
    PromptTokens As Long
    CompletionTokens As Long
    LatencyMs As Long
    ErrorMessage As String
End Type

Public Type EmbedResult
    OK As Boolean
    Vectors() As Double ' flat: [vec0_d0, vec0_d1, ..., vec1_d0, ...]
    Dim_ As Long
    Count_ As Long
    PromptTokens As Long
    LatencyMs As Long
    ErrorMessage As String
End Type

' ----------------------------------------------------------------------------
' Public: Embed a batch of texts
' ----------------------------------------------------------------------------
Public Function Embed(ByRef texts() As String) As EmbedResult
    Dim r As EmbedResult
    Dim mode As String: mode = UCase$(modConfig.GetString("api", "mode", "B"))

    Dim url As String
    Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
    BuildEmbedRequest mode, url, headers
    If LenB(url) = 0 Then
        r.OK = False
        r.ErrorMessage = "Embed URL not configured"
        Embed = r
        Exit Function
    End If

    Dim body As String
    body = BuildEmbedBody(texts)

    Dim resp As HttpResponse
    resp = modHttpClient.PostJson(url, body, headers, 90000)
    r.LatencyMs = resp.LatencyMs
    If resp.Status < 200 Or resp.Status >= 300 Then
        r.OK = False
        r.ErrorMessage = "HTTP " & resp.Status & " " & Left$(resp.Body, 500)
        If LenB(resp.ErrorMessage) > 0 Then
            r.ErrorMessage = r.ErrorMessage & " | " & resp.ErrorMessage
        End If
        Embed = r
        Exit Function
    End If

    On Error GoTo ParseFail
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(resp.Body)
    Dim items As Object: Set items = parsed("data")
    Dim n As Long: n = items.Count
    Dim d As Long
    Dim firstVec As Object: Set firstVec = items(1)("embedding")
    d = firstVec.Count

    ReDim r.Vectors(0 To n * d - 1)
    Dim i As Long, j As Long, base As Long
    For i = 1 To n
        Dim vec As Object: Set vec = items(i)("embedding")
        base = (i - 1) * d
        For j = 1 To d
            r.Vectors(base + j - 1) = CDbl(vec(j))
        Next j
    Next i
    r.Dim_ = d
    r.Count_ = n

    If parsed.Exists("usage") Then
        r.PromptTokens = CLng(parsed("usage")("prompt_tokens"))
    End If
    r.OK = True
    Embed = r
    Exit Function

ParseFail:
    r.OK = False
    r.ErrorMessage = "JSON parse failed: " & Err.Description
    Embed = r
End Function

' ----------------------------------------------------------------------------
' Public: Chat completion (single turn). Messages built from system + user.
' ----------------------------------------------------------------------------
Public Function Chat(ByVal systemPrompt As String, ByVal userPrompt As String) As ChatResult
    Dim r As ChatResult
    Dim mode As String: mode = UCase$(modConfig.GetString("api", "mode", "B"))

    Dim url As String
    Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
    BuildChatRequest mode, url, headers
    If LenB(url) = 0 Then
        r.OK = False
        r.ErrorMessage = "Chat URL not configured"
        Chat = r
        Exit Function
    End If

    Dim body As String
    body = BuildChatBody(systemPrompt, userPrompt)

    Dim resp As HttpResponse
    resp = modHttpClient.PostJson(url, body, headers, 120000)
    r.LatencyMs = resp.LatencyMs
    If resp.Status < 200 Or resp.Status >= 300 Then
        r.OK = False
        r.ErrorMessage = "HTTP " & resp.Status & " " & Left$(resp.Body, 500)
        If LenB(resp.ErrorMessage) > 0 Then
            r.ErrorMessage = r.ErrorMessage & " | " & resp.ErrorMessage
        End If
        Chat = r
        Exit Function
    End If

    On Error GoTo ParseFail
    Dim parsed As Object: Set parsed = JsonConverter.ParseJson(resp.Body)
    Dim choices As Object: Set choices = parsed("choices")
    If choices.Count = 0 Then
        r.OK = False
        r.ErrorMessage = "No choices in response"
        Chat = r
        Exit Function
    End If
    r.Content = CStr(choices(1)("message")("content"))
    If parsed.Exists("usage") Then
        r.PromptTokens = CLng(parsed("usage")("prompt_tokens"))
        r.CompletionTokens = CLng(parsed("usage")("completion_tokens"))
    End If
    r.OK = True
    Chat = r
    Exit Function

ParseFail:
    r.OK = False
    r.ErrorMessage = "JSON parse failed: " & Err.Description
    Chat = r
End Function

' ----------------------------------------------------------------------------
' Request builders (per mode)
' ----------------------------------------------------------------------------

Private Sub BuildEmbedRequest(ByVal mode As String, ByRef url As String, ByVal headers As Object)
    Select Case mode
        Case "A"
            url = modConfig.GetString("api", "relay_url", "") & "/embeddings"
            ApplyRelayAuth headers
        Case "B", "C"
            url = BuildAzureUrl(modConfig.GetString("api", "embed_deployment", ""), "embeddings")
            ApplyAzureAuth mode, headers
    End Select
End Sub

Private Sub BuildChatRequest(ByVal mode As String, ByRef url As String, ByVal headers As Object)
    Select Case mode
        Case "A"
            url = modConfig.GetString("api", "relay_url", "") & "/chat/completions"
            ApplyRelayAuth headers
        Case "B", "C"
            url = BuildAzureUrl(modConfig.GetString("api", "chat_deployment", ""), "chat/completions")
            ApplyAzureAuth mode, headers
    End Select
End Sub

Private Function BuildAzureUrl(ByVal deployment As String, ByVal route As String) As String
    Dim endpoint As String, ver As String
    endpoint = modConfig.GetString("api", "endpoint", "")
    ver = modConfig.GetString("api", "api_version", "2024-10-21")
    If LenB(endpoint) = 0 Or LenB(deployment) = 0 Then
        BuildAzureUrl = ""
        Exit Function
    End If
    If Right$(endpoint, 1) = "/" Then endpoint = Left$(endpoint, Len(endpoint) - 1)
    BuildAzureUrl = endpoint & "/openai/deployments/" & deployment & "/" & route & "?api-version=" & ver
End Function

Private Sub ApplyAzureAuth(ByVal mode As String, ByVal headers As Object)
    Select Case mode
        Case "B"
            headers("api-key") = modKeyVault.GetKey()
        Case "C"
            headers("Authorization") = "Bearer " & modKeyVault.GetKey()
    End Select
End Sub

Private Sub ApplyRelayAuth(ByVal headers As Object)
    ' Per-user token if the relay requires one. Empty for IP-allowlist relays.
    Dim tok As String: tok = modKeyVault.GetKey()
    If LenB(tok) > 0 Then headers("Authorization") = "Bearer " & tok
End Sub

' ----------------------------------------------------------------------------
' Body builders (Azure OpenAI / OpenAI use the same chat & embeddings schema)
' ----------------------------------------------------------------------------

Private Function BuildEmbedBody(ByRef texts() As String) As String
    Dim sb As String, i As Long
    sb = "{""input"":["
    For i = LBound(texts) To UBound(texts)
        If i > LBound(texts) Then sb = sb & ","
        sb = sb & JsonString(texts(i))
    Next i
    sb = sb & "]}"
    BuildEmbedBody = sb
End Function

Private Function BuildChatBody(ByVal systemPrompt As String, ByVal userPrompt As String) As String
    Dim sb As String
    sb = "{""messages"":["
    sb = sb & "{""role"":""system"",""content"":" & JsonString(systemPrompt) & "},"
    sb = sb & "{""role"":""user"",""content"":" & JsonString(userPrompt) & "}"
    sb = sb & "],""temperature"":0.2}"
    BuildChatBody = sb
End Function

Private Function JsonString(ByVal s As String) As String
    Dim out As String, i As Long, ch As String, code As Long
    out = """"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 34:  out = out & "\"""
            Case 92:  out = out & "\\"
            Case 8:   out = out & "\b"
            Case 9:   out = out & "\t"
            Case 10:  out = out & "\n"
            Case 12:  out = out & "\f"
            Case 13:  out = out & "\r"
            Case Else
                If code < 32 Then
                    out = out & "\u" & Right$("0000" & Hex(code), 4)
                Else
                    out = out & ch
                End If
        End Select
    Next i
    JsonString = out & """"
End Function
