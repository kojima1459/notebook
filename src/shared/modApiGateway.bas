Attribute VB_Name = "modApiGateway"
Option Explicit

' ============================================================================
' modApiGateway - Provider-agnostic Embed() / Chat() facade
' ----------------------------------------------------------------------------
' Hides the provider (Azure OpenAI / Gemini / relay) and key mode (A/B/C/P)
' behind two functions. Callers only know about texts and messages.
'
' Selected via config.ini [api] provider = azure_openai | gemini
' Default is azure_openai for production builds; the demo build flips it to
' gemini in config.sample.ini.
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

Public Function Embed(ByRef texts() As String) As EmbedResult
    Select Case LCase$(modConfig.GetString("api", "provider", "azure_openai"))
        Case "gemini": Embed = EmbedGemini(texts)
        Case Else:     Embed = EmbedAzureOpenAI(texts)
    End Select
End Function

Public Function Chat(ByVal systemPrompt As String, ByVal userPrompt As String) As ChatResult
    Select Case LCase$(modConfig.GetString("api", "provider", "azure_openai"))
        Case "gemini": Chat = ChatGemini(systemPrompt, userPrompt)
        Case Else:     Chat = ChatAzureOpenAI(systemPrompt, userPrompt)
    End Select
End Function

' ============================================================================
' Gemini (Google Generative Language API)
' ----------------------------------------------------------------------------
' Auth: header `x-goog-api-key: <key>` (URL query also works but stays out of
' proxy logs this way).
' Chat: POST .../v1beta/models/{model}:generateContent
' Embed: POST .../v1beta/models/{model}:batchEmbedContents
' ============================================================================

Private Function GeminiBaseUrl() As String
    GeminiBaseUrl = "https://generativelanguage.googleapis.com/v1beta/models/"
End Function

Private Function EmbedGemini(ByRef texts() As String) As EmbedResult
    Dim r As EmbedResult
    Dim model As String
    model = modConfig.GetString("api", "embed_deployment", "gemini-embedding-001")
    Dim outDim As Long: outDim = modConfig.GetLong("api", "embed_dim", 768)
    Dim url As String: url = GeminiBaseUrl() & model & ":batchEmbedContents"

    Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
    headers("x-goog-api-key") = modKeyVault.GetKey()

    Dim body As String, i As Long
    body = "{""requests"":["
    For i = LBound(texts) To UBound(texts)
        If i > LBound(texts) Then body = body & ","
        body = body & "{""model"":""models/" & model & """,""content"":{""parts"":[{""text"":" & _
               JsonString(texts(i)) & "}]},""taskType"":""RETRIEVAL_QUERY"",""outputDimensionality"":" & outDim & "}"
    Next i
    body = body & "]}"

    Dim resp As HttpResponse
    resp = modHttpClient.PostJson(url, body, headers, 90000)
    r.LatencyMs = resp.LatencyMs
    If resp.Status < 200 Or resp.Status >= 300 Then
        r.OK = False
        r.ErrorMessage = "HTTP " & resp.Status & " " & Left$(resp.Body, 500)
        If LenB(resp.ErrorMessage) > 0 Then r.ErrorMessage = r.ErrorMessage & " | " & resp.ErrorMessage
        EmbedGemini = r
        Exit Function
    End If

    On Error GoTo ParseFail
    Dim parsed As Object: Set parsed = JsonConverter.ParseJson(resp.Body)
    Dim arr As Object: Set arr = parsed("embeddings")
    Dim n As Long: n = arr.Count
    If n = 0 Then
        r.OK = False
        r.ErrorMessage = "Empty embeddings array"
        EmbedGemini = r
        Exit Function
    End If
    Dim firstVec As Object: Set firstVec = arr(1)("values")
    Dim d As Long: d = firstVec.Count

    ReDim r.Vectors(0 To n * d - 1)
    Dim j As Long, base As Long
    For i = 1 To n
        Dim vec As Object: Set vec = arr(i)("values")
        base = (i - 1) * d
        For j = 1 To d
            r.Vectors(base + j - 1) = CDbl(vec(j))
        Next j
    Next i
    r.Dim_ = d
    r.Count_ = n
    r.OK = True
    EmbedGemini = r
    Exit Function
ParseFail:
    r.OK = False
    r.ErrorMessage = "Gemini embed parse failed: " & Err.Description
    EmbedGemini = r
End Function

Private Function ChatGemini(ByVal systemPrompt As String, ByVal userPrompt As String) As ChatResult
    Dim r As ChatResult
    Dim model As String
    model = modConfig.GetString("api", "chat_deployment", "gemini-2.5-flash")
    Dim url As String: url = GeminiBaseUrl() & model & ":generateContent"

    Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
    headers("x-goog-api-key") = modKeyVault.GetKey()

    Dim maxOut As Long: maxOut = modConfig.GetLong("api", "max_output_tokens", 16384)
    Dim thinkBudget As Long: thinkBudget = modConfig.GetLong("api", "thinking_budget", 4096)
    Dim body As String
    body = "{""system_instruction"":{""parts"":[{""text"":" & JsonString(systemPrompt) & "}]},"
    body = body & """contents"":[{""role"":""user"",""parts"":[{""text"":" & JsonString(userPrompt) & "}]}],"
    body = body & """generationConfig"":{""temperature"":0.15,""maxOutputTokens"":" & maxOut & _
           ",""topP"":0.95,""thinkingConfig"":{""thinkingBudget"":" & thinkBudget & "}}}"

    Dim resp As HttpResponse
    resp = modHttpClient.PostJson(url, body, headers, 240000)
    r.LatencyMs = resp.LatencyMs
    If resp.Status < 200 Or resp.Status >= 300 Then
        r.OK = False
        r.ErrorMessage = "HTTP " & resp.Status & " " & Left$(resp.Body, 500)
        If LenB(resp.ErrorMessage) > 0 Then r.ErrorMessage = r.ErrorMessage & " | " & resp.ErrorMessage
        ChatGemini = r
        Exit Function
    End If

    On Error GoTo ParseFail
    Dim parsed As Object: Set parsed = JsonConverter.ParseJson(resp.Body)
    Dim cand As Object: Set cand = parsed("candidates")
    If cand.Count = 0 Then
        r.OK = False
        r.ErrorMessage = "Gemini returned no candidates"
        ChatGemini = r
        Exit Function
    End If
    Dim content As Object: Set content = cand(1)("content")
    Dim parts As Object: Set parts = content("parts")
    Dim sb As String, i As Long
    For i = 1 To parts.Count
        If parts(i).Exists("text") Then sb = sb & CStr(parts(i)("text"))
    Next i
    r.Content = sb

    If parsed.Exists("usageMetadata") Then
        Dim u As Object: Set u = parsed("usageMetadata")
        If u.Exists("promptTokenCount") Then r.PromptTokens = CLng(u("promptTokenCount"))
        If u.Exists("candidatesTokenCount") Then r.CompletionTokens = CLng(u("candidatesTokenCount"))
    End If
    r.OK = True
    ChatGemini = r
    Exit Function
ParseFail:
    r.OK = False
    r.ErrorMessage = "Gemini chat parse failed: " & Err.Description
    ChatGemini = r
End Function

' ============================================================================
' Azure OpenAI (unchanged from production design)
' ============================================================================

Private Function EmbedAzureOpenAI(ByRef texts() As String) As EmbedResult
    Dim r As EmbedResult
    Dim mode As String: mode = UCase$(modConfig.GetString("api", "mode", "B"))

    Dim url As String
    Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
    BuildAzureEmbedRequest mode, url, headers
    If LenB(url) = 0 Then
        r.OK = False
        r.ErrorMessage = "Embed URL not configured"
        EmbedAzureOpenAI = r
        Exit Function
    End If

    Dim body As String: body = BuildAzureEmbedBody(texts)

    Dim resp As HttpResponse
    resp = modHttpClient.PostJson(url, body, headers, 90000)
    r.LatencyMs = resp.LatencyMs
    If resp.Status < 200 Or resp.Status >= 300 Then
        r.OK = False
        r.ErrorMessage = "HTTP " & resp.Status & " " & Left$(resp.Body, 500)
        If LenB(resp.ErrorMessage) > 0 Then r.ErrorMessage = r.ErrorMessage & " | " & resp.ErrorMessage
        EmbedAzureOpenAI = r
        Exit Function
    End If

    On Error GoTo ParseFail
    Dim parsed As Object: Set parsed = JsonConverter.ParseJson(resp.Body)
    Dim items As Object: Set items = parsed("data")
    Dim n As Long: n = items.Count
    Dim firstVec As Object: Set firstVec = items(1)("embedding")
    Dim d As Long: d = firstVec.Count

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
    If parsed.Exists("usage") Then r.PromptTokens = CLng(parsed("usage")("prompt_tokens"))
    r.OK = True
    EmbedAzureOpenAI = r
    Exit Function

ParseFail:
    r.OK = False
    r.ErrorMessage = "JSON parse failed: " & Err.Description
    EmbedAzureOpenAI = r
End Function

Private Function ChatAzureOpenAI(ByVal systemPrompt As String, ByVal userPrompt As String) As ChatResult
    Dim r As ChatResult
    Dim mode As String: mode = UCase$(modConfig.GetString("api", "mode", "B"))

    Dim url As String
    Dim headers As Object: Set headers = CreateObject("Scripting.Dictionary")
    BuildAzureChatRequest mode, url, headers
    If LenB(url) = 0 Then
        r.OK = False
        r.ErrorMessage = "Chat URL not configured"
        ChatAzureOpenAI = r
        Exit Function
    End If

    Dim body As String: body = BuildAzureChatBody(systemPrompt, userPrompt)

    Dim resp As HttpResponse
    resp = modHttpClient.PostJson(url, body, headers, 120000)
    r.LatencyMs = resp.LatencyMs
    If resp.Status < 200 Or resp.Status >= 300 Then
        r.OK = False
        r.ErrorMessage = "HTTP " & resp.Status & " " & Left$(resp.Body, 500)
        If LenB(resp.ErrorMessage) > 0 Then r.ErrorMessage = r.ErrorMessage & " | " & resp.ErrorMessage
        ChatAzureOpenAI = r
        Exit Function
    End If

    On Error GoTo ParseFail
    Dim parsed As Object: Set parsed = JsonConverter.ParseJson(resp.Body)
    Dim choices As Object: Set choices = parsed("choices")
    If choices.Count = 0 Then
        r.OK = False
        r.ErrorMessage = "No choices in response"
        ChatAzureOpenAI = r
        Exit Function
    End If
    r.Content = CStr(choices(1)("message")("content"))
    If parsed.Exists("usage") Then
        r.PromptTokens = CLng(parsed("usage")("prompt_tokens"))
        r.CompletionTokens = CLng(parsed("usage")("completion_tokens"))
    End If
    r.OK = True
    ChatAzureOpenAI = r
    Exit Function

ParseFail:
    r.OK = False
    r.ErrorMessage = "JSON parse failed: " & Err.Description
    ChatAzureOpenAI = r
End Function

Private Sub BuildAzureEmbedRequest(ByVal mode As String, ByRef url As String, ByVal headers As Object)
    Select Case mode
        Case "A"
            url = modConfig.GetString("api", "relay_url", "") & "/embeddings"
            ApplyRelayAuth headers
        Case Else
            url = BuildAzureUrl(modConfig.GetString("api", "embed_deployment", ""), "embeddings")
            ApplyAzureAuth mode, headers
    End Select
End Sub

Private Sub BuildAzureChatRequest(ByVal mode As String, ByRef url As String, ByVal headers As Object)
    Select Case mode
        Case "A"
            url = modConfig.GetString("api", "relay_url", "") & "/chat/completions"
            ApplyRelayAuth headers
        Case Else
            url = BuildAzureUrl(modConfig.GetString("api", "chat_deployment", ""), "chat/completions")
            ApplyAzureAuth mode, headers
    End Select
End Sub

Private Function BuildAzureUrl(ByVal deployment As String, ByVal route As String) As String
    Dim endpoint As String, ver As String
    endpoint = modConfig.GetString("api", "endpoint", "")
    ver = modConfig.GetString("api", "api_version", "2024-10-21")
    If LenB(endpoint) = 0 Or LenB(deployment) = 0 Then Exit Function
    If Right$(endpoint, 1) = "/" Then endpoint = Left$(endpoint, Len(endpoint) - 1)
    BuildAzureUrl = endpoint & "/openai/deployments/" & deployment & "/" & route & "?api-version=" & ver
End Function

Private Sub ApplyAzureAuth(ByVal mode As String, ByVal headers As Object)
    Select Case mode
        Case "C": headers("Authorization") = "Bearer " & modKeyVault.GetKey()
        Case Else: headers("api-key") = modKeyVault.GetKey()
    End Select
End Sub

Private Sub ApplyRelayAuth(ByVal headers As Object)
    Dim tok As String: tok = modKeyVault.GetKey()
    If LenB(tok) > 0 Then headers("Authorization") = "Bearer " & tok
End Sub

Private Function BuildAzureEmbedBody(ByRef texts() As String) As String
    Dim sb As String, i As Long
    sb = "{""input"":["
    For i = LBound(texts) To UBound(texts)
        If i > LBound(texts) Then sb = sb & ","
        sb = sb & JsonString(texts(i))
    Next i
    sb = sb & "]}"
    BuildAzureEmbedBody = sb
End Function

Private Function BuildAzureChatBody(ByVal systemPrompt As String, ByVal userPrompt As String) As String
    Dim sb As String
    sb = "{""messages"":["
    sb = sb & "{""role"":""system"",""content"":" & JsonString(systemPrompt) & "},"
    sb = sb & "{""role"":""user"",""content"":" & JsonString(userPrompt) & "}"
    sb = sb & "],""temperature"":0.2}"
    BuildAzureChatBody = sb
End Function

' Shared JSON string escaper
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
