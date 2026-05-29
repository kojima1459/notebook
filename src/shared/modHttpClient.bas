Attribute VB_Name = "modHttpClient"
Option Explicit

' ============================================================================
' modHttpClient - WinHTTP-based HTTP wrapper
' ----------------------------------------------------------------------------
' Late-binds WinHttp.WinHttpRequest.5.1 to avoid reference issues on user PCs.
' Honors [proxy] section of config.ini (auto / direct / manual).
' Returns raw body string for callers to parse with JsonConverter.
' ============================================================================

Private Const HTTPREQUEST_PROXYSETTING_DEFAULT  As Long = 0
Private Const HTTPREQUEST_PROXYSETTING_PRECONFIG As Long = 0
Private Const HTTPREQUEST_PROXYSETTING_DIRECT    As Long = 1
Private Const HTTPREQUEST_PROXYSETTING_PROXY     As Long = 2

Private Const AUTOLOGON_POLICY_ALWAYS As Long = 0
Private Const AUTOLOGON_POLICY_NEVER  As Long = 2

Public Type HttpResponse
    Status As Long
    Body As String
    LatencyMs As Long
    ErrorMessage As String
End Type

Public Function PostJson(ByVal url As String, _
                         ByVal jsonBody As String, _
                         ByVal headers As Object, _
                         Optional ByVal timeoutMs As Long = 60000) As HttpResponse
    Dim r As HttpResponse
    Dim http As Object
    Dim t0 As Single: t0 = Timer

    On Error GoTo Failed
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    ApplyProxy http
    http.SetTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept", "application/json"

    If Not headers Is Nothing Then
        Dim key As Variant
        For Each key In headers.Keys
            http.SetRequestHeader CStr(key), CStr(headers(key))
        Next key
    End If

    ' Allow transparent proxy auth on corporate networks
    On Error Resume Next
    http.Option(6) = True ' WinHttpRequestOption_EnableRedirects
    http.SetAutoLogonPolicy AUTOLOGON_POLICY_ALWAYS
    On Error GoTo Failed

    http.Send jsonBody
    r.Status = http.Status
    r.Body = http.ResponseText
    r.LatencyMs = CLng((Timer - t0) * 1000)
    PostJson = r
    Exit Function

Failed:
    r.Status = -1
    r.Body = ""
    r.LatencyMs = CLng((Timer - t0) * 1000)
    r.ErrorMessage = Err.Description
    PostJson = r
End Function

Public Function GetText(ByVal url As String, _
                        Optional ByVal timeoutMs As Long = 60000) As HttpResponse
    Dim r As HttpResponse
    Dim http As Object
    Dim t0 As Single: t0 = Timer

    On Error GoTo Failed
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    ApplyProxy http
    http.SetTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    http.Open "GET", url, False
    On Error Resume Next
    http.SetAutoLogonPolicy AUTOLOGON_POLICY_ALWAYS
    On Error GoTo Failed
    http.Send
    r.Status = http.Status
    r.Body = http.ResponseText
    r.LatencyMs = CLng((Timer - t0) * 1000)
    GetText = r
    Exit Function

Failed:
    r.Status = -1
    r.Body = ""
    r.LatencyMs = CLng((Timer - t0) * 1000)
    r.ErrorMessage = Err.Description
    GetText = r
End Function

Public Function DownloadBinary(ByVal url As String, _
                               ByVal targetPath As String, _
                               Optional ByVal timeoutMs As Long = 120000) As Boolean
    Dim http As Object
    On Error GoTo Failed
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    ApplyProxy http
    http.SetTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    http.Open "GET", url, False
    On Error Resume Next
    http.SetAutoLogonPolicy AUTOLOGON_POLICY_ALWAYS
    On Error GoTo Failed
    http.Send
    If http.Status < 200 Or http.Status >= 300 Then
        DownloadBinary = False
        Exit Function
    End If

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile targetPath, 2 ' Overwrite
    stream.Close
    DownloadBinary = True
    Exit Function

Failed:
    DownloadBinary = False
End Function

Private Sub ApplyProxy(ByVal http As Object)
    Dim mode As String
    mode = LCase$(modConfig.GetString("proxy", "mode", "auto"))
    Select Case mode
        Case "direct"
            http.SetProxy HTTPREQUEST_PROXYSETTING_DIRECT
        Case "manual"
            Dim url As String, bypass As String
            url = modConfig.GetString("proxy", "proxy_url", "")
            bypass = modConfig.GetString("proxy", "bypass_list", "<local>")
            If Len(url) > 0 Then
                http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, url, bypass
            Else
                http.SetProxy HTTPREQUEST_PROXYSETTING_PRECONFIG
            End If
        Case Else ' auto
            http.SetProxy HTTPREQUEST_PROXYSETTING_PRECONFIG
    End Select
End Sub
