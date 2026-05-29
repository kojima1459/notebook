Attribute VB_Name = "modKeyVault"
Option Explicit

' ============================================================================
' modKeyVault - API key resolver (3-mode abstraction)
' ----------------------------------------------------------------------------
'   Mode A (relay):  no client-side key; auth is at the relay server.
'   Mode B (direct + obfuscated): key is XOR+Base64 encoded into a hidden
'                                 named range. Decoded only inside this module.
'   Mode C (Entra ID): an OAuth2 access token is obtained per-session and
'                      cached until expiry. Implementation TBD - placeholder.
'
' SECURITY NOTE
' -------------
' Mode B obfuscation prevents casual viewing only. A determined user can
' extract the key from the xlsm by unzipping the package and dumping the
' VBA project. Treat this as "out of sight, not out of reach" and document
' the limitation in docs/security.md. The right answer is Mode A.
' ============================================================================

Private Const KEY_STORAGE_NAME As String = "__nblm_obf_key__"
Private mCachedKey As String

Public Function GetKey() As String
    If LenB(mCachedKey) > 0 Then
        GetKey = mCachedKey
        Exit Function
    End If

    Dim mode As String
    mode = UCase$(modConfig.GetString("api", "mode", "B"))
    Select Case mode
        Case "A"
            ' Relay mode: clients should not need an API key. We may return
            ' a per-user bearer if the relay requires one - left empty here.
            mCachedKey = ""
        Case "C"
            ' Entra ID token cache - to be implemented.
            mCachedKey = ResolveEntraToken()
        Case Else
            mCachedKey = DecodeObfuscated(LoadObfuscatedBlob())
    End Select
    GetKey = mCachedKey
End Function

Public Sub ClearCache()
    mCachedKey = ""
End Sub

' ----------------------------------------------------------------------------
' Mode B: obfuscation store
' ----------------------------------------------------------------------------
' The obfuscated blob lives in a hidden defined name in ThisWorkbook so that
' it travels with the .xlsm file. The plaintext key never touches a sheet.
' Use modKeyEnroller (admin) to set it.

Public Sub StoreObfuscated(ByVal blob As String)
    On Error Resume Next
    ThisWorkbook.Names(KEY_STORAGE_NAME).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=KEY_STORAGE_NAME, _
                           RefersTo:="=""" & EscapeForName(blob) & """", _
                           Visible:=False
End Sub

Private Function LoadObfuscatedBlob() As String
    On Error Resume Next
    Dim nm As Name
    Set nm = ThisWorkbook.Names(KEY_STORAGE_NAME)
    If nm Is Nothing Then
        LoadObfuscatedBlob = ""
        Exit Function
    End If
    Dim ref As String: ref = nm.RefersTo
    ' RefersTo looks like: ="payload"
    Dim q1 As Long, q2 As Long
    q1 = InStr(ref, """")
    q2 = InStrRev(ref, """")
    If q1 > 0 And q2 > q1 Then
        LoadObfuscatedBlob = Mid$(ref, q1 + 1, q2 - q1 - 1)
        LoadObfuscatedBlob = Replace(LoadObfuscatedBlob, """""", """")
    End If
End Function

Private Function EscapeForName(ByVal s As String) As String
    EscapeForName = Replace(s, """", """""")
End Function

' ----------------------------------------------------------------------------
' Encode / decode (XOR + Base64). The XOR pad is derived from a per-workbook
' salt named range so two different distribution builds yield different blobs.
' ----------------------------------------------------------------------------

Public Function EncodeObfuscated(ByVal plaintext As String) As String
    EncodeObfuscated = Base64Encode(XorBytes(StrToBytes(plaintext), Pad()))
End Function

Private Function DecodeObfuscated(ByVal blob As String) As String
    If LenB(blob) = 0 Then
        DecodeObfuscated = ""
        Exit Function
    End If
    DecodeObfuscated = BytesToStr(XorBytes(Base64Decode(blob), Pad()))
End Function

Private Function Pad() As Byte()
    ' Deterministic pad seeded purely from LITERAL_SALT so admin (encoder)
    ' and chatbot (decoder) produce identical bytes. Rotate the salt per
    ' distribution build (build.ps1 can sed it in).
    Const LITERAL_SALT As String = "nblm-v1-salt-please-rotate-in-build"
    Dim h As Long, i As Long
    Dim out(0 To 255) As Byte
    h = 0
    For i = 1 To Len(LITERAL_SALT)
        h = ((h * 131) + Asc(Mid$(LITERAL_SALT, i, 1))) And &H7FFFFFFF
    Next i
    For i = 0 To 255
        h = ((h * 1103515245) + 12345) And &H7FFFFFFF
        out(i) = CByte((h \ 65536) And &HFF)
    Next i
    Pad = out
End Function

Private Function XorBytes(ByRef data() As Byte, ByRef pad() As Byte) As Byte()
    Dim i As Long, n As Long, padLen As Long
    n = UBound(data) - LBound(data) + 1
    padLen = UBound(pad) - LBound(pad) + 1
    Dim out() As Byte
    ReDim out(0 To n - 1)
    For i = 0 To n - 1
        out(i) = data(LBound(data) + i) Xor pad(i Mod padLen)
    Next i
    XorBytes = out
End Function

Private Function StrToBytes(ByVal s As String) As Byte()
    ' UTF-8 encode via ADODB.Stream
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' text
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1 ' binary
    st.Position = 3 ' skip BOM
    StrToBytes = st.Read
    st.Close
End Function

Private Function BytesToStr(ByRef b() As Byte) As String
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write b
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    BytesToStr = st.ReadText
    st.Close
End Function

Private Function Base64Encode(ByRef data() As Byte) As String
    Dim dom As Object, el As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set el = dom.createElement("b64")
    el.DataType = "bin.base64"
    el.nodeTypedValue = data
    Base64Encode = Replace(Replace(el.Text, vbCrLf, ""), vbLf, "")
End Function

Private Function Base64Decode(ByVal s As String) As Byte()
    Dim dom As Object, el As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set el = dom.createElement("b64")
    el.DataType = "bin.base64"
    el.Text = s
    Base64Decode = el.nodeTypedValue
End Function

' ----------------------------------------------------------------------------
' Mode C placeholder
' ----------------------------------------------------------------------------
Private Function ResolveEntraToken() As String
    ' To be implemented in Phase 2 (device code flow or integrated auth).
    Err.Raise vbObjectError + &H6001, "modKeyVault", _
              "Mode C (Entra ID) is not yet implemented. Use Mode A or B."
End Function
