Attribute VB_Name = "modMockRibbon"
Option Explicit

' ============================================================================
' modMockRibbon - ニセリボンちゃん(Windows実機テスト用スタブアドイン)
' ----------------------------------------------------------------------------
' 目的:
'   本物の社内AIリボンが無い個人Windows PCでも、「マイ本棚AI」の全配管
'   (Application.Run呼び出し・引数の型と順序・戻り値の形式・Word起動・
'   クリップボード画像・アドイン検出)を本物のWindows Excelで検証するための、
'   確定台帳 RIBBON_API_CONFIRMED.md §1 と同一シグネチャのスタブ実装。
'
' 使い方(docs参照): このモジュールを空のブックにインポートし、
'   「リボンちゃん(検証用).xlam」という名前でExcelアドインとして保存→有効化する。
'   ファイル名に「リボンちゃん」を含めること(マイ本棚AIのアドイン検出は
'   アドイン名の部分一致のため)。
'
' 注意:
'   ・AIの回答品質は検証できない(全て決め打ちのダミー応答)。
'   ・ConvToJpegは実変換しない(パス受け渡しの配管確認のみ)。
'   ・本物のリボンちゃんと同居させないこと(会社PCには入れない)。
' ============================================================================

Private Const MOCK_TAG As String = "(ニセリボン応答)"

' --- 0) 起動制御 -----------------------------------------------------------

Public Function LimitCheck() As Boolean
    ' True=続行不可 の解釈(台帳§2 D3)。スタブは常に続行可。
    LimitCheck = False
End Function

' --- 1) ChatGPT (台帳§1 #1 + 互換後方追加のeffort/verbosity) ---------------

Public Function ChatGPT(ByVal Text As String, _
        Optional ByVal roleSystem As String, _
        Optional ByVal Temperature As Double = 0.4, _
        Optional ByVal MaxTokens As Long = 4096, _
        Optional ByVal Wait As Long = 120, _
        Optional ByVal optModel As String, _
        Optional ByVal prevU As String, _
        Optional ByVal prevA As String, _
        Optional ByVal toolN As String, _
        Optional ByVal reasoning_effort As String, _
        Optional ByVal verbosity As String) As String

    Dim s As String
    s = MOCK_TAG & " model=" & optModel & " tool=" & toolN & vbLf

    ' 会話履歴が渡ってきたことを応答に反映する(prevU/prevA配管の検証点)
    If LenB(prevU) > 0 Then
        s = s & "前回の質問「" & FirstToken(prevU) & "」を踏まえた続きの回答です。" & vbLf
    End If

    s = s & "ご質問について、本棚の資料から確認しました。" & vbLf & _
            "・ここに回答本文が入ります [本棚:サンプル資料.pdf p.1]" & vbLf & _
            "・(effort=" & reasoning_effort & " verbosity=" & verbosity & ")"

    ' プロンプトが深掘り候補を要求していればマーカーを付ける(D11配管の検証点)
    If InStr(Text, "FOLLOWUP") > 0 Then
        s = s & vbLf & "[[FOLLOWUP: (ニセ)この手続きの期限は? | (ニセ)例外になるケースは?]]"
    End If

    ChatGPT = s
End Function

' prevU(新しい順;;;区切り)の先頭要素を返す
Private Function FirstToken(ByVal joined As String) As String
    Dim p As Long
    p = InStr(joined, ";;;")
    If p > 0 Then
        FirstToken = Left$(joined, p - 1)
    Else
        FirstToken = joined
    End If
    If Len(FirstToken) > 40 Then FirstToken = Left$(FirstToken, 40) & "…"
End Function

' --- 2) Embeddings (台帳§1 #2〜#4) -----------------------------------------

Public Function GetEmbeddings(ByVal Text As String) As String
    ' 同じテキストからは常に同じ1536次元ベクトル(カンマ区切り)を返す決定的スタブ
    Const DIMS As Long = 1536
    Dim seed As Double
    Dim i As Long
    For i = 1 To Len(Text)
        seed = seed + AscW(Mid$(Text, i, 1)) * i
    Next i
    seed = seed - Int(seed / 2147483647#) * 2147483647#
    If seed < 1 Then seed = seed + 12345

    Dim parts() As String
    ReDim parts(0 To DIMS - 1)
    Const A As Double = 1664525#
    Const C As Double = 1013904223#
    Const M As Double = 4294967296#
    For i = 0 To DIMS - 1
        seed = A * seed + C
        seed = seed - Int(seed / M) * M
        parts(i) = Format$((seed / M) * 2# - 1#, "0.000000")
    Next i
    GetEmbeddings = Join(parts, ",")
End Function

Public Function CosineSimilarityN2(ByVal str1 As String, ByVal str2 As String) As Double
    Dim v1() As String, v2() As String
    v1 = Split(str1, ","): v2 = Split(str2, ",")
    Dim n As Long: n = UBound(v1)
    If UBound(v2) < n Then n = UBound(v2)
    Dim dot As Double, n1 As Double, n2 As Double, i As Long, a As Double, b As Double
    For i = 0 To n
        a = Val(v1(i)): b = Val(v2(i))
        dot = dot + a * b: n1 = n1 + a * a: n2 = n2 + b * b
    Next i
    If n1 = 0 Or n2 = 0 Then Exit Function
    CosineSimilarityN2 = dot / (Sqr(n1) * Sqr(n2))
End Function

' --- 4) 画像解析系 (台帳§1 #5〜#9) -----------------------------------------

Public Function ChatGPTV(ByVal Text As String, ByVal imageInputs As String, _
        Optional ByVal roleSystem As String, _
        Optional ByVal resolution As String, _
        Optional ByVal toolN As String) As String
    ' Base64を実際に受け取ったことを応答に反映する(D5/D13配管の検証点)
    Dim imgCount As Long
    imgCount = 1 + (Len(imageInputs) - Len(Replace(imageInputs, ",", ""))) ' カンマ数+1(Base64に,は出ない)
    ChatGPTV = MOCK_TAG & " 画像から読み取ったテキストです。" & vbLf & _
        "(受領: 画像" & imgCount & "枚 / base64長=" & Len(imageInputs) & "字 / resolution=" & resolution & _
        " / tool=" & toolN & ")" & vbLf & _
        "経費精算マニュアル 12ページ" & vbLf & _
        "1. 領収書は原本を添付すること" & vbLf & _
        "2. 提出期限は利用月の翌月5営業日まで"
End Function

Public Function IsImageInCB() As Boolean
    Dim f As Variant
    Dim fmts As Variant
    fmts = Application.ClipboardFormats
    If Not IsArray(fmts) Then Exit Function
    For Each f In fmts
        If f = xlClipboardFormatBitmap Or f = xlClipboardFormatPICT Then
            IsImageInCB = True
            Exit Function
        End If
    Next f
End Function

Public Function Base64FromCB(Optional ByVal Ptn As Long = 0) As String
    ' クリップボード画像を一時Chart経由でjpg書き出しする(VBA定番トリック)。
    ' Ptn=1: jpgパスを返す / Ptn=0: そのjpgをBase64化して返す(台帳§1 #7)。
    Dim tmp As String
    tmp = Environ$("TEMP") & "\mockcb_" & Format$(Now, "yyyymmddhhnnss") & ".jpg"

    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    On Error GoTo Fail
    Dim cht As ChartObject
    Set cht = wb.Worksheets(1).ChartObjects.Add(0, 0, 800, 600)
    cht.Chart.Paste
    cht.Chart.Export tmp, "JPG"
    cht.Delete
    wb.Close SaveChanges:=False

    If Ptn = 1 Then
        Base64FromCB = tmp
    Else
        Base64FromCB = Base64FromFile(tmp)
    End If
    Exit Function
Fail:
    On Error Resume Next
    wb.Close SaveChanges:=False
    Base64FromCB = ""
End Function

Public Function Base64FromFile(ByVal filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.LoadFromFile filePath
    Dim bytes() As Byte
    bytes = stm.Read
    stm.Close

    Dim xml As Object, node As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    Base64FromFile = Replace(Replace(node.Text, vbLf, ""), vbCr, "")
End Function

Public Function ConvToJpeg(ByVal imagePath As String) As String
    ' スタブ: 実変換はせずTempへコピーして拡張子だけ.jpgにする(配管確認用)
    Dim tmp As String
    tmp = Environ$("TEMP") & "\" & Format$(Now, "yyyymmddhhnnss") & ".jpg"
    FileCopy imagePath, tmp
    ConvToJpeg = tmp
End Function

' --- 5) ユーティリティ (台帳§1 #10〜#12) -----------------------------------

Public Sub OpenWordMark(ByVal Text As String)
    ' 本物のWordを起動してテキストを流し込む(Word連携配管の実検証)
    Dim wd As Object
    Set wd = CreateObject("Word.Application")
    wd.Visible = True
    wd.Documents.Add
    wd.Selection.TypeText MOCK_TAG & vbCrLf & Text
End Sub

Public Sub OpenMemo(ByVal Text As String)
    Dim p As String
    p = Environ$("TEMP") & "\mockmemo_" & Format$(Now, "yyyymmddhhnnss") & ".txt"
    Dim n As Integer
    n = FreeFile
    Open p For Output As #n
    Print #n, Text
    Close #n
    Shell "notepad.exe """ & p & """", vbNormalFocus
End Sub

Public Sub CellMarkDown(ByVal rng As Range, Optional ByVal isComment As Boolean = False)
    ' スタブ: セル内の「# 」始まり行を太字化するだけの最小装飾
    ' (Rangeオブジェクトが引数として正しく届くことの検証が主目的)
    Dim s As String
    s = CStr(rng.Value)
    If LenB(s) = 0 Then Exit Sub
    Dim lines() As String
    lines = Split(s, vbLf)
    Dim pos As Long: pos = 1
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Left$(lines(i), 2) = "# " Or Left$(lines(i), 3) = "## " Then
            rng.Characters(pos, Len(lines(i))).Font.Bold = True
        End If
        pos = pos + Len(lines(i)) + 1
    Next i
End Sub
