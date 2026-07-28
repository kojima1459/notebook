Attribute VB_Name = "modP2PIo"
Option Explicit

' ========================================
' modP2PIo - 共有フォルダへの小さなファイルI/Oと、ID/文字列の正規化
'
' modP2P から切り出した下請け。UTF-8での読み書き(リトライ付き)、削除、
' フォルダ作成、ファイル名に使えるIDへの変換、nonce の生成。
' 感謝状・品質報告・専門家への質問など複数の機能が同じ作法を要求するので、
' 作法そのものをここ1箇所に置く。
'
' 切り出しの理由(2026-07-28): modP2P が契約上限30,000字に対し残り450字と
' なり、次の修正が入らなくなっていた(レビュー I-2)。
' ========================================

Public Function CnFromDn(ByVal dn As String) As String
    If LenB(dn) = 0 Then Exit Function
    Dim p As Long: p = InStr(1, dn, "CN=", vbTextCompare)
    If p = 0 Then Exit Function
    Dim rest As String: rest = Mid$(dn, p + 3)
    Dim c As Long: c = InStr(rest, ",")
    If c > 0 Then rest = Left$(rest, c - 1)
    CnFromDn = Trim$(rest)
End Function

Public Function SanitizeId(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", ",", vbTab, vbCr, vbLf)
    Dim t As String: t = s
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i
    SanitizeId = modUtil.SafeLeft(Trim$(t), 64)
End Function

Public Sub EnsureDir(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

Public Function NewNonce(ByVal myId As String) As String
    NewNonce = myId & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & _
               Format$(Int(Timer * 1000) Mod 100000, "00000")
End Function

Public Function SanitizeField(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    SanitizeField = t
End Function

' 3回リトライしても書けなかった場合、握りつぶさずE0705として記録する
' (SRE監査Phase2.1の方針をEmitNoiseVote以外の全書込み経路にも拡張。
' 実機環境の壁(共有フォルダのアクセス権限・セキュリティソフトのブロック等=
' エラー52/70想定)を、利用者が🩺診断の「コピー」ボタンでそのまま
' 開発者へ伝えられるようにするため、err_number構造化フィールドに残す)。
' context: 実際に呼んでいるPublic関数名を "modP2P.XxxYyy" 形式で呼び出し元が渡す
' (LogErrorのcontext引数は「modX.Y」形式だとYがPublicか契約チェックされるため
' 「§7」、Private助手関数自身の名前は使えない。呼び出し元ごとに正しい実体を
' 渡すことで、どの機能から起きた失敗かerr_log上で見分けられるようにする)。
Public Function WriteUtf8Retry(ByVal filePath As String, ByVal content As String, _
                                ByVal context As String) As Boolean
    Dim attempt As Long
    Dim lastNum As Long, lastDesc As String
    For attempt = 1 To 3
        If TryWriteUtf8(filePath, content, lastNum, lastDesc) Then
            WriteUtf8Retry = True
            Exit Function
        End If
        WaitMs 250 * attempt   ' 250 / 500 / 750ms バックオフ
    Next attempt
    On Error Resume Next
    modLog.LogError "E0705", context, _
        "3回リトライしても書込み失敗: " & modUtil.SafeLeft(filePath, 300) & " : " & lastDesc, lastNum
    On Error GoTo 0
End Function

Public Function ReadUtf8Retry(ByVal filePath As String, ByRef outText As String, _
                               ByVal context As String) As Boolean
    Dim attempt As Long
    Dim lastNum As Long, lastDesc As String
    For attempt = 1 To 3
        If TryReadUtf8(filePath, outText, lastNum, lastDesc) Then
            ReadUtf8Retry = True
            Exit Function
        End If
        WaitMs 250 * attempt
    Next attempt
    ' エラー53(ファイルが見つかりません)は、列挙後に他ユーザーの並行GCで
    ' ファイルが消えた正常な競合(このモジュール冒頭のKillRetryコメント参照)。
    ' 想定内の自己解決ケースなのでログを汚さない。それ以外(52/70等の
    ' アクセス権限・ネットワーク瞬断)のみE0705として記録する。
    If lastNum <> 53 Then
        On Error Resume Next
        modLog.LogError "E0705", context, _
            "3回リトライしても読込み失敗: " & modUtil.SafeLeft(filePath, 300) & " : " & lastDesc, lastNum
        On Error GoTo 0
    End If
End Function

' 処理済み感謝状の削除(GC)。ロック時はバックオフして最大3回。失敗しても無害
' (nonce重複排除で二重加算は起きない)。err_logは汚さず、診断用にusage_logへ
' だけ記録する(GC失敗は非致命なので🩺診断の「直近のエラー」には出さない)。
Public Function KillRetry(ByVal filePath As String) As Boolean
    Dim attempt As Long
    Dim lastNum As Long
    For attempt = 1 To 3
        On Error Resume Next
        Err.Clear
        ' 並行GC耐性: 複数ユーザーが同時に同じ投票/感謝状をGCすると、後着のKillは
        ' error 53(ファイルなし)になる。「既に無い=削除目的は達成」なので成功扱いにし、
        ' 無駄な3回×最大1.5秒のリトライ(その間ロック保持)を避ける。
        If LenB(Dir(filePath)) = 0 Then
            On Error GoTo 0
            KillRetry = True
            Exit Function
        End If
        Kill filePath
        lastNum = Err.Number
        If lastNum = 0 Then
            On Error GoTo 0
            KillRetry = True
            Exit Function
        End If
        On Error GoTo 0
        WaitMs 250 * attempt
    Next attempt
    On Error Resume Next
    modLog.LogUsage "p2p_gc_failed", "", "3回リトライしても削除失敗 err#" & lastNum & ": " & modUtil.SafeLeft(filePath, 300)
    On Error GoTo 0
End Function
