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

' ----------------------------------------------------------------------------
' TeamCodeOf / DeptOf(2026-08-03 R13-7c): チーム/部の推定。
'   ユーザーIDの末尾が "_" + 英大文字/数字4〜6桁のとき、そこをチームコードと
'   みなす(実機で観測された「氏名_E2T22」規約)。純粋な文字列判定なので
'   COM/Excelに一切触れず、LOの実行テストからそのまま検証できる。
'   アンダースコアが複数含まれる名前(例: "山田_太郎_E2T22")は、InStrRevで
'   【最後の】"_"を基準にする(末尾優先。先頭寄りの"_"に釣られない)。
'   小文字はこの規約に無いため一致しない(全社共通の綴りゆれを増やさない)。
' ----------------------------------------------------------------------------
Public Function TeamCodeOf(ByVal userId As String) As String
    Dim s As String: s = Trim$(userId)
    If LenB(s) = 0 Then Exit Function
    Dim p As Long: p = InStrRev(s, "_")
    If p = 0 Then Exit Function
    Dim suf As String: suf = Mid$(s, p + 1)
    Dim n As Long: n = Len(suf)
    If n < 4 Or n > 6 Then Exit Function
    Dim i As Long
    For i = 1 To n
        Dim c As Long: c = AscW(Mid$(suf, i, 1))
        If Not ((c >= 48 And c <= 57) Or (c >= 65 And c <= 90)) Then Exit Function
    Next i
    TeamCodeOf = suf
End Function

' ----------------------------------------------------------------------------
' IsTeamCode(2026-08-03 R13 F12): 文字列がチームコード規約に合うか。
' ----------------------------------------------------------------------------
' config user_department は自由記述で、実機には「営業部」のような部署名が
' そのまま入る。それを無検証でチームコードとして採用すると、
'   ・DeptOf が「営業部」の先頭3字を部コードとして扱う
'   ・ビーコンの team 列に日本語が乗り、他端末の部別集計と噛み合わない
' という、誰にも気付かれないまま集計だけが狂う状態になる。採用の前にここで
' 規約(TeamCodeOf が末尾サフィックスに課すものと同じ「英大文字/数字4〜6字」)
' を確かめる。加えて、採用の可否を決める場面では「英字だけ」「数字だけ」も
' 弾く: 実機で観測された規約は3字の部コード+2桁の連番("E2T22")であり、
' 英字のみ/数字のみの文字列は部署名やコード断片の混入である可能性が高く、
' 誤って全社集計へ流し込むより採用しない方が害が小さい。
' 純粋な文字列判定なのでCOM/Excelに触れず、LOの実行テストから直接検証できる。
Public Function IsTeamCode(ByVal s As String) As Boolean
    Dim t As String: t = Trim$(s)
    Dim n As Long: n = Len(t)
    If n < 4 Or n > 6 Then Exit Function
    Dim hasAlpha As Boolean, hasDigit As Boolean
    Dim i As Long
    For i = 1 To n
        Dim c As Long: c = AscW(Mid$(t, i, 1))
        If c >= 48 And c <= 57 Then
            hasDigit = True
        ElseIf c >= 65 And c <= 90 Then
            hasAlpha = True
        Else
            Exit Function
        End If
    Next i
    IsTeamCode = (hasAlpha And hasDigit)
End Function

' 部コードはチームコードの先頭3字(例: "E2T22" -> "E2T")。空入力は空を返す。
Public Function DeptOf(ByVal teamCode As String) As String
    If LenB(teamCode) = 0 Then Exit Function
    DeptOf = Left$(teamCode, 3)
End Function

' ----------------------------------------------------------------------------
' MinutesPerSelfsolve(2026-08-03 R13-7d): 「自己解決1件=何分の節約か」の
'   換算係数を config(既定15)から読む唯一の窓口。modStats/modBoard/
'   modDashStat の3箇所に同じ値(15)がPrivate Constとして重複していた
'   (憲章§4-5「同型の問題は共通部品で一度だけ解決する」違反)ものを統合する。
'   置き場所を容量の逼迫していないここ(modP2PIo)にしたのは、3箇所のうち
'   どの層からも「中間層(部品層)→中間層」または「UI層→中間層」の一方向
'   参照で済み、R1(依存層順序)を壊さないため。
' ----------------------------------------------------------------------------
Public Function MinutesPerSelfsolve() As Long
    On Error Resume Next
    MinutesPerSelfsolve = modConfig.GetLong("minutes_per_selfsolve", 15)
    On Error GoTo 0
    If MinutesPerSelfsolve <= 0 Then MinutesPerSelfsolve = 15
End Function

' ----------------------------------------------------------------------------
' Beacon* (2026-08-03 R13-7c): modBoard の統計ビーコン(1行タブ区切り)の
'   組み立てとteam列の取り出し。列数で新旧形式を分岐する部分は間違えると
'   「みんなの節約」が静かに0のままになる箇所なので、modBoard内に埋め込まず
'   ここへ切り出してLOの実行テストで固定する。
'   書式(旧): myId, thanksN, dk, dMin, mk, mMin, yk, yMin            (8列, idx0-7)
'   書式(新): 旧8列 + team                                          (9列, idx0-8)
'   WriteBeacon が末尾へ送信時刻(NowStamp)をさらに1列足すため、ファイル上の
'   実列数は旧10列/新11列相当だが、Split後のUBoundで見るのはteam列の有無だけ。
' ----------------------------------------------------------------------------
Public Function BeaconDataText(ByVal myId As String, ByVal thanksN As Long, _
        ByVal dk As String, ByVal dMin As Long, ByVal mk As String, ByVal mMin As Long, _
        ByVal yk As String, ByVal yMin As Long, ByVal teamCode As String) As String
    ' R13 F12: team列は必ず SanitizeId を通してから書く。この行はタブ区切りで
    ' あり、team値にタブやCR/LFが混じると読み手側の列がまるごとズレて、
    ' 別人の節約時間が team として解釈される(1文字で集計が壊れる)。
    ' SanitizeId は禁止文字を "_" に置換し64字で打ち切る既存の共通部品。
    BeaconDataText = myId & vbTab & thanksN & vbTab & dk & vbTab & dMin & vbTab & _
        mk & vbTab & mMin & vbTab & yk & vbTab & yMin & vbTab & SanitizeId(teamCode)
End Function

' 旧形式(team列なし。8フィールド+送信時刻でUBound=8)を読んでも例外に
' ならず単に空文字を返す(新形式はteam列が増えるぶんUBound=9以上になる。
' idx8は旧形式では送信時刻なので、team扱いしないよう境界を9で切る)。
Public Function BeaconTeamField(ByRef fields() As String) As String
    On Error Resume Next
    If UBound(fields) >= 9 Then BeaconTeamField = Trim$(fields(8))
    On Error GoTo 0
End Function

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

' ----------------------------------------------------------------------------
' 下請け(2026-07-28): modP2P から WriteUtf8Retry / ReadUtf8Retry をここへ
' 切り出したとき、これらを向こうに置いたままにしていた。実機で
' 「Sub または Function が定義されていません」になる。使う側へ一緒に運ぶ。
' ----------------------------------------------------------------------------
' UTF-8書き出しの実体は modUtilText.WriteTextFileUtf8(2026-07-31 R11-F2で
' 同型実装を1本化)。err番号/説明の受け渡し(呼び出し元のリトライ判定と
' E0705の材料)はそのまま素通しする。
Private Function TryWriteUtf8(ByVal filePath As String, ByVal content As String, _
                              Optional ByRef outErrNum As Long, Optional ByRef outErrDesc As String) As Boolean
    TryWriteUtf8 = modUtilText.WriteTextFileUtf8(filePath, content, outErrNum, outErrDesc)
End Function

' UTF-8読み取りの実体は modUtilText.ReadTextFileUtf8(2026-07-31 R11-F2)。
Private Function TryReadUtf8(ByVal filePath As String, ByRef outText As String, _
                             Optional ByRef outErrNum As Long, Optional ByRef outErrDesc As String) As Boolean
    TryReadUtf8 = modUtilText.ReadTextFileUtf8(filePath, outText, outErrNum, outErrDesc)
End Function

' Timer基準の短時間待機(DoEventsで応答性維持。Sleep API宣言を避けbitness非依存)。
Private Sub WaitMs(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do   ' 日跨ぎ(深夜0時)ガード
    Loop
End Sub
