Attribute VB_Name = "modTelemetry"
Option Explicit

' ============================================================================
' modTelemetry - 利用データの集計と、匿名フィードバック
' ----------------------------------------------------------------------------
' 何を集めるか(プロダクト改善のため):
'   ・画面別の表示回数(Hub/チャット/ナレッジ/ダッシュボード)
'   ・質問回数、モード別、解決/微妙/違うの比率
'   ・所属部署(config user_department)と利用者名
'
' 何を集めないか(ここが信頼の分かれ目):
'   ・質問の全文は共有側へ送らない。先頭40字までに切る。
'     傾向の把握には十分で、個別の中身を覗く用途には足りない粒度にしてある。
'   ・共有側に送っていることは「使い方」シートに明記する。黙って集めない。
'   ・ローカルの usage_log には従来どおり全文が残る(自分の履歴は自分のもの)。
'
' 送り方:
'   1人1ファイルを日次で上書きするだけ(telemetry\<user>_<yyyymm>.txt)。
'   追記ではなく上書きなので、何度書いても行が増えず、競合も起きない。
'   朝の一斉起動と重ならないよう、送信は終了時(Auto_Close)に行う。
'
' 匿名フィードバック:
'   既存のご意見箱はメール送信のため実名になる。率直な意見は実名では出にくい。
'   feedback\ に、利用者IDを含まないファイル名(乱数)で書き込む。
'   誰が書いたかは記録しないし、復元もできない。
' ============================================================================

Private Const TELE_SUBDIR As String = "telemetry"
Private Const FB_SUBDIR As String = "feedback"
Private Const Q_PREVIEW_CHARS As Long = 40

' ----------------------------------------------------------------------------
' TrackScreen - 画面表示を1件記録する(ローカルのmy_statsに貯めるだけ)。
'   共有フォルダI/Oはここでは行わない。画面遷移のたびにネットワークへ
'   書きに行くと、遷移が目に見えて遅くなる。
' ----------------------------------------------------------------------------
Public Sub TrackScreen(ByVal screenName As String)
    On Error Resume Next
    modStats.Bump "scr:" & screenName
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Publish - 集めた統計を共有フォルダへ1ファイル上書きする。
'   Auto_Close から呼ぶ(終了時なので数秒かかっても業務を止めない)。
' ----------------------------------------------------------------------------
Public Sub Publish()
    On Error Resume Next
    If Not modConfig.GetBool("telemetry_enabled", True) Then Exit Sub

    Dim dirPath As String: dirPath = SubDir(TELE_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Sub
    EnsureDir dirPath

    Dim uid As String: uid = SafeUserId()
    If LenB(uid) = 0 Then Exit Sub

    Dim sb As String
    sb = "v1" & vbTab & uid & vbTab & AuthorName() & vbTab & DeptName() & vbTab & _
         Left$(modUtilText.IsoDateTime(Now), 16) & vbLf
    ' モード別は集計側で内訳を見たいので残す。総計(ask_total)を別行で足すのは、
    ' 2026-08-03(R14-1a)まで ask_quick+ask_deep しか送っておらず、入念モードの
    ' 質問が組織集計から丸ごと欠けていたため(実機第3報 RC1)。モードが増えても
    ' 集計側を直さずに済むよう、総計は modStats.AskTotalAll を1行で送る。
    sb = sb & "ask_quick" & vbTab & modStats.GetStat("ask_quick_total") & vbLf
    sb = sb & "ask_deep" & vbTab & modStats.GetStat("ask_deep_total") & vbLf
    sb = sb & "ask_thorough" & vbTab & modStats.GetStat("ask_thorough_total") & vbLf
    sb = sb & "ask_total" & vbTab & modStats.AskTotalAll() & vbLf
    sb = sb & "solved" & vbTab & modStats.GetStat("selfsolve_total") & vbLf
    sb = sb & "unsure" & vbTab & modStats.GetStat("unsure_total") & vbLf
    sb = sb & "wrong" & vbTab & modStats.GetStat("fail_total") & vbLf
    sb = sb & "correction" & vbTab & modStats.GetStat("correction_total") & vbLf
    sb = sb & "saved_min" & vbTab & modStats.SavedMinutesEstimate() & vbLf
    sb = sb & "streak" & vbTab & modStats.GetStat("streak_days") & vbLf
    sb = sb & "level" & vbTab & modStats.Level() & vbLf
    sb = sb & "scr_hub" & vbTab & modStats.GetStat("scr:hub") & vbLf
    sb = sb & "scr_chat" & vbTab & modStats.GetStat("scr:chat") & vbLf
    sb = sb & "scr_knowledge" & vbTab & modStats.GetStat("scr:knowledge") & vbLf
    sb = sb & "scr_dash" & vbTab & modStats.GetStat("scr:dash") & vbLf

    ' 2026-07-28(レビュー M-19): ファイル名から月を外した。
    ' 送っているのは【生涯累計】のスナップショットなのに、ファイルを
    ' 「ユーザー×月」で分けていたため、2ヶ月目以降は同じ人の累計が
    ' 月数ぶん足し合わされ、利用者数・質問総数・節約時間が過大に出ていた
    ' (集計側は素直に全ファイルを足す)。1人1ファイルの上書きにする。
    Dim outPath As String
    outPath = dirPath & SafeName(uid) & ".txt"
    WriteShared outPath, sb

    ' 旧形式(<uid>_yyyymm.txt)が残っていると二重計上が続くので、
    ' 自分のぶんだけ片付ける。他人のファイルには触らない。
    CleanupLegacyMonthlyFiles dirPath, SafeName(uid)
    On Error GoTo 0
End Sub

' 旧「ユーザー×月」形式の自分のファイルを消す(移行用・レビュー M-19)。
' 列挙中に Kill すると列挙が壊れるので「集めてから消す」。
Private Sub CleanupLegacyMonthlyFiles(ByVal dirPath As String, ByVal uidSafe As String)
    On Error Resume Next
    Dim names() As String: ReDim names(0 To 63)
    Dim n As Long
    Dim fn As String: fn = Dir(dirPath & uidSafe & "_*.txt")
    Do While LenB(fn) > 0
        If n > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 64)
        names(n) = fn
        n = n + 1
        fn = Dir()
    Loop
    Dim i As Long
    For i = 0 To n - 1
        Kill dirPath & names(i)
        Err.Clear
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' SendAnonymousFeedback - 匿名の投書。利用者IDは一切書かない。
'   ファイル名も乱数なので、誰が書いたかは後から辿れない。
' ----------------------------------------------------------------------------
Public Function SendAnonymousFeedback(ByVal bodyText As String) As Boolean
    On Error Resume Next
    If LenB(Trim$(bodyText)) = 0 Then Exit Function

    Dim dirPath As String: dirPath = SubDir(FB_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Function
    EnsureDir dirPath

    ' 本文と日付だけ。所属も名前も書かない(書けば匿名ではなくなる)。
    Dim sb As String
    sb = "v1" & vbTab & Left$(modUtilText.IsoDateTime(Now), 16) & vbTab & _
         Clean1(bodyText)

    Randomize
    Dim fname As String
    fname = "fb_" & Format$(Now, "yyyymmdd") & "_" & _
            Format$(Int(Rnd() * 1000000), "000000") & ".txt"
    SendAnonymousFeedback = WriteShared(dirPath & fname, sb)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' オーナー向け集計。共有フォルダのtelemetryを読んで要約テキストを返す。
' ----------------------------------------------------------------------------
Public Function SummaryText() As String
    On Error Resume Next
    Dim dirPath As String: dirPath = SubDir(TELE_SUBDIR)
    If LenB(dirPath) = 0 Then
        ' 2026-07-31(R8b B14): SubDir は modShare の関所越しなので、
        ' 空が返る理由は「未設定」と「今は届かない」の両方がある。
        SummaryText = "共有フォルダが未設定か、共有フォルダに接続できません。" & vbLf & _
                      "(設定済みの場合は、社内ネットワークへの接続をご確認ください)"
        Exit Function
    End If
    If Len(Dir(dirPath, vbDirectory)) = 0 Then
        SummaryText = "まだ利用データが届いていません。"
        Exit Function
    End If

    Dim users As Long, askTotal As Long, solved As Long, savedMin As Long
    Dim wrong As Long, corr As Long
    Dim deptSummary As String

    Dim f As String: f = Dir(dirPath & "*.txt")
    Dim guard As Long
    Do While LenB(f) > 0 And guard < 500
        guard = guard + 1
        Dim txt As String
        If ReadShared(dirPath & f, txt) Then
            users = users + 1
            askTotal = askTotal + Val(FieldOf(txt, "ask_quick")) + Val(FieldOf(txt, "ask_deep"))
            solved = solved + Val(FieldOf(txt, "solved"))
            wrong = wrong + Val(FieldOf(txt, "wrong"))
            corr = corr + Val(FieldOf(txt, "correction"))
            savedMin = savedMin + Val(FieldOf(txt, "saved_min"))

            Dim d As String: d = HeaderDept(txt)
            If LenB(d) > 0 Then
                If InStr(1, deptSummary, "[" & d & "]") = 0 Then deptSummary = deptSummary & "[" & d & "]"
            End If
        End If
        f = Dir()
    Loop

    Dim sb As String
    sb = "■ 利用状況(共有フォルダに届いている分)" & vbLf & _
         "  利用者数: " & users & " 名" & vbLf & _
         "  質問の総数: " & askTotal & " 件" & vbLf & _
         "  解決した: " & solved & " 件" & vbLf & _
         "  違うと報告: " & wrong & " 件" & vbLf & _
         "  正しい内容の提供: " & corr & " 件" & vbLf & _
         "  節約した時間の合計: " & (savedMin \ 60) & " 時間 " & (savedMin Mod 60) & " 分" & vbLf & vbLf
    If askTotal > 0 Then
        sb = sb & "  解決率: " & CLng(solved * 100# / askTotal) & "%" & vbLf & vbLf
    End If
    sb = sb & "■ 参加している部署" & vbLf & "  " & Replace(deptSummary, "][", "] [") & vbLf
    SummaryText = sb
    On Error GoTo 0
End Function

' 匿名フィードバックの一覧(新しい順・最大20件)。
Public Function FeedbackText() As String
    On Error Resume Next
    Dim dirPath As String: dirPath = SubDir(FB_SUBDIR)
    If LenB(dirPath) = 0 Then Exit Function
    If Len(Dir(dirPath, vbDirectory)) = 0 Then
        FeedbackText = "まだ届いていません。"
        Exit Function
    End If

    Dim names() As String: ReDim names(0 To 255)
    Dim n As Long
    Dim f As String: f = Dir(dirPath & "fb_*.txt")
    Do While LenB(f) > 0 And n < 256
        names(n) = f
        n = n + 1
        f = Dir()
    Loop
    If n = 0 Then
        FeedbackText = "まだ届いていません。"
        Exit Function
    End If

    Dim a As Long, b As Long
    For a = 0 To n - 2
        For b = 0 To n - 2 - a
            If names(b) < names(b + 1) Then
                Dim t As String: t = names(b): names(b) = names(b + 1): names(b + 1) = t
            End If
        Next b
    Next a

    Dim sb As String
    Dim i As Long
    For i = 0 To n - 1
        If i >= 20 Then Exit For
        Dim txt As String
        If ReadShared(dirPath & names(i), txt) Then
            Dim parts() As String: parts = Split(txt, vbTab)
            If UBound(parts) >= 2 Then
                sb = sb & "・" & parts(1) & vbLf & "  " & parts(2) & vbLf & vbLf
            End If
        End If
    Next i
    FeedbackText = sb
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' 内部
' ----------------------------------------------------------------------------
Private Function FieldOf(ByVal txt As String, ByVal key As String) As String
    Dim lines_() As String: lines_ = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    Dim i As Long
    For i = LBound(lines_) To UBound(lines_)
        If Left$(lines_(i), Len(key) + 1) = key & vbTab Then
            FieldOf = Mid$(lines_(i), Len(key) + 2)
            Exit Function
        End If
    Next i
End Function

Private Function HeaderDept(ByVal txt As String) As String
    Dim lines_() As String: lines_ = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    If UBound(lines_) < 0 Then Exit Function
    Dim parts() As String: parts = Split(lines_(0), vbTab)
    If UBound(parts) >= 3 Then HeaderDept = Trim$(parts(3))
End Function

' 到達性の判定は modShare が1セッション1回だけ行う(レビュー M-20)。
' 終了時のテレメトリ送信がここで数十秒ブロックすると、
' 「閉じたのにExcelが残る」という最悪の別れ方になる。
Private Function SubDir(ByVal leaf As String) As String
    SubDir = modShare.SubDir(leaf)
End Function

Private Sub EnsureDir(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Function SafeUserId() As String
    On Error Resume Next
    SafeUserId = modP2P.CurrentUserId()
    On Error GoTo 0
End Function

Private Function AuthorName() As String
    On Error Resume Next
    AuthorName = Trim$(modConfig.GetString("pack_author", ""))
    On Error GoTo 0
End Function

Private Function DeptName() As String
    On Error Resume Next
    DeptName = Trim$(modConfig.GetString("user_department", ""))
    On Error GoTo 0
End Function

Private Function SafeName(ByVal s As String) As String
    Dim t As String: t = s
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ")
        t = Replace(t, CStr(bad), "_")
    Next bad
    SafeName = modUtil.SafeLeft(t, 32)
End Function

' 共有側へ出す文字列は改行とタブを潰し、長さも切る。
' 質問文の全文が共有フォルダに残らないようにするための最後の関門。
Private Function Clean1(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    Clean1 = modUtil.SafeLeft(Trim$(t), 1000)
End Function

' 質問文を共有用に切り詰める(先頭40字)。傾向把握には足り、
' 個別の中身を覗くには足りない粒度。
Public Function QuestionPreview(ByVal q As String) As String
    QuestionPreview = modUtil.SafeLeft(Trim$(q), Q_PREVIEW_CHARS)
End Function

Private Function WriteShared(ByVal filePath As String, ByVal content As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To 3
        If TryWrite(filePath, content) Then
            WriteShared = True
            Exit Function
        End If
        Wait_ 200 * attempt
    Next attempt
End Function

' UTF-8書き出しの実体は modUtilText.WriteTextFileUtf8(2026-07-31 R11-F2)。
Private Function TryWrite(ByVal filePath As String, ByVal content As String) As Boolean
    TryWrite = modUtilText.WriteTextFileUtf8(filePath, content)
End Function

' UTF-8読み取りの実体は modUtilText.ReadTextFileUtf8(2026-07-31 R11-F2)。
Private Function ReadShared(ByVal filePath As String, ByRef outText As String) As Boolean
    ReadShared = modUtilText.ReadTextFileUtf8(filePath, outText)
End Function

Private Sub Wait_(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do
    Loop
End Sub

' ============================================================================
' 組織集計スナップショットの引き継ぎ(2026-08-16 R33H M12)
' ----------------------------------------------------------------------------
' 【症状】F18 で走査窓を回すようにした結果、スナップショットに載るのは
'   その回の窓に入った人だけの合計になった。称号(累積)は引き継いだが、
'   組織合計(今日/今月/今年)と部別(D行)は引き継いでいないため、
'   「今月 約1,200時間」が翌日「約780時間」へ【減る】。月間累計として
'   ありえない動きを全端末が見る。
'
' 【規則】(ユーザー裁定: 前回の値を引き継いで大きい方を残す)
'   ・前回の値と今回の値の大きい方を採る(称号と同じ考え方=単調増加)。
'   ・ただし引き継ぎ元のキーが今のキーと違えば【1つも引き継がない】。
'     日・月・年をそれぞれ独立に判定する。日が替われば今日だけ0から、
'     月が替われば今日と今月が0から、年が替われば3つとも0から。
'     12/31→1/1 では3つが同時に替わるが、独立判定なので取りこぼしが無い。
'     ここを1つのフラグでまとめると、月が替わっても先月の値が max で
'     生き残り【今月が永久に減らない】= 元の症状より重い壊れ方になる。
'   ・D行(部別)はヘッダに自分のキーを持たないが、中身は「今月ぶんの合算」
'     (modBoard.BuildSnapshot が f(4)=月キーに一致する行の f(5) だけを
'     足している)。したがって月キーで一括して判定してよい ―― 月が替われば
'     部別も1行も引き継がない。
'   ・引き継ぎ元は「読めた回」だけ(M7 の見送りと整合)。読めなければ
'     modShare.CarryText() が空を返し、この経路は素通しになる。
'   ・利用者の操作は一切要らない(発行者端末が集計を作るたびに自動で効く)。
'
' 【なぜこのモジュールに在るのか】容量(憲章§4-6)。書式の持ち主 modShare は
'   残207字、呼び口の modBoard は残36字で1文字も入らない。集計値の意味づけを
'   扱うこのモジュールが統計層で最も主題が近く、modBoard 側は呼び先の
'   モジュール名1語の差し替えだけで届く。
' ============================================================================

' BoardCarryNum - 引き継ぎの1マス【純関数】。
'   prevKey/curKey が違えば引き継がない(今回の値をそのまま返す)。
'   同じなら大きい方を返す。curKey が空(壊れた入力)なら引き継がない。
Public Function BoardCarryNum(ByVal prevKey As String, ByVal curKey As String, _
                              ByVal prevVal As Long, ByVal curVal As Long) As Long
    BoardCarryNum = curVal
    If LenB(curKey) = 0 Then Exit Function
    If StrComp(prevKey, curKey, vbBinaryCompare) <> 0 Then Exit Function
    If prevVal > curVal Then BoardCarryNum = prevVal
End Function

' BoardBodyCarry - 前回ぶんを引き継いでから本文を組み立てる(modBoard の
'   BuildSnapshot が modShare.BoardBodyText の代わりに呼ぶ唯一の入口)。
Public Function BoardBodyCarry(ByVal headLine As String, ByRef deptAgg As Object, _
                               ByRef titleAgg As Object) As String
    Dim prevText As String
    On Error Resume Next
    prevText = modShare.CarryText()
    On Error GoTo 0
    If LenB(prevText) = 0 Then
        BoardBodyCarry = modShare.BoardBodyText(headLine, deptAgg, titleAgg)
        Exit Function
    End If
    Dim prevHead As String: prevHead = modShare.BoardHeadLine(prevText)
    CarryDeptRows prevText, prevHead, headLine, deptAgg
    BoardBodyCarry = modShare.BoardBodyText(BoardCarriedHead(prevHead, headLine), _
                                            deptAgg, titleAgg)
End Function

' BoardCarriedHead - ヘッダの今日/今月/今年を引き継いだ形へ組み直す。
'   列番号は modShare.BoardHeadText の見出しにある11列の定義。
Public Function BoardCarriedHead(ByVal prevHead As String, ByVal headLine As String) As String
    BoardCarriedHead = headLine
    On Error GoTo Bad
    Dim dk As String: dk = modShare.BoardHeadField(headLine, 2)
    Dim mk As String: mk = modShare.BoardHeadField(headLine, 4)
    Dim yk As String: yk = modShare.BoardHeadField(headLine, 6)
    If LenB(dk) = 0 Or LenB(mk) = 0 Or LenB(yk) = 0 Then Exit Function
    Dim d As Long, m As Long, y As Long
    d = BoardCarryNum(modShare.BoardHeadField(prevHead, 2), dk, _
        modShare.BoardNum(modShare.BoardHeadField(prevHead, 3)), _
        modShare.BoardNum(modShare.BoardHeadField(headLine, 3)))
    m = BoardCarryNum(modShare.BoardHeadField(prevHead, 4), mk, _
        modShare.BoardNum(modShare.BoardHeadField(prevHead, 5)), _
        modShare.BoardNum(modShare.BoardHeadField(headLine, 5)))
    y = BoardCarryNum(modShare.BoardHeadField(prevHead, 6), yk, _
        modShare.BoardNum(modShare.BoardHeadField(prevHead, 7)), _
        modShare.BoardNum(modShare.BoardHeadField(headLine, 7)))
    BoardCarriedHead = modShare.BoardHeadText(modShare.BoardHeadField(headLine, 1), _
        dk, d, mk, m, yk, y, _
        modShare.BoardNum(modShare.BoardHeadField(headLine, 8)), _
        (modShare.BoardHeadField(headLine, 9) = "1"), _
        modShare.BoardNum(modShare.BoardHeadField(headLine, 10)))
Bad:
End Function

' BoardCarryDeptOk - D行(部別)を引き継いでよいか【純関数】。D行はヘッダに
'   自分のキーを持たないが、中身は「今月ぶんの合算」(modBoard.BuildSnapshot が
'   f(4)=月キーに一致する行の f(5) だけを足す)なので、月キーで一括判定できる。
Public Function BoardCarryDeptOk(ByVal prevHead As String, ByVal headLine As String) As Boolean
    Dim mk As String: mk = modShare.BoardHeadField(headLine, 4)
    If LenB(mk) = 0 Then Exit Function
    BoardCarryDeptOk = (StrComp(modShare.BoardHeadField(prevHead, 4), mk, vbBinaryCompare) = 0)
End Function

' CarryDeptRows - 前回の D行(部別=今月ぶん)を deptAgg へ引き継ぐ。
'   月キーが違えば1行も引き継がない(月が替われば部別も0から積み直す)。
Private Sub CarryDeptRows(ByVal prevText As String, ByVal prevHead As String, _
                          ByVal headLine As String, ByRef deptAgg As Object)
    If deptAgg Is Nothing Then Exit Sub
    If Not BoardCarryDeptOk(prevHead, headLine) Then Exit Sub
    On Error GoTo Bad
    Dim rows() As String: rows = Split(Replace$(prevText, vbCrLf, vbLf), vbLf)
    Dim i As Long
    For i = 1 To UBound(rows)
        Dim c() As String: c = Split(rows(i), vbTab)
        If UBound(c) >= 2 Then
            If c(0) = "D" Then
                Dim k As String: k = Trim$(c(1))
                Dim v As Long: v = modShare.BoardNum(c(2))
                If LenB(k) > 0 Then
                    If deptAgg.Exists(k) Then
                        If v > CLng(deptAgg(k)) Then deptAgg(k) = v
                    Else
                        deptAgg(k) = v
                    End If
                End If
            End If
        End If
    Next i
Bad:
End Sub
