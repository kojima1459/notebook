Attribute VB_Name = "modTestsPure6"
Option Explicit

' ============================================================================
' modTestsPure6 - modTestsPure5の分割先(2026-07-31 R8b: 敵対的レビュー対応)
' ----------------------------------------------------------------------------
' 役割:
'   R8の敵対的レビューで見つかった穴(B1〜B16)のうち、純ロジックで固定できる
'   ものをここで押さえる。入口は modTestsPure5.RunAll5 の末尾から呼ばれる
'   Public Sub RunAll6()。
'   (modTestsPure5 が30,000字上限まで残り約200字となったための分割。
'    modTestsPure3→4→5 と同じ中継方式。)
'
' ここに集めたものの共通点:
'   「一度も再現していないが、起きたら取り返しがつかない」種類の穴。
'   ・B1  失効判定の戻り値が想定外(空文字など)のときに知識を消してしまう
'   ・B7b 端末の時計が共有サーバとずれている人が永久に発行できない
'   ・B10 UNC共有ルートを直接指した人が永久に到達不能になる
'   どれも「特定の環境の人だけが、原因の分からないまま被害を受ける」形で、
'   実機で踏んでも報告として上がってきにくい。だから机上で固定する。
' ============================================================================

' ----------------------------------------------------------------------------
' B1: 失効判定の戻り値のホワイトリスト
' ----------------------------------------------------------------------------
' modGuard.EnforceExpiry は WipeKnowledge の直前で
'   If act <> "wipe" Then Exit Function
' と【ホワイトリスト】で守る。従来はここが「1つずつ弾くブラックリスト」で、
' 既定の振る舞いが「消す」だった。EnforceExpiry の冒頭は On Error Resume Next
' なので、my_stats のセルがエラー値(#REF! 等)で GetStatText が失敗すると
' act が空文字のまま消去へ落ちる。
'
' ここでは「ExpiryDecision が返し得る値の集合」を固定する。呼び出し側の
' ホワイトリストは、この集合のうち "wipe" だけを通す実装であることが前提。
' 集合に知らない値が増えたらこのテストが落ち、ホワイトリストの見直しを促す。
Private Sub TestExpiryDecisionVocabulary()
    Dim reach As String: reach = "2026-06-01"
    Dim seenWipe As Boolean
    Dim bad As String
    Dim L As Long, d As Long, w As Long

    ' limitDays / daysSince / 予告済み / 消去済み を広く振って、戻り値が
    ' 想定した6語のいずれかであることを確かめる。
    For L = 0 To 35
        For d = 0 To 40
            For w = 0 To 1
                Dim warned As String
                Dim wiped As String
                If w = 0 Then
                    warned = ""
                    wiped = ""
                Else
                    warned = CStr(d)
                    wiped = ""
                End If
                Dim act As String
                act = modShareRule.ExpiryDecision(reach, d, L, warned, wiped)
                Select Case act
                    Case "off", "never", "none", "warn", "warn_first"
                        ' 想定内。消さない側。
                    Case "wipe"
                        seenWipe = True
                    Case Else
                        If LenB(bad) = 0 Then _
                            bad = "limit=" & L & " days=" & d & " warned=" & warned & " -> [" & act & "]"
                End Select
            Next w
        Next d
    Next L

    modTestRunner.Check "B1: ExpiryDecisionは想定した6語しか返さない", _
        (LenB(bad) = 0), "想定外の戻り値=" & bad
    modTestRunner.Check "B1: 上の総当たりに wipe が実際に含まれる(テストが空回りしていない)", _
        (seenWipe = True), "wipeが1度も出ていない = 条件の作り方が間違っている"

    ' (C5: ここにあった『空文字は wipe ではない』はリテラル同士の比較で
    '  常に真になるトートロジーだった。modGuard 側のホワイトリスト
    '  If act <> "wipe" Then Exit Function を守るのは上の語彙テストであり、
    '  何も検証していない行を残すとテスト件数だけが増えて安心感を偽装する。)
    modTestRunner.Check "B1: 未到達端末は総当たりでも never のみ", _
        (modShareRule.ExpiryDecision("", 0, 30, "", "") = "never" And _
         modShareRule.ExpiryDecision("", 400, 1, "400", "") = "never"), ""
End Sub

' ----------------------------------------------------------------------------
' 同時発行の見張り(F12)+ B7b: 発行ロックの時計ズレ
' ----------------------------------------------------------------------------
' (F12のロック境界テストは modTestsPure5 の容量逼迫によりこちらへ移設。
'  同じ関数の境界は1箇所にまとめておく方が、次に触る人が見落とさない。)
'
' B7b: ファイルサーバと端末の時計が数分ずれるのは社内では普通にある。
' その端末から見ると他人のロックは常に「未来の時刻」に見え、経過時間が負になる。
' これを wait にすると、その端末だけが永久に発行できない
' (しかも画面には「いま別の方が発行中です」と出るので、本人は待てば直ると思う)。
' 残骸側へ倒し、呼び出し側で err_log(E0807)へ1行残す。
Private Sub TestPublishLockClockSkew()
    ' --- F12: ロックが無い / 新しい / 古い の基本 ---
    modTestRunner.Check "発行ロック: ロック無しなら go", _
        (modShareRule.PublishLockAction(False, 0, 10) = "go"), ""
    modTestRunner.Check "発行ロック: ロック無しは経過時間を見ない", _
        (modShareRule.PublishLockAction(False, 999, 10) = "go"), ""
    modTestRunner.Check "発行ロック: 直後(0分)は wait", _
        (modShareRule.PublishLockAction(True, 0, 10) = "wait"), ""
    modTestRunner.Check "発行ロック: 9.9分は wait", _
        (modShareRule.PublishLockAction(True, 9.9, 10) = "wait"), ""
    ' 境界: ちょうど10分は残骸とみなす(発行は数十秒で終わるため)。
    modTestRunner.Check "発行ロック: ちょうど10分は stale", _
        (modShareRule.PublishLockAction(True, 10, 10) = "stale"), _
        "実際=" & modShareRule.PublishLockAction(True, 10, 10)
    modTestRunner.Check "発行ロック: 60分は stale", _
        (modShareRule.PublishLockAction(True, 60, 10) = "stale"), ""

    ' --- B7b/C2: 経過が負(ロックの更新時刻が未来)は、ズレの大きさで分ける ---
    ' 小さなズレ(-staleMinutes < age < 0)は「たった今誰かが作った」の方が
    ' ありそうなので待つ。踏み潰して同時発行になる方が害が大きい。
    modTestRunner.Check "C2: 小さな時計ズレ(-1分)は wait", _
        (modShareRule.PublishLockAction(True, -1, 10) = "wait"), _
        "実際=" & modShareRule.PublishLockAction(True, -1, 10)
    modTestRunner.Check "C2: 小さな時計ズレ(-5分)は wait", _
        (modShareRule.PublishLockAction(True, -5, 10) = "wait"), _
        "実際=" & modShareRule.PublishLockAction(True, -5, 10)
    modTestRunner.Check "C2: -9.99分は wait(境界の内側)", _
        (modShareRule.PublishLockAction(True, -9.99, 10) = "wait"), ""

    ' 大きなズレ(age <= -staleMinutes)は残骸扱い。ここを wait にすると
    ' 時計が大きく狂った端末が永久に発行できなくなる。
    modTestRunner.Check "C2: -10分ちょうどは stale(境界)", _
        (modShareRule.PublishLockAction(True, -10, 10) = "stale"), _
        "実際=" & modShareRule.PublishLockAction(True, -10, 10)
    modTestRunner.Check "C2: -600分は stale", _
        (modShareRule.PublishLockAction(True, -600, 10) = "stale"), ""

    ' 待たせるのは「ロックがあり、-10分 < 経過 < 10分」のときだけ。
    ' つまり wait の窓は原点をはさんで staleMinutes ぶんずつ、が新仕様。
    Dim m As Long
    Dim badWait As Boolean
    Dim badStale As Boolean
    For m = -20 To 20
        Dim act As String
        act = modShareRule.PublishLockAction(True, CDbl(m), 10)
        If act = "wait" Then
            If m <= -10 Or m >= 10 Then badWait = True
        End If
        If act = "stale" Then
            If m > -10 And m < 10 Then badStale = True
        End If
    Next m
    modTestRunner.Check "C2: waitになるのは -10分 < 経過 < 10分 のときだけ", _
        (badWait = False), "大きな時計ズレや残骸で待たせている"
    modTestRunner.Check "C2: staleになるのは経過が±10分の外だけ", _
        (badStale = False), "発行中のロックを踏み潰す経路がある"
End Sub

' ----------------------------------------------------------------------------
' B10: UNC共有ルート直指定のときのプローブ再試行
' ----------------------------------------------------------------------------
' ProbeTargetPath は GetAttr のために末尾の "\" を1つ落とす。ところが
' 共有ルートを直接指した場合、環境によっては "\\srv\share" では
' 実行時エラー52/76 が返り、"\\srv\share\" でしか通らない。
' 1回目が失敗したときだけ、末尾を残した形でもう一度試す。
' 「形が変わらないなら再試行しない」= 遅い共有でタイムアウトを二重に払わない。
Private Sub TestProbeRetryPath()
    ' 末尾 "\" を落とした = 形が変わった → もう1回試す価値がある。
    modTestRunner.Check "B10: UNC共有ルートは末尾\付きで再試行する", _
        (modShareRule.ProbeRetryPath("\\srv\share\") = "\\srv\share\"), _
        "実際=[" & modShareRule.ProbeRetryPath("\\srv\share\") & "]"
    modTestRunner.Check "B10: 深いUNCパスも末尾\付きで再試行する", _
        (modShareRule.ProbeRetryPath("\\srv\share\Nexus\") = "\\srv\share\Nexus\"), _
        "実際=[" & modShareRule.ProbeRetryPath("\\srv\share\Nexus\") & "]"
    modTestRunner.Check "B10: ローカルの深いパスも再試行対象", _
        (modShareRule.ProbeRetryPath("C:\data\nexus\") = "C:\data\nexus\"), ""

    ' 末尾に "\" が無い = 1回目がそのままの形 → 再試行しない。
    modTestRunner.Check "B10: 末尾\が無ければ再試行しない", _
        (LenB(modShareRule.ProbeRetryPath("\\srv\share")) = 0), _
        "実際=[" & modShareRule.ProbeRetryPath("\\srv\share") & "]"
    ' ドライブルート "C:\" は ProbeTargetPath が "\" を落とさない
    ' = 1回目と同じ実引数になるので、再試行しても意味が無い。
    modTestRunner.Check "B10: ドライブルート C:\ は再試行しない(1回目と同形)", _
        (LenB(modShareRule.ProbeRetryPath("C:\")) = 0), _
        "実際=[" & modShareRule.ProbeRetryPath("C:\") & "]"
    ' "\\host\" も同様(ホスト名だけになるので ProbeTargetPath が落とさない)。
    modTestRunner.Check "B10: \\host\ は再試行しない(1回目と同形)", _
        (LenB(modShareRule.ProbeRetryPath("\\host\")) = 0), _
        "実際=[" & modShareRule.ProbeRetryPath("\\host\") & "]"
    modTestRunner.Check "B10: 空文字は再試行しない", _
        (LenB(modShareRule.ProbeRetryPath("")) = 0), ""

    ' 再試行の実引数は「1回目と必ず違う」ことが存在意義。同じものを2回投げると
    ' 遅い共有に対して OS のタイムアウトを二重に払うだけになる。
    Dim samples As Variant
    samples = Array("\\srv\share\", "\\srv\share\Nexus\", "C:\data\nexus\", _
                    "\\srv\share", "C:\", "\\host\", "", "   ")
    Dim i As Long
    Dim badSame As Boolean
    For i = LBound(samples) To UBound(samples)
        Dim rp As String: rp = modShareRule.ProbeRetryPath(CStr(samples(i)))
        If LenB(rp) > 0 Then
            If StrComp(rp, modShareRule.ProbeTargetPath(CStr(samples(i))), vbBinaryCompare) = 0 Then
                badSame = True
            End If
        End If
    Next i
    modTestRunner.Check "B10: 再試行パスは1回目の実引数と必ず異なる", _
        (badSame = False), "同じ実引数で2回GetAttrを呼ぶ = 無駄なタイムアウト"

    ' 再試行する場合、その形は必ず末尾 "\" 付き(=元の設定値そのもの)。
    modTestRunner.Check "B10: 再試行パスは末尾\付き", _
        (Right$(modShareRule.ProbeRetryPath("\\srv\share\"), 1) = "\"), ""
End Sub

' ----------------------------------------------------------------------------
' R10-2: GsCandidatePaths の http(s) 境界(実機初報A・GS解決の堅牢化)
' ----------------------------------------------------------------------------
' ThisWorkbook.Path はOneDrive上のブックだと "https://d.docs.live.net/..." の
' ようなURLになり得る。従来はそのまま "\Ghostscript\gswin32c.exe" を連結して
' 無意味な候補を作っていた(実在確認は必ず失敗するだけの死に候補)。
' wbDir が "http" 始まり(大小無視)のときは同梱候補(stage2)を作らないことを
' 純ロジックで固定する。
Private Sub TestGsCandidatePathsHttpBoundary()
    Dim cHttps As String
    cHttps = optOcrCore.GsCandidatePaths("", "https://d.docs.live.net/xxx", "")
    modTestRunner.Check "R10-2: httpsのwbDirは同梱Ghostscript候補を含まない", _
        (InStr(1, cHttps, "Ghostscript", vbTextCompare) = 0), "実際=[" & cHttps & "]"

    Dim cHttpUpper As String
    cHttpUpper = optOcrCore.GsCandidatePaths("", "HTTP://example.com/x", "")
    modTestRunner.Check "R10-2: 大文字HTTPも同梱Ghostscript候補を含まない(大小無視)", _
        (InStr(1, cHttpUpper, "Ghostscript", vbTextCompare) = 0), "実際=[" & cHttpUpper & "]"

    Dim cNormal As String
    cNormal = optOcrCore.GsCandidatePaths("", "C:\work", "")
    modTestRunner.Check "R10-2: 通常のローカルパスは同梱Ghostscript候補を含む", _
        (InStr(1, cNormal, "Ghostscript", vbTextCompare) > 0), "実際=[" & cNormal & "]"

    ' cfgPath/searchDirsはhttp判定と無関係(候補として素通り)であることも確認する。
    Dim cCfgOnly As String
    cCfgOnly = optOcrCore.GsCandidatePaths("C:\gs\gswin32c.exe", "https://d.docs.live.net/xxx", "")
    modTestRunner.Check "R10-2: httpsのwbDirでもcfgPath候補は残る", _
        (InStr(1, cCfgOnly, "C:\gs\gswin32c.exe", vbTextCompare) > 0), "実際=[" & cCfgOnly & "]"
End Sub

' ----------------------------------------------------------------------------
' R10-3: txtwriteによるテキストPDF抽出の純ロジック
' ----------------------------------------------------------------------------
' 実機の管理端末は Word/Acrobat の COM が塞がれており、さらに使えている端末でも
' WordのPDF Reflowが「'Word'がOLE操作を完了するのを待っています」を頻発させて
' いた。そこで同梱Ghostscriptの txtwrite を PDF本文抽出の第1選択にした。
' 実GSはLibreOffice環境で動かせないので、机上で担保できるのは
'   (1) コマンド文字列の組み立て(引用符の付け忘れが最大の地雷)
'   (2) 「採用するか、画像PDFとしてOCRへ回すか」の境界
' の2つ。どちらも1文字単位で固定する。
Private Sub TestGsTextCommandGolden()
    Dim cmdText As String
    cmdText = optOcrCore.BuildGsTextCommand("C:\Program Files\gs\gswin32c.exe", _
        "C:\My Docs\約款.pdf", "C:\Temp\nxocr_x\gstext.txt")

    modTestRunner.Check "R10-3: BuildGsTextCommandのゴールデン(全パスが二重引用符)", _
        (cmdText = Chr$(34) & "C:\Program Files\gs\gswin32c.exe" & Chr$(34) & _
            " -dSAFER -dNOPAUSE -dBATCH" & _
            " -sDEVICE=txtwrite" & _
            " -sOutputFile=" & Chr$(34) & "C:\Temp\nxocr_x\gstext.txt" & Chr$(34) & _
            " " & Chr$(34) & "C:\My Docs\約款.pdf" & Chr$(34)), _
        "実際=[" & cmdText & "]"

    ' 画像化(jpeg)側と取り違えていないこと。txtwrite側に -sDEVICE=jpeg や
    ' 解像度指定が混ざると、テキストのつもりで画像を吐いて必ず空振りする。
    modTestRunner.Check "R10-3: txtwriteコマンドにjpeg/解像度指定が混ざらない", _
        (InStr(1, cmdText, "jpeg", vbTextCompare) = 0 And _
         InStr(1, cmdText, " -r", vbTextCompare) = 0), "実際=[" & cmdText & "]"

    ' 完了フラグ付きのcmd.exeラップは画像化側と共用する(BuildRunCommand)。
    ' ここが繋がっていないと監視ループがタイムアウトまで解けない。
    Dim runCmd As String
    runCmd = optOcrCore.BuildRunCommand(cmdText, "C:\Temp\nxocr_x\done.flag")
    modTestRunner.Check "R10-3: txtwriteコマンドもBuildRunCommandで完了フラグを付けられる", _
        (Left$(runCmd, 14) = "cmd.exe /s /c " And _
         InStr(1, runCmd, "echo done>" & Chr$(34) & "C:\Temp\nxocr_x\done.flag" & Chr$(34), _
               vbTextCompare) > 0), "実際=[" & runCmd & "]"
End Sub

' GsTextVerdict の境界(しきい値40字)。テキストPDFとして採用するか、
' 画像PDF(ほぼ空)としてOCR経路へ回すかの分かれ目そのもの。ここが緩いと
' 画像PDFを「文字が取れた」と誤認してOCRへ回らなくなり、逆に厳しすぎると
' 表紙だけの短いテキストPDFがOCR送りになる。
Private Sub TestGsTextVerdictBoundary()
    modTestRunner.Check "R10-3: 0字はimage(画像PDF疑い)", _
        (optOcrCore.GsTextVerdict(0) = "image"), "実際=" & optOcrCore.GsTextVerdict(0)
    modTestRunner.Check "R10-3: 39字はimage(しきい値の1つ手前)", _
        (optOcrCore.GsTextVerdict(39) = "image"), "実際=" & optOcrCore.GsTextVerdict(39)
    modTestRunner.Check "R10-3: 40字はok(しきい値ちょうどは採用)", _
        (optOcrCore.GsTextVerdict(40) = "ok"), "実際=" & optOcrCore.GsTextVerdict(40)
    modTestRunner.Check "R10-3: 41字はok", _
        (optOcrCore.GsTextVerdict(41) = "ok"), "実際=" & optOcrCore.GsTextVerdict(41)
    modTestRunner.Check "R10-3: 十分に長ければok", _
        (optOcrCore.GsTextVerdict(120000) = "ok"), "実際=" & optOcrCore.GsTextVerdict(120000)

    ' 負の値(想定外)は「採用しない」側へ倒す。数え方を間違えたときに、
    ' 画像PDFを本文ありと誤認するより、OCRへ回す方が被害が小さい。
    modTestRunner.Check "R10-3: 負の値はimage(安全側へ倒す)", _
        (optOcrCore.GsTextVerdict(-1) = "image"), "実際=" & optOcrCore.GsTextVerdict(-1)
End Sub

Public Sub RunAll6()
    On Error GoTo VocabFail
    TestExpiryDecisionVocabulary
NextSkew:
    On Error GoTo SkewFail
    TestPublishLockClockSkew
NextRetry:
    On Error GoTo RetryFail
    TestProbeRetryPath
NextGsHttp:
    On Error GoTo GsHttpFail
    TestGsCandidatePathsHttpBoundary
NextGsTextCmd:
    On Error GoTo GsTextCmdFail
    TestGsTextCommandGolden
NextGsVerdict:
    On Error GoTo GsVerdictFail
    TestGsTextVerdictBoundary
NextDone6:
    On Error GoTo 0
    Exit Sub

VocabFail:
    modTestRunner.Check "TestExpiryDecisionVocabulary(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextSkew
SkewFail:
    modTestRunner.Check "TestPublishLockClockSkew(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRetry
RetryFail:
    modTestRunner.Check "TestProbeRetryPath(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGsHttp
GsHttpFail:
    modTestRunner.Check "TestGsCandidatePathsHttpBoundary(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGsTextCmd
GsTextCmdFail:
    modTestRunner.Check "TestGsTextCommandGolden(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGsVerdict
GsVerdictFail:
    modTestRunner.Check "TestGsTextVerdictBoundary(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone6
End Sub
