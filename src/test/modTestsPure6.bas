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
            " -sOutputFile=" & Chr$(34) & "C:\Temp\nxocr_x\gstext_%04d.txt" & Chr$(34) & _
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
    runCmd = optOcrCore.BuildRunCommand(cmdText, "C:\Temp\nxocr_x\done.flag", _
                                        "C:\Temp\nxocr_x\gs_out.log")
    modTestRunner.Check "R10-3: txtwriteコマンドもBuildRunCommandで完了フラグを付けられる", _
        (Left$(runCmd, 14) = "cmd.exe /s /c " And _
         InStr(1, runCmd, " >" & Chr$(34) & "C:\Temp\nxocr_x\done.flag" & Chr$(34), _
               vbTextCompare) > 0), "実際=[" & runCmd & "]"
End Sub

' ----------------------------------------------------------------------------
' R11-D(監査3 H-2): GS実行の観測性
' ----------------------------------------------------------------------------
' 実機(管理端末)では MOTW / AppLocker / EDR に gswin32c.exe の実行を
' 止められることがあるが、旧実装は標準出力も標準エラーも捨てていたため
' 「フラグは出たのに1枚も画像ができていない」としか観測できなかった。
' 新仕様: (1) GS本体を丸括弧でまとめて gs_out.log へ 1> と 2>&1 で落とす
'         (2) 完了フラグの中身を "done" 固定から【終了コード】へ変える
'         (3) フラグ作成は従来どおり `&`(無条件)連結=必ず出る
' ここでは「文字列の組み立て」と「フラグの読み方」だけを固定する。
Private Sub TestGsRunCommandObservability()
    Dim gsCmd As String
    gsCmd = Chr$(34) & "C:\Program Files\gs\gswin32c.exe" & Chr$(34) & " -dSAFER " & _
            Chr$(34) & "C:\My Docs\約款.pdf" & Chr$(34)
    Dim runCmd As String
    runCmd = optOcrCore.BuildRunCommand(gsCmd, "C:\Temp\nxocr_x\done.flag", _
                                        "C:\Temp\nxocr_x\gs_out.log")

    ' (1) 標準出力・標準エラーの両方がログへ落ちること。
    modTestRunner.Check "R11-D: stdoutをgs_out.logへ落とす", _
        (InStr(runCmd, " 1>" & Chr$(34) & "C:\Temp\nxocr_x\gs_out.log" & Chr$(34)) > 0), _
        "実際=[" & runCmd & "]"
    modTestRunner.Check "R11-D: stderrをstdoutへ合流させる(2>&1)", _
        (InStr(runCmd, " 2>&1") > 0), "実際=[" & runCmd & "]"

    ' (2) GS本体は丸括弧でまとめる(まとめないとリダイレクトが最終トークン
    '     にしか掛からず、PDFのパスがログ名として解釈されうる)。
    modTestRunner.Check "R11-D: GS本体が丸括弧でまとめられている", _
        (InStr(runCmd, "/c " & Chr$(34) & "(" & gsCmd & ")") > 0), "実際=[" & runCmd & "]"

    ' (3) フラグへ書くのは終了コード。`call` が無いと初回解析で親プロセスの
    '     0 に展開されてしまい、常に「成功」に見える。
    modTestRunner.Check "R11-D: 終了コードをcall経由でフラグへ書く", _
        (InStr(runCmd, " & (if errorlevel 1 (echo 1) else if errorlevel 0 (echo 0) else (echo 255)) >") > 0), "実際=[" & runCmd & "]"

    ' (4) `>` の直前に必ず空白がある。空白が無いと echo 1>… の 1 が
    '     リダイレクト先ハンドル番号として食われ、フラグが空になる。
    modTestRunner.Check "R11-D: フラグへのリダイレクト直前に空白がある", _
        (InStr(runCmd, ") >" & Chr$(34)) > 0 And InStr(runCmd, "echo 1)") > 0), "実際=[" & runCmd & "]"

    ' (5) 連結は `&`(無条件)のまま=GSが異常終了してもフラグは必ず出る。
    modTestRunner.Check "R11-D: フラグ作成の連結は無条件の&のまま(&&にしない)", _
        (InStr(runCmd, " && ") = 0 And InStr(runCmd, " & (if errorlevel") > 0), _
        "実際=[" & runCmd & "]"

    ' (6) ログのパス規約。
    modTestRunner.Check "R11-D: GsLogForはフォルダ直下のgs_out.log", _
        (optOcrCore.GsLogFor("C:\Temp\nxocr_x") = "C:\Temp\nxocr_x\gs_out.log"), _
        "実際=" & optOcrCore.GsLogFor("C:\Temp\nxocr_x")
    modTestRunner.Check "R11-D: GsLogForは末尾の区切り文字を正規化する", _
        (optOcrCore.GsLogFor("C:\Temp\nxocr_x\") = "C:\Temp\nxocr_x\gs_out.log"), _
        "実際=" & optOcrCore.GsLogFor("C:\Temp\nxocr_x\")
End Sub

' 完了フラグの中身から終了コードを読む。判定不能(-1)は「失敗と決めつけない」
' 側へ倒す約束なので、そこを含めて境界を固定する。
Private Sub TestGsExitCodeFromFlag()
    modTestRunner.Check "R11-D: 正常終了は0", _
        (optOcrCore.GsExitCodeFromFlag("0 " & vbCrLf) = 0), _
        "実際=" & optOcrCore.GsExitCodeFromFlag("0 " & vbCrLf)
    modTestRunner.Check "R11-D: 失敗コードはそのまま読む", _
        (optOcrCore.GsExitCodeFromFlag("1 " & vbCrLf) = 1), _
        "実際=" & optOcrCore.GsExitCodeFromFlag("1 " & vbCrLf)
    modTestRunner.Check "R11-D: 空白や改行だけが混ざっても読める", _
        (optOcrCore.GsExitCodeFromFlag(vbCrLf & "  9009  " & vbCrLf) = 9009), _
        "実際=" & optOcrCore.GsExitCodeFromFlag(vbCrLf & "  9009  " & vbCrLf)
    ' 旧形式("done")や、%^ERRORLEVEL% が展開できなかった端末の残骸は
    ' 【判定不能=-1】。ここを0や失敗へ倒すと、取込を壊すか嘘の成功になる。
    modTestRunner.Check "R11-D: 旧形式doneは判定不能(-1)", _
        (optOcrCore.GsExitCodeFromFlag("done") = -1), _
        "実際=" & optOcrCore.GsExitCodeFromFlag("done")
    modTestRunner.Check "R11-D: 展開されなかった変数名は判定不能(-1)", _
        (optOcrCore.GsExitCodeFromFlag("%^ERRORLEVEL% ") = -1), _
        "実際=" & optOcrCore.GsExitCodeFromFlag("%^ERRORLEVEL% ")
    modTestRunner.Check "R11-D: 空のフラグは判定不能(-1)", _
        (optOcrCore.GsExitCodeFromFlag("") = -1), ""
    ' クラッシュ系の負の終了コードは「非0の失敗」として255へ丸める。
    modTestRunner.Check "R11-D: 負の終了コードは失敗(255)へ丸める", _
        (optOcrCore.GsExitCodeFromFlag("-1073741819 ") = 255), _
        "実際=" & optOcrCore.GsExitCodeFromFlag("-1073741819 ")
    modTestRunner.Check "R11-D: 桁あふれする長い数字は判定不能(-1)", _
        (optOcrCore.GsExitCodeFromFlag("12345678901") = -1), _
        "実際=" & optOcrCore.GsExitCodeFromFlag("12345678901")
End Sub

' ----------------------------------------------------------------------------
' R10c(M1/M2): GsTextVerdict の3値化と、ページ数に比例するしきい値
' ----------------------------------------------------------------------------
' R10-3の実装は固定40字の2値判定だった。これだと「100ページのスキャンPDFに
' 透明テキストが50字だけ乗っている」資料を "ok" と判定して取り込んでしまい、
' 中身の無い本が本棚に並ぶ(しかもOCRへ回らないので二度と読めない)。
' 逆に一律で厳しくすると、表紙1枚だけの正当な短いPDFがOCR送りになる。
' そこで3値にした:
'   image  文字層ゼロ  → 呼び出し元は即E0303(OCR経路)
'   sparse 薄すぎる    → 呼び出し元は Word/Acrobat を先に試し、両方失敗時のみE0303
'   ok     採用
' しきい値は Max(40, pageCount×10)。境界の上下両側を固定する。
Private Sub TestGsTextVerdictBoundary()
    ' --- 文字層ゼロは常に image(ページ数によらない) ---
    modTestRunner.Check "R10c: 0字はimage(1ページ)", _
        (optOcrCore.GsTextVerdict(0, 1) = "image"), "実際=" & optOcrCore.GsTextVerdict(0, 1)
    modTestRunner.Check "R10c: 0字はimage(100ページ)", _
        (optOcrCore.GsTextVerdict(0, 100) = "image"), "実際=" & optOcrCore.GsTextVerdict(0, 100)
    ' 負の値(想定外)も image 側へ倒す。数え方を誤ったときに本文ありと
    ' 誤認するより、OCRへ回す方が被害が小さい。
    modTestRunner.Check "R10c: 負の値はimage(安全側へ倒す)", _
        (optOcrCore.GsTextVerdict(-1, 1) = "image"), "実際=" & optOcrCore.GsTextVerdict(-1, 1)

    ' --- 少ページ側は固定40字が効く(下限40) ---
    modTestRunner.Check "R10c: 1字はsparse(1ページ・40未満)", _
        (optOcrCore.GsTextVerdict(1, 1) = "sparse"), "実際=" & optOcrCore.GsTextVerdict(1, 1)
    modTestRunner.Check "R10c: 39字はsparse(しきい値の1つ手前)", _
        (optOcrCore.GsTextVerdict(39, 1) = "sparse"), "実際=" & optOcrCore.GsTextVerdict(39, 1)
    modTestRunner.Check "R10c: 40字はok(しきい値ちょうどは採用)", _
        (optOcrCore.GsTextVerdict(40, 1) = "ok"), "実際=" & optOcrCore.GsTextVerdict(40, 1)
    modTestRunner.Check "R10c: 40字はok(4ページでも40が下限)", _
        (optOcrCore.GsTextVerdict(40, 4) = "ok"), "実際=" & optOcrCore.GsTextVerdict(40, 4)

    ' --- 多ページ側はページ比例(pageCount×10)が効く ---
    modTestRunner.Check "R10c: 100ページ×50字はsparse(透明テキスト付きスキャンPDF)", _
        (optOcrCore.GsTextVerdict(50, 100) = "sparse"), "実際=" & optOcrCore.GsTextVerdict(50, 100)
    modTestRunner.Check "R10c: 100ページ×999字はsparse(比例しきい値1000の直下)", _
        (optOcrCore.GsTextVerdict(999, 100) = "sparse"), "実際=" & optOcrCore.GsTextVerdict(999, 100)
    modTestRunner.Check "R10c: 100ページ×1000字はok(比例しきい値ちょうど)", _
        (optOcrCore.GsTextVerdict(1000, 100) = "ok"), "実際=" & optOcrCore.GsTextVerdict(1000, 100)
    modTestRunner.Check "R10c: 十分に長ければok", _
        (optOcrCore.GsTextVerdict(120000, 100) = "ok"), "実際=" & optOcrCore.GsTextVerdict(120000, 100)

    ' pageCountが0(数えられなかった)ときは固定40字だけで判定する
    ' (比例分を0にして、しきい値が消えたり負になったりしないこと)。
    modTestRunner.Check "R10c: pageCount=0なら40字でok", _
        (optOcrCore.GsTextVerdict(40, 0) = "ok"), "実際=" & optOcrCore.GsTextVerdict(40, 0)
    modTestRunner.Check "R10c: pageCount=0でも39字はsparse", _
        (optOcrCore.GsTextVerdict(39, 0) = "sparse"), "実際=" & optOcrCore.GsTextVerdict(39, 0)
End Sub

' ----------------------------------------------------------------------------
' R10c(L1): CleanTextLen / GsPageBounds / BuildPagesFromGsText の境界
' ----------------------------------------------------------------------------
' txtwriteの出力はページ区切りが改ページ文字(Chr(12))。数え方・割り方がズレると
' (a) 採否しきい値の分母が狂う (b) 出典ページ番号が丸ごとずれる。特に(b)は
' 「出典 p.5」を開いても別のページが出る形で、製品全体の信用を失う。
' 2026-07-31(R11-F2): 添字計算を modUtilText.GsPageBounds へ一本化したので、
' 突き合わせではなく唯一の実装の単体テストとして境界を固定する(委譲側の
' optOcrCore.GsPageCount が同値を返すことも併せて見る)。
'
' ただし BuildPagesFromGsText は ExtractedPage() を ReDim するため、LibreOffice
' では既知の制限(実行時エラー420。CanUseTypeArrays参照)に当たる。当たる環境では
' そこだけスキップし、純文字列側(GsPageBounds)で境界を固定する。
Private Sub TestGsTextPageSplit()
    Dim ff As String: ff = Chr$(12)

    ' --- CleanTextLen: 空白類は数えない ---
    modTestRunner.Check "R10c: CleanTextLenは改ページ・改行・タブ・空白を数えない", _
        (optOcrCore.CleanTextLen(ff & vbCrLf & vbTab & " " & ChrW(&H3000)) = 0), _
        "実際=" & optOcrCore.CleanTextLen(ff & vbCrLf & vbTab & " " & ChrW(&H3000))
    modTestRunner.Check "R10c: CleanTextLenは本文だけを数える", _
        (optOcrCore.CleanTextLen("あい" & ff & " う" & vbCrLf & "え") = 4), _
        "実際=" & optOcrCore.CleanTextLen("あい" & ff & " う" & vbCrLf & "え")
    modTestRunner.Check "R10c: CleanTextLenの空文字は0", _
        (optOcrCore.CleanTextLen("") = 0), "実際=" & optOcrCore.CleanTextLen("")

    ' --- ページ分割の境界(先頭FF/末尾FF/FFのみ/改行だけのページ/FF無し) ---
    CheckPageSplit "改ページ無しの単一ページ", "本文だけ", 1
    CheckPageSplit "末尾に改ページ(GSの通常出力)", "1ページ目" & ff, 1
    CheckPageSplit "先頭に改ページ(本文前に1枚吐く形)", ff & "本文", 1
    CheckPageSplit "先頭と末尾の両方に改ページ", ff & "本文" & ff, 1
    CheckPageSplit "2ページ", "1枚目" & ff & "2枚目", 2
    CheckPageSplit "先頭空+2ページ+末尾空", ff & "1枚目" & ff & "2枚目" & ff, 2
    ' 中間の空ページは落とさない(本当に白紙のページがあり得るので、
    ' 落とすと以降のページ番号が全部ずれる)。
    CheckPageSplit "中間の空ページは残す", "1枚目" & ff & ff & "3枚目", 3
    ' 改行だけのページ(実質白紙)も先頭・末尾なら落とす(番号側はCheckPageBuild)。
    CheckPageSplit "改行だけの先頭ページは空扱い", vbCrLf & ff & "本文", 1

    ' --- 中身が無い入力は0ページ ---
    CheckPageSplit "改ページだけの入力", ff & ff, 0
    CheckPageSplit "空文字", "", 0
    CheckPageSplit "空白だけの入力", " " & vbCrLf, 0

    If Not CanUseTypeArrays() Then
        modTestRunner.Check "[SKIP] R10c: BuildPagesFromGsTextはLO環境の既知の制限によりスキップ", True, _
            "ExtractedPage()のReDimがLibreOfficeで実行時エラー420になる" & _
            "(modTestsPure冒頭コメント/CanUseTypeArrays参照)。ページ分割の規約は" & _
            "同一規約の純文字列版 optOcrCore.GsPageCount で固定済み。" & _
            "Excel実機受入チェック(§11.3)で出典ページ番号を必ず目視確認すること。"
        Exit Sub
    End If

    CheckPageBuild "改ページ無しの単一ページ", "本文だけ", 1, "本文だけ"
    CheckPageBuild "先頭に改ページ(本文前に1枚吐く形)", ff & "本文", 1, "本文"
    CheckPageBuild "先頭空+2ページ+末尾空", ff & "1枚目" & ff & "2枚目" & ff, 2, "1枚目"
    CheckPageBuild "中間の空ページは残す", "1枚目" & ff & ff & "3枚目", 3, "1枚目"

    Dim emptyPages() As ExtractedPage
    Dim emptyTrunc As Boolean
    modTestRunner.Check "R10c: 改ページだけの入力はBuildPagesFromGsText=False", _
        (modExtractor.BuildPagesFromGsText(ff & ff, 20, emptyPages, emptyTrunc) = False), ""
    modTestRunner.Check "R10c: 空文字はBuildPagesFromGsText=False", _
        (modExtractor.BuildPagesFromGsText("", 20, emptyPages, emptyTrunc) = False), ""

    ' --- 上限ページの打ち切り ---
    Dim capPages() As ExtractedPage
    Dim capTrunc As Boolean
    Dim okCap As Boolean
    okCap = modExtractor.BuildPagesFromGsText("A" & ff & "B" & ff & "C", 2, capPages, capTrunc)
    modTestRunner.Check "R10c: 上限2ページで打ち切りtruncated=True", _
        (okCap And capTrunc And (UBound(capPages) - LBound(capPages) + 1) = 2), _
        "ok=" & okCap & " trunc=" & capTrunc
    If okCap Then
        modTestRunner.Check "R10c: 打ち切っても先頭ページは1ページ目のまま", _
            (capPages(0).page = 1 And capPages(0).Text = "A"), _
            "page=" & capPages(0).page & " text=[" & capPages(0).Text & "]"
    End If
End Sub

' ページ分割の唯一の実装(modUtilText.GsPageBounds)の単体テスト+委譲確認。
Private Sub CheckPageSplit(ByVal label As String, ByVal txt As String, ByVal wantN As Long)
    Dim firstIdx As Long, lastIdx As Long
    Dim gotN As Long: gotN = modUtilText.GsPageBounds(txt, firstIdx, lastIdx)
    modTestRunner.Check "R11-F2: GsPageBounds " & label, (gotN = wantN), _
        "期待=" & wantN & " 実際=" & gotN

    If gotN > 0 Then
        modTestRunner.Check "R11-F2: GsPageBounds " & label & "(添字幅=件数)", _
            ((lastIdx - firstIdx + 1) = wantN), "first=" & firstIdx & " last=" & lastIdx
    End If
    modTestRunner.Check "R11-F2: GsPageCount(optOcrCore)は委譲先と同値 " & label, _
        (optOcrCore.GsPageCount(txt) = gotN), "GsPageCount=" & optOcrCore.GsPageCount(txt)
End Sub

' コア側(modExtractor.BuildPagesFromGsText)が同じページ数・同じ1ページ目に
' なることを確かめる(2箇所に分かれた実装のズレ=出典ページずれの検出)。
Private Sub CheckPageBuild(ByVal label As String, ByVal txt As String, _
                           ByVal wantN As Long, ByVal wantFirst As String)
    Dim pages() As ExtractedPage
    Dim trunc As Boolean
    Dim okBuild As Boolean
    Dim buildN As Long

    okBuild = modExtractor.BuildPagesFromGsText(txt, 100, pages, trunc)
    buildN = 0
    If okBuild Then buildN = UBound(pages) - LBound(pages) + 1

    Dim firstIdx As Long, lastIdx As Long
    modTestRunner.Check "R10c: BuildPagesFromGsText " & label & "(件数が共通部品と一致)", _
        (okBuild And buildN = modUtilText.GsPageBounds(txt, firstIdx, lastIdx) And buildN = wantN), _
        "期待=" & wantN & " 実際=" & buildN

    If okBuild Then
        ' R12-3-6: 先頭の空ページも番号を消費する=先頭要素は firstIdx+1 ページ目。
        modTestRunner.Check "R12-3-6: BuildPages " & label & "(先頭の中身と物理番号)", _
            (Trim$(pages(0).Text) = wantFirst And pages(0).page = firstIdx + 1), _
            "実際=[" & Trim$(pages(0).Text) & "] page=" & pages(0).page & " first=" & firstIdx
    End If
End Sub

' ----------------------------------------------------------------------------
' CanUseTypeArrays - modTypes の Public Type 配列を ReDim できる実行環境か
'   (modTestsPure の同名関数と同じ実測プローブ。LibreOfficeでは420になる)。
' ----------------------------------------------------------------------------
Private Function CanUseTypeArrays() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As ExtractedPage
    ReDim probe(0 To 0)
    CanUseTypeArrays = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

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
NextGsRunObs:
    On Error GoTo GsRunObsFail
    TestGsRunCommandObservability
NextGsExitCode:
    On Error GoTo GsExitCodeFail
    TestGsExitCodeFromFlag
NextGsVerdict:
    On Error GoTo GsVerdictFail
    TestGsTextVerdictBoundary
NextGsSplit:
    On Error GoTo GsSplitFail
    TestGsTextPageSplit
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
    Resume NextGsRunObs
GsRunObsFail:
    modTestRunner.Check "TestGsRunCommandObservability(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGsExitCode
GsExitCodeFail:
    modTestRunner.Check "TestGsExitCodeFromFlag(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGsVerdict
GsVerdictFail:
    modTestRunner.Check "TestGsTextVerdictBoundary(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGsSplit
GsSplitFail:
    modTestRunner.Check "TestGsTextPageSplit(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone6
End Sub
