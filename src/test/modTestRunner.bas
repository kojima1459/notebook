Attribute VB_Name = "modTestRunner"
Option Explicit

' ============================================================================
' modTestRunner - 純ロジックテストの実行エンジン(MASTER_SPEC §7.8)
' ----------------------------------------------------------------------------
' 役割:
'   modTestsPure(および将来増える各 modTests*)が呼ぶ「Check」を受け止め、
'   合否を集計してレポート文字列を作る、テストの土台。
'   ExcelでもLibreOffice(ヘッドレス)でも同じコードで動くことが最重要な
'   モジュールなので、他のどのモジュールより厳格にR4(純ロジック)を守る。
'
' 設計判断:
'   ・R4準拠: Worksheets/Range/Application/ThisWorkbook/MsgBoxに一切触れない。
'     結果はすべてPrivateなモジュール変数(配列)に蓄積し、ReportTextが
'     文字列として返す。呼び出し側(Excelなら診断シート、LOならファイル
'     書き出し)が好きな方法で表示すればよい。
'   ・改行は vbLf で統一(§12: 純ロジックモジュールはvbLf基準。LOの
'     StarBasicはvbCrLfも解釈できるが、テキストファイル比較やLO側での
'     表示ズレを避けるため vbLf に統一する)。
'   ・失敗一覧は「あとから追記」しやすいよう固定長配列を倍々に伸長する方式
'     (ReDim Preserve)にした。Collectionでも書けるが、配列の方が
'     LibreOffice Basicでの挙動の揺れが少なく、要素数もFailures()で
'     即座に読めるため採用した。
'   ・RunAllPureTestsは「まだ実装されていないmodTests*」を落とさず許容する
'     設計にしている(将来modTests*が増えても、個別モジュールの有無で
'     全体が落ちないようにするための一般的な安全策)。
'     On Error Resume Next で1行だけ囲み、失敗したらそれ自体を1件の
'     テスト失敗として記録する(R5: エラー握りつぶし禁止に配慮し、
'     「無かったことにする」のではなく「失敗として可視化する」)。
'   ・tools/run_lo_tests.py はこのモジュールを modTestsPure 等と一緒に
'     LibreOffice側の一時ライブラリへ注入し、生成した TestMain.Main から
'     RunAllPureTests → ReportText の順で呼び出してファイルへ書き出す
'     (詳細は tools/README.md)。
' ============================================================================

' ----------------------------------------------------------------------------
' SKIP集計(2026-08-15 R33波1 W1-1)
'   実行環境の制限で本体を1行も実行できないテスト群が、これまで
'   `Check "…スキップ", True` として【PASSに計上】されていた。レポートには
'   「実行されなかった」という情報が一切残らず、PASS件数が水増しされていた
'   (実測: LibreOffice では CanUseTypeArrays()=False が確定しており、
'    modChunker/modPrompts などの群が丸ごと未実行のままPASSになっていた)。
'   本モジュールは §7 契約が closed=True で、Public を1本も増やせないため、
'   新しいエントリポイントを足さずに Check の【テスト名の接頭辞】で振り分ける。
'   呼び出し側は  Check "[SKIP] modChunker: …", True, "理由"  と書く。
'   接頭辞つき かつ cond=True のものは PASS にも総数にも数えず、SKIP として
'   別に数えて ReportText の1行目と末尾の一覧に出す。
' ----------------------------------------------------------------------------
Private Const SKIP_PREFIX As String = "[SKIP]"

Private mTotalCount As Long     ' これまでのCheck呼び出し総数(SKIPは含めない)
Private mFailCount As Long      ' 失敗数
Private mFailLines() As String  ' 失敗の詳細(1件1行)
Private mFailCap As Long        ' mFailLinesの現在容量
Private mSkipCount As Long      ' 未実行(SKIP)数
Private mSkipLines() As String  ' 未実行の一覧(1件1行)
Private mSkipCap As Long        ' mSkipLinesの現在容量
Private mStarted As Boolean     ' ResetTests済みかどうか(未Reset時の誤集計防止)

' ----------------------------------------------------------------------------
' ResetTests: 集計をゼロから始める
' ----------------------------------------------------------------------------
Public Sub ResetTests()
    mTotalCount = 0
    mFailCount = 0
    mFailCap = 16
    ReDim mFailLines(1 To mFailCap)
    mSkipCount = 0
    mSkipCap = 16
    ReDim mSkipLines(1 To mSkipCap)
    mStarted = True
End Sub

' ----------------------------------------------------------------------------
' Check: 1件のテスト結果を記録する
'   testName : テスト名(例: "Fnv1a64Hex_決定性")
'   cond     : True=成功 / False=失敗
'   detail   : 失敗時に添える補足(期待値/実際値など)。省略可
' ----------------------------------------------------------------------------
Public Sub Check(ByVal testName As String, ByVal cond As Boolean, Optional ByVal detail As String = "")
    If Not mStarted Then ResetTests

    ' 未実行の申告([SKIP]接頭辞)は PASS にも総数にも数えない。
    ' cond=False で来た場合だけは本物の失敗として下へ流す(接頭辞を付けた
    ' まま落ちるテストを黙って隠さないため)。
    If cond Then
        If Left(testName, Len(SKIP_PREFIX)) = SKIP_PREFIX Then
            mSkipCount = mSkipCount + 1
            If mSkipCount > mSkipCap Then
                mSkipCap = mSkipCap * 2
                ReDim Preserve mSkipLines(1 To mSkipCap)
            End If
            mSkipLines(mSkipCount) = "SKIP: " & Trim(Mid(testName, Len(SKIP_PREFIX) + 1))
            Exit Sub
        End If
    End If

    mTotalCount = mTotalCount + 1

    If Not cond Then
        mFailCount = mFailCount + 1
        If mFailCount > mFailCap Then
            mFailCap = mFailCap * 2
            ReDim Preserve mFailLines(1 To mFailCap)
        End If

        Dim line As String
        line = "NG: " & testName
        If Len(detail) > 0 Then line = line & " -- " & detail
        mFailLines(mFailCount) = line
    End If
End Sub

' ----------------------------------------------------------------------------
' Failures: 現在までの失敗件数
' ----------------------------------------------------------------------------
Public Function Failures() As Long
    Failures = mFailCount
End Function

' ----------------------------------------------------------------------------
' ReportText: "PASS n / FAIL m / SKIP k" に続けて失敗一覧・未実行一覧を返す
'   例:
'     PASS 41 / FAIL 1 / SKIP 12
'     NG: Fnv1a64Hex_決定性 -- 同じ文字列で異なるhexが出た
'     SKIP: modChunker: LO環境の既知の制限により未実行
'   SKIP は「テストが1行も走っていない」という事実の可視化であり、合否では
'   ない。k>0 のときに何が守られていないかを司令塔が毎回見るための行。
' ----------------------------------------------------------------------------
Public Function ReportText() As String
    If Not mStarted Then ResetTests

    Dim passCount As Long
    passCount = mTotalCount - mFailCount

    Dim parts() As String
    ReDim parts(0 To mFailCount + mSkipCount)
    parts(0) = "PASS " & passCount & " / FAIL " & mFailCount & " / SKIP " & mSkipCount

    Dim i As Long
    For i = 1 To mFailCount
        parts(i) = mFailLines(i)
    Next i
    For i = 1 To mSkipCount
        parts(mFailCount + i) = mSkipLines(i)
    Next i

    ReportText = Join(parts, vbLf)
End Function

' ----------------------------------------------------------------------------
' RunAllPureTests: 各 modTests* の Run系(現状は modTestsPure.RunAll)を
'   列挙呼び出しする。呼び出し先が未実装/実行時エラーでも本Subは落ちない。
' ----------------------------------------------------------------------------
Public Sub RunAllPureTests()
    ResetTests

    On Error Resume Next
    Err.Clear
    modTestsPure.RunAll
    If Err.Number <> 0 Then
        Check "modTestsPure.RunAll", False, "呼び出しでエラー: " & Err.Description & _
              " (Err=" & Err.Number & ") ※未実装/未注入の可能性"
        Err.Clear
    End If
    On Error GoTo 0

    ' 2026-08-14(R32波2 W2-1): modTestsPure31は、modTestsPure.RunAllが辿る
    ' 既存チェーン(→modTestsPure2→…→modTestsPure30)には繋がない別枝として
    ' ここから直接呼ぶ。modTestsPure30はR32波1が直したばかりで触らない方針
    ' (CLAUDE.md「禁止」)のため、末尾へ1行足す形の分割ができなかった。
    On Error Resume Next
    Err.Clear
    modTestsPure31.RunAll31
    If Err.Number <> 0 Then
        Check "modTestsPure31.RunAll31", False, "呼び出しでエラー: " & Err.Description & _
              " (Err=" & Err.Number & ") ※未実装/未注入の可能性"
        Err.Clear
    End If
    On Error GoTo 0

    ' 2026-08-14(R32波4 W4-5): modTestsPure32 も同じ理由(modTestsPure30/31は
    ' 波1・波2が直したばかりで触らない方針)で、既存チェーンへは繋がず
    ' ここから直接呼ぶ別枝にする。固定するのは背景画像方式の純関数
    ' (BMPバイト列の生成と、敷き直しを弾く冪等メモ)。
    On Error Resume Next
    Err.Clear
    modTestsPure32.RunAll32
    If Err.Number <> 0 Then
        Check "modTestsPure32.RunAll32", False, "呼び出しでエラー: " & Err.Description & _
              " (Err=" & Err.Number & ") ※未実装/未注入の可能性"
        Err.Clear
    End If
    On Error GoTo 0

    ' 2026-08-16(R33波2): modTestsPure33 も同じ別枝。既存チェーンの末尾
    ' (modTestsPure30)はR33波1が直したばかりで触らない方針。固定するのは
    ' PII検知の全角ゴールデン(W2-1)と、失効判定の「つながっていれば消さない」
    ' 不変条件(W2-2)の2群。
    On Error Resume Next
    Err.Clear
    modTestsPure33.RunAll33
    If Err.Number <> 0 Then
        Check "modTestsPure33.RunAll33", False, "呼び出しでエラー: " & Err.Description & _
              " (Err=" & Err.Number & ") ※未実装/未注入の可能性"
        Err.Clear
    End If
    On Error GoTo 0

    ' 2026-08-16(R33波3): modTestsPure33 が残997字になったための分割先。
    ' 同じ別枝の作法で直接呼ぶ。固定するのは opt機能の "#ERR:" の理由を
    ' UIが握り潰さないこと(W3-9)。
    On Error Resume Next
    Err.Clear
    modTestsPure34.RunAll34
    If Err.Number <> 0 Then
        Check "modTestsPure34.RunAll34", False, "呼び出しでエラー: " & Err.Description & _
              " (Err=" & Err.Number & ") ※未実装/未注入の可能性"
        Err.Clear
    End If
    On Error GoTo 0

    ' 2026-08-16(R33波5c): modTestsPure34 が残1,676字になったための分割先。
    ' 同じ別枝の作法で直接呼ぶ。固定するのは「部門ごとの削除」が消してよい
    ' ものだけを数えること(W5-23。self / pack: を巻き込まない線引き)。
    On Error Resume Next
    Err.Clear
    modTestsPure35.RunAll35
    If Err.Number <> 0 Then
        Check "modTestsPure35.RunAll35", False, "呼び出しでエラー: " & Err.Description & _
              " (Err=" & Err.Number & ") ※未実装/未注入の可能性"
        Err.Clear
    End If
    On Error GoTo 0
End Sub
