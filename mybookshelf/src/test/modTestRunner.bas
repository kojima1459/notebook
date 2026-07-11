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
'     設計にしている(Wave1時点ではmodTestsPureがまだ存在しない)。
'     On Error Resume Next で1行だけ囲み、失敗したらそれ自体を1件の
'     テスト失敗として記録する(R5: エラー握りつぶし禁止に配慮し、
'     「無かったことにする」のではなく「失敗として可視化する」)。
'   ・tools/run_lo_tests.py はこのモジュールを modTestsPure 等と一緒に
'     LibreOffice側の一時ライブラリへ注入し、生成した TestMain.Main から
'     RunAllPureTests → ReportText の順で呼び出してファイルへ書き出す
'     (詳細は tools/README.md)。
' ============================================================================

Private mTotalCount As Long     ' これまでのCheck呼び出し総数
Private mFailCount As Long      ' 失敗数
Private mFailLines() As String  ' 失敗の詳細(1件1行)
Private mFailCap As Long        ' mFailLinesの現在容量
Private mStarted As Boolean     ' ResetTests済みかどうか(未Reset時の誤集計防止)

' ----------------------------------------------------------------------------
' ResetTests: 集計をゼロから始める
' ----------------------------------------------------------------------------
Public Sub ResetTests()
    mTotalCount = 0
    mFailCount = 0
    mFailCap = 16
    ReDim mFailLines(1 To mFailCap)
    mStarted = True
End Sub

' ----------------------------------------------------------------------------
' Check: 1件のテスト結果を記録する
'   name   : テスト名(例: "Fnv1a64Hex_決定性")
'   cond   : True=成功 / False=失敗
'   detail : 失敗時に添える補足(期待値/実際値など)。省略可
' ----------------------------------------------------------------------------
Public Sub Check(ByVal name As String, ByVal cond As Boolean, Optional ByVal detail As String = "")
    If Not mStarted Then ResetTests

    mTotalCount = mTotalCount + 1

    If Not cond Then
        mFailCount = mFailCount + 1
        If mFailCount > mFailCap Then
            mFailCap = mFailCap * 2
            ReDim Preserve mFailLines(1 To mFailCap)
        End If

        Dim line As String
        line = "NG: " & name
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
' ReportText: "PASS n / FAIL m" に続けて失敗一覧(1行1件)を返す
'   例:
'     PASS 41 / FAIL 1
'     NG: Fnv1a64Hex_決定性 -- 同じ文字列で異なるhexが出た
' ----------------------------------------------------------------------------
Public Function ReportText() As String
    If Not mStarted Then ResetTests

    Dim passCount As Long
    passCount = mTotalCount - mFailCount

    Dim parts() As String
    ReDim parts(0 To mFailCount)
    parts(0) = "PASS " & passCount & " / FAIL " & mFailCount

    Dim i As Long
    For i = 1 To mFailCount
        parts(i) = mFailLines(i)
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
End Sub
