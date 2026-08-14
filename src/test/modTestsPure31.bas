Attribute VB_Name = "modTestsPure31"
Option Explicit

' ============================================================================
' modTestsPure31 - R32波2(⑦PIIオフスイッチ+付随バグ)の純ロジック回帰テスト。
'   modTestsPure30はR32波1が直したばかりで触らない方針(CLAUDE.md「禁止」)
'   のため、既存チェーン(modTestsPure.RunAll→…→modTestsPure30.RunAll30)へは
'   繋がず、独立の新規モジュールとして新設した。入口は
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll31 の1本
'   (modTestsPure.RunAllの呼び出し列とは別枝)。
' ----------------------------------------------------------------------------
' 【本ハーネスの既知の制約】(modTestsPure.bas冒頭コメントに記録済みの制約と
'   同型): LibreOffice Basic純ロジック実行環境にはExcelのconfigシートが
'   存在しない。そのため modConfig.GetBool/GetLong はどのキーを渡しても
'   必ず Fallback(=呼び出し側が渡した既定値)経路を通る。ここで固定できるのは
'   「configシートが無い/読めない環境で、コード側の既定値そのものが正しいか」
'   までで、configシートに実際にTRUE/FALSEや別の数値を書いたときの分岐
'   (LookupValueが値を返す側の経路)はこの実行テストでは検証できない
'   (Excel実機での確認が必要。docs/45_実機スモークテスト手順.md参照)。
' ============================================================================

' ----------------------------------------------------------------------------
' R32 W2-1: pii_scan_enabled は新設キーで既定FALSE(既定オフ)。
'   modPackExport.ExportPackToFile の関所
'   `If modConfig.GetBool("pii_scan_enabled", False) Then` が読む既定値と
'   同じ呼び出し形(キー名・デフォルト引数とも同一)で固定する。
'   config行が無い/読めない環境では必ず安全側(=誤検知で発行を止めない側)
'   へ倒れることを保証する。
' ----------------------------------------------------------------------------
Private Sub TestPiiScanEnabledDefault32()
    modTestRunner.Check "R32-W2-1_pii_scan_enabledの既定はFALSE(configシート無しの環境)", _
        (modConfig.GetBool("pii_scan_enabled", False) = False), _
        "実際=" & modConfig.GetBool("pii_scan_enabled", False)
End Sub

' ----------------------------------------------------------------------------
' R32 W2-1: 波1が使う2キー(gap_keep_days/gap_dup_hours)を波2でconfigへ追加。
'   既定値は modInsightIo.bas(GcOldNonces手前の gap_keep_days)・
'   modInsight.bas(GapDupBlocked手前の gap_dup_hours)のGetLong呼び出しの
'   既定値と一致していることを固定する(configの既定値とコードの既定値が
'   片方だけ変わってズレる事故を検知する唯一の場所)。
' ----------------------------------------------------------------------------
Private Sub TestGapConfigDefaults32()
    modTestRunner.Check "R32-W2-1_gap_keep_daysの既定は30日(modInsightIoの既定値と一致)", _
        (modConfig.GetLong("gap_keep_days", 30) = 30), _
        "実際=" & modConfig.GetLong("gap_keep_days", 30)
    modTestRunner.Check "R32-W2-1_gap_dup_hoursの既定は24時間(modInsightの既定値と一致)", _
        (modConfig.GetLong("gap_dup_hours", 24) = 24), _
        "実際=" & modConfig.GetLong("gap_dup_hours", 24)
End Sub

' ============================================================================
Public Sub RunAll31()
    On Error GoTo Fail31
    TestPiiScanEnabledDefault32
    TestGapConfigDefaults32
NextDone31:
    On Error GoTo 0
    Exit Sub

Fail31:
    modTestRunner.Check "TestPiiScanEnabledDefault32/TestGapConfigDefaults32(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone31
End Sub
