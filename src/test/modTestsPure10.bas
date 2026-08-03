Attribute VB_Name = "modTestsPure10"
Option Explicit

' ============================================================================
' modTestsPure10 - R13-7c(チーム/部・共有ビーコンのteam列)の純ロジック回帰テスト
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか:
'   modTestsPure9(26,883字)は次の追加を入れると28,000字のWARN帯に触れる。
'   憲章§4-6「WARN帯のモジュールに機能を足さない。足す前に分割を裁定する」
'   に従い、新規分割先とした。入口は modTestsPure9.RunAll9 の末尾から呼ばれる
'   Public Sub RunAll10()。modTestRunner.RunAllPureTests は modTestsPure.RunAll
'   だけを呼ぶ契約なので、ここへの導線は modTestsPure9 内の1行だけ。消すと
'   このモジュールのテストは「実行されないまま」全部PASSに見える。
'
' 固定する事実:
'   ・modP2PIo.TeamCodeOf: ユーザーIDの末尾 "_"+英大文字/数字4〜6桁だけを
'     チームコードとみなす(実機で観測された「氏名_E2T22」規約)。小文字は
'     不一致、アンダースコアが複数あるときは【最後の】"_"を基準にする
'     (末尾優先で、手前の区切りにフォールバックしない)。
'   ・modP2PIo.DeptOf: チームコードの先頭3字。空入力は空を返す。
'   ・modP2PIo.BeaconDataText/BeaconTeamField: 共有ビーコン(タブ区切り1行)の
'     組み立てとteam列の取り出し。team列が無い旧形式ファイルを新しい読み手が
'     読んでも、次の列(送信時刻)をteamと誤読しないこと(R13-7cの本旨=
'     「旧読み手は新形式の余分な列を無視する/新読み手は列数で分岐する」の
'     後半を固定する)。
' ============================================================================

Private Sub TestTeamCodeOf()
    ' 正常: 「氏名_E2T22」規約どおりの末尾チームコード。
    modTestRunner.Check "TeamCodeOf_正常", _
        modP2PIo.TeamCodeOf("山田太郎_E2T22") = "E2T22"

    ' サフィックス無し("_"自体が無い)。
    modTestRunner.Check "TeamCodeOf_サフィックス無し", _
        modP2PIo.TeamCodeOf("山田太郎") = ""

    ' 空入力・末尾が"_"だけ(サフィックスが空)。
    modTestRunner.Check "TeamCodeOf_空入力", modP2PIo.TeamCodeOf("") = ""
    modTestRunner.Check "TeamCodeOf_末尾アンダースコアのみ", _
        modP2PIo.TeamCodeOf("山田太郎_") = ""

    ' 小文字は不一致(規約は英大文字/数字のみ)。
    modTestRunner.Check "TeamCodeOf_小文字は不一致", _
        modP2PIo.TeamCodeOf("山田太郎_e2t22") = ""

    ' 長さ境界: 4桁(下限)・6桁(上限)は一致、3桁・7桁は不一致。
    modTestRunner.Check "TeamCodeOf_長さ境界4は一致", _
        modP2PIo.TeamCodeOf("山田_AB12") = "AB12"
    modTestRunner.Check "TeamCodeOf_長さ境界6は一致", _
        modP2PIo.TeamCodeOf("山田_ABC123") = "ABC123"
    modTestRunner.Check "TeamCodeOf_長さ境界3は不一致", _
        modP2PIo.TeamCodeOf("山田_AB1") = ""
    modTestRunner.Check "TeamCodeOf_長さ境界7は不一致", _
        modP2PIo.TeamCodeOf("山田_ABCD123") = ""

    ' アンダースコア複数は【末尾】優先。手前の区切りにフォールバックしない
    ' ことも合わせて確認する(末尾サフィックスが規約外なら全体で不一致)。
    modTestRunner.Check "TeamCodeOf_複数アンダースコアは末尾優先", _
        modP2PIo.TeamCodeOf("山田_太郎_E2T22") = "E2T22"
    modTestRunner.Check "TeamCodeOf_末尾が規約外なら手前へフォールバックしない", _
        modP2PIo.TeamCodeOf("E2T22_ab") = ""
End Sub

Private Sub TestDeptOf()
    modTestRunner.Check "DeptOf_先頭3字", modP2PIo.DeptOf("E2T22") = "E2T"
    modTestRunner.Check "DeptOf_空入力は空", modP2PIo.DeptOf("") = ""
    modTestRunner.Check "DeptOf_3字未満はそのまま", modP2PIo.DeptOf("AB") = "AB"
End Sub

' ビーコン行の組み立て→(WriteBeaconが末尾に送信時刻を足す想定)→Split→
' team列の取り出し、を実際の書式のまま往復させる。
Private Sub TestBeaconTeamRoundTrip()
    ' team列あり(新形式)。RefreshBoardが読む主要列(f(0)・f(3)・f(5))も
    ' ズレていないことを合わせて確認する。
    Dim dataText As String
    dataText = modP2PIo.BeaconDataText("alice", 3, "20260803", 15, "202608", 45, "2026", 200, "E2T22")
    Dim recNew As String: recNew = dataText & vbTab & "2026-08-03 10:00:00"
    Dim fNew() As String: fNew = Split(recNew, vbTab)

    modTestRunner.Check "Beacon_新形式_myId", fNew(0) = "alice"
    modTestRunner.Check "Beacon_新形式_今日分", CLng(Val(fNew(3))) = 15
    modTestRunner.Check "Beacon_新形式_今月分", CLng(Val(fNew(5))) = 45
    modTestRunner.Check "Beacon_新形式_team列を取り出せる", _
        modP2PIo.BeaconTeamField(fNew) = "E2T22"

    ' team列なし(旧形式。BeaconDataTextを使わず素の8列+送信時刻を組み立てる
    ' ことで、実際に流通しうる旧バージョンのファイルを模す)。
    Dim recOld As String
    recOld = "bob" & vbTab & "1" & vbTab & "20260803" & vbTab & "10" & vbTab & _
             "202608" & vbTab & "30" & vbTab & "2026" & vbTab & "120" & vbTab & _
             "2026-08-03 09:00:00"
    Dim fOld() As String: fOld = Split(recOld, vbTab)

    modTestRunner.Check "Beacon_旧形式_列数は9(idx0-8)", UBound(fOld) = 8
    modTestRunner.Check "Beacon_旧形式_送信時刻をteamと誤読しない", _
        modP2PIo.BeaconTeamField(fOld) = ""

    ' team列ありの新形式は列数が10(idx0-9)になっていること自体も固定する
    ' (この差が新旧の分岐条件そのものなので、境界がズレたら即検知できる)。
    modTestRunner.Check "Beacon_新形式_列数は10(idx0-9)", UBound(fNew) = 9
End Sub

Public Sub RunAll10()
    On Error GoTo TeamCodeFail
    TestTeamCodeOf
NextDeptOf:
    On Error GoTo DeptOfFail
    TestDeptOf
NextBeaconRoundTrip:
    On Error GoTo BeaconRoundTripFail
    TestBeaconTeamRoundTrip
NextDone10:
    On Error GoTo 0
    Exit Sub

TeamCodeFail:
    modTestRunner.Check "TestTeamCodeOf(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDeptOf
DeptOfFail:
    modTestRunner.Check "TestDeptOf(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBeaconRoundTrip
BeaconRoundTripFail:
    modTestRunner.Check "TestBeaconTeamRoundTrip(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone10
End Sub
