Attribute VB_Name = "modNoiseReport"
Option Explicit

' ============================================================================
' modNoiseReport - ⚠️ノイズ報告(品質が低いと報告)の一連の流れと、その取り消し。
' ----------------------------------------------------------------------------
' 2026-09-19(R49 監査 A-A-3): 新設。理由は2つある。
'
' (1) 【誤操作が既定の動作になっていた】ナレッジ詳細のダイアログは
'     vbYesNoCancel + vbDefaultButton2 で出ていた。vbDefaultButton2 は
'     [いいえ]、つまり【ノイズ報告】。資料カードを開いて Enter や Space を
'     押しただけで、
'       ・その資料が自分の検索から即座に消える(個人ミュート)
'       ・共有フォルダへ「この資料は品質が低い」という票が1つ飛ぶ
'     という2つが起きる。全社12,000名が使う前提では誤操作は事故ではなく
'     確定事象で、しかも2つ目は他人の検索結果を変える(憲章§2)。
'     既定は【キャンセル=閉じる】へ。Enter が何も壊さない状態にする。
'
' (2) 【取り消す道が1本も無かった】報告したあと、本人が戻す導線は製品の
'     どこにも無かった。復帰できるのは config admin_users に載った管理者が
'     modP2P.ClearNoise を呼ぶ経路だけで、これは【その資料への全員の票】を
'     消すもの ―― 1人の誤操作を直すのに管理者が組織全体の判断を巻き戻す、
'     という釣り合わない操作しか用意されていなかった。
'     報告の直後に「取り消す」を出し、自分の票だけを引っ込められるようにする。
'
' 置き場所: modVaultGallery は残271字(§12 逼迫=分割裁定必須)で、この流れを
'   そのまま持てない。呼び出し元には1行だけ残し、実体をこちらへ出した。
' ============================================================================

' ----------------------------------------------------------------------------
' ReportWithUndo - ノイズ報告を実行し、その場で取り消せる確認を出す。
'   呼び出し元(modVaultGallery.OnVaultCardClick)は結果を見ない。画面の
'   再描画は呼び出し元が続けて行う(取り消したときも、報告したときと同じく
'   描き直せばよい ―― 除外集合を読み直すだけなので冪等)。
' ----------------------------------------------------------------------------
Public Sub ReportWithUndo(ByVal srcName As String)
    If LenB(srcName) = 0 Then Exit Sub

    modStats.ReportNoise srcName          ' 個人ミュート(即時・自分の検索からのみ除外)
    On Error Resume Next
    modP2P.EmitNoiseVote srcName          ' 組織的除外への1票(共有フォルダ・同期時に集計)
    On Error GoTo 0

    Dim ans As Long
    ans = MsgBox("『" & srcName & "』を品質報告しました。" & vbLf & vbLf & _
                 "・あなたの検索からは今すぐ除外されます。" & vbLf & _
                 "・異なる" & modStats.NoiseThreshold() & _
                 "人以上が報告すると、組織全体の検索から除外されます。" & vbLf & vbLf & _
                 "間違えて押した場合は [いいえ] で取り消せます。", _
                 vbYesNo + vbInformation + vbDefaultButton1, _
                 modAppDef.APP_NAME)
    ' R49 Fix(敵対的レビュー R49-REV-08): vbOKCancel だと **Esc が「取り消す」に
    ' なる**。Esc はダイアログを閉じる万能操作として使われるので、意図して
    ' 報告した人が Esc で閉じると報告が黙って取り消される ―― 既定ボタンの罠を
    ' vbDefaultButton3 で潰した直後に、Esc 側へ同じ罠を作っていた。
    ' vbYesNo には Cancel が無いので Esc では閉じられない＝誤爆の道が消える。
    If ans = vbNo Then Undo srcName
End Sub

' ----------------------------------------------------------------------------
' Undo - 自分の報告だけを引っ込める。
'   ・個人ミュート(my_stats "noise:<資料名>")を消す。
'   ・共有フォルダの自分の票ファイルを消す。
'   他人の票・確定フラグ(gexcl)には触らない。
'
'   【R49 Fix・敵対的レビュー R49-REV-01/02】初版は3つとも嘘をつきえた。
'   どれも「取り消しの処理は正しいが、文面が結果を確かめずに断言していた」型。
'
'   (a) **「あなたの検索に戻ります」が戻らないことがある。**
'       modRetrieve が見る modStats.ExcludedSources は noise: と gexcl: の【和】。
'       組織的除外が確定した資料は gexcl が立っているので、個人ミュートを
'       消しても検索には戻らない。初版は gexcl を一度も見ていなかった。
'   (b) **「1票も引っ込めました」が引っ込めていないことがある。**
'       CollectNoiseVotes は閾値に達した資料の確定フラグを書いたあと、
'       GcNoiseVotesForSource で【個別票を全部削除する】。以後こちらの
'       票ファイルは存在せず、KillRetry は「既に無い＝削除目的は達成」として
'       True を返す(並行GC耐性の意図的な仕様)。さらに次の集計は確定フラグを
'       読んで【票を数え直さない】ので、1票引いても復帰しない。
'   (c) **ミュートが元から無かったときも「取り消しました」と言っていた。**
'       muteCleared を取っておきながら、文面に一度も使っていなかった。
'
'   → 結果を見てから言う。確定済みのときは「管理者に相談」まで案内する
'      (憲章§3-3・できないことを言わない)。
' ----------------------------------------------------------------------------
Public Sub Undo(ByVal srcName As String)
    If LenB(srcName) = 0 Then Exit Sub

    Dim muteCleared As Boolean
    muteCleared = modStats.UnreportNoise(srcName)

    Dim voteCleared As Boolean
    On Error Resume Next
    voteCleared = modP2P.RetractNoiseVote(srcName)
    On Error GoTo 0

    ' 組織的除外が確定しているか。ここが True なら、個人ミュートを消しても
    ' 検索には戻らないし、1票引いても集計は確定フラグを優先する。
    Dim orgExcluded As Boolean
    On Error Resume Next
    orgExcluded = modStats.IsGloballyExcluded(srcName)
    On Error GoTo 0

    On Error Resume Next
    modLog.LogUsage "noise_undo", "", _
        "mute=" & muteCleared & " vote=" & voteCleared & _
        " gexcl=" & orgExcluded & " src=" & modUtil.SafeLeft(srcName, 120)
    On Error GoTo 0

    Dim msg As String
    msg = "『" & srcName & "』の品質報告を取り消しました。" & vbLf

    ' (1) 自分の検索に戻るか
    If orgExcluded Then
        msg = msg & modEmj.Warn() & "ただし、この資料は" & _
              "【部内の複数の人からの報告で除外が確定している】ため、" & _
              "あなたの検索にはまだ戻りません。" & vbLf & _
              "戻すには管理者による解除が必要です。" & vbLf
    ElseIf muteCleared Then
        msg = msg & "・あなたの検索に戻ります。" & vbLf
    Else
        msg = msg & "(この資料はもともとあなたの検索から除外されていません" & _
              "でした。)" & vbLf
    End If

    ' (2) 部内へ送った1票
    If orgExcluded Then
        msg = msg & "・除外が確定したあとなので、1票だけを引っ込めることは" & _
              "できません。"
    ElseIf voteCleared Then
        msg = msg & "・部内へ送った1票も引っ込めました" & _
              "(全員の画面へは次の同期で反映されます)。"
    Else
        msg = msg & modEmj.Warn() & "部内へ送った1票は" & _
              "引っ込められませんでした" & _
              "(共有フォルダに繋がっていない可能性があります)。" & vbLf & _
              "ネットワークに繋がってから、もう一度取り消してください。"
    End If
    MsgBox msg, vbInformation, modAppDef.APP_NAME
End Sub
