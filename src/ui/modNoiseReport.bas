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
                 "間違えて押した場合は [キャンセル] で取り消せます。", _
                 vbOKCancel + vbInformation + vbDefaultButton1, _
                 modAppDef.APP_NAME)
    If ans = vbCancel Then Undo srcName
End Sub

' ----------------------------------------------------------------------------
' Undo - 自分の報告だけを引っ込める。
'   ・個人ミュート(my_stats "noise:<資料名>")を消す → 自分の検索に戻る。
'   ・共有フォルダの自分の票ファイルを消す → 次の同期で全員の集計から外れる。
'   他人の票・確定フラグ(gexcl)には触らない。既に閾値へ達していた資料は、
'   自分が1票引いた結果として閾値を下回れば次の集計で自然に復帰する
'   (modP2P.CollectNoiseVotes は毎回ファイルから数え直す設計)。
'   そこまでは約束できないので、文面では「あなたの検索」と「あなたの1票」
'   だけを言う。できないことを言わない(憲章§3-3)。
' ----------------------------------------------------------------------------
Public Sub Undo(ByVal srcName As String)
    If LenB(srcName) = 0 Then Exit Sub

    Dim muteCleared As Boolean
    muteCleared = modStats.UnreportNoise(srcName)

    Dim voteCleared As Boolean
    On Error Resume Next
    voteCleared = modP2P.RetractNoiseVote(srcName)
    On Error GoTo 0

    On Error Resume Next
    modLog.LogUsage "noise_undo", "", _
        "mute=" & muteCleared & " vote=" & voteCleared & _
        " src=" & modUtil.SafeLeft(srcName, 120)
    On Error GoTo 0

    ' 票を引っ込められたかどうかで文面を変える。共有フォルダに届いていない
    ' (ネットワーク断・共有フォルダ未設定)のに「取り消しました」と言うと、
    ' 画面が嘘をつくことになる。
    Dim msg As String
    msg = "『" & srcName & "』の品質報告を取り消しました。" & vbLf & _
          "・あなたの検索に戻ります。"
    If voteCleared Then
        msg = msg & vbLf & "・部内へ送った1票も引っ込めました" & _
              "(全員の画面へは次の同期で反映されます)。"
    Else
        msg = msg & vbLf & modEmj.Warn() & _
              "部内へ送った1票は引っ込められませんでした" & _
              "(共有フォルダに繋がっていない可能性があります)。" & vbLf & _
              "ネットワークに繋がってから、もう一度取り消してください。"
    End If
    MsgBox msg, vbInformation, modAppDef.APP_NAME
End Sub
