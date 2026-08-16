Attribute VB_Name = "modPublishUI"
Option Explicit

' ========================================
' modPublishUI - 部門正典の「発行」と「巻き戻し」の画面側
'
' modKnowledge から切り出した、発行者だけが使う導線。
' 利用者(読む人)が触るのは modKnowledge、発行者(書く人)が触るのはここ、
' と分けることで、片方の修正がもう片方の画面を壊す事故を減らす。
'
' 切り出しの理由(2026-07-28): modKnowledge が契約上限30,000字に対し
' 残り295字となり、部門名の検証(レビュー L-13)を入れた時点で
' 次の修正が入らなくなっていた(レビュー I-2)。
' ========================================

' いま発行ロックを握っている部門名(2026-07-31 R8 F12)。
' 途中でエラーが起きても Done ハンドラから確実に外すために持つ。
' VBAはモジュールレベル宣言をプロシージャより前に置く必要がある。
Private mLockedChannel As String

' ----------------------------------------------------------------------------
' 正典の発行(1操作で完結)。
'   以前は「保存先を手で貼り付けて保存 → もう一度発行を押す」の2段だった。
'   実機で共有フォルダに channels\<部門>\ が作られず止まった原因がこれ。
'   保存先はアプリが知っているのだから、人に貼らせる理由が無い。
'   ボタン1回で 書き出し → 旧版退避 → 配置 → version.txt まで通す。
' ----------------------------------------------------------------------------
Public Sub OnPublish()
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim chName As String
    chName = Trim$(InputBox( _
        "部門の公式ナレッジ(正典)を発行します。" & vbCrLf & vbCrLf & _
        "【ボタン1回で完了します】" & vbCrLf & _
        "  今の本棚の内容がそのまま部内へ配られます。" & vbCrLf & _
        "  保存先を選ぶ操作は要りません。" & vbCrLf & vbCrLf & _
        "【まちがえても大丈夫です】" & vbCrLf & _
        "  前の版は自動で保存されます。もう一度この画面を開いて" & vbCrLf & _
        "  「いいえ」を選ぶと、直前の版に戻せます。" & vbCrLf & vbCrLf & _
        "発行する部門名を入力してください(例: 商品部)。" & vbCrLf & _
        "※ 更新するときも、前と同じ名前を入れてください。", _
        modAppDef.APP_NAME & " - 正典を発行"))
    ' 2026-07-31(R11-A C3): 取りやめの経路は Done: ではなく DoneCleanup0 へ
    ' 直行する。Done: の Resume は「エラーが起きていない」状態で実行すると
    ' 実行時エラー20を起こし、【取りやめただけなのに例外扱い】になるため
    ' (Done: に例外の通知を足した結果、その違いが表に出るようになった)。
    If LenB(chName) = 0 Then GoTo DoneCleanup0

    If Not modPublish.VerifyKey(InputBox( _
            "発行キーを入力してください。" & vbCrLf & _
            "(分からない場合は、このツールの管理担当者に確認してください)", _
            modAppDef.APP_NAME & " - 発行キー")) Then
        modUiLock.Leave
        MsgBox "発行キーが違います。発行は行いませんでした。", vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' 共有フォルダに書ける状態かを、書き出す前に確かめる。
    ' 権限が無いフォルダを指しているのが実機で一番多い失敗なので、
    ' 重い書き出しをやってから失敗させない。
    ' 2026-07-28(レビュー L-13): 部門名にファイル名として使えない文字が
    ' 入っていないか、重い書き出しの前に確かめる。従来は無検証で進み、
    ' 後段の SaveAs が失敗したときに「共有フォルダへの書き込み権限を
    ' ご確認ください」という【見当違いの案内】を出していた。
    ' 権限を疑って調べ回った末に、原因が名前の "/" だった、では時間の無駄。
    Dim badChar As String
    badChar = FirstInvalidPathChar(chName)
    If LenB(badChar) > 0 Then
        modUiLock.Leave
        MsgBox "部門名に使えない文字が含まれています: " & badChar & vbCrLf & vbCrLf & _
               "次の文字はフォルダ名に使えません:" & vbCrLf & _
               "  \ / : * ? "" < > |" & vbCrLf & vbCrLf & _
               "別の名前でもう一度お試しください。", vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    Dim dest As String
    On Error Resume Next
    dest = modPublish.PackDestPath(chName)
    On Error GoTo Done
    If LenB(dest) = 0 Then
        modUiLock.Leave
        ' 2026-07-31(R8b B14): dest が空になる理由は「未設定」だけではない。
        ' modShare の関所を通すようにしたため、【設定済みだが今は届かない】
        ' (VPN未接続・サーバ停止・セッション降格)でも空が返る。
        ' 「未設定です」と断定すると、設定済みの発行者が config を疑って
        ' 探し回ることになる。両方の可能性をそのまま示す。
        MsgBox "共有フォルダが未設定か、共有フォルダに接続できません。" & vbCrLf & vbCrLf & _
               "・まだ設定していない場合: Hubのお知らせから設定してください" & vbCrLf & _
               "・設定済みの場合: 社内ネットワーク(VPN)に接続できているか" & vbCrLf & _
               "  ご確認のうえ、ファイルを開き直してからお試しください", _
               vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' 2026-07-28(レビュー H-4): 発行するのは「自分で入れた資料」だけ
    ' (origin="self")。購読中の他部門の正典まで自部門のパックへ混ぜて
    ' 再配布しないため。確認ダイアログの件数も、本棚の総数ではなく
    ' 実際に出る件数に揃える(数字が違うと発行者が気付けない)。
    Dim total As Long
    On Error Resume Next
    total = modShelfStore.CountRowsByOrigin("self")
    Dim log_ As String: log_ = modPublish.RecentLog(chName)
    On Error GoTo Done

    ' 2026-08-14(R32 W2-1): PII走査はconfig pii_scan_enabled(既定FALSE)で
    ' 有効化したときだけ動く。オフのまま「自動で中止します」とだけ案内すると
    ' 実際には動いていないチェックを期待させてしまうため、オフ時は現在の
    ' 状態が分かる一文へ出し分ける。
    Dim piiLine As String
    If modConfig.GetBool("pii_scan_enabled", False) Then
        piiLine = "  ・個人情報が見つかった場合は自動で中止します"
    Else
        piiLine = "  ・(個人情報の自動チェックは現在オフです)"
    End If

    Dim msg As String
    msg = "【" & chName & "】として発行します。" & vbCrLf & vbCrLf & _
          "  発行する件数: " & total & " チャンク(自分で入れた資料のみ)" & vbCrLf & _
          "  配置先: " & dest & vbCrLf & vbCrLf & _
          "【発行前の確認】" & vbCrLf & _
          "  ・正典には『確認済みQ&A・要点』を入れてください" & vbCrLf & _
          piiLine & vbCrLf & vbCrLf
    If LenB(log_) > 0 Then msg = msg & "【この部門の発行履歴】" & vbCrLf & log_ & vbCrLf
    msg = msg & "このまま発行しますか?" & vbCrLf & _
          "(「いいえ」= 直前の版に戻す操作に進みます)"

    Dim ans As VbMsgBoxResult
    ans = MsgBox(msg, vbYesNoCancel + vbQuestion, modAppDef.APP_NAME & " - 発行の確認")
    If ans = vbCancel Then GoTo DoneCleanup0
    If ans = vbNo Then
        modUiLock.Leave
        ' DoRollback は内部で自分のロックを取り直す(レビュー M-21)。
        DoRollback chName
        Exit Sub
    End If

    ' 2026-07-31(レビュー R8 F12): 同時発行の見張り。
    ' 2人が同じ部門を同時に発行すると、片方の pack.xlsx ともう片方の
    ' version.txt が組み合わさった状態になり得る。どちらの発行者にも
    ' エラーは出ないまま、購読者だけが「新しい版番号の古い中身」を掴む。
    Dim holder As String
    Dim gotLock As Boolean
    On Error Resume Next
    gotLock = modPublish.AcquireLock(chName, holder)
    On Error GoTo Done
    If Not gotLock Then
        modUiLock.Leave
        ' 2026-07-31(C3): 取れなかった理由は2つあり、対処がまるで違う。
        '   (a) 他の人のロックが実在する(holder が読めている)→ 待てば解ける
        '   (b) ロックを作れなかった(権限不足・共有断。R8b B2 で False を返す)
        '       → いくら待っても解けない。権限と接続を見るしかない
        ' holder が空のまま「いま別の方が発行中です」と出すと、実在しない
        ' 相手を探して待ち続けることになる(発行担当は1人のことも多く、
        ' 「自分しかいないのに誰が発行中なのか」で必ず詰まる)。
        If LenB(holder) > 0 Then
            MsgBox "いま別の方が【" & chName & "】を発行中です。" & vbCrLf & vbCrLf & _
                   "  発行を始めた方: " & holder & vbCrLf & vbCrLf & _
                   "同時に発行すると内容と版番号が食い違うことがあるため、" & vbCrLf & _
                   "今回は発行しませんでした。数分おいてもう一度お試しください。", _
                   vbExclamation, modAppDef.APP_NAME
        Else
            MsgBox "発行の見張り(publish.lock)を作成できませんでした。" & vbCrLf & vbCrLf & _
                   "共有フォルダへの書き込み権限と、社内ネットワークへの接続を" & vbCrLf & _
                   "ご確認のうえ、もう一度お試しください。" & vbCrLf & vbCrLf & _
                   "安全のため、見張りを置けない状態では発行しません" & vbCrLf & _
                   "(2人が同時に発行すると、内容と版番号が食い違います)。", _
                   vbExclamation, modAppDef.APP_NAME
        End If
        Exit Sub
    End If
    mLockedChannel = chName        ' 失敗経路でも必ず外すための控え

    ' ここから先は待たせるので、必ず砂時計と進捗を出す。
    ' 2026-07-31(レビュー R8 F12): 重い書き出しは【ローカルの%TEMP%】へ行い、
    ' 出来上がったものだけを共有へ据える。共有フォルダ上に「書きかけの
    ' pack.xlsx」が見えている時間を、数十秒から1回のコピーへ縮める。
    Dim stage As String
    On Error Resume Next
    stage = modPublish.StagePackPath(chName)
    On Error GoTo Done
    If LenB(stage) = 0 Then stage = dest      ' %TEMP%が引けない環境は従来どおり
    Dim arcFailed As Boolean
    If stage = dest Then
        ' 共有へ直接書く従来経路。書き出しの前に旧版を退避しておかないと
        ' 巻き戻し先が消える。
        On Error Resume Next
        modPublish.ArchiveCurrent chName, arcFailed
        On Error GoTo Done
        ' 2026-07-31(R11-A C3): 退避に失敗したまま進めると、発行画面で
        ' 約束した「前の版に戻せます」が果たせなくなる。黙って進めない。
        If arcFailed Then
            If Not ConfirmArchiveFailure() Then
                AbortPublish chName, stage, dest
                Exit Sub
            End If
        End If
    End If

    On Error Resume Next
    Application.Cursor = 2
    Application.StatusBar = "【" & chName & "】を書き出しています..."
    On Error GoTo Done

    Dim wrote As Long
    Dim ok As Boolean
    Dim piiAborted As Boolean
    On Error Resume Next
    ok = modPackExport.ExportPackToFile(stage, "", True, wrote, "self", piiAborted)
    Application.Cursor = -4143
    Application.StatusBar = False
    On Error GoTo Done

    If ok And stage <> dest Then
        ' 共有側を触るのはここから。退避も、据え付ける直前に行う
        ' (書き出しに失敗したときに旧版を動かさないため)。
        On Error Resume Next
        Application.StatusBar = "【" & chName & "】を共有フォルダへ置いています..."
        modPublish.ArchiveCurrent chName, arcFailed   ' 旧版を退避(戻せるように)
        Application.StatusBar = False
        On Error GoTo Done
        ' 据え付ける(=旧版を上書きする)【前】に確認する。ここを通り過ぎた
        ' あとでは、もう戻せる版は残っていない(R11-A C3)。
        If arcFailed Then
            If Not ConfirmArchiveFailure() Then
                AbortPublish chName, stage, dest
                Exit Sub
            End If
        End If

        On Error Resume Next
        Application.StatusBar = "【" & chName & "】を共有フォルダへ置いています..."
        ok = modPublish.CommitPack(chName, stage)
        Application.StatusBar = False
        On Error GoTo Done
    End If

    If Not ok Then
        On Error Resume Next
        modPublish.ReleaseLock chName
        mLockedChannel = ""
        On Error GoTo 0
        modUiLock.Leave
        ' 2026-08-14(R32 W2-2): PII検知による中止は、原因も対処も
        ' modPackExport.ExportPackToFile側のMsgBox(E0703)で既に案内済み。
        ' ここで続けて「共有フォルダへの書込権限/パスをご確認ください」を
        ' 出すと、原因確定済みの中止に対して見当違いの案内を重ねることになる
        ' (実機で原因確定済みの不具合)。PII由来のときは無言で退出する。
        If piiAborted Then Exit Sub
        MsgBox "発行できませんでした。" & vbCrLf & vbCrLf & _
               "・共有フォルダに書き込む権限があるか" & vbCrLf & _
               "・パスが正しいか(" & dest & ")" & vbCrLf & _
               "をご確認ください。", vbExclamation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' pack.xlsx を置き終えてから version.txt を書く。順序が逆だと
    ' 「新しい版番号なのに中身が古い」を掴む人が出る。
    Dim ver As String
    On Error Resume Next
    ver = modPublish.FinalizePublish(chName, wrote)
    modStats.Bump "publish_total"
    modPublish.ReleaseLock chName
    mLockedChannel = ""
    On Error GoTo Done

    modUiLock.Leave
    If LenB(ver) > 0 Then
        MsgBox "発行しました。" & vbCrLf & vbCrLf & _
               "  部門: " & chName & vbCrLf & _
               "  件数: " & wrote & " チャンク" & vbCrLf & _
               "  版: " & ver & vbCrLf & vbCrLf & _
               "部内の全員が、次にファイルを開いたときに受け取れます。" & vbCrLf & _
               "内容に誤りが見つかったら、この画面から「いいえ」で戻せます。", _
               vbInformation, modAppDef.APP_NAME
    Else
        MsgBox "ファイルは置けましたが、版の記録に失敗しました。" & vbCrLf & _
               "共有フォルダの書き込み権限をご確認のうえ、もう一度発行してください。", _
               vbExclamation, modAppDef.APP_NAME
    End If
    ' R11-A2(Med-1): 正常終了も後始末を1箇所に集める。ここで Exit Sub して
    ' しまうと、あとから発行ロックを使う処理をこの手前に足したときに
    ' 「この経路だけ外し忘れ」が起きる(10分間その部門の発行が止まる)。
    ' 発行ロックも modUiLock も既に外してあり、どちらの解放も繰り返して
    ' 安全なので、通る道筋は変わらない(failNum=0 のため通知も出ない)。
    GoTo DoneCleanup0
Done:
    ' 2026-07-31(R11-A C3): 例外でここへ来た場合、従来は後始末だけして
    ' 【完全に無言で】終わっていた。押した人から見れば「発行を押したのに
    ' 何も起きない」で、err_log にも1行も残らない(憲章§3-1/§4-1違反)。
    ' Resume はハンドラを抜けると同時に Err を消すので、先に控えておく。
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume DoneCleanup0
DoneCleanup0:
    On Error Resume Next
    Application.Cursor = -4143
    Application.StatusBar = False
    ' 2026-07-31(レビュー R8 F12): どの経路で抜けても発行ロックは必ず外す。
    ' 外し忘れると、その部門は10分間(LOCK_STALE_MIN)誰も発行できなくなる。
    If LenB(mLockedChannel) > 0 Then
        modPublish.ReleaseLock mLockedChannel
        mLockedChannel = ""
    End If
    On Error GoTo 0
    modUiLock.Leave

    ' 例外で中断したときだけ、記録と通知を必ず出す。
    ' failNum=20 は「エラーが起きていないのに Resume した」ときの番号で、
    ' 取りやめの経路が Done: を踏んだ場合にだけ出る自作自演の値。
    ' 利用者の操作は正常なので通知しない(取りやめの経路自体も
    ' DoneCleanup0 直行に変えてあるが、二重の安全弁として残す)。
    If failNum <> 0 And failNum <> 20 Then
        On Error Resume Next
        modLog.LogError "E0808", "modPublishUI.OnPublish", _
            "発行の処理が例外で中断: " & modUtil.SafeLeft(failDesc, 300), failNum
        On Error GoTo 0
        MsgBox modLog.FriendlyMessage("E0808") & vbLf & "(コード: E0808)" & vbLf & vbLf & _
               "「診断」ボタン" & ChrW(&H2192) & "「直近のエラーをコピー」で、" & _
               "詳しい情報をそのまま担当者に送れます。", _
               vbExclamation, modAppDef.APP_NAME
    End If
End Sub

' ----------------------------------------------------------------------------
' ConfirmArchiveFailure - 旧版の退避に失敗したときの確認(R11-A C3)。
'   True = 戻せなくなることを承知のうえで発行を続ける。
' ----------------------------------------------------------------------------
Private Function ConfirmArchiveFailure() As Boolean
    ConfirmArchiveFailure = (MsgBox( _
        "以前の版の控え(アーカイブ)を保存できませんでした。" & vbCrLf & vbCrLf & _
        "このまま発行すると前の版に戻せません。" & vbCrLf & vbCrLf & _
        "・共有フォルダの空き容量と書き込み権限" & vbCrLf & _
        "・社内ネットワーク(VPN)への接続" & vbCrLf & _
        "をご確認のうえ、あとでやり直すこともできます。" & vbCrLf & vbCrLf & _
        "続行しますか?", vbYesNo + vbExclamation, _
        modAppDef.APP_NAME & " - 前の版の控えを保存できません") = vbYes)
End Function

' ----------------------------------------------------------------------------
' AbortPublish - 発行を取りやめて後始末する(R11-A C3)。
'   stage が共有の本番パス(dest)と同じときは消してはいけない。それは
'   いま配られている正典そのもので、消すと全員の受け取り先が無くなる。
' ----------------------------------------------------------------------------
Private Sub AbortPublish(ByVal chName As String, ByVal stage As String, ByVal dest As String)
    On Error Resume Next
    If LenB(stage) > 0 And stage <> dest Then Kill stage
    Application.Cursor = -4143
    Application.StatusBar = False
    modPublish.ReleaseLock chName
    mLockedChannel = ""
    On Error GoTo 0
    modUiLock.Leave
    MsgBox "発行を取りやめました。前の版はそのまま残っています。", _
           vbInformation, modAppDef.APP_NAME
End Sub

' 直前の版に戻す。事故を止める最終手段なので、操作は最短手数にする。
Private Sub DoRollback(ByVal chName As String)
    ' 2026-07-28(レビュー M-21): 呼び出し側が Leave してから来るので、
    ' 実処理はここでロックを取り直す。取れなければ何もしない
    ' (巻き戻しの二重実行は共有フォルダの版を壊す)。
    If Not modUiLock.Enter() Then GoTo RollbackDone
    On Error GoTo RollbackDone
    On Error Resume Next
    Dim list_ As String
    list_ = modPublish.ArchiveList(chName)
    If LenB(list_) = 0 Then
        MsgBox "戻せる過去版がありません(まだ1度も発行していない部門です)。", _
               vbInformation, modAppDef.APP_NAME
        GoTo RollbackDone
    End If

    ' ------------------------------------------------------------------------
    ' 2026-08-16(R33 W2-4・データ破壊): 巻き戻しを2回続けて押すと、1回目が
    ' 取り下げた【誤った版】が復活し、しかも「戻しました」としか出ない。
    '
    ' 機構: modPublish.Rollback は巻き戻す【前】に今の版を _archive へ退避する
    ' (巻き戻し自体を取り消せるように)。退避名は pack_<時刻>.xlsx で、
    ' ArchiveList は名前の降順=時刻の降順に並べる。つまり1回目の巻き戻しが
    ' 作った退避(=いま消したい誤った版)が、次の瞬間から「最新の退避」になる。
    ' ここが無条件に parts(0) を選ぶため、2回目は必然的に誤った版を復元し、
    ' version.txt も必ず新しい値で書かれるので購読者全員が取り込み直す。
    ' 画面は時刻だけのファイル名を見せるので、発行担当は中身を判別できない。
    ' modPublish 冒頭が「承認は入れず、代わりに誰でも即座に巻き戻せることに
    ' 全振りしている」と宣言している、その安全装置が反転する経路。
    '
    ' 直し方(spec W2-4 の「2回目を抑止する」を採る): 直前の操作が巻き戻し
    ' だったら、ここで止める。判定材料は共有フォルダの publish_log.txt なので、
    ' ファイルを開き直しても、別の担当者の端末から押されても効く
    ' (セッション内の変数で覚える方式だと、そのどちらも抜ける)。
    ' 巻き戻しの後に正しい版を発行し直せば履歴の先頭は「発行」に戻るので、
    ' 次の事故のときの巻き戻しは従来どおり1回目として通る。
    ' ------------------------------------------------------------------------
    If LastOpWasRollback(chName) Then
        MsgBox "【" & chName & "】は直前に巻き戻しを実行済みです。" & vbCrLf & vbCrLf & _
               "いま一番新しい控えは、その巻き戻しで【取り下げた版】です。" & vbCrLf & _
               "ここでもう一度戻すと、取り下げたはずの版が全員へ" & vbCrLf & _
               "配信し直されてしまうため、この操作は行いません。" & vbCrLf & vbCrLf & _
               "・戻した内容で問題なければ、このまま何もしないでください" & vbCrLf & _
               "・さらに前の版に戻したい場合は、正しい資料で発行し直すか、" & vbCrLf & _
               "  共有フォルダの _archive フォルダから管理担当者に" & vbCrLf & _
               "  選んでもらってください", _
               vbExclamation, modAppDef.APP_NAME
        GoTo RollbackDone
    End If

    Dim parts() As String: parts = Split(list_, "|")
    Dim newest As String: newest = parts(0)

    If MsgBox("【" & chName & "】を直前の版に戻します。" & vbCrLf & vbCrLf & _
              "  戻す版: " & newest & vbCrLf & vbCrLf & _
              "戻すと、部内の全員が次にファイルを開いたときに" & vbCrLf & _
              "その版へ自動で置き換わります(誤った内容は各PCから消えます)。" & vbCrLf & vbCrLf & _
              "実行しますか?", vbOKCancel + vbExclamation, _
              modAppDef.APP_NAME & " - 直前の版に戻す") <> vbOK Then GoTo RollbackDone

    Dim rbArcFailed As Boolean
    If modPublish.Rollback(chName, newest, rbArcFailed) Then
        ' R11-D(A波発見事項3): 巻き戻しの直前に「いま配っている版」の控えを
        ' 取れなかった場合、この巻き戻しはもう取り消せない。E0808はログに
        ' 残っているが利用者には完全に無言だった。データ保全に関わる知らせ
        ' なので、完了メッセージに必ず併記する(憲章§3-5)。
        If rbArcFailed Then
            MsgBox "戻しました。全員に自動で配信されます。" & vbCrLf & vbCrLf & _
                   "ただし、いま配っていた版の控え(アーカイブ)を保存できません" & _
                   "でした。この巻き戻しを取り消して元の版へ戻すことはできません。" & vbCrLf & _
                   "共有フォルダの空き容量と書き込み権限をご確認ください。" & vbCrLf & _
                   "(コード: E0808)", vbExclamation, modAppDef.APP_NAME
        Else
            MsgBox "戻しました。全員に自動で配信されます。", vbInformation, modAppDef.APP_NAME
        End If
    Else
        MsgBox "戻せませんでした。共有フォルダへの書き込み権限をご確認ください。", _
               vbExclamation, modAppDef.APP_NAME
    End If
    On Error GoTo 0
RollbackDone:
    modUiLock.Leave
End Sub

' 直前の操作が巻き戻しだったか(発行履歴 publish_log.txt の最新1行で判定)。
'   行の形は「<日時>タブ<発行者>タブ<操作>タブ…」。操作欄が "巻き戻し" の
'   ときだけ True。RecentLog は新しい順に返すので、最初の非空行が最新。
'   履歴が読めない/空のときは False ―― 判らないことを理由に、事故を止める
'   最終手段(巻き戻し)そのものを塞いでしまう方が害が大きい(R33 W2-4)。
Private Function LastOpWasRollback(ByVal chName As String) As Boolean
    On Error Resume Next
    Dim log_ As String: log_ = modPublish.RecentLog(chName)
    If LenB(log_) = 0 Then Exit Function

    Dim lines_() As String: lines_ = Split(Replace(log_, vbCrLf, vbLf), vbLf)
    Dim i As Long
    Dim fields_() As String
    For i = LBound(lines_) To UBound(lines_)
        If LenB(Trim$(lines_(i))) > 0 Then
            fields_ = Split(lines_(i), vbTab)
            If UBound(fields_) >= 2 Then
                LastOpWasRollback = (Trim$(fields_(2)) = "巻き戻し")
            End If
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

' 部門名に含まれる、Windowsのフォルダ名として使えない最初の文字を返す。
' 無ければ空文字(レビュー L-13)。
'
' 2026-08-16(R33 W2-5): 先頭要素は "\\" ではなく "\" が正しい。
'   VBA の文字列リテラルにはエスケープが無いので、"\\" は【バックスラッシュ
'   2個】という2文字の文字列になる。InStr は連続した \\ にしか当たらず、
'   `商品部\商品1課` のような単独の \ を含む名前は「使えない文字なし」と
'   判定されて素通りしていた。すり抜けた \ はそのままパスへ連結され、
'   多階層を作れない MkDir が黙って失敗した末に、最後の見張りファイル作成で
'   E0807「共有フォルダへの書き込み権限と、社内ネットワークへの接続を
'   ご確認ください」が出る ―― レビュー L-13 が潰したはずの見当違い案内
'   そのもの(原因は名前の1文字なのに、権限とネットワークを疑わせる)。
'   同型の実装は modChannel.SafeName と modP2PIo.SanitizeId にあり、
'   そちらは正しく1文字の "\" を使っている。書き方をそちらへ揃える。
'   なお 73行目の利用者向け文言は当初から単独の \ を禁止と告知しており、
'   意図(文言)と実装が食い違っていた側を実装へ合わせた形。
Private Function FirstInvalidPathChar(ByVal s As String) As String
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        If InStr(1, s, CStr(bad)) > 0 Then
            FirstInvalidPathChar = CStr(bad)
            Exit Function
        End If
    Next bad
End Function
