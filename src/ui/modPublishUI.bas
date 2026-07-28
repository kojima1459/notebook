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
    If LenB(chName) = 0 Then GoTo Done

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
        MsgBox "共有フォルダが未設定です。" & vbCrLf & _
               "Hubのお知らせから設定してから、もう一度お試しください。", _
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

    Dim msg As String
    msg = "【" & chName & "】として発行します。" & vbCrLf & vbCrLf & _
          "  発行する件数: " & total & " チャンク(自分で入れた資料のみ)" & vbCrLf & _
          "  配置先: " & dest & vbCrLf & vbCrLf & _
          "【発行前の確認】" & vbCrLf & _
          "  ・正典には『確認済みQ&A・要点』を入れてください" & vbCrLf & _
          "  ・個人情報が見つかった場合は自動で中止します" & vbCrLf & vbCrLf
    If LenB(log_) > 0 Then msg = msg & "【この部門の発行履歴】" & vbCrLf & log_ & vbCrLf
    msg = msg & "このまま発行しますか?" & vbCrLf & _
          "(「いいえ」= 直前の版に戻す操作に進みます)"

    Dim ans As VbMsgBoxResult
    ans = MsgBox(msg, vbYesNoCancel + vbQuestion, modAppDef.APP_NAME & " - 発行の確認")
    If ans = vbCancel Then GoTo Done
    If ans = vbNo Then
        modUiLock.Leave
        ' DoRollback は内部で自分のロックを取り直す(レビュー M-21)。
        DoRollback chName
        Exit Sub
    End If

    ' ここから先は待たせるので、必ず砂時計と進捗を出す。
    On Error Resume Next
    Application.Cursor = 2
    Application.StatusBar = "【" & chName & "】を書き出しています..."
    modPublish.ArchiveCurrent chName          ' 旧版を退避(戻せるように)
    On Error GoTo Done

    Dim wrote As Long
    Dim ok As Boolean
    On Error Resume Next
    ok = modPackExport.ExportPackToFile(dest, "", True, wrote, "self")
    Application.Cursor = -4143
    Application.StatusBar = False
    On Error GoTo Done

    If Not ok Then
        modUiLock.Leave
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
    Exit Sub
Done:
    On Error Resume Next
    Application.Cursor = -4143
    Application.StatusBar = False
    On Error GoTo 0
    modUiLock.Leave
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

    Dim parts() As String: parts = Split(list_, "|")
    Dim newest As String: newest = parts(0)

    If MsgBox("【" & chName & "】を直前の版に戻します。" & vbCrLf & vbCrLf & _
              "  戻す版: " & newest & vbCrLf & vbCrLf & _
              "戻すと、部内の全員が次にファイルを開いたときに" & vbCrLf & _
              "その版へ自動で置き換わります(誤った内容は各PCから消えます)。" & vbCrLf & vbCrLf & _
              "実行しますか?", vbOKCancel + vbExclamation, _
              modAppDef.APP_NAME & " - 直前の版に戻す") <> vbOK Then GoTo RollbackDone

    If modPublish.Rollback(chName, newest) Then
        MsgBox "戻しました。全員に自動で配信されます。", vbInformation, modAppDef.APP_NAME
    Else
        MsgBox "戻せませんでした。共有フォルダへの書き込み権限をご確認ください。", _
               vbExclamation, modAppDef.APP_NAME
    End If
    On Error GoTo 0
RollbackDone:
    modUiLock.Leave
End Sub

' 部門名に含まれる、Windowsのフォルダ名として使えない最初の文字を返す。
' 無ければ空文字(レビュー L-13)。
Private Function FirstInvalidPathChar(ByVal s As String) As String
    Dim bad As Variant
    For Each bad In Array("\\", "/", ":", "*", "?", """", "<", ">", "|")
        If InStr(1, s, CStr(bad)) > 0 Then
            FirstInvalidPathChar = CStr(bad)
            Exit Function
        End If
    Next bad
End Function
