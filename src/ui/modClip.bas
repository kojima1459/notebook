Attribute VB_Name = "modClip"
Option Explicit

' ============================================================================
' modClip - クリップボード書き込み(文字化けしないUnicode安全方式)。盲点C2/D6。
' ----------------------------------------------------------------------------
' 役割:
'   AIバブルの本文テキストをクリップボードへ格納し、ユーザーがCtrl+Vで貼り付け
'   できるようにする(Shape=図形の文字は手で綺麗にコピーできないため、明示的な
'   「コピー」ボタンが実務上必須)。
'
' 設計判断(盲点D6=文字化け回避):
'   ・cmd.exe /c echo ... | clip 方式は使わない。clip.exeがコンソールの
'     コードページで受け取るため、日本語(特にShift-JIS外の文字)や改行が
'     「??」に化ける。DOS窓が一瞬光る副作用もある。
'   ・MSForms.DataObject を採用。ただしライブラリ参照(FM20.dll)を前提にすると
'     自己インストーラ配布(標準モジュールのみ注入)と両立しないため、参照無しで
'     生成できる GetObject("new:{CLSID}") 方式でLate Bindingする。SetTextは
'     Unicodeをそのまま保持するので化けない。
'   ・{1C3B4210-F441-11CE-B9EA-00AA006B1A69} = "Microsoft Forms 2.0 DataObject"。
'   ・生成/PutInClipboardが失敗し得る環境(FM20.dll未登録等)に備え、戻り値Boolで
'     成否を返し、呼び出し側が代替案内を出せるようにする。COMは正常・異常とも
'     Set = Nothing で解放(§メモリリーク対策と同じ規律)。
' ============================================================================

Private Const DATAOBJECT_CLSID As String = "new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"

' ----------------------------------------------------------------------------
' SetClipboardText - textをクリップボードへ格納。成功=True。
' ----------------------------------------------------------------------------
Public Function SetClipboardText(ByVal text As String) As Boolean
    Dim dobj As Object
    On Error GoTo Fail
    Set dobj = GetObject(DATAOBJECT_CLSID)
    dobj.SetText text
    dobj.PutInClipboard
    Set dobj = Nothing
    SetClipboardText = True
    Exit Function
Fail:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup0
FailCleanup0:
    On Error Resume Next
    Set dobj = Nothing
    On Error GoTo 0
    SetClipboardText = False
End Function

' ----------------------------------------------------------------------------
' CopyOrGuide - クリップボードへ入れ、結果に【合った】案内文を返す。
' ----------------------------------------------------------------------------
'   【直したこと】(2026-08-16 R33 W4-7)
'   SetClipboardText は上のヘッダーどおり「失敗し得る前提」で戻り値を返す
'   契約なのに、4箇所(modHelp のご意見箱 / modHub の共有失敗時の逃がし /
'   modPeek のパス案内 / modHubStat のポータルURL)が Sub 呼び出しの形で
'   戻り値を捨て、その直後に「クリップボードに入っています」と断言して
'   いた。RDPのクリップボード同期中や履歴ツールが掴んでいる最中は
'   PutInClipboard が実際に失敗する(そのとき本関数は例外を投げずに
'   False を返して正常復帰するので、呼び出し側の On Error では気づけない)。
'   利用者は案内どおり Ctrl+V を押し、【その前にあった別の内容】を貼り、
'   ご意見・不具合報告の本文は完全に失われる。しかも本人は送ったつもりで、
'   EXP も usage_log も付くので後から失敗と分からない。
'
'   案内文の分岐を各呼び出し側に書くと、字数の逼迫したモジュール
'   (modHub 残159字 / modHubStat 残103字)に入らないうえ、失敗時の文面が
'   4通りに散る。コピーの成否を知っているのはこのモジュールなので、
'   「成否に応じた案内文を返す」までを契約に含める。
'   失敗文には本文そのものを添える ―― 手で控える道を残さないと、
'   利用者は自分が書いた文章を復元できない(MsgBox の1024字制限と
'   読みやすさのため 400字で切る)。
' ----------------------------------------------------------------------------
'
'   【2026-08-16 R33H F32: 主因の説明まで一緒に消していた】
'   okMsg には「何が起きたか」と「だからこうしてください」が1本に混ざって
'   いることがある(modHubStat は FriendlyMessage("E0905")=「社内ポータルを
'   このパソコンから開けませんでした。アドレスをコピーしましたので…」を
'   渡す)。コピーに失敗すると okMsg ごと差し替わるので、【ポータルが開けな
'   かったという主因そのもの】が画面から消え、利用者には「クリップボードに
'   入れられませんでした」だけが出る ―― 押したボタンと無関係な話に見える。
'   主因の1文は failLead として別に受け取り、失敗文の【前置き】に必ず残す。
'   省略可能にしてあるのは、主因が案内文と分けられない残り3箇所(ご意見箱・
'   共有失敗の逃がし・パス案内)の呼び方を変えないため。
' ----------------------------------------------------------------------------
Public Function CopyOrGuide(ByVal text As String, ByVal okMsg As String, _
                            Optional ByVal failLead As String = "") As String
    If SetClipboardText(text) Then
        CopyOrGuide = okMsg
        Exit Function
    End If
    Dim lead As String
    If LenB(Trim$(failLead)) > 0 Then lead = Trim$(failLead) & vbCrLf & vbCrLf
    CopyOrGuide = lead & "クリップボードに入れられませんでした" & _
                  "(他のアプリが使用中の可能性があります)。" & vbCrLf & _
                  "お手数ですが、次の内容を手で控えてください。" & vbCrLf & vbCrLf & _
                  modUtil.SafeLeft(text, 400)
End Function

' CopyOrGuideBox - 上の結果をそのまま MsgBox で見せるだけの薄い包み。
'   呼び出し元(modHubStat)は残72字しか無く、引数を1本増やすと入らない。
'   憲法の「残り300字未満は実体を余裕モジュールへ置き、1行呼び出しに留める」
'   に従い、MsgBox の型(vbInformation とタイトル)ごとこちらへ持ってきた。
Public Sub CopyOrGuideBox(ByVal text As String, ByVal okMsg As String, _
                          Optional ByVal failLead As String = "")
    MsgBox CopyOrGuide(text, okMsg, failLead), vbInformation, modAppDef.APP_NAME
End Sub
