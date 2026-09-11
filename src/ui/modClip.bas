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

' R46「前の回答をコピー」用。一覧に並べる本数と、1行見出しの字数。
Private Const PAST_MAX As Long = 10
Private Const PAST_HEAD_CHARS As Long = 28

' OnCopyPastAnswer のエラー退避。宣言はモジュール先頭へ集約する
' (プロシージャの後ろに置くと実機VBAでコンパイルエラーになる)。
Private mFailNum As Long
Private mFailDesc As String

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

' ============================================================================
' 前の回答をコピー(R46)
' ----------------------------------------------------------------------------
' 実機報告「最新の出力内容はコピーできるけど、それより前の出力内容はコピー
' できない」。原因は modUINexusDraw.DrawContextActions の1行目が
' ClearContextActions で、コピー等の pill が【最新のAIバブルの直下にだけ】
' 描かれること。本文そのものは Shape として最大40件(modUI.MAX_BUBBLES)
' 画面に残っており、消えているのはボタンだけだった。
'
' 保存先として「チャット履歴」シートも存在するが、そちらは本文を 5,000字で
' 【印も付けずに】切る(modChatLog.MAX_ANS_CHARS)。切れたことが読み手に
' 分からない写しを配るのは、この製品が守るべき「出力の誠実さ」と逆向きなので
' 採らない。画面の Shape から読めば、表示されているものと1字も違わない。
' 代償はセッション内(直近40件)に限られること。
'
' 導線はヘルプカード(? ガイド)の7段目。文脈 pill を8個目にする案は
' modUINexusDraw が残21字で入らず、そこへ手を出すと分割裁定=大手術になる。
' ============================================================================

' OnCopyPastAnswer - 「前の回答をコピー」ハンドラ。番号で選ばせて写す。
Public Sub OnCopyPastAnswer()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Failed

    Dim names() As String, heads() As String
    Dim n As Long
    n = CollectAiBubbles(names, heads)
    If n < 1 Then
        MsgBox "コピーできる回答がまだありません。", vbInformation, modAppDef.APP_NAME
        GoTo Done
    End If

    Dim shown As Long: shown = n
    If shown > PAST_MAX Then shown = PAST_MAX

    Dim listText As String
    Dim i As Long
    For i = 1 To shown
        listText = listText & CStr(i) & ". " & heads(i) & vbLf
    Next i

    Dim v As Variant
    v = Application.InputBox( _
        "コピーしたい回答の番号を入れてください(1 = いちばん新しい回答)。" & _
        vbLf & vbLf & listText, "前の回答をコピー", 1, Type:=1)
    If VarType(v) = vbBoolean Then GoTo Done    ' キャンセル

    Dim idx As Long: idx = CLng(v)
    If idx < 1 Then GoTo Done
    If idx > shown Then GoTo Done

    CopyOrGuideBox modUI.BubbleTextOf(names(idx)), _
        CStr(idx) & "番目の回答をコピーしました。貼り付け先で Ctrl+V を押してください。"

Done:
    modUiLock.Leave
    Exit Sub

Failed:
    ' Err は On Error / Exit で消えるので、後始末の前に退避する(CLAUDE.md §11)。
    mFailNum = Err.Number
    mFailDesc = Err.Description
    ' Resume で【ハンドラ稼働中】の状態を解除してから後始末へ行く。
    ' ハンドラの中で On Error Resume Next を書いても効かず、そこで起きた
    ' エラーが呼び出し元へ飛んで本来の原因を上書きする(modEmbed の EscOrErr と
    ' 同じ作法)。
    Resume Cleanup

Cleanup:
    On Error Resume Next
    modLog.LogError "E0906", "modClip.OnCopyPastAnswer", _
        CStr(mFailNum) & " " & mFailDesc
    modUiLock.Leave
    On Error GoTo 0
End Sub

' CollectAiBubbles - Nexus 上のAI回答バブルを【新しい順】で集める。
'   「新しい=下にある」なので Top の降順。AddChatBubble は積み上げ式で、
'   modUI.LatestAiBubbleName も同じ基準(Top+Height の最大)を使っている。
Private Function CollectAiBubbles(ByRef namesOut() As String, _
                                  ByRef headsOut() As String) As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Nexus")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim nm() As String, tp() As Double
    Dim cnt As Long
    cnt = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 9) = "nx_msg_a_" Then
            cnt = cnt + 1
            ReDim Preserve nm(1 To cnt)
            ReDim Preserve tp(1 To cnt)
            nm(cnt) = shp.Name
            tp(cnt) = shp.Top
        End If
    Next shp
    If cnt = 0 Then Exit Function

    ' Top の降順へ単純挿入で並べ替える(件数は最大40。modUI.MAX_BUBBLES)。
    Dim i As Long, j As Long
    Dim tmpN As String, tmpT As Double
    For i = 2 To cnt
        tmpN = nm(i): tmpT = tp(i)
        j = i - 1
        Do While j >= 1
            If tp(j) >= tmpT Then Exit Do
            nm(j + 1) = nm(j)
            tp(j + 1) = tp(j)
            j = j - 1
        Loop
        nm(j + 1) = tmpN
        tp(j + 1) = tmpT
    Next i

    ReDim namesOut(1 To cnt)
    ReDim headsOut(1 To cnt)
    For i = 1 To cnt
        namesOut(i) = nm(i)
        headsOut(i) = HeadLineOf(modUI.BubbleTextOf(nm(i)))
    Next i
    CollectAiBubbles = cnt
End Function

' HeadLineOf - 一覧へ出す1行見出し(純関数)。先頭行だけを取り、長ければ
'   末尾に … を付けて切る。空なら「(空の回答)」。
Public Function HeadLineOf(ByVal s As String) As String
    Dim t As String
    t = Replace(Replace(s, vbCr, vbLf), vbLf & vbLf, vbLf)
    Dim p As Long
    p = InStr(1, t, vbLf, vbBinaryCompare)
    If p > 0 Then t = Left$(t, p - 1)
    t = Trim$(t)
    If Len(t) > PAST_HEAD_CHARS Then t = Left$(t, PAST_HEAD_CHARS) & ChrW(&H2026)
    If LenB(t) = 0 Then t = "(空の回答)"
    HeadLineOf = t
End Function
