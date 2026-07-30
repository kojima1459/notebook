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
