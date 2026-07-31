Attribute VB_Name = "modShelfBatch"
Option Explicit

' ============================================================================
' modShelfBatch - 「資料を追加」の一括取込オーケストレーション(R10dで分離)
' ----------------------------------------------------------------------------
' 役割:
'   複数選択のFileDialogを開き、選ばれたファイルを1件ずつ modShelf.IngestFile
'   へ渡し、進捗表示・上限チェック・失敗理由の内訳集計・完了メッセージまでを
'   まとめて面倒みる層。1ファイルぶんの取込そのもの(抽出→チャンク化→
'   my_knowledge追記→manifest→埋め込み)は modShelf の仕事で、ここは
'   「何件を、どの順で、どう見せながら回すか」だけを持つ。
'
' なぜ分けたか(2026-07-31 R10d):
'   modShelf が §7.1 の「1モジュール30,000字以内」に対して残り440字となり、
'   バグ修正1件すら入らない状態になっていた(R10cのH2=バッチ全体の再入
'   ガードが容量不足で実装できず保留になったのが直接のきっかけ)。
'   AddFilesViaDialog / AddFilesResult とその補助(集計文字列の組み立て・
'   失敗理由カウンタ・FileDialogのフィルタ生成)は、modShelf の他の機能
'   (IngestFile / DeleteSource / SourceList / TotalChunks)から
'   【この一式の中だけで完結する】凝集した塊で、外から使われているのは
'   AddFilesViaDialog と AddFilesResult の2本だけだった。切り口として最も
'   素直なのでここを分けた。
'
' 依存:
'   ・modShelf.IngestFile / modShelf.TotalChunks(取込本体と件数)
'   ・modExtractor.SupportedExts(FileDialogのフィルタ生成)
'   ・表示は modUIMain.ShowProgress/HideProgress・modSkin.ShowToast・
'     modUIShelf.RenderShelf(いずれもR1例外の通知コールバック)
'   ・opt名は書かず modFeatures.InvokeFeature 経由(R2)
' ============================================================================

'
' AddFilesViaDialog - 複数選択FileDialog(フィルタ=SupportedExts)→各IngestFile
'   OnAction互換の後方互換ラッパー(リボン/ボタンから直接呼ぶ用)。
'   結果表示はAddFilesResultの既定どおりMsgBoxで行う。
'
Public Sub AddFilesViaDialog()
    AddFilesResult True
End Sub

'
' AddFilesResult - 複数選択FileDialog→各IngestFileの結果をまとめて返す(要件E)。
'   2026-07-30(背景5): modShelf自前のMsgBoxと、呼び出し元modApp.OnAddDocsの
'   チャットバブルが同時に出ていた。呼び出し元が表示を担うときは
'   showMsgBox:=Falseを渡せば、戻り値の集計だけで文言を組み立てられる。
'
'   戻り値: "ok=<n>;ng=<n>;capped=<n>;chunks=<n>;reasons=<code>:<count>,..."
'   reasonsはE0504(同名衝突)/image_pdf/E0302等、失敗理由コードごとの内訳。
'   ngにもcappedにも入らない「キャンセル」は ok=0;ng=0;capped=0;chunks=0 で返す。
'
Public Function AddFilesResult(Optional ByVal showMsgBox As Boolean = True) As String
    ' 集計値の宣言はハンドラより前に置く(途中で落ちても集計を返すため)。
    Dim okCount As Long: okCount = 0
    Dim ngCount As Long: ngCount = 0
    Dim cappedN As Long: cappedN = 0
    Dim chunksAdded As Long: chunksAdded = 0
    ' 失敗理由の内訳(代表的なE0504/image_pdf/E0302等をコード単位で数える)。
    Dim reasonKeys() As String: ReDim reasonKeys(0 To 15)
    Dim reasonCounts() As Long: ReDim reasonCounts(0 To 15)
    Dim reasonN As Long: reasonN = 0

    ' レビュー4-B: プロシージャレベルのハンドラ。従来は素通りで、途中の
    ' 実行時エラーは呼び出し元(modApp.OnAddDocsのOn Error Resume Next)に
    ' 握り潰され、戻り値が空文字=「何も言わない」になっていた。
    On Error GoTo AddFailed

    Dim fd As Object
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker(名前付き定数は使わない)
    fd.AllowMultiSelect = True
    fd.Title = "本棚に追加する資料を選んでください"
    fd.Filters.Clear
    fd.Filters.Add "対応ファイル", BuildFilterPattern()

    If fd.Show <> -1 Then
        AddFilesResult = "ok=0;ng=0;capped=0;chunks=0;reasons="
        Exit Function   ' キャンセル
    End If

    ' R10-2: ファイル選択が確定した(キャンセルではない)ので、GS未検出の
    ' 案内カード(セッション1回きり)を再提示可能に戻す。利用者が能動的に
    ' 取込を実行する入口のみが対象で、自動同期(modShelfSync)からは呼ばない。
    ' vision機能が無効なビルドでもInvokeFeature側で安全に無視される。
    On Error Resume Next
    modFeatures.InvokeFeature "vision", "ResetGsGuidance", Array()
    On Error GoTo AddFailed

    Dim capMax As Long: capMax = modConfig.GetLong("shelf_max_chunks", 10000)
    If capMax < 1 Then capMax = 10000

    ' R10-5(実機報告「止まってるのか分からない」): 経過秒÷完了件数の単純平均で
    ' 目安時間を出す。1件目はまだ実績が無いのでETA無し(modUtil.EtaText任せ)。
    Dim tIngStart As Double: tIngStart = Timer

    Dim i As Long
    For i = 1 To fd.SelectedItems.count
        Dim etaPart As String: etaPart = ""
        If i > 1 Then
            Dim elapsedSec As Double: elapsedSec = Timer - tIngStart
            If elapsedSec >= 0 Then   ' 日跨ぎのTimerロールオーバーはETA非表示に落とす
                etaPart = modUtil.EtaText(fd.SelectedItems.count - (i - 1), _
                    (elapsedSec * 1000#) / (i - 1))
            End If
        End If
        modUIMain.ShowProgress modUtil.ProgressText(i, fd.SelectedItems.count, etaPart) & " " & _
            modUtil.SafeLeft(modUtil.FileNameOf(CStr(fd.SelectedItems(i))), 40)

        If modShelf.TotalChunks() >= capMax Then
            cappedN = cappedN + 1
        Else
            Dim beforeChunks As Long: beforeChunks = modShelf.TotalChunks()
            Dim st As String: st = "failed"
            Dim errCd As String: errCd = ""
            ' silent:=True(レビュー4-A): 1件ごとのモーダルを全廃する。同名衝突
            ' (E0504)だけが silent を見ずにモーダルを出し、選んだ件数ぶん出た
            ' 上で集計メッセージが同じことをもう一度言っていた(二重報告)。
            ' 情報は戻り値の reasons に載るので失わない。
            On Error Resume Next
            st = modShelf.IngestFile(CStr(fd.SelectedItems(i)), "self", True, errCd)
            On Error GoTo AddFailed
            chunksAdded = chunksAdded + (modShelf.TotalChunks() - beforeChunks)
            If st = "done" Or st = "partial" Then
                okCount = okCount + 1
            Else
                ngCount = ngCount + 1
                Dim reasonKey As String
                If st = "image_pdf" Then
                    reasonKey = "image_pdf"
                ElseIf LenB(errCd) > 0 Then
                    reasonKey = errCd
                Else
                    reasonKey = "failed"
                End If
                BumpReasonCount reasonKeys, reasonCounts, reasonN, reasonKey
            End If
        End If
    Next i

    ' R10-5: 進捗バナーを閉じる前に完了を1回だけトーストで知らせる
    ' (連呼はしない。ここは完了1回のみなので1.1秒のブロッキングも許容)。
    On Error Resume Next
    modSkin.ShowToast "取り込みが完了しました(成功" & okCount & "件/失敗" & ngCount & "件)", "info"
    On Error GoTo AddFailed
    modUIMain.HideProgress

    ' silent にしたぶん1件ごとの再描画も走らない(IngestFileのFinishは
    ' silent時RenderShelfを呼ばない)。まとめて1回だけ描き直す(R1例外)。
    On Error Resume Next
    modUIShelf.RenderShelf
    On Error GoTo AddFailed

    Dim msg As String
    msg = okCount & "件を本棚に追加しました。"
    If ngCount > 0 Then
        msg = msg & vbLf & ngCount & "件は取込めませんでした。マイ本棚の一覧で状態をご確認ください。"
    End If
    ' E0504はper-fileモーダルをやめた(4-A)ので、ここで必ず件数を伝える。
    ' 伝えないと「入れたはずの資料が無い」理由がどこにも出ない。
    Dim dupN As Long
    dupN = ReasonCountOf(reasonKeys, reasonCounts, reasonN, "E0504")
    If dupN > 0 Then
        msg = msg & vbLf & "同じ名前の資料が別の場所から登録済みのものが" & dupN & "件あります。" & _
            vbLf & "ファイル名を変えて入れ直すか、マイ本棚で古いほうを削除してからお試しください。"
    End If
    If cappedN > 0 Then
        msg = msg & vbLf & vbLf & _
            "※本棚の上限に達したため" & cappedN & "件は見送りました。" & vbLf & _
            "configシートの shelf_max_chunks を大きくすると上限を増やせます。"
        On Error Resume Next
        modLog.LogError "E0501", "modShelfBatch.AddFilesResult", _
            "上限" & capMax & "到達で" & cappedN & "件見送り"
        On Error GoTo AddFailed
    End If
    If showMsgBox Then
        MsgBox msg, vbInformation, modAppDef.APP_NAME
    End If

    AddFilesResult = ComposeAddResult(okCount, ngCount, cappedN, chunksAdded, _
        ReasonsText(reasonKeys, reasonCounts, reasonN))
    Exit Function

AddFailed:
    ' Err は Resume で消えるので先に控える。
    Dim failNum As Long: failNum = Err.Number
    Dim failDesc As String: failDesc = Err.Description
    ' R6: ハンドラ稼働中は On Error Resume Next が効かない。
    ' 後始末(ログ書込)の前に Resume でハンドラを抜ける。
    Resume AddFailedCleanup
AddFailedCleanup:
    On Error Resume Next
    modLog.LogError "E0801", "modShelfBatch.AddFilesResult", _
        "err#" & failNum & ": " & failDesc & _
        " (ok=" & okCount & " ng=" & ngCount & " capped=" & cappedN & ")"
    modUIMain.HideProgress   ' R10-5: 異常終了経路でも進捗バナーを必ず閉じる
    On Error GoTo 0
    ' 途中で落ちても、そこまでの集計を返す。空文字で返すと呼び出し元は
    ' 「キャンセル」と区別できず、利用者には何も表示されない(レビュー4-B)。
    AddFilesResult = ComposeAddResult(okCount, ngCount, cappedN, chunksAdded, _
        ReasonsText(reasonKeys, reasonCounts, reasonN))
End Function

' 集計結果の文字列化(正常終了と途中失敗の両方から使う。書式は1箇所に持つ)。
Private Function ComposeAddResult(ByVal okCount As Long, ByVal ngCount As Long, _
                                  ByVal cappedN As Long, ByVal chunksAdded As Long, _
                                  ByVal reasonsPart As String) As String
    ComposeAddResult = "ok=" & okCount & ";ng=" & ngCount & ";capped=" & cappedN & _
        ";chunks=" & chunksAdded & ";reasons=" & reasonsPart
End Function

' 失敗理由の内訳を "E0504:2,image_pdf:1" の形へ。
Private Function ReasonsText(ByRef keys() As String, ByRef counts() As Long, _
                             ByVal n As Long) As String
    Dim s As String
    Dim r As Long
    For r = 0 To n - 1
        If LenB(s) > 0 Then s = s & ","
        s = s & keys(r) & ":" & counts(r)
    Next r
    ReasonsText = s
End Function

' ---- 内部ヘルパー ----
' 失敗理由の内訳へ1件加算する(要件E: AddFilesResultの"reasons="組立用)。
' 既存キーがあれば加算、無ければ新規キーとして追加(容量不足時は倍々拡張)。
Private Sub BumpReasonCount(ByRef keys() As String, ByRef counts() As Long, _
                            ByRef n As Long, ByVal reasonKey As String)
    Dim i As Long
    For i = 0 To n - 1
        If keys(i) = reasonKey Then
            counts(i) = counts(i) + 1
            Exit Sub
        End If
    Next i
    If n > UBound(keys) Then
        ReDim Preserve keys(0 To (UBound(keys) + 1) * 2 - 1)
        ReDim Preserve counts(0 To (UBound(counts) + 1) * 2 - 1)
    End If
    keys(n) = reasonKey
    counts(n) = 1
    n = n + 1
End Sub

' 失敗理由の内訳から特定コードの件数を取り出す(レビュー4-A)。
Private Function ReasonCountOf(ByRef keys() As String, ByRef counts() As Long, _
                               ByVal n As Long, ByVal reasonKey As String) As Long
    Dim i As Long
    For i = 0 To n - 1
        If keys(i) = reasonKey Then
            ReasonCountOf = counts(i)
            Exit Function
        End If
    Next i
End Function

Private Function BuildFilterPattern() As String
    Dim exts() As String: exts = Split(modExtractor.SupportedExts(), ",")
    Dim parts() As String: ReDim parts(LBound(exts) To UBound(exts))
    Dim i As Long
    For i = LBound(exts) To UBound(exts)
        parts(i) = "*." & Trim$(exts(i))
    Next i
    BuildFilterPattern = Join(parts, ";")
End Function
