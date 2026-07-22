Attribute VB_Name = "modLog"
Option Explicit

' ============================================================================
' modLog - エラーログ/利用ログの記録と、ユーザー向けメッセージの一元管理
' ----------------------------------------------------------------------------
' 役割:
'   err_log/usage_logシートへの追記と、MASTER_SPEC §6のエラーコード表に
'   基づく「何が起きたか。どうすればよいか。」形式のメッセージ生成を担当する。
'
' 設計判断:
'   ・R5(エラーを握りつぶさない)の唯一の例外として、ログ書き込み自体の
'     失敗はアプリの処理を止めないために黙って捨てる(「ログで死なない」)。
'     その場合もDebug.Printにだけは残す(実行時にVBEを見れば追える)。
'   ・err_log/usage_logシートは初回の書き込み時にその場で生成する
'     (ヘッダ付き、非表示)。ビルド時に必ず作られる想定だが、万一シートが
'     無い状態でも動作を止めないための保険。
'   ・FriendlyMessageは§6の全20コードを網羅し、必ず2文構成
'     (「何が起きたか。どうすればよいか。」)で返す。末尾の呼び出し側
'     (ShowError)がコードを追記して「(コード: E0xxx)」を表示することで、
'     保守者がerr_logと突合できるようにする。
'   ・modConfig(debug_mode判定)にのみ依存する。循環参照を避けるため
'     modConfigからmodLogを呼び返すことはしない(modConfig側のコメント参照)。
' ============================================================================

' err_number/http_status(共にOptional・既定0=情報なし): 実機環境の壁
' (社内プロキシのエラー407、共有フォルダ/P2Pのエラー52・70等)を、利用者が
' コードを読まずにそのまま開発者へ伝えられるようにするための生の診断情報。
' 追加は末尾Optionalのため、この2つを渡さない既存の全呼び出し元は無改修で
' 動く(後方互換)。err_number=Err.Number(VBA実行時エラー番号)、
' http_status=HTTPレスポンスコード(direct埋め込み等、生のHTTP応答がある
' 場合のみ)。
Public Sub LogError(ByVal code As String, ByVal context As String, ByVal detail As String, _
                    Optional ByVal err_number As Long = 0, Optional ByVal http_status As Long = 0)
    On Error GoTo Fail
    Dim ws As Worksheet: Set ws = EnsureLogSheet(modAppDef.SH_ERRLOG, ErrLogHeader())
    If ws Is Nothing Then GoTo Fail

    Dim r As Long: r = NextRow(ws)
    ws.Cells(r, 1).Value = modUtil.NowStamp()
    ws.Cells(r, 2).Value = code
    ws.Cells(r, 3).Value = modUtil.SafeLeft(context, 255)
    ws.Cells(r, 4).Value = modUtil.SafeLeft(detail, 2000)
    ' 2026-07-21: version列にビルド識別子(日時+gitコミット)を併記する。
    ' 「今テストしているファイルは本当に最新ビルドか」を実機報告から即座に
    ' 判別できるようにするため(build/build_mybookshelf.py compute_build_stamp)。
    Dim buildStamp As String
    On Error Resume Next
    buildStamp = modConfig.GetString("build_stamp", "")
    On Error GoTo 0
    ws.Cells(r, 5).Value = modAppDef.APP_VERSION & IIf(LenB(buildStamp) > 0, " " & buildStamp, "")
    ws.Cells(r, 6).Value = err_number
    ws.Cells(r, 7).Value = http_status

    If modConfig.GetBool("debug_mode", False) Then
        Debug.Print "[modLog.LogError] " & code & " " & context & " : " & detail & _
            " (err#" & err_number & " http=" & http_status & ")"
    End If
    Exit Sub
Fail:
    ' ログ書き込み自体の失敗はアプリを止めない(R5の唯一の例外運用)。
    Debug.Print "[modLog.LogError:書込失敗] " & code & " " & context & " : " & detail
End Sub

Public Sub LogUsage(ByVal event_name As String, ByVal mode As String, ByVal detail As String, _
                    Optional ByVal latency_ms As Long = 0, Optional ByVal hit_count As Long = 0)
    On Error GoTo Fail
    Dim ws As Worksheet: Set ws = EnsureLogSheet(modAppDef.SH_USAGE, UsageLogHeader())
    If ws Is Nothing Then GoTo Fail

    Dim r As Long: r = NextRow(ws)
    ws.Cells(r, 1).Value = modUtil.NowStamp()
    ws.Cells(r, 2).Value = event_name
    ws.Cells(r, 3).Value = mode
    ws.Cells(r, 4).Value = modUtil.SafeLeft(detail, 2000)
    ws.Cells(r, 5).Value = latency_ms
    ws.Cells(r, 6).Value = hit_count
    Exit Sub
Fail:
    Debug.Print "[modLog.LogUsage:書込失敗] " & event_name & " " & mode & " : " & detail
End Sub

' MASTER_SPEC §6のエラーコード表に対応する、ユーザー向けの2文構成メッセージ。
' 「何が起きたか。どうすればよいか。」の順で必ず2文になるようにしている。
Public Function FriendlyMessage(ByVal code As String) As String
    Select Case UCase$(Trim$(code))
        Case "E0101"
            FriendlyMessage = "設定ファイル(configシート)に必要な項目が見つかりません。" & _
                "配布元にこのファイルの再入手を依頼してください。"
        Case "E0201"
            FriendlyMessage = "AIリボンが見つからず、AIに問い合わせできませんでした。" & _
                "AIリボン入りのExcelで開き直すか、" & ChrW(&HD83E) & ChrW(&HDE7A) & "診断ボタンで状態を確認してください。"
        Case "E0202"
            FriendlyMessage = "AIとの通信が混み合っているようです(あなたの操作に問題はありません)。" & _
                "数十秒おいてもう一度お試しください。続くようなら診断結果を管理者へ連絡してください。"
        Case "E0203"
            FriendlyMessage = "文章をAIが読める形に変換する処理(埋め込み取得)に失敗しました。" & _
                "時間を置いてからもう一度お試しください。続くようなら診断結果を管理者へ連絡してください。"
        Case "E0204"
            FriendlyMessage = "本日のAI利用上限に達した可能性があります。" & _
                "今日はここまでにして、明日「" & ChrW(&HD83D) & ChrW(&HDD04) & "同期」を押せば続きから再開できます。"
        Case "E0301"
            FriendlyMessage = "対応していない種類のファイルです。" & _
                "対応形式(txt/md/csv/pdf/docx/doc/xlsx等)のファイルをお使いください。"
        Case "E0302"
            FriendlyMessage = "ファイルを開けませんでした(他のアプリで開いている、または権限がない可能性があります)。" & _
                "ファイルを閉じてから、もう一度お試しください。"
        Case "E0303"
            FriendlyMessage = "このPDFは画像として保存されていて、文字を読み取れませんでした。" & _
                "画像PDFの取込には対応していません(将来のVision対応をお待ちください)。"
        Case "E0304"
            FriendlyMessage = "ファイルは開けましたが、中身の文章が空でした。" & _
                "中身が空か、保護されているファイルの可能性があります。別の資料でお試しください。"
        Case "E0401"
            FriendlyMessage = "文章を区切って保存する処理で、保存できた内容が0件でした。" & _
                "文字数が少なすぎる資料の可能性があります。別の資料でお試しください。"
        Case "E0501"
            FriendlyMessage = "本棚に入れられる資料の上限を超えています。" & _
                "使っていない資料を削除してから、もう一度追加してください。" & vbLf & _
                "(上限そのものを増やすこともできます: configシートの shelf_max_chunks の数字を大きくしてください)"
        Case "E0502"
            ' E0502はMsgBox専用コード(shelf_folder関連)のため絵文字を使わない
            ' (MsgBoxでの絵文字表示問題はShowErrorのコメント参照)。
            FriendlyMessage = "同期するフォルダが見つかりませんでした。" & _
                "「フォルダを選ぶ」からフォルダを選び直してください。"
        Case "E0503"
            FriendlyMessage = "今、別の取込処理が実行中です。" & _
                "処理が完了するまで、少々お待ちください。"
        Case "E0504"
            FriendlyMessage = "同じ名前の資料が、別の場所からすでに登録されています。" & _
                "ファイル名を変えて追加するか、先に元の資料を削除してから追加し直してください。"
        Case "E0601"
            FriendlyMessage = "本棚の中に手がかりが見つかりませんでした。" & _
                "資料が本棚に入っているか確認するか、資料を追加してから試してください。"
        Case "E0602"
            FriendlyMessage = "うまく回答をまとめられませんでした(こちら側の処理の都合です)。" & _
                "お手数ですが、もう一度送信してください。続く場合は質問の言い回しを少し変えると通ることがあります。"
        Case "E0701"
            FriendlyMessage = "パックの形式が正しくないか、バージョンが合っていません。" & _
                "相手にパックの再書き出しを依頼してください。"
        Case "E0702"
            FriendlyMessage = "パックのベクトルの次元数が一致しませんでした。" & _
                "AIリボンのバージョンが違う可能性があります。管理者へ連絡してください。"
        Case "E0703"
            FriendlyMessage = "書き出す内容に個人情報らしきものが見つかりました。" & _
                "内容を確認してから、書き出しを続けるかどうか判断してください。"
        Case "E0705"
            FriendlyMessage = "共有フォルダへの書き込み/読み込みに失敗しました" & _
                "(ネットワークの瞬断・アクセス権限・セキュリティソフトのブロック等の可能性があります)。" & _
                "この処理は自動でスキップされましたが、通常は次回の" & ChrW(&HD83D) & ChrW(&HDD04) & "同期で再試行されます。"
        Case "E0801"
            FriendlyMessage = "画面の組み立てに失敗しました(データは失われていませんのでご安心ください)。" & _
                "ブックを一度閉じて開き直せば、元どおり使えます。"
        Case "E0901"
            FriendlyMessage = "診断で問題が見つかりました。" & _
                "診断レポートの指示に従って対応してください。"
        Case Else
            FriendlyMessage = "予期しない問題が発生しました(あなたの操作のせいではありません)。" & _
                "時間を置いてもう一度お試しいただき、続くようなら右上の ? から「ご意見・不具合報告」で教えてください。"
    End Select
End Function

' 実機報告(2026-07-21): ここのChrW絵文字がMsgBox上で「??」表示になっていた。
' Nexus画面のShape文字(TextFrame2)では正しく描画されるが、ネイティブMsgBox
' (Win32 MessageBox)は既定フォントの絵文字グリフ対応が弱く、Excel側の描画
' 経路とは別問題。MsgBoxはここだけ絵文字を使わない。
Public Sub ShowError(ByVal code As String, ByVal context As String, ByVal detail As String)
    LogError code, context, detail
    MsgBox FriendlyMessage(code) & vbLf & "(コード: " & code & ")" & vbLf & vbLf & _
        "「診断」ボタン" & ChrW(&H2192) & "「直近のエラーをコピー」で、詳しい情報をそのまま担当者に送れます。", _
        vbExclamation, modAppDef.APP_NAME
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------
Private Function ErrLogHeader() As Variant
    ErrLogHeader = Array("timestamp", "code", "context", "detail", "version", "err_number", "http_status")
End Function

Private Function UsageLogHeader() As Variant
    UsageLogHeader = Array("timestamp", "event", "mode", "detail", "latency_ms", "hit_count")
End Function

Private Function EnsureLogSheet(ByVal sheetName As String, ByVal header As Variant) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
        Dim i As Long
        For i = LBound(header) To UBound(header)
            ws.Cells(1, i + 1).Value = header(i)
        Next i
        HideSheetQuietly ws
        On Error GoTo 0
    End If
    Set EnsureLogSheet = ws
    Exit Function
Fail:
    Set EnsureLogSheet = Nothing
End Function

Private Sub HideSheetQuietly(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Visible = 0   ' xlSheetHidden (名前付き定数への依存を避けるV2の慣習に合わせる)
    On Error GoTo 0
End Sub

Private Function NextRow(ByVal ws As Worksheet) As Long
    Dim r As Long: r = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If r < 1 Then r = 1
    NextRow = r + 1
End Function
