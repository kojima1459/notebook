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
    TrimLog ws

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
    TrimLog ws
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
        Case "E0102"
            ' 2026-07-31 R11-D(監査3 H-4): Scripting.Dictionary(Windows
            ' Script Runtime)が端末ポリシーで塞がれている状態。検索・同期・
            ' 集計の中核がここに依存しており、利用者にはどれも「静かに
            ' 空振りする」形でしか現れない。宛先を明示する。
            FriendlyMessage = "この端末では検索機能に必要な部品" & _
                "(Windows Script Runtime)が利用できません。" & _
                "管理者へご連絡ください。"
        Case "E0201"
            FriendlyMessage = "AIリボンが見つからず、AIに問い合わせできませんでした。" & _
                "AIリボン入りのExcelで開き直すか、「診断」ボタンで状態を確認してください。"
        Case "E0202"
            ' 2026-08-04(R15-2d・実機第4報 RC8): E0202はAI通信の不調とは限らず、
            ' 実体は【実行中の処理と操作が重なったこと】による一時的な失敗の
            ' ことが多い(GSの完了待ち中に押されたボタンが入れ子で走る)。
            ' 「AIが混んでいる」とだけ言うと、待っても直らない別の話として
            ' 受け取られ、利用者は同じ操作を繰り返してしまう。
            FriendlyMessage = "AIとの通信が混み合っているようです(あなたの操作に問題はありません)。" & _
                "実行中の処理と操作が重なった可能性もあります。" & _
                "少し待ってから、もう一度お試しください。" & _
                "続くようなら診断結果を管理者へ連絡してください。"
        Case "E0203"
            FriendlyMessage = "文章をAIが読める形に変換する処理(埋め込み取得)に失敗しました。" & _
                "時間を置いてからもう一度お試しください。続くようなら診断結果を管理者へ連絡してください。"
        Case "E0204"
            FriendlyMessage = "本日のAI利用上限に達した可能性があります。" & _
                "今日はここまでにして、明日「同期」を押せば続きから再開できます。"
        Case "E0301"
            FriendlyMessage = "対応していない種類のファイルです。" & _
                "対応形式(txt/md/csv/pdf/docx/doc/xlsx等)のファイルをお使いください。"
        Case "E0302"
            FriendlyMessage = "ファイルを開けませんでした(他のアプリで開いている、または権限がない可能性があります)。" & _
                "ファイルを閉じてから、もう一度お試しください。" & _
                "PDFの場合は Ghostscript による読み取り(Word不要の経路)も試したうえでの結果です。" & _
                "このファイルと同じ場所に Ghostscript フォルダがあるかもご確認ください。"
        Case "E0303"
            FriendlyMessage = "このPDFは画像として保存されていて、そのままでは文字を読み取れませんでした。" & _
                "画像解析(config の feature_vision)を有効にし、Ghostscript を配置すると" & _
                "AIが1ページずつ読み取れます(設定は「43_画像PDFのOCR取込設定」を参照)。"
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
                "この処理は自動でスキップされましたが、通常は次回の「同期」で再試行されます。"
        Case "E0801"
            FriendlyMessage = "画面の組み立てに失敗しました(データは失われていませんのでご安心ください)。" & _
                "ブックを一度閉じて開き直せば、元どおり使えます。"
        Case "E0805"
            ' 2026-08-04(R15-3b): 読み取り専用で開かれている状態。取込も質問も
            ' 動くが、結果が1つも残らない(閉じた瞬間に消える)。表の未使用
            ' コードだったE0805をここで初めて使う。文言の実体は
            ' ReadOnlyWarnMsg(純関数)に置き、起動時の案内と共用する。
            FriendlyMessage = ReadOnlyWarnMsg()
        Case "E0806"
            ' 2026-07-31(R8b B16): R8 F13 で使い始めたコードが表に無く、
            ' 「予期しない問題が発生しました」という無意味な案内になっていた。
            ' 利用者が自力で直せる種類の問題ではないので、宛先を明示する。
            FriendlyMessage = "部門チャンネルの数が、このツールで扱える上限を超えています。" & _
                "上限を超えた部門の公式ナレッジは受け取れません。" & _
                "お手数ですが、このツールの管理担当者にご連絡ください(設定の変更が必要です)。"
        Case "E0807"
            ' 2026-07-31(R8b B2/B7): 正典の発行ロック(publish.lock)まわり。
            ' 発行者だけが見るコードなので、次に何をすればよいかを具体的に書く。
            FriendlyMessage = "正典の発行の見張り(publish.lock)を扱えませんでした。" & _
                "共有フォルダへの書き込み権限、または端末の時計のずれをご確認ください。" & _
                "発行が繰り返し止まる場合は、共有フォルダの channels\<部門名>\publish.lock を" & _
                "手動で削除すると再開できます。"
        Case "E0808"
            ' 2026-07-31(R11-A C3): 発行の予期しない中断と、旧版の退避失敗。
            ' 発行の入口(modPublishUI.OnPublish)は例外で抜けると完全に無言で
            ' 終わっていたため、押した人には「何も起きなかった」ようにしか
            ' 見えなかった。発行者が次に何を確かめればよいかまで書く。
            FriendlyMessage = "発行の処理中に予期しない問題が発生しました。" & _
                "もう一度お試しください。" & _
                "繰り返す場合は共有フォルダの接続と書込権限をご確認ください。"
        Case "E0901"
            FriendlyMessage = "診断で問題が見つかりました。" & _
                "診断レポートの指示に従って対応してください。"
        Case "E0904"
            ' 2026-08-05(R16-2a/波3の是正): 「作業用Excelを開く」が起動に失敗した
            ' 状態。端末ポリシーで WScript.Shell が塞がれていることが多く、利用者
            ' 自身で今すぐ打てる代替手段(タスクバーのExcelを Alt 押しながら
            ' クリック=新しいプロセスで起動)を必ず添える。取込中に出る文言なので
            ' 1文を短く保つ(StatusBar は長文を切り落とす)。
            FriendlyMessage = "作業用Excelを起動できませんでした。" & _
                "タスクバーのExcelをAltキーを押しながらクリックしても開けます。"
        Case "E0905"
            ' 2026-08-05(R18-5b): Hubフッターから社内ポータルを開けなかった
            ' 状態。FollowHyperlink と WScript.Shell の両方が塞がれている
            ' (端末ポリシー・既定ブラウザ未設定等)。呼び出し側がURLを
            ' クリップボードへ入れてから出すので、次にやることを1文で書く。
            ' MsgBoxへ渡る文言なので絵文字(非BMP)は使わない(検査14)。
            FriendlyMessage = "社内ポータルをこのパソコンから開けませんでした。" & _
                "アドレスをクリップボードにコピーしましたので、" & _
                "ブラウザのアドレス欄に貼り付けて(Ctrl+V)開いてください。"
        Case Else
            FriendlyMessage = "予期しない問題が発生しました(あなたの操作のせいではありません)。" & _
                "時間を置いてもう一度お試しいただき、続くようなら右上の ? から「ご意見・不具合報告」で教えてください。"
    End Select
End Function

' ============================================================================
' FriendlyFailMsg - 取込失敗を利用者へ伝える1文を選ぶ(2026-08-03 R13-3b)。
' ----------------------------------------------------------------------------
' 実機第2報 RC6: modUtil.DescribeComError が作った
' 「Wordを開いたままにして、もう一度お試しください」という【その場で打てる
' 次の一手】が、modShelf でE0302の汎用文言(Ghostscript前提)に上書きされ、
' 利用者には一度も届いていなかった。しかもその汎用文言は .docx の失敗にも
' 出るため、Wordの話をしているのにGhostscriptの確認を求めていた。
'
' 決め方は2段:
'   (1) errDetail に行動可能な案内が入っていれば、それを優先して採用する。
'       生のCOM説明文(英語や「型が一致しません」)は採用しない。目印は
'       DescribeComError が必ず書く導入句(「この端末では」「この環境では」)
'       と、462の案内だけが持つ「<アプリ名>を操作できませんでした」(R13-F11)。
'   (2) 無ければコード表の汎用文言。ただしE0302だけは拡張子で分岐し、
'       docx/doc には Ghostscript の話を一切しない。
' 純粋な文字列処理なので modTestsPure9 が分岐を固定する。
' ============================================================================
Public Function FriendlyFailMsg(ByVal errCode As String, ByVal errDetail As String, _
                                ByVal ext As String) As String
    Dim hint As String: hint = ActionableHint(errDetail)
    If LenB(hint) > 0 Then
        FriendlyFailMsg = hint & "(コード: " & errCode & ")"
        Exit Function
    End If

    If UCase$(Trim$(errCode)) = "E0302" Then
        Dim e As String: e = LCase$(Trim$(ext))
        If e = "docx" Or e = "doc" Then
            FriendlyFailMsg = "Wordの文書を読み取れませんでした" & _
                "(他のアプリで開いている、または権限がない可能性があります)。" & _
                "そのファイルを閉じてから、もう一度お試しください。" & _
                "それでも取り込めない場合は、Wordを起動したままにしてお試しください" & _
                "(この端末ではWordが開いていると取り込めることがあります)。" & _
                "(コード: " & errCode & ")"
            Exit Function
        End If
    End If

    FriendlyFailMsg = FriendlyMessage(errCode) & "(コード: " & errCode & ")"
End Function

' ----------------------------------------------------------------------------
' SharedReadFailMsg - 共有フォルダのファイルを読み取れなかったときの1文
'   (2026-08-03 R14-3b)。取込経路(modExtractorPdf)とOCR経路(optGsTxt)の
'   両方が同じ文を出す必要があるが、opt層はコアの基盤層しか参照できない(§7.7)
'   ため、文言の置き場としてはここが唯一の交点になる(憲章§4-5: 同じ文を
'   2箇所に書かない)。ActionableHint もこの文を目印として使う。
' ----------------------------------------------------------------------------
Public Function SharedReadFailMsg() As String
    SharedReadFailMsg = "共有のファイルを読み取れませんでした" & _
        "(ファイル名の特殊文字が原因の可能性)。ファイル名を変えて再度お試しください。"
End Function

' ----------------------------------------------------------------------------
' ReadOnlyWarnMsg - このブックが読み取り専用で開かれているときの1文
'   (2026-08-04 R15-3b・実機第4報 RC9)。
'   起動時の案内(modBoot)とコード表(E0805)の両方が同じ文を出す必要があり、
'   文言の一次情報は SharedReadFailMsg と同じ理由でここに1つだけ置く。
'   「保存されません」で終わらせず、原因として一番多い【別のExcelで同じ
'   ファイルを開いている】を名指しする(利用者がその場で確かめられる)。
'   純粋な文字列なので modTestsPure12 が文言を固定する。
' ----------------------------------------------------------------------------
Public Function ReadOnlyWarnMsg() As String
    ReadOnlyWarnMsg = "このファイルは読み取り専用で開かれています。" & _
        "取り込んだ資料は保存されません" & _
        "(既に別のExcelで開いていないか確認してください)。"
End Function

' ----------------------------------------------------------------------------
' SaveFailMsg - 中間保存に失敗したときの1文(2026-08-04 R15-3a)。
'   保存の失敗は利用者の操作で直せることが多い(別のExcelを閉じる・
'   共有ドライブへ再接続する)。ここで打てる一手は「終了ボタンから保存」
'   なので、それだけを言う。詳細は usage_log の save_fail 行に残す。
' ----------------------------------------------------------------------------
Public Function SaveFailMsg() As String
    SaveFailMsg = "保存に失敗しました。終了ボタンからの保存をお試しください。"
End Function

' ----------------------------------------------------------------------------
' CopyFailMsgOf - 一時コピーに失敗した【種類】ごとの、利用者向けの1文
'   (2026-08-03 R14-F3/F11)。
'   kind: "src_empty" / "locked" / "too_big" / "name_busy" / その他
'   R14-3b では全ての失敗に「ファイル名の特殊文字が原因の可能性」という
'   1文を出していた。0バイトの元ファイル・他アプリのロック・巨大ファイルまで
'   「名前を変えてください」と案内するのは、その場で打てない一手を言うのと
'   同じで、利用者は名前を変えて何度も試すことになる(§3-3)。
'   文言をここに置くのは SharedReadFailMsg と同じ理由(取込経路とOCR経路の
'   両方から出す必要があり、opt層はコア基盤層しか参照できない)。
'   ActionableHint はこの表の全文を目印にして、汎用文言への潰しを防ぐ。
' ----------------------------------------------------------------------------
Public Function CopyFailMsgOf(ByVal kind As String) As String
    Select Case LCase$(Trim$(kind))
        Case "src_empty"
            CopyFailMsgOf = "共有上のファイルが空(0バイト)です。" & _
                "元のファイルが正しく保存されているか確認してください。"
        Case "locked"
            CopyFailMsgOf = "他のアプリで開かれている可能性があります。" & _
                "ファイルを閉じてからもう一度お試しください。"
        Case "too_big"
            CopyFailMsgOf = "ファイルが大きすぎて読み込めませんでした。"
        Case "name_busy"
            CopyFailMsgOf = "一時ファイルが混み合っています。" & _
                "しばらくしてからもう一度お試しください。"
        Case Else
            CopyFailMsgOf = SharedReadFailMsg()
    End Select
End Function

' errDetail から「利用者がその場で打てる次の一手」を含む文だけを取り出す。
' 取り出せなければ ""(=汎用文言へ落とす)。
' errDetail は経路によって「[段階/開き方N] 本文 (詳細: …) [localcopy=ok]」や
' 「GS: … / Word: … / Acrobat: …」のように前後へ診断情報が付く。導入句から
' 始め、技術情報が始まる所で切ることで、利用者に見せてよい部分だけを残す。
Private Function ActionableHint(ByVal errDetail As String) As String
    If LenB(errDetail) = 0 Then Exit Function

    Dim p As Long: p = InStr(errDetail, "この端末では")
    Dim q As Long: q = InStr(errDetail, "この環境では")
    If p = 0 Or (q > 0 And q < p) Then p = q

    ' R13-F11: 462(相手のCOMサーバが居ない/セキュリティ製品に止められた)の
    ' 案内文だけは導入句を持たず「<アプリ名>を操作できませんでした。」で始まる。
    ' 導入句だけを目印にしていたため、実機で最も多いブロックの案内
    ' (手で起動できるか確かめる/セキュリティ製品の可能性)が汎用文言に
    ' 潰されていた。アプリ名の先頭まで戻って、そこを開始点として拾う。
    Dim r As Long: r = HintStartOfOperateFail(errDetail)
    If p = 0 Or (r > 0 And r < p) Then p = r

    ' R14-3b: 共有読みの失敗(実機第3報 RC3)は導入句を持たない独自の1文。
    ' これを拾えないと、E0302の汎用文言(Ghostscriptの話)に潰されて
    ' 「ファイル名を変えてみる」という唯一の一手が利用者へ届かない。
    ' R14-F3/F11: 失敗の種類ごとに文言が分かれたので、表の全文を目印にする
    ' (1つでも漏れると、その種類だけ汎用文言に潰されて案内が消える)。
    Dim s As Long: s = CopyFailHintStart(errDetail)
    If p = 0 Or (s > 0 And s < p) Then p = s
    If p = 0 Then Exit Function

    Dim d As String: d = Mid$(errDetail, p)
    d = CutBefore(d, "(詳細:")
    d = CutBefore(d, " [")
    d = CutBefore(d, " / ")
    ActionableHint = Trim$(d)
End Function

' 一時コピー失敗の案内文(CopyFailMsgOf の表)が errDetail のどこから
' 始まるかを返す(見つからなければ 0)。複数見つかったら最も手前を採る。
Private Function CopyFailHintStart(ByVal errDetail As String) As Long
    Dim kinds As Variant
    kinds = Array("src_empty", "locked", "too_big", "name_busy", "")
    Dim i As Long
    For i = LBound(kinds) To UBound(kinds)
        Dim p As Long: p = InStr(errDetail, CopyFailMsgOf(CStr(kinds(i))))
        If p > 0 Then
            If CopyFailHintStart = 0 Or p < CopyFailHintStart Then CopyFailHintStart = p
        End If
    Next i
End Function

' 「<アプリ名>を操作できませんでした」の【アプリ名の先頭】の位置を返す
' (見つからなければ 0)。アプリ名は "Word"/"Excel"/"Acrobat"/"Office" の
' ような半角英字なので、その並びだけを手前へ遡る(R13-F11)。
' 遡りすぎて診断情報を巻き込まないよう、英字以外に当たったら即やめる。
Private Function HintStartOfOperateFail(ByVal errDetail As String) As Long
    Dim r As Long: r = InStr(errDetail, "を操作できませんでした")
    If r = 0 Then Exit Function

    Dim i As Long: i = r
    Do While i > 1
        Dim c As Long: c = AscW(Mid$(errDetail, i - 1, 1))
        If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Then
            i = i - 1
        Else
            Exit Do
        End If
    Loop
    If i = r Then Exit Function      ' アプリ名が無い=別文脈。拾わない
    HintStartOfOperateFail = i
End Function

' marker が見つかったらその手前まで(見つからなければそのまま)。
Private Function CutBefore(ByVal s As String, ByVal marker As String) As String
    Dim p As Long: p = InStr(s, marker)
    If p > 0 Then
        CutBefore = Left$(s, p - 1)
    Else
        CutBefore = s
    End If
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

' ----------------------------------------------------------------------------
' TrimLog - ログシートの古い行を落として上限行数に収める。
'
' 2026-07-28(レビュー I-1): err_log / usage_log は無限追記だった。
' このブックは自己インストーラが開くたび ThisWorkbook.Save するので、
' 行が増えるほど毎起動の保存が重くなり、最後はファイルサイズそのものが
' 配布の邪魔になる。チャット履歴(modChatLog)は既に100件でローテして
' いたのに、ログ側だけ野放しだった。
'
' 新しい行は末尾へ積む(NextRow)。したがって落とすのは【上側=古い方】で、
' 直近の記録は必ず残す。ここを逆にすると、障害が起きた直後に最も見たい
' 行から消えるという最悪の挙動になる。
' 毎回 EntireRow.Delete すると重いので、上限を1割超えてから
' 上限ちょうどまで一気に削る(削る頻度を下げる)。
' 保持件数は config log_max_rows(既定2000・0以下でローテ無効)。
' ----------------------------------------------------------------------------
Private Sub TrimLog(ByVal ws As Worksheet)
    On Error Resume Next
    Dim maxRows As Long
    maxRows = modConfig.GetLong("log_max_rows", 2000)
    If maxRows <= 0 Then Exit Sub

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim dataRows As Long: dataRows = lastRow - 1
    If dataRows <= maxRows + (maxRows \ 10) Then Exit Sub

    ' 残すのは末尾 maxRows 行。ヘッダ(1行目)の直下から、余った分だけ消す。
    Dim dropCount As Long: dropCount = dataRows - maxRows
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + dropCount, 1)).EntireRow.Delete
    On Error GoTo 0
End Sub
