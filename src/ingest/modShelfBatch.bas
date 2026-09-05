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
' なぜ分けたか(2026-07-31 R10d): modShelf が30,000字上限まで残り440字になり
'   バグ修正1件すら入らなくなった(R10cのH2=バッチ全体の再入ガードが容量不足で
'   保留になったのが直接のきっかけ)。AddFilesViaDialog / AddFilesResult と
'   その補助(集計文字列・失敗理由カウンタ・FileDialogのフィルタ生成)は
'   【この一式の中だけで完結する】凝集した塊で、外から使われるのは2本だけ。
'
' 依存: modShelf.IngestFile / TotalChunks(取込本体と件数)、
'   modExtractor.SupportedExts(FileDialogのフィルタ生成)、表示は
'   modUIMain.ShowProgress/HideProgress・modSkin.ShowToast・
'   modUIShelf.RenderShelf(いずれもR1例外の通知コールバック)。
'   opt名は書かず modFeatures.InvokeFeature 経由(R2)。
' ============================================================================

' R10c(H2): バッチ取込【全体】を覆う再入ガード。modShelf の mIngesting は
' IngestFile 1件ぶんしか覆っておらず、ファイル間の隙間(バナー更新やトーストの
' DoEvents 中)はガードが下りていなかった。そこで発火したクリックは
' modUiLock.BlockIfIngesting を素通りし、取込ループの途中から画面遷移が入れ子で
' 走り出す(実機「取込中にボタンを押すと応答なし」の残り火)。modShelf.IsBusy が
' このフラグも OR で見る。焼き付いても全ボタンが永久に死なないよう mIngesting と
' 同じ作法(GUARD_EXPIRY_MIN=30分)で自動失効させる。
Private mBatchIngesting As Boolean
Private mBatchSince As Date
Private Const GUARD_EXPIRY_MIN As Long = 30

' R15-1a(実機第4報 RC5): 4本の再入ガード(modShelf/modShelfBatch/modEmbed/
' modShelfSync)が共有する「最後に生きていた時刻」。取込は実機で85〜127分かかる
' のに失効判定が「開始から30分」で、途中で全ボタンが解放され二重取込が構造的に
' 可能だった(憲章§3-5)。ビートから数えれば動いている間は何時間でも守り、
' 止まってから30分で自己回復する(§4-5: 答えは1箇所)。
Private mLastBeat As Date

' R15-3a: 中間保存の失敗を利用者へ告げたかどうか(1セッション1回だけ告げる)。
Private mSaveFailToasted As Boolean

' R15-FixB(FB-7・レビューA-L1): 読み取り専用による save_fail を usage_log へ
' 書いたか(1セッション1回だけ)。読み取り専用は【セッションを通じて変わらない
' 1つの事実】で、毎回書くと数百行が同じ内容で並び、usage_log の行数ローテで
' 本当に調べたい行(取込・中断・再開)が先に流れて消える。
Private Const RO_FAIL_DESC As String = "ReadOnly(読み取り専用で開かれているため保存しません)"
Private mReadOnlyLogged As Boolean

' R15-FixA(FA-2): 前回の中間保存を試みた時刻。頁OCRの控え保存から呼ばれる
' 経路だけスロットル(120秒)をかけるための唯一の材料。0=まだ一度も無い。
Private mLastSaveAt As Date

' R15-6a(実機第4報 RC3): 中断の申し出。進捗バナー脇の中断ボタンが立て、
' optOcrPage(頁/バッチ境界)と下の取込ループ(ファイル境界)が読む。
' Esc/err18 方式を採らないのは、modGateway.TryRibbonRun/CallLLM の全捕捉
' ハンドラが Err 18 を握りつぶす構造になっており、共有コアを書き換える
' 爆風半径が大きすぎるため(仕様書「記録のみ」の裁定)。
Private mCancelRequested As Boolean
Private mCancelAt As Date

' ----------------------------------------------------------------------------
' TouchBusy - 「まだ生きている」印を1つ進める(R15-1b)。
'   呼び出し点は modUIMain.SetStage と本モジュールの StageBanner の2箇所だけ
'   (長時間ループは必ずどちらかを通る)。副作用はモジュール変数の代入のみで、
'   例外を出さない=どこから呼ばれても既存の処理を壊さない。
' ----------------------------------------------------------------------------
Public Sub TouchBusy()
    mLastBeat = Now
End Sub

' LastBeat - 最終ビート(まだ1度も打たれていなければ0)。
Public Function LastBeat() As Date
    LastBeat = mLastBeat
End Function

' ----------------------------------------------------------------------------
' OnCancelIngest - 進捗バナー脇の「中断」ボタン(2026-08-04 R15-6a)。
' ----------------------------------------------------------------------------
' やることは3つだけ: 印を立てる / まだ生きていることを知らせる(TouchBusy) /
' 何が起きるのかを1行で伝える。確認ダイアログは出さない。
'   ・取込中はモーダルを開くこと自体が事故の種で(実機第4報 RC8 の入れ子実行)、
'     しかも「本当に中断しますか?」に答える前に次の頁が始まってしまう。
'   ・押した瞬間には止まらない。同期の Application.Run(ChatGPTV)は VBA から
'     中断できないので、止まるのは今の頁が終わってから。だから「中断して
'     います…」と現在形で言い、いつ止まるのかまで書く(憲章§3-2)。
' 2回押されても、取込していないときに押されても壊れない(印を立て直すだけ)。
' 再入の関所(BlockIfIngesting)は【付けない】。取込中でも必ず動かなければ
' 意味が無いハンドラで、vba_lint の ONACTION_GUARD_ALLOWLIST に理由つきで
' 登録してある。
' 2026-08-04(R15-FixA FA-8): (a) SetStage → ShowProgress。SetStage の出力先は
'   Nexus画面では実質不可視で、押した人には何も起きていないように見えていた。
'   (b) TouchBusy を外す。ビートは【実作業だけ】が打つ一本の線にする。
'   文言から「頁」を外した: 中断が効くのは頁境界だけではない(画像化の待ち・
'   ベクトル化・ファイル境界でも効く)。
Public Sub OnCancelIngest()
    On Error Resume Next
    mCancelRequested = True
    mCancelAt = Now
    modUIMain.ShowProgress "中断しています… 現在の処理が終わり次第停止します", True
    modLog.LogUsage "ingest_cancel", "", "利用者が中断ボタンを押しました"
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' CancelRequested - 中断が申し出られているか(R15-6a)。
'   印は取込の外へ持ち越さない。AddFilesResult の入口と全ての出口で必ず
'   下ろすが、それだけだと「手動同期の途中で中断した」印が次の取込まで
'   残り得る(modShelfSync は本ラウンド接触禁止)ので、GUARD_EXPIRY_MIN と
'   同じ30分で自然に失効させる。取込は押された直後に読むので、この失効で
'   本来効くべき中断が取りこぼされることはない。
' ----------------------------------------------------------------------------
Public Function CancelRequested() As Boolean
    If Not mCancelRequested Then Exit Function
    On Error Resume Next
    If DateDiff("n", mCancelAt, Now) > GUARD_EXPIRY_MIN Then
        mCancelRequested = False
    Else
        CancelRequested = True
    End If
    On Error GoTo 0
End Function

' ResetCancel - 中断の印を下ろす(取込の入口と出口で必ず呼ぶ)。
Public Sub ResetCancel()
    mCancelRequested = False
End Sub

' ----------------------------------------------------------------------------
' GuardExpiredNow - 「開始時刻 startAt のガードが今もう失効しているか」。
'   判定式そのものは純関数 modUtilText.GuardExpired(テスト済み)に持たせ、
'   ここは現在時刻と共有ビートを差し込むだけの薄い口にする。4本のガードが
'   同じ関数を呼ぶことで、判定が箇所ごとにズレることを構造的に防ぐ。
' ----------------------------------------------------------------------------
'   absLimitMin:=0(2026-08-04 R15-FixA FA-9): 開始からの絶対上限(480分)は
'   外す。ビートを打つのは実作業だけになった(FA-8)ので自己延命ループは起き
'   ない。一方で8時間超の取込は正当に起こり得(254頁の資料を何件もまとめて
'   選ぶ)、そこでガードが解けると動いている取込の上に二重取込の入口ができる。
'   無音30分の自己回復だけで足りる(§3-5)。
Public Function GuardExpiredNow(ByVal startAt As Date, ByVal limitMin As Long) As Boolean
    GuardExpiredNow = modUtilText.GuardExpired(startAt, mLastBeat, Now, limitMin, 0)
End Function

' ----------------------------------------------------------------------------
' SaveCheckpoint - 中間保存(R15-3a・実機第4報 RC9)。
' ----------------------------------------------------------------------------
' 保存の機会は「自己インストーラのopen時Save」と「終了ボタン」だけで、取り込んだ
' チャンクとログは数時間メモリ上だけに置かれ、途中でExcelが落ちれば127分かけた
' 取込が丸ごと消えた(憲章§3-5)。1件ごと・同期末尾ごとに保存し失う範囲を1件へ
' 抑える。保存は失敗し得る(共有ロック・読み取り専用・容量)ので失敗は必ず1行
' 残し(save_fail)、利用者には次の一手だけ伝える。成功時は無言(毎回トーストは
' 1.1秒×件数の待ち)。R18-2c: 呼び出し元は modShelf.IngestFile の Finish
' (成功時のみ・silent含む)へ一本化=どの入口から取り込んでも保存される。
'
' changedN: 呼び出し元が数えた「今回変わった件数」。0なら保存しない(何も
'   変わっていないのに5分ごとの自動同期で数百KBを書き戻すと遅く、共有
'   ドライブでは競合の種)。省略時は1=「1件入った直後」。
'
' 読み取り専用のときは保存を試さない: 必ず失敗するうえ、Excelが「名前を付けて
' 保存」を出す余地を作らない(取込ループの途中でモーダルが立つと全部止まる)。
' 起動時に modDiag.WarnIfReadOnly が理由を伝えるので、ここは記録だけ。
' トーストは1セッション1回だけ: 20件取り込めば20回失敗し、そのたびに1.1秒の
' トーストを出すと「押すほど固まる」を自分で作る(R10c M3と同じ轍)。
'
' throttleSec(2026-08-04 R15-FixA FA-2): この秒数以内に前回保存していたら黙って
'   戻る。頁OCRの控え保存と silent な同期の1件ごと(R18H FA-7)からは120を渡す:
'   数十秒に1回書き戻すとEDRのスキャンが取込より重くなる端末がある。「失うのは
'   最大2分ぶん」なら127分が丸ごと消える(RC9)壊れ方は起きない。既定0=毎回。
Public Sub SaveCheckpoint(Optional ByVal changedN As Long = 1, _
                          Optional ByVal throttleSec As Long = 0)
    If changedN < 1 Then Exit Sub
    If throttleSec > 0 And mLastSaveAt > 0 Then
        On Error Resume Next
        Dim sinceSec As Long: sinceSec = 0
        sinceSec = DateDiff("s", mLastSaveAt, Now)
        On Error GoTo 0
        ' 時計が巻き戻った端末(負)は「間隔不明」なので保存する側に倒す。
        If sinceSec >= 0 And sinceSec < throttleSec Then Exit Sub
    End If
    mLastSaveAt = Now

    On Error Resume Next
    Dim failNum As Long: failNum = 0
    Dim failDesc As String: failDesc = ""
    Dim isReadOnly As Boolean: isReadOnly = False
    isReadOnly = ThisWorkbook.ReadOnly
    Err.Clear
    If isReadOnly Then
        failNum = -1
        failDesc = RO_FAIL_DESC
    Else
        ' R18-2d→R18H FA-3(A-M2): マークは Save の【直前】に書く。後に書くと
        ' ui_state のマークだけが必ず1世代古いままディスクへ残り、次の起動で
        ' 「前回127件→今回0件」の虚偽警告の材料になった。先に書けばマークと
        ' 本体が同じ世代で保存される(保存に失敗しても、メモリ上のマークが
        ' 最新の行数を指すだけで害は無い)。
        modIntegrity.RecordSaveMark
        ThisWorkbook.Save
        failNum = Err.Number
        failDesc = Err.Description
        Err.Clear
    End If

    If failNum <> 0 Then
        ' R15-FixB(FB-7): 読み取り専用だけは1セッション1行に畳む(理由は上の
        ' mReadOnlyLogged の注記)。他の失敗(共有ロック・容量)は毎回残す:
        ' そちらは回ごとに起きたり起きなかったりする=回数が手がかりになる。
        Dim skipLog As Boolean: skipLog = False
        If isReadOnly Then
            skipLog = mReadOnlyLogged
            mReadOnlyLogged = True
        End If
        If Not skipLog Then _
            modLog.LogUsage "save_fail", "", _
                "中間保存に失敗 err#" & failNum & ": " & modUtil.SafeLeft(failDesc, 300)
        If Not mSaveFailToasted Then
            mSaveFailToasted = True
            modSkin.ShowToast modLog.SaveFailMsg(), "error"
        End If
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' IsBatchBusy - バッチ取込の最中か(modShelf.IsBusy から OR で参照される)。
'   期限切れのガードは busy とみなさない(modShelf.IsBusy と同じ考え方)。
' ----------------------------------------------------------------------------
Public Function IsBatchBusy() As Boolean
    If Not mBatchIngesting Then Exit Function
    On Error Resume Next
    IsBatchBusy = Not GuardExpiredNow(mBatchSince, GUARD_EXPIRY_MIN)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' StageBanner - 取込の【段階】を進捗バナー(nx_progress)へ流す薄いAPI
'   (2026-08-03 R13-4a。最初の呼び手はR13-1dのGS本文抽出の待ちループ)。
'
' 規約はただ1つ: 【進捗バナーが今すでに出ているときだけ】更新する。
'   ・手動の「資料を追加」や手動同期では、ループ側が先に ShowProgress を
'     呼んでバナーを出しているので、ここからの更新は素直に届く。
'   ・自動同期(silent)や起動時の裏処理ではバナーが出ていない。そこで
'     ShowProgress を呼ぶと、利用者が何もしていないのに画面へ黒い帯が
'     突然生えることになる(憲章§3-4「利用者を不安にさせない」に反する)。
'     出ていないなら黙る、というこの1点だけで silent は silent のままになる。
'
' なぜフラグを配線しないのか: 呼び出し元(取込→抽出→GS待ち)は層をまたぐ長い
' 経路で、silent を引数で運ぶと modShelfSync まで波及する。可視判定は「今の
' シートに nx_progress のShapeがあるか」で行う。表示の実体(modProgressBar)が
' Shapeを作って消すので、これが唯一の事実(modUIMain は状態を持っていない)。
' 表示系の失敗が取込を壊してはならない(憲章§4-4)ので全体をOERNで包む。
' ----------------------------------------------------------------------------
Public Sub StageBanner(ByVal text As String)
    ' R15-1b: 段階が1つ進んだ=取込は生きている。バナーが見えていない
    ' (silent同期・別ブックを見ている)ときも印だけは必ず進める。
    TouchBusy
    On Error Resume Next
    If LenB(text) = 0 Then Exit Sub
    If Not IsProgressBannerVisible() Then Exit Sub
    ShowIngestBanner text
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ShowIngestBanner - 【取込経路の】進捗バナーを出す(2026-08-04 R15-FixA FA-6)。
' ----------------------------------------------------------------------------
' modUIMain.ShowProgress と同じ2つのこと(状態行/StatusBarへの実況+画面上の
' バナー)をするが、バナーには「中断」ボタンを添える。
' なぜ分けるのか: 中断ボタンは PaintProgress が【常に】描いていたため、
' 質問の準備・Q&Aの読込・ナレッジ登録など、中断の仕組みが一切無い処理の
' バナーにも生えていた(A-M4/B-H4)。押せば取込用の印が立つが何も止まらない=
' 「押しても何も起きないボタン」で、故障と区別が付かない。取込の入口3箇所
' (この下の AddFilesResult / StageBanner / modShelfSync のファイルループ)
' だけがここを通り、他は従来どおり modUIMain.ShowProgress を使う。
' modUIMain 側で分岐しないのは、あちらが残り206字で引数1つ入らないため。
Public Sub ShowIngestBanner(ByVal text As String)
    On Error Resume Next
    modUIMain.SetStage text          ' ShowProgress の前半(既存チャネルへの実況)
    modProgressBar.PaintProgress text, True ' 後半(バナー)。中断ボタンつき
    On Error GoTo 0
End Sub

' 進捗バナーが今この画面に出ているか(StageBanner専用の可視判定)。
' 別ブックを見ている間は「出ていない」とみなす(PaintProgressの誤爆ガードと
' 同じ考え方)。Shapesの取得はShapeが無いとエラーになるので、OERNで受けて
' Nothing のままかどうかで判定する。
Private Function IsProgressBannerVisible() As Boolean
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Function
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Function
    Dim shp As Object
    Set shp = ws.Shapes("nx_progress")
    IsProgressBannerVisible = Not (shp Is Nothing)
    On Error GoTo 0
End Function

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
'   どれにも入らない「キャンセル」は ok=0;ng=0;capped=0;chunks=0;reasons= 。
'   別の取込が動いていて受け付けなかったときは reasons=busy:1(R11-H Med1)。
'
Public Function AddFilesResult(Optional ByVal showMsgBox As Boolean = True) As String
    ' R11-C(H-6): 入口での再入ガード。DoEvents経由で同時に2本目が走ることは
    ' 構造的に無い(FileDialog.Showはモーダルでイベントループを回さない)が、
    ' 前回の呼び出しがAddFailed経路で解除し損ねた場合の保険としてキャンセル
    ' 相当を返す(mBatchIngestingの退避/復元はせず、入口で弾くだけで足りる)。
    If IsBatchBusy() Then
        ' 2026-07-31(R11-H Med1): 従来はキャンセルと同じ戻り値だったため、
        ' 呼び出し元(modApp.OnAddDocs)は「押し間違い」とみなして何も言わず、
        ' 利用者には【押しても何も起きない】としか見えなかった(憲章§3-1)。
        ' reasons=busy を立てて、呼び出し元が理由を説明できるようにする。
        modLog.LogUsage "add_files_busy", "", _
            "取込中に「資料を追加」が押されたため受け付けませんでした"
        AddFilesResult = "ok=0;ng=0;capped=0;chunks=0;reasons=busy:1"
        Exit Function
    End If
    modShelfScan.EnsureManifestTextFormat
    ' 集計値の宣言はハンドラより前に置く(途中で落ちても集計を返すため)。
    Dim okCount As Long: okCount = 0
    Dim ngCount As Long: ngCount = 0
    ' R13-3a: status="partial" で入った件数(打ち切り・埋め込み未了・薄い抽出)。
    ' 「成功」に混ぜたまま黙っていると、本文が薄いまま入った資料に誰も
    ' 気付けない(実機第2報 RC1)。完了トーストで件数だけは必ず言う。
    Dim warnN As Long: warnN = 0
    Dim cappedN As Long: cappedN = 0
    Dim chunksAdded As Long: chunksAdded = 0
    ' R15-6c: 中断のため手を付けずに見送ったファイル数。黙って減らすと
    ' 「選んだのに入っていない」資料が理由も無く生まれる(憲章§4-1)。
    Dim cancelSkipN As Long: cancelSkipN = 0
    ' R15-FixA(FA-7): 中断が【あったか】。見送り件数(cancelSkipN)とは別物で、
    ' 最後の1件の途中で押されたときは見送り0件でも中断は起きている。
    ' 従来はcancelSkipNだけを見ていたため、その場合に「取り込みが完了
    ' しました」と言っていた(利用者が止めたことを、こちらが忘れた顔をする)。
    Dim cancelHit As Boolean: cancelHit = False
    ' R15-FixB(FB-5): 事前確認で利用者が「いいえ」を選んだ件数。失敗ではない
    ' ので ngCount には入れない(入れると「1件取り込めませんでした。状態を
    ' ご確認ください」と、自分で見送っただけの人に不具合を疑わせる)。
    Dim declinedN As Long: declinedN = 0
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

    ' R19-5b(実機第6報⑤): 他のブックと同じプロセスに居るなら、取込を始める前に
    ' もう一度だけ確認する(判定・文言・モーダルは modIntegrity 側。ここは
    ' 戻り値だけを見る)。「いいえ」はキャンセルと同じ扱いにする=取込は
    ' 1件も始まらず、mBatchIngesting も立てないので後始末も要らない。
    ' 無人同期(modShelfSync)はこの関数を通らないので silent は silent のまま。
    If Not modIntegrity.ConfirmIngestWhenCohabit() Then
        AddFilesResult = "ok=0;ng=0;capped=0;chunks=0;reasons="
        Exit Function
    End If

    ' R10c(H2): ここから下はバッチ取込中。全終了経路(正常/AddFailed)で必ず解除する。
    mBatchIngesting = True
    mBatchSince = Now
    ' R15-6c: 前回の中断の印を必ず下ろしてから始める。残っていると、押した
    ' 覚えの無い取込が1件目の頁境界でいきなり止まる。
    ResetCancel

    ' R12-4: 取込はシートを大きく書き換える。検索用のベクトルキャッシュ
    ' (32bit Excelで最大126MB)を抱えたまま取込に入ると、取込側の配列と
    ' ピークが重なってメモリ不足(err7)を起こしやすい。先に解放しておく
    ' (次の検索で作り直される)。
    On Error Resume Next
    modVecCache.ResetVecCache
    On Error GoTo AddFailed

    ' R10-2: ファイル選択が確定した(キャンセルではない)ので、GS未検出の
    ' 案内カード(セッション1回きり)を再提示可能に戻す。利用者が能動的に
    ' 取込を実行する入口のみが対象で、自動同期(modShelfSync)からは呼ばない。
    ' vision機能が無効なビルドでもInvokeFeature側で安全に無視される。
    On Error Resume Next
    modFeatures.InvokeFeature "vision", "ResetGsGuidance", Array()
    On Error GoTo AddFailed

    Dim capMax As Long: capMax = modConfig.GetLong("shelf_max_chunks", modAppDef.DEFAULT_SHELF_MAX_CHUNKS)
    If capMax < 1 Then capMax = modAppDef.DEFAULT_SHELF_MAX_CHUNKS

    ' R10-5(実機報告「止まってるのか分からない」): 経過秒÷完了件数の単純平均で
    ' 目安時間を出す。1件目はまだ実績が無いのでETA無し(modUtil.EtaText任せ)。
    Dim tIngStart As Double: tIngStart = Timer

    Dim i As Long
    For i = 1 To fd.SelectedItems.count
        ' R15-6c: ファイル境界の中断確認。1件目に手を付ける前でも効く。
        ' 残りは【手を付けずに】見送り、件数を必ず利用者へ伝える。
        If CancelRequested() Then
            cancelHit = True
            cancelSkipN = fd.SelectedItems.count - (i - 1)
            Exit For
        End If

        Dim etaPart As String: etaPart = ""
        If i > 1 Then
            Dim elapsedSec As Double: elapsedSec = Timer - tIngStart
            If elapsedSec >= 0 Then   ' 日跨ぎのTimerロールオーバーはETA非表示に落とす
                etaPart = modUtil.EtaText(fd.SelectedItems.count - (i - 1), _
                    (elapsedSec * 1000#) / (i - 1))
            End If
        End If
        ' R10c(M5): 表示系は modEmbed/modShelfSync と同形にOERNで挟む
        ' (バナー描画の失敗で AddFailed へ飛ばし取込を止めてはならない)。
        ' R13-4b: ETAが出せない最初の1件(1件だけの取込では最後まで)は、
        ' 従来「1/1件 <ファイル名>」で止まって見えた。段階バナー(コピー中/
        ' 本文抽出中/…)が引き継ぐまでの間を「処理中…」で埋め、
        ' 「動いているのか止まっているのか分からない」を作らない(憲章§3-2)。
        Dim tailPart As String: tailPart = ""
        If LenB(etaPart) = 0 Then tailPart = " 処理中…"
        ' R15-FixA(FA-6): 中断できるのは取込だけ。ここは中断ボタン付きで出す。
        On Error Resume Next
        ShowIngestBanner modUtil.ProgressText(i, fd.SelectedItems.count, etaPart) & " " & _
            modUtil.SafeLeft(modUtil.FileNameOf(CStr(fd.SelectedItems(i))), 40) & tailPart
        On Error GoTo AddFailed

        ' 2026-07-31(§7裁定3件目 採用1): 手動「資料を追加」でこのブック自身
        ' (フルパス一致、または同名の一時コピー)を選んだ場合は取り込まない。
        ' 自動同期側の判定は modShelfScan.IsExcludedFile(フォルダ走査専用)が
        ' 別に持っており、こちらは手動選択という別経路の共通入口を担うため
        ' 重複ではない。
        If IsSelfWorkbookFile(CStr(fd.SelectedItems(i))) Then
            ngCount = ngCount + 1
            BumpReasonCount reasonKeys, reasonCounts, reasonN, "self"
        ElseIf modShelf.TotalChunks() >= capMax Then
            cappedN = cappedN + 1
        Else
            ' R12-3-3: 恒久失敗(failed_permanent)で自動同期から外れたファイルも、
            ' 利用者がここで明示的に選び直したのなら、もう一度試すのが正しい。
            ' 連続失敗の記録を消してから取込へ入る。
            On Error Resume Next
            modShelfStore.ResetFailCountForPath CStr(fd.SelectedItems(i))
            On Error GoTo AddFailed

            Dim beforeChunks As Long: beforeChunks = modShelf.TotalChunks()
            Dim st As String: st = "failed"
            Dim errCd As String: errCd = ""
            ' silent:=True(レビュー4-A): 1件ごとのモーダルを全廃する。同名衝突
            ' (E0504)だけが silent を見ずにモーダルを出し、選んだ件数ぶん出た
            ' 上で集計メッセージが同じことをもう一度言っていた(二重報告)。
            ' 情報は戻り値の reasons に載るので失わない。
            ' R15-FixA(FA-1): interactive:=True。silent(1件ごとのモーダル抑止)
            ' と、【利用者がその場にいる取込か】は別の問いで、両方を silent 1本で
            ' 表していたために「何時間かかります」の事前確認(R15-7b)が全経路で
            ' 死んでいた(A-H1/B-H1)。手動で選んだこの経路だけが True。
            On Error Resume Next
            st = modShelf.IngestFile(CStr(fd.SelectedItems(i)), "self", True, errCd, True)
            On Error GoTo AddFailed
            chunksAdded = chunksAdded + (modShelf.TotalChunks() - beforeChunks)
            If st = "done" Or st = "partial" Then
                okCount = okCount + 1
                If st = "partial" Then warnN = warnN + 1
                ' R18-2c: ここにあった SaveCheckpoint は modShelf.IngestFile の
                ' Finish へ移した(この経路にしか保存が無く、スクショ・ナレッジ
                ' 登録・パック取込など6経路が無保存だった。二重呼び出しを避ける
                ' ため、ここは純減=削除)。
            ElseIf errCd = "declined" Then
                ' R15-FixB(FB-5): 見送りは失敗と別に数える。内訳には残すので
                ' (reasons=declined:N)、呼び出し元も件数を説明できる。
                declinedN = declinedN + 1
                BumpReasonCount reasonKeys, reasonCounts, reasonN, "declined"
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

    ' R10-5: 完了を1回だけトーストで知らせる(連呼はしない。ここは完了1回のみ
    ' なので1.1秒のブロッキングも許容)。
    ' R10c(M4): 「バナーを閉じる → トースト」の順にする。逆順だと1.1秒のあいだ
    ' 2枚がならび、どちらが今の状態なのか読み手に判断できない(バナーには
    ' 最後のファイル名が残っている)。modShelfSync も同じ順序。
    ' R15-FixA(FA-7): ループを抜けたあとにもう一度見る。最後の1件の途中で
    ' 押された中断は、ループ先頭の判定を二度と通らないのでここでしか拾えない。
    If CancelRequested() Then cancelHit = True

    Dim warnPart As String: warnPart = ""
    If warnN > 0 Then warnPart = "/注意" & warnN & "件"
    ' R15-FixB(FB-5): 見送りは0件なら言わない(常に「見送り0件」と出すと、
    ' 見送りという言葉だけが毎回目に入って意味を失う)。
    If declinedN > 0 Then warnPart = warnPart & "/見送り" & declinedN & "件"
    Dim doneToast As String
    ' R15-6c: 中断したのに「完了しました」と言わない(利用者が止めたことを
    ' こちらが忘れた顔をするのは不誠実)。トーストは1回だけ=待ちは増やさない。
    If cancelHit Then
        ' R15-FixB(FB-5): 中断で手を付けなかった件数は「未取込」と言う。
        ' 「見送り」は【利用者が確認で自分から見送った件数】に譲る
        ' (同じ言葉が1つのトーストに2つ並ぶと、どちらが何なのか読めない)。
        doneToast = "中断しました(成功" & okCount & "件/失敗" & ngCount & "件" & _
            warnPart & "/未取込" & cancelSkipN & "件)"
    Else
        doneToast = "取り込みが完了しました(成功" & okCount & "件/失敗" & ngCount & "件" & _
            warnPart & ")"
    End If
    On Error Resume Next
    modUIMain.HideProgress
    modSkin.ShowToast doneToast, "info"
    On Error GoTo AddFailed

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
    If warnN > 0 Then
        ' R13-3a: 「成功」の中に partial が混ざっていることを黙らない。
        ' 何がどう足りないのかは資料ごとに違うので、行き先(メモ)だけ示す。
        msg = msg & vbLf & "うち" & warnN & "件は一部だけの取込です" & _
            "(マイ本棚の一覧で、そのカードのメモをご確認ください)。"
    End If
    ' R15-FixB(FB-5): 見送りは失敗の文とは別の文で言う。理由(推定時間)は
    ' カードのメモに載っているので、ここでは件数と「次はどうすれば入るか」だけ。
    If declinedN > 0 Then
        msg = msg & vbLf & declinedN & "件は時間がかかるため取り込みませんでした" & _
            "(確認で「いいえ」を選んだ資料です)。もう一度選んで「はい」を選ぶと取り込めます。"
    End If
    ' R15-FixA(FA-7): 中断はモーダルでも必ず言う。未取込0件(最後の1件の
    ' 途中で押した)の場合も「止まったこと」と「その資料がどうなったか」を
    ' 黙らない(途中まで読めた頁は partial として本棚カードのメモに出る)。
    If cancelHit Then
        msg = msg & vbLf & "利用者の操作で中断しました。"
        If cancelSkipN > 0 Then
            msg = msg & vbLf & "うち" & cancelSkipN & "件は取り込んでいません。" & _
                "もう一度「資料を追加」から選び直せば取り込めます。"
        End If
    End If
    Dim dupN As Long
    dupN = ReasonCountOf(reasonKeys, reasonCounts, reasonN, "E0504")
    If dupN > 0 Then
        msg = msg & vbLf & "同じ名前の資料が別の場所から登録済みのものが" & dupN & "件あります。" & _
            vbLf & "ファイル名を変えて入れ直すか、マイ本棚で古いほうを削除してからお試しください。"
    End If
    Dim selfN As Long
    selfN = ReasonCountOf(reasonKeys, reasonCounts, reasonN, "self")
    If selfN > 0 Then
        msg = msg & vbLf & "このファイルは本ツール自身のため取り込めません。"
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
        ' R13-8: 取込直後は「読める状態になった」ことしか伝えておらず、次に
        ' 何をすればよいかの導線が無かった。チャットへボタンの存在を1文添える
        ' (このモーダルの表示先だけに足す。戻り値の集計文字列は変えない)。
        ' R13-L6: 絵文字は付けない。ネイティブMsgBoxではサロゲートペアが
        ' 「??」で描かれる(modLog.bas 末尾の実機報告)。
        If okCount > 0 Then
            msg = msg & vbLf & "「チャットへ」ボタンからそのまま質問できます。"
        End If
        MsgBox msg, vbInformation, modAppDef.APP_NAME
    End If

    mBatchIngesting = False   ' R10c(H2): 正常終了経路の解除
    ResetCancel               ' R15-6c: 中断の印を取込の外へ持ち越さない
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
    mBatchIngesting = False   ' R10c(H2): 異常終了経路でも必ず解除する
    ResetCancel               ' R15-6c: 異常終了経路でも印を持ち越さない
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

' このブック自身、または同名の一時コピーかどうか(§7裁定3件目 採用1)。
'   フルパス一致(完全一致)に加え、拡張子を問わずファイル名だけが
'   ThisWorkbook.Nameと一致する場合も自分自身とみなす(共有フォルダ経由で
'   コピーされた同名の一時ファイルを選んでしまうケースを含めて防ぐ)。
Private Function IsSelfWorkbookFile(ByVal path As String) As Boolean
    On Error Resume Next
    Dim me_ As String: me_ = ThisWorkbook.FullName
    Dim myName As String: myName = ThisWorkbook.Name
    If LenB(me_) = 0 Then Exit Function
    If StrComp(path, me_, vbTextCompare) = 0 Then
        IsSelfWorkbookFile = True
    ElseIf StrComp(modUtil.FileNameOf(path), myName, vbTextCompare) = 0 Then
        IsSelfWorkbookFile = True
    End If
    On Error GoTo 0
End Function

Private Function BuildFilterPattern() As String
    Dim exts() As String: exts = Split(modExtractor.SupportedExts(), ",")
    Dim parts() As String: ReDim parts(LBound(exts) To UBound(exts))
    Dim i As Long
    For i = LBound(exts) To UBound(exts)
        parts(i) = "*." & Trim$(exts(i))
    Next i
    BuildFilterPattern = Join(parts, ";")
    ' R36波1(§4): vision(画像解析)が有効な端末だけ、📁 追加のフィルタへ
    ' png/jpg/jpegを足す。SUPPORTED_EXTS自体は変えない(足すとE0301→
    ' TryVisionFallbackの委譲条件が壊れるため。フィルタ文字列だけの追加)。
    If modFeatures.FeatureEnabled("vision") Then
        BuildFilterPattern = BuildFilterPattern & ";" & modShelfVision.ImageDialogPattern()
    End If
End Function
