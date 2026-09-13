Attribute VB_Name = "modIntegrity"
Option Explicit

' ============================================================================
' modIntegrity - データ整合性の観測点(2026-08-05 R18-2b/2d・実機第5報⑧)
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   「昨日入れた資料が今日は0件になっている」という報告に対して、アプリ側に
'   【気付く仕組みが1つも無かった】。表示は my_manifest の chunk_count を
'   そのまま出すだけで実データと突き合わせず、保存が次回に残ったかどうかを
'   確かめる記録も無かった(調査agent1)。憲章§4-2「観測性ファースト」。
'
' ここが持つのは2つだけ:
'   (1) 台帳と実データの突合(ReconcileChunkCount) — マイ本棚のカードが
'       my_knowledge の実行数と食い違ったら、実行数を正として台帳を直す。
'   (2) 保存の履歴と起動時の突合(RecordSaveMark / WarnAtStartup) —
'       保存が成功した時点の (my_knowledge行数, ブックのフルパス) を ui_state に
'       控え、次の起動で行数が減っていたら1回だけ警告する。あわせて、ブックが
'       %TEMP%・Temp1_*.zip 配下(=zipの中から直接開いた/一時展開コピー)なら
'       「保存が次回に残らない場所」であることを起動時に1回伝える。
'
' 置き場が基盤層(src/core)である理由:
'   modBoot は30,000字上限まで残り僅かで判定本体を置けず、modShelfStore /
'   modShelf も同様に逼迫している(憲章§4-6)。判定材料は my_knowledge /
'   my_manifest / ui_state のシートだけで、上位層を1つも呼ばない(R1)。
'   modDiag.WarnIfReadOnly が同じ理由で同じ層に置かれているのと同型。
'
' 判定そのもの(DataShrunk / IsVolatilePath / ReconcileStatText)は Excel に
' 触れない純ロジックへ切り出し、modTestsPure15 がゴールデンで固定する。
' ============================================================================

' ui_state のキー(modState.SaveState/LoadState 経由。作法は nexus_theme と同じ)。
Private Const KEY_LAST_ROWS As String = "integrity_last_rows"
Private Const KEY_LAST_PATH As String = "integrity_last_path"

' 起動時の警告は1セッション1回まで(Bootが二度走っても二度は出さない)。
Private mWarned As Boolean

' 直近の一括削除が「消してから書けなかった」(=行が欠けた)側で失敗したか。
' R33H M8。ResetPurgeMark で落とし、PurgePartial で読む。
Private mPurgePartial As Boolean

' R29H F2b: WarnAtStartupはInitUI(modApp.LaunchNexus内)より前(modBoot.bas)に
' 呼ばれるため、この時点ではAddChatBubbleが描けない。文言をここへ保留し、
' InitUI後の既存経路(modBoard.BootBoard)からFlushPendingBubblesで1回だけ出す。
Private mPendingBubbles As Collection

' UsedRange の焼き付き判定のしきい値(pt)。R19H FA-5(ii)。根拠は
' IsUsedRangeBloated の見出しコメント(倍率をやめて絶対値にした理由)を参照。
Private Const BLOAT_W_PT As Double = 2400
Private Const BLOAT_H_PT As Double = 12000

Private Sub QueueBubble(ByVal msg As String)
    If mPendingBubbles Is Nothing Then Set mPendingBubbles = New Collection
    mPendingBubbles.Add msg
End Sub

' FlushPendingBubbles - InitUI後(modBoard.BootBoardの先頭)から1回だけ呼ばれる。
Public Sub FlushPendingBubbles()
    On Error Resume Next
    If mPendingBubbles Is Nothing Then Exit Sub
    Dim v As Variant
    For Each v In mPendingBubbles
        modUI.AddChatBubble "ai", CStr(v)
    Next v
    Set mPendingBubbles = Nothing
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' ReconcileStatText - 台帳の統計文字列の chunk_count を actualN へ直す(純)。
' ----------------------------------------------------------------------------
' statText は modShelf.SourceList が組み立てた
' "status|ingested_at|chunk_count|error_note|origin"。3つ目が actualN と
' 一致していれば statText をそのまま返す(=呼び出し側は何もしない合図)。
' error_note に "|" が混じっていても Join で元どおりに戻る(触るのは3つ目だけ)。
' 5つ未満の壊れた文字列は触らない(直せないものを壊さない)。
Public Function ReconcileStatText(ByVal statText As String, ByVal actualN As Long) As String
    ReconcileStatText = statText
    Dim f() As String: f = Split(statText, "|")
    If UBound(f) - LBound(f) + 1 < 5 Then Exit Function
    Dim idx As Long: idx = LBound(f) + 2
    Dim shown As Long: shown = -1
    If IsNumeric(f(idx)) Then shown = CLng(f(idx))
    If shown = actualN Then Exit Function
    f(idx) = CStr(actualN)
    ReconcileStatText = Join(f, "|")
End Function

' ----------------------------------------------------------------------------
' IndexOfName - names(0..n-1) から target の位置を探す(無ければ -1。純ロジック)。
' ----------------------------------------------------------------------------
' R18-2b の突合で my_knowledge の全行に対して呼ばれるため、直前に一致した位置
' (hint)を先に1回だけ試す。my_knowledge は source 単位の連続ブロックで追記
' されるので、ほぼ全ての行がこの1回の比較で決まり、実質 O(行数)で済む
' (毎回先頭から探すと 資料数×行数 の比較になり、2万行の本棚で描画が固まる)。
' hint が範囲外・不一致でも結果は変わらない(速さだけの助言で、正しさは
' 常に線形探索が保証する)。比較は StrComp(vbTextCompare)=大小無視で、
' manifest 側の照合(FindConflictingManifestPath 等)と作法を揃える。
Public Function IndexOfName(ByRef names() As String, ByVal n As Long, _
                            ByVal target As String, ByVal hint As Long) As Long
    IndexOfName = -1
    If n < 1 Then Exit Function
    If LenB(target) = 0 Then Exit Function
    If hint >= 0 And hint < n Then
        If StrComp(names(hint), target, vbTextCompare) = 0 Then
            IndexOfName = hint
            Exit Function
        End If
    End If
    Dim i As Long
    For i = 0 To n - 1
        If StrComp(names(i), target, vbTextCompare) = 0 Then
            IndexOfName = i
            Exit Function
        End If
    Next i
End Function

' ----------------------------------------------------------------------------
' ReconcileChunkCount - (2b の本体)カード1枚ぶんの突合と自動修復。
' ----------------------------------------------------------------------------
' 一致していればシートにも触らずログも書かない(通常はここで即戻る)。
' 食い違うときだけ実行数を正として manifest の5列目を直し、usage_log に
' manifest_fix を1行残す。台帳を直せなくても(シート無し・行が見つからない・
' 書込み失敗)戻り値は実行数に揃うので、少なくとも画面が嘘をつくことは無い。
' 突合の失敗が一覧の描画を止めてはならないので全体を OERN で包む(憲章§4-4)。
Public Function ReconcileChunkCount(ByVal sourceName As String, ByVal statText As String, _
                                    ByVal actualN As Long) As String
    ReconcileChunkCount = statText
    On Error Resume Next
    Dim fixed As String: fixed = ReconcileStatText(statText, actualN)
    If fixed = statText Then Exit Function
    ReconcileChunkCount = fixed

    Dim wsM As Worksheet: Set wsM = GetSheet(modAppDef.SH_MANIFEST)
    If wsM Is Nothing Then Exit Function
    Dim r As Long: r = ManifestRowByName(wsM, sourceName)
    If r < 2 Then Exit Function
    Dim shownTxt As String: shownTxt = CStr(wsM.Cells(r, 5).Value)
    wsM.Cells(r, 5).Value = actualN
    modLog.LogUsage "manifest_fix", "", modUtil.SafeLeft(sourceName, 200) & _
        ": 台帳" & shownTxt & "件 → 実データ" & actualN & "件へ修復"
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' RecordSaveMark - 保存が成功した瞬間の事実を控える(2d)。
' ----------------------------------------------------------------------------
' modShelfBatch.SaveCheckpoint が ThisWorkbook.Save の【直前】に1行呼ぶ
' (R18H FA-3: 後に呼ぶとディスク上のマークだけが常に1世代古くなる)。
' 加えて、意図した削除の直後にも呼ぶ(modShelf.DeleteSource /
' modShelfStore.RemoveRowsByOrigin(Prefix)。保存はしない=削除は保存の合図では
' ない。ui_state はブックと一緒に保存/破棄されるのでどちらでも辻褄が合う)。
' 控えるのは (my_knowledge の行数, ThisWorkbook.FullName) の2つだけ。
' 「保存した」と「次に開いたときに残っていた」は別の事実で、後者を確かめる
' 手段がこれまで1つも無かった(調査agent1 §4: zip直開き・一時展開コピーは
' 保存も成功し ReadOnly でもないため、アプリからは完全に正常に見える)。
' 記録の失敗が保存を壊してはならないので OERN で包む。
Public Sub RecordSaveMark()
    On Error Resume Next
    modState.SaveState KEY_LAST_ROWS, CStr(KnowledgeRowCount())
    modState.SaveState KEY_LAST_PATH, ThisWorkbook.FullName
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' RewriteRowsAfterPurge - 一括削除の後始末: 生存行だけを書き戻す(R33H F2)
' ----------------------------------------------------------------------------
' 旧実装(modShelfStore.RemoveRowsByOrigin)は「生存行を上へ詰めて一括代入 →
' 末尾を ClearContents」の順だった。代入の途中で落ちると【生存行が二重に
' 存在】したまま残る ―― 重複は件数にも画面にも出ないので誰も気付けず、
' 検索結果だけが静かに汚れる(このモジュールが観測しようとしている
' 「台帳と実データの食い違い」を、削除側が自分で作っていた)。
' 先に消してから書けば、途中で落ちても残るのは「消えすぎ」側になり、
' 部門チャンネルから取り込み直せば戻る=取り返しがつく側へ倒す。
' 書き込みは200行バッチ(数千行×長文の1回代入は実行時エラー7になり得る。
' modPack.WriteRowsBatched・modShelfStore.FlushNormBackfill と同じ作法)。
'   ws       : 対象シート。1行目は見出しなのでデータは2行目から。
'   colCount : 運ぶ列数(全列を運ばないと別チャンクの列が混ざる)。
'   survivors: (1 To n以上, 1 To colCount以上) の2次元配列。
'   n        : 生存行数。oldRows: 消す前のデータ行数。
' 戻り値: 全行を書けたら True。1行でも書けなければ False(呼び出し元が
'   「削除に失敗した可能性があります」と言うための材料。旧実装は失敗を
'   握り潰し「0件 削除しました」と表示する余地があった)。
' 置き場が基盤層なのは modShelfStore が残819字で分岐を書けないため
'   (このモジュールの冒頭に書いた理由と同じ。憲章§4-6)。
' 2026-08-16(R33H M8): 失敗の【向き】を呼び出し元へ残す。ClearContents に失敗
'   したなら1行も消していない(未着手=資料はそのまま。やり直せば直る)が、
'   消してから書けなかったなら行が欠けている(やり直しても戻らない)。同じ
'   「失敗」でも利用者への案内が正反対になるので、後者だけ印を立てる。
'   一連の削除操作の前に ResetPurgeMark を呼ぶこと(複数タグを回す経路が
'   あるため、印は操作の単位でクリアする)。
Public Function RewriteRowsAfterPurge(ByVal ws As Worksheet, ByVal colCount As Long, _
                                      ByRef survivors As Variant, ByVal n As Long, _
                                      ByVal oldRows As Long) As Boolean
    Const PURGE_BATCH_ROWS As Long = 200
    If ws Is Nothing Then Exit Function
    If colCount < 1 Then Exit Function
    If n < 0 Then Exit Function
    If oldRows < n Then Exit Function

    On Error Resume Next
    Err.Clear
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + oldRows, colCount)).ClearContents
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If

    Dim st As Long, k As Long, wrote As Long, r As Long, c As Long
    For st = 1 To n Step PURGE_BATCH_ROWS
        k = n - st + 1
        If k > PURGE_BATCH_ROWS Then k = PURGE_BATCH_ROWS
        Dim buf() As Variant: ReDim buf(1 To k, 1 To colCount)
        For r = 1 To k
            For c = 1 To colCount
                buf(r, c) = survivors(st + r - 1, c)
            Next c
        Next r
        Err.Clear
        ws.Range(ws.Cells(1 + st, 1), ws.Cells(st + k, colCount)).Value = buf
        If Err.Number = 0 Then wrote = wrote + k
    Next st
    RewriteRowsAfterPurge = (wrote = n)
    If wrote <> n Then mPurgePartial = True   ' 消した後で書けなかった=欠落側
    On Error GoTo 0
End Function

' PurgePartial - 直近の一括削除が【欠落】側で失敗したか(R33H M8)。
'   False は「1行も消していない(未着手)」= やり直せば戻る側。
Public Function PurgePartial() As Boolean
    PurgePartial = mPurgePartial
End Function

' ResetPurgeMark - 一括削除の一連の操作を始める前に印を落とす(R33H M8)。
Public Sub ResetPurgeMark()
    mPurgePartial = False
End Sub

' ----------------------------------------------------------------------------
' DataShrunk - 「前回保存した時より資料が減ったか」(純ロジック)。
' ----------------------------------------------------------------------------
' prevRows <= 0(記録が無い・壊れている)は判定しない=初回起動や旧ブックから
' の移行で、根拠なく警告を出さない。増えている/同じは当然 False。
Public Function DataShrunk(ByVal prevRows As Long, ByVal curRows As Long) As Boolean
    If prevRows <= 0 Then Exit Function
    DataShrunk = (curRows < prevRows)
End Function

' ----------------------------------------------------------------------------
' IsVolatilePath - 「保存しても次回に残らない場所」か(純ロジック)。
' ----------------------------------------------------------------------------
' zip をダブルクリックして中の xlsm を直接開くと、Windows は
' %TEMP%\Temp1_<zip名>.zip\ のような一時フォルダへ展開してそこを開く。
' 保存は成功し ReadOnly でもないので、アプリからは完全に正常に見えるが、
' 次に開くのは別の展開コピーで、前回の取込は1件も残っていない
' (調査agent1 §5 順位2「検知ゼロ」)。判定材料はパス文字列だけなので、
' 環境に依らない素朴な一致で見る(正規表現も FileSystemObject も使わない
' =管理端末で塞がれても必ず動く)。
'   ・"temp1_"  : zip 直開きの展開フォルダ名(Temp1_, Temp2_ ... と続く)
'   ・".zip\"   : zip の中を直接開いている
'   ・tempA/tempB(=Environ("TEMP")/Environ("TMP") の実値)配下
'
' 2026-08-05(R18H FA-4 / A-M3・B-M5): "\temp\" の部分一致を廃止した。
' D:\temp\約款\ のような【正規の作業フォルダ】や \\share\Temp\ の共有を
' 一時フォルダと誤判定し、正しく保存できている端末に「保存が次回に残らない
' 場所です」という嘘を毎回出していた(憲章§3-3の逆・その機能はもう使われない)。
' 判定は実際の %TEMP%/%TMP% との【前方一致】に限る。環境変数はここでは読まず
' 引数で受け取る(呼び出し元 WarnAtStartup が読む)=この関数は純ロジックの
' ままで、テストが端末の環境変数に左右されない。
' 大文字小文字は無視する(パスの表記は端末ごとに揺れる)。
Public Function IsVolatilePath(ByVal folderPath As String, _
                               Optional ByVal tempA As String = "", _
                               Optional ByVal tempB As String = "") As Boolean
    Dim p As String: p = LCase$(Trim$(folderPath))
    If LenB(p) = 0 Then Exit Function
    If InStr(p, "temp1_") > 0 Then IsVolatilePath = True
    If InStr(p, ".zip\") > 0 Then IsVolatilePath = True
    If UnderDir(p, tempA) Then IsVolatilePath = True
    If UnderDir(p, tempB) Then IsVolatilePath = True
End Function

' p(小文字化済み)が dirPath そのもの、またはその配下か(前方一致・R18H FA-4)。
' 末尾に区切りを補ってから比べる: 補わないと "C:\Temp" が "C:\Temp2\..." にも
' 当たる(1文字違いのフォルダを巻き込む誤検知)。
Private Function UnderDir(ByVal p As String, ByVal dirPath As String) As Boolean
    Dim d As String: d = LCase$(Trim$(dirPath))
    If LenB(d) = 0 Then Exit Function
    If Right$(d, 1) <> "\" Then d = d & "\"
    Dim pp As String: pp = p
    If Right$(pp, 1) <> "\" Then pp = pp & "\"
    If Len(pp) < Len(d) Then Exit Function
    UnderDir = (Left$(pp, Len(d)) = d)
End Function

' Environ の読み出し(失敗しても空文字で返る)。判定そのものは純ロジック側。
Private Function SafeEnv(ByVal keyName As String) As String
    On Error Resume Next
    SafeEnv = Environ$(keyName)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' ShrinkWarnMsg / VolatileWarnMsg - 警告文(純ロジック・BMPの文字だけ)。
' ----------------------------------------------------------------------------
' 非BMPの絵文字は CP932 の実行文へ変換される往復で化け、ダイアログには
' 化けたまま出る(実機第5報⑤)。ここでは記号を一切使わず日本語だけで書く。
' 「何が起きたか」+「次の一手」を必ず両方言う(憲章§3-3)。
Public Function ShrinkWarnMsg(ByVal prevRows As Long, ByVal curRows As Long, _
                              ByVal prevPath As String, ByVal curPath As String) As String
    ShrinkWarnMsg = "前回このアプリを保存したときは資料の中身が " & prevRows & _
        " 件ありましたが、今は " & curRows & " 件になっています。" & vbLf & _
        "別の場所にあるコピーを開いている可能性があります。" & vbLf & vbLf & _
        "前回保存した場所: " & modUtil.SafeLeft(prevPath, 200) & vbLf & _
        "今開いている場所: " & modUtil.SafeLeft(curPath, 200) & vbLf & vbLf & _
        "心当たりが無いときは、この画面を管理者に見せてください。"
End Function

Public Function VolatileWarnMsg(ByVal curPath As String) As String
    VolatileWarnMsg = "保存が次回に残らない場所でこのファイルを開いています。" & vbLf & _
        "取り込んだ資料は、次にこのファイルを開いたときに残りません。" & vbLf & vbLf & _
        "今の場所: " & modUtil.SafeLeft(curPath, 200) & vbLf & vbLf & _
        "配布された zip を右クリックして「すべて展開」し、" & vbLf & _
        "展開されたフォルダの中のファイルを開き直してください。"
End Function

' ----------------------------------------------------------------------------
' R19-5(実機第6報⑤): 作業用Excelとの「同居」の検出と警告。
' ----------------------------------------------------------------------------
' 機序(調査⑤班2章): Excelは既定で複数ブックを1プロセスへ結合(マージ)する。
' 本体xlsmをエクスプローラーからダブルクリックすると、生き残っていた作業用
' Excelのプロセスへ本体が吸い込まれ、以後は本体のOCR同期呼び出しがプロセス
' 全体のメッセージポンプを止める=作業用Excelも「■中断」も丸ごと無反応になる
' (実機第6報⑤の「①も中断ボタンも全滅」)。DisableMergeInstance はダブル
' クリックには効かないことが公式に明記されており(調査⑤ウェブ班3-(b))、
' VBAだけで結合を防ぐ完全な実装解は存在しない。できるのは「検出して正直に
' 伝え、危険な操作の前に一度止める」ことだけ(仕様 R19-5 の裁定)。
'
' 検出の対象は Application.Workbooks。これは【自分のプロセス内で開いている
' ブックだけ】を指すというSDIの仕様がそのまま「自分以外が同居しているか」の
' 答えになる(調査⑤ウェブ班3-(c)。Application.Hwnd はアクティブウィンドウしか
' 返さないので使わない)。
'
' 2026-08-06(R19H FA-3 / A-H③・B-H②): 旧実装は Workbooks.Count > 1 だった。
' しかし Workbooks には【利用者から見えないブック】まで並ぶ:
'   ・PERSONAL.XLSB(個人用マクロブック。Excelが常に非表示で開く。国内の
'     業務端末では珍しくない)
'   ・アドイン(IsAddin=True。参照設定・組織配布のxlam)
'   ・可視ウィンドウを持たないブック(他マクロが Visible=False で開いたもの)
' これらは「取込中に一緒に固まる相手」ではないので、同居の警告を出す理由が
' 無い。PERSONAL.XLSB を使っている人は【毎回・単独で開いても】警告が出て、
' しかも言われたとおりに他のExcelを全部閉じても消えない=直せない警告になる
' (憲章§3-3の逆。狼少年になった警告は次から読まれない)。
' 数える対象は「自分以外・アドインでない・可視ウィンドウを持つ」ブックだけ。
'
' 相手のブック名は【列挙しない】: 個人情報になり得るうえ、20冊開いている人の
' ダイアログが読めない長文になる。利用者に要る情報は「同居している」事実と
' 「どうすればよいか」の2つだけ(憲章§3-3)。
Public Function CohabitCount() As Long
    On Error Resume Next
    CohabitCount = CohabitOtherCount(CollectBookLines(), ThisWorkbook.Name)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' CohabitOtherCount - 同居している「可視の他ブック」の数(純ロジック)。
' ----------------------------------------------------------------------------
' bookLines: 1行1冊の "ブック名<TAB>可視ウィンドウ数<TAB>アドインなら1" を
'            vbLf で連ねた文字列(自分自身の行が含まれていてよい)。
' selfName : 自分のブック名(ThisWorkbook.Name)。大小無視で除く。
' Excel を1つも触らないのでテストで固定できる(modTestsPure18)。Hit() 型のような
' 「LOで組み立てられない引数」も使わない=文字列1本で全ケースを再現できる。
' 壊れた行(TABが足りない・数値でない)は【数えない】: 数え漏らしても出るのは
' 「警告が出ない」だけだが、数え過ぎると出せない警告を出し続けることになる。
Public Function CohabitOtherCount(ByVal bookLines As String, _
                                  ByVal selfName As String) As Long
    If LenB(bookLines) = 0 Then Exit Function
    Dim lines() As String: lines = Split(bookLines, vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim f() As String: f = Split(lines(i), vbTab)
        If UBound(f) - LBound(f) + 1 >= 3 Then
            Dim nm As String: nm = Trim$(f(LBound(f)))
            Dim vis As Long: vis = CLng(Val(f(LBound(f) + 1)))
            Dim isAdd As Long: isAdd = CLng(Val(f(LBound(f) + 2)))
            If LenB(nm) > 0 And vis > 0 And isAdd = 0 Then
                If StrComp(nm, selfName, vbTextCompare) <> 0 And _
                   StrComp(nm, "personal.xlsb", vbTextCompare) <> 0 Then
                    CohabitOtherCount = CohabitOtherCount + 1
                End If
            End If
        End If
    Next i
End Function

' Workbooks を「1行1冊」の文字列へ畳む。Excel に触れるのはここだけで、
' 数える規則そのものは上の純ロジックが持つ(2箇所に分かれない=憲章§4-5)。
Private Function CollectBookLines() As String
    Dim sb As String
    On Error Resume Next
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        Dim vis As Long: vis = 0
        Dim k As Long
        For k = 1 To wb.Windows.count
            If wb.Windows(k).Visible Then vis = vis + 1
        Next k
        Dim isAdd As Long: isAdd = 0
        If wb.IsAddin Then isAdd = 1
        If LenB(sb) > 0 Then sb = sb & vbLf
        sb = sb & wb.Name & vbTab & vis & vbTab & isAdd
    Next wb
    On Error GoTo 0
    CollectBookLines = sb
End Function

' 同居しているか(純ロジック。CohabitCount は【自分以外の可視ブック】を
' 数えるので、1冊でもあれば同居)。
Public Function IsCohabiting(ByVal otherCount As Long) As Boolean
    IsCohabiting = (otherCount > 0)
End Function

' 起動時のモーダル文(BMPの文字だけ。非BMPはCP932往復で化ける=実機第5報⑤)。
' 2026-08-06(R19H FA-7 / A-M⑦・B-M⑤): 「同梱の『MyBookshelfを起動』から
' 開き直してください」だけだと、そのランチャーを持っていない人(旧版のzip・
' xlsm だけを転送してもらった人・展開せず1ファイルだけ取り出した人)には
' 【存在しないものへの誘導】になり、言われたとおりにできない=直せない警告に
' なる。有る場合と無い場合の両方に効く1文へ揃える(ガードシート・READMEと同文)。
Public Function CohabitWarnMsg() As String
    CohabitWarnMsg = "他のExcelブックと同じプロセスで開かれています。" & vbLf & _
        "このまま取込を行うと、そのブックも一緒に固まります。" & vbLf & vbLf & _
        "同梱の「MyBookshelfを起動.bat」があればそれから、" & vbLf & _
        "無ければ他のExcelを全て閉じてから開き直してください。"
End Function

' 取込直前の再確認文(vbYesNo)。起動時の警告を見落とした人への最後の関所。
Public Function CohabitIngestMsg() As String
    CohabitIngestMsg = "他のブックが同じプロセスにあります。" & vbLf & _
        "取込中はそれらも操作できなくなります。" & vbLf & vbLf & _
        "続行しますか?"
End Function

' ----------------------------------------------------------------------------
' ConfirmIngestWhenCohabit - 取込入口の関所(R19-5b)。続行してよければTrue。
' ----------------------------------------------------------------------------
' 判定も文言もモーダルもここに置き、呼び出し側(modShelfBatch.AddFilesResult
' 冒頭)は「戻り値がFalseなら中止」の1つだけを知っていればよい(§4-5)。
' 同居していなければ何も出さずTrue=通常の取込は1ミリも変わらない。
' 表示に失敗しても取込を止めない(既定はTrue=続行)。
Public Function ConfirmIngestWhenCohabit() As Boolean
    ConfirmIngestWhenCohabit = True
    On Error Resume Next
    If Not IsCohabiting(CohabitCount()) Then Exit Function
    modLog.LogUsage "cohabit_detected", "ingest", "他ブックと同居した状態で取込が始まろうとしています"
    If MsgBox(CohabitIngestMsg(), vbExclamation + vbYesNo + &H10000, modAppDef.APP_NAME) <> vbYes Then
        ConfirmIngestWhenCohabit = False
        modLog.LogUsage "cohabit_abort", "ingest", "同居の確認で「いいえ」が選ばれ取込を中止しました"
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' WarnAtStartup - 起動時の突合(2d 本体)。modBoot から1行だけ呼ばれる。
' ----------------------------------------------------------------------------
' 順番は「置き場所の警告」が先: 一時フォルダで開いていることが分かっていれば、
' 件数が減っているのは当然の結果で、原因を先に伝えた方が利用者は動ける。
' どちらも1セッション1回まで。機能は止めない(閲覧・質問はそのまま使える)。
' 起動を止めないよう例外は外へ出さない。
'
' 2026-08-05(R18H FA-4 / A-M3・B-M5): (i) 置き場所の判定に渡すのは
' ThisWorkbook.Path(フォルダ)。FullName を渡すと、ファイル名に "temp1_" を
' 含むだけの資料まで一時扱いになる。(ii) 一時フォルダの警告を出しても
' 【減少チェックは打ち切らない】。従来はここで Exit Sub していたため、
' zip直開きの端末では「資料が減っている」という本命の異常を最後まで一度も
' 検知できなかった(2つは別の事実で、片方が他方を隠す理由は無い)。
Public Sub WarnAtStartup()
    On Error Resume Next
    If mWarned Then Exit Sub
    mWarned = True

    Dim curPath As String: curPath = ThisWorkbook.FullName
    Dim bookDir As String: bookDir = ThisWorkbook.Path
    Dim curRows As Long: curRows = KnowledgeRowCount()

    If IsVolatilePath(bookDir, SafeEnv("TEMP"), SafeEnv("TMP")) Then
        modLog.LogUsage "integrity_warn", "volatile_path", modUtil.SafeLeft(curPath, 300)
        MsgBox VolatileWarnMsg(curPath), vbExclamation, modAppDef.APP_NAME
    End If

    ' R19-5a(実機第6報⑤): 同居の検出。置き場所の次に伝えるのは「今このプロセスは
    ' 自分だけのものではない」という事実で、これを知らないまま取込に入ると
    ' 相手のブックごと固まる(上の CohabitWarnMsg の見出しコメント参照)。
    ' モーダルは起動1回だけ(WarnAtStartup 自体が mWarned で1セッション1回)。
    ' 相手のブック名は出さない。記録には件数だけ残す(名前はログにも書かない)。
    Dim wbN As Long: wbN = CohabitCount()
    If IsCohabiting(wbN) Then
        modLog.LogUsage "cohabit_detected", "startup", "同一プロセスの可視な他ブック数=" & wbN
        MsgBox CohabitWarnMsg(), vbExclamation, modAppDef.APP_NAME
    End If

    ' 2026-08-05(R17H FB-7 / B-M): 既に本棚があるのに chunk_meta が0行=R17より
    ' 前に取り込んだ資料しか無い本棚。構造検索(条文参照・俯瞰)は「無ければ
    ' 従来どおり」で静かに効かないだけなので、放っておくと新機能があることも、
    ' 有効にする方法(取り込み直し)も一生伝わらない(実機第5報のレビューB)。
    ' 起動時1回・非モーダルのトーストで1文だけ言う。モーダルにしないのは、
    ' これが「壊れている」ではなく「もっと良くできる」の知らせだから
    ' (WarnAtStartup 自体が1セッション1回=mWarned なので回数も1回)。
    If curRows > 0 And ChunkMetaRowCount() = 0 Then
        modLog.LogUsage "integrity_hint", "no_chunk_meta", "rows=" & curRows
        QueueBubble "資料を取り込み直すと新しい構造検索(条文参照・俯瞰)が" & _
            "有効になります"   ' R29H F2b: InitUI前のためバブルは保留(BootBoardでFlush)
    End If

    ' R19-1e(実機第6報①): 既存ブックの UsedRange は保存するまで縮まない。
    ' R19-1b で塗りを実使用範囲へ縮めても、それ以前に全域書式が焼き付いた
    ' ブックでは「右にも下にも無限にスクロールできる」がそのまま残るため、
    ' 利用者には「直っていない」としか見えない(調査①班1-e)。検知して、
    ' 直し方(一度保存して開き直す)を1文で伝える。
    WarnIfUsedRangeBloated

    Dim prevRows As Long: prevRows = 0
    Dim prevTxt As String: prevTxt = modState.LoadState(KEY_LAST_ROWS, "")
    If IsNumeric(prevTxt) Then prevRows = CLng(Val(prevTxt))
    If Not DataShrunk(prevRows, curRows) Then Exit Sub

    Dim prevPath As String: prevPath = modState.LoadState(KEY_LAST_PATH, "(記録なし)")
    modLog.LogUsage "integrity_warn", "rows_shrunk", _
        "前回" & prevRows & "件 → 今回" & curRows & "件 / 前回パス=" & _
        modUtil.SafeLeft(prevPath, 200)
    MsgBox ShrinkWarnMsg(prevRows, curRows, prevPath, curPath), _
        vbExclamation, modAppDef.APP_NAME
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' R19-1e: UsedRange が桁違いに広いなら、焼き付いた全域書式が残っている。
' ----------------------------------------------------------------------------
' 判定そのものは数の比較だけの純関数(modTestsPure16 が境界を固定する)。
'
' 2026-08-06(R19H FA-5(ii) / A-M⑤): 旧実装は「可視の4倍超」という【倍率】
' だった。これはチャットで必ず誤発動する: R19-1b 以降、チャットの塗りと境界は
' 会話の実下端まで【正しく】伸びるので、20往復も話せば縦は可視の4倍
' (700pt×4=2,800pt)を普通に超える。つまり「長く話した人ほど、正常に動いて
' いる画面に対して『一度保存して開き直してください』と言われ続ける」ことに
' なる(憲章§3-3の逆で、しかも保存しても消えない=直せない案内)。
' 焼き付いたブックは桁が違う(旧チャット A1:P2000=約31,200pt / 全域書式は
' 1,048,576行)ので、可視サイズと無関係な【絶対値】で切る。
'   ・縦12,000pt = 1画面700pt の約17画面ぶん。会話の長さで届く値ではない
'     (CapBubbles がバブル数を間引くため、正常な会話の下端はこの手前で頭打ち)。
'   ・横2,400pt = 一番広い画面(ダッシュボード約1,020pt)の倍以上、かつ
'     ViewportWidth のクランプ上限1,600ptより広い=正常には作れない幅。
Public Function IsUsedRangeBloated(ByVal usedW As Double, ByVal usedH As Double) As Boolean
    IsUsedRangeBloated = (usedW > BLOAT_W_PT) Or (usedH > BLOAT_H_PT)
End Function

' ----------------------------------------------------------------------------
' SwapFileWithBackup - 出来上がった tmp を dst へ差し替える1回ぶん(R33H M6)
' ----------------------------------------------------------------------------
' 呼び口は modShare.BoardWriteSummary(集約スナップショットの置き換え)。
' 置き場が基盤層なのは modShare が残509字で分岐を書けないため(憲章§4-6)。
' 「壊れた中身を読ませない・旧版を失わない」がこのモジュールの主題そのもの。
'
' 【不変条件は1本だけ】: dst が存在しない間は bak を絶対に消さない。
'   Windows の Name は宛先が既に在ると失敗するので、手順は
'   「dst を bak へ退避 → tmp を dst へ改名 → 成功したら bak を捨てる」。
'   1回目に (a)退避は成功 (b)改名は失敗 (c)復旧の改名も失敗、で抜けると
'   【旧版は bak にしか無い】。R33H F15/F16 の実装はこの状態で再試行の先頭に
'   無条件の Kill bakPath があり、唯一のコピーを自分で消していた ―― 3回とも
'   失敗すれば summary.txt が完全に消える(その日の集計だけでなく、称号の
'   累積状態も消える。累積は置き換え対象のファイルにしか無いため)。
'   なので先に dst の在否を見て、無ければ bak を宛先へ戻してから始める。
' 戻り値: 差し替えられたら True。
Public Function SwapFileWithBackup(ByVal tmpPath As String, ByVal dstPath As String, _
                                   ByVal bakPath As String) As Boolean
    On Error Resume Next
    Err.Clear
    If LenB(Dir(dstPath)) > 0 Then
        Kill bakPath            ' dst が在る=bak は前回の残骸。捨ててよい
    ElseIf LenB(Dir(bakPath)) > 0 Then
        Name bakPath As dstPath ' dst が無く bak が在る=まず旧版を戻す
    End If
    Err.Clear
    Name dstPath As bakPath     ' 宛先が無ければ 53 で失敗するだけ(初回)
    Dim hadOld As Boolean: hadOld = (Err.Number = 0)
    Err.Clear
    Name tmpPath As dstPath
    Dim swapErr As Long: swapErr = Err.Number
    Err.Clear
    If swapErr <> 0 Then
        ' 旧版を戻す。ここも失敗したら bak に残す(次回の先頭で戻す)。
        If hadOld Then Name bakPath As dstPath
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    If hadOld Then Kill bakPath
    Err.Clear
    On Error GoTo 0
    SwapFileWithBackup = True
End Function

' PauseMs - 共有I/Oの再試行のあいだだけ待つ(R33H M12 で modShare から移設。
'   置き場の理由は SwapFileWithBackup と同じ=modShare が残9字。作りは
'   modBoard.BoardWait と同型で、Timer が日跨ぎで0へ戻ったら即抜ける)。
Public Sub PauseMs(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do
    Loop
End Sub

' ----------------------------------------------------------------------------
' UsedLastRow - 使用済み範囲の最終行を測る(2026-08-16 R33H M5)。
' ----------------------------------------------------------------------------
' 測り方に2つの作法を畳んである。どちらも実測由来で、抜けると【無言で
' 間違った値】が返る種類のもの:
'   (1) 捨て読み: 行削除や条件付き書式の剥がしだけでは内部の使用済み範囲
'       (xlCellTypeLastCell)が縮まない端末がある。UsedRange を1回参照して
'       再計算させてから測る(戻り値は捨てる。modViewport.ReleaseSheetRowsBelow
'       に同じ1行がある ―― R33H M5 はその1行が新経路に無かった件)。
'   (2) 参照は1回に畳む: 「ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1」は
'       同一式の中で UsedRange を2回読む。評価順は保証されないので、伸び縮み
'       の途中で読むと存在しない行番号が出る。
' 置き場が基盤層なのは modBackdrop が残155字で分岐を書けないため(憲章§4-6)。
' 自己検算そのものが「台帳(測った値)と実データの突合」なので主題も合う。
' 失敗時は0(呼び出し側は0を「測れなかった」として扱う)。
Public Function UsedLastRow(ByVal ws As Worksheet) As Long
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    Dim discard As Long
    discard = ws.UsedRange.Rows.Count
    Err.Clear
    Dim ur As Range
    Set ur = ws.UsedRange
    If Err.Number = 0 Then
        If Not ur Is Nothing Then UsedLastRow = ur.Row + ur.Rows.Count - 1
    End If
    Err.Clear
    On Error GoTo 0
End Function

' 画面4枚(Hub/チャット/マイ本棚/ダッシュボード)のどれかが膨らんでいたら
' 1回だけ案内する(セッション1回=WarnAtStartup 自体が1回)。
Private Sub WarnIfUsedRangeBloated()
    On Error Resume Next
    ' R19H FA-5(ii): 可視サイズはもう見ない(判定が絶対値になったため)。
    ' ついでに ActiveWindow への依存も消えた ―― 他ブックが前面のときに
    ' 他人の窓を測っていた経路がここからは無くなる(FB-5 と同じ筋)。
    Dim names As Variant
    names = Array(modAppDef.SH_HOME, "Nexus", modAppDef.SH_SHELF, modAppDef.SH_NEXUS_DASH)
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim ws As Worksheet
        Set ws = Nothing
        Set ws = GetSheet(CStr(names(i)))
        If Not ws Is Nothing Then
            Dim ur As Range
            Set ur = Nothing
            Set ur = ws.UsedRange
            If Not ur Is Nothing Then
                If IsUsedRangeBloated(ur.Left + ur.Width, ur.Top + ur.Height) Then
                    modLog.LogUsage "integrity_hint", "usedrange_bloated", _
                        ws.Name & " " & CLng(ur.Left + ur.Width) & "x" & CLng(ur.Top + ur.Height)
                    QueueBubble "一度保存して開き直すと、画面のスクロール範囲が" & _
                        "正常になります"   ' R29H F2b: InitUI前のためバブルは保留(BootBoardでFlush)
                    Exit Sub
                End If
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 内部ヘルパー(シートを直接読む。上位層は呼ばない=R1)
' ----------------------------------------------------------------------------

' my_knowledge の行数(ヘッダ除く)。シートが無ければ0。
Private Function KnowledgeRowCount() As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function
    Dim lastK As Long: lastK = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastK < 2 Then Exit Function
    KnowledgeRowCount = lastK - 1
    On Error GoTo 0
End Function

' chunk_meta の行数(ヘッダ除く)。シートが無ければ0(R17H FB-7)。
' 判定材料はシート1枚だけで、上位層は1つも呼ばない(KnowledgeRowCount と同型)。
Private Function ChunkMetaRowCount() As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = GetSheet(modAppDef.SH_CHUNK_META)
    If ws Is Nothing Then Exit Function
    Dim lastM As Long: lastM = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Function
    ChunkMetaRowCount = lastM - 1
    On Error GoTo 0
End Function

' file_name(2列目)でmanifest行を探す(見つからなければ0)。
Private Function ManifestRowByName(ByVal wsM As Worksheet, ByVal fileName As String) As Long
    Dim lastM As Long: lastM = wsM.Cells(wsM.Rows.count, 1).End(xlUp).row
    If lastM < 2 Then Exit Function
    Dim arr As Variant: arr = wsM.Range(wsM.Cells(2, 2), wsM.Cells(lastM, 2)).Value
    If Not IsArray(arr) Then
        If StrComp(CStr(arr), fileName, vbTextCompare) = 0 Then ManifestRowByName = 2
        Exit Function
    End If
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(CStr(arr(i, 1)), fileName, vbTextCompare) = 0 Then
            ManifestRowByName = i + 1
            Exit Function
        End If
    Next i
End Function

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function
