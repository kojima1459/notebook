Attribute VB_Name = "modP2P"
Option Explicit

' my_stats に置く nonce 行の接頭辞(レビュー L-15)。GcOldNonces が掃除する。
Private Const NONCE_PREFIX As String = "thx:"

' ============================================================================
' modP2P - Phase 4: P2P 感謝状(EXP交換)+ AD連携ユーザーID
' ----------------------------------------------------------------------------
' 不正防止の核: 「感謝EXP(thumbup)」は "他者のExcelから共有フォルダ経由で
' 感謝状(✅解決由来)を受領した時のみ" 加算する。自分で自分の回答に✅を押しても
' 感謝EXPは一切増えない(自己申告不可)。活動EXP(質問/登録/共有)はローカルで
' 貯まり、感謝EXPと合算してレベルになる(オーナー裁定: 活動+感謝の合算型)。
'
' 仕組み:
'   ・出所: 取り込んだチャンクの origin は2つの名前空間を持つ(レビュー C-1)。
'       手渡しパック   … "pack:<作者表示名>"
'       部門チャンネル … "channel:<部門名>"
'     この出所情報から「誰へ感謝するか」を決める。
'     2026-07-31(レビュー R8 F1): ここが "pack:" しか解決していなかったため、
'     チャンネル経由で解決した人の感謝は【構造的に一度も送られていなかった】。
'     部門名は人ではないので、チャンネルは取込時に控えた発行者ID
'     ("chauth:<部門名>" → modPack が記録)を引いて宛先にする。
'     発行者IDを持たない旧い正典パックは、従来どおり静かに感謝なしで終わる。
'   ・発行: ✅解決時、直近回答の最上位引用ソースの作者(自分以外)へ、共有フォルダ
'     <nexus_share_path>\thanks\ に感謝状(1件1ファイル・nonce付きTSV)を書く。
'   ・受領: 同期/起動時に thanks\ を走査し、宛先=自分かつ未処理(nonce重複排除)
'     の感謝状ごとに感謝EXP(+exp_thumbup)を加算する。
'
' 実行環境: Windows版Excelのみ。ADSystemInfo/ADODB.Stream/共有フォルダは実機
' 前提のため、本モジュールは実機(2台+共有フォルダ)での結合テストが必要。
' 本リポジトリのlint/LO構文チェックは構文のみ検証(実行検証は不可)。
' ============================================================================

Private Const THANKS_SUBDIR As String = "thanks"   ' サブフォルダ名(区切り"\"は連結時に付与)

' 届いた「ありがとう」を会話で伝えるための控え(NoticeTextが引き取ると消える)。
' VBAはモジュールレベル宣言をプロシージャより前に置く必要がある。
Private Const MAX_NOTICE As Long = 3
Private mThanksFrom(0 To 2) As String
Private mThanksSrc(0 To 2) As String
Private mThanksN As Long
Private Const NOISE_SUBDIR As String = "noise"      ' サブフォルダ名(区切り"\"は連結時に付与)
Private mUserIdCache As String

' ----------------------------------------------------------------------------
' CurrentUserId - AD(ADSystemInfo)のCN → 失敗時 USERNAME。共有内での宛先/差出人。
' ----------------------------------------------------------------------------
Public Function CurrentUserId() As String
    If LenB(mUserIdCache) > 0 Then
        CurrentUserId = mUserIdCache
        Exit Function
    End If

    ' 2026-07-31 R11-D(監査3 M-1): ADSystemInfo は「失敗しても静かに
    ' USERNAME へ落ちる」設計だが、落ちたこと自体がどこにも残らなかった。
    ' 感謝状の宛先IDが端末ごとに AD の CN と USERNAME で食い違うと、送った
    ' 感謝が誰にも届かない(誰も気付けない)。落ちたら1行残す。
    ' もう1つの罠が【遅延】: 切断されたドメインでは ADSystemInfo の生成が
    ' 数十秒ブロックすることがあり、起動が固まったように見える。閾値2秒を
    ' 超えたら、成功していても1行残す(次の調査の起点になる)。
    Dim uid As String
    Dim adT0 As Double: adT0 = Timer
    Dim adErrNum As Long: adErrNum = 0
    Dim adErrDesc As String: adErrDesc = ""
    On Error Resume Next
    Dim adsi As Object
    Set adsi = CreateObject("ADSystemInfo")
    uid = modP2PIo.CnFromDn(CStr(adsi.UserName))   ' 例 "CN=山田 太郎,OU=..,DC=.."
    adErrNum = Err.Number
    adErrDesc = Err.Description
    Err.Clear
    Set adsi = Nothing                    ' COM解放(正常・異常ともOn Error Resume Next配下)
    On Error GoTo 0

    Dim adElapsed As Double: adElapsed = Timer - adT0
    If adElapsed < 0 Then adElapsed = 0        ' 深夜0時のTimerロールオーバー
    On Error Resume Next
    If LenB(uid) = 0 Then
        modLog.LogUsage "p2p_userid_fallback", "", "ADSystemInfo失敗のためUSERNAMEを使用 " & _
            "err#" & adErrNum & ": " & modUtil.SafeLeft(adErrDesc, 200) & _
            " (" & Format$(adElapsed, "0.0") & "秒)"
    End If
    If adElapsed >= 2# Then
        modLog.LogUsage "p2p_userid_slow", "", "ADSystemInfo の応答に " & _
            Format$(adElapsed, "0.0") & "秒かかりました(閾値2秒)"
    End If
    On Error GoTo 0

    If LenB(uid) = 0 Then uid = Environ$("USERNAME")
    If LenB(uid) = 0 Then uid = "user"
    uid = modP2PIo.SanitizeId(uid)

    mUserIdCache = uid
    CurrentUserId = uid
End Function


' ----------------------------------------------------------------------------
' EmitThanksForLastAnswer - ✅解決時に modAsk.FeedbackGreen から呼ぶ。
'   直近回答の最上位ソースの作者(自分以外・出所判明)へ感謝状を書き出す。
' ----------------------------------------------------------------------------
Public Sub EmitThanksForLastAnswer()
    On Error Resume Next
    EmitThanks modAsk.LastTopSource()
    On Error GoTo 0
End Sub

Public Sub EmitThanks(ByVal topSource As String)
    On Error GoTo Done
    If LenB(topSource) = 0 Then Exit Sub

    ' 出所(origin)をそのまま取り出し、名前空間ごとに宛先を解決する。
    ' 2026-07-31(レビュー R8 F1): 以前はここで "pack:" 接頭辞だけを剥がして
    ' いたため、origin="channel:<部門名>" は空文字になり即Exitしていた。
    Dim origin As String: origin = OriginOfSource(topSource)
    Dim author As String: author = modShareRule.OriginName(origin)
    If LenB(author) = 0 Then Exit Sub                          ' 自作/出所不明 → 感謝なし
    If StrComp(author, "不明", vbTextCompare) = 0 Then Exit Sub

    Dim myId As String: myId = CurrentUserId()
    If StrComp(author, myId, vbTextCompare) = 0 Then Exit Sub  ' 自分自身には送らない

    ' 盲点D5: ファイル名・照合キーは必ずSanitizeIdを通す。作者名にWindowsの
    ' ファイル名禁止文字(\ / : * ? " < > | , タブ/改行)が混じると、生のままでは
    ' (a) 共有フォルダでのファイル作成が実行時エラーになる、(b) 受信側は sanitize済の
    '     自分IDで thx_<myId>_*.txt を集めるため前方一致が崩れ感謝EXPが永久に届かない。
    ' そこで送信側も宛先キーをsanitizeし、ファイル名・payloadの綴りを両側で一致させる。
    ' 2026-07-28(レビュー H-5): 宛先キーは「作者のID」を使う。
    ' 以前は表示名(初回起動で本人が打つ pack_author)をそのまま宛先に
    ' していたが、受け取る側の照合キーは AD の CN または %USERNAME% で、
    ' 両者が偶然一致しない限り感謝は誰にも届かなかった。
    ' パック取込時に控えた 表示名→ID の対応(my_stats "pkauth:")を引く。
    ' 対応が無い(author_id を持たない旧いパック)ときだけ、従来どおり
    ' 表示名で宛先を作る。
    ' 2026-07-31(レビュー R8 F1): チャンネル("channel:<部門名>")は
    ' 表示名での代用を許さない。部門名は人ではないので、代用すると
    ' 「人事部」宛の感謝状が誰にも受け取られず共有フォルダに溜まるだけになる。
    Dim resolvedId As String
    resolvedId = ResolveAuthorId(origin, author)
    If LenB(resolvedId) = 0 Then Exit Sub
    ' 解決したIDが自分なら送らない(自部門の正典を自分で発行している人)。
    If StrComp(resolvedId, myId, vbTextCompare) = 0 Then Exit Sub

    Dim authorKey As String
    authorKey = modP2PIo.SanitizeId(resolvedId)
    If LenB(authorKey) = 0 Then Exit Sub

    Dim folderPath As String: folderPath = ThanksDir()
    If LenB(folderPath) = 0 Then Exit Sub
    modP2PIo.EnsureDir folderPath

    ' MAX_PATH対策(SRE監査Phase2.2): ファイル名に載せるID成分は Fnv1a64Hex(16桁)へ
    ' 圧縮する。SanitizeIdは最大64字で、深い共有UNCパス(例 \\host\部\課\...)配下では
    ' thx_<author64>_<myId64...> が260字を超えて Dir/Kill/SaveToFile がクラッシュし得る。
    ' payload(TSV)側は生の sanitize済ID のまま保持し、受信側 CollectThanks の
    ' StrComp(f(2), myId) 照合を壊さない(ファイル名だけを短縮する)。
    Dim nonce As String: nonce = modP2PIo.NewNonce(modUtil.Fnv1a64Hex(myId))
    Dim rowText As String
    rowText = nonce & vbTab & myId & vbTab & authorKey & vbTab & _
              modP2PIo.SanitizeField(topSource) & vbTab & modUtil.NowStamp()

    ' ネットワークドライブのロック(実行時エラー70等)に耐えるリトライ書込み
    If modP2PIo.WriteUtf8Retry(folderPath & "thx_" & modUtil.Fnv1a64Hex(authorKey) & "_" & nonce & ".txt", rowText, "modP2P.EmitThanks") Then
        On Error Resume Next
        modLog.LogUsage "thanks_emit", "", "to=" & author & " src=" & modUtil.SafeLeft(topSource, 120)
        On Error GoTo 0
    End If
Done:
End Sub

' origin から、感謝状の宛先に使うIDを解決する。
'   pack:<表示名>   … my_stats "pkauth:<表示名>" を引く。無ければ表示名で代用
'                      (author_id を持たない旧いパックとの過渡期互換。H-5)
'   channel:<部門名> … my_stats "chauth:<部門名>" を引く。無ければ空文字
'                      (代用不可。裁定 R8 F1: 旧い正典は静かに感謝なし)
' キーの綴りは modShareRule.AuthorStatKey が唯一の決定者。書く側(modPack)と
' 読む側(ここ)でずれると感謝が永久に届かないため、必ず同じ関数を通す。
Private Function ResolveAuthorId(ByVal origin As String, ByVal authorName As String) As String
    Dim key As String: key = modShareRule.AuthorStatKey(origin)
    If LenB(key) = 0 Then Exit Function

    Dim v As String
    On Error Resume Next
    v = Trim$(modStats.GetStatText(key))
    On Error GoTo 0
    If LenB(v) > 0 Then
        ResolveAuthorId = v
        Exit Function
    End If
    If modShareRule.NeedsResolvedId(origin) Then Exit Function
    ResolveAuthorId = authorName
End Function


' ----------------------------------------------------------------------------
' CollectThanks - 同期/起動時に呼ぶ。宛先=自分の未処理感謝状ごとに感謝EXP加算。
'   戻り値=今回付与した件数(dedupはnonceで実施)。
' ----------------------------------------------------------------------------
Public Function CollectThanks(Optional ByVal silent As Boolean = False) As Long
    On Error GoTo Done
    Dim folderPath As String: folderPath = ThanksDir()
    If LenB(folderPath) = 0 Then Exit Function

    Dim myId As String: myId = CurrentUserId()
    Dim awarded As Long: awarded = 0

    ' 1) 宛先=自分のファイル名を先に全部集める。Dir()列挙中にKill(GC)すると
    '    列挙がスキップ・破綻するため、必ず「集めてから処理」にする。
    Dim names() As String: ReDim names(0 To 63)
    Dim nFiles As Long: nFiles = 0
    ' MAX_PATH対策(Phase2.2)で送信側はファイル名の宛先IDをFnv1a64Hexで綴るため、
    ' 受信側の前方一致も同じハッシュで揃える(payload f(2)は生IDのままStrComp照合)。
    Dim fn As String: fn = Dir(folderPath & "thx_" & modUtil.Fnv1a64Hex(myId) & "_*.txt")
    Do While LenB(fn) > 0
        If nFiles > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 64)
        names(nFiles) = fn
        nFiles = nFiles + 1
        fn = Dir()   ' このループ内で他のDir()を呼ばないこと
    Loop

    ' 2026-07-28(レビュー H-5)の過渡期対応: 修正前のビルドから送られた
    ' 感謝状は、宛先が「表示名」でハッシュされている。自分の表示名でも
    ' 一度だけ集めて拾う(表示名とIDが同じ端末では二重に集まるが、
    ' nonce の重複排除があるのでEXPは二重加算されない)。
    Dim myName As String
    On Error Resume Next
    myName = modP2PIo.SanitizeId(Trim$(modConfig.GetString("pack_author", "")))
    On Error GoTo Done
    If LenB(myName) > 0 Then
        If StrComp(myName, myId, vbTextCompare) <> 0 Then
            fn = Dir(folderPath & "thx_" & modUtil.Fnv1a64Hex(myName) & "_*.txt")
            Do While LenB(fn) > 0
                If nFiles > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 64)
                names(nFiles) = fn
                nFiles = nFiles + 1
                fn = Dir()
            Loop
        End If
    End If

    ' 2) 処理: リトライ読取り→加点(nonce dedup)→GC(処理済みは削除してDir肥大化防止)
    Dim i As Long
    For i = 0 To nFiles - 1
        Dim full As String: full = folderPath & names(i)
        Dim rec As String
        If modP2PIo.ReadUtf8Retry(full, rec, "modP2P.CollectThanks") Then
            Dim f() As String: f = Split(rec, vbTab)
            If UBound(f) >= 4 Then
                Dim nonce As String: nonce = f(0)
                ' 宛先はIDか、過渡期は表示名(上のコメント参照)。
                Dim isForMe As Boolean
                isForMe = (StrComp(f(2), myId, vbTextCompare) = 0)
                If Not isForMe And LenB(myName) > 0 Then
                    isForMe = (StrComp(f(2), myName, vbTextCompare) = 0)
                End If
                If isForMe Then
                    If Not SeenNonce(nonce) Then
                        MarkNonce nonce
                        modStats.AddExp "thumbup"
                        modStats.Bump "thanks_received_total"
                        On Error Resume Next
                        modLog.LogUsage "thanks_recv", "", "from=" & f(1) & " src=" & modUtil.SafeLeft(f(3), 120)
                        ' 誰が・どの資料で解決したかを控える。EXPの数字ではなく
                        ' この2つが、届けるべき中身そのもの(NoticeText)。
                        RememberThanks f(1), f(3)
                        On Error GoTo Done
                        awarded = awarded + 1
                    End If
                    ' 自分宛の処理済み(または既知)ファイルはGC。失敗しても
                    ' nonce重複排除があるので二重加算にはならない(次回再スキップ)。
                    modP2PIo.KillRetry full
                End If
            End If
        End If
    Next i

    ' 2026-07-27: ここにあった祝福MsgBoxは到達不能コードだった。呼び出し元は
    ' modBoot・modShelfSync とも全て silent:=True で、silent:=False は
    ' ソース全体に1か所も無かった。つまり「あなたの資料で誰かが解決した」という、
    ' このアプリで唯一の人対人の瞬間が、誰にも届かないまま消えていた。
    ' ファイル数を数えるバッジにはモーダルが出るのに、である。
    '
    ' 起動処理の途中でダイアログを重ねるのは元の設計判断として正しいので、
    ' MsgBoxは復活させない。代わりに内容を控えておき(NoticeText)、チャットが
    ' 描き終わったあとに会話の中で伝える。相手の名前と資料名が本体で、
    ' 感謝EXPの数字はそこに要らない。
    CollectThanks = awarded
    ' 宛先不明の感謝状のGC。宛先の綴りが変わった端末や、退職・異動で
    ' もう誰も取りに来ないファイルは、放置すると共有フォルダに無限に
    ' 溜まり、Dir列挙が毎起動で重くなる(レビュー H-5)。
    ' 既定60日。0以下でGC無効(config thanks_gc_days)。
    GcOldThanks folderPath
    GcOldNonces
Done:
End Function

' 作成からN日以上経った thx_ ファイルを消す。自分宛かどうかは見ない
' (自分宛は上の処理で読まれた時点で消えているため、ここに残っているのは
'  【まだ受け取りに来ていない】ファイル)。列挙中にKillすると列挙が壊れるので、
' 必ず「集めてから消す」。
'
' 2026-07-31(レビュー R8 F14): 保持を30日から60日へ延ばした。
' 以前のコメントは「退職・異動でもう誰も取りに来ないファイル」と書いていたが、
' これは前提が間違っている。ここに残っているファイルの大半は
' 「宛先の人がまだこの30日間ブックを開いていない」だけで、育休・長期出張・
' 長期休職なら普通に起こる。感謝は届かなかったことにも気付けないまま消える。
' 共有フォルダのDir列挙が重くなるのを避ける目的は60日でも十分達成できる。
Private Sub GcOldThanks(ByVal folderPath As String)
    On Error Resume Next
    Dim days As Long: days = modConfig.GetLong("thanks_gc_days", 60)
    If days <= 0 Then Exit Sub

    Dim old() As String: ReDim old(0 To 63)
    Dim n As Long
    Dim limit As Date: limit = DateAdd("d", -days, Now)

    Dim fn As String: fn = Dir(folderPath & "thx_*.txt")
    Do While LenB(fn) > 0
        If n > UBound(old) Then ReDim Preserve old(0 To UBound(old) + 64)
        old(n) = fn
        n = n + 1
        fn = Dir()
    Loop

    Dim i As Long
    For i = 0 To n - 1
        Dim full As String: full = folderPath & old(i)
        Dim stamp As Date
        Err.Clear
        stamp = FileDateTime(full)
        If Err.Number = 0 Then
            If stamp < limit Then modP2PIo.KillRetry full
        End If
        Err.Clear
    Next i
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 届いた「ありがとう」を、会話で伝えるための控え
' ----------------------------------------------------------------------------
' 起動中に集めた分をモジュール変数に貯め、UI側(modApp.LaunchNexus)が
' 描画完了後に1回だけ引き取る。引き取ったら消える(二度は出さない)。

Private Sub RememberThanks(ByVal fromName As String, ByVal srcName As String)
    If mThanksN >= MAX_NOTICE Then
        mThanksN = mThanksN + 1     ' 件数だけは数え続ける(「ほか N件」用)
        Exit Sub
    End If
    mThanksFrom(mThanksN) = Trim$(fromName)
    mThanksSrc(mThanksN) = Trim$(srcName)
    mThanksN = mThanksN + 1
End Sub

' NoticeText - 会話に出す文面を返して控えを消す。無ければ空文字。
Public Function NoticeText() As String
    If mThanksN <= 0 Then Exit Function

    Dim sb As String
    sb = ChrW(&HD83C) & ChrW(&HDF89) & " あなたが登録した資料で、他の方の疑問が解決しました。"

    Dim shown As Long
    shown = mThanksN
    If shown > MAX_NOTICE Then shown = MAX_NOTICE

    Dim i As Long
    For i = 0 To shown - 1
        Dim who As String: who = mThanksFrom(i)
        If LenB(who) = 0 Then who = "どなたか"
        Dim what As String: what = mThanksSrc(i)
        If LenB(what) = 0 Then
            sb = sb & vbLf & "　・" & who & " さん"
        Else
            sb = sb & vbLf & "　・" & who & " さん ―「" & modUtil.SafeLeft(what, 40) & "」"
        End If
    Next i
    If mThanksN > shown Then sb = sb & vbLf & "　　ほか " & (mThanksN - shown) & "件"

    sb = sb & vbLf & "資料を入れておくと、こうして自分がいない場面でも誰かの助けになります。"

    mThanksN = 0
    NoticeText = sb
End Function

' ----------------------------------------------------------------------------
' EmitNoiseVote - 「品質報告(ノイズ投票)」を共有フォルダへ書き出す。
'   投票者×ソースにつき1ファイル(再投票は上書き=重複排除)。
' ----------------------------------------------------------------------------
Public Sub EmitNoiseVote(ByVal source As String)
    On Error GoTo Done
    If LenB(source) = 0 Then Exit Sub

    Dim folderPath As String: folderPath = NoiseDir()
    If LenB(folderPath) = 0 Then Exit Sub
    modP2PIo.EnsureDir folderPath

    Dim myId As String: myId = CurrentUserId()
    Dim srcHash As String: srcHash = modUtil.Fnv1a64Hex(source)

    Dim content As String
    content = myId & vbTab & modP2PIo.SanitizeField(source) & vbTab & modUtil.NowStamp()

    ' MAX_PATH対策(Phase2.2): reporter IDもFnv1a64Hex(16桁)で綴る。(reporter,source)毎に
    ' 決定的なので再投票は同一ファイルを上書き=重複排除は不変。集計/GCはsrcHash前方一致で
    ' 行うためreporter部の綴りには非依存(payloadに生myIdを保持し報告者一覧は復元可能)。
    Dim votePath As String
    votePath = folderPath & "noise_" & srcHash & "_" & modUtil.Fnv1a64Hex(myId) & ".txt"
    ' 3回リトライしても書けなかった場合のE0705記録はWriteUtf8Retry内で一元化済み
    ' (重複排除により二重投票にはならないので、ここでは追加の対応は不要)。
    modP2PIo.WriteUtf8Retry votePath, content, "modP2P.EmitNoiseVote"
Done:
End Sub

' ----------------------------------------------------------------------------
' CollectNoiseVotes - 共有フォルダの投票を集計し、組織的除外(gexcl)を
'   ローカルで再計算する。戻り値=今回の組織的除外件数。
'   注意: 投票/解除ファイルは削除しない(毎回の同期で再計算するため)。
' ----------------------------------------------------------------------------
Public Function CollectNoiseVotes(Optional ByVal silent As Boolean = False) As Long
    On Error GoTo Done
    Dim folderPath As String: folderPath = NoiseDir()
    If LenB(folderPath) = 0 Then Exit Function

    ' fail-safe: 共有ルートが到達不能なとき(一時的なNW断)は「投票ゼロ」と区別
    ' できないため、ResetGlobalExcludedで既存の組織的除外を消さない=前回状態を保持。
    ' 共有は到達できるが noise\ が未作成(まだ誰も報告していない)なら通常どおり
    ' 集計する(=正当に空へ再計算)。
    '
    ' 2026-07-31(レビュー R8 F6): ここには「末尾"\"無しのパスを Dir で見る」
    ' 独自プローブがあった。判定式が modShare と違ううえ、空の共有ルートを
    ' 到達不能と誤判定する(F6の本題)。到達判定は modShare に一本化する。
    ' NoiseDir() が空でない時点で modShare.Reachable() は真だが、意図
    '(「共有そのものへ届くか」を確かめてから集計する)を残すため明示する。
    If Not modShare.Reachable() Then Exit Function   ' 共有到達不能 → 除外状態を保持して撤退

    ' 1) 確定除外フラグ(gexcl_*.txt)を集める。これは閾値到達で確定し、個別投票を
    '    GC(削除)して1ファイルへ圧縮した「軽量な確定除外フラグ」。中身は
    '    source\t確定日時\t報告者一覧(監査用)。フラグ有り=確定除外(票の再集計不要)。
    Dim flagNames() As String: ReDim flagNames(0 To 63)
    Dim nFlags As Long: nFlags = 0
    Dim fn As String: fn = Dir(folderPath & "gexcl_*.txt")
    Do While LenB(fn) > 0
        If nFlags > UBound(flagNames) Then ReDim Preserve flagNames(0 To UBound(flagNames) + 64)
        flagNames(nFlags) = fn
        nFlags = nFlags + 1
        fn = Dir()   ' このループ内で他のDir()を呼ばないこと
    Loop

    Dim confirmed As Object: Set confirmed = CreateObject("Scripting.Dictionary")
    confirmed.CompareMode = 0   ' vbBinaryCompare
    Dim i As Long
    For i = 0 To nFlags - 1
        Dim frec As String
        If modP2PIo.ReadUtf8Retry(folderPath & flagNames(i), frec, "modP2P.CollectNoiseVotes") Then
            Dim ff() As String: ff = Split(frec, vbTab)
            If UBound(ff) >= 0 Then
                If LenB(ff(0)) > 0 Then
                    If Not confirmed.Exists(ff(0)) Then confirmed.Add ff(0), True
                End If
            End If
        End If
    Next i

    ' 2) 未確定のノイズ投票(noise_*.txt)を集める→ソース毎の異なる報告者数+一覧
    '    (ファイル名がreporter×source単位で一意なため、ファイル数=異なる報告者数)
    Dim voteNames() As String: ReDim voteNames(0 To 63)
    Dim nVotes As Long: nVotes = 0
    fn = Dir(folderPath & "noise_*.txt")
    Do While LenB(fn) > 0
        If nVotes > UBound(voteNames) Then ReDim Preserve voteNames(0 To UBound(voteNames) + 64)
        voteNames(nVotes) = fn
        nVotes = nVotes + 1
        fn = Dir()   ' このループ内で他のDir()を呼ばないこと
    Loop

    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary"): counts.CompareMode = 0
    Dim reporters As Object: Set reporters = CreateObject("Scripting.Dictionary"): reporters.CompareMode = 0
    For i = 0 To nVotes - 1
        Dim vrec As String
        If modP2PIo.ReadUtf8Retry(folderPath & voteNames(i), vrec, "modP2P.CollectNoiseVotes") Then
            Dim vf() As String: vf = Split(vrec, vbTab)
            If UBound(vf) >= 1 Then
                Dim who As String: who = vf(0)
                Dim src As String: src = vf(1)
                If counts.Exists(src) Then
                    counts(src) = counts(src) + 1
                    reporters(src) = CStr(reporters(src)) & "," & who
                Else
                    counts.Add src, 1
                    reporters.Add src, who
                End If
            End If
        End If
    Next i

    ' 3) 適用: 全解除→(確定フラグ有 または 票数>=閾値)のソースを組織的除外。
    '    新規確定(フラグ無し&票数>=閾値)はフラグを書いて個別投票をGC(圧縮)。
    modStats.ResetGlobalExcluded
    Dim threshold As Long: threshold = modStats.NoiseThreshold()

    Dim excluded As Object: Set excluded = CreateObject("Scripting.Dictionary"): excluded.CompareMode = 0
    Dim k As Variant
    For Each k In confirmed.Keys
        If Not excluded.Exists(CStr(k)) Then excluded.Add CStr(k), True
    Next k
    For Each k In counts.Keys
        If CLng(counts(k)) >= threshold Then
            If Not excluded.Exists(CStr(k)) Then excluded.Add CStr(k), True
        End If
    Next k

    Dim awarded As Long: awarded = 0
    For Each k In excluded.Keys
        Dim s As String: s = CStr(k)
        modStats.MarkGlobalExcluded s
        awarded = awarded + 1
        If Not confirmed.Exists(s) Then
            ' 新規確定: フラグを書いて投票を圧縮(報告者一覧を監査用に保持)
            Dim repList As String
            repList = ""
            If reporters.Exists(s) Then repList = CStr(reporters(s))
            modP2PIo.WriteUtf8Retry folderPath & "gexcl_" & modUtil.Fnv1a64Hex(s) & ".txt", _
                s & vbTab & modUtil.NowStamp() & vbTab & modP2PIo.SanitizeField(repList), "modP2P.CollectNoiseVotes"
        End If
        GcNoiseVotesForSource folderPath, s   ' 個別投票をGC(Dir肥大化=遅延を防止)
    Next k

    If awarded > 0 And Not silent Then
        MsgBox "組織で" & awarded & "件のナレッジが検索対象から除外されています" & _
               "(品質報告の集計結果)。", vbInformation, modAppDef.APP_NAME
    End If
    CollectNoiseVotes = awarded
Done:
End Function

' 指定ソースの個別投票(noise_<hash>_*.txt)を共有フォルダから削除(GC)。
' 列挙中の削除でスキップが起きないよう「集めてから削除」。
Private Sub GcNoiseVotesForSource(ByVal folderPath As String, ByVal source As String)
    On Error GoTo Done
    Dim hash As String: hash = modUtil.Fnv1a64Hex(source)
    Dim names() As String: ReDim names(0 To 63)
    Dim n As Long: n = 0
    Dim fn As String: fn = Dir(folderPath & "noise_" & hash & "_*.txt")
    Do While LenB(fn) > 0
        If n > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 64)
        names(n) = fn
        n = n + 1
        fn = Dir()
    Loop
    Dim i As Long
    For i = 0 To n - 1
        modP2PIo.KillRetry folderPath & names(i)
    Next i
Done:
End Sub

' ----------------------------------------------------------------------------
' ClearNoise - 管理者専用。指定ソースの組織的除外を「復帰」する。確定フラグと
'   個別投票を共有フォルダから削除するため、全ユーザーの次回同期で除外が解ける。
'   (個人ミュートは本人の意思として各自ローカルに残る。)
' ----------------------------------------------------------------------------
Public Function ClearNoise(ByVal source As String) As Boolean
    On Error GoTo Done
    If Not IsAdmin() Then Exit Function
    If LenB(source) = 0 Then Exit Function

    Dim folderPath As String: folderPath = NoiseDir()
    If LenB(folderPath) = 0 Then Exit Function

    Dim srcHash As String: srcHash = modUtil.Fnv1a64Hex(source)
    modP2PIo.KillRetry folderPath & "gexcl_" & srcHash & ".txt"   ' 確定フラグを削除
    GcNoiseVotesForSource folderPath, source              ' 残存する個別投票も削除

    On Error Resume Next
    modLog.LogUsage "noise_cleared", "", "by=" & CurrentUserId() & " src=" & modUtil.SafeLeft(source, 120)
    On Error GoTo Done
    ClearNoise = True
Done:
End Function

' ----------------------------------------------------------------------------
' IsAdmin - config "admin_users"(カンマ区切りADID)に自分が含まれるか。
' ----------------------------------------------------------------------------
Public Function IsAdmin() As Boolean
    IsAdmin = IsAdminId(CurrentUserId())
End Function

Private Function IsAdminId(ByVal candidate As String) As Boolean
    On Error GoTo Done
    Dim list As String: list = modConfig.GetString("admin_users", "")
    If LenB(list) = 0 Then Exit Function
    If LenB(candidate) = 0 Then Exit Function

    Dim parts() As String: parts = Split(list, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If StrComp(Trim$(parts(i)), Trim$(candidate), vbTextCompare) = 0 Then
            IsAdminId = True
            Exit Function
        End If
    Next i
Done:
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' ソース名 → origin セルの生値("pack:<作者>" / "channel:<部門名>" / "self" 等)。
' 名前空間の解釈は modShareRule に一本化してある(2026-07-31 R8 F1)。
' 以前はここで "pack:" だけを剥がしていたため、チャンネル由来のソースは
' 常に空文字になり、感謝状が一度も発行されなかった。
Private Function OriginOfSource(ByVal srcName As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastR
        If StrComp(CStr(ws.Cells(r, 2).Value), srcName, vbTextCompare) = 0 Then
            OriginOfSource = CStr(ws.Cells(r, 3).Value)
            Exit Function
        End If
    Next r
End Function

' 到達性の判定は modShare が1セッション1回だけ行う(レビュー M-20)。
' 届かない共有に毎回挨拶しに行くと、起動と終了が数十秒ブロックする。
Private Function ThanksDir() As String
    ThanksDir = modShare.SubDir(THANKS_SUBDIR)
End Function

' 到達性の判定は modShare が1セッション1回だけ行う(レビュー M-20)。
' 届かない共有に毎回挨拶しに行くと、起動と終了が数十秒ブロックする。
Private Function NoiseDir() As String
    NoiseDir = modShare.SubDir(NOISE_SUBDIR)
End Function


' 2026-07-28(レビュー L-15): nonce の記録に日付を持たせ、古いものを掃除する。
'
' 従来は my_stats に "thx:<nonce>" 行が無限に溜まり、
' modStats.FindKeyRow は線形探索なので、使い込むほど全ての統計参照が
' 遅くなっていた(感謝が届くほど遅くなる、という逆向きの設計)。
' nonce の役目は「同じ感謝状を二度数えない」ことだけで、送信側が
' thanks_gc_days 日で共有フォルダのファイルを掃除する以上、
' それより長く覚えている必要が無い。
Private Function SeenNonce(ByVal nonce As String) As Boolean
    SeenNonce = (LenB(modStats.GetStatText(NONCE_PREFIX & nonce)) > 0)
End Function

Private Sub MarkNonce(ByVal nonce As String)
    modStats.SetStatText NONCE_PREFIX & nonce, Format$(Date, "yyyy-mm-dd")
End Sub

' 期限を過ぎた nonce 行を my_stats から取り除く(CollectThanks の最後に呼ぶ)。
' 保持日数は共有フォルダのGCと同じ既定60日＋余裕7日(R8 F14で30→60)。
Private Sub GcOldNonces()
    On Error Resume Next
    Dim keepDays As Long
    keepDays = modConfig.GetLong("thanks_gc_days", 60) + 7
    If keepDays < 1 Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_STATS)
    If ws Is Nothing Then Exit Sub
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastR < 2 Then Exit Sub

    Dim limit As Date: limit = DateAdd("d", -keepDays, Date)
    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 2)).Value

    ' 下から消す(上から消すと行番号がずれる)。
    Dim i As Long
    For i = UBound(arr, 1) To LBound(arr, 1) Step -1
        Dim k As String: k = CStr(arr(i, 1))
        If Left$(k, Len(NONCE_PREFIX)) = NONCE_PREFIX Then
            Dim v As String: v = Trim$(CStr(arr(i, 2)))
            Dim drop As Boolean: drop = True        ' 日付が読めない旧形式も掃除対象
            If LenB(v) > 0 Then
                Err.Clear
                Dim d As Date: d = CDate(v)
                If Err.Number = 0 Then drop = (d < limit)
                Err.Clear
            End If
            If drop Then ws.Rows(i + 1).Delete
        End If
    Next i
    On Error GoTo 0
End Sub


' ネットワークドライブのロック(実行時エラー70/55/75等)に耐えるリトライI/O。
' 失敗時は短時間バックオフ待機して最大3回試行。全滅でも例外は出さず False。


