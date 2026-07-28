Attribute VB_Name = "modP2P"
Option Explicit

' ============================================================================
' modP2P - Phase 4: P2P 感謝状(EXP交換)+ AD連携ユーザーID
' ----------------------------------------------------------------------------
' 不正防止の核: 「感謝EXP(thumbup)」は "他者のExcelから共有フォルダ経由で
' 感謝状(✅解決由来)を受領した時のみ" 加算する。自分で自分の回答に✅を押しても
' 感謝EXPは一切増えない(自己申告不可)。活動EXP(質問/登録/共有)はローカルで
' 貯まり、感謝EXPと合算してレベルになる(オーナー裁定: 活動+感謝の合算型)。
'
' 仕組み:
'   ・出所: パック取込時に各チャンクの origin が "pack:<作者>" になっている
'     (modPack実装済み)。この作者情報をそのまま「誰へ感謝するか」に使う。
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

    Dim uid As String
    On Error Resume Next
    Dim adsi As Object
    Set adsi = CreateObject("ADSystemInfo")
    uid = CnFromDn(CStr(adsi.UserName))   ' 例 "CN=山田 太郎,OU=..,DC=.."
    Set adsi = Nothing                    ' COM解放(正常・異常ともOn Error Resume Next配下)
    On Error GoTo 0

    If LenB(uid) = 0 Then uid = Environ$("USERNAME")
    If LenB(uid) = 0 Then uid = "user"
    uid = SanitizeId(uid)

    mUserIdCache = uid
    CurrentUserId = uid
End Function

Private Function CnFromDn(ByVal dn As String) As String
    If LenB(dn) = 0 Then Exit Function
    Dim p As Long: p = InStr(1, dn, "CN=", vbTextCompare)
    If p = 0 Then Exit Function
    Dim rest As String: rest = Mid$(dn, p + 3)
    Dim c As Long: c = InStr(rest, ",")
    If c > 0 Then rest = Left$(rest, c - 1)
    CnFromDn = Trim$(rest)
End Function

Private Function SanitizeId(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", ",", vbTab, vbCr, vbLf)
    Dim t As String: t = s
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i
    SanitizeId = modUtil.SafeLeft(Trim$(t), 64)
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

    Dim author As String: author = AuthorOfSource(topSource)
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
    Dim authorKey As String
    authorKey = SanitizeId(ResolveAuthorId(author))
    If LenB(authorKey) = 0 Then Exit Sub

    Dim folderPath As String: folderPath = ThanksDir()
    If LenB(folderPath) = 0 Then Exit Sub
    EnsureDir folderPath

    ' MAX_PATH対策(SRE監査Phase2.2): ファイル名に載せるID成分は Fnv1a64Hex(16桁)へ
    ' 圧縮する。SanitizeIdは最大64字で、深い共有UNCパス(例 \\host\部\課\...)配下では
    ' thx_<author64>_<myId64...> が260字を超えて Dir/Kill/SaveToFile がクラッシュし得る。
    ' payload(TSV)側は生の sanitize済ID のまま保持し、受信側 CollectThanks の
    ' StrComp(f(2), myId) 照合を壊さない(ファイル名だけを短縮する)。
    Dim nonce As String: nonce = NewNonce(modUtil.Fnv1a64Hex(myId))
    Dim rowText As String
    rowText = nonce & vbTab & myId & vbTab & authorKey & vbTab & _
              SanitizeField(topSource) & vbTab & modUtil.NowStamp()

    ' ネットワークドライブのロック(実行時エラー70等)に耐えるリトライ書込み
    If WriteUtf8Retry(folderPath & "thx_" & modUtil.Fnv1a64Hex(authorKey) & "_" & nonce & ".txt", rowText, "modP2P.EmitThanks") Then
        On Error Resume Next
        modLog.LogUsage "thanks_emit", "", "to=" & author & " src=" & modUtil.SafeLeft(topSource, 120)
        On Error GoTo 0
    End If
Done:
End Sub

' 表示名から、感謝状の宛先に使うIDを解決する。
' パック取込時に my_stats へ控えた "pkauth:<表示名>" を引き、無ければ
' 表示名をそのまま返す(author_id を持たない旧いパックとの過渡期互換)。
Private Function ResolveAuthorId(ByVal authorName As String) As String
    ResolveAuthorId = authorName
    On Error Resume Next
    Dim v As String
    v = Trim$(modStats.GetStatText("pkauth:" & LCase$(Trim$(authorName))))
    If LenB(v) > 0 Then ResolveAuthorId = v
    On Error GoTo 0
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
    myName = SanitizeId(Trim$(modConfig.GetString("pack_author", "")))
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
        If ReadUtf8Retry(full, rec, "modP2P.CollectThanks") Then
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
                    KillRetry full
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
    ' 既定30日。0以下でGC無効(config thanks_gc_days)。
    GcOldThanks folderPath
Done:
End Function

' 作成からN日以上経った thx_ ファイルを消す。自分宛かどうかは見ない
' (自分宛は上の処理で読まれた時点で消えているため、ここに残っているのは
'  誰も取りに来ていないファイル)。列挙中にKillすると列挙が壊れるので、
' 必ず「集めてから消す」。
Private Sub GcOldThanks(ByVal folderPath As String)
    On Error Resume Next
    Dim days As Long: days = modConfig.GetLong("thanks_gc_days", 30)
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
            If stamp < limit Then KillRetry full
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
            sb = sb & vbLf & "　・" & who & " さん —「" & modUtil.SafeLeft(what, 40) & "」"
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
    EnsureDir folderPath

    Dim myId As String: myId = CurrentUserId()
    Dim srcHash As String: srcHash = modUtil.Fnv1a64Hex(source)

    Dim content As String
    content = myId & vbTab & SanitizeField(source) & vbTab & modUtil.NowStamp()

    ' MAX_PATH対策(Phase2.2): reporter IDもFnv1a64Hex(16桁)で綴る。(reporter,source)毎に
    ' 決定的なので再投票は同一ファイルを上書き=重複排除は不変。集計/GCはsrcHash前方一致で
    ' 行うためreporter部の綴りには非依存(payloadに生myIdを保持し報告者一覧は復元可能)。
    Dim votePath As String
    votePath = folderPath & "noise_" & srcHash & "_" & modUtil.Fnv1a64Hex(myId) & ".txt"
    ' 3回リトライしても書けなかった場合のE0705記録はWriteUtf8Retry内で一元化済み
    ' (重複排除により二重投票にはならないので、ここでは追加の対応は不要)。
    WriteUtf8Retry votePath, content, "modP2P.EmitNoiseVote"
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
    Dim shareRoot As String: shareRoot = modConfig.GetString("nexus_share_path", "")
    On Error Resume Next
    Dim probe As String: probe = Dir(shareRoot, vbDirectory)
    On Error GoTo Done
    If LenB(probe) = 0 Then Exit Function   ' 共有到達不能 → 除外状態を保持して撤退

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
        If ReadUtf8Retry(folderPath & flagNames(i), frec, "modP2P.CollectNoiseVotes") Then
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
        If ReadUtf8Retry(folderPath & voteNames(i), vrec, "modP2P.CollectNoiseVotes") Then
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
            WriteUtf8Retry folderPath & "gexcl_" & modUtil.Fnv1a64Hex(s) & ".txt", _
                s & vbTab & modUtil.NowStamp() & vbTab & SanitizeField(repList), "modP2P.CollectNoiseVotes"
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
        KillRetry folderPath & names(i)
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
    KillRetry folderPath & "gexcl_" & srcHash & ".txt"   ' 確定フラグを削除
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

' ソース名 → 作者(origin が "pack:<作者>" のときのみ。自作等は "")
Private Function AuthorOfSource(ByVal srcName As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastR
        If StrComp(CStr(ws.Cells(r, 2).Value), srcName, vbTextCompare) = 0 Then
            Dim origin As String: origin = CStr(ws.Cells(r, 3).Value)
            If LCase$(Left$(origin, 5)) = "pack:" Then AuthorOfSource = Trim$(Mid$(origin, 6))
            Exit Function
        End If
    Next r
End Function

Private Function ThanksDir() As String
    Dim basePath As String: basePath = modConfig.GetString("nexus_share_path", "")
    If LenB(basePath) = 0 Then Exit Function
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    ThanksDir = basePath & THANKS_SUBDIR & "\"
End Function

Private Function NoiseDir() As String
    Dim basePath As String: basePath = modConfig.GetString("nexus_share_path", "")
    If LenB(basePath) = 0 Then Exit Function
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    NoiseDir = basePath & NOISE_SUBDIR & "\"
End Function

Private Sub EnsureDir(ByVal folderPath As String)
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Function NewNonce(ByVal myId As String) As String
    NewNonce = myId & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & _
               Format$(Int(Timer * 1000) Mod 100000, "00000")
End Function

Private Function SeenNonce(ByVal nonce As String) As Boolean
    SeenNonce = (modStats.GetStat("thx:" & nonce) > 0)
End Function

Private Sub MarkNonce(ByVal nonce As String)
    modStats.Bump "thx:" & nonce
End Sub

Private Function SanitizeField(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    SanitizeField = t
End Function

' ネットワークドライブのロック(実行時エラー70/55/75等)に耐えるリトライI/O。
' 失敗時は短時間バックオフ待機して最大3回試行。全滅でも例外は出さず False。

' 3回リトライしても書けなかった場合、握りつぶさずE0705として記録する
' (SRE監査Phase2.1の方針をEmitNoiseVote以外の全書込み経路にも拡張。
' 実機環境の壁(共有フォルダのアクセス権限・セキュリティソフトのブロック等=
' エラー52/70想定)を、利用者が🩺診断の「コピー」ボタンでそのまま
' 開発者へ伝えられるようにするため、err_number構造化フィールドに残す)。
' context: 実際に呼んでいるPublic関数名を "modP2P.XxxYyy" 形式で呼び出し元が渡す
' (LogErrorのcontext引数は「modX.Y」形式だとYがPublicか契約チェックされるため
' 「§7」、Private助手関数自身の名前は使えない。呼び出し元ごとに正しい実体を
' 渡すことで、どの機能から起きた失敗かerr_log上で見分けられるようにする)。
Private Function WriteUtf8Retry(ByVal filePath As String, ByVal content As String, _
                                ByVal context As String) As Boolean
    Dim attempt As Long
    Dim lastNum As Long, lastDesc As String
    For attempt = 1 To 3
        If TryWriteUtf8(filePath, content, lastNum, lastDesc) Then
            WriteUtf8Retry = True
            Exit Function
        End If
        WaitMs 250 * attempt   ' 250 / 500 / 750ms バックオフ
    Next attempt
    On Error Resume Next
    modLog.LogError "E0705", context, _
        "3回リトライしても書込み失敗: " & modUtil.SafeLeft(filePath, 300) & " : " & lastDesc, lastNum
    On Error GoTo 0
End Function

Private Function TryWriteUtf8(ByVal filePath As String, ByVal content As String, _
                              Optional ByRef outErrNum As Long, Optional ByRef outErrDesc As String) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    st.Close
    Set st = Nothing            ' COM解放(正常パス)
    TryWriteUtf8 = True
    Exit Function
Fail:
    ' Err.Clear相当が起きる前に必ず先頭で退避する(以降のOn Error Resume Next/
    ' st.Closeで上書きされないようにするため)。
    outErrNum = Err.Number
    outErrDesc = Err.Description
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing            ' COM解放(異常パス=半開きも確実に解放)
    On Error GoTo 0
End Function

Private Function ReadUtf8Retry(ByVal filePath As String, ByRef outText As String, _
                               ByVal context As String) As Boolean
    Dim attempt As Long
    Dim lastNum As Long, lastDesc As String
    For attempt = 1 To 3
        If TryReadUtf8(filePath, outText, lastNum, lastDesc) Then
            ReadUtf8Retry = True
            Exit Function
        End If
        WaitMs 250 * attempt
    Next attempt
    ' エラー53(ファイルが見つかりません)は、列挙後に他ユーザーの並行GCで
    ' ファイルが消えた正常な競合(このモジュール冒頭のKillRetryコメント参照)。
    ' 想定内の自己解決ケースなのでログを汚さない。それ以外(52/70等の
    ' アクセス権限・ネットワーク瞬断)のみE0705として記録する。
    If lastNum <> 53 Then
        On Error Resume Next
        modLog.LogError "E0705", context, _
            "3回リトライしても読込み失敗: " & modUtil.SafeLeft(filePath, 300) & " : " & lastDesc, lastNum
        On Error GoTo 0
    End If
End Function

Private Function TryReadUtf8(ByVal filePath As String, ByRef outText As String, _
                             Optional ByRef outErrNum As Long, Optional ByRef outErrDesc As String) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    outText = st.ReadText
    st.Close
    Set st = Nothing            ' COM解放(正常パス)
    TryReadUtf8 = True
    Exit Function
Fail:
    outErrNum = Err.Number
    outErrDesc = Err.Description
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing            ' COM解放(異常パス=半開きも確実に解放)
    On Error GoTo 0
End Function

' 処理済み感謝状の削除(GC)。ロック時はバックオフして最大3回。失敗しても無害
' (nonce重複排除で二重加算は起きない)。err_logは汚さず、診断用にusage_logへ
' だけ記録する(GC失敗は非致命なので🩺診断の「直近のエラー」には出さない)。
Private Function KillRetry(ByVal filePath As String) As Boolean
    Dim attempt As Long
    Dim lastNum As Long
    For attempt = 1 To 3
        On Error Resume Next
        Err.Clear
        ' 並行GC耐性: 複数ユーザーが同時に同じ投票/感謝状をGCすると、後着のKillは
        ' error 53(ファイルなし)になる。「既に無い=削除目的は達成」なので成功扱いにし、
        ' 無駄な3回×最大1.5秒のリトライ(その間ロック保持)を避ける。
        If LenB(Dir(filePath)) = 0 Then
            On Error GoTo 0
            KillRetry = True
            Exit Function
        End If
        Kill filePath
        lastNum = Err.Number
        If lastNum = 0 Then
            On Error GoTo 0
            KillRetry = True
            Exit Function
        End If
        On Error GoTo 0
        WaitMs 250 * attempt
    Next attempt
    On Error Resume Next
    modLog.LogUsage "p2p_gc_failed", "", "3回リトライしても削除失敗 err#" & lastNum & ": " & modUtil.SafeLeft(filePath, 300)
    On Error GoTo 0
End Function

' Timer基準の短時間待機(DoEventsで応答性維持。Sleep API宣言を避けbitness非依存)。
Private Sub WaitMs(ByVal ms As Long)
    Dim t0 As Double: t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        DoEvents
        If Timer < t0 Then Exit Do   ' 日跨ぎ(深夜0時)ガード
    Loop
End Sub
