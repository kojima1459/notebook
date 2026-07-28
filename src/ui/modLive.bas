Attribute VB_Name = "modLive"
Option Explicit

' ============================================================================
' modLive - 待ち時間の実況(Nexusチャット画面)
' ----------------------------------------------------------------------------
' なぜこれが要るのか:
'   このアプリで最も価値のある1秒は、回答が出た瞬間ではなく「どの資料に
'   答えがあるか判明した瞬間」にある。検索は送信から1～2秒で終わっていて、
'   その時点で資料名もページ番号も手元にある。残りの10～20秒(しっかり
'   調べるなら1～2分)は、LLMがそれを日本語の文章に整えているだけの時間。
'
'   ところが従来、その一番おいしい情報は modUIMain.RenderSourcesPreview が
'   旧ホームシートのセルへ書いていた。既定UIがNexusチャットになった今、
'   利用者はそのシートを一度も見ない。結果として、送信してから回答が出る
'   までのあいだ、画面には「考えています…」が1個あるだけになっていた。
'   アプリは既に答えを見つけているのに、黙っていた。
'
'   ここが「すごいですね」と「えっ、もう見つけたの?」を分ける。
'   待ち時間は1秒も縮まらない。縮まらないまま、体感だけが別物になる。
'
' 仕組み:
'   modApp が送信直後に作る「考えています…」バブルの名前をここへ預ける
'   (Begin)。以降、modAsk からの進捗通知は modUIMain 経由でここへ流れ、
'   そのバブルの本文が書き換わる。modAsk 側は1行も変えていない
'   (§7.3のUIコールバック契約と R1 のレイヤ規約をそのまま守る)。
'
' 表示の設計:
'   資料が確定したら、その一覧は最後まで上に据え置く。下の1行(いま何を
'   しているか)だけが進む。据え置くのが肝心で、「見つけた」という事実が
'   途中で消えてしまうと、待ち時間はまた元のただの空白に戻ってしまう。
'
' 安全弁:
'   実況先が預けられていなければ全て即return(旧ホーム画面から質問された
'   場合や、起動直後などバブルが無い場合に何も起こさない)。
'   描画の失敗は握って捨てる。実況が出ないことはあっても、実況の失敗で
'   回答そのものが落ちることは絶対にない。
' ============================================================================

Private Const MAX_SOURCE_LINES As Long = 4

Private mBubble As String     ' 実況先のバブル名("" のときは何もしない)
Private mSources As String    ' 検索で確定した「見つかった資料」ブロック(据え置き)

' ----------------------------------------------------------------------------
' Begin - 実況先のバブルを預かる(modApp が送信直後に呼ぶ)
' ----------------------------------------------------------------------------
Public Sub Begin(ByVal bubbleName As String)
    mBubble = bubbleName
    mSources = ""
End Sub

' ----------------------------------------------------------------------------
' Finish - 実況先を手放す。回答が出た/失敗した の両方から必ず呼ぶこと
'   (持ち越すと、次のターンで存在しないバブルを書きに行くことになる)。
' ----------------------------------------------------------------------------
Public Sub Finish()
    mBubble = ""
    mSources = ""
End Sub

' ----------------------------------------------------------------------------
' PaintStage - 「いま何をしているか」の1行を差し替える。
'   modUIMain.SetStage から転送されてくる(検索中/回答作成中/検証中…)。
' ----------------------------------------------------------------------------
Public Sub PaintStage(ByVal stageMsg As String)
    If LenB(mBubble) = 0 Then Exit Sub
    If LenB(stageMsg) = 0 Then Exit Sub

    Dim shown As String
    shown = Humanize(stageMsg)

    Dim body As String
    If LenB(mSources) > 0 Then
        body = mSources & vbLf & vbLf & shown
    Else
        body = ChrW(&HD83D) & ChrW(&HDCAD) & " " & shown
    End If

    On Error Resume Next
    modUI.UpdateBubbleText mBubble, body
    DoEvents            ' ここで初めて実際に描画される(LLM呼び出しは同期で止まるため)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' PaintSources - 検索が終わった瞬間の実況。ここが体感を決める。
'   modUIMain.RenderSourcesPreview から転送されてくる。
' ----------------------------------------------------------------------------
Public Sub PaintSources(hits() As Hit, ByVal nHits As Long)
    If LenB(mBubble) = 0 Then Exit Sub

    mSources = BuildSourceBlock(hits, nHits)
    If LenB(mSources) = 0 Then Exit Sub

    PaintStage ChrW(&H270D) & ChrW(&HFE0F) & " いま、この内容を読んで回答をまとめています…"
End Sub

' ============================================================================
' 回答の「まとい」: 所要時間・読んだ量・資料が無いときの言い方
' ============================================================================

' ----------------------------------------------------------------------------
' Footer - 回答バブルの末尾に付ける1行。
' ----------------------------------------------------------------------------
' modAsk は所要秒も出典件数も計算していたが、書き出す先が旧ホームシートの
' セルだったため、チャットを見ている利用者には一度も届いていなかった。
' 分母の無い「15秒」はただ遅い。「3冊・8か所を12.4秒で」は速い。
' 同じ待ち時間が、言い方ひとつで自慢になる。
Public Function Footer(ByVal secs As Double, ByVal grounded As Boolean) As String
    Dim t As String
    t = ChrW(&H26A1) & " " & Format$(secs, "0.0") & "秒"

    If Not grounded Then
        Footer = t & " ・ 社内資料は未使用"
        Exit Function
    End If

    Dim spotN As Long
    Dim srcN As Long
    srcN = UniqueSourceCount(spotN)
    If srcN <= 0 Then
        Footer = t
    Else
        Footer = t & " ・ " & srcN & "冊 / " & spotN & "か所を読みました"
    End If
End Function

' StyleFooter - フッター段落だけを小さく淡くする。
'   段落の区切りは vbCr(vbLfでは1段落のままで Paragraphs が分かれない)。
Public Sub StyleFooter(ByVal bubbleName As String)
    If LenB(bubbleName) = 0 Then Exit Sub
    On Error Resume Next
    Dim shp As Shape
    Set shp = ThisWorkbook.Worksheets("Nexus").Shapes(bubbleName)
    If shp Is Nothing Then Exit Sub
    Dim pcount As Long
    pcount = shp.TextFrame2.TextRange.Paragraphs.count
    If pcount < 2 Then Exit Sub
    With shp.TextFrame2.TextRange.Paragraphs(pcount)
        .Font.Size = 8.5
        .Font.Fill.ForeColor.RGB = modUI.UiColor("muted")
    End With
    On Error GoTo 0
End Sub

' 直近回答が根拠にした資料の異なり数と、参照箇所の総数。
Private Function UniqueSourceCount(ByRef spotCount As Long) As Long
    Dim n As Long
    On Error Resume Next
    n = modAsk.LastHitCount()
    On Error GoTo 0
    If n <= 0 Then Exit Function

    Dim seen As String
    Dim uniq As Long
    Dim i As Long
    On Error Resume Next
    For i = 0 To n - 1
        Dim nm As String
        nm = Trim$(modAsk.LastHitSource(i))
        If LenB(nm) > 0 Then
            spotCount = spotCount + 1
            If InStr(1, "|" & seen & "|", "|" & nm & "|", vbTextCompare) = 0 Then
                seen = seen & IIf(LenB(seen) > 0, "|", "") & nm
                uniq = uniq + 1
            End If
        End If
    Next i
    On Error GoTo 0
    UniqueSourceCount = uniq
End Function

' ----------------------------------------------------------------------------
' 本棚が空のときの言い方
' ----------------------------------------------------------------------------
' 従来は「根拠になる資料がありませんでした」と断り、社内ポータルで資料を
' 探してくるよう12行の指示を出していた。配慮としては正しいのだが、初めて
' 開いた人が最初の質問で受け取る返事がそれでは、二度目は無い。断ったところで
' 利用者の疑問は消えず、結局どこかで勘に頼るだけになる。
'
' 「答えない」のではなく「何に基づく答えかをはっきりさせて答える」に変えた。
' 境界の見えている答えは、沈黙より安全で、かつ役に立つ。

' EmptyShelfGuard - 社内資料が無い状態で、社内固有の数字を断定させない指示。
'   これが無いと「うちの規程では30日です」のような、もっともらしくて出所の
'   無い一文が出る。損保の実務でそれは事故そのもの。
Public Function EmptyShelfGuard() As String
    EmptyShelfGuard = _
        "・この回答には社内資料が一切使えない。会社固有の条件・金額・期限・様式は断定せず、" & _
        "『社内資料での確認が必要』と明示すること。" & vbLf & _
        "・一般的な仕組みや考え方、確認すべき観点の整理に徹すること。"
End Function

' EmptyShelfWrap - 何に基づく答えかを前後で挟む。前置きを先に置くのは、
'   読み終えてから種明かしされると、読んだ内容の格が下がって見えるため。
Public Function EmptyShelfWrap(ByVal body As String) As String
    EmptyShelfWrap = _
        ChrW(&H26A0) & ChrW(&HFE0F) & " 社内資料がまだ1件も入っていないため、" & _
        "一般知識だけでお答えします(出典は付きません)。" & vbLf & vbLf & _
        body & vbLf & vbLf & _
        ChrW(&HD83D) & ChrW(&HDCC1) & " 入力欄の左のボタンから約款やマニュアルを入れると、" & _
        "次からは「どの資料の何ページか」まで付けて答えられます。"
End Function

' ============================================================================
' 血の通った余白: 時間帯の挨拶 / 弱音への返し
' ============================================================================
' modAppから移設(2026-07-27)。この画面が「どう感じられるか」を担う責務は
' すべてこのモジュールに寄せる。
Public Function TimeGreeting() As String
    Dim h As Long: h = Hour(Now)
    If h >= 5 And h < 10 Then
        TimeGreeting = "おはようございます。今日もスムーズにいきましょう。"
    ElseIf h >= 20 Or h < 5 Then
        TimeGreeting = "こんな時間までお疲れ様です。キリのいいところで切り上げてくださいね。"
    Else
        TimeGreeting = "こんにちは。"
    End If
End Function

' OpeningLine - 起動時の最初の1言。状態で変える。
'   資料が無いなら「入れてください」ではなく「そのまま聞ける」を先に言う。
'   何もできない状態から始めさせないための一文。
Public Function OpeningLine() As String
    Dim n As Long
    On Error Resume Next
    n = modShelf.TotalChunks()
    On Error GoTo 0

    ' 同梱の初期ナレッジがあるときは、必ず資料名を名指しする。
    ' 「何か聞いてください」では、利用者はまだ何を聞くか考えないといけない。
    ' 入っている物の名前が出ていれば、そこから連想するだけで済む。
    Dim docs As String
    On Error Resume Next
    If modSeed.HasSeed() Then docs = modSeed.SeedDocList()
    On Error GoTo 0

    If LenB(docs) > 0 Then
        Dim parts() As String: parts = Split(docs, "|")
        Dim body As String
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            body = body & vbLf & "　・" & parts(i)
        Next i
        Dim pending As Boolean
        On Error Resume Next
        pending = modSeed.SeedNeedsEmbedding()
        On Error GoTo 0

        If pending Then
            ' ベクトル未生成のまま配られた場合。「資料はあるのに答えられない」
            ' という状態を黙って見せるのが最悪なので、必ず先に言う。
            OpeningLine = TimeGreeting() & vbLf & _
                "この端末には、次の " & (UBound(parts) + 1) & "つの資料が入っています。" & body & vbLf & vbLf & _
                ChrW(&H23F3) & " いまAIが読み込み中です。終わるとこの資料から" & _
                "「どの資料の何ページか」まで付けて答えられるようになります。" & vbLf & _
                "読み込み中でも、一般的な内容ならそのままお答えします。"
        Else
            OpeningLine = TimeGreeting() & vbLf & _
                "この端末には、次の " & (UBound(parts) + 1) & "つの資料がすでに入っています。" & body & vbLf & vbLf & _
                "この内容なら、いま聞けば「どの資料の何ページか」まで付けてお答えします。" & vbLf & _
                "下の質問はどれも押すだけで試せます。まず1つどうぞ。"
        End If
    ElseIf n = 0 Then
        OpeningLine = TimeGreeting() & vbLf & _
            "知りたいことを、ふだんの言葉のまま書いてください。そのままお答えします。" & vbLf & _
            ChrW(&HD83D) & ChrW(&HDCC1) & " 左のボタンで約款やマニュアルを入れると、" & _
            "「どの資料の何ページか」まで付けて答えられるようになります。"
    Else
        OpeningLine = TimeGreeting() & vbLf & _
            "知りたいことを、ふだんの言葉のまま書いてください。" & _
            "本棚の資料から、出典付きでお答えします。"
    End If
End Function

Public Function IsTiredWords(ByVal q As String) As Boolean
    Dim t As String: t = Trim$(q)
    If Len(t) > 12 Then Exit Function   ' 長文は業務の質問(誤発動防止)
    IsTiredWords = (InStr(t, "疲れた") > 0 Or InStr(t, "つかれた") > 0 Or _
                    InStr(t, "しんどい") > 0 Or InStr(t, "眠い") > 0)
End Function

Public Function ComfortMessage() As String
    Dim pick As Long: pick = (Minute(Now) Mod 3)   ' 乱数を使わない決定的な出し分け
    Select Case pick
        Case 0
            ComfortMessage = "お疲れ様です！今日はずいぶん頑張ってはりますね。" & vbLf & _
                "温かいお茶でも飲んで、ちょっと一息つきましょか。" & ChrW(&HD83C) & ChrW(&HDF75)
        Case 1
            ComfortMessage = "ようやってはりますよ、ほんまに。" & vbLf & _
                "5分だけ肩の力抜いて、深呼吸してからまたいきましょ。" & ChrW(&H2615)
        Case Else
            ComfortMessage = "無理は禁物でっせ。仕事は明日も待ってくれます。" & vbLf & _
                "今日はここまでにして、はよ休んでくださいね。" & ChrW(&HD83C) & ChrW(&HDF19)
    End Select
End Function

' ----------------------------------------------------------------------------
' 内部: 実況文の言い換え
' ----------------------------------------------------------------------------
' modAsk が流してくる進捗文字列は、処理の名前がそのまま出ている
' (「回答作成中…」「検証中…」)。これは実装者の語彙であって、待っている人
' への説明にはなっていない。しかも資料が見つかった直後、modAsk はすぐ
' 「回答作成中…」を送ってくるので、せっかくの
' 「いま、この内容を読んで…」が一瞬で上書きされてしまう。
'
' 待ち時間に読む文章は、何が起きているかではなく「自分の質問がどう
' 扱われているか」を伝えるべきなので、UI層でここだけ書き換える。
' qa層は1行も変えない(どちらの語彙も、それぞれの層では正しい)。
Private Function Humanize(ByVal msg As String) As String
    Humanize = msg

    If InStr(msg, "検索中") > 0 Then
        Humanize = "本棚ぜんぶを見ています…"
    ElseIf InStr(msg, "質問を分析") > 0 Then
        Humanize = "ご質問の意図を読み取っています…"
    ElseIf InStr(msg, "関連度を精査") > 0 Then
        Humanize = "見つかった中から、いちばん確かなものを選んでいます…"
    ElseIf InStr(msg, "下書き") > 0 Then
        Humanize = "見つけた内容を読んで、下書きを書いています…"
    ElseIf InStr(msg, "検証") > 0 Then
        ' ここが「しっかり調べる」の価値そのもの。何をしているか伝われば、
        ' 1～2分は「遅い」ではなく「そこまでやるのか」に変わる。
        Humanize = "書いた内容を、資料と1行ずつ突き合わせて確認しています…"
    ElseIf InStr(msg, "回答作成") > 0 Then
        Humanize = "見つけた内容を読んで、回答をまとめています…"
    ElseIf InStr(msg, "件の資料がヒット") > 0 Then
        ' 直後に PaintSources が資料名まで出すので、ここは何も足さない。
        Humanize = "答えのありかを絞り込んでいます…"
    End If
End Function

' ----------------------------------------------------------------------------
' 内部: 見つかった資料のブロックを組み立てる。
'   同じ資料の別チャンクがそのまま並ぶと「3件見つかった」の重みが消えるので、
'   資料名で束ねて、最も関連が高かった箇所のページだけを出す
'   (hits は関連度の高い順に並んでいるので、最初に出会ったものが最良)。
' ----------------------------------------------------------------------------
Private Function BuildSourceBlock(hits() As Hit, ByVal nHits As Long) As String
    If nHits <= 0 Then Exit Function

    Dim lines_() As String
    ReDim lines_(0 To MAX_SOURCE_LINES - 1)
    Dim seen As String
    Dim shown As Long
    Dim uniq As Long

    Dim i As Long
    On Error Resume Next
    For i = 1 To nHits
        Dim nm As String
        nm = Trim$(hits(i).source)
        If LenB(nm) > 0 Then
            If InStr(1, "|" & seen & "|", "|" & nm & "|", vbTextCompare) = 0 Then
                seen = seen & IIf(LenB(seen) > 0, "|", "") & nm
                uniq = uniq + 1
                If shown < MAX_SOURCE_LINES Then
                    Dim ln As String
                    ln = "　・" & modUtil.SafeLeft(nm, 34)
                    If hits(i).page > 0 Then ln = ln & "  p." & hits(i).page
                    lines_(shown) = ln
                    shown = shown + 1
                End If
            End If
        End If
    Next i
    On Error GoTo 0

    If shown = 0 Then Exit Function

    Dim body As String
    body = ChrW(&HD83D) & ChrW(&HDCC4) & " " & uniq & "件の資料に、答えがありそうです"
    For i = 0 To shown - 1
        body = body & vbLf & lines_(i)
    Next i
    If uniq > shown Then body = body & vbLf & "　　ほか " & (uniq - shown) & "件"

    BuildSourceBlock = body
End Function
