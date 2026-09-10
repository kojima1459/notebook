Attribute VB_Name = "modOutlineBuild"
Option Explicit

' ============================================================================
' modOutlineBuild - 章単位要約の作成(R17 Phase2・設計書§3 Phase2)
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   検索は「点」しか拾えない。dense+sparse の上位k件は、質問の語に近い
'   チャンクを並べたものであって、「この資料には何が書いてあるか」という
'   問いには原理的に答えられない(俯瞰質問=「〜を全部教えて」「全体像は?」)。
'   規程・マニュアルは既に人の手で【章】という単位に割れているので、
'   GraphRAG のコミュニティ要約の代わりに、その章をそのまま要約単位にする。
'   取込時に章ごと1回だけ CallLLM して doc_outline へ貯め、質問時は
'   章の要約だけを読んで「どの章を読むか」を選ぶ(modAskGlobal)。
'
' 設計判断:
'   ・章キー = section_path の第1要素(ChapterKeyOf)。章見出しの無い資料では
'     これが「第12条(…)」のような条単位になるが、それが正しい保守的動作。
'     見出しの無い資料に勝手な章立てを与えるより、条ごとの要約が並ぶほうが
'     選択段の材料として正直で、外れても「条を1つ余計に読む」だけで済む。
'   ・LLM呼び出しは章数ぶん(254頁の規程で20〜30回=取込+3〜8分)。R15の枠組み
'     (中断ボタン・進捗バナー・ETA・中間保存)にそのまま乗せる。中断されたら
'     【そこまでの章を保存して正常終了】する: 20章のうち12章まで終えた成果を
'     捨てて0章に戻すのは、待った時間の全損になる(R15の一貫した判断)。
'   ・1章の失敗は「(要約失敗)」の行として保存して次の章へ進む。行ごと落とすと、
'     利用者にも次に触る人にも「その章だけ要約が無い」ことが見えない(憲章§3-3)。
'   ・1章の本文は max_context_chars で打ち切る(BudgetTake)。章はいくらでも
'     長くなり得るのに、1回のLLM呼び出しに載る量は決まっている。
'   ・失敗もゲート(config graph_outline)も【この層に閉じる】。呼び出し元
'     (modShelf.IngestFile)は残227字しかなく、判断を1つも持てない(憲章§4-6)。
'   ・プロンプトは modPrompts ではなくここの Private に置く(modPrompts は残386字。
'     章要約プロンプトは1本も入らない=R17 Phase2 の容量裁定)。
'
' 2026-08-07(R21-3 E2・実機第8報②): 章検出の根治(⚡仕上げの射程内=既存資料も
' 再取込なしで改善)。
'   ・ChapterKeyOf: 目次行キー(「第3章 総則……12」)と本文見出しキー
'     (「第3章 総則」)を末尾のリーダー記号+頁番号の除去で統合する。
'   ・GroupChapters: MAX_CHAPTERS超過時の「先着60で無言切り捨て」を廃止し、
'     1〜2チャンクの章を隣接章へ統合する2パス縮退へ置き換えた
'     (CoalesceSmallChapters)。OUTLINE_LOGIC_VERはこの改善の世代キーで、
'     doc_outlineへ記録し旧世代をmodBackfill.DetectLegacyDocsが「未仕上げ」
'     扱いにする(ロジック改善後に⚡仕上げが再提案される)。
' ============================================================================

' my_knowledge の列(modShelfStore の定義と同じ並び。ここは読むだけ)。
Private Const COL_ID As Long = 1
Private Const COL_SOURCE As Long = 2
Private Const COL_FULLTEXT As Long = 7

' 1資料あたりの章数の天井(=LLM呼び出し回数の天井)。254頁の規程で20〜30章。
' 見出しが取れない資料で条単位に割れたときの暴走(数百回の呼び出し)を止める。
Private Const MAX_CHAPTERS As Long = 60

' 章の最小サイズ(R21-3 E1後方互換ガード): これ未満のチャンク数の章は隣接
' (文書順で次)の章へ統合する。Visionが階層指示に完全に従わない場合の
' フォールバック(1〜2チャンクの「章」を作らせない)。
Private Const MIN_CHAPTER_CHUNKS As Long = 3

' doc_outlineの章検出ロジック世代(R21-3 E2)。GroupChapters/ChapterKeyOfの
' 改善(縮退統合・末尾正規化・目次抑制と対)を適用して書いた行にはこの値を
' 記録する。modBackfill.DetectLegacyDocsはこれより小さい(=未記録=0を含む)
' 資料を「未仕上げ」として再提案する単一情報源(§4-5)。
Public Const OUTLINE_LOGIC_VER As Long = 2

' 縮退統合(CoalesceSmallChapters)が複数の生キーを1つの代表キーへ束ねる際の
' 連結記号(R21-3 E2)。"|" にしないのは、この代表キーが doc_outline経由で
' modAskGlobal.BuildPickPromptの一覧に出て【LLMの<pick>応答へそのまま
' 書き写される】文字列だから: modRagParse.ParseSubqueries は複数章の選択を
' "|" で分割する(章キー自身に"|"が入ると、LLMが正しく丸ごと書き写しても
' 応答パース側で章キーの後半が「別の(不正な)pick」として切り離されてしまう。
' 憲章§4-5: 単一情報源の式のつもりが、別の文脈の区切り文字と衝突していた
' 実装ミスの是正)。全角の｜(U+FF5C)は半角"|"と字形が近いが別コードポイントで、
' 章見出しにまず出現せず"::"/半角"|"いずれとも衝突しない。Const初期化に
' ChrW()等の関数呼び出しは使えない(VBA仕様。実行はできてもコンパイルが
' 通らない=LO実行テストがexit=-9でハングして発覚した)ため、リテラル文字を
' 直接埋め込む(tools/vba_lint.py check_cp932_safeでCP932往復可能=文字化け
' しないことを固定済み)。
Private Const CHAPTER_KEY_SEP As String = "｜"

' 1章の本文へ載せる最低限(max_context_chars が異常に小さい設定でも、
' 要約の材料が数十字しか渡らない状態にはしない)。
Private Const CAP_MIN As Long = 2000

Private Const SUMMARY_MAX As Long = 600
Private Const KEYWORDS_MAX As Long = 200
Private Const FAIL_TEXT As String = "(要約失敗)"

' 無人実行(自動同期)中は章要約を作らない資料の章数(2026-08-05 R17H FA-7)。
Private Const UNATTENDED_MAX_CH As Long = 12

' いま走っているのが「誰も見ていない処理」か(2026-08-05 R17H FA-7 / A-M8)。
' 自動同期(modShelfSync.SyncNow silent:=True)は OnTime で勝手に始まり、
' 進捗バナーも中断ボタンも利用者の目の前には無い。そこで章数の多い資料の
' 章要約(=章数ぶんのLLM呼び出し。254頁で20〜30回・数分〜十数分)を始めると、
' 「何もしていないのにExcelが数分固まる」だけの体験になる。手動取込・手動
' 同期(silent=False)では従来どおり全量作る。印の上げ下げは modShelfSync の
' 入口/Finish の2行だけで、こちらからは読むだけ。
Private mUnattended As Boolean

' ----------------------------------------------------------------------------
' SetUnattended - 無人実行(自動同期)の印を上げ下げする(R17H FA-7)。
'   呼ぶのは modShelfSync.SyncNow の入口(silent の値)と Finish(False)だけ。
' ----------------------------------------------------------------------------
Public Sub SetUnattended(ByVal b As Boolean)
    mUnattended = b
End Sub

' ----------------------------------------------------------------------------
' BuildOutlineFor - 1資料ぶんの章要約を作って doc_outline へ保存する。
'   config graph_outline=off / chunk_meta が0行 / その資料のチャンクが無い /
'   章キーが1つも取れない のいずれかなら【完全に無操作】(=LLMも呼ばない)。
'   呼び出しは modShelf.IngestFile の Finish(done/partial)から1行、および
'   2026-08-06(R20-3): modBackfill.BackfillOne(再取込ゼロの「資料の仕上げ」。
'   Phase1の構造メタを保存済みfull_textから再計算した直後)から1行の計2箇所。
'   どちらも「その資料のchunk_metaが直前に揃った」直後に呼ぶだけで、この
'   関数自体の判断(ゲート・中断・失敗握り)は1文字も変えない。
' ----------------------------------------------------------------------------
Public Sub BuildOutlineFor(ByVal sourceName As String)
    If Not GateOn() Then Exit Sub
    If LenB(Trim$(sourceName)) = 0 Then Exit Sub

    On Error GoTo Quiet

    ' 1) 構造メタが無ければ章に割れない(=まだ取り込み直していない本棚)。
    Dim mIds() As String, mPaths() As String, mRefs() As String
    Dim metaN As Long
    metaN = modChunkMetaStore.ReadAllMeta(mIds, mPaths, mRefs)
    If metaN < 1 Then Exit Sub

    ' 2) その資料のチャンクを文書順(=シートの並び順)に集める。
    Dim rowIdx() As Long, ids() As String
    Dim nSrc As Long
    nSrc = CollectSourceRows(sourceName, rowIdx, ids)
    If nSrc < 1 Then Exit Sub

    ' 3) chunk_id → section_path を引き当てる(chunk_meta 側を1周するだけ)。
    Dim paths() As String: ReDim paths(1 To nSrc)
    MapPaths ids, nSrc, mIds, mPaths, metaN, paths

    ' 4) 章キーで束ねる(出現順=文書順)。上限超過時は無言切り捨てず縮退統合
    ' する(R21-3 E2)ので、その旨をusage_logへ残す(現状は無言だった)。
    Dim chKeys() As String, chCount() As Long
    Dim nCh As Long, uniqRaw As Long, mergedN As Long, droppedN As Long
    nCh = GroupChapters(paths, nSrc, chKeys, chCount, uniqRaw, mergedN, droppedN)
    If nCh < 1 Then Exit Sub
    If uniqRaw > MAX_CHAPTERS Then
        On Error Resume Next
        modLog.LogUsage "outline_capped", "ingest", _
            "source=" & sourceName & " uniq=" & uniqRaw & " merged=" & mergedN & _
            " dropped=" & droppedN
        On Error GoTo 0
    End If

    ' 無人実行(自動同期)で章数が多い資料は、章要約を先送りする(R17H FA-7)。
    ' 手動取込・手動同期(非silent)なら同じ資料でも全量作られる=先送りは
    ' 「作らない」ではなく「人が見ているときに作る」。名寄せ(name_dedup)は
    ' 1資料につき最大1回なので、無人でもそのまま続ける。
    If mUnattended And nCh > UNATTENDED_MAX_CH Then
        On Error Resume Next
        modLog.LogUsage "outline_deferred", "ingest", _
            "source=" & sourceName & " chapters=" & nCh & " (unattended)"
        On Error GoTo 0
        GoTo SynStep
    End If

    Dim cap As Long: cap = SafeCap()
    Dim sums() As String: ReDim sums(1 To nCh)
    Dim kws() As String: ReDim kws(1 To nCh)
    Dim srcs() As String: ReDim srcs(1 To nCh)

    Dim eff As String: eff = modConfig.GetString("quick_effort", "low")
    Dim vrb As String: vrb = modConfig.GetString("quick_verbosity", "low")
    Dim mdl As String: mdl = modConfig.GetString("quick_model", "gpt-5.5")

    Dim msPerItem As Double
    Dim doneN As Long
    Dim i As Long

    On Error GoTo ChapFail
    For i = 1 To nCh
        ' 中断は章の境界で拾う(LLM呼び出しの内側では modGateway が握り潰す)。
        If modShelfBatch.CancelRequested() Then Exit For
        Banner i, nCh, sourceName, msPerItem

        Dim t0 As Double: t0 = Timer
        srcs(i) = sourceName
        sums(i) = FAIL_TEXT      ' 先に失敗を書いておく=どこで落ちても正直に残る
        kws(i) = ""

        Dim body As String
        body = ChapterBody(rowIdx, paths, nSrc, chKeys(i), cap)
        If LenB(body) > 0 Then
            Dim lat As Long
            Dim resp As String
            resp = modGateway.CallLLM(BuildChapterPrompt(sourceName, chKeys(i), body), _
                                      "chapter_summary", eff, vrb, mdl, lat)
            Dim sm As String, kw As String
            If modRagParse.ParseOutlineResp(resp, sm, kw) Then
                sums(i) = modUtil.SafeLeft(sm, SUMMARY_MAX)
                kws(i) = modUtil.SafeLeft(kw, KEYWORDS_MAX)
            End If
        End If

        doneN = i
        msPerItem = modUtilText.BlendPerItemMs(msPerItem, t0, 1)
        modShelfBatch.SaveCheckpoint 1, 120
    Next i
    GoTo LoopDone

ChapFail:
    If Err.Number = 18 Then
        ' ESC。ここまでの章を保存して正常終了する(待った時間を捨てない)。
        Err.Clear
        Resume LoopDone
    End If
    ' 想定外の失敗はその章だけ「(要約失敗)」のまま次へ(1章の事故で全部を失わない)。
    Err.Clear
    Resume Next

LoopDone:
    On Error GoTo Quiet
    If doneN < 1 Then Exit Sub

    ' 2026-08-16(R33波3 W3-4): 世代キーを刻むのは【完走したときだけ】。
    ' このループは中断(Exit For)・ESC(Resume LoopDone)・章単位の想定外エラー
    ' (Resume Next で sums(i)=FAIL_TEXT のまま)でも doneN>=1 なら書き込む。
    ' そこへ現行世代を刻むと、読み手(modBackfill.BuildOutlineSourceSet)は
    ' 「1章でも現行世代で書けていれば仕上げ済み」と見なすので、60章のうち
    ' 数章しか無い資料が【永久に】⚡資料の仕上げの候補から外れる。
    ' 俯瞰質問はその数章ぶんしか見ないまま固定され、欠落は誰にも見えない。
    Dim failN As Long
    For i = 1 To doneN
        If sums(i) = FAIL_TEXT Then failN = failN + 1
    Next i

    ' 作り直し: 同じ資料の古い章要約を落としてから今回ぶんを1回で書く
    ' (取込フック側でも掃除しているが、ここが唯一の書込み口なので二重に守る。
    '  同じ資料の行が2組並ぶと、章選択の段で同じ章が2回候補に出る)。
    modOutlineStore.RemoveOutlineForSource sourceName
    modOutlineStore.WriteOutlineRows srcs, chKeys, sums, kws, chCount, doneN, _
        OutlineVerFor(doneN, nCh, failN)

    On Error Resume Next
    modUIMain.SetStage ""
    modLog.LogUsage "outline_built", "ingest", _
        "source=" & sourceName & " chapters=" & doneN & "/" & nCh
    On Error GoTo 0

SynStep:
    ' R17 Phase3: 章要約に続けて名寄せ辞書のバッチ生成を1回だけ(ゲート・失敗握り
    ' ・usage_log("synonyms_fail")はmodSynonymStore側に閉じる。ここは1行)。
    On Error GoTo Quiet
    If Not modShelfBatch.CancelRequested() Then
        On Error Resume Next
        modSynonymStore.BuildSynonymsFor sourceName
        ' R37 §3: 資料間リンク(章の重心。LLMは1回も呼ばない)。ゲート
        ' (config xdoc_links)・上限・失敗握りは向こうに閉じるのでここは1行。
        ' 埋め込み(modShelf.IngestFile 手順9 EmbedPending)はこの時点で
        ' 済んでいるので、重心の材料になるベクトルは揃っている。
        modXDocStore.BuildLinksFor sourceName
        On Error GoTo 0
    End If
    Exit Sub

Quiet:
    ' ハンドラ稼働中は On Error Resume Next が効かないので、必ず Resume で
    ' 抜けてから記録する(2026-07-30 実機err#462 と同型の作法)。
    Resume QuietDone
QuietDone:
    Err.Clear
    On Error Resume Next
    modLog.LogUsage "outline_fail", "ingest", "build source=" & sourceName
    On Error GoTo 0
End Sub

' ============================================================================
' 純ロジック(LibreOfficeの実行テストで固定する)
' ============================================================================

' ----------------------------------------------------------------------------
' ChapterKeyOf - section_path("第3章 総則>第12条(免責)")から章キーを取る。
'   戻り値 = 第1要素("第3章 総則")に、末尾の「リーダー記号列+頁番号」だけを
'   素手走査で除去した緩やか正規化を掛けたもの(R21-3 E2)。section_path は
'   取込時に modSparse.NormalizeForSearch を通っているので、ここへ足すのは
'   【この関数1本】に閉じた追加正規化に留める限り安全: ChapterKeyOfは
'   GroupChapters/ChapterBody/CollectChapterHits(modAskGlobal)の全員が
'   必ずこの関数を通してから比較するため、保存側と照合側の式は常に一致する
'   (憲章§4-5。「別の式」を各呼び出し側へ書かないことが単一情報源の条件)。
'   ・目次行由来のキー("第3章 総則……12")と本文見出し由来のキー
'     ("第3章 総則")が、この正規化で同一キーへ統合される。
'   ・リーダー記号が1字だけ(文中の句点等)/頁番号が無い末尾には触れない
'     (誤って章名を削らないための保守的なガード)。
'   ・章見出しが無い資料では第1要素が条になる(=条単位の要約)。それが正しい
'     保守的動作で、無い章立てを推測するよりも外れ方が小さい。
'   ・空/">"だけ/先頭が空の path は空文字("章が分からない"という正しい答え。
'     呼び出し元はその行を束ねの対象から外す)。
' ----------------------------------------------------------------------------
' OutlineVerFor - doc_outline へ刻む世代キー(純ロジック)。
'   2026-08-16(R33波3 W3-4): 世代キーは「どのロジックで作ったか」しか表せない
'   のに、読み手(modBackfill.BuildOutlineSourceSet)は「仕上がっているか」の
'   判定に使う。したがって【完走していない章要約に現行世代を刻んではならない】。
'   0 を返すと、読み手の後方互換読み(Val("")=0)と同じ「未仕上げ」扱いになり、
'   ⚡資料の仕上げの候補として正しく再提案される(移行処理は不要)。
'   未完走とみなす条件は2つ:
'     ・doneN < nCh …… 中断・ESCで章の途中で抜けた
'     ・failN > 0   …… 章単位の想定外エラーで "(要約失敗)" のまま残った章がある
' ----------------------------------------------------------------------------
Public Function OutlineVerFor(ByVal doneN As Long, ByVal nCh As Long, _
                              ByVal failN As Long) As Long
    If doneN < 1 Then Exit Function
    If doneN < nCh Then Exit Function
    If failN > 0 Then Exit Function
    OutlineVerFor = OUTLINE_LOGIC_VER
End Function

' ----------------------------------------------------------------------------
Public Function ChapterKeyOf(ByVal sectionPath As String) As String
    Dim t As String: t = Trim$(sectionPath)
    If LenB(t) = 0 Then Exit Function
    Dim p As Long: p = InStr(t, ">")
    If p > 0 Then t = Left$(t, p - 1)
    ChapterKeyOf = StripTocTail(Trim$(t))
End Function

' ChapterKeyOfRaw - ChapterKeyOfからStripTocTail(末尾リーダー+頁番号除去)
'   だけを外したもの(R21H F5)。OUTLINE_LOGIC_VER<2で保存された旧世代の
'   doc_outlineの章キーはこの正規化が導入される前に作られているため、現行の
'   ChapterKeyOfで再計算したキーとは一致しない(俯瞰が「no_chunk」で無音全滅
'   する)。modAskGlobal.CollectChapterHitsが正規化前後の2段でこの関数と
'   ChapterKeyOfの両方を試すためのフォールバック(⚡仕上げ実行までの橋渡し)。
Public Function ChapterKeyOfRaw(ByVal sectionPath As String) As String
    Dim t As String: t = Trim$(sectionPath)
    If LenB(t) = 0 Then Exit Function
    Dim p As Long: p = InStr(t, ">")
    If p > 0 Then t = Left$(t, p - 1)
    ChapterKeyOfRaw = Trim$(t)
End Function

' 末尾の「リーダー記号2字以上+頁番号」を除去する(目次行キーの統合用)。
' 除去条件を満たさなければ入力をそのまま返す(安全側)。RegExp不使用。
Private Function StripTocTail(ByVal s As String) As String
    StripTocTail = s
    Dim n As Long: n = Len(s)
    Dim i As Long: i = n
    ' VBAのAndは短絡評価しない(i=0でもMid$(s,i,1)が評価されてErr5になる)ため、
    ' 境界チェックと文字判定を同じDo While条件に混ぜない(ループ本体側で判定)。
    Do While i >= 1
        If Not IsAsciiDigit(Mid$(s, i, 1)) Then Exit Do
        i = i - 1
    Loop
    If i = n Then Exit Function              ' 末尾が数字でない=頁番号なし
    Dim leaderN As Long, j As Long: j = i
    Do While j >= 1
        Dim c As String: c = Mid$(s, j, 1)
        If c = "." Or c = ChrW(&HFF0E) Or c = ChrW(&H30FB) _
                Or c = ChrW(&H2026) Or c = ChrW(&H2025) Then
            leaderN = leaderN + 1: j = j - 1
        ElseIf c = " " Then
            j = j - 1
        Else
            Exit Do
        End If
    Loop
    If leaderN < 2 Then Exit Function          ' リーダー1字だけ=誤爆防止
    Dim head As String: head = Trim$(Left$(s, j))
    If LenB(head) = 0 Then Exit Function        ' 全部食うと章名が消える
    StripTocTail = head
End Function

Private Function IsAsciiDigit(ByVal ch As String) As Boolean
    IsAsciiDigit = (ch >= "0" And ch <= "9")
End Function

' ----------------------------------------------------------------------------
' BudgetTake - 「あと何字入れられるか」の算数(1章の本文の打ち切り)。
'   usedLen: すでに積んだ字数 / addLen: 次に足したい字数 / cap: 上限
'   戻り値 0 = もう1字も入らない(呼び出し元はそこで打ち切る)。
'   cap を超えた分を黙って捨てるのではなく【手前で切る】ことにしているのは、
'   LLM側で切られると章の後半が理由も分からず消えるため(こちらで切れば
'   「どこまで読んだか」が呼び出し元の手に残る)。
' ----------------------------------------------------------------------------
Public Function BudgetTake(ByVal usedLen As Long, ByVal addLen As Long, _
                           ByVal cap As Long) As Long
    If cap <= 0 Or addLen <= 0 Then Exit Function
    Dim room As Long: room = cap - usedLen
    If room <= 0 Then Exit Function
    If addLen < room Then
        BudgetTake = addLen
    Else
        BudgetTake = room
    End If
End Function

' ============================================================================
' 内部ヘルパー
' ============================================================================

' config graph_outline(既定on)。off のときだけ切る(graph_refs と同じ作法)。
Private Function GateOn() As Boolean
    GateOn = (LCase$(Trim$(modConfig.GetString("graph_outline", "on"))) <> "off")
End Function

' 1章の本文へ載せる上限。LLM呼び出し1回ぶんの上限と同じ値を使う。
Private Function SafeCap() As Long
    Dim v As Long: v = 40000
    On Error Resume Next
    v = modConfig.GetLong("max_context_chars", 40000)
    On Error GoTo 0
    If v < CAP_MIN Then v = CAP_MIN
    SafeCap = v
End Function

' その資料の行(シート順=文書順)の行番号と chunk_id を集める。
' 読むのは chunk_id 列と source 列の2列だけ(本文まで全部読むと20,000件規模で
' 数百MBになる。modAskFocus と同じ作法)。
Private Function CollectSourceRows(ByVal sourceName As String, ByRef outRows() As Long, _
                                   ByRef outIds() As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function

    Dim lastK As Long
    lastK = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).row
    If lastK < 2 Then Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, COL_ID), ws.Cells(lastK, COL_SOURCE)).Value

    ReDim outRows(1 To lastK - 1)
    ReDim outIds(1 To lastK - 1)
    Dim n As Long
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(Trim$(CStr(arr(i, COL_SOURCE))), sourceName, vbTextCompare) = 0 Then
            Dim id As String: id = Trim$(CStr(arr(i, COL_ID)))
            If LenB(id) > 0 Then
                n = n + 1
                outRows(n) = i + 1          ' arr は2行目起点
                outIds(n) = id
            End If
        End If
    Next i
    CollectSourceRows = n
End Function

' chunk_id → section_path の引き当て。chunk_meta 側を1周し、その資料の
' chunk_id の箱に入っているものだけ位置を割り出して書き込む(箱は数十KB、
' meta は本棚全体ぶん。modAskFocus.RefsExpand と同じ「箱へのInStr」流儀)。
Private Sub MapPaths(ByRef ids() As String, ByVal nSrc As Long, _
                     ByRef mIds() As String, ByRef mPaths() As String, _
                     ByVal metaN As Long, ByRef outPaths() As String)
    Dim box As String
    Dim i As Long
    For i = 1 To nSrc
        box = box & vbLf & ids(i)
    Next i
    box = box & vbLf

    For i = 0 To metaN - 1
        If LenB(mPaths(i)) > 0 Then
            Dim k As Long: k = IndexInBox(box, mIds(i))
            If k >= 1 And k <= nSrc Then outPaths(k) = mPaths(i)
        End If
    Next i
End Sub

' vbLf区切りの箱の中で id が何番目か(1始まり。無ければ0)。
' 位置は「一致した所より前にある vbLf の数」で決まる(箱は必ず vbLf で始まる)。
Private Function IndexInBox(ByVal box As String, ByVal id As String) As Long
    If LenB(id) = 0 Then Exit Function
    Dim p As Long
    p = InStr(1, box, vbLf & id & vbLf, vbBinaryCompare)
    If p = 0 Then Exit Function
    Dim head As String: head = Left$(box, p)
    IndexInBox = Len(head) - Len(Replace(head, vbLf, ""))
End Function

' ----------------------------------------------------------------------------
' GroupChapters - 章キーで束ねる(文書順=出現順)。戻り値=最終的な章数。
'   outKeys/outCount は最終(縮退統合後)の章。outUniqRaw/outMerged/outDropped
'   は診断用(R21-3 E2・呼び出し元のusage_log用): uniqRaw=縮退前のユニーク
'   章キー数、merged=uniqRaw-最終章数(統合で消えた章数)、dropped=0固定
'   (統合はするが取りこぼしはしない=全チャンクが必ずどこかの章に属す。
'    「先着60で無言切り捨て」だった旧実装からの脱却)。
'   縮退の2段構え:
'     (1) 常に適用する最小サイズガード(MIN_CHAPTER_CHUNKS未満の章を隣接の
'         次の章へ統合。Visionが階層指示に従わない場合のフォールバック)。
'     (2) それでもMAX_CHAPTERSを超えるときだけ、章数から逆算した粒度で
'         もう一段統合する(2パス縮退。CoalesceSmallChaptersを閾値を上げて
'         再適用。上げても収まらなければ閾値を段階的に引き上げて必ず収束
'         させる=安全弁)。
'   統合で複数キーが1章にまとまった場合、代表キーは全メンバーを
'   CHAPTER_KEY_SEP区切りで連結した文字列になる(ChapterKeyMatchesが照合時に
'   分解して見る。単一キーのままの章は従来どおりの素の文字列で、後方互換を
'   壊さない)。"|"を使わない理由はCHAPTER_KEY_SEPの定義コメント参照。
' ----------------------------------------------------------------------------
Public Function GroupChapters(ByRef paths() As String, ByVal nSrc As Long, _
                              ByRef outKeys() As String, ByRef outCount() As Long, _
                              ByRef outUniqRaw As Long, ByRef outMerged As Long, _
                              ByRef outDropped As Long) As Long
    ReDim outKeys(1 To nSrc)
    ReDim outCount(1 To nSrc)
    outUniqRaw = 0: outMerged = 0: outDropped = 0

    ' R21H F9(敵対的レビュー確定): ユニーク化がoutKeysを毎回1件ずつ線形走査
    ' しており、章見出しの無い資料(条ごとに割れて実質nSrc件に近い)ではO(n²)
    ' になる。既存使用例(modBoard/modVaultGallery/modCluster等)に倣い
    ' Scripting.Dictionaryでkey→index引きをO(1)化する。ただしDictionaryは
    ' Windows Script Runtime(COM)でLibreOfficeには無く、GroupChaptersは
    ' modTestsPure23からLO純ロジックテストとして直接実行される(既存の
    ' Excel専用モジュールでのDictionary利用と違い、コンパイルだけでなく
    ' 実行までされる)。生成に失敗したら黙って旧来のO(n²)線形走査へ落とす
    ' (結果は完全に同一。速度だけの最適化)。
    Dim n As Long
    Dim totalGrouped As Long
    Dim i As Long, j As Long
    Dim dict As Object, useDict As Boolean
    On Error Resume Next
    Set dict = CreateObject("Scripting.Dictionary")
    If Not dict Is Nothing Then dict.CompareMode = 1   ' vbTextCompare相当
    useDict = (Not dict Is Nothing) And (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    For i = 1 To nSrc
        Dim k As String: k = ChapterKeyOf(paths(i))
        If LenB(k) > 0 Then
            totalGrouped = totalGrouped + 1
            Dim at As Long: at = 0
            If useDict Then
                On Error Resume Next
                If dict.Exists(k) Then at = CLng(dict(k))
                On Error GoTo 0
            Else
                ' ChapterKeyMatches側をvbTextCompareへ揃えたのに合わせ、ここ
                ' (章のユニーク化)もvbTextCompareへ統一(大小文字/全角半角の
                ' ゆれで同じ章が2つに割れるのを防ぐ。統一しないとGroupChapters
                ' が別々に数えた章を、後段のChapterKeyMatchesは同一章として
                ' 拾ってしまい、章数と実際の紐付けが食い違う)。
                For j = 1 To n
                    If StrComp(outKeys(j), k, vbTextCompare) = 0 Then
                        at = j
                        Exit For
                    End If
                Next j
            End If
            If at > 0 Then
                outCount(at) = outCount(at) + 1
            Else
                n = n + 1
                outKeys(n) = k
                outCount(n) = 1
                If useDict Then
                    On Error Resume Next
                    dict(k) = n
                    On Error GoTo 0
                End If
            End If
        End If
    Next i
    outUniqRaw = n

    ' (1) 最小サイズガード(常時)
    CoalesceSmallChapters outKeys, outCount, n, MIN_CHAPTER_CHUNKS

    ' (2) それでも上限超なら、粒度を上げて再集約(無言切り捨ての廃止)
    If n > MAX_CHAPTERS Then
        Dim minSize2 As Long: minSize2 = CeilDivLong(totalGrouped, MAX_CHAPTERS)
        If minSize2 <= MIN_CHAPTER_CHUNKS Then minSize2 = MIN_CHAPTER_CHUNKS + 1
        Dim guard As Long
        Do While n > MAX_CHAPTERS And guard < 40
            CoalesceSmallChapters outKeys, outCount, n, minSize2
            minSize2 = minSize2 + (minSize2 \ 2) + 1
            guard = guard + 1
        Loop
    End If

    outMerged = outUniqRaw - n
    GroupChapters = n
End Function

' 隣接する章(文書順)を、合計チャンク数がminSize以上になるまで束ねる。
' 束ねた章の代表キーは、束ねたメンバー全員のキーをCHAPTER_KEY_SEP区切りで
' 連結したもの(単独で条件を満たす章は連結されず、素のキーのまま=通常時は
' 無変化)。
Private Sub CoalesceSmallChapters(ByRef keys() As String, ByRef counts() As Long, _
                                  ByRef n As Long, ByVal minSize As Long)
    If n <= 1 Then Exit Sub
    Dim newKeys() As String: ReDim newKeys(1 To n)
    Dim newCounts() As Long: ReDim newCounts(1 To n)
    Dim outN As Long
    Dim accKey As String, accCount As Long
    Dim i As Long
    For i = 1 To n
        If accCount = 0 Then
            accKey = keys(i)
        Else
            accKey = accKey & CHAPTER_KEY_SEP & keys(i)
        End If
        accCount = accCount + counts(i)
        If accCount >= minSize Or i = n Then
            outN = outN + 1
            newKeys(outN) = accKey
            newCounts(outN) = accCount
            accCount = 0
            accKey = ""
        End If
    Next i
    Dim r As Long
    For r = 1 To outN
        keys(r) = newKeys(r)
        counts(r) = newCounts(r)
    Next r
    n = outN
End Sub

' 切り上げ除算(a/bの天井)。b<=0は異常値として a をそのまま返す(安全側)。
Private Function CeilDivLong(ByVal a As Long, ByVal b As Long) As Long
    If b <= 0 Then CeilDivLong = a: Exit Function
    CeilDivLong = (a + b - 1) \ b
End Function

' ----------------------------------------------------------------------------
' ChapterKeyMatches - 生の章キー(1件)が、章グループのキー(単一 or
'   CoalesceSmallChaptersがCHAPTER_KEY_SEP連結した複数)に属するか。単一
'   キーのグループは従来どおりの完全一致(後方互換: 旧doc_outlineの章キーも
'   そのまま動く)。
' ----------------------------------------------------------------------------
' 2026-08-07(R21H F5・敵対的レビュー確定): vbBinaryCompareだと大小文字/全角
' 半角の違いだけで同じ章が別扱いになり(暗黙の狭窄)、章のチャンクが一部
' しか引けなかった。vbTextCompareへ統一する(GroupChaptersの束ね判定・
' CollectChapterHits/ChapterBodyはこの関数を経由するので自動的に揃う)。
Public Function ChapterKeyMatches(ByVal rawKey As String, ByVal groupKey As String) As Boolean
    If LenB(rawKey) = 0 Or LenB(groupKey) = 0 Then Exit Function
    If InStr(groupKey, CHAPTER_KEY_SEP) = 0 Then
        ChapterKeyMatches = (StrComp(rawKey, groupKey, vbTextCompare) = 0)
        Exit Function
    End If
    Dim members() As String: members = Split(groupKey, CHAPTER_KEY_SEP)
    Dim i As Long
    For i = LBound(members) To UBound(members)
        If StrComp(rawKey, members(i), vbTextCompare) = 0 Then
            ChapterKeyMatches = True
            Exit Function
        End If
    Next i
End Function

' 1章ぶんの本文(文書順に連結し cap で打ち切る)。本文は採用した行だけ
' セル単位で読む(全列を読み直すと本棚全体の本文をもう一度メモリへ載せる)。
Private Function ChapterBody(ByRef rowIdx() As Long, ByRef paths() As String, _
                             ByVal nSrc As Long, ByVal chKey As String, _
                             ByVal cap As Long) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(modAppDef.SH_KNOWLEDGE)
    If ws Is Nothing Then Exit Function

    Dim sb As String
    Dim used As Long
    Dim i As Long
    For i = 1 To nSrc
        If ChapterKeyMatches(ChapterKeyOf(paths(i)), chKey) Then
            Dim t As String
            t = CStr(ws.Cells(rowIdx(i), COL_FULLTEXT).Value)
            Dim take As Long: take = BudgetTake(used, Len(t), cap)
            If take < 1 Then Exit For
            If take < Len(t) Then t = modUtil.SafeLeft(t, take)
            sb = sb & t & vbLf
            used = used + take
        End If
    Next i
    ChapterBody = sb
End Function

' 進捗(バナー+実況行)。ETAは1章あたりの移動平均から出す(modEnrich と同じ
' 部品 modUtilText.BlendPerItemMs / modUtil.EtaText を使う)。
Private Sub Banner(ByVal idx As Long, ByVal total As Long, ByVal sourceName As String, _
                   ByVal msPerItem As Double)
    On Error Resume Next
    Dim t As String
    t = "まとまりの要約中… " & idx & "/" & total & "まとまり(" & modUtil.SafeLeft(sourceName, 40) & ")"
    Dim eta As String: eta = modUtil.EtaText(total - idx + 1, msPerItem)
    If LenB(eta) > 0 Then t = t & " " & eta
    modShelfBatch.StageBanner t
    modUIMain.SetStage t
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' BuildChapterPrompt - 章要約のプロンプト(Private。置き場の理由は冒頭)。
'   出力契約はタグ形式(modRagParse.ParseOutlineResp が読む):
'     <summary>…</summary><keywords>a|b|c</keywords>
'   要約に「具体名を残す」と指示するのは、この要約が【章を選ぶための材料】
'   だから。「本章は手続について定める」とだけ書かれた要約が20本並ぶと、
'   どの章を読むべきかを選ぶ手掛かりが1つも無くなる。
' ----------------------------------------------------------------------------
Private Function BuildChapterPrompt(ByVal sourceName As String, ByVal chKey As String, _
                                    ByVal body As String) As String
    Dim sb As String
    sb = "あなたは社内資料の要約担当です。次に示すのは資料「" & sourceName & "」の" & _
         "「" & chKey & "」の本文です。後で【どの章を読むべきか】を選ぶための" & _
         "材料として、この範囲に何が書いてあるかを要約してください。" & vbLf
    sb = sb & "・要約は200～300字の日本語。対象・原則・例外の順に、具体名" & _
         "(条番号・書類名・期限・金額)を必ず残すこと。" & vbLf
    sb = sb & "・本文に書かれていないことは書かない(推測・一般論・前置き・感想は禁止)。" & vbLf
    sb = sb & "・キーワードは、この範囲を探すときに使われそうな語を最大8個。" & vbLf
    sb = sb & vbLf & "## 本文" & vbLf & body & vbLf & vbLf
    sb = sb & "## 出力形式(この形式のみで出力。説明文・前置きは一切禁止)" & vbLf
    sb = sb & "<summary>200～300字の要約</summary>" & vbLf
    sb = sb & "<keywords>語1|語2|語3</keywords>" & vbLf
    BuildChapterPrompt = sb
End Function
