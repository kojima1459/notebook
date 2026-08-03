Attribute VB_Name = "modTestsPure9"
Option Explicit

' ============================================================================
' modTestsPure9 - R12-4(検索スケール恒久対策)の純ロジック回帰テスト
' ----------------------------------------------------------------------------
' なぜ新しいモジュールなのか:
'   modTestsPure8 に足すと28,000字(WARN帯)を超える。憲章§4-6「WARN帯の
'   モジュールに機能を足さない。足す前に分割を裁定する」に従った分割先で、
'   modTestsPure2〜8 と同じ理由・同じ作法である。
'   入口は modTestsPure8.RunAll8 の末尾から呼ばれる Public Sub RunAll9()。
'   この1行が唯一の導線で、消すと本モジュールのテストは「実行されないまま」
'   全部PASSに見える(分割群と同型の事故)。
'
' 固定する事実:
'   ・modVecCache: キャッシュ経路と直接パース経路のスコア【完全一致】。
'     R12-4の検収条件そのもので、加算順序を変える"最適化"を機械で止める。
'   ・modVecCache: 世代カウンタによる無効化。件数も先頭/末尾idも変わらない
'     再埋め込みは、印(Stamp)だけでは検知できない。
'   ・modVecCache: 次元混在時に載らなかった行を-2で区別すること
'     (-1と同じ扱いにするとE0702が出ないまま資料が静かに検索から消える)。
'   ・modBitwiseOpt.ShouldPrefilter: binary_rag_auto の真理値表。
'   ・modSparse.MatchDocText: my_knowledge.norm_text(取込時の前計算)と
'     検索時のその場計算が1文字も違わないこと。
'   ・modExtractorPdf.TempBaseNameFor(2026-08-03 R13-2): 一時コピー名が
'     ASCIIだけで出来ていること・拡張子を落とさないこと・決定的であること。
'   ・optOcrCore.ClassifyGsTextResult(2026-08-03 R13-1b): 「GSは完了したのに
'     本文が空」の理由を image/gsfail/flagdelay のどれに倒すかの真理表。
'     実機第2報 RC1(スキャンPDFがWordのゴミ本文で「登録成功」)の再発防止。
'   ・optOcrCore.GsTotalPagesFromLog / GsPagesFromLog / GsWaitBanner
'     (2026-08-03 R13-1c/1d): gs_out.log から進捗を読む目と、その見せ方。
'   ・modExtractorPdf.IsThinExtract("pdf", 2026-08-03 R13-3a): 「取り込めてはいるが
'     本文が薄すぎる」の閾値。RC1(44頁の約款が chunks=1 で登録成功)の
'     二段目の防衛で、閾値は仕様から動かさない約束をここで固定する。
'   ・modLog.FriendlyFailMsg(2026-08-03 R13-3b): 取込失敗の文言の選び方。
'     RC6(「Wordを開いたままに…」という正確な案内が汎用文言で上書きされ、
'     しかも docx に Ghostscript の確認を求めていた)の再発防止。
' ============================================================================

' テスト用の my_vectors 相当(chunk_id, vector_csv)の2列配列を作る。
Private Function FakeVData() As Variant
    Dim v As Variant
    ReDim v(1 To 5, 1 To 2)
    v(1, 1) = "bs::a::p1::c1": v(1, 2) = "0.1234567,-0.7654321,0.3333333,0.5555555"
    v(2, 1) = "bs::b::p1::c2": v(2, 2) = "-0.9876543,0.1111111,0.2222222,-0.4444444"
    v(3, 1) = "bs::c::p2::c1": v(3, 2) = ""                       ' ベクトル未生成の行
    v(4, 1) = "bs::d::p2::c2": v(4, 2) = "0.5000001,0.5000002,0.4999999,0.4999998"
    v(5, 1) = "bs::e::p3::c1": v(5, 2) = "not,a,number,x"         ' 壊れたCSV
    FakeVData = v
End Function

' ----------------------------------------------------------------------------
' 等価テスト(R12-4 検収条件): キャッシュ経路と直接パース経路のスコアが
'   「近い」ではなく【完全に同じ】であること。
' ----------------------------------------------------------------------------
' 浮動小数の加算は順序が変われば結果も変わる。キャッシュ側(modVecCache.DotAt)を
' 「速くしよう」として加算順序や中間型を変えると、順位が静かに入れ替わる。
' ここは近似(誤差以内)ではなく完全一致(=)で固定する。
Private Sub TestVecCacheScoreEquivalence()
    Dim vData As Variant: vData = FakeVData()
    Dim qv(0 To 3) As Double
    qv(0) = 0.2672612: qv(1) = -0.5345225: qv(2) = 0.8017837: qv(3) = -0.0123456

    modVecCache.ResetVecCache
    modTestRunner.Check "VecCache_構築できる", _
        modVecCache.BuildFrom(vData, "s1", False), "BuildFrom=False"
    modTestRunner.Check "VecCache_載るのはパースできた行だけ", _
        (modVecCache.SlotCount() = 3), "slots=" & modVecCache.SlotCount()
    modTestRunner.Check "VecCache_次元は先頭の有効行で確定", _
        (modVecCache.CachedDim() = 4), "dim=" & modVecCache.CachedDim()

    Dim r As Long, diffN As Long: diffN = 0
    Dim worst As String: worst = ""
    For r = 1 To 5
        Dim slot As Long: slot = modVecCache.SlotOfRow(r)
        Dim vv() As Double
        If modUtil.CsvToVector(CStr(vData(r, 2)), vv) Then
            Dim direct As Double: direct = modUtil.DotProduct(qv, vv)
            Dim cached As Double: cached = modVecCache.DotAt(qv, slot)
            If cached <> direct Then
                diffN = diffN + 1
                worst = "row=" & r & " direct=" & direct & " cached=" & cached
            End If
        Else
            ' パースできない行はキャッシュにも載らない(=従来経路と同じく飛ばす)
            If slot >= 0 Then diffN = diffN + 1
        End If
    Next r
    modTestRunner.Check "VecCache_全行でスコアが完全一致(キャッシュ経路=直接パース経路)", _
        (diffN = 0), "不一致=" & diffN & " " & worst

    ' 順位(top1)まで同じであることも直接見る。「検索結果が変わらない」という
    ' 検収条件は、スコアの一致だけでなく並びで確かめる方が読み手に伝わる。
    Dim bestDirect As String, bestCached As String
    Dim maxD As Double: maxD = -1E+30
    Dim maxC As Double: maxC = -1E+30
    For r = 1 To 5
        Dim vv2() As Double
        If modUtil.CsvToVector(CStr(vData(r, 2)), vv2) Then
            Dim d2 As Double: d2 = modUtil.DotProduct(qv, vv2)
            If d2 > maxD Then
                maxD = d2
                bestDirect = CStr(vData(r, 1))
            End If
            Dim c2 As Double: c2 = modVecCache.DotAt(qv, modVecCache.SlotOfRow(r))
            If c2 > maxC Then
                maxC = c2
                bestCached = CStr(vData(r, 1))
            End If
        End If
    Next r
    modTestRunner.Check "VecCache_top1が一致", _
        (bestDirect = bestCached And LenB(bestDirect) > 0), _
        "direct=" & bestDirect & " cached=" & bestCached
    modVecCache.ResetVecCache
End Sub

' ----------------------------------------------------------------------------
' 世代カウンタと無効化判定(純ロジック)
' ----------------------------------------------------------------------------
' 件数も先頭/末尾idも変わらないのに中身だけ変わる「再埋め込み」は、
' 印(Stamp)だけでは検知できない。世代で必ず落ちることを固定する。
Private Sub TestVecCacheGeneration()
    Dim ids As Variant
    ReDim ids(1 To 3, 1 To 1)
    ids(1, 1) = "bs::a::p1::c1": ids(2, 1) = "bs::b::p1::c2": ids(3, 1) = "bs::c::p2::c1"
    modTestRunner.Check "VecCache_StampOfは件数と先頭末尾idで作る", _
        (modVecCache.StampOf(ids) = "3|bs::a::p1::c1|bs::c::p2::c1"), _
        "stamp=" & modVecCache.StampOf(ids)

    modTestRunner.Check "VecCache_未構築は無効", _
        modVecCache.IsStale(False, "s1", 1, "s1", 1), "built=Falseなのに有効"
    modTestRunner.Check "VecCache_印も世代も同じなら有効", _
        (Not modVecCache.IsStale(True, "s1", 1, "s1", 1)), "同一なのに無効"
    modTestRunner.Check "VecCache_印が変われば無効(取込/削除)", _
        modVecCache.IsStale(True, "s1", 1, "s2", 1), "印違いを見逃した"
    modTestRunner.Check "VecCache_世代が変われば無効(再埋め込み)", _
        modVecCache.IsStale(True, "s1", 1, "s1", 2), "世代違いを見逃した"

    ' 実体でも同じことが起きるか(BumpGenerationで即無効・ResetVecCacheで解放)
    Dim vData As Variant: vData = FakeVData()
    modVecCache.ResetVecCache
    modVecCache.BuildFrom vData, "s1", False
    modTestRunner.Check "VecCache_構築直後は有効", modVecCache.ValidFor("s1"), "直後に無効"
    Dim g0 As Long: g0 = modVecCache.Generation()
    modVecCache.BumpGeneration
    modTestRunner.Check "VecCache_世代は単調に増える", _
        (modVecCache.Generation() = g0 + 1), "g0=" & g0 & " g1=" & modVecCache.Generation()
    modTestRunner.Check "VecCache_世代を進めると同じ印でも無効", _
        (Not modVecCache.ValidFor("s1")), "再埋め込みを検知できていない"

    modVecCache.BuildFrom vData, "s1", False
    modVecCache.ResetVecCache
    modTestRunner.Check "VecCache_解放後は参照できない", _
        (modVecCache.Ready() = False And modVecCache.SlotOfRow(1) = -1), "解放されていない"

    ' 1行1列のRange読みはスカラーになる。配列へ揃える前処理の固定。
    Dim one As Variant: one = modVecCache.AsColumnArray("bs::x::p1::c1")
    modTestRunner.Check "VecCache_単一セル読みも(1,1)配列へ揃う", _
        (IsArray(one) And CStr(one(1, 1)) = "bs::x::p1::c1"), "スカラーのまま"
End Sub

' 次元混在(embed_dim変更後の途中状態)で、載らなかった行が-2で区別されること。
' -1(ベクトル無し)と同じ扱いにすると、E0702が1件も出ないまま検索結果から
' 資料が静かに消える(従来経路はここでE0702を残していた)。
Private Sub TestVecCacheDimMismatch()
    Dim v As Variant
    ReDim v(1 To 3, 1 To 2)
    v(1, 1) = "id1": v(1, 2) = "0.1,0.2,0.3,0.4"
    v(2, 1) = "id2": v(2, 2) = "0.5,0.5,0.5"          ' 次元が違う
    v(3, 1) = "id3": v(3, 2) = "0.9,0.1,0.1,0.1"

    modVecCache.ResetVecCache
    modVecCache.BuildFrom v, "sx", False
    modTestRunner.Check "VecCache_次元不一致行は-2で返る", _
        (modVecCache.SlotOfRow(2) = -2), "slot=" & modVecCache.SlotOfRow(2)
    modTestRunner.Check "VecCache_不一致行の次元を記録する(E0702の記録内容)", _
        (modVecCache.DimOfRow(2) = 3), "dim=" & modVecCache.DimOfRow(2)
    modTestRunner.Check "VecCache_不一致件数を数える(粗選別の辞退判断に使う)", _
        (modVecCache.MismatchCount() = 1), "n=" & modVecCache.MismatchCount()
    modTestRunner.Check "VecCache_正常行は載る", _
        (modVecCache.SlotOfRow(1) >= 0 And modVecCache.SlotOfRow(3) >= 0), "正常行が落ちた"
    modVecCache.ResetVecCache
End Sub

' ----------------------------------------------------------------------------
' binary_rag_auto(R12-4): 粗選別を使うかの判定(configから切り離した純ロジック)
' ----------------------------------------------------------------------------
Private Sub TestPrefilterAutoDecision()
    modTestRunner.Check "粗選別_小規模では自動でも使わない", _
        (modBitwiseOpt.ShouldPrefilter(4999, 5000, False, True) = False), "小規模で作動した"
    modTestRunner.Check "粗選別_大規模かつ自動ONで有効(R12-4の新既定)", _
        modBitwiseOpt.ShouldPrefilter(5000, 5000, False, True), "自動が効いていない"
    modTestRunner.Check "粗選別_明示TRUEは従来どおり有効", _
        modBitwiseOpt.ShouldPrefilter(20500, 5000, True, False), "明示指定が効かない"
    modTestRunner.Check "粗選別_両方OFFなら大規模でも使わない(止める手段が残る)", _
        (modBitwiseOpt.ShouldPrefilter(20500, 5000, False, False) = False), "止められない"
    modTestRunner.Check "粗選別_binary_rag_minが0以下でも既定5000で守る", _
        (modBitwiseOpt.ShouldPrefilter(100, 0, False, True) = False), "閾値が壊れている"
End Sub

' ----------------------------------------------------------------------------
' norm_text(前計算した照合テキスト)の等価性
' ----------------------------------------------------------------------------
' 保存済みの行(第10列に値がある)と未保存の行(その場で計算)で、採点対象の
' 文字列が1文字も違わないことを固定する。ここがズレると、同じ本棚なのに
' 取込時期によって順位が変わる。
Private Sub TestMatchDocTextEquivalence()
    Dim summary As String: summary = "保険金の 支払い"
    Dim keywords As String: keywords = "支払,免責"
    Dim src As String: src = "ＡＢＣ約款2025.pdf"
    Dim body As String: body = "第１２条　保 険 金 は 30 日以内に支払う"

    Dim legacy As String
    legacy = modSparse.CompactForMatch(summary & " " & keywords & " " & src & " " & body)
    Dim precomputed As String
    precomputed = modSparse.MatchDocText(summary, keywords, src, body)
    modTestRunner.Check "norm_text_前計算と従来の照合テキストが完全一致", _
        (precomputed = legacy), "legacy=[" & legacy & "] pre=[" & precomputed & "]"

    Dim keys As String: keys = modSparse.DistinctiveKeys("第12条の保険金はいつ支払われますか")
    modTestRunner.Check "norm_text_採点結果も一致(保存済み行=未保存行)", _
        (modSparse.KeyScore(keys, precomputed) = modSparse.KeyScore(keys, legacy)), _
        "スコアが割れた"

    ' 取込時点(要約・キーワードは空)と富化後では照合テキストが変わる。
    ' だから modEnrich は要約を書いたら norm_text を空へ戻す(作り直させる)。
    Dim atIngest As String: atIngest = modSparse.MatchDocText("", "", src, body)
    modTestRunner.Check "norm_text_富化で内容が変わる(空へ戻す根拠)", _
        (atIngest <> precomputed), "富化前後で同じになっている"
    modTestRunner.Check "norm_text_空白は全て落ちる(PDF字詰めに強い)", _
        (InStr(precomputed, " ") = 0), "空白が残っている"
End Sub

' ----------------------------------------------------------------------------
' modExtractorPdf.TempBaseNameFor - 一時コピー名の導出(R13-2)
' ----------------------------------------------------------------------------
' 実機第2報 RC4: 元ファイル名をそのまま連結した mbtmp_ 名は、Mac由来のNFD分解
' 濁点(U+3099)のようなCP932に無い文字を含むと、クラシックOpen(ANSI経路)が
' 実際に作るファイル名とGhostscriptへ渡す文字列が食い違い /undefinedfilename に
' なる。名前がASCIIだけで出来ていること・拡張子を落とさないこと・同じパスなら
' 必ず同じ名前になることを固定する(ここが崩れると取込が黙って失敗する)。
Private Sub TestTempBaseNameFor()
    ' "か"(U+304B)+結合濁点(U+3099) = Mac(NFD)が作る「が」。CP932に無い文字。
    Dim nfd As String
    nfd = "C:\Users\eigyo\" & ChrW(&H304B&) & ChrW(&H3099&) & "いよう.PDF"

    Dim nm As String: nm = modExtractorPdf.TempBaseNameFor(nfd)
    modTestRunner.Check "TempBaseNameFor_NFD結合文字_ASCIIのみ", _
        IsAsciiOnly(nm), "nm=" & nm
    modTestRunner.Check "TempBaseNameFor_NFD結合文字_書式(mbtmp_+16桁+拡張子)", _
        (Left$(nm, 6) = "mbtmp_") And (Len(nm) = 26) And _
        IsHex16Lower(Mid$(nm, 7, 16)) And (Right$(nm, 4) = ".pdf"), "nm=" & nm
    modTestRunner.Check "TempBaseNameFor_決定性(同じパスは同じ名前)", _
        (modExtractorPdf.TempBaseNameFor(nfd) = nm), "nm=" & nm

    ' 既存のFNV実装(modUtil.Fnv1a64Hex・ゴールデン値固定済み)をそのまま使う。
    ' 別のハッシュへ差し替えると既存の tmp が全部別名になるので機械で止める。
    modTestRunner.Check "TempBaseNameFor_ハッシュ入力はフルパスそのもの", _
        (Mid$(nm, 7, 16) = modUtil.Fnv1a64Hex(nfd)), "nm=" & nm

    ' NFD(か+濁点)とNFC(が)は別のパス。ディスク上も別ファイルなので別名になる。
    Dim nfc As String
    nfc = "C:\Users\eigyo\" & ChrW(&H304C&) & "いよう.PDF"
    modTestRunner.Check "TempBaseNameFor_NFDとNFCは別名", _
        (modExtractorPdf.TempBaseNameFor(nfc) <> nm), "nm=" & nm

    ' 空白入りパス: コマンドラインへ渡す際の引用符崩れを名前側から無くす。
    Dim spaced As String: spaced = "\\pgiofs01\share\E2T\第 3 部\ご 案内 資料.pdf"
    Dim nmSp As String: nmSp = modExtractorPdf.TempBaseNameFor(spaced)
    modTestRunner.Check "TempBaseNameFor_空白入りパス_名前に空白が残らない", _
        IsAsciiOnly(nmSp) And (InStr(nmSp, " ") = 0), "nmSp=" & nmSp

    ' 長いパス: 元の長さに関係なく名前は「mbtmp_+16桁+拡張子」だけ
    ' (%TEMP%側でMAX_PATHを超えないための要)。
    Dim longPath As String
    longPath = "C:\" & String$(300, "a") & "\" & String$(120, ChrW(&H3042&)) & ".docx"
    Dim nmLong As String: nmLong = modExtractorPdf.TempBaseNameFor(longPath)
    modTestRunner.Check "TempBaseNameFor_長パスでも名前の長さは元に依存しない", _
        (Len(nmLong) = 27) And (Right$(nmLong, 5) = ".docx") And IsAsciiOnly(nmLong), _
        "len=" & Len(nmLong) & " nmLong=" & nmLong

    ' 拡張子は抽出の振り分けに使うので必ず残す(小文字化される)。
    Dim nmNoExt As String: nmNoExt = modExtractorPdf.TempBaseNameFor("C:\tmp\README")
    modTestRunner.Check "TempBaseNameFor_拡張子なしパスは点を付けない", _
        (InStr(nmNoExt, ".") = 0) And (Len(nmNoExt) = 22), "nmNoExt=" & nmNoExt

    ' 衝突時の連番(既存方式の維持)。全て "mbtmp_" 始まりなので掃除の
    ' パターンは mbtmp_* のままでよい。
    Dim nm1 As String: nm1 = modExtractorPdf.TempBaseNameFor(nfd, 1)
    modTestRunner.Check "TempBaseNameFor_連番_別名かつmbtmp_始まり", _
        (nm1 <> nm) And (Left$(nm1, 6) = "mbtmp_") And (Right$(nm1, 4) = ".pdf") And _
        IsAsciiOnly(nm1), "nm1=" & nm1
End Sub

' ----------------------------------------------------------------------------
' optOcrCore.ClassifyGsTextResult の真理表(2026-08-03 R13-1b)。
'   実機第2報 RC1 の事故は「finished かつ txt空」を一律E0302にしていたこと。
'   スキャンPDFがOCRへ回らずWord経路へ落ち、Wordのリフローが作ったゴミ本文が
'   ザルゲートを通過して44頁の約款が chunks=1 / done で「登録成功」になった。
'   分類の順序そのものが事故の本体なので、ここで全組み合わせを固定する。
'   特に大事な2点:
'     ・rc不明(-1)を image にしてはならない。実際にはGSが落ちていた資料が
'       「画像PDF」として記録され、原因が永久に分からなくなる。
'     ・rc=0でもページ処理の痕跡が無いなら image ではない(EDRに実行を
'       止められた等、GSがPDFを開く前に終わった場合)。
' ----------------------------------------------------------------------------
Private Sub TestClassifyGsTextResult()
    ' 本文が取れているときは分類の対象外。rcもログの中身も見ない。
    modTestRunner.Check "Classify_本文ありはok(rc=0)", _
        (optOcrCore.ClassifyGsTextResult(0, 1234, True) = "ok"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(0, 1234, True)
    modTestRunner.Check "Classify_本文ありはok(rc非0でもok)", _
        (optOcrCore.ClassifyGsTextResult(1, 1234, False) = "ok"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(1, 1234, False)
    modTestRunner.Check "Classify_本文ありはok(rc不明でもok)", _
        (optOcrCore.ClassifyGsTextResult(-1, 1, False) = "ok"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(-1, 1, False)

    ' 空 x rc=0 x ページ処理あり = 文字層なし(スキャンPDF)。ERR_303→OCRへ。
    modTestRunner.Check "Classify_空_rc0_ページ処理ありはimage", _
        (optOcrCore.ClassifyGsTextResult(0, 0, True) = "image"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(0, 0, True)

    ' 空 x rc=0 x ページ処理なし = GSがPDFを開く前に終わっている。imageではない。
    modTestRunner.Check "Classify_空_rc0_ページ処理なしはgsfail", _
        (optOcrCore.ClassifyGsTextResult(0, 0, False) = "gsfail"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(0, 0, False)

    ' 空 x rc<>0 = GS実行失敗。ページ処理の有無によらず gsfail(E0302)。
    modTestRunner.Check "Classify_空_rc1_ページ処理ありでもgsfail", _
        (optOcrCore.ClassifyGsTextResult(1, 0, True) = "gsfail"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(1, 0, True)
    modTestRunner.Check "Classify_空_rc1_ページ処理なしもgsfail", _
        (optOcrCore.ClassifyGsTextResult(1, 0, False) = "gsfail"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(1, 0, False)
    modTestRunner.Check "Classify_空_rc255(クラッシュ系)はgsfail", _
        (optOcrCore.ClassifyGsTextResult(255, 0, True) = "gsfail"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(255, 0, True)

    ' 空 x rc不明(-1) = フラグ書込み遅延。ページ処理の有無によらず flagdelay。
    ' ここを image に倒すと「GSが落ちた資料」が画像PDF扱いで記録される。
    modTestRunner.Check "Classify_空_rc不明_ページ処理ありはflagdelay", _
        (optOcrCore.ClassifyGsTextResult(-1, 0, True) = "flagdelay"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(-1, 0, True)
    modTestRunner.Check "Classify_空_rc不明_ページ処理なしもflagdelay", _
        (optOcrCore.ClassifyGsTextResult(-1, 0, False) = "flagdelay"), _
        "実際=" & optOcrCore.ClassifyGsTextResult(-1, 0, False)
End Sub

' ----------------------------------------------------------------------------
' gs_out.log からの進捗読み取り(R13-1c)と待ちバナーの文面(R13-1d)。
'   ここが0を返し続けると「進んでいるのに無進捗と判定して打ち切る」ため、
'   実際のGS出力と同じ形(改行はスペースへ潰した後の姿)で固定する。
' ----------------------------------------------------------------------------
Private Sub TestGsLogProgressParse()
    Dim logText As String
    logText = "Processing pages 1 through 44. Page 1 Page 2 Page 3 Page 12 "

    modTestRunner.Check "GsTotalPages_Processing行から総ページ数", _
        (optOcrCore.GsTotalPagesFromLog(logText) = 44), _
        "実際=" & optOcrCore.GsTotalPagesFromLog(logText)
    modTestRunner.Check "GsPages_最大のPage番号を拾う", _
        (optOcrCore.GsPagesFromLog(logText) = 12), _
        "実際=" & optOcrCore.GsPagesFromLog(logText)

    ' "Processing pages" の pages を Page として拾ってはいけない(大小を区別)。
    Dim headOnly As String: headOnly = "Processing pages 1 through 44. "
    modTestRunner.Check "GsPages_Processing行だけなら0", _
        (optOcrCore.GsPagesFromLog(headOnly) = 0), _
        "実際=" & optOcrCore.GsPagesFromLog(headOnly)

    ' 末尾500字だけを渡された場合(Processing行は切れて見えない)。
    Dim tailOnly As String: tailOnly = "Page 41 Page 42 Page 43 Page 44 "
    modTestRunner.Check "GsPages_末尾だけでも最大値が取れる", _
        (optOcrCore.GsPagesFromLog(tailOnly) = 44), _
        "実際=" & optOcrCore.GsPagesFromLog(tailOnly)
    modTestRunner.Check "GsTotalPages_Processing行が無ければ0", _
        (optOcrCore.GsTotalPagesFromLog(tailOnly) = 0), _
        "実際=" & optOcrCore.GsTotalPagesFromLog(tailOnly)

    ' 空・無関係な出力で誤って数字を作らないこと(EDRのエラー文言など)。
    modTestRunner.Check "GsTotalPages_空は0", _
        (optOcrCore.GsTotalPagesFromLog("") = 0), ""
    modTestRunner.Check "GsPages_空は0", _
        (optOcrCore.GsPagesFromLog("") = 0), ""
    Dim denied As String: denied = "アクセスが拒否されました。 "
    modTestRunner.Check "GsPages_無関係な出力は0", _
        (optOcrCore.GsPagesFromLog(denied) = 0), _
        "実際=" & optOcrCore.GsPagesFromLog(denied)
End Sub

Private Sub TestGsWaitBanner()
    ' 総ページ数が読めるまでは経過秒だけ(嘘の分母を出さない)。
    Dim b0 As String: b0 = optOcrCore.GsWaitBanner(0, 0, 3)
    modTestRunner.Check "Banner_総頁不明なら分母を出さない", _
        (InStr(b0, "ページ") = 0) And (InStr(b0, "経過 3秒") > 0), "実際=" & b0

    ' 1-2ページでは残り時間を出さない(実績が足りず表示が跳ねるため)。
    Dim b1 As String: b1 = optOcrCore.GsWaitBanner(2, 44, 10)
    modTestRunner.Check "Banner_2頁では残り時間を出さない", _
        (InStr(b1, "2/44ページ") > 0) And (InStr(b1, "残り約") = 0), "実際=" & b1

    ' 3ページ以上進んだら残り時間を出す。10頁/20秒 → 残り34頁 = 68秒。
    Dim b2 As String: b2 = optOcrCore.GsWaitBanner(10, 44, 20)
    modTestRunner.Check "Banner_頁レートから残り秒を出す", _
        (InStr(b2, "10/44ページ") > 0) And (InStr(b2, "残り約68秒") > 0), "実際=" & b2

    ' 全ページ到達後は残り時間を出さない(0秒や負値を出さない)。
    Dim b3 As String: b3 = optOcrCore.GsWaitBanner(44, 44, 90)
    modTestRunner.Check "Banner_全頁到達で残り時間なし", _
        (InStr(b3, "44/44ページ") > 0) And (InStr(b3, "残り約") = 0), "実際=" & b3

    ' 異常値(負の経過秒・分母超えのページ数)でも壊れた表示を出さない。
    Dim b4 As String: b4 = optOcrCore.GsWaitBanner(99, 44, -5)
    modTestRunner.Check "Banner_異常値でも破綻しない", _
        (InStr(b4, "44/44ページ") > 0) And (InStr(b4, "経過 0秒") > 0), "実際=" & b4
End Sub

' ----------------------------------------------------------------------------
' modExtractorPdf.IsThinExtract("pdf", 2026-08-03 R13-3a): 薄い抽出の検出。
'   実機第2報 RC1 では44ページの約款が chunks=1 / status=done で
'   「登録成功」になっていた。閾値は仕様R13-3aから動かさない約束なので、
'   ここで境界を固定して「正当に薄い資料まで partial にする」改変も止める。
' ----------------------------------------------------------------------------
Private Sub TestIsThinExtract()
    ' 実機の事故そのもの: 44ページ / 1チャンク。
    modTestRunner.Check "Thin_44頁1チャンクは薄い", _
        modExtractorPdf.IsThinExtract("pdf", 44, 1, 30000), "Falseになった"

    ' 9ページ以下は見ない(表紙+ポンチ絵のような正当に薄い資料を守る)。
    modTestRunner.Check "Thin_9頁は判定しない", _
        Not modExtractorPdf.IsThinExtract("pdf", 9, 0, 0), "Trueになった"

    ' 境界: 10ページ・チャンク0件は薄い(10\20=0 なので 0<=0)。
    modTestRunner.Check "Thin_10頁0チャンクは薄い", _
        modExtractorPdf.IsThinExtract("pdf", 10, 0, 99999), "Falseになった"

    ' 境界: 40ページで2チャンクは薄い(40\20=2)、3チャンクなら文字数次第。
    modTestRunner.Check "Thin_40頁2チャンクは薄い", _
        modExtractorPdf.IsThinExtract("pdf", 40, 2, 99999), "Falseになった"
    modTestRunner.Check "Thin_40頁3チャンク_文字十分なら薄くない", _
        Not modExtractorPdf.IsThinExtract("pdf", 40, 3, 2400), "Trueになった"

    ' 文字数側の条件: ページ数x60 未満なら薄い(40x60=2400 が境界)。
    modTestRunner.Check "Thin_文字数が頁x60未満なら薄い", _
        modExtractorPdf.IsThinExtract("pdf", 40, 3, 2399), "Falseになった"

    ' 正常な資料(44ページ・88チャンク・十分な文字数)は薄くない。
    modTestRunner.Check "Thin_正常な資料は薄くない", _
        Not modExtractorPdf.IsThinExtract("pdf", 44, 88, 61600), "Trueになった"

    ' 拡張子による足切り(R13-F7)は modTestsPure10.TestThinExtractExt が持つ
    ' (このモジュールは30,000字上限まで残りが少ないため。憲章§4-6)。
End Sub

' ----------------------------------------------------------------------------
' modLog.FriendlyFailMsg(2026-08-03 R13-3b): 取込失敗の文言の選び方。
'   実機第2報 RC6: DescribeComError が作った「Wordを開いたままにして…」という
'   正確な案内が、E0302の汎用文言(Ghostscript前提)で上書きされて届いて
'   いなかった。しかもその汎用文言は .docx の失敗にも出ていた。
'   ここで固定するのは3点: 案内文の優先採用 / 技術情報を混ぜないこと /
'   docx・doc に Ghostscript の話をしないこと。
' ----------------------------------------------------------------------------
Private Sub TestFriendlyFailMsg()
    ' (1) 行動可能な案内(429・相乗りモード)がある場合は、それを採用する。
    Dim d1 As String
    d1 = "[Word起動/開き方4] この端末ではExcelからWordを起動できないため、" & _
         "すでに開いているWordを使おうとしましたが、Wordが開いていませんでした。" & _
         "Wordを開いたままにして、もう一度お試しください。" & _
         "(詳細: ActiveX component can't create object) [localcopy=ok]"
    Dim m1 As String: m1 = modLog.FriendlyFailMsg("E0302", d1, "docx")
    modTestRunner.Check "FailMsg_案内文を優先採用する", _
        InStr(m1, "Wordを開いたままにして") > 0, "実際=" & m1
    modTestRunner.Check "FailMsg_案内文に技術情報を混ぜない", _
        (InStr(m1, "(詳細:") = 0) And (InStr(m1, "localcopy") = 0), "実際=" & m1
    modTestRunner.Check "FailMsg_案内文にコードを添える", _
        InStr(m1, "(コード: E0302)") > 0, "実際=" & m1

    ' (2) 生のCOM説明文しか無い .docx は、Wordの話だけをする。
    Dim d2 As String: d2 = "[本文取り出し/開き方1] 型が一致しません。 [localcopy=ok]"
    Dim m2 As String: m2 = modLog.FriendlyFailMsg("E0302", d2, "docx")
    modTestRunner.Check "FailMsg_docxにGhostscriptの話をしない", _
        InStr(m2, "Ghostscript") = 0, "実際=" & m2
    modTestRunner.Check "FailMsg_docxはWordを開いて再試行を案内する", _
        (InStr(m2, "Word") > 0) And (InStr(m2, "もう一度お試しください") > 0), "実際=" & m2
    modTestRunner.Check "FailMsg_docxで生のCOM文言を見せない", _
        InStr(m2, "型が一致しません") = 0, "実際=" & m2

    ' .doc も同じ扱い(拡張子だけが違う同じ経路)。
    modTestRunner.Check "FailMsg_docもGhostscriptの話をしない", _
        InStr(modLog.FriendlyFailMsg("E0302", d2, "doc"), "Ghostscript") = 0, "docで混入"

    ' (3) PDFは従来どおり(GhostscriptもWordも試したうえでの結果)。
    Dim m3 As String: m3 = modLog.FriendlyFailMsg("E0302", "GS: 応答なし / Word: x", "pdf")
    modTestRunner.Check "FailMsg_pdfは従来の汎用文言のまま", _
        InStr(m3, "Ghostscript") > 0, "実際=" & m3

    ' (4) 未知のコードでも必ずコード付きの2文で返す(無言にしない)。
    Dim m4 As String: m4 = modLog.FriendlyFailMsg("E9999", "", "txt")
    modTestRunner.Check "FailMsg_未知コードでもコードを添える", _
        (LenB(m4) > 0) And (InStr(m4, "(コード: E9999)") > 0), "実際=" & m4

    ' (5) errDetail が空でも落ちない。
    modTestRunner.Check "FailMsg_errDetail空でも文言を返す", _
        LenB(modLog.FriendlyFailMsg("E0302", "", "pdf")) > 0, "空文字が返った"
End Sub

' 名前がASCII印字可能文字だけで出来ているか(CP932変換で1文字も化けない条件)。
Private Function IsAsciiOnly(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        Dim c As Long: c = AscW(Mid$(s, i, 1))
        If c < 32 Or c > 126 Then Exit Function
    Next i
    IsAsciiOnly = True
End Function

Private Function IsHex16Lower(ByVal s As String) As Boolean
    If Len(s) <> 16 Then Exit Function
    Dim i As Long
    For i = 1 To 16
        If InStr(1, "0123456789abcdef", Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsHex16Lower = True
End Function

Public Sub RunAll9()
    On Error GoTo VecEquivFail
    TestVecCacheScoreEquivalence
NextVecGen:
    On Error GoTo VecGenFail
    TestVecCacheGeneration
NextVecDim:
    On Error GoTo VecDimFail
    TestVecCacheDimMismatch
NextPrefilterAuto:
    On Error GoTo PrefilterAutoFail
    TestPrefilterAutoDecision
NextNormText:
    On Error GoTo NormTextFail
    TestMatchDocTextEquivalence
NextTempName:
    On Error GoTo TempNameFail
    TestTempBaseNameFor
NextClassifyGs:
    On Error GoTo ClassifyGsFail
    TestClassifyGsTextResult
NextGsLogParse:
    On Error GoTo GsLogParseFail
    TestGsLogProgressParse
NextGsBanner:
    On Error GoTo GsBannerFail
    TestGsWaitBanner
NextThin:
    On Error GoTo ThinFail
    TestIsThinExtract
NextFailMsg:
    On Error GoTo FailMsgFail
    TestFriendlyFailMsg
NextRun10:
    ' 2026-08-03(R13-7c): 容量のための分割先(modTestsPure10)。ここが唯一の
    ' 導線で、消すとTeamCodeOf/DeptOf/ビーコンteam列のテストが「実行されない
    ' まま」全部PASSに見える。
    On Error GoTo Run10Fail
    modTestsPure10.RunAll10
NextDone9:
    On Error GoTo 0
    Exit Sub

VecEquivFail:
    modTestRunner.Check "TestVecCacheScoreEquivalence(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextVecGen
VecGenFail:
    modTestRunner.Check "TestVecCacheGeneration(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextVecDim
VecDimFail:
    modTestRunner.Check "TestVecCacheDimMismatch(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPrefilterAuto
PrefilterAutoFail:
    modTestRunner.Check "TestPrefilterAutoDecision(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextNormText
NormTextFail:
    modTestRunner.Check "TestMatchDocTextEquivalence(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextTempName
TempNameFail:
    modTestRunner.Check "TestTempBaseNameFor(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextClassifyGs
ClassifyGsFail:
    modTestRunner.Check "TestClassifyGsTextResult(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGsLogParse
GsLogParseFail:
    modTestRunner.Check "TestGsLogProgressParse(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGsBanner
GsBannerFail:
    modTestRunner.Check "TestGsWaitBanner(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextThin
ThinFail:
    modTestRunner.Check "TestIsThinExtract(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextFailMsg
FailMsgFail:
    modTestRunner.Check "TestFriendlyFailMsg(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRun10
Run10Fail:
    modTestRunner.Check "modTestsPure10.RunAll10(モジュール全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone9
End Sub
