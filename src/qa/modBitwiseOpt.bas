Attribute VB_Name = "modBitwiseOpt"
Option Explicit

' ============================================================================
' modBitwiseOpt - バイナリ量子化ハイブリッドRAG(狂気案Lv.1)。完全独立の最適化層。
' ----------------------------------------------------------------------------
' 役割:
'   大規模ナレッジ(数万～10万件)でも高速に検索するため、埋め込みベクトルを
'   「符号ビット(v>=0→1)」で1bit量子化し、Long型(32bit整数)配列へパックする。
'   XOR+ビットカウント(popcount)でハミング距離を高速計算して上位N件を粗選別し、
'   その上位に対してだけ従来のFloatコサイン(modRetrieveの中核)を適用する。
'   → 全件のCSVパース+内積という重処理を上位N件に限定でき、精度を保ったまま高速化。
'
' 設計判断(最重要=既存Float中核を1バイトも壊さない):
'   ・本モジュールはmodRetrieveの中核ロジックを変更しない。modRetrieve.Searchは
'     ループ先頭に「候補ガード1行」を足すだけで、スコアリング/ランキングは不変。
'     binary_rag=FALSE(既定)またはPrefilterが辞退した場合は従来どおり全件Floatスキャン。
'   ・量子化は符号ベース(v(i)>=0→bit1)。L2正規化済み埋め込みは0付近に分布するため
'     符号量子化がハミング距離とコサイン順位の良い近似になる(標準的なBinary Quantization)。
'   ・VBAのLongは符号付き32bit。bit31=符号ビットの扱いを誤ると全計算が狂うため、
'     popcountは「符号ビットを+1で数え、残り0..30bitをバイト表で数える」符号安全実装。
'     アルゴリズムはPythonリファレンスとdim768/1536/3072・全正/全負/bit31境界で
'     一致検証済み(オフライン)。
'   ・バイナリコードはセッション内メモリキャッシュ(mCodes)に保持。キャッシュ鍵は
'     my_vectorsの件数+先頭/末尾chunk_idで、ナレッジが変わると自動再構築する。
'     my_vectorsのスキーマ・modEmbed・ビルドには一切手を加えない(完全独立)。
'   ・次元不一致・パース失敗・小規模KBでは静かに辞退(Prefilter=False)し、
'     呼び出し側が従来の厳密Floatへフォールバックする(安全側に倒す)。
'   ・R4対象外(PURE_LOGIC_MODULESに含めない)。Dictionary/配列を使うため。
' ============================================================================

' 粗選別(バイナリハミング距離)で残す候補数の既定値。実機のスペック・データ量に
' 合わせて後からここ1箇所で即チューニング可能(config binary_rag_prefilter で
' 実行時上書きも可)。topK(quick=6/deep=12)より十分大きくFloat再ランクの再現率を担保。
Public Const TOP_K_ROUGH As Long = 200

Private mBit(0 To 31) As Long        ' ビットマスク(bit31は&H80000000=符号ビット)
Private mPop(0 To 255) As Long       ' バイトpopcountテーブル
Private mTablesReady As Boolean

' セッション内バイナリコードキャッシュ
Private mNLongs As Long              ' 1ベクトルあたりのLong数(=ceil(dim/32))
Private mNCached As Long             ' キャッシュ済みベクトル数
Private mCodes() As Long             ' フラット格納: slot s は mCodes(s*mNLongs .. +mNLongs-1)
Private mRowOf() As Long             ' mRowOf(s) = vData上の行番号(r)
Private mStamp As String             ' キャッシュ鍵(件数+先頭末尾id)
Private mLastPerf As String          ' 直近の爆速証明ログ(UI層がConsumePerfLogで取得)

' ----------------------------------------------------------------------------
' Enabled - この検索でバイナリ粗選別を使うか。config binary_rag=TRUE かつ
'   件数 >= binary_rag_min のときだけ有効(小規模KBは厳密Floatのまま=挙動不変)。
' ----------------------------------------------------------------------------
Public Function Enabled(ByVal rowCount As Long) As Boolean
    On Error GoTo Off
    If Not modConfig.GetBool("binary_rag", False) Then Exit Function
    Dim minN As Long: minN = modConfig.GetLong("binary_rag_min", 5000)
    If rowCount < minN Then Exit Function
    Enabled = True
    Exit Function
Off:
    Enabled = False
End Function

' 粗選別で残す候補数。既定は定数TOP_K_ROUGH、config binary_rag_prefilterで実行時上書き可。
Public Function PrefilterN() As Long
    On Error Resume Next
    PrefilterN = modConfig.GetLong("binary_rag_prefilter", TOP_K_ROUGH)
    On Error GoTo 0
    If PrefilterN < 1 Then PrefilterN = TOP_K_ROUGH
End Function

' ----------------------------------------------------------------------------
' MicroTimerMs - 経過時間計測用のミリ秒タイマー(Timer基準)。Windows APIの
'   GetTickCount/QueryPerformanceCounterはDeclareが必要で32/64bit宣言地雷+本
'   コードベースの無API方針に反するため、bitness非依存のTimer(精度~15ms)を使う。
'   検索が数百ms規模の大規模デモでは十分。深夜0時のロールオーバーは呼び出し側で
'   差分が負なら0に丸める。
' ----------------------------------------------------------------------------
Public Function MicroTimerMs() As Double
    MicroTimerMs = Timer * 1000#
End Function

' ----------------------------------------------------------------------------
' LogPerf - 「爆速証明」ログ。バイナリ粗選別[X]ms→Float再ランク[Y]ms/候補[N]/総[Z]件を
'   イミディエイトウィンドウ(Debug.Print)へ出力し、binary_rag_debug=TRUEのときは
'   Toastでも画面に出す(審査員向けの可視化)。数値がマイナス(0時跨ぎ)なら0に丸める。
' ----------------------------------------------------------------------------
Public Sub LogPerf(ByVal binMs As Double, ByVal floatMs As Double, _
                   ByVal candN As Long, ByVal totalN As Long)
    On Error Resume Next
    If binMs < 0 Then binMs = 0
    If floatMs < 0 Then floatMs = 0
    Dim msg As String
    msg = ChrW(&HD83D) & ChrW(&HDD0D) & " ハイブリッド検索: バイナリ選別 " & Format$(binMs, "0") & "ms" & _
          " -> Float再ランク " & Format$(floatMs, "0") & "ms / 候補 " & candN & "件 / 総 " & totalN & "件"
    Debug.Print msg
    mLastPerf = msg   ' UI層(modApp)がConsumePerfLogで取り出してToast表示する(層の向きを守る)
    On Error GoTo 0
End Sub

' 直近のLogPerfメッセージを取り出して消す(1回だけToast表示させ、古い値の再表示を防ぐ)。
Public Function ConsumePerfLog() As String
    ConsumePerfLog = mLastPerf
    mLastPerf = vbNullString
End Function

' ----------------------------------------------------------------------------
' Prefilter - qvに対しハミング距離で上位prefilterN件を粗選別し、候補行(vData行番号)を
'   candRows(Dictionary: r -> True)へ入れる。適用できたらTrue、辞退時False(全件Floatへ)。
'   vData: modRetrieveが読み込んだ my_vectors の (chunk_id, vector_csv) 2列配列。
' ----------------------------------------------------------------------------
Public Function Prefilter(ByRef qv() As Double, ByRef vData As Variant, _
                          ByVal poolN As Long, ByRef candRows As Object) As Boolean
    On Error GoTo Fail
    EnsureTables

    If Not BuildCache(vData) Then GoTo Fail
    If mNCached < 1 Then GoTo Fail

    Dim qcode() As Long
    QuantizeToLongs qv, qcode
    If (UBound(qcode) - LBound(qcode) + 1) <> mNLongs Then GoTo Fail   ' 次元不一致→辞退

    Dim n As Long: n = poolN
    If n < 1 Then n = 1
    If n >= mNCached Then
        ' 候補数がキャッシュ全件以上なら粗選別の意味が無い→全件を候補に
        Set candRows = CreateObject("Scripting.Dictionary")
        Dim s0 As Long
        For s0 = 0 To mNCached - 1
            candRows(mRowOf(s0)) = True
        Next s0
        Prefilter = True
        Exit Function
    End If

    ' ストリーミングtop-N(ハミング距離が小さいN件を保持。最大距離のスロットを追跡)。
    Dim keepRow() As Long: ReDim keepRow(0 To n - 1)
    Dim keepDist() As Long: ReDim keepDist(0 To n - 1)
    Dim filled As Long: filled = 0
    Dim worstIdx As Long: worstIdx = 0
    Dim worstDist As Long: worstDist = 0

    Dim s As Long
    For s = 0 To mNCached - 1
        Dim d As Long
        d = HammingAt(qcode, s)
        If filled < n Then
            keepRow(filled) = mRowOf(s)
            keepDist(filled) = d
            filled = filled + 1
            If filled = n Then RecomputeWorst keepDist, n, worstIdx, worstDist
        ElseIf d < worstDist Then
            keepRow(worstIdx) = mRowOf(s)
            keepDist(worstIdx) = d
            RecomputeWorst keepDist, n, worstIdx, worstDist
        End If
    Next s

    Set candRows = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To filled - 1
        candRows(keepRow(i)) = True
    Next i
    Prefilter = True
    Exit Function

Fail:
    Prefilter = False
End Function

' ----------------------------------------------------------------------------
' 内部: バイナリコードキャッシュ
' ----------------------------------------------------------------------------
Private Function BuildCache(ByRef vData As Variant) As Boolean
    Dim stamp As String: stamp = CacheStamp(vData)
    If mStamp = stamp And mNCached > 0 Then
        BuildCache = True
        Exit Function
    End If

    Dim rLo As Long, rHi As Long
    rLo = LBound(vData, 1): rHi = UBound(vData, 1)

    ' 次元をまず1件確定(nLongs算出)
    Dim nLongs As Long: nLongs = 0
    Dim r As Long
    For r = rLo To rHi
        Dim probe() As Double
        If modUtil.CsvToVector(CStr(vData(r, 2)), probe) Then
            Dim dimN As Long: dimN = UBound(probe) - LBound(probe) + 1
            If dimN > 0 Then
                nLongs = (dimN + 31) \ 32
                Exit For
            End If
        End If
    Next r
    If nLongs < 1 Then Exit Function

    Dim maxRows As Long: maxRows = rHi - rLo + 1
    ReDim mCodes(0 To maxRows * nLongs - 1)
    ReDim mRowOf(0 To maxRows - 1)
    Dim slot As Long: slot = 0

    For r = rLo To rHi
        Dim vid As String: vid = CStr(vData(r, 1))
        If LenB(vid) = 0 Then GoTo NextRow
        Dim vcsv As String: vcsv = CStr(vData(r, 2))
        If LenB(vcsv) = 0 Then GoTo NextRow
        Dim vv() As Double
        If Not modUtil.CsvToVector(vcsv, vv) Then GoTo NextRow
        If (UBound(vv) - LBound(vv) + 1 + 31) \ 32 <> nLongs Then GoTo NextRow   ' 次元不一致はキャッシュ対象外

        Dim code() As Long
        QuantizeToLongs vv, code
        Dim j As Long
        For j = 0 To nLongs - 1
            mCodes(slot * nLongs + j) = code(j)
        Next j
        mRowOf(slot) = r
        slot = slot + 1
NextRow:
    Next r

    mNLongs = nLongs
    mNCached = slot
    mStamp = stamp
    BuildCache = (slot > 0)
End Function

' キャッシュ鍵: 件数+先頭/末尾chunk_id(ナレッジが変わると変化→自動再構築)。
Private Function CacheStamp(ByRef vData As Variant) As String
    Dim rLo As Long, rHi As Long
    rLo = LBound(vData, 1): rHi = UBound(vData, 1)
    CacheStamp = CStr(rHi - rLo + 1) & "|" & CStr(vData(rLo, 1)) & "|" & CStr(vData(rHi, 1))
End Function

' キャッシュ内スロットsのコードとqcodeのハミング距離。
Private Function HammingAt(ByRef qcode() As Long, ByVal slot As Long) As Long
    Dim baseI As Long: baseI = slot * mNLongs
    Dim h As Long: h = 0
    Dim j As Long
    For j = 0 To mNLongs - 1
        h = h + PopcountLong(qcode(LBound(qcode) + j) Xor mCodes(baseI + j))
    Next j
    HammingAt = h
End Function

' ----------------------------------------------------------------------------
' 内部: ビット演算コア(符号付きLong安全。Pythonリファレンスと一致検証済み)
' ----------------------------------------------------------------------------

' 符号ビット量子化: v(i)>=0 のビットを立て、32bitずつLong配列へパックする。
Public Sub QuantizeToLongs(ByRef v() As Double, ByRef code() As Long)
    EnsureTables
    Dim lo As Long: lo = LBound(v)
    Dim dimN As Long: dimN = UBound(v) - lo + 1
    Dim nLongs As Long: nLongs = (dimN + 31) \ 32
    ReDim code(0 To nLongs - 1)   ' ReDimで0クリア
    Dim i As Long
    For i = 0 To dimN - 1
        If v(lo + i) >= 0# Then
            code(i \ 32) = code(i \ 32) Or mBit(i Mod 32)
        End If
    Next i
End Sub

' 符号付きLongのpopcount。bit31(符号)は+1で数え、残り0..30bitをバイト表で数える。
Public Function PopcountLong(ByVal x As Long) As Long
    Dim c As Long: c = 0
    If x < 0 Then
        c = 1
        x = x And &H7FFFFFFF   ' 符号ビットを落とす→x>=0(bits0..30)
    End If
    PopcountLong = c + mPop(x And &HFF&) _
                     + mPop((x \ &H100&) And &HFF&) _
                     + mPop((x \ &H10000) And &HFF&) _
                     + mPop((x \ &H1000000) And &HFF&)
End Function

' 2つのLong配列のハミング距離(XOR→popcountの総和)。
Public Function HammingLongs(ByRef a() As Long, ByRef b() As Long) As Long
    EnsureTables
    Dim h As Long: h = 0
    Dim i As Long
    For i = LBound(a) To UBound(a)
        h = h + PopcountLong(a(i) Xor b(LBound(b) + (i - LBound(a))))
    Next i
    HammingLongs = h
End Function

Private Sub EnsureTables()
    If mTablesReady Then Exit Sub
    ' ビットマスク: 0..30 は 1<<i、31 は &H80000000(符号ビット)
    mBit(0) = 1
    Dim i As Long
    For i = 1 To 30
        mBit(i) = mBit(i - 1) * 2
    Next i
    mBit(31) = &H80000000
    ' バイトpopcountテーブル
    For i = 0 To 255
        Dim v As Long: v = i
        Dim cc As Long: cc = 0
        Do While v > 0
            cc = cc + (v And 1)
            v = v \ 2
        Loop
        mPop(i) = cc
    Next i
    mTablesReady = True
End Sub

' 保持中のtop-N内で最大ハミング距離のスロットと値を再計算(ストリーミング用)。
Private Sub RecomputeWorst(ByRef dist() As Long, ByVal n As Long, _
                           ByRef outIdx As Long, ByRef outDist As Long)
    outIdx = 0
    outDist = dist(0)
    Dim i As Long
    For i = 1 To n - 1
        If dist(i) > outDist Then
            outDist = dist(i)
            outIdx = i
        End If
    Next i
End Sub
