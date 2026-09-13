Attribute VB_Name = "modVecCache"
Option Explicit

' ============================================================================
' modVecCache - セッション内ベクトルキャッシュ(2026-08-01 R12-4)
' ----------------------------------------------------------------------------
' 役割:
'   my_vectors の vector_csv を「セッション中に1回だけ」パースして Double 配列
'   (次元×行)で保持し、以降の検索は配列参照だけで内積を取れるようにする。
'
' なぜ要るか(R12-4 要件定義の背景):
'   検索は毎クエリ my_vectors 全行に対して modUtil.CsvToVector を呼んでいた。
'   公称上限20,500チャンク×768次元では 1クエリあたり約1,574万要素(内部反復は
'   その10倍規模)のパースになり、既定の多段検索(retrieve_mode=multi)では
'   クエリ本数ぶん繰り返される。監査の定量では1質問60〜300秒。内積そのものは
'   支配項ではなく、支配項は「毎回やり直すCSVパース」だった。
'   同じ答えを出すのに同じ計算を毎回やり直さない、というだけの対策である。
'
' 設計判断:
'   ・スコアの完全一致を最優先する。DotAt の加算順序は modUtil.DotProduct と
'     1バイトも変えない(i=0 から昇順・Double累算)。パース自体も同じ
'     modUtil.CsvToVector を通す。よってキャッシュ経路と直接パース経路の
'     スコアは「近い」ではなく「同一」になる(modTestsPure8 で固定)。
'   ・世代カウンタ(Generation)を持つ。取込・削除・再埋め込みの各書込点が
'     BumpGeneration を呼び、キャッシュは自分が作られた世代と現世代が
'     食い違ったら無効になる。件数と先頭/末尾 chunk_id の印(Stamp)だけでは
'     「同じ行を上書きした再埋め込み」を検知できないため、世代を併用する。
'   ・32bitメモリ: 上限20,500×768次元の Double で約126MB。ReDim が
'     実行時エラー7(メモリ不足)になり得るので必ず捕捉し、キャッシュ無しの
'     従来経路へ自動で戻す(機能は落とさない)。断念した印(mFailStamp)を
'     残し、同じ内容に対して毎回やり直して二重に遅くなることを防ぐ。
'   ・構築中は ShowProgress + DoEvents(PROGRESS_STEP 行ごと)。憲章§3-2。
'   ・キャッシュの解放口(ResetVecCache)を持ち、取込バッチの開始時に呼ぶ。
'     取込は my_knowledge/my_vectors を大きく書き換えるので、古いキャッシュを
'     抱えたままにするとピークメモリが二重になる。
'   ・純ロジック部(StampOf/IsStale/BuildFrom/SlotOfRow/DotAt)はシートに
'     触れないので LO実行テストから直接検証できる。シートを読むのは
'     PrepareVectors だけで、ここが唯一の Excel 依存点である。
' ============================================================================

' 進捗更新+DoEventsの間隔(行)。憲章§3-2「全ての待ち時間に進捗と目安を示す」。
Private Const PROGRESS_STEP As Long = 2000

' 1セルに収まる上限(Excelは32,767字)。照合テキストの保存側と同じ考え方。
Private Const COL_V_ID As Long = 1
Private Const COL_V_VEC As Long = 2

' 埋め込み世代。取込/削除/再埋め込みでインクリメントされる。
Private mGeneration As Long

' キャッシュ本体
Private mReady As Boolean
Private mStamp As String            ' 構築時の印(件数|先頭id|末尾id)
Private mBuiltGen As Long           ' 構築時の世代
Private mDim As Long                ' キャッシュ次元(最初にパースできた行の次元)
Private mSlots As Long              ' 実際に載った本数
Private mRowLo As Long, mRowHi As Long
Private mMismatch As Long           ' 次元不一致で載らなかった行数
Private mFlat() As Double           ' slot*mDim + i
Private mSlotOf() As Long           ' 行r -> slot(>=0) / -1=載らない / -2=次元不一致
Private mDimOf() As Long            ' 行r -> その行の次元(0=パース不可)
Private mRowOf() As Long            ' slot -> 行r
Private mFailStamp As String        ' 構築を断念した印(同じ内容では再試行しない)

' ----------------------------------------------------------------------------
' 世代カウンタ
' ----------------------------------------------------------------------------
Public Function Generation() As Long
    Generation = mGeneration
End Function

' 取込(EmbedPendingの書込)・削除(RemoveVectorsByIds)・再埋め込み
' (MarkAllForReembed)から呼ぶ。呼び忘れると「古いベクトルで検索し続ける」
' という無言の間違いになるため、書込点そのものに置くこと。
Public Sub BumpGeneration()
    On Error Resume Next
    mGeneration = mGeneration + 1
    If mGeneration < 0 Then mGeneration = 1   ' 桁溢れ時も単調に見えれば十分
    mReady = False
    mFailStamp = vbNullString
    On Error GoTo 0
End Sub

' キャッシュを解放する(取込バッチ開始時に呼ぶ=ピークメモリの二重取りを防ぐ)。
Public Sub ResetVecCache()
    On Error Resume Next
    mReady = False
    mStamp = vbNullString
    mFailStamp = vbNullString
    mDim = 0
    mSlots = 0
    mMismatch = 0
    mRowLo = 0
    mRowHi = -1
    Erase mFlat
    Erase mSlotOf
    Erase mDimOf
    Erase mRowOf
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' 純ロジック(LO実行テスト対象)
' ----------------------------------------------------------------------------

' キャッシュの印。件数+先頭/末尾chunk_id。modBitwiseOpt.CacheStamp と同じ考え方。
Public Function StampOf(ByRef ids As Variant) As String
    On Error GoTo Bad
    Dim rLo As Long, rHi As Long
    rLo = LBound(ids, 1): rHi = UBound(ids, 1)
    StampOf = CStr(rHi - rLo + 1) & "|" & CStr(ids(rLo, 1)) & "|" & CStr(ids(rHi, 1))
    Exit Function
Bad:
    StampOf = vbNullString
End Function

' 無効化判定。印か世代のどちらかが違えば作り直す(未構築も無効扱い)。
Public Function IsStale(ByVal built As Boolean, ByVal builtStamp As String, ByVal builtGen As Long, _
                        ByVal curStamp As String, ByVal curGen As Long) As Boolean
    If Not built Then
        IsStale = True
    ElseIf LenB(curStamp) = 0 Then
        IsStale = True
    ElseIf builtStamp <> curStamp Then
        IsStale = True
    ElseIf builtGen <> curGen Then
        IsStale = True
    End If
End Function

' 1列×1行のRange読みはスカラーになるため、常に(1..n, 1..1)の配列へ揃える。
Public Function AsColumnArray(ByVal v As Variant) As Variant
    If IsArray(v) Then
        AsColumnArray = v
        Exit Function
    End If
    Dim one() As Variant
    ReDim one(1 To 1, 1 To 1)
    one(1, 1) = v
    AsColumnArray = one
End Function

' このキャッシュが stamp と現世代に対して使えるか。
Public Function ValidFor(ByVal stampKey As String) As Boolean
    ValidFor = Not IsStale(mReady, mStamp, mBuiltGen, stampKey, mGeneration)
End Function

' 直前に構築を断念した内容か(err7フォールバック中は再試行しない)。
Public Function DeclinedFor(ByVal stampKey As String) As Boolean
    If LenB(stampKey) = 0 Then Exit Function
    DeclinedFor = (mFailStamp = stampKey & "|g" & mGeneration)
End Function

Public Function Ready() As Boolean
    Ready = mReady
End Function

Public Function Stamp() As String
    Stamp = mStamp
End Function

Public Function CachedDim() As Long
    CachedDim = mDim
End Function

Public Function SlotCount() As Long
    SlotCount = mSlots
End Function

' 次元不一致でキャッシュに載らなかった行数(>0ならキャッシュは本棚を代表しない)。
Public Function MismatchCount() As Long
    MismatchCount = mMismatch
End Function

' 行r -> slot。0以上=キャッシュ内の位置 / -1=ベクトル無し・パース不可 /
' -2=次元不一致(呼び出し側がE0702を記録して飛ばす)。
Public Function SlotOfRow(ByVal r As Long) As Long
    SlotOfRow = -1
    If Not mReady Then Exit Function
    If r < mRowLo Or r > mRowHi Then Exit Function
    SlotOfRow = mSlotOf(r)
End Function

' 行rのベクトル次元(0=パースできなかった行)。E0702の記録内容に使う。
Public Function DimOfRow(ByVal r As Long) As Long
    If Not mReady Then Exit Function
    If r < mRowLo Or r > mRowHi Then Exit Function
    DimOfRow = mDimOf(r)
End Function

Public Function RowOfSlot(ByVal s As Long) As Long
    RowOfSlot = -1
    If Not mReady Then Exit Function
    If s < 0 Or s >= mSlots Then Exit Function
    RowOfSlot = mRowOf(s)
End Function

' ----------------------------------------------------------------------------
' DotAt - キャッシュ上のベクトルとqvの内積。
' ----------------------------------------------------------------------------
' 【等価性の要】加算順序・型は modUtil.DotProduct と同一(i=0から昇順・Double)。
' 浮動小数は加算順序が変われば値も変わるため、ここを「速く」しようとして
' 並べ替えたり分割加算にしたりしてはいけない。順位が静かに変わる。
Public Function DotAt(ByRef qv() As Double, ByVal slot As Long) As Double
    If Not mReady Then Exit Function
    If slot < 0 Or slot >= mSlots Then Exit Function
    Dim baseI As Long: baseI = slot * mDim
    Dim lo As Long: lo = LBound(qv)
    If (UBound(qv) - lo + 1) <> mDim Then Exit Function   ' 次元不一致は0(DotProductと同じ)
    Dim s As Double
    Dim i As Long
    For i = 0 To mDim - 1
        s = s + qv(lo + i) * mFlat(baseI + i)
    Next i
    DotAt = s
End Function

' キャッシュ上のベクトルを取り出す(粗選別の量子化・テスト用)。
Public Function VectorAt(ByVal slot As Long, ByRef outVec() As Double) As Boolean
    If Not mReady Then Exit Function
    If slot < 0 Or slot >= mSlots Then Exit Function
    Dim tmp() As Double: ReDim tmp(0 To mDim - 1)
    Dim baseI As Long: baseI = slot * mDim
    Dim i As Long
    For i = 0 To mDim - 1
        tmp(i) = mFlat(baseI + i)
    Next i
    outVec = tmp
    VectorAt = True
End Function

' ----------------------------------------------------------------------------
' BuildFrom - vData(my_vectorsの2列配列)から一括構築する。
'   withProgress=Falseならシート/UIに一切触れない(=純ロジックとしてテスト可能)。
'   戻り値 False = 構築できなかった(呼び出し側は従来経路へ)。
' ----------------------------------------------------------------------------
Public Function BuildFrom(ByRef vData As Variant, ByVal stampKey As String, _
                          ByVal withProgress As Boolean) As Boolean
    On Error GoTo Fail
    ResetVecCache

    Dim rLo As Long, rHi As Long
    rLo = LBound(vData, 1): rHi = UBound(vData, 1)
    Dim nRows As Long: nRows = rHi - rLo + 1
    If nRows < 1 Then GoTo Abort

    ' 次元は「最初にパースできた行」で確定する(modBitwiseOpt.BuildCacheと同じ規約)。
    Dim dimN As Long: dimN = 0
    Dim probe() As Double
    Dim r As Long
    For r = rLo To rHi
        If modUtil.CsvToVector(CStr(vData(r, COL_V_VEC)), probe) Then
            dimN = UBound(probe) - LBound(probe) + 1
            If dimN > 0 Then Exit For
        End If
    Next r
    If dimN < 1 Then GoTo Abort

    ' ここが実行時エラー7(メモリ不足)になり得る唯一の場所。捕捉して従来経路へ。
    ReDim mFlat(0 To nRows * dimN - 1)
    ReDim mSlotOf(rLo To rHi)
    ReDim mDimOf(rLo To rHi)
    ReDim mRowOf(0 To nRows - 1)

    Dim slot As Long: slot = 0
    Dim vv() As Double
    Dim i As Long
    For r = rLo To rHi
        mSlotOf(r) = -1
        mDimOf(r) = 0
        If withProgress Then
            If ((r - rLo) Mod PROGRESS_STEP) = 0 And r > rLo Then Tick r - rLo, nRows
        End If

        Dim vid As String: vid = CStr(vData(r, COL_V_ID))
        Dim vcsv As String: vcsv = CStr(vData(r, COL_V_VEC))
        If LenB(vid) > 0 And LenB(vcsv) > 0 Then
            If modUtil.CsvToVector(vcsv, vv) Then
                Dim d As Long: d = UBound(vv) - LBound(vv) + 1
                mDimOf(r) = d
                If d = dimN Then
                    Dim baseI As Long: baseI = slot * dimN
                    Dim lo As Long: lo = LBound(vv)
                    For i = 0 To dimN - 1
                        mFlat(baseI + i) = vv(lo + i)
                    Next i
                    mSlotOf(r) = slot
                    mRowOf(slot) = r
                    slot = slot + 1
                Else
                    mSlotOf(r) = -2          ' 次元不一致(呼び出し側がE0702を残す)
                    mMismatch = mMismatch + 1
                End If
            End If
        End If
    Next r

    If slot < 1 Then GoTo Abort

    mDim = dimN
    mSlots = slot
    mRowLo = rLo
    mRowHi = rHi
    mStamp = stampKey
    mBuiltGen = mGeneration
    mReady = True
    If withProgress Then
        On Error Resume Next
        modUIMain.HideProgress
        On Error GoTo 0
    End If
    BuildFrom = True
    Exit Function

Abort:
    ' 例外ではない断念(行が無い/1本もパースできない)。ここは Resume を
    ' 使ってはいけない(エラーが起きていないときの Resume は実行時エラー20)。
    ' 2026-08-01(R12-H-10): 例外ではないので err# を書かない。err#0 と書くと
    ' 「原因不明の失敗」に見え、読み手が存在しないエラーを探すことになる。
    On Error GoTo 0
    Cleanup stampKey, "ベクトルの行が無いか、1本も読み取れませんでした", withProgress
    BuildFrom = False
    Exit Function

Fail:
    Dim origNum As Long: origNum = Err.Number
    ' ハンドラ稼働中は On Error Resume Next が効かない。後始末の前に Resume で
    ' ハンドラを抜ける(2026-07-30 実機err#462と同型の作法)。
    Resume FailTail
FailTail:
    Cleanup stampKey, "メモリが足りませんでした(err#" & origNum & ")", withProgress
    BuildFrom = False
End Function

' 断念したときの後始末: 解放して「この内容ではもう試さない」印を残し、記録する。
Private Sub Cleanup(ByVal stampKey As String, ByVal reason As String, ByVal withProgress As Boolean)
    ResetVecCache
    mFailStamp = stampKey & "|g" & mGeneration
    ReportFallback reason, withProgress
End Sub

' 構築を断念したことの記録(無言の失敗禁止・憲章§4-1)。機能は落とさないが
' 「なぜ今日は遅いのか」を後から特定できるよう usage_log に1行だけ残す。
Private Sub ReportFallback(ByVal reason As String, ByVal withProgress As Boolean)
    On Error Resume Next
    modLog.LogUsage "veccache_fallback", "", _
        "ベクトルキャッシュを作れなかったため従来の検索経路で続行します: " & reason
    If withProgress Then modUIMain.HideProgress
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' PrepareVectors - 検索の入口。my_vectorsの読み方をここ1箇所に集約する。
'   戻り値 True  = キャッシュ経路が使える(vDataは読み込まないまま=メモリ節約)
'   戻り値 False = 従来経路(vDataへ2列を読み込んで返す)
'   idData は chunk_id列だけの配列。キャッシュが生きているときは
'   vector_csv列(20,500行なら100MB規模の文字列)を読まずに済む。
' ----------------------------------------------------------------------------
Public Function PrepareVectors(ByVal wsV As Worksheet, ByVal lastV As Long, ByVal qDim As Long, _
                               ByRef idData As Variant, ByRef vData As Variant) As Boolean
    On Error GoTo Fallback
    If wsV Is Nothing Then GoTo Direct
    If IsEmpty(idData) Then
        idData = AsColumnArray(wsV.Range(wsV.Cells(2, COL_V_ID), wsV.Cells(lastV, COL_V_ID)).Value)
    End If

    Dim stampKey As String: stampKey = StampOf(idData)
    If LenB(stampKey) = 0 Then GoTo Direct

    If ValidFor(stampKey) Then
        If qDim <= 0 Or qDim = mDim Then
            PrepareVectors = True
            Exit Function
        End If
        ' 質問側の次元がキャッシュと違う(embed_dim変更後の混在など)。
        ' キャッシュに載っていない次元の行が正解かもしれないので、
        ' このクエリだけは従来経路で厳密に見る(挙動等価を守る)。
        GoTo Direct
    End If

    If DeclinedFor(stampKey) Then GoTo Direct

    If IsEmpty(vData) Then
        vData = wsV.Range(wsV.Cells(2, COL_V_ID), wsV.Cells(lastV, COL_V_VEC)).Value
    End If
    If BuildFrom(vData, stampKey, True) Then
        If qDim <= 0 Or qDim = mDim Then
            vData = Empty        ' 生CSVを解放(キャッシュとの二重保持を避ける)
            PrepareVectors = True
            Exit Function
        End If
    End If

Direct:
    On Error GoTo 0
    LoadDirect wsV, lastV, idData, vData
    PrepareVectors = False
    Exit Function

Fallback:
    ' ここへ来るのは想定外の例外。ハンドラを Resume で抜けてから読み直す。
    Resume Direct2
Direct2:
    LoadDirect wsV, lastV, idData, vData
    PrepareVectors = False
End Function

' 従来経路(毎回パース)で必要な2列を読む。読めなくても検索側が0件で
' 安全に終われるよう、失敗は握って空のまま返す。
' 2026-08-01(R12-H-2): 走査の途中でキャッシュが解放されたとき、検索側が
' 「その場で従来経路へ切り替えて最初からやり直す」ために公開している
' (my_vectors の読み方をこのモジュール1箇所に保つ)。
Public Sub LoadDirectVectors(ByVal wsV As Worksheet, ByVal lastV As Long, _
                             ByRef idData As Variant, ByRef vData As Variant)
    LoadDirect wsV, lastV, idData, vData
End Sub

Private Sub LoadDirect(ByVal wsV As Worksheet, ByVal lastV As Long, _
                       ByRef idData As Variant, ByRef vData As Variant)
    On Error Resume Next
    If wsV Is Nothing Then Exit Sub
    If IsEmpty(idData) Then
        idData = AsColumnArray(wsV.Range(wsV.Cells(2, COL_V_ID), wsV.Cells(lastV, COL_V_ID)).Value)
    End If
    If IsEmpty(vData) Then
        vData = wsV.Range(wsV.Cells(2, COL_V_ID), wsV.Cells(lastV, COL_V_VEC)).Value
    End If
    On Error GoTo 0
End Sub

' 構築中の実況。表示の失敗が構築を止めてはならない(憲章§4-4)。
Private Sub Tick(ByVal doneRows As Long, ByVal totalRows As Long)
    On Error Resume Next
    modUIMain.ShowProgress ChrW(&HD83D) & ChrW(&HDD0D) & " 検索の準備をしています… " & _
        modUtil.ProgressText(doneRows, totalRows, "")
    DoEvents
    On Error GoTo 0
End Sub
