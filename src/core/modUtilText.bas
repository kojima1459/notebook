Attribute VB_Name = "modUtilText"
Option Explicit

' ============================================================================
' modUtilText - テキストファイル入出力と、複数モジュールに散っていた
'   「同じ算数」を1本にまとめるための共通部品置き場(基盤層)。
' ----------------------------------------------------------------------------
' 2026-07-31(R11-F1/F2): 憲章§4-5「同型の問題は共通部品で一度だけ解決する。
'   2箇所で違う答えを出す実装を残さない」に基づく統合先。modUtil は27,822字で
'   余地が少ないため、新しい共通部品はこちらへ置く。
'
' 置くもの / 置かないもの:
'   ・置く: 2箇所以上で同じ答えを出す必要がある処理。ADODB.Stream による
'     UTF-8の読み書き、経過ミリ秒の算出(Timerの日跨ぎ)、1件あたり所要時間の
'     移動平均、GS txtwrite 出力のページ分割添字。
'   ・置かない: 画面(Shape/Range)を触るもの、業務ロジックの判断。
'     基盤層なので上位層(ingest/qa/pack/stats/ui)を参照してはならない。
' ============================================================================

' 1日のミリ秒。VBAのTimerは0時に0へ戻るため、経過が負になったらこれを足す。
Private Const MS_PER_DAY As Double = 86400000#

' ----------------------------------------------------------------------------
' ReadTextFileUtf8 - UTF-8テキストファイルを読む(ADODB.Stream)。
'   戻り値: 成功=True。失敗時は False を返し、outErrNum/outErrDesc に理由を
'   入れる(例外は外へ伝播させない=呼び出し側が必ず分岐で受けられる)。
'   先頭にBOM(U+FEFF)が残っていたら落とす。ADODB.Stream は Charset="utf-8"
'   のとき通常はBOMを落とすが、環境差で残ることが実機で確認されている
'   (旧 optGsTxt.ExtractPdfTextNoOcr の個別対処をここへ集約した)。
'
'   2026-07-31(R11-F2): 同じ形の実装が10箇所にあった(modExtractor/optDiffDoc/
'   optGsTxt/modChannel/modP2PIo/modPublish/modTelemetry/modBoard/modMentor/
'   modInsightIo)。COMの後始末(Resumeでハンドラを抜けてからClose/解放)を
'   1箇所でだけ正しく書けばよい状態にする。
' ----------------------------------------------------------------------------
Public Function ReadTextFileUtf8(ByVal filePath As String, ByRef outText As String, _
                                 Optional ByRef outErrNum As Long, _
                                 Optional ByRef outErrDesc As String) As Boolean
    Dim st As Object
    Dim txt As String

    outText = ""
    outErrNum = 0
    outErrDesc = ""

    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    txt = CStr(st.ReadText(-1))   ' -1 = adReadAll
    st.Close
    Set st = Nothing

    If Len(txt) > 0 Then
        If Left$(txt, 1) = ChrW(&HFEFF) Then txt = Mid$(txt, 2)
    End If
    outText = txt
    ReadTextFileUtf8 = True
    Exit Function

Fail:
    ' Err は Resume でクリアされるので先に控える。
    outErrNum = Err.Number
    outErrDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きたエラーは
    ' 呼び出し元へ飛んで本来の原因を上書きする。後始末の前に Resume で
    ' ハンドラを抜ける(2026-07-30 実機err#462)。
    Resume ReadCleanup
ReadCleanup:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing     ' COM解放(異常パス=半開きも確実に解放)
    On Error GoTo 0
    outText = ""
    ReadTextFileUtf8 = False
End Function

' ----------------------------------------------------------------------------
' WriteTextFileUtf8 - UTF-8テキストファイルを書く(ADODB.Stream・既存は上書き)。
'   戻り値: 成功=True。失敗時は False を返し、outErrNum/outErrDesc に理由を
'   入れる。BOMは ADODB.Stream の既定どおり付く(Excelが日本語CSVを正しく
'   開けるのはこのBOMのおかげなので、落とさない。modAnalytics のCSV書き出しが
'   これに依存している)。
'
'   2026-07-31(R11-F2): 同じ形の実装が9箇所にあった(modChannel/modP2PIo/
'   modPublish/modTelemetry/modAnalytics/modBoard/modMentor/modVault/
'   modInsightIo)。
' ----------------------------------------------------------------------------
Public Function WriteTextFileUtf8(ByVal filePath As String, ByVal content As String, _
                                  Optional ByRef outErrNum As Long, _
                                  Optional ByRef outErrDesc As String) As Boolean
    Dim st As Object

    outErrNum = 0
    outErrDesc = ""

    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2          ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    st.Close
    Set st = Nothing     ' COM解放(正常パス)
    WriteTextFileUtf8 = True
    Exit Function

Fail:
    outErrNum = Err.Number
    outErrDesc = Err.Description
    Resume WriteCleanup
WriteCleanup:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Set st = Nothing     ' COM解放(異常パス=半開きも確実に解放)
    On Error GoTo 0
    WriteTextFileUtf8 = False
End Function

' ----------------------------------------------------------------------------
' ElapsedMsSince - t0(Timerの値)からの経過ミリ秒。
'   VBAの Timer は0時に0へ戻るため、素の (Timer - t0) は日付をまたいだ瞬間に
'   大きな負の値になる。latency_ms が負になるとログの集計が壊れ、進捗の残り
'   時間見積りも跳ねる。負になったら1日ぶんを足して実経過へ戻す。
'   2026-07-31(R11-F2): modGateway 12箇所・modAsk 2箇所が素の式で、その他
'   9モジュールは各自でガードしていた(答えが2種類ある状態)。ここへ一本化する。
' ----------------------------------------------------------------------------
Public Function ElapsedMsSince(ByVal t0 As Double) As Double
    Dim ms As Double
    ms = (Timer - t0) * 1000#
    If ms < 0 Then ms = ms + MS_PER_DAY     ' 日跨ぎ(Timerが0へ戻った)
    If ms < 0 Then ms = 0                   ' それでも負なら0(時計の巻き戻し)
    ElapsedMsSince = ms
End Function

' ----------------------------------------------------------------------------
' BlendPerItemMs - 1件あたり所要ミリ秒の更新(R7 B-1)。直近バッチの実測と
'   現在値の平均(=直近2バッチの移動平均)を返す。
'   2026-07-31(R11-F2): modEmbed と modEnrich に同じ実装が2本あった。
'   経過時間の算出は ElapsedMsSince に任せる(日跨ぎもそこで吸収される)。
' ----------------------------------------------------------------------------
Public Function BlendPerItemMs(ByVal curMs As Double, ByVal t0 As Double, _
                               ByVal itemCount As Long) As Double
    BlendPerItemMs = curMs
    If itemCount < 1 Then Exit Function

    Dim thisMs As Double
    thisMs = ElapsedMsSince(t0) / CDbl(itemCount)
    If curMs <= 0 Then
        BlendPerItemMs = thisMs
    Else
        BlendPerItemMs = (curMs + thisMs) / 2#
    End If
End Function

' ----------------------------------------------------------------------------
' GsPageIsBlank - 空白類しか無いページか。
'   Trim$ は半角スペースしか落とさないため使わない(改行だけのページを
'   「中身あり」と数えると出典ページ番号が丸ごと1つずれる)。
' ----------------------------------------------------------------------------
Public Function GsPageIsBlank(ByVal s As String) As Boolean
    GsPageIsBlank = (CleanTextLen(s) = 0)
End Function

' ----------------------------------------------------------------------------
' CleanTextLen - 空白類(改行/タブ/改ページ/半角空白/全角空白)を除いた文字数。
' ----------------------------------------------------------------------------
Public Function CleanTextLen(ByVal s As String) As Long
    Dim t As String

    t = Replace(s, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, Chr$(12), "")   ' 改ページ(GSのtxtwriteはページをChr(12)で区切る)
    t = Replace(t, " ", "")
    t = Replace(t, ChrW(&H3000), "")
    CleanTextLen = Len(t)
End Function

' ----------------------------------------------------------------------------
' GsPageBounds - Ghostscript txtwrite の出力を改ページで割ったとき、
'   実質何ページぶんあるかと、その範囲を返す。
'   戻り値: ページ数(中身が全く無ければ0)。
'   firstIdx/lastIdx: Split(txt, Chr(12)) の結果配列における先頭/末尾の添字
'   (先頭側・末尾側の空ページを対称に読み飛ばした位置)。中間の空ページは
'   落とさない(本当に白紙のページがあり得るので、落とすと以降のページ番号が
'   全部ずれる)。
'
'   2026-07-31(R11-F2): この添字計算は optOcrCore.GsPageCount(採否しきい値の
'   分母)と modExtractor.BuildPagesFromGsText(出典ページ番号)に同じものが
'   2本あり、ズレると「出典 p.5 を開いても別のページが出る」という誰も
'   気付けない壊れ方をするため、突き合わせテストで守っていた。実装を1本に
'   したので、突き合わせではなくこの関数の単体テストで固定する。
' ----------------------------------------------------------------------------
Public Function GsPageBounds(ByVal txt As String, ByRef firstIdx As Long, _
                             ByRef lastIdx As Long) As Long
    firstIdx = 0
    lastIdx = -1
    If CleanTextLen(txt) = 0 Then Exit Function   ' 全体が空白類だけ=0ページ

    Dim parts() As String
    parts = Split(txt, Chr$(12))
    firstIdx = LBound(parts)
    lastIdx = UBound(parts)

    ' 上の早期returnにより、必ずどこかに中身のあるページがある=ループは必ず解ける。
    Do While firstIdx < lastIdx
        If CleanTextLen(parts(firstIdx)) > 0 Then Exit Do
        firstIdx = firstIdx + 1
    Loop
    Do While lastIdx > firstIdx
        If CleanTextLen(parts(lastIdx)) > 0 Then Exit Do
        lastIdx = lastIdx - 1
    Loop

    GsPageBounds = lastIdx - firstIdx + 1
End Function
