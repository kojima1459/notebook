Attribute VB_Name = "modUtil"
Option Explicit

' ============================================================================
' modUtil - 純ロジックのユーティリティ集(R4準拠)
' ----------------------------------------------------------------------------
' 役割:
'   文字列/数値/ベクトルまわりの小さな純関数を集約する。ハッシュ・CSV変換・
'   ベクトル演算・時刻表示・パス解析など、Excelを一切必要としない処理のみ。
'
' 設計判断(R4: Excelオブジェクト禁止):
'   ・Worksheets/Range/Application/ThisWorkbook/MsgBox には一切触れない。
'     このモジュールは LibreOffice headless でもそのまま実行できる
'     (tools/run_lo_tests.py の RunAllPureTests 対象)。
'   ・Now()/Format$()/Timer など「Excelオブジェクトではないが実行環境に
'     依存するVBAランタイム関数」はR4の対象外として許可する
'     (MASTER_SPEC §7.1 modUtil注記のとおり)。
'   ・数値⇔文字列の相互変換はロケール非依存にするため、CStr/Format$では
'     なく Str$/Val を使う(Str$は常に "." を小数点として出力し先頭に
'     符号用の空白を付ける仕様のため Trim$ と組み合わせる。Valは "." を
'     小数点として解釈しロケール設定の影響を受けない。V2 modEmbeddings.bas
'     のコメントと同じ理由)。
'
' ■ Fnv1a64Hex の実装方針(VBAにUInt64が無いことへの対処)
'   64bit値を「上位32bit・下位32bitの2つのLong(ビットパターンとして解釈。
'   符号は無視)」で保持する。32bit同士の掛け算は16bit単位に分割して
'   Double(53bitまで整数を誤差なく表現できる)で計算し、桁上げ(キャリー)
'   を手動で伝播させることで、VBAの `*` 演算子によるオーバーフロー例外
'   (実行時エラー6)を一切発生させずに計算する。アルゴリズムの詳細は
'   Fnv1a64Hex関数の直前のコメント、および内部関数 MulU32 / MulU64ByPrime
'   のコメントを参照。Python(整数演算)で同一アルゴリズムを再現し、複数の
'   テスト文字列(空文字/ASCII/日本語/絵文字サロゲートペア/5000字)で
'   愚直なFNV実装と一致することを確認済み。
' ============================================================================

' 注意(2026-07-11 Wave3で特定・LO互換のための最小修正): LibreOffice Basicは
' モジュールレベルの Private Const を「宣言より前の行での参照」を解決できず
' 実行時エラー12 "Variable not defined" になることを実験で確認した
' (VBAでは宣言順序に関係なくモジュール内どこからでも参照できるため、
' Excel実機ではこの問題は起きない)。Fnv1a64HexはFNV_OFFSET_HI等をモジュール
' 末尾の定数より先に使っていたため、LO実行時テストで検出された。
' 定数の値そのものは変えず、単に「最初の使用箇所より前」に宣言位置を
' 移動しただけである(ロジックへの影響なし)。
Private Const FNV_OFFSET_HI As Long = &HCBF29CE4
Private Const FNV_OFFSET_LO As Long = &H84222325
Private Const FNV_PRIME_HI As Long = &H100    ' = 256
Private Const FNV_PRIME_LO As Long = &H1B3    ' = 435

' DeobfuscateSecret用(下記関数の直前ではなくモジュール先頭に置く。理由は
' すぐ上のコメントと同じ: LO Basicは宣言より前の行での参照を解決できない)。
Private Const OBF_PREFIX As String = "OBF1:"
Private Const OBF_KEY As String = "NexusAgentBuildObfuscationKey2026"

' JoinPagedText/SplitPagedText用のマーカー(同上の理由でモジュール先頭に置く)。
' 「1本の文字列」しか受け渡せない境界(modFeatures.InvokeFeatureの戻り値)を
' 越えて複数ページを運ぶための行単位マーカー。行頭から行末までが完全一致した
' 行だけをマーカーとして扱うので、本文中に似た文字列があっても壊れない
' (最悪でもそこでページが割れるだけ。本文は1文字も失わない)。
Private Const PAGED_MARK_PRE As String = "@@NEXUS_PAGE:"
Private Const PAGED_MARK_SUF As String = "@@"
Private Const PAGED_TRUNC_MARK As String = "@@NEXUS_TRUNCATED@@"

' ============================================================================
' Fnv1a64Hex - FNV-1a 64bit ハッシュを16進16桁(小文字)の文字列で返す。
'   決定的: 同一の入力文字列 s に対して常に同一の出力を返す。
' ----------------------------------------------------------------------------
' ■ アルゴリズム(FNV-1a 64bit、標準定義に準拠)
'   offset_basis = 0xCBF29CE484222325
'   prime        = 0x100000001B3            (= 2^40 + 435)
'   文字列 s を先頭から1文字ずつ、UTF-16コードユニット(VBAのAscW値。
'   負値は+65536して0～65535の符号なし相当に補正)として取り出し、
'   各コードユニットを「下位バイト→上位バイト」の順に2バイトとして
'   FNVへ投入する:
'     hash = offset_basis
'     for each byte b (下位バイトが先):
'       hash = hash Xor b            ' 下位32bitワードの最下位バイトのみ影響
'       hash = (hash * prime) Mod 2^64
'   最終的な hash(上位32bit:下位32bit)を8桁+8桁=16桁の16進文字列にする。
' ============================================================================
Public Function Fnv1a64Hex(ByVal s As String) As String
    Dim hHi As Long: hHi = FNV_OFFSET_HI
    Dim hLo As Long: hLo = FNV_OFFSET_LO

    Dim n As Long: n = Len(s)
    Dim i As Long
    For i = 1 To n
        Dim code As Long: code = AscW(Mid$(s, i, 1))
        If code < 0 Then code = code + 65536   ' 符号なし0～65535に補正

        Dim byteLo As Long: byteLo = code And &HFF
        Dim byteHi As Long: byteHi = (code \ 256) And &HFF

        ' UTF-16コードユニット1個=2バイトとして、下位バイト→上位バイトの順に投入
        hLo = hLo Xor byteLo
        MulU64ByPrime hHi, hLo
        hLo = hLo Xor byteHi
        MulU64ByPrime hHi, hLo
    Next i

    Fnv1a64Hex = U32ToHex8(hHi) & U32ToHex8(hLo)
End Function

' 空白圧縮(半角空白・タブの連続を1個に)+改行統一(CRLF/CRをLFへ)+Trim。
' 改行(LF)自体は圧縮せず保持する(段落境界の情報を壊さないため)。
Public Function NormalizeForHash(ByVal s As String) As String
    Dim t As String: t = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
    Dim n As Long: n = Len(t)
    If n = 0 Then
        NormalizeForHash = ""
        Exit Function
    End If

    Dim outArr() As String: ReDim outArr(1 To n)
    Dim outCount As Long: outCount = 0
    Dim prevWasSpace As Boolean: prevWasSpace = False
    Dim i As Long
    For i = 1 To n
        Dim c As String: c = Mid$(t, i, 1)
        If c = " " Or c = vbTab Then
            If Not prevWasSpace Then
                outCount = outCount + 1
                outArr(outCount) = " "
            End If
            prevWasSpace = True
        Else
            outCount = outCount + 1
            outArr(outCount) = c
            prevWasSpace = False
        End If
    Next i

    Dim joined As String
    If outCount > 0 Then
        ReDim Preserve outArr(1 To outCount)
        joined = Join(outArr, "")
    End If
    NormalizeForHash = Trim$(joined)
End Function

Public Function VectorToCsv(vec() As Double) As String
    Dim n As Long: n = ArrLenD(vec)
    If n <= 0 Then Exit Function
    Dim parts() As String: ReDim parts(0 To n - 1)
    Dim lo As Long: lo = LBound(vec)
    Dim i As Long
    For i = 0 To n - 1
        parts(i) = Trim$(Str$(vec(lo + i)))   ' Str$は常に"."区切り(ロケール非依存)
    Next i
    VectorToCsv = Join(parts, ",")
End Function

' ----------------------------------------------------------------------------
' VectorToCsvPrec - 各成分を小数点以下decimals桁へ丸めてカンマ結合(設計書§F-1)。
'   ロケール非依存(Str$/Val方針)。丸めは四捨五入(負値対応)。空配列は""。
'   6桁丸めでcosine類似度への影響は1e-5未満(サイズ約-42%)。
' ----------------------------------------------------------------------------
Public Function VectorToCsvPrec(vec() As Double, ByVal decimals As Long) As String
    Dim n As Long: n = ArrLenD(vec)
    If n <= 0 Then Exit Function

    Dim d As Long: d = decimals
    If d < 0 Then d = 0
    If d > 12 Then d = 12
    Dim f As Double: f = 10 ^ d

    Dim parts() As String: ReDim parts(0 To n - 1)
    Dim lo As Long: lo = LBound(vec)
    Dim i As Long
    For i = 0 To n - 1
        Dim v As Double: v = vec(lo + i)
        Dim r As Double
        r = Fix(v * f + 0.5 * Sgn(v)) / f
        parts(i) = Trim$(Str$(r))
    Next i
    VectorToCsvPrec = Join(parts, ",")
End Function

' ----------------------------------------------------------------------------
' TruncateAndRenorm - ベクトルを先頭dims成分へ切詰め+L2再正規化(設計書§F-1、
'   Matryoshka特性を利用した次元圧縮)。dims以下ならL2正規化のみ。
'   未初期化/空/ゼロベクトルはFalseで無変更。
' ----------------------------------------------------------------------------
Public Function TruncateAndRenorm(ByRef vec() As Double, ByVal dims As Long) As Boolean
    Dim n As Long: n = ArrLenD(vec)
    If n <= 0 Or dims < 1 Then Exit Function

    Dim lo As Long: lo = LBound(vec)
    If n > dims Then
        Dim cut() As Double: ReDim cut(0 To dims - 1)
        Dim i As Long
        For i = 0 To dims - 1
            cut(i) = vec(lo + i)
        Next i
        vec = cut
    End If

    TruncateAndRenorm = L2Normalize(vec)
End Function

Public Function CsvToVector(ByVal s As String, ByRef vec() As Double) As Boolean
    Dim t As String: t = Trim$(s)
    If LenB(t) = 0 Then Exit Function
    Dim parts() As String: parts = Split(t, ",")
    Dim n As Long: n = UBound(parts) - LBound(parts) + 1
    If n < 1 Then Exit Function

    Dim tmp() As Double: ReDim tmp(0 To n - 1)
    Dim i As Long, j As Long: j = 0
    For i = LBound(parts) To UBound(parts)
        Dim p As String: p = Trim$(parts(i))
        If LenB(p) = 0 Then Exit Function   ' 空要素(連続カンマ等)=不正なCSV
        ' 2026-07-28(レビュー L-5): Val() は解釈できない文字列を黙って 0 に
        ' するため、壊れたベクトルCSVが「全要素0の正常なベクトル」として
        ' 通っていた。ゼロベクトルは全ての質問と無関係になるので、
        ' そのチャンクは検索から静かに消える(誰も気付けない)。
        ' 数値として読める形かを先に確かめ、読めなければ不正なCSVとして落とす。
        If Not LooksNumeric(p) Then Exit Function
        tmp(j) = Val(p)                      ' Valは"."小数点固定でロケール非依存
        j = j + 1
    Next i
    vec = tmp
    CsvToVector = True
End Function

' 次元不一致・空配列の場合は0を返す(呼び出し側が別途次元検査を行う規約。
' MASTER_SPEC §7.1 modUtil注記のとおり単純化)。
Public Function DotProduct(a() As Double, b() As Double) As Double
    Dim na As Long: na = ArrLenD(a)
    Dim nb As Long: nb = ArrLenD(b)
    If na <> nb Or na = 0 Then Exit Function

    Dim oa As Long: oa = LBound(a)
    Dim ob As Long: ob = LBound(b)
    Dim s As Double
    Dim i As Long
    For i = 0 To na - 1
        s = s + a(oa + i) * b(ob + i)
    Next i
    DotProduct = s
End Function

' ゼロベクトル(全要素0または空配列)はFalseを返し、vecは変更しない。
Public Function L2Normalize(ByRef vec() As Double) As Boolean
    Dim n As Long: n = ArrLenD(vec)
    If n = 0 Then Exit Function

    Dim lo As Long: lo = LBound(vec)
    Dim sumSq As Double
    Dim i As Long
    For i = lo To UBound(vec)
        sumSq = sumSq + vec(i) * vec(i)
    Next i
    If sumSq <= 0 Then Exit Function

    Dim norm As Double: norm = Sqr(sumSq)
    For i = lo To UBound(vec)
        vec(i) = vec(i) / norm
    Next i
    L2Normalize = True
End Function

' 未初期化/空配列判定。ArrLenD内部でOn Errorを使い、未割り当て配列の
' LBound/UBound例外(実行時エラー9)を安全に0件として扱う。
Public Function HasVector(vec() As Double) As Boolean
    HasVector = (ArrLenD(vec) > 0)
End Function

' 空配列(0件)は Split(vbNullString) で作る。
' 【2026-07-16 実機バグの恒久修正】以前は「ReDim outArr(0 To -1)」で0要素
' 配列を作っていたが、これはLibreOffice Basic特有に通る書き方で、実機の
' Excel VBAでは ReDim の上限<下限は実行時エラー9「インデックスが有効範囲に
' ありません」になる(VB.NETの空配列宣言との混同に由来する誤解だった)。
' Split(vbNullString) は言語仕様上どちらの環境でも LBound=0/UBound=-1 の
' 正当な空配列を返すため、For i = LBound To UBound が0回ループで安全に成立する。
Public Function SplitKeepNonEmpty(ByVal s As String, ByVal sep As String) As String()
    Dim raw() As String: raw = Split(s, sep)
    Dim maxN As Long: maxN = UBound(raw) - LBound(raw) + 1

    Dim outArr() As String
    If maxN <= 0 Then
        SplitKeepNonEmpty = Split(vbNullString)
        Exit Function
    End If

    ReDim outArr(0 To maxN - 1)
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = LBound(raw) To UBound(raw)
        If LenB(Trim$(raw(i))) > 0 Then
            outArr(cnt) = raw(i)
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then
        SplitKeepNonEmpty = Split(vbNullString)
    Else
        ReDim Preserve outArr(0 To cnt - 1)
        SplitKeepNonEmpty = outArr
    End If
End Function

Public Function HumanBytes(ByVal n As Double) As String
    Dim v As Double: v = n
    If v < 0 Then v = 0
    Dim units() As String: units = Split("B,KB,MB,GB,TB", ",")
    Dim idx As Long: idx = 0
    Do While v >= 1024# And idx < UBound(units)
        v = v / 1024#
        idx = idx + 1
    Loop
    If idx = 0 Then
        HumanBytes = CStr(CLng(v)) & " " & units(idx)
    Else
        HumanBytes = Format$(v, "0.0") & " " & units(idx)
    End If
End Function

Public Function HumanSeconds(ByVal sec As Double) As String
    Dim s As Long: s = CLng(sec)
    If s < 0 Then s = 0
    If s < 60 Then
        HumanSeconds = "約" & s & "秒"
    Else
        Dim m As Long: m = s \ 60
        Dim r As Long: r = s Mod 60
        If r = 0 Then
            HumanSeconds = "約" & m & "分"
        Else
            HumanSeconds = "約" & m & "分" & r & "秒"
        End If
    End If
End Function

Public Function SafeLeft(ByVal s As String, ByVal n As Long) As String
    Dim lim As Long: lim = n
    If lim < 0 Then lim = 0
    If Len(s) <= lim Then
        SafeLeft = s
    Else
        SafeLeft = Left$(s, lim)
    End If
End Function

' "yyyy-mm-dd hh:nn:ss"。Now()はExcel/LO両対応のVBAランタイム関数。
Public Function NowStamp() As String
    NowStamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Public Function FileNameOf(ByVal path As String) As String
    Dim i1 As Long: i1 = InStrRev(path, "\")
    Dim i2 As Long: i2 = InStrRev(path, "/")
    Dim i As Long: i = i1
    If i2 > i Then i = i2
    If i = 0 Then
        FileNameOf = path
    Else
        FileNameOf = Mid$(path, i + 1)
    End If
End Function

Public Function ExtOf(ByVal path As String) As String
    Dim fn As String: fn = FileNameOf(path)
    Dim i As Long: i = InStrRev(fn, ".")
    If i = 0 Or i = Len(fn) Then
        ExtOf = ""
    Else
        ExtOf = LCase$(Mid$(fn, i + 1))
    End If
End Function

' OneDrive/FATファイルシステムの秒精度誤差を吸収するため2秒丸めで比較する。
Public Function IsSameTimestamp(ByVal a As Date, ByVal b As Date) As Boolean
    Dim diffSec As Double: diffSec = Abs(CDbl(a - b)) * 86400#
    IsSameTimestamp = (diffSec <= 2#)
End Function

' ============================================================================
' JoinPagedText / SplitPagedText - ページ付きテキストの符号化・復号(R6)
' ----------------------------------------------------------------------------
'   画像PDFのOCRは「1ページ=1回のAI呼び出し」で複数ページ分の結果を作るが、
'   opt層とコア層の境界(modFeatures.InvokeFeature)は文字列1本しか運べない。
'   そこで行単位のマーカーでページ境界を埋め込み、受け取った側が
'   ExtractedPage() へ戻す。改行はLFへ統一する(CRLF/CR混在の資料でも
'   往復で形が変わらないようにするため)。
'
'   例(2ページ・上限で打ち切った場合):
'     @@NEXUS_TRUNCATED@@      ← 打ち切りがあったときだけ先頭行に付く
'     @@NEXUS_PAGE:1@@
'     1ページ目の本文
'     @@NEXUS_PAGE:2@@
'     2ページ目の本文
'
'   マーカーは「行全体が完全一致」した行だけを境界として扱う。本文中に
'   マーカーそっくりの行があっても壊れず、最悪でもそこでページが割れるだけで
'   本文の文字は1文字も失われない。
' ============================================================================
Public Function JoinPagedText(pages() As ExtractedPage, Optional ByVal truncated As Boolean = False) As String
    Dim head As String
    If truncated Then head = PAGED_TRUNC_MARK

    Dim n As Long: n = ArrLenP(pages)
    If n <= 0 Then
        JoinPagedText = head
        Exit Function
    End If

    Dim parts() As String: ReDim parts(0 To n - 1)
    Dim lo As Long: lo = LBound(pages)
    Dim i As Long
    For i = 0 To n - 1
        parts(i) = PAGED_MARK_PRE & CStr(pages(lo + i).page) & PAGED_MARK_SUF & vbLf & _
                   NormalizeEol(pages(lo + i).Text)
    Next i

    JoinPagedText = Join(parts, vbLf)
    If LenB(head) > 0 Then JoinPagedText = head & vbLf & JoinPagedText
End Function

' 復号。1ページも見つからなければFalse(pagesは触らない=呼び出し側は
' 「ページ付きではない普通の全文」として扱えばよい)。
Public Function SplitPagedText(ByVal s As String, ByRef pages() As ExtractedPage, ByRef truncated As Boolean) As Boolean
    truncated = False

    Dim body As String: body = NormalizeEol(s)
    If LenB(body) = 0 Then Exit Function

    Dim rows() As String: rows = Split(body, vbLf)
    Dim first As Long: first = LBound(rows)
    If rows(first) = PAGED_TRUNC_MARK Then
        truncated = True
        first = first + 1
    End If

    ' 1回目の走査でページ数を数える(ReDim Preserveの繰り返しを避ける)。
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = first To UBound(rows)
        If PagedMarkNumber(rows(i)) > 0 Then cnt = cnt + 1
    Next i
    If cnt = 0 Then
        truncated = False
        Exit Function
    End If

    Dim tmp() As ExtractedPage: ReDim tmp(0 To cnt - 1)
    Dim buf() As String: ReDim buf(0 To UBound(rows) - first)
    Dim bufN As Long: bufN = 0
    Dim idx As Long: idx = -1

    For i = first To UBound(rows)
        Dim pno As Long: pno = PagedMarkNumber(rows(i))
        If pno > 0 Then
            If idx >= 0 Then tmp(idx).Text = JoinFirstN(buf, bufN)
            idx = idx + 1
            bufN = 0
            tmp(idx).page = pno
        ElseIf idx >= 0 Then
            buf(bufN) = rows(i)
            bufN = bufN + 1
        End If
    Next i
    If idx >= 0 Then tmp(idx).Text = JoinFirstN(buf, bufN)

    pages = tmp
    SplitPagedText = True
End Function

' 行がページマーカーなら実ページ番号(1以上)、そうでなければ0を返す。
Private Function PagedMarkNumber(ByVal rowText As String) As Long
    Dim preLen As Long: preLen = Len(PAGED_MARK_PRE)
    Dim sufLen As Long: sufLen = Len(PAGED_MARK_SUF)
    If Len(rowText) <= preLen + sufLen - 1 Then Exit Function
    If Left$(rowText, preLen) <> PAGED_MARK_PRE Then Exit Function
    If Right$(rowText, sufLen) <> PAGED_MARK_SUF Then Exit Function

    Dim digits As String
    digits = Mid$(rowText, preLen + 1, Len(rowText) - preLen - sufLen)
    If LenB(digits) = 0 Or Len(digits) > 9 Then Exit Function   ' 9桁超はCLng溢れ回避

    Dim i As Long, ch As String
    For i = 1 To Len(digits)
        ch = Mid$(digits, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    PagedMarkNumber = CLng(digits)
End Function

Private Function JoinFirstN(arr() As String, ByVal n As Long) As String
    If n <= 0 Then Exit Function
    Dim keep() As String: ReDim keep(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        keep(i) = arr(i)
    Next i
    JoinFirstN = Join(keep, vbLf)
End Function

Private Function NormalizeEol(ByVal s As String) As String
    NormalizeEol = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
End Function

' 未割り当て/空のExtractedPage配列でも例外にせず0を返す(ArrLenDと同じ趣旨)。
Private Function ArrLenP(arr() As ExtractedPage) As Long
    On Error GoTo Empty0
    ArrLenP = UBound(arr) - LBound(arr) + 1
    Exit Function
Empty0:
    ArrLenP = 0
End Function

' ============================================================================
' DeobfuscateSecret - configシートに平文で置かないための軽い難読化の解除。
' ----------------------------------------------------------------------------
'   build/build_mybookshelf.py の obfuscate_secret() と対になる実装
'   (XOR + 16進エンコード)。これは暗号的な秘匿ではなく、configシートを
'   セルとして開いた人の目に平文キーが直接触れないようにする程度の対策
'   (VBAプロジェクト自体に触れる人には無意味。VBAプロジェクトへの
'   アクセス自体を前提とするこのアプリの配布モデル上、それ以上の防御は
'   このモジュール単体では不可能)。
'   "OBF1:"接頭辞が無い値(手動でconfigに平文キーを入力した場合等)は
'   そのまま返す(後方互換)。
' ============================================================================
Public Function DeobfuscateSecret(ByVal raw As String) As String
    Dim s As String: s = Trim$(raw)
    If Left$(s, Len(OBF_PREFIX)) <> OBF_PREFIX Then
        DeobfuscateSecret = s
        Exit Function
    End If
    DeobfuscateSecret = XorWithObfKey(HexDecodeBytes(Mid$(s, Len(OBF_PREFIX) + 1)))
End Function

Private Function HexDecodeBytes(ByVal hexStr As String) As String
    Dim n As Long: n = Len(hexStr) \ 2
    Dim outStr As String: outStr = ""
    Dim i As Long
    On Error GoTo Fail
    For i = 1 To n
        outStr = outStr & Chr$(CLng("&H" & Mid$(hexStr, (i - 1) * 2 + 1, 2)))
    Next i
    HexDecodeBytes = outStr
    Exit Function
Fail:
    HexDecodeBytes = ""   ' 壊れた16進文字列は空扱い(呼び出し側がE0203として処理)
End Function

Private Function XorWithObfKey(ByVal s As String) As String
    Dim outStr As String: outStr = ""
    Dim kLen As Long: kLen = Len(OBF_KEY)
    Dim i As Long
    For i = 1 To Len(s)
        outStr = outStr & Chr$(Asc(Mid$(s, i, 1)) Xor Asc(Mid$(OBF_KEY, ((i - 1) Mod kLen) + 1, 1)))
    Next i
    XorWithObfKey = outStr
End Function

' ============================================================================
' 内部ヘルパー(Fnv1a64Hex専用): 64bit整数演算のDouble安全実装
' ============================================================================
' (FNV_OFFSET_HI/LO・FNV_PRIME_HI/LOの定義はLO互換のためファイル冒頭に移動済み。
'  値の由来: offset_basis = 0xCBF29CE484222325 を上位/下位32bitワードに分解。
'  prime = 0x100000001B3 = 2^40 + 435 なので上位ワードは非常に小さい値になる。
'  16進リテラルはビットパターンとしてそのままLongへ格納される(符号は無視)。)

' 未割り当て配列でも例外にせず0を返す(LBound/UBoundは未割り当て配列に
' 対して実行時エラー9を出すため、1行スコープのOn Errorで吸収する)。
Private Function ArrLenD(arr() As Double) As Long
    On Error GoTo Empty0
    ArrLenD = UBound(arr) - LBound(arr) + 1
    Exit Function
Empty0:
    ArrLenD = 0
End Function

' Longのビットパターン(0～2^32-1相当、符号は無視)をDoubleの数値に変換する。
Private Function U32ToDouble(ByVal L As Long) As Double
    If L < 0 Then
        U32ToDouble = CDbl(L) + 4294967296#     ' 2^32
    Else
        U32ToDouble = CDbl(L)
    End If
End Function

' [0, 2^32) のDouble値を、その値を表すLongのビットパターンに変換する。
Private Function DoubleToU32Bits(ByVal d As Double) As Long
    Dim v As Double: v = d - Int(d / 4294967296#) * 4294967296#   ' Mod 2^32相当
    If v >= 2147483648# Then                     ' 2^31
        DoubleToU32Bits = CLng(v - 4294967296#)
    Else
        DoubleToU32Bits = CLng(v)
    End If
End Function

' 32bit×32bit → 64bit(hiOut:loOut)の符号なし乗算。オーバーフローなし。
' ---------------------------------------------------------------------------
' a, b をそれぞれ上位16bit/下位16bit(aHi,aLo / bHi,bLo)に分割し、
'   t0=aLo*bLo, t1=aHi*bLo, t2=aLo*bHi, t3=aHi*bHi
' を計算する(各項は最大 65535*65535 ≒ 4.29e9 なのでDoubleで厳密に表現できる。
' Doubleが誤差なく表せる整数の上限は 2^53 ≒ 9.007e15 であり、この計算に
' 現れる最大の中間値でもそれを大きく下回る)。
' これらを16bit単位の桁(digit0～digit3)としてキャリー(繰り上がり)を
' 手動で伝播しながら合算し、
'   lo32 = digit1*65536 + digit0
'   hi32 = digit3*65536 + digit2
' とする(digit3から溢れた分は64bitを超えるため切り捨てる=Mod 2^64相当)。
Private Sub MulU32(ByVal aBits As Long, ByVal bBits As Long, ByRef hiOut As Long, ByRef loOut As Long)
    Dim a As Double: a = U32ToDouble(aBits)
    Dim b As Double: b = U32ToDouble(bBits)

    Dim aHi As Double: aHi = Int(a / 65536#)
    Dim aLo As Double: aLo = a - aHi * 65536#
    Dim bHi As Double: bHi = Int(b / 65536#)
    Dim bLo As Double: bLo = b - bHi * 65536#

    Dim t0 As Double: t0 = aLo * bLo
    Dim t1 As Double: t1 = aHi * bLo
    Dim t2 As Double: t2 = aLo * bHi
    Dim t3 As Double: t3 = aHi * bHi

    Dim digit0 As Double: digit0 = t0 - Int(t0 / 65536#) * 65536#
    Dim carry As Double: carry = Int(t0 / 65536#)

    Dim sum1 As Double
    sum1 = carry + (t1 - Int(t1 / 65536#) * 65536#) + (t2 - Int(t2 / 65536#) * 65536#)
    Dim digit1 As Double: digit1 = sum1 - Int(sum1 / 65536#) * 65536#
    carry = Int(sum1 / 65536#) + Int(t1 / 65536#) + Int(t2 / 65536#)

    Dim sum2 As Double
    sum2 = carry + (t3 - Int(t3 / 65536#) * 65536#)
    Dim digit2 As Double: digit2 = sum2 - Int(sum2 / 65536#) * 65536#
    carry = Int(sum2 / 65536#) + Int(t3 / 65536#)

    Dim digit3 As Double: digit3 = carry - Int(carry / 65536#) * 65536#   ' 64bit超過分は破棄

    loOut = DoubleToU32Bits(digit1 * 65536# + digit0)
    hiOut = DoubleToU32Bits(digit3 * 65536# + digit2)
End Sub

' hash(hHi:hLo) = hash * FNV_PRIME Mod 2^64 を破壊的に計算する。
' ---------------------------------------------------------------------------
' prime = PHi*2^32 + PLo (PHi=256, PLo=435という小さい定数)であることを
' 利用すると、64bit×64bitの積 mod 2^64 は次式で求まる
' (HHi*PHi*2^64 の項は mod 2^64 で必ず0になるため計算不要):
'   newLo = 下位32bit( HLo * PLo )
'   newHi = ( 上位32bit(HLo*PLo) + 下位32bit(HHi*PLo) + 下位32bit(HLo*PHi) ) Mod 2^32
Private Sub MulU64ByPrime(ByRef hHi As Long, ByRef hLo As Long)
    Dim hiA As Long, loA As Long
    MulU32 hLo, FNV_PRIME_LO, hiA, loA          ' HLo * PLo -> 64bit(hiA:loA)

    Dim hiB As Long, loB As Long
    MulU32 hHi, FNV_PRIME_LO, hiB, loB          ' HHi * PLo -> 下位32bitのみ使用

    Dim hiC As Long, loC As Long
    MulU32 hLo, FNV_PRIME_HI, hiC, loC          ' HLo * PHi -> 下位32bitのみ使用

    Dim sumHi As Double
    sumHi = U32ToDouble(hiA) + U32ToDouble(loB) + U32ToDouble(loC)

    hLo = loA
    hHi = DoubleToU32Bits(sumHi)
End Sub

' Longのビットパターンを8桁小文字16進文字列(ゼロ埋め)にする。
' VBAのHex$()はビットパターンをそのまま16進化するため符号を気にせず使える。
Private Function U32ToHex8(ByVal L As Long) As String
    U32ToHex8 = LCase$(Right$("00000000" & Hex$(L), 8))
End Function

' ベクトルCSVの1要素が数値として読める形か(純関数)。
' 許すのは 先頭の符号 / 数字 / 小数点1つ / 指数表記(e|E と符号)。
' ロケール差を避けるため IsNumeric は使わない(全角数字や通貨記号を通す)。
Private Function LooksNumeric(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    Dim seenDigit As Boolean, seenDot As Boolean, seenExp As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            seenDigit = True
        ElseIf ch = "." Then
            If seenDot Or seenExp Then Exit Function
            seenDot = True
        ElseIf ch = "e" Or ch = "E" Then
            If seenExp Or Not seenDigit Then Exit Function
            seenExp = True
            seenDigit = False        ' 指数部にも数字が要る
        ElseIf ch = "+" Or ch = "-" Then
            ' 符号は先頭か、指数の直後だけ許す
            If i > 1 Then
                Dim prev As String: prev = Mid$(s, i - 1, 1)
                If prev <> "e" And prev <> "E" Then Exit Function
            End If
        Else
            Exit Function
        End If
    Next i
    LooksNumeric = seenDigit
End Function
