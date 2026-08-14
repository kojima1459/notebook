Attribute VB_Name = "modBackdrop"
Option Explicit

' ============================================================================
' modBackdrop - 画面の「地(背景)」だけを受け持つモジュール(2026-08-14 R32波4)
' ----------------------------------------------------------------------------
' なぜ新設したのか(容量と主題の両方):
'   R32の実機第17報②で、余白まわりの受け皿になり得るモジュールが軒並み
'   上限30,000字に張り付いていた(実測: modSkin残27 / modHub残159 /
'   modUIShelf残217 / modHubStat残450 / modChrome残837 / modViewport残1,038)。
'   憲章§4-6「入らなければ実体を余裕モジュールへ置いて1行呼び出し」の
'   受け皿がもう無い状態で、地色の話だけを凝集させた新しい置き場所を作る。
'
' 持ち分は「セルの文字・Shapeではなく、その【背後】に何を敷くか」だけ:
'   ・RestoreShelfHeaderBg  … 一覧表の一律塗りで消える見出し行のグレーを戻す
'                             (W4-2。塗る順序を変えて解くための復元点)
'   ・LogStyleFailOnce      … 標準スタイル方式の失敗を1セッション1回ログへ
'                             (W4-4。無言失敗の根絶)
'   ・Apply / BmpHex / MemoKey / MemoPut
'                           … 背景画像方式(W4-5)。テーマの地色で1×1のBMPを
'                             一時フォルダへ生成し SetBackgroundPicture で
'                             敷く。セルを1つも使わないので使用済み範囲
'                             (=ホイールの停止線の素)が増えない。
'
' 【この画面の余白問題についての確定事実】(R32実機プローブ。ここに集約する)
'   1. ScrollArea はホイールを止めない。止まるのはセル選択とスクロールバー
'      だけで、"A1:L15" にしても60行付近まで転がる(R32実機実証)。
'      ホイールの停止線は S = B + k(B=使用済み末尾行 / k=1画面の行数)。
'   2. したがって「塗る深さ d を増やす」対策は原理的に効かない。塗った行は
'      使用済みになるので B が増え、S も同じだけ下がる。S − B = k は d に
'      依存しない ―― R18〜R27の7ラウンドが同じ壁に当たり続けた真因。
'   3. ThisWorkbook.Styles("Normal")/("標準") はどちらも実機で 1004。
'      無保護シートを Activate してからでも同じ。標準スタイル方式は
'      実機では使用不能(modChrome.ApplyNormalStyleBg はR28以来一度も
'      機能していなかった。R32 W4-4で失敗をログに残すよう是正)。
' ============================================================================

' ---- モジュールレベル宣言(実機VBAでは宣言部が必ず全プロシージャより先) ----
' 標準スタイル方式の失敗を1セッション1回だけログに残したか(W4-4)。
Private mStyleFailLogged As Boolean
' 背景画像の失敗を1セッション1回だけログに残したか(W4-5)。
Private mBackdropFailLogged As Boolean
' どのシートへどの色の背景画像を敷いたか。"|シート名=色|…" の1本の文字列
' (MemoKey/MemoPut がこの形の唯一の持ち主)。W4-5。
Private mApplied As String

' ----------------------------------------------------------------------------
' RestoreShelfHeaderBg - 一覧表のカード見出し行だけ、明示塗りを戻す。
' ----------------------------------------------------------------------------
'   R32 W4-2【確定バグ・実機第17報②「本棚に古い地色が残る」】:
'   マイ本棚の一覧表モードだけが地色を塗り直していなかった。
'   modUIShelf.RenderShelf には ws.Cells.Clear が無く(あるのは EnsureLayout
'   側だけ)、後始末の ApplyShelfExtent が modKnowledge.ApplyShelfBound を
'   paintBg:=False で呼んでいたため、直前にギャラリー/解決事例モードが
'   【別テーマの地色】で塗ったセルがそのまま残る ―― これがテーマを戻しても
'   一覧表だけ古い地色に見える正体。
'
'   paintBg:=False にしていた元の理由(R20-1c のコメント)は
'   「描き終えた後の一律塗りが、見出し行のグレーとカードの塗り分けを潰す」。
'   そこで塗る【順序】を変えて解く:
'     ・塗り直しはカードを描く【前】の呼び出し(ApplyShelfExtent の
'       withFont=True 側)だけで行う。カードはその後に描かれるので潰れない。
'     ・その一律塗りが消してしまう唯一の明示塗り=カード見出し行のグレーを、
'       ここで即座に戻す。
'   カード行そのものは Interior を持たない(modUIShelf 内の明示塗りは
'   見出し行の1箇所だけ。実測で確認)ので、「塗り分け」が潰れることはない。
'   描き終えた後の呼び出し(withFont=False)は従来どおり塗らない。
'
'   15921906 = RGB(242,242,242)。値の持ち主は modUIShelf.EnsureLayout で、
'   ここはその復元だけを行う(同じ値を2箇所で「決めない」)。
'   10列(A:J)は modUIShelf が一覧表で使う全列
'   (modChrome.ApplyShelfTableTextColor と同じ範囲)。
'   見出し行のグレーはテーマに依らず維持する既存の裁定(R28H F1)をそのまま
'   踏襲する ―― 文字色側(ApplyShelfTableTextColor)がこの行を除外している
'   ため、ここだけテーマ地色に変えると dark で「明るい地に明るい文字」になる。
' ----------------------------------------------------------------------------
Public Sub RestoreShelfHeaderBg(ByVal ws As Worksheet, ByVal headerRow As Long)
    If ws Is Nothing Then Exit Sub
    If headerRow < 1 Then Exit Sub
    On Error Resume Next
    ws.Cells(headerRow, 1).Resize(1, 10).Interior.Color = 15921906
    Err.Clear
    On Error GoTo 0
End Sub

' ============================================================================
' W4-4: 無言失敗の根絶
' ============================================================================

' ----------------------------------------------------------------------------
' LogStyleFailOnce - modChrome.ApplyNormalStyleBg の失敗を1行残す。
' ----------------------------------------------------------------------------
'   R32 W4-4【是正・8ラウンド気づけなかった構造的原因】:
'   modChrome.ApplyNormalStyleBg(R28波1)は
'   `ws.Parent.Styles("Normal").Interior.Color = 地色` で画面の地を一括で
'   塗るはずだったが、R32の実機プローブで
'     ・Styles("Normal") → 実行時エラー1004
'     ・Styles("標準")   → 実行時エラー1004
'     ・無保護のホームシートを Activate してから叩いても両方1004
'   が確定した。つまりこの関数はR28以降【一度も機能していない】。
'   にもかかわらず `On Error Resume Next` で握り潰され、usage_log にも
'   err_log にも1行も出ていなかったため、余白問題の調査は「地はNormal
'   スタイルで塗れている」という誤った前提の上を8ラウンド走り続けた。
'
'   よって関数自体は残す(将来のExcel/環境で使える可能性があり、削除すると
'   「試したが駄目だった」という事実まで消える)が、【黙って失敗しない】。
'   errNum<>0 のときだけ、1セッションに1回 usage_log へ残す:
'     event=normal_style_bg_failed / mode=シート名 / detail=err番号と説明
'   1回に絞る理由は、この関数が画面を描くたびに4経路から呼ばれるため
'   (毎描画で書くと usage_log が実用にならないほど膨らむ)。
'   ログ自体が失敗しても画面は落とさない(全体を On Error Resume Next 配下)。
' ----------------------------------------------------------------------------
Public Sub LogStyleFailOnce(ByVal ws As Worksheet, ByVal errNum As Long, _
                            ByVal errDesc As String)
    If errNum = 0 Then Exit Sub
    If mStyleFailLogged Then Exit Sub
    mStyleFailLogged = True          ' 先に立てる(ログ側で落ちても再入しない)
    On Error Resume Next
    Dim nm As String
    If Not ws Is Nothing Then nm = ws.Name
    modLog.LogUsage "normal_style_bg_failed", nm, _
        "Styles(Normal)への地色代入が失敗 err=" & errNum & " " & errDesc
    Err.Clear
    On Error GoTo 0
End Sub

' ============================================================================
' W4-5: 背景画像方式(SetBackgroundPicture)
' ============================================================================
' なぜこれなのか(採用理由と、これ以外に道が無い理由):
'   冒頭の確定事実1・2のとおり、ホイールの停止線 S = B + k は「使用済み末尾
'   B」で決まり、S − B = k は塗る深さに依存しない。セルを塗って地色を届かせる
'   方式は、塗ったぶんだけ B が増えて S も下がるので【原理的に】届かない。
'   届かせる手は「セルを1つも使わずに地を色づける」しか残っておらず、
'   標準スタイル方式(確定事実3)が実機で使えない以上、
'   Worksheet.SetBackgroundPicture が最後の1本になる。
'   背景画像はシート全体の背後にタイル表示されるだけで、使用済み範囲を
'   1行も増やさない ―― B が増えないので、転がる距離も伸びない。
'
' 実装方針:
'   テーマの bg 色で 1×1 ピクセルの 24bit BMP を一時フォルダへ動的生成し、
'   SetBackgroundPicture で敷く。1×1 がタイルされるので、どんな窓サイズでも
'   全面がその1色になる。PNGと違い BMP は圧縮が無く、ヘッダ+画素を素直に
'   並べるだけで作れる(=VBAから外部ライブラリ無しで生成できる)。
'
' ★ BMPバイト列の検算(BmpHex の中身。ここが唯一の技術リスクなので明記する)
'   全長58バイト = ファイルヘッダ14 + 情報ヘッダ40 + 画素4。
'   [ファイルヘッダ BITMAPFILEHEADER 14バイト]
'     +0  'B''M'            = 42 4D
'     +2  bfSize      = 58  = 3A 00 00 00   (リトルエンディアン)
'     +6  bfReserved1 = 0   = 00 00
'     +8  bfReserved2 = 0   = 00 00
'     +10 bfOffBits   = 54  = 36 00 00 00   (14+40=54。画素データの開始位置)
'   [情報ヘッダ BITMAPINFOHEADER 40バイト]
'     +14 biSize          = 40 = 28 00 00 00
'     +18 biWidth         = 1  = 01 00 00 00
'     +22 biHeight        = 1  = 01 00 00 00 (正=ボトムアップ。1行なので同じ)
'     +26 biPlanes        = 1  = 01 00       (2バイト)
'     +28 biBitCount      = 24 = 18 00       (2バイト。24bit=BGR各1バイト)
'     +30 biCompression   = 0  = 00 00 00 00 (BI_RGB=無圧縮)
'     +34 biSizeImage     = 4  = 04 00 00 00 (下の行サイズと一致させる)
'     +38 biXPelsPerMeter = 2835 = 13 0B 00 00 (72dpi。2835=&H0B13)
'     +42 biYPelsPerMeter = 2835 = 13 0B 00 00
'     +46 biClrUsed       = 0  = 00 00 00 00
'     +50 biClrImportant  = 0  = 00 00 00 00
'   [画素データ 4バイト]
'     行サイズ = ((biWidth * biBitCount + 31) \ 32) * 4
'              = ((1 * 24 + 31) \ 32) * 4 = (55 \ 32) * 4 = 1 * 4 = 4
'     → 画素3バイト(B, G, R の順。RGBではない)+ 4バイト境界へのパディング1。
'   検算例(darkテーマの地 RGB(15,23,42) = VBAのLong 2758415):
'     B=42=2A / G=23=17 / R=15=0F → 末尾4バイトは 2A 17 0F 00。
'     全体 = 424D3A0000000000000036000000 28000000 01000000 01000000 0100 1800
'            00000000 04000000 130B0000 130B0000 00000000 00000000 2A170F00
'     (modTestsPure32 がこの116字をゴールデンとして固定する)
'
' 適用範囲:
'   ホーム(Hub)/ Dashboard / マイ本棚 の3枚だけ。チャット(Nexus)は保護シート
'   のため SetBackgroundPicture が通るか実機で未検証 ―― 通らない場合に
'   モーダルや例外を出す可能性を排除できないので、この波では対象外にする
'   (対象に加えるときは TargetSheet に SH_NEXUS を足す1行だけで済む)。
'   入口は modChrome.ApplyNormalStyleBg の先頭1行。あそこは Setup*Columns
'   4本と modSkin.ApplyTheme から呼ばれる「このシートの地を決める」唯一の
'   合流点で、テーマ切替後の再描画も必ずそこを通る(W4-1で入口を
'   modHub.OnThemeToggle 1本へ一本化済み)。
' ============================================================================

' ----------------------------------------------------------------------------
' BmpHex - 1×1・24bit BMP ファイル全体(58バイト)を16進文字列で返す【純関数】。
' ----------------------------------------------------------------------------
'   bgColor: VBAのLong色(RGB(r,g,b) = r + g*256 + b*65536)
'   戻り値 : 116字(58バイト×2)の大文字16進。Excel/ファイルI/Oに一切触れない
'            ので LibreOffice の純ロジックテストからそのまま固定できる
'            (SetBackgroundPicture 自体はLOで検証できないため、
'             「敷く中身が正しいか」だけはここで機械的に守る)。
' ----------------------------------------------------------------------------
Public Function BmpHex(ByVal bgColor As Long) As String
    Dim r As Long, g As Long, b As Long
    r = bgColor And &HFF&
    g = (bgColor \ 256) And &HFF&
    b = (bgColor \ 65536) And &HFF&
    BmpHex = "424D" & Hex32(58) & "00000000" & Hex32(54) & _
             Hex32(40) & Hex32(1) & Hex32(1) & "0100" & "1800" & _
             Hex32(0) & Hex32(4) & Hex32(2835) & Hex32(2835) & _
             Hex32(0) & Hex32(0) & _
             Hex8(b) & Hex8(g) & Hex8(r) & "00"
End Function

' Hex8  - 1バイトを2字の16進へ(0〜255。範囲外は下位8bitだけ見る)。純関数。
Private Function Hex8(ByVal v As Long) As String
    Hex8 = Right$("0" & Hex$(v And &HFF&), 2)
End Function

' Hex32 - 4バイト値をリトルエンディアンの8字16進へ。純関数。
'   BMPの数値フィールドは全てリトルエンディアン(下位バイトが先)。
Private Function Hex32(ByVal v As Long) As String
    Hex32 = Hex8(v) & Hex8(v \ 256) & Hex8(v \ 65536) & Hex8(v \ 16777216)
End Function

' ----------------------------------------------------------------------------
' MemoKey / MemoPut - 「どのシートにどの色を敷いたか」のメモ【純関数】。
' ----------------------------------------------------------------------------
'   Apply はシート描画のたびに呼ばれる(Setup*Columns 経由)ので、同じ色なら
'   ファイル生成も SetBackgroundPicture も走らせない。メモは
'   "|シート名=色|シート名=色|" の1本の文字列で持つ(Dictionaryを増やさない)。
'   シート名に "|" は本アプリでは使わない(ホーム/Dashboard/マイ本棚の3枚)。
'   MemoPut は同じシートの古い1件を必ず落としてから足す(重複させない)。
' ----------------------------------------------------------------------------
Public Function MemoKey(ByVal sheetName As String, ByVal bgColor As Long) As String
    MemoKey = "|" & sheetName & "=" & CStr(bgColor) & "|"
End Function

Public Function MemoPut(ByVal memo As String, ByVal sheetName As String, _
                        ByVal bgColor As Long) As String
    Dim head As String: head = "|" & sheetName & "="
    Dim s As String: s = memo
    Dim p As Long: p = InStr(1, s, head, vbBinaryCompare)
    If p > 0 Then
        Dim q As Long: q = InStr(p + Len(head), s, "|")
        If q = 0 Then
            s = Left$(s, p - 1)
        Else
            s = Left$(s, p - 1) & Mid$(s, q)
        End If
    End If
    If LenB(s) = 0 Then s = "|"
    MemoPut = s & sheetName & "=" & CStr(bgColor) & "|"
End Function

' ----------------------------------------------------------------------------
' Apply - このシートの背後にテーマの地色を敷く(冪等)。
' ----------------------------------------------------------------------------
'   呼び口: modChrome.ApplyNormalStyleBg の先頭(4画面の Setup*Columns と
'   modSkin.ApplyTheme が必ず通る合流点)。描画のたびに呼ばれる前提なので、
'   同じシートへ同じ色を敷き直さない(mApplied のメモで弾く)。
'   失敗しても画面は落とさない・黙らない(W4-4と同じ思想。1セッション1回ログ)。
' ----------------------------------------------------------------------------
Public Sub Apply(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next

    Dim nm As String
    nm = ws.Name
    If Err.Number <> 0 Then GoTo CleanExit
    If Not TargetSheet(nm) Then GoTo CleanExit

    Dim c As Long
    c = modUI.UiColor("bg")
    If Err.Number <> 0 Then GoTo CleanExit
    ' 0 は modSkin.ResolveColor の「未知のキー」センチネル値でもある
    ' (modChrome.ApplyNormalStyleBg の R28H F8(m-3) と同じ理由で異常値扱い)。
    If c = 0 Then GoTo CleanExit
    If InStr(1, mApplied, MemoKey(nm, c), vbBinaryCompare) > 0 Then GoTo CleanExit

    Dim p As String
    p = BmpPath(c)
    If LenB(p) = 0 Then GoTo CleanExit
    Err.Clear
    If Not WriteBmpFile(p, c) Then
        LogBackdropFailOnce nm, "bmp_write", Err.Number, Err.Description
        GoTo CleanExit
    End If

    Err.Clear
    ws.SetBackgroundPicture p
    Dim bpNum As Long, bpDesc As String
    bpNum = Err.Number: bpDesc = Err.Description
    Err.Clear
    ' 成否に関わらずメモは進める。失敗する環境で描画のたびにファイルI/Oを
    ' 繰り返さないため(失敗の事実はログに1行残る)。
    mApplied = MemoPut(mApplied, nm, c)
    If bpNum <> 0 Then
        LogBackdropFailOnce nm, "set_background", bpNum, bpDesc
    Else
        SweepOldBmp p
    End If

CleanExit:
    Err.Clear
    On Error GoTo 0
End Sub

' 背景画像を敷く対象シートか。チャット(Nexus)は保護シートで未検証のため除外
' (実機で通ることが確認できたら modAppDef.SH_NEXUS を1行足すだけでよい)。
Private Function TargetSheet(ByVal sheetName As String) As Boolean
    TargetSheet = (sheetName = modAppDef.SH_HOME) Or _
                  (sheetName = modAppDef.SH_NEXUS_DASH) Or _
                  (sheetName = modAppDef.SH_SHELF)
End Function

' 一時ファイルのフルパス。色ごとに別名にする理由:
'   (a) 同名を上書きすると、Excelが同じパスの画像をキャッシュして色が
'       変わらない可能性を排除できない(実機で検証できないので安全側)。
'   (b) 直前のファイルが何らかの理由でロックされていても、新しい名前なら
'       書き込みが必ず成功する。
' 残骸は SweepOldBmp が掃除する(色は最大6テーマ=58バイト×6で実害は無い)。
Private Function BmpPath(ByVal bgColor As Long) As String
    Dim d As String
    d = Environ$("TEMP")
    If LenB(d) = 0 Then d = Environ$("TMP")
    If LenB(d) = 0 Then Exit Function
    If Right$(d, 1) <> "\" Then d = d & "\"
    BmpPath = d & BmpFileName(bgColor)
End Function

' ファイル名だけ(掃除の照合にも使うので分けてある)。純関数。
Private Function BmpFileName(ByVal bgColor As Long) As String
    BmpFileName = "MyBookshelf_bg_" & Right$("000000" & Hex$(bgColor), 6) & ".bmp"
End Function

' BmpHex の16進をバイト列へ戻してファイルへ書く。既に正しいサイズで在るなら
' 書き直さない(58バイト固定なのでサイズ一致で十分)。
Private Function WriteBmpFile(ByVal filePath As String, ByVal bgColor As Long) As Boolean
    On Error Resume Next
    Dim hx As String: hx = BmpHex(bgColor)
    Dim n As Long: n = Len(hx) \ 2

    If LenB(Dir$(filePath)) > 0 Then
        If FileLen(filePath) = n And Err.Number = 0 Then
            WriteBmpFile = True
            Err.Clear
            Exit Function
        End If
        Err.Clear
        Kill filePath
        Err.Clear
    End If
    Err.Clear

    Dim bytes() As Byte
    ReDim bytes(1 To n)
    Dim i As Long
    For i = 1 To n
        bytes(i) = CByte(CLng("&H" & Mid$(hx, i * 2 - 1, 2)))
    Next i
    If Err.Number <> 0 Then Exit Function

    Dim fn As Long: fn = FreeFile
    Open filePath For Binary Access Write As #fn
    If Err.Number <> 0 Then Exit Function
    Put #fn, 1, bytes
    Close #fn
    WriteBmpFile = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' 今使っているBMP以外の MyBookshelf_bg_*.bmp を消す(失敗は握って続行)。
' Dir$ の列挙中に Kill すると列挙状態が壊れるため、名前を集めてから消す。
Private Sub SweepOldBmp(ByVal keepPath As String)
    On Error Resume Next
    Dim sep As Long: sep = InStrRev(keepPath, "\")
    If sep < 1 Then Exit Sub
    Dim dir_ As String: dir_ = Left$(keepPath, sep)
    Dim keepName As String: keepName = Mid$(keepPath, sep + 1)

    Dim victims As String
    Dim f As String
    f = Dir$(dir_ & "MyBookshelf_bg_*.bmp")
    Do While LenB(f) > 0
        If StrComp(f, keepName, vbTextCompare) <> 0 Then victims = victims & f & vbTab
        f = Dir$()
    Loop
    Err.Clear

    Dim parts() As String
    parts = Split(victims, vbTab)
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If LenB(parts(i)) > 0 Then
            Kill dir_ & parts(i)
            Err.Clear
        End If
    Next i
    On Error GoTo 0
End Sub

' 背景画像の失敗を1セッション1回だけ usage_log へ(W4-4と同じ「無言禁止」)。
Private Sub LogBackdropFailOnce(ByVal sheetName As String, ByVal stage As String, _
                                ByVal errNum As Long, ByVal errDesc As String)
    If mBackdropFailLogged Then Exit Sub
    mBackdropFailLogged = True
    On Error Resume Next
    modLog.LogUsage "backdrop_failed", sheetName, _
        "stage=" & stage & " err=" & errNum & " " & errDesc
    Err.Clear
    On Error GoTo 0
End Sub
