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
'                           … 背景画像方式(W4-5)。テーマの地色で8×8のBMPを
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
' 失敗した回数(2026-08-14 R32 Fix波 F7 m-3)。従来は失敗してもメモを進めて
' いたため、一過性の失敗(一時フォルダが一瞬ロックされていた・別プロセスが
' 掴んでいた等)が【そのセッションの恒久失敗】になっていた。3回までは
' 敷き直しを試し、それでも駄目なら諦める(描画のたびにファイルI/Oを
' 繰り返さない、という元の意図はこの上限で守る)。
Private mFailCount As Long
Private Const MAX_APPLY_RETRY As Long = 3
' 上の mFailCount が今どの (シート,色) 組を数えているか(R33波5a W5-3)。
' MemoKey と同じ書式の文字列。組が変わったら mFailCount を0に戻す。
Private mFailKey As String

' ---- 条件付き書式方式(W5-1)の宣言。設計の根拠は下の「W5-1」節 ----
Private mCfOff As Boolean        ' 検算に落ちた=この端末では不成立。以後張らない
Private mCfOkLogged As Boolean   ' 「効いている」を1セッション1回だけログしたか
Private mCfFailLogged As Boolean ' 不首尾を1セッション1回だけログしたか
' ルールの目印。数式に埋めて、掃除のとき【自分が張った1本だけ】を見分ける。
Private Const CF_MARK As String = "MBSBG"
' XlFormatConditionType.xlExpression(定数名は実機Excelにしか無いので数値)。
Private Const CF_TYPE_EXPRESSION As Long = 2
Public Const CF_DEPTH_ROWS As Long = 400      ' 張る深さ(行)。根拠はW5-1節
Public Const CF_MAX_ROW As Long = 1048576     ' Excel 2007以降のシート最終行

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
'   テーマの bg 色で 8×8 ピクセルの 24bit BMP を一時フォルダへ動的生成し、
'   SetBackgroundPicture で敷く。全画素が同じ色のタイルなので、どんな窓
'   サイズでも全面がその1色になる。PNGと違い BMP は圧縮が無く、ヘッダ+画素を
'   素直に並べるだけで作れる(=VBAから外部ライブラリ無しで生成できる)。
'
' ★ BMPバイト列の検算(BmpHex の中身。ここが唯一の技術リスクなので明記する)
'   2026-08-14(R32 Fix波 F13・予防): 1×1 → 8×8 へ拡大した。1×1だとタイルの
'   敷き詰め回数が画素数ぶん必要になり、32bit Excelでの描画コストが未検証
'   (実機で「背景を敷いた瞬間に重い」が出たら原因の切り分けが難しい)。
'   8×8にすればタイル回数は 1/64 になり、ファイルは 58 → 246バイトで済む。
'
'   全長246バイト = ファイルヘッダ14 + 情報ヘッダ40 + 画素192。
'   [ファイルヘッダ BITMAPFILEHEADER 14バイト]
'     +0  'B''M'            = 42 4D
'     +2  bfSize      = 246 = F6 00 00 00   (リトルエンディアン。246=&HF6)
'     +6  bfReserved1 = 0   = 00 00
'     +8  bfReserved2 = 0   = 00 00
'     +10 bfOffBits   = 54  = 36 00 00 00   (14+40=54。画素データの開始位置)
'   [情報ヘッダ BITMAPINFOHEADER 40バイト]
'     +14 biSize          = 40 = 28 00 00 00
'     +18 biWidth         = 8  = 08 00 00 00
'     +22 biHeight        = 8  = 08 00 00 00 (正=ボトムアップ。全行同色なので
'                                             上下の向きは結果に影響しない)
'     +26 biPlanes        = 1  = 01 00       (2バイト)
'     +28 biBitCount      = 24 = 18 00       (2バイト。24bit=BGR各1バイト)
'     +30 biCompression   = 0  = 00 00 00 00 (BI_RGB=無圧縮)
'     +34 biSizeImage     = 192 = C0 00 00 00 (下の行サイズ×高さと一致させる)
'     +38 biXPelsPerMeter = 2835 = 13 0B 00 00 (72dpi。2835=&H0B13)
'     +42 biYPelsPerMeter = 2835 = 13 0B 00 00
'     +46 biClrUsed       = 0  = 00 00 00 00
'     +50 biClrImportant  = 0  = 00 00 00 00
'   [画素データ 192バイト]
'     行サイズ = ((biWidth * biBitCount + 31) \ 32) * 4
'              = ((8 * 24 + 31) \ 32) * 4 = (223 \ 32) * 4 = 6 * 4 = 24
'     8画素 × 3バイト = 24 ちょうどなので【パディングは0バイト】
'       (1×1のときは 3バイト+パディング1バイト だった。ここが変更点)
'     画素データ = 24バイトの行 × 8行 = 192バイト
'     bfSize = 54 + 192 = 246 / biSizeImage = 192 ―― 上の2値と一致すること。
'   検算例(darkテーマの地 RGB(15,23,42) = VBAのLong 2758415):
'     B=42=2A / G=23=17 / R=15=0F → 1画素は 2A 17 0F(B,G,Rの順。RGBではない)
'     1行 = 2A170F × 8 / 全体 = その行 × 8 = 2A170F の 64回くり返し。
'     ヘッダ = 424D F6000000 00000000 36000000 28000000 08000000 08000000
'              0100 1800 00000000 C0000000 130B0000 130B0000 00000000 00000000
'              (=108字=54バイト)
'     全体の16進は 108 + 384 = 492字(=246バイト)。
'     (modTestsPure32 がこの492字をゴールデンとして固定する)
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
' BmpHex - 8×8・24bit BMP ファイル全体(246バイト)を16進文字列で返す【純関数】。
' ----------------------------------------------------------------------------
'   bgColor: VBAのLong色(RGB(r,g,b) = r + g*256 + b*65536)
'   戻り値 : 492字(246バイト×2)の大文字16進。Excel/ファイルI/Oに一切触れない
'            ので LibreOffice の純ロジックテストからそのまま固定できる
'            (SetBackgroundPicture 自体はLOで検証できないため、
'             「敷く中身が正しいか」だけはここで機械的に守る)。
'   8×8 の1行は 8画素×3バイト = 24バイトちょうどで、4バイト境界への
'   パディングが要らない(1×1 のときだけ必要だった。冒頭の検算表を参照)。
' ----------------------------------------------------------------------------
Public Function BmpHex(ByVal bgColor As Long) As String
    Dim r As Long, g As Long, b As Long
    r = bgColor And &HFF&
    g = (bgColor \ 256) And &HFF&
    b = (bgColor \ 65536) And &HFF&
    Dim px As String: px = Hex8(b) & Hex8(g) & Hex8(r)      ' 1画素=BGR順
    Dim ln As String                                        ' 1行=8画素=24バイト
    ln = px & px & px & px & px & px & px & px
    BmpHex = "424D" & Hex32(246) & "00000000" & Hex32(54) & _
             Hex32(40) & Hex32(8) & Hex32(8) & "0100" & "1800" & _
             Hex32(0) & Hex32(192) & Hex32(2835) & Hex32(2835) & _
             Hex32(0) & Hex32(0) & _
             ln & ln & ln & ln & ln & ln & ln & ln          ' 8行=192バイト
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

' MemoDrop - メモから当該シートの1件だけを落とす【純関数】。
'   R33波5a W5-2 で切り出した。従来 MemoPut の前半に埋まっていた「古い1件を
'   落とす」処理そのもので、失敗時に背景画像を剥がしたとき(=もう敷いていない)
'   に「敷いてある」という嘘のメモを消すために、単独で呼べる必要が出た。
'   1件も無いメモ・空メモを渡しても落ちない。全部落ちたら "" を返す
'   (呼び出し側の InStr は "" に対して 0 を返すので、そのまま扱える)。
Public Function MemoDrop(ByVal memo As String, ByVal sheetName As String) As String
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
    MemoDrop = s
End Function

Public Function MemoPut(ByVal memo As String, ByVal sheetName As String, _
                        ByVal bgColor As Long) As String
    Dim s As String: s = MemoDrop(memo, sheetName)
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
    ' R32 Fix波 F7 m-2: On Error Resume Next の直後に Err.Clear が無く、
    ' 上流(呼び出し元の On Error Resume Next 配下)で起きた残留エラーを
    ' 自分のものと誤読して、この関数が【一度も何もせずに】無言で撤退する
    ' 経路があった。しかも下の GoTo CleanExit はログも残さないので、
    ' 「背景が敷かれない」という症状だけが残り原因が追えない。
    Err.Clear

    Dim nm As String
    nm = ws.Name
    If Err.Number <> 0 Then GoTo CleanExit
    If Not TargetSheet(nm) Then GoTo CleanExit

    Dim c As Long
    Err.Clear
    c = modUI.UiColor("bg")
    If Err.Number <> 0 Then GoTo CleanExit
    ' 0 は modSkin.ResolveColor の「未知のキー」センチネル値でもある
    ' (modChrome.ApplyNormalStyleBg の R28H F8(m-3) と同じ理由で異常値扱い)。
    If c = 0 Then GoTo CleanExit
    If InStr(1, mApplied, MemoKey(nm, c), vbBinaryCompare) > 0 Then GoTo CleanExit

    Dim p As String
    p = BmpPath(c)
    If LenB(p) = 0 Then GoTo CleanExit

    ' R32 Fix波 F7 m-1: 失敗の理由は WriteBmpFile に【ByRefで返させる】。
    ' 従来はあちらが最後に Err.Clear してから戻るため、呼び出し側で読む
    ' Err.Number は必ず 0 で、ログに "err=0 " とだけ書かれていた
    ' (=無言失敗を潰したはずのログが、何も語らないログになっていた)。
    Dim wNum As Long, wDesc As String
    If Not WriteBmpFile(p, c, wNum, wDesc) Then
        LogBackdropFailOnce nm, "bmp_write", wNum, wDesc
        GoTo Failed
    End If

    Err.Clear
    ws.SetBackgroundPicture p
    Dim bpNum As Long, bpDesc As String
    bpNum = Err.Number: bpDesc = Err.Description
    Err.Clear
    If bpNum <> 0 Then
        LogBackdropFailOnce nm, "set_background", bpNum, bpDesc
        GoTo Failed
    End If

    ' 成功したときだけメモを進める(F7 m-3)。
    mApplied = MemoPut(mApplied, nm, c)
    SweepOldBmp p
    GoTo CleanExit

Failed:
    ' R33波5a W5-2【R32配布物のリスク・実コードで確認済み】: 失敗経路に
    ' 「既に貼ってある背景画像を剥がす」処理が無かった。背景画像は成功させるか
    ' "" で明示解除するまで残り続けるので、旧テーマ色で敷けている状態から
    ' 敷き直しに失敗すると 境界内=新テーマ色 / 下=旧テーマ色 が確定する ――
    ' 「白い余白」より目立つ「明るいテーマなのに下半分だけ濃紺」になる。
    ' ここで剥がせば最悪でも従来どおり白。
    Err.Clear
    ws.SetBackgroundPicture ""
    Err.Clear
    ' 剥がした以上「敷いてある」というメモも嘘になる。落としておかないと、
    ' テーマを元の色へ戻したときにメモが一致して敷き直しごと弾かれ、
    ' そのシートだけ白いまま戻らなくなる。
    mApplied = MemoDrop(mApplied, nm)

    ' R32 Fix波 F7 m-3: 失敗はメモへ進めない。ただし無制限に再試行すると
    ' 描画のたびにファイルI/Oが走るので、3回で打ち切る。
    '
    ' R33波5a W5-3: その3回は【(シート,色)の組ごと】に数える。従来はモジュール
    ' 共有の1本だったため、最初の組で3回使い切ると以後ずっと3以上のままで、
    ' 2組目以降(3シート×6テーマ=最大18組)は【初回の失敗で即座に】諦めていた
    ' ―― F7 m-3 が潰したはずの欠陥が2組目以降に残っていた。直前の組の鍵だけ
    ' 持ち、組が変わったら数え直す。
    Dim fkey As String
    fkey = MemoKey(nm, c)
    If StrComp(fkey, mFailKey, vbBinaryCompare) <> 0 Then
        mFailKey = fkey
        mFailCount = 0
    End If
    mFailCount = mFailCount + 1
    If mFailCount >= MAX_APPLY_RETRY Then mApplied = MemoPut(mApplied, nm, c)

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
' 残骸は SweepOldBmp が掃除する(色は最大6テーマ=246バイト×6で実害は無い)。
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
' 書き直さない(246バイト固定なのでサイズ一致で十分)。
'
' R32 Fix波 F7 m-1: 失敗した理由を outNum/outDesc で【呼び出し側へ返す】。
'   従来は最後に Err.Clear してから戻っていたため、呼び出し側が読む Err は
'   常に 0 で、ログに "err=0" としか残らなかった。
' R32 Fix波 F8: 存在確認から Dir$ を外した。VBAの Dir はプロセス全体で
'   【状態を1つしか持たない】ので、外側で列挙中に別の Dir$ を始めると
'   その列挙が壊れる。ここは描画経路(Setup*Columns 経由)から呼ばれるため、
'   将来どこかの列挙の内側に入り込む可能性を先に潰しておく。
'   存在確認は FileLen(無ければ実行時エラー53=Err.Numberで判る)で足りる。
Private Function WriteBmpFile(ByVal filePath As String, ByVal bgColor As Long, _
                              ByRef outNum As Long, ByRef outDesc As String) As Boolean
    On Error Resume Next
    Err.Clear
    outNum = 0
    outDesc = ""
    Dim hx As String: hx = BmpHex(bgColor)
    Dim n As Long: n = Len(hx) \ 2

    ' 既存ファイルの長さを見る。無ければ FileLen がエラーになるので、
    ' そのときは「無い」として書きに行く(Dir$ を使わない理由は上の注記)。
    Dim curLen As Long
    Err.Clear
    curLen = FileLen(filePath)
    If Err.Number = 0 Then
        If curLen = n Then
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
    If Err.Number <> 0 Then
        outNum = Err.Number: outDesc = Err.Description
        Err.Clear
        Exit Function
    End If

    Dim fn As Long: fn = FreeFile
    Open filePath For Binary Access Write As #fn
    If Err.Number <> 0 Then
        outNum = Err.Number: outDesc = Err.Description
        Err.Clear
        Exit Function
    End If
    Put #fn, 1, bytes
    ' Put の失敗を先に退避してから Close する(Close は成功して Err を
    ' 0 に戻すため、順番を逆にすると書き込み失敗が消える)。
    outNum = Err.Number: outDesc = Err.Description
    Err.Clear
    Close #fn
    If outNum = 0 Then
        outNum = Err.Number: outDesc = Err.Description
    End If
    WriteBmpFile = (outNum = 0)
    Err.Clear
    On Error GoTo 0
End Function

' 今使っているBMP以外の MyBookshelf_bg_*.bmp を消す(失敗は握って続行)。
'
' R32マイクロ修正波 F18: %TEMP%全体を列挙するのをやめ、列挙ゼロにした。
'   【F8で直したはずが列挙が残っていたこと】F8は「VBAのDir$はプロセスに
'   状態を1つしか持たないので、外側のDir$列挙を壊す」問題を、
'   Scripting.FileSystemObjectでの列挙(fld.Files)へ置き換えて直した。だが
'   これは「壊さない列挙」であって「列挙しない」ではない。長期運用PCの
'   %TEMP%は数千〜数万件が普通で、この関数は描画経路(ScreenUpdating=False
'   中・テーマ切替時)から呼ばれるため、%TEMP%全件ぶんのCOMオブジェクト
'   (f.Nameアクセスのたび)を毎回生成するコストがそこに乗っていた。しかも
'   消す対象は色ごと最大6テーマ×246バイトで、コメント(旧BmpPathの注記)
'   自身が「残骸は実害が無い」と自認している ―― 実害の無い掃除のために
'   描画経路を数千〜数万件ぶん重くする理由が無い。
'
'   直し方: BmpFileName(色)が色から一意に決まる【決定的な名前】である
'   性質を使う。「今どのファイルが在るか」を尋ねる(列挙する)必要は無く、
'   「在り得る名前を先に数え上げて、使っていなければ消してみる」だけで足りる
'   (無ければ Kill が失敗するだけで実害は無い。On Error Resume Nextで握る)。
'   6テーマの bg 色は modSkin.ResolveColor(Public関数)から取る ―― 色の
'   一次情報を modBackdrop 側へ複製すると、将来 modSkin 側で配色が変わった
'   ときにここだけ古いままになる(RGB値の二重管理を避ける。modSkinは
'   残27字で編集不可のため、呼ぶだけで実体は増やさない)。テーマ名6件の
'   並びは modSkin.CycleSkin の orderList / modTestsPure23 と同じ固定リスト
'   (新テーマが増えたらこの配列に1件足すだけでよい)。
'   light/msadは実際は同色(RGB(243,244,246))だが、そのぶんBmpFileNameが
'   同じ名前になるだけで、2回目のKillは「無い」失敗として握られるだけ。
Private Sub SweepOldBmp(ByVal keepPath As String)
    On Error Resume Next
    Err.Clear
    Dim sep As Long: sep = InStrRev(keepPath, "\")
    If sep < 1 Then Exit Sub
    Dim dir_ As String: dir_ = Left$(keepPath, sep)
    Dim keepName As String: keepName = Mid$(keepPath, sep + 1)

    Dim themes As Variant
    themes = Array("msad", "dark", "light", "sakura", "ocean", "gold")

    Dim i As Long
    For i = LBound(themes) To UBound(themes)
        Dim c As Long
        Err.Clear
        c = modSkin.ResolveColor("bg", CStr(themes(i)))
        If Err.Number = 0 Then
            Dim candName As String: candName = BmpFileName(c)
            If StrComp(candName, keepName, vbTextCompare) <> 0 Then
                Kill dir_ & candName
                Err.Clear
            End If
        End If
        Err.Clear
    Next i
    On Error GoTo 0
End Sub

' ============================================================================
' W5-1: 条件付き書式方式(2026-08-16 R33波5a)。背景画像との【二重防御】。
' ============================================================================
' なぜ第2案が要るのか: :137 の「SetBackgroundPicture が最後の1本」は誤り
'   だった。条件付き書式は8ラウンド一度も検討されていない(src/ を grep して
'   利用は0件。唯一の言及は tools/README §4 の「LO では検証できない」)。
'   背景画像は %TEMP% へのファイル生成に依存し、書込不可・ディスク満杯・
'   ウイルス対策の隔離・Environ 空 のどれか1つで不発になる。条件付き書式は
'   そのどれにも依存しない。排他にせず両方を掛ける。
'
' 機構: 条件付き書式は「セルごとの書式レコード」を作らず、シート単位に
'   「適用範囲(sqref)1件+ルール1件」として保持され、ルール側の Interior が
'   地を上書き描画する。本当なら B を1行も増やさずに下の地を色づけられる。
'
' ★成立条件はただ1つ:【張っても UsedRange が伸びないこと】。伸びれば B が
'   下がり停止線 S=B+k も同じだけ下がるので元の木阿弥(R18〜R27を溶かした壁)。
'   これは LO では検証できない(tools/README §4)ので【実行時に自分で検算する】:
'   張る直前と直後の使用済み末尾を測り、伸びていたら剥がし、伸びたぶんの行を
'   Rows.Delete で戻し(解放できるのはこれだけ=R30実機実証)、以後張らない。
'   成否どちらでも usage_log に1行残す(無言failで8ラウンド溶かしている)。
'
' 深さ(仕様との差分・報告済み): 仕様は Rows(B+1:1048576) だが CF_DEPTH_ROWS
'   行に留めた。(a)要る深さは k(=1画面。窓高750pt/行高18ptで約42行)までで、
'   400行は行高換算7,200pt=最大窓高の約10倍。(b)全域(約17億セル)の描画コスト
'   が未検証で、最も重い苦情「フリーズ」を新たに作りかねない。
'
' 対象: Apply と同じ TargetSheet の3枚。チャット(Nexus)は modUI.bas:169 が
'   UserInterfaceOnly:=True で Protect されており設定自体は通るはずだが実機
'   未検証。背景画像と同じくこの波では対象外(範囲を広げない)。
'
' 呼び口: modViewport.ReleaseSheetRowsBelow の末尾。境界が確定した【直後】に
'   走る必要があり(Setup*Columns 経由の ApplyNormalStyleBg は描画前で B が
'   未確定)、かつ4画面が必ず通る唯一の合流点。呼び出し側4本は容量が無い。
' ============================================================================

' CfFormula - 常に真になる数式【純関数】。ISTEXT("MBSBG") は引数が文字列
'   リテラルなのでどのセルで評価されても必ずTRUE。セル参照が無いので相対参照
'   のズレも起きない。数式に目印(CF_MARK)が残るのでClearOwnCFが自分のルールを
'   見分けられる。関数名はVBAから設定する限り常に英語(日本語Excelでも同じ)。
Public Function CfFormula() As String
    CfFormula = "=ISTEXT(""" & CF_MARK & """)"
End Function

' CfStartRow - 張り始める行【純関数】。boundRow=解放の下端行 / usedLast=解放
'   【後】に実測した使用済み末尾行。使用済みの行に絶対に被せない(被せると本文が
'   ルールの地色で潰れ、カードや見出しの塗り分けが消える)ので大きい方を採る。
'   行1に掛かる指定・最終行超えは 0(=張らない)。
Public Function CfStartRow(ByVal boundRow As Long, ByVal usedLast As Long) As Long
    Dim s As Long
    s = boundRow
    If usedLast + 1 > s Then s = usedLast + 1
    If s < 2 Then Exit Function
    If s > CF_MAX_ROW Then Exit Function
    CfStartRow = s
End Function

' CfRowsAddr - "開始行:終了行" の行アドレス【純関数】。depth 行ぶん下まで
'   (最終行でクランプ)。異常な指定(行1以下・深さ0以下)は空文字=何もしない。
Public Function CfRowsAddr(ByVal startRow As Long, ByVal depth As Long) As String
    If startRow < 2 Then Exit Function
    If startRow > CF_MAX_ROW Then Exit Function
    If depth < 1 Then Exit Function
    Dim e As Long
    e = startRow + depth - 1
    If e > CF_MAX_ROW Then e = CF_MAX_ROW
    If e < startRow Then Exit Function
    CfRowsAddr = CStr(startRow) & ":" & CStr(e)
End Function

' ApplyCF - 境界より下の地を条件付き書式1本でテーマ色にする(冪等)。
'   ws=対象シート(TargetSheet 以外は何もしない) / boundRow=解放の下端行。
'   冪等化は Delete→Add の張り替え。消すのは自分が張った1本だけ(ClearOwnCF)。
Public Sub ApplyCF(ByVal ws As Worksheet, ByVal boundRow As Long)
    If ws Is Nothing Then Exit Sub
    If mCfOff Then Exit Sub
    On Error Resume Next
    ' Apply と同じ理由(F7 m-2)で、上流の残留エラーを自分のものと誤読しない。
    Err.Clear

    Dim nm As String
    nm = ws.Name
    If Err.Number <> 0 Then GoTo CfExit
    If Not TargetSheet(nm) Then GoTo CfExit

    Dim c As Long
    Err.Clear
    c = modUI.UiColor("bg")
    If Err.Number <> 0 Then GoTo CfExit
    ' 0 は modSkin.ResolveColor の「未知のキー」センチネル(Apply と同じ扱い)。
    If c = 0 Then GoTo CfExit

    ' 行解放の【後】の実測値。ここが停止線 S = B + k の B。
    Dim usedLast As Long
    Err.Clear
    usedLast = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If Err.Number <> 0 Then GoTo CfExit

    Dim addr As String
    addr = CfRowsAddr(CfStartRow(boundRow, usedLast), CF_DEPTH_ROWS)
    If LenB(addr) = 0 Then GoTo CfExit

    ClearOwnCF ws

    Err.Clear
    Dim fc As Object
    Set fc = ws.Rows(addr).FormatConditions.Add(CF_TYPE_EXPRESSION, , CfFormula())
    Dim aNum As Long, aDesc As String
    aNum = Err.Number: aDesc = Err.Description
    Err.Clear
    If aNum <> 0 Then
        LogCfOnce nm, "cf_add", aNum, aDesc
        GoTo CfExit
    End If

    fc.Interior.Color = c
    aNum = Err.Number: aDesc = Err.Description
    Err.Clear
    If aNum <> 0 Then
        LogCfOnce nm, "cf_interior", aNum, aDesc
        ClearOwnCF ws
        GoTo CfExit
    End If

    ' ★必須検算。ここが成否を分ける唯一の点(この波の主題)。
    Dim after As Long
    Err.Clear
    after = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If Err.Number = 0 Then
        If after > usedLast Then
            ' 伸びた=この端末では条件付き書式もセルを使う。剥がして、伸びた
            ' ぶんの行を消して元の B に戻し、以後このセッションでは張らない。
            ClearOwnCF ws
            Err.Clear
            ws.Rows((usedLast + 1) & ":" & after).Delete
            Err.Clear
            mCfOff = True
            LogCfOnce nm, "usedrange_grew", after - usedLast, _
                "before=" & usedLast & " after=" & after & " range=" & addr
            GoTo CfExit
        End If
    End If
    Err.Clear

    ' 成立した。1セッション1回だけ「効いている」ことも残す ―― 実機で
    ' 「本当に張れたのか」を後から1行で確かめられるようにするため。
    If Not mCfOkLogged Then
        mCfOkLogged = True
        modLog.LogUsage "backdrop_cf", nm, _
            "range=" & addr & " color=" & c & " used=" & usedLast
        Err.Clear
    End If

CfExit:
    Err.Clear
    On Error GoTo 0
End Sub

' ClearOwnCF - 自分が張ったルールだけ落とす。ws.Cells.FormatConditions.Delete
'   は【全ての】条件付き書式を消すので使わない(src全体の利用は現時点で本
'   モジュールだけ=R33波5a実測。将来他機能を巻き込む作りにしない)。
'   目印はCF_MARK。後ろから回す(添字ズレ回避)。
Private Sub ClearOwnCF(ByVal ws As Worksheet)
    On Error Resume Next
    Err.Clear
    Dim fcs As Object
    Set fcs = ws.Cells.FormatConditions
    If Err.Number <> 0 Then GoTo ClearDone
    Dim n As Long
    n = fcs.Count
    If Err.Number <> 0 Then GoTo ClearDone
    Dim i As Long
    For i = n To 1 Step -1
        Dim f As String
        Err.Clear
        f = fcs.Item(i).Formula1
        If Err.Number = 0 Then
            If InStr(1, f, CF_MARK, vbTextCompare) > 0 Then
                fcs.Item(i).Delete
            End If
        End If
        Err.Clear
    Next i
ClearDone:
    Err.Clear
    On Error GoTo 0
End Sub

' 条件付き書式まわりの不首尾を1セッション1回だけ usage_log へ。検算に落ちた
'   ("usedrange_grew")場合もここを通る(成立しなかったことこそ残す)。
Private Sub LogCfOnce(ByVal sheetName As String, ByVal stage As String, _
                      ByVal errNum As Long, ByVal errDesc As String)
    If mCfFailLogged Then Exit Sub
    mCfFailLogged = True
    On Error Resume Next
    modLog.LogUsage "backdrop_cf_failed", sheetName, _
        "stage=" & stage & " n=" & errNum & " " & errDesc
    Err.Clear
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
