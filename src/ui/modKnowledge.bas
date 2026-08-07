Attribute VB_Name = "modKnowledge"
Option Explicit

' modKnowledge - ナレッジ画面の共通クロム(ヘッダー+モード切替+ツールバー)。
' 2026-07-26再設計: Vault(カードギャラリー)/マイ本棚(一覧表)の2シートに
' 同じヘッダー・ツールバーを描き、上部ピルで切り替える1画面2モードに見せる。
' ツールバーは既存のmodVault/modUIShelf/modPack/modShelfSyncへ配線するだけ。
' 設計の鉄則: Shape座標は実セル幾何から導く(旧実装の決め打ち座標が画面外に
' はみ出していた教訓)。配色はmodUI.UiColor()。Shape名はnxk_接頭辞。

Private Const HDR_H As Double = 40
Private Const PILL_H As Double = 26
' 右肩ピル(R7 A-5): モード3つ + 🎨着せ替え + 🚪終了 の5個。
' 2026-07-31(R11-F2): 「?」ヘルプを足して6個(監査1 L-3。ナレッジ画面から
' ヘルプへ行く手段が無く、困った人が詰む状態だった。R11-Eでは容量不足で保留)。
' 2026-08-05(R18-3c): 「💬 チャットへ」を下段ツールバーから移して7個
' (実機第5報②。下段はボタン数で位置が動くため、常設の移動導線に向かない)。
Private Const PILL_N As Long = 7
Private Const PILL_PITCH As Double = 9
Private Const PILL_PAD As Double = 14
Private Const PILL_MIN As Double = 44

' 旧ギャラリー/解決事例の描画先だった実行時生成シート。2026-07-30(R4要件A)で
' 描画先を「マイ本棚」へ統合したため、既存ブックに残っているものを消すためだけに
' 名前を持っている(新規に作ることは二度と無い)。
Private Const LEGACY_VAULT_SHEET As String = "Vault"

' 上部クロムが占める行(1..6)。本文はDrawChrome後の ContentTop から下に描く。
' 行3がツールバーの帯で、段数に応じて高さが伸びる(行数は変えない)。
Public Const CHROME_ROWS As Long = 6

' R18-3a/3b → R19-1b: 「マイ本棚」シート(table/gallery/shared の3モードが
' 共有)が使うセル範囲。列は3モードとも A:N で、余りは最終列Nが吸う。
' 行の【上限】はカードの最終行(modUIShelf.MAX_CARD_ROWS=400 の末尾=412行)
' だが、常に412行(約6,200pt=8画面ぶん)を塗ると「ほぼ無制限に下へ
' スクロールできる」状態になるため、実際の範囲は描いた内容の下端から
' ShelfBound で決める。3モジュールで同じ算数を書くとズレるので、共通クロムを
' 持つここが単一情報源(モードごとに実下端が違うので引数で受ける)。
Public Const SHELF_MAX_ROW As Long = 412
Public Const SHELF_PAD_COL As String = "N"
Public Const SHELF_BAND As String = "A1:N1"

' 直近に描いたモード("gallery"/"table"/"shared")。
' 3モードとも同じ「マイ本棚」シートに描くようになった(2026-07-30 R4要件A)ため、
' 「今どのモードか」はシート名からは分からない。ここが唯一の情報源。
Private mMode As String

' 直近に行高を明示した最終行(3モード共有。R20-1d。詳細は NormalizeShelfRows)。
Private mShelfRowHigh As Long

' ShelfBound - マイ本棚の実使用範囲 "A1:N<行>"。塗り・ScrollArea が共有する。
'   contentBottom: そのモードが描き切った実下端(pt)。0なら1画面ぶん。
Public Function ShelfBound(ByVal ws As Worksheet, ByVal contentBottom As Double) As String
    If ws Is Nothing Then Exit Function
    ShelfBound = modViewport.BoundAddr(ws, SHELF_PAD_COL, contentBottom, SHELF_MAX_ROW)
End Function

' ----------------------------------------------------------------------------
' 行高の後始末(R20-1d・実機第7報⑦の層2)。3モードが共有する状態なのでここ。
' ----------------------------------------------------------------------------
' 3モードとも冒頭で Rows("7:412").RowHeight = 15 と、406行ぶんの行高を毎回
' 書き直していた。行高を明示した行はExcelから見れば「使用済み」なので、
' 資料が3冊しか無くても常に412行(約6,200pt=8画面ぶん)下へホイールで
' 転がれる状態が残る ―― 塗りと ScrollArea を実下端まで縮めても消えない
' (ScrollArea はホイールを止めない。modViewport 冒頭参照)。
' 「前回どこまで使ったか」を覚えて、そこまでだけ均す。
'   0 = このセッションではまだ均していない。既存ブックに焼き付いた行高を
'   救済するため、初回だけは従来どおり最終行まで均す(セッション1回だけ)。
'   (mShelfRowHigh の宣言は他のモジュール変数と一緒に先頭へ置いてある)

' NormalizeShelfRows - 前モードが残した可変行高を15ptへ戻す(fromRow以降)。
Public Sub NormalizeShelfRows(ByVal ws As Worksheet, ByVal fromRow As Long)
    If ws Is Nothing Then Exit Sub
    Dim r0 As Long: r0 = fromRow
    If r0 < 1 Then r0 = 1
    Dim r1 As Long: r1 = ShelfRowHigh()
    If r1 > SHELF_MAX_ROW Then r1 = SHELF_MAX_ROW
    If r1 < r0 Then Exit Sub
    On Error Resume Next
    ws.Rows(r0 & ":" & r1).RowHeight = 15
    On Error GoTo 0
End Sub

' ShelfRowHigh - 直近に行高を明示した最終行(未初期化なら最終行=全域救済)。
Public Function ShelfRowHigh() As Long
    ShelfRowHigh = mShelfRowHigh
    If ShelfRowHigh < 1 Then ShelfRowHigh = SHELF_MAX_ROW
End Function

' ApplyShelfBound - 描き終えた実下端から、塗り・ScrollArea・行高の後始末を
'   1本で行う(3モードが同じ後始末を3通りに書かないための単一情報源)。
'   contentBottom: そのモードが描き切った実下端(pt)
'   paintBg      : 地色も塗り直すか。一覧表モードは False(見出し行のグレーや
'                  カードの塗り分けを、描き終えた後の一律塗りで潰さないため)。
'   戻り値: 実際に適用した範囲アドレス("" なら何もできなかった)。
Public Function ApplyShelfBound(ByVal ws As Worksheet, ByVal contentBottom As Double, _
                                Optional ByVal paintBg As Boolean = True) As String
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Dim addr As String: addr = ShelfBound(ws, contentBottom)
    If LenB(addr) = 0 Then Exit Function
    If paintBg Then ws.Range(addr).Interior.Color = modUI.UiColor("bg")
    modViewport.ApplyScrollBound ws, addr
    ' 受け入れ基準(R20-1d): この境界の下端行より下に、行高カスタムを残さない。
    ' addr は必ず A1 起点なので、行数がそのまま下端行になる。
    Dim lastR As Long: lastR = ws.Range(addr).Rows.Count
    If lastR > 0 Then
        modViewport.ResetRowsBelow ws, lastR + 1, SHELF_MAX_ROW
        mShelfRowHigh = lastR
    End If
    ApplyShelfBound = addr
    ' R21-S7: フィット直後の5値観測点(窓幅/可視幅/帯実幅/中身右端/境界下端)。
    modViewport2.LogFit ws, "shelf-" & CurrentMode(), SHELF_BAND, addr
    On Error GoTo 0
End Function

' DrawChrome - ヘッダー+モードピル+ツールバーを描く(冪等)。
'   mode: "gallery"(ナレッジ倉庫のカード) / "table"(マイ本棚の一覧)
Public Sub DrawChrome(ByVal ws As Worksheet, ByVal mode As String)
    If ws Is Nothing Then Exit Sub
    On Error GoTo Fail

    RemoveChrome ws
    mMode = LCase$(mode)

    ' 幾何を先に確定させる(順序が逆だとShape座標がズレる)。
    ' 行1(ヘッダー)と行2(隙間)だけ先に決めれば 行3.Top は確定する
    ' (行3自身の高さはツールバーが何段になったかが分かってから入れる)。
    ws.Rows(1).RowHeight = HDR_H
    ws.Rows(2).RowHeight = 6

    ' 罫線と行列番号を隠す。ここを消さないと、どれだけ整えても
    ' 画面が「Excelのシート」にしか見えない(実機要望: エクセル感を消す)。
    On Error Resume Next
    If ThisWorkbook.ActiveSheet Is ws Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If
    On Error GoTo Fail

    Dim L As Double, cellW As Double, W As Double
    L = ws.Range("A1").Left
    ' R19-1b: 余りを最終列Nに吸わせて A:N の合計を可視幅ぴったりにする
    ' (右の白い余白は寸法の問題で、ScrollAreaでは消せない)。帯の右端も
    ' modViewport.ContentRight 1本から取る。
    ' R21-S3: 吸収列はモードごとに1つだけ(2段階Fitの廃止。ShelfPadCol参照)。
    modViewport.FitBandToViewport ws, SHELF_BAND, modViewport2.ShelfPadCol(mMode, SHELF_PAD_COL)
    cellW = modViewport.ContentRight(ws, SHELF_BAND, 0) - L
    ' R19H FB-6(A-L⑬): 帯の【塗り幅】だけは max(可視幅, 内容幅) にする。
    ' ContentRight は可視幅で頭打ちにするのが正しい(操作系を画面外へ出さない
    ' ため)が、本棚系は列の最小幅の制約で帯 A:N が可視幅より広くなることが
    ' ある(狭い窓)。そのとき可視幅で切ると、横スクロールした先で帯が途中で
    ' 終わり、右側だけ地色の白が出る=「帯が切れて見える」。塗りは内容幅まで
    ' 伸ばしても押せるものが画面外へ出ないので、広い側へ倒して構わない
    ' (ContentRight の返り値そのものは変えない=操作系の右端は従来どおり)。
    Dim bandW As Double: bandW = ws.Range(SHELF_BAND).Width
    If bandW > cellW Then cellW = bandW
    ' R11-B(#30本丸): セル範囲幅(cellW)基準のクロムは本棚系で772〜882ptに
    ' なり、実可視域(約600pt)を大幅超過して右肩ピルが画面外へ出ていた。
    ' 帯の背景(下のhdr)は従来どおりcellWいっぱいのまま、操作系を置く
    ' 右端(W)だけを実可視幅(modUIMain.ViewportWidth)へクランプする。
    W = modChrome.BarWidth(cellW, modUIMain.ViewportWidth(), 8)

    ' --- 右肩ピルの配置を先に決める(帯の高さがこれで決まるため) ---
    ' 2026-07-31(R7 A-5): 🎨着せ替え / 🚪終了 をここへ統合し、
    ' 「どの画面でも右上は同じ」にする(並びはチャット画面と同じで一番右が🚪)。
    ' 位置は固定座標をやめて modChrome.FlowRight に任せる。1段に入らなければ
    ' 段を増やし、帯の高さもそれに合わせて伸ばす=画面外や見切れが起きない。
    Dim pcaps(0 To PILL_N - 1) As String, pnames(0 To PILL_N - 1) As String
    Dim pacts(0 To PILL_N - 1) As String, pwid(0 To PILL_N - 1) As Double
    Dim pactive(0 To PILL_N - 1) As Boolean
    Dim md As String: md = LCase$(mode)
    Dim isTable As Boolean: isTable = (md = "table")
    Dim isShared As Boolean: isShared = (md = "shared")
    PillSpec md, isTable, isShared, pcaps, pnames, pacts, pactive, pwid

    ' R14-2a(実機第3報 RC10): ピルは帯の右端(L+W-8)へ密着させていたが、
    ' ツールバーはボタン数が減ると帯の右端まで届かず、両者の右端が
    ' 食い違って見えていた。ツールバーの実際の右端へピルの右アンカーを
    ' 合わせる。isTable/isSharedは下のDrawToolbar呼び出しと同じ値
    ' (このSub内で唯一の判定)。
    ' R14-G13: フォールバック条件を「L+200未満」から「値が縮退しているとき」
    ' へ変えた。ボタンが少ない端末ほどツールバーの右端は左に来るのに、
    ' L+200 未満だと丸ごとフォールバックして帯の右端(L+W-8)へ戻り、
    ' 【この修正が一番効くはずの構成】でだけ整列しないという逆転が起きていた。
    ' 縮退=ボタン0個(ToolbarContentRightは0を返す)か、左余白(TB_PAD=8)にも
    ' 届かない値。そのときだけ従来の右端へ倒す。
    ' R20-1c(実機第7報⑦の層1): アンカーを【帯の右端】へ移す。ツールバーは
    ' modKnowledgeBar が左詰めで流し込む固定幅の並びで、窓を広げても伸びない。
    ' そこへピルを揃えると、窓が広いほどピルだけが画面の左寄りで止まり、
    ' 帯の右に何百ptもの白が残る(R14-2a は「窓が狭い」前提での整列合わせだった)。
    ' 右端の単一情報源は modViewport.ContentRight ―― 帯・ナビ・フッターと同じ式。
    Dim tbRight As Double
    On Error Resume Next
    tbRight = modViewport.ContentRight(ws, SHELF_BAND, 8)
    On Error GoTo Fail
    If tbRight <= L + 8 Then tbRight = L + W - 8

    Dim pxs() As Double, prows() As Long, pws() As Double
    Dim pillRowN As Long
    pillRowN = modChrome.FlowRight(pwid, PILL_N, tbRight, _
                                   L + modChrome.TitleReserve(W, 160), L + 8, 6, _
                                   pxs, prows, pws)
    If pillRowN < 1 Then pillRowN = 1
    Dim hdrH As Double: hdrH = HDR_H + (pillRowN - 1) * (PILL_H + 2)
    If hdrH > 200 Then hdrH = 200          ' 行高の異常値でDrawChrome全体を落とさない
    ws.Rows(1).RowHeight = hdrH

    ' --- ヘッダーバー ---
    Dim hdr As Shape
    Set hdr = ws.Shapes.AddShape(5, L, 0, cellW, hdrH)
    hdr.Name = "nxk_hdr"
    hdr.Adjustments(1) = 0.02
    hdr.Line.Visible = 0
    hdr.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    modSkin.ApplyHeaderDepth hdr          ' §9: 濃紺の2色グラデーション
    With hdr.TextFrame2
        .TextRange.Text = ChrW(&HD83D) & ChrW(&HDCDA) & " ナレッジ"
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .MarginLeft = 82
        .VerticalAnchor = 3
    End With

    ' 実機報告(2026-07-27)「←Hubがヘッダーと同色で目立たない」対策。
    ' 戻り導線は一番見つけやすくなければならないので、白ピルで強調する。
    Pill ws, "nxk_back", ChrW(&H2190) & " Hub", L + 8, 72, 0, _
         "modKnowledge.OnBackHub", True

    ' --- モード切替ピル+共通ヘッダー機能(右肩) ---
    Dim pi As Long
    For pi = 0 To PILL_N - 1
        Pill ws, pnames(pi), pcaps(pi), pxs(pi), pws(pi), _
             prows(pi) * (PILL_H + 2), pacts(pi), pactive(pi)
    Next pi

    On Error Resume Next
    modTelemetry.TrackScreen "knowledge"
    On Error GoTo Fail

    ' --- ツールバー(行3の帯) ---
    ' 2026-07-30(R4要件C): 折り返しが「1段目だけ」だったため、2段目に
    ' あふれたボタンは幅を見ずに右へ描き続けられ、画面外で見切れていた
    ' (実機写真の「🗑削除が『除』しか見えない」)。段数無制限の流し込みに
    ' 統一し、実際に使った高さをそのまま帯(行3)の高さにする。
    ' こうすると検索欄(行5)もカード領域(行7以降)も自動で下がり、
    ' 何段になっても重ならない。
    Dim barH As Double
    barH = modKnowledgeBar.DrawToolbar(ws, isTable, isShared, L, W, ws.Rows(3).Top)
    If barH < modKnowledgeBar.BAR_H + 2 Then barH = modKnowledgeBar.BAR_H + 2
    ' 3-C: Excelの行高上限は409.5pt。段数無制限にしたので、極端に狭い帯
    ' (異常な列幅設定など)では計算上それを超えうる。超えた値を代入すると
    ' 1004になり、そこから先(検索欄・カード)が丸ごと描かれない。
    ' 手前でクランプして「崩れても描き切る」ほうを選ぶ。
    If barH > 400 Then barH = 400
    ws.Rows(3).RowHeight = barH
    ws.Rows(4).RowHeight = 2
    ws.Rows(5).RowHeight = 22
    ws.Rows(6).RowHeight = 8

    ' R18-3b: 3モード(table/gallery/shared)とも同じシートで同じ範囲を使う。
    ' 共通クロムを描くここが唯一の宣言点(モードごとに書くとズレる)。
    ' R19-1b: この時点では本文がまだ無いので「最低1画面」ぶん。本文を描き
    ' 終えた各モードが、実下端で ApplyScrollBound を上書きする。
    ' R21-S7: 観測点は本文を描き終えた ApplyShelfBound 側へ移した。
    modViewport.ApplyScrollBound ws, ShelfBound(ws, 0)

    On Error Resume Next
    modUI.FreezeShapePlacement ws
    On Error GoTo 0
    Exit Sub

Fail:
    modLog.LogError "E0801", "modKnowledge.DrawChrome", Err.Description, Err.Number
End Sub

' ----------------------------------------------------------------------------
' PrepareScreenView - 「マイ本棚」「Hub」を表示・再描画するときの共通儀式。
' ----------------------------------------------------------------------------
' 2026-07-30(R4要件B): 実機で「🗑削除が『除』しか見えない」「ボタンが右で
' 切れている」と報告された画面は、多くが「一度右へスクロールした状態が
' 持ち越されているだけ」だった。Nexus起動時に水平スクロールバーを消して
' あるため、利用者には戻す手段が無い。表示のたびに必ず左上へ戻す。
' あわせて、旧"Vault"シートが残っているブックの移行掃除もここで行う
' (表示経路すべてがここを通るので、掃除の呼び忘れが起きない)。
' 罫線・行列番号の表示制御は各画面の既存処理に任せる(ここでは触らない)。
Public Sub PrepareScreenView(ByVal ws As Worksheet)
    PurgeLegacyVaultSheet
    ' 2026-07-31(R7 A-2): 全画面/数式バー/罫線の崩れをここでも自己修復する。
    ' 「閉じる→キャンセル」で壊れた表示が、次にどのボタンを押しても戻る。
    On Error Resume Next
    modUI.EnsureAppView
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ' ActiveWindow系は、そのシートが実際に前面のときだけ触る
    ' (別シートの表示状態を巻き添えで変えないため)。
    If ThisWorkbook.ActiveSheet Is ws Then
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        ActiveWindow.Zoom = 100
    End If
    On Error GoTo 0
End Sub

' 旧"Vault"シートの移行削除。ギャラリーと解決事例の描画先を「マイ本棚」へ
' 統合した(R4要件A)ので、既存ブックに残る空の"Vault"タブは
' 「同じ資料が2つの画面にある」という元の混乱をそのまま残してしまう。
Private Sub PurgeLegacyVaultSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LEGACY_VAULT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' 2026-07-30(レビュー2-D/7-B): DisplayAlerts は変更前の値を控えて戻す
    ' (True 決め打ちで戻すと、警告を切って一括処理している最中に呼ばれた
    ' ときに、その処理の途中から警告ダイアログが出るようになる)。
    ' また Delete に失敗した場合、可視化したままにすると旧"Vault"タブが
    ' 画面に出てしまう(消せなかったうえに混乱だけが増える)ので、
    ' まだ残っていたら veryHidden へ戻す。
    Dim prevAlerts As Boolean: prevAlerts = True
    On Error Resume Next
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Visible = -1          ' veryHiddenのままだとDeleteが1004になる環境がある
    ws.Delete
    Application.DisplayAlerts = prevAlerts
    On Error GoTo 0

    ' 消えたかどうかを取り直して確認する(消えていれば参照は失敗する)。
    Dim still As Worksheet
    On Error Resume Next
    Set still = ThisWorkbook.Worksheets(LEGACY_VAULT_SHEET)
    On Error GoTo 0
    If still Is Nothing Then Exit Sub
    On Error Resume Next
    still.Visible = 2        ' xlSheetVeryHidden(消せなかったので隠し直す)
    modLog.LogError "E0801", "modKnowledge(PurgeLegacyVaultSheet)", _
        "旧Vaultシートを削除できなかったため veryHidden へ戻しました"
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' CurrentMode - 現在モードの正規化(このモジュールで唯一の解釈)。
' ----------------------------------------------------------------------------
' 2026-07-30(レビュー1-A): mMode が "" (起動直後・未描画)のときの解釈が
' IsTableMode(=一覧表とみなす)と RefreshCurrent(Case Else でギャラリー)で
' 食い違っていた。取込直後の自動再描画が「一覧表のつもりで書いた上に
' ギャラリーを描く」といった噛み合わない動きの原因になる。
' 既定の画面は一覧表なので、"" は "table" に寄せる。両方ここを使う。
Private Function CurrentMode() As String
    Dim m As String: m = LCase$(Trim$(mMode))
    If m = "gallery" Or m = "shared" Then
        CurrentMode = m
    Else
        CurrentMode = "table"
    End If
End Function

' 今「マイ本棚」シートに描かれているのが一覧表モードかどうか。
' 3モードが同じシートを共有する(R4要件A)ため、取込・同期の完了時に
' 走る自動再描画(modShelf/modShelfSync → modUIShelf.RenderShelf)が、
' ギャラリー表示中に一覧表をカードの裏へ書き込んでしまう経路ができた。
' RenderShelf側でここを見て空振りさせる。
' 未描画(起動直後)は一覧表とみなす ―― 既定の画面が一覧表だから。
Public Function IsTableMode() As Boolean
    IsTableMode = (CurrentMode() = "table")
End Function

' 本文(カード/表)を描き始めてよいY座標。
Public Function ContentTop(ByVal ws As Worksheet) As Double
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    ContentTop = ws.Rows(CHROME_ROWS + 1).Top
    On Error GoTo 0
End Function

' 検索キーワードのセル(ギャラリーだけが使う)。
Public Function SearchCellAddress() As String
    SearchCellAddress = "B5:E5"
End Function

' ----------------------------------------------------------------------------
' PillSpec - 右肩に並べるものの唯一の定義(2026-07-31 R7 A-5)。
'   FlowRightは配列の先頭から右詰めで置くので、画面上の並びは配列の逆順:
'     ← Hub | 💬チャットへ | ギャラリー | マイ本棚 | みんなの解決事例 |
'     🎨着せ替え | 🚪終了 | ❓
'   これはダッシュボード(modDash.HeaderSpec・FlowLeftで同じ見た目の並び)と
'   同一の順序。ナレッジ画面だけが違う並びだと「どの画面でも右上は同じ」が
'   崩れる(2026-08-05 R18-3c。従来は❓が左端・🚪終了が右端だった)。
'   幅はキャプションの実文字から出す(固定幅の予約と実物が食い違うと、
'   実機で「🗑削除が『除』しか見えない」類の見切れになる。R4要件Cの教訓)。
' ----------------------------------------------------------------------------
Private Sub PillSpec(ByVal md As String, ByVal isTable As Boolean, ByVal isShared As Boolean, _
                     ByRef caps() As String, ByRef nms() As String, ByRef acts() As String, _
                     ByRef actives() As Boolean, ByRef widths() As Double)
    ' 「?」= 使い方。ヘルプカードはチャット画面(Nexus)の上に描く作りなので、
    ' modHub.OnHelp と同じくチャットへ移ってから開く(2026-07-31 R11-F2)。
    caps(0) = ChrW(&H2753)
    nms(0) = "nxk_help": acts(0) = "modKnowledge.OnHelp": actives(0) = False
    caps(1) = ChrW(&HD83D) & ChrW(&HDEAA) & " 終了"
    nms(1) = "nxk_exit": acts(1) = "modApp.OnSaveAndExit": actives(1) = False
    caps(2) = ChrW(&HD83C) & ChrW(&HDFA8) & " 着せ替え"
    nms(2) = "nxk_skin": acts(2) = "modHub.OnThemeToggle": actives(2) = False
    caps(3) = ChrW(&HD83C) & ChrW(&HDF81) & " みんなの解決事例"
    nms(3) = "nxk_m_shared": acts(3) = "modKnowledge.OnGoShared": actives(3) = isShared
    caps(4) = ChrW(&HD83D) & ChrW(&HDCCB) & " マイ本棚"
    nms(4) = "nxk_m_table": acts(4) = "modKnowledge.OnGoTable": actives(4) = isTable
    caps(5) = ChrW(&HD83C) & ChrW(&HDCCF) & " ギャラリー"
    nms(5) = "nxk_m_gallery": acts(5) = "modKnowledge.OnGoGallery"
    actives(5) = (Not isTable) And (Not isShared)
    ' R18-3c: チャットへの導線。ハンドラは既存のOnToChatを再利用する
    ' (下段ツールバーから移しただけで、押したときの動きは変わらない)。
    caps(6) = ChrW(&HD83D) & ChrW(&HDCAC) & " チャットへ"
    nms(6) = "nxk_chat": acts(6) = "modKnowledge.OnToChat": actives(6) = False

    Dim i As Long
    For i = 0 To PILL_N - 1
        widths(i) = modChrome.PillWidth(caps(i), PILL_PITCH, PILL_PAD, PILL_MIN)
    Next i
End Sub

' ヘッダー上のピル。active:=Trueで「今いるモード」を塗りつぶして示す。
' yOff は段送り(1段に入りきらなかったぶんを下の段へ置くための縦オフセット)。
Private Sub Pill(ByVal ws As Worksheet, ByVal shapeName As String, _
                 ByVal caption As String, ByVal x As Double, ByVal w As Double, _
                 ByVal yOff As Double, ByVal action As String, ByVal active As Boolean)
    On Error Resume Next
    Dim p As Shape
    Set p = ws.Shapes.AddShape(5, x, (HDR_H - PILL_H) / 2 + yOff, w, PILL_H)
    If p Is Nothing Then Exit Sub
    p.Name = shapeName
    p.Placement = 3          ' 行高を後から変えてもピルは動かさない
    p.Adjustments(1) = 0.35
    p.Line.Visible = 0
    If active Then
        p.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        p.Fill.ForeColor.RGB = modUI.UiColor("sidebarActive")
    End If
    With p.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = -1
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    If active Then
        p.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modUI.UiColor("sidebar")
    Else
        p.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End If
    p.OnAction = action
    On Error GoTo 0
End Sub

Private Sub RemoveChrome(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.Count)
    Dim n As Long
    Dim shp As Shape
    For Each shp In ws.Shapes
        ' 旧マイ本棚のボタン(btn_/lbl_)も一緒に消す。残すと新ツールバーの上に
        ' 浮いたまま二重表示になる(Hub移植時に踏んだのと同じ罠)。
        ' 2026-07-30(R4要件A): 3モードが同じ「マイ本棚」シートを共有する
        ' ようになったので、前のモードのカード(nxg_)とチェックボックス(nxs_)も
        ' ここで必ず落とす。残すとギャラリーのカードの上に一覧表が重なる。
        If Left$(shp.Name, 4) = "nxk_" Or Left$(shp.Name, 4) = "btn_" _
           Or Left$(shp.Name, 4) = "lbl_" Or Left$(shp.Name, 4) = "nxg_" _
           Or Left$(shp.Name, 4) = "nxs_" Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        On Error Resume Next
        ws.Shapes(names(i)).Delete
        On Error GoTo 0
    Next i
End Sub

' ---- ツールバーのハンドラ(既存エンジンへの配線に徹する) ----
'
' 2026-07-31(R7 B-2): 画面遷移・資料操作の入口は必ず
' modUiLock.BlockIfIngesting を最初に通す。取込/同期の最中は
' 抽出ループのDoEventsでここが【取込の途中から入れ子で】走り出し、
' 実機の「取込中にボタンを押すとExcelが応答なし」になっていた。
' busy のときはモーダルを出さず、実況行だけ出して即Exitする。

Public Sub OnGoGallery()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    ' R14-2b: 既にギャラリー表示中の再押下は無反応に見えていた(押しても
    ' 見た目が変わらない=憲章§3-1違反)。押下前のモードを CurrentMode()
    ' (右肩ピルのactives()判定と同じ唯一の情報源)で控えておき、既に
    ' ギャラリーだった場合だけ再描画後にトーストで「効いた」ことを伝える。
    ' R14-G8: トーストを出すのは【描き直しに成功したとき】だけ。失敗しても
    ' 「表示を更新しました」を出していたため、ShowVaultGallery が出した
    ' 「描き直せませんでした」を自分で上書きし、壊れているのに成功に見えていた。
    Dim wasGallery As Boolean: wasGallery = (CurrentMode() = "gallery")
    On Error Resume Next
    Dim drawOk As Boolean
    drawOk = modVaultGallery.ShowVaultGallery()
    If wasGallery And drawOk Then modSkin.ShowToast "表示を更新しました", "info", True
    On Error GoTo 0
    modUiLock.Leave
End Sub

' みんなのQ&A(選択式取り込み)へ。
Public Sub OnGoShared()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    ' 2026-07-31(R7実装後の追加是正・発見事項4): Leave→Showの順で
    ' 順序が入れ替わっていたうえ、On Error Resume Nextの直後にGoTo 0が
    ' 無く、その後のExit Subまで丸ごと保護区間に入っていた。modShared.Show
    ' が失敗しても完全に無音になり、他の兄弟ハンドラ(OnGoGallery等)と
    ' 挙動が食い違っていた。
    On Error Resume Next
    modShared.Show
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnGoTable()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUIShelf.EnsureLayout
    modUI.GoToNativeSheet modAppDef.SH_SHELF, "modKnowledge.OnGoTable"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnBackHub()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modHub.EnsureHubLayout activate:=True
    On Error GoTo 0
    modUiLock.Leave
End Sub

' 「?」= 使い方。ヘルプカードはNexus(チャット)シート上に描かれるため、
' 先にチャットへ移ってから modHelp.OnHelpClick を呼ぶ(modHub.OnHelpと同型)。
' OnHelpClick 自身が modUiLock を取るので、ここではロックを取らない。
Public Sub OnHelp()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modKnowledge.OnHelp"
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnHelp", Err.Description, Err.Number
    On Error GoTo 0
    modHelp.OnHelpClick
End Sub

Public Sub OnToChat()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUI.GoToNexus "modKnowledge.OnToChat"
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnSearch()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modVaultGallery.OnVaultSearch
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnSearch", Err.Description, Err.Number
    On Error GoTo 0
    modUiLock.Leave
End Sub

' ----------------------------------------------------------------------------
' 💡 みんなの困りごと: 組織で答えが見つからなかった質問の一覧。
'   資料を書ける人(商品部)がここを見て、その場でナレッジを書けるようにする。
'   営業の「分からない」が、商品部の「書くべきこと」に直結する一番短い経路。
' ----------------------------------------------------------------------------
Public Sub OnGapBoard()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim body As String
    On Error Resume Next
    body = modInsight.GapListText()
    On Error GoTo Done

    Dim n As Long
    On Error Resume Next
    n = modInsight.GapCount()
    On Error GoTo Done

    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "みんなが質問して、本棚に答えが無かった質問です(新しい順・最大20件)。" & vbCrLf & _
        "ここに並ぶ質問に答える資料を用意すると、部内の全員がすぐ答えを得られます。" & vbCrLf & vbCrLf & _
        body & vbCrLf & _
        "この内容に答える資料を、今すぐ登録しますか?", _
        vbYesNo + vbInformation, modAppDef.APP_NAME & " - みんなの困りごと (" & n & "件)")

    If resp = vbYes Then
        modUiLock.Leave
        modVault.ShowVaultInput
        Exit Sub
    End If
Done:
    modUiLock.Leave
End Sub


' ----------------------------------------------------------------------------
' 部門チャンネル: 購読・更新・チャンク予算をひとまとめに扱う入口。
'   全社共通/商品/システム/人事… と部門ごとに正典が発行される。ここで
'   必要なものだけ購読する。全部入れないのが既定なので、部門が増えても
'   ひとりのブックが際限なく膨らむことはない。
' ----------------------------------------------------------------------------
Public Sub OnChannels()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Done

    Dim all As String
    On Error Resume Next
    all = modChannel.ListChannels()
    On Error GoTo Done

    If LenB(all) = 0 Then
        modUiLock.Leave
        ' 2026-08-06 R20H FA-16: 内部のフォルダ・ファイル構造(channels\<部門名>\
        ' pack.xlsx/version.txt)は運用担当者向けの情報で、一般利用者への
        ' 案内としては細かすぎたため簡略化(詳細はdocs/70データ管理者向け)。
        MsgBox "部門の公式ナレッジがまだ1つも見つかりません。" & vbCrLf & vbCrLf & _
            "各部門の担当者が共有フォルダへ発行すると、自動でここに現れます。" & vbCrLf & vbCrLf & _
            "共有フォルダ自体が未設定の場合は、Hubのお知らせから設定してください。", _
            vbInformation, modAppDef.APP_NAME
        Exit Sub
    End If

    ' 2026-07-27: 「どの部門か」を利用者に選ばせるのをやめた。
    '
    ' 以前はここで、アプリが既に知っている部門名を一覧表示したうえで、その
    ' どれかを利用者にInputBoxへ打ち直させていた。さらに「常時つないでおける
    ' のは1部門だけ」「切り替えると前の部門は本棚から外れます」と説明していた。
    ' これは modChannel 冒頭のコメントが「誤りだった」と名指ししている設計
    ' そのもの ―― 聞く前に分野を自分で判断させた時点で、ポータルを探し回る
    ' のと同じ認知負荷が生まれる。実務は「今日は商品、明日はシステム」であり、
    ' 分野の判定は人間ではなく検索側の仕事。
    Dim parts() As String: parts = Split(all, "|")
    Dim n As Long: n = UBound(parts) - LBound(parts) + 1

    Dim listText As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        listText = listText & "  ・" & parts(i) & vbCrLf
    Next i

    modUiLock.Leave
    If MsgBox("見つかった " & n & " 部門の公式ナレッジを、まとめて本棚に読み込みます。" & vbCrLf & vbCrLf & _
              listText & vbCrLf & _
              "以後はどの分野の質問でも、部門を選ばずにそのまま聞けます。" & vbCrLf & _
              "(あなたが自分で入れた資料はそのまま残ります)" & vbCrLf & vbCrLf & _
              "分量によっては数分かかることがあります。よろしいですか?", _
              vbOKCancel + vbQuestion, modAppDef.APP_NAME & " - 部門の公式ナレッジ") <> vbOK Then Exit Sub

    ' 2026-07-28(レビュー M-21): 実処理の区間はロックを取り直す。
    ' 確認ダイアログの前に Leave しているため、ここは無防備だった。
    ' 取り込みは共有I/Oのリトライ待ちで DoEvents を回すので、その間に
    ' もう一度押されると同じ処理が入れ子で走り、purge と import が
    ' 交錯して本棚が中途半端な状態になる。
    If Not modUiLock.Enter() Then Exit Sub

    Dim result As String
    On Error Resume Next
    ' 取り込み中の ESC を実行時エラー18として捕捉できるようにする
    ' (捕捉しないと Excel が処理を強制中断し、purge 直後で止まり得る)。
    Application.EnableCancelKey = 2      ' xlErrorHandler
    modUIMain.SetStage "" & ChrW(&HD83D) & ChrW(&HDCE1) & " 部門の公式ナレッジを読み込んでいます…"
    result = modChannel.SubscribeAllAvailable()
    modUIMain.SetStage ""
    Application.EnableCancelKey = 1      ' xlInterrupt(既定へ戻す)
    On Error GoTo 0
    modUiLock.Leave

    On Error Resume Next
    modHub.EnsureHubLayout
    On Error GoTo 0

    MsgBox result & vbCrLf & vbCrLf & _
           "本棚の使用量: " & modChannel.ChunkUsagePercent() & "%", _
           vbInformation, modAppDef.APP_NAME
    Exit Sub
Done:
    modUiLock.Leave
End Sub


Public Sub OnRegister()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modVault.ShowVaultInput
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnRegister", Err.Description, Err.Number
    On Error GoTo 0
    modUiLock.Leave
End Sub

Public Sub OnAddFiles()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modShelfBatch.AddFilesViaDialog
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnAddFiles", Err.Description, Err.Number
    On Error GoTo 0
    RefreshCurrent
    modUiLock.Leave
End Sub

' OnPackOut/OnPackIn: 例外ハンドラ必須(監査1 M-9・VBA生ダイアログ防止)。
Public Sub OnPackOut()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Fail
    modPackExport.ExportPackDialog
    modUiLock.Leave
    Exit Sub
Fail:
    modUiLock.Leave
    modLog.ShowError "E0801", "modKnowledge.OnPackOut", Err.Description
End Sub

Public Sub OnPackIn()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error GoTo Fail
    modPack.ImportPackDialog
    RefreshCurrent
    modUiLock.Leave
    Exit Sub
Fail:
    modUiLock.Leave
    modLog.ShowError "E0801", "modKnowledge.OnPackIn", Err.Description
End Sub

Public Sub OnSync()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modShelfSync.SyncNow
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnSync", Err.Description, Err.Number
    On Error GoTo 0
    RefreshCurrent
    modUiLock.Leave
End Sub

Public Sub OnPickFolder()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modShelfSync.PickShelfFolder
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnPickFolder", Err.Description, Err.Number
    On Error GoTo 0
    RefreshCurrent
    modUiLock.Leave
End Sub

Public Sub OnDelete()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    If Not modUiLock.Enter() Then Exit Sub
    On Error Resume Next
    modUIShelf.OnDeleteSource
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.OnDelete", Err.Description, Err.Number
    On Error GoTo 0
    RefreshCurrent
    modUiLock.Leave
End Sub

' 今表示しているモードだけを描き直す(モード切替をまたいで表示がズレないよう、
' 資料を足した/消した直後は必ずここを通す)。
'
' 2026-07-30(R4要件A/F): 3モードとも描画先が「マイ本棚」シート1枚になったので、
' 旧実装のように ActiveSheet.Name では現在モードを判別できない。DrawChromeが
' 記録した mMode を唯一の情報源にする。modApp.OnRefreshUI(🔄再描画)の
' フォールバック先でもあるため Public。
Public Sub RefreshCurrent()
    On Error Resume Next
    ' モードの解釈は CurrentMode() 1箇所に集約する(レビュー1-A)。
    ' ここで mMode を直接見ると、"" の扱いが IsTableMode とずれる。
    Select Case CurrentMode()
        Case "shared"
            modShared.Show
        Case "gallery"
            modVaultGallery.ShowVaultGallery
        Case Else
            modUIShelf.RenderShelf
    End Select
    If Err.Number <> 0 Then modLog.LogError "E0801", "modKnowledge.RefreshCurrent", Err.Description, Err.Number
    On Error GoTo 0
End Sub

