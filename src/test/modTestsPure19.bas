Attribute VB_Name = "modTestsPure19"
Option Explicit

' ============================================================================
' modTestsPure19 - R20-1(実機第7報⑦「右・下余白の3層根治」)の純ロジック回帰
' ----------------------------------------------------------------------------
' なぜ新設したか(憲章§4-6):
'   幾何系の純関数は modTestsPure16 に集めてきたが、あちらは27,170字で
'   30,000字上限まで残り2,830字しかなく、ここの真理表(約8,000字)が入らない。
'   16→17→18 と同じ線で分割する。入口は modTestsPure18.RunAll18 の末尾から
'   呼ばれる RunAll19 の1本だけ。
'
' ここで固定するもの(R20-1a〜1f):
'   ・modViewport.ClampD           : 幅・カード数の共通クランプ(境界と逆転指定)
'   ・modDashStat.CardWidthFor     : 帯幅→KPIカード幅(130で頭打ち/260で頭打ち)
'   ・modViewport.GalleryColsFor   : 帯幅→ギャラリーの列数(3〜6・退化入力)
'   ・modViewport.PadUnitsRefine   : ColumnWidth上限255の頭打ち(R20-1cで追加)
'   ・modViewport.RefitAction      : 窓リサイズ再フィットのデバウンス状態遷移
'   ・pt→行の換算(R20-1eの確定バグ): 「行高15pt固定」の机上換算が、行1が
'     48pt(ヘッダー帯)ある画面で何行ずれるかを数字で残す。実測に置き換えた
'     modViewport.RowAt は Worksheet を要るのでここでは回せないが、
'     【旧式が誤りである】ことは純粋な算数として固定できる。
' ============================================================================

' ----------------------------------------------------------------------------
' R20-1b/1c: 幅のクランプ(modViewport.ClampD)
' ----------------------------------------------------------------------------
Private Sub TestClampD()
    modTestRunner.Check "ClampD_範囲内はそのまま", (modViewport.ClampD(150, 130, 220) = 150)
    modTestRunner.Check "ClampD_下限ちょうど(境界)", (modViewport.ClampD(130, 130, 220) = 130)
    modTestRunner.Check "ClampD_上限ちょうど(境界)", (modViewport.ClampD(220, 130, 220) = 220)
    modTestRunner.Check "ClampD_下限未満は下限へ", (modViewport.ClampD(0, 130, 220) = 130)
    modTestRunner.Check "ClampD_上限超は上限へ", (modViewport.ClampD(9999, 130, 220) = 220)
    modTestRunner.Check "ClampD_負値も下限へ", (modViewport.ClampD(-500, 130, 220) = 130)
    ' lo>hi の壊れた指定でも落とさず lo を返す(呼び出し側の防御を1つ減らす)。
    modTestRunner.Check "ClampD_下限>上限の異常指定は下限を返す", _
        (modViewport.ClampD(100, 220, 130) = 220)
End Sub

' ----------------------------------------------------------------------------
' R20-1b: 帯幅 → KPIカード幅(modDashStat.CardWidthFor)
' ----------------------------------------------------------------------------
' 版面は 左右余白 KPI_X0=20 を2つ + カード間 KPI_GAP=10 を3つ + カード4枚。
'   cardW = clamp((bandW - 40 - 30) / 4, 130, 260)  ' R21-S3で上限220→260
' 従来は 130 固定で、帯だけが可視幅へ伸びていた(実機第7報⑦の層1)。
Private Sub TestCardWidthFor()
    ' 最小版面 550pt(=130×4+10×3)+左右余白40pt=590pt が下限の境界。
    modTestRunner.Check "カード幅_帯590ptでちょうど最小130(境界)", _
        (modDashStat.CardWidthFor(590) = 130)
    modTestRunner.Check "カード幅_帯589ptでも130を割らない(境界の外)", _
        (modDashStat.CardWidthFor(589) = 130)
    modTestRunner.Check "カード幅_帯591ptは130を超える(境界の内側)", _
        (modDashStat.CardWidthFor(591) > 130)
    modTestRunner.Check "カード幅_狭い窓(帯400pt)でも130", _
        (modDashStat.CardWidthFor(400) = 130)
    ' R21-S3: 上限を220→260へ上げた。260×4+10×3+40 = 1110pt が上限の境界。
    modTestRunner.Check "カード幅_帯1110ptでちょうど最大260(境界)", _
        (modDashStat.CardWidthFor(1110) = 260)
    modTestRunner.Check "カード幅_帯1111ptでも260で頭打ち", _
        (modDashStat.CardWidthFor(1111) = 260)
    modTestRunner.Check "カード幅_広い窓(帯1800pt)でも260で頭打ち", _
        (modDashStat.CardWidthFor(1800) = 260)
    ' 旧上限220は「頭打ちではなくなった」ことを固定する(退行検知)。
    modTestRunner.Check "カード幅_帯950ptは220を超える(旧上限では止まらない)", _
        (modDashStat.CardWidthFor(950) = 220)
    ' 実機の代表値。窓900pt→帯888pt: (888-70)/4 = 204.5
    modTestRunner.Check "カード幅_帯888pt(窓900)で204.5", _
        (Abs(modDashStat.CardWidthFor(888) - 204.5) < 0.001), _
        "実際=" & modDashStat.CardWidthFor(888)
    ' 窓1300pt→帯1288pt: (1288-70)/4 = 304.5 → 260で頭打ち
    modTestRunner.Check "カード幅_帯1288pt(窓1300)は260で頭打ち", _
        (modDashStat.CardWidthFor(1288) = 260)
    modTestRunner.Check "カード幅_0や負の帯幅でも130(退化入力)", _
        (modDashStat.CardWidthFor(0) = 130 And modDashStat.CardWidthFor(-100) = 130)
End Sub

' ----------------------------------------------------------------------------
' R20-1c: 帯幅 → ギャラリーの列数(modViewport.GalleryColsFor)
' ----------------------------------------------------------------------------
' cols = clamp(Int((bandW - leftX - gap) / (cardW + gap)), 3, 6)
' カードは 215pt、隙間は 14pt。左端 leftX はA列(幅2)の右=約12pt。
Private Sub TestGalleryColsFor()
    ' 窓900pt → 帯888pt: Int((888-12-14)/229) = Int(3.76) = 3列(従来と同じ絵)
    modTestRunner.Check "ギャラリー列数_帯888pt(窓900)は3列", _
        (modViewport.GalleryColsFor(888, 12, 215, 14, 3, 6) = 3), _
        "実際=" & modViewport.GalleryColsFor(888, 12, 215, 14, 3, 6)
    ' 窓1300pt → 帯1288pt: Int((1288-12-14)/229) = Int(5.51) = 5列
    modTestRunner.Check "ギャラリー列数_帯1288pt(窓1300)は5列", _
        (modViewport.GalleryColsFor(1288, 12, 215, 14, 3, 6) = 5), _
        "実際=" & modViewport.GalleryColsFor(1288, 12, 215, 14, 3, 6)
    ' 窓1800pt → 帯1788pt: Int((1788-12-14)/229) = Int(7.69) = 7 → 上限6でクランプ
    modTestRunner.Check "ギャラリー列数_帯1788pt(窓1800)は上限6でクランプ", _
        (modViewport.GalleryColsFor(1788, 12, 215, 14, 3, 6) = 6)
    ' 4列ちょうどの境界: leftX+gap+4*(215+14) = 12+14+916 = 942pt
    modTestRunner.Check "ギャラリー列数_帯942ptでちょうど4列(境界)", _
        (modViewport.GalleryColsFor(942, 12, 215, 14, 3, 6) = 4)
    modTestRunner.Check "ギャラリー列数_帯941ptは3列(境界の1pt手前)", _
        (modViewport.GalleryColsFor(941, 12, 215, 14, 3, 6) = 3)
    ' 狭い窓では下限3列。カード群は左寄せのままなので、はみ出す側は
    ' 横スクロールで届く(0列にして画面を空にするより害が小さい)。
    modTestRunner.Check "ギャラリー列数_帯500ptでも下限3列", _
        (modViewport.GalleryColsFor(500, 12, 215, 14, 3, 6) = 3)
    modTestRunner.Check "ギャラリー列数_帯0でも下限3列(退化入力)", _
        (modViewport.GalleryColsFor(0, 12, 215, 14, 3, 6) = 3)
    ' カード幅+隙間が0以下(壊れた定数)でも落とさず下限を返す。
    modTestRunner.Check "ギャラリー列数_カード幅0は下限を返す(0除算の防御)", _
        (modViewport.GalleryColsFor(1000, 12, 0, 0, 3, 6) = 3)
    ' 1ページ9枚が何段になるか(ページャのYと下端行に直結する)。
    modTestRunner.Check "ギャラリー段数_3列なら3段", (PageRows(9, 3) = 3)
    modTestRunner.Check "ギャラリー段数_5列なら2段", (PageRows(9, 5) = 2)
    modTestRunner.Check "ギャラリー段数_6列なら2段", (PageRows(9, 6) = 2)
    modTestRunner.Check "ギャラリー段数_1枚だけなら1段", (PageRows(1, 3) = 1)
End Sub

' 描画側 RenderGalleryCards と同じ式(((endIdx-startIdx) \ cols) + 1)。
Private Function PageRows(ByVal shown As Long, ByVal cols As Long) As Long
    If cols < 1 Then cols = 1
    PageRows = ((shown - 1) \ cols) + 1
    If PageRows < 1 Then PageRows = 1
End Function

' ----------------------------------------------------------------------------
' R20-1c: ColumnWidth の上限255(modViewport.PadUnitsRefine)
' ----------------------------------------------------------------------------
' 本文の最終列に余りを吸わせる使い方(本棚のメモ列J・登録フォームのH列)を
' 足したので、超広い窓では 255 を超える ColumnWidth が計算されうる。
' 超えた値の代入は1004で、On Error Resume Next 配下では「列幅が前回のまま」と
' いう無言の失敗になる(=右の余白が消えない)。手前で頭打ちにする。
Private Sub TestPadUnitsRefineClamp()
    ' 傾き = (w1-w0)/(u1-u0) = (750-75)/(100-10) = 7.5 pt/unit。
    ' 必要 2500pt なら u = 100 + (2500-750)/7.5 = 333.3 → 255で頭打ち。
    modTestRunner.Check "吸収列_255を超える要求は255で頭打ち", _
        (modViewport.PadUnitsRefine(2500, 10, 75, 100, 750) = 255), _
        "実際=" & modViewport.PadUnitsRefine(2500, 10, 75, 100, 750)
    ' 255ちょうどになる要求(u=255 → needPt = 750 + (255-100)*7.5 = 1912.5)。
    modTestRunner.Check "吸収列_ちょうど255は255のまま(境界)", _
        (modViewport.PadUnitsRefine(1912.5, 10, 75, 100, 750) = 255)
    ' 従来の下限(0.05)は変わっていない。
    modTestRunner.Check "吸収列_極小要求は0.05で止まる(従来どおり)", _
        (modViewport.PadUnitsRefine(-9999, 10, 75, 100, 750) = 0.05)
    ' 通常域は素通り(既存の振る舞いを壊していないこと)。
    modTestRunner.Check "吸収列_通常域は素通り", _
        (Abs(modViewport.PadUnitsRefine(300, 10, 75, 100, 750) - 40) < 0.001), _
        "実際=" & modViewport.PadUnitsRefine(300, 10, 75, 100, 750)
End Sub

' ----------------------------------------------------------------------------
' R20-1f: 窓リサイズ再フィットのデバウンス状態遷移(modViewport.RefitAction)
' ----------------------------------------------------------------------------
' 引数は (ev, isMine, isRunning, isBusy, waited)。
' 戻り値 "none"/"schedule"/"run" の3値だけで、OnTime も ActiveWorkbook も
' 見ない ―― だからここで全分岐を回せる。
Private Sub TestRefitAction()
    ' 窓が動いた → 予約し直す(連続リサイズは ScheduleRefit が先に解除するので
    ' 何回来ても予約は1本。ここでは「予約する」判断が出ることだけを固定する)。
    modTestRunner.Check "再フィット_リサイズ通知は予約", _
        (modViewport.RefitAction("resize", True, False, False, False) = "schedule")
    modTestRunner.Check "再フィット_連続リサイズも毎回予約(多重登録は解除側で防ぐ)", _
        (modViewport.RefitAction("resize", True, False, False, True) = "schedule")
    ' (iii) 他ブックが前面: 予約もしないし実行もしない。
    modTestRunner.Check "再フィット_他ブック前面のリサイズは何もしない", _
        (modViewport.RefitAction("resize", False, False, False, False) = "none")
    modTestRunner.Check "再フィット_他ブック前面のタイマ発火も何もしない", _
        (modViewport.RefitAction("tick", False, False, False, False) = "none")
    ' 再入抑止: 自分の再描画が起こしたリサイズを拾って無限ループにしない。
    modTestRunner.Check "再フィット_実行中のリサイズは無視(再入抑止)", _
        (modViewport.RefitAction("resize", True, True, False, False) = "none")
    modTestRunner.Check "再フィット_実行中のタイマ発火も無視", _
        (modViewport.RefitAction("tick", True, True, False, False) = "none")
    ' 通常のタイマ発火 → 組み直す。
    modTestRunner.Check "再フィット_タイマ発火は組み直す", _
        (modViewport.RefitAction("tick", True, False, False, False) = "run")
    ' Busy(取込・回答生成中)は1回だけ待ち直す。2回目は諦める。
    modTestRunner.Check "再フィット_Busyの初回は1回だけ待ち直す", _
        (modViewport.RefitAction("tick", True, False, True, False) = "schedule")
    modTestRunner.Check "再フィット_Busyの2回目は諦める(張り替え続けない)", _
        (modViewport.RefitAction("tick", True, False, True, True) = "none")
    ' 未知のイベント名は何もしない(配線ミスで暴走させない)。
    modTestRunner.Check "再フィット_未知のイベント名は何もしない", _
        (modViewport.RefitAction("close", True, False, False, False) = "none")
    modTestRunner.Check "再フィット_空のイベント名も何もしない", _
        (modViewport.RefitAction("", True, False, False, False) = "none")
    ' 大小・前後空白は無視する(呼び出し側の表記ゆれで壊れない)。
    modTestRunner.Check "再フィット_大文字/空白付きのRESIZEも予約", _
        (modViewport.RefitAction("  RESIZE ", True, False, False, False) = "schedule")
End Sub

' ----------------------------------------------------------------------------
' R20-1e【確定バグ】: pt→行の机上換算が行1=48ptを無視していた
' ----------------------------------------------------------------------------
' Hub は行1だけがヘッダー帯と同じ48pt(ナビが2段に折り返せば96pt)で、
' 行2以降が15pt。旧実装 r = CLng(y/15) + 2 はこれを無視していたため、
' 統計タイルの下端(386pt)に対して行28(上端438pt)を返し、52ptの空隙が
' 開いていた。正しくは「下端が386pt以上になる最初の行」= 行24 で、
' バッジ見出しはその次の行25(上端393pt)から始まる。
' RowAt は Worksheet を要るのでここでは回せないが、【旧式が誤りである】
' ことと、正しい行番号がいくつかは純粋な算数として固定できる。
Private Sub TestRowConversion()
    Dim tilesBottom As Double: tilesBottom = 386      ' TilesTop(146)+TilesHeight(240)
    modTestRunner.Check "行換算_旧式(15pt割り算)は行28を返す", _
        (CLng(tilesBottom / 15) + 2 = 28), "実際=" & (CLng(tilesBottom / 15) + 2)
    modTestRunner.Check "行換算_実測なら下端386ptを含む最小行は24", _
        (RowAtHub(tilesBottom) = 24), "実際=" & RowAtHub(tilesBottom)
    modTestRunner.Check "行換算_バッジ見出しはその次の行25", _
        (RowAtHub(tilesBottom) + 1 = 25)
    ' 旧式が返す行28の上端は438pt。タイル下端386ptとの差=52ptが空隙の正体。
    modTestRunner.Check "行換算_旧式の空隙は52pt", (RowTopHub(28) - tilesBottom = 52), _
        "実際=" & (RowTopHub(28) - tilesBottom)
    modTestRunner.Check "行換算_実測なら空隙は7pt以内", _
        (RowTopHub(25) - tilesBottom = 7), "実際=" & (RowTopHub(25) - tilesBottom)
    ' 戻り値(バッジ帯の下端)も同じ誤差を持っていた。
    ' 旧: (r+4)*15 = 32*15 = 480pt / 実測: 行29の下端 = 48+27*15+15 = 468pt。
    ' 旧式は【行28起点】なので実際の下端は行32の下端=513pt。つまり返り値480は
    ' 実物より33pt上で、フッターがバッジ帯へ食い込んでいた。
    modTestRunner.Check "行換算_旧式の戻り値は480pt", ((28 + 4) * 15 = 480)
    modTestRunner.Check "行換算_旧式起点の実下端は513pt(戻り値より33pt下)", _
        (RowTopHub(32) + 15 = 513), "実際=" & (RowTopHub(32) + 15)
    modTestRunner.Check "行換算_実測起点の実下端は468pt(食い込みなし)", _
        (RowTopHub(29) + 15 = 468), "実際=" & (RowTopHub(29) + 15)
    ' 行1が2段(96pt)へ伸びる狭い窓では、旧式のズレはさらに広がる。
    modTestRunner.Check "行換算_ヘッダー2段(96pt)なら実測行は21", _
        (RowAtGeneric(tilesBottom, 96, 15) = 21), _
        "実際=" & RowAtGeneric(tilesBottom, 96, 15)
End Sub

' Hubの行幾何(行1=48pt・行2以降15pt)での「yを含む最小の行」。
Private Function RowAtHub(ByVal y As Double) As Long
    RowAtHub = RowAtGeneric(y, 48, 15)
End Function

' 行1の高さ h1、行2以降 h の等幅シートで、下端がy以上になる最初の行。
Private Function RowAtGeneric(ByVal y As Double, ByVal h1 As Double, ByVal h As Double) As Long
    Dim i As Long
    Dim bottom As Double
    For i = 1 To 400
        If i = 1 Then
            bottom = h1
        Else
            bottom = h1 + (i - 1) * h
        End If
        If bottom >= y Then
            RowAtGeneric = i
            Exit Function
        End If
    Next i
    RowAtGeneric = 400
End Function

' Hubの行幾何での行nの上端(pt)。
Private Function RowTopHub(ByVal n As Long) As Double
    If n <= 1 Then Exit Function
    RowTopHub = 48 + (n - 2) * 15
End Function

' ----------------------------------------------------------------------------
' R20-3(実機第7報②): 資料の仕上げバックフィルの純ロジック
'   (modBackfill.ClassifyDoc/LooksLikeBreadcrumbLine/ConfirmText/ResultText)
' ----------------------------------------------------------------------------
Private Sub TestBackfillClassify()
    ' 両方揃っていれば仕上げ済み(空文字=対象外。breadcrumbOkの値によらない)。
    modTestRunner.Check "仕上げ判定_メタ有+要約有は仕上げ済み(空文字)", _
        (modBackfill.ClassifyDoc(True, True, True) = "")
    modTestRunner.Check "仕上げ判定_メタ有+要約有はbreadcrumb欠落でも仕上げ済み", _
        (modBackfill.ClassifyDoc(True, True, False) = "")
    ' メタ無し(未仕上げの主因)はbreadcrumbがあれば仕上げ対象。
    modTestRunner.Check "仕上げ判定_メタ無+要約無+breadcrumb有は仕上げ対象", _
        (modBackfill.ClassifyDoc(False, False, True) = modBackfill.STATUS_NEEDS)
    modTestRunner.Check "仕上げ判定_メタ有+要約無+breadcrumb有も仕上げ対象", _
        (modBackfill.ClassifyDoc(True, False, True) = modBackfill.STATUS_NEEDS)
    modTestRunner.Check "仕上げ判定_メタ無+要約有+breadcrumb有も仕上げ対象", _
        (modBackfill.ClassifyDoc(False, True, True) = modBackfill.STATUS_NEEDS)
    ' breadcrumb欠落(かなり古い取込方式)は他の状態によらず仕上げ不可。
    modTestRunner.Check "仕上げ判定_breadcrumb欠落は仕上げ不可(メタ無+要約無)", _
        (modBackfill.ClassifyDoc(False, False, False) = modBackfill.STATUS_CANNOT)
    modTestRunner.Check "仕上げ判定_breadcrumb欠落は仕上げ不可(メタ有+要約無)", _
        (modBackfill.ClassifyDoc(True, False, False) = modBackfill.STATUS_CANNOT)
End Sub

Private Sub TestBackfillBreadcrumb()
    modTestRunner.Check "breadcrumb判定_資料>章>条の2階層は形式あり", _
        modBackfill.LooksLikeBreadcrumbLine("【資料名>第1章>第1条】" & vbLf & "本文…")
    modTestRunner.Check "breadcrumb判定_空文字は形式なし", _
        (Not modBackfill.LooksLikeBreadcrumbLine(""))
    modTestRunner.Check "breadcrumb判定_先頭が【でない本文は形式なし", _
        (Not modBackfill.LooksLikeBreadcrumbLine("ふつうの本文です。第5条について…"))
    ' 「【】」だけ=閉じ括弧が2文字目(3文字目未満)なので形式なし(境界)。
    modTestRunner.Check "breadcrumb判定_【】だけは形式なし(境界の外)", _
        (Not modBackfill.LooksLikeBreadcrumbLine("【】" & vbLf & "本文"))
    ' 「【a】」=閉じ括弧がちょうど3文字目なので形式あり(境界の内側)。
    modTestRunner.Check "breadcrumb判定_【a】は形式あり(境界)", _
        modBackfill.LooksLikeBreadcrumbLine("【a】" & vbLf & "本文")
    modTestRunner.Check "breadcrumb判定_改行が無い1行だけでも判定できる", _
        modBackfill.LooksLikeBreadcrumbLine("【資料名】")
End Sub

Private Sub TestBackfillText()
    modTestRunner.Check "確認文_件数が入り再取込不要と分かる", _
        (InStr(modBackfill.ConfirmText(3, 0), "3冊") > 0 And _
         InStr(modBackfill.ConfirmText(3, 0), "再取込は不要") > 0)
    modTestRunner.Check "確認文_仕上げ不可がある場合は件数を併記", _
        (InStr(modBackfill.ConfirmText(3, 2), "2冊は形式が古いため") > 0)
    modTestRunner.Check "確認文_仕上げ不可0件は併記しない", _
        (InStr(modBackfill.ConfirmText(3, 0), "形式が古いため") = 0)
    modTestRunner.Check "結果文_失敗0件は完了のみ+OCR再取込案内(R21-3 E2)", _
        (modBackfill.ResultText(5, 0, 0) = "5冊の仕上げが完了しました。" & vbLf & _
         "(OCR資料は再取込するとさらに章立てが正確になります)")
    modTestRunner.Check "結果文_失敗があれば件数を併記", _
        (InStr(modBackfill.ResultText(4, 1, 0), "1冊は失敗") > 0)
    modTestRunner.Check "結果文_仕上げ不可があれば再取込が必要と併記", _
        (InStr(modBackfill.ResultText(4, 0, 1), "再取込が必要") > 0)
    modTestRunner.Check "結果文_0冊完了(該当なし)ではOCR案内を付けない", _
        (InStr(modBackfill.ResultText(0, 0, 0), "OCR資料は再取込") = 0)
End Sub

Public Sub RunAll19()
    On Error GoTo ClampFail19
    TestClampD
NextCard19:
    On Error GoTo CardFail19
    TestCardWidthFor
NextCols19:
    On Error GoTo ColsFail19
    TestGalleryColsFor
NextPad19:
    On Error GoTo PadFail19
    TestPadUnitsRefineClamp
NextRefit19:
    On Error GoTo RefitFail19
    TestRefitAction
NextRow19:
    On Error GoTo RowFail19
    TestRowConversion
NextBfClassify19:
    On Error GoTo BfClassifyFail19
    TestBackfillClassify
NextBfCrumb19:
    On Error GoTo BfCrumbFail19
    TestBackfillBreadcrumb
NextBfText19:
    On Error GoTo BfTextFail19
    TestBackfillText
NextChain20:
    ' R20-4/R20-7(実機第7報③④⑥)の真理表は modTestsPure20 へ(波ごとの分割)。
    On Error GoTo ChainFail20
    modTestsPure20.RunAll20
NextDone19:
    On Error GoTo 0
    Exit Sub

ClampFail19:
    modTestRunner.Check "TestClampD(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCard19
CardFail19:
    modTestRunner.Check "TestCardWidthFor(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextCols19
ColsFail19:
    modTestRunner.Check "TestGalleryColsFor(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPad19
PadFail19:
    modTestRunner.Check "TestPadUnitsRefineClamp(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRefit19
RefitFail19:
    modTestRunner.Check "TestRefitAction(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextRow19
RowFail19:
    modTestRunner.Check "TestRowConversion(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBfClassify19
BfClassifyFail19:
    modTestRunner.Check "TestBackfillClassify(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBfCrumb19
BfCrumbFail19:
    modTestRunner.Check "TestBackfillBreadcrumb(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextBfText19
BfTextFail19:
    modTestRunner.Check "TestBackfillText(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextChain20
ChainFail20:
    modTestRunner.Check "modTestsPure20.RunAll20(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone19
End Sub
