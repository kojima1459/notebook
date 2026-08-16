Attribute VB_Name = "modTestsPure36"
Option Explicit

' ============================================================================
' modTestsPure36 - R33H Fix波3(テスト・文言・記録)の純ロジック回帰テスト。
'   modTestsPure34 が残45字になったための分割先。既存チェーン
'   (modTestsPure.RunAll→…→modTestsPure30.RunAll30)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll36 の1本が入口
'   (modTestsPure31 / 32 / 33 / 34 / 35 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【F22: 辞書が使えない環境で、称号行の解析をどう固定したか】
'   裁定書は「BoardReadRows に実 Dictionary を渡す」と書いていたが、
'   **その方式は取れない**。LibreOffice には Scripting.Dictionary が無く、
'   CreateObject した時点で Err=323 になって群ごと落ちる(Fix波2の実測)。
'   そこで採った方式:
'     行の【読み分け】だけを辞書に触らない純関数の対
'     (modUtilText.BoardRowText / BoardRowParse)へ出し、
'     書く側(modShare.BoardBodyText)と読む側(modShare.BoardReadRows)の
'     両方がその1対を通るようにしたうえで、往復をゴールデンで固定する。
'
'   ★この方式で守れるもの:
'     ・称号行("T")の【キーの正規化】(Trim + LCase)。ここがズレると
'       同じ人の称号が誰の画面にも出ない(読む側は modP2PIo.IdKey で
'       小文字化して引く)。従来は1文字も撃たれていなかった。
'     ・行の【採否】(列不足・未知の種別・空行を落とすこと)。
'     ・部の行("D")と称号行("T")の取り違えが起きないこと。
'     ・書く側と読む側の【書式の一致】(往復して元へ戻ること)。
'       片方だけ書式を変えた瞬間に落ちる。
'     ・BoardBodyText が終端行まで含めて組めること / BoardHeadLine が
'       CRLF・LF どちらでもヘッダ1行を取り出すこと / BoardStateText が
'       状態ごとに違う説明を返すこと(F22 が「一度も呼ばれていない」と
'       指摘した3本すべて)。
'
'   ★この方式でも守れないもの(正直に書く。恒真アサートで埋めない):
'     ・`titlesOut(key) = value` という【Dictionary への書き込みそのもの】。
'       LO では実辞書を作れないため、この1行だけは実機スモーク担当のまま。
'       症状が出るとすれば「称号が誰にも出ない」で、実機第18報の確認項目
'       (Hub の称号表示)で拾う。
'     ・BoardBodyText が deptAgg.Keys / titleAgg.Keys を回す部分。同上。
'     ・共有フォルダI/O(BoardWriteSummary / BoardCarryTitles)。
'
' 【このモジュールが守るもの(F22 以外)】
'   F31 modUtilText.AutoNeedsRetry / AutoTextPick: 文字コードの採否判断。
'   F27 modShelfScan.HeaderCheckable: 先頭バイトを読む価値がある拡張子か。
'   F30 modP2PIo.IdKey / IdHash: 送受信で同じキー・同じハッシュになること。
' ============================================================================

' ---- F22(1): 1行を組む/読む の往復 ---------------------------------------
'   discriminate の作り:
'   ・Parse から LCase を外すと「称号キーは小文字へ揃える」が落ちる。
'   ・Parse から Trim を外すと「前後の空白を落とす」2本が落ちる。
'   ・部の行にも LCase を掛けると「部の行は大小をそのまま残す」が落ちる
'     (部門名は表示にも使うので潰してはいけない)。
'   ・列不足/未知種別のガードを外すと空文字を期待する3本が落ちる。
'   ・Text 側の区切りをタブ以外へ変えると往復2本が落ちる。
Private Sub TestBoardRow36()
    ChkStr36 "F22_組む: 種別/キー/数値をタブで繋ぐ", _
        modUtilText.BoardRowText("D", "営業部", "300"), "D" & vbTab & "営業部" & vbTab & "300"

    ChkStr36 "F22_読む: 部の行はそのまま", _
        modUtilText.BoardRowParse("D" & vbTab & "営業部" & vbTab & "300"), _
        "D" & vbTab & "営業部" & vbTab & "300"
    ChkStr36 "F22_読む: 称号キーは小文字へ揃える", _
        modUtilText.BoardRowParse("T" & vbTab & "Yamada_Taro" & vbTab & "7"), _
        "T" & vbTab & "yamada_taro" & vbTab & "7"
    ChkStr36 "F22_読む: 部の行は大小をそのまま残す", _
        modUtilText.BoardRowParse("D" & vbTab & "Sales" & vbTab & "12"), _
        "D" & vbTab & "Sales" & vbTab & "12"
    ChkStr36 "F22_読む: 称号キーの前後の空白は落とす", _
        modUtilText.BoardRowParse("T" & vbTab & "  U1  " & vbTab & "5"), _
        "T" & vbTab & "u1" & vbTab & "5"
    ChkStr36 "F22_読む: 部名の前後の空白も落とす", _
        modUtilText.BoardRowParse("D" & vbTab & " 総務部 " & vbTab & "9"), _
        "D" & vbTab & "総務部" & vbTab & "9"

    ChkStr36 "F22_読む: 列が足りない行は捨てる", _
        modUtilText.BoardRowParse("T" & vbTab & "u1"), ""
    ChkStr36 "F22_読む: 知らない種別は捨てる", _
        modUtilText.BoardRowParse("X" & vbTab & "u1" & vbTab & "5"), ""
    ChkStr36 "F22_読む: 空行は捨てる", modUtilText.BoardRowParse(""), ""
    ChkStr36 "F22_読む: 終端行は行として読まない", _
        modUtilText.BoardRowParse(modShare.BOARD_END_TAG & vbTab & "2"), ""

    ' 往復: 組んだものを読むと、キーの正規化ぶんを除いて元へ戻る。
    ChkStr36 "F22_往復: 部の行は組んで読むと同じ", _
        modUtilText.BoardRowParse(modUtilText.BoardRowText("D", "営業部", "300")), _
        modUtilText.BoardRowText("D", "営業部", "300")
    ChkStr36 "F22_往復: 称号行は小文字化ぶんだけ変わる", _
        modUtilText.BoardRowParse(modUtilText.BoardRowText("T", "U9", "20")), _
        modUtilText.BoardRowText("T", "u9", "20")
End Sub

' ---- F22(2): 本文を通した称号行の経路 -------------------------------------
'   BoardReadRows は辞書を渡さない(Nothing)ままでも、称号行を【読み飛ばさず
'   に解析して】部の行と取り違えないことが確かめられる。
'   discriminate:
'   ・"T" と "D" の分岐を入れ替えると「称号行を部の合計に数えない」が落ちる。
'   ・myDept の照合を前方一致等へ緩めると「別の部は0」が落ちる。
'   ・BoardNum のガードを外すと負値・巨大値の2本が落ちる。
'   ・キーの Trim を外すと「前後に空白のある部名でも一致する」が落ちる。
Private Sub TestBoardReadRows36()
    Dim h As String
    h = modShare.BoardHeadText("2026-08-16 09:00:00", "20260816", 120, _
        "202608", 900, "2026", 5000, 42, False)
    Dim b As String
    b = h & vbLf & modUtilText.BoardRowText("T", "u1", "7") & _
            vbLf & modUtilText.BoardRowText("D", "営業部", "300") & _
            vbLf & modUtilText.BoardRowText("T", "u2", "20") & _
            vbLf & modUtilText.BoardRowText("D", " 総務部 ", "150") & _
            vbLf & modShare.BOARD_END_TAG & vbTab & "4"

    Dim td As Object      ' 実辞書は LO で作れない(Err=323)ので Nothing のまま
    ChkLong36 "F22_自部署の今月合計だけ返す", modShare.BoardReadRows(b, "営業部", td), 300
    ChkLong36 "F22_前後に空白のある部名でも一致する", modShare.BoardReadRows(b, "総務部", td), 150
    ChkLong36 "F22_居ない部は0", modShare.BoardReadRows(b, "経理部", td), 0
    ChkLong36 "F22_部名が空なら0(部の行を出さない)", modShare.BoardReadRows(b, "", td), 0
    ' 称号キーを部名として渡しても部の合計にはならない(種別の取り違え検知)。
    ChkLong36 "F22_称号行を部の合計に数えない", modShare.BoardReadRows(b, "u2", td), 0
    ' 終端行まで揃っていること(データ行4件)。
    ChkLong36 "F22_本文は終端行まで揃っている", modShare.BoardEndCount(b), 4

    ' 壊れた数値欄(共有フォルダのファイルは誰でも書ける前提)。
    Dim bad As String
    bad = h & vbLf & modUtilText.BoardRowText("D", "営業部", "-5") & _
              vbLf & modShare.BOARD_END_TAG & vbTab & "1"
    ChkLong36 "F22_負の値は0として捨てる", modShare.BoardReadRows(bad, "営業部", td), 0
    Dim huge As String
    huge = h & vbLf & modUtilText.BoardRowText("D", "営業部", "999999999") & _
               vbLf & modShare.BOARD_END_TAG & vbTab & "1"
    ChkLong36 "F22_1億分超も0として捨てる", modShare.BoardReadRows(huge, "営業部", td), 0
End Sub

' ---- F22(3): 一度も呼ばれていなかった3本 ---------------------------------
'   BoardBodyText / BoardHeadLine / BoardStateText。
'   discriminate:
'   ・BodyText の終端行を落とすと BoardEndCount が -1 になり2本落ちる。
'   ・HeadLine が改行を見なくなると「2行目以降を混ぜない」が落ちる。
'   ・StateText が状態を見ずに1つの文言を返すと、4件が同じ文字列になり
'     「状態ごとに違う説明」の3本が落ちる。
Private Sub TestBoardTexts36()
    Dim h As String
    h = modShare.BoardHeadText("2026-08-16 09:00:00", "20260816", 120, _
        "202608", 900, "2026", 5000, 42, False)

    ' BoardBodyText: 辞書が無い(Nothing)ときはヘッダ+終端行だけの完全な本文。
    Dim body As String: body = modShare.BoardBodyText(h, Nothing, Nothing)
    ChkStr36 "F22_本文の1行目はヘッダそのもの", modShare.BoardHeadLine(body), h
    ChkLong36 "F22_データ0件でも終端行があるので完全", modShare.BoardEndCount(body), 0
    ChkStr36 "F22_入口は完全で新しければok", _
        modShare.BoardTextStatus(body, "2026-08-15 09:00:00"), "ok"

    ' BoardHeadLine: CRLF でも LF でも1行目だけ。
    ChkStr36 "F22_ヘッダ抽出はLFで切る", _
        modShare.BoardHeadLine(h & vbLf & "D" & vbTab & "営業部" & vbTab & "300"), h
    ChkStr36 "F22_ヘッダ抽出はCRLFでも切る", _
        modShare.BoardHeadLine(h & vbCrLf & "D" & vbTab & "営業部" & vbTab & "300"), h
    ChkStr36 "F22_改行の無い1行はそのまま", modShare.BoardHeadLine(h), h

    ' BoardStateText: 状態ごとに違う説明。人が読む文なので「何が入るか」で見る。
    Dim sStale As String: sStale = modShare.BoardStateText("stale", 24)
    Dim sBroken As String: sBroken = modShare.BoardStateText("broken", 24)
    Dim sNone As String: sNone = modShare.BoardStateText("none", 24)
    ChkBool36 "F22_古い集計は時間数を添えて言う", (InStr(sStale, "24") > 0), True
    ChkBool36 "F22_壊れた集計は書き込み中の可能性を言う", _
        (InStr(sBroken, "書き込み中") > 0), True
    ChkBool36 "F22_集計無しは「まだありません」と言う", _
        (InStr(sNone, "まだありません") > 0), True
    ChkBool36 "F22_3つの状態は互いに違う文言", _
        (StrComp(sStale, sBroken, vbBinaryCompare) <> 0 And _
         StrComp(sBroken, sNone, vbBinaryCompare) <> 0 And _
         StrComp(sStale, sNone, vbBinaryCompare) <> 0), True
    ' 知らない状態は「まだありません」側へ倒す(Case Else)。ここを「ok と
    ' 同じ扱い」にすると、読めなかった集計が数字として出るようになる。
    ChkStr36 "F22_知らない状態はまだありません側へ倒す", _
        modShare.BoardStateText("zzz", 24), sNone
End Sub

' ---- F31: 文字コードの採否判断 --------------------------------------------
'   閾値は2%(FFFD_LIMIT)。境界の意味は modUtilText の見出しを参照。
'   discriminate:
'   ・`>` を `>=` にすると「2%ちょうどは読み直さない」が落ちる。
'   ・altRatio < utf8Ratio を <= にすると「同率は採らない」が落ちる。
'   ・altOk を見ない実装にすると「読み直せなければ UTF-8 のまま」が落ちる。
'   ・「常に cp932」実装は化けていない側の2本で落ち、「常に utf8」実装は
'     真に減る1本で落ちる(両方向を対で置いてある)。
Private Sub TestAutoCharset36()
    ChkBool36 "F31_化けていなければ読み直さない(0%)", _
        modUtilText.AutoNeedsRetry(0#), False
    ChkBool36 "F31_2%ちょうどは読み直さない", _
        modUtilText.AutoNeedsRetry(0.02), False
    ChkBool36 "F31_2%を少しでも超えたら読み直す", _
        modUtilText.AutoNeedsRetry(0.0201), True
    ChkBool36 "F31_CP932を丸ごとUTF-8で読んだ形(90%)は読み直す", _
        modUtilText.AutoNeedsRetry(0.9), True

    ChkStr36 "F31_化けていなければUTF-8のまま", _
        modUtilText.AutoTextPick(0.01, True, 0#), "utf8"
    ChkStr36 "F31_2%ちょうどでもUTF-8のまま", _
        modUtilText.AutoTextPick(0.02, True, 0#), "utf8"
    ChkStr36 "F31_化けていて読み直しで真に減ればCP932", _
        modUtilText.AutoTextPick(0.9, True, 0.01), "cp932"
    ChkStr36 "F31_減っても化けたままなら採る(比率だけで決める)", _
        modUtilText.AutoTextPick(0.9, True, 0.5), "cp932"
    ChkStr36 "F31_同率は採らない(判定できなかった扱い)", _
        modUtilText.AutoTextPick(0.9, True, 0.9), "utf8"
    ChkStr36 "F31_増えるなら採らない", _
        modUtilText.AutoTextPick(0.5, True, 0.6), "utf8"
    ChkStr36 "F31_読み直せなければUTF-8のまま", _
        modUtilText.AutoTextPick(0.9, False, 0#), "utf8"
End Sub

' ---- F27: 先頭バイトを読む価値がある拡張子か ------------------------------
'   これが False の拡張子では1バイトも読まない(旧実装は ADODB.Stream の
'   LoadFromFile がファイル【全体】をメモリへ載せてから捨てていた。PDF 取込は
'   必ずここを通るので、32bit Excel では数百MBのPDFが OOM 域だった)。
'   discriminate:
'   ・大小の正規化を外すと ".XLSX" が落ちる。
'   ・一覧に "pdf" を足すと「PDFは読まない」が落ちる。
'   ・一覧から OOXML を1つ落とすと該当の1本が落ちる。
Private Sub TestHeaderCheckable36()
    ChkBool36 "F27_xlsxは読む価値がある", modShelfScan.HeaderCheckable("xlsx"), True
    ChkBool36 "F27_docxも読む", modShelfScan.HeaderCheckable("docx"), True
    ChkBool36 "F27_pptxも読む", modShelfScan.HeaderCheckable("pptx"), True
    ChkBool36 "F27_大文字でも読む", modShelfScan.HeaderCheckable("XLSX"), True
    ChkBool36 "F27_前後の空白も均す", modShelfScan.HeaderCheckable(" xlsm "), True
    ChkBool36 "F27_PDFは読まない(判別不能と分かっている)", _
        modShelfScan.HeaderCheckable("pdf"), False
    ChkBool36 "F27_旧xlsは読まない(常にOLEで判別不能)", _
        modShelfScan.HeaderCheckable("xls"), False
    ChkBool36 "F27_旧docも読まない", modShelfScan.HeaderCheckable("doc"), False
    ChkBool36 "F27_txtは読まない", modShelfScan.HeaderCheckable("txt"), False
    ChkBool36 "F27_拡張子が無くても落ちない", modShelfScan.HeaderCheckable(""), False

    ' 判定側(EncryptedByHeader)が同じ門を使っていること=読んだ後も一貫する。
    ChkStr36 "F27_対象外の拡張子はヘッダが何でも判別不能", _
        modShelfScan.EncryptedByHeader("pdf", "D0CF11E0"), ""
    ChkStr36 "F27_対象の拡張子ならOLEは暗号化", _
        modShelfScan.EncryptedByHeader("xlsx", "D0CF11E0"), "enc"
    ChkStr36 "F27_対象の拡張子でZIPなら素のまま", _
        modShelfScan.EncryptedByHeader("xlsx", "504B0304"), "plain"
End Sub

' ---- F30: 宛先ハッシュと称号キーの単一情報源 ------------------------------
'   送信側だけが SanitizeId を通していたため、CN のエスケープ(`\`)や64字超で
'   送信ファイル名と受信側の探すファイル名が別物になり、「送りました」が
'   嘘になっていた。両側を IdHash に通す。
'   discriminate:
'   ・IdHash から SanitizeId を外すと「禁止文字の有無で同じ答え」が落ちる。
'   ・IdKey から LCase を外すと「大小を同一視する」が落ちる。
'   ・IdHash に LCase を足すと「ハッシュは大小を潰さない」が落ちる
'     (潰すと共有フォルダに既に置かれた送信済みファイルが孤立する)。
Private Sub TestIdKeyHash36()
    ChkStr36 "F30_キーは小文字へ揃える", modP2PIo.IdKey("Yamada"), "yamada"
    ChkStr36 "F30_キーは禁止文字を均す", modP2PIo.IdKey("a\b"), "a_b"
    ChkStr36 "F30_キーは前後の空白を落とす", modP2PIo.IdKey("  U1  "), "u1"
    ' 片側だけ SanitizeId を通しても答えが変わらない(=どちらの側から呼んでも
    ' 同じキーになる)。これが「送信側だけ通していた」問題の再発防止の型。
    ChkBool36 "F30_送信側と受信側は同じキーになる", _
        (StrComp(modP2PIo.IdKey("Yamada\, Taro"), _
                 modP2PIo.IdKey(modP2PIo.SanitizeId("Yamada\, Taro")), _
                 vbBinaryCompare) = 0), True

    ' ハッシュ: 生IDと SanitizeId 済みIDで同じ答えになる(=片側だけ通しても
    ' ファイル名が食い違わない)。これが F30 の症状そのものの再現防止。
    ChkBool36 "F30_生IDとサニタイズ済みIDのハッシュが一致する", _
        (StrComp(modP2PIo.IdHash("yamada\taro"), _
                 modP2PIo.IdHash(modP2PIo.SanitizeId("yamada\taro")), _
                 vbBinaryCompare) = 0), True
    ChkBool36 "F30_64字超でも両側で一致する", _
        (StrComp(modP2PIo.IdHash(String$(80, "a")), _
                 modP2PIo.IdHash(modP2PIo.SanitizeId(String$(80, "a"))), _
                 vbBinaryCompare) = 0), True
    ChkBool36 "F30_ハッシュは16桁", (Len(modP2PIo.IdHash("u1")) = 16), True
    ' ハッシュは大小を潰さない(潰すと送信済みの既存ファイルが孤立する)。
    ChkBool36 "F30_ハッシュは大小を潰さない", _
        (StrComp(modP2PIo.IdHash("Yamada"), modP2PIo.IdHash("yamada"), _
                 vbBinaryCompare) <> 0), True
    ' 別人が同じハッシュにならない(照合キーとして成立している)。
    ChkBool36 "F30_別のIDは別のハッシュ", _
        (StrComp(modP2PIo.IdHash("u1"), modP2PIo.IdHash("u2"), _
                 vbBinaryCompare) <> 0), True
End Sub

' ----------------------------------------------------------------------------
' R33H M12: 組織合計の引き継ぎ(単調増加を守る / キーが変われば0から積み直す)
' ----------------------------------------------------------------------------
'   F18 で走査窓を回した結果、スナップショットはその回の窓に入った人だけの
'   合計になり、「今月 約1,200時間」が翌日「約780時間」へ【減る】。前回の値と
'   大きい方を採って月内は減らないようにする(ユーザー裁定)。
'   ただしキーが変われば1つも引き継がない ―― ここを外すと月が替わっても
'   先月の値が max で生き残り【今月が永久に減らない】= 元より重い壊れ方。
'   discriminate(両方向を対で置く):
'   ・max をやめる(常に今回の値)と「同日内の再実行」「日替わり」が落ちる。
'   ・キー判定を外す(常に max)と「月替わり」「年替わり」「年末」が落ちる。
'   ・日/月/年を1つのフラグでまとめると、年末(3つ同時に替わる)か
'     日替わり(今日だけ替わる)のどちらかが必ず落ちる。
' ----------------------------------------------------------------------------
Private Sub TestBoardCarry36()
    ' (1) 引き継ぎの1マス
    ChkLong36 "M12_同じキーなら大きい方(前回が大)", _
        modTelemetry.BoardCarryNum("202608", "202608", 1200, 780), 1200
    ChkLong36 "M12_同じキーなら大きい方(今回が大)", _
        modTelemetry.BoardCarryNum("202608", "202608", 780, 1200), 1200
    ChkLong36 "M12_キーが違えば引き継がない", _
        modTelemetry.BoardCarryNum("202607", "202608", 1200, 80), 80
    ChkLong36 "M12_今のキーが空なら引き継がない", _
        modTelemetry.BoardCarryNum("202608", "", 1200, 80), 80

    ' (2) ヘッダ全体。日・月・年をそれぞれ独立に判定する。
    Dim cur As String
    cur = modShare.BoardHeadText("2026-08-16 09:10:00", "20260816", 80, _
        "202608", 780, "2026", 4000, 500, True, 12000)

    ' 同じ日にもう一度走った: 3つとも減らない。
    ChkStr36 "M12_同日内の再実行は3つとも減らない", _
        modTelemetry.BoardCarriedHead(modShare.BoardHeadText("2026-08-16 09:00:00", _
            "20260816", 100, "202608", 1200, "2026", 5000, 500, True, 12000), cur), _
        modShare.BoardHeadText("2026-08-16 09:10:00", "20260816", 100, _
            "202608", 1200, "2026", 5000, 500, True, 12000)

    ' 日が替わった(同じ月): 今日は0から、今月と今年は引き継ぐ。
    ChkStr36 "M12_日替わりは今日だけ0から", _
        modTelemetry.BoardCarriedHead(modShare.BoardHeadText("2026-08-15 09:00:00", _
            "20260815", 100, "202608", 1200, "2026", 5000, 500, True, 12000), cur), _
        modShare.BoardHeadText("2026-08-16 09:10:00", "20260816", 80, _
            "202608", 1200, "2026", 5000, 500, True, 12000)

    ' 月が替わった(同じ年): 今日と今月は0から、今年だけ引き継ぐ。
    ChkStr36 "M12_月替わりは今日と今月が0から", _
        modTelemetry.BoardCarriedHead(modShare.BoardHeadText("2026-07-31 09:00:00", _
            "20260731", 100, "202607", 1200, "2026", 5000, 500, True, 12000), cur), _
        modShare.BoardHeadText("2026-08-16 09:10:00", "20260816", 80, _
            "202608", 780, "2026", 5000, 500, True, 12000)

    ' 年末→年始(日・月・年が同時に替わる): 3つとも0から。
    Dim newYear As String
    newYear = modShare.BoardHeadText("2027-01-01 09:10:00", "20270101", 80, _
        "202701", 780, "2027", 4000, 500, True, 12000)
    ChkStr36 "M12_年末をまたぐと3つとも0から", _
        modTelemetry.BoardCarriedHead(modShare.BoardHeadText("2026-12-31 09:00:00", _
            "20261231", 100, "202612", 1200, "2026", 5000, 500, True, 12000), newYear), _
        newYear

    ' 年だけ替わって月キーが同じという入力は現実には無いが、独立判定なので
    ' 「今年だけ0から」に倒れる(1つのフラグでまとめた実装はここで落ちる)。
    ChkStr36 "M12_年だけ替われば今年だけ0から", _
        modTelemetry.BoardCarriedHead(modShare.BoardHeadText("2026-08-16 09:00:00", _
            "20260816", 100, "202608", 1200, "2025", 5000, 500, True, 12000), cur), _
        modShare.BoardHeadText("2026-08-16 09:10:00", "20260816", 100, _
            "202608", 1200, "2026", 4000, 500, True, 12000)

    ' 壊れた引き継ぎ元(ヘッダが空)からは何も引き継がない。
    ChkStr36 "M12_引き継ぎ元が空なら今回の値のまま", _
        modTelemetry.BoardCarriedHead("", cur), cur

    ' (3) 部別(D行)は月キーで一括判定する。
    ChkBool36 "M12_同じ月なら部別も引き継ぐ", _
        modTelemetry.BoardCarryDeptOk(modShare.BoardHeadText("2026-08-15 09:00:00", _
            "20260815", 100, "202608", 1200, "2026", 5000, 500, True, 12000), cur), True
    ChkBool36 "M12_月が替われば部別は1行も引き継がない", _
        modTelemetry.BoardCarryDeptOk(modShare.BoardHeadText("2026-07-31 09:00:00", _
            "20260731", 100, "202607", 1200, "2026", 5000, 500, True, 12000), cur), False
    ChkBool36 "M12_引き継ぎ元が空なら部別も引き継がない", _
        modTelemetry.BoardCarryDeptOk("", cur), False

    ' (4) 部の行にも概算の断りを出す(母数はヘッダ側の BoardApproxSuffix)。
    ChkStr36 "M12_概算のときは部の行にも断りを付ける", _
        modShare.BoardDeptLine("営業", 120, True), vbLf & "  部(営業)で今月 約2時間(概算)"
    ChkStr36 "M12_概算でなければ従来どおり", _
        modShare.BoardDeptLine("営業", 120, False), vbLf & "  部(営業)で今月 約2時間"
End Sub

Private Sub ChkLong36(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R33H-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkBool36(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R33H-" & label, (got = want), "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkStr36(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R33H-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Public Sub RunAll36()
    On Error GoTo H01Fail36
    TestBoardRow36
H02Next36:
    On Error GoTo H02Fail36
    TestBoardReadRows36
H03Next36:
    On Error GoTo H03Fail36
    TestBoardTexts36
H04Next36:
    On Error GoTo H04Fail36
    TestAutoCharset36
H05Next36:
    On Error GoTo H05Fail36
    TestHeaderCheckable36
H06Next36:
    On Error GoTo H06Fail36
    TestIdKeyHash36
H07Next36:
    On Error GoTo H07Fail36
    TestBoardCarry36
H01Done36:
    On Error GoTo 0
    Exit Sub

H01Fail36:
    modTestRunner.Check "TestBoardRow36(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next36
H02Fail36:
    modTestRunner.Check "TestBoardReadRows36(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next36
H03Fail36:
    modTestRunner.Check "TestBoardTexts36(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next36
H04Fail36:
    modTestRunner.Check "TestAutoCharset36(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H05Next36
H05Fail36:
    modTestRunner.Check "TestHeaderCheckable36(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H06Next36
H06Fail36:
    modTestRunner.Check "TestIdKeyHash36(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H07Next36
H07Fail36:
    modTestRunner.Check "TestBoardCarry36(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done36
End Sub
