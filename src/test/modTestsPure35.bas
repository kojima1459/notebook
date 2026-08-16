Attribute VB_Name = "modTestsPure35"
Option Explicit

' ============================================================================
' modTestsPure35 - R33波5c(部門ごとの削除・W5-23)の純ロジック回帰テスト。
'   modTestsPure34 が残1,676字になったための分割先。既存チェーン
'   (modTestsPure.RunAll→…→modTestsPure30.RunAll30)へは繋がず、
'   modTestRunner.RunAllPureTests から直接呼ばれる RunAll35 の1本が入口
'   (modTestsPure31 / 32 / 33 / 34 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【このモジュールが何を守るのか】
'   W5-23: ナレッジ画面の「部門の資料を削除」が、消してよいものだけを
'     数え、消してはいけないものを1件も巻き込まないこと。
'     ・数えるのは origin が "channel:" で始まる行だけ。
'       self(自作) / pack:(手渡しパック) / 空 は絶対に数に入らない。
'     ・部門名の切り出しは前置きの後ろ全部。"channel:" だけの行は
'       名前が無いので数えない(空の部門名で削除ダイアログを出さない)。
'     ・番号→部門名の対応が1つでもずれると、利用者は「商品部を消すつもりで
'       人事部を消す」ことになる。ここは取り消せない操作なので、番号の
'       境界(0 / 1 / 最終 / 最終+1)を全部固定する。
'
'   ※ 実際に本棚の行が消えるのは modShelfStore.RemoveRowsByOrigin(完全一致)で、
'     こちらは Worksheet を触るため LO では検証できない(tools/README §4)。
'     恒真アサートで代替せず、純関数側の線引きだけをここで固定する。
' ============================================================================

' 「章立ての違う3部門 + 自作 + 手渡しパック + 空行 + 名前無しchannel」を
' 1本に混ぜたゴールデン入力。実機の origin 列に実際に並びうる形。
Private Function Sample35() As String
    Dim a(0 To 10) As String
    a(0) = "channel:商品部"
    a(1) = "self"
    a(2) = "channel:人事部"
    a(3) = "pack:商品部"          ' 手渡しパック(部門名が同じでも別物)
    a(4) = "channel:商品部"
    a(5) = ""
    a(6) = "channel:"             ' 前置きだけ=部門名が無い
    a(7) = "pack:田中"
    a(8) = "channel:システム部"
    a(9) = "self"
    a(10) = "channel:商品部"
    Sample35 = Join(a, vbLf)
End Function

' ---- W5-23(1): 数えるのは channel: だけ -----------------------------------
'   discriminate:
'   ・前置き判定を「前方一致ではなく部分一致」へ緩めると pack:商品部 を拾い、
'     商品部が 3→4 件になって落ちる。
'   ・Len(o) > pl を Len(o) >= pl へ緩めると "channel:" 単独行を空名で拾い、
'     件数が4部門になって落ちる。
'   ・prefix を "self" にした行で0件になることを対で置いているので、
'     「常に全部数える」実装も通らない。
Private Sub TestOriginCounts35()
    Dim got As String
    got = modShareRule.OriginCountsText(Sample35(), "channel:")
    ChkStr35 "W5-23_channelだけを出現順に数える", got, _
        "商品部" & vbTab & "3|人事部" & vbTab & "1|システム部" & vbTab & "1"

    ChkStr35 "W5-23_pack:は別名前空間として別に数える", _
        modShareRule.OriginCountsText(Sample35(), "pack:"), _
        "商品部" & vbTab & "1|田中" & vbTab & "1"

    ChkStr35 "W5-23_selfは前置きの後ろが無いので0件", _
        modShareRule.OriginCountsText(Sample35(), "self"), ""

    ChkStr35 "W5-23_空入力は空", modShareRule.OriginCountsText("", "channel:"), ""
    ChkStr35 "W5-23_前置きが空なら何も数えない", _
        modShareRule.OriginCountsText(Sample35(), ""), ""
    ChkStr35 "W5-23_一致が無ければ空", _
        modShareRule.OriginCountsText(Sample35(), "channel:営業部"), ""

    ' 1行だけ(本棚に1資料しかない端末)でも配列化に頼らず数えられること。
    ChkStr35 "W5-23_1行だけでも数えられる", _
        modShareRule.OriginCountsText("channel:全社共通", "channel:"), _
        "全社共通" & vbTab & "1"
    ' 前後の空白は落とす(セル書式で空白が入っている実機の行)。
    ChkStr35 "W5-23_前後の空白は落として同じ部門として数える", _
        modShareRule.OriginCountsText("channel:商品部" & vbLf & "  channel:商品部  ", "channel:"), _
        "商品部" & vbTab & "2"
End Sub

' ---- W5-23(2): 番号→部門名の対応がずれない ---------------------------------
'   discriminate:
'   ・PurgeMenuPick を 0始まりへ変えると「1番」が人事部を返して落ちる。
'   ・上限チェック(choice > n)を外すと4番が実行時エラー/空以外になり落ちる。
Private Sub TestPurgeMenuPick35()
    Dim c As String
    c = modShareRule.OriginCountsText(Sample35(), "channel:")

    ChkStr35 "W5-23_1番は先頭(商品部)", modShareRule.PurgeMenuPick(c, 1), _
        "商品部" & vbTab & "3"
    ChkStr35 "W5-23_2番は人事部", modShareRule.PurgeMenuPick(c, 2), _
        "人事部" & vbTab & "1"
    ChkStr35 "W5-23_3番は末尾(システム部)", modShareRule.PurgeMenuPick(c, 3), _
        "システム部" & vbTab & "1"
    ChkStr35 "W5-23_0番は選べない", modShareRule.PurgeMenuPick(c, 0), ""
    ChkStr35 "W5-23_負の番号は選べない", modShareRule.PurgeMenuPick(c, -1), ""
    ChkStr35 "W5-23_件数を超える番号は選べない", modShareRule.PurgeMenuPick(c, 4), ""
    ChkStr35 "W5-23_一覧が空なら何番でも選べない", modShareRule.PurgeMenuPick("", 1), ""
End Sub

' ---- W5-23(3): 選択肢の見出しは番号付きで、上限を超えたら明示する ---------
'   MsgBox/InputBox のプロンプトは1,024字までで、部門は最大40まで有りうる。
'   全部並べると末尾が無言で切れるため、上限を超えたぶんは件数だけ添える。
'   discriminate: 上限処理を消すと「…ほか」行が出ず、3行目の期待が落ちる。
Private Sub TestPurgeMenuText35()
    Dim c As String
    c = modShareRule.OriginCountsText(Sample35(), "channel:")

    ChkStr35 "W5-23_番号と件数を並べる", modShareRule.PurgeMenuText(c, 20), _
        "  1) 商品部  (3件)" & vbLf & "  2) 人事部  (1件)" & vbLf & _
        "  3) システム部  (1件)"

    ChkStr35 "W5-23_上限を超えたら残数を明示する", modShareRule.PurgeMenuText(c, 2), _
        "  1) 商品部  (3件)" & vbLf & "  2) 人事部  (1件)" & vbLf & _
        "  …ほか 1 部門(番号を直接入力すると選べます)"

    ' 上限を超えて隠れた部門も、番号を直接打てば選べる(選べなくはしない)。
    ChkStr35 "W5-23_隠れた3番も番号指定なら選べる", _
        modShareRule.PurgeMenuPick(c, 3), "システム部" & vbTab & "1"

    ChkStr35 "W5-23_一覧が空なら見出しも空", modShareRule.PurgeMenuText("", 20), ""
End Sub

' ---- W5-27: 📊利用状況を開けるかの真理表 ----------------------------------
'   admin_users を設定した組織では名簿だけを見る。未設定のときに限り
'   発行者(publish_key あり)へ開放する。
'   discriminate(両方向を対で置く):
'   ・フォールバックを常時有効(= isAdmin Or canPublish を無条件)に壊すと、
'     「名簿あり×名簿外×発行者」が True になって落ちる。
'   ・フォールバックを削る(= 常に isAdmin だけ)と、
'     「名簿なし×発行者」が False になって落ちる。
Private Sub TestCanViewUsage35()
    ' (1) admin_users 設定あり … 名簿だけを見る(発行者フラグで上書きしない)
    ChkBool35 "W5-27_名簿あり×名簿に載っている → 開ける", _
        modShareRule.CanViewUsage(True, True, False), True
    ChkBool35 "W5-27_名簿あり×名簿に載っている×発行者でもある → 開ける", _
        modShareRule.CanViewUsage(True, True, True), True
    ChkBool35 "W5-27_名簿あり×名簿外×発行者 → 開けない(意図を上書きしない)", _
        modShareRule.CanViewUsage(True, False, True), False
    ChkBool35 "W5-27_名簿あり×名簿外×一般 → 開けない", _
        modShareRule.CanViewUsage(True, False, False), False

    ' (2) admin_users 未設定 … 名簿による制限が存在しないので発行者へ開放
    ChkBool35 "W5-27_名簿なし×発行者 → 開ける(フォールバック)", _
        modShareRule.CanViewUsage(False, False, True), True
    ChkBool35 "W5-27_名簿なし×一般 → 開けない", _
        modShareRule.CanViewUsage(False, False, False), False
    ChkBool35 "W5-27_名簿なし×発行者でない管理者判定 → 開ける", _
        modShareRule.CanViewUsage(False, True, False), True
End Sub

' ---- R33H F1(BLOCKER): 削除の一致判定は全角/半角・かな/カナを畳まない -----
'   日本語ロケールの VBA では vbTextCompare が全角/半角・ひらがな/カタカナを
'   同一視する。「営業1課」を消すと「営業１課」まで消え、しかも版の記録
'   (my_stats "ch:")は LCase$ しか通さないので巻き添え側だけ記録が残り、
'   PendingUpdates が「最新です」と言って【永久に戻らない】。
'   LO も vba_lint もこの差を検出しないので、ここで固定する。
'   discriminate(両方向を対で置く):
'   ・OriginMatches を vbTextCompare へ戻すと「全角/半角」「かな/カナ」
'     「半角カナ/全角カナ」の3本が True になって落ちる。
'   ・OriginKeyNorm から LCase$ を外すと「ASCIIの大小は畳む」2本が落ちる。
'   ・ChannelStatKey を "ch:" & chName(正規化なし)へ戻すと大小の1本が落ちる。
Private Sub TestOriginMatches35()
    ChkBool35 "F1_同じタグは一致する", _
        modShareRule.OriginMatches("channel:営業1課", "channel:営業1課"), True
    ChkBool35 "F1_全角数字の部門は別物(巻き添え禁止)", _
        modShareRule.OriginMatches("channel:営業1課", "channel:営業１課"), False
    ChkBool35 "F1_ひらがなとカタカナは別物", _
        modShareRule.OriginMatches("channel:さくら課", "channel:サクラ課"), False
    ChkBool35 "F1_半角カナと全角カナは別物", _
        modShareRule.OriginMatches("channel:ｻｸﾗ課", "channel:サクラ課"), False
    ChkBool35 "F1_全角英字と半角英字は別物", _
        modShareRule.OriginMatches("channel:ＡＢ課", "channel:AB課"), False
    ChkBool35 "F1_ASCIIの大小だけは畳む(NTFSは大小を区別しない)", _
        modShareRule.OriginMatches("channel:Sales", "channel:sales"), True
    ChkBool35 "F1_前後の空白は畳む(セル書式の空白)", _
        modShareRule.OriginMatches("  channel:商品部  ", "channel:商品部"), True
    ChkBool35 "F1_別部門は一致しない", _
        modShareRule.OriginMatches("channel:商品部", "channel:人事部"), False
    ChkBool35 "F1_前方一致では一致しない(完全一致)", _
        modShareRule.OriginMatches("channel:商品部第2", "channel:商品部"), False
    ChkBool35 "F1_名前空間が違えば一致しない", _
        modShareRule.OriginMatches("pack:商品部", "channel:商品部"), False

    ' 削除の一致判定と my_stats のキーが【同じ正規化】を通ること。
    ChkStr35 "F1_版の記録キーは同じ正規化を通る", _
        modShareRule.ChannelStatKey("営業1課"), "ch:営業1課"
    ChkStr35 "F1_全角の部門は別のキーになる", _
        modShareRule.ChannelStatKey("営業１課"), "ch:営業１課"
    ChkStr35 "F1_ASCIIの大小はキーでも畳む", _
        modShareRule.ChannelStatKey("Sales"), "ch:sales"
    ChkStr35 "F1_前後の空白はキーでも落とす", _
        modShareRule.ChannelStatKey("  商品部 "), "ch:商品部"
    ' 一致判定とキーの同値関係が一致すること(= 消した部門と記録を消す部門が
    ' 必ず同じになる)。ここがずれると永久復旧不能が再発する。
    ChkBool35 "F1_一致するタグ同士は同じ版キーを持つ", _
        (modShareRule.ChannelStatKey("Sales") = modShareRule.ChannelStatKey("sales")), True
    ChkBool35 "F1_一致しないタグ同士は別の版キーを持つ", _
        (modShareRule.ChannelStatKey("営業1課") = modShareRule.ChannelStatKey("営業１課")), False

    ' 集計(表示件数)側も同じ線引きであること。ここが vbTextCompare のままだと
    ' 「3件消します」と言って1件しか消えない(または別部門を巻き込む)。
    ChkStr35 "F1_全角と半角の同名部門は別行として数える", _
        modShareRule.OriginCountsText( _
            "channel:営業1課" & vbLf & "channel:営業１課" & vbLf & "channel:営業1課", _
            "channel:"), _
        "営業1課" & vbTab & "2|営業１課" & vbTab & "1"
    ChkStr35 "F1_ASCII大小違いは1つの部門にまとめる(表示は初出の綴り)", _
        modShareRule.OriginCountsText( _
            "channel:Sales" & vbLf & "channel:sales", "channel:"), _
        "Sales" & vbTab & "2"
End Sub

' ---- R33H F4: 購読しない部門の一覧はカンマではなく "|" で区切る ------------
'   Windows のフォルダ名にカンマは使えるので、カンマ区切りだと
'   「品質保証,監査」を解除したときに「品質保証」と「監査」の両方が届かなく
'   なる。区切りをフォルダ名に使えない "|" へ変え、旧カンマ形式は移行して読む。
'   discriminate(両方向を対で置く):
'   ・区切りをカンマへ戻すと「カンマ入りの部門名を1件として扱う」3本が落ちる。
'   ・旧形式の移行を外す(常に "|" で分ける)と「旧形式の端末が壊れない」
'     4本が落ちる。
'   ・突き合わせを vbTextCompare へ戻すと「全角の同名部門は別扱い」が落ちる。
Private Sub TestUnsubList35()
    ' (1) 新形式(| 区切り)
    ChkStr35 "F4_空は空", modShareRule.UnsubListNorm(""), ""
    ChkStr35 "F4_正規形は先頭にも区切りが付く(新形式である印)", _
        modShareRule.UnsubListNorm("商品部|人事部"), "|商品部|人事部"
    ChkStr35 "F4_正規化は冪等", _
        modShareRule.UnsubListNorm("|商品部|人事部"), "|商品部|人事部"
    ChkStr35 "F4_空要素と前後空白は落とす", _
        modShareRule.UnsubListNorm(" 商品部 ||  人事部|"), "|商品部|人事部"
    ChkStr35 "F4_重複は1つにまとめる", _
        modShareRule.UnsubListNorm("商品部|商品部"), "|商品部"

    ' (2) 旧カンマ形式からの移行(手編集で設定済みの端末が壊れない)
    ChkStr35 "F4_旧カンマ形式は移行して読む", _
        modShareRule.UnsubListNorm("商品部,人事部"), "|商品部|人事部"
    ChkBool35 "F4_旧形式で解除した部門は解除のまま", _
        modShareRule.UnsubListHas("商品部,人事部", "人事部"), True
    ChkBool35 "F4_旧形式で1件だけ書いた端末も解除のまま", _
        modShareRule.UnsubListHas("商品部", "商品部"), True
    ChkBool35 "F4_旧形式に無い部門は届く", _
        modShareRule.UnsubListHas("商品部,人事部", "システム部"), False
    ' 新形式が1つでも混ざっていれば以後カンマは区切りにしない
    ' (= カンマを含む部門名を1件として扱えるようになる)。
    ChkStr35 "F4_新形式が混ざればカンマは名前の一部", _
        modShareRule.UnsubListNorm("品質保証,監査|商品部"), "|品質保証,監査|商品部"

    ' (3) F4 の本題: カンマ入りの部門名が他部門を巻き込まない
    Dim after As String
    after = modShareRule.UnsubListAdd("", "品質保証,監査")
    ChkStr35 "F4_カンマ入りの部門名は1件として入る", after, "|品質保証,監査"
    ChkBool35 "F4_カンマ入りの部門は解除されている", _
        modShareRule.UnsubListHas(after, "品質保証,監査"), True
    ChkBool35 "F4_巻き添えにされない(品質保証は届く)", _
        modShareRule.UnsubListHas(after, "品質保証"), False
    ChkBool35 "F4_巻き添えにされない(監査は届く)", _
        modShareRule.UnsubListHas(after, "監査"), False

    ' (4) 足す・外すの往復
    ChkStr35 "F4_足す", modShareRule.UnsubListAdd("商品部", "人事部"), "|商品部|人事部"
    ChkStr35 "F4_既に載っていれば増やさない", _
        modShareRule.UnsubListAdd("商品部|人事部", "商品部"), "|商品部|人事部"
    ChkStr35 "F4_空名は足さない", modShareRule.UnsubListAdd("商品部", "  "), "|商品部"
    ChkStr35 "F4_外す", modShareRule.UnsubListRemove("商品部|人事部", "商品部"), "|人事部"
    ChkStr35 "F4_最後の1件を外すと空", _
        modShareRule.UnsubListRemove("商品部", "商品部"), ""
    ChkStr35 "F4_載っていない部門を外しても変わらない", _
        modShareRule.UnsubListRemove("商品部|人事部", "システム部"), "|商品部|人事部"
    ChkBool35 "F4_足して外せば元に戻る", _
        (modShareRule.UnsubListRemove( _
            modShareRule.UnsubListAdd("商品部", "人事部"), "人事部") = "|商品部"), True

    ' (5) 突き合わせは F1 と同じ正規化(全角/半角を畳まない・ASCII大小は畳む)
    ChkBool35 "F4_全角の同名部門は別部門として届く", _
        modShareRule.UnsubListHas("営業1課", "営業１課"), False
    ChkBool35 "F4_ASCIIの大小は同じ部門とみなす", _
        modShareRule.UnsubListHas("Sales", "sales"), True
End Sub

' ---- R33H F6: 本棚の警告に「78%です」と出さない ---------------------------
'   W5-17 で表示の分母は shelf_max_chunks へ統一されたが、警告の判定は
'   chunk_limit の8割のまま。文面が表示用のパーセントを読んでいたので、
'   既定のまま16,000件で「本棚の使用量が 78% です」と出ていた。
'   discriminate:
'   ・文面へパーセントを戻すと「%を含まない」の2本が落ちる。
'   ・目安の件数を出さない(固定文にする)と、20,000/40,000で同じ文字列に
'     なって「組織ごとの目安が出る」が落ちる。
'   ・判定 IsBudgetTightAt は【変えない】ことを同時に固定する
'     (modTestsPure34 の W5-17 と対になる)。
Private Sub TestBudgetWarn35()
    ChkStr35 "F6_警告文は目安の件数で語る(既定20,000)", _
        modShareRule.BudgetWarnCaption(20000), _
        ChrW(&H26A0) & " 棚卸しの目安(20,000件)の8割を超えました" & vbCr & _
        "使っていない資料を減らすと空きます(マイ本棚から削除できます)"
    ChkStr35 "F6_目安を広げた組織はその件数が出る", _
        modShareRule.BudgetWarnCaption(40000), _
        ChrW(&H26A0) & " 棚卸しの目安(40,000件)の8割を超えました" & vbCr & _
        "使っていない資料を減らすと空きます(マイ本棚から削除できます)"
    ChkBool35 "F6_警告文にパーセントは出さない(既定)", _
        (InStr(1, modShareRule.BudgetWarnCaption(20000), "%", vbBinaryCompare) > 0), False
    ChkBool35 "F6_警告文にパーセントは出さない(拡張時)", _
        (InStr(1, modShareRule.BudgetWarnCaption(40000), "%", vbBinaryCompare) > 0), False
    ' 判定ロジックは変えない(表示だけを変えた裁定であることの確認)。
    ChkBool35 "F6_警告線は従来どおり目安の8割ちょうどで立つ", _
        modShareRule.IsBudgetTightAt(16000, 20000), True
    ChkBool35 "F6_8割の1件手前では立たない", _
        modShareRule.IsBudgetTightAt(15999, 20000), False
End Sub

' ---- R33H F7: 「検索も生成もしなかった」と「生成が成立しなかった」を分ける ----
'   W5-21 は後者の抑止に mLastMode="" を使ったが、その印は
'   modUIMain.RenderAnswer が【空質問ターン専用】として先に使っており、
'   API失敗と逆質問がそこへ流れ込むと (a)通信失敗の直後に状態セルが
'   「準備できています」になり (b)「Wordで開く」が前のターンの回答を出した。
'   mLastMode には従来どおりモード名を入れ、根拠表示(信頼度バッジ・出典
'   チップ)の抑止だけを別の1ビットで行う。
'   discriminate(両方向を対で置く):
'   ・NoteAnswered を AnsweredMode と同じ実装(失敗時に "" を返す)へ戻すと
'     「不成立ターンでもモード名は空にしない」が落ちる。
'   ・mGrounded を常に True にすると「不成立ターンは根拠表示を止める」が落ちる。
'   ・空質問の印(AnsweredMode が "" を返すこと)も対で固定しているので、
'     2つの印を1つに戻す実装は必ずどちらかで落ちる。
Private Sub TestNoteAnswered35()
    ChkStr35 "F7_成立ターンはモード名をそのまま返す", _
        modMode.NoteAnswered(True, "deep"), "deep"
    ChkBool35 "F7_成立ターンは根拠表示を許す", modMode.GroundingAllowed(), True

    ChkStr35 "F7_不成立ターンでもモード名は空にしない", _
        modMode.NoteAnswered(False, "deep"), "deep"
    ChkBool35 "F7_不成立ターンは根拠表示を止める", modMode.GroundingAllowed(), False

    ChkStr35 "F7_モード名は素通しする(quick)", _
        modMode.NoteAnswered(True, "quick"), "quick"
    ChkBool35 "F7_見るのは直近の1ターンだけ", modMode.GroundingAllowed(), True

    ' 空質問ターンの印は従来どおり別の関数(mLastMode="")が持つ。
    ChkStr35 "F7_空質問の印は別物(検索も生成もしなかった)", _
        modMode.AnsweredMode(False, "deep"), ""
    ChkStr35 "F7_成立ターンの印は従来どおり", _
        modMode.AnsweredMode(True, "deep"), "deep"

    ' 後始末: 次のテスト群へ状態を持ち越さない。
    ChkStr35 "F7_閉じ直せる", modMode.NoteAnswered(False, ""), ""
End Sub

Private Sub ChkBool35(ByVal label As String, ByVal got As Boolean, ByVal want As Boolean)
    modTestRunner.Check "R33-" & label, (got = want), _
        "実際=" & got & " 期待=" & want
End Sub

Private Sub ChkStr35(ByVal label As String, ByVal got As String, ByVal want As String)
    modTestRunner.Check "R33-" & label, (StrComp(got, want, vbBinaryCompare) = 0), _
        "実際=[" & got & "] 期待=[" & want & "]"
End Sub

Public Sub RunAll35()
    On Error GoTo H01Fail35
    TestOriginCounts35
H02Next35:
    On Error GoTo H02Fail35
    TestPurgeMenuPick35
H03Next35:
    On Error GoTo H03Fail35
    TestPurgeMenuText35
H04Next35:
    On Error GoTo H04Fail35
    TestCanViewUsage35
H05Next35:
    On Error GoTo H05Fail35
    TestOriginMatches35
H06Next35:
    On Error GoTo H06Fail35
    TestUnsubList35
H07Next35:
    On Error GoTo H07Fail35
    TestBudgetWarn35
H08Next35:
    On Error GoTo H08Fail35
    TestNoteAnswered35
H01Done35:
    On Error GoTo 0
    Exit Sub

H01Fail35:
    modTestRunner.Check "TestOriginCounts35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H02Next35
H02Fail35:
    modTestRunner.Check "TestPurgeMenuPick35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H03Next35
H03Fail35:
    modTestRunner.Check "TestPurgeMenuText35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H04Next35
H04Fail35:
    modTestRunner.Check "TestCanViewUsage35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H05Next35
H05Fail35:
    modTestRunner.Check "TestOriginMatches35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H06Next35
H06Fail35:
    modTestRunner.Check "TestUnsubList35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H07Next35
H07Fail35:
    modTestRunner.Check "TestBudgetWarn35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H08Next35
H08Fail35:
    modTestRunner.Check "TestNoteAnswered35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done35
End Sub
