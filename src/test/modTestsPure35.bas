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
'   ・成立/不成立を続けて撃ち、その間で GroundingAllowed の答えが反転する
'     ことまで見るので、「印を1つに戻す(モード名で兼ねる)」実装は必ず落ちる。
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

    ' 空質問ターンの印は従来どおり mLastMode="" が持つ(R33H Fix波3: その印を
    ' 返していた modMode.AnsweredMode は呼び出し元0件になったため撤去した)。

    ' 後始末: 次のテスト群へ状態を持ち越さない。
    ChkStr35 "F7_閉じ直せる", modMode.NoteAnswered(False, ""), ""
End Sub

' ----------------------------------------------------------------------------
' R33H F15: 集約スナップショットの終端行(半端なファイルを "ok" と言わない)
' ----------------------------------------------------------------------------
'   置き換えが FileCopy だった頃、読み手は「ヘッダだけ揃った状態」を掴み得た。
'   ヘッダは80〜100バイトしかないので部分読みの大多数はヘッダが完全になり、
'   列数と種別印しか見ない BoardHeadStatus はそれを "ok" と判定する ――
'   部の合算0・他人の称号が全滅したまま「集計時点/N名ぶん」と表示し、
'   TTL(600秒)のあいだそれが続いた。
'   discriminate の作り:
'   ・「常に -1(broken)」実装にすると、完全な本文の ok/行数2 が落ちる。
'   ・「常に 0以上」実装にすると、ヘッダだけ・行数改ざんの3件が落ちる。
'   ・終端行の【行数の突き合わせ】を外すと「行が1本欠けた本文」が落ちる。
'   ・最後の1本(旧入口はokと言う)が、この検査が実際に何かを変えている
'     ことの証拠になる(恒真アサートでないことの対)。
' ----------------------------------------------------------------------------
'   ※ Scripting.Dictionary は LibreOffice に存在しない(CreateObject で群ごと
'     Err=323 になる)。したがって BoardBodyText へ実辞書を渡す形は LO では
'     組めないので、辞書ゼロ件(Nothing)の1本だけを本物の BoardBodyText で
'     撃ち、行が在る形は同じ書式の文字列を手で組んで撃つ。
Private Sub TestBoardEnd35()
    Dim h As String
    h = modShare.BoardHeadText("2026-08-16 09:00:00", "20260816", 120, _
        "202608", 900, "2026", 5000, 42, False)
    Dim body As String
    body = h & vbLf & "D" & vbTab & "営業部" & vbTab & "300" & _
           vbLf & "T" & vbTab & "u1" & vbTab & "7" & _
           vbLf & modShare.BOARD_END_TAG & vbTab & "2"

    ChkStr35 "F15_終端行はデータ行数を宣言する", _
        Right$(modShare.BoardBodyText(h, Nothing, Nothing), _
               Len(modShare.BOARD_END_TAG) + 2), modShare.BOARD_END_TAG & vbTab & "0"
    ChkLong35b "F15_完全な本文は行数を返す", modShare.BoardEndCount(body), 2
    ChkLong35b "F15_ヘッダだけ(書きかけ)は-1", modShare.BoardEndCount(h), -1
    ChkLong35b "F15_終端行が無い本文は-1", _
        modShare.BoardEndCount(h & vbLf & "D" & vbTab & "営業部" & vbTab & "300"), -1
    ChkLong35b "F15_宣言と実数が食い違えば-1", _
        modShare.BoardEndCount(h & vbLf & "D" & vbTab & "営業部" & vbTab & "300" & _
        vbLf & modShare.BOARD_END_TAG & vbTab & "5"), -1
    ChkLong35b "F15_行が欠けても宣言と合わなければ-1", _
        modShare.BoardEndCount(h & vbLf & modShare.BOARD_END_TAG & vbTab & "2"), -1
    ChkLong35b "F15_データ0件でも終端行があれば0", _
        modShare.BoardEndCount(modShare.BoardBodyText(h, Nothing, Nothing)), 0

    ChkStr35 "F15_入口: 完全で新しければok", _
        modShare.BoardTextStatus(body, "2026-08-15 09:00:00"), "ok"
    ChkStr35 "F15_入口: 完全でも古ければstale", _
        modShare.BoardTextStatus(body, "2026-08-16 09:00:01"), "stale"
    ChkStr35 "F15_入口: ヘッダだけならbroken", _
        modShare.BoardTextStatus(h, "2026-08-15 09:00:00"), "broken"
    ' ★この1本が「検査が実際に何かを変えている」ことの証拠。旧入口
    '   (ヘッダ1行だけを見る BoardHeadStatus)は同じ入力を ok と言う。
    ChkStr35 "F15_旧入口は同じ半端な入力をokと言っていた", _
        modShare.BoardHeadStatus(modShare.BoardHeadLine(h), "2026-08-15 09:00:00"), "ok"
End Sub

' ----------------------------------------------------------------------------
' R33H F18: 走査の打ち切りに「恒久除外」を作らない
' ----------------------------------------------------------------------------
'   ビーコン名は stats_<不変ハッシュ>.txt で NTFS は名前順に返すため、先頭から
'   上限本で切る旧実装では【打ち切られる顔ぶれが毎回同じ】になり、その人たちは
'   感謝を20件集めても称号が誰の画面にも付かなかった。起点を日ごとにずらす。
'   discriminate の作り: 最後の2本が対になっている ―― 回す実装なら
'   ceil(本数/窓)日で全員が入り(miss=0)、旧実装(起点固定)なら4人が
'   永久に外れる(miss=4)。「常に0を返す」実装は前者で落ち、「常にずらす」
'   実装は「本数が上限以下なら0」の3本で落ちる。
' ----------------------------------------------------------------------------
Private Sub TestScanRotate35()
    ChkLong35b "F18_本数が上限以下なら回さない", modShare.BoardScanStart(500, 500, 12345), 0
    ChkLong35b "F18_本数0なら0", modShare.BoardScanStart(0, 500, 12345), 0
    ChkLong35b "F18_上限0なら0(壊れた指定)", modShare.BoardScanStart(700, 0, 12345), 0

    ChkLong35b "F18_1日ごとに窓ぶん進む(1日目)", modShare.BoardScanStart(7, 3, 1), 3
    ChkLong35b "F18_1日ごとに窓ぶん進む(2日目)", modShare.BoardScanStart(7, 3, 2), 6
    ChkLong35b "F18_輪になって戻る(3日目=9 mod 7)", modShare.BoardScanStart(7, 3, 3), 2

    Dim hit(0 To 6) As Boolean
    Dim fixedHit(0 To 6) As Boolean
    Dim d As Long, i As Long, idx As Long
    For d = 0 To 2                        ' ceil(7/3)=3日ぶん
        Dim st As Long: st = modShare.BoardScanStart(7, 3, d)
        For i = 0 To 2
            idx = (st + i) Mod 7
            hit(idx) = True
            fixedHit((0 + i) Mod 7) = True    ' 旧実装=起点が常に0
        Next i
    Next d
    Dim miss As Long, missFixed As Long
    For i = 0 To 6
        If Not hit(i) Then miss = miss + 1
        If Not fixedHit(i) Then missFixed = missFixed + 1
    Next i
    ChkLong35b "F18_3日で全員が窓に入る(恒久除外なし)", miss, 0
    ChkLong35b "F18_起点を固定すると4人が恒久的に外れる(旧実装の再現)", missFixed, 4
End Sub

' ----------------------------------------------------------------------------
' R33H F18/F17: 画面へ出す文字列(概算の母数・部の行)
' ----------------------------------------------------------------------------
'   「(概算)」としか出ないと、何名ぶんの数字なのかが画面から分からない。
'   部の行は60分未満を時間へ丸めると「約0時間」になる(R13 L-batch)。
' ----------------------------------------------------------------------------
Private Sub TestBoardTexts35()
    ChkStr35 "F18_概算でなければ何も付けない", _
        modShare.BoardApproxSuffix(500, 12000, False), ""
    ChkStr35 "F18_概算なら母数を入れる", _
        modShare.BoardApproxSuffix(500, 12000, True), "(概算500/12000名)"
    ChkStr35 "F18_母数が取れない旧版は概算だけ", _
        modShare.BoardApproxSuffix(500, 0, True), "(概算)"

    ChkStr35 "F18_部が不明なら行を出さない", modShare.BoardDeptLine("", 300), ""
    ChkStr35 "F18_0分なら行を出さない", modShare.BoardDeptLine("営業", 0), ""
    ChkStr35 "F18_60分未満は分のまま", modShare.BoardDeptLine("営業", 59), _
        vbLf & "  部(営業)で今月 約59分"
    ChkStr35 "F18_60分以上は時間", modShare.BoardDeptLine("営業", 120), _
        vbLf & "  部(営業)で今月 約2時間"
End Sub

' ----------------------------------------------------------------------------
' R33H F21: 冷えたVPN復帰を「到達不能」で確定させない(第3引数の追加分)
' ----------------------------------------------------------------------------
'   W4-6 の1秒閾値は「構文起因なら即座に返る」を根拠にしているが、その裏返し
'   (即座でないなら到達性の問題)は成り立たない ―― SMBセッション未確立では
'   1回目の GetAttr が数秒かけて失敗する。その1回で共有機能がセッション丸ごと
'   死に、W2-2 の reachableNow=False 経由で知識の消去にまで連鎖する。
'   既存の2引数の答え(modTestsPure34 が999/1000/1001で固定)は1つも変えず、
'   「取り違えの代償が大きい呼び口」だけ 1秒〜PROBE_COLD_FAIL_MS の窓を開ける。
'   discriminate(両方向の対):
'   ・第3引数を無視する実装に戻すと「冷えたSMB」の2本が落ちる。
'   ・第3引数で無条件に再試行する実装にすると「タイムアウト級」の2本が落ちる。
'   ・1秒の側を動かすと modTestsPure34 の境界3本(999/1000/1001)が落ちる。
' ----------------------------------------------------------------------------
Private Sub TestRetryProbe35()
    ' --- 既定(第3引数なし)は W4-6 のまま ---------------------------------
    ChkBool35 "F21_既定は1秒以内だけ再試行", _
        modShareRule.ShouldRetryProbe(76, 1000), True
    ChkBool35 "F21_既定は1秒を超えたら再試行しない", _
        modShareRule.ShouldRetryProbe(76, 1001), False
    ChkBool35 "F21_既定は3秒でも再試行しない(冷えたVPNを撃ち抜いていた)", _
        modShareRule.ShouldRetryProbe(76, 3000), False

    ' --- 代償の大きい呼び口では、冷えたSMBの幅まで許す ---------------------
    ChkBool35 "F21_代償が大きい経路は3秒でも再試行する", _
        modShareRule.ShouldRetryProbe(76, 3000, True), True
    ChkBool35 "F21_境界ちょうど(5秒)は再試行する", _
        modShareRule.ShouldRetryProbe(76, modShareRule.PROBE_COLD_FAIL_MS, True), True

    ' --- タイムアウト級は、どちらの経路でも二度払わない --------------------
    ChkBool35 "F21_境界の1つ先(5001ms)は再試行しない", _
        modShareRule.ShouldRetryProbe(76, modShareRule.PROBE_COLD_FAIL_MS + 1, True), False
    ChkBool35 "F21_名前解決のタイムアウト(15秒)は再試行しない", _
        modShareRule.ShouldRetryProbe(53, 15000, True), False
    ChkBool35 "F21_SMBのタイムアウト(30秒)は再試行しない", _
        modShareRule.ShouldRetryProbe(76, 30000, True), False

    ' --- 答えを決めるのは経過時間と呼び口だけで、エラー番号ではない --------
    ChkBool35 "F21_番号が違っても即失敗なら再試行", _
        modShareRule.ShouldRetryProbe(0, 10, False), True
    ChkBool35 "F21_番号が違っても遅ければ代償の大きい経路のみ", _
        modShareRule.ShouldRetryProbe(52, 2000, True), True
    ChkBool35 "F21_同じ材料で既定なら再試行しない", _
        modShareRule.ShouldRetryProbe(52, 2000, False), False
End Sub

Private Sub ChkLong35b(ByVal label As String, ByVal got As Long, ByVal want As Long)
    modTestRunner.Check "R33-" & label, (got = want), "実際=" & got & " 期待=" & want
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
H09Next35:
    On Error GoTo H09Fail35
    TestBoardEnd35
H10Next35:
    On Error GoTo H10Fail35
    TestScanRotate35
H11Next35:
    On Error GoTo H11Fail35
    TestBoardTexts35
H12Next35:
    On Error GoTo H12Fail35
    TestRetryProbe35
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
    Resume H09Next35
H09Fail35:
    modTestRunner.Check "TestBoardEnd35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H10Next35
H10Fail35:
    modTestRunner.Check "TestScanRotate35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H11Next35
H11Fail35:
    modTestRunner.Check "TestBoardTexts35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H12Next35
H12Fail35:
    modTestRunner.Check "TestRetryProbe35(グループ全体)", False, _
        "群の実行中に例外: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume H01Done35
End Sub
