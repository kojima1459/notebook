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
    Resume H01Done35
End Sub
