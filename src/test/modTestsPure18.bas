Attribute VB_Name = "modTestsPure18"
Option Explicit

' ============================================================================
' modTestsPure18 - R17H(レビュー裁定 Fix波)の純ロジック回帰
' ----------------------------------------------------------------------------
' なぜ新設したか(憲章§4-6):
'   modTestsPure17 が25,032字で、ここの真理表(約3,400字)を足すと WARN帯
'   (28,000字)へ入る。17 を新設したときと同じ線で分割する。
'   入口は modTestsPure17.RunAll17 の末尾から呼ばれる RunAll18 の1本だけ。
'
' ここで固定するもの:
'   ・modRagParse.MergeSynPairs(FA-1 / A-H1): 名寄せ辞書のマージ規則。
'     実体は modSynonymStore.MergeAndSave にあり、ParseSynResp が返す
'     【0始まり】の配列を 1〜n で読んでいたため、実データでは毎回
'     「添字が範囲外」で握り潰され synonyms が永久に0行だった。純関数へ
'     切り出した以上、規則そのものをゴールデンで縛る(新規のみ/旧のみ/
'     上書き/先頭要素の保持/空入力/壊れた要素/大小無視/順序安定)。
'   ・modRagParse.HasGlobalSignal(FA-6 / A-M7・B-H2): 俯瞰語彙の検知。
'     拾う語と【拾わない語】の両方を固定する(誤爆しやすい短語を後から
'     足されないようにするのがこのテストの半分の目的)。
'   ・modAskMulti.DecomposeGate(FA-6): 俯瞰シグナルのOR。短い俯瞰質問でも
'     段0が呼ばれること、off が最優先であることの2点。
' ============================================================================

' ----------------------------------------------------------------------------
' FA-1: 名寄せ辞書のマージ規則(旧CSV + 新CSV → 統合CSV)。
' ----------------------------------------------------------------------------
Private Sub TestMergeSynPairs()
    ' 旧が空 = 初回の取込。新規ぶんがそのまま順序どおり残る。
    modTestRunner.Check "名寄せ統合_旧が空なら新規がそのまま", _
        (modRagParse.MergeSynPairs("", "回収>リコール|中途解約>解約") = _
         "回収>リコール|中途解約>解約"), _
        "実際=" & modRagParse.MergeSynPairs("", "回収>リコール|中途解約>解約")

    ' 新が空 = 表記ゆれの無い資料。既存の辞書を1行も落とさない。
    modTestRunner.Check "名寄せ統合_新が空なら旧がそのまま", _
        (modRagParse.MergeSynPairs("回収>リコール", "") = "回収>リコール"), _
        "実際=" & modRagParse.MergeSynPairs("回収>リコール", "")

    ' 両方空。
    modTestRunner.Check "名寄せ統合_両方空なら空", _
        (modRagParse.MergeSynPairs("", "") = "")

    ' 【先頭要素の保持】= A-H1 の本体。旧実装は 0始まりの配列を 1〜n で
    ' 読んでいたため、先頭が落ちて末尾で範囲外エラーになっていた。
    Dim r1 As String
    r1 = modRagParse.MergeSynPairs("A>a|B>b|C>c", "D>d")
    modTestRunner.Check "名寄せ統合_旧の先頭要素が落ちない", _
        (r1 = "A>a|B>b|C>c|D>d"), "実際=" & r1

    ' 既存termは上書き: 同じ term を持つ旧行は落とし、新しい行が末尾へ。
    Dim r2 As String
    r2 = modRagParse.MergeSynPairs("A>a|B>b|C>c", "B>bb")
    modTestRunner.Check "名寄せ統合_既存termは上書き(旧行が消える)", _
        (r2 = "A>a|C>c|B>bb"), "実際=" & r2

    ' 順序安定: 残った旧行の並びは変わらない。
    Dim r3 As String
    r3 = modRagParse.MergeSynPairs("A>a|B>b|C>c|D>d", "C>cc|A>aa")
    modTestRunner.Check "名寄せ統合_残った旧行の順序は変わらない", _
        (r3 = "B>b|D>d|C>cc|A>aa"), "実際=" & r3

    ' 大小は同じ term として扱う(termは検索キーで、大小の違いは表記ゆれ)。
    Dim r4 As String
    r4 = modRagParse.MergeSynPairs("abc>x", "ABC>y")
    modTestRunner.Check "名寄せ統合_termの大小違いは同じ語として上書き", _
        (r4 = "ABC>y"), "実際=" & r4

    ' 壊れた要素(">"無し・どちらか空・空要素)は両側とも1件ずつ捨てる。
    Dim r5 As String
    r5 = modRagParse.MergeSynPairs("A>a|こわれた|>x|y>", "||B>b")
    modTestRunner.Check "名寄せ統合_壊れた要素は捨てて読めた分だけ残す", _
        (r5 = "A>a|B>b"), "実際=" & r5

    ' 同じ term が新CSVの中で重複していたら先勝ち(2行目以降は捨てる)。
    Dim r6 As String
    r6 = modRagParse.MergeSynPairs("", "A>1|A>2|B>3")
    modTestRunner.Check "名寄せ統合_新CSV内の重複termは先勝ち", _
        (r6 = "A>1|B>3"), "実際=" & r6

    ' 前後の空白は落として保存する(シートの表示と照合の両方を揃える)。
    Dim r7 As String
    r7 = modRagParse.MergeSynPairs(" A > a ", "")
    modTestRunner.Check "名寄せ統合_前後の空白は落とす", _
        (r7 = "A>a"), "実際=" & r7

    ' 統合結果は ParseSynResp がそのまま読み戻せる形(Store が書き戻す経路)。
    Dim tms() As String, cns() As String, n As Long
    n = modRagParse.ParseSynResp("<syn>" & r2 & "</syn>", tms, cns)
    modTestRunner.Check "名寄せ統合_結果をParseSynRespで読み戻せる", _
        (n = 3), "実際=" & n
    If n = 3 Then
        modTestRunner.Check "名寄せ統合_読み戻しの先頭がAのまま(0始まり)", _
            (tms(LBound(tms)) = "A" And cns(LBound(cns)) = "a"), _
            "実際=" & tms(LBound(tms)) & ">" & cns(LBound(cns))
    End If
End Sub

' ----------------------------------------------------------------------------
' FA-6: 俯瞰語彙の検知。拾う語と【拾わない語】の両方を固定する。
' ----------------------------------------------------------------------------
Private Sub TestHasGlobalSignal()
    modTestRunner.Check "俯瞰語_全体像は?", _
        (modRagParse.HasGlobalSignal("全体像は?") = True)
    modTestRunner.Check "俯瞰語_この規程を全部教えて", _
        (modRagParse.HasGlobalSignal("この規程を全部教えてください") = True)
    modTestRunner.Check "俯瞰語_すべて教えて", _
        (modRagParse.HasGlobalSignal("補償の種類をすべて教えて") = True)
    modTestRunner.Check "俯瞰語_一覧", _
        (modRagParse.HasGlobalSignal("提出書類の一覧") = True)
    modTestRunner.Check "俯瞰語_どんな種類", _
        (modRagParse.HasGlobalSignal("どんな種類がありますか") = True)
    modTestRunner.Check "俯瞰語_どんなこと", _
        (modRagParse.HasGlobalSignal("この資料にはどんなことが書いてありますか") = True)
    modTestRunner.Check "俯瞰語_概要", _
        (modRagParse.HasGlobalSignal("概要を知りたい") = True)
    modTestRunner.Check "俯瞰語_何が書いて", _
        (modRagParse.HasGlobalSignal("何が書いてある?") = True)
    modTestRunner.Check "俯瞰語_全体の流れ", _
        (modRagParse.HasGlobalSignal("申請の全体の流れは?") = True)
    modTestRunner.Check "俯瞰語_どういう構成", _
        (modRagParse.HasGlobalSignal("どういう構成ですか") = True)
    modTestRunner.Check "俯瞰語_まとめて", _
        (modRagParse.HasGlobalSignal("免責事項をまとめて知りたい") = True)

    ' 【拾わない】: 単独で普通に使う短語で誤爆させない。ここが緩むと、
    ' ふつうの条文質問まで俯瞰扱いになり、聞き返しも効かなくなる。
    modTestRunner.Check "俯瞰語_『全て』単独では拾わない", _
        (modRagParse.HasGlobalSignal("全ての書類を提出しますか") = False)
    modTestRunner.Check "俯瞰語_『全部』単独では拾わない", _
        (modRagParse.HasGlobalSignal("全部で何日かかりますか") = False)
    modTestRunner.Check "俯瞰語_条文の質問は拾わない", _
        (modRagParse.HasGlobalSignal("第5条の免責は?") = False)
    modTestRunner.Check "俯瞰語_空文字は拾わない", _
        (modRagParse.HasGlobalSignal("") = False)
    modTestRunner.Check "俯瞰語_空白だけは拾わない", _
        (modRagParse.HasGlobalSignal("   ") = False)
End Sub

' ----------------------------------------------------------------------------
' FA-6: 段0(分解判定)を呼ぶ門に俯瞰シグナルをORで足したこと。
'   実装(modAskMulti.DecomposeGate)とテストが同じ1本を見る(憲章§4-5)。
' ----------------------------------------------------------------------------
Private Sub TestDecomposeGateGlobal()
    ' 6字の俯瞰質問。文字数ゲート(既定25)も複合シグナルも通らないが、
    ' 俯瞰シグナルで段0が呼ばれる=俯瞰が試される。
    modTestRunner.Check "段0門_短い俯瞰質問でも呼ばれる", _
        (modAskMulti.DecomposeGate("auto", "全体像は?", 25) = True)
    modTestRunner.Check "段0門_一覧を求める短文でも呼ばれる", _
        (modAskMulti.DecomposeGate("auto", "補償の一覧", 25) = True)

    ' 俯瞰でも複合でもない短文は従来どおり呼ばない(無駄な段0を増やさない)。
    modTestRunner.Check "段0門_ふつうの短文は従来どおり呼ばない", _
        (modAskMulti.DecomposeGate("auto", "更新日は?", 25) = False)

    ' off は最優先(機能まるごとのエスケープハッチ)。
    modTestRunner.Check "段0門_offは俯瞰シグナルより優先", _
        (modAskMulti.DecomposeGate("off", "全体像は?", 25) = False)

    ' 複合シグナルのORは従来どおり効いたまま(片方を足して片方を壊さない)。
    modTestRunner.Check "段0門_複合シグナルのORは不変", _
        (modAskMulti.DecomposeGate("auto", "免責は?保険料は?", 25) = True)
End Sub

Public Sub RunAll18()
    On Error GoTo MergeFail18
    TestMergeSynPairs
NextGlobal18:
    On Error GoTo GlobalFail18
    TestHasGlobalSignal
NextGate18:
    On Error GoTo GateFail18
    TestDecomposeGateGlobal
NextDone18:
    On Error GoTo 0
    Exit Sub

MergeFail18:
    modTestRunner.Check "TestMergeSynPairs(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGlobal18
GlobalFail18:
    modTestRunner.Check "TestHasGlobalSignal(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextGate18
GateFail18:
    modTestRunner.Check "TestDecomposeGateGlobal(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone18
End Sub
