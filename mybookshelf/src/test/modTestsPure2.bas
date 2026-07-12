Attribute VB_Name = "modTestsPure2"
Option Explicit

' ============================================================================
' modTestsPure2 - modTestsPureの分割先(MASTER_SPEC §7.1「1モジュール30,000字
'   以内」を超えたための分割。§7.8のテスト範囲そのものは変わらない)
' ----------------------------------------------------------------------------
' 役割:
'   modPrompts / modShelfSync.DiffDecision / modPack.ValidatePackMeta の
'   純ロジック部のテストをここに置く(modUtil/modChunker/modPiiは
'   modTestsPure.bas側)。入口は modTestsPure.RunAll の末尾から呼ばれる
'   Public Sub RunAll2()。modTestRunner.RunAllPureTests は modTestsPure.RunAll
'   だけを呼ぶ契約(modTestRunner.bas §7.8。担当外につき変更しない)なので、
'   本モジュールへの導線は modTestsPure.RunAll 内に置く。
'
' 設計判断(R4準拠・グループ単位の失敗隔離): modTestsPure.bas冒頭コメントと
'   同じ方針(1グループの想定外エラーが他グループを道連れにしない)。
'
' ■ CanUseTypeArraysの複製について:
'   modTestsPure.bas の Private Function CanUseTypeArrays() は別モジュールの
'   Privateであるため、ここから呼べない。判定ロジック自体は3行程度の軽量な
'   実測プローブなので、モジュールをまたいだ複製を許容する(共有したいだけの
'   ために新たにPublicの公開契約を増やすより、テストモジュール限定の軽微な
'   重複の方が実害が小さいと判断)。挙動・コメントはmodTestsPure.bas側の
'   オリジナルと同一にしてある。
' ============================================================================

Public Sub RunAll2()
    On Error GoTo PromptsFail
    TestModPrompts
NextShelfSync:
    On Error GoTo ShelfSyncFail
    TestModShelfSync
NextPack:
    On Error GoTo PackFail
    TestModPack
NextDone:
    On Error GoTo 0
    Exit Sub

PromptsFail:
    modTestRunner.Check "TestModPrompts(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextShelfSync
ShelfSyncFail:
    modTestRunner.Check "TestModShelfSync(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextPack
PackFail:
    modTestRunner.Check "TestModPack(グループ全体)", False, "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone
End Sub

' modTestsPure.bas の CanUseTypeArrays と同一実装(モジュール冒頭コメント参照)。
Private Function CanUseTypeArrays() As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe() As ShelfChunk
    ReDim probe(0 To 0)
    CanUseTypeArrays = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ============================================================================
' modPrompts
' ============================================================================
Private Sub TestModPrompts()
    ' BuildEnrichPromptはHit()配列を取らないため、LO実行環境のPublic Type配列の
    ' 制限(modTestsPure.bas冒頭コメント参照)の影響を受けない。常にフル検証する。
    TestBuildEnrichPrompt

    If Not CanUseTypeArrays() Then
        modTestRunner.Check "modPrompts(Hit配列を使う3関数): LO環境の既知の制限によりスキップ", True, _
            "BuildQuickPrompt/BuildDeepDraftPrompt/BuildDeepVerifyPromptはHit()配列を" & _
            "引数に取るため、LibreOffice実行環境のPublic Type配列制限(modTestsPure.bas" & _
            "冒頭コメント参照)の影響を受ける。出典形式([本棚:.. p.N] / [パック(作成者):..])・" & _
            "full_text根拠の使用・max_context_chars打切り時の「(一部省略)」挿入・" & _
            "answer_language既定値挿入はコードレビューで確認済みだが、" & _
            "Excel実機受入チェック(§11.3)で必ず再確認すること。"
        Exit Sub
    End If

    TestBuildQuickPrompt_CitationAndLanguage
    TestBuildQuickPrompt_FullTextIsUsedAsBody
    TestBuildQuickPrompt_Truncation
    TestBuildDeepDraftAndVerify_Citation
End Sub

Private Sub TestBuildEnrichPrompt()
    Dim r As String
    r = modPrompts.BuildEnrichPrompt("チャンク1本文" & vbLf & "チャンク2本文")
    modTestRunner.Check "BuildEnrichPrompt_バッチ本文を含む", (InStr(r, "チャンク1本文") > 0), "r=" & r
    modTestRunner.Check "BuildEnrichPrompt_JSON形式の指示を含む", (InStr(r, "summary") > 0) And (InStr(r, "keywords") > 0), "r=" & r
    ' answer_language既定の挿入(modConfig未接続でも既定値"日本語"で動くこと)
    modTestRunner.Check "BuildEnrichPrompt_既定言語_日本語", (InStr(r, "日本語") > 0), "r=" & r
End Sub

Private Sub TestBuildQuickPrompt_CitationAndLanguage()
    Dim h(0 To 1) As Hit
    h(0).chunk_id = "bs::aaa::p3::c1": h(0).score = 0.9
    h(0).source = "ファイルA.pdf": h(0).page = 3
    h(0).preview = "ここに本文の抜粋が入ります。": h(0).origin = "self"

    h(1).chunk_id = "bs::bbb::p1::c1": h(1).score = 0.8
    h(1).source = "ファイルB.docx": h(1).page = 1
    h(1).preview = "別の抜粋です。": h(1).origin = "pack:山田太郎"

    Dim r As String
    r = modPrompts.BuildQuickPrompt("何か質問", h, 2)

    modTestRunner.Check "BuildQuickPrompt_本棚出典形式", (InStr(r, "[本棚:ファイルA.pdf p.3]") > 0), "r=" & r
    modTestRunner.Check "BuildQuickPrompt_パック出典形式", (InStr(r, "[パック(山田太郎):ファイルB.docx]") > 0), "r=" & r
    ' answer_language既定の挿入(modConfig未接続でも既定値"日本語"で動くこと)
    modTestRunner.Check "BuildQuickPrompt_既定言語_日本語", (InStr(r, "日本語") > 0), "r=" & r
    modTestRunner.Check "BuildQuickPrompt_質問文を含む", (InStr(r, "何か質問") > 0), "r=" & r
    ' 打ち切っていない通常ケースでは「(一部省略)」は出ないこと
    modTestRunner.Check "BuildQuickPrompt_非打切り時は省略表記なし", (InStr(r, "(一部省略)") = 0), "r=" & r
End Sub

' Wave3 PM裁定1(Hit.full_text追加)の検証: 本棚抜粋ブロックの本文には
' preview(先頭120字・出典先出し表示専用)ではなく full_text(チャンク本文
' 全体)が使われること。preview と full_text に別々の目印文字列を仕込み、
' 出力に full_text 側の目印だけが含まれ、preview 側の目印は含まれないことを
' 確認する(modPrompts.SourceBodyがfull_text優先・空時のみpreviewへ
' フォールバックする設計であることのテスト。§7.3/modTypes.Hit参照)。
Private Sub TestBuildQuickPrompt_FullTextIsUsedAsBody()
    Dim h(0 To 0) As Hit
    h(0).chunk_id = "bs::ft::p9::c1": h(0).score = 0.9
    h(0).source = "全文根拠資料.pdf": h(0).page = 9
    h(0).preview = "PREVIEW_ONLY_MARKER_短い先頭抜粋"
    h(0).full_text = "FULLTEXT_MARKER_これがチャンク本文全体の根拠テキストです。" & _
        "本来はここに数百字の実際の抜粋が入る想定。"
    h(0).origin = "self"

    Dim r As String
    r = modPrompts.BuildQuickPrompt("何か質問", h, 1)

    modTestRunner.Check "BuildQuickPrompt_full_text根拠が含まれる", _
        (InStr(r, "FULLTEXT_MARKER_これがチャンク本文全体の根拠テキストです。") > 0), "r=" & r
    modTestRunner.Check "BuildQuickPrompt_full_text優先時はpreviewを使わない", _
        (InStr(r, "PREVIEW_ONLY_MARKER") = 0), "r=" & r
    modTestRunner.Check "BuildQuickPrompt_full_text使用時も本棚出典形式", _
        (InStr(r, "[本棚:全文根拠資料.pdf p.9]") > 0), "r=" & r

    ' 防御的フォールバック確認: full_textが空(旧データ・テストダブル等)なら
    ' previewへフォールバックすること。
    Dim h2(0 To 0) As Hit
    h2(0).chunk_id = "bs::fb::p1::c1": h2(0).score = 0.5
    h2(0).source = "フォールバック資料.pdf": h2(0).page = 1
    h2(0).preview = "PREVIEW_FALLBACK_MARKER"
    h2(0).full_text = ""   ' 空
    h2(0).origin = "self"

    Dim r2 As String
    r2 = modPrompts.BuildQuickPrompt("別の質問", h2, 1)
    modTestRunner.Check "BuildQuickPrompt_full_text空時はpreviewへフォールバック", _
        (InStr(r2, "PREVIEW_FALLBACK_MARKER") > 0), "r2=" & r2
End Sub

Private Sub TestBuildQuickPrompt_Truncation()
    ' previewを非常に長くして確実にmax_context_chars(既定40000。modConfig未接続時も
    ' SafeMaxContextCharsの既定値40000が使われる)を超えさせ、打ち切りと
    ' 「(一部省略)」挿入を検証する。
    Dim bigPreview As String: bigPreview = String(30000, "x")
    Dim h(0 To 1) As Hit
    h(0).chunk_id = "bs::a::p1::c1": h(0).score = 0.9
    h(0).source = "大きい資料.pdf": h(0).page = 1
    h(0).preview = bigPreview: h(0).origin = "self"

    h(1).chunk_id = "bs::b::p1::c1": h(1).score = 0.8
    h(1).source = "大きい資料2.pdf": h(1).page = 1
    h(1).preview = bigPreview: h(1).origin = "self"

    Dim r As String
    r = modPrompts.BuildQuickPrompt("質問", h, 2)

    modTestRunner.Check "BuildQuickPrompt_打切り時は省略表記あり", (InStr(r, "(一部省略)") > 0), _
        "len(r)=" & Len(r)
    ' 打ち切られているので2件目の資料名までは含まれない(先頭60000字級のpreviewの後半は
    ' コンテキストブロックに入りきらないはず)。少なくとも全文がそのまま連結されて
    ' いない(=本当に打ち切られている)ことを、長さの上限チェックで確認する。
    modTestRunner.Check "BuildQuickPrompt_本文が無制限に伸びていない", (Len(r) < (Len(bigPreview) * 2 + 5000)), _
        "len(r)=" & Len(r)
End Sub

Private Sub TestBuildDeepDraftAndVerify_Citation()
    Dim h(0 To 0) As Hit
    h(0).chunk_id = "bs::c::p2::c1": h(0).score = 0.7
    h(0).source = "資料C.xlsx": h(0).page = 2
    h(0).preview = "抜粋C": h(0).origin = "self"

    Dim rDraft As String
    rDraft = modPrompts.BuildDeepDraftPrompt("質問2", h, 1, "")
    modTestRunner.Check "BuildDeepDraftPrompt_出典形式", (InStr(rDraft, "[本棚:資料C.xlsx p.2]") > 0), "rDraft=" & rDraft
    modTestRunner.Check "BuildDeepDraftPrompt_既定言語_日本語", (InStr(rDraft, "日本語") > 0), "rDraft=" & rDraft

    Dim rVerify As String
    rVerify = modPrompts.BuildDeepVerifyPrompt("質問2", "下書き回答本文", h, 1)
    modTestRunner.Check "BuildDeepVerifyPrompt_出典形式", (InStr(rVerify, "[本棚:資料C.xlsx p.2]") > 0), "rVerify=" & rVerify
    modTestRunner.Check "BuildDeepVerifyPrompt_下書きを含む", (InStr(rVerify, "下書き回答本文") > 0), "rVerify=" & rVerify
End Sub

' ============================================================================
' modShelfSync.DiffDecision(純関数)
' ============================================================================
Private Sub TestModShelfSync()
    modTestRunner.Check "DiffDecision_新規", _
        (modShelfSync.DiffDecision(False, False, False) = "ingest")
    modTestRunner.Check "DiffDecision_新規_サイズ時刻フラグ無視", _
        (modShelfSync.DiffDecision(False, True, True) = "ingest")
    modTestRunner.Check "DiffDecision_サイズ変化", _
        (modShelfSync.DiffDecision(True, True, False) = "replace")
    modTestRunner.Check "DiffDecision_時刻変化", _
        (modShelfSync.DiffDecision(True, False, True) = "replace")
    modTestRunner.Check "DiffDecision_サイズ時刻両方変化", _
        (modShelfSync.DiffDecision(True, True, True) = "replace")
    modTestRunner.Check "DiffDecision_不変", _
        (modShelfSync.DiffDecision(True, False, False) = "keep")
End Sub

' ============================================================================
' modPack.ValidatePackMeta(純関数)
' ============================================================================
Private Sub TestModPack()
    Dim reason As String

    ' 正常(バージョン一致・次元1以上)
    Dim okNormal As Boolean
    okNormal = modPack.ValidatePackMeta(modAppDef.PACK_FORMAT_VERSION, 1536, reason)
    modTestRunner.Check "ValidatePackMeta_正常", okNormal, "reason=" & reason

    ' バージョン不一致
    reason = ""
    Dim okVerMismatch As Boolean
    okVerMismatch = modPack.ValidatePackMeta(modAppDef.PACK_FORMAT_VERSION + 1, 1536, reason)
    modTestRunner.Check "ValidatePackMeta_バージョン不一致False", (Not okVerMismatch), "reason=" & reason
    modTestRunner.Check "ValidatePackMeta_バージョン不一致理由あり", (Len(reason) > 0), "reason=" & reason

    ' 次元不一致(0以下は不正)
    reason = ""
    Dim okDimInvalid As Boolean
    okDimInvalid = modPack.ValidatePackMeta(modAppDef.PACK_FORMAT_VERSION, 0, reason)
    modTestRunner.Check "ValidatePackMeta_次元0はFalse", (Not okDimInvalid), "reason=" & reason

    reason = ""
    Dim okDimNegative As Boolean
    okDimNegative = modPack.ValidatePackMeta(modAppDef.PACK_FORMAT_VERSION, -5, reason)
    modTestRunner.Check "ValidatePackMeta_次元負値はFalse", (Not okDimNegative), "reason=" & reason
End Sub
