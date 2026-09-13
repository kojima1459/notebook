Attribute VB_Name = "modFeatures"
Option Explicit

' ============================================================================
' modFeatures - opt機能(読み上げ/画像PDF/Markdown表示/約款差分)の分離実行
' ----------------------------------------------------------------------------
' 役割:
'   opt層モジュール(optTts等)への参照は、コアのどのモジュールにも直接
'   書かせない(MASTER_SPEC §3 R2)。呼び出しは全てここを経由し、
'   文字列("opt" & 機能名)からApplication.Runで遅延バインドする。
'   これにより、opt機能の撤去は「ビルドリストから1行削除+configの
'   feature_xxxフラグをFALSE」だけで完結し、コア側のコードは無傷のまま
'   保てる(§1のDoD5)。
'
' 設計判断:
'   ・featureId→モジュール名の対応表はここ(modFeatures内Private)に
'     持つ(modAppDefではない。MASTER_SPEC §7.1の指示どおり)。
'   ・ModulePresentの結果はfeatureIdごとにキャッシュする。opt側の
'     Application.Run "optX.Ping" は「モジュールが実在しビルドに
'     同梱されているか」を確認するためだけの軽い呼び出しであり、
'     何度も呼ぶ必要が無いため。
'   ・4つのfeatureIdは固定(tts/vision/markdown/diffdoc)なので、汎用的な
'     文字列キャッシュではなく、featureId毎に専用のPrivate変数を持つ
'     単純な実装にした(コードは長くなるが、バグの温床になりやすい
'     文字列パース処理を避けられる)。
' ============================================================================

Private mChkTts As Boolean, mValTts As Boolean
Private mChkVision As Boolean, mValVision As Boolean
Private mChkMarkdown As Boolean, mValMarkdown As Boolean
Private mChkDiffdoc As Boolean, mValDiffdoc As Boolean

' config "feature_<id>"=TRUE かつ ModulePresent の両方を満たす時のみTrue。
Public Function FeatureEnabled(ByVal featureId As String) As Boolean
    Dim flagName As String: flagName = "feature_" & LCase$(featureId)
    Dim onFlag As Boolean: onFlag = modConfig.GetBool(flagName, False)
    FeatureEnabled = onFlag And ModulePresent(featureId)
End Function

' Application.Run "opt<Pascal>.Ping" をOn Errorで試す(結果はキャッシュ)。
' opt側はPublic Function Ping() As Booleanを必ず持つ契約(§7.7)。
Public Function ModulePresent(ByVal featureId As String) As Boolean
    Dim key As String: key = LCase$(featureId)

    Dim already As Boolean, cachedVal As Boolean
    GetCache key, already, cachedVal
    If already Then
        ModulePresent = cachedVal
        Exit Function
    End If

    Dim modName As String: modName = ModuleNameOf(key)
    Dim ok As Boolean: ok = False
    If LenB(modName) > 0 Then
        Dim errNum As Long
        On Error Resume Next
        Dim res As Variant: res = Application.Run(modName & ".Ping")
        errNum = Err.Number
        On Error GoTo 0
        If errNum = 0 Then ok = CBool(res)
    End If

    SetCache key, ok
    ModulePresent = ok
End Function

' Application.Run("opt<Pascal>." & procName, ...) 遅延バインド。
' 機能が無効/モジュール不在 -> "#ERR:FEATURE_UNAVAILABLE"(呼び出しUI側が
' 「この機能は現在利用できません(管理者が有効化すると使えます)」を表示する)。
' 呼べたが失敗した場合 -> opt側が返した "#ERR:..." をそのまま返す(R6要件9)。
Public Function InvokeFeature(ByVal featureId As String, ByVal procName As String, ByVal args As Variant) As Variant
    If Not FeatureEnabled(featureId) Then
        InvokeFeature = "#ERR:FEATURE_UNAVAILABLE"
        Exit Function
    End If

    Dim modName As String: modName = ModuleNameOf(LCase$(featureId))
    If LenB(modName) = 0 Then
        InvokeFeature = "#ERR:FEATURE_UNAVAILABLE"
        Exit Function
    End If

    ' TryRibbonRunの「Variant配列を展開してApplication.Runする」ロジックを
    ' そのまま再利用する(呼び先がリボンではなくoptモジュールという違いのみ)。
    '
    ' 2026-07-31(R6要件9): opt側が返した "#ERR:..." を一律
    ' "#ERR:FEATURE_UNAVAILABLE" へ潰すのをやめ、元の文字列をそのまま返す。
    ' 潰していたせいで「Ghostscriptが見つからないので社内ポータルのzipから
    ' 置いてください」のような、利用者がその場で打てる一手の案内が全部消えて
    ' 「この機能は現在利用できません」だけになっていた。"#ERR:" の接頭辞は
    ' 維持されるため、失敗を判定している既存の呼び出し元はすべて互換。
    InvokeFeature = modGateway.TryRibbonRun(modName & "." & procName, args)
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' featureId -> モジュール名 対応表(MASTER_SPEC §7.1)。
Private Function ModuleNameOf(ByVal featureId As String) As String
    Select Case featureId
        ' 2026-07-28(レビュー I-15): optTts は実体が無い(AIリボンが読み上げAPIを
        ' 公開していないため作れなかった)。config feature_tts は既定FALSEで
        ' 握り潰されるので実害は無いが、対応表に載っていると「あるはず」に
        ' 見える。名前は残しつつ、無いことをここに明記する。
        Case "tts": ModuleNameOf = "optTts"   ' 実体なし(feature_tts=False固定)
        Case "vision": ModuleNameOf = "optVision"
        Case "markdown": ModuleNameOf = "optMarkdown"
        ' 2026-08-16(R33 W4-4): diffdoc は【同梱されているのに動かない】。
        ' optDiffDoc.bas はビルド対象(build/modules.json)に載っていて
        ' 注入もされ、config の feature_diffdoc が有効なら診断画面にも
        ' 「モジュールあり/設定有効」と出る。しかし src 全体を探しても
        ' InvokeFeature("diffdoc", …) も FeatureEnabled("diffdoc") も
        ' 呼び出しが0件で、利用者から起動する導線がどの画面にも無い
        ' (modUIShelf.bas:37-43 が「仕様に明記の無いUI追加はしない」と
        ' 判断してボタンを置かず、config キー側だけが残った)。
        ' 導線を足すのは機能追加なので今ラウンドの範囲外。宿題は
        ' docs/dev/TODO.md「F. R33からの持ち越し」に記録してある。
        Case "diffdoc": ModuleNameOf = "optDiffDoc"   ' 呼び出し口0件(導線なし)
        Case Else: ModuleNameOf = ""
    End Select
End Function

Private Sub GetCache(ByVal key As String, ByRef already As Boolean, ByRef v As Boolean)
    Select Case key
        Case "tts"
            already = mChkTts: v = mValTts
        Case "vision"
            already = mChkVision: v = mValVision
        Case "markdown"
            already = mChkMarkdown: v = mValMarkdown
        Case "diffdoc"
            already = mChkDiffdoc: v = mValDiffdoc
        Case Else
            already = True: v = False   ' 未知のfeatureIdは常に不在扱い
    End Select
End Sub

Private Sub SetCache(ByVal key As String, ByVal v As Boolean)
    Select Case key
        Case "tts"
            mChkTts = True: mValTts = v
        Case "vision"
            mChkVision = True: mValVision = v
        Case "markdown"
            mChkMarkdown = True: mValMarkdown = v
        Case "diffdoc"
            mChkDiffdoc = True: mValDiffdoc = v
    End Select
End Sub
