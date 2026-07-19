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
' 無効/不在/失敗 -> "#ERR:FEATURE_UNAVAILABLE" を返し、呼び出しUI側が
' 「この機能は現在利用できません(管理者が有効化すると使えます)」を表示する。
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
    Dim result As Variant
    result = modGateway.TryRibbonRun(modName & "." & procName, args)
    If VarType(result) = vbString Then
        If Left$(CStr(result), 5) = "#ERR:" Then
            InvokeFeature = "#ERR:FEATURE_UNAVAILABLE"
            Exit Function
        End If
    End If
    InvokeFeature = result
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' featureId -> モジュール名 対応表(MASTER_SPEC §7.1)。
Private Function ModuleNameOf(ByVal featureId As String) As String
    Select Case featureId
        Case "tts": ModuleNameOf = "optTts"
        Case "vision": ModuleNameOf = "optVision"
        Case "markdown": ModuleNameOf = "optMarkdown"
        Case "diffdoc": ModuleNameOf = "optDiffDoc"
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
