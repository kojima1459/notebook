Attribute VB_Name = "modPii"
Option Explicit

' ============================================================================
' modPii - パック書き出し前の簡易PII(個人情報らしき文字列)検知(純ロジックR4)
' ----------------------------------------------------------------------------
' 役割:
'   modPack.ExportPackDialog が書き出し前に全チャンクへ適用する、ローカルの
'   簡易ヒューリスティック検知。ブロックはしない(警告であり停止ではない。
'   MASTER_SPEC §6 E0703)。検知した種別をカンマ区切りで返し、呼び出し側の
'   UIが件数・例示を出して続行確認する。
'
' 流用元:
'   /home/user/notebook/src/chatbot_v2/modPii.bas (DetectPii)。
'   ロジック(メールらしきパターン・10桁以上の連続数字)はそのまま踏襲しつつ、
'   MASTER_SPEC §7.4 の契約シグネチャ
'   `Public Function ScanText(ByVal s As String) As String` に合わせて、
'   Booleanの単純検知から「検知種別のカンマ列挙 or 空文字列」を返す形へ
'   適応した(呼び出し側が件数だけでなく「何が検知されたか」を案内できる
'   ようにするため)。
'
' 設計判断(R4: Excelオブジェクト禁止):
'   ・Worksheets/Range/Application/ThisWorkbook/MsgBox には一切触れない。
'     文字列処理のみの純関数。tools/run_lo_tests.py のモード1
'     (RunAllPureTests)対象。
'   ・§7.8 は「pack検証」だけでなくPII検知もmodTestsPureから直接呼べる
'     ことを前提にしている(Excel非依存)ため、本モジュール全体を
'     R4準拠のまま保つ。
'   ・検知は「確実な個人情報検出」ではなく、あくまで書き出し前の注意喚起
'     (advisory)。誤検知(false positive)は安全側(検知多め)で許容する
'     設計とし、逆に見逃し(false negative)を過度に恐れて複雑な正規表現を
'     持ち込むことはしない(VBAの正規表現は標準では使えないため、Like演算子
'     と文字走査だけで完結させるV2の方針を踏襲)。
'   ・区切り文字を含む連続数字(電話番号のハイフン区切り等)を1つのランとして
'     数えるため、"-"・半角スペース・全角相当のハイフン(U+2010)は
'     ランを継続する区切りとして扱う(V2 modPii.bas と同じ扱い)。
' ============================================================================

' ScanText - 検知した種別をカンマ区切りで返す(検知無しは空文字列)。
Public Function ScanText(ByVal s As String) As String
    If LenB(s) = 0 Then Exit Function

    Dim cats() As String: ReDim cats(0 To 1)
    Dim n As Long: n = 0

    If HasEmailPattern(s) Then
        cats(n) = "メールアドレス"
        n = n + 1
    End If

    If HasLongDigitRun(s) Then
        cats(n) = "電話番号やマイナンバー等の長い数字列"
        n = n + 1
    End If

    If n = 0 Then
        ScanText = ""
    Else
        ReDim Preserve cats(0 To n - 1)
        ScanText = Join(cats, ",")
    End If
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

' メールアドレスらしきパターン(V2 modPii.bas の DetectPii をそのまま踏襲)。
Private Function HasEmailPattern(ByVal s As String) As Boolean
    If InStr(s, "@") = 0 Then Exit Function
    HasEmailPattern = (s Like "*[A-Za-z0-9.]@[A-Za-z0-9.]*")
End Function

' 10桁以上の連続する数字(区切り文字を挟んでも継続とみなす)を検知する。
' 電話番号・マイナンバー・契約番号等をまとめてこのヒューリスティックで拾う
' (V2 modPii.bas の DetectPii と同じ設計)。
Private Function HasLongDigitRun(ByVal s As String) As Boolean
    Dim runLen As Long: runLen = 0
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            runLen = runLen + 1
            If runLen >= 10 Then
                HasLongDigitRun = True
                Exit Function
            End If
        ElseIf ch = "-" Or ch = " " Or ch = "‐" Then
            ' 区切り文字はランを継続扱い(桁数はリセットしない)
        Else
            runLen = 0
        End If
    Next i
End Function
