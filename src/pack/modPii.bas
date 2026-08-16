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
'   ・走査の入口で全角(数字・＠・ハイフン類)だけを半角へ均す(NormalizeWidth)。
'     判定閾値や区切り規則そのものは変えない。理由は NormalizeWidth 参照。
' ============================================================================

' ScanText - 検知した種別をカンマ区切りで返す(検知無しは空文字列)。
Public Function ScanText(ByVal s As String) As String
    If LenB(s) = 0 Then Exit Function

    ' R33 W2-1: 判定に渡す前に幅だけを均す(閾値・区切り規則は不変)。
    s = NormalizeWidth(s)

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

' ----------------------------------------------------------------------------
' NormalizeWidth - 走査前に全角の数字・＠・ハイフン類だけを半角へ均す(純関数)。
' ----------------------------------------------------------------------------
'   【直したこと】(2026-08-15 R33 W2-1)
'   日本語IMEを全角のまま質問・資料を書くと、電話番号は
'   "０９０－１２３４－５６７８"、メールは "taro＠example.co.jp" のように
'   全角で入る。ところが HasLongDigitRun の数字判定は Binary比較の
'   `ch >= "0" And ch <= "9"` なので、全角数字(U+FF10〜FF19=65296〜65305)は
'   "9"(57) より大きく False になる。しかも区切り文字にも該当しないので
'   Else へ落ちてランを【リセット】する ―― 全角の電話番号は1桁も数えられない。
'   HasEmailPattern も半角@しか見ないので全角＠も素通りする。
'   結果、氏名+携帯番号を含む文面がPII検知を素通りして共有経路へ流れていた。
'
'   【なぜ入口で均すだけに留めるのか】判定閾値(10桁)と既存の区切り規則を
'   変えると、R32波2の裁定「PII緩和ロジックはオンに戻す際に再検討」を
'   先回りして踏み越えることになる。よってここは「半角で書かれていれば
'   従来どおり検知できたはずの文面を、全角というだけで見逃さない」ための
'   幅合わせに限定する。半角入力に対する挙動は1文字も変わらない。
'
'   均す対象(spec R33 W2-1):
'     ・U+FF10〜U+FF19(全角数字)             -> "0"〜"9"
'     ・U+FF20(全角＠)                       -> "@"
'     ・U+FF0D/U+2212/U+30FC/U+FF5E          -> "-"(既存の区切り文字へ寄せる)
'   全角ハイフン類を "-" へ寄せるのは、既存の区切り規則(ランを継続する)を
'   そのまま再利用するため。区切り文字の一覧そのものは変えていない。
'
'   非ASCIIは必ず ChrW で組む(CP932 では U+FF0D と U+2212 が同じバイトへ
'   潰れうるため、ソースへ生の文字を置くと注入後にどちらか一方が消える)。
'
'   【Public である理由】(2026-08-16 R33 W2-7)
'   共有発信側の関所 modInsightGate.PiiBlocked は
'   「幅を均す → StripDateLike → ScanText」の順で呼ぶ。StripDateLike は
'   半角数字専用なので、ScanText の【内側】でだけ均していると、全角で
'   書かれた日付が日付潰しを素通りしたまま12桁ランとして検知され、
'   その質問が無言で共有見送りになる(R32 F4 が半角側で潰した誤検知の
'   全角版)。幅の正規化は関所の【入口に1回】置くのが正しい位置なので、
'   前処理の並びの先頭から呼べるよう Public にしてある(§7契約表にも登録)。
'   ScanText 内側の均しは残す ―― 本関数は冪等(出力に全角は残らない)なので
'   二重に通しても結果は変わらず、modPii を直接叩く他の経路
'   (modPackExport.ScanChunksForPii)を守り続けられる。
' ----------------------------------------------------------------------------
Public Function NormalizeWidth(ByVal s As String) As String
    Dim t As String: t = s

    Dim d As Long
    For d = 0 To 9
        t = Replace(t, ChrW(&HFF10& + d), Chr$(48 + d))
    Next d

    t = Replace(t, ChrW(&HFF20&), "@")
    t = Replace(t, ChrW(&HFF0D&), "-")
    t = Replace(t, ChrW(&H2212&), "-")
    t = Replace(t, ChrW(&H30FC&), "-")
    t = Replace(t, ChrW(&HFF5E&), "-")

    NormalizeWidth = t
End Function

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
