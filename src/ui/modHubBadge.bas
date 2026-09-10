Attribute VB_Name = "modHubBadge"
Option Explicit

' ============================================================================
' modHubBadge - Hubのバッジ帯(獲得済みだけ色を変える)の実体(R43 2-10)。
' ----------------------------------------------------------------------------
' 背景: modHub.DrawBadgesは獲得済み🏅/未獲得🔒の記号を1本の文字列に連結し、
'   結合セル1個(B(r+1):F(r+br))へまとめて書いていた。文字色は帯全体で
'   同じ("text")のため、獲得済みかどうかは記号(🏅/🔒)だけでしか分からず、
'   離れて見ると全バッジが同じ濃さに見える(実機報告)。
'   modHubは29,551字(直接書くと30,002字で上限超過)・modHubStatは残77字
'   (受け皿にできない)ため、本モジュールを新設して実体を持つ。
'
' 方式: 帯全体をmuted(薄色)にしたうえで、獲得済みの記号+資料名の区間だけ
'   Characters(start, length) で text色+太字に塗り直す(未獲得は触らない
'   =薄いまま)。結合セルへの Characters は【左上セル1個】に対してのみ
'   行う(複数セルにまたがる Range.Characters は実 Excel で 1004 になる。
'   modHub.DrawBadgesはB列を左上として結合するので、対象は Cells(topRow,"B")
'   固定)。
'
' 手順の順序(唯一の落とし穴): .Value → 帯全体の色(muted) → Characters。
'   Characters を先に当てると直後の .Value 代入で消える。帯全体の色より
'   前に Characters を当てても、あとから帯色で一括上書きされて消える。
'
' 文字位置の算出(BadgeSpans・純関数): 呼び出し元(modHub.DrawBadges)の
'   組み立て規約「mark(2文字・🏅/🔒はどちらもサロゲートペア2文字) & " "
'   (1文字) & title & "   "(区切り3文字)」を1文字も変えずに踏襲した位置
'   計算。Excelオブジェクトに一切触れないため modTestsPure47 が直接固定する。
' ============================================================================

' BadgeSpans - 獲得済みバッジの文字区間(開始位置・長さ)を返す純関数。
'   区間は「mark(2)+空白(1)+資料名」までで、末尾の区切り空白3個は含めない
'   (太字化する意味が無い装飾の間隔のため)。戻り値は区間の件数。
'   titles/earnedはn件ぶん(0 To n-1)。n<1は0を返し配列に触れない。
Public Function BadgeSpans(ByRef titles() As String, ByRef earned() As Boolean, _
                            ByVal n As Long, ByRef starts() As Long, ByRef lens() As Long) As Long
    If n < 1 Then Exit Function
    ReDim starts(0 To n - 1)
    ReDim lens(0 To n - 1)
    Dim cnt As Long
    Dim pos As Long: pos = 1
    Dim i As Long
    For i = 0 To n - 1
        Dim segLen As Long: segLen = 3 + Len(titles(i))   ' mark(2)+空白(1)+資料名
        If earned(i) Then
            starts(cnt) = pos
            lens(cnt) = segLen
            cnt = cnt + 1
        End If
        pos = pos + segLen + 3                             ' 区切りの半角空白3個
    Next i
    BadgeSpans = cnt
End Function

' PaintEarnedSpans - 獲得済み区間だけをtext色+太字で塗り直す(未獲得は
'   帯全体のmutedのまま)。topRow/topColは結合セルの左上セル(modHub.
'   DrawBadgesの結合は"B"列始まり)。Characters関連は1個の失敗で残りを
'   道連れにしないOn Error Resume Next配下。
Public Sub PaintEarnedSpans(ByVal ws As Worksheet, ByVal topRow As Long, _
                             ByRef titles() As String, ByRef earned() As Boolean, ByVal n As Long)
    On Error Resume Next
    Dim starts() As Long, lens() As Long
    Dim cnt As Long: cnt = BadgeSpans(titles, earned, n, starts, lens)
    If cnt < 1 Then Exit Sub
    Dim topCell As Range: Set topCell = ws.Cells(topRow, "B")
    Dim i As Long
    For i = 0 To cnt - 1
        With topCell.Characters(starts(i), lens(i)).Font
            .Color = modUI.UiColor("text")
            .Bold = True
        End With
    Next i
    On Error GoTo 0
End Sub
