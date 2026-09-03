Attribute VB_Name = "modInstallCheck"
Option Explicit

' ============================================================================
' modInstallCheck - 行数照合の純関数だけを持つモジュール
'                   (2026-08-10 R23b で新設・2026-09-03 R35波1で手術)
' ----------------------------------------------------------------------------
' R35 で方式Bへ転換した(spec_20260903_R35_配布方式転換.md)。旧方式(自己
' インストーラが起動のたびに vba_src シートの本文をVBEオブジェクト操作API
' + 旧文字列注入APIで注入する形)は 2026-09-02 に社内AVの AMSI で検知され
' (VDI強制停止)、採用禁止になった(裁定書27)。方式Bはビルド時に完成品の
' vbaProject.bin を書くため、起動時の注入・注入結果の検証(旧 VI /
' ModuleLineCount / WarnBroken と、それが読んでいた vba_src の列定数)は
' 配布物に一切残らない(=ThisWorkbookのVBAプロジェクト参照そのものが消える)。
'
' このモジュールに残っているのは、その検証が使っていた「行数」の数え方を
' 固定する純関数(ExpectedLineCount / LineCountMismatch)だけである。
' ビルド時の期待行数計算(build/build_mybookshelf.py の _expected_line_count)
' と実装が乖離しないことを固定するテスト資産(src/test/modTestsPure23.bas)
' として、実行時には誰にも呼ばれない状態のまま残置する。
'
' 行数比較の仕様(なぜ「行数」で、なぜ「末尾空行を落とす」のか。旧VIの設計
' 判断をそのまま記録): 本文そのものの完全一致ではなく行数にしたのは、
' VBEモジュール本文取得APIで全文を取り出して比較すると巨大な文字列結合が走るうえ、
' VBE側が行末の空白を落とす等の無害な差でも不一致になり、偽陽性を招きかねない
' ため。旧文字列注入APIは末尾の空行の扱いが環境で1本ぶれることが知られていた
' ので、両側とも末尾の空行を落としてから厳密一致で比較する(±1の黙認はしない)。
' ============================================================================

' ----------------------------------------------------------------------------
' LineCountMismatch - 期待ソースと実測行数が食い違うか(純関数)
' ----------------------------------------------------------------------------
' expectedSrc は比較対象のソース本文(改行は vbLf。念のため vbCrLf/vbCr も
' 正規化する)。actualLineCount は呼び出し側が【末尾の空行を落としたあと】の
' 行数を渡す。expectedSrc 側の末尾空行はこの関数が落とす。
' True = 不一致。
' ----------------------------------------------------------------------------
Public Function LineCountMismatch(ByVal expectedSrc As String, ByVal actualLineCount As Long) As Boolean
    LineCountMismatch = (ExpectedLineCount(expectedSrc) <> actualLineCount)
End Function

' ----------------------------------------------------------------------------
' ExpectedLineCount - ソース文字列の「末尾空行を落とした行数」(純関数)
' ----------------------------------------------------------------------------
' 空文字列は 0 行。"a" は 1 行。"a" & vbLf は 1 行(末尾空行を落とす)。
' "a" & vbLf & vbLf も 1 行。"a" & vbLf & "b" は 2 行。
' ----------------------------------------------------------------------------
Public Function ExpectedLineCount(ByVal expectedSrc As String) As Long
    Dim s As String
    s = Replace(expectedSrc, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    If LenB(s) = 0 Then Exit Function

    Dim parts As Variant
    parts = Split(s, vbLf)

    Dim lo As Long
    lo = LBound(parts)
    Dim n As Long
    n = UBound(parts) - lo + 1

    ' 末尾から空行(空白・タブだけの行を含む)を落とす。
    Do While n > 0
        If Not IsBlankText(CStr(parts(lo + n - 1))) Then Exit Do
        n = n - 1
    Loop
    ExpectedLineCount = n
End Function

' ----------------------------------------------------------------------------
' IsBlankText - 空白・タブだけの行か(Trim$ はタブを落とさないため自前)
' ----------------------------------------------------------------------------
Private Function IsBlankText(ByVal t As String) As Boolean
    Dim s As String
    s = Replace(t, vbTab, " ")
    IsBlankText = (LenB(Trim$(s)) = 0)
End Function
