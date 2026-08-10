Attribute VB_Name = "modInstallCheck"
Option Explicit

' ============================================================================
' modInstallCheck - 自己インストーラの注入結果を「行数」で検算する
'                   (2026-08-10 R23b・実機第9報①の再発対策 MA-3)
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   この配布物は起動のたびに vba_src シートの本文を VBE へ注入して自分自身を
'   組み立てる(自己インストーラ型)。R23 で「名前だけ付いた空モジュール」
'   (CountOfLines < 1)は検出できるようにしたが、それでは
'   【部分注入】= AddFromString が途中まで書いて終わった状態を1件も拾えない。
'   1行でも入っていれば CountOfLines >= 1 なので素通りし、そのまま Save まで
'   到達して壊れたブックがファイルへ焼き付く。実機で繰り返し出ている
'   「modViewport2.BadgeRowsFor が見つかりません」は、まさに
'   「そのモジュールは在るが中身が足りない」形のコンパイルエラーである。
'
'   そこで注入の直後に、全モジュールについて
'       vba_src の本文の行数(期待値)  ==  VBE 上の CodeModule の行数(実測値)
'   を突き合わせる。1件でも食い違えば「セットアップ未完了」と判定し、
'   インストーラ側の既存ガード(f > 0 なら Save しない)へ合流させる。
'   ペイロード(vba_src)はファイル上で無傷なので、保存せずに閉じて開き直せば
'   全量が再試行される。利用者にはその手順と壊れたモジュール名を見せる。
'
' 置き場が基盤層(src/core)である理由:
'   起動の最初期(modBoot すら注入直後)に呼ばれるため、他モジュールへ一切
'   依存してはいけない。実際このモジュールは自分自身の Private 関数しか
'   呼ばない(R1 の最も内側)。modIntegrity/modDiag が同じ理由で同じ層に
'   置かれているのと同型。
'
' 呼ばれ方:
'   自己インストーラ(ThisWorkbook ストリーム)の注入ループ完了後から
'       f = f + Application.Run("modInstallCheck.VI")
'   の1行だけで呼ばれる。Application.Run 自体が失敗する場合
'   (=このモジュールの注入に失敗した/コンパイルできない)も、
'   インストーラ側が Err.Number を見て f を加算する。つまり
'   「検証が動かなかった」も不合格側に倒れる。
'
' 行数比較の仕様(なぜ「行数」で、なぜ「末尾空行を落とす」のか):
'   ・本文そのものの完全一致ではなく行数にしたのは、CodeModule.Lines で
'     全文を取り出して比較すると 135 本ぶんの巨大な文字列結合が起動ごとに
'     走るうえ、VBE 側が行末の空白を落とす等の無害な差でも不一致になり、
'     偽陽性で「絶対に起動できないブック」を作りかねないため。
'     部分注入は必ず行数が減る(途中で切れる)ので、行数だけで確実に捕まる。
'   ・AddFromString は末尾の空行の扱いが環境で1本ぶれることが知られている。
'     そこで【両側とも末尾の空行を落としてから】厳密一致で比較する。
'     ±1 の黙認はしない(1行足りない部分注入を見逃す穴になるため)。
'
' 期待行数の出所(2026-08-10 R23bH-F1): 以前の VI は起動のたびに vba_src の
' C列(全136本・本文合計約260万字)を読み直し、Replace/Split で期待行数を
' その場で再計算していた。これは注入直後で32bit Excelのメモリが最も逼迫
' しているタイミングに、検証器自身が新たな失敗点(OOM/低速化)になり得る。
' そこで期待行数は【ビルド時】に build/build_mybookshelf.py の
' _make_vba_src が _expected_line_count(=このモジュールの ExpectedLineCount
' と完全同一規則)で計算し、vba_src!D列(row2以降)へ Long であらかじめ
' 焼き込む。verify_build 側も同じ規則で全数突合するので、D列の値そのものが
' 出荷前に検算済みである。VI はもう C列本文を読まず、D列の数値を読んで
' 実測行数(ModuleLineCount)と比較するだけになった。
' ExpectedLineCount/LineCountMismatch は削除していない: ビルド時の
' _expected_line_count と実装が乖離しないことを固定する純関数テスト資産
' (src/test/modTestsPure23.bas)としてそのまま残す。
' ============================================================================

' vba_src シートの名前と読む列。インストーラ本体(ThisWorkbookストリーム)と
' 同じ約束。A列=モジュール名(インストーラも読む)/ D列=期待行数(Long。
' build/build_mybookshelf.py の _make_vba_src がビルド時に焼き込む。
' row1のD1は書かない=E1のOnTime予約時刻と同じ理由でrow2以降のみを使う)。
' C列(ソース本文)はインストーラの注入に使うが、VI はもう読まない(F1)。
Private Const SHEET_VBA_SRC As String = "vba_src"
Private Const COL_NAME As Long = 1
Private Const COL_EXPECTED As Long = 4
Private Const FIRST_DATA_ROW As Long = 2

' xlUp の値。定数名で書かずに数値なのは、この関数が「まだ何も注入できて
' いないかもしれない」最初期に走るためで、インストーラ本体の書き方に揃える。
Private Const XL_UP As Long = -4162

' MsgBox に名前を並べる上限(全滅時に135行のダイアログを出さないため)。
Private Const MAX_LISTED As Long = 5

' ----------------------------------------------------------------------------
' LineCountMismatch - 期待ソースと実測行数が食い違うか(純関数)
' ----------------------------------------------------------------------------
' expectedSrc は vba_src セルの本文(改行は vbLf。念のため vbCrLf/vbCr も
' 正規化する)。actualLineCount は呼び出し側が【末尾の空行を落としたあと】の
' CodeModule 行数を渡す。expectedSrc 側の末尾空行はこの関数が落とす。
' True = 不一致(=不完全な注入の疑い)。
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

' ----------------------------------------------------------------------------
' VI - 注入結果の検証本体。不一致モジュール数を返す(0 = 合格)
' ----------------------------------------------------------------------------
' 名前が2文字なのは、呼び出し側が自己インストーラ(圧縮後1,148バイト上限の
' ThisWorkbook ストリーム)で、1バイトでも短い名前が要るため
' (Verify Injection の略)。
'
' 異常時(vba_src が無い・VBProject を読めない等)は 1 を返す。
' 「検証できなかった」を合格側へ倒すと、この仕組みそのものが無意味になる。
'
' 期待行数は D列の値をそのまま使う(2026-08-10 R23bH-F1。C列本文の
' 実行時再読込・Replace/Splitはもうしない)。D列が空/非数値のときは
' 「期待行数が読めない」こと自体を不一致として不合格側(bad+1)へ倒す。
' IsNumeric で先に型を確かめてから CLng するのは、VBA の And が短絡評価
' しないため「IsNumeric(v) And CLng(v)=actual」のような1行に書くと
' 非数値のときに CLng 側で型不一致の実行時エラーになりかねないからで、
' 素直に If を分けて保守的に倒す。
' ----------------------------------------------------------------------------
Public Function VI() As Long
    Dim vbp As Object
    Dim ws As Object
    Dim lastRow As Long
    Dim r As Long
    Dim nm As String
    Dim dVal As Variant
    Dim expected As Long
    Dim actual As Long
    Dim bad As Long
    Dim listed As String

    On Error GoTo Fail

    Set vbp = ThisWorkbook.VBProject
    Set ws = ThisWorkbook.Worksheets(SHEET_VBA_SRC)
    lastRow = ws.Cells(ws.Rows.Count, COL_NAME).End(XL_UP).Row

    For r = FIRST_DATA_ROW To lastRow
        nm = CStr(ws.Cells(r, COL_NAME).Value)
        If LenB(nm) > 0 Then
            ' 見つからない/読めないモジュールは -1(必ず不一致になる)。
            actual = ModuleLineCount(vbp, nm)
            dVal = ws.Cells(r, COL_EXPECTED).Value
            If IsNumeric(dVal) Then
                expected = CLng(dVal)
                If expected <> actual Then
                    bad = bad + 1
                    If bad <= MAX_LISTED Then
                        If LenB(listed) > 0 Then listed = listed & vbLf
                        listed = listed & "  " & nm & " (" & actual & " / " & expected & " 行)"
                    End If
                End If
            Else
                ' D列が空/非数値: 期待行数そのものが読めない=保守的に不合格へ倒す。
                bad = bad + 1
                If bad <= MAX_LISTED Then
                    If LenB(listed) > 0 Then listed = listed & vbLf
                    listed = listed & "  " & nm & " (D列不正: 実測" & actual & " 行)"
                End If
            End If
        End If
    Next r

    If bad > 0 Then WarnBroken bad, listed
    VI = bad
    Exit Function

Fail:
    ' 検証そのものが実行できなかった。合格にはしない。
    VI = 1
End Function

' ----------------------------------------------------------------------------
' ModuleLineCount - VBE 上のモジュールの「末尾空行を落とした行数」
' ----------------------------------------------------------------------------
' 見つからない・読めない場合は -1(期待値は必ず 0 以上なので確実に不一致)。
' 別 Function に切り出しているのは、VBComponents(nm) が存在しないときの
' 実行時エラーを【呼び出し元のエラーハンドラを稼働させずに】受けるため
' (稼働中のハンドラの内側では On Error が効かない: MASTER_SPEC の
'  modExtractorWord 462 の教訓)。
' ----------------------------------------------------------------------------
Private Function ModuleLineCount(ByVal vbp As Object, ByVal nm As String) As Long
    Dim cm As Object
    Dim n As Long

    On Error GoTo NoMod
    Set cm = vbp.VBComponents(nm).CodeModule
    n = cm.CountOfLines
    Do While n > 0
        If Not IsBlankText(cm.Lines(n, 1)) Then Exit Do
        n = n - 1
    Loop
    ModuleLineCount = n
    Exit Function

NoMod:
    ModuleLineCount = -1
End Function

' ----------------------------------------------------------------------------
' WarnBroken - 壊れたモジュール名を見せて、やるべき操作だけを伝える
' ----------------------------------------------------------------------------
' インストーラ側の MsgBox は英字1行しか置けない(サイズ上限)ので、
' 日本語の詳細説明はこちらが担う。ここで出してから件数を返す。
' ----------------------------------------------------------------------------
Private Sub WarnBroken(ByVal bad As Long, ByVal listed As String)
    Dim msg As String
    msg = "セットアップ検証NG: 以下のモジュールが不完全です。" & vbLf
    msg = msg & "保存せずに閉じて、開き直してください(自動でやり直します)。" & vbLf & vbLf
    msg = msg & listed & vbLf
    If bad > MAX_LISTED Then msg = msg & "  ...ほか" & (bad - MAX_LISTED) & "件" & vbLf
    msg = msg & vbLf & "不一致 " & bad & " 件 / 表示は実測行数 / 期待行数。"
    On Error Resume Next
    MsgBox msg, vbCritical
    Err.Clear
    On Error GoTo 0
End Sub
