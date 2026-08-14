Attribute VB_Name = "modTestsPure32"
Option Explicit

' ============================================================================
' modTestsPure32 - R32波4 W4-5(背景画像方式)の純ロジック回帰テスト。
'   modTestsPure30(波1)/ modTestsPure31(波2)は直したばかりで触らない方針
'   (CLAUDE.md「禁止」)のため、既存チェーン(modTestsPure.RunAll→…→
'   modTestsPure30.RunAll30)へは繋がず、独立の新規モジュールとして新設した。
'   入口は modTestRunner.RunAllPureTests から直接呼ばれる RunAll32 の1本
'   (modTestsPure31 と同型の別枝)。
' ----------------------------------------------------------------------------
' 【本ハーネスの守備範囲】
'   Worksheet.SetBackgroundPicture そのものは LibreOffice では検証できない
'   (Excel固有API。--mode compile がコンパイル可能であることだけを確認する)。
'   したがってここで固定するのは「敷く中身=BMPのバイト列」と「描画のたびに
'   呼ばれても敷き直さないための冪等メモ」の2つ、すなわち modBackdrop の
'   純関数だけ。実機での見え方の確認はユーザーが行う。
'   BMPの構造(58バイトの内訳とリトルエンディアンの並び)は
'   modBackdrop.bas 冒頭「★ BMPバイト列の検算」に全バイト分の根拠がある。
' ============================================================================

' ----------------------------------------------------------------------------
' R32 W4-5(A): 1×1・24bit BMP の全58バイト。
'   ゴールデンは modBackdrop.bas 冒頭の検算表と1バイトずつ対応している:
'     424D                             'B''M'
'     3A000000                         bfSize=58
'     00000000                         bfReserved1/2=0
'     36000000                         bfOffBits=54
'     28000000                         biSize=40
'     01000000 01000000                biWidth=1 / biHeight=1
'     0100 1800                        biPlanes=1 / biBitCount=24
'     00000000                         biCompression=BI_RGB(0)
'     04000000                         biSizeImage=4(=((1*24+31)\32)*4)
'     130B0000 130B0000                biX/YPelsPerMeter=2835(72dpi)
'     00000000 00000000                biClrUsed / biClrImportant=0
'     <BB><GG><RR>00                   画素(BGR順)+4バイト境界パディング
'   darkテーマの地 RGB(15,23,42) は VBA の Long で 15 + 23*256 + 42*65536
'   = 2758415。画素は B=&H2A / G=&H17 / R=&H0F の順で 2A170F、末尾に 00。
' ----------------------------------------------------------------------------
Private Sub TestBmpHexDark32()
    Const HEAD As String = "424D3A0000000000000036000000280000000100000001000000010018000000000004000000130B0000130B00000000000000000000"
    modTestRunner.Check "R32-W4-5_BMP58バイト(dark RGB(15,23,42))が検算どおり", _
        (modBackdrop.BmpHex(2758415) = HEAD & "2A170F00"), _
        "実際=" & modBackdrop.BmpHex(2758415)

    ' 長さは常に58バイト=116字(色に依らない)。
    modTestRunner.Check "R32-W4-5_BMPの長さは常に116字(58バイト)", _
        (Len(modBackdrop.BmpHex(2758415)) = 116), _
        "実際=" & Len(modBackdrop.BmpHex(2758415))

    ' 白 RGB(255,255,255)=16777215。画素は FF FF FF + パディング00。
    modTestRunner.Check "R32-W4-5_白(16777215)の画素はFFFFFF00", _
        (modBackdrop.BmpHex(16777215) = HEAD & "FFFFFF00"), _
        "実際=" & modBackdrop.BmpHex(16777215)

    ' 黒(0)。Applyは0を異常値として弾くが、バイト列生成そのものは成立する。
    modTestRunner.Check "R32-W4-5_黒(0)の画素は00000000", _
        (modBackdrop.BmpHex(0) = HEAD & "00000000"), _
        "実際=" & modBackdrop.BmpHex(0)

    ' 【BGR順であること】の単独固定。RGB順に書いてしまう取り違えは、実機では
    ' 「青系テーマなのに赤い背景」という形でしか現れず、コンパイルもLOも
    ' 素通りする。R=1・G=2・B=3 の非対称な色で順序だけを撃つ。
    ' RGB(1,2,3) = 1 + 2*256 + 3*65536 = 197121 → 画素は 03 02 01 00。
    modTestRunner.Check "R32-W4-5_画素はBGR順(RGB(1,2,3)→030201)", _
        (modBackdrop.BmpHex(197121) = HEAD & "03020100"), _
        "実際=" & modBackdrop.BmpHex(197121)
End Sub

' ----------------------------------------------------------------------------
' R32 W4-5(B): 冪等メモ。Apply は画面描画のたびに呼ばれる(Setup*Columns 経由)
'   ので、「同じシートへ同じ色」なら一時ファイル生成も SetBackgroundPicture も
'   走らせない。その判定に使う純関数を固定する。
'   ・MemoKey は照合用の "|シート名=色|"
'   ・MemoPut は同じシートの古い1件を必ず落としてから足す(重複させない)
' ----------------------------------------------------------------------------
Private Sub TestBackdropMemo32()
    modTestRunner.Check "R32-W4-5_MemoKeyは|シート名=色|の形", _
        (modBackdrop.MemoKey("ホーム", 123) = "|ホーム=123|"), _
        "実際=" & modBackdrop.MemoKey("ホーム", 123)

    Dim m As String
    m = modBackdrop.MemoPut("", "ホーム", 123)
    modTestRunner.Check "R32-W4-5_空メモへの1件目", (m = "|ホーム=123|"), "実際=" & m

    m = modBackdrop.MemoPut(m, "Dashboard", 456)
    modTestRunner.Check "R32-W4-5_別シートは追加される", _
        (m = "|ホーム=123|Dashboard=456|"), "実際=" & m

    ' 同じシートの色替え: 古い1件が消え、新しい1件が末尾へ付く(重複しない)。
    m = modBackdrop.MemoPut(m, "ホーム", 789)
    modTestRunner.Check "R32-W4-5_同一シートの再登録で古い1件が落ちる", _
        (m = "|Dashboard=456|ホーム=789|"), "実際=" & m

    ' 冪等判定: 敷いた色は当たり、敷いていない色は当たらない。
    modTestRunner.Check "R32-W4-5_敷いた色はメモに当たる", _
        (InStr(1, m, modBackdrop.MemoKey("ホーム", 789), vbBinaryCompare) > 0), _
        "メモ=" & m
    modTestRunner.Check "R32-W4-5_別の色はメモに当たらない(敷き直しが走る)", _
        (InStr(1, m, modBackdrop.MemoKey("ホーム", 123), vbBinaryCompare) = 0), _
        "メモ=" & m
    ' 前方一致の取り違え防止: "ホーム=78" は "ホーム=789" に当たってはいけない
    ' (キーの両端を "|" で挟んでいることの確認)。
    modTestRunner.Check "R32-W4-5_色の前方一致では誤ヒットしない", _
        (InStr(1, m, modBackdrop.MemoKey("ホーム", 78), vbBinaryCompare) = 0), _
        "メモ=" & m
End Sub

' ============================================================================
Public Sub RunAll32()
    On Error GoTo Fail32
    TestBmpHexDark32
    TestBackdropMemo32
NextDone32:
    On Error GoTo 0
    Exit Sub

Fail32:
    modTestRunner.Check "TestBmpHexDark32/TestBackdropMemo32(グループ全体)", False, _
        "実行時エラー: " & Err.Description & " (Err=" & Err.Number & ")"
    Resume NextDone32
End Sub
