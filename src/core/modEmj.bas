Attribute VB_Name = "modEmj"
Option Explicit

' ============================================================================
' modEmj - 絵文字の単一情報源(R47)
' ----------------------------------------------------------------------------
' なぜ要るのか
' ----------------------------------------------------------------------------
' Unicode の絵文字には「既定がカラー表示」のものと「既定がテキスト字形」の
' ものがある(Emoji_Presentation プロパティ)。後者は異体字セレクタ U+FE0F を
' 付けないと、レンダラが白黒の記号字形を探しにいく。Windows はその字形を持つ
' フォントを削っているため、【何も描かれない】。
'
' R46 の実機報告「削除ボタンの絵文字が出ない」の原因がこれだった。そして
' R46 の修正は 🗑 を2箇所直しただけで、同じ文字の残り6箇所と、⚠(20箇所中
' 15箇所)・🖼・⚙・☀・◀▶ など他の EP=No 文字を取りこぼしていた。
' R46 のコミットには「⚠ には全箇所で FE0F を付けている」と書いたが、
' それは事実ではなかった(実測 20箇所中5箇所)。**誤った記述だったので訂正する。**
'
' 直書きを禁止し、ここを通す
' ----------------------------------------------------------------------------
' 呼び出し側で ChrW を並べる限り、新しい絵文字を足すたびに同じ取りこぼしが
' 起きる。文字ごとに名前の付いた関数をここに置き、呼び出しは1つの式にする。
' 効果は2つ:
'   (1) FE0F の有無をこの1ファイルで保証できる(付け忘れようがない)
'   (2) 呼び出し側の【字数が減る】。ChrW(&HD83D) & ChrW(&HDDD1) & ChrW(&HFE0F)
'       は40字だが modEmj.Trash() は15字。残り10字しか無い modUIShelf のような
'       モジュールへ、容量を増やさずに FE0F を入れられる唯一の道。
'
' 機械で守る
' ----------------------------------------------------------------------------
' tools/vba_lint.py が「EP=No の絵文字を ChrW で直書きしていて FE0F が
' 続いていない」箇所を ERROR にする。このモジュール自身だけが除外される。
'
' R4準拠(純ロジック): Excel オブジェクトに触れない。CP932 準拠: 本文は ASCII と
' CP932 内の文字だけで書き、絵文字そのものはソースに直書きしない(ChrW で組む)。
' ============================================================================

' 異体字セレクタ。Const に ChrW は書けないので関数で持つ。
Private Function VS16() As String
    VS16 = ChrW(&HFE0F)
End Function

' ----------------------------------------------------------------------------
' 非BMP(サロゲートペアで組む文字)
' ----------------------------------------------------------------------------

' 🗑 U+1F5D1 WASTEBASKET(削除)
Public Function Trash() As String
    Trash = ChrW(&HD83D) & ChrW(&HDDD1) & VS16()
End Function

' 🖼 U+1F5BC FRAME WITH PICTURE(画像)
Public Function Picture() As String
    Picture = ChrW(&HD83D) & ChrW(&HDDBC) & VS16()
End Function

' 🏷 U+1F3F7 LABEL(ラベル・富化)
Public Function Label() As String
    Label = ChrW(&HD83C) & ChrW(&HDFF7) & VS16()
End Function

' 🗔 U+1F5D4 DESKTOP WINDOW(作業用Excel)
Public Function WindowIcon() As String
    WindowIcon = ChrW(&HD83D) & ChrW(&HDDD4) & VS16()
End Function

' ----------------------------------------------------------------------------
' BMP(1文字で組む文字)
' ----------------------------------------------------------------------------

' ⚠ U+26A0 WARNING SIGN(警告)
Public Function Warn() As String
    Warn = ChrW(&H26A0) & VS16()
End Function

' ⚙ U+2699 GEAR(設定)
Public Function Gear() As String
    Gear = ChrW(&H2699) & VS16()
End Function

' ✍ U+270D WRITING HAND(執筆中)
Public Function Writing() As String
    Writing = ChrW(&H270D) & VS16()
End Function

' ☀ U+2600 BLACK SUN WITH RAYS(ライトテーマ)
Public Function Sun() As String
    Sun = ChrW(&H2600) & VS16()
End Function

' ↩ U+21A9 LEFTWARDS ARROW WITH HOOK(復帰)
Public Function Undo() As String
    Undo = ChrW(&H21A9) & VS16()
End Function

' ⌨ U+2328 KEYBOARD(ショートカット)
Public Function Keyboard() As String
    Keyboard = ChrW(&H2328) & VS16()
End Function

' ⏱ U+23F1 STOPWATCH(時間)
Public Function Stopwatch() As String
    Stopwatch = ChrW(&H23F1) & VS16()
End Function

' ✉ U+2709 ENVELOPE(メール)
Public Function Mail() As String
    Mail = ChrW(&H2709) & VS16()
End Function

' ◀ U+25C0 BLACK LEFT-POINTING TRIANGLE(前へ)
Public Function ArrowLeft() As String
    ArrowLeft = ChrW(&H25C0) & VS16()
End Function

' ▶ U+25B6 BLACK RIGHT-POINTING TRIANGLE(次へ)
Public Function ArrowRight() As String
    ArrowRight = ChrW(&H25B6) & VS16()
End Function

' ⏹ U+23F9 BLACK SQUARE FOR STOP(中断)
Public Function StopMark() As String
    StopMark = ChrW(&H23F9) & VS16()
End Function
