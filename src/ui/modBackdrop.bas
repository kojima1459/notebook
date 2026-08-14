Attribute VB_Name = "modBackdrop"
Option Explicit

' ============================================================================
' modBackdrop - 画面の「地(背景)」だけを受け持つモジュール(2026-08-14 R32波4)
' ----------------------------------------------------------------------------
' なぜ新設したのか(容量と主題の両方):
'   R32の実機第17報②で、余白まわりの受け皿になり得るモジュールが軒並み
'   上限30,000字に張り付いていた(実測: modSkin残27 / modHub残159 /
'   modUIShelf残217 / modHubStat残450 / modChrome残837 / modViewport残1,038)。
'   憲章§4-6「入らなければ実体を余裕モジュールへ置いて1行呼び出し」の
'   受け皿がもう無い状態で、地色の話だけを凝集させた新しい置き場所を作る。
'
' 持ち分は「セルの文字・Shapeではなく、その【背後】に何を敷くか」だけ:
'   ・RestoreShelfHeaderBg  … 一覧表の一律塗りで消える見出し行のグレーを戻す
'                             (W4-2。塗る順序を変えて解くための復元点)
'   ・LogStyleFailOnce      … 標準スタイル方式の失敗を1セッション1回ログへ
'                             (W4-4。無言失敗の根絶)
'   ・(W4-5で背景画像方式をここへ追加する)
'
' 【この画面の余白問題についての確定事実】(R32実機プローブ。ここに集約する)
'   1. ScrollArea はホイールを止めない。止まるのはセル選択とスクロールバー
'      だけで、"A1:L15" にしても60行付近まで転がる(R32実機実証)。
'      ホイールの停止線は S = B + k(B=使用済み末尾行 / k=1画面の行数)。
'   2. したがって「塗る深さ d を増やす」対策は原理的に効かない。塗った行は
'      使用済みになるので B が増え、S も同じだけ下がる。S − B = k は d に
'      依存しない ―― R18〜R27の7ラウンドが同じ壁に当たり続けた真因。
'   3. ThisWorkbook.Styles("Normal")/("標準") はどちらも実機で 1004。
'      無保護シートを Activate してからでも同じ。標準スタイル方式は
'      実機では使用不能(modChrome.ApplyNormalStyleBg はR28以来一度も
'      機能していなかった。R32 W4-4で失敗をログに残すよう是正)。
' ============================================================================

' ---- モジュールレベル宣言(実機VBAでは宣言部が必ず全プロシージャより先) ----
' 標準スタイル方式の失敗を1セッション1回だけログに残したか(W4-4)。
Private mStyleFailLogged As Boolean

' ----------------------------------------------------------------------------
' RestoreShelfHeaderBg - 一覧表のカード見出し行だけ、明示塗りを戻す。
' ----------------------------------------------------------------------------
'   R32 W4-2【確定バグ・実機第17報②「本棚に古い地色が残る」】:
'   マイ本棚の一覧表モードだけが地色を塗り直していなかった。
'   modUIShelf.RenderShelf には ws.Cells.Clear が無く(あるのは EnsureLayout
'   側だけ)、後始末の ApplyShelfExtent が modKnowledge.ApplyShelfBound を
'   paintBg:=False で呼んでいたため、直前にギャラリー/解決事例モードが
'   【別テーマの地色】で塗ったセルがそのまま残る ―― これがテーマを戻しても
'   一覧表だけ古い地色に見える正体。
'
'   paintBg:=False にしていた元の理由(R20-1c のコメント)は
'   「描き終えた後の一律塗りが、見出し行のグレーとカードの塗り分けを潰す」。
'   そこで塗る【順序】を変えて解く:
'     ・塗り直しはカードを描く【前】の呼び出し(ApplyShelfExtent の
'       withFont=True 側)だけで行う。カードはその後に描かれるので潰れない。
'     ・その一律塗りが消してしまう唯一の明示塗り=カード見出し行のグレーを、
'       ここで即座に戻す。
'   カード行そのものは Interior を持たない(modUIShelf 内の明示塗りは
'   見出し行の1箇所だけ。実測で確認)ので、「塗り分け」が潰れることはない。
'   描き終えた後の呼び出し(withFont=False)は従来どおり塗らない。
'
'   15921906 = RGB(242,242,242)。値の持ち主は modUIShelf.EnsureLayout で、
'   ここはその復元だけを行う(同じ値を2箇所で「決めない」)。
'   10列(A:J)は modUIShelf が一覧表で使う全列
'   (modChrome.ApplyShelfTableTextColor と同じ範囲)。
'   見出し行のグレーはテーマに依らず維持する既存の裁定(R28H F1)をそのまま
'   踏襲する ―― 文字色側(ApplyShelfTableTextColor)がこの行を除外している
'   ため、ここだけテーマ地色に変えると dark で「明るい地に明るい文字」になる。
' ----------------------------------------------------------------------------
Public Sub RestoreShelfHeaderBg(ByVal ws As Worksheet, ByVal headerRow As Long)
    If ws Is Nothing Then Exit Sub
    If headerRow < 1 Then Exit Sub
    On Error Resume Next
    ws.Cells(headerRow, 1).Resize(1, 10).Interior.Color = 15921906
    Err.Clear
    On Error GoTo 0
End Sub

' ============================================================================
' W4-4: 無言失敗の根絶
' ============================================================================

' ----------------------------------------------------------------------------
' LogStyleFailOnce - modChrome.ApplyNormalStyleBg の失敗を1行残す。
' ----------------------------------------------------------------------------
'   R32 W4-4【是正・8ラウンド気づけなかった構造的原因】:
'   modChrome.ApplyNormalStyleBg(R28波1)は
'   `ws.Parent.Styles("Normal").Interior.Color = 地色` で画面の地を一括で
'   塗るはずだったが、R32の実機プローブで
'     ・Styles("Normal") → 実行時エラー1004
'     ・Styles("標準")   → 実行時エラー1004
'     ・無保護のホームシートを Activate してから叩いても両方1004
'   が確定した。つまりこの関数はR28以降【一度も機能していない】。
'   にもかかわらず `On Error Resume Next` で握り潰され、usage_log にも
'   err_log にも1行も出ていなかったため、余白問題の調査は「地はNormal
'   スタイルで塗れている」という誤った前提の上を8ラウンド走り続けた。
'
'   よって関数自体は残す(将来のExcel/環境で使える可能性があり、削除すると
'   「試したが駄目だった」という事実まで消える)が、【黙って失敗しない】。
'   errNum<>0 のときだけ、1セッションに1回 usage_log へ残す:
'     event=normal_style_bg_failed / mode=シート名 / detail=err番号と説明
'   1回に絞る理由は、この関数が画面を描くたびに4経路から呼ばれるため
'   (毎描画で書くと usage_log が実用にならないほど膨らむ)。
'   ログ自体が失敗しても画面は落とさない(全体を On Error Resume Next 配下)。
' ----------------------------------------------------------------------------
Public Sub LogStyleFailOnce(ByVal ws As Worksheet, ByVal errNum As Long, _
                            ByVal errDesc As String)
    If errNum = 0 Then Exit Sub
    If mStyleFailLogged Then Exit Sub
    mStyleFailLogged = True          ' 先に立てる(ログ側で落ちても再入しない)
    On Error Resume Next
    Dim nm As String
    If Not ws Is Nothing Then nm = ws.Name
    modLog.LogUsage "normal_style_bg_failed", nm, _
        "Styles(Normal)への地色代入が失敗 err=" & errNum & " " & errDesc
    Err.Clear
    On Error GoTo 0
End Sub
