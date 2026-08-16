Attribute VB_Name = "modVault"
Option Explicit


' ============================================================================
' modVault - ナレッジ登録フォーム(DOCS_NEXUS_SPEC Phase 2・裁定①)。
'   UserForm代替(専用シートをSPA遷移でアクティブ化。フォーカス制御が
'   不安定なためShape入力は使わず、入力欄はセルで作る)。登録は既存の
'   取込パイプライン(modShelf.IngestFile)を再利用する。
' ============================================================================

Private Const VAULT_SHEET As String = "VaultInput"
' R18-3a/3b: この画面が実際に使うセル範囲(列A:I・内容はC2:H18で完結)。
' 書式の適用範囲とScrollAreaの唯一の情報源。
Private Const VAULT_BOUND As String = "A1:I30"
' R20-1g: 帯(列A〜吸収列I)の1行ぶん。列幅はここを可視幅へ合わせて決める。
Private Const VAULT_BAND As String = "A1:I1"
Private Const CELL_TITLE As String = "C6"
Private Const CELL_BODY As String = "C8"
Private Const CELL_TAGS As String = "C18"


' ----------------------------------------------------------------------------
' ShowVaultInput - 登録フォームを描画して表示(SPA遷移)
' ----------------------------------------------------------------------------
Public Sub ShowVaultInput()
    Dim ws As Worksheet
    Set ws = GetOrCreateVaultSheet()
    If ws Is Nothing Then Exit Sub

    On Error GoTo Finish
    Application.ScreenUpdating = False
    modUiLock.AlertsOff   ' R25-1a-1: 描画中のMerge警告を出さない(対はFinishCleanup0)

    ' 冪等再構築
    RemoveVaultShapes ws
    ws.Cells.Clear
    ' R18-3a: 全域(ws.Cells)への書式はUsedRangeをシート最大へ膨らませる
    ' (無限スクロールの主因・調査agent2 §1.3)。実使用範囲だけに当てる。
    ws.Range(VAULT_BOUND).Font.Name = "Yu Gothic UI"
    ws.Range(VAULT_BOUND).Interior.Color = RGB(243, 244, 246)

    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 3
    ws.Columns("C:G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 1
    ' R20-1g(実機第7報⑦の層1): 列幅の合計527ptが固定で、窓を広げるほど
    ' その右が丸ごと灰色の余白になっていた(常用画面なので同梱)。余りは
    ' 本文の最終列Hに吸わせ、吸収列Iは最小のまま置く ―― こうすると
    ' 白いカード(C2:H18)の右端そのものが窓幅へ追随する(帯だけ伸ばしても
    ' カードは527ptで止まったままで、症状は消えない)。
    modViewport.FitBandToViewport ws, VAULT_BAND, "H"
    ' R18-3b: 行ける範囲を宣言(この小画面は C2:H18 で完結する)。
    modViewport.ApplyScrollBound ws, VAULT_BOUND

    ' カード風の背景(白)。Shapeはセルより必ず手前に描画される(ZOrderは
    ' Shape同士の前後関係にしか効かない)ため、Shapeで背景を作るとラベル等の
    ' セル文字が完全に隠れてしまう(実機報告のバグ)。セルの塗りつぶしで代替する。
    ws.Range("C2:H18").Interior.Color = RGB(255, 255, 255)

    ' タイトル
    With ws.Range("C2:H2")
        .Merge
        .Value = ChrW(&H2795) & " 新規ナレッジの登録"
        .Font.Size = 14
        .Font.Bold = True
    End With
    With ws.Range("C3:H3")
        .Merge
        .Value = "登録した内容はベクトル化され、社内ナレッジ検索の回答に使われます。"
        .Font.Size = 9
        .Font.Color = RGB(107, 114, 128)
    End With

    ' タイトル入力
    With ws.Range("C5:H5")
        .Merge
        .Value = "タイトル(例: 漁船保険の特例について)"
        .Font.Size = 9.5
        .Font.Bold = True
    End With
    With ws.Range("C6:H6")
        .Merge
        .Interior.Color = RGB(249, 250, 251)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(229, 231, 235)
    End With
    ws.Rows(6).RowHeight = 24

    ' 本文入力
    With ws.Range("C7:H7")
        .Merge
        .Value = "ナレッジ本文(またはAIへの修正指示)"
        .Font.Size = 9.5
        .Font.Bold = True
    End With
    With ws.Range("C8:H16")
        .Merge
        .WrapText = True
        .VerticalAlignment = -4160   ' xlTop
        .Interior.Color = RGB(249, 250, 251)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(229, 231, 235)
    End With

    ' タグ入力
    With ws.Range("C17:H17")
        .Merge
        .Value = "タグ(カンマ区切り。例: 約款解釈,特約)"
        .Font.Size = 9.5
        .Font.Bold = True
    End With
    With ws.Range("C18:H18")
        .Merge
        .Interior.Color = RGB(249, 250, 251)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(229, 231, 235)
    End With
    ws.Rows(18).RowHeight = 24

    ' ボタン(登録=青 / キャンセル=白)
    Dim submitBtn As Shape
    Set submitBtn = ws.Shapes.AddShape(5, 420, 388, 140, 32)
    submitBtn.Name = "nxv_submit"
    submitBtn.Line.Visible = 0
    submitBtn.Fill.ForeColor.RGB = RGB(37, 99, 235)
    With submitBtn.TextFrame2
        .TextRange.Text = "ベクトル化して登録"
        .TextRange.Font.Size = 10.5
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    submitBtn.OnAction = "modVault.OnVaultSubmit"

    Dim cancelBtn As Shape
    Set cancelBtn = ws.Shapes.AddShape(5, 320, 388, 90, 32)
    cancelBtn.Name = "nxv_cancel"
    cancelBtn.Fill.ForeColor.RGB = RGB(255, 255, 255)
    cancelBtn.Line.ForeColor.RGB = RGB(229, 231, 235)
    With cancelBtn.TextFrame2
        .TextRange.Text = "キャンセル"
        .TextRange.Font.Size = 10.5
        .TextRange.Font.Fill.ForeColor.RGB = RGB(17, 24, 39)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    cancelBtn.OnAction = "modVault.OnVaultCancel"

    modUI.FreezeShapePlacement ws   ' 全Shapeを絶対配置に固定(ズレ防止)

    ' SPA遷移(表示してアクティブ化・枠線等は非表示)
    ws.Visible = -1   ' xlSheetVisible
    If modUI.ActivateSheetRobust(ws, "modVault.ShowVaultInput") Then
        On Error Resume Next
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ws.Range(CELL_TITLE).Select
        On Error GoTo 0
    Else
        modUI.RestoreExcelUI
    End If

    ' 正常系はハンドラ本体(Resume)を跨いで後始末へ入る
    ' (Resume はエラーが起きていないと実行時エラー20になる)。
    GoTo FinishCleanup0
Finish:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FinishCleanup0
FinishCleanup0:
    On Error Resume Next
    modUiLock.AlertsOn
    Application.ScreenUpdating = True   ' 例外時も必ず画面更新を戻す(暗転固定を防ぐ)
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OnVaultSubmit - 入力内容を既存取込パイプラインへ流して登録
' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------
' OnVaultSubmit - 「登録」ボタン(2026-07-31 R11-F2で進捗表示と二重送信ガードを
'   追加。R11-Eで容量不足のため保留になっていた項目)。
'   登録は本文のチャンク分割+埋め込みAPI呼び出しを伴い、本文が長いと十数秒
'   かかる。従来はその間まったくの無反応で、利用者が「効いていない」と思って
'   もう一度押すと同じナレッジが2件登録されていた(憲章§3-1/§3-2)。
' ----------------------------------------------------------------------------
Public Sub OnVaultSubmit()
    If modUiLock.BlockIfIngesting() Then Exit Sub
    ' 二重送信ガード: 1回目の登録が終わるまで2回目を受け付けない
    ' (modUiLock.Enter が False を返す=すでに何かが走っている)。
    If Not modUiLock.Enter() Then Exit Sub

    Dim ws As Worksheet
    Dim ok As Boolean
    Dim titleText As String, bodyText As String, tagsText As String

    Set ws = GetVaultSheet()
    If ws Is Nothing Then GoTo SubmitDone

    titleText = Trim$(CStr(ws.Range(CELL_TITLE).Value))
    bodyText = Trim$(CStr(ws.Range(CELL_BODY).Value))
    tagsText = Trim$(CStr(ws.Range(CELL_TAGS).Value))

    If LenB(titleText) = 0 Or LenB(bodyText) = 0 Then
        MsgBox "タイトルと本文を入力してください。", vbExclamation, modAppDef.APP_NAME
        GoTo SubmitDone
    End If

    On Error GoTo SubmitFail
    modUIMain.ShowProgress "ナレッジを登録しています…(長い本文は少し時間がかかります)"
    ok = RegisterKnowledgeText(titleText, TagsHeader(tagsText) & bodyText, tagsText)
    modUIMain.HideProgress

    If ok Then
        ' 2026-08-06 R20H FA-16: 「ベクトル化されました」は生ジャーゴン。
        ' タイトルも他の全ハンドラと同じmodAppDef.APP_NAME参照へ揃える
        ' (このMsgBoxだけ"Nexus Agent"のハードコードだった)。
        MsgBox "ナレッジデータベースに追加され、AIが読み込みました。", vbInformation, modAppDef.APP_NAME
        ClearInputs ws
        CloseVault ws
    Else
        MsgBox "登録に失敗しました。マイ本棚の一覧で状態をご確認ください。", vbExclamation, modAppDef.APP_NAME
    End If
    GoTo SubmitDone

SubmitFail:
    Dim subDesc As String: subDesc = Err.Description
    Dim subNum As Long: subNum = Err.Number
    ' ハンドラ稼働中は On Error Resume Next が効かない。後始末の前に Resume で抜ける。
    Resume SubmitCleanup
SubmitCleanup:
    On Error Resume Next
    modUIMain.HideProgress
    modLog.LogError "E0801", "modVault.OnVaultSubmit", subDesc, subNum
    MsgBox "登録に失敗しました。もう一度お試しください。" & vbLf & _
           "(繰り返す場合は、本文を短くして分けて登録してみてください)", _
           vbExclamation, modAppDef.APP_NAME
    On Error GoTo 0
SubmitDone:
    modUiLock.Leave
End Sub

Public Sub OnVaultCancel()
    Dim ws As Worksheet
    Set ws = GetVaultSheet()
    If ws Is Nothing Then Exit Sub
    ClearInputs ws
    CloseVault ws
End Sub

' ----------------------------------------------------------------------------
' RegisterKnowledgeText - テキストナレッジを既存パイプラインで登録する共通口
'   (modAppの👎自己学習からも使う)。成功=True(done/partial)。
'
' 2026-08-11(R26-3): 第4引数 docBase を追加(省略時は従来と1文字も変わらない)。
'   本棚に並ぶ「資料名」は取り込んだファイル名そのものなので、資料名を呼び出し
'   側が決めたいとき(考察メモの `考察メモ_題名` 固定)は一時ファイルの名前を
'   決められる必要がある。従来の "ナレッジ_題名_yyyymmddhhnnss.txt" は
'   毎回別名になり資料名が一意にならないため、既存呼び出しの挙動は据え置いて
'   別名を渡せる口だけを開ける。docBase も SanitizeName(60字上限・
'   ファイル名禁止文字を"_"へ)を必ず通す。
' ----------------------------------------------------------------------------
' 2026-08-16(R33 W5-19): 第5引数 silent を追加(省略時は従来どおり False で
'   1文字も挙動が変わらない)。一括取込の呼び出し元が「1件ごとの再描画」を
'   止められるようにするためのもの。modShelf.IngestFile は silent=False だと
'   1件ごとに modUIShelf.RenderShelf を呼ぶので、共有モードで回すと
'   Shape全削除→Cells.Clear→クロム再生成→ScreenUpdating復帰までが件数ぶん
'   繰り返され、画面が激しく明滅する。同型の一括経路(modShelfSync/
'   modShelfBatch)は既に silent:=True + 末尾で1回描く形へ手当済みだった。
Public Function RegisterKnowledgeText(ByVal titleText As String, ByVal bodyText As String, _
                                      ByVal tagsText As String, _
                                      Optional ByVal docBase As String = "", _
                                      Optional ByVal silent As Boolean = False) As Boolean
    On Error GoTo Fail

    Dim tempDir As String: tempDir = Environ$("TEMP")
    If LenB(tempDir) = 0 Then tempDir = Environ$("TMP")
    If LenB(tempDir) = 0 Then Exit Function
    If Right$(tempDir, 1) <> "\" Then tempDir = tempDir & "\"

    Dim filePath As String
    If LenB(docBase) > 0 Then
        filePath = tempDir & SanitizeName(docBase) & ".txt"
    Else
        filePath = tempDir & "ナレッジ_" & SanitizeName(titleText) & "_" & _
                   Format$(Now, "yyyymmddhhnnss") & ".txt"
    End If

    Dim content As String
    content = "【" & titleText & "】" & vbLf & bodyText

    ' UTF-8で保存(modExtractorのtxt読取りがUTF-8のため整合)。実体は
    ' modUtilText.WriteTextFileUtf8(2026-07-31 R11-F2で9箇所の同型実装を
    ' 1本化)。書けなかったときは従来と同じく E0801 を1行残して False を返す
    ' (無言で「登録できませんでした」だけを出さない。憲章§4-1)。
    Dim wErrNum As Long, wErrDesc As String
    If Not modUtilText.WriteTextFileUtf8(filePath, content, wErrNum, wErrDesc) Then
        modLog.LogError "E0801", "modVault.RegisterKnowledgeText", _
            "一時ファイルの書き出しに失敗: " & wErrDesc, wErrNum
        Exit Function
    End If

    Dim resultStatus As String
    resultStatus = modShelf.IngestFile(filePath, "self", silent)

    ' 一時ファイルは掃除(取込済みなので不要。失敗しても無視)
    ' R26H F6(m-1): Kill はパスをANSI(CP932)へ落として渡すため、変換できない
    ' 非BMP文字が "?" に化ける。考察メモの資料名は 💭 で始まる(modInsightCard.
    ' DocNameFor)ので TEMP\?考察メモ_… となり、Kill はその "?" を1文字
    ' ワイルドカードとして解釈する=TEMP直下の無関係な同型名まで巻き込みうる。
    ' FileSystemObject はパスをUnicodeのまま受け取るのでこの化けが起きない。
    ' SanitizeName が "*" "?" を "_" へ潰しているので、こちら側でワイルドカードが
    ' 残ることも無い。On Error の作法(失敗しても無視)は従来のまま。
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFile filePath, True
    On Error GoTo 0

    RegisterKnowledgeText = (resultStatus = "done" Or resultStatus = "partial")
    Exit Function

Fail:
    ' Err の内容は Resume でクリアされるので先に控える。
    Dim failDesc As String: failDesc = Err.Description
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup3
FailCleanup3:
    On Error Resume Next
    modLog.LogError "E0801", "modVault.RegisterKnowledgeText", failDesc
    On Error GoTo 0
    RegisterKnowledgeText = False
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー
' ----------------------------------------------------------------------------

Private Function TagsHeader(ByVal tagsText As String) As String
    If LenB(tagsText) = 0 Then Exit Function
    TagsHeader = "タグ: " & tagsText & vbLf & vbLf
End Function

Private Sub ClearInputs(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range(CELL_TITLE).Value = ""
    ws.Range(CELL_BODY).Value = ""
    ws.Range(CELL_TAGS).Value = ""
    On Error GoTo 0
End Sub

' レビュー1-C: 順序を逆にした。アクティブなシートは隠せず1004になるため、
' 旧実装は「隠そうとして失敗→隠れないままギャラリーを描く」状態だった。
' 先に行き先(登録前に見ていたモード)を描いて前面へ出し、そのあとで隠す。
' ギャラリー固定をやめたのは、一覧表から登録した人が毎回ギャラリーへ
' 飛ばされていたため。
Private Sub CloseVault(ByVal ws As Worksheet)
    On Error Resume Next
    modKnowledge.RefreshCurrent
    ' RenderShelf(一覧表)はシートをアクティブにしないので必ず前面へ出す。
    modUI.GoToNativeSheet modAppDef.SH_SHELF, "modVault(CloseVault)"
    ws.Visible = 2   ' xlSheetVeryHidden
    On Error GoTo 0
End Sub

Private Function SanitizeName(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    Dim t As String: t = s
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        t = Replace(t, CStr(bad(i)), "_")
    Next i
    SanitizeName = modUtil.SafeLeft(Trim$(t), 60)
End Function

Private Function GetVaultSheet() As Worksheet
    On Error Resume Next
    Set GetVaultSheet = ThisWorkbook.Worksheets(VAULT_SHEET)
    On Error GoTo 0
End Function

Private Function GetOrCreateVaultSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetVaultSheet()
    If ws Is Nothing Then
        On Error GoTo Fail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = VAULT_SHEET
        On Error GoTo 0
    End If
    Set GetOrCreateVaultSheet = ws
    Exit Function
Fail:
    ' ハンドラ稼働中は On Error Resume Next が効かず、ここで起きた
    ' エラーは呼び出し元へ飛んで本来の原因を上書きする。
    ' 後始末の前に Resume でハンドラを抜ける(2026-07-30 実機err#462)。
    Resume FailCleanup26
FailCleanup26:
    If Not ws Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If
    Set GetOrCreateVaultSheet = Nothing
End Function

Private Sub RemoveVaultShapes(ByVal ws As Worksheet)
    Dim names() As String
    ReDim names(0 To ws.Shapes.count)
    Dim n As Long: n = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "nxv_" Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp
    Dim i As Long
    For i = 0 To n - 1
        On Error Resume Next
        ws.Shapes(names(i)).Delete
        On Error GoTo 0
    Next i
End Sub
