Attribute VB_Name = "modMode"
Option Explicit

' ============================================================================
' modMode - 回答モードの方針(すぐ聞く / 通常 / 入念)
' ----------------------------------------------------------------------------
' なぜ3段階なのか(実測に基づく)
' ----------------------------------------------------------------------------
' 実物6資料180チャンク31問で採点した結果(tools/bench_retrieval.py):
'
'   手法                        R@1   R@3   R@5   R@10  R@20   ページ
'   単発検索(InStr・空白除去)     84%   84%   90%   90%   94%    81%
'   多クエリ→RRF統合            77%   84%   90%   97%   97%    71%
'
' ここから2つのことが分かる:
'   1. 「LLMに渡す件数」を増やしても効かない。R@5 90% → R@10 90% で頭打ち。
'      Qwen案の「quick 6→10件 / deep 12→24件」は、この実測では0%改善だった。
'      むしろ無関係なチャンクが増えて Lost in the middle を招く。
'   2. 「引き方(角度)」を増やすと効く。多クエリにすると R@10 が 90%→97%。
'      ただし1位の精度は落ちる(84%→77%)。角度を増やすほど、
'      本命以外も上位に混ざるため。
'
' つまりモードごとに最適な戦略が違う:
'   ・すぐ聞く … 少数しか渡さないので R@1 が命。単発検索が最良(84%)
'   ・通常     … 中間。単発+質問拡張
'   ・入念     … 多く渡してLLMと検証段に選ばせるので R@10 が命。
'                多クエリ(97%)+再ランク+下書き→照合
'
' 金融ドメインでは「ここ一番で外さない」ことが信用の全てなので、
' 入念モードは時間を一切気にせず精度側に全振りする。
'
' R4準拠(純ロジック): Excelオブジェクトに触れないので LibreOffice でテスト可能。
' ============================================================================

' モード識別子。文字列はui_state/configと共有するので変更しないこと。
Private Const MODE_QUICK As String = "quick"
Private Const MODE_DEEP As String = "deep"
Private Const MODE_THOROUGH As String = "thorough"

' ----------------------------------------------------------------------------
' Normalize - 文字列を3つのモードのいずれかへ正規化する(既定=すぐ聞く)。
' ----------------------------------------------------------------------------
Public Function Normalize(ByVal mode As String) As String
    Dim m As String: m = LCase$(Trim$(mode))
    If m = MODE_THOROUGH Then
        Normalize = MODE_THOROUGH
    ElseIf m = MODE_DEEP Then
        Normalize = MODE_DEEP
    Else
        Normalize = MODE_QUICK
    End If
End Function

' ----------------------------------------------------------------------------
' NextMode - ボタンを押したときの巡回順(すぐ聞く → 通常 → 入念 → すぐ聞く)。
' ----------------------------------------------------------------------------
Public Function NextMode(ByVal mode As String) As String
    Select Case Normalize(mode)
        Case MODE_QUICK:    NextMode = MODE_DEEP
        Case MODE_DEEP:     NextMode = MODE_THOROUGH
        Case Else:          NextMode = MODE_QUICK
    End Select
End Function

' ----------------------------------------------------------------------------
' Caption - 画面に出す名前。何をするモードかが分かる言い方にする。
'   「速い/遅い」ではなく「何をするか」が違う、と伝えるのが目的。
' ----------------------------------------------------------------------------
Public Function Caption(ByVal mode As String) As String
    Select Case Normalize(mode)
        Case MODE_THOROUGH
            Caption = ChrW(&HD83D) & ChrW(&HDD2C) & " 入念に調べる"
        Case MODE_DEEP
            Caption = ChrW(&HD83D) & ChrW(&HDD0D) & " しっかり調べる"
        Case Else
            Caption = ChrW(&H26A1) & " すぐ聞く"
    End Select
End Function

' ----------------------------------------------------------------------------
' Description - 待ち時間の目安と、そのモードが何をするか(利用者向け)。
'   時間を隠さない。隠すと「固まった」と思われる。
' ----------------------------------------------------------------------------
Public Function Description(ByVal mode As String) As String
    Select Case Normalize(mode)
        Case MODE_THOROUGH
            Description = "多方向から検索 → 関連度を精査 → 下書き → 資料と1行ずつ照合。" & _
                          "数分かかりますが、精度を最優先します。"
        Case MODE_DEEP
            Description = "質問を分析して検索 → 下書き → 資料と照合。1〜2分ほど。"
        Case Else
            Description = "そのまま検索して答えます。数秒〜20秒ほど。"
    End Select
End Function

' ----------------------------------------------------------------------------
' TopK - LLMへ渡すチャンク数。
'   実測で「増やしても効かない」ことが分かっているので闇雲に増やさない
'   (R@5 90% → R@10 90%)。入念だけ多クエリの成果を拾える幅を取る。
' ----------------------------------------------------------------------------
Public Function TopK(ByVal mode As String, ByVal quickK As Long, _
                     ByVal deepK As Long, ByVal thoroughK As Long) As Long
    Select Case Normalize(mode)
        Case MODE_THOROUGH: TopK = thoroughK
        Case MODE_DEEP:     TopK = deepK
        Case Else:          TopK = quickK
    End Select
    If TopK < 1 Then TopK = 6
End Function

' ----------------------------------------------------------------------------
' 各段を実行するか。入念は「時間より精度」なので全段を必ず通す。
' ----------------------------------------------------------------------------

' 質問拡張(1つの質問を複数の角度へ展開する)
Public Function UseExpand(ByVal mode As String, ByVal cfgEnabled As Boolean, _
                          ByVal quickOptIn As Boolean) As Boolean
    Select Case Normalize(mode)
        Case MODE_THOROUGH: UseExpand = True
        Case MODE_DEEP:     UseExpand = cfgEnabled
        Case Else:          UseExpand = cfgEnabled And quickOptIn
    End Select
End Function

' 再ランク(候補を関連度順に並べ直す)
Public Function UseRerank(ByVal mode As String, ByVal cfgEnabled As Boolean, _
                          ByVal quickOptIn As Boolean) As Boolean
    Select Case Normalize(mode)
        Case MODE_THOROUGH: UseRerank = True
        Case MODE_DEEP:     UseRerank = cfgEnabled
        Case Else:          UseRerank = cfgEnabled And quickOptIn
    End Select
End Function

' 検証段(下書きを資料と1行ずつ突き合わせる)。ここが「間違えない」の要。
Public Function UseVerify(ByVal mode As String) As Boolean
    Dim m As String: m = Normalize(mode)
    UseVerify = (m = MODE_DEEP Or m = MODE_THOROUGH)
End Function

' 拡張の軽量版を使うか(すぐ聞くのときだけ。入念で軽くしては意味が無い)
Public Function UseLightExpand(ByVal mode As String, ByVal cfgLight As Boolean) As Boolean
    UseLightExpand = (Normalize(mode) = MODE_QUICK) And cfgLight
End Function

' サブクエリ数。入念は角度の数がそのまま精度になる(実測 R@10 90%→97%)。
Public Function SubQueryCount(ByVal mode As String, ByVal normalN As Long, _
                              ByVal thoroughN As Long) As Long
    If Normalize(mode) = MODE_THOROUGH Then
        SubQueryCount = thoroughN
    Else
        SubQueryCount = normalN
    End If
    If SubQueryCount < 1 Then SubQueryCount = 1
End Function
