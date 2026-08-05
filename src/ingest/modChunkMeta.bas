Attribute VB_Name = "modChunkMeta"
Option Explicit

' ============================================================================
' modChunkMeta - チャンクの構造ラベル(section_path)と明示参照(refs_out)の抽出
'                2026-08-05 R17 Phase1(設計書 docs/dev/design_20260805_R17)
' ----------------------------------------------------------------------------
' なぜ要るのか:
'   規程・約款のQAで最も多い質問は「第5条の免責は?」「8条との関係は?」の形を
'   している。ところが取込は章・条の【ラベル】も、本文中の「第8条による」という
'   【参照関係】も捨ててフラットなチャンクにしていた。結果、検索は当たった点
'   だけを返し、条文を読むのに必要な「面」(その条+参照先+別表)が揃わない。
'   ここはその2つを、LLMを1回も呼ばずに(=取込時間も費用も増やさずに)
'   文字列処理だけで取り出す層。
'
' 設計判断:
'   ・VBScript.RegExp / ScriptControl は使わない(政策ブロックのリスクがあり、
'     LibreOffice実行テストに載せられない=回帰を機械で固定できない)。
'     modSparse.DistinctiveKeys の「InStr起点→数字ラン→単位漢字」と同型の
'     素手の走査で書く。
'   ・全角/半角・大小文字の吸収は modSparse.NormalizeForSearch へ一本化する。
'     ここで独自の正規化を書くと、取込側(section_path)と質問側(参照ラベル)で
'     別々の式になり、「第１２条」と「第12条」が別物として静かに外れる。
'   ・breadcrumb はチャンク本文の【1行目】にしか無い(modChunker.FlushBlock が
'     BuildBreadcrumb の結果 + vbLf を先頭に置く)。2行目以降の「【…】」は
'     本文なので見ない。
'   ・抽出元は modShelf.ApplyCrumb を通す【前】の生チャンク。ApplyCrumb は
'     config embed_prefix_breadcrumb が False のとき breadcrumb 行を丸ごと
'     消すため、保存後の my_knowledge から読むと設定ひとつで構造が全滅する。
'   ・Excelオブジェクトに触れない(PURE_LOGIC_MODULES 登録)。シートI/Oは
'     modChunkMetaStore、検索への合流は modAskFocus が持つ。
' ============================================================================

' 1チャンクから拾う参照ラベルの上限。条文の羅列(「第1条から第40条まで」を
' 個別に書いた目次など)で refs_out が1セルを埋め尽くすのを防ぐ。上限に達した
' 後の参照は捨てるが、精読の材料は近傍展開(modAskFocus.NeighborExpand)側でも
' 拾えるので、ここを厚くしても効果は伸びない。
Private Const MAX_REFS As Long = 24

' section_path の長さ上限。breadcrumb の各要素は modChunker.BuildBreadcrumb が
' 80字で切るので、章>条の2階層なら理論上161字。余裕を見て200。
Private Const MAX_PATH As Long = 200

' ----------------------------------------------------------------------------
' ExtractSectionPath - 生チャンク先頭の breadcrumb から「章>条」を取り出す。
'   入力例: "【〔資料〕 > 第3章 総則 > 第12条(免責)】" & vbLf & "本文…"
'   戻り値: "第3章 総則>第12条(免責)"
'   ・先頭要素(資料名)は必ず捨てる。資料名は my_knowledge.source が持って
'     いる情報で、ここに混ぜると同じ条文が資料名の表記ゆれで別ラベルになる。
'   ・breadcrumb が無い(構造チャンク以外)/閉じ括弧が無い/中身が空の
'     どれでも空文字を返す。空文字は「構造が分からない」という正しい答えで、
'     検索側はその行を無いものとして従来どおり動く(フェイルセーフ)。
' ----------------------------------------------------------------------------
Public Function ExtractSectionPath(ByVal rawChunkText As String) As String
    If LenB(rawChunkText) = 0 Then Exit Function
    If Left$(rawChunkText, 1) <> "【" Then Exit Function

    Dim line1 As String
    Dim lf As Long: lf = InStr(rawChunkText, vbLf)
    If lf > 0 Then
        line1 = Left$(rawChunkText, lf - 1)
    Else
        line1 = rawChunkText
    End If

    Dim cl As Long: cl = InStr(line1, "】")
    If cl < 3 Then Exit Function          ' 「【】」だけ/閉じ無し=形式崩れ

    Dim parts() As String
    parts = Split(modSparse.NormalizeForSearch(Mid$(line1, 2, cl - 2)), ">")

    Dim outS As String
    Dim i As Long
    For i = LBound(parts) + 1 To UBound(parts)
        Dim seg As String: seg = Trim$(parts(i))
        If LenB(seg) > 0 Then
            If LenB(outS) > 0 Then outS = outS & ">"
            outS = outS & seg
        End If
    Next i
    ExtractSectionPath = modUtil.SafeLeft(outS, MAX_PATH)
End Function

' ----------------------------------------------------------------------------
' ExtractRefs - 本文中の明示参照を "|" 区切りで返す(重複なし・最大 MAX_REFS)。
'   拾うもの: 「第N条」「第N項」「第N章」「別表N」「様式N」
'             (「別表第2」「様式第3」のように間に「第」が入る書き方も、
'               原文どおりのラベルとして拾う。同じ資料の中では表記が揃って
'               いるのが普通で、勝手に寄せると逆に一致しなくなる)
'   ・自分自身の見出し番号(そのチャンクが第6条なら「第6条」)も混ざるが、
'     ここでは落とさない。落とす判断は「どのチャンクの refs か」を知っている
'     検索側(RefLabelsFor)が持つ。抽出側は本文の事実だけを記録する。
' ----------------------------------------------------------------------------
Public Function ExtractRefs(ByVal bodyText As String) As String
    If LenB(bodyText) = 0 Then Exit Function

    Dim s As String: s = modSparse.NormalizeForSearch(bodyText)
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Function

    Dim box As String
    ScanDaiN s, n, box
    ScanHeadN s, n, "別表", box
    ScanHeadN s, n, "様式", box
    ExtractRefs = box
End Function

' ----------------------------------------------------------------------------
' MetaOf - 1チャンクぶんの section_path と refs_out を一度に返す。
'   取込ループ(modShelf.IngestFile)は残り字数が数百字しかなく(憲章§4-6)、
'   2行の呼び出しを置く余裕が無いためのまとめ役。同じ生テキストから両方を
'   採ることを型で保証する副次効果もある(片方だけ ApplyCrumb 後のテキストを
'   渡す、という取り違えが起こらない)。
' ----------------------------------------------------------------------------
Public Sub MetaOf(ByVal rawChunkText As String, ByRef outPath As String, _
                  ByRef outRefs As String)
    outPath = ExtractSectionPath(rawChunkText)
    outRefs = ExtractRefs(rawChunkText)
End Sub

' ----------------------------------------------------------------------------
' PathHasLabel - section_path が参照ラベルを含むか。
'   単純な部分一致でよい理由: ラベルは必ず数字の直後に単位の漢字が来る形
'   (第12条 / 別表2)なので、「第1条」が「第12条」に誤って当たることはない
'   (第・1・条 の並びは 第・1・2・条 の中に現れない)。逆に「第12条の2」の
'   ような枝番の見出しには当たってほしいので、前方一致では狭すぎる。
' ----------------------------------------------------------------------------
Public Function PathHasLabel(ByVal sectionPath As String, ByVal label As String) As Boolean
    If LenB(sectionPath) = 0 Or LenB(label) = 0 Then Exit Function
    PathHasLabel = (InStr(1, sectionPath, label, vbBinaryCompare) > 0)
End Function

' ----------------------------------------------------------------------------
' RefLabelsFor - refs_out から「そのチャンク自身の見出し番号」を除いたラベル列。
'   第6条の本文には「第6条」自身が何度も出る(見出し行がチャンクに入るため)。
'   除かないと、参照展開が毎回「自分と同じ条の他のチャンク」で上限を使い切り、
'   本来足したい第8条・別表2に届かない。
' ----------------------------------------------------------------------------
Public Function RefLabelsFor(ByVal refsOut As String, ByVal ownPath As String) As String
    If LenB(refsOut) = 0 Then Exit Function

    Dim arr() As String: arr = Split(refsOut, "|")
    Dim outS As String
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Dim t As String: t = Trim$(arr(i))
        If LenB(t) > 0 Then
            If Not PathHasLabel(ownPath, t) Then
                If InStr(1, "|" & outS & "|", "|" & t & "|", vbBinaryCompare) = 0 Then
                    If LenB(outS) > 0 Then outS = outS & "|"
                    outS = outS & t
                End If
            End If
        End If
    Next i
    RefLabelsFor = outS
End Function

' ----------------------------------------------------------------------------
' GraphActive - 参照グラフの層が働く条件(1つでも欠ければ【無操作】)。
'   metaCount=0 は「chunk_meta が無い/まだ再取込していない本棚」で、
'   R17設計書§3の明文の前提(既存資料は再取込で生成・移行処理は書かない)。
'   このとき参照展開も条番号の直接ヒット保証も何もせず、回答は R16 までと
'   1文字も変わらない。この1本を両方の入口(modAskFocus)から呼ぶことで、
'   フェイルセーフの条件が2箇所に割れないようにする。
' ----------------------------------------------------------------------------
'   allowZeroHits(2026-08-05 R17H FA-8 / A-M9): 検索が0件でも、質問が条文・
'   別表・様式を名指ししているときだけ True にしてよい経路がある(直接キーで
'   最大2件だけ材料を用意する modAskRetrieve.EnsureArticleSeed)。条件を
'   もう1本の関数へ割らず、【意図を引数で渡す】形にして単一情報源を保つ。
Public Function GraphActive(ByVal metaCount As Long, ByVal nHits As Long, _
                            Optional ByVal allowZeroHits As Boolean = False) As Boolean
    GraphActive = (metaCount > 0 And (nHits > 0 Or allowZeroHits))
End Function

' ============================================================================
' 内部ヘルパー(走査)
' ============================================================================

' 「第」起点: 第 → 数字ラン → 単位漢字(条/項/章)。
' modSparse.DistinctiveKeys の同名処理と同型(向こうは質問から、こちらは本文
' から拾う。拾う対象の語彙もこちらは「号」を含まない=号は条の内側の細目で、
' 章>条の粒度で作る section_path とは噛み合わないため)。
Private Sub ScanDaiN(ByVal s As String, ByVal n As Long, ByRef box As String)
    Dim p As Long: p = InStr(s, "第")
    Do While p > 0
        Dim j As Long: j = p + 1
        Dim dgt As String: dgt = ""
        Do While j <= n
            Dim dch As String: dch = Mid$(s, j, 1)
            If dch >= "0" And dch <= "9" Then
                dgt = dgt & dch
                j = j + 1
            Else
                Exit Do
            End If
        Loop
        If LenB(dgt) > 0 And j <= n Then
            Dim unit As String: unit = Mid$(s, j, 1)
            If unit = "条" Or unit = "項" Or unit = "章" Then
                AddRef box, "第" & dgt & unit
            End If
        End If
        p = InStr(p + 1, s, "第")
    Loop
End Sub

' 「別表」「様式」起点: 見出し語 → (任意の「第」) → 数字ラン。
Private Sub ScanHeadN(ByVal s As String, ByVal n As Long, ByVal head As String, _
                      ByRef box As String)
    Dim hl As Long: hl = Len(head)
    Dim p As Long: p = InStr(s, head)
    Do While p > 0
        Dim j As Long: j = p + hl
        Dim dai As String: dai = ""
        If j <= n Then
            If Mid$(s, j, 1) = "第" Then
                dai = "第"
                j = j + 1
            End If
        End If
        Dim dgt As String: dgt = ""
        Do While j <= n
            Dim dch As String: dch = Mid$(s, j, 1)
            If dch >= "0" And dch <= "9" Then
                dgt = dgt & dch
                j = j + 1
            Else
                Exit Do
            End If
        Loop
        If LenB(dgt) > 0 Then AddRef box, head & dai & dgt
        p = InStr(p + 1, s, head)
    Loop
End Sub

' "|" 区切りの箱へ重複なしで1件足す(上限 MAX_REFS)。
Private Sub AddRef(ByRef box As String, ByVal label As String)
    If LenB(box) > 0 Then
        If RefCount(box) >= MAX_REFS Then Exit Sub
        If InStr(1, "|" & box & "|", "|" & label & "|", vbBinaryCompare) > 0 Then Exit Sub
        box = box & "|" & label
    Else
        box = label
    End If
End Sub

' 箱に入っている件数("|"の数+1。空なら0)。
Private Function RefCount(ByVal box As String) As Long
    If LenB(box) = 0 Then Exit Function
    RefCount = Len(box) - Len(Replace(box, "|", "")) + 1
End Function
