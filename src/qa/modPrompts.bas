Attribute VB_Name = "modPrompts"
Option Explicit

' ============================================================================
' modPrompts - プロンプト文字列の組み立て(純文字列モジュール)
' ----------------------------------------------------------------------------
' 役割:
'   ⚡すぐ聞く/🔍しっかり調べる(下書き・検証)/バッチ富化の4種のプロンプトを
'   組み立てて文字列として返す(MASTER_SPEC §7.3)。
'
' 設計判断:
'   ・R4準拠(§3): Worksheets/Range(/Application./ThisWorkbook/MsgBox/
'     ActiveSheet の各トークンはこのモジュールのソースに一切書かない。ただし
'     modConfig.GetString/GetLong の呼び出しはR4対象外として明示許可されている
'     (MASTER_SPEC発注時の指示:「ExcelトークンがmodPrompts自身のソースに
'     現れなければR4適合」)。modConfig自体はシートを読むが、それはmodConfig.bas
'     の中の話であり、modPromptsのソースにはExcelオブジェクトトークンが一切
'     現れないためLintのトークン検査は通る。
'   ・出典指示形式(MASTER_SPEC §7.3): 本棚由来は
'     "[本棚:ファイル名 p.ページ番号]"、パック由来は
'     "[パック(作成者):ファイル名]"。origin文字列が"pack:"で始まるかどうかで
'     判定する(§4: origin = self | pack:<作成者名>)。
'   ・「資料に無いことは『資料には見当たらない』と述べる」指示は、QAで
'     実際にLLMが本棚抜粋から回答するBuildQuickPrompt/BuildDeepDraftPrompt/
'     BuildDeepVerifyPromptの全モードに挿入する。BuildEnrichPromptは既知の
'     チャンク本文を要約するだけの用途で「資料に無い」概念が無いため対象外。
'   ・本文組み込みの打ち切り: hits()の各full_text(チャンク本文全体。最大
'     32000字)を連結し、config max_context_chars(既定40000)を超える手前で
'     打ち切り、打ち切った場合のみ末尾に「(一部省略)」を挿入する(§7.3)。
'   ・Wave3 PM裁定1: modTypes.Hit に full_text が追加され、modRetrieve.Search が
'     my_knowledge読取時に格納するようになったため、「本棚抜粋」の本文として
'     full_text を使う(旧実装はpreview=先頭120字しか渡せず根拠が不十分だった)。
'     preview は出典先出し表示(modUIMain.RenderSourcesPreview)専用としてHit型に
'     残っており、プロンプト本文には使わない。ただしfull_textが空(旧データ・
'     テストダブル等)の場合はpreviewへフォールバックし、空の抜粋を混ぜない。
'   ・深掘り候補(裁定D11): 利用者に実際に表示される回答を生成する
'     BuildQuickPrompt(すぐ聞くの最終応答)とBuildDeepVerifyPrompt(しっかり
'     調べるの最終応答)の末尾に、[[FOLLOWUP: 候補1 | 候補2]] 形式で深掘り質問
'     候補を2つ付けさせる指示行(FollowupInstruction)を追加する。マーカーは
'     modAsk側でパース・除去され「深掘り候補(『続けて質問』でそのまま聞けます)」
'     ブロックに整形されるため、利用者の目に生マーカーは触れない(LLMが指示を
'     無視しても候補ブロックが付かないだけで壊れない)。BuildDeepDraftPromptは
'     表示されない中間生成物(検証段の入力)のため対象外。
' ============================================================================

' 分解した論点のうち「資料を確認できなかった」節の文言(2026-08-05 R16-3A)。
' 節を作る側(BuildPartSection)と、それを消させない側(BuildMergePrompt)と、
' 成否を数える側(modAskMulti)が同じ文字列を見るための単一情報源。
Public Const PART_FAIL_TEXT As String = "資料からは確認できませんでした(検索失敗)"

' 深掘り(続けて質問)のターンだけ下書き段へ足す1文(2026-08-05 R16-3D)。
' 検索側は既出チャンクを降格して新しい材料を渡すが、材料が変わっても
' 「前回と同じ構成でもう一度説明する」癖は指示しないと直らない。
Private Const FOLLOWUP_DEPTH_LINE As String = "・前回の回答と重複する説明は繰り返さず、新しい詳細・根拠・例外を優先すること。"

Public Function BuildQuickPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long, _
                                 Optional ByVal strictGrounding As Boolean = False, _
                                 Optional ByVal answerTags As Boolean = False) As String
    Dim lang As String
    lang = SafeAnswerLanguage()
    Dim maxChars As Long
    maxChars = SafeMaxContextChars()
    Dim ctx As String
    ctx = BuildSourceBlock(hits, nHits, maxChars)

    Dim sb As String
    sb = "以下の本棚抜粋だけを根拠に、質問へ" & lang & "で回答してください(「すぐ聞く」モード=速さと簡潔さ優先)。" & vbLf
    sb = sb & StyleInstruction() & vbLf
    sb = sb & CitationInstruction() & vbLf
    sb = sb & NotFoundInstruction() & vbLf
    sb = sb & DomainGuardInstruction() & vbLf
    If strictGrounding Then sb = sb & GroundingInstruction() & vbLf
    If answerTags Then sb = sb & AnswerTagsInstruction() & vbLf
    sb = sb & vbLf
    sb = sb & UserContextBlock()
    sb = sb & "## 本棚抜粋" & vbLf & ctx & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf
    sb = sb & vbLf & FollowupInstruction() & vbLf
    BuildQuickPrompt = sb
End Function

' digest(2026-08-03 R14-8a): 入念モードが先に作る「資料の要点」。空なら従来と完全に
'   同じプロンプトになる。要点を足したぶんだけ原文の枠を減らし、合計が
'   max_context_chars を超えないようにする(減らさないと入念だけ本文が2倍になり、
'   上限で黙って切れて後半の資料が丸ごと消える)。
Public Function BuildDeepDraftPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long, ByVal history As String, _
                                     Optional ByVal strictGrounding As Boolean = False, _
                                     Optional ByVal answerTags As Boolean = False, _
                                     Optional ByVal digest As String = "") As String
    Dim lang As String
    lang = SafeAnswerLanguage()
    Dim maxChars As Long
    maxChars = SafeMaxContextChars()
    Dim budget As Long
    budget = maxChars - Len(digest)
    If budget < 2000 Then budget = 2000
    Dim ctx As String
    ctx = BuildSourceBlock(hits, nHits, budget)

    Dim sb As String
    sb = "あなたは社内の資料検索AIアシスタントです。「しっかり調べる」モードの下書き回答を作成します。" & _
         "以下の本棚抜粋を根拠に、質問へ" & lang & "で丁寧に、根拠を示しながら回答してください。" & vbLf
    sb = sb & CitationInstruction() & vbLf
    sb = sb & NotFoundInstruction() & vbLf
    sb = sb & DomainGuardInstruction() & vbLf
    If strictGrounding Then sb = sb & GroundingInstruction() & vbLf
    If answerTags Then sb = sb & AnswerTagsInstruction() & vbLf
    ' R16-3D: 深掘りのターンだけ「前回の繰り返しをしない」を明示する。
    If modFollowup.IsFollowupTurn() Then sb = sb & FOLLOWUP_DEPTH_LINE & vbLf
    If LenB(history) > 0 Then
        sb = sb & vbLf & "## これまでの会話(参考。続きの質問なら踏まえて回答する)" & vbLf & history
    End If
    sb = sb & vbLf & UserContextBlock()
    If LenB(Trim$(digest)) > 0 Then
        sb = sb & "## 資料の要点(質問に関係する部分を先に整理したもの。" & _
             "数値・条文は必ず下の本棚抜粋の原文で裏を取ること)" & vbLf & digest & vbLf & vbLf
    End If
    sb = sb & "## 本棚抜粋" & vbLf & ctx & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf
    sb = sb & vbLf & "(この回答は下書きです。この後、別の検証ステップで事実確認されます。" & _
        "根拠が弱い部分は無理に断定せず、その旨を書いてください。)"
    BuildDeepDraftPrompt = sb
End Function

' critique(2026-08-03 R14-8a): 入念モードの自己批判段が挙げた指摘。空なら
'   従来と完全に同じプロンプトになる(深掘りは指摘段を持たない)。
Public Function BuildDeepVerifyPrompt(ByVal q As String, ByVal draft As String, hits() As Hit, ByVal nHits As Long, _
                                      Optional ByVal strictGrounding As Boolean = False, _
                                      Optional ByVal answerTags As Boolean = False, _
                                      Optional ByVal critique As String = "") As String
    Dim lang As String
    lang = SafeAnswerLanguage()
    Dim maxChars As Long
    maxChars = SafeMaxContextChars()
    Dim ctx As String
    ctx = BuildSourceBlock(hits, nHits, maxChars)

    Dim sb As String
    sb = "下書き回答を本棚抜粋と照合し、" & lang & "で最終回答を作成してください。" & _
         "本棚抜粋で裏付けられない断定や事実と異なる記載は、修正するか削除してください。" & vbLf
    sb = sb & StyleInstruction() & vbLf
    sb = sb & CitationInstruction() & vbLf
    ' 2026-07-28(解説書 §11-9): 検証段にも DomainGuard を入れる。下書き段
    ' (quick/deep draft)には入っているのに検証段だけ抜けており、数値の厳格さと
    ' 「(要確認)」の付与ルールを知らないモデルが最終回答を書き直していた。下書きが
    ' 正しく付けた (要確認) を検証段が「不要な但し書き」と判断して落とすと、
    ' 【確認が要る数字が、確認不要の顔をして残る】。保険の金額・期限・料率で
    ' これが起きると実害が出る。
    sb = sb & DomainGuardInstruction() & vbLf
    sb = sb & NotFoundInstruction() & vbLf
    If strictGrounding Then sb = sb & GroundingInstruction() & vbLf
    If answerTags Then sb = sb & AnswerTagsInstruction() & vbLf
    sb = sb & vbLf
    sb = sb & "## 本棚抜粋" & vbLf & ctx & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf
    sb = sb & vbLf & "## 下書き回答" & vbLf & draft & vbLf
    If LenB(Trim$(critique)) > 0 Then
        sb = sb & vbLf & "## 査読で挙がった指摘" & vbLf & critique & vbLf
        sb = sb & vbLf & "(指示: 以下の指摘を踏まえて修正せよ。指摘が抜粋と食い違う場合は" & _
            "抜粋を優先し、元の記述を残してよい。指摘そのものへの言及・弁明は" & _
            "最終回答に書かないこと。)" & vbLf
    End If
    sb = sb & vbLf & "(注意: 下書きに付いている「(要確認)」は、根拠が弱い箇所に" & _
        "意図的に付けたものです。抜粋で裏付けられない限り外さないでください。)" & vbLf
    sb = sb & vbLf & "(指示: 検証済みの最終回答のみを出力してください。下書きとの差分説明や、" & _
        "検証過程の説明は不要です。)"
    sb = sb & vbLf & FollowupInstruction() & vbLf
    BuildDeepVerifyPrompt = sb
End Function

' ----------------------------------------------------------------------------
' BuildExpandPrompt - クエリ拡張段(多段RAG・設計書§C-1)のプロンプト。
'   出力契約: <standalone>/<subqueries>/<hyde>。lightMode=Trueはstandaloneのみ
'   (すぐ聞くモードの速度優先)。パースはmodRagParse.ParseExpand(寛容退化)。
' ----------------------------------------------------------------------------
Public Function BuildExpandPrompt(ByVal q As String, ByVal history As String, _
                                  ByVal subqueryCount As Long, ByVal lightMode As Boolean) As String
    Dim nSub As Long: nSub = subqueryCount
    If nSub < 0 Then nSub = 0
    If nSub > 8 Then nSub = 8

    Dim sb As String
    sb = "あなたは社内資料検索システムの検索プランナーです。利用者の質問を、" & _
         "ベクトル検索でヒットしやすい形に変換してください。" & vbLf
    If LenB(history) > 0 Then
        sb = sb & vbLf & "## これまでの会話(代名詞や『それ』の解決に使う)" & vbLf & history & vbLf
    End If
    sb = sb & vbLf & "## 利用者の質問" & vbLf & q & vbLf & vbLf
    sb = sb & "## 出力形式(この形式のみで出力。説明文・前置きは一切禁止)" & vbLf
    sb = sb & "<standalone>会話の文脈を織り込み、単体で意味が通る独立した質問文</standalone>" & vbLf
    If Not lightMode Then
        If nSub > 0 Then
            sb = sb & "<subqueries>言い換えや下位概念・関連語での検索文を" & nSub & _
                 "本、「 | 」区切りで</subqueries>" & vbLf
        End If
        sb = sb & "<hyde>この質問に理想的に答える文章を1段落(実在資料の文体を想像した検索用の仮回答)</hyde>" & vbLf
    End If
    BuildExpandPrompt = sb
End Function

' ----------------------------------------------------------------------------
' BuildRerankPrompt - 再ランク段(多段RAG・設計書§C-3)のプロンプト。
'   候補に1..nの連番を振り、<rank>3,1,7</rank>形式のみで返させる。
'   パースはmodRagParse.ParseRankOrder(寛容退化: 崩れたら元順維持)。
' ----------------------------------------------------------------------------
Public Function BuildRerankPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long, _
                                  ByVal maxContextChars As Long) As String
    Dim lim As Long: lim = maxContextChars
    If lim <= 0 Then lim = 40000

    Dim sb As String
    sb = "あなたは社内資料検索システムの関連度審査員です。以下の候補チャンクを、" & _
         "質問への関連度が高い順に並べ替えてください。" & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf & vbLf
    sb = sb & "## 候補チャンク" & vbLf

    Dim used As Long: used = Len(sb)
    Dim i As Long
    For i = 1 To nHits
        Dim entry As String
        ' R34 B2(裁定・司令塔承認済みの凍結モジュール最小手術): 抜粋300→700字。
        entry = "[" & i & "] " & SourceTag(hits(i)) & " " & _
                Replace(modUtil.SafeLeft(SourceBody(hits(i)), 700), vbLf, " ") & vbLf
        If used + Len(entry) > lim Then
            sb = sb & "(以下省略)" & vbLf
            Exit For
        End If
        sb = sb & entry
        used = used + Len(entry)
    Next i

    sb = sb & vbLf & "## 出力形式(この形式のみ。説明禁止)" & vbLf
    sb = sb & "関連度が高い順に候補番号をカンマ区切りで並べ、次の形式で出力: <rank>3,1,7</rank>" & vbLf
    sb = sb & "質問と無関係な候補は含めなくてよい。最低1件は含めること。"
    BuildRerankPrompt = sb
End Function

' ----------------------------------------------------------------------------
' BuildSourceDigestPrompt - 入念モード(1)資料の要点整理(2026-08-03 R14-8a)
' ----------------------------------------------------------------------------
' 本棚抜粋を1回のLLM呼び出しで「質問に関係する部分だけ」資料ごと2～3行へ畳む。
' 下書き段へは この要点 + 上位の原文 のハイブリッドを渡す(要点だけだと数値や
' 条文が丸まり、原文だけだと件数が増えたとき中盤が読み飛ばされる)。出力の粒度を
' JSON等にしないのは、この結果を人ではなく次のプロンプトが読むからで、素の日本語
' のまま差し込むのがいちばん壊れない(BuildEnrichPromptがJSONなのは読み手がVBAの
' パーサだから)。
Public Function BuildSourceDigestPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long) As String
    Dim maxChars As Long
    maxChars = SafeMaxContextChars()

    Dim sb As String
    sb = "以下は社内資料から検索で拾った抜粋です。質問に答えるうえで意味のある部分だけを、" & _
         "資料ごとに2～3行へ要約してください。" & vbLf
    sb = sb & "・各行の末尾に、その内容の出典タグ([本棚:ファイル名 p.ページ番号] または " & _
         "[パック(作成者名):ファイル名])を抜粋にあるとおり書き写すこと。" & vbLf
    sb = sb & "・条文番号・日数・金額・料率・期限は要約せず、抜粋の値をそのまま書き写すこと。" & vbLf
    sb = sb & "・質問と関係の無い抜粋は「(質問とは関係なし)」の1行だけにすること。" & vbLf
    sb = sb & "・抜粋に書かれていないことは推測で補わないこと。" & vbLf
    sb = sb & "・前置き・総括・結論は書かず、資料ごとの要約だけを出力すること。" & vbLf & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf & vbLf
    sb = sb & "## 本棚抜粋" & vbLf & BuildSourceBlock(hits, nHits, maxChars) & vbLf
    BuildSourceDigestPrompt = sb
End Function

' ----------------------------------------------------------------------------
' BuildCritiquePrompt - 入念モード(3)自己批判(2026-08-03 R14-8a)
' ----------------------------------------------------------------------------
' 下書きの問題点【だけ】を挙げさせる。書き直しを同時に頼まないのが要点で、「指摘
' して直せ」と言うとモデルは自分の文章を守るために指摘を軽くする。指摘だけを吐かせ
' 別の呼び出し(検証段)に直させたほうが、実際に消える未検証の断定の数が多い。
Public Function BuildCritiquePrompt(ByVal q As String, ByVal draft As String, _
                                    hits() As Hit, ByVal nHits As Long) As String
    Dim maxChars As Long
    maxChars = SafeMaxContextChars()

    Dim sb As String
    sb = "あなたは社内資料の査読者です。下書き回答を本棚抜粋と突き合わせ、" & _
         "問題点だけを指摘してください。書き直した文章や修正版の提示は禁止です。" & vbLf
    sb = sb & "見る観点は次の4つ:" & vbLf
    sb = sb & "1. 未検証の断定 … 抜粋で裏付けられないのに言い切っている" & vbLf
    sb = sb & "2. 出典不備 … 出典が付いていない/別の資料の出典が付いている/ページが違う" & vbLf
    sb = sb & "3. 論点漏れ … 質問のうち答えていない部分がある" & vbLf
    sb = sb & "4. 憶測 … 抜粋に無い数値・条件・例外を補っている" & vbLf & vbLf
    sb = sb & "## 出力形式(この形式のみ。前置き・総括・修正文は書かない)" & vbLf
    sb = sb & "1. [観点] 該当箇所の引用(20字程度) → 何が問題か" & vbLf
    sb = sb & "2. [観点] …" & vbLf
    sb = sb & "問題が1つも無ければ「指摘なし」とだけ出力してください。" & vbLf & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf & vbLf
    sb = sb & "## 本棚抜粋" & vbLf & BuildSourceBlock(hits, nHits, maxChars) & vbLf
    sb = sb & "## 下書き回答" & vbLf & draft & vbLf
    BuildCritiquePrompt = sb
End Function

' ----------------------------------------------------------------------------
' BuildDecomposePrompt - 複合質問の分解判定(2026-08-05 R16-3A・段0)
' ----------------------------------------------------------------------------
' 「AとBの違いと、Cの手続き」を1本のクエリで検索すると、全部の論点に薄く当たった
' 資料が上位に来てどの論点も詰め切れない。先に論点へ割るための1回きりの判定。
' 出力はパーサ(modRagParse.ParseDecomposeVerdict/ParseParts)が読む3タグで、
' 崩れても single へ寛容退化する=この段が失敗しても従来の入念フローが動く。
Public Function BuildDecomposePrompt(ByVal q As String, ByVal history As String, _
                                     ByVal maxParts As Long) As String
    Dim lim As Long: lim = maxParts
    If lim < 2 Then lim = 2
    If lim > 5 Then lim = 5

    Dim sb As String
    sb = "あなたは社内資料検索システムの質問アナリストです。利用者の質問が" & _
         "「1つの論点か」「複数の論点を含むか」「読み方が定まらないか」を判定してください。" & vbLf
    sb = sb & "・parts = 独立して調べるべき論点が2つ以上ある(例: AとBの違い【と】Cの手続き)。" & vbLf
    sb = sb & "・clarify = 質問の読み方が複数あり、どれを調べるべきか決められない。" & vbLf
    sb = sb & "・global = 全体像・一覧・「全部教えて」型(資料を章ごとに俯瞰する必要がある)。" & vbLf
    sb = sb & "・single = 上のどれでもない(1つの論点として調べられる)。迷ったら single。" & vbLf
    sb = sb & "・分解するときは元の質問の言葉を使い、それ単体で資料を検索できる文にすること。" & vbLf
    sb = sb & "・論点は最大" & lim & "個まで。1つの論点を言い換えて水増ししないこと。" & vbLf
    If LenB(history) > 0 Then
        sb = sb & vbLf & "## これまでの会話(代名詞や『それ』の解決に使う)" & vbLf & history & vbLf
    End If
    sb = sb & vbLf & "## 利用者の質問" & vbLf & q & vbLf & vbLf
    sb = sb & "## 出力形式(この形式のみで出力。説明文・前置きは一切禁止)" & vbLf
    sb = sb & "<verdict>single/parts/clarify/global のどれか1つ</verdict>" & vbLf
    sb = sb & "<parts>論点1 | 論点2 | 論点3</parts>" & vbLf
    sb = sb & "<options>読み方の候補1 | 候補2 | 候補3</options>" & vbLf
    sb = sb & "(parts のときは parts だけ、clarify のときは options だけを埋め、" & _
         "single と global ではどちらも空にする。)" & vbLf
    BuildDecomposePrompt = sb
End Function

' ----------------------------------------------------------------------------
' BuildPartDraftPrompt - 分解した1論点だけの副下書き(2026-08-05 R16-3A)
' ----------------------------------------------------------------------------
' 「その論点だけに答えさせる」が全て。元の質問全体に答え始めると、他の論点と
' 重複した薄い文章がN本できるだけになる。見出しを付けさせないのは統合段が
' ■見出し=副質問 で組み直すため(2箇所で作ると二重になる)。
Public Function BuildPartDraftPrompt(ByVal q As String, ByVal part As String, _
                                     ByVal idx As Long, hits() As Hit, ByVal nHits As Long, _
                                     Optional ByVal strictGrounding As Boolean = False, _
                                     Optional ByVal answerTags As Boolean = False) As String
    Dim lang As String
    lang = SafeAnswerLanguage()
    Dim ctx As String
    ctx = BuildSourceBlock(hits, nHits, SafeMaxContextChars())

    Dim sb As String
    sb = "以下の本棚抜粋だけを根拠に、【この論点だけ】へ" & lang & "で答えてください。" & _
         "これは1つの質問を論点ごとに分けて調べているうちの" & idx & "番目で、" & _
         "後で他の論点の答えと統合されます。" & vbLf
    sb = sb & "・元の質問の全体には答えないこと(他の論点は別に調べています)。" & vbLf
    sb = sb & "・250字程度。結論を先に書き、前置き・総括・締めの挨拶は書かない。" & vbLf
    sb = sb & "・見出し(■で始まる行)は付けないこと。統合するときに付けます。" & vbLf
    sb = sb & CitationInstruction() & vbLf
    sb = sb & NotFoundInstruction() & vbLf
    sb = sb & DomainGuardInstruction() & vbLf
    If strictGrounding Then sb = sb & GroundingInstruction() & vbLf
    If answerTags Then sb = sb & AnswerTagsInstruction() & vbLf
    sb = sb & vbLf
    sb = sb & "## 元の質問(文脈の参考。答えるのは下の論点だけ)" & vbLf & q & vbLf & vbLf
    sb = sb & "## この論点" & vbLf & part & vbLf & vbLf
    sb = sb & "## 本棚抜粋" & vbLf & ctx & vbLf
    BuildPartDraftPrompt = sb
End Function

' ----------------------------------------------------------------------------
' BuildMergePrompt - 論点ごとの下書きを1つの回答へ統合する(2026-08-05 R16-3A)
' ----------------------------------------------------------------------------
' 出典タグが1つでも消えると、後段の機械的突合(modAskThorough.
' AnnotateAgainstHits)が「タグが無い回答」として何も言えなくなり、根拠の追跡が
' 丸ごと切れる。だから最も強く書くのは文章の巧さではなく「タグを一字一句残せ」。
' 新事実の追加禁止も同じ理由(統合は並べ替えと接続であって取材ではない)。
' StyleInstruction を入れないのは、あちらの長さ指示(200～400字)が論点N本ぶんの
' 本文と正面から衝突するため(記法だけをここで短く指示する)。
Public Function BuildMergePrompt(ByVal q As String, ByVal sections As String) As String
    Dim lang As String
    lang = SafeAnswerLanguage()

    Dim sb As String
    sb = "以下は、1つの質問を論点ごとに分けて調べた下書きです。これを1つの回答へ" & _
         lang & "でまとめてください。" & vbLf
    sb = sb & "【厳守】出典タグ([本棚:ファイル名 p.ページ番号] と " & _
         "[パック(作成者名):ファイル名])は一字一句そのまま残すこと。" & _
         "書き換え・削除・末尾へのまとめ直しは禁止。" & vbLf
    sb = sb & "【厳守】下書きに書かれていない事実・数値・条件を足さないこと。" & vbLf
    sb = sb & "【厳守】「" & PART_FAIL_TEXT & "」の節は、その文言のまま残すこと" & _
         "(調べられなかったことを消さない)。" & vbLf
    sb = sb & "・論点ごとの見出しは「■」で始まる行として残し、順番も変えないこと。" & vbLf
    sb = sb & "・重複する説明は1つにまとめてよい(そのとき出典タグは両方とも残す)。" & vbLf
    sb = sb & "・最初の1～2行で、質問全体に対する結論を先に書くこと。" & vbLf
    sb = sb & "・Markdown記号(#、**、`、表)は使わない。箇条書きは「・」、" & _
         "最重要語だけ【 】で囲む。ブロックの間は空行1つ。" & vbLf
    sb = sb & vbLf & "## 元の質問" & vbLf & q & vbLf
    sb = sb & vbLf & "## 論点ごとの下書き" & vbLf & sections & vbLf
    sb = sb & vbLf & FollowupInstruction() & vbLf
    BuildMergePrompt = sb
End Function

Public Function BuildEnrichPrompt(ByVal batchText As String) As String
    Dim lang As String
    lang = SafeAnswerLanguage()

    Dim sb As String
    sb = "以下は社内資料から抽出した複数のチャンクです。それぞれに要約(30字程度)と" & _
         "キーワード(2～5個、カンマ区切り)を" & lang & "で付けてください。" & vbLf
    sb = sb & "出力は必ず次のJSON配列形式のみで返してください" & _
         "(前後に説明文やコードフェンス、余計な文字を付けない):" & vbLf
    sb = sb & "[{""i"":1,""summary"":""..."",""keywords"":""a,b""}," & _
         "{""i"":2,""summary"":""..."",""keywords"":""a,b""}]" & vbLf & vbLf
    sb = sb & "## チャンク一覧" & vbLf & batchText
    BuildEnrichPrompt = sb
End Function

' ----------------------------------------------------------------------------
' BuildQuestionsPrompt - 質問例のオンデマンド生成(2026-08-03 R14-7a)
' ----------------------------------------------------------------------------
' シード資料が無い端末でも「押すだけで聞ける質問」を出すための1回きりの
' 生成。呼び出し元(modStarter)が「資料名: 冒頭の抜粋」を1行ずつまとめた
' 文字列を渡す(BuildEnrichPromptと同型)。出力は1行1問・番号無しの
' プレーンテキストで、modRagParse.ParseQuestionLinesが解析する。
Public Function BuildQuestionsPrompt(ByVal sourceDigest As String) As String
    Dim lang As String
    lang = SafeAnswerLanguage()

    Dim sb As String
    sb = "以下は社内資料の一覧(資料名と冒頭の抜粋)です。それぞれの資料の内容" & _
         "だけで答えられる、実用的な質問を" & lang & "で5件作ってください。" & vbLf
    sb = sb & "・1行に1問だけ書くこと(番号・記号・箇条書き記号を付けない)。" & vbLf
    sb = sb & "・抜粋に書かれていないことを聞く質問は作らないこと。" & vbLf
    sb = sb & "・質問文以外(前置き・見出し・総括)は一切書かないこと。" & vbLf
    sb = sb & "・1問は40字程度までの短い質問文にすること。" & vbLf & vbLf
    sb = sb & "## 資料一覧" & vbLf & sourceDigest
    BuildQuestionsPrompt = sb
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: modPromptsの公開契約はBuild*関数群のみ)
' ----------------------------------------------------------------------------

' SafeAnswerLanguage/SafeMaxContextChars:
'   modConfig.GetString/GetLong への呼び出しを1行スコープのOn Errorで守る
'   (R5: 直後にOn Error GoTo 0)。modConfig自体はconfigシートが無くても既定値へ
'   フォールバックするが、modPromptsはR4のPure Logicモジュール群(vba_lint.py
'   PURE_LOGIC_MODULES / run_lo_tests.py PURE_ALLOWLIST)の一員として、modConfig
'   モジュールそのものが読み込まれていない実行環境(LibreOffice純ロジックテストの
'   一時ライブラリ等)でも動く必要がある。その環境では "modConfig.GetString" の
'   呼び出し自体が実行時エラー(Err=420 Invalid object reference)になることを実測
'   で確認したため、On Errorで包み失敗時は契約既定値へ倒す。Excel実機(modConfigが
'   常に存在する環境)では通常どおりconfigシートの値を返す(エラー時のみ働く)。
Private Function SafeAnswerLanguage() As String
    Dim v As String: v = "日本語"
    On Error Resume Next
    v = modConfig.GetString("answer_language", "日本語")
    On Error GoTo 0
    If LenB(v) = 0 Then v = "日本語"
    SafeAnswerLanguage = v
End Function

Private Function SafeMaxContextChars() As Long
    Dim v As Long: v = 40000
    On Error Resume Next
    v = modConfig.GetLong("max_context_chars", 40000)
    On Error GoTo 0
    If v <= 0 Then v = 40000
    SafeMaxContextChars = v
End Function

' NotebookLM級の出力品質指示(利用者に表示されるQuick/DeepVerifyのみに注入)。
' 重要: この画面(Excel Shape)はMarkdownを描画しないため、## や ** は禁止し、
' プレーンテキストで視覚構造を作る記法(■/・/【】)を強制する。
Private Function StyleInstruction() As String
    StyleInstruction = _
        "【あなたの人格】あなたはMS&ADの最上位ナレッジコンシェルジュです。" & _
        "プロフェッショナルで、簡潔で、温かい。機械的な言い回しはしない。" & vbLf & _
        "【意図の深読み】質問の言葉面だけでなく「質問者が実務で何に困っているか」を一歩深く" & _
        "解釈し、その課題に効く回答をする(解釈がぶれる場合は最有力の解釈で答え、末尾に別解釈を1行)。" & vbLf & _
        "【結論先行】必ず最初の1～2行で結論を言い切る。前置き・挨拶・言い訳から始めない。" & vbLf & _
        "【構成】結論 → 根拠や詳細(箇条書き) → 注意点・例外(あれば) の順。" & vbLf & _
        "【記法・厳守】Markdown記号(#、**、`、表)は一切使わない(この画面では装飾されず崩れて見える)。" & _
        "代わりに: 見出しは「■ 」で始める / 箇条書きは「・」 / 最重要語だけ【 】で囲む / " & _
        "ブロックの間は空行1つ。1ブロックは3行以内。" & vbLf & _
        "【長さ】全体をおおむね200～400字(複雑な質問でも600字まで)。同じ内容の言い換え、" & _
        "冗長な前置き、締めの挨拶は書かない。短くても情報が濃いことが最高の親切。" & vbLf & _
        "【平易さ】保険・社内用語には短い補足を()で添え、初めて読む人にも一度で伝わる言葉を選ぶ。" & vbLf & _
        "【理想の出力例(One-Shot。この型・トーン・粒度に従う)】" & vbLf & _
        "【結論】〇〇の申請には、規定第X条に基づきAとBの手続きが必要です。[本棚:規約集 p.12]" & vbLf & _
        "■ 必要な手順" & vbLf & _
        "・手順1: 〇〇を提出する。[本棚:規約集 p.12]" & vbLf & _
        "・手順2: 〇〇の承認を得る(承認者は課長職以上)。[パック(山田):承認フロー]" & vbLf & _
        "■ 注意点" & vbLf & _
        "・提出期限は事由発生からX日以内。過ぎた場合は個別協議になります。[本棚:規約集 p.13]"
End Function

Private Function CitationInstruction() As String
    CitationInstruction = "回答の根拠として使った情報には、その文の直後に必ず出典を付けてください。" & vbLf & _
        "本棚に自分で入れた資料は [本棚:ファイル名 p.ページ番号] の形式、" & vbLf & _
        "他の人から受け取ったパック由来の資料は [パック(作成者名):ファイル名] の形式で示してください。" & vbLf & _
        "出典は情報と1対1で紐づけ、まとめて末尾に並べるだけの書き方はしないでください。"
End Function

Private Function NotFoundInstruction() As String
    NotFoundInstruction = "本棚抜粋に書かれていないことは、推測で埋めずに" & _
        "「資料には見当たらない」とはっきり述べてください。" & _
        "やむを得ず推測で補う場合は、それが推測であることを明示してください。"
End Function

' 金融・保険ドメインのガードレール(2026-07-26)。社内の既存RAGツールとの差は
' 「速さ」だけでは作れない。約款・規程の条文番号や日数・金額を記憶で補って答えると
' 実務では致命傷になる。「どこまでが資料の裏付けで、どこからが確認が必要か」を
' 回答自身に語らせることが、利用者が正誤を判断できる唯一の現実的な手段になる。
Private Function DomainGuardInstruction() As String
    DomainGuardInstruction = _
        "【数値の厳格性】条文番号・日数・金額・料率・期限は、抜粋に書かれた値だけを" & _
        "そのまま使う。抜粋に無い数値は絶対に書かず「資料に記載なし」と述べる。" & vbLf & _
        "【確認マーク】抜粋から完全には裏付けられない記述の文末に (要確認) と付ける。" & _
        "裏付けのある記述には付けない。全文に付けるのは禁止(意味が消えるため)。" & vbLf & _
        "【実務での使いどころ】お客さま対応に関わる内容では、最後に1行だけ" & _
        "「お客さまへ案内する前に確認すべき点」を書く(無ければ書かない)。" & vbLf & _
        "【断定の禁止】例外規定・特約・経過措置の有無が抜粋から読み取れないときは、" & _
        "断定せず「この抜粋の範囲では」と限定して述べる。"
End Function

' グラウンディング強制文(strict_grounding=TRUE時。設計書§D-1)。
Private Function GroundingInstruction() As String
    GroundingInstruction = "【厳守】本棚抜粋に書かれた情報のみで回答し、外部知識や推測での補完は禁止。" & _
        "各主張の直後に出典を必ず付け、出典を付けられない主張は書かない。" & _
        "抜粋から判断できない場合は、無理に答えず「資料からは判断できません」とだけ述べ、" & _
        "どんな資料を追加すれば答えられるかを1行添えること。"
End Function

' 構造化出力指示(answer_tags=TRUE時。設計書§D-1)。<thinking>は利用者非表示。
Private Function AnswerTagsInstruction() As String
    AnswerTagsInstruction = "出力は次の構造にすること: まず<thinking>タグ内に、どの抜粋が根拠か・" & _
        "矛盾が無いかの検討を書く(利用者には表示されない)。次に<answer>タグ内に、利用者へ見せる" & _
        "最終回答のみを書く。[[FOLLOWUP:...]]の行は</answer>を閉じた後(タグの外)に置くこと。"
End Function

' 深掘り候補の要求指示(裁定D11)。[[FOLLOWUP:...]]マーカーはmodAsk側で
' パース・除去して「深掘り候補」ブロックに整形される(利用者には見せない)。
Private Function FollowupInstruction() As String
    FollowupInstruction = "最後に、この回答をさらに深掘りするための質問候補を2つ考え、" & _
        "回答本文の一番最後に次の1行だけを追加してください" & _
        "(この行は利用者向け表示からは自動的に取り除かれます。良い候補が無ければ [[FOLLOWUP: なし]] と書く):" & vbLf & _
        "[[FOLLOWUP: 候補1 | 候補2]]"
End Function

' hits()由来の情報から「## 本棚抜粋」ブロックを組み立てる。
' maxCharsを超える手前で打ち切り、打ち切った場合のみ末尾に「(一部省略)」を挿入する。
Private Function BuildSourceBlock(hits() As Hit, ByVal nHits As Long, ByVal maxChars As Long) As String
    Dim lim As Long
    lim = maxChars
    If lim < 0 Then lim = 0

    If nHits < 1 Then
        BuildSourceBlock = "(本棚抜粋なし)"
        Exit Function
    End If

    Dim sb As String
    Dim used As Long: used = 0
    Dim truncated As Boolean: truncated = False
    Dim i As Long
    For i = 1 To nHits
        Dim entry As String
        entry = SourceTag(hits(i)) & vbLf & SourceBody(hits(i)) & vbLf & vbLf

        If used + Len(entry) > lim Then
            Dim remain As Long
            remain = lim - used
            If remain > 0 Then
                ' 2026-08-01(R12-1-6): 生の Left$ ではサロゲートペアの真ん中で
                ' 切れて孤立サロゲートがプロンプトへ混入する(R11-I で SafeLeft へ
                ' 寄せた際の適用漏れ箇所)。
                sb = sb & modUtil.SafeLeft(entry, remain)
                used = used + remain
            End If
            truncated = True
            Exit For
        End If

        sb = sb & entry
        used = used + Len(entry)
    Next i

    If truncated Then
        sb = sb & vbLf & "(一部省略)" & vbLf
    End If

    BuildSourceBlock = sb
End Function

' 本棚抜粋ブロックの本文: full_text(チャンク本文全体)を根拠として使う(Wave3 PM
' 裁定1)。full_textが空(旧データ・テストダブル等)の場合のみpreviewへ退避する
' (空の抜粋を本文に混ぜないための防御)。
Private Function SourceBody(ByRef h As Hit) As String
    If LenB(h.full_text) > 0 Then
        SourceBody = h.full_text
    Else
        SourceBody = h.preview
    End If
End Function

' origin="pack:<作成者>"なら [パック(作成者):ファイル名]、それ以外(self)は
' [本棚:ファイル名 p.N] の出典タグ文字列を返す(§4/§7.3)。
'
' 2026-08-03(R14-8a): Publicにした。入念モードの出典突合(modAskThorough.
' CiteIndexFrom)が「回答に書かれたタグが検索結果に在るか」を機械的に照合するとき、
' LLMへ指示している形と検査に使う形が別々の実装だと、一致するはずのものが全件
' 不一致になるか、その逆になる。タグの形はここが唯一の持ち主(憲章§4-5)。
Public Function SourceTag(ByRef h As Hit) As String
    If LCase$(Left$(h.origin, 5)) = "pack:" Then
        Dim authorName As String
        authorName = Mid$(h.origin, 6)
        SourceTag = "[パック(" & authorName & "):" & h.source & "]"
    Else
        SourceTag = "[本棚:" & h.source & " p." & CStr(h.page) & "]"
    End If
End Function

' UserContextBlock - 質問者の属性をプロンプトへ注入し、回答の粒度・トーンを
'   LLM側で調整させる。参照キーは実在するものだけを使う(config pack_author /
'   user_department、統計は modStats.AskTotalAll = 全モードの質問回数の合算)。
'   何も設定されていなければ空文字を返し、プロンプトに一切影響させない。
Private Function UserContextBlock() As String
    Dim nm As String, dept As String
    Dim askCount As Long, streak As Long, lvl As Long
    On Error Resume Next
    nm = Trim$(modConfig.GetString("pack_author", ""))
    dept = Trim$(modConfig.GetString("user_department", ""))
    ' R14-1a: quick+deep 決め打ちをやめ、全モードの合算(modStats)へ委譲する。
    askCount = modStats.AskTotalAll()
    streak = modStats.GetStat("streak_days")
    lvl = modStats.Level()
    On Error GoTo 0

    If nm = "名称未設定" Then nm = ""
    If LenB(nm) = 0 And LenB(dept) = 0 And askCount = 0 Then Exit Function

    Dim sb As String
    sb = "## 質問者の背景(回答の粒度・専門用語の量を調整するために使う。" & _
         "回答本文でこの情報自体には言及しないこと)" & vbLf
    If LenB(dept) > 0 Then sb = sb & "・所属: " & dept & vbLf
    If askCount > 0 Then
        sb = sb & "・このツールの利用回数: " & askCount & "回"
        If askCount <= 5 Then
            sb = sb & "(不慣れ。前提から補って丁寧に)"
        ElseIf askCount >= 50 Then
            sb = sb & "(習熟。前置きを省いて要点から)"
        End If
        sb = sb & vbLf
    End If
    If streak >= 3 Then sb = sb & "・連続利用: " & streak & "日" & vbLf
    UserContextBlock = sb
End Function
