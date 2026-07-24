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
'     ActiveSheet の各トークンはこのモジュールのソースに一切書かない。
'     ただし modConfig.GetString/GetLong の呼び出しはR4対象外として明示
'     許可されている(MASTER_SPEC発注時の指示: 「ExcelトークンがmodPrompts
'     自身のソースに現れなければR4適合」)。modConfig自体はシートを読むが、
'     それはmodConfig.basの中の話であり、modPromptsのソースには
'     Excelオブジェクトトークンが一切現れないためLintのトークン検査は通る。
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
'   ・Wave3 PM裁定1: modTypes.Hit に full_text フィールドが追加され、
'     modRetrieve.Searchがmy_knowledge読取時にfull_textを格納するように
'     なったため、本モジュールは「本棚抜粋」の本文として full_text を使う
'     (旧実装はpreview=先頭120字しか渡せず、回答生成の根拠が不十分だった)。
'     preview は出典先出し表示(modUIMain.RenderSourcesPreview)専用として
'     Hit型に残っており、本モジュールのプロンプト本文には使わない。ただし
'     full_textが空(旧データ・テストダブル等で未設定)の場合はpreviewへ
'     フォールバックする防御的実装にし、空の抜粋が本文に混じらないようにする。
'   ・深掘り候補(裁定D11): 利用者に実際に表示される回答を生成する
'     BuildQuickPrompt(すぐ聞くの最終応答)とBuildDeepVerifyPrompt(しっかり
'     調べるの最終応答)の末尾に、[[FOLLOWUP: 候補1 | 候補2]] 形式で深掘り
'     質問候補を2つ付けさせる指示行(FollowupInstruction)を追加する。
'     マーカーはmodAsk側でパース・除去され「深掘り候補(『続けて質問』で
'     そのまま聞けます)」ブロックに整形されるため、利用者の目に生マーカーが
'     触れることはない(LLMが指示を無視してマーカーを出さなくても、候補
'     ブロックが付かないだけで壊れない)。BuildDeepDraftPromptは表示されない
'     中間生成物(検証段の入力)のため対象外。
' ============================================================================

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
    If strictGrounding Then sb = sb & GroundingInstruction() & vbLf
    If answerTags Then sb = sb & AnswerTagsInstruction() & vbLf
    sb = sb & vbLf
    Dim uctx As String: uctx = UserContextBlock(): If LenB(uctx) > 0 Then sb = sb & vbLf & uctx
    sb = sb & "## 本棚抜粋" & vbLf & ctx & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf
    sb = sb & vbLf & FollowupInstruction() & vbLf
    BuildQuickPrompt = sb
End Function

Public Function BuildDeepDraftPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long, ByVal history As String, _
                                     Optional ByVal strictGrounding As Boolean = False, _
                                     Optional ByVal answerTags As Boolean = False) As String
    Dim lang As String
    lang = SafeAnswerLanguage()
    Dim maxChars As Long
    maxChars = SafeMaxContextChars()
    Dim ctx As String
    ctx = BuildSourceBlock(hits, nHits, maxChars)

    Dim sb As String
    sb = "あなたは社内の資料検索AIアシスタントです。「しっかり調べる」モードの下書き回答を作成します。" & _
         "以下の本棚抜粋を根拠に、質問へ" & lang & "で丁寧に、根拠を示しながら回答してください。" & vbLf
    sb = sb & CitationInstruction() & vbLf
    sb = sb & NotFoundInstruction() & vbLf
    If strictGrounding Then sb = sb & GroundingInstruction() & vbLf
    If answerTags Then sb = sb & AnswerTagsInstruction() & vbLf
    If LenB(history) > 0 Then
        sb = sb & vbLf & "## これまでの会話(参考。続きの質問なら踏まえて回答する)" & vbLf & history
    End If
    Dim uctx2 As String: uctx2 = UserContextBlock(): If LenB(uctx2) > 0 Then sb = sb & vbLf & uctx2
    sb = sb & vbLf & "## 本棚抜粋" & vbLf & ctx & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf
    sb = sb & vbLf & "(この回答は下書きです。この後、別の検証ステップで事実確認されます。" & _
        "根拠が弱い部分は無理に断定せず、その旨を書いてください。)"
    BuildDeepDraftPrompt = sb
End Function

Public Function BuildDeepVerifyPrompt(ByVal q As String, ByVal draft As String, hits() As Hit, ByVal nHits As Long, _
                                      Optional ByVal strictGrounding As Boolean = False, _
                                      Optional ByVal answerTags As Boolean = False) As String
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
    sb = sb & NotFoundInstruction() & vbLf
    If strictGrounding Then sb = sb & GroundingInstruction() & vbLf
    If answerTags Then sb = sb & AnswerTagsInstruction() & vbLf
    sb = sb & vbLf
    sb = sb & "## 本棚抜粋" & vbLf & ctx & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf
    sb = sb & vbLf & "## 下書き回答" & vbLf & draft & vbLf
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
        entry = "[" & i & "] " & SourceTag(hits(i)) & " " & _
                Replace(modUtil.SafeLeft(SourceBody(hits(i)), 300), vbLf, " ") & vbLf
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

Public Function BuildEnrichPrompt(ByVal batchText As String) As String
    Dim lang As String
    lang = SafeAnswerLanguage()

    Dim sb As String
    sb = "以下は社内資料から抽出した複数のチャンクです。それぞれに要約(30字程度)と" & _
         "キーワード(2〜5個、カンマ区切り)を" & lang & "で付けてください。" & vbLf
    sb = sb & "出力は必ず次のJSON配列形式のみで返してください" & _
         "(前後に説明文やコードフェンス、余計な文字を付けない):" & vbLf
    sb = sb & "[{""i"":1,""summary"":""..."",""keywords"":""a,b""}," & _
         "{""i"":2,""summary"":""..."",""keywords"":""a,b""}]" & vbLf & vbLf
    sb = sb & "## チャンク一覧" & vbLf & batchText
    BuildEnrichPrompt = sb
End Function

' ----------------------------------------------------------------------------
' 内部ヘルパー(すべてPrivate: modPromptsの公開契約はBuild*4関数のみ)
' ----------------------------------------------------------------------------

' SafeAnswerLanguage/SafeMaxContextChars:
'   modConfig.GetString/GetLong への呼び出しを1行スコープのOn Errorで守る
'   (R5: 直後にOn Error GoTo 0)。modConfig自体はconfigシートが無くても
'   既定値にフォールバックする実装だが、modPromptsはR4のPure Logicモジュール
'   群(vba_lint.py PURE_LOGIC_MODULES / run_lo_tests.py PURE_ALLOWLIST)
'   の一員として、modConfigモジュールそのものが読み込まれていない実行環境
'   (LibreOffice純ロジックテストの一時ライブラリ等)でも動く必要がある。
'   その環境では "modConfig.GetString" の呼び出し自体が実行時エラー
'   (Err=420 Invalid object reference)になることを実測で確認したため、
'   この関数呼び出しをOn Errorで包み、失敗時は契約既定値にフォールバックする。
'   Excel実機(modConfigが常に存在する環境)では通常どおりconfigシートの値を
'   返す(挙動は変えない。エラー発生時のみフォールバックが働く)。
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
        "【結論先行】必ず最初の1〜2行で結論を言い切る。前置き・挨拶・言い訳から始めない。" & vbLf & _
        "【構成】結論 → 根拠や詳細(箇条書き) → 注意点・例外(あれば) の順。" & vbLf & _
        "【記法・厳守】Markdown記号(#、**、`、表)は一切使わない(この画面では装飾されず崩れて見える)。" & _
        "代わりに: 見出しは「■ 」で始める / 箇条書きは「・」 / 最重要語だけ【 】で囲む / " & _
        "ブロックの間は空行1つ。1ブロックは3行以内。" & vbLf & _
        "【長さ】全体をおおむね200〜400字(複雑な質問でも600字まで)。同じ内容の言い換え、" & _
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
                sb = sb & Left$(entry, remain)
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

' 本棚抜粋ブロックの本文: full_text(チャンク本文全体)を根拠として使う
' (Wave3 PM裁定1)。full_textが空(旧データ・テストダブル等)の場合のみ
' previewへフォールバックする(空の抜粋を本文に混ぜないための防御)。
Private Function SourceBody(ByRef h As Hit) As String
    If LenB(h.full_text) > 0 Then
        SourceBody = h.full_text
    Else
        SourceBody = h.preview
    End If
End Function

' origin="pack:<作成者>"なら [パック(作成者):ファイル名]、それ以外(self)は
' [本棚:ファイル名 p.N] の出典タグ文字列を返す(§4/§7.3)。
Private Function SourceTag(ByRef h As Hit) As String
    If LCase$(Left$(h.origin, 5)) = "pack:" Then
        Dim authorName As String
        authorName = Mid$(h.origin, 6)
        SourceTag = "[パック(" & authorName & "):" & h.source & "]"
    Else
        SourceTag = "[本棚:" & h.source & " p." & CStr(h.page) & "]"
    End If
End Function

' ----------------------------------------------------------------------------
' UserContextBlock - 質問者の属性情報をプロンプトに注入(回答精度向上)。
'   部署・役職・利用回数・連続利用日数から、LLMが回答の粒度・トーンを
'   自動調整できるようにする。modConfig/modStatsへの呼び出しはR4対象外
'   (modConfig自体はシートを読むが、modPromptsのソースにはExcelオブジェクト
'   トークンが一切現れないためLintのトークン検査は通る)。
' ----------------------------------------------------------------------------
Private Function UserContextBlock() As String
    Dim dept As String, userName As String
    Dim askCount As Long, streak As Long, lvl As Long
    On Error Resume Next
    dept = modConfig.GetString("user_dept", "")
    userName = modConfig.GetString("user_name", "")
    askCount = modStats.GetStat("question_total")
    streak = modStats.GetStat("streak_days")
    lvl = modStats.Level()
    On Error GoTo 0

    If LenB(userName) = 0 And LenB(dept) = 0 And askCount = 0 Then
        UserContextBlock = ""
        Exit Function
    End If

    Dim sb As String
    sb = "## 質問者コンテキスト(回答の粒度・トーン調整に使う。本人には言及しない)" & vbLf
    If LenB(userName) > 0 Then sb = sb & "・氏名: " & userName & vbLf
    If LenB(dept) > 0 Then sb = sb & "・部署: " & dept & vbLf
    If lvl > 0 Then sb = sb & "・アプリLv: " & lvl & vbLf
    If askCount > 0 Then
        sb = sb & "・累計質問数: " & askCount & "回"
        If askCount <= 5 Then
            sb = sb & "(初心者=基礎から丁寧に)"
        ElseIf askCount >= 50 Then
            sb = sb & "(ヘビーユーザー=簡潔に本質だけ)"
        End If
        sb = sb & vbLf
    End If
    If streak >= 3 Then sb = sb & "・連続利用: " & streak & "日(定着ユーザー)" & vbLf
    UserContextBlock = sb
End Function
