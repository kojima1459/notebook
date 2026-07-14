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

Public Function BuildQuickPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long) As String
    Dim lang As String
    lang = SafeAnswerLanguage()
    Dim maxChars As Long
    maxChars = SafeMaxContextChars()
    Dim ctx As String
    ctx = BuildSourceBlock(hits, nHits, maxChars)

    Dim sb As String
    sb = "あなたは社内の資料検索AIアシスタントです。「すぐ聞く」モードとして、" & _
         "以下の本棚抜粋だけを根拠に、質問へ" & lang & "で簡潔に回答してください。" & vbLf
    sb = sb & CitationInstruction() & vbLf
    sb = sb & NotFoundInstruction() & vbLf & vbLf
    sb = sb & "## 本棚抜粋" & vbLf & ctx & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf
    sb = sb & vbLf & FollowupInstruction() & vbLf
    BuildQuickPrompt = sb
End Function

Public Function BuildDeepDraftPrompt(ByVal q As String, hits() As Hit, ByVal nHits As Long, ByVal history As String) As String
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
    If LenB(history) > 0 Then
        sb = sb & vbLf & "## これまでの会話(参考。続きの質問なら踏まえて回答する)" & vbLf & history
    End If
    sb = sb & vbLf & "## 本棚抜粋" & vbLf & ctx & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf
    sb = sb & vbLf & "(この回答は下書きです。この後、別の検証ステップで事実確認されます。" & _
        "根拠が弱い部分は無理に断定せず、その旨を書いてください。)"
    BuildDeepDraftPrompt = sb
End Function

Public Function BuildDeepVerifyPrompt(ByVal q As String, ByVal draft As String, hits() As Hit, ByVal nHits As Long) As String
    Dim lang As String
    lang = SafeAnswerLanguage()
    Dim maxChars As Long
    maxChars = SafeMaxContextChars()
    Dim ctx As String
    ctx = BuildSourceBlock(hits, nHits, maxChars)

    Dim sb As String
    sb = "あなたは社内の資料検索AIアシスタントの検証担当です。下書き回答を本棚抜粋と照合し、" & _
         lang & "で最終回答を作成してください。本棚抜粋で裏付けられない断定や事実と異なる記載は、" & _
         "修正するか削除してください。" & vbLf
    sb = sb & CitationInstruction() & vbLf
    sb = sb & NotFoundInstruction() & vbLf & vbLf
    sb = sb & "## 本棚抜粋" & vbLf & ctx & vbLf
    sb = sb & "## 質問" & vbLf & q & vbLf
    sb = sb & vbLf & "## 下書き回答" & vbLf & draft & vbLf
    sb = sb & vbLf & "(指示: 検証済みの最終回答のみを出力してください。下書きとの差分説明や、" & _
        "検証過程の説明は不要です。)"
    sb = sb & vbLf & FollowupInstruction() & vbLf
    BuildDeepVerifyPrompt = sb
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

Private Function CitationInstruction() As String
    CitationInstruction = "回答の根拠として使った箇所には、必ず出典を付けてください。" & vbLf & _
        "本棚に自分で入れた資料は [本棚:ファイル名 p.ページ番号] の形式、" & vbLf & _
        "他の人から受け取ったパック由来の資料は [パック(作成者名):ファイル名] の形式で示してください。"
End Function

Private Function NotFoundInstruction() As String
    NotFoundInstruction = "本棚抜粋に書かれていないことは、推測で埋めずに" & _
        "「資料には見当たらない」とはっきり述べてください。" & _
        "やむを得ず推測で補う場合は、それが推測であることを明示してください。"
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
