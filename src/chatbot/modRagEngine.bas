Attribute VB_Name = "modRagEngine"
Option Explicit

' ============================================================================
' modRagEngine - retrieve-then-answer pipeline
' ----------------------------------------------------------------------------
' Steps:
'   1. embed(question) -> query vector
'   2. L2-normalize query
'   3. TopK over the in-memory index
'   4. assemble prompt with citation markers
'   5. chat(systemPrompt, userPrompt)
'   6. return answer + citations + token usage
' ============================================================================

Public Type Citation
    Source As String
    Page As Long
    Score As Double
    Snippet As String
End Type

Public Type AnswerResult
    OK As Boolean
    Answer As String
    Citations() As Citation
    PromptTokens As Long
    CompletionTokens As Long
    EmbedLatencyMs As Long
    ChatLatencyMs As Long
    ErrorMessage As String
End Type

Public Function Answer(ByRef idx As LoadedIndex, ByVal question As String) As AnswerResult
    Dim r As AnswerResult
    If Not idx.OK Then
        r.OK = False
        r.ErrorMessage = "Index not loaded"
        Answer = r
        Exit Function
    End If

    ' 1. embed query
    Dim qArr(0 To 0) As String
    qArr(0) = question
    Dim emb As EmbedResult
    emb = modApiGateway.Embed(qArr)
    r.EmbedLatencyMs = emb.LatencyMs
    If Not emb.OK Then
        r.OK = False
        r.ErrorMessage = "Embed failed: " & emb.ErrorMessage
        Answer = r
        Exit Function
    End If
    If emb.Dim_ <> idx.Dim_ Then
        r.OK = False
        r.ErrorMessage = "Embedding dim mismatch (query=" & emb.Dim_ & ", index=" & idx.Dim_ & ")"
        Answer = r
        Exit Function
    End If

    ' 2. L2 normalize query
    Dim q() As Double
    ReDim q(0 To emb.Dim_ - 1)
    Dim i As Long, sumsq As Double
    For i = 0 To emb.Dim_ - 1
        q(i) = emb.Vectors(i)
        sumsq = sumsq + q(i) * q(i)
    Next i
    If sumsq > 0 Then
        Dim inv As Double: inv = 1# / Sqr(sumsq)
        For i = 0 To emb.Dim_ - 1
            q(i) = q(i) * inv
        Next i
    End If

    ' 3. TopK
    Dim k As Long: k = modConfig.GetLong("retrieval", "top_k", 5)
    Dim topIdx() As Long, topScore() As Double
    modSimilarity.TopK idx.Flat, idx.Dim_, idx.Count_, q, k, topIdx, topScore

    ' 4. assemble prompt
    Dim contextStr As String
    Dim maxCtx As Long: maxCtx = modConfig.GetLong("retrieval", "max_context_chars", 6000)
    Dim used As Long: used = 0
    ReDim r.Citations(0 To k - 1)
    Dim cited As Long: cited = 0
    For i = 0 To k - 1
        If topIdx(i) < 0 Then Exit For
        Dim ch As Chunk: ch = idx.Chunks(topIdx(i))
        Dim block As String
        block = "[#" & (i + 1) & " " & ch.Source
        If ch.page > 0 Then block = block & " p." & ch.page
        block = block & "]" & vbCrLf & ch.Text & vbCrLf & vbCrLf
        If used + Len(block) > maxCtx Then Exit For
        contextStr = contextStr & block
        used = used + Len(block)
        r.Citations(cited).Source = ch.Source
        r.Citations(cited).page = ch.page
        r.Citations(cited).Score = topScore(i)
        r.Citations(cited).Snippet = Left$(ch.Text, 200)
        cited = cited + 1
    Next i
    If cited > 0 Then
        ReDim Preserve r.Citations(0 To cited - 1)
    Else
        ReDim r.Citations(-1 To -1)
    End If

    ' 5. chat
    Dim sys As String, usr As String
    sys = SystemPrompt()
    usr = "## ナレッジ" & vbCrLf & contextStr & vbCrLf & _
          "## 質問" & vbCrLf & question

    Dim chatRes As ChatResult
    chatRes = modApiGateway.Chat(sys, usr)
    r.ChatLatencyMs = chatRes.LatencyMs
    If Not chatRes.OK Then
        r.OK = False
        r.ErrorMessage = "Chat failed: " & chatRes.ErrorMessage
        Answer = r
        Exit Function
    End If

    r.Answer = chatRes.Content
    r.PromptTokens = chatRes.PromptTokens
    r.CompletionTokens = chatRes.CompletionTokens
    r.OK = True
    Answer = r
End Function

Private Function SystemPrompt() As String
    SystemPrompt = _
        "あなたは社内アンダーライターの業務支援AIです。" & vbCrLf & _
        "次のルールを厳守してください:" & vbCrLf & _
        "1. 回答は提示されたナレッジのみを根拠とし、推測を述べる場合は『参考意見』と明記する。" & vbCrLf & _
        "2. 個人情報(契約番号・氏名・電話番号など)が質問に含まれる場合は『個人情報を含めないでください』と返答し、回答しない。" & vbCrLf & _
        "3. 回答末尾に出典マーカー([#1], [#2] など)を必ず記載する。" & vbCrLf & _
        "4. ナレッジに無い情報は『社内ナレッジに該当する記載がありません。アンダーライターへ確認してください』と返す。"
End Function
