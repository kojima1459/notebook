# 公式OCRツール調査 (2026-07-31)

## 1. ツール概要

### Excel画像RAG (gazou_macros.txt)

PDFファイルを JPEG 画像に変換し、その画像に対してOCR処理を実行するツール。取得したテキストをベクトル化（embeddings）してExcelに格納し、ユーザーの質問に対して画像内容に基づいた回答を生成AIから取得する。見開きページの自動分割機能、テキスト抽出後の類似度計算による画像検索、複数画像を基に質問に答えるRAG（Retrieval Augmented Generation）機能を備える。

### Excel帳票OCR (chouhyou_macros.txt)

PDFまたはフォルダ内の画像ファイルから帳票項目を自動認識し、指定した項目名に該当する値を抽出するツール。項目の自動検出モード、項目名を事前指定して値を抽出するモード、生成AIで表をHTML形式に変換するモードの3パターンをサポート。抽出結果は新しいExcelブックに表形式で出力される。

---

## 2. Ghostscript呼び出しの完全引用

### gazou版 (790行付近)

```vba
Function ConvertPDFToJPEG(pdfPath As String, folderPath As String) As String
    Dim gsPath As String
    Dim outputPath As String
    Dim shellCommand As String

    ' Ghostscriptの実行ファイルのパス
    gsPath = ThisWorkbook.path & "\Ghostscript\gswin32c.exe"

    outputPath = folderPath & "\page_%03d.jpg" ' %03dでページ番号を3桁で出力
    ' Ghostscriptのコマンドを構築
    shellCommand = Chr(34) & gsPath & Chr(34) & " -sDEVICE=jpeg -r300 -dNOPAUSE -dBATCH -dSAFER " & _
                   "-dTextAlphaBits=4 -dGraphicsAlphaBits=4 " & _
                   "-sOutputFile=" & Chr(34) & outputPath & Chr(34) & " " & Chr(34) & pdfPath & Chr(34)
    ' シェルコマンドを実行
    Dim wsh As Object, result As Long
    Set wsh = CreateObject("WScript.Shell")
    result = wsh.Run(shellCommand, 0, True)

    If result = 0 Then ConvertPDFToJPEG = folderPath

    Set wsh = Nothing
End Function
```

**特徴**: `Chr(34) & gsPath & Chr(34)` で gsPath を引用符で囲む（正規形式）、解像度は `-r300`

### chouhyou版 (372行付近)

```vba
Function ConvertPDFToJPEG(pdfPath As String, folderPath As String) As String
    Dim gsPath As String
    Dim outputPath As String
    Dim shellCommand As String

    ' Ghostscriptの実行ファイルのパス
    gsPath = ThisWorkbook.path & "\Ghostscript\gswin32c.exe"

    outputPath = folderPath & "\page_%03d.jpg" ' %03dでページ番号を3桁で出力
    ' Ghostscriptのコマンドを構築
    shellCommand = gsPath & " -sDEVICE=jpeg -r150 -dNOPAUSE -dBATCH -dSAFER " & _
                   "-dTextAlphaBits=4 -dGraphicsAlphaBits=4 " & _
                   "-sOutputFile=" & Chr(34) & outputPath & Chr(34) & " " & Chr(34) & pdfPath & Chr(34)
    ' シェルコマンドを実行
    Dim wsh As Object, result As Long
    Set wsh = CreateObject("WScript.Shell")
    result = wsh.Run(shellCommand, 0, True)

    If result = 0 Then ConvertPDFToJPEG = folderPath

    Set wsh = Nothing
End Function
```

**特徴**: `gsPath` に引用符がない（バグの要因）、解像度は `-r150`

---

## 3. ChatGPTV呼び出しとプロンプト文言

### gazou版 GetOCR関数 (463-473行)

```vba
Function GetOCR(imagePath As String) As String
    Dim Prompt As String
    Prompt = "次の画像をOCRしてください。画像やイラストがあった場合は、その内容をテキストにしてください。回答は結果だけを回答、「OCRの結果は以下の通りです」等の回答文言は不要です"

    Dim strImage As String
    'OCRテキスト
    strImage = Application.Run("Base64FromFile", imagePath)
    strImage = Application.Run("EscapeJSON", strImage)
    GetOCR = Application.Run("ChatGPTV", Prompt, strImage, , , ToolName & "（取込）")

End Function
```

**プロンプト**:
> 次の画像をOCRしてください。画像やイラストがあった場合は、その内容をテキストにしてください。回答は結果だけを回答、「OCRの結果は以下の通りです」等の回答文言は不要です

### chouhyou版 GetOCR関数 (347-360行)

```vba
Function GetOCR(Items As String, imagePath As String) As String
    
    Dim Prompt As String
    Prompt = "画像から、次の「;;;」で区切った項目名に該当する値を抽出してください。" & _
            "抽出した値は「;;;」で区切って回答して下さい。" & _
            "項目名に該当する値が見つからない場合は、ブランクを回答、そのまま「;;;」で区切ってください" & _
            "項目名の()は配列ではないので単なる文字として扱ってください。" & _
            "最後に、項目数と、値の数が一致しているか確認、不一致の場合はやり直して、数が一致してから回答してください" & _
            "回答に「抽出しました」などの文言は不要です。値のみを「;;;」区切りで回答してください。＜以下項目名＞"
    Dim strImage As String
    strImage = Application.Run("Base64FromFile", imagePath)
    GetOCR = Application.Run("ChatGPTV", Prompt & Items, strImage, , , ToolName)

End Function
```

**プロンプト**:
> 画像から、次の「;;;」で区切った項目名に該当する値を抽出してください。抽出した値は「;;;」で区切って回答して下さい。項目名に該当する値が見つからない場合は、ブランクを回答、そのまま「;;;」で区切ってください項目名の()は配列ではないので単なる文字として扱ってください。最後に、項目数と、値の数が一致しているか確認、不一致の場合はやり直して、数が一致してから回答してください回答に「抽出しました」などの文言は不要です。値のみを「;;;」区切りで回答してください。＜以下項目名＞

---

## 4. 同梱 Ghostscript の構成

配布ZIPファイルに同梱される Ghostscript バイナリの構成は以下の通り：

| ファイル | サイズ | 説明 |
|--------|--------|------|
| `gswin32c.exe` | 87 KB | Ghostscript コマンドラインインターフェース |
| `gsdll32.dll` | 22 MB | Ghostscript 動的ライブラリ |
| `license.txt` | - | AGPLv3 ライセンス |

**配置規約**: `ThisWorkbook.path & "\Ghostscript\gswin32c.exe"`

ツール起動時に、このパスの存在チェックが実行される（ThisWorkbook.cls の Workbook_Open プロシージャで確認）。

---

## 5. 注意点：chouhyou版のバグ

chouhyou版の ConvertPDFToJPEG 関数には **公式ツール側のバグ** が存在する。

### 問題の詳細

- **gazou版**: gsPath を引用符で囲む実装 `Chr(34) & gsPath & Chr(34)` により、パスに空白が含まれる場合でも正常に動作する
- **chouhyou版**: gsPath に引用符がなく、そのまま shellCommand に連結されている

### 実装の比較

**gazou版（正規）**:
```vba
shellCommand = Chr(34) & gsPath & Chr(34) & " -sDEVICE=jpeg -r300 ...
```

**chouhyou版（バグ）**:
```vba
shellCommand = gsPath & " -sDEVICE=jpeg -r150 ...
```

### パスに空白がある場合の影響

`ThisWorkbook.path` に空白が含まれる場合（例：`C:\Users\User Name\Document\Tool\`）、chouhyou版では Ghostscript コマンドが正しく実行されない。

例えば、以下のようなコマンドが生成される場合：
```
C:\Users\User Name\Document\Tool\Ghostscript\gswin32c.exe -sDEVICE=jpeg ...
```

シェルが最初の空白で分割してしまい、`C:\Users\User` を実行ファイルと解釈するため実行失敗となる。

### 推奨される対応

chouhyou版を使用する場合は、以下のいずれかの対応が必要：

1. **パス変更**：Excel ファイルを空白を含まないパス（例：`C:\Tool\` など）に配置する
2. **パッチ適用**：以下のように chouhyou版の ConvertPDFToJPEG を修正する

```vba
' 修正案：gazou版と同じ形式に統一
shellCommand = Chr(34) & gsPath & Chr(34) & " -sDEVICE=jpeg -r150 -dNOPAUSE -dBATCH -dSAFER " & _
               "-dTextAlphaBits=4 -dGraphicsAlphaBits=4 " & _
               "-sOutputFile=" & Chr(34) & outputPath & Chr(34) & " " & Chr(34) & pdfPath & Chr(34)
```
