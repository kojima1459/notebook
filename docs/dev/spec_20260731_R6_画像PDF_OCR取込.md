# 要件定義 R6: 画像PDFのOCR取込（Ghostscript連携）（実装: Opus / 起草: Fable 司令塔）

## 目的

スキャンPDF等の画像PDF(E0303)を、会社公式配布ツール同梱の Ghostscript で
ページ毎にJPEG化し、既存の ChatGPTV(Vision) でOCRして本棚へ取り込めるようにする。
optVision.bas に明記された唯一の穴「PDF→画像変換の手段はVBA単体に無い」を塞ぐ。

前提資料(必読):
- docs/dev/reference_公式OCRツール調査_20260731.md
  (公式ツールのGSコマンドライン。帳票OCR版のgsPath引用符漏れは踏襲しない)
- docs/dev/RIBBON_API_CONFIRMED.md §1 #5(ChatGPTV) #8(Base64FromFile)
- src/opt/optVision.bas / src/ingest/modShelf.bas:279-322 / src/core/modFeatures.bas

## 確定済みの設計判断（司令塔裁定。変更しない）

1. Ghostscript本体はリポジトリに同梱しない。実行ファイルの解決順:
   (1) config `ghostscript_path`(gswin32c.exe のフルパス。既定"")
   (2) `ThisWorkbook.Path & "\Ghostscript\gswin32c.exe"`(公式ツールと同じ規約)
   見つからなければ、案内文つきの失敗
   (「社内ポータルの『Excel帳票OCR』zipにある Ghostscript フォルダを
     このファイルの隣に置くか、config の ghostscript_path で場所を指定」)。
2. 1ページ=1回の ChatGPTV 呼び出し(既存 VISION_PROMPT を使用、resolution="high")。
   複数枚バッチ(最大10)はページ境界が崩れるため v1 では使わない(コメントで将来課題として明記)。
3. ページ上限 config `vision_pdf_max_pages`(既定20)。
   GSは `-dFirstPage=1 -dLastPage=(上限+1)` で cap+1 ページまで描画し、
   出力ファイル数が 上限+1 なら「切り詰めあり」と確定して最後の1枚を捨てる
   (総ページ数を知らずに truncated を正確に判定するため)。
   truncated は既存の PARTIAL_PAGES 系(status="partial")へ接続する。
4. dpi は config `vision_pdf_dpi`(既定150。公式帳票OCR版と同じ。300は倍重い)。
5. GS実行は**ブロッキング待ちにしない**(壊れたPDF1つでExcelが永久フリーズするため)。
   `WScript.Shell.Run "cmd.exe /c <GS…> & 完了フラグファイル作成", 0(非表示), False`
   で起動し、完了フラグ(またはタイムアウト config `vision_pdf_timeout_sec` 既定120)
   を DoEvents つきループで監視する。タイムアウト時はそのファイルを失敗として返し、
   一時フォルダの掃除は試みる(ロック中のKill失敗は無視)。プロセスのkillはしない。
6. すべてのパスは二重引用符で囲む(公式帳票OCR版の引用符漏れバグを踏襲しない)。
   Windowsのファイル名に " は使えないため、引用エスケープは考えなくてよい。
7. 一時出力は `Environ("TEMP") & "\nxocr_" & 一意名 & "\page_%03d.jpg"`。
   成功・失敗どちらでもJPEGとフォルダを削除する(R6規約準拠の別Sub後始末)。
8. レイヤー配置(R2: コアはopt名を書けない):
   - **modUtil**(基盤・純): ページ付きテキストの符号化/復号
     `JoinPagedText(pages() As ExtractedPage) As String` /
     `SplitPagedText(ByVal s As String, ByRef pages() As ExtractedPage, ByRef truncated As Boolean) As Boolean`。
     区切りは行単位マーカー(例: `@@NEXUS_PAGE:<実ページ番号>@@`、
     先頭行に `@@NEXUS_TRUNCATED@@` があれば truncated=True)。
     CONTRACT(closed)への追記を忘れない。
   - **optOcrCore**(src/opt・新規・純ロジック): GSコマンド文字列の組み立て
     `BuildGsCommand(gsExe, pdfPath, outPattern, dpi, lastPage) As String`、
     `BuildRunCommand(gsCommand, doneFlagPath) As String`(cmd.exe /c ラップ)、
     `PageJpgName(idx) As String`(page_001.jpg形式)、
     `RenderCapFor(maxPages) As Long`(=maxPages+1) 等。
     Worksheets/Range/Application等に一切触れない。
     PURE_LOGIC_MODULES / PURE_ALLOWLIST / CONTRACT / modules.json へ登録
     (src/ui/modChrome の登録が名前ベースの前例)。
   - **optVision**(opt・拡張): 新Public `ExtractPdfOcrPagedText(path) As String`。
     GS解決→一時フォルダ描画(非同期+タイムアウト)→JPEG列挙→ページ毎に
     Base64FromFile+ChatGPTV→modUtil.JoinPagedTextで1本の文字列に→掃除→返す。
     失敗は "#ERR:E0303:<具体的な理由と対処>"。進捗は modUIMain.SetStage
     (「🖼 OCR中… i/n ページ」)。CONTRACT(closed)更新。
   - **modShelfVision**(src/ingest・新規・コア): modShelf の vision フォール
     バックブロック(E0301の既存分+E0303の新設)を丸ごと引き受ける。
     `Public Function TryVisionFallback(ByVal path As String, ByVal errCode As String, _
        ByRef pages() As ExtractedPage, ByRef truncated As Boolean, _
        ByRef outNote As String) As Boolean`
     E0301: 従来どおり ExtractImagePdfText(単一ページ)。
     E0303: InvokeFeature("vision","ExtractPdfOcrPagedText", path) →
     modUtil.SplitPagedText で pages() へ。失敗時は outNote に理由を返す。
     modules.json / CONTRACT 登録。
   - **modShelf**(変更最小): 既存の vision ブロック(279-322相当)を
     modShelfVision.TryVisionFallback 呼び出しへ置換(**正味で文字数が減ること**。
     残1,344字なので増やさない)。truncated が返れば既存の pagesTruncated/
     partial 系へ接続。outNote が非空なら manifest の失敗メモに使う。
9. **modFeatures.InvokeFeature の互換修正**: 現在 "#ERR:…" を一律
   "#ERR:FEATURE_UNAVAILABLE" に潰しており、利用者向けの対処案内
   (GS未配置の説明等)が消える。元の "#ERR:…" 文字列をそのまま返すよう変更
   ("#ERR:" 接頭辞は維持されるため既存呼び出し元の判定はすべて互換)。
10. config 追加(build/build_mybookshelf.py の build_config_rows へ3行):
    ghostscript_path("", 説明) / vision_pdf_max_pages(20) /
    vision_pdf_dpi(150) / vision_pdf_timeout_sec(120)。※4行
11. 文言更新:
    - modLog.FriendlyMessage("E0303"): 「対応していません」→OCR取込の存在と
      有効化条件(feature_vision + Ghostscript配置)を案内。
    - modUIShelf.BuildMemo の image_pdf 文言、modApp の image_pdf バブル文言も
      同様に更新(スクショ誘導は残しつつ、GS配置で自動OCRできる旨を追加)。
    - docs/40番台に利用者向け手順 `docs/43_画像PDFのOCR取込設定.md` を新規作成
      (ポータルからのDL→Ghostscriptフォルダ配置→config、コスト注意:
       1ページ=1回のAI呼び出し、既定は先頭20ページ)。

## 純ロジックテスト(modTestsPure3 へ追記。容量超過なら modTestsPure4 新設+登録)

- BuildGsCommand のゴールデンテスト: 空白と日本語を含むパスで期待文字列と
  完全一致(引用符の位置まで)。-dSAFER -dNOPAUSE -dBATCH -dFirstPage/-dLastPage
  -sDEVICE=jpeg -r<dpi> -dTextAlphaBits=4 -dGraphicsAlphaBits=4 を含むこと。
- BuildRunCommand: cmd.exe /c ラップと完了フラグの合成のゴールデンテスト。
- JoinPagedText/SplitPagedText 往復: 0/1/N ページ、truncated あり/なし、
  本文にマーカー風の行が含まれる場合の挙動(壊れず、最悪でもページ割れのみ)。
- RenderCapFor / 切り詰め判定(n=cap+1 → truncated)の境界。

## 制約

- CP932外文字禁止 / Integer禁止 / 宣言は先頭 / R6ハンドラ規約遵守。
- opt層はコア参照が基盤層+modUIMain.SetStageのみ(lintが強制)。
- modShelf は文字数を増やさない(純減させる)。modTestsPure2 へは追記しない。
- 検証: vba_lint ERROR 0 / run_lo_tests 全PASS(226+新規) /
  build --prod/--dev 自己検証OK。コミットは司令塔が行う。

## 受入条件(報告書に含めること)

- 各設計判断1〜11の実装箇所(file:line)。
- GSコマンドの最終形サンプル(実際に組み立てた文字列を1つ)。
- modShelf の文字数 before/after。
- lint/テスト/ビルド結果。逸脱があれば理由。
