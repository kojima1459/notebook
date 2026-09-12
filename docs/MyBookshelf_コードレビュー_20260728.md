# MyBookshelf.xlsm コードレビュー報告（バグチェック＋改善提案）

- レビュー日: 2026-07-28
- 対象: `MyBookshelf.xlsm`（Nexus Agent / build `20260727-011626Z+8ad78dd`、VBA 標準モジュール64本＋インストーラ、約26,400行）
- 方法: 静的解析（机上検証）。実機（Windows + Excel + 社内AIリボン + 共有フォルダ）での動作確認は未実施のため、再現手順は「コード上そうなる」ことの確認であり、実機テスト前の修正判断材料として使ってください。
  - vbaProject.bin（コンパイル済み実体）と `vba_src` シート（注入用ソース）の両方を抽出し、実バイトレベルで照合
  - サブシステム別に6系統（起動基盤／RAG検索・AI呼び出し／取込・本棚／配信・共有フォルダ／UI層／横断整合性）の精査を並列実施し、重要所見はすべて該当コードを再読して裏取り
  - config シート実データ（91キー）・全シート実データ・ツール解説書（§11 既知の課題12件）・小島さんのテストチェックリストと突合
- 重大度の定義: **CRITICAL**=主要機能が設計どおり動かない／データ破壊。**HIGH**=データ喪失・誤動作・デモ品質を大きく損なう。**MEDIUM**=特定条件で誤動作・運用リスク。**LOW**=限定的な不具合・磨き込み。**IMPROVE**=動くが見直し推奨。

---

## 0. 要約

**総評**: 26,000行規模のVBAとしては異例に丁寧な作りです（エラーコード体系、一括Range読み、Dir列挙の2段化、リトライ付き共有I/O、再入ガード、テスト群）。設計契約（MASTER_SPEC）がコメントに残っており、大半のモジュールはその契約を守っています。一方で、**プロダクトの差別化の核である「チャンネル配信」に、部門切替・更新配信を机上で確実に失敗させる致命バグが1件**（C-1）あり、小島さんが最も不安視している STEP5／STEP6 はこのビルドでは通りません。また出典プレビュー（C-2）、起動シーケンスの二重実行（H-1/H-2）、絵文字・記号の「?」化（H-16）など、デモの第一印象を直撃する問題が複数あります。いずれも修正自体は小さく、8/3 提出前に潰せる規模です。

| 重大度 | 件数 | 代表例 |
|---|---|---|
| CRITICAL | 2 | チャンネル切替/更新の削除が空振り（C-1）、出典チップ全滅（C-2） |
| HIGH | 17 | APIキー平文露出、二重Boot・旧UI重畳、発行が本棚全量書出し、感謝EXP恒久不達、データ喪失系4件、文字化け ほか |
| MEDIUM | 26 | 聞き返し誤マージ、Word可視残存、Acrobat経路死亡、テレメトリ二重計上 ほか |
| LOW | 23 | 入力消失、Ctrl+Enterの挙動、境界条件系 |
| IMPROVE | 15 | ログローテ、モジュールサイズ逼迫、未配線機能の整理 ほか |

（重大度は静的解析に基づく相対評価です。実機での再現頻度・影響は環境に依存します）

**提出前に必ず直すべき Top 5**: ① C-1（origin タグ不一致。H-14とセットで）、② C-2（出典チップ。実質2文字の修正）、③ H-2＋H-1（起動二重実行と旧UI重畳。既知の「画面ずれ」の正体候補）、④ H-16（「?」文字化け。デモの見た目）、⑤ H-10（#N/A を含む Excel が取り込めない。実務データで高頻度）。**加えて配布前に H-17（Azureキーのローテーション）が必須**（コードを1行も直さなくても、いま配ればキー配布になる）。

---

## 1. CRITICAL（主要機能が設計どおり動かない）

### C-1 チャンネルの削除タグ不一致 — 部門切替・更新配信・巻き戻し配信がすべて空振りする

- 場所: `modChannel.PurgeChannelChunks`（333行）↔ `modPack.ImportChunksDedup`（644行）／`modChannel.SyncChannel`（207-210行）／`modChannel.SwitchTo`（447-472行）
- 事象: チャンク削除は `origin = "pack:" & チャンネル名`（例 `pack:商品部`）の完全一致で探すが、取込時に書かれる origin は `"pack:" & 作者名`（pack_meta の author = 発行者の config `pack_author`、例 `pack:noriaki`）。`SyncChannel` は `ImportPackFile(localPath, True)` を呼ぶだけでチャンネル名を渡さないため、**部門名≠作者名である限り削除対象は常に0件**。解説書自体が §5.3「pack:\<作者名\>」／§6.3「origin=pack:\<部門名\>を削除」と矛盾しており、その矛盾がそのまま実装されている（SyncChannel のコメントも「origin が pack:\<チャンネル名\> のチャンクを消す」と書いている）。
- 帰結（時系列の机上検証。3系統すべて確認済み）:
  1. **STEP5（切替）が失敗**: 商品部→人事部へ切替 → 商品部の全チャンクが残留したまま人事部が追加される。前部門の分野を聞くと商品部の出典で答え続け、切替のたびに使用量が増えて `chunk_limit`（20,000）に向かう。
  2. **一度離れた部門に再接続できない**: 人事部→商品部へ戻すと、SwitchTo が商品部のローカル版数を消しているため再同期に入るが、残留チャンクとハッシュ重複で全件スキップ → 取込0件 → SwitchTo が失敗と解釈して `active_channel` を空に戻す →「読み込めませんでした」。ヘッダは未接続なのに中身は残る、という説明不能な状態になる。
  3. **STEP6（更新配信）で新旧混在**: 再発行後の更新で旧版が消えず、変更チャンクだけ追加される。コード内コメントが「最悪の事故」と呼ぶ「AIが古い条文を根拠に答える」状態がまさに起きる。全チャンク不変の再発行では取込0件のため版数が記録されず、「更新があります」が消えない。STEP3 の巻き戻し配信も同様に購読者側へ反映されない。
- 修正案: `ImportPackFile` に Optional の originOverride 引数を追加し、`SyncChannel` から `"channel:" & chName` を渡す（手渡しパックは従来どおり `pack:作者名`）。`PurgeChannelChunks` のタグも `"channel:" & chName` に変更。感謝状の宛先解決（`modP2P.AuthorOfSource`）は `pack:` のみを対象とするため壊れない。**修正すると今まで空振りで隠れていた H-14（SwitchTo の破壊的順序）が顕在化するため、必ずセットで直すこと。** 既存利用者の本棚に残った `pack:` 残骸の掃除導線（全削除→再同期）も用意する。

### C-2 出典アクセサの添字ズレ — 出典チップ（Peek View）が全滅し、空名の「専門家に質問」ボタンが出る

- 場所: `modAsk.LastHitSource/LastHitPage/LastHitOrigin/LastHitPeek`（300-311行）。消費側 `modPeek.RenderCitations`（87行）・`ShowPeek`（148行）、`modMentor.FindExpert`（346-355行）
- 事象: 検索ヒット配列 `mLastHits` は常に `ReDim (1 To n)`（modRetrieve:251,393 / modAsk:528）で1始まりだが、アクセサ4本は「添字0始まり」としてガード `If i >= 0 And i < mLastNHits` の下で `mLastHits(i)` を直接参照する。i=0 で実行時エラー9（添字範囲外）。すぐ上の `LastTopSource`（283行）のコメント自身が「mLastHits は ReDim(1 To n)。0始まり走査は添字エラーになる（2026-07-20実バグ）」と明記しており、アクセサ側だけ直し漏れている。
- 帰結: (a) RAG回答のたびに `modPeek.RenderCitations` が i=0 で即エラー → `On Error GoTo Done` で脱出し、「出典(クリックで原文を確認):」の見出しだけ描かれて**出典チップが1枚も出ない**（Peek View 機能全体が死んでいる。「答えの下の資料名ボタンを押して原文確認」という使い方シートの中核導線が成立しない）。(b) `modMentor.OfferMentor` は冒頭 `On Error Resume Next` のため FindExpert 内のエラーで判定行ごとスキップされ、パック由来ヒットがある回答に**「この分野は さんが詳しいです」という空名ボタン**が出る（押しても無反応）。
- 修正案: アクセサ4本を `mLastHits(i + 1)` に変更（または呼び出し側3箇所を1始まりに統一）。`LastHitCount>0` のとき `LastHitSource(0)` を呼ぶスモークテストを modTestsExcel に追加。

---

## 2. HIGH

### 起動シーケンス（3件・相互連動）

**H-1 Boot が毎起動2回走る（Auto_Open ＋ インストーラの OnTime）／その OnTime は解除不能**
- 場所: `ThisWorkbook.Install`（37行 `Application.OnTime Now+1秒, "modBoot.Boot"`）、`modBoot.Auto_Open`（60行）、`modBoot.Boot`（gBootDone 設定は277行=最後）
- 2回目以降の起動では、Workbook_Open→Install が OnTime(+1秒) を予約した直後に、保存済みの `Auto_Open` が Boot を同期実行する。Boot は共有フォルダ同期＋起動ジッタ（DoEvents ループ、平均1.5秒）で1秒を超えるため、**ジッタ中の DoEvents で期限到来済みの OnTime Boot が再入発火**し得る（gBootDone はまだ False → フル初期化が二重実行。MsgBox 2連発・SyncNow 二重・`sync_interval_min>0` なら AutoSyncTick の予約が追跡不能になり「閉じたブックが勝手に再オープン」する）。Boot 完了後に発火した場合は H-2 の経路に入る。また Boot 予約の時刻はどこにも保持されないため `Auto_Close` でキャンセルできず、**発火前にブックを閉じると数秒後に勝手に開き直して保存までされる**（解説書§9 が「防ぐ」と明記している事故そのもの）。
- 修正案: Boot 冒頭にモジュール変数の再入ガード（`mBootRunning`）を追加。Install 側の予約時刻を ui_state に記録し、Boot 先着側と Auto_Close で `Schedule:=False` キャンセル。

**H-2 gBootDone 後の再実行パスが「旧ホームUI」を Hub の上に再描画する（既知の「画面ずれ」の正体候補）**
- 場所: `modBoot.Boot`（67-78行）
- 初回パスは `modUIMain.EnsureLayout` の後に `modHub.EnsureHubLayout` で上書きする規約だが、`If gBootDone Then` の再実行パスは Hub 化以前のまま取り残されており、旧ホームUI（btn_/lbl_＋セル装飾）だけを再構築して Exit する。H-1 の2回目 Boot がこのパスを踏むと、**Hub タイルの上に旧「質問する」画面が重なった二重画面**になり、起動約1秒後にホーム/マイ本棚が一瞬前面に出るフラッシュも発生する。小島さんメモの「タブが常時表示されない」「位置ずれたら画面再描画で直る」という既知症状の発生源として最有力（EnsureHubLayout を通ると自己修復されるため「再描画で直る」とも整合する）。
- 修正案: 再実行パスに `modHub.EnsureHubLayout` を1行追加（初回パスと同型化）。H-1 の再入ガードとセットで。

**H-3 Boot 中盤の `On Error GoTo 0` が Failed ハンドラを恒久解除 — 素のエラーダイアログ＋EnableEvents 焼き付き経路**
- 場所: `modBoot.Boot`（160行ほか計9箇所）
- VBAの `On Error GoTo 0` は「直前の状態に戻す」ではなく「当該プロシージャのハンドラを無効化する」。86行で張った `On Error GoTo Failed` が160行で解除され、以降ステージ4〜8入口の広い区間が無ハンドラで走る（182行・245行の `bootStage` 設定はハンドラ前提の死にコード）。この区間で実行時エラーが出ると（例: config セルに #REF! が混入 → `CStr(エラー値)` で型不一致、M-7参照）素の「実行時エラー」ダイアログで Boot が即死し、92行で False にした **`Application.EnableEvents` がセッション全体で焼き付く**（Excel の全ブックでイベント死）。
- 修正案: 各 `On Error Resume Next` ブロックの終端を `On Error GoTo 0` ではなく `On Error GoTo Failed` に戻す（機械的な置換で完了）。

### 配信・共有（5件）

**H-4 正典発行が本棚全量を書き出す — 他部門の正典・手渡しパックまで自部門の正典に混入**
- 場所: `modKnowledge.OnPublish`（769行 `ExportPackToFile(dest, "", True, wrote)`）→ `modPack.LoadChunksForExport`（420行、source名フィルタのみで origin 無視）
- 発行時に origin フィルタが無く、購読中の他部門正典（origin=pack:…）も含む my_knowledge 全行が pack.xlsx に入る。発行担当者が他部門を購読した状態で発行すると、**他部門の正典全文が自部門パックとして配布**され、以後元部門が改定してもコピーは古いまま残る。小島さんメモの「PC1台で発行者と利用者を両方テストする方法」はまさに発行端末で購読する手順なので、PoC 初日に踏む。
- 修正案: 発行経路では origin="self" の行のみ書き出す（LoadChunksForExport に originFilter 追加）。発行前の件数表示も同じフィルタ後の数に揃える。

**H-5 感謝EXPが構造的に誰にも届かない — 宛先は表示名、受取判定は AD ID**
- 場所: `modP2P.EmitThanks`（88-118行）／`CollectThanks`（144-161行）／`AuthorOfSource`（412-428行）
- 感謝状の宛先キーは origin 由来の**パック作者名**（初回起動で入力する表示名。config 実データでは "noriaki"）、受取側の照合キーは **AD の CN または %USERNAME%**。両者が偶然一致しない限り、感謝EXP・thanks_received_total・称号（modBoard.TitleFor の源泉）は一切付与されず、宛先不明の thx_ ファイルは誰も削除しないため共有フォルダに無限堆積する。「使うと作った人に感謝が届く」というプロダクトの中核ループが実質無効。
- 修正案: pack_meta に author_id（`modP2P.CurrentUserId()`）を追加して宛先はそちらを使う（表示は従来の author）。過渡期は受取側で両ハッシュを走査。作成後N日の宛先不明ファイルGCも追加。

**H-13 「更新があります」が別チャンネル起因で永久点灯し、表示部門名も誤る**
- 場所: `modChannel.PendingUpdates`（161-181行）／`modHub.DrawInbox`（486-506行）
- 判定は全チャンネル横断で、一度も同期していない部門はローカル版数が空のため常に「更新あり」。さらに SwitchTo が旧部門の版数を空に戻すため、切替直後から旧部門が保留扱いに戻る。Hub は保留の実体を見ずに**アクティブ部門名**で「【商品部】に更新があります」と表示 → 押しても「既に最新です」→ バッジは消えない無限ループ。elseif 連鎖のため解決済みQ&A等の他のお知らせカードも永久に隠される。
- 修正案: 判定を ActiveChannel のみに変更（または版数が空=未接続チャンネルを除外）。表示は保留リストの実際の先頭要素名を使う。

**H-14 SwitchTo の破壊的順序 — 新部門の取得可否を確かめる前に旧部門を消す（C-1修正後に顕在化）**
- 場所: `modChannel.SwitchTo`（455行 purge → 461行 SyncChannel）
- 版読取・pack存在確認・%TEMP%コピー（リトライ込み）はすべて SyncChannel 内なのに、旧部門の purge はその前。共有瞬断・発行者の上書き中ロック等でコピーが失敗すると、**旧部門は消え・新部門は入らず・active_channel 空**の三重苦で終わる。現ビルドでは C-1 の空振りに隠れて起きないが、C-1 を直すと朝の一斉更新で実際に踏む。
- 修正案: SyncChannel 内で「ローカルコピー確保後に purge → import」の順に組み替える（同居時間ゼロは import 直前 purge で維持できる）。失敗時メッセージに「前の部門の内容が残っています/外れました」の実状を明記。

**H-15 共有Q&A・困りごと回収の恒久取りこぼし — 読取失敗でも既読化＋累計60ファイルで走査停止**
- 場所: `modInsight.CollectFrom`（350-364行）／`MAX_COLLECT=60`（41行）／`ReadShared`（430-448行、リトライ無し）
- 読取が失敗（AVスキャン・書込直後ロック）しても nonce を既読化するため、その投稿はその端末に二度と入らない。また走査上限60は既読・自分発ファイルも数え、qa\・gap\ のファイルは誰も消さないため、**部内の累計投稿が60件を超えた日から新着が誰にも届かなくなる**（エラー表示なし）。共有知フライホイールが約60投稿で静かに停止する。
- 修正案: 既読化を読取成功パス内へ移動し、ReadShared に modP2P 同等の3回リトライを実装。guard は「未読を処理した件数」でカウントし、共有側の取込済みファイルGC（または日付別サブフォルダ化）を導入。

### データ保全（4件）

**H-6 再取込が「旧データ削除→抽出」の順 — 抽出失敗で既存ナレッジが消える**
- 場所: `modShelf.IngestFile`（139行 削除 → 146行 抽出）
- 同名資料の置換で先に旧チャンク/ベクトルを消してから抽出する。ユーザーがそのPDFを開いていてWordが開けない・一時的なネットワーク断などで抽出が失敗すると、旧データは戻らず**その資料の検索が即座に全滅**する。
- 修正案: 抽出・チャンク化成功後に旧行を削除して書き込む順序へ（重複排除用ハッシュ集合は当該sourceを除外して構築すれば後置できる）。

**H-7 フォルダ列挙の失敗が「空フォルダ」と区別されず、スコープ内全資料を削除**
- 場所: `modShelfSync.EnumFolderFiles`（529-531行）＋ SyncNow の消失判定（304-309行）
- 初回 `Dir$` を `On Error Resume Next` で握るため、アクセス拒否・瞬断・AVブロック時に「0件」となり、消失判定が同期スコープの全資料を `DeleteSource` する。フォルダ不存在は保護済みだが、存在確認後の列挙失敗には保護がない。
- 修正案: 初回 Dir$ のエラーを検知したら同期を中断（missing扱い）。「ディスク0件かつmanifest多数」の全削除には確認かスキップの安全弁を追加。

**H-8 失効/ドメイン外ワイプが my_manifest を残す — 「自動で戻せます」の案内どおりに復元されない**
- 場所: `modGuard.WipeKnowledge`（145-174行。my_knowledge/my_vectors/ch:版のみ削除）↔ `modShelfSync.DiffDecision`（keep判定）
- ワイプ後も manifest は status=done のまま残るため、次回同期は「変化なし=keep」でスキップし、**自分で取り込んだ資料は shelf_folder が健在でも二度と自動復元されない**（チャンネル分だけ ch: 版クリアで復元される）。EnforceExpiry のダイアログ「接続して開き直すと自動的に取り込み直せます」と実装が矛盾。本棚UIはカードを表示し続けるのに検索0件という不整合になる。§11-10（発火条件バグ）とは独立の、正当な失効時にも起きる欠陥。
- 修正案: WipeKnowledge で manifest のデータ行も削除（または全行 status="failed" 化して replace 経路に乗せる）。

**H-9 スクショ取込（jpg）が翌朝の同期で「消失」と誤判定され削除される**
- 場所: `modUIShelf.OnIngestScreenshot`（342-352行、保存先=shelf_folder）↔ `modShelfSync.IsSupportedExtLocal`（551-556行、jpg対象外）＋消失判定
- スクショは shelf_folder に jpg として保存・取込されるが、同期のディスク列挙は対応拡張子のみ、manifest スコープ判定は拡張子無フィルタ。結果 jpg は「manifest に有りディスクに無い＝消失」として**次回同期で無言削除**される（ファイル自体は残るが再取込もされない）。STEP5 の「自分で入れた資料が消えていないか」に直結。
- 修正案: 消失判定側にも同じ拡張子フィルタを適用（1行）。またはスクショ保存先を shelf_folder 外へ。

### 取込・表示品質（4件）

**H-10 エラー値セル（#N/A 等）1個で Excel ファイル全体が取込失敗**
- 場所: `modExtractorExcel.ExtractSheetText`（107-109行 `CStr(v)` に IsError ガード無し）
- セル値が #N/A/#REF! 等のとき Variant は vbError で、`CStr()` が型不一致(13)を投げ、抽出済みシート分も破棄して E0302 で全体失敗する。VLOOKUP の #N/A を含む一覧表は実務で頻出のため、**Excel 取込の実用性を大きく下げる**。エラー詳細からも原因が読めない。
- 修正案: `If IsError(v) Then`（スキップ or 固定文字列化）を IsEmpty 判定の前に追加（実質2行）。

**H-11 取込ファイルのマクロが実行される／編集中ブックが巻き添えで閉じられる**
- 場所: `modExtractorExcel.Extract`（38-63行）、`modExtractorWord`（AutomationSecurity 未設定）
- xlsm/xls/doc を開く際に `AutomationSecurity=ForceDisable`・`EnableEvents=False` を設定していないため、**取り込むファイルの Auto_Open/Workbook_Open マクロが MyBookshelf と同一セッションで実行される**（出所不明ファイルの任意コード実行経路）。また対象が同一Excelで既に開かれていると %TEMP% コピーが失敗して原本を掴み、未保存変更の破棄ダイアログ→**ユーザーが編集中のブックを `wb.Close False` で閉じてしまう**。
- 修正案: Open 前に AutomationSecurity=3 と EnableEvents=False（終了時復元）。既オープンのブックはスキップ（partial扱い）。コピー失敗（=ロック中）時は原本フォールバックせず失敗させる。

**H-12 聞き返し（曖昧確認）への返信に数字が1つでも含まれると、新しい質問が黙って破棄される**
- 場所: `modClarify.MergeAnswer`（111-146行）／`PickSource`（158行 `InStr(1, reply, num)`）
- 聞き返し保留中の次送信は必ず MergeAnswer を通り、返答のどこかに半角 1〜4 が含まれるだけで「番号選択」とみなされ、**利用者が実際に打った文章は捨てられて「前回の質問＋対象の資料: 候補N」が送信される**。保険実務の質問は「第1条」「3日以内」等で数字をほぼ確実に含む。保留はチャットクリアまでブック再起動を跨いで残るため、翌日の無関係な質問が乗っ取られるシナリオが現実的。
- 修正案: 番号選択とみなす条件を「返答全体が短い番号パターンのみ」（例: 6字以内・数字/丸数字/区切りのみ）に限定し、それ以外は書き直しとして新文を優先。保留に有効期限（保存時刻）を付ける。

**H-16 CP932 に無い文字が注入時に「?」化 — チャット回答・ボタン・ダイアログに「?」が出る（20箇所）**
- 原因: 本ブックは開くたび `vba_src` シートのソースを VBE へ注入するが、VBE はコードを CP932 で保持するため、CP932 に無い文字（〜 U+301C、— U+2014、🗑、ế/ệ）はリテラル `?`（0x3F）として保存される。**vbaProject.bin の実バイトで「?」を確認済み**（olevba の表示問題ではない）。ほとんどの絵文字は ChrW 組み立てで無事であり、規約（modUINexusDraw 冒頭「非ASCIIはChrWで組む」）から漏れた箇所だけが化けている。
- 場所: 下表（コメント内は除く、実行時文字列のみ）。

| 箇所 | 内容 | ユーザーへの見え方 |
|---|---|---|
| modAsk:349,352 | 信頼度表示「部分的に一致 — …」 | **全RAG回答**に「?」が出る |
| modUIMain:692,694 | モードボタン「(1〜2分)」「(10〜20秒)」 | 常時見えるボタンに「?」 |
| modUIShelf:152 | 「🗑削除」ラベル | 「??削除」 |
| modApp:629,630 | 言語「Tiếng Việt」 | ボタン表示・config 保存値が「Ti?ng Vi?t」化（LLMへの言語指定も破損） |
| modApp:764,767 / modPrompts:203,255,260 | プロンプト内「1〜2行」「200〜400字」等 | LLM への指示文に「?」 |
| modKnowledge:523,611,650 / modShared:276 | MsgBox「1〜2分かかります」 | ダイアログに「?」 |
| modUIDashboard:251,253 | バッジ行「— 獲得しました」 | ダッシュボードに「?」 |
| modEmbed:151 / modGateway:405 | 進捗表示・エラー文言 | 「?」 |

- 修正案: vba_src 側で一括置換 — 〜→～(U+FF5E) ／ —→－ or ─（─ U+2500 は生存実績あり）／ 🗑→`ChrW(&HD83D) & ChrW(&HDDD1)` ／ Tiếng Việt→ChrW 組み立て定数（既に化けて保存された config `answer_language` の移行として、旧 "Ti?ng Vi?t" も Case に暫定併記）。コメント内の化けは動作無害だが、可読性のため同時置換を推奨。

### セキュリティ（1件）

**H-17 Azure APIキーが実質平文で3経路に露出 — 配布ブロッカー（解説書§12.3 B1 の記載より深刻）**
- 場所: `modTestsPure`（437-439行）／config `azure_embed_key`／`modUtil`（51行 `OBF_KEY` 平文定数）
- 事象: 解説書B1は「難読化テスト固定値」と表現しているが、実体は二重の露出。(a) modTestsPure のテストコードに、config の `azure_embed_key` と**バイト単位で同一**の OBF1 文字列と、その**復号後の平文キーそのもの**（32桁hex）が期待値として直書きされている——難読化を解く必要すらなく、モジュールを開けばキーが読める。(b) 難読化は XOR＋16進の可逆変換で、鍵文字列 `OBF_KEY` も modUtil.bas:51 に平文で同梱。復号を実際に再現し、テスト内の平文と一致することを検証済み（本報告書には実キー値は転記しません）。`azure_embed_url`（社内Azureリソース名・デプロイ名）も config に平文で載っており、組み合わせると即時利用可能。vba_src シート（veryHidden だがパスワード保護なし）にも同じ文字列が入るため、VBA プロジェクト保護の有無に関係なく露出する。
- 帰結: **コードを1行も直さなくても、現状のまま配布した時点で全受領者に Azure キーが渡る。** 解説書 §11-2／§12.3 B1 が既に指摘する配布ブロッカーだが、想定より突破が容易（難読化解除すら不要）。
- 修正案（優先度順）: ① キーを即時ローテーション（露出済み前提で無効化）② テスト固定値を本番と無関係なダミー生成値へ差し替え ③ 一般配布は `embed_transport=ribbon` に倒す ④ 配布ビルドは `azure_embed_key`／`azure_embed_url` を空にする。あわせて個人メールアドレスの直書き（modHub:785／modHelp:309 の `m-kojima@…`）も config キー化を推奨。

---

## 3. MEDIUM（特定条件で誤動作・運用リスク）

**RAG・AI呼び出し**

- **M-1 利用上限誤検知**: `modGateway.LooksLikeLimitError`（420-426行）は120字以下の応答に「上限」「回数」等が含まれるだけで E0204 扱いにする。「請求回数の上限はありません。[出典]」のような正当な短文回答が誤って「利用上限」エラーに置換される。→ 出典タグ（`[本棚:` 等）や構造タグを含む応答を除外してから判定。
- **M-2 回答タグ抽出の乗っ取り／thinking露出**: `modRagParse.TagInner`（126-142行）は最初の `<answer>` 出現位置から切り出すため、資料本文に `<answer>…</answer>` という文字列が含まれモデルが thinking 内で引用すると、ニセ回答が表示される（資料由来のプロンプトインジェクション経路）。`<answer>` 欠落時のフォールバックは thinking を剥がさず応答全体を表示する。→ `</thinking>` より後の最初の `<answer>` を採用、フォールバックは thinking ブロック除去後を返す。
- **M-3 埋め込みの非アトミック確定**: `modEmbed.EmbedPending`（176-183行）はベクトル追記→embedded=1 の間に ESC/エラーで中断すると、次回再埋め込みで my_vectors に**同一 chunk_id の重複行**を作る（単段検索では同一チャンクが topK を2枠占有）。→ 追記前に既存行を Find して upsert。
- **M-4 偽Q&Aペアの部内共有**: `modAsk.FeedbackGreen`（373-375行）の共有ペアは「今回の質問＋**前回成功ターンの回答**」になり得る（回答本文は成功ターンでしか更新されないが質問は毎ターン更新されるため）。0件回答・聞き返しターンの直後に✅を押すと不整合ペアが「人が確認したQ&A」として配信される。→ 非成功ターンで回答バッファをクリアし、空なら発信しない。
- **M-5 ナレッジ地図の逐次セル読み**: `modCluster.LoadVectors`（132-139行）は my_knowledge 全行×3セルを1セルずつ読む（2万チャンクで約6万回のCOM呼び出し×ダッシュボード表示と分析CSVの2回）。大規模時にダッシュボードが数十秒フリーズ。→ 一括 Range 読みへ（modRetrieve と同型に）。

**起動・基盤**

- **M-6 ドメインチェックの脆弱性と非対称な帰結**: `modGuard.CheckDomain`（40-62行）は USERDOMAIN の NetBIOS 名しか実質見ず（FQDN設定・AzureAD参加機で正規端末が全滅）、全角カンマ・全角スペースを正規化しない。かつ不一致の帰結が失効タイマー（7日前予告）と違い**予告なし即ワイプ**（H-8 と複合で実質不可逆）。→ 正規化＋両変数照合＋「初回は表示ブロックのみ、ワイプはN回連続時」へ。
- **M-7 config のエラー値・桁あふれで既定値フォールバック契約が破れる**: `modConfig`（40-123行）は A/B列にエラー値が1つあると `CStr` で型不一致になり、それより下の全キーの取得が例外化する（H-3 の無ハンドラ区間と複合すると EnableEvents 焼き付き）。`GetLong` は `CLng("99999999999")` でオーバーフロー。config は非エンジニアが直接編集する設計なので現実的な入力。→ IsError ガードと On Error フォールバックを追加。
- **M-8 Install が注入失敗を握りつぶして Save**: `ThisWorkbook.Install` は Add/Remove/Name の失敗を無視して**無条件 Save** するため、モジュール欠落・二重注入などの破損状態がファイルに固定化される（次回以降「マクロを実行できません」）。→ 失敗モジュール名を収集し、1件でもあれば Save をスキップして再入手を案内。
- **M-9 modUiLock にタイムアウト自動解除が無い**: Leave 漏れが1箇所でもあると全ボタンが**無音・無期限**に死ぬ（modShelfSync の同種ガードは30分自動解除あり）。今回 Enter/Leave 対応の全数調査では漏れは見つからなかったが、将来の1バグで全滅する構造は危うい。→ Enter 時刻を保持し N 分で自動解除＋busy 時の軽い通知。

**取込・COM**

- **M-10 パック取込の無ハンドラ区間**: `modPack.ImportPackFile`（182-216行）は Open 後〜Close までハンドラが無く、pack 内のエラー値セルで素のVBAエラー→ **pack.xlsx が開きっぱなし**になる。Open/Close 中の ScreenUpdating 抑止も無くちらつく。→ 区間を On Error で覆い Failed で必ず Close。
- **M-11 パスワード付き文書でモーダル停止**: Word/Excel の Open にダミーパスワードを渡していないため、暗号化ファイルに遭遇すると入力ダイアログで同期が停止する（Wordはリトライで2度出る）。→ `PasswordDocument:="__dummy__"` 等でエラー化して E0302 扱いに。
- **M-12 Word が可視のまま起動する診断コードの残存**: `modExtractorWord.TryExtractOnce`（63-78行）に「一時的に Visible=True で診断」「原因確定次第 False へ戻す」というコメント付きの診断コードが本番ビルドに残っており、PDF/docx 取込のたびに Word ウィンドウが開閉する。ユーザーが触ると COM エラー・WINWORD 残留の誘因。→ Visible=False に戻す（1行）。
- **M-13 Acrobat フォールバックが常に失敗する**: `modExtractorAcrobat.ExtractPageWords`（113-123行）が使う `PDPage.GetNumWords/GetWord` は Acrobat IAC に存在しない API（正規手段は JSObject の getPageNumWords/getPageNthWord）。エラー438で必ず失敗し、「Word失敗→Acrobatで救済」の経路が Acrobat Pro 環境でも機能しない。→ GetJSObject 経由へ置換し実機確認。
- **M-14 Word ページ抽出の off-by-one**: `modExtractorWord.ExtractPageText`（142-148行）が `endRange.Start - 1` とするため、自然なページ送りの境界で**各ページ末尾の1文字が欠落**する（「100万円」の「円」が消える等。300ページ約款で最大299箇所）。→ `= endRange.Start` にして改ページ文字は文字列側で除去。
- **M-15 enrich 機能が未配線**: `modEnrich.EnrichPending` はどこからも呼ばれておらず（全モジュール grep で確認）、config `enrich_mode=light/full` は**完全に無効**。さらに配線した場合も行番号ズレ書込みのバグが内在（modEmbed 型の chunk_id 再解決が無い）。→ 配線するか config 説明を「未実装」に修正。
- **M-16 ギャラリー描画の O(資料数×全行) 逐次読み**: `modVault.PreviewOf`（667-681行）が資料ごとに my_knowledge を1セルずつ走査。資料100件×1万チャンクでナレッジ画面が分単位のフリーズ。→ 一括読み＋Dictionary 化。
- **M-17 パック取込の一括書込み**: `modPack.ImportChunksDedup`（677-686行）は数千行×9列（full_text 32,000字含む）を1回の Range 代入で書く（modShelf は同じ処理を err#7 対策で200行バッチ化済み）。大型正典の初回接続で実行時エラー7の懸念。またベクトル空のチャンクにも空行を書き、後の埋め込みで**同一 chunk_id の重複行**が恒久残留する。→ 200行バッチ化の共通化＋空ベクトル行のスキップ。

**配信・集計**

- **M-18 版番号が分精度**: `modPublish.FinalizePublish`（112行）の版が `yyyymmdd-hhnn`+作者のため、同一分内の「発行→巻き戻し→再発行」で版文字列が初回と一致し、購読者に配信されない。テストの連続発行で現実に踏む。→ 秒精度（hhnnss）へ。
- **M-19 テレメトリの二重計上**: 生涯累計のスナップショットを「ユーザー×月」ファイルで置くため、2ヶ月目以降は同一ユーザーの累計が月数分合算され、利用者数・質問総数・節約時間が過大表示。→ uid ごと最新月のみ集計 or 単一ファイル上書き。
- **M-20 共有不達時の起動/終了ブロッキング**: 不達判定を OS の SMB タイムアウトに委ねており、存在しないホストでは最初の UNC タッチが10〜30秒ブロックし得る。起動時（thanks/noise回収→insights→channels）と終了時（テレメトリ送信）が直列に走るため「閉じたのに数十秒残る」体感になる。→ セッション単位の到達フラグ（最初の probe 失敗で以後の共有I/Oをスキップ）。
- **M-21 チャンネル切替・巻き戻しが UI ロック外**: `modKnowledge.OnChannels/OnPublish` は InputBox 後に Leave してから DoSwitch/DoRollback を呼ぶため、リトライ待機中の DoEvents でクリックが素通りし、**SwitchTo が入れ子実行**され得る（purge/import の交錯で本棚が中途状態に）。→ 実処理区間で Enter/Leave を取り直し、purge+import を EnableCancelKey=xlErrorHandler で保護。
- **M-22 自分発ファイルの前方一致誤爆**: `modInsight.CollectFrom`（354行）の自分除外が `InStr(1, nc, myId) <> 1` の前方一致のため、ID "鈴木" は "鈴木一" さんの投稿も自分発と誤認して受信しない。→ `myId & "-"` で照合（またはpayloadの発信者IDと比較）。
- **M-23 起動ジッタが主要I/Oの後**: 既定（sync_on_open=True）では thanks/noise の全ファイル走査が StartupJitter より**先に**走り、一斉起動分散が主要負荷に効いていない。→ ジッタを SyncNow より前へ移動。

**UI**

- **M-24 vision 無効時にクリップボタンが無言で死ぬ**: `modApp.OnAttachImage`（482行）の `If VarType(hasImg) = vbString Or Not CBool(hasImg)` は、VBA の Or が短絡しないため機能無効時（InvokeFeature が "#ERR:…" 文字列を返す）に `CBool("#ERR:…")` で型不一致 → Fail が無言 Leave。ボタンは機能フラグ無関係に常時表示されるため「出るのに動かない」。→ modUIShelf.OnIngestScreenshot と同型の2段 If へ。
- **M-25 チャットバブルが選択・移動・Delete できる**: シート保護が `DrawingObjects:=False`（modUI:115）で、バブルの OnAction 結線は廃止済みのため、クリック→白ハンドル→Delete で回答が消える（Ctrl+Z は無効化済みで復元不能）。しかも操作案内文言が「吹き出しをクリックして選択」とこの操作へ誘導している。→ `DrawingObjects:=True` で保護（UserInterfaceOnly:=True でマクロ操作は可）＋文言修正。
- **M-26 ヘッダーのモード/速度/言語/テーマ切替が UI ロック非適用**: `modApp.OnToggleMode` ほか4ハンドラだけ Enter() を通らず、送信処理中でも状態が切り替わる。OnSend はモードを2回読む（158行と179行）ため、待機中に切り替えると**一般回答の下に前回RAGの出典チップが表示される**（出典の誤提示）。→ OnSend 冒頭でモードをローカル確定＋4トグルに IsBusy ガード。

---

## 4. LOW（限定的な不具合）

| ID   | 場所                                                    | 事象                                                                                                               | 修正の要点                                |
| ---- | ----------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- | ------------------------------------ |
| L-1  | modBitwiseOpt.BuildCache（186-211行）                    | 先頭行の次元でキャッシュ次元を確定するため、次元混在ストアでは残行が検索から消える（binary_rag=TRUE＋単段時のみ）                                                 | クエリ次元でプローブし、一致0件なら全件Floatへ辞退         |
| L-2  | modAsk.AppendFollowupPair（866-872行）                   | followup_max_pairs≤0 でメモリだけ消し ui_state を消さないため、無効化後も古い履歴が送られ続ける                                                  | 無効化分岐でも空を SaveState                  |
| L-3  | modAsk.AppendFollowupPair（883-897行）                   | VBAリセット後に新規質問が先行すると保存済み履歴を上書き消去（読む前に書く）                                                                          | 冒頭に CanFollowup と同じ遅延ロード             |
| L-4  | modGateway.DirectEmbedSlice（485-487行）                 | 空チャンクをダミー1字 " " で埋め込み成功として保存（ribbon/mock 経路は失敗扱いで非対称）                                                            | 空要素はリクエストから除外し結果は空のまま                |
| L-5  | modUtil.CsvToVector（190-207行）                         | Val() のため破損要素が黙って 0.0 になり検出不能                                                                                    | 形式検査を追加                              |
| L-6  | modDiag.RecreateDiagSheet（192-203行）                   | Delete 失敗時に既定名の孤児シートが可視で残る（modBoot.RemoveOrphanDefaultSheets が「原因未特定」と注記して掃除している現象の有力な発生源）                       | On Error で包み Name 失敗時はロールバック         |
| L-7  | modGuard.EnforceExpiry（104-139行）                      | expire_days≤7 で到達済みでも毎回警告／期限超過後は毎回ワイプ通知／時計前進ジャンプで無警告即ワイプ                                                         | d≥1 条件・wiped フラグ・「前回警告済み」条件          |
| L-8  | modBoot.Boot（104-110行）                                | ドメインブロック経路が gBootDone を立てず Exit → 二重 Boot 時に消去・案内が2回                                                             | 脱出前に gBootDone=True                  |
| L-9  | modShelfSync.SafeFileLen/SafeFileDateTime（679-696行）   | `On Error Resume Next` 文で Err がリセットされた後にログへ渡すため、詳細ゼロのログになる                                                       | 先に変数へ退避                              |
| L-10 | modShelf.ApplyCrumb（487行）                             | パンくず実名置換後の32,000字クランプで上限チャンクの末尾が欠け得る（legacy＋巨大設定時のみ）                                                             | room 計算に実ファイル名長を織り込む                 |
| L-11 | modShelf.IngestFile（118-124行）                         | silent 同期中でも E0504（同名衝突）はモーダルを出す。同期サマリは失敗も成功数に加算                                                                 | silent 化＋戻り値で集計                      |
| L-12 | modShelf.IngestFile（151-158行）                         | 画像PDF（E0303）の vision フォールバックは optVision が PDF を拒否するため必ず失敗する死に経路                                                  | 最初からスクショ案内を出す                        |
| L-13 | modKnowledge.OnPublish                                | 発行部門名の禁止文字（/ : 等）が無検証で、後段の SaveAs 失敗が「権限/パス」の誤誘導メッセージになる                                                         | 入口で検査し具体的に指摘                         |
| L-14 | modPublish.Rollback（176-184行）                         | pack 差替え後の version 書込3連続失敗で「戻せませんでした」と言うが pack は旧版に置換済み。FileCopy にリトライ無し                                         | FileCopy リトライ＋実状を伝える文言               |
| L-15 | modP2P.MarkNonce ほか                                   | my_stats に nonce 行が無限蓄積し、線形探索の FindKeyRow が経年劣化                                                                  | 別シート化＋期限GC＋Dictionary キャッシュ          |
| L-16 | modPublish.AppendLog（228-240行）／modInsight.EnsureSheet | publish_log.txt が read-modify-write で同時発行時に行消失。insight_inbox 再生成時に10列目 selected のヘッダ欠落                           | OS追記モード化／ヘッダに selected 追加            |
| L-17 | modUI.InitUI（54-108行）                                 | ws.Activate 失敗時に ActiveWindow 設定と FreezePanes が別シートへ適用される                                                        | Repaint と同じシート一致ガード                  |
| L-18 | modApp.OnClearChat                                    | クリアで nx_peek／Mentor ボタンが残存。ui_state の履歴も残るため再起動でクリア済み会話が復元される                                                    | HideCitations/ClearMentor 呼び出し＋履歴空保存 |
| L-19 | modBoot.Boot（274行）                                    | Ctrl+Enter はセル編集中は効かず（1回目=確定、2回目=送信）、テンキー Enter 未登録。ヒント文と実挙動がずれる                                                 | `^{ENTER}` 追加登録＋文言調整                 |
| L-20 | modUiLock.Enter（46行）ほか                                | 処理中表示（StatusBar）は InitUI が DisplayStatusBar=False にするため Nexus 中は見えない。SetStage の実況も同様                             | 画面内トースト等へ置換 or StatusBar を隠さない       |
| L-21 | modApp.OnActDrill／OnAttachImage の Fail                | エラー時にログも通知も出さない無言経路（OnSend の Fail は両方やる）                                                                         | LogError＋エラーバブルを追加                   |
| L-22 | modApp.OnSend（141-142行）                               | 送信直後に入力欄をクリアするため、失敗時に長文質問が失われる                                                                                   | 失敗経路で入力欄へ書き戻し                        |
| L-23 | modShelfSync.CancelAutoSync（420-426行）                 | 予約解除がモジュール変数 mScheduled/mNextRunTime 依存で、VBAリセットで消えると Auto_Close が解除できず予約が残る（§9の保証と食い違い。sync_interval_min>0 時のみ） | 予約時刻を ui_state に退避し、それを使って解除         |

---

## 5. IMPROVE（動くが見直し推奨）

- **I-1 ログのローテーション欠如**: err_log／usage_log は無限追記（modChatLog は100件ローテ済み）。Install が開くたび Save する設計と相まって、長期利用でファイル肥大→毎起動の保存遅延に直結。チャット履歴と同じ TrimToMaxRows 方式を推奨。
- **I-2 モジュールサイズが注入上限に逼迫**: vba_src は1モジュール=1セル格納（Excel セル上限 32,767字、設計上限 30,000字）。改行を CRLF 換算すると **modAsk 30,797／modShelf 30,780／modDash 30,748／modKnowledge 30,744／modHub 30,634／modPack 30,592／modUIMain 29,807／modApp 29,214 の8本が29,000字超**（各モジュール残り約60行）。解説書§11-3 が挙げる4本（modApp/modHub/modKnowledge/modDash）より実態は広く、modAsk・modShelf・modPack・modUIMain も同等以上に逼迫。**PoC 中の UI ポリッシュはまさにこれらを触る作業**で、上限を超えると vba_src セルへの書き戻し時に無言で切り詰められ、次回起動の再注入でソースが破損（コンパイルエラーで全機能停止）する。既存パターン（新規モジュール＋フック1行）の厳守と、ビルドスクリプトへの「30,000字超で失敗」ガード追加を推奨。§11-3 の対象も8本へ更新を。
- **I-3 VBProject 内の残骸 Module1**: `Sub Test/End Sub` のみのモジュールが vba_src 管理外で残存（Install は vba_src 記載分しか Remove しないため永久に残る）。動作無害だが不純物。
- **I-4 ブック内「使い方」シートが未配線機能を案内**: 「購読の操作は要りません。開くだけで自動的に届きます」「使っていない部門の購読を外せます(Hubから)」— 実装は手動の単一チャンネル切替のみ（購読の設定UI＝Subscribe/Unsubscribe は未配線、まとめ同期 SyncSubscribed も未配線）。応答時間の記載（すぐ聞く10〜20秒）もボタン表記（1〜2分）と不一致。ユーザーの期待値がズレたまま PoC に入ると問い合わせのもと。**補足（解説書§11-12の訂正）**: `unsubscribed_channels` は「いかなる効果も持たない」と§11-12にあるが、実際は `modChannel.IsSubscribed`→`PendingUpdates`（modHub:486 から Hub の更新通知に反映）で読まれており、**config を手で書けば当該部門を更新通知から除外する効果はある**（設定UIが無いだけ）。設定台帳で「死に設定」と誤記しないこと。
- **I-14 到達不能な Public プロシージャ群（旧UI併存の実害）**: 整合性の全数照合で、どこからも呼ばれない Public プロシージャ15件を検出。内訳: 旧マイ本棚シートの `modUIShelf.OnAddFiles/OnSyncNow/OnPickFolder/OnExportPack/OnImportPack/OnBackToChat`（描画側 AddButton が死んでおりボタンが生成されない）、`modVault.OnVaultExportPack/OnVaultImportPack/OnVaultPickFolder/OnVaultSyncNow/OnVaultAddFiles`（ツールバー未描画）、`modChannel.Subscribe/Unsubscribe/WriteVersion`、`modApp.OnActGood/OnNavVault/OnNavDash/OnSelectBubble/OnImportSharedQA_Legacy` 等。コンパイルは通るため潜伏し、I-2 の30,000字上限を圧迫する死荷重。§11-4（新旧UI併存）の具体的な実態として、配線するか削除するかを機能単位で判断を。
- **I-15 解説書の軽微な数値・記載の訂正**: ①config は**90キー**（§8/§5.1 の「91キー」は行数91＝ヘッダ込みの取り違え）②`modFeatures` のビルドリストに実体の無い `optTts` が1件残る（feature_tts=False で握り潰され無害）③エラーコード表 §9: `E0704` は欠番、`E0702` は検索時の次元不一致にも使用（意図的拡張）、`E0801` は画面組み立て以外の内部失敗5件でも使用 — いずれも文書追記が安全。
- **I-5 チャンネル系の重複実装**: `modChannel.WriteVersion/PrepareChannelDir` は未使用で modPublish 側と重複（版フォーマット変更時の片側修正漏れリスク）。削除または委譲を。
- **I-6 modBoard の全ビーコン直列読み**: 起動毎に board\ 全ファイルを読む（1件3リトライ）。数百人規模では起動遅延要因。件数上限か当日分フィルタを。あわせて感謝の連打による EXP ファーミング（同一人の✅連打で +10 ずつ）は閾値運用の検討を。
- **I-7 direct 埋め込みの 429/5xx リトライ無し・応答 index 未照合**: 非200は即リボンフォールバック（128件を1件ずつ叩き直すためレート制限をむしろ悪化させ得る）。Retry-After 待機→同一スライス再送を1回挟む。応答の `data[].index` 照合も契約上は正。
- **I-8 多段検索のベクトル再パース**: SearchExpanded はクエリごとに全ベクトルCSVを再パースする（deep で最大5回）。パースを1回に括り出すだけで多段検索は約クエリ数倍高速化（§11-7 の Prefilter 結線と独立に効く）。
- **I-9 容量上限キーの二重管理**: 強制は `shelf_max_chunks`、Hub の使用率表示は `chunk_limit`（2026-07-28 当時: 両方20000で偶然一致、コード既定値は10000で不一致。2026-09-12 現在の出荷は chunk_limit=20,000 / shelf_max_chunks=20,500）。片方に統一を。
- **I-10 一時ファイルの掃除漏れ**: スクショ取込後の %TEMP% jpg、消し残り mbtmp_*（Kill 失敗の蓄積で連番501以後コピー放棄）、Kill 済みパスを指す manifest 行。起動時スイープを推奨。
- **I-11 フォルダ同期中の per-file 全面再描画**: SyncNow ループ内で毎ファイル RenderShelf（manifest+knowledge 全読＋最大400カード描画）が走り、大量取込が O(N×総チャンク) 化。同期中は抑止し完了時に1回へ。
- **I-12 文言・導線の残骸**: modHelp が旧UI（👍/👎・サイドバー）を案内／ガチャの「もう一度引く」案内が旧導線／optDiffDoc（約款差分）は feature_diffdoc=TRUE なのに UI から到達不能（未結線。結線時は diff_report シートに脱出ボタンも必要）／Hub のアイコン下ラベルに OnAction 無し（誤クリックで図形選択）／EnsureHubLayout が Unprotect したまま再保護しない。
- **I-13 小粒の整理**: modConfig.mLoaded は書くだけの死に変数／WipeKnowledge の戻り値が Delete 失敗でも件数を返す（kill switch の空振りが無音）／HexDecodeBytes が奇数長16進の最終ニブルを黙って捨てる／modAnalytics.CsvField が LF 単独をガードしない／ScrollToBottom の行高換算が約2行ずれる。

---

## 6. 小島さんテストチェックリストとの突合（机上検証）

| STEP | 観点 | 机上検証の結果 |
|---|---|---|
| 0 | config 設定 | 手順書どおりで可。ただし publish_key を入れた端末は H-4（全量書出し）に注意。`mock_llm` 等のキー綴りは実装と全数一致を確認済み |
| 1 | 起動と基本動作 | ⚠️ H-1/H-2 により、起動約1秒後の画面フラッシュ・Hub 上への旧UI重畳が発生し得る（既知の「位置ずれ」の正体候補）。5ナビ・右上6アイコンの OnAction 結線は全数照合で正しい。ボタンはブック名非依存のため **MyBookshelf_利用者.xlsm へのリネーム運用でも壊れない**（手順書の懸念に対して安全確認済み）。ツアー3枚は同位置・同サイズをコードで確認 |
| 2 | 正典の発行 | 手順書記載の「クリップボード貼り付け2ステップ」は旧仕様（解説書§11-5どおり現ビルドは全自動1操作）。**手順書の更新が必要**。version.txt/publish_log の生成は机上で正しい。⚠️ H-4（他部門購読中の発行で混入）に注意 |
| 3 | 巻き戻し | 発行側の「新しい番号に進む」はコードで確認（正しい）。⚠️ M-18（同一分内の再発行は配信されない）、L-14（差替え後の版書込失敗で半端状態）。受信側への反映は C-1 により機能しない |
| 4 | 受け取り側 | 初回接続は動く見込み（未接続→接続は purge 対象が無いため C-1 の影響を受けない）。ヘッダ表示・使用量%表示の実装は確認済み |
| 5 | 切替 | ❌ **C-1 により失敗する**（前部門の内容が本棚から消えない／一度離れた部門へ戻れない）。「自分で入れた資料が消えていないか」は purge 機構としては温存が正しいが、別経路の H-9（スクショ）・H-6〜H-8（データ保全）に注意 |
| 6 | 更新配信 | ❌ **C-1 により新旧混在が起きる**（チェックリストが恐れている事象そのもの）。H-13 により「更新があります」の表示自体も誤動作 |
| 7 | 異常系 | 共有不達・channels 不在・権限なしで「落ちない」ことは全経路で確認（例外はすべて握られる）。ただし ⚠️ M-20（起動/終了が数十秒ブロックし得る）、H-7（列挙失敗時の全削除）、ESC 連打は取込系は EnableCancelKey 保護済みだが M-21（切替中は無保護）。送信ボタン連打の二重送信防止は Enter/Leave 全数調査で成立を確認 |

---

## 7. ツール解説書 §11（既知の課題12件）の検証結果

| §11 | 検証結果 |
|---|---|
| 1〜6 | 記載どおりであることをコード上で確認（§11-5 の「発行は現ビルドでは全自動」も確認。テスト手順書 STEP2 の2ステップ記述は旧仕様） |
| 7（binary_rag 未結線） | 確認。Prefilter 呼出しは単段 Search のみ。既定 retrieve_mode=multi では作動しない |
| 8（ボーナス上限+0.64） | 定数を確認（0.15+0.10+0.15+0.24）。記載どおり |
| 9（単段/多段のスコア差・検証段 DomainGuard 無し） | 確認。GramBonus は多段専用、deep 検証プロンプトに DomainGuardInstruction 注入なし |
| 10（TouchReach 発火条件） | 確認。呼出しは modBoot:232 の1箇所のみ・チャンネル存在が条件。**追加で H-8（ワイプ後に manifest が残り自動復元も塞がる）という未記載の欠陥を発見** — 修正は §11-10 とセットで |
| 11（バッジ4種非表示） | 確認。判定12種 vs 表示8種のハードコードを3箇所で確認 |
| 12（Subscribe 系未配線） | 一部訂正あり。Subscribe/Unsubscribe/SyncSubscribed の未配線は確認。ただし `unsubscribed_channels` は「効果なし」ではなく、config 手編集時は Hub 更新通知の抑止として作用する（I-4 参照）。また I-4 のとおりブック内「使い方」シートも自動購読前提の文言のため要修正 |

**整合性の全数照合で確認できた健全性（この規模のVBAとしては極めて良好）**: Application.Run（14箇所）・Shape.OnAction（42箇所・実マクロ70件）・OnTime/OnKey・Worksheets参照（145箇所）・モジュール間呼び出し（1,891箇所）のいずれも**参照切れ・綴り違い・全半角違い・引数個数不一致・Private越境呼び出しがゼロ**。config キー参照も90キーと全数一致でtypoなし。エラーコード22種に重複割当なし。Option Explicit は Module1（残骸）を除く全モジュールにあり。**唯一の残課題**は、AutoSyncTick の OnTime 予約解除がモジュール変数（VBAリセットで消える）に依存する点（解説書§9の保証と食い違うが、既定 sync_interval_min=0 では予約自体が作られないため限定的。L-23）。

---

## 8. 検証して問題なしを確認できた点（安心材料）

- **二重送信防止**: OnSend/HotSend は modUiLock を先頭で通過し、連打は棄却される。Enter/Leave の対応は全ハンドラで成立（全数調査）
- **ボタン結線**: 全 OnAction・OnKey・OnTime のマクロ名実在と綴りを照合し不整合なし。ブック名修飾が無いためファイル名変更に強い
- **発行/購読の順序**: 発行は pack→version の順、購読は version→pack＋%TEMP% コピーで、部分ファイル取込は実質防止されている（TOCTOU は次回同期で自己修復する方向に倒れる）
- **巻き戻し（発行側）**: version が必ず新番号に進む設計は正しい。アーカイブ命名・直前版選択・巻き戻しの取消も整合
- **ノイズ報告の閾値**: 「異なる2人」を本当に保証する（ファイル名が reporter ハッシュで決定的なため再投票は1票のまま）
- **ベクトル数学**: 保存・クエリ双方 L2 正規化済みの内積=コサイン、次元不一致の明示検査、ゼロベクトル・空配列ガードあり
- **JSON エスケープ**: 質問に引用符・改行・バックスラッシュ・サロゲートペアが含まれても direct 埋め込みのリクエストは壊れない
- **Dir の入れ子**: 全列挙が「収集→後処理」の2段方式で統一され、列挙子破壊なし
- **行削除の方向**: すべて下から／Union 一括／配列書き戻し方式で行ズレなし
- **連続利用日数・holidays・レベル計算**: 営業日ベースの判定（同日2回・金→月・連休・年跨ぎ）を机上検証し正しい。holidays は CDate 非依存でロケール安全
- **32,767字セル上限**: 全書込み経路が 32,000 字でクランプ済み
- **config キー**: 全モジュールが参照するキー名を config 実データ91キーと突合し、綴り違いなし
- **Randomize**: Rnd 使用4箇所すべて直前に Randomize あり（起動ジッタが全端末同一値になる問題は無い）
- **再入・状態復元**: mIngesting/mSyncRunning 等のガードは全出口で解除（30分自己回復付き）。EnableCancelKey・ScreenUpdating・Calculation・DisplayAlerts は主要経路で復元（例外は H-3 の区間のみ）

---

## 9. 修正順序の提案

**フェーズ1: 8/3 提出前（デモを壊すもの・修正が小さいもの）**
1. C-2 出典チップ（アクセサ4本の添字。実質 `i + 1` への修正4行）
2. H-2 旧UI重畳（`modHub.EnsureHubLayout` 1行追加）＋ H-1 再入ガード（数行）
3. H-16 文字化け（vba_src の一括置換。〜→～、—→－、🗑→ChrW、Tiếng Việt→ChrW 定数）
4. C-1 origin タグ（引数1本追加＋タグ変更）＋ H-14 順序入替（セット修正・要注意）
5. H-10 #N/A ガード（IsError 2行）／ M-12 Word Visible=False（1行）
6. H-4 発行の origin=self フィルタ（1引数追加）

**フェーズ2: PoC 開始前（データ保全・運用系）**
- H-6/H-7/H-8/H-9（データ喪失4件）、H-3（ハンドラ復元）、H-12（聞き返し誤爆）、H-13（更新バッジ）、H-15（Insight 取りこぼし）、H-5（感謝ID。pack_meta 拡張を伴うため設計判断込み）、H-11（AutomationSecurity）、M-7（config 防御）、M-18（版の秒精度）
- 手順書・使い方シートの現行仕様への更新（STEP2 の2ステップ記述削除、I-4）

**フェーズ3: PoC 期間中（品質・性能・磨き込み）**
- MEDIUM 残り（M-1〜M-26）、LOW 全件、IMPROVE（特に I-1 ログローテ、I-2 モジュールサイズ管理、I-8 検索高速化）

**修正時の注意**
- C-1 と H-14 は必ず同時に直す（現状は2つのバグが相互にマスクし合っている）
- 修正は必ず vba_src シート側に行う（VBE 直編集は次回起動の再注入で失われる）。I-2 のとおり6モジュールは残容量が少ないため、行を足す修正は新規モジュールへの切り出しを優先
- 解説書 §12.3 B3（更新配布時にユーザーデータの移行手段が無い）のとおり、PoC 開始後の修正配布はコストが跳ね上がる。フェーズ1・2 はできる限り配布前に

---

## 付録: 検査環境と制約

- 検査は macOS 上の静的解析で実施（olevba 0.60.2 / openpyxl 3.1.5）。Windows 実機・社内AIリボン・共有フォルダでの動作、Word/Acrobat COM の実挙動、画面描画は未検証
- vba_src シートと vbaProject.bin の照合により、両者はコメント・文字化け（H-16）・VBEの大文字小文字正規化を除いて一致していることを確認（ビルドの自己整合性は良好）
- 行番号は vba_src シート格納ソースの行位置（VBE 表示とほぼ一致）
