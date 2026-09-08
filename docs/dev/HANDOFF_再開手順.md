# 再開手順（セッション中断対策・最終更新: R40 着手時点）

中断したら、次のセッションはこのファイルから読むこと。
**docs/dev/00_プロダクト憲章.md が全裁定の判定基準(必読)。**
リポジトリ: `kojima1459/notebook`、ブランチ: `claude/internal-notebook-lm-chatbot-B6BE7`。

## 0R40. R40（実機報告5件・PDF 取込の AMSI 強制終了ほか・2026-09-08・**検問中**）

**起点**: ユーザーの実機報告（R39 build `e08bb6f`）。bat 起動（シェアポ／OneDrive）・Word／Excel／スクショ取込・本文ボタン・ページ数表示・パック／引き継ぎの出力と取込は**全部動作確認済み**。残り5件を全裁定 = `spec_20260908_R40_実機報告5件.md`（§0 裁定表・§1 F1 設計・§4 実機確認）。ユーザー指示「最優先で修正」「修正実装して」（層1 の F1 も指示済み・一本道）。
**F1（最優先）**: PDF 取込で「悪意のあるマクロが検出されました…Office を終了します」＝ Defender AMSI が【WMI `Win32_Process.Create` → `cmd.exe /s /c "(gs) 1>log 2>&1 & (if errorlevel…) >flag"`】をマクロ型マルウェアの手口と誤検知。→ `optGsProc` を全面書き換え: `Shell()` で gswin32c を直起動＋kernel32（OpenProcess/GetExitCodeProcess/TerminateProcess）。完了フラグは VBA が終了コードで書く（新 `SyncDoneFlag`・optGsTxt の2つの待ちループ先頭）。`BuildRunCommand` は cmd ラップを捨て `-sstdout=<log>` を実行ファイル直後に挿すだけ。cmd.exe／WMI／WScript.Shell／taskkill は GS 経路から消滅。`_前回` ファイルは起動 bat の正常な退避。**この環境に Windows／Defender／GS が無いので AMSI が黙るかは実機でしか確認できない**（§4-1。再発したら保護の履歴の脅威名を聞く）。
**F2** 出典タグ `[本棚:…]` を 8pt・primary 色に（`modLive.CiteTagSpans` 純関数＋`StyleAnswerParas` 末尾。文字列は不変）／**F3** Excel 各行の先頭に `[A6] ` 番地（`modExtractorExcel.CellAddressOf/RowPrefix`）＋表示「シートN」（`modLive.PageLabel`・modPeek/BuildSourceBlock）＋1回読み・広域プロンプトに番地を添える1行。**既取込の Excel は削除→再取込で番地が付く**／**F4** パック出力・取込の `Bump` 直後に `EvaluateBadges`／**F5** `onepass_max_chars` 300000→**450000**（≒30万トークン。遅ければ 300000 へ）。
**検問**: 着手 `4ebf758` → F1 `80630c4` → F2〜F5 `0092290`。lint 0。LO compile／pure・敵対的レビュー（Opus）→ 司令塔 Fix → final-gates → 配布。数字は完了時にここへ。
**容量（09-08 python len）**: optGsProc 6,827／optGsTxt 27,183／**optOcrCore 29,338（残 662）**／modLive 25,617／modExtractorExcel 16,995／modAskOnePass 20,775／modAskGlobal 21,088／modPack 25,414／modPackExport 16,996／Pure4 26,383／**Pure6 28,788（残 1,212）**／Pure43 12,448。

## 0R39. R39（外部受入テストの指摘対応・2026-09-07・**検問全緑・実機再試験待ち**）

**final-gates（09-07・司令塔）**: lint ERROR 0／LO compile 170/170／pure **PASS 3211 / FAIL 0 / SKIP 14**（下限 3211。ネガティブ確認: Pure44 A_i1 を壊して FAIL 1→復元）／`--dev`・`--prod --zip` 自己検証 OK（24シート・170本）／`bin_roundtrip --book` 6条件／`lo_xlsm --book` 3条件 OK。配布物 `dist/MyBookshelf_配布.zip`（build `20260907-074450Z+e08bb6f`・14,658,267B・xlsm 1,964,179B・図解ガイド R39 版表記・`mock_llm` FALSE・キー空・docs 6本＝05/45/00/41/**10/44**・README に「1台1系統」）。**ユーザーへ送付済み。**
**敵対的レビュー（Opus）→ Fix は司令塔自身**（spec §6）: BLOCKER 2（空本文の分類が到達不能／ファイルハンドル漏れ）・MAJOR 4（`%` の対ずれ→完了フラグを `if errorlevel` の3値に／進捗信号の単調性／`Dir$` 無限ループ／O(n²)）を全部採用。**版跨ぎの注意（R1）**: F002 でハッシュ基準が変わったので、R38 のブックで作ったパックを R39 に入れると重複排除が1回効かない＝正典を配り直すときは版を揃える。

**起点**: 髙橋さんが私物 Windows（ARM64・Excel 16.0・社内リボン無し）で OpenAI Astra に R38 配布物（build `cd61029`）の受入テストを実施（実Excel で純ロジック 3,277・E2E 16・取込/起動/保存/パック/移行）。報告 zip は `scratchpad/astra/r/受入テスト_20260907`（リポジトリ外・会社資料なし）。指摘6件＋懸念1件。**全裁定** = `spec_20260907_R39_受入指摘対応.md` §0（司令塔が全件を再現して採用）。ユーザー GO（09-07）。
**致命（F001）**: 文字 PDF のページ番号が全部 p.1（同梱 Ghostscript の txtwrite 出力に改ページ文字が無い。R10 以来）。→ GS にページ別ファイル `gstext_%04d.txt` で出させ、読み側（`optOcrCore.ReadPageFilesJoined`）が各ページ末尾に Chr$(12) を付けて連結。連結後の処理は不変。**この環境に GS が無いので実機でしか検証できない**（§4 再試験）。cmd.exe 1行の中の `%04d` と `%^ERRORLEVEL%` の対（変数展開）はコマンドライン文脈では未定義名が原文のまま残る規則に依拠＝**実機要確認**（完了フラグに数字が入るか）。
**その他**: F002 別名同内容の重複（`modShelf` 凍結解除1行: ハッシュを見出し前の本文で）／F003 Pure8 の固定長配列（LO は通り実Excel は Err 13＝§10 に追記）／F004 E2E の質問文／F005 手順書 41・45 の10箇所（Gemini 3.1 Pro が突合→司令塔照合）／F006 zip に 10・44 同梱／C001 README に「1台1系統」。
**検問（09-07）**: 検問1 `2610cd4`（司令塔分）／検問2 `efe2d42`（F001・Sonnet 実装＋司令塔検収）／Fix `e08bb6f`（レビュー対応・司令塔）→ final-gates 全緑。
**実機再試験（§4・髙橋さんの環境で可能）**: ①5ページ PDF で page 1〜5・上限2で partial ②画像表紙＋文字本文 ③途中空白ページ ④同内容別名で件数不変・E2E 16/16 ⑤Pure8 が実Excel で PASS ⑥zip に 10・44。
**体制メモ**: Gemini 3.1 Pro は文書突合10件を正確に出した（引用が1件だけ不正確→司令塔が原文で修正）。

## 0R38. R38（「入念に調べる」の1回読み＝章丸ごと・メガコンテキスト・2026-09-06・**検問全緑・実機受入待ち**）

**final-gates（09-06・司令塔）**: lint ERROR 0／LO compile 169/169／pure **PASS 3193 / FAIL 0 / SKIP 14**（下限 3155→3193。ネガティブ確認2回: Pure43 C_正常等分 10001→FAIL 1、Fix2 の「E_見出し記号■が残る」を反転→FAIL 1、いずれも下限割れの赤を確認して復元）／`--dev`・`--prod --zip` 自己検証 OK（**24シート・169本**）／`bin_roundtrip --book` 6条件 OK／`lo_xlsm --book` 3条件 OK。配布物 `dist/MyBookshelf_配布.zip`（build `20260905-170004Z+cd61029`・14,632,175B・xlsm 1,958,224B・`mock_llm` FALSE・`azure_embed_key`/`publish_key` 空・`thorough_onepass` on・`onepass_max_chars` 300000・`onepass_max_chapters` 8・`onepass_seed_hits` 6・`onepass_neighbor` 6・`thorough_onepass_effort` high・`max_context_chars` 60000 のまま）。図解ガイドの版表記は R38（dist 実測後に更新）。**ユーザーへ送付済み。**
**レビュー2周目（Opus・Fix 検証）→ Fix2 は司令塔自身**（ユーザー裁定「最終検問と手直しは Fable」）: ■ の消失（表示側が行頭 ■ を見出しにする）／F1 の追加分が予算切れで載らない→章の予算8割／フッターの印は Footer 冒頭で必ず消費。裁定表は spec §7-2。

**起点**: 外部レビュー（Gemini）の裏どり（`review_20260905_外部レビュー裏どり_R37.md`）→ 髙橋情報「同じリボン経由で 20 万トークンまで1回で渡せる（本人検証済）」→ ユーザー GO「A（config だけ）を挟まず一気に B」。**入念＝精度／しっかり＝中間／すぐ聞く＝速度**を守り、変えたのは入念の単発経路だけ。**戻し方は config `thorough_onepass=off`**。
**仕様と全裁定** = `spec_20260906_R38_入念1回読み.md`（§0 守るもの／§2 新モジュール `modAskOnePass`／§3 配線／§5 実機受入・§5-2 テスター手順／§7 レビュー1周目の3ストリーム裁定）。
**体制の実験（ユーザー指示「Gemini 3.8 flash も奴隷の如く」）**: Gemini 3.8 flash を API 直叩き（`scratchpad/gemini_worker.py`・キーは scratchpad の `.gemini_key` のみ・リポジトリに入れない）でワーカーに使った。**起草（modAskOnePass＋Pure43・31秒・司令塔検収で3点修正）は有効**。**レビュー（3ストリーム目）は要注意**: git もファイル計測も実行できないのに「HEAD 3e6ecb9」「18,348字」と**推定値を事実のように書く**。所見は根拠つきのものだけ採る（採用1件=ESC、不採用1件=配列参照共有の誤警報）。文書下書き（使い方・テスター手順）は典拠照合すれば使える。配線・Fix は Sonnet、壊す班・2周目は Opus。**09-06 追記（ユーザー裁定）**: 同じキーで `gemini-3.1-pro-preview` も呼べる（実測41秒・正答）→ **難易度1〜4＝Gemini 3.1 Pro（既定・軽いタスクは全部）／5〜7＝Sonnet／8〜9＝Opus／10と司令塔・PO＝Fable**。3.8 flash は速さ優先の雑務だけ。`delegate` スキル・CLAUDE.md §1・ワーカーの既定モデルを更新済み。
**経過**: 仕様（`8b622d3`）→ 波1 Gemini 起草＋Sonnet 配線（`614701d`）→ 波2 LO 注入・ESC・文書（`da069fe`）→ 1周目3ストリーム → Fix 波 F1〜F8（`e6fcd43`）→ 2周目（Fix 検証・Opus）→ Fix2 司令塔（`cd61029`）→ final-gates 全緑 → 配布。
**検問の数字（Fix 波時点）**: lint 0／LO compile OK／pure **PASS 3192 / FAIL 0 / SKIP 14**（下限 3155→3192。ネガティブ確認: Pure43 C_正常等分 を 10001 に壊して FAIL 1＋下限割れ→復元）／`--dev`・`--prod` 自己検証 OK。容量: `modAskOnePass` 約20,100／`modAskThorough` 17,271／`modLive` 23,046／`modAskGlobal` 20,927／`modXDoc` 14,543。
**実機受入（§5-2・未）**: ①入念1問で実況「🔎 入念(1回読み)」・usage_log `thorough_onepass`（`mode=chapter|neighbor`・`chars=`＝プロンプト全長・`extra=`）・`ask_steps` の `thorough_onepass=<ms>`・フッターの段数が −3 ②`chars=` 20万超で `#ERR` にならない ③on/off で3問（表・但し書き・章またぎ。**1文で聞く**・usage_log の最新行が `thorough_onepass` であることを確認してから採点）④章に割れていない本棚で `mode=neighbor`（`chars=` が小さいのは正常・最大約100件）⑤続けて質問。**著しく遅い／精度低下 → `thorough_onepass=off`**。
**記録のみ／R39 候補**: 章ごとの `CollectChapterHits` 呼び出し×章数のシート走査コスト（A-m5・実機秒数を見てから走査1回化）／`allHits` 複製と突合 O(n²)（A-m7）／`modXDoc` は CONTRACT 外（記録1）／しっかり（deep）への1回読み適用は実機結果後／短い章の余り予算を長い章へ回す／`UserContextBlock` は1回読みに載らない。

## 0R37. R37（是正の仕上げ＝誤根拠の降格・資料間リンク（LLM 0回の軽量グラフ）・測定の材料・2026-09-05・**検問全緑・実機受入待ち**）

**最新状態（09-05 夜）**: 波A（Sonnet: 誤根拠の降格＋`ask_sources`）／波B（Opus: 資料間リンク＝`doc_centroids`/`doc_links`・`modXDoc`/`modXDocStore`/`modXDocBuild`）→ レビュー1周目（班A/B・Opus）→ Fix 波（Opus）→ 2周目（班C・Opus）→ Fix2（Sonnet）**まで全消化**。全裁定は `spec_20260905_R37_是正の仕上げと資料間リンク.md` §8〜§12。司令塔のマイクロ修正3件（`takeN`/`taken` 大小同名の二重宣言／`InStrRev` 4引数形／`Optional Object = Nothing`＝いずれも LO compile のタイムアウトか pure の FAIL で検出。CLAUDE.md §10 に追記）。
**final-gates（09-05）**: lint ERROR 0／LO compile 167/167／pure **PASS 3,155 / FAIL 0 / SKIP 14**（下限 3,155。ネガティブ確認: Pure42 の期待値を壊して FAIL 1／下限割れ→復元）／`--dev`・`--prod --zip` OK（**24シート**・167本）／`bin_roundtrip` 6条件／`lo_xlsm` 3条件 OK。配布物 `dist/MyBookshelf_配布.zip`（build `20260905-115417Z+8105569`・14,615,053B・`mock_llm` FALSE・キー空・`xdoc_links` on・`xdoc_min_sim` 80・`xdoc_add` 2・`log_max_rows` 4000）。**ユーザーへ送付済み。**
**実機受入（spec §3-2・未）**: ①関連する2資料を取り込むと 📖本文 のメタ行に「関連する資料」が出る（無関係な2冊では出ない） ②A の言葉で聞いて B に答えがある質問で末尾に B が出ることがある（⚡すぐ聞くか入念の単発で。続けて質問では出ない） ③❌違う のあと同じ質問で誤答の根拠だった資料が後ろへ下がる（完全一致のときだけ） ④`usage_log` に `ask_sources`・`xdoc_built`（`cap=`/`loadfail=`）が残る ⑤取込時間の増分（章重心の総当たり。4,000章相当で Err 7 が出ないか）。**R36 の受入①〜⑥と R35 追補の実機1往復も未確認のまま。**
**記録のみ／R38 送り**（spec §10・§12）: `IndexInBox` と `AppendFromPool` のコスト（実機で測ってから）／`CosineCsv` が本番経路で未使用／`|` 区切りの脆さ／非正規化ベクトル混入の防御なし／パック・チャンネル由来の資料は links を持たない（復旧は ⚡仕上げ）→ 全量再構築ボタン／`bench_retrieval.py` を `ask_sources`＋feedback から採点できるよう拡張／`vba_lint` に大小同名 Dim 検査（`check_dim_case_collision`）／`modAskRetrieve` 残126＝次は分割が先／`modLog` のフォールバック 2000。
**09-05 深夜・外部レビュー（Gemini）の裏どり**: `review_20260905_外部レビュー裏どり_R37.md`。バグ3件のうち**再現1件**（📖本文の書き込み中に `ScreenUpdating` を止めていない → 層2で修正済 **R38 M1**・`modTextView` 15,827字）、机上再現1件（降格同士の順序＝記録のみ）、**再現せず1件**（「降格が逆効果」＝並べ替え範囲 ⊆ 残る枠なので残る集合は不変）。「未修正4件」は **3件が事実誤認**（700字切りは rerank 段だけ／履歴は `prevU/prevA` で渡している／出典突合は R34 B1 で quick/deep 済）、1件（入念4段）は正しい。**メガコンテキスト全面移行は不採用**（リボン経由の入力上限が未知・大きい1冊で30万トークン超・測る土台がまだ空）。R38 候補: 入念1パスを config スイッチで A/B／`max_context_chars` を config で1段ずつ上げて実測／引用クリック→本文ジャンプ。
**起点**: ユーザー依頼「①是正の完成 ②NotebookLM 越えの精度（無理はしない）。動的グラフナレッジは効果があるか」。**裁定** = `spec_20260905_R37_是正の仕上げと資料間リンク.md`（§0 結論: フル GraphRAG はやらない／①の残り＝誤根拠の降格（A）／資料同士の関連は章重心の類似度で LLM 0回のリンク表（B）／精度の測定材料を残す（C））。波A（Sonnet）・波B（Opus）を並行投入 → レビュー2周 → final-gates。

## 0R36. R36（是正メモ・スクショ本文表示・画像を📁から・版上げ時の自動引き継ぎ・2026-09-05・**検問全緑・実機受入待ち**）

**最新状態（09-05 夜）**: 髙橋さんフィードバックを起点に4件を実装。波1（Sonnet: §3 本文表示＋§4 画像📁）／波2（Opus: §2 是正メモ）／波4（Sonnet: §1-A 版上げ時の自動引き継ぎ）／波3（Haiku: 文書）→ 敵対的レビュー1周目（班A壊す／班B突合）→ Fix 波（Opus）→ 2周目（班C Fix検証／班D 波4。Opus 上限で Sonnet に切替）→ Fix2（Sonnet）**まで全消化**。全裁定は `spec_20260905_R36_是正と本文表示と置き場所.md` §9（9-1〜9-8）。
**final-gates（09-05・司令塔）**: lint ERROR 0／LO compile 163/163／pure **PASS 3,090 / FAIL 0 / SKIP 14**（`EXPECTED_PASS_MIN` 3009→3090。ネガティブ確認: Pure41 の期待値を1分ずらして FAIL 1／下限割れの赤を確認→復元）／`--dev`・`--prod --zip` 自己検証 OK（163本・CP932 置換 371 字）／`bin_roundtrip` 6条件 OK／`lo_xlsm` 3条件 OK。配布物 `dist/MyBookshelf_配布.zip`（build `20260905-095232Z+f27af62`・xlsm 1,894,719B・22シート・`mock_llm` FALSE・`azure_embed_key`/`publish_key` 空・`screenshot_show_text` TRUE・`correct_inject` on・`correct_key_min` 60・bat 5,333B に `_旧版` 複製行と `_前回` 退避が分離・README に `MyBookshelf_旧版.xlsm`・docs に 05_図解ガイド.html 同梱）。**ユーザーへ送付済み。**
**実機受入（spec §6-6・未）**: ①📁 追加で png を選べて取り込める ②📸 取込直後に 📖本文 が自動で開く（← 戻る で戻る） ③❌違う→正しい内容→同じ質問で是正メモが出典の先頭に出る（🤔・深掘り・一般アシスタントでは旧経路） ④前版を `MyBookshelf_旧版.xlsm` に改名→新版の初回に「前の版の本棚が見つかりました（N件）」→はい→閉じて開き直すと本棚が戻る（0件の前版では聞かれない） ⑤pre-R35 の前版でも `Workbook_Open`/`Auto_Open` が走らない ⑥text_view を開いたまま保存→再起動で残骸が無い。**加えて R35 追補の実機1往復（OneDrive の bat → D: → 閉じる → 書き戻し・`_前回`）は依然未確認**（bat と Ghostscript は髙橋さん経由の SharePoint zip でしか会社PCへ入らない）。
**記録のみ／R37 送り**（spec §8・§9）: 1-B 本格データブック分離（起動・終了の秒数実測後）／誤チャンクの減点／他人の correction 受信箱の取り込み／是正メモ一覧／読み取り結果の編集保存／フォルダ同期での画像取込／`WScript.Shell`×17・`new:{`×2・`ServerXMLHTTP`×2 の撤去と `--vba-mode installer` 削除（姉妹PJの「0件照合」に追随）／精度は `bench_retrieval.py` で数字を取ってから／`HitSourceList` の是正メモ除外は4呼び出し元に及ぶ（意図どおり）／新規導入期間は起動ごとに `_旧版` の存在確認1回。
**運用メモ**: Opus のセッション上限（09-05 09:10 UTC リセット）で2周目レビューが起動直後に落ちた→ Sonnet で再投入して完走。worktree のサブエージェントは基点が古い `main` になる事故が3回続いた（各自 `git fetch && reset --hard` で復旧。**次回から指示文に「最初に HEAD を確認」を必ず入れる**）。マージ時の末尾追記衝突は「両方残す」で解決するが、共通部分が消える型（modTestRunner の `End If`・modules.json の `},{`）があるので、**解決後は必ず lint＋`--dev` ビルド＋LO compile まで回してから次へ**。

## 0R35. R35（配布方式の転換: 自己インストーラ → 完成品 vbaProject.bin・2026-09-03・**検問全緑・実Excel確認待ち**）

**最新状態（09-03 夜）**: 波1〜3 → 敵対的レビュー1周目（班A壊す／班B残骸）→ Fix波 F1（Sheet1 ごみ）／F2a（Auto_Open 一本化・初回 Save・禁止語・コメント・班A採択6件）／F2b（文書）→ 2周目（班C／班D）→ マイクロ修正 M1（installer 正規名ガードの値判定・`ReadOnly` ガード＋`RecordSaveMark`・MASTER_SPEC 矛盾・CLAUDE.md 容量）**まで全消化**。
`final-gates` 実施（09-03）: lint ERROR 0／LO compile OK／pure **PASS 3,004 / SKIP 14**／`--dev`・`--prod --zip` 自己検証 PASS／`bin_roundtrip` prod・dev OK（6条件）／`lo_xlsm` 157/157・155本コンパイル OK。配布物 `dist/MyBookshelf.xlsm`（`azure_embed_key` 空・`mock_llm` FALSE・22シート・157モジュール）と `MyBookshelf_配布.zip` を生成。
**実Excel（Windows）合格（2026-09-04 実機第20報）**: 最終候補 zip（build `9539809`）を D: に展開し bat から起動 → VBOM 無しで全画面が組み上がり、🩺診断は全✅（AIリボン検出・22シート）。**R35 の本題は達成。** 同報で新規不具合1件: パック出力は成功するが**パック取込が E0701**（`modPack.ImportPackFile`「'Open' メソッドは失敗しました: 'Workbooks' オブジェクト」）→ R36 の先頭で扱う（下記）。
**R35 追補（09-04）**: F3a＝パック取込・引き継ぎ読み込みの Open 失敗を是正（開く前に先頭4バイトで暗号化判別→普通のファイルは Password 引数なしで開く。`modPack.OpenGateReason` 純関数＋Pure37 に5件。PASS 下限 3,009）。F3b＝**ランチャー往復 bat**（OneDrive 原本→D:\MyBookshelf\ 複製→30秒監視・10分ごと書き戻し→閉じたら最終書き戻し＋`_前回` 退避。D: 上に展開した運用は従来どおり起動のみ。3プロセス・最小化・失敗は通常窓）。**bat は当環境で実行不能＝実機1往復（spec §10-3 の2）を確認するまでテスターへ配らない。** spec §10 参照。
（経過）09-03 に送った1本目はテンプレートの `Sheet1` を PerformanceCache のごみで焼いていた（`read_modules` の MODULEOFFSET バグ。riskconsulting へ報告済み・F1 で是正）。是正版 V1〜V4 と最終候補をユーザーへ送付済み。**テスターへの再配布は実Excel 合格後**。姉妹PJの Mac 実Excel では「標準モジュール＋ThisWorkbook は OK・クラスモジュールだけ属性8行と MODULEPRIVATE が要る」（うちはクラス無し）。
**09-05 追記**: `docs/05_図解ガイド.html`（出荷ビルド典拠の1枚もの説明書）を追加し `--zip` に同梱（README 先頭で案内）。**髙橋フィードバック（09-05）の司令塔調査結果（実装で再現済み）**: ①スクショ取込＝リボン `ChatGPTV`（モデル指定なし・resolution=high）で全文転写→`my_knowledge.full_text`（veryHidden）。**読み取り結果の全文ビューア・編集→保存し直しは無い**（抜粋 150/300/600 字のみ。`modVaultGallery:730/260`・`modPeek:32`）。png/jpg ファイルは `📁 追加` の FileDialog で選べない（`SUPPORTED_EXTS` に無い。入口は 📸 のみ）。②❌違う→「修正ナレッジ」は普通の資料として追加されるだけ（`modAppAct.RecordCorrection:118-120`。**元の質問文が本文に入らない・優先度なし・元の誤資料は残る**。トースト「次から…この内容で答えます」は過大表現）。訂正の注入は凍結の modAsk/modRetrieve/modPrompts 外＝`modAskRetrieve`/`modSparse` 側で設計する。③削除は 🗑1件／🗑部門ぶん一括／カード[はい]の3経路。全削除・複数選択削除は無い。④`publish_key` は発行者ブックだけの片側鍵（ボタン表示＋発行直前の再入力照合のみ）。**一般ブックは pack.xlsx を形式と次元数だけ見て無条件に取り込む**（署名なし・運用は共有フォルダ権限で担保）。⑤フィードバックは2系統（Hub 📮＝匿名・共有フォルダ `feedback\fb_*.txt`／ヘルプ 📮ご意見・不具合報告＝mailto のみ・`feedback_mail_to` 既定空）。→ R36 候補として裁定待ち（ユーザーへ選択肢提示済み）。
**姉妹PJ riskconsulting からの伝言（09-05）**: ①環境変数で OneDrive を探す方式は向こうも撤去中（W11-b。`%OneDrive%` が別利用者を指す件は安全上の欠陥として扱われた）②bat はメールで配らず OneDrive/SharePoint 配布フォルダのみ。初回持ち込みは「手順書に bat 平文（約30行）を載せ、会社PCのメモ帳で ANSI 保存」＋出荷検査で手順書の写しと生成 bat を機械照合 ③修正版 `ovba_write.py`（クラス属性8行・MODULEPRIVATE をクラスにだけ・`read_modules` の MODULEOFFSET 尊重・非ASCII名 fail-closed）を送付予定 → **添付がセッションに届いていない（要再送）**。届いたら `build/ovba_write.py` を差し替え（本体不触の運用）→ 検問チェーン → **クラス1本入りのスモーク xlsm を作ってユーザーの Windows 実Excel で開けるかを確認し、向こうへ返す**（向こうは Mac 実Excel でしか未確認）④向こうは配布 bin から `ServerXMLHTTP` を撤去（直接API経路を開発ビルド専用へ）。うちは `modGatewayDirect:78` の1箇所（`embed_transport=direct` 時のみ・出荷既定 ribbon）。**REPORT_BIN_STRINGS に追加して件数報告に載せた（09-05）。R36 で WScript.Shell と一緒に扱う。**
**向こうの質問への回答（09-05）**: (1) SharePoint 経由の bat: **前版は bat 入り zip を SharePoint から各テスターがダウンロードして使えた実績あり**（ユーザー報告）。bat 単体のアップロード可否は未確認 (2) OneDrive 直接起動時の `ThisWorkbook.Path` が URL になる件: **うちは実測なし**（spec §10-2 の前提として設計で避けた） (3) `data_dir.txt` 方式は不採用。データは xlsm 内部のまま、bat が xlsm ごと `D:\MyBookshelf\` へ複製して書き戻す。保存先は bat の `%~dp0`（原本）と固定 `D:\MyBookshelf\` で完結し、VBA は OneDrive を知らない。R36 のデータブック分離で置き場所の決め方を裁定する（data_dir.txt 方式も候補）。
**09-05 夕**: `ovba_write.py` を W9.3 版へ差し替え（spec R35 §11。差分は MSForms 参照の除去だけ・全モジュール一致）。`bin_roundtrip` OK／`lo_xlsm` 本番 157/157 OK／クラス入りスモーク S2 も LO OK。**S1/S2 をユーザーへ送付（Windows 実Excel 待ち。判定表は spec R35 §11）。** **R36 起草**: `spec_20260905_R36_是正と本文表示と置き場所.md`（§1 置き場所は 1-A 推奨で GO 待ち／§2 是正・§3 本文表示・§4 画像📁 は先行実装）。
**R36 実装波投入（09-05 夜）**: spec `ba86e58` で契約確定。波1（Sonnet・§3 本文表示＋§4 画像📁・worktree）と波2（Opus・§2 是正・worktree）を並行投入。LO は司令塔が直列で回す取り決め。波の完了後: 司令塔がマージ→lint→LO compile/pure→`EXPECTED_PASS_MIN` 更新→build→bin_roundtrip→lo_xlsm→敵対的レビュー1周目→Fix→2周目。§1（置き場所）は 1-A/1-B の分岐でユーザー GO 待ち。
**09-05 実機第22報**: スモーク **S1 ○／S2 ○**（Windows 実Excel・D: と共有サーバーの両方。spec R35 §11 判定結果）→ 新版 `ovba_write.py` で確定。**§1 置き場所はユーザーが 1-A（版上げ時の自動引き継ぎ）を選択（「任せる」）→ 波4 を Fix 波の後に投入。**
**発行者用ブックは未再生成**（`MYBOOKSHELF_PUBLISH_KEY` が要る。ユーザー判断待ち）。R36 の本題は「消えない置き場所」（spec §7。髙橋案＝エンジンと蓄積の分離。推奨 (a) ランチャー往復 → (b) データブック分離）。

**起点**: 09-02 事務局の VBOM 異議 + Excel 落ち報告 / 09-02 AMSI 検知（受付番号 202609020376）/ 09-03 D: シャットダウン消滅の実測 / 09-03 姉妹PJ riskconsulting から `build/ovba_write.py` 提供。
**裁定と実測のすべて** = `spec_20260903_R35_配布方式転換.md`（実装波はそれだけ読めば着手できる）。

**状態（2026-09-03 夕）**: 波1（ビルド側・Opus→レート制限でSonnetへ引継）・波2（検問移植+AMSI縮小・Sonnet）・波3（文書・Haiku）**完了**。敵対的レビュー1周目（Opus 2班: 壊す／残骸と縫い目）**実行中**。
司令塔の再検証（全部自分で再現）: lint ERROR 0／LO compile OK・pure **PASS 3,004 / SKIP 14**／`--dev`・`--prod`・`--prod --zip`・`--prod --vba-mode installer` 自己検証 PASS／`tools/bin_roundtrip.py`（olevba）prod・dev OK／`tools/lo_xlsm.py` **LO 列挙 157/157・155本コンパイル PASS**／src の `VBProject`・`AddFromString`・`ExecuteExcel4Macro` 0件／容量 modUI 残280・modUIShelf 残96。
配布物: `dist/MyBookshelf.xlsm` 1.36MB・bin 2,148,864B・157モジュール（155+ThisWorkbook+Sheet1）・22シート（`vba_src` 無し）・CP932 置換 330 字。
**実Excel は未検証**（検証用ファイルをユーザーへ送付済み・結果待ち）。**`dist/MyBookshelf_発行者用.xlsm` は 09-01 の旧方式のまま**（再ビルドに `MYBOOKSHELF_PUBLISH_KEY` が要る。ユーザー判断待ち）。
記録: コミット `3a042df` は仕様書のコミットだが、前任 Opus が死ぬ直前にステージしていた `build/ovba_write.py`（759e41b 版）と `build_mybookshelf.py` の import 2行を同梱している（中身は意図どおり。履歴は書き換えない）。
**環境の前提**: LO 検問（`lo_xlsm.py`）には `libreoffice-calc` が要る（無いと Excel 製の xlsm すらライブラリ0本）。`bin_roundtrip.py` には `oletools`（pip）が要る。どちらも無ければ exit 2（環境不備）で緑にならない。
**会社PC の環境変数（09-05 実測）**: `%OneDriveCommercial%` は未定義、`%OneDrive%` は**別利用者（win11admin）のフォルダ**を指す。環境変数で OneDrive を探す設計は不採用（spec §10-2 記録）。bat は `%~dp0`（bat を置いた場所）を原本とする現行のまま。
riskconsulting の一時公開は **09-05 に private へ戻した**（ユーザー報告）。
**会社PCへの持ち込み経路（09-05 実測）**: Gmail→会社メールは `.bat` も `.bat.txt` も受信時に自動削除される（xlsm は通る）。zip は Gmail 側が送信を拒否（bat/exe 入り）。**bat と Ghostscript が入った配布物は髙橋さん経由（Teams/SharePoint）でしか会社PCへ入らない。** 会社PCのローカル（D:・Documents・Downloads）は日跨ぎ/シャットダウンで消え、残るのは OneDrive・共有サーバ・デスクトップのショートカットだけ。

## 0R34. R34（外部レビュー裁定と検索精度強化・2026-08-20完了）

**起点**: ユーザーが Gemini 3.1 Pro に依頼した外部レビュー（バグ指摘5件+RAG/UI改善提案）の裁定依頼。
仕様と全裁定 = `spec_20260820_R34_外部レビュー裁定.md`（検証5班の結果・採用/却下理由・レビュー2周の裁定まで全部ここ）。

**検証結果**: バグ指摘5件中、完全成立0件（棄却3・一部確定2）。前提の事実誤認多数
（モデルは gpt-5.5・チャットは社内リボン`ChatGPT()`経由でAPI直叩きではない・UI提案4件は既存実装済み）。
**外部レビューは必ず実コードで反証してから採用すること（R33の46/102と同じ教訓）。**

**実装したもの**:
- A1: modGenPipe.ParseVerdict の「PASS+後続文」無言消失を可視化（trailing>100字で改稿へ昇格・findingsは指摘本文のみ）
- A2: 古い `summary_*.bak`（24時間超）の掃除。**実体は modChatLog**（core層都合。BakIsStale/BoardSweepStaleBak。
  「集めてから消す」2周構成 = modInsightIo.GcOldInsights と同作法。列挙中Killは禁止）。呼び出しは modShare:657 の1行
- A3: modChunker:606 の生 Left$ → SafeLeft（src全体で唯一残っていた非全角安全の切り詰め）
- B0: **modApp 分割手術**（SaveTurnForRestore の実体を modAppState へ。modApp 残36→488）
- B1: **機械的出典突合を ⚡すぐ聞く/🔍しっかり にも適用**。実体は modMode.AnnotateIfNeeded
  （modAsk の LastHit* アクセサから突合表を再構築。書式は modPrompts.SourceTag と一字一句一致をテストで固定）。
  配線は modApp:219 の1行 + 共有境界（modInsightIo.EmitVerifiedQA 入口 / modAppAct の訂正共有）
- B2: 再ランク抜粋 300→700字（凍結 modPrompts の数値1個の最小手術・検算29,550字<60,000）
- B3: 🔍に前後チャンク結合。**1周目レビューでBLOCKER3件**（RunMultiRetrieve内で件数を増やした=
  件数バッジ水増し/スコープ広げ直しゲート恒真化/DemoteUsedで実ヒット押し出し）→ Fix波F1で
  **凍結 modAsk.RunDeepFlow への最小手術**（入念の nUse と同型「近傍は根拠でありヒットではない」）へ位置替え。
  `mLastGenN`/`LastGenHitCount()` 新設 = 突合表だけ生成に使った件数まで読む（近傍出典の誤判定防止）
- C: max_context_chars 既定 40,000→60,000（VBA側フォールバック40,000は意図的に据え置き）+ 管理者ページに戻し方/上げ方手引き

**重要な不変条件（R34で新設・壊すな）**:
- 件数バッジ・実況・usage_log・スコープゲート・DemoteUsed は**実ヒット数**(mLastNHits)。
  プロンプトと出典突合表だけが**生成に使った件数**(mLastGenN/LastGenHitCount)を読む
- AnnotateCitations は**冪等ではない**。AnnotateIfNeeded を同一文字列に二度通す経路を作らないこと
  （現状: modAppローカルansに1回 / 共有境界は未注記のmLastCleanAnswerに1回 / thoroughはShouldAnnotate=Falseで素通り）
- 凍結 modAsk は R34 で modAskFocus への依存を獲得（残732字=WARN帯。次に手術するなら分割が先）

**テスト**: PASS 3,004 + SKIP 14（modTestsPure37/38 新設。SKIP+2は Pure38 のHit型2群=LOの既知制限、実機で走る）。
**レビュー2周**: 1周目 BLOCKER3/MAJOR1/MINOR4 → Fix波F1〜F5 → 2周目 BLOCKER0/MAJOR0/MINOR2 → マイクロ修正1件。

**R34 記録のみ（次期）**: 訂正共有は先頭200字切りのため注記がほぼ届かない（次ラウンドで注記件数の末尾1行を検討）/
multi_candidates を80超へ上げると再ランクの土俵に載らない候補が出る（既定40は安全）/
bak の mtime は Name で引き継がれる=「24時間」は中身の鮮度/旧UI(nexus_ui=FALSE)は画面に注記が出ず共有本文にだけ付く（安全側の非対称）/
Gemini提案のうち次期候補: 訂正共有の注記・Ctrl+K・Ctrl+Shift+C（容量解消後）。

**実機第19報で見るべき観点（R34分）**:
1. ⚡/🔍で存在しない資料名を出典に挙げたとき「(出典確認できず)」が付くか
2. 🔍の回答が表の但し書き・前後の文脈を拾えるようになったか（B2+B3の効果）
3. 🔍・🧠の応答時間が体感で悪化していないか（max_context 60,000の影響。悪化したら管理者ページの手順で40,000へ）
4. 入念の一般アシスタントで、まれに回答が1段丁寧になる（A1の改稿昇格）ことがあるが誤動作ではない

## 0. R33H 容量実測(2026-08-16 R33H Fix波3・F26。python の `len` 基準)

**R33 の容量裁定表から最逼迫3本が漏れていた件(F26)の是正。以後はこの節が最新。**
`wc -m` は当環境のロケールでバイト数を返す(日本語で約3倍)ので**必ず python の `len` で測る**。
上限30,000字／WARN28,000字／**残り300字未満は分割裁定必須**。

- **実質凍結(残100字未満・1行も入らない)**: `modSkin`27 / `modApp`36 / **`modBoard`36** / `modUINexusDraw`37 / `modChunker`72 / **`modHubStat`77** / `modUIShelf`85
- **逼迫(残300字未満=分割裁定必須)**: `modViewport`118 / `modGateway`125 / `modUIMain`150 / **`modBackdrop`155** / `modUI`167 / `modClarify`233 / `modVaultGallery`286
- **準逼迫(残1,000字未満)**: `modTestsPure24`347 / `modChrome`**350** / `modTestsPure25`365 / `modExtractor`377 / `modTestsPure18`402 / **`modHub`449** / `modTestsPure2`451 / `modChannel`509 / **`modShare`509** / `modShelfStore`576 / **`modUtilText`584** / `modViewport2`688 / `modTestsPure34`705 / `modKnowledge`708 / `modSparse`953 / `modTestsPure33`997
- **その他の要注意**: `modShareRule`1,238 / `modAskRetrieve`1,343
- **受け皿(残り多い順)**: `modTypes`28,831 / `modAppDef`27,268 / `modState`26,645 / `optTts`25,907 / `modChatLog`25,621 / `modClip`25,603 / `modSetupWizard`24,999 / `modFeatures`24,307 / `modExtractorAcrobat`24,291 / `modUIMainShape`22,799 / `modGatewayDirect`22,314 / `modConfig`21,980 / `modInstallCheck`20,946 / `modProgressBar`17,542 / `modTelemetry`17,381 / `modInsightGate`14,668 / `modExtractorExcel`14,347 / `modStarter`13,389 / `modPublishUI`10,536 / `modGuard`9,342 / `modDiag`8,097 / `modDash`5,879 / `modShared`3,448

**注意: `modChrome` はもう受け皿ではない**(残350字)。R32 の HANDOFF/CLAUDE.md にあった「実体は modChrome(open契約)へ」という記述は**失効**しているので、そのつもりで割り当てないこと。
**R33H Fix波3 で減った主なもの**: `modBoard` 104→36(F30の称号キー) / `modShare` 684→509(F22) / `modUtilText` 2,964→584(F31+F22) / `modShareRule` 1,238(不変) / `modHubStat` 72→77(F32でむしろ回復)。

## 1. 現在地

> ⚠️ **この節より上の「最新ラウンド節」（現在は §0R34）が現在地の正。**
> この §1 は R33 時点の記述が残っており、更新されていない。
> 2026-09-02、同じファイルの中に「R34完了」と「R33完了」が並存していて、
> 再開したセッションがどちらを信じるべきか分からない状態だった。
> **ラウンドを閉じるときは、新しい節を冒頭に足すだけでなく、この §1 の
> 見出し直下に「最新は §0Rxx」と書き換えること**（`final-gates` スキル参照）。
> 以下は R33 当時の記録として残す（削らない。当時の判断根拠が入っているため）。

**【R33当時の記録】R33完了(ユーザー要望「全体を徹底的にコードレビューして根治」=26観点の全体監査→所見102件→確定46件+テスト基盤→実装6波→敵対的レビュー1周目4班45件→Fix波3本32件→Fix検証パス2周目2班14件→マイクロ修正12件 全消化)。R33版配布済み+発行者用xlsm個別送付済み・実機第18報待ち。**
仕様=spec_20260815_R33_全体監査.md / 裁定=spec_20260816_R33H_レビュー裁定.md / 生報告=audit_20260815_R33_全体監査_生報告.md。
**テスト2,956件PASS + SKIP 12(実数)・lint ERROR 0/WARN 43(容量WARNのみ)・モジュール155本**(modTestsPure33/34/35/36 新設)。

### R33で最も重要な発見: 「テストが嘘をついていた」

開始時の「2,497件全PASS」は**保証として薄かった**。波1で実測したところ:
- **12群が未実行のままPASSに計上**(LOで原理的に走らない本体が `Check(..., True)` に落ちていた)。SKIPとして可視化し PASS から外した。
- **恒真アサート4件**(実装を呼ばず定数を比較)。恒真を恒真へ置き換えず削除。
- **`RunAll30/31/32` の単一ハンドラで「みんなの困りごと」10群75アサートが無言消失**。先頭に実行時エラーを仕込むと PASS 2497→2421 になることで実証。**R32で新設した51件がここで死んでいた。**
- `modTestsPure16` の `HasSurrogate16` が16進リテラルの `&` 欠落で**常にFalse**=非BMP絵文字の検査が恒真だった(波5bが発見・波5cで根治)。
- **再発防止**: `run_lo_tests.py` に `EXPECTED_SKIP_MAX` / `EXPECTED_PASS_MIN` のラチェットを実装。テストを消して静かにすることも `[SKIP]` を貼って集計から消すこともできない。

### R33の骨子(6波)

- **波1 テスト基盤**: 上記。土台が嘘のままでは他の46件の検証が信用できないため最初に置いた。
- **波2 個人情報と共有の関所(10件)**: PII検知の全角素通り(全角数字/全角ハイフン/全角＠。`ChrW`で組む=CP932でU+FF0DとU+2212が同バイトへ潰れるため)/**社内NW復帰で本棚が消える**(到達性を先に見る。3,660通り総当たり+反転で両方向検算)/巻き戻し2回で誤版復活/パス脱出(`..`と前後空白ピリオド)/**メール判定の根治**(＠より後ろにドット=単価表記`商品Ａ＠１００円`の誤検知を消す)。
- **波3 取込のデータ破損(11件)**: **Excel表の列ズレ**(空セルを詰めて値が別見出しのものになる。git履歴でH-10の適用漏れと証明)/CP932をUTF-8で読む(判定を`ReadTextFileAuto`へ集約+U+FFFD検知)/章要約の世代印/`.Find`の`LookIn`省略3件全部+**lintルール新設**/`KillGsTree`のFunction化/**暗号化ファイルを開く前に先頭バイトで判別**(`.xls`は判別不能として従来経路へ流す分岐をテストで固定=正常な旧形式が全滅する事故の防止)。
- **波4 設定・状態(8件)**: **`GetBool`だけが既定値へ落ちない**(スペース1個・全角ＴＲＵＥで機能が無言全オフ)/`nexus_ui`と`mock_llm`の既定値不一致(mock_llmは13箇所同時に揃えた=1箇所だけ直すと「回答は本物・ベクトルはダミー」の混在を作るため)/`ui_state`の`=`始まり質問が無言失敗し履歴が1つズレる。
- **波5a/b/c 余白の第2案+UI(26件)**: **条件付き書式方式**(8ラウンド誰も検討していなかった第2案。実装は「測る→張る→測る→伸びたら剥がして戻す」の実行時自己検算)/専門家召喚が原理的に出ない・宛先未解決なのに成功表示/ギャラリーの検索欄と絞り込み/取込完了でチャットから飛ばされる/**部門ごとの削除を新設**(取り消せない破壊操作。件数提示+「取り消せません」+既定いいえ+削除後に「もう受け取らないか」)/管理者ページの断定文を事実に合わせる/📊利用状況を`admin_users`優先・未設定時のみ発行者へ。
- **波6 全社12,000人(2件)**: 組織集計が**全社員のビーコンを1本ずつ開いて起動が分単位で固まる**問題を、発行者端末だけが `board\summary.txt` を1本書き他端末はそれだけ読む**集約スナップショット**へ(12,000オープン→1)。

### レビュー2周が止めた実害(省略してはならない理由の更新)

**1周目(4班45件)**: BLOCKER1(**部門削除が全半角同一視で別部門を巻き添えにし、巻き添え側だけ版記録が残って永久復旧不能**)+MAJOR24。複数班が独立到達したものが9件。
**2周目(2班14件)は全件が「Fix波が作った新しい縫い目」**:
- F7が旧UIに**「回答ができました。出典もあわせてご確認ください」と嘘をつかせた**(修正前の「準備できています」より悪化)。
- F7が`ShouldEmitInsight`の門も開け、**聞き返し直後の🔴が質問を部内へ発信**するようになった。
- F2の順序反転(重複→**欠落**)で、最頻経路(`modPack:189`)が`outOk`を受け取らず、かつ`MarkRowsChanged`が起動時の「資料が減りました」警告を**黙らせていた**。
- F10の検算に**確立作法の捨て読み1行が欠落**(`modViewport:435-437`にある「UsedRangeを1回参照して再計算させる」)。無いと8ラウンド溶かしたのと同じ無言失敗に戻る。
- `BoardSwap`の再試行が**唯一生き残った旧版を自分で消す**。
- `BoardCarryTitles`の1発読みで**称号が最大24日消える**。

### R33 記録のみ(次期)

**容量が限界**: `modBoard`残19 / `modUIMain`残83 / `modBackdrop`残197 / `modShare`残248。**次に1行足す前に分割裁定が必須**(modShareの集計スナップショット書式一式を別モジュールへ切るのが自然)。冒頭の「0. R33H 容量実測」節とCLAUDE.md「凍結・容量」を見ること。
その他: M12の引き継ぎは発行者端末の自画面(`mOrgDay`等)に効かない(modBoard残19字で分岐が入らない。12,000人中1台だけ今月が下がって見え得る)/称号が一方通行のラチェット(退職者の称号が消えない・端末交換で0に戻った人が下がらない)/旧形式summaryが混在期間にbrokenと出る文言/`modP2P`の感謝状が`IdHash`の非対称のまま(移行が要る)/`Open For Binary`の共有モード未指定/F23のラチェットが全体1本なので「60件足して60件消す」波は素通り/`modUtilText.BomCharset`も3バイトのために全読み/`modTestsPure35`が`PURE_LOGIC_MODULES`未登録/`CanPublish`が`publish_key`非空だけなので発行者が複数台あるとsummaryを書き合う(リース方式の裁定が要る)/波1の申し送り`[MODEL]`8件(模擬しか通らないテスト)の回収。

### 実機第18報の観点(R33版)

①**余白(最重要)**: 🩺診断の**[余白の敷き詰め(条件付き書式)]**2行を読む。`backdrop_cf`に記録あり＋`backdrop_cf_failed`が「記録なし」なら条件付き書式が効いている。`stage=usedrange_grew`が出ていたら条件付き書式は使えず(自己検算が自動で巻き戻した)背景画像側が塗っている ―― **どちらでも正常**。両方記録なし＋`backdrop_failed`もあるなら塗られていないのでスクショと📥ログを送付。②きせかえで**画面の下半分だけ色が違う**状態にならないか。③チャットで👎→本棚へ飛ばされないか。④ダークでギャラリー検索欄が白いまま・文字が読めるか。⑤専門家ボタンが出て**相手の端末に実際に届く**か(宛先不明なら**ボタン自体が出ない**のが正常)。⑥Hubのアイコン**下の文字**が押せるか。⑦Hubとダッシュボードの使用率が**一致し100%を超えない**か。⑧「🗑 部門ぶん一括」で、いいえ→何も消えない/はい→**自作とパック由来は残る**。⑨📊利用状況が**発行者用ファイルでは出て一般配布では出ない**か。⑩ダッシュボードの「更新」を2回続けて押すと2回目に**1行案内が出る**(無反応でない)。⑪発行者端末を1度開いた後、他端末で「みんなの節約」に数字と「集計時点/N名ぶん」が出るか。⑫OCR取込後に`%TEMP%`に`nxocr_`フォルダが残っていないか。

---
**R32完了(実機第17報8件=調査4班→4波→敵対的レビュー1周目(B1+M5+m12)→Fix波15件→Fix検証パス2周目(MAJOR3+MINOR5)→マイクロ修正7件 全消化)。**
仕様=spec_20260814_R32_実機第17報.md。テスト2,497件全PASS・lint ERROR 0/WARN 31(容量WARNのみ)・モジュール150本(modInsightGate/modBackdrop/modTestsPure31/32新設)。
**余白の最終決着(8ラウンド・数学的証明)**: 実機プローブで(a)`ScrollArea="A1:L15"`でも60行転がる=**ScrollAreaはホイールを縛らない**(R29の「止める」は二度目の誤帰属・CLAUDE.md訂正済み) (b)`Styles("Normal")`も`Styles("標準")`も**エラー1004=標準スタイル方式は実機で使用不能**(R28以来一度も機能していなかった。無言failが8ラウンド気づけなかった構造的原因) を確定。**停止線 S = B + k**(B=使用済み末尾、k=1画面の行数)で、**塗り深さdを増やしても未塗り到達域は常にk行残る(S−B=k、dに非依存)**=塗りでは原理的に解けないことを数式で証明。R18〜R27の7ラウンド失敗の真因。**唯一残った手段=セルを使わずに色を付ける**→`SetBackgroundPicture`(modBackdrop新設・8×8単色BMPを動的生成しHub/Dash/本棚へ敷く。チャットは保護シートのため対象外)。**実機未検証**=第18報で判定。
R32の骨子:
- **波1 みんなの困りごと全面修正(Opus)**: **一度も通しで動作確認されていなかった機能**(IsMineで自分の投稿は自分の板に出ない=1人テストでは板が常に空)。BLOCKER2(訂正投稿がgap板へ混入し訂正本文が投稿者名欄に出る/67日後の一斉再受信→本棚二重登録)+MAJOR7+MINOR群+**匿名化**+テスト51件をゼロから新設。
- **波2 PIIオフスイッチ(Sonnet)**: config `pii_scan_enabled` 既定FALSE(関所は`ExportPackToFile`1箇所)+`gap_keep_days=30`/`gap_dup_hours=24`追加。発行失敗時の見当違いメッセージ(PII中止なのに「共有フォルダの権限を確認」)修正。E0703文言と動作の矛盾+ドキュメント4箇所訂正。**検知ロジックの実態**: 氏名は見ておらず「メール形式」or「10桁以上の数字列」のOR、1件でも当たれば全件中止。ハイフン/半角スペースはランを継続するため部品番号・条番号・フリーダイヤルが誤爆。
- **波3 ページ拡充(Sonnet・VBA容量ゼロ)**: 使い方⑧FAQ新設(TOC_N 7→8・エラーコード表19件)+④拡充(対応形式/第1階層/E0504/量と速さ)。管理者向け6→9章(**②発行者になるには・⑥ファイルの置き場所・⑧よくある質問**を新設)。configシート直編集の手順は意図的に書かない。
- **波4 余白+確定バグ(Opus)**: 着せ替え入口の一本化(`modHelp.OnCycleSkin`が3画面を塗り直さない=「ライトに戻しても黒いまま」の直接原因)/本棚一覧表の古い地色/Hubバッジ帯の下2行が毎描画削除(R31 W2-2の副作用)/無言failのログ化+法則訂正/背景画像方式。
- **レビュー2周+マイクロ修正**: 1周目BLOCKER1(nonce冪等化が保持期間の非対称60日<67日で空振り)+MAJOR5(匿名化がファイル名とuser_idに実名残存/900字クランプでも1046字>MsgBox上限/PII走査がISOタイムスタンプで必ず誤検知し無言停止/「届きます」と出るのに3経路で無言ドロップ/波2の新設テスト3件が恒真)→Fix波F1-F15。2周目がMAJOR3(**日付除外が市外局番4桁の固定電話`0463-12-3456`を丸ごと消し本物のPIIが素通り**=F4の退行/走査対象をClean1に寄せたため改行が空白化し新たな偽陽性/SweepOldBmpが%TEMP%全列挙)+MINOR5→マイクロ修正F16-F22。
R32 記録のみ(次期): **F1の穴は消えたのでなく1周ぶん遠のいた**(行の期限=発信者のcreated_at基準、既読印=自端末の収集時刻基準で延命→既定約day134に再来。`thanks_gc_days<=0`ならday15。根治は(a)保持期間より古いcreated_atを最初から拾わない(新任者が古い投稿を受け取れない副作用)(b)既読印の無期限化(my_stats肥大)のタクシー提示が要る)/F5のトーストとmodAskのMsgBoxが同時表示(modAsk凍結のため限界)/狭窓クランプ時にフッターを描かない=**社内ポータルへの唯一の導線が消える**(窓を広げれば復帰)/Python側config検査はVBA側の既定値リテラルまで見張れない/modBackdropの再試行カウンタはモジュール全体で成功リセット無し(実効1〜3回・安全側)/**匿名IDは同一人物→常に同一ハッシュ=名寄せ可能**(誰かは分からないが投稿群は束ねられる。仕様として許容)/gap手動「解決済み」UIと自己投稿の確認・取消UI/MAX_COLLECT=60/非BMP絵文字の豆腐化/`modUIMain.bas:116-117`の`A1:I40` Font焼き付け/管理者⑧と⑨の軽微な重複/`build_mybookshelf.py:721-722`の古い相互参照コメント。
**容量逼迫(R32後実測。R33H Fix波3で全モジュール再実測したので下記は失効。最新は本ファイル冒頭の「R33H 容量実測」節と CLAUDE.md「凍結・容量」を見ること)**: modSkin残27・modUIShelf残84・modHubStat残103・modHub残159/modChrome残350/modViewport残534/modKnowledge残708/modInsightIo残2,983。受け皿: modInsightGate残14,990(新設)/modBackdrop残10,377(新設)/modShared残11,147/modDash残7,255。
**実機第18報の観点**: ①**背景画像(最重要・未検証)**: ダークにしてホームで下まで思い切りスクロール→**転がった先まで全部濃紺なら成功**、途中から真っ白な帯が出たら不発。Dashboard・マイ本棚も同様(チャットは対象外で白くて正常)。不発なら📥ログに`backdrop_failed`(stage/err番号)が出るので送付。②きせかえ: ダーク→ライトで**3画面とも地色が戻るか**(以前は黒いまま)。本棚はギャラリー→一覧表に切り替えてからきせかえ→見出し行のグレーは残るか。③バッジ名が途中で切れていないか。④正典発行: **発行者用ファイル**で「📤正典を発行」が出るか・部門名/合言葉→発行できるか。⑤管理者ページ: 使い方末尾リンク→②発行者になるには・⑥ファイルの置き場所・⑧FAQが読めるか。⑥困りごと: **2台以上**でA端末が0件ヒット質問→B端末を再起動→板に出るか(1台では原理的に検証不能)。⑦パック出力/発行が個人情報で止まらなくなったか。⑧`normal_style_bg_failed`(err=1004)が📥ログに1行出るのは**想定どおり**。⑨**余白の敷き詰め(R33H F24でDoD8から移送)**: 🩺診断の**[余白の敷き詰め(条件付き書式)]**の2行を読む。`backdrop_cf`に記録あり＋`backdrop_cf_failed`が「記録なし」なら条件付き書式が効いている。`backdrop_cf_failed`に`stage=usedrange_grew`が出ていたら、この端末では条件付き書式が使えず(自己検算が自動で巻き戻した)背景画像側が塗っている ―― **どちらでも正常**。両方とも記録なし＋`backdrop_failed`もあるなら余白は塗られていないので、その画面のスクショと📥ログを送付。
---
**R31完了(実機第16報6件=調査3班→3波→敵対的レビュー1周目(M5+m6)→Fix波9件→Fix検証パス2周目(F2の縫い目2件検出)→マイクロ修正5件 全消化)。R31版配布済み・実機第17報待ち。**
仕様=spec_20260813_R31_実機第16報.md。テスト2,385件全PASS・lint ERROR 0/WARN 31(容量WARNのみ)・モジュール147本(modTestsPure30新設)。
**前提の再訂正(2ラウンド連続の法則転換・最重要)**: R29の「ScrollAreaはホイールを止める」も誤帰属だった。**ホイールの実際の停止線は"焼き付いたUsedRange末尾"**(R31 F-B)。R29実機実証は画面間比較で、FreezePanes有無と焼き付き深度(Nexus400行vs Hub40行)の交絡が分離できていなかった。Hub=毎描画40行(600pt≒1画面弱)、本棚=毎セッション初回412行(6,090pt≒8画面)を焼いており、余白の体感差と完全整合。ResetRowsBelow(UseStandardHeight)は解放力ゼロと確定し撤去。
**LO死角の新種(CLAUDE.mdに追記済み)**: Shapeにハイパーリンクを付けるとOnActionが死ぬ(HyperlinkがOnActionに優先。R30のScreenTipツールチップが全ボタンクリック不能=第16報③⑤の真因。LOでは検出不能)。
R31の骨子:
- **波1 ツールチップ復旧+❓凡例(Sonnet)**: Hyperlinks.Add全撤去(クリック復活・掃除波不要=RemoveChromeが次の再描画で残骸ごと消す)。tipTextはAlternativeTextへ退避し、ツールバー「❓説明」(幅58・Fix波F3でピル❓=使い方ヘルプとの重複解消)で全ボタン説明を`modChrome.FormatLegendBody/ShowLegendCard`のカード表示(塗り必須・自クリックで閉じる・増殖防止・テーマ追随・可視域クランプ=F6/F13)。shared画面にも❓説明(F7)。取込中ガード(F8)。
- **波2 余白3画面根治(Opus)**: `modViewport.ReleaseSheetRowsBelow`(汎用行削除解放・冪等・行1-12絶対保護・ScrollArea掛け直し・削除前FreezeShapePlacement)へHub/Dash/本棚を接続、Nexus版ReleaseRowsBelowは委譲で挙動同一。本棚の412行フォールバックは`modViewport2.SeedRowCap/ShelfSeedRow`(窓ぶん/使用済み下端の大きい方・412上限)で頭打ち。`ShelfRowHigh(ws)`シグネチャ変更(呼び出し2箇所更新)。
- **波3 伸縮両端揃え+管理者ページ(Sonnet)**: `StretchToolbarRows`(段ごと比例配分・最終段除外・+25%上限・rowRightリセット=F1)でツールバー右端をヘッダーピルと同一式(L+W-8。代数的に同一と検証済み=R25の12pt欠陥は本件非該当)へ整列。`_make_admin_guide`=「管理者向け」シート新設(6章+2パス目次・誰でも可読・合言葉非記載・使い方ページと相互リンク・VBA容量ゼロ)。
- **レビュー2周+マイクロ修正**: 1周目M5(rowRight持ち越し/Hub・Dash非冪等往復/❓重複/使い方ページの陳腐化/管理者ページ事実誤認3件=「確認済みフィルタは実在しない」等の危険側誤解)+m6→Fix波F1-F9。2周目がF2の縫い目2件(seed四捨五入vs境界切り下げのズレ/Hubヘッダー高48ptで焼き範囲が実境界より広い)を検出→マイクロ修正F10-F14(`SeedBurn`切り下げ+maxRowCap=境界上限に統一+Optional hdrH/`PrimePaint`切り出し/テスト実効化=ネガティブ2種/凡例カード高さ=min(自然高,可視高-16,480)/早見表文言)。検算: 窓600pt・ヘッダー48ptで焼き=残し=38行が一致し往復ゼロ。
R31 記録のみ(次期): 最終段の孤立空白は設計裁定どおり残る(W=1000の発行者構成でスクショ取込が2段目に孤立・右に約850pt空白。第17報で指摘されうる)/modHubStat.ClearBadgeArea(B10:F60固定)がUsedRangeを押し上げる別経路の疑い(実機で往復ゼロが崩れたらここ)/F12テストの限界=modHub/modDash呼び出し引数の配線ミスは純テストで検知不能/チャットのReleaseRowsBelowにFreezeShapePlacement全Shape走査が乗った(性能のみ・機能無害)/SeedBurnの15pt直書き(DASH_ROW_H変更で無言ズレ)/lastUsed救済項はClear後で実質死にコード/LEGEND_W_MIN=200はvisW<240で横はみ出し残(縮退)/管理者向けタブ色6B7280の黒文字4.31:1(既存ホームタブと同水準)/modHubコメントの「modViewport2.PrimePaint」表記は実体modViewportと不一致(軽微ドリフト)/凡例カードに❓説明自身も並ぶ(自己言及・害なし)/E0204二重decompose疑い=再現手順をユーザーへ伝達済み(入念モード+社内ナレッジ検索+複数論点質問→📥ログでdecomposed行数確認)・第17報待ち。
**容量逼迫(R31後実測)**: **modSkin残27・modUINexusDraw残37・modApp残36=完全満杯**/modHub残159/modKnowledge残146/modUIShelf残217/modViewport2残688/modChrome残837/modViewport残1,038/modTestsPure11残1,244=**次にこの帯を触る波は分割裁定必須**。余裕: modKnowledgeBar残7,134/modDash残7,255/modTestsPure30残20,210。
**実機第17報の観点**: ①ボタン復旧(最重要): 本棚ツールバーの全ボタンがクリックで動くか・**資料の登録/取込ができるか**・「❓説明」を押すと説明カードが出てカードクリックで閉じるか(みんなの解決事例画面でも)。②余白: Hub/ダッシュボード/本棚(一覧・ギャラリー・共有)でホイール最下部→コンテンツのすぐ下で止まるか・表示崩れがないか。③配置: ツールバーの右端がヘッダーと揃ったか・窓幅を変えても崩れないか(狭い窓で最終段に孤立ボタン+右空白が残るのは設計どおり)。④管理者ページ: 使い方ページ末尾のリンクから飛べるか・目次ジャンプ・説明どおりに発行画面へたどり着けるか。⑤表記揺れ検索: 登録復活後に「クマ/熊」等で再テスト。⑥E0204: 入念モード+社内ナレッジ検索で複数論点質問→📥ログのdecomposed行数(1行か2行以上か)+その時間帯の行を送付。⑦usage_logの`row_release`がHub/Dash/本棚表示のたびに出ていないか(出ていたら往復が残っている証拠・ClearBadgeArea疑い)。
---
**R30完了(実機第15報10件=調査5班+ユーザー実機プローブ3回→3波→敵対的レビュー→R30F Fix波9件→Fix検証パス→R30F2マイクロ修正3件 全消化)。R30版配布済み・実機第16報待ち。**
仕様=spec_20260813_R30_実機第15報.md。テスト2,345件全PASS・lint ERROR 0/WARN 27(容量WARNのみ)・モジュール146本。
**余白の真犯人確定(R18以来7ラウンド越しの決着・最重要)**: `modUI.InitUI` の `Rows("1:400").RowHeight=18` が**起動のたびにUsedRangeを400行へ焼き付け**(実機実測$A$1:$M$400=定数一致)、FreezePanes素通りホイールがそこまで転がれた。**行高の明示設定は内部使用範囲へ焼き付き、ClearFormats・保存では消えない。`Rows.Delete`のみが即時解放**(いずれもユーザー実機イミディエイトで実証)。「使うほど伸びる」はR18以来の錯覚で最初から400行あった。Split(分割ウィンドウ)方式は実機で両ペイン独立スクロール+入力欄流出を確認し不採用。**修正後の余白は最大約1画面弱=Excel仕様の理論限界**(最終使用行が画面上端に来るまで転がせる。ゼロ化はFreezePanes廃止=入力欄固定の喪失としか交換できず不採用)。
R30の骨子:
- **波1 余白根治(Opus)**: 400行焼き付け廃止→`modViewport2.FitChatRows/ReleaseRowsBelow/ReleaseRange`新設(暫定焼き→bound+3超の行を削除解放・起動時移行掃除は冪等・旧ブックの$M$400も1回で解放)。ExtendChatBand拡大分岐で新規バンド行へ18pt追随(バンド内全行18pt不変式)・縮小分岐とClearChatで行削除解放。ClearChatはReleaseRowsBelowでなくFitChatRows(上位互換・裁定承認)。
- **波2 VBA一括(Sonnet)**: トーストAutoSize転換(幅算数=実機比約20%過小+icon分未算入が見切れの真因→`modChrome.FitToastHeight`=TextFrame2実測・34〜92ptクランプ・AutoSize不発時はToastHeightForフォールバック。約45箇所全トーストに効く)/E0204誤検知3件目の救済(`<verdict>`タグを実回答認定。「支払い上限額」の「上限」がLooksLikeLimitErrorに誤爆していた)/文言「チャット入力欄(画面いちばん上の白い枠)」4箇所+👉番号ヒント赤太字(dangerトークン6テーマ新設・5.88〜7.07:1検算済み)/おもてなし2件(解決事例「前へ」ガード・仕上げ0件トーストへ機能説明)/本棚ツールバー20ボタンへHyperlinks.Add(ScreenTip)ツールチップ(**🗑削除のみ除外**=OnDeleteSourceのActiveCell.Row依存との相性リスク回避・実機確認後に戻すか判断)/表記揺れ道1=expand段へ他言語・別表記の言い換えヒント連結(`modAskThorough.ExpandPromptWithSynonymHint`・凍結modPrompts不触・thorough系のみ)+modSynonymStore名寄せプロンプトへ英語/ローマ字例。
- **波3 使い方ページ(Python・VBA容量ゼロ)**: `_make_howto`を7章+クリック目次14リンクへ全面刷新(機械検証済み)。マイ本棚ボタン早見表(新設)・削除方法・仕上げの説明・部門+共有フォルダ4ステップ・文脈引き継ぎユースケース・検索のコツ。陳腐化2件も訂正(旧称「ナレッジ倉庫」/存在しないピル名「一覧表」)。
- **レビュー2周+マイクロ修正**: 1周目MAJOR4+MINOR6→採択9(R30F1〜F9: ReleaseRange上限8000クランプ/RefitChatBand+ToggleThemeへFitChatRows=バンド行高不変式の回復/削除ボタンtip除外/使い方ページ文言2件/FitToastHeightフォールバック/縮小分岐slack+3/FitChatRows末尾ScrollBound掛け直し/row_release観測ログ)。2周目(Fix検証パス)がF2+F9の**非冪等**(毎リサイズ約1,961行削除+ログ)を検出→R30F2マイクロ修正(F10 `ChatSeedRow`純関数で焼き範囲をbound+3へ頭打ち=通常経路削除ゼロ/F11 逆転範囲ガード/F12 「どのボタンも」断言の限定)。2周目の「ツールチップ中途半端」指摘はホバー(表示のみ)とクリック(リンク発動)の混同として**棄却**・現状維持。
R30 記録のみ(次期): mChatBandRow desync(窓縮小→ExtendChatBand縮小分岐で18pt保証欠落→ScrollToBottom約±2行ずれ。modSkin分割時に根治)/StyleHintCellのHasPending二重呼び+Err残留→DrawInputArea偽陽性ログの可能性(modUINexusDraw満杯のため)/ScrollToBottomの18pt割り算はバンド内不変式依存(適法・依存明記)/modHubの`Rows("1:40")`焼き付き同型+本棚系ResetRowsBelowのUseStandardHeight毎描画=**いずれもScrollArea防波堤で余白到達不能・無害**/danger赤2種併存(#DC2626固定と#B91C1Cトークン)/row_releaseログがExportAnalyticsCsvへ混入・保持期間押し出し(発火は移行掃除時のみへ激減済み)/CHAT_ROW_SLACKがPrivate ConstのためmodSkinからリテラル3(コメント訂正済み)/E0204の裏の二重decompose疑い=ユーザーへ9:41-9:47のusage_log依頼済み・未受領/チャット画面ピルにはScreenTip機構なし(本棚のみ)。
**容量逼迫(R30後実測)**: **modSkin残27・modUINexusDraw残37=完全満杯(1字も不可・次に触る前に分割必須)**/modApp残36/modHub残50/modGateway残129/modUIMain残129/modKnowledge残160/modUI残167/modClarify残233/modShelfStore残44。余裕: modViewport2残2,639/modChrome残5,845/modAskThorough残13,368。modAsk 28,381/modAskRetrieve 28,657/modSparse 29,047=WARN許容継続。
**実機第16報の観点**: ①余白: 起動→チャットでホイール最下部→**最後の吹き出しのすぐ下+最大1画面弱で停止**するか/会話5往復→🗑クリア→同様か/再起動後も同様か/窓サイズを変えた後の送信で最新回答が見える位置に来るか(記録のみ事項の監視)。②トースト: モード切替3種+着せ替えで全文読めるか・短いトーストが不自然に縮んでいないか。③**ツールチップ(最重要)**: 本棚ツールバー各ボタンにマウスを乗せると説明が出るか+**クリックで従来どおり動くか**(両立を必ず確認。壊れていたら❓凡例方式へ転換裁定済み)。④E0204: 「支払い上限額」を含む質問でエラーログが出ないか。⑤逆質問: 「チャット入力欄(画面いちばん上の白い枠)」案内+👉番号ヒントが赤太字で見えるか。⑥使い方ページ: 目次クリックでジャンプするか・「▲目次へ戻る」が効くか・説明どおりにボタンが見つかるか。⑦検索: 表記違い(クマ/熊等)の質問でヒットの体感が改善したか。
---
**R29完了(実機第14報6件=調査4班+実機3分検証→2波→敵対的レビュー→R29H Fix波4件+F2b→検証パス計3周 全消化)。R29版配布済み・実機第15報待ち。**
仕様=spec_20260812_R29_実機第14報.md。テスト2,322件全PASS・lint ERROR 0/WARN 27(容量WARNのみ)・モジュール146本。
**前提の恒久訂正(最重要)**: 「ScrollAreaはマウスホイールを止めない」はR19のウェブ伝聞が仕様書上「実機実証済み」へ格上げされていた誤り(引用ロンダリング・D班調査)。ユーザー実機で**ScrollAreaはホイールも止める**ことを実証(Hub/Dash/本棚は止まる)。転がるのはチャットのみ=**FreezePanes併用シートではホイールが素通りする**のが真相。CLAUDE.md・modViewportコメント訂正済み(W2-6)。→ **下余白問題クローズ**(Hub/Dash/本棚=完治実証、チャット=会話画面として自然+地色継続)。⑥本棚削除も正常動作確認でクローズ(行クリック→🗑の2手順設計。「状態」列=✅⏳⚠️🖼の表示専用アイコンと説明済み)。
R29の骨子:
- **波1 ④文脈引き継ぎ根治**: 一般thoroughの起草→査読→改稿のうち**履歴を知るのは起草だけ**で、査読が文脈由来の情報を「根拠不明」指摘→改稿が削って一般論化していた(実機usage_logの失敗4件=loops=2/pass=0と整合)。GenHistoryBlock新設で3プロンプト本文へ会話ブロック埋め込み+査読/改稿に「先行会話の情報を根拠不明扱いしない」+共通sysに指示語解釈の1文。履歴空ならR28とプロンプト同一(検証パスで代数確認・ただし守るテストなし=記録)。「直近1往復」トースト誤表記も是正。5往復化は元々実効(config=5実測・CFG_KEEP_KEYS外)。
- **波2**: ②入力欄案内の訂正(「下の入力欄」→実位置は**画面最上部のFreezePanes固定領域**=R28文言が実装と矛盾していた)+逆質問pending中のC4ヒント「👉ここに番号」(InputHintText+RedrawInputHint)／③トースト表示時間可変(ToastWaitMsFor: 全角15字/秒・下限3,000/上限9,000ms)／⑤着せ替え巡回順=darkを2番目へ(msad⇄lightがほぼ同色で1回目が空振りに見えていた)／豊田合成対策=DiversitySwapPickに相対スコア下限70%(config diversify_min_ratio_x100・0で旧挙動。「1資料だけ正しくヒット」と「偏り」を区別できず無関係資料を混ぜていた)／削除ボタン右端+danger赤(#DC2626白字4.83:1・実体modKnowledgeBar.ToolbarSpec)／modTestsPure29新設。
- **R29H Fix**: F1=ヒント残留3経路(クリア/モード切替/失敗パス)へRedrawInputHint配線／F2=トースト下限3秒の波及で起動最大14秒固まり→**F2bで起動系お知らせ4本(modIntegrity×2/modBoard/modMentor)をトーストからチャットバブルへ転換**(QueueBubble/FlushPendingBubblesパターン=InitUI前確定・InitUI後表示の再利用可能な型)+CycleSkinティザーは成功トースト末尾付記へ統合(🔒解放条件の動的列挙)／F3=履歴保存の";;;"サニタイズ／F4=多様性下限の負値穴を安全側(非正スコアでは介入停止)へ。
R29 記録のみ(次期): modBoot内QuickHealthCheck警告トースト1本のみ旧仕様(凍結のため・警告発生時のみ発火)／トースト同名Shape掛け替えレース(待機9秒化で窓拡大・Tag照合が根治)／F3テストが実装式の書き写し検証=実装を守っていない／波1の「履歴空なら同一」を守るテストなし(Private群のためラッパPublic化が要る)／quick/deepはリボン引数のみで本文会話ブロックなし(thoroughのみ埋め込み)／HISTORY_MAX_TURNS=3ハードコード(modAsk凍結)／nx_mentor_replyがmChatBottom未更新+modStarter早期Exit時SettleChat未到達(既存)／描画がui_state書込を伴う副作用(InputHintText→HasPending→TTL失効時ClearPending)／nexus_ui=False構成では起動案内が出ない(既定TRUEで実害なし)／テーマ名の表記揺れ(さくら/海/金 vs フルネーム)／lint誤字QueueBuble／絵文字幅は切れる側。
**容量逼迫(R29後実測)**: **modApp残36(実質凍結・次は実体退避が先)**／modUIMain残129／modSkin残163／modClarify残229(波2で+80)／modUIShelf残223／modSparse残953／modHub残50／modKnowledge残160／modShelfStore残44＝**次に触る波は分割裁定必須**。modAsk 28,381/modAskRetrieve 28,493/modVaultGallery 28,356/modSparse 29,047=WARN許容(R28/R29裁定)。
**実機第15報の観点**: ④RAGで特約の金額回答→一般へ切替→「請求3万円なら？」で**直前の数字を使った計算**(自己負担1,000円→2万円請求なら9,000円上限)が返るか／「要約してくれよ」で直前回答の要約が出るか／一般モード内で3ターン以上話が続くか／②逆質問→「画面上部の入力欄」案内+入力欄下の「👉ここに番号」→迷わず番号を送れるか(全角「１-①」)／③モード切替トーストが全文読み切れるか／⑤着せ替え1回目でダークになるか／熊の質問で無関係資料が出典に混ざらないか／起動時に画面が固まらず、案内(週次サマリー等)がチャットの吹き出しで出るか／削除ボタンが右端の赤ボタンで見つけやすいか。
---
**R28完了(実機第13報7件=調査4班→プリフライトGO3件→4波→敵対的レビュー→R28H Fix波→Fix検証パス2周目 全消化)。R28版配布済み・実機第14報待ち。**
仕様=spec_20260812_R28_実機第13報.md。テスト2,286件全PASS・lint ERROR 0/WARN 26(容量WARNのみ)・モジュール145本。GO裁定: ①下余白=道1(Normalスタイル方式)/⑥modAsk=B'最小手術/②〜⑦一括GO。
R28の骨子:
- **波1 ①下余白の設計反転(7ラウンド続いた真相)**: R18以来の「塗りを内容下端で打ち止め」はホイール素通し(実機実証済み)の下で原理的に無力(1ノッチ目で塗りの外の白が見える)。R27の「塗り継続」実体はdash限定の埋め草13ptのみだった。反転=`Styles("Normal").Interior.Color=bg`(modChrome.ApplyNormalStyleBg・冪等・On Error保護・want=0黒焼き付きガード)をSetup*Columns4関数+modSkin.ApplyThemeの5配線で全6描画経路へ。既存塗りは二重防御で温存。**LOはNormalスタイルが空振りするため効果は実機でのみ検証可能**。副産物=一覧表モードの背景ゼロ塗り(darkで全面白)も同時解消。Fix: darkで一覧表の文字が読めないB-1(黒文字on濃紺1.17:1)→ApplyShelfTableTextColorでtext色明示(17.06:1)・ExtendChatBand縮小分岐xlNone→bg明示塗り(M-1)。
- **波2 ②逆質問根治(2段構えの故障)**: 一次=番号パーサ2本併存(quick系PickSourceは半角のみ/thorough系ParseChoiceNumbersは全角正規化)で全角「１-①」の資料番号だけ消失→ParseClarifyReplyへ統一(丸数字=意図番号の契約・ClarifyLoneIntentで単独数の意図読み)。二次=「対象の資料:」はただの文字列で検索は本棚全体のまま→clarify_pick_src(ui_state・1ターン限り・TakeClarifyScopeで読んだ瞬間消す)でmodRetrieveの既存scope引数へ実配線(不足時は全体へ広げ直し+clarify_scope_widenログ・modAsk/modRetrieve不触)。深掘りボタンはpending中「下の入力欄に番号を」トースト誘導。入力欄NumberFormat="@"(半角1-3の日付化根絶)。Fix: モード切替でpending印が残り次の質問が1資料に閉じ込められるM-3→OnToggleModeでClearPending(=モードを跨いだら逆質問は破棄)。
- **波3 小物+5往復**: 3モード強化(QuickRules新設/ThoroughDraftRules拡充800-1200字/gen_thorough_verify_verbosity config化)・トースト高さを段数テーブルから幅の算数へ(全角10.5pt/半角5.25pt・AscW負値補正・34pt床/92pt上限=「社内ナ」切れ根治)・ギャラリーPageCapFor=cols×2(9固定廃止=5列で10個)・幽霊文字IsHomeActiveガード5経路(SetStage/RenderSourcesPreview/ShowTip+Fix波でRenderAnswer/ShowEmptyShelfHint。StatusBarは維持)・「AI整理中(本棚全体の未処理分)」文言・**followup_max_pairs既定3→5+橋渡しを1往復→全保持往復運搬へ**(1往復4,000字/総量20,016字・CarryKeepMaskで切替往復の逆流重複を除外=Q+A両方一致で判定しR26H F3「回答が違えば別往復」契約を維持)。
- **波4 ⑥modAsk最小手術(B'案・凍結解除は新規2関数の純追加のみ)**: SetPrevMemory/ResetPrevMemory(git diff 削除0行を証明)+BridgeConvMemory直結+OnClearChat配線。これで(i)切替後の深掘りに古いRAG文脈が混ざる実害(R28調査Dで新発見) (ii)クリア後のmPrevU亡霊(旧M-4)の両方を根治。modAsk 28,381字=**WARN超過はR28裁定で許容**(上限内)。lint CONTRACT/LOテストPURE_ALLOWLISTへ最小追記(検査は弱めていないことを2周目で確認済み)。
- レビュー2周の実績: 1周目BLOCKER1+MAJOR5+MINOR7(うち採択10=R28H F1〜F10+F6b)、2周目(Fix検証パス)要修正0・記録のみ8。
R28 記録のみ(次期): ToastHeightForにアイコン+空白15.75pt未算入(「しっかり調べる」が3行境界まで残4.5pt=次に文言を伸ばすと切れる)/RenderAnswerのIsHomeActiveガードがmLastAnswerText・StatusBar解除も落とす(現状到達不能・Hubに質問導線を戻すと罠)/ExtendChatBand縮小分岐の塗り解放機構喪失(裁定済みトレードオフ・上限2000行×A:M)/ClarifyLoneIntent「2-9」型で1個目を意図採用(コメントと実挙動の齟齬・実害極小)/絵文字はToastHeightForで全角1字扱い=切れる側/modUIShelf呼び口のリテラル7が2箇所/逆質問バブルに👍💾等が作用/合成後クエリが画面に出ない/橋渡しヘッダーが最大5個入るノイズ/lint迂回(ws.Parent.Styles)の例外明文化。
**容量逼迫(R28後実測)**: modUIMain残129/modUIShelf残223/modApp残226/modClarify残309/modSkin残340/modHub残50/modKnowledge残160/modShelfStore残44=**次に触る波は分割裁定必須(特にmodUIMainは実体退避が先)**。modAsk 28,381(WARN許容)/modAskRetrieve 28,319(WARN許容)/modVaultGallery 28,356(WARN)。
**実機第14報の観点**: ①各画面(Hub/Dash/本棚一覧/ギャラリー/チャット)でホイールを下・右へ何回転がしても地の色が続くか(1ノッチ目が勝負)。**ダークテーマで本棚一覧の文字が読めるか。ダークで会話クリア後に白い矩形が出ないか**/②「個別特約は？」→逆質問→全角「１-①」で選んだ資料(興行中止特約等)から回答が出るか。逆質問中に深掘りボタン→案内トーストが出るか。半角「1-3」が日付にならないか/③モード切替トーストが全文見えるか(社内ナ解消)/⑤ナレッジ12個でギャラリー1ページ目に10個出るか/⑦入念モード実行中にHubへ移動→幽霊文字が出ないか/文脈: RAG⇄一般を往復して話が続くか・同じ会話が二重に出ないか・クリア→切替で前の文脈が出ないか・3〜5往復前の話を踏まえるか。
---
(以下は過去ラウンドの記録)
**R23完了(実機第9報=調査3班→波A/B順次→敵対的レビュー→R23H Fix波 全消化)。実機配布可。**
R1〜R21・R23まで全ラウンド完了・検収済み・push済み(R22=一般アシスタント3段化+文脈引き継ぎ構想はGO待ちで仕様化未着手)。
テスト1,995件・lint ERROR 0/WARN 17(容量WARNのみ)・モジュール137本。
R23の仕様: spec_20260810_R23_実機第9報.md。骨子:
①コンパイルエラー根治=**実機第9報の「modViewport2.BadgeRowsForが見つからない」はソース起源ではなく自己インストーラの無言失敗**(ペイロードはsrcと完全一致を実測確認・BadgeRowsForを欠くmodViewport2はgit史上不存在)。原因は_INSTALLER_SRC_TEXT(build_mybookshelf.py内・ThisWorkbookストリーム外科パッチ)のOn Error Resume Next下でc.Name/AddFromString失敗がf計上されず、空モジュールのままSaveで恒久破損(R12監査指摘5が未実装だった)。対策=(1a)Err計上+LenB(s)>0なのにCountOfLines<1の実測検証(1b)f>0でMsgBox"Close WITHOUT saving, then reopen to retry"+**ThisWorkbook.Saved=True**(Excel終了時プロンプトの[保存]反射押し対策=レビューBL-1)+CountOfLines読取例外もf計上(MI-1)。圧縮後1,142B/上限1,148B(残6B・インストーラへの機能追加はもう不可能)(1c)verify_buildへペイロード本文完全一致検査(全135本・_vba_src_text共通化で二重実装なし・CRLF正規化・未指定時はerrors積み=黙殺不可)。
②ヘッダーボタン視認性=帯(sidebar+ApplyHeaderDepthグラデ#0B7D6E→#014D44)とボタン(sidebarActive)が全テーマでコントラスト1.2〜1.4:1。対策=4画面5箇所(modHub円形/modDash・modKnowledge Pill/modUINexusDraw HeaderButton+HeaderTheme)へ白枠線0.75pt(帯に対し5.03/9.78:1)。**FA-R23-2b(塗りの明度引き上げ)は敵対的レビューで幾何的不成立(帯色域内でmsad白文字4.25:1へ退行・gold帯明端3:1割れ)と判明し全面撤回・全6テーマ旧値へ差し戻し**。テストはmsad/light/darkベタ固定+sakura/ocean/goldはThanksCountゲートのためmsadフォールバック検証(Pure環境はThisWorkbook不可のため)。
③ホームタブ右余白=修正なし。実機はプロジェクト全体コンパイルエラー(1行も実行されない)状態の古い描画残存であり現行S2の評価材料にならないと裁定。①解消後の実機再確認で残る場合のみ次期(SbWidthFrom上限側防御が候補)。
R21の仕様: spec_20260807_R21_実機第8報.md(裁定はFix波プロンプト+本節に記録)。
**実機検証は未実施**(要検証: 余白の完全消滅=帯≤可視幅/境界≤窓高・⚡仕上げ再提案(世代キー2)・俯瞰の章立て・逆質問rel判定・回答フッターのモード表示)。
R21の骨子: ⑦S1-S7構造完治=測定の単一化(EnsureViewState後のみ測定・dash/galleryの描く→活性化を反転・HSB確定後測定で24pt嘘解消)/帯≤可視幅(SCROLLBAR_W決め打ち廃止・事後検証・HSCroll条件化)/中身右端=ContentRight不変条件(dashセンタリング廃止・gallery弾性5列・table J一発Fit)/境界RowAtFloor切り下げ(塗りはRowAt切り上げで分離)/窓高適応圧縮(CompressFactor=固定・可変分離)/再フィット穴(table/chat)/5値LogFit。
⑧D1-D4=分散判定をpool(絞り込み前40件)基準+相対gap((b1-b2)/b1×100)へ計器修理(新config dispersion_rel_gap_x100=15/thorough_…=20・素cos分離はmodRetrieve凍結下で不可と裁定し合成スコア相対化で承認)/E0204誤爆修正(査読形式[論点漏れ]等+「1.[」前置きを実回答扱い)/thoroughのdigest/draft段へ免責事由網羅の観点追記/一般モード中のトグル注記。
②E1-E3=Visionプロンプト階層化(#=最上位章のみ/##=節/目次行に付けない)+ClassifyLine階層判定+LooksLikeTocPage(目次頁の見出し抑制)+ChapterKeyOf正規化(StripTocTail)+GroupChapters縮退統合(先着60無言切り捨て廃止・outline_cappedログ)+OUTLINE_LOGIC_VER=2世代キー(doc_outline6列目・旧世代は⚡仕上げ再提案・旧5列後方互換)。24頁資料の机上検証: 旧60章→新3章。
R21H Fix=BLOCKER1(チャットヘッダーShape増殖=RefitChatBandがClearChatHeader不通)+MAJOR4(Hub FitsInView契約破れ/CompressFactor分母/E0204[出典不備]欠落+テスト形骸/章キーvbBinary狭窄+旧outline非互換)+MINOR6(未塗り帯/shared・table S1漏れ/校正過大target/O(n²)Dictionary化/テスト形骸/⚡確認文言にAI回数)。
運用変更(本ラウンドから): 実装波は**検問方式**(着手空コミット+項目ごと即コミット)=エージェント死亡時の損失を1検問分に限定。死亡個体へのSendMessage蘇生pingと再投入の併用は衝突リスク(実際に波Cで2個体並走)→**再投入前に旧個体の完全死を確認するか明示停止**。
R20の骨子: ⑦余白3層根治=帯だけでなく(層1)中身の窓幅追随(Dashカード幅130-220pt可変+
帯内センタリング/本棚gallery3-6列可変+table J列吸収/ピルContentRight統一)・(層2)縦の
行高明示範囲を実寸+αへ縮小(Hub60→40行/Dash120→40行/本棚412→実測。UseStandardHeightで
既存ブックも毎回救済)・(層3)測定頑健化(ViewportWidth他ブックガード=FB-5同型/クランプ
1600→3000pt/Dash全画面化順序)+Workbook_WindowResizeデバウンス0.7s自動再フィット
(dev/prodインストーラ両対応・OnTimeゾンビ防止・busy3情報源=modUiLock+modShelf+modShelfSync)/
①深掘りの一般アシスタント対応(ゲートをmGenPrevU系へ分岐・先消し順序逆転・ShowToast化・
OnClearChatの一般履歴クリア新設)/⑧モード可観測性(回答フッターにモード名+段数+所要を永続
表示・dispersionログのmode列正常化・自己エコー防御(全角対応)・thorough専用分散閾値
thorough_dispersion_gap_x100=25・thorough文体差別化はmodAskThorough側追記連結)/
②再取込ゼロの⚡仕上げ(modBackfill新設: 保存済みfull_text+breadcrumbからPhase1(AI0回)→
BuildOutlineFor委譲でPhase2/3。本棚ツールバー⚡+初回トースト)/③④設定導線(modSetupWizard
新設: 部門=商品部/リスコン部/自由入力・共有フォルダFolderPicker化・Hub/タイル/カードの
到達不能config案内全廃・初回2ステップウィザード(nexus_wizard_doneで1回性))/⑤引継ぎ
ファイル説明ダイアログ+パック表記統一/⑥MS&AD配色(msadテーマ既定: PRIMARY #01675B/
DARK #014D44/LIGHT #0B7D6E・自分バブル緑グラデ白字(6.8:1/5.0:1 AA)・☀🌙はmsad⇄dark)/
APP_NAME="MyBookshelf"統一(表示リテラル18箇所もHaiku波で置換)。
R19の骨子: ①余白根治=ScrollArea前提(R18)を廃し寸法で解く(modViewport:
FitBandToViewportのアフィン2点補正で吸収列を可視幅ぴったりへ・ContentRight一本化・
BoundAddr=内容下端+24pt・mChatBottom連動塗り。modSkinのA1:P2000全面塗り廃止)/
②取込完了時「型が一致しません」=Split直渡し3箇所(modSynonymStore)を型付き配列受けへ
+lint検査15(配列型引数へのSplit/Filter/Array直渡し検知。LO Basicは検査しない穴を塞ぐ)/
③「入念に調べる」バナー残留=nx_toastをClearChat+ClearProgressで掃除/
④短文×複数資料の分散シグナル聞き返し(DispersionGapX100+発動非発動両方でusage_log
"dispersion"・gap既定10・資料名明示なら発動せず)/⑤作業用Excel同居フリーズ根治=
Excelの既定インスタンスマージが原因。可視他ブック検知(CohabitOtherCount:
personal.xlsb/アドイン/不可視は除外)+取込前関所(スクショ・手動同期へも拡大)+
ランチャー『MyBookshelfを起動.bat』(excel.exe /x=別プロセス起動)同梱+docs00/10是正+
起動後案内はShowToast。UsedRange膨張判定は倍率→絶対値(横2,400/縦12,000pt)。
R19H中にAPI上限で実装エージェントが死亡→WIP保全コミット(ff1f5b8)→上限回復後に
レジュームで完走(§8の手順が実際に機能した)。
R17Hの骨子: 名寄せ辞書の永久0行バグ(添字)根治+MergeSynPairs純関数化/俯瞰回答の
二重警告排除(WasGlobalTurn+🔭バッジ・全モード入口リセット)/ArticleEnsureのスコープ遵守・
score=0化・0件時seed/同名章キー×複数資料の箱詰め修正/無人同期の章要約を12章超で先送り
(outline_deferred)/HasGlobalSignal(俯瞰語彙14語)で短文俯瞰を発火/chunk_metaセッション
キャッシュ/取込確認文言に章要約時間を明記/既存本棚への再取込案内トースト。
仕様: docs/dev/spec_20260805_R18_実機第5報.md + spec_20260805_R18H_レビュー裁定.md。
R18の骨子: ①バナー可変幅/2行化+ボタン常時前面(modProgressBar新設)+砂時計廃止+2段目
「作業用Excelを開く?」(セッション1回)/⑧manifest偽0根治(chunk_count -1保持+SourceList
実数突合自動修復+保存一本化+置換反転+modIntegrity起動時突合+Temp検知)/⑤非BMP絵文字の
ダイアログ排除+lint検査14+手動同期のinteractive配線/⑦HasCompoundSignal(？×2・。×2)で
短い複合質問も段0へ/E0303分類のusage_log格下げ/②全域書式の範囲限定+modViewport
(ScrollArea)+チャットへ上段移設/③ナレッジ地図の可視化撤去(分析CSV系は温存)/
④hubチップのカード化/⑨フッター+社内ポータル導線(E0905)。
R16の骨子(参考): ①文言統一/②作業用Excel+白画面抑止(初Declare)/③複合質問分解(modAskMulti)
・逆質問番号選択肢・精読(modAskFocus)・深掘り既出降格。
R16の骨子: ①文言統一(節約した時間)/②取込中のPC作業両立(🗔作業用Excel=WScript.Shellで
excel.exe /x・白画面抑止=リポジトリ初のDeclare(user32.DisableProcessWindowsGhosting、
config freeze_keep_banner 既定on)・案内文とdocs)/③入念モードの複合質問分解(modAskMulti、
段0判定→論点別軽量検索→統合→critique/verify→union出典突合。auto/最大3論点。質問全体で
LLM論点数+6回)・逆質問=番号選択肢(clarify、ok=False非回答契約・「1と3」複数選択・TTL30分)・
精読=近傍チャンク束ね(modAskFocus、source基準・thoroughのみ・deep_neighbor既定2)・
deep深掘りの既出チャンク降格(modFollowup、followup全検索に適用)。
次: 利用者の実機検証(R26版で第13報=R27の余白・RAG根治の確認と、R26の3新機能の確認を同時に)。R26H Fix+検証パス2周までクローズ済み。
**R26(知の循環・本命セット)**: 仕様=spec_20260810_R26_知の循環_設計.md+§7追補。開発憲法CLAUDE.md/スキル3本(preflight/adversarial-review-vba/final-gates)/編集後lintフックはこのラウンド前に整備済み(コミット3c1e16d/b0f01a3)。
①一般3段化=modGenPipe新設(起草→別ペルソナ査読→改稿・最大2周・verdict:PASS早期終了・パース失敗時は起草を返す退行)。AskGeneral第3引数(quick/deep/thorough)。config6キー(gen_deep_effort=medium等+conv_bridge)。フッター「🌐一般アシスタント・(モード名)・検証n回」(**レビューB-1: 初版は既存の早期脱出で到達不能=出荷ブロッカーだった。是正済み**)。LLM回数: すぐ1/しっかり1/入念2〜5(仕様の2〜4は誤りと裁定)。mock査読は交互応答(指摘→PASS)でdev検証可能。
②文脈橋渡し=modConvBridge新設(切替時に直前1往復を先頭差し込み+KeepNewestPairsで丸め。**レビューM-1: 初版の置換方式は切替先の3往復を消す退行→差し込みへ是正**)。ヘッダー「【直前の○○での文脈】」冪等付与。4,000字サロゲート安全切り詰め。クリアで完全失効。**M-2裁定: 一般→RAG方向は凍結modAskの遅延ロード仕様により「切替後に🔍深掘り(続けて質問)を押したときのみ」効く非対称。トースト・docs10は実態通りに正直化済み。双方向根治はmodAsk凍結解除が必要=次ラウンドのタクシー裁定案件**。
③洞察カード=modInsightCard新設。💾ピルを評価アクション行7個目に(境界関所の内側=構造的に安全)。取込はmodVault.RegisterKnowledgeText再利用(既存経路・新規実装ゼロ)。資料名💭考察メモ_題名(40字)+本文冒頭に汚染防止定型文。同題名は上書き確認。Q/Aはnexus_hist_u/aを第一情報源+同一ターンならフル本文優先(700字切れ回避のPickFullerAnswer)。💭非BMPファイル名は取込経路(ADODB=Unicode安全)を通ることをレビューで全数確認済み(Kill→FSO.DeleteFile化)。
R26 記録のみ(次期): **modAsk凍結解除の裁定2件**(一般→RAG双方向化/クリア後のmPrevU亡霊がM-4=クリア→RAG回答→続けて質問の窓で復活)/mock4回経路のdev実機確認(usage_log gen_mode loops=2 pass=1)/同会話別題名の2件目がチャンク重複排除で痩せる(m-5)/RegisterKnowledgeTextのtagsText未反映(m-6・既存)/manifest行がサイズ0・日時Now(m-2・💭パスのANSI副作用)/hist_aに";;;"を含む回答の700字版採用(極小)/ParseVerdictグレーケースのテスト追加/**modApp残350・modGateway残453・modUINexusDraw残900・modTestsPure25残412=次はPure26へ・modHub残50/modKnowledge残160/modShelfStore残44は分割裁定必須**。
実機第13報の観点: (R27分)特約2資料の削除→再取込後に「個別特約は？」で特約から回答or逆質問/各画面右端フィット/下スクロール先が塗り継続/ヘルプ・Peek開いたまま送信・クリアで崩れない。(R26分)同じ質問を一般3モードで=深さ・待ち時間・フッターの違い/RAG→一般切替で文脈が続く/一般→RAGは🔍深掘りで文脈が効く/💾保存→本棚に💭カード→検索でヒット→出典に💭表示/会話クリア→切替→前文脈が出ない。
**R27(実機第12報)**: 仕様=spec_20260810_R27_実機第12報.md。調査3班(A=RAG/B=余白設計/C=バグ型横展開監査)→3実装波→敵対的レビュー→Fix波→Fix検証パス(2周目)→マイクロ修正の完全プロセス。
波1 RAG根治: **SparseBoost無上限が主因**(KeyScore≒330×0.06=b1 20.26でcos無力化・曖昧/低ヒット/確信度の安全装置3つが休眠)→sparse_keyscore_cap=10(config・0で旧挙動)+Len clamp 8。**GarbleRatioのAscW符号バグ**(U+8000以降の常用漢字が制御文字扱い=約款文の13.3%が誤bad)+Word構造制御文字(Chr(7)等)誤爆を是正、私用領域U+E000-F8FFは検出継続。制御文字は空白置換(行頭字下げ1字温存)。.doc全滅時はContent.Text一括+Chr(12)分割でページ復元・maxPages/truncated遵守。embed_stepにok=N(成功数観測)。dispersionにtop3資料名。逆質問の名指し規則(3)は2件以上で無効化。多様性は**最終hits確定後のガード付き最小介入**(全件1資料のときのみ最下位1件を次点資料代表と差し替え+hits_diversifyログ。敵対的レビューB-1裁定=pool再配列はtopK枠を代表で埋めて上位を押し出す+分散判定は順序非依存で効かないため作り直し)。
波2 余白: **下余白の正体=境界関所を通らず境界外へ積まれた実コンテンツ(9経路・チャット8)**。SettleChat→ExtendChatBand 1行が本丸+ヘルプ/ツアー/Peekの開閉両側配線+ヘッダートグル3経路の境界再適用。右=列幅1のHub方式横展開(modChromeへSetupDash/Shelf/ChatColumns)。埋め草PadRowToWindow(dash配線のみ=Hub残4字/modKnowledge残160字のため。チャットは対象外)。重ね表示はoverlay floor(modSkin)で衝突封鎖(HideCitations経路含む=検証パス2周目で発見)。行HiddenはR19裁定維持で不採用(保存関所なし+8KB→3MB膨張実測)。
波3 監査15/16件: Enter拒否トーストwaitless/RefitTickロック/Workbooks.Openダミーパスワード2箇所/manifest列"@"(modShelfScan経由=modShelfStore残44不触)/**OnClearChat後の🟢が消したQ&Aを部内共有する穴をNoteAnswerFailed流用で凍結不触封鎖**/作業用Excel失敗トースト化/診断ボタン無言死/Enterハンドラ3本のLeave漏れ/専門家の未表示質問ファイル削除限定/引継ぎ文言/言語トースト集約(modHub -46字)/FollowHyperlink前トーストwaitless3箇所/TrackScreen dash。
R27 記録のみ(次期): 短文質問の聞き返し率上昇の観測(dispersionのb1/rel分布1週間→ambiguous_score_x100再校正。ロールバック=sparse_keyscore_cap=0)/PUA以外の未知化けパターン/config成功だが表示更新失敗時のトースト精度/埋め草のHub・チャット未配線/同期中に入れ子で走るEnter専用ハンドラ20本(裁定必須)/ValidateMigFile・ValidatePackの行数上限とRestoreSheetのトランザクション化/共有フォルダI/OのDoEvents+進捗/telemetry_enabledオプトアウト導線が物理的に不存在/**modHub残50字・modKnowledge残160字・modShelfStore残44字・modAskRetrieve WARN残24字=分割裁定必須**。
実機第13報の観点: ②「個別特約は？」で特約資料がヒットし逆質問または特約からの回答が出るか(切り分け: それでも出なければmy_vectorsに特約のbs::行があるか目視/化け再取込は⚡仕上げでなく再取込が必要)/①各画面band=vis-2か・下スクロールで白でなく塗りが続くか・ヘルプ/Peek開いたままクリアで崩れないか/散発: 言語切替トースト・作業用Excel失敗時トースト・ダッシュ再描画連打。**特約2資料は本文が化けたまま保存されている可能性が高いため、修正版で再取込(削除→取込)が必要**。
**R25(実機第11報)**: 仕様=spec_20260810_R25_実機第11報.md。
③④RAGフリーズ根治=Excel標準のセル結合警告(複数値Merge)が原因。描画系Merge40箇所超が全て無保護だった。対策2層: modUiLock.AlertsOff/On(ネスト対応・深さカウンタ)を全描画エントリ11箇所+**RAG回答描画のWriteSafe(modUIMainShape=敵対的レビューが座標検算で特定した本丸。圧縮窓でHubバッジ結合セルが回答領域A14:H25内に入る)**へ+DrawBadges旧領域のUnMerge/Clear。パイプライン中断時の後始末(バナー消去+ロック解放+ask_abortログ)で「無言砂時計」を構造的に排除。
①右余白=Dashラチェット(吸収列K未リセット→FitTargetのFullyVisibleWidthが前回帯幅を自己参照しSAFE_MARGIN2ptが複利)を吸収列込みリセットで根治/Hub列幅ブロックをmodChromeへ移設(容量)+M:BZ幅1(VisibleRange右端列の差し引き上限保証)/ギャラリー固定列合計815pt→411pt(Fit実効化)。shelf-table25pt下振れとFullyVisibleWidth共有ロジックは意図的に不触(チャットの完璧fit保護)。下スクロール1-2回はExcelホイール仕様限界=対応不能とユーザーへ説明済み。
⑥バッジ16枠=死にバッジ2種修理(qa_shared_total←EmitVerifiedQA成功時/gapfill_total←EmitCorrection成功時のBump配線が存在しなかった)+新3種(thanks5/streak30/thorough10)。単一情報源(BadgeCatalog)でHub/Dash自動追随。modTestsPure24新設(16種固定+サロゲート検査。**&H8000以上の16進リテラルは&サフィックス無しでInteger負値になるVBA罠**をテスト自身が踏んでいたのを検証パスで発見・修正)。
⑤解決事例=壊れていない(設計通り: 一般モードの解決は共有対象外・自分の投稿は自分に非表示の自作自演防止)。空状態文言に説明追記。単独PCテスト手順はspec R25末尾。
R25H=敵対的レビューBL1+M3+m8を全消化+Fix検証パス(2周目)でF5の深さカウンタ未リセット(保護機構の恒久無効化窓)まで検出・是正。
R25 記録のみ(次期): FullyVisibleWidthのキャップ過大(hub23pt/table25pt下振れの共通構造・実機ログuw値で再裁定)/ギャラリー最終ページのカード高(rowsNeed固定)/Hubバッジ帯16種が圧縮窓(br=3)で末尾切れ/ask_abortの早期Exit経路/AlertsOff区間内DoEventsの保存確認抑止/uw>3800pt級ウルトラワイドでのM:BZ範囲外/**modHub残4字・modKnowledge残160字=次に触る波は分割裁定必須**。
**サブエージェント運用の教訓(R25中に確立)**: 監視信号は波の性質で選ぶ——実装波=検問コミット+ソースmtime、調査・レビュー波=トランスクリプトmtime。ただし基盤不調時はトランスクリプトファイルが更新されない偽死があり得る(検問コミットが真実)。3連続初手死亡時は司令塔が直接調査を実施した(調査・原因特定は司令塔の本務)。
**R24(実機第10報)**: R23cで起動成功後の3件。仕様=spec_20260810_R24_実機第10報.md。
(1)本棚ギャラリーで「引数は省略できません」=R21H F3のCompressFactor 3引数化の際にmodVaultGallery:337だけ旧式2引数のまま取り残し(LO Basicは引数数を照合しない=LO死角その2)。修正+**vba_lintに引数数照合検査を新設**(修飾呼び出し7,642件を照合・偽陽性ゼロ設計・名前付き引数6件のみスキップ)。
(2)ボタン視認性=白枠0.75ptは実機で視認不可と判定→**白地反転**(Hub円形6個/Dash HeaderPill/Chat HeaderButton+HeaderTheme=白塗り+RGB(1,77,68)文字・枠。Knowledgeはモードタブ3個のみセグメンテッド文法=active白ベタ/inactive透明+白枠、非タブ4個は白ベタ)。
(3)Dash下余白(3画面スクロール)=**未修正**。調査で(a)ApplyDashScrollBoundは毎描画で呼ばれScrollAreaも再設定 (b)有力仮説=UsedRange焼き付き(DASH_ROWS 120×15pt=1800pt=ちょうど3画面)or ApplyDashScrollBoundのShapeループが過大bottomYを拾う、まで絞ったが一意確定不可→**実機usage_logの"viewport"(dash)5値を第11報で取り寄せて裁定**。
R24H(敵対的レビュー6件全消化): BL-1=Knowledge inactive PillのFill.Visible=0はクリック透過(R18H FA-8の既決違反)→Visible=-1+Transparency=1へ/BL-2=**modSkin.ApplyThemeがチャットの操作pill9個の塗りをsidebarActiveへ毎回上書き**(初回描画のInitUIですら白地が消える。出荷していたらR23より悪化)→白地仕様へ統一/M-3=lintの括弧付き第1引数`modX.Proc (a+1), b`偽陽性→SKIPへ/M-4=Hub絵文字の文字色が白のまま(白地で沈む)→RGB(1,77,68)/M-5=Knowledge非タブ4個へのセグメンテッド誤適用(画面間一貫性破れ)→isTab引数で分離/MI-6=lintのコメント行末`_`が次行を飲む穴。
R24 記録のみ(次期): ギャラリー最終ページのカード高が満ページ想定の圧縮率で不必要に低い(rowsNeed=9固定vs実rowsUsed)/lint引数数検査はWithブロック内`.Proc`と引数順序・型の変化は見えない/🌙絵文字は白地で沈み気味(🎨等への差し替え候補)/modKnowledge残373字・modHub残187字(次に触る波は分割裁定必須)。
**R23c(真犯人確定・最重要教訓)**: R21以来の「modViewport2.BadgeRowsForが見つからない」コンパイルエラーの真因は**BadgeRowsForの引数名`scale`**だった。`Scale`はMS-VBAL公式仕様のreserved-name/special-form(VB伝統のグラフィック命令)で、**本物のExcel VBAパーサは識別子として拒否**(本文`If scale < 0.92`が構文エラー→関数がシンボル登録から脱落→参照側で「メソッドまたはデータメンバーが見つかりません」)。**LibreOffice Basicはこれを通す=LO検査の構造的死角**。実機プローブ(イミディエイトでの直接呼び出し失敗+Debug>コンパイルがBadgeRowsFor自身に着地+該当行の赤色表示)で確定。修正=引数`scale`→`sc`(2行)。恒久対策=vba_lintにMS-VBAL予約名検査(check_msvbal_reserved_names、公式リスト転記)を追加しERROR化。`Dim line`3箇所はリスト非該当+実機実績ありで除外。副次修正=_VBA_PROJECTストリームのMS-OVBA違反解消(Version=0xFFFF化+PerformanceCache 3,054Bゼロ埋め=どのExcelでも必ずソースから再コンパイル)+verify_build検査追加。インストーラDoEvents追加は圧縮後1,144B(pad不能差分4B)で断念。
**R21〜R23bの誤診の記録(再発防止)**: 第9報を「インストーラの無言失敗→空モジュール」と誤診し、R23(失敗検出強化)・R23b(行数検算modInstallCheck)を実装した。これらは実際のバグとは無関係だったが、注入の無言失敗・部分注入への防御としては正しく機能する堅牢化であり残す。教訓: (1)「LOコンパイルOK」は実Excelコンパイル成功を保証しない (2)実機VBEのイミディエイト窓(CountOfLines/Lines/直接呼び出し)とDebug>コンパイルは、リモートから原因を確定できる最強の診断手段(ユーザーに依頼する価値がある) (3)コンパイルエラーのダイアログ位置(呼び出し側)は真の発生源(定義側)と別の場所を指すことがある。
**R23b(実機第9報の再発対応)**: R23配布後もユーザー実機で同一コンパイルエラーが再報告された(旧破損ファイル開き直しか部分注入の再発かは未確定)。敵対的レビューMA-3(部分注入=AddFromString途中切れはCountOfLines>=1で素通り)を根治: modInstallCheck新設(src/core・非凍結・136本目)。ビルド時にvba_srcのD列へ期待行数(Long)を焼き込み(_expected_line_count、_make_vba_srcと_verify_vba_src_bodiesが同一関数を使用)、起動時にVI()が全モジュールのCodeModule実測行数(末尾空行トリム)とD列を突合。不一致ならモジュール名を日本語ダイアログで提示し、インストーラ側でf計上→Save絶対禁止。インストーラからは`f = f + Application.Run("modInstallCheck.VI")`の1行(Run失敗=modInstallCheck自体の注入失敗やコンパイル不能もErr経由でf計上=**事実上の全体コンパイル検問**)。インストーラ圧縮後1,137B/1,148B(残11B)。README刷新(旧xlsm削除必須・Setup NG時はダイアログ2枚→保存せず閉じて開き直し)。テスト2,012件PASS。
R23b 記録のみ(次期): 無修飾Application.Run 4箇所目(modInstallCheck.VI)は失敗時に起動を落とす側なので昇格リスク(次回スケルトン差し替え時に一括ブック名修飾)/f>0時のマクロ無効ガードシート文言とMsgBoxの食い違い/「わざと1本壊したdevビルド」での実機スモーク推奨(Application.Runの全体コンパイル誘発効果の実証)/C列破損偽陰性はD列化で解消済み。
**実機第10報の読み方**: (1)エラー無く開けた→根治成功 (2)「セットアップ検証NG: モジュール名」→部分注入がその機で実在する証拠。モジュール名がログ代わりになる。保存せず閉じて開き直しで自己回復 (3)R23と同じVBEコンパイルエラーがまた出た→ほぼ確実に旧破損ファイルを開いている(新版はSave前に必ず検問される構造のため)。ファイルサイズとconfig!build_stampで版を確認。
R23H 記録のみ(次期): 部分注入検出(AddFromStringが途中で切れた場合はCountOfLines>=1で素通り。根治はmodBoot側検証関数が必要だがmodBoot凍結+インストーラ残6Bのため見送り=MA-3)/インストーラのn=CStr()失敗時の無通知終了(Done直行・既存経路)/信頼設定エラー時のf=135無限リトライ誘導文言/MsgBox不能環境の完全無通知/ペイロード先頭"="の数式解釈リスク(現状該当0本・_make_vba_src側で弾く1行が候補)/modKnowledge Pillのactive=白塗り+白枠=枠不可視(視認性の害なしと裁定済み)。
**R23の実機で最重要**: 開いて「Setup incomplete (N). Close WITHOUT saving...」が出た場合は保存せず閉じて開き直し(ファイルのペイロードは無傷=開き直しで全量再試行)。エラー無く開けたらコンパイルエラー消滅とヘッダーボタンの白枠線、右余白の有無を確認。
R21H 記録のみ(次期): F6の塗り/境界分離はdash/knowledge系のみ適用(modHub残364字等の逼迫でHub/チャット/modVault/modUIMainは未展開・次期追随)/**modHub(残364)・modChunker(残72)・modUIShelf(残773)は次に触る前に分割裁定必須**/俯瞰pickの連結キー截断リスク(BudgetTakeで章キーが切れると照合不能・無言退避)/LogFitのusage_log肥大(4000字リセット後の再記録)/CoalesceSmallChaptersの最終章連結キーがUI露出/SparseBoost無上限の設計見直し(検索品質側)/無言失敗の全数監査(次ラウンド独立項目・ユーザー指示)。
R20H 記録のみ(次期): 部門ラベルのクリック配線(modUINexusDrawピル化・分割裁定とセット)/
Hub左カラム(統計タイル/バッジ帯)の幅追随/ACCENT #07A963の白字未達箇所(既存継承)/
darkテーマのヘッダー帯ブランド色固定/ToolbarContentRight孤児化の契約掃除/
breadcrumb診断の複数チャンク多数決化/MaybeShowBackfillToastの全走査の重さ(2万行級)/
OnTime予約のui_state二重化+mRefitRunning失効ガード/ブック名「'」のOnTime修飾(既存作法全体)/
「ナレッジ」⇔「資料」全面用語統一/docs10目次の全面再構成/modBackfill残りDictionaryの
比較方式統一/modAnalytics等のNexus残存(表示外)/診断画面の「チャンク数(my_knowledge)」は
技術者向けとして意図的に温存。
R19H 記録のみ(次期): 分散聞き返しの意図5区分の冗長(実機フィードバック後に省略検討)/
検査15の名前付き引数・単一行If(現リポ0件)/ViewportWidthの他ブックガード(modUIMain
凍結解除後)/ambiguous_dispersion_gap_x100=10の校正(usage_log "dispersion"の観測待ち)/
**cohabit_detectedの値の意味が「可視な他ブック数」へ変わった(次報のログ読み取り注意)**。
R17H 記録のみ(次期): 同義語追記が DistinctiveKeys 8枠を圧迫し得る(発火後に実測評価)/
名寄せの outline 非依存化(modShelf 凍結解除後)/既存本棚カードへの「再取込で新機能有効」
バッジ表示。Phase1(構造メタ+参照エッジ)は R17波0/波1、
Phase2(章単位要約=疑似グローバル検索)は R17波2、Phase3(enrich常時ON+用語名寄せ辞書)は
R17波3 で実装完了(いずれも下記「R17 で増えたもの」)。**modPrompts の分割は
Phase2/Phase3 とも見送った**(章要約・章選択・俯瞰回答・名寄せの4本とも新モジュール内の
Private プロンプトにしたため。調査agent7 §5-4 が前提にしていた分割は今回も不要だった。
modPromptsは残321字のまま=Phase3でも1文字も触っていない)。
残: 利用者の実機テスト(R21分: 余白の完全消滅(右・下スクロール不能)・⚡仕上げ再提案と俯瞰の章立て(全部教えての出典が本文ページになる)・「免責は？」の逆質問(rel判定)・回答フッターのモード表示+usage_log回収(viewport 5値/dispersion b1,b2,rel/outline_capped))。
本棚gallery列数追随)・⚡仕上げ(旧形式資料の検出→実行→32-34機能有効化)・初回ウィザード
(2回目起動で出ないこと)・MS&AD配色・深掘りボタンの一般モード動作・回答フッターの
モード名表示(3モードの体感差)+usage_log "dispersion"(mode列が実モードに変更済み)/
"viewport"/"cohabit_detected"とask_stepsの回収→thorough_dispersion_gap_x100=25と
ambiguous_score_x100の校正)。
容量の分割必須ライン(次に触る波は先に分割裁定。R21H後の実測。上限30,000/WARN28,000):
**上限接近・分割裁定必須: modChunker(29,928・残72)/modHub(29,636・残364)/
modTestsPure18(29,598)/modTestsPure2(29,542)/modUIShelf(29,227・残773)**。
WARN帯: modKnowledge(28,918)/modTestsPure(28,906)/modHubStat(28,760)/modSkin(28,690)/
modTestsPure5(28,622)/modUI(28,530)/modGateway(28,522)/modChannel(28,476)/
modUIMain(28,358)/optVision(28,312)/modTestsPure12(28,280)。
WARN線直下の凍結・逼迫: modBoot(残8)/modUtil(残8)/modUINexusDraw(残13程度)/
modAsk(残43)/modShelfSync(残55)/modShelf(残65)ほかR20H時点の一覧から不変。
余裕: modViewport2(19,013→残10,987)/modOutlineBuild(25,470)/modBackfill(18,304)/
modDashStat(24,634)/modViewport(25,160)/modAskRetrieve/modClarify(各~26,000)。

| R | 実装者 | 内容 | コミット | 状態 |
|---|---|---|---|---|
| R26H+検証 | Opus/Sonnet | レビュー裁定8件(フッター到達不能=BL/橋渡し置換退行/トースト正直化/mock交互/FSO化/verdict緩和/情報源統一+700字回避)+検証パス合格 | 96de9d7〜f238484 | 完了 |
| R26波C | Opus | 洞察カード(modInsightCard/💾ピル/RegisterKnowledgeText再利用/汚染防止2層/docs10) | 083044c〜2a89dd6 | 完了 |
| R26波B | Sonnet | 文脈橋渡し(modConvBridge/クリア完全性確認/mockPASS/入念案内出し分け) | 60d0720〜357872e | 完了 |
| R26波A | Opus | 一般3段化(modGenPipe/AskGeneral3段/config5キー/フッター/mode説明) | f7e5271〜6770208 | 完了 |
| R27H+検証 | Opus/Sonnet | レビュー裁定6件(多様性作り直し/overlay floor/PUA検出/Chr(12)ページ復元/サニタイズ最適化/言語トースト集約)+2周目指摘2件 | d991f3a〜79e5a21 | 完了 |
| R27波3 | Opus | 監査15/16件(モーダル保護/無言失敗/状態漏れ/再入/化け素通し/UIUX) | b10bd44〜4aa042e | 完了 |
| R27波2 | Opus | 余白構造完治(関所全数配線/列幅1横展開/埋め草/開閉両側復帰) | 〜87431d6他 | 完了 |
| R27波1 | Opus | RAG根治(KeyScore cap/AscW符号バグ/GarbleRatio誤爆/サニタイズ/.doc代替/観測3種) | ac5152c〜e7c0833 | 完了 |
| R25H-Fix+F5 | Sonnet | レビュー裁定9件(WriteSafe保護/Err退避/EnsureLayout保護/M:BZ/自己修復/範囲拡張/絵文字/stage/False窓)+検証パス指摘F5是正 | e45802c〜a271ff0 | 完了 |
| R25波C | Sonnet | バッジ16枠(死に2修理+新3種)+Pure24+解決事例文言+uwログ+M:P幅1 | e3d602d〜0b87f2d | 完了 |
| R25波B | Sonnet | 余白ラチェット根治(Dash吸収列K/Hub列幅modChrome移設/ギャラリー列幅411pt化) | 66349d7〜deee38f | 完了 |
| R25波A | Sonnet | Merge警告フリーズ根治(AlertsOff/On層11箇所+DrawBadgesクリア+ask_abort後始末) | 1b0c265〜9424a26 | 完了 |
| R24H-Fix | Sonnet | レビュー裁定6件(ApplyTheme上書き/クリック透過/非タブ誤適用/絵文字色/lint偽陽性2種) | 0af2a93〜eaa256b | 完了 |
| R24波2 | Sonnet | ボタン白地化(Hub/Dash/Chat/Knowledgeタブ)+Dash下余白調査(未修正・実機ログ待ち) | e22f9dd〜2293eb0 | 完了 |
| R24波1 | Opus | CompressFactor取り残し修正+vba_lint引数数照合検査(7,642件) | 20e0fca〜4cab916 | 完了 |
| R23c | Opus | 真因修正(scale→sc改名)+lint予約名検査+_VBA_PROJECT無害化+verify_build検査 | 5a19ef2〜0e6e597 | 完了 |
| R23bH-Fix | Sonnet | レビュー裁定(D列期待行数焼き込み=実行時再読込全廃/README2枚表記/テストPrivate化/先頭空行ガード) | 1599856〜8ead794 | 完了 |
| R23b | Opus×2 | 部分注入検出modInstallCheck新設+インストーラVI呼出+テスト+README刷新(初回投入は途中死亡→検問から再開) | 8e3b058〜4420dd5 | 完了 |
| R23H-Fix | Sonnet | レビュー裁定(Saved=True/CountOfLines例外計上/sidebarActive全戻し+テスト強化/CRLF正規化/検査スキップ黙殺防止) | 703143a〜b187bc3 | 完了 |
| R23波B | Sonnet | ②ヘッダーボタン白枠線(4画面5箇所)+sidebarActive明度(→Fix波で撤回) | 3a521d7〜27f7e4c | 完了 |
| R23波A | Opus | ①インストーラ失敗検出(1a/1b)+verify_buildペイロード本文一致検査(1c) | 314d725〜09ded26 | 完了 |
| R21H-Fix | Sonnet | レビュー裁定F11(Shape増殖/FitsInView契約/圧縮分母/E0204出典不備/章キーvbText/塗り境界分離/shared・table S1/校正頭打ち/Dictionary化/テスト是正/⚡確認文言) | a27c014〜ecb7477 | 完了 |
| R21波C | Sonnet | ②章検出根治(Vision階層+目次抑制+縮退統合+世代キー2) | c1dbbaa〜ae37768 | 完了 |
| R21波B | Sonnet | ⑧分散判定pool+相対gap/E0204/入念網羅性/一般トグル注記 | d65baea〜ff86b32 | 完了 |
| R21波A | Opus | ⑦S1-S7構造完治(測定単一化/帯≤可視幅/中身右端/境界切り下げ/適応圧縮/再フィット穴/5値ログ) | a61f87f〜c87772c | 完了 |
| R20H2 | Haiku | "Nexus Agent"表示リテラル18箇所→APP_NAME参照+分析CSV名 | 62c10af | 完了 |
| R20-Fix | Sonnet | R20H裁定FA16(ビルド台帳3件/busy3情報源/仕上げロック/行高高水位分離/テーマトグル/管理者行RowX0/段数リーク/ウィザード1回性/全角エコー/部門ピッカー/APP_NAME/文言・docs是正) | 86a65cc〜2ec9a5f | 完了 |
| R20波D | Sonnet | ③④設定ウィザード(modSetupWizard/FolderPicker/導線全付替)+⑥MS&AD配色(msadテーマ)+Dashセンタリング | 74ba6a4〜f4f3cbb | 完了 |
| R20波C | Sonnet | ②⚡仕上げバックフィル(modBackfill)+⑤引継ぎ/パック説明 | 637b28c〜0fb0754 | 完了 |
| R20波B | Sonnet | ①深掘り一般対応+⑧モード可観測性・聞き返し強化 | 8ff2571〜2c43493 | 完了 |
| R20波A | Opus | ⑦余白3層根治(測定頑健化/縦掃除/幅追随/Hub縦詰め/リサイズ再フィット) | f8ec173〜13232dd | 完了 |
| R19-Fix | Opus | R19H裁定FA9(アフィン2点補正/shared下端/同居誤検知除外/HUB_BAND統一/塗り巻き戻し+絶対値判定/関所拡大/文言両対応/ShowToast/docs00)+FB10 | ff1f5b8/6daffd9 | 完了 |
| R19波C | Sonnet | ⑤同居検知+取込前関所+ランチャーbat生成/④分散シグナル聞き返し(HasScoreDispersion/config gap=10) | 9d2736e〜f457f85 | 完了 |
| R19波B | Opus | ①modViewport新設+全画面の余白根治(吸収列/ContentRight/BoundAddr/mChatBottom連動塗り) | 177c79b〜d4b37b8 | 完了 |
| R19波A | Opus | ②Split直渡し3箇所の型不一致根治+lint検査15(検知実証→修正→0件)/③nx_toast掃除(ClearChat+ClearProgress) | 7fdd2ba〜3d23d56 | 完了 |
| R17-Fix | Opus | R17H裁定FA10(名寄せ0行根治/俯瞰二重警告/スコープ遵守ほか)+FB9+FA-2補 | 28c7838/0bf73cd/0a1f7aa | 完了 |
| R1〜R7 | Opus/Sonnet | レビュー対応/取込/統計/UI統合/OCR/UIUX修復 | 〜 | 完了 |
| R8/R8b/R8c | Opus/Sonnet | P2P修正14件+敵対的レビュー16件+再レビュー6件。テスト272→417 | 〜 | 完了 |
| R9 | Opus/Sonnet | Ghostscript同梱(dist常置)+検出4段階+案内カード+--zip+docs | 〜 | 完了 |
| R10-1 | Sonnet | Activate堅牢化とギャラリー救済(err91×4・ギャラリー無反応) | f2239ba | 完了 |
| R10-2 | Sonnet | GS解決の観測性と堅牢化(E0303に候補パス・カード状態記録) | 94c27d7 | 完了 |
| R10-3/3b | Opus | GhostscriptによるテキストPDF抽出第1選択+移設 | 3872a0a/f01b3d2 | 完了 |
| R10-5/5b | Sonnet | 取込進捗の可視化(進捗バナー・完了トースト) | b833692/91e0ac4 | 完了 |
| R10-4 | Haiku | docs追随 | 7cd686e | 完了 |
| R10c | Opus | 敵対的レビュー裁定の修正8件(H1/M1-M5/L1/L5/L6) | e117750 | 完了 |
| R10d | Opus | modShelf分割(modShelfBatch新設)+H2/M4/M5残 | 69a3364/16eeab4 | 完了 |
| R11-A | Opus | データ保全Critical(C1取込中終了禁止/C2Word所有判定/C3失敗検知/C4silent伝播/C5統計防御) | eded8b2 | 完了 |
| R11-B | Sonnet | #30恒久対策+Activate全数(C5ビューポート/C6素のActivate) | 0d9fa66 | 完了 |
| R11-A2 | Opus | Word設定復元の無条件化(外部レビューHigh)+発行後始末統一 | b3d446e | 完了 |
| R11-C | Sonnet | 関所とログの全数配線(BlockIfIngesting/再入ガード/ツールバーログ/委譲先ログ) | ed977de | 完了 |
| R11-D | Opus | GS/COM/実機層堅牢化(観測性/タイムアウト/Dictionary/ComError/Word所有事実判定) | 6fb3707 | 完了 |
| R11-E | Sonnet | 進捗・文言・ヘルプ・UI細部(進捗配線/ヘルプ刷新/死にコード削除/自己取込ガード) | 130e314 | 完了 |
| R11-F1 | Opus | 容量救済の分割6本(純移設。実装側WARNゼロ化) | 3087a33 | 完了 |
| R11-F2 | Opus | 重複統合(UTF-8/移動平均/ページ分割/Timer)+保留回収 | c7587b7 | 完了 |
| R11-G | Haiku | docs同期+実機スモークテスト手順書(docs/45)新規 | 7f2216a | 完了 |
| R11-H | Opus | 敵対的レビュー裁定の最終修正8件(Word保全対称化/バッジ遅延通知ほか) | 1855097 | 完了 |

| R12-4 | Opus | 検索スケール恒久対策(ベクトルキャッシュ/norm_text前計算/binary_rag自動化/進捗+DoEvents) | 13fe5e3ほか | 完了 |
| R13-A1 | Opus | modExtractor分割(modExtractorPdf新設)+一時コピー名ASCII化(mbtmp_fnv16.ext) | 0789a87 | 完了 |
| R13-A2 | Opus | GS根治: フラグ内容検証/空出力の画像PDF分類/生存監視型待機/WMI起動+PID/kill/診断tail化/StageBanner | 3a89aa7 | 完了 |
| R13-A3 | Opus | 薄い抽出の品質ゲート/FriendlyFailMsg/OCR経路の安全コピー/全段階バナー/送信・終了保護/Word試行短縮 | c5313d3 | 完了 |
| R13-B | Opus | 深掘りのスコープ限定検索+会話出典メモリ+armed followup(チップ)+LLM段階計測・実況 | 8515d61 | 完了 |
| R13-C | Sonnet | Hub「みんな」意味化+未接続/未設定表示+部別集計(ビーコンteam列)+係数config化+チャット導線前方化 | edb91ce | 完了 |
| R13-Fix1 | Opus | レビュー裁定: GS判定窓4000字/kill前PID本人確認/OCR上限是正/生存監視の実効化/optGsProc分割 | 16c5d50 | 完了 |
| R13-Fix2 | Opus | レビュー裁定: armed漏れ止め/終了承諾取消/未設定表示/段ログ集約(ask_steps)/スコープ時prefilter回避 | bdae2b2 | 完了 |
| R14-A | Opus | コピー根治(Stream+サイズ突合)+OCR20頁バッチ全量化(optOcrPage新設・上限100頁・正直メモ) | d2a8ba8 | 完了 |
| R14-B | Opus | 質問カウンタ3箇所+一般モードfeedback解禁(誤共有ゲート)+入念モード6段(modAskThorough新設)+可読性 | df5f535 | 完了 |
| R14-C | Sonnet | ヘッダー整列+ギャラリー可視化+着せ替え5件+質問例オンデマンド生成+チャンク拡張子別設定 | 83387a4 | 完了 |
| R14-FixA | Opus | レビュー裁定: GS起動失敗の切り分け/バッチ中断の可視化/共有読みコピー復活+理由4分岐/PID kill/上限メモ正直化 | 6e6a4db | 完了 |
| R14-FixB | Opus | レビュー裁定: 発信の関所(修正/gap/感謝状)/出典タグ]対応/質問例後始末+メモリ/着せ替えガード/docs追随 | 282f183 | 完了 |
| R15波1〜4+Fix | Opus/Sonnet | OCR254頁対応(ETA/中断/再開/確認)+E0202/save_fail根治+レビュー裁定FA9/FB14 | 0ce3f7a〜f65251a | 完了 |
| R16波1 | Sonnet | ①文言統一+②作業用Excelボタン/白画面抑止(初Declare)/案内+docs | 635cfef/24e6f18/74712a7 | 完了 |
| R16波2 | Opus | ③-A複合質問の分解→統合(modAskMulti+段0判定+union出典突合+テスト57件) | d528286/37352c2/461e1d2 | 完了 |
| R16波3 | Opus | ③-B逆質問番号選択肢+③-C精読(modAskFocus)+③-D既出降格+裁定1〜3+docs+テスト82件 | 7251a25〜f72f12a | 完了 |
| R16-Fix | Opus | R16H裁定FA8(精読source基準化/逆質問非回答化/降格全followup化ほか)+FB12 | e71b940/e5e8507/bbef504 | 完了 |
| R18波A | Opus | ①⑪バナー根治(modProgressBar新設)+⑧永続化(modIntegrity新設・保存一本化・置換反転) | ea6b8ed〜cceb2d6 | 完了 |
| R17波0 | Sonnet | chunk_metaシート基盤(SH_CHUNK_META/EXPECTED_SHEETS/modChunkMetaStore)+modShelf圧縮 | 729a75f〜4746a61 | 完了 |
| R17波1 | Opus | Phase1: 構造メタ+参照エッジ(modChunkMeta新設/取込フック/RefsExpand/ArticleEnsure/別表・様式キー/graph_refs/テスト43件) | (本波) | 完了 |
| R17波2 | Opus | Phase2: 章単位要約=疑似グローバル検索(doc_outline/modOutlineStore/modOutlineBuild/modAskGlobal/verdict=global/graph_outline/テスト43件) | (本波) | 完了 |
| R17波3 | Sonnet | Phase3: enrich常時ON(EnrichPending既定maxCount30+modShelfSync小口呼び出し)+用語名寄せ辞書(synonyms/modSynonymStore/BuildSynonymsFor/name_dedup)+クエリ展開(ExpandQueryBySyn/ParseSynResp・modAskRetrieve入口)+graph_synonyms/テスト17件 | (本波) | 完了 |
| R18波B | Sonnet | ⑤非BMP排除+lint検査14+同期interactive+⑦HasCompoundSignal+E0303格下げ | 7d99933〜5e5b2cf | 完了 |
| R18波C | Opus | ②範囲限定+modViewport+チャットへ上段+③地図撤去+④カード化+⑨フッター | 5b1d786〜a41bcfa | 完了 |
| R18-Fix | Opus | R18H裁定FA8(2段目1回化/バナー2行/虚偽警告根治/全重複保護/フッター当たり判定ほか)+FB8 | 294afa4/f9664ef | 完了 |

### R17 で増えたもの(Phase1 分。次に触る人が最初に知るべき5点)

0. **新シート `chunk_meta`(chunk_id / section_path / refs_out)** = 規程の「章>条」と
   本文中の明示参照(「第8条による」「別表2のとおり」)の保存先。**列を足すのではなく
   新シートにした**のは、my_knowledge に列を増やすと配布済み本棚の移行が要るため
   (R17設計書§3)。ビルドが headers-only で焼き込み(`EXPECTED_SHEETS` + `_make_headers_only`)、
   実行時の `Worksheets.Add` は壊れたブックの自己修復専用。
1. **フェイルセーフが全ての前提**: chunk_meta が0行(=まだ取り込み直していない
   既存本棚)なら、参照展開も条番号の直接ヒット保証も**1行も足さずに戻る**。
   その条件は `modChunkMeta.GraphActive(metaCount, nHits)` の**1本だけ**にある。
   ここを2箇所へ書き写してはならない(片方だけ直すと古い本棚で実行時エラーになる)。
   移行処理は書かない=既存資料は「もう一度取り込む」ことでだけ構造が付く。
   これは設計書の明文の判断で、**版移行(modMigrate)のたびに効果を失う**という
   運用上の摩擦もそのまま受け入れている(調査agent7 §5-1)。
2. **新モジュール `modChunkMeta`(src/ingest・PURE_LOGIC_MODULES)** = 抽出の純ロジック。
   `ExtractSectionPath` は breadcrumb 行から、`ExtractRefs` は本文から InStr 走査だけで
   拾う。**VBScript.RegExp / ScriptControl は使わない**(政策ブロックのリスクと
   LO実行テスト不能=回帰を機械で固定できない)。全角/半角の吸収は
   `modSparse.NormalizeForSearch` に一本化してある(取込側と質問側で式が割れると
   「第１２条」が「第12条」で外れる)。
3. **section_path は ApplyCrumb を通す【前】の生チャンクから採る**
   (`modShelf.IngestFile` のループ内 `modChunkMeta.MetaOf` 1行)。保存後の
   my_knowledge から読んではいけない: config `embed_prefix_breadcrumb` が FALSE だと
   ApplyCrumb が breadcrumb 行を丸ごと消すため、admin 設定ひとつで構造が全滅する。
4. **旧メタ行の掃除は「消す前」に呼ぶ**。`modChunkMetaStore.RemoveMetaForSource` は
   my_knowledge を引いて chunk_id を集めるので、`RemoveKnowledgeAndVectorsForSource`
   の**後**に置くと何も引けない。呼び出しは `IngestFile` 手順7.5 と `DeleteSource` の
   各1行。chunk_meta 側の失敗は全て握って `usage_log("chunk_meta_fail")` 1行に留め、
   取込は止めない(検索精度の上積みであってデータ保全ではない)。
5. **検索合流は modAskFocus の2本だけ**: `RefsExpand`(精読束へ参照先を1ホップ・
   **同じ資料の中だけ**・既定8件・score=0)と `ArticleEnsure`(質問が名指しした条番号の
   チャンクが1件も無いときだけ先頭へ最大2件・score は先頭ヒット同値)。
   config `graph_refs`(既定on)のゲートもこの層に閉じてあるので、呼び出し元
   (modAskThorough / modAskMulti / modAskRetrieve)はどれも1行。**LLM呼び出しは
   1回も増えていない**。RefsExpand が資料を跨がないのは、跨いだ瞬間に無関係な規程の
   第8条が正しい出典タグ付きで根拠に混ざるため(利用者が気付けない外し方)。
   - modShelf は圧縮後 27,773字(残227)。次に触る波は先に分割を裁定すること。

### R17 Phase2 で増えたもの(俯瞰質問。Phase1 の上に積んである)

6. **新シート `doc_outline`(source / section_key / summary / keywords / chunk_n)**
   = 章ごとの要約。取込時に章の数だけ AI を呼んで作り(254頁で+3〜8分)、
   質問時は**要約だけ**を読んで「どの章を読むか」を選ぶ。GraphRAG のコミュニティ
   要約を「マニュアルが既に持つ章」で代替する、という R17設計書§3 Phase2 の骨子。
   ビルドが headers-only で焼き込む(chunk_meta と同じ)。
   - **`modBoot.HideInternalSheets` へは足していない**(modBoot は残8字で1行も
     入らない。憲章§4-6)。veryHidden はビルドの焼き込みと
     `modOutlineStore.EnsureOutlineSheet` の自己設定の**二重**で守っている。
     次に modBoot を触る波は、分割と同時にここへ `doc_outline` を足すこと。
7. **章キーは `section_path` の第1要素をそのまま**(`modOutlineBuild.ChapterKeyOf`)。
   追加の正規化を掛けてはいけない(section_path は取込時に
   `modSparse.NormalizeForSearch` を通っており、別の式を足すと**保存側と照合側で
   章キーが割れて、章を選べたのに本文が1件も引けない**=俯瞰が無音で死ぬ)。
   章見出しの無い資料では第1要素が条になる=条単位の要約になるが、それが正しい
   保守的動作(無い章立てを推測するより外れ方が小さい)。
8. **フェイルセーフは `modAskGlobal.OutlineActive` の1本**。doc_outline が0行
   (=取り込み直していない本棚)なら俯瞰は一切動かず、回答は R16 までと完全に同じ。
   章が選ばれない・章のチャンクが引けない・回答生成が失敗も**全て False** で
   従来の入念フローへ落ちる(2回呼んだ後の失敗でも同じ。確実に答えが出る側へ倒す)。
   不発の理由は `usage_log("global_zero" why=…)` に必ず1行残る。
9. **取込側の中断・保存は R15 の枠組みそのまま**: 章の境界で
   `modShelfBatch.CancelRequested`(と Err18)を見て、**そこまでの章を保存して
   正常終了**する。失敗章は `(要約失敗)` の行として保存して続行(行ごと落とすと
   「その章だけ要約が無い」ことが誰にも見えない)。`SaveCheckpoint(1, 120)`。
10. **プロンプト3本(章要約 / 章選択 / 俯瞰回答)は新モジュールの Private**。
   modPrompts(残321字)には置けないという容量裁定で、Phase2 では
   **modPrompts の分割はしていない**。出典タグの書式だけは
   `modPrompts.SourceTag`(唯一の持ち主)を必ず通すこと=ここを自前で書くと
   出典突合(`modAskThorough.CiteIndexFrom`)が全件不一致になる。
   段0の verdict に `global` を足したのは modPrompts / modRagParse / modAskMulti の
   各1〜3行だけで、知らない語は従来どおり single へ落ちる。

### R17 Phase3 で増えたもの(enrich常時ON+用語名寄せ辞書。Phase1/2 の上に積んである)

11. **enrich常時ON化(config `enrich_mode` 既定 off→light)**。取込・同期のたびに
    チャンクの要約・キーワードを少しずつ作る。254頁≒500チャンク規模を一括処理すると
    1時間級になる(設計書§3 Phase3)ため、**取込直後の同期実行では一気にやらない**。
    `modEnrich.EnrichPending` の `Optional maxCount` 既定値を `-1`(無制限)から
    `30` へ変えただけで小口化を実現した(呼び出し元 `modShelf.IngestFile` は
    既存どおり引数無しで呼ぶ=無改修)。`modShelfSync.SyncNow` にも
    `EnrichPending` の1行を新規追加し(既存 `EmbedPending` 直後)、同期のたびに
    少しずつ追いつく後追い巡回にした。**modShelf/modEnrich本体は無改修**
    (司令塔裁定どおり、凍結モジュールに1文字も足していない)。
12. **新シート `synonyms`(term, canonical)** = 用語の表記ゆれ辞書。
    「回収」⇔「リコール」のように意味が同じでも書き方が違う語を吸収する
    (docs/45 項目34)。chunk_meta/doc_outline と同じくビルドが headers-only で
    焼き込む。**`modBoot.HideInternalSheets` へは足していない**(残8字で
    1行も入らないため。次に modBoot を触る波が分割と同時に足すこと)。
13. **新モジュール `modSynonymStore`(src/ingest)** = synonymsのシートI/O
    (Ensure/Write/Read/RemoveAll)**と**名寄せバッチ(`BuildSynonymsFor`)の
    両方を1本に置いた。modChunkMetaStore/modOutlineStore(シートI/Oのみ)と
    modOutlineBuild(業務ロジック)を分けた前例とは違う構成だが、司令塔裁定で
    Phase3の新設モジュールをこの1本だけに絞った指示どおり。当該資料の
    chunk_meta(section_path)・doc_outline(keywords)・my_knowledge(keywords列)
    から用語候補を集め(重複排除・最大200語)、`CallLLM(step="name_dedup")`を
    資料1本の取込につき最大1回だけ呼ぶ。出力契約
    `<syn>表記>正規形|表記>正規形</syn>` は `modRagParse.ParseSynResp`。
    **既存termは上書き**: `ReadMapCsv`で読んだ既存分から新規termと重なる行を
    除き、`RemoveAll`してから全件を書き直す(1つの表に追記と上書きの2つの
    書き方を混在させない設計)。呼び出しは `modOutlineBuild.BuildOutlineFor`
    の末尾に1行(config `graph_synonyms`・失敗握り・`CancelRequested`確認も
    その1行の内側=呼び出し元は増えない)。
14. **クエリ展開は「質問文への同義語追記」方式**。`modRetrieve`/`modSparse`の
    スコアリング本体は無改修(司令塔裁定どおり)で、`modAskRetrieve.
    RunMultiRetrieve`の入口が質問文 `q` を書き換えてから既存の多段RAGへ渡す
    だけ。展開そのものは純関数 `modRagParse.ExpandQueryBySyn(q, mapCsv, maxAdd)`
    で、`modSparse.NormalizeForSearch`を両辺に通して全角/半角の表記ゆれも
    吸収し、双方向(表記→正規形・正規形→表記)・自己一致除外・最大3語・
    質問文に既にある語や追記済みの語の二重追記防止を1本の関数に閉じてある。
    `modSynonymStore.ReadMapCsv`は**1セッション1回だけ**呼び、モジュール変数
    (`mSynMapCsv`/`mSynLoaded`)へ控える。**取込・同期でsynonymsが更新されても、
    開いたままのセッションには次にブックを開き直すまで反映されない**
    (docs/10・docs/30に明記した既知のトレードオフ)。
15. **mock対応**: `modGateway.MockLLMResponse`に`Case "name_dedup"`を追加
    (`<syn></syn>`=0グループ。R16H FB-2の decompose/chapter_summary と同じ
    「Case Elseの汎用ダミーはタグを含まないためパーサが読めない」教訓の踏襲)。
    `enrich`のCaseは既存(R15以前)にあったため今回の追加は不要だった。

### R18 で増えたもの(次に触る人が最初に知るべき5点)

0. **新モジュール `modViewport`(src/ui)** = 各画面の「行ける範囲」の宣言
   (`Worksheet.ScrollArea`)。R18-3b で新設。呼び出しは Hub / マイ本棚(共通クロム
   `modKnowledge.DrawChrome` 1箇所で3モード分)/ チャット(`modUINexusDraw.
   DrawInputArea` 末尾。modUI に1行も入らないため)/ ダッシュボード(実測下端)/
   ナレッジ登録フォームの5箇所だけ。範囲文字列は画面ごとに1つの定数
   (`modHub.HUB_BOUND` / `modKnowledge.SHELF_BOUND` / `modUINexusDraw.NEXUS_BOUND`
   / `modVault.VAULT_BOUND`)で、**書式を当てる範囲と同じもの**を指す。
   ScrollArea を締めるときは必ず「実際に描いた最大到達点+余白」で決めること。
   きつく締めるとボタンが境界の外に取り残され、憲章§3-1違反(見えない=押せない)
   になる。**チャットだけは行方向を締めない**(バブルは行と無関係にpt座標で
   下へ伸び続けるため。A1:P2000)。

1. **新モジュール `modProgressBar`(src/ui)** = 進捗バナー(nx_progress)と
   ■中断 / 作業用Excel の描画・撤去。R18-1a で modSkin(残り10字)から忠実移設。
   呼び出し口は `modUIMain.ShowProgress/HideProgress`、
   `modShelfBatch.ShowIngestBanner`、`modHub.EnsureHubLayout`(SweepOrphans)の4本だけ。
   **バナーを最前面化したら必ず両ボタンを前面へ戻す**(R18-1c)。ここを外すと
   cancellable=False の実況が走るたびにボタンが不透明バナーの下へ埋まり、
   クリックがバナーに吸われて完全に無反応になる(実機第5報①の主犯)。
   幅は `BarWidthFor(modUIMain.ViewportWidth())`= min(760, viewport-16)。
   固定幅に戻してはならない(#30恒久対策の適用漏れが本文を殺した)。
2. **新モジュール `modIntegrity`(src/core)** = データ整合性の観測点。
   ・`ReconcileChunkCount`: カードの件数と my_knowledge の実行数の突合+自動修復
   (`modShelf.SourceList` から1行)。**カードは manifest の chunk_count を出す**
   という事実がここで初めて安全になった。
   ・`RecordSaveMark` / `WarnAtStartup`: 保存成功時に (行数, FullName) を ui_state へ
   控え、起動時に突合する。**zip直開き・一時展開コピーは保存も成功しReadOnlyでも
   ないため、これが唯一の検知手段**(調査agent1 §4「検知ゼロ」)。
3. **manifest の chunk_count は -1 で「前値保持」**(R18-2a)。実データを消して
   いない失敗経路が 0 を書くと、カードだけが「0件」に化ける。新しい失敗経路を
   足すときは必ず -1 を渡すこと。
4. **中間保存(SaveCheckpoint)の呼び出し点は modShelf.IngestFile の Finish 1箇所**
   (R18-2c)。取込の入口をいくつ増やしても保存が漏れない構造にしてある。
   AddFilesResult のループへ戻してはならない(二重保存になる)。

### R12-4 で増えたもの(次に触る人が最初に知るべき3点)

1. **新モジュール `modVecCache`(src/qa)** = セッション内ベクトルキャッシュ+
   **埋め込み世代カウンタ**。my_vectors の vector_csv は1セッション1回しか
   パースしない。世代を進める場所は3つだけ: `modEmbed.EmbedPending`(書込あり時)・
   `modEmbed.MarkAllForReembed`・`modShelfStore.RemoveVectorsByIds`。
   **ベクトルを書く/消すコードを足したらここを呼ぶこと**(呼び忘れると
   「古いベクトルで検索し続ける」という無言の誤り)。取込・同期の開始時
   (modShelfBatch/modShelfSync)は `ResetVecCache` でメモリを解放している。
   構築失敗(err7)は捕捉して従来経路へ落ち、usage_log に `veccache_fallback`。
2. **my_knowledge に第10列 `norm_text`** = 照合用の正規化済みテキスト
   (`modSparse.MatchDocText` が唯一の作り手)。取込時に前計算し、空欄の行は
   検索時に遅延バックフィルする。**行の詰め直しは必ず10列を運ぶこと**
   (modShelfStore の KNOWLEDGE_COLS)。要約・キーワードを書き換えたら
   norm_text を空へ戻す(modEnrich がそうしている)。
3. **config `binary_rag_auto`(既定TRUE)** = binary_rag=FALSE でも
   チャンク数 >= binary_rag_min なら粗選別を自動有効化。止めるには
   binary_rag_auto=FALSE。判定は `modBitwiseOpt.ShouldPrefilter`(純ロジック)に
   1本化してあり、真理値表は modTestsPure9 が固定している。

### R13 で増えたもの(次に触る人が最初に知るべき6点)

1. **新モジュール `modExtractorPdf`(src/ingest)** = PDFフォールバック連鎖+一時コピー
   (旧modExtractorから純移設)。一時コピー名は `mbtmp_<fnv1a64hex16>.<拡張子>`
   (TempBaseNameFor)。元ファイル名は表示/ログ専用。**NFD分解文字・CP932非対応文字
   対策の要**なので、抽出経路を足すときは必ず CopyToLocalTemp 経由にすること
   (OCR経路は modShelfVision が自前でコピー取得+全出口Kill)。
2. **新モジュール `optGsProc`(src/opt)** = GSプロセスの起動(WMI優先→wsh.Run退避)と
   停止(taskkill前に Win32_Process で本人確認)。優待機は optGsTxt.WaitGsTextDone
   = アイドル(vision_pdf_timeout_sec)+絶対上限(gs_abs_timeout_sec 既定1200)の
   二段で、gs_out.log のページ進行/出力サイズが進む限り待つ。どのシグナルが
   効いたかは usage_log `gs_progress_signal`(page=..;size=..;none=..)で観測できる
   (実機初回はこれを必ず確認。両シグナル0なら固定タイムアウト相当に退化している)。
3. **空txtwrite出力の分類** = optOcrCore.ClassifyGsTextResult(純・真理表固定)。
   rc=0+ページ進行あり→「画像PDF」としてERR_303→OCRへ。**「txtwrite出力を
   読めず」でWordへ流れてゴミ登録される事故(RC1)の再発防止の要**。判定材料の
   gs_out読み窓は4000字(300字に戻すとバナー+xref警告で再発する)。
4. **armed followup** = 「続けて質問/深掘り」はInputBoxではなく、チップ
   (Shape `nx_fchip`)を出して nx_input へ誘導し、OnSend が消費する方式。
   消費はモード分岐より先(一般モードなら黙って解除)。RAG↔一般トグルで解除、
   エフォートトグルでチップ文言更新。deepのfollowupは会話出典スコープ内
   多段検索(modAskRetrieve.RunDeepScoped、<2件で無スコープへ自動退避)。
   **スコープ指定時は binary_rag 粗選別をバイパスする**(EnabledScoped)。
5. **ask_steps 1行集約** = LLM/埋め込みの段階レイテンシは usage_log に
   毎段書かず、modGateway のバッファへ積んで modAsk が質問1回につき
   `ask_steps` 1行で書く(usage_log 2000行ローテを圧迫しない)。
6. **ビーコン第10列 team** = 部別集計用(modP2PIo.TeamCodeOf/IsTeamCode/
   SanitizeId経由)。user_department は規約(英大数4-6字)に合う時だけ採用。
   共有パス既定は空文字になった(未設定=「未設定」表示、設定済み到達不能=
   「未接続」表示。ダミーパス出荷は廃止)。

### R14 で増えたもの(次に触る人が最初に知るべき7点)

1. **新モジュール `optOcrPage`** = OCRのバッチ制御(GSを20頁ずつ複数回起動、バッチごとに
   JPEG逐次削除、頁間DoEvents、PID捕捉+タイムアウト時KillGsTree)。上限は
   vision_pdf_max_pages(既定100)。RenderBatch は3値(完了/起動失敗/タイムアウト)で、
   起動失敗は ghostscript_path 案内、途中中断は truncated=True+中断メモ+partial
   (「頁が黙って欠ける」経路は全て塞いだ)。gs_abs_timeout_sec は資料単位の絶対予算
   (バッチ間で残額を配分)。
2. **コピーは2段構え** = Dir$でANSI可視なら旧来の1MB分割・共有読み(ロック中でも読める・
   メモリ一定)、不可視(NFD等)や空読み時のみ ADODB.Stream。サイズ突合必須。失敗理由は
   src_empty/locked/too_big/name_busy/特殊文字 の5系統で文言が分かれる
   (modLog.CopyFailMsgOf が一元管理)。0バイトコピーがGSへ届く経路は消滅。
3. **新モジュール `modAskThorough`** = 入念モード専用6段(再ランク強化→資料要約→下書き→
   自己批判→批判反映検証→機械的出典突合)。deep/quickは不変。出典タグ照合は
   modPrompts.SourceTag が唯一の書式源。検証段エラー注記は Decorate 後に付く
   (履歴/共有に混入しない)。
4. **発信の関所** = 部内への発信(EmitVerifiedQA/EmitThanks/EmitGap/EmitCorrection)は
   全て modAsk.CanShareInsight()(=ShouldEmitInsight(mLastMode, mLastNHits))を通る。
   一般モード・生成失敗ターン・検索0件残留からの誤共有は構造的に不可能になった。
   一般モードの「解決した」はカウント/節約時間のみ加算(共有なし)。
5. **回答の可読性** = modLive.NormalizeAnswerText(markdown混入の保険変換)+vbCr段落化+
   ■見出し段落の太字化(StyleAnswerParas)。バブル本文はvbCr、外へ出す時は
   TargetText がvbLfへ戻す(コピーはCRLF)。復元バブルも同経路。
6. **質問例のオンデマンド生成** = modStarter がシード無しでも my_knowledge から
   1回のLLM呼び出しで生成しmodStateへキャッシュ(指紋=manifest件数+日付シリアル)。
   プレビュー読みは対象8行のみ(全列一括読み禁止)。
7. **チャンクは拡張子別** = chunk_*_chars_{pdf,docx,doc}(target450/overlap100/max900)。
   グローバル既定とExcelは不変。既存資料は再取込しない限り不変(入れ直す場合は先に削除)。

R14で受容した次期課題: ADODB.Stream共有モードの実機検証(開いたままのxlsx/pdf取込=
docs/45項目15で確認)/全角コロン等の異形出典タグは検査対象外(偽陽性なし)/
OCR中断時のJPEG削除は実行される(OCR済みのため実害なし)/一般モードのno_hit gapは
RAG限定のまま/config実キー数は約120(MASTER_SPECは固定値を書かない)。

配布方法: GitHubの「Code → Download ZIP」→解凍→ dist/MyBookshelf.xlsm を開く
(dist/Ghostscript が隣にあるのでOCRも追加作業なし)。
大規模配布は `python3 build/build_mybookshelf.py --prod --zip`。

## 2. 残タスク(優先順)

2. **実機テスト待ち(利用者)**: docs/45(15分スモーク・1台)と docs/44(P2P・2台)。
   不具合報告が来たら次ラウンドとして裁定。
3. 記録済みの次期課題(裁定で受容・未対応):
   - ロック取得のTOCTOU(PoCでは受容。本格展開時に Open For Output 排他)
   - Dir(vbDirectory) イディオム実測24箇所の統一(次期クリーンアップ)
   - SwitchTo / Subscribe/Unsubscribe(意図的未使用・契約表で可視化済み)
   - ダミーローカルパスが Reachable=True(既定expire=0で無害化済み)
   - B9副作用: ホーム表示→Hub描画の間に最大3秒の間(経過観察)
   - 発行者不在部門の共有肥大(docs/30 §9-1 に注意記載済み)
   - nx_progress が nx_ 一括削除対象(受容・再作成されるので無害)
   - PathExists のopt層内相互参照 / 画像PDFでGS起動2回(受容)
   - E0801の OnToChat/OnGoChat側2件(activate_recovered ログで経過観察)
   - State Loss時の mFeedbackDone 消失(二重FB可能になるのみ・受容)
   - チャンク重複排除の先着帰属(設計判断・受容)
   - Excel強制終了時のWordゾンビ(プラットフォーム限界・受容)
   - NextEmbeddingArray のAPI形式変更耐性(現行形式では正・将来課題)
   - PACK_SAVE_FAILED の独自コード(E0704化は次期。docs/30に注記済み)
   - check_raw_activate の回避余地強化(マーカー位置限定等・次期)
   - ArchiveCurrent 初回_archive未作成時の過剰警告エッジ(次期)
   - AddFilesViaDialog の busy 早期戻りが無言(通常経路は関所が先に停止・次期)
   - modBoot 残2,244字(次の機能追加時に分割裁定)
   - modBoot.Boot(:165でApplication.EnableEvents=False)→SyncNow(:464でTrueへ
     復帰)のEnableEvents入れ子崩れ: Boot内5)でSyncNowを呼ぶと、SyncNow自身が
     戻り際にEnableEventsをTrueへ戻してしまうため、Boot手順6)〜8.6)が
     Trueのまま走る(Bootが意図した「起動シーケンス全体を抑止」が5)以降
     効かなくなる)。R16-2e(2026-08-05)で発見・記録のみ。R16では触らない。
   - Q&A長時間経路への🗔/中断の到達手段(R16H FB-4/B-H1): 質問の実行中は❓も
     バナー内ボタンもクリックが届かないため、「🗔作業用Excel」は取込中にしか
     押せない。docs/10は「質問を送る前に開いておく」運用で暫定回避したが、
     Q&A側にも到達可能な導線(送信前の常設ボタン・ESC以外の中断手段)が要る。
   - **R13で受容した次期課題**: 深掘りfollowupは retrieve_mode=single でも
     スコープ内多段を実行(機能仕様として明記・escape hatchの例外)/
     ReadTextTail のDBCS境界で先頭1文字化けの可能性(診断限定・受容)/
     ナレッジ画面ツールバー総幅は未解決(チャットへは前方化で救済済み。
     折返し or リサイズ再描画は次期)/embed_transport=direct 時の質問側
     embed_step 行は未集約(既定ribbonでは無影響)/OCR経路のWMI起動はPID破棄
     (kill不能はOCR側は従来どおり)/ThinExtract のページ数は maxPages
     打ち切り後の値/OnSend弱音キーワード経路はarmed維持/本番ビルドの
     ×ボタン終了は無防備のまま(BeforeCloseはdev構成のみ。終了ボタン側のみ保護)
   - **R18-1gで事実を確定・次期課題化(2026-08-05)**: 上記の「BeforeCloseはdev
     構成のみ」を調査agent0がfile:lineで裏取りした(build_mybookshelf.py:1002-1010
     が Workbook_Open だけを書き込む/Auto_Close に Cancel 引数が無い)。
     **「取込中終了禁止ガードは開発構成のみ」**を MASTER_SPEC §7.6 に明記した。
     実機で×・最小化が効かないのはガードではなく DisableProcessWindowsGhosting の
     副作用。本番でも取込中の終了を守る手段(インストーラ側で BeforeClose を
     注入する/Application.OnKey で退避する等)の設計は**次期**に回す(R18では
     事実の記録のみ。修正はしない=仕様どおり)。
   - **FindKeyRow裁定前提の更新(2026-08-01 R12-5-11)**: spec_20260731_R11
     §9「FindKeyRow線形探索(103キーでms級)」は modConfig.FindKeyRow
     (config = 103キーで固定・安定)の前提であり、modStats.FindKeyRowには
     そのまま適用できない。my_statsは "thx:"/"ins:" のnonce行が記録先
     (modP2P.MarkNonce/modInsight系)で、GC(thanks_gc_days・GcOldNonces)は
     入っているものの周期実行のため、GCが効くまでの間はS(my_statsの行数)が
     数百〜千行に成長し得る事実がある。線形探索の即時最適化は見送り継続
     (体感数百ms。まず前提行数の事実更新のみ)だが、「103キーでms級だから
     全FindKeyRowが軽い」という読み方はしないこと。

## 3. 運用メモ

- 失効タイマー: 出荷既定 knowledge_expire_days=0(消えない)。本格展開時に30以上。
  利用者のconfig手動変更は不要になった。
- 体制: Fable=司令塔(要件・裁定・検収のみ)/Opus・Sonnet=実装/Haiku=雑務。
  敵対的レビュー(読取専用)→裁定→修正→再レビューのループが有効に機能した。
- LibreOfficeテスト(tools/run_lo_tests.py)は多重起動禁止。
- API月次上限で実装エージェントが落ちた場合: WIPを即コミットして保全し、
  軽作業のみで待つ(今回それで乗り切った)。
- モジュール数: 98本(実装)/テスト: 491件/WARN: 3本(全てテスト系)。
  ※R12-4時点: 実装90本+テスト11本=101本(modVecCache/modTestsPure9を追加)、
    LOモード1のPASSは669件、WARNは3本(modTestsPure/2/5)のまま。
  ※R13完了時点: 105本(modExtractorPdf/optGsProc/modTestsPure10を追加)、
    LOモード1のPASSは818件、lint ERROR 0/WARN 3(modTestsPure/2/5)のまま。
  ※R14完了時点: 109本(optOcrPage/modAskThorough/modTestsPure11/modTestsPure12を追加)、
    LOモード1のPASSは1,036件、lint ERROR 0/WARN 3(modTestsPure/2/5)のまま。
  ※R15完了時点: 114本(optOcrEta/optOcrCache/modTestsPure13/modTestsPure14を追加)、
    LOモード1のPASSは1,228件、lint ERROR 0/WARN 4(modTestsPure/2/5/12。全てテスト系)。
  ※R16完了時点: 117本(modWorkExcel/modAskMulti/modAskFocus/modTestsPure15を追加)、
    LOモード1のPASSは1,391件、lint ERROR 0/WARN 4(同上)のまま。
  ※R18完了時点: 121本(modProgressBar/modIntegrity/modViewport/modTestsPure16を追加)、
    LOモード1のPASSは1,486件、lint ERROR 0/WARN 4(同上)のまま。
    R18で記録した確定事項: 全重複再取込はfail_countを進め3回でfailed_permanent(復帰口
    あり・仕様)。ナレッジ地図の可視化コードは a41bcfa の親(75ba385^)から復活可能。
  ※R17完了時点: 129本(modChunkMeta/modChunkMetaStore/modOutlineStore/modOutlineBuild/
    modAskGlobal/modSynonymStore/modTestsPure17/modTestsPure18を追加)、
    LOモード1のPASSは1,626件、lint ERROR 0/WARN 4(同上)のまま。
- **R15で記録した次期課題・確定事項**:
  - ChatGPTV は待ち秒数引数なし・同期永久待ち(実機確認済み)。VBA側の根治は不可能。
    リボン側に引数が追加されたら optOcrPage の呼び出し1行で反映可能。
  - 複数の大型PDF一括取込では事前確認が1件ごとに出る(Nが抽出後にしか判明しないため
    構造上不可避・受容)。確認ダイアログ表示中は後続ファイルも停止する(モーダルの性質)。
  - 手動/自動同期は事前確認なしで長時間OCRが走り得る(silent設計・R11-A C4準拠。
    自動同期は既定OFF。チェックポイントで被害限定)。
  - modEmbed 単独実行(同期の埋め込み再開等)では中断ボタンは出ない(ESC中断は既存)。
    OCR起点で出たボタンはベクトル化中も残り、押せば埋め込みループも止まる。
  - modGateway 全捕捉による ESC(err18)握りつぶし(modEmbed/modEnrich のESC中断が
    CallLLM内で失敗文字列化され得る): 共有コアのため未着手。フラグ方式が代替。
  - SourceList の manifest/pack 同名非合算(pack利用時のみの表示問題・受容)。
  - vba_lint が手続き内 Const を module-level と誤認する(KEEP_DAYS衝突で実測。
    回避は改名。lint側の修正は次期)。
  - SleepMs が modEmbed/modExtractorPdf に private 複製2箇所(共通化するなら基盤層へ)。
  - modShelfSync の分割(現27,819字。次に本格的に触る波は先に分割裁定)。
  - EndClock の「翌 」は2日以上先でも「翌」(現行最大ETAでは到達不能・受容)。
  - キャッシュ復元は「そのバッチをGSが描けた」ことが前提(描画0枚バッチの頁は
    キャッシュがあっても打ち切り。頁番号整合を優先した設計判断)。
  - 32,000字超で切り詰めた頁はキャッシュ保存しない(次回読み直し・FB-10)。
- **R16で記録した次期課題・確定事項**(詳細は spec_20260805_R16H_レビュー裁定.md の「記録のみ」):
  - usage_log のリング(log_max_rows=2000)回転が新タグ(decomposed/multi_*/neighbor_*等)で
    速まる。観測価値優先で容認。次期でタグ整理を検討(節約時間の前月比が早く欠ける)。
  - StatusBar 通知(E0904等)はQ&A中は数秒で上書きされ実質見えない(err_logで追跡可。
    次期でトースト化検討)。取込中は modUiLock の文言を上書きしたまま残る。
  - source型逆質問(「2-②」教育)と topic型(「1と3」)の文言完全統一は次期
    (入力互換は R16H FA-7 の区切り拡張で確保済み)。
  - 逆質問TTL失効直後の番号入力への能動ガード(現状は普通の質問として流れる。
    案内文の「30分で無効」明記のみ実施)。
  - single-thorough/deep の経過時間表示(分解経路のみ実装済み。B-M7残)。
  - 既存の文字列リテラル末尾区切り文字パターン2件(modHelp.bas OnShareSetup付近/
    modTelemetry)。LO実測は通過・実害未確認(EDGE_CASES §1.3 のトリガ条件が実際は
    狭い可能性)。新規コードは Chr$(92) 経由を徹底。
  - mock_llm では decompose は常に single(FB-2で固定)。R16-3系のデモは実LLMのみ。
  - 精読(NeighborExpand)の資料キーは source(表示・スコープと同じ一意基準)。
    chunk_id のハッシュ部は「チャンク本文ハッシュ」であり資料キーではない(FA-1の教訓。
    テストは実データ形状=行ごと異ハッシュで書くこと)。
  - deep深掘りの精読適用は取り下げ(nHits表示契約と衝突・R16H §3C改訂)。
  - 非#ERR文字列は BuildErrorAnswer を素通しする契約(FA-2。非回答ターン=逆質問の表示経路)。
- **R18で記録した次期課題(2026-08-05)**:
  - **ナレッジ地図の再導入案**(R18-4で可視化は撤去。調査agent3 §1-6):
    もし戻すなら「見るだけの絵」ではなく行動につながる形にすること。最小差分は
    (a) 各クラスタ円に `OnAction` を付け、代表ラベルで穴埋めテンプレを作って
    Nexus入力欄へ流し込む(`modHub.OnQuickAsk` と同型。新規ロジックほぼ不要)、
    (b) `ingest_files_total` や総チャンク数を鍵にした再計算キャッシュを入れて
    「開くたびにK-Means+最大約800回のCOM単発読み」をやめる、の2点。
    (b)を入れずに戻すと撤去理由の半分がそのまま戻る。削除したコード
    (DrawClusterMap/DrawBubbles/ClusterColor/MdsCoords/JacobiEigen/TopTwo/
    BuildLabels/TokenizeKw/NearestSource)は 81aa2d9 の親コミットに残っている。
  - `modCluster.LoadVectors` の my_vectors 側は今も1セルずつ読み(M-5の片割れ・
    最大約800回)。分析CSV経路だけになったので体感は消えたが、CSV出力は遅いまま。
    なお `kw()`(keywords列)の読みは R18H FB-3 で削除済み(消費者ゼロだった)。
- **R18H(敵対的レビュー裁定)で記録した次期課題(2026-08-05・FB-4 / B-M3)**:
  1. **本番ビルドに `Workbook_BeforeClose` が無い**(=「取込中は終了できない」
     ガードは**開発構成でしか効かない**)。事実は R18-1g で file:line まで裏取り
     済み(`build_mybookshelf.py:1002-1010` が `Workbook_Open` だけを書き込む /
     `Auto_Close` に `Cancel` 引数が無い)。実機で×・最小化が効かないのはこの
     ガードではなく `DisableProcessWindowsGhosting` の副作用。本番でも取込中の
     終了を守るには、インストーラ側で `BeforeClose` を注入するか
     `Application.OnKey` で退避する等の設計が要る。R18では**事実の記録のみ**。
     MASTER_SPEC §7.6 に「取込中終了禁止ガードは開発構成のみ」と明記済み。
  2. **`my_stats` のキーGC**(調査agent5)。`thx:` / `ins:` の nonce 行が単調
     増加する。最大25KB程度で実害は無いと裁定したが、`modStats.FindKeyRow` は
     線形探索なので行数が増えるほど遅くなる(HANDOFF「FindKeyRow裁定前提の
     更新」も参照)。GCの周期・保持日数は次期に決める。
  3. **deep / quick 向けの軽量 clarify**。論点分解と番号での聞き返しは入念
     モードだけの機能で、deep/quick では複合質問が1本のクエリのまま流れる。
     R18H FB-5 で「回答後に入念モードを1回案内するトースト」を入れたのは
     導線の応急処置であって、聞き返しそのものではない。LLM追加呼び出し1回に
     収まる軽量版(選択肢を出さず「どちらの論点ですか?」だけ聞く等)を次期に
     設計する。判定材料は `modRagParse.HasCompoundSignal` が既にある。
- R11での事実確認・修正メモ:
  - LibreOffice Private Const の参照不可: Public Const へ揃えて回避(modDashStatで実測)。
  - LogError context ラベル: Public エントリ名を指すこと(lintの参照チェックが文字列リテラル内も見る)。
  - modFeatures.InvokeFeature 引数: 最大6個。opt側の引数追加は末尾Optional固定。
  - Dir(vbDirectory): 実測24箇所(記録済み・次期統一)。
