# Nexus Agent — 機能一覧

実装済み機能をカテゴリ別に整理する。各機能の実装モジュール・関連config・
関連ドキュメントを併記する。

## 1. コア: ナレッジ取込・RAG検索

| 機能 | 説明 | 主モジュール |
|---|---|---|
| マルチ形式取込 | PDF/Word/Excel/テキストをフォルダ投入 or ダイアログで取込 | `modExtractor*`, `modShelf` |
| 構造認識チャンク化 | 見出し・条文単位での分割(`chunk_mode=structure`) | `modChunker` |
| 画像PDFのOCR取込 | スキャンPDF(E0303)を同梱外のGhostscriptでページ毎にJPEG化し、1ページ=1回のChatGPTVで文字起こし。上限は`vision_pdf_max_pages`(既定100・20ページずつのバッチ処理)で超過分はpartial。全ページ文字化けのPDFも自動でこの経路へ回る。手順は`docs/43_画像PDFのOCR取込設定.md` | `optVision`, `optOcrCore`, `modShelfVision` |
| 自動同期 | 指定フォルダとの差分同期(`sync_interval_min`でOnTime自動化) | `modShelfSync` |
| 同期完了表示(静かな完了) | 手動同期が正常完了(上限見送りなし)のとき、状態表示行とステータスバーに「✅ 同期が完了しました(…)」と出るだけでダイアログを出さない。上限見送りがあった場合のみ従来どおりダイアログ表示 | `modShelfSync` |
| ハイブリッド検索(標準) | 全件Float内積スコアリング+キーワードボーナス | `modRetrieve` |
| バイナリ量子化ハイブリッド(狂気案Lv.1) | 大規模時にXORハミング距離で粗選別→Float再ランク。`binary_rag=TRUE`でオプトイン | `modBitwiseOpt` |
| 多段RAG | クエリ拡張→マルチクエリ検索→AI再ランク | `modAsk`, `modPrompts`, `modRagParse` |
| 2速度モード | ⚡すぐ聞く(10〜20秒) / 🔍しっかり調べる(下書き+検証の2段) | `modAsk` |
| 続けて質問(深掘り) | 直前の会話を踏まえた追加質問 | `modFollowup`, `modAsk` |
| 出典厳格化 | 資料に無いことは「見当たらない」と明言(ハルシネーション抑制) | `modPrompts.GroundingInstruction` |
| One-Shot品質固定 | プロンプトに理想的な出力例を1つ同梱し、トーン・型・粒度を固定 | `modPrompts.StyleInstruction` |
| Peek View(出典ポップアップ) | 回答直下の出典チップをクリック→元チャンク本文をその場でポップアップ表示 | `modPeek` |
| 低関連度警告 | 検索ヒットの関連度が低い場合(config `low_hit_warn_score` 既定0.3未満)、回答の先頭に「⚠️ 手元の資料との関連が薄い可能性があります」と警告表示 | `modAsk` |

## 2. Nexus SPA UI

| 機能 | 説明 | 主モジュール |
|---|---|---|
| ネイティブUI完全隠蔽 | 数式バー・グリッド線・シートタブ・スクロールバーを非表示 | `modUI.InitUI` |
| デザインシステム | MS&ADグリーン基調・Yu Gothic UI・浅い角丸・柔らかいドロップシャドウ | `modSkin` |
| きせかえスキン(称号連動) | 感謝受領数で解放される5スキン(標準/ダーク/サクラ/オーシャン/ゴールド) | `modSkin.CycleSkin/ResolveColor` |
| Toast通知 | MsgBoxのハイブリッド代替。完了/情報は非ブロッキング通知、確認/エラー/起動時はMsgBox維持 | `modSkin.ShowToast` |
| Empty State | 検索0件時に透かしアイコン+誘導CTA(単なる空白にしない) | `modVault` |
| 考え中バブル | 送信直後に即時表示し、応答到着で本物に置換(体感速度向上) | `modApp.OnSend` |
| グローバルUIロック | 連打・多重発火防止の単一関所 | `modUiLock` |
| ホットキー | Ctrl+Shift+Q(どこからでも一撃召喚)/ Ctrl+Enter(送信) | `modApp`, `modBoot` |
| オンボーディングツアー | 初回起動時のみ3ステップ案内(質問→出典→評価) | `modTour` |
| ヘルプ | 「?」ボタン→ガイドカード+詳細マニュアル導線+ツアー再実行 | `modHelp` |
| 軽量マクロ無効ガード | マクロ無効起動時に案内シートを表示(壊れたUIを見せない) | `modBoot`, ビルド側 |
| チャット履歴シート | 質問と回答(日時・モード付き)が自動で記録される見える化シート。最新が一番上、100件で自動削除。config `chat_log_enabled=FALSE` で無効化可能 | `modChatLog` |

## 3. ゲーミフィケーション・組織連帯

| 機能 | 説明 | 主モジュール |
|---|---|---|
| EXP/レベル | 質問・登録・自己解決・パック共有・ご意見箱で加点。感謝EXPは他者受領時のみ | `modStats` |
| バッジ | 利用実績に応じた達成バッジ | `modStats.EvaluateBadges` |
| 連続起動ストリーク🔥 | 昨日利用していれば+1、途切れたらリセット | `modStats.TouchToday` |
| 節約時間トラッキング | ✅解決1件=15分。日/月/年キーで自動リセット+履歴保持 | `modAsk.FeedbackGreen` |
| チーム連帯ボード | 個人+組織全体の節約時間、月間目標プログレスバー、履歴ポップアップ | `modBoard` |
| ナレッジ自浄(ノイズ投票) | ⚠️報告が閾値(既定2人)に達した資料を全員の検索から論理除外 | `modStats`, `modP2P.CollectNoiseVotes`, `modRetrieve` |
| 管理者による復帰 | `admin_users`該当者はダッシュボードから組織的除外を解除できる | `modDash`, `modP2P.ClearNoise` |
| ガバナンス分析CSV | 質問・節約時間・EXP履歴・クラスタID等をCSVエクスポート | `modAnalytics` |
| ナレッジ地図(K-Means) | 蓄積チャンクをクラスタリングし円で可視化 | `modCluster` |
| ナレッジガチャ🎲 | ランダムに1件を「今日のワンポイント」として表示 | `modApp.OnGacha` |
| 質問テンプレチップ(穴埋め化) | 「改定ポイントは?」等のテンプレをクリック→「【知りたい改定】: (資料名や年度を記入)」のような穴埋めフォームが入力欄に自動入力される。(ここに記入)を埋めてCtrl+Enterで送信 | `modApp.OnQuickAsk` |

## 4. P2P協調機能(ファイルベース・サーバー不要)

| 機能 | 説明 | 主モジュール |
|---|---|---|
| パック共有 | 本棚を`.mbpack`としてエクスポート/インポート(PII検査付き) | `modPack`, `modPii` |
| 感謝状 | ✅解決時、他者由来ナレッジの作者へ感謝状を自動送信。受領でEXP加算 | `modP2P.EmitThanks/CollectThanks` |
| Mentor(専門家召喚) | 回答の出所から専門家を自動特定し、質問を直接送れるボタンを表示 | `modMentor` |
| Mentor受信/返信 | 起動時に自分宛の質問を回収してバブル表示。「✉返信する」で対話化 | `modMentor.CollectQuestions/OnReplyQuestion` |
| 称号(偽装不可) | 感謝受領数5件で💡、20件で🌟を専門家名に自動付与 | `modBoard.TitleFor` |

## 5. 品質・信頼性(利用者からは見えない基盤)

| 機能 | 説明 | 主モジュール |
|---|---|---|
| リトライI/O | 共有フォルダアクセスをAVロック等に対して3回リトライ | `modP2P`, `modMentor`, `modBoard`, `modAnalytics` |
| COMオブジェクト確実解放 | 外部COM(ADODB.Stream/Word/Acrobat等)を正常・異常両経路でSet=Nothing | 全ingest/pack層 |
| APIタイムアウト | HTTP直送信経路にタイムアウト設定(無限フリーズ防止) | `modGateway` |
| 数式インジェクション対策 | CSV/セル出力の先頭記号エスケープ+セル書式のテキスト固定 | `modAnalytics`, `modShelf`, ビルド側 |
| 静的Lintゲート | 契約違反・LO罠・レイヤリング違反を自動検出 | `tools/vba_lint.py` |
| LibreOffice構文検証 | 全モジュールの構文コンパイルを実行なしで検証 | `tools/run_lo_tests.py` |

## 6. 遊び心(依存性を生む余白)

| 機能 | 説明 | 主モジュール |
|---|---|---|
| 時間帯挨拶 | 朝/深夜で異なる労いの一言 | `modApp.TimeGreeting` |
| 弱音への即応(関西弁) | 「疲れた」等の短文入力にAPI非通信でローカル応答 | `modApp.IsTiredWords/ComfortMessage` |
| 関西弁モード | 言語切替の隠し味として追加 | `modApp.OnLangCycle` |
| 週次サマリー | 週初回起動時「先週◯分取り戻しました」 | `modBoard.ShowWeeklySummary` |

## 7. サポート・フィードバック導線

| 機能 | 説明 | 主モジュール |
|---|---|---|
| ご意見箱(バグバウンティ) | 感想・不具合を作成者へメール送信、EXP+5(1日1回) | `modHelp.OnFeedback` |
| P2P接続設定UI | 隠しconfigシートを直接触らせずに共有フォルダパスを設定 | `modHelp.OnShareSetup` |
| フレンドリーエラー | エラーコードを「あなたのせいではありません+次の一手」トーンへ翻訳 | `modLog.FriendlyMessage` |
| 診断の設定サマリー | 🩺診断レポート冒頭に、mock_llm/embed_transport/shelf_folder/共有パス/モデル名/自動同期間隔の現在値をサマリー表示 | `modDiag` |
| エラーダイアログの直近エラーコピー | エラーダイアログ末尾に案内を表示。🩺診断→📋ボタンで、err_logの直近5件(エラーコード・発生箇所・生エラー番号・HTTPステータス・詳細)を整形テキストでクリップボードへコピー | `modLog` |

## 8. 未実装・保留中の機能

`docs/dev/TODO.md` を参照。
