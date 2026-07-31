# 再開手順（セッション中断対策・最終更新: R8がAPI月次上限で中断した直後）

中断したら、次のセッションはこのファイルから読むこと。
リポジトリ: `kojima1459/notebook`、ブランチ: `claude/internal-notebook-lm-chatbot-B6BE7`。

## 1. まず最初にやること

```bash
git log --oneline -10
git status --porcelain     # クリーンなはず(WIPは全てコミット済み)
python3 tools/vba_lint.py  # ERROR 0 のはず
```

**実機配布に使ってよい最後の完成コミットは `284c5af`(R7検収確定)。**
それより後(`3925409` = R8 WIP)は未完成なので、配布ビルドの基点にしないこと。

## 2. 現在地(ラウンド表)

| R | 内容 | 状態 |
|---|---|---|
| R1〜R5 | レビュー対応・取込コア・統計/バッジ・UI統合・敵対レビュー | 完了・push済み |
| R6 | 画像PDF OCR取込(Ghostscript連携)+全ページ化けPDF振替 | 完了・push済み(`5e2a578`) |
| R7 | UI/UX修復5件+応答性2件+司令塔追加是正4件 | **完了・検収済み・push済み(`284c5af`)** |
| R8 | P2P/共有系修正12件(裁定書 `spec_20260731_R8_P2P修正裁定.md`) | **WIPで中断(`3925409`)。API月次上限でエージェント停止**。下記3 |
| R9 | Ghostscript配布・自動検出(要件 `spec_20260731_R9_Ghostscript配布.md`) | **要件定義済み・未着手**。下記4 |

`3925409`(R8 WIP)の状態: lint ERROR 0 / 272テストPASS / 両ビルドOK
=**既存機能は無傷**。ただしR8として**未完成**(新判定式のテスト未追加・
末尾の低優先項目未了・敵対的レビュー未実施)。新モジュール
`src/core/modShareRule.bas`(共有系の判定式を純ロジック分離)は3点登録済み。

## 3. R8の再開方法

裁定書 `docs/dev/spec_20260731_R8_P2P修正裁定.md` の【修正】F1〜F14 のうち、
どこまで実装されたかを **コードで確認**する(エージェントの最終報告は
上限エラーで届いていない)。確認手順:
```bash
git diff 284c5af 3925409 -- src/ | less   # R8 WIP の全変更
```
- `modShareRule.bas` に判定式(OriginKind/AuthorStatKey/ProbeIsReachable/
  ExpiryDecision/StandardSubDirs 等)がある。呼び出し側(modShare/modP2P/
  modChannel/modPack/modGuard)がこれらを正しく使っているか確認。
- **最重要F4(失効タイマー: 未到達端末は消さない)** が
  `modShareRule.ExpiryDecision` + `modGuard` で正しく実装され、
  guard_last_reach が空のとき「消さない」に分岐するかを最優先で確認。
- **F1(感謝状のchannel対応)** の chauth 機構が入ったか。
- 未実装のF項目 + 新ロジックの純ロジックテスト(modTestsPure4/5)追加 +
  敵対的レビュー(読取専用Opus) を、新しいエージェントへ委任して完了させる。
  委任プロンプトは本セッションのR8投入時の型を踏襲(裁定書に従う/
  自己修正禁止・発見は報告/容量制約/検証の排他 を必ず含める)。

## 4. R9(Ghostscript)の進め方

要件書 `docs/dev/spec_20260731_R9_Ghostscript配布.md` が確定済み。
**利用者(PM)の合意事項**: 「別DL・手動配置ゼロ」が絶対要件。
司令塔の裁定: 既定=配布zip一体化+自動検出強化(EDR安全)、
埋め込み(embed-and-drop)はconfig任意ON(既定OFF)。
- 利用者が「多少重くても真の1ファイルがいい」と言えば、埋め込みを既定へ
  切替(その場合は実機1台のEDR隔離テストを先行)。この確認が未了なら
  再開後に一度確認する。

## 5. 実装リソースの制約(重要)

- 本セッション終盤で **Anthropic APIの組織月次上限**に到達し、Opus実装
  エージェントが停止した。再開時、上限がリセット/引上げされていないと
  サブエージェント(Opus/Sonnet)が同じエラーで落ちる可能性がある。
  まず軽い実行(lint等)で通るか確かめ、駄目なら上限リセットを待つ。
- 上限が生きている間は、司令塔(Fable)自身が軽微な安定化
  (WIP保全・lint修正・登録)だけ行い、大きな実装は待つ。

## 6. ユーザーへの未確認・未伝達事項

- **失効タイマーの緊急回避**: 実機configに `knowledge_expire_days=0` を
  追加するようお願い済みだが、実施したか未確認。R8のF4が入るまでの暫定策。
- **Ghostscript配布方式**: R9の既定(zip一体化)で進める旨を伝達済み。
  「真の1ファイル埋め込みを既定にするか」の最終意思は保留(§4)。
- R6完了時に案内した「Ghostscriptフォルダ配置での実機OCR確認」の結果は未報告。

## 7. 参照ファイル一覧(全てpush済み)

```
docs/dev/spec_20260730_R2〜R5 …(完了分)
docs/dev/spec_20260731_R6_画像PDF_OCR取込.md            (完了)
docs/dev/spec_20260731_R7_UIUX修復と応答性.md          (完了)
docs/dev/spec_20260731_R8_P2P修正裁定.md               ← 実装途中(WIP 3925409)
docs/dev/spec_20260731_R9_Ghostscript配布.md           ← 次
docs/dev/reference_公式OCRツール調査_20260731.md
docs/82_実機不具合対応記録_20260729.md
```

## 8. タスクリスト

`TaskList` を確認。未完了の主要タスク: R8完了(#29)/
R9実装/R7-発見2(#30, ウィンドウ幅基準の恒久対策)。無ければ本§2から再作成。
