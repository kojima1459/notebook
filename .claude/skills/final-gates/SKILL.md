---
name: final-gates
description: 配布前の最終検問と配布物確定の定型シーケンス。ユーザーが「納品して」「配って」「zipちょうだい」と言ったとき、またはレビュー2周消化後のラウンドクローズ時に司令塔が自動で発火する。ラウンドのクローズ時(レビュー2周消化後)に司令塔が実行する。
---

# 最終検問→配布(定型)

全て直列で実行。1つでも落ちたら配布せず原因を潰す。

```bash
python3 tools/vba_lint.py --path src                 # ERROR 0
python3 tools/run_lo_tests.py --mode compile         # OK(絶対に並列にしない)
python3 tools/run_lo_tests.py --mode pure            # 全PASS(同上)
python3 build/build_mybookshelf.py --dev             # 自己検証PASS
python3 build/build_mybookshelf.py --prod            # 自己検証PASS
python3 build/build_mybookshelf.py --prod --zip      # dist/MyBookshelf_配布.zip 再生成
```

その後:
1. `docs/dev/HANDOFF_再開手順.md` を更新(現在地/ラウンド骨子/記録のみ/容量実測/ラウンド履歴表/次報の確認観点)
2. `git add`(dist 3点+HANDOFF+仕様書)→コミット(末尾に Co-Authored-By と Claude-Session の2行)→ `git push -u origin <ブランチ>`(失敗時2s/4s/8s/16sリトライ)
3. SendUserFileで配布zipを送付(キャプション: 旧xlsm削除→展開の注意)
4. クローズ報告: 原因→対策→検証結果→実機での確認手順、を非エンジニア向けの日本語で。途中の実況はしない(成果物で示す)

## 注意

- 配布zipは必ず `--prod --zip` で公式生成(手作業でzipを組まない。R23でGhostscript同梱を壊しかけた実例)
- コミットメッセージ・PR・コードコメントにモデル識別子を入れない
