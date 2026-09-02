---
name: orient
description: セッション開始時・再開時の現在地確認。ユーザーが「現在地」「いまどこ」「再開」「続きから」「状況教えて」「何やってたっけ」「前回の続き」と言ったとき、またはセッションの最初の実作業に入る前に司令塔が自動で発火する。ローカルとoriginのズレ・未コミット・検問の状態・HANDOFFの現在地を3分で確定させ、「無い」「存在しない」と断言する前の探索範囲を固める。
---

# 現在地確認（着手前の3分）

**このスキルを通さずに「〜は無い」「〜は存在しない」「〜は直っている」と断言してはならない。**

2026-09-02のセッションで、ローカルが古いコミットに取り残されたまま
「R34/MyBookshelf はこのリポジトリに存在しない」とユーザーと別セッションへ
**繰り返し誤報**した。origin は force-update 済みで、実際には存在していた。
`git push` が拒否されるまで気付かなかった。ここは毎回・機械的に通す。

## 手順1: リポジトリの現在地（必ず最初）

```bash
git fetch origin claude/internal-notebook-lm-chatbot-B6BE7 2>&1 | tail -2
git status --short | head -20
git log --oneline -3
git log --oneline -3 origin/claude/internal-notebook-lm-chatbot-B6BE7
git rev-list --left-right --count HEAD...origin/claude/internal-notebook-lm-chatbot-B6BE7
```

判定:

| `--left-right --count` の出力 | 意味 | やること |
|---|---|---|
| `0 0` | 同期済み | そのまま進む |
| `0 N`（右が正） | **ローカルが古い** | `git reset --hard origin/...` の前に、ローカル固有の変更が無いか `git status` と `git stash list` を確認してから追いつく |
| `N 0` | 未push | 内容を確認してから push |
| `N M` | **分岐している** | **勝手に rebase / reset しない。** ユーザーへ両側の差分を提示して選ばせる |

> **`git rebase` / `git reset --hard` を回す前に必ず `git diff` で失うものを見る。**
> 2026-09-02、rebase が `.gitignore` の育ててきたルール（`*.building.xlsm`・
> `発行者用`・`seed/`）を丸ごと巻き戻す寸前だった。差分を見て abort し、
> 追記だけの外科手術に切り替えて助かった。

## 手順2: 検問の現在地

```bash
python3 tools/vba_lint.py --path src 2>&1 | tail -4
```

テスト（`run_lo_tests.py`）は**約3〜4分かかる**。着手前に流すなら
バックグラウンドで走らせ、その間に読み物・調査を進める（**直列は
run_lo_tests 同士の話であって、調査や文書作業との並行は禁じていない**）:

```bash
nohup bash -c 'python3 tools/run_lo_tests.py --mode compile > /tmp/lo_c.log 2>&1 \
  && python3 tools/run_lo_tests.py --mode pure > /tmp/lo_p.log 2>&1; \
  echo DONE_$? > /tmp/lo_done.txt' >/dev/null 2>&1 &
```

## 手順3: 仕事の現在地

1. `docs/dev/HANDOFF_再開手順.md` の **冒頭の最新ラウンド節**を読む
   （§1「現在地」は過去ラウンドの記述が残っていることがある。
   **節番号ではなく日付・ラウンド番号が最新のものを正とする**）
2. `docs/dev/INDEX.md` で、今回の話題がどの文書の管轄かを確認する
3. 未消化の「記録のみ」「残タスク」を拾う

## 手順4: ユーザーへ1画面で報告

```
【現在地】
・ブランチ: <名前>／origin と <同期済み / N件遅れ / 分岐>
・作業ツリー: <クリーン / 未コミットN件（内訳）>
・直近のラウンド: <R34 など>・状態: <完了 / 実機報待ち>
・検問: lint <結果>／テスト <実行中 / 前回R34時点でPASS N件>
・未消化: <記録のみ・残タスクから重要なものだけ2-3件>
```

## 「無い」と言うときの作法

**探索範囲を書かずに「ゼロ」「無い」「存在しない」と言ってはならない。**

2026-09-02、キー混入調査で「全45コミットでゼロ」と断言したが、
実際には調査対象から外していたブランチに存在した。
**範囲を書かなかったことが誤りの本体**で、調査自体は間違っていなかった。

```
✅ 「main と B6BE7 の全45コミットを対象に調べた範囲では見つからなかった。
    product/nexus-agent は未調査」
❌ 「全コミットでゼロ」
```
