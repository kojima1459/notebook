#!/usr/bin/env bash
# セッション開始時の現在地表示(SessionStart フック)。
#
# なぜ要るのか: 2026-09-02、ローカルが古いコミットに取り残されたまま
# 「R34/MyBookshelf はこのリポジトリに存在しない」と繰り返し誤報した。
# origin は force-update 済みで実在していた。push が拒否されるまで気付かなかった。
# セッションの1行目でズレが見えていれば起きなかった事故なので、機械に見張らせる。
#
# 掟: 何があっても exit 0。ここで失敗してセッションが始まらないのが最悪。
set +e

ROOT="${CLAUDE_PROJECT_DIR:-$(git rev-parse --show-toplevel 2>/dev/null)}"
[ -d "$ROOT" ] || exit 0
cd "$ROOT" || exit 0
git rev-parse --git-dir >/dev/null 2>&1 || exit 0

BR=$(git branch --show-current 2>/dev/null)
echo "【現在地】ブランチ: ${BR:-(detached)}"

# fetch はネットワーク次第で待たされるので上限を切る。落ちても続行する。
if command -v timeout >/dev/null 2>&1; then
  timeout 20 git fetch origin "$BR" --quiet 2>/dev/null
else
  git fetch origin "$BR" --quiet 2>/dev/null
fi

UP="origin/$BR"
if git rev-parse --verify --quiet "$UP" >/dev/null 2>&1; then
  read -r AHEAD BEHIND <<<"$(git rev-list --left-right --count "HEAD...$UP" 2>/dev/null)"
  if [ "${AHEAD:-0}" = "0" ] && [ "${BEHIND:-0}" = "0" ]; then
    echo "  origin と同期済み"
  elif [ "${AHEAD:-0}" = "0" ]; then
    echo "  ⚠️ ローカルが ${BEHIND} 件遅れています。リポジトリの中身を断定する前に追いついてください"
  elif [ "${BEHIND:-0}" = "0" ]; then
    echo "  未push ${AHEAD} 件"
  else
    echo "  ⚠️ 分岐しています(未push ${AHEAD} / 未取込 ${BEHIND})。勝手に rebase/reset せずユーザーへ差分を出すこと"
  fi
else
  echo "  (origin/$BR が見当たりません)"
fi

DIRTY=$(git status --porcelain 2>/dev/null | wc -l | tr -d ' ')
if [ "${DIRTY:-0}" != "0" ]; then
  echo "  未コミット ${DIRTY} 件:"
  git status --short 2>/dev/null | head -8 | sed 's/^/    /'
  # dist/ は APIキーと社内限資料を含む。コミット対象ではない(CLAUDE.md §13)。
  if git status --porcelain 2>/dev/null | grep -q ' dist/'; then
    echo "    ⚠️ dist/ が見えています。ビルド成果物はコミットしないこと"
  fi
fi

HANDOFF="docs/dev/HANDOFF_再開手順.md"
if [ -f "$HANDOFF" ]; then
  # cut -c は当環境ではバイト単位で切れて日本語が壊れるため、切らずに出す。
  echo "  直近のラウンド: $(grep -m1 '^## 0R' "$HANDOFF" 2>/dev/null | sed 's/^## //')"
fi
echo "  詳しく見るときは orient スキル / 文書の地図は docs/dev/INDEX.md"
exit 0
