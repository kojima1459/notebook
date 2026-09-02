#!/usr/bin/env bash
# .bas を触ったら vba_lint を即座に流す(PostToolUse フック)。
#
# 旧版は `cd /home/user/notebook` をベタ書きしていた。これはクラウド実行環境の
# パスで、ユーザーの Mac には存在しない。つまり Mac のデスクトップアプリでは
# このフックは【何もしていなかった】(2026-09-02 発見)。パスは
# CLAUDE_PROJECT_DIR から取り、無ければ git に聞く。
#
# lint を早い段階で当てる理由: LO のテストは、予約語 Enum と衝突する識別子で
# コンパイルエラーではなく120秒ハングする(R33波1)。lint は同じものを ERROR で
# 即座に捕まえるので、LO へ入る前にここで落としたほうが速い。
#
# 掟: 何があっても exit 0。lint の結果は「お知らせ」であって、編集を止める門ではない。
set +e

ROOT="${CLAUDE_PROJECT_DIR:-$(git rev-parse --show-toplevel 2>/dev/null)}"
[ -d "$ROOT" ] || exit 0
cd "$ROOT" || exit 0
[ -f tools/vba_lint.py ] || exit 0

# src/ 配下の .bas/.cls に変更が無ければ黙る(毎回の編集で走らせない)。
git status --porcelain 2>/dev/null | grep -qE 'src/.*\.(bas|cls)$' || exit 0

OUT=$(python3 tools/vba_lint.py --path src 2>&1)
echo "$OUT" | grep -E '^ERROR|^結果:|ERROR: [0-9]+ 件' | tail -3

# ERROR が出ているときだけ、該当行を数行見せる(WARN は容量警告が常時43件あるので出さない)。
if ! echo "$OUT" | grep -q 'ERROR: 0 件'; then
  echo "$OUT" | grep -B1 '  ERROR' | head -12
fi
exit 0
