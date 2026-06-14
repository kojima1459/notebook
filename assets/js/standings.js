/* =========================================================
   サンガキッズ - J1順位表（本物・自動取得）
   /api/standings（Vercel関数）からJ1リーグの順位表を取得して表示。
   京都サンガの行をハイライト。サンガの最新ニュースも表示。
   ========================================================= */

const OFFICIAL_STANDINGS = "https://www.jleague.jp/standings/j1/";

function formBadges(form) {
  if (!form || !form.length) return "";
  return (
    '<span class="form">' +
    form
      .map((f) => {
        const cls = f === "W" ? "w" : f === "L" ? "l" : "d";
        return `<i class="fb ${cls}">${f}</i>`;
      })
      .join("") +
    "</span>"
  );
}

function renderGroup(group) {
  const rows = group.rows
    .map((r) => {
      const cls = r.isSanga ? "stand-row sanga" : "stand-row";
      return `
      <tr class="${cls}">
        <td class="c-rank">${r.rank ?? "-"}</td>
        <td class="c-team">${r.isSanga ? "💜 " : ""}${r.team}</td>
        <td>${r.played ?? "-"}</td>
        <td class="c-pts">${r.points ?? "-"}</td>
        <td>${r.goalDiff > 0 ? "+" + r.goalDiff : r.goalDiff ?? "-"}</td>
      </tr>`;
    })
    .join("");

  return `
    <div class="stand-card">
      <h3 class="stand-group">${group.name === "EAST" ? "🟦 EASTグループ" : group.name === "WEST" ? "🟥 WESTグループ" : "J1リーグ"}</h3>
      <table class="stand-table">
        <thead>
          <tr><th>順位</th><th>クラブ</th><th>試合</th><th>勝点</th><th>得失</th></tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

function renderSangaSummary(groups) {
  let sanga = null, groupName = "";
  groups.forEach((g) => {
    const s = g.rows.find((r) => r.isSanga);
    if (s) { sanga = s; groupName = g.name; }
  });
  if (!sanga) return "";
  return `
    <div class="sanga-summary">
      <div class="ss-rank"><span>${groupName}</span><b>${sanga.rank}</b>位</div>
      <div class="ss-stats">
        <div><b>${sanga.points}</b><span>勝点</span></div>
        <div><b>${sanga.win}</b><span>勝</span></div>
        <div><b>${sanga.lose}</b><span>負</span></div>
        <div><b>${sanga.goalsFor}-${sanga.goalsAgainst}</b><span>得失点</span></div>
      </div>
    </div>`;
}

async function loadStandings() {
  const wrap = document.querySelector("[data-standings]");
  if (!wrap) return;
  wrap.innerHTML = `<div class="feed-loading">順位表をよみこみ中… ⚽</div>`;
  try {
    const res = await fetch("/api/standings", { cache: "no-store" });
    const data = await res.json();
    if (!data.ok || !data.groups.length) throw new Error("no data");
    const summary = document.querySelector("[data-sanga-summary]");
    if (summary) summary.innerHTML = renderSangaSummary(data.groups);
    wrap.innerHTML = data.groups.map(renderGroup).join("");
    const upd = document.querySelector("[data-standings-updated]");
    if (upd) {
      const d = new Date(data.updated);
      upd.textContent = `🟢 ${d.getMonth() + 1}月${d.getDate()}日 ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")} 取得`;
    }
  } catch (e) {
    wrap.innerHTML = `
      <div class="feed-fallback">
        <p>📡 順位表をよみこめませんでした。<br>公式サイトでチェックしてね。</p>
        <div class="feed-links"><a href="${OFFICIAL_STANDINGS}" target="_blank" rel="noopener noreferrer">Jリーグ公式 順位表 →</a></div>
      </div>`;
  }
}

// サンガの最新ニュース（試合結果・日程ふくむ）
async function loadSangaNews() {
  const el = document.querySelector("[data-sanga-news]");
  if (!el) return;
  try {
    const res = await fetch("/api/news?cat=sanga", { cache: "no-store" });
    const data = await res.json();
    if (!data.items || !data.items.length) throw new Error("none");
    el.innerHTML = data.items.slice(0, 6).map((n) => `
      <a class="feed-item" href="${n.link}" target="_blank" rel="noopener noreferrer">
        <div class="feed-title">${n.title}</div>
        <div class="feed-meta">${n.source || "ニュース"}</div>
      </a>`).join("");
  } catch (e) {
    el.innerHTML = `<div class="feed-links"><a href="https://www.jleague.jp/club/kyoto/" target="_blank" rel="noopener noreferrer">京都サンガ（Jリーグ公式）→</a></div>`;
  }
}

// サンガの次の試合・直近の結果・クラブプロフィール
async function loadSangaInfo() {
  const matchEl = document.querySelector("[data-sanga-match]");
  const profEl = document.querySelector("[data-sanga-profile]");
  try {
    const res = await fetch("/api/sanga", { cache: "no-store" });
    const data = await res.json();

    if (matchEl) {
      const cards = [];
      if (data.nextMatch) {
        const n = data.nextMatch;
        cards.push(`
          <a class="match-info next" ${n.link ? `href="${n.link}" target="_blank" rel="noopener noreferrer"` : ""}>
            <div class="mi-label">📅 次の試合</div>
            <div class="mi-teams">京都サンガ <span class="mi-vs">VS</span> ${n.opponent}</div>
            <div class="mi-date">${n.year}年${n.dateLabel}${n.round ? "・第" + n.round + "節" : ""}</div>
          </a>`);
      }
      if (data.lastResult) {
        const r = data.lastResult;
        cards.push(`
          <a class="match-info last" ${r.link ? `href="${r.link}" target="_blank" rel="noopener noreferrer"` : ""}>
            <div class="mi-label">📝 直近の試合</div>
            <div class="mi-teams">京都サンガ <span class="mi-vs">VS</span> ${r.opponent}</div>
            <div class="mi-date">${r.dateLabel}${r.round ? "・第" + r.round + "節" : ""}・結果は記事でチェック →</div>
          </a>`);
      }
      if (!cards.length) {
        cards.push(`
          <div class="match-info none">
            <div class="mi-label">📅 次の試合</div>
            <div class="mi-date">日程が決まると、ここに表示されるよ。<br>
            <a href="https://www.jleague.jp/club/kyoto/day/" target="_blank" rel="noopener noreferrer">公式の日程・結果を見る →</a></div>
          </div>`);
      }
      matchEl.innerHTML = cards.join("");
    }

    if (profEl && data.profile) {
      const p = data.profile;
      profEl.innerHTML = `
        <div class="profile-card">
          <h3>💜 ${p.name}</h3>
          <ul class="profile-list">
            <li><b>ホームタウン</b>${p.hometown}</li>
            <li><b>スタジアム</b>${p.stadium}</li>
            <li><b>クラブカラー</b>${p.color}</li>
            <li><b>Jリーグ</b>${p.joinedJ}</li>
            <li><b>マスコット</b>${p.mascot}</li>
          </ul>
          <p class="profile-note">${p.note}</p>
        </div>`;
    }
  } catch (e) {
    if (matchEl) matchEl.innerHTML = `<div class="feed-links"><a href="https://www.jleague.jp/club/kyoto/day/" target="_blank" rel="noopener noreferrer">公式の日程・結果を見る →</a></div>`;
  }
}

// 選手名鑑（主な選手）
async function loadRoster() {
  const el = document.querySelector("[data-roster]");
  if (!el) return;
  el.innerHTML = `<div class="feed-loading">選手をよみこみ中… ⚽</div>`;
  try {
    const res = await fetch("/api/roster", { cache: "no-store" });
    const data = await res.json();
    if (!data.ok || !data.groups.length) throw new Error("none");
    const POS_JP = { GK: "GK ゴールキーパー", DF: "DF ディフェンダー", MF: "MF ミッドフィルダー", FW: "FW フォワード" };
    el.innerHTML = data.groups.map((g) => `
      <div class="roster-group">
        <div class="roster-pos">${POS_JP[g.pos] || g.pos}</div>
        <div class="roster-grid">
          ${g.players.map((p) => `
            <div class="player-card">
              <div class="player-num">${p.number || "–"}</div>
              <div class="player-info">
                <div class="player-name">${p.flag} ${p.name}</div>
                <div class="player-nat">${p.nationality || ""}</div>
              </div>
            </div>`).join("")}
        </div>
      </div>`).join("");
  } catch (e) {
    el.innerHTML = `<div class="feed-links"><a href="https://www.jleague.jp/club/kyoto/player/" target="_blank" rel="noopener noreferrer">公式の選手一覧を見る →</a></div>`;
  }
}

loadStandings();
loadSangaInfo();
loadSangaNews();
loadRoster();
