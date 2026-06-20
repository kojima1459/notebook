/* =========================================================
   サンガキッズ - メインスクリプト
   モバイルナビ・年号・ホームのライブニュース見出し
   ========================================================= */

// ---------- モバイルナビの開閉 ----------
(function setupNav() {
  const toggle = document.querySelector(".nav-toggle");
  const nav = document.querySelector(".main-nav");
  if (!toggle || !nav) return;

  toggle.addEventListener("click", () => {
    const open = nav.classList.toggle("open");
    toggle.setAttribute("aria-expanded", open ? "true" : "false");
    toggle.textContent = open ? "✕" : "☰";
  });

  nav.querySelectorAll("a").forEach((a) => {
    a.addEventListener("click", () => {
      nav.classList.remove("open");
      toggle.setAttribute("aria-expanded", "false");
      toggle.textContent = "☰";
    });
  });
})();

// ---------- フッターの年号 ----------
(function setYear() {
  const el = document.querySelector("[data-year]");
  if (el) el.textContent = new Date().getFullYear();
})();

// ---------- ホームの「つぎの試合」（本物の日程を /api/sanga から）----------
(async function homeNextMatch() {
  const el = document.querySelector("[data-home-match]");
  if (!el) return;
  try {
    const res = await fetch("/api/sanga");
    const data = await res.json();
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
    el.innerHTML = cards.join("");
  } catch (e) {
    el.innerHTML = `
      <div class="match-info none">
        <div class="mi-label">📅 次の試合</div>
        <div class="mi-date"><a href="https://www.jleague.jp/club/kyoto/day/" target="_blank" rel="noopener noreferrer">公式の日程・結果を見る →</a></div>
      </div>`;
  }
})();

// ---------- ホームのライブニュース見出し（京都サンガ） ----------
(async function liveNews() {
  const el = document.querySelector("[data-livenews]");
  if (!el) return;
  const cat = el.dataset.livenews || "sanga";
  try {
    const res = await fetch(`/api/news?cat=${cat}`, { cache: "no-store" });
    const data = await res.json();
    if (!data.items || !data.items.length) throw new Error("none");
    el.innerHTML = data.items.slice(0, 4).map((n) => `
      <a class="feed-item" href="${n.link}" target="_blank" rel="noopener noreferrer">
        <div class="feed-title">${n.title}</div>
        <div class="feed-meta">${n.source || "ニュース"}</div>
      </a>`).join("");
  } catch (e) {
    el.innerHTML = `<div class="feed-links"><a href="news.html">ニュースページを見る →</a></div>`;
  }
})();
