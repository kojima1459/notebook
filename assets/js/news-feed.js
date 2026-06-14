/* =========================================================
   サンガキッズ - サッカーニュース（自動取得・本物）
   Vercel関数 /api/news からニュースを取得（サーバー側でGoogleニュースRSSを取得）。
   サンガ／Jリーグ／海外サッカー／海外組／速報をタブで切替。
   ========================================================= */

const FEEDS = [
  { key: "sanga", label: "💜 京都サンガ" },
  { key: "jleague", label: "🏆 Jリーグ" },
  { key: "official", label: "📢 Jリーグ公式" },
  { key: "overseas", label: "🌍 海外サッカー" },
  { key: "japan", label: "🇯🇵 海外組" },
  { key: "flash", label: "⚡ 速報" },
];

const NEWS_SOURCE_LINKS = [
  { name: "Jリーグ公式ニュース", url: "https://www.jleague.jp/news/" },
  { name: "京都サンガF.C. 公式", url: "https://www.sanga-fc.jp/" },
  { name: "スポーツナビ サッカー", url: "https://soccer.yahoo.co.jp/" },
];

(function setupNewsFeed() {
  const root = document.querySelector("[data-newsfeed]");
  if (!root) return;
  const tabsEl = root.querySelector("[data-feed-tabs]");
  const listEl = root.querySelector("[data-feed-list]");
  const statusEl = root.querySelector("[data-feed-status]");
  let active = null;

  function relTime(pub) {
    if (!pub) return "";
    const t = new Date(pub).getTime();
    if (isNaN(t)) return "";
    const h = Math.floor((Date.now() - t) / 3600000);
    if (h < 1) return "さっき";
    if (h < 24) return `${h}時間前`;
    return `${Math.floor(h / 24)}日前`;
  }

  function renderItems(items) {
    listEl.innerHTML = items.map((n) => `
      <a class="feed-item" href="${n.link}" target="_blank" rel="noopener noreferrer">
        <div class="feed-title">${n.title}</div>
        <div class="feed-meta">${n.source || "ニュース"} ・ ${relTime(n.pub)}</div>
      </a>`).join("");
    statusEl.textContent = "🟢 最新ニュースを自動取得しました";
  }

  function renderFallback() {
    listEl.innerHTML = `
      <div class="feed-fallback">
        <p>📡 いまニュースを読み込めませんでした。下のサイトでチェックしてね。</p>
        <div class="feed-links">
          ${NEWS_SOURCE_LINKS.map((s) => `<a href="${s.url}" target="_blank" rel="noopener noreferrer">${s.name} →</a>`).join("")}
        </div>
      </div>`;
    statusEl.textContent = "";
  }

  async function loadFeed(feed) {
    active = feed.key;
    tabsEl.querySelectorAll("button").forEach((b) => b.classList.toggle("active", b.dataset.key === feed.key));
    statusEl.textContent = "⏳ 取得中…";
    listEl.innerHTML = `<div class="feed-loading">よみこみ中… ⚽</div>`;
    try {
      const res = await fetch(`/api/news?cat=${feed.key}`, { cache: "no-store" });
      const data = await res.json();
      if (active !== feed.key) return;
      if (data.items && data.items.length) renderItems(data.items);
      else renderFallback();
    } catch (e) {
      if (active === feed.key) renderFallback();
    }
  }

  tabsEl.innerHTML = FEEDS.map((f) => `<button type="button" class="feed-tab" data-key="${f.key}">${f.label}</button>`).join("");
  tabsEl.querySelectorAll("button").forEach((b) => {
    b.addEventListener("click", () => {
      const f = FEEDS.find((x) => x.key === b.dataset.key);
      if (f) loadFeed(f);
    });
  });
  loadFeed(FEEDS[0]);
})();
