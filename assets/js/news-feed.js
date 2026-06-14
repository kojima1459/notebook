/* =========================================================
   サンガキッズ - サッカーニュース（自動取得・本物）
   Vercel関数 /api/news からニュースを取得（サーバー側でGoogleニュースRSSを取得）。
   サンガ／Jリーグ／海外サッカー／海外組／速報をタブで切替。
   ========================================================= */

const FEEDS = [
  { key: "sanga",     label: "💜 京都サンガ" },
  { key: "jleague",   label: "🏆 Jリーグ" },
  { key: "official",  label: "📢 Jリーグ公式" },
  { key: "overseas",  label: "🌍 海外サッカー" },
  { key: "japan",     label: "🇯🇵 海外組" },
  { key: "flash",     label: "⚡ 速報" },
  { key: "koukou",    label: "🏫 高校サッカー" },
  { key: "transfer",  label: "⭐ 移籍情報" },
  { key: "champions", label: "🌟 欧州サッカー" },
  { key: "junior",    label: "👦 ジュニア" },
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

  const PAGE = 8;            // さいしょに出す件数／「もっと見る」で足す件数
  let currentItems = [];     // いま開いているタブの全ニュース
  let shown = 0;             // いま表示している件数

  function relTime(pub) {
    if (!pub) return "";
    const t = new Date(pub).getTime();
    if (isNaN(t)) return "";
    const h = Math.floor((Date.now() - t) / 3600000);
    if (h < 1) return "さっき";
    if (h < 24) return `${h}時間前`;
    return `${Math.floor(h / 24)}日前`;
  }

  function itemHtml(n) {
    return `
      <a class="feed-item" href="${n.link}" target="_blank" rel="noopener noreferrer">
        <div class="feed-title">${n.title}</div>
        <div class="feed-meta">${n.source || "ニュース"} ・ ${relTime(n.pub)}</div>
      </a>`;
  }

  // いま表示している件数ぶんを描画し、続きがあれば「もっと見る」ボタンを付ける
  function renderList() {
    const slice = currentItems.slice(0, shown);
    let html = slice.map(itemHtml).join("");
    const rest = currentItems.length - shown;
    if (rest > 0) {
      html += `<button type="button" class="feed-more" data-feed-more>
        ＋ もっと見る（あと${rest}件）</button>`;
    }
    listEl.innerHTML = html;
    const moreBtn = listEl.querySelector("[data-feed-more]");
    if (moreBtn) moreBtn.addEventListener("click", () => {
      shown = Math.min(shown + PAGE, currentItems.length);
      renderList();
    });
    statusEl.textContent = `🟢 最新ニュース ${slice.length}／${currentItems.length}件`;
  }

  function renderItems(items) {
    currentItems = items;
    shown = Math.min(PAGE, items.length);
    renderList();
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
