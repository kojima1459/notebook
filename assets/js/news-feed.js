/* =========================================================
   サンガキッズ - サッカーニュース（自動取得）
   - Googleニュースの検索RSSから、本物のサッカーニュースを自動表示
   - 手動キュレーション不要。サンガ／Jリーグ／海外／海外組／速報
   - CORSプロキシ経由で取得し、取得結果は localStorage にキャッシュ
   ========================================================= */

const FEEDS_FALLBACK = [
  { key: "sanga", label: "💜 京都サンガ", query: "京都サンガ" },
  { key: "jleague", label: "🏆 Jリーグ", query: "Jリーグ" },
  { key: "overseas", label: "🌍 海外サッカー", query: "海外サッカー OR プレミアリーグ OR チャンピオンズリーグ" },
  { key: "japan", label: "🇯🇵 海外組(日本人)", query: "サッカー 日本人選手 海外 OR 日本代表" },
  { key: "flash", label: "⚡ 速報", query: "サッカー 速報" }
];

// 取得できなかったときに案内する、本物のニュースサイト
const SOURCE_LINKS = [
  { name: "Jリーグ公式ニュース", url: "https://www.jleague.jp/news/" },
  { name: "京都サンガF.C. 公式", url: "https://www.sanga-fc.jp/" },
  { name: "京都サンガ（Jリーグ）", url: "https://www.jleague.jp/club/kyoto/" },
  { name: "スポーツナビ サッカー", url: "https://soccer.yahoo.co.jp/" },
  { name: "ゲキサカ", url: "https://web.gekisaka.jp/" }
];

(function setupNewsFeed() {
  const root = document.querySelector("[data-newsfeed]");
  if (!root) return;

  const tabsEl = root.querySelector("[data-feed-tabs]");
  const listEl = root.querySelector("[data-feed-list]");
  const statusEl = root.querySelector("[data-feed-status]");

  let feeds = [];
  let active = null;

  function gnewsUrl(query) {
    const q = encodeURIComponent(`${query} when:14d`);
    return `https://news.google.com/rss/search?q=${q}&hl=ja&gl=JP&ceid=JP:ja`;
  }

  // CORSプロキシ（順番にためす）
  const proxies = [
    (u) => `https://api.allorigins.win/raw?url=${encodeURIComponent(u)}`,
    (u) => `https://api.codetabs.com/v1/proxy/?quest=${encodeURIComponent(u)}`,
    (u) => `https://api.allorigins.win/get?url=${encodeURIComponent(u)}` // JSON .contents
  ];

  async function fetchFeed(query) {
    const target = gnewsUrl(query);
    for (const make of proxies) {
      try {
        const res = await fetch(make(target), { cache: "no-store" });
        if (!res.ok) continue;
        const ctype = res.headers.get("content-type") || "";
        let text;
        if (ctype.includes("application/json")) {
          const j = await res.json();
          text = j.contents || "";
        } else {
          text = await res.text();
        }
        const items = parseRss(text);
        if (items.length) return items;
      } catch (e) { /* つぎのプロキシへ */ }
    }
    return null;
  }

  function parseRss(xmlText) {
    const out = [];
    try {
      const doc = new DOMParser().parseFromString(xmlText, "text/xml");
      doc.querySelectorAll("item").forEach((it) => {
        const titleRaw = (it.querySelector("title")?.textContent || "").trim();
        const link = (it.querySelector("link")?.textContent || "").trim();
        const pub = it.querySelector("pubDate")?.textContent || "";
        const source = it.querySelector("source")?.textContent || "";
        if (!titleRaw || !link) return;
        // Googleニュースのタイトルは "見出し - 媒体名" の形が多い
        let title = titleRaw, src = source;
        const idx = titleRaw.lastIndexOf(" - ");
        if (!src && idx > 0) { title = titleRaw.slice(0, idx); src = titleRaw.slice(idx + 3); }
        out.push({ title, link, src, pub });
      });
    } catch (e) { /* パース失敗 */ }
    return out.slice(0, 15);
  }

  function relTime(pub) {
    if (!pub) return "";
    const t = new Date(pub).getTime();
    if (isNaN(t)) return "";
    const diff = Date.now() - t;
    const h = Math.floor(diff / 3600000);
    if (h < 1) return "さっき";
    if (h < 24) return `${h}時間前`;
    const d = Math.floor(h / 24);
    return `${d}日前`;
  }

  function esc(s) {
    return (s || "").replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  function renderItems(items, fromCache) {
    listEl.innerHTML = items.map((n) => `
      <a class="feed-item" href="${esc(n.link)}" target="_blank" rel="noopener noreferrer">
        <div class="feed-title">${esc(n.title)}</div>
        <div class="feed-meta">${esc(n.src) || "ニュース"} ・ ${relTime(n.pub)}</div>
      </a>`).join("");
    statusEl.innerHTML = fromCache
      ? "📑 ほぞんしてあるニュースを表示中（さいしんは通信がひつようです）"
      : "🟢 さいしんのニュースを じどう取得しました";
  }

  function renderFallback() {
    listEl.innerHTML = `
      <div class="feed-fallback">
        <p>📡 いま ニュースをよみこめませんでした。<br>下の本物のニュースサイトから見てね（あたらしいタブでひらきます）。</p>
        <div class="feed-links">
          ${SOURCE_LINKS.map((s) =>
            `<a href="${s.url}" target="_blank" rel="noopener noreferrer">${s.name} →</a>`).join("")}
        </div>
      </div>`;
    statusEl.textContent = "";
  }

  function cacheKey(key) { return `sangakids_news_${key}`; }

  async function loadFeed(feed) {
    active = feed.key;
    tabsEl.querySelectorAll("button").forEach((b) =>
      b.classList.toggle("active", b.dataset.key === feed.key));
    statusEl.textContent = "⏳ ニュースを取得中…";
    listEl.innerHTML = `<div class="feed-loading">よみこみ中… ⚽</div>`;

    const items = await fetchFeed(feed.query);
    if (active !== feed.key) return; // タブが切りかわっていたら無視

    if (items && items.length) {
      try { localStorage.setItem(cacheKey(feed.key), JSON.stringify({ t: Date.now(), items })); } catch (e) {}
      renderItems(items, false);
      return;
    }
    // 失敗 → キャッシュ → フォールバック
    try {
      const cached = JSON.parse(localStorage.getItem(cacheKey(feed.key)));
      if (cached && cached.items && cached.items.length) { renderItems(cached.items, true); return; }
    } catch (e) {}
    renderFallback();
  }

  function render() {
    tabsEl.innerHTML = feeds.map((f) =>
      `<button type="button" class="feed-tab" data-key="${f.key}">${f.label}</button>`).join("");
    tabsEl.querySelectorAll("button").forEach((b) => {
      b.addEventListener("click", () => {
        const f = feeds.find((x) => x.key === b.dataset.key);
        if (f) loadFeed(f);
      });
    });
    if (feeds[0]) loadFeed(feeds[0]);
  }

  (async function init() {
    try {
      const res = await fetch("data/feeds.json", { cache: "no-store" });
      if (!res.ok) throw new Error("not ok");
      const json = await res.json();
      feeds = json.feeds || FEEDS_FALLBACK;
    } catch (e) {
      feeds = FEEDS_FALLBACK;
    }
    render();
  })();
})();
