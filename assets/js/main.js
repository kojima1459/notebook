/* =========================================================
   サンガキッズ - メインスクリプト
   モバイルナビ・ニュース表示・年号の自動更新など
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

  // リンクをタップしたら閉じる（モバイル）
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

// ---------- ニュースデータ（オフラインでも動く埋め込みデータ） ----------
const NEWS_FALLBACK = [
  { category: "⚽", tag: "しあい", date: "2026-06-13",
    title: "ホームで大かんせい！サンガが元気にプレー",
    summary: "サンガスタジアム by KYOCERA にたくさんのファンがあつまったよ。みんなで紫のタオルマフラーをふって応援しよう！" },
  { category: "🌟", tag: "せんしゅ", date: "2026-06-10",
    title: "ゴールをきめた選手にインタビュー",
    summary: "「みんなの応援がいちばんの力です！」とニッコリ。次のしあいも楽しみだね。" },
  { category: "🎉", tag: "イベント", date: "2026-06-07",
    title: "キッズサッカー教室がひらかれたよ",
    summary: "ボールのけり方やドリブルを楽しく練習。マスコットもいっしょにあそんでくれたよ！" },
  { category: "🏟️", tag: "スタジアム", date: "2026-06-03",
    title: "サンガスタジアムってどんなところ？",
    summary: "選手とピッチがとっても近いスタジアム。サッカーのめいろやひろばもあって一日たのしめるよ。" },
  { category: "💜", tag: "おうえん", date: "2026-05-30",
    title: "みんなで作った応援メッセージ",
    summary: "子どもたちのかいた応援ボードがスタジアムにとうじょう。選手たちもパワーをもらったみたい！" },
  { category: "🎽", tag: "グッズ", date: "2026-05-25",
    title: "あたらしいキッズユニフォームが登場",
    summary: "サイズもいろいろ。お気に入りの番号で、きみもサンガの一員になろう！" }
];

function renderNews(list, limit) {
  const container = document.querySelector("[data-news]");
  if (!container) return;
  const items = limit ? list.slice(0, limit) : list;
  container.innerHTML = items
    .map(
      (n) => `
      <article class="news-item">
        <div class="cat" aria-hidden="true">${n.category}</div>
        <div class="body">
          <div class="meta">${formatDate(n.date)}<span class="tag">${n.tag}</span></div>
          <h3>${escapeHtml(n.title)}</h3>
          <p>${escapeHtml(n.summary)}</p>
        </div>
      </article>`
    )
    .join("");
}

function formatDate(iso) {
  const [y, m, d] = iso.split("-");
  return `${y}年${Number(m)}月${Number(d)}日`;
}

function escapeHtml(s) {
  return s.replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
  );
}

// JSONを読みにいき、失敗したら埋め込みデータを使う
(async function loadNews() {
  const container = document.querySelector("[data-news]");
  if (!container) return;
  const limit = container.dataset.news === "home" ? 3 : null;
  try {
    const res = await fetch("data/news.json", { cache: "no-store" });
    if (!res.ok) throw new Error("not ok");
    const json = await res.json();
    renderNews(json.news || NEWS_FALLBACK, limit);
  } catch (e) {
    renderNews(NEWS_FALLBACK, limit);
  }
})();
