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
