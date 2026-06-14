/* =========================================================
   サンガキッズ - サッカーショート（サイト内 動画プレーヤー）
   - おうちの人がえらんだサッカー動画だけを表示
   - 無限スクロールではなく「○本中○本目」の有げんリスト
   ========================================================= */

// オフラインでも動く埋め込みリスト（data/shorts.json と同じ内容）
const SHORTS_FALLBACK = [
  { id: "SzZ7Ecql-sg", title: "リフティングのきほん（しょしんしゃ・キッズむけ）", category: "技れんしゅう" },
  { id: "CtgXlExS8qc", title: "リフティング やり方（はじめてさん）", category: "技れんしゅう" },
  { id: "I2I8Hq_w7lQ", title: "ボールのあつかい方・ジャグリング入もん", category: "技れんしゅう" },
  { id: "ftfP45kKfPw", title: "かんたんトリック 8つにちょうせん", category: "技れんしゅう" }
];

(function setupShorts() {
  const root = document.querySelector("[data-shorts]");
  if (!root) return;

  const frame = root.querySelector(".shorts-frame");
  const titleEl = root.querySelector(".shorts-title");
  const countEl = root.querySelector(".shorts-count");
  const catEl = root.querySelector(".shorts-cat");
  const prevBtn = root.querySelector("[data-shorts-prev]");
  const nextBtn = root.querySelector("[data-shorts-next]");
  const emptyEl = root.querySelector(".shorts-empty");
  const stageEl = root.querySelector(".shorts-stage");

  let list = [];
  let i = 0;

  function isValidId(id) {
    return typeof id === "string" && /^[A-Za-z0-9_-]{6,20}$/.test(id);
  }

  function render() {
    const item = list[i];
    // youtube-nocookie（プライバシー強化）＋ rel=0 で関連動画をへらす
    frame.innerHTML =
      `<iframe src="https://www.youtube-nocookie.com/embed/${item.id}?rel=0&modestbranding=1"
        title="${item.title}"
        allow="accelerometer; encrypted-media; gyroscope; picture-in-picture"
        allowfullscreen
        loading="lazy"></iframe>`;
    titleEl.textContent = item.title;
    catEl.textContent = item.category || "サッカー";
    countEl.textContent = `${list.length}本中 ${i + 1}本目`;
    prevBtn.disabled = i === 0;
    nextBtn.disabled = i === list.length - 1;
  }

  function showEmpty() {
    stageEl.hidden = true;
    emptyEl.hidden = false;
  }

  function start(data) {
    list = (data || []).filter((s) => s && isValidId(s.id));
    if (!list.length) return showEmpty();
    emptyEl.hidden = true;
    stageEl.hidden = false;
    render();
  }

  prevBtn.addEventListener("click", () => { if (i > 0) { i--; render(); } });
  nextBtn.addEventListener("click", () => { if (i < list.length - 1) { i++; render(); } });

  (async function load() {
    try {
      const res = await fetch("data/shorts.json", { cache: "no-store" });
      if (!res.ok) throw new Error("not ok");
      const json = await res.json();
      start(json.shorts || SHORTS_FALLBACK);
    } catch (e) {
      start(SHORTS_FALLBACK);
    }
  })();
})();
