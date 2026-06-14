/* =========================================================
   サンガキッズ - サッカー動画ゾーン
   - サッカー専門チャンネルの動画を自動でながす（手動キュレーション不要）
   - 各カテゴリ＝チャンネルのアップロード再生リスト or 動画あつめ
   - サッカーだけ。YouTubeの関連動画でジャンクに流れないように rel=0
   ========================================================= */

const CHANNELS_FALLBACK = [
  { key: "skills", label: "⚽ 技・れんしゅう", type: "playlist", id: "UUGJRstFI6eWJTW-DtmgpfGA", desc: "ドリブル・リフティングなど 上たつの動画" },
  { key: "highlights", label: "🏆 Jリーグ ハイライト", type: "playlist", id: "UUyzs0YrgWiL2wdROpajnO1Q", desc: "Jリーグの しあいハイライト" },
  { key: "jfa", label: "🇯🇵 日本代表", type: "playlist", id: "UUgIeUSV91-FfmCayG4lSBcw", desc: "日本代表・JFA の動画" },
  { key: "jleague", label: "📺 Jリーグ公式", type: "playlist", id: "UUWc-XpFHPK1SwGcvpFPZ8NA", desc: "Jリーグ公式チャンネル" },
  { key: "world", label: "🌍 Jリーグ国際", type: "playlist", id: "UUmQp6ZaAejJKKkXc_Y_lh1A", desc: "海外むけ ハイライト・とくしゅう" },
  { key: "collection", label: "🤹 キッズ技あつめ", type: "videos", ids: ["SzZ7Ecql-sg", "CtgXlExS8qc", "I2I8Hq_w7lQ", "ftfP45kKfPw"], desc: "はじめてさん向けの れんしゅう動画" }
];

const VZ_KEY = "sangakids_videozone_tab";

(function setupVideoZone() {
  const root = document.querySelector("[data-videozone]");
  if (!root) return;

  const tabsEl = root.querySelector("[data-vz-tabs]");
  const frame = root.querySelector(".video-frame");
  const descEl = root.querySelector("[data-vz-desc]");

  let cats = [];
  let active = localStorage.getItem(VZ_KEY) || null;

  const base = "https://www.youtube-nocookie.com/embed/";
  const opts = "rel=0&modestbranding=1&playsinline=1";

  function embedUrl(cat) {
    if (cat.type === "playlist") {
      return `${base}videoseries?list=${cat.id}&${opts}`;
    }
    // type: videos（動画ID あつめ）
    const ids = cat.ids || [];
    const first = ids[0];
    const rest = ids.slice(1).join(",");
    return `${base}${first}?${opts}${rest ? "&playlist=" + rest : ""}`;
  }

  function show(cat) {
    active = cat.key;
    localStorage.setItem(VZ_KEY, active);
    frame.innerHTML =
      `<iframe src="${embedUrl(cat)}"
        title="${cat.label}"
        allow="accelerometer; encrypted-media; gyroscope; picture-in-picture; fullscreen"
        allowfullscreen loading="lazy"></iframe>`;
    descEl.textContent = cat.desc || "";
    tabsEl.querySelectorAll("button").forEach((b) =>
      b.classList.toggle("active", b.dataset.key === cat.key)
    );
  }

  function render() {
    tabsEl.innerHTML = cats
      .map((c) => `<button type="button" class="vz-tab" data-key="${c.key}">${c.label}</button>`)
      .join("");
    tabsEl.querySelectorAll("button").forEach((b) => {
      b.addEventListener("click", () => {
        const cat = cats.find((c) => c.key === b.dataset.key);
        if (cat) show(cat);
      });
    });
    const start = cats.find((c) => c.key === active) || cats[0];
    if (start) show(start);
  }

  (async function load() {
    try {
      const res = await fetch("data/channels.json", { cache: "no-store" });
      if (!res.ok) throw new Error("not ok");
      const json = await res.json();
      cats = json.categories || CHANNELS_FALLBACK;
    } catch (e) {
      cats = CHANNELS_FALLBACK;
    }
    render();
  })();
})();
