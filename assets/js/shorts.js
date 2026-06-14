/* =========================================================
   サンガキッズ - サッカー動画ゾーン
   - サッカー専門チャンネルの動画を自動再生（手動キュレーション不要）
   - YouTube IFrame APIで「いま流れている動画そのもの」を保存し、
     あとからサイト内で見返せる（YouTubeサイトには飛ばさない）
   - サッカー以外へ流れないよう rel=0
   ========================================================= */

const CHANNELS_FALLBACK = [
  { key: "sanga",      label: "💜 京都サンガ",       type: "playlist", id: "UUThC5l5XIVUEljKnCTMgI8A", desc: "京都サンガ公式チャンネルの最新動画" },
  { key: "skills",     label: "⚽ 技・れんしゅう",   type: "playlist", id: "UUGJRstFI6eWJTW-DtmgpfGA", desc: "ドリブル・リフティングなど 上たつの動画" },
  { key: "highlights", label: "🏆 Jリーグ ハイライト",type: "playlist", id: "UUyzs0YrgWiL2wdROpajnO1Q", desc: "Jリーグの しあいハイライト" },
  { key: "jfa",        label: "🇯🇵 日本代表",        type: "playlist", id: "UUgIeUSV91-FfmCayG4lSBcw", desc: "日本代表・JFA の動画" },
  { key: "jleague",    label: "📺 Jリーグ公式",      type: "playlist", id: "UUWc-XpFHPK1SwGcvpFPZ8NA", desc: "Jリーグ公式チャンネル" },
  { key: "world",      label: "🌍 Jリーグ国際",      type: "playlist", id: "UUmQp6ZaAejJKKkXc_Y_lh1A", desc: "海外むけ ハイライト・とくしゅう" },
  { key: "premier",    label: "🏴 プレミアリーグ",   type: "playlist", id: "UUqZQlzSHbVJrwrn5XvzrzcA", desc: "イングランドのトップリーグ。ハイライトと得点集" },
  { key: "champions",  label: "🌟 チャンピオンズL",  type: "playlist", id: "UUJ2PKPB0Lnz19sCioNSGJ9Q", desc: "ヨーロッパ最高峰・UEFAチャンピオンズリーグの動画" },
  { key: "laliga",     label: "🇪🇸 ラ・リーガ",     type: "playlist", id: "UUcwdO2MRiKEFRIWCELh4b3A", desc: "スペインの一部リーグ。メッシやクリロナのリーグ" },
  { key: "bundesliga", label: "🇩🇪 ブンデスリーガ", type: "playlist", id: "UUGNon7f4RLTkHLVQqXUMI_Q", desc: "ドイツの一部リーグ。スピードと激しさが魅力！" },
  { key: "super",      label: "⚡ スーパープレー",   type: "videos",   ids: ["MmBj0gFsUaE","bJSpOIRuSn4","tYEW7nhXvVs","UHQCe-iMcSc","LNK4hPKqPmo"], desc: "世界の選手のすごいプレー集。一時停止してまねしてみよう！" },
  { key: "collection", label: "🤹 キッズ技あつめ",  type: "videos",   ids: ["SzZ7Ecql-sg","CtgXlExS8qc","I2I8Hq_w7lQ","ftfP45kKfPw"], desc: "はじめてさん向けの れんしゅう動画" }
];

const VZ_KEY = "sangakids_videozone_tab";
const VZ_FAV = "sangakids_videozone_favs";
const VZ_SAVED = "sangakids_saved_videos";

// カテゴリ別 練習おすすめ（動画→練習の動線）
const PRACTICE_SUGGESTIONS = {
  sanga:      { icon: "🎯", name: "シュート",                 text: "サンガ選手みたいに正確なシュートを練習しよう" },
  skills:     { icon: "🤹", name: "リフティング",             text: "テクニックの土台！まずはリフティングから" },
  highlights: { icon: "👟", name: "シザース（フェイント）",   text: "ハイライトみたいな1対1、フェイントで抜こう" },
  jfa:        { icon: "🏃", name: "ドリブル",                 text: "代表選手みたいに速いドリブルを身につけよう" },
  jleague:    { icon: "🎯", name: "シュート",                 text: "Jリーグ選手みたいな強いシュートにちょうせん" },
  world:      { icon: "🌀", name: "ルーレット",               text: "海外の選手みたいなかっこいい技を覚えよう" },
  premier:    { icon: "💨", name: "ドリブルダッシュ",         text: "プレミアの選手みたいにスピードで抜こう！" },
  champions:  { icon: "🌟", name: "シュート精度",             text: "チャンピオンズリーグの決定力！コースを狙おう" },
  laliga:     { icon: "🌀", name: "ルーレット",               text: "ラ・リーガのテクニカルな技、まねしてみよう" },
  bundesliga: { icon: "🏃", name: "スプリント＆ターン",       text: "ブンデスのスピードプレー、切り返しを練習しよう" },
  super:      { icon: "⚡", name: "スーパープレーのコピー",   text: "見た技を一時停止して、体で覚えてみよう！" },
  collection: { icon: "🤹", name: "リフティング",             text: "キッズ技動画を見たら、自分でも挑戦してみよう" },
  saved:      { icon: "⭐", name: "気になった技",             text: "保存した動画の技、体で覚えよう" },
};

(function setupVideoZone() {
  const root = document.querySelector("[data-videozone]");
  if (!root) return;

  const tabsEl = root.querySelector("[data-vz-tabs]");
  const descEl = root.querySelector("[data-vz-desc]");
  const favBar = root.querySelector("[data-vz-fav]");
  const starBtn = root.querySelector("[data-vz-star]");
  const saveBtn = root.querySelector("[data-vz-save]");
  const savedMgr = root.querySelector("[data-vz-saved]");

  let cats = [];
  let active = localStorage.getItem(VZ_KEY) || null;
  let favs = read(VZ_FAV, []);
  let saved = read(VZ_SAVED, []); // [{id, title}]
  let player = null, apiReady = false, pending = null, currentVideo = null;
  let errorSkips = 0; // 再生エラー（埋め込み不可など）の連続スキップ数

  // 視聴時間トラッキング（セッション用）
  let playTimer = null, totalPlaySec = 0, lastNudgeSec = -999, timeNudgeDone = false;

  // ---- 1日の視聴上限（1時間 = 3600秒）----
  const WATCH_KEY = "sangakids_watch_v1";
  const WATCH_LIMIT = 3600;  // 1時間
  const BREAK_SEC   = 600;   // 休憩タイマー 10分
  let lockInterval  = null;

  function loadWatchData() {
    try {
      const d = JSON.parse(localStorage.getItem(WATCH_KEY)) || {};
      if (d.date !== new Date().toDateString()) return { date: new Date().toDateString(), total: 0, unlockedUntil: 0 };
      return d;
    } catch (e) { return { date: new Date().toDateString(), total: 0, unlockedUntil: 0 }; }
  }
  function saveWatchData(d) { try { localStorage.setItem(WATCH_KEY, JSON.stringify(d)); } catch (e) {} }

  let watchData = loadWatchData();

  function read(k, d) { try { return JSON.parse(localStorage.getItem(k)) || d; } catch (e) { return d; } }
  function write(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }

  // ---- 練習への誘導（視聴後ナッジ）& 1時間制限 ----
  function startPlayTimer() {
    if (playTimer) return;
    playTimer = setInterval(() => {
      totalPlaySec++;
      watchData.total++;
      // 30秒ごとに保存（localStorage 負荷軽減）
      if (watchData.total % 30 === 0) saveWatchData(watchData);

      const el = root.querySelector("[data-vz-playtime]");
      if (el) {
        const m = Math.floor(totalPlaySec / 60);
        const s = totalPlaySec % 60;
        el.textContent = `⏱ 視聴 ${m}分${s < 10 ? "0" : ""}${s}秒`;
      }
      if (!timeNudgeDone && totalPlaySec >= 300) {
        timeNudgeDone = true;
        showNudge("time");
      }
      // 1時間チェック
      if (watchData.total >= WATCH_LIMIT && !document.querySelector(".vz-lock-overlay:not([hidden])")) {
        triggerLock();
      }
    }, 1000);
  }

  function triggerLock() {
    stopPlayTimer();
    saveWatchData(watchData);
    if (player && player.pauseVideo) try { player.pauseVideo(); } catch (e) {}
    const overlay = document.querySelector(".vz-lock-overlay");
    if (!overlay) return;
    overlay.hidden = false;

    const drills = [
      "🤹 リフティング：下からやさしく ポンポン 20回",
      "🏃 その場ドリブル：インサイドで 左右こうたい 30秒",
      "🦩 片足バランス：左右それぞれ 30秒ずつ",
      "🔄 ボールまわし：体のまわりでぐるぐる 右→左",
      "🧱 かべパス：インサイドでパス→トラップ 20回",
      "💨 ダッシュ＆ストップ：5歩ダッシュ→ぴたっと止まる 5回"
    ];
    const drill = drills[Math.floor(Math.random() * drills.length)];
    overlay.querySelector("[data-lock-drill]").textContent = drill;

    let remaining = BREAK_SEC;
    const timerEl = overlay.querySelector("[data-lock-timer]");
    function tick() {
      remaining--;
      if (timerEl) {
        const m = Math.floor(remaining / 60);
        const s = remaining % 60;
        timerEl.textContent = `${m}:${s < 10 ? "0" : ""}${s}`;
      }
      if (remaining <= 0) { clearInterval(lockInterval); doUnlock(overlay); }
    }
    lockInterval = setInterval(tick, 1000);

    const doneBtn = overlay.querySelector("[data-lock-done]");
    if (doneBtn) {
      const fresh = doneBtn.cloneNode(true);
      doneBtn.replaceWith(fresh);
      fresh.addEventListener("click", () => { clearInterval(lockInterval); doUnlock(overlay); });
    }
  }

  function doUnlock(overlay) {
    overlay.hidden = true;
    watchData.total = 0; // カウンターリセット（30分後また制限）
    saveWatchData(watchData);
  }

  function stopPlayTimer() {
    if (playTimer) { clearInterval(playTimer); playTimer = null; }
  }

  function showNudge(trigger) {
    const nudge = root.querySelector(".vz-practice-nudge");
    if (!nudge) return;
    lastNudgeSec = totalPlaySec;
    const sugg = PRACTICE_SUGGESTIONS[active] || PRACTICE_SUGGESTIONS.saved;
    const msg = trigger === "time"
      ? "⏰ 5分見たね！練習もしてみよう 💪"
      : trigger === "ended"
      ? "🎬 動画を見終わった！次は体で覚えよう"
      : "⏸ ちょっと休憩のすきに練習してみよう";
    nudge.hidden = false;
    nudge.innerHTML = `
      <div class="nudge-inner">
        <span class="nudge-icon" aria-hidden="true">${sugg.icon}</span>
        <span class="nudge-body">
          <b>${msg}</b>
          <span>${sugg.name}のやり方 → 練習ページへ</span>
        </span>
        <a href="training.html#skills" class="nudge-go">練習へ →</a>
        <button class="nudge-close" type="button" aria-label="閉じる">✕</button>
      </div>`;
    nudge.querySelector(".nudge-close").addEventListener("click", () => { nudge.hidden = true; });
    if (trigger !== "ended") setTimeout(() => { nudge.hidden = true; }, 20000);
  }

  // ---- 保存ずみ動画を「カテゴリ」として先頭に足す ----
  function allCats() {
    const list = cats.slice();
    list.unshift({
      key: "saved",
      label: `💾 保存した動画 (${saved.length})`,
      type: "videos",
      ids: saved.map((v) => v.id),
      desc: "あなたが保存した動画。サイトの中で見返せるよ。",
    });
    return list;
  }

  // ---- YouTube IFrame API ----
  function loadAPI() {
    if (window.YT && window.YT.Player) { apiReady = true; createPlayer(); return; }
    const tag = document.createElement("script");
    tag.src = "https://www.youtube.com/iframe_api";
    document.head.appendChild(tag);
    window.onYouTubeIframeAPIReady = () => {
      apiReady = true;
      createPlayer();
    };
  }

  function createPlayer() {
    player = new YT.Player("vz-player", {
      host: "https://www.youtube-nocookie.com",
      width: "100%", height: "100%",
      playerVars: { rel: 0, modestbranding: 1, playsinline: 1 },
      events: {
        onReady: () => { if (pending) { applyCat(pending.cat, pending.play); pending = null; } },
        onError: () => {
          // 埋め込み不可・削除ずみ・地域せいげんの動画は飛ばして次へ
          errorSkips++;
          if (errorSkips <= 6 && player && player.nextVideo) {
            try { player.nextVideo(); } catch (e2) {}
          }
        },
        onStateChange: (e) => {
          const state = (e && typeof e.data !== "undefined") ? e.data : -1;
          if (state === 1) { // PLAYING
            errorSkips = 0; // 再生できたのでリセット
            startPlayTimer();
          } else {
            stopPlayTimer();
            if (state === 0) { // ENDED
              showNudge("ended");
            } else if (state === 2 && totalPlaySec - lastNudgeSec > 90) { // PAUSED
              showNudge("paused");
            }
          }
          try {
            const d = player.getVideoData();
            if (d && d.video_id) currentVideo = { id: d.video_id, title: d.title || "" };
          } catch (e2) {}
          updateSaveBtn();
        },
      },
    });
  }

  // play=true なら すぐ再生（タブ切り替え時）／false なら読み込みだけ（初回）
  function applyCat(cat, play) {
    const frame = root.querySelector(".video-frame");
    const empty = root.querySelector(".vz-empty");
    if (cat.type === "videos" && (!cat.ids || !cat.ids.length)) {
      // 空（保存ゼロなど）
      if (player && player.stopVideo) player.stopVideo();
      frame.classList.add("is-empty");
      if (empty) empty.textContent = cat.key === "saved"
        ? "まだ保存した動画はないよ。動画の下の「💾 この動画を保存」でためていこう！"
        : "動画がありません。";
      return;
    }
    frame.classList.remove("is-empty");
    errorSkips = 0;
    if (cat.type === "playlist") {
      // UC...（チャンネルID）が来たら UU...（アップロード再生リスト）に変換
      const list = String(cat.id).replace(/^UC/, "UU");
      const opts = { listType: "playlist", list: list, index: 0 };
      if (play) player.loadPlaylist(opts); else player.cuePlaylist(opts);
    } else {
      if (play) player.loadPlaylist(cat.ids, 0); else player.cuePlaylist(cat.ids, 0);
    }
    currentVideo = null;
    updateSaveBtn();
  }

  // ---- 表示 ----（play=true でタブ切り替え時はすぐ再生）
  function show(cat, play) {
    active = cat.key;
    if (cat.key !== "saved") localStorage.setItem(VZ_KEY, active);
    descEl.textContent = cat.desc || "";
    tabsEl.querySelectorAll("button").forEach((b) => b.classList.toggle("active", b.dataset.key === cat.key));
    savedMgr.hidden = cat.key !== "saved";
    if (cat.key === "saved") renderSavedManager();
    updateStar();
    if (apiReady && player && player.cuePlaylist) applyCat(cat, play);
    else pending = { cat: cat, play: play };
  }

  function renderTabs() {
    const list = allCats();
    tabsEl.innerHTML = list.map((c) => `<button type="button" class="vz-tab${c.key === "saved" ? " saved" : ""}" data-key="${c.key}">${c.label}</button>`).join("");
    tabsEl.querySelectorAll("button").forEach((b) => {
      b.addEventListener("click", () => { const c = list.find((x) => x.key === b.dataset.key); if (c) show(c, true); });
    });
  }

  // ---- チャンネルお気に入り（タグ） ----
  function renderFavBar() {
    const items = favs.map((k) => cats.find((c) => c.key === k)).filter(Boolean);
    if (!items.length) { favBar.hidden = true; favBar.innerHTML = ""; return; }
    favBar.hidden = false;
    favBar.innerHTML = '<span class="vz-fav-label">⭐ お気に入りジャンル</span>' +
      items.map((c) => `<button type="button" class="vz-fav-chip" data-key="${c.key}">${c.label}</button>`).join("");
    favBar.querySelectorAll(".vz-fav-chip").forEach((b) => {
      b.addEventListener("click", () => { const c = cats.find((x) => x.key === b.dataset.key); if (c) show(c, true); });
    });
  }
  function updateStar() {
    if (!starBtn) return;
    const on = favs.includes(active);
    starBtn.classList.toggle("on", on);
    starBtn.style.display = active === "saved" ? "none" : "";
    starBtn.textContent = on ? "★ ジャンル登録ずみ" : "☆ ジャンルお気に入り";
  }
  if (starBtn) starBtn.addEventListener("click", () => {
    if (!active || active === "saved") return;
    const i = favs.indexOf(active);
    if (i >= 0) favs.splice(i, 1); else favs.push(active);
    write(VZ_FAV, favs); updateStar(); renderFavBar();
  });

  // ---- 動画そのものを保存 ----
  function updateSaveBtn() {
    if (!saveBtn) return;
    const v = currentVideo;
    const has = v && saved.some((s) => s.id === v.id);
    saveBtn.disabled = !v;
    saveBtn.classList.toggle("on", !!has);
    saveBtn.textContent = !v ? "💾 再生すると保存できるよ" : has ? "✓ 保存ずみ" : "💾 この動画を保存";
  }
  if (saveBtn) saveBtn.addEventListener("click", () => {
    if (!currentVideo) return;
    const idx = saved.findIndex((s) => s.id === currentVideo.id);
    if (idx >= 0) saved.splice(idx, 1);
    else saved.unshift({ id: currentVideo.id, title: currentVideo.title || "サッカー動画" });
    write(VZ_SAVED, saved);
    updateSaveBtn();
    renderTabs();
    tabsEl.querySelectorAll("button").forEach((b) => b.classList.toggle("active", b.dataset.key === active));
    if (active === "saved") { renderSavedManager(); }
  });

  function renderSavedManager() {
    if (!saved.length) { savedMgr.innerHTML = ""; return; }
    savedMgr.innerHTML = '<div class="vz-saved-title">💾 保存リスト（タップで再生／✕で削除）</div>' +
      saved.map((v, i) => `
        <div class="saved-item">
          <button type="button" class="saved-play" data-i="${i}">▶ ${escapeHtml(v.title || "動画")}</button>
          <button type="button" class="saved-del" data-del="${i}" aria-label="削除">✕</button>
        </div>`).join("");
    savedMgr.querySelectorAll(".saved-play").forEach((b) => b.addEventListener("click", () => {
      const i = +b.dataset.i;
      if (apiReady && player) { errorSkips = 0; player.loadPlaylist(saved.map((s) => s.id), i); }
    }));
    savedMgr.querySelectorAll(".saved-del").forEach((b) => b.addEventListener("click", () => {
      saved.splice(+b.dataset.del, 1); write(VZ_SAVED, saved);
      renderTabs(); renderSavedManager(); updateSaveBtn();
      const savedCat = allCats()[0];
      tabsEl.querySelectorAll("button").forEach((x) => x.classList.toggle("active", x.dataset.key === "saved"));
      if (apiReady && player) applyCat(savedCat, false);
    }));
  }

  function escapeHtml(s) { return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

  // ---- 初期化 ----
  (async function init() {
    try {
      const res = await fetch("data/channels.json", { cache: "no-store" });
      if (!res.ok) throw new Error();
      cats = (await res.json()).categories || CHANNELS_FALLBACK;
    } catch (e) { cats = CHANNELS_FALLBACK; }
    renderTabs();
    renderFavBar();
    loadAPI();
    const start = cats.find((c) => c.key === active) || cats[0];
    show(start);
  })();
})();
