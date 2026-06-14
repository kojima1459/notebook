/* =========================================================
   サンガキッズ - れんしゅうメニュー & 毎日チャレンジ
   - スター選手の技（見本さがしリンク付き）
   - レベル別 週間プログラム（今日やることを自動表示）
   - 毎日チャレンジ（ストリーク）＆ ごほうびバッジ
   データは localStorage に保存
   ========================================================= */

/* ---------- 技・練習メニューのデータ ---------- */
const SKILLS = [
  {
    icon: "🤹", name: "リフティング", level: "★☆☆ かんたん",
    goal: "ボールを足でポンポン！まずは10回をめざそう",
    steps: [
      "ボールを手でもって、ひざのたかさからそっと落とす",
      "足の甲（くつのひもの上）でまっすぐ上にけりあげる",
      "1回ごとに手でキャッチ → なれたら2回つづけてみよう",
      "目ひょう：5回 → 10回 → 30回とふやしていこう！"
    ]
  },
  {
    icon: "🏃", name: "ドリブル（まっすぐ）", level: "★☆☆ かんたん",
    goal: "ボールを足もとからはなさず、まっすぐ運ぼう",
    steps: [
      "ボールを足のすぐ前におく（はなれすぎない）",
      "インサイド（足の内がわ）で小さくタッチしながら進む",
      "顔を上げて、まわりを見ながら歩くスピードで",
      "なれたら少しずつスピードアップ！"
    ]
  },
  {
    icon: "🎯", name: "シュート", level: "★★☆ ふつう",
    goal: "ねらったところに、つよくけろう",
    steps: [
      "ボールの横に、けらない方の足をおく（じくあし）",
      "ける足はくつの甲でボールのまん中をミート",
      "けったあと、足をターゲットにむかってふりぬく",
      "かべや小さなゴールにねらってシュート練習！"
    ]
  },
  {
    icon: "👟", name: "シザース（またぎフェイント）", level: "★★☆ ふつう",
    goal: "ボールをまたいで、あいてをだますフェイント",
    steps: [
      "ボールの前で、足を外から内へ「またぐ」",
      "またいだ足を地面につけたら、ぎゃくの足でよこへタッチ",
      "ゆっくり→だんだん速く。リズムが大事！",
      "実せん：あいての前でまたいで、すきまをぬけよう"
    ]
  },
  {
    icon: "🌀", name: "ルーレット（マルセイユターン）", level: "★★★ じょうきゅう",
    goal: "くるっと回ってあいてをかわす、かっこいい技",
    steps: [
      "足のうらでボールを手前に引く",
      "そのまま体を半回てんさせる",
      "ぎゃくの足のうらでボールをまた引いて、回りきる",
      "スター選手もつかう技！ゆっくりから練習しよう"
    ]
  },
  {
    icon: "⚡", name: "エラシコ", level: "★★★ じょうきゅう",
    goal: "ボールを外→内へ一しゅんで動かすフェイント",
    steps: [
      "つま先の外がわでボールをちょっと外へおす",
      "すぐに同じ足の内がわで内へひっぱりもどす",
      "「トン・トン」と一しゅんで！むずかしいから何回も練習",
      "できたらかなりのテクニシャン！"
    ]
  }
];

/* ---------- レベル別 週間プログラム ---------- */
const WEEK_LABELS = ["日", "月", "火", "水", "木", "金", "土"];
const PROGRAMS = {
  beginner: {
    name: "初級（はじめてさん）",
    days: [
      "公園でいっぱい走る＆ボールであそぶ（おやすみDAY）",
      "リフティング 10回にちょうせん",
      "まっすぐドリブル 往ふく5回",
      "インサイドパス（かべ）20回",
      "リフティング＋ドリブル ミックス",
      "シュート 10本（かべ・ゴール）",
      "好きな技を1つえらんで練習！"
    ]
  },
  intermediate: {
    name: "中級（もっと上手に）",
    days: [
      "ミニゲーム or おにごっこで体を動かす",
      "リフティング 30回にちょうせん",
      "コーンドリブル（ジグザグ）5往ふく",
      "シザース → ドリブル 10回",
      "両足インサイドパス 各30回",
      "シュート 20本（左右の足）",
      "おぼえた技で1分間チャレンジ"
    ]
  },
  advanced: {
    name: "上級（テクニシャン）",
    days: [
      "ランニング＆ストレッチでコンディション",
      "リフティング 50回＋もも・頭ミックス",
      "ルーレットを左右 各10回",
      "エラシコ 20回（ゆっくり→速く）",
      "フェイント→シュートのコンビ 15回",
      "1対1のかけひき練習（あいて役とでも）",
      "今週おぼえた技をぜんぶつなげてみる！"
    ]
  }
};

/* ---------- 毎日チャレンジ（ミッション） ---------- */
const CHALLENGES = [
  { id: "lift", icon: "🤹", text: "リフティング 10回 ちょうせん" },
  { id: "dribble", icon: "🏃", text: "コーン（ペットボトル）ドリブル 5往ふく" },
  { id: "shoot", icon: "🎯", text: "かべ または ゴールに シュート 10本" },
  { id: "wall", icon: "🧱", text: "かべパス（インサイド）20回" },
  { id: "run", icon: "💨", text: "ダッシュ → ストップ を 5回" },
  { id: "stretch", icon: "🤸", text: "じゅんび体そう & ストレッチ" }
];

/* ---------- ごほうびバッジ ---------- */
// got(state) が true になると獲得（かくとく）
const BADGES = [
  { icon: "🌱", name: "はじめの一歩", desc: "1日 ぜんぶ達成", got: (s) => s.totalDays >= 1 },
  { icon: "🔥", name: "3日れんぞく", desc: "3日つづけて達成", got: (s) => s.best >= 3 },
  { icon: "⭐", name: "1週間れんぞく", desc: "7日つづけて達成", got: (s) => s.best >= 7 },
  { icon: "🏅", name: "がんばり10", desc: "ぜんぶ達成 累計10日", got: (s) => s.totalDays >= 10 },
  { icon: "💎", name: "2週間れんぞく", desc: "14日つづけて達成", got: (s) => s.best >= 14 },
  { icon: "🏆", name: "テクニシャン", desc: "ぜんぶ達成 累計30日", got: (s) => s.totalDays >= 30 },
  { icon: "👑", name: "サンガマスター", desc: "30日れんぞく達成", got: (s) => s.best >= 30 }
];

const STORE_KEY = "sangakids_training_v1";

function loadState() {
  try { return JSON.parse(localStorage.getItem(STORE_KEY)) || {}; }
  catch (e) { return {}; }
}
function saveState(s) {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(s)); } catch (e) {}
}
function dateKey(offset) {
  const d = new Date();
  if (offset) d.setDate(d.getDate() + offset);
  return `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`;
}
const todayStr = () => dateKey(0);
const yesterdayStr = () => dateKey(-1);

/* ---------- 技カードの描画（見本さがしリンク付き） ---------- */
(function renderSkills() {
  const wrap = document.querySelector("[data-skills]");
  if (!wrap) return;
  wrap.innerHTML = SKILLS.map((s, i) => {
    const query = encodeURIComponent(`サッカー ${s.name.replace(/（.*）/, "")} やり方 子ども`);
    const url = `https://www.youtube.com/results?search_query=${query}`;
    return `
    <div class="skill-card">
      <button class="skill-head" type="button" aria-expanded="false" data-skill="${i}">
        <span class="skill-icon" aria-hidden="true">${s.icon}</span>
        <span class="skill-title"><strong>${s.name}</strong><small>${s.level}</small></span>
        <span class="skill-arrow" aria-hidden="true">▼</span>
      </button>
      <div class="skill-body" hidden>
        <p class="skill-goal">🎯 めあて：${s.goal}</p>
        <ol class="skill-steps">${s.steps.map((t) => `<li>${t}</li>`).join("")}</ol>
        <a class="video-link" href="${url}" target="_blank" rel="noopener noreferrer">
          🎥 見本どうがを さがす（おうちの人と いっしょに）
        </a>
      </div>
    </div>`;
  }).join("");

  wrap.querySelectorAll(".skill-head").forEach((btn) => {
    btn.addEventListener("click", () => {
      const body = btn.nextElementSibling;
      const open = btn.getAttribute("aria-expanded") === "true";
      btn.setAttribute("aria-expanded", String(!open));
      body.hidden = open;
      btn.querySelector(".skill-arrow").textContent = open ? "▼" : "▲";
    });
  });
})();

/* ---------- レベル別 週間プログラム ---------- */
(function renderProgram() {
  const wrap = document.querySelector("[data-program]");
  if (!wrap) return;

  const state = loadState();
  state.level = state.level || "beginner";

  const tabsEl = wrap.querySelector("[data-program-tabs]");
  const todayEl = wrap.querySelector("[data-program-today]");
  const listEl = wrap.querySelector("[data-program-week]");
  const dow = new Date().getDay();

  function render() {
    const prog = PROGRAMS[state.level];
    // タブ
    tabsEl.innerHTML = Object.keys(PROGRAMS).map((key) =>
      `<button type="button" class="lv-tab ${key === state.level ? "active" : ""}" data-level="${key}">
        ${PROGRAMS[key].name}
      </button>`
    ).join("");
    // 今日やること
    todayEl.innerHTML = `
      <div class="today-label">📅 きょう（${WEEK_LABELS[dow]}よう日）やること</div>
      <div class="today-task">${prog.days[dow]}</div>`;
    // 週間リスト
    listEl.innerHTML = prog.days.map((task, i) =>
      `<li class="week-day ${i === dow ? "is-today" : ""}">
        <span class="wd-name">${WEEK_LABELS[i]}</span>
        <span class="wd-task">${task}</span>
      </li>`
    ).join("");

    tabsEl.querySelectorAll(".lv-tab").forEach((b) => {
      b.addEventListener("click", () => {
        state.level = b.dataset.level;
        saveState(state);
        render();
      });
    });
  }
  render();
})();

/* ---------- 毎日チャレンジ ＆ バッジ ---------- */
(function renderChallenge() {
  const wrap = document.querySelector("[data-challenge]");
  if (!wrap) return;

  const state = loadState();
  if (state.date !== todayStr()) {
    state.date = todayStr();
    state.done = {};
    saveState(state);
  }
  state.streak = state.streak || 0;
  state.best = state.best || 0;
  state.totalDays = state.totalDays || 0;

  const listEl = wrap.querySelector("[data-challenge-list]");
  const streakEl = wrap.querySelector("[data-streak]");
  const bestEl = wrap.querySelector("[data-best]");
  const totalEl = wrap.querySelector("[data-total]");
  const msgEl = wrap.querySelector("[data-challenge-msg]");
  const badgesEl = document.querySelector("[data-badges]");

  const countDone = () =>
    CHALLENGES.filter((c) => state.done && state.done[c.id]).length;

  function update() {
    streakEl.textContent = state.streak;
    bestEl.textContent = state.best;
    if (totalEl) totalEl.textContent = state.totalDays;
    const done = countDone();
    msgEl.textContent = done === CHALLENGES.length
      ? "🎉 ぜんぶクリア！今日のれんしゅう おつかれさま！"
      : `あと ${CHALLENGES.length - done} こ！がんばれ💪`;
    renderBadges();
  }

  function renderBadges() {
    if (!badgesEl) return;
    badgesEl.innerHTML = BADGES.map((b) => {
      const got = b.got(state);
      return `
        <div class="badge ${got ? "got" : "locked"}" title="${b.desc}">
          <div class="badge-icon">${got ? b.icon : "🔒"}</div>
          <div class="badge-name">${b.name}</div>
          <div class="badge-desc">${b.desc}</div>
        </div>`;
    }).join("");
  }

  function render() {
    listEl.innerHTML = CHALLENGES.map((c) => {
      const checked = state.done && state.done[c.id];
      return `
        <li>
          <label class="mission ${checked ? "done" : ""}">
            <input type="checkbox" data-mission="${c.id}" ${checked ? "checked" : ""} />
            <span class="mission-icon" aria-hidden="true">${c.icon}</span>
            <span class="mission-text">${c.text}</span>
            <span class="mission-check" aria-hidden="true">✓</span>
          </label>
        </li>`;
    }).join("");

    listEl.querySelectorAll("input[data-mission]").forEach((box) => {
      box.addEventListener("change", () => {
        const id = box.dataset.mission;
        state.done = state.done || {};
        const before = countDone();
        state.done[id] = box.checked;
        box.closest(".mission").classList.toggle("done", box.checked);

        const after = countDone();
        if (before < CHALLENGES.length && after === CHALLENGES.length) {
          if (state.lastComplete !== todayStr()) {
            state.streak = state.lastComplete === yesterdayStr() ? state.streak + 1 : 1;
            state.lastComplete = todayStr();
            state.best = Math.max(state.best, state.streak);
            state.totalDays += 1;
            celebrate();
          }
        }
        saveState(state);
        update();
      });
    });
  }

  function celebrate() {
    const c = document.createElement("div");
    c.className = "confetti";
    c.textContent = "🎉⚽🌟💜✨🏆";
    document.body.appendChild(c);
    setTimeout(() => c.remove(), 1500);
  }

  render();
  update();

  const resetBtn = wrap.querySelector("[data-challenge-reset]");
  if (resetBtn) {
    resetBtn.addEventListener("click", () => {
      state.done = {};
      saveState(state);
      render();
      update();
    });
  }
})();
