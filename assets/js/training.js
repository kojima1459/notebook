/* =========================================================
   サンガキッズ - れんしゅうメニュー & 毎日チャレンジ
   スター選手の技と、上達のための練習メニュー
   毎日チャレンジは localStorage で連続記録（ストリーク）を保存
   ========================================================= */

/* ---------- 技・練習メニューのデータ ---------- */
const SKILLS = [
  {
    icon: "🤹",
    name: "リフティング",
    level: "★☆☆ かんたん",
    goal: "ボールを足でポンポン！まずは10回をめざそう",
    steps: [
      "ボールを手でもって、ひざのたかさからそっと落とす",
      "足の甲（くつのひもの上）でまっすぐ上にけりあげる",
      "1回ごとに手でキャッチ → なれたら2回つづけてみよう",
      "目ひょう：5回 → 10回 → 30回とふやしていこう！"
    ]
  },
  {
    icon: "🏃",
    name: "ドリブル（まっすぐ）",
    level: "★☆☆ かんたん",
    goal: "ボールを足もとからはなさず、まっすぐ運ぼう",
    steps: [
      "ボールを足のすぐ前におく（はなれすぎない）",
      "インサイド（足の内がわ）で小さくタッチしながら進む",
      "顔を上げて、まわりを見ながら歩くスピードで",
      "なれたら少しずつスピードアップ！"
    ]
  },
  {
    icon: "🎯",
    name: "シュート",
    level: "★★☆ ふつう",
    goal: "ねらったところに、つよくけろう",
    steps: [
      "ボールの横に、けらない方の足をおく（じくあし）",
      "ける足はくつの甲でボールのまん中をミート",
      "けったあと、足をターゲットにむかってふりぬく",
      "かべや小さなゴールにねらってシュート練習！"
    ]
  },
  {
    icon: "👟",
    name: "シザース（またぎフェイント）",
    level: "★★☆ ふつう",
    goal: "ボールをまたいで、あいてをだますフェイント",
    steps: [
      "ボールの前で、足を外から内へ「またぐ」",
      "またいだ足を地面につけたら、ぎゃくの足でよこへタッチ",
      "ゆっくり→だんだん速く。リズムが大事！",
      "実せん：あいての前でまたいで、すきまをぬけよう"
    ]
  },
  {
    icon: "🌀",
    name: "ルーレット（マルセイユターン）",
    level: "★★★ じょうきゅう",
    goal: "くるっと回ってあいてをかわす、かっこいい技",
    steps: [
      "足のうらでボールを手前に引く",
      "そのまま体を半回てんさせる",
      "ぎゃくの足のうらでボールをまた引いて、回りきる",
      "スター選手もつかう技！ゆっくりから練習しよう"
    ]
  },
  {
    icon: "⚡",
    name: "エラシコ",
    level: "★★★ じょうきゅう",
    goal: "ボールを外→内へ一しゅんで動かすフェイント",
    steps: [
      "つま先の外がわでボールをちょっと外へおす",
      "すぐに同じ足の内がわで内へひっぱりもどす",
      "「トン・トン」と一しゅんで！むずかしいから何回も練習",
      "できたらかなりのテクニシャン！"
    ]
  }
];

/* ---------- 毎日チャレンジ（ミッション） ---------- */
const CHALLENGES = [
  { id: "lift", icon: "🤹", text: "リフティング 10回 ちょうせん" },
  { id: "dribble", icon: "🏃", text: "コーン（ペットボトル）ドリブル 5往ふく" },
  { id: "shoot", icon: "🎯", text: "かべ または ゴールに シュート 10本" },
  { id: "wall", icon: "🧱", text: "かべパス（インサイド）20回" },
  { id: "run", icon: "💨", text: "ダッシュ → ストップ を 5回" },
  { id: "stretch", icon: "🤸", text: "じゅんび体そう & ストレッチ" }
];

const STORE_KEY = "sangakids_training_v1";

function loadState() {
  try {
    return JSON.parse(localStorage.getItem(STORE_KEY)) || {};
  } catch (e) {
    return {};
  }
}
function saveState(s) {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify(s));
  } catch (e) {}
}
function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`;
}
function yesterdayStr() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`;
}

/* ---------- 技カードの描画 ---------- */
(function renderSkills() {
  const wrap = document.querySelector("[data-skills]");
  if (!wrap) return;
  wrap.innerHTML = SKILLS.map(
    (s, i) => `
    <div class="skill-card">
      <button class="skill-head" type="button" aria-expanded="false" data-skill="${i}">
        <span class="skill-icon" aria-hidden="true">${s.icon}</span>
        <span class="skill-title">
          <strong>${s.name}</strong>
          <small>${s.level}</small>
        </span>
        <span class="skill-arrow" aria-hidden="true">▼</span>
      </button>
      <div class="skill-body" hidden>
        <p class="skill-goal">🎯 めあて：${s.goal}</p>
        <ol class="skill-steps">
          ${s.steps.map((t) => `<li>${t}</li>`).join("")}
        </ol>
      </div>
    </div>`
  ).join("");

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

/* ---------- 毎日チャレンジの描画 ---------- */
(function renderChallenge() {
  const wrap = document.querySelector("[data-challenge]");
  if (!wrap) return;

  const state = loadState();
  // 日付がかわったらチェックをリセット（ストリークは保持）
  if (state.date !== todayStr()) {
    state.date = todayStr();
    state.done = {};
    saveState(state);
  }
  state.streak = state.streak || 0;
  state.best = state.best || 0;

  const listEl = wrap.querySelector("[data-challenge-list]");
  const streakEl = wrap.querySelector("[data-streak]");
  const bestEl = wrap.querySelector("[data-best]");
  const msgEl = wrap.querySelector("[data-challenge-msg]");

  function countDone() {
    return CHALLENGES.filter((c) => state.done && state.done[c.id]).length;
  }

  function update() {
    streakEl.textContent = state.streak;
    bestEl.textContent = state.best;
    const done = countDone();
    if (done === CHALLENGES.length) {
      msgEl.textContent = "🎉 ぜんぶクリア！今日のれんしゅう おつかれさま！";
    } else {
      msgEl.textContent = `あと ${CHALLENGES.length - done} こ！がんばれ💪`;
    }
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
        // 「ぜんぶ達成」になった瞬間にストリークを＋1
        if (before < CHALLENGES.length && after === CHALLENGES.length) {
          if (state.lastComplete !== todayStr()) {
            state.streak = state.lastComplete === yesterdayStr() ? state.streak + 1 : 1;
            state.lastComplete = todayStr();
            state.best = Math.max(state.best, state.streak);
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
