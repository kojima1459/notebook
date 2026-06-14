/* =========================================================
   サンガキッズ - れんしゅうメニュー & 毎日チャレンジ
   - スター選手の技（見本さがしリンク付き）
   - レベル別 週間プログラム（今日やることを自動表示）
   - 毎日チャレンジ（ストリーク）＆ ごほうびバッジ
   データは localStorage に保存
   ========================================================= */

/* ---------- 技・練習メニューのデータ ----------
   1つ1つに「なぜ大事か(why)」「どこでやるか(place)」「つまずきポイント(tips)」
   「できたら次は(next)」を付けて、意味がわかってから練習できるようにする。 */
const SKILLS = [
  {
    icon: "🤹", name: "リフティング", level: "★☆☆ きほん", place: "🏠家 ・ 🌳外",
    why: "リフティングは“ボールと友だちになる”練習。ボールのどこを さわれば まっすぐ上がるかが、体でわかるようになるよ。これができると トラップ（止める）・ドリブル・シュートが ぜんぶ上手くなる。サッカーの“いちばん大事な土台”なんだ。",
    goal: "回数より「1回をていねいに」。目標は“昨日より1回多く”だけでOK！",
    steps: [
      "ボールを手でもって、ひざの高さからそっと落とす",
      "足の甲（くつのひもの上）の“たいらな所”でまっすぐ上にポンとける",
      "へそくらいの高さまで、やさしく上げるのがコツ（強すぎNG）",
      "1回けったら手でキャッチ → なれたら2回つづけてみよう"
    ],
    tips: "つまずきポイント：強くけりすぎると前に飛んでいく。「下から やさしく押し上げる」イメージで。10回できなくても、1回きれいに上がれば大せいこう！",
    next: "1回 → 2回 → 5回 → 左右の足こうたい → ももやatamaも使ってみよう"
  },
  {
    icon: "🏃", name: "ドリブル（まっすぐ）", level: "★☆☆ きほん", place: "🌳外",
    why: "ドリブルが上手だと、ボールを取られずに前へ運べる。「ボールを足もとに置いておく感覚」は、試合でいちばん使う力。まっすぐ運べるようになってから、ジグザグやフェイントに進もう。",
    goal: "ボールを足もとから はなさず、自分の思った所へ運ぶ",
    steps: [
      "ボールを足のすぐ前におく（1歩で とどく きょり）",
      "インサイド（足の内がわ）で 小さく・たくさん タッチして進む",
      "顔を上げて、前やまわりを見ながら 歩くスピードで",
      "コーン（ペットボトル）を 5本ならべて ジグザグも やってみよう"
    ],
    tips: "つまずきポイント：ボールをけりすぎて 足からはなれる子が多い。「ボールに手をそえて歩く」くらい 小さなタッチを意しきしよう。",
    next: "歩く → 小走り → ダッシュ。左足だけ・右足だけ でもチャレンジ"
  },
  {
    icon: "🧱", name: "かべパス（トラップ＆パス）", level: "★☆☆ きほん", place: "🏠家(やわらかいボール) ・ 🌳外",
    why: "パスとトラップ（止める）は サッカーの“会話”。味方とボールをつなぐ ためのいちばん大事な技術。かべは いつでも 正しい所へ返してくれる“最高の練習あいて”だよ。",
    goal: "インサイドで ねらった所へパス → かえってきたボールを ピタッと止める",
    steps: [
      "かべから 3〜4歩 はなれて立つ",
      "足の内がわ（くるぶしの下）で、かべの同じ所をねらってパス",
      "かえってきたら、足の内がわで“ふわっと”引いて止める（トラップ）",
      "止める→けるを リズムよく くりかえす"
    ],
    tips: "つまずきポイント：トラップで ボールが はねて しまう。止めるしゅんかんに 足を少し引くと、ボールが足もとに おちつくよ。",
    next: "止めないで“ワンタッチ”で返す → 左足でも → きょりを長くする"
  },
  {
    icon: "🎯", name: "シュート", level: "★★☆ ふつう", place: "🌳外",
    why: "サッカーは ゴールを決めると いちばん楽しい！ 強く正かくにけるには “けり方の形”が大事。なんとなくけるのと、形をおぼえてけるのでは、上たつのスピードが ぜんぜん ちがうよ。",
    goal: "ねらった所へ、強くまっすぐ ける",
    steps: [
      "ボールの横（10cmくらい）に、けらない方の足をおく＝じく足",
      "ける足は つま先を下にのばし、くつの甲でボールのまん中をミート",
      "けったあとも、足を ターゲットへ むかって ふりぬく",
      "かべや 小さなゴールに ねらって 10本"
    ],
    tips: "つまずきポイント：ボールの下をけると 上に飛ぶ。ボールの“まん中”を まっすぐ ミートしよう。じく足を ターゲットに向けるのも コツ。",
    next: "右足10本→左足10本 → 走りながら → かべに当てて返ってきた所をシュート"
  },
  {
    icon: "👟", name: "シザース（またぎフェイント）", level: "★★☆ ふつう", place: "🌳外",
    why: "あいてを“だます”フェイント。サッカーは1対1で ぬけるとチャンスが生まれる。シザースは プロもよく使う、いちばん有名で かっこいいフェイントだよ。",
    goal: "ボールをまたいで、あいてを 反対がわに ひっかける",
    steps: [
      "ドリブル中、ボールの前で 足を 外から内へ「またぐ」",
      "またいだ足を 地面につけたら、その足で 体重を かたむける（行くふり）",
      "すぐに ぎゃくの足の外がわで 反対方向へ タッチして ぬける",
      "ゆっくり→だんだん速く。リズムが大事！"
    ],
    tips: "つまずきポイント：またぐだけで終わると あいては だまされない。「行くふり」を 大きく見せて、ぎゃくへ すばやく！",
    next: "止まってまたぐ → 歩きながら → ドリブル中に → あいて役の前で実せん"
  },
  {
    icon: "🌀", name: "ルーレット（マルセイユターン）", level: "★★★ じょうきゅう", place: "🌳外",
    why: "くるっと回って あいてに せなかを向けながら ボールをまもる技。せまい所で かこまれても 脱出できる。メッシやジダンも使った“魔法のターン”だよ。",
    goal: "回りながら ボールをまもって、あいてを かわす",
    steps: [
      "足のうらで ボールを手前に引く",
      "引いた いきおいで 体を 半回てんさせる",
      "ぎゃくの足のうらで ボールを また引いて、回りきる",
      "回ったあと そのまま 前へドリブル"
    ],
    tips: "つまずきポイント：ボールを置いていきがち。足のうらで“ボールに ふれ続ける”のがコツ。最初は その場で ゆっくり 練習しよう。",
    next: "その場で → 歩きながら → スピードに乗って"
  },
  {
    icon: "⚡", name: "エラシコ", level: "★★★ じょうきゅう", place: "🌳外",
    why: "ボールを 外→内へ 一しゅんで動かす、上きゅうフェイント。むずかしいけど、できると“テクニシャン”！ むずかしい技に ちょうせんする気もちが、上たつを はやくするよ。",
    goal: "外へ行くふり → 一しゅんで内へ。あいてを 反対へ ひっかける",
    steps: [
      "つま先の外がわで ボールを ちょっとだけ 外へおす",
      "すぐに 同じ足の内がわで 内へ ひっぱりもどす",
      "「トン・トン」と 一しゅんで！　足首を やわらかく使う",
      "むずかしいから、その場で 何回も くりかえそう"
    ],
    tips: "つまずきポイント：2タッチが ゆっくりだと フェイントにならない。“1つの動き”に見えるくらい すばやく。まずは ゆっくり 形をおぼえてから スピードアップ。",
    next: "その場で形 → 歩きながら → ドリブルから → シュートにつなげる"
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
const SKILL_DONE_KEY = "sangakids_skill_done";

function loadState() {
  try { return JSON.parse(localStorage.getItem(STORE_KEY)) || {}; }
  catch (e) { return {}; }
}
function saveState(s) {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(s)); } catch (e) {}
}
function loadSkillDone() {
  try { return JSON.parse(localStorage.getItem(SKILL_DONE_KEY)) || {}; }
  catch (e) { return {}; }
}
function saveSkillDone(d) {
  try { localStorage.setItem(SKILL_DONE_KEY, JSON.stringify(d)); } catch (e) {}
}
function dateKey(offset) {
  const d = new Date();
  if (offset) d.setDate(d.getDate() + offset);
  return `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`;
}
const todayStr = () => dateKey(0);
const yesterdayStr = () => dateKey(-1);

/* ---------- できた！記録の状態を更新 ---------- */
function renderSkillDoneStates() {
  const done = loadSkillDone();
  document.querySelectorAll(".skill-card").forEach((card, i) => {
    const skill = SKILLS[i];
    if (!skill) return;
    const doneDate = done[skill.name];
    const doneToday = doneDate === todayStr();
    const badge = card.querySelector(".skill-done-badge");
    if (badge) {
      badge.textContent = doneToday ? "✅ 今日できた！" : doneDate ? `📅 ${doneDate} にできた` : "";
      badge.hidden = !doneDate;
    }
    const btn = card.querySelector("[data-skill-done]");
    if (btn) {
      btn.textContent = doneToday ? "✅ 今日できた！（取り消す）" : "✓ できた！記録する";
      btn.classList.toggle("done", doneToday);
    }
  });
}

/* ---------- 技カードの描画（見本さがしリンク＋できた！ボタン） ---------- */
(function renderSkills() {
  const wrap = document.querySelector("[data-skills]");
  if (!wrap) return;
  wrap.innerHTML = SKILLS.map((s, i) => {
    const query = encodeURIComponent(`サッカー ${s.name.replace(/（.*）/, "")} やり方 子ども`);
    const url = `https://www.youtube.com/results?search_query=${query}`;
    return `
    <div class="skill-card">
      <div class="skill-done-badge" hidden></div>
      <button class="skill-head" type="button" aria-expanded="false" data-skill="${i}">
        <span class="skill-icon" aria-hidden="true">${s.icon}</span>
        <span class="skill-title">
          <strong>${s.name}</strong>
          <small>${s.level}${s.place ? ` ・ ${s.place}` : ""}</small>
        </span>
        <span class="skill-arrow" aria-hidden="true">▼</span>
      </button>
      <div class="skill-body" hidden>
        ${s.why ? `<p class="skill-why"><b>💡 なぜ大事？</b>${s.why}</p>` : ""}
        <p class="skill-goal">🎯 めあて：${s.goal}</p>
        <p class="skill-steplabel">📋 正しいやり方（じゅんばん）</p>
        <ol class="skill-steps">${s.steps.map((t) => `<li>${t}</li>`).join("")}</ol>
        ${s.tips ? `<p class="skill-tips"><b>⚠️ つまずきポイント＆コツ</b><br>${s.tips}</p>` : ""}
        ${s.next ? `<p class="skill-next"><b>📈 できたら つぎは</b><br>${s.next}</p>` : ""}
        <a class="video-link" href="${url}" target="_blank" rel="noopener noreferrer">
          🎥 見本どうがを さがす（おうちの人と いっしょに）
        </a>
        <button type="button" class="skill-done-btn" data-skill-done="${i}">✓ できた！記録する</button>
      </div>
    </div>`;
  }).join("");

  wrap.querySelectorAll(".skill-head").forEach((btn) => {
    btn.addEventListener("click", () => {
      const body = btn.closest(".skill-card").querySelector(".skill-body");
      const open = btn.getAttribute("aria-expanded") === "true";
      btn.setAttribute("aria-expanded", String(!open));
      body.hidden = open;
      btn.querySelector(".skill-arrow").textContent = open ? "▼" : "▲";
    });
  });

  wrap.querySelectorAll("[data-skill-done]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const i = +btn.dataset.skillDone;
      const skill = SKILLS[i];
      const d = loadSkillDone();
      if (d[skill.name] === todayStr()) {
        delete d[skill.name];
      } else {
        d[skill.name] = todayStr();
        celebrate();
      }
      saveSkillDone(d);
      renderSkillDoneStates();
      renderTodayDrill();
    });
  });

  renderSkillDoneStates();
})();

/* ---------- 今日の1つ おすすめドリル ---------- */
function renderTodayDrill() {
  const wrap = document.querySelector("[data-today-drill]");
  if (!wrap) return;

  const state = loadState();
  const level = state.level || "beginner";
  const dow = new Date().getDay();

  // レベル別・曜日別のおすすめ（SKILLSの添字）
  const rotation = {
    beginner:     [0, 1, 2, 0, 1, 3, 0],
    intermediate: [0, 1, 4, 2, 3, 3, 1],
    advanced:     [0, 5, 4, 6, 3, 2, 5],
  };
  const idx = (rotation[level] || rotation.beginner)[dow];
  const drill = SKILLS[idx];
  if (!drill) return;

  const done = loadSkillDone();
  const doneToday = done[drill.name] === todayStr();

  wrap.innerHTML = `
    <div class="today-drill ${doneToday ? "is-done" : ""}">
      <div class="td-label">📅 今日やる1つ ${doneToday ? "✅ できた！" : "← まずこれだけ！"}</div>
      <div class="td-main">
        <span class="td-icon">${drill.icon}</span>
        <div class="td-info">
          <b class="td-name">${drill.name}</b>
          <span class="td-where">${drill.place || ""}</span>
        </div>
        ${doneToday ? "" : `<button type="button" class="td-done-btn" data-td-done="${idx}">✓ できた！</button>`}
      </div>
      <p class="td-why">${drill.why ? drill.why.slice(0, 100) + "…" : drill.goal}</p>
      <a href="#skills" class="td-go">📋 正しいやり方を見る ↓</a>
    </div>`;

  const btn = wrap.querySelector("[data-td-done]");
  if (btn) btn.addEventListener("click", () => {
    const i = +btn.dataset.tdDone;
    const d = loadSkillDone();
    d[SKILLS[i].name] = todayStr();
    saveSkillDone(d);
    celebrate();
    renderTodayDrill();
    renderSkillDoneStates();
  });
}
renderTodayDrill();

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
