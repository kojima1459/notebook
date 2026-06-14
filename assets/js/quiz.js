/* =========================================================
   サンガキッズ - サッカークイズ
   子供向けのやさしいサッカークイズ
   ========================================================= */

const QUIZ = [
  {
    q: "サッカーは1チーム、なん人でプレーするかな？",
    options: ["9人", "11人", "13人", "15人"],
    answer: 1,
    hint: "ゴールキーパーをいれた数だよ！"
  },
  {
    q: "サッカーのしあいは、まえはんと後はんで合計なん分かな？",
    options: ["60分", "70分", "90分", "120分"],
    answer: 2,
    hint: "45分ずつ、2回プレーするよ。"
  },
  {
    q: "京都サンガF.C. のチームカラーはなに色かな？",
    options: ["あか", "みどり", "むらさき", "あお"],
    answer: 2,
    hint: "ヒント：パープル！"
  },
  {
    q: "手をつかってボールをとめてもいい人はだれ？",
    options: ["フォワード", "ゴールキーパー", "しんぱん", "かんとく"],
    answer: 1,
    hint: "ゴールをまもるせんしゅだよ。"
  },
  {
    q: "ボールがゴールにはいると、なにが入るかな？",
    options: ["ホームラン", "トライ", "ゴール（とくてん）", "アウト"],
    answer: 2,
    hint: "サッカーは足でゴールをねらうスポーツ！"
  },
  {
    q: "わるいことをしたときに出される赤いカードのなまえは？",
    options: ["レッドカード", "ブルーカード", "ゴールドカード", "ストップカード"],
    answer: 0,
    hint: "あかい色のカードだよ。"
  },
  {
    q: "京都サンガのホームスタジアムがある町はどこ？",
    options: ["京都市", "亀岡市", "大阪市", "神戸市"],
    answer: 1,
    hint: "サンガスタジアム by KYOCERA があるよ。"
  }
];

(function setupQuiz() {
  const root = document.querySelector("[data-quiz]");
  if (!root) return;

  let current = 0;
  let score = 0;
  let answered = false;

  const qEl = root.querySelector(".quiz-question");
  const optsEl = root.querySelector(".quiz-options");
  const fbEl = root.querySelector(".quiz-feedback");
  const nextBtn = root.querySelector(".quiz-next");
  const progEl = root.querySelector(".quiz-progress");
  const barEl = root.querySelector(".quiz-bar > i");

  function render() {
    answered = false;
    const item = QUIZ[current];
    progEl.textContent = `だい ${current + 1} もん / ${QUIZ.length} もん`;
    barEl.style.width = `${(current / QUIZ.length) * 100}%`;
    qEl.textContent = item.q;
    fbEl.textContent = "";
    fbEl.style.color = "";
    nextBtn.disabled = true;
    nextBtn.textContent = current === QUIZ.length - 1 ? "けっか を みる →" : "つぎ の もんだい →";

    optsEl.innerHTML = "";
    item.options.forEach((opt, i) => {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = opt;
      b.addEventListener("click", () => choose(i, b));
      optsEl.appendChild(b);
    });
  }

  function choose(i, btn) {
    if (answered) return;
    answered = true;
    const item = QUIZ[current];
    const buttons = optsEl.querySelectorAll("button");
    buttons.forEach((b) => (b.disabled = true));

    if (i === item.answer) {
      btn.classList.add("correct");
      fbEl.textContent = "せいかい！🎉 すごい！";
      fbEl.style.color = "#1f9d6b";
      score++;
    } else {
      btn.classList.add("wrong");
      buttons[item.answer].classList.add("correct");
      fbEl.textContent = "ざんねん… " + item.hint;
      fbEl.style.color = "#d6336c";
    }
    nextBtn.disabled = false;
  }

  function showResult() {
    barEl.style.width = "100%";
    progEl.textContent = "おわり！";
    const ratio = score / QUIZ.length;
    let medal = "🎽";
    let msg = "またチャレンジしてね！";
    if (ratio === 1) { medal = "🏆"; msg = "パーフェクト！サッカーはかせだ！"; }
    else if (ratio >= 0.7) { medal = "🥇"; msg = "すごい！もうサンガ博士だね！"; }
    else if (ratio >= 0.4) { medal = "🥈"; msg = "いいかんじ！あと少し！"; }
    else { medal = "🥉"; msg = "これからもっとくわしくなろう！"; }

    root.innerHTML = `
      <div class="quiz-result">
        <div class="medal">${medal}</div>
        <p>きみの スコアは…</p>
        <div class="score">${score} / ${QUIZ.length}</div>
        <p>${msg}</p>
        <button class="quiz-next" type="button" style="max-width:280px;margin:18px auto 0;">
          もういちど あそぶ 🔄
        </button>
      </div>`;
    root.querySelector(".quiz-next").addEventListener("click", () => location.reload());
  }

  nextBtn.addEventListener("click", () => {
    if (!answered) return;
    current++;
    if (current >= QUIZ.length) showResult();
    else render();
  });

  render();
})();
