/* =========================================================
   サンガキッズ - サッカークイズ
   子供向けのやさしいサッカークイズ
   ========================================================= */

const QUIZ = [
  {
    q: "京都サンガF.C. のチームカラーは？",
    options: ["赤", "緑", "紫（パープル）", "青"],
    answer: 2,
    hint: "サンガといえばこの色！"
  },
  {
    q: "京都サンガのホームスタジアムがある市は？",
    options: ["京都市", "亀岡市", "大阪市", "宇治市"],
    answer: 1,
    hint: "サンガスタジアム by KYOCERA があるよ。"
  },
  {
    q: "京都サンガが所属しているリーグは？",
    options: ["J1", "J2", "J3", "プレミアリーグ"],
    answer: 0,
    hint: "日本の一番上のカテゴリだよ。"
  },
  {
    q: "サンガスタジアム by KYOCERA の特長は？",
    options: ["屋内コート", "ピッチと客席がとても近い", "野球もできる", "海の上にある"],
    answer: 1,
    hint: "選手との距離が近い、サッカー専用スタジアム。"
  },
  {
    q: "2022年 W杯（カタール大会）で優勝した国は？",
    options: ["フランス", "ブラジル", "アルゼンチン", "ドイツ"],
    answer: 2,
    hint: "メッシのいる国だよ。"
  },
  {
    q: "リオネル・メッシの出身国は？",
    options: ["スペイン", "アルゼンチン", "ブラジル", "ポルトガル"],
    answer: 1,
    hint: "南米の国。W杯2022優勝国でもある。"
  },
  {
    q: "クリスティアーノ・ロナウドの出身国は？",
    options: ["スペイン", "ブラジル", "ポルトガル", "イタリア"],
    answer: 2,
    hint: "スペインのとなりの国。"
  },
  {
    q: "「プレミアリーグ」がある国は？",
    options: ["スペイン", "イングランド", "ドイツ", "フランス"],
    answer: 1,
    hint: "世界一人気とも言われるリーグ。"
  },
  {
    q: "「ラ・リーガ」がある国は？",
    options: ["スペイン", "イタリア", "ドイツ", "オランダ"],
    answer: 0,
    hint: "レアル・マドリードやバルセロナの国。"
  },
  {
    q: "ヨーロッパのクラブ世界一を決める大会は？",
    options: ["W杯", "UEFAチャンピオンズリーグ", "ルヴァンカップ", "天皇杯"],
    answer: 1,
    hint: "略して『CL』と呼ばれるよ。"
  },
  {
    q: "サッカーのW杯は何年ごとに開かれる？",
    options: ["毎年", "2年ごと", "4年ごと", "10年ごと"],
    answer: 2,
    hint: "オリンピックと同じ間かく。"
  },
  {
    q: "退場（たいじょう）になるカードは？",
    options: ["イエローカード", "レッドカード", "グリーンカード", "ホワイトカード"],
    answer: 1,
    hint: "2枚もらってもダメな、あの色。"
  },
  {
    q: "「セリエA」があるのはどの国？",
    options: ["スペイン", "イタリア", "フランス", "ポルトガル"],
    answer: 1,
    hint: "ピザやパスタで有名な国。"
  },
  {
    q: "「ブンデスリーガ」があるのはどの国？",
    options: ["オランダ", "ベルギー", "ドイツ", "スイス"],
    answer: 2,
    hint: "バイエルン・ミュンヘンの国。"
  },
  {
    q: "京都サンガの「サンガ」の意味に近いのは？",
    options: ["勝利", "仲間", "太陽", "金色"],
    answer: 1,
    hint: "みんなで力を合わせる、という気持ち。"
  },
  {
    q: "コーナーキックになるのはどんなとき？",
    options: ["相手がボールをゴールラインの外に出した", "ファウルをした", "オフサイド", "時間切れ"],
    answer: 0,
    hint: "守る側が最後にさわって、自分のゴール側の外に出たとき。"
  },
  {
    q: "PK（ペナルティーキック）はゴールから何メートル？",
    options: ["約7m", "約11m", "約16m", "約20m"],
    answer: 1,
    hint: "11メートルマークとも呼ばれるよ。"
  },
  {
    q: "1チームが試合中に同時にピッチに立てる人数は？",
    options: ["10人", "11人", "12人", "9人"],
    answer: 1,
    hint: "GKを入れた数。退場すると減るよ。"
  },
  {
    q: "日本代表のことを愛称で何と呼ぶ？",
    options: ["サムライブルー", "レッドドラゴンズ", "ゴールデンイーグルス", "ブルーロック"],
    answer: 0,
    hint: "青いユニフォームの侍たち。"
  },
  {
    q: "ハットトリックとは？",
    options: ["1試合で1人が3得点", "3人で点を取る", "3回パスする", "3回交代する"],
    answer: 0,
    hint: "ひとりで3点はすごい！"
  },
  {
    q: "オフサイドがあるのはどのスポーツ？",
    options: ["野球", "サッカー", "テニス", "卓球"],
    answer: 1,
    hint: "今やっているこのスポーツだよ。"
  },
  {
    q: "Jリーグで一番上のカテゴリ（ディビジョン）は？",
    options: ["J1", "J2", "J3", "JFL"],
    answer: 0,
    hint: "京都サンガもここを目指して戦う。"
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
