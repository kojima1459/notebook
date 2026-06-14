/* =========================================================
   サンガキッズ - サッカークイズ
   - 京都サンガ・世界サッカー・ルール・戦術/戦略
   - 出題は毎回シャッフル。AIがつくるフレッシュ問題も混ぜる。
   - 解説つきで、遊びながら学べる。
   ========================================================= */

// ---- ベース問題バンク（確かな事実・解説つき） ----
const STATIC_BANK = [
  // 京都サンガ
  { q: "京都サンガF.C. のチームカラーは？", options: ["赤", "緑", "紫（パープル）", "青"], answer: 2, explain: "サンガのシンボルカラーは紫。スタジアムが紫に染まるよ。", cat: "サンガ" },
  { q: "京都サンガのホームスタジアムがある市は？", options: ["京都市", "亀岡市", "大阪市", "宇治市"], answer: 1, explain: "亀岡市の『サンガスタジアム by KYOCERA』が本拠地だよ。", cat: "サンガ" },
  { q: "京都サンガの公式マスコット（鳳凰の男の子）は？", options: ["パーサくん", "サンガくん", "コトノちゃん", "しかお"], answer: 0, explain: "鳳凰がモチーフのパーサくん。仲間のコトノちゃんもいるよ。", cat: "サンガ" },
  { q: "「サンガ」の意味に近いのは？", options: ["勝利", "仲間", "太陽", "金色"], answer: 1, explain: "サンガは“仲間・集まり”を表すことば。みんなで戦うクラブ。", cat: "サンガ" },
  { q: "サンガスタジアム by KYOCERA の特長は？", options: ["屋内コート", "ピッチと客席がとても近い", "野球もできる", "海の上"], answer: 1, explain: "サッカー専用で、選手との距離が近い迫力スタジアム。", cat: "サンガ" },

  // ルール
  { q: "1チームが同時にピッチに立てる人数は？", options: ["10人", "11人", "12人", "9人"], answer: 1, explain: "GKを入れて11人。退場すると人数が減るよ。", cat: "ルール" },
  { q: "一発退場になるカードは？", options: ["イエロー", "レッド", "グリーン", "ホワイト"], answer: 1, explain: "レッドカードは退場。イエロー2枚でもレッドになるよ。", cat: "ルール" },
  { q: "ボールがタッチラインを出たとき、手で投げ入れるプレーは？", options: ["フリーキック", "コーナーキック", "スローイン", "ゴールキック"], answer: 2, explain: "横の線から出たらスローイン。両手で頭の後ろから投げるよ。", cat: "ルール" },
  { q: "PK（ペナルティーキック）はゴールから約何メートル？", options: ["約7m", "約11m", "約16m", "約20m"], answer: 1, explain: "11mマークから蹴る、GKと1対1の大チャンス。", cat: "ルール" },
  { q: "守る側が最後に触ってゴールラインの外に出すと？", options: ["コーナーキック", "スローイン", "ゴールキック", "PK"], answer: 0, explain: "攻める側にコーナーキックが与えられるよ。", cat: "ルール" },
  { q: "オフサイドがあるスポーツは？", options: ["野球", "サッカー", "テニス", "卓球"], answer: 1, explain: "オフサイドはサッカーの大事なルール。待ち伏せを防ぐよ。", cat: "ルール" },
  { q: "ハットトリックとは？", options: ["1人が1試合3得点", "3人で得点", "3回パス", "3回交代"], answer: 0, explain: "ひとりで3ゴール！とてもすごい記録だよ。", cat: "ルール" },

  // 戦術・戦略
  { q: "相手に近づいて強くボールを奪いにいく守り方は？", options: ["プレス", "ポゼッション", "クリア", "スローイン"], answer: 0, explain: "プレス（プレッシング）。早く奪うほどチャンスになる。", cat: "戦術" },
  { q: "ボールを奪って一気に速く攻めることを何という？", options: ["カウンター", "ビルドアップ", "リフティング", "クリア"], answer: 0, explain: "カウンター（速攻）。相手が戻る前にゴールをねらう。", cat: "戦術" },
  { q: "味方でボールをたくさん回して試合を支配する戦い方は？", options: ["ポゼッション", "ロングボール", "マンマーク", "クリア"], answer: 0, explain: "ポゼッション（ボール保持）。パスで相手を動かすよ。", cat: "戦術" },
  { q: "「4-4-2」などの数字が表すのは？", options: ["選手の身長", "フォーメーション（並び方）", "背番号", "得点数"], answer: 1, explain: "DF-MF-FWの人数の並び。作戦で形が変わるよ。", cat: "戦術" },
  { q: "相手のFWを一人ひとり担当して守るのは？", options: ["マンマーク", "ゾーン", "カウンター", "プレス"], answer: 0, explain: "マンマークは“人につく”守り。エースを止めたい時に使う。", cat: "戦術" },
  { q: "DFが同時に上がって相手を罠にかける守備は？", options: ["オフサイドトラップ", "ビルドアップ", "ドリブル", "スローイン"], answer: 0, explain: "ラインを上げて相手をオフサイドにする高度な作戦。", cat: "戦術" },
  { q: "後ろから丁寧にパスをつないで攻撃を組み立てるのは？", options: ["ビルドアップ", "クリア", "ロングスロー", "カウンター"], answer: 0, explain: "GKやDFから始める攻撃の組み立て。落ち着きが大事。", cat: "戦術" },
  { q: "ピンチのとき遠くへ大きく蹴り出して危険を消すのは？", options: ["クリア", "ドリブル", "フェイント", "トラップ"], answer: 0, explain: "まず安全に！クリアでピンチをしのぐ守備の基本。", cat: "戦術" },

  // 世界のサッカー
  { q: "2022年W杯（カタール）で優勝した国は？", options: ["フランス", "ブラジル", "アルゼンチン", "ドイツ"], answer: 2, explain: "メッシのいるアルゼンチンが優勝したよ。", cat: "世界" },
  { q: "リオネル・メッシの出身国は？", options: ["スペイン", "アルゼンチン", "ブラジル", "ポルトガル"], answer: 1, explain: "南米アルゼンチン出身。世界最高の選手のひとり。", cat: "世界" },
  { q: "クリスティアーノ・ロナウドの出身国は？", options: ["スペイン", "ブラジル", "ポルトガル", "イタリア"], answer: 2, explain: "ポルトガル代表のスーパースター。", cat: "世界" },
  { q: "「プレミアリーグ」がある国は？", options: ["スペイン", "イングランド", "ドイツ", "フランス"], answer: 1, explain: "世界一人気とも言われるイングランドのリーグ。", cat: "世界" },
  { q: "「ラ・リーガ」がある国は？", options: ["スペイン", "イタリア", "ドイツ", "オランダ"], answer: 0, explain: "レアル・マドリードやバルセロナの国、スペイン。", cat: "世界" },
  { q: "「セリエA」がある国は？", options: ["スペイン", "イタリア", "フランス", "ポルトガル"], answer: 1, explain: "イタリアのトップリーグ。守備がうまいチームが多い。", cat: "世界" },
  { q: "「ブンデスリーガ」がある国は？", options: ["オランダ", "ベルギー", "ドイツ", "スイス"], answer: 2, explain: "ドイツのリーグ。日本人選手も多く活躍してきたよ。", cat: "世界" },
  { q: "ヨーロッパのクラブ世界一を決める大会は？", options: ["W杯", "UEFAチャンピオンズリーグ", "天皇杯", "ルヴァンカップ"], answer: 1, explain: "略して『CL』。強豪クラブが集まる大舞台。", cat: "世界" },
  { q: "W杯は何年ごとに開かれる？", options: ["毎年", "2年ごと", "4年ごと", "10年ごと"], answer: 2, explain: "4年に一度。オリンピックと同じ間かくだよ。", cat: "世界" },
  { q: "日本代表の愛称は？", options: ["サムライブルー", "レッドドラゴンズ", "ゴールデンイーグルス", "ブルーロック"], answer: 0, explain: "青いユニフォームの侍たち＝サムライブルー。", cat: "世界" },

  // Jリーグ
  { q: "Jリーグで一番上のカテゴリは？", options: ["J1", "J2", "J3", "JFL"], answer: 0, explain: "J1が最上位。サンガもここで戦うよ。", cat: "Jリーグ" },
  { q: "手を使ってボールを止めてよい選手は？", options: ["FW", "ゴールキーパー", "審判", "監督"], answer: 1, explain: "GKだけが自陣のペナルティーエリア内で手を使えるよ。", cat: "ルール" }
];

function shuffle(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

// AIのフレッシュ問題を取得（失敗してもOK）
async function fetchFreshQuestions() {
  try {
    const res = await fetch("/api/quiz?n=6", { cache: "no-store" });
    const data = await res.json();
    if (data.ok && Array.isArray(data.questions)) {
      return data.questions.map((q) => ({ q: q.q, options: q.options, answer: q.answer, explain: q.explain, cat: q.category || "AI" }));
    }
  } catch (e) {}
  return [];
}

// 出題セットをつくる（フレッシュ問題＋ベース問題をシャッフルして12問）
async function buildQuizSet() {
  const fresh = await fetchFreshQuestions();
  const pool = shuffle(STATIC_BANK);
  const set = shuffle([...fresh, ...pool]).slice(0, 12);
  return set.length ? set : shuffle(STATIC_BANK).slice(0, 12);
}

(async function setupQuiz() {
  const root = document.querySelector("[data-quiz]");
  if (!root) return;

  const qEl = root.querySelector(".quiz-question");
  const optsEl = root.querySelector(".quiz-options");
  const fbEl = root.querySelector(".quiz-feedback");
  const nextBtn = root.querySelector(".quiz-next");
  const progEl = root.querySelector(".quiz-progress");
  const barEl = root.querySelector(".quiz-bar > i");

  progEl.textContent = "問題をよみこみ中… ⚽";
  let QUIZ = await buildQuizSet();

  let current = 0, score = 0, answered = false;

  function render() {
    answered = false;
    const item = QUIZ[current];
    progEl.textContent = `第 ${current + 1} 問 / ${QUIZ.length} 問${item.cat ? "（" + item.cat + "）" : ""}`;
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
      fbEl.innerHTML = `せいかい！🎉<br><span class="quiz-explain">💡 ${item.explain || ""}</span>`;
      fbEl.style.color = "#1f9d6b";
      score++;
    } else {
      btn.classList.add("wrong");
      buttons[item.answer].classList.add("correct");
      fbEl.innerHTML = `ざんねん…<br><span class="quiz-explain">💡 ${item.explain || ""}</span>`;
      fbEl.style.color = "#d6336c";
    }
    nextBtn.disabled = false;
  }

  function showResult() {
    barEl.style.width = "100%";
    progEl.textContent = "おわり！";
    const ratio = score / QUIZ.length;
    let medal = "🎽", msg = "またチャレンジしてね！";
    if (ratio === 1) { medal = "🏆"; msg = "パーフェクト！サッカー博士だ！"; }
    else if (ratio >= 0.7) { medal = "🥇"; msg = "すごい！かなりの実力者！"; }
    else if (ratio >= 0.4) { medal = "🥈"; msg = "いいね！あと少し！"; }
    else { medal = "🥉"; msg = "これからもっとくわしくなろう！"; }

    root.innerHTML = `
      <div class="quiz-result">
        <div class="medal">${medal}</div>
        <p>きみの スコアは…</p>
        <div class="score">${score} / ${QUIZ.length}</div>
        <p>${msg}</p>
        <button class="quiz-next" type="button" style="max-width:300px;margin:18px auto 0;">
          新しい問題でもう一度 🔄
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
