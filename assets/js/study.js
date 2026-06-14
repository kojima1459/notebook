/* =========================================================
   サンガキッズ - サッカーのべんきょう（座学・あたまを鍛える）
   歴史 / ルール / ポジション / フォーメーション / 動き方 /
   戦術 / テクニックの理論 / フィジカルの大切さ
   各トピックから「練習」への動線をはる。
   ========================================================= */

/* フォーメーション図のデータ（DF→MF→FWの人数。GKは別で1人） */
const FORMATIONS = [
  {
    name: "4-4-2", nick: "バランス型",
    df: 4, mf: 4, fw: 2,
    desc: "まもりと せめが バランスよし。むかしから 人気の形。れつがそろっていて、まもりやすい。はじめて おぼえるなら これ！",
  },
  {
    name: "4-3-3", nick: "こうげき型",
    df: 4, mf: 3, fw: 3,
    desc: "前に 3人いるので せめが強い。両サイドの FW（ウイング）が ドリブルで しかける。今の世界の トップチームに 多い形。",
  },
  {
    name: "3-4-3", nick: "せめ重視",
    df: 3, mf: 4, fw: 3,
    desc: "まもりを3人にして、前に 人数をかける ぜめの形。サイドを 上下に走る 選手が だいかつやく。スタミナが いるよ。",
  },
];

/* 各トピック（アコーディオン） */
const TOPICS = [
  {
    id: "history", icon: "🌍", title: "サッカーの歴史",
    lead: "サッカーって いつ・どこで 生まれたの？",
    blocks: [
      { h: "イングランドで ルールができた", p: "ボールをけるあそびは 大むかしから 世界中にあったけれど、今のサッカーの ルールは 1863年にイングランド（イギリス）で 決められたよ。「手を使わない」「11人でやる」などが この時に まとまったんだ。" },
      { h: "ワールドカップと FIFA", p: "世界の国が きそう大会が ワールドカップ（W杯）。4年に1度ひらかれ、世界中が もりあがる お祭りだよ。サッカーを まとめている 世界の団体が FIFA（フィファ）。" },
      { h: "日本とJリーグ・京都サンガ", p: "日本のプロリーグ「Jリーグ」は 1993年にスタート。京都サンガF.C. も Jリーグで 戦うチームだよ。日本代表も W杯に 出場する 強いチームに なってきたんだ。" },
    ],
    practice: { text: "🎮 歴史クイズで力だめし →", href: "quiz.html" },
  },
  {
    id: "rules", icon: "📏", title: "ルールと きほんの考え方",
    lead: "なぜ この ルールが あるの？を知ると、見るのも たのしい！",
    blocks: [
      { h: "11人 × 90分", p: "1チーム 11人（その中の1人が ゴールキーパー）。試合は ぜんはん・こうはん 45分ずつ＝90分。手を使えるのは キーパーだけ（自分のペナルティエリア内）。" },
      { h: "オフサイドって？", p: "あいてゴール近くで、ボールより前・なおかつ あいての うしろから2番目の選手より前で パスを もらうと「オフサイド」で 反則。これがあるから みんな 前に かたまらず、いいタイミングで 動く くふうをするんだ。" },
      { h: "ファウルとカード", p: "あぶないプレーや 手で つかむのは ファウル。わるいファウルは「イエローカード（けいこく）」、とても わるいと「レッドカード（たいじょう）」。フェアプレーが いちばん かっこいい！" },
    ],
    practice: { text: "📺 試合を見て ルールをかくにん →", href: "shorts.html" },
  },
  {
    id: "positions", icon: "🧭", title: "ポジションと 役わり",
    lead: "サッカーには いろんな「お仕事」がある。きみは どこ向き？",
    blocks: [
      { h: "GK（ゴールキーパー）", p: "ゴールを まもる さいごのとりで。手を使える ゆいいつの ポジション。声で みんなを うごかす リーダーでもあるよ。" },
      { h: "DF（ディフェンダー）", p: "あいてのこうげきを 止める まもりの選手。ボールを うばって 味方に つなぐ。体の強さと、よむ力が 大事。" },
      { h: "MF（ミッドフィルダー）", p: "まん中で まもりと せめを つなぐ「チームのエンジン」。たくさん走って、パスを くばる。サッカーIQ（あたま）が いちばん 出る ポジション。" },
      { h: "FW（フォワード）", p: "ゴールを 決める せめの先頭。シュートの うまさ・スピード・うらへ ぬけ出す かしこさが 武器。" },
    ],
    practice: { text: "🔥 ポジションに役立つ練習へ →", href: "training.html#skills" },
  },
  {
    id: "formations", icon: "🔢", title: "フォーメーション（ならび方）",
    lead: "数字は うしろ（DF）→前（FW）の 人数。図で見てみよう！",
    formations: true,
    blocks: [
      { h: "数字の よみ方", p: "「4-4-2」なら、まもり4人・まん中4人・前2人（＋キーパー1人）という意味。チームの 作せん によって、ならび方を かえるんだ。" },
      { h: "どれが いいの？", p: "「これが せいかい」は ないよ。あいてや 自分たちの とくいに 合わせて えらぶ。まずは バランスのいい 4-4-2 から おぼえると わかりやすい！" },
    ],
    practice: { text: "🏃 動き方のページも 見てみよう →", href: "#movement" },
  },
  {
    id: "movement", icon: "🏃", title: "動き方（ボールを持ってない時）",
    lead: "じつは 試合の 9わりは「ボールを 持ってない 時間」なんだ！",
    blocks: [
      { h: "スペースを さがす", p: "あいてが いない場所＝スペース。そこへ 走りこむと、パスを もらいやすい。「ボールを 見るだけ」じゃなく、あいた所を さがして 動こう。" },
      { h: "もらう前に 体の向きを作る", p: "パスを もらう前に、半身（はんみ）に なって 前を見ておくと、もらってすぐ 前へ いける。これを「準備（じゅんび）」というよ。じょうずな選手は みんな やってる。" },
      { h: "サポート＝味方を助ける位置", p: "ボールを持った 味方の となりや ななめに 動いて、パスの にがし道を 作ってあげる。三角形（トライアングル）を 作ると パスが まわりやすい。" },
    ],
    practice: { text: "🧱 かべパス＆ダッシュで練習 →", href: "training.html#out-drill" },
  },
  {
    id: "tactics", icon: "♟️", title: "戦術・作せん",
    lead: "チームみんなで「どう せめる・どう まもる」を きめる ことが戦術。",
    blocks: [
      { h: "プレス（ボールを とりにいく）", p: "あいてが ボールを持ったら、みんなで すばやく かこんで うばう。早く とれば、あいてゴール近くで せめられる。" },
      { h: "ビルドアップ（つないで運ぶ）", p: "うしろから パスを つないで、あわてず ボールを 前へ 運ぶこと。あいてを ひきつけて、あいたスペースを 使うのが コツ。" },
      { h: "カウンター（すばやい はんげき）", p: "ボールを うばったら、あいてが もどる前に いっきに せめる。スピードが ぶき。守ってから 速く せめる チームは こわいよ！" },
      { h: "セットプレー", p: "コーナーキックや フリーキックなど、止まった ボールから せめる場面。サインを 決めておくと チャンスに なる。" },
    ],
    practice: { text: "⚽ ミニゲームで ためそう →", href: "training.html#out-drill" },
  },
  {
    id: "technic", icon: "🎯", title: "テクニックの 理ゆう",
    lead: "「なぜ そうやるの？」が わかると、れんしゅうの 質が上がる！",
    blocks: [
      { h: "なぜ インサイドパス？", p: "足の内がわ（インサイド）は たいらで 広いから、ボールに 当てやすく 正かく。だから 近くへの パスは ほとんど インサイドを 使うんだ。" },
      { h: "なぜ トラップで 足を引く？", p: "ボールが 当たる しゅんかんに 足を そっと 引くと、ボールの いきおいを 消せる。だから 足もとに ピタッと 止まる。かべに 当たると はね返るのと 反対だね。" },
      { h: "軸足（じくあし）の 向きが 大事", p: "ける時、ボールの横に おく じく足の つま先を、飛ばしたい方向へ 向けると、ボールも そっちへ いく。シュートが ずれる時は たいてい じく足の 向きが げんいん。" },
      { h: "体の向き（半身）", p: "正面を むくより、ななめ（半身）に かまえると、前も後ろも 見えて、つぎのプレーが 速くなる。うまい選手は つねに 半身。" },
    ],
    practice: { text: "⭐ 技の 正しいやり方へ →", href: "training.html#skills" },
  },
  {
    id: "physical", icon: "💪", title: "フィジカル（体づくり）の 大切さ",
    lead: "技を「ほんとうに 使える」ようにするのが、強い体。",
    blocks: [
      { h: "なぜ 体づくり？", p: "どんなに うまい技も、あたりに 負けたり、つかれて 動けなくなったら 使えない。強い体は、技を 試合で 出すための「土台」なんだ。" },
      { h: "スピードと アジリティ", p: "速く 走る力＋ すばやく 向きをかえる力（アジリティ）が あると、あいてより 先に ボールに とどく。1対1にも 強くなる。" },
      { h: "体幹（たいかん）", p: "おなか・せなかの まん中の力。これが 強いと、あたられても ブレない・たおれない。シュートや パスも 安定するよ。" },
      { h: "スタミナ", p: "90分 走りきる 力。後半に つかれて 足が止まると、いいプレーが できない。走れる選手は それだけで チームを 助ける。" },
    ],
    practice: { text: "💪 フィジカル練習へ →", href: "training.html#physical" },
  },
];

(function renderStudy() {
  const wrap = document.querySelector("[data-study]");
  if (!wrap) return;

  function pitchHtml(f) {
    const row = (n, cls, lbl) =>
      `<div class="pitch-row ${cls}">${Array.from({ length: n })
        .map(() => '<span class="dot"></span>').join("")}<span class="pitch-lbl">${lbl}</span></div>`;
    return `
      <div class="formation">
        <div class="formation-name">${f.name} <small>${f.nick}</small></div>
        <div class="pitch">
          ${row(f.fw, "fw", "FW")}
          ${row(f.mf, "mf", "MF")}
          ${row(f.df, "df", "DF")}
          ${row(1, "gk", "GK")}
        </div>
        <p class="formation-desc">${f.desc}</p>
      </div>`;
  }

  wrap.innerHTML = TOPICS.map((t, i) => {
    const body = t.blocks.map((b) => `<div class="study-block"><b>${b.h}</b><p>${b.p}</p></div>`).join("");
    const forms = t.formations ? `<div class="formation-grid">${FORMATIONS.map(pitchHtml).join("")}</div>` : "";
    const practice = t.practice
      ? `<a class="study-go" href="${t.practice.href}">${t.practice.text}</a>` : "";
    return `
      <div class="study-card" id="${t.id}">
        <button class="study-head" type="button" aria-expanded="false" data-study-head="${i}">
          <span class="study-icon" aria-hidden="true">${t.icon}</span>
          <span class="study-title"><strong>${t.title}</strong><small>${t.lead}</small></span>
          <span class="study-arrow" aria-hidden="true">▼</span>
        </button>
        <div class="study-body" hidden>
          ${body}
          ${forms}
          ${practice}
        </div>
      </div>`;
  }).join("");

  wrap.querySelectorAll(".study-head").forEach((btn) => {
    btn.addEventListener("click", () => {
      const body = btn.closest(".study-card").querySelector(".study-body");
      const open = btn.getAttribute("aria-expanded") === "true";
      btn.setAttribute("aria-expanded", String(!open));
      body.hidden = open;
      btn.querySelector(".study-arrow").textContent = open ? "▼" : "▲";
    });
  });

  // ページ内リンク（#movement など）で来たら そのカードを自動でひらく
  function openFromHash() {
    const id = location.hash.replace("#", "");
    if (!id) return;
    const card = document.getElementById(id);
    if (card && card.classList.contains("study-card")) {
      const btn = card.querySelector(".study-head");
      const body = card.querySelector(".study-body");
      if (btn && body && body.hidden) {
        btn.setAttribute("aria-expanded", "true");
        body.hidden = false;
        btn.querySelector(".study-arrow").textContent = "▲";
      }
      card.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }
  window.addEventListener("hashchange", openFromHash);
  openFromHash();
})();
