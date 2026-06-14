// Vercel Serverless Function: サッカー博士チャットボット（Gemini）
// あなたの Gemini API キーをサーバー側で使って回答します。
// 環境変数 GEMINI_API_KEY を Vercel に設定してください（ブラウザには出ません）。
//
// リクエスト: POST { question, context, history }
//  - question: 子どもの質問
//  - context : いま画面で見ている内容（ページ名・ニュース見出し・順位など）
//  - history : [{role:'user'|'model', text}] 直近の会話
// Google検索グラウンディングを有効にして、最新ニュース・試合結果にも答えられるようにする。

const MODEL = "gemini-flash-latest";

const SYSTEM = `あなたは「サッカー博士」。京都サンガと世界のサッカーが大好きな小学生（10〜11歳が中心）の子ども向けサイトのアシスタントです。
ルール:
- やさしい日本語で、短く（2〜4文くらい）answerします。むずかしい漢字にはふりがなの代わりにやさしい言い回しを。
- サッカー、Jリーグ、京都サンガ、海外サッカー、選手、ルール、このサイトのニュースや順位に答えます。
- 「いま画面で見ている内容(context)」が渡されたら、それを最優先の手がかりにして答えます。
- 試合結果・得点者・順位・最新ニュースなど“最新の事実”は、必ずGoogle検索で確認してから答えます。わからない時は「公式サイトで見てみよう」と正直に言います。
- サッカーに関係ない質問や、子どもにふさわしくない話題には、やさしく「サッカーのことを聞いてね！」と返します。
- 応援したくなるような、前向きであたたかい口調で。`;

async function callGemini(key, body) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${key}`;
  const r = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return r;
}

module.exports = async (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  if (req.method !== "POST") {
    res.status(405).end(JSON.stringify({ ok: false, error: "POST only" }));
    return;
  }
  const key = process.env.GEMINI_API_KEY;
  if (!key) {
    res.status(200).end(JSON.stringify({
      ok: false,
      answer: "（セットアップ中だよ）おうちの人へ：Vercel の環境変数 GEMINI_API_KEY を設定してね。",
    }));
    return;
  }

  // リクエストボディの取得
  let payload = req.body;
  if (typeof payload === "string") { try { payload = JSON.parse(payload); } catch (e) { payload = {}; } }
  if (!payload || typeof payload !== "object") payload = {};
  const question = String(payload.question || "").slice(0, 1000);
  const context = String(payload.context || "").slice(0, 4000);
  const history = Array.isArray(payload.history) ? payload.history.slice(-6) : [];

  if (!question.trim()) {
    res.status(200).end(JSON.stringify({ ok: false, answer: "なにを聞きたいかな？サッカーのこと、聞いてね！⚽" }));
    return;
  }

  const contents = [];
  for (const h of history) {
    if (h && (h.role === "user" || h.role === "model") && h.text) {
      contents.push({ role: h.role, parts: [{ text: String(h.text).slice(0, 1500) }] });
    }
  }
  const userText = context
    ? `【いま画面で見ている内容】\n${context}\n\n【質問】\n${question}`
    : question;
  contents.push({ role: "user", parts: [{ text: userText }] });

  const baseBody = {
    system_instruction: { parts: [{ text: SYSTEM }] },
    contents,
    generationConfig: { temperature: 0.5, maxOutputTokens: 700 },
  };

  try {
    // まずGoogle検索グラウンディングつきで試す
    let r = await callGemini(key, { ...baseBody, tools: [{ google_search: {} }] });
    if (!r.ok) {
      // ツール非対応などのときはツール無しで再試行
      r = await callGemini(key, baseBody);
    }
    const data = await r.json();
    if (!r.ok) {
      const msg = (data && data.error && data.error.message) || `HTTP ${r.status}`;
      res.status(200).end(JSON.stringify({ ok: false, answer: "ごめんね、いまうまく答えられなかったよ。もう一度きいてみて！", detail: msg }));
      return;
    }
    const parts = (((data.candidates || [])[0] || {}).content || {}).parts || [];
    const answer = parts.map((p) => p.text || "").join("").trim() ||
      "うーん、うまく答えが見つからなかったよ。べつの聞き方でためしてね！";
    res.status(200).end(JSON.stringify({ ok: true, answer }));
  } catch (e) {
    res.status(200).end(JSON.stringify({ ok: false, answer: "ごめんね、通信がうまくいかなかったみたい。もう一度ためしてね。", detail: String((e && e.message) || e) }));
  }
};
