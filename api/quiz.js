// Vercel Serverless Function: クイズ問題の自動生成（Gemini）
// 京都サンガ・世界サッカー・ルール・戦術/戦略 のフレッシュな4択クイズをJSONで返す。
// 環境変数 GEMINI_API_KEY を使用。CDNキャッシュで生成コストを抑える。

const MODEL = "gemini-flash-latest";

const PROMPT = `小学校高学年（10〜11歳）向けの、サッカー4択クイズを {N} 問つくってください。
条件:
- テーマをバランスよく混ぜる：京都サンガF.C.、Jリーグ、海外サッカー（有名クラブ・選手・大会）、サッカーのルール、そして「戦術・戦略」（フォーメーション、プレス、カウンター、ポゼッション、オフサイドトラップ、ビルドアップ、マンマーク/ゾーンなど）。
- 戦術・戦略の問題は、子どもが読んで「なるほど！」と学べる内容にする。
- むずかしすぎず、でも歯ごたえのある“コアな”問題も入れる。難易度はバラけさせる。
- 事実は、確実で一般によく知られている内容だけ（あいまいな最新情報は避ける）。正解は1つだけ。
- 各問題に、やさしい日本語の解説(explain)を必ずつける（1〜2文、学びになるように）。
- 選択肢は必ず4つ。answerは正解の番号(0〜3)。
- 同じような問題のくり返しは避け、毎回ちがう切り口にする。`;

const schema = {
  type: "object",
  properties: {
    questions: {
      type: "array",
      items: {
        type: "object",
        properties: {
          q: { type: "string" },
          options: { type: "array", items: { type: "string" } },
          answer: { type: "integer" },
          explain: { type: "string" },
          category: { type: "string" },
        },
        required: ["q", "options", "answer", "explain"],
      },
    },
  },
  required: ["questions"],
};

module.exports = async (req, res) => {
  // 毎回ちがう問題が欲しいのでキャッシュは短め＋seedでばらつかせる
  res.setHeader("Cache-Control", "public, s-maxage=120, stale-while-revalidate=600");
  res.setHeader("Content-Type", "application/json; charset=utf-8");

  const key = process.env.GEMINI_API_KEY;
  const n = Math.min(Math.max(parseInt((req.query && req.query.n) || "8", 10) || 8, 3), 12);
  if (!key) {
    res.status(200).end(JSON.stringify({ ok: false, questions: [], error: "no-key" }));
    return;
  }
  const seed = Math.floor(Math.random() * 100000);
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${key}`;
    const r = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ role: "user", parts: [{ text: PROMPT.replace("{N}", n) + `\n(ランダムシード:${seed})` }] }],
        generationConfig: {
          temperature: 1.1,
          maxOutputTokens: 3600,
          responseMimeType: "application/json",
          responseSchema: schema,
        },
      }),
    });
    const data = await r.json();
    if (!r.ok) throw new Error((data.error && data.error.message) || `HTTP ${r.status}`);
    const text = ((((data.candidates || [])[0] || {}).content || {}).parts || []).map((p) => p.text || "").join("");
    const parsed = JSON.parse(text);
    const questions = (parsed.questions || [])
      .filter((q) => q && q.q && Array.isArray(q.options) && q.options.length === 4 &&
        Number.isInteger(q.answer) && q.answer >= 0 && q.answer < 4)
      .map((q) => ({ q: q.q, options: q.options, answer: q.answer, explain: q.explain || "", category: q.category || "サッカー" }));
    res.status(200).end(JSON.stringify({ ok: questions.length > 0, questions }));
  } catch (e) {
    res.status(200).end(JSON.stringify({ ok: false, questions: [], error: String((e && e.message) || e) }));
  }
};
