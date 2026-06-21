// Vercel Serverless Function: 京都サンガ 次の試合・最新結果・クラブプロフィール
// サンガ公式ニュースのタイトル（「第N節 YYYY年M月D日 ◯◯戦 試合情報/マッチレポート」）を解析して
// 次節・直近の結果カードを作る。プロフィールは確認済みの事実テキスト。

const UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36";

const PROFILE = {
  name: "京都サンガF.C.",
  hometown: "京都府（全市町村）",
  stadium: "サンガスタジアム by KYOCERA（京都府亀岡市）",
  color: "紫（パープル）",
  joinedJ: "1996年 Jリーグ加盟",
  mascot: "パーサくん・コトノ",
  note: "「サンガ」は“仲間”を意味することば。京都をホームに戦うクラブだよ。",
};

function decode(s) {
  return s
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, "&")
    .replace(/<[^>]+>/g, "").trim();
}

function parseItems(xml) {
  const items = [];
  const re = /<item>([\s\S]*?)<\/item>/g;
  let m;
  while ((m = re.exec(xml))) {
    const block = m[1];
    const t = /<title[^>]*>([\s\S]*?)<\/title>/.exec(block);
    const l = /<link[^>]*>([\s\S]*?)<\/link>/.exec(block);
    const p = /<pubDate[^>]*>([\s\S]*?)<\/pubDate>/.exec(block);
    if (t) items.push({ title: decode(t[1]), link: l ? decode(l[1]) : "", pub: p ? p[1].trim() : "" });
  }
  return items;
}

// 公式の試合タイトルを解析 → {round, date, opponent, kind}
function parseMatch(title) {
  // 例: J1百年構想 第1節 2026年6月6日 柏レイソル戦 マッチレポート
  const dm = title.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  // 日付の直後にくるチーム名（全角文字・記号もそのまま拾う）
  const om = title.match(/日\s*([^\s　0-9][^\s　]{0,15}?)戦/);
  const rm = title.match(/第(\d+)節/);
  if (!dm || !om) return null;
  const kind = /マッチレポート|試合結果|ハイライト/.test(title)
    ? "result"
    : /試合情報|プレビュー|チケット|みどころ/.test(title)
    ? "preview"
    : "other";
  const date = new Date(Number(dm[1]), Number(dm[2]) - 1, Number(dm[3]));
  return {
    round: rm ? Number(rm[1]) : null,
    dateLabel: `${dm[2]}月${dm[3]}日`,
    year: Number(dm[1]),
    ts: date.getTime(),
    opponent: om[1].replace(/Ｆ．Ｃ．?|F\.C\.?/g, "").trim(),
    kind,
    title,
  };
}

const { setSecurityHeaders } = require("./_guard");

module.exports = async (req, res) => {
  setSecurityHeaders(res);
  res.setHeader("Cache-Control", "public, s-maxage=600, stale-while-revalidate=3600");
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  try {
    const q = encodeURIComponent("京都サンガ 試合 OR 第 節 OR マッチレポート when:30d");
    const url = `https://news.google.com/rss/search?q=${q}&hl=ja&gl=JP&ceid=JP:ja`;
    const resp = await fetch(url, { headers: { "User-Agent": UA } });
    const xml = await resp.text();
    const items = parseItems(xml);

    const matches = [];
    const seen = new Set();
    for (const it of items) {
      const mm = parseMatch(it.title);
      if (!mm) continue;
      const key = `${mm.ts}-${mm.opponent}-${mm.kind}`;
      if (seen.has(key)) continue;
      seen.add(key);
      mm.link = it.link;
      matches.push(mm);
    }

    const now = Date.now();
    const dayMs = 86400000;
    // 結果：resultのうち最新
    const results = matches.filter((m) => m.kind === "result").sort((a, b) => b.ts - a.ts);
    const lastResult = results[0] || null;
    // 次の試合：previewのうち、まだ先（または直近1日内）の最も早いもの
    const upcoming = matches
      .filter((m) => m.kind === "preview" && m.ts >= now - dayMs)
      .sort((a, b) => a.ts - b.ts);
    const nextMatch = upcoming[0] || null;

    res.status(200).end(JSON.stringify({
      ok: true, updated: new Date().toISOString(),
      profile: PROFILE, nextMatch, lastResult,
    }));
  } catch (e) {
    res.status(200).end(JSON.stringify({ ok: false, profile: PROFILE, nextMatch: null, lastResult: null, error: String((e && e.message) || e) }));
  }
};

module.exports.parseMatch = parseMatch;
