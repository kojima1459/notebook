// Vercel Serverless Function: サッカーニュース
// GoogleニュースのRSS検索をサーバー側で取得し、JSONで返す（CORS不要・安定）。
// ?cat=sanga|jleague|overseas|japan|flash

const UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36";

const QUERIES = {
  sanga: "京都サンガ",
  jleague: "Jリーグ",
  official: "Jリーグ OR 京都サンガ site:jleague.jp",
  overseas: "海外サッカー OR プレミアリーグ OR チャンピオンズリーグ OR ラリーガ",
  japan: "サッカー 日本人選手 海外 OR 日本代表",
  flash: "サッカー 速報",
};

function decode(s) {
  return s
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/<[^>]+>/g, "")
    .trim();
}

function parseRss(xml) {
  const items = [];
  const itemRe = /<item>([\s\S]*?)<\/item>/g;
  let m;
  while ((m = itemRe.exec(xml))) {
    const block = m[1];
    const pick = (tag) => {
      const r = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`).exec(block);
      return r ? r[1] : "";
    };
    const titleRaw = decode(pick("title"));
    const link = decode(pick("link"));
    const pub = pick("pubDate").trim();
    let source = decode(pick("source"));
    let title = titleRaw;
    if (!source) {
      const idx = titleRaw.lastIndexOf(" - ");
      if (idx > 0) {
        source = titleRaw.slice(idx + 3);
        title = titleRaw.slice(0, idx);
      }
    }
    if (title && link) items.push({ title, link, source, pub });
  }
  // 新しい順（pubDate降順）に並べ替え
  items.sort((a, b) => {
    const ta = new Date(a.pub).getTime() || 0;
    const tb = new Date(b.pub).getTime() || 0;
    return tb - ta;
  });
  return items.slice(0, 60);
}

module.exports = async (req, res) => {
  const cat = String((req.query && req.query.cat) || "sanga");
  const query = QUERIES[cat] || QUERIES.sanga;
  res.setHeader("Cache-Control", "public, s-maxage=600, stale-while-revalidate=3600");
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  try {
    const q = encodeURIComponent(`${query} when:21d`);
    const url = `https://news.google.com/rss/search?q=${q}&hl=ja&gl=JP&ceid=JP:ja`;
    const resp = await fetch(url, { headers: { "User-Agent": UA } });
    const xml = await resp.text();
    const items = parseRss(xml);
    res.status(200).end(JSON.stringify({ ok: true, cat, updated: new Date().toISOString(), items }));
  } catch (e) {
    res.status(200).end(JSON.stringify({ ok: false, cat, error: String((e && e.message) || e), items: [] }));
  }
};

module.exports.parseRss = parseRss;
