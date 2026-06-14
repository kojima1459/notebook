// Vercel Serverless Function: J1リーグ順位表
// jleague.jp の公式順位表（2026 百年構想リーグ＝EAST/WESTグループ制）をサーバー側で取得・解析してJSONで返す。
// CDNキャッシュ(s-maxage)で上流アクセスを最小化。

const SOURCE = "https://www.jleague.jp/standings/j1/";
const UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36";

function stripTags(s) {
  return s.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
}

// 1チーム分の行をパース
function parseRow(rowHtml) {
  const tds = [];
  const re = /<td[^>]*>([\s\S]*?)<\/td>/g;
  let m;
  while ((m = re.exec(rowHtml))) tds.push(m[1]);
  if (tds.length < 12) return null;

  // チーム名は span.embS の中身（重複表記を避ける）
  const span = tds[2].match(/<span[^>]*class="[^"]*emb[^"]*"[^>]*>([\s\S]*?)<\/span>/i);
  const team = span ? stripTags(span[1]) : stripTags(tds[2]);

  // 直近5試合の勝敗アイコン（win/lose/draw）を拾う
  const formIcons = (tds[12] || "").match(/ico_match_(win|lose|draw)/g) || [];
  const form = formIcons.map((c) =>
    c.includes("win") ? "W" : c.includes("lose") ? "L" : "D"
  );

  const num = (i) => {
    const v = stripTags(tds[i] || "");
    const n = parseInt(v.replace(/[^\-0-9]/g, ""), 10);
    return isNaN(n) ? null : n;
  };

  return {
    rank: num(1),
    team,
    points: num(3),
    played: num(4),
    win: num(5),
    pkWin: num(6),
    pkLose: num(7),
    lose: num(8),
    goalsFor: num(9),
    goalsAgainst: num(10),
    goalDiff: num(11),
    form,
    isSanga: team.includes("京都") || team.includes("サンガ"),
  };
}

function parseStandings(html) {
  const groups = [];
  // EAST / WEST の見出し＋テーブルを順に拾う
  const groupRe = /<h4[^>]*>(EAST|WEST)<\/h4>[\s\S]*?<table[^>]*J1table[^>]*>([\s\S]*?)<\/table>/g;
  let g;
  while ((g = groupRe.exec(html))) {
    const name = g[1];
    const tableHtml = g[2];
    const rows = [];
    const rowRe = /<tr[^>]*class="row[^"]*"[^>]*>([\s\S]*?)<\/tr>/g;
    let r;
    while ((r = rowRe.exec(tableHtml))) {
      const row = parseRow(r[1]);
      if (row && row.team) rows.push(row);
    }
    if (rows.length) groups.push({ name, rows });
  }
  // グループ見出しが無い構成（通常の1テーブル）にもフォールバック
  if (!groups.length) {
    const t = html.match(/<table[^>]*J1table[^>]*>([\s\S]*?)<\/table>/);
    if (t) {
      const rows = [];
      const rowRe = /<tr[^>]*class="row[^"]*"[^>]*>([\s\S]*?)<\/tr>/g;
      let r;
      while ((r = rowRe.exec(t[1]))) {
        const row = parseRow(r[1]);
        if (row && row.team) rows.push(row);
      }
      if (rows.length) groups.push({ name: "J1", rows });
    }
  }
  return groups;
}

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "public, s-maxage=900, stale-while-revalidate=86400");
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  try {
    const resp = await fetch(SOURCE, { headers: { "User-Agent": UA } });
    const buf = Buffer.from(await resp.arrayBuffer());
    const html = buf.toString("utf-8");
    const groups = parseStandings(html);
    if (!groups.length) throw new Error("parse-empty");
    res.status(200).end(
      JSON.stringify({ ok: true, updated: new Date().toISOString(), source: SOURCE, groups })
    );
  } catch (e) {
    res.status(200).end(
      JSON.stringify({ ok: false, error: String(e && e.message || e), groups: [], source: SOURCE })
    );
  }
};

// テスト用にパーサを公開
module.exports.parseStandings = parseStandings;
