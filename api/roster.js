// Vercel Serverless Function: 京都サンガ 選手名鑑（主な選手）
// TheSportsDB から京都サンガの選手を取得して、ポジション別に整理して返す。
// ※無料APIのため一部選手。全選手は公式サイトで。

const TEAM_ID = "139888"; // Kyoto Sanga
const POS_MAP = {
  Goalkeeper: "GK", Goalkeepers: "GK",
  Defender: "DF", Defence: "DF", "Right Back": "DF", "Left Back": "DF", "Centre Back": "DF",
  Midfielder: "MF", Midfield: "MF", "Defensive Midfield": "MF", "Central Midfield": "MF",
  "Attacking Midfield": "MF", "Left Midfield": "MF", "Right Midfield": "MF", "Left Wing": "MF", "Right Wing": "MF",
  Attacker: "FW", Forward: "FW", "Centre Forward": "FW", Striker: "FW",
};
const FLAG = { Japan: "🇯🇵", Brazil: "🇧🇷", Korea: "🇰🇷", "South Korea": "🇰🇷", Spain: "🇪🇸", Germany: "🇩🇪", Brazilian: "🇧🇷" };
const ORDER = ["GK", "DF", "MF", "FW"];

const { setSecurityHeaders } = require("./_guard");

module.exports = async (req, res) => {
  setSecurityHeaders(res);
  res.setHeader("Cache-Control", "public, s-maxage=21600, stale-while-revalidate=86400");
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  try {
    const r = await fetch(`https://www.thesportsdb.com/api/v1/json/3/lookup_all_players.php?id=${TEAM_ID}`);
    const data = await r.json();
    const players = (data.player || []).map((p) => ({
      number: (p.strNumber || "").trim() || null,
      name: p.strPlayer || "",
      pos: POS_MAP[p.strPosition] || "—",
      rawPos: p.strPosition || "",
      flag: FLAG[p.strNationality] || "⚽",
      nationality: p.strNationality || "",
    })).filter((p) => p.name);

    const groups = ORDER.map((g) => ({
      pos: g,
      players: players
        .filter((p) => p.pos === g)
        .sort((a, b) => (parseInt(a.number) || 99) - (parseInt(b.number) || 99)),
    })).filter((g) => g.players.length);

    res.status(200).end(JSON.stringify({ ok: true, count: players.length, groups }));
  } catch (e) {
    res.status(200).end(JSON.stringify({ ok: false, groups: [], error: String((e && e.message) || e) }));
  }
};
