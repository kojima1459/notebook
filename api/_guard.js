// 共有ガード：API濫用（課金あらし）対策
//  - 同一IPからのリクエスト数をしぼる（サーバー側レート制限）
//  - サイト自身からのアクセスだけ通す（Origin / Referer チェック）
//  - エラーの中身はbrowserに出さず、サーバーログにだけ残す
//
// 注意: Vercelのサーバーレスは複数インスタンスに分かれるため、
//       メモリ上のレート制限は「ベストエフォート」。CDNキャッシュ＋Originチェックと
//       組み合わせて、現実的な濫用をしっかり防ぐ。さらにGoogle側で
//       APIキーに「制限」をかけるのが最強（READMEのセキュリティ項目を参照）。

// このサイトからのアクセスだけ許可するホスト名
const ALLOWED_HOST_SUFFIXES = [
  ".vercel.app", // 自分のVercelドメイン（sanga-kids / notebook どちらも）
];
const ALLOWED_EXACT_HOSTS = [
  "localhost",
  "127.0.0.1",
];

function hostFromUrl(value) {
  if (!value) return "";
  try { return new URL(value).hostname.toLowerCase(); }
  catch (e) { return ""; }
}

// Origin / Referer がサイト自身かどうか
function isAllowedOrigin(req) {
  const origin = req.headers["origin"];
  const referer = req.headers["referer"];
  // どちらのヘッダーも無い場合は通す（プライバシー設定でreferer等が落ちることがあるため）
  if (!origin && !referer) return true;
  const host = hostFromUrl(origin) || hostFromUrl(referer);
  if (!host) return true;
  if (ALLOWED_EXACT_HOSTS.includes(host)) return true;
  return ALLOWED_HOST_SUFFIXES.some((suf) => host.endsWith(suf));
}

// ---- かんたんなレート制限（IPごと・スライディングウィンドウ）----
const HITS = new Map(); // ip -> number[] (timestamps ms)
const WINDOW_MS = 60 * 1000; // 1分
const MAX_PER_WINDOW = 15;   // 1分あたり最大15回／IP

function clientIp(req) {
  const xff = req.headers["x-forwarded-for"];
  if (xff) return String(xff).split(",")[0].trim();
  return (req.socket && req.socket.remoteAddress) || "unknown";
}

function rateLimited(req) {
  const ip = clientIp(req);
  const now = Date.now();
  const arr = (HITS.get(ip) || []).filter((t) => now - t < WINDOW_MS);
  arr.push(now);
  HITS.set(ip, arr);
  // たまにメモリ掃除
  if (HITS.size > 5000) {
    for (const [k, v] of HITS) {
      if (!v.length || now - v[v.length - 1] > WINDOW_MS) HITS.delete(k);
    }
  }
  return arr.length > MAX_PER_WINDOW;
}

// 共通セキュリティヘッダー
function setSecurityHeaders(res) {
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("Referrer-Policy", "strict-origin-when-cross-origin");
  res.setHeader("X-Frame-Options", "SAMEORIGIN");
}

// ガードを通す。OKなら null、ダメなら {status, body} を返す。
function guard(req, res) {
  setSecurityHeaders(res);
  if (!isAllowedOrigin(req)) {
    return { status: 403, body: { ok: false, error: "forbidden" } };
  }
  if (rateLimited(req)) {
    res.setHeader("Retry-After", "60");
    return { status: 429, body: { ok: false, error: "rate-limited" } };
  }
  return null;
}

module.exports = { guard, setSecurityHeaders };
