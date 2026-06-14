/* =========================================================
   サンガキッズ - サッカー博士チャットボット
   画面で見ている内容（ニュース見出し・順位表など）を文脈として /api/chat に送り、
   Gemini が回答する。画面共有は不要：DOMから文脈を自動で集める。
   ========================================================= */

(function setupChat() {
  // ---- ウィジェットのDOMを生成 ----
  const root = document.createElement("div");
  root.className = "chatbot";
  root.innerHTML = `
    <button class="chat-fab" type="button" aria-label="サッカー博士に質問">💬<span>はかせに質問</span></button>
    <div class="chat-panel" hidden>
      <div class="chat-head">
        <span>⚽ サッカー博士</span>
        <button class="chat-close" type="button" aria-label="閉じる">✕</button>
      </div>
      <div class="chat-msgs" data-msgs></div>
      <form class="chat-form">
        <input class="chat-input" type="text" placeholder="サッカーのこと、聞いてね！" autocomplete="off" maxlength="300" />
        <button class="chat-send" type="submit" aria-label="送信">➤</button>
      </form>
    </div>`;
  document.body.appendChild(root);

  const fab = root.querySelector(".chat-fab");
  const panel = root.querySelector(".chat-panel");
  const closeBtn = root.querySelector(".chat-close");
  const msgsEl = root.querySelector("[data-msgs]");
  const form = root.querySelector(".chat-form");
  const input = root.querySelector(".chat-input");

  const history = [];
  let greeted = false;

  // ---- 1日の利用回数を制限（既定30回） ----
  const DAILY_LIMIT = 30;
  function todayKey() { const d = new Date(); return `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`; }
  function getUsage() {
    try { const u = JSON.parse(localStorage.getItem("sangakids_chat_usage")) || {}; return u.date === todayKey() ? u.count : 0; }
    catch (e) { return 0; }
  }
  function bumpUsage() {
    const c = getUsage() + 1;
    try { localStorage.setItem("sangakids_chat_usage", JSON.stringify({ date: todayKey(), count: c })); } catch (e) {}
    return c;
  }

  function open() {
    panel.hidden = false;
    fab.classList.add("hidden");
    if (!greeted) {
      greeted = true;
      addMsg("model", "こんにちは！⚽ サッカー博士だよ。京都サンガや世界のサッカー、いま見ているニュースのこと、なんでも聞いてね！");
    }
    setTimeout(() => input.focus(), 100);
  }
  function close() { panel.hidden = true; fab.classList.remove("hidden"); }

  fab.addEventListener("click", open);
  closeBtn.addEventListener("click", close);

  function addMsg(role, text) {
    const el = document.createElement("div");
    el.className = "chat-msg " + (role === "user" ? "me" : "bot");
    el.textContent = text;
    msgsEl.appendChild(el);
    msgsEl.scrollTop = msgsEl.scrollHeight;
    return el;
  }

  // ---- 画面の文脈を自動で集める ----
  function gatherContext() {
    const parts = ["ページ: " + document.title];
    const heads = Array.from(document.querySelectorAll(".feed-title"))
      .slice(0, 12).map((e) => e.textContent.trim()).filter(Boolean);
    if (heads.length) parts.push("画面に出ているニュース見出し:\n- " + heads.join("\n- "));
    const sum = document.querySelector(".sanga-summary");
    if (sum) parts.push("京都サンガの順位: " + sum.textContent.replace(/\s+/g, " ").trim());
    const stand = document.querySelector("[data-standings]");
    if (stand) {
      const rows = Array.from(stand.querySelectorAll(".stand-row"))
        .slice(0, 10).map((r) => r.textContent.replace(/\s+/g, " ").trim());
      if (rows.length) parts.push("順位表(一部): " + rows.join(" / "));
    }
    const match = document.querySelector("[data-sanga-match]");
    if (match && match.textContent.trim()) parts.push("試合情報: " + match.textContent.replace(/\s+/g, " ").trim());
    return parts.join("\n");
  }

  async function send(question) {
    if (getUsage() >= DAILY_LIMIT) {
      addMsg("user", question);
      addMsg("bot", "今日はたくさんお話できたね！⚽ また明日きいてね（1日30回まで）。");
      return;
    }
    bumpUsage();
    addMsg("user", question);
    history.push({ role: "user", text: question });
    const typing = addMsg("bot", "考え中… ⚽");
    typing.classList.add("typing");
    try {
      const res = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ question, context: gatherContext(), history: history.slice(0, -1) }),
      });
      const data = await res.json();
      const answer = data.answer || "うまく答えられなかったよ。もう一度ためしてね！";
      typing.classList.remove("typing");
      typing.textContent = answer;
      history.push({ role: "model", text: answer });
    } catch (e) {
      typing.classList.remove("typing");
      typing.textContent = "ごめんね、つながらなかったみたい。もう一度ためしてね。";
    }
    msgsEl.scrollTop = msgsEl.scrollHeight;
  }

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const q = input.value.trim();
    if (!q) return;
    input.value = "";
    send(q);
  });
})();
