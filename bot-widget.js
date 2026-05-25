/* ============================================================
   Pascal Press — Copilot Bot Widget
   Bottom-right floating chat panel
   Copilot Studio embed will replace sendToBot() when ready.
   ============================================================ */
(function () {
  // ── Inject CSS ──────────────────────────────────────────────
  const style = document.createElement('style');
  style.textContent = `
    /* Move toasts up so they don't sit behind the bot button */
    #toast-container {
      bottom: 100px !important;
    }

    /* Toggle button — bottom RIGHT */
    #pp-bot-toggle {
      position: fixed; bottom: 28px; right: 28px; z-index: 9990;
      width: 56px; height: 56px; border-radius: 50%;
      background: linear-gradient(135deg, #00a4a6, #008385);
      border: none; cursor: pointer;
      box-shadow: 0 4px 16px rgba(0,164,166,.45);
      display: flex; align-items: center; justify-content: center;
      transition: transform .2s ease, box-shadow .2s ease;
      color: white; font-size: 1.4rem;
    }
    #pp-bot-toggle:hover { transform: scale(1.08); box-shadow: 0 6px 20px rgba(0,164,166,.55); }
    #pp-bot-toggle .bot-notif {
      position: absolute; top: 2px; right: 2px;
      width: 14px; height: 14px; border-radius: 50%;
      background: #f4851f; border: 2px solid white;
    }

    /* Chat panel — anchored bottom RIGHT */
    #pp-bot-panel {
      position: fixed; bottom: 96px; right: 28px; z-index: 9989;
      width: 340px; height: 480px;
      background: white; border-radius: 14px;
      box-shadow: 0 12px 40px rgba(30,58,95,.22);
      display: flex; flex-direction: column;
      overflow: hidden;
      transform: scale(.92) translateY(16px);
      opacity: 0; pointer-events: none;
      transition: transform .22s ease, opacity .22s ease;
      border: 1px solid #d8dde6;
    }
    #pp-bot-panel.open {
      transform: scale(1) translateY(0);
      opacity: 1; pointer-events: all;
    }

    /* Header */
    #pp-bot-panel .bot-header {
      background: linear-gradient(135deg, #1e3a5f 0%, #2a4d7a 100%);
      padding: 14px 16px;
      display: flex; align-items: center; gap: 10px;
      flex-shrink: 0;
    }
    .bot-header .bot-avatar {
      width: 36px; height: 36px; border-radius: 50%;
      background: #00a4a6;
      display: flex; align-items: center; justify-content: center;
      font-size: .95rem; color: white; flex-shrink: 0;
      box-shadow: 0 2px 8px rgba(0,164,166,.4);
    }
    .bot-header .bot-info { flex: 1; }
    .bot-header .bot-info .bot-name { font-size: .88rem; font-weight: 700; color: white; }
    .bot-header .bot-info .bot-status {
      font-size: .72rem; color: rgba(255,255,255,.6);
      display: flex; align-items: center; gap: 4px;
    }
    .bot-header .bot-info .bot-status::before {
      content: ''; width: 6px; height: 6px; border-radius: 50%;
      background: #27ae60; display: inline-block;
    }
    .bot-header .bot-close {
      background: none; border: none; color: rgba(255,255,255,.6);
      cursor: pointer; font-size: 1rem; padding: 4px;
      border-radius: 6px; transition: color .15s;
      line-height: 1;
    }
    .bot-header .bot-close:hover { color: white; background: rgba(255,255,255,.1); }

    /* Messages area */
    #pp-bot-messages {
      flex: 1; overflow-y: auto; padding: 14px 12px;
      display: flex; flex-direction: column; gap: 10px;
      background: #f5f7fa;
    }
    #pp-bot-messages::-webkit-scrollbar { width: 4px; }
    #pp-bot-messages::-webkit-scrollbar-thumb { background: #d8dde6; border-radius: 4px; }

    /* Message bubbles */
    .bot-msg { display: flex; gap: 7px; align-items: flex-end; }
    .bot-msg.user { flex-direction: row-reverse; }
    .bot-msg .msg-avatar {
      width: 26px; height: 26px; border-radius: 50%;
      background: #00a4a6; flex-shrink: 0;
      display: flex; align-items: center; justify-content: center;
      font-size: .65rem; color: white; font-weight: 700;
    }
    .bot-msg.user .msg-avatar { background: #1e3a5f; }
    .bot-msg .bubble {
      max-width: 230px; padding: 9px 12px;
      border-radius: 14px; font-size: .83rem; line-height: 1.45;
      color: #2d3a4a;
    }
    .bot-msg .bubble { background: white; border: 1px solid #edf0f4; border-bottom-left-radius: 4px; box-shadow: 0 1px 3px rgba(30,58,95,.06); }
    .bot-msg.user .bubble { background: #1e3a5f; color: white; border: none; border-bottom-right-radius: 4px; }
    .bot-msg .msg-time { font-size: .65rem; color: #9ba5b4; padding: 0 2px; }

    /* Typing indicator */
    .typing-indicator { display: flex; gap: 4px; align-items: center; padding: 4px 2px; }
    .typing-indicator span {
      width: 7px; height: 7px; border-radius: 50%; background: #9ba5b4;
      animation: bounce .9s infinite ease-in-out;
    }
    .typing-indicator span:nth-child(2) { animation-delay: .15s; }
    .typing-indicator span:nth-child(3) { animation-delay: .3s; }
    @keyframes bounce { 0%,80%,100% { transform: translateY(0); } 40% { transform: translateY(-6px); } }

    /* Quick replies */
    .quick-replies { display: flex; flex-wrap: wrap; gap: 6px; padding: 0 0 2px 33px; }
    .quick-reply-btn {
      background: white; border: 1.5px solid #00a4a6; color: #00a4a6;
      border-radius: 20px; padding: 4px 12px; font-size: .78rem; font-weight: 600;
      cursor: pointer; transition: all .15s;
    }
    .quick-reply-btn:hover { background: #00a4a6; color: white; }

    /* Input area */
    #pp-bot-panel .bot-input-area {
      padding: 10px 12px; border-top: 1px solid #edf0f4;
      display: flex; gap: 8px; align-items: center; flex-shrink: 0;
      background: white;
    }
    .bot-input-area input {
      flex: 1; border: 1.5px solid #d8dde6; border-radius: 20px;
      padding: 8px 14px; font-size: .83rem; outline: none;
      background: #f5f7fa; transition: border-color .15s;
    }
    .bot-input-area input:focus { border-color: #00a4a6; background: white; }
    .bot-input-area .send-btn {
      width: 36px; height: 36px; border-radius: 50%;
      background: #00a4a6; border: none; cursor: pointer;
      display: flex; align-items: center; justify-content: center;
      color: white; font-size: .9rem; transition: background .15s; flex-shrink: 0;
    }
    .bot-input-area .send-btn:hover { background: #008385; }
    .bot-input-area .send-btn:disabled { background: #d8dde6; cursor: not-allowed; }

    /* Powered by */
    .bot-powered {
      text-align: center; padding: 6px; font-size: .68rem;
      color: #9ba5b4; background: white; border-top: 1px solid #edf0f4; flex-shrink: 0;
    }
    .bot-powered a { color: #00a4a6; text-decoration: none; }
  `;
  document.head.appendChild(style);

  // ── Inject HTML ─────────────────────────────────────────────
  const html = `
    <button id="pp-bot-toggle" title="Chat with Pascal Assistant">
      <i class="fa-solid fa-robot"></i>
      <span class="bot-notif"></span>
    </button>

    <div id="pp-bot-panel">
      <div class="bot-header">
        <div class="bot-avatar"><i class="fa-solid fa-robot"></i></div>
        <div class="bot-info">
          <div class="bot-name">Pascal Assistant</div>
          <div class="bot-status">Online — Ready to help</div>
        </div>
        <button class="bot-close" id="pp-bot-close" title="Close">✕</button>
      </div>

      <div id="pp-bot-messages"></div>

      <div class="bot-input-area">
        <input type="text" id="pp-bot-input" placeholder="Ask me anything…" autocomplete="off" />
        <button class="send-btn" id="pp-bot-send"><i class="fa-solid fa-paper-plane"></i></button>
      </div>
      <div class="bot-powered">Powered by <a href="https://copilotstudio.microsoft.com" target="_blank">Microsoft Copilot Studio</a></div>
    </div>
  `;
  const wrap = document.createElement('div');
  wrap.innerHTML = html;
  document.body.appendChild(wrap);

  // ── State ───────────────────────────────────────────────────
  const panel    = document.getElementById('pp-bot-panel');
  const toggle   = document.getElementById('pp-bot-toggle');
  const closeBtn = document.getElementById('pp-bot-close');
  const input    = document.getElementById('pp-bot-input');
  const sendBtn  = document.getElementById('pp-bot-send');
  const messages = document.getElementById('pp-bot-messages');
  let isOpen = false;
  let conversationId = null;
  let watermark = null;
  let pollTimer = null;

  // ── Auth token — prefer Power Platform token for bot ────────
  function getToken() {
    return localStorage.getItem('pp_bot_token') || localStorage.getItem('pp_token') || '';
  }

  // ── Helpers ─────────────────────────────────────────────────
  function getTime() {
    return new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  }

  function addMessage(text, sender = 'bot', opts = {}) {
    const isUser = sender === 'user';
    const initials = isUser ? 'You' : 'PA';
    const msg = document.createElement('div');
    msg.className = `bot-msg${isUser ? ' user' : ''}`;
    msg.innerHTML = `
      <div class="msg-avatar">${initials}</div>
      <div>
        <div class="bubble">${text}</div>
        <div class="msg-time">${getTime()}</div>
      </div>
    `;
    messages.appendChild(msg);

    if (opts.quickReplies && opts.quickReplies.length) {
      const qr = document.createElement('div');
      qr.className = 'quick-replies';
      opts.quickReplies.forEach(label => {
        const btn = document.createElement('button');
        btn.className = 'quick-reply-btn';
        btn.textContent = label;
        btn.onclick = () => { qr.remove(); handleUserMessage(label); };
        qr.appendChild(btn);
      });
      messages.appendChild(qr);
    }

    messages.scrollTop = messages.scrollHeight;
    return msg;
  }

  function showTyping() {
    if (document.getElementById('pp-typing')) return;
    const el = document.createElement('div');
    el.className = 'bot-msg';
    el.id = 'pp-typing';
    el.innerHTML = `
      <div class="msg-avatar">PA</div>
      <div class="bubble" style="padding:10px 14px;">
        <div class="typing-indicator"><span></span><span></span><span></span></div>
      </div>
    `;
    messages.appendChild(el);
    messages.scrollTop = messages.scrollHeight;
  }

  function hideTyping() {
    const el = document.getElementById('pp-typing');
    if (el) el.remove();
  }

  // ── Copilot Studio DirectLine via server proxy ───────────────
  async function startConversation() {
    const token = getToken();
    const headers = { 'Content-Type': 'application/json' };
    if (token) headers['Authorization'] = `Bearer ${token}`;
    const r = await fetch('/api/bot/conversations', { method: 'POST', headers });
    if (!r.ok) throw new Error(`Bot start failed: ${r.status}`);
    const data = await r.json();
    conversationId = data.conversationId;
    watermark = null;
    return data;
  }

  async function sendActivity(text) {
    const token = getToken();
    const headers = { 'Content-Type': 'application/json' };
    if (token) headers['Authorization'] = `Bearer ${token}`;
    await fetch(`/api/bot/conversations/${conversationId}/activities`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ type: 'message', text })
    });
  }

  async function pollActivities() {
    const token = getToken();
    const headers = {};
    if (token) headers['Authorization'] = `Bearer ${token}`;
    const wm = watermark != null ? `?watermark=${watermark}` : '';
    const r = await fetch(`/api/bot/conversations/${conversationId}/activities${wm}`, { headers });
    if (!r.ok) return;
    const data = await r.json();
    if (data.watermark) watermark = data.watermark;

    (data.activities || []).forEach(act => {
      if (act.type === 'message' && act.from?.role !== 'user') {
        hideTyping();
        const text = act.text || '';
        const cards = act.attachments || [];
        if (text) addMessage(text.replace(/\n/g, '<br>'), 'bot');
        cards.forEach(c => {
          if (c.content?.text) addMessage(c.content.text.replace(/\n/g, '<br>'), 'bot');
        });
        sendBtn.disabled = false;
        input.focus();
      }
    });
  }

  function startPolling() {
    stopPolling();
    pollTimer = setInterval(pollActivities, 1500);
    setTimeout(stopPolling, 30000);
  }

  function stopPolling() {
    if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
  }

  // ── Send message flow ────────────────────────────────────────
  async function handleUserMessage(text) {
    if (!text.trim()) return;
    addMessage(text, 'user');
    input.value = '';
    sendBtn.disabled = true;
    showTyping();

    try {
      if (!conversationId) await startConversation();
      await sendActivity(text);
      startPolling();
    } catch (err) {
      hideTyping();
      addMessage('⚠️ Could not reach Pascal Assistant. Please try again.', 'bot');
      console.error('[Bot]', err);
      sendBtn.disabled = false;
    }
  }

  // ── Welcome message ──────────────────────────────────────────
  async function showWelcome() {
    try {
      await startConversation();
      // Send a silent greeting to trigger bot welcome message
      await sendActivity('hello');
      showTyping();
      startPolling();
    } catch {
      addMessage(
        "Hi there! 👋 I'm <b>Pascal Assistant</b>. I can help with sales orders, customers, products and more.",
        'bot',
        { quickReplies: ['Sales Orders', 'Customers', 'Products', 'Help'] }
      );
    }
  }

  // ── Events ───────────────────────────────────────────────────
  toggle.addEventListener('click', () => {
    isOpen = !isOpen;
    panel.classList.toggle('open', isOpen);
    toggle.querySelector('.bot-notif').style.display = isOpen ? 'none' : '';
    if (isOpen && messages.childElementCount === 0) showWelcome();
    if (isOpen) setTimeout(() => input.focus(), 250);
  });

  closeBtn.addEventListener('click', () => {
    isOpen = false;
    panel.classList.remove('open');
    stopPolling();
  });

  sendBtn.addEventListener('click', () => handleUserMessage(input.value));
  input.addEventListener('keydown', e => { if (e.key === 'Enter') handleUserMessage(input.value); });
})();
