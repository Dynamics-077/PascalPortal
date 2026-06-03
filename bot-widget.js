/* ============================================================
   Pascal Press — Copilot Studio Bot Widget
   M365 Agents SDK Direct Connect (authenticated, no iframe)
   ============================================================ */
(function () {

  const DIRECT_CONNECT_URL = 'https://1db737e7f1f2e6ee8744c917393a84.c5.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/cr2d9_AISalesBot/conversations?api-version=2022-03-01-preview';
  const WEBCHAT_CDN        = 'https://cdn.botframework.com/botframework-webchat/latest/webchat.js';

  // ── CSS ────────────────────────────────────────────────────
  const style = document.createElement('style');
  style.textContent = `
    #toast-container { bottom: 100px !important; }

    /* ── Toggle button ── */
    #pp-bot-toggle {
      position: fixed; bottom: 28px; right: 28px; z-index: 9990;
      width: 56px; height: 56px; border-radius: 50%;
      background: linear-gradient(135deg, #00a4a6, #008385);
      border: none; cursor: pointer;
      box-shadow: 0 4px 16px rgba(0,164,166,.45);
      display: flex; align-items: center; justify-content: center;
      transition: transform .2s, box-shadow .2s;
      color: white; font-size: 1.4rem;
    }
    #pp-bot-toggle:hover { transform: scale(1.08); box-shadow: 0 6px 20px rgba(0,164,166,.55); }
    #pp-bot-toggle .bot-notif {
      position: absolute; top: 2px; right: 2px;
      width: 14px; height: 14px; border-radius: 50%;
      background: #f4851f; border: 2px solid white;
    }

    /* ── Chat panel ── */
    #pp-bot-panel {
      position: fixed; bottom: 96px; right: 28px; z-index: 9989;
      width: 340px; height: 480px;
      background: white; border-radius: 14px;
      box-shadow: 0 12px 40px rgba(30,58,95,.22);
      display: flex; flex-direction: column; overflow: hidden;
      transform: scale(.92) translateY(16px);
      opacity: 0; pointer-events: none;
      transition: transform .22s ease, opacity .22s ease;
      border: 1px solid #d8dde6;
    }
    #pp-bot-panel.open { transform: scale(1) translateY(0); opacity: 1; pointer-events: all; }

    /* ── Custom header ── */
    #pp-bot-panel .bot-header {
      background: linear-gradient(135deg, #1e3a5f 0%, #2a4d7a 100%);
      padding: 14px 16px; display: flex; align-items: center; gap: 10px; flex-shrink: 0;
    }
    .bot-header .bot-avatar {
      width: 36px; height: 36px; border-radius: 50%; background: #00a4a6;
      display: flex; align-items: center; justify-content: center;
      font-size: .95rem; color: white; flex-shrink: 0;
      box-shadow: 0 2px 8px rgba(0,164,166,.4);
    }
    .bot-header .bot-info { flex: 1; }
    .bot-header .bot-info .bot-name  { font-size: .88rem; font-weight: 700; color: white; }
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
      border-radius: 6px; transition: color .15s; line-height: 1;
    }
    .bot-header .bot-close:hover { color: white; background: rgba(255,255,255,.1); }

    /* ── Chat area ── */
    #pp-bot-chat {
      flex: 1; overflow: hidden; position: relative;
      display: flex; flex-direction: column;
    }

    /* ── Loading / Error state inside chat area ── */
    #pp-bot-status {
      position: absolute; inset: 0;
      display: flex; flex-direction: column;
      align-items: center; justify-content: center;
      gap: 12px; padding: 24px; text-align: center;
      background: white; z-index: 1;
      font-size: 13px; color: #555;
    }
    #pp-bot-status.hidden { display: none; }
    #pp-bot-status .bot-spinner {
      border: 3px solid #e0e0e0;
      border-top: 3px solid #00a4a6;
      border-radius: 50%;
      width: 32px; height: 32px;
      animation: pp-spin 0.8s linear infinite;
    }
    @keyframes pp-spin { to { transform: rotate(360deg); } }
    #pp-bot-status .bot-err {
      color: #dc2626; font-size: 12px; line-height: 1.5;
    }

    /* ── WebChat overrides — fill panel ── */
    #pp-bot-chat > div { height: 100% !important; }
  `;
  document.head.appendChild(style);

  // ── HTML ───────────────────────────────────────────────────
  const wrap = document.createElement('div');
  wrap.innerHTML = `
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
      <div id="pp-bot-chat">
        <div id="pp-bot-status">
          <div class="bot-spinner"></div>
          <span id="pp-bot-status-text">Connecting to assistant...</span>
        </div>
      </div>
    </div>
  `;
  document.body.appendChild(wrap);

  // ── Toggle logic ───────────────────────────────────────────
  const panel  = document.getElementById('pp-bot-panel');
  const toggle = document.getElementById('pp-bot-toggle');
  const close  = document.getElementById('pp-bot-close');
  let isOpen   = false;
  let chatReady = false;

  toggle.addEventListener('click', () => {
    isOpen = !isOpen;
    panel.classList.toggle('open', isOpen);
    toggle.querySelector('.bot-notif').style.display = isOpen ? 'none' : '';
    if (isOpen && !chatReady) initWebChat();
  });

  close.addEventListener('click', () => {
    isOpen = false;
    panel.classList.remove('open');
  });

  // ── Init WebChat (runs once on first open) ─────────────────
  async function initWebChat() {
    const statusEl  = document.getElementById('pp-bot-status');
    const statusTxt = document.getElementById('pp-bot-status-text');

    function showStatus(msg, isError) {
      statusEl.classList.remove('hidden');
      if (isError) {
        statusEl.innerHTML = `<div class="bot-err">⚠️ ${msg}</div>`;
      } else {
        statusTxt.textContent = msg;
      }
    }

    try {
      // 1. Get bot token stored during login
      const token = localStorage.getItem('pp_bot_token');
      if (!token) {
        showStatus('Session token missing.<br>Please sign out and sign in again.', true);
        return;
      }

      showStatus('Loading chat engine...');

      // 2. Load WebChat SDK if not already loaded
      await loadScript(WEBCHAT_CDN);

      showStatus('Connecting to assistant...');

      // 3. Create authenticated Direct Line connection
      // createDirectLineAppServiceExtension uses WebSocket streaming (ASE protocol)
      // If WebSocket fails, fall back to standard DirectLine polling
      let directLine;
      try {
        directLine = await window.WebChat.createDirectLineAppServiceExtension({
          domain: DIRECT_CONNECT_URL,
          token:  token,
        });
      } catch (aseErr) {
        console.warn('[Bot] ASE failed, trying DirectLine:', aseErr.message);
        directLine = window.WebChat.createDirectLine({
          domain: DIRECT_CONNECT_URL.replace('/conversations', ''),
          token:  token,
        });
      }

      // 4. Hide status overlay
      statusEl.classList.add('hidden');
      chatReady = true;

      // 5. Render WebChat inside the chat div
      const chatDiv = document.getElementById('pp-bot-chat');
      window.WebChat.renderWebChat(
        {
          directLine,
          locale: 'en-US',
          styleOptions: {
            accent:                     '#00a4a6',
            backgroundColor:            '#f9fafb',
            bubbleBackground:           '#ffffff',
            bubbleBorderColor:          '#e0e0e0',
            bubbleBorderRadius:         10,
            bubbleFromUserBackground:   '#00a4a6',
            bubbleFromUserTextColor:    '#ffffff',
            bubbleFromUserBorderRadius: 10,
            sendBoxBackground:          '#ffffff',
            sendBoxBorderTop:           '1px solid #e5e7eb',
            fontSizeSmall:              '12px',
            hideUploadButton:           true,
            hideScrollToEndButton:      false,
          }
        },
        chatDiv
      );

    } catch (err) {
      console.error('[Bot] Init failed:', err);
      showStatus('Could not connect: ' + err.message, true);
    }
  }

  // ── Helper: load external script once ─────────────────────
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) return resolve();
      const s = document.createElement('script');
      s.src = src; s.onload = resolve; s.onerror = reject;
      document.head.appendChild(s);
    });
  }

})();
