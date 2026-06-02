/* ============================================================
   Pascal Press — Copilot Studio Bot Widget
   Uses official Copilot Studio webchat iframe embed.
   ============================================================ */
(function () {

  const WEBCHAT_URL = 'https://copilotstudio.microsoft.com/environments/1db737e7-f1f2-e6ee-8744-c917393a84c5/bots/cr2d9_PascalPortal/webchat?__version__=2';

  // ── CSS ────────────────────────────────────────────────────
  const style = document.createElement('style');
  style.textContent = `
    #toast-container { bottom: 100px !important; }

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

    #pp-bot-iframe {
      flex: 1; border: none; width: 100%; display: block;
    }
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
      <iframe
        id="pp-bot-iframe"
        src="${WEBCHAT_URL}"
        allow="microphone"
        title="Pascal Assistant Chat">
      </iframe>
    </div>
  `;
  document.body.appendChild(wrap);

  // ── Toggle logic ───────────────────────────────────────────
  const panel  = document.getElementById('pp-bot-panel');
  const toggle = document.getElementById('pp-bot-toggle');
  const close  = document.getElementById('pp-bot-close');
  let isOpen   = false;

  toggle.addEventListener('click', () => {
    isOpen = !isOpen;
    panel.classList.toggle('open', isOpen);
    toggle.querySelector('.bot-notif').style.display = isOpen ? 'none' : '';
  });

  close.addEventListener('click', () => {
    isOpen = false;
    panel.classList.remove('open');
  });

})();
