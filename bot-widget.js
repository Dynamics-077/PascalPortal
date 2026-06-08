/* ============================================================
   Pascal Press — Copilot Studio Bot Widget (iframe embed)
   ============================================================ */
(function () {

  const WEBCHAT_URL = 'https://copilotstudio.microsoft.com/environments/1db737e7-f1f2-e6ee-8744-c917393a84c5/bots/cr2d9_PascalPortal/webchat?__version__=2';

  // ── CSS ────────────────────────────────────────────────────
  const style = document.createElement('style');
  style.textContent = `
    #pp-bot-toggle {
      position: fixed; bottom: 24px; right: 24px; z-index: 9990;
      width: 52px; height: 52px; border-radius: 50%;
      background: linear-gradient(135deg, #00a4a6 0%, #007b7d 100%);
      border: none; cursor: pointer;
      box-shadow: 0 4px 18px rgba(0,164,166,.5);
      display: flex; align-items: center; justify-content: center;
      color: white; font-size: 1.35rem;
      transition: transform .2s, box-shadow .2s;
    }
    #pp-bot-toggle:hover {
      transform: scale(1.1);
      box-shadow: 0 6px 22px rgba(0,164,166,.6);
    }
    #pp-bot-toggle .pp-notif {
      position: absolute; top: 3px; right: 3px;
      width: 12px; height: 12px; border-radius: 50%;
      background: #f4851f; border: 2px solid white;
    }

    #pp-bot-panel {
      position: fixed; bottom: 88px; right: 24px; z-index: 9989;
      width: min(360px, calc(100vw - 32px)); height: 500px;
      border-radius: 16px; overflow: hidden;
      box-shadow: 0 16px 48px rgba(20,40,70,.22), 0 2px 8px rgba(0,0,0,.08);
      border: 1px solid rgba(0,164,166,.18);
      transform: scale(.93) translateY(14px); opacity: 0; pointer-events: none;
      transition: transform .22s cubic-bezier(.4,0,.2,1), opacity .22s ease;
    }
    #pp-bot-panel.open {
      transform: scale(1) translateY(0); opacity: 1; pointer-events: all;
    }

    #pp-bot-iframe {
      width: 100%; height: 100%; border: none; display: block;
    }

    @media (max-width: 480px) {
      #pp-bot-panel {
        width: calc(100vw - 16px);
        right: 8px; bottom: 80px;
        height: min(520px, calc(100vh - 100px));
      }
    }
  `;
  document.head.appendChild(style);

  // ── HTML ───────────────────────────────────────────────────
  const wrap = document.createElement('div');
  wrap.innerHTML = `
    <button id="pp-bot-toggle" title="Chat with Pascal Assistant">
      <i class="fa-solid fa-robot"></i>
      <span class="pp-notif"></span>
    </button>
    <div id="pp-bot-panel">
      <iframe
        id="pp-bot-iframe"
        src="${WEBCHAT_URL}"
        allow="microphone"
        title="Pascal Assistant">
      </iframe>
    </div>
  `;
  document.body.appendChild(wrap);

  // ── Toggle ─────────────────────────────────────────────────
  const panel  = document.getElementById('pp-bot-panel');
  const toggle = document.getElementById('pp-bot-toggle');
  const notif  = toggle.querySelector('.pp-notif');

  toggle.addEventListener('click', function () {
    var open = panel.classList.toggle('open');
    notif.style.display = open ? 'none' : '';
  });

})();
