/* ============================================================
   Pascal Press — Copilot Studio Bot Widget
   WebSocket proxy through backend server
   ============================================================ */
(function () {

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
      transition: transform .2s, box-shadow .2s; color: white; font-size: 1.4rem;
    }
    #pp-bot-toggle:hover { transform: scale(1.08); }
    #pp-bot-toggle .bot-notif {
      position: absolute; top: 2px; right: 2px; width: 14px; height: 14px;
      border-radius: 50%; background: #f4851f; border: 2px solid white;
    }

    #pp-bot-panel {
      position: fixed; bottom: 96px; right: 28px; z-index: 9989;
      width: 340px; height: 480px; background: white; border-radius: 14px;
      box-shadow: 0 12px 40px rgba(30,58,95,.22);
      display: flex; flex-direction: column; overflow: hidden;
      transform: scale(.92) translateY(16px); opacity: 0; pointer-events: none;
      transition: transform .22s ease, opacity .22s ease; border: 1px solid #d8dde6;
    }
    #pp-bot-panel.open { transform: scale(1) translateY(0); opacity: 1; pointer-events: all; }

    .bot-header {
      background: linear-gradient(135deg, #1e3a5f 0%, #2a4d7a 100%);
      padding: 14px 16px; display: flex; align-items: center; gap: 10px; flex-shrink: 0;
    }
    .bot-header .bot-avatar {
      width: 36px; height: 36px; border-radius: 50%; background: #00a4a6;
      display: flex; align-items: center; justify-content: center;
      font-size: .95rem; color: white; flex-shrink: 0;
    }
    .bot-header .bot-info { flex: 1; }
    .bot-header .bot-name  { font-size: .88rem; font-weight: 700; color: white; }
    .bot-header .bot-status {
      font-size: .72rem; color: rgba(255,255,255,.7);
      display: flex; align-items: center; gap: 4px;
    }
    .bot-header .bot-status::before {
      content: ''; width: 6px; height: 6px; border-radius: 50%;
      background: #27ae60; display: inline-block;
    }
    .bot-header .bot-close {
      background: none; border: none; color: rgba(255,255,255,.6);
      cursor: pointer; font-size: 1rem; padding: 4px; border-radius: 6px; line-height: 1;
    }
    .bot-header .bot-close:hover { color: white; background: rgba(255,255,255,.1); }

    #pp-bot-messages {
      flex: 1; overflow-y: auto; padding: 12px;
      display: flex; flex-direction: column; gap: 8px; background: #f9fafb;
    }
    #pp-bot-messages::-webkit-scrollbar { width: 4px; }
    #pp-bot-messages::-webkit-scrollbar-thumb { background: #d1d5db; border-radius: 4px; }

    .pp-msg { max-width: 82%; padding: 9px 13px; border-radius: 12px; font-size: 13px; line-height: 1.5; word-wrap: break-word; }
    .pp-msg.bot  { background: white; border: 1px solid #e5e7eb; color: #1f2937; align-self: flex-start; border-bottom-left-radius: 4px; }
    .pp-msg.user { background: #00a4a6; color: white; align-self: flex-end; border-bottom-right-radius: 4px; }
    .pp-msg.sys  { background: none; color: #9ca3af; font-size: 11px; align-self: center; border: none; padding: 2px 0; }

    .pp-typing { align-self: flex-start; padding: 10px 14px; background: white; border: 1px solid #e5e7eb; border-radius: 12px; border-bottom-left-radius: 4px; display: flex; gap: 4px; align-items: center; }
    .pp-typing span { width: 6px; height: 6px; background: #9ca3af; border-radius: 50%; animation: pp-b 1.2s infinite; }
    .pp-typing span:nth-child(2) { animation-delay: .2s; }
    .pp-typing span:nth-child(3) { animation-delay: .4s; }
    @keyframes pp-b { 0%,80%,100%{transform:scale(.8);opacity:.5} 40%{transform:scale(1.1);opacity:1} }

    #pp-bot-input-row {
      padding: 10px 12px; border-top: 1px solid #e5e7eb;
      display: flex; gap: 8px; align-items: flex-end;
      background: white; flex-shrink: 0;
    }
    #pp-bot-input {
      flex: 1; border: 1px solid #d1d5db; border-radius: 8px;
      padding: 8px 11px; font-size: 13px; resize: none; outline: none;
      line-height: 1.4; max-height: 80px; font-family: inherit; background: #f9fafb;
    }
    #pp-bot-input:focus { border-color: #00a4a6; background: white; }
    #pp-bot-send {
      width: 34px; height: 34px; border-radius: 8px; background: #00a4a6;
      border: none; color: white; cursor: pointer; display: flex;
      align-items: center; justify-content: center; flex-shrink: 0; transition: background .15s;
    }
    #pp-bot-send:hover { background: #008385; }
    #pp-bot-send:disabled { background: #d1d5db; cursor: default; }

    #pp-bot-overlay {
      position: absolute; inset: 52px 0 0 0;
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      gap: 12px; padding: 24px; text-align: center;
      background: #f9fafb; z-index: 5; font-size: 13px; color: #6b7280;
    }
    #pp-bot-overlay.hidden { display: none; }
    #pp-bot-overlay .pp-spin {
      border: 3px solid #e5e7eb; border-top: 3px solid #00a4a6;
      border-radius: 50%; width: 28px; height: 28px; animation: pp-s .8s linear infinite;
    }
    #pp-bot-overlay .pp-err { color: #dc2626; font-size: 12px; line-height: 1.6; }
    @keyframes pp-s { to { transform: rotate(360deg); } }
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
          <div class="bot-status">Online</div>
        </div>
        <button class="bot-close" id="pp-bot-close">✕</button>
      </div>
      <div id="pp-bot-messages">
        <div id="pp-bot-overlay">
          <div class="pp-spin"></div>
          <span id="pp-bot-overlay-txt">Connecting...</span>
        </div>
      </div>
      <div id="pp-bot-input-row">
        <textarea id="pp-bot-input" rows="1" placeholder="Type a message..." disabled></textarea>
        <button id="pp-bot-send" disabled title="Send">
          <i class="fa-solid fa-paper-plane" style="font-size:.75rem"></i>
        </button>
      </div>
    </div>
  `;
  document.body.appendChild(wrap);

  // ── Refs ───────────────────────────────────────────────────
  const panel   = document.getElementById('pp-bot-panel');
  const toggle  = document.getElementById('pp-bot-toggle');
  const closeB  = document.getElementById('pp-bot-close');
  const msgs    = document.getElementById('pp-bot-messages');
  const overlay = document.getElementById('pp-bot-overlay');
  const oTxt    = document.getElementById('pp-bot-overlay-txt');
  const input   = document.getElementById('pp-bot-input');
  const sendBtn = document.getElementById('pp-bot-send');

  let isOpen = false, ws = null, ready = false, _msalInst = null;

  // ── MSAL config — same app registration as the portal ──────
  var BOT_MSAL = {
    clientId: '6bc856af-36d8-424d-a089-1860d402627b',
    tenantId: '132fee41-6bf5-4f91-be3e-5c3b2a2fb1b8',
    scopes:   ['https://api.powerplatform.com/CopilotStudio.Copilots.Invoke'],
  };

  // ── Toggle ─────────────────────────────────────────────────
  toggle.addEventListener('click', function() {
    isOpen = !isOpen;
    panel.classList.toggle('open', isOpen);
    toggle.querySelector('.bot-notif').style.display = isOpen ? 'none' : '';
    if (isOpen && !ready) connect();
    if (isOpen) setTimeout(function() { input.focus(); }, 300);
  });
  closeB.addEventListener('click', function() { isOpen = false; panel.classList.remove('open'); });

  // ── Load MSAL dynamically if not already on the page ───────
  function loadMsal(cb) {
    if (typeof msal !== 'undefined') { cb(); return; }
    var s = document.createElement('script');
    s.src = 'https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js';
    s.onload  = cb;
    s.onerror = function() { showOverlay('Auth library failed to load. Check your connection.', true); };
    document.head.appendChild(s);
  }

  // ── Acquire a real Power Platform token via MSAL ────────────
  function acquirePPToken(onSuccess) {
    showOverlay('Authenticating...');
    loadMsal(function() {
      try {
        if (!_msalInst) {
          _msalInst = new msal.PublicClientApplication({
            auth: {
              clientId:    BOT_MSAL.clientId,
              authority:   'https://login.microsoftonline.com/' + BOT_MSAL.tenantId,
              redirectUri: window.location.origin + window.location.pathname,
            },
            cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: true },
          });
        }

        _msalInst.handleRedirectPromise().then(function(resp) {
          // If redirect brought back a PP token, use it immediately
          if (resp && resp.accessToken) { onSuccess(resp.accessToken); return; }

          var accounts = _msalInst.getAllAccounts();
          if (accounts.length === 0) {
            // No cached session — trigger interactive login then return to this page
            _msalInst.loginRedirect({ scopes: BOT_MSAL.scopes, prompt: 'select_account' });
            return;
          }

          _msalInst.acquireTokenSilent({ scopes: BOT_MSAL.scopes, account: accounts[0] })
            .then(function(result) { onSuccess(result.accessToken); })
            .catch(function(err) {
              console.warn('[Bot] Silent token failed, trying redirect:', err.message);
              _msalInst.loginRedirect({ scopes: BOT_MSAL.scopes });
            });

        }).catch(function(err) {
          console.error('[Bot] Redirect promise error:', err);
          showOverlay('Auth error: ' + err.message, true);
        });

      } catch (err) {
        console.error('[Bot] MSAL init error:', err);
        showOverlay('Auth init error: ' + err.message, true);
      }
    });
  }

  // ── Connect via WebSocket proxy ────────────────────────────
  function connect() {
    acquirePPToken(function(ppToken) {
      showOverlay('Connecting to assistant...');
      var proto = location.protocol === 'https:' ? 'wss' : 'ws';
      var wsUrl = proto + '://' + location.host + '/api/bot/stream?token=' + encodeURIComponent(ppToken);
      console.log('[Bot] Connecting WebSocket proxy...');

      ws = new WebSocket(wsUrl);
      ws.onopen = function() { console.log('[Bot] WS proxy connected'); };

      ws.onmessage = function(e) {
        try {
          var msg = JSON.parse(e.data);
          if (msg.type === 'ready') {
            hideOverlay();
            ready = true;
            input.disabled = false;
            sendBtn.disabled = false;
            input.focus();
          } else if (msg.type === 'activities') {
            (msg.activities || []).forEach(renderAct);
          } else if (msg.type === 'error') {
            addMsg('⚠️ ' + msg.message, 'sys');
          }
        } catch (_) {}
      };

      ws.onerror = function(e) {
        console.error('[Bot] WS error', e);
        msgs.querySelectorAll('.pp-typing').forEach(function(el) { el.remove(); });
        sendBtn.disabled = false;
        showOverlay('Could not connect. Please try again.', true);
      };

      ws.onclose = function(e) {
        console.log('[Bot] WS closed:', e.code, e.reason);
        msgs.querySelectorAll('.pp-typing').forEach(function(el) { el.remove(); });
        sendBtn.disabled = false;
        if (ready) addMsg('Connection closed.', 'sys');
      };
    });
  }

  // ── Send ───────────────────────────────────────────────────
  function send() {
    var text = input.value.trim();
    if (!text || !ws || ws.readyState !== WebSocket.OPEN) return;
    input.value = ''; input.style.height = 'auto';
    addMsg(text, 'user');
    var typing = addTyping();
    sendBtn.disabled = true;
    ws.send(JSON.stringify({ type: 'message', text: text }));
    var origMsg = ws.onmessage;
    ws.onmessage = function(e) {
      typing.remove();
      sendBtn.disabled = false;
      ws.onmessage = origMsg;
      origMsg(e);
    };
  }

  input.addEventListener('keydown', function(e) { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); } });
  input.addEventListener('input',   function() { input.style.height = 'auto'; input.style.height = Math.min(input.scrollHeight, 80) + 'px'; });
  sendBtn.addEventListener('click', send);

  // ── Render ─────────────────────────────────────────────────
  function renderAct(act) {
    if (act.type !== 'message') return;
    if (act.from && act.from.role === 'user') return;
    var text = act.text || act.speak || '';
    if (text) addMsg(text, 'bot');
  }

  function addMsg(text, cls) {
    var d = document.createElement('div');
    d.className = 'pp-msg ' + cls;
    d.textContent = text;
    msgs.appendChild(d);
    msgs.scrollTop = msgs.scrollHeight;
    return d;
  }

  function addTyping() {
    var d = document.createElement('div');
    d.className = 'pp-typing';
    d.innerHTML = '<span></span><span></span><span></span>';
    msgs.appendChild(d);
    msgs.scrollTop = msgs.scrollHeight;
    return d;
  }

  function showOverlay(msg, err) {
    overlay.classList.remove('hidden');
    overlay.innerHTML = err
      ? '<div class="pp-err">⚠️ ' + msg + '</div>'
      : '<div class="pp-spin"></div><span>' + msg + '</span>';
  }
  function hideOverlay() { overlay.classList.add('hidden'); }

})();
