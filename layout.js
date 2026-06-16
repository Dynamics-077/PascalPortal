/* =============================================================
   layout.js — Shared Sidebar + RBAC for Pascal Press Portal
   Include this before </body> on every portal page.
   Each page needs: <div id="pp-sidebar"></div> as sidebar placeholder.
   ============================================================= */
(function () {
  'use strict';

  // ── Role hierarchy ──────────────────────────────────────────
  // salesrep (1) < manager (2) < admin (3)
  const ROLE_LEVELS = { salesrep: 1, manager: 2, admin: 3 };

  // ── Navigation definition ───────────────────────────────────
  // minRole: only show item if user role >= this level
  const NAV = [
    { section: 'Main' },
    { id: 'dashboard',  href: 'dashboard.html',                icon: 'fa-chart-bar',      label: 'Dashboard' },
    { id: 'sales-order',href: 'sales-order.html',              icon: 'fa-plus-circle',    label: 'New Order' },
    { id: 'all-orders', href: 'all-orders.html',               icon: 'fa-clipboard-list', label: 'All Orders' },
    { id: 'all-quotes', href: 'all-quotes.html',               icon: 'fa-file-invoice-dollar', label: 'Quotes' },
    { id: 'quote',      href: 'quote.html',                    icon: 'fa-plus-square',    label: 'New Quote' },
    { id: 'data-admin', href: 'data-admin.html',               icon: 'fa-database',       label: 'Data Admin', minRole: 'admin' },

    { section: 'Dynamics 365' },
    { id: 'customers',  href: 'customers.html',                icon: 'fa-building',       label: 'Customers' },
    { id: 'products',   href: 'customers.html#products',       icon: 'fa-box',            label: 'Products', navId: 'navProducts' },

    { section: 'Support' },
    { id: 'contact',    href: 'mailto:orders@pascalpress.com.au', icon: 'fa-envelope',    label: 'Contact' },
  ];

  // Role display config
  const ROLE_CONFIG = {
    admin:    { label: 'Admin',        color: '#ef4444' },
    manager:  { label: 'Sales Manager',color: '#f97316' },
    salesrep: { label: 'Sales Rep',    color: '#00a4a6' },
  };

  // ── Detect active page from URL ─────────────────────────────
  function getActivePage() {
    const file = location.pathname.split('/').pop() || 'dashboard.html';
    const base = file.replace('.html', '') || 'dashboard';
    if (base === 'customers' && location.hash === '#products') return 'products';
    return base;
  }

  // ── Role access check ───────────────────────────────────────
  function canSee(item, role) {
    if (!item.minRole) return true;
    return (ROLE_LEVELS[role] || 1) >= (ROLE_LEVELS[item.minRole] || 1);
  }

  // ── Build initials from name ────────────────────────────────
  function initials(name) {
    return (name || 'SR').split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
  }

  // ── Render sidebar HTML ─────────────────────────────────────
  function buildSidebar(user, role) {
    const active   = getActivePage();
    const rc       = ROLE_CONFIG[role] || ROLE_CONFIG.salesrep;
    const userInit = initials(user.name);

    const navHtml = NAV.map(item => {
      if (item.section) {
        return `<div class="nav-section-label" style="margin-top:8px;">${item.section}</div>`;
      }
      if (!canSee(item, role)) return '';

      const isActive = active === item.id;
      return `<a href="${item.href}"
                 class="nav-item${isActive ? ' active' : ''}"
                 ${item.navId ? `id="${item.navId}"` : ''}>
                <i class="fas ${item.icon} nav-icon"></i> ${item.label}
              </a>`;
    }).join('');

    return `
      <aside class="sidebar">
        <div class="sidebar-logo">
          <div class="logo-mark">
            <div class="logo-icon">PP</div>
            <div class="logo-text">
              <div class="brand">Pascal Press</div>
              <div class="sub">Sales Rep Portal</div>
            </div>
          </div>
        </div>

        <nav class="sidebar-nav">${navHtml}</nav>

        <div class="sidebar-footer">
          <button class="logout-btn" onclick="logout()">
            <i class="fas fa-sign-out-alt"></i> Sign Out
          </button>
        </div>
      </aside>`;
  }

  // ── Mobile hamburger + backdrop ─────────────────────────────
  function setupMobileNav() {
    if (document.querySelector('.sidebar-backdrop')) return;

    const bd = document.createElement('div');
    bd.className = 'sidebar-backdrop';
    bd.addEventListener('click', closeSidebar);
    document.body.appendChild(bd);

    const topbar = document.querySelector('.topbar');
    if (topbar && !topbar.querySelector('.hamburger-btn')) {
      const btn = document.createElement('button');
      btn.className = 'hamburger-btn';
      btn.setAttribute('aria-label', 'Open menu');
      btn.innerHTML = '<i class="fas fa-bars"></i>';
      btn.addEventListener('click', toggleSidebar);
      topbar.insertBefore(btn, topbar.firstChild);
    }
  }

  function toggleSidebar() {
    const sb     = document.querySelector('.sidebar');
    const bd     = document.querySelector('.sidebar-backdrop');
    const isOpen = sb && sb.classList.toggle('open');
    if (bd) bd.classList.toggle('open', !!isOpen);
    document.body.style.overflow = isOpen ? 'hidden' : '';
    const icon = document.querySelector('.hamburger-btn i');
    if (icon) icon.className = isOpen ? 'fas fa-times' : 'fas fa-bars';
  }

  function closeSidebar() {
    const sb = document.querySelector('.sidebar');
    const bd = document.querySelector('.sidebar-backdrop');
    if (sb) sb.classList.remove('open');
    if (bd) bd.classList.remove('open');
    document.body.style.overflow = '';
    const icon = document.querySelector('.hamburger-btn i');
    if (icon) icon.className = 'fas fa-bars';
  }

  // Close sidebar when a nav item is tapped on mobile
  document.addEventListener('click', function (e) {
    const item = e.target.closest('.nav-item');
    if (item && window.innerWidth <= 900) closeSidebar();
  });

  // ── Inject sidebar into placeholder ────────────────────────
  function inject(user, role) {
    const el = document.getElementById('pp-sidebar');
    if (el) {
      el.outerHTML = buildSidebar(user, role);
    } else {
      // Fallback: replace existing sidebar if placeholder missing
      const existing = document.querySelector('.sidebar');
      if (existing) existing.outerHTML = buildSidebar(user, role);
    }
  }

  // ── Page-level RBAC guard ───────────────────────────────────
  // Call this after fetching the user role.
  // Pages that require a minimum role redirect if user doesn't qualify.
  const PAGE_ROLES = {
    'data-admin': 'admin',
  };

  function enforcePageAccess(role) {
    const page    = getActivePage();
    const minRole = PAGE_ROLES[page];
    if (!minRole) return; // page open to all
    if ((ROLE_LEVELS[role] || 1) < (ROLE_LEVELS[minRole] || 1)) {
      window.location.replace('dashboard.html');
    }
  }

  // ── Bootstrap: fetch user + inject ─────────────────────────
  async function boot() {
    // 1. Show cached NAME immediately to avoid flash — but NEVER
    //    cache role, always fetch from server so .env changes
    //    take effect on next page load without clearing storage.
    let cachedName = '';
    try {
      const raw = localStorage.getItem('pp_salesrep');
      if (raw) cachedName = JSON.parse(raw).name || '';
    } catch (_) {}

    inject({ name: cachedName }, 'salesrep');
    setupMobileNav();

    // 2. Fetch authoritative role + user from server
    try {
      const token = localStorage.getItem('pp_token');
      const res   = await fetch('/api/bootstrap', {
        headers: { Authorization: token ? `Bearer ${token}` : '' }
      });
      if (!res.ok) return;
      const data = await res.json();
      const rep  = data.salesRep || data.rep || {};
      const role = rep.role || 'salesrep';

      // Only cache name — never cache role (always read fresh from server)
      localStorage.setItem('pp_salesrep', JSON.stringify({ name: rep.name }));

      // Re-render sidebar with correct role
      inject(rep, role);

      // Enforce page access after real role confirmed
      enforcePageAccess(role);

    } catch (_) {}
  }

  // ── Run when DOM ready ──────────────────────────────────────
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }

  // Expose for manual calls if needed
  window.Layout = { refresh: boot };

})();
