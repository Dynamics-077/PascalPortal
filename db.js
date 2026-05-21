/* ============================================================
   Pascal Press Sales Rep Portal — Data Loader (db.js)

   Loads data from:
     1. /api/bootstrap  (SharePoint via server — primary)
     2. MOCK data in app.js (fallback)

   Four tables: Company | AllCustomers | SalesOrderHeader | SalesOrderLines
   ============================================================ */

'use strict';

const DB = (function () {

  let _ready    = false;
  let salesRep  = null;
  let companies = [];
  let customers = [];
  let orders    = [];

  /* ----------------------------------------------------------
     load() — fetch bootstrap data from API, fall back to MOCK
     Call once at page startup: await DB.load();
     ---------------------------------------------------------- */
  async function load() {
    if (_ready) return;
    // Clear any stale mock data saved by old testing code
    localStorage.removeItem('pp_mock_orders');

    try {
      const res = await fetch('/api/bootstrap');
      if (res.ok) {
        const snapshot = await res.json();
        if (snapshot && snapshot.sharepointConfigured === false) {
          throw new Error('SharePoint not configured — using MOCK data');
        }
        _loadFromSnapshot(snapshot);
        _ready = true;
        _syncPublicMock();
        console.info('[DB] Loaded from API bootstrap');
        return;
      }
      throw new Error(`Cannot fetch /api/bootstrap (${res.status})`);
    } catch (err) {
      console.warn('[DB] API bootstrap failed — using MOCK data.', err.message);
      _loadFromMock();
      _ready = true;
      _syncPublicMock();
    }
  }

  function _loadFromSnapshot(snapshot = {}) {
    salesRep  = snapshot.salesRep  || null;
    companies = Array.isArray(snapshot.companies)   ? snapshot.companies   : [];
    customers = Array.isArray(snapshot.customers)   ? snapshot.customers   : [];
    orders    = Array.isArray(snapshot.salesOrders) ? snapshot.salesOrders : [];
  }

  function _loadFromMock() {
    if (typeof PP === 'undefined' || !PP.MOCK) return;
    salesRep  = PP.MOCK.salesRep || null;
    companies = Array.isArray(PP.MOCK.companies)   ? PP.MOCK.companies   : [];
    customers = Array.isArray(PP.MOCK.customers)   ? PP.MOCK.customers   : [];
    orders    = Array.isArray(PP.MOCK.salesOrders) ? PP.MOCK.salesOrders : [];
    console.info('[DB] Using built-in MOCK data');
  }

  function _syncPublicMock() {
    if (typeof window === 'undefined' || !window.PP || !window.PP.MOCK) return;
    Object.assign(window.PP.MOCK, {
      salesRep:    salesRep    || window.PP.MOCK.salesRep,
      companies,
      customers,
      salesOrders: orders,
    });
  }

  function searchCustomers(query) {
    if (!query || query.length < 1) return customers.slice(0, 10);
    const q = query.toLowerCase();
    return customers
      .filter(c => [c.name, c.id, c.accountNum, c.contactPerson]
        .filter(Boolean)
        .join(' ')
        .toLowerCase()
        .includes(q))
      .slice(0, 10);
  }

  return {
    load,
    searchCustomers,
    get ready()     { return _ready;     },
    get salesRep()  { return salesRep;   },
    get companies() { return companies;  },
    get customers() { return customers;  },
    get orders()    { return orders;     },
  };

})();
