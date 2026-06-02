require('dotenv').config();
const express = require('express');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
const axios = require('axios');
const sharepoint = require('./sharepoint');
const d365       = require('./d365');

const app = express();
const PORT = process.env.PORT || 5000;

// Enable CORS for localhost:3000 (React/HTML frontend)
app.use(cors({
    origin: true,
    credentials: true,
    optionsSuccessStatus: 200
}));

app.use(express.json());
app.use(express.static(__dirname));

/* ============================================================
   MICROSOFT ENTRA ID (AZURE AD) TOKEN VALIDATION
   Verify the incoming token against Microsoft's public keys
   ============================================================ */
const client = jwksClient({
    jwksUri: `https://login.microsoftonline.com/${process.env.ENTRA_TENANT_ID}/discovery/v2.0/keys`
});

function getKey(header, callback) {
    client.getSigningKey(header.kid, function (err, key) {
        if (err) return callback(err);
        const signingKey = key.publicKey || key.rsaPublicKey;
        callback(null, signingKey);
    });
}

// Middleware to Protect Routes
const validateToken = (req, res, next) => {
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
        // For development/mocking purposes when testing without real token:
        console.warn('⚠️ No token provided. Passing through for dev mock testing.');
        req.user = {
            oid:   'dev-user',
            name:  process.env.DEV_REP_NAME  || 'Sales Rep',
            email: process.env.DEV_REP_EMAIL || '',
        };
        return next();
        
        // Use this in production:
        // return res.status(401).json({ error: 'Missing Authorization header' });
    }

    const token = authHeader.split(' ')[1];
    
    jwt.verify(token, getKey, {
        audience: process.env.ENTRA_CLIENT_ID,
        issuer: `https://login.microsoftonline.com/${process.env.ENTRA_TENANT_ID}/v2.0`
    }, (err, decoded) => {
        if (err) {
            console.error('Token validation failed:', err.message);
            // In dev mode, we'll let it pass anyway. Un-comment the return in production.
            req.user = {
                oid:   'dev-user',
                name:  process.env.DEV_REP_NAME  || 'Sales Rep',
                email: process.env.DEV_REP_EMAIL || '',
            };
            return next();
            // return res.status(401).json({ error: 'Invalid Token' });
        }
        
        // Token is valid — attach decoded claims + extract role
        req.user      = decoded;
        req.user.role = (decoded.roles && decoded.roles[0]) || 'salesrep';
        console.log(`✅ Authenticated: ${req.user.preferred_username || req.user.name} | role: ${req.user.role}`);
        next();
    });
};

// Protect all /api/ routines
app.use('/api', validateToken);


// ── RBAC middleware ─────────────────────────────────────────────
// Usage: router.get('/admin-route', requireRole('admin'), handler)
const ROLE_LEVELS = { salesrep: 1, manager: 2, admin: 3 };

function requireRole(minRole) {
    return (req, res, next) => {
        const userRole  = req.user?.role || 'salesrep';
        const userLevel = ROLE_LEVELS[userRole]  || 1;
        const minLevel  = ROLE_LEVELS[minRole]   || 1;
        if (userLevel >= minLevel) return next();
        res.status(403).json({ error: `Access denied — requires role: ${minRole}` });
    };
}

function buildFallbackSalesRep(user = {}) {
    const name = user.name || user.given_name || user.preferred_username || 'Mock Rep';
    const firstName = name.split(' ')[0] || 'Rep';
    // Role: read from JWT 'roles' claim, or from env DEV_REP_ROLE, default salesrep
    const role = (user.roles && user.roles[0]) || user.role || process.env.DEV_REP_ROLE || 'salesrep';

    return {
        id: 'REP-003',
        name,
        firstName,
        email: user.preferred_username || user.email || 'sarah.thompson@pascalpress.com.au',
        phone: '0412 345 678',
        territory: 'NSW Metro Schools',
        d365WorkerId: 'D365-WORKER-0042',
        entraOid: user.oid || 'mock-user-id',
        role,
        targets: {
            monthly: 50000.00,
            ytd: 350000.00
        },
        performance: {
            monthly: 37420.00,
            ytd: 228650.00
        }
    };
}

function generateId(prefix) {
    const now = new Date();
    const date = now.toISOString().slice(0, 10).replace(/-/g, '');
    const rand = Math.floor(Math.random() * 9000) + 1000; // 1000-9999, matches PA rand(1000,9999)
    return `${prefix}-${date}-${rand}`;
}

function buildFallbackBootstrap(user = {}) {
    const salesRep = buildFallbackSalesRep(user);
    return {
        sharepointConfigured: false,
        salesRep,
        company: null,
        companies: [],
        customer: null,
        customers: [],
        salesOrders: [],
        dashboardSummary: {
            ordersToday: 0,
            ordersThisWeek: 0,
            ordersThisMonth: 0,
            revenueThisMonth: 0,
            pendingApprovals: 0,
            monthlyTarget: 50000.00,
            ytdRevenue: 0,
            ytdTarget: 350000.00
        }
    };
}

// 0. Bootstrap data for the portal
app.get('/api/bootstrap', async (req, res) => {
    try {
        const data = await sharepoint.getBootstrapData(req.user);
        res.json(data);
    } catch (error) {
        console.warn('[SharePoint] Bootstrap unavailable:', error.message);
        res.status(503).json(buildFallbackBootstrap(req.user));
    }
});

// 0b. Current authenticated user profile
app.get('/api/auth/me', async (req, res) => {
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        res.json(bootstrap.salesRep || buildFallbackSalesRep(req.user));
    } catch (error) {
        res.json(buildFallbackSalesRep(req.user));
    }
});

/* ============================================================
   API ENDPOINTS
   Below are the endpoints defined in the project plan.
   Currently, they return MOCK data. You will replace the mock 
   responses with axios.get/post calls to D365 FinOps APIs.
   ============================================================ */

// 1. Dashboard Settings
app.get('/api/dashboard/summary', async (req, res) => {
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        res.json(bootstrap.dashboardSummary || buildFallbackBootstrap(req.user).dashboardSummary);
    } catch (error) {
        res.json(buildFallbackBootstrap(req.user).dashboardSummary);
    }
});

// 2. Sales Orders POST — save to SharePoint
app.post('/api/salesorders', async (req, res) => {
    const payload = req.body;
    console.log('📦 New sales order received:', payload.salesId || 'DRAFT', '|', payload.customerName || payload.customerId);
    try {
        const created = await sharepoint.createSalesOrder(payload, req.user);
        res.status(201).json(created);
    } catch (error) {
        console.error('Order save error:', error.message);
        if (error.response) console.error('SharePoint response:', error.response.status, JSON.stringify(error.response.data));
        res.status(500).json({ error: error.message || 'Failed to save order' });
    }
});

// 3. Customers Search
app.get('/api/customers/search', async (req, res) => {
    const query = req.query.q || '';
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        const filtered = (bootstrap.customers || []).filter(customer => {
            const haystack = [customer.id, customer.name, customer.contactPerson, customer.email]
                .filter(Boolean)
                .join(' ')
                .toLowerCase();
            return haystack.includes(query.toLowerCase());
        });
        res.json(filtered);
    } catch (error) {
        res.json([]);
    }
});


// 5. Get all customers (full details)
app.get('/api/customers', async (req, res) => {
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        res.json(bootstrap.customers || []);
    } catch (error) {
        console.error('Error fetching customers:', error.message);
        res.status(500).json({ error: 'Failed to fetch customers' });
    }
});

// 6. Get single customer by ID (with addresses & order history)
app.get('/api/customers/:id', async (req, res) => {
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        const customer = (bootstrap.customers || []).find(c => c.id === req.params.id || c.accountNum === req.params.id);
        
        if (!customer) {
            return res.status(404).json({ error: 'Customer not found' });
        }

        // Enrich with order history
        const orders = (bootstrap.salesOrders || []).filter(o => o.customerId === customer.id);
        
        res.json({
            ...customer,
            orderHistory: orders
        });
    } catch (error) {
        console.error('Error fetching customer:', error.message);
        res.status(500).json({ error: 'Failed to fetch customer' });
    }
});

// 7. Get all sales orders
app.get('/api/salesorders', async (req, res) => {
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        res.json(bootstrap.salesOrders || []);
    } catch (error) {
        console.error('Error fetching sales orders:', error.message);
        res.status(500).json({ error: 'Failed to fetch sales orders' });
    }
});

// 8. Get sales order by ID
app.get('/api/salesorders/:id', async (req, res) => {
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        const order = (bootstrap.salesOrders || []).find(o => o.id === req.params.id || o.salesId === req.params.id);
        
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        res.json(order);
    } catch (error) {
        console.error('Error fetching order:', error.message);
        res.status(500).json({ error: 'Failed to fetch order' });
    }
});

// 9. Data sync status (check if SharePoint is configured) — admin only
app.get('/api/sync/status', requireRole('admin'), async (req, res) => {
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        res.json({
            configured: bootstrap.sharepointConfigured,
            lastSync: new Date().toISOString(),
            dataLoaded: {
                companies: (bootstrap.companies || []).length,
                customers: (bootstrap.customers || []).length,
                orders: (bootstrap.salesOrders || []).length
            }
        });
    } catch (error) {
        res.json({
            configured: false,
            error: error.message
        });
    }
});


// ============================================================
// Copilot Studio Bot — DirectLine Proxy
// ============================================================
const BOT_BASE = 'https://1db737e7f1f2e6ee8744c917393a84.c5.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/cr2d9_AISalesBot';
const BOT_API  = '2022-03-01-preview';

// Start a new bot conversation (No Authentication mode)
app.post('/api/bot/conversations', async (req, res) => {
    try {
        const r = await axios.post(
            `${BOT_BASE}/conversations?api-version=${BOT_API}`,
            {},
            { headers: { 'Content-Type': 'application/json' } }
        );
        res.json(r.data);
    } catch (e) {
        console.error('[Bot] Start error:', e.response?.status, JSON.stringify(e.response?.data) || e.message);
        res.status(e.response?.status || 500).json({ error: e.message, detail: e.response?.data });
    }
});

// Send message to bot
app.post('/api/bot/conversations/:id/activities', async (req, res) => {
    try {
        const r = await axios.post(
            `${BOT_BASE}/conversations/${req.params.id}/activities?api-version=${BOT_API}`,
            req.body,
            { headers: { 'Content-Type': 'application/json' } }
        );
        res.json(r.data);
    } catch (e) {
        console.error('[Bot] Send error:', e.response?.status, JSON.stringify(e.response?.data) || e.message);
        res.status(e.response?.status || 500).json({ error: e.message, detail: e.response?.data });
    }
});

// Poll bot responses
app.get('/api/bot/conversations/:id/activities', async (req, res) => {
    try {
        const wm = req.query.watermark ? `&watermark=${req.query.watermark}` : '';
        const r = await axios.get(
            `${BOT_BASE}/conversations/${req.params.id}/activities?api-version=${BOT_API}${wm}`
        );
        res.json(r.data);
    } catch (e) {
        console.error('[Bot] Poll error:', e.response?.status, JSON.stringify(e.response?.data) || e.message);
        res.status(e.response?.status || 500).json({ error: e.message, detail: e.response?.data });
    }
});

// Start Server
app.listen(PORT, () => {
    console.log(`\n🚀 Backend API Middleware started!`);
    console.log(`   Listening directly on http://localhost:${PORT}`);
    console.log(`   Ready to route requests towards Dynamics 365.\n`);

    // Pre-warm D365 customer cache so the first browser load is fast
    d365.getCustomers({ top: 500 })
        .then(data => console.log(`[D365] Customer cache pre-warmed — ${(data.value || []).length} customers ready`))
        .catch(err  => console.warn('[D365] Customer pre-warm failed (will retry on first request):', err.message));

    // Pre-warm D365 product cache
    d365.getProducts({ top: 1000 })
        .then(data => console.log(`[D365] Product cache pre-warmed — ${(data.value || []).length} products ready`))
        .catch(err  => console.warn('[D365] Product pre-warm failed (will retry on first request):', err.message));
});

// -------------------------------
// Admin CRUD for Company
// -------------------------------
app.post('/api/companies', requireRole('admin'), async (req, res) => {
    try {
        const fields = Object.assign({}, req.body);
        if (!fields.CompanyId) fields.CompanyId = generateId('COMP');
        if (!fields.Title) fields.Title = fields.CompanyId;
        const created = await sharepoint.createListItem('Company', fields);
        res.status(201).json(created);
    } catch (error) {
        console.error('Create company error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

app.put('/api/companies/:id', requireRole('admin'), async (req, res) => {
    try {
        const itemId = req.params.id;
        const fields = req.body;
        const updated = await sharepoint.updateListItem('Company', itemId, fields);
        res.json(updated);
    } catch (error) {
        console.error('Update company error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

app.delete('/api/companies/:id', requireRole('admin'), async (req, res) => {
    try {
        const itemId = req.params.id;
        await sharepoint.deleteListItem('Company', itemId);
        res.json({ success: true });
    } catch (error) {
        console.error('Delete company error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

// -------------------------------
// Admin CRUD for Customers (AllCustomers)
// -------------------------------
app.post('/api/customers', requireRole('admin'), async (req, res) => {
    try {
        const fields = Object.assign({}, req.body);
        if (!fields.AccountNum) fields.AccountNum = generateId('CUST');
        if (!fields.Title) fields.Title = fields.AccountNum;
        const created = await sharepoint.createListItem('AllCustomers', fields);
        res.status(201).json(created);
    } catch (error) {
        console.error('Create customer error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

app.put('/api/customers/:id', requireRole('admin'), async (req, res) => {
    try {
        const id = req.params.id;
        let item = null;
        // if numeric id, treat as item id; otherwise try to find by AccountNum
        if (/^\d+$/.test(id)) {
            item = { id };
        } else {
            item = await sharepoint.findItemByField('AllCustomers', 'AccountNum', id);
            if (!item) {
                return res.status(404).json({ error: 'Customer not found' });
            }
        }
        const updated = await sharepoint.updateListItem('AllCustomers', item.id, req.body);
        res.json(updated);
    } catch (error) {
        console.error('Update customer error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

app.delete('/api/customers/:id', requireRole('admin'), async (req, res) => {
    try {
        const id = req.params.id;
        let item = null;
        if (/^\d+$/.test(id)) {
            item = { id };
        } else {
            item = await sharepoint.findItemByField('AllCustomers', 'AccountNum', id);
            if (!item) return res.status(404).json({ error: 'Customer not found' });
        }
        await sharepoint.deleteListItem('AllCustomers', item.id);
        res.json({ success: true });
    } catch (error) {
        console.error('Delete customer error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

// -------------------------------
// Admin: Sales Orders update/delete
// -------------------------------
app.post('/api/salesorders/header', async (req, res) => {
    try {
        const salesId = generateId('PAS'); // always server-generated, never from client
        const fields = Object.assign({}, req.body, { SalesId: salesId });
        delete fields.CustName; // SP Graph API rejects CustName — field not writable via items endpoint
        if (!fields.Title) fields.Title = salesId;
        console.log('[SP] Creating SalesOrderHeader with fields:', JSON.stringify(fields));
        const created = await sharepoint.createListItem('SalesOrderHeader', fields);
        const spItemId = created?.id || created?.fields?.id || null;
        console.log('[SP] SalesOrderHeader created, salesId:', salesId, 'spItemId:', spItemId);
        // Return salesId + spItemId so client can do PUT/DELETE by SP item ID (bypasses SalesId column lookup)
        res.status(201).json({ ...created, salesId, spItemId });
    } catch (error) {
        console.error('[SP] Create order header error:', error.message);
        if (error.response) {
            console.error('[SP] Graph status:', error.response.status);
            console.error('[SP] Graph body:', JSON.stringify(error.response.data));
        }
        const status = error.response?.status === 400 ? 400 : 500;
        res.status(status).json({ error: error.message, detail: error.response?.data });
    }
});

app.put('/api/salesorders/:salesId', async (req, res) => {
    try {
        const salesId = req.params.salesId;
        const item = await sharepoint.findItemByField('SalesOrderHeader', 'SalesId', salesId);
        if (!item) return res.status(404).json({ error: 'Order not found' });
        console.log('[SP] Updating SalesOrderHeader', salesId, 'fields:', JSON.stringify(req.body));
        const updated = await sharepoint.updateListItem('SalesOrderHeader', item.id, req.body);
        res.json(updated);
    } catch (error) {
        console.error('[SP] Update order error:', error.message);
        if (error.response) {
            console.error('[SP] Graph status:', error.response.status);
            console.error('[SP] Graph body:', JSON.stringify(error.response.data));
        }
        const status = error.response?.status === 400 ? 400 : 500;
        res.status(status).json({ error: error.message, detail: error.response?.data });
    }
});

app.delete('/api/salesorders/:salesId', requireRole('manager'), async (req, res) => {
    try {
        const salesId = req.params.salesId;
        const item = await sharepoint.findItemByField('SalesOrderHeader', 'SalesId', salesId);
        if (!item) return res.status(404).json({ error: 'Order not found' });
        await sharepoint.deleteListItem('SalesOrderHeader', item.id);
        res.json({ success: true });
    } catch (error) {
        console.error('Delete order error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

// PUT/DELETE by SharePoint item ID directly — bypasses SalesId column lookup
app.put('/api/salesorders/item/:itemId', async (req, res) => {
    try {
        const updated = await sharepoint.updateListItem('SalesOrderHeader', req.params.itemId, req.body);
        res.json(updated);
    } catch (error) {
        console.error('[SP] Update order error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

app.delete('/api/salesorders/item/:itemId', requireRole('manager'), async (req, res) => {
    try {
        await sharepoint.deleteListItem('SalesOrderHeader', req.params.itemId);
        res.json({ success: true });
    } catch (error) {
        console.error('[SP] Delete order error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

// -------------------------------
// Admin: Sales Order Lines CRUD
// -------------------------------
app.post('/api/salesorders/:salesId/lines', async (req, res) => {
    try {
        const salesId     = req.params.salesId;
        const lineNumber  = parseInt(req.body.lineNumber) || 1;
        const lineSalesId = `${salesId}-L${lineNumber}`;
        const fields = Object.assign({}, req.body, { SalesId: lineSalesId });
        console.log('[SP] Creating SalesOrderLine', lineSalesId, 'fields:', JSON.stringify(fields));
        const created = await sharepoint.createListItem('SalesOrderLines', fields);
        console.log('[SP] Line created, item id:', created?.id);
        res.status(201).json(created);
    } catch (error) {
        console.error('[SP] Create line error:', error.message);
        if (error.response) {
            console.error('[SP] Graph status:', error.response.status);
            console.error('[SP] Graph body:', JSON.stringify(error.response.data));
        }
        const status = error.response?.status === 400 ? 400 : 500;
        res.status(status).json({ error: error.message, detail: error.response?.data });
    }
});

app.put('/api/salesorders/lines/:itemId', async (req, res) => {
    try {
        const itemId = req.params.itemId;
        const updated = await sharepoint.updateListItem('SalesOrderLines', itemId, req.body);
        res.json(updated);
    } catch (error) {
        console.error('Update order line error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

app.delete('/api/salesorders/lines/:itemId', requireRole('manager'), async (req, res) => {
    try {
        const itemId = req.params.itemId;
        await sharepoint.deleteListItem('SalesOrderLines', itemId);
        res.json({ success: true });
    } catch (error) {
        console.error('Delete order line error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

// Debug: show SharePoint list column names (internal vs display)
app.get('/api/debug/list-schema/:listName', async (req, res) => {
    try {
        const columns = await sharepoint.getListColumns(req.params.listName);
        res.json({ listName: req.params.listName, columns });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 5b. Get all companies
app.get('/api/companies', async (req, res) => {
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        res.json(bootstrap.companies || []);
    } catch (error) {
        console.error('Error fetching companies:', error.message);
        res.status(500).json({ error: 'Failed to fetch companies' });
    }
});

// 5c. Get raw order lines
app.get('/api/orderlines', async (req, res) => {
    try {
        const rows = await sharepoint.getRawListItemsByName('SalesOrderLines', ['SHAREPOINT_SALES_ORDER_LINES_LIST_ID', 'SHAREPOINT_ORDER_LINES_LIST_ID']);
        res.json(rows.map(row => ({
            id: row.id,
            ...row.fields
        })));
    } catch (error) {
        console.error('Error fetching order lines:', error.message);
        res.status(500).json({ error: 'Failed to fetch order lines' });
    }
});

// ============================================================
// D365 F&O — Customers & Products
// ============================================================

// GET /api/d365/customers?search=&top=&skip=
app.get('/api/d365/customers', async (req, res) => {
    try {
        const { search = '', top = 500, skip = 0 } = req.query;
        const data = await d365.getCustomers({ search, top: +top, skip: +skip });

        // Return only the fields the UI needs — keeps payload small & avoids @odata.etag issues
        const slim = (data.value || []).map(c => ({
            CustomerAccount:     c.CustomerAccount     || '',
            OrganizationName:    c.OrganizationName    || '',
            NameAlias:           c.NameAlias           || '',
            CustomerGroupId:     c.CustomerGroupId     || '',
            CurrencyCode:        c.SalesCurrencyCode   || c.CurrencyCode || '',
            AddressCity:         c.AddressCity         || '',
            AddressState:        c.AddressState        || '',
            AddressCountryRegionId: c.AddressCountryRegionId || '',
            AddressStreet:       c.AddressStreet       || '',
            AddressZipCode:      c.AddressZipCode      || '',
            PrimaryContactEmail: c.PrimaryContactEmail || '',
            PrimaryContactPhone: c.PrimaryContactPhone || '',
            PaymentTermsName:    c.PaymentTerms        || '',
            DeliveryTerms:       c.DeliveryTerms       || '',
            SalesTaxGroup:       c.SalesTaxGroup       || '',
            LanguageId:          c.LanguageId          || '',
        }));

        res.json({ value: slim, count: slim.length });
    } catch (error) {
        console.error('[D365] Customers error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

// GET /api/d365/customers/:account
app.get('/api/d365/customers/:account', async (req, res) => {
    try {
        const customer = await d365.getCustomer(req.params.account);
        res.json(customer);
    } catch (error) {
        console.error('[D365] Customer detail error:', error.message);
        const status = error.message.includes('not found') ? 404 : 500;
        res.status(status).json({ error: error.message });
    }
});

// GET /api/d365/products?search=&top=&skip=&itemNumber=
app.get('/api/d365/products', async (req, res) => {
    try {
        const { search = '', top = 20, skip = 0, itemNumber = '' } = req.query;
        const data = await d365.getProducts({ search, top: +top, skip: +skip, itemNumber });
        const slim = (data.value || []).map(p => ({
            ItemNumber:          p.ItemNumber          || '',
            ProductNumber:       p.ProductNumber       || '',
            SearchName:          p.SearchName          || p.ProductSearchName || '',
            ProductGroupId:      p.ProductGroupId      || '',
            ProductType:         p.ProductType         || '',
            SalesPrice:          p.SalesPrice          || 0,
            UnitCost:            p.UnitCost            || 0,
            PurchasePrice:       p.PurchasePrice       || 0,
            SalesUnitSymbol:     p.SalesUnitSymbol     || 'ea',
            InventoryUnitSymbol: p.InventoryUnitSymbol || 'ea',
            dataAreaId:          p.dataAreaId          || '',
        }));
        res.json({ value: slim, count: slim.length });
    } catch (error) {
        console.error('[D365] Products error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

// GET /api/d365/salesorders?salesOrderNumber=&customerAccount=&top=&skip=
app.get('/api/d365/salesorders', async (req, res) => {
    try {
        const { salesOrderNumber = '', customerAccount = '', top = 50, skip = 0 } = req.query;
        const data = await d365.getSalesOrders({ salesOrderNumber, customerAccount, top: +top, skip: +skip });
        res.json(data);
    } catch (error) {
        console.error('[D365] Sales orders error:', error.message);
        if (error.response) console.error('[D365] Detail:', JSON.stringify(error.response.data));
        res.status(error.response?.status || 500).json({ error: error.message });
    }
});

// GET /api/d365/salesorders/:salesOrderNumber/lines
app.get('/api/d365/salesorders/:salesOrderNumber/lines', async (req, res) => {
    try {
        const data = await d365.getSalesOrderLines(req.params.salesOrderNumber);
        res.json(data);
    } catch (error) {
        console.error('[D365] Sales order lines error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

// GET /api/d365/products/:itemNumber
app.get('/api/d365/products/:itemNumber', async (req, res) => {
    try {
        const product = await d365.getProduct(req.params.itemNumber);
        res.json({
            ItemNumber:          product.ItemNumber          || '',
            ProductNumber:       product.ProductNumber       || '',
            SearchName:          product.SearchName          || product.ProductSearchName || '',
            ProductGroupId:      product.ProductGroupId      || '',
            ProductType:         product.ProductType         || '',
            SalesPrice:          product.SalesPrice          || 0,
            UnitCost:            product.UnitCost            || 0,
            PurchasePrice:       product.PurchasePrice       || 0,
            SalesUnitSymbol:     product.SalesUnitSymbol     || 'ea',
            InventoryUnitSymbol: product.InventoryUnitSymbol || 'ea',
            dataAreaId:          product.dataAreaId          || '',
        });
    } catch (error) {
        console.error('[D365] Product detail error:', error.message);
        const status = error.message.includes('not found') ? 404 : 500;
        res.status(status).json({ error: error.message });
    }
});