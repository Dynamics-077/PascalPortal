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
app.use(express.static(__dirname, { etag: false, lastModified: false, setHeaders: (res) => { res.setHeader('Cache-Control', 'no-store'); } }));

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
        applyAdminWhitelist(req.user);
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
            // Token signature/audience invalid — but decode payload to get real user identity
            // so the email filter still works correctly with the actual logged-in user.
            try {
                const unverified = jwt.decode(token);
                const uemail = unverified?.preferred_username || unverified?.email || unverified?.upn || '';
                if (uemail) {
                    req.user = {
                        oid:                unverified.oid  || 'unverified',
                        name:               unverified.name || process.env.DEV_REP_NAME || 'Sales Rep',
                        preferred_username: uemail,
                        email:              uemail,
                        role:               (unverified.roles && unverified.roles[0]) || process.env.DEV_REP_ROLE || 'salesrep',
                    };
                    applyAdminWhitelist(req.user);
                    console.warn(`⚠️ Token unverified — identity from payload: ${uemail} | role: ${req.user.role}`);
                    return next();
                }
            } catch (_) { /* ignore decode error */ }
            // No identity in token — fall back to dev mock
            req.user = {
                oid:   'dev-user',
                name:  process.env.DEV_REP_NAME  || 'Sales Rep',
                email: process.env.DEV_REP_EMAIL || '',
            };
            return next();
        }

        // Token is valid — attach decoded claims + extract role
        req.user      = decoded;
        req.user.role = (decoded.roles && decoded.roles[0]) || 'salesrep';
        applyAdminWhitelist(req.user);
        console.log(`✅ Authenticated: ${req.user.preferred_username || req.user.name} | role: ${req.user.role}`);
        next();
    });
};

// Protect all /api/ routines
app.use('/api', validateToken);


// ── RBAC middleware ─────────────────────────────────────────────
// Usage: router.get('/admin-route', requireRole('admin'), handler)
const ROLE_LEVELS = { salesrep: 1, manager: 2, admin: 3 };

// ADMIN_EMAILS acts as a strict allowlist:
// - if the email IS in the list → promote to admin
// - if the email is NOT in the list → demote any admin claim to salesrep
function applyAdminWhitelist(user) {
    const list = (process.env.ADMIN_EMAILS || '')
        .split(',').map(e => e.trim().toLowerCase()).filter(Boolean);
    if (!list.length) return user; // no whitelist configured — leave role unchanged
    const email = (user.preferred_username || user.email || '').toLowerCase();
    if (list.includes(email)) {
        user.role = 'admin';
    } else if (user.role === 'admin') {
        user.role = 'salesrep'; // strip admin from anyone not in the whitelist
    }
    return user;
}

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
        const userEmail = req.user?.preferred_username || req.user?.email || '';
        console.log(`[Order] Creating header — user: ${userEmail || '(no email — token missing?)'} | auth header present: ${!!req.headers.authorization}`);
        // Whitelist: only columns that exist in SalesOrderHeader SharePoint list
        const HEADER_COLS = ['Title','CustAccount','Currency','CustGroup','InvoiceAccount','Status','DeliveryTerms','PaymentTerms'];
        const fields = {};
        HEADER_COLS.forEach(k => { if (req.body[k] !== undefined) fields[k] = req.body[k]; });
        fields.SalesId = salesId;
        fields.Email   = userEmail;
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
        const userEmail   = req.user?.preferred_username || req.user?.email || '';
        // Whitelist: only columns that exist in SalesOrderLines SharePoint list
        const LINE_COLS = ['Title','lineNumber','CustAccount','Currency','CustGroup','SalesUnit','SalesPrice','SalesQty','OrderLineStatus','ItemCode'];
        const fields = {};
        LINE_COLS.forEach(k => { if (req.body[k] !== undefined && req.body[k] !== '') fields[k] = req.body[k]; });
        fields.SalesId = salesId;
        fields.Email   = userEmail;
        console.log('[SP] Creating SalesOrderLine', salesId, 'fields:', JSON.stringify(fields));
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

// ============================================================
// Quotes (SalesQuoteHeader + SalesQuoteLines)
// ============================================================

app.get('/api/quotes', async (req, res) => {
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        res.json(bootstrap.quotes || []);
    } catch (error) {
        console.error('Error fetching quotes:', error.message);
        res.status(500).json({ error: 'Failed to fetch quotes' });
    }
});

app.get('/api/quotes/:quoteId', async (req, res) => {
    try {
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        const quote = (bootstrap.quotes || []).find(
            q => q.id === req.params.quoteId || q.quoteId === req.params.quoteId
        );
        if (!quote) return res.status(404).json({ error: 'Quote not found' });
        res.json(quote);
    } catch (error) {
        console.error('Error fetching quote:', error.message);
        res.status(500).json({ error: 'Failed to fetch quote' });
    }
});

app.post('/api/quotes/header', async (req, res) => {
    try {
        const quoteId   = generateId('QUO');
        const userEmail = req.user?.preferred_username || req.user?.email || '';
        const b = req.body;
        // Map frontend/code field names → actual SharePoint column names
        const fields = {
            Title:           b.Title         || b.CustomerName || quoteId,
            QuoteId:         quoteId,
            CustomerAccount: b.CustAccount   || b.CustomerAccount || '',
            CustomerName:    b.CustomerName  || '',
            Currency:        b.Currency      || 'AUD',
            CustomerGroup:   b.CustGroup     || b.CustomerGroup  || '',
            Status:          b.Status        || 'Draft',
            DeliveryTerms:   b.DeliveryTerms || '',
            PaymentTerms:    b.PaymentTerms  || '',
            ValidUntil:      b.ValidUntil    || '',
            Note:            b.Notes         || b.Note || '',
            QuoteRevision:   b.QuoteRevision || 1,
            ParentQuoteId:   b.ParentQuoteId || '',
            ConvertedOrderId:b.ConvertedOrderId || '',
            Warehouse:       b.Warehouse        || '',
            DeliveryAddress: b.DeliveryAddress  || '',
            Email:           userEmail,
        };
        // Remove empty strings to avoid SP validation issues
        Object.keys(fields).forEach(k => { if (fields[k] === '') delete fields[k]; });
        console.log('[SP] Creating SalesQuoteHeader:', JSON.stringify(fields));
        const created  = await sharepoint.createListItem('SalesQuoteHeader', fields);
        const spItemId = created?.id || null;
        res.status(201).json({ ...created, quoteId, spItemId });
    } catch (error) {
        console.error('[SP] Create quote header error:', error.message);
        if (error.response) console.error('[SP] Graph body:', JSON.stringify(error.response.data));
        res.status(error.response?.status === 400 ? 400 : 500).json({ error: error.message, detail: error.response?.data });
    }
});

app.put('/api/quotes/item/:itemId', async (req, res) => {
    try {
        const b = req.body;
        // Map frontend field names → actual SharePoint column names
        const fields = {};
        if (b.Title         !== undefined) fields.Title           = b.Title;
        if (b.Status        !== undefined) fields.Status          = b.Status;
        if (b.CustAccount   !== undefined) fields.CustomerAccount = b.CustAccount;
        if (b.CustomerAccount !== undefined) fields.CustomerAccount = b.CustomerAccount;
        if (b.CustomerName  !== undefined) fields.CustomerName    = b.CustomerName;
        if (b.CustGroup     !== undefined) fields.CustomerGroup   = b.CustGroup;
        if (b.CustomerGroup !== undefined) fields.CustomerGroup   = b.CustomerGroup;
        if (b.Currency      !== undefined) fields.Currency        = b.Currency;
        if (b.DeliveryTerms !== undefined) fields.DeliveryTerms   = b.DeliveryTerms;
        if (b.PaymentTerms  !== undefined) fields.PaymentTerms    = b.PaymentTerms;
        if (b.ValidUntil    !== undefined) fields.ValidUntil      = b.ValidUntil;
        if (b.Notes         !== undefined) fields.Note            = b.Notes;
        if (b.Note          !== undefined) fields.Note            = b.Note;
        if (b.QuoteRevision !== undefined) fields.QuoteRevision   = b.QuoteRevision;
        if (b.ParentQuoteId    !== undefined) fields.ParentQuoteId    = b.ParentQuoteId;
        if (b.ConvertedOrderId !== undefined) fields.ConvertedOrderId = b.ConvertedOrderId;
        if (b.Warehouse        !== undefined) fields.Warehouse        = b.Warehouse;
        if (b.DeliveryAddress  !== undefined) fields.DeliveryAddress  = b.DeliveryAddress;
        const updated = await sharepoint.updateListItem('SalesQuoteHeader', req.params.itemId, fields);
        res.json(updated);
    } catch (error) {
        console.error('[SP] Update quote error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/quotes/:quoteId/lines', async (req, res) => {
    try {
        const quoteId   = req.params.quoteId;
        const userEmail = req.user?.preferred_username || req.user?.email || '';
        const lineIdx = parseInt(req.body.lineNumber || req.body.linenumber) || 1;
        const b = req.body;
        // Map frontend/code field names → actual SharePoint column names
        const fields = {
            Title:           b.Title        || `${quoteId}-L${lineIdx}`,
            QuotationId:     `${quoteId}-L${lineIdx}`,
            CustomerAccount: b.CustAccount  || b.CustomerAccount || '',
            CustomerGroup:   b.CustGroup    || b.CustomerGroup   || '',
            Currency:        b.Currency     || 'AUD',
            linenumber:      String(lineIdx),
            Itemnumber:      b.ItemCode     || b.Itemnumber      || '',
            productname:     b.ItemName     || b.productname     || '',
            ItemCategory:    b.ItemCategory || '',
            SalesQuantity:   String(parseFloat(b.SalesQty || b.SalesQuantity || 0) || 0),
            // SP internal name confusion: "SalesUnit" column has internal name "SalesPrice"
            //                            "SalesPrice" column has internal name "SalesPrice0"
            SalesPrice:      b.SalesUnit || 'ea',
            SalesPrice0:     parseFloat(b.SalesPrice  || b.SalesPrice0 || 0) || 0,
            Discount:        parseFloat(b.Discount    || 0) || 0,
            DeliveryType:    b.DeliveryType || 'Stock',
            Email:           userEmail,
        };
        // Remove empty strings to avoid SP validation issues
        Object.keys(fields).forEach(k => { if (fields[k] === '') delete fields[k]; });
        console.log('[SP] Creating SalesQuoteLine', quoteId, 'fields:', JSON.stringify(fields));
        const created = await sharepoint.createListItem('SalesQuoteLines', fields);
        res.status(201).json(created);
    } catch (error) {
        console.error('[SP] Create quote line error:', error.message);
        if (error.response) console.error('[SP] Graph body:', JSON.stringify(error.response.data));
        res.status(error.response?.status === 400 ? 400 : 500).json({ error: error.message, detail: error.response?.data });
    }
});

// POST /api/quotes/:quoteId/revise — lock old Active quote as Revised, create new Draft
app.post('/api/quotes/:quoteId/revise', async (req, res) => {
    try {
        const quoteId   = req.params.quoteId;
        const userEmail = req.user?.preferred_username || req.user?.email || '';
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        const original  = (bootstrap.quotes || []).find(q => q.quoteId === quoteId || q.id === quoteId);
        if (!original)                    return res.status(404).json({ error: 'Quote not found' });
        if (original.status !== 'Active') return res.status(400).json({ error: 'Only Active quotes can be revised' });

        // 1. Lock old quote as Revised (immutable audit record)
        await sharepoint.updateListItem('SalesQuoteHeader', original.spItemId, { Status: 'Revised' });

        // 2. Create new header (incremented revision)
        const newQuoteId  = generateId('QUO');
        const newRevision = (original.quoteRevision || 1) + 1;
        const parentId    = original.parentQuoteId || quoteId;
        const newHeader   = {
            Title:           original.customerName  || original.custAccount || newQuoteId,
            QuoteId:         newQuoteId,
            CustomerAccount: original.custAccount   || '',
            CustomerName:    original.customerName  || '',
            Currency:        original.currency      || 'AUD',
            CustomerGroup:   original.customerGroup || '',
            Status:          'Draft',
            DeliveryTerms:   original.deliveryTerms || '',
            PaymentTerms:    original.paymentTerms  || '',
            ValidUntil:      original.validUntil    || '',
            Note:            original.notes         || '',
            QuoteRevision:   newRevision,
            ParentQuoteId:   parentId,
            Email:           userEmail,
        };
        Object.keys(newHeader).forEach(k => { if (newHeader[k] === '') delete newHeader[k]; });
        const created = await sharepoint.createListItem('SalesQuoteHeader', newHeader);
        console.log(`[SP] Revision ${newRevision} created: ${newQuoteId} (parent: ${parentId})`);

        // 3. Copy all lines to the new revision
        for (let i = 0; i < (original.lines || []).length; i++) {
            const l  = original.lines[i];
            const lf = {
                Title:           l.name         || `${newQuoteId}-L${i + 1}`,
                QuotationId:     `${newQuoteId}-L${i + 1}`,
                CustomerAccount: original.custAccount   || '',
                CustomerGroup:   original.customerGroup || '',
                Currency:        original.currency      || 'AUD',
                linenumber:      String(i + 1),
                Itemnumber:      l.itemNo                || '',
                productname:     l.name                  || '',
                ItemCategory:    l.category              || '',
                SalesQuantity:   String(parseFloat(l.qty)    || 0),
                SalesPrice:      l.unit                  || 'ea',
                SalesPrice0:     parseFloat(l.price)     || 0,
                Discount:        parseFloat(l.discount)  || 0,
                DeliveryType:    l.deliveryType          || 'Stock',
                Email:           userEmail,
            };
            Object.keys(lf).forEach(k => { if (lf[k] === '') delete lf[k]; });
            await sharepoint.createListItem('SalesQuoteLines', lf);
        }

        res.status(201).json({
            quoteId:       newQuoteId,
            spItemId:      created?.id,
            revision:      newRevision,
            parentQuoteId: parentId,
            linesCopied:   (original.lines || []).length,
        });
    } catch (error) {
        console.error('[SP] Revise quote error:', error.message);
        if (error.response) console.error('[SP] Graph body:', JSON.stringify(error.response.data));
        res.status(500).json({ error: error.message });
    }
});

// POST /api/quotes/:quoteId/convert — convert Accepted quote into a new Sales Order
app.post('/api/quotes/:quoteId/convert', async (req, res) => {
    try {
        const quoteId   = req.params.quoteId;
        const bootstrap = await sharepoint.getBootstrapData(req.user);
        const quote     = (bootstrap.quotes || []).find(q => q.quoteId === quoteId || q.id === quoteId);
        if (!quote)                      return res.status(404).json({ error: 'Quote not found' });
        if (quote.status !== 'Accepted') return res.status(400).json({ error: 'Only Accepted quotes can be converted to orders' });

        const salesId = generateId('PAS');
        console.log(`[SP] Converting quote ${quoteId} → order ${salesId}`);

        // Build payload using same structure as createSalesOrder
        const orderPayload = {
            salesId,
            customerName:  quote.customerName  || quote.custAccount,
            customerId:    quote.custAccount,
            custAccount:   quote.custAccount,
            currency:      quote.currency      || 'AUD',
            customerGroup: quote.customerGroup || '',
            custGroup:     quote.customerGroup || '',
            invoiceAccount:quote.custAccount,
            status:        'In Progress',
            deliveryTerms: quote.deliveryTerms || '',
            paymentTerms:  quote.paymentTerms  || '',
            lines: (quote.lines || []).map((l, i) => ({
                lineNo:          i + 1,
                itemNo:          l.itemNo       || '',
                name:            l.name         || '',
                category:        l.category     || '',
                qty:             l.qty          || 1,
                unit:            l.unit         || 'ea',
                unitPrice:       l.price        || 0,
                price:           l.price        || 0,
                discount:        l.discount     || 0,
                deliveryType:    l.deliveryType || 'Stock',
                orderLineStatus: 'In Progress',
            })),
        };

        const result = await sharepoint.createSalesOrder(orderPayload, req.user);
        if (!result.success) throw new Error(result.message || 'Order creation failed');

        // Update quote: Converted + store back-reference
        await sharepoint.updateListItem('SalesQuoteHeader', quote.spItemId, {
            Status:          'Converted',
            ConvertedOrderId: salesId,
        });

        res.json({ salesId, quoteId, linesCreated: (quote.lines || []).length });
    } catch (error) {
        console.error('[SP] Convert quote error:', error.message);
        if (error.response) console.error('[SP] Graph body:', JSON.stringify(error.response.data));
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

// 5c. Get raw order lines — filtered by current user email
app.get('/api/orderlines', async (req, res) => {
    try {
        const rows = await sharepoint.getRawListItemsByName('SalesOrderLines', ['SHAREPOINT_SALES_ORDER_LINES_LIST_ID', 'SHAREPOINT_ORDER_LINES_LIST_ID']);
        const userEmail = (req.user?.preferred_username || req.user?.email || '').toLowerCase();
        const isAdmin = req.user?.role === 'admin';
        const lines = rows.map(row => ({ id: row.id, ...row.fields }));
        if (!isAdmin && userEmail) {
            return res.json(lines.filter(l => (l.Email || '').toLowerCase() === userEmail));
        }
        res.json(lines);
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