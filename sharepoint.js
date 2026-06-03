require('dotenv').config();
const axios = require('axios');

const graphBaseUrl = 'https://graph.microsoft.com/v1.0';

let cachedToken = null;
let cachedTokenExpiresAt = 0;
const listIdCache = new Map();

function isConfigured() {
    return Boolean(
        process.env.SHAREPOINT_TENANT_ID &&
        process.env.SHAREPOINT_CLIENT_ID &&
        process.env.SHAREPOINT_CLIENT_SECRET &&
        process.env.SHAREPOINT_SITE_ID
    );
}

function requireEnv(name) {
    const value = process.env[name];
    if (!value) {
        throw new Error(`Missing environment variable: ${name}`);
    }
    return value;
}

function cleanText(value) {
    if (value === null || value === undefined) {
        return '';
    }
    return String(value).trim();
}

function asNumber(value) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
}

function asBoolean(value) {
    return String(value).toLowerCase() === 'true';
}

function generateId(prefix) {
    const now = new Date();
    const date = now.toISOString().slice(0, 10).replace(/-/g, '');
    const rand = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
    return `${prefix}-${date}-${rand}`;
}

function buildSalesRep(user = {}) {
    const name      = cleanText(user.name || user.given_name || user.preferred_username || process.env.DEV_REP_NAME || 'Sales Rep');
    const firstName = name.split(' ')[0] || 'Rep';
    const role      = (user.roles && user.roles[0]) || user.role || process.env.DEV_REP_ROLE || 'salesrep';

    return {
        id:           cleanText(user.oid || 'dev-user'),
        name,
        firstName,
        email:        cleanText(user.preferred_username || user.email || process.env.DEV_REP_EMAIL || ''),
        phone:        '',
        territory:    cleanText(process.env.DEV_REP_TERRITORY || ''),
        d365WorkerId: cleanText(user.d365WorkerId || ''),
        entraOid:     cleanText(user.oid || 'dev-user'),
        role,
    };
}

async function getAccessToken() {
    if (!isConfigured()) {
        throw new Error('SharePoint configuration is incomplete');
    }

    if (cachedToken && Date.now() < cachedTokenExpiresAt - 60000) {
        return cachedToken;
    }

    const tenantId = requireEnv('SHAREPOINT_TENANT_ID');
    const clientId = requireEnv('SHAREPOINT_CLIENT_ID');
    const clientSecret = requireEnv('SHAREPOINT_CLIENT_SECRET');

    const body = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: 'client_credentials',
        scope: 'https://graph.microsoft.com/.default'
    });

    const response = await axios.post(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        body,
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    cachedToken = response.data.access_token;
    cachedTokenExpiresAt = Date.now() + ((response.data.expires_in || 3600) * 1000);
    return cachedToken;
}

async function graphRequest(method, url, data) {
    const token = await getAccessToken();
    const response = await axios({
        method,
        url: `${graphBaseUrl}${url}`,
        data,
        headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
        }
    });
    return response.data;
}

async function readListItems(listId, select = '*') {
    if (!listId) {
        return [];
    }

    const data = await graphRequest(
        'get',
        `/sites/${requireEnv('SHAREPOINT_SITE_ID')}/lists/${listId}/items?expand=fields($select=${encodeURIComponent(select)})&$top=999`
    );

    return (data.value || []).map(item => ({
        id: item.id,
        createdDateTime: item.createdDateTime,
        lastModifiedDateTime: item.lastModifiedDateTime,
        webUrl: item.webUrl,
        fields: item.fields || {}
    }));
}

async function getRawListItemsByName(listName, envKeys = []) {
    const listId = await resolveListId(listName, envKeys);
    return readListItems(listId);
}

async function resolveListId(listName, envKeys = []) {
    for (const envKey of envKeys) {
        const value = process.env[envKey];
        if (value) {
            return value;
        }
    }

    const normalized = cleanText(listName).toLowerCase();
    if (!normalized) {
        return '';
    }

    if (listIdCache.has(normalized)) {
        return listIdCache.get(normalized);
    }

    const data = await graphRequest(
        'get',
        `/sites/${requireEnv('SHAREPOINT_SITE_ID')}/lists?$select=id,displayName,name`
    );

    const match = (data.value || []).find(list => {
        const displayName = cleanText(list.displayName || list.name).toLowerCase();
        return displayName === normalized;
    });

    const resolvedId = match ? match.id : '';
    if (resolvedId) {
        listIdCache.set(normalized, resolvedId);
    }
    return resolvedId;
}

function getDateKey(value) {
    if (!value) {
        return '';
    }
    return String(value).slice(0, 10);
}

function mapCompany(item) {
    return {
        id: cleanText(item.id || ''),
        title: cleanText(item.fields.Title || ''),
        companyId: cleanText(item.fields.CompanyId || item.fields.CompanyID || item.fields.Title),
        name: cleanText(item.fields.Name || item.fields.CompanyName || item.fields.Title),
        createdDateTime: cleanText(item.createdDateTime || ''),
    };
}

function mapCustomer(item) {
    return {
        id: cleanText(item.fields.AccountNum || item.fields.AccountNumber || item.fields.Title),
        accountNum: cleanText(item.fields.AccountNum || item.fields.AccountNumber || item.fields.Title),
        name: cleanText(item.fields.Name || item.fields.Title),
        title: cleanText(item.fields.Title || ''),
        customerGroup: cleanText(item.fields.CustGroup || item.fields.CustomerGroup),
        companyId: cleanText(item.fields.CompanyId || item.fields.CompanyID),
        contactPerson: cleanText(item.fields.Name || item.fields.Title),
        creditLimit: 0,
        currentBalance: 0,
        overdueAmount: 0,
        paymentTerms: 'Net 30',
        salesRep: '',
        creditStatus: 'Good',
        notes: '',
        addresses: []
    };
}

function mapOrder(item, customerMap) {
    const fields     = item.fields || {};
    const customerId = cleanText(fields.CustAccount || fields.CustomerAccount || fields.AccountNum);
    const customer   = customerMap.get(customerId);

    // Title pattern written by createSalesOrder: "SalesId | CustomerName | PO:Ref"
    const titleParts      = cleanText(fields.Title || '').split(' | ');
    const custNameInTitle = titleParts.length >= 2 ? titleParts[1].replace(/^PO:.*/, '').trim() : '';

    return {
        id:       cleanText(fields.SalesId || fields.Title || item.id),
        spItemId: item.id,
        salesId:  cleanText(fields.SalesId || fields.Title || item.id),
        customerId,
        custAccount:  customerId,
        customerName: cleanText(customer?.name || custNameInTitle || fields.CustomerName || customerId),
        customerRef:  titleParts.find(p => p.startsWith('PO:'))?.replace('PO:', '') || '',
        userEmail:    cleanText(fields.Email || fields.UserEmail || ''),
        date: getDateKey(item.createdDateTime),
        createdDateTime: cleanText(item.createdDateTime || ''),
        status:         cleanText(fields.Status        || 'In Progress'),
        deliveryTerms:  cleanText(fields.DeliveryTerms || ''),
        paymentTerms:   cleanText(fields.PaymentTerms  || ''),
        currency:       cleanText(fields.Currency      || 'AUD'),
        dateTime:       fields.DateTime || '',
        customerGroup:  cleanText(fields.CustGroup     || ''),
        invoiceAccount: cleanText(fields.InvoiceAccount || ''),
        deliveryDate: '',
        reference: '',
        total: 0,
        warehouse: '',
        salesRepId: '',
        salesRepName: '',
        lines: []
    };
}

function mapOrderLine(item) {
    const fields = item.fields || {};
    // Strip the "-Ln" suffix we append on write (e.g. "PAS-001-L1" → "PAS-001")
    const rawSalesId = cleanText(fields.SalesId || fields.OrderId || fields.Title || '');
    const orderId    = rawSalesId.replace(/-L\d+$/, '');
    return {
        id: cleanText(item.id || ''),
        orderId,
        lineNo: asNumber(fields.lineNumber || fields.LineNumber),
        itemNo: cleanText(fields.ItemId || fields.ItemNo || fields.Title),
        name: cleanText(fields.Name || fields.ItemName || fields.Title),
        category: cleanText(fields.Category || ''),
        qty: asNumber(fields.SalesQty || fields.Qty),
        unit: cleanText(fields.SalesUnit || fields.Unit || 'ea'),
        price: asNumber(fields.SalesPrice || fields.Price),
        discount: asNumber(fields.Discount || 0),
        deliveryType:    cleanText(fields.DeliveryType    || 'Stock'),
        status:          cleanText(fields.Status          || 'In Progress'),
        orderLineStatus: cleanText(fields.OrderLineStatus || 'In Progress'),
        customerId:      cleanText(fields.CustAccount     || fields.CustomerAccount || ''),
        currency:        cleanText(fields.Currency        || 'AUD'),
        customerGroup:   cleanText(fields.CustGroup       || '')
    };
}


function computeDashboardSummary(orders) {
    const now = new Date();
    const monthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const weekAgo = new Date(now);
    weekAgo.setDate(weekAgo.getDate() - 7);

    const ordersThisMonth = orders.filter(order => cleanText(order.date).startsWith(monthKey)).length;
    const ordersThisWeek  = orders.filter(order => new Date(order.date || order.createdDateTime) >= weekAgo).length;
    const revenueThisMonth = orders
        .filter(order => cleanText(order.date).startsWith(monthKey))
        .reduce((sum, order) => sum + asNumber(order.total), 0);
    const ytdRevenue = orders
        .filter(order => cleanText(order.date).startsWith(String(now.getFullYear())))
        .reduce((sum, order) => sum + asNumber(order.total), 0);

    // openOrders = orders that are not shipped or cancelled (real count, not fake "pending approval")
    const openOrders = orders.filter(order => !['Shipped','Cancelled','Invoiced'].includes(order.status)).length;

    return {
        ordersToday:        orders.filter(order => cleanText(order.date) === now.toISOString().slice(0, 10)).length,
        ordersThisWeek,
        ordersThisMonth,
        revenueThisMonth:   Number(revenueThisMonth.toFixed(2)),
        pendingApprovals:   openOrders,
        ytdRevenue:         Number(ytdRevenue.toFixed(2)),
    };
}

async function getBootstrapData(user = {}) {
    if (!isConfigured()) {
        throw new Error('SharePoint is not configured');
    }

    const companyListName = cleanText(process.env.SHAREPOINT_COMPANY_LIST_NAME || 'Company') || 'Company';
    const customersListName = cleanText(process.env.SHAREPOINT_ALL_CUSTOMERS_LIST_NAME || 'AllCustomers') || 'AllCustomers';
    const ordersListName = cleanText(process.env.SHAREPOINT_SALES_ORDER_HEADER_LIST_NAME || 'SalesOrderHeader') || 'SalesOrderHeader';
    const orderLinesListName = cleanText(process.env.SHAREPOINT_SALES_ORDER_LINES_LIST_NAME || 'SalesOrderLines') || 'SalesOrderLines';

    const companiesListId = await resolveListId(companyListName, ['SHAREPOINT_COMPANY_LIST_ID']);
    const customersListId = await resolveListId(customersListName, ['SHAREPOINT_ALL_CUSTOMERS_LIST_ID', 'SHAREPOINT_CUSTOMERS_LIST_ID']);
    const ordersListId = await resolveListId(ordersListName, ['SHAREPOINT_SALES_ORDER_HEADER_LIST_ID', 'SHAREPOINT_ORDERS_LIST_ID']);
    const orderLinesListId = await resolveListId(orderLinesListName, ['SHAREPOINT_SALES_ORDER_LINES_LIST_ID', 'SHAREPOINT_ORDER_LINES_LIST_ID']);

    const [companyRows, customerRows, orderRows, orderLineRows] = await Promise.all([
        readListItems(companiesListId),
        readListItems(customersListId),
        readListItems(ordersListId),
        readListItems(orderLinesListId)
    ]);

    const companies = companyRows.map(mapCompany);
    let customers = customerRows.map(mapCustomer);
    const customerMap = new Map(customers.map(customer => [customer.id, customer]));
    let salesOrders = orderRows.map(item => mapOrder(item, customerMap));

    // Filter by current user email
    const currentUserEmail = cleanText(user?.preferred_username || user?.email || '').toLowerCase();
    if (user?.role !== 'admin' && currentUserEmail) {
        salesOrders = salesOrders.filter(order => (order.userEmail || '').toLowerCase() === currentUserEmail);
    }

    const orderLines = orderLineRows.map(mapOrderLine);

    // If customer master list is empty, derive lightweight customer records from order history.
    if (customers.length === 0) {
        const seenCustomers = new Set();
        customers = salesOrders
            .filter(order => order.customerId)
            .map(order => ({
                id: order.customerId,
                accountNum: order.customerId,
                name: order.customerName || order.customerId,
                title: order.customerName || order.customerId,
                customerGroup: order.customerGroup || '',
                companyId: '',
                contactPerson: order.customerName || order.customerId,
                creditLimit: 0,
                currentBalance: 0,
                overdueAmount: 0,
                paymentTerms: 'Net 30',
                salesRep: '',
                creditStatus: 'Good',
                notes: 'Derived from SalesOrderHeader.',
                addresses: []
            }))
            .filter(customer => {
                if (seenCustomers.has(customer.id)) return false;
                seenCustomers.add(customer.id);
                return true;
            });
    }

    const orderMap = new Map(salesOrders.map(order => [order.id, order]));
    orderLines.forEach(line => {
        const order = orderMap.get(line.orderId);
        if (order) {
            order.lines.push({
                lineNo: line.lineNo,
                itemNo: line.itemNo,
                name: line.name,
                category: line.category,
                qty: line.qty,
                unit: line.unit,
                price: line.price,
                discount: line.discount,
                deliveryType: line.deliveryType
            });
        }
    });

    // Compute order total from attached lines
    salesOrders.forEach(order => {
        if (order.lines.length > 0) {
            order.total = Number(order.lines.reduce((s, l) => {
                return s + (l.price * l.qty * (1 - (l.discount || 0) / 100));
            }, 0).toFixed(2));
        }
    });

    return {
        sharepointConfigured: true,
        salesRep: buildSalesRep(user),
        company: companies[0] || null,
        companies,
        customer: customers[0] || null,
        customers,
        salesOrders,
        dashboardSummary: computeDashboardSummary(salesOrders)
    };
}

async function createSalesOrder(orderPayload, user = {}) {
    if (!isConfigured()) {
        return {
            success: true,
            sharepointConfigured: false,
            message: 'SharePoint is not configured. Order was not saved to SharePoint.',
            orderId: `PAS${Math.floor(Math.random() * 100000)}`
        };
    }

    const ordersListId = await resolveListId('SalesOrderHeader', ['SHAREPOINT_SALES_ORDER_HEADER_LIST_ID', 'SHAREPOINT_ORDERS_LIST_ID']);
    const orderLinesListId = await resolveListId('SalesOrderLines', ['SHAREPOINT_SALES_ORDER_LINES_LIST_ID', 'SHAREPOINT_ORDER_LINES_LIST_ID']);
    const salesId = cleanText(orderPayload.salesId || orderPayload.id || generateId('PAS'));
    const custName = cleanText(orderPayload.customerName || orderPayload.customerId || orderPayload.custAccount);
    const custRef  = cleanText(orderPayload.customerRef || orderPayload.reference || '');
    const orderTitle = custRef
        ? `${salesId} | ${custName} | PO:${custRef}`
        : `${salesId} | ${custName}`;

    const orderFields = {
        Title:          orderTitle,
        SalesId:        salesId,
        CustAccount:    cleanText(orderPayload.customerId   || orderPayload.custAccount),
        Currency:       cleanText(orderPayload.currency     || 'AUD'),
        CustGroup:      cleanText(orderPayload.customerGroup || orderPayload.custGroup),
        InvoiceAccount: cleanText(orderPayload.invoiceAccount || orderPayload.customerId || orderPayload.custAccount),
        Status:         cleanText(orderPayload.status        || 'In Progress'),
        DeliveryTerms:  cleanText(orderPayload.deliveryTerms || ''),
        PaymentTerms:   cleanText(orderPayload.paymentTerms  || ''),
        Email:          cleanText(user?.preferred_username || user?.email || ''),
    };

    if (ordersListId) {
        await graphRequest(
            'post',
            `/sites/${requireEnv('SHAREPOINT_SITE_ID')}/lists/${ordersListId}/items`,
            { fields: orderFields }
        );
    }

    if (orderLinesListId && Array.isArray(orderPayload.lines)) {
        for (const [index, line] of orderPayload.lines.entries()) {
            // SalesId column has unique constraint — store as "ORDERID-Ln" so each row is unique.
            // mapOrderLine strips the "-Ln" suffix back to get the parent orderId.
            const lineSalesId = `${salesId}-L${index + 1}`;
            const lineTitle   = cleanText(line.name || line.itemNo)
                ? `${cleanText(line.name || '')} (${cleanText(line.itemNo || '')})`.replace(/^\s*\(/, '(').trim()
                : lineSalesId;
            const lineFields = {
                Title:           lineTitle,
                SalesId:         lineSalesId,
                lineNumber:      index + 1,
                CustAccount:     cleanText(orderPayload.customerId || orderPayload.custAccount),
                Currency:        cleanText(orderPayload.currency   || 'AUD'),
                CustGroup:       cleanText(orderPayload.customerGroup || orderPayload.custGroup),
                SalesUnit:       cleanText(line.unit               || 'ea'),
                SalesPrice:      asNumber(line.unitPrice           || line.price),
                SalesQty:        asNumber(line.qty),
                OrderLineStatus: cleanText(line.orderLineStatus    || line.status || 'In Progress'),
                Email:           cleanText(user?.preferred_username || user?.email || ''),
            };

            await graphRequest(
                'post',
                `/sites/${requireEnv('SHAREPOINT_SITE_ID')}/lists/${orderLinesListId}/items`,
                { fields: lineFields }
            );
        }
    }

    return {
        success: true,
        sharepointConfigured: true,
        message: 'Order created successfully in SharePoint',
        orderId: salesId,
        order: orderFields
    };
}

module.exports = {
    isConfigured,
    getBootstrapData,
    createSalesOrder,
    generateId,
    getRawListItemsByName,
    // Generic list CRUD helpers
    createListItem: async function(listName, fields) {
        const listId = await resolveListId(listName);
        if (!listId) throw new Error(`List ${listName} not found`);
        const result = await graphRequest('post', `/sites/${requireEnv('SHAREPOINT_SITE_ID')}/lists/${listId}/items`, { fields });
        return result;
    },
    updateListItem: async function(listName, itemId, fields) {
        const listId = await resolveListId(listName);
        if (!listId) throw new Error(`List ${listName} not found`);
        const result = await graphRequest('patch', `/sites/${requireEnv('SHAREPOINT_SITE_ID')}/lists/${listId}/items/${itemId}/fields`, fields);
        return result;
    },
    deleteListItem: async function(listName, itemId) {
        const listId = await resolveListId(listName);
        if (!listId) throw new Error(`List ${listName} not found`);
        const result = await graphRequest('delete', `/sites/${requireEnv('SHAREPOINT_SITE_ID')}/lists/${listId}/items/${itemId}`);
        return result;
    },
    findItemByField: async function(listName, fieldKey, value) {
        const listId = await resolveListId(listName);
        if (!listId) return null;
        const rows = await readListItems(listId);
        const found = rows.find(r => {
            const v = r.fields && r.fields[fieldKey];
            return String(v || '').toLowerCase() === String(value || '').toLowerCase();
        });
        return found || null;
    },
    getListColumns: async function(listName) {
        const listId = await resolveListId(listName);
        if (!listId) return [];
        const data = await graphRequest(
            'get',
            `/sites/${requireEnv('SHAREPOINT_SITE_ID')}/lists/${listId}/columns?$select=displayName,name,type,readOnly,hidden`
        );
        return (data.value || []).filter(c => !c.hidden).map(c => ({
            displayName: c.displayName,
            internalName: c.name,
            type: c.type,
            readOnly: c.readOnly
        }));
    }
};