/* ============================================================
   Dynamics 365 F&O — OData API Client (d365.js)
   Auth: OAuth 2.0 client_credentials (v1 endpoint + resource)
   Entities: CustomersV3, ReleasedProductsV2,
             SalesOrderHeadersV2, SalesOrderLines
   ============================================================ */

'use strict';

const axios = require('axios');

const TENANT_ID     = process.env.D365_TENANT_ID;
const CLIENT_ID     = process.env.D365_CLIENT_ID;
const CLIENT_SECRET = process.env.D365_CLIENT_SECRET;
const RESOURCE      = process.env.D365_RESOURCE || 'https://dynaone.sandbox.operations.dynamics.com';

let _cachedToken  = null;
let _tokenExpires = 0;

// Full customer list cache (5-minute TTL) — D365 sandbox is ~2.5s cold
let _custCache    = null;
let _custCacheExp = 0;
const CUST_TTL_MS = 5 * 60 * 1000;

// Full product list cache (10-minute TTL)
let _prodCache    = null;
let _prodCacheExp = 0;
const PROD_TTL_MS = 10 * 60 * 1000;

// Fields to fetch from ReleasedProductsV2
const PRODUCT_SELECT = [
  'ItemNumber', 'ProductNumber', 'SearchName', 'ProductSearchName',
  'ProductGroupId', 'ProductType', 'ProductSubType',
  'SalesPrice', 'UnitCost', 'PurchasePrice',
  'SalesUnitSymbol', 'InventoryUnitSymbol', 'PurchaseUnitSymbol',
  'dataAreaId',
].join(',');

// Fields to fetch from SalesOrderHeadersV2
const SO_HEADER_SELECT = [
  'SalesOrderNumber', 'OrderingCustomerAccountNumber', 'SalesOrderName',
  'CurrencyCode', 'SalesOrderProcessingStatus', 'SalesOrderStatus',
  'RequestedShippingDate', 'OrderCreationDateTime',
  'dataAreaId', 'OrderTotalAmount', 'PaymentTermsName',
  'InvoiceCustomerAccountNumber', 'DeliveryTermsCode', 'DeliveryModeCode',
].join(',');

/* ----------------------------------------------------------
   getToken — OAuth 2.0 client_credentials (D365 v1 endpoint)
   Token is cached until 60 s before expiry
   ---------------------------------------------------------- */
async function getToken() {
  if (_cachedToken && Date.now() < _tokenExpires - 60_000) return _cachedToken;

  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('D365 credentials not configured in .env');
  }

  const params = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     CLIENT_ID,
    client_secret: CLIENT_SECRET,
    resource:      RESOURCE,
  });

  const res = await axios.post(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/token`,
    params.toString(),
    { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
  );

  _cachedToken  = res.data.access_token;
  _tokenExpires = Date.now() + res.data.expires_in * 1000;
  console.info('[D365] Token refreshed, expires in', res.data.expires_in, 's');
  return _cachedToken;
}

/* ----------------------------------------------------------
   _get — raw OData GET with auth header
   ---------------------------------------------------------- */
async function _get(entity, params = {}) {
  const token = await getToken();
  const res = await axios.get(`${RESOURCE}/data/${entity}`, {
    params,
    headers: {
      Authorization:       `Bearer ${token}`,
      'OData-MaxVersion':  '4.0',
      'OData-Version':     '4.0',
      Accept:              'application/json;odata.metadata=minimal',
    },
  });
  return res.data;
}

/* ----------------------------------------------------------
   _post — raw OData POST (create)
   ---------------------------------------------------------- */
async function _post(entity, body) {
  const token = await getToken();
  const res = await axios.post(`${RESOURCE}/data/${entity}`, body, {
    headers: {
      Authorization:       `Bearer ${token}`,
      'OData-MaxVersion':  '4.0',
      'OData-Version':     '4.0',
      Accept:              'application/json;odata.metadata=minimal',
      'Content-Type':      'application/json',
    },
  });
  return res.data;
}

/* ----------------------------------------------------------
   _patch — raw OData PATCH (update)
   entityWithKey e.g. "SalesOrderHeadersV2(SalesOrderNumber='000811',dataAreaId='PASCAL')"
   ---------------------------------------------------------- */
async function _patch(entityWithKey, body) {
  const token = await getToken();
  const res = await axios.patch(`${RESOURCE}/data/${entityWithKey}`, body, {
    headers: {
      Authorization:       `Bearer ${token}`,
      'OData-MaxVersion':  '4.0',
      'OData-Version':     '4.0',
      Accept:              'application/json;odata.metadata=minimal',
      'Content-Type':      'application/json',
    },
  });
  return res.data;
}

/* ----------------------------------------------------------
   getDefaultDataAreaId — read from env or derive from cache
   ---------------------------------------------------------- */
function getDefaultDataAreaId() {
  if (process.env.D365_DATA_AREA_ID) return process.env.D365_DATA_AREA_ID;
  if (_prodCache && _prodCache.length > 0 && _prodCache[0].dataAreaId) {
    return _prodCache[0].dataAreaId;
  }
  return '';
}

/* ----------------------------------------------------------
   getCustomers — search / paginate CustomersV3
   ---------------------------------------------------------- */
async function _getAllCustomersCached() {
  if (_custCache && Date.now() < _custCacheExp) return _custCache;
  const data = await _get('CustomersV3', { '$top': 500 });
  _custCache    = data.value || [];
  _custCacheExp = Date.now() + CUST_TTL_MS;
  console.info(`[D365] Customer cache refreshed — ${_custCache.length} records, TTL 5 min`);
  return _custCache;
}

async function getCustomers({ top = 20, skip = 0, search = '', account = '' } = {}) {
  // Exact account lookup — OData eq filter is supported
  if (account) {
    return _get('CustomersV3', {
      '$top': 1,
      '$filter': `CustomerAccount eq '${account.replace(/'/g, "''")}'`,
    });
  }

  // All other cases: use cached full list and slice / filter in Node.js.
  // D365 F&O rejects contains()/startswith() with "not Queryable" on string fields.
  const all = await _getAllCustomersCached();

  if (search) {
    const q = search.toLowerCase();
    const filtered = all.filter(c =>
      (c.OrganizationName || '').toLowerCase().includes(q) ||
      (c.CustomerAccount  || '').toLowerCase().startsWith(q) ||
      (c.NameAlias        || '').toLowerCase().includes(q)
    );
    return { value: filtered.slice(skip, skip + top) };
  }

  return { value: all.slice(skip, skip + top) };
}

/* ----------------------------------------------------------
   getCustomer — single customer by account number
   ---------------------------------------------------------- */
async function getCustomer(accountId) {
  const data = await getCustomers({ account: accountId, top: 1 });
  const records = data.value || [];
  if (records.length === 0) throw new Error(`Customer '${accountId}' not found`);
  return records[0];
}

/* ----------------------------------------------------------
   getProducts — search / paginate ReleasedProductsV2
   Uses in-memory cache + filter (same pattern as customers)
   ---------------------------------------------------------- */
async function _getAllProductsCached() {
  if (_prodCache && Date.now() < _prodCacheExp) return _prodCache;
  const data = await _get('ReleasedProductsV2', {
    '$top':    1000,
    '$select': PRODUCT_SELECT,
  });
  _prodCache    = data.value || [];
  _prodCacheExp = Date.now() + PROD_TTL_MS;
  console.info(`[D365] Product cache refreshed — ${_prodCache.length} records, TTL 10 min`);
  return _prodCache;
}

async function getProducts({ top = 20, skip = 0, search = '', itemNumber = '' } = {}) {
  // Exact item lookup
  if (itemNumber) {
    const data = await _get('ReleasedProductsV2', {
      '$top':    1,
      '$select': PRODUCT_SELECT,
      '$filter': `ItemNumber eq '${itemNumber.replace(/'/g, "''")}'`,
    });
    return data;
  }

  // In-memory search from cache
  const all = await _getAllProductsCached();

  if (search) {
    const q = search.toLowerCase();
    const filtered = all.filter(p =>
      (p.ItemNumber       || '').toLowerCase().startsWith(q) ||
      (p.SearchName       || '').toLowerCase().includes(q)   ||
      (p.ProductSearchName|| '').toLowerCase().includes(q)
    );
    return { value: filtered.slice(skip, skip + top) };
  }

  return { value: all.slice(skip, skip + top) };
}

/* ----------------------------------------------------------
   getProduct — single product by item number
   ---------------------------------------------------------- */
async function getProduct(itemNumber) {
  const data = await getProducts({ itemNumber, top: 1 });
  const records = data.value || [];
  if (records.length === 0) throw new Error(`Product '${itemNumber}' not found`);
  return records[0];
}

/* ----------------------------------------------------------
   getSalesOrders — list / search SalesOrderHeadersV2
   ---------------------------------------------------------- */
async function getSalesOrders({ top = 50, skip = 0, salesOrderNumber = '', customerAccount = '' } = {}) {
  const params = { '$top': top, '$skip': skip, '$select': SO_HEADER_SELECT };
  if (salesOrderNumber) {
    params['$filter'] = `SalesOrderNumber eq '${salesOrderNumber.replace(/'/g, "''")}'`;
  } else if (customerAccount) {
    params['$filter'] = `OrderingCustomerAccountNumber eq '${customerAccount.replace(/'/g, "''")}'`;
  }
  return _get('SalesOrderHeadersV2', params);
}

/* ----------------------------------------------------------
   createSalesOrderHeader — POST to SalesOrderHeadersV2
   Returns the created record including D365-assigned SalesOrderNumber
   ---------------------------------------------------------- */
async function createSalesOrderHeader({ customerAccount, orderName, currency, paymentTerms, dataAreaId } = {}) {
  const area = dataAreaId || getDefaultDataAreaId();
  const body = {
    dataAreaId:                    area,
    OrderingCustomerAccountNumber: customerAccount,
    SalesOrderName:                orderName || customerAccount,
    CurrencyCode:                  currency  || 'AUD',
  };
  if (paymentTerms) body.PaymentTermsName = paymentTerms;
  const result = await _post('SalesOrderHeadersV2', body);
  console.info(`[D365] Sales order header created: ${result.SalesOrderNumber}`);
  return result;
}

/* ----------------------------------------------------------
   updateSalesOrderHeader — PATCH SalesOrderHeadersV2
   ---------------------------------------------------------- */
async function updateSalesOrderHeader(salesOrderNumber, areaId, updates) {
  const area = areaId || getDefaultDataAreaId();
  const key  = `SalesOrderHeadersV2(SalesOrderNumber='${salesOrderNumber.replace(/'/g, "''")}',dataAreaId='${area}')`;
  return _patch(key, updates);
}

/* ----------------------------------------------------------
   getSalesOrderLines — GET SalesOrderLines by order number
   ---------------------------------------------------------- */
async function getSalesOrderLines(salesOrderNumber) {
  if (!salesOrderNumber) throw new Error('salesOrderNumber is required');
  return _get('SalesOrderLines', {
    '$filter': `SalesOrderNumber eq '${salesOrderNumber.replace(/'/g, "''")}'`,
    '$top':    200,
  });
}

/* ----------------------------------------------------------
   createSalesOrderLine — POST to SalesOrderLines
   Returns the created line including InventoryLotId (used for PATCH key)
   ---------------------------------------------------------- */
async function createSalesOrderLine({ salesOrderNumber, itemNumber, salesPrice, quantity, siteId, warehouseId, dataAreaId } = {}) {
  const area = dataAreaId || getDefaultDataAreaId();
  const body = {
    dataAreaId:           area,
    SalesOrderNumber:     salesOrderNumber,
    ItemNumber:           itemNumber,
    SalesPrice:           salesPrice  || 0,
    OrderedSalesQuantity: quantity    || 1,
  };
  if (siteId)      body.ShippingSiteId      = siteId;
  if (warehouseId) body.ShippingWarehouseId = warehouseId;
  return _post('SalesOrderLines', body);
}

/* ----------------------------------------------------------
   updateSalesOrderLine — PATCH SalesOrderLines by InventoryLotId
   ---------------------------------------------------------- */
async function updateSalesOrderLine(inventoryLotId, areaId, updates) {
  const area = areaId || getDefaultDataAreaId();
  const key  = `SalesOrderLines(InventoryLotId='${inventoryLotId}',dataAreaId='${area}')`;
  return _patch(key, updates);
}

module.exports = {
  getToken,
  getCustomers, getCustomer,
  getProducts,  getProduct,
  getSalesOrders, createSalesOrderHeader, updateSalesOrderHeader,
  getSalesOrderLines, createSalesOrderLine, updateSalesOrderLine,
  getDefaultDataAreaId,
};
