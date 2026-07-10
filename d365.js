/* ============================================================
   Dynamics 365 F&O — OData API Client (d365.js)
   Auth: OAuth 2.0 client_credentials (v1 endpoint + resource)
   Entities: CustomersV3, ReleasedProductsV2,
             SalesOrderHeadersV2, SalesOrderLines
   ============================================================ */

'use strict';


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

// Units of measure cache (1-hour TTL)
let _unitsCache    = null;
let _unitsCacheExp = 0;
const UNITS_TTL_MS = 60 * 60 * 1000;

// Fields to fetch from ReleasedProductsV2
const PRODUCT_SELECT = [
  'ItemNumber', 'ProductNumber', 'SearchName', 'ProductSearchName',
  'ProductGroupId', 'ProductType', 'ProductSubType',
  'SalesPrice', 'UnitCost', 'PurchasePrice',
  'SalesUnitSymbol', 'InventoryUnitSymbol', 'PurchaseUnitSymbol',
  'dataAreaId',
  'SalesLineDiscountProductGroupCode',
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

  const res  = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });
  if (!res.ok) throw new Error(`[D365] Token fetch failed: ${res.status}`);
  const data = await res.json();

  _cachedToken  = data.access_token;
  _tokenExpires = Date.now() + data.expires_in * 1000;
  console.info('[D365] Token refreshed, expires in', data.expires_in, 's');
  return _cachedToken;
}

/* ----------------------------------------------------------
   _get — raw OData GET with auth header
   ---------------------------------------------------------- */
async function _get(entity, params = {}) {
  const token = await getToken();
  const qs    = new URLSearchParams(params).toString();
  const url   = `${RESOURCE}/data/${entity}${qs ? '?' + qs : ''}`;
  const res   = await fetch(url, {
    headers: {
      Authorization:       `Bearer ${token}`,
      'OData-MaxVersion':  '4.0',
      'OData-Version':     '4.0',
      Accept:              'application/json;odata.metadata=minimal',
    },
  });
  if (!res.ok) { const t = await res.text(); throw new Error(`[D365] GET ${entity} ${res.status}: ${t}`); }
  return res.json();
}

/* ----------------------------------------------------------
   _post — raw OData POST (create)
   ---------------------------------------------------------- */
async function _post(entity, body) {
  const token = await getToken();
  const res   = await fetch(`${RESOURCE}/data/${entity}`, {
    method:  'POST',
    headers: {
      Authorization:       `Bearer ${token}`,
      'OData-MaxVersion':  '4.0',
      'OData-Version':     '4.0',
      Accept:              'application/json;odata.metadata=minimal',
      'Content-Type':      'application/json',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) { const t = await res.text(); throw new Error(`[D365] POST ${entity} ${res.status}: ${t}`); }
  return res.json();
}

/* ----------------------------------------------------------
   _patch — raw OData PATCH (update)
   entityWithKey e.g. "SalesOrderHeadersV2(SalesOrderNumber='000811',dataAreaId='PASCAL')"
   ---------------------------------------------------------- */
async function _patch(entityWithKey, body) {
  const token = await getToken();
  const res   = await fetch(`${RESOURCE}/data/${entityWithKey}`, {
    method:  'PATCH',
    headers: {
      Authorization:       `Bearer ${token}`,
      'OData-MaxVersion':  '4.0',
      'OData-Version':     '4.0',
      Accept:              'application/json;odata.metadata=minimal',
      'Content-Type':      'application/json',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) { const t = await res.text(); throw new Error(`[D365] PATCH ${entityWithKey} ${res.status}: ${t}`); }
  if (res.status === 204) return {};
  return res.json();
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

  // Fetch products + translations in parallel
  const [prodData, transData] = await Promise.all([
    _get('ReleasedProductsV2', { '$top': 1000, '$select': PRODUCT_SELECT }),
    _get('ProductTranslations', {
      '$top':    5000,
      '$select': 'ProductNumber,ProductName,LanguageId',
      '$filter': "LanguageId eq 'en-us'",
    }).catch(() => ({ value: [] })),
  ]);

  // Build translation map: ProductNumber → ProductName
  const nameMap = {};
  (transData.value || []).forEach(t => {
    if (t.ProductName) nameMap[t.ProductNumber] = t.ProductName;
  });

  // Merge ProductName into each product record
  _prodCache = (prodData.value || []).map(p => ({
    ...p,
    ProductName: nameMap[p.ProductNumber] || p.SearchName || p.ItemNumber,
  }));

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

/* ----------------------------------------------------------
   SalesPriceAgreements — 4-case price cascade
   ---------------------------------------------------------- */
const PRICE_SELECT = 'ItemNumber,Price,PriceCurrencyCode,CustomerAccountNumber,PriceCustomerGroupCode,FromQuantity,ToQuantity,QuantityUnitySymbol';

function _esc(s) { return String(s || '').replace(/'/g, "''"); }

async function getSalesPriceAgreements({ itemNumber, customerAccount = '', customerGroupCodes = [], customerGroupCode = '', quantity = 1, caseOnly = 0 } = {}) {
  if (!itemNumber) throw new Error('itemNumber is required');

  // Normalise group codes — accept both array and legacy single string
  const groupCodes = [...new Set(
    (Array.isArray(customerGroupCodes) ? customerGroupCodes : [customerGroupCodes])
      .concat(customerGroupCode ? [customerGroupCode] : [])
      .filter(Boolean)
  )];

  // Run all 4 cases in parallel
  const [r1, r2all, r3, r4] = await Promise.all([
    // Case 1 — D365 smart lookup: CustomerAccountNumber eq 'US-010' resolves both
    //           direct customer agreements AND group agreements the customer belongs to.
    //           orderby CustomerAccountNumber desc puts direct matches first.
    customerAccount ? _get('SalesPriceAgreements', {
      '$filter':  `ItemNumber eq '${_esc(itemNumber)}' and CustomerAccountNumber eq '${_esc(customerAccount)}' and FromQuantity le ${quantity}`,
      '$select':  PRICE_SELECT,
      '$orderby': 'CustomerAccountNumber desc,FromQuantity desc',
      '$top':     1,
    }) : Promise.resolve({ value: [] }),

    // Case 2 — all group prices for this item (reverse lookup)
    // Also covers the case where Case 1 resolved a PriceCustomerGroupCode
    _get('SalesPriceAgreements', {
      '$filter':  `ItemNumber eq '${_esc(itemNumber)}' and CustomerAccountNumber eq '' and PriceCustomerGroupCode ne ''`,
      '$select':  PRICE_SELECT,
      '$orderby': 'FromQuantity desc',
    }),

    // Case 3 — global / all-customer price
    _get('SalesPriceAgreements', {
      '$filter':  `ItemNumber eq '${_esc(itemNumber)}' and CustomerAccountNumber eq '' and PriceCustomerGroupCode eq ''`,
      '$select':  PRICE_SELECT,
      '$orderby': 'FromQuantity desc',
      '$top':     1,
    }),

    // Case 4 — product base price from ReleasedProductsV2
    _get('ReleasedProductsV2', {
      '$filter': `ItemNumber eq '${_esc(itemNumber)}'`,
      '$select': 'ItemNumber,SalesPrice,SalesPriceQuantity,SalesUnitSymbol,UnitCost,PurchasePrice,BaseSalesPriceSource',
      '$top':    1,
    }),
  ]);

  const case1 = (r1.value || [])[0] || null;

  if (caseOnly === 1) {
    return case1 ? { source: case1.CustomerAccountNumber ? 'customer' : 'group', record: case1 } : { source: 'none', record: null };
  }

  // Case 1 — direct customer-specific agreement (CustomerAccountNumber = custAccount)
  if (!caseOnly || caseOnly === 1) {
    if (case1 && case1.CustomerAccountNumber === customerAccount) {
      return { source: 'customer', record: case1 };
    }
  }

  // Case 2 — group price lookup
  // Sources: (a) D365 smart lookup returned a group agreement in Case 1
  //          (b) reverse lookup across all item group agreements using known group codes
  if (!caseOnly || caseOnly === 2) {
    // (a) Case 1 resolved a group agreement for this customer (D365 smart lookup)
    if (case1 && !case1.CustomerAccountNumber && case1.PriceCustomerGroupCode) {
      return { source: 'group', record: case1 };
    }
    // (b) Reverse lookup: match customer's known group codes against all group agreements
    const groupAgreements = r2all.value || [];
    const allCodes = [...new Set([...groupCodes,
      case1?.PriceCustomerGroupCode,
    ].filter(Boolean))];
    for (const code of allCodes) {
      const match = groupAgreements
        .filter(r => r.PriceCustomerGroupCode === code && (r.FromQuantity || 0) <= quantity)
        .sort((a, b) => (b.FromQuantity || 0) - (a.FromQuantity || 0))[0];
      if (match) return { source: 'group', record: match };
    }
    if (caseOnly === 2) return { source: 'none', record: null };
  }

  // Case 3 — global
  if (!caseOnly || caseOnly === 3) {
    const case3 = (r3.value || [])[0] || null;
    if (case3) return { source: 'global', record: case3 };
    if (caseOnly === 3) return { source: 'none', record: null };
  }

  // Case 4 — base product price
  if (!caseOnly || caseOnly === 4) {
    const prod = (r4.value || [])[0] || null;
    if (prod) {
      const effectivePrice = prod.SalesPrice || prod.PurchasePrice || prod.UnitCost || 0;
      return { source: 'base', record: { ...prod, Price: effectivePrice, _salesPrice: prod.SalesPrice, _purchasePrice: prod.PurchasePrice, _unitCost: prod.UnitCost, PriceCurrencyCode: '', QuantityUnitySymbol: prod.SalesUnitSymbol } };
    }
  }

  return { source: 'none', record: null };
}

async function getAllPriceAgreementsForItem(itemNumber) {
  if (!itemNumber) throw new Error('itemNumber is required');
  return _get('SalesPriceAgreements', {
    '$filter':  `ItemNumber eq '${_esc(itemNumber)}'`,
    '$orderby': 'CustomerAccountNumber desc,FromQuantity desc',
    '$top':     100,
  });
}

/* ----------------------------------------------------------
   getSalesLineDiscounts — 9-case line discount cascade
   Runs all 9 Party×Product combinations against
   SalesLineDiscountAgreements, sums DiscountPercentage1.
   ---------------------------------------------------------- */
const DISC_SELECT = [
  'ItemNumber', 'LineDiscountProductGroupCode',
  'CustomerAccountNumber', 'LineDiscountCustomerGroupCode',
  'DiscountPercentage1', 'DiscountPercentage2', 'DiscountAmount',
  'FromQuantity', 'ToQuantity',
  'DiscountApplicableFromDate', 'DiscountApplicableToDate', 'DiscountCurrencyCode',
].join(',');

async function getSalesLineDiscounts({ itemNumber, customerAccount = '', customerGroupCode = '', productGroupCode = '' } = {}) {
  if (!itemNumber) throw new Error('itemNumber is required');
  const a = _esc(customerAccount);
  const cg = _esc(customerGroupCode);
  const i = _esc(itemNumber);
  const pg = _esc(productGroupCode);
  const _f = filter => _get('SalesLineDiscountAgreements', { '$filter': filter, '$select': DISC_SELECT, '$top': 50 });
  const empty = { value: [] };

  const results = await Promise.allSettled([
    // Case 1: Cust=Table + Item=Table
    customerAccount                         ? _f(`CustomerAccountNumber eq '${a}' and ItemNumber eq '${i}'`) : empty,
    // Case 2: Cust=Group + Item=Table
    customerGroupCode                       ? _f(`LineDiscountCustomerGroupCode eq '${cg}' and ItemNumber eq '${i}'`) : empty,
    // Case 3: Cust=All + Item=Table
    _f(`LineDiscountCustomerGroupCode eq '' and CustomerAccountNumber eq '' and ItemNumber eq '${i}'`),
    // Case 4: Cust=Table + Item=Group
    customerAccount && productGroupCode     ? _f(`CustomerAccountNumber eq '${a}' and LineDiscountProductGroupCode eq '${pg}'`) : empty,
    // Case 5: Cust=Group + Item=Group
    customerGroupCode && productGroupCode   ? _f(`LineDiscountCustomerGroupCode eq '${cg}' and LineDiscountProductGroupCode eq '${pg}'`) : empty,
    // Case 6: Cust=All + Item=Group
    productGroupCode                        ? _f(`CustomerAccountNumber eq '' and LineDiscountCustomerGroupCode eq '' and LineDiscountProductGroupCode eq '${pg}'`) : empty,
    // Case 7: Cust=Table + Item=All
    customerAccount                         ? _f(`CustomerAccountNumber eq '${a}' and LineDiscountCustomerGroupCode eq '' and ItemNumber eq '' and LineDiscountProductGroupCode eq ''`) : empty,
    // Case 8: Cust=Group + Item=All
    customerGroupCode                       ? _f(`LineDiscountCustomerGroupCode eq '${cg}' and CustomerAccountNumber eq '' and ItemNumber eq '' and LineDiscountProductGroupCode eq ''`) : empty,
    // Case 9: Cust=All + Item=All
    _f(`LineDiscountCustomerGroupCode eq '' and CustomerAccountNumber eq '' and ItemNumber eq '' and LineDiscountProductGroupCode eq ''`),
  ]);

  const allRecords = results.flatMap(r => (r.status === 'fulfilled' ? r.value.value : null) || []);
  const totalDiscountPct = allRecords.reduce((sum, r) => sum + (parseFloat(r.DiscountPercentage1) || 0), 0);
  const totalDiscountAmt = allRecords.reduce((sum, r) => sum + (parseFloat(r.DiscountAmount) || 0), 0);
  return {
    totalDiscount:    +totalDiscountPct.toFixed(4),
    totalDiscountAmt: +totalDiscountAmt.toFixed(4),
    records: allRecords,
  };
}

async function getUnitsOfMeasure() {
  if (_unitsCache && Date.now() < _unitsCacheExp) return _unitsCache;
  const res = await _get('UnitsOfMeasure', {
    $select: 'UnitSymbol,UnitDescription,DecimalPrecision'
  });
  _unitsCache    = res.value || [];
  _unitsCacheExp = Date.now() + UNITS_TTL_MS;
  return _unitsCache;
}

module.exports = {
  getToken,
  getCustomers, getCustomer,
  getProducts,  getProduct,
  getSalesOrders, createSalesOrderHeader, updateSalesOrderHeader,
  getSalesOrderLines, createSalesOrderLine, updateSalesOrderLine,
  getSalesPriceAgreements, getAllPriceAgreementsForItem,
  getSalesLineDiscounts,
  getDefaultDataAreaId,
  getUnitsOfMeasure,
};
