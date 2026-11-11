// ════════════════════════════════════════════════════════════════════
//  Stock Portfolio Tracker — Google Apps Script Backend  (Code.gs)
//  Deploy: Extensions → Apps Script → Deploy → New Deployment
//    Type: Web App | Execute as: Me | Access: Anyone
// ════════════════════════════════════════════════════════════════════

const SHEET_DEF = {
  ACCOUNTS:    { name: 'Accounts',         headers: ['id','name','created_at'] },
  BROKERS:     { name: 'Brokers',          headers: ['id','account_id','name','currency','created_at'] },
  CASH:        { name: 'Cash_Balances',    headers: ['id','broker_id','balance','updated_at'] },
  HOLDINGS:    { name: 'Holdings',         headers: ['id','broker_id','symbol','quantity','avg_cost','updated_at'] },
  PIGGY_BANKS: { name: 'Piggy_Banks',      headers: ['id','name','goal_amount','currency','description','created_at'] },
  GOAL_ALLOCS: { name: 'Goal_Allocations', headers: ['id','piggy_bank_id','broker_id','symbol','quantity','updated_at'] },
  ORDERS:      { name: 'Orders',           headers: ['id','account_id','broker_id','symbol','order_type','exec_type',
                                                     'quantity','price','fees','net_total','avg_cost','tax_term',
                                                     'status','piggy_bank_id','realized_pnl','notes','created_at','updated_at'] },
  WATCHLIST:   { name: 'Watchlist',        headers: ['id','symbol','notes','added_at'] },
};

/* ─────────────────────────────── Utilities ─────────────────────────────── */

const generateId = () => Utilities.getUuid();
const nowISO     = () => new Date().toISOString();

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function wb() { return SpreadsheetApp.getActiveSpreadsheet(); }

/* ──────────────────────────── Sheet Bootstrap ──────────────────────────── */

function initSheets() {
  const ss = wb();
  Object.values(SHEET_DEF).forEach(cfg => {
    if (ss.getSheetByName(cfg.name)) return;
    const sh = ss.insertSheet(cfg.name);
    const hdr = sh.getRange(1, 1, 1, cfg.headers.length);
    hdr.setValues([cfg.headers]);
    hdr.setBackground('#3b4db8').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  if (!ss.getSheetByName('_Scratch')) ss.insertSheet('_Scratch').hideSheet();
  return { success: true, message: 'All sheets are ready.' };
}

function getSheet(key) {
  const s = wb().getSheetByName(SHEET_DEF[key].name);
  if (s) return s;
  initSheets();
  return wb().getSheetByName(SHEET_DEF[key].name);
}

/* ─────────────────────────── Generic CRUD Helpers ─────────────────────── */

function rows(key) {
  const sh   = getSheet(key);
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];
  const hdrs = vals[0];
  return vals.slice(1).map(r => {
    const o = {};
    hdrs.forEach((h, i) => { o[h] = r[i]; });
    return o;
  });
}

function appendRow(key, obj) {
  getSheet(key).appendRow(SHEET_DEF[key].headers.map(h => (obj[h] !== undefined && obj[h] !== null) ? obj[h] : ''));
  return obj;
}

function updateRow(key, id, patch) {
  const sh    = getSheet(key);
  const hdrs  = SHEET_DEF[key].headers;
  const idCol = hdrs.indexOf('id');
  const data  = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] !== id) continue;
    Object.entries(patch).forEach(([k, v]) => {
      const c = hdrs.indexOf(k);
      if (c >= 0) sh.getRange(i + 1, c + 1).setValue(v);
    });
    return true;
  }
  return false;
}

function removeRow(key, id) {
  const sh    = getSheet(key);
  const idCol = SHEET_DEF[key].headers.indexOf('id');
  const data  = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idCol] === id) { sh.deleteRow(i + 1); return true; }
  }
  return false;
}

function findOne(key, field, value) {
  return rows(key).find(r => r[field] === value) || null;
}

/* ─────────────────────────────── doGet ────────────────────────────────── */
// All frontend calls go through doGet to avoid the POST redirect/CORS issue.
// Write actions are sent as GET ?payload=<JSON> to stay in the simple-request
// category (no preflight, no redirect chain breakage).

function doGet(e) {
  try {
    const p = e.parameter || {};

    // ── Write path: ?payload={"action":"createAccount","name":"Me"} ──────────
    if (p.payload) {
      const d = JSON.parse(p.payload);
      return jsonResponse(dispatchWrite(d));
    }

    // ── Read path: ?action=getAll ─────────────────────────────────────────────
    const action = p.action || 'getAll';
    const readHandlers = {
      init:               () => initSheets(),
      getAll:             () => getAll(),
      getAccounts:        () => ({ success: true, data: rows('ACCOUNTS') }),
      getBrokers:         () => ({ success: true, data: rows('BROKERS') }),
      getCash:            () => ({ success: true, data: rows('CASH') }),
      getHoldings:        () => ({ success: true, data: rows('HOLDINGS') }),
      getPiggyBanks:      () => ({ success: true, data: rows('PIGGY_BANKS') }),
      getGoalAllocs:      () => ({ success: true, data: rows('GOAL_ALLOCS') }),
      getOrders:          () => ({ success: true, data: rows('ORDERS') }),
      getWatchlist:       () => ({ success: true, data: rows('WATCHLIST') }),
      getWatchlistPrices: () => watchlistPrices(),
      getPortfolioPrices: () => portfolioPrices(),
    };
    const fn = readHandlers[action];
    if (!fn) return jsonResponse({ success: false, error: 'Unknown action: ' + action });
    return jsonResponse(fn());
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// Shared write dispatcher used by both doGet(?payload) and doPost.
function dispatchWrite(d) {
  const handlers = {
    createAccount:   () => createAccount(d),
    deleteAccount:   () => deleteAccount(d),
    createBroker:    () => createBroker(d),
    deleteBroker:    () => deleteBroker(d),
    setCash:         () => setCash(d),
    addHolding:      () => addHolding(d),
    removeHolding:   () => removeHolding(d),
    createPiggyBank: () => createPiggyBank(d),
    deletePiggyBank: () => deletePiggyBank(d),
    setGoalAlloc:    () => setGoalAlloc(d),
    placeOrder:      () => placeOrder(d),
    fulfillOrder:    () => fulfillOrder(d),
    cancelOrder:     () => cancelOrder(d),
    deleteOrder:     () => deleteOrder(d),
    addWatchlist:    () => addWatchlist(d),
    removeWatchlist: () => removeWatchlist(d),
  };
  const fn = handlers[d.action];
  if (!fn) return { success: false, error: 'Unknown action: ' + d.action };
  return fn();
}

/* ─────────────────────────────── doPost ───────────────────────────────── */
// Kept for completeness / direct API testing; the browser app uses doGet+payload.

function doPost(e) {
  try {
    const d = JSON.parse((e.postData && e.postData.contents) || '{}');
    return jsonResponse(dispatchWrite(d));
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

/* ──────────────────────────── Aggregation ──────────────────────────────── */

function getAll() {
  initSheets();
  return {
    success: true,
    data: {
      accounts:   rows('ACCOUNTS'),
      brokers:    rows('BROKERS'),
      cash:       rows('CASH'),
      holdings:   rows('HOLDINGS'),
      piggyBanks: rows('PIGGY_BANKS'),
      goalAllocs: rows('GOAL_ALLOCS'),
      orders:     rows('ORDERS'),
      watchlist:  rows('WATCHLIST'),
    }
  };
}

/* ──────────────────────────── Accounts ─────────────────────────────────── */

function createAccount(d) {
  if (!d.name || !d.name.trim()) return { success: false, error: 'Account name is required' };
  return { success: true, data: appendRow('ACCOUNTS', { id: generateId(), name: d.name.trim(), created_at: nowISO() }) };
}

function deleteAccount(d) {
  if (!d.id) return { success: false, error: 'id required' };
  removeRow('ACCOUNTS', d.id);
  return { success: true };
}

/* ──────────────────────────── Brokers ──────────────────────────────────── */

function createBroker(d) {
  if (!d.account_id || !d.name || !d.name.trim()) return { success: false, error: 'account_id and name required' };
  if (!['USD', 'INR'].includes(d.currency)) return { success: false, error: 'currency must be USD or INR' };
  const rec = { id: generateId(), account_id: d.account_id, name: d.name.trim(), currency: d.currency, created_at: nowISO() };
  appendRow('BROKERS', rec);
  // Auto-create zero cash balance
  appendRow('CASH', { id: generateId(), broker_id: rec.id, balance: 0, updated_at: nowISO() });
  return { success: true, data: rec };
}

function deleteBroker(d) {
  if (!d.id) return { success: false, error: 'id required' };
  removeRow('BROKERS', d.id);
  const cr = findOne('CASH', 'broker_id', d.id);
  if (cr) removeRow('CASH', cr.id);
  return { success: true };
}

/* ──────────────────────────── Cash ─────────────────────────────────────── */

function setCash(d) {
  if (!d.broker_id) return { success: false, error: 'broker_id required' };
  const bal = parseFloat(d.balance) || 0;
  const ex  = findOne('CASH', 'broker_id', d.broker_id);
  if (ex) updateRow('CASH', ex.id, { balance: bal, updated_at: nowISO() });
  else    appendRow('CASH', { id: generateId(), broker_id: d.broker_id, balance: bal, updated_at: nowISO() });
  return { success: true };
}

function _adjustCash(brokerId, delta) {
  const ex = findOne('CASH', 'broker_id', brokerId);
  if (ex) {
    const nb = (parseFloat(ex.balance) || 0) + delta;
    updateRow('CASH', ex.id, { balance: nb, updated_at: nowISO() });
    return nb;
  }
  appendRow('CASH', { id: generateId(), broker_id: brokerId, balance: delta, updated_at: nowISO() });
  return delta;
}

/* ──────────────────────────── Holdings ─────────────────────────────────── */

function addHolding(d) {
  if (!d.broker_id || !d.symbol) return { success: false, error: 'broker_id and symbol required' };
  const sym = d.symbol.trim().toUpperCase();
  const qty = parseFloat(d.quantity) || 0;
  const avg = parseFloat(d.avg_cost) || 0;
  if (qty <= 0) return { success: false, error: 'quantity must be > 0' };

  const ex = rows('HOLDINGS').find(h => h.broker_id === d.broker_id && h.symbol === sym);
  if (ex) {
    const eqty = parseFloat(ex.quantity) || 0;
    const eavg = parseFloat(ex.avg_cost) || 0;
    const nqty = eqty + qty;
    const navg = ((eqty * eavg) + (qty * avg)) / nqty;
    updateRow('HOLDINGS', ex.id, { quantity: nqty, avg_cost: navg, updated_at: nowISO() });
    return { success: true, data: { id: ex.id, quantity: nqty, avg_cost: navg } };
  }
  const rec = { id: generateId(), broker_id: d.broker_id, symbol: sym, quantity: qty, avg_cost: avg, updated_at: nowISO() };
  appendRow('HOLDINGS', rec);
  return { success: true, data: rec };
}

function removeHolding(d) {
  if (!d.id) return { success: false, error: 'id required' };
  removeRow('HOLDINGS', d.id);
  return { success: true };
}

function _adjustHolding(brokerId, symbol, delta, buyPrice) {
  const ex = rows('HOLDINGS').find(h => h.broker_id === brokerId && h.symbol === symbol);
  if (ex) {
    const eq = parseFloat(ex.quantity) || 0;
    const ea = parseFloat(ex.avg_cost) || 0;
    const nq = eq + delta;
    if (nq < -0.00001) throw new Error(`Insufficient holdings for ${symbol}: have ${eq}, need ${Math.abs(delta)}`);
    let na = ea;
    if (delta > 0 && buyPrice != null) na = ((eq * ea) + (delta * buyPrice)) / nq;
    if (nq <= 0.00001) removeRow('HOLDINGS', ex.id);
    else updateRow('HOLDINGS', ex.id, { quantity: nq, avg_cost: na, updated_at: nowISO() });
    return { avg_cost: na, quantity: Math.max(0, nq) };
  }
  if (delta < 0) throw new Error(`No holdings record found for ${symbol}`);
  const rec = { id: generateId(), broker_id: brokerId, symbol, quantity: delta, avg_cost: buyPrice || 0, updated_at: nowISO() };
  appendRow('HOLDINGS', rec);
  return { avg_cost: buyPrice || 0, quantity: delta };
}

/* ──────────────────────────── Piggy Banks ──────────────────────────────── */

function createPiggyBank(d) {
  if (!d.name || !d.name.trim()) return { success: false, error: 'Name required' };
  const rec = {
    id:          generateId(),
    name:        d.name.trim(),
    goal_amount: parseFloat(d.goal_amount) || 0,
    currency:    d.currency || 'USD',
    description: d.description || '',
    created_at:  nowISO(),
  };
  appendRow('PIGGY_BANKS', rec);
  return { success: true, data: rec };
}

function deletePiggyBank(d) {
  if (!d.id) return { success: false, error: 'id required' };
  removeRow('PIGGY_BANKS', d.id);
  rows('GOAL_ALLOCS').filter(a => a.piggy_bank_id === d.id).forEach(a => removeRow('GOAL_ALLOCS', a.id));
  return { success: true };
}

/* ──────────────────────────── Goal Allocations ─────────────────────────── */

function setGoalAlloc(d) {
  if (!d.piggy_bank_id || !d.broker_id || !d.symbol) {
    return { success: false, error: 'piggy_bank_id, broker_id, and symbol are required' };
  }
  const sym = d.symbol.trim().toUpperCase();
  const qty = parseFloat(d.quantity) || 0;
  const ex  = rows('GOAL_ALLOCS').find(a =>
    a.piggy_bank_id === d.piggy_bank_id && a.broker_id === d.broker_id && a.symbol === sym
  );

  if (qty <= 0) {
    if (ex) removeRow('GOAL_ALLOCS', ex.id);
    return { success: true };
  }
  if (ex) {
    updateRow('GOAL_ALLOCS', ex.id, { quantity: qty, updated_at: nowISO() });
  } else {
    appendRow('GOAL_ALLOCS', { id: generateId(), piggy_bank_id: d.piggy_bank_id, broker_id: d.broker_id, symbol: sym, quantity: qty, updated_at: nowISO() });
  }
  return { success: true };
}

/* ──────────────────────────── Orders ───────────────────────────────────── */

function placeOrder(d) {
  const required = ['account_id','broker_id','symbol','order_type','exec_type','quantity','price'];
  for (const f of required) if (!d[f]) return { success: false, error: `'${f}' is required` };

  const qty   = parseFloat(d.quantity);
  const price = parseFloat(d.price);
  const fees  = parseFloat(d.fees) || 0;
  // BUY cost = (qty * price) + fees; SELL proceeds = (qty * price) - fees
  const netTotal = d.order_type === 'BUY' ? (qty * price) + fees : (qty * price) - fees;

  const hld    = rows('HOLDINGS').find(h => h.broker_id === d.broker_id && h.symbol === d.symbol.toUpperCase());
  const avgCost = hld ? parseFloat(hld.avg_cost) : price;

  const rec = {
    id:           generateId(),
    account_id:   d.account_id,
    broker_id:    d.broker_id,
    symbol:       d.symbol.trim().toUpperCase(),
    order_type:   d.order_type,
    exec_type:    d.exec_type,
    quantity:     qty,
    price:        price,
    fees:         fees,
    net_total:    Math.abs(netTotal),
    avg_cost:     avgCost,
    tax_term:     d.tax_term || '',
    status:       'Pending',
    piggy_bank_id: d.piggy_bank_id || '',
    realized_pnl: '',
    notes:        d.notes || '',
    created_at:   nowISO(),
    updated_at:   nowISO(),
  };
  appendRow('ORDERS', rec);
  return { success: true, data: rec };
}

function fulfillOrder(d) {
  const order = findOne('ORDERS', 'id', d.id);
  if (!order)                      return { success: false, error: 'Order not found' };
  if (order.status !== 'Pending')  return { success: false, error: 'Only Pending orders can be fulfilled' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const qty   = parseFloat(order.quantity);
    const price = parseFloat(order.price);
    const fees  = parseFloat(order.fees) || 0;
    const net   = parseFloat(order.net_total);

    if (order.order_type === 'BUY') {
      _adjustCash(order.broker_id, -net);
      _adjustHolding(order.broker_id, order.symbol, qty, price);

      // Auto-allocate to tagged piggy bank
      if (order.piggy_bank_id) {
        const ex = rows('GOAL_ALLOCS').find(a =>
          a.piggy_bank_id === order.piggy_bank_id &&
          a.broker_id === order.broker_id &&
          a.symbol === order.symbol
        );
        if (ex) updateRow('GOAL_ALLOCS', ex.id, { quantity: (parseFloat(ex.quantity) || 0) + qty, updated_at: nowISO() });
        else    appendRow('GOAL_ALLOCS', { id: generateId(), piggy_bank_id: order.piggy_bank_id, broker_id: order.broker_id, symbol: order.symbol, quantity: qty, updated_at: nowISO() });
      }
      updateRow('ORDERS', order.id, { status: 'Fulfilled', updated_at: nowISO() });

    } else { // SELL
      const hld = rows('HOLDINGS').find(h => h.broker_id === order.broker_id && h.symbol === order.symbol);
      if (!hld || parseFloat(hld.quantity) < qty - 0.00001) {
        return { success: false, error: `Insufficient holdings: have ${hld ? hld.quantity : 0}, selling ${qty}` };
      }
      const avgCost    = parseFloat(hld.avg_cost) || 0;
      const realizedPnl = ((price - avgCost) * qty) - fees;

      _adjustHolding(order.broker_id, order.symbol, -qty);
      _adjustCash(order.broker_id, net);

      // Deduct piggy bank allocations as directed by frontend
      (d.piggy_bank_deductions || []).forEach(ded => {
        const dq = parseFloat(ded.quantity) || 0;
        if (dq <= 0) return;
        const ex = rows('GOAL_ALLOCS').find(a =>
          a.piggy_bank_id === ded.piggy_bank_id &&
          a.broker_id === order.broker_id &&
          a.symbol === order.symbol
        );
        if (!ex) return;
        const nq = (parseFloat(ex.quantity) || 0) - dq;
        if (nq <= 0.00001) removeRow('GOAL_ALLOCS', ex.id);
        else               updateRow('GOAL_ALLOCS', ex.id, { quantity: nq, updated_at: nowISO() });
      });

      updateRow('ORDERS', order.id, {
        status:       'Fulfilled',
        realized_pnl: realizedPnl,
        avg_cost:     avgCost,
        updated_at:   nowISO(),
      });
    }
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function cancelOrder(d) {
  const order = findOne('ORDERS', 'id', d.id);
  if (!order)                     return { success: false, error: 'Order not found' };
  if (order.status !== 'Pending') return { success: false, error: 'Only Pending orders can be cancelled' };
  updateRow('ORDERS', d.id, { status: 'Cancelled', updated_at: nowISO() });
  return { success: true };
}

function deleteOrder(d) {
  if (!d.id) return { success: false, error: 'id required' };
  removeRow('ORDERS', d.id);
  return { success: true };
}

/* ──────────────────────────── Watchlist ────────────────────────────────── */

function addWatchlist(d) {
  if (!d.symbol) return { success: false, error: 'symbol required' };
  const sym = d.symbol.trim().toUpperCase();
  if (rows('WATCHLIST').find(w => w.symbol === sym)) return { success: false, error: `${sym} is already in watchlist` };
  const rec = { id: generateId(), symbol: sym, notes: d.notes || '', added_at: nowISO() };
  appendRow('WATCHLIST', rec);
  return { success: true, data: rec };
}

function removeWatchlist(d) {
  if (!d.id) return { success: false, error: 'id required' };
  removeRow('WATCHLIST', d.id);
  return { success: true };
}

/* ──────────────────────────── Price Fetching ───────────────────────────── */

function _fetchPrices(symbols) {
  if (!symbols.length) return {};
  const ss = wb();
  let sc = ss.getSheetByName('_Scratch');
  if (!sc) { sc = ss.insertSheet('_Scratch'); sc.hideSheet(); }

  symbols.forEach((sym, i) => {
    sc.getRange(i + 1, 1).setFormula(`=IFERROR(GOOGLEFINANCE("${sym}","price"),"")`);
    sc.getRange(i + 1, 2).setFormula(`=IFERROR(GOOGLEFINANCE("${sym}","changepct"),"")`);
    sc.getRange(i + 1, 3).setFormula(`=IFERROR(GOOGLEFINANCE("${sym}","name"),"")`);
    sc.getRange(i + 1, 4).setFormula(`=IFERROR(GOOGLEFINANCE("${sym}","priceopen"),"")`);
    sc.getRange(i + 1, 5).setFormula(`=IFERROR(GOOGLEFINANCE("${sym}","high"),"")`);
    sc.getRange(i + 1, 6).setFormula(`=IFERROR(GOOGLEFINANCE("${sym}","low"),"")`);
    sc.getRange(i + 1, 7).setFormula(`=IFERROR(GOOGLEFINANCE("${sym}","high52"),"")`);
    sc.getRange(i + 1, 8).setFormula(`=IFERROR(GOOGLEFINANCE("${sym}","low52"),"")`);
  });
  SpreadsheetApp.flush();
  Utilities.sleep(2500); // allow GOOGLEFINANCE to resolve

  const vals = sc.getRange(1, 1, symbols.length, 8).getValues();
  sc.getRange(1, 1, symbols.length, 8).clearContent();

  const out = {};
  symbols.forEach((s, i) => {
    out[s] = {
      price:      typeof vals[i][0] === 'number' ? vals[i][0] : null,
      change_pct: typeof vals[i][1] === 'number' ? vals[i][1] : null,
      name:       vals[i][2] || s,
      open:       typeof vals[i][3] === 'number' ? vals[i][3] : null,
      high:       typeof vals[i][4] === 'number' ? vals[i][4] : null,
      low:        typeof vals[i][5] === 'number' ? vals[i][5] : null,
      high52:     typeof vals[i][6] === 'number' ? vals[i][6] : null,
      low52:      typeof vals[i][7] === 'number' ? vals[i][7] : null,
    };
  });
  return out;
}

function watchlistPrices() {
  const syms = rows('WATCHLIST').map(w => w.symbol);
  return { success: true, data: _fetchPrices(syms) };
}

function portfolioPrices() {
  const syms = [...new Set(rows('HOLDINGS').map(h => h.symbol))];
  return { success: true, data: _fetchPrices(syms) };
}
