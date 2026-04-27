const crypto = require('crypto');

const SPREADSHEET_ID = process.env.GOOGLE_SHEETS_SPREADSHEET_ID;
const cors = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

// --- Google service account JWT auth (no npm deps) ---
async function getAccessToken() {
  const sa = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const now = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify({
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now,
  })).toString('base64url');
  const unsigned = `${header}.${payload}`;
  const sign = crypto.createSign('RSA-SHA256');
  sign.update(unsigned);
  const sig = sign.sign(sa.private_key, 'base64url');
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${unsigned}.${sig}`,
  });
  const { access_token } = await res.json();
  return access_token;
}

// --- Sheets API helpers ---
async function sheetsGet(token, range) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  return res.json();
}

async function sheetsUpdate(token, range, values) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;
  const res = await fetch(url, {
    method: 'PUT',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values }),
  });
  return res.json();
}

async function sheetsClear(token, range) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}:clear`;
  await fetch(url, { method: 'POST', headers: { Authorization: `Bearer ${token}` } });
}

async function sheetsAppend(token, range, values) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values }),
  });
  return res.json();
}

// --- Data transformation ---
const INV_HEADERS = ['id','name','cat','scope','qty','location','container','notes','lastBy','lastDate','retired'];

function rowsToInventory(rows) {
  if (!rows || rows.length < 2) return null;
  const [, ...data] = rows; // skip header row
  return data.map(row => {
    const obj = {};
    INV_HEADERS.forEach((h, i) => obj[h] = row[i] ?? '');
    obj.id = parseInt(obj.id) || 0;
    obj.qty = parseInt(obj.qty) || 0;
    obj.retired = obj.retired === 'true';
    return obj;
  });
}

function inventoryToRows(inventory) {
  const rows = [INV_HEADERS];
  inventory.forEach(i => rows.push(INV_HEADERS.map(h => {
    if (h === 'retired') return i.retired ? 'true' : 'false';
    return (i[h] !== undefined && i[h] !== null) ? String(i[h]) : '';
  })));
  return rows;
}

const PL_HEADERS = ['type', 'label', 'qty', 'price', 'amount'];

function rowsToPLPlan(rows) {
  if (!rows || rows.length < 2) return null;
  const rev = {}, exp = {};
  const [, ...data] = rows;
  data.forEach(row => {
    const [type, label, qty, price, amount] = row;
    if (!label) return;
    if (type === 'rev') rev[label] = { q: parseFloat(qty) || 0, p: parseFloat(price) || 0 };
    else if (type === 'exp') exp[label] = { a: parseFloat(amount) || 0 };
  });
  return (Object.keys(rev).length || Object.keys(exp).length) ? { rev, exp } : null;
}

function plPlanToRows(pl) {
  const rows = [PL_HEADERS];
  Object.entries(pl.rev).forEach(([label, r]) =>
    rows.push(['rev', label, String(r.q), String(r.p), String(r.q * r.p)]));
  Object.entries(pl.exp).forEach(([label, e]) =>
    rows.push(['exp', label, '', '', String(e.a)]));
  return rows;
}

// --- Handler ---
exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 200, headers: cors, body: '' };

  try {
    const token = await getAccessToken();
    const body = JSON.parse(event.body || '{}');
    const { action } = body;

    // On page load — read everything at once
    if (action === 'readAll') {
      const [invData, plData, actualsData] = await Promise.all([
        sheetsGet(token, 'Inventory!A:K'),
        sheetsGet(token, 'PL_Plan!A:E'),
        sheetsGet(token, 'PL_Actuals!A:D'),
      ]);
      const inventory = rowsToInventory(invData.values);
      const plPlan = rowsToPLPlan(plData.values);
      let plActuals = null;
      if (actualsData.values && actualsData.values.length >= 2) {
        const row = actualsData.values[1];
        plActuals = { r: parseFloat(row[0]) || 0, e: parseFloat(row[1]) || 0 };
      }
      return {
        statusCode: 200,
        headers: { ...cors, 'Content-Type': 'application/json' },
        body: JSON.stringify({ inventory, plPlan, plActuals }),
      };
    }

    // Append a row to the inventory move log
    if (action === 'appendInventoryLog') {
      const { item, from, to, who, ts } = body;
      await sheetsAppend(token, 'InventoryLog!A:E', [[ts, item, from, to, who]]);
      return { statusCode: 200, headers: cors, body: JSON.stringify({ ok: true }) };
    }

    // Overwrite full inventory state (called after any item update)
    if (action === 'writeInventory') {
      const rows = inventoryToRows(body.inventory);
      await sheetsClear(token, 'Inventory!A:K');
      await sheetsUpdate(token, 'Inventory!A1', rows);
      return { statusCode: 200, headers: cors, body: JSON.stringify({ ok: true }) };
    }

    // Overwrite P&L plan (debounced, called on any price/qty edit)
    if (action === 'writePLPlan') {
      const rows = plPlanToRows(body.pl);
      await sheetsClear(token, 'PL_Plan!A:E');
      await sheetsUpdate(token, 'PL_Plan!A1', rows);
      return { statusCode: 200, headers: cors, body: JSON.stringify({ ok: true }) };
    }

    // Overwrite P&L actuals (called when actuals inputs change)
    if (action === 'writePLActuals') {
      const { actuals, who } = body;
      await sheetsClear(token, 'PL_Actuals!A:D');
      await sheetsUpdate(token, 'PL_Actuals!A1', [
        ['actualRevenue', 'actualExpenses', 'updatedBy', 'updatedAt'],
        [String(actuals.r), String(actuals.e), who || '', new Date().toISOString()],
      ]);
      return { statusCode: 200, headers: cors, body: JSON.stringify({ ok: true }) };
    }

    return { statusCode: 400, headers: cors, body: JSON.stringify({ error: 'Unknown action' }) };
  } catch (err) {
    console.error('sheets.js error:', err);
    return {
      statusCode: 500,
      headers: { ...cors, 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: err.message }),
    };
  }
};
