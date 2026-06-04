import path from 'path';
import fs from 'fs';
import XLSX from 'xlsx';

const EXCEL_PATH = path.join(process.cwd(), 'data', 'budget.xlsx');

// Map our app's data keys to Excel sheet/cell locations
const SECTION_SHEETS = {
  committed: 'App - Budget',
  living: 'App - Budget',
  save: 'App - Budget',
  subs: 'App - Subscriptions',
  utils: 'App - Utilities',
  expenses: 'App - Expenses',
  flex: 'App - Flex',
  trips: 'App - Trips',
  goals: 'App - Goals',
  cards: 'App - Cards',
};

function ensureAppSheets(wb) {
  const needed = [...new Set(Object.values(SECTION_SHEETS))];
  needed.forEach(name => {
    if (!wb.Sheets[name]) {
      wb.Sheets[name] = {};
      wb.SheetNames.push(name);
    }
  });
}

function sheetToJson(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: null });
}

function jsonToSheet(wb, sheetName, rows) {
  const ws = XLSX.utils.json_to_sheet(rows);
  wb.Sheets[sheetName] = ws;
}

function readAllData(wb) {
  ensureAppSheets(wb);

  const budget = { committed: [], living: [], save: [] };
  const budgetRows = sheetToJson(wb, 'App - Budget');
  budgetRows.forEach(r => {
    if (r.section && budget[r.section]) {
      budget[r.section].push({
        emoji: r.emoji || '📌',
        name: r.name || '',
        amount: parseFloat(r.amount) || 0,
        recurring: r.recurring === true || r.recurring === 'TRUE' || r.recurring === 1,
        locked: r.locked === true || r.locked === 'TRUE' || r.locked === 1,
        ...(r.feed ? { feed: r.feed } : {}),
        ...(r.cat ? { cat: r.cat } : {}),
      });
    }
  });

  const subs = sheetToJson(wb, 'App - Subscriptions').map(r => ({
    name: r.name || '',
    amount: parseFloat(r.amount) || 0,
    cycle: r.cycle || 'monthly',
    ess: r.ess || 'essential',
  }));

  const utils = sheetToJson(wb, 'App - Utilities').map(r => ({
    name: r.name || '',
    amount: parseFloat(r.amount) || 0,
    split: r.split || '',
  }));

  // Expenses keyed by month
  const expenses = {};
  sheetToJson(wb, 'App - Expenses').forEach(r => {
    const mk = r.monthKey;
    if (!mk) return;
    if (!expenses[mk]) expenses[mk] = [];
    expenses[mk].push({
      desc: r.desc || '',
      amount: parseFloat(r.amount) || 0,
      cat: r.cat || 'Misc',
      card: r.card || 'Cash',
      date: r.date || '',
    });
  });

  // Flex keyed by month
  const flex = {};
  sheetToJson(wb, 'App - Flex').forEach(r => {
    const mk = r.monthKey;
    if (!mk) return;
    if (!flex[mk]) flex[mk] = [];
    flex[mk].push({
      desc: r.desc || '',
      amount: parseFloat(r.amount) || 0,
      cat: r.cat || 'Other',
      date: r.date || '',
    });
  });

  // Trips
  const tripsRaw = sheetToJson(wb, 'App - Trips');
  const tripsMap = {};
  tripsRaw.forEach(r => {
    if (!r.id) return;
    if (!tripsMap[r.id]) {
      tripsMap[r.id] = { name: r.name, when: r.when, budget: parseFloat(r.budget) || 0, saved: parseFloat(r.saved) || 0, expenses: [] };
    }
    if (r.expDesc) {
      tripsMap[r.id].expenses.push({ desc: r.expDesc, amt: parseFloat(r.expAmt) || 0, cat: r.expCat || 'Misc', month: r.expMonth || '' });
    }
  });
  const trips = Object.values(tripsMap);

  const goals = sheetToJson(wb, 'App - Goals').map(r => ({
    name: r.name || '',
    target: parseFloat(r.target) || 0,
    saved: parseFloat(r.saved) || 0,
    monthly: parseFloat(r.monthly) || 0,
  }));

  const cards = sheetToJson(wb, 'App - Cards').map(r => ({
    name: r.name || '',
    issuer: r.issuer || '',
    color: r.color || '#1a1916',
  }));

  const income = parseFloat(sheetToJson(wb, 'App - Budget').find(r => r.section === '__income')?.amount) || 5222.38;
  const loan = parseFloat(sheetToJson(wb, 'App - Budget').find(r => r.section === '__loan')?.amount) || 38550;

  return { budget, subs, utils, expenses, flex, trips, goals, cards, income, loan };
}

function writeAllData(wb, data) {
  ensureAppSheets(wb);

  // Budget sheet (committed + living + save + income + loan meta rows)
  const budgetRows = [];
  budgetRows.push({ section: '__income', amount: data.income });
  budgetRows.push({ section: '__loan', amount: data.loan });
  ['committed', 'living', 'save'].forEach(sec => {
    (data.budget[sec] || []).forEach(it => {
      budgetRows.push({
        section: sec,
        emoji: it.emoji,
        name: it.name,
        amount: it.amount,
        recurring: it.recurring ? 1 : 0,
        locked: it.locked ? 1 : 0,
        feed: it.feed || '',
        cat: it.cat || '',
      });
    });
  });
  jsonToSheet(wb, 'App - Budget', budgetRows);

  jsonToSheet(wb, 'App - Subscriptions', (data.subs || []).map(s => ({
    name: s.name, amount: s.amount, cycle: s.cycle, ess: s.ess
  })));

  jsonToSheet(wb, 'App - Utilities', (data.utils || []).map(u => ({
    name: u.name, amount: u.amount, split: u.split
  })));

  // Flatten expenses (keyed by month)
  const expRows = [];
  Object.entries(data.expenses || {}).forEach(([mk, items]) => {
    (items || []).forEach(it => expRows.push({ monthKey: mk, ...it }));
  });
  jsonToSheet(wb, 'App - Expenses', expRows);

  // Flatten flex
  const flexRows = [];
  Object.entries(data.flex || {}).forEach(([mk, items]) => {
    (items || []).forEach(it => flexRows.push({ monthKey: mk, ...it }));
  });
  jsonToSheet(wb, 'App - Flex', flexRows);

  // Trips — flatten trips+expenses with a shared id
  const tripRows = [];
  (data.trips || []).forEach((t, i) => {
    const id = i;
    if (!t.expenses || t.expenses.length === 0) {
      tripRows.push({ id, name: t.name, when: t.when, budget: t.budget, saved: t.saved });
    } else {
      t.expenses.forEach(e => {
        tripRows.push({ id, name: t.name, when: t.when, budget: t.budget, saved: t.saved, expDesc: e.desc, expAmt: e.amt, expCat: e.cat, expMonth: e.month });
      });
    }
  });
  jsonToSheet(wb, 'App - Trips', tripRows);

  jsonToSheet(wb, 'App - Goals', (data.goals || []).map(g => ({
    name: g.name, target: g.target, saved: g.saved, monthly: g.monthly
  })));

  jsonToSheet(wb, 'App - Cards', (data.cards || []).map(c => ({
    name: c.name, issuer: c.issuer, color: c.color
  })));
}

export default async function handler(req, res) {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      return res.status(404).json({ error: 'budget.xlsx not found in /data folder' });
    }

    const wb = XLSX.readFile(EXCEL_PATH);

    if (req.method === 'GET') {
      const data = readAllData(wb);
      // If App sheets are empty (first run), seed with defaults
      const isEmpty = data.budget.committed.length === 0 && data.budget.save.length === 0;
      if (isEmpty) {
        return res.status(200).json({ data: null, firstRun: true });
      }
      return res.status(200).json({ data });
    }

    if (req.method === 'POST') {
      const { data } = req.body;
      writeAllData(wb, data);
      XLSX.writeFile(wb, EXCEL_PATH);
      return res.status(200).json({ ok: true });
    }

    res.status(405).json({ error: 'Method not allowed' });
  } catch (err) {
    console.error('Excel API error:', err);
    res.status(500).json({ error: err.message });
  }
}
