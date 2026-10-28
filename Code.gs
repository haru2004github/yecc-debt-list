const SHEET_DEBTS = 'Debts';
const SHEET_PAYMENTS = 'Payments';

function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Canonical API names; existing worksheet headers are preserved.
function canonicalHeader(header, sheetName) {
  const key = String(header).trim().toLowerCase().replace(/\s+/g, '_');
  if (sheetName === SHEET_DEBTS && key === 'debt_id') return 'id';
  if (sheetName === SHEET_PAYMENTS) {
    if (key === 'payment_id' || key === 'pyament_id') return 'id';
    if (key === 'payment_date') return 'date';
    if (key === 'amount_paid') return 'amount';
  }
  return key;
}

function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemas = [
    [SHEET_DEBTS, ['debt_id', 'date', 'name', 'year', 'total_debt', 'paid_amount', 'balance', 'status', 'category']],
    [SHEET_PAYMENTS, ['payment_id', 'debt_id', 'payment_date', 'amount_paid']]
  ];
  schemas.forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  });
}

function readTable(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => canonicalHeader(h, name));
  const required = name === SHEET_DEBTS
    ? ['id', 'date', 'name', 'year', 'total_debt', 'paid_amount', 'balance', 'status', 'category']
    : ['id', 'debt_id', 'date', 'amount'];
  required.forEach(key => {
    if (headers.indexOf(key) < 0) throw new Error(name + ': missing column ' + key);
    if (headers.indexOf(key) !== headers.lastIndexOf(key)) throw new Error(name + ': duplicate column ' + key);
  });
  return { sheet, headers, rows: values.slice(1) };
}

function getSheetData(name) {
  const table = readTable(name);
  return table.rows.filter(row => row.some(v => v !== '')).map(row => {
    const obj = {};
    table.headers.forEach((key, i) => { if (key) obj[key] = row[i]; });
    if (name === SHEET_DEBTS) {
      obj.total_debt = Number(obj.total_debt) || 0;
      obj.paid_amount = Number(obj.paid_amount) || 0;
      obj.balance = obj.balance === '' ? obj.total_debt - obj.paid_amount : Number(obj.balance) || 0;
    } else {
      obj.amount = Number(obj.amount) || 0;
    }
    return obj;
  });
}

function appendObject(table, obj) {
  table.sheet.appendRow(table.headers.map(key => obj[key] === undefined ? '' : obj[key]));
}

function positiveAmount(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount) || amount <= 0) throw new Error('Amount must be greater than zero');
  return amount;
}

function doGet() {
  try {
    initSheets();
    return createResponse({ status: 'success', debts: getSheetData(SHEET_DEBTS), payments: getSheetData(SHEET_PAYMENTS) });
  } catch (error) {
    return createResponse({ status: 'error', message: String(error) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createResponse({ status: 'error', message: 'System is busy. Please try again.' });
  try {
    initSheets();
    const request = JSON.parse(e.postData.contents);
    const data = request.data || {};
    const debts = readTable(SHEET_DEBTS);
    const payments = readTable(SHEET_PAYMENTS);
    const idColumn = debts.headers.indexOf('id');
    if (request.action === 'addDebt') {
      const amount = positiveAmount(data.total_debt);
      const name = String(data.name || '').trim();
      if (!name) throw new Error('Name is required');
      const id = String(data.debt_id || data.id || ('D-' + Utilities.getUuid()));
      if (debts.rows.some(row => String(row[idColumn]) === id)) throw new Error('Debt ID already exists');
      appendObject(debts, {
        id, user_id: data.user_id || '', date: data.date || new Date().toISOString(),
        name, year: data.year || 'N/A', total_debt: amount, paid_amount: 0,
        balance: amount, status: 'Unpaid', category: data.category || 'Canteen'
      });
    } else if (request.action === 'addPayment' || request.action === 'increaseDebt') {
      const amount = positiveAmount(data.amount);
      const matches = debts.rows.map((row, index) => ({ row, index }))
        .filter(item => String(item.row[idColumn]) === String(data.debt_id));
      if (matches.length !== 1) throw new Error('Debt ID is missing or not unique');
      const { row, index } = matches[0];
      const totalCol = debts.headers.indexOf('total_debt');
      const paidCol = debts.headers.indexOf('paid_amount');
      const balanceCol = debts.headers.indexOf('balance');
      const statusCol = debts.headers.indexOf('status');
      const balance = row[balanceCol] === '' ? Number(row[totalCol]) - Number(row[paidCol]) : Number(row[balanceCol]);
      if (!Number.isFinite(balance)) throw new Error('Invalid debt balance');
      const increase = request.action === 'increaseDebt';
      if (!increase && amount > balance) throw new Error('Payment exceeds the remaining balance');
      const id = String(data.id || ((increase ? 'INC-' : 'P-') + Utilities.getUuid()));
      if (payments.rows.some(r => String(r[payments.headers.indexOf('id')]) === id)) throw new Error('Transaction ID already exists');
      if (increase) row[totalCol] = (Number(row[totalCol]) || 0) + amount;
      else row[paidCol] = (Number(row[paidCol]) || 0) + amount;
      row[balanceCol] = balance + (increase ? amount : -amount);
      row[statusCol] = row[balanceCol] <= 0 ? 'Paid' : 'Unpaid';
      // Write only accounting fields, preserving unrelated cells and formulas.
      [totalCol, paidCol, balanceCol, statusCol].forEach(col => {
        debts.sheet.getRange(index + 2, col + 1).setValue(row[col]);
      });
      appendObject(payments, { id, debt_id: data.debt_id, date: data.date || new Date().toISOString(), amount: increase ? -amount : amount });
    } else if (request.action === 'updateDebtDetails') {
      updateDebtDetails(debts, data);
    } else if (['updateTransaction', 'deleteTransaction'].includes(request.action)) {
      mutateTransaction(debts, payments, data, request.action === 'deleteTransaction');
    } else {
      throw new Error('Invalid action');
    }
    return createResponse({ status: 'success', message: 'Saved' });
  } catch (error) {
    return createResponse({ status: 'error', message: String(error) });
  } finally {
    lock.releaseLock();
  }
}

// Edits preserve historical opening balances by applying only the transaction delta.
function mutateTransaction(debts, payments, data, deleting) {
  const initial = data.kind === 'initial';
  if (!initial && data.kind !== 'payment') throw new Error('Invalid transaction kind');
  const table = initial ? debts : payments;
  const matches = table.rows.map((row, index) => ({ row, index }))
    .filter(item => String(item.row[table.headers.indexOf('id')]) === String(data.id));
  if (!data.id || matches.length !== 1) throw new Error('Transaction is missing or not unique');
  const item = matches[0];
  const debtId = initial ? data.id : item.row[payments.headers.indexOf('debt_id')];
  const parents = debts.rows.map((row, index) => ({ row, index }))
    .filter(item => String(item.row[debts.headers.indexOf('id')]) === String(debtId));
  if (parents.length !== 1) throw new Error('Debt is missing or not unique');
  const parent = parents[0];
  const related = payments.rows.filter(row => String(row[payments.headers.indexOf('debt_id')]) === String(debtId));
  const increases = related.reduce((sum, row) => sum + Math.max(0, -Number(row[payments.headers.indexOf('amount')])), 0);
  const total = Number(parent.row[debts.headers.indexOf('total_debt')]);
  const paid = Number(parent.row[debts.headers.indexOf('paid_amount')]);
  const balance = Number(parent.row[debts.headers.indexOf('balance')]);
  const oldAmount = initial ? total - increases : Number(item.row[payments.headers.indexOf('amount')]);
  const oldDate = new Date(item.row[table.headers.indexOf('date')]).toISOString();
  if (Number(data.expected_amount) !== oldAmount || data.expected_date !== oldDate)
    throw new Error('Transaction changed. Refresh before editing again.');
  if (initial && deleting) {
    if (related.length || paid !== 0) throw new Error('Delete linked payments and increases before deleting this debt.');
    debts.sheet.deleteRow(parent.index + 2);
    return;
  }
  const date = deleting ? oldDate : new Date(data.date).toISOString();
  const amount = deleting ? 0 : positiveAmount(data.amount) * (oldAmount < 0 ? -1 : 1);
  const delta = amount - oldAmount;
  const newTotal = total + (initial ? delta : oldAmount < 0 ? -delta : 0);
  const newPaid = paid + (!initial && oldAmount >= 0 ? delta : 0);
  const newBalance = balance + (initial ? delta : -delta);
  if (![newTotal, newPaid, newBalance].every(Number.isFinite) || newTotal < 0 || newPaid < 0 || newBalance < 0)
    throw new Error('This change would create a negative total, paid amount or balance. Correct related transactions first.');
  const updates = {total_debt: newTotal, paid_amount: newPaid, balance: newBalance, status: newBalance <= 0 ? 'Paid' : 'Unpaid'};
  // Validate everything before writing. Keep unrelated sheet columns intact.
  Object.keys(updates).forEach(key => debts.sheet.getRange(parent.index + 2, debts.headers.indexOf(key) + 1).setValue(updates[key]));
  if (deleting) payments.sheet.deleteRow(item.index + 2);
  else {
    table.sheet.getRange(item.index + 2, table.headers.indexOf('date') + 1).setValue(date);
    if (!initial) payments.sheet.getRange(item.index + 2, payments.headers.indexOf('amount') + 1).setValue(amount);
  }
}

function updateDebtDetails(debts, data) {
  const matches = debts.rows.map((row, index) => ({row, index}))
    .filter(item => String(item.row[debts.headers.indexOf('id')]) === String(data.id));
  if (!data.id || matches.length !== 1) throw new Error('Debt is missing or not unique');
  const {row, index} = matches[0];
  const fields = ['name', 'year', 'category'];
  const values = {};
  fields.forEach(key => {
    if (!data.expected || String(data.expected[key] || '') !== String(row[debts.headers.indexOf(key)] || ''))
      throw new Error('Student details changed. Refresh and try again.');
    values[key] = String(data[key] || '').trim();
    if (!values[key] || values[key].length > 150 || /^[=+@-]/.test(values[key]))
      throw new Error('Enter valid ' + key);
  });
  if (!['Canteen', 'Kitchen'].includes(values.category)) throw new Error('Invalid category');
  fields.forEach(key => debts.sheet.getRange(index + 2, debts.headers.indexOf(key) + 1).setValue(values[key]));
}
