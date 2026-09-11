/**
 * Expense Tracker — Apps Script backend
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me
 *   Who has access: Anyone
 */

// https://docs.google.com/spreadsheets/d/1zfPcBORv3b9nzwHU-a6E5X2yflDAhRmgjf7ZPhjZHHA/edit
const SHEET_NAME   = 'Expense Tracker';
const HEADERS      = ['Date', 'Amount', 'Category', 'Note'];
const INCOME_SHEET = 'Income';
const INCOME_HDR   = ['Date', 'Amount', 'Source', 'Note'];

// ── Expense sheet ──────────────────────────────────────────────────────────

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  ensureFormatted_(sheet);
  return sheet;
}

function ensureFormatted_(sheet) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('formatted') === 'true') return;

  const GREEN = '#1E6B3C';
  sheet.setTabColor(GREEN);

  const hdr = sheet.getRange('A1:D1');
  hdr.setBackground(GREEN).setFontColor('#FFFFFF').setFontWeight('bold')
     .setFontFamily('Arial').setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('F1').setBackground(GREEN).setFontColor('#FFFFFF')
       .setFontWeight('bold').setFontFamily('Arial').setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 170); sheet.setColumnWidth(4, 240);
  sheet.setColumnWidth(5, 20);  sheet.setColumnWidth(6, 170);

  sheet.getRange('B2:B2000').setNumberFormat('#,##0.00" €"');
  sheet.getRange('A2:A2000').setNumberFormat('dd-mmm-yyyy').setFontFamily('Arial');
  sheet.getRange('C2:D2000').setFontFamily('Arial');
  sheet.getRange('F2:F200').setFontFamily('Arial');

  const catRange = sheet.getRange('F2:F200');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(catRange, true).setAllowInvalid(false).build();
  sheet.getRange('C2:C2000').setDataValidation(rule);

  sheet.getRange('A1:F1').setBorder(false, false, true, false, false, false,
    '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  props.setProperty('formatted', 'true');
}

function getCategories_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
  return values.map(r => r[0]).filter(v => v !== '');
}

// ── Income sheet ───────────────────────────────────────────────────────────

function getIncomeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INCOME_SHEET);
  if (!sheet) sheet = ss.insertSheet(INCOME_SHEET);
  if (sheet.getLastRow() === 0) sheet.appendRow(INCOME_HDR);
  ensureIncomeFormatted_(sheet);
  return sheet;
}

function ensureIncomeFormatted_(sheet) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('income_formatted') === 'true') return;

  const BLUE = '#1A5276';
  sheet.setTabColor(BLUE);

  const hdr = sheet.getRange('A1:D1');
  hdr.setBackground(BLUE).setFontColor('#FFFFFF').setFontWeight('bold')
     .setFontFamily('Arial').setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 170); sheet.setColumnWidth(4, 240);

  sheet.getRange('B2:B2000').setNumberFormat('#,##0.00" €"');
  sheet.getRange('A2:A2000').setNumberFormat('dd-mmm-yyyy').setFontFamily('Arial');
  sheet.getRange('C2:D2000').setFontFamily('Arial');

  sheet.getRange('A1:D1').setBorder(false, false, true, false, false, false,
    '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  props.setProperty('income_formatted', 'true');
}

function addIncome_(sheet, params) {
  const amount = parseFloat(params.amount);
  if (isNaN(amount)) return { status: 'error', message: 'amount is required and must be a number' };
  const date   = params.date ? new Date(params.date) : new Date();
  const source = params.source || 'Salary';
  const note   = params.note || '';
  sheet.appendRow([date, amount, source, note]);
  return { status: 'ok', message: 'Income added' };
}

function listIncome_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'ok', data: [] };
  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const rows = values
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      date:   (r[0] instanceof Date) ? r[0].toISOString() : String(r[0]),
      amount: Number(r[1]) || 0,
      source: r[2] || 'Salary',
      note:   r[3] || ''
    }));
  return { status: 'ok', data: rows };
}

// ── Router ─────────────────────────────────────────────────────────────────

function doGet(e) {
  const params = e.parameter || {};
  const action = params.action || 'list';

  let result;
  try {
    if (action === 'add') {
      result = addExpense_(getSheet_(), params);
    } else if (action === 'list') {
      result = listExpenses_(getSheet_());
    } else if (action === 'categories') {
      result = { status: 'ok', categories: getCategories_(getSheet_()) };
    } else if (action === 'add_income') {
      result = addIncome_(getIncomeSheet_(), params);
    } else if (action === 'list_income') {
      result = listIncome_(getIncomeSheet_());
    } else {
      result = { status: 'error', message: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { status: 'error', message: err.message };
  }

  const json = JSON.stringify(result);
  if (params.callback) {
    return ContentService
      .createTextOutput(params.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ── Helpers ────────────────────────────────────────────────────────────────

function addExpense_(sheet, params) {
  const amount = parseFloat(params.amount);
  if (isNaN(amount)) return { status: 'error', message: 'amount is required and must be a number' };
  const date     = params.date ? new Date(params.date) : new Date();
  const category = params.category || 'Uncategorized';
  const note     = params.note || '';
  sheet.appendRow([date, amount, category, note]);
  return { status: 'ok', message: 'Expense added' };
}

function listExpenses_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'ok', data: [] };
  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const rows = values
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      date:     (r[0] instanceof Date) ? r[0].toISOString() : String(r[0]),
      amount:   Number(r[1]) || 0,
      category: r[2] || 'Uncategorized',
      note:     r[3] || ''
    }));
  return { status: 'ok', data: rows };
}

function setup() { getSheet_(); getIncomeSheet_(); }
