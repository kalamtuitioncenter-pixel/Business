/**
 * ╔══════════════════════════════════════════════════════════════╗
 * ║       MEENATCHI TRADERS — Google Apps Script v2.0           ║
 * ║       Business Management System — Complete Integration      ║
 * ╠══════════════════════════════════════════════════════════════╣
 * ║  SHEETS CREATED AUTOMATICALLY:                               ║
 * ║  DAILY SALES | Products | Customers | Orders | Tea Shop      ║
 * ║  Tuition | Staff | Suppliers | Reminders | Product Rules     ║
 * ║  Loyalty Points | Birthday Gifts | Referrals | Daily Report  ║
 * ╠══════════════════════════════════════════════════════════════╣
 * ║  SETUP (do once):                                            ║
 * ║  1. Extensions > Apps Script > paste this entire file        ║
 * ║  2. Run setupSheets() from editor to create all sheet tabs   ║
 * ║  3. Run createTrigger() once for 9 PM daily auto-report      ║
 * ║  4. Deploy > New Deployment > Web App                        ║
 * ║     Execute as: Me | Who has access: Anyone                  ║
 * ║  5. Copy Web App URL > paste in app Settings page            ║
 * ╚══════════════════════════════════════════════════════════════╝
 */

// =============================================================
//  SHEET COLUMN DEFINITIONS
//  Every sheet name maps to its exact column headers.
//  addRow functions write columns IN THIS ORDER.
// =============================================================

var SHEET_HEADERS = {

  'DAILY SALES': [
    'ID','Date','Customer','Phone','Products','Total Qty',
    'First Price','First Discount','Total','Amount Received','Pending',
    'Payment Mode','Order Status','Est Profit','Invoice','Timestamp'
  ],

  'Products': [
    'ID','Name','Category','Unit','Buy Unit','Buy Qty',
    'Buy Price','Ship Cost','Cost Per Unit','Sell Price',
    'Sticker Cost','Cover Cost','Profit Per Unit',
    'Stock','Sold','Balance','Alert Level',
    'Wholesale','Pack Size (g)','Parent Product','Timestamp'
  ],

  'Customers': [
    'ID','Name','Phone','Address','Last Product',
    'Payment Mode','Total Paid','Total Pending',
    'Loyalty Points','Birthday','Referred By','Status','Added On'
  ],

  'Orders': [
    'ID','Date','Customer','Phone','Products',
    'Qty','Amount','Status','Updated At','Timestamp'
  ],

  'Tea Shop': [
    'ID','Date','Time','Cups Sold','Sell Price Per Cup',
    'Milk Used ml','Sugar Used g','Tea Powder g',
    'Gas Expense','Other Expense',
    'Cost Per Cup','Profit Per Cup','Total Income','Daily Profit',
    'Milk Rate','Sugar Rate','Powder Rate','Notes','Timestamp'
  ],

  'Tuition': [
    'ID','Name','Class','Phone','Monthly Fee',
    'Paid','Pending','Fee Status','Added On','Timestamp'
  ],

  'Staff': [
    'ID','Name','Role','Phone','Monthly Salary',
    'Paid','Due','Last Paid Date','Added On','Timestamp'
  ],

  'Suppliers': [
    'ID','Supplier Name','Phone','Products Supplied',
    'Total Purchased','Last Order Date','Notes','Added On','Timestamp'
  ],

  'Reminders': [
    'ID','Customer','Phone','Product','Last Purchase Date',
    'Remind After Days','Due Date','Priority',
    'Notes','Message','Sent Count','Snoozed','Created At'
  ],

  'Product Rules': [
    'ID','Product','Remind After Days','Priority',
    'Message','Auto Create','Times Used','Created At'
  ],

  'Loyalty Points': [
    'Customer','Phone','Date','Points Earned',
    'Points Used','Balance','Reason','Timestamp'
  ],

  'Birthday Gifts': [
    'Customer Name','Gift Product','Qty','Message','Set On'
  ],

  'Referrals': [
    'Referrer','Referred Customer','Date','Bonus Points'
  ],

  'Daily Report': [
    'Date','Orders','Total Sales','Total Profit',
    'Tea Income','Tea Profit','Pending Payments',
    'Margin Percent','Top Product','Generated At'
  ]
};

// =============================================================
//  WEB APP ENTRY POINTS  (doGet + doPost)
// =============================================================

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || '';
    var result;

    if      (action === 'test')         result = testConnection();
    else if (action === 'setup')        result = setupSheets();
    else if (action === 'syncAll')      result = syncAllData(body.data);
    else if (action === 'addSale')      result = addSale(body.data);
    else if (action === 'addProduct')   result = saveProduct(body.data);
    else if (action === 'addCustomer')  result = saveCustomer(body.data);
    else if (action === 'addOrder')     result = saveOrder(body.data);
    else if (action === 'updateOrder')  result = updateOrderStatus(body.data);
    else if (action === 'addTeaEntry')  result = addTeaEntry(body.data);
    else if (action === 'addStudent')   result = saveStudent(body.data);
    else if (action === 'addStaff')     result = saveStaff(body.data);
    else if (action === 'addSupplier')  result = saveSupplier(body.data);
    else if (action === 'addReminder')  result = saveReminder(body.data);
    else if (action === 'addRule')      result = saveProductRule(body.data);
    else if (action === 'logPoints')    result = logLoyaltyPoints(body.data);
    else if (action === 'saveBdayGift') result = saveBirthdayGift(body.data);
    else if (action === 'addReferral')  result = addReferral(body.data);
    else if (action === 'deleteRow')    result = deleteRowById(body.sheet, body.id);
    else if (action === 'clearSheet')   result = clearSheetData(body.sheet);
    else if (action === 'getAll')       result = getAllData();
    else if (action === 'report')       result = generateDailyReport();
    else                                result = fail('Unknown action: ' + action);

    return jsonOut(result);
  } catch (ex) {
    return jsonOut(fail('doPost error: ' + ex.toString()));
  }
}

function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) ? e.parameter.action : 'test';
    var result;
    if      (action === 'test')   result = testConnection();
    else if (action === 'getAll') result = getAllData();
    else if (action === 'report') result = generateDailyReport();
    else                          result = testConnection();
    return jsonOut(result);
  } catch (ex) {
    return jsonOut(fail('doGet error: ' + ex.toString()));
  }
}

// =============================================================
//  SETUP — run once from editor
// =============================================================

function setupSheets() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var names = Object.keys(SHEET_HEADERS);
  var created = [];

  names.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      created.push(name);
    }
    if (sheet.getLastRow() === 0) {
      writeHeaders(sheet, SHEET_HEADERS[name]);
    }
  });

  // Remove blank Sheet1 if it exists
  var blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) {
    try { ss.deleteSheet(blank); } catch(e2) {}
  }

  Logger.log('Setup complete. Created: ' + created.join(', '));
  return ok('Setup complete. ' + names.length + ' sheets ready.', { created: created });
}

function writeHeaders(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var r = sheet.getRange(1, 1, 1, headers.length);
  r.setBackground('#0d1117');
  r.setFontColor('#f9c846');
  r.setFontWeight('bold');
  r.setFontSize(10);
  r.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  for (var i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 140);
  }
}

// =============================================================
//  SYNC ALL DATA (full overwrite from app export)
// =============================================================

function syncAllData(data) {
  if (!data) return fail('syncAll: no data received');
  var ts = now();
  var out = {};

  if (arr(data.sales))        out.sales        = bulkWrite('DAILY SALES',    data.sales,        saleRow);
  if (arr(data.products))     out.products      = bulkWrite('Products',       data.products,     productRow);
  if (arr(data.customers))    out.customers     = bulkWrite('Customers',      data.customers,    customerRow);
  if (arr(data.orders))       out.orders        = bulkWrite('Orders',         data.orders,       orderRow);
  if (arr(data.teaEntries))   out.teaEntries    = bulkWrite('Tea Shop',       data.teaEntries,   teaRow);
  if (arr(data.students))     out.students      = bulkWrite('Tuition',        data.students,     studentRow);
  if (arr(data.staff))        out.staff         = bulkWrite('Staff',          data.staff,        staffRow);
  if (arr(data.suppliers))    out.suppliers     = bulkWrite('Suppliers',      data.suppliers,    supplierRow);
  if (arr(data.reminders))    out.reminders     = bulkWrite('Reminders',      data.reminders,    reminderRow);
  if (arr(data.productRules)) out.productRules  = bulkWrite('Product Rules',  data.productRules, ruleRow);
  if (arr(data.bdayGifts))    out.bdayGifts     = bulkWrite('Birthday Gifts', data.bdayGifts,    giftRow);
  if (arr(data.referrals))    out.referrals     = bulkWrite('Referrals',      data.referrals,    referralRow);

  return ok('Sync complete', { syncTime: ts, results: out });
}

// =============================================================
//  ROW FORMATTERS — one function per sheet
// =============================================================

function saleRow(s) {
  var prods = '', qty = 0, price = 0, disc = 0;
  if (arr(s.products) && s.products.length > 0) {
    prods = s.products.map(function(p) { return str(p.name) + ' x' + n(p.qty); }).join(', ');
    qty   = s.products.reduce(function(a, p) { return a + n(p.qty); }, 0);
    price = n(s.products[0].price);
    disc  = n(s.products[0].disc);
  }
  return [str(s.id), str(s.date), str(s.customer), str(s.phone),
          prods, qty, price, disc,
          n(s.total), n(s.recv), n(s.pending),
          str(s.mode), str(s.status), n(s.profit),
          s.invoice ? 'Yes' : 'No', now()];
}

function productRow(p) {
  return [str(p.id), str(p.name), str(p.cat), str(p.unit), str(p.buyUnit || p.unit),
          n(p.bqty), n(p.bprice), n(p.ship), n(p.cunit), n(p.sprice),
          n(p.stickerCost), n(p.coverCost), n(p.sprice) - n(p.cunit),
          n(p.stock), n(p.sold), n(p.stock) - n(p.sold), n(p.alert),
          p.wholesale ? 'Yes' : 'No', n(p.packSize), str(p.parentProduct), now()];
}

function customerRow(c) {
  return [str(c.id), str(c.name), str(c.phone), str(c.addr),
          str(c.lastProd), str(c.mode || 'Cash'),
          n(c.paid), n(c.pending), n(c.points),
          str(c.bday), str(c.referredBy),
          n(c.pending) > 0 ? 'Pending' : 'Clear',
          todayStr()];
}

function orderRow(o) {
  return [str(o.id), str(o.date), str(o.customer), str(o.phone),
          str(o.product), n(o.qty), n(o.amount),
          str(o.status), '', now()];
}

function teaRow(e) {
  return [str(e.id || now()), str(e.date), str(e.time || ''),
          n(e.cups), n(e.sellPrice || 10),
          n(e.milk || 0), n(e.sugar || 0), n(e.powder || 0),
          n(e.gas || 0), n(e.other || 0) + n(e.expense || 0),
          n(e.cpc || 0), n(e.ppc || 0),
          n(e.income), n(e.profit),
          n(e.milkRate || 60), n(e.sugarRate || 45), n(e.powderRate || 400),
          str(e.notes), now()];
}

function studentRow(s) {
  return [str(s.id), str(s.name), str(s.cls), str(s.phone),
          n(s.fee), n(s.paid), n(s.pending),
          n(s.pending) > 0 ? 'Pending' : 'Paid',
          todayStr(), now()];
}

function staffRow(s) {
  return [str(s.id), str(s.name), str(s.role), str(s.phone),
          n(s.sal), n(s.paid), Math.max(n(s.sal) - n(s.paid), 0),
          '', todayStr(), now()];
}

function supplierRow(s) {
  return [str(s.id), str(s.name), str(s.phone), str(s.prods),
          n(s.total), str(s.lastOrder || ''), str(s.notes || ''),
          todayStr(), now()];
}

function reminderRow(r) {
  return [str(r.id), str(r.custName), str(r.phone), str(r.product),
          str(r.lastDate), n(r.days), str(r.dueDate), str(r.priority),
          str(r.notes), str(r.msg), n(r.sentCount),
          r.snoozed ? 'Yes' : 'No', str(r.createdAt)];
}

function ruleRow(r) {
  return [str(r.id), str(r.product), n(r.days), str(r.priority),
          str(r.msg), r.autoCreate ? 'Yes' : 'No',
          n(r.timesUsed), now()];
}

function giftRow(g) {
  return [str(g.custName), str(g.product), n(g.qty), str(g.msg), todayStr()];
}

function referralRow(r) {
  return [str(r.referrer), str(r.referred), str(r.date), n(r.bonus)];
}

// =============================================================
//  INDIVIDUAL SAVE FUNCTIONS (add or update by ID)
// =============================================================

function addSale(s) {
  if (!s || !s.customer) return fail('addSale: missing customer');
  var sheet = getSheet('DAILY SALES');
  sheet.appendRow(saleRow(s));
  styleRow(sheet, sheet.getLastRow());
  autoUpdateCustomer(s);
  return ok('Sale saved', { row: sheet.getLastRow() });
}

function saveProduct(p) {
  if (!p || !p.name) return fail('saveProduct: missing name');
  return upsert('Products', p, productRow);
}

function saveCustomer(c) {
  if (!c || !c.name) return fail('saveCustomer: missing name');
  return upsert('Customers', c, customerRow);
}

function saveOrder(o) {
  if (!o || !o.customer) return fail('saveOrder: missing customer');
  return upsert('Orders', o, orderRow);
}

function updateOrderStatus(data) {
  if (!data || !data.id) return fail('updateOrderStatus: missing id');
  var sheet = getSheet('Orders');
  var row   = findById(sheet, data.id);
  if (row < 2) return fail('Order not found: ' + data.id);
  sheet.getRange(row, 8).setValue(str(data.status));
  sheet.getRange(row, 9).setValue(now());
  return ok('Order status updated');
}

function addTeaEntry(e) {
  if (!e || !e.date) return fail('addTeaEntry: missing date');
  var sheet = getSheet('Tea Shop');
  sheet.appendRow(teaRow(e));
  styleRow(sheet, sheet.getLastRow());
  return ok('Tea entry saved');
}

function saveStudent(s) {
  if (!s || !s.name) return fail('saveStudent: missing name');
  return upsert('Tuition', s, studentRow);
}

function saveStaff(s) {
  if (!s || !s.name) return fail('saveStaff: missing name');
  return upsert('Staff', s, staffRow);
}

function saveSupplier(s) {
  if (!s || !s.name) return fail('saveSupplier: missing name');
  return upsert('Suppliers', s, supplierRow);
}

function saveReminder(r) {
  if (!r || !r.custName) return fail('saveReminder: missing custName');
  return upsert('Reminders', r, reminderRow);
}

function saveProductRule(r) {
  if (!r || !r.product) return fail('saveProductRule: missing product');
  return upsert('Product Rules', r, ruleRow);
}

function logLoyaltyPoints(data) {
  if (!data || !data.customer) return fail('logPoints: missing customer');
  var sheet = getSheet('Loyalty Points');
  sheet.appendRow([
    str(data.customer), str(data.phone), str(data.date),
    n(data.earned), n(data.used || 0), n(data.balance),
    str(data.reason), now()
  ]);
  styleRow(sheet, sheet.getLastRow());
  return ok('Points logged');
}

function saveBirthdayGift(g) {
  if (!g || !g.custName) return fail('saveBirthdayGift: missing custName');
  var sheet = getSheet('Birthday Gifts');
  deleteByColValue(sheet, 1, g.custName); // Remove old record for this customer
  sheet.appendRow(giftRow(g));
  styleRow(sheet, sheet.getLastRow());
  return ok('Birthday gift saved for ' + g.custName);
}

function addReferral(r) {
  if (!r || !r.referrer) return fail('addReferral: missing referrer');
  var sheet = getSheet('Referrals');
  sheet.appendRow(referralRow(r));
  styleRow(sheet, sheet.getLastRow());
  return ok('Referral saved');
}

function deleteRowById(sheetName, id) {
  if (!sheetName || id === undefined) return fail('deleteRow: sheetName and id required');
  var sheet = getSheet(sheetName);
  var row   = findById(sheet, id);
  if (row < 2) return fail('Row not found in ' + sheetName + ' for id: ' + id);
  sheet.deleteRow(row);
  return ok('Row deleted from ' + sheetName);
}

function clearSheetData(sheetName) {
  if (!sheetName) return fail('clearSheet: sheetName required');
  var sheet = getSheet(sheetName);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  return ok(sheetName + ' cleared');
}

// =============================================================
//  DAILY AUTO REPORT — triggered at 9 PM
// =============================================================

function generateDailyReport() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var tz     = Session.getScriptTimeZone();
  var today  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var totSales  = 0, totProfit = 0, orders = 0;
  var totTea    = 0, totTeaP  = 0;
  var totPend   = 0;
  var prodCount = {};

  // Sales sheet
  var sSheet = ss.getSheetByName('DAILY SALES');
  if (sSheet && sSheet.getLastRow() > 1) {
    var sd   = sSheet.getDataRange().getValues();
    var sh   = sd[0];
    var dcol = indexOf(sh, 'Date');
    var tcol = indexOf(sh, 'Total');
    var pcol = indexOf(sh, 'Est Profit');
    var rcol = indexOf(sh, 'Products');
    for (var i = 1; i < sd.length; i++) {
      if (String(sd[i][dcol]).indexOf(today) === 0) {
        totSales  += n(sd[i][tcol]);
        totProfit += n(sd[i][pcol]);
        orders++;
        String(sd[i][rcol]).split(',').forEach(function(item) {
          var nm = item.split(' x')[0].trim();
          if (nm) prodCount[nm] = (prodCount[nm] || 0) + 1;
        });
      }
    }
  }

  // Tea Shop sheet
  var tSheet = ss.getSheetByName('Tea Shop');
  if (tSheet && tSheet.getLastRow() > 1) {
    var td   = tSheet.getDataRange().getValues();
    var th   = td[0];
    var tdcol = indexOf(th, 'Date');
    var ticol = indexOf(th, 'Total Income');
    var tpcol = indexOf(th, 'Daily Profit');
    for (var j = 1; j < td.length; j++) {
      if (String(td[j][tdcol]).indexOf(today) === 0) {
        totTea  += n(td[j][ticol]);
        totTeaP += n(td[j][tpcol]);
      }
    }
  }

  // Customers pending
  var cSheet = ss.getSheetByName('Customers');
  if (cSheet && cSheet.getLastRow() > 1) {
    var cd    = cSheet.getDataRange().getValues();
    var ch    = cd[0];
    var pendC = indexOf(ch, 'Total Pending');
    for (var k = 1; k < cd.length; k++) {
      totPend += n(cd[k][pendC]);
    }
  }

  // Top product
  var topP = '', topN = 0;
  Object.keys(prodCount).forEach(function(nm) {
    if (prodCount[nm] > topN) { topN = prodCount[nm]; topP = nm; }
  });

  var margin = totSales > 0 ? ((totProfit / totSales) * 100).toFixed(1) + '%' : '0%';

  var repSheet = getSheet('Daily Report');
  repSheet.appendRow([
    today, orders, totSales, totProfit,
    totTea, totTeaP, totPend,
    margin, topP, new Date().toLocaleString()
  ]);
  styleRow(repSheet, repSheet.getLastRow());

  Logger.log('Report: ' + today + ' | Sales=' + totSales + ' | Profit=' + totProfit);
  return ok('Daily report generated', {
    date: today, orders: orders,
    totalSales: totSales, totalProfit: totProfit,
    teaIncome: totTea, pending: totPend, margin: margin
  });
}

/**
 * Run ONCE from editor to schedule 9 PM daily report.
 */
function createTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'generateDailyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('generateDailyReport').timeBased().everyDays(1).atHour(21).create();
  Logger.log('Trigger created: generateDailyReport runs at 9 PM daily.');
}

// =============================================================
//  GET ALL DATA — returns everything to app on load
// =============================================================

function getAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ok('Data loaded', {
    sales:        readSheet(ss, 'DAILY SALES'),
    products:     readSheet(ss, 'Products'),
    customers:    readSheet(ss, 'Customers'),
    orders:       readSheet(ss, 'Orders'),
    teaEntries:   readSheet(ss, 'Tea Shop'),
    students:     readSheet(ss, 'Tuition'),
    staff:        readSheet(ss, 'Staff'),
    suppliers:    readSheet(ss, 'Suppliers'),
    reminders:    readSheet(ss, 'Reminders'),
    productRules: readSheet(ss, 'Product Rules'),
    bdayGifts:    readSheet(ss, 'Birthday Gifts'),
    referrals:    readSheet(ss, 'Referrals'),
    fetchTime:    now()
  });
}

function readSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h); });
  var result  = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row.every(function(c) { return c === '' || c === null || c === undefined; })) continue;
    var obj = {};
    headers.forEach(function(h, j) { obj[h] = row[j] !== undefined ? row[j] : ''; });
    result.push(obj);
  }
  return result;
}

// =============================================================
//  INTERNAL HELPERS
// =============================================================

function getSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (SHEET_HEADERS[name]) writeHeaders(sheet, SHEET_HEADERS[name]);
  }
  return sheet;
}

function upsert(sheetName, data, rowFn) {
  var sheet = getSheet(sheetName);
  var row   = findById(sheet, data.id);
  var built = rowFn(data);
  if (row >= 2) {
    sheet.getRange(row, 1, 1, built.length).setValues([built]);
  } else {
    sheet.appendRow(built);
    styleRow(sheet, sheet.getLastRow());
  }
  return ok(sheetName + ' saved');
}

function findById(sheet, id) {
  if (sheet.getLastRow() < 2) return -1;
  var col1 = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < col1.length; i++) {
    if (String(col1[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function deleteByColValue(sheet, col, value) {
  if (sheet.getLastRow() < 2) return;
  var vals = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).toLowerCase() === String(value).toLowerCase()) {
      sheet.deleteRow(i + 2);
    }
  }
}

function bulkWrite(sheetName, dataArr, rowFn) {
  var sheet = getSheet(sheetName);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  if (!dataArr || dataArr.length === 0) return { rows: 0, sheet: sheetName };
  var rows = [];
  dataArr.forEach(function(item) {
    try { rows.push(rowFn(item)); }
    catch (ex) { Logger.log('bulkWrite row error in ' + sheetName + ': ' + ex); }
  });
  if (rows.length === 0) return { rows: 0, sheet: sheetName };
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  for (var i = 0; i < rows.length; i++) {
    var bg = i % 2 === 0 ? '#0d1117' : '#161b22';
    sheet.getRange(i + 2, 1, 1, rows[0].length)
      .setBackground(bg).setFontColor('#e6edf3').setFontSize(10);
  }
  return { rows: rows.length, sheet: sheetName };
}

function styleRow(sheet, rowNum) {
  var cols = Math.max(sheet.getLastColumn(), 1);
  var bg   = (rowNum % 2 === 0) ? '#161b22' : '#0d1117';
  sheet.getRange(rowNum, 1, 1, cols)
    .setBackground(bg).setFontColor('#e6edf3').setFontSize(10);
}

function autoUpdateCustomer(s) {
  var sheet    = getSheet('Customers');
  var custName = str(s.customer);
  if (!custName) return;
  if (sheet.getLastRow() < 2) { appendNewCustomer(sheet, s); return; }

  var nameCol = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  var rowIdx  = -1;
  for (var i = 0; i < nameCol.length; i++) {
    if (String(nameCol[i][0]).toLowerCase() === custName.toLowerCase()) {
      rowIdx = i + 2; break;
    }
  }
  var paid    = n(s.total) - n(s.pending);
  var pending = n(s.pending);
  var lastP   = arr(s.products) && s.products.length > 0 ? str(s.products[0].name) : '';

  if (rowIdx >= 2) {
    var curPaid = n(sheet.getRange(rowIdx, 7).getValue());
    var curPend = n(sheet.getRange(rowIdx, 8).getValue());
    sheet.getRange(rowIdx, 7).setValue(curPaid + paid);
    sheet.getRange(rowIdx, 8).setValue(curPend + pending);
    if (lastP) sheet.getRange(rowIdx, 5).setValue(lastP);
    sheet.getRange(rowIdx, 12).setValue((curPend + pending) > 0 ? 'Pending' : 'Clear');
  } else {
    appendNewCustomer(sheet, s);
  }
}

function appendNewCustomer(sheet, s) {
  var lastP = arr(s.products) && s.products.length > 0 ? str(s.products[0].name) : '';
  var paid  = n(s.total) - n(s.pending);
  sheet.appendRow([
    str(s.id), str(s.customer), str(s.phone), '',
    lastP, str(s.mode || 'Cash'), paid, n(s.pending), 0,
    '', '', n(s.pending) > 0 ? 'Pending' : 'Clear', todayStr()
  ]);
  styleRow(sheet, sheet.getLastRow());
}

function indexOf(arr, val) {
  for (var i = 0; i < arr.length; i++) {
    if (String(arr[i]).toLowerCase() === String(val).toLowerCase()) return i;
  }
  return -1;
}

// =============================================================
//  RESPONSE HELPERS
// =============================================================

function ok(message, data) {
  var res = { status: 'ok', message: message || 'Success' };
  if (data !== undefined) res.data = data;
  return res;
}

function fail(message) {
  return { status: 'error', message: message || 'Error' };
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================
//  TYPE HELPERS
// =============================================================

function str(v)  { return (v === null || v === undefined) ? '' : String(v); }
function n(v)    { var x = Number(v); return isNaN(x) ? 0 : x; }
function arr(v)  { return Array.isArray(v); }
function now()   { return new Date().toISOString(); }
function todayStr() { return new Date().toLocaleDateString('en-IN'); }

function testConnection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ok('Meenatchi Traders connected!', {
    spreadsheet: ss.getName(),
    timestamp:   now(),
    sheets:      ss.getSheets().map(function(s) { return s.getName(); })
  });
}

// =============================================================
//  RUN ALL TESTS — run from editor to verify everything works
// =============================================================

function runAllTests() {
  Logger.log('============ MEENATCHI TRADERS TESTS ============');

  function test(name, fn) {
    try {
      var r = fn();
      Logger.log((r && r.status === 'ok' ? '✅ PASS' : '❌ FAIL') + ' — ' + name + (r ? ': ' + r.message : ''));
    } catch(e) {
      Logger.log('❌ ERROR — ' + name + ': ' + e.toString());
    }
  }

  test('Setup sheets',         function() { return setupSheets(); });
  test('Connection test',      function() { return testConnection(); });

  test('Add sale',             function() { return addSale({
    id:'T001', date:'2026-03-15', customer:'Test Customer', phone:'9876543210',
    products:[{name:'Tea n31 100g', qty:2, price:45, disc:0, total:90}],
    total:90, recv:90, pending:0, mode:'Cash', status:'Delivered', profit:53
  }); });

  test('Save product',         function() { return saveProduct({
    id:'P001', name:'Tea n31 100g', cat:'Tea', unit:'packs', buyUnit:'kg',
    bqty:10, bprice:1800, ship:50, cunit:18.5, sprice:45,
    stickerCost:0.75, coverCost:0.40, stock:40, sold:2, alert:5,
    wholesale:false, packSize:100, parentProduct:'Tea n31'
  }); });

  test('Save customer',        function() { return saveCustomer({
    id:'C001', name:'Test Customer', phone:'9876543210',
    addr:'Chennai', lastProd:'Tea n31 100g', mode:'Cash',
    paid:90, pending:0, points:90, bday:'1990-03-15', referredBy:''
  }); });

  test('Save order',           function() { return saveOrder({
    id:1, date:'2026-03-15', customer:'Test Customer', phone:'9876543210',
    product:'Tea n31 100g', qty:2, amount:90, status:'Delivered'
  }); });

  test('Add tea entry',        function() { return addTeaEntry({
    date:'2026-03-15', time:'09:00 AM', cups:50, sellPrice:10,
    milk:2000, sugar:500, powder:100, gas:30, other:10,
    cpc:2.5, ppc:7.5, income:500, profit:375,
    milkRate:60, sugarRate:45, powderRate:400, notes:'Morning batch'
  }); });

  test('Save student',         function() { return saveStudent({
    id:'S001', name:'Test Student', cls:'10th', phone:'9123456789',
    fee:500, paid:0, pending:500
  }); });

  test('Save staff',           function() { return saveStaff({
    id:'ST001', name:'Test Teacher', role:'Teacher', phone:'9000000001',
    sal:5000, paid:0
  }); });

  test('Log loyalty points',   function() { return logLoyaltyPoints({
    customer:'Test Customer', phone:'9876543210',
    date:'2026-03-15', earned:90, used:0, balance:90,
    reason:'Sale Rs.90'
  }); });

  test('Save birthday gift',   function() { return saveBirthdayGift({
    custName:'Test Customer', product:'Tea n31 100g', qty:1,
    msg:'Happy Birthday! Enjoy your free Tea!'
  }); });

  test('Add referral',         function() { return addReferral({
    referrer:'Test Customer', referred:'New Customer',
    date:'2026-03-15', bonus:50
  }); });

  test('Save reminder',        function() { return saveReminder({
    id:'R001', custName:'Test Customer', phone:'9876543210',
    product:'Tea n31 100g', lastDate:'2026-03-15',
    days:30, dueDate:'2026-04-14', priority:'medium',
    notes:'Regular buyer', msg:'Hi Test Customer, time to reorder!',
    sentCount:0, snoozed:false, createdAt:now()
  }); });

  test('Daily report',         function() { return generateDailyReport(); });
  test('Get all data',         function() {
    var r = getAllData();
    if (r && r.status === 'ok') {
      Logger.log('   Sheets fetched: ' + Object.keys(r.data || {}).join(', '));
    }
    return r;
  });

  Logger.log('============ TESTS COMPLETE ============');
}
