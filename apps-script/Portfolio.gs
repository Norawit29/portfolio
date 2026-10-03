/**
 * Portfolio quick entry for the "Portfolio" Google Sheet.
 *
 * Adds a "📈 Portfolio" menu with:
 *   - a form (sidebar, or a web app on the phone) that appends rows to the
 *     Transaction sheet and copies the calculation formulas (H:P) for you
 *   - a one-shot DCA entry for every asset marked "dca" in Portfolio Summary
 *   - a Monthly Summary snapshot (manual or on a monthly timer)
 *   - an audit of existing transactions (future dates, prices that look swapped)
 *
 * The sheet layout is not changed: rows are written exactly the way they are
 * typed in by hand today, so Cost / Portfolio Summary keep working.
 */

var PF = {
  TX: 'Transaction',
  PRICE: 'Price',
  SUMMARY: 'Portfolio Summary',
  MONTHLY: 'Monthly Summary',
  TX_WIDTH: 17,            // A:Q
  FIRST_TEMPLATE_ROW: 3,   // row 2 has #REF! in H/K, never copy formulas from it
  PRICE_WARN: 0.3,         // warn when a price is >30% away from the reference
  OLD_DATE_WARN_DAYS: 45,
  SNAPSHOT_DAY: 20,        // day of month for the automatic snapshot
  SNAPSHOT_HOUR: 20,
  // Monthly Summary label -> Portfolio Summary label, after pfNorm_()
  LABEL_ALIASES: { 'cash(usd)': 'cashininnovestx(usd)' },
};

// 1-based column numbers in Transaction
var TXC = { DATE: 1, TYPE: 2, SYMBOL: 3, UNITS: 4, PRICE: 5, FEE: 6, SPLIT: 7,
            CUM_UNITS: 9, CUM_COST: 14, FX: 17 };

/* ------------------------------------------------------------------ menu -- */

function pfOnOpen() {
  try {
    PropertiesService.getScriptProperties()
      .setProperty('PF_SPREADSHEET_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  } catch (err) { /* simple trigger without permission; the menu still works */ }
  SpreadsheetApp.getUi().createMenu('📈 Portfolio')
    .addItem('➕ บันทึกซื้อ / ขาย', 'pfShowEntry')
    .addItem('📅 บันทึก DCA', 'pfShowDca')
    .addSeparator()
    .addItem('📸 Snapshot ลง Monthly Summary ตอนนี้', 'pfSnapshotNow')
    .addItem('⏰ ตั้ง Snapshot อัตโนมัติทุกเดือน', 'pfInstallMonthlyTrigger')
    .addSeparator()
    .addItem('🔍 ตรวจข้อมูล Transaction', 'pfAudit')
    .addToUi();
}

function pfShowEntry() { pfShowSidebar_('single'); }
function pfShowDca() { pfShowSidebar_('dca'); }

function pfShowSidebar_(tab) {
  PropertiesService.getScriptProperties()
    .setProperty('PF_SPREADSHEET_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  var html = pfPage_(tab).setTitle('บันทึกรายการ');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Web app entry: same form, full screen, for use on the phone. */
function pfDoGet(e) {
  var tab = (e && e.parameter && e.parameter.tab) === 'dca' ? 'dca' : 'single';
  return pfPage_(tab)
    .setTitle('Portfolio – บันทึกรายการ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function pfPage_(tab) {
  var t = HtmlService.createTemplateFromFile('QuickAdd');
  t.initialTab = tab;
  return t.evaluate();
}

/* ------------------------------------------------------------- reading -- */

function pfSs_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('PF_SPREADSHEET_ID');
  if (!id) throw new Error('เปิดไฟล์ Portfolio ในเบราว์เซอร์หนึ่งครั้งก่อน เพื่อให้สคริปต์จำไฟล์ได้');
  return SpreadsheetApp.openById(id);
}

function pfSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต "' + name + '"');
  return sh;
}

function pfNum_(v) { return typeof v === 'number' && isFinite(v) ? v : null; }

function pfLastTxRow_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var syms = sh.getRange(2, TXC.SYMBOL, last - 1, 1).getValues();
  for (var i = syms.length - 1; i >= 0; i--) {
    if (String(syms[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

/** Transactions with computed values, plus the running state per symbol. */
function pfReadTx_(ss) {
  var sh = pfSheet_(ss, PF.TX);
  var last = pfLastTxRow_(sh);
  var rows = last >= 2 ? sh.getRange(2, 1, last - 1, PF.TX_WIDTH).getValues() : [];
  var bySymbol = {};
  var list = [];
  rows.forEach(function (r, i) {
    var sym = String(r[TXC.SYMBOL - 1]).trim();
    if (!sym) return;
    var tx = {
      row: i + 2,
      date: r[TXC.DATE - 1] instanceof Date ? r[TXC.DATE - 1] : null,
      type: String(r[TXC.TYPE - 1]).trim(),
      symbol: sym,
      units: pfNum_(r[TXC.UNITS - 1]),
      price: pfNum_(r[TXC.PRICE - 1]),
      cumUnits: pfNum_(r[TXC.CUM_UNITS - 1]),
      cumCost: pfNum_(r[TXC.CUM_COST - 1]),
    };
    list.push(tx);
    var s = bySymbol[sym] || (bySymbol[sym] = { count: 0 });
    s.count++;
    if (tx.cumUnits !== null) s.units = tx.cumUnits;
    if (tx.cumCost !== null) s.cost = tx.cumCost;
    if (tx.price !== null && (tx.type === 'Buy' || tx.type === 'Sell')) {
      s.lastPrice = tx.price;
      s.lastDate = tx.date;
    }
  });
  return { sheet: sh, lastRow: last, list: list, bySymbol: bySymbol };
}

/** Price sheet: USD/THB rate and, per symbol, whether it is priced in USD. */
function pfReadPrices_(ss) {
  var sh = pfSheet_(ss, PF.PRICE);
  var last = sh.getLastRow();
  var out = { fx: null, bySymbol: {} };
  if (last < 2) return out;
  var rng = sh.getRange(2, 1, last - 1, 3);
  var vals = rng.getValues();
  var forms = rng.getFormulas();
  vals.forEach(function (r, i) {
    var label = String(r[0]).trim();
    if (!label) return;
    if (/^USD\s*\/\s*THB$/i.test(label)) { out.fx = pfNum_(r[1]); return; }
    if (label.indexOf('/') >= 0) return; // other currency pairs
    out.bySymbol[label] = {
      isUsd: pfNum_(r[1]) !== null || forms[i][1] !== '',
      priceUsd: pfNum_(r[1]),
      priceThb: pfNum_(r[2]),
    };
  });
  return out;
}

/** Assets marked "dca" in Portfolio Summary (column G) with the amount in F. */
function pfReadDca_(ss) {
  var sh = pfSheet_(ss, PF.SUMMARY);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 7).getValues()
    .filter(function (r) { return String(r[6]).trim().toLowerCase() === 'dca'; })
    .map(function (r) { return { symbol: String(r[0]).trim(), amount: pfNum_(r[5]) }; });
}

function pfCollect_(ss) {
  return { tx: pfReadTx_(ss), prices: pfReadPrices_(ss), dca: pfReadDca_(ss) };
}

/** Everything the form needs, in one call. */
function pfGetFormData() {
  var ss = pfSs_();
  var tz = ss.getSpreadsheetTimeZone();
  var d = pfCollect_(ss);
  var names = {};
  Object.keys(d.tx.bySymbol).forEach(function (s) { names[s] = true; });
  Object.keys(d.prices.bySymbol).forEach(function (s) { names[s] = true; });
  var dcaAmount = {};
  d.dca.forEach(function (x) { dcaAmount[x.symbol] = x.amount; names[x.symbol] = true; });

  var symbols = Object.keys(names).map(function (s) {
    var t = d.tx.bySymbol[s] || {};
    var p = d.prices.bySymbol[s] || {};
    return {
      symbol: s,
      isUsd: !!p.isUsd,
      units: t.units || 0,
      cost: t.cost || 0,
      count: t.count || 0,
      lastPrice: t.lastPrice || null,
      lastDate: t.lastDate ? Utilities.formatDate(t.lastDate, tz, 'yyyy-MM-dd') : null,
      priceThb: p.priceThb || null,
      priceUsd: p.priceUsd || null,
      dca: s in dcaAmount ? dcaAmount[s] : null,
    };
  }).sort(function (a, b) {
    return (b.dca !== null) - (a.dca !== null) || b.count - a.count || a.symbol.localeCompare(b.symbol);
  });

  var recent = d.tx.list.slice(-6).reverse().map(function (t) {
    return {
      row: t.row, type: t.type, symbol: t.symbol, units: t.units, price: t.price,
      date: t.date ? Utilities.formatDate(t.date, tz, 'yyyy-MM-dd') : '',
    };
  });

  return {
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    fx: d.prices.fx,
    symbols: symbols,
    recent: recent,
  };
}

/* ------------------------------------------------------------- writing -- */

/**
 * entries: [{date:'yyyy-MM-dd', type:'Buy'|'Sell'|'Div', symbol, units, price, fx, fee}]
 *   price is per unit in the asset's own currency (USD when fx is given).
 *   For 'Div', price is the amount received in THB and units is ignored.
 * Returns {ok, rows} or {ok:false, errors|needConfirm, warnings}.
 */
function pfSaveTransactions(entries, confirmed) {
  if (!entries || !entries.length) return { ok: false, errors: ['ไม่มีรายการให้บันทึก'], warnings: [] };
  var ss = pfSs_();
  var tz = ss.getSpreadsheetTimeZone();
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var d = pfCollect_(ss);
    var check = pfValidate_(entries, d, tz);
    if (check.errors.length) return { ok: false, errors: check.errors, warnings: check.warnings };
    if (check.warnings.length && !confirmed) return { ok: false, needConfirm: true, warnings: check.warnings };

    var sh = d.tx.sheet;
    var last = d.tx.lastRow;
    if (last < PF.FIRST_TEMPLATE_ROW) throw new Error('ต้องมีรายการใน Transaction อย่างน้อย 2 แถวเพื่อใช้เป็นแม่แบบสูตร');
    var start = last + 1;
    var n = entries.length;
    var need = start + n - 1 - sh.getMaxRows();
    if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);

    // Formulas (H:P), formats and validation come from the last row, as if
    // the row had been dragged down by hand. Inputs are then overwritten.
    sh.getRange(last, 1, 1, PF.TX_WIDTH).copyTo(sh.getRange(start, 1, n, PF.TX_WIDTH));
    sh.getRange(start, TXC.DATE, n, 7).setValues(entries.map(function (e) { return pfRowInputs_(e, tz); }));
    sh.getRange(start, TXC.FX, n, 1).setValues(entries.map(function (e) {
      return [e.type !== 'Div' && pfNum_(e.fx) ? e.fx : ''];
    }));
    SpreadsheetApp.flush();
    return { ok: true, rows: [start, start + n - 1], warnings: check.warnings };
  } finally {
    lock.releaseLock();
  }
}

/** Values for columns A:G of one Transaction row. */
function pfRowInputs_(e, tz) {
  var date = Utilities.parseDate(e.date, tz, 'yyyy-MM-dd');
  var fee = pfNum_(e.fee) || 0;
  if (e.type === 'Div') return [date, 'Div', e.symbol, 1, e.price, fee, 1];
  // USD assets keep the "=usdPrice*fx" form used in the sheet so far.
  var price = pfNum_(e.fx) ? '=' + e.price + '*' + e.fx : e.price;
  return [date, e.type, e.symbol, e.units, price, fee, 1];
}

function pfValidate_(entries, d, tz) {
  var errors = [];
  var warnings = [];
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var oldLimit = Utilities.formatDate(new Date(Date.now() - PF.OLD_DATE_WARN_DAYS * 864e5), tz, 'yyyy-MM-dd');
  var holding = {};

  entries.forEach(function (e, i) {
    var tag = (entries.length > 1 ? '#' + (i + 1) + ' ' : '') + (e.symbol || '?') + ': ';
    e.symbol = String(e.symbol || '').trim();
    e.units = pfNum_(Number(e.units));
    e.price = pfNum_(Number(e.price));
    e.fee = pfNum_(Number(e.fee)) || 0;
    e.fx = e.fx === '' || e.fx === null || e.fx === undefined ? null : pfNum_(Number(e.fx));

    if (['Buy', 'Sell', 'Div'].indexOf(e.type) < 0) errors.push(tag + 'ประเภทต้องเป็น Buy, Sell หรือ Div');
    if (!e.symbol) errors.push(tag + 'ยังไม่ได้เลือกสินทรัพย์');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(e.date || '')) errors.push(tag + 'วันที่ไม่ถูกต้อง');
    else if (e.date > today) errors.push(tag + 'วันที่ ' + e.date + ' อยู่ในอนาคต');
    else if (e.date < oldLimit) warnings.push(tag + 'วันที่ ' + e.date + ' ย้อนหลังเกิน ' + PF.OLD_DATE_WARN_DAYS + ' วัน');

    if (e.type === 'Div') {
      if (!(e.price > 0)) errors.push(tag + 'กรอกจำนวนเงินปันผล');
      return;
    }
    if (!(e.units > 0)) errors.push(tag + 'จำนวนหน่วยต้องมากกว่า 0');
    if (!(e.price > 0)) errors.push(tag + 'ราคาต่อหน่วยต้องมากกว่า 0');
    if (e.fx !== null && !(e.fx > 0)) errors.push(tag + 'อัตราแลกเปลี่ยนไม่ถูกต้อง');
    if (!(e.units > 0) || !(e.price > 0)) return;

    var t = d.tx.bySymbol[e.symbol];
    var p = d.prices.bySymbol[e.symbol];
    if (!t && !p) warnings.push(tag + 'เป็นสินทรัพย์ใหม่ อย่าลืมเพิ่มแถวในชีต Price / Cost / Portfolio Summary');
    if (p && p.isUsd && e.fx === null) warnings.push(tag + 'สินทรัพย์นี้ซื้อเป็น USD แต่ไม่ได้ใส่อัตราแลกเปลี่ยน');

    // Sanity check against today's price (Price sheet), else the last price
    // typed in. The market price comes first so one bad row can't hide the next.
    var thb = e.price * (e.fx || 1);
    var ref = (p && p.priceThb > 0) ? p.priceThb : (t && t.lastPrice > 0 ? t.lastPrice : null);
    if (ref && Math.abs(thb / ref - 1) > PF.PRICE_WARN) {
      warnings.push(tag + 'ราคา ฿' + pfFmt_(thb) + ' ต่างจากราคาอ้างอิง ฿' + pfFmt_(ref) + ' อยู่ ' +
        Math.round(Math.abs(thb / ref - 1) * 100) + '% — กรอกสลับช่องหรือใส่หน่วยผิดหรือเปล่า?');
    }

    if (!(e.symbol in holding)) holding[e.symbol] = (t && t.units) || 0;
    if (e.type === 'Sell' && e.units > holding[e.symbol] + 1e-9) {
      errors.push(tag + 'ขาย ' + e.units + ' หน่วย แต่ถืออยู่ ' + holding[e.symbol] + ' หน่วย');
    }
    holding[e.symbol] += e.type === 'Sell' ? -e.units : e.units;

    var dup = d.tx.list.some(function (x) {
      return x.symbol === e.symbol && x.date && x.units !== null &&
        Utilities.formatDate(x.date, tz, 'yyyy-MM-dd') === e.date && Math.abs(x.units - e.units) < 1e-9;
    });
    if (dup) warnings.push(tag + 'มีรายการวันที่และจำนวนหน่วยเดียวกันอยู่แล้ว (บันทึกซ้ำหรือเปล่า?)');
  });
  return { errors: errors, warnings: warnings };
}

function pfFmt_(n) {
  return Number(n).toLocaleString('en-US', { maximumFractionDigits: 2 });
}

/* ------------------------------------------------------------ snapshot -- */

function pfNorm_(s) {
  var k = String(s).toLowerCase().replace(/\s+/g, '');
  return PF.LABEL_ALIASES[k] || k;
}

function pfParseHeaderDate_(v) {
  if (v instanceof Date) return v;
  var m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/); // d/m/yy as typed in the sheet
  if (!m) return null;
  var y = Number(m[3]);
  return new Date(y < 100 ? 2000 + y : y, Number(m[2]) - 1, Number(m[1]));
}

/** Works out what a new Monthly Summary column would contain. */
function pfBuildSnapshot_(ss) {
  var mon = pfSheet_(ss, PF.MONTHLY);
  var sum = pfSheet_(ss, PF.SUMMARY);
  var lastRow = mon.getLastRow();
  var lastCol = Math.max(mon.getLastColumn(), 2);
  var grid = mon.getRange(1, 1, lastRow, lastCol).getValues();

  // Next column = first one after the last column holding any numbers.
  var col = 2;
  for (var c = lastCol; c >= 2; c--) {
    var hasData = false;
    for (var r = 1; r < lastRow; r++) if (pfNum_(grid[r][c - 1]) !== null) { hasData = true; break; }
    if (hasData) { col = c + 1; break; }
  }
  var prevHeader = col > 2 ? pfParseHeaderDate_(grid[0][col - 2]) : null;

  // Labels above "Total" are the holdings; the rows below it (gain, FX) are not.
  var current = {};
  var holdings = [];
  var pastTotal = false;
  sum.getRange(1, 1, sum.getLastRow(), 3).getValues().forEach(function (r) {
    var label = String(r[0]).trim();
    if (!label) return;
    current[pfNorm_(label)] = r[2];
    if (pfNorm_(label) === 'total') pastTotal = true;
    if (!pastTotal && pfNum_(r[2])) holdings.push(label); // skip empty/zero rows
  });
  var monthlyKeys = {};
  for (var j = 1; j < lastRow; j++) monthlyKeys[pfNorm_(grid[j][0])] = true;
  // e.g. a new stock added to Portfolio Summary but not to Monthly Summary
  var missing = holdings.filter(function (l) { return !monthlyKeys[pfNorm_(l)]; });

  var values = [];
  var unmatched = [];
  var broken = [];
  for (var i = 1; i < lastRow; i++) {
    var label = String(grid[i][0]).trim();
    if (!label) { values.push(['']); continue; }
    var k = pfNorm_(label);
    if (!(k in current)) { unmatched.push(label); values.push(['']); continue; }
    var v = current[k];
    if (typeof v === 'string' && v.charAt(0) === '#') { broken.push(label + ' (' + v + ')'); v = ''; }
    values.push([v]);
  }
  return { sheet: mon, col: col, lastRow: lastRow, prevHeader: prevHeader,
           values: values, unmatched: unmatched, broken: broken, missing: missing };
}

function pfWriteSnapshot_(snap, tz) {
  var sh = snap.sheet;
  if (snap.col > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), snap.col - sh.getMaxColumns());
  if (snap.col > 2) {
    sh.getRange(1, snap.col - 1, snap.lastRow, 1)
      .copyTo(sh.getRange(1, snap.col, snap.lastRow, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
  sh.getRange(1, snap.col).setValue(Utilities.parseDate(
    Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), tz, 'yyyy-MM-dd'));
  sh.getRange(2, snap.col, snap.values.length, 1).setValues(snap.values);
}

function pfSnapshotNow() {
  var ui = SpreadsheetApp.getUi();
  var ss = pfSs_();
  var snap = pfBuildSnapshot_(ss);
  var notes = [];
  var now = new Date();
  if (snap.prevHeader && snap.prevHeader.getFullYear() === now.getFullYear() &&
      snap.prevHeader.getMonth() === now.getMonth()) {
    notes.push('เดือนนี้มี snapshot แล้ว — จะเพิ่มคอลัมน์ใหม่อีกคอลัมน์');
  }
  if (snap.broken.length) notes.push('ค่าที่เป็น error (จะเว้นว่างไว้):\n  ' + snap.broken.join('\n  '));
  if (snap.unmatched.length) notes.push('หาแถวที่ตรงกันใน Portfolio Summary ไม่เจอ:\n  ' + snap.unmatched.join('\n  '));
  if (snap.missing.length) notes.push('มีใน Portfolio Summary แต่ยังไม่มีแถวใน Monthly Summary (มูลค่าจะไม่ถูกบันทึก — เพิ่มแถวก่อน):\n  ' + snap.missing.join('\n  '));
  var colName = sheetColumnName_(snap.col);
  var msg = 'จะบันทึกค่าปัจจุบันจาก Portfolio Summary ลงคอลัมน์ ' + colName + ' ของ Monthly Summary' +
    (notes.length ? '\n\n' + notes.join('\n\n') : '') + '\n\nดำเนินการต่อ?';
  if (ui.alert('Snapshot', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  pfWriteSnapshot_(snap, ss.getSpreadsheetTimeZone());
  ss.toast('บันทึก snapshot ลงคอลัมน์ ' + colName + ' แล้ว', 'Portfolio');
}

/** Time-driven version: skips if this month already has a column. */
function pfSnapshotAuto() {
  var ss = pfSs_();
  var snap = pfBuildSnapshot_(ss);
  var now = new Date();
  if (snap.prevHeader && snap.prevHeader.getFullYear() === now.getFullYear() &&
      snap.prevHeader.getMonth() === now.getMonth()) return;
  pfWriteSnapshot_(snap, ss.getSpreadsheetTimeZone());
}

function pfInstallMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pfSnapshotAuto') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pfSnapshotAuto').timeBased()
    .onMonthDay(PF.SNAPSHOT_DAY).atHour(PF.SNAPSHOT_HOUR).create();
  SpreadsheetApp.getUi().alert('ตั้งเวลาแล้ว: snapshot ลง Monthly Summary ทุกวันที่ ' +
    PF.SNAPSHOT_DAY + ' เวลาประมาณ ' + PF.SNAPSHOT_HOUR + ':00 น.');
}

function sheetColumnName_(n) {
  var s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

/* --------------------------------------------------------------- audit -- */

/** Lists rows that look wrong: future dates, error values, odd prices. */
function pfAuditFindings_(ss) {
  var tz = ss.getSpreadsheetTimeZone();
  var d = pfReadTx_(ss);
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var out = [];
  var bySym = {};
  d.list.forEach(function (t) {
    if (t.date && Utilities.formatDate(t.date, tz, 'yyyy-MM-dd') > today) {
      out.push('แถว ' + t.row + ' ' + t.symbol + ': วันที่ ' + Utilities.formatDate(t.date, tz, 'd/M/yyyy') + ' อยู่ในอนาคต');
    }
    if (t.cumUnits === null || t.cumCost === null) {
      out.push('แถว ' + t.row + ' ' + t.symbol + ': คอลัมน์ Cumulative Units / Cost คำนวณไม่ได้');
    }
    if (t.type === 'Buy' && t.price !== null) (bySym[t.symbol] = bySym[t.symbol] || []).push(t);
  });
  // A price far from both the previous and the next price of the same asset.
  Object.keys(bySym).forEach(function (s) {
    var a = bySym[s];
    a.forEach(function (t, i) {
      var refs = [a[i - 1], a[i + 1]].filter(Boolean).map(function (x) { return x.price; });
      if (!refs.length) return;
      var diff = Math.min.apply(null, refs.map(function (r) { return Math.abs(t.price / r - 1); }));
      if (diff > PF.PRICE_WARN) {
        out.push('แถว ' + t.row + ' ' + s + ': ราคา ' + pfFmt_(t.price) + ' ต่างจากรายการก่อน/หลัง (' +
          refs.map(pfFmt_).join(', ') + ') ' + Math.round(diff * 100) + '%');
      }
    });
  });
  return out;
}

function pfAudit() {
  var out = pfAuditFindings_(pfSs_());
  SpreadsheetApp.getUi().alert('ตรวจข้อมูล Transaction',
    out.length ? out.join('\n') : 'ไม่พบสิ่งผิดปกติ 👍', SpreadsheetApp.getUi().ButtonSet.OK);
}
