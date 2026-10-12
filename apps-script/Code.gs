/**
 * Staff Bottle App – Google Sheets backend (template)
 * ---------------------------------------------------
 * Paste this whole file into Extensions > Apps Script of your bottle sheet, then
 * reload the sheet and click  Bottle List > Set up this sheet (first time).
 * Full instructions: README.md in the template repository.
 *
 * Adds a "Bottle List" menu to the sheet:
 *   - Make menu tab   – rebuilds the "Menu (print)" tab from the Bottles tab.
 *   - Make folded PDF – builds the folded, double-sided bottle list PDF and saves it
 *                       in the same Drive folder as this spreadsheet.
 * Rules for the printed list:
 *   - prints rows with status menu, insert or needs-page, and never a row with Y in "Out of stock"
 *   - each line uses the four MENU columns; Vintage and Price come from the main columns
 *   - sections go on the pages listed in the "Menu layout" tab; inside a section,
 *     whites first, then a gap, then reds, each in Sort order
 *   - ** is added for Serve chilled (Y) or Skin contact (Y)
 * It is also the backend of the staff app (publish it as a web app – see README).
 */

var CFG = {
  BOTTLES: 'Bottles',
  LAYOUT: 'Menu layout',
  MENU_TAB: 'Menu (print)',
  SETUP: 'Setup',       // tab with the bar name, page template, cover logo and photo folder
  FONT: 'Lora',
  BODY: 10, HEAD: 12, FOOT: 9.5,
  REDS: ['Light & chillable reds', 'Bright & silky reds', 'Savory & spicy reds',
         'Firm & structured reds', 'Dark & structured reds'],
  DEFAULT_LAYOUT: [[2, 'Sparkling'], [3, 'White'], [4, 'Orange & Pink'], [5, 'Red']],
  KEEP_SLIDES: false   // true keeps the editable Slides copy next to the PDF
};

// The sheet this script belongs to (remembered by "Set up this sheet" so the web app can find it).
function book_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) {}
  return ss || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
}

// A value from the Setup tab (column A = setting name, column B = value).
function setting_(name) {
  var sh = book_().getSheetByName(CFG.SETUP);
  if (!sh) return '';
  var v = sh.getDataRange().getDisplayValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim().toLowerCase() === name.toLowerCase()) return String(v[i][1] || '').trim();
  return '';
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bottle List')
    .addItem('Make menu tab', 'makeMenuTab')
    .addItem('Make folded PDF', 'makeFoldedPdf')
    .addSeparator()
    .addItem('Set up this sheet (first time)', 'setupSheet')
    .addToUi();
}

/* ------------------------------------------------------------------ data */

function readMenu_() {
  var ss = book_();
  var sh = ss.getSheetByName(CFG.BOTTLES);
  if (!sh) throw new Error('No tab called "' + CFG.BOTTLES + '".');
  var vals = sh.getDataRange().getDisplayValues();
  var hdr = vals[0];
  function col(prefix) {
    for (var i = 0; i < hdr.length; i++) if (String(hdr[i]).indexOf(prefix) === 0) return i;
    return -1;
  }
  var C = {
    status: col('menu / off-menu'), section: col('Menu section'), order: col('Sort order'),
    producer: col('Producer'), wine: col('Wine / cuv'), grapes: col('Grapes'), place: col('Place'),
    vintage: col('Vintage'), price: col('Price'), style: col('Taste group'),
    chilled: col('Serve chilled'), skin: col('Skin contact'), oos: col('Out of stock'),
    mProd: col('MENU: producer'), mCuv: col('MENU: cuv'), mStyle: col('MENU: grape'), mPlace: col('MENU: place')
  };
  if (C.status < 0 || C.section < 0 || C.price < 0) throw new Error('Could not find the Status / Menu section / Price columns on the Bottles tab.');
  var t = function (r, k) { return C[k] < 0 ? '' : String(r[C[k]] == null ? '' : r[C[k]]).trim(); };

  var warnings = [], items = [], seen = {};
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i], rowNo = i + 1;
    if (!t(r, 'producer')) continue;
    var st = t(r, 'status').toLowerCase();
    if (st !== 'menu' && st !== 'insert' && st !== 'needs-page') continue;   // sold-out, off-menu, removed never print
    if (t(r, 'oos').toUpperCase() === 'Y') continue;
    var prod, cuv, sty, pl, pendingMenuWarn = '';
    if (t(r, 'mProd')) { prod = t(r, 'mProd'); cuv = t(r, 'mCuv'); sty = t(r, 'mStyle'); pl = t(r, 'mPlace'); }
    else {
      prod = t(r, 'producer'); cuv = t(r, 'wine'); sty = t(r, 'grapes'); pl = t(r, 'place');
      var pendingMenuWarn = 'Row ' + rowNo + ' (' + prod + '): the four MENU columns are empty, so the long booklet wording was used.';
    }
    var v = t(r, 'vintage'), vin = '';
    if (/^NV/i.test(v)) vin = 'NV'; else if (/^\d{4}$/.test(v)) vin = v;
    else if (v) warnings.push('Row ' + rowNo + ' (' + prod + '): vintage "' + v + '" printed without a year.');
    var price = t(r, 'price').replace(/[^0-9.]/g, '').replace(/\.0+$/, '');
    if (!price) warnings.push('Row ' + rowNo + ' (' + prod + '): no price.');
    var so = parseFloat(t(r, 'order'));
    var soWarn = '';
    if (isNaN(so)) { so = 9999; soWarn = 'Row ' + rowNo + ' (' + prod + '): no Sort order – printed at the end of its section.'; }
    var lastWord = t(r, 'producer').split(/\s+/).pop();
    var keys = [[prod, cuv, vin, price].join('|').toLowerCase(),
                ['b', lastWord, t(r, 'wine'), vin, price].join('|').toLowerCase()];
    var dup = keys.map(function (k) { return seen[k]; }).filter(Boolean)[0];
    if (dup) { warnings.push('Row ' + rowNo + ' (' + t(r, 'producer') + ' ' + t(r, 'wine') + ') looks like a duplicate of row ' + dup + ' – skipped. Delete one of them.'); continue; }
    keys.forEach(function (k) { seen[k] = rowNo; });
    if (pendingMenuWarn) warnings.push(pendingMenuWarn);
    if (st === 'needs-page') warnings.push('Row ' + rowNo + ' (' + prod + '): on the guest list, but its booklet page isn\u2019t written yet (status needs-page).');
    if (soWarn) warnings.push(soWarn);
    var star = t(r, 'chilled').toUpperCase() === 'Y' || t(r, 'skin').toUpperCase() === 'Y';
    items.push({ sec: t(r, 'section'), so: so, red: CFG.REDS.indexOf(t(r, 'style')) >= 0,
      prod: prod, cuv: cuv, sty: sty, pl: pl, vin: vin, price: price, star: star, row: rowNo });
  }

  // layout
  var layout = [], footnote = '**skin contact or chilled red';
  var lsh = ss.getSheetByName(CFG.LAYOUT);
  if (lsh) {
    lsh.getDataRange().getDisplayValues().forEach(function (r) {
      var a = String(r[0]).trim(), b = String(r[1] || '').trim();
      if (a.toUpperCase() === 'FOOTNOTE') footnote = b;
      else if (/^\d+$/.test(a) && b) layout.push([parseInt(a, 10), b]);
    });
  }
  if (!layout.length) layout = CFG.DEFAULT_LAYOUT.slice();
  var panels = {}, placed = {};
  layout.forEach(function (p) { (panels[p[0]] = panels[p[0]] || []).push(p[1]); placed[p[1]] = true; });
  var secs = {};
  items.forEach(function (it) { (secs[it.sec] = secs[it.sec] || []).push(it); });
  var nums = Object.keys(panels).map(Number).sort(function (a, b) { return a - b; });
  var last = nums[nums.length - 1];
  Object.keys(secs).forEach(function (s) {
    if (!placed[s]) { panels[last].push(s); warnings.push('Section "' + s + '" is not in the Menu layout tab – added to page ' + last + '.'); }
  });
  Object.keys(secs).forEach(function (s) {
    secs[s].sort(function (a, b) { return (a.red - b.red) || (a.so - b.so); });
  });
  return { panels: panels, nums: nums, secs: secs, footnote: footnote, warnings: warnings, count: items.length };
}

// One menu line as text plus the [start,end) of the bold-italic grape/style part.
function lineText_(it) {
  var head = it.prod + (it.cuv ? ' ‘' + it.cuv + '’' : '');
  var s = head + ' ';
  var biStart = s.length;
  s += it.sty;
  var biEnd = s.length;
  var tail = [it.pl, it.vin].filter(String).join(' ');
  s += ' ' + (tail ? tail + ' ' : '') + '| ' + it.price + (it.star ? '**' : '');
  return { text: s, bi: [biStart, biEnd] };
}

/* ------------------------------------------------------------ menu tab */

function makeMenuTab() {
  var d = readMenu_();
  var ss = book_();
  var sh = ss.getSheetByName(CFG.MENU_TAB) || ss.insertSheet(CFG.MENU_TAB);
  sh.clear();
  sh.setHiddenGridlines(true);
  var rows = [];   // {text, bi, kind}
  var stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MMM d, yyyy h:mm a');
  rows.push({ text: 'Bottle list – built ' + stamp + ' – ' + d.count + ' bottles', kind: 'title' });
  if (d.warnings.length) {
    rows.push({ text: 'CHECK BEFORE PRINTING:', kind: 'warnhead' });
    d.warnings.forEach(function (w) { rows.push({ text: '• ' + w, kind: 'warn' }); });
  } else rows.push({ text: 'No problems found.', kind: 'ok' });
  d.nums.forEach(function (p, pi) {
    rows.push({ text: '', kind: 'gap' });
    rows.push({ text: 'PAGE ' + p, kind: 'page' });
    d.panels[p].forEach(function (s, si) {
      var list = d.secs[s] || [];
      if (!list.length) return;
      if (si) rows.push({ text: '', kind: 'gap' });
      rows.push({ text: s.toUpperCase(), kind: 'head' });
      var prevRed = null;
      list.forEach(function (it) {
        if (prevRed === false && it.red) rows.push({ text: '', kind: 'gap' });
        prevRed = it.red;
        var L = lineText_(it);
        rows.push({ text: L.text, bi: L.bi, kind: 'line' });
      });
    });
    if (pi === d.nums.length - 1 && d.footnote) { rows.push({ text: '', kind: 'gap' }); rows.push({ text: d.footnote, kind: 'foot' }); }
  });

  var n = rows.length;
  var rich = rows.map(function (r) {
    var b = SpreadsheetApp.newRichTextValue().setText(r.text);
    if (r.bi && r.bi[1] > r.bi[0]) b.setTextStyle(r.bi[0], r.bi[1], SpreadsheetApp.newTextStyle().setBold(true).setItalic(true).build());
    return [b.build()];
  });
  var rng = sh.getRange(1, 2, n, 1);
  rng.setFontFamily(CFG.FONT).setFontSize(CFG.BODY).setVerticalAlignment('middle').setWrap(true);
  var sizes = [], weights = [], colors = [];
  rows.forEach(function (r) {
    sizes.push([r.kind === 'head' ? CFG.HEAD : r.kind === 'title' ? 12 : r.kind === 'page' ? 8 : r.kind === 'foot' ? CFG.FOOT : CFG.BODY]);
    weights.push([(r.kind === 'head' || r.kind === 'title' || r.kind === 'warnhead' || r.kind === 'foot') ? 'bold' : 'normal']);
    colors.push([r.kind === 'warn' || r.kind === 'warnhead' ? '#B3261E' : r.kind === 'page' ? '#8A8A8A' : r.kind === 'ok' ? '#2C7447' : '#231F20']);
  });
  rng.setFontSizes(sizes).setFontWeights(weights).setFontColors(colors);
  rng.setRichTextValues(rich);   // last, so the bold-italic grape survives
  sh.setColumnWidth(1, 24);
  sh.setColumnWidth(2, 560);
  sh.setRowHeights(1, n, 21);
  ss.setActiveSheet(sh);
  try { book_().toast(d.count + ' bottles' + (d.warnings.length ? ' – ' + d.warnings.length + ' things to check (top of the tab)' : ''), 'Menu tab rebuilt', 8); } catch (e) {}
}

/* ---------------------------------------------------------- folded PDF */

function makeFoldedPdf(opt) {
  var d = readMenu_();
  var ss = book_();
  var maxPanel = Math.max.apply(null, d.nums);
  var N = Math.ceil((maxPanel + 1) / 4) * 4;          // panels incl. cover and back
  var footPanel = d.nums[d.nums.length - 1];
  var W = 792, H = 612, HALF = 396;

  // 1) copy the empty 11 x 8.5 in template deck (the Slides API can't set a page size on a new deck)
  if (typeof Slides === 'undefined') throw new Error('Turn on the Slides service first: in Apps Script click Services (+) \u2192 Google Slides API \u2192 Add.');
  var folder = DriveApp.getFileById(ss.getId()).getParents().next();
  var title = 'Bottle List \u2013 ' + ss.getName() + ' (building)';
  var tpl = setting_('Page template ID');
  if (!tpl) throw new Error('Add the ID of your 11 x 8.5 in Google Slides page template on the Setup tab first (see README, step 5).');
  var pid = DriveApp.getFileById(tpl).makeCopy(title, folder).getId();
  var pres = Slides.Presentations.get(pid);
  var firstSlide = pres.slides && pres.slides[0] && pres.slides[0].objectId;
  var layoutId = pres.slides && pres.slides[0] && pres.slides[0].slideProperties.layoutObjectId;

  // 2) imposition: sheet k front = [N-2k | 2k+1], back = [2k+2 | N-2k-1]
  var sides = [];
  for (var k = 0; k < N / 4; k++) { sides.push([N - 2 * k, 2 * k + 1]); sides.push([2 * k + 2, N - 2 * k - 1]); }

  var req = [], uid = 0, pageIds = [];
  function id(p) { return p + '_' + (uid++); }
  sides.forEach(function (lr, i) {
    var sid = 'sheet' + i;
    pageIds.push(sid);
    req.push({ createSlide: { objectId: sid, insertionIndex: i, slideLayoutReference: layoutId ? { layoutId: layoutId } : { predefinedLayout: 'BLANK' } } });
  });
  if (firstSlide) req.push({ deleteObject: { objectId: firstSlide } });

  var coverAt = null;
  sides.forEach(function (lr, i) {
    [0, 1].forEach(function (half) {
      var panel = lr[half], x0 = half * HALF;
      if (panel === 1) { coverAt = { slide: i, x: x0 }; return; }
      var secs = d.panels[panel];
      if (!secs) return;
      buildPanelRequests_(req, id, pageIds[i], x0, secs, d.secs, H, panel === footPanel ? d.footnote : '', d.warnings, panel);
    });
  });
  Slides.Presentations.batchUpdate({ requests: req }, pid);

  // 3) cover logo (same position as the printed cover)
  var deck = SlidesApp.openById(pid);
  if (coverAt) {
    try {
      var blob = DriveApp.getFileById(setting_('Cover logo file ID')).getBlob();
      deck.getSlides()[coverAt.slide].insertImage(blob, coverAt.x + 94.7, 54, 202.3, 281.9);
    } catch (e) {
      var tb = deck.getSlides()[coverAt.slide].insertTextBox((setting_('Bar name') || 'Our').toUpperCase() + "\n\nBOTTLE LIST", coverAt.x + 60, 200, 276, 120);
      tb.getText().getTextStyle().setFontFamily(CFG.FONT).setFontSize(28);
      d.warnings.push('No cover logo set on the Setup tab – the cover prints the bar name as text.');
    }
  }
  deck.saveAndClose();

  // 4) export the PDF next to the spreadsheet
  var stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HHmm');
  var pdf = DriveApp.getFileById(pid).getAs('application/pdf').setName('Bottle List PRINT – ' + stamp + '.pdf');
  var file = folder.createFile(pdf);
  var deckFile = DriveApp.getFileById(pid);
  if (CFG.KEEP_SLIDES) deckFile.setName('Bottle List (editable slides) – ' + stamp); else deckFile.setTrashed(true);

  var msg = '<p style="font-family:Arial">Saved <b>' + file.getName() + '</b> in the same Drive folder as this sheet (' + N / 4 +
    ' sheets).<br><a href="' + file.getUrl() + '" target="_blank">Open the PDF</a></p>' +
    '<p style="font-family:Arial">Print it double-sided, <b>flip on short edge</b>, actual size, then fold.</p>';
  if (d.warnings.length) msg += '<p style="font-family:Arial;color:#B3261E"><b>Check:</b><br>' + d.warnings.map(function (w) { return '• ' + w; }).join('<br>') + '</p>';
  if (opt && opt.quiet) return { file: file, warnings: d.warnings, sheets: N / 4 };
  try { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(msg).setWidth(460).setHeight(300), 'Folded PDF ready'); }
  catch (e) { Logger.log(file.getUrl()); }
  return file.getUrl();
}

// Text box for one page of the menu, with styles, sized to fit.
function buildPanelRequests_(req, id, pageId, x0, secNames, secs, H, footnote, warnings, panelNo) {
  var LEFT = 22, TOP = 18, WIDTH = 352, BOTTOM = footnote ? 46 : 24;
  // paragraphs: {text, kind, bi, gap}
  var paras = [];
  secNames.forEach(function (s, si) {
    var list = secs[s] || [];
    if (!list.length) return;
    paras.push({ text: s.toUpperCase(), kind: 'head', gap: si && paras.length ? 'sec' : 'first' });
    var prevRed = null;
    list.forEach(function (it, j) {
      var L = lineText_(it);
      var gap = j === 0 ? 'afterhead' : (prevRed === false && it.red ? 'red' : 'line');
      prevRed = it.red;
      paras.push({ text: L.text, bi: L.bi, kind: 'line', gap: gap });
    });
  });
  if (!paras.length) return;

  // estimate height to pick a scale (Lora 10 pt ≈ 5.05 pt per character on average)
  var avail = H - TOP - BOTTOM;
  var GAP = { first: 0, sec: 13, afterhead: 7.5, red: 15, line: 5 };
  function estimate(sc) {
    var h = 0, cap = Math.floor((WIDTH - 14) / (5.05 * sc));
    paras.forEach(function (p) {
      var lines = p.kind === 'head' ? 1 : Math.ceil(p.text.length / cap);
      h += GAP[p.gap] * sc + lines * 15 * sc;
    });
    return h;
  }
  var sc = 1;
  while (estimate(sc) > avail && sc > 0.84) sc -= 0.01;
  if (estimate(sc) > avail) warnings.push('Page ' + panelNo + ' may run off the bottom – move a section to another page in the Menu layout tab.');

  var text = paras.map(function (p) { return p.text; }).join('\n');
  var box = id('box');
  req.push({ createShape: { objectId: box, shapeType: 'TEXT_BOX', elementProperties: { pageObjectId: pageId,
    size: { width: { magnitude: WIDTH, unit: 'PT' }, height: { magnitude: avail, unit: 'PT' } },
    transform: { scaleX: 1, scaleY: 1, translateX: x0 + LEFT, translateY: TOP, unit: 'PT' } } } });
  req.push({ insertText: { objectId: box, insertionIndex: 0, text: text } });
  req.push({ updateTextStyle: { objectId: box, textRange: { type: 'ALL' },
    style: { fontFamily: CFG.FONT, fontSize: { magnitude: CFG.BODY * sc, unit: 'PT' }, bold: false, italic: false,
      foregroundColor: { opaqueColor: { rgbColor: { red: 0.137, green: 0.122, blue: 0.125 } } } },
    fields: 'fontFamily,fontSize,bold,italic,foregroundColor' } });
  req.push({ updateParagraphStyle: { objectId: box, textRange: { type: 'ALL' },
    style: { lineSpacing: 117, spaceBelow: { magnitude: 0, unit: 'PT' }, spaceAbove: { magnitude: GAP.line * sc, unit: 'PT' } },
    fields: 'lineSpacing,spaceBelow,spaceAbove' } });
  var pos = 0;
  paras.forEach(function (p) {
    var start = pos, end = pos + p.text.length;
    if (p.gap !== 'line') req.push({ updateParagraphStyle: { objectId: box, textRange: { type: 'FIXED_RANGE', startIndex: start, endIndex: end },
      style: { spaceAbove: { magnitude: GAP[p.gap] * sc, unit: 'PT' } }, fields: 'spaceAbove' } });
    if (p.kind === 'head') req.push({ updateTextStyle: { objectId: box, textRange: { type: 'FIXED_RANGE', startIndex: start, endIndex: end },
      style: { bold: true, fontSize: { magnitude: CFG.HEAD * sc, unit: 'PT' } }, fields: 'bold,fontSize' } });
    if (p.bi && p.bi[1] > p.bi[0]) req.push({ updateTextStyle: { objectId: box, textRange: { type: 'FIXED_RANGE', startIndex: start + p.bi[0], endIndex: start + p.bi[1] },
      style: { bold: true, italic: true }, fields: 'bold,italic' } });
    pos = end + 1;
  });
  if (footnote) {
    var fb = id('foot');
    req.push({ createShape: { objectId: fb, shapeType: 'TEXT_BOX', elementProperties: { pageObjectId: pageId,
      size: { width: { magnitude: WIDTH, unit: 'PT' }, height: { magnitude: 22, unit: 'PT' } },
      transform: { scaleX: 1, scaleY: 1, translateX: x0 + LEFT - 4, translateY: H - 40, unit: 'PT' } } } });
    req.push({ insertText: { objectId: fb, insertionIndex: 0, text: footnote } });
    req.push({ updateTextStyle: { objectId: fb, textRange: { type: 'ALL' },
      style: { fontFamily: CFG.FONT, fontSize: { magnitude: CFG.FOOT, unit: 'PT' }, bold: true }, fields: 'fontFamily,fontSize,bold' } });
  }
}

/* ============================================================ STAFF APP
 * Backend for the staff bottle app (published as a web app:
 * Deploy > Manage deployments). The app reads the Bottles + Glossary tabs and
 * marks bottles 86'd by writing Y in "Out of stock". The staff PIN lives in
 * the "Staff app" tab (cell B1); every 86 is logged on the "86 log" tab.
 */
var STAFF = { SETTINGS: 'Staff app', LOG: '86 log', ON_APP: ['menu', 'insert', 'needs-page', 'off-menu'] };

/* Staff codes: the "Staff app" tab holds the app PIN in B1 and a table (header row starting "Name")
 * with each person's 86'd / Add-edit / Print codes. The person's name comes from their code. */
function staffInfo_() {
  var v = book_().getSheetByName(STAFF.SETTINGS).getDataRange().getDisplayValues();
  var appPin = String(v[0] && v[0][1] || '').trim(), people = [], hdr = -1;
  for (var i = 0; i < v.length; i++) {
    if (hdr < 0) { if (String(v[i][0]).trim() === 'Name' && /86/.test(v[i][1])) hdr = i; continue; }
    var name = String(v[i][0]).trim();
    if (!name || /^HOW IT WORKS$/i.test(name)) { if (/^HOW IT WORKS$/i.test(name)) break; continue; }
    people.push({ name: name, c86: String(v[i][1]).trim(), edit: String(v[i][2]).trim(), print: String(v[i][3]).trim() });
  }
  return { appPin: appPin, people: people };
}
// kind: 'c86' | 'edit' | 'print'  -> person's name, or '' if no match
function whoFor_(code, kind, info) {
  code = String(code || '').trim().toLowerCase(); if (!code) return '';
  info = info || staffInfo_();
  for (var i = 0; i < info.people.length; i++) if (info.people[i][kind] && info.people[i][kind].toLowerCase() === code) return info.people[i].name;
  return '';
}
function badTry_(cache, fails) { cache.put('pinfails', String(fails + 1), 600); }

function doGet(e) {
  var p = (e && e.parameter) || {}, a = p.action || 'data';
  try {
    var cache = CacheService.getScriptCache(), fails = Number(cache.get('pinfails') || 0);
    if (fails >= 12) return json_({ ok: false, error: 'locked' });
    var info = staffInfo_();
    if (!info.appPin || String(p.pin || '').trim() !== info.appPin) { badTry_(cache, fails); return json_({ ok: false, error: 'login' }); }
    if (a === 'stock') return json_({ ok: true, stock: stockMap_() });
    if (a === 'log') return json_(changeLog_());
    return json_(catalog_());
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (x) { return json_({ ok: false, error: 'bad_request' }); }
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pinfails') || 0);
  if (fails >= 12) return json_({ ok: false, error: 'locked' });
  if (body.action === 'printmenu') return json_(printMenu_(body, cache, fails));
  if (/^b(auth|check|add|edit)$/.test(body.action || '')) return json_(botRoute_(body, cache, fails));
  if (body.action === 'undo') return json_(undo_(body, cache, fails));
  // 86'd: a person's own 86'd code (name from the table) or the app PIN (name typed)
  var info = staffInfo_(), code = String(body.pin || '').trim(), name = whoFor_(code, 'c86', info), typed = false;
  if (!name && info.appPin && code === info.appPin) { name = String(body.name || '').trim().slice(0, 40); typed = true; }
  if (!name && !typed) { badTry_(cache, fails); return json_({ ok: false, error: 'bad_pin' }); }
  if (body.action === 'check') return json_({ ok: true, name: name, typed: typed });
  if (body.action === '86') {
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try { return json_(set86_(String(body.id || ''), !!body.out, name + (typed && name ? ' (app PIN)' : typed ? '(app PIN)' : ''))); }
    finally { lock.releaseLock(); }
  }
  return json_({ ok: false, error: 'unknown_action' });
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function bottlesSheet_() {
  var ss = book_();
  var sh = ss.getSheetByName(CFG.BOTTLES);
  var vals = sh.getDataRange().getDisplayValues();
  var hdr = vals[0];
  function col(p) { for (var i = 0; i < hdr.length; i++) if (String(hdr[i]).indexOf(p) === 0) return i; return -1; }
  var C = {
    status: col('menu / off-menu'), section: col('Menu section'), order: col('Sort order'), producer: col('Producer'),
    wine: col('Wine / cuv'), grapes: col('Grapes'), place: col('Place'), vintage: col('Vintage'), price: col('Price'),
    taste: col('Taste group'), chilled: col('Serve chilled'), skin: col('Skin contact'), oos: col('Out of stock'),
    remember: col('REMEMBER IT BY'), tasting: col('TASTING'), terroir: col('TERROIR'), producerNotes: col('PRODUCER & WINEMAKING'),
    glossary: col('Glossary terms'), id: col('ID (do not edit)'), mProd: col('MENU: producer'), mCuv: col('MENU: cuv'),
    mStyle: col('MENU: grape'), mPlace: col('MENU: place'), sub: col('Taste sub-group'), isNew: col('New this month'),
    reductive: col('Slightly reductive'), addedBy: col('Added by'), addedOn: col('Added on'), editedBy: col('Last edited by'), editedOn: col('Last edited on'),
    photo: col('Photo file')
  };
  return { ss: ss, sh: sh, vals: vals, C: C };
}

function rowId_(r, C, rowNo) { var v = C.id >= 0 ? String(r[C.id]).trim() : ''; return v || ('row' + rowNo); }

function catalog_() {
  var B = bottlesSheet_(), C = B.C;
  var t = function (r, k) { return C[k] < 0 ? '' : String(r[C[k]] == null ? '' : r[C[k]]).trim(); };
  var lines = function (s) { return s.split('\n').map(function (x) { return x.replace(/^[•\-*]\s*/, '').trim(); }).filter(String); };
  var REDS = ['Light & chillable reds', 'Bright & silky reds', 'Savory & spicy reds', 'Firm & structured reds', 'Dark & structured reds'];
  var color = function (tg) { if (REDS.indexOf(tg) >= 0) return 'Red'; if (tg === 'Bubbles') return 'Bubbles'; if (/^Orange/.test(tg)) return 'Orange'; if (tg === 'Pink') return 'Pink'; return 'White'; };
  // glossary
  var gl = {}, lut = {};
  var gsh = B.ss.getSheetByName('Glossary');
  if (gsh) gsh.getDataRange().getDisplayValues().slice(1).forEach(function (g) {
    if (!g[0]) return; gl[g[0]] = g[1];
    [g[0]].concat(String(g[2] || '').split('|')).forEach(function (k) { if (k.trim()) lut[k.trim().toLowerCase()] = g[0]; });
  });
  var bottles = [];
  for (var i = 1; i < B.vals.length; i++) {
    var r = B.vals[i], st = t(r, 'status').toLowerCase();
    if (!t(r, 'producer') || STAFF.ON_APP.indexOf(st) < 0) continue;
    var v = t(r, 'vintage'), m = v.match(/^(\d{4}|NV)/i);
    var p = t(r, 'price').replace(/[^0-9.]/g, '');
    var so = parseFloat(t(r, 'order'));
    var terms = [];
    t(r, 'glossary').split('·').forEach(function (g) { var k = lut[g.trim().toLowerCase()]; if (k && terms.indexOf(k) < 0) terms.push(k); });
    var red = t(r, 'reductive').toUpperCase() === 'Y';
    if (red && gl['Oxidative, reductive'] && terms.indexOf('Oxidative, reductive') < 0) terms.push('Oxidative, reductive');
    bottles.push({
      id: rowId_(r, C, i + 1), status: st, section: t(r, 'section'), order: isNaN(so) ? 9999 : so,
      producer: t(r, 'producer'), wine: t(r, 'wine'), grapes: t(r, 'grapes'), place: t(r, 'place'),
      vintage: m ? m[1].toUpperCase() : v, price: p ? Math.round(Number(p)) : null,
      taste: t(r, 'taste'), sub: t(r, 'sub'), color: color(t(r, 'taste')),
      chilled: t(r, 'chilled').toUpperCase() === 'Y', skin: t(r, 'skin').toUpperCase() === 'Y',
      remember: t(r, 'remember'), tasting: lines(t(r, 'tasting')), terroir: t(r, 'terroir'), producerNotes: lines(t(r, 'producerNotes')),
      glossary: terms, menu: { prod: t(r, 'mProd'), cuv: t(r, 'mCuv'), style: t(r, 'mStyle'), place: t(r, 'mPlace') },
      isNew: st === 'insert' || st === 'needs-page' || !!t(r, 'isNew'), guestList: st !== 'off-menu',
      reductive: red, out: t(r, 'oos').toUpperCase() === 'Y',
      addedBy: t(r, 'addedBy'), addedOn: t(r, 'addedOn'), editedBy: t(r, 'editedBy'), editedOn: t(r, 'editedOn'),
      photoUrl: /^drive:/.test(t(r, 'photo')) ? 'https://drive.google.com/thumbnail?id=' + t(r, 'photo').slice(6) + '&sz=w1000' : ''
    });
  }
  return { ok: true, edition: B.ss.getName().replace(/^.*?[–-]\s*/, ''), glossary: gl, bottles: bottles, stock: stockMap_(B), sections: sectionOrder_(B.ss) };
}

// Menu sections in the order of the "Menu layout" tab (page by page).
function sectionOrder_(ss) {
  var lsh = ss.getSheetByName(CFG.LAYOUT), rows = [];
  if (lsh) lsh.getDataRange().getDisplayValues().forEach(function (r, i) {
    var a = String(r[0]).trim(), b = String(r[1] || '').trim();
    if (/^\d+$/.test(a) && b) rows.push([parseInt(a, 10), i, b]);
  });
  if (!rows.length) rows = CFG.DEFAULT_LAYOUT.map(function (p, i) { return [p[0], i, p[1]]; });
  rows.sort(function (x, y) { return (x[0] - y[0]) || (x[1] - y[1]); });
  return rows.map(function (r) { return r[2]; });
}

function stockMap_(B) {
  B = B || bottlesSheet_();
  var C = B.C, out = {};
  var last = lastLog_();
  for (var i = 1; i < B.vals.length; i++) {
    var r = B.vals[i];
    if (C.oos >= 0 && String(r[C.oos]).trim().toUpperCase() === 'Y') {
      var id = rowId_(r, C, i + 1);
      out[id] = { out: true, by: (last[id] && last[id].out) ? last[id].name : '', at: (last[id] && last[id].out) ? last[id].at : '' };
    }
  }
  return out;
}

function lastLog_() {
  var sh = book_().getSheetByName(STAFF.LOG), res = {};
  if (!sh || sh.getLastRow() < 2) return res;
  var n = Math.min(400, sh.getLastRow() - 1);
  var v = sh.getRange(sh.getLastRow() - n + 1, 1, n, 5).getValues();
  v.forEach(function (r) { if (r[4]) res[r[4]] = { at: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), out: r[2] === "86'd", name: String(r[3] || '') }; });
  return res;
}

function set86_(id, out, name) {
  var B = bottlesSheet_(), C = B.C;
  if (C.oos < 0) return { ok: false, error: 'no_out_of_stock_column' };
  for (var i = 1; i < B.vals.length; i++) {
    var r = B.vals[i];
    if (rowId_(r, C, i + 1) !== id) continue;
    B.sh.getRange(i + 1, C.oos + 1).setValue(out ? 'Y' : '');
    var label = [r[C.producer], r[C.wine] ? '‘' + r[C.wine] + '’' : '', r[C.vintage]].filter(String).join(' ');
    var log = B.ss.getSheetByName(STAFF.LOG) || B.ss.insertSheet(STAFF.LOG);
    log.appendRow([new Date(), label, out ? "86'd" : 'Back in stock', name || '', id]);
    log.getRange('A:A').setNumberFormat('ddd mm/dd/yy h:mm AM/PM');
    return { ok: true, id: id, out: out, at: new Date().toISOString(), by: name || '' };
  }
  return { ok: false, error: 'not_found' };
}


/* Print menu (from the app): needs a person's print code (Staff app tab). Rebuilds the Menu (print) tab and the folded PDF,
 * saves the PDF next to this sheet and sends it back to the phone. */
function printMenu_(body, cache, fails) {
  var who = whoFor_(body.pw, 'print');
  if (!who) { badTry_(cache, fails); return { ok: false, error: 'bad_pw' }; }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, error: 'busy' };
  try {
    makeMenuTab();
    var r = makeFoldedPdf({ quiet: true });
    var blob = r.file.getBlob();
    botLog_(book_(), r.file.getName(), 'Printed menu PDF', who, '');
    return { ok: true, by: who, name: r.file.getName(), url: r.file.getUrl(), sheets: r.sheets, warnings: r.warnings,
             pdf: Utilities.base64Encode(blob.getBytes()) };
  } finally { lock.releaseLock(); }
}

/* ============================================================ ADD / EDIT BOTTLES (from the app)
 * Needs a person's add / edit code (Staff app tab). New bottles go straight onto the menu. Every add / edit / re-add is
 * written to the "Bottle log" tab and stamped on the row (Added by / on, Last edited by / on).
 * Label photos go to the Drive folder named on the Setup tab (or the sheet's own folder), viewable
 * with the link so the app can show them.
 */
var BOT = { LOG: 'Bottle log', STAMP: 'EEE MM/dd/yy · h:mm a' };

function botAuth_(body, cache, fails) {
  var who = whoFor_(body.pw, 'edit');
  if (!who) { badTry_(cache, fails); return false; }
  body.name = who;   // the name always comes from the person's code
  return true;
}

function botStamp_(ss) { return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), BOT.STAMP); }

function norm_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[‘’'\x60".,()\-–—\/]/g, ' ').replace(/\b(domaine|dom|chateau|ch|maison|weingut|cantina|azienda agricola|the|les|le|la|de|du|des|di|da)\b/g, ' ')
    .replace(/\s+/g, ' ').trim();
}
function words_(s) { return norm_(s).split(' ').filter(function (w) { return w.length > 2; }); }
function overlap_(a, b) { var B = words_(b); return words_(a).filter(function (w) { return B.indexOf(w) >= 0; }).length; }

function botLog_(ss, label, action, name, id, undo) {
  var log = ss.getSheetByName(BOT.LOG);
  if (!log) {
    log = ss.insertSheet(BOT.LOG);
    log.appendRow(['When', 'Bottle', 'Action', 'Name', 'ID', 'Undo data (for the app – don\'t edit)', 'Undone']);
    log.setFrozenRows(1);
  }
  if (!log.getRange(1, 6).getValue()) log.getRange(1, 6, 1, 2).setValues([['Undo data (for the app – don\'t edit)', 'Undone']]);
  log.appendRow([new Date(), label, action, name || '', id, undo || '', '']);
  log.getRange('A:A').setNumberFormat('ddd mm/dd/yy h:mm AM/PM');
}

function botLabel_(r, C) { return [r[C.producer], r[C.wine] ? '‘' + r[C.wine] + '’' : '', r[C.vintage]].filter(String).join(' '); }

// Duplicate / 86'd / sold-out check across EVERY row of the Bottles tab.
function botCheck_(b) {
  var B = bottlesSheet_(), C = B.C, last = lastLog_(), out = [];
  var p = b.producer || '', w = b.wine || '';
  for (var i = 1; i < B.vals.length; i++) {
    var r = B.vals[i];
    if (!r[C.producer]) continue;
    var rp = r[C.producer], rw = r[C.wine];
    var np = norm_(p), nrp = norm_(rp), nw = norm_(w), nrw = norm_(rw);
    var prodSame = np && nrp && (np === nrp || np.indexOf(nrp) >= 0 || nrp.indexOf(np) >= 0);
    var prodNear = prodSame || overlap_(p, rp) > 0;
    var wineSame = nw === nrw;
    var wineNear = wineSame || (nw && nrw && (nw.indexOf(nrw) >= 0 || nrw.indexOf(nw) >= 0)) || overlap_(w, rw) > 0;
    var kind = prodSame && wineSame ? 'same' : (prodNear && wineNear) ? 'similar' : '';
    if (!kind) continue;
    var id = rowId_(r, C, i + 1), st = String(r[C.status]).trim().toLowerCase(), oos = String(r[C.oos]).trim().toUpperCase() === 'Y';
    var L = last[id] || {};
    out.push({ id: id, row: i + 1, kind: kind, producer: rp, wine: rw, vintage: r[C.vintage], price: r[C.price],
      status: st, out: oos, outBy: (oos && L.out) ? L.name : '', outAt: (oos && L.out) ? L.at : '',
      sameVintage: String(r[C.vintage]).trim() === String(b.vintage || '').trim() });
  }
  out.sort(function (a, b) { return (a.kind === 'same' ? 0 : 1) - (b.kind === 'same' ? 0 : 1); });
  return { ok: true, matches: out.slice(0, 8) };
}

function botPhoto_(b64, label) {
  if (!b64) return '';
  var ss = book_();
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', label.replace(/[\/\\:*?"<>|]/g, '') + ' (app ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HHmm') + ').jpg');
  var fid = setting_('Label photo folder ID');
  var folder = fid ? DriveApp.getFolderById(fid) : DriveApp.getFileById(ss.getId()).getParents().next();
  var f = folder.createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'drive:' + f.getId();
}

// Column map for the extra app columns (found by header text).
function botCols_(hdr) {
  function col(p) { for (var i = 0; i < hdr.length; i++) if (String(hdr[i]).indexOf(p) === 0) return i; return -1; }
  return { addedBy: col('Added by'), addedOn: col('Added on'), editedBy: col('Last edited by'), editedOn: col('Last edited on'),
           photo: col('Photo file'), verified: col('Facts verified'), questions: col('Open questions') };
}

var BOT_FIELDS = ['producer', 'wine', 'grapes', 'place', 'vintage', 'price', 'section', 'taste', 'sub', 'chilled', 'skin', 'mStyle', 'mPlace'];
var BOT_NAMES = { producer: 'producer', wine: 'wine', grapes: 'grapes', place: 'place', vintage: 'vintage', price: 'price', section: 'section', taste: 'style group',
  sub: 'sub-group', chilled: 'serve chilled', skin: 'skin contact', mStyle: 'menu grape/style', mPlace: 'menu place' };

function botSet_(row, C, k, v) {
  var c = C[k]; if (c == null || c < 0) return;
  if (k === 'chilled' || k === 'skin') v = v ? 'Y' : '';
  if (k === 'price') v = String(v).replace(/[^0-9.]/g, '') ? Number(String(v).replace(/[^0-9.]/g, '')) : '';
  row[c] = v;
}

function botAdd_(body) {
  var B = bottlesSheet_(), C = B.C, X = botCols_(B.vals[0]), ss = B.ss, b = body.bottle || {};
  if (!b.producer || !b.vintage || !b.price || !b.section || !b.taste) return { ok: false, error: 'missing' };
  var width = B.vals[0].length, row = [];
  for (var k = 0; k < width; k++) row.push('');
  BOT_FIELDS.forEach(function (f) { if (b[f] != null) botSet_(row, C, f, b[f]); });
  row[C.status] = 'menu';
  // sort order: after the last bottle of the same kind (white or red) in that section
  var isRed = CFG.REDS.indexOf(b.taste) >= 0, best = null, any = null, maxNew = 0;
  for (var i = 1; i < B.vals.length; i++) {
    var r = B.vals[i], so = parseFloat(r[C.order]);
    var m = String(r[C.id] || '').match(/^new(\d+)$/); if (m) maxNew = Math.max(maxNew, +m[1]);
    if (r[C.section] !== b.section || isNaN(so)) continue;
    any = any == null ? so : Math.max(any, so);
    if ((CFG.REDS.indexOf(r[C.taste]) >= 0) === isRed) best = best == null ? so : Math.max(best, so);
  }
  row[C.order] = Math.round(((best != null ? best : any != null ? any : 0) + 0.1) * 100) / 100;
  var id = 'new' + (maxNew + 1);
  row[C.id] = id;
  row[C.mProd] = b.producer; row[C.mCuv] = b.wine || ''; row[C.mStyle] = b.grapes || ''; row[C.mPlace] = b.place || '';
  if (C.isNew >= 0) row[C.isNew] = 'new on menu';
  var label = [b.producer, b.wine ? '‘' + b.wine + '’' : '', b.vintage].filter(String).join(' ');
  var stamp = botStamp_(ss);
  if (X.addedBy >= 0) row[X.addedBy] = body.name || '';
  if (X.addedOn >= 0) row[X.addedOn] = stamp;
  if (X.verified >= 0) row[X.verified] = 'NO';
  if (X.questions >= 0) row[X.questions] = 'Added from the app by ' + (body.name || '?') + ' (' + stamp + '). Booklet page (remember it by, tasting, terroir, producer) still to be researched.';
  if (X.photo >= 0 && body.photo) row[X.photo] = botPhoto_(body.photo, label);
  B.sh.getRange(B.vals.length + 1, 1, 1, width).setValues([row]);
  botLog_(ss, label, 'Added to the menu', body.name, id, JSON.stringify({ __add: true }));
  return { ok: true, id: id, catalog: catalog_() };
}

function botEdit_(body) {
  var B = bottlesSheet_(), C = B.C, X = botCols_(B.vals[0]), ss = B.ss, b = body.bottle || {};
  for (var i = 1; i < B.vals.length; i++) {
    var r = B.vals[i];
    if (rowId_(r, C, i + 1) !== body.id) continue;
    var changes = [], set = {};   // column index -> new value (only changed cells are written)
    BOT_FIELDS.forEach(function (f) {
      if (b[f] === undefined || C[f] < 0) return;
      var before = String(r[C[f]]).trim();
      var tmp = []; botSet_(tmp, C, f, b[f]); var after = String(tmp[C[f]] == null ? '' : tmp[C[f]]).trim();
      if (f === 'price') before = before.replace(/[^0-9.]/g, '');
      if (before === after) return;
      set[C[f]] = tmp[C[f]];
      changes.push(BOT_NAMES[f] + ' ' + (before || '–') + ' → ' + (after || '–'));
      // keep the printed-menu wording in step
      if (f === 'producer') set[C.mProd] = after;
      if (f === 'wine') set[C.mCuv] = after;
      if (f === 'grapes') set[C.mStyle] = after;
      if (f === 'place') set[C.mPlace] = after;
    });
    var action = 'Edited';
    if (body.restore) {
      if (String(r[C.oos]).trim().toUpperCase() === 'Y') { set[C.oos] = ''; changes.push('back in stock'); }
      var st = String(r[C.status]).trim().toLowerCase();
      if (st !== 'menu' && st !== 'insert' && st !== 'needs-page') { set[C.status] = 'menu'; changes.push('status ' + st + ' → menu'); }
      action = 'Re-added';
    }
    var label = botLabel_(r, C);
    if (body.photo && X.photo >= 0) { set[X.photo] = botPhoto_(body.photo, label); changes.push('new label photo'); }
    if (!changes.length) return { ok: true, id: body.id, nochange: true, catalog: catalog_() };
    var stamp = botStamp_(ss);
    if (X.editedBy >= 0) set[X.editedBy] = body.name || '';
    if (X.editedOn >= 0) set[X.editedOn] = stamp;
    var undo = {};
    Object.keys(set).forEach(function (c) { if (Number(c) === X.editedBy || Number(c) === X.editedOn) return; undo[B.vals[0][c]] = r[c]; });
    Object.keys(set).forEach(function (c) { B.sh.getRange(i + 1, Number(c) + 1).setValue(set[c]); });
    botLog_(ss, label, action + ': ' + changes.join('; '), body.name, body.id, JSON.stringify(undo));
    return { ok: true, id: body.id, changes: changes, catalog: catalog_() };
  }
  return { ok: false, error: 'not_found' };
}

function botRoute_(body, cache, fails) {
  if (!botAuth_(body, cache, fails)) return { ok: false, error: 'bad_pw' };
  if (body.action === 'bauth') return { ok: true, name: body.name };
  if (body.action === 'bcheck') return botCheck_(body.bottle || {});
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'busy' };
  try {
    if (body.action === 'badd') return botAdd_(body);
    if (body.action === 'bedit') return botEdit_(body);
  } finally { lock.releaseLock(); }
  return { ok: false, error: 'unknown_action' };
}


/* ============================================================ CHANGE LOG + UNDO (from the app)
 * The Change Log tab in the app lists the "86 log" and "Bottle log" tabs together, newest first.
 * Undo needs a person's add / edit code. 86 log column F and Bottle log column G record who undid a change. */
function changeLog_() {
  var ss = book_(), tz = ss.getSpreadsheetTimeZone(), out = [];
  function fmt(d) { return d instanceof Date ? Utilities.formatDate(d, tz, BOT.STAMP) : String(d || ''); }
  var a = ss.getSheetByName(STAFF.LOG);
  if (a && a.getLastRow() > 1) a.getRange(2, 1, a.getLastRow() - 1, 6).getValues().forEach(function (r, k) {
    if (!r[0] && !r[1]) return;
    out.push({ src: '86', row: k + 2, t: r[0] instanceof Date ? r[0].getTime() : 0, when: fmt(r[0]), bottle: String(r[1]), action: String(r[2]),
      who: String(r[3] || ''), id: String(r[4] || ''), undone: String(r[5] || ''), canUndo: !r[5] && !!r[4] && (r[2] === "86'd" || r[2] === 'Back in stock') });
  });
  var b = ss.getSheetByName(BOT.LOG);
  if (b && b.getLastRow() > 1) b.getRange(2, 1, b.getLastRow() - 1, 7).getValues().forEach(function (r, k) {
    if (!r[0] && !r[1]) return;
    out.push({ src: 'bottle', row: k + 2, t: r[0] instanceof Date ? r[0].getTime() : 0, when: fmt(r[0]), bottle: String(r[1]), action: String(r[2]),
      who: String(r[3] || ''), id: String(r[4] || ''), undone: String(r[6] || ''), canUndo: !r[6] && !!r[5] });
  });
  out.sort(function (x, y) { return y.t - x.t; });
  return { ok: true, entries: out.slice(0, 300) };
}

function undo_(body, cache, fails) {
  var who = whoFor_(body.pw, 'edit');
  if (!who) { badTry_(cache, fails); return { ok: false, error: 'bad_pw' }; }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'busy' };
  try {
    var ss = book_(), stamp = botStamp_(ss), row = Number(body.row);
    if (body.src === '86') {
      var sh = ss.getSheetByName(STAFF.LOG), r = sh.getRange(row, 1, 1, 6).getValues()[0];
      if (r[5]) return { ok: false, error: 'already' };
      if (!r[4] || (r[2] !== "86'd" && r[2] !== 'Back in stock')) return { ok: false, error: 'cant' };
      var res = set86_(String(r[4]), r[2] !== "86'd", who);
      if (!res.ok) return res;
      sh.getRange(1, 6).setValue('Undone');
      sh.getRange(row, 6).setValue('Undone by ' + who + ' · ' + stamp);
      return { ok: true, catalog: catalog_() };
    }
    if (body.src === 'bottle') {
      var lg = ss.getSheetByName(BOT.LOG), q = lg.getRange(row, 1, 1, 7).getValues()[0], u = null;
      if (q[6]) return { ok: false, error: 'already' };
      try { u = JSON.parse(q[5] || ''); } catch (x) {}
      if (!u) return { ok: false, error: 'cant' };
      var B = bottlesSheet_(), C = B.C, hdr = B.vals[0], rn = -1;
      for (var i = 1; i < B.vals.length; i++) if (rowId_(B.vals[i], C, i + 1) === String(q[4])) { rn = i; break; }
      if (rn < 0) return { ok: false, error: 'not_found' };
      var redo = {};
      if (u.__add) { redo[hdr[C.status]] = B.vals[rn][C.status]; B.sh.getRange(rn + 1, C.status + 1).setValue('removed'); }
      else Object.keys(u).forEach(function (h) {
        var c = hdr.indexOf(h); if (c < 0) return;
        redo[h] = B.vals[rn][c]; B.sh.getRange(rn + 1, c + 1).setValue(u[h]);
      });
      botLog_(ss, String(q[1]), 'Undid: ' + q[2], who, String(q[4]), JSON.stringify(redo));
      lg.getRange(row, 7).setValue('Undone by ' + who + ' · ' + stamp);
      return { ok: true, catalog: catalog_() };
    }
    return { ok: false, error: 'cant' };
  } finally { lock.releaseLock(); }
}


/* ============================================================ FIRST-TIME SETUP
 * Bottle List > Set up this sheet (first time). Safe to run again: it only adds
 * tabs and headers that are missing and never deletes anything. */
var TASTE_GROUPS = ['Bubbles', 'Crisp & lean whites', 'Round & textured whites', 'Off-dry & sweet whites', 'Orange & grippy', 'Pink',
  'Light & chillable reds', 'Bright & silky reds', 'Savory & spicy reds', 'Firm & structured reds', 'Dark & structured reds'];
var BOTTLE_HEADERS = ["menu / off-menu / insert / sold-out / needs-page", "Menu section (exact)", "Sort order", "Producer", "Wine / cuvée name", "Grapes", "Place", "Vintage", "Price $ (must match chalk)", "Taste group", "Serve chilled (Y)", "Skin contact (Y)", "Out of stock (Y)", "REMEMBER IT BY - one unique hook", "TASTING (one bullet per line)", "TERROIR - because [site fact], the wine [taste]", "PRODUCER & WINEMAKING (one bullet per line)", "Glossary terms (separate with ·)", "Photo file name (in photos folder)", "Facts verified? (YES/NO)", "Sources for facts", "Open questions for manager (never printed)", "ID (do not edit)", "Old Instinct page", "MENU: producer as printed", "MENU: cuvée as printed (blank if none)", "MENU: grape / style (prints in bold italic)", "MENU: place as printed", "Taste sub-group", "New this month (blank / new page / new on menu)", "Slightly reductive (Y)", "Added by (app)", "Added on (app)", "Last edited by (app)", "Last edited on (app)"];
var STARTER_GLOSSARY = [["AOC, IGP, Vin de France", "France's three origin tiers, strictest first. AOC (or AOP) is a named appellation with rules on grapes, yields and borders; IGP is a wider region with looser rules; Vin de France (VdF) names no place at all. Many natural producers choose VdF on purpose - to use a grape, blend or style the appellation rules don't allow - so VdF often means independent, not cheap.", "AOC|IGP|Vin de France|VdF"], ["Blanc de blancs / blanc de noirs", "Blanc de blancs is white wine from white grapes (in Champagne, usually Chardonnay). Blanc de noirs is white wine pressed gently from red grapes (Pinot Noir or Meunier) so the skins add little or no colour - it tends to taste rounder and more red-fruited.", "Blanc de noirs|Blanc de blancs"], ["Brut, extra brut, brut nature", "How much sugar is added to a sparkling wine after the second fermentation (the dosage). Brut: up to 12 g/l. Extra brut: 0-6 g/l. Brut nature (also 'zero dosage' or 'non dosé'): under 3 g/l with none added. Less sugar = drier, more mineral, more exposed.", "Brut, extra brut, brut nature|Dosage"], ["Carbonic maceration", "Whole, uncrushed red grapes are sealed in a tank full of CO2 and start fermenting inside each berry. The result is juicy, bright, low-tannin red with bubblegum, cherry or banana notes - the classic Beaujolais and 'glou-glou' style.", "Carbonic"], ["Cru, clos, lieu-dit", "A cru is a vineyard or village officially recognised for quality (premier cru, grand cru). A clos is a walled vineyard. A lieu-dit is a traditionally named plot, whether or not it is classified. On a label, any of these means the wine comes from one specific piece of ground.", "Cru, clos, lieu-dit"], ["Cuvée", "A particular named wine or blend from a producer. Not a legal quality term.", "Cuvée"], ["GG, Grosse Lage and the VDP", "The VDP is an association of top German estates with its own site ladder: Gutswein (estate), Ortswein (village), Erste Lage (premier-cru level) and Grosse Lage (grand-cru level, 'GL'). A GG (Grosses Gewächs) is the dry wine from a Grosse Lage site - the estate's top dry Riesling.", "Grosses Gewächs and the VDP tiers|GG|Grosse Lage"], ["Kabinett, Spätlese", "German ripeness levels at harvest. Kabinett is the lightest - delicate, lower alcohol, often a little sweet unless labelled trocken. Spätlese means 'late harvest': riper and fuller. Neither word alone tells you whether the wine is dry.", "Kabinett, Spätlese, Auslese"], ["Trocken, feinherb", "Trocken means dry. Feinherb is an unofficial term for off-dry: a touch of sweetness balanced by high acidity, so it usually tastes less sweet than it measures.", "Trocken and feinherb"], ["Fuder", "The traditional 1,000-litre oval oak cask of the Mosel. Old Fuder add almost no oak flavour; they let the wine breathe slowly.", "Fuder"], ["Sekt / Winzersekt", "German sparkling wine. Winzersekt is grower sparkling: made from the estate's own grapes by the traditional (bottle-fermented) method.", "Sekt"], ["Traditional method", "The Champagne method: the second fermentation happens inside the bottle you buy, and the wine then rests on the spent yeast (lees) for months or years. That time on lees builds the bready, nutty, toasty notes and the fine bubble.", "Traditional method"], ["Reserve wine and solera", "Reserve wine is older wine kept back and blended into a new release for depth and consistency. A solera is a perpetual blend: each year part is bottled and topped up with new wine, so every bottle contains traces of many vintages.", "Reserve wine and solera"], ["Lees, sur lie", "Lees are the spent yeast that settles after fermentation. Ageing 'on the lees' (sur lie) adds texture, creaminess and savoury, bready depth. Stirring them back up is called bâtonnage.", "Lees, sur lie, bâtonnage"], ["Old vines", "No legal definition, but usually 40+ years. Old vines give fewer, smaller bunches, so the wine tends to be more concentrated, and deep roots help them handle drought.", "Old vines"], ["Massal selection", "Replanting a vineyard with cuttings taken from its own best old vines, instead of buying identical nursery clones. It keeps a site's genetic diversity - and, growers argue, its character.", "Massal selection"], ["Yields and hectares", "A hectare is about 2.5 acres. Yield is how much wine a vineyard gives, in hectolitres per hectare (hl/ha). Lower yields usually mean more concentrated fruit.", "Yields and hectares"], ["Organic, biodynamic", "Organic: no synthetic pesticides, herbicides or fertilisers. Biodynamic: organic plus herbal preparations and a farming calendar that treats the farm as one living system. Demeter and Biodyvin certify biodynamics; Ecocert certifies organic in France.", "Organic, biodynamic, Demeter|Organic"], ["Wild yeast", "Fermentation started by the yeasts already living on the grapes and in the cellar, rather than a packaged yeast. Slower and less predictable, but more particular to the place. Also called native or ambient yeast.", "Wild yeast|Ambient yeast"], ["Whole cluster", "Fermenting red grapes with their stems still attached. Adds spice, freshness, a floral or herbal lift and finer-grained tannin.", "Whole cluster"], ["Infusion", "A gentle way to make red wine: the skins are left to steep in the juice with little or no punching-down or pumping-over, so colour and flavour come out slowly and tannin stays soft.", "Infusion"], ["Skin contact", "White grapes fermented on their skins like a red wine - from hours to months. The skins add colour (gold to amber, hence 'orange wine'), tannin, grip and savoury, tea-like flavours.", "Skin contact|Skin contact marker|Amphora and qvevri"], ["Serve chilled", "Light, low-tannin reds show brighter fruit and freshness at cellar temperature (about 55°F). Give them 20-30 minutes in the fridge before pouring.", "Chilled red marker"], ["Sulphur (SO2), zero-zero", "Sulphur dioxide protects wine from oxygen and stray microbes; almost every winemaker adds some. 'Low sulphur' usually means a small dose at bottling. 'Zero-zero' means nothing added and nothing taken away - no added sulphur, no fining, no filtering. These wines need cool, careful storage.", "Sulphur (SO2)|Zero-zero and low sulphur"], ["Unfiltered, unfined", "Filtering strains a wine clear; fining uses a binding agent to pull out haze or harsh tannin. Skipping both keeps more texture and flavour but can leave a little haze or sediment - which is harmless.", "Unfiltered and unfined"], ["Oxidative, reductive", "Oxidative handling lets the wine meet air on purpose, giving nutty, dried-fruit, deeper-coloured wines. Reductive handling keeps air out, giving fresher, fruitier and sometimes flinty or struck-match wines.", "Oxidative and reductive"], ["Slate", "Thin, layered rock that holds the day's heat and forces roots deep into its cracks. The classic Mosel and Saar soil; its wines are often described as taut, smoky and stony.", "Slate"], ["Limestone, chalk, marl", "Limestone and chalk are porous: they hold water in dry years and drain in wet ones, and their wines tend toward high acidity, tension and a chalky, saline finish. Marl is limestone mixed with clay, which adds weight. Kimmeridgian is a fossil-rich limestone found in Chablis and the Côte des Bar.", "Limestone, chalk and marl|Limestone"], ["Granite", "Hard rock that weathers into coarse sand and holds little water. Vines root deep; wines tend to be fresh, fine-grained and aromatic rather than heavy. Found under northern Rhône Syrah (Saint-Joseph, Cornas) and Beaujolais crus.", "Granite|Granitic soils"], ["Volcanic soils", "Soils from lava and ash - often rich in minerals and fast-draining. Wines are commonly described as smoky, saline or savoury.", "Volcanic soils"], ["Loess", "Fine, wind-blown silt that is deep and easy to root in. It gives rounder, softer, fruitier wines than rocky soils.", "Loess"], ["Tuffeau", "The soft, pale limestone of Saumur and Touraine in the Loire. The caves quarried from it became the region's cellars.", "Tuffeau"], ["Vintage and NV", "The vintage is the harvest year on the label. NV (non-vintage) wines, common in Champagne, blend several years for a consistent house style.", "Vintage"]];

function setupSheet() {
  var ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  function tab(name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

  // Bottles
  var b = tab(CFG.BOTTLES);
  if (!b.getRange(1, 1).getValue()) {
    b.getRange(1, 1, 1, BOTTLE_HEADERS.length).setValues([BOTTLE_HEADERS]).setFontWeight('bold').setWrap(true);
    b.setFrozenRows(1);
    var ex = []; for (var k = 0; k < BOTTLE_HEADERS.length; k++) ex.push('');
    var put = function (prefix, val) { for (var c = 0; c < BOTTLE_HEADERS.length; c++) if (BOTTLE_HEADERS[c].indexOf(prefix) === 0) { ex[c] = val; return; } };
    put('menu / off-menu', 'menu'); put('Menu section', 'White'); put('Sort order', 1); put('Producer', 'Example Producer');
    put('Wine / cuv', 'Example Cuvée'); put('Grapes', 'Chenin Blanc'); put('Place', 'Example Town, Example Region'); put('Vintage', '2023');
    put('Price', 60); put('Taste group', 'Crisp & lean whites'); put('REMEMBER IT BY', 'Replace this example row with your first real bottle.');
    put('TASTING', 'One short tasting note per line.\nSecond line.'); put('ID (do not edit)', 'b1');
    put('MENU: producer', 'Example Producer'); put('MENU: cuv', 'Example Cuvée'); put('MENU: grape', 'Chenin Blanc'); put('MENU: place', 'Example Town');
    b.getRange(2, 1, 1, ex.length).setValues([ex]);
    var col = function (prefix) { for (var c = 0; c < BOTTLE_HEADERS.length; c++) if (BOTTLE_HEADERS[c].indexOf(prefix) === 0) return c + 1; return 0; };
    var list = function (c, vals) { if (c) b.getRange(2, c, 998, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(false).build()); };
    list(col('menu / off-menu'), ['menu', 'off-menu', 'insert', 'needs-page', 'sold-out', 'removed']);
    list(col('Taste group'), TASTE_GROUPS);
    ['Serve chilled', 'Skin contact', 'Out of stock', 'Slightly reductive'].forEach(function (p) { list(col(p), ['Y']); });
  }

  // Setup
  var st = tab(CFG.SETUP);
  if (!st.getRange(1, 1).getValue()) st.getRange(1, 1, 5, 3).setValues([
    ['Setting', 'Value', 'What it is'],
    ['Bar name', 'Your Bar', 'Printed on the PDF cover if there is no cover logo.'],
    ['Page template ID', '', 'ID of a Google Slides file with page size 11 x 8.5 in (File > Page setup > Custom). Needed for Make folded PDF. See README step 5.'],
    ['Cover logo file ID', '', 'Optional: ID of your logo image in Google Drive, for the PDF cover.'],
    ['Label photo folder ID', '', 'Optional: Drive folder for label photos taken in the app. Blank = the same folder as this sheet.']
  ]).setWrap(true);

  // Staff app (PIN + staff codes)
  var sa = tab(STAFF.SETTINGS);
  if (!sa.getRange(1, 1).getValue()) sa.getRange(1, 1, 14, 4).setValues([
    ['App PIN', "'0000", "Opens the staff app (asked once per phone). Also works as a backup 86'd code – the person types their name.", ''],
    ['', '', '', ''],
    ['STAFF CODES', '', '', ''],
    ['Name', "86'd code", 'Add / edit code', 'Print code'],
    ['Manager (change me)', 'MGR-86', 'MGR-EDIT', 'MGR-PRINT'],
    ['', '', '', ''], ['', '', '', ''], ['', '', '', ''],
    ['HOW IT WORKS', '', '', ''],
    ['Codes', '', 'Each person uses their own codes, so the app logs them by name. Codes work in upper or lower case. Leave a code blank to take that permission away.', ''],
    ['Add someone', '', 'Type their name and codes on an empty row in the table above. Works right away.', ''],
    ['Remove someone', '', 'Delete their row (or clear their codes). Their codes stop working right away.', ''],
    ['Logs', '', "Every 86 / back in stock goes on the \"86 log\" tab; every add, edit, undo and print on the \"Bottle log\" tab. The app's Change log shows both.", ''],
    ['Keep private', '', 'Codes live only on this tab. Never put them in the app or its files.', '']
  ]).setWrap(true);

  // Menu layout
  var ml = tab(CFG.LAYOUT);
  if (!ml.getRange(1, 1).getValue()) ml.getRange(1, 1, 7, 2).setValues([
    ['Page', 'Section (must match "Menu section" on the Bottles tab exactly)'],
    [2, 'Sparkling'], [3, 'White'], [4, 'Orange & Pink'], [5, 'Red'], ['', ''],
    ['FOOTNOTE', '**skin contact or chilled red']
  ]);

  // Glossary
  var g = tab('Glossary');
  if (!g.getRange(1, 1).getValue()) {
    g.getRange(1, 1, 1, 3).setValues([['Term', 'Explanation', 'Other names (separate with |)']]).setFontWeight('bold');
    g.getRange(2, 1, STARTER_GLOSSARY.length, 3).setValues(STARTER_GLOSSARY).setWrap(true);
  }

  // Logs
  var l86 = tab(STAFF.LOG);
  if (!l86.getRange(1, 1).getValue()) { l86.getRange(1, 1, 1, 6).setValues([['When', 'Bottle', 'Action', 'Name', 'Bottle ID', 'Undone']]); l86.setFrozenRows(1); }
  var lb = tab(BOT.LOG);
  if (!lb.getRange(1, 1).getValue()) { lb.getRange(1, 1, 1, 7).setValues([['When', 'Bottle', 'Action', 'Name', 'ID', "Undo data (for the app – don't edit)", 'Undone']]); lb.setFrozenRows(1); }

  try { SpreadsheetApp.getUi().alert('Setup done', 'Your tabs are ready. Next: fill in the Setup and Staff app tabs, then publish the web app (README step 4).', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
}
