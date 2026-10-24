// ===== متابعة الدائنين والمدينين — Google Apps Script =====
// ضع رقم الشيت هنا (الجزء بين /d/ و /edit في الرابط)
const SHEET_ID = '1zuE_DMx82fnqsKQyN22sqdhfob-mrTGTMkVt_JXrKUE';
const CREDITOR = 'أبو سعود';   // اسم الدائن الذي يظهر في التطبيق
const CHUNK = 40000;

function ws_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function readData_() {
  const sh = ws_('_data'), n = sh.getLastRow();
  const t = n ? sh.getRange(1, 1, n, 1).getValues().map(r => r[0]).join('') : '';
  try { const j = JSON.parse(t); return { debts: j.debts || [], phones: j.phones || {} }; }
  catch (e) { return { debts: [], phones: {} }; }
}

// حفظ النسخة الكاملة + جدولان مقروءان (الديون / الدفعات)
function writeAll_(D) {
  const raw = JSON.stringify(D), sh = ws_('_data');
  sh.clear();
  const parts = [];
  for (let i = 0; i < raw.length; i += CHUNK) parts.push([raw.slice(i, i + CHUNK)]);
  sh.getRange(1, 1, parts.length, 1).setValues(parts);

  const paid = d => d.pays.reduce((a, p) => a + +p.a, 0);
  const rows = [['الدائن', 'المدين', 'جوال المدين', 'المبلغ', 'المسدد', 'المتبقي', 'تاريخ الدين', 'الاستحقاق', 'ملاحظات']]
    .concat(D.debts.map(d => [d.cr, d.db, (D.phones || {})[d.db] || '', +d.amt, paid(d), d.amt - paid(d), d.date, d.due, d.note]));
  const v = ws_('الديون'); v.clear();
  v.getRange(1, 1, rows.length, 9).setValues(rows);
  v.setRightToLeft(true); v.setFrozenRows(1);

  const pr = [['المدين', 'الدائن', 'المبلغ', 'التاريخ', 'ملاحظة']];
  D.debts.forEach(d => d.pays.forEach(p => pr.push([d.db, d.cr, +p.a, p.d, p.n])));
  const p = ws_('الدفعات'); p.clear();
  p.getRange(1, 1, pr.length, 5).setValues(pr);
  p.setRightToLeft(true); p.setFrozenRows(1);
}

function doGet() {
  const sh = ws_('_data'), n = sh.getLastRow();
  const txt = n ? sh.getRange(1, 1, n, 1).getValues().map(r => r[0]).join('') : '{}';
  return ContentService.createTextOutput(txt || '{}').setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  writeAll_(JSON.parse(e.postData.contents));
  return ContentService.createTextOutput('ok');
}

// ===== استيراد تلقائي من تبويبات "سلف ..." الموجودة في الشيت =====
// شغّل هذه الدالة يدوياً مرة واحدة (Run ▶ importFromTabs)، ثم من التطبيق: الإعدادات ← تحميل من الشيت
// تقرأ كل تبويب فيه جدول (التاريخ / القسط) وخانة "المتبقي"، وتتخطى الباقي وتكتب السبب في ورقة _سجل_الاستيراد
function num_(x) { return parseFloat(String(x).replace(/[^\d.\-]/g, '')); }

function iso_(s) {
  s = String(s || '').trim();
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);   // يوم/شهر/سنة
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function importFromTabs() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const skip = ['_data', 'الديون', 'الدفعات', '_سجل_الاستيراد'];
  const D = readData_();
  D.debts = D.debts.filter(d => !String(d.id).startsWith('tab_'));
  const log = [['التبويب', 'النتيجة']];

  ss.getSheets().forEach(sh => {
    const nm = sh.getName();
    if (skip.includes(nm)) return;
    const v = sh.getDisplayValues();
    let hr = -1, cD = -1, cT = -1;
    for (let r = 0; r < v.length && hr < 0; r++) {
      const t = v[r].map(x => String(x).trim());
      const a = t.indexOf('القسط'), b = t.indexOf('التاريخ');
      if (a >= 0 && b >= 0) { hr = r; cD = a; cT = b; }
    }
    if (hr < 0) { log.push([nm, 'تم التخطي: لا يوجد جدول أقساط']); return; }

    let rem = NaN, remR = -1, remC = -1;
    for (let r = 0; r < hr && isNaN(rem); r++)
      for (let c = 1; c < v[r].length; c++)
        if (String(v[r][c]).trim() === 'المتبقي') { rem = num_(v[r][c - 1]); remR = r; remC = c; break; }
    if (isNaN(rem)) { log.push([nm, 'تم التخطي: لم أجد خانة المتبقي']); return; }

    const pays = [];
    for (let r = hr + 1; r < v.length; r++) {
      const a = num_(v[r][cD]);
      if (!(a > 0)) continue;
      const n0 = String(v[r][cD + 2] || '').trim();
      pays.push({ a: a, d: iso_(v[r][cT]), n: /^[\d.,\s\-]*$/.test(n0) ? '' : n0 });
    }
    const paid = pays.reduce((s, p) => s + p.a, 0), amt = paid + rem;

    let title = '';
    for (let r = 0; r < hr; r++) v[r].forEach(x => {
      x = String(x).trim();
      if (/[^\d.,\s]/.test(x) && x !== 'المتبقي' && x.length > title.length) title = x;
    });
    let det = [];
    for (let r = 0; r < v.length; r++) {
      const c = v[r].findIndex(x => String(x).trim() === 'البيان');
      if (c >= 0) {
        for (let q = r + 1; q < v.length && String(v[q][c]).trim(); q++)
          det.push(String(v[q][c]).trim() + ' ' + String(v[q][c - 1]).trim());
        break;
      }
    }
    const note = title + (det.length ? ' | التفصيل: ' + det.join('، ') : '');

    let warn = '';
    if (remR > 0) {
      const tot = num_(v[remR - 1][remC]);
      if (isFinite(tot) && tot !== amt) warn = ' ⚠ الإجمالي في الشيت ' + tot + ' يختلف عن المحسوب ' + amt;
    }
    const name = nm.replace(/^سلف\s+/, '').trim();
    D.debts.push({ id: 'tab_' + sh.getSheetId(), cr: CREDITOR, db: name, amt: amt, date: '', due: '', note: note, pays: pays });
    log.push([nm, 'تم: ' + name + ' | المبلغ ' + amt + ' | مسدد ' + paid + ' | متبقي ' + rem + ' | دفعات ' + pays.length + warn]);
  });

  writeAll_(D);
  const L = ws_('_سجل_الاستيراد'); L.clear();
  L.getRange(1, 1, log.length, 2).setValues(log);
  L.setRightToLeft(true);
}
