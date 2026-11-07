/**
 * طلب كود الفروع — الجزء اللي يشتغل داخل قوقل (Apps Script)
 *
 *  - الأكواد  : من الشيت، ورقة «الأكواد»  (A الفرع | B الكود) — كود واحد ثابت لكل فرع لين تغيّره
 *  - الأسئلة  : من فورم قوقل (FORM_URL) — أي تعديل في الفورم يطلع في الصفحة خلال دقيقة
 *  - الإجابات : تنحفظ في الشيت، ورقة «الردود»  (الوقت | الفرع | الكود | عمود لكل سؤال)
 *
 * الصفحة (index.html) تكلّم هذا السكربت من برّا عن طريق رابط الـ Web app.
 */

// ===== الإعدادات =====

// رقم الشيت (من رابطه: /spreadsheets/d/<هذا>/edit).
// إذا السكربت مفتوح من داخل الشيت (Extensions ← Apps Script) تقدر تخليه فاضي.
const SHEET_ID = '';

// رابط فورم قوقل اللي ناخذ منه الأسئلة (رابط التعديل: .../forms/d/<id>/edit).
// لازم الحساب اللي عليه السكربت يكون «محرّر» في الفورم.
// تقدر تخليه فاضي وتحط الرابط من قائمة «⚙️ تجهيز» داخل الشيت.
// إذا ما فيه رابط أبداً، التجهيز يسوي فورم تجريبي ويستخدمه.
const FORM_URL = '';

// الصفحة العامة (GitHub Pages). رابطك للناس = هذي الصفحة + ?api=<رابط النشر حقك>
const PAGE_URL = 'https://24hmood24.github.io/branch-codes/';

const CODES = 'الأكواد';
const RESP  = 'الردود';

// ===== أدوات =====

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
}

function formUrl_() {
  const p = PropertiesService.getScriptProperties();
  return FORM_URL || p.getProperty('FORM_URL') || p.getProperty('DEMO_FORM_URL') || '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== الـ API =====

// GET ?action=start            → عنوان الفورم + أسماء الفروع + الأسئلة (بدون أي كود)
// GET ?action=code&branch=...  → كود الفرع المختار فقط
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'start') return json_(Object.assign({ ok: true }, getStart()));
    if (p.action === 'code')  return json_({ ok: true, code: getCode(p.branch) });
    return json_({ ok: true, msg: 'Branch codes API' });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

// POST {action:'submit', branch, answers:{id: جواب}} → يحفظ في ورقة الردود
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'submit') {
      submit(body.branch, body.answers);
      return json_({ ok: true });
    }
    throw new Error('طلب غير معروف');
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

// ===== الفروع والأكواد =====

function rows_() {
  const sh = ss_().getSheetByName(CODES);
  if (!sh) throw new Error('ورقة «' + CODES + '» غير موجودة — شغّل setup');
  return sh.getDataRange().getValues().slice(1)
    .map(r => [String(r[0]).trim(), String(r[1]).trim()])
    .filter(r => r[0]);
}

function getCode(branch) {
  const row = rows_().find(r => r[0] === String(branch || '').trim());
  if (!row || !row[1]) throw new Error('ما فيه كود لهذا الفرع حالياً');
  return row[1];
}

// ===== الأسئلة من فورم قوقل =====

// يقرأ الفورم ويحوّل كل سؤال لشكل بسيط تفهمه الصفحة.
// يتخزّن دقيقة وحدة عشان الصفحة تفتح بسرعة.
function readForm_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('form');
  if (hit) return JSON.parse(hit);

  const url = formUrl_();
  if (!url) throw new Error('ما فيه فورم — حط FORM_URL أو شغّل setup');
  const form = FormApp.openByUrl(url);
  const T = FormApp.ItemType;
  const qs = [];

  form.getItems().forEach(item => {
    const id = String(item.getId());
    const title = item.getTitle();
    const help = item.getHelpText();
    let q = null;
    switch (item.getType()) {
      case T.TEXT: {
        const it = item.asTextItem();
        const phone = /جوال|هاتف|موبايل|phone|mobile/i.test(title);
        const mail = /بريد|ايميل|إيميل|email/i.test(title);
        q = { type: phone ? 'جوال' : mail ? 'بريد' : 'نص', required: it.isRequired() };
        break;
      }
      case T.PARAGRAPH_TEXT:
        q = { type: 'فقرة', required: item.asParagraphTextItem().isRequired() };
        break;
      case T.MULTIPLE_CHOICE: {
        const it = item.asMultipleChoiceItem();
        q = { type: 'اختيار', required: it.isRequired(), options: it.getChoices().map(c => c.getValue()) };
        break;
      }
      case T.LIST: {
        const it = item.asListItem();
        q = { type: 'اختيار', required: it.isRequired(), options: it.getChoices().map(c => c.getValue()) };
        break;
      }
      case T.CHECKBOX: {
        const it = item.asCheckboxItem();
        q = { type: 'متعدد', required: it.isRequired(), options: it.getChoices().map(c => c.getValue()) };
        break;
      }
      case T.SCALE: {
        const it = item.asScaleItem();
        const opts = [];
        for (let i = it.getLowerBound(); i <= it.getUpperBound(); i++) opts.push(String(i));
        q = { type: 'اختيار', required: it.isRequired(), options: opts };
        break;
      }
      case T.DATE:
        q = { type: 'تاريخ', required: item.asDateItem().isRequired() };
        break;
      case T.TIME:
        q = { type: 'وقت', required: item.asTimeItem().isRequired() };
        break;
      case T.SECTION_HEADER:
      case T.PAGE_BREAK:
        q = { type: 'عنوان', required: false };
        break;
      default:
        q = null; // صور/فيديو/جداول: ما ندعمها
    }
    if (q) {
      q.id = id; q.q = title; q.help = help || '';
      q.options = q.options || [];
      qs.push(q);
    }
  });

  const data = { title: form.getTitle(), description: form.getDescription(), questions: qs };
  cache.put('form', JSON.stringify(data), 60);
  return data;
}

function getStart() {
  const f = readForm_();
  return {
    title: f.title,
    description: f.description,
    branches: rows_().map(r => r[0]),
    questions: f.questions
  };
}

// ===== حفظ الإجابات =====

function submit(branch, answers) {
  const row = rows_().find(r => r[0] === String(branch || '').trim());
  if (!row) throw new Error('الفرع غير موجود');
  answers = answers || {};
  const qs = readForm_().questions.filter(q => q.type !== 'عنوان');

  qs.forEach(q => {
    if (q.required && !String(answers[q.id] || '').trim()) throw new Error('عبّ: ' + q.q);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    const sh = ss.getSheetByName(RESP) || ss.insertSheet(RESP);
    let head = sh.getLastRow() > 0 && sh.getLastColumn() > 0
      ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String)
      : [];
    while (head.length && !head[head.length - 1]) head.pop();
    if (!head.length) head = ['الوقت', 'الفرع', 'الكود'];
    qs.forEach(q => { if (head.indexOf(q.q) < 0) head.push(q.q); }); // سؤال جديد = عمود جديد
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#ede7f6');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);

    const byTitle = {};
    qs.forEach(q => { byTitle[q.q] = answers[q.id]; });
    const vals = head.map(h => {
      if (h === 'الوقت') return new Date();
      if (h === 'الفرع') return row[0];
      if (h === 'الكود') return row[1];
      let v = String(byTitle[h] == null ? '' : byTitle[h]).slice(0, 2000);
      if (/^[=+\-@0]/.test(v)) v = "'" + v; // يحفظ الصفر في أول الجوال ويمنع المعادلات
      return v;
    });
    sh.appendRow(vals);
  } finally {
    lock.releaseLock();
  }
}

// ===== التجهيز (شغّلها مرة وحدة من المحرر) =====
// تسوي ورقتين: «الأكواد» و«الردود»، وتحط فروع تجربة إذا الورقة فاضية.
// وإذا FORM_URL فاضي تسوي فورم تجريبي وتطبع رابطه في السجل (Execution log).
function setup() {
  const ss = ss_();
  const purple = '#ede7f6';

  const c = ss.getSheetByName(CODES) || ss.insertSheet(CODES, 0);
  c.setRightToLeft(true);
  c.getRange('A1:B1').setValues([['الفرع', 'الكود']]).setFontWeight('bold').setBackground(purple);
  if (c.getLastRow() < 2) {
    c.getRange('A2:B4').setValues([
      ['فرع تجربة 1', 'TEST-111'],
      ['فرع تجربة 2', 'TEST-222'],
      ['فرع تجربة 3', 'TEST-333']
    ]);
  }
  c.setFrozenRows(1);
  c.setColumnWidths(1, 2, 180);

  const r = ss.getSheetByName(RESP) || ss.insertSheet(RESP, 1);
  r.setRightToLeft(true);
  if (r.getLastRow() < 1) {
    r.getRange('A1:C1').setValues([['الوقت', 'الفرع', 'الكود']]).setFontWeight('bold').setBackground(purple);
  }
  r.setFrozenRows(1);

  if (!formUrl_()) {
    const f = FormApp.create('نموذج تجريبي — طلب كود');
    f.setTitle('طلب كود');
    f.setDescription('اختر الفرع واطلب الكود، بعدين عبّ معلوماتك وأرسلها.');
    f.addTextItem().setTitle('الاسم الكامل').setRequired(true);
    f.addTextItem().setTitle('رقم الجوال').setRequired(true);
    f.addTextItem().setTitle('رقم الفاتورة');
    f.addMultipleChoiceItem().setTitle('نوع الطلب').setChoiceValues(['شراء', 'استرجاع', 'استبدال']).setRequired(true);
    f.addParagraphTextItem().setTitle('ملاحظات');
    PropertiesService.getScriptProperties().setProperty('DEMO_FORM_URL', f.getEditUrl());
  }
  CacheService.getScriptCache().remove('form');
  Logger.log('الفورم المستخدم: ' + formUrl_());
}

// ===== قائمة «⚙️ تجهيز» داخل الشيت =====

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ تجهيز')
    .addItem('١) تجهيز وربط الفورم', 'menuSetup')
    .addItem('٢) طريقة النشر', 'menuDeployHelp')
    .addItem('٣) رابط الصفحة للناس', 'menuPageLink')
    .addToUi();
}

function menuSetup() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const now = formUrl_();
  const res = ui.prompt('ربط فورم قوقل',
    'الصق رابط فورم قوقل حقك (رابط التعديل اللي ينتهي بـ /edit).\n' +
    'إذا خليته فاضي نسوي لك فورم تجريبي.' + (now ? '\n\nالفورم الحالي:\n' + now : ''),
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const url = res.getResponseText().trim();
  if (url) {
    if (!/docs\.google\.com\/forms\/d\//.test(url) || /\/forms\/d\/e\//.test(url)) {
      ui.alert('الرابط غير صحيح', 'افتح الفورم للتعديل وانسخ الرابط من شريط العنوان.\nلازم يكون فيه /forms/d/ ومو رابط المشاركة (/forms/d/e/).', ui.ButtonSet.OK);
      return;
    }
    try { FormApp.openByUrl(url).getTitle(); } catch (e) {
      ui.alert('ما قدرت أفتح الفورم', 'تأكد إن حسابك هذا «محرّر» في الفورم.\n\n' + e.message, ui.ButtonSet.OK);
      return;
    }
    props.setProperty('FORM_URL', url);
  }
  setup();
  const f = readForm_();
  ui.alert('تم التجهيز ✓',
    'الشيت فيه الحين: «الأكواد» و«الردود».\n' +
    'الفورم: ' + (f.title || formUrl_()) + '\n' +
    'عدد الأسئلة: ' + f.questions.filter(q => q.type !== 'عنوان').length + '\n\n' +
    'الخطوة الجاية: ⚙️ تجهيز ← ٢) طريقة النشر', ui.ButtonSet.OK);
}

function menuDeployHelp() {
  const edit = 'https://script.google.com/home/projects/' + ScriptApp.getScriptId() + '/edit';
  const html = HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:14px;line-height:1.9">' +
    '<ol style="padding-right:18px;margin:0">' +
    '<li>افتح المحرّر: <a href="' + edit + '" target="_blank">اضغط هنا</a> (أو Extensions ← Apps Script)</li>' +
    '<li>فوق يمين: <b>Deploy ← New deployment</b></li>' +
    '<li>اضغط ⚙️ جنب «Select type» واختر <b>Web app</b></li>' +
    '<li><b>Execute as:</b> Me — <b>Who has access:</b> Anyone</li>' +
    '<li><b>Deploy</b> ← وافق على الصلاحيات إذا طلب ← <b>Done</b></li>' +
    '<li>ارجع للشيت: <b>⚙️ تجهيز ← ٣) رابط الصفحة للناس</b></li>' +
    '</ol>' +
    '<p style="color:#5f6368;font-size:12px">إذا المحرّر طلع «الصفحة غير موجودة»: المتصفح فيه أكثر من حساب قوقل — افتحه في نافذة خاصة (Incognito).</p>' +
    '</div>').setWidth(460).setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(html, 'طريقة النشر');
}

function menuPageLink() {
  const ui = SpreadsheetApp.getUi();
  let api = '';
  try { api = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  if (!/\/exec$/.test(api)) {
    const res = ui.prompt('رابط النشر',
      'ما لقيت رابط النشر تلقائياً.\nالصق رابط الـ Web app (ينتهي بـ /exec) من Deploy ← Manage deployments:',
      ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    api = res.getResponseText().trim();
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(api)) {
      ui.alert('الرابط لازم يبدأ بـ https://script.google.com وينتهي بـ /exec');
      return;
    }
  }
  const link = PAGE_URL + '?api=' + encodeURIComponent(api);
  const html = HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:14px">' +
    '<p>هذا رابطك، أرسله للناس:</p>' +
    '<textarea readonly onclick="this.select()" style="width:100%;height:90px;direction:ltr;font-size:12px">' + link + '</textarea>' +
    '<p><a href="' + link + '" target="_blank">افتح الصفحة وجرّبها</a></p>' +
    '</div>').setWidth(460).setHeight(260);
  ui.showModalDialog(html, 'رابط الصفحة للناس');
}
