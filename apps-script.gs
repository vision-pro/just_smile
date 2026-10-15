// رؤية للإنتاج الفني — استقبال ردود صفحات العروض في Google Sheets
// عام لكل زبون: أسماء الأعمدة تجي من الصفحة نفسها (_labels)، والمفتاح محفوظ كملاحظة على خلية العنوان
const NOTIFY_EMAIL = "v.production.iq@gmail.com"; // "" يوقف إيميل التنبيه
const SHEET_ID = "1d2TrJ_2wU1Pz2l-goDZcAEg5q1KCH0-_HCocUMA56xM"; // مشروع مستقل؛ "" إذا السكربت مربوط بالشيت نفسه
const ACCENT = "#F37420";

function doGet() {
  return ContentService.createTextOutput("ok");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const p = (e && e.parameter) || {};
    let labels = {};
    try { labels = JSON.parse(p._labels || "{}"); } catch (err) {}
    labels.submitted_at = labels.submitted_at || "وقت الإرسال";
    const incoming = ["submitted_at"].concat(Object.keys(labels), Object.keys(p))
      .filter((k, i, a) => k.charAt(0) !== "_" && a.indexOf(k) === i);
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheets()[0];
    const hasHead = sh.getLastRow() > 0;
    const head = hasHead ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    const keys = hasHead ? sh.getRange(1, 1, 1, sh.getLastColumn()).getNotes()[0] : [];
    const before = keys.length;
    incoming.forEach(k => { if (keys.indexOf(k) === -1) keys.push(k); });
    if (keys.length !== before) {
      sh.getRange(1, 1, 1, keys.length)
        .setValues([keys.map((k, i) => labels[k] || head[i] || k)])
        .setNotes([keys])
        .setFontWeight("bold").setBackground(ACCENT).setFontColor("#FFFFFF");
      sh.setFrozenRows(1);
      sh.setRightToLeft(true);
    }
    // أرقام الهواتف تبقى نص (ما يضيع الصفر)، وأي قيمة تبدأ بـ = + - @ ما تنحسب معادلة
    sh.appendRow(keys.map(k => { const v = String(p[k] || ""); return /^[=+\-@]|^\d[\d\s]*$/.test(v) ? "'" + v : v; }));
    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, p._title || "رد جديد",
        keys.filter(String).map(k => (labels[k] || k) + ":\n" + (p[k] || "—")).join("\n\n"));
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
