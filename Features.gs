/**
 * Features.gs - dashboard, status, history and notifications.
 * In Apps Script: click + next to "Files" -> Script -> name it Features -> paste this whole file.
 * Then make the 3 small edits to Code.gs listed in README.md and redeploy as a NEW VERSION.
 */
const NOTIFY_EMAILS      = "";   // stock team e-mails, comma separated, e.g. "a@gmail.com,b@gmail.com" (leave "" for no e-mail)
const TELEGRAM_BOT_TOKEN = "";   // optional: Telegram bot token from @BotFather
const TELEGRAM_CHAT_ID   = "";   // optional: chat / group id that receives the message
const SITE_URL           = "";   // optional: link to the form, added to the message
const STOCK_ROLE         = /^(stock|admin)$/i;               // Users tab, column "Role"
const STATUSES           = ["Pending", "Approved", "Issued", "Completed"];
const EXTRA_HEADERS      = ["Status", "Status updated", "Updated by"];

// ---- Main Stock approval & feedback ----
// Telegram chat id of each department (the department's group, or a person who pressed Start on the bot).
// Department name must match the department in the Users tab. Leave {} to notify inside the system only.
const DEPT_TELEGRAM = {};   // e.g. { "Pharmacy": "-1001234567890", "Laboratory": "123456789" }
const DEPT_EMAILS   = {};   // optional, e.g. { "Pharmacy": "pharmacy@gmail.com" }
const APPROVAL_TAB  = "Approvals";   // one row per item per approval (history, never edited)
const FEEDBACK_TAB  = "Feedback";    // one row per approval submitted (remarks + whether the department has seen it)
const AP_HEADERS = ["Time", "Ref", "Version", "Line", "Item code", "Item name", "Requested qty", "Approved qty", "Unit", "Reason"];
const FB_HEADERS = ["Time", "Ref", "Version", "Reviewed by", "Department", "Remarks", "Total requested", "Total approved", "Seen", "Seen time", "Seen by"];

function roleOf(user) {
  const ss = USERS_SHEET_ID ? SpreadsheetApp.openById(USERS_SHEET_ID) : openMain();
  const v = ss.getSheetByName(USERS_TAB).getDataRange().getValues();
  const h = v[0].map(x => String(x).trim().toLowerCase());
  const iu = h.indexOf("username"), ir = h.indexOf("role");
  if (ir < 0) return "";
  const row = v.slice(1).filter(r => String(r[iu]).trim().toLowerCase() === String(user).trim().toLowerCase())[0];
  return row ? String(row[ir]).trim() : "";
}
const isStock = me => STOCK_ROLE.test(roleOf(me.user));

function ensureCols(sh) {
  let cols = columnMap(sh);
  EXTRA_HEADERS.forEach(h => {
    if (!cols[h]) { sh.getRange(1, Math.max(sh.getLastColumn(), 1) + 1).setValue(h).setFontWeight("bold"); cols = columnMap(sh); }
  });
  return cols;
}

function featureRoute(d, me) {
  if (d.action === "list") return listRequests(me);
  if (d.action === "setStatus") return setStatus(d, me);
  if (d.action === "approve") return approveRequest(d, me);
  if (d.action === "notices") return noticesFor(me);
  if (d.action === "seen") return markSeen(d, me);
  return null;
}

// Stock team: all requests. Everyone else: only their own department.
function listRequests(me) {
  const sh = getSheet(), cols = ensureCols(sh), stock = isStock(me), last = sh.getLastRow();
  if (last < 2) return { ok: true, stock: stock, statuses: STATUSES, requests: [] };
  const v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const g = (r, h) => r[cols[h] - 1];
  const apr = readApprovals_();
  const map = {}, order = [];
  v.forEach(r => {
    const ref = String(g(r, "Ref")); if (!ref) return;
    if (!stock && String(g(r, "Department")).trim().toLowerCase() !== me.dept.toLowerCase()) return;
    if (!map[ref]) {
      const t = g(r, "Timestamp");
      map[ref] = { ref: ref, time: t instanceof Date ? Utilities.formatDate(t, TIMEZONE, "yyyy-MM-dd HH:mm") : String(t),
        name: String(g(r, "Name")), dept: String(g(r, "Department")), priority: String(g(r, "Priority")),
        notes: String(g(r, "Notes")), status: String(g(r, "Status") || "Pending"), by: String(g(r, "Updated by")), items: [] };
      order.push(ref);
    }
    map[ref].items.push({ code: String(g(r, "Item code")), name: String(g(r, "Item name")), qty: g(r, "Qty"), unit: String(g(r, "Unit")) });
  });
  return { ok: true, stock: stock, statuses: STATUSES, requests: order.reverse().slice(0, 5000).map(k => {
    const q = map[k];
    q.approvals = apr[k] || [];
    q.unseen = !stock && q.approvals.some(v => !v.seen);
    return q;
  }) };
}

function setStatus(d, me) {
  if (!isStock(me)) return { ok: false, error: "Only the stock team can change the status" };
  if (STATUSES.indexOf(d.status) < 0) return { ok: false, error: "Unknown status" };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet(), cols = ensureCols(sh), last = sh.getLastRow();
    if (last < 2) return { ok: false, error: "Request not found" };
    const refs = sh.getRange(2, cols["Ref"], last - 1, 1).getValues();
    const when = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm");
    let n = 0;
    refs.forEach((r, i) => {
      if (String(r[0]) !== String(d.ref)) return;
      sh.getRange(i + 2, cols["Status"]).setValue(d.status);
      sh.getRange(i + 2, cols["Status updated"]).setValue(when);
      sh.getRange(i + 2, cols["Updated by"]).setValue(me.user);
      n++;
    });
    return n ? { ok: true, status: d.status } : { ok: false, error: "Request not found" };
  } finally { lock.releaseLock(); }
}

// Called by saveRequest after a NEW request is written
function afterSave(sh, start, count, ref, d, dept) {
  try { sh.getRange(start, ensureCols(sh)["Status"], count, 1).setValue("Pending"); } catch (e) { console.error(e); }
  notifyStock(ref, d, dept);
}

function notifyStock(ref, d, dept) {
  const lines = d.items.map(i => "• " + i.name + " (" + i.code + ") x " + i.qty + " " + i.unit).join("\n");
  const msg = "New item request " + ref + "\nDepartment: " + dept + "\nBy: " + d.name + "\nPriority: " + (d.priority || "Normal") +
    "\n\n" + lines + (d.notes ? "\n\nNotes: " + d.notes : "") + (SITE_URL ? "\n\n" + SITE_URL : "");
  try { if (NOTIFY_EMAILS) MailApp.sendEmail(NOTIFY_EMAILS, (d.priority === "Urgent" ? "[URGENT] " : "") + "New request " + ref + " - " + dept, msg); }
  catch (e) { console.error("mail", e); }
  try { if (TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage",
      { method: "post", payload: { chat_id: TELEGRAM_CHAT_ID, text: msg }, muteHttpExceptions: true }); }
  catch (e) { console.error("telegram", e); }
}

/* Run once from the editor: asks for permission and sends a test notification */
function testNotify() {
  notifyStock("REQ-TEST-00000", { name: "Test", priority: "Normal", notes: "", items: [{ code: "0001", name: "Test item", qty: 1, unit: "pcs" }] }, "Test department");
}


/* ===================== Approval & feedback ===================== */
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}
const fmtTime_ = t => t instanceof Date ? Utilities.formatDate(t, TIMEZONE, "yyyy-MM-dd HH:mm") : String(t);
const norm_ = c => String(c).trim().replace(/^0+/, "").toLowerCase();
function lookupDept_(map, dept) {
  const k = Object.keys(map).filter(x => x.trim().toLowerCase() === String(dept).trim().toLowerCase())[0];
  return k ? map[k] : "";
}

// All approvals grouped by request ref: { REF: [ {ver, time, by, remarks, seen, items:[...]}, ... ] }
function readApprovals_() {
  const out = {}, ss = openMain(), fb = ss.getSheetByName(FEEDBACK_TAB), ap = ss.getSheetByName(APPROVAL_TAB);
  if (!fb || fb.getLastRow() < 2) return out;
  const idx = {};
  fb.getRange(2, 1, fb.getLastRow() - 1, FB_HEADERS.length).getValues().forEach(r => {
    const ref = String(r[1]); if (!ref) return;
    const o = { ver: Number(r[2]) || 1, time: fmtTime_(r[0]), by: String(r[3]), dept: String(r[4]), remarks: String(r[5]),
      totReq: Number(r[6]) || 0, totAppr: Number(r[7]) || 0, seen: String(r[8]) === "Y", items: [] };
    (out[ref] = out[ref] || []).push(o);
    idx[ref + "|" + o.ver] = o;
  });
  if (ap && ap.getLastRow() > 1) {
    ap.getRange(2, 1, ap.getLastRow() - 1, AP_HEADERS.length).getValues().forEach(r => {
      const o = idx[String(r[1]) + "|" + (Number(r[2]) || 1)];
      if (o) o.items[(Number(r[3]) || 1) - 1] = { code: String(r[4]), name: String(r[5]), req: Number(r[6]) || 0,
        approved: Number(r[7]) || 0, unit: String(r[8]), reason: String(r[9] || "") };
    });
  }
  Object.keys(out).forEach(k => out[k].sort((a, b) => a.ver - b.ver));
  return out;
}

// Main Stock submits approved quantities. The original request rows are NOT changed (only the Status columns).
function approveRequest(d, me) {
  if (!isStock(me)) return { ok: false, error: "Only the stock team can approve requests" };
  if (!d.ref || !Array.isArray(d.items)) return { ok: false, error: "Missing fields" };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let res, dept = "", lines = [], out = [], remarks = "", ver = 1, by = me.user;
  try {
    const sh = getSheet(), cols = ensureCols(sh), last = sh.getLastRow();
    if (last < 2) return { ok: false, error: "Request not found" };
    const v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const g = (r, h) => r[cols[h] - 1], rowIdx = [];
    let cur = "";
    v.forEach((r, i) => {
      if (String(g(r, "Ref")) !== String(d.ref)) return;
      if (!lines.length) { dept = String(g(r, "Department")); cur = String(g(r, "Status") || "Pending"); }
      lines.push({ code: String(g(r, "Item code")), name: String(g(r, "Item name")), qty: Number(g(r, "Qty")) || 0, unit: String(g(r, "Unit")) });
      rowIdx.push(i + 2);
    });
    if (!lines.length) return { ok: false, error: "Request not found" };
    if (d.items.length !== lines.length) return { ok: false, error: "The request has changed. Reload and try again" };
    for (let i = 0; i < lines.length; i++) {
      const a = d.items[i] || {}, ap = a.approved, reason = String(a.reason || "").trim().slice(0, 300);
      if (norm_(a.code) !== norm_(lines[i].code)) return { ok: false, error: "The request has changed. Reload and try again" };
      if (typeof ap !== "number" || !isFinite(ap) || Math.floor(ap) !== ap || ap < 0 || ap > 9999999) return { ok: false, error: "Invalid approved quantity for " + lines[i].name };
      if (ap !== lines[i].qty && !reason) return { ok: false, error: "A reason is required for " + lines[i].name };
      out.push({ approved: ap, reason: ap !== lines[i].qty ? reason : reason });
    }
    remarks = String(d.remarks || "").trim().slice(0, 1000);

    const ss = openMain(), fb = ensureSheet_(ss, FEEDBACK_TAB, FB_HEADERS), apSh = ensureSheet_(ss, APPROVAL_TAB, AP_HEADERS);
    if (fb.getLastRow() > 1) {
      fb.getRange(2, 2, fb.getLastRow() - 1, 1).getValues().forEach(r => { if (String(r[0]) === String(d.ref)) ver++; });
    }
    const now = new Date();
    const rows = lines.map((l, i) => [now, d.ref, ver, i + 1, l.code, l.name, l.qty, out[i].approved, l.unit, out[i].reason]);
    const a0 = apSh.getLastRow() + 1;
    apSh.getRange(a0, 5, rows.length, 2).setNumberFormat("@");        // keep item codes like 0001
    apSh.getRange(a0, 10, rows.length, 1).setNumberFormat("@");       // reasons are text, never formulas
    apSh.getRange(a0, 1, rows.length, AP_HEADERS.length).setValues(rows);
    const f0 = fb.getLastRow() + 1;
    fb.getRange(f0, 6, 1, 1).setNumberFormat("@");
    fb.getRange(f0, 1, 1, FB_HEADERS.length).setValues([[now, d.ref, ver, me.user, dept, remarks,
      lines.reduce((s, l) => s + l.qty, 0), out.reduce((s, o) => s + o.approved, 0), "", "", ""]]);

    // Pending -> Approved. Issued / Completed are left alone, so Main Stock never has to wait for anyone.
    let status = cur;
    if (cur === "Pending" || cur === "") {
      status = "Approved";
      const when = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd HH:mm");
      rowIdx.forEach(r => {
        sh.getRange(r, cols["Status"]).setValue(status);
        sh.getRange(r, cols["Status updated"]).setValue(when);
        sh.getRange(r, cols["Updated by"]).setValue(me.user);
      });
    }
    res = { ok: true, ref: d.ref, version: ver, status: status };
  } finally { lock.releaseLock(); }

  res.tg = notifyDeptApproval_(d.ref, dept, ver, by, lines, out, remarks);
  return res;
}

function notifyDeptApproval_(ref, dept, ver, by, lines, out, remarks) {
  const rows = lines.map((l, i) => {
    const a = out[i].approved, mark = a === l.qty ? "✅" : a === 0 ? "❌" : "⚠️";
    const q = a === l.qty ? l.qty + " " + l.unit + " (as requested)" : "requested " + l.qty + " → approved " + a + " " + l.unit;
    return mark + " " + l.name + " (" + l.code + "): " + q + (out[i].reason ? "\n    Reason: " + out[i].reason : "");
  }).join("\n");
  const msg = "Main Stock feedback on request " + ref + (ver > 1 ? " (revised, v" + ver + ")" : "") + "\nDepartment: " + dept + "\nReviewed by: " + by +
    "\n\n" + rows + (remarks ? "\n\nRemarks: " + remarks : "") + "\n\nOpen the form to view / print the approved slip." + (SITE_URL ? "\n" + SITE_URL : "");
  let tg = "none";
  const chat = lookupDept_(DEPT_TELEGRAM, dept);
  if (TELEGRAM_BOT_TOKEN && chat) {
    try {
      const r = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage",
        { method: "post", payload: { chat_id: String(chat), text: msg.slice(0, 4000) }, muteHttpExceptions: true });
      tg = r.getResponseCode() === 200 ? "sent" : "failed";
      if (tg === "failed") console.error("telegram", r.getContentText());
    } catch (e) { console.error("telegram", e); tg = "failed"; }
  }
  const mail = lookupDept_(DEPT_EMAILS, dept);
  try { if (mail) MailApp.sendEmail(mail, "Approved: request " + ref + " - " + dept, msg); } catch (e) { console.error("mail", e); }
  return tg;
}

// Refs with feedback this department has not opened yet (drives the bell / badge in the form)
function noticesFor(me) {
  const out = {}, ss = openMain(), fb = ss.getSheetByName(FEEDBACK_TAB);
  if (fb && fb.getLastRow() > 1) {
    fb.getRange(2, 1, fb.getLastRow() - 1, FB_HEADERS.length).getValues().forEach(r => {
      if (String(r[4]).trim().toLowerCase() === String(me.dept).trim().toLowerCase() && String(r[8]) !== "Y") out[String(r[1])] = 1;
    });
  }
  return { ok: true, refs: Object.keys(out) };
}

function markSeen(d, me) {
  const refs = Array.isArray(d.refs) ? d.refs.map(String) : [];
  const ss = openMain(), fb = ss.getSheetByName(FEEDBACK_TAB);
  if (!refs.length || !fb || fb.getLastRow() < 2) return { ok: true };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const v = fb.getRange(2, 1, fb.getLastRow() - 1, FB_HEADERS.length).getValues();
    v.forEach((r, i) => {
      if (refs.indexOf(String(r[1])) < 0 || String(r[8]) === "Y") return;
      if (String(r[4]).trim().toLowerCase() !== String(me.dept).trim().toLowerCase()) return;
      fb.getRange(i + 2, 9, 1, 3).setValues([["Y", new Date(), me.user]]);
    });
  } finally { lock.releaseLock(); }
  return { ok: true };
}

/* Run once from the editor to test the department Telegram message: change "Pharmacy" to a department in DEPT_TELEGRAM */
function testDeptNotify() {
  console.log(notifyDeptApproval_("REQ-TEST-00000", "Pharmacy", 1, "Main Stock",
    [{ code: "0001", name: "Test item", qty: 10, unit: "box" }], [{ approved: 6, reason: "Only 6 in stock" }], "Please collect tomorrow"));
}
