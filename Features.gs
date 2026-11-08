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
  return null;
}

// Stock team: all requests. Everyone else: only their own department.
function listRequests(me) {
  const sh = getSheet(), cols = ensureCols(sh), stock = isStock(me), last = sh.getLastRow();
  if (last < 2) return { ok: true, stock: stock, statuses: STATUSES, requests: [] };
  const v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const g = (r, h) => r[cols[h] - 1];
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
  return { ok: true, stock: stock, statuses: STATUSES, requests: order.reverse().slice(0, 1000).map(k => map[k]) };
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
