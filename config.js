/* ===========================================================
   Item Request Form - settings
   Edit this file only if you need to change a setting.
   To change ITEMS, just edit the two Excel files instead.
   =========================================================== */
window.FORM_CONFIG = {

  // Google Apps Script web-app URL (requests are sent here -> Google Sheet)
  SHEET_URL: "https://script.google.com/macros/s/AKfycbz-4lyWtHYqPRWo1oAFNKktZlSPjxaisQ3QIvNxI4o2ST_o-PtOMz0iEYuq2Z0cXq2qBg/exec",
  // Shown on the form and on the printed slip
  ORG_NAME: "Domrey Mother & Child Clinic",

  // Time zone used for the time printed on the slip
  TIMEZONE: "Asia/Phnom_Penh",

  // The Excel file with the main item list (same folder as index.html).
  // File names are case-sensitive on GitHub.
  ACTIVE_FILE: "Active_ALL.xlsx",   // all items: code | name | (optional) category

  // Departments that also get their own item list (shown first, with their own filter button).
  //   dept = department name exactly as in the account, file = Excel file, cat = filter button name
  // To add another department, add one more line here and upload its Excel file.
  DEPT_ITEMS: [
    { dept: "Laboratory",   file: "LAb_items.xlsx", cat: "Lab" },          // "CODE - Name" in column A
    { dept: "Housekeeping", file: "HK_items.xlsx",  cat: "Housekeeping" }  // code | name | category | default unit
  ],

  // Dropdown choices
  DEPARTMENTS: ["Pharmacy", "Laboratory", "Housekeeping", "Other"],
  UNITS: ["pcs", "box", "bottle", "vial", "tube", "sachet", "pack", "amp", "pair", "roll", "case", "container", "bundle"],

  // Category filter buttons are worked out automatically for Active_ALL:
  //   - code starting with S + digits  -> Supplies
  //   - name contains "vaccine"        -> Vaccines
  //   - name contains a word below     -> Milk
  //   - everything else                -> Medicines
  // (or type a category in column D of Active_ALL.xlsx to override any item)
  MILK_WORDS: ["milk", "btnature", "enfamilk", "enfalac"],
  CATEGORY_ORDER: ["Medicines", "Vaccines", "Milk", "Supplies"]
};
