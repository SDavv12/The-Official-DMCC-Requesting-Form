# Item Request Form

A static website: `index.html` reads the item lists straight from two Excel files and sends each request to a Google Sheet.

## Files

| File | What it is | Edit it? |
|---|---|---|
| `index.html` | The form | No |
| `config.js` | Google Sheet URL, departments, units | Rarely |
| `Active_ALL.xlsx` | All pharmacy / vaccine / milk / supply items | **Yes – this is your item list** |
| `LAb_items.xlsx` | Lab items (only shown for the Laboratory department) | **Yes – this is your lab list** |
| `Code.gs` | Google Apps Script for the Google Sheet (paste into Google, not needed on GitHub) | Once |

## Change items

1. Edit the Excel file (add, rename or delete rows) and save with the **same file name**.
2. On GitHub: open the repo -> **Add file -> Upload files** -> drop the Excel file -> **Commit changes**.
3. Wait 1-2 minutes, refresh the website. Done - no code involved.

**Active_ALL.xlsx** - first sheet is read. Per row: code | name | optional category. Blank rows and a header row are ignored.
Category is automatic (`S...` codes = Supplies, name has "vaccine" = Vaccines, milk brands = Milk, rest = Medicines).
To force a category (a new filter button appears automatically), type it in the column after the name.

**LAb_items.xlsx** - first sheet is read. One item per row in column A, written like `T0001 - Influenza A&B`.

Keep item codes unique. Two items with code `0255` currently exist (see browser console warnings).

## Put it online (GitHub Pages)

1. Create a repo on github.com and upload `index.html`, `config.js`, `README.md` and both Excel files.
2. **Settings -> Pages -> Build and deployment**: Source = *Deploy from a branch*, Branch = `main`, folder = `/ (root)` -> Save.
3. After a minute your site is at `https://<your-username>.github.io/<repo-name>/`.

The page must be opened from that address (or any web server). Double-clicking `index.html` on your computer cannot load Excel files because of browser security.

## Google Sheet

`config.js` already contains a Google Apps Script URL. If requests already arrive in your sheet, keep it.
Otherwise follow the steps at the top of `Code.gs` and paste the new URL into `SHEET_URL`.
