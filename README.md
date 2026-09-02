# ENMAEYA Funding Dashboard — How it works

This is your dashboard built the "Option A" way: a public web page whose
numbers and text live in one editable file. No database, no server to run, no
logins to maintain. It's cheap, hard to break, and perfect for data that
changes about once a quarter.

## The three files

| File          | What it is                          | Who touches it |
|---------------|-------------------------------------|----------------|
| `index.html`  | The design + the code that draws the page | **Graphic designer** (look & feel) |
| `data.json`   | All the numbers, labels and text    | **Whoever updates data each quarter** |
| `README.md`   | This guide                          | Nobody edits |

The key idea: **design and data are separate.** The designer can restyle the
whole page (colours, fonts, spacing, layout) in `index.html` without ever
seeing a number. The person doing the quarterly update only edits `data.json`
and never has to touch code or design.

---

## Updating the data each quarter (the main task)

1. Open `data.json` in any text editor (even Notepad, but VS Code is nicer).
2. Change the values. For example, to update the top metric:
   ```json
   { "label": "Highest funding", "value": "$52M", "sub": "Ministry of Public Health" }
   ```
   Just edit the text   de the quotes.
3. Update the reporting period at the top:
   ```json
   "reportingPeriod": "1 January–21 July 2026",
   "lastRecord": "21 July 2026"
   ```
4. Save the file. Re-upload it to your host (see "Publishing" below). Done.

### Rules that keep it from breaking
- Keep the **quotes** and **commas** exactly as they are — only change the
  text between quotes, or the numbers.
- Bar widths use `"pct"` (0–100) — that's the length of the bar, not the money.
  Set the biggest item to 100 and scale the others relative to it.
- Heatmap cells accept a number 0–5 or `"—"` for none.
- The scatter chart is automatic: just edit each ministry's `projects`,
  `fundingM`, and `partners`, and the bubbles reposition themselves.
- After editing, paste the whole file into <https://jsonlint.com> to check
  there are no missing commas. This catches 99% of mistakes.

---

## Previewing it on your computer

Because the page reads `data.json`, you can't just double-click `index.html`
(browsers block that for security). Run a tiny local server instead:

**Option 1 — Python (already on most Macs/Linux):**
```
cd folder-with-these-files
python -m http.server 8000
```
Then open <http://localhost:8000> in your browser.

**Option 2 — VS Code:** install the "Live Server" extension, right-click
`index.html`, choose "Open with Live Server".

Once it's hosted online (below), this step isn't needed — it just works.

---

## Publishing it to the public (free)

Any of these host static sites for free and are ideal for this:

- **Netlify** — drag-and-drop the folder onto <https://app.netlify.com/drop>.
  You get a public link in seconds. To update, drag the folder again.
- **Vercel** — connect a GitHub repo or drag-and-drop; auto-publishes.
- **GitHub Pages** — push the files to a repo and enable Pages.

For quarterly updates: edit `data.json`, then re-deploy (drag the folder again
on Netlify, or push to GitHub). The public site updates immediately.

You can also point your own domain (e.g. `data.enmaeya.com`) at any of these.

---

## What this does NOT include (and when you'd want it)

This simple setup is intentionally minimal. You'd only outgrow it if you later
need:

- **A friendly admin screen** so non-technical staff never open `data.json` →
  move the data into **Airtable** or **Google Sheets** and have the page read
  from there. Small upgrade, same page.
- **Live/frequent updates, many editors, permissions, an audit trail, or user
  accounts** → then you'd add a real backend + database. Overkill for
  quarterly data, but the path is open.

For now, Option A covers everything you described.
