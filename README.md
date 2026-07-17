# EdiFinance

A personal finance tracker built as a single HTML file. No backend, no accounts, no subscriptions — just open it in a browser and start using it.

Built for the Iranian Toman, with support for importing directly from a structured Excel workbook.

---

## What it does

- Track multiple income streams and see them broken down visually
- Log expenses against budget categories and watch the progress bars move
- Import your existing data from an Excel workbook (matches the Khordad Budget sheet structure)
- Export everything back to Excel — the exported file can be re-imported without any data loss
- All data lives in your browser's `localStorage`. Nothing leaves your machine.

---

## Getting started

Download `finance-tracker.html` and open it in any modern browser. That's it.

No build step. No `npm install`. No server. Just a file.

---

## Importing from Excel

Click **Import Excel** in the top right and upload your workbook. If the file has multiple sheets, you'll be asked to choose one.

The importer expects this column layout (matching the Khordad Budget format):

| Col | Content              |
| --- | -------------------- |
| B   | Category name        |
| C   | Budget total         |
| D   | Spent                |
| G   | Income source name   |
| H   | Income amount        |
| K   | Transaction name     |
| L   | Transaction amount   |
| N   | Transaction category |

---

## Exporting

Click **Export Excel** to download your current data as a `.xlsx` file. The exported file follows the same column structure above, so you can re-import it later or share it with someone else using the same app.

---

## Data & privacy

Everything is stored in `localStorage` — meaning it stays on your device, in your browser, on that machine. Clearing browser data will erase it. If you want a backup, use the export button.

There is no sync, no cloud, no account. The footer says it best:

> _localStorage, not "the cloud" — there is no cloud, it's just someone else's computer_

---

## Stack

Pure HTML, CSS, and vanilla JavaScript. Three CDN dependencies:

- [Tailwind CSS](https://tailwindcss.com) — utility classes
- [Chart.js](https://www.chartjs.org) — donut and bar charts
- [SheetJS (xlsx)](https://sheetjs.com) — Excel import and export

Fonts: [Fraunces](https://fonts.google.com/specimen/Fraunces), [Inter](https://fonts.google.com/specimen/Inter), [JetBrains Mono](https://fonts.google.com/specimen/JetBrains+Mono) via Google Fonts.

---

## License

MIT. Do whatever you want with it.
