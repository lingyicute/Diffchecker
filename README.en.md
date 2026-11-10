# Diffchecker

English | [简体中文](README.md)

A dependency-free, front-end-only toolkit for comparing text, images and spreadsheets offline. Everything runs locally in your browser: nothing is uploaded and no network connection is needed.

| Page | Purpose |
| --- | --- |
| `text-compare.html` | Text comparison |
| `image-compare.html` | Image comparison |
| `excel-compare.html` | Spreadsheet comparison (`.xlsx` / `.csv` / `.tsv`) |

## Getting started

There is nothing to install or build. Open `index.html` in a browser (it redirects to the text comparison), or drop the whole folder onto any static web server.

The interface language follows your browser by default. Use the buttons at the top right to switch between English and 中文, and between dark and light themes.

## Features

### Text comparison

- Split and unified views; inline differences at smart, word or character granularity.
- Syntax highlighting (JSON, JavaScript / TypeScript, Python, XML / HTML, CSS, SQL, with auto-detection). Multi-line comments and strings are coloured correctly.
- Pre-processing and ignore rules: hide whitespace changes, lowercase, sort lines, trim whitespace, replace line breaks with spaces, ignore quote / dash differences, and custom ignore words.
- "Hide unchanged lines" folds long identical stretches. Click a fold bar (or focus it with `Tab` and press `Enter` / `Space`) to expand it.
- When a changed block has a different number of removed and added lines, similar lines are matched up first, so an inserted or deleted line does not misalign the modified lines after it.
- Open files or drag them onto an input box; optionally compare in real time as you type.
- Export to PDF, Excel, PNG or a `.diffchecker` file.

### Image comparison

- Six modes: split, fade, flicker, slider, highlight and subtract.
- Automatic or manual alignment (shift, rotate, scale, flip). If the two images differ in size you are offered automatic alignment.
- Differing regions are detected and listed; click one to jump to it. Sensitivity is adjustable and several scans can be combined.
- Paste images from the clipboard (`Ctrl` / `⌘` + `V`).
- Export the current view as PNG, or a PDF report with a summary, the source images and a change list.

### Spreadsheet comparison

- Reads `.xlsx` (including `.xlsm`, `.xltx`, `.xltm`), `.csv` and `.tsv`. You can pick the sheet and the header row (a title block above the header is skipped automatically).
- Match by header (detects added, removed and moved columns) or compare row by row; show cell values or formulas.
- Options: ignore whitespace, ignore case, normalize dates, sort rows / columns before comparing, hide unchanged rows / columns, show the old value in changed cells.
- Export to Excel, CSV, PDF or a `.diffchecker` file.

## Keyboard shortcuts

| Key | Action |
| --- | --- |
| `Ctrl` / `⌘` + `Enter` | Text page: find differences |
| `Esc` | Close the side panel and dialogs |
| `1` / `2` (hold) | Image page, flicker mode: temporarily show the original / changed image |

## Files and formats

- **Encoding**: text and CSV files are decoded as UTF-8 first. If that fails, GB18030 is tried (so common GBK Chinese files are read correctly), then windows-1252 (Western European text). UTF-16 is also recognised, with or without a BOM (without one it is detected from the distribution of NUL bytes).
- **Binary files**: the text page rejects files containing NUL bytes instead of filling the editor with garbage.
- **`.diffchecker` files**: store both sides and the current options, and can be reopened on the matching page. Files from the text page and the spreadsheet page are not interchangeable.
- **Unsupported formats**: legacy binary `.xls` files must be saved as `.xlsx` or `.csv` first; encrypted or password-protected workbooks must have the password removed.
- **Dates and numbers**: numbers in `.xlsx` files are read with 15 significant digits, matching what Excel displays. With "Normalize dates" on, text such as `2024/3/5`, `15.03.2024` or `2024-03-05 00:00:00` is compared as an ISO date, and `9:30` as `09:30:00`, so comparing a CSV against an `.xlsx` does not report false differences. For dates written as `x/y/yyyy`, the second number greater than 12 means month/day/year; otherwise day/month/year is assumed.
- **CSV export**: non-numeric text starting with `=`, `+`, `-` or `@` is prefixed with a single quote, so Excel does not execute it as a formula.
- **Large inputs**: the diff algorithm has a cost cap for files with many repeated lines and huge numbers of differences. The result is still a valid diff but may not be the shortest one. Very long lines or cells only have their common prefix and suffix trimmed, and the middle is marked as changed.
- **PNG export**: a text comparison taller than 32000 px cannot be exported as a single PNG. Turn on "Hide unchanged lines", or export a PDF / Excel file instead.

## Privacy and security

- The pages ship with a strict Content Security Policy: the text and spreadsheet pages use `default-src 'none'` and load no external resources; the image page additionally allows same-origin scripts, `blob:` and `wasm-unsafe-eval`.
- Image alignment uses a built-in FFT phase-correlation algorithm by default. If you inline `opencv.js` into the page (or include it with a `<script>` before the page script), ORB feature-based alignment is used automatically instead.

## Development and testing

The project itself has no dependencies. The regression tests extract the diff core, spreadsheet comparison, CSV / date parsing, zip reading and writing, and the image page's region detection and alignment code straight from the pages' inline scripts, and run them with Node's built-in test runner (Node 18 or later):

```sh
npm test
```

GitHub Actions runs the same tests on every push and pull request.

## License

[AGPL-3.0](LICENSE)
