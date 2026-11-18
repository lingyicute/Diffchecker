# Diffchecker

English | [简体中文](README.md)

A dependency-free, front-end-only toolkit for comparing text, images and spreadsheets offline. Everything runs locally in your browser: nothing is uploaded and no network connection is needed.

| Page | Purpose |
| --- | --- |
| `text-compare.html` | Text comparison |
| `image-compare.html` | Image comparison |
| `excel-compare.html` | Spreadsheet comparison (`.xlsx` / `.csv` / `.tsv`) |

## Getting started

There is nothing to install or build. Open `index.html` in a modern browser that supports native dialogs (it redirects to the text comparison), or drop the whole folder onto any static web server.

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

- Six modes: split, fade, flicker, slider, highlight and subtract. Flicker can be paused and starts paused when reduced motion is preferred. Use the arrow keys or Home / End to move a focused slider.
- Automatic or manual alignment (shift, rotate, scale, flip). If the two images differ in size you are offered automatic alignment. The comparison frame includes both images after transforming them, without clipping pixels outside the original frame.
- Differing regions are detected and listed; click one to jump to it. Sensitivity is adjustable and several scans can be combined.
- Paste images from the clipboard (`Ctrl` / `⌘` + `V`). The focused image input takes precedence.
- Compare image transparency as well as colour. Export a PNG at full or working resolution, or print a PDF report with a summary, the unaligned source images and a change list.

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
- **`.diffchecker` files**: store both sides and the current options, and can be reopened on the matching page. Unsupported session versions are not loaded. Files from the text page and the spreadsheet page are not interchangeable.
- **Unsupported formats**: legacy binary `.xls` files must be saved as `.xlsx` or `.csv` first; encrypted or password-protected workbooks must have the password removed.
- **Dates and numbers**: numbers in `.xlsx` files are read with 15 significant digits, matching what Excel displays. With "Normalize dates" on, text such as `2024/3/5`, `15.03.2024` or `2024-03-05 00:00:00` is compared as an ISO date, and `9:30` as `09:30:00`, so comparing a CSV against an `.xlsx` does not report false differences. For dates written as `x/y/yyyy`, the second number greater than 12 means month/day/year; otherwise day/month/year is assumed. XLSX ISO dates and time formats with fractional seconds preserve up to three decimal places. Non-standard ISO dates with time zones are kept as written.
- **CSV export**: non-numeric text starting with `=`, `+`, `-` or `@` (including full-width variants and after leading whitespace) is prefixed with a single quote to keep it as text. Signed numeric values remain unchanged. Import recognises quoted multiline fields and rejects unclosed fields or extra text after a closing quote. If delimiters are ambiguous, save as `.tsv` or `.xlsx` first.
- **Large inputs**: the diff algorithm has a cost cap for files with many repeated lines and huge numbers of differences. The result is still a valid diff but may not be the shortest one. Very long lines or cells only have their common prefix and suffix trimmed, and the middle is marked as changed.
- **Text PNG**: long lines wrap. The canvas is 1600 px wide and limited to 16 Mi pixels. If the rendered image exceeds this limit, turn on "Hide unchanged lines" or export a PDF / Excel file instead.
- **Image limits**: each source can have up to 32 Mi pixels, with sides no longer than 16384 px. Preview and region detection cap the longest side at 2600 px. PNG export allows up to 16 Mi pixels; select working resolution if the full-size PNG is too large.
- **Spreadsheet limits**: an uncompressed XLSX part may be up to 64 MiB, and each sheet can import up to 1 million cells or fields. The expanded comparison can cover up to 500,000 positions across both sides, including row numbers. Chart sheets are not compared. Use another export format if the result exceeds Excel’s row, column or cell-length limits.

## Privacy and security

- The pages ship with a strict Content Security Policy: the text and spreadsheet pages use `default-src 'none'` and load no external resources; the image page additionally allows same-origin scripts, `blob:` and `wasm-unsafe-eval`.
- Image alignment uses a built-in FFT phase-correlation algorithm by default. If you inline `opencv.js` into the page (or include it with a `<script>` before the page script), ORB feature-based alignment is attempted when rotation and scaling are allowed and no flip is active. The built-in algorithm covers the other cases.

## Development and testing

The project itself has no dependencies. The regression tests extract the diff core, spreadsheet comparison, CSV / date parsing, zip reading and writing, and the image page's region detection and alignment code straight from the pages' inline scripts, and run them with Node's built-in test runner (Node 18 or later):

```sh
npm test
```

GitHub Actions runs the same tests on every push and pull request.

## License

[AGPL-3.0](LICENSE)
