# What Cells puts on the clipboard (for Inkspire)

PLAN.md Phase 5, item 4; Inkspire's side is `docs/inkspire-todo.md` I4. Everything travels as `text/plain` and
`text/html`, which every app and every webview (WebKitGTK included) passes between programs. A custom MIME type
would not reliably reach another app from a webview, so the full-fidelity part rides inside the HTML as data
attributes.

## A range

- `text/plain`: tab-separated values, as Excel writes them (what each cell shows).
- `text/html`: a table as Excel's: `<col width>` per column, row heights, each `<td>` with its font, size, bold,
  italic, underline, strike, colour, fill, alignment, indent, wrap and borders inline, merged cells as
  `colspan`/`rowspan`. Hidden rows and columns are left out. Word, LibreOffice and Inkspire paste it as a formatted
  table.
- On the `<table>`, `data-cells` (JSON, HTML-escaped):

  ```json
  {
    "app": "cells",
    "version": 1,
    "file": "/home/alex/Q1 Budget.xlsx",
    "sheet": "Sales",
    "range": "A1:D10",
    "link": "cells://open?file=%2Fhome%2Falex%2FQ1%20Budget.xlsx&range=Sales%21A1%3AD10"
  }
  ```

  `file` and `link` are null for a workbook that was never saved. Inkspire's Paste Special › Paste Link turns this
  into a `LINK Excel.Sheet.12 "<file>" "Sales!R1C1:R10C4" \a \f 4 \h` field, and "Edit in Cells" opens `link`.

## A chart

- `text/plain`: the chart's title (or its name, "Chart 1").
- `text/html`: Inkspire's own chart markup, so its editor reads it as an editable chart with no extra code:

  ```html
  <div
    data-figure=""
    data-kind="chart"
    data-variant="column"
    data-chart="{…ChartSpec…}"
    data-width="480"
    data-height="288"
    data-alt="Sales by quarter"
    data-cells-chart="{…}"
  >
    <img src="data:image/png;base64,…" width="480" height="288" alt="Sales by quarter" />
  </div>
  ```

  `data-chart` is the whole chart spec (the `@cells/chart` package is Inkspire's chart spec), with the numbers read
  from the cells at the moment of copying. The picture inside is what other apps paste (a PNG drawn when the chart is
  selected; an SVG if the PNG isn't ready). `data-cells-chart` holds the source:

  ```json
  {
    "app": "cells",
    "version": 1,
    "token": "…",
    "file": "/home/alex/Sales.xlsx",
    "sheet": "Sheet1",
    "name": "Chart 1",
    "range": "=Sheet1!$A$1:$C$4"
  }
  ```

  Pasting it back into Cells makes a second chart with the same ranges (as Excel does); the `token` is how Cells
  knows its own copy.

Code: `packages/core/src/clipboard-html.ts`, `app/src/sheet/chart-clipboard.ts`, `app/src/grid/chart-png.ts`.
