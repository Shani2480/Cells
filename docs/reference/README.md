# Reference material

- `excel-ribbon.json`: Excel's ribbon (tabs, groups, controls in order, their ids and labels) and its right-click
  menus for cells, rows, columns and sheet tabs. The spec for names, grouping and order of every Cells ribbon tab and
  menu. Control ids and placement come from Microsoft's
  [Office Fluent UI command identifiers](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)
  (MIT License, Copyright (c) Microsoft Corporation); labels and the right-click menus were read from Excel 16.0
  (build 17932) itself. Groups without a label are variants Excel shows only in some editions (Python, Power Query
  previews) and are not built.
