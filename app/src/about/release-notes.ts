/**
 * What changed in each version, in plain words, newest first. Shown on the What's new page (File → Account, with
 * the backstage in Phase 3) and once after an update. A test checks the newest entry matches the version built.
 */
export type ReleaseNote = { version: string; date: string; notes: string[] }

export const RELEASE_NOTES: ReleaseNote[] = [
  {
    version: "0.7.0",
    date: "2026-11-17",
    notes: [
      "Shapes (Insert tab): rectangles, circles, arrows, stars, flowchart shapes, callouts and lines. Pick one, then drag on the sheet where it goes (Shift keeps it even), or click for the standard size. Double-click a shape to type in it. Drag a shape to move it, its handles to size it, and the round handle above it to turn it (Shift turns in 15° steps).",
      "Text Box and WordArt (Insert tab): a box of text you can put anywhere on the sheet, and decorative text in eight styles with fill, outline, glow, shadow or reflection.",
      "Pictures (Insert tab): put a photo, logo or other picture from this device on the sheet. Icons (Insert tab): pick from more than 2,000 icons that come with Cells, searchable, several at once. They go in as sharp SVG pictures that Excel shows too.",
      "Shape Format and Picture Format tabs appear while a shape or a picture is selected: fill, outline (colour, weight and dashes), shadow, text colour, outline and effects, picture border, Reset Picture, Alt Text, and the exact height and width. Theme colours follow the workbook's theme.",
      "Arrange (on the Format tabs and Page Layout): Bring Forward, Send Backward, to front and to back, Align and Distribute, Group and Ungroup, Rotate and Flip. Ctrl+click or Shift+click selects more than one object; the arrow keys nudge them (Ctrl for one pixel at a time), Ctrl+D duplicates and Delete removes them. Copy and paste work on them, and other apps get a picture.",
      "Selection Pane (Page Layout › Arrange, or Home › Find & Select): every shape, picture and chart on the sheet, top first. Select, rename (double-click), hide or show each one, or reorder them with the arrows. Select Objects selects them all.",
      "Shapes, text boxes, WordArt, pictures, icons and groups open and save with Excel files: the ones made in Excel come back unchanged unless you edit them, and the ones made in Cells open in Excel. They print, count as Images in Workbook Statistics, and Check Accessibility now asks for alt text on pictures and shapes too.",
      "Equation (Insert tab): put a formula in the sheet, from Office's built-in ones or your own. Type it as you would in Office (x^2, a/b, √(x), ∑_(k=1)^n) and it is drawn like a printed formula when you click away. The Equation tab adds fractions, scripts, roots, integrals, brackets, matrices and math symbols where you are typing. Equations open and save with Excel files, both ways.",
      "3D Models (Insert tab): put a 3D model from a file in the sheet (.glb, .gltf, .fbx, .obj, .3mf, .ply or .stl) and turn it to show any side: drag the handle in its middle, pick one of Office's views, or type the turns in degrees on the 3D Model tab. Reset 3D Model turns it back. Models open and save with Excel files, and Excel shows the same view.",
      "SmartArt (Insert tab): pick a list, process, cycle or matrix and type your points in the Text Pane beside it. Enter adds a point, Tab and Shift+Tab move it down or up a level, and the shapes, arrows and text size follow as you type. The SmartArt Design tab adds and moves shapes, changes the layout or the colours, resets the graphic or turns it into ordinary shapes. SmartArt opens and saves with Excel files: Excel's own show as Excel drew them, and the ones made in Cells open in Excel as SmartArt you can keep editing there.",
      "Object (Insert tab): put a file of any kind in the workbook. It shows as an icon with the file's name; double-click it to open the file in its own app. Embedded files open and save with Excel files, so Excel opens them too, and files embedded in Excel open in Cells.",
      "Signature Line (Insert tab): add a line that says who should sign, with their title, email address and instructions. Double-click it to change them. Signature lines open and save with Excel files. Signing it with a digital ID isn't there yet.",
      "Windows (View tab): New Window shows the workbook in a second window, with its own sheet, place and zoom; edits show in both. Arrange All lays the windows out side by side, one above the other, tiled or cascaded, inside the Cells window; click a window to work in it. View Side by Side compares two windows and scrolls them together. Hide, Unhide and Switch Windows work too, and Ctrl+F6 goes to the next window.",
      "Themes (Page Layout tab): give the workbook one of Office's themes, or just its Colors or Fonts. Everything in theme colours and fonts changes with it, as in Excel: cells coloured from the Theme Colors part of a colour menu, cell styles, table styles, charts, shapes, and the font of cells without one of their own. Colours you picked under Standard Colors stay as they are. Browse for Themes uses the theme of an Office theme file (.thmx) or of another workbook. Themes open and save with Excel files.",
      "The colour menus show the workbook theme's colours, with the same lighter and darker steps Excel shows.",
      "Background (Page Layout tab): put a picture behind the sheet, repeated across it like a wallpaper, with your data on top. It shows on screen only and doesn't print, as in Excel; in dark mode it's dimmed so the text stays readable. The same button then reads Delete Background. Backgrounds open and save with Excel files.",
      "Spelling (Review tab, or F7): checks the text in your sheet, or in the cells you selected, with Ignore, Add to Dictionary, Change and Change All, like Excel. Words you add are shared with Inkspire. The English dictionary comes with Cells, so nothing goes online.",
      "Thesaurus (Review tab, or Shift+F7): synonyms and opposites for the word in the cell, grouped by meaning. Click a word to look it up, or Insert to put it in the cell. Works offline.",
      "Check Accessibility (Review tab, or the Accessibility note in the status bar): finds charts without alt text, tables without a header row, hard-to-read text colours, merged cells, unclear link text and default sheet names, takes you to each one and offers a fix, with why it matters.",
      "Alt Text for charts (Chart Design tab): describe a chart for people who use a screen reader, or mark it as decorative. Saved in Excel files and read from them.",
      "Allow Edit Ranges (Review tab): pick ranges of a sheet that people can still edit after you protect it, each with its own password if you like. A range with a password asks for it once, then stays open until you close the workbook. Works the same in Excel.",
      "Workbook Statistics (Review tab): how many cells with data, formulas, tables, charts, comments and notes the sheet and the whole workbook have, and where the sheet ends.",
      "Page Layout view (View tab, or the middle button by the zoom): the sheet as printed pages, with their margins, the header and footer, and blank pages after your data to type into. You can edit cells right on the pages; click the header or footer to change it. Rulers above and beside the pages measure in inches or centimetres (View › Ruler turns them off).",
      "Custom Views (View tab): save how the workbook looks under a name (each sheet's selection, scroll, zoom, frozen panes, gridlines and headings, and if you like its hidden rows and columns and print settings) and switch back to it later with Show. Custom views are saved in Excel files and work in Excel too, and the ones made in Excel open in Cells.",
      "Page Break Preview (View tab, or the buttons beside the zoom in the status bar): see the sheet as it will print, page by page, with what doesn't print greyed out, each page outlined in blue and numbered across. Automatic page breaks are dashed and the ones you set are solid. Drag a blue line to move a page break (drag it off the pages to remove it), or the outer edge to change what prints. The preview has its own zoom (60% to start), and Normal gets its zoom back. Excel files keep the view and both zooms.",
      "Split (View tab): divide the window into two or four panes that scroll on their own, to see distant parts of a sheet at once. The split goes above and left of the cell you're on (through the middle of the window from the top-left cell). Drag a split bar to move it, drag it to the edge or double-click it to take it away, or click Split again. The wheel and the scroll bars scroll the pane they belong to. Freeze Panes freezes where the split is, and Split on frozen panes splits there. Split panes are saved in Excel files, and Excel's open in Cells.",
      "Grouped sheets: Shift+click a sheet tab to group the sheets from the one you're on to that one, Ctrl+click to add or leave out one sheet, or right-click a tab and choose Select All Sheets (Ctrl+Shift+PgDn and PgUp add the next or previous sheet). While sheets are grouped, the title bar says [Group], and what you type, paste, clear or format goes into the same cells of every grouped sheet, as one step to undo; so do row heights, column widths, inserted and deleted rows and columns, and page setup. Delete and Hide on a grouped tab act on all of them. Home › Fill › Across Worksheets copies the selected cells (everything, the contents or the formats) to the other grouped sheets. Features that belong to one sheet, such as tables, charts, filters and comments, are greyed out until you ungroup (click a tab outside the group, or Ungroup Sheets). The group is saved in Excel files.",
      "Links (Insert › Link, or Ctrl+K): link a cell to a web page, a file, a place in this workbook or an e-mail address, with a ScreenTip of your own. Click a link to follow it, or click and hold to just select the cell; a workbook opens in Cells at the place the link names, and other files open in their own app after a warning. The cell menu has Edit, Open and Remove Hyperlink, and Home › Clear has Clear Hyperlinks and Remove Hyperlinks. Typing a web or e-mail address makes it a link, as in Excel. Links open and save with Excel files, .xls included, and so does the HYPERLINK function.",
      "Symbol (Insert tab): pick a character that isn't on your keyboard, such as © € → or ½, from a grid of Unicode sets, or a dash, space or quote from Special Characters, and Insert puts it in the cell you are typing in. The ones you use most stay at hand under Recently used symbols.",
      "Header & Footer (Insert tab) opens Page Setup on its Header/Footer page.",
      "Cells opens Excel 97-2003 workbooks (.xls): values, formulas, formatting, column widths, merged cells, frozen panes, names and notes. It opens a copy, and saving makes an .xlsx; the .xls stays as it was. On Android, Cells now shows up in Open with for .xls files, and the Open picker no longer greys them out.",
      "The cell menu's Filter and Sort work: filter by the cell's value, colour or font colour, clear one column's filter, reapply, Sort A to Z or Z to A, put the cell's colour on top, and Custom Sort. On a phone they open in the long-press panel.",
      "Paste Special is in the cell menu.",
      "The Office Clipboard (the small arrow by Home › Clipboard): the last 24 things you copied in Cells, to paste any of them again, or all at once.",
      "Show Comments (Review tab): every comment on the sheet in a pane beside it, in reading order. Pick one to go to its cell and reply; New starts one on the selected cell.",
      "Pick From Drop-down List (in the cell menu, or Alt+Down on a cell): choose from the words already typed in the column above and below.",
      "Sparklines (Insert tab): tiny line, column and win/loss charts inside cells. Select your data, pick Line, Column or Win/Loss and say where they go; one sparkline goes in each cell, each showing its own row (or column) of data. The Sparkline tab changes their type, shows high, low, first, last and negative points or markers, applies one of Excel's 36 styles or your own colours and line weight, sets the axis (same scale for the whole group, your own minimum and maximum, dates, right to left), decides how empty and hidden cells show, groups and ungroups them, and clears them. Sparklines from Excel files show as Excel draws them, follow inserted and deleted rows and columns, and are kept when you save.",
      "Slicers (Insert › Slicer, or Table Design › Insert Slicer): tick the table's columns and a slicer for each appears. The Slicer tab sets its caption, style, buttons (columns, height, width) and size, and arranges it like a shape; Slicer Settings renames it, hides its header, sorts its items the other way, and hides or shows items with no data.",
      "Slicers: the button panels Excel puts beside a table now show and work in Cells, in all 14 of Excel's styles. Click a button to show only those rows, Ctrl+click or Multi-Select to pick several, Shift+click for a run, and Clear Filter to show everything again; values the table's other filters leave no rows for are greyed and listed last. A slicer's picks are the table's own filter, so the header's dropdown always agrees, and both are kept when you save.",
      "Quick Analysis (Ctrl+Q, the small button at the corner of a selection, or right-click › Quick Analysis): one-click analyses of the data you selected, on five tabs. Formatting adds data bars, a colour scale, icons or highlights (different choices for numbers, text and dates); Charts shows the charts that suit the data; Totals adds sums, averages, counts, % of the total or running totals in the row below or the column to the right; Tables turns the data into a table; Sparklines puts a tiny chart beside each row. Point at a choice to see it on the sheet first; click to keep it, and Undo takes it back in one step.",
      "Dialogs opened from the ribbon now get the keyboard focus, so you can type in them straight away.",
      "PivotTables in Excel files are now known to Cells as PivotTables: they are kept exactly as Excel saved them, move along when you insert rows or columns above or before them, and their cells can't be typed over or pasted into by mistake (Excel's message explains why), so saving no longer risks breaking them. Refreshing and changing them comes next.",
      "Refresh PivotTables: Data › Refresh All (Ctrl+Alt+F5) reads the data of every PivotTable again, and Refresh (Alt+F5, or right-click in a PivotTable) the one you are in, along with the others made from the same data. New items join the PivotTable, ones no longer in the data leave it, and your hidden items, sorting and layout stay. If the PivotTable would grow over other cells, Cells asks first. PivotTables set to refresh when the file opens do so.",
      "Filter and sort by conditional formatting icons: a filter button's Filter by Color and Sort by Color list the icons the column shows (and No Cell Icon), the cell menu has Filter by Selected Cell's Icon and Put Selected Cell Icon On Top, and the Sort dialog sorts on Conditional Formatting Icon. Filter and Sort by Color now also see the colours conditional formatting gives cells, as Excel does. Icon filters open and save with Excel files.",
      "Sort A to Z and Z to A work again on the Data tab and in Home's Sort & Filter menu.",
      "Outlines (Data tab): Group and Ungroup rows or columns (Alt+Shift+Right and Left), Auto Outline from your SUM formulas, Clear Outline, Show and Hide Detail. The outline bar beside the headings has a level button for each level, and a − or + to fold each group away or open it again; Ctrl+8 hides or shows it. Settings chooses whether summaries sit below or above, right or left. Outlines are read from and saved to Excel files, .xls included.",
      "Subtotal (Data tab): a total at each change in a column, with the function you choose (Sum, Count, Average and more), a grand total, and the outline set up so level 2 shows only the totals. Remove All takes them out again.",
      "Goal Seek (Data › What-If Analysis): find the value for one cell that makes a formula give the result you want, such as the loan you can afford for a monthly payment.",
      "Advanced Filter (Data › Sort & Filter › Advanced): filter a list by conditions you write on the sheet (a row of headings, and under it the values or comparisons such as >50, each row another choice, or a formula), in place or as a copy elsewhere, and leave out repeated rows with Unique records only. Clear and Reapply work on it too.",
      "Formulas that use other workbooks, such as =[Prices.xlsx]Sheet1!B2, work: in workbooks from Excel they no longer turn into #NAME? errors, and they calculate from the values saved with the link, as Excel does when the other file is closed. Data › Workbook Links lists those workbooks: Update Values reads the file again, Change Source points the formulas at another file, and Break Link turns them into their values.",
      "Data from a text or CSV file (Data › From Text/CSV) stays linked to the file. Refresh (Alt+F5) or Refresh All (Ctrl+Alt+F5) reads it again when the file changes. Queries & Connections lists every link, with its file and where its data is; Properties renames it and can make it refresh each time the workbook opens. The links are kept in .xlsx files the way Excel keeps them, so Excel can refresh them too.",
      "Forecast Sheet (Data › Forecast): pick a column of dates or numbers and the values beside it, and Cells makes a new sheet that carries them forward: a table of forecasts with their confidence bounds and a chart of both, worked out by exponential smoothing that finds repeating seasons by itself. Options choose where the forecast starts and ends, the confidence level, the season length and how gaps and repeated dates are handled. The new functions behind it, FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY and FORECAST.ETS.STAT, work in any formula. Line charts now leave a gap for empty cells, as Excel does, instead of dropping to zero.",
      "Consolidate (Data › Data Tools): add up (or average, count, find the largest of…) ranges from several sheets into one table where you are, matching rows and columns by their labels or by position. With Create links to source data, the totals stay linked to the sheets they came from, with the detail folded away in an outline. Cells remembers the ranges for next time and saves them in Excel files.",
      "Scenario Manager (Data › What-If Analysis): save named sets of values for the cells your formulas use, such as a low and a high estimate, show them on the sheet in turn, merge them from another sheet, and compare them on a Scenario Summary sheet. Scenarios open and save with Excel files.",
      "Data Table (Data › What-If Analysis): see a formula's result for a list of values, or for every pair of two lists, such as the monthly payment for each interest rate and loan length. The table updates as the sheet changes, and Excel's data tables open and save with it.",
      "Pinching on a touchpad in the Linux desktop app zooms the sheet, about the pointer, instead of magnifying the whole window like a web page.",
      "Scrolling and panning are much smoother in the desktop app on Linux (and on phones and tablets): scrolling no longer redraws the ribbon and the panes with every step.",
      "AutoFit Column Width sizes numbers as the cells show them, not with every hidden decimal.",
      "Formula auditing (Formulas tab): Trace Precedents and Trace Dependents draw arrows on the sheet to the cells a formula uses or the formulas that use a cell, one level more with each click, red where an error travels and from a sheet icon for other sheets; Remove Arrows clears them. Trace Error follows an error back to where it starts.",
      "Name Manager (Ctrl+F3) lists every name with its value and what it refers to, to add, edit or delete names. Create from Selection (Ctrl+Shift+F3) names rows and columns from their labels.",
      "Evaluate Formula works a formula out one step at a time, underlining the part it works out next. Error Checking goes through the sheet's errors and numbers stored as text, explains each, and can convert, ignore or show the steps. The Watch Window keeps chosen cells' values in view while you work elsewhere.",
    ],
  },
  {
    version: "0.6.0",
    date: "2026-11-16",
    notes: [
      "Installers for Linux (.deb, AppImage and Flatpak), Windows and macOS, all made from one release.",
      "Cells registers itself for workbooks (.xlsx, .xlsm and templates), and offers itself for .csv and .tsv files: double-click a workbook in your file manager and it opens in Cells, in the Cells that's already running if there is one.",
      "On a Mac, Cells has its own menu bar (Cells, File, Edit, View, Format, Window, Help) with the usual ⌘ shortcuts. macOS keeps ⌘H, ⌘M and ⌘Q, so Replace is ⌃H and Go To is ⌃G, as in Excel for Mac. Quit asks about unsaved changes first.",
      "Cells on Android tablets, with the full ribbon. Open and Save As use Android's own file picker, so workbooks in Downloads, Drive or OneDrive open and save back where they are, and Recent opens them again later. A workbook you open with Cells from a file manager, mail or Drive opens in Cells. With a keyboard attached, the usual Ctrl shortcuts work, and the on-screen keyboard only comes up once you start typing in a cell.",
      "Cells on Android phones, with a layout of its own: a compact top bar (File, Undo, Redo, Find), bigger column and row headings, and a bottom bar with your sheets and the Home, Insert, Formulas, Data and View tools. Each tool opens the ribbon's commands in a panel from the bottom, which you can switch to any tab and swipe down to close. While you type in a cell, a row of formula keys (= ( ) , : $ + - * / & \" and fx) sits above the keyboard. Dialogs and the File view take the whole screen. A tablet keeps the full ribbon, and a split screen switches between the two as it's resized, without losing anything.",
      "On a phone, pressing and holding a cell (or a row or column heading) opens a panel from the bottom: Cut, Copy and Paste as big buttons, then Insert, Delete, Clear Contents, notes and comments, Format Cells and the rest.",
      "On a touch screen the sheet glides on after a quick swipe, and stops when you touch it.",
      "On Android, Back closes what's open (a panel, a dialog, a menu, the File view or the cell you're typing in), and with nothing open it puts Cells in the background with your workbook as it was.",
    ],
  },
  {
    version: "0.5.0",
    date: "2026-11-16",
    notes: [
      "A short setup the first time Cells opens: your name (notes and comments are signed with it), how you write dates and numbers, and light or dark. It starts from what your system says, or from what you already set in Inkspire, and Skip keeps those.",
      "File › Options: your user name and initials, the Office Theme (your system's setting, White or Black), and on the Advanced page your region, short date and separators. Run setup again is there too.",
      "Dates and numbers follow your region: in a day-first region 31/1/2026 is a date, with a decimal comma 1,5 is a number, and cells show dates, decimals and thousands your way. Formulas keep Excel's English form, so files mean the same in Excel anywhere.",
      "Only one Cells runs at a time. A cells:// link (Inkspire's coming \"Edit in Cells\" uses one) or a workbook opened from your file manager opens in the Cells that's already running and comes to the front. A link can point at a range, and Cells selects it.",
      "Copy keeps the formatting: a range pastes into Inkspire, Word or mail as a formatted table (fonts, colours, fills, borders, widths and merged cells), and remembers where it came from so Inkspire can link to it.",
      "Charts can be copied (Ctrl+C, or cut with Ctrl+X). In Inkspire the chart pastes as a chart you can still edit; other apps get a picture. Pasting it back into Cells makes a second chart of the same data.",
      "Ready for Inkspire's next update: tables it links from a workbook follow your changes in Cells as you make them.",
      "Also for Inkspire's next update: Edit Data on an Inkspire chart opens the chart's numbers in Cells (\"Chart in Report.docx\"); every change goes back into the document by itself.",
      "AI tools, if you want them (off until you turn them on in the setup or File › Options › AI): Home › AI Tools writes a formula from your words, explains a formula, suggests a chart or summarizes your data. Use an AI on your own computer (Ollama, LM Studio) or a service you have a key for; Cells shares the choice and the keys with Inkspire, and always says where your cells are sent.",
      "These settings are kept in a file Cells shares with Inkspire, so they're set once for both. Cells picks up a change the next time you switch to it. Inkspire starts using the file with its next update.",
    ],
  },
  {
    version: "0.4.0",
    date: "2026-11-15",
    notes: [
      "Open and save Excel workbooks (.xlsx, .xlsm and templates). Files made in Excel come back out with their charts, pictures, macros and other parts Cells doesn't show yet kept as they were.",
      "File › Home, New, Open, Save As and Info, with a Recent list you can pin files to. Ctrl+N, Ctrl+O, Ctrl+S, F12 and Ctrl+W work as in Excel.",
      "Cells asks to save your changes before closing or opening another file, and the title bar shows the workbook's name.",
      "AutoRecover keeps a copy of unsaved work every two minutes. If Cells closes unexpectedly, Document Recovery offers it the next time.",
      "CSV and text files: a .csv opens at once (a semicolon CSV too) and saves back in its own delimiter and encoding. A .txt opens through the Text Import Wizard: delimited or fixed width, the first row, the file's encoding, and each column as General, Text, a date in any order, or skipped. Data › From Text/CSV puts a file on a new sheet.",
      "Save As CSV UTF-8 or Text (Tab delimited), written as Excel writes them; with several sheets Cells asks first and saves the active one.",
      "Drop a workbook or CSV on the window to open it.",
      "Big files show their progress while opening. If a file can't be opened, Cells says why (damaged, password-protected, the old .xls format) and lists anything it left out.",
      "Sort: Sort A to Z and Z to A on the data around the cell (headers stay on top), and the Sort dialog for several levels, cell or font colour, custom lists like Jan, Feb, Mar, case-sensitive sorting and left-to-right. The order is checked against Excel.",
      "Filter (Ctrl+Shift+L): a dropdown on each header with sorting, a search box, a list of values to tick (dates by year, month and day), Text, Number and Date Filters, Top 10, Above Average and Filter by Color. The status bar counts the rows shown; Clear and Reapply are on the Data tab. Filters open and save with Excel files, and give the same rows as Excel.",
      "Find and Replace (Ctrl+F, Ctrl+H, Shift+F4 for the next match): in the sheet or the whole workbook, in formulas, values, notes or comments, with wildcards, Match case and Match entire cell contents, and Find All's list of matches. The dialog stays open while you work in the sheet.",
      "Go To Special (Ctrl+G › Special…): blanks, constants, formulas, notes, current region and array, row and column differences, precedents, dependents, the last cell, visible cells only and conditional formats.",
      'Remove Duplicates (Data tab): pick the columns to compare, with or without a header row; duplicate rows go and the rest move up, compared the way Excel does (12 and 12.0 match, "apple" and "Apple" match).',
      "Text to Columns (Data tab): split a column by a delimiter or at fixed widths, set each new column's format, and choose where the result goes. It splits the way Excel does, formulas included.",
      "The formula bar shows a date as 3/19/2026 and a percentage as 100%, as Excel does, instead of the number underneath.",
      "Data Validation (Data tab): limit what a cell takes (whole numbers, decimals, a list, dates, times, text length or a formula of your own), show a message when the cell is selected, and choose what happens when someone types something else: stop it, warn, or just tell them. A list gets a dropdown beside the cell (Alt+Down opens it). Circle Invalid Data rings the cells that break their rule, and Go To Special finds cells with rules. The rules open and save with Excel files.",
      "Tables: Insert › Table (Ctrl+T) or Home › Format as Table turns your data into a table with one of Excel's 60 styles, a filter button on every header and banded rows. The Table Design tab renames it, resizes it, adds a total row (pick Sum, Average, Count and more from each total's dropdown), turns the header row, banded rows and columns and first or last column on and off, and converts it back to a normal range. Formulas can name the table's columns (=SUM(Sales[Amount]), [@Qty]*[@Price]) and follow when you rename them. The table grows as you type beside or under it, Tab on its last cell adds a row, and a formula typed into an empty column fills the whole column. Tables open and save with Excel files.",
      "Flash Fill (Ctrl+E, on the Data tab and under Home › Fill): type one or two examples of what you want next to your data, such as first names from full names, initials, an email address made from two columns or the digits of a phone number, and Cells fills in the rest of the column the same way.",
      "Charts: Insert › Charts (or Alt+F1) draws a chart of the data around the selected cell: column, bar, line, area, pie and doughnut, scatter and more, with Recommended Charts first. Drag a chart to move it, drag its handles to size it (Shift keeps its shape), Delete removes it. The Chart Design tab adds titles, labels, legends and gridlines, applies quick layouts, colours and styles, changes the chart type, switches rows and columns and picks the data with Select Data. Charts redraw as you change the numbers, follow inserted rows and renamed sheets, and open and save with Excel files.",
      "Printing: File › Print (Ctrl+P) shows each page as it will print, with the active sheet, the whole workbook or the selection, a page range, orientation, paper size, margins and scaling. Page Layout sets margins, orientation, paper size, the print area, page breaks, titles that repeat on every page, Fit to a number of pages, and printed gridlines and headings. The Page Setup dialog adds headers and footers (page numbers, date, file and sheet name) and centring on the page. After printing or setting up the page, dashed lines on the sheet show where the pages break. All of it opens and saves with Excel files.",
      "File › Export saves the sheet, the workbook or the selection as a PDF.",
    ],
  },
  {
    version: "0.3.0",
    date: "2026-11-14",
    notes: [
      "The ribbon: Home, Insert, Page Layout, Formulas, Data, Review, View and Help, with KeyTips (press Alt), Ctrl+F1 to collapse it, a Quick Access Toolbar you can add commands to, and Search (Alt+Q) to find any command. Commands that aren't built yet are greyed out and say so.",
      "Cell formatting: fonts, colours, fills, borders, alignment, wrap, rotation, indents, merging, cell styles, Format Painter, and the Format Cells dialog (Ctrl+1) with all six tabs.",
      "Number formats work like Excel's, including custom codes, dates and times, fractions and conditions; checked against Excel on 594 cases.",
      "Conditional formatting: highlight rules, top/bottom, data bars, colour scales, icon sets, formula rules, and Manage Rules.",
      "Sheet tabs: add, rename, move and copy (drag, or Ctrl+drag), colour, hide and unhide.",
      "View: freeze panes, zoom (status bar, Ctrl+wheel, pinch), gridlines and headings on and off, and dark mode.",
      "Notes and threaded comments; Protect Sheet and Protect Workbook, with passwords.",
      "The formula bar grows to several lines and can be resized. The Name Box lists and creates names, and fx opens Insert Function.",
      "On a touch screen: tap to select, drag the round handles to stretch the selection, long-press for the menu, double-tap to edit, pinch to zoom. The sheet scrolls so the keyboard doesn't cover the cell you're typing in.",
      "This page opens by itself once after an update.",
    ],
  },
  {
    version: "0.2.0",
    date: "2026-11-14",
    notes: [
      "Formulas calculate: 262 Excel functions, dynamic arrays that spill, circular reference warnings and manual calculation with F9.",
      "References in a formula are coloured, with matching boxes on the grid; point at cells with the mouse or the arrow keys.",
      "While typing a formula, a list suggests functions (Tab inserts one) and a tip shows the function's arguments.",
      "F4 cycles a reference through $A$1, A$1, $A1 and A1. Ctrl+` shows the formulas instead of their results.",
      "Inserting or deleting rows and columns, moving cells and copying formulas update the references, as in Excel.",
      "Calculation runs in the background, so the sheet stays responsive while it works.",
    ],
  },
  {
    version: "0.1.0",
    date: "2026-11-14",
    notes: [
      "The first version: a fast grid of 1,048,576 rows by 16,384 columns with Excel's keyboard shortcuts.",
      "Typing, editing, fill handle, copy and paste, insert and delete, row heights and column widths, hide and unhide.",
      "Undo and redo up to 100 steps.",
    ],
  },
]

export const CURRENT_VERSION = RELEASE_NOTES[0].version

const SEEN_KEY = "cells.seen-version"

/**
 * Whether to open What's New on start: once after an update (a different version was seen before). A first start
 * only records the version. Storage can be missing or blocked; then nothing shows.
 */
export function whatsNewOnStart(storage: Pick<Storage, "getItem" | "setItem"> | null = safeStorage()): boolean {
  if (!storage) return false
  try {
    const seen = storage.getItem(SEEN_KEY)
    storage.setItem(SEEN_KEY, CURRENT_VERSION)
    return seen !== null && seen !== CURRENT_VERSION
  } catch {
    return false
  }
}

function safeStorage(): Storage | null {
  try {
    return typeof localStorage === "undefined" ? null : localStorage
  } catch {
    return null
  }
}
