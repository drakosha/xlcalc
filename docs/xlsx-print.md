# Worksheet printing

`xlcalc/xlsx/print` reads and writes XLSX print settings and calculates the page regions a host can preview or render.
The entry point works with workbook snapshots and sheet geometry. A host can import it when the user opens print preview
or exports a printed report.

## Read and write settings

`readXlsxWithPrint(bytes, options?)` returns the usual workbook, cached values and reader issues, plus `print`,
`printIssues` and `sheetDefaults`. `print` contains the print settings by worksheet name. It includes areas and titles, setup,
margins, print flags, headers and footers, and horizontal and vertical breaks. A `printIssues` entry identifies malformed
print settings or a device-specific printer part or header/footer drawing that requires separate handling.

`writeXlsxWithPrint(snapshot, print, options?)` takes the same snapshot and writer options as `writeXlsx`. Print settings
are worksheet metadata, passed separately from the calculation model. The entry writes a new workbook in one pass.
Use `xlcalc/xlsx/print/keep` to edit settings in an existing file while preserving its drawings and other source parts.
`openKeptWithPrint` reads the settings and adds the workbook carriers used by `writeKept`.
`writeKeptWithPrint(original, snapshot, print, options?)` replaces the modeled settings for each listed worksheet.
Start from the corresponding entry returned by the reader to retain fields you want to keep. A sheet can be cleared
by passing only its name. The opaque printer relationship on an existing sheet stays in the package. The new names
for areas and titles are scoped to the worksheet's position in the written file, which can include chart sheets.

Print area references use sheet-local A1 notation. Row and column numbers in titles and breaks are one-based. Margins
are in inches, as in SpreadsheetML. Paper and printer settings remain available to the host when it chooses the print
target.

## Layout

`printSheetGeometry(sheet, { font, usedRange, ...defaults })` converts workbook row and column spans into point
coordinates, including zero-size hidden lines and the sheet's default column width. The optional defaults are the
matching entry of `readXlsxWithPrint`'s `sheetDefaults`. They give the row height and the visible spans on a sheet that
hides empty rows by default. Pass the printable used range from the session or the report builder. For areas spanning whole columns or rows it supplies the last used line on the open
axis. A host that has more precise font or row metrics can pass its own `PrintSheetGeometry` to the layout function.

`layoutPrintPages(settings, geometry, target)` returns pages in print order. The host supplies the effective, oriented
paper size, the device insets and, when enabled, the unscaled heading dimensions, all in points. The layout uses the
paper body, selected scale, repeated titles and manual breaks to place sheet regions. A manual break that overlaps a
print area makes a cut across that area. Fit-to-page computes a common integer percentage for the sheet and ignores
manual breaks. The returned issues identify a fit limit reached at 10% and a line exceeding the available body.

`printPaperSize(setup, fallback)` resolves the requested Excel paper code and orientation into points. The renderer can
supply the effective paper and insets when the printer changes that request. A device that rounds line sizes at each
scale can pass `target.deviceGeometry` for the row and column edges in print points and `target.paginationReserve`
for the trailing clearance it holds back when breaking pages. The ordinary sheet edges continue to describe the
source rectangles.

Each region identifies an A1 range and rectangles in sheet and paper points. A renderer can use the range to fetch cells
and paint them at the target rectangle, clipping to the page body. Drawings with anchors in sheet points will use the
same projection. `projectPrintRect(rect, page, geometry, device?)` returns the intersections of a sheet rectangle with
the page's regions and their paper positions. When the device changes the line pitch, it interpolates within each row
and column. The renderer clips those positions to the page body. The host can include drawing extents in the used
range when it supplies geometry for them.

## Pictures and note boxes

`readXlsxDrawings(bytes, options?)` reads the file's drawing anchors, embedded picture references, note boxes and
chart frames. Join a note box to the corresponding `sheet.cells[cell].note` for its text. The box describes its
position, visibility and saved fill and stroke. A component can use `printDrawingBounds(anchor, geometry)` for the
sheet coordinates, including for on-screen display. `readImage(part)` loads the bytes of a reported picture on demand
and shares them across calls. A picture also carries its saved crop, rotation and flips when set.

`projectPrintDrawings(drawings, page, geometry, { setup, device })` projects one sheet's objects in drawing order.
For a page returned by `layoutPrintWorkbook`, it selects the page's sheet from the supplied drawings. Pass the
matching `settings.setup` and `target.deviceGeometry`. Each result has the full box in sheet points and its projected
regions. Pictures and chart frames are clipped to the printed regions. A shown note box in `asDisplayed` mode
uses its full position on every page where it overlaps the print area. Apply the paper clip when rendering those
boxes. The host can include object extents in the printable used range it passes to `printSheetGeometry`.

Current object coverage includes embedded pictures anchored over the grid, VML note boxes and chart frames.
Pictures placed inside cells, header/footer pictures and general DrawingML shapes use additional formats.

## Chart data

`xlcalc/xlsx/charts` is a separate entry for hosts that display or print charts. A chart frame from
`readXlsxDrawings` identifies the source part and anchor. `readXlsxCharts(bytes, options?)` reads the corresponding
ChartML and returns frames with their groups, series, axes, cached points and basic fill and line styles. Match
a chart to a drawing by `frame.id`. It reads bar and column, line, area, pie, doughnut and scatter groups, including
combinations of these groups. Chart sheets are identified as `frame.sheetKind: 'chartsheet'`. The worksheet
paginator lays out objects over worksheet grids.

Pass a current session to `resolveChartData(chart, session)` to read ordinary worksheet ranges from that state.
Pass a reader with `getValues` and `resolve(reference, sheet)` to supply results for other reference forms, such as
defined names. Unresolved sources and hierarchical category levels keep the file's saved data. The component
draws the chart using the model, including the axis IDs, sparse point indexes and visible-cells setting.

## Print several sheets

`layoutPrintWorkbook(sheets, options?)` lays out the supplied visible worksheets in order. Each entry supplies its
settings, geometry and target. The returned pages identify the source sheet, the physical `position` and the numeric
`number` printed by the header/footer codes. Set `fromPage` and `toPage` to output a physical interval. Use `totalPages` for the `&N` field. Excel uses the physical
position within the job for first, odd and even variants. Manual page starts affect the printed numbers. If the
first printed sheet specifies a starting number, Excel uses that number again for the next automatic sheet. An
explicit start on a later sheet advances the automatic numbering by the number of pages on that sheet.

## Print notes at the end of a sheet

When `setup.cellComments` is `'atEnd'`, pass `{ cells: sheet.cells, font: book.font }` as the fourth argument to
`layoutPrintPages` or as the worksheet's `notes` in `layoutPrintWorkbook`. The cells can come directly from
`readXlsxWithPrint`. The layout appends pages with `area: 0` and `noteLines` after the sheet's grid. These pages take part in
page numbers, header/footer variants, page selection and job limits. Fit-to-page applies to the grid. The note pages
use that scale and can add pages beyond the fit count.

Each line identifies its cell and gives the text and label positions in paper points. The positions describe line
boxes. The host applies its font's baseline and the device's page clip. Excel can start the final line close enough
to the bottom margin for its glyphs to extend beyond the grid body. The default labels are `Cell:` and `Note:`.
Supply `notes.labels` for the print locale. For precise line breaks and placement, pass `notes.metrics(scale)` with
the renderer's scaled line height and text measurements in paper points. The default uses nominal Normal-font
measurements.

## Excel compatibility work

The layout selects header/footer variants by physical page position and positions their bands.
`formatPrintHeaderFooter(text, context)` resolves text into left, center and right sections with formatted runs. The
context supplies page counters, sheet and file labels, date and time. Picture markers identify the position where a
renderer will place an image when it has the source part.

The renderer still needs to draw the cells, gridlines, headings and error display as requested. Printer-specific defaults
and mixed print targets need their own compatibility cases. A host supplies rendering for the cell grid and the
objects it reads, including note boxes, pictures and charts.

The XLSX reader test uses an Excel-saved print area and break. For pagination, compare the generated plan with Excel PDF
output under a fixed paper size, printer, Excel version and font set. Record the scale, row and column boundaries,
repeated titles and footer text for each page. Include narrow and wide sheets, both page orders, multiple print areas,
hidden lines, manual breaks and fit-to-width with unrestricted height. Keep the same tests when adding drawings that
cross a page boundary.

### Windows Excel PDF comparison

The Windows harness writes synthetic workbooks, has Excel export them to PDF, then records page membership and line
positions alongside the COM geometry. Run it in the oracle's current locale with `pdftotext` available.

```sh
XLWIN_HOST=your-excel-host XLWIN_CODEPAGE=any npx tsx bench/windows/print.ts ru-RU /tmp/xlcalc-print
```

Excel 16.0.20430 opened the landscape A4 report and printed `A1:A125` on four pages, with rows 1 and 2 repeated and
a numbered footer. Its last data rows were 35, 68, 101 and 125. The nominal sheet geometry produced 34, 66, 98 and 125.

On the measured A4 PDF device, a fixed 100% print of nominal 15 pt rows advances about 14.523 pt per row. The
different zooms in the matrix agree with rounding a 121-unit row at the requested scale. A 36th row fits at a bottom
margin of 0.496 inches and moves to the next page at 0.497 inches. In that case the measured grid and about 0.72 pt
of trailing clearance reproduce the boundaries. The workbook's font and the printer both affect those metrics.

The wide fit-to-width print matches the fixed 68% case on both axes and has the same rows on each page. The screen
column is 262.5 pt, while its unscaled print width is about 282 pt on the same device. The tests use those measured
axes to reproduce Excel's scale and its print areas. A Letter request on this driver exported an A4 PDF with rescaled
content. Use the requested paper to reproduce the break calculation in that case and handle the output adaptation
when rendering.

With page numbering set to start at 4, Excel uses the first, even, odd and even footers on the first four physical
pages. The numeric labels are 4 through 7. The offset tests also pin Excel's treatment of a digit at the end of a
header or footer field. Nine worksheet combinations printed as a workbook matched the layout's footer variants,
explicit and automatic page numbers, and total page count.

The captured matrix can be replayed against the measured profile with
`npx tsx bench/windows/print-replay.ts <control-matrix-directory> [additional-trace-directories]`.
The replay checks page count and the last visible row wherever the PDF includes a row marker. It reproduced
all 76 cases in the A4, margin, formatting and range batches measured on this printer. The profile is a benchmark
input. Other printers and unmeasured fonts and row heights need their own measurements for the same precision.

### Windows Excel note comparison

The `notes` cases retain text and word positions from the PDFs. The replay checks page count, footers, line text,
horizontal offsets and vertical pitch against the measured device. The eleven end-note cases cover hidden cells,
overlapping print areas, a margin change, fixed and fitted zooms, both grid page orders, text breaks and a long note
that flows across pages.

The end list follows print area order, then row and column order within each area. Notes in overlapping areas appear once
per area. Excel includes notes on hidden lines inside an area and leaves out notes that appear only in the repeated
titles. It places the cell address and the note body on separate lines, with a column for their labels. A page break
can fall between those lines or inside the note body. A shown note in `asDisplayed` mode prints as a text box
where the VML drawing puts it, including on a page that omits its source cell.

Windows Excel also exported an Excel-saved worksheet with a picture through five one-column print areas and a chart
through three adjacent areas. The picture appeared in pages 2–4 and the chart in pages 1–2, with each object clipped
to the respective area. The drawing test uses those page memberships and the note box that appears across area
boundaries. Exact device coordinates depend on the supplied print geometry.
