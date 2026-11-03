# External references

A formula can read another workbook: `[1]Sheet1!A1`, `'[1]My Sheet'!A1:B9`, `[1]!Rate`. Alongside such formulas, Excel stores in the file itself what it last read from that workbook: where it is, its sheets, the values of the cells the formulas read, and its names. The workbook keeps the same data as `externalBooks`, and the engine calculates the references from it, as Excel calculates a file it opens without updating its links. The other workbook is never opened: a grid displays a financial model that depends on a dozen other files with the values those files had when the model was saved.

```js
const externalBooks = [
  {
    target: '../data/Prices.xlsx',
    sheets: [{ name: 'Sheet1', cells: { A1: 5, A2: 7, B1: 'pear', B2: null } }, { name: 'Rates' }],
    names: [{ name: 'Rate', ref: 'Sheet1!$A$2' }],
  },
]
```

Element `i` of `externalBooks` corresponds to `[i+1]`. `target` is where the link points, as the file stores it: a file name, a relative or an absolute path, a `file:` or web address. A file stores a space as `%20` (`Report%202025.xlsx`); a host may give it as a space too. A formula displays and finds the link by the target with its escapes undone (`[Report 2025.xlsx]`), and `writeXlsx` and `writeKept` write a space as `%20`, as Excel does, so a link given with a space comes back from the file with `%20`. A sheet's `cells` are its stored values by address (a number, text, a boolean or an error), with `null` for a cell stored as empty. `readXlsx` reads them from the file's links and `writeXlsx` writes them back ([docs/xlsx.md](xlsx.md)).

## In short

```js
import { compute, toLocal } from 'xlcalc'

const externalBooks = [{ target: 'Prices.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 5, A2: 7, B1: 'pear' } }] }]
const book = { sheets: [{ name: 'S', cells: { A1: { formula: "[1]Sheet1!A1+1" }, A2: { formula: 'SUM([1]Sheet1!A:A)' }, A3: { formula: "COUNTIF([1]Sheet1!A:A,5)" } } }], externalBooks }
const cells = compute(book).sheets[0].cells
console.log(cells.A1.value, cells.A2.value, cells.A3.value, cells.A1.formula) // → 6 12 { error: 'VALUE' } '[1]Sheet1'!A1+1
console.log(toLocal(cells.A1.formula, 'en-US', { externalBooks })) // → { ok: true, formula: '[Prices.xlsx]Sheet1!A1+1' }
```

A cell the stored values do not include is blank, as Excel treats it, unless the sheet's last update failed (`refreshError`) or nothing from the sheet is stored: then it is `#REF!`. The criteria functions' ranges (`SUMIF`, `COUNTIFS` and the others), `COUNTBLANK` and the database functions accept only a reference to an open workbook and return `#VALUE!` over a link, as in Excel; `INDEX`, `MATCH`, `VLOOKUP`, `XLOOKUP`, `SUMPRODUCT` and the rest read the stored values. Entering a formula with a link in a session makes Excel mark that workbook as unreadable, and the session does the same: from then on, its cells that the stored values do not include are `#REF!`. Formulas are stored with the link's number; `toLocal`, `toInvariant`, `parseEntry`, a viewport's `formulaLocal` and `editText` display and parse them as Excel does, by file name (`[Prices.xlsx]Sheet1!A1`), and reject a workbook the workbook has no link to. [docs/links.md](links.md) has the rules and how they were verified.

## Values

- A cell stored in the link returns its value: `[1]Sheet1!A1+1` is 6. Excel stores `#SPILL!` and `#CALC!` from another workbook as `#VALUE!`, and the engine calculates them the same way (along with its own `#CYCLE!`, which a self-referencing name returns, and `#ERROR!`).
- A cell not stored in the link is blank, as Excel treats it: `[1]Sheet1!A9` is 0, `[1]Sheet1!A9&"|"` is `|`, and `SUM([1]Sheet1!A:A)` sums the stored cells of the column. Excel stores the non-empty cells of the part of a sheet that the formulas read, so the cells it leaves out were empty.
- On a sheet whose last update failed, which Excel marks as `refreshError` (`ExternalSheet.refreshError`), a cell not stored in the link is `#REF!`: `SUM([1]Sheet1!A:A)` is `#REF!`, while `MATCH(3,[1]Sheet1!A:A,0)` still finds a stored 3. A sheet the link names but stores nothing from (no `cells`) is `#REF!` everywhere.
- A range is read cell by cell, so `INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `LOOKUP`, `SUMPRODUCT`, `COUNT`, `COUNTA`, `AVERAGE`, `RANK`, `SUBTOTAL`, `AGGREGATE`, `ROWS`, `ROW`, `ISREF` and a dynamic array such as `[1]Sheet1!A1:A3*2` work over a link as over a sheet of the workbook.
- The criteria functions' ranges (`SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`), `COUNTBLANK` and the database functions' range accept only a reference to an open workbook: over a link they return `#VALUE!`, as Excel does over a closed workbook, whether referenced directly, through a name or through a LET name. A linked value used as a criterion is read like any value: `COUNTIF(A1:A3,[1]Sheet1!A1)` counts.
- A name in the linked workbook returns what it refers to: `[1]!Rate` the workbook-level name, `[1]Sheet1!Local` the sheet's own name, otherwise the workbook-level one. A name that contains a constant (`=5`) or a lost reference is `#REF!`, as in Excel. A name in this workbook that refers to a link (`Ext` = `'[1]Sheet1'!$A$2`) is a reference like any other. A reference that Excel lost in a linked workbook (`[1]Sheet1!#REF!`) is `#REF!`.
- A reference to a workbook that is not among the links, or to a sheet missing from its link, is `#REF!` with the error code `external_ref`.
- A table of a linked workbook (`[1]!Sales[[#This Row],[Qty]]`, `SUM([1]!Sales[Qty])`) is `#REF!`, as Excel calculates it while that workbook is closed, whatever the link stores; the formula keeps the reference, and a file gets it back as Excel writes it. Written without `!` (`[1]Sales[Qty]`), as some writers leave it, the reference is lost as Excel reads the file: the formula becomes `#REF!`.

The error of a cell that reads a stored error names the linked cell: `value not found, from '[1]Sheet1'!C1`. `getCellIssue` reports the formula that reads the link as the cell where the error originated, and `explain` shows the linked cell as Excel shows it.

## Editing

The stored values are not editable: no edit of the workbook changes them, and inserting or deleting rows and columns, moving a range, and renaming or deleting a sheet leave a link's references as written, even when a sheet of the workbook has the same name as the link's sheet.

A copy adjusts a link's relative references, as Excel adjusts them: `[1]Sheet1!A1` filled one row down becomes `[1]Sheet1!A2`, while `[1]Sheet1!$A$1` stays. Beyond the edge of the sheet it becomes `#REF!`.

Entering a formula that refers to a linked workbook, whether typed, set or pasted (`setCells`, `setCellText`, `pasteText`, `copyRange` and the rest), makes Excel try to read that workbook, which is closed: it marks every sheet of that link as `refreshError`, and from then on its cells that the stored values do not include are `#REF!`, including in formulas that already read them. A session does the same: its workbook's link gets the marker, and `getBook()` returns it, so a full calculation of the session's workbook matches the session. Other links keep their state. Renaming a sheet or inserting rows does not set the marker.

A host that reads the other workbook itself says so with `open: true` on the link, given with the values it read (`setExternalBooks`). Excel reads a linked workbook that is open, or a closed one it finds on the disk, in the same way: a cell missing from the values is blank, as an empty cell of that workbook is (`SUM('[1]Sales'!H:H)` sums the column, `ISBLANK('[1]Sales'!Z99)` is TRUE), a sheet without `cells` is an empty sheet, and entering a formula over it sets no `refreshError`, which its sheets must not have (an `InputError`). Given again without `open`, the link keeps its values unmarked until a formula over it is entered, as Excel keeps what it read after the workbook is closed. A file stores no such flag: `writeXlsx` and `writeKept` write the values without the marker and a sheet without values as an empty one, as Excel saves a workbook it read, and `readXlsx` gives the link without `open`.

```js
s.setExternalBooks([{ target: 'Sales.xlsx', sheets: [{ name: 'Sales', cells: { H1: 10, H2: 20 } }], open: true }])
s.setCellText('S', 'B1', '=SUM([Sales.xlsx]Sales!H:H)+[Sales.xlsx]Sales!Z99')
console.log(s.getCell('S', 'B1')?.value) // → 30
```

## Formula text

The stored formula numbers the link and quotes its sheet, as every stored reference is quoted: `'[1]Sheet1'!A1`, `'[1]My Sheet'!A1:B9`, the workbook-level name `[1]!Rate`, and a sheet-level name `'[1]Sheet1'!Local`. A file writes `[1]Sheet1!A1`, with quotes only where the sheet name needs them, and reading it back produces the same formula.

`toLocal` with `externalBooks`, a viewport's `formulaLocal` and `editText`, and the text of `explain` display a link's reference as Excel displays it, with the link's file name in brackets and the path the link stores:

| stored | link's target | displayed |
|---|---|---|
| `'[1]Sheet1'!A1` | `Prices.xlsx` | `[Prices.xlsx]Sheet1!A1` |
| `'[1]My Sheet'!A1` | `Prices.xlsx` | `'[Prices.xlsx]My Sheet'!A1` |
| `'[1]Sheet1'!A1` | `../data/Prices.xlsx` | `'../data/[Prices.xlsx]Sheet1'!A1` |
| `'[1]Sheet1'!A1` | `file:///C:\data\Prices.xlsx` | `'C:\data\[Prices.xlsx]Sheet1'!A1` |
| `'[1]Sheet1'!A1` | `Prices%202025.xlsx` | `'[Prices 2025.xlsx]Sheet1'!A1` |
| `[1]!Rate` | `Prices.xlsx` | `Prices.xlsx!Rate` |
| `[1]!Rate` | `../data/Prices.xlsx` | `'../data/Prices.xlsx'!Rate` |
| `[1]!Sales[[#This Row],[Qty]]` | `Prices.xlsx` | `Prices.xlsx!Sales[@Qty]` |

Typed text can use either form or the number: `toInvariant` with `externalBooks`, `parseEntry` with `externalBooks` and a session's `setCellText` find the link by its file name, matching the full path first, in any case, and number it (`=[Prices.xlsx]Sheet1!A1` becomes `'[1]Sheet1'!A1`). They compare typed text with the target as displayed: `[Report 2025.xlsx]` finds the target `Report%202025.xlsx`, while typed text is taken as written, since a file may have `%` in its name. A user who opens a cell for editing and presses Enter keeps the formula. A workbook that the workbook has no link to is rejected with `external_ref` and a message in the workbook's language: the other file is outside the model. A host that has read the other workbook gives the link itself, with the values it needs: in `BookInput.externalBooks` before it opens the workbook, or in a session with `setExternalBooks(books)`, which replaces the links, calculates the workbook again and is one step of undo. A formula reads a link by its number, so a new link goes after the others, which keep their places; then typed text finds it by its file name (`='[Report.xlsx]Sheet'!$I$8`). The same call gives a link more values, or values read again. `writeXlsx` writes every link, and `writeKept` writes a new link and the values of a changed one over the original file ([Writing over the original file](xlsx-keep.md#links-to-other-workbooks)).

```js
import { createSession } from 'xlcalc'

const s = createSession({ sheets: [{ name: 'S', cells: {} }] })
s.setExternalBooks([{ target: 'Report.xlsx', sheets: [{ name: 'Sheet', cells: { I8: 42 } }] }])
s.setCellText('S', 'A1', "='[Report.xlsx]Sheet'!$I$8*2")
const cell = s.getCell('S', 'A1')
console.log(cell?.formula, cell?.value) // → '[1]Sheet'!$I$8*2 84
```

`validateFormula` and `validateNames` check references against `ctx.externalBooks`; without it, a reference to another workbook is rejected.

## Known differences

- Excel displays a reference to a closed workbook with the absolute path it resolves from the folder of the opened file (`'/Users/…/[Prices.xlsx]Sheet1'!A1`, `'C:\work\[Prices.xlsx]Sheet1'!A1`), which a browser does not know: the engine displays the path the link stores, or just the file name if it stores none.
- Only entering or pasting a formula with a link was verified to mark the link. Whether moving such a formula, undoing the edit, or entering a formula that reads a link only through a name marks it was not verified: the engine marks the link for formulas whose text names it, on entering and pasting, and an undo keeps the marker.
- On a marked sheet, a range is read only to one row and one column beyond the stored values, so that an entire column stays cheap; a sum over it is `#REF!` as in Excel, but a count of errors over it (`COUNTA`) counts one `#REF!` where Excel counts every cell missing from the values.
- A reference across sheets of another workbook (`[1]Jan:Mar!A1`) and a spill reference into one (`[1]Sheet1!A1#`) are rejected.
- `INDIRECT` into a linked workbook is `#REF!`, and `OFFSET` and `CELL` over one are `#VALUE!`, as Excel calculates them over a closed workbook, also over a link held `open`, whose workbook Excel would calculate them over. A table of a linked workbook is `#REF!` there too.
- Excel shows a reference to a workbook that is open by its name alone (`[Sales.xlsx]Sales!A1`); a link held `open` shows the path it stores, as any link.
- A link of a kind other than a workbook (DDE, OLE) keeps its number and is reported by `readXlsx`; its references are `#REF!`.

## Checked against Excel

Tested in Excel for Mac and Excel for Windows, each opening every file without updating its links and returning the same values everywhere; the engine returns their values and their formula text.

Excel for Windows also typed formulas over a linked workbook placed beside the file: opened, opened and closed before calculating, left closed on the disk, and missing (`test/excel/win/links/open-held.json`). In the first three it read the workbook, a cell the workbook does not hold was blank, no sheet was marked, and the file it saved stored the values it read; a missing workbook marked the link. A session holding the link `open` returns the same values. The predictions, committed before the run, matched 15 of 16 cells: a formula typed before over a closed workbook on the disk was read again too.
