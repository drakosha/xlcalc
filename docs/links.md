# Links to other workbooks

A formula can read another workbook: `[1]Sheet1!A1`, `'[1]My Sheet'!A1:B9`, `[1]!Rate`. Excel stores beside such formulas, in the file itself, what it last read of that workbook: where it is, its sheets, the values of the cells the formulas read and its names. The book keeps the same as `externalBooks`, and the engine computes the references from it, as Excel computes a file it opens without updating its links. The other workbook is never opened: a grid shows a financial model that leans on a dozen other files with the values those files had when the model was saved.

```js
const externalBooks = [
  {
    target: '../data/Prices.xlsx',
    sheets: [{ name: 'Sheet1', cells: { A1: 5, A2: 7, B1: 'pear', B2: null } }, { name: 'Rates' }],
    names: [{ name: 'Rate', ref: 'Sheet1!$A$2' }],
  },
]
```

Element `i` of `externalBooks` answers `[i+1]`. `target` is where the link points, as the file holds it: a file name, a relative or an absolute path, a `file:` or web address. A sheet's `cells` are its stored values by address (a number, text, a boolean or an error), `null` for a cell stored empty. `readXlsx` reads them from the file's links and `writeXlsx` writes them back ([docs/xlsx.md](xlsx.md)).

## In short

```js
import { compute, toLocal } from 'xlcalc'

const externalBooks = [{ target: 'Prices.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 5, A2: 7, B1: 'pear' } }] }]
const book = { sheets: [{ name: 'S', cells: { A1: { formula: "[1]Sheet1!A1+1" }, A2: { formula: 'SUM([1]Sheet1!A:A)' }, A3: { formula: "COUNTIF([1]Sheet1!A:A,5)" } } }], externalBooks }
const cells = compute(book).sheets[0].cells
console.log(cells.A1.value, cells.A2.value, cells.A3.value, cells.A1.formula) // → 6 12 { error: 'VALUE' } '[1]Sheet1'!A1+1
console.log(toLocal(cells.A1.formula, 'en-US', { externalBooks })) // → { ok: true, formula: '[Prices.xlsx]Sheet1!A1+1' }
```

A cell the stored values do not hold is blank, as Excel takes it, unless the sheet's last update failed (`refreshError`) or nothing of the sheet is stored: then it is `#REF!`. The criteria functions' ranges (`SUMIF`, `COUNTIFS` and the others), `COUNTBLANK` and the database functions take only an open book's reference and give `#VALUE!` over a link, as in Excel; `INDEX`, `MATCH`, `VLOOKUP`, `XLOOKUP`, `SUMPRODUCT` and the rest read the stored values. A formula with a link entered in a session makes Excel mark that workbook as it cannot read it, and so does the session: its cells the stored values lack are `#REF!` from then on. Formulas are stored with the link's number; `toLocal`, `toInvariant`, `parseEntry`, a window's `formulaLocal` and `editText` show and read them as Excel does, by the file name (`[Prices.xlsx]Sheet1!A1`), and refuse a workbook the book has no link to. [docs/links.md](links.md) has the rules and how they were measured.

## Values

- A cell the link stores gives its value: `[1]Sheet1!A1+1` is 6. Excel stores `#SPILL!` and `#CALC!` of another workbook as `#VALUE!`, and so does the engine compute them (and its own `#CYCLE!`, which a name that uses itself gives, and `#ERROR!`).
- A cell the link does not store is blank, as Excel takes it: `[1]Sheet1!A9` is 0, `[1]Sheet1!A9&"|"` is `|`, and `SUM([1]Sheet1!A:A)` sums the stored cells of the column. Excel stores the non-empty cells of the part of a sheet the formulas read, so the cells it leaves out were empty.
- On a sheet whose last update failed, which Excel marks `refreshError` (`ExternalSheet.refreshError`), a cell the link does not store is `#REF!`: `SUM([1]Sheet1!A:A)` is `#REF!`, `MATCH(3,[1]Sheet1!A:A,0)` still finds a stored 3. A sheet the link names but stores nothing of (no `cells`) is `#REF!` everywhere.
- A range is read cell by cell, so `INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `LOOKUP`, `SUMPRODUCT`, `COUNT`, `COUNTA`, `AVERAGE`, `RANK`, `SUBTOTAL`, `AGGREGATE`, `ROWS`, `ROW`, `ISREF` and a dynamic array such as `[1]Sheet1!A1:A3*2` work over a link as over a sheet of the book.
- The criteria functions' ranges (`SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`), `COUNTBLANK` and the database functions' range take only an open book's reference: over a link they give `#VALUE!`, as Excel gives them over a closed book, directly, through a name or a LET name. A linked value as a criterion is read as any value: `COUNTIF(A1:A3,[1]Sheet1!A1)` counts.
- A name of the linked workbook gives what it refers to: `[1]!Rate` its book's name, `[1]Sheet1!Local` the sheet's own name, else the book's. A name that holds a constant (`=5`) or a lost reference is `#REF!`, as in Excel. A name of this book over a link (`Ext` = `'[1]Sheet1'!$A$2`) is a reference like any other. A reference Excel lost into a linked book (`[1]Sheet1!#REF!`) is `#REF!`.
- A reference to a book the links do not hold, or to a sheet its link lacks, is `#REF!` with the check code `external_ref`.

The error of a cell that reads a stored error names the linked cell: `value not found, from '[1]Sheet1'!C1`. `getCellIssue` gives the formula that reads the link as the cell the error began in, and `explain` shows the linked cell as Excel shows it.

## Editing

The stored values are not edited: no edit of the book reaches them, and inserting or deleting rows and columns, moving a range and renaming or deleting a sheet leave a link's references as written, a sheet of the book named like the link's sheet too.

A copy moves a link's relative references, as Excel moves them: `[1]Sheet1!A1` filled a row down is `[1]Sheet1!A2`, `[1]Sheet1!$A$1` stays. Off the sheet it is `#REF!`.

Entering a formula that refers to a linked book, typed, set or pasted (`setCells`, `setCellText`, `pasteText`, `copyRange` and the rest), makes Excel try to read that book, which is closed: it marks every sheet of that link `refreshError`, and its cells the stored values lack are `#REF!` from then on, in the formulas that read them already too. A session does the same: its book's link takes the mark, and `getBook()` gives it, so a full computation of the session's book agrees with the session. Other links keep their state. Renaming a sheet or inserting rows marks nothing.

## Formula text

The stored formula numbers the link and quotes its sheet, as every stored reference is quoted: `'[1]Sheet1'!A1`, `'[1]My Sheet'!A1:B9`, the book's own name `[1]!Rate`, a sheet's `'[1]Sheet1'!Local`. A file writes `[1]Sheet1!A1`, with quotes only where the sheet needs them, and reading it gives the same formula.

`toLocal` with `externalBooks`, a window's `formulaLocal` and `editText` and the text of `explain` show a link's reference as Excel shows it, by the link's file in brackets, with the path the link holds:

| stored | link's target | shown |
|---|---|---|
| `'[1]Sheet1'!A1` | `Prices.xlsx` | `[Prices.xlsx]Sheet1!A1` |
| `'[1]My Sheet'!A1` | `Prices.xlsx` | `'[Prices.xlsx]My Sheet'!A1` |
| `'[1]Sheet1'!A1` | `../data/Prices.xlsx` | `'../data/[Prices.xlsx]Sheet1'!A1` |
| `'[1]Sheet1'!A1` | `file:///C:\data\Prices.xlsx` | `'C:\data\[Prices.xlsx]Sheet1'!A1` |
| `[1]!Rate` | `Prices.xlsx` | `Prices.xlsx!Rate` |
| `[1]!Rate` | `../data/Prices.xlsx` | `'../data/Prices.xlsx'!Rate` |

Typed text reads either form and the number: `toInvariant` with `externalBooks`, `parseEntry` with `externalBooks` and a session's `setCellText` find the link by its file name, the whole path first, in any case, and number it (`=[Prices.xlsx]Sheet1!A1` is `'[1]Sheet1'!A1`). A person who opens a cell for editing and presses Enter keeps the formula. A workbook the book has no link to is refused with `external_ref` and a message in the book's language: making a new link would need the other file, which is outside the model. `validateFormula` and `validateNames` check references against `ctx.externalBooks`; without it a reference to another workbook is refused.

## Known differences

- Excel shows a closed book's reference with the absolute path it resolves from the folder of the opened file (`'/Users/…/[Prices.xlsx]Sheet1'!A1`, `'C:\work\[Prices.xlsx]Sheet1'!A1`), which a browser does not know: the engine shows the path the link holds, a file name alone without one.
- Only entering or pasting a formula with a link was measured to mark the link. Whether moving such a formula, undoing the edit or entering a formula that reads a link only through a name marks it was not: the engine marks for the formulas whose text names the link, on entering and pasting, and an undo keeps the mark.
- On a marked sheet a range is read to one row and one column past the stored values, so that a whole column stays cheap; a sum over it is `#REF!` as in Excel, but a count of errors over it (`COUNTA`) counts one `#REF!` where Excel counts every cell the values lack.
- A reference across another workbook's sheets (`[1]Jan:Mar!A1`) and a spill reference into one (`[1]Sheet1!A1#`) are refused.
- `INDIRECT`, `OFFSET` and `CELL` are not in the engine (`#NAME?` in any book): over a closed book Excel gives `#REF!`, `#VALUE!` and no text.
- A link of another kind than a workbook (DDE, OLE) keeps its number and is reported by `readXlsx`; its references are `#REF!`.

## How it was checked

`npx tsx bench/coverage/run.ts links <set>` writes each case of `bench/coverage/links/<set>.ts` as `writeXlsx` writes its own cells, with the formulas over links and the link parts put in by hand as Excel writes them, and has Excel for Mac open it with `update links do not update links`; `npx tsx bench/windows/run.ts links <set>` has Excel for Windows open it with `UpdateLinks` 0. Neither showed a window about the links. Each reads the rectangles after opening, takes the case's operations (a fill down, a copy, rows inserted, a sheet renamed), reads them again after a full calculation, and saves the book. The answers are in `test/excel/links/` and `test/excel/win/links/` with the files Excel saved; the predictions in each set were committed before Excel ran it.

| set | what it covers |
|---|---|
| `first` | one cell, the shape of the step |
| `basics-held` | cells, ranges, whole columns and rows, cells not stored, text, booleans and errors stored, sheets with a space and in Cyrillic, the linked book's names, this book's names over a link, lookups, criteria, `INDIRECT`, `OFFSET`, a dynamic array, a fill down, a copy, a renamed sheet, inserted rows, failed updates, two links, six forms of a target |
| `edits-held` | which edit makes a cell not stored `#REF!`: none, fills and copies onto stored and unstored cells, a rename, inserted rows |
| `refresh-held` | the mark on stored values, on empty data, a sheet without data, and which sheets and links a fill marks |
| `functions-held` | the functions that take only a reference, a linked criterion, a linked range through a name, more functions that read values |

Both Excels gave the same values everywhere. `test/links-excel.test.ts` checks, for each platform, the values after opening against `compute` of the case's book, the values after the operations against a session that takes them, the formulas Excel showed against `toLocal` with the links' targets in the opened file's folder, and that Excel's answer was the predicted value or one of its alternatives; the predictions it did not confirm are named there with the reason (a cell not stored was predicted blank, and is, until a formula with the link is entered). Excel also built a link itself, to a book it had open: it stored only the non-empty cells of the source's used part, an empty text as `t="str"` with an empty value, and opened without the source gave the stored values, blank elsewhere. `test/links-session.test.ts` runs an edit under formulas with links, an entered link and a refused new link in a synchronous session, in the same thread and in a pool, in both execution modes, against a full computation of the session's book. `test/links.test.ts` covers the text, the check, the input check, the input hash and structure operations.
