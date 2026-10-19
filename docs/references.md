# Reference operators

The kinds of references that Excel files use. `@A1:A10` takes one cell of a range. `(A1:A3,C1:C3)` combines ranges on one sheet. `'Jan:Mar'!A1:B2` is the same rectangle on every sheet from Jan to Mar. `A1:C3 B2:D4` is the cells two ranges share, and `A1:INDEX(B1:B9,3)` is the range from one reference to another.

The rules below were verified in Excel 16.113 for Mac (en-US) with dynamic array formulas.

```js
import { compute } from 'xlcalc'

const month = (n) => ({ cells: { A1: { value: n }, A2: { value: n * 10 } } })
const book = {
  sheets: [
    { name: 'Jan', ...month(1) },
    { name: 'Feb', ...month(2) },
    { name: 'Mar', ...month(3) },
    {
      name: 'Sum',
      cells: {
        B1: { value: 5 }, B2: { value: 'text' }, B3: { value: 7 }, D1: { value: 100 }, D3: { value: 300 },
        // @ takes the cell of B1:B3 in the formula's row, as a reference
        C2: { formula: 'SUM(@B1:B3)' }, C3: { formula: '@B1:B3*2' },
        E1: { formula: 'SUM((B1:B3,D1:D3))' }, E2: { formula: 'INDEX((B1:B3,D1:D3),3,1,2)' },
        F1: { formula: "SUM('Jan:Mar'!A1:A2)" }, F2: { formula: 'SUM(B1:D3 D1:D3)' }, F3: { formula: 'SUM(B1:INDEX(B1:B3,3))' },
      },
    },
  ],
}
const s = compute(book).sheets[3].cells
console.log(s.C2.value, s.C3.value, s.E1.value, s.E2.value)
// → 0 14 412 300
console.log(s.F1.value, s.F2.value, s.F3.value)
// → 66 400 12
```

## The @ operator

`@x` is implicit intersection: a single value where `x` could produce several.

| x | `@x` returns |
|---|---|
| a column `A1:A10` | the cell in the formula's row, or `#VALUE!` when the row is outside it |
| a row `C1:E1` | the cell in the formula's column |
| a block `G1:H2` | the cell in the formula's row and column |
| an entire column `A:A`, an entire row `1:1` | the cell in the formula's row or column |
| a range on another sheet | the cell in the same row and column there |
| a spill reference `C3#` | the cell of the spill range in the formula's row |
| an array, or a call that returns an array | its top-left value (`@{7,8}` is 7, `@SEQUENCE(3)` is 1) |
| a call that returns a reference (`INDEX(A1:B10,0,1)`) | the cell of that reference |
| a union | `#VALUE!` |
| a reference across sheets | `#REF!` |
| a single value | that value |

The result is a reference to that cell, so `SUM(@A1:A10)` skips text there, `ROW(@A1:A10)` is the formula's row and `ISREF(@A1:A10)` is TRUE.

`@` binds to what follows it: a reference, a name, a call, a constant or a parenthesis. `@A1:A10*2` is `(@A1:A10)*2`. An xlsx file stores it as `_xlfn.SINGLE(A1:A10)`, and `readXlsx` and `writeXlsx` convert between the two forms.

## Unions

A union is written in parentheses inside a call: `SUM((A1:A3,C1:C3))`. A name's formula can be a union without them. All parts must be references on one sheet; otherwise the union is `#VALUE!`.

| formula | returns |
|---|---|
| `SUM((A1:A3,C1:C3))`, `COUNT`, `MAX`, `AVERAGE`, `LARGE`, `MEDIAN`, `STDEV` | the function over every part |
| `SUBTOTAL(9,(A1:A3,C1:C3))`, `AGGREGATE(9,6,(A1:A3,C1:C3))` | the function over every part, skipping nested `SUBTOTAL` cells and hidden rows as in a single range; overlapping parts are counted twice |
| `FREQUENCY((A1:A3,C1:C3),E1:E3)`, or a union of bins | the counts over the numbers of every part |
| `SORTBY((A1:A3,C1:C3),…)`, or a union as a sort order | `#VALUE!` |
| `INDEX((A1:A3,C1:C3),2,1,2)` | C2, from the second area |
| `INDEX((A1:A3,C1:C3),2,1,3)` | `#REF!`: there is no third area |
| `AREAS((A1:A3,C1:C3))` | 2 |
| `VSTACK((A1:A3,C1:C3))`, `TOCOL` | the parts one after another |
| `ROWS((A1:A3,C1:C3))` | `#REF!` |
| `(A1:A3,C1:C3)` alone, or `+1` | `#VALUE!` |

## References across sheets

`'Jan:Mar'!A1:B2` covers the sheets from Jan to Mar in the workbook's order, including Feb between them. The two ends can be in either order. Sheet names that look like cell addresses (`S1`, `Q3`) must be quoted there: Excel parses `S1:S3!A1` as the range operator between cell S1 of the formula's own sheet and A1 on S3, so it is `#VALUE!` on any other sheet and the range A1:S1 on S3. A single such sheet needs no quotes (`S1!A1`).

The list functions take each sheet's rectangle in turn: `SUM`, `COUNT`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`, `PRODUCT`, `STDEV`, `VAR` and the others over lists of numbers, `AND`, `OR`, and `FREQUENCY`, as its data or its bins. `VSTACK`, `HSTACK`, `TOCOL` and `TOROW` stack the sheets' rectangles. `CHOOSE` and `IF` pass one through.

On its own or with an operator it is `#REF!` (`=Jan:Mar!A1+1`). `INDEX`, `AREAS`, `ROWS`, `SUBTOTAL` and `AGGREGATE` return `#VALUE!`, `SORTBY` and a value compared by `SWITCH` return `#REF!`, and `ISREF` returns FALSE.

## Sheet names in references

A formula puts a sheet's name in quotes before its reference only where the name needs them, and `quoteSheetIfNeeded` writes a name the way Excel does. Excel quotes a name:

- that is not a letter or `_` followed by letters, digits, `_` and `.`: `'Q1 data'`, `'2024'`, `'Sheet-1'`, `'€uro'`, `'.a'`; letters of any script count (`Лист1`, `Straße`, `日本`, `a.b`, `_x` stay unquoted), and an apostrophe inside is doubled (`'O''Brien'`);
- that is a cell of the sheet in A1 notation, in either case and with leading zeros (`'A1'`, `'ab12'`, `'A01'`, `'XFD1048576'`, `'RCC1'`); a name beyond the edge of the sheet is not a cell (`XFE1`, `A1048577`, `A0`), and neither is one with more characters after the cell (`A1B`, `AB12x`);
- that starts with an R1C1 reference containing a row or column number within the sheet, unless a `.` follows the reference: `'R2'`, `'R2D2'`, `'C3PO'`, `'R1x'`, `'RC1C'`, `'R1C0'` (a reference to the row, ignoring column 0), but `R0`, `R1048577`, `C16385x`, `C1.x` and `R1.` stay unquoted;
- that is, as a whole, an R1C1 reference without numbers: `'R'`, `'C'`, `'RC'`, `'R1C'`; with more characters after it, it stays unquoted (`Region`, `Cost`, `RCx`, `R_`);
- that is TRUE or FALSE in any case (`'TRUE'`, `'false'`) in the stored text (`Range.Formula`), or the locale's words in the text displayed in a locale (`Range.FormulaLocal`): `'ИСТИНА'` and `'ЛОЖЬ'` in ru-RU, where `TRUE` stays unquoted. `quoteSheetIfNeeded(name, locale)` accepts a locale object for its words, and uses en-US's without one.

Letters of other scripts that look like R and C do not form references (`Р1`, `С1`, `Ř1`, `К1С1`), and words that are not booleans stay unquoted (`TRUEx`, `FALSE1`). The engine's canonical text quotes every sheet (`'Sales'!A1`), and its tokenizer reads an unquoted name by the first rule above, so both Excel's text and the canonical text parse back. Text displayed to users quotes by these rules, as Excel displays it: `toLocal`, a viewport's `formulaLocal` and `editText`, and the text of `explain` show `=Sales!A1`, and a reference across sheets without quotes when neither end needs them (`Jan:Mar!A1`, but `'S1:S3'!A1`).

## The intersection and range operators

A space between two references is the intersection operator: the cells both cover. Line breaks count as spaces. A space next to an operator, an opening parenthesis or a comma is not an operator (`SUM( A1:A3 , B1 )`).

| formula | returns |
|---|---|
| `SUM(A1:C3 B2:D4)` | the sum of B2:C3 |
| `A1:C3 C2:C2` | C2's value |
| `A1:A3 C1:C3` | `#NULL!`: the ranges do not intersect |
| `ERROR.TYPE(A1:A3 C1:C3)`, `ISREF(A1:A3 C1:C3)` | 1, FALSE |
| `(A1:A3,C1:C3) A2:C2` | A2 and C2, a union |
| `A1:C3 INDEX(A1:D4,0,2)` | B1:B3: a call that returns a reference takes part |
| `A1:A3 'Jan'!A1:A3`, `'Jan:Mar'!A1:A3 'Jan:Mar'!A2:B2` | `#VALUE!`: different sheets, or references across sheets |

`:` between two references that are not a single range is the range operator: the smallest rectangle containing both. `A1:INDEX(B1:B9,3)` is A1:B3, `INDEX(B1:B9,2):B4` is B2:B4, and `Sales[Qty]:Sales[Price]` is the columns from Qty to Price. A reference without a sheet is on the formula's sheet, so `'D'!B1:B2:C3` combines two sheets and is `#VALUE!`, as `A1:'Jan'!A1` is. A spill reference can be an end (`A1:E2#` extends as far as the array spills).

A formula with the range operator reads every cell between the ends, so a formula in that range is calculated first, an edit that changes it reaches the formula that reads it, and a range around the formula's own cell is a cycle (`SUM(A1:A1:C1)` in B1). Where an end comes from a call, the cells in between are those around the references the call can return (for example `INDEX`'s array), not those it only reads, such as `MATCH`'s lookup range in `A1:INDEX(A1:A9,MATCH(5,Z1:Z9,0))`. A spill reference extends as far as its array spilled in the last calculation, and when the array grows, the formula that reads it is recalculated.

`:` binds tightest, then the space, then `@` and the other operators: `-A1:C3 C2:C2` is minus C2, and `@A1:C3 C2:C2` is C2. The comma of a union combines whole operands. The canonical text keeps one space for the intersection.

## Other functions

A function that does not accept several areas returns an error, which depends on the function. By default it is `#VALUE!` for a union, and for a reference across sheets `#REF!` where the function accepts an array (`SUMPRODUCT`, `MMULT`, `CORREL`) and `#VALUE!` elsewhere (the criteria functions, the database functions). The lookups return `#N/A` for a union, and `ROWS`, `COLUMNS`, `ROW`, `COLUMN`, `FILTER` and `CHOOSECOLS` return `#REF!`. `src/functions/areas.ts` lists the functions found to differ from the default.

## Structure operations

| operation | a reference across sheets |
|---|---|
| rows or columns inserted or deleted, on any sheet | stays unchanged |
| an end sheet renamed | follows the new name |
| an end sheet deleted | the next sheet inward becomes the end; with one sheet left it becomes a plain reference (`Feb:Mar!B1` without Feb becomes `Mar!B1`) |
| a sheet between the ends deleted | stays; the sheet is no longer read |
| a sheet inserted or moved between the ends | stays; the sheet is read |
| an end sheet moved | as if deleted from its position: the next sheet inward becomes the end |
| a copy | shifts by the copy's offset where it is relative |

The `insertSheet` and `moveSheet` operations of `editStructure` follow these rules.

## Checked against Excel

Calculated in Excel: a union and a reference across sheets in every argument of every whitelisted function that accepts a range (1,137 cases), with held-out sets for the functions that handle them differently; sheet names in references were tested in Excel for Mac and for Windows.
