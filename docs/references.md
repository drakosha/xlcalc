# The @ operator, unions, references across sheets, sheet names and the reference operators

Kinds of references that Excel files use. `@A1:A10` takes one cell of a range. `(A1:A3,C1:C3)` joins ranges of one sheet. `'Jan:Mar'!A1:B2` is the same rectangle on every sheet from Jan to Mar. `A1:C3 B2:D4` is the cells two ranges share, and `A1:INDEX(B1:B9,3)` the range from one reference to another.

The rules below were measured in Excel 16.113 for Mac (en-US) with dynamic array formulas.

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

`@x` is implicit intersection: one value where `x` could give several.

| x | `@x` gives |
|---|---|
| a column `A1:A10` | the cell in the formula's row, `#VALUE!` when the row is outside it |
| a row `C1:E1` | the cell in the formula's column |
| a block `G1:H2` | the cell in the formula's row and column |
| a whole column `A:A`, a whole row `1:1` | the cell in the formula's row or column |
| a range of another sheet | the cell by the same row and column there |
| a spill reference `C3#` | the cell of the spill range in the formula's row |
| an array, a call with an array answer | its top-left value (`@{7,8}` is 7, `@SEQUENCE(3)` is 1) |
| a call answering a reference (`INDEX(A1:B10,0,1)`) | the cell of that reference |
| a union | `#VALUE!` |
| a reference across sheets | `#REF!` |
| one value | that value |

The answer is a reference to that cell, so `SUM(@A1:A10)` skips text there, `ROW(@A1:A10)` is the formula's row and `ISREF(@A1:A10)` is TRUE.

`@` binds to what follows it: a reference, a name, a call, a constant or a parenthesis. `@A1:A10*2` is `(@A1:A10)*2`. An xlsx file stores it as `_xlfn.SINGLE(A1:A10)`, and `readXlsx` and `writeXlsx` turn one into the other.

## Unions

A union is written in parentheses inside a call: `SUM((A1:A3,C1:C3))`. A name's formula may be a union without them. All parts must be references on one sheet; otherwise the union is `#VALUE!`.

| formula | gives |
|---|---|
| `SUM((A1:A3,C1:C3))`, `COUNT`, `MAX`, `AVERAGE`, `LARGE`, `MEDIAN`, `STDEV` | the function over every part |
| `SUBTOTAL(9,(A1:A3,C1:C3))`, `AGGREGATE(9,6,(A1:A3,C1:C3))` | the function over every part, nested `SUBTOTAL` cells and hidden rows skipped as in one range; an area counted twice where parts overlap |
| `FREQUENCY((A1:A3,C1:C3),E1:E3)`, a union of bins | the counts over every part's numbers |
| `SORTBY((A1:A3,C1:C3),…)`, a union as a sort order too | `#VALUE!` |
| `INDEX((A1:A3,C1:C3),2,1,2)` | C2, from the second area |
| `INDEX((A1:A3,C1:C3),2,1,3)` | `#REF!`, no third area |
| `AREAS((A1:A3,C1:C3))` | 2 |
| `VSTACK((A1:A3,C1:C3))`, `TOCOL` | the parts one after another |
| `ROWS((A1:A3,C1:C3))` | `#REF!` |
| `(A1:A3,C1:C3)` alone, `+1` | `#VALUE!` |

## References across sheets

`'Jan:Mar'!A1:B2` covers the sheets from Jan to Mar in the book's order, Feb between them included. The ends can come in either order. Sheet names that read as cell addresses (`S1`, `Q3`) must be quoted there: Excel reads `S1:S3!A1` as the range operator between the cell S1 of the formula's own sheet and A1 on S3, so it is `#VALUE!` on any other sheet and the range A1:S1 on S3. One such sheet alone needs no quotes (`S1!A1`).

The list functions take each sheet's rectangle in turn: `SUM`, `COUNT`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`, `PRODUCT`, `STDEV`, `VAR` and the others over lists of numbers, `AND`, `OR`, and `FREQUENCY`, as its data or its bins. `VSTACK`, `HSTACK`, `TOCOL` and `TOROW` stack the sheets' rectangles. `CHOOSE` and `IF` pass one on.

Alone or under an operator it is `#REF!` (`=Jan:Mar!A1+1`). `INDEX`, `AREAS`, `ROWS`, `SUBTOTAL` and `AGGREGATE` give `#VALUE!`, `SORTBY` and a value `SWITCH` compares give `#REF!`, `ISREF` gives FALSE.

## Sheet names in references

A formula shows a sheet's name before its reference in quotes only where the name needs them, and `quoteSheetIfNeeded` writes a name the way Excel does. Excel quotes a name:

- that is not a letter or `_` followed by letters, digits, `_` and `.`: `'Q1 data'`, `'2024'`, `'Sheet-1'`, `'€uro'`, `'.a'`; letters of any script count (`Лист1`, `Straße`, `日本`, `a.b`, `_x` stay as they are), and an apostrophe inside is doubled (`'O''Brien'`);
- that is a cell of the sheet in A1 notation, in either case and with leading zeros (`'A1'`, `'ab12'`, `'A01'`, `'XFD1048576'`, `'RCC1'`); a name past the sheet's edge is no cell (`XFE1`, `A1048577`, `A0`), nor one with more after the cell (`A1B`, `AB12x`);
- that begins with an R1C1 reference holding a row or column number of the sheet, unless a `.` follows the reference: `'R2'`, `'R2D2'`, `'C3PO'`, `'R1x'`, `'RC1C'`, `'R1C0'` (the row's reference, the column 0 aside), but `R0`, `R1048577`, `C16385x`, `C1.x` and `R1.` stay as they are;
- that is an R1C1 reference without numbers as a whole: `'R'`, `'C'`, `'RC'`, `'R1C'`; with more after it, it stays (`Region`, `Cost`, `RCx`, `R_`);
- that is TRUE or FALSE in any case (`'TRUE'`, `'false'`) in the stored text (`Range.Formula`), and the locale's words in the text shown in a locale (`Range.FormulaLocal`): `'ИСТИНА'` and `'ЛОЖЬ'` in ru-RU, where `TRUE` stays as it is. `quoteSheetIfNeeded(name, locale)` takes a locale object for its words, en-US's without one.

Letters of other scripts that look like R and C are no references (`Р1`, `С1`, `Ř1`, `К1С1`), and words that are no boolean stay as they are (`TRUEx`, `FALSE1`). The engine's canonical text quotes every sheet (`'Sales'!A1`), and its tokenizer reads a name without quotes by the first rule above, so both Excel's text and the canonical one read back. The text shown to a person quotes by these rules, as Excel shows it: `toLocal`, a window's `formulaLocal` and `editText` and the text of `explain` show `=Sales!A1`, and a reference across sheets without quotes when neither end needs them (`Jan:Mar!A1`, but `'S1:S3'!A1`).

## The intersection and the range operator

A space between two references is the intersection operator: the cells both cover. Line breaks count as spaces. A space next to an operator, a parenthesis that opens or a comma is no operator (`SUM( A1:A3 , B1 )`).

| formula | gives |
|---|---|
| `SUM(A1:C3 B2:D4)` | the sum of B2:C3 |
| `A1:C3 C2:C2` | C2's value |
| `A1:A3 C1:C3` | `#NULL!`: the ranges do not meet |
| `ERROR.TYPE(A1:A3 C1:C3)`, `ISREF(A1:A3 C1:C3)` | 1, FALSE |
| `(A1:A3,C1:C3) A2:C2` | A2 and C2, a union |
| `A1:C3 INDEX(A1:D4,0,2)` | B1:B3: a call that answers a reference takes part |
| `A1:A3 'Jan'!A1:A3`, `'Jan:Mar'!A1:A3 'Jan:Mar'!A2:B2` | `#VALUE!`: other sheets, references across sheets |

`:` between two references that are not one range is the range operator: the rectangle around both. `A1:INDEX(B1:B9,3)` is A1:B3, `INDEX(B1:B9,2):B4` is B2:B4, `Sales[Qty]:Sales[Price]` the columns from Qty to Price. A reference without a sheet is on the formula's sheet, so `'D'!B1:B2:C3` joins two sheets and is `#VALUE!`, as `A1:'Jan'!A1` is. A spill reference can be an end (`A1:E2#` reaches as far as the array spills).

A formula with the range operator reads every cell between the ends, so a formula there is computed first, an edit that changes it reaches the reader, and a range around the formula's own cell is a cycle (`SUM(A1:A1:C1)` in B1). Where an end comes from a call, the cells between are those around the references the call can answer, `INDEX`'s array say, not those it only reads, such as `MATCH`'s lookup range in `A1:INDEX(A1:A9,MATCH(5,Z1:Z9,0))`. A spill reference reaches as far as its array spilled at the last computation, and when the array grows the reader is computed again.

`:` binds tightest, then the space, then `@` and the other operators: `-A1:C3 C2:C2` is minus C2, `@A1:C3 C2:C2` is C2. The comma of a union joins whole operands. The canonical text keeps one space for the intersection.

## Other functions

A function that does not take several areas gives an error, which depends on the function. By default it is `#VALUE!` for a union and `#REF!` for a reference across sheets where the function takes an array (`SUMPRODUCT`, `MMULT`, `CORREL`), and `#VALUE!` elsewhere (the criteria functions, the database functions). The lookups give `#N/A` for a union, `ROWS`, `COLUMNS`, `ROW`, `COLUMN`, `FILTER` and `CHOOSECOLS` give `#REF!`. `src/functions/areas.ts` lists the functions measured to differ from the default.

## Structure operations

| operation | a reference across sheets |
|---|---|
| rows or columns inserted or deleted, on any sheet | stays as it is |
| an end renamed | follows the new name |
| an end deleted | the next sheet inward is the end; one sheet left gives a plain reference (`Feb:Mar!B1` without Feb is `Mar!B1`) |
| a sheet between the ends deleted | stays; the sheet is no longer read |
| a sheet inserted or moved between the ends | stays; the sheet is read |
| an end moved | as if deleted from its place: the next sheet inward is the end |
| a copy | moves by the copy's offset where it is relative |

`insertSheet` and `moveSheet` of `editStructure` do this.

## Checked against Excel

Computed in Excel: a union and a reference across sheets in every argument of every whitelisted function that takes a range (1,137 cases), with held-out sets for the functions that take them otherwise; sheet names in references were measured in Excel for Mac and for Windows.
