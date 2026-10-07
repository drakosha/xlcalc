# Formulas of the old kind

Files saved before dynamic arrays hold formulas of the old kind. Excel 365 computes them with implicit intersection and shows them with `@`. The engine computes every formula as a dynamic one, so a reader turns an old formula into the dynamic formula Excel shows for it with `fromLegacy`. Array formulas entered with Ctrl+Shift+Enter keep their own kind in the book: `CellInput.array`.

The rules below were measured in Excel 16.113 for Mac (en-US).

## fromLegacy

```js
import { fromLegacy } from 'xlcalc'

fromLegacy('A1:A10*2')                // → '@A1:A10*2'
fromLegacy('SUM(A1:A10*2)')           // → 'SUM(@A1:A10*2)'
fromLegacy('SUMPRODUCT(A1:A10*2)')    // → 'SUMPRODUCT(A1:A10*2)'
fromLegacy('INDEX(A1:C3,0,1)+1')      // → '@INDEX(A1:C3,0,1)+1'
fromLegacy('rng*2', { names, sheet }) // → '@rng*2' when rng is a range
```

An xlsx cell formula without dynamic array metadata (`cm`) is of the old kind. `fromLegacy` takes its text after the prefixes of newer functions are dropped and returns canonical text. A formula it does not read comes back as it is.

`readXlsx` does this for every formula of the old kind in a file and keeps the text as written where no `@` is needed. It reads an array formula without the metadata into `CellInput.array`, and `writeXlsx` writes it back so.

## Where Excel puts @

An old formula is computed by values. Where one value goes, a range gives the cell in the formula's row or column, and an array gives its top-left value. Only an argument that takes an array keeps a range whole.

| old formula | Excel 365 shows |
|---|---|
| `A1:A10`, `A:A`, `A1:A1` | `@A1:A10`, `@A:A`, `@A1:A1` (a single cell stays `A1`) |
| `ABS(A1:A10)`, `IF(A1:A10>5,1,0)` | the range with `@` |
| `SUM(A1:A10)`, `SUMIF(A1:A10,">3")` | as it is: these arguments take a reference |
| `SUM(A1:A10*2)` | `SUM(@A1:A10*2)`: an expression is computed by values there |
| `SUMPRODUCT(A1:A10*2)`, `ROWS(A1:A10*2)` | as it is: these arguments take an array |
| `SUMPRODUCT(IF(A1:A10>5,1,0))` | `SUMPRODUCT(IF(@A1:A10>5,1,0))`: IF's condition stops the array |
| `={1,2,3}*2`, `ABS({-1,-2})` | `@{1,2,3}*2`, `ABS(@{-1,-2})` |
| `SUM(ABS({-1,-2}))` | as it is: an array stays an array inside SUM |
| `SEQUENCE(3)`, `TRANSPOSE(A1)`, `IFS(…)`, `SWITCH(…)` | `@` before the call |
| `INDEX(A1:C3,2,2)`, `INDEX(A1:C3,2)` | as it is: one cell |
| `INDEX(A1:C3,0,1)`, `INDEX(A1:C3,B1,1)` | `@INDEX(…)`: a row or column index that is not a positive number may give several |
| `IF(TRUE,A1:A10)`, `CHOOSE(1,A1:A10)` | `@IF(…)`, `@CHOOSE(…)`: the call can answer a range |
| `ROW(A1:A3)`, `SUM(ROW(A1:A3))` | `@ROW(…)`, `SUM(@ROW(…))`: ROW gives several values only inside an array |
| `LET(x,A1:A10,x)`, `LET(x,A1:A10,x*2)` | `@LET(…)`, `LET(x,A1:A10,@x*2)` |
| a name over a range or an array | `@name`; a name of one value stays |

Each argument of each whitelisted function has its class: one value, a reference, or an array, and whether an array evaluation goes on into it. `src/legacy/params.ts` holds the classes, written from Excel's text of the `legacy/params` and `legacy/nesting` sets by `bench/coverage/legacy/table.ts`.

## Array formulas of the old kind

A cell with `array: 'B2:C5'` holds an array formula entered over B2:C5. The formula is computed as an array, without implicit intersection, and fills its area:

| answer | B2:C5 shows |
|---|---|
| an array of the area's size | the array |
| a smaller array of one row or one column | that row or column repeated |
| a smaller array of several rows and columns | the array, `#N/A` past it |
| a larger array | its top-left part |
| one value | that value in every cell |

`ROW()` and `COLUMN()` give the rows and columns of the area: `{=ROW()}` over E1:E3 shows 1, 2, 3. A single cell (`array: 'B2'`) shows the top-left value: `{=A1:A3}` shows A1, where an old plain formula would intersect.

The area never spills and never gives `#SPILL!`. A dynamic array that meets it gives `#SPILL!`. In the result the formula's cell has `spill`, the area, and the other cells `spillOf`. `B2#` refers to the area.

The other cells of the area hold no input. A session refuses edits that change some cells of an area without the rest: `setCells` throws `InputError` and changes nothing, as Excel does. Clearing every cell of the area removes the formula. An edit of the formula's cell that keeps an array over the same area changes the whole array.

`restructure` moves an area with its cell. It refuses with `array` to insert lines inside an area, to delete part of one, to move part of one and to paste over part of one. A copy of a whole area copies the array; a copy of part of an area pastes each cell as an array formula of its own, the area's formula shifted from the area's cell.

## How it was checked

- `legacy/params` (3,071 formulas): every whitelisted function with a range and an expression over a range put into each of its arguments, read back as Excel 365 shows them (`formula2`) and computed.
- `legacy/nesting` (3,469) and `legacy/blocks` (346): the same inside SUMPRODUCT, SUM, INDEX, ROWS and MMULT.
- `legacy/calls` (113): calls whose answer depends on their arguments, array constants, names.
- The coverage measurement computed with formulas of the old kind: every whitelisted case gives Excel's value at its row.
- `spill/legacy` and `spill/legacy-held`: intersection by row and column, on other sheets and through names. The 13 cells of `legacy-held` were predicted before Excel computed them, and all came out as predicted.
- `spill/cse`, `spill/cse-held` and `struct/arrays` for array formulas of the old kind. The 25 cells of `cse-held` were predicted before Excel computed them, all as predicted.
- `xlsx intersect`: 75 formulas of the old kind in a file Excel computed and saved. Read by `readXlsx`, each computes the value Excel stored, but two with `OFFSET`, which is not in the whitelist, and one with the intersection operator.

`fromLegacy` gives Excel's text for every formula of these sets but one, which uses the intersection operator.

## Known differences

- The intersection operator (a space between two ranges) is not read.
- Excel computes an old formula by its old rules, and in a few places the dynamic text it shows computes otherwise: a range where an argument takes one value and Excel shows no `@` (GROWTH's and TREND's last argument, MATCH's type) is an error in the old formula and lifted in the dynamic one. The engine gives the dynamic formula's value.
- Inside SUMPRODUCT Excel does not lift ACCRINT, CUMIPMT, PRICE, XIRR, GROWTH and TREND over an array in a one-value argument; the engine lifts them.
- An error as NETWORKDAYS.INTL's weekend is `#VALUE!` in Excel and `#NUM!` here; XIRR's guess given a range is `#VALUE!` in Excel and lifted here.
