# Arrays inside a cell

Design of array values inside a formula, before spilling. Reference semantics: Excel 365 with dynamic arrays (Excel 16.113 for Mac is one), en-US; every rule below that is not plain Excel documentation was measured by the `arrays` case set. The contract (`src/contract.ts`) does not change: every cell still holds one value.

## The array value

An array is a rectangle of values, `rows × cols`, stored row by row (`Arr` in `src/arrays/value.ts`). An element is any value a cell can hold: a number, text, a boolean, an error, or empty (`null`). Errors inside an array are values and do not stop the computation of the other elements.

A whole column has 1,048,576 rows in Excel, and so does `B:B*C:C` (`SUM(B:B+1)` adds 1,048,576). To keep that cheap the array keeps an occupied part, `h × w` values from the top-left corner, and one `fill` value for every position outside it, the same layout as a range's `Grid` (whose fill is blank). An operator over `B:B` and `C:C` computes the occupied rows and the fill once (`null*null` is 0); a function that needs plain values gets the array materialised.

A 1×1 array is its single value wherever one value is expected (`TYPE({1})` is still 64).

An empty element is what a blank cell becomes inside an array (`TRANSPOSE(A1:A2)` over blanks): `ISBLANK` is TRUE, `COUNTA` counts it (unlike a blank cell of a range), it reads as 0 or `""`. A function lifted over a blank cell answers 0 (`IF({1;1},A1:A2)`), INDEX lifted over cells keeps them empty.

## Lifting

An operator, or a function argument that takes one value, applied to an array or a range of several cells is computed element by element:

- the result has the larger number of rows and the larger number of columns of the operands;
- an operand with one row (one column) is repeated down (across);
- otherwise a position outside an operand is `#N/A` (`{1,2,3}+{1,2}` is `{2,4,#N/A}`).

Function arguments are lifted by kind (`src/args.ts`, `src/arrays/call.ts`): the one-value kinds (`n i x d t s b v a`) are lifted, the range and list kinds (`N A Y Z V G M R D`) take the whole array as they take a range, `r` (VLOOKUP's column number, MATCH's type) lifts an array and passes a range, `p` (CHOOSE's values, IF's branches, IFERROR's alternative) is passed as it is unless the call is lifted. So `SUM(B1:B9*C1:C9)` multiplies element by element and SUM reads the product as a range (numbers only), `SUMIF(A:A,{"a","b"},B:B)` gives two sums, `IF(A1:A9>0,A1:A9)` keeps the positive ones. `IF`, `IFERROR`, `IFNA` and `SWITCH` stay lazy unless their condition, value or expression (for SWITCH also a value to match) is an array; then every argument is computed and the function is lifted over all of them.

Measured exceptions, all from the `lift` cases, which give every whitelisted function's typical call a range of two cells and a two-value array in each one-value argument:

- 108 functions refuse a range in a one-value argument with `#VALUE!` but lift an array (`EOMONTH(A1:A2,0)` is `#VALUE!`, `EOMONTH(A1:A2+0,0)` an array): the former Analysis ToolPak (BESSEL, radix and complex-number functions, bond, coupon and discount functions, EDATE, EOMONTH, WORKDAY, YEARFRAC and others) and the field of the database functions; WORKDAY.INTL lifts a range of weekends. The list is `REFUSE_RANGES` in `src/functions.ts`.
- `N` and `T` take the top-left cell of a range and lift an array.
- In an array, `MAXA`, `MINA`, `STDEVA`, `STDEVPA`, `VARA` and `VARPA` take only the numbers, `AVERAGEA` counts text as 0 (kind `Y`); booleans are skipped by all of them.
- `VLOOKUP` with a range as the column number is `#REF!`, `MATCH` reads its top-left cell as the match type.
- `INDEX`, `TAKE`, `DROP`, `XLOOKUP` and `CHOOSE` answer references when given references (`SUMIF(INDEX(A1:C9,0,1),">2")`, `COUNTA(TAKE(A1:A8,2))` skips blank cells); the other functions with an array answer give arrays.

## A formula whose result is an array

Before spilling, a formula whose result is an array of more than one value gives `#SPILL!`, with an issue that names the size and has the check code `unsupported` ("the result is an array of 3×1 values: spilling into neighbouring cells is not supported yet; wrap it in a function that takes an array, such as SUM or INDEX"). This is what Excel shows when the cells an array would spill into are not empty, so the engine behaves as Excel does on a sheet where every neighbour is taken; when spilling arrives, the same cells start spilling where there is room. The alternatives are worse: the top-left value hides the problem (the decision of 06.10), `#VALUE!` is Excel's answer to other things, `#ERROR!` loses the reason.

`compute` finds this when the cell is computed, so a result that is an array only for some data (`INDEX(A1:C3,0,B1)`, `IF(B1,A1:A3,0)`, `UNIQUE(A1:A9)`, `SEQUENCE(B1)`) is treated the same as one that always is. `validateFormula` answers `unsupported` with the same message where the result is an array whatever the data (a range of several cells or an array constant reaching the top through operators and lifted arguments, `ROW(A1:A3)`), as it did for array arithmetic before.

## Array constants

`{1,2;3,4}`: rows separated by `;`, values by `,`, all rows of the same length. Values are numbers (with a sign), text, TRUE/FALSE and error literals; no references, expressions or nested arrays, as in Excel. The canonical form prints them the same way. A constant is built once per compiled formula.

## Functions with an array answer

`INDEX` with row or column 0 answers the whole column, row or range; `ROW` and `COLUMN` of several rows or columns answer the list of numbers. New functions, usable inside a formula: TRANSPOSE, MMULT, MUNIT, MINVERSE, SEQUENCE, CHOOSECOLS, CHOOSEROWS, TAKE, DROP, EXPAND, HSTACK, VSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS, SORT, SORTBY, UNIQUE, FILTER, FREQUENCY, MODE.MULT, LINEST, LOGEST, TREND, GROWTH, TEXTSPLIT, RANDARRAY. `RANDARRAY` is volatile and draws from the clock seed, the cell and the element's position, as `RAND` does (RAND's own values do not change). Sorting is stable and orders numbers, text, booleans, errors, then blanks, which stay last in a descending sort too; text compares with the en-US collation (Latin before Cyrillic, as Excel en-US sorts), which `<` and `>` now use as well. LINEST and its kin solve by Householder QR.

Excel answers `#CALC!` for an empty array (`FILTER` with nothing kept, `TAKE` of zero rows, `DROP` of everything). The contract has no such code; the engine answers `#VALUE!` (as for an over-long `CONCAT`), listed under Known differences.

## Graph and recomputation

Nothing changes in the graph: a cell still holds one value, the references of a formula are the references in its text, and an array lives only while its formula runs. A session recomputes a formula with arrays like any other; the session test runs edits with array formulas against full computations.

## Compilation and speed

The code generator knows statically whether a node can be an array (`src/arrays/shape.ts`): ranges of several cells in a value position, array constants, functions with an array answer, and what is lifted over them. Only such nodes go through the lifting operators (`L1`, `L2`) and callers (`FL`, `LZ`); a formula without arrays compiles to the same code as before. The shape pass costs about 5 ms on the synthetic book's 3,907 formulas.

## Checking against Excel

The case set `bench/coverage/sets/arrays.ts` is computed in Excel as dynamic array formulas. The harness writes a case formula as Excel 365 saves one: `<c cm="1"><f t="array" ref="A1">…</f>` plus `xl/metadata.xml` with the `XLDAPR` metadata type and `dynamicArrayProperties fDynamic="1"` (the layout XlsxWriter also writes). A first run confirmed the difference: `SUM(IF(D!B1:B8>0,D!B1:B8))` is 114.75 as an old formula (implicit intersection) and 118.25 as a dynamic one, `ROW(D!C1:C8)` is 1 and `#SPILL!`. The `_xlfn._xlws.` prefixes only name newer functions and do not make a formula dynamic; a scalar wrapper alone does not either, an old formula still intersects inside it (`SUM(D!B1:B8*2)` in row 2 is 16). A Ctrl+Shift+Enter array formula (what the coverage measurement used for array answers) evaluates the same way inside, but keeps the top-left value of an array answer instead of spilling, so it cannot check the rule above.

A case whose answer is an array is checked through a scalar wrapper (SUM, INDEX, ROWS, COLUMNS, COUNTA, CONCAT, TEXTJOIN), or at the top level, where both Excel and the engine give `#SPILL!`: in the case book every neighbour of a case is taken, and a value under the last case blocks a downward spill.

Older answers stored from formulas of the old kind (the coverage measurement, the `text` and `matheng` sets) part from Excel 365 where implicit intersection or an array result matters; those cases are listed in `KNOWN` with the reason `legacyFormula`, and the array set checks the same formulas as dynamic ones.

## Questions for Mikhail

1. `#CALC!`: add `CALC` to `ErrorCode` (a `contract:` change) so that empty arrays and over-long `CONCAT` answer as Excel does, or keep `#VALUE!` as a known difference?
2. The JSDoc of `CheckCode` in `src/contract.ts` says `unsupported` means "array arithmetic"; it now means a result that is an array (and external calls with non-literal arguments). A one-line `contract:` comment fix after your word.
3. Old-kind answers: 19 cases of the coverage measurement and of the `text` and `matheng` sets are in `KNOWN` because Excel computed them as old formulas. Re-measure those sets as dynamic formulas (`DYNAMIC = true`, a full rerun of each) so that they leave `KNOWN`, or keep them as they are?
4. The text order changed from ru-RU to en-US collation for `<`, `>`, criteria and lookups (Latin before Cyrillic), because Excel en-US sorts that way (`SORT`, and `{"x*y"}<"Аренда"` is TRUE in Excel). No existing case depended on the old order. Keep it?
