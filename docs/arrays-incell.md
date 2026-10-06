# Arrays inside a cell

Design of array values inside a formula; a formula whose result is an array spills, as `docs/spill.md` describes. Reference semantics: Excel 365 with dynamic arrays (Excel 16.113 for Mac is one), en-US; every rule below that is not plain Excel documentation was measured by the `arrays` case set.

## The array value

An array is a rectangle of values, `rows × cols`, stored row by row (`Arr` in `src/arrays/value.ts`). An element is any value a cell can hold: a number, text, a boolean, an error, or empty (`null`). Errors inside an array are values and do not stop the computation of the other elements.

A whole column has 1,048,576 rows in Excel, and so does `B:B*C:C` (`SUM(B:B+1)` adds 1,048,576). To keep that cheap the array keeps an occupied part, `h × w` values from the top-left corner, and one `fill` value for every position outside it, the same layout as a range's `Grid` (whose fill is blank). An operator over `B:B` and `C:C` computes the occupied rows and the fill once (`null*null` is 0); a function that needs plain values gets the array materialised.

A 1×1 array stays an array: an operator or a lifted call over it gives a 1×1 array (`TYPE({1}+0)` is 64, so `DEVSQ({1}>=TRUE)` skips the boolean and is `#NUM!`, `SUM({TRUE}=TRUE)` is 0), as do `INDEX` of a line selected by 0 (`INDEX({1;2},1,0)`) and the functions with an array answer (`SORT(5)`, `VSTACK(5)`, `MMULT(2,3)`); `TRANSPOSE` of a value or of one cell is that value, and `INDEX` of one element its value. A one-cell reference is a value under an operator (`TYPE(A1:A1+0)` is 1). A cell shows a 1×1 array as its value. Measured by the `1x1` cases of the `core` set.

An empty element is what a blank cell becomes inside an array (`TRANSPOSE(A1:A2)` over blanks): `ISBLANK` is TRUE, `COUNTA` counts it (unlike a blank cell of a range), it reads as 0 or `""`. A function lifted over a blank cell answers 0 (`IF({1;1},A1:A2)`), INDEX lifted over cells keeps them empty.

## Lifting

An operator, or a function argument that takes one value, applied to an array or a range of several cells is computed element by element:

- the result has the larger number of rows and the larger number of columns of the operands;
- an operand with one row (one column) is repeated down (across);
- otherwise a position outside an operand is `#N/A` (`{1,2,3}+{1,2}` is `{2,4,#N/A}`).

Function arguments are lifted by kind (`src/args.ts`, `src/arrays/call.ts`): the one-value kinds (`n i w x d t s b v a`) are lifted, the range and list kinds (`N A Y Z V G M R D`) take the whole array as they take a range, `r` (MATCH's type) lifts an array and passes a range, `p` (IFERROR's alternative, FILTER's `if_empty`) and `P` (CHOOSE's values, the branches of IF, IFS and SWITCH, HSTACK's and VSTACK's arrays) are passed as they are unless the call is lifted. `P` passes a reference on as a reference, one cell too; where such an answer goes to a one-value position, the walk reads it as a value (`read` of the backends), so `SUM(IF(TRUE,A2))` skips text in A2 while `IF(TRUE,A2)+0` reads it. So `SUM(B1:B9*C1:C9)` multiplies element by element and SUM reads the product as a range (numbers only), `SUMIF(A:A,{"a","b"},B:B)` gives two sums, `IF(A1:A9>0,A1:A9)` keeps the positive ones. `IF`, `IFERROR`, `IFNA` and `SWITCH` stay lazy unless their condition, value or expression (for SWITCH also a value to match) is an array; then every argument is computed and the function is lifted over all of them.

Measured exceptions, all from the `lift` cases, which give every whitelisted function's typical call a range of two cells and a two-value array in each one-value argument:

- 108 functions refuse a range in a one-value argument with `#VALUE!` but lift an array (`EOMONTH(A1:A2,0)` is `#VALUE!`, `EOMONTH(A1:A2+0,0)` an array): the former Analysis ToolPak (BESSEL, radix and complex-number functions, bond, coupon and discount functions, EDATE, EOMONTH, WORKDAY, YEARFRAC and others) and the field of the database functions; WORKDAY.INTL lifts a range of weekends. The list is `REFUSE_RANGES` in `src/functions.ts`.
- `N` and `T` take the top-left cell of a range and lift an array.
- In an array, `MAXA`, `MINA`, `STDEVA`, `STDEVPA`, `VARA` and `VARPA` take only the numbers, `AVERAGEA` counts text as 0 (kind `Y`); booleans are skipped by all of them.
- `VLOOKUP` and `HLOOKUP` take an array or a range of several cells as the index whole: the answer is the found line at each number, shaped as the index, and the first number in order that is not valid answers for all (`#VALUE!` below 1, `#REF!` past the table, an error or text that is not a number). A call lifted by an array key or match mode reads the index's top-left number instead (a 1 × 1 array of it: text or an error there is `#REF!`), so another number of the index does not count (the `lookup` cases of the `core` set). `MATCH` reads its top-left cell as the match type. Measured by the `index` cases.
- `INDEX`, `TAKE`, `DROP`, `XLOOKUP` (`if_not_found` too), `CHOOSE`, `IF`, `IFS` and `SWITCH` answer references when given references, one cell too (`SUMIF(INDEX(A1:C9,0,1),">2")`, `COUNTA(TAKE(A1:A8,2))` and `COUNTA(IF(TRUE,A1))` skip blank cells, `COLUMN(CHOOSE(1,C1))` is 3); `IFERROR`, `IFNA`, `FILTER`'s `if_empty` and unary `+` give values; the other functions with an array answer give arrays. `HSTACK` and `VSTACK` answer the first error given as a value and stack references and arrays, errors inside included (the `stack` and `ref` cases of the `core` set).

## A formula whose result is an array

It spills into the cells to its right and below, or gives `#SPILL!` with the cell in its way: `docs/spill.md`. Before spilling (until 07.10) such a formula gave `#SPILL!` everywhere, with the check code `unsupported`, and `validateFormula` refused a result that is an array whatever the data; both are gone. `compute` finds an array result when the cell is computed, so a result that is an array only for some data (`INDEX(A1:C3,0,B1)`, `SEQUENCE(B1)`) spills like one that always is.

## Array constants

`{1,2;3,4}`: rows separated by `;`, values by `,`, all rows of the same length. Values are numbers (with a sign), text, TRUE/FALSE and error literals; no references, expressions or nested arrays, as in Excel. The canonical form prints them the same way. A constant is built once per compiled formula.

## Functions with an array answer

`INDEX` with row or column 0 answers the whole column, row or range; `ROW` and `COLUMN` of several rows or columns answer the list of numbers. New functions, usable inside a formula: TRANSPOSE, MMULT, MUNIT, MINVERSE, SEQUENCE, CHOOSECOLS, CHOOSEROWS, TAKE, DROP, EXPAND, HSTACK, VSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS, SORT, SORTBY, UNIQUE, FILTER, FREQUENCY, MODE.MULT, LINEST, LOGEST, TREND, GROWTH, TEXTSPLIT, RANDARRAY. `RANDARRAY` is volatile and draws from the clock seed, the cell and the element's position, as `RAND` does (RAND's own values do not change). Sorting is stable and orders numbers, text, booleans, errors, then blanks, which stay last in a descending sort too; text compares with the en-US collation (Latin before Cyrillic, as Excel en-US sorts), which `<` and `>` now use as well. LINEST and its kin solve by Householder QR.

An empty array (`FILTER` with nothing kept, `TAKE` of zero rows, `DROP` of everything) is `#CALC!`, as in Excel (the error code `CALC` joined the contract for it).

## Graph and recomputation

An array inside a formula lives only while its formula runs; a spilled one puts its formula before the formulas reading its range (`docs/spill.md`). A session recomputes a formula with arrays like any other; the session tests run edits with array formulas and with spills against full computations.

## Compilation and speed

The compiler's walk (`src/exec/walk.ts`), shared by both execution modes, knows statically whether a node can be an array (`src/arrays/shape.ts`): ranges of several cells in a value position, array constants, functions with an array answer, and what is lifted over them. Only such nodes go through the lifting operators (`L1`, `L2`) and callers (`FL`, `LZ`), as closures or as generated text; a formula without arrays compiles to plain scalar code. The shape pass costs about 5 ms on the synthetic book's 3,907 formulas.

## Checking against Excel

The case set `bench/coverage/sets/arrays.ts` is computed in Excel as dynamic array formulas. The harness writes a case formula as Excel 365 saves one: `<c cm="1"><f t="array" ref="A1">…</f>` plus `xl/metadata.xml` with the `XLDAPR` metadata type and `dynamicArrayProperties fDynamic="1"` (the layout XlsxWriter also writes). A first run confirmed the difference: `SUM(IF(D!B1:B8>0,D!B1:B8))` is 114.75 as an old formula (implicit intersection) and 118.25 as a dynamic one, `ROW(D!C1:C8)` is 1 and `#SPILL!`. The `_xlfn._xlws.` prefixes only name newer functions and do not make a formula dynamic; a scalar wrapper alone does not either, an old formula still intersects inside it (`SUM(D!B1:B8*2)` in row 2 is 16). A Ctrl+Shift+Enter array formula (what the coverage measurement used for array answers) evaluates the same way inside, but keeps the top-left value of an array answer instead of spilling, so it cannot check the rule above.

A case whose answer is an array is checked through a scalar wrapper (SUM, INDEX, ROWS, COLUMNS, COUNTA, CONCAT, TEXTJOIN), or at the top level, where both Excel and the engine give `#SPILL!`: in the case book every neighbour of a case is taken, and a value under the last case blocks a downward spill (the engine's case books keep a value right of each case and under the last, `caseCells` in `test/excel-cases.ts`).

The coverage measurement's Excel columns (en-US and ru-RU) and the `text` and `matheng` sets are computed as dynamic array formulas too since 06.10. Measured again, 9 of the measurement's 5,275 answers changed in each column and 10 of the sets' 4,392, all where an old formula intersected a range with its row or showed the first value of an array (`ROW(D!C1:C8)` was 1, `TYPE(D!B1:B8)` 16, `INDEX(D!L1:N6,0,3)` `#VALUE!`); the 19 cases kept in `KNOWN` for that (`legacyFormula`) leave it: 18 match Excel as they are, `text:VLOOKUP#39` (`VLOOKUP(5,L!A1:B8,D!L1:L2)`, `#SPILL!`) after `VLOOKUP` and `HLOOKUP` learned an index array. The old columns were `docs/coverage-formulajs/excel-en-legacy.json` and `excel-ru-legacy.json`, removed since (in git at `85522cf`). The sets `criteria`, `findate`, `stat` and `subtotal` and the probes followed (06.10, the `parity/core` branch): of their 5,824 answers one changed, `AGGREGATE(14,6,T!M1:M10,T!A1:A3)`, `#VALUE!` by implicit intersection before and 9 now (a range as `k` gives its top-left cell, which the engine now does), so all acceptance is on Excel 365 formulas.

## Decisions taken (Mikhail, 06.10)

1. `#CALC!`: added to `ErrorCode` (a `contract:` commit after review of 06.10); empty arrays and over-long `CONCAT`, `TEXTJOIN` answer it.
2. The JSDoc of `CheckCode` `unsupported` says what it means now (the same `contract:` commit).
3. Old-kind answers: 19 cases of the coverage measurement and of the `text` and `matheng` sets stayed in `KNOWN` (`legacyFormula`) until those were measured again as dynamic formulas (done 06.10, see Checking against Excel).
4. The text order is en-US for `<`, `>`, criteria, lookups and sorting (Latin before Cyrillic, as Excel en-US); a locale of the book will set it in the locale work's second part.
