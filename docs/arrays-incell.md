# Arrays inside a cell

Design of array values inside a formula, before spilling. Reference semantics: Excel 365 with dynamic arrays (Excel 16.113 for Mac is one), en-US. The contract (`src/contract.ts`) does not change: every cell still holds one value.

## The array value

An array is a rectangle of values, `rows × cols`, stored row by row (`Arr` in `src/arrays/value.ts`). An element is any value a cell can hold: a number, text, a boolean, an error, or blank (`null`, from an empty cell of a referenced range; an operator reads it as 0, `""` or FALSE as usual). Errors inside an array are values and do not stop the computation of the other elements.

A whole column has 1,048,576 rows in Excel, and so does `B:B*C:C`. To keep that cheap the array keeps an occupied part, `h × w` values from the top-left corner, and one `fill` value for every position outside it, the same layout as a range's `Grid` (whose fill is blank). An operator applied to `B:B` and `C:C` computes the occupied rows and the fill once (`null*null` is 0). Consumers that read every position (SUM over the array, a lookup in it) see the fill; a function that needs plain values gets the array materialised.

A 1×1 array is a single value wherever it ends up.

## Lifting

An operator, or a function argument that takes one value, applied to an array or a range of several cells is computed element by element, with Excel's rules for sizes:

- the result has the larger number of rows and the larger number of columns of the operands;
- an operand with one row (one column) is repeated down (across);
- otherwise a position outside an operand is `#N/A` (`{1,2,3}+{1,2}` is `{2,4,#N/A}`).

Function arguments are lifted by kind (`src/args.ts`): the one-value kinds (`n i x d t s b v a`) are lifted, the range and list kinds (`N A Z V G M R D`) take the whole array as they take a range. So `SUM(B1:B9*C1:C9)` multiplies element by element and SUM reads the product as it reads a range (numbers only), `SUMIF(A:A,{"a","b"},B:B)` gives two sums, `IF(A1:A9>0,A1:A9)` keeps the positive ones. `IF`, `IFERROR`, `IFNA` and `SWITCH` stay lazy unless their first argument is an array; then every argument is computed and the function is lifted over all of them, as Excel does. `CHOOSE` with a single index returns the chosen argument as it is (a range stays a range).

Measured exception: functions of the former Analysis ToolPak (BESSEL, the radix conversions, the complex-number functions, EOMONTH and others; the list is in the code and comes from the `lift` cases of the array set) do not lift a range of several cells in a one-value argument: Excel 365 answers `#VALUE!` there, while an array (`D!B1:B2+0`, `{1;2}`) is lifted. `FnDef.rangeLift: false` marks them.

## A formula whose result is an array

Before spilling, a formula whose result is an array of more than one value gives `#SPILL!`, with an issue that names the size and the check code `unsupported` ("the result is an array of 3×1 values: spilling into neighbouring cells is not supported yet; wrap it in a function that takes an array, such as SUM or INDEX"). This is what Excel shows when the cells an array would spill into are not empty, so the engine behaves as Excel does on a sheet where every neighbour is taken; when spilling arrives, the same cells start spilling where there is room. The alternatives are worse: the top-left value hides the problem (the decision of 06.10), `#VALUE!` is Excel's answer to other things, `#ERROR!` loses the reason.

Where the result is an array regardless of the data (a range of several cells or an array constant reaching the top through operators and lifted arguments, `ROW(A1:A3)`), `validateFormula` already answers `unsupported` with the same message, as it did for array arithmetic before; `compute` gives `#SPILL!` there too. Where it depends on the data (`INDEX(A1:C3,0,B1)`, `IF(B1,A1:A3,0)`, `UNIQUE(A1:A9)`, `SEQUENCE(B1)`), the check happens when the cell is computed.

## Array constants

`{1,2;3,4}`: rows separated by `;`, values by `,`, all rows of the same length. Values are numbers (with a sign), text, TRUE/FALSE and error literals; no references, expressions or nested arrays, as in Excel. Canonical form prints them the same way. An array constant is built once per compiled formula.

## Functions with an array answer

`INDEX` with row or column 0 answers the whole column, row or range (a reference when given a reference, so it still works where a range is required); `ROW` and `COLUMN` of several rows or columns answer the list of numbers. New functions, usable inside a formula (`SUM(UNIQUE(...))`, `INDEX(SORT(...),1,1)`, `COUNTA(FILTER(...))`): TRANSPOSE, MMULT, MUNIT, MINVERSE, SEQUENCE, CHOOSECOLS, CHOOSEROWS, TAKE, DROP, EXPAND, HSTACK, VSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS, SORT, SORTBY, UNIQUE, FILTER, FREQUENCY, MODE.MULT, LINEST, LOGEST, TREND, GROWTH, TEXTSPLIT, RANDARRAY. `RANDARRAY` is volatile and draws from the clock seed, the cell and the element's position, as `RAND` does.

Excel answers `#CALC!` for an empty array (`FILTER` with nothing kept, `DROP` of everything). The contract has no such code; the engine answers `#VALUE!` (as for an over-long `CONCAT`), listed under Known differences.

## Graph and recomputation

Nothing changes in the graph: a cell still holds one value, the references of a formula are the references in its text, and an array lives only while its formula runs. Array constants have no references. A session recomputes a formula with arrays like any other; the session test runs edits with array formulas against full computations.

## Compilation and speed

The code generator knows statically whether a node can be an array (`src/arrays/shape.ts`): ranges of several cells in a value position, array constants, functions with array answers, and what is lifted over them. Only such nodes go through the lifting operators and callers; a formula without arrays compiles to the same code as before, so its speed does not change. The same pass gives `validateFormula` its check of a result that is always an array.

## Checking against Excel

The case set `bench/coverage/sets/arrays.ts` is computed in Excel as dynamic array formulas. The harness writes a case formula as Excel 365 saves one: `<c cm="1"><f t="array" ref="A1">…</f>` plus `xl/metadata.xml` with the `XLDAPR` metadata type and `dynamicArrayProperties fDynamic="1"` (the layout XlsxWriter also writes). A first run confirmed the difference: `SUM(IF(D!B1:B8>0,D!B1:B8))` is 114.75 as an old formula (implicit intersection) and 118.25 as a dynamic one, `ROW(D!C1:C8)` is 1 and `#SPILL!`. The `_xlfn._xlws.` prefixes only name newer functions and do not make a formula dynamic; a scalar wrapper alone (`SUM(...)`) does not either, an old formula still intersects inside it. A Ctrl+Shift+Enter array formula (what the coverage measurement used for array answers) evaluates the same way inside, but keeps the top-left value of an array answer instead of spilling, so it cannot check the rule above.

A case whose answer is an array is checked through a scalar wrapper (SUM, INDEX, ROWS, COLUMNS, COUNTA, CONCAT, TEXTJOIN), or at the top level, where both Excel and the engine give `#SPILL!`: in the case book every neighbour of a case is taken, and a value under the last case blocks a downward spill.

## Questions for Mikhail

1. `#CALC!`: add `CALC` to `ErrorCode` (a `contract:` change) so that empty arrays and over-long `CONCAT` answer as Excel does, or keep `#VALUE!` as a known difference?
2. The JSDoc of `CheckCode` in `src/contract.ts` says `unsupported` means "array arithmetic"; it now means a result that is an array (and external calls with non-literal arguments). A one-line `contract:` comment fix after your word.
3. Old measurement answers: the coverage measurement (and the `matheng` set before this branch) stored Excel's answers to formulas of the old kind. Where the old and the new semantics part (`ROW(D!C1:C8)` is 1 as an old formula, an array as a dynamic one), the engine follows Excel 365 and the case is listed in `KNOWN` with that reason, instead of re-measuring the coverage measurement. Re-measure them as dynamic formulas instead?
4. Spilling: this branch keeps every cell to one value. The next task (spilling, `E2#`) changes the contract; the issue text and `#SPILL!` stay as they are for blocked spills.
