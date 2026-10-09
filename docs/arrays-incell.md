# Arrays inside a formula

Operators and functions work on whole arrays and ranges, as in Excel 365 with dynamic arrays: `SUM(B1:B9*C1:C9)`, `SUM(IF(A1:A9>0,A1:A9))`, `INDEX(SORT(A1:A9),1)`. A formula whose final result is an array spills into neighbouring cells, which [spill.md](spill.md) describes.

Rules that go beyond Excel's documentation were measured in Excel 16.113 for Mac (en-US).

## Array values

An array is a rectangle of values stored row by row (`Arr` in `src/arrays/value.ts`). An element is anything a cell can hold: a number, text, a boolean, an error or empty. An error inside an array is a value like any other and does not stop the other elements from being computed.

In Excel a whole column has 1,048,576 rows, and so does `B:B*C:C`: `SUM(B:B+1)` adds 1,048,576. To keep this cheap, an array stores its occupied part from the top-left corner and one fill value for every position outside it. An operator over `B:B` and `C:C` computes the occupied rows and the fill once (`null*null` is 0). A function that needs plain values gets the array expanded.

An array of one element stays an array. An operator or a lifted call over it returns a 1×1 array: `TYPE({1}+0)` is 64, so `DEVSQ({1}>=TRUE)` skips the boolean and is `#NUM!`, and `SUM({TRUE}=TRUE)` is 0. `INDEX` of a line selected by 0 and functions with an array answer (`SORT(5)`, `VSTACK(5)`, `MMULT(2,3)`) return 1×1 arrays too. `TRANSPOSE` of a value or of one cell is that value, and `INDEX` of one element is its value. A reference to one cell is a value under an operator (`TYPE(A1:A1+0)` is 1). A cell shows a 1×1 array as its value.

A blank cell inside an array becomes an empty element (`TRANSPOSE(A1:A2)` over blanks). `ISBLANK` of it is TRUE, `COUNTA` counts it, unlike a blank cell of a range, and it reads as 0 or `""`. A function lifted over a blank cell returns 0 (`IF({1;1},A1:A2)`), and `INDEX` lifted over cells keeps them empty.

`asArray` in `src/arrays/value.ts` wraps scalar values and preserves existing arrays. `selectArray` is the common row/column selection used by the reshaping, sorting and filtering functions; each caller retains its allocation-limit checks. Operators and old-style array formulas share `broadcastAt` for repeating one row or column and detecting positions outside an operand.

## Element by element

An operator, or a function argument that takes one value, given an array or a range of several cells, is computed element by element:

- the result has as many rows and columns as the larger operand;
- an operand with one row or one column is repeated;
- any other position outside an operand is `#N/A`: `{1,2,3}+{1,2}` is `{2,4,#N/A}`.

Each function declares the kinds of its arguments (`src/args.ts`), and the kind decides what an array does there:

- arguments that take one value (`n i w x d t s b v a`) are lifted, so the function runs once per element;
- range and list arguments (`N A Y Z V G M R D`) take the whole array, as they take a range;
- `r` (MATCH's type) lifts an array and passes a range;
- `p` (IFERROR's alternative, FILTER's `if_empty`) and `P` (CHOOSE's values, the branches of IF, IFS and SWITCH, the arrays of HSTACK and VSTACK) are passed as they are unless the call is lifted. `P` passes a reference on as a reference, even one cell, so `SUM(IF(TRUE,A2))` skips text in A2 while `IF(TRUE,A2)+0` reads it.

So `SUM(B1:B9*C1:C9)` multiplies element by element and sums the products, `SUMIF(A:A,{"a","b"},B:B)` gives two sums, and `IF(A1:A9>0,A1:A9)` keeps the positive values. `IF`, `IFERROR`, `IFNA` and `SWITCH` evaluate only the branch they need, unless their condition or value is an array; then every argument is computed and the function is lifted.

Numeric list kinds share one collector inside `list` in `src/args.ts`. The collector selects the array or reference rule once per argument and uses the same conversion for stored values and the array's fill. The conversion rules remain distinct, including MAXA versus AVERAGEA and the handling of date lists. `test/args.test.ts` checks those differences and the order of errors when a fill lies between stored rows.

### Exceptions

Every whitelisted function's typical call was run with a two-cell range and a two-value array in each one-value argument. The exceptions:

- 108 functions refuse a range in a one-value argument with `#VALUE!` but lift an array: `EOMONTH(A1:A2,0)` is `#VALUE!`, while `EOMONTH(A1:A2+0,0)` is an array. These are the former Analysis ToolPak functions (BESSEL, radix and complex-number functions, bond, coupon and discount functions, EDATE, EOMONTH, WORKDAY, YEARFRAC and others) and the field of the database functions. WORKDAY.INTL lifts a range of weekends. The list is `REFUSE_RANGES` in `src/functions.ts`.
- `N` and `T` take the top-left cell of a range and lift an array.
- Given an array, `MAXA`, `MINA`, `STDEVA`, `STDEVPA`, `VARA` and `VARPA` take only its numbers, and `AVERAGEA` counts its text as 0. All of them skip booleans there.
- `VLOOKUP` and `HLOOKUP` take an array or range of column numbers as a whole: the answer has the found row's value at each number, shaped like the numbers, and the first invalid number decides for all (`#VALUE!` below 1, `#REF!` past the table). When the call is lifted by an array of keys, each call reads only the top-left number. `MATCH` reads the top-left cell of a range as its match type.
- `INDEX`, `TAKE`, `DROP`, `XLOOKUP` (`if_not_found` too), `CHOOSE`, `IF`, `IFS` and `SWITCH` return references when given references, one cell too: `SUMIF(INDEX(A1:C9,0,1),">2")` works, `COUNTA(TAKE(A1:A8,2))` and `COUNTA(IF(TRUE,A1))` skip blank cells, `COLUMN(CHOOSE(1,C1))` is 3. `IFERROR`, `IFNA`, FILTER's `if_empty` and unary `+` return values. Other functions with an array answer return arrays.
- `HSTACK` and `VSTACK` return the first error passed as a value, but stack references and arrays with the errors inside them.
- `SORTBY` with an array or a range of several cells as a sort order is not lifted: it answers an array as large as the orders, the first value of the sort by the orders' first values in the first place and blanks in the others, `#VALUE!` where an order is not 1 or −1 or has no value (an order's row or column is not repeated) and an order's error where it holds one (`SORTBY(A1:A6,B1:B6,{1;-1})` is the first value and a blank, `SORTBY(A1:A6,B1:B6,{1;2})` the first value and `#VALUE!`). This happens before a union or a reference across sheets among its arguments counts.

## Array constants

`{1,2;3,4}`: rows are separated by `;`, values by `,`, and all rows have the same length. Values are numbers (with a sign), text, TRUE, FALSE and error literals. References, expressions and nested arrays are not allowed, as in Excel.

## Functions with an array answer

`INDEX` with row or column 0 returns the whole column, row or range. `ROW` and `COLUMN` of several rows or columns return the list of numbers. TRANSPOSE, MMULT, MUNIT, MINVERSE, SEQUENCE, CHOOSECOLS, CHOOSEROWS, TAKE, DROP, EXPAND, HSTACK, VSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS, SORT, SORTBY, UNIQUE, FILTER, FREQUENCY, MODE.MULT, LINEST, LOGEST, TREND, GROWTH, TEXTSPLIT and RANDARRAY return arrays.

`RANDARRAY` is volatile and draws from the clock's seed, the cell and the element's position, as `RAND` does. Sorting is stable and orders numbers, text, booleans, errors, then blanks; blanks stay last in a descending sort too. Text is compared in the book's locale order (en-US puts Latin before Cyrillic). LINEST and its relatives solve by Householder QR.

An empty array, such as `FILTER` that keeps nothing, `TAKE` of zero rows or `DROP` of everything, is `#CALC!`, as in Excel.

## Compilation

The compiler knows statically whether a node of a formula can be an array (`src/arrays/shape.ts`): ranges of several cells where one value is expected, array constants, functions with an array answer, and anything lifted over them. A LET name has the shape of its value, and an omitted argument is one value. Only such nodes go through the lifting operators and callers; a formula without arrays compiles to plain scalar code. The shape pass costs about 5 ms on the 3,907 formulas of the synthetic book.

An array inside a formula exists only while the formula runs, so arrays add nothing to the reference graph. A session recomputes formulas with arrays like any others.

## How it was checked

The `arrays` case set (`bench/coverage/sets/arrays.ts`) is computed in Excel as dynamic array formulas. The harness writes each formula the way Excel 365 saves one: `<c cm="1"><f t="array" ref="A1">…</f>` plus `xl/metadata.xml` with the `XLDAPR` metadata and `dynamicArrayProperties fDynamic="1"`. Without that metadata Excel reads a formula the old way, with implicit intersection: `SUM(IF(D!B1:B8>0,D!B1:B8))` is 114.75 as an old formula and 118.25 as a dynamic one. The `_xlfn._xlws.` prefixes only name newer functions and do not make a formula dynamic.

A case whose answer is an array is checked through a wrapper that returns one value (SUM, INDEX, ROWS, COLUMNS, COUNTA, CONCAT, TEXTJOIN), or directly, where both Excel and the engine give `#SPILL!` because every neighbour of the case is occupied.

All case sets of the library are computed as dynamic array formulas, so the whole acceptance follows Excel 365.

The shared collector and array helpers were additionally compared with frozen build `69b4d1f` in both execution modes and locales: 2,600 complete results matched in values, issues and both hashes, including edits, whole-column fills, row/column selection, old-style arrays and conditional aggregates. Across 672 limit checks, the result or cell-limit error also matched at limits of 4, 20 and 100 values. These used local builds and the saved Excel fixtures, without a live Excel run.

On Node 26.7, 2026-10-09, seven alternating measured pairs after two warmup pairs (ten calls per sample) gave these medians for arrays of 34,000 elements. MAXA and AVERAGEA used mixed numbers, text and booleans; FILTER retained every third row; CHOOSECOLS selected the same column three times.

| Call | Before | Shared implementation |
|---|---:|---:|
| SUM | 0.332 ms | 0.347 ms |
| MAXA | 0.297 ms | 0.253 ms |
| AVERAGEA | 0.251 ms | 0.272 ms |
| FILTER | 0.295 ms | 0.248 ms |
| SORT | 3.564 ms | 3.429 ms |
| CHOOSECOLS | 1.064 ms | 0.940 ms |

The SUM and AVERAGEA microbenchmarks retain a small overhead. The collector stays local to the argument handler: moving it to a module-level function slowed mixed-array coercion substantially in the trial. V8's inlining trace showed the original per-element closure being inlined; the module-level collector remained a separate function.

Five alternating pairs of `synthBook({ seed: 1, rows: 34000 })`, with closures and fixed clock/seed, gave 1,561/1,547 ms for full compute and 3.71/3.53 ms for edit-A p95 (before/after). This workload showed no overall regression from the shared helpers.
