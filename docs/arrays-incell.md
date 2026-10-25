# Array operations in formulas

Operators and functions work on entire arrays and ranges, as in Excel 365 with dynamic arrays: `SUM(B1:B9*C1:C9)`, `SUM(IF(A1:A9>0,A1:A9))`, `INDEX(SORT(A1:A9),1)`. A formula whose final result is an array spills into neighboring cells, as described in [spill.md](spill.md).

Rules that go beyond Excel's documentation were verified in Excel 16.113 for Mac (en-US).

```js
import { compute } from 'xlcalc'

const cells = {
  A1: { value: 'a' }, A2: { value: 'b' }, A3: { value: 'a' },
  B1: { value: 2 }, B2: { value: 3 }, B3: { value: 4 },
  C1: { value: 10 }, C2: { value: 100 }, C3: { value: 1000 },
  D1: { formula: 'SUM(B1:B3*C1:C3)' },
  D2: { formula: 'SUM(IF(A1:A3="a",B1:B3))' },
  D3: { formula: 'TYPE({1}+0)' },
  E1: { formula: 'SUMIF(A1:A3,{"a","b"},B1:B3)' },
  E3: { formula: '{1,2,3}+{1,2}' },
}
const out = compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells
console.log(out.D1.value, out.D2.value, out.D3.value)
// → 4320 6 64
// SUMIF over two criteria gives two sums, spilled to F1; the shorter operand leaves #N/A
console.log([out.E1.value, out.F1.value], [out.E3.value, out.F3.value, out.G3.value])
// → [ 6, 3 ] [ 2, 4, { error: 'NA' } ]
```

## Array values

An array is a rectangle of values stored row by row (`Arr` in `src/arrays/value.ts`). An element can be anything a cell can hold: a number, text, a boolean, an error or empty. An error inside an array is a value like any other and does not stop the other elements from being calculated.

In Excel an entire column has 1,048,576 rows, and so does `B:B*C:C`: `SUM(B:B+1)` adds 1,048,576. To keep this cheap, an array stores its occupied part from the top-left corner and one fill value for every position outside it. An operator over `B:B` and `C:C` calculates the occupied rows and the fill once (`null*null` is 0). A function that needs plain values gets the array expanded.

An array of one element stays an array. An operator or a lifted call over it returns a 1×1 array: `TYPE({1}+0)` is 64, so `DEVSQ({1}>=TRUE)` skips the boolean and is `#NUM!`, and `SUM({TRUE}=TRUE)` is 0. `INDEX` of a row or column selected by 0 and functions that return arrays (`SORT(5)`, `VSTACK(5)`, `MMULT(2,3)`) return 1×1 arrays too. `TRANSPOSE` of a value or of one cell is that value, and `INDEX` of one element is its value. A reference to one cell is a value under an operator (`TYPE(A1:A1+0)` is 1). A cell displays a 1×1 array as its value.

A blank cell inside an array becomes an empty element (`TRANSPOSE(A1:A2)` over blanks). `ISBLANK` of it is TRUE, `COUNTA` counts it, unlike a blank cell of a range, and it is read as 0 or `""`. A function lifted over a blank cell returns 0 (`IF({1;1},A1:A2)`), and `INDEX` lifted over cells keeps them empty.

`asArray` in `src/arrays/value.ts` wraps scalar values and preserves existing arrays. `selectArray` is the common row/column selection used by the reshaping, sorting and filtering functions; each caller keeps its own allocation-limit checks. Operators and legacy array formulas share `broadcastAt` for repeating one row or column and detecting positions outside an operand.

## Element-wise calculation

An operator, or a function argument that expects a single value, given an array or a range of several cells, is calculated element-wise:

- the result has as many rows and columns as the larger operand;
- an operand with one row or one column is repeated;
- any other position outside an operand is `#N/A`: `{1,2,3}+{1,2}` is `{2,4,#N/A}`. A function gets `#N/A` there as that argument's value and evaluates it in turn, so an earlier argument can determine the result first: `IFERROR({1;2;3},{7;8})` is `{1;2;3}`, `IF({TRUE;FALSE;FALSE},{1;2},3)` is `{1;2;3}`, and `LEFT(A1:A3,{1;2})` with an error in A3 returns that error in the third row.

Each function declares the kinds of its arguments (`src/args.ts`), and the kind decides what an array does there:

- arguments that expect a single value (`n i w x d t s b v a`) are lifted, so the function runs once per element;
- range and list arguments (`N A Y Z V G M R D`) take the whole array, as they take a range;
- `r` (MATCH's match type) lifts an array and passes a range through;
- `p` (IFERROR's alternative, FILTER's `if_empty`) and `P` (CHOOSE's values, the branches of IF, IFS and SWITCH, the arrays of HSTACK and VSTACK) are passed as they are unless the call is lifted. `P` passes a reference on as a reference, even a single cell, so `SUM(IF(TRUE,A2))` skips text in A2 while `IF(TRUE,A2)+0` reads it.

So `SUM(B1:B9*C1:C9)` multiplies element-wise and sums the products, `SUMIF(A:A,{"a","b"},B:B)` returns two sums, and `IF(A1:A9>0,A1:A9)` keeps the positive values. `IF`, `IFERROR`, `IFNA` and `SWITCH` evaluate only the branch they need, unless their condition or value is an array; then every argument is calculated and the function is lifted.

Numeric list kinds share one collector inside `list` in `src/args.ts`. The collector selects the array or reference rule once per argument and uses the same conversion for stored values and the array's fill. The conversion rules remain distinct, including MAXA versus AVERAGEA and the handling of date lists. `test/args.test.ts` checks those differences and the order of errors when a fill lies between stored rows.

### Exceptions

Every whitelisted function's typical call was run with a two-cell range and a two-value array in each single-value argument. The exceptions:

- 109 functions reject a range in a single-value argument with `#VALUE!` but lift an array: `EOMONTH(A1:A2,0)` is `#VALUE!`, while `EOMONTH(A1:A2+0,0)` is an array. These are the former Analysis ToolPak functions (BESSEL, radix and complex-number functions, bond, coupon and discount functions, EDATE, EOMONTH, WORKDAY, YEARFRAC and others) and the field argument of the database functions. WORKDAY.INTL lifts a range of weekends. The list is `REFUSE_RANGES` in `src/functions.ts`.
- `N` and `T` take the top-left cell of a range and lift an array.
- Given an array, `MAXA`, `MINA`, `STDEVA`, `STDEVPA`, `VARA` and `VARPA` take only its numbers, and `AVERAGEA` counts its text as 0. All of them skip booleans there.
- `VLOOKUP` and `HLOOKUP` take an array or range of column numbers as a whole: the result has the found row's value at each number, shaped like the numbers, and the first invalid number decides the result for all (`#VALUE!` below 1, `#REF!` beyond the table). When the call is lifted by an array of lookup values, each call reads only the top-left number. `MATCH` reads the top-left cell of a range as its match type.
- `INDEX`, `TAKE`, `DROP`, `XLOOKUP` (including `if_not_found`), `CHOOSE`, `IF`, `IFS` and `SWITCH` return references when given references, even for a single cell: `SUMIF(INDEX(A1:C9,0,1),">2")` works, `COUNTA(TAKE(A1:A8,2))` and `COUNTA(IF(TRUE,A1))` skip blank cells, and `COLUMN(CHOOSE(1,C1))` is 3. `IFERROR`, `IFNA`, FILTER's `if_empty` and unary `+` return values. Other functions that return arrays return arrays.
- `HSTACK` and `VSTACK` return the first error passed as a value, but stack references and arrays with the errors inside them.
- `SORTBY` with an array or a range of several cells as a sort order is not lifted element-wise: it returns an array as large as the sort orders, which repeat as an operator's operands do (an order of one row or column is repeated along the other direction, and positions beyond a shorter one get `#N/A`). Each position, in row order, returns the error of the call with the orders at that position: an error passed as an argument, `#VALUE!` where an order is not 1 or −1 or is empty, or `#VALUE!` for a by-array whose size differs from the array. The first position whose call sorts returns its first sorted value, and later ones return a blank, or `#VALUE!` when `SORTBY` has more than one by-array (`SORTBY(A1:A6,B1:B6,{1;-1})` is the first value and a blank, `SORTBY(A1:A6,B1:B6,{1;2})` the first value and `#VALUE!`, `SORTBY(A1:A6,B1:B6,{2;-1})` `#VALUE!` and the first value of the descending sort, and `SORTBY(A1:A6,B1:B6,{1;-1},A1:A6,1)` the first value and `#VALUE!`). This is decided before a union or a reference across sheets among its arguments is considered.

## Array constants

`{1,2;3,4}`: rows are separated by `;`, values by `,`, and all rows have the same length. Values are numbers (with a sign), text, TRUE, FALSE and error literals. References, expressions and nested arrays are not allowed, as in Excel.

## Functions that return arrays

`INDEX` with row or column 0 returns the entire column, row or range. `ROW` and `COLUMN` of several rows or columns return the list of numbers. TRANSPOSE, MMULT, MUNIT, MINVERSE, SEQUENCE, CHOOSECOLS, CHOOSEROWS, TAKE, DROP, EXPAND, HSTACK, VSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS, SORT, SORTBY, UNIQUE, FILTER, FREQUENCY, MODE.MULT, LINEST, LOGEST, TREND, GROWTH, TEXTSPLIT and RANDARRAY return arrays.

`RANDARRAY` is volatile and draws from the clock's seed, the cell and the element's position, as `RAND` does. Sorting is stable and orders numbers, text, booleans, errors, then blanks; blanks stay last in a descending sort too. Text is compared in the workbook's locale order (en-US puts Latin before Cyrillic). LINEST and its relatives solve by Householder QR, centered on the means of y and the variables when the model has a constant.

An empty array, such as a `FILTER` that keeps nothing, `TAKE` of zero rows or `DROP` of everything, is `#CALC!`, as in Excel.

An array can have more columns than a sheet: `COLUMNS(SEQUENCE(1,16385))` is 16,385, as in Excel. More rows than a sheet has are refused as Excel refuses them: `SEQUENCE` and `RANDARRAY` of 1,048,577 rows are `#VALUE!`, and `VSTACK` or `TOCOL` that would make them `#NUM!`. Whatever its shape, an array of more than 134,217,728 values (2^27, what one JavaScript array holds) is `#VALUE!`.

## Compilation

The compiler knows statically whether a node of a formula can be an array (`src/arrays/shape.ts`): ranges of several cells where a single value is expected, array constants, functions that return arrays, and anything lifted over them. A LET name has the shape of its value, and an omitted argument is a single value. Only such nodes go through the lifting operators and callers; a formula without arrays compiles to plain scalar code. The shape pass takes about 5 ms on the 3,907 formulas of the synthetic workbook.

An array inside a formula exists only while the formula runs, so arrays add nothing to the reference graph. A session recalculates formulas with arrays like any others.

## Checked against Excel

Every test case set of the library is calculated in Excel as dynamic array formulas, so the rules above are Excel 365's. Held-out sets, each predicted before Excel calculated it, settled `SORTBY` and `LOOKUP`; one known difference remains: a `LOOKUP` placed among the cells it may read returns `#CYCLE!`, where Excel calculates it.
