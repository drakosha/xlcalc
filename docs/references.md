# The @ operator, unions, references across sheets and the reference operators

Kinds of references that Excel files use. `@A1:A10` takes one cell of a range. `(A1:A3,C1:C3)` joins ranges of one sheet. `'Jan:Mar'!A1:B2` is the same rectangle on every sheet from Jan to Mar. `A1:C3 B2:D4` is the cells two ranges share, and `A1:INDEX(B1:B9,3)` the range from one reference to another.

The rules below were measured in Excel 16.113 for Mac (en-US) with dynamic array formulas.

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

`:` between two references that are not one range is the range operator: the rectangle around both. `A1:INDEX(B1:B9,3)` is A1:B3, `INDEX(B1:B9,2):B4` is B2:B4, `Sales[Qty]:Sales[Price]` the columns from Qty to Price. A reference without a sheet is on the formula's sheet, so `'D'!B1:B2:C3` joins two sheets and is `#VALUE!`, as `A1:'Jan'!A1` is.

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

`insertSheet` and `moveSheet` of `restructure` do this.

## How it was checked

The set `areas` (1,137 cases) puts a union and a reference across sheets into every argument of every whitelisted function that takes a range, and adds the forms above; its case ids name the function and the argument (`fn:union:SUBTOTAL#2`). The set `areas-each` (52 cases) measures how `SUBTOTAL`, `AGGREGATE`, `FREQUENCY` and `SORTBY` take them: 40 of its 47 predicted answers came out as predicted, and the misses were all `SORTBY`, which takes neither. The set `intersect` (44 cases) measures the intersection and the range operator, 40 of its 41 predicted answers as predicted (the miss: `'D'!B1:B2:C3`, whose C3 is on the formula's sheet), and `sheet-cells` (13 cases) the sheets named like cells, all 11 predictions as predicted. The spill sets `at` (8 books) and `at-held` measure `@`. The struct set `cube` (20 operations) measures structure operations, sheets inserted and moved included. The sets `at-held` (7 cells) and `cube-held` (14 cells) were predicted before Excel computed them, and all cells came out as predicted.
