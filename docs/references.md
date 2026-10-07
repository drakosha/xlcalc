# The @ operator, unions and references across sheets

Three kinds of references that Excel files use. `@A1:A10` takes one cell of a range. `(A1:A3,C1:C3)` joins ranges of one sheet. `'Jan:Mar'!A1:B2` is the same rectangle on every sheet from Jan to Mar.

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
| one value | that value |

The answer is a reference to that cell, so `SUM(@A1:A10)` skips text there, `ROW(@A1:A10)` is the formula's row and `ISREF(@A1:A10)` is TRUE.

`@` binds to what follows it: a reference, a name, a call, a constant or a parenthesis. `@A1:A10*2` is `(@A1:A10)*2`. An xlsx file stores it as `_xlfn.SINGLE(A1:A10)`.

## Unions

A union is written in parentheses inside a call: `SUM((A1:A3,C1:C3))`. A name's formula may be a union without them. All parts must be references on one sheet; otherwise the union is `#VALUE!`.

| formula | gives |
|---|---|
| `SUM((A1:A3,C1:C3))`, `COUNT`, `MAX`, `AVERAGE`, `LARGE`, `MEDIAN`, `STDEV` | the function over every part |
| `INDEX((A1:A3,C1:C3),2,1,2)` | C2, from the second area |
| `INDEX((A1:A3,C1:C3),2,1,3)` | `#REF!`, no third area |
| `AREAS((A1:A3,C1:C3))` | 2 |
| `VSTACK((A1:A3,C1:C3))`, `TOCOL` | the parts one after another |
| `ROWS((A1:A3,C1:C3))` | `#REF!` |
| `(A1:A3,C1:C3)` alone, `+1` | `#VALUE!` |

## References across sheets

`'Jan:Mar'!A1:B2` covers the sheets from Jan to Mar in the book's order, Feb between them included. The ends can come in either order. Sheet names that read as cell addresses (`S1`, `Q3`) must be quoted: Excel does not read `S1:S3!A1` as a reference across sheets.

The list functions take each sheet's rectangle in turn: `SUM`, `COUNT`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`, `PRODUCT`, `STDEV`, `VAR` and the others over lists of numbers, `AND`, `OR`. `VSTACK`, `HSTACK`, `TOCOL` and `TOROW` stack the sheets' rectangles. `CHOOSE` and `IF` pass one on.

Alone or under an operator it is `#REF!` (`=Jan:Mar!A1+1`). `INDEX`, `AREAS` and `ROWS` give `#VALUE!`, `ISREF` gives FALSE.

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

The set `areas` (989 cases) puts a union and a reference across sheets into every argument of every whitelisted function that takes a range, and adds the forms above. The spill sets `at` (8 books) and `at-held` measure `@`. The struct set `cube` (20 operations) measures structure operations, sheets inserted and moved included. The held-out sets `at-held` (7 cells) and `cube-held` (14 cells) were predicted before Excel computed them, and all cells came out as predicted.

## Known differences

- `SUBTOTAL` and `AGGREGATE` take each area of a union in Excel and give `#VALUE!` here.
- `FREQUENCY` and `SORTBY` take a union or a reference across sheets in Excel and answer an array; here they give `#VALUE!` or `#REF!`.
- A reference across sheets that `SWITCH` compares is `#REF!` in Excel and `#VALUE!` here.
- The intersection operator, a space between two ranges, is not read.
