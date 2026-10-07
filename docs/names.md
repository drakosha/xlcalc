# Defined names

A defined name stands for a formula, as Excel's "Refers to" does. The formula can be a reference, a constant, a calculation or a union of references. A formula that uses the name gets what the name's formula gives at the cell where it is used.

```js
names: [
  { name: 'rate', ref: '0.2' },
  { name: 'col', ref: "INDEX('Data'!$A$1:$C$9,0,2)" },
  { name: 'both', ref: "'Data'!$A$1:$A$3,'Data'!$C$1:$C$3" },
  { name: 'left', ref: "'Data'!XFD1" },
]
```

The rules below were measured in Excel 16.113 for Mac (en-US).

## What a name can hold

| `ref` | used as | gives |
|---|---|---|
| `0.2`, `"text"`, `TRUE`, `#N/A` | `=rate*100` | the constant |
| `{1,2,3}` | `=arr` | the array, which spills; `SUM(arr)` is 6, `COLUMNS(arr)` is 3 |
| `'Data'!$A$2*2` | `=dbl` | the value of the calculation |
| `'Data'!$A$1:$A$3*10` | `=SUM(sq)`, `=sq` | an array: summed, or spilled |
| `INDEX('Data'!$A$1:$C$9,0,2)` | `=SUMIF(col,">1")` | a reference: functions that take ranges take it |
| `'Data'!$A$1:$A$3,'Data'!$C$1:$C$3` | `=SUM(both)` | a union: see [References](references.md) |
| `'Data'!$B$1#` | `=SUM(sp)` | the spill range of B1 |

A name's formula is written as a file stores it, without `=`. Every reference in it has its sheet. A name whose formula does not parse, or has a reference without a sheet, is left out with a name issue.

## References without `$`

A reference without `$` is relative to A1. Used in a cell, it moves by that cell's offset from A1:

| name | used in | refers to |
|---|---|---|
| `'Data'!B1` | C5 | `'Data'!D5` |
| `'Data'!XFD1` | D3 | `'Data'!C3`, the cell to the left |
| `'Data'!$A1048576` | E5 | `'Data'!A4`, the row above |

The offset wraps round the edges of the sheet, so a name can point to the cell above or to the left. A part with `$` stays where it is. `ROW()` and `COLUMN()` in a name give the row and column of the cell that uses it.

## Scope

A name with a `scope` sheet is seen on that sheet only and wins over a global name of the same name. `'T'!rate` is the name `rate` as sheet T sees it: T's own, else the global one.

Names inside a name's formula are the global ones, as a file stores them. For a sheet's own name the formula says `'T'!rate`. So a global `g = y+0` gives the global `y` on every sheet, and a local `L = y*2` on T uses the global `y` too.

## Errors and cycles

| the name's formula | where it is used |
|---|---|
| a lost reference (`'Data'!#REF!`) | `#REF!` |
| an unknown sheet | `#REF!`, with a name issue |
| an unknown function (`FOO(1)`) | `#NAME?`, with a name issue |
| the name itself, directly or through other names | `#CYCLE!` |

A name that looks like a cell address (`nm2`, `col2`) is not a name. Excel reads such a word as the cell, and so does the engine; the name gets an issue.

A function that takes only a reference (SUMIF's range) and gets an array through a name gives `#VALUE!` for each value of the array, as Excel does.

## Structure operations

Operations rewrite the references in names' formulas as in cells' formulas: a union's parts each on its own, a reference across sheets by its ends. A reference without `$` is an offset, so inserted and deleted rows and columns leave it as it is (`'S'!XFD1` stays when row 1 is deleted).

## How it was checked

The sets `spill/names` (16 books) and `struct/cube` (20 operations) were computed in Excel. The held-out set `spill/names-held` (10 books, 24 cells) was predicted and committed before Excel computed it, and all 24 cells came out as predicted.

## Known differences

- A cycle through names alone is `#CYCLE!`. Excel shows a circular reference warning for it, so the sets do not measure it.
