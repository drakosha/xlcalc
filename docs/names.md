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

## A name

A name follows Excel's rules for names, which a table's name follows too: a letter of any script, `_` or `\` first, then letters, digits, `_`, `.`, `\` and `?`, at most 255 characters. It must not read as something else in a formula: a cell address (`Q1`, `AB12`), an R1C1 reference (`R2C3`, and `R` or `C` alone), TRUE or FALSE. Cell addresses and R1C1 are in Latin letters, so the same shape in Cyrillic is a name. In a ru-RU book its booleans `ИСТИНА` and `ЛОЖЬ` are no names either, in any case: Excel ru-RU refuses them and does not open a file that holds such a name, while Excel en-US takes them.

| name | |
|---|---|
| `Ставка`, `НДС_2025`, `НДС.2025`, `_Ставка`, `\Ставка`, `Став?ка`, `Став\ка`, `Τιμή` | a name |
| `АБ12`, `К2С3`, `С` in Cyrillic letters | a name |
| `Сумма`, `СУММ`, `SUM` | a name: a function is a word followed by `(` |
| `ИСТИНА`, `ложь` | a name in an en-US book, not in a ru-RU one |
| `1Ставка`, `?Ставка`, `Став ка`, `AB12`, `R2C3`, `R`, `TRUE` | not a name |

```js
import { compute, validateName } from 'xlcalc'

console.log(validateName('Ставка').ok, validateName('1Ставка').ok) // → true false
const r = compute({ sheets: [{ name: 'S', cells: { A1: { formula: 'СТАВКА*2' } } }], names: [{ name: 'Ставка', ref: '0.2' }] })
console.log(r.sheets[0].cells.A1.value, r.sheets[0].cells.A1.formula) // → 0.4 Ставка*2
```

Names are compared in any case, Cyrillic too: `ставка` and `СТАВКА` in a formula are the name `Ставка`, and the formula's text shows it as it is defined (`Ставка*2`), as Excel shows it. A second name `СТАВКА` beside it is a duplicate, left out with a name issue. `Ё` and `Е` are two letters, so `Ёлка` and `Елка` are two names.

A name is written as it is in every locale. In ru-RU a Russian function name followed by `(` is the function and the same word alone is a name: `=СУММ(Сумма;1)` sums the name `Сумма`. A stored formula can hold the word `ИСТИНА` as a name only from an en-US book; `toLocal` shows it in ru-RU in quotes, `'ИСТИНА'`, as Excel ru-RU does, and `toInvariant` reads `'ИСТИНА'` back as the name, while `ИСТИНА` typed in Russian is TRUE. A table named `ИСТИНА` keeps no quotes before its brackets (`ИСТИНА[Qty]`).

These are the rules of Microsoft's pages on defined names and on table names, with what Excel was measured to do beyond them: `\` and `?` inside a name, and the booleans of ru-RU. `validateName` checks the rules of an en-US book.

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

A name's formula is written as a file stores it, without `=`, and `readXlsx` reads it so. Every reference in it has its sheet. A name whose formula does not parse, or has a reference without a sheet, is left out with a name issue.

## References without `$`

A reference without `$` is relative to A1. Used in a cell, it moves by that cell's offset from A1:

| name | used in | refers to |
|---|---|---|
| `'Data'!B1` | C5 | `'Data'!D5` |
| `'Data'!XFD1` | D3 | `'Data'!C3`, the cell to the left |
| `'Data'!$A1048576` | E5 | `'Data'!A4`, the row above |

The offset wraps round the edges of the sheet, so a name can point to the cell above or to the left. A part with `$` stays where it is. `ROW()` and `COLUMN()` in a name give the row and column of the cell that uses it.

A name cannot be a table's: tables and names share one namespace ([Tables](tables.md)). A name's formula can hold a structured reference (`Sales[Qty]`); `Sales[[#This Row],[Price]]` takes the row of the cell that uses the name.

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

The sets `spill/names` (16 books) and `struct/cube` (20 operations) were computed in Excel. The 24 cells of the set `spill/names-held` (10 books) were predicted before Excel computed them, and all came out as predicted.

The rules of a name were measured in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US and ru-RU, through the table sets ([Tables](tables.md)): names in formulas of the file and typed, names added through Excel's object model (`Names.Add`), which refuses a name Excel does not take, tables renamed and books Excel loads or not. Six sets were predicted before Excel computed them:

| set | cases | as predicted (Mac) | misses |
|---|---:|---:|---|
| `names-held` | 17 | 13 | a name typed in another case shown as defined (2, fixed); tables renamed to names Excel does not take (2, known) |
| `names-ru-held` (ru-RU) | 8 | 6 | the name's case (fixed); a file with the name `ИСТИНА`, which Excel ru-RU does not open (fixed) |
| `names-more-held` | 12 | 7 | `?` and `\` inside a name, added and in a file (4, fixed); a 255-letter name in a typed formula, cut by AppleScript (known, the Mac only) |
| `names-ru-more-held` (ru-RU) | 10 | 5 | `ИСТИНА` and `ЛОЖЬ` refused and TRUE and FALSE taken (TRUE known), `?` and `\` (fixed), files with `ИСТИНА`, `истина`, `ЛОЖЬ` not opened (3, fixed); one more the predictions held no text for: the word `ИСТИНА` shown as `'ИСТИНА'` (fixed) |
| `names-tables-held` | 4 | 3 | a table renamed `?Продажи` (known) |
| `names-ru-tables-held` (ru-RU) | 3 | 0 | tables renamed `ИСТИНА`, `ложь`, `TRUE`, where the two Excels part (known) |

In all, 34 of the 54 cases came out as predicted on the Mac and 35 on Windows, which agreed with the Mac on every case but the 255-letter name and the tables of `names-ru-tables-held`.

## Known differences

- A cycle through names alone is `#CYCLE!`. Excel shows a circular reference warning for it, so the sets do not measure it.
- Excel ru-RU takes TRUE and FALSE as defined names through its object model. Formulas are kept in en-US text, where they are the booleans, so a name TRUE could not be written in one; the engine refuses them in every locale.
- A table renamed to a name Excel does not take: Excel's object model makes one (`_1Ставка`, `_?Продажи`, `СТАВКА_1` beside a name `Ставка`), where `renameTable` refuses. In ru-RU Excel for Mac keeps a table named `ИСТИНА` and makes `TRUE_` of `TRUE`, Excel for Windows makes `ИСТИНА_` and keeps `TRUE`; the engine refuses both, following Windows for `ИСТИНА`.
