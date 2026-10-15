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

## In short

```js
import { compute } from 'xlcalc'

const result = compute({
  sheets: [
    { name: 'Data', cells: { A1: { value: 'tea' }, B1: { value: 3 }, A2: { value: 'coffee' }, B2: { value: 5 } } },
    { name: 'Report', cells: { A1: { formula: 'SUMIF(items,"tea",amounts)' } } },
  ],
  names: [
    { name: 'items', ref: "'Data'!$A:$A" },
    { name: 'amounts', ref: "'Data'!$B:$B" },
  ],
})
console.log(result.sheets[1].cells.A1.value) // → 3
```

A name follows Excel's rules for names, in letters of any script (`Ставка`, `НДС_2025`, `Став?ка`), and is compared in any case: `ставка` is the name `Ставка`, and a formula's text shows it as it is defined (`ставка*2` is `Ставка*2`). A name with `scope: 'Sheet'` is visible on that sheet only and shadows a global name. Invalid names are reported in the result's `diagnostics` (`kind: 'name'`), `dropped` telling a name left out from one kept with an error.

A name can hold any formula: a constant (`0.2`), a calculation (`'Data'!$B$1*2`), a reference from `INDEX`, a union. A reference without `$` moves with the cell that uses the name. [docs/names.md](names.md) has the rules.

## A name

A name follows Excel's rules for names, which a table's name follows too: a letter of any script, `_` or `\` first, then letters, digits, `_`, `.`, `\` and `?`, at most 255 characters. It must not read as something else in a formula: a cell address of the sheet (`Q1`, `AB12`, `XFD1`), an R1C1 reference (`R2C3`, and `R` or `C` alone), TRUE or FALSE. A cell's shape past the sheet is a name, `XFE1` and `A0`, and a formula reads it as one. Cell addresses and R1C1 are in Latin letters, so the same shape in Cyrillic is a name. In a ru-RU book its booleans `ИСТИНА` and `ЛОЖЬ` are no names either, in any case: Excel ru-RU refuses them and does not open a file that holds such a name, while Excel en-US takes them.

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

Names are compared in any case, Cyrillic too: `ставка` and `СТАВКА` in a formula are the name `Ставка`, and the formula's text shows it as it is defined (`Ставка*2`), as Excel shows it. A second name `СТАВКА` beside it is a duplicate, left out with a name issue. `Ё` and `Е` are two letters, so `Ёлка` and `Елка` are two names. A letter compares by its whole upper case, so `ß` is `SS` and `ﬁ` is `FI` (`straße` is the name `STRASSE`), but the dotless `ı` and the long `ſ` are no `I` and `S`, as Excel for Windows compares names; a LET name compares letter by letter, so `LET(ß,1,SS)` is `#NAME?`.

A name is written as it is in every locale. In ru-RU a Russian function name followed by `(` is the function and the same word alone is a name: `=СУММ(Сумма;1)` sums the name `Сумма`. A stored formula can hold the word `ИСТИНА` as a name only from an en-US book; `toLocal` shows it in ru-RU in quotes, `'ИСТИНА'`, as Excel ru-RU does, and `toInvariant` reads `'ИСТИНА'` back as the name, while `ИСТИНА` typed in Russian is TRUE. A table named `ИСТИНА` keeps no quotes before its brackets (`ИСТИНА[Qty]`).

These are the rules of Microsoft's pages on defined names and on table names, with what Excel was measured to do beyond them: `\` and `?` inside a name, and the booleans of ru-RU. `validateName(name, { locale })` checks them for a book of that locale, en-US without one, and refuses a name with the code of the rule it breaks: `chars` (its first character, a character inside or its length), `reference` (it reads as a cell or R1C1: `AB12`, `R2C3`, `R`) or `boolean` (`TRUE`, `FALSE`, and `ИСТИНА` and `ЛОЖЬ` in ru-RU).

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

A name's formula is written as a file stores it, without `=` (a leading one is dropped, as Excel's "Refers to" shows one), and `readXlsx` reads it so. Every reference in it has its sheet. A name whose formula does not parse, or has a reference without a sheet, is left out with a name issue.

## Names left out and names with errors

A name the book cannot have is left out: one Excel would not take, a second one of the same name and scope, a scope sheet the book lacks, a formula that does not parse or a reference without its sheet. A formula that uses its word finds no name and gives `#NAME?`. A name whose formula refers to a sheet, a table or a function the book lacks stays a name and gives its error where it is used ([Errors and cycles](#errors-and-cycles)).

A result's `diagnostics` list both, of kind `name`, and `dropped` tells them apart. `validateNames(names, { sheets, tables, locale })` gives the same list without the book, by the same check, so a host can refuse a name before it stores it: whether a name is left out depends only on the names, the sheets' names and the locale.

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

A LET name has rules of its own: any cell's word is one, written without `$` (`LET(A1,1,A1)` is 1, the calculation's `A1` the name), but one that reads as R1C1's row or column too (`R2`, `C5`, `RC5`) and TRUE and FALSE; R1C1's other shapes (`R`, `RC`, `R1C1`) are names here too.

A function that takes only a reference (SUMIF's range) and gets an array through a name gives `#VALUE!` for each value of the array, as Excel does.

## Structure operations

Operations rewrite the references in names' formulas as in cells' formulas: a union's parts each on its own, a reference across sheets by its ends. A reference without `$` is an offset, so inserted and deleted rows and columns leave it as it is (`'S'!XFD1` stays when row 1 is deleted).

## Checked against Excel

Measured in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US and ru-RU, through the spill, structure, table and entry sets, the later ones predicted before Excel ran them. Where the two Excels part (the long `ſ` in a name), the engine follows Excel for Windows.

## Known differences

- A cycle through names alone is `#CYCLE!`. Excel shows a circular reference warning for it, so the sets do not measure it.
- Excel ru-RU takes TRUE and FALSE as defined names through its object model. Formulas are kept in en-US text, where they are the booleans, so a name TRUE could not be written in one; the engine refuses them in every locale.
- A table renamed to a name Excel does not take: Excel's object model makes one (`_1Ставка`, `_?Продажи`, `СТАВКА_1` beside a name `Ставка`), where `renameTable` refuses. In ru-RU Excel for Mac keeps a table named `ИСТИНА` and makes `TRUE_` of `TRUE`, Excel for Windows makes `ИСТИНА_` and keeps `TRUE`; the engine refuses both, following Windows for `ИСТИНА`.
