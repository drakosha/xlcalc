# Defined names

A defined name stands for a formula, as Excel's "Refers to" does. The formula can be a reference, a constant, a calculation or a union of references. A formula that uses the name gets the result of the name's formula at the cell where it is used.

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

A name follows Excel's naming rules, in letters of any script (`Ставка`, `НДС_2025`, `Став?ка`), and is case-insensitive: `ставка` is the name `Ставка`, and a formula's text shows it as it is defined (`ставка*2` becomes `Ставка*2`). A name with `scope: 'Sheet'` is visible on that sheet only and shadows a global name. Invalid names are reported in the result's `diagnostics` (`kind: 'name'`), where `dropped` distinguishes a name that was left out from one that was kept with an error.

A name can contain any formula: a constant (`0.2`), a calculation (`'Data'!$B$1*2`), a reference from `INDEX`, or a union. A reference without `$` moves with the cell that uses the name. [docs/names.md](names.md) has the rules.

## A name

A name follows Excel's naming rules, which table names follow too: it starts with a letter of any script, `_` or `\`, continues with letters, digits, `_`, `.`, `\` and `?`, and is at most 255 characters long. It must not be parsed as something else in a formula: a cell address on the sheet (`Q1`, `AB12`, `XFD1`), an R1C1 reference (`R2C3`, and `R` or `C` alone), TRUE or FALSE. Something shaped like a cell but beyond the sheet is a name, such as `XFE1` and `A0`, and a formula parses it as one. Cell addresses and R1C1 references use Latin letters, so the same shape in Cyrillic is a name. In a ru-RU workbook, its booleans `ИСТИНА` and `ЛОЖЬ` are not valid names either, in any case: Excel ru-RU rejects them and does not open a file that contains such a name, while Excel en-US accepts them.

| name | |
|---|---|
| `Ставка`, `НДС_2025`, `НДС.2025`, `_Ставка`, `\Ставка`, `Став?ка`, `Став\ка`, `Τιμή` | a name |
| `АБ12`, `К2С3`, `С` in Cyrillic letters | a name |
| `Сумма`, `СУММ`, `SUM` | a name: a function is a word followed by `(` |
| `ИСТИНА`, `ложь` | a name in an en-US workbook, not in a ru-RU one |
| `1Ставка`, `?Ставка`, `Став ка`, `AB12`, `R2C3`, `R`, `TRUE` | not a name |

```js
import { compute, validateName } from 'xlcalc'

console.log(validateName('Ставка').ok, validateName('1Ставка').ok) // → true false
const r = compute({ sheets: [{ name: 'S', cells: { A1: { formula: 'СТАВКА*2' } } }], names: [{ name: 'Ставка', ref: '0.2' }] })
console.log(r.sheets[0].cells.A1.value, r.sheets[0].cells.A1.formula) // → 0.4 Ставка*2
```

Names are case-insensitive, Cyrillic included: `ставка` and `СТАВКА` in a formula are the name `Ставка`, and the formula's text shows it as it is defined (`Ставка*2`), as Excel shows it. A second name `СТАВКА` alongside it is a duplicate and is left out with a name issue. `Ё` and `Е` are two different letters, so `Ёлка` and `Елка` are two names. Letters are compared by their full uppercase form, so `ß` is `SS` and `ﬁ` is `FI` (`straße` is the name `STRASSE`), but the dotless `ı` and the long `ſ` are not `I` and `S`, as in Excel for Windows; a LET name is compared letter by letter, so `LET(ß,1,SS)` is `#NAME?`.

A name is written the same way in every locale. In ru-RU, a Russian function name followed by `(` is the function, and the same word alone is a name: `=СУММ(Сумма;1)` sums the name `Сумма`. A stored formula can contain the word `ИСТИНА` as a name only if it comes from an en-US workbook; `toLocal` displays it in ru-RU in quotes, `'ИСТИНА'`, as Excel ru-RU does, and `toInvariant` parses `'ИСТИНА'` back as the name, while `ИСТИНА` typed in Russian is TRUE. A table named `ИСТИНА` gets no quotes before its brackets (`ИСТИНА[Qty]`).

These are the rules from Microsoft's documentation on defined names and table names, plus what Excel was found to do beyond them: `\` and `?` inside a name, and the booleans of ru-RU. `validateName(name, { locale })` checks them for a workbook in that locale (en-US if omitted) and rejects a name with the code of the rule it breaks: `chars` (its first character, a character inside it, or its length), `reference` (it would be parsed as a cell or an R1C1 reference: `AB12`, `R2C3`, `R`) or `boolean` (`TRUE`, `FALSE`, and `ИСТИНА` and `ЛОЖЬ` in ru-RU).

The rules below were verified in Excel 16.113 for Mac (en-US).

## What a name can contain

| `ref` | used as | returns |
|---|---|---|
| `0.2`, `"text"`, `TRUE`, `#N/A` | `=rate*100` | the constant |
| `{1,2,3}` | `=arr` | the array, which spills; `SUM(arr)` is 6, `COLUMNS(arr)` is 3 |
| `'Data'!$A$2*2` | `=dbl` | the value of the calculation |
| `'Data'!$A$1:$A$3*10` | `=SUM(sq)`, `=sq` | an array: summed, or spilled |
| `INDEX('Data'!$A$1:$C$9,0,2)` | `=SUMIF(col,">1")` | a reference, which functions that accept ranges accept |
| `'Data'!$A$1:$A$3,'Data'!$C$1:$C$3` | `=SUM(both)` | a union: see [References](references.md) |
| `'Data'!$B$1#` | `=SUM(sp)` | the spill range of B1 |

A name's formula is written as a file stores it, without `=` (a leading `=` is dropped, although Excel's "Refers to" shows one), and `readXlsx` reads it that way. Every reference in it includes its sheet. A name whose formula does not parse, or contains a reference without a sheet, is left out with a name issue.

## Names left out and names with errors

A name the workbook cannot have is left out: one Excel would not accept, a second one with the same name and scope, a scope sheet the workbook does not have, a formula that does not parse, or a reference without its sheet. A formula that uses its word finds no name and evaluates to `#NAME?`. A name whose formula refers to a sheet, a table or a function the workbook does not have stays a name and returns its error where it is used ([Errors and cycles](#errors-and-cycles)).

A result's `diagnostics` list both, with kind `name`, and `dropped` tells them apart. `validateNames(names, { sheets, tables, locale })` returns the same list without the workbook, using the same check, so a host can reject a name before storing it: whether a name is left out depends only on the names, the sheet names and the locale.

A session's `setNames` (and the operation `setNames` of `applyChanges`) refuses a name the workbook would leave out with an `InputError` at that name (`names[1].name`), whose message gives `validateNames`' reason, and changes nothing. A name the workbook already has as it stands is taken back: `setNames` replaces the whole list, so a host sends back the names the workbook was opened with, hidden ones among them, unchanged. `compute` and `createSession` still take such a name and report it in `diagnostics`.

## References without `$`

A reference without `$` is relative to A1. When used in a cell, it moves by that cell's offset from A1:

| name | used in | refers to |
|---|---|---|
| `'Data'!B1` | C5 | `'Data'!D5` |
| `'Data'!XFD1` | D3 | `'Data'!C3`, the cell to the left |
| `'Data'!$A1048576` | E5 | `'Data'!A4`, the row above |

The offset wraps around the edges of the sheet, so a name can point to the cell above or to the left. A part with `$` stays fixed. `ROW()` and `COLUMN()` in a name return the row and column of the cell that uses it.

A name cannot have the same name as a table: tables and names share one namespace ([Tables](tables.md)). A name's formula can contain a structured reference (`Sales[Qty]`); `Sales[[#This Row],[Price]]` takes the row of the cell that uses the name.

## Scope

A name with a `scope` sheet is visible on that sheet only and takes precedence over a global name with the same name. `'T'!rate` is the name `rate` as sheet T sees it: T's own, otherwise the global one.

Names inside a name's formula are the global ones, as a file stores them. For a sheet's own name, the formula says `'T'!rate`. So a global `g = y+0` uses the global `y` on every sheet, and a local `L = y*2` on T uses the global `y` too.

## Hidden names

A name with `hidden: true` stays out of Excel's Name Manager, as a file marks it with `hidden="1"`; a host's list of names should leave it out too. It calculates like any other name: the flag changes no value and no values hash, only `inputHash`. Structure operations, sessions with their undo, and `readXlsx` and `writeXlsx` keep it, and a local name that a moved or copied formula takes to another sheet is hidden there too. `hidden: false` is a visible name, as one without the field.

```js
names: [{ name: 'rate', ref: "'Data'!$B$1", hidden: true }]
```

## Errors and cycles

| the name's formula | where it is used |
|---|---|
| a lost reference (`'Data'!#REF!`) | `#REF!` |
| an unknown sheet | `#REF!`, with a name issue |
| an unknown function (`FOO(1)`) | `#NAME?`, with a name issue |
| the name itself, directly or through other names | `#CYCLE!` |

A name that looks like a cell address (`nm2`, `col2`) is not a name. Excel parses such a word as the cell, and so does the engine; the name gets an issue.

A LET name has its own rules: any cell address can be one, written without `$` (`LET(A1,1,A1)` is 1, and the calculation's `A1` is the name), except one that is also an R1C1 row or column reference (`R2`, `C5`, `RC5`), and TRUE and FALSE; the other R1C1 shapes (`R`, `RC`, `R1C1`) are valid LET names too.

A function that accepts only a reference (SUMIF's range) and gets an array through a name returns `#VALUE!` for each value of the array, as Excel does.

## Structure operations

Operations rewrite the references in names' formulas as in cells' formulas: each part of a union separately, and a reference across sheets by its ends. A reference without `$` is an offset, so inserted and deleted rows and columns leave it unchanged (`'S'!XFD1` stays when row 1 is deleted), and so is the part without `$` of a mixed reference (`'S'!$A1` stays when row 1 is inserted, and becomes `'S'!$B1` when a column is inserted before A). The part with `$` moves alone, as one line: deleting row 2 makes `'S'!$B$2:$B1` `'S'!#REF!`, as Excel does, though the same range in a cell's formula would shrink. Renaming a sheet rewrites its local names written with it (`'S'!rate`).

## Checked against Excel

Tested in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US and ru-RU, through the spill, structure, table and entry sets, the later ones predicted before Excel ran them. Where the two versions of Excel differ (the long `ſ` in a name), the engine follows Excel for Windows.

## Known differences

- A cycle through names alone is `#CYCLE!`. Excel shows a circular reference warning for it, so the test sets do not cover it.
- Excel ru-RU accepts TRUE and FALSE as defined names through its object model. Formulas are stored in en-US syntax, where these are the booleans, so a name TRUE could not be written in one; the engine rejects them in every locale.
- A table renamed to a name Excel does not accept: Excel's object model generates one (`_1Ставка`, `_?Продажи`, `СТАВКА_1` next to a name `Ставка`), whereas `renameTable` rejects it. In ru-RU, Excel for Mac keeps a table named `ИСТИНА` and turns `TRUE` into `TRUE_`, while Excel for Windows turns `ИСТИНА` into `ИСТИНА_` and keeps `TRUE`; the engine rejects both, following Windows for `ИСТИНА`.
