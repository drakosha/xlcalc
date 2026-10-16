# Data validation

A sheet's rules of data validation (Excel's Data Validation) are part of a book: what a cell takes, the message shown while it is selected, the alert for a value it does not take. The engine keeps them, reads and writes them in xlsx files, moves them with structure operations as Excel moves them, gives a cell's rule and the items of its list, and checks a value against a rule as Excel checks it. Computation does not read them.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [
    {
      name: 'Orders',
      cells: { A1: { value: 'Status' }, B1: { value: 'Qty' }, A2: { value: 'Open' }, E1: { value: 100 } },
      validations: [
        { range: 'A2:A100', type: 'list', formula1: '"Open,Shipped,Closed"', promptTitle: 'Status', prompt: 'Pick from the list' },
        { range: 'B2:B100', type: 'whole', formula1: '1', formula2: '$E$1', errorStyle: 'warning', error: 'From 1 to the limit in E1' },
      ],
    },
  ],
})
console.log(session.getValidation('Orders', 'A5').list.map((i) => i.text).join(' | ')) // → Open | Shipped | Closed
console.log(session.validateEntry('Orders', 'B5', { value: 120 }).ok, session.validateEntry('Orders', 'A5', { value: 'Done' }).ok) // → false false
```

## In short

A sheet's rules of data validation (`validations`) say what its cells take: a list, whole numbers or dates between bounds, a text length, a custom formula, with the message shown while a cell is selected and the alert for a value it does not take. `session.getValidation(sheet, cell)` gives a cell's rule and the items of its list, computed in the book as it is now, each with the `input` Excel writes when it is picked from the dropdown (a cell's value as it is, a literal item read as typed in en-US; no rule checked); `session.validateEntry(sheet, cell, input)` tells whether a value meets the rule, as Excel checks it: `{ ok: true }`, or the rule with its alert's style (`severity`) and message, the rule's `error` or, without one, the text of Excel's alert in the book's locale (`This value doesn't match the data validation restrictions defined for this cell.`); `setCellText` checks the value it writes so.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'Orders',
    cells: { A1: { value: 'Status' }, A2: { value: 'Open' }, B2: { value: 5 }, B3: { value: 7 } },
    validations: [
      { range: 'A2:A100', type: 'list', formula1: '"Open,Shipped,Closed"', prompt: 'Pick a status' },
      { range: 'B2:B3', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning' },
    ],
  }],
})
console.log(session.getValidation('Orders', 'A7').list.map((i) => i.value)) // → [ 'Open', 'Shipped', 'Closed' ]
console.log(session.validateEntry('Orders', 'A7', { value: 'Lost' }).ok, session.validateEntry('Orders', 'B3', { value: -2 }).rule.errorStyle) // → false warning
```

Rules and outlines change through `setLayout` ([Layout](grid.md#layout)), move with structure operations as Excel moves them and go both ways in xlsx files. [docs/validation.md](validation.md) has how a value is checked and how it was measured in Excel.

Rows and columns have outline levels (`outlineLevel`, 1 to 7), Excel's groups: a collapsed group is hidden rows with its summary row, the row past it, marked `collapsed`. `outlineGroups(sheet)` (also from `xlcalc/a1`, without the engine) gives a sheet's groups as Excel draws them, each `{ first, last, level, summary, collapsed }`, from a sheet or its layout; Excel's commands on them are structure operations: `group` and `ungroup` of rows or columns, `hideDetail` and `showDetail` at a line (a group's `−` and `+`), `showLevels` (the level buttons), each one step of the undo in a session. The marks are read and written as Excel keeps them: a group whose rows are all hidden, through `setLayout` too, is collapsed, and an inner group collapsed inside an outer one stays collapsed when the outer one opens.

```js
import { createSession, outlineGroups } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'S', cells: {}, rows: { '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2 }, '8:10': { outlineLevel: 1 } } }],
})
session.editStructure({ op: 'hideDetail', sheet: 'S', row: 8 })
session.editStructure({ op: 'hideDetail', sheet: 'S', row: 11 })
session.editStructure({ op: 'showDetail', sheet: 'S', row: 11 })
console.log(outlineGroups(session.getLayout().sheets[0]).rows.map((g) => [`${g.first}:${g.last}`, g.collapsed])) // → [ [ '3:10', false ], [ '5:7', true ] ]
console.log(session.editStructure({ op: 'showDetail', sheet: 'S', row: 11 }).code) // → detail
```

[docs/outline.md](outline.md) has the rules, measured in Excel for Windows.

## Rules

A rule is `DataValidation`, as an xlsx file holds it:

| field | |
|---|---|
| `range` | the cells, ranges separated by spaces: `B2:B20 D2`, whole columns `C:C`, whole rows `5:9`. No two rules of a sheet meet |
| `type` | `any` (a rule for its input message alone), `whole`, `decimal`, `date`, `time`, `textLength`, `list`, `custom` |
| `operator` | for the first five: `between` (the default), `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`, `greaterThanOrEqual`, `lessThanOrEqual` |
| `formula1`, `formula2` | the value or the bounds, the list, the custom formula; `formula2` for `between` and `notBetween` |
| `allowBlank` | blank cells pass, Excel's Ignore blank (true by default) |
| `dropdown` | a list shows its dropdown in the cell (true by default) |
| `promptTitle`, `prompt`, `showPrompt` | the input message shown while the cell is selected, and whether it is (true by default) |
| `errorStyle`, `errorTitle`, `error`, `showError` | the alert: `stop` (the default), `warning` or `information`, its title and text, and whether it is shown at all (true by default) |

The defaults are those of a rule Excel makes, so a rule made in Excel reads back short. Excel takes titles of 32 characters, an input message of 255 and an alert's text of 225; the writer cuts longer ones.

Formulas are stored text without `=` (a leading one is dropped), as a cell's: English function names, a comma between arguments. A rule's formulas are written for its first cell, the top-left cell of the first range of `range`, and a relative reference moves to each other cell as a copy would move it: `{ range: 'B2:B9', type: 'custom', formula1: 'B2>A2' }` checks `B5>A5` for B5.

A literal list is a text constant of items separated by commas, `'"Open,Shipped,Closed"'`, as a file stores it in every locale. Excel ru-RU shows and takes the same list as `Open;Shipped;Closed` in its window, and a file's `"a;b;c"` is one item in both locales. A list of cells is a reference to one row or column (`'$E$2:$E$9'`, `"'Lists'!$A$1:$A$20"`), a defined name, a spill reference (`'$G$2#'`) or a formula giving a reference.

## In a session

`session.getValidation(sheet, cell)` gives the rule over a cell, or null: the rule, its formulas as at the cell, and for a list the items its dropdown offers, computed in the book as it is now. A literal list's items are its texts; a list of cells gives each cell's value and the text it shows by its number format (`1.50`, `01.01.2025`), without the blank cells past the last value. A list whose formula gives no reference has its error (`{ error: 'REF' }` for a name lost with its sheet).

`session.validateEntry(sheet, cell, input)` tells whether the cell holding `input` (a `CellInput`, or null for a cleared cell) meets its rule, as Excel's `Validation.Value` checks a cell: `{ ok: true }`, or `{ ok: false, code: 'validation', message, rule, severity }`, the message the rule's `error` or, without one, the text Excel's alert shows in the book's locale (`This value doesn't match the data validation restrictions defined for this cell.`, `Это значение не соответствует ограничениям по проверке данных, установленным для этой ячейки.`) and the severity its `errorStyle` (stop by default). The input goes into the book for the check, so a custom formula that counts the column sees it, and comes out again: nothing changes. Text a user types is `parseEntry`'s input first, or goes through `setCellText`, which writes the value and gives the rule it breaks as `validation` ([Editing a grid](edits.md#typing-text)). A grid takes it from there as Excel does:

- the rule's `errorStyle` decides the alert: `stop` refuses the value, `warning` asks whether to keep it, `information` tells and keeps it;
- with `showError: false` Excel shows no alert and keeps any value;
- a value pasted, or put in by a formula's change, is not checked by Excel; Circle Invalid Data checks it later, which is `validateEntry` with the cell's own input.

`setLayout([{ sheet, validations: [...] }])` replaces a sheet's rules and recomputes nothing. Rules go with a session's book: `getBook()` has them, and `editStructure` moves them.

## How a value is checked

These are the rules measured in Excel for Mac and for Windows ([Checked against Excel](#checked-against-excel)).

### Blanks

A blank cell passes while blanks are ignored, the default. Where they are not (`allowBlank: false`), a blank is checked as a blank: 0 for a number, a text of length 0, in no list, and a custom formula computes with the cell blank. `whole` greater than 0 refuses a blank, between -1 and 1 takes it; `textLength` equal to 0 takes it.

While blanks are ignored, a formula that refers to a blank cell lets any value in: `whole` between `$Z$1` and 10 takes 50 when Z1 is empty, and so does `$Z$1+1`, and the custom `B2+$Z$1>5` takes 1. A range does not count: `SUM($Z$1:$Z$3)>5` checks the value. A list through a defined name whose cells hold a blank takes anything too, while the same cells as a reference (`$U$1:$U$5`) do not.

### Numbers, dates and times

`whole`, `decimal`, `date` and `time` take a number, or a boolean, which compares above every number as Excel's operators compare (TRUE is greater than 5 and not between 1 and 10); text never passes, `"5"` neither. `whole` takes a whole number alone. The value compares with the bounds to 15 significant digits, so `0.1+0.2` equals 0.3. Dates and times are serial numbers and compare as they are: a date with a time on the last day is past it, so not between the two dates, and a time of 1.25 (a day and six hours) is not less than 0.5. Bounds in the wrong order take nothing: between 10 and 1 refuses 5.

### Text length

`textLength` compares the length of the value as text: `abc` has 3, 12345 has 5, TRUE has 4, 1.5 has 3.

### Lists

A list of cells takes a value equal to one of its cells of the same type, text in any case: `APPLE` is in a list holding `apple`, the number 2 is not in a list holding the text `2`.

A literal list takes a value whose text equals an item, in its case: `No` is in `"Yes,No"`, `no` is not. A value's spaces at either end do not count, and neither do an item's leading spaces: `b` and ` b` are in `"a, b"`, while an item `c ` with a trailing space takes nothing. A number is its text as `&` writes it: 1.5 is in `"1.5,2"`, but 1000 is not in `"1e3"`, 2 not in `"+2"`, 0.5 not in `"50%"` and a date not in `"1/2/2025"`.

### Custom formulas

A custom rule takes a value for which its formula, computed at the cell with the value in it, is TRUE or a number other than 0. A text, an error, FALSE or 0 refuses it. An error in the cell fails every type but `any`.

## Under structure operations

A rule's ranges move as references do: they grow with lines inserted after their first line and move with lines inserted at or before it, and lines inserted right after a range take its rule, as they take the format of the line before them. Deleted lines leave the ranges, and a range ending on the sheet's last row or column keeps ending there, one cell too; a range that loses every line is gone, and a rule that loses every range. A move takes the rules of its cells along and leaves the cells it lands on their own rules where the moved cells have none; a copy pastes rules on every tile, and a pasted cell without a rule clears the target's, as Excel's Paste All does. A rule's formulas change as a cell's formulas do, written for its new first cell when that moves; moved to another sheet, they move as a copy moves them. [docs/structure.md](structure.md#data-validation-and-outlines) has the details.

## In xlsx files

`readXlsx` reads `<dataValidations>`, and the rules Excel writes in the sheet's extension (`x14:dataValidations`) when their formulas refer to another sheet; `writeXlsx` writes every rule in `<dataValidations>`, which Excel opens. [docs/xlsx.md](xlsx.md#data-validation-and-outlines) has what is reported.

### An item picked from the dropdown

Picking an item of a cell's dropdown writes it without checking the rule, and not as typed text: a list of cells gives the cell's value as it is, and a literal list's item is read as typed in en-US, the locale a file holds the list in, whatever the book's locale. With `001`, `1/2` and `0.5` shown as `50%` in the list's cells, a cell in General takes the text `001`, the text `1/2` and the number 0.5, keeping its own format; from the literal list `"001,1/2,x"` it takes the number 1 and the 2nd of January in the format `d-mmm`, in ru-RU too. Each item of `getValidation`'s `list` has that as `input`, which a grid passes to `setCells`:

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'S', cells: { B2: { value: 0 } }, validations: [{ range: 'B2', type: 'list', formula1: '"001,1/2,x"' }] }],
}, { clock: { now: '2026-10-16T12:00:00Z', seed: 1 } }) // a date without a year takes the clock's
const { list } = session.getValidation('S', 'B2')
console.log(list.map((item) => item.input))
// → [ { value: 1 }, { value: 46024, numberFormat: 16 }, { value: 'x' } ]
session.setCells([{ sheet: 'S', cell: 'B2', input: list[1].input }])
console.log(session.getCell('S', 'B2'))
// → {
//     value: 46024,
//     text: '2-Jan',
//     numberFormat: 16,
//     numberFormatCode: 'd-mmm'
//   }
```

## Checked against Excel

Measured in Excel for Mac and Excel for Windows, en-US and ru-RU, on 74 books, each set predicted before Excel ran it. Not measured: what Excel does with an item picked in a dropdown and the text its dropdown shows, which no object model gives.

## Known differences

- A rule's texts over Excel's limits (titles of 32 characters, an input message of 255, an alert's text of 225) are cut by the writer, as Excel does not open longer ones.
- A list item holding a comma, typed in ru-RU where the comma is no list separator, is one item in Excel while the book is open, and the file Excel saves holds the list split at the comma, with the one item only in an extension (`x12ac:list`); the reader reads the list split and reports it.
- Excel saves a list's reference to a deleted sheet as `#REF!` alone; the engine keeps `#REF!$A$1:$A$4`, as Excel shows it while the book is open. Both are lost references.
