# Data validation

A sheet's data validation rules (Excel's Data Validation) are part of a workbook: what a cell accepts, the message shown while it is selected, and the alert for a value it does not accept. The engine keeps them, reads and writes them in xlsx files, moves them with structure operations as Excel moves them, returns a cell's rule and the items of its list, and checks a value against a rule as Excel checks it. Calculation does not read them.

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

A sheet's data validation rules (`validations`) define what its cells accept: a list, whole numbers or dates between bounds, a text length, a custom formula, with the message shown while a cell is selected and the alert for a value it does not accept. `session.getValidation(sheet, cell)` returns a cell's rule and the items of its list, calculated in the workbook's current state, each with the `input` Excel writes when it is picked from the dropdown (a cell's value as is, or a literal item parsed as typed in en-US; no rule is checked). `session.validateEntry(sheet, cell, input)` tells whether a value satisfies the rule, as Excel checks it: it returns `{ ok: true }`, or the rule with its alert's style (`severity`) and message, which is the rule's `error` or, if it has none, the text of Excel's alert in the workbook's locale (`This value doesn't match the data validation restrictions defined for this cell.`). `setCellText` checks the value it writes the same way.

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

Rules and outlines change through `setLayout` ([Layout](grid.md#layout)), move with structure operations as Excel moves them, and round-trip through xlsx files. [docs/validation.md](validation.md) describes how a value is checked and how this was verified in Excel.

Rows and columns have outline levels (`outlineLevel`, 1 to 7), which are Excel's groups: a collapsed group consists of hidden rows plus its summary row, the row after it, marked `collapsed`. `outlineGroups(sheet)` (also in `xlcalc/a1`, without the engine) returns a sheet's groups as Excel draws them, each as `{ first, last, level, summary, collapsed }`, from a sheet or its layout. Excel's commands on them are structure operations: `group` and `ungroup` of rows or columns, `hideDetail` and `showDetail` at a row or column (a group's `−` and `+` buttons), and `showLevels` (the level buttons), each one undo step in a session. The markers are read and written as Excel keeps them: a group whose rows are all hidden, including through `setLayout`, is collapsed, and an inner group collapsed inside an outer one stays collapsed when the outer one is expanded.

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

[docs/outline.md](outline.md) has the rules, verified in Excel for Windows.

## Rules

A rule is a `DataValidation`, as an xlsx file stores it:

| field | |
|---|---|
| `range` | the cells, with ranges separated by spaces: `B2:B20 D2`, entire columns `C:C`, entire rows `5:9`. No two rules on a sheet overlap; the ranges of one rule may (`A1:A5 A3:A7`) |
| `type` | `any` (a rule for its input message only), `whole`, `decimal`, `date`, `time`, `textLength`, `list`, `custom` |
| `operator` | for the first five types: `between` (the default), `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`, `greaterThanOrEqual`, `lessThanOrEqual` |
| `formula1`, `formula2` | the value or the bounds, the list, or the custom formula; `formula2` for `between` and `notBetween` |
| `allowBlank` | blank cells pass, Excel's Ignore blank (true by default) |
| `dropdown` | a list shows its dropdown in the cell (true by default) |
| `promptTitle`, `prompt`, `showPrompt` | the input message shown while the cell is selected, and whether it is shown (true by default) |
| `errorStyle`, `errorTitle`, `error`, `showError` | the alert: `stop` (the default), `warning` or `information`, its title and text, and whether it is shown at all (true by default) |

The defaults are those of a rule Excel creates, so a rule created in Excel reads back concisely. Excel accepts titles of up to 32 characters, an input message of 255 and an alert text of 225; the writer truncates longer ones.

Formulas are stored as text without `=` (a leading `=` is dropped), like a cell's formula: English function names, a comma between arguments. A rule's formulas are written for the top-left cell of the rectangle enclosing its ranges, whatever their order, as a conditional format's are, and a relative reference shifts for each other cell as copying would shift it: `{ range: 'B2:B9', type: 'custom', formula1: 'B2>A2' }` checks `B5>A5` for B5, and `{ range: 'H3:H6 J1:J3', type: 'custom', formula1: '$A1>2' }` checks `$A3>2` for H3 and `$A2>2` for J2. That corner need not be a cell of the rule. Excel reads a file's rule this way and saves the rule it makes so, with the ranges in the order they were selected.

A literal list is a text constant of comma-separated items, `'"Open,Shipped,Closed"'`, as a file stores it in every locale. Excel ru-RU displays and accepts the same list as `Open;Shipped;Closed` in its dialog, and a file's `"a;b;c"` is a single item in both locales. A list of cells is a reference to one row or column (`'$E$2:$E$9'`, `"'Lists'!$A$1:$A$20"`), a defined name, a spill reference (`'$G$2#'`) or a formula that returns a reference.

## In a session

`session.getValidation(sheet, cell)` returns the rule that covers a cell, or null: the rule, its formulas adjusted to the cell, and for a list the items its dropdown offers, calculated in the workbook's current state. A literal list's items are its texts; a list of cells gives each cell's value and the text it displays in its number format (`1.50`, `01.01.2025`), without the blank cells after the last value. A list whose formula does not return a reference has its error (`{ error: 'REF' }` for a name lost along with its sheet).

`session.validateEntry(sheet, cell, input)` tells whether the cell containing `input` (a `CellInput`, or null for a cleared cell) satisfies its rule, as Excel's `Validation.Value` checks a cell. It returns `{ ok: true }`, or `{ ok: false, code: 'validation', message, rule, severity }`, where the message is the rule's `error` or, if it has none, the text Excel's alert shows in the workbook's locale (`This value doesn't match the data validation restrictions defined for this cell.`, `Это значение не соответствует ограничениям по проверке данных, установленным для этой ячейки.`), and the severity is its `errorStyle` (stop by default). The input is placed in the workbook for the check, so a custom formula that counts the column sees it, and is then removed: nothing changes. Text a user types first goes through `parseEntry`, or through `setCellText`, which writes the value and returns the rule it breaks as `validation` ([Editing a grid](edits.md#typing-text)). A grid continues from there as Excel does:

- the rule's `errorStyle` decides the alert: `stop` rejects the value, `warning` asks whether to keep it, and `information` notifies and keeps it;
- with `showError: false`, Excel shows no alert and keeps any value;
- Excel does not check a value that is pasted, or that changes because a formula changed; Circle Invalid Data checks it later, which corresponds to `validateEntry` with the cell's own input.

`setLayout([{ sheet, validations: [...] }])` replaces a sheet's rules and recalculates nothing. Rules are part of a session's workbook: `getBook()` includes them, and `editStructure` moves them.

## How a value is checked

These are the rules as verified in Excel for Mac and for Windows ([Checked against Excel](#checked-against-excel)).

### Blanks

A blank cell passes while blanks are ignored, which is the default. When they are not (`allowBlank: false`), a blank is checked as a blank: 0 for a number, a text of length 0, not in any list, and a custom formula calculates with the cell blank. `whole` greater than 0 rejects a blank, while between -1 and 1 accepts it; `textLength` equal to 0 accepts it.

While blanks are ignored, a formula that refers to a blank cell lets any value through: `whole` between `$Z$1` and 10 accepts 50 when Z1 is empty, and so does `$Z$1+1`, and the custom `B2+$Z$1>5` accepts 1. A range does not count: `SUM($Z$1:$Z$3)>5` checks the value. A list through a defined name whose cells include a blank also accepts anything, while the same cells as a reference (`$U$1:$U$5`) do not.

### Numbers, dates and times

`whole`, `decimal`, `date` and `time` accept a number, or a boolean, which compares greater than every number, as Excel's operators compare (TRUE is greater than 5 and not between 1 and 10); text never passes, and neither does `"5"`. `whole` accepts only a whole number. The value is compared with the bounds to 15 significant digits, so `0.1+0.2` equals 0.3. Dates and times are serial numbers and compare as such: a date with a time on the last day is beyond it, so it is not between the two dates, and a time of 1.25 (one day and six hours) is not less than 0.5. Bounds in the wrong order accept nothing: between 10 and 1 rejects 5.

### Text length

`textLength` compares the length of the value as text: `abc` has 3, 12345 has 5, TRUE has 4, and 1.5 has 3.

### Lists

A list of cells accepts a value equal to one of its cells of the same type, with text compared case-insensitively: `APPLE` is in a list containing `apple`, while the number 2 is not in a list containing the text `2`.

A literal list accepts a value whose text equals an item, case-sensitively: `No` is in `"Yes,No"`, and `no` is not. Spaces at either end of the value do not count, and neither do an item's leading spaces: `b` and ` b` are in `"a, b"`, while an item `c ` with a trailing space accepts nothing. A number is compared as its text, as `&` produces it: 1.5 is in `"1.5,2"`, but 1000 is not in `"1e3"`, 2 is not in `"+2"`, 0.5 is not in `"50%"` and a date is not in `"1/2/2025"`.

### Custom formulas

A custom rule accepts a value for which its formula, calculated at the cell with the value in it, returns TRUE or a number other than 0. Text, an error, FALSE or 0 rejects it. An error in the cell fails every type except `any`.

## Under structure operations

A rule's ranges move as references do: they grow when rows or columns are inserted after their first row or column, and move when they are inserted at or before it, and rows or columns inserted immediately after a range get its rule, as they get the format of the row or column before them. Deleted rows and columns are removed from the ranges, and a range that ends on the sheet's last row or column keeps ending there, even for a single cell; a range that loses all its rows or columns is removed, and so is a rule that loses all its ranges. A move takes the rules of its cells along and leaves the destination cells their own rules where the moved cells have none; a copy pastes rules on every tile, and a pasted cell without a rule clears the destination's, as Excel's Paste All does. A rule's formulas change as a cell's formulas do, rewritten for its new first cell when that moves; when moved to another sheet, they are adjusted as a copy adjusts them. [docs/structure.md](structure.md#data-validation-and-outlines) has the details.

## In xlsx files

`readXlsx` reads `<dataValidations>`, and also the rules Excel writes in the sheet's extension (`x14:dataValidations`) when their formulas refer to another sheet; `writeXlsx` writes every rule in `<dataValidations>`, which Excel opens. [docs/xlsx.md](xlsx.md#data-validation-and-outlines) describes what is reported.

### An item picked from the dropdown

Picking an item from a cell's dropdown writes it without checking the rule, and not as typed text: a list of cells gives the cell's value as is, and a literal list's item is parsed as typed in en-US, the locale in which a file stores the list, whatever the workbook's locale. With `001`, `1/2` and `0.5` displayed as `50%` in the list's cells, a General cell gets the text `001`, the text `1/2` and the number 0.5, keeping its own format; from the literal list `"001,1/2,x"` it gets the number 1 and January 2 in the format `d-mmm`, in ru-RU too. Each item of the `list` returned by `getValidation` has this as `input`, which a grid passes to `setCells`:

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

Tested in Excel for Mac and Excel for Windows, en-US and ru-RU, on 74 workbooks, each set predicted before Excel ran it. Not tested: what Excel does with an item picked from a dropdown and the text its dropdown displays, which no object model exposes.

## Known differences

- The writer truncates a rule's texts that exceed Excel's limits (titles of 32 characters, an input message of 255, an alert text of 225), as Excel does not open longer ones.
- A list item containing a comma, typed in ru-RU where the comma is not the list separator, is a single item in Excel while the workbook is open, but the file Excel saves stores the list split at the comma, with the single item only in an extension (`x12ac:list`); the reader reads the split list and reports it.
- Excel saves a list's reference to a deleted sheet as `#REF!` alone; the engine keeps `#REF!$A$1:$A$4`, as Excel displays it while the workbook is open. Both are lost references.
