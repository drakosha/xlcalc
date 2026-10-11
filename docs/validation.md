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

`session.validateEntry(sheet, cell, input)` tells whether the cell holding `input` (a `CellInput`, or null for a cleared cell) meets its rule, as Excel's `Validation.Value` checks a cell: `{ ok: true }`, or `{ ok: false, code: 'validation', message, rule, severity }`, the message the rule's `error` or the library's text in the book's locale and the severity its `errorStyle` (stop by default). The input goes into the book for the check, so a custom formula that counts the column sees it, and comes out again: nothing changes. Text a user types is `parseEntry`'s input first, or goes through `setCellText`, which writes the value and gives the rule it breaks as `validation` ([Editing a grid](edits.md#typing-text)). A grid takes it from there as Excel does:

- the rule's `errorStyle` decides the alert: `stop` refuses the value, `warning` asks whether to keep it, `information` tells and keeps it;
- with `showError: false` Excel shows no alert and keeps any value;
- a value pasted, or put in by a formula's change, is not checked by Excel; Circle Invalid Data checks it later, which is `validateEntry` with the cell's own input.

`setLayout([{ sheet, validations: [...] }])` replaces a sheet's rules and recomputes nothing. Rules go with a session's book: `getBook()` has them, and `editStructure` moves them.

## How a value is checked

These are the rules measured in Excel for Mac and for Windows (`How it was checked`).

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

## How it was checked

`npx tsx bench/coverage/run.ts valid <set> [ru]` writes each book of `bench/coverage/valid/<set>.ts` with `writeXlsx`. Excel for Mac opens it, makes what the case makes through its object model (`add data validation` with the range's first cell selected, groups, a collapsed group), applies the case's structure operation, and over a window of each sheet reads, with each cell selected in turn, the cell's rule (type, operator, formulas, settings, texts) and `Validation.Value`, whether the cell's value meets it; it reads the outline too and saves the book. `npx tsx bench/windows/run.ts valid <set> [ru]` does the same in Excel for Windows through COM. `test/validation-excel.test.ts` and `test/validation-win.test.ts` check, without Excel, that Excel read what the engine holds after each case, `validateEntry` with each cell's own input against `Validation.Value`, and that the reader reads the file Excel saved as Excel read it.

| set | cases | what |
|---|---|---|
| `files` | 8 | books written with every type and setting, whole columns, several ranges, another sheet, texts past Excel's limits, outlines; books written after structure operations in the engine |
| `made` | 4 | rules of every type, groups and collapsed groups made by Excel itself |
| `struct-held` | 31 | lines inserted and deleted around and inside rules and groups, moves and copies of rules and onto them, sheets renamed and deleted under lists |
| `struct-held-2` | 15 | held out after `struct-held`: lines inserted right after rules, moves onto rules and to another sheet, lines inserted into collapsed groups |
| `entry-held` | 96 values | values of every kind against every type, bounds on blank cells, lists of cells holding numbers as text |
| `lists` | 30 values | lists through names, another sheet, a spill, a formula, a whole column, literal lists with `;` and spaces |
| `entry-held-2` | 39 values | held out after `entry-held` and `lists` |
| `held-3` | 30 values, 2 moves | held out after the second sets: blanks where blanks are not ignored, literal numbers, moves and copies to another sheet |
| `ru` | 2 | in ru-RU: literal lists with commas and semicolons in a file, and lists, bounds and formulas typed into the object model in Russian |
| `edge-held` | 8 | rules at the sheet's last row and column, one cell and two, under lines deleted and inserted far before them and the last line deleted |
| `sqref-held` | 5 | a rule's cells as another writer may spell them: in lower case, with `$`, both, several ranges, Excel 2010's `<xm:sqref>` (put into the file by hand) |

The engine's answers were committed before Excel ran each set (`bench/coverage/valid/*.predicted.json`). Excel loaded every written book. Of the predictions, `files` held in 2 of 8 books (a literal list matches text in its case; texts the writer cuts), `made` in 3 of 4 (lines inserted into a collapsed group open it), `struct-held` in 26 of 31 (lines inserted right after a rule take it; a move keeps the target's rules; a formula referring to a blank cell lets anything in), `struct-held-2` in 14 of 15 (a rule moved to another sheet moves its formulas as a copy), the checks of values in 87 of 96 (`entry-held`), 26 of 30 (`lists`), 36 of 39 (`entry-held-2`) and 27 of 30 (`held-3`), and `edge-held` in 4 of 8 (a rule of one cell on the last row or column stretches over the line before it when lines above are deleted, as a range ending there does; Excel refuses lines inserted that push its cells off the sheet, as predicted). Each miss gave one of the rules above, fitted on the set that showed it. In `sqref-held` both Excels took every spelling as the range it names, where the engine predicted four of the five rules lost (it read lower case but not `$`); it now reads them as Excel does. Excel for Windows read what Excel for Mac read in every set in en-US.

In ru-RU both Excels read a file's literal lists as en-US does, split at commas. Typed through the object model in Russian, a list `x;y;z` is three items in Excel for Windows and one in Excel for Mac's AppleScript, which also made no rule of a formula with a Russian function name, and Excel for Mac's object model shows a rule's formulas now in Russian, now in English, so the ru-RU set compares Excel for Mac's values and settings and leaves the formulas to the file it saved.

What Excel does with a picked item of a dropdown, and the text its dropdown shows, was not measured: no object model gives them. Lists of cells were checked by which values they take.

## Known differences

- A rule's texts over Excel's limits (titles of 32 characters, an input message of 255, an alert's text of 225) are cut by the writer, as Excel does not open longer ones.
- A list item holding a comma, typed in ru-RU where the comma is no list separator, is one item in Excel while the book is open, and the file Excel saves holds the list split at the comma, with the one item only in an extension (`x12ac:list`); the reader reads the list split and reports it.
- Excel saves a list's reference to a deleted sheet as `#REF!` alone; the engine keeps `#REF!$A$1:$A$4`, as Excel shows it while the book is open. Both are lost references.
