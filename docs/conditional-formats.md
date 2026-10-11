# Conditional formats

A sheet can hold Excel's conditional formatting rules: a fill for numbers over a limit, a bold row where a formula holds, a colour scale, data bars, icon sets. The book keeps the rules, a session tells a grid what each cell of a window shows by them, the xlsx subpath reads and writes them, and structure operations move them as Excel does.

```js
import { openSession } from 'xlcalc'

const session = openSession({
  sheets: [{
    name: 'Sales',
    cells: { A1: { value: 120 }, A2: { value: 80 }, A3: { value: 45 } },
    conditionalFormats: [
      { type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['100'], style: { fill: { color: '#FFC7CE' }, font: { color: '#9C0006' } } },
      { type: 'colorScale', ranges: ['A1:A3'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    ],
  }],
})
const cells = session.cells('Sales', 'A1:A3')
console.log(cells.A1.look, cells.A2.look, cells.A3.look)
// → { style: { font: { color: '#9C0006' }, fill: { color: '#FFC7CE' } } } { style: { fill: { color: '#B39072' } } } { style: { fill: { color: '#F8696B' } } }
```

The rules are `SheetInput.conditionalFormats`, in priority order: the first one comes first. They change no value and take no part in the values hash; the input hash counts them.

## The rules

Each rule has `type`, `ranges` (the cells it applies to: `B2:B20`, `D2`, whole columns `F:F`, whole rows `3:5`, one or several, as Excel's "Applies to" lists them) and may have `stopIfTrue`. A rule that holds for a cell gives it its `style`, Excel's differential format: a fill, a font's colour, bold, italic, underline and strike (`bold: false` takes away a bold the cell has), borders and a number format. Formulas are written as cells hold them, without `=`; a leading one is dropped.

| type | the cell holds when | example |
|---|---|---|
| `cellIs` | its value compared with one formula, or two for `between` and `notBetween` | `{ operator: 'between', formulas: ['10', '$Z$1'] }` |
| `expression` | the formula holds | `{ formula: '$C2="late"' }` |
| `containsText`, `notContainsText`, `beginsWith`, `endsWith` | its text has `text`, in any case | `{ text: 'urgent' }` |
| `containsBlanks`, `notContainsBlanks`, `containsErrors`, `notContainsErrors` | it is blank or holds spaces alone; it holds an error | `{}` |
| `timePeriod` | it holds a date of the period, by the book's clock | `{ period: 'lastWeek' }` |
| `top10` | it is among the `rank` highest numbers (lowest with `bottom`), or the `rank` percent of them | `{ rank: 10, percent: true }` |
| `aboveAverage` | it is above the average (below with `below`, or equal too with `equal`), or `stdDev` deviations past it | `{ below: true }` |
| `duplicateValues`, `uniqueValues` | its value occurs more than once among the rule's cells, or once | `{}` |
| `colorScale` | it holds a number: a colour between the points around it | `{ points: [...] }` |
| `dataBar` | it holds a number: a bar as long as it is far between `min` and `max` | `{ min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6' }` |
| `iconSet` | it holds a number: the icon of the last threshold it reaches | `{ icons: '3Arrows', points: [...] }` |

A threshold of a colour scale, a data bar or an icon set is `{ type, value }`: `min` and `max` are the lowest and highest number of the rule's cells, `num` the number its formula gives, `percent` that many percent of the way from the lowest to the highest, `percentile` that percentile of the numbers, `formula` what its formula gives; a data bar's `autoMin` and `autoMax` are 0, or the lowest number below 0 (the highest above it).

### Where a formula reads

A rule's formulas are written for the top-left cell of the rectangle around its ranges, and a reference without `$` moves from there to each cell, as a copied formula does. With the ranges `H3:H6` and `J1:J3`, the formula `$A3>2` is written for H1, so H3 reads A5 and J2 reads A4. Excel reads it so, whatever the order of the ranges, and writes it back unchanged when it saves them as `J1:J3 H3:H6`.

A formula's answer holds as IF reads its condition: TRUE, a number other than 0 and the text TRUE hold; other text, a blank and an error do not. A formula that answers a range is not taken at the cell's row: `$A$1:$A$6>3` gives every cell the answer of A1. A formula may read other sheets and defined names.

### Comparisons

A cell value rule compares as `=` and `<` compare: text in any case, text above numbers and booleans above text, a blank cell as 0 or empty text. A blank cell that a rule's formula refers to is 0 as the formula's answer, as a cell's formula shows it: `lessThanOrEqual $Z$1` holds for a blank cell and for 0 when Z1 is blank, but not for empty text. `between` holds with its ends either way round.

### Text, blanks and periods

Excel writes a formula beside a text rule (`NOT(ISERROR(SEARCH("urgent",A1)))`) and computes the formula, not the text: a rule whose text says one thing and whose formula another goes by the formula. A rule without `formula` has Excel's. SEARCH takes `*` and `?` as wildcards, and a number reads as its text (`123` contains `2`; in ru-RU `1.5` reads `1,5` and contains a comma).

Blank and error rules go by the cell's value: blank, empty text and spaces alone are blank, an error is not blank. Excel writes `LEN(TRIM(A1))>0` for a rule of cells that are not blank, which an error does not satisfy, but it formats the error.

A time period is computed by the formula Excel writes for it (`FLOOR(A1,1)=TODAY()-1` for yesterday), by the book's clock; a week begins on Sunday.

### Statistics

Top, average, duplicate rules and the thresholds of colour scales, data bars and icon sets read the rule's cells together. Numbers count; text, booleans, blanks and errors are left out. A share of the numbers is cut (ten percent of 12 numbers is one of them), and numbers equal to the last one taken count too. The deviation of an average rule is the population's. Duplicates compare text in any case and text that reads as a number as that number (`"1"` is a duplicate of 1); empty text is a duplicate of a blank cell, and blank cells themselves are never formatted.

### Colour scales

A number gets the colour of the points around it, each channel moving from the lower point's by its share of the way, cut toward the lower point's: 2 between 1 (`#F8696B`) and 10 (`#63BE7B`) is `#E8726C`. A number below the first point gets its colour, one above the last point the last colour, and equal numbers all get the last colour.

### Data bars and icon sets

A data bar gives `look.bar`: where the bar starts and ends as fractions of the cell's width, its colour, fill and border, and for numbers on both sides of 0 the axis. An Excel 2010 bar (the default) starts at its axis, by default at the numbers' zero, and is from `minLength` to `maxLength` percent long, 0 and 100 by default; an Excel 2007 bar (`excel2007`) starts at the cell's left edge and is from 10 to 90 percent long. An icon set gives `look.icon`, the set and the icon's number in it in Excel's order (for `3Arrows` 0 is the red arrow down, 2 the green arrow up), the last one for the highest numbers, reversed by `reverse`; a threshold's own `icon` replaces it, `null` shows none. `showValue: false` gives `look.hideValue`.

## What a cell shows

`Session.cells(sheet, range)` gives each cell its `look` by the current values:

- `style`: the styles of the rules that hold, merged in priority order: a field an earlier rule sets wins (a red fill from the first rule and an italic from the second make a red italic cell); a colour scale's colour is a fill.
- `bar`, `icon` and `hideValue`, from the first data bar and icon set.

A rule with `stopIfTrue` that holds leaves out the rules after it for that cell. A number format of the look shows the cell's `text` by it, as Excel's `Range.Text` does. An empty cell with a look is listed with the value null, so a grid paints a highlighted empty row; in a rectangle of more than 65,536 cells only the cells the rectangle lists get looks. `compute` and `result()` give no looks.

A session's edits do not list cells whose look changed: an edit can recolour a whole colour scale. A host with rules on a sheet reads its window again after an edit, which costs what the window holds. The statistics of a rule's cells are kept until an edit changes a cell among them, and the code of a rule's formula is made once per cell along the axes it moves and kept with the session.

On a sheet of 34,000 rows with a formula rule over 4 columns, a three-colour scale and a data bar, a window of 20,000 cells takes about 60 ms the first time and 25 ms after (15 ms without the rules), an edit of a cell under 1 ms with the rules or without them, and the window read again after an edit that changes the data bar's numbers 35 ms, on an Apple M5 Pro.

`setLayout([{ sheet, conditionalFormats: [...] }])` replaces a sheet's rules and recomputes nothing.

## What is computed and what is kept

Every kind of rule is computed and kept. What a data bar and an icon set give a cell (the bar's length, the icon) is computed by Excel's documented rules, which Excel's object model does not report, so it was not measured. Kept and written but not drawn into the look: borders and a font's underline and strike are in `style` for a host to draw. A dxf's alignment, protection, font name and size, pattern and gradient fills are left out of the book and reported.

## Files

`readXlsx` reads the rules of `<conditionalFormatting>` with their differential formats (`<dxfs>` of styles.xml), and Excel 2010's rules of the sheet's extension list: a data bar's 2010 part (its fill, border, direction, negative bars and axis) joined to its 2007 rule by their id, and the rules only Excel 2010 has, the icon sets `3Stars`, `3Triangles` and `5Boxes`, custom icons and formulas over other sheets, with their own style. `writeXlsx` writes them the same way. A rule a pivot table's formatting holds, or a kind the model does not have, is left out and reported (`conditionalFormat`).

## Structure operations

The operations change rules as Excel changes them:

- Inserted lines that follow a line of a rule join it, as they take the format of the line before them; a rule after them moves.
- Deleted lines leave its ranges, and a rule with nothing left is gone. A rule over statistics (a colour scale, a data bar, an icon set, top, average and duplicate rules) that loses lines goes after the others, as Excel takes it out and puts it back.
- A move takes the rules of its cells along; the cells it lands on keep their own rules too. On another sheet the moved part is a rule of that sheet, first in its order. A rule whose formula the move rewrote goes after the others.
- A copy pastes its cells' rules as new rules over the pasted cells, first in the target's order, and the pasted cells lose the target's rules; a paste that goes on down a rule's own columns (a fill down) grows the rule instead.
- Formulas change as the formula of a cell at the rule's first cell would, and are made for a new first cell as a copy would make them. A reference to a deleted sheet is `#REF!`.

## How it was checked

`npx tsx bench/coverage/run.ts cond <set>` writes each book of `bench/coverage/cond/<set>.ts` with `writeXlsx`; Excel 16.113 for Mac opens it, applies the case's structure operation, computes the book, reads the display format of the listed cells (`Range.Text`, the fill and font colours, bold, italic: what conditional formats show) and saves the book. `npx tsx bench/windows/run.ts cond <set>` does the same in Excel 16.0.20430 for Windows through COM (`DisplayFormat`), and also makes rules through its object model for the set `made`. `test/cond-excel.test.ts` and `test/cond-win.test.ts` check without Excel that the engine shows what Excel showed, that the reader reads from each saved file the rules the engine holds after the case, and that the engine shows on the saved book what Excel showed.

| set | cases | what |
|---|---|---|
| `looks` | 7 | rules of every kind over numbers, text, booleans, blanks, empty text, errors and dates, priorities and `stopIfTrue`, a rule's number format; cases between readings of Excel's rules |
| `looks-ru` | 1 | text rules over Cyrillic and numbers with a decimal comma, text in the locale's order, dates, a number format, in ru-RU |
| `files` | 6 | every kind and option of rule in a file, Excel 2010's rules, many rules, whole columns and rows, books written after structure operations |
| `struct-held` | 24 | rows and columns inserted and deleted, moves and copies within a sheet and to another, sheets renamed and deleted |
| `held`, `held-2` | 9, 4 | held out after the sets above: the rules they gave on other cases |
| `both` | 3 | a sheet with conditional formats, data validation and grouped rows, written, with a row inserted and one deleted |
| `made` | 1 | 15 rules Excel for Windows made through COM (`FormatConditions.Add`, `AddTop10`, `AddAboveAverage`, `AddUniqueValues`, `AddColorScale`, `AddDatabar`, `AddIconSetCondition`), over two ranges too |

Both Excels loaded every book. The engine's answers were committed before each run. Excel for Mac showed the looks as predicted in 594 of 676 cells of `looks`, 191 of 200 of `files`, 1,699 of 1,885 of `struct-held` and 79 of 80 of `looks-ru`, and saved the rules of 12 of the 24 cases of `struct-held` as predicted; the misses gave the rules above (how a colour scale cuts its channels, errors left out of statistics, the cell a formula is written for, IF's reading of an answer, a blank operand, blank rules, duplicates, and how operations grow, keep and order rules), and the held-out sets, made after them, showed 708 of 708 and 336 of 336 cells as predicted and the rules of 11 of 13 cases, whose misses gave the order rules of top rules and where a paste grows a rule; `both`, 132 of 132 cells and its rules and data validation. Now the engine shows what Excel for Mac showed in 4,015 of 4,017 cells, the two pinned below. Excel for Windows, predicted after the Mac's answers, showed `looks`, `struct-held`, `held`, `held-2` and `both` as predicted in every cell, and agreed with Excel for Mac in 3,988 of the 4,017 cells both read (the others `#` counted in pixels, and the order of Latin and Cyrillic); the rules Excel for Windows made itself, read from its file, showed as Excel showed them in 203 of 204 cells. Of the rules each Excel saved, the reader reads what the engine holds after the case in every case, but the bar lengths and the order of one rule pinned below.

Not measured: a data bar's length and an icon, which no object model reports (the rules' parameters were checked, the bar and the icon follow Excel's documented rules); borders, underline and strike of a style (written and read back, not read from the display format); rules made through Excel for Mac's AppleScript, whose `make new format condition` stopped Excel; a copy across sheets of a rule with relative references beyond the cases above.

## Known differences

- Excel keeps an Excel 2010 data bar's lengths only when they are 0 and 100: a bar of 5 and 95 comes back without them, read as 10 and 90.
- After rows are deleted, Excel put a cell value rule whose reference the deletion lost after the others; the engine keeps its place, as Excel did deleting a column under another such rule.
- A colour exactly a third or two fifths of the way from a percentile point can differ by one in a channel: Excel's percentile differs in its last bit.
- Excel for Mac orders Cyrillic before Latin in ru-RU; the engine follows Excel for Windows, Latin first, in rules as in formulas.
- Excel for Windows lays text out in pixels, so `#` across a column can count otherwise ([docs/formats.md](formats.md)).
- A dxf's number format that is a built-in one comes back from Excel as its code (`0.00%` for 10).
