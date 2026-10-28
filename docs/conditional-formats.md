# Conditional formats

A sheet can contain Excel's conditional formatting rules: a fill for numbers above a limit, a bold row where a formula is true, a color scale, data bars, icon sets. The workbook keeps the rules, a session tells a grid how each cell of a viewport is displayed under them, the xlsx subpath reads and writes them, and structure operations move them as Excel does.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'Sales',
    cells: { A1: { value: 120 }, A2: { value: 80 }, A3: { value: 45 } },
    conditionalFormats: [
      { type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['100'], style: { fill: { color: '#FFC7CE' }, font: { color: '#9C0006' } } },
      { type: 'colorScale', ranges: ['A1:A3'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    ],
  }],
})
const cells = session.getCells('Sales', 'A1:A3')
console.log(cells.A1.appearance, cells.A2.appearance, cells.A3.appearance)
// → { style: { font: { color: '#9C0006' }, fill: { color: '#FFC7CE' } } } { style: { fill: { color: '#B39072' } } } { style: { fill: { color: '#F8696B' } } }
```

The rules are `SheetInput.conditionalFormats`, in priority order: the first one has the highest priority. They change no value and are not part of the values hash; the input hash includes them.

## The rules

Each rule has `type` and `ranges` (the cells it applies to: `B2:B20`, `D2`, entire columns `F:F`, entire rows `3:5`, one or several, as Excel's "Applies to" lists them), and can have `stopIfTrue`. A rule that applies to a cell gives it its `style`, Excel's differential format: a fill, a font color, bold, italic, underline and strikethrough (`bold: false` removes bold that the cell has), borders and a number format. Formulas are written as cells store them, without `=`; a leading `=` is dropped.

| type | the rule applies when | example |
|---|---|---|
| `cellIs` | the cell's value, compared with one formula, or two for `between` and `notBetween` | `{ operator: 'between', formulas: ['10', '$Z$1'] }` |
| `expression` | the formula is true | `{ formula: '$C2="late"' }` |
| `containsText`, `notContainsText`, `beginsWith`, `endsWith` | the cell's text contains `text`, case-insensitively | `{ text: 'urgent' }` |
| `containsBlanks`, `notContainsBlanks`, `containsErrors`, `notContainsErrors` | the cell is blank or contains only spaces; the cell contains an error | `{}` |
| `timePeriod` | the cell contains a date within the period, by the workbook's clock | `{ period: 'lastWeek' }` |
| `top10` | the cell is among the `rank` highest numbers (lowest with `bottom`), or the top `rank` percent of them | `{ rank: 10, percent: true }` |
| `aboveAverage` | the cell is above the average (below with `below`, or also equal with `equal`), or `stdDev` standard deviations beyond it | `{ below: true }` |
| `duplicateValues`, `uniqueValues` | the cell's value occurs more than once among the rule's cells, or exactly once | `{}` |
| `colorScale` | the cell contains a number: a color between the points around it | `{ points: [...] }` |
| `dataBar` | the cell contains a number: a bar whose length reflects where it falls between `min` and `max` | `{ min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6' }` |
| `iconSet` | the cell contains a number: the icon of the last threshold it reaches | `{ icons: '3Arrows', points: [...] }` |

A threshold of a color scale, a data bar or an icon set is `{ type, value }`: `min` and `max` are the lowest and highest numbers among the rule's cells, `num` is the number its formula returns, `percent` is that many percent of the way from the lowest to the highest, `percentile` is that percentile of the numbers, and `formula` is what its formula returns; a data bar's `autoMin` and `autoMax` are 0, or the lowest number if it is below 0 (the highest if it is above).

### Where a formula reads

A rule's formulas are written for the top-left cell of the rectangle enclosing its ranges, and a reference without `$` shifts from there to each cell, as a copied formula does. With the ranges `H3:H6` and `J1:J3`, the formula `$A3>2` is written for H1, so H3 reads A5 and J2 reads A4. Excel reads it this way whatever the order of the ranges, and writes it back unchanged when it saves them as `J1:J3 H3:H6`.

A formula's result counts as true the way IF evaluates its condition: TRUE, a number other than 0 and the text TRUE are true; other text, a blank and an error are not. A formula that returns a range is not intersected with the cell's row: `$A$1:$A$6>3` gives every cell the result for A1. A formula can read other sheets and defined names.

### Comparisons

A cell value rule compares as `=` and `<` compare: text case-insensitively, text greater than numbers and booleans greater than text, and a blank cell as 0 or empty text. A blank cell that a rule's formula refers to is 0 as the formula's result, as a cell's formula displays it: `lessThanOrEqual $Z$1` applies to a blank cell and to 0 when Z1 is blank, but not to empty text. `between` works with its bounds in either order.

### Text, blanks and periods

Excel writes a formula alongside a text rule (`NOT(ISERROR(SEARCH("urgent",A1)))`) and evaluates the formula, not the text: a rule whose text says one thing and whose formula says another follows the formula. A rule without `formula` gets Excel's. SEARCH treats `*` and `?` as wildcards, and a number is matched as its text (`123` contains `2`; in ru-RU `1.5` becomes `1,5` and contains a comma).

Blank and error rules depend on the cell's value: blank, empty text and only spaces count as blank, and an error is not blank. Excel writes `LEN(TRIM(A1))>0` for a rule matching non-blank cells, which an error does not satisfy, but it formats the error anyway.

A time period is evaluated by the formula Excel writes for it (`FLOOR(A1,1)=TODAY()-1` for yesterday), with the workbook's clock; a week starts on Sunday.

### Statistics

Top, average and duplicate rules, and the thresholds of color scales, data bars and icon sets, read the rule's cells together. Numbers count; text, booleans, blanks and errors are ignored. A share of the numbers is truncated (ten percent of 12 numbers is one of them), and numbers equal to the last one included are included too. The standard deviation of an average rule is the population standard deviation. Duplicates compare text case-insensitively, and text that looks like a number as that number (`"1"` is a duplicate of 1); empty text is a duplicate of a blank cell, and blank cells themselves are never formatted.

### Color scales

A number gets a color between the points around it, each channel moving from the lower point's color by its share of the way, truncated toward the lower point's: 2 between 1 (`#F8696B`) and 10 (`#63BE7B`) is `#E8726C`. A number below the first point gets its color, one above the last point gets the last color, and if all the numbers are equal they all get the last color.

### Data bars and icon sets

A data bar gives `appearance.bar`: where the bar starts and ends as fractions of the cell's width, its color, fill and border, and, for numbers on both sides of 0, the axis. An Excel 2010 bar (the default) starts at its axis, by default at zero, and is from `minLength` to `maxLength` percent long, 0 and 100 by default; an Excel 2007 bar (`excel2007`) starts at the cell's left edge and is from 10 to 90 percent long. An icon set gives `appearance.icon`: the set and the icon's index in it in Excel's order (for `3Arrows`, 0 is the red down arrow and 2 the green up arrow), the last one for the highest numbers, reversed by `reverse`; a threshold's own `icon` replaces it, and `null` shows none. `showValue: false` gives `appearance.hideValue`.

## What a cell displays

`Session.getCells(sheet, range)` gives each cell its `appearance` for the current values (`getViewport` returns the same):

- `style`: the styles of the rules that apply, merged in priority order: a field set by a higher-priority rule wins (a red fill from the first rule and italic from the second make a red italic cell); a color scale's color is a fill.
- `bar`, `icon` and `hideValue`, from the first data bar and icon set.

A rule with `stopIfTrue` that applies skips the rules after it for that cell. A number format in the appearance formats the cell's `text`, as Excel's `Range.Text` does. A built-in format given by number is displayed using its en-US code in every locale, as Excel interprets a differential format by the code it stores: `{ numberFormat: 14 }` displays `3.15.2026` in ru-RU, where a cell's own format 14 displays `15.03.2026`. An empty cell with an appearance is listed with the value null, so a grid can paint a highlighted empty row; in a rectangle of more than 65,536 cells, only the cells the rectangle lists get one. `compute` and `getResult()` return none.

### A saved version

A result has no appearance, so `viewportOf` returns a viewport over a saved workbook and its result without one, unless `appearance: true` is given. Then the rules are evaluated as a session evaluates them, over a state of the workbook whose cells contain the result's values, without calculating its formulas, and with the result's clock, so a time period or TODAY in a rule follows the day the version was calculated. The viewport is then the one `getViewport` returns in a session of the workbook with that clock, including the appearance and the text in a rule's number format, except where the result holds less than a session:

- A result rounds a number to 15 digits: `0.1+0.2` is 0.3 there, a duplicate of 0.3 and tied with it for a top rule, where a session's full-precision number is neither, and a rule formula that calculates with it can come out differently; a comparison in a cell value rule treats them as equal in both cases.
- A cell of the state contains a value, not a formula: SUBTOTAL and AGGREGATE in a rule's formula count a cell that contains a SUBTOTAL of its own, which they skip in a session.
- A result holds no values of external functions: a rule formula that calls one returns an error, and the rule formats no cell because of it.

A session's edits do not list the cells whose appearance changed: an edit can recolor an entire color scale. Instead, they return the rules' ranges as `repaint` after a value or a rule changed, and a host reads its viewport again where it overlaps them, which costs as much as the viewport contains ([Editing a grid](edits.md#what-an-edit-returns)). The statistics of a rule's cells are kept until an edit changes one of those cells, and the code of a rule's formula is compiled once and kept with the session. Where only its references and position-dependent functions such as ROW() shift it, one piece of code reads the cell it runs at; a formula with a name, `@`, a structured reference, a relative reference to other sheets, or LOOKUP over a shifting reference is compiled once per cell along the axes it shifts, and a session keeps about 8 MB of such code, evicting the least recently used first.

On a sheet of 34,000 rows with a formula rule over 4 columns, a three-color scale and a data bar, a viewport of 20,000 cells takes about 60 ms the first time and 25 ms afterwards (15 ms without the rules), editing a cell takes under 1 ms with or without the rules, and reading the viewport again after an edit that changes the data bar's numbers takes 35 ms, on an Apple M5 Pro.

`setLayout([{ sheet, conditionalFormats: [...] }])` replaces a sheet's rules and recalculates nothing.

## What is evaluated and what is kept

Every kind of rule is evaluated and kept. What a data bar and an icon set give a cell (the bar's length, the icon) is calculated by Excel's documented rules, which Excel's object model does not expose, so it was not verified. Kept and written but not resolved into the display: borders and a font's underline and strikethrough are in `style` for a host to draw. A dxf's alignment, protection, font name and size, and pattern and gradient fills are left out of the workbook and reported.

## Files

`readXlsx` reads the rules in `<conditionalFormatting>` with their differential formats (`<dxfs>` in styles.xml), and the Excel 2010 rules in the sheet's extension list: a data bar's 2010 part (its fill, border, direction, negative bars and axis), joined to its 2007 rule by id, and the rules only Excel 2010 has (the icon sets `3Stars`, `3Triangles` and `5Boxes`, custom icons, and formulas referring to other sheets), with their own style. `writeXlsx` writes them the same way. A rule that belongs to a pivot table's formatting, or of a kind the model does not support, is left out and reported (`conditionalFormat`). Writing over the original file, `writeKept` keeps such rules among the workbook's ([Writing over the original file](xlsx-keep.md#conditional-formats-the-workbook-does-not-read)).

## Structure operations

The operations change rules as Excel changes them:

- Inserted rows or columns that follow a row or column of a rule join it, as they get the format of the one before them; a rule after them moves.
- Deleted rows and columns are removed from its ranges, and a rule with nothing left is removed; a range that ends on the sheet's last row or column keeps ending there, even a single cell (`B1048576` becomes `B1048575:B1048576` after a row above it is deleted). A statistics rule (a color scale, a data bar, an icon set, or a top, average or duplicate rule) that loses rows or columns moves after the others, as Excel removes it and adds it back.
- A move takes the rules of its cells along; the destination cells keep their own rules too. On another sheet, the moved part becomes a rule of that sheet, first in its order. A rule whose formula the move rewrote moves after the others.
- A copy pastes its cells' rules as new rules over the pasted cells, first in the destination's order, and the pasted cells lose the destination's rules; a paste that continues down a rule's own columns (a fill down) extends the rule instead.
- Formulas change as the formula of a cell at the rule's first cell would, and are rewritten for a new first cell as a copy would rewrite them, except that a reference moved past the sheet's edge comes round the other edge (`$B3<>$B2` filled up to A1 becomes `$B1<>$B1048576`), as the rule computes it, instead of becoming `#REF!`. A reference to a deleted sheet becomes `#REF!`.

## Checked against Excel

Tested in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US and ru-RU, on 74 workbooks, each set predicted before Excel ran it: the engine displays what Excel displayed in 4,015 of 4,017 cells, and the other two are known differences. Not tested: a data bar's length and an icon, which no object model exposes, and the borders, underline and strikethrough of a rule's style.

## Known differences

- Excel keeps an Excel 2010 data bar's lengths only when they are 0 and 100: a bar of 5 and 95 comes back without them, and is read as 10 and 90.
- After rows are deleted, Excel moved a cell value rule whose reference the deletion lost after the others; the engine keeps its position, as Excel did when deleting a column under another such rule.
- A color exactly a third or two fifths of the way from a percentile point can differ by one in a channel: Excel's percentile differs in its last bit.
- Excel for Mac sorts Cyrillic before Latin in ru-RU; the engine follows Excel for Windows, Latin first, in rules as in formulas.
- Excel for Windows lays out text in pixels, so the number of `#` across a column can differ ([docs/formats.md](formats.md)).
- A dxf number format that is a built-in one comes back from Excel as its code (`0.00%` for 10).
- A gradient fill is read as no fill. A rule fill with `patternType="none"`, which removes the cell's own fill in Excel, and a solid fill without a color, which Excel displays as white on a cell without a fill, are read as no fill. Each is reported.
- Excel for Mac writes the short month names of ru-RU differently (`15-март-26`); the engine follows Excel for Windows (`15-мар-26`).
