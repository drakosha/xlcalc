# Moving to the new names

The package's calls were renamed so that a name says what a call does: a function that makes a session is `createSession`, a session's methods that read start with `get`, and a field is named by what it holds (`numberFormat`, `appearance`, `diagnostics`). The old names of functions, methods and input fields still work as deprecated synonyms, the same functions as the new ones, so a host moves call by call; the types mark each old name `@deprecated` with its new one. Fields of answers have the new names alone, and a few calls answer otherwise than before. This page lists both, a line each.

## Renamed

| old | new | the old name |
|---|---|---|
| `openSession(book, opts)` | `createSession(book, opts)` | a synonym |
| `calc.openSession(book, opts)` (`createInlineCalc`, `createCalcPool`) | `calc.createSession(book, opts)` | a synonym |
| `restructure(book, op)` | `editStructure(book, op)` | a synonym |
| `session.restructure(op)` | `session.editStructure(op)` | a synonym |
| `RestructureResult`, `SessionRestructureResult`, `AsyncRestructureResult` | `StructureResult`, `SessionStructureResult`, `AsyncStructureResult` | type aliases |
| `session.book()` | `session.getBook()` | a synonym |
| `session.layout()` | `session.getLayout()` | a synonym |
| `session.extent(sheet)` | `session.getUsedRange(sheet)` | a synonym |
| `session.cells(sheet, range)` | `session.getCells(sheet, range)` | a synonym |
| `session.validation(sheet, cell)` | `session.getValidation(sheet, cell)` | a synonym |
| `session.result()` | `session.getResult()` | a synonym |
| `session.hash()` | `session.getHash()` | a synonym |
| `Viewport.extent` | `Viewport.usedRange` | none |
| `call: 'restructure'` in the answer of `undo` and `redo` | `call: 'editStructure'` | none |
| `format` of cells, rows, columns, conditional formats' styles and their edits (`setCells`, `replaceCells`, `patchCells`, `setLayout`) | `numberFormat` | taken in input; both with two values are an `InputError` |
| `format` of `displayValue`'s and `entryWidth`'s options and of `parseEntry`'s `cell` | `numberFormat` | taken |
| `format` of a result's cell (`CellResult`, `CellView`) and of a book a session, `parseEntry` or `readXlsx` gives | `numberFormat` | none |
| `CellView.look`, the type `ConditionalLook` | `CellView.appearance`, `ConditionalAppearance` | none; the type an alias |
| `result.issues` and `result.nameIssues` | `result.diagnostics`, one list: the cells with an error (`kind: 'cell'`), then the names with a problem (`kind: 'name'`), each with the fields it had | none |
| the object `engine` | the functions by name: `import { compute, createSession } from 'xlcalc'` | deprecated whole |
| `fromLegacy` from `xlcalc` | `fromLegacy` from `xlcalc/xlsx` | a synonym |

A book hashes as before whichever name it gives the number format: `inputHash` writes the field by its old name, and neither the values hash nor `inputHash` depends on `diagnostics`.

## Answering otherwise

| before | now |
|---|---|
| `setCells` replaced each cell whole | `setCells` sets the fields an edit gives and keeps the cell's others, as Excel keeps a cell's look when a value is typed over it; `replaceCells` replaces a cell whole |
| `input: null` in `setCells` removed the cell | it clears the content and the hyperlink and keeps the style, the number format, the note and `meta`, as Excel's Delete key; `replaceCells` with null removes the cell |
| an edit answered the changed cells, an array | every call that changes the book answers an `EditResult`: `{ revision, changes, layout?, styles?, moved?, repaint? }` |
| `readXlsx`'s `cached` was a list of sheets with a value by address (`cached[0].cells.B1` is 200) | `cached` has the shape of a result: `cached.sheets[0].cells.B1` is `{ value: 200 }` |
| ru-RU was built in and the root exported `ruRU` | `import { ruRU } from 'xlcalc/locales/ru-RU'`, and `registerLocale(ruRU)` once for books in ru-RU and for its tag; a locale object needs no loading where no book is needed |
| `writeXlsx(book, { result })` | `writeXlsx({ book, result })`, or `writeXlsx(session.getSnapshot())`; the older form writes the same file |
| `validateEntry` refused with `{ ok: false, rule }` | `{ ok: false, code: 'validation', message, rule, severity }`, as every check refuses |
| after a structure operation `changes` listed the cells whose value changed | they also list the cells whose formula's text, own style or own number format differ from the cell that was at their old place, and `moved` says how addresses moved, so a window moves by `moved` and applies `changes` instead of reading itself again |
| `patchCells` listed the cells whose shown text changed | it also lists the cells whose own style or number format it set, with `style` and `numberFormat` |
| a synchronous session's `editStructure` answered the book too | it answers `mapCell` and what an edit answers without the book, as an asynchronous session's does; `getBook()` gives the book |
| `mapCell` threw `InputError` for a sheet the book before the operation does not have | it answers null, as for a deleted cell, in every place it is given: `editStructure`'s answer, a session's, and `undo`'s and `redo`'s |
| the books of answers (`getBook()`, `getLayout()`, `getSnapshot().book`, `editStructure(book, op).book`, `readXlsx(…).book`) were typed `BookInput`, so reading `format` from one compiled and gave undefined | they are typed `Book` (its sheets `Sheet`, its cells `Cell`; `BookLayout` without the cells), which have no `format`: reading the old name does not compile; a `Book` is a `BookInput` and goes back into every call that takes a book |
| `getViewport` gave a cell's `style` and `effectiveStyle` as style objects | as numbers of the book's `styles`, as `getCells`, the changes of edits and the book give a style; `Viewport.styles` has the objects of the styles the window shows, by number |
| `explain`'s text wrote an error by its code (`#DIV0`) and an empty cell as `null` | the error as the cells show it in the book's locale (`#DIV/0!`, `#ДЕЛ/0!`) with the cell it began in (`(error from 'S'!A2)`), and an empty cell as `empty`; the tree's error nodes have `origin` |

## How it was checked

`test/old-names.test.ts` calls every old name of a function and a method in a synchronous session, in the same thread and in a thread of the pool, and checks that it is the function of its new name and answers the same. `test/number-format-name.test.ts` gives every input that takes a number format the old name, both names and two different values, in the three kinds of session with closures and eval, and checks the input hash of the synthetic book and of a book with formats of cells, rows, columns and conditional formats against the hashes they had before the field was renamed. `test/result-diagnostics.test.ts` checks the list of diagnostics, its order and the hashes of a book with errors and bad names, and the appearance of a cell's conditional formats. Of the answers that changed their shape, in the three kinds of session with closures and eval: `test/session-viewport.test.ts` compares a window's styles by number with `getCells` and the book's table, `test/session-geometry.test.ts` checks that `editStructure` answers without the book and that `mapCell` answers null for a sheet it does not know, after an undo and a redo too, and `test/number-format-name.test.ts` reads `format` from every book an answer gives under `@ts-expect-error`, which `npm run typecheck` refuses to compile if a type lets it through.
