# Changelog

## Unreleased

The first public release. Versions up to 0.34.0 were internal builds and were not published to npm; their changes are folded into this entry, and the README describes the package as it is.

What is in it:

- `compute` and `createSession`: a book's formulas compiled once, to closures or with `execution: 'eval'`, recalculated incrementally, with deterministic values and a SHA-256 `hash` of them (configuration version `xlcalc/1;canon15`), and `inputHash` of the input.
- A whitelist of functions, each admitted after giving Excel's answer on all of its cases, computed in Excel 16 for Mac and checked again in Excel for Windows; arrays inside formulas and spilled arrays as in Excel 365; tables and structured references; names; links to other workbooks from their stored values.
- Sessions: edits, undo and redo, subscribers told of every change of the book (`subscribe`), structure operations (insert, delete and move rows, columns and sheets), typed entry as Excel reads it, fill and paste, Find and Replace, outline, AutoFit, conditional formats, data validation, and the text each cell shows in its number format.
- `xlcalc/xlsx`: reading and writing xlsx without dependencies, with what a book cannot hold listed as issues.
- `createInlineCalc` and `createCalcPool` from `xlcalc/worker`: computing without blocking, in portions on the same thread or in worker threads, with deadlines, cancellation and limits.
- Localization: formulas, number formats and messages in ru-RU through `xlcalc/locales/ru-RU`.
- `xlcalc/a1` and `xlcalc/text`: addresses, references of formula text and number format categories without the engine.
