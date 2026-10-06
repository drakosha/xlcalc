# formulajs coverage against Excel

Measured on 5 October 2026, branch `probe/coverage`. The question: how many functions of `@formulajs/formulajs` 4.6.1 give Excel's answer when called through a shared adapter following the engine's runtime rules (`src/runtime.ts`), and how many without it. The reference for the engine is Excel en-US. Excel ru-RU (for an interface with Russian input) and LibreOffice en-US (the oracle of the engine's tests) were computed alongside. The measurement does not touch the engine's whitelist.

The occasion was the WB1 measurement in the application the engine came from (its `docs/probes/WB1-engine.md`, section "List semantics"): bare formulajs went wrong not on plain calls but at the edges (wildcards and date text in criteria, text and booleans in ranges, case, JS Date in the local zone). So every function here gets edge cases besides its typical calls.

The measurement's code, the rules for working with Excel and the command to repeat it: `bench/coverage/README.md`. Raw results by column: `test/excel/coverage/` (`cases.json`, the cases with their formulas; `excel-en.json`, `excel-ru.json`) and `bench/coverage/data/` (`formulajs.json`, `libreoffice.json`, `summary.json` with the verdicts by case and function). The former Excel columns `excel-en-legacy.json`, `excel-ru-legacy.json` and `excel-en-ru-ui.json` were removed; they are in git at commit `85522cf`, under `docs/coverage-formulajs/`. The tables between the `auto` marks below are rebuilt by `npx tsx bench/coverage/run.ts report`.

## Summary

The set has 402 functions and 5,275 cases. The Excel en-US reference was computed with the English interface and locale (`-AppleLanguages (en) -AppleLocale en_US`). Through the adapter all cases match it for 105 functions; for 213 the typical cases match and edge cases differ; for 81 typical cases differ already; 3 functions hang on a huge argument. Bare formulajs matches fully for 55 functions. By case the adapter raises the match from 3,355 to 3,976 (from 64 to 75 percent); 817 of 965 typical cases match, and 3,159 of 4,310 edge cases. A complex number as text with an imaginary unit is compared by parts within the same tolerance as numbers (since 7 October; before that character for character, and 15 IM* functions differed only by writing 17 digits). The proposal (last section): take 105 functions as they are, 40 after the adapter's shared rules, 169 after the adapter's rules and a short argument check per function, write 85 anew, close the 3 that hang with a guard.

The report's first edition computed the en-US column with Excel's Russian interface. The interface language changes computation, not only display (section "The interface language"): 32 cases of 23 functions, which moved the result from 97 to 105 functions matching fully.

Since 6 October 2026 the Excel en-US and ru-RU columns are computed as dynamic array formulas, as Excel 365 saves them; the former ones, with formulas of the old kind, are in git (`docs/coverage-formulajs/excel-en-legacy.json` and `excel-ru-legacy.json` at commit `85522cf`). Measured again, 9 of the 5,275 answers changed in each column, all where the old formula intersected a range with its own row or showed the first value of an array: ROW and COLUMN of a range (ROW#1, ROW#2, ROW#4, COLUMN#2, COLUMN#4: a number before, `#SPILL!` now), INDEX with zero or a blank cell as the number (INDEX#6, INDEX#7, INDEX#13: `#VALUE!` before, `#SPILL!` now), TYPE of a range (TYPE#4: 16 before, 64 now). The cases with an array answer, written before as array formulas (Ctrl+Shift+Enter), did not change. One more case matches through the adapter (TYPE#4), and TYPE moved from "own" to "adapter".

<!-- auto:summary -->
| | through the adapter | bare formulajs |
|---|---|---|
| functions in the set | 402 | 402 |
| all cases match Excel en-US | 105 | 55 |
| only the typical cases match | 213 | 259 |
| typical cases differ already | 81 | 85 |
| hangs (on one of the cases) | 3 | 3 |
| cases that match | 3976 of 5275 | 3355 of 5275 |

Within the soft tolerance 1e-6, 3 more match fully through the adapter: F.INV.RT, GAMMA, TDIST.
The adapter brought 50 functions to a full match where bare formulajs differs: ACOTH, ATANH, AVEDEV, AVERAGE, BINOM.DIST.RANGE, BINOMDIST, CHOOSE, CHOOSECOLS, CHOOSEROWS, CLEAN, COLUMNS, COSH, DROP, EXP, EXPAND, F.DIST.RT, FIND, FISHER, FIXED, GAMMALN, GAMMALN.PRECISE, IF, IFS, ISPMT, LEFT, LN, LOG, LOG10, MAX, MEDIAN, MIN, NORM.DIST, NORM.S.DIST, NORMDIST, NOT, NPV, PDURATION, PHI, PROB, PRODUCT, REPT, RIGHT, ROUNDDOWN, ROUNDUP, ROWS, SINH, SUBTOTAL, SUM, TRIM, UNICODE.
Worse through the adapter than bare: 15: DAYS360 (9 against 10), GCD (8 against 10), IMCSC (6 against 7), IMCSCH (6 against 7), IRR (4 against 5), LCM (6 against 8), MULTINOMIAL (5 against 8), NETWORKDAYS (5 against 6), NETWORKDAYS.INTL (7 against 9), TBILLEQ (11 against 12), TEXT (8 against 9), VALUE (10 against 11), VLOOKUP (11 against 12), WEEKDAY (9 against 11), WEEKNUM (11 against 12).
<!-- /auto:summary -->

Most of the adapter's gain comes from the SUM rule for folds (only numbers from a range, a direct argument converted), JS Date to a serial number, NaN and infinity to `#NUM!` and exceptions to `#VALUE!`. Bare formulajs gives NaN or infinity 228 times, throws 132 times and gives a JS Date instead of a number 65 times. Through the adapter no NaN or JS Date is left; exceptions become `#VALUE!`, and in 94 cases Excel answers otherwise.

For 15 functions the adapter is worse than bare formulajs, and these show directly which runtime rules must not be applied to everything:
- a date argument through `toNum` does not take date text (`YEAR("2025-01-31")` is 2025 in Excel, `#VALUE!` through the adapter), while bare formulajs reads such text itself; these are DAY, MONTH, YEAR, WEEKDAY, WEEKNUM, ISOWEEKNUM, DAYS, DAYS360, ACCRINT, COUPDAYS, TBILL* and others;
- truncating whole-number arguments is wrong for the bit functions: `BITAND(1,2.7)` is `#NUM!` in Excel;
- the SUM rule is wrong for GCD, LCM and MULTINOMIAL: Excel gives `#VALUE!` on text in a range and on a boolean, 0 or 1 on a blank range;
- formulajs takes the holidays of NETWORKDAYS and WORKDAY as serial numbers and refuses a two-dimensional array of JS Dates (`#VALUE!` on a typical case).

The runtime's `toBool` takes the text `"TRUE"`, and so does Excel en-US with the English interface: `IF("TRUE","yes","no")` is `"yes"`, `NOT("TRUE")` is FALSE. The engine's own IF (`compute` gives `"yes"`) matches the reference. With the Russian interface Excel answers `#VALUE!` to the same text and takes `"ИСТИНА"` instead.

Kinds of difference by case and function:

<!-- auto:kinds -->
| kind of difference | cases through the adapter | functions through the adapter | cases bare | functions bare |
|---|---|---|---|---|
| number | 349 | 107 | 373 | 121 |
| value instead of an error | 334 | 164 | 379 | 183 |
| other error code | 278 | 113 | 441 | 212 |
| error instead of a value | 191 | 98 | 215 | 112 |
| exception | 94 | 29 | 132 | 59 |
| string | 22 | 10 | 23 | 10 |
| string: number format | 14 | 14 | 0 | 0 |
| string: case | 8 | 3 | 12 | 8 |
| other type | 6 | 5 | 3 | 3 |
| hangs | 3 | 3 | 3 | 3 |
| NaN or infinity | 0 | 0 | 228 | 101 |
| other | 0 | 0 | 46 | 17 |
| JS Date | 0 | 0 | 65 | 6 |
<!-- /auto:kinds -->

"value instead of an error" and "other error code" are almost all edge cases: Excel checks the domain and answers `#NUM!`, `#DIV/0!` or `#N/A`, while formulajs computes on or answers with another code. "number" falls in two parts: at the edges it is a different reading of the input (a boolean, a blank cell, the dates 0 and 60), on typical cases a wrong formula or an old name with another function's semantics (section "Where typical cases differ"). "string: number format" means the same numbers written differently (formulajs prints a real IM* answer with 17 digits, Excel with 15; a complex one with an imaginary unit is compared by parts), "string: case" means hexadecimal digits in lower case.

Within the soft tolerance 1e-6 three more functions match fully: F.INV.RT, GAMMA and TDIST (GAMMA in formulajs is accurate to about 1e-8). XIRR matches its typical cases only softly (0.3733625352 against 0.3733625335) and differs at the edges too.

## How it was measured

**The set.** formulajs has 406 names of Excel's function reference as functions (dotted names are taken by path: `fj.NORM.S.DIST`). Four volatile ones are left out: NOW, RAND, RANDBETWEEN, TODAY. formulajs has no environment functions (CELL, INFO). formulajs names Excel lacks (synonyms such as CEILINGMATH and CHIDISTRT) are listed at the end and are not measured. The reference: Microsoft's page "Excel functions (alphabetical)", 521 names with the `*B` variants, saved as `bench/coverage/excel-functions.json`.

**The cases.** The signature table (`bench/coverage/signatures.ts`) gives each function its argument kinds and one to seven typical calls, mostly the examples of Excel's reference. From the first typical call the generator makes edge cases, putting values into an argument by its kind: for numbers numeric text, a boolean, a blank cell, an error input, text, a negative number, zero, a fraction, a large number, numeric text from a cell, a date, `"1.5"`; for dates ISO text, `"1/2/2025"`, 2024-02-29, 1900-02-28, the false 1900-02-29 (60), 1900-03-01, 0, −1, a date with a time; for text Cyrillic, case, a number, a boolean, blanks, spaces; for folds a mixed range (a number, numeric text, text, a boolean, a blank), a range with an error, a blank range, a boolean and text as arguments, a reference to numeric text. Criteria (SUMIF and its kin) and databases (the D functions) follow a shared hand-written template: the wildcards `*`, `?`, `~*`, Cyrillic in any case, "й" against "и", date text, `<>`, an empty criterion, an error in the range, a field in any case, field number 0. The choice is deterministic (seed 20261005), up to 15 cases per function plus the elements of an array answer: the median is 15, and 15 functions have fewer than eight (functions without arguments and reference functions). 965 typical cases and 4,310 edge cases in all.

**The columns.**
- formulajs through the adapter: the shared rules of `runtime.ts`, spread by argument kind. An error in a scalar argument or in a numeric array is the answer at once; a number through `toNum`, a whole number also truncated (as in ROUND), text through `toText`, a boolean through `toBool`; a date as a JS Date at local midnight (a negative one is `#NUM!`), a time with the fraction of the day; a fold of numbers by the SUM rule; other ranges as two-dimensional arrays, errors inside them as formulajs error objects. The answer: an error to its code, a Date to the serial number of the local date, NaN and infinity to `#NUM!`, an exception to `#VALUE!`, anything else (null, an array) to `#VALUE!`.
- bare formulajs: arguments as they lie in the cells (one cell as its value, a range as a two-dimensional array, a blank as `null`, a date as a serial number, an error as a formulajs object), the answer as it is.
- Excel en-US and Excel ru-RU: the same book, Excel 16.113 for Mac. For en-US the interface language and the locale come from the launch arguments (`-AppleLanguages (en) -AppleLocale en_US`), for ru-RU they are the system's (Russian).
- LibreOffice en-US: the same xlsx, an en-US profile, as in `test/oracle/libreoffice.ts`.

**Comparison.** Numbers within |a−b| ≤ 1e-9·max(1,|a|), and apart from that a soft 1e-6; strings and booleans exactly; errors by code (`ERROR.TYPE` in Excel, not the error's text). The reference for the engine's formulas is Excel en-US: the English interface and locale. Excel ru-RU was computed with the Russian interface and locale (the system's), as for a user with Russian input.

**The book.** Functions newer than Excel 2007 are written with `_xlfn.` (SORT with `_xlfn._xlws.`); no function gave `#NAME?` in Excel, so the prefix list is right. An array answer is compared by element (`INDEX(F(…),i,j)`) and by size (`ROWS`, `COLUMNS`). All cases are written as dynamic array formulas (`cm="1"`, the XLDAPR metadata); until 6 October typical cases were written as formulas of the old kind and cases with an array answer as array formulas, since otherwise old-kind Excel applied implicit intersection to TRANSPOSE's argument. The LibreOffice column was computed from a book of the old kind. A one-dimensional formulajs array counts as a column.

<!-- auto:meta -->
- formulajs 4.6.1, Node v26.7.0, time zone Europe/Moscow, 15.914 s to compute;
- Excel 16.113.3 for Mac: en-US 17.6 s, ru-RU 18.3 s (14 parts, Excel's restarts not counted);
- LibreOffice 26.2.4.2, en-US profile, 2.891 s.
<!-- /auto:meta -->

A full `run.ts all` run with two Excel restarts took 66 s. A second run from scratch gave the same values in all five columns.

## By function

"match" counts cases; the "result" through the adapter is all, typical only, typical differ or hangs; "N more soft" means that N of the differing cases match within 1e-6. The last column holds the tags of the cases where the adapter differs from Excel en-US.

<!-- auto:functions -->
| function | category | cases | match through the adapter | match bare | result through the adapter | kind of difference through the adapter | on which cases |
|---|---|---|---|---|---|---|---|
| BETADIST | Compatibility | 15 | 9 | 7 | typical only | number 3, value instead of an error 3 | numeric text; boolean; fractional; blank cell; zero; negative |
| BETAINV | Compatibility | 15 | 10 | 7 | typical only | value instead of an error 5 | boolean; numeric text; blank cell; zero; negative |
| BINOMDIST | Compatibility | 15 | 15 | 11 | all |  |  |
| CEILING | Compatibility | 15 | 15 | 15 | all |  |  |
| CHIDIST | Compatibility | 15 | 4 | 2 | typical differ | number 9, value instead of an error 2 | typical; numeric text; boolean; numeric text in a cell; blank cell; negative; date in a cell; zero; fractional |
| CHIINV | Compatibility | 15 | 4 | 2 | hangs | number 6, value instead of an error 4, hangs 1 | typical; boolean; number text with a point; blank cell; large; negative; zero; above 1; date in a cell; numeric text |
| CHITEST | Compatibility | 4 | 1 | 0 | typical differ, 1 more soft | number 2, other error code 1 | typical; range: mixed; different size |
| COVAR | Compatibility | 9 | 3 | 1 | typical only | value instead of an error 3, error instead of a value 2, other error code 1 | range: mixed; blank range; different size |
| CRITBINOM | Compatibility | 15 | 9 | 7 | typical only | value instead of an error 4, other error code 2 | negative; above 1; blank cell; zero |
| EXPONDIST | Compatibility | 15 | 12 | 10 | typical only | value instead of an error 3 | zero; negative |
| FDIST | Compatibility | 15 | 5 | 2 | typical differ | number 9, value instead of an error 1 | typical; numeric text; boolean; large; numeric text in a cell; blank cell; negative; date in a cell; zero; fractional |
| FINV | Compatibility | 15 | 11 | 7 | typical differ | number 3, error instead of a value 1 | typical; boolean; date in a cell |
| FLOOR | Compatibility | 15 | 15 | 15 | all |  |  |
| FTEST | Compatibility | 9 | 2 | 0 | typical differ | number 3, error instead of a value 2, value instead of an error 1, other error code 1 | typical; range: mixed; different size; blank range |
| GAMMADIST | Compatibility | 15 | 14 | 7 | typical only | other error code 1 | negative |
| GAMMAINV | Compatibility | 15 | 13 | 8 | typical only | value instead of an error 2 | boolean; large |
| HYPGEOMDIST | Compatibility | 15 | 12 | 6 | typical only | exception 3 | negative; large |
| LOGINV | Compatibility | 15 | 9 | 7 | typical only | value instead of an error 6 | boolean; zero; blank cell; negative; above 1 |
| LOGNORMDIST | Compatibility | 15 | 3 | 1 | typical differ | number 9, value instead of an error 3 | typical; numeric text; boolean; number text with a point; blank cell; negative; zero; fractional |
| NEGBINOMDIST | Compatibility | 15 | 12 | 9 | typical only | value instead of an error 3 | negative; above 1 |
| NORM.INV | Compatibility | 15 | 9 | 6 | typical only | value instead of an error 6 | boolean; blank cell; negative; zero; above 1 |
| NORMDIST | Compatibility | 15 | 15 | 13 | all |  |  |
| NORMSDIST | Compatibility | 13 | 2 | 1 | typical differ | number 11 | typical; numeric text; boolean; blank cell; negative; zero; fractional; large; numeric text in a cell; date in a cell; number text with a point |
| NORMSINV | Compatibility | 10 | 5 | 4 | typical only | value instead of an error 5 | boolean; blank cell; negative; zero; above 1 |
| STDEVP | Compatibility | 12 | 11 | 8 | typical only | other error code 1 | blank range |
| TDIST | Compatibility | 15 | 12 | 7 | typical only, 3 more soft | number 3 | boolean; blank cell; zero |
| TINV | Compatibility | 15 | 6 | 1 | typical differ | number 6, error instead of a value 2, value instead of an error 1 | typical; numeric text; boolean; numeric text in a cell; date in a cell; above 1; number text with a point; fractional |
| TTEST | Compatibility | 15 | 5 | 2 | typical differ | value instead of an error 5, number 4, error instead of a value 1 | typical; range: mixed; blank range; different size; date in a cell; number text with a point |
| VARP | Compatibility | 12 | 11 | 8 | typical only | other error code 1 | blank range |
| ZTEST | Compatibility | 15 | 13 | 11 | typical only | error instead of a value 1, value instead of an error 1 | range: mixed; blank range |
| DAVERAGE | Database | 10 | 1 | 1 | typical differ | exception 5, error instead of a value 2, other error code 2 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field number 0; no such field |
| DCOUNT | Database | 10 | 1 | 1 | typical differ | exception 5, number 2, value instead of an error 2 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field number 0; no such field |
| DCOUNTA | Database | 10 | 0 | 0 | typical differ | exception 5, number 3, value instead of an error 2 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field number 0; no such field; field as header text |
| DGET | Database | 10 | 4 | 4 | typical differ | other error code 3, exception 2, error instead of a value 1 | typical; field, any case; one row; nothing matches; field as header text |
| DMAX | Database | 10 | 2 | 0 | typical differ | exception 6, error instead of a value 2 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field as header text |
| DMIN | Database | 10 | 2 | 0 | typical differ | exception 6, error instead of a value 2 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field as header text |
| DPRODUCT | Database | 10 | 0 | 0 | typical differ | exception 5, number 3, value instead of an error 2 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field number 0; no such field; field as header text |
| DSTDEV | Database | 10 | 0 | 0 | typical differ | exception 5, value instead of an error 4, number 1 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field number 0; no such field; field as header text |
| DSTDEVP | Database | 10 | 0 | 0 | typical differ | exception 5, other error code 3, error instead of a value 2 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field number 0; no such field; field as header text |
| DSUM | Database | 10 | 1 | 1 | typical differ | exception 5, number 2, value instead of an error 2 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field number 0; no such field |
| DVAR | Database | 10 | 0 | 0 | typical differ | exception 5, value instead of an error 4, number 1 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field number 0; no such field; field as header text |
| DVARP | Database | 10 | 0 | 0 | typical differ | exception 5, other error code 3, error instead of a value 2 | typical; field, any case; wildcard in the criterion; one row; nothing matches; empty criterion (all rows); field number 0; no such field; field as header text |
| DATE | Date and time | 15 | 13 | 2 | typical differ | number 2 | typical |
| DATEDIF | Date and time | 15 | 10 | 9 | typical only | other error code 4, value instead of an error 1 | ISO date text; m/d/yyyy date text; empty string; leap day 2024-02-29; spaces |
| DATEVALUE | Date and time | 13 | 12 | 8 | typical only | value instead of an error 1 | number |
| DAY | Date and time | 14 | 9 | 8 | typical only | number 3, error instead of a value 2 | ISO date text; m/d/yyyy date text; false leap day 1900-02-29; zero; blank cell |
| DAYS | Date and time | 15 | 12 | 10 | typical only | error instead of a value 2, number 1 | ISO date text; m/d/yyyy date text; 1900-03-01 |
| DAYS360 | Date and time | 15 | 9 | 10 | typical differ | number 4, error instead of a value 2 | typical; ISO date text; m/d/yyyy date text; 1900-02-28; zero |
| EDATE | Date and time | 15 | 12 | 1 | typical only | error instead of a value 2, number 1 | ISO date text; m/d/yyyy date text; false leap day 1900-02-29 |
| EOMONTH | Date and time | 15 | 11 | 1 | typical only | error instead of a value 2, number 2 | ISO date text; m/d/yyyy date text; false leap day 1900-02-29; zero |
| HOUR | Date and time | 12 | 11 | 8 | typical only | error instead of a value 1 | time text |
| ISOWEEKNUM | Date and time | 15 | 13 | 13 | typical only | error instead of a value 2 | ISO date text; m/d/yyyy date text |
| MINUTE | Date and time | 11 | 10 | 7 | typical only | error instead of a value 1 | time text |
| MONTH | Date and time | 15 | 10 | 9 | typical only | number 3, error instead of a value 2 | ISO date text; m/d/yyyy date text; false leap day 1900-02-29; zero; blank cell |
| NETWORKDAYS | Date and time | 15 | 5 | 6 | typical differ | number 6, error instead of a value 3, value instead of an error 1 | typical; ISO date text; m/d/yyyy date text; leap day 2024-02-29; false leap day 1900-02-29; boolean; 1900-02-28; 1900-03-01 |
| NETWORKDAYS.INTL | Date and time | 15 | 7 | 9 | typical differ | error instead of a value 4, number 4 | typical; ISO date text; m/d/yyyy date text; zero; leap day 2024-02-29; false leap day 1900-02-29; 1900-02-28 |
| SECOND | Date and time | 11 | 10 | 7 | typical only | error instead of a value 1 | time text |
| TIME | Date and time | 15 | 10 | 9 | typical differ | value instead of an error 3, number 1, error instead of a value 1 | typical; large; date in a cell; negative |
| TIMEVALUE | Date and time | 12 | 7 | 5 | typical differ | error instead of a value 2, other error code 2, value instead of an error 1 | typical; number; blank cell; empty string |
| WEEKDAY | Date and time | 15 | 9 | 11 | typical only | number 4, error instead of a value 2 | ISO date text; m/d/yyyy date text; false leap day 1900-02-29; 1900-02-28; zero; blank cell |
| WEEKNUM | Date and time | 15 | 11 | 12 | typical only | error instead of a value 2, number 2 | ISO date text; m/d/yyyy date text; zero; blank cell |
| WORKDAY | Date and time | 15 | 10 | 1 | typical differ | error instead of a value 3, number 2 | typical; ISO date text; m/d/yyyy date text; false leap day 1900-02-29; 1900-02-28 |
| WORKDAY.INTL | Date and time | 15 | 9 | 2 | typical differ | error instead of a value 3, number 2, other error code 1 | typical; ISO date text; m/d/yyyy date text; #DIV/0!; false leap day 1900-02-29; 1900-02-28 |
| YEAR | Date and time | 15 | 11 | 11 | typical only | error instead of a value 2, number 2 | ISO date text; m/d/yyyy date text; zero; blank cell |
| YEARFRAC | Date and time | 15 | 8 | 8 | typical only | number 5, error instead of a value 2 | ISO date text; m/d/yyyy date text; 1900-02-28; leap day 2024-02-29; false leap day 1900-02-29; 1900-03-01 |
| BESSELI | Engineering | 15 | 13 | 10 | typical only | value instead of an error 2 | boolean |
| BESSELJ | Engineering | 15 | 14 | 12 | typical only | value instead of an error 1 | boolean |
| BESSELK | Engineering | 15 | 13 | 9 | typical only | value instead of an error 2 | boolean |
| BESSELY | Engineering | 15 | 13 | 7 | typical only | value instead of an error 2 | boolean |
| BIN2DEC | Engineering | 12 | 9 | 8 | typical only | error instead of a value 2, other error code 1 | boolean; blank cell; empty string |
| BIN2HEX | Engineering | 15 | 7 | 5 | typical differ | string: case 4, other error code 2, error instead of a value 2 | typical; numeric text; boolean; blank cell; empty string |
| BIN2OCT | Engineering | 15 | 12 | 11 | typical only | error instead of a value 2, other error code 1 | boolean; blank cell; empty string |
| BITAND | Engineering | 15 | 13 | 13 | typical only | value instead of an error 2 | fractional |
| BITLSHIFT | Engineering | 15 | 13 | 13 | typical only | value instead of an error 1, number 1 | fractional; large |
| BITOR | Engineering | 15 | 12 | 12 | typical only | value instead of an error 2, number 1 | number text with a point; fractional; large |
| BITRSHIFT | Engineering | 15 | 13 | 13 | typical only | value instead of an error 1, number 1 | fractional; large |
| BITXOR | Engineering | 15 | 13 | 12 | typical only | value instead of an error 1, number 1 | fractional; large |
| COMPLEX | Engineering | 15 | 14 | 14 | typical only | value instead of an error 1 | boolean |
| CONVERT | Engineering | 15 | 13 | 11 | typical differ | error instead of a value 1, value instead of an error 1 | typical; boolean |
| DEC2BIN | Engineering | 15 | 12 | 10 | typical only | value instead of an error 1, other error code 1, exception 1 | boolean; large |
| DEC2HEX | Engineering | 15 | 9 | 8 | typical differ | string: case 3, value instead of an error 2, other error code 1 | typical; boolean; negative; date in a cell |
| DEC2OCT | Engineering | 15 | 13 | 11 | typical only | value instead of an error 1, exception 1 | boolean; large |
| DELTA | Engineering | 15 | 14 | 13 | typical only | value instead of an error 1 | boolean |
| ERF | Engineering | 14 | 12 | 11 | typical differ | number 1, value instead of an error 1 | typical; boolean |
| ERFC | Engineering | 13 | 12 | 8 | typical only | value instead of an error 1 | boolean |
| GESTEP | Engineering | 15 | 14 | 13 | typical only | value instead of an error 1 | boolean |
| HEX2BIN | Engineering | 15 | 11 | 9 | typical only | error instead of a value 2, exception 1, other error code 1 | large; boolean; blank cell; empty string |
| HEX2DEC | Engineering | 12 | 9 | 8 | typical only | error instead of a value 2, other error code 1 | boolean; blank cell; empty string |
| HEX2OCT | Engineering | 15 | 12 | 11 | typical only | error instead of a value 2, other error code 1 | boolean; blank cell; empty string |
| IMABS | Engineering | 10 | 5 | 3 | typical only | other error code 4, value instead of an error 1 | Cyrillic; case; empty string; spaces; Cyrillic in a cell |
| IMAGINARY | Engineering | 12 | 10 | 9 | typical only | other error code 1, value instead of an error 1 | boolean; empty string |
| IMARGUMENT | Engineering | 10 | 5 | 3 | typical only | other error code 5 | Cyrillic; case; blank cell; spaces; Cyrillic in a cell |
| IMCONJUGATE | Engineering | 10 | 4 | 2 | typical only | other error code 4, string 1, value instead of an error 1 | Cyrillic; case; blank cell; empty string; spaces; Cyrillic in a cell |
| IMCOS | Engineering | 10 | 4 | 2 | typical only | other error code 4, string: number format 1, value instead of an error 1 | Cyrillic; case; number; empty string; spaces; Cyrillic in a cell |
| IMCOSH | Engineering | 10 | 4 | 2 | typical only | other error code 4, string: number format 1, value instead of an error 1 | Cyrillic; case; number; empty string; spaces; Cyrillic in a cell |
| IMCOT | Engineering | 10 | 3 | 2 | typical only | other error code 4, exception 2, string: number format 1 | Cyrillic; case; number; blank cell; empty string; spaces; Cyrillic in a cell |
| IMCSC | Engineering | 10 | 6 | 7 | typical only | exception 2, string: number format 1, other error code 1 | number; boolean; blank cell; empty string |
| IMCSCH | Engineering | 10 | 6 | 7 | typical only | exception 2, string: number format 1, other error code 1 | number; boolean; blank cell; empty string |
| IMDIV | Engineering | 15 | 6 | 3 | typical only | other error code 7, other type 1, value instead of an error 1 | Cyrillic; case; blank cell; empty string; spaces; Cyrillic in a cell |
| IMEXP | Engineering | 10 | 4 | 2 | typical only | other error code 4, string: number format 1, value instead of an error 1 | Cyrillic; case; number; empty string; spaces; Cyrillic in a cell |
| IMLN | Engineering | 10 | 5 | 2 | typical only | other error code 4, string: number format 1 | Cyrillic; case; number; spaces; Cyrillic in a cell |
| IMLOG10 | Engineering | 10 | 5 | 2 | typical only | other error code 4, string: number format 1 | Cyrillic; case; number; spaces; Cyrillic in a cell |
| IMLOG2 | Engineering | 10 | 5 | 2 | typical only | other error code 4, string: number format 1 | Cyrillic; case; number; spaces; Cyrillic in a cell |
| IMPOWER | Engineering | 15 | 7 | 4 | typical only | other error code 5, value instead of an error 2, error instead of a value 1 | Cyrillic; case; blank cell; empty string; boolean; spaces; Cyrillic in a cell; date in a cell |
| IMPRODUCT | Engineering | 15 | 7 | 6 | typical only | other error code 6, other type 1, value instead of an error 1 | Cyrillic; case; blank cell; empty string; Cyrillic in a cell; spaces |
| IMREAL | Engineering | 10 | 6 | 6 | typical only | other type 2, other error code 1, value instead of an error 1 | number; boolean; blank cell; empty string |
| IMSEC | Engineering | 10 | 4 | 2 | typical only | other error code 4, string: number format 1, value instead of an error 1 | Cyrillic; case; number; empty string; spaces; Cyrillic in a cell |
| IMSECH | Engineering | 10 | 4 | 2 | typical only | other error code 4, string: number format 1, value instead of an error 1 | Cyrillic; case; number; empty string; spaces; Cyrillic in a cell |
| IMSIN | Engineering | 10 | 3 | 2 | typical only | other error code 4, string: number format 1, other type 1, value instead of an error 1 | Cyrillic; case; number; blank cell; empty string; spaces; Cyrillic in a cell |
| IMSINH | Engineering | 10 | 3 | 2 | typical only | other error code 4, string: number format 1, other type 1, value instead of an error 1 | Cyrillic; case; number; blank cell; empty string; spaces; Cyrillic in a cell |
| IMSQRT | Engineering | 10 | 3 | 2 | typical only | other error code 5, string: number format 1, error instead of a value 1 | Cyrillic; case; number; blank cell; empty string; spaces; Cyrillic in a cell |
| IMSUB | Engineering | 15 | 6 | 2 | typical only | other error code 7, value instead of an error 2 | Cyrillic; case; empty string; spaces; Cyrillic in a cell |
| IMSUM | Engineering | 15 | 8 | 5 | typical only | other error code 6, value instead of an error 1 | Cyrillic; case; empty string; spaces; Cyrillic in a cell |
| IMTAN | Engineering | 10 | 4 | 2 | typical only | other error code 4, exception 2 | Cyrillic; case; blank cell; empty string; spaces; Cyrillic in a cell |
| OCT2BIN | Engineering | 15 | 11 | 9 | typical only | other error code 2, error instead of a value 2 | boolean; blank cell; empty string |
| OCT2DEC | Engineering | 11 | 8 | 7 | typical only | error instead of a value 2, other error code 1 | boolean; blank cell; empty string |
| OCT2HEX | Engineering | 15 | 10 | 9 | typical differ | error instead of a value 2, string: case 1, other error code 1, value instead of an error 1 | typical; boolean; date in a cell; blank cell; empty string |
| ACCRINT | Financial | 15 | 11 | 6 | typical only | other error code 2, value instead of an error 1, number 1 | ISO date text; m/d/yyyy date text; boolean; false leap day 1900-02-29 |
| COUPDAYS | Financial | 15 | 10 | 10 | typical only | number 3, other error code 2 | ISO date text; m/d/yyyy date text; false leap day 1900-02-29; 1900-02-28; 1900-03-01 |
| CUMIPMT | Financial | 15 | 12 | 8 | typical only | value instead of an error 2, number 1 | numeric text; boolean; numeric text in a cell |
| CUMPRINC | Financial | 15 | 13 | 9 | typical only | number 1, value instead of an error 1 | numeric text; boolean |
| DB | Financial | 15 | 12 | 10 | typical differ | number 2, error instead of a value 1 | typical; numeric text; boolean |
| DDB | Financial | 15 | 14 | 13 | typical only | value instead of an error 1 | blank cell |
| DISC | Financial | 15 | 5 | 5 | typical differ, 4 more soft | number 7, error instead of a value 2, value instead of an error 1 | typical; ISO date text; m/d/yyyy date text; boolean; numeric text; leap day 2024-02-29; false leap day 1900-02-29; number text with a point; 1900-02-28 |
| DOLLARDE | Financial | 15 | 12 | 11 | typical only | number 2, value instead of an error 1 | boolean; numeric text; number text with a point |
| DOLLARFR | Financial | 15 | 14 | 13 | typical only | value instead of an error 1 | boolean |
| EFFECT | Financial | 15 | 13 | 11 | typical only | value instead of an error 2 | boolean |
| FV | Financial | 15 | 14 | 11 | typical only | number 1 | fractional |
| FVSCHEDULE | Financial | 13 | 12 | 11 | typical only | value instead of an error 1 | boolean |
| IPMT | Financial | 15 | 12 | 12 | typical only | value instead of an error 2, number 1 | fractional; blank cell; negative |
| IRR | Financial | 7 | 4 | 5 | typical only | other error code 2, number 1 | range: mixed; range with an error; blank range |
| ISPMT | Financial | 15 | 15 | 13 | all |  |  |
| MIRR | Financial | 15 | 13 | 11 | typical only | number 1, other error code 1 | range: mixed; no sign change |
| NOMINAL | Financial | 15 | 14 | 13 | typical only | value instead of an error 1 | boolean |
| NPER | Financial | 15 | 13 | 11 | typical only | number 2 | numeric text; numeric text in a cell |
| NPV | Financial | 15 | 15 | 14 | all |  |  |
| PDURATION | Financial | 15 | 15 | 12 | all |  |  |
| PMT | Financial | 15 | 14 | 12 | typical only | value instead of an error 1 | negative |
| PPMT | Financial | 15 | 12 | 8 | typical only | value instead of an error 3 | zero; blank cell; negative |
| PRICEDISC | Financial | 15 | 10 | 8 | typical only | other error code 3, error instead of a value 1, number 1 | ISO date text; m/d/yyyy date text; leap day 2024-02-29; 1900-03-01 |
| PV | Financial | 15 | 13 | 11 | typical only | number 2 | fractional; negative |
| RATE | Financial | 15 | 14 | 13 | typical only | value instead of an error 1 | negative |
| RRI | Financial | 15 | 14 | 12 | typical only | value instead of an error 1 | negative |
| SLN | Financial | 15 | 14 | 13 | typical only | other error code 1 | blank cell |
| SYD | Financial | 15 | 14 | 13 | typical only | value instead of an error 1 | negative |
| TBILLEQ | Financial | 15 | 11 | 12 | typical differ | other error code 2, number 1, value instead of an error 1 | typical; ISO date text; m/d/yyyy date text; large |
| TBILLPRICE | Financial | 15 | 10 | 9 | typical differ | other error code 3, number 2 | typical; ISO date text; m/d/yyyy date text; numeric text |
| TBILLYIELD | Financial | 15 | 10 | 10 | typical differ | number 3, other error code 2 | typical; ISO date text; m/d/yyyy date text; number text with a point; large |
| XIRR | Financial | 8 | 2 | 1 | typical differ, 4 more soft | number 4, value instead of an error 2 | typical; dates out of order; different size; guess -0.5; cash flow: mixed |
| XNPV | Financial | 15 | 9 | 7 | typical only | value instead of an error 6 | dates out of order; different size; cash flow: mixed; boolean; blank cell; zero |
| ERROR.TYPE | Information | 9 | 9 | 9 | all |  |  |
| ISBLANK | Information | 9 | 9 | 9 | all |  |  |
| ISERR | Information | 9 | 9 | 9 | all |  |  |
| ISERROR | Information | 9 | 9 | 9 | all |  |  |
| ISEVEN | Information | 15 | 14 | 12 | typical only | value instead of an error 1 | boolean |
| ISLOGICAL | Information | 9 | 9 | 9 | all |  |  |
| ISNA | Information | 9 | 9 | 9 | all |  |  |
| ISNONTEXT | Information | 9 | 9 | 9 | all |  |  |
| ISNUMBER | Information | 9 | 9 | 9 | all |  |  |
| ISODD | Information | 15 | 14 | 12 | typical only | value instead of an error 1 | boolean |
| ISTEXT | Information | 9 | 9 | 9 | all |  |  |
| N | Information | 11 | 11 | 11 | all |  |  |
| NA | Information | 1 | 1 | 1 | all |  |  |
| TYPE | Information | 11 | 10 | 10 | typical only | error instead of a value 1 | blank cell |
| AND | Logical | 15 | 15 | 15 | all |  |  |
| FALSE | Logical | 1 | 1 | 1 | all |  |  |
| IF | Logical | 15 | 15 | 14 | all |  |  |
| IFERROR | Logical | 15 | 12 | 12 | typical differ | error instead of a value 3 | typical; blank cell |
| IFNA | Logical | 15 | 14 | 14 | typical only | error instead of a value 1 | blank cell |
| IFS | Logical | 15 | 15 | 13 | all |  |  |
| NOT | Logical | 9 | 9 | 8 | all |  |  |
| OR | Logical | 13 | 13 | 13 | all |  |  |
| SWITCH | Logical | 15 | 12 | 12 | typical only | error instead of a value 1, number 1, value instead of an error 1 | case; Cyrillic and case; #DIV/0! |
| TRUE | Logical | 1 | 1 | 1 | all |  |  |
| XOR | Logical | 13 | 13 | 13 | all |  |  |
| CHOOSE | Lookup and reference | 15 | 15 | 10 | all |  |  |
| CHOOSECOLS | Lookup and reference | 19 | 19 | 16 | all |  |  |
| CHOOSEROWS | Lookup and reference | 19 | 19 | 16 | all |  |  |
| COLUMN | Lookup and reference | 4 | 0 | 0 | typical differ | error instead of a value 2, other error code 2 | typical; one cell; table by rows |
| COLUMNS | Lookup and reference | 5 | 5 | 4 | all |  |  |
| DROP | Lookup and reference | 16 | 16 | 15 | all |  |  |
| EXPAND | Lookup and reference | 19 | 19 | 18 | all |  |  |
| HLOOKUP | Lookup and reference | 15 | 13 | 11 | typical only | other error code 2 | blank cell; zero |
| HSTACK | Lookup and reference | 7 | 7 | 7 | all |  |  |
| INDEX | Lookup and reference | 15 | 12 | 9 | typical only | value instead of an error 3 | blank cell; zero |
| LOOKUP | Lookup and reference | 14 | 12 | 12 | typical differ | number 1, other error code 1 | typical; #DIV/0! |
| MATCH | Lookup and reference | 15 | 12 | 11 | typical only | other error code 2, number 1 | descending -1; error input; #DIV/0! |
| ROW | Lookup and reference | 4 | 0 | 0 | typical differ | other error code 3, error instead of a value 1 | typical; one cell; table by rows |
| ROWS | Lookup and reference | 5 | 5 | 4 | all |  |  |
| SORT | Lookup and reference | 16 | 10 | 10 | typical differ | number 5, string 1 | typical, element 2,1; typical, element 7,1; typical |
| TAKE | Lookup and reference | 12 | 12 | 12 | all |  |  |
| TRANSPOSE | Lookup and reference | 8 | 8 | 8 | all |  |  |
| UNIQUE | Lookup and reference | 10 | 1 | 1 | typical differ | error instead of a value 5, number 4 | typical; typical, element 4,1; typical, row count; case: a and A one value; by columns; Cyrillic and case; blank cells |
| VLOOKUP | Lookup and reference | 15 | 11 | 12 | typical only | error instead of a value 2, other error code 2 | wildcard *; wildcard ?; zero; error input |
| VSTACK | Lookup and reference | 7 | 7 | 7 | all |  |  |
| ABS | Math and trigonometry | 14 | 14 | 14 | all |  |  |
| ACOS | Math and trigonometry | 14 | 14 | 14 | all |  |  |
| ACOSH | Math and trigonometry | 14 | 14 | 14 | all |  |  |
| ACOT | Math and trigonometry | 13 | 12 | 12 | typical only | number 1 | negative |
| ACOTH | Math and trigonometry | 13 | 13 | 12 | all |  |  |
| AGGREGATE | Math and trigonometry | 15 | 13 | 7 | typical differ | error instead of a value 1, value instead of an error 1 | typical; date in a cell |
| ARABIC | Math and trigonometry | 11 | 10 | 10 | typical differ | error instead of a value 1 | typical |
| ASIN | Math and trigonometry | 13 | 13 | 13 | all |  |  |
| ASINH | Math and trigonometry | 13 | 13 | 13 | all |  |  |
| ATAN | Math and trigonometry | 13 | 13 | 13 | all |  |  |
| ATAN2 | Math and trigonometry | 15 | 6 | 6 | typical only | number 9 | numeric text; fractional; blank cell; negative; numeric text in a cell; zero; large |
| ATANH | Math and trigonometry | 14 | 14 | 13 | all |  |  |
| BASE | Math and trigonometry | 15 | 13 | 12 | typical only | value instead of an error 1, exception 1 | negative; boolean |
| CEILING.MATH | Math and trigonometry | 15 | 15 | 15 | all |  |  |
| CEILING.PRECISE | Math and trigonometry | 15 | 15 | 15 | all |  |  |
| COMBIN | Math and trigonometry | 15 | 13 | 12 | typical only | exception 2 | negative; large |
| COMBINA | Math and trigonometry | 15 | 12 | 12 | typical only | error instead of a value 2, exception 1 | boolean; fractional; large |
| COS | Math and trigonometry | 13 | 12 | 12 | typical only | value instead of an error 1 | large |
| COSH | Math and trigonometry | 13 | 13 | 11 | all |  |  |
| COT | Math and trigonometry | 13 | 12 | 12 | typical only | value instead of an error 1 | large |
| COTH | Math and trigonometry | 13 | 11 | 11 | typical only | error instead of a value 2 | large; date in a cell |
| CSC | Math and trigonometry | 13 | 12 | 12 | typical only | value instead of an error 1 | large |
| CSCH | Math and trigonometry | 13 | 13 | 13 | all |  |  |
| DECIMAL | Math and trigonometry | 15 | 13 | 13 | typical only | value instead of an error 2 | number; spaces |
| DEGREES | Math and trigonometry | 13 | 13 | 13 | all |  |  |
| EVEN | Math and trigonometry | 15 | 15 | 15 | all |  |  |
| EXP | Math and trigonometry | 14 | 14 | 12 | all |  |  |
| FACT | Math and trigonometry | 14 | 11 | 11 | typical only | exception 3 | negative; large; date in a cell |
| FACTDOUBLE | Math and trigonometry | 14 | 10 | 10 | typical only | value instead of an error 2, exception 2 | boolean; negative; large; date in a cell |
| FLOOR.MATH | Math and trigonometry | 15 | 15 | 15 | all |  |  |
| FLOOR.PRECISE | Math and trigonometry | 15 | 15 | 15 | all |  |  |
| GCD | Math and trigonometry | 13 | 8 | 10 | typical only | value instead of an error 4, error instead of a value 1 | range: mixed; blank range; boolean argument; reference to a boolean; negative and fractional |
| INT | Math and trigonometry | 14 | 14 | 14 | all |  |  |
| ISO.CEILING | Math and trigonometry | 15 | 5 | 5 | typical differ | number 9, error instead of a value 1 | typical; numeric text; boolean; negative; fractional; large; numeric text in a cell; date in a cell; number text with a point |
| LCM | Math and trigonometry | 12 | 6 | 8 | typical only | value instead of an error 4, error instead of a value 1, number 1 | range: mixed; blank range; boolean argument; reference to numeric text; reference to a boolean; negative and fractional |
| LN | Math and trigonometry | 14 | 14 | 13 | all |  |  |
| LOG | Math and trigonometry | 15 | 15 | 14 | all |  |  |
| LOG10 | Math and trigonometry | 14 | 14 | 13 | all |  |  |
| MMULT | Math and trigonometry | 8 | 8 | 8 | all |  |  |
| MOD | Math and trigonometry | 15 | 14 | 14 | typical differ | number 1 | typical |
| MROUND | Math and trigonometry | 15 | 14 | 14 | typical only | value instead of an error 1 | boolean |
| MULTINOMIAL | Math and trigonometry | 11 | 5 | 8 | typical only | value instead of an error 3, error instead of a value 1, number 1, exception 1 | range: mixed; blank range; boolean argument; reference to numeric text; reference to a boolean; negative and fractional |
| MUNIT | Math and trigonometry | 7 | 7 | 7 | all |  |  |
| ODD | Math and trigonometry | 15 | 15 | 15 | all |  |  |
| PI | Math and trigonometry | 1 | 1 | 1 | all |  |  |
| POWER | Math and trigonometry | 15 | 15 | 15 | all |  |  |
| PRODUCT | Math and trigonometry | 12 | 12 | 10 | all |  |  |
| QUOTIENT | Math and trigonometry | 15 | 12 | 12 | typical only | value instead of an error 1, other error code 1, number 1 | boolean; zero; large |
| RADIANS | Math and trigonometry | 13 | 13 | 13 | all |  |  |
| ROMAN | Math and trigonometry | 15 | 10 | 9 | typical differ | string 3, value instead of an error 2 | typical; negative; fractional |
| ROUND | Math and trigonometry | 15 | 14 | 14 | typical only | error instead of a value 1 | date in a cell |
| ROUNDDOWN | Math and trigonometry | 15 | 15 | 14 | all |  |  |
| ROUNDUP | Math and trigonometry | 15 | 15 | 14 | all |  |  |
| SEC | Math and trigonometry | 13 | 12 | 12 | typical only | value instead of an error 1 | large |
| SECH | Math and trigonometry | 13 | 13 | 13 | all |  |  |
| SERIESSUM | Math and trigonometry | 15 | 12 | 10 | typical only | value instead of an error 3 | boolean; blank cell; zero |
| SIGN | Math and trigonometry | 14 | 14 | 14 | all |  |  |
| SIN | Math and trigonometry | 14 | 13 | 13 | typical only | value instead of an error 1 | large |
| SINH | Math and trigonometry | 14 | 14 | 12 | all |  |  |
| SQRT | Math and trigonometry | 14 | 14 | 14 | all |  |  |
| SQRTPI | Math and trigonometry | 14 | 13 | 12 | typical only | value instead of an error 1 | boolean |
| SUBTOTAL | Math and trigonometry | 15 | 15 | 9 | all |  |  |
| SUM | Math and trigonometry | 13 | 13 | 9 | all |  |  |
| SUMIF | Math and trigonometry | 15 | 6 | 6 | typical differ | number 9 | typical; wildcard *; escaped ~*; wildcard ?; Cyrillic, any case; date text; not equal; empty criterion; boolean criterion |
| SUMIFS | Math and trigonometry | 16 | 6 | 6 | typical differ | number 10 | typical; wildcard *; escaped ~*; wildcard ?; Cyrillic, any case; date text; not equal; empty criterion; boolean criterion; two criteria |
| SUMPRODUCT | Math and trigonometry | 10 | 7 | 6 | typical only | error instead of a value 2, value instead of an error 1 | range: mixed; different size |
| SUMSQ | Math and trigonometry | 12 | 11 | 9 | typical only | error instead of a value 1 | blank range |
| SUMX2MY2 | Math and trigonometry | 10 | 4 | 2 | typical only | value instead of an error 3, error instead of a value 2, other error code 1 | range: mixed; blank range; different size |
| SUMX2PY2 | Math and trigonometry | 10 | 4 | 2 | typical only | value instead of an error 3, error instead of a value 2, other error code 1 | range: mixed; blank range; different size |
| SUMXMY2 | Math and trigonometry | 10 | 4 | 2 | typical only | value instead of an error 3, error instead of a value 2, other error code 1 | range: mixed; different size; blank range |
| TAN | Math and trigonometry | 14 | 13 | 13 | typical only | value instead of an error 1 | large |
| TANH | Math and trigonometry | 14 | 12 | 12 | typical only | error instead of a value 2 | large; date in a cell |
| TRUNC | Math and trigonometry | 15 | 15 | 15 | all |  |  |
| AVEDEV | Statistical | 12 | 12 | 10 | all |  |  |
| AVERAGE | Statistical | 12 | 12 | 10 | all |  |  |
| AVERAGEA | Statistical | 12 | 10 | 10 | typical only | number 1, value instead of an error 1 | numeric text argument; text argument |
| AVERAGEIF | Statistical | 15 | 5 | 5 | typical differ | error instead of a value 6, number 3, other error code 1 | typical; wildcard *; escaped ~*; wildcard ?; Cyrillic, any case; date text; not equal; empty criterion; boolean criterion; nothing matches |
| AVERAGEIFS | Statistical | 16 | 5 | 5 | typical differ | number 10, value instead of an error 1 | typical; wildcard *; escaped ~*; wildcard ?; Cyrillic, any case; date text; not equal; empty criterion; boolean criterion; nothing matches; two criteria |
| BETA.DIST | Statistical | 15 | 13 | 10 | typical differ | number 1, value instead of an error 1 | typical; blank cell |
| BETA.INV | Statistical | 15 | 11 | 8 | typical only | value instead of an error 4 | boolean; blank cell; date in a cell |
| BINOM.DIST | Statistical | 15 | 14 | 9 | typical only | value instead of an error 1 | large |
| BINOM.DIST.RANGE | Statistical | 15 | 15 | 7 | all |  |  |
| BINOM.INV | Statistical | 15 | 10 | 8 | typical only | value instead of an error 4, other error code 1 | zero; boolean; negative |
| CHISQ.DIST | Statistical | 15 | 14 | 11 | typical only | value instead of an error 1 | negative |
| CHISQ.DIST.RT | Statistical | 15 | 10 | 6 | typical only | error instead of a value 3, other error code 1, number 1 | blank cell; zero; fractional; number text with a point |
| CHISQ.INV | Statistical | 15 | 10 | 9 | hangs | value instead of an error 4, hangs 1 | boolean; negative; fractional; above 1; large |
| CHISQ.INV.RT | Statistical | 15 | 11 | 5 | hangs | other error code 3, hangs 1 | blank cell; zero; large |
| CHISQ.TEST | Statistical | 4 | 1 | 0 | typical differ, 1 more soft | number 2, other error code 1 | typical; range: mixed; different size |
| CONFIDENCE.NORM | Statistical | 15 | 10 | 5 | typical only | value instead of an error 5 | boolean; blank cell; negative; zero; above 1 |
| CONFIDENCE.T | Statistical | 15 | 10 | 4 | typical differ, 3 more soft | number 3, value instead of an error 2 | typical; boolean; numeric text; large; above 1 |
| CORREL | Statistical | 9 | 3 | 1 | typical only | other error code 3, error instead of a value 2, value instead of an error 1 | range: mixed; blank range; different size |
| COUNT | Statistical | 12 | 10 | 10 | typical only | number 2 | numeric text argument; boolean argument |
| COUNTA | Statistical | 12 | 12 | 12 | all |  |  |
| COUNTBLANK | Statistical | 7 | 7 | 7 | all |  |  |
| COUNTIF | Statistical | 15 | 6 | 6 | typical differ | number 9 | typical; wildcard *; escaped ~*; wildcard ?; Cyrillic, any case; date text; not equal; empty criterion; boolean criterion |
| COUNTIFS | Statistical | 16 | 6 | 6 | typical differ | number 10 | typical; wildcard *; escaped ~*; wildcard ?; Cyrillic, any case; date text; not equal; empty criterion; boolean criterion; two criteria |
| COVARIANCE.P | Statistical | 9 | 3 | 1 | typical only | value instead of an error 3, error instead of a value 2, other error code 1 | range: mixed; blank range; different size |
| COVARIANCE.S | Statistical | 9 | 3 | 1 | typical only | value instead of an error 3, error instead of a value 2, other error code 1 | range: mixed; blank range; different size |
| DEVSQ | Statistical | 12 | 11 | 8 | typical only | other error code 1 | blank range |
| EXPON.DIST | Statistical | 15 | 13 | 10 | typical only | value instead of an error 2 | negative; zero |
| F.DIST | Statistical | 15 | 12 | 9 | typical only, 1 more soft | number 2, value instead of an error 1 | large; negative; number text with a point |
| F.DIST.RT | Statistical | 15 | 15 | 9 | all |  |  |
| F.INV | Statistical | 15 | 12 | 9 | typical only, 1 more soft | error instead of a value 2, number 1 | blank cell; large; zero |
| F.INV.RT | Statistical | 15 | 14 | 8 | typical only, 1 more soft | number 1 | large |
| F.TEST | Statistical | 9 | 2 | 0 | typical differ | number 3, error instead of a value 2, value instead of an error 1, other error code 1 | typical; range: mixed; different size; blank range |
| FISHER | Statistical | 12 | 12 | 5 | all |  |  |
| FISHERINV | Statistical | 13 | 11 | 11 | typical only | error instead of a value 2 | large; date in a cell |
| FORECAST | Statistical | 15 | 11 | 8 | typical only | value instead of an error 2, error instead of a value 1, other error code 1 | range: mixed; different size; blank range |
| FREQUENCY | Statistical | 8 | 6 | 5 | typical only | error instead of a value 1, number 1 | range: mixed; blank range |
| GAMMA | Statistical | 15 | 10 | 8 | typical differ, 5 more soft | number 5 | typical; negative; fractional; number text with a point |
| GAMMA.DIST | Statistical | 15 | 12 | 6 | typical only | other error code 3 | blank cell; negative |
| GAMMA.INV | Statistical | 15 | 14 | 8 | typical only | value instead of an error 1 | boolean |
| GAMMALN | Statistical | 13 | 13 | 10 | all |  |  |
| GAMMALN.PRECISE | Statistical | 13 | 13 | 8 | all |  |  |
| GAUSS | Statistical | 13 | 13 | 13 | all |  |  |
| GEOMEAN | Statistical | 12 | 11 | 7 | typical only | other error code 1 | blank range |
| GROWTH | Statistical | 10 | 8 | 8 | typical only | other error code 2 | different size; zero in known_y |
| HARMEAN | Statistical | 12 | 10 | 7 | typical only | other error code 1, value instead of an error 1 | blank range; negative and fractional |
| HYPGEOM.DIST | Statistical | 15 | 13 | 7 | typical only | exception 2 | blank cell; date in a cell |
| INTERCEPT | Statistical | 9 | 5 | 3 | typical only | error instead of a value 2, other error code 2 | range: mixed; blank range |
| KURT | Statistical | 12 | 4 | 4 | typical only | other error code 8 | range: mixed; blank range; numeric text argument; boolean argument; reference to numeric text; reference to a boolean; blank cell argument; negative and fractional |
| LARGE | Statistical | 15 | 8 | 6 | typical only | other error code 5, number 2 | negative; blank range; number text with a point; blank cell; fractional; large; date in a cell |
| LINEST | Statistical | 10 | 3 | 3 | typical differ | number 4, error instead of a value 3 | typical, element 1,2; typical, row count; typical, column count; typical, element 3,1 |
| LOGEST | Statistical | 10 | 1 | 1 | typical differ, 2 more soft | number 6, error instead of a value 3 | typical; typical, element 1,2; typical, row count; typical, column count; typical, element 3,1 |
| LOGNORM.DIST | Statistical | 15 | 13 | 10 | typical only | value instead of an error 2 | blank cell; negative |
| LOGNORM.INV | Statistical | 15 | 9 | 8 | typical only | value instead of an error 6 | boolean; blank cell; negative; zero; above 1 |
| MAX | Statistical | 12 | 12 | 10 | all |  |  |
| MAXA | Statistical | 12 | 10 | 10 | typical only | value instead of an error 1, number 1 | text argument; reference to numeric text |
| MAXIFS | Statistical | 16 | 8 | 8 | typical differ | number 8 | typical; wildcard *; escaped ~*; wildcard ?; date text; empty criterion; boolean criterion; two criteria |
| MEDIAN | Statistical | 12 | 12 | 7 | all |  |  |
| MIN | Statistical | 12 | 12 | 11 | all |  |  |
| MINA | Statistical | 12 | 9 | 9 | typical only | number 2, value instead of an error 1 | text argument; reference to numeric text; blank cell argument |
| MINIFS | Statistical | 16 | 7 | 7 | typical only | number 9 | wildcard *; escaped ~*; wildcard ?; Cyrillic, any case; date text; not equal; empty criterion; boolean criterion; two criteria |
| MODE.MULT | Statistical | 14 | 2 | 2 | typical differ | value instead of an error 7, number 3, other error code 2 | typical; typical, element 2,1; typical, element 3,1; range: mixed; blank range; numeric text argument; boolean argument; text argument; reference to numeric text; reference to a boolean; blank cell argument; negative and fractional |
| MODE.SNGL | Statistical | 12 | 3 | 3 | typical only | value instead of an error 7, other error code 2 | range: mixed; blank range; numeric text argument; boolean argument; text argument; reference to numeric text; reference to a boolean; blank cell argument; negative and fractional |
| NEGBINOM.DIST | Statistical | 15 | 14 | 13 | typical only | value instead of an error 1 | negative |
| NORM.DIST | Statistical | 15 | 15 | 12 | all |  |  |
| NORM.S.DIST | Statistical | 15 | 15 | 13 | all |  |  |
| NORM.S.INV | Statistical | 10 | 5 | 4 | typical only | value instead of an error 5 | boolean; blank cell; negative; zero; above 1 |
| NORMINV | Statistical | 15 | 10 | 9 | typical only | value instead of an error 5 | boolean; blank cell; negative; zero; above 1 |
| PEARSON | Statistical | 9 | 3 | 1 | typical only | other error code 3, error instead of a value 2, value instead of an error 1 | range: mixed; blank range; different size |
| PERCENTILE.EXC | Statistical | 15 | 14 | 11 | typical only | other error code 1 | blank range |
| PERCENTILE.INC | Statistical | 15 | 14 | 9 | typical only | other error code 1 | blank range |
| PERCENTRANK.EXC | Statistical | 15 | 13 | 11 | typical only | other error code 2 | range: mixed; blank range |
| PERCENTRANK.INC | Statistical | 15 | 10 | 8 | typical differ | other error code 5 | typical; range: mixed; blank range; date in a cell; large |
| PERMUT | Statistical | 15 | 8 | 6 | typical only | exception 7 | boolean; blank cell; large; negative; zero; fractional |
| PERMUTATIONA | Statistical | 15 | 13 | 10 | typical only | value instead of an error 2 | negative; large |
| PHI | Statistical | 12 | 12 | 11 | all |  |  |
| POISSON.DIST | Statistical | 15 | 14 | 13 | typical only | value instead of an error 1 | negative |
| PROB | Statistical | 14 | 14 | 13 | all |  |  |
| QUARTILE.EXC | Statistical | 15 | 14 | 12 | typical only | other error code 1 | blank range |
| QUARTILE.INC | Statistical | 15 | 12 | 10 | typical differ | error instead of a value 2, other error code 1 | typical; blank range |
| RANK.AVG | Statistical | 15 | 5 | 4 | typical only | value instead of an error 10 | numeric text; boolean; blank cell; negative; zero; fractional; large; numeric text in a cell; date in a cell; number text with a point |
| RANK.EQ | Statistical | 15 | 5 | 4 | typical only | value instead of an error 10 | numeric text; boolean; blank cell; negative; zero; fractional; large; numeric text in a cell; date in a cell; number text with a point |
| RSQ | Statistical | 9 | 3 | 1 | typical only | other error code 3, error instead of a value 2, value instead of an error 1 | range: mixed; blank range; different size |
| SKEW | Statistical | 12 | 6 | 5 | typical only | other error code 6 | blank range; numeric text argument; boolean argument; reference to numeric text; reference to a boolean; blank cell argument |
| SKEW.P | Statistical | 12 | 6 | 5 | typical only | other error code 5, value instead of an error 1 | blank range; numeric text argument; boolean argument; reference to numeric text; reference to a boolean; blank cell argument |
| SLOPE | Statistical | 9 | 3 | 1 | typical only | error instead of a value 2, value instead of an error 2, other error code 2 | range: mixed; blank range; different size |
| SMALL | Statistical | 15 | 11 | 7 | typical only | other error code 4 | range: mixed; blank range; zero; blank cell |
| STANDARDIZE | Statistical | 15 | 14 | 12 | typical only | value instead of an error 1 | negative |
| STDEV.P | Statistical | 12 | 11 | 8 | typical only | other error code 1 | blank range |
| STDEV.S | Statistical | 12 | 8 | 4 | typical only | other error code 3, value instead of an error 1 | blank range; reference to numeric text; reference to a boolean; blank cell argument |
| STDEVA | Statistical | 12 | 7 | 7 | typical only | other error code 3, number 1, value instead of an error 1 | range: mixed; range with an error; blank range; text argument; blank cell argument |
| STDEVPA | Statistical | 12 | 7 | 7 | typical only | number 2, other error code 2, value instead of an error 1 | range: mixed; range with an error; blank range; text argument; blank cell argument |
| STEYX | Statistical | 9 | 3 | 1 | typical only | error instead of a value 2, value instead of an error 2, other error code 2 | range: mixed; different size; blank range |
| T.DIST | Statistical | 15 | 4 | 1 | typical differ | error instead of a value 11 | typical; numeric text; boolean; number text with a point; blank cell; negative; boolean in a cell; zero |
| T.DIST.2T | Statistical | 15 | 9 | 6 | typical only, 5 more soft | number 6 | boolean; blank cell; number text with a point; large; zero; fractional |
| T.DIST.RT | Statistical | 15 | 8 | 3 | typical only, 6 more soft | number 6, error instead of a value 1 | boolean; blank cell; date in a cell; negative; large; zero; fractional |
| T.INV | Statistical | 15 | 13 | 4 | typical only, 1 more soft | number 2 | boolean; number text with a point |
| T.INV.2T | Statistical | 15 | 10 | 9 | typical only, 3 more soft | number 4, error instead of a value 1 | numeric text; number text with a point; above 1 |
| T.TEST | Statistical | 15 | 4 | 0 | typical differ | value instead of an error 5, number 4, error instead of a value 2 | typical; range: mixed; numeric text in a cell; blank range; different size; date in a cell |
| TREND | Statistical | 10 | 3 | 3 | typical differ | number 3, error instead of a value 2, other error code 1, value instead of an error 1 | typical; typical, element 2,1; typical, row count; without known_x; without the constant; different size; text in known_y |
| TRIMMEAN | Statistical | 15 | 8 | 4 | typical differ | number 5, other error code 1, value instead of an error 1 | typical; range: mixed; blank range; different size; blank cell; zero; negative |
| VAR.P | Statistical | 12 | 11 | 8 | typical only | other error code 1 | blank range |
| VAR.S | Statistical | 12 | 8 | 4 | typical only | other error code 3, value instead of an error 1 | blank range; reference to numeric text; reference to a boolean; blank cell argument |
| VARA | Statistical | 12 | 7 | 7 | typical only | other error code 3, number 1, value instead of an error 1 | range: mixed; range with an error; blank range; text argument; blank cell argument |
| VARPA | Statistical | 12 | 7 | 7 | typical only | number 2, other error code 2, value instead of an error 1 | range: mixed; range with an error; blank range; text argument; blank cell argument |
| WEIBULL.DIST | Statistical | 15 | 13 | 11 | typical only | value instead of an error 2 | negative; large |
| Z.TEST | Statistical | 15 | 13 | 11 | typical only | error instead of a value 1, value instead of an error 1 | range: mixed; blank range |
| CHAR | Text | 15 | 10 | 10 | typical differ | value instead of an error 4, string 1 | typical; negative; large; date in a cell |
| CLEAN | Text | 11 | 11 | 9 | all |  |  |
| CODE | Text | 12 | 8 | 6 | typical differ | number 4 | typical; Cyrillic; case; Cyrillic in a cell |
| CONCAT | Text | 12 | 12 | 12 | all |  |  |
| CONCATENATE | Text | 15 | 15 | 15 | all |  |  |
| DOLLAR | Text | 15 | 12 | 10 | typical differ | string 3 | typical; negative |
| EXACT | Text | 15 | 15 | 15 | all |  |  |
| FIND | Text | 15 | 15 | 14 | all |  |  |
| FIXED | Text | 15 | 15 | 14 | all |  |  |
| LEFT | Text | 15 | 15 | 14 | all |  |  |
| LEN | Text | 12 | 12 | 12 | all |  |  |
| LOWER | Text | 11 | 11 | 11 | all |  |  |
| MID | Text | 15 | 13 | 12 | typical only | error instead of a value 1, value instead of an error 1 | blank cell; large |
| NUMBERVALUE | Text | 15 | 6 | 5 | typical differ | other error code 7, error instead of a value 2 | typical; Cyrillic; case; boolean; number; blank cell; Cyrillic in a cell |
| PROPER | Text | 13 | 9 | 9 | typical differ | string 4 | typical; case |
| REPLACE | Text | 15 | 14 | 9 | typical only | value instead of an error 1 | negative |
| REPT | Text | 15 | 15 | 14 | all |  |  |
| RIGHT | Text | 15 | 15 | 14 | all |  |  |
| SEARCH | Text | 15 | 12 | 11 | typical only | error instead of a value 3 | wildcard ?; wildcard *; escaped ~? |
| SUBSTITUTE | Text | 15 | 14 | 10 | typical differ | string 1 | typical |
| T | Text | 11 | 11 | 11 | all |  |  |
| TEXT | Text | 15 | 8 | 9 | typical differ | string 5, exception 2 | typical; text; boolean; blank cell |
| TEXTJOIN | Text | 15 | 12 | 10 | typical only | string 2, value instead of an error 1 | range: mixed; range with an error |
| TRIM | Text | 12 | 12 | 11 | all |  |  |
| UNICHAR | Text | 15 | 13 | 13 | typical only | value instead of an error 2 | negative; large |
| UNICODE | Text | 12 | 12 | 10 | all |  |  |
| UPPER | Text | 12 | 11 | 11 | typical differ | string 1 | typical |
| VALUE | Text | 15 | 10 | 11 | typical differ | value instead of an error 3, error instead of a value 2 | typical; boolean; empty string |
<!-- /auto:functions -->

## Where edge cases differ

For the 206 functions whose typical cases match: the edge cases they differ on, and for how many functions.

<!-- auto:tags -->
| edge case | functions | which |
|---|---|---|
| blank cell | 61 | ATAN2, BETA.INV, BETADIST, BETAINV, BIN2DEC, BIN2OCT, CHISQ.DIST.RT, CONFIDENCE.NORM, CRITBINOM, DAY, DDB, F.INV, GAMMA.DIST, HEX2BIN, HEX2DEC, HEX2OCT, HLOOKUP, HYPGEOM.DIST, IFNA, IMARGUMENT, IMCONJUGATE, IMCOT, IMCSC, IMCSCH, IMDIV, IMPOWER, IMPRODUCT, IMREAL, IMSIN, IMSINH, IMSQRT, IMTAN, INDEX, IPMT, LARGE, LOGINV, LOGNORM.DIST, LOGNORM.INV, MID, MONTH, NORM.INV, NORM.S.INV, NORMINV, NORMSINV, OCT2BIN, OCT2DEC, PERMUT, PPMT, RANK.AVG, RANK.EQ, SERIESSUM, SLN, SMALL, T.DIST.2T, T.DIST.RT, TDIST, TYPE, WEEKDAY, WEEKNUM, XNPV, YEAR |
| boolean | 60 | ACCRINT, BASE, BESSELI, BESSELJ, BESSELK, BESSELY, BETA.INV, BETADIST, BETAINV, BIN2DEC, BIN2OCT, BINOM.INV, COMBINA, COMPLEX, CONFIDENCE.NORM, CUMIPMT, CUMPRINC, DEC2BIN, DEC2OCT, DELTA, DOLLARDE, DOLLARFR, EFFECT, ERFC, FACTDOUBLE, FVSCHEDULE, GAMMA.INV, GAMMAINV, GESTEP, HEX2BIN, HEX2DEC, HEX2OCT, IMAGINARY, IMCSC, IMCSCH, IMPOWER, IMREAL, ISEVEN, ISODD, LOGINV, LOGNORM.INV, MROUND, NOMINAL, NORM.INV, NORM.S.INV, NORMINV, NORMSINV, OCT2BIN, OCT2DEC, PERMUT, QUOTIENT, RANK.AVG, RANK.EQ, SERIESSUM, SQRTPI, T.DIST.2T, T.DIST.RT, T.INV, TDIST, XNPV |
| negative | 45 | ACOT, ATAN2, BASE, BETADIST, BETAINV, BINOM.INV, CHISQ.DIST, COMBIN, CONFIDENCE.NORM, CRITBINOM, EXPON.DIST, EXPONDIST, F.DIST, FACT, FACTDOUBLE, GAMMA.DIST, GAMMADIST, HYPGEOMDIST, IPMT, LARGE, LOGINV, LOGNORM.DIST, LOGNORM.INV, NEGBINOM.DIST, NEGBINOMDIST, NORM.INV, NORM.S.INV, NORMINV, NORMSINV, PERMUT, PERMUTATIONA, PMT, POISSON.DIST, PPMT, PV, RANK.AVG, RANK.EQ, RATE, REPLACE, RRI, STANDARDIZE, SYD, T.DIST.RT, UNICHAR, WEIBULL.DIST |
| blank range | 44 | CORREL, COVAR, COVARIANCE.P, COVARIANCE.S, DEVSQ, FORECAST, FREQUENCY, GCD, GEOMEAN, HARMEAN, INTERCEPT, IRR, KURT, LARGE, LCM, MODE.SNGL, MULTINOMIAL, PEARSON, PERCENTILE.EXC, PERCENTILE.INC, PERCENTRANK.EXC, QUARTILE.EXC, RSQ, SKEW, SKEW.P, SLOPE, SMALL, STDEV.P, STDEV.S, STDEVA, STDEVP, STDEVPA, STEYX, SUMSQ, SUMX2MY2, SUMX2PY2, SUMXMY2, VAR.P, VAR.S, VARA, VARP, VARPA, Z.TEST, ZTEST |
| large | 38 | ATAN2, BINOM.DIST, BITLSHIFT, BITOR, BITRSHIFT, BITXOR, COMBIN, COMBINA, COS, COT, COTH, CSC, DEC2BIN, DEC2OCT, F.DIST, F.INV, F.INV.RT, FACT, FACTDOUBLE, FISHERINV, GAMMAINV, HEX2BIN, HYPGEOMDIST, LARGE, MID, PERMUT, PERMUTATIONA, QUOTIENT, RANK.AVG, RANK.EQ, SEC, SIN, T.DIST.2T, T.DIST.RT, TAN, TANH, UNICHAR, WEIBULL.DIST |
| zero | 36 | ATAN2, BETADIST, BETAINV, BINOM.INV, CHISQ.DIST.RT, CONFIDENCE.NORM, CRITBINOM, DAY, EOMONTH, EXPON.DIST, EXPONDIST, F.INV, HLOOKUP, INDEX, LOGINV, LOGNORM.INV, MONTH, NORM.INV, NORM.S.INV, NORMINV, NORMSINV, PERMUT, PPMT, QUOTIENT, RANK.AVG, RANK.EQ, SERIESSUM, SMALL, T.DIST.2T, T.DIST.RT, TDIST, VLOOKUP, WEEKDAY, WEEKNUM, XNPV, YEAR |
| range: mixed | 31 | CORREL, COVAR, COVARIANCE.P, COVARIANCE.S, FORECAST, FREQUENCY, GCD, INTERCEPT, IRR, KURT, LCM, MIRR, MODE.SNGL, MULTINOMIAL, PEARSON, PERCENTRANK.EXC, RSQ, SLOPE, SMALL, STDEVA, STDEVPA, STEYX, SUMPRODUCT, SUMX2MY2, SUMX2PY2, SUMXMY2, TEXTJOIN, VARA, VARPA, Z.TEST, ZTEST |
| empty string | 29 | BIN2DEC, BIN2OCT, DATEDIF, HEX2BIN, HEX2DEC, HEX2OCT, IMABS, IMAGINARY, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMCSC, IMCSCH, IMDIV, IMEXP, IMPOWER, IMPRODUCT, IMREAL, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN, OCT2BIN, OCT2DEC |
| spaces | 23 | DATEDIF, DECIMAL, IMABS, IMARGUMENT, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN |
| case | 22 | IMABS, IMARGUMENT, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN, SWITCH |
| Cyrillic | 21 | IMABS, IMARGUMENT, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN |
| Cyrillic in a cell | 21 | IMABS, IMARGUMENT, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN |
| fractional | 18 | ATAN2, BETADIST, BITAND, BITLSHIFT, BITOR, BITRSHIFT, BITXOR, CHISQ.DIST.RT, COMBINA, FV, IPMT, LARGE, PERMUT, PV, RANK.AVG, RANK.EQ, T.DIST.2T, T.DIST.RT |
| number | 17 | DATEVALUE, DECIMAL, IMCOS, IMCOSH, IMCOT, IMCSC, IMCSCH, IMEXP, IMLN, IMLOG10, IMLOG2, IMREAL, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT |
| different size | 15 | CORREL, COVAR, COVARIANCE.P, COVARIANCE.S, FORECAST, GROWTH, PEARSON, RSQ, SLOPE, STEYX, SUMPRODUCT, SUMX2MY2, SUMX2PY2, SUMXMY2, XNPV |
| ISO date text | 14 | ACCRINT, COUPDAYS, DATEDIF, DAY, DAYS, EDATE, EOMONTH, ISOWEEKNUM, MONTH, PRICEDISC, WEEKDAY, WEEKNUM, YEAR, YEARFRAC |
| m/d/yyyy date text | 14 | ACCRINT, COUPDAYS, DATEDIF, DAY, DAYS, EDATE, EOMONTH, ISOWEEKNUM, MONTH, PRICEDISC, WEEKDAY, WEEKNUM, YEAR, YEARFRAC |
| date in a cell | 13 | BETA.INV, COTH, FACT, FACTDOUBLE, FISHERINV, HYPGEOM.DIST, IMPOWER, LARGE, RANK.AVG, RANK.EQ, ROUND, T.DIST.RT, TANH |
| blank cell argument | 11 | KURT, MINA, MODE.SNGL, SKEW, SKEW.P, STDEV.S, STDEVA, STDEVPA, VAR.S, VARA, VARPA |
| numeric text | 10 | ATAN2, BETADIST, BETAINV, CUMIPMT, CUMPRINC, DOLLARDE, NPER, RANK.AVG, RANK.EQ, T.INV.2T |
| number text with a point | 10 | BITOR, CHISQ.DIST.RT, DOLLARDE, F.DIST, LARGE, RANK.AVG, RANK.EQ, T.DIST.2T, T.INV, T.INV.2T |
| above 1 | 10 | CONFIDENCE.NORM, CRITBINOM, LOGINV, LOGNORM.INV, NEGBINOMDIST, NORM.INV, NORM.S.INV, NORMINV, NORMSINV, T.INV.2T |
| reference to numeric text | 10 | KURT, LCM, MAXA, MINA, MODE.SNGL, MULTINOMIAL, SKEW, SKEW.P, STDEV.S, VAR.S |
| reference to a boolean | 9 | GCD, KURT, LCM, MODE.SNGL, MULTINOMIAL, SKEW, SKEW.P, STDEV.S, VAR.S |
| false leap day 1900-02-29 | 8 | ACCRINT, COUPDAYS, DAY, EDATE, EOMONTH, MONTH, WEEKDAY, YEARFRAC |
| text argument | 8 | AVERAGEA, MAXA, MINA, MODE.SNGL, STDEVA, STDEVPA, VARA, VARPA |
| boolean argument | 8 | COUNT, GCD, KURT, LCM, MODE.SNGL, MULTINOMIAL, SKEW, SKEW.P |
| numeric text argument | 6 | AVERAGEA, COUNT, KURT, MODE.SNGL, SKEW, SKEW.P |
| negative and fractional | 6 | GCD, HARMEAN, KURT, LCM, MODE.SNGL, MULTINOMIAL |
| range with an error | 6 | IRR, STDEVA, STDEVPA, TEXTJOIN, VARA, VARPA |
| numeric text in a cell | 5 | ATAN2, CUMIPMT, NPER, RANK.AVG, RANK.EQ |
| 1900-03-01 | 4 | COUPDAYS, DAYS, PRICEDISC, YEARFRAC |
| 1900-02-28 | 3 | COUPDAYS, WEEKDAY, YEARFRAC |
| leap day 2024-02-29 | 3 | DATEDIF, PRICEDISC, YEARFRAC |
| time text | 3 | HOUR, MINUTE, SECOND |
| wildcard * | 3 | MINIFS, SEARCH, VLOOKUP |
| wildcard ? | 3 | MINIFS, SEARCH, VLOOKUP |
| error input | 2 | MATCH, VLOOKUP |
| #DIV/0! | 2 | MATCH, SWITCH |
| zero in known_y | 1 | GROWTH |
| descending -1 | 1 | MATCH |
| escaped ~* | 1 | MINIFS |
| Cyrillic, any case | 1 | MINIFS |
| date text | 1 | MINIFS |
| not equal | 1 | MINIFS |
| empty criterion | 1 | MINIFS |
| boolean criterion | 1 | MINIFS |
| two criteria | 1 | MINIFS |
| no sign change | 1 | MIRR |
| escaped ~? | 1 | SEARCH |
| Cyrillic and case | 1 | SWITCH |
| dates out of order | 1 | XNPV |
| cash flow: mixed | 1 | XNPV |
<!-- /auto:tags -->

Recurring causes, from a sample of the differing cases:
- a boolean literal in a numeric argument: the functions of the former Analysis ToolPak (BESSEL*, COMPLEX, CONVERT, the base conversions, ACCRINT, CUMIPMT, CUMPRINC, DOLLARDE and others) give `#VALUE!` in Excel (37 cases), while the ordinary math functions take TRUE as 1;
- the domain: a negative number, zero, a probability above 1, a blank cell as 0 give `#NUM!` in Excel's distribution and inverse distribution functions, where formulajs computes a number (40 cases on a negative number, 22 on a blank cell, 20 on zero); trigonometry of 1E10 is `#NUM!` in Excel and a number in formulajs;
- a blank range: Excel answers `#DIV/0!`, `#NUM!` or `#N/A` depending on the function, formulajs with another code or a number;
- a mix in paired ranges (CORREL, SLOPE, PEARSON, COVARIANCE.*): Excel drops the pairs where at least one value is not a number, formulajs answers `#VALUE!`; on ranges of different sizes Excel gives `#N/A`, formulajs computes over the matching part or answers `#NUM!`;
- date text in a date argument, and dates at the 1900 boundary: serial number 60 (the false 1900-02-29) cannot be a JS Date and becomes 1 March, and Excel reads 0 as 0 January 1900;
- the error code on wrong input: the complex functions (IM*) give `#NUM!` in Excel on text that is not a complex number (Cyrillic, spaces), `#VALUE!` in formulajs; Excel refuses an empty string (`#NUM!`), formulajs reads it as zero.

## Where typical cases differ

Typical cases differ through the adapter for 82 functions: the 81 of the row "typical cases differ already" and CHIINV, which hangs. For three of them only the writing of the answer differs: hexadecimal digits in lower case (BIN2HEX, DEC2HEX, OCT2HEX); the proposal puts them with the adapter. A complex number with 17 significant digits instead of 15 (15 IM* functions) now matches by parts. The rest by group:
- old names: formulajs puts in the new functions with other semantics. CHIDIST and CHIINV (the right tail in Excel, CHISQ.DIST and CHISQ.INV in formulajs), FDIST and FINV, TINV (two-tailed in Excel), NORMSDIST and LOGNORMDIST (the old names have no cumulative argument, and without it formulajs computes the density); TTEST and T.TEST ignore the number of tails and the test type, FTEST and F.TEST do not give a probability, CHITEST and CHISQ.TEST round the answer to six digits (as LOGEST does);
- databases: all 12 D functions. formulajs expects the database and the criteria by column (`database[column][row]`), while Excel's range goes by row; the field number is off by one;
- criteria: SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS (and MINIFS at the edges): case counts, and there are no wildcards or date text. The engine has its own versions of these already;
- dates and times: DATE (years 0 to 1899 and 1900-02-29), TIME (hours over 24 do not wrap), TIMEVALUE (`"2:24 AM"`, `"18:30"` give `#VALUE!`), DAYS360 (the end of February), NETWORKDAYS (dates in reverse order), NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL (holidays);
- financial: DB (the seventh period with month 7), DISC, TBILLEQ, TBILLPRICE, TBILLYIELD (the day count), XIRR;
- arrays and references: COLUMN and ROW (formulajs gets values and does not know the address), LOOKUP in its array form, SORT (numbers sorted as text), UNIQUE, MODE.MULT (another order), LINEST and LOGEST (a one-dimensional array instead of a row), TREND without new x, AGGREGATE (option 6 does not skip errors), IFERROR (a reference to a blank cell gives `null`, which the adapter turns into `#VALUE!`; Excel gives 0);
- math and statistics: ARABIC (lower case), ISO.CEILING (0 without a significance), MOD (a negative divisor), ROMAN (forms 1 and 4), BETA.DIST (the density with bounds twice as large), T.DIST (`#NUM!`), TRIMMEAN, QUARTILE.INC (quarts 0 and 4), PERCENTRANK.INC (the error code), ERF with two arguments, CONVERT (`"F"` to `"C"`), GAMMA and CONFIDENCE.T (softly only);
- text: PROPER (a capital after a digit and a hyphen; formulajs leaves Cyrillic alone: `PROPER("привет мир")` stays lower case), UPPER (`ß` stays `ß` in Excel and becomes `SS` in formulajs), CHAR and CODE (Excel for Mac takes the Mac encoding, `CHAR(200)` is `»`; `CODE("Ж")` is 95 with the English interface and 134 with the Russian one), DOLLAR (a negative number in parentheses), NUMBERVALUE (percentages, a space as separator), SUBSTITUTE (the instance number), TEXT (date formats), VALUE (time and date text, `"1,5"`).

## formulajs hangs

<!-- auto:hangs -->
| case | formula | what | Excel en-US |
|---|---|---|---|
| CHIINV#7 | `CHIINV(0.050001,10000000000)` | hangs over 5 s | #NUM |
| CHISQ.INV#14 | `_xlfn.CHISQ.INV(0.93,10000000000)` | hangs over 5 s | #NUM |
| CHISQ.INV.RT#15 | `_xlfn.CHISQ.INV.RT(0.050001,10000000000)` | hangs over 5 s | #NUM |
<!-- /auto:hangs -->

In the current set formulajs hangs on three inverse distributions with 1E10 degrees of freedom; Excel answers `#NUM!`. Before edge values were capped, `BESSELJ(1.9,1E10)`, `CUMPRINC(…,13,1E10,0)`, `WORKDAY(39722,1E10)`, `BIN2OCT("1001",45688)`, `DEC2OCT(58,45688)`, `HEX2BIN("F",45688)`, `HEX2OCT("F",45688)` and `MUNIT(45688)` hung or ran out of memory too. Excel checks these arguments at once; in formulajs they are loops over the number of digits, days or the order, and arrays of that size. Such functions need an argument guard before the call if they go into the whitelist.

## LibreOffice against Excel

<!-- auto:libreoffice -->
LibreOffice differs from Excel en-US in 1345 cases of 5275, for 319 functions of 402. By kind: other error code 775, value instead of an error 208, error instead of a value 163, number 153, string 38, other type 6, string: number format 1, boolean 1.

On typical cases it differs 59 times, for 38 functions:

| case | formula | Excel en-US | LibreOffice |
|---|---|---|---|
| ACCRINT#2 | `ACCRINT(39508,39691,39569,0.1,1000,2,0,FALSE)` | 16.6666666667 | #VALUE |
| ACCRINT#3 | `ACCRINT(39512,39691,39569,0.1,1000,2,0,TRUE)` | 15.5555555556 | #VALUE |
| BESSELJ#1 | `BESSELJ(1.9,2)` | 0.32992582867 | 0.329925727692 |
| BINOM.DIST.RANGE#1 | `_xlfn.BINOM.DIST.RANGE(60,0.75,48)` | 0.083974967429 | #NAME |
| BINOM.DIST.RANGE#2 | `_xlfn.BINOM.DIST.RANGE(60,0.75,45,50)` | 0.523629793472 | #NAME |
| CHAR#3 | `CHAR(200)` | "»" | "�" |
| CHAR#4 | `CHAR(0)` | #VALUE | "_x0000_" |
| CODE#3 | `CODE("Ж")` | 95 | 208 |
| CODE#4 | `CODE("")` | #VALUE | 0 |
| COLUMN#2 | `COLUMN(D!L1:N6)` | #SPILL | 12 |
| CONCATENATE#3 | `CONCATENATE(1.5,TRUE)` | "1.5TRUE" | "1.51" |
| CONVERT#1 | `CONVERT(1,"lbm","kg")` | 0.45359237 | 0.453592309749 |
| CONVERT#3 | `CONVERT(2.5,"ft","sec")` | #NA | #VALUE |
| DATE#4 | `DATE(108,1,2)` | 39449 | #VALUE |
| DATE#5 | `DATE(1900,2,29)` | 60 | 61 |
| DATEVALUE#4 | `DATEVALUE("2011/02/23")` | 40597 | #VALUE |
| DGET#1 | `DGET(D!U1:Y7,"Yield",D!AA1:AB2)` | #NUM | #VALUE |
| DGET#2 | `DGET(D!U1:Y7,4,D!AC1:AC2)` | #NUM | #VALUE |
| DOLLAR#3 | `DOLLAR(-1234.567,-2)` | "($1,200)" | "-$1,200" |
| DOLLAR#4 | `DOLLAR(-0.123,4)` | "($0.1230)" | "-$0.1230" |
| FIND#5 | `FIND("","abc")` | 1 | #VALUE |
| FLOOR#3 | `FLOOR(2.5,-2)` | #NUM | #VALUE |
| GAMMA#3 | `_xlfn.GAMMA(0)` | #NUM | #VALUE |
| GAMMA#4 | `_xlfn.GAMMA(-2)` | #NUM | #VALUE |
| IMCOSH#1 | `_xlfn.IMCOSH("4+3i")` | "-27.0349456030742+3.85115333481178i" | #NAME |
| IMCOT#1 | `_xlfn.IMCOT("4+3i")` | "0.00490118239430447-0.999266927805902i" | #NAME |
| IMCSC#1 | `_xlfn.IMCSC("4+3i")` | "-0.0754898329158637+0.0648774713706355i" | #NAME |
| IMCSCH#1 | `_xlfn.IMCSCH("4+3i")` | "-0.036275889628626-0.0051744731840194i" | #NAME |
| IMSEC#1 | `_xlfn.IMSEC("4+3i")` | "-0.0652940278579471-0.0752249603027732i" | #NAME |
| IMSECH#1 | `_xlfn.IMSECH("4+3i")` | "-0.0362534969158689-0.00516434460775318i" | #NAME |
| IMSINH#1 | `_xlfn.IMSINH("4+3i")` | "-27.0168132580039+3.85373803791938i" | #NAME |
| IMTAN#1 | `_xlfn.IMTAN("4+3i")` | "0.00490825806749606+1.00070953606723i" | #NAME |
| ISO.CEILING#1 | `_xlfn.ISO.CEILING(4.3)` | 5 | #NAME |
| ISO.CEILING#2 | `_xlfn.ISO.CEILING(-4.3,2)` | -4 | #NAME |
| ISO.CEILING#3 | `_xlfn.ISO.CEILING(4.3,-2)` | 6 | #NAME |
| MROUND#4 | `MROUND(5,-2)` | #NUM | 6 |
| NETWORKDAYS.INTL#1 | `_xlfn.NETWORKDAYS.INTL(38718,38748)` | 22 | #NAME |
| NETWORKDAYS.INTL#2 | `_xlfn.NETWORKDAYS.INTL(38718,38749,7,D!AP1:AP8)` | 22 | #NAME |
| NETWORKDAYS.INTL#3 | `_xlfn.NETWORKDAYS.INTL(38718,38749,"0010001")` | 22 | #NAME |
| NUMBERVALUE#2 | `_xlfn.NUMBERVALUE("3.5%")` | 0.035 | #VALUE |
| NUMBERVALUE#3 | `_xlfn.NUMBERVALUE("1 234.5")` | 1234.5 | #VALUE |
| NUMBERVALUE#5 | `_xlfn.NUMBERVALUE("2.5")` | 2.5 | #VALUE |
| PERCENTILE.EXC#3 | `_xlfn.PERCENTILE.EXC(D!B1:B8,0.05)` | #NUM | #VALUE |
| PERCENTRANK.EXC#3 | `_xlfn.PERCENTRANK.EXC(D!B1:B8,5.5)` | 0.263 | 0.264 |
| PERCENTRANK.INC#1 | `_xlfn.PERCENTRANK.INC(D!B1:B8,8)` | 0.285 | 0.286 |
| PERCENTRANK.INC#2 | `_xlfn.PERCENTRANK.INC(D!B1:B8,5.5,2)` | 0.19 | 0.2 |
| PERCENTRANK.INC#3 | `_xlfn.PERCENTRANK.INC(D!B1:B8,100)` | #NA | #VALUE |
| ROW#1 | `ROW(D!C1:C8)` | #SPILL | 1 |
| ROW#2 | `ROW(D!L1:N6)` | #SPILL | 1 |
| SQRT#2 | `SQRT(-16)` | #NUM | #VALUE |
| SUBTOTAL#4 | `SUBTOTAL(2,D!D1:D8)` | 3 | 5 |
| TYPE#3 | `TYPE(TRUE)` | 4 | 1 |
| UNICHAR#4 | `_xlfn.UNICHAR(0)` | #VALUE | "_x0000_" |
| UPPER#3 | `UPPER("ß")` | "ß" | "ẞ" |
| WORKDAY.INTL#1 | `_xlfn.WORKDAY.INTL(40909,90,11)` | 41013 | #NAME |
| WORKDAY.INTL#2 | `_xlfn.WORKDAY.INTL(40909,30,17)` | 40944 | #NAME |
| WORKDAY.INTL#3 | `_xlfn.WORKDAY.INTL(40909,30,"0000011",D!AP1:AP8)` | 40949 | #NAME |
| XIRR#1 | `XIRR(D!J1:J5,D!AU1:AU5)` | 0.373362535238 | 0.373362533519 |
| XIRR#2 | `XIRR(D!J1:J5,D!AU1:AU5,0.1)` | 0.373362535238 | 0.373362533519 |

Every function where LibreOffice differs in at least one case:

| function | cases that differ | example |
|---|---|---|
| BINOM.DIST.RANGE | 15 of 15 | `_xlfn.BINOM.DIST.RANGE(60,0.75,48)`: Excel 0.083974967429, LibreOffice #NAME |
| ISO.CEILING | 15 of 15 | `_xlfn.ISO.CEILING(4.3)`: Excel 5, LibreOffice #NAME |
| NETWORKDAYS.INTL | 15 of 15 | `_xlfn.NETWORKDAYS.INTL(38718,38748)`: Excel 22, LibreOffice #NAME |
| WORKDAY.INTL | 15 of 15 | `_xlfn.WORKDAY.INTL(40909,90,11)`: Excel 41013, LibreOffice #NAME |
| TBILLEQ | 12 of 15 | `TBILLEQ("2025-01-31",39600,0.0914)`: Excel #NUM, LibreOffice #VALUE |
| TBILLYIELD | 12 of 15 | `TBILLYIELD("2025-01-31",39600,98.45)`: Excel #NUM, LibreOffice #VALUE |
| IMDIV | 11 of 15 | `IMDIV("Привет Мир","10+24i")`: Excel #NUM, LibreOffice #VALUE |
| IMSUB | 11 of 15 | `IMSUB("Привет Мир","5+3i")`: Excel #NUM, LibreOffice #VALUE |
| TBILLPRICE | 11 of 15 | `TBILLPRICE("2025-01-31",39600,0.09)`: Excel #NUM, LibreOffice #VALUE |
| IMCOSH | 10 of 10 | `_xlfn.IMCOSH("4+3i")`: Excel "-27.0349456030742+3.85115333481178i", LibreOffice #NAME |
| IMCOT | 10 of 10 | `_xlfn.IMCOT("4+3i")`: Excel "0.00490118239430447-0.999266927805902i", LibreOffice #NAME |
| IMCSC | 10 of 10 | `_xlfn.IMCSC("4+3i")`: Excel "-0.0754898329158637+0.0648774713706355i", LibreOffice #NAME |
| IMCSCH | 10 of 10 | `_xlfn.IMCSCH("4+3i")`: Excel "-0.036275889628626-0.0051744731840194i", LibreOffice #NAME |
| IMPRODUCT | 10 of 15 | `IMPRODUCT("Привет Мир","5-3i")`: Excel #NUM, LibreOffice #VALUE |
| IMSEC | 10 of 10 | `_xlfn.IMSEC("4+3i")`: Excel "-0.0652940278579471-0.0752249603027732i", LibreOffice #NAME |
| IMSECH | 10 of 10 | `_xlfn.IMSECH("4+3i")`: Excel "-0.0362534969158689-0.00516434460775318i", LibreOffice #NAME |
| IMSINH | 10 of 10 | `_xlfn.IMSINH("4+3i")`: Excel "-27.0168132580039+3.85373803791938i", LibreOffice #NAME |
| IMTAN | 10 of 10 | `_xlfn.IMTAN("4+3i")`: Excel "0.00490825806749606+1.00070953606723i", LibreOffice #NAME |
| PERCENTRANK.INC | 10 of 15 | `_xlfn.PERCENTRANK.INC(D!B1:B8,8)`: Excel 0.285, LibreOffice 0.286 |
| ACCRINT | 9 of 15 | `ACCRINT(39508,39691,39569,0.1,1000,2,0,FALSE)`: Excel 16.6666666667, LibreOffice #VALUE |
| IMPOWER | 9 of 15 | `IMPOWER("Привет Мир",3)`: Excel #NUM, LibreOffice #VALUE |
| IMSUM | 9 of 15 | `IMSUM("Привет Мир","5-3i")`: Excel #NUM, LibreOffice #VALUE |
| OCT2HEX | 9 of 15 | `OCT2HEX("Привет Мир",4)`: Excel #NUM, LibreOffice #VALUE |
| BESSELK | 8 of 15 | `BESSELK(TRUE,1)`: Excel #VALUE, LibreOffice 0.601907227223 |
| CONVERT | 8 of 15 | `CONVERT(1,"lbm","kg")`: Excel 0.45359237, LibreOffice 0.453592309749 |
| CUMIPMT | 8 of 15 | `CUMIPMT("3",360,125000,13,24,0)`: Excel -4500000, LibreOffice -4502604 |
| DECIMAL | 8 of 15 | `_xlfn.DECIMAL("Привет Мир",16)`: Excel #NUM, LibreOffice #VALUE |
| F.INV.RT | 8 of 15 | `_xlfn.F.INV.RT(0.01,0.75,4)`: Excel #NUM, LibreOffice #VALUE |
| HEX2BIN | 8 of 15 | `HEX2BIN("Привет Мир",8)`: Excel #NUM, LibreOffice #VALUE |
| HEX2OCT | 8 of 15 | `HEX2OCT("Привет Мир",3)`: Excel #NUM, LibreOffice #VALUE |
| IMLOG2 | 8 of 10 | `IMLOG2("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| MODE.MULT | 8 of 14 | `INDEX(_xlfn.MODE.MULT(D!D1:D8),1,1)`: Excel #NA, LibreOffice 1 |
| MODE.SNGL | 8 of 12 | `_xlfn.MODE.SNGL(D!D1:D8)`: Excel #NA, LibreOffice 1 |
| NUMBERVALUE | 8 of 15 | `_xlfn.NUMBERVALUE("3.5%")`: Excel 0.035, LibreOffice #VALUE |
| XIRR | 8 of 8 | `XIRR(D!J1:J5,D!AU1:AU5)`: Excel 0.373362535238, LibreOffice 0.373362533519 |
| ATANH | 7 of 14 | `ATANH("3")`: Excel #NUM, LibreOffice #VALUE |
| BESSELI | 7 of 15 | `BESSELI(TRUE,1)`: Excel #VALUE, LibreOffice 0.565159103992 |
| BESSELY | 7 of 15 | `BESSELY(TRUE,1)`: Excel #VALUE, LibreOffice -0.7812128213 |
| CHAR | 7 of 15 | `CHAR(200)`: Excel "»", LibreOffice "�" |
| CODE | 7 of 12 | `CODE("Ж")`: Excel 95, LibreOffice 208 |
| COMBIN | 7 of 15 | `COMBIN(TRUE,2)`: Excel #NUM, LibreOffice #VALUE |
| COMBINA | 7 of 15 | `_xlfn.COMBINA(TRUE,3)`: Excel 1, LibreOffice #VALUE |
| COUPDAYS | 7 of 15 | `COUPDAYS("2025-01-31",40862,2,1)`: Excel #NUM, LibreOffice #VALUE |
| DATE | 7 of 15 | `DATE(108,1,2)`: Excel 39449, LibreOffice #VALUE |
| F.INV | 7 of 15 | `_xlfn.F.INV(TRUE,6,4)`: Excel #NUM, LibreOffice 1.34967413836e+162 |
| FINV | 7 of 15 | `FINV(D!A1,6,4)`: Excel #NUM, LibreOffice #VALUE |
| FISHER | 7 of 12 | `FISHER("3")`: Excel #NUM, LibreOffice #VALUE |
| HYPGEOMDIST | 7 of 15 | `HYPGEOMDIST(1,D!A6,8,20)`: Excel #NUM, LibreOffice #VALUE |
| IMABS | 7 of 10 | `IMABS("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| IMAGINARY | 7 of 12 | `IMAGINARY("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| IMARGUMENT | 7 of 10 | `IMARGUMENT("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| IMCONJUGATE | 7 of 10 | `IMCONJUGATE("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| IMCOS | 7 of 10 | `IMCOS("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| IMEXP | 7 of 10 | `IMEXP("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| IMLN | 7 of 10 | `IMLN("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| IMLOG10 | 7 of 10 | `IMLOG10("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| IMREAL | 7 of 10 | `IMREAL("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| IMSIN | 7 of 10 | `IMSIN("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| IMSQRT | 7 of 10 | `IMSQRT("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| LARGE | 7 of 15 | `LARGE(D!B1:B8,-2.5)`: Excel #NUM, LibreOffice #VALUE |
| PERCENTILE.EXC | 7 of 15 | `_xlfn.PERCENTILE.EXC(D!B1:B8,0.05)`: Excel #NUM, LibreOffice #VALUE |
| PERCENTRANK.EXC | 7 of 15 | `_xlfn.PERCENTRANK.EXC(D!B1:B8,5.5)`: Excel 0.263, LibreOffice 0.264 |
| PERMUT | 7 of 15 | `PERMUT(TRUE,3)`: Excel #NUM, LibreOffice #VALUE |
| RATE | 7 of 15 | `RATE(48,0.75,8000)`: Excel #NUM, LibreOffice #NA |
| SMALL | 7 of 15 | `SMALL(D!D1:D8,4)`: Excel #NUM, LibreOffice 3.5 |
| T.INV | 7 of 15 | `_xlfn.T.INV(TRUE,2)`: Excel #NUM, LibreOffice #VALUE |
| T.TEST | 7 of 15 | `_xlfn.T.TEST(D!D1:D8,D!C1:C8,2,1)`: Excel 0.323995741228, LibreOffice 0.155516420789 |
| UNICHAR | 7 of 15 | `_xlfn.UNICHAR(0)`: Excel #VALUE, LibreOffice "_x0000_" |
| XNPV | 7 of 15 | `XNPV(0.09,D!J1:J5,D!AU7:AU11)`: Excel #NUM, LibreOffice 2298.32224185 |
| BETAINV | 6 of 15 | `BETAINV(TRUE,8,10,1,3)`: Excel #NUM, LibreOffice 3 |
| BIN2DEC | 6 of 12 | `BIN2DEC("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| BIN2HEX | 6 of 15 | `BIN2HEX("Привет Мир",4)`: Excel #NUM, LibreOffice #VALUE |
| BIN2OCT | 6 of 15 | `BIN2OCT("Привет Мир",3)`: Excel #NUM, LibreOffice #VALUE |
| CHIINV | 6 of 15 | `CHIINV(D!A1,10)`: Excel #NUM, LibreOffice #VALUE |
| CHISQ.INV.RT | 6 of 15 | `_xlfn.CHISQ.INV.RT(D!A1,10)`: Excel #NUM, LibreOffice #VALUE |
| CONFIDENCE.NORM | 6 of 15 | `_xlfn.CONFIDENCE.NORM(TRUE,2.5,50)`: Excel #NUM, LibreOffice #VALUE |
| CONFIDENCE.T | 6 of 15 | `_xlfn.CONFIDENCE.T(TRUE,1,50)`: Excel #NUM, LibreOffice #VALUE |
| CORREL | 6 of 9 | `CORREL(D!D1:D8,D!C1:C8)`: Excel 0.880913833004, LibreOffice 0.784116609751 |
| COVAR | 6 of 9 | `COVAR(D!D1:D8,D!C1:C8)`: Excel 24, LibreOffice 21.52 |
| COVARIANCE.P | 6 of 9 | `_xlfn.COVARIANCE.P(D!D1:D8,D!C1:C8)`: Excel 24, LibreOffice 21.52 |
| COVARIANCE.S | 6 of 9 | `_xlfn.COVARIANCE.S(D!D1:D8,D!C1:C8)`: Excel 36, LibreOffice 26.9 |
| CRITBINOM | 6 of 15 | `CRITBINOM(6,-0.1,0.75)`: Excel #NUM, LibreOffice #VALUE |
| CUMPRINC | 6 of 15 | `CUMPRINC("3",360,125000,13,24,0)`: Excel -6.37900438848e-198, LibreOffice 2604 |
| DATEDIF | 6 of 15 | `DATEDIF("2025-01-31",37622,"Y")`: Excel #NUM, LibreOffice #VALUE |
| DAY | 6 of 14 | `DAY(60)`: Excel 29, LibreOffice 28 |
| DDB | 6 of 15 | `DDB("3",300,3650,1)`: Excel 0, LibreOffice #VALUE |
| DGET | 6 of 10 | `DGET(D!U1:Y7,"Yield",D!AA1:AB2)`: Excel #NUM, LibreOffice #VALUE |
| GAMMA.INV | 6 of 15 | `_xlfn.GAMMA.INV(TRUE,9,2)`: Excel #NUM, LibreOffice #VALUE |
| HEX2DEC | 6 of 12 | `HEX2DEC("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| INTERCEPT | 6 of 9 | `INTERCEPT(D!D1:D8,D!B1:B8)`: Excel 1.54654021777, LibreOffice 0.717032284768 |
| LCM | 6 of 12 | `LCM(D!D1:D8)`: Excel #VALUE, LibreOffice 0 |
| LOGINV | 6 of 15 | `LOGINV(TRUE,3.5,1.2)`: Excel #NUM, LibreOffice #VALUE |
| LOGNORM.INV | 6 of 15 | `_xlfn.LOGNORM.INV(TRUE,3.5,1.2)`: Excel #NUM, LibreOffice #VALUE |
| NOMINAL | 6 of 15 | `NOMINAL(TRUE,4)`: Excel #VALUE, LibreOffice 0.756828460011 |
| NORM.INV | 6 of 15 | `_xlfn.NORM.INV(TRUE,40,1.5)`: Excel #NUM, LibreOffice #VALUE |
| OCT2BIN | 6 of 15 | `OCT2BIN("Привет Мир",3)`: Excel #NUM, LibreOffice #VALUE |
| OCT2DEC | 6 of 11 | `OCT2DEC("Привет Мир")`: Excel #NUM, LibreOffice #VALUE |
| PEARSON | 6 of 9 | `PEARSON(D!D1:D8,D!C1:C8)`: Excel 0.880913833004, LibreOffice 0.784116609751 |
| QUARTILE.EXC | 6 of 15 | `_xlfn.QUARTILE.EXC(D!D1:D8,1)`: Excel 1, LibreOffice 0.5 |
| RSQ | 6 of 9 | `RSQ(D!D1:D8,D!C1:C8)`: Excel 0.776009181179, LibreOffice 0.614838857687 |
| SLOPE | 6 of 9 | `SLOPE(D!D1:D8,D!B1:B8)`: Excel 0.086406743941, LibreOffice 0.100579470199 |
| STEYX | 6 of 9 | `STEYX(D!D1:D8,D!B1:B8)`: Excel 1.40249163595, LibreOffice 1.28835740006 |
| SUMX2MY2 | 6 of 10 | `SUMX2MY2(D!D1:D8,D!C1:C8)`: Excel -1646.75, LibreOffice -2007.75 |
| SUMX2PY2 | 6 of 10 | `SUMX2PY2(D!D1:D8,D!C1:C8)`: Excel 1723.25, LibreOffice 2086.25 |
| SUMXMY2 | 6 of 10 | `SUMXMY2(D!D1:D8,D!C1:C8)`: Excel 1256.25, LibreOffice #VALUE |
| T.INV.2T | 6 of 15 | `_xlfn.T.INV.2T(D!A1,60)`: Excel #NUM, LibreOffice #VALUE |
| TINV | 6 of 15 | `TINV(D!A1,60)`: Excel #NUM, LibreOffice #VALUE |
| TTEST | 6 of 15 | `TTEST(D!D1:D8,D!C1:C8,2,1)`: Excel 0.323995741228, LibreOffice 0.155516420789 |
| BESSELJ | 5 of 15 | `BESSELJ(1.9,2)`: Excel 0.32992582867, LibreOffice 0.329925727692 |
| BETA.INV | 5 of 15 | `_xlfn.BETA.INV(TRUE,8,10,1,3)`: Excel #NUM, LibreOffice 3 |
| BINOM.INV | 5 of 15 | `_xlfn.BINOM.INV(6,0,0.75)`: Excel #NUM, LibreOffice 0 |
| CHISQ.INV | 5 of 15 | `_xlfn.CHISQ.INV(TRUE,1)`: Excel #NUM, LibreOffice #VALUE |
| DAYS360 | 5 of 15 | `DAYS360(40573,59)`: Excel -39932, LibreOffice -39933 |
| DB | 5 of 15 | `DB("3",100000,6,1,7)`: Excel -8.17775, LibreOffice #VALUE |
| DEC2BIN | 5 of 15 | `DEC2BIN(TRUE,4)`: Excel #VALUE, LibreOffice "0001" |
| EFFECT | 5 of 15 | `EFFECT(TRUE,4)`: Excel #VALUE, LibreOffice 1.44140625 |
| GAMMAINV | 5 of 15 | `GAMMAINV(TRUE,9,2)`: Excel #NUM, LibreOffice #VALUE |
| ISEVEN | 5 of 15 | `ISEVEN("3")`: Excel FALSE, LibreOffice #VALUE |
| ISODD | 5 of 15 | `ISODD("3")`: Excel TRUE, LibreOffice #VALUE |
| MULTINOMIAL | 5 of 11 | `MULTINOMIAL(D!D1:D8)`: Excel #VALUE, LibreOffice 5040 |
| NORM.S.INV | 5 of 10 | `_xlfn.NORM.S.INV(TRUE)`: Excel #NUM, LibreOffice #VALUE |
| NORMINV | 5 of 15 | `NORMINV(TRUE,40,1.5)`: Excel #NUM, LibreOffice #VALUE |
| NORMSINV | 5 of 10 | `NORMSINV(TRUE)`: Excel #NUM, LibreOffice #VALUE |
| PDURATION | 5 of 15 | `_xlfn.PDURATION(0.025,2000,D!A1)`: Excel #NUM, LibreOffice #VALUE |
| PRICEDISC | 5 of 15 | `PRICEDISC("2025-01-31",39508,0.0525,100,2)`: Excel #NUM, LibreOffice #VALUE |
| QUARTILE.INC | 5 of 15 | `_xlfn.QUARTILE.INC(D!D1:D8,1)`: Excel 2.25, LibreOffice 1 |
| TIME | 5 of 15 | `TIME(12,0,10000000000)`: Excel #NUM, LibreOffice 0.240740740741 |
| TRIMMEAN | 5 of 15 | `TRIMMEAN(D!D1:D8,0.2)`: Excel 3.16666666667, LibreOffice 2.1 |
| ACOSH | 4 of 14 | `ACOSH(D!A1)`: Excel #NUM, LibreOffice #VALUE |
| ACOTH | 4 of 13 | `_xlfn.ACOTH(TRUE)`: Excel #NUM, LibreOffice #VALUE |
| BINOM.DIST | 4 of 15 | `_xlfn.BINOM.DIST(6,"1.5",0.5,FALSE)`: Excel #NUM, LibreOffice #VALUE |
| CHOOSECOLS | 4 of 19 | `INDEX(_xlfn.CHOOSECOLS(D!L1:N6,"3",3),1,1)`: Excel 10, LibreOffice #VALUE |
| DEC2HEX | 4 of 15 | `DEC2HEX(TRUE,4)`: Excel #VALUE, LibreOffice "0001" |
| DEC2OCT | 4 of 15 | `DEC2OCT(TRUE,3)`: Excel #VALUE, LibreOffice "001" |
| EDATE | 4 of 15 | `EDATE(40558,10000000000)`: Excel #NUM, LibreOffice #VALUE |
| EXPAND | 4 of 19 | `INDEX(_xlfn.EXPAND(D!P1:Q2,3,3,D!A7),1,1)`: Excel 1, LibreOffice #DIV0 |
| F.DIST.RT | 4 of 15 | `_xlfn.F.DIST.RT(15.2069,0,4)`: Excel #NUM, LibreOffice #VALUE |
| F.TEST | 4 of 9 | `_xlfn.F.TEST(D!D1:D8,D!C1:C8)`: Excel 0.046918420262, LibreOffice 0.0030714686204 |
| FACTDOUBLE | 4 of 14 | `FACTDOUBLE(TRUE)`: Excel #VALUE, LibreOffice 1 |
| FORECAST | 4 of 15 | `FORECAST(30,D!D1:D8,D!B1:B8)`: Excel 4.138742536, LibreOffice 3.73441639073 |
| FTEST | 4 of 9 | `FTEST(D!D1:D8,D!C1:C8)`: Excel 0.046918420262, LibreOffice 0.0030714686204 |
| GCD | 4 of 13 | `GCD(D!D1:D8)`: Excel #VALUE, LibreOffice 1 |
| HARMEAN | 4 of 12 | `HARMEAN(D!D1:D8)`: Excel 2.01923076923, LibreOffice #VALUE |
| HYPGEOM.DIST | 4 of 15 | `_xlfn.HYPGEOM.DIST(1,D!A1,8,20,TRUE)`: Excel 1, LibreOffice #VALUE |
| LOGNORMDIST | 4 of 15 | `LOGNORMDIST(D!A1,3.5,1.2)`: Excel #NUM, LibreOffice 0 |
| MONTH | 4 of 15 | `MONTH(0)`: Excel 1, LibreOffice 12 |
| PERCENTILE.INC | 4 of 15 | `_xlfn.PERCENTILE.INC(D!D1:D8,0.3)`: Excel 2.5, LibreOffice 1 |
| PPMT | 4 of 15 | `PPMT(0.00833333333333333,0,24,2000)`: Excel #NUM, LibreOffice #VALUE |
| RRI | 4 of 15 | `_xlfn.RRI(D!A1,10000,11000)`: Excel #NUM, LibreOffice #VALUE |
| SYD | 4 of 15 | `SYD(30000,7500,0.75,1)`: Excel #NUM, LibreOffice 25714.2857143 |
| TDIST | 4 of 15 | `TDIST(1.959999998,60,"3")`: Excel #NUM, LibreOffice #VALUE |
| YEAR | 4 of 15 | `YEAR(0)`: Excel 1900, LibreOffice 1899 |
| YEARFRAC | 4 of 15 | `YEARFRAC(40909,59)`: Excel 111.836111111, LibreOffice 111.844444444 |
| BASE | 3 of 15 | `_xlfn.BASE(-2.5,2)`: Excel #NUM, LibreOffice #VALUE |
| BETADIST | 3 of 15 | `BETADIST(D!A1,8,10,1,3)`: Excel #NUM, LibreOffice 0 |
| BINOMDIST | 3 of 15 | `BINOMDIST(6,"3",0.5,FALSE)`: Excel #NUM, LibreOffice #VALUE |
| BITAND | 3 of 15 | `_xlfn.BITAND(1,2.7)`: Excel #NUM, LibreOffice 0 |
| BITLSHIFT | 3 of 15 | `_xlfn.BITLSHIFT(-2.5,2)`: Excel #NUM, LibreOffice #VALUE |
| BITOR | 3 of 15 | `_xlfn.BITOR(23,"1.5")`: Excel #NUM, LibreOffice 23 |
| BITRSHIFT | 3 of 15 | `_xlfn.BITRSHIFT(13,-2.5)`: Excel 52, LibreOffice 104 |
| CHISQ.DIST | 3 of 15 | `_xlfn.CHISQ.DIST(0.5,D!A1,TRUE)`: Excel #NUM, LibreOffice #VALUE |
| CONCAT | 3 of 12 | `_xlfn.CONCAT(D!D1:D8)`: Excel "12abcTRUE5FALSE3.5", LibreOffice "12abc1503.5" |
| COT | 3 of 13 | `_xlfn.COT(D!A1)`: Excel #DIV0, LibreOffice #NUM |
| CSC | 3 of 13 | `_xlfn.CSC(D!A1)`: Excel #DIV0, LibreOffice #NUM |
| DEVSQ | 3 of 12 | `DEVSQ(D!D1:D8)`: Excel 8.16666666667, LibreOffice 17.2 |
| DISC | 3 of 15 | `DISC(43282,54058,TRUE,100,1)`: Excel #VALUE, LibreOffice 0.0335565593793 |
| DOLLAR | 3 of 15 | `DOLLAR(-1234.567,-2)`: Excel "($1,200)", LibreOffice "-$1,200" |
| DOLLARDE | 3 of 15 | `DOLLARDE(TRUE,16)`: Excel #VALUE, LibreOffice 1 |
| DSTDEV | 3 of 10 | `DSTDEV(D!U1:Y7,"Age",D!AE1:AE2)`: Excel #DIV0, LibreOffice #NUM |
| DVAR | 3 of 10 | `DVAR(D!U1:Y7,"Age",D!AE1:AE2)`: Excel #DIV0, LibreOffice #NUM |
| EOMONTH | 3 of 15 | `EOMONTH(40544,10000000000)`: Excel #NUM, LibreOffice #VALUE |
| EXPONDIST | 3 of 15 | `EXPONDIST(0.2,0,TRUE)`: Excel #NUM, LibreOffice #VALUE |
| F.DIST | 3 of 15 | `_xlfn.F.DIST(15.2069,D!A1,4,TRUE)`: Excel #NUM, LibreOffice #VALUE |
| FACT | 3 of 14 | `FACT(-2.5)`: Excel #NUM, LibreOffice #VALUE |
| FDIST | 3 of 15 | `FDIST(15.2069,10000000000,4)`: Excel 0.00792646915308, LibreOffice #VALUE |
| FIND | 3 of 15 | `FIND("","abc")`: Excel 1, LibreOffice #VALUE |
| GAMMA | 3 of 15 | `_xlfn.GAMMA(0)`: Excel #NUM, LibreOffice #VALUE |
| GAMMA.DIST | 3 of 15 | `_xlfn.GAMMA.DIST(10.00001131,D!A1,2,FALSE)`: Excel #NUM, LibreOffice #VALUE |
| GAMMALN | 3 of 13 | `GAMMALN(D!A1)`: Excel #NUM, LibreOffice #VALUE |
| GAMMALN.PRECISE | 3 of 13 | `_xlfn.GAMMALN.PRECISE(D!A1)`: Excel #NUM, LibreOffice #VALUE |
| GEOMEAN | 3 of 12 | `GEOMEAN(D!D1:D8)`: Excel 2.59624705093, LibreOffice 0 |
| INDEX | 3 of 15 | `INDEX(D!L1:N6,2,D!A1)`: Excel #SPILL, LibreOffice #VALUE |
| IRR | 3 of 7 | `IRR(D!I13:I17)`: Excel #VALUE, LibreOffice #NA |
| LN | 3 of 14 | `LN(D!A1)`: Excel #NUM, LibreOffice #VALUE |
| LOG | 3 of 15 | `LOG(D!A1)`: Excel #NUM, LibreOffice #VALUE |
| LOG10 | 3 of 14 | `LOG10(D!A1)`: Excel #NUM, LibreOffice #VALUE |
| MAXA | 3 of 12 | `MAXA("2",1)`: Excel 2, LibreOffice 1 |
| MEDIAN | 3 of 12 | `MEDIAN(D!D1:D8)`: Excel 3.5, LibreOffice 1 |
| MIN | 3 of 12 | `MIN(D!D1:D8)`: Excel 1, LibreOffice 0 |
| MINA | 3 of 12 | `MINA("2",1)`: Excel 1, LibreOffice 0 |
| MROUND | 3 of 15 | `MROUND(5,-2)`: Excel #NUM, LibreOffice 6 |
| NEGBINOMDIST | 3 of 15 | `NEGBINOMDIST(10,5,-0.1)`: Excel #NUM, LibreOffice #VALUE |
| NETWORKDAYS | 3 of 15 | `NETWORKDAYS(41183,-1)`: Excel #NUM, LibreOffice -29417 |
| ROW | 3 of 4 | `ROW(D!C1:C8)`: Excel #SPILL, LibreOffice 1 |
| SERIESSUM | 3 of 15 | `SERIESSUM(TRUE,0,2,D!AR1:AR4)`: Excel #VALUE, LibreOffice 0.540277777778 |
| STDEV.S | 3 of 12 | `_xlfn.STDEV.S(D!D1:D8)`: Excel 2.02072594216, LibreOffice 2.07364413533 |
| SUM | 3 of 13 | `SUM(D!D1:D8)`: Excel 9.5, LibreOffice 10.5 |
| SUMSQ | 3 of 12 | `SUMSQ(D!D1:D8)`: Excel 38.25, LibreOffice 39.25 |
| T.DIST.2T | 3 of 15 | `_xlfn.T.DIST.2T(-2.5,60)`: Excel #NUM, LibreOffice #VALUE |
| TEXTJOIN | 3 of 15 | `_xlfn.TEXTJOIN(" ",TRUE,D!D1:D8)`: Excel "1 2 abc TRUE 5 FALSE 3.5", LibreOffice "1 2 abc 1 5 0 3.5" |
| TYPE | 3 of 11 | `TYPE(TRUE)`: Excel 4, LibreOffice 1 |
| VAR.S | 3 of 12 | `_xlfn.VAR.S(D!D1:D8)`: Excel 4.08333333333, LibreOffice 4.3 |
| WEEKNUM | 3 of 15 | `WEEKNUM(0)`: Excel 0, LibreOffice 52 |
| AND | 2 of 15 | `AND("2",1)`: Excel TRUE, LibreOffice #VALUE |
| AVEDEV | 2 of 12 | `AVEDEV(D!D1:D8)`: Excel 1.44444444444, LibreOffice 1.72 |
| AVERAGE | 2 of 12 | `AVERAGE(D!D1:D8)`: Excel 3.16666666667, LibreOffice 2.1 |
| AVERAGEA | 2 of 12 | `AVERAGEA("2",1)`: Excel 1.5, LibreOffice 0.5 |
| AVERAGEIF | 2 of 15 | `AVERAGEIF(D!D1:D8,2,D!B1:B8)`: Excel 8, LibreOffice #DIV0 |
| AVERAGEIFS | 2 of 16 | `AVERAGEIFS(D!B1:B8,D!D1:D8,2)`: Excel 8, LibreOffice #DIV0 |
| BITXOR | 2 of 15 | `_xlfn.BITXOR(-2.5,3)`: Excel #NUM, LibreOffice #VALUE |
| CHIDIST | 2 of 15 | `CHIDIST(18.307,D!A1)`: Excel #NUM, LibreOffice #VALUE |
| CHISQ.DIST.RT | 2 of 15 | `_xlfn.CHISQ.DIST.RT(-2.5,10)`: Excel #NUM, LibreOffice #VALUE |
| CHISQ.TEST | 2 of 4 | `_xlfn.CHISQ.TEST(D!AK5:AL7,D!AM1:AN3)`: Excel 0.0168018232972, LibreOffice #VALUE |
| CHITEST | 2 of 4 | `CHITEST(D!AK5:AL7,D!AM1:AN3)`: Excel 0.0168018232972, LibreOffice #VALUE |
| CHOOSEROWS | 2 of 19 | `INDEX(_xlfn.CHOOSEROWS(D!L1:N6,1,"3"),1,1)`: Excel 1, LibreOffice #VALUE |
| COLUMN | 2 of 4 | `COLUMN(D!L1:N6)`: Excel #SPILL, LibreOffice 12 |
| CONCATENATE | 2 of 15 | `CONCATENATE(1.5,TRUE)`: Excel "1.5TRUE", LibreOffice "1.51" |
| COTH | 2 of 13 | `_xlfn.COTH(D!A1)`: Excel #DIV0, LibreOffice #NUM |
| COUNT | 2 of 12 | `COUNT(D!D1:D8)`: Excel 3, LibreOffice 5 |
| COUNTIF | 2 of 15 | `COUNTIF(D!D1:D8,2)`: Excel 1, LibreOffice 0 |
| COUNTIFS | 2 of 16 | `COUNTIFS(D!D1:D8,2)`: Excel 1, LibreOffice 0 |
| CSCH | 2 of 13 | `_xlfn.CSCH(D!A1)`: Excel #DIV0, LibreOffice #NUM |
| DAYS | 2 of 15 | `_xlfn.DAYS(-1,44228)`: Excel #NUM, LibreOffice -44229 |
| DOLLARFR | 2 of 15 | `DOLLARFR(TRUE,16)`: Excel #VALUE, LibreOffice 1 |
| DSTDEVP | 2 of 10 | `DSTDEVP(D!U1:Y7,"Yield",D!AF1:AF2)`: Excel #DIV0, LibreOffice #NUM |
| DVARP | 2 of 10 | `DVARP(D!U1:Y7,"Yield",D!AF1:AF2)`: Excel #DIV0, LibreOffice #NUM |
| EXPON.DIST | 2 of 15 | `_xlfn.EXPON.DIST(-2.5,10,TRUE)`: Excel #NUM, LibreOffice 0 |
| FLOOR | 2 of 15 | `FLOOR(2.5,-2)`: Excel #NUM, LibreOffice #VALUE |
| FREQUENCY | 2 of 8 | `INDEX(FREQUENCY(D!D1:D8,D!AT1:AT3),1,1)`: Excel 0, LibreOffice 1 |
| GROWTH | 2 of 10 | `INDEX(GROWTH(D!BB1:BB6,D!AZ1:AZ5),1,1)`: Excel #REF, LibreOffice #VALUE |
| IPMT | 2 of 15 | `IPMT(0.00833333333333333,1,D!A1,8000)`: Excel #NUM, LibreOffice #VALUE |
| KURT | 2 of 12 | `KURT(D!D1:D8)`: Excel #DIV0, LibreOffice -1.44605191996 |
| LOGNORM.DIST | 2 of 15 | `_xlfn.LOGNORM.DIST(D!A1,3.5,1.2,TRUE)`: Excel #NUM, LibreOffice 0 |
| MATCH | 2 of 15 | `MATCH(5,D!L1:L6,D!A7)`: Excel #REF, LibreOffice #DIV0 |
| MAX | 2 of 12 | `MAX("2",1)`: Excel 2, LibreOffice #VALUE |
| MINIFS | 2 of 16 | `_xlfn.MINIFS(D!B1:B8,D!D1:D8,2)`: Excel 8, LibreOffice 0 |
| MIRR | 2 of 15 | `MIRR(D!I13:I17,0.1,0.12)`: Excel #NA, LibreOffice 0.138383084319 |
| NPV | 2 of 15 | `NPV(0.1,D!D1:D8)`: Excel 7.67092411721, LibreOffice 7.6653358253 |
| OR | 2 of 13 | `OR("2",1)`: Excel TRUE, LibreOffice #VALUE |
| PERMUTATIONA | 2 of 15 | `_xlfn.PERMUTATIONA(-2.5,2)`: Excel #NUM, LibreOffice #VALUE |
| PRODUCT | 2 of 12 | `PRODUCT(D!D1:D8)`: Excel 17.5, LibreOffice 0 |
| QUOTIENT | 2 of 15 | `QUOTIENT(TRUE,2)`: Excel #VALUE, LibreOffice 0 |
| RIGHT | 2 of 15 | `RIGHT(TRUE,5)`: Excel "TRUE", LibreOffice "1" |
| SKEW | 2 of 12 | `SKEW(D!D1:D8)`: Excel -0.722108645721, LibreOffice 0.710747415234 |
| SKEW.P | 2 of 12 | `_xlfn.SKEW.P(D!D1:D8)`: Excel -0.294799620145, LibreOffice 0.476783860588 |
| SQRT | 2 of 14 | `SQRT(-16)`: Excel #NUM, LibreOffice #VALUE |
| SQRTPI | 2 of 14 | `SQRTPI(TRUE)`: Excel #VALUE, LibreOffice 1.77245385091 |
| STDEV.P | 2 of 12 | `_xlfn.STDEV.P(D!D1:D8)`: Excel 1.64991582277, LibreOffice 1.8547236991 |
| STDEVP | 2 of 12 | `STDEVP(D!D1:D8)`: Excel 1.64991582277, LibreOffice 1.8547236991 |
| SUMIF | 2 of 15 | `SUMIF(D!D1:D8,2,D!B1:B8)`: Excel 8, LibreOffice 0 |
| SUMIFS | 2 of 16 | `SUMIFS(D!B1:B8,D!D1:D8,2)`: Excel 8, LibreOffice 0 |
| SUMPRODUCT | 2 of 10 | `SUMPRODUCT(D!D1:D8,D!C1:C8)`: Excel 233.5, LibreOffice 252.5 |
| SWITCH | 2 of 15 | `_xlfn.SWITCH(2,D!A7,"one",2,"two","other")`: Excel #DIV0, LibreOffice "two" |
| UPPER | 2 of 12 | `UPPER("ß")`: Excel "ß", LibreOffice "ẞ" |
| VAR.P | 2 of 12 | `_xlfn.VAR.P(D!D1:D8)`: Excel 2.72222222222, LibreOffice 3.44 |
| VARP | 2 of 12 | `VARP(D!D1:D8)`: Excel 2.72222222222, LibreOffice 3.44 |
| VLOOKUP | 2 of 15 | `VLOOKUP(TRUE,D!L1:N6,2,FALSE)`: Excel #NA, LibreOffice "one" |
| WEIBULL.DIST | 2 of 15 | `_xlfn.WEIBULL.DIST(-2.5,20,100,TRUE)`: Excel #NUM, LibreOffice #VALUE |
| XOR | 2 of 13 | `_xlfn.XOR("2",1)`: Excel TRUE, LibreOffice #VALUE |
| Z.TEST | 2 of 15 | `_xlfn.Z.TEST(D!D1:D8,4)`: Excel 0.762474737973, LibreOffice 0.979760271943 |
| ZTEST | 2 of 15 | `ZTEST(D!D1:D8,4)`: Excel 0.762474737973, LibreOffice 0.979760271943 |
| AGGREGATE | 1 of 15 | `_xlfn.AGGREGATE(D!A7,4,D!B1:B8)`: Excel #DIV0, LibreOffice #VALUE |
| BETA.DIST | 1 of 15 | `_xlfn.BETA.DIST(D!A1,8,10,TRUE,1,3)`: Excel #NUM, LibreOffice #VALUE |
| CLEAN | 1 of 11 | `CLEAN(TRUE)`: Excel "TRUE", LibreOffice "1" |
| COMPLEX | 1 of 15 | `COMPLEX(TRUE,4)`: Excel #VALUE, LibreOffice "1+4i" |
| COS | 1 of 13 | `COS(10000000000)`: Excel #NUM, LibreOffice 0.873119622677 |
| DATEVALUE | 1 of 13 | `DATEVALUE("2011/02/23")`: Excel 40597, LibreOffice #VALUE |
| DCOUNT | 1 of 10 | `DCOUNT(D!U1:Y7,0,D!AC1:AC2)`: Excel #VALUE, LibreOffice 2 |
| DCOUNTA | 1 of 10 | `DCOUNTA(D!U1:Y7,0,D!AC1:AC2)`: Excel #VALUE, LibreOffice 2 |
| DELTA | 1 of 15 | `DELTA(TRUE,4)`: Excel #VALUE, LibreOffice 0 |
| DROP | 1 of 16 | `INDEX(_xlfn.DROP(D!L1:N6,100),1,1)`: Excel #CALC, LibreOffice 1 |
| ERF | 1 of 14 | `ERF(TRUE)`: Excel #VALUE, LibreOffice 0.84270079295 |
| ERFC | 1 of 13 | `ERFC(TRUE)`: Excel #VALUE, LibreOffice 0.15729920705 |
| FVSCHEDULE | 1 of 13 | `FVSCHEDULE(TRUE,D!AQ1:AQ3)`: Excel #VALUE, LibreOffice 1.33089 |
| GAMMADIST | 1 of 15 | `GAMMADIST(-2.5,9,2,FALSE)`: Excel #NUM, LibreOffice 0 |
| GESTEP | 1 of 15 | `GESTEP(TRUE,4)`: Excel #VALUE, LibreOffice 0 |
| HLOOKUP | 1 of 15 | `HLOOKUP(TRUE,D!L8:Q10,2,FALSE)`: Excel #NA, LibreOffice "one" |
| HOUR | 1 of 12 | `HOUR(-0.5)`: Excel #NUM, LibreOffice 12 |
| HSTACK | 1 of 7 | `INDEX(_xlfn.HSTACK(D!P1:Q2,D!D1:D8),5,3)`: Excel 0, LibreOffice "" |
| ISNUMBER | 1 of 9 | `ISNUMBER(TRUE)`: Excel FALSE, LibreOffice TRUE |
| ISOWEEKNUM | 1 of 15 | `_xlfn.ISOWEEKNUM(-1)`: Excel #NUM, LibreOffice 52 |
| LEFT | 1 of 15 | `LEFT(TRUE,4)`: Excel "TRUE", LibreOffice "1" |
| LEN | 1 of 12 | `LEN(TRUE)`: Excel 4, LibreOffice 1 |
| LOOKUP | 1 of 14 | `LOOKUP(TRUE,D!L1:L6,D!N1:N6)`: Excel #NA, LibreOffice 10 |
| LOWER | 1 of 11 | `LOWER(TRUE)`: Excel "true", LibreOffice "1" |
| MAXIFS | 1 of 16 | `_xlfn.MAXIFS(D!B1:B8,D!D1:D8,2)`: Excel 8, LibreOffice 0 |
| MID | 1 of 15 | `MID(TRUE,1,5)`: Excel "TRUE", LibreOffice "1" |
| MINUTE | 1 of 11 | `MINUTE(-0.5)`: Excel #NUM, LibreOffice 0 |
| NEGBINOM.DIST | 1 of 15 | `_xlfn.NEGBINOM.DIST(-2.5,5,0.25,TRUE)`: Excel #NUM, LibreOffice #VALUE |
| NORM.DIST | 1 of 15 | `_xlfn.NORM.DIST(42,40,0,TRUE)`: Excel #NUM, LibreOffice #VALUE |
| POISSON.DIST | 1 of 15 | `_xlfn.POISSON.DIST(-2.5,5,TRUE)`: Excel #NUM, LibreOffice #VALUE |
| PROPER | 1 of 13 | `PROPER(TRUE)`: Excel "True", LibreOffice "1" |
| REPLACE | 1 of 15 | `REPLACE(TRUE,6,5,"*")`: Excel "TRUE*", LibreOffice "1*" |
| REPT | 1 of 15 | `REPT(TRUE,3)`: Excel "TRUETRUETRUE", LibreOffice "111" |
| ROMAN | 1 of 15 | `ROMAN(499,TRUE)`: Excel "CDXCIX", LibreOffice "LDVLIV" |
| SEARCH | 1 of 15 | `SEARCH(D!A1,"Statements",6)`: Excel 6, LibreOffice #VALUE |
| SEC | 1 of 13 | `_xlfn.SEC(10000000000)`: Excel #NUM, LibreOffice 1.14531843522 |
| SECOND | 1 of 11 | `SECOND(-0.5)`: Excel #NUM, LibreOffice 0 |
| SIN | 1 of 14 | `SIN(10000000000)`: Excel #NUM, LibreOffice -0.487506025088 |
| STANDARDIZE | 1 of 15 | `STANDARDIZE(42,40,-2.5)`: Excel #NUM, LibreOffice #VALUE |
| STDEVA | 1 of 12 | `STDEVA("abc",1)`: Excel #VALUE, LibreOffice 0.707106781187 |
| STDEVPA | 1 of 12 | `STDEVPA("abc",1)`: Excel #VALUE, LibreOffice 0.5 |
| SUBSTITUTE | 1 of 15 | `SUBSTITUTE(TRUE,"Sales","Cost")`: Excel "TRUE", LibreOffice "1" |
| SUBTOTAL | 1 of 15 | `SUBTOTAL(2,D!D1:D8)`: Excel 3, LibreOffice 5 |
| T.DIST | 1 of 15 | `_xlfn.T.DIST(60,0.75,TRUE)`: Excel #NUM, LibreOffice #VALUE |
| T.DIST.RT | 1 of 15 | `_xlfn.T.DIST.RT(1.959999998,10000000000)`: Excel 0.024997895279, LibreOffice 0.0249978929515 |
| TAKE | 1 of 12 | `INDEX(_xlfn.TAKE(D!L1:N6,0),1,1)`: Excel #CALC, LibreOffice #NA |
| TAN | 1 of 14 | `TAN(10000000000)`: Excel #NUM, LibreOffice -0.558349637811 |
| TEXT | 1 of 15 | `TEXT(TRUE,"$#,##0.00")`: Excel "TRUE", LibreOffice "$1.00" |
| TREND | 1 of 10 | `INDEX(TREND(D!AY1:AY5,D!BC1:BC6),1,1)`: Excel #REF, LibreOffice #VALUE |
| TRIM | 1 of 12 | `TRIM(TRUE)`: Excel "TRUE", LibreOffice "1" |
| UNICODE | 1 of 12 | `_xlfn.UNICODE(TRUE)`: Excel 84, LibreOffice 49 |
| UNIQUE | 1 of 10 | `ROWS(_xlfn.UNIQUE(D!D1:D8))`: Excel 8, LibreOffice 7 |
| VALUE | 1 of 15 | `VALUE(TRUE)`: Excel #VALUE, LibreOffice TRUE |
| VARA | 1 of 12 | `VARA("abc",1)`: Excel #VALUE, LibreOffice 0.5 |
| VARPA | 1 of 12 | `VARPA("abc",1)`: Excel #VALUE, LibreOffice 0.25 |
| VSTACK | 1 of 7 | `INDEX(_xlfn.VSTACK(D!L1:L6,D!D1:D8),11,1)`: Excel 0, LibreOffice "" |
| WEEKDAY | 1 of 15 | `WEEKDAY(-1)`: Excel #NUM, LibreOffice 6 |
| WORKDAY | 1 of 15 | `WORKDAY(-1,151)`: Excel #NUM, LibreOffice 212 |
<!-- /auto:libreoffice -->

LibreOffice is no oracle for edge cases: it differs from Excel in 1,337 of 5,275 cases, 772 of them by error code (569 times `#VALUE!` where Excel gives `#NUM!`). LibreOffice does not know twelve functions with the `_xlfn.` prefix (BINOM.DIST.RANGE, ISO.CEILING, NETWORKDAYS.INTL, WORKDAY.INTL and eight IM*) and answers `#NAME?`. On typical cases it differs 56 times for 36 functions; besides those twelve, these are the precision of BESSELJ and CONVERT, CHAR and CODE (another encoding), DATE near 1900, ACCRINT with the calc_method argument, DOLLAR (a minus instead of parentheses), FLOOR and GAMMA (the error code), XIRR. Among the engine's whitelisted functions LibreOffice differs from Excel where it did in WB1 (a boolean in SUM's range, numeric text as SUM's argument, DATE near 1900), and also in SUMIF, SUMIFS and COUNTIFS on a number criterion against numeric text in the range and on a boolean criterion; ROUND matched everywhere, EOMONTH on typical cases, TEXT everywhere except on a boolean.

## Excel ru-RU against en-US

<!-- auto:ru -->
Excel ru-RU (Russian interface and locale) differs from en-US (English ones) in 280 cases, for 168 functions.

| case | differ | functions | example (en-US and ru-RU) |
|---|---|---|---|
| number text with a point | 101 | ABS, ACOSH, ACOT, ACOTH, ASINH, ATAN, BETA.DIST, BETA.INV, BINOM.DIST, BINOM.INV, BITOR, CEILING.PRECISE, CHIINV, CHISQ.DIST.RT, CHISQ.INV, CHISQ.INV.RT, COMBIN, CONFIDENCE.NORM, COS, COSH, COT, COTH, CRITBINOM, CSC, CSCH, DATE, DEGREES, DISC, DOLLARDE, EFFECT, ERF, ERFC, EVEN, EXP, F.DIST, FACT, FACTDOUBLE, FISHERINV, FIXED, FLOOR.PRECISE, FVSCHEDULE, GAMMA, GAMMA.INV, GAMMALN, GAMMALN.PRECISE, GAUSS, HYPGEOM.DIST, HYPGEOMDIST, INDEX, INT, ISO.CEILING, LARGE, LN, LOG, LOG10, LOGNORM.DIST, LOGNORMDIST, MID, MOD, NORM.INV, NORM.S.INV, NORMSDIST, NORMSINV, PDURATION, PERCENTILE.EXC, PERCENTILE.INC, PERCENTRANK.EXC, PERCENTRANK.INC, PHI, POWER, PRICEDISC, QUOTIENT, RADIANS, RRI, SEC, SECH, SERIESSUM, SIN, SINH, SMALL, SQRT, SQRTPI, T.DIST, T.DIST.2T, T.INV, T.INV.2T, TAN, TANH, TBILLYIELD, TINV, TRIMMEAN, TTEST, UNICHAR, ZTEST | `ABS("1.5")`: 1.5 and 46143 |
| typical | 43 | CODE, CONCATENATE, DATEVALUE, DOLLAR, FIXED, IMCOS, IMCOSH, IMCOT, IMCSC, IMCSCH, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMTAN, NUMBERVALUE, SORT, TEXT, TIMEVALUE, VALUE | `CODE("Ж")`: 95 and 134 |
| number | 40 | CLEAN, CONCATENATE, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMCSC, IMCSCH, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN, LEFT, LOWER, MID, NUMBERVALUE, PROPER, REPLACE, REPT, RIGHT, SUBSTITUTE, TEXT, TEXTJOIN, TRIM, UPPER | `CLEAN(12.5)`: "12.5" and "12,5" |
| numeric text | 28 | BETA.INV, BETAINV, CHIINV, CHISQ.INV, CHISQ.INV.RT, CONFIDENCE.NORM, CONFIDENCE.T, DOLLAR, F.INV, F.INV.RT, FINV, FIXED, GAMMA.INV, GAMMAINV, LOGINV, LOGNORM.INV, NEGBINOM.DIST, NORM.INV, NORM.S.INV, NORMINV, NORMSINV, PERCENTILE.EXC, PERCENTILE.INC, T.INV, T.INV.2T, TINV, TRIMMEAN | `_xlfn.BETA.INV("0.5",8,10,1,3)`: 1.88468463676 and #VALUE |
| boolean | 20 | CLEAN, CODE, CONCATENATE, DOLLAR, FIXED, LEFT, LEN, LOWER, MID, NUMBERVALUE, PROPER, REPLACE, REPT, RIGHT, SUBSTITUTE, TEXT, TEXTJOIN, TRIM, UNICODE, UPPER | `CLEAN(TRUE)`: "TRUE" and "ИСТИНА" |
| m/d/yyyy date text | 17 | DAY, DAYS, DAYS360, DISC, EDATE, EOMONTH, ISOWEEKNUM, MONTH, NETWORKDAYS, NETWORKDAYS.INTL, PRICEDISC, WEEKDAY, WEEKNUM, WORKDAY, WORKDAY.INTL, YEARFRAC | `DAY("1/2/2025")`: 2 and 1 |
| text TRUE | 5 | CHISQ.DIST, IF, IFS, NOT, POISSON.DIST | `_xlfn.CHISQ.DIST(0.5,1,"TRUE")`: 0.520499877813 and #VALUE |
| negative | 4 | COMPLEX, DOLLAR, FIXED, IMPOWER | `COMPLEX(-2.5,4)`: "-2.5+4i" and "-2,5+4i" |
| blank cell | 4 | DOLLAR, FIXED, TEXT | `DOLLAR(1234.567,D!A1)`: "$1,235" and "1 235 ₽" |
| range: mixed | 3 | CONCAT, TEXTJOIN | `_xlfn.CONCAT(D!D1:D8)`: "12abcTRUE5FALSE3.5" and "12abcИСТИНА5ЛОЖЬ3,5" |
| zero | 3 | DOLLAR, FIXED | `DOLLAR(1234.567,0)`: "$1,235" and "1 235 ₽" |
| fractional | 2 | COMPLEX, FIXED | `COMPLEX(0.75,4)`: "0.75+4i" and "0,75+4i" |
| Cyrillic | 1 | CODE | `CODE("Привет Мир")`: 95 and 143 |
| case | 1 | CODE | `CODE("пРИВЕТ мИР")`: 95 and 239 |
| Cyrillic in a cell | 1 | CODE | `CODE(D!A9)`: 95 and 143 |
| boolean argument | 1 | CONCAT | `_xlfn.CONCAT(TRUE,1)`: "TRUE1" and "ИСТИНА1" |
| reference to a boolean | 1 | CONCAT | `_xlfn.CONCAT(D!A4,1)`: "TRUE1" and "ИСТИНА1" |
| negative and fractional | 1 | CONCAT | `_xlfn.CONCAT(-2.5,0.75,0)`: "-2.50.750" and "-2,50,750" |
| numeric text in a cell | 1 | FIXED | `FIXED(1234.567,D!A2)`: "1,234.57" and "1 234,57" |
| typical, element 2,1 | 1 | SORT | `INDEX(_xlfn._xlws.SORT(D!F1:F7),2,1)`: "Арeнда*" and "аренда" |
| typical, element 7,1 | 1 | SORT | `INDEX(_xlfn._xlws.SORT(D!F1:F7),7,1)`: "Чай" and "x*y" |
| text | 1 | TEXT | `TEXT("abc","$#,##0.00")`: "abc" and #VALUE |

<details><summary>All cases</summary>

| case | formula | en-US | ru-RU |
|---|---|---|---|
| ABS#14 | `ABS("1.5")` | 1.5 | 46143 |
| ACOSH#14 | `ACOSH("1.5")` | 0.962423650119 | 11.4326477296 |
| ACOT#13 | `_xlfn.ACOT("1.5")` | 0.588002603548 | 0.0000216717595268 |
| ACOTH#13 | `_xlfn.ACOTH("1.5")` | 0.804718956217 | 0.0000216717595335 |
| ASINH#13 | `ASINH("1.5")` | 1.19476321729 | 11.4326477298 |
| ATAN#13 | `ATAN("1.5")` | 0.982793723247 | 1.57077465504 |
| BETA.DIST#13 | `_xlfn.BETA.DIST(2,"1.5",10,TRUE,1,3)` | 0.997333151583 | 0 |
| BETA.DIST#14 | `_xlfn.BETA.DIST(2,8,"1.5",TRUE,1,3)` | 0.00971090389227 | 1 |
| BETA.INV#3 | `_xlfn.BETA.INV("0.5",8,10,1,3)` | 1.88468463676 | #VALUE |
| BETA.INV#7 | `_xlfn.BETA.INV(0.685470581,8,10,"1.5",3)` | 2.24999999997 | #NUM |
| BETAINV#2 | `BETAINV("0.5",8,10,1,3)` | 1.88468463676 | #VALUE |
| BINOM.DIST#5 | `_xlfn.BINOM.DIST(6,"1.5",0.5,FALSE)` | #NUM | 0 |
| BINOM.DIST#11 | `_xlfn.BINOM.DIST(6,10,"0.25",FALSE)` | 0.0162220001221 | #VALUE |
| BINOM.INV#5 | `_xlfn.BINOM.INV(6,0.5,"0.25")` | 2 | #VALUE |
| BITOR#4 | `_xlfn.BITOR(23,"1.5")` | #NUM | 46143 |
| CEILING.PRECISE#15 | `_xlfn.CEILING.PRECISE("1.5")` | 2 | 46143 |
| CHIINV#2 | `CHIINV("0.5",10)` | 9.34181776559 | #VALUE |
| CHIINV#4 | `CHIINV(0.050001,"1.5")` | 3.8414252859 | 46643.8151591 |
| CHIINV#14 | `CHIINV("0.25",10)` | 12.5488613969 | #VALUE |
| CHISQ.DIST#10 | `_xlfn.CHISQ.DIST(0.5,1,"TRUE")` | 0.520499877813 | #VALUE |
| CHISQ.DIST.RT#13 | `_xlfn.CHISQ.DIST.RT(18.307,"1.5")` | 0.0000188015040226 | 1 |
| CHISQ.INV#3 | `_xlfn.CHISQ.INV("0.5",1)` | 0.45493642312 | #VALUE |
| CHISQ.INV#15 | `_xlfn.CHISQ.INV("0.25",1)` | 0.101531044268 | #VALUE |
| CHISQ.INV.RT#2 | `_xlfn.CHISQ.INV.RT("0.5",10)` | 9.34181776559 | #VALUE |
| CHISQ.INV.RT#14 | `_xlfn.CHISQ.INV.RT("0.25",10)` | 12.5488613969 | #VALUE |
| CLEAN#5 | `CLEAN(12.5)` | "12.5" | "12,5" |
| CLEAN#6 | `CLEAN(TRUE)` | "TRUE" | "ИСТИНА" |
| CODE#3 | `CODE("Ж")` | 95 | 134 |
| CODE#5 | `CODE("Привет Мир")` | 95 | 143 |
| CODE#6 | `CODE("пРИВЕТ мИР")` | 95 | 239 |
| CODE#8 | `CODE(TRUE)` | 84 | 136 |
| CODE#12 | `CODE(D!A9)` | 95 | 143 |
| COMBIN#4 | `COMBIN(8,"1.5")` | 8 | #NUM |
| COMPLEX#12 | `COMPLEX(-2.5,4)` | "-2.5+4i" | "-2,5+4i" |
| COMPLEX#15 | `COMPLEX(0.75,4)` | "0.75+4i" | "0,75+4i" |
| CONCAT#3 | `_xlfn.CONCAT(D!D1:D8)` | "12abcTRUE5FALSE3.5" | "12abcИСТИНА5ЛОЖЬ3,5" |
| CONCAT#7 | `_xlfn.CONCAT(TRUE,1)` | "TRUE1" | "ИСТИНА1" |
| CONCAT#10 | `_xlfn.CONCAT(D!A4,1)` | "TRUE1" | "ИСТИНА1" |
| CONCAT#12 | `_xlfn.CONCAT(-2.5,0.75,0)` | "-2.50.750" | "-2,50,750" |
| CONCATENATE#3 | `CONCATENATE(1.5,TRUE)` | "1.5TRUE" | "1,5ИСТИНА" |
| CONCATENATE#7 | `CONCATENATE(12.5,"b")` | "12.5b" | "12,5b" |
| CONCATENATE#8 | `CONCATENATE(TRUE,"b")` | "TRUEb" | "ИСТИНАb" |
| CONCATENATE#12 | `CONCATENATE("a",12.5)` | "a12.5" | "a12,5" |
| CONFIDENCE.NORM#2 | `_xlfn.CONFIDENCE.NORM("0.5",2.5,50)` | 0.238468138102 | #VALUE |
| CONFIDENCE.NORM#8 | `_xlfn.CONFIDENCE.NORM(0.05,"1.5",50)` | 0.415771147305 | 12789.9520334 |
| CONFIDENCE.T#2 | `_xlfn.CONFIDENCE.T("0.5",1,50)` | 0.0961000040365 | #VALUE |
| COS#13 | `COS("1.5")` | 0.0707372016677 | 0.75647103573 |
| COSH#13 | `COSH("1.5")` | 2.35240961524 | #NUM |
| COT#13 | `_xlfn.COT("1.5")` | 0.0709148443027 | -1.15663544459 |
| COTH#13 | `_xlfn.COTH("1.5")` | 1.10479139298 | 1 |
| CRITBINOM#13 | `CRITBINOM(6,0.5,"0.25")` | 2 | #VALUE |
| CSC#13 | `_xlfn.CSC("1.5")` | 1.00251130425 | -1.52898840796 |
| CSCH#13 | `_xlfn.CSCH("1.5")` | 0.469642440595 | 0 |
| DATE#13 | `DATE(2025,1,"1.5")` | 45658 | 78424 |
| DATEVALUE#2 | `DATEVALUE("8/22/2011")` | 40777 | #VALUE |
| DATEVALUE#3 | `DATEVALUE("22-MAY-2011")` | 40685 | #VALUE |
| DAY#4 | `DAY("1/2/2025")` | 2 | 1 |
| DAYS#4 | `_xlfn.DAYS("1/2/2025",44228)` | 1431 | 1461 |
| DAYS360#5 | `DAYS360("1/2/2025",40908)` | -4681 | -4710 |
| DEGREES#13 | `DEGREES("1.5")` | 85.9436692696 | 2643799.15407 |
| DISC#4 | `DISC("1/2/2025",54058,97.975,100,1)` | 0.000880618228361 | 0.000883774943243 |
| DISC#12 | `DISC(43282,54058,"1.5",100,1)` | 0.0333870818066 | -15.6065117525 |
| DOLLAR#1 | `DOLLAR(1234.567,2)` | "$1,234.57" | "1 234,57 ₽" |
| DOLLAR#2 | `DOLLAR(1234.567,-2)` | "$1,200" | "1 200 ₽" |
| DOLLAR#3 | `DOLLAR(-1234.567,-2)` | "($1,200)" | "-1 200 ₽" |
| DOLLAR#4 | `DOLLAR(-0.123,4)` | "($0.1230)" | "-0,1230 ₽" |
| DOLLAR#5 | `DOLLAR(99.888)` | "$99.89" | "99,89 ₽" |
| DOLLAR#6 | `DOLLAR("3",2)` | "$3.00" | "3,00 ₽" |
| DOLLAR#7 | `DOLLAR(TRUE,2)` | "$1.00" | "1,00 ₽" |
| DOLLAR#8 | `DOLLAR(1234.567,D!A1)` | "$1,235" | "1 235 ₽" |
| DOLLAR#9 | `DOLLAR(D!A1,2)` | "$0.00" | "0,00 ₽" |
| DOLLAR#13 | `DOLLAR(-2.5,2)` | "($2.50)" | "-2,50 ₽" |
| DOLLAR#14 | `DOLLAR(1234.567,0)` | "$1,235" | "1 235 ₽" |
| DOLLAR#15 | `DOLLAR(0,2)` | "$0.00" | "0,00 ₽" |
| DOLLARDE#8 | `DOLLARDE(1.02,"1.5")` | 1.02 | 1.04334351906 |
| EDATE#6 | `EDATE("1/2/2025",1)` | 45690 | 45717 |
| EFFECT#13 | `EFFECT(0.0525,"1.5")` | 0.0525 | 0.0539025306067 |
| EOMONTH#5 | `EOMONTH("1/2/2025",1)` | 45716 | 45747 |
| ERF#14 | `ERF("1.5")` | 0.966105146475 | 1 |
| ERFC#13 | `ERFC("1.5")` | 0.0338948535247 | 0 |
| EVEN#15 | `EVEN("1.5")` | 2 | 46144 |
| EXP#14 | `EXP("1.5")` | 4.48168907034 | #NUM |
| F.DIST#15 | `_xlfn.F.DIST(15.2069,"1.5",4,TRUE)` | 0.982452632418 | 0.99207324541 |
| F.INV#2 | `_xlfn.F.INV("0.5",6,4)` | 1.06166887827 | #VALUE |
| F.INV.RT#2 | `_xlfn.F.INV.RT("0.5",6,4)` | 1.06166887827 | #VALUE |
| FACT#14 | `FACT("1.5")` | 1 | #NUM |
| FACTDOUBLE#14 | `FACTDOUBLE("1.5")` | 1 | #NUM |
| FINV#2 | `FINV("0.5",6,4)` | 1.06166887827 | #VALUE |
| FISHERINV#13 | `FISHERINV("1.5")` | 0.905148253645 | 1 |
| FIXED#1 | `FIXED(1234.567,1)` | "1,234.6" | "1 234,6" |
| FIXED#2 | `FIXED(1234.567,-1)` | "1,230" | "1 230" |
| FIXED#4 | `FIXED(44.332)` | "44.33" | "44,33" |
| FIXED#5 | `FIXED("3",1)` | "3.0" | "3,0" |
| FIXED#6 | `FIXED(TRUE,1)` | "1.0" | "1,0" |
| FIXED#7 | `FIXED(1234.567,D!A2)` | "1,234.57" | "1 234,57" |
| FIXED#8 | `FIXED(D!A1,1)` | "0.0" | "0,0" |
| FIXED#10 | `FIXED(1234.567,"3")` | "1,234.567" | "1 234,567" |
| FIXED#12 | `FIXED(-2.5,1)` | "-2.5" | "-2,5" |
| FIXED#13 | `FIXED(1234.567,"1.5")` | "1,234.6" | #VALUE |
| FIXED#14 | `FIXED(0,1)` | "0.0" | "0,0" |
| FIXED#15 | `FIXED(0.75,1)` | "0.8" | "0,8" |
| FLOOR.PRECISE#13 | `_xlfn.FLOOR.PRECISE(-3.2,"1.5")` | -4.5 | -46143 |
| FVSCHEDULE#13 | `FVSCHEDULE("1.5",D!AQ1:AQ3)` | 1.996335 | 61411.25727 |
| GAMMA#15 | `_xlfn.GAMMA("1.5")` | 0.886226925453 | #NUM |
| GAMMA.INV#2 | `_xlfn.GAMMA.INV("0.5",9,2)` | 17.3379023687 | #VALUE |
| GAMMA.INV#13 | `_xlfn.GAMMA.INV(0.068094,9,"1.5")` | 7.50000839358 | 230715.258203 |
| GAMMAINV#2 | `GAMMAINV("0.5",9,2)` | 17.3379023687 | #VALUE |
| GAMMALN#13 | `GAMMALN("1.5")` | -0.120782237635 | 449405.323028 |
| GAMMALN.PRECISE#13 | `_xlfn.GAMMALN.PRECISE("1.5")` | -0.120782237635 | 449405.323028 |
| GAUSS#13 | `_xlfn.GAUSS("1.5")` | 0.433192798731 | 0.5 |
| HYPGEOM.DIST#7 | `_xlfn.HYPGEOM.DIST(1,4,8,"1.5",TRUE)` | #NUM | 0.999999842216 |
| HYPGEOM.DIST#11 | `_xlfn.HYPGEOM.DIST(1,"1.5",8,20,TRUE)` | 1 | #NUM |
| HYPGEOMDIST#9 | `HYPGEOMDIST(1,"1.5",8,20)` | 0.4 | #NUM |
| IF#8 | `IF("TRUE","yes","no")` | "yes" | #VALUE |
| IFS#8 | `_xlfn.IFS("TRUE",1,TRUE,2)` | 1 | #VALUE |
| IMCONJUGATE#4 | `IMCONJUGATE(12.5)` | "12.5" | "12,5" |
| IMCOS#1 | `IMCOS("1+i")` | "0.833730025131149-0.988897705762865i" | "0,833730025131149-0,988897705762865i" |
| IMCOS#4 | `IMCOS(12.5)` | "0.997798279178581" | "0,997798279178581" |
| IMCOSH#1 | `_xlfn.IMCOSH("4+3i")` | "-27.0349456030742+3.85115333481178i" | "-27,0349456030742+3,85115333481178i" |
| IMCOSH#4 | `_xlfn.IMCOSH(12.5)` | "134168.643262301" | "134168,643262301" |
| IMCOT#1 | `_xlfn.IMCOT("4+3i")` | "0.00490118239430447-0.999266927805902i" | "0,00490118239430447-0,999266927805902i" |
| IMCOT#4 | `_xlfn.IMCOT(12.5)` | "-15.0447788593086" | "-15,0447788593086" |
| IMCSC#1 | `_xlfn.IMCSC("4+3i")` | "-0.0754898329158637+0.0648774713706355i" | "-0,0754898329158637+0,0648774713706355i" |
| IMCSC#4 | `_xlfn.IMCSC(12.5)` | "-15.0779763537916" | "-15,0779763537916" |
| IMCSCH#1 | `_xlfn.IMCSCH("4+3i")` | "-0.036275889628626-0.0051744731840194i" | "-0,036275889628626-0,0051744731840194i" |
| IMCSCH#4 | `_xlfn.IMCSCH(12.5)` | "7.45330634426085E-06" | "7,45330634426085E-06" |
| IMDIV#5 | `IMDIV(12.5,"10+24i")` | "0.18491124260355-0.443786982248521i" | "0,18491124260355-0,443786982248521i" |
| IMEXP#1 | `IMEXP("1+i")` | "1.46869393991589+2.28735528717884i" | "1,46869393991589+2,28735528717884i" |
| IMEXP#4 | `IMEXP(12.5)` | "268337.286520874" | "268337,286520874" |
| IMLN#1 | `IMLN("3+4i")` | "1.6094379124341+0.927295218001612i" | "1,6094379124341+0,927295218001612i" |
| IMLN#4 | `IMLN(12.5)` | "2.52572864430826" | "2,52572864430826" |
| IMLOG10#1 | `IMLOG10("3+4i")` | "0.698970004336019+0.402719196273373i" | "0,698970004336019+0,402719196273373i" |
| IMLOG10#4 | `IMLOG10(12.5)` | "1.09691001300806" | "1,09691001300806" |
| IMLOG2#1 | `IMLOG2("3+4i")` | "2.32192809488736+1.33780421245098i" | "2,32192809488736+1,33780421245098i" |
| IMLOG2#4 | `IMLOG2(12.5)` | "3.64385618977473" | "3,64385618977473" |
| IMPOWER#1 | `IMPOWER("2+3i",3)` | "-46+9.00000000000001i" | "-46+9,00000000000001i" |
| IMPOWER#4 | `IMPOWER("2+3i",-2.5)` | "-0.0313823701431344-0.0256178001612619i" | "-0,0313823701431344-0,0256178001612619i" |
| IMPOWER#5 | `IMPOWER(12.5,3)` | "1953.125" | "1953,125" |
| IMPRODUCT#6 | `IMPRODUCT(12.5,"5-3i")` | "62.5-37.5i" | "62,5-37,5i" |
| IMPRODUCT#14 | `IMPRODUCT("3+4i",12.5)` | "37.5+50i" | "37,5+50i" |
| IMSEC#1 | `_xlfn.IMSEC("4+3i")` | "-0.0652940278579471-0.0752249603027732i" | "-0,0652940278579471-0,0752249603027732i" |
| IMSEC#4 | `_xlfn.IMSEC(12.5)` | "1.00220657909255" | "1,00220657909255" |
| IMSECH#1 | `_xlfn.IMSECH("4+3i")` | "-0.0362534969158689-0.00516434460775318i" | "-0,0362534969158689-0,00516434460775318i" |
| IMSECH#4 | `_xlfn.IMSECH(12.5)` | "7.45330634405383E-06" | "7,45330634405383E-06" |
| IMSIN#1 | `IMSIN("4+3i")` | "-7.61923172032141-6.548120040911i" | "-7,61923172032141-6,548120040911i" |
| IMSIN#4 | `IMSIN(12.5)` | "-0.0663218973512007" | "-0,0663218973512007" |
| IMSINH#1 | `_xlfn.IMSINH("4+3i")` | "-27.0168132580039+3.85373803791938i" | "-27,0168132580039+3,85373803791938i" |
| IMSINH#4 | `_xlfn.IMSINH(12.5)` | "134168.643258574" | "134168,643258574" |
| IMSQRT#1 | `IMSQRT("1+i")` | "1.09868411346781+0.455089860562227i" | "1,09868411346781+0,455089860562227i" |
| IMSQRT#4 | `IMSQRT(12.5)` | "3.53553390593274" | "3,53553390593274" |
| IMSUB#5 | `IMSUB(12.5,"5+3i")` | "7.5-3i" | "7,5-3i" |
| IMSUM#5 | `IMSUM(12.5,"5-3i")` | "17.5-3i" | "17,5-3i" |
| IMSUM#15 | `IMSUM("3+4i",12.5)` | "15.5+4i" | "15,5+4i" |
| IMTAN#1 | `_xlfn.IMTAN("4+3i")` | "0.00490825806749606+1.00070953606723i" | "0,00490825806749606+1,00070953606723i" |
| IMTAN#4 | `_xlfn.IMTAN(12.5)` | "-0.0664682418632742" | "-0,0664682418632742" |
| INDEX#10 | `INDEX(D!L1:N6,2,"1.5")` | 2 | #REF |
| INT#14 | `INT("1.5")` | 1 | 46143 |
| ISO.CEILING#15 | `_xlfn.ISO.CEILING("1.5")` | 2 | 46143 |
| ISOWEEKNUM#5 | `_xlfn.ISOWEEKNUM("1/2/2025")` | 1 | 5 |
| LARGE#9 | `LARGE(D!B1:B8,"1.5")` | 23 | #NUM |
| LEFT#8 | `LEFT(12.5,4)` | "12.5" | "12,5" |
| LEFT#9 | `LEFT(TRUE,4)` | "TRUE" | "ИСТИ" |
| LEN#8 | `LEN(TRUE)` | 4 | 6 |
| LN#14 | `LN("1.5")` | 0.405465108108 | 10.7395005491 |
| LOG#15 | `LOG("1.5")` | 0.176091259056 | 4.66410582688 |
| LOG10#14 | `LOG10("1.5")` | 0.176091259056 | 4.66410582688 |
| LOGINV#2 | `LOGINV("0.5",3.5,1.2)` | 33.1154519587 | #VALUE |
| LOGNORM.DIST#10 | `_xlfn.LOGNORM.DIST(4,"1.5",1.2,TRUE)` | 0.462254834906 | 0 |
| LOGNORM.INV#2 | `_xlfn.LOGNORM.INV("0.5",3.5,1.2)` | 33.1154519587 | #VALUE |
| LOGNORMDIST#5 | `LOGNORMDIST(4,3.5,"1.5")` | 0.0793973173314 | 0.499981725364 |
| LOWER#5 | `LOWER(12.5)` | "12.5" | "12,5" |
| LOWER#6 | `LOWER(TRUE)` | "true" | "истина" |
| MID#9 | `MID(12.5,1,5)` | "12.5" | "12,5" |
| MID#10 | `MID(TRUE,1,5)` | "TRUE" | "ИСТИН" |
| MID#11 | `MID("Fluid Flow","1.5",5)` | "Fluid" | "" |
| MOD#13 | `MOD(3,"1.5")` | 0 | 3 |
| MONTH#4 | `MONTH("1/2/2025")` | 1 | 2 |
| NEGBINOM.DIST#6 | `_xlfn.NEGBINOM.DIST(10,5,"0.5",TRUE)` | 0.940765380859 | #VALUE |
| NETWORKDAYS#5 | `NETWORKDAYS("1/2/2025",41334)` | -3090 | -3111 |
| NETWORKDAYS.INTL#5 | `_xlfn.NETWORKDAYS.INTL("1/2/2025",38748)` | -4938 | -4959 |
| NETWORKDAYS.INTL#12 | `_xlfn.NETWORKDAYS.INTL(38718,"1/2/2025")` | 4959 | 4980 |
| NORM.INV#2 | `_xlfn.NORM.INV("0.5",40,1.5)` | 40 | #VALUE |
| NORM.INV#4 | `_xlfn.NORM.INV(0.908789,"1.5",1.5)` | 3.50000200957 | 46145.000002 |
| NORM.S.INV#2 | `_xlfn.NORM.S.INV("0.5")` | 0 | #VALUE |
| NORM.S.INV#10 | `_xlfn.NORM.S.INV("0.25")` | -0.674489750196 | #VALUE |
| NORMINV#2 | `NORMINV("0.5",40,1.5)` | 40 | #VALUE |
| NORMSDIST#13 | `NORMSDIST("1.5")` | 0.933192798731 | 1 |
| NORMSINV#2 | `NORMSINV("0.5")` | 0 | #VALUE |
| NORMSINV#10 | `NORMSINV("0.25")` | -0.674489750196 | #VALUE |
| NOT#5 | `NOT("TRUE")` | FALSE | #VALUE |
| NUMBERVALUE#2 | `_xlfn.NUMBERVALUE("3.5%")` | 0.035 | 461.45 |
| NUMBERVALUE#3 | `_xlfn.NUMBERVALUE("1 234.5")` | 1234.5 | #VALUE |
| NUMBERVALUE#5 | `_xlfn.NUMBERVALUE("2.5")` | 2.5 | 46144 |
| NUMBERVALUE#8 | `_xlfn.NUMBERVALUE("2.500,27",TRUE,".")` | #VALUE | 2500.27 |
| NUMBERVALUE#10 | `_xlfn.NUMBERVALUE(12.5,",",".")` | 125 | 12.5 |
| PDURATION#13 | `_xlfn.PDURATION(0.025,"1.5",2200)` | 295.260272867 | -123.246898182 |
| PERCENTILE.EXC#10 | `_xlfn.PERCENTILE.EXC(D!B1:B8,"0.25")` | 5 | #VALUE |
| PERCENTILE.EXC#11 | `_xlfn.PERCENTILE.EXC(D!B1:B8,"0.5")` | 12.625 | #VALUE |
| PERCENTILE.INC#13 | `_xlfn.PERCENTILE.INC(D!B1:B8,"0.25")` | 7 | #VALUE |
| PERCENTILE.INC#15 | `_xlfn.PERCENTILE.INC(D!B1:B8,"0.5")` | 12.625 | #VALUE |
| PERCENTRANK.EXC#10 | `_xlfn.PERCENTRANK.EXC(D!B1:B8,"1.5")` | 0.185 | #NA |
| PERCENTRANK.INC#6 | `_xlfn.PERCENTRANK.INC(D!B1:B8,"1.5")` | 0.095 | #NA |
| PHI#12 | `_xlfn.PHI("1.5")` | 0.129517595666 | 0 |
| POISSON.DIST#6 | `_xlfn.POISSON.DIST(2,5,"TRUE")` | 0.124652019483 | #VALUE |
| POWER#9 | `POWER(5,"1.5")` | 11.1803398875 | #NUM |
| PRICEDISC#4 | `PRICEDISC(39494,"1/2/2025",0.0525,100,2)` | 10.09375 | 9.65625 |
| PRICEDISC#11 | `PRICEDISC(39494,39508,"1.5",100,2)` | 94.1666666667 | -179345 |
| PROPER#7 | `PROPER(12.5)` | "12.5" | "12,5" |
| PROPER#8 | `PROPER(TRUE)` | "True" | "Истина" |
| QUOTIENT#12 | `QUOTIENT(5,"1.5")` | 3 | 0 |
| RADIANS#13 | `RADIANS("1.5")` | 0.0261799387799 | 805.347276748 |
| REPLACE#9 | `REPLACE("abcdefghijk",6,5,12.5)` | "abcde12.5k" | "abcde12,5k" |
| REPLACE#10 | `REPLACE(12.5,6,5,"*")` | "12.5*" | "12,5*" |
| REPLACE#11 | `REPLACE(TRUE,6,5,"*")` | "TRUE*" | "ИСТИН*" |
| REPT#7 | `REPT(12.5,3)` | "12.512.512.5" | "12,512,512,5" |
| REPT#8 | `REPT(TRUE,3)` | "TRUETRUETRUE" | "ИСТИНАИСТИНАИСТИНА" |
| RIGHT#7 | `RIGHT(12.5,5)` | "12.5" | "12,5" |
| RIGHT#8 | `RIGHT(TRUE,5)` | "TRUE" | "СТИНА" |
| RRI#9 | `_xlfn.RRI(96,10000,"1.5")` | -0.0876370993397 | 0.0160562906913 |
| SEC#13 | `_xlfn.SEC("1.5")` | 14.136832903 | 1.32192767835 |
| SECH#13 | `_xlfn.SECH("1.5")` | 0.425096034942 | 0 |
| SERIESSUM#13 | `SERIESSUM(0.785398163,0,"1.5",D!AR1:AR4)` | 0.671697526921 | 1 |
| SIN#14 | `SIN("1.5")` | 0.997494986604 | -0.654027195231 |
| SINH#14 | `SINH("1.5")` | 2.12927945509 | #NUM |
| SMALL#15 | `SMALL(D!B1:B8,"1.5")` | -3.5 | #NUM |
| SORT#9 | `INDEX(_xlfn._xlws.SORT(D!F1:F7),1,1)` | "x*y" | "Аренда" |
| SORT#10 | `INDEX(_xlfn._xlws.SORT(D!F1:F7),2,1)` | "Арeнда*" | "аренда" |
| SORT#11 | `INDEX(_xlfn._xlws.SORT(D!F1:F7),7,1)` | "Чай" | "x*y" |
| SQRT#14 | `SQRT("1.5")` | 1.22474487139 | 214.809217679 |
| SQRTPI#14 | `SQRTPI("1.5")` | 2.17080376367 | 380.739425086 |
| SUBSTITUTE#10 | `SUBSTITUTE(12.5,"Sales","Cost")` | "12.5" | "12,5" |
| SUBSTITUTE#11 | `SUBSTITUTE(TRUE,"Sales","Cost")` | "TRUE" | "ИСТИНА" |
| T.DIST#5 | `_xlfn.T.DIST(60,"1.5",TRUE)` | 0.994695326367 | 1 |
| T.DIST.2T#7 | `_xlfn.T.DIST.2T(1.959999998,"1.5")` | 0.300342892039 | 0.0500017999147 |
| T.INV#2 | `_xlfn.T.INV("0.5",2)` | 0 | #VALUE |
| T.INV#13 | `_xlfn.T.INV(0.75,"1.5")` | 1 | 0.674495067075 |
| T.INV#14 | `_xlfn.T.INV("0.25",2)` | -0.816496580928 | #VALUE |
| T.INV.2T#2 | `_xlfn.T.INV.2T("0.5",60)` | 0.678600720648 | #VALUE |
| T.INV.2T#7 | `_xlfn.T.INV.2T(0.546449,"1.5")` | 0.863773679737 | 0.603094183201 |
| T.INV.2T#14 | `_xlfn.T.INV.2T("0.25",60)` | 1.16159552688 | #VALUE |
| TAN#14 | `TAN("1.5")` | 14.1014199472 | -0.864576651767 |
| TANH#14 | `TANH("1.5")` | 0.905148253645 | 1 |
| TBILLYIELD#5 | `TBILLYIELD(39538,39600,"1.5")` | 381.290322581 | -5.7938680106 |
| TEXT#1 | `TEXT(1234.567,"$#,##0.00")` | "$1,234.57" | #VALUE |
| TEXT#2 | `TEXT(45658,"YYYY-MM-DD")` | "2025-01-01" | "YYYY-MM-DD" |
| TEXT#3 | `TEXT(0.285,"0.0%")` | "28.5%" | #VALUE |
| TEXT#4 | `TEXT(45658,"dd.mm.yyyy")` | "01.01.2025" | "dd.mm.yyyy" |
| TEXT#5 | `TEXT(1234.5,"0.00")` | "1234.50" | #VALUE |
| TEXT#7 | `TEXT(45688.75,"h:mm AM/PM")` | "6:00 PM" | "h:mm PM" |
| TEXT#8 | `TEXT(4.34,"# ?/?")` | "4 1/3" | #VALUE |
| TEXT#9 | `TEXT(2.5,"$#,##0.00")` | "$2.50" | #VALUE |
| TEXT#10 | `TEXT("abc","$#,##0.00")` | "abc" | #VALUE |
| TEXT#12 | `TEXT(D!A1,"$#,##0.00")` | "$0.00" | #VALUE |
| TEXT#13 | `TEXT(TRUE,"$#,##0.00")` | "TRUE" | #VALUE |
| TEXTJOIN#4 | `_xlfn.TEXTJOIN(" ",TRUE,D!D1:D8)` | "1 2 abc TRUE 5 FALSE 3.5" | "1 2 abc ИСТИНА 5 ЛОЖЬ 3,5" |
| TEXTJOIN#8 | `_xlfn.TEXTJOIN(", ",TRUE,D!D1:D8)` | "1, 2, abc, TRUE, 5, FALSE, 3.5" | "1, 2, abc, ИСТИНА, 5, ЛОЖЬ, 3,5" |
| TEXTJOIN#9 | `_xlfn.TEXTJOIN(12.5,TRUE,"a","","b")` | "a12.5b" | "a12,5b" |
| TEXTJOIN#10 | `_xlfn.TEXTJOIN(TRUE,TRUE,"a","","b")` | "aTRUEb" | "aИСТИНАb" |
| TIMEVALUE#2 | `TIMEVALUE("22-Aug-2011 6:35 AM")` | 0.274305555555 | #VALUE |
| TINV#2 | `TINV("0.5",60)` | 0.678600720648 | #VALUE |
| TINV#14 | `TINV("0.25",60)` | 1.16159552688 | #VALUE |
| TRIM#6 | `TRIM(12.5)` | "12.5" | "12,5" |
| TRIM#7 | `TRIM(TRUE)` | "TRUE" | "ИСТИНА" |
| TRIMMEAN#13 | `TRIMMEAN(D!B1:B8,"0.25")` | 12.7083333333 | #VALUE |
| TRIMMEAN#15 | `TRIMMEAN(D!B1:B8,"0.5")` | 12.3125 | #VALUE |
| TTEST#14 | `TTEST(D!B1:B8,D!C1:C8,"1.5",1)` | 0.176013703593 | #NUM |
| TTEST#15 | `TTEST(D!B1:B8,D!C1:C8,2,"1.5")` | 0.352027407187 | #NUM |
| UNICHAR#15 | `_xlfn.UNICHAR("1.5")` | "\u0001" | "됿" |
| UNICODE#8 | `_xlfn.UNICODE(TRUE)` | 84 | 1048 |
| UPPER#6 | `UPPER(12.5)` | "12.5" | "12,5" |
| UPPER#7 | `UPPER(TRUE)` | "TRUE" | "ИСТИНА" |
| VALUE#1 | `VALUE("$1,000")` | 1000 | #VALUE |
| VALUE#3 | `VALUE("1.5")` | 1.5 | 46143 |
| VALUE#4 | `VALUE("1,5")` | #VALUE | 1.5 |
| WEEKDAY#7 | `WEEKDAY("1/2/2025")` | 5 | 7 |
| WEEKNUM#5 | `WEEKNUM("1/2/2025")` | 1 | 5 |
| WORKDAY#5 | `WORKDAY("1/2/2025",151)` | 45870 | 45901 |
| WORKDAY.INTL#5 | `_xlfn.WORKDAY.INTL("1/2/2025",90,11)` | 45764 | 45794 |
| YEARFRAC#7 | `YEARFRAC("1/2/2025",41120)` | 12.4222222222 | 12.5027777778 |
| ZTEST#12 | `ZTEST(D!B1:B8,"1.5")` | 0.00414427367499 | 1 |

</details>
<!-- /auto:ru -->

For an interface with Russian input the difference comes down to a few rules:
- text like `"1.5"` becomes a date in ru-RU (1 May of the current year, 46143), so any numeric text with a point turns into a number near 46,000 on the Russian side (101 cases of 94 functions); `"1/2/2025"` reads as 1 February, not 2 January (17 cases of 16 date functions);
- a number turned into text is written with a comma: CONCATENATE, CONCAT, TEXTJOIN, the complex functions (`"-27,0349456030742+3,85115333481178i"`), string functions of a number;
- DOLLAR and FIXED give roubles and a space between thousands in ru-RU (`"1 235 ₽"` instead of `"$1,235"`);
- TEXT in ru-RU does not know the English codes: `TEXT(45658,"YYYY-MM-DD")` returns `"YYYY-MM-DD"`, and `"$#,##0.00"`, `"0.00"` and `"0.0%"` give `#VALUE!`;
- NUMBERVALUE, VALUE, DATEVALUE and TIMEVALUE read text by the locale;
- SORT orders a mix of Cyrillic and Latin differently;
- from the interface language: a boolean turned into text comes out as `ИСТИНА`/`ЛОЖЬ` (20 cases: CONCATENATE, TEXTJOIN, LEFT, LEN and other string functions of TRUE), the text `"TRUE"` in a boolean argument gives `#VALUE!` (IF, IFS, NOT, CHISQ.DIST, POISSON.DIST), CODE of Cyrillic takes the Mac Cyrillic encoding.

## The interface language

Excel takes the interface language not only for display (the texts of errors) but for computation as well: it decides how a boolean becomes text, which text counts as a boolean and which encoding CODE works in. The report's first edition computed the en-US column with the en_US locale but the Russian interface; that column is in git (`docs/coverage-formulajs/excel-en-ru-ui.json` at commit `85522cf`). With the English interface `IF("TRUE",1,2)` is 1 and `IF("ИСТИНА",1,2)` is `#VALUE!` (a separate check on 5 October); with the Russian one it is the other way round. With the English interface Excel shows errors in English too. The difference on the measurement's cases (measured on 5 October, both columns with formulas of the old kind; this table is no longer rebuilt):

With the Russian interface Excel en-US's answer was different in 32 cases of 23 functions. Through the adapter 3950 cases matched the former column and 3975 the one with the English interface; bare, 3342 and 3354.

| case | formula | Russian interface | English interface |
|---|---|---|---|
| CHISQ.DIST#10 | `_xlfn.CHISQ.DIST(0.5,1,"TRUE")` | #VALUE | 0.520499877813 |
| CLEAN#6 | `CLEAN(TRUE)` | "ИСТИНА" | "TRUE" |
| CODE#3 | `CODE("Ж")` | 134 | 95 |
| CODE#5 | `CODE("Привет Мир")` | 143 | 95 |
| CODE#6 | `CODE("пРИВЕТ мИР")` | 239 | 95 |
| CODE#8 | `CODE(TRUE)` | 136 | 84 |
| CODE#12 | `CODE(D!A9)` | 143 | 95 |
| CONCAT#3 | `_xlfn.CONCAT(D!D1:D8)` | "12abcИСТИНА5ЛОЖЬ3.5" | "12abcTRUE5FALSE3.5" |
| CONCAT#7 | `_xlfn.CONCAT(TRUE,1)` | "ИСТИНА1" | "TRUE1" |
| CONCAT#10 | `_xlfn.CONCAT(D!A4,1)` | "ИСТИНА1" | "TRUE1" |
| CONCATENATE#3 | `CONCATENATE(1.5,TRUE)` | "1.5ИСТИНА" | "1.5TRUE" |
| CONCATENATE#8 | `CONCATENATE(TRUE,"b")` | "ИСТИНАb" | "TRUEb" |
| IF#8 | `IF("TRUE","yes","no")` | #VALUE | "yes" |
| IFS#8 | `_xlfn.IFS("TRUE",1,TRUE,2)` | #VALUE | 1 |
| LEFT#9 | `LEFT(TRUE,4)` | "ИСТИ" | "TRUE" |
| LEN#8 | `LEN(TRUE)` | 6 | 4 |
| LOWER#6 | `LOWER(TRUE)` | "истина" | "true" |
| MID#10 | `MID(TRUE,1,5)` | "ИСТИН" | "TRUE" |
| NOT#5 | `NOT("TRUE")` | #VALUE | FALSE |
| POISSON.DIST#6 | `_xlfn.POISSON.DIST(2,5,"TRUE")` | #VALUE | 0.124652019483 |
| PROPER#8 | `PROPER(TRUE)` | "Истина" | "True" |
| REPLACE#11 | `REPLACE(TRUE,6,5,"*")` | "ИСТИН*" | "TRUE*" |
| REPT#8 | `REPT(TRUE,3)` | "ИСТИНАИСТИНАИСТИНА" | "TRUETRUETRUE" |
| RIGHT#8 | `RIGHT(TRUE,5)` | "СТИНА" | "TRUE" |
| SUBSTITUTE#11 | `SUBSTITUTE(TRUE,"Sales","Cost")` | "ИСТИНА" | "TRUE" |
| TEXT#13 | `TEXT(TRUE,"$#,##0.00")` | "ИСТИНА" | "TRUE" |
| TEXTJOIN#4 | `_xlfn.TEXTJOIN(" ",TRUE,D!D1:D8)` | "1 2 abc ИСТИНА 5 ЛОЖЬ 3.5" | "1 2 abc TRUE 5 FALSE 3.5" |
| TEXTJOIN#8 | `_xlfn.TEXTJOIN(", ",TRUE,D!D1:D8)` | "1, 2, abc, ИСТИНА, 5, ЛОЖЬ, 3.5" | "1, 2, abc, TRUE, 5, FALSE, 3.5" |
| TEXTJOIN#10 | `_xlfn.TEXTJOIN(TRUE,TRUE,"a","","b")` | "aИСТИНАb" | "aTRUEb" |
| TRIM#7 | `TRIM(TRUE)` | "ИСТИНА" | "TRUE" |
| UNICODE#8 | `_xlfn.UNICODE(TRUE)` | 1048 | 84 |
| UPPER#7 | `UPPER(TRUE)` | "ИСТИНА" | "TRUE" |

## Excel functions formulajs lacks

<!-- auto:missing -->
115 of the reference's 521 functions. An asterisk: formulajs has the name, but as an object of subfunctions, not a function (it cannot be called).

| category | count | functions |
|---|---|---|
| Financial | 22 | ACCRINTM, AMORDEGRC, AMORLINC, COUPDAYBS, COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD, DURATION, INTRATE, MDURATION, ODDFPRICE, ODDFYIELD, ODDLPRICE, ODDLYIELD, PRICE, PRICEMAT, RECEIVED, VDB, YIELD, YIELDDISC, YIELDMAT |
| Text | 22 | ARRAYTOTEXT, ASC, BAHTTEXT, DBCS, DETECTLANGUAGE, FINDB, JIS, LEFTB, LENB, MIDB, PHONETIC, REGEXEXTRACT, REGEXREPLACE, REGEXTEST, REPLACEB, RIGHTB, SEARCHB, TEXTAFTER, TEXTBEFORE, TEXTSPLIT, TRANSLATE, VALUETOTEXT |
| Lookup and reference | 20 | ADDRESS, AREAS, FILTER, FORMULATEXT, GETPIVOTDATA, GROUPBY, HYPERLINK, IMAGE, INDIRECT, OFFSET, PIVOTBY, RTD, SORTBY, TOCOL, TOROW, TRIMRANGE, WRAPCOLS, WRAPROWS, XLOOKUP, XMATCH |
| Compatibility | 10 | CONFIDENCE*, MODE*, PERCENTILE*, PERCENTRANK*, POISSON*, QUARTILE*, RANK*, STDEV*, VAR*, WEIBULL* |
| Logical | 8 | BYCOL, BYROW, LAMBDA, LET, MAKEARRAY, MAP, REDUCE, SCAN |
| Information | 8 | CELL, INFO, ISFORMULA, ISOMITTED, ISREF, SHEET, SHEETS, STOCKHISTORY |
| Cube | 7 | CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY, CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE |
| Statistical | 5 | FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY, FORECAST.ETS.STAT, FORECAST.LINEAR |
| Math and trigonometry | 5 | MDETERM, MINVERSE, PERCENTOF, RANDARRAY, SEQUENCE |
| Add-in and Automation | 3 | CALL, EUROCONVERT, REGISTER.ID |
| Web | 3 | ENCODEURL, FILTERXML, WEBSERVICE |
| Engineering | 2 | ERF.PRECISE, ERFC.PRECISE |
<!-- /auto:missing -->

Of what the engine needs, these are XLOOKUP and XMATCH, OFFSET and INDIRECT, SEQUENCE, FILTER and SORTBY, LET and LAMBDA. Ten old statistical names (STDEV, VAR, MODE, RANK and others) cannot be called in formulajs: under these names it has an object with the subfunctions `.S`, `.P` and so on.

formulajs names Excel lacks (not measured):

<!-- auto:notinexcel -->
| formulajs name | same implementation as the Excel function |
|---|---|
| BINOMDIST.RANGE | BINOM.DIST.RANGE |
| CEILINGMATH | CEILING.MATH |
| CEILINGPRECISE | CEILING.PRECISE |
| CHIDIST.RT | CHISQ.DIST.RT |
| CHIDISTRT | CHISQ.DIST.RT |
| CHIINV.RT | CHISQ.INV.RT |
| CHIINVRT | CHISQ.INV.RT |
| COVARIANCEP | COVAR, COVARIANCE.P |
| COVARIANCES | COVARIANCE.S |
| FDIST.RT | F.DIST.RT |
| FDISTRT | F.DIST.RT |
| FINV.RT | F.INV.RT |
| FINVRT | F.INV.RT |
| FLOORMATH | FLOOR.MATH |
| FLOORPRECISE | FLOOR.PRECISE |
| GAMMALNPRECISE | GAMMALN.PRECISE |
| ISO.CEILING.MATH | CEILING.MATH |
| ISO.CEILING.PRECISE | CEILING.PRECISE |
| LOGNORMINV | LOGINV, LOGNORM.INV |
| MODEMULT | MODE.MULT |
| MODESNGL | MODE.SNGL |
| NETWORKDAYSINTL | NETWORKDAYS.INTL |
| PERCENTILEEXC | PERCENTILE.EXC |
| PERCENTILEINC | PERCENTILE.INC |
| PERCENTRANKEXC | PERCENTRANK.EXC |
| PERCENTRANKINC | PERCENTRANK.INC |
| POISSONDIST | POISSON.DIST |
| QUARTILEEXC | QUARTILE.EXC |
| QUARTILEINC | QUARTILE.INC |
| RANKAVG | RANK.AVG |
| RANKEQ | RANK.EQ |
| SKEWP | SKEW.P |
| STDEVS | STDEV.S |
| TDIST.2T | T.DIST.2T |
| TDIST.RT | T.DIST.RT |
| TDISTRT | T.DIST.RT |
| TINV.2T | T.INV.2T |
| VARS | VAR.S |
| WEIBULLDIST | WEIBULL.DIST |
| WORKDAYINTL | WORKDAY.INTL |
<!-- /auto:notinexcel -->

## Proposal

The classes are set automatically by what the function differs on through the adapter: "adapter" means every difference is in argument conversion, blank input, errors inside ranges, date text or the answer's format; "adapter and a domain check" means there are also differences on values outside the domain, on dates at the 1900 boundary or only in the error code; "own" means typical cases differ, or edge cases differ in meaning (wildcards, case, Cyrillic, the order of dates).

<!-- auto:plan -->
| proposal | functions | which |
|---|---|---|
| as is | 105 | BINOMDIST, CEILING, FLOOR, NORMDIST, ISPMT, NPV, PDURATION, ERROR.TYPE, ISBLANK, ISERR, ISERROR, ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISTEXT, N, NA, AND, FALSE, IF, IFS, NOT, OR, TRUE, XOR, CHOOSE, CHOOSECOLS, CHOOSEROWS, COLUMNS, DROP, EXPAND, HSTACK, ROWS, TAKE, TRANSPOSE, VSTACK, ABS, ACOS, ACOSH, ACOTH, ASIN, ASINH, ATAN, ATANH, CEILING.MATH, CEILING.PRECISE, COSH, CSCH, DEGREES, EVEN, EXP, FLOOR.MATH, FLOOR.PRECISE, INT, LN, LOG, LOG10, MMULT, MUNIT, ODD, PI, POWER, PRODUCT, RADIANS, ROUNDDOWN, ROUNDUP, SECH, SIGN, SINH, SQRT, SUBTOTAL, SUM, TRUNC, AVEDEV, AVERAGE, BINOM.DIST.RANGE, COUNTA, COUNTBLANK, F.DIST.RT, FISHER, GAMMALN, GAMMALN.PRECISE, GAUSS, MAX, MEDIAN, MIN, NORM.DIST, NORM.S.DIST, PHI, PROB, CLEAN, CONCAT, CONCATENATE, EXACT, FIND, FIXED, LEFT, LEN, LOWER, REPT, RIGHT, T, TRIM, UNICODE |
| adapter | 40 | ZTEST, DATEVALUE, HOUR, ISOWEEKNUM, MINUTE, SECOND, BESSELI, BESSELJ, BESSELK, BESSELY, COMPLEX, DELTA, ERFC, GESTEP, CUMIPMT, CUMPRINC, DDB, DOLLARDE, DOLLARFR, EFFECT, FVSCHEDULE, NOMINAL, NPER, ISEVEN, ISODD, TYPE, IFNA, MROUND, SQRTPI, SUMPRODUCT, SUMSQ, AVERAGEA, COUNT, FREQUENCY, GAMMA.INV, MAXA, MINA, T.INV, Z.TEST, TEXTJOIN |
| adapter and a domain check | 169 | BETADIST, BETAINV, COVAR, CRITBINOM, EXPONDIST, GAMMADIST, GAMMAINV, HYPGEOMDIST, LOGINV, NEGBINOMDIST, NORM.INV, NORMSINV, STDEVP, TDIST, VARP, DATEDIF, DAY, DAYS, EDATE, EOMONTH, MONTH, WEEKDAY, WEEKNUM, YEAR, YEARFRAC, BIN2DEC, BIN2HEX, BIN2OCT, BITAND, BITLSHIFT, BITOR, BITRSHIFT, BITXOR, DEC2BIN, DEC2HEX, DEC2OCT, HEX2BIN, HEX2DEC, HEX2OCT, IMABS, IMAGINARY, IMARGUMENT, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMCSC, IMCSCH, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMREAL, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN, OCT2BIN, OCT2DEC, OCT2HEX, ACCRINT, COUPDAYS, FV, IPMT, IRR, MIRR, PMT, PPMT, PRICEDISC, PV, RATE, RRI, SLN, SYD, HLOOKUP, INDEX, ACOT, ATAN2, BASE, COMBIN, COMBINA, COS, COT, COTH, CSC, FACT, FACTDOUBLE, GCD, LCM, MULTINOMIAL, QUOTIENT, ROUND, SEC, SERIESSUM, SIN, SUMX2MY2, SUMX2PY2, SUMXMY2, TAN, TANH, BETA.INV, BINOM.DIST, BINOM.INV, CHISQ.DIST, CHISQ.DIST.RT, CONFIDENCE.NORM, CORREL, COVARIANCE.P, COVARIANCE.S, DEVSQ, EXPON.DIST, F.DIST, F.INV, F.INV.RT, FISHERINV, FORECAST, GAMMA.DIST, GEOMEAN, GROWTH, HARMEAN, HYPGEOM.DIST, INTERCEPT, KURT, LARGE, LOGNORM.DIST, LOGNORM.INV, MODE.SNGL, NEGBINOM.DIST, NORM.S.INV, NORMINV, PEARSON, PERCENTILE.EXC, PERCENTILE.INC, PERCENTRANK.EXC, PERMUT, PERMUTATIONA, POISSON.DIST, QUARTILE.EXC, RANK.AVG, RANK.EQ, RSQ, SKEW, SKEW.P, SLOPE, SMALL, STANDARDIZE, STDEV.P, STDEV.S, STDEVA, STDEVPA, STEYX, T.DIST.2T, T.DIST.RT, T.INV.2T, VAR.P, VAR.S, VARA, VARPA, WEIBULL.DIST, MID, REPLACE, UNICHAR |
| own (edges differ in meaning) | 7 | XNPV, SWITCH, MATCH, VLOOKUP, DECIMAL, MINIFS, SEARCH |
| own (typical cases differ) | 78 | CHIDIST, CHITEST, FDIST, FINV, FTEST, LOGNORMDIST, NORMSDIST, TINV, TTEST, DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR, DVARP, DATE, DAYS360, NETWORKDAYS, NETWORKDAYS.INTL, TIME, TIMEVALUE, WORKDAY, WORKDAY.INTL, CONVERT, ERF, DB, DISC, TBILLEQ, TBILLPRICE, TBILLYIELD, XIRR, IFERROR, COLUMN, LOOKUP, ROW, SORT, UNIQUE, AGGREGATE, ARABIC, ISO.CEILING, MOD, ROMAN, SUMIF, SUMIFS, AVERAGEIF, AVERAGEIFS, BETA.DIST, CHISQ.TEST, CONFIDENCE.T, COUNTIF, COUNTIFS, F.TEST, GAMMA, LINEST, LOGEST, MAXIFS, MODE.MULT, PERCENTRANK.INC, QUARTILE.INC, T.DIST, T.TEST, TREND, TRIMMEAN, CHAR, CODE, DOLLAR, NUMBERVALUE, PROPER, SUBSTITUTE, TEXT, UPPER, VALUE |
| own or a guard (hangs) | 3 | CHIINV, CHISQ.INV, CHISQ.INV.RT |
<!-- /auto:plan -->

**As is, 105 functions.** Take them into the whitelist with a test against Excel, as SUM. Among them SUM, PRODUCT, AVERAGE, MIN, MAX, MEDIAN, COUNTA, ROUNDDOWN, ROUNDUP, TRUNC, INT, ABS, POWER, SQRT, LN, LOG, EXP, CHOOSE, IF, IFS, NOT, AND, OR, ISBLANK, ISERROR and the other IS* except ISEVEN and ISODD, TRIM, FIND, LEFT, RIGHT, LEN, LOWER, CONCATENATE, CONCAT, REPT, FIXED, MMULT, TRANSPOSE and the array shape functions (CHOOSECOLS, DROP, EXPAND, HSTACK, TAKE, VSTACK). After any change to the adapter they need checking again by the same measurement.

**Fixing the adapter, 39 functions.** These rules are needed by this group and the next; each comes from observed differences:
1. Conversion by group, not one for all. The Analysis ToolPak functions (engineering, some financial) do not take a boolean literal: `#VALUE!`. Folds like COUNT, AVERAGEA, MAXA, STDEVA convert numeric text and booleans given directly as arguments, but not from a reference. GCD, LCM and MULTINOMIAL do not let text and booleans from a range through. The text `"TRUE"` in a boolean argument is taken, as by the runtime's `toBool` (Excel en-US with the English interface).
2. Date text in a date argument: read it by the locale (`"2025-01-31"` and `"1/2/2025"` in en-US) before turning it into a JS Date.
3. Blank input: a blank range in statistics gives `#DIV/0!` (CORREL, SLOPE, COVARIANCE, KURT, SKEW), `#NUM!` (LARGE, SMALL, PERCENTILE, QUARTILE, GEOMEAN, DEVSQ) or `#N/A` (MODE, HARMEAN, PERCENTRANK); a formulajs answer `null` (a reference to a blank cell in IF and IFERROR) is 0, not `#VALUE!`.
4. Pairs in paired arrays (CORREL, SLOPE, INTERCEPT, PEARSON, RSQ, STEYX, COVARIANCE.*, FORECAST, SUMX2*): drop the pairs where at least one value is not a number; different sizes are `#N/A`.
5. The answer's format: hexadecimal digits in upper case (BIN2HEX, DEC2HEX, OCT2HEX), a complex number as a string of 15 significant digits (IM*).
6. Holidays and payment dates as a flat list of serial numbers, not a two-dimensional array of JS Dates.

The rules come from the differences; no run checked them: before a function goes into the whitelist, its rule is checked by the same measurement.

**Adapter and a domain check, 169 functions.** The same rules plus an argument check before the call and the error code for wrong input: the domains of distributions (a probability in (0,1), positive parameters, whole numbers where Excel requires one), trigonometry of large arguments (1E10 in the measurement), whole-number arguments of the bit functions without truncation, the dates 0 and 60, `#NUM!` instead of `#VALUE!` on text that is not a number in the form needed (the complex IM*, the base conversions). The check is per function, but short; formulajs still computes.

**Own implementation, 86 functions.** Typical cases differ (78 functions) or the meaning at the edges (8), and no input rule fixes that: the old statistics names, the D functions, criteria (the engine's own already), DATE, TIME, TIMEVALUE, DAYS360, NETWORKDAYS, WORKDAY, TEXT, VALUE, NUMBERVALUE, DOLLAR, CHAR, CODE, PROPER, UPPER, SUBSTITUTE, MOD, ROMAN, ARABIC, ISO.CEILING, SORT, UNIQUE, LOOKUP, COLUMN, ROW, AGGREGATE, LINEST, LOGEST, TREND, MODE.MULT, XIRR; at the edges SWITCH (case), MATCH, VLOOKUP and SEARCH (wildcards, error input), DECIMAL, TYPE, XNPV, MINIFS. For some of them (the D functions: transpose the database and shift the field number; LINEST, LOGEST, TREND: the answer's shape) an adapter for one function would do, but that is a wrapper per function rather than a shared adapter.

**Hangs, 3 functions.** CHIINV, CHISQ.INV, CHISQ.INV.RT: a guard on the degrees of freedom or an own implementation.

## What could not be checked

- Excel for Mac only. CHAR and CODE on the Mac work in the Mac encoding, which depends on the interface language (`CHAR(200)` is `»`, `CODE("Ж")` 95 or 134); Excel for Windows answers otherwise, and the verdict on them is the Mac's.
- Two pairs of language and locale: English (en-US) and Russian (ru-RU). Mixed pairs, except the former column (en_US with the Russian interface), and other locales were not measured.
- Volatile functions are left out. ROW and COLUMN without arguments (they depend on the formula's cell) were not checked.
- Array sizes from an argument (MUNIT, EXPAND, TAKE, DROP) were checked with small literals only; large sizes and dates in that position were not: `MUNIT(45688)` and `MUNIT("1.5")` in ru-RU ran Excel out of resources and raised a modal dialog. For the other size and count arguments "large" is 1000.
- The Europe/Moscow time zone only. Bare formulajs gives midnight UTC for `DATEVALUE("2025-01-31")` (03:00 in Moscow); in a zone with a negative offset that would be the day before, and no such run was made.
- Array answers were compared by element and size, not as spilled ranges.
- Up to 15 cases per function (up to 19 with the answer's elements for functions with an array answer): rare argument modes (day-count bases, week types) are covered by one or two calls.
- The adapter rules of the proposal come from the differences and were not run.
- LibreOffice: the character `CHAR(0)` is written as `_x0000_` in its xlsx and is not decoded; this does not change the verdict (Excel gives `#VALUE!`).

## Repeating it

From the repository root:

```sh
npx tsx bench/coverage/run.ts all
```

It needs LibreOffice (`soffice`), Excel for Mac and permission to control Excel and System Events. Excel must have no open books before the run; the step restarts Excel with the interface language and locale it needs and leaves it running without arguments. Details and the rules for working with Excel: `bench/coverage/README.md`.
