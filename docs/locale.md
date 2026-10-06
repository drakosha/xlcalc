# Locales

Excel depends on its language and regional settings in two ways: it shows and reads formulas in the user's language, and it computes some values by the regional rules (text read as a number or a date, numbers and booleans turned into text, format codes, currency). Books keep formulas in the invariant form of xlsx files (English names, `,` between arguments, `.` as the decimal separator) whatever the locale. Part 1 covers formula text, part 2 computation; both are done for en-US and ru-RU.

## Part 1: formula text (done)

`src/locale/`: a locale is a plain object (`Locale` in `types.ts`); `en-US.ts` and `ru-RU.ts` are the two that exist. `toLocal`, `toInvariant` and `validateLocalFormula` translate piece by piece (`lex.ts`, `translate.ts`): function names, TRUE/FALSE, error literals, the decimal separator, the separators of arguments and of array constants. Strings, sheet names, defined names, references and whitespace are copied. The output is lexed again on the target side, so text that would change its meaning in translation (`1.5` in ru-RU) is a syntax error at its place. The lexer does not parse; array constants translate piece by piece like the rest.

The ru-RU data come from Excel 16.113 for Mac with the Russian interface and locale (`bench/locale/snapshot.ts`): 6,381 formulas written to books, read back as `formula` and `formula local`, plus Excel's international settings, all stored in `test/excel/locale/ru-RU.json`. `bench/locale/derive.ts` pairs the calls of each English formula with those of its Russian text, in order, and `bench/locale/generate.ts` writes `src/locale/ru-RU.ts`. 513 of the 521 functions of Microsoft's list have a Russian name from Excel (468 differ from the English one); Excel for Mac does not know 6 (`ENCODEURL`, `FILTERXML`, `WEBSERVICE`, `REGEXEXTRACT`, `REGEXREPLACE`, `REGEXTEST`) and deletes 2 from loaded books (`CALL`, `REGISTER.ID`). Excel takes 8 error literals in formulas; `#SPILL!`, `#CALC!` and the newer ones are not literals for it. In ru-RU on the Mac the list separator is `;`, the decimal separator `,`, and array constants are `{1;2:3;4}` (`;` between columns, `:` between rows).

Tests: `test/locale.test.ts` (behaviour), `test/locale-excel.test.ts` (the data equal what Excel's text gives; both translations reproduce Excel's text of all 6,381 formulas), `test/locale-roundtrip.test.ts` (the 15,266 canonical formulas of the measurement, the case sets and the probes come back unchanged from ru-RU and en-US).

## Part 2: computing by a book's locale (done)

### Where the locale lives

A book carries its locale: `BookInput.locale?: string`, a BCP 47 tag of one of the package's locales (`'en-US'` by default, `'ru-RU'`). An xlsx file stores no calculation locale, so the same file gives different values in Excel on a Russian and on an English machine (`VALUE("1.5")` is 1.5 or 46143); a book that carries its locale gives the same values to everyone, while each viewer still sees formulas in their own language through `toLocal`. A session takes the locale with the book.

The field is a tag, not a `Locale` object: `checkBook` looks it up in `LOCALES` (`src/locale/registry.ts`, one line per locale file) and an unknown tag is an `InputError` at `book.locale`. A host may still build `Locale` objects of its own for formula text; computation takes only the package's locales.

Hashes: `inputHash` takes the line `L "<tag>"` after the style table when the tag is given and is not `en-US`; the values `hash` takes `;locale=<tag>` after the configuration version under the same condition (`src/canon.ts`, `valuesHash`). Books without the field, and en-US books, hash as before (the synthetic book's hash is pinned in `test/sha256.test.ts`), and `CONFIG_VERSION` did not change for locales.

### The rules of a locale

`src/locale/rules.ts` defines `Rules`, one object per book, and assembles it: `EN_US` is the engine's own en-US code (`values.ts`, `dates.ts`, `format.ts`, `criteria.ts`), unchanged; another locale is built by `regionalRules` (`src/locale/rules/regional.ts`) from its data (`RegionalData`; ru-RU's in `rules/ru-RU.ts`: separators, currency, month and day names, TEXT's codes, the collation tag, and the formula-text `Locale` for TRUE, FALSE and error texts) and the book's clock. The pieces of a day-month-year locale: `rules/numbers.ts` (number text), `rules/dates.ts` (the date part of date text; times and the frame of date text are shared with en-US in `dates.ts`), `rules/format.ts` (TEXT's codes, read into the tokens `format.ts` formats). A new locale with ru-RU's shape is a data file and a line in `REGIONAL`.

The engine builds the rules when it creates a book's state (`rulesOf` in `createState`) and passes them on:

- to the operators (`&`, arithmetic on text, comparisons) and to criteria selection through `RuntimeContext.rules` (`runtime.ts`);
- to every function through `Cx.rules`: the argument kinds of `args.ts` coerce with it, and the functions that read or write text themselves are declared with `cx: true` (`TEXT`, `DOLLAR`, `FIXED`, `VALUE`, `NUMBERVALUE`, `DATEVALUE`, `TIMEVALUE`, the complex functions, `CONCAT`, `TEXTJOIN`, `VALUETOTEXT`, `ARRAYTOTEXT`, the lookups and sorting for the text order, `IF`, `IFS`, `AND`, `OR`, `XOR`, the database functions, the ToolPak wrapper `atp`, and the rest that coerced on their own);
- code outside a book (the parser, `unstable/internal`'s criteria index) works by en-US.

### What depends on the locale, and the ru-RU rules

The measurement (`docs/coverage-formulajs.md`, "Excel ru-RU против en-US") has 280 cases where Excel ru-RU differs from en-US, 272 of them in whitelisted functions; the `locale` set (1,109 cases computed in Excel ru-RU) measured the rules behind them. By mechanism:

| mechanism | ru-RU, as Excel 16.113 for Mac computes | where |
|---|---|---|
| A. numeric text | `1 234,5`, `1 2345` (a plain space before groups of three digits or more; a no-break space is no separator), `,5`, `(1,5)`, `1,5%`, `%(1,5)`, `1,5E+3`, `1 1/2`, `1 234,5 ₽` (the ruble after the amount, not with `%`); `$`, `р.`, `1.234,5` are no numbers | `rules/numbers.ts` |
| B. date and time text | day-month-year with `.`, `/` or `-` (`31.01.2025`, `1/2/2025` is 1 February, `31.01.25`); `2025-01-31`, `2025.1.31`; a month by three letters or more of its nominative name, any case (`31 январь 2025`, `31.янв.2025`, `15 июн. 2025`, `сентяб`; not the genitive `января`, not English names); `янв 2025` and `янв-25` are the 1st; two numbers are a day and a month of the clock's year (`1.5`, `15 янв`), or a month and a year when the second is past 12 (`1.13` is January 2013, `1.2025`); times as in en-US with a decimal comma in seconds (`18:30:15,5`) and the English `AM`/`PM` | `rules/dates.ts`, `dates.ts` |
| C1, H. numbers as text | General with a decimal comma: `a1,5`, `1,23456789012345E+20`; complex parts the same (`1,5-2,25i`) | `regional.ts`, `complex.ts` |
| C2, D. booleans | `ИСТИНА`, `ЛОЖЬ` both ways, any case when read; `"TRUE"` is #VALUE! where a boolean is expected | `regional.ts` |
| E. TEXT codes | `Д`, `Г`, `ч`, `с` in either case; `М` the month, or minutes after hours or before seconds; `м` always minutes; `,` the decimal point, a space the group separator (grouping between placeholders, scaling by 1000 after them; before `₽` a literal), `Основной` the General format, among literals too; Latin letters are literals (`yyyy` stays `yyyy`); month names nominative (`январь`, short `янв`, `февр`, `сент`, `нояб`), days `пятница`, `Пт`; a format with a point before a digit placeholder, or a space in a fraction, is #VALUE! whatever the value | `rules/format.ts`, `format.ts` |
| F. DOLLAR, FIXED | `1 234,57 ₽`, `-1 234,57 ₽`, `-0,00 ₽` (a minus, even rounded to zero), `1 234 567,89` | `functions/text/numbers.ts` with the locale's data |
| G. VALUE and the like | by A and B; NUMBERVALUE's defaults are the locale's (`,` and a space), the rest of its text read as VALUE reads it (`NUMBERVALUE("1.5")` is a date) | `functions/text/value.ts`, `functions/date/` |
| H. complex text | each part read as VALUE reads it (`IMSUM("1.5+2i")` is `46143+2i`) | `complex.ts` (`Rules.complexPart`) |
| errors, arrays | error texts by the formula-text locale (`#Н/Д`), in VALUETOTEXT, ARRAYTOTEXT and criteria; ARRAYTOTEXT writes `{1,5;2:ИСТИНА}` and `1,5; 2` | `regional.ts`, `astext.ts`, `criteria.ts` |
| criteria | numbers, dates and words of the locale in criterion text and in cell text (`">1,5"`, `">=01.01.2026"`, `"01.05"`, `"ИСТИНА"`; `"TRUE"` is text) | `criteria.ts` (`CritRules`) |

### Dates without a year and the clock

Excel completes a date without a year with the current year; the engine takes the year of the book's clock (`Clock.now`), so a book's values stay a function of its input and clock. Whether a cell reads the year depends on the data (`VALUE(A1)` reads it only when A1 holds such text), so it is found while computing: the rules read the year through a `RulesClock` that the engine creates per book (`createState`), and reading it sets a flag. `evaluate` clears the flag before each formula and records the formulas that set it in `EngineState.clocked`; `volatile` is the closure of the volatile functions and these formulas over their dependents, so `setClock` and live sessions recompute them, and `ComputeResult.volatile` lists them. The cost is one flag write and one read per formula computed.

Caches would hide a read: a criterion parsed once and an index built once serve later formulas without reading the year again. So a parsed criterion or a criteria index whose building read the year is not cached (`readsClock` in `rules/data.ts`, `parseCrit`, the runtime's `indexOf`); each formula that needs it reads the year itself. Such books are rare, and the rest are cached as before. `test/locale-clock.test.ts` checks that a session after `setClock` equals a full computation with the same clock, also through repeated criteria and a shared index.

### Text order

Excel en-US sorts Latin before Cyrillic, ru-RU Cyrillic before Latin; both compare case-insensitively and accent-sensitively (`ё` after `е`, `й` after `и`, equal neither). The `collation` set (174 cases of operators, SORT, criteria, MATCH and UNIQUE on mixed scripts, computed in both) matches `Intl.Collator` with the locale's tag (`en-US`, `ru-RU`) and `sensitivity: 'accent'` on every case; the order comes from `Rules.compareText` for operators, criteria, lookups and sorting.

### What the measurement showed besides the locale

Some differences came from rules of Excel that the engine had wrong in en-US too, measured in both locales and fixed for both:

- criteria read their text and a cell's text as VALUE does, and a number criterion's `<>` holds for all text (`criteria-text` set, 37 cases in each locale);
- AND, OR and XOR take a direct text argument that is a boolean word as that boolean (`logical-text`, 51 cases each);
- DATE takes a day outside −32768…32767 as 32767, a month of 32767 or more is #NUM! (`date-day`, en-US; ru-RU's `DATE(2025,1,"1.5")` led to it);
- TEXT rounds a number multiplied by `%` or scaled to 15 significant digits first (`format-scale`, en-US; `TEXT(0.285,"0%")` is `29%`).

### Acceptance

`test/excel-ru.test.ts` computes the measurement's book with `locale: 'ru-RU'` and the clock of the case books (2026, the year Excel computed `excel-ru.json` in) and compares every whitelisted function's cases with `test/excel/coverage/excel-ru.json` as `test/excel.test.ts` does; `test/excel-ru-probes.test.ts` does the same for the sets in `test/excel/cases/ru-RU/` (`locale`, `collation`, `criteria-text`, `logical-text`; `run.ts probes <set> ru` computes a set in Excel ru-RU). Known differences are pinned in `RU_KNOWN` (`test/excel-ru-cases.ts`) with README lines: the four en-US `gammaInvConvergence` cases of the measurement, `#NULL!` read as `#ERROR!` (1), exponent formats and elapsed time in TEXT, unsupported in both locales (5), and a second decimal separator among a format's decimals (1). en-US acceptance is unchanged, with the new en-US sets added.

## Decisions

The questions of the plan, with the answers taken as defaults (each is a local change if it is decided otherwise):

1. The locale is a field of the book, `BookInput.locale`, not an option of `compute` or `openSession`.
2. Only the package's locales compute (a tag); a host's own `Locale` object is for formula text only. Accepting one would need its rules data checked and hashed.
3. Hashes take the tag only when it is not en-US, so existing hashes stay and `CONFIG_VERSION` does not change.
4. A date without a year follows Excel with the clock's year, and such cells depend on the clock as described above.
5. Excel for Mac is the reference for ru-RU; a measurement in Excel for Windows is not done.
6. `validateLocalFormula` messages stay English (`ROUND takes 2 arguments` for `ОКРУГЛ(1,5)`), positions in the typed text; Russian messages can come with an interface.
7. `toInvariant` keeps reading English function names, TRUE/FALSE and error texts in Russian text.
8. `CHAR` and `CODE` stay out of the whitelist until a code page per locale is decided (Mac Roman, Mac Cyrillic, Windows-1251 differ).
