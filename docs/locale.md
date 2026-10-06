# Locales

Excel depends on its language and regional settings in two ways: it shows and reads formulas in the user's language, and it computes some values by the regional rules (text read as a number or a date, numbers and booleans turned into text, format codes, currency). Books keep formulas in the invariant form of xlsx files (English names, `,` between arguments, `.` as the decimal separator) whatever the locale. Part 1 (done) covers formula text; part 2 (planned) covers computation.

## Part 1: formula text (done)

`src/locale/`: a locale is a plain object (`Locale` in `types.ts`); `en-US.ts` and `ru-RU.ts` are the two that exist. `toLocal`, `toInvariant` and `validateLocalFormula` translate piece by piece (`lex.ts`, `translate.ts`): function names, TRUE/FALSE, error literals, the decimal separator, the separators of arguments and of array constants. Strings, sheet names, defined names, references and whitespace are copied. The output is lexed again on the target side, so text that would change its meaning in translation (`1.5` in ru-RU) is a syntax error at its place. The lexer does not parse; array constants translate piece by piece like the rest.

The ru-RU data come from Excel 16.113 for Mac with the Russian interface and locale (`bench/locale/snapshot.ts`): 6,381 formulas written to books, read back as `formula` and `formula local`, plus Excel's international settings, all stored in `docs/locale/excel-ru-RU.json`. `bench/locale/derive.ts` pairs the calls of each English formula with those of its Russian text, in order, and `bench/locale/generate.ts` writes `src/locale/ru-RU.ts`. 513 of the 521 functions of Microsoft's list have a Russian name from Excel (468 differ from the English one); Excel for Mac does not know 6 (`ENCODEURL`, `FILTERXML`, `WEBSERVICE`, `REGEXEXTRACT`, `REGEXREPLACE`, `REGEXTEST`) and deletes 2 from loaded books (`CALL`, `REGISTER.ID`). Excel takes 8 error literals in formulas; `#SPILL!`, `#CALC!` and the newer ones are not literals for it. In ru-RU on the Mac the list separator is `;`, the decimal separator `,`, and array constants are `{1;2:3;4}` (`;` between columns, `:` between rows).

Tests: `test/locale.test.ts` (behaviour), `test/locale-excel.test.ts` (the data equal what Excel's text gives; both translations reproduce Excel's text of all 6,381 formulas), `test/locale-roundtrip.test.ts` (the 15,266 canonical formulas of the measurement, the case sets and the probes come back unchanged from ru-RU and en-US).

## Part 2: computing by a book's locale (plan)

### Where the locale lives

Proposal: a field of the book, `BookInput.locale?: string`, a BCP 47 tag of a built-in locale (`'en-US'` by default, `'ru-RU'`).

The locale of computation belongs to the book, not to the viewer or to the call. An xlsx file stores no calculation locale, so the same file gives different values in Excel on a Russian and on an English machine (`VALUE("1.5")` is 1.5 or 46143); a book that carries its locale gives the same values to everyone, while each viewer still sees formulas in their own language through `toLocal`. An option of `compute` would repeat Excel's problem: values would depend on who computes. A session takes the locale with the book; another locale is another session.

The field is a tag, not a `Locale` object: the engine resolves it in a registry of the package's locales (`LOCALES` in `src/locale/index.ts`, one line per locale file), and an unknown tag is an `InputError`. Locales of formula text stay objects that any host may build; a host-built locale for computation is a question below.

Hashes: `inputHash` takes the tag when it is given and not `en-US`; the values `hash` takes `;locale=<tag>` after the configuration version under the same condition. Books without the field, and en-US books, hash as now: the synthetic book keeps `8d63bad2c247ae26…`, and `CONFIG_VERSION` does not change, because en-US semantics do not.

### What depends on the locale

The measurement (`docs/coverage-formulajs.md`, "Excel ru-RU против en-US") has 280 cases where Excel ru-RU differs from en-US; 272 of them are in 166 whitelisted functions (the other 8 are `CODE` and `SORT`). By mechanism, each case counted once under its main one:

| mechanism | place in the engine | cases | Excel ru-RU |
|---|---|---|---|
| A. numeric text in a numeric argument | `values.ts` `toNum`, `args.ts` kinds `n`, `i`, `x` | 125 | `"1,5"` is 1.5; `"1.5"` is a date without a year (1 May of the current year, 46143 in 2026); `"0.5"` is no date and no number, `#VALUE!` |
| B. date text | `dates.ts` `parseDateText`, kinds `d`, `t`, `D` | 17 | day before month: `"1/2/2025"` is 1 February; English month names are not read (`DATEVALUE("22-MAY-2011")` is `#VALUE!`) |
| C1. number to text | `format.ts` `generalText`, `values.ts` `toText` | 20 | decimal comma: `CONCATENATE("a",12.5)` is `a12,5` |
| C2. boolean to text | `values.ts` `toText` | 17 | `ИСТИНА`, `ЛОЖЬ`: `LEN(TRUE)` is 6 |
| D. text to boolean | `values.ts` `toBool` | 5 | `"ИСТИНА"` is TRUE, `"TRUE"` is `#VALUE!` (`IF("TRUE",1,2)`) |
| E. TEXT format codes | `format.ts` `formatText` | 11 | Russian codes (`ДД.ММ.ГГГГ`, `ч:мм`, `# ##0,00`); English ones are literals or `#VALUE!` (`TEXT(45658,"YYYY-MM-DD")` is `YYYY-MM-DD`, `TEXT(1234.5,"0.00")` is `#VALUE!`) |
| F. DOLLAR, FIXED | `functions/text/numbers.ts` | 24 | `1 235 ₽`, `-1 200 ₽`, `1 234,6` (space between thousands, comma) |
| G. VALUE, NUMBERVALUE, DATEVALUE, TIMEVALUE | `functions/text/value.ts` | 11 | read by A and B; NUMBERVALUE's default separators are `,` and space |
| H. complex numbers as text | `functions/engineering/complex.ts` | 42 | `IMSUM(12.5,"5-3i")` is `17,5-3i` |

Outside the measurement's differences, the same rules reach operators (`&`, arithmetic and comparison on text in `runtime.ts`), criteria text (`">1,5"`, dates in criteria, `criteria.ts`), and the order of text: `criteria.ts` compares text with `Intl.Collator('ru-RU')` for every book now, which puts Cyrillic before Latin as Excel ru-RU does in `SORT`, while Excel en-US puts Latin first. `CHAR` and `CODE` follow the platform and language (Mac Roman, Mac Cyrillic) and could join the whitelist with a code page per locale.

### What the locale object gets

`Locale` gains the rules of computation, generated like the formula data. Shared with part 1: the decimal separator and the words TRUE and FALSE (Excel uses the same in `TRUE&""`).

- `groupSeparator`: ` ` in ru-RU (a plain space in Excel for Mac).
- `dateOrder` (`'mdy'`, `'dmy'`, `'ymd'`), date and time separators, the month and weekday names Excel reads and writes (full and short, with the case forms `ММММ` gives), what a year-less date means.
- `formatCodes`: year, month, day, hour, minute, second (`Г`, `М`, `Д`, `ч`, `м`, `с`), the name of the General format (`Основной`).
- `currency`: symbol, position, space, the negative form, digits (`1 234,57 ₽`, `-1 200 ₽`).
- `collation`: the `Intl.Collator` tag that orders text as this Excel does.
- `codePage` for `CHAR` and `CODE`, if they join the whitelist.

The international settings in `docs/locale/excel-ru-RU.json` already give the separators, the date order (1, day-month-year), the format codes, the currency symbol and its position. The rest needs a set of probes computed in Excel ru-RU (see "Data to measure").

### How each place gets the rules

The engine builds one rules object per book from the locale, as `translate.ts` builds its tables: `toNum`, `toText`, `toBool`, `parseDate`, `formatText`, `compareText`, the currency and fixed formats. The functions in `values.ts`, `dates.ts` and `format.ts` become methods of the en-US rules, with the same code and results, so callers outside a book (the parser, the criteria index) keep working. Inside a book the rules reach:

- function arguments through `Cx` (`cx.rules`): `args.ts` coerces every kind with `cx` at hand already, so `n`, `i`, `x`, `d`, `t`, `s`, `b` and the lists change in one place;
- operators through the compiled book: `runtime.ts` helpers for `&`, arithmetic on text and comparisons are created per book with its rules;
- functions that format or parse on their own (`TEXT`, `DOLLAR`, `FIXED`, `VALUE` and the rest, the complex functions) through `cx.rules`, declared with `cx: true`;
- criteria through the rules passed to `parseCrit` and `select`.

### Contract change

One `contract:` commit: `BookInput.locale?: string` with its JSDoc (a tag of the package's locales, `en-US` by default; it changes computed values, not formula text), `checkBook` checks the tag, `inputHash` and `valuesHash` take it as above, the README's API and Limitations sections say so.

### Acceptance

`test/excel-ru.test.ts`: the measurement's book with `locale: 'ru-RU'` and a clock in 2026 (Excel computed `excel-ru.json` on 5 October 2026, so a year-less date is in 2026), every whitelisted function's cases against `docs/coverage-formulajs/excel-ru.json` with the same comparison as `test/excel.test.ts`; cases that still differ go to a `KNOWN` list with reasons, as for en-US. The probes of the data step below get the same test. en-US tests and hashes do not change.

### Order of work, after the arrays branch is merged

The rules thread through `engine.ts`, `runtime.ts` and `args.ts`, which `arrays/incell` rewrites, so the work starts on top of it.

1. The contract commit and the registry of locales; the field does nothing yet (tests: tag check, hashes, en-US unchanged).
2. Rules object with the en-US methods, passed through `Cx` and the runtime; no value changes (all tests and the synthetic hash as before).
3. ru-RU rules by mechanism, each step with its cases from `excel-ru.json`: A and B with `dates.ts` taking the order and names (142 cases), C1, C2 and H through one number-to-text rule (79), D (5), F (24), G (11), E (11, the largest change: the format tokenizer reads the locale's codes).
4. Text order by locale, after measuring en-US order on mixed scripts; criteria text.

### Data to measure

A set of probes computed in Excel ru-RU, stored like `docs/excel-cases/`, with a step that runs a set in either locale (`run.ts probes` runs en-US only now):

- date text: `31.01.2025`, `31/01/2025`, `31-01-2025`, `2025-01-31`, `31 января 2025`, `31.янв.2025`, `янв 2025`, `1.5`, `1.5.25`, times with and without AM/PM;
- numeric text: `1 234,5` (space and no-break space), `1234,5 ₽`, `(1,5)`, `1,5%`, `1,5E+3`, `1.234,5`;
- TEXT with Russian codes and with English ones: dates, times, numbers, percent, currency, `Основной`, month and weekday names;
- criteria: `">1,5"`, `">=01.01.2026"`, `"=ИСТИНА"`;
- the order of mixed Cyrillic and Latin text in `MATCH` and comparisons, in ru-RU and en-US;
- `CHAR` and `CODE` over 1–255 if they are to join the whitelist.

## Questions

1. The book's locale as a field of `BookInput` (proposed) or as an option of `compute` and `openSession`?
2. Only the package's locales for computation (a tag), or also a host's own `Locale` object? The object would need its data checked and a hash of its rules in `hash`.
3. Hashes: take the tag only when it is not `en-US`, so existing hashes stay (proposed)?
4. A date without a year (`"1.5"` in ru-RU is 1 May of the current year, 101 cases of the measurement): follow Excel with the clock's year, which makes such cells depend on the clock like `TODAY`, or keep them `#VALUE!` as en-US does now with `"1/31"`?
5. Excel for Mac is the reference for ru-RU (array constants `{1;2:3;4}`, `CODE` in Mac Cyrillic). Excel for Windows may differ in these details; is the Mac enough, or is a Windows measurement needed?
6. `validateLocalFormula` messages name functions in English (`ROUND takes 2 arguments` for `ОКРУГЛ(1,5)`); positions are in the typed text. Should messages use the locale's names?
7. `toInvariant` also reads English function names, TRUE/FALSE and error texts in Russian text, which Excel ru-RU does not. Keep this leniency?
8. `CHAR` and `CODE`: with a code page per locale they could join the whitelist; worth it?
