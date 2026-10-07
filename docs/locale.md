# Locales

Excel depends on its language and regional settings in two ways. It shows and reads formulas in the user's language, and it computes some values by regional rules: text read as a number or a date, numbers and booleans turned into text, format codes, currency. A book here stores formulas in the invariant form of xlsx files (English function names, `,` between arguments, `.` as the decimal separator) whatever its locale. en-US and ru-RU are supported for both formula text and computation.

## Formula text

A locale is a plain object (`Locale`, `src/locale/types.ts`); the package has `enUS` and `ruRU`. `toLocal`, `toInvariant` and `validateLocalFormula` translate a formula piece by piece: function names, TRUE and FALSE, error literals, the decimal separator, and the separators of arguments and array constants. Strings, sheet names, defined names, references and whitespace are copied as they are.

The output is read again in the target locale, so text that would change its meaning in translation is a syntax error at its place: `1.5` typed in ru-RU, where `.` is not a decimal separator.

In ru-RU on the Mac the list separator is `;`, the decimal separator `,`, and array constants are written `{1;2:3;4}`, with `;` between columns and `:` between rows.

`toInvariant` also accepts English function names, TRUE, FALSE and error texts inside Russian text. Messages of `validateLocalFormula` are in English, with positions in the typed text.

### Where the ru-RU names come from

The Russian names were read from Excel 16.113 for Mac with the Russian interface. `bench/locale/snapshot.ts` wrote 6,381 formulas to books and read each back both as `formula` and as `formula local`, together with Excel's international settings (`test/excel/locale/ru-RU.json`). `bench/locale/derive.ts` pairs the function calls of each English formula with those of its Russian text, and `bench/locale/generate.ts` writes `src/locale/ru-RU.ts`.

513 of the 521 functions on Microsoft's list got a Russian name this way (468 differ from the English one). Excel for Mac does not know 6 functions (`ENCODEURL`, `FILTERXML`, `WEBSERVICE`, `REGEXEXTRACT`, `REGEXREPLACE`, `REGEXTEST`) and removes 2 from the books it loads (`CALL`, `REGISTER.ID`); these keep their English names. Excel accepts 8 error literals in formulas; `#SPILL!`, `#CALC!` and newer errors are not literals for it.

## Computing in a book's locale

A book carries its locale in `BookInput.locale`, a tag of one of the package's locales: `'en-US'` (the default) or `'ru-RU'`. An xlsx file stores no calculation locale, so the same file gives different values in Excel on a Russian and on an English machine: `VALUE("1.5")` is 1.5 in one and 46143 in the other. A book that carries its locale gives the same values to everyone, while each viewer still sees formulas in their own language through `toLocal`.

An unknown tag is an `InputError`. A host may build its own `Locale` objects for formula text, but computation uses only the package's locales.

The locale enters the hashes only when it is set and is not en-US: `inputHash` gets a line `L "<tag>"`, and the values hash gets `;locale=<tag>` after the configuration version. Books without a locale and en-US books hash as they did before.

### What changes in ru-RU

| | ru-RU, as Excel 16.113 for Mac computes |
|---|---|
| numbers in text | `1 234,5` and `1 2345` (a plain space before groups of three digits or more; a no-break space is no separator), `,5`, `(1,5)`, `1,5%`, `%(1,5)`, `1,5E+3`, `1 1/2`, `1 234,5 ₽` (the ruble after the amount). `$`, `р.` and `1.234,5` are not numbers |
| dates in text | day, month, year with `.`, `/` or `-`: `31.01.2025`, `1/2/2025` is 1 February, `31.01.25`; also `2025-01-31` and `2025.1.31`. A month by three letters or more of its nominative name in any case: `31 январь 2025`, `31.янв.2025`, `15 июн. 2025` (not the genitive `января`, not English names). `янв 2025` and `янв-25` are the 1st |
| two numbers | a day and a month of the current year (`1.5`, `15 янв`), or a month and a year when the second is past 12 (`1.13` is January 2013, `1.2025`) |
| times | as in en-US, with a decimal comma in seconds (`18:30:15,5`) and the English `AM`/`PM` |
| numbers as text | a decimal comma: `"a"&1.5` is `a1,5`, `1,23456789012345E+20`; complex numbers too (`1,5-2,25i`) |
| booleans | `ИСТИНА` and `ЛОЖЬ` both ways, any case when read; `"TRUE"` is `#VALUE!` where a boolean is expected |
| TEXT codes | `Д`, `Г`, `ч`, `с` in either case; `М` is the month, or minutes after hours or before seconds; `м` is always minutes; `,` is the decimal point and a space the group separator; `Основной` is General; Latin letters are literals (`yyyy` stays `yyyy`); month names in the nominative (`январь`, short `янв`, `февр`, `сент`, `нояб`), days `пятница`, `Пт` |
| DOLLAR, FIXED | `1 234,57 ₽`, `-1 234,57 ₽`, `-0,00 ₽` (a minus even when rounded to zero), `1 234 567,89` |
| VALUE and NUMBERVALUE | read as above; NUMBERVALUE's default separators are `,` and a space, and the rest of its text is read as VALUE reads it (`NUMBERVALUE("1.5")` is a date) |
| complex text | each part is read as VALUE reads it: `IMSUM("1.5+2i")` is `46143+2i` |
| errors, arrays as text | `#Н/Д` and the other Russian error texts in VALUETOTEXT, ARRAYTOTEXT and criteria; ARRAYTOTEXT writes `{1,5;2:ИСТИНА}` and `1,5; 2` |
| criteria | numbers, dates and words of the locale in criterion text and in cell text: `">1,5"`, `">=01.01.2026"`, `"01.05"`, `"ИСТИНА"` (`"TRUE"` is text) |
| text order | Cyrillic before Latin (en-US: Latin first) |
| CHAR, CODE | Windows-1251 (en-US: Windows-1252) |

### Dates without a year

Excel completes a date without a year with the current year: `1.5` in ru-RU, `1/31` in en-US. The engine takes the year of the book's clock, so a book's values stay a function of its input and clock.

Whether a cell reads the year depends on its data (`VALUE(A1)` reads it only when A1 holds such text), so the engine finds out while computing. The rules read the year through a clock object that sets a flag; the engine notes every formula that set it and treats it as volatile. Such cells are listed in `ComputeResult.volatile` and recomputed by `setClock` and in live sessions. A criterion or criteria index whose parsing read the year is not cached, so every formula that needs it reads the year itself.

### Text order

Excel en-US sorts Latin before Cyrillic, ru-RU Cyrillic before Latin. Both compare without regard to case but with regard to accents: `ё` comes after `е` and `й` after `и`, and neither pair is equal. `Intl.Collator` with the locale's tag and `sensitivity: 'accent'` gives Excel's order on all 174 measured cases. Operators, criteria, lookups and sorting all use it.

## How the engine does it

`src/locale/rules.ts` defines `Rules`, one object per book. en-US is the engine's own code. Another locale is built by `regionalRules` (`src/locale/rules/regional.ts`) from its data: separators, currency, month and day names, TEXT codes, the collation tag, and the words for TRUE, FALSE and errors. A new locale shaped like ru-RU is a data file and one line in `REGIONAL`.

The engine builds the rules when it creates a book's state and passes them to operators and criteria through the runtime, and to every function through `Cx.rules`. The argument kinds of `src/args.ts` convert values with the rules, and functions that read or write text themselves (TEXT, DOLLAR, FIXED, VALUE, the complex functions, the lookups and others) use them directly. Code outside a book, such as the parser, works in en-US.

## How it was checked

`test/excel-ru.test.ts` computes the measurement's book with `locale: 'ru-RU'` and compares every whitelisted function's cases with Excel ru-RU's answers in `test/excel/coverage/excel-ru.json`. `test/excel-ru-probes.test.ts` does the same for the ru-RU case sets in `test/excel/cases/ru-RU/`, among them `locale` (1,109 cases) and `collation` (174). `npx tsx bench/coverage/run.ts probes <set> ru` computes a set in Excel ru-RU. Known differences are listed in `RU_KNOWN` with a line each in the README.

`test/locale-excel.test.ts` checks that both translations reproduce Excel's own text of all 6,381 formulas, and `test/locale-roundtrip.test.ts` that 15,266 formulas of the case sets come back unchanged from ru-RU and en-US.

Excel for Mac is the reference for ru-RU. Excel for Windows may differ in details such as array constants, and was not measured.
