# Locales

Excel depends on its language and regional settings in two ways. It shows and reads formulas in the user's language, and it computes some values by regional rules: text read as a number or a date, numbers and booleans turned into text, format codes, currency. A book here stores formulas in the invariant form of xlsx files (English function names, `,` between arguments, `.` as the decimal separator) whatever its locale. en-US and ru-RU are supported for both formula text and computation.

## Formula text

A locale is a plain object (`Locale`, `src/locale/types.ts`); the package has `enUS` and `ruRU`. `toLocal`, `toInvariant` and `validateLocalFormula` translate a formula piece by piece: function names, TRUE and FALSE, error literals, the decimal separator, and the separators of arguments and array constants. Strings, sheet names, defined names, references and whitespace are copied as they are.

The output is read again in the target locale, so text that would change its meaning in translation is a syntax error at its place: `1.5` typed in ru-RU, where `.` is not a decimal separator.

In ru-RU on the Mac the list separator is `;`, the decimal separator `,`, and array constants are written `{1;2:3;4}`, with `;` between columns and `:` between rows.

`toInvariant` also accepts English function names, TRUE, FALSE and error texts inside Russian text. Messages of `validateLocalFormula` are in English, with positions in the typed text.

Structured references translate too: the specifiers by the locale's `tableItems` (in ru-RU `#Все`, `#Данные`, `#Заголовки`, `#Итоги` and `#Эта строка`, which Excel shows as `@`) and the separator between their items (`Sales[[#Заголовки];[Qty]]`). `toLocal` shows them as Excel does, and given the table of the formula's cell drops its name where Excel drops it ([Tables](tables.md)).

### Where the ru-RU names come from

The Russian names were read from Excel 16.113 for Mac with the Russian interface. `bench/locale/snapshot.ts` wrote 6,381 formulas to books and read each back both as `formula` and as `formula local`, together with Excel's international settings (`test/excel/locale/ru-RU.json`). `bench/locale/derive.ts` pairs the function calls of each English formula with those of its Russian text, and `bench/locale/generate.ts` writes `src/locale/ru-RU.ts`.

513 of the 521 functions on Microsoft's list got a Russian name this way (468 differ from the English one). Excel for Mac does not know 6 functions (`ENCODEURL`, `FILTERXML`, `WEBSERVICE`, `REGEXEXTRACT`, `REGEXREPLACE`, `REGEXTEST`) and removes 2 from the books it loads (`CALL`, `REGISTER.ID`); these keep their English names. Excel accepts 8 error literals in formulas; `#SPILL!`, `#CALC!` and newer errors are not literals for it.

## Computing in a book's locale

A book carries its locale in `BookInput.locale`, a tag of one of the package's locales: `'en-US'` (the default) or `'ru-RU'`. An xlsx file stores no calculation locale, so the same file gives different values in Excel on a Russian and on an English machine: `VALUE("1.5")` is 1.5 in one and 46143 in the other. A book that carries its locale gives the same values to everyone, while each viewer still sees formulas in their own language through `toLocal`.

An unknown tag is an `InputError`. A host may build its own `Locale` objects for formula text, but computation uses only the package's locales.

The locale enters the hashes only when it is set and is not en-US: `inputHash` gets a line `L "<tag>"`, and the values hash gets `;locale=<tag>` after the configuration version. Books without a locale and en-US books hash as they did before.

### What changes in ru-RU

| | ru-RU, as Excel 16.113 for Mac computes, and as Excel for Windows where the two part |
|---|---|
| numbers in text | `1 234,5` and `1 2345` (a plain space before groups of three digits or more; a no-break space is no separator), `,5`, `(1,5)`, `1,5%`, `%(1,5)`, `1,5E+3`, `1 1/2`, `1 234,5 ₽` (the ruble after the amount). `$`, `р.` and `1.234,5` are not numbers |
| dates in text | day, month, year with `.`, `/` or `-`: `31.01.2025`, `1/2/2025` is 1 February, `31.01.25`; also `2025-01-31` and `2025.1.31`. A month by three letters or more of its nominative name in any case: `31 январь 2025`, `31.янв.2025`, `15 июн. 2025` (not the genitive `января`, not English names). `янв 2025` and `янв-25` are the 1st |
| two numbers | a day and a month of the current year (`1.5`, `15 янв`), or a month and a year when the second is past 12 (`1.13` is January 2013, `1.2025`) |
| times | as in en-US, with a decimal comma in seconds (`18:30:15,5`) and the English `AM`/`PM` |
| numbers as text | a decimal comma: `"a"&1.5` is `a1,5`, `1,23456789012345E+20`; complex numbers too (`1,5-2,25i`) |
| booleans | `ИСТИНА` and `ЛОЖЬ` both ways, any case when read; `"TRUE"` is `#VALUE!` where a boolean is expected |
| TEXT codes | `Д`, `Г`, `ч`, `с` in either case; `М` is the month, or minutes after hours or before seconds; `м` is always minutes; `,` is the decimal point and a space the group separator; `Основной` is General; Latin letters are literals (`yyyy` stays `yyyy`); month names in the nominative with a capital (`Январь`, `Я` for `МММММ`), short ones of three letters (`янв`, `фев`, `сен`, `ноя`), as Excel for Windows writes them; days `пятница`, `Пт` |
| DOLLAR, FIXED | `1 234,57 ₽`, `-1 234,57 ₽`, `-0,00 ₽` (a minus even when rounded to zero), `1 234 567,89` |
| VALUE and NUMBERVALUE | read as above; NUMBERVALUE's default separators are `,` and a space, and the rest of its text is read as VALUE reads it (`NUMBERVALUE("1.5")` is a date) |
| complex text | each part is read as VALUE reads it: `IMSUM("1.5+2i")` is `46143+2i` |
| errors, arrays as text | `#Н/Д` and the other Russian error texts in VALUETOTEXT, ARRAYTOTEXT and criteria; ARRAYTOTEXT writes `{1,5;2:ИСТИНА}` and `1,5; 2` |
| criteria | numbers, dates and words of the locale in criterion text and in cell text: `">1,5"`, `">=01.01.2026"`, `"01.05"`, `"ИСТИНА"` (`"TRUE"` is text) |
| text order | Latin before Cyrillic, as in en-US and as Excel for Windows sorts (Excel for Mac ru-RU puts Cyrillic first) |
| times of built-in formats | the hour without a leading zero (`6:00` for the formats 20 and 21 and `[$-F400]`), as Excel for Windows shows them |
| CHAR, CODE | Windows-1251 (en-US: Windows-1252) |

### Dates without a year

Excel completes a date without a year with the current year: `1.5` in ru-RU, `1/31` in en-US. The engine takes the year of the book's clock, so a book's values stay a function of its input and clock.

Whether a cell reads the year depends on its data (`VALUE(A1)` reads it only when A1 holds such text), so the engine finds out while computing. The rules read the year through a clock object that sets a flag; the engine notes every formula that set it and treats it as volatile. Such cells are listed in `ComputeResult.volatile` and recomputed by `setClock` and in live sessions. A criterion or criteria index whose parsing read the year is not cached, so every formula that needs it reads the year itself.

### Text order

Text sorts Latin before Cyrillic in both locales, as Excel for Windows sorts in every locale; Excel for Mac ru-RU puts Cyrillic first, and the engine follows Windows there. The order is case-insensitive and accent-sensitive: `"аб"="АБ"` is TRUE, `"ё"="е"` and `"й"="и"` are FALSE. An accent and `ё` weigh only where the letters are otherwise equal, while `й` is a letter of its own after `и`:

- `"ёа"<"еб"` is TRUE, as `"éa"<"eb"`: the second letter decides before the accent.
- `"шел"<"шёл"` is TRUE: with the letters equal, the plain one comes first.
- `"йа"<"иб"` is FALSE: `й` comes after every word that starts with `и`.
- `ъ`, `ы`, `ь` come in that order, after `й` and before `э`, `ю`, `я`.

`Intl.Collator` with the en-US order (`collation: 'en-US'` in both locales) and `sensitivity: 'accent'` gives Excel for Windows's order on all 236 measured cases of ru-RU and 174 of en-US. Operators, criteria, lookups and sorting all use it.

## How the engine does it

`src/locale/rules.ts` defines `Rules`, one object per book. en-US is the engine's own code. Another locale is built by `regionalRules` (`src/locale/rules/regional.ts`) from its data: separators, currency, month and day names, TEXT codes, the collation tag, the code page of CHAR and CODE, the word of new table columns' names, and the words for TRUE, FALSE and errors. Some rules are still code written for ru-RU, so a new locale is its data files and their registration, plus code where its rules differ from those ([Adding a locale](#adding-a-locale)).

The engine builds the rules when it creates a book's state and passes them to operators and criteria through the runtime, and to every function through `Cx.rules`. The argument kinds of `src/args.ts` convert values with the rules, and functions that read or write text themselves (TEXT, DOLLAR, FIXED, VALUE, the complex functions, the lookups and others) use them directly. Code outside a book, such as the parser, works in en-US.

## Adding a locale

A locale added to the package lets books in it compute as Excel in that language and region does, and lets hosts show formulas in its language. Most of a locale is data in two files. Some rules are still code written for ru-RU, listed under [What needs code](#what-needs-code), so a locale such as de-DE or tr-TR is its data, the code its rules need, and a check against Excel as ru-RU had.

### Two layers

Formula text is a `Locale` object (`src/locale/types.ts`): the separators, the words for TRUE and FALSE, error literals, function names and the specifiers of structured references. ru-RU's was taken from Excel: `bench/locale/snapshot.ts` reads Excel's own text of 6,381 formulas and its international settings into `test/excel/locale/ru-RU.json`, `bench/locale/derive.ts` derives the locale's data from them, and `bench/locale/generate.ts` writes `src/locale/ru-RU.ts`. The snapshot's books have no tables, so `generate.ts` keeps the specifiers by hand (`TABLE_ITEMS`), as the table set `locale` measured them.

Computation is a `RegionalData` object (`src/locale/rules/data.ts`), ru-RU's in `src/locale/rules/ru-RU.ts`: the separators of numbers in text, the currency, month and day names, TEXT's codes and colour names, the date separator of cells, Excel's code of the locale (`419`), the long date and time, the collation tag, the built-in formats, the code page of CHAR and CODE (`codePage`, Windows-1251) and the word of new table columns' names (`columnWord`, `Столбец`, which makes `Столбец1`). `regionalRules` (`src/locale/rules/regional.ts`) builds a book's rules from it. en-US has no data of its own: its rules are the engine's code in `src/locale/rules.ts`, where Windows-1252 and `Column` are.

### Registering it

A new locale `xx-YY` takes these places:

- `src/locale/xx-YY.ts`, its `Locale`, exported from `src/locale/index.ts` and `src/index.ts` as `ruRU` is, for `toLocal` and `toInvariant`.
- A line in `LOCALES` (`src/locale/registry.ts`). A book's `locale` and the locale stored in an xlsx file must be one of its tags, and a locale's words for TRUE and FALSE are no names in its books.
- `src/locale/rules/xx-YY.ts`, its `RegionalData`, and a line in `REGIONAL` (`src/locale/rules.ts`). A tag in `LOCALES` without one would compute by en-US's rules; `test/locale-book.test.ts` checks that both maps list the same locales besides en-US.
- The published table of its code page in `PAGES` of `test/codepage.test.ts`, which fails until every locale's page is checked against one.
- The tags listed in the comments of `BookInput.locale` and `DisplayOptions.locale` (`src/contract.ts`), and in the message `test/locale-book.test.ts` expects for an unknown tag.

### Where the data come from

Function names and the rest of formula text are read from Excel for Mac in the locale's language, as ru-RU's were. The Windows harness has no snapshot of formula text.

Regional data (the separators, the currency, month names, TEXT codes, the order of text, the code page, the word of column names) are checked against Excel for Windows too. Where Excel for Mac and Excel for Windows part, Windows is the reference. ru-RU follows it for the order of text (Latin before Cyrillic), month names (`Январь`, not `январь`), the short and long time, and CHAR and CODE (Windows-1251, not Mac Cyrillic); it keeps one value of the Mac's, the long date `[$-F800]`.

### Accepting a locale

A locale is ready when the engine gives Excel's answers in it, and every difference left is listed with its reason. For ru-RU that is:

- the coverage measurement computed in Excel ru-RU (`npx tsx bench/coverage/run.ts excel ru`) and compared by `test/excel-ru.test.ts`;
- the case sets (`npx tsx bench/coverage/run.ts probes <set> ru`), what cells show (`display <set> ru`) and tables (`tables <set> ru`), with their tests, as `bench/coverage/README.md` describes;
- the same books in Excel for Windows (`npx tsx bench/windows/run.ts all ru`, `bench/windows/README.md`);
- the differences: `RU_KNOWN` in `test/excel-ru-cases.ts` pins the engine's answer of each case that differs from Excel ru-RU, with a reason from `RU_REASONS`, and `WIN_KNOWN` does the same against Excel for Windows. A pinned case that comes to match Excel fails until it is taken off the list.

A new locale needs the same: the measurement and the sets computed in its Excel, and its own list of differences.

### What needs code

These rules are written for ru-RU, or for en-US and ru-RU only. A new locale needs code for them, not only data.

#### Dates in text

`regionalRules` reads the date text of every locale besides en-US as day, month, year (`dmyDates`, `src/locale/rules/dates.ts`): `31.01.2025`, `31/1/25`, also `2025-01-31`, and a month by three letters or more of its nominative name. A locale that writes the month or the year first needs its own reader.

#### Currency in number text

VALUE and the other readers of number text take the currency symbol after the amount only (`numberReader`, `src/locale/rules/numbers.ts`), as in `1 234,5 ₽`. `currency.before` changes what DOLLAR writes, not what is read, so a locale with the symbol before the amount needs code there.

#### Widths of characters

What a cell shows by its column's width adds up the widths of characters measured for Calibri 11 in Excel for Mac (`src/numfmt/glyphs.ts`): the basic Latin letters, the Russian alphabet, digits and some signs. Any other character counts as a digit, so text with `ä`, `ß`, `ş` or `ğ` may fit a column here where Excel shows `#`, or the other way round.

#### Differences from Excel

The machinery that keeps differences from Excel is ru-RU's alone: `RU_KNOWN`, `RU_SETS` and `runRuSet` in `test/excel-ru-cases.ts`, used by `test/excel-ru.test.ts` and `test/excel-ru-probes.test.ts`. A new locale needs its own, or this one made to take a locale.

#### Case of names

Sheet names, table names and the words of formula text are compared in upper case by Russian rules, `toLocaleUpperCase('ru-RU')`: in `src/syntax/print.ts` (`sheetKey`, which table names use too), `src/syntax/table.ts` (the specifiers of structured references), `src/xlsx/read/book.ts` (table names in a file), `src/locale/lex.ts` and `src/locale/translate.ts` (function names, TRUE, FALSE and specifiers in typed text). Defined names are compared with `toUpperCase()` (`src/names.ts`), which is the same for Latin and Cyrillic. Both take `i` and `ı` to `I` and keep `İ` apart, while Turkish rules pair `i` with `İ` and `ı` with `I`; how Excel tr-TR compares such names has not been measured.

#### The harnesses

The harnesses know en-US and ru-RU only. On the Mac, `bench/coverage/excel.ts` has `type Locale = 'en' | 'ru'` and Excel's launch arguments by locale (`LAUNCH`): en-US starts Excel with `-AppleLanguages (en) -AppleLocale en_US`, ru-RU with none, in the system's Russian. `bench/locale/snapshot.ts` takes `en` and `ru`, `bench/locale/generate.ts` `ru` alone. On Windows, `bench/windows/remote.ts` has `Tag = 'en-US' | 'ru-RU'` with each locale's code and code page, and the machine needs the language: Excel's interface language follows the Windows display language, so a new language needs its Windows language pack, and Office's language pack for it as well, as the English one is installed there.

## How it was checked

`test/excel-ru.test.ts` computes the measurement's book with `locale: 'ru-RU'` and compares every whitelisted function's cases with Excel ru-RU's answers in `test/excel/coverage/excel-ru.json`. `test/excel-ru-probes.test.ts` does the same for the ru-RU case sets in `test/excel/cases/ru-RU/`, among them `locale` (1,109 cases) and `collation` (174). `npx tsx bench/coverage/run.ts probes <set> ru` computes a set in Excel ru-RU. Known differences are listed in `RU_KNOWN` with a line each in the README; among them are the cases where the engine follows Excel for Windows and the Mac answers otherwise (the order of text, month names).

`test/excel-win.test.ts` checks the same books against Excel for Windows (`test/excel/win/`, `bench/windows`), with the set `collation-more` (62 cases of the order inside Cyrillic and Latin) measured in Windows alone.

`test/locale-excel.test.ts` checks that both translations reproduce Excel's own text of all 6,381 formulas, and `test/locale-roundtrip.test.ts` that 15,266 formulas of the case sets come back unchanged from ru-RU and en-US.

Excel for Mac and Excel for Windows (16.0.20430) compute ru-RU alike but for their regional data. Where those part, the engine follows Windows: the order of text, month names, the times of built-in formats, `CHAR` and `CODE`. It keeps the Mac's long date `[$-F800]` (`пятница, 31 января 2025 г.`; Windows's has no weekday). Formula text in ru-RU (array constants and the like) was not compared with Windows.
