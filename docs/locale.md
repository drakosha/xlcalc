# Locales

Excel depends on its language and regional settings in two ways. It shows and reads formulas in the user's language, and it computes some values by regional rules: text read as a number or a date, numbers and booleans turned into text, format codes, currency. A book here stores formulas in the invariant form of xlsx files (English function names, `,` between arguments, `.` as the decimal separator) whatever its locale. en-US and ru-RU are supported for both formula text and computation.

## Loading a locale

en-US is built in. ru-RU is a module of its own, `xlcalc/locales/ru-RU`, so a bundle that has no ru-RU book does not carry its 52 KB of formula text, rules and messages. A host imports it and loads it once, before a book in it computes:

```js
import { compute, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
compute({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { formula: 'VALUE("1,5")' } } }] }) // A1 is 1.5
```

- `LOCALES` lists the loaded locales by tag: `en-US`, then each locale registered.
- A book's `locale` is a tag of a loaded locale. A book in a locale that is not loaded is an `InputError` that says what to do: `book.locale: the locale ru-RU is not loaded: import { ruRU } from 'xlcalc/locales/ru-RU' and call registerLocale(ruRU) (loaded: en-US)`. So is the tag given to a function that takes a locale. Another tag lists the loaded ones, and names the tag meant when only its case differs (`ru-ru`): tags are compared as written.
- Every function that takes a locale takes the tag of a loaded one or the locale object: `toLocal`, `toInvariant`, `validateLocalFormula`, `parseEntry`, `displayValue`, `entryWidth`, `builtinFormatCode`, `entryText`, `needsApostrophe`, `errorText`, `scanReferences`, `validateSheetName`, `validateName` and `validateNames`. An object needs no loading there, since no book is computed: `toLocal(formula, ruRU)` works in a bundle that never calls `registerLocale`.
- Importing the module does nothing by itself, so a bundler keeps it out of a bundle that does not use it. Loading the same object again does nothing; another object with a loaded tag is an `InputError`, as is a locale without the rules a book computes by.
- `readXlsx` keeps a file's locale (`ru-RU`) as it reads it, loaded or not; computing or writing the book needs the locale loaded.
- A pool's threads load the locales loaded where the pool runs, those loaded after it started too ([docs/async.md](async.md#locales)).

What counts in books of every locale stays in the core, so an en-US book shows the same whatever is loaded: the widths of Cyrillic letters (`src/numfmt/glyphs.ts`) and ru-RU's month and day names, which the format code `[$-419]` gives in a book of any locale (`src/locale/lcid.ts`).

## Formula text

A locale is a plain object (`Locale`, `src/locale/types.ts`): its formula text, and for a locale a book computes in, its rules of computation (`regional`) and messages. en-US's is `enUS` of the root, ru-RU's `ruRU` of its module, and `LOCALES` has the loaded ones by tag (`LOCALES.get('ru-RU')`). `toLocal`, `toInvariant` and `validateLocalFormula` take a loaded locale's tag or a locale object, and translate a formula piece by piece: function names, TRUE and FALSE, error literals, the decimal separator, and the separators of arguments and array constants. Strings, sheet names, defined names, references and whitespace are copied as they are.

The output is read again in the target locale, so text that would change its meaning in translation is a syntax error at its place: `1.5` typed in ru-RU, where `.` is not a decimal separator.

In ru-RU on the Mac the list separator is `;`, the decimal separator `,`, and array constants are written `{1;2:3;4}`, with `;` between columns and `:` between rows.

`toInvariant` also accepts English function names, TRUE, FALSE and error texts inside Russian text. Messages of `toLocal`, `toInvariant` and `validateLocalFormula` are in the locale's language ([Messages](#messages)), with positions in the typed text.

Structured references translate too: the specifiers by the locale's `tableItems` (in ru-RU `#Все`, `#Данные`, `#Заголовки`, `#Итоги` and `#Эта строка`, which Excel shows as `@`) and the separator between their items (`Sales[[#Заголовки];[Qty]]`). `toLocal` shows them as Excel does, and given the table of the formula's cell drops its name where Excel drops it ([Tables](tables.md)).

### Where the ru-RU names come from

The Russian names were read from Excel 16.113 for Mac with the Russian interface. `bench/locale/snapshot.ts` wrote 6,381 formulas to books and read each back both as `formula` and as `formula local`, together with Excel's international settings (`test/excel/locale/ru-RU.json`). `bench/locale/derive.ts` pairs the function calls of each English formula with those of its Russian text, and `bench/locale/generate.ts ru-RU` writes them into ru-RU's locale file, `src/locales/ru-RU.ts`.

513 of the 521 functions on Microsoft's list got a Russian name this way (468 differ from the English one). Excel for Mac does not know 6 functions (`ENCODEURL`, `FILTERXML`, `WEBSERVICE`, `REGEXEXTRACT`, `REGEXREPLACE`, `REGEXTEST`) and removes 2 from the books it loads (`CALL`, `REGISTER.ID`); these keep their English names. Excel accepts 8 error literals in formulas; `#SPILL!`, `#CALC!` and newer errors are not literals for it.

## Computing in a book's locale

A book carries its locale in `BookInput.locale`, the tag of a loaded locale: `'en-US'` (the default) or `'ru-RU'` once loaded. An xlsx file stores no calculation locale, so the same file gives different values in Excel on a Russian and on an English machine: `VALUE("1.5")` is 1.5 in one and 46143 in the other. A book that carries its locale gives the same values to everyone, while each viewer still sees formulas in their own language through `toLocal`.

A tag that is not loaded is an `InputError` ([Loading a locale](#loading-a-locale)). A host may build its own `Locale` objects for formula text; one without rules of computation reads and shows values as en-US does, and `registerLocale` refuses it, so a book computes only in a locale with its own rules.

In every locale a percent in text keeps the digits written: `VALUE("1,1%")` in ru-RU, as `VALUE("1.1%")` in en-US, is 0.011, the number the text reads as once the decimal point moves two places, as Excel reads it in VALUE, operators and numeric arguments, not 1.1 divided by 100 (0.011000000000000001; the difference shows when it is magnified, `(VALUE("1.1%")-0.011)*1E+20` is 0). `NUMBERVALUE` divides by 100 for each `%` in turn and the operator `%` divides its operand, as Excel does (`NUMBERVALUE("1.1%%")` is 1.1/100/100). Money with percent is no number (`"$1.1%"`, `"1,1 ₽%"`). The set `pct-text` measured these in both Excels and both locales.

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

## Messages

A person working in a book's locale reads the engine's messages in its language: why a typed formula does not read, why a structure operation is refused, where a value comes from. In ru-RU they are Russian:

| where | the message in ru-RU |
|---|---|
| `parseEntry('=СУММ(1', { locale: 'ru-RU' })` | `Ожидалась ), а формула закончилась` |
| `parseEntry('=ОКРУГЛ(1;2;3)', { locale: 'ru-RU' })` | `Функция ОКРУГЛ принимает 2 аргумента, а указано 3` |
| `validateLocalFormula('=СУММ(1.5)', ruRU)` | `Здесь не может стоять «.»` |
| `editStructure` deleting rows through an old-kind array | `Нельзя изменить часть массива S!C1 (C1:C3)` |
| `editStructure` renaming a sheet to a name taken | `Лист «Итоги» уже есть (имена листов не различают регистр)` |
| `explain` | `'План'!B2 = 480  ← ОКРУГЛ(B1*1,2;0)` |

### Which locale

- `parseEntry` writes in `opts.locale`, and `toLocal`, `toInvariant` and `validateLocalFormula` in the locale they are given. A locale object speaks by its own `messages`; without them, as the loaded locale of its tag, and in English for another tag.
- `editStructure`, and a session's `editStructure` alike in the same thread, in a thread of the pool or synchronously, write in the book's `locale`.
- `explain`'s `text` is written in the book's locale: its words (`(2 ячейки)`, `перенос из`, `(ошибка из 'План'!A2)`), numbers with the locale's decimal separator, formulas as `toLocal` shows them, booleans in the locale's words (`ИСТИНА`), errors as the book's cells show them (`#Н/Д`), and an empty cell as `пусто`. The tree's `ref`, `value`, `formula` and `origin` stay as the book stores them, for programs.

Codes (`code`) and positions (`at`) are the same in every language, so a host that chose its own words by the code may go on doing so, or show `message` as it is.

### What stays in English

- `validateFormula`, which has no locale, and every message of an en-US book or call.
- The reasons of cell errors and the issues of names (`diagnostics[].reason`) and `validateName`'s message.
- `InputError`, a program's mistake rather than a person's.

### How the Russian is written

As an interface writes to a person: short sentences, a capital letter first and no full stop, no long dash. Where Excel ru-RU has a message of its own for the same refusal, the text stays close to it (`Нельзя изменить часть массива`). What comes from the typed formula is written in the locale: function names (`ОКРУГЛ`), separators (`Разделитель «;» вне скобок функции`), specifiers of structured references (`#Данные`). Names of sheets, tables and defined names stand in «» (`Лист «Итоги» уже есть`), references as they are (`S!C1`). A number takes its noun in the right form: `1 строку`, `2 строки`, `5 строк`.

### The catalog

Each message is a key with parameters (`MessageParams` in `src/locale/messages.ts`): the parser, the checks and the structure operations name the message where they refuse, and its text is written only then, so messages cost nothing while a book computes. A locale's `messages` write the text of each key from its parameters and the locale's words (`Words`: function names, specifiers, numbers, values and formulas as the locale writes them). en-US's file has every message, the engine's English text; a locale without the text of a message gives the English one.

`test/messages-catalog.test.ts` checks that ru-RU writes every message of en-US and shows every parameter the English text shows, and `test/messages-paths.test.ts` that each way a message reaches a person (a typed formula, each refusal of a structure operation in every kind of session, `explain`) gives it in Russian in ru-RU and in English elsewhere.

## How the engine does it

`src/locale/rules.ts` defines `Rules`, one object per book. en-US is the engine's own code. Another locale is built by `regionalRules` (`src/locale/rules/regional.ts`) from its data (`Locale.regional`): separators, the order of dates, the currency and where its symbol stands, month and day names, TEXT codes, the collation tag, the code page of CHAR and CODE, the word of new table columns' names, and the words for TRUE, FALSE and errors.

The engine builds the rules when it creates a book's state and passes them to operators and criteria through the runtime, and to every function through `Cx.rules`. The argument kinds of `src/args.ts` convert values with the rules, and functions that read or write text themselves (TEXT, DOLLAR, FIXED, VALUE, the complex functions, the lookups and others) use them directly. Code outside a book, such as the parser, works in en-US.

## Adding a locale

A locale added to the package lets books in it compute as Excel in that language and region does, and lets hosts show formulas in its language. It takes one module and three lines:

- the module `src/locales/xx-YY.ts`, which holds everything the engine knows of the locale and exports it as `xxYY`, the tag without its hyphen;
- its tag in `PACKAGE_LOCALES` (`src/locale/registry.ts`), which names the module in the message of a book in it that is not loaded, and lets `readXlsx` keep it;
- its import in `LOCALE_MODULES` (`src/worker/locales.ts`), through which a pool's thread takes its messages, and the tests and harnesses load every locale of the package;
- its subpath in `package.json`, `"./locales/xx-YY"`, as ru-RU's.

`test/locale-modules.test.ts` checks that the four agree. No code of the engine changes, as long as the locale's rules are ones the engine reads from data ([What the data cannot say](#what-the-data-cannot-say)), and none of the tests and harnesses that compare it with Excel ([Tests and harnesses](#tests-and-harnesses)).

### The locale module

A locale module exports one `Locale` (`src/locale/types.ts`) and has no effect of its own. ru-RU's, `src/locales/ru-RU.ts`, is the one to copy. It has three parts.

The formula text: the tag, the separators, the words for TRUE and FALSE, error literals, function names and the specifiers of structured references. The file holds most of it as the constant `FORMULA_TEXT` between two marks:

```
// ---- formula text, generated by bench/locale/generate.ts: do not edit ----
...
// ---- end of formula text ----
```

`npx tsx bench/locale/generate.ts xx-YY` rewrites the lines between the marks from Excel's snapshot of the locale (`test/excel/locale/xx-YY.json`) and leaves the rest of the file as it is. The tag and the specifiers are written by hand after the marks: the snapshot's books have no tables, so ru-RU's specifiers (`#Все`, `#Данные`, `#Заголовки`, `#Итоги`, `#Эта строка`) were taken from the table set `locale`.

`regional` is the rules of computation, a `RegionalData` (`src/locale/rules/data.ts`) written by hand:

| field | what it gives | ru-RU |
|---|---|---|
| `decimal`, `group` | the separators of numbers in text, FIXED and DOLLAR, NUMBERVALUE's defaults | `,` and a space: `1 234,5` |
| `dateOrder` | the order of day, month and year in date text: `dmy`, `mdy` or `ymd` | `dmy`: `31.01.2025`, and `1.5` is 1 May |
| `currency` | DOLLAR's format, and where VALUE reads the symbol | `₽` after the amount, a space between, a minus: `-1 234,57 ₽` |
| `months`, `monthsShort`, `days`, `daysShort` | the names TEXT writes; date text reads three letters or more of a full month name | `Январь`, `янв`, `пятница`, `Пт` |
| `codes`, `colors` | TEXT's letters of dates and times, the name of General, the colour names | `ДД.ММ.ГГГГ ч:мм`, `Основной`, `[Красный]` |
| `dateSeparator` | what `/` in a cell's date format shows | `.`: `dd/mm/yyyy` shows `31.01.2025` |
| `lcid`, `longDate`, `longTime`, `longMonths` | the locale code (`[$-419]`), the long date and time (`[$-F800]`, `[$-F400]`) | `пятница, 31 января 2025 г.` |
| `collation` | the `Intl.Collator` tag that orders text | `en-US`: Latin before Cyrillic, as Excel for Windows sorts |
| `builtin` | the built-in formats the locale shows its own way, as en-US codes | 14 is `dd.mm.yyyy` |
| `codePage` | CHAR's and CODE's Windows code page: the name of its published table and its characters | `windows-1251` |
| `columnWord` | the word of a new table column's name | `Столбец`, which makes `Столбец1` |

Two kinds of data count in books of every locale, so they live in the core rather than in the module, and an en-US book shows the same whether the locale is loaded or not. The widths of the letters of the locale's script in Calibri 11, for what a cell shows by its column's width, go into `src/numfmt/glyphs.ts`, beside the Latin letters and ru-RU's Cyrillic ones: text of any script may stand in any book, and a character that no table has counts as a digit. The month and day names that the locale's code gives in a format (`[$-419]mmmm`) go into `src/locale/lcid.ts`, which `rules.ts` reads for the code, and the module takes them from there (`...RU_NAMES` in ru-RU's `regional`).

`messages` are the engine's messages for people in the locale's language ([Messages](#messages)): a function for each key of `MessageParams` that writes the text from its parameters and the locale's `Words`. A locale may leave out any of them, and those show in English; `test/messages-catalog.test.ts` checks ru-RU's against en-US's keys, and a new locale's full set would join it there.

en-US's module holds its formula text and its messages, the English ones every locale falls back to, and is built into the root: its rules are the engine's code in `src/locale/rules.ts`, where its code page (`windows-1252`) and the word `Column` are.

### Where the data come from

Formula text is read from Excel for Mac in the locale's language: `bench/locale/snapshot.ts` writes the snapshot, Excel's own text of 6,381 formulas and its international settings, and `generate.ts` the file's generated part from it. The Windows harness takes no snapshot of formula text.

Regional data start from Excel's international settings, which the snapshot keeps (separators, the currency), and from case sets computed in Excel (names, TEXT codes, colours, the long date). They are checked against Excel for Windows too. Where Excel for Mac and Excel for Windows part, Windows is the reference. ru-RU follows it for the order of text (Latin before Cyrillic), month names (`Январь`, not `январь`), the short and long time, and CHAR and CODE (Windows-1251, not Mac Cyrillic); it keeps one value of the Mac's, the long date `[$-F800]`.

### What the engine reads from data

These rules were written for ru-RU and are now read from a locale's data. Each was compared with Excel in ru-RU only; in the other shapes the data allow they are the engine's own rule, made by analogy with en-US, until a locale's Excel shows otherwise.

#### Dates in text

`dateOrder` places the numbers: day first in ru-RU (`31.01.2025`, `31/1/25`), month first in `mdy` (`1/31/2025`), year first in `ymd` (`2025/1/31`, `25.1.31`). In every order the numbers are separated by `.`, `/` or `-`, `2025-01-31` is a date, and month names are read with the day first (`31 янв 2025`, `янв 2025`). Two numbers are a day and a month in day-first text (`1.5`), or a month and a year when the second is past 12 (`1.13` is January 2013); in the other orders they are a month and then a day or a year, as en-US reads `1/31` and `1/2025`.

#### Currency in number text

VALUE and the other readers of number text take the symbol where `currency.before` puts it: after the amount in ru-RU (`1 234,5 ₽`), before it otherwise, with one sign before or after the symbol (`-€5`, `€-5`) and no fraction, as en-US reads `$`.

#### Case of names

Defined names, table names and sheet names are compared in upper case by Unicode's default rules, the same in every locale (`caseKey`, `src/collate.ts`), a letter's whole upper case: `straße` is the name `STRASSE`, `ﬁ` is `FI`, as Excel for Mac and for Windows compare defined names and tables. The dotless `ı` and the long `ſ` stay themselves, no `I` and `S`, as Excel for Windows keeps them (Excel for Mac takes `ſ` for `S`). The words of formula text, function names, `TRUE` and `FALSE`, LET's names and a locale's words, compare letter by letter (`wordKey`): a letter whose upper case is two letters stays itself, so `gauß(1)` is no `GAUSS` and `falſe` no FALSE (the sets `name-case-held`). Excel compares sheet names letter by letter too and takes a sheet `ß` beside `SS`; the engine compares them as names, a known difference. For Latin and Cyrillic both keys are what Russian and English rules give, character for character (`test/case-key.test.ts`). Turkish rules pair `i` with `İ` and `ı` with `I`, while these rules take `i` to `I` and keep `ı` and `İ` apart; how Excel tr-TR compares such names has not been measured.

#### Number formats

TEXT refuses a point before a digit placeholder where the point is neither the decimal nor the group separator (`0.00` in ru-RU), and the group separator in a fraction format (`# ?/?` in ru-RU, where the group separator is a space).

#### Code pages

CHAR and CODE take codes 128 to 255 from the code page in the locale's data; `test/codepage.test.ts` checks every listed locale's page against the published table its data names (`test/fixtures/unicode/`), so a new page needs its table there.

### What the data cannot say

Some locales need code as well as data:

- Scripts written right to left (Arabic, Hebrew): neither the translation of typed formulas nor what cells show has been checked with them.
- Digits of a script's own (Arabic-Indic, Devanagari): number and date text is read with the digits 0 to 9 only, and TEXT's `[DBNum1]` and the like are not read.
- Calendars other than the Gregorian (`B2`, `[$-x-sysdate]`, Japanese eras): dates are Gregorian.
- Groups of digits other than three (`1,00,000` in hi-IN): FIXED and DOLLAR group by three, and number text takes groups of three digits or more.
- Double-byte code pages (East Asian locales): CHAR and CODE take one byte.
- Date text with the month name before the day (`Jan 31, 2025`): read only by en-US's code.
- Rules the data do not describe. The readers of number and date text, TEXT's codes and the rest were measured in ru-RU and en-US; a locale's own Excel may answer by a rule neither has, and that rule is code.

### Tests and harnesses

A new locale needs no change to the tests or the harnesses. They load every module of `LOCALE_MODULES` (`test/locales.ts`, `bench/coverage/locales.ts`), and the tests compare the engine with Excel in every loaded locale whose answers are stored, and name a locale without them as skipped: `acceptance against Excel xx-YY: no answers of Excel xx-YY to the coverage measurement (test/excel/coverage/excel-xx-YY.json)`. The harnesses take the locale's tag and store its answers in files of their own.

#### Where a locale's answers are

Answers are found by tag (`bench/coverage/locales.ts`). en-US's are at the top of each directory of `test/excel/`, another locale's in a directory named by its tag:

- the case sets: `cases/<set>.json` in en-US, `cases/ru-RU/<set>.json` in ru-RU, `cases/xx-YY/<set>.json` in xx-YY; the same for what cells show (`display/`) and tables (`tables/`);
- the coverage measurement: `coverage/excel-en.json` and `coverage/excel-ru.json`, as they were first named, and `coverage/excel-xx-YY.json` in another locale;
- the formula text: `locale/xx-YY.json`;
- Excel for Windows's answers: the same layout under `win/`, and the built-in formats in `win/xlsx/xx-YY/formats.json`.

A directory of stored answers named by a tag that the package lacks fails a test (`bench/coverage/locales.test.ts`), so answers are never left unchecked.

#### Taking the answers on the Mac

The steps of `bench/coverage/run.ts` take a locale as their last argument: `en`, `ru` or a tag. The snapshot comes first, as the locale's file is generated from it:

```sh
npx tsx bench/locale/snapshot.ts xx-YY             # the formula text, into test/excel/locale/xx-YY.json
npx tsx bench/coverage/run.ts excel xx-YY          # the coverage measurement, into test/excel/coverage/excel-xx-YY.json
npx tsx bench/coverage/run.ts probes <set> xx-YY   # a case set, into test/excel/cases/xx-YY/<set>.json
npx tsx bench/coverage/run.ts display <set> xx-YY  # what cells show
npx tsx bench/coverage/run.ts tables <set> xx-YY   # tables, with formulas typed and read in the locale's language
```

Excel for Mac computes in a locale by its launch arguments. en-US starts with `-AppleLanguages (en) -AppleLocale en_US`, and ru-RU with none, in the Russian of the Mac the answers were taken on. Another tag starts with `-AppleLanguages (xx) -AppleLocale xx_YY`; that form has not yet been tried with a third language. The harness knows Excel's window "We found a problem with some content" by its English and Russian text. In another interface language a book Excel refuses waits out the step's three minutes instead of being declined at once, until the window's text is added to `REPAIR` and `declineRepair` in `bench/coverage/excel.ts`.

#### Taking the answers on Windows

`bench/windows/run.ts` takes the same arguments, for the package's locales only (`npx tsx bench/windows/run.ts all xx-YY`). It sets the machine to the locale's language, regional format and system locale, and checks Excel's interface language against the locale's code (`lcid`) and the system code page against its `codePage`. The machine needs the language: Excel's interface language follows the Windows display language, so a new language needs its Windows language pack, and Office's language pack for it as well, as the English one is installed there.

#### Known differences

A case where the engine knowingly differs from Excel in a locale is pinned with its reason: the test passes on the engine's answer, and fails once the case comes to match Excel, until it is taken off the list.

- ru-RU's list is `RU_KNOWN` in `test/excel-ru-known.ts`, with its own reasons in `RU_REASONS`. A new locale's list goes into a file of its own beside it, named in `KNOWN_BY_LOCALE` (`test/excel-locale-cases.ts`); its new reasons join `LOCALE_REASONS` there, and each has a line in the README's known differences.
- Against Excel for Windows, `WIN_KNOWN` (`test/excel-win-known.ts`) holds every locale's, by keys that start with the tag: `ru-RU:locale:format#244`.
- What cells show and tables keep their lists by set: `MAC_SHOWN` in `test/display-excel.test.ts`, `KNOWN` in `test/tables-excel.test.ts`.

#### Formula text and code pages

`test/locale-excel.test.ts` checks the formula text of every locale with a snapshot, `test/locale-roundtrip.test.ts` takes the case sets' formulas round every locale's text, and `test/codepage.test.ts` checks each locale's code page against its published table. A new page needs that table in `test/fixtures/unicode/`; the test names the missing file.

### Accepting a locale

A locale is ready when the engine gives Excel's answers in it, and every difference left is listed with its reason. For ru-RU that is:

- the coverage measurement computed in Excel ru-RU (`npx tsx bench/coverage/run.ts excel ru`) and compared by `test/excel-locales.test.ts`;
- the case sets (`npx tsx bench/coverage/run.ts probes <set> ru`), what cells show (`display <set> ru`) and tables (`tables <set> ru`), with their tests, as `bench/coverage/README.md` describes;
- the same books in Excel for Windows (`npx tsx bench/windows/run.ts all ru`, `bench/windows/README.md`);
- the differences: `RU_KNOWN` in `test/excel-ru-known.ts` pins the engine's answer of each case that differs from Excel ru-RU, with a reason, and `WIN_KNOWN` does the same against Excel for Windows.

A new locale needs the same: the measurement and the sets computed in its Excel, and its own list of differences ([Tests and harnesses](#tests-and-harnesses)).

## How it was checked

`test/excel-locales.test.ts` computes the measurement's book with `locale: 'ru-RU'` and compares every whitelisted function's cases with Excel ru-RU's answers in `test/excel/coverage/excel-ru.json`. `test/excel-locales-sets.test.ts` does the same for the ru-RU case sets in `test/excel/cases/ru-RU/`, among them `locale` (1,109 cases) and `collation` (174). `npx tsx bench/coverage/run.ts probes <set> ru` computes a set in Excel ru-RU. Known differences are listed in `RU_KNOWN` with a line each in the README; among them are the cases where the engine follows Excel for Windows and the Mac answers otherwise (the order of text, month names).

`test/excel-win.test.ts` checks the same books against Excel for Windows (`test/excel/win/`, `bench/windows`), with the set `collation-more` (62 cases of the order inside Cyrillic and Latin) measured in Windows alone.

`test/locale-excel.test.ts` checks that both translations reproduce Excel's own text of all 6,381 formulas in ru-RU, and `test/locale-roundtrip.test.ts` that 15,266 formulas of the case sets come back unchanged from every locale.

`test/locale-modules.test.ts` loads nothing until its last part: a ru-RU book is the `InputError` that names the module in `compute`, a session, `editStructure`, `fitColumns`, the same thread and `writeXlsx`, and so is the tag in every function that takes a locale; `ruRU` as an object gives these functions their answers without loading; an en-US book gives Excel's `Январь` for `[$-419]mmmm` and Excel for Mac's widths of Cyrillic letters; `readXlsx` keeps a file's ru-RU; once loaded, a ru-RU book computes as before in a session and in the same thread, with closures and eval; `registerLocale` refuses another object with a loaded tag, en-US, and a locale without rules or with wrong messages; the files of `src/locales/`, `PACKAGE_LOCALES`, `LOCALE_MODULES` and the subpaths agree. `test/locale-args.test.ts` gives every function with a locale the tag and the object and gets the same answer, and `test/locale-pool.test.ts` checks the threads of a pool. `test/bundle-size.test.ts` bundles `compute` alone with esbuild: 253 KB minified, with en-US alone of the locales, under its budget of 300 KB.

Excel for Mac and Excel for Windows (16.0.20430) compute ru-RU alike but for their regional data. Where those part, the engine follows Windows: the order of text, month names, the times of built-in formats, `CHAR` and `CODE`. It keeps the Mac's long date `[$-F800]` (`пятница, 31 января 2025 г.`; Windows's has no weekday). Formula text in ru-RU (array constants and the like) was not compared with Windows.
